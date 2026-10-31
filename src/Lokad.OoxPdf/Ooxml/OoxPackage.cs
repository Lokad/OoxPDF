using System.Buffers.Binary;
using System.IO.Compression;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal sealed class OoxPackage : IDisposable
{
    private const int MaxEntryCount = 10_000;
    internal const long MaxPartBytes = 64L * 1024L * 1024L;
    internal const long MaxTotalBytes = 256L * 1024L * 1024L;
    internal const long MaxContentTypesBytes = 4L * 1024L * 1024L;
    internal const long MaxCompressedBytes = 512L * 1024L * 1024L;
    internal const long MaxCentralDirectoryBytes = 32L * 1024L * 1024L;

    private readonly Dictionary<string, OoxPart> parts;
    private IMemoryScope? storageScope;
    private IMemoryReservation? xmlStorage;
    private bool disposed;
    private const int ScratchBytes = 81920;
    // Coarse capacity/headroom, not object accounting or a managed-heap ceiling.
    // Sparse text and dense elements expand differently. Leave conservative
    // source-size headroom for cached records and temporary normalization.
    private const int XmlRetainedFactor = 32;
    private const int XmlConstructionFactor = 64;
    private const int XmlReaderHeadroom = 64 * 1024;

    // Conversion-local parse caches for LoadXml/relationship consumers. Callers
    // read the shared DOMs or work on detached copies. One conversion owns the
    // package on one thread; these caches need no shared synchronization.
    private readonly Dictionary<string, XDocument> xmlCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, IReadOnlyDictionary<string, OoxRelationship>> relationshipCache = new(StringComparer.OrdinalIgnoreCase);

    private OoxPackage(Dictionary<string, OoxPart> parts, OoxContentTypes contentTypes, IMemoryScope? storageScope, IMemoryReservation? xmlStorage)
    {
        this.parts = parts;
        ContentTypes = contentTypes;
        this.storageScope = storageScope;
        this.xmlStorage = xmlStorage;
    }

    public OoxContentTypes ContentTypes { get; }

    public IReadOnlyCollection<OoxPart> Parts
    {
        get { ObjectDisposedException.ThrowIf(disposed, this); return parts.Values; }
    }

    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        // Source arrays are borrowed through PDF resource emission. The owner
        // closes the package only after those readers finish, before refunding.
        parts.Clear();
        xmlCache.Clear();
        relationshipCache.Clear();
        IMemoryScope? held = storageScope;
        storageScope = null;
        xmlStorage = null;
        held?.Dispose();
    }

    public static OoxPackage Open(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using FileStream stream = File.OpenRead(path);
        return Open(stream, cancellationToken);
    }

    public static OoxPackage Open(Stream stream, CancellationToken cancellationToken) =>
        OpenAsync(stream, cancellationToken).GetAwaiter().GetResult();

    public static async ValueTask<OoxPackage> OpenAsync(Stream stream, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // bound archive intake before ZipArchive buffers it. For forward-only
        // (non-seekable) input, ZipArchive would otherwise consume an unbounded or
        // never-ending stream without observing cancellation. Stage non-seekable input
        // through a bounded, cancellation-aware copy first; seekable inputs are checked
        // by length. .NET 11 reads them directly; .NET 10 also stages them below
        // because its ZIP directory API reads synchronously. Caller retains ownership.
        MemoryStream? stagedBytes = null;
        Stream archiveStream = stream;
        // one scratch buffer is allocated for the whole open call instead of
        // allocating an 81,920-byte array per retained part plus staging.
        IMemoryScope? memory = OoxConversionBudget.Current?.Memory?.CreateScope();
        try
        {
            var working = OoxMemory.Reserve(memory, ScratchBytes);
            var compressed = OoxMemory.Reserve(memory, 0);
            var retained = OoxMemory.Reserve(memory, 0);
            var parsedXml = OoxMemory.Reserve(memory, 0);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] scratch = new byte[ScratchBytes];
            if (!stream.CanSeek)
            {
                stagedBytes = await StageCompressedInputAsync(stream, MaxCompressedBytes, scratch, compressed, cancellationToken).ConfigureAwait(false);
                archiveStream = stagedBytes;
            }
            else
            {
                bool stageSeekable = false;
                try
                {
                    long length = stream.Length;
                    if (length > MaxCompressedBytes)
                    {
                        throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported compressed size.");
                    }
                }
                catch (Exception ex) when (ex is NotSupportedException or IOException)
                {
                    // Limit failures must propagate; only length-query failures are handled here.
                    if (ex is OoxPdfLimitExceededException)
                    {
                        throw;
                    }

                    // R05: a seekable stream that cannot report length is staged like
                    // forward-only input under the same compressed-size quota, since
                    // ZipArchive itself requires Length. Caller retains ownership of
                    // the original stream; the staging buffer is disposed below.
                    stageSeekable = true;
                }

                if (stageSeekable)
                {
                    stagedBytes = await StageCompressedInputAsync(stream, MaxCompressedBytes, scratch, compressed, cancellationToken).ConfigureAwait(false);
                    archiveStream = stagedBytes;
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            await PreflightCentralDirectoryAsync(archiveStream, scratch, cancellationToken).ConfigureAwait(false);
#if NET10_0
            // .NET 10 lazily reads ZipArchive.Entries synchronously, even after
            // CreateAsync. Keep those reads on owned memory; external input is
            // copied with the existing bounded, admitted async staging path.
            if (stagedBytes is null)
            {
                stream.Seek(0, SeekOrigin.Begin);
                stagedBytes = await StageCompressedInputAsync(stream, MaxCompressedBytes, scratch, compressed, cancellationToken).ConfigureAwait(false);
                archiveStream = stagedBytes;
            }
#endif
            OoxContentTypes contentTypes;
            Dictionary<string, OoxPart> parts;
            ZipArchive archive = await ZipArchive.CreateAsync(archiveStream, ZipArchiveMode.Read,
                leaveOpen: true, entryNameEncoding: null, cancellationToken).ConfigureAwait(false);
            await using (archive.ConfigureAwait(false))
            {
                if (archive.Entries.Count > MaxEntryCount)
                {
                    throw new OoxPdfLimitExceededException($"OOXML package has too many ZIP entries: {archive.Entries.Count}.");
                }

                // Bound central-directory claims before extracting retained parts.
                long claimedCompressed = 0;
                foreach (ZipArchiveEntry claimed in archive.Entries)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    long entryCompressed = claimed.CompressedLength;
                    if (entryCompressed < 0)
                    {
                        continue;
                    }

                    try
                    {
                        claimedCompressed = checked(claimedCompressed + entryCompressed);
                    }
                    catch (OverflowException ex)
                    {
                        throw new OoxPdfLimitExceededException("OOXML package central directory exceeds the maximum supported size.", ex);
                    }

                    if (claimedCompressed > MaxCompressedBytes)
                    {
                        throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported compressed size.");
                    }
                }

                ZipArchiveEntry? contentTypesEntry = archive.GetEntry("[Content_Types].xml")
                    ?? throw new InvalidDataException("OOXML package is missing [Content_Types].xml.");

                // The declared entry length is untrusted: copy through a bounded buffer first
                // so an oversized Content_Types part fails before XML parsing amplifies it.
                byte[] contentTypesBytes;
                Stream contentTypesStream = await contentTypesEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
                await using (contentTypesStream.ConfigureAwait(false))
                {
                    contentTypesBytes = await CopyBoundedAsync(contentTypesStream, MaxContentTypesBytes, "[Content_Types].xml", scratch, working, retained, cancellationToken).ConfigureAwait(false);
                }
                OoxMemory.Resize(parsedXml, (long)contentTypesBytes.Length * XmlConstructionFactor + XmlReaderHeadroom);
                cancellationToken.ThrowIfCancellationRequested();
                using (var contentTypesBuffer = new MemoryStream(contentTypesBytes, writable: false))
                {
                    contentTypes = OoxContentTypes.Parse(contentTypesBuffer, cancellationToken);
                }
                // Content-type lookup records survive the temporary DOM/reader.
                OoxMemory.Resize(parsedXml, (long)contentTypesBytes.Length * XmlRetainedFactor);

                cancellationToken.ThrowIfCancellationRequested();

                long totalBytes = 0;
                parts = new Dictionary<string, OoxPart>(StringComparer.OrdinalIgnoreCase);

                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (entry.FullName.EndsWith('/'))
                    {
                        continue;
                    }

                    string partName = OoxPath.NormalizePartName(entry.FullName);
                    if (entry.Length < 0 || entry.Length > MaxPartBytes)
                    {
                        throw new OoxPdfLimitExceededException($"OOXML part '{partName}' exceeds the maximum supported size.");
                    }

                    string? contentType = partName == "/[Content_Types].xml"
                        ? "application/xml"
                        : contentTypes.GetContentType(partName);
                    if (contentType is null)
                    {
                        continue;
                    }

                    // Two archive entries normalizing to one part name are ambiguous: keep neither guess.
                    if (parts.ContainsKey(partName))
                    {
                        throw new InvalidDataException($"OOXML package contains a duplicate part {partName}.");
                    }

                    // The [Content_Types].xml entry was already copied and parsed above: reuse
                    // those bytes instead of copying the entry a second time (G01).
                    byte[] partBytes;
                    if (string.Equals(entry.FullName, "[Content_Types].xml", StringComparison.Ordinal))
                    {
                        partBytes = contentTypesBytes;
                    }
                    else
                    {
                        // Declared lengths are untrusted: enforce the budgets on actual copied
                        // bytes, capped by the remaining aggregate capacity so a single part
                        // cannot overshoot the package budget by its full size before rejection.
                        long remaining = MaxTotalBytes - totalBytes;
                        if (remaining < 0)
                        {
                            throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported uncompressed size.");
                        }

                        long partBudget = Math.Min(MaxPartBytes, remaining);
                        bool cappedByTotal = partBudget < MaxPartBytes;
                        try
                        {
                            Stream entryStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
                            await using (entryStream.ConfigureAwait(false))
                            {
                                partBytes = await CopyBoundedAsync(entryStream, partBudget, partName, scratch, working, retained, cancellationToken).ConfigureAwait(false);
                            }
                        }
                        catch (OoxPdfLimitExceededException exception) when (cappedByTotal && exception.Kind == OoxPdfLimitKind.DocumentBound)
                        {
                            throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported uncompressed size.");
                        }
                    }

                    totalBytes = checked(totalBytes + partBytes.LongLength);
                    if (totalBytes > MaxTotalBytes)
                    {
                        throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported uncompressed size.");
                    }

                    parts.Add(partName, new OoxPart(partName, contentType, partBytes));
                }
            }
            // Archive/entry disposal has completed. Refund intake buffers before
            // transferring the package's retained-byte scope to the result.
            stagedBytes?.Dispose();
            stagedBytes = null;
            OoxMemory.Resize(compressed, 0);
            OoxMemory.Resize(working, 0);
            OoxPackage result = new OoxPackage(parts, contentTypes, memory, parsedXml);
            memory = null;
            return result;
        }
        finally
        {
            try { stagedBytes?.Dispose(); }
            finally { memory?.Dispose(); }
        }
    }

    // R05: bound central-directory metadata before ZipArchive materializes one entry
    // object per record. The End-of-Central-Directory carries the entry count and
    // directory size, so rejecting from those bytes alone avoids allocating millions
    // of entry objects for a hostile directory. Non-seekable input is already staged
    // to memory above, so this covers both intake paths; streams that cannot seek or
    // report length, multi-disk archives, and Zip64 shapes the reader does not parse
    // fall through to the existing post-construction checks. The stream position is
    // restored. Limit failures always propagate; only metadata-query failures fall
    // through (OoxPdfLimitExceededException derives from IOException, so every catch
    // below rethrows it explicitly).
    private static async ValueTask PreflightCentralDirectoryAsync(Stream archiveStream, byte[] scratch, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!archiveStream.CanSeek)
        {
            return;
        }

        long length;
        try
        {
            length = archiveStream.Length;
        }
        catch (Exception ex) when (ex is NotSupportedException or IOException)
        {
            if (ex is OoxPdfLimitExceededException)
            {
                throw;
            }

            return;
        }

        const int EndOfCentralDirectorySize = 22;
        if (length < EndOfCentralDirectorySize)
        {
            return;
        }

        long savedPosition;
        try
        {
            savedPosition = archiveStream.Position;
        }
        catch (Exception ex) when (ex is NotSupportedException or IOException)
        {
            if (ex is OoxPdfLimitExceededException)
            {
                throw;
            }

            return;
        }

        try
        {
            // The record hides behind at most a 65,535-byte comment; the scratch
            // buffer allocated for the whole open call fits the scan window.
            long scanSize = Math.Min(length, 65_535 + EndOfCentralDirectorySize);
            long scanStart = length - scanSize;
            int scanLength = checked((int)scanSize);
            if (scratch.Length < scanLength)
            {
                return;
            }

            if (!TrySeekCentralDirectory(archiveStream, scanStart)) return;
            int read = 0;
            while (read < scanLength)
            {
                int chunk = await archiveStream.ReadAsync(scratch.AsMemory(read, scanLength - read), cancellationToken).ConfigureAwait(false);
                if (chunk == 0)
                {
                    return;
                }

                read += chunk;
            }

            // The true record is the last signature whose comment length reaches the end.
            long recordOffset = -1;
            for (long candidate = scanLength - EndOfCentralDirectorySize; candidate >= 0; candidate--)
            {
                if (scratch[candidate] == 0x50 && scratch[candidate + 1] == 0x4B &&
                    scratch[candidate + 2] == 0x05 && scratch[candidate + 3] == 0x06)
                {
                    int commentLength = scratch[candidate + 20] | (scratch[candidate + 21] << 8);
                    if (candidate + EndOfCentralDirectorySize + commentLength == scanLength)
                    {
                        recordOffset = scanStart + candidate;
                        break;
                    }
                }
            }

            if (recordOffset < 0)
            {
                return;
            }

            int recordBase = checked((int)(recordOffset - scanStart));
            int diskNumber = scratch[recordBase + 4] | (scratch[recordBase + 5] << 8);
            int directoryDisk = scratch[recordBase + 6] | (scratch[recordBase + 7] << 8);
            if (diskNumber != 0 || directoryDisk != 0)
            {
                return;
            }

            long entryCount = (long)(scratch[recordBase + 10] | (scratch[recordBase + 11] << 8));
            long directorySize = (long)scratch[recordBase + 12] | ((long)scratch[recordBase + 13] << 8) |
                ((long)scratch[recordBase + 14] << 16) | ((long)scratch[recordBase + 15] << 24);
            if (entryCount == 0xFFFF || directorySize == 0xFFFFFFFF)
            {
                var counts = await ReadZip64CountsAsync(archiveStream, recordOffset, scratch, cancellationToken).ConfigureAwait(false);
                if (counts is not { } zip64) return;
                entryCount = zip64.EntryCount;
                directorySize = zip64.DirectorySize;
            }

            if (entryCount > MaxEntryCount)
            {
                throw new OoxPdfLimitExceededException($"OOXML package has too many ZIP entries: {entryCount}.");
            }

            if (directorySize > MaxCentralDirectoryBytes)
            {
                throw new OoxPdfLimitExceededException("OOXML package central directory exceeds the maximum supported size.");
            }
        }
        finally
        {
            _ = TrySeekCentralDirectory(archiveStream, savedPosition);
        }
    }

    // Only unavailable metadata/seek operations fall back. An actual read failure
    // belongs to the caller's transport and must escape without a second read.
    private static bool TrySeekCentralDirectory(Stream stream, long position)
    {
        try { stream.Seek(position, SeekOrigin.Begin); return true; }
        catch (Exception ex) when (ex is NotSupportedException or IOException)
        {
            if (ex is OoxPdfLimitExceededException) throw;
            return false;
        }
    }

    private static async ValueTask<(long EntryCount, long DirectorySize)?> ReadZip64CountsAsync(
        Stream archiveStream, long recordOffset, byte[] scratch, CancellationToken cancellationToken)
    {
        // The Zip64 locator sits 20 bytes before the End-of-Central-Directory.
        if (recordOffset < 20 || scratch.Length < 56 || !TrySeekCentralDirectory(archiveStream, recordOffset - 20)) return null;
        int read = 0;
        while (read < 20)
        {
            int chunk = await archiveStream.ReadAsync(scratch.AsMemory(read, 20 - read), cancellationToken).ConfigureAwait(false);
            if (chunk == 0) return null;
            read += chunk;
        }
        if (!(scratch[0] == 0x50 && scratch[1] == 0x4B && scratch[2] == 0x06 && scratch[3] == 0x07)) return null;
        long offset = BinaryPrimitives.ReadInt64LittleEndian(scratch.AsSpan(8, 8));
        if (offset < 0 || !TrySeekCentralDirectory(archiveStream, offset)) return null;
        read = 0;
        while (read < 56)
        {
            int chunk = await archiveStream.ReadAsync(scratch.AsMemory(read, 56 - read), cancellationToken).ConfigureAwait(false);
            if (chunk == 0) return null;
            read += chunk;
        }
        if (!(scratch[0] == 0x50 && scratch[1] == 0x4B && scratch[2] == 0x06 && scratch[3] == 0x06)) return null;
        long entries = BinaryPrimitives.ReadInt64LittleEndian(scratch.AsSpan(32, 8));
        long size = BinaryPrimitives.ReadInt64LittleEndian(scratch.AsSpan(40, 8));
        return entries >= 0 && size >= 0 ? (entries, size) : null;
    }

    public OoxPart? GetPart(string partName)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        return parts.TryGetValue(OoxPath.NormalizePartName(partName), out OoxPart? part) ? part : null;
    }

    public IReadOnlyList<OoxRelationship> GetRelationships(string sourcePartName, CancellationToken cancellationToken)
    {
        // Fresh list over the shared dictionary: same elements in document order, no reparse.
        return [.. GetRelationshipDictionary(sourcePartName, cancellationToken).Values];
    }

    public IReadOnlyDictionary<string, OoxRelationship> GetRelationshipDictionary(string sourcePartName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string relationshipPartName = OoxPath.GetRelationshipPartName(sourcePartName);
        if (relationshipCache.TryGetValue(relationshipPartName, out IReadOnlyDictionary<string, OoxRelationship>? cached))
        {
            return cached;
        }

        OoxPart? relationshipPart = GetPart(relationshipPartName);
        if (relationshipPart is null)
        {
            // Missing relationship parts share a cached empty dictionary.
            IReadOnlyDictionary<string, OoxRelationship> empty = new Dictionary<string, OoxRelationship>(StringComparer.Ordinal);
            relationshipCache[relationshipPartName] = empty;
            return empty;
        }

        IReadOnlyDictionary<string, OoxRelationship> relationships = ParseRelationshipDocument(LoadXml(relationshipPart, cancellationToken), sourcePartName, cancellationToken);
        relationshipCache[relationshipPartName] = relationships;
        return relationships;
    }

    public XDocument LoadXml(OoxPart part, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        cancellationToken.ThrowIfCancellationRequested();
        if (xmlCache.TryGetValue(part.Name, out XDocument? cached))
        {
            return cached;
        }

        long prior = xmlStorage?.Bytes ?? 0;
        OoxMemory.Resize(xmlStorage, prior + part.Bytes.LongLength * XmlConstructionFactor + XmlReaderHeadroom);
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            XDocument document;
            using (Stream stream = part.OpenRead())
            {
                document = SafeXml.Load(stream, cancellationToken);
            }
            // One package reservation owns all cached DOMs. Refund construction
            // headroom only after readers/normalization finish; cache hits reuse it.
            OoxMemory.Resize(xmlStorage, prior + part.Bytes.LongLength * XmlRetainedFactor);
            xmlCache[part.Name] = document;
            return document;
        }
        catch
        {
            OoxMemory.Resize(xmlStorage, prior);
            throw;
        }
    }

    private static Dictionary<string, OoxRelationship> ParseRelationshipDocument(XDocument document, string sourcePartName, CancellationToken cancellationToken)
    {
        var relationships = new Dictionary<string, OoxRelationship>(StringComparer.Ordinal);
        var seenIds = new HashSet<string>(StringComparer.Ordinal);

        foreach (XElement element in document.Root?.Elements(OoxNamespaces.PackageRelationshipsNamespace + "Relationship") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string id = RequiredAttribute(element, "Id");
            if (!seenIds.Add(id))
            {
                throw new InvalidDataException("OOXML package contains a duplicate relationship id.");
            }
            string type = RequiredAttribute(element, "Type");
            type = OoxNamespaces.NormalizeRelationshipType(type);
            string target = RequiredAttribute(element, "Target");
            string? targetMode = (string?)element.Attribute("TargetMode");
            string? resolvedTarget = targetMode?.Equals("External", StringComparison.OrdinalIgnoreCase) == true
                ? null
                : OoxPath.ResolveRelationshipTarget(sourcePartName, target);

            relationships.Add(id, new OoxRelationship(id, type, target, targetMode, resolvedTarget));
        }

        return relationships;
    }

    private static async ValueTask<MemoryStream> StageCompressedInputAsync(Stream source, long maxBytes, byte[] scratch, IMemoryReservation? storage, CancellationToken cancellationToken)
    {
        var destination = new MemoryStream();
        try
        {
            byte[] buffer = scratch;
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int read = await source.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false);
                if (read == 0)
                {
                    destination.Position = 0;
                    return destination;
                }

                if (checked(destination.Length + read) > maxBytes)
                {
                    throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported compressed size.");
                }

                GrowBuffer(destination, checked((int)(destination.Length + read)), maxBytes, storage, 0, cancellationToken);
                destination.Write(buffer, 0, read);
            }
        }
        catch
        {
            destination.Dispose();
            throw;
        }
    }

    private static async ValueTask<byte[]> CopyBoundedAsync(Stream source, long maxBytes, string what, byte[] scratch,
        IMemoryReservation? working, IMemoryReservation? retained, CancellationToken cancellationToken)
    {
        var destination = new MemoryStream();
        try
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int read = await source.ReadAsync(scratch.AsMemory(), cancellationToken).ConfigureAwait(false);
                if (read == 0)
                {
                    // The final array and construction buffer overlap until
                    // ToArray finishes. The package owns every retained array.
                    OoxMemory.Resize(retained, checked((retained?.Bytes ?? 0) + destination.Length));
                    cancellationToken.ThrowIfCancellationRequested();
                    return destination.ToArray();
                }
                if (checked(destination.Length + read) > maxBytes)
                    throw new OoxPdfLimitExceededException($"OOXML part {what} exceeds the maximum supported size.");
                GrowBuffer(destination, checked((int)(destination.Length + read)), maxBytes,
                    working, ScratchBytes, cancellationToken);
                destination.Write(scratch, 0, read);
            }
        }
        finally
        {
            destination.Dispose();
            OoxMemory.Resize(working, ScratchBytes);
        }
    }

    private static void GrowBuffer(MemoryStream buffer, int required, long maximum,
        IMemoryReservation? reservation, int otherBytes, CancellationToken cancellationToken)
    {
        int oldCapacity = buffer.Capacity;
        if (required <= oldCapacity) return;
        int capacity = checked((int)Math.Min(maximum, Math.Max(required, Math.Max(256L, 2L * oldCapacity))));
        OoxMemory.Resize(reservation, (long)otherBytes + oldCapacity + capacity);
        cancellationToken.ThrowIfCancellationRequested();
        buffer.Capacity = capacity;
        OoxMemory.Resize(reservation, (long)otherBytes + capacity);
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        return (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required Relationship attribute '{name}'.");
    }
}
