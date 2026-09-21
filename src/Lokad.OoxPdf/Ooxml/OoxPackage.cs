using System.IO.Compression;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Ooxml;

internal sealed class OoxPackage
{
    private const int MaxEntryCount = 10_000;
    internal const long MaxPartBytes = 64L * 1024L * 1024L;
    internal const long MaxTotalBytes = 256L * 1024L * 1024L;
    internal const long MaxContentTypesBytes = 4L * 1024L * 1024L;
    internal const long MaxCompressedBytes = 512L * 1024L * 1024L;

    private readonly Dictionary<string, OoxPart> parts;

    // PLAN W01: conversion-local immutable parse indexes. Shared masters, layouts, and
    // relationship parts were reparsed per slide (and visibility discovery reparsed
    // slides the scene later parsed again). Each unique part now parses once; every
    // consumer shares the resulting DOM/dictionary. Entries are immutable parse results:
    // downstream code only reads them (chart cache hydration is idempotent and
    // deterministic) or works on explicit copies. Single-owner affinity: the package
    // serves one conversion on one thread; these caches are not synchronized.
    // XmlParseCount/RelationshipParseCount instrument the W01 acceptance (one parse per
    // unique immutable part).
    private readonly Dictionary<string, XDocument> xmlCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, IReadOnlyDictionary<string, OoxRelationship>> relationshipCache = new(StringComparer.OrdinalIgnoreCase);
    private int xmlParseCount;
    private int relationshipParseCount;

    private OoxPackage(Dictionary<string, OoxPart> parts, OoxContentTypes contentTypes)
    {
        this.parts = parts;
        ContentTypes = contentTypes;
    }

    public OoxContentTypes ContentTypes { get; }

    public IReadOnlyCollection<OoxPart> Parts => parts.Values;

    // Counts every part-XML parse through LoadXml: content parts and relationship
    // parts alike. Together with RelationshipParseCount it shows each unique part
    // parsed once (a 3-slide shared-master deck parses 6 content + 6 rels parts).
    internal int XmlParseCount => xmlParseCount;

    internal int RelationshipParseCount => relationshipParseCount;

    public static OoxPackage Open(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using FileStream stream = File.OpenRead(path);
        return Open(stream, cancellationToken);
    }

    public static OoxPackage Open(Stream stream, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // PLAN M01: bound archive intake before ZipArchive buffers it. For forward-only
        // (non-seekable) input, ZipArchive would otherwise consume an unbounded or
        // never-ending stream without observing cancellation. Stage non-seekable input
        // through a bounded, cancellation-aware copy first; seekable inputs are checked
        // by length and read directly to avoid an extra full copy. Caller retains ownership.
        MemoryStream? stagedBytes = null;
        Stream archiveStream = stream;
        // PLAN G01: one scratch buffer is rented for the whole open call instead of
        // allocating an 81,920-byte array per retained part plus staging.
        byte[] scratch = new byte[81920];
        try
        {
            if (!stream.CanSeek)
            {
                stagedBytes = StageCompressedInput(stream, MaxCompressedBytes, scratch, cancellationToken);
                archiveStream = stagedBytes;
            }
            else
            {
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
                    // Length unavailable: fall through to bounded per-part enforcement below.
                    // Limit failures must propagate; only length-query failures are ignored here.
                    if (ex is OoxPdfLimitExceededException)
                    {
                        throw;
                    }
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read, leaveOpen: true);
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
        using (Stream contentTypesStream = contentTypesEntry.Open())
        {
            contentTypesBytes = CopyBounded(contentTypesStream, MaxContentTypesBytes, "[Content_Types].xml", scratch, cancellationToken);
        }

        OoxContentTypes contentTypes;
        using (var contentTypesBuffer = new MemoryStream(contentTypesBytes, writable: false))
        {
            contentTypes = OoxContentTypes.Parse(contentTypesBuffer, cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        long totalBytes = 0;
        var parts = new Dictionary<string, OoxPart>(StringComparer.OrdinalIgnoreCase);

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
                    using (Stream entryStream = entry.Open())
                    {
                        partBytes = CopyBounded(entryStream, partBudget, partName, scratch, cancellationToken);
                    }
                }
                catch (OoxPdfLimitExceededException) when (cappedByTotal)
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

            OoxPackage result = new OoxPackage(parts, contentTypes);
            return result;
        }
        finally
        {
            stagedBytes?.Dispose();
        }
    }

    public OoxPart? GetPart(string partName)
    {
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
            // Missing relationship parts are cached as empty without counting as parses.
            IReadOnlyDictionary<string, OoxRelationship> empty = new Dictionary<string, OoxRelationship>(StringComparer.Ordinal);
            relationshipCache[relationshipPartName] = empty;
            return empty;
        }

        IReadOnlyDictionary<string, OoxRelationship> relationships = ParseRelationshipDocument(LoadXml(relationshipPart, cancellationToken), sourcePartName, cancellationToken);
        relationshipCache[relationshipPartName] = relationships;
        relationshipParseCount++;
        return relationships;
    }

    public XDocument LoadXml(OoxPart part, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (xmlCache.TryGetValue(part.Name, out XDocument? cached))
        {
            return cached;
        }

        using Stream stream = part.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        xmlCache[part.Name] = document;
        xmlParseCount++;
        return document;
    }

    public static IReadOnlyList<OoxRelationship> ParseRelationships(Stream stream, string sourcePartName, CancellationToken cancellationToken)
    {

        XDocument document = SafeXml.Load(stream, cancellationToken);
        return [.. ParseRelationshipDocument(document, sourcePartName, cancellationToken).Values];
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

    private static MemoryStream StageCompressedInput(Stream source, long maxBytes, byte[] scratch, CancellationToken cancellationToken)
    {
        var destination = new MemoryStream();
        try
        {
            byte[] buffer = scratch;
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0)
                {
                    destination.Position = 0;
                    return destination;
                }

                if (checked(destination.Length + read) > maxBytes)
                {
                    throw new OoxPdfLimitExceededException("OOXML package exceeds the maximum supported compressed size.");
                }

                destination.Write(buffer, 0, read);
            }
        }
        catch
        {
            destination.Dispose();
            throw;
        }
    }

    private static byte[] CopyBounded(Stream source, long maxBytes, string what, byte[] scratch, CancellationToken cancellationToken)
    {
        byte[] buffer = scratch;
        using var destination = new MemoryStream();
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = source.Read(buffer, 0, buffer.Length);
            if (read == 0)
            {
                return destination.ToArray();
            }

            if (checked(destination.Length + read) > maxBytes)
            {
                throw new OoxPdfLimitExceededException($"OOXML part {what} exceeds the maximum supported size.");
            }

            destination.Write(buffer, 0, read);
        }
    }

    private static string RequiredAttribute(XElement element, string name)
    {
        return (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required Relationship attribute '{name}'.");
    }
}
