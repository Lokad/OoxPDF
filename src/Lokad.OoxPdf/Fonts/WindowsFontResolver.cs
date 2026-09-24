using System.Diagnostics.CodeAnalysis;
namespace Lokad.OoxPdf.Fonts;

/// <summary>
/// Resolves installed system fonts with a process-static discovery snapshot per
/// directory set.
/// </summary>
/// <remarks>
/// Retention contract: the snapshot retains one lazy <see cref="FileFontProgramSource"/>
/// per discovered file (shared by all its faces), and each source retains its full
/// program bytes after first use (files are capped individually; see
/// FileFontProgramSource). R13 ownership: the process-static snapshot owns the
/// sources, resolver instances share them, and conversions only borrow the arrays;
/// discovery spans (R13) bound the initial read. Aggregate retention is capped by
/// MaxSnapshotRetainedBytes with LRU eviction, so repeated conversions re-read from
/// local disk only after eviction. Hosts that rotate font directories or must
/// release memory call <see cref="InvalidateDiscoveryCaches"/>; existing resolver
/// instances keep their snapshot.
/// </remarks>
public sealed class WindowsFontResolver : IFontResolver, IFontCatalog
{
    private static readonly object CacheLock = new();
    private static readonly Dictionary<string, Lazy<DiscoverySnapshot>> DiscoveryCaches = new(StringComparer.OrdinalIgnoreCase);
    private readonly Lazy<DiscoverySnapshot> cache;

    // per-instance (usually per-conversion) resolution cache. Requests repeat
    // per run; the static snapshot underneath is shared. Capped with clear-all so a
    // long-lived directly-held resolver cannot grow without bound on adversarial
    // distinct families; per-conversion instances never approach the cap.
    internal const int MaxCachedResolutions = 4096;

    // R13: aggregate ceiling on snapshot-retained program bytes. A conversion
    // typically touches a handful of families; hundreds of megabytes only
    // accumulate across adversarial font directories, where LRU eviction re-reads
    // evicted files from local disk on demand.
    internal const long MaxSnapshotRetainedBytes = 512L * 1024L * 1024L;

    // R13: snapshot-owned LRU over per-file program sources. Mirrors the pack
    // eviction contract: retained entries move to the back on every access,
    // overflow evicts from the front, evicted bytes re-read on demand, and arrays
    // already handed out stay alive via GC. Eviction takes no source locks, so it
    // cannot deadlock against in-progress loads.
    internal sealed class RetainedFileTracker
    {
        private readonly long maxTotalBytes;
        private readonly IReadOnlyDictionary<string, FileFontProgramSource> sourcesByPath;
        private readonly object sync = new();
        private readonly Dictionary<string, long> retainedSizes = new(StringComparer.OrdinalIgnoreCase);
        private readonly LinkedList<string> retainedOrder = new();
        private long retainedBytes;

        internal RetainedFileTracker(long maxTotalBytes, IReadOnlyDictionary<string, FileFontProgramSource> sourcesByPath)
        {
            this.maxTotalBytes = maxTotalBytes;
            this.sourcesByPath = sourcesByPath;
        }

        internal void NoteLoaded(string path, long byteCount)
        {
            lock (sync)
            {
                if (retainedSizes.TryGetValue(path, out long previous))
                {
                    retainedBytes -= previous;
                    retainedOrder.Remove(path);
                }

                retainedSizes[path] = byteCount;
                retainedOrder.AddLast(path);
                retainedBytes += byteCount;

                while (retainedBytes > maxTotalBytes && retainedOrder.Count > 0)
                {
                    string eldest = retainedOrder.First!.Value;
                    if (retainedOrder.Count == 1)
                    {
                        // Only the just-loaded source is left; a single file always
                        // fits by the per-file cap, so this is unreachable in practice.
                        break;
                    }

                    retainedOrder.RemoveFirst();
                    retainedBytes -= retainedSizes.GetValueOrDefault(eldest);
                    retainedSizes.Remove(eldest);
                    if (sourcesByPath.TryGetValue(eldest, out FileFontProgramSource? source))
                    {
                        source.EvictCachedBytes();
                    }
                }
            }
        }

        internal void NoteAccessed(string path)
        {
            lock (sync)
            {
                if (retainedSizes.ContainsKey(path))
                {
                    retainedOrder.Remove(path);
                    retainedOrder.AddLast(path);
                }
            }
        }
    }
    private readonly Dictionary<FontRequest, FontFaceResolution> requestCache = new(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);

    private sealed record DiscoverySnapshot(
        IReadOnlyList<FontFaceResolution> Faces,
        Dictionary<string, FontFaceResolution[]> ByFamily,
        FontFaceResolution[] TextFaces);

    public WindowsFontResolver()
        : this(GetDefaultFontDirectories())
    {
    }

    internal WindowsFontResolver(string fontsDirectory)
        : this([fontsDirectory])
    {
    }

    private WindowsFontResolver(IReadOnlyList<string> fontDirectories)
    {
        cache = GetOrCreateCache();

        Lazy<DiscoverySnapshot> GetOrCreateCache()
        {
            string cacheKey = string.Join(
                "|",
                fontDirectories
                    .Where(d => !string.IsNullOrWhiteSpace(d))
                    .Select(Path.GetFullPath)
                    .Order(StringComparer.OrdinalIgnoreCase));
            lock (CacheLock)
            {
                if (!DiscoveryCaches.TryGetValue(cacheKey, out Lazy<DiscoverySnapshot>? cached))
                {
                    cached = new Lazy<DiscoverySnapshot>(() => Discover());
                    DiscoveryCaches[cacheKey] = cached;
                }

                return cached;

            DiscoverySnapshot Discover()
            {
                var fonts = new List<FontFaceResolution>();
                var sourcesByPath = new Dictionary<string, FileFontProgramSource>(StringComparer.OrdinalIgnoreCase);
                var retention = new RetainedFileTracker(MaxSnapshotRetainedBytes, sourcesByPath);
                foreach (string fontsDirectory in fontDirectories.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    SearchOption searchOption = fontsDirectory.Contains("CloudFonts", StringComparison.OrdinalIgnoreCase)
                        ? SearchOption.AllDirectories
                        : SearchOption.TopDirectoryOnly;
                    foreach (string path in Directory.EnumerateFiles(fontsDirectory, "*.*", searchOption)
                                 .Where(p => p.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) ||
                                     p.EndsWith(".otf", StringComparison.OrdinalIgnoreCase) ||
                                     p.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
                                 .Order(StringComparer.OrdinalIgnoreCase))
                    {
                        try
                        {
                            // discovery parses only the directory plus eight
                            // table extents, never outlines or layout tables. Read exactly
                            // those bytes instead of the whole program per file.
                            (byte[] bytes, bool complete, long fileLength) = ReadDiscoveryBytes(path);
                            if (!sourcesByPath.TryGetValue(path, out FileFontProgramSource? source))
                            {
                                source = new FileFontProgramSource(path, retention);
                                sourcesByPath[path] = source;
                            }
                            bool isCollection = OpenTypeFont.IsTrueTypeCollectionHeader(bytes);
                            int faceCount = OpenTypeFont.GetCollectionFontCount(bytes);
                            for (int faceIndex = 0; faceIndex < faceCount; faceIndex++)
                            {
                                OpenTypeFont.FontDiscoveryHeaders headers;
                                try
                                {
                                    // Collections parse each face in place instead of
                                    // repackaging per-face sfnt copies (G02).
                                    headers = isCollection
                                        ? OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(bytes, faceIndex, complete ? null : fileLength)
                                        : OpenTypeFont.ReadDiscoveryHeaders(bytes, faceIndex, complete ? null : fileLength);
                                }
                                catch (Exception spanEx) when (!complete && spanEx is InvalidDataException or ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException)
                                {
                                    // A span read can truncate bytes a malformed font's
                                    // parser would otherwise touch. Re-parse the full file
                                    // so accept/reject matches whole-program reads exactly.
                                    // Malformed files pay one extra bounded read; valid
                                    // fonts never take this path.
                                    bytes = File.ReadAllBytes(path);
                                    complete = true;
                                    headers = OpenTypeFont.ReadDiscoveryHeaders(bytes, faceIndex);
                                }
                                if (!string.IsNullOrWhiteSpace(headers.FamilyName))
                                {
                                    fonts.Add(new FontFaceResolution(
                                        headers.FamilyName,
                                        headers.FamilyName,
                                        new FontStyleKey(
                                            Bold: headers.WeightClass >= 600,
                                            Italic: Math.Abs(headers.ItalicAngle) > 0.01d,
                                            WeightClass: headers.WeightClass,
                                            FaceIndex: faceIndex,
                                            HasMathTable: headers.HasMathTable),
                                        source,
                                        IsFallback: false));
                                }
                        }
                        }
                        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException)
                        {
                            // Ignore fonts outside the minimal parser's current scope.
                        }
                    }
                }

                FontFaceResolution[] faces = fonts
                    .OrderBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                var byFamily = new Dictionary<string, FontFaceResolution[]>(StringComparer.OrdinalIgnoreCase);
                foreach (IGrouping<string, FontFaceResolution> group in faces.GroupBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase))
                {
                    byFamily[group.Key] = group.ToArray();
                }

                return new DiscoverySnapshot(faces, byFamily, faces.Where(f => !f.HasMathTable).ToArray());
            }
            }
        }
    }


    private static (byte[] Bytes, bool Complete, long FileLength) ReadDiscoveryBytes(string path)
    {
        byte[] header = new byte[12];
        long fileLength;
        using (FileStream stream = File.OpenRead(path))
        {
            fileLength = stream.Length;
            // Same per-font ceiling as first use (FileFontProgramSource): oversized
            // files are skipped at stat time instead of buffered.
            FileFontProgramSource.CheckLocalFontSize(fileLength, path);
            if (fileLength < 12 || fileLength > int.MaxValue)
            {
                return (File.ReadAllBytes(path), true, fileLength);
            }

            ReadExactly(stream, header, 0, header.Length);
            if (OpenTypeFont.IsTrueTypeCollectionHeader(header))
            {
                // R13: collections read the header, face offsets, face directories,
                // and needed table extents instead of the whole program per file.
                if (TryReadCollectionDiscoverySpan(stream, header, fileLength, out byte[]? span))
                {
                    return (span, false, fileLength);
                }

                return (File.ReadAllBytes(path), true, fileLength);
            }

            ushort tableCount = OpenTypeFont.U16(header, 4);
            if (tableCount != 0 && tableCount <= 256)
            {
                int directoryLength = tableCount * 16;
                var prefix = new byte[12 + directoryLength];
                Buffer.BlockCopy(header, 0, prefix, 0, header.Length);
                ReadExactly(stream, prefix, header.Length, directoryLength);
                if (OpenTypeFont.TryGetDiscoveryByteBudget(prefix, fileLength, out long requiredEnd))
                {
                    if (requiredEnd <= prefix.Length)
                    {
                        return (prefix, false, fileLength);
                    }

                    if (requiredEnd <= fileLength)
                    {
                        var span = new byte[(int)requiredEnd];
                        Buffer.BlockCopy(prefix, 0, span, 0, prefix.Length);
                        ReadExactly(stream, span, prefix.Length, (int)requiredEnd - prefix.Length);
                        return (span, false, fileLength);
                    }
                }
            }
        }

        return (File.ReadAllBytes(path), true, fileLength);
    }

    private static bool TryReadCollectionDiscoverySpan(FileStream stream, byte[] header, long fileLength, [NotNullWhen(true)] out byte[]? span)
    {
        span = null;
        try
        {
            uint faceCount = OpenTypeFont.U32(header, 8);
            if (faceCount == 0 || faceCount > 256)
            {
                return false;
            }

            int offsetTableLength = checked((int)faceCount * 4);
            long offsetsEnd = checked(12L + offsetTableLength);
            if (offsetsEnd > fileLength)
            {
                return false;
            }

            var offsets = new byte[offsetTableLength];
            ReadExactly(stream, offsets, 0, offsets.Length);
            long requiredEnd = offsetsEnd;
            var faceWindow = new byte[12];
            for (int faceIndex = 0; faceIndex < faceCount; faceIndex++)
            {
                long faceOffset = OpenTypeFont.U32(offsets, faceIndex * 4);
                if (faceOffset > fileLength - 12)
                {
                    return false;
                }

                stream.Seek(faceOffset, SeekOrigin.Begin);
                ReadExactly(stream, faceWindow, 0, faceWindow.Length);
                ushort faceTableCount = OpenTypeFont.U16(faceWindow, 4);
                if (faceTableCount == 0 || faceTableCount > 256)
                {
                    return false;
                }

                int directoryLength = checked(faceTableCount * 16);
                if (faceOffset > fileLength - 12 - directoryLength)
                {
                    return false;
                }

                var window = new byte[12 + directoryLength];
                Buffer.BlockCopy(faceWindow, 0, window, 0, faceWindow.Length);
                stream.Seek(faceOffset + 12, SeekOrigin.Begin);
                ReadExactly(stream, window, 12, directoryLength);
                if (!OpenTypeFont.TryGetTableDirectoryByteBudget(window, headerOffset: 0, fileLength, out long faceEnd))
                {
                    return false;
                }

                requiredEnd = Math.Max(requiredEnd, faceEnd);
            }

            if (requiredEnd > fileLength || requiredEnd > int.MaxValue)
            {
                return false;
            }

            stream.Seek(0, SeekOrigin.Begin);
            var prefix = new byte[(int)requiredEnd];
            ReadExactly(stream, prefix, 0, prefix.Length);
            span = prefix;
            return true;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException or NotSupportedException)
        {
            span = null;
            return false;
        }
    }

    private static void ReadExactly(Stream stream, byte[] buffer, int offset, int count)
    {
        while (count > 0)
        {
            int read = stream.Read(buffer, offset, count);
            if (read == 0)
            {
                throw new EndOfStreamException("Font file ended unexpectedly during discovery.");
            }

            offset += read;
            count -= read;
        }
    }

    /// <summary>
    /// Drops every shared discovery snapshot, so resolvers created afterwards
    /// rediscover installed fonts. Existing instances keep their snapshot.
    /// </summary>
    public static void InvalidateDiscoveryCaches()
    {
        lock (CacheLock)
        {
            DiscoveryCaches.Clear();
        }
    }

    public FontFaceResolution Resolve(FontRequest request)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.FamilyName);
        if (requestCache.TryGetValue(request, out FontFaceResolution? cached))
        {
            return RepopulateRequestedFamily(cached, request);
        }

        FontFaceResolution resolved = ResolveCore(request);
        if (requestCache.Count >= MaxCachedResolutions)
        {
            requestCache.Clear();
        }

        requestCache[request] = resolved;
        return RepopulateRequestedFamily(resolved, request);
    }

    private static FontFaceResolution RepopulateRequestedFamily(FontFaceResolution cached, FontRequest request)
    {
        // Fallback results bake in the first-seen request spelling; re-stamp the current
        // spelling so shared resolvers report deterministically per request. Exact hits
        // already carry discovery spelling and return allocation-free.
        return cached.IsFallback && !string.Equals(cached.RequestedFamily, request.FamilyName, StringComparison.Ordinal)
            ? cached with { RequestedFamily = request.FamilyName }
            : cached;
    }

    private FontFaceResolution ResolveCore(FontRequest request)
    {
        DiscoverySnapshot snapshot = cache.Value;
        if (snapshot.ByFamily.TryGetValue(request.FamilyName, out FontFaceResolution[]? exact) && exact.Length != 0)
        {
            return SelectBest(exact, request);
        }

        if (snapshot.TextFaces.Length != 0)
        {
            return SelectBest(snapshot.TextFaces, request) with { RequestedFamily = request.FamilyName, IsFallback = true };
        }

        return snapshot.Faces.FirstOrDefault() is { } first
            ? first with { RequestedFamily = request.FamilyName, IsFallback = true }
            : new FontFaceResolution(
                request.FamilyName,
                request.FamilyName,
                new FontStyleKey(request.Bold, request.Italic, 400, 0, false),
                new MemoryFontProgramSource("missing:" + request.FamilyName, ReadOnlyMemory<byte>.Empty),
                IsFallback: true);
    }

    internal FontFaceResolution ResolvePresentationTextFace(FontRequest request)
    {
        return Resolve(request);
    }

    internal IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        return cache.Value.Faces;
    }

    IReadOnlyList<FontFaceResolution> IFontCatalog.GetDiscoveredFonts()
    {
        return GetDiscoveredFonts();
    }


    private static IReadOnlyList<string> GetDefaultFontDirectories()
    {
        string windowsFonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string cloudFonts = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft",
            "FontCache",
            "4",
            "CloudFonts");
        return [windowsFonts, cloudFonts];
    }

    private static FontFaceResolution SelectBest(IReadOnlyList<FontFaceResolution> candidates, FontRequest request)
    {
        // single-pass minimum instead of a five-key sort per resolve. OrderBy
        // is stable and First takes the earliest minimum; a strict-less-than scan keeps
        // the first minimal candidate, which is exactly equivalent.
        if (candidates.Count == 0)
        {
            throw new InvalidOperationException("Sequence contains no elements.");
        }

        int targetWeight = request.Bold ? 700 : 400;
        FontFaceResolution best = candidates[0];
        for (int i = 1; i < candidates.Count; i++)
        {
            if (CompareCandidates(candidates[i], best, request, targetWeight) < 0)
            {
                best = candidates[i];
            }
        }

        return best;
    }

    private static int CompareCandidates(FontFaceResolution left, FontFaceResolution right, FontRequest request, int targetWeight)
    {
        // D02: shared style/weight mechanics; family/source/face tie-breakers below
        // stay Windows-specific.
        int result = FontCandidateScoring.CompareStyleWeight(left.Italic, left.WeightClass, right.Italic, right.WeightClass, request.Italic, targetWeight);
        if (result != 0)
        {
            return result;
        }

        result = string.Compare(left.FamilyName, right.FamilyName, StringComparison.OrdinalIgnoreCase);
        if (result != 0)
        {
            return result;
        }

        result = string.Compare(left.Source.StableId, right.Source.StableId, StringComparison.OrdinalIgnoreCase);
        if (result != 0)
        {
            return result;
        }

        return left.FontFaceIndex.CompareTo(right.FontFaceIndex);
    }
}
