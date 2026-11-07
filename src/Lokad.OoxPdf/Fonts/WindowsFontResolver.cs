namespace Lokad.OoxPdf.Fonts;

/// <summary>
/// Resolves installed system fonts using a shared snapshot or an explicitly
/// owned, host-funded catalog.
/// </summary>
/// <remarks>
/// Retention contract: the snapshot retains one lazy <see cref="FileFontProgramSource"/>
/// per discovered file (shared by all its faces). The synchronized, bounded
/// FontProgramCache owns loaded program arrays; sources identify entries and
/// conversions borrow immutable bytes. Individual file bounds apply before loading;
/// LRU eviction bounds cached bytes, so later requests can re-read evicted programs.
/// Hosts that rotate font directories or must
/// release memory call <see cref="InvalidateDiscoveryCaches"/>; existing resolver
/// instances keep their snapshot.
/// The cache-budget factory owns an independent catalog; dispose that resolver
/// after its users finish. Existing program leases retain their own lifetime.
/// </remarks>
public sealed class WindowsFontResolver : IFontResolver, IFontCatalog, IDisposable
{
    private static readonly object CacheLock = new();
    private static readonly Dictionary<string, DiscoveryState> DiscoveryCaches = new(StringComparer.OrdinalIgnoreCase);
    private DiscoverySnapshot? cache;
    private readonly IMemoryReservation? ownedMetadata;
    private readonly FontProgramCache? ownedPrograms;
    private readonly long catalogCapacity;
    private int requestCapacity;

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

    private readonly Dictionary<FontRequest, FontFaceResolution> requestCache = new(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);

    private sealed record DiscoverySnapshot(
        IReadOnlyList<FontFaceResolution> Faces,
        Dictionary<string, FontFaceResolution[]> ByFamily,
        FontFaceResolution[] TextFaces);

    // Only initialization performs file reads. A canceled initializer publishes
    // nothing; another caller can retry the same directory set with its own token.
    private sealed class DiscoveryState(IReadOnlyList<string> directories)
    {
        private readonly SemaphoreSlim gate = new(1, 1);
        private DiscoverySnapshot? snapshot;

        public async ValueTask<DiscoverySnapshot> GetAsync(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (Volatile.Read(ref snapshot) is { } ready) return ready;
            await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                if (snapshot is null)
                {
                    DiscoverySnapshot discovered = await DiscoverAsync(directories, cancellationToken).ConfigureAwait(false);
                    cancellationToken.ThrowIfCancellationRequested();
                    Volatile.Write(ref snapshot, discovered);
                }
                return snapshot;
            }
            finally { gate.Release(); }
        }
    }

    /// <summary>
    /// Prepares installed-font metadata synchronously. Use <see cref="CreateAsync(CancellationToken)"/>
    /// to await font-file reads. Resolution afterwards uses only this snapshot.
    /// </summary>
    public WindowsFontResolver()
        : this(GetOrCreateDiscovery(GetDefaultFontDirectories()).GetAsync(default).GetAwaiter().GetResult())
    {
    }

    internal WindowsFontResolver(string fontsDirectory)
        : this(GetOrCreateDiscovery([fontsDirectory]).GetAsync(default).GetAwaiter().GetResult())
    {
    }

    private WindowsFontResolver(DiscoverySnapshot snapshot) => cache = snapshot;

    /// <summary>Prepares installed-font metadata, awaiting cancellable font-file reads.</summary>
    /// <remarks>
    /// Directory enumeration uses the platform's synchronous directory API.
    /// Concurrent initializers share one published snapshot. Cancellation does not
    /// publish a partial snapshot or cancel another caller's initialization.
    /// </remarks>
    public static ValueTask<WindowsFontResolver> CreateAsync(CancellationToken cancellationToken = default)
        => CreateAsync(GetDefaultFontDirectories(), cancellationToken);

    /// <summary>Prepares an independently owned catalog and bounded program cache funded by the host.</summary>
    /// <remarks>
    /// Keep the resolver alive while using it. Dispose releases catalog retention
    /// and closes the program cache; outstanding program leases remain valid and funded.
    /// The budget belongs to the cache lifetime, independently of conversion budgets.
    /// </remarks>
    public static ValueTask<WindowsFontResolver> CreateAsync(IMemoryBudget cacheBudget, CancellationToken cancellationToken)
        => CreateOwnedAsync(GetDefaultFontDirectories(), cacheBudget, cancellationToken);

    private WindowsFontResolver(DiscoverySnapshot snapshot, IMemoryReservation metadata, FontProgramCache programs)
    {
        cache = snapshot; ownedMetadata = metadata; ownedPrograms = programs; catalogCapacity = metadata.Bytes;
    }

    private static async ValueTask<WindowsFontResolver> CreateOwnedAsync(IReadOnlyList<string> directories,
        IMemoryBudget cacheBudget, CancellationToken token)
    {
        ArgumentNullException.ThrowIfNull(cacheBudget);
        token.ThrowIfCancellationRequested();
        IMemoryReservation? metadata = null;
        bool published = false;
        using var construction = OoxMemory.Reserve(cacheBudget, 0, OoxPdfLimitScope.SharedResource);
        var programs = new FontProgramCache(MaxSnapshotRetainedBytes, StringComparer.OrdinalIgnoreCase, cacheBudget);
        try
        {
            metadata = OoxMemory.Reserve(cacheBudget, 0, OoxPdfLimitScope.SharedResource)!;
            var snapshot = await DiscoverAsync(directories, token, programs, metadata, construction, cacheBudget).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            var result = new WindowsFontResolver(snapshot, metadata, programs);
            published = true;
            return result;
        }
        catch (OoxPdfLimitExceededException ex) when (ex.Kind == OoxPdfLimitKind.MemoryAdmission && ex.Scope != OoxPdfLimitScope.SharedResource)
        {
            throw new OoxPdfLimitExceededException("Host refused shared-memory capacity for font discovery.",
                OoxPdfLimitKind.MemoryAdmission, OoxPdfLimitScope.SharedResource, ex);
        }
        finally
        {
            if (!published) { programs.Dispose(); metadata?.Dispose(); }
        }
    }

    private static async ValueTask<WindowsFontResolver> CreateAsync(IReadOnlyList<string> fontDirectories, CancellationToken cancellationToken)
        => new(await GetOrCreateDiscovery(fontDirectories).GetAsync(cancellationToken).ConfigureAwait(false));

    private static DiscoveryState GetOrCreateDiscovery(IReadOnlyList<string> fontDirectories)
    {
        string cacheKey = string.Join("|", fontDirectories
            .Where(d => !string.IsNullOrWhiteSpace(d))
            .Select(Path.GetFullPath).Order(StringComparer.OrdinalIgnoreCase));
        lock (CacheLock)
        {
            if (!DiscoveryCaches.TryGetValue(cacheKey, out DiscoveryState? state))
            {
                state = new DiscoveryState(fontDirectories);
                DiscoveryCaches.Add(cacheKey, state);
            }
            return state;
        }
    }

    private static async ValueTask<DiscoverySnapshot> DiscoverAsync(IReadOnlyList<string> fontDirectories, CancellationToken cancellationToken,
        FontProgramCache? programs = null, IMemoryReservation? metadata = null,
        IMemoryReservation? construction = null, IMemoryBudget? readBudget = null)
    {
        var fonts = new List<FontFaceResolution>();
        var aliases = new Dictionary<string, List<FontFaceResolution>>(StringComparer.OrdinalIgnoreCase);
        var sourcesByPath = new Dictionary<string, FileFontProgramSource>(StringComparer.OrdinalIgnoreCase);
        var retention = programs ?? new FontProgramCache(MaxSnapshotRetainedBytes, StringComparer.OrdinalIgnoreCase);
        foreach (string fontsDirectory in fontDirectories.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            SearchOption searchOption = fontsDirectory.Contains("CloudFonts", StringComparison.OrdinalIgnoreCase)
                ? SearchOption.AllDirectories
                : SearchOption.TopDirectoryOnly;
            cancellationToken.ThrowIfCancellationRequested();
            foreach (string path in Directory.EnumerateFiles(fontsDirectory, "*.*", searchOption)
                         .Where(p => p.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) ||
                             p.EndsWith(".otf", StringComparison.OrdinalIgnoreCase) ||
                             p.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
                         .Order(StringComparer.OrdinalIgnoreCase))
            {
                cancellationToken.ThrowIfCancellationRequested();
                // Discovery buffers are invocation-local. The static catalog and
                // lazy program cache do not inherit this short-lived reservation.
                using IMemoryReservation? discovery = OoxMemory.Reserve(readBudget ?? OoxConversionBudget.Current?.Memory, 0);
                try
                {
                    // Read the prefix covering discovery tables. This can include
                    // unrelated tables between them; trailing program data stays on disk.
                    (byte[] bytes, bool complete, long fileLength) = await ReadDiscoveryBytesAsync(path, false, discovery, cancellationToken).ConfigureAwait(false);
                    bool isCollection = OpenTypeFont.IsTrueTypeCollectionHeader(bytes);
                    int faceCount = OpenTypeFont.GetCollectionFontCount(bytes);
                    for (int faceIndex = 0; faceIndex < faceCount; faceIndex++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        long precedingCapacity = metadata?.Bytes ?? 0;
                        bool retained = false;
                        try
                        {
                            AdmitParsingWork();
                            OpenTypeFont.FontDiscoveryHeaders headers;
                            try
                            {
                                // Collections parse each face in place instead of
                                // repackaging per-face sfnt copies (G02).
                                headers = isCollection
                                    ? OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(bytes, faceIndex, complete ? null : fileLength, metadata is null ? null : AdmitMetadata)
                                    : OpenTypeFont.ReadDiscoveryHeaders(bytes, faceIndex, complete ? null : fileLength, metadata is null ? null : AdmitMetadata);
                            }
                            catch (Exception spanEx) when (!complete && spanEx is InvalidDataException or ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException)
                            {
                                // A span read can truncate bytes a malformed font's
                                // parser would otherwise touch. Re-parse the full file
                                // so accept/reject matches whole-program reads exactly.
                                // Malformed files pay one extra bounded read; valid
                                // fonts never take this path.
                                (bytes, _, _) = await ReadDiscoveryBytesAsync(path, true, discovery, cancellationToken).ConfigureAwait(false);
                                complete = true;
                                AdmitParsingWork();
                                headers = OpenTypeFont.ReadDiscoveryHeaders(bytes, faceIndex, null, metadata is null ? null : AdmitMetadata);
                            }
                            if (!string.IsNullOrWhiteSpace(headers.FamilyName))
                            {
                                if (!sourcesByPath.TryGetValue(path, out FileFontProgramSource? source))
                                {
                                    source = new FileFontProgramSource(path, retention);
                                    sourcesByPath[path] = source;
                                }
                                var face = new FontFaceResolution(
                                    headers.FamilyName,
                                    headers.FamilyName,
                                    new FontStyleKey(
                                        Bold: headers.WeightClass >= 600,
                                        Italic: Math.Abs(headers.ItalicAngle) > 0.01d,
                                        WeightClass: headers.WeightClass,
                                        FaceIndex: faceIndex,
                                        HasMathTable: headers.HasMathTable),
                                    source,
                                    IsFallback: false);
                                fonts.Add(face);
                                foreach (string family in headers.FamilyNames) {
                                    if (!aliases.TryGetValue(family, out List<FontFaceResolution>? familyFaces)) {
                                        familyFaces = new();
                                        aliases.Add(family, familyFaces);
                                    }
                                    familyFaces.Add(face);
                                }
                                retained = true;
                            }
                        }
                        finally
                        {
                            if (!retained && metadata is not null)
                                OoxMemory.Resize(metadata, precedingCapacity, OoxPdfLimitScope.SharedResource);
                        }

                        void AdmitMetadata(long capacity)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            long next = checked(precedingCapacity + capacity + 2L * path.Length +
                                (precedingCapacity == 0 ? 64L * 1024 : 0));
                            OoxMemory.Resize(metadata, next, OoxPdfLimitScope.SharedResource);
                            AdmitParsingWork();
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        void AdmitParsingWork()
                        {
                            if (construction is null) return;
                            // Validation can materialize maps or a repackaged
                            // collection face before names are decoded.
                            OoxMemory.Resize(construction, checked((metadata?.Bytes ?? 0) + 2L * bytes.LongLength + 128L * 1024),
                                OoxPdfLimitScope.SharedResource);
                            cancellationToken.ThrowIfCancellationRequested();
                        }
                    }
                }
                catch (Exception ex) when (ex is not OoxPdfLimitExceededException &&
                    (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException))
                {
                    // Ignore fonts outside the minimal parser's current scope.
                }
            }
        }

        FontFaceResolution[] faces = fonts
            .OrderBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var byFamily = aliases.ToDictionary(pair => pair.Key, pair => pair.Value.ToArray(), StringComparer.OrdinalIgnoreCase);

        if (metadata is not null)
        {
            // One catalog-family estimate, not a managed-object ledger. Source
            // declarations admitted construction before decoding; publication
            // refunds duplicate/unused names and the old/new index overlap.
            long retained = faces.Length == 0 ? 0 : checked(64L * 1024 + faces.Length * 256L +
                byFamily.Sum(pair => 128L + 2L * pair.Key.Length + 16L * pair.Value.Length) +
                faces.Sum(face => 2L * face.FamilyName.Length) +
                sourcesByPath.Keys.Sum(path => 512L + 2L * path.Length));
            OoxMemory.Resize(metadata, retained, OoxPdfLimitScope.SharedResource);
        }

        return new DiscoverySnapshot(faces, byFamily, faces.Where(f => !f.HasMathTable).ToArray());
    }

    private static async ValueTask<(byte[] Bytes, bool Complete, long FileLength)> ReadDiscoveryBytesAsync(
        string path, bool wholeFile, IMemoryReservation? discovery, CancellationToken cancellationToken)
    {
        (byte[] Bytes, bool Complete, long FileLength) result;
        await using (var stream = new FileStream(path, new FileStreamOptions { Mode = FileMode.Open,
            Access = FileAccess.Read, Share = FileShare.Read, BufferSize = 1,
            Options = FileOptions.Asynchronous | FileOptions.RandomAccess }))
        {
            result = await ReadDiscoveryStreamAsync(stream, path, wholeFile, discovery, cancellationToken).ConfigureAwait(false);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    // Borrows the seekable stream; the path overload owns and closes the file.
    private static async ValueTask<(byte[] Bytes, bool Complete, long FileLength)> ReadDiscoveryStreamAsync(
        Stream stream, string path, bool wholeFile, IMemoryReservation? discovery, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        long fileLength = stream.Length;
        FileFontProgramSource.CheckLocalFontSize(fileLength, path);
        // A parser fallback can still read the previous returned span while the
        // replacement is constructed. Each attempt keeps that existing window.
        long retained = discovery?.Bytes ?? 0;
        if (!wholeFile && fileLength >= 12)
        {
            AdmitDiscoveryBuffer(discovery, retained, 12, cancellationToken);
            byte[] header = new byte[12];
            await stream.ReadExactlyAsync(header, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (OpenTypeFont.IsTrueTypeCollectionHeader(header))
            {
                byte[]? span = await TryReadCollectionDiscoverySpanAsync(stream, header, fileLength, discovery, retained, cancellationToken).ConfigureAwait(false);
                if (span is not null) return (span, false, fileLength);
            }
            else
            {
                ushort tableCount = OpenTypeFont.U16(header, 4);
                if (tableCount != 0 && tableCount <= 256)
                {
                    int directoryLength = tableCount * 16;
                    var prefix = new byte[12 + directoryLength];
                    Buffer.BlockCopy(header, 0, prefix, 0, header.Length);
                    await stream.ReadExactlyAsync(prefix.AsMemory(header.Length, directoryLength), cancellationToken).ConfigureAwait(false);
                    cancellationToken.ThrowIfCancellationRequested();
                    if (OpenTypeFont.TryGetDiscoveryByteBudget(prefix, fileLength, out long requiredEnd))
                    {
                        if (requiredEnd <= prefix.Length) return (prefix, false, fileLength);
                        if (requiredEnd <= fileLength)
                        {
                            AdmitDiscoveryBuffer(discovery, retained, requiredEnd, cancellationToken);
                            var span = new byte[(int)requiredEnd];
                            Buffer.BlockCopy(prefix, 0, span, 0, prefix.Length);
                            await stream.ReadExactlyAsync(span.AsMemory(prefix.Length), cancellationToken).ConfigureAwait(false);
                            cancellationToken.ThrowIfCancellationRequested();
                            return (span, false, fileLength);
                        }
                    }
                }
            }
        }

        // Reuse this bounded open-file length even on malformed metadata fallback;
        // reopening by path could race replacement/growth and lose the size bound.
        stream.Seek(0, SeekOrigin.Begin);
        AdmitDiscoveryBuffer(discovery, retained, fileLength, cancellationToken);
        byte[] bytes = new byte[checked((int)fileLength)];
        await stream.ReadExactlyAsync(bytes, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return (bytes, true, fileLength);
    }

    private static async ValueTask<byte[]?> TryReadCollectionDiscoverySpanAsync(Stream stream, byte[] header, long fileLength,
        IMemoryReservation? discovery, long retained, CancellationToken cancellationToken)
    {
        try
        {
            uint faceCount = OpenTypeFont.U32(header, 8);
            if (faceCount == 0 || faceCount > 256)
            {
                return null;
            }

            int offsetTableLength = checked((int)faceCount * 4);
            long offsetsEnd = checked(12L + offsetTableLength);
            if (offsetsEnd > fileLength)
            {
                return null;
            }

            var offsets = new byte[offsetTableLength];
            await stream.ReadExactlyAsync(offsets, cancellationToken).ConfigureAwait(false);
            long requiredEnd = offsetsEnd;
            var faceWindow = new byte[12];
            for (int faceIndex = 0; faceIndex < faceCount; faceIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                long faceOffset = OpenTypeFont.U32(offsets, faceIndex * 4);
                if (faceOffset > fileLength - 12)
                {
                    return null;
                }

                stream.Seek(faceOffset, SeekOrigin.Begin);
                await stream.ReadExactlyAsync(faceWindow, cancellationToken).ConfigureAwait(false);
                ushort faceTableCount = OpenTypeFont.U16(faceWindow, 4);
                if (faceTableCount == 0 || faceTableCount > 256)
                {
                    return null;
                }

                int directoryLength = checked(faceTableCount * 16);
                if (faceOffset > fileLength - 12 - directoryLength)
                {
                    return null;
                }

                var window = new byte[12 + directoryLength];
                Buffer.BlockCopy(faceWindow, 0, window, 0, faceWindow.Length);
                stream.Seek(faceOffset + 12, SeekOrigin.Begin);
                await stream.ReadExactlyAsync(window.AsMemory(12, directoryLength), cancellationToken).ConfigureAwait(false);
                if (!OpenTypeFont.TryGetTableDirectoryByteBudget(window, headerOffset: 0, fileLength, out long faceEnd))
                {
                    return null;
                }

                requiredEnd = Math.Max(requiredEnd, faceEnd);
            }

            if (requiredEnd > fileLength || requiredEnd > int.MaxValue)
            {
                return null;
            }

            stream.Seek(0, SeekOrigin.Begin);
            AdmitDiscoveryBuffer(discovery, retained, requiredEnd, cancellationToken);
            var prefix = new byte[(int)requiredEnd];
            await stream.ReadExactlyAsync(prefix, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return prefix;
        }
        catch (Exception ex) when (ex is not OoxPdfLimitExceededException &&
            (ex is IOException or InvalidDataException or ArgumentOutOfRangeException or IndexOutOfRangeException or OverflowException or ArgumentException or NotSupportedException))
        {
            return null;
        }
    }

    private static void AdmitDiscoveryBuffer(IMemoryReservation? discovery, long retained, long bytes, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (discovery is not null)
        {
            // Reuse one capacity window across collection directories/faces.
            // Their bounded headers/directories fit in fixed working headroom;
            // selected spans and old/new replacement overlap add actual lengths.
            long capacity = checked(retained + bytes + 128L * 1024);
            if (capacity > discovery.Bytes) OoxMemory.Resize(discovery, capacity);
        }
        token.ThrowIfCancellationRequested();
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
        // Hits, publication and capacity eviction share ownership of the dictionary.
        lock (requestCache)
        {
            ObjectDisposedException.ThrowIf(cache is null, this);
            // An empty catalog has no useful program/style choice to memoize.
            if (cache.Faces.Count == 0) return ResolveCore(request);
            if (requestCache.TryGetValue(request, out FontFaceResolution? cached))
            {
                return RepopulateRequestedFamily(cached, request);
            }

            FontFaceResolution resolved = ResolveCore(request);
            if (requestCache.Count >= MaxCachedResolutions)
            {
                requestCache.Clear();
            }

            if (ownedMetadata is not null && requestCache.Count >= requestCapacity)
            {
                int capacity = Math.Min(MaxCachedResolutions, Math.Max(256, requestCapacity * 2));
                long retained = checked(catalogCapacity + capacity * 256L);
                long prior = ownedMetadata.Bytes;
                try
                {
                    OoxMemory.Resize(ownedMetadata, checked(catalogCapacity + capacity * 512L), OoxPdfLimitScope.SharedResource);
                    requestCache.EnsureCapacity(capacity);
                    requestCapacity = capacity;
                }
                finally { OoxMemory.Resize(ownedMetadata, requestCapacity == capacity ? retained : prior, OoxPdfLimitScope.SharedResource); }
            }
            requestCache[request] = resolved;
            return RepopulateRequestedFamily(resolved, request);
        }
    }

    private static FontFaceResolution RepopulateRequestedFamily(FontFaceResolution cached, FontRequest request)
    {
        // Fallback and alias requests keep caller spelling. Canonical exact matches
        // retain discovery spelling and return allocation-free.
        return (cached.IsFallback || !string.Equals(cached.ResolvedFamily, request.FamilyName, StringComparison.OrdinalIgnoreCase)) &&
            !string.Equals(cached.RequestedFamily, request.FamilyName, StringComparison.Ordinal)
            ? cached with { RequestedFamily = request.FamilyName }
            : cached;
    }

    private FontFaceResolution ResolveCore(FontRequest request)
    {
        DiscoverySnapshot snapshot = cache ?? throw new ObjectDisposedException(nameof(WindowsFontResolver));
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

    internal IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        lock (requestCache)
        {
            ObjectDisposedException.ThrowIf(cache is null, this);
            // Owned-provider results are caller values, not a borrowed cache array.
            return ownedMetadata is null ? cache.Faces : cache.Faces.ToArray();
        }
    }

    /// <summary>Releases this resolver; shared default snapshots retain their independent lifetime.</summary>
    public void Dispose()
    {
        lock (requestCache)
        {
            if (cache is null) return;
            cache = null; requestCache.Clear(); requestCache.TrimExcess();
        }
        ownedPrograms?.Dispose();
        ownedMetadata?.Dispose();
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
