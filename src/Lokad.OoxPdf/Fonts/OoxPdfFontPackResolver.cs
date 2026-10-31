using System.Security.Cryptography;
using System.Text.Json;

namespace Lokad.OoxPdf.Fonts;

public sealed class OoxPdfFontPackResolver : IFontResolver, IFontCatalog, IDisposable
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    internal const long MaxManifestBytes = 1024L * 1024L;
    internal const long MaxFontFileBytes = 64L * 1024L * 1024L;

    // Cache retention is LRU-bounded. Leases keep evicted or shutdown programs
    // and their funding alive until the final reader. Payload limits separately
    // bound downloads; acquisition buffers and replacement overlap are admitted.
    internal const long MaxTotalFontPackBytes = 256L * 1024L * 1024L;
    private readonly FontPackFileSource fileSource;
    private bool disposed;
    private IReadOnlyList<FontPackFace> faces;
    private IReadOnlyList<string> fallbackFamilies;
    private readonly Dictionary<string, FontPackFace[]> facesByFamily;
    private FontPackFace[] textFaces;
    private readonly Dictionary<FontRequest, FontFaceResolution> requestCache = new(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);

    private OoxPdfFontPackResolver(
        string packId,
        Uri packRootUri,
        HttpClient httpClient,
        IReadOnlyList<FontPackFace> faces,
        IReadOnlyList<string> fallbackFamilies,
        IMemoryBudget? cacheBudget, IMemoryReservation? metadataStorage, CancellationToken token)
    {
        PackId = packId;
        PackRootUri = packRootUri;
        this.faces = faces;
        this.fallbackFamilies = fallbackFamilies;
        facesByFamily = IndexFacesByFamily(faces, token);
        textFaces = faces.Where(face => !face.Style.HasMathTable).ToArray();
        token.ThrowIfCancellationRequested();
        fileSource = new FontPackFileSource(packId, packRootUri, httpClient, MaxTotalFontPackBytes, cacheBudget, metadataStorage);
    }

    private static Dictionary<string, FontPackFace[]> IndexFacesByFamily(IReadOnlyList<FontPackFace> faces, CancellationToken token)
    {
        var grouped = new Dictionary<string, List<FontPackFace>>(StringComparer.OrdinalIgnoreCase);
        foreach (FontPackFace face in faces)
        {
            token.ThrowIfCancellationRequested();
            if (!grouped.TryGetValue(face.RequestedFamily, out List<FontPackFace>? requested))
            {
                requested = [];
                grouped[face.RequestedFamily] = requested;
            }

            requested.Add(face);
            if (!face.ResolvedFamily.Equals(face.RequestedFamily, StringComparison.OrdinalIgnoreCase))
            {
                if (!grouped.TryGetValue(face.ResolvedFamily, out List<FontPackFace>? resolved))
                {
                    resolved = [];
                    grouped[face.ResolvedFamily] = resolved;
                }

                resolved.Add(face);
            }
        }

        var index = new Dictionary<string, FontPackFace[]>(StringComparer.OrdinalIgnoreCase);
        foreach ((string family, List<FontPackFace> group) in grouped)
        {
            token.ThrowIfCancellationRequested();
            index[family] = group.ToArray();
        }

        return index;
    }

    public string PackId { get; }

    public Uri PackRootUri { get; }

    /// <summary>Loads pack metadata; programs are acquired lazily.</summary>
    /// <remarks>Cache admission is borrowed until shutdown and outstanding leases/I/O finish. HttpClient remains caller-owned.</remarks>
    public static async Task<OoxPdfFontPackResolver> CreateHttpAsync(
        string packId,
        Uri sourceUri,
        HttpClient httpClient,
        IMemoryBudget? cacheBudget = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(packId);
        ArgumentNullException.ThrowIfNull(sourceUri);
        ArgumentNullException.ThrowIfNull(httpClient);
        cancellationToken.ThrowIfCancellationRequested();

        ValidatePackId(packId);
        Uri sourceRootUri = NormalizeSourceUri(sourceUri);
        Uri packRootUri = new(sourceRootUri, packId + "/");
        Uri manifestUri = new(packRootUri, "manifest.json");

        using FontProgramBuffer manifestBytes = await DownloadBytesAsync(
            httpClient,
            manifestUri,
            "font pack manifest",
            MaxManifestBytes,
            cacheBudget,
            OoxPdfLimitScope.SharedResource,
            cancellationToken).ConfigureAwait(false);
        // Coarse metadata/index capacity plus the bounded request-cache window.
        // Deserialization, validation and indexing coexist before publication;
        // their extra window refunds once the resolver owns the retained graph.
        long retainedCapacity = checked(1024L * 1024 + 16L * manifestBytes.Bytes.Length);
        IMemoryReservation? metadataStorage = OoxMemory.Reserve(cacheBudget, checked(2 * retainedCapacity), OoxPdfLimitScope.SharedResource);
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            FontPackManifest manifest = DeserializeManifest();
            ValidatedManifest validated = ValidateManifest(packId, manifest, cancellationToken);
            IMemoryReservation? owned = metadataStorage;
            var resolver = new OoxPdfFontPackResolver(packId, packRootUri, httpClient,
                validated.Faces, validated.FallbackFamilies, cacheBudget, owned, cancellationToken);
            metadataStorage = null; // The existing file-source owner now holds it.
            try
            {
                OoxMemory.Resize(owned, retainedCapacity, OoxPdfLimitScope.SharedResource);
                cancellationToken.ThrowIfCancellationRequested();
                return resolver;
            }
            catch { resolver.Dispose(); throw; }
        }
        finally { metadataStorage?.Dispose(); }

        FontPackManifest DeserializeManifest()
        {
            try
            {
                return JsonSerializer.Deserialize<FontPackManifest>(manifestBytes.Bytes.Span, JsonOptions)
                    ?? throw new OoxPdfFontPackException(
                        OoxPdfFontPackDiagnosticIds.FontPackInvalid,
                        "The font pack manifest is empty.");
            }
            catch (JsonException ex)
            {
                throw new OoxPdfFontPackException(
                    OoxPdfFontPackDiagnosticIds.FontPackInvalid,
                    "The font pack manifest is not valid JSON.",
                    ex);
            }
        }
    }

    public FontFaceResolution Resolve(FontRequest request)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.FamilyName);
        // Hits, publication and capacity eviction share ownership of the dictionary.
        lock (requestCache)
        {
            ObjectDisposedException.ThrowIf(disposed, this);
            if (requestCache.TryGetValue(request, out FontFaceResolution? cached))
            {
                return RepopulateRequestedFamily(cached, request);
            }

            FontFaceResolution resolved = ResolveCore(request);
            if (requestCache.Count >= WindowsFontResolver.MaxCachedResolutions)
            {
                requestCache.Clear();
            }

            requestCache[request] = resolved;
            return RepopulateRequestedFamily(resolved, request);
        }
    }

    private static FontFaceResolution RepopulateRequestedFamily(FontFaceResolution cached, FontRequest request)
    {
        return cached.IsFallback && !string.Equals(cached.RequestedFamily, request.FamilyName, StringComparison.Ordinal)
            ? cached with { RequestedFamily = request.FamilyName }
            : cached;
    }

    private FontFaceResolution ResolveCore(FontRequest request)
    {
        if (facesByFamily.TryGetValue(request.FamilyName, out FontPackFace[]? exact) && exact.Length != 0)
        {
            return SelectBest(exact, request, isFallback: false);
        }

        foreach (string fallback in fallbackFamilies)
        {
            if (facesByFamily.TryGetValue(fallback, out FontPackFace[]? configuredFallbacks) && configuredFallbacks.Length != 0)
            {
                return SelectBest(configuredFallbacks, request, isFallback: true);
            }
        }

        return SelectBest(textFaces.Length == 0 ? faces : textFaces, request, isFallback: true);
    }

    /// <summary>Returns caller-owned resolution values; their sources refuse new acquisition after shutdown.</summary>
    public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        lock (requestCache)
        {
            ObjectDisposedException.ThrowIf(disposed, this);
            return faces.Select(face => face.ToResolution(fileSource, isFallback: false)).ToArray();
        }
    }

    /// <summary>Ends cache retention and refuses new acquisition; active leases remain valid.</summary>
    /// <remarks>Keep the borrowed cache budget alive until in-flight acquisitions and leases finish. HttpClient remains caller-owned.</remarks>
    public void Dispose()
    {
        lock (requestCache)
        {
            if (disposed) return;
            disposed = true;
            requestCache.Clear(); requestCache.TrimExcess();
            facesByFamily.Clear(); facesByFamily.TrimExcess();
            faces = []; textFaces = []; fallbackFamilies = [];
        }
        fileSource.Dispose();
    }

    private FontFaceResolution SelectBest(IReadOnlyList<FontPackFace> candidates, FontRequest request, bool isFallback)
    {
        // single-pass minimum instead of a five-key sort per resolve (see
        // WindowsFontResolver.SelectBest for the equivalence argument).
        if (candidates.Count == 0)
        {
            throw new InvalidOperationException("Sequence contains no elements.");
        }

        int targetWeight = request.Bold ? 700 : 400;
        FontPackFace face = candidates[0];
        for (int i = 1; i < candidates.Count; i++)
        {
            if (CompareCandidates(candidates[i], face, request, targetWeight) < 0)
            {
                face = candidates[i];
            }
        }

        return face.ToResolution(fileSource, isFallback) with
        {
            RequestedFamily = request.FamilyName
        };
    }

    private static int CompareCandidates(FontPackFace left, FontPackFace right, FontRequest request, int targetWeight)
    {
        // D02: shared style/weight mechanics; family/file/face tie-breakers below
        // stay pack-specific.
        int result = FontCandidateScoring.CompareStyleWeight(left.Style.Italic, left.Style.WeightClass, right.Style.Italic, right.Style.WeightClass, request.Italic, targetWeight);
        if (result != 0)
        {
            return result;
        }

        result = string.Compare(left.ResolvedFamily, right.ResolvedFamily, StringComparison.OrdinalIgnoreCase);
        if (result != 0)
        {
            return result;
        }

        result = string.Compare(left.RelativeFontFile, right.RelativeFontFile, StringComparison.Ordinal);
        if (result != 0)
        {
            return result;
        }

        return left.Style.FaceIndex.CompareTo(right.Style.FaceIndex);
    }

    private static Uri NormalizeSourceUri(Uri sourceUri)
    {
        if (!sourceUri.IsAbsoluteUri ||
            (sourceUri.Scheme != Uri.UriSchemeHttp && sourceUri.Scheme != Uri.UriSchemeHttps))
        {
            throw new ArgumentException("The font pack source URI must be an absolute HTTP or HTTPS URI.", nameof(sourceUri));
        }

        if (!string.IsNullOrEmpty(sourceUri.Query) || !string.IsNullOrEmpty(sourceUri.Fragment))
        {
            throw new ArgumentException("The font pack source URI cannot include a query string or fragment.", nameof(sourceUri));
        }

        string source = sourceUri.AbsoluteUri;
        return source.EndsWith("/", StringComparison.Ordinal)
            ? sourceUri
            : new Uri(source + "/", UriKind.Absolute);
    }

    private static void ValidatePackId(string packId)
    {
        if (IsUnsafePathSegment(packId))
        {
            throw new ArgumentException("The font pack identifier must be a single relative path segment.", nameof(packId));
        }
    }

    // R13: process-wide bound on concurrent font downloads. Response payloads
    // are capped separately; accumulation and the final copy need headroom beyond
    // payload size. This is not an exact heap ceiling. Retries run
    // inside one slot; cancellation always throws.
    internal const int MaxConcurrentFontDownloads = 4;
    private static readonly SemaphoreSlim DownloadConcurrency = new(MaxConcurrentFontDownloads, MaxConcurrentFontDownloads);

    private static async Task<FontProgramBuffer> DownloadBytesAsync(
        HttpClient httpClient,
        Uri uri,
        string description,
        long maxBytes,
        IMemoryBudget? memoryBudget,
        OoxPdfLimitScope admissionScope,
        CancellationToken cancellationToken)
    {
        // Transient transport failures retry immediately and bounded.
        // Server answers (status, declared size, over-cap bodies) fail fast;
        // cancellation always throws.
        await DownloadConcurrency.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return await DownloadWithRetriesAsync(httpClient, uri, description, maxBytes, memoryBudget, admissionScope, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            DownloadConcurrency.Release();
        }
    }

    private static async Task<FontProgramBuffer> DownloadWithRetriesAsync(
        HttpClient httpClient,
        Uri uri,
        string description,
        long maxBytes,
        IMemoryBudget? memoryBudget,
        OoxPdfLimitScope admissionScope,
        CancellationToken cancellationToken)
    {
        const int maxAttempts = 3;
        for (int attempt = 1; ; attempt++)
        {
            try
            {
                return await DownloadBytesSingleAttemptAsync(httpClient, uri, description, maxBytes, memoryBudget, admissionScope, cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (OoxPdfLimitExceededException) { throw; }
            catch (OoxPdfFontPackException)
            {
                throw;
            }
            catch (Exception ex) when (IsTransientDownloadFailure(ex) && attempt < maxAttempts)
            {
                continue;
            }
            catch (Exception ex) when (ex is HttpRequestException or IOException or InvalidOperationException or TaskCanceledException)
            {
                throw new OoxPdfFontPackException(
                    OoxPdfFontPackDiagnosticIds.FontPackDownloadFailed,
                    $"Unable to download {description} from '{uri}'.",
                    ex);
            }
        }
    }

    private static async Task<FontProgramBuffer> DownloadBytesSingleAttemptAsync(
        HttpClient httpClient,
        Uri uri,
        string description,
        long maxBytes,
        IMemoryBudget? memoryBudget,
        OoxPdfLimitScope admissionScope,
        CancellationToken cancellationToken)
    {
        FontProgramBuffer? result = null;
        try
        {
            using HttpResponseMessage response = await httpClient.GetAsync(
                uri,
                HttpCompletionOption.ResponseHeadersRead,
                cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                throw new OoxPdfFontPackException(
                    OoxPdfFontPackDiagnosticIds.FontPackDownloadFailed,
                    $"Unable to download {description} from '{uri}'. HTTP {(int)response.StatusCode} {response.ReasonPhrase}.");
            }

            if (response.Content.Headers.ContentLength is long declared && declared > maxBytes)
            {
                throw new OoxPdfFontPackException(
                    OoxPdfFontPackDiagnosticIds.FontPackDownloadFailed,
                    $"Unable to download {description} from \u0027{uri}\u0027: declared size {declared} bytes exceeds the limit of {maxBytes} bytes.");
            }

            await using Stream content = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
            result = await CopyCappedAsync(content, maxBytes, description, uri.ToString(), memoryBudget, admissionScope, cancellationToken).ConfigureAwait(false);
        }
        catch { result?.Dispose(); throw; }
        if (cancellationToken.IsCancellationRequested)
        {
            result?.Dispose();
            cancellationToken.ThrowIfCancellationRequested();
        }
        return result!;
    }

    private static bool IsTransientDownloadFailure(Exception ex)
    {
        return ex is HttpRequestException or IOException or InvalidOperationException or TaskCanceledException;
    }

    private static ValidatedManifest ValidateManifest(string packId, FontPackManifest manifest, CancellationToken token)
    {
        if (!packId.Equals(manifest.PackId, StringComparison.Ordinal))
        {
            throw InvalidManifest("The font pack manifest packId does not match the requested pack identifier.");
        }

        if (manifest.Files is null || manifest.Files.Length == 0)
        {
            throw MissingFonts("The font pack manifest does not list any font files.");
        }

        var files = new Dictionary<string, FontPackFile>(StringComparer.Ordinal);
        foreach (FontPackFileManifest? file in manifest.Files)
        {
            token.ThrowIfCancellationRequested();
            if (file is null)
            {
                throw InvalidManifest("The font pack manifest contains an empty font file entry.");
            }

            string relativePath = ValidateRelativePath(file.RelativePath, "font file");
            if (!files.TryAdd(relativePath, new FontPackFile(
                relativePath,
                ValidateByteSize(file.ByteSize, relativePath),
                ValidateSha256(file.Sha256, relativePath))))
            {
                throw InvalidManifest($"The font pack manifest contains duplicate font file path '{relativePath}'.");
            }
        }

        if (manifest.Families is null || manifest.Families.Length == 0)
        {
            throw MissingFonts("The font pack manifest does not list any font families.");
        }

        var faces = new List<FontPackFace>();
        foreach (FontPackFamilyManifest? family in manifest.Families)
        {
            token.ThrowIfCancellationRequested();
            if (family is null)
            {
                throw InvalidManifest("The font pack manifest contains an empty font family entry.");
            }

            string requestedFamily = RequiredText(family.RequestedFamily, "requested font family");
            string resolvedFamily = RequiredText(family.ResolvedFamily, "resolved font family");
            string relativeFontFile = ValidateRelativePath(family.RelativeFontFile, "font family file");
            if (!files.ContainsKey(relativeFontFile))
            {
                throw InvalidManifest($"Font family '{requestedFamily}' references missing file '{relativeFontFile}'.");
            }

            int faceIndex = family.FaceIndex ?? 0;
            if (faceIndex < 0)
            {
                throw InvalidManifest($"Font family '{requestedFamily}' has a negative face index.");
            }

            int weight = family.Weight ?? 400;
            if (weight <= 0)
            {
                throw InvalidManifest($"Font family '{requestedFamily}' has an invalid weight.");
            }

            bool italic = family.Italic ?? false;
            faces.Add(new FontPackFace(
                requestedFamily,
                resolvedFamily,
                relativeFontFile,
                new FontStyleKey(
                    Bold: weight >= 600,
                    Italic: italic,
                    WeightClass: weight,
                    FaceIndex: faceIndex,
                    HasMathTable: family.HasMathTable ?? false),
                files[relativeFontFile]));
        }

        if (faces.Count == 0)
        {
            throw MissingFonts("The font pack manifest does not list any usable font faces.");
        }

        string[] fallbackFamilies = manifest.Fallbacks is null
            ? []
            : manifest.Fallbacks
                .Select(fallback => RequiredText(fallback?.Family, "fallback font family"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

        return new ValidatedManifest(faces.ToArray(), fallbackFamilies);
    }

    private static string RequiredText(string? value, string fieldName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw InvalidManifest($"The font pack manifest has an empty {fieldName}.");
        }

        return value;
    }

    private static string ValidateRelativePath(string? value, string fieldName)
    {
        string relativePath = RequiredText(value, fieldName);
        if (relativePath[0] == '/' ||
            relativePath.Contains('\\', StringComparison.Ordinal) ||
            relativePath.Contains(':', StringComparison.Ordinal) ||
            relativePath.Contains('?', StringComparison.Ordinal) ||
            relativePath.Contains('#', StringComparison.Ordinal) ||
            relativePath.Contains("//", StringComparison.Ordinal))
        {
            throw InvalidManifest($"The font pack manifest has an unsafe {fieldName} path '{relativePath}'.");
        }

        string[] segments = relativePath.Split('/');
        if (segments.Any(IsUnsafePathSegment))
        {
            throw InvalidManifest($"The font pack manifest has an unsafe {fieldName} path '{relativePath}'.");
        }

        return relativePath;
    }

    private static bool IsUnsafePathSegment(string segment)
    {
        if (segment.Length == 0 ||
            segment is "." or ".." ||
            segment.Contains('/', StringComparison.Ordinal) ||
            segment.Contains('\\', StringComparison.Ordinal) ||
            segment.Contains(':', StringComparison.Ordinal) ||
            segment.Contains('?', StringComparison.Ordinal) ||
            segment.Contains('#', StringComparison.Ordinal))
        {
            return true;
        }

        string unescaped;
        try
        {
            unescaped = Uri.UnescapeDataString(segment);
        }
        catch (UriFormatException)
        {
            return true;
        }

        return unescaped is "." or ".." ||
            unescaped.Contains('/', StringComparison.Ordinal) ||
            unescaped.Contains('\\', StringComparison.Ordinal);
    }

    internal static async Task<FontProgramBuffer> CopyCappedAsync(Stream source, long maxBytes, string description,
        string uri, IMemoryBudget? memoryBudget, OoxPdfLimitScope admissionScope, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        const int scratchBytes = 81920;
        IMemoryReservation? storage = OoxMemory.Reserve(memoryBudget, scratchBytes, admissionScope);
        byte[]? buffer = null;
        using var destination = new MemoryStream();
        try
        {
            buffer = new byte[scratchBytes];
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int read = await source.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                if (read == 0) break;
                long length = checked(destination.Length + read);
                if (length > maxBytes)
                    throw new OoxPdfFontPackException(OoxPdfFontPackDiagnosticIds.FontPackDownloadFailed,
                        "Unable to download " + description + " from " + uri + ": response exceeds the limit.");
                if (length > destination.Capacity)
                {
                    int capacity = checked((int)Math.Min(maxBytes, Math.Max(length, Math.Max(256L, (long)destination.Capacity * 2))));
                    OoxMemory.Resize(storage, checked((long)scratchBytes + destination.Capacity + capacity), admissionScope);
                    destination.Capacity = capacity;
                    OoxMemory.Resize(storage, (long)scratchBytes + capacity, admissionScope);
                }
                destination.Write(buffer, 0, read);
            }
            OoxMemory.Resize(storage, checked((long)scratchBytes + destination.Capacity + destination.Length), admissionScope);
            byte[] bytes = destination.ToArray();
            destination.SetLength(0); destination.Capacity = 0; buffer = null;
            OoxMemory.Resize(storage, bytes.Length, admissionScope);
            var result = new FontProgramBuffer(bytes, storage);
            storage = null;
            return result;
        }
        finally
        {
            destination.SetLength(0); destination.Capacity = 0; buffer = null;
            storage?.Dispose();
        }
    }

    private static long ValidateByteSize(long? byteSize, string relativePath)
    {
        if (byteSize is null or < 0)
        {
            throw InvalidManifest($"Font file '{relativePath}' has an invalid byte size.");
        }

        return byteSize.Value;
    }

    private static string ValidateSha256(string? sha256, string relativePath)
    {
        if (sha256 is null ||
            sha256.Length != 64 ||
            sha256.Any(ch => !Uri.IsHexDigit(ch)))
        {
            throw InvalidManifest($"Font file '{relativePath}' has an invalid SHA-256 hash.");
        }

        return sha256.ToUpperInvariant();
    }

    private static OoxPdfFontPackException InvalidManifest(string message)
    {
        return new OoxPdfFontPackException(OoxPdfFontPackDiagnosticIds.FontPackInvalid, message);
    }

    private static OoxPdfFontPackException MissingFonts(string message)
    {
        return new OoxPdfFontPackException(OoxPdfFontPackDiagnosticIds.FontPackMissing, message);
    }

    internal sealed class FontPackFileSource : IDisposable
    {
        private readonly string packId;
        private readonly Uri packRootUri;
        private readonly HttpClient httpClient;
        private readonly FontProgramCache cache;
        private readonly object sync = new();
        private readonly Dictionary<string, IFontProgramSource> sources = new(StringComparer.Ordinal);
        private IMemoryReservation? metadataStorage;
        private int readers = 1; // Resolver retention; acquisitions add temporary readers.
        private bool disposed;

        internal FontPackFileSource(string packId, Uri packRootUri, HttpClient httpClient, long maxTotalBytes = MaxTotalFontPackBytes, IMemoryBudget? cacheBudget = null, IMemoryReservation? metadataStorage = null)
        {
            this.packId = packId;
            this.packRootUri = packRootUri;
            this.httpClient = httpClient;
            cache = new FontProgramCache(maxTotalBytes, StringComparer.Ordinal, cacheBudget);
            this.metadataStorage = metadataStorage;
        }

        public IFontProgramSource Create(FontPackFile file)
        {
            lock (sync)
            {
                ObjectDisposedException.ThrowIf(disposed, this);
                if (!sources.TryGetValue(file.RelativePath, out IFontProgramSource? source))
                {
                    source = new OoxPdfFontPackProgramSource(packId, packRootUri, httpClient, file, this);
                    sources[file.RelativePath] = source;
                }

                return source;
            }
        }

        internal long RetainedBytes => cache.RetainedBytes;

        internal IMemoryBudget? Budget => cache.Budget;
        internal bool TryAcquire(string path, IMemoryBudget? operationBudget, out IOoxPdfFontProgramLease? lease) => cache.TryAcquire(path, operationBudget, out lease);
        internal void Store(string path, FontProgramBuffer program, IMemoryBudget? operationBudget) => cache.Store(path, program, operationBudget);
        internal void AcquireReader()
        {
            lock (sync)
            {
                ObjectDisposedException.ThrowIf(disposed, this);
                readers = checked(readers + 1);
            }
        }

        internal void ReleaseReader()
        {
            IMemoryReservation? retired = null;
            lock (sync)
            {
                if (--readers == 0) { retired = metadataStorage; metadataStorage = null; }
            }
            retired?.Dispose();
        }

        public void Dispose()
        {
            lock (sync)
            {
                if (disposed) return;
                disposed = true;
                sources.Clear(); sources.TrimExcess();
            }
            // Retention remains a reader until cache shutdown finishes. Refunds
            // happen outside locks, after this and any pending acquisitions end.
            try { cache.Dispose(); }
            finally { ReleaseReader(); }
        }
    }

    private sealed class OoxPdfFontPackProgramSource(
        string packId,
        Uri packRootUri,
        HttpClient httpClient,
        FontPackFile file,
        FontPackFileSource owner) : IFontProgramSource
    {
        private readonly SemaphoreSlim gate = new(1, 1);

        public string StableId => "ooxpdf-font-pack:" + packId + ":" + file.Sha256;

        public async ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            owner.AcquireReader();
            try
            {
                if (owner.TryAcquire(file.RelativePath, operationBudget, out var cached))
                {
                    return cached!;
                }

                await gate.WaitAsync(ct).ConfigureAwait(false);
                try
                {
                    if (owner.TryAcquire(file.RelativePath, operationBudget, out var rechecked))
                    {
                        return rechecked!;
                    }

                    using FontProgramBuffer bytes = await DownloadBytesAsync(
                        httpClient,
                        new Uri(packRootUri, file.RelativePath),
                        "font file '" + file.RelativePath + "'",
                        Math.Min(file.ByteSize, MaxFontFileBytes),
                        owner.Budget ?? operationBudget,
                        owner.Budget is null ? OoxPdfLimitScope.Conversion : OoxPdfLimitScope.SharedResource,
                        ct).ConfigureAwait(false);

                    ct.ThrowIfCancellationRequested();
                    ValidateFontBytes(file, bytes.Bytes.Span);
                    owner.Store(file.RelativePath, bytes, operationBudget);
                    return bytes.Acquire();
                }
                finally
                {
                    gate.Release();
                }
            }
            finally { owner.ReleaseReader(); }
        }

        private static void ValidateFontBytes(FontPackFile file, ReadOnlySpan<byte> bytes)
        {
            if (bytes.Length != file.ByteSize)
            {
                throw new OoxPdfFontPackException(
                    OoxPdfFontPackDiagnosticIds.FontPackHashMismatch,
                    $"Font file '{file.RelativePath}' size mismatch: expected {file.ByteSize} bytes, got {bytes.Length} bytes.");
            }

            string actual = Convert.ToHexString(SHA256.HashData(bytes));
            if (!actual.Equals(file.Sha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new OoxPdfFontPackException(
                    OoxPdfFontPackDiagnosticIds.FontPackHashMismatch,
                    $"Font file '{file.RelativePath}' SHA-256 mismatch.");
            }
        }
    }

    private sealed record ValidatedManifest(
        IReadOnlyList<FontPackFace> Faces,
        IReadOnlyList<string> FallbackFamilies);

    internal sealed record FontPackFile(
        string RelativePath,
        long ByteSize,
        string Sha256);

    private sealed record FontPackFace(
        string RequestedFamily,
        string ResolvedFamily,
        string RelativeFontFile,
        FontStyleKey Style,
        FontPackFile File)
    {
        public FontFaceResolution ToResolution(FontPackFileSource source, bool isFallback)
        {
            return new FontFaceResolution(
                RequestedFamily,
                ResolvedFamily,
                Style,
                source.Create(File),
                isFallback);
        }
    }

    private sealed class FontPackManifest
    {
        public string? PackId { get; set; }
        public FontPackFileManifest?[]? Files { get; set; }
        public FontPackFamilyManifest?[]? Families { get; set; }
        public FontPackFallbackManifest?[]? Fallbacks { get; set; }
    }

    private sealed class FontPackFileManifest
    {
        public string? RelativePath { get; set; }
        public long? ByteSize { get; set; }
        public string? Sha256 { get; set; }
    }

    private sealed class FontPackFamilyManifest
    {
        public string? RequestedFamily { get; set; }
        public string? ResolvedFamily { get; set; }
        public string? RelativeFontFile { get; set; }
        public int? Weight { get; set; }
        public bool? Italic { get; set; }
        public int? FaceIndex { get; set; }
        public bool? HasMathTable { get; set; }
    }

    private sealed class FontPackFallbackManifest
    {
        public string? Family { get; set; }
    }
}
