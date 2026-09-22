using System.Security.Cryptography;
using System.Text.Json;

namespace Lokad.OoxPdf.Fonts;

public sealed class OoxPdfFontPackResolver : IFontResolver, IFontCatalog
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    internal const long MaxManifestBytes = 1024L * 1024L;
    internal const long MaxFontFileBytes = 64L * 1024L * 1024L;

    // downloaded programs persist for the resolver lifetime. The face
    // population is manifest-finite, but retaining every downloaded face is still
    // unbounded live memory for long-lived resolvers, so retained downloads are
    // capped in aggregate with LRU eviction. Evicted sources re-download on demand
    // (hash-verified every time, same retry/cancellation path); concurrently active
    // conversions simply keep their already-handed-out array alive via GC.
    // R13 ownership: the resolver owns retained downloads, conversions borrow the
    // arrays, and in-flight download memory is bounded separately by the process-wide
    // MaxConcurrentFontDownloads throttle (each slot holds at most MaxFontFileBytes).
    internal const long MaxTotalFontPackBytes = 256L * 1024L * 1024L;
    private readonly FontPackFileSource fileSource;
    private readonly IReadOnlyList<FontPackFace> faces;
    private readonly IReadOnlyList<string> fallbackFamilies;
    private readonly Dictionary<string, FontPackFace[]> facesByFamily;
    private readonly FontPackFace[] textFaces;
    private readonly Dictionary<FontRequest, FontFaceResolution> requestCache = new(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);

    private OoxPdfFontPackResolver(
        string packId,
        Uri packRootUri,
        HttpClient httpClient,
        IReadOnlyList<FontPackFace> faces,
        IReadOnlyList<string> fallbackFamilies)
    {
        PackId = packId;
        PackRootUri = packRootUri;
        fileSource = new FontPackFileSource(packId, packRootUri, httpClient);
        this.faces = faces;
        this.fallbackFamilies = fallbackFamilies;
        facesByFamily = IndexFacesByFamily(faces);
        textFaces = faces.Where(face => !face.Style.HasMathTable).ToArray();
    }

    private static Dictionary<string, FontPackFace[]> IndexFacesByFamily(IReadOnlyList<FontPackFace> faces)
    {
        var grouped = new Dictionary<string, List<FontPackFace>>(StringComparer.OrdinalIgnoreCase);
        foreach (FontPackFace face in faces)
        {
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
            index[family] = group.ToArray();
        }

        return index;
    }

    public string PackId { get; }

    public Uri PackRootUri { get; }

    public static async Task<OoxPdfFontPackResolver> CreateHttpAsync(
        string packId,
        Uri sourceUri,
        HttpClient httpClient,
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

        byte[] manifestBytes = await DownloadBytesAsync(
            httpClient,
            manifestUri,
            "font pack manifest",
            MaxManifestBytes,
            cancellationToken).ConfigureAwait(false);
        FontPackManifest manifest = DeserializeManifest();
        ValidatedManifest validated = ValidateManifest(packId, manifest);

        var resolver = new OoxPdfFontPackResolver(
            packId,
            packRootUri,
            httpClient,
            validated.Faces,
            validated.FallbackFamilies);

        return resolver;

        FontPackManifest DeserializeManifest()
        {
            try
            {
                return JsonSerializer.Deserialize<FontPackManifest>(manifestBytes, JsonOptions)
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

    public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        return faces.Select(face => face.ToResolution(fileSource, isFallback: false)).ToArray();
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

    // R13: process-wide bound on concurrent font downloads. Each download holds up
    // to MaxFontFileBytes in flight, so the semaphore caps aggregate in-flight
    // download memory separately from the retained-download total. Retries run
    // inside one slot; cancellation always throws.
    internal const int MaxConcurrentFontDownloads = 4;
    private static readonly SemaphoreSlim DownloadConcurrency = new(MaxConcurrentFontDownloads, MaxConcurrentFontDownloads);

    private static async Task<byte[]> DownloadBytesAsync(
        HttpClient httpClient,
        Uri uri,
        string description,
        long maxBytes,
        CancellationToken cancellationToken)
    {
        // Transient transport failures retry immediately and bounded.
        // Server answers (status, declared size, over-cap bodies) fail fast;
        // cancellation always throws.
        await DownloadConcurrency.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return await DownloadWithRetriesAsync(httpClient, uri, description, maxBytes, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            DownloadConcurrency.Release();
        }
    }

    private static async Task<byte[]> DownloadWithRetriesAsync(
        HttpClient httpClient,
        Uri uri,
        string description,
        long maxBytes,
        CancellationToken cancellationToken)
    {
        const int maxAttempts = 3;
        for (int attempt = 1; ; attempt++)
        {
            try
            {
                return await DownloadBytesSingleAttemptAsync(httpClient, uri, description, maxBytes, cancellationToken).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
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

    private static async Task<byte[]> DownloadBytesSingleAttemptAsync(
        HttpClient httpClient,
        Uri uri,
        string description,
        long maxBytes,
        CancellationToken cancellationToken)
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

        using Stream content = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        return await CopyCappedAsync(content, maxBytes, description, uri.ToString(), cancellationToken).ConfigureAwait(false);
    }

    private static bool IsTransientDownloadFailure(Exception ex)
    {
        return ex is HttpRequestException or IOException or InvalidOperationException or TaskCanceledException;
    }

    private static ValidatedManifest ValidateManifest(string packId, FontPackManifest manifest)
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

    internal static async Task<byte[]> CopyCappedAsync(Stream source, long maxBytes, string description, string uri, CancellationToken cancellationToken)
    {
        byte[] buffer = new byte[81920];
        using var destination = new MemoryStream();
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int read = await source.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            if (read == 0)
            {
                return destination.ToArray();
            }

            if (checked(destination.Length + read) > maxBytes)
            {
                throw new OoxPdfFontPackException(
                    OoxPdfFontPackDiagnosticIds.FontPackDownloadFailed,
                    "Unable to download " + description + " from " + uri + ": response exceeds the limit.");
            }

            destination.Write(buffer, 0, read);
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

    internal sealed class FontPackFileSource
    {
        private readonly string packId;
        private readonly Uri packRootUri;
        private readonly HttpClient httpClient;
        private readonly long maxTotalBytes;
        private readonly object sync = new();
        private readonly Dictionary<string, IFontProgramSource> sources = new(StringComparer.Ordinal);
        private readonly Dictionary<string, long> retainedSizes = new(StringComparer.Ordinal);
        private readonly LinkedList<string> retainedOrder = new();
        private long retainedBytes;

        internal FontPackFileSource(string packId, Uri packRootUri, HttpClient httpClient)
            : this(packId, packRootUri, httpClient, MaxTotalFontPackBytes)
        {
        }

        internal FontPackFileSource(string packId, Uri packRootUri, HttpClient httpClient, long maxTotalBytes)
        {
            this.packId = packId;
            this.packRootUri = packRootUri;
            this.httpClient = httpClient;
            this.maxTotalBytes = maxTotalBytes;
        }

        public IFontProgramSource Create(FontPackFile file)
        {
            lock (sync)
            {
                if (!sources.TryGetValue(file.RelativePath, out IFontProgramSource? source))
                {
                    source = new OoxPdfFontPackProgramSource(packId, packRootUri, httpClient, file, this);
                    sources[file.RelativePath] = source;
                }

                return source;
            }
        }

        // R13: cache hits refresh recency so the order evicts least-recently-used
        // files instead of least-recently-downloaded ones. Entries missing from
        // the retained set (never downloaded or already evicted) are ignored.
        internal void NoteAccessed(string relativePath)
        {
            lock (sync)
            {
                if (retainedSizes.ContainsKey(relativePath))
                {
                    retainedOrder.Remove(relativePath);
                    retainedOrder.AddLast(relativePath);
                }
            }
        }

        internal void NoteDownloaded(string relativePath, long byteCount)
        {
            lock (sync)
            {
                if (retainedSizes.TryGetValue(relativePath, out long previous))
                {
                    retainedBytes -= previous;
                    retainedOrder.Remove(relativePath);
                }

                retainedSizes[relativePath] = byteCount;
                retainedOrder.AddLast(relativePath);
                retainedBytes += byteCount;

                while (retainedBytes > maxTotalBytes && retainedOrder.Count > 0)
                {
                    string eldest = retainedOrder.First!.Value;
                    if (retainedOrder.Count == 1)
                    {
                        // Only the just-downloaded source is left; a single font always
                        // fits by the per-font cap, so this is unreachable in practice.
                        break;
                    }

                    retainedOrder.RemoveFirst();
                    retainedBytes -= retainedSizes.GetValueOrDefault(eldest);
                    retainedSizes.Remove(eldest);
                    if (sources.TryGetValue(eldest, out IFontProgramSource? source) &&
                        source is OoxPdfFontPackProgramSource programSource)
                    {
                        programSource.EvictCachedBytes();
                    }
                }
            }
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
        private ReadOnlyMemory<byte>? cachedBytes;

        public string StableId => "ooxpdf-font-pack:" + packId + ":" + file.Sha256;

        public async ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct)
        {
            if (cachedBytes is ReadOnlyMemory<byte> cached)
            {
                owner.NoteAccessed(file.RelativePath);
                return cached;
            }

            await gate.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                if (cachedBytes is ReadOnlyMemory<byte> rechecked)
                {
                    owner.NoteAccessed(file.RelativePath);
                    return rechecked;
                }

                byte[] bytes = await DownloadBytesAsync(
                    httpClient,
                    new Uri(packRootUri, file.RelativePath),
                    "font file '" + file.RelativePath + "'",
                    Math.Min(file.ByteSize, MaxFontFileBytes),
                    ct).ConfigureAwait(false);

                ValidateFontBytes(file, bytes);
                cachedBytes = bytes;
                owner.NoteDownloaded(file.RelativePath, bytes.LongLength);
                return bytes;
            }
            finally
            {
                gate.Release();
            }
        }

        // Clears retained bytes so the owner aggregate stays bounded. Arrays already
        // handed out stay alive (and valid: every download is hash-verified) via GC;
        // the next access simply downloads again through the per-source gate.
        internal void EvictCachedBytes()
        {
            cachedBytes = null;
        }

        private static void ValidateFontBytes(FontPackFile file, byte[] bytes)
        {
            if (bytes.LongLength != file.ByteSize)
            {
                throw new OoxPdfFontPackException(
                    OoxPdfFontPackDiagnosticIds.FontPackHashMismatch,
                    $"Font file '{file.RelativePath}' size mismatch: expected {file.ByteSize} bytes, got {bytes.LongLength} bytes.");
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
