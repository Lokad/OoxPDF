using System.Security.Cryptography;
using System.Text.Json;

namespace Lokad.OoxPdf.Fonts;

public sealed class OoxPdfFontPackResolver : IFontResolver, IFontCatalog
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly FontPackFileSource fileSource;
    private readonly IReadOnlyList<FontPackFace> faces;
    private readonly IReadOnlyList<string> fallbackFamilies;

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
            cancellationToken).ConfigureAwait(false);
        FontPackManifest manifest = DeserializeManifest(manifestBytes);
        ValidatedManifest validated = ValidateManifest(packId, manifest);

        var resolver = new OoxPdfFontPackResolver(
            packId,
            packRootUri,
            httpClient,
            validated.Faces,
            validated.FallbackFamilies);

        return resolver;
    }

    public FontFaceResolution Resolve(FontRequest request)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(request.FamilyName);

        FontPackFace[] exact = faces
            .Where(face =>
                face.RequestedFamily.Equals(request.FamilyName, StringComparison.OrdinalIgnoreCase) ||
                face.ResolvedFamily.Equals(request.FamilyName, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (exact.Length != 0)
        {
            return SelectBest(exact, request, isFallback: false);
        }

        foreach (string fallback in fallbackFamilies)
        {
            FontPackFace[] configuredFallbacks = faces
                .Where(face =>
                face.RequestedFamily.Equals(fallback, StringComparison.OrdinalIgnoreCase) ||
                face.ResolvedFamily.Equals(fallback, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (configuredFallbacks.Length != 0)
            {
                return SelectBest(configuredFallbacks, request, isFallback: true);
            }
        }

        FontPackFace[] textFaces = faces.Where(face => !face.Style.HasMathTable).ToArray();
        return SelectBest(textFaces.Length == 0 ? faces : textFaces, request, isFallback: true);
    }

    public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        return faces.Select(face => face.ToResolution(fileSource, isFallback: false)).ToArray();
    }

    private FontFaceResolution SelectBest(IReadOnlyList<FontPackFace> candidates, FontRequest request, bool isFallback)
    {
        int targetWeight = request.Bold ? 700 : 400;
        FontPackFace face = candidates
            .OrderBy(candidate => candidate.Style.Italic == request.Italic ? 0 : 1000)
            .ThenBy(candidate => Math.Abs(candidate.Style.WeightClass - targetWeight))
            .ThenBy(candidate => candidate.ResolvedFamily, StringComparer.OrdinalIgnoreCase)
            .ThenBy(candidate => candidate.RelativeFontFile, StringComparer.Ordinal)
            .ThenBy(candidate => candidate.Style.FaceIndex)
            .First();

        return face.ToResolution(fileSource, isFallback) with
        {
            RequestedFamily = request.FamilyName
        };
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

    private static async Task<byte[]> DownloadBytesAsync(
        HttpClient httpClient,
        Uri uri,
        string description,
        CancellationToken cancellationToken)
    {
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

            return await response.Content.ReadAsByteArrayAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (OoxPdfFontPackException)
        {
            throw;
        }
        catch (Exception ex) when (ex is HttpRequestException or IOException or InvalidOperationException or TaskCanceledException)
        {
            throw new OoxPdfFontPackException(
                OoxPdfFontPackDiagnosticIds.FontPackDownloadFailed,
                $"Unable to download {description} from '{uri}'.",
                ex);
        }
    }

    private static FontPackManifest DeserializeManifest(byte[] manifestBytes)
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

    private sealed class FontPackFileSource(string packId, Uri packRootUri, HttpClient httpClient)
    {
        public IFontProgramSource Create(FontPackFile file)
        {
            return new OoxPdfFontPackProgramSource(packId, packRootUri, httpClient, file);
        }
    }

    private sealed class OoxPdfFontPackProgramSource(
        string packId,
        Uri packRootUri,
        HttpClient httpClient,
        FontPackFile file) : IFontProgramSource
    {
        private ReadOnlyMemory<byte>? cachedBytes;

        public string StableId => "ooxpdf-font-pack:" + packId + ":" + file.Sha256;

        public async ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default)
        {
            if (cachedBytes is ReadOnlyMemory<byte> cached)
            {
                return cached;
            }

            byte[] bytes = await DownloadBytesAsync(
                httpClient,
                new Uri(packRootUri, file.RelativePath),
                "font file '" + file.RelativePath + "'",
                ct).ConfigureAwait(false);

            ValidateFontBytes(file, bytes);
            cachedBytes = bytes;
            return bytes;
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

    private sealed record FontPackFile(
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
