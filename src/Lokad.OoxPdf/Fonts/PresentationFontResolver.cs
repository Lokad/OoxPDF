namespace Lokad.OoxPdf.Fonts;

internal sealed class PresentationFontResolver
{
    private readonly IFontResolver primary;
    private readonly WindowsFontResolver windowsCatalog;
    private readonly Dictionary<string, OpenTypeFont?> openTypeFonts = new(StringComparer.OrdinalIgnoreCase);

    public PresentationFontResolver(IFontResolver? primary = null)
    {
        this.primary = primary ?? new WindowsFontResolver();
        windowsCatalog = this.primary as WindowsFontResolver ?? new WindowsFontResolver();
    }

    public FontFaceResolution Resolve(FontRequest request)
    {
        return primary.Resolve(request);
    }

    public FontFaceResolution ResolvePresentationTextFace(FontRequest request)
    {
        return primary is WindowsFontResolver windows
            ? windows.ResolvePresentationTextFace(request)
            : primary.Resolve(request);
    }

    public (FontFaceResolution Resolution, OpenTypeFont Font)? ResolvePresentationOpenTypeFont(FontRequest request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        FontFaceResolution resolution = ResolvePresentationTextFace(request);
        OpenTypeFont? font = LoadOpenTypeFont(resolution, cancellationToken);
        return font is null ? null : (resolution, font);
    }

    public IReadOnlyList<FontFaceResolution> GetDiscoveredFonts()
    {
        return windowsCatalog.GetDiscoveredFonts();
    }

    private OpenTypeFont? LoadOpenTypeFont(FontFaceResolution resolution, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string key = resolution.Source.StableId + "\u001f" + resolution.FontFaceIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (openTypeFonts.TryGetValue(key, out OpenTypeFont? cached))
        {
            return cached;
        }

        cached = FontProgramLoader.Load(resolution, cancellationToken);
        openTypeFonts[key] = cached;
        return cached;
    }
}
