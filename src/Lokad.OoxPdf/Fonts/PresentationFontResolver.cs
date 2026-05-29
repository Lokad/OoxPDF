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

    public FontResolution Resolve(FontRequest request)
    {
        return primary.Resolve(request);
    }

    public FontResolution ResolvePresentationTextFace(FontRequest request)
    {
        return primary is WindowsFontResolver windows
            ? windows.ResolvePresentationTextFace(request)
            : primary.Resolve(request);
    }

    public (FontResolution Resolution, OpenTypeFont Font)? ResolvePresentationOpenTypeFont(FontRequest request)
    {
        FontResolution resolution = ResolvePresentationTextFace(request);
        OpenTypeFont? font = LoadOpenTypeFont(resolution);
        return font is null ? null : (resolution, font);
    }

    public IReadOnlyList<FontResolution> GetDiscoveredFonts()
    {
        return windowsCatalog.GetDiscoveredFonts();
    }

    private OpenTypeFont? LoadOpenTypeFont(FontResolution resolution)
    {
        if (resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
        {
            return null;
        }

        string key = resolution.FontFilePath + "\u001f" + resolution.FontFaceIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (openTypeFonts.TryGetValue(key, out OpenTypeFont? cached))
        {
            return cached;
        }

        try
        {
            cached = OpenTypeFont.Load(resolution.FontFilePath, resolution.FontFaceIndex);
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
        {
            cached = null;
        }

        openTypeFonts[key] = cached;
        return cached;
    }
}
