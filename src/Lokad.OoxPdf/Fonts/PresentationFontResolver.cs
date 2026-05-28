namespace Lokad.OoxPdf.Fonts;

internal sealed class PresentationFontResolver
{
    private readonly IFontResolver primary;
    private readonly WindowsFontResolver windowsCatalog;

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

    public IReadOnlyList<FontResolution> GetDiscoveredFonts()
    {
        return windowsCatalog.GetDiscoveredFonts();
    }
}
