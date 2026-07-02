namespace Lokad.OoxPdf.Fonts;

internal interface IFontCatalog
{
    IReadOnlyList<FontFaceResolution> GetDiscoveredFonts();
}
