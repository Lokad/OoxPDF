namespace Lokad.OoxPdf.Fonts;

public interface IFontResolver
{
    FontResolution Resolve(FontRequest request);
}
