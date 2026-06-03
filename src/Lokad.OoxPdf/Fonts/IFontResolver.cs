namespace Lokad.OoxPdf.Fonts;

public interface IFontResolver
{
    FontFaceResolution Resolve(FontRequest request);
}
