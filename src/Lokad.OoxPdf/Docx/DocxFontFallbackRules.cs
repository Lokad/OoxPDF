using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Docx;

internal static class DocxFontFallbackRules
{
    public const string DefaultDocumentTypefaceRequest = "OOXPDF_DOCUMENT_DEFAULT";

    private static readonly string[] DefaultLatinTypefaceCandidates =
    [
        "Aptos",
        "Calibri",
        "Arial"
    ];

    public static FontFaceResolution ResolveDefaultDocumentTypeface(
        IFontResolver fontResolver,
        bool bold = false,
        bool italic = false)
    {
        FontFaceResolution defaultResolution = fontResolver.Resolve(new FontRequest(DefaultDocumentTypefaceRequest, bold, italic));
        if (!defaultResolution.IsFallback ||
            defaultResolution.FamilyName.Equals(DefaultDocumentTypefaceRequest, StringComparison.OrdinalIgnoreCase))
        {
            return defaultResolution;
        }

        foreach (string family in DefaultLatinTypefaceCandidates)
        {
            FontFaceResolution resolution = fontResolver.Resolve(new FontRequest(family, bold, italic));
            if (!resolution.IsFallback)
            {
                return resolution;
            }
        }

        return defaultResolution;
    }
}
