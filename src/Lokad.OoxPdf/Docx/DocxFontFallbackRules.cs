using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Docx;

internal static class DocxFontFallbackRules
{
    public const string DefaultDocumentTypefaceRequest = "OOXPDF_DOCUMENT_DEFAULT";

    private static readonly string[] DefaultLatinTypefaceCandidates =
    [
        // Stable library fallback for undeclared Latin typefaces. Office's
        // application default varies; it is not always Times New Roman.
        // Explicit/theme families and a resolver's document default come first.
        "Times New Roman",
        "Aptos",
        "Calibri",
        "Arial"
    ];

    public static FontFaceResolution ResolveDefaultDocumentTypeface(
        IFontResolver fontResolver,
        bool bold,
        bool italic)
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

    private static readonly string[] PerCharacterFallbackOrder =
    [
        "Symbol",
        "Segoe UI Symbol",
        "Aptos",
        "Calibri",
        "Arial"
    ];

    public static IReadOnlyList<string> PerCharacterFallbackFamilies(string? primaryFamily)
    {
        var families = new List<string>();
        foreach (string family in PerCharacterFallbackOrder)
        {
            if (!string.Equals(family, primaryFamily, StringComparison.OrdinalIgnoreCase))
            {
                families.Add(family);
            }
        }
        return families;
    }

    public static IReadOnlyList<FontFaceResolution> ResolveCandidateResolutions(IFontResolver fontResolver, DocxTextRun? run, string? primaryFamily, FontFaceResolution primaryResolution, FontFaceResolution? documentFallback)
    {
        var resolutions = new List<FontFaceResolution> { primaryResolution };
        bool bold = run is not null && run.Bold;
        bool italic = run is not null && run.Italic;
        foreach (string family in PerCharacterFallbackFamilies(primaryFamily))
        {
            FontFaceResolution candidate = fontResolver.Resolve(new FontRequest(family, bold, italic));
            if (!ContainsResolution(resolutions, candidate))
            {
                resolutions.Add(candidate);
            }
        }

        if (documentFallback is not null && !ContainsResolution(resolutions, documentFallback))
        {
            resolutions.Add(documentFallback);
        }

        return resolutions;
    }

    private static bool ContainsResolution(IReadOnlyList<FontFaceResolution> resolutions, FontFaceResolution candidate)
    {
        foreach (FontFaceResolution existing in resolutions)
        {
            if (existing.Source.StableId == candidate.Source.StableId && existing.FontFaceIndex == candidate.FontFaceIndex)
            {
                return true;
            }
        }

        return false;
    }
}
