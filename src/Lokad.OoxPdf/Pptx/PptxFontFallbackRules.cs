namespace Lokad.OoxPdf.Pptx;

internal static class PptxFontFallbackRules
{
    public const string DefaultLatinTypeface = "Arial";

    public static string ResolveDefaultLatinTypeface(string? familyName)
    {
        return string.IsNullOrWhiteSpace(familyName)
            ? DefaultLatinTypeface
            : familyName;
    }
}
