using System.Globalization;

namespace Lokad.OoxPdf.Diagnostics;

internal static class FontDiagnostics
{
    public static OoxPdfDiagnostic MissingGlyphs(string family, IReadOnlyCollection<int> sortedCodePoints)
    {
        const int maxListedCodePoints = 12;
        string listed = string.Join(", ", sortedCodePoints.Take(maxListedCodePoints)
            .Select(codePoint => "U+" + codePoint.ToString("X4", CultureInfo.InvariantCulture)));
        if (sortedCodePoints.Count > maxListedCodePoints)
        {
            listed += ", and " + (sortedCodePoints.Count - maxListedCodePoints).ToString(CultureInfo.InvariantCulture) + " more";
        }

        return new OoxPdfDiagnostic(
            "FONT_MISSING_GLYPHS",
            OoxPdfSeverity.Warning,
            sortedCodePoints.Count.ToString(CultureInfo.InvariantCulture) + " characters of " + family +
                " have no glyph in any usable face (" + listed + ") and show as question mark.",
            Feature: family,
            Fallback: "Question-mark substitution");
    }
}
