using System.Globalization;

namespace Lokad.OoxPdf.Diagnostics;

internal static class FontDiagnostics
{
    public static OoxPdfDiagnostic MissingGlyphs(string family, IReadOnlyCollection<int> sortedCodePoints)
        => new(
            "FONT_MISSING_GLYPHS",
            OoxPdfSeverity.Warning,
            sortedCodePoints.Count.ToString(CultureInfo.InvariantCulture) + " characters of " + family +
                " have no glyph in any usable face (" + FormatCodePoints(sortedCodePoints) +
                "). A question mark is used when the selected font provides one.",
            Feature: family,
            Fallback: "Question-mark substitution");

    public static OoxPdfDiagnostic MissingBalloonTitleGlyphs(IReadOnlyCollection<int> sortedCodePoints)
        => new(
            "DOCX_BALLOON_TITLE_MISSING_GLYPHS",
            OoxPdfSeverity.Warning,
            "A DOCX review balloon title contains characters absent from its selected font (" +
                FormatCodePoints(sortedCodePoints) + "). Unavailable glyphs use a question mark when the font has one.",
            Feature: "DOCX review balloon title",
            Fallback: "Question-mark substitution");

    private static string FormatCodePoints(IReadOnlyCollection<int> sortedCodePoints)
    {
        const int maxListedCodePoints = 12;
        string listed = string.Join(", ", sortedCodePoints.Take(maxListedCodePoints)
            .Select(codePoint => "U+" + codePoint.ToString("X4", CultureInfo.InvariantCulture)));
        if (sortedCodePoints.Count > maxListedCodePoints)
        {
            listed += ", and " + (sortedCodePoints.Count - maxListedCodePoints).ToString(CultureInfo.InvariantCulture) + " more";
        }

        return listed;
    }
}
