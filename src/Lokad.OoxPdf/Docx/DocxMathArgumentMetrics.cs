using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static class DocxMathArgumentMetrics
{
    public static bool TryGetScriptFontSize(OpenTypeFont font, double nominalSize, int level, out double emittedSize)
    {
        emittedSize = 0d;
        if (level is < 1 or > 2 || !double.IsFinite(nominalSize) || nominalSize <= 0d ||
            !font.TryGetMathConstants(out OpenTypeMathConstants constants)) { return false; }
        int percent = level == 1 ? constants.ScriptPercentScaleDown : constants.ScriptScriptPercentScaleDown;
        if (percent <= 0) { return false; }
        // Independent Word8/12/16/20/24pt contrasts: resolve the font's MATH
        // scale, round the nominal argument down to half points, then apply
        // the existing Office PDF export grid. A constant71% fails20/24pt.
        double nominalArgument = Math.Floor(nominalSize * (percent / 100d) * 2d) / 2d;
        if (!double.IsFinite(nominalArgument) || nominalArgument <= 0d) { return false; }
        emittedSize = OfficePdfTextEmissionProfile.FontSize(nominalArgument);
        return double.IsFinite(emittedSize) && emittedSize > 0d;
    }
}
