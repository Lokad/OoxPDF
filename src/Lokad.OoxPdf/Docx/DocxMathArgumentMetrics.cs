using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static class DocxMathArgumentMetrics
{
    public static bool TryGetScriptLevel(DocxMathRow argument, out int level, int startingLevel = 1)
    {
        level = startingLevel;
        if (startingLevel is < 0 or > 2) { return false; }
        if (!argument.HasSupportedProperties && !argument.HasKnownArgumentProperties) { return false; }
        int increment = argument.ArgumentProperties?.Size ?? 0;
        if (increment is < -2 or > 2) { return false; }
        // Source increments adjust the role's starting style, bounded to the
        // font's text/script/script-script sizes. Other row roles stay neutral.
        level = Math.Clamp(startingLevel - increment, 0, 2);
        return true;
    }

    public static bool TryGetScriptFontSize(OpenTypeFont font, double nominalSize, int level, out double emittedSize)
    {
        emittedSize = 0d;
        if (!TryGetNominalScriptFontSize(font, nominalSize, level, out double argumentSize)) return false;
        emittedSize = OfficePdfTextEmissionProfile.FontSize(argumentSize);
        return double.IsFinite(emittedSize) && emittedSize > 0d;
    }

    internal static bool TryGetNominalScriptFontSize(OpenTypeFont font, double nominalSize, int level, out double argumentSize)
    {
        argumentSize = 0d;
        if (level is < 1 or > 2 || !double.IsFinite(nominalSize) || nominalSize <= 0d ||
            !font.TryGetMathConstants(out OpenTypeMathConstants constants)) { return false; }
        int percent = level == 1 ? constants.ScriptPercentScaleDown : constants.ScriptScriptPercentScaleDown;
        if (percent <= 0) { return false; }
        // Independent Word8/12/16/20/24pt contrasts: resolve the font's MATH
        // scale, round the nominal argument down to half points, then apply
        // the existing Office PDF export grid. A constant71% fails20/24pt.
        argumentSize = Math.Floor(nominalSize * (percent / 100d) * 2d) / 2d;
        return double.IsFinite(argumentSize) && argumentSize > 0d;
    }
}
