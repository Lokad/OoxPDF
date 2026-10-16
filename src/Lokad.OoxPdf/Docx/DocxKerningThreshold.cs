using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxKerningThreshold
{
    internal static bool IsMet(ushort minimum, double nominalFontSize) =>
        minimum > 0 && nominalFontSize * 2d >= minimum;

    internal static bool TryGetMinimum(DocxTextRun run, out ushort minimum)
    {
        minimum = 0;
        return run.KerningMinimumHalfPointsValue is not { } token ||
            ushort.TryParse(token, NumberStyles.None, CultureInfo.InvariantCulture, out minimum) && minimum <= 32767;
    }

    internal static bool TryResolve(DocxTextRun run, out bool enabled)
    {
        enabled = false;
        if (!TryGetMinimum(run, out ushort minimum) || !double.IsFinite(run.FontSize) || run.FontSize <= 0d) return false;
        enabled = IsMet(minimum, run.FontSize);
        return true;
    }
}
