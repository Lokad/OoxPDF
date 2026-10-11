using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxKerningThreshold
{
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
        enabled = minimum > 0 && run.FontSize * 2d >= minimum;
        return true;
    }
}
