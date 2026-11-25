using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxLineStyleReader
{
    internal const double OfficeDefaultLineWidthPoints = 0.75d;

    public static double? ReadWidth(XElement? line) => line?.Attribute("w") is { } width
        ? OoxUnits.EmuToPoints(long.Parse(width.Value, CultureInfo.InvariantCulture))
        : null;

    public static bool TryReadLineWithAlpha(
        XElement line,
        PptxTheme theme,
        PptxColorMap colorMap,
        out RgbColor color,
        out double lineWidth,
        out double alpha,
        double? fallbackLineWidth)
    {
        lineWidth = ReadWidth(line) ?? fallbackLineWidth ?? OfficeDefaultLineWidthPoints;
        return PptxColorResolver.TryReadSolidColorWithAlpha(line, theme, colorMap, out color, out alpha);
    }
}
