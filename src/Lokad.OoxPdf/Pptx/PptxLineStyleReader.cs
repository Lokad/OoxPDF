using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxLineStyleReader
{
    private const double OfficeDefaultLineWidthPoints = 0.75d;

    public static bool TryReadLineWithAlpha(
        XElement line,
        PptxTheme theme,
        PptxColorMap colorMap,
        out RgbColor color,
        out double lineWidth,
        out double alpha,
        double? fallbackLineWidth)
    {
        lineWidth = line.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : fallbackLineWidth ?? OfficeDefaultLineWidthPoints;
        return PptxColorResolver.TryReadSolidColorWithAlpha(line, theme, colorMap, out color, out alpha);
    }
}
