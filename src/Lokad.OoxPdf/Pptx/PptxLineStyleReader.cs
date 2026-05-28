using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxLineStyleReader
{
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static bool TryReadLineWithAlpha(
        XElement shapeProperties,
        PptxTheme theme,
        PptxColorMap colorMap,
        out RgbColor color,
        out double lineWidth,
        out double alpha,
        double? fallbackLineWidth = null)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : fallbackLineWidth ?? 1d;
        return PptxColorResolver.TryReadSolidColorWithAlpha(line, theme, colorMap, out color, out alpha);
    }
}
