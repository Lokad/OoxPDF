using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool TryReadSolidColor(XElement? element, PptxTheme theme, out RgbColor color)
    {
        return PptxColorResolver.TryReadSolidColor(element, theme, out color);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, out RgbColor color, out double alpha)
    {
        return PptxColorResolver.TryReadSolidColorWithAlpha(element, theme, out color, out alpha);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, XElement? placeholderColorContainer, out RgbColor color, out double alpha)
    {
        return PptxColorResolver.TryReadSolidColorWithAlpha(element, theme, placeholderColorContainer, out color, out alpha);
    }

    private static double ReadAlpha(XElement? colorContainer)
    {
        return PptxColorResolver.ReadAlpha(colorContainer);
    }

    private static byte ToByte(double value)
    {
        return PptxColorResolver.ToByte(value);
    }

    private static bool TryReadLine(XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth)
    {
        return TryReadLineWithAlpha(shapeProperties, theme, out color, out lineWidth, out _);
    }

    private static bool TryReadLineWithAlpha(
        XElement shapeProperties,
        PptxTheme theme,
        out RgbColor color,
        out double lineWidth,
        out double alpha,
        double? fallbackLineWidth = null)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : fallbackLineWidth ?? 1d;
        return TryReadSolidColorWithAlpha(line, theme, out color, out alpha);
    }
}
