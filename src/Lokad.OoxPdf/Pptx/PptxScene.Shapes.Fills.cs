using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    internal static PptxSceneGradientFill ReadShapeGradientFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? gradientFill = shapeProperties?.Element(DrawingNamespace + "gradFill");
        XElement? gradientStopList = gradientFill?.Element(DrawingNamespace + "gsLst");
        if (gradientFill is null)
        {
            return new PptxSceneGradientFill(false, false, 0d, []);
        }

        if (gradientStopList is null)
        {
            return new PptxSceneGradientFill(false, true, 0d, []);
        }

        XElement[] rawStops = gradientStopList
            .Elements(DrawingNamespace + "gs")
            .ToArray();
        PptxSceneGradientStop[] stops = rawStops
            .Select(stop => TryReadGradientStop(stop, theme, colorMap, out PptxSceneGradientStop parsed) ? parsed : (PptxSceneGradientStop?)null)
            .OfType<PptxSceneGradientStop>()
            .OrderBy(stop => stop.Offset)
            .ToArray();
        if (stops.Length < 2 ||
            gradientFill.Element(DrawingNamespace + "lin") is not { } linear)
        {
            return new PptxSceneGradientFill(false, true, 0d, []);
        }

        double angleDegrees = OoxXml.ParseOptionalLong(linear, "ang", 0) / 60000d;
        bool hasUnsupportedGradient =
            stops.Length != rawStops.Length ||
            !PptxColorResolver.TryGetUniformGradientAlpha(stops, static stop => stop.Alpha, out _);
        return new PptxSceneGradientFill(true, hasUnsupportedGradient, angleDegrees, stops);
    }

    internal static bool HasSupportedAuthoredGradientFill(XElement gradientFill)
    {
        XName? ownerName = gradientFill.Parent?.Name;
        if ((ownerName != PresentationNamespace + "spPr" && ownerName != PresentationNamespace + "bgPr") ||
            gradientFill.Element(DrawingNamespace + "lin") is null ||
            gradientFill.Element(DrawingNamespace + "gsLst") is not { } stopList)
        {
            return false;
        }

        XElement[] stops = stopList.Elements(DrawingNamespace + "gs").ToArray();
        return stops.Length >= 2 &&
            stops.All(stop => stop.Elements().Any(PptxColorResolver.IsDrawingColorElement)) &&
            PptxColorResolver.HasSupportedGradientStopAlpha(stops);
    }

    private static bool TryReadGradientStop(XElement gradientStop, PptxTheme theme, PptxColorMap colorMap, out PptxSceneGradientStop stop)
    {
        if (!PptxColorResolver.TryReadSolidColorWithAlpha(gradientStop, theme, colorMap, out RgbColor color, out double alpha))
        {
            stop = default;
            return false;
        }

        double offset = gradientStop.Attribute("pos") is { } position
            ? Math.Clamp(OoxXml.ParsePercentage(position.Value), 0d, 1d) : 0d;
        stop = new PptxSceneGradientStop(offset, color, alpha);
        return true;
    }

    private static PptxScenePatternFill ReadShapePatternFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        if (patternFill is null)
        {
            return default;
        }

        string? preset = (string?)patternFill.Attribute("prst");
        if (preset is null || !IsSupportedDiagonalPatternFill(preset))
        {
            return new PptxScenePatternFill(false, true, preset ?? string.Empty, default, default);
        }

        return ReadPatternFill(patternFill, preset, theme, colorMap);
    }

    private static PptxScenePatternFill ReadPatternFill(
        XElement patternFill, string preset, PptxTheme theme, PptxColorMap colorMap)
    {
        RgbColor foreground = PptxColorResolver.TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "fgClr"), theme, colorMap, out RgbColor foregroundColor, out _)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = PptxColorResolver.TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "bgClr"), theme, colorMap, out RgbColor backgroundColor, out _)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        return new PptxScenePatternFill(true, false, preset, foreground, background);
    }

    internal static bool IsSupportedDiagonalPatternFill(string? preset)
    {
        return preset is not null &&
            (preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase) ||
             preset.Contains("DnDiag", StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryReadShapeFill(
        XElement? shapeProperties,
        PptxTheme theme,
        PptxColorMap colorMap,
        PptxFormatSchemeReference fillReference,
        out RgbColor color,
        out double alpha)
    {
        if (shapeProperties?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            alpha = 1d;
            return false;
        }

        if (PptxColorResolver.TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out color, out alpha))
        {
            return true;
        }

        if (fillReference.Style is not null &&
            PptxColorResolver.TryReadSolidColorWithAlpha(fillReference.Style, theme, colorMap, fillReference.Reference, out color, out alpha))
        {
            return true;
        }

        return fillReference.Index > 0 && PptxColorResolver.TryReadSolidColorWithAlpha(fillReference.Reference, theme, colorMap, out color, out alpha);
    }

    private static bool TryReadInheritedGroupFill(
        XElement shape,
        PptxTheme theme,
        PptxColorMap colorMap,
        out RgbColor color,
        out double alpha)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        if (shapeProperties?.Element(DrawingNamespace + "noFill") is not null ||
            shapeProperties?.Element(DrawingNamespace + "grpFill") is null)
        {
            color = default;
            alpha = 1d;
            return false;
        }

        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp"))
        {
            XElement? groupProperties = group.Element(PresentationNamespace + "grpSpPr");
            if (groupProperties?.Element(DrawingNamespace + "noFill") is not null)
            {
                color = default;
                alpha = 1d;
                return false;
            }

            if (PptxColorResolver.TryReadSolidColorWithAlpha(groupProperties, theme, colorMap, out color, out alpha))
            {
                return true;
            }
        }

        color = default;
        alpha = 1d;
        return false;
    }
}
