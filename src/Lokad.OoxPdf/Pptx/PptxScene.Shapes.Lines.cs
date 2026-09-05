using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static bool TryReadShapeLine(
        XElement? shapeProperties,
        PptxTheme theme,
        PptxColorMap colorMap,
        PptxFormatSchemeReference lineReference,
        out RgbColor color,
        out double lineWidth,
        out double alpha)
    {
        XElement? explicitLine = shapeProperties?.Element(DrawingNamespace + "ln");
        if (explicitLine?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        double? styleLineWidth = TryReadStyleLineWidth(lineReference, out double inheritedLineWidth)
            ? inheritedLineWidth
            : null;
        if (shapeProperties is not null && explicitLine is not null && PptxLineStyleReader.TryReadLineWithAlpha(shapeProperties, theme, colorMap, out color, out lineWidth, out alpha, styleLineWidth))
        {
            return true;
        }

        if (explicitLine?.Attribute("w") is { } explicitWidthAttribute &&
            lineReference.Style is not null &&
            PptxColorResolver.TryReadSolidColorWithAlpha(lineReference.Style, theme, colorMap, lineReference.Reference, out color, out alpha))
        {
            lineWidth = OoxUnits.EmuToPoints(long.Parse(explicitWidthAttribute.Value, CultureInfo.InvariantCulture));
            return true;
        }

        if (lineReference.Style is null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        if (TryReadThemeLineReference(lineReference, theme, colorMap, out PptxSceneLineStyle line))
        {
            color = line.Color;
            lineWidth = line.Width;
            alpha = line.Alpha;
            return true;
        }

        color = default;
        lineWidth = 0d;
        alpha = 1d;
        return false;
    }

    private static bool TryReadStyleLineWidth(PptxFormatSchemeReference lineReference, out double lineWidth)
    {
        if (lineReference.Style?.Attribute("w") is { } widthAttribute)
        {
            lineWidth = OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture));
            return true;
        }

        lineWidth = 0d;
        return false;
    }

    private static bool TryReadThemeLineReference(XElement? lineReference, PptxTheme theme, out PptxSceneLineStyle line)
    {
        return TryReadThemeLineReference(
            new PptxFormatSchemeReference(
                lineReference,
                PptxFormatSchemeResolver.ReadIndex(lineReference),
                null),
            theme,
            PptxColorMap.Default,
            out line);
    }

    private static bool TryReadThemeLineReference(PptxFormatSchemeReference lineReference, PptxTheme theme, out PptxSceneLineStyle line)
    {
        return TryReadThemeLineReference(lineReference, theme, PptxColorMap.Default, out line);
    }

    private static bool TryReadThemeLineReference(PptxFormatSchemeReference lineReference, PptxTheme theme, PptxColorMap colorMap, out PptxSceneLineStyle line)
    {
        XElement? lineStyle = lineReference.Style;
        if (lineStyle is null &&
            lineReference.Index > 0 &&
            theme.TryGetLineStyle(lineReference.Index, out XElement? resolvedLineStyle))
        {
            lineStyle = resolvedLineStyle;
        }

        if (lineStyle is null ||
            lineStyle.Element(DrawingNamespace + "noFill") is not null ||
            !PptxColorResolver.TryReadSolidColorWithAlpha(lineStyle, theme, colorMap, lineReference.Reference, out RgbColor color, out double alpha))
        {
            line = default;
            return false;
        }

        double lineWidth = lineStyle.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        XElement shapeProperties = new(DrawingNamespace + "spPr", new XElement(lineStyle));
        string? lineCap = ReadLineCap(shapeProperties);
        line = new PptxSceneLineStyle(
            true,
            color,
            lineWidth,
            alpha,
            TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern) ? dashPattern : [],
            ReadPresetDashValue(shapeProperties),
            ReadLineCompound(shapeProperties),
            ReadLineCompoundValue(shapeProperties),
            lineCap switch
            {
                "rnd" => 1,
                "sq" => 2,
                _ => null
            },
            lineCap,
            ReadLineJoin(shapeProperties),
            ReadLineJoinValue(shapeProperties), true);
        return true;
    }

    private static bool TryReadPresetDash(XElement? shapeProperties, double lineWidth, out IReadOnlyList<double> dashPattern)
    {
        string? presetDash = ReadPresetDashValue(shapeProperties);
        double w = Math.Max(lineWidth, MinimumStrokeWidth);
        dashPattern = presetDash switch
        {
            "dot" or "sysDot" => [w, w * 2d],
            "dash" or "sysDash" => [w * 4d, w * 3d],
            "lgDash" => [w * 8d, w * 3d],
            "dashDot" or "sysDashDot" => [w * 4d, w * 3d, w, w * 3d],
            "lgDashDot" => [w * 8d, w * 3d, w, w * 3d],
            "lgDashDotDot" or "sysDashDotDot" => [w * 8d, w * 3d, w, w * 3d, w, w * 3d],
            _ => []
        };
        return dashPattern.Count > 0;
    }

    private static string? ReadPresetDashValue(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + "prstDash")
            ?.Attribute("val");
    }

    private static PptxSceneLineCompound? ReadLineCompound(XElement? shapeProperties)
    {
        string? compound = ReadLineCompoundValue(shapeProperties);
        return compound switch
        {
            "sng" => PptxSceneLineCompound.Single,
            "dbl" => PptxSceneLineCompound.Double,
            "thickThin" => PptxSceneLineCompound.ThickThin,
            "thinThick" => PptxSceneLineCompound.ThinThick,
            "tri" => PptxSceneLineCompound.Triple,
            _ => null
        };
    }

    private static string? ReadLineCompoundValue(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Attribute("cmpd");
    }

    private static string? ReadLineCap(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Attribute("cap");
    }

    private static int? ReadLineJoin(XElement? shapeProperties)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return 1;
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return 2;
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return 0;
        }

        return null;
    }

    private static string? ReadLineJoinValue(XElement? shapeProperties)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return "round";
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return "bevel";
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return "miter";
        }

        return null;
    }

    private static PptxSceneLineEnd ReadLineEnd(XElement? shapeProperties, string elementName)
    {
        XElement? end = shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + elementName);
        string? type = (string?)end?.Attribute("type");
        string? width = (string?)end?.Attribute("w");
        string? length = (string?)end?.Attribute("len");
        return new PptxSceneLineEnd(
            ReadLineEndKind(type),
            type,
            ReadLineEndScale(width),
            width,
            ReadLineEndScale(length),
            length);
    }

    private static PptxSceneLineEndKind ReadLineEndKind(string? type)
    {
        return type switch
        {
            "triangle" => PptxSceneLineEndKind.Triangle,
            "arrow" => PptxSceneLineEndKind.Arrow,
            "stealth" => PptxSceneLineEndKind.Stealth,
            "diamond" => PptxSceneLineEndKind.Diamond,
            "oval" => PptxSceneLineEndKind.Oval,
            _ => PptxSceneLineEndKind.None
        };
    }

    private static double ReadLineEndScale(string? value)
    {
        return value switch
        {
            "sm" => 0.5d,
            "lg" => 1.5d,
            _ => 1d
        };
    }
}
