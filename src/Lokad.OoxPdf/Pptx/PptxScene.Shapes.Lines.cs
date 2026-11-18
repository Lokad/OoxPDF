using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneLineStyle ReadShapeLine(
        XElement? shapeProperties,
        PptxTheme theme,
        PptxColorMap colorMap,
        PptxFormatSchemeReference lineReference,
        bool widthSpecifiedByDefault)
    {
        XElement? explicitLine = shapeProperties?.Element(DrawingNamespace + "ln");
        if (explicitLine is null)
        {
            return lineReference.Style is not null &&
                TryReadThemeLineReference(lineReference, theme, colorMap, out PptxSceneLineStyle themed)
                ? themed
                : default;
        }

        if (explicitLine.Element(DrawingNamespace + "noFill") is not null)
        {
            return default;
        }

        XElement? style = lineReference.Style;
        double? styleLineWidth = style?.Attribute("w") is { } inheritedWidth
            ? OoxUnits.EmuToPoints(long.Parse(inheritedWidth.Value, CultureInfo.InvariantCulture))
            : null;
        bool hasLocalWidth = explicitLine.Attribute("w") is not null;
        bool widthSpecified = widthSpecifiedByDefault || hasLocalWidth;
        if (PptxLineStyleReader.TryReadLineWithAlpha(explicitLine, theme, colorMap, out RgbColor color,
                out double lineWidth, out double alpha, styleLineWidth) ||
            (hasLocalWidth && style is not null &&
                PptxColorResolver.TryReadSolidColorWithAlpha(style, theme, colorMap, lineReference.Reference, out color, out alpha)))
        {
            return ReadSceneLineStyle(explicitLine, color, lineWidth, alpha, widthSpecified);
        }

        if (style is not null && TryReadThemeLineReference(lineReference, theme, colorMap, out PptxSceneLineStyle inherited))
        {
            return ReadSceneLineStyle(explicitLine, inherited.Color, inherited.Width, inherited.Alpha, widthSpecified);
        }

        return default;
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
        line = ReadSceneLineStyle(lineStyle, color, lineWidth, alpha, widthSpecified: true);
        return true;
    }

    private static PptxSceneLineStyle ReadSceneLineStyle(
        XElement line, RgbColor color, double lineWidth, double alpha, bool widthSpecified)
    {
        string? presetDash = (string?)line.Element(DrawingNamespace + "prstDash")?.Attribute("val");
        double w = Math.Max(lineWidth, MinimumStrokeWidth);
        IReadOnlyList<double> dashPattern = presetDash switch
        {
            "dot" or "sysDot" => [w, w * 2d],
            "dash" or "sysDash" => [w * 4d, w * 3d],
            "lgDash" => [w * 8d, w * 3d],
            "dashDot" or "sysDashDot" => [w * 4d, w * 3d, w, w * 3d],
            "lgDashDot" => [w * 8d, w * 3d, w, w * 3d],
            "lgDashDotDot" or "sysDashDotDot" => [w * 8d, w * 3d, w, w * 3d, w, w * 3d],
            _ => []
        };
        string? compoundValue = (string?)line.Attribute("cmpd");
        PptxSceneLineCompound? compound = compoundValue switch
        {
            "sng" => PptxSceneLineCompound.Single,
            "dbl" => PptxSceneLineCompound.Double,
            "thickThin" => PptxSceneLineCompound.ThickThin,
            "thinThick" => PptxSceneLineCompound.ThinThick,
            "tri" => PptxSceneLineCompound.Triple,
            _ => null
        };
        string? capValue = (string?)line.Attribute("cap");
        int? cap = capValue switch
        {
            "rnd" => 1,
            "sq" => 2,
            _ => null
        };
        (int? Number, string? Token) join =
            line.Element(DrawingNamespace + "round") is not null ? (1, "round") :
            line.Element(DrawingNamespace + "bevel") is not null ? (2, "bevel") :
            line.Element(DrawingNamespace + "miter") is not null ? (0, "miter") : (null, null);

        return new PptxSceneLineStyle(true, color, lineWidth, alpha, dashPattern, presetDash,
            compound, compoundValue, cap, capValue, join.Number, join.Token, widthSpecified);
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
