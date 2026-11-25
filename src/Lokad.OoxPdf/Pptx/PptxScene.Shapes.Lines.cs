using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneLineStyle ReadShapeLine(
        XElement? explicitLine,
        PptxTheme theme,
        PptxColorMap colorMap,
        PptxFormatSchemeReference lineReference,
        bool widthSpecifiedByDefault)
    {
        if (explicitLine?.Element(DrawingNamespace + "noFill") is not null)
        {
            return default;
        }

        XElement? themeLine = lineReference.Style;
        double? localWidth = PptxLineStyleReader.ReadWidth(explicitLine);
        bool hasLocalColor = PptxColorResolver.TryReadSolidColorWithAlpha(explicitLine, theme, colorMap, out RgbColor color, out double alpha);
        if (!hasLocalColor && !TryReadThemeLineColor(themeLine, lineReference.Reference, theme, colorMap, out color, out alpha))
        {
            return default;
        }

        double lineWidth = localWidth ?? PptxLineStyleReader.ReadWidth(themeLine) ??
            (hasLocalColor ? PptxLineStyleReader.OfficeDefaultLineWidthPoints : 1d);
        return ReadSceneLineStyle(explicitLine, color, lineWidth, alpha,
            widthSpecifiedByDefault || localWidth.HasValue, themeLine);
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

        if (!TryReadThemeLineColor(lineStyle, lineReference.Reference, theme, colorMap, out RgbColor color, out double alpha))
        {
            line = default;
            return false;
        }

        double lineWidth = PptxLineStyleReader.ReadWidth(lineStyle) ?? 1d;
        line = ReadSceneLineStyle(lineStyle, color, lineWidth, alpha, widthSpecified: true);
        return true;
    }

    private static bool TryReadThemeLineColor(XElement? line, XElement? reference, PptxTheme theme, PptxColorMap colorMap,
        out RgbColor color, out double alpha)
    {
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            alpha = 0d;
            return false;
        }

        return PptxColorResolver.TryReadSolidColorWithAlpha(line, theme, colorMap, reference, out color, out alpha);
    }

    private static PptxSceneLineStyle ReadSceneLineStyle(
        XElement? line, RgbColor color, double lineWidth, double alpha, bool widthSpecified, XElement? themeLine = null)
    {
        // Resolve the dash choice before scaling its segments to the effective width.
        XElement? dash = line?.Element(DrawingNamespace + "prstDash");
        if (dash is null && line?.Element(DrawingNamespace + "custDash") is null)
        {
            dash = themeLine?.Element(DrawingNamespace + "prstDash");
        }

        string? presetDash = (string?)dash?.Attribute("val");
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
        string? compoundValue = (string?)(line?.Attribute("cmpd") ?? themeLine?.Attribute("cmpd"));
        PptxSceneLineCompound? compound = compoundValue switch
        {
            "sng" => PptxSceneLineCompound.Single,
            "dbl" => PptxSceneLineCompound.Double,
            "thickThin" => PptxSceneLineCompound.ThickThin,
            "thinThick" => PptxSceneLineCompound.ThinThick,
            "tri" => PptxSceneLineCompound.Triple,
            _ => null
        };
        string? capValue = (string?)(line?.Attribute("cap") ?? themeLine?.Attribute("cap"));
        int? cap = capValue switch
        {
            "rnd" => 1,
            "sq" => 2,
            _ => null
        };
        var join = ReadJoin(line) ?? ReadJoin(themeLine);

        return new PptxSceneLineStyle(true, color, lineWidth, alpha, dashPattern, presetDash,
            compound, compoundValue, cap, capValue, join?.Number, join?.Token, widthSpecified);

        static (int Number, string Token)? ReadJoin(XElement? owner) =>
            owner?.Element(DrawingNamespace + "round") is not null ? (1, "round") :
            owner?.Element(DrawingNamespace + "bevel") is not null ? (2, "bevel") :
            owner?.Element(DrawingNamespace + "miter") is not null ? (0, "miter") : null;
    }

    private static PptxSceneLineEnd ReadLineEnd(XElement? line, XElement? themeLine, string elementName)
    {
        XName name = DrawingNamespace + elementName;
        XElement? end = line?.Element(name);
        XElement? themeEnd = themeLine?.Element(name);
        return new PptxSceneLineEnd(
            ReadLineEndKind((string?)(end?.Attribute("type") ?? themeEnd?.Attribute("type"))),
            ReadLineEndScale((string?)(end?.Attribute("w") ?? themeEnd?.Attribute("w"))),
            ReadLineEndScale((string?)(end?.Attribute("len") ?? themeEnd?.Attribute("len"))));
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
