using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using static Lokad.OoxPdf.Pptx.PptxRunTextAttributeReaders;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneTextBody? ReadTextBody(XElement element, IReadOnlyList<XDocument> placeholderSources, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? textBody = element.Element(PresentationNamespace + "txBody");
        if (textBody is null)
        {
            return null;
        }

        IReadOnlyList<XElement> inheritedPlaceholderShapes = FindInheritedPlaceholderShapes(element, placeholderSources);
        IReadOnlyList<XElement> inheritedTextBodies = inheritedPlaceholderShapes
            .Select(shape => shape.Element(PresentationNamespace + "txBody"))
            .Where(textBody => textBody is not null)
            .Cast<XElement>()
            .ToArray();
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        return new PptxSceneTextBody(
            bodyProperties,
            textBody.Element(DrawingNamespace + "lstStyle"),
            HasUnsupportedTextOrientation(bodyProperties),
            HasUnsupportedTextVerticalOverflow(bodyProperties),
            textBody.Elements(DrawingNamespace + "p").Select(paragraph => ReadParagraph(paragraph, element, textBody, inheritedTextBodies, inheritedPlaceholderShapes, placeholderSources, theme, colorMap)).ToArray());
    }

    private static bool HasUnsupportedTextOrientation(XElement? bodyProperties)
    {
        string? orientation = (string?)bodyProperties?.Attribute("vert");
        return !string.IsNullOrEmpty(orientation) &&
            !orientation.Equals("horz", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("vert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("vert270", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("eaVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("mongolianVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("wordArtVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("wordArtVertRtl", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnsupportedTextVerticalOverflow(XElement? bodyProperties)
    {
        string? overflow = (string?)bodyProperties?.Attribute("vertOverflow");
        return overflow?.Equals("ellipsis", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static PptxSceneTextParagraph ReadParagraph(
        XElement paragraph,
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedTextBodies,
        IReadOnlyList<XElement> inheritedPlaceholderShapes,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        PptxColorMap colorMap)
    {
        XElement? properties = paragraph.Element(DrawingNamespace + "pPr");
        int level = properties?.Attribute("lvl") is { } levelAttribute
            ? int.Parse(levelAttribute.Value, CultureInfo.InvariantCulture)
            : 0;
        XElement? defaultParagraphProperties = ResolveDefaultParagraphProperties(
            level,
            shape,
            textBody,
            inheritedTextBodies,
            placeholderSources);
        XElement? defaultRunProperties = properties?.Element(DrawingNamespace + "defRPr") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
        PptxSceneParagraphStyle resolvedStyle = ResolveParagraphStyle(level, properties, defaultParagraphProperties, defaultRunProperties, shape, theme, colorMap);
        string levelName = "lvl" + Math.Clamp(level + 1, 1, 9).ToString(CultureInfo.InvariantCulture) + "pPr";
        var cascadeLayers = new List<PptxSceneCascadeLayer>
        {
            new("shape.lstStyle", "ShapeListStyle", textBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)),
        };
        cascadeLayers.AddRange(inheritedPlaceholderShapes
            .Select((placeholder, index) => new PptxSceneCascadeLayer(
                PptxTextStyleInheritance.PlaceholderListStyleLayerName(placeholder, index, placeholderSources.Count),
                PptxTextStyleInheritance.PlaceholderListStyleLayerKindName(placeholder, index, placeholderSources.Count),
                placeholder.Element(PresentationNamespace + "txBody")?.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)))
            .Reverse());
        cascadeLayers.Add(new PptxSceneCascadeLayer(
            "inherited.txStyle",
            "InheritedTextStyle",
            PptxTextStyleInheritance.FindInheritedTextStyle(shape, placeholderSources, levelName))
);
        cascadeLayers.Add(new PptxSceneCascadeLayer(
            "defaultTextStyle",
            "DefaultTextStyle",
            PptxTextStyleInheritance.FindDefaultTextStyle(placeholderSources, levelName))
);
        return new PptxSceneTextParagraph(
            properties,
            paragraph.Element(DrawingNamespace + "endParaRPr"),
            level,
            resolvedStyle,
            paragraph.Elements().Select(run => ReadRun(run, defaultRunProperties, resolvedStyle, theme, colorMap)).Where(run => run is not null).Cast<PptxSceneTextRun>().ToArray(),
            defaultParagraphProperties,
            defaultRunProperties,
            cascadeLayers,
            paragraph);
    }

    private static PptxSceneTextRun? ReadRun(XElement element, XElement? defaultRunProperties, PptxSceneParagraphStyle paragraphStyle, PptxTheme theme, PptxColorMap colorMap)
    {
        if (element.Name == DrawingNamespace + "r")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(
                PptxSceneTextRunKind.Text,
                (string?)element.Element(DrawingNamespace + "t") ?? string.Empty,
                runProperties,
                ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme, colorMap),
                element);
        }

        if (element.Name == DrawingNamespace + "br")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(PptxSceneTextRunKind.Break, "\n", runProperties, ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme, colorMap), element);
        }

        if (element.Name == DrawingNamespace + "fld")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(
                PptxSceneTextRunKind.Field,
                (string?)element.Element(DrawingNamespace + "t") ?? string.Empty,
                runProperties,
                ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme, colorMap),
                element);
        }

        return null;
    }

    private static XElement? ResolveDefaultParagraphProperties(
        int level,
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedTextBodies,
        IReadOnlyList<XDocument> placeholderSources)
    {
        string levelName = $"lvl{Math.Clamp(level + 1, 1, 9).ToString(CultureInfo.InvariantCulture)}pPr";
        var sources = new List<XElement?>();
        sources.Add(textBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName));
        sources.AddRange(inheritedTextBodies
            .Reverse()
            .Select(inheritedTextBody => inheritedTextBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)));
        sources.Add(FindInheritedTextStyle(shape, placeholderSources, levelName));
        sources.Add(FindDefaultTextStyle(placeholderSources, levelName));
        return MergeParagraphProperties(sources.ToArray());
    }

    private static XElement? MergeParagraphProperties(params XElement?[] sources)
    {
        return PptxParagraphPropertyMerger.MergeSceneDefaultProperties(sources);
    }

    private static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        return PptxTextStyleInheritance.FindInheritedTextStyle(shape, placeholderSources, levelName);
    }

    private static XElement? FindDefaultTextStyle(IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        return PptxTextStyleInheritance.FindDefaultTextStyle(placeholderSources, levelName);
    }

    private static IReadOnlyList<XElement> FindInheritedPlaceholderShapes(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        return PptxPlaceholderMatcher.FindInheritedPlaceholderShapes(shape, placeholderSources);
    }

    private static PptxSceneParagraphStyle ResolveParagraphStyle(
        int level,
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        XElement? defaultRunProperties,
        XElement shape,
        PptxTheme theme,
        PptxColorMap colorMap)
    {
        double fontSize = ReadFontSize(defaultRunProperties, null);
        RgbColor color = PptxColorResolver.TryReadSolidColorWithAlpha(defaultRunProperties, theme, colorMap, out RgbColor defaultColor, out double alpha)
            ? defaultColor
            : TryReadShapeFontColor(shape, theme, colorMap, out RgbColor shapeColor)
                ? shapeColor
                : new RgbColor(0, 0, 0);
        if (!PptxColorResolver.TryReadSolidColorWithAlpha(defaultRunProperties, theme, colorMap, out _, out alpha))
        {
            alpha = 1d;
        }

        PptxThemeTypefaceResolution typeface = theme.ResolveTypefaceWithSource((string?)defaultRunProperties?.Element(DrawingNamespace + "latin")?.Attribute("typeface"));
        return new PptxSceneParagraphStyle(
            level,
            (string?)(paragraphProperties?.Attribute("algn") ?? defaultParagraphProperties?.Attribute("algn")) ?? "l",
            fontSize,
            color,
            alpha,
            typeface.Typeface,
            typeface.Source,
            OoxXml.ParseOptionalBool(defaultRunProperties, "b"),
            OoxXml.ParseOptionalBool(defaultRunProperties, "i"),
            ReadCharacterSpacing(defaultRunProperties, null));
    }

    private static PptxSceneRunStyle ResolveRunStyle(XElement? runProperties, XElement? defaultRunProperties, PptxSceneParagraphStyle paragraphStyle, PptxTheme theme, PptxColorMap colorMap)
    {
        double fontSize = ReadFontSize(runProperties, defaultRunProperties);
        double alpha = paragraphStyle.Alpha;
        RgbColor color = paragraphStyle.Color;
        bool hasHyperlinkClick = HasRunHyperlinkClick();
        if (hasHyperlinkClick && theme.TryResolveColor("hlink", colorMap, out RgbColor hyperlinkColor))
        {
            color = hyperlinkColor;
            alpha = 1d;
        }
        else if (PptxColorResolver.TryReadSolidColorWithAlpha(runProperties, theme, colorMap, out RgbColor runColor, out double runAlpha))
        {
            color = runColor;
            alpha = runAlpha;
        }
        else if (PptxColorResolver.TryReadSolidColorWithAlpha(defaultRunProperties, theme, colorMap, out RgbColor defaultColor, out double defaultAlpha))
        {
            color = defaultColor;
            alpha = defaultAlpha;
        }

        string? requestedTypeface = (string?)(runProperties?.Element(DrawingNamespace + "latin") ??
            defaultRunProperties?.Element(DrawingNamespace + "latin"))?.Attribute("typeface");
        PptxThemeTypefaceResolution? typeface = string.IsNullOrWhiteSpace(requestedTypeface)
            ? null
            : theme.ResolveTypefaceWithSource(requestedTypeface);
        bool bold = OoxXml.ParseOptionalBool(runProperties, "b") ||
            (runProperties?.Attribute("b") is null && paragraphStyle.Bold);
        bool italic = OoxXml.ParseOptionalBool(runProperties, "i") ||
            (runProperties?.Attribute("i") is null && paragraphStyle.Italic);
        string? underlineValue = ReadUnderlineValue(runProperties, defaultRunProperties);
        string? strikeValue = ReadStrikeValue(runProperties, defaultRunProperties);
        string? capsValue = ReadTextCapsValue(runProperties, defaultRunProperties);
        bool underline = underlineValue is null
            ? hasHyperlinkClick
            : !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
        bool strike = IsStrikeEnabled(strikeValue);
        return new PptxSceneRunStyle(
            fontSize,
            color,
            alpha,
            typeface?.Typeface ?? paragraphStyle.Typeface,
            typeface?.Source ?? paragraphStyle.TypefaceSource,
            bold,
            italic,
            underline,
            underlineValue ?? (hasHyperlinkClick ? "sng" : null),
            strike,
            strikeValue,
            capsValue,
            ReadCharacterSpacing(runProperties, defaultRunProperties),
            ReadBaselineOffset(runProperties, defaultRunProperties, fontSize),
            TryReadHighlightColor(runProperties, out RgbColor highlight) ? highlight : null);

        bool HasRunHyperlinkClick()
        {
            return runProperties?.Element(DrawingNamespace + "hlinkClick") is not null;
        }
    }

    private static double? ReadOptionalChartCharacterSpacing(XElement? runProperties)
    {
        return runProperties?.Attribute("spc") is { } spacing &&
            int.TryParse(spacing.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int spacingHundredths)
                ? spacingHundredths / 100d
                : null;
    }

    private static bool TryReadHighlightColor(XElement? runProperties, out RgbColor color)
    {
        XElement? highlight = runProperties?.Element(DrawingNamespace + "highlight");
        string? hex = (string?)highlight?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        return RgbColor.TryParse(hex, out color);
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, PptxColorMap colorMap, out RgbColor color)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return PptxColorResolver.TryReadSolidColorWithAlpha(fontRef, theme, colorMap, out color, out _);
    }
}
