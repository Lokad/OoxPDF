using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxScene(PptxDocument Document, PptxTheme Theme, IReadOnlyList<PptxSceneSlide> Slides);

internal sealed record PptxTextRunSnapshot(
    string Text,
    double X,
    double Y,
    double Width,
    double FontSize,
    double CharacterSpacing,
    RgbColor Color,
    double Alpha,
    RgbColor? Highlight,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Strike,
    string Alignment,
    string? FontFamily);

internal sealed record PptxTextFrameModelSnapshot(
    double TextX,
    double TextWidth,
    double FontScale,
    IReadOnlyList<PptxTextParagraphModelSnapshot> Paragraphs);

internal sealed record PptxTextParagraphModelSnapshot(
    int Level,
    string CascadeLevelName,
    int ResolvedCascadeSourceCount,
    IReadOnlyList<string> CascadeLayerNames,
    string Alignment,
    double FontSize,
    IReadOnlyList<PptxTextRunModelSnapshot> Runs);

internal sealed record PptxTextRunModelSnapshot(
    string Kind,
    string Text,
    double FontSize,
    double CharacterSpacing,
    string? Typeface,
    bool Underline,
    RgbColor? Highlight);

internal sealed record PptxTextLayoutSnapshot(IReadOnlyList<PptxTextFrameLayoutSnapshot> Frames);

internal sealed record PptxTextFrameLayoutSnapshot(IReadOnlyList<PptxTextParagraphLayoutSnapshot> Paragraphs);

internal sealed record PptxTextParagraphLayoutSnapshot(
    int Level,
    IReadOnlyList<PptxTextLineLayoutSnapshot> Lines);

internal sealed record PptxTextLineLayoutSnapshot(
    double TopY,
    double BaselineY,
    double Advance,
    double BaselineOffset,
    double MaxFontSize,
    string LineSpacingKind,
    double StartX,
    double EndX,
    string Alignment,
    IReadOnlyList<PptxTextSpanLayoutSnapshot> Spans);

internal sealed record PptxTextSpanLayoutSnapshot(
    string? SourceText,
    string Text,
    double X,
    double Y,
    double Width,
    double FontSize,
    IReadOnlyList<PptxTextAtomLayoutSnapshot> Atoms);

internal sealed record PptxTextAtomLayoutSnapshot(
    string Kind,
    string Text,
    double X,
    double Width,
    bool Draw);

internal sealed record PptxTextGlyphRunSnapshot(
    string Text,
    double X,
    double BaselineY,
    double Width,
    int GlyphCount,
    double FirstAdjustmentAfterOrigin);

internal sealed record PptxSceneSlide(
    int Index,
    string PartName,
    XDocument SlideXml,
    IReadOnlyList<PptxSceneNode> MasterNodes,
    IReadOnlyList<PptxSceneNode> LayoutNodes,
    IReadOnlyList<PptxSceneNode> SlideNodes);

internal sealed record PptxSceneNode(
    PptxSceneNodeKind Kind,
    string Id,
    string Name,
    bool IsPlaceholder,
    PptxSceneBounds? Bounds,
    PptxSceneTextBody? TextBody,
    XElement Source);

internal sealed record PptxSceneTextBody(
    XElement? BodyProperties,
    XElement? ListStyle,
    IReadOnlyList<PptxSceneTextParagraph> Paragraphs);

internal sealed record PptxSceneTextParagraph(
    XElement? Properties,
    XElement? EndParagraphProperties,
    int Level,
    PptxSceneParagraphStyle ResolvedStyle,
    IReadOnlyList<PptxSceneTextRun> Runs);

internal sealed record PptxSceneTextRun(
    PptxSceneTextRunKind Kind,
    string Text,
    XElement? Properties,
    PptxSceneRunStyle ResolvedStyle,
    XElement Source);

internal sealed record PptxSceneParagraphStyle(
    int Level,
    string Alignment,
    double FontSize,
    RgbColor Color,
    double Alpha,
    string? Typeface,
    bool Bold,
    bool Italic,
    double CharacterSpacing);

internal sealed record PptxSceneRunStyle(
    double FontSize,
    RgbColor Color,
    double Alpha,
    string? Typeface,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Strike,
    double CharacterSpacing,
    double BaselineOffset,
    RgbColor? Highlight);

internal enum PptxSceneTextRunKind
{
    Text,
    Break,
    Field
}

internal sealed record PptxSceneBounds(
    double X,
    double Y,
    double Width,
    double Height,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical);

internal enum PptxSceneNodeKind
{
    Shape,
    Picture,
    Table,
    Chart,
    Group,
    Connector,
    UnknownGraphicFrame,
    Unknown
}

internal sealed class PptxSceneBuilder
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string SlideLayoutRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

    public PptxScene Build(PptxDocument document, OoxPackage package)
    {
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        var slides = new List<PptxSceneSlide>(document.Slides.Count);
        foreach (PptxSlide slide in document.Slides)
        {
            OoxPart? slidePart = package.GetPart(slide.PartName);
            if (slidePart is null)
            {
                slides.Add(new PptxSceneSlide(slide.Index, slide.PartName, new XDocument(), [], [], []));
                continue;
            }

            XDocument slideXml = LoadXml(slidePart);
            OoxPart? layoutPart = GetRelatedPart(package, slide.PartName, SlideLayoutRelationshipType);
            OoxPart? masterPart = layoutPart is null ? null : GetRelatedPart(package, layoutPart.Name, SlideMasterRelationshipType);
            XDocument? masterXml = masterPart is null ? null : LoadXml(masterPart);
            XDocument? layoutXml = layoutPart is null ? null : LoadXml(layoutPart);
            IReadOnlyList<XDocument> layoutSources = masterXml is null ? [] : [masterXml];
            IReadOnlyList<XDocument> slideSources = masterXml is null
                ? layoutXml is null ? [] : [layoutXml]
                : layoutXml is null ? [masterXml] : [masterXml, layoutXml];
            slides.Add(new PptxSceneSlide(
                slide.Index,
                slide.PartName,
                slideXml,
                masterXml is null ? [] : ReadNodes(masterXml, [], theme),
                layoutXml is null ? [] : ReadNodes(layoutXml, layoutSources, theme),
                ReadNodes(slideXml, slideSources, theme)));
        }

        return new PptxScene(document, theme, slides);
    }

    private static XDocument LoadXml(OoxPart part)
    {
        using Stream stream = part.OpenRead();
        return SafeXml.Load(stream);
    }

    private static OoxPart? GetRelatedPart(OoxPackage package, string sourcePartName, string relationshipType)
    {
        OoxRelationship? relationship = package.GetRelationships(sourcePartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null ? null : package.GetPart(relationship.ResolvedTarget);
    }

    private static IReadOnlyList<PptxSceneNode> ReadNodes(XDocument xml, IReadOnlyList<XDocument> placeholderSources, PptxTheme theme)
    {
        var nodes = new List<PptxSceneNode>();
        foreach (XElement shapeTree in xml.Descendants(PresentationNamespace + "spTree"))
        {
            foreach (XElement child in shapeTree.Elements())
            {
                PptxSceneNodeKind kind = ReadKind(child);
                if (kind == PptxSceneNodeKind.Unknown)
                {
                    continue;
                }

                (string id, string name) = ReadNonVisualProperties(child);
                nodes.Add(new PptxSceneNode(kind, id, name, IsPlaceholder(child), ReadBounds(child), ReadTextBody(child, placeholderSources, theme), child));
            }
        }

        return nodes;
    }

    private static PptxSceneNodeKind ReadKind(XElement element)
    {
        if (element.Name == PresentationNamespace + "sp")
        {
            return PptxSceneNodeKind.Shape;
        }

        if (element.Name == PresentationNamespace + "cxnSp")
        {
            return PptxSceneNodeKind.Connector;
        }

        if (element.Name == PresentationNamespace + "pic")
        {
            return PptxSceneNodeKind.Picture;
        }

        if (element.Name == PresentationNamespace + "grpSp")
        {
            return PptxSceneNodeKind.Group;
        }

        if (element.Name != PresentationNamespace + "graphicFrame")
        {
            return PptxSceneNodeKind.Unknown;
        }

        XElement? graphicData = element
            .Descendants(DrawingNamespace + "graphicData")
            .FirstOrDefault();
        string uri = (string?)graphicData?.Attribute("uri") ?? string.Empty;
        if (graphicData?.Descendants(DrawingNamespace + "tbl").Any() == true)
        {
            return PptxSceneNodeKind.Table;
        }

        return uri.Contains("chart", StringComparison.OrdinalIgnoreCase)
            ? PptxSceneNodeKind.Chart
            : PptxSceneNodeKind.UnknownGraphicFrame;
    }

    private static (string Id, string Name) ReadNonVisualProperties(XElement element)
    {
        XElement? nonVisual = element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "cNvPr");
        return ((string?)nonVisual?.Attribute("id") ?? string.Empty, (string?)nonVisual?.Attribute("name") ?? string.Empty);
    }

    private static bool IsPlaceholder(XElement element)
    {
        return element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
    }

    private static PptxSceneBounds? ReadBounds(XElement element)
    {
        XElement? transform = element
            .Element(PresentationNamespace + "spPr")?
            .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "grpSpPr")?
                .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "xfrm");
        if (transform is null)
        {
            return null;
        }

        XElement? offset = transform.Element(DrawingNamespace + "off");
        XElement? extents = transform.Element(DrawingNamespace + "ext");
        if (offset is null || extents is null)
        {
            return null;
        }

        return new PptxSceneBounds(
            OoxUnits.EmuToPoints(ReadLong(offset, "x")),
            OoxUnits.EmuToPoints(ReadLong(offset, "y")),
            OoxUnits.EmuToPoints(ReadLong(extents, "cx")),
            OoxUnits.EmuToPoints(ReadLong(extents, "cy")),
            transform.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            ReadBool(transform, "flipH"),
            ReadBool(transform, "flipV"));
    }

    private static PptxSceneTextBody? ReadTextBody(XElement element, IReadOnlyList<XDocument> placeholderSources, PptxTheme theme)
    {
        XElement? textBody = element.Element(PresentationNamespace + "txBody");
        if (textBody is null)
        {
            return null;
        }

        IReadOnlyList<XElement> inheritedTextBodies = FindInheritedPlaceholderShapes(element, placeholderSources)
            .Select(shape => shape.Element(PresentationNamespace + "txBody"))
            .Where(textBody => textBody is not null)
            .Cast<XElement>()
            .ToArray();
        return new PptxSceneTextBody(
            textBody.Element(DrawingNamespace + "bodyPr"),
            textBody.Element(DrawingNamespace + "lstStyle"),
            textBody.Elements(DrawingNamespace + "p").Select(paragraph => ReadParagraph(paragraph, element, textBody, inheritedTextBodies, placeholderSources, theme)).ToArray());
    }

    private static PptxSceneTextParagraph ReadParagraph(
        XElement paragraph,
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedTextBodies,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme)
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
        PptxSceneParagraphStyle resolvedStyle = ResolveParagraphStyle(level, properties, defaultParagraphProperties, defaultRunProperties, shape, theme);
        return new PptxSceneTextParagraph(
            properties,
            paragraph.Element(DrawingNamespace + "endParaRPr"),
            level,
            resolvedStyle,
            paragraph.Elements().Select(run => ReadRun(run, defaultRunProperties, resolvedStyle, theme)).Where(run => run is not null).Cast<PptxSceneTextRun>().ToArray());
    }

    private static PptxSceneTextRun? ReadRun(XElement element, XElement? defaultRunProperties, PptxSceneParagraphStyle paragraphStyle, PptxTheme theme)
    {
        if (element.Name == DrawingNamespace + "r")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(
                PptxSceneTextRunKind.Text,
                (string?)element.Element(DrawingNamespace + "t") ?? string.Empty,
                runProperties,
                ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme),
                element);
        }

        if (element.Name == DrawingNamespace + "br")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(PptxSceneTextRunKind.Break, "\n", runProperties, ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme), element);
        }

        if (element.Name == DrawingNamespace + "fld")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(
                PptxSceneTextRunKind.Field,
                (string?)element.Element(DrawingNamespace + "t") ?? string.Empty,
                runProperties,
                ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme),
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
        XElement? merged = null;
        foreach (XElement source in sources.Reverse().Where(source => source is not null).Cast<XElement>())
        {
            merged ??= new XElement(source.Name);
            foreach (XAttribute attribute in source.Attributes())
            {
                merged.SetAttributeValue(attribute.Name, attribute.Value);
            }

            foreach (XElement child in source.Elements())
            {
                XElement? existing = merged.Element(child.Name);
                if (existing is null)
                {
                    merged.Add(new XElement(child));
                }
                else
                {
                    MergeElementInto(existing, child);
                }
            }
        }

        return merged;
    }

    private static void MergeElementInto(XElement target, XElement source)
    {
        foreach (XAttribute attribute in source.Attributes())
        {
            target.SetAttributeValue(attribute.Name, attribute.Value);
        }

        foreach (XElement child in source.Elements())
        {
            target.Elements(child.Name).Remove();
            target.Add(new XElement(child));
        }
    }

    private static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        string? placeholderType = (string?)shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph")
            ?.Attribute("type");
        string styleName = placeholderType switch
        {
            "title" or "ctrTitle" => "titleStyle",
            "body" or "subTitle" => "bodyStyle",
            _ => "otherStyle"
        };

        foreach (XDocument source in placeholderSources)
        {
            XElement? style = source.Root?
                .Element(PresentationNamespace + "txStyles")
                ?.Element(PresentationNamespace + styleName);
            XElement? level = style?.Element(DrawingNamespace + levelName) ??
                style?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    private static XElement? FindDefaultTextStyle(IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        foreach (XDocument source in placeholderSources)
        {
            XElement? defaultTextStyle = source.Root?.Element(PresentationNamespace + "defaultTextStyle");
            XElement? level = defaultTextStyle?.Element(DrawingNamespace + levelName) ??
                defaultTextStyle?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    private static IReadOnlyList<XElement> FindInheritedPlaceholderShapes(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        XElement? placeholder = shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph");
        if (placeholder is null)
        {
            return [];
        }

        var matches = new List<XElement>();
        string? type = (string?)placeholder.Attribute("type");
        string? index = (string?)placeholder.Attribute("idx");
        foreach (XDocument source in placeholderSources)
        {
            foreach (XElement candidate in source.Descendants(PresentationNamespace + "sp"))
            {
                XElement? candidatePlaceholder = candidate
                    .Element(PresentationNamespace + "nvSpPr")
                    ?.Element(PresentationNamespace + "nvPr")
                    ?.Element(PresentationNamespace + "ph");
                if (candidatePlaceholder is null)
                {
                    continue;
                }

                string? candidateType = (string?)candidatePlaceholder.Attribute("type");
                string? candidateIndex = (string?)candidatePlaceholder.Attribute("idx");
                bool indexMatches = index is not null && candidateIndex == index;
                bool typeMatches = index is null && type is not null && candidateType == type;
                if (indexMatches || typeMatches)
                {
                    matches.Add(candidate);
                    break;
                }
            }
        }

        return matches;
    }

    private static PptxSceneParagraphStyle ResolveParagraphStyle(
        int level,
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        XElement? defaultRunProperties,
        XElement shape,
        PptxTheme theme)
    {
        double fontSize = ReadFontSize(defaultRunProperties, null);
        RgbColor color = TryReadSolidColorWithAlpha(defaultRunProperties, theme, out RgbColor defaultColor, out double alpha)
            ? defaultColor
            : TryReadShapeFontColor(shape, theme, out RgbColor shapeColor)
                ? shapeColor
                : new RgbColor(0, 0, 0);
        if (!TryReadSolidColorWithAlpha(defaultRunProperties, theme, out _, out alpha))
        {
            alpha = 1d;
        }

        string? typeface = theme.ResolveTypeface((string?)defaultRunProperties?.Element(DrawingNamespace + "latin")?.Attribute("typeface"));
        return new PptxSceneParagraphStyle(
            level,
            (string?)(paragraphProperties?.Attribute("algn") ?? defaultParagraphProperties?.Attribute("algn")) ?? "l",
            fontSize,
            color,
            alpha,
            typeface,
            ParseOptionalBoolAttribute(defaultRunProperties, "b"),
            ParseOptionalBoolAttribute(defaultRunProperties, "i"),
            ReadCharacterSpacing(defaultRunProperties, null));
    }

    private static PptxSceneRunStyle ResolveRunStyle(XElement? runProperties, XElement? defaultRunProperties, PptxSceneParagraphStyle paragraphStyle, PptxTheme theme)
    {
        double fontSize = ReadFontSize(runProperties, defaultRunProperties);
        double alpha = paragraphStyle.Alpha;
        RgbColor color = paragraphStyle.Color;
        if (TryReadSolidColorWithAlpha(runProperties, theme, out RgbColor runColor, out double runAlpha))
        {
            color = runColor;
            alpha = runAlpha;
        }
        else if (TryReadSolidColorWithAlpha(defaultRunProperties, theme, out RgbColor defaultColor, out double defaultAlpha))
        {
            color = defaultColor;
            alpha = defaultAlpha;
        }

        string? typeface = theme.ResolveTypeface((string?)(runProperties?.Element(DrawingNamespace + "latin") ??
            defaultRunProperties?.Element(DrawingNamespace + "latin"))?.Attribute("typeface")) ?? paragraphStyle.Typeface;
        bool bold = ParseOptionalBoolAttribute(runProperties, "b") ||
            (runProperties?.Attribute("b") is null && paragraphStyle.Bold);
        bool italic = ParseOptionalBoolAttribute(runProperties, "i") ||
            (runProperties?.Attribute("i") is null && paragraphStyle.Italic);
        bool underline = ((string?)(runProperties?.Attribute("u") ?? defaultRunProperties?.Attribute("u"))) is { } underlineValue &&
            !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
        bool strike = IsStrikeEnabled(runProperties, defaultRunProperties);
        return new PptxSceneRunStyle(
            fontSize,
            color,
            alpha,
            typeface,
            bold,
            italic,
            underline,
            strike,
            ReadCharacterSpacing(runProperties, defaultRunProperties),
            ReadBaselineOffset(runProperties, defaultRunProperties, fontSize),
            TryReadHighlightColor(runProperties, out RgbColor highlight) ? highlight : null);
    }

    private static double ReadFontSize(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("sz") ?? defaultRunProperties?.Attribute("sz")) is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
            : 18d;
    }

    private static double ReadCharacterSpacing(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("spc") ?? defaultRunProperties?.Attribute("spc")) is { } spacing
            ? int.Parse(spacing.Value, CultureInfo.InvariantCulture) / 100d
            : 0d;
    }

    private static double ReadBaselineOffset(XElement? runProperties, XElement? defaultRunProperties, double fontSize)
    {
        return (runProperties?.Attribute("baseline") ?? defaultRunProperties?.Attribute("baseline")) is { } baseline
            ? fontSize * int.Parse(baseline.Value, CultureInfo.InvariantCulture) / 100000d
            : 0d;
    }

    private static bool IsStrikeEnabled(XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = (string?)(runProperties?.Attribute("strike") ?? defaultRunProperties?.Attribute("strike"));
        return value is not null && !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ParseOptionalBoolAttribute(XElement? element, string attributeName)
    {
        string? value = (string?)element?.Attribute(attributeName);
        return value is "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryReadHighlightColor(XElement? runProperties, out RgbColor color)
    {
        XElement? highlight = runProperties?.Element(DrawingNamespace + "highlight");
        string? hex = (string?)highlight?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        return RgbColor.TryParse(hex, out color);
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, out RgbColor color)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return TryReadSolidColorWithAlpha(fontRef, theme, out color, out _);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, out RgbColor color, out double alpha)
    {
        XElement? solidFill = element?.Element(DrawingNamespace + "solidFill");
        XElement? colorContainer = solidFill ?? element;
        alpha = ReadAlpha(colorContainer);
        XElement? srgbColor = colorContainer?.Element(DrawingNamespace + "srgbClr");
        if (RgbColor.TryParse((string?)srgbColor?.Attribute("val"), out color))
        {
            color = ApplyColorTransforms(srgbColor, color);
            return true;
        }

        XElement? schemeColorElement = colorContainer?.Element(DrawingNamespace + "schemeClr");
        string? schemeColor = (string?)schemeColorElement?.Attribute("val");
        if (schemeColor is not null && theme.TryResolveColor(schemeColor, out color))
        {
            color = ApplyColorTransforms(schemeColorElement, color);
            return true;
        }

        return false;
    }

    private static double ReadAlpha(XElement? colorContainer)
    {
        XElement? alpha = colorContainer?
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName is "srgbClr" or "schemeClr")
            ?.Element(DrawingNamespace + "alpha");
        if (alpha?.Attribute("val") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
        {
            return Math.Clamp(parsed / 100000d, 0d, 1d);
        }

        return 1d;
    }

    private static RgbColor ApplyColorTransforms(XElement? colorElement, RgbColor color)
    {
        if (colorElement is null)
        {
            return color;
        }

        double red = color.Red;
        double green = color.Green;
        double blue = color.Blue;
        foreach (XElement transform in colorElement.Elements())
        {
            double value = ReadLong(transform, "val", 100000) / 100000d;
            switch (transform.Name.LocalName)
            {
                case "lumMod":
                case "shade":
                    red *= value;
                    green *= value;
                    blue *= value;
                    break;
                case "lumOff":
                    red += 255d * value;
                    green += 255d * value;
                    blue += 255d * value;
                    break;
                case "tint":
                    red += (255d - red) * value;
                    green += (255d - green) * value;
                    blue += (255d - blue) * value;
                    break;
            }
        }

        return new RgbColor(ToByte(red), ToByte(green), ToByte(blue));
    }

    private static byte ToByte(double value)
    {
        return (byte)Math.Clamp((int)Math.Round(value, MidpointRounding.AwayFromZero), 0, 255);
    }

    private static long ReadLong(XElement element, string name)
    {
        return element.Attribute(name) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : 0L;
    }

    private static long ReadLong(XElement element, string name, long defaultValue)
    {
        return element.Attribute(name) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultValue;
    }

    private static bool ReadBool(XElement element, string name)
    {
        string? value = (string?)element.Attribute(name);
        return value is "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }
}
