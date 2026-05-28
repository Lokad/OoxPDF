using System.Globalization;
using System.Xml.Linq;

using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static IReadOnlyList<PptxTextFrameModelSnapshot> InspectTextFrameModels(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null);
        if (context is null)
        {
            return [];
        }

        return context.InheritedXml
            .SelectMany(xml => BuildTextFrameModels(context, xml, includePlaceholders: false, placeholderSources: []))
            .Concat(BuildTextFrameModels(context, context.SlideXml, includePlaceholders: true, context.InheritedXml))
            .Select(ToSnapshot)
            .ToArray();
    }

    private static PptxTextFrameModelSnapshot ToSnapshot(PptxTextFrameModel frame)
    {
        return new PptxTextFrameModelSnapshot(
            frame.TextX,
            frame.TextWidth,
            frame.FontScale,
            frame.BodyProperties.Orientation.ToString(),
            frame.BodyProperties.OrientationValue,
            frame.BodyProperties.VerticalAnchor.ToString(),
            frame.BodyProperties.VerticalAnchorValue,
            frame.BodyProperties.AnchorCenter,
            frame.BodyProperties.AnchorCenterValue,
            frame.BodyProperties.WrapMode.ToString(),
            frame.BodyProperties.WrapValue,
            frame.BodyProperties.VerticalOverflow.ToString(),
            frame.BodyProperties.VerticalOverflowValue,
            frame.BodyProperties.ColumnCount,
            frame.BodyProperties.ColumnSpacing,
            frame.Paragraphs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextParagraphModelSnapshot ToSnapshot(PptxTextParagraphModel paragraph)
    {
        return new PptxTextParagraphModelSnapshot(
            paragraph.Level,
            paragraph.Cascade.LevelName,
            paragraph.Cascade.Sources.Count(source => source is not null),
            paragraph.Cascade.Layers.Select(layer => layer.Name).ToArray(),
            paragraph.Cascade.Layers.Select(layer => layer.Kind.ToString()).ToArray(),
            paragraph.Style.Alignment.ToString(),
            paragraph.Style.AlignmentValue,
            paragraph.Style.FontSize,
            paragraph.Style.Indent.MarginLeft,
            paragraph.Style.Indent.Hanging,
            paragraph.Runs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextRunModelSnapshot ToSnapshot(PptxTextRunModel run)
    {
        return new PptxTextRunModelSnapshot(
            run.Kind.ToString(),
            run.Text,
            run.Cascade.Sources.Count(source => source is not null),
            run.Cascade.Layers.Select(layer => layer.Name).ToArray(),
            run.Cascade.Layers.Select(layer => layer.Kind.ToString()).ToArray(),
            run.Style.FontSize,
            run.Style.CharacterSpacing,
            run.Style.Typeface,
            run.Style.HasHyperlinkClick,
            run.Style.HyperlinkClickId,
            run.Style.Underline,
            run.Style.UnderlineValue,
            run.Style.Strike,
            run.Style.StrikeValue,
            run.Style.CapsValue,
            run.Style.Highlight);
    }

    private static IReadOnlyList<PptxTextFrameModel> BuildTextFrameModels(
        PptxRenderContext context,
        XDocument slideXml,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return BuildTextFrameModels(slideXml, context.Document, context.Theme, context.SlideNumber, includePlaceholders, placeholderSources);
    }

    private static IReadOnlyList<PptxTextFrameModel> BuildTextFrameModels(
        XDocument slideXml,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        var frames = new List<PptxTextFrameModel>();
        foreach (XElement shape in slideXml.Descendants(PresentationNamespace + "sp"))
        {
            PptxTextFrameModel? frame = BuildTextFrameModel(shape, document, theme, slideNumber, includePlaceholders, placeholderSources);
            if (frame is not null)
            {
                frames.Add(frame);
            }
        }

        return frames;
    }

    private static PptxTextFrameModel? BuildTextFrameModel(
        XElement shape,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        if (!includePlaceholders && IsPlaceholder(shape))
        {
            return null;
        }

        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        XElement? textBody = shape.Element(PresentationNamespace + "txBody");
        IReadOnlyList<XElement> inheritedPlaceholders = FindInheritedPlaceholderShapes(shape, placeholderSources);
        XElement? inheritedPlaceholder = inheritedPlaceholders.LastOrDefault();
        XElement? inheritedTextBody = inheritedPlaceholder?.Element(PresentationNamespace + "txBody");
        ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
        bounds ??= inheritedPlaceholder?.Element(PresentationNamespace + "spPr") is { } inheritedProperties
            ? ReadBounds(inheritedProperties)
            : null;
        if (bounds is null || textBody is null)
        {
            return null;
        }

        bounds = ReadAncestorGroupTransform(shape).Apply(bounds.Value);
        double x = OoxUnits.EmuToPoints(bounds.Value.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Value.Y);
        double width = OoxUnits.EmuToPoints(bounds.Value.Width);
        double height = OoxUnits.EmuToPoints(bounds.Value.Height);
        PptxTextBodyProperties bodyProperties = ReadTextBodyProperties(textBody, inheritedTextBody);
        TextInsets insets = bodyProperties.Insets;
        PptxTextOrientation orientation = bodyProperties.Orientation;
        double fontScale = bodyProperties.FontScale;
        double lineSpacingScale = bodyProperties.LineSpacingScale;
        bool compatibleLineSpacing = bodyProperties.CompatibleLineSpacing;
        double rotationCenterX = x + width / 2d;
        double rotationCenterY = document.SlideHeightPoints - yTop - height / 2d;
        double flowX = x;
        double flowYTop = yTop;
        double flowWidth = width;
        double flowHeight = height;
        TextInsets presetTextInsets = ReadPresetTextRectInsets(shape, width, height);
        if (!presetTextInsets.IsEmpty)
        {
            flowX += presetTextInsets.Left;
            flowYTop += presetTextInsets.Top;
            flowWidth = Math.Max(1d, flowWidth - presetTextInsets.Left - presetTextInsets.Right);
            flowHeight = Math.Max(1d, flowHeight - presetTextInsets.Top - presetTextInsets.Bottom);
        }

        double? explicitTextRotationDegrees = bodyProperties.RotationDegrees;
        double textRotationDegrees = explicitTextRotationDegrees ?? bounds.Value.RotationDegrees;
        if (explicitTextRotationDegrees is null && bounds.Value.FlipHorizontal != bounds.Value.FlipVertical)
        {
            textRotationDegrees = NormalizeRotationDegrees(textRotationDegrees + 180d);
        }

        bool textFlipHorizontal = false;
        bool textFlipVertical = false;
        if (orientation is PptxTextOrientation.Vertical or
            PptxTextOrientation.Vertical270 or
            PptxTextOrientation.EastAsianVertical or
            PptxTextOrientation.MongolianVertical or
            PptxTextOrientation.WordArtVertical or
            PptxTextOrientation.WordArtVerticalRightToLeft)
        {
            flowWidth = height;
            flowHeight = width;
            flowX = rotationCenterX - flowWidth / 2d;
            flowYTop = document.SlideHeightPoints - rotationCenterY - flowHeight / 2d;
            textRotationDegrees += TextOrientationRotationDegrees(orientation);
        }

        double textX = flowX + insets.Left;
        double textWidth = Math.Max(1d, flowWidth - insets.Left - insets.Right);
        double textWrapWidth = bodyProperties.ExplicitWrapWidth ?? textWidth;
        double textHeight = Math.Max(1d, flowHeight - insets.Top - insets.Bottom);
        int columnCount = bodyProperties.ColumnCount;
        double columnSpacing = bodyProperties.ColumnSpacing;
        bool clipsVerticalOverflow = bodyProperties.VerticalOverflow == PptxTextVerticalOverflow.Clip;
        bool canOverflowToSlideClip = columnCount == 1;
        bool clipsTextLocally = clipsVerticalOverflow || !canOverflowToSlideClip;
        double textClipX = clipsTextLocally ? textX : 0d;
        double textClipWidth = clipsTextLocally ? textWidth : document.SlideWidthPoints;
        double textClipY = 0d;
        double textClipHeight = document.SlideHeightPoints;
        if (clipsVerticalOverflow)
        {
            (textClipY, textClipHeight) = IntersectVerticalTextClipWithSlide(
                document.SlideHeightPoints - flowYTop - insets.Top - textHeight,
                textHeight,
                document.SlideHeightPoints);
        }
        RgbColor? shapeFontColor = TryReadShapeFontColor(shape, theme, out RgbColor fontColor)
            ? fontColor
            : null;
        bool useOfficeBaselineFloor = TextFrameUsesOfficeBaselineFloor(shape, bodyProperties);
        XElement? defaultParagraphProperties = MergeParagraphProperties(
            BuildParagraphStyleCascade(shape, textBody, inheritedPlaceholders, placeholderSources, "lvl1pPr").Sources.ToArray());
        double verticalOffset = bodyProperties.VerticalAnchor switch
        {
            TextVerticalAnchor.Top when inheritedTextBody is not null => ReadTextVerticalAnchor(inheritedTextBody) switch
            {
                TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(textBody, defaultParagraphProperties, theme, textWrapWidth)) / 2d),
                TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(textBody, defaultParagraphProperties, theme, textWrapWidth)),
                _ => 0d
            },
            TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(textBody, defaultParagraphProperties, theme, textWrapWidth)) / 2d),
            TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(textBody, defaultParagraphProperties, theme, textWrapWidth)),
            _ => 0d
        };
        IReadOnlyList<PptxTextParagraphModel> paragraphs = BuildParagraphModels(
            shape,
            textBody,
            inheritedPlaceholders,
            placeholderSources,
            theme,
            slideNumber,
            fontScale,
            lineSpacingScale,
            compatibleLineSpacing,
            shapeFontColor);

        return new PptxTextFrameModel(
            shape,
            textBody,
            inheritedTextBody,
            theme,
            bodyProperties,
            bounds.Value,
            insets,
            fontScale,
            lineSpacingScale,
            textX,
            textWidth,
            textWrapWidth,
            textHeight,
            textClipX,
            textClipWidth,
            textClipY,
            textClipHeight,
            columnCount,
            columnSpacing,
            rotationCenterX,
            rotationCenterY,
            textRotationDegrees,
            textFlipHorizontal,
            textFlipVertical,
            useOfficeBaselineFloor,
            flowYTop,
            verticalOffset,
            orientation,
            shapeFontColor,
            paragraphs);
    }

    private static PptxTextFrameModel BuildTextFrameModel(
        PptxTableCellTextFrame tableFrame,
        PptxDocument document,
        PptxTheme theme,
        int slideNumber,
        IReadOnlyList<XDocument> placeholderSources)
    {
        XElement textBody = tableFrame.TextBody;
        PptxTextBodyProperties baseBodyProperties = ReadTextBodyProperties(textBody, inheritedTextBody: null);
        PptxTextBodyProperties bodyProperties = baseBodyProperties with
        {
            Insets = tableFrame.Insets,
            VerticalAnchor = tableFrame.VerticalAnchor,
            VerticalAnchorValue = tableFrame.VerticalAnchor switch
            {
                TextVerticalAnchor.Middle => "ctr",
                TextVerticalAnchor.Bottom => "b",
                _ => "t"
            },
            VerticalOverflow = PptxTextVerticalOverflow.Clip,
            VerticalOverflowValue = "clip",
            ExplicitWrapWidth = Math.Max(1d, tableFrame.Width - tableFrame.Insets.Left)
        };

        long shapeX = PointsToEmu(tableFrame.X);
        long shapeY = PointsToEmu(document.SlideHeightPoints - tableFrame.Y - tableFrame.Height);
        long shapeWidth = PointsToEmu(tableFrame.Width);
        long shapeHeight = PointsToEmu(tableFrame.Height);
        var bounds = new ShapeBounds(shapeX, shapeY, shapeWidth, shapeHeight, RotationDegrees: 0d, FlipHorizontal: false, FlipVertical: false);
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        TextInsets insets = bodyProperties.Insets;
        PptxTextOrientation orientation = bodyProperties.Orientation;
        double fontScale = bodyProperties.FontScale;
        double lineSpacingScale = bodyProperties.LineSpacingScale;
        bool compatibleLineSpacing = bodyProperties.CompatibleLineSpacing;
        double rotationCenterX = x + width / 2d;
        double rotationCenterY = document.SlideHeightPoints - yTop - height / 2d;
        double flowX = x;
        double flowYTop = yTop;
        double flowWidth = width;
        double flowHeight = height;
        double? explicitTextRotationDegrees = bodyProperties.RotationDegrees;
        double textRotationDegrees = explicitTextRotationDegrees ?? 0d;
        bool textFlipHorizontal = false;
        bool textFlipVertical = false;
        if (orientation is PptxTextOrientation.Vertical or
            PptxTextOrientation.Vertical270 or
            PptxTextOrientation.EastAsianVertical or
            PptxTextOrientation.MongolianVertical or
            PptxTextOrientation.WordArtVertical or
            PptxTextOrientation.WordArtVerticalRightToLeft)
        {
            flowWidth = height;
            flowHeight = width;
            flowX = rotationCenterX - flowWidth / 2d;
            flowYTop = document.SlideHeightPoints - rotationCenterY - flowHeight / 2d;
            textRotationDegrees += TextOrientationRotationDegrees(orientation);
        }

        double textX = flowX + insets.Left;
        double textWidth = Math.Max(1d, flowWidth - insets.Left - insets.Right);
        double textWrapWidth = bodyProperties.ExplicitWrapWidth ?? textWidth;
        double textHeight = Math.Max(1d, flowHeight - insets.Top - insets.Bottom);
        int columnCount = bodyProperties.ColumnCount;
        double columnSpacing = bodyProperties.ColumnSpacing;
        bool clipsVerticalOverflow = bodyProperties.VerticalOverflow == PptxTextVerticalOverflow.Clip;
        bool clipsTextLocally = clipsVerticalOverflow || columnCount != 1;
        double textClipX = clipsTextLocally ? textX : 0d;
        double textClipWidth = clipsTextLocally ? textWidth : document.SlideWidthPoints;
        double textClipY = 0d;
        double textClipHeight = document.SlideHeightPoints;
        if (clipsVerticalOverflow)
        {
            (textClipY, textClipHeight) = IntersectVerticalTextClipWithSlide(
                document.SlideHeightPoints - flowYTop - insets.Top - textHeight,
                textHeight,
                document.SlideHeightPoints);
        }

        XElement? defaultParagraphProperties = MergeParagraphProperties(
            BuildParagraphStyleCascade(textBody, textBody, inheritedPlaceholders: [], placeholderSources, "lvl1pPr").Sources.ToArray());
        double verticalOffset = bodyProperties.VerticalAnchor switch
        {
            TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(textBody, defaultParagraphProperties, theme, textWrapWidth, tableFrame.TextStyle)) / 2d),
            TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(textBody, defaultParagraphProperties, theme, textWrapWidth, tableFrame.TextStyle)),
            _ => 0d
        };
        IReadOnlyList<PptxTextParagraphModel> paragraphs = BuildParagraphModels(
            textBody,
            textBody,
            inheritedPlaceholders: [],
            placeholderSources,
            theme,
            slideNumber,
            fontScale,
            lineSpacingScale,
            compatibleLineSpacing,
            shapeFontColor: null,
            tableFrame.TextStyle);

        return new PptxTextFrameModel(
            textBody,
            textBody,
            InheritedTextBody: null,
            theme,
            bodyProperties,
            bounds,
            insets,
            fontScale,
            lineSpacingScale,
            textX,
            textWidth,
            textWrapWidth,
            textHeight,
            textClipX,
            textClipWidth,
            textClipY,
            textClipHeight,
            columnCount,
            columnSpacing,
            rotationCenterX,
            rotationCenterY,
            textRotationDegrees,
            textFlipHorizontal,
            textFlipVertical,
            UseOfficeBaselineFloor: bodyProperties.VerticalAnchor == TextVerticalAnchor.Top,
            flowYTop,
            verticalOffset,
            orientation,
            ShapeFontColor: null,
            paragraphs);
    }

    private static bool TextFrameUsesOfficeBaselineFloor(XElement shape, PptxTextBodyProperties bodyProperties)
    {
        if (bodyProperties.VerticalAnchor != TextVerticalAnchor.Top)
        {
            return false;
        }

        XElement? geometry = shape
            .Element(PresentationNamespace + "spPr")
            ?.Element(DrawingNamespace + "prstGeom");
        string? preset = (string?)geometry?.Attribute("prst");
        return string.IsNullOrEmpty(preset) || string.Equals(preset, "rect", StringComparison.Ordinal);
    }

    private static (double Y, double Height) IntersectVerticalTextClipWithSlide(double y, double height, double slideHeight)
    {
        double minY = Math.Max(0d, y);
        double maxY = Math.Min(slideHeight, y + Math.Max(0d, height));
        return maxY <= minY
            ? (minY, 0d)
            : (minY, maxY - minY);
    }

    private static PptxTextBodyProperties ReadTextBodyProperties(XElement textBody, XElement? inheritedTextBody)
    {
        (int columnCount, double columnSpacing) = ReadTextColumns(textBody);
        string? orientation = ReadTextOrientationValue(textBody, inheritedTextBody);
        string? verticalAnchor = ReadTextBodyAttribute(textBody, "anchor");
        string? anchorCenter = ReadTextBodyAttribute(textBody, "anchorCtr");
        string? wrap = ReadTextBodyAttribute(textBody, "wrap");
        string? verticalOverflow = ReadTextBodyAttribute(textBody, "vertOverflow");
        XElement? bodyPr = textBody.Element(DrawingNamespace + "bodyPr");
        return new PptxTextBodyProperties(
            ReadTextInsets(textBody),
            ParseTextOrientation(orientation),
            orientation,
            ParseTextVerticalAnchor(verticalAnchor),
            verticalAnchor,
            OoxBoolean.ParseOptionalAttribute(bodyPr, "anchorCtr"),
            anchorCenter,
            ParseTextWrapMode(wrap),
            wrap,
            ParseTextVerticalOverflow(verticalOverflow),
            verticalOverflow,
            columnCount,
            columnSpacing,
            ReadNormAutofitFontScale(textBody),
            ReadNormAutofitLineSpacingScale(textBody),
            HasCompatibleLineSpacing(textBody),
            ReadTextBodyRotationDegrees(textBody),
            ExplicitWrapWidth: null);
    }

    private static TextInsets ReadPresetTextRectInsets(XElement shape, double width, double height)
    {
        if (IsTextBoxShape(shape))
        {
            return TextInsets.Empty;
        }

        string? preset = shape
            .Element(PresentationNamespace + "spPr")
            ?.Element(DrawingNamespace + "prstGeom")
            ?.Attribute("prst")
            ?.Value;

        return preset switch
        {
            "ellipse" => new TextInsets(
                width * PptxTextMetricRules.EllipseTextRectInsetRatio,
                width * PptxTextMetricRules.EllipseTextRectInsetRatio,
                0d,
                0d),
            _ => TextInsets.Empty
        };
    }

    private static bool IsTextBoxShape(XElement shape)
    {
        return string.Equals(
            (string?)shape
                .Element(PresentationNamespace + "nvSpPr")
                ?.Element(PresentationNamespace + "cNvSpPr")
                ?.Attribute("txBox"),
            "1",
            StringComparison.Ordinal);
    }

    private static IReadOnlyList<PptxTextParagraphModel> BuildParagraphModels(
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedPlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        int slideNumber,
        double fontScale,
        double lineSpacingScale,
        bool compatibleLineSpacing,
        RgbColor? shapeFontColor,
        PptxSceneTableCellTextStyle tableStyleTextStyle = default)
    {
        var paragraphs = new List<PptxTextParagraphModel>();
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
            int paragraphLevel = paragraphProperties?.Attribute("lvl") is { } levelAttribute
                ? int.Parse(levelAttribute.Value, CultureInfo.InvariantCulture)
                : 0;
            string levelName = $"lvl{Math.Clamp(paragraphLevel + 1, 1, 9).ToString(CultureInfo.InvariantCulture)}pPr";
            PptxParagraphStyleCascade cascade = BuildParagraphStyleCascade(shape, textBody, inheritedPlaceholders, placeholderSources, levelName);
            XElement? defaultParagraphProperties = MergeParagraphProperties(cascade.Sources.ToArray());
            ResolvedParagraphTextStyle paragraphStyle = ResolveParagraphTextStyle(paragraph, paragraphProperties, defaultParagraphProperties, fontScale, lineSpacingScale, compatibleLineSpacing);
            IReadOnlyList<PptxTextRunModel> runs = BuildRunModels(paragraph, paragraphStyle, shapeFontColor, theme, slideNumber, fontScale, tableStyleTextStyle);
            paragraphs.Add(new PptxTextParagraphModel(
                paragraph,
                paragraphProperties,
                defaultParagraphProperties,
                paragraphLevel,
                cascade,
                paragraphStyle,
                runs));
        }

        return paragraphs;
    }

    private static PptxParagraphStyleCascade BuildParagraphStyleCascade(
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedPlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        string levelName)
    {
        var layers = new List<PptxParagraphStyleLayer>
        {
            new(
                "shape.lstStyle",
                PptxParagraphStyleLayerKind.ShapeListStyle,
                textBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName))
        };
        layers.AddRange(inheritedPlaceholders
            .Select((placeholder, index) => new PptxParagraphStyleLayer(
                $"placeholder[{index}].lstStyle",
                PptxParagraphStyleLayerKind.InheritedPlaceholderListStyle,
                placeholder.Element(PresentationNamespace + "txBody")?.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)))
            .Reverse());
        layers.Add(new PptxParagraphStyleLayer(
            "inherited.txStyle",
            PptxParagraphStyleLayerKind.InheritedTextStyle,
            FindInheritedTextStyle(shape, placeholderSources, levelName)));
        layers.Add(new PptxParagraphStyleLayer(
            "defaultTextStyle",
            PptxParagraphStyleLayerKind.DefaultTextStyle,
            FindDefaultTextStyle(placeholderSources, levelName)));
        return new PptxParagraphStyleCascade(levelName, layers);
    }

    private static IReadOnlyList<PptxTextRunModel> BuildRunModels(
        XElement paragraph,
        ResolvedParagraphTextStyle paragraphStyle,
        RgbColor? shapeFontColor,
        PptxTheme theme,
        int slideNumber,
        double fontScale,
        PptxSceneTableCellTextStyle tableStyleTextStyle)
    {
        var runs = new List<PptxTextRunModel>();
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == DrawingNamespace + "br")
            {
                XElement? breakProperties = child.Element(DrawingNamespace + "rPr");
                PptxRunStyleCascade breakCascade = BuildRunStyleCascade("break.rPr", breakProperties, paragraphStyle.DefaultRunProperties);
                runs.Add(new PptxTextRunModel(
                    PptxTextRunKind.Break,
                    child,
                    breakProperties,
                    breakCascade,
                    "\n",
                    ResolveRunTextStyle(breakProperties, paragraphStyle.DefaultRunProperties, shapeFontColor, theme, fontScale, tableStyleTextStyle)));
                continue;
            }

            if (!IsTextRunElement(child))
            {
                continue;
            }

            XElement? runProperties = child.Element(DrawingNamespace + "rPr");
            PptxRunStyleCascade textRunCascade = BuildRunStyleCascade("run.rPr", runProperties, paragraphStyle.DefaultRunProperties);
            runs.Add(new PptxTextRunModel(
                child.Name == DrawingNamespace + "fld" ? PptxTextRunKind.Field : PptxTextRunKind.Text,
                child,
                runProperties,
                textRunCascade,
                ReadTextElementText(child, slideNumber),
                ResolveRunTextStyle(runProperties, paragraphStyle.DefaultRunProperties, shapeFontColor, theme, fontScale, tableStyleTextStyle)));
        }

        return runs;
    }

    private static PptxRunStyleCascade BuildRunStyleCascade(
        string runPropertiesLayerName,
        XElement? runProperties,
        XElement? paragraphDefaultRunProperties)
    {
        return new PptxRunStyleCascade(
        [
            new PptxRunStyleLayer(runPropertiesLayerName, PptxRunStyleLayerKind.RunProperties, runProperties),
            new PptxRunStyleLayer(
                "paragraph.defRPr",
                PptxRunStyleLayerKind.ParagraphDefaultRunProperties,
                paragraphDefaultRunProperties)
        ]);
    }
}
