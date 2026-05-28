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
            frame.TextHeight,
            frame.VerticalOffset,
            frame.Insets.Left,
            frame.Insets.Right,
            frame.Insets.Top,
            frame.Insets.Bottom,
            frame.BodyProperties.InsetValues.Left,
            frame.BodyProperties.InsetValues.Right,
            frame.BodyProperties.InsetValues.Top,
            frame.BodyProperties.InsetValues.Bottom,
            frame.FontScale,
            frame.BodyProperties.FontScaleValue,
            frame.BodyProperties.FontScaleSource.ToString(),
            frame.BodyProperties.LineSpacingScale,
            frame.BodyProperties.LineSpacingReductionValue,
            frame.BodyProperties.LineSpacingScaleSource.ToString(),
            frame.BodyProperties.CompatibleLineSpacing,
            frame.BodyProperties.CompatibleLineSpacingSource.ToString(),
            frame.BodyProperties.RotationDegrees,
            frame.BodyProperties.RotationValue,
            frame.BodyProperties.RotationDegreesSource.ToString(),
            frame.InheritedPlaceholderCount,
            frame.InheritedTextBody is not null,
            frame.UsesInheritedShapeBounds,
            frame.BodyProperties.InsetSources.Left.ToString(),
            frame.BodyProperties.InsetSources.Right.ToString(),
            frame.BodyProperties.InsetSources.Top.ToString(),
            frame.BodyProperties.InsetSources.Bottom.ToString(),
            frame.BodyProperties.Orientation.ToString(),
            frame.BodyProperties.OrientationValue,
            frame.BodyProperties.OrientationSource.ToString(),
            frame.BodyProperties.VerticalAnchor.ToString(),
            frame.BodyProperties.VerticalAnchorValue,
            frame.BodyProperties.VerticalAnchorSource.ToString(),
            frame.BodyProperties.AnchorCenter,
            frame.BodyProperties.AnchorCenterValue,
            frame.BodyProperties.AnchorCenterSource.ToString(),
            frame.BodyProperties.WrapMode.ToString(),
            frame.BodyProperties.WrapValue,
            frame.BodyProperties.WrapSource.ToString(),
            frame.BodyProperties.VerticalOverflow.ToString(),
            frame.BodyProperties.VerticalOverflowValue,
            frame.BodyProperties.VerticalOverflowSource.ToString(),
            frame.BodyProperties.ColumnCount,
            frame.BodyProperties.ColumnSpacing,
            frame.BodyProperties.ColumnSource.ToString(),
            frame.BodyProperties.ColumnCountSource.ToString(),
            frame.BodyProperties.ColumnSpacingSource.ToString(),
            frame.BodyProperties.ColumnCountValue,
            frame.BodyProperties.ColumnSpacingValue,
            frame.BodyProperties.AutofitModeValue,
            frame.BodyProperties.AutofitModeSource.ToString(),
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
            paragraph.ResolvedStyleCascade.Sources.Count(source => source is not null),
            paragraph.ResolvedStyleCascade.Layers.Select(layer => layer.Name).ToArray(),
            paragraph.ResolvedStyleCascade.Layers.Select(layer => layer.Kind.ToString()).ToArray(),
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
            run.Style.ColorSource.ToString(),
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
        bool usesInheritedShapeBounds = bounds is null && inheritedPlaceholder?.Element(PresentationNamespace + "spPr") is not null;
        if (bounds is null && inheritedPlaceholder?.Element(PresentationNamespace + "spPr") is { } inheritedProperties)
        {
            bounds = ReadBounds(inheritedProperties);
        }
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
        double verticalOffset = bodyProperties.VerticalAnchor switch
        {
            TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(paragraphs, textWrapWidth, bodyProperties)) / 2d),
            TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(paragraphs, textWrapWidth, bodyProperties)),
            _ => 0d
        };

        return new PptxTextFrameModel(
            shape,
            textBody,
            inheritedTextBody,
            inheritedPlaceholders.Count,
            usesInheritedShapeBounds,
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
            InsetSources = new TextInsetSources(
                PptxTextBodyPropertySource.TableCellStyle,
                PptxTextBodyPropertySource.TableCellStyle,
                PptxTextBodyPropertySource.TableCellStyle,
                PptxTextBodyPropertySource.TableCellStyle),
            VerticalAnchor = tableFrame.VerticalAnchor,
            VerticalAnchorValue = tableFrame.VerticalAnchor switch
            {
                TextVerticalAnchor.Middle => "ctr",
                TextVerticalAnchor.Bottom => "b",
                _ => "t"
            },
            VerticalAnchorSource = PptxTextBodyPropertySource.TableCellStyle,
            VerticalOverflow = PptxTextVerticalOverflow.Clip,
            VerticalOverflowValue = "clip",
            VerticalOverflowSource = PptxTextBodyPropertySource.TableCellStyle,
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
        double verticalOffset = bodyProperties.VerticalAnchor switch
        {
            TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - EstimateTextHeight(paragraphs, textWrapWidth, bodyProperties)) / 2d),
            TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - EstimateTextHeight(paragraphs, textWrapWidth, bodyProperties)),
            _ => 0d
        };

        return new PptxTextFrameModel(
            textBody,
            textBody,
            InheritedTextBody: null,
            InheritedPlaceholderCount: 0,
            UsesInheritedShapeBounds: false,
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
        (int columnCount, double columnSpacing, PptxTextBodyPropertySource columnCountSource, PptxTextBodyPropertySource columnSpacingSource, string? columnCountValue, string? columnSpacingValue) =
            ReadTextColumns(textBody, inheritedTextBody);
        PptxTextBodyPropertySource columnSource = MergeTextBodyPropertySources(columnCountSource, columnSpacingSource);
        (TextInsets insets, TextInsetSources insetSources, TextInsetValues insetValues) = ReadTextInsets(textBody, inheritedTextBody);
        (XElement? autofit, string autofitMode, PptxTextBodyPropertySource autofitModeSource) = ReadTextAutofit(textBody, inheritedTextBody);
        (double fontScale, PptxTextBodyPropertySource fontScaleSource, string? fontScaleValue) = ReadNormAutofitFontScale(autofit, autofitModeSource);
        (double lineSpacingScale, PptxTextBodyPropertySource lineSpacingScaleSource, string? lineSpacingReductionValue) = ReadNormAutofitLineSpacingScale(autofit, autofitModeSource);
        (string? orientation, PptxTextBodyPropertySource orientationSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "vert", inherit: true);
        (string? verticalAnchor, PptxTextBodyPropertySource verticalAnchorSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "anchor", inherit: true);
        (string? anchorCenter, PptxTextBodyPropertySource anchorCenterSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "anchorCtr", inherit: true);
        (string? wrap, PptxTextBodyPropertySource wrapSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "wrap", inherit: true);
        (string? verticalOverflow, PptxTextBodyPropertySource verticalOverflowSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "vertOverflow", inherit: true);
        (string? compatibleLineSpacing, PptxTextBodyPropertySource compatibleLineSpacingSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "compatLnSpc", inherit: true);
        (string? rotation, PptxTextBodyPropertySource rotationSource) = ReadTextBodyAttributeWithSource(textBody, inheritedTextBody, "rot", inherit: true);
        return new PptxTextBodyProperties(
            insets,
            insetSources,
            insetValues,
            ParseTextOrientation(orientation),
            orientation,
            orientationSource,
            ParseTextVerticalAnchor(verticalAnchor),
            verticalAnchor,
            verticalAnchorSource,
            anchorCenter is null ? null : OoxBoolean.IsTrue(anchorCenter),
            anchorCenter,
            anchorCenterSource,
            ParseTextWrapMode(wrap),
            wrap,
            wrapSource,
            ParseTextVerticalOverflow(verticalOverflow),
            verticalOverflow,
            verticalOverflowSource,
            columnCount,
            columnSpacing,
            columnSource,
            columnCountSource,
            columnSpacingSource,
            columnCountValue,
            columnSpacingValue,
            autofitMode,
            autofitModeSource,
            fontScale,
            fontScaleValue,
            fontScaleSource,
            lineSpacingScale,
            lineSpacingReductionValue,
            lineSpacingScaleSource,
            compatibleLineSpacing is not null && OoxBoolean.IsTrue(compatibleLineSpacing),
            compatibleLineSpacingSource,
            ParseTextBodyRotationDegrees(rotation),
            rotation,
            rotationSource,
            ExplicitWrapWidth: null);
    }

    private static (string? Value, PptxTextBodyPropertySource Source) ReadTextBodyAttributeWithSource(
        XElement textBody,
        XElement? inheritedTextBody,
        string attributeName,
        bool inherit)
    {
        XElement? bodyPr = textBody.Element(DrawingNamespace + "bodyPr");
        if (bodyPr?.Attribute(attributeName) is { } directAttribute)
        {
            return (directAttribute.Value, PptxTextBodyPropertySource.DirectBodyPr);
        }

        XElement? inheritedBodyPr = inheritedTextBody?.Element(DrawingNamespace + "bodyPr");
        if (inherit && inheritedBodyPr?.Attribute(attributeName) is { } inheritedAttribute)
        {
            return (inheritedAttribute.Value, PptxTextBodyPropertySource.InheritedBodyPr);
        }

        return (null, PptxTextBodyPropertySource.DefaultValue);
    }

    private static PptxTextBodyPropertySource MergeTextBodyPropertySources(
        PptxTextBodyPropertySource first,
        PptxTextBodyPropertySource second)
    {
        if (first == second)
        {
            return first;
        }

        if (first == PptxTextBodyPropertySource.DirectBodyPr || second == PptxTextBodyPropertySource.DirectBodyPr)
        {
            return PptxTextBodyPropertySource.DirectBodyPr;
        }

        if (first == PptxTextBodyPropertySource.InheritedBodyPr || second == PptxTextBodyPropertySource.InheritedBodyPr)
        {
            return PptxTextBodyPropertySource.InheritedBodyPr;
        }

        if (first == PptxTextBodyPropertySource.TableCellStyle || second == PptxTextBodyPropertySource.TableCellStyle)
        {
            return PptxTextBodyPropertySource.TableCellStyle;
        }

        return PptxTextBodyPropertySource.DefaultValue;
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
            PptxParagraphStyleCascade resolvedStyleCascade = BuildResolvedParagraphStyleCascade(cascade, paragraphProperties);
            IReadOnlyList<PptxTextRunModel> runs = BuildRunModels(paragraph, paragraphStyle, shapeFontColor, theme, slideNumber, fontScale, tableStyleTextStyle);
            paragraphs.Add(new PptxTextParagraphModel(
                paragraph,
                paragraphProperties,
                defaultParagraphProperties,
                paragraphLevel,
                cascade,
                resolvedStyleCascade,
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
                InheritedPlaceholderLayerName(index, placeholderSources.Count),
                InheritedPlaceholderLayerKind(index, placeholderSources.Count),
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

    private static string InheritedPlaceholderLayerName(int sourceIndex, int sourceCount)
    {
        return (sourceCount, sourceIndex) switch
        {
            (2, 0) => "master.placeholder.lstStyle",
            (2, _) => "layout.placeholder.lstStyle",
            _ => $"inherited.placeholder[{sourceIndex}].lstStyle"
        };
    }

    private static PptxParagraphStyleLayerKind InheritedPlaceholderLayerKind(int sourceIndex, int sourceCount)
    {
        return (sourceCount, sourceIndex) switch
        {
            (2, 0) => PptxParagraphStyleLayerKind.MasterPlaceholderListStyle,
            (2, _) => PptxParagraphStyleLayerKind.LayoutPlaceholderListStyle,
            _ => PptxParagraphStyleLayerKind.InheritedPlaceholderListStyle
        };
    }

    private static PptxParagraphStyleCascade BuildResolvedParagraphStyleCascade(
        PptxParagraphStyleCascade defaultCascade,
        XElement? paragraphProperties)
    {
        var layers = new List<PptxParagraphStyleLayer>(defaultCascade.Layers)
        {
            new(
                "paragraph.pPr",
                PptxParagraphStyleLayerKind.ParagraphProperties,
                paragraphProperties)
        };
        return new PptxParagraphStyleCascade(defaultCascade.LevelName, layers);
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
