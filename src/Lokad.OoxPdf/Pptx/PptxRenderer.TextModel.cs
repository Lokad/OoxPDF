using System.Globalization;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Pptx.PptxRunTextAttributeReaders;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static IReadOnlyList<PptxTextFrameModelSnapshot> InspectTextFrameModels(PptxDocument document, OoxPackage package, int slideIndex, PptxScene? sharedScene = null, PresentationFontResolver? inspectionFonts = null)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null, cancellationToken: CancellationToken.None, sharedScene: sharedScene, inspectionFonts: inspectionFonts);
        if (context is null)
        {
            return [];
        }

        return EnumerateInheritedPaintSources(context)
            .SelectMany(source => (BuildTextFrameModelsAsync(context, source, includePlaceholders: false, placeholderSources: []).GetAwaiter().GetResult()))
            .Concat((BuildTextFrameModelsAsync(context, context.SlideSource, includePlaceholders: true, context.InheritedXml).GetAwaiter().GetResult()))
            .Select(ToSnapshot)
            .ToArray();
    }

    // Inherited XML still supplies placeholder defaults when master paint is
    // suppressed. Only placed inspection content follows the paint selection.
    private static IEnumerable<PptxRenderSource> EnumerateInheritedPaintSources(PptxRenderContext context) =>
        context.InheritedSources.Where(source => source.Kind != PptxRenderSourceKind.Master || context.SceneSlide.ShowMasterShapes);

    internal static IReadOnlyList<PptxTextFrameModelSnapshot> InspectTableTextFrameModels(PptxDocument document, OoxPackage package, int slideIndex, PptxScene? sharedScene = null, PresentationFontResolver? inspectionFonts = null)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null, cancellationToken: CancellationToken.None, sharedScene: sharedScene, inspectionFonts: inspectionFonts);
        if (context is null)
        {
            return [];
        }

        return (ReadSceneTableTextFramesAsync(context).GetAwaiter().GetResult())
            .Select(tableFrame => BuildTableCellTextFrameModelAsync(tableFrame, context).GetAwaiter().GetResult())
            .Select(ToSnapshot)
            .ToArray();
    }

    private static PptxTextFrameModelSnapshot ToSnapshot(PptxTextFrameModel frame)
    {
        return new PptxTextFrameModelSnapshot(
            OoxUnits.EmuToPoints(frame.Bounds.X),
            OoxUnits.EmuToPoints(frame.Bounds.Y),
            OoxUnits.EmuToPoints(frame.Bounds.Width),
            OoxUnits.EmuToPoints(frame.Bounds.Height),
            frame.TextX,
            frame.TextWidth,
            frame.TextWrapWidth,
            frame.TextHeight,
            frame.TableRowIndex,
            frame.TableColumnIndex,
            frame.TableRowSpan,
            frame.TableColumnSpan,
            frame.TableDeclaredRowHeight,
            frame.TableDeclaredRowSpanHeight,
            frame.TableDeclaredHeight,
            frame.TableHeightSlackFactor,
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
            frame.BodyProperties.CompatibleLineSpacingValue,
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
            paragraph.EndParagraphProperties is not null,
            paragraph.EndParagraphStyle.FontSize,
            paragraph.EndParagraphStyle.Typeface,
            paragraph.EndParagraphStyle.Bold,
            paragraph.EndParagraphStyle.Italic,
            paragraph.EmptySpacingBefore,
            paragraph.EmptySpacingAfter,
            paragraph.HasLayoutContent,
            paragraph.HasVisibleContent,
            paragraph.HasManualLineBreak,
            paragraph.FirstLineFallbackFontSize,
            paragraph.Cascade.LevelName,
            paragraph.Cascade.Layers.Count(layer => layer.Source is not null),
            paragraph.Cascade.Layers.Select(layer => layer.Name).ToArray(),
            paragraph.Cascade.Layers.Select(layer => layer.Kind.ToString()).ToArray(),
            paragraph.ResolvedStyleCascade.Layers.Count(layer => layer.Source is not null),
            paragraph.ResolvedStyleCascade.Layers.Select(layer => layer.Name).ToArray(),
            paragraph.ResolvedStyleCascade.Layers.Select(layer => layer.Kind.ToString()).ToArray(),
            paragraph.Style.Alignment.ToString(),
            paragraph.Style.AlignmentValue,
            paragraph.Style.FontSize,
            paragraph.Bullet.Kind.ToString(),
            paragraph.Bullet.Character,
            paragraph.Bullet.ResolvedCharacter,
            paragraph.Bullet.AutoNumberType,
            paragraph.Bullet.AutoNumberStartAtValue,
            paragraph.Bullet.FontTypeface,
            paragraph.Bullet.FontCharset,
            paragraph.Bullet.ResolvedFontTypeface,
            paragraph.Bullet.FontTypefaceSource.ToString(),
            FormatColor(paragraph.Bullet.Paint?.Color),
            paragraph.Bullet.SizeKind.ToString(),
            paragraph.Bullet.SizeValue,
            paragraph.Style.SpacingBefore,
            paragraph.Style.SpacingAfter,
            paragraph.Style.LineSpacing.Value,
            paragraph.Style.LineSpacing.IsAbsolute ? "Absolute" : paragraph.Style.LineSpacing.IsExplicit ? "Multiple" : "Default",
            paragraph.Style.LineSpacing.UseNormalLineAdvance,
            paragraph.Style.Indent.MarginLeft,
            paragraph.Style.Indent.Hanging,
            paragraph.Style.TabStops,
            paragraph.Runs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextRunModelSnapshot ToSnapshot(PptxTextRunModel run)
    {
        return new PptxTextRunModelSnapshot(
            run.RunIndex,
            run.Kind.ToString(),
            run.Text,
            run.Cascade.Sources.Count(source => source is not null),
            run.Cascade.Layers.Select(layer => layer.Name).ToArray(),
            run.Cascade.Layers.Select(layer => layer.Kind.ToString()).ToArray(),
            run.Style.FontSize,
            run.Style.CharacterSpacing,
            run.Style.Typeface,
            run.Style.TypefaceSource.ToString(),
            run.Style.ColorSource.ToString(),
            run.Style.HasHyperlinkClick,
            run.Style.HyperlinkClickId,
            run.Style.Underline,
            run.Style.UnderlineValue,
            run.Style.Strike,
            run.Style.StrikeValue,
            run.Style.CapsValue,
            run.Style.Highlight?.Color);
    }

    private static async ValueTask<IReadOnlyList<PptxTextFrameModel>> BuildTextFrameModelsAsync(
        PptxRenderContext context,
        PptxRenderSource source,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return (await BuildTextFrameModelsAsync(source.Xml, context.Document, context.Theme, source.ColorMap, context.DisplayedSlideNumber, includePlaceholders, placeholderSources, context.FontResolver, context.CancellationToken).ConfigureAwait(false));
    }

    private static async ValueTask<IReadOnlyList<PptxTextFrameModel>> BuildTextFrameModelsAsync(
        XDocument slideXml,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        long slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PresentationFontResolver? fontResolver = null,
        CancellationToken cancellationToken = default)
    {
        var frames = new List<PptxTextFrameModel>();
        foreach (XElement shape in slideXml.Descendants(PresentationNamespace + "sp"))
        {
            if (!PptxSceneBuilder.IsVisibleDrawingContent(shape)) continue;
            PptxTextFrameModel? frame = (await BuildTextFrameModelAsync(shape, document, theme, colorMap, slideNumber, includePlaceholders, placeholderSources, fontResolver, cancellationToken).ConfigureAwait(false));
            if (frame is not null)
            {
                frames.Add(frame);
            }
        }

        return frames;
    }

    private static async ValueTask<PptxTextFrameModel?> BuildTextFrameModelAsync(
        XElement shape,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        long slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PresentationFontResolver? fontResolver = null,
        CancellationToken cancellationToken = default,
        bool prepareSceneFrame = false)
    {
        if (!includePlaceholders && IsPlaceholder(shape))
        {
            return null;
        }

        XElement? textBody = shape.Element(PresentationNamespace + "txBody");
        IReadOnlyList<XElement> inheritedPlaceholders = FindInheritedPlaceholderShapes(shape, placeholderSources);
        XElement? inheritedTextBody = BuildEffectiveInheritedTextBody();
        ShapeBounds? bounds = ReadShapeBounds(shape, inheritedPlaceholders, out bool usesInheritedShapeBounds);
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
        double rotationCenterX = x + width / 2d;
        double rotationCenterY = document.SlideHeightPoints - yTop - height / 2d;
        double orientationRotation = TextOrientationRotationDegrees(orientation);
        var (flowX, flowYTop, flowWidth, flowHeight) = ResolveTextFlowRectangle(x, yTop, width, height,
            document.SlideHeightPoints - rotationCenterY, orientationRotation, ReadPresetTextRectInsets(shape, width, height));

        double textRotationDegrees = (bodyProperties.RotationDegrees ?? DefaultShapeTextRotationDegrees(bounds.Value)) + orientationRotation;

        bool textFlipHorizontal = false;
        bool textFlipVertical = false;

        // Vertical text wraps and centers against the top/bottom insets: the flow column
        // spans the shape height, so Office measures it with the vertical insets (six Office
        // renders agree on tokens and centering). Tables keep Left/Right (unobserved).
        double columnInsetStart = orientation == PptxTextOrientation.Vertical ? insets.Top : insets.Left;
        double columnInsetEnd = orientation == PptxTextOrientation.Vertical ? insets.Bottom : insets.Right;
        double textX = flowX + columnInsetStart;
        double textWidth = Math.Max(1d, flowWidth - columnInsetStart - columnInsetEnd);
        double textWrapWidth = bodyProperties.ExplicitWrapWidth ?? textWidth;
        double textHeight = Math.Max(1d, flowHeight - insets.Top - insets.Bottom);
        int columnCount = bodyProperties.ColumnCount;
        double columnSpacing = bodyProperties.ColumnSpacing;
        bool clipsVerticalOverflow = ClipsTextVerticalOverflow(bodyProperties.VerticalOverflow);
        bool clipsTextLocally = bodyProperties.ClipsTextHorizontally;
        double textClipX = clipsTextLocally ? textX : 0d;
        double textClipWidth = clipsTextLocally ? textWidth : document.SlideWidthPoints;
        double textClipY = 0d;
        double textClipHeight = document.SlideHeightPoints;
        if (clipsVerticalOverflow)
        {
            // F01: the vertical clip spans the shape text area (bounds minus regular
            // insets), not the preset-inscribed flow rect. Office keeps a glyph whose
            // baseline sits inside shape bounds even when preset inscription shrinks the
            // flow rect past it (small-label-origin: Office Y=359.11 inside shape bounds,
            // outside the ellipse-inscribed clip). Horizontal placement and clip keep the
            // inscribed rect, which Office matches to 0.03pt. Shapes without preset
            // insets are unaffected (both rects coincide).
            if (ShapeWritingModeClipRotation(bounds.Value, orientation) != 0d)
            {
                // Cull baselines in the writing frame. Its clip is projected
                // to the slide only after placement.
                textClipHeight = textHeight;
                textClipY = document.SlideHeightPoints - flowYTop - insets.Top - textClipHeight;
            }
            else
            {
                double clipTextHeight = Math.Max(1d, height - insets.Top - insets.Bottom);
                (textClipY, textClipHeight) = IntersectVerticalTextClipWithSlide(
                    document.SlideHeightPoints - yTop - insets.Top - clipTextHeight,
                    clipTextHeight,
                    document.SlideHeightPoints);
            }
        }
        var shapeFontPaint = ReadShapeFontPaint(shape, theme, colorMap);
        bool useOfficeBaselineFloor = true;
        IReadOnlyList<PptxTextParagraphModel> paragraphs = prepareSceneFrame && !NeedsShapeTextMeasurement(bodyProperties)
            ? []
            : BuildParagraphModels(shape, textBody, inheritedPlaceholders, placeholderSources, theme, colorMap,
                slideNumber, bodyProperties, shapeFontPaint, default,
                disableZeroKerning: UsesNoWrapZeroKerning(bodyProperties, textRotationDegrees),
                defaultTextStyle: document.DefaultTextStyle);
        fontResolver ??= new PresentationFontResolver(await WindowsFontResolver.CreateAsync(cancellationToken).ConfigureAwait(false), cancellationToken);
        // Vertical middle/bottom anchors resolve from laid-out actuals below (the estimate
        // uses the wrong axis and advance); other orientations keep estimated offsets.
        TextVerticalAnchor anchorForEstimate = orientation == PptxTextOrientation.Vertical ? TextVerticalAnchor.Top : bodyProperties.VerticalAnchor;
        double verticalOffset = anchorForEstimate switch
        {
            TextVerticalAnchor.Middle => Math.Max(0d, (textHeight - (await EstimateTextHeightAsync(paragraphs, textWrapWidth, bodyProperties, fontResolver, cancellationToken).ConfigureAwait(false))) / 2d),
            TextVerticalAnchor.Bottom => Math.Max(0d, textHeight - (await EstimateTextHeightAsync(paragraphs, textWrapWidth, bodyProperties, fontResolver, cancellationToken).ConfigureAwait(false))),
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
            TableRowIndex: null,
            TableColumnIndex: null,
            TableRowSpan: null,
            TableColumnSpan: null,
            TableDeclaredRowHeight: null,
            TableDeclaredRowSpanHeight: null,
            TableDeclaredHeight: null,
            TableHeightSlackFactor: null,
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
            shapeFontPaint,
            paragraphs);

        XElement? BuildEffectiveInheritedTextBody()
        {
            XElement[] inheritedTextBodies = inheritedPlaceholders
                .Select(placeholder => placeholder.Element(PresentationNamespace + "txBody"))
                .Where(candidateTextBody => candidateTextBody is not null)
                .Cast<XElement>()
                .ToArray();
            if (inheritedTextBodies.Length == 0)
            {
                return null;
            }

            XElement effectiveTextBody = new(PresentationNamespace + "txBody");
            XElement effectiveBodyProperties = new(DrawingNamespace + "bodyPr");
            foreach (XElement mergedBodyProperties in inheritedTextBodies
                .Select(candidateTextBody => candidateTextBody.Element(DrawingNamespace + "bodyPr"))
                .Where(mergedBodyProperties => mergedBodyProperties is not null)
                .Cast<XElement>())
            {
                foreach (XAttribute attribute in mergedBodyProperties.Attributes())
                {
                    effectiveBodyProperties.SetAttributeValue(attribute.Name, attribute.Value);
                }

                foreach (XElement child in mergedBodyProperties.Elements())
                {
                    effectiveBodyProperties.Elements(child.Name).Remove();
                    effectiveBodyProperties.Add(new XElement(child));
                }
            }

            effectiveTextBody.Add(effectiveBodyProperties);
            return effectiveTextBody;
        }
    }

    private static async ValueTask<PptxTextFrameModel> BuildTableCellTextFrameModelAsync(
        PptxTableCellTextFrame tableFrame,
        PptxRenderContext context,
        bool estimateVerticalOffset = true,
        TableCellTextMeasurement? measurement = null)
    {
        PptxDocument document = context.Document;
        XElement textBody = tableFrame.TextBody;
        PptxTextBodyProperties baseBodyProperties = ReadTextBodyProperties(textBody, inheritedTextBody: null);
        PptxTextBodyProperties bodyProperties = baseBodyProperties with
        {
            Insets = tableFrame.Insets,
            InsetSources = tableFrame.InsetSources,
            InsetValues = tableFrame.InsetValues,
            VerticalAnchor = tableFrame.VerticalAnchor,
            VerticalAnchorValue = tableFrame.VerticalAnchorValue,
            VerticalAnchorSource = tableFrame.VerticalAnchorSource,
            VerticalOverflow = PptxTextVerticalOverflow.Clip,
            VerticalOverflowValue = "clip",
            VerticalOverflowSource = PptxTextBodyPropertySource.TableCellStyle,
            ClipsHorizontalOverflow = true,
            ExplicitWrapWidth = Math.Max(1d, tableFrame.Width - tableFrame.Insets.Left - tableFrame.Insets.Right)
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
        double rotationCenterX = x + width / 2d;
        double rotationCenterY = document.SlideHeightPoints - yTop - height / 2d;
        double orientationRotation = TextOrientationRotationDegrees(orientation);
        var (flowX, flowYTop, flowWidth, flowHeight) = ResolveTextFlowRectangle(x, yTop, width, height,
            document.SlideHeightPoints - rotationCenterY, orientationRotation);
        double textRotationDegrees = (bodyProperties.RotationDegrees ?? 0d) + orientationRotation;
        bool textFlipHorizontal = false;
        bool textFlipVertical = false;

        double textX = flowX + insets.Left;
        double textWidth = Math.Max(1d, flowWidth - insets.Left - insets.Right);
        double textWrapWidth = bodyProperties.ExplicitWrapWidth ?? textWidth;
        double textHeight = Math.Max(1d, flowHeight - insets.Top - insets.Bottom);
        int columnCount = bodyProperties.ColumnCount;
        double columnSpacing = bodyProperties.ColumnSpacing;
        bool clipsVerticalOverflow = ClipsTextVerticalOverflow(bodyProperties.VerticalOverflow);
        bool clipsTextLocally = bodyProperties.ClipsTextHorizontally;
        double textClipX = clipsTextLocally ? flowX : 0d;
        double textClipWidth = clipsTextLocally ? flowWidth : document.SlideWidthPoints;
        double textClipY = 0d;
        double textClipHeight = document.SlideHeightPoints;
        if (clipsVerticalOverflow)
        {
            (textClipY, textClipHeight) = IntersectVerticalTextClipWithSlide(
                document.SlideHeightPoints - flowYTop - flowHeight,
                flowHeight,
                document.SlideHeightPoints);
        }

        IReadOnlyList<PptxTextParagraphModel> paragraphs = measurement?.Paragraphs ?? BuildParagraphModels(
            textBody,
            textBody,
            inheritedPlaceholders: [],
            context.InheritedXml,
            context.Theme,
            tableFrame.ColorMap,
            context.DisplayedSlideNumber,
            bodyProperties,
            shapeFontPaint: null,
            tableFrame.TextStyle,
            defaultTextStyle: document.DefaultTextStyle);
        double verticalOffset = 0d;
        if (estimateVerticalOffset && bodyProperties.VerticalAnchor is TextVerticalAnchor.Middle or TextVerticalAnchor.Bottom)
        {
            double contentHeight = measurement?.Height ??
                await EstimateTextHeightAsync(paragraphs, textWrapWidth, bodyProperties, context.FontResolver, context.CancellationToken).ConfigureAwait(false);
            verticalOffset = Math.Max(0d, textHeight - contentHeight);
            if (bodyProperties.VerticalAnchor == TextVerticalAnchor.Middle)
            {
                verticalOffset /= 2d;
            }
        }

        return new PptxTextFrameModel(
            textBody,
            textBody,
            InheritedTextBody: null,
            InheritedPlaceholderCount: 0,
            UsesInheritedShapeBounds: false,
            context.Theme,
            bodyProperties,
            bounds,
            tableFrame.RowIndex,
            tableFrame.ColumnIndex,
            tableFrame.RowSpan,
            tableFrame.ColumnSpan,
            tableFrame.DeclaredRowHeight,
            tableFrame.DeclaredRowSpanHeight,
            tableFrame.DeclaredTableHeight,
            tableFrame.TableHeightSlackFactor,
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
            ShapeFontPaint: null,
            paragraphs);
    }

    private static async ValueTask PrepareParagraphFontsAsync(IReadOnlyList<PptxTextParagraphModel> paragraphs,
        TextAdvanceEstimator estimator, CancellationToken cancellationToken)
    {
        foreach (PptxTextParagraphModel paragraph in paragraphs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!paragraph.HasVisibleContent)
            {
                if (paragraph.HasLayoutContent)
                    await estimator.PrepareAsync(paragraph.EndParagraphStyle.Typeface, paragraph.EndParagraphStyle.Bold,
                        paragraph.EndParagraphStyle.Italic, string.Empty).ConfigureAwait(false);
                continue;
            }
            foreach (PptxTextRunModel run in paragraph.Runs)
            {
                // Use the same caps transform as flow. Tabs occupy spaces; break
                // markers select metrics but never require a newline glyph.
                if (run.Kind == PptxTextRunKind.Break)
                    await estimator.PrepareAsync(run.Style.Typeface, run.Style.Bold, run.Style.Italic, string.Empty).ConfigureAwait(false);
                else
                    foreach (TextCapsFragment fragment in ApplyTextCaps(run.Text.Replace('\t', ' '), run.Properties, paragraph.Style.DefaultRunProperties))
                        await estimator.PrepareAsync(run.Style.Typeface, run.Style.Bold, run.Style.Italic, fragment.Text).ConfigureAwait(false);
            }
        }
    }

    private static bool UsesNoWrapZeroKerning(PptxTextBodyProperties bodyProperties, double textRotationDegrees)
    {
        // Office's zero threshold disables pair kerning in these controls.
        // Wrapped and aligned frames still need their existing advance policy:
        // changing it can alter line breaks or worsen alignment.
        return bodyProperties.WrapMode == PptxTextWrapMode.None &&
            bodyProperties.AutofitMode == PptxTextAutofitMode.None &&
            bodyProperties.Orientation == PptxTextOrientation.Horizontal &&
            bodyProperties.ColumnCount == 1 &&
            Math.Abs(textRotationDegrees) < PptxTextMetricRules.CoordinateTolerance;
    }

    private static (double X, double YTop, double Width, double Height) ResolveTextFlowRectangle(
        double x, double yTop, double width, double height, double centerYFromTop, double orientationRotation,
        TextInsets presetInsets = default)
    {
        if (orientationRotation != 0d)
        {
            return (x + width / 2d - height / 2d, centerYFromTop - width / 2d, height, width);
        }

        return presetInsets.IsEmpty
            ? (x, yTop, width, height)
            : (x + presetInsets.Left, yTop + presetInsets.Top,
                Math.Max(1d, width - presetInsets.Left - presetInsets.Right),
                Math.Max(1d, height - presetInsets.Top - presetInsets.Bottom));
    }

    private static (double Y, double Height) IntersectVerticalTextClipWithSlide(double y, double height, double slideHeight)
    {
        double minY = Math.Max(0d, y);
        double maxY = Math.Min(slideHeight, y + Math.Max(0d, height));
        return maxY <= minY
            ? (minY, 0d)
            : (minY, maxY - minY);
    }

    private static IReadOnlyList<PptxTextParagraphModel> BuildParagraphModels(
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedPlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        PptxColorMap colorMap,
        long slideNumber,
        PptxTextBodyProperties bodyProperties,
        (RgbColor Color, double Alpha)? shapeFontPaint,
        PptxSceneTableCellTextStyle tableStyleTextStyle,
        bool disableZeroKerning = false,
        XElement? defaultTextStyle = null)
    {
        double compatibleDefaultLineSpacingFactor = ResolveCompatibleDefaultLineSpacingFactor(bodyProperties);
        var paragraphs = new List<PptxTextParagraphModel>();
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            XElement? paragraphProperties = paragraph.Element(DrawingNamespace + "pPr");
            int paragraphLevel = PptxTextStyleInheritance.ReadParagraphLevel(paragraphProperties);
            string levelName = PptxTextStyleInheritance.ParagraphLevelName(paragraphLevel);
            PptxParagraphStyleCascade cascade = BuildParagraphStyleCascade(levelName);
            paragraphs.Add(BuildParagraphModel(paragraph, paragraphProperties, paragraph.Element(DrawingNamespace + "endParaRPr"),
                paragraphLevel, cascade, theme, colorMap, slideNumber, bodyProperties, bodyProperties.FontScale,
                compatibleDefaultLineSpacingFactor, shapeFontPaint, tableStyleTextStyle, disableZeroKerning));
        }

        return paragraphs;


        PptxParagraphStyleCascade BuildParagraphStyleCascade(string levelName)
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
                    InheritedPlaceholderLayerName(placeholder, index, placeholderSources.Count),
                    InheritedPlaceholderLayerKind(placeholder, index, placeholderSources.Count),
                    placeholder.Element(PresentationNamespace + "txBody")?.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)))
                .Reverse());
            layers.Add(new PptxParagraphStyleLayer(
                "inherited.txStyle",
                PptxParagraphStyleLayerKind.InheritedTextStyle,
                PptxTextStyleInheritance.FindInheritedTextStyle(shape, placeholderSources, levelName)));
            layers.Add(new PptxParagraphStyleLayer(
                "defaultTextStyle",
                PptxParagraphStyleLayerKind.DefaultTextStyle,
                PptxTextStyleInheritance.FindDefaultTextStyle(defaultTextStyle, levelName)));
            return new PptxParagraphStyleCascade(levelName, layers);

            PptxParagraphStyleLayerKind InheritedPlaceholderLayerKind(XElement placeholder, int sourceIndex, int sourceCount)
            {
                return PptxTextStyleInheritance.PlaceholderListStyleLayerKindName(placeholder, sourceIndex, sourceCount) switch
                {
                    "MasterPlaceholderListStyle" => PptxParagraphStyleLayerKind.MasterPlaceholderListStyle,
                    "LayoutPlaceholderListStyle" => PptxParagraphStyleLayerKind.LayoutPlaceholderListStyle,
                    _ => PptxParagraphStyleLayerKind.InheritedPlaceholderListStyle,
                };
            }

            string InheritedPlaceholderLayerName(XElement placeholder, int sourceIndex, int sourceCount)
            {
                return PptxTextStyleInheritance.PlaceholderListStyleLayerName(placeholder, sourceIndex, sourceCount);
            }
        }


    }

    // The scene retains the paragraph cascade, including its placeholder and master sources.
    private static IReadOnlyList<PptxTextParagraphModel> BuildSceneFedParagraphModels(
        IReadOnlyList<PptxSceneTextParagraph> sceneParagraphs,
        PptxTextFrameModel frameModel,
        PptxTheme theme,
        PptxColorMap colorMap,
        long slideNumber)
    {
        double compatibleDefaultLineSpacingFactor = ResolveCompatibleDefaultLineSpacingFactor(frameModel.BodyProperties);
        bool disableZeroKerning = UsesNoWrapZeroKerning(frameModel.BodyProperties, frameModel.TextRotationDegrees);
        var paragraphs = new List<PptxTextParagraphModel>();
        foreach (PptxSceneTextParagraph sceneParagraph in sceneParagraphs)
        {
            if (sceneParagraph.Source is null)
            {
                throw new InvalidOperationException("Expected a retained paragraph element for scene-fed layout.");
            }

            XElement paragraph = sceneParagraph.Source;
            XElement? paragraphProperties = sceneParagraph.Properties;
            string levelName = PptxTextStyleInheritance.ParagraphLevelName(sceneParagraph.Level);
            var cascade = new PptxParagraphStyleCascade(
                levelName,
                sceneParagraph.CascadeLayers.Select(layer => new PptxParagraphStyleLayer(layer.Name, SceneCascadeLayerKind(layer.Kind), layer.Source)).ToList());
            paragraphs.Add(BuildParagraphModel(paragraph, paragraphProperties, sceneParagraph.EndParagraphProperties,
                sceneParagraph.Level, cascade, theme, colorMap, slideNumber, frameModel.BodyProperties, frameModel.FontScale,
                compatibleDefaultLineSpacingFactor, frameModel.ShapeFontPaint, default, disableZeroKerning));
        }

        return paragraphs;
    }

    private static PptxTextParagraphModel BuildParagraphModel(
        XElement paragraph,
        XElement? paragraphProperties,
        XElement? endParagraphProperties,
        int paragraphLevel,
        PptxParagraphStyleCascade cascade,
        PptxTheme theme,
        PptxColorMap colorMap,
        long slideNumber,
        PptxTextBodyProperties bodyProperties,
        double fontScale,
        double compatibleDefaultLineSpacingFactor,
        (RgbColor Color, double Alpha)? shapeFontPaint,
        PptxSceneTableCellTextStyle tableStyleTextStyle,
        bool disableZeroKerning)
    {
        XElement? defaultParagraphProperties = cascade.ResolveDefaultProperties();
        PptxParagraphStyleCascade resolvedStyleCascade = BuildResolvedParagraphStyleCascade(cascade, paragraphProperties);
        XElement? effectiveParagraphProperties = paragraphProperties is null ? defaultParagraphProperties : resolvedStyleCascade.ResolveDefaultProperties();
        ResolvedParagraphTextStyle paragraphStyle = ResolveParagraphTextStyle(paragraph, paragraphProperties,
            effectiveParagraphProperties, fontScale, bodyProperties.LineSpacingScale, bodyProperties.CompatibleLineSpacing,
            compatibleDefaultLineSpacingFactor);
        PptxParagraphBulletModel bullet = BuildParagraphBulletModel(effectiveParagraphProperties, theme, colorMap);
        IReadOnlyList<PptxTextRunModel> runs = BuildRunModels(paragraph, paragraphStyle, resolvedStyleCascade,
            shapeFontPaint, theme, colorMap, fontScale, slideNumber, tableStyleTextStyle,
            disableZeroKerning && paragraphStyle.Alignment == TextAlignment.Left && bullet.Kind == PptxParagraphBulletKind.None);
        ResolvedEndParagraphTextStyle endParagraphStyle = ResolveEndParagraphTextStyle(endParagraphProperties, paragraphStyle.DefaultRunProperties, fontScale);
        return new PptxTextParagraphModel(
            paragraph,
            paragraphProperties,
            endParagraphProperties,
            endParagraphStyle,
            ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", endParagraphStyle.FontSize),
            ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", endParagraphStyle.FontSize),
            paragraphProperties is not null || endParagraphProperties is not null,
            runs.Count > 0,
            runs.Any(run => run.Kind == PptxTextRunKind.Break || TextContainsManualLineBreak(run.Text)),
            paragraphStyle.FontSize,
            paragraphLevel,
            cascade,
            resolvedStyleCascade,
            paragraphStyle,
            bullet,
            runs);
    }

    private static IReadOnlyList<PptxTextRunModel> BuildRunModels(
        XElement paragraph,
        ResolvedParagraphTextStyle paragraphStyle,
        PptxParagraphStyleCascade resolvedParagraphStyleCascade,
        (RgbColor Color, double Alpha)? shapeFontPaint,
        PptxTheme theme,
        PptxColorMap colorMap,
        double fontScale,
        long slideNumber,
        PptxSceneTableCellTextStyle tableStyleTextStyle,
        bool disableZeroKerning = false)
    {
        var runs = new List<PptxTextRunModel>();
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == DrawingNamespace + "br")
            {
                XElement? breakProperties = child.Element(DrawingNamespace + "rPr");
                PptxRunStyleCascade breakCascade = BuildRunStyleCascade("break.rPr", breakProperties, resolvedParagraphStyleCascade, paragraphStyle.DefaultRunProperties);
                runs.Add(new PptxTextRunModel(
                    runs.Count,
                    PptxTextRunKind.Break,
                    child,
                    breakProperties,
                    breakCascade,
                    "\n",
                    ResolveRunTextStyle(breakCascade, shapeFontPaint, theme, colorMap, fontScale, tableStyleTextStyle, disableZeroKerning)));
                continue;
            }

            if (!IsTextRunElement(child))
            {
                continue;
            }

            XElement? runProperties = child.Element(DrawingNamespace + "rPr");
            PptxRunStyleCascade textRunCascade = BuildRunStyleCascade("run.rPr", runProperties, resolvedParagraphStyleCascade, paragraphStyle.DefaultRunProperties);
            runs.Add(new PptxTextRunModel(
                runs.Count,
                child.Name == DrawingNamespace + "fld" ? PptxTextRunKind.Field : PptxTextRunKind.Text,
                child,
                runProperties,
                textRunCascade,
                ReadTextElementText(child, slideNumber),
                ResolveRunTextStyle(textRunCascade, shapeFontPaint, theme, colorMap, fontScale, tableStyleTextStyle, disableZeroKerning)));
        }

        return runs;
    }

    private static PptxParagraphStyleLayerKind SceneCascadeLayerKind(string kind)
    {
        return kind switch
        {
            "ShapeListStyle" => PptxParagraphStyleLayerKind.ShapeListStyle,
            "MasterPlaceholderListStyle" => PptxParagraphStyleLayerKind.MasterPlaceholderListStyle,
            "LayoutPlaceholderListStyle" => PptxParagraphStyleLayerKind.LayoutPlaceholderListStyle,
            "InheritedPlaceholderListStyle" => PptxParagraphStyleLayerKind.InheritedPlaceholderListStyle,
            "InheritedTextStyle" => PptxParagraphStyleLayerKind.InheritedTextStyle,
            "DefaultTextStyle" => PptxParagraphStyleLayerKind.DefaultTextStyle,
            _ => throw new InvalidOperationException("Unknown scene cascade layer kind."),
        };
    }

    internal static async ValueTask<IReadOnlyList<PptxPositionedTextSpan>> BuildSceneFedTextSpansAsync(
        PptxSceneNode node,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        long slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PresentationFontResolver? fontResolver,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        PptxSceneTextBody? textBody = node.TextBody;
        if (textBody is null || node.Source.Name != PresentationNamespace + "sp" ||
            !PptxSceneBuilder.IsVisibleDrawingContent(node.Source))
        {
            return [];
        }

        fontResolver ??= new PresentationFontResolver(await WindowsFontResolver.CreateAsync(cancellationToken).ConfigureAwait(false), cancellationToken);
        PptxTextFrameModel? frameModel = await BuildTextFrameModelAsync(node.Source, document, theme, colorMap,
            slideNumber, includePlaceholders, placeholderSources, fontResolver, cancellationToken,
            prepareSceneFrame: true).ConfigureAwait(false);
        if (frameModel is null)
        {
            return [];
        }

        var estimator = new TextAdvanceEstimator(fontResolver, cancellationToken);
        if (NeedsShapeTextMeasurement(frameModel.BodyProperties))
        {
            frameModel = (await BuildTextFrameLayoutAsync(frameModel, document, estimator).ConfigureAwait(false)).Model;
        }

        IReadOnlyList<PptxTextParagraphModel> fedParagraphs = BuildSceneFedParagraphModels(textBody.Paragraphs, frameModel, theme, colorMap, slideNumber);
        var fedFrame = frameModel with { Paragraphs = fedParagraphs };
        PptxTextFlowFrame fedFlow = await BuildTextFlowFrameAsync(fedFrame, document, estimator).ConfigureAwait(false);
        PptxTextFrameLayout fedLayout = (await BuildTextFrameLayoutAsync(fedFlow, document, estimator, allowWrapping: true).ConfigureAwait(false));
        return RemapMongolianVerticalSpans(await FlattenTextLayoutToSpansAsync(new PptxTextLayoutModel(new[] { fedLayout }), fontResolver, document).ConfigureAwait(false), node.Source);
    }

    // Autofit and center/bottom anchors can change the frame before scene text is placed.
    private static bool NeedsShapeTextMeasurement(PptxTextBodyProperties properties) =>
        HasShapeAutoFit(properties) || properties.VerticalAnchor is TextVerticalAnchor.Middle or TextVerticalAnchor.Bottom;

    private static PptxParagraphBulletModel BuildParagraphBulletModel(XElement? paragraphProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        if (paragraphProperties is null || paragraphProperties.Element(DrawingNamespace + "buNone") is not null)
        {
            return new PptxParagraphBulletModel(PptxParagraphBulletKind.None, null, null, null, null, null, null, null, null, PptxThemeTypefaceSource.DefaultMinorLatin, null, PptxParagraphBulletSizeKind.Text, null);
        }

        XElement? bulletFont = FindBulletProperty("buFont");
        XElement? bulletColor = FindBulletProperty("buClr");
        XElement? bulletSizePercent = FindBulletProperty("buSzPct");
        XElement? bulletSizePoints = FindBulletProperty("buSzPts");
        string? fontTypeface = (string?)bulletFont?.Attribute("typeface");
        string? fontCharset = (string?)bulletFont?.Attribute("charset");
        PptxThemeTypefaceResolution FontFaceResolution = FindBulletProperty("buFontTx") is not null
            ? new(null, null, PptxThemeTypefaceSource.FollowText)
            : theme.ResolveTypefaceWithSource(fontTypeface);
        (RgbColor Color, double Alpha)? paint = bulletColor is not null &&
            PptxColorResolver.TryReadSolidColorWithAlpha(bulletColor, theme, colorMap, out RgbColor resolvedColor, out double alpha)
            ? (resolvedColor, alpha)
            : null;
        PptxParagraphBulletSizeKind sizeKind = PptxParagraphBulletSizeKind.Text;
        string? sizeValue = null;
        if ((string?)bulletSizePercent?.Attribute("val") is { } percentValue)
        {
            sizeKind = PptxParagraphBulletSizeKind.Percent;
            sizeValue = percentValue;
        }
        else if ((string?)bulletSizePoints?.Attribute("val") is { } pointValue)
        {
            sizeKind = PptxParagraphBulletSizeKind.Points;
            sizeValue = pointValue;
        }

        if ((string?)paragraphProperties.Element(DrawingNamespace + "buChar")?.Attribute("char") is { } character)
        {
            string resolvedCharacter = IsSymbolBulletFont(bulletFont)
                ? MapSymbolBulletText(character)
                : character;
            return new PptxParagraphBulletModel(PptxParagraphBulletKind.Character, character, resolvedCharacter, null, null, null, fontTypeface, fontCharset, FontFaceResolution.Typeface, FontFaceResolution.Source, paint, sizeKind, sizeValue);
        }

        if (paragraphProperties.Element(DrawingNamespace + "buAutoNum") is { } autoNumber)
        {
            string? startAtValue = (string?)autoNumber.Attribute("startAt");
            int? startAt = int.TryParse(startAtValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedStartAt) && parsedStartAt > 0
                ? parsedStartAt
                : null;
            return new PptxParagraphBulletModel(
                PptxParagraphBulletKind.AutoNumber,
                null,
                null,
                (string?)autoNumber.Attribute("type"),
                startAtValue,
                startAt,
                fontTypeface,
                fontCharset,
                FontFaceResolution.Typeface,
                FontFaceResolution.Source,
                paint,
                sizeKind,
                sizeValue);
        }

        if (paragraphProperties.Element(DrawingNamespace + "buBlip") is not null)
        {
            return new PptxParagraphBulletModel(PptxParagraphBulletKind.Blip, null, null, null, null, null, fontTypeface, fontCharset, FontFaceResolution.Typeface, FontFaceResolution.Source, paint, sizeKind, sizeValue);
        }

        return new PptxParagraphBulletModel(PptxParagraphBulletKind.None, null, null, null, null, null, fontTypeface, fontCharset, FontFaceResolution.Typeface, FontFaceResolution.Source, paint, sizeKind, sizeValue);

        XElement? FindBulletProperty(string localName)
        {
            if (paragraphProperties is null)
            {
                return null;
            }

            XName propertyName = DrawingNamespace + localName;
            XElement? marker = paragraphProperties
                .Elements()
                .FirstOrDefault(element => element.Name == DrawingNamespace + "buChar" ||
                    element.Name == DrawingNamespace + "buAutoNum" ||
                    element.Name == DrawingNamespace + "buBlip");
            IEnumerable<XElement> candidates = marker is null
                ? paragraphProperties.Elements()
                : paragraphProperties.Elements().TakeWhile(element => element != marker);
            return candidates.FirstOrDefault(element => element.Name == propertyName);
        }
    }

    private static PptxParagraphStyleCascade BuildResolvedParagraphStyleCascade(PptxParagraphStyleCascade defaultCascade, XElement? paragraphProperties)
    {
        var layers = new List<PptxParagraphStyleLayer>
        {
            new(
                "paragraph.pPr",
                PptxParagraphStyleLayerKind.ParagraphProperties,
                paragraphProperties)
        };
        layers.AddRange(defaultCascade.Layers);
        return new PptxParagraphStyleCascade(defaultCascade.LevelName, layers);
    }

    private static bool TextContainsManualLineBreak(string text)
    {
        return text.Contains('\n', StringComparison.Ordinal) || text.Contains('\r', StringComparison.Ordinal);
    }

    private static double ResolveCompatibleDefaultLineSpacingFactor(PptxTextBodyProperties bodyProperties)
    {
        return HasShapeAutoFit(bodyProperties)
            ? PptxTextMetricRules.OfficeCompatibleDefaultLineSpacingFactor
            : PptxTextMetricRules.OfficeCompatibleNoAutoFitDefaultLineSpacingFactor;
    }

    private static string? FormatColor(RgbColor? color)
    {
        return color is { } value
            ? string.Create(CultureInfo.InvariantCulture, $"{value.Red:X2}{value.Green:X2}{value.Blue:X2}")
            : null;
    }

    private static ResolvedEndParagraphTextStyle ResolveEndParagraphTextStyle(XElement? endRunProperties, XElement? defaultRunProperties, double fontScale)
    {
        double fontSize = ReadFontSize(endRunProperties, defaultRunProperties) * fontScale;
        string? typeface = ReadTypeface(endRunProperties) ?? ReadTypeface(defaultRunProperties);
        bool bold = OoxXml.ParseOptionalBool(endRunProperties, "b") ||
            (endRunProperties?.Attribute("b") is null && OoxXml.ParseOptionalBool(defaultRunProperties, "b"));
        bool italic = OoxXml.ParseOptionalBool(endRunProperties, "i") ||
            (endRunProperties?.Attribute("i") is null && OoxXml.ParseOptionalBool(defaultRunProperties, "i"));
        return new ResolvedEndParagraphTextStyle(fontSize, typeface, bold, italic);
    }

    private static PptxRunStyleCascade BuildRunStyleCascade(
        string runPropertiesLayerName,
        XElement? runProperties,
        PptxParagraphStyleCascade resolvedParagraphStyleCascade,
        XElement? paragraphDefaultRunProperties)
    {
        var layers = new List<PptxRunStyleLayer>
        {
            new(runPropertiesLayerName, PptxRunStyleLayerKind.RunProperties, runProperties),
            new(
                "paragraph.defRPr",
                PptxRunStyleLayerKind.ParagraphDefaultRunProperties,
                paragraphDefaultRunProperties)
        };
        layers.AddRange(resolvedParagraphStyleCascade.Layers
            .Select(layer => new PptxRunStyleLayer(
                $"{layer.Name}.defRPr",
                RunDefaultLayerKind(layer.Kind),
                layer.Source?.Element(DrawingNamespace + "defRPr"))));
        return new PptxRunStyleCascade(layers, paragraphDefaultRunProperties);
    }

    private static PptxRunStyleLayerKind RunDefaultLayerKind(PptxParagraphStyleLayerKind paragraphLayerKind)
    {
        return paragraphLayerKind switch
        {
            PptxParagraphStyleLayerKind.ShapeListStyle => PptxRunStyleLayerKind.ShapeListStyleDefaultRunProperties,
            PptxParagraphStyleLayerKind.MasterPlaceholderListStyle => PptxRunStyleLayerKind.MasterPlaceholderDefaultRunProperties,
            PptxParagraphStyleLayerKind.LayoutPlaceholderListStyle => PptxRunStyleLayerKind.LayoutPlaceholderDefaultRunProperties,
            PptxParagraphStyleLayerKind.InheritedPlaceholderListStyle => PptxRunStyleLayerKind.InheritedPlaceholderDefaultRunProperties,
            PptxParagraphStyleLayerKind.InheritedTextStyle => PptxRunStyleLayerKind.InheritedTextStyleDefaultRunProperties,
            PptxParagraphStyleLayerKind.DefaultTextStyle => PptxRunStyleLayerKind.DefaultTextStyleDefaultRunProperties,
            PptxParagraphStyleLayerKind.ParagraphProperties => PptxRunStyleLayerKind.ParagraphPropertiesDefaultRunProperties,
            _ => PptxRunStyleLayerKind.ParagraphDefaultRunProperties
        };
    }
}
