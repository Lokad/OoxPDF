using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal static IReadOnlyList<PptxTextRunSnapshot> InspectTextRuns(PptxDocument document, OoxPackage package, int slideIndex)
    {
        return ReadSlideTextRunsForInspection(document, package, slideIndex)
            .Select(ToSnapshot)
            .ToArray();
    }

    private static IReadOnlyList<TextRun> ReadSlideTextRunsForInspection(PptxDocument document, OoxPackage package, int slideIndex)
    {
        return ReadSlideTextSpansForInspection(document, package, slideIndex).Select(span => span.Run).ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadSlideTextSpansForInspection(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null);
        if (context is null)
        {
            return [];
        }

        return context.InheritedSources
            .SelectMany(source => ReadTextSpans(context, source, includePlaceholders: false, placeholderSources: []))
            .Concat(ReadTextSpans(context, context.SlideSource, includePlaceholders: true, context.InheritedXml))
            .Concat(ReadSceneTableTextSpans(context))
            .ToArray();
    }

    internal static PptxTextLayoutSnapshot InspectTextLayout(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null);
        if (context is null)
        {
            return new PptxTextLayoutSnapshot([]);
        }

        PptxTextLayoutModel inheritedLayout = BuildTextLayoutModelForSources(context.InheritedSources, context);
        PptxTextLayoutModel slideLayout = BuildTextLayoutModel(context, context.SlideSource, includePlaceholders: true, context.InheritedXml);
        return ToSnapshot(new PptxTextLayoutModel(inheritedLayout.Frames.Concat(slideLayout.Frames).ToArray()));
    }

    internal static PptxTextFlowSnapshot InspectTextFlow(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null);
        if (context is null)
        {
            return new PptxTextFlowSnapshot([]);
        }

        PptxTextFlowModel inheritedFlow = BuildTextFlowModelForSources(context.InheritedSources, context);
        PptxTextFlowModel slideFlow = BuildTextFlowModel(context, context.SlideSource, includePlaceholders: true, context.InheritedXml);
        return ToSnapshot(new PptxTextFlowModel(inheritedFlow.Frames.Concat(slideFlow.Frames).ToArray()));
    }

    private static PptxTextLayoutModel BuildTextLayoutModelForSources(
        IReadOnlyList<PptxRenderSource> sources,
        PptxRenderContext context)
    {
        var frames = new List<PptxTextFrameLayout>();
        foreach (PptxRenderSource source in sources)
        {
            frames.AddRange(BuildTextLayoutModel(context, source, includePlaceholders: false, placeholderSources: []).Frames);
        }

        return new PptxTextLayoutModel(frames);
    }

    private static PptxTextFlowModel BuildTextFlowModelForSources(
        IReadOnlyList<PptxRenderSource> sources,
        PptxRenderContext context)
    {
        var frames = new List<PptxTextFlowFrame>();
        foreach (PptxRenderSource source in sources)
        {
            frames.AddRange(BuildTextFlowModel(context, source, includePlaceholders: false, placeholderSources: []).Frames);
        }

        return new PptxTextFlowModel(frames);
    }

    private static PptxTextRunSnapshot ToSnapshot(TextRun run)
    {
        return new PptxTextRunSnapshot(
            run.Text,
            run.X,
            run.Y,
            run.Width,
            run.FontSize,
            run.CharacterSpacing,
            run.Color,
            run.Alpha,
            run.HighlightColor,
            run.Bold,
            run.Italic,
            run.Underline,
            run.Strike,
            run.Alignment.ToString(),
            run.FontFamily);
    }

    private static PptxTextLayoutSnapshot ToSnapshot(PptxTextLayoutModel layout)
    {
        return new PptxTextLayoutSnapshot(layout.Frames.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowSnapshot ToSnapshot(PptxTextFlowModel flow)
    {
        return new PptxTextFlowSnapshot(flow.Frames.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowFrameSnapshot ToSnapshot(PptxTextFlowFrame frame)
    {
        return new PptxTextFlowFrameSnapshot(
            frame.Box.TextX,
            frame.Box.TextWidth,
            frame.Box.TextHeight,
            frame.Box.ClipY,
            frame.Box.ClipHeight,
            frame.Box.CursorTop,
            frame.Paragraphs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowParagraphSnapshot ToSnapshot(PptxTextFlowParagraph paragraph)
    {
        return new PptxTextFlowParagraphSnapshot(
            paragraph.Model.Level,
            paragraph.Style.Alignment.ToString(),
            paragraph.Style.FontSize,
            paragraph.Runs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowRunSnapshot ToSnapshot(PptxTextFlowRun run)
    {
        return new PptxTextFlowRunSnapshot(
            run.Source.Kind.ToString(),
            run.Source.Text,
            run.Style.FontSize,
            run.Style.Typeface,
            run.Style.HasHyperlinkClick,
            run.Style.HyperlinkClickId,
            run.Segments.Select(ToSnapshot).ToArray());
    }

    private static PptxTextFlowSegmentSnapshot ToSnapshot(PptxTextFlowSegment segment)
    {
        return new PptxTextFlowSegmentSnapshot(
            segment.Kind.ToString(),
            segment.Text,
            segment.AdvanceText,
            segment.Draw,
            segment.PreventCoalesce,
            segment.FontScale);
    }

    private static PptxTextFrameLayoutSnapshot ToSnapshot(PptxTextFrameLayout frame)
    {
        return new PptxTextFrameLayoutSnapshot(frame.Paragraphs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextParagraphLayoutSnapshot ToSnapshot(PptxTextParagraphLayout paragraph)
    {
        return new PptxTextParagraphLayoutSnapshot(paragraph.Model.Level, paragraph.Lines.Select(ToSnapshot).ToArray());
    }

    private static PptxTextLineLayoutSnapshot ToSnapshot(PptxTextLineLayout line)
    {
        return new PptxTextLineLayoutSnapshot(
            line.Box.TopY,
            line.Box.BaselineY,
            line.Box.Advance,
            line.Box.BaselineOffset,
            line.Box.MaxFontSize,
            line.Box.LineSpacing.IsAbsolute ? "Absolute" : line.Box.LineSpacing.IsExplicit ? "Multiple" : "Default",
            ToSnapshot(line.Box.BaselineMetric),
            line.StartX,
            line.EndX,
            line.NaturalEndX,
            line.Alignment.ToString(),
            line.Spans.Select(ToSnapshot).ToArray());
    }

    private static PptxTextBaselineMetricSnapshot ToSnapshot(PptxTextBaselineMetricLayout metric)
    {
        return new PptxTextBaselineMetricSnapshot(
            metric.Source,
            metric.Typeface,
            metric.Bold,
            metric.Italic,
            metric.FontSize,
            metric.Ratio,
            metric.UnitsPerEm,
            metric.WindowsAscender,
            metric.WindowsDescender,
            metric.TypographicAscender,
            metric.TypographicDescender,
            metric.TypographicLineGap);
    }

    private static PptxTextSpanLayoutSnapshot ToSnapshot(PptxTextSpanLayout span)
    {
        return new PptxTextSpanLayoutSnapshot(
            span.SourceRun?.Text,
            span.Run.Text,
            span.Run.X,
            span.Run.Y,
            span.Run.Width,
            span.Run.FontSize,
            span.Atoms.Select(ToSnapshot).ToArray(),
            ToSnapshot(span.GlyphSpan));
    }

    private static PptxTextAtomLayoutSnapshot ToSnapshot(PptxTextAtomLayout atom)
    {
        return new PptxTextAtomLayoutSnapshot(atom.Kind.ToString(), atom.Text, atom.X, atom.Width, atom.Draw);
    }

    private static PptxTextGlyphSpanLayoutSnapshot ToSnapshot(PptxTextGlyphSpanLayout span)
    {
        return new PptxTextGlyphSpanLayoutSnapshot(
            span.Text,
            span.Typeface,
            span.FontSize,
            span.LeadingAdjustment,
            span.NaturalWidth,
            span.LayoutWidth,
            span.Glyphs.Count,
            span.Glyphs.Skip(1).FirstOrDefault()?.AdjustmentBefore ?? 0d,
            span.Glyphs.Select(ToSnapshot).ToArray());
    }

    private static PptxTextGlyphLayoutSnapshot ToSnapshot(PptxTextGlyphLayout glyph)
    {
        return new PptxTextGlyphLayoutSnapshot(
            glyph.CodePoint,
            glyph.Typeface,
            glyph.TypefaceResolutionSource.ToString(),
            glyph.GlyphId,
            glyph.Advance,
            glyph.AdjustmentBefore);
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadSceneShapeTextSpans(PptxRenderContext context)
    {
        var textSpans = new List<PptxPositionedTextSpan>();
        AddSceneShapeTextSpans(context.SceneSlide.MasterNodes, context, textSpans, context.MasterColorMap, renderPlaceholders: false);
        AddSceneShapeTextSpans(context.SceneSlide.LayoutNodes, context, textSpans, context.LayoutColorMap, renderPlaceholders: false);
        AddSceneShapeTextSpans(context.SceneSlide.SlideNodes, context, textSpans, context.SlideColorMap, renderPlaceholders: true);
        return textSpans;
    }

    private static void AddSceneShapeTextSpans(
        IReadOnlyList<PptxSceneNode> nodes,
        PptxRenderContext context,
        List<PptxPositionedTextSpan> textSpans,
        PptxColorMap colorMap,
        bool renderPlaceholders)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Kind == PptxSceneNodeKind.Shape)
            {
                if (renderPlaceholders || !node.IsPlaceholder)
                {
                    textSpans.AddRange(ReadTextSpansForSceneNode(node, context, colorMap, renderPlaceholders));
                }

                continue;
            }

            if (node.Kind == PptxSceneNodeKind.Group)
            {
                AddSceneShapeTextSpans(node.Children, context, textSpans, colorMap, renderPlaceholders);
            }
        }
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpans(
        PptxRenderContext context,
        PptxRenderSource source,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return FlattenTextLayoutToSpans(BuildTextLayoutModel(context, source, includePlaceholders, placeholderSources), context.FontResolver);
    }

    private static PptxTextLayoutModel BuildTextLayoutModel(
        PptxRenderContext context,
        PptxRenderSource source,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return BuildTextLayoutModel(source.Xml, context.Document, context.Theme, source.ColorMap, context.SlideNumber, includePlaceholders, placeholderSources, context.FontResolver);
    }

    private static PptxTextLayoutModel BuildTextLayoutModel(
        XDocument slideXml,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PresentationFontResolver? fontResolver = null)
    {
        var advanceEstimator = new TextAdvanceEstimator(fontResolver);
        var frames = new List<PptxTextFrameLayout>();
        IReadOnlyList<PptxTextFrameModel> frameModels = BuildTextFrameModels(slideXml, document, theme, colorMap, slideNumber, includePlaceholders, placeholderSources);
        foreach (PptxTextFrameModel frameModel in frameModels)
        {
            frames.Add(BuildTextFrameLayout(frameModel, document, advanceEstimator));
        }

        return new PptxTextLayoutModel(frames);
    }

    private static PptxTextFrameLayout BuildTextFrameLayout(PptxTextFrameModel frameModel, PptxDocument document, TextAdvanceEstimator advanceEstimator)
    {
        frameModel = ResetEstimatedVerticalAnchorOffset(frameModel);
        PptxTextFlowFrame flowFrame = BuildTextFlowFrame(frameModel, document, advanceEstimator);
        PptxTextFrameLayout layout = BuildTextFrameLayout(flowFrame, document, advanceEstimator);
        if (HasShapeAutoFit(frameModel.BodyProperties) && UsesRotatedFrameAutoFit(frameModel.Orientation))
        {
            PptxTextFrameLayout unwrappedLayout = BuildTextFrameLayout(flowFrame, document, advanceEstimator, allowWrapping: false);
            if (TextLayoutOverflows(unwrappedLayout, flowFrame.Box))
            {
                PptxTextFrameModel fitted = FitShapeAutoFitFrame(frameModel, document, advanceEstimator, allowWrapping: false);
                return ApplyActualVerticalAnchorOffsetIfNeeded(
                    BuildTextFrameLayout(BuildTextFlowFrame(fitted, document, advanceEstimator), document, advanceEstimator, allowWrapping: false),
                    document,
                    advanceEstimator,
                    allowWrapping: false);
            }

            return ApplyActualVerticalAnchorOffsetIfNeeded(unwrappedLayout, document, advanceEstimator, allowWrapping: false);
        }

        if (HasShapeAutoFit(frameModel.BodyProperties) &&
            frameModel.Orientation == PptxTextOrientation.Horizontal &&
            TextLayoutOverflowsHorizontally(layout, flowFrame.Box, PptxTextMetricRules.ShapeAutoFitWrapTolerance(ResolveLayoutMaxFontSize(layout), flowFrame.Box.TextWidth)))
        {
            PptxTextFrameModel fitted = FitShapeAutoFitFrame(frameModel, document, advanceEstimator, allowWrapping: true);
            return ApplyActualVerticalAnchorOffsetIfNeeded(
                BuildTextFrameLayout(BuildTextFlowFrame(fitted, document, advanceEstimator), document, advanceEstimator),
                document,
                advanceEstimator,
                allowWrapping: true);
        }

        return ApplyActualVerticalAnchorOffsetIfNeeded(layout, document, advanceEstimator, allowWrapping: true);
    }

    private static PptxTextFrameModel ResetEstimatedVerticalAnchorOffset(PptxTextFrameModel frame)
    {
        if (frame.VerticalOffset <= PptxTextMetricRules.CoordinateTolerance ||
            frame.Orientation != PptxTextOrientation.Horizontal ||
            frame.ColumnCount != 1 ||
            !UsesActualLineBoxVerticalAnchor(frame) ||
            IsTableCellVerticalAnchorSource(frame.BodyProperties.VerticalAnchorSource) ||
            frame.BodyProperties.VerticalAnchor is not (TextVerticalAnchor.Middle or TextVerticalAnchor.Bottom))
        {
            return frame;
        }

        return frame with { VerticalOffset = 0d };
    }

    private static bool UsesActualLineBoxVerticalAnchor(PptxTextFrameModel frame)
    {
        if (HasNoAutoFit(frame.BodyProperties))
        {
            return true;
        }

        return HasShapeAutoFit(frame.BodyProperties) &&
            frame.BodyProperties.CompatibleLineSpacing &&
            frame.Paragraphs.Any(paragraph => paragraph.HasManualLineBreak && HasExplicitParagraphSpacing(paragraph.Properties));
    }

    private static PptxTextFlowModel BuildTextFlowModel(
        PptxRenderContext context,
        PptxRenderSource source,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return BuildTextFlowModel(source.Xml, context.Document, context.Theme, source.ColorMap, context.SlideNumber, includePlaceholders, placeholderSources);
    }

    private static PptxTextFlowModel BuildTextFlowModel(
        XDocument slideXml,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return BuildTextFlowModel(BuildTextFrameModels(slideXml, document, theme, colorMap, slideNumber, includePlaceholders, placeholderSources), document);
    }

    private static PptxTextFlowModel BuildTextFlowModel(IReadOnlyList<PptxTextFrameModel> frames, PptxDocument document)
    {
        var advanceEstimator = new TextAdvanceEstimator();
        return new PptxTextFlowModel(frames.Select(frame => BuildTextFlowFrame(frame, document, advanceEstimator)).ToArray());
    }

    private static PptxTextFlowFrame BuildTextFlowFrame(PptxTextFrameModel frame, PptxDocument document, TextAdvanceEstimator advanceEstimator)
    {
        var box = new PptxTextFlowBox(
            frame.FlowYTop,
            document.SlideHeightPoints - frame.FlowYTop - frame.Insets.Top - frame.VerticalOffset,
            frame.TextX,
            frame.TextWidth,
            frame.TextWrapWidth,
            frame.TextHeight,
            frame.TextClipX,
            frame.TextClipWidth,
            frame.TextClipY,
            frame.TextClipHeight,
            frame.RotationCenterX,
            frame.RotationCenterY);
        bool attachSpacesToFollowingWord = HasNoAutoFit(frame.BodyProperties);
        return new PptxTextFlowFrame(frame, box, frame.Paragraphs.Select(paragraph => BuildTextFlowParagraph(paragraph, attachSpacesToFollowingWord, advanceEstimator)).ToArray());
    }

    private static bool HasShapeAutoFit(PptxTextBodyProperties bodyProperties)
    {
        return bodyProperties.AutofitModeValue == "spAutoFit";
    }

    private static bool HasNoAutoFit(PptxTextBodyProperties bodyProperties)
    {
        return bodyProperties.AutofitModeValue == "noAutofit";
    }

    private static (XElement? Element, string Mode, PptxTextBodyPropertySource Source) ReadTextAutofit(
        XElement textBody,
        XElement? inheritedTextBody)
    {
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        if (TryReadTextAutofit(bodyProperties) is { } directAutofit)
        {
            return (directAutofit.Element, directAutofit.Mode, PptxTextBodyPropertySource.DirectBodyPr);
        }

        XElement? inheritedBodyProperties = inheritedTextBody?.Element(DrawingNamespace + "bodyPr");
        if (TryReadTextAutofit(inheritedBodyProperties) is { } inheritedAutofit)
        {
            return (inheritedAutofit.Element, inheritedAutofit.Mode, PptxTextBodyPropertySource.InheritedBodyPr);
        }

        return (null, string.Empty, PptxTextBodyPropertySource.DefaultValue);
    }

    private static (XElement Element, string Mode)? TryReadTextAutofit(XElement? bodyProperties)
    {
        if (bodyProperties?.Element(DrawingNamespace + "spAutoFit") is { } shapeAutofit)
        {
            return (shapeAutofit, "spAutoFit");
        }

        if (bodyProperties?.Element(DrawingNamespace + "noAutofit") is { } noAutofit)
        {
            return (noAutofit, "noAutofit");
        }

        if (bodyProperties?.Element(DrawingNamespace + "normAutofit") is { } normalAutofit)
        {
            return (normalAutofit, "normAutofit");
        }

        return null;
    }

    private static bool TextBodyAllowsWrapping(PptxTextBodyProperties bodyProperties)
    {
        return bodyProperties.WrapMode != PptxTextWrapMode.None;
    }

    private static bool IsCenteredTableCellText(PptxTextFrameModel frame, ResolvedParagraphTextStyle paragraphStyle)
    {
        return frame.TableRowIndex.HasValue &&
            paragraphStyle.Alignment == TextAlignment.Center &&
            TextBodyAllowsWrapping(frame.BodyProperties);
    }

    private static PptxTextFrameModel FitShapeAutoFitFrame(
        PptxTextFrameModel frame,
        PptxDocument document,
        TextAdvanceEstimator advanceEstimator,
        bool allowWrapping)
    {
        double high = frame.FontScale;
        double low = PptxTextMetricRules.MinimumAutofitScale;
        PptxTextFrameModel best = ScaleTextFrameModel(frame, low);
        for (int i = 0; i < PptxTextMetricRules.ShapeAutoFitSearchIterations; i++)
        {
            double candidateScale = (low + high) / 2d;
            PptxTextFrameModel candidate = ScaleTextFrameModel(frame, candidateScale);
            PptxTextFlowFrame candidateFlow = BuildTextFlowFrame(candidate, document, advanceEstimator);
            PptxTextFrameLayout candidateLayout = BuildTextFrameLayout(candidateFlow, document, advanceEstimator, allowWrapping);
            if (TextLayoutOverflows(candidateLayout, candidateFlow.Box))
            {
                high = candidateScale;
            }
            else
            {
                low = candidateScale;
                best = candidate;
            }
        }

        return best;
    }

    private static bool TextLayoutOverflows(PptxTextFrameLayout layout, PptxTextFlowBox box)
    {
        double bottom = box.CursorTop - box.TextHeight;
        double right = box.TextX + box.TextWidth;
        return layout.Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .Any(line =>
                line.Box.TopY - line.Box.Advance < bottom - PptxTextMetricRules.TextStateTolerance ||
                line.EndX > right + PptxTextMetricRules.TextStateTolerance);
    }

    private static bool TextLayoutOverflowsHorizontally(PptxTextFrameLayout layout, PptxTextFlowBox box)
    {
        return TextLayoutOverflowsHorizontally(layout, box, PptxTextMetricRules.TextStateTolerance);
    }

    private static bool TextLayoutOverflowsHorizontally(PptxTextFrameLayout layout, PptxTextFlowBox box, double tolerance)
    {
        double right = box.TextX + box.TextWidth;
        return layout.Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .Any(line => line.EndX > right + Math.Max(PptxTextMetricRules.TextStateTolerance, tolerance));
    }

    private static PptxTextFrameLayout ApplyActualVerticalAnchorOffsetIfNeeded(
        PptxTextFrameLayout layout,
        PptxDocument document,
        TextAdvanceEstimator advanceEstimator,
        bool allowWrapping)
    {
        if (!TryResolveActualVerticalAnchorOffset(layout, out double verticalOffset) ||
            Math.Abs(verticalOffset - layout.Model.VerticalOffset) <= PptxTextMetricRules.CoordinateTolerance)
        {
            return layout;
        }

        PptxTextFrameModel anchored = layout.Model with { VerticalOffset = verticalOffset };
        return BuildTextFrameLayout(BuildTextFlowFrame(anchored, document, advanceEstimator), document, advanceEstimator, allowWrapping);
    }

    private static bool TryResolveActualVerticalAnchorOffset(PptxTextFrameLayout layout, out double verticalOffset)
    {
        verticalOffset = 0d;
        PptxTextFrameModel frame = layout.Model;
        if (frame.Orientation != PptxTextOrientation.Horizontal ||
            frame.ColumnCount != 1 ||
            frame.VerticalOffset > PptxTextMetricRules.CoordinateTolerance ||
            IsTableCellVerticalAnchorSource(frame.BodyProperties.VerticalAnchorSource))
        {
            return false;
        }

        double multiplier = frame.BodyProperties.VerticalAnchor switch
        {
            TextVerticalAnchor.Middle => PptxTextMetricRules.MiddleVerticalAnchorSlackMultiplier,
            TextVerticalAnchor.Bottom => 1d,
            _ => 0d
        };
        if (multiplier <= 0d)
        {
            return false;
        }

        PptxTextLineLayout[] lines = layout.Paragraphs.SelectMany(paragraph => paragraph.Lines).ToArray();
        if (lines.Length == 0)
        {
            return false;
        }

        double top = lines.Max(line => line.Box.TopY);
        double bottom = lines.Min(line => line.Box.TopY - line.Box.Advance);
        double occupiedHeight = Math.Max(0d, top - bottom);
        double slack = frame.TextHeight - occupiedHeight;
        if (Math.Abs(slack) <= PptxTextMetricRules.CoordinateTolerance)
        {
            return false;
        }

        verticalOffset = slack * multiplier;
        return true;
    }

    private static double ResolveLayoutMaxFontSize(PptxTextFrameLayout layout)
    {
        return layout.Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .Select(span => span.Run.FontSize)
            .DefaultIfEmpty(18d)
            .Max();
    }

    private static PptxTextFrameModel ScaleTextFrameModel(PptxTextFrameModel frame, double fontScale)
    {
        if (Math.Abs(frame.FontScale - fontScale) <= PptxTextMetricRules.TextStateTolerance)
        {
            return frame;
        }

        double ratio = frame.FontScale <= 0d ? fontScale : fontScale / frame.FontScale;
        return frame with
        {
            FontScale = fontScale,
            Paragraphs = frame.Paragraphs.Select(paragraph => ScaleTextParagraphModel(paragraph, ratio)).ToArray()
        };
    }

    private static PptxTextParagraphModel ScaleTextParagraphModel(PptxTextParagraphModel paragraph, double ratio)
    {
        return paragraph with
        {
            FirstLineFallbackFontSize = paragraph.FirstLineFallbackFontSize * ratio,
            Style = ScaleParagraphStyle(paragraph.Style, ratio),
            Runs = paragraph.Runs.Select(run => run with { Style = ScaleRunStyle(run.Style, ratio) }).ToArray()
        };
    }

    private static ResolvedParagraphTextStyle ScaleParagraphStyle(ResolvedParagraphTextStyle style, double ratio)
    {
        return style with
        {
            FontSize = style.FontSize * ratio,
            SpacingBefore = style.SpacingBefore * ratio,
            SpacingAfter = style.SpacingAfter * ratio,
            LineSpacing = style.LineSpacing.ScaleExplicit(ratio)
        };
    }

    private static ResolvedRunTextStyle ScaleRunStyle(ResolvedRunTextStyle style, double ratio)
    {
        return style with
        {
            NominalFontSize = style.NominalFontSize * ratio,
            FontSize = style.FontSize * ratio,
            CharacterSpacing = style.CharacterSpacing * ratio,
            BaselineOffset = style.BaselineOffset * ratio
        };
    }

    private static PptxTextFlowParagraph BuildTextFlowParagraph(PptxTextParagraphModel paragraph, bool attachSpacesToFollowingWord, TextAdvanceEstimator advanceEstimator)
    {
        var runs = new List<PptxTextFlowRun>(paragraph.Runs.Count);
        PptxTextFlowRun? previousDrawableRun = null;
        bool hideLeadingSpacesAfterBoundary = false;
        foreach (PptxTextRunModel run in paragraph.Runs)
        {
            PptxTextFlowRun flowRun = BuildTextFlowRun(run, paragraph.Style.DefaultRunProperties, attachSpacesToFollowingWord);
            bool hideLeadingSpacesAfterStyleBoundary =
                previousDrawableRun is not null &&
                StartsWithDrawableRegularSpace(flowRun.Segments) &&
                !CanCoalesceFlowRunStyles(previousDrawableRun.Style, flowRun.Style);
            if (hideLeadingSpacesAfterStyleBoundary)
            {
                hideLeadingSpacesAfterBoundary = true;
            }

            PptxTextFlowRun rewritten = flowRun with { Segments = HideSpacesAfterBoundaryPunctuation(flowRun.Segments, ref hideLeadingSpacesAfterBoundary) };
            runs.Add(rewritten);
            if (HasDrawableText(rewritten.Segments))
            {
                previousDrawableRun = rewritten;
            }
        }

        if (UsesHighlightedSyntheticBoldItalicParagraphSpacing(runs, advanceEstimator))
        {
            for (int i = 0; i < runs.Count; i++)
            {
                PptxTextFlowRun run = runs[i];
                if (UsesSyntheticBoldItalicSpacing(run.Style, advanceEstimator))
                {
                    runs[i] = run with
                    {
                        Style = run.Style with
                        {
                            CharacterSpacing = PptxTextMetricRules.OfficeSyntheticBoldItalicCharacterSpacing(run.Style.FontSize)
                        }
                    };
                }
            }
        }

        return new PptxTextFlowParagraph(paragraph, paragraph.Style, runs.ToArray());
    }

    private static bool UsesHighlightedSyntheticBoldItalicParagraphSpacing(IReadOnlyList<PptxTextFlowRun> runs, TextAdvanceEstimator advanceEstimator)
    {
        bool hasHighlightedRun = false;
        bool hasMatchingDrawableRun = false;
        foreach (PptxTextFlowRun run in runs)
        {
            bool hasDrawableText = HasDrawableText(run.Segments);
            if (!hasDrawableText)
            {
                continue;
            }

            hasHighlightedRun |= run.Style.Highlight is not null;
            hasMatchingDrawableRun |= UsesSyntheticBoldItalicSpacing(run.Style, advanceEstimator);
        }

        return hasHighlightedRun && hasMatchingDrawableRun;
    }

    private static bool UsesSyntheticBoldItalicSpacing(ResolvedRunTextStyle style, TextAdvanceEstimator advanceEstimator)
    {
        return Math.Abs(style.CharacterSpacing) <= PptxTextMetricRules.TextStateTolerance &&
            style.Bold &&
            style.Italic &&
            (advanceEstimator.RequestedStyleRequiresSyntheticBold(style.Typeface, style.Bold, style.Italic) ||
             advanceEstimator.RequestedStyleRequiresSyntheticItalic(style.Typeface, style.Bold, style.Italic));
    }

    private static bool StartsWithDrawableRegularSpace(IReadOnlyList<PptxTextFlowSegment> segments)
    {
        foreach (PptxTextFlowSegment segment in segments)
        {
            if (!segment.Draw)
            {
                continue;
            }

            return segment.Kind == PptxTextFlowSegmentKind.Text &&
                segment.Text.Length != 0 &&
                segment.AdvanceText.Length != 0 &&
                segment.Text[0] == ' ' &&
                segment.AdvanceText[0] == ' ';
        }

        return false;
    }

    private static bool HasDrawableText(IReadOnlyList<PptxTextFlowSegment> segments)
    {
        return segments.Any(static segment => segment.Draw && segment.Kind == PptxTextFlowSegmentKind.Text && segment.Text.Length != 0);
    }

    private static bool CanCoalesceFlowRunStyles(ResolvedRunTextStyle left, ResolvedRunTextStyle right)
    {
        return Math.Abs(left.FontSize - right.FontSize) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.CharacterSpacing - right.CharacterSpacing) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.BaselineOffset - right.BaselineOffset) < PptxTextMetricRules.CoordinateTolerance &&
            left.Color.Equals(right.Color) &&
            Math.Abs(left.Alpha - right.Alpha) < PptxTextMetricRules.TextStateTolerance &&
            TextOutlinesEqual(left.Outline, right.Outline) &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.Underline == right.Underline &&
            string.Equals(left.UnderlineValue, right.UnderlineValue, StringComparison.OrdinalIgnoreCase) &&
            left.Strike == right.Strike &&
            string.Equals(left.StrikeValue, right.StrikeValue, StringComparison.OrdinalIgnoreCase) &&
            left.KerningEnabled == right.KerningEnabled &&
            string.Equals(left.Typeface, right.Typeface, StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<PptxTextFlowSegment> HideSpacesAfterBoundaryPunctuation(IReadOnlyList<PptxTextFlowSegment> segments, ref bool hideLeadingSpaces)
    {
        var rewritten = new List<PptxTextFlowSegment>(segments.Count);
        foreach (PptxTextFlowSegment segment in segments)
        {
            foreach (PptxTextFlowSegment current in HideLeadingSpacesIfNeeded(segment, hideLeadingSpaces))
            {
                rewritten.Add(current);
                if (current.Kind == PptxTextFlowSegmentKind.BoundaryPunctuation)
                {
                    hideLeadingSpaces = true;
                }
                else if (current.AdvanceText.Any(static c => c != ' '))
                {
                    hideLeadingSpaces = false;
                }
            }
        }

        return rewritten.ToArray();
    }

    private static IReadOnlyList<PptxTextFlowSegment> HideLeadingSpacesIfNeeded(PptxTextFlowSegment segment, bool hideLeadingSpaces)
    {
        if (!hideLeadingSpaces ||
            !segment.Draw ||
            segment.Kind != PptxTextFlowSegmentKind.Text ||
            segment.Text.Length == 0 ||
            segment.AdvanceText.Length == 0 ||
            segment.Text[0] != ' ' ||
            segment.AdvanceText[0] != ' ')
        {
            return [segment];
        }

        int hiddenLength = 0;
        while (hiddenLength < segment.AdvanceText.Length &&
            hiddenLength < segment.Text.Length &&
            segment.AdvanceText[hiddenLength] == ' ' &&
            segment.Text[hiddenLength] == ' ')
        {
            hiddenLength++;
        }

        string hiddenAdvance = segment.AdvanceText[..hiddenLength];
        PptxTextFlowSegment hidden = new(string.Empty, hiddenAdvance, PptxTextFlowSegmentKind.HiddenAdvance, Draw: false, PreventCoalesce: true, segment.FontScale);
        if (hiddenLength < segment.Text.Length)
        {
            return
            [
                hidden,
                segment with
                {
                    Text = segment.Text[hiddenLength..],
                    AdvanceText = segment.AdvanceText[hiddenLength..]
                }
            ];
        }

        return [hidden];
    }

    private static PptxTextFlowRun BuildTextFlowRun(PptxTextRunModel run, XElement? defaultRunProperties, bool attachSpacesToFollowingWord)
    {
        if (run.Kind == PptxTextRunKind.Break)
        {
            return new PptxTextFlowRun(run, run.Style, [new PptxTextFlowSegment("\n", "\n", PptxTextFlowSegmentKind.Break, Draw: false, PreventCoalesce: true)]);
        }

        var segments = new List<PptxTextFlowSegment>();
        string[] tabParts = run.Text.Split('\t');
        for (int tabPartIndex = 0; tabPartIndex < tabParts.Length; tabPartIndex++)
        {
            if (tabPartIndex > 0)
            {
                segments.Add(new PptxTextFlowSegment(" ", " ", PptxTextFlowSegmentKind.Tab, Draw: true, PreventCoalesce: true));
            }

            foreach (TextCapsFragment fragment in ApplyTextCaps(tabParts[tabPartIndex], run.Properties, defaultRunProperties))
            {
                if (fragment.Text.Length == 0)
                {
                    continue;
                }

                foreach (PptxTextFlowSegment segment in SplitFlowSegments(fragment.Text, attachSpacesToFollowingWord))
                {
                    segments.Add(segment with
                    {
                        FontScale = fragment.FontScale
                    });
                }
            }
        }

        return new PptxTextFlowRun(run, run.Style, segments);
    }

    private static IReadOnlyList<TextRun> FlattenTextLayout(PptxTextLayoutModel layout)
    {
        return FlattenTextLayoutToSpans(layout).Select(span => span.Run).ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> FlattenTextLayoutToSpans(PptxTextLayoutModel layout, PresentationFontResolver? fontResolver = null)
    {
        PptxPositionedTextSpan[] spans = layout.Frames
            .SelectMany((frame, frameIndex) => frame.Paragraphs.Select((paragraph, paragraphIndex) => new
            {
                Frame = frame,
                FrameIndex = frameIndex,
                Paragraph = paragraph,
                ParagraphIndex = paragraphIndex
            }))
            .SelectMany(paragraphState => paragraphState.Paragraph.Lines.Select((line, lineIndex) => new
            {
                paragraphState.Frame,
                paragraphState.FrameIndex,
                paragraphState.Paragraph,
                paragraphState.ParagraphIndex,
                Line = line,
                LineIndex = lineIndex
            }))
            .SelectMany(lineState => lineState.Line.Spans.Select((span, spanIndex) => new PptxPositionedTextSpan(
                span.SourceRun,
                lineState.Line.Box,
                lineState.FrameIndex,
                lineState.ParagraphIndex,
                lineState.Paragraph.Model.Bullet.Kind.ToString(),
                lineState.Paragraph.Model.Bullet.AutoNumberType,
                lineState.Paragraph.Model.Bullet.AutoNumberStartAt,
                lineState.LineIndex,
                spanIndex,
                lineState.Line.Spans.Count,
                lineState.Frame.Model.FontScale,
                OoxUnits.EmuToPoints(lineState.Frame.Model.Bounds.X),
                OoxUnits.EmuToPoints(lineState.Frame.Model.Bounds.Y),
                OoxUnits.EmuToPoints(lineState.Frame.Model.Bounds.Width),
                OoxUnits.EmuToPoints(lineState.Frame.Model.Bounds.Height),
                lineState.Frame.Model.TableRowIndex,
                lineState.Frame.Model.TableColumnIndex,
                lineState.Frame.Model.TableRowSpan,
                lineState.Frame.Model.TableColumnSpan,
                lineState.Frame.Model.Insets.Left,
                lineState.Frame.Model.Insets.Right,
                lineState.Frame.Model.Insets.Top,
                lineState.Frame.Model.Insets.Bottom,
                lineState.Frame.Model.BodyProperties.WrapMode.ToString(),
                lineState.Frame.Model.BodyProperties.WrapValue,
                lineState.Frame.Model.BodyProperties.VerticalOverflow.ToString(),
                lineState.Frame.Model.BodyProperties.VerticalOverflowValue,
                lineState.Frame.Model.BodyProperties.VerticalOverflowSource.ToString(),
                lineState.Frame.Model.BodyProperties.AutofitModeValue,
                lineState.Frame.Model.TextX,
                lineState.Frame.Model.TextWidth,
                lineState.Frame.Model.TextWrapWidth,
                lineState.Frame.Model.TextHeight,
                lineState.Frame.Model.TextClipX,
                lineState.Frame.Model.TextClipWidth,
                lineState.Frame.Model.TextClipY,
                lineState.Frame.Model.TextClipHeight,
                lineState.Frame.Model.ColumnCount,
                lineState.Frame.Model.ColumnSpacing,
                lineState.Line.Alignment,
                span.Run,
                span.EndX,
                span.Atoms,
                span.GlyphSpan,
                PdfCharacterSpacingOverride: null)))
            .ToArray();
        return AddEllipsisOverflowMarkers(spans, fontResolver);
    }

    private static IReadOnlyList<PptxPositionedTextSpan> AddEllipsisOverflowMarkers(
        IReadOnlyList<PptxPositionedTextSpan> spans,
        PresentationFontResolver? fontResolver)
    {
        if (spans.Count == 0)
        {
            return spans;
        }

        var result = new List<PptxPositionedTextSpan>(spans.Count);
        var advanceEstimator = new TextAdvanceEstimator(fontResolver);
        foreach (IGrouping<int, PptxPositionedTextSpan> frameSpans in spans.GroupBy(span => span.FrameIndex))
        {
            PptxPositionedTextSpan[] frame = frameSpans.ToArray();
            result.AddRange(frame);
            if (!frame.Any(span => string.Equals(span.FrameVerticalOverflowMode, nameof(PptxTextVerticalOverflow.Ellipsis), StringComparison.Ordinal)))
            {
                continue;
            }

            PptxPositionedTextSpan[] visible = frame
                .Where(span => BaselineIntersectsClip(span.Run, span.Run.Y + span.Run.BaselineOffset))
                .ToArray();
            if (visible.Length == 0 || visible.Length == frame.Length)
            {
                continue;
            }

            PptxPositionedTextSpan last = visible
                .OrderBy(span => span.ParagraphIndex)
                .ThenBy(span => span.LineIndex)
                .ThenBy(span => span.SpanIndex)
                .Last();
            result.Add(CreateEllipsisOverflowMarker(last, advanceEstimator));
        }

        return result
            .OrderBy(span => span.FrameIndex)
            .ThenBy(span => span.ParagraphIndex)
            .ThenBy(span => span.LineIndex)
            .ThenBy(span => span.SpanIndex)
            .ToArray();
    }

    private static PptxPositionedTextSpan CreateEllipsisOverflowMarker(PptxPositionedTextSpan anchor, TextAdvanceEstimator advanceEstimator)
    {
        const string ellipsis = "…";
        double width = PptxTextMetricRules.MinimumWidth(
            advanceEstimator.Measure(ellipsis, anchor.Run.FontSize, anchor.Run.FontFamily, anchor.Run.Bold, anchor.Run.Italic, anchor.Run.CharacterSpacing, anchor.Run.KerningEnabled));
        TextRun run = anchor.Run with
        {
            Text = ellipsis,
            X = anchor.EndX,
            Width = width,
            PreventCoalesce = true,
            HighlightColor = null,
            Underline = false,
            Strike = false,
            Outline = null
        };

        return anchor with
        {
            SpanIndex = anchor.LineSpanCount,
            LineSpanCount = anchor.LineSpanCount + 1,
            Run = run,
            EndX = run.X + width,
            Atoms = BuildTextAtoms(run, advanceEstimator, PptxTextAtomKind.Word),
            GlyphSpan = BuildGlyphSpan(run, advanceEstimator)
        };
    }

    private static PptxTextFrameLayout BuildTextFrameLayout(PptxTextFlowFrame flowFrame, PptxDocument document, TextAdvanceEstimator advanceEstimator, bool allowWrapping = true)
    {
        PptxTextFrameLayout layout = BuildTextFrameLayout(flowFrame, document, advanceEstimator, allowWrapping, PptxTextColumnBreakMode.StrictFit);
        if (TryResolveOfficeOverflowColumnLineBalance(layout, out int lineBalanceTarget, out int lineBalanceStartColumn))
        {
            return BuildTextFrameLayout(flowFrame, document, advanceEstimator, allowWrapping, PptxTextColumnBreakMode.LineCountBalance, lineBalanceTarget, lineBalanceStartColumn);
        }

        return ShouldUseOfficeOverflowColumnBalance(layout)
            ? BuildTextFrameLayout(flowFrame, document, advanceEstimator, allowWrapping, PptxTextColumnBreakMode.OverflowBalance)
            : layout;
    }

    private static PptxTextFrameLayout BuildTextFrameLayout(PptxTextFlowFrame flowFrame, PptxDocument document, TextAdvanceEstimator advanceEstimator, bool allowWrapping, PptxTextColumnBreakMode columnBreakMode, int lineBalanceTarget = 0, int lineBalanceStartColumn = 0)
    {
        PptxTextFrameModel frame = flowFrame.Model;
        allowWrapping &= TextBodyAllowsWrapping(frame.BodyProperties);
        double cursorLineTop = flowFrame.Box.CursorTop;
        int columnIndex = 0;
        int linesInCurrentColumn = 0;
        double totalColumnSpacing = frame.ColumnSpacing * (frame.ColumnCount - 1);
        double columnWidth = frame.ColumnCount <= 1
            ? frame.TextWidth
            : Math.Max(1d, (frame.TextWidth - totalColumnSpacing) / frame.ColumnCount);
        double columnWrapWidth = frame.ColumnCount <= 1
            ? frame.TextWrapWidth
            : Math.Max(1d, (frame.TextWrapWidth - totalColumnSpacing) / frame.ColumnCount);
        double columnStartX = frame.TextX;
        bool strictClip = ClipsTextVerticalOverflow(frame.BodyProperties.VerticalOverflow);
        int autoNumberValue = 1;
        bool hasPlacedParagraph = false;
        var paragraphLayouts = new List<PptxTextParagraphLayout>();

        foreach (PptxTextFlowParagraph flowParagraph in flowFrame.Paragraphs)
        {
            PptxTextParagraphModel paragraph = flowParagraph.Model;
            var lineLayouts = new List<PptxTextLineLayout>();
            ResolvedParagraphTextStyle paragraphStyle = flowParagraph.Style;
            if (!paragraph.HasVisibleContent)
            {
                if (paragraph.HasLayoutContent)
                {
                    double emptyFontSize = paragraph.EndParagraphStyle.FontSize;
                    cursorLineTop -= (hasPlacedParagraph ? paragraph.EmptySpacingBefore : 0d) + ReadParagraphAdvance(paragraphStyle.LineSpacing, emptyFontSize) + paragraph.EmptySpacingAfter;
                    hasPlacedParagraph = true;
                }

                paragraphLayouts.Add(new PptxTextParagraphLayout(paragraph, lineLayouts));
                continue;
            }

            if (paragraph.Bullet.Kind != PptxParagraphBulletKind.AutoNumber)
            {
                autoNumberValue = 1;
            }

            string? bulletText = ReadBulletText(paragraph.Bullet, ref autoNumberValue);
            bool bulletPending = bulletText is not null;
            double effectiveTextWidth = columnWrapWidth;
            double bulletX = columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging);
            double paragraphTextX = bulletText is null
                ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
            bool clipsLocally = frame.TextClipX != 0d ||
                frame.TextClipWidth < document.SlideWidthPoints - PptxTextMetricRules.CoordinateTolerance;
            double columnClipX = clipsLocally ? columnStartX : frame.TextClipX;
            double columnClipWidth = clipsLocally ? columnWidth : frame.TextClipWidth;
            if (hasPlacedParagraph)
            {
                cursorLineTop -= paragraphStyle.SpacingBefore;
                double nextLineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, paragraphStyle.FontSize);
                MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, nextLineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: false);
                columnClipX = clipsLocally ? columnStartX : frame.TextClipX;
                columnClipWidth = clipsLocally ? columnWidth : frame.TextClipWidth;
                bulletX = columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging);
                paragraphTextX = bulletText is null
                    ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                    : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
            }

            bool afterManualLineBreak = false;
            bool afterLeadingManualLineBreak = false;
            bool shapeAutoFit = HasShapeAutoFit(frame.BodyProperties);
            bool useExplicitMultipleBaselineOffset = ShouldUseExplicitMultipleBaselineOffset(frame, paragraphStyle.LineSpacing);
            double cursorY = cursorLineTop - ReadFirstLineBaselineOffset(paragraph, paragraphStyle.LineSpacing, advanceEstimator, frame.UseOfficeBaselineFloor, shapeAutoFit, useExplicitMultipleBaselineOffset);
            double cursorX = paragraphTextX;
            double maxFontSize = 0d;
            var line = new TextLayoutLine(paragraphTextX);
            int? previousAdvanceCodePoint = null;
            double pendingVisibleLeadingAdjustment = 0d;
            PptxTextSpanLayout? noBreakAnchorSpan = null;
            string pendingNoBreakAdvanceText = string.Empty;
            int remainingDrawableSegments = CountDrawableTextSegments(flowParagraph);
            foreach (PptxTextFlowRun flowRun in flowParagraph.Runs)
            {
                PptxTextRunModel modelRun = flowRun.Source;
                if (modelRun.Kind == PptxTextRunKind.Break)
                {
                    double lineFontSize = ResolveLineFontSize(maxFontSize, flowRun.Style.FontSize);
                    bool leadingManualBreak = line.Spans.Count == 0;
                    bool useManualBreakFallback = leadingManualBreak || !shapeAutoFit;
                    AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: false, advanceEstimator);
                    double lineAdvance = useManualBreakFallback
                        ? ReadManualBreakLineAdvance(paragraphStyle.LineSpacing, lineFontSize)
                        : ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                    cursorLineTop -= lineAdvance;
                    MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, lineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
                    columnClipX = clipsLocally ? columnStartX : frame.TextClipX;
                    columnClipWidth = clipsLocally ? columnWidth : frame.TextClipWidth;
                    paragraphTextX = bulletText is null
                        ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                        : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
                    cursorY = double.NaN;
                    afterManualLineBreak = true;
                    afterLeadingManualLineBreak = useManualBreakFallback;
                    cursorX = paragraphTextX;
                    line.Reset(paragraphTextX);
                    maxFontSize = 0d;
                    previousAdvanceCodePoint = null;
                    pendingVisibleLeadingAdjustment = 0d;
                    noBreakAnchorSpan = null;
                    pendingNoBreakAdvanceText = string.Empty;
                    continue;
                }

                ResolvedRunTextStyle runStyle = flowRun.Style;
                if (double.IsNaN(cursorY))
                {
                    cursorY = cursorLineTop - (afterLeadingManualLineBreak
                        ? ManualBreakBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset)
                        : LineBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset));
                    afterManualLineBreak = false;
                    afterLeadingManualLineBreak = false;
                }

                if (bulletPending)
                {
                    BulletStyle bulletStyle = ReadBulletStyle(paragraph.Bullet, runStyle.FontSize, runStyle.Color, runStyle.Typeface);
                    maxFontSize = Math.Max(maxFontSize, bulletStyle.FontSize);
                    double bulletWidth = PptxTextMetricRules.MinimumWidth(effectiveTextWidth - (bulletX - columnStartX));
                    double bulletEndX = bulletX + advanceEstimator.Measure(bulletText!, bulletStyle.FontSize, bulletStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing);
                    TextRun bulletRun = new(bulletText!, bulletX, cursorY, bulletWidth, frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, bulletStyle.FontSize, runStyle.CharacterSpacing, 0d, bulletStyle.Color, 1d, null, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, bulletStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, StrictClip: strictClip);
                    line.Add(modelRun, bulletRun, bulletEndX, BuildTextAtoms(bulletRun, advanceEstimator, PptxTextAtomKind.Word), BuildGlyphSpan(bulletRun, advanceEstimator));
                    bulletPending = false;
                }

                foreach (PptxTextFlowSegment flowSegment in flowRun.Segments)
                {
                    if (flowSegment.Kind == PptxTextFlowSegmentKind.Break)
                    {
                        double lineFontSize = ResolveLineFontSize(maxFontSize, runStyle.FontSize);
                        bool leadingManualBreak = line.Spans.Count == 0;
                        bool useManualBreakFallback = leadingManualBreak || !shapeAutoFit;
                        AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: false, advanceEstimator);
                        double lineAdvance = useManualBreakFallback
                            ? ReadManualBreakLineAdvance(paragraphStyle.LineSpacing, lineFontSize)
                            : ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                        cursorLineTop -= lineAdvance;
                        MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, lineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
                        columnClipX = clipsLocally ? columnStartX : frame.TextClipX;
                        columnClipWidth = clipsLocally ? columnWidth : frame.TextClipWidth;
                        paragraphTextX = bulletText is null
                            ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                            : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
                        cursorY = double.NaN;
                        afterManualLineBreak = true;
                        afterLeadingManualLineBreak = useManualBreakFallback;
                        cursorX = paragraphTextX;
                        line.Reset(paragraphTextX);
                        maxFontSize = 0d;
                        previousAdvanceCodePoint = null;
                        pendingVisibleLeadingAdjustment = 0d;
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                        continue;
                    }

                    if (double.IsNaN(cursorY))
                    {
                        cursorY = cursorLineTop - (afterLeadingManualLineBreak
                            ? ManualBreakBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset)
                            : LineBaselineOffset(runStyle.NominalFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset));
                        afterManualLineBreak = false;
                        afterLeadingManualLineBreak = false;
                    }

                    if (flowSegment.Kind == PptxTextFlowSegmentKind.Tab)
                    {
                        double tabSpaceWidth = advanceEstimator.Measure(" ", runStyle.FontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        TextRun tabRun = new(" ", cursorX, cursorY, PptxTextMetricRules.MinimumWidth(tabSpaceWidth), frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, runStyle.FontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, PreventCoalesce: true, Outline: runStyle.Outline, StrictClip: strictClip);
                        line.Add(modelRun, tabRun, cursorX + tabSpaceWidth, BuildTextAtoms(tabRun, advanceEstimator, PptxTextAtomKind.Tab), BuildGlyphSpan(tabRun, advanceEstimator));
                        cursorX = ResolveNextTabX(cursorX, paragraphTextX, paragraphStyle.TabStops);
                        line.AdvanceTo(cursorX);
                        previousAdvanceCodePoint = null;
                        pendingVisibleLeadingAdjustment = 0d;
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                        continue;
                    }

                    double fragmentFontSize = runStyle.FontSize * flowSegment.FontScale;
                    string currentSegment = flowSegment.Text;
                    string currentAdvanceText = flowSegment.AdvanceText;
                    bool isNoBreakHiddenAdvance = flowSegment.Kind == PptxTextFlowSegmentKind.NoBreakHiddenAdvance;
                    bool isDrawableTextSegment = flowSegment.Draw && currentAdvanceText.TrimStart().Length > 0;
                    bool isFinalShortWordSegment = isDrawableTextSegment &&
                        remainingDrawableSegments == 1 &&
                        lineLayouts.Count == 0 &&
                        IsShortWordSegment(currentAdvanceText);
                    double segmentIntrinsicWidth = advanceEstimator.Measure(currentAdvanceText, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                    double segmentBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, currentAdvanceText, previousAdvanceCodePoint, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                    double segmentWidth = Math.Max(0d, segmentIntrinsicWidth + segmentBoundaryAdjustment);
                    bool splitOverwideFirstSegment = allowWrapping &&
                        flowSegment.Kind == PptxTextFlowSegmentKind.Text &&
                        flowSegment.Draw &&
                        currentSegment == currentAdvanceText &&
                        currentSegment.Length > 1 &&
                        line.Spans.Count == 0 &&
                        segmentWidth > effectiveTextWidth + PptxTextMetricRules.CoordinateTolerance;
                    if ((frame.Orientation != PptxTextOrientation.Horizontal &&
                            flowSegment.Kind == PptxTextFlowSegmentKind.Text &&
                            flowSegment.Draw &&
                            currentSegment == currentAdvanceText &&
                            currentSegment.Length > 1 &&
                            segmentWidth > frame.TextWidth) ||
                        splitOverwideFirstSegment)
                    {
                        double chunkMaxWidth = frame.Orientation == PptxTextOrientation.Horizontal
                            ? effectiveTextWidth
                            : frame.TextWidth;
                        double chunkFitTolerance = PptxTextMetricRules.WrapFitTolerance(fragmentFontSize);
                        string[] chunks = SplitTextIntoFittingChunks(currentSegment, chunkMaxWidth + chunkFitTolerance, fragmentFontSize, runStyle, advanceEstimator);
                        for (int chunkIndex = 0; chunkIndex < chunks.Length; chunkIndex++)
                        {
                            string chunk = chunks[chunkIndex];
                            double chunkWidth = advanceEstimator.Measure(chunk, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                            double chunkBoundaryAdjustment = chunkIndex == 0
                                ? MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, chunk, previousAdvanceCodePoint, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled)
                                : 0d;
                            double chunkTotalWidth = Math.Max(0d, chunkWidth + chunkBoundaryAdjustment);
                            maxFontSize = Math.Max(maxFontSize, fragmentFontSize);
                            double chunkX = cursorX + chunkBoundaryAdjustment;
                            double chunkClipX = frame.Orientation == PptxTextOrientation.Horizontal ? columnClipX : frame.TextClipX;
                            double chunkClipWidth = frame.Orientation == PptxTextOrientation.Horizontal ? columnClipWidth : frame.TextClipWidth;
                            TextRun textRun = new(chunk, chunkX, cursorY, PptxTextMetricRules.MinimumWidth(chunkWidth), frame.TextHeight, chunkClipX, frame.TextClipY, chunkClipWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, flowSegment.PreventCoalesce, Outline: runStyle.Outline, StrictClip: strictClip);
                            double chunkLeadingAdjustment = pendingVisibleLeadingAdjustment + chunkBoundaryAdjustment;
                            line.Add(modelRun, textRun, cursorX + chunkTotalWidth, BuildTextAtoms(textRun, advanceEstimator), BuildGlyphSpan(textRun, advanceEstimator, chunkLeadingAdjustment));
                            pendingVisibleLeadingAdjustment = 0d;
                            cursorX += chunkTotalWidth;
                            line.AdvanceTo(cursorX);
                            if (chunkIndex < chunks.Length - 1)
                            {
                                double lineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
                                double lineTextX = frame.Orientation == PptxTextOrientation.Horizontal ? columnStartX : frame.TextX;
                                double lineTextWidth = frame.Orientation == PptxTextOrientation.Horizontal ? effectiveTextWidth : frame.TextWidth;
                                AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, lineTextX, lineTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator);
                                double lineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                                cursorLineTop -= lineAdvance;
                                MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, lineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
                                columnClipX = clipsLocally ? columnStartX : frame.TextClipX;
                                columnClipWidth = clipsLocally ? columnWidth : frame.TextClipWidth;
                                paragraphTextX = bulletText is null
                                    ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                                    : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
                                cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
                                cursorX = paragraphTextX;
                                line.Reset(paragraphTextX);
                                maxFontSize = 0d;
                                previousAdvanceCodePoint = null;
                                pendingVisibleLeadingAdjustment = 0d;
                                noBreakAnchorSpan = null;
                                pendingNoBreakAdvanceText = string.Empty;
                            }
                        }

                        previousAdvanceCodePoint = LastCodePoint(currentAdvanceText);
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                        if (isDrawableTextSegment)
                        {
                            remainingDrawableSegments--;
                        }

                        continue;
                    }

                    bool usesCenteredShapeAutoFit = HasShapeAutoFit(frame.BodyProperties) &&
                        paragraphStyle.Alignment == TextAlignment.Center;
                    double wrapTolerance = usesCenteredShapeAutoFit
                        ? PptxTextMetricRules.CoordinateTolerance
                        : IsCenteredTableCellText(frame, paragraphStyle)
                        ? PptxTextMetricRules.CenteredTableCellWrapTolerance(fragmentFontSize, effectiveTextWidth)
                        : bulletText is not null
                        ? PptxTextMetricRules.BulletWrapFitTolerance(fragmentFontSize)
                        : HasShapeAutoFit(frame.BodyProperties)
                        ? PptxTextMetricRules.ShapeAutoFitWrapTolerance(fragmentFontSize, effectiveTextWidth)
                        : !HasNoAutoFit(frame.BodyProperties) ||
                        IsWordJustifiedAlignment(paragraphStyle.Alignment) ||
                        paragraphStyle.Alignment == TextAlignment.Distributed
                        ? PptxTextMetricRules.CoordinateTolerance
                        : PptxTextMetricRules.CoordinateTolerance;
                    if (isFinalShortWordSegment &&
                        HasShapeAutoFit(frame.BodyProperties) &&
                        !usesCenteredShapeAutoFit)
                    {
                        wrapTolerance = PptxTextMetricRules.FinalWordWrapTolerance(fragmentFontSize, effectiveTextWidth);
                    }

                    bool overflowsLine = allowWrapping &&
                        cursorX > paragraphTextX &&
                        cursorX + segmentWidth > columnStartX + effectiveTextWidth + wrapTolerance;
                    if (overflowsLine)
                    {
                        bool movedNoBreakCluster = false;
                        PptxTextSpanLayout? movedNoBreakSpan = null;
                        string movedNoBreakAdvanceText = string.Empty;
                        if (flowSegment.Draw &&
                            noBreakAnchorSpan is not null &&
                            pendingNoBreakAdvanceText.Length != 0 &&
                            line.Spans.Count > 0 &&
                            line.Spans[^1].Equals(noBreakAnchorSpan) &&
                            line.TryRemoveLastSpan(out PptxTextSpanLayout? removedNoBreakSpan))
                        {
                            movedNoBreakCluster = true;
                            movedNoBreakSpan = removedNoBreakSpan;
                            movedNoBreakAdvanceText = pendingNoBreakAdvanceText;
                            cursorX = line.EndX;
                        }

                        if (flowSegment.Draw)
                        {
                            int leadingSpaceCount = CountLeadingSpaces(currentSegment);
                            if (leadingSpaceCount > 0)
                            {
                                string lineEndSpaces = currentSegment[..leadingSpaceCount];
                                double lineEndSpaceWidth = advanceEstimator.Measure(lineEndSpaces, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                                TextRun spaceRun = new(lineEndSpaces, cursorX, cursorY, PptxTextMetricRules.MinimumWidth(lineEndSpaceWidth), frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, PreventCoalesce: false, Outline: runStyle.Outline, StrictClip: strictClip);
                                line.Add(modelRun, spaceRun, cursorX + lineEndSpaceWidth, BuildTextAtoms(spaceRun, advanceEstimator, PptxTextAtomKind.Space), BuildGlyphSpan(spaceRun, advanceEstimator));
                                cursorX += lineEndSpaceWidth;
                                line.AdvanceTo(cursorX);
                            }
                        }

                        double lineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
                        AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, lineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: IsWordJustifiedAlignment(paragraphStyle.Alignment), distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator);
                        double lineAdvance = ReadLineAdvance(paragraphStyle.LineSpacing, lineFontSize);
                        cursorLineTop -= lineAdvance;
                        MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, lineAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
                        columnClipX = clipsLocally ? columnStartX : frame.TextClipX;
                        columnClipWidth = clipsLocally ? columnWidth : frame.TextClipWidth;
                        paragraphTextX = bulletText is null
                            ? columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft + paragraphStyle.Indent.Hanging)
                            : columnStartX + PptxTextMetricRules.ClampNonNegative(paragraphStyle.Indent.MarginLeft);
                        cursorY = cursorLineTop - LineBaselineOffset(fragmentFontSize, paragraphStyle.LineSpacing, runStyle, advanceEstimator, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
                        cursorX = paragraphTextX;
                        line.Reset(paragraphTextX);
                        maxFontSize = 0d;
                        previousAdvanceCodePoint = null;
                        pendingVisibleLeadingAdjustment = 0d;
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                        if (movedNoBreakCluster)
                        {
                            TextRun movedRun = movedNoBreakSpan!.Run with { X = cursorX, Y = cursorY };
                            double movedEndX = cursorX + movedRun.Width;
                            maxFontSize = Math.Max(maxFontSize, movedRun.FontSize);
                            line.Add(
                                movedNoBreakSpan.SourceRun,
                                movedRun,
                                movedEndX,
                                BuildTextAtoms(movedRun, advanceEstimator),
                                BuildGlyphSpan(movedRun, advanceEstimator));
                            cursorX = movedEndX;
                            previousAdvanceCodePoint = LastCodePoint(movedRun.Text);

                            double hiddenIntrinsicWidth = advanceEstimator.Measure(movedNoBreakAdvanceText, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                            double hiddenBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, movedNoBreakAdvanceText, previousAdvanceCodePoint, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                            double hiddenWidth = Math.Max(0d, hiddenIntrinsicWidth + hiddenBoundaryAdjustment);
                            pendingVisibleLeadingAdjustment = hiddenBoundaryAdjustment;
                            cursorX += hiddenWidth;
                            line.AdvanceTo(cursorX);
                            previousAdvanceCodePoint = LastCodePoint(movedNoBreakAdvanceText);
                        }

                        currentSegment = currentSegment.TrimStart();
                        currentAdvanceText = currentAdvanceText.TrimStart();
                        segmentIntrinsicWidth = advanceEstimator.Measure(currentAdvanceText, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        segmentBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(advanceEstimator, currentAdvanceText, previousAdvanceCodePoint, fragmentFontSize, runStyle.Typeface, runStyle.Bold, runStyle.Italic, runStyle.CharacterSpacing, runStyle.KerningEnabled);
                        segmentWidth = Math.Max(0d, segmentIntrinsicWidth + segmentBoundaryAdjustment);
                    }

                    if (currentAdvanceText.Length == 0)
                    {
                        continue;
                    }

                    if (flowSegment.Draw && currentSegment.Length != 0)
                    {
                        maxFontSize = Math.Max(maxFontSize, fragmentFontSize);
                        double textRunX = cursorX + segmentBoundaryAdjustment;
                        TextRun textRun = new(currentSegment, textRunX, cursorY, PptxTextMetricRules.MinimumWidth(segmentIntrinsicWidth), frame.TextHeight, columnClipX, frame.TextClipY, columnClipWidth, frame.TextClipHeight, fragmentFontSize, runStyle.CharacterSpacing, runStyle.BaselineOffset, runStyle.Color, runStyle.Alpha, runStyle.Highlight, runStyle.Bold, runStyle.Italic, runStyle.Underline, runStyle.Strike, runStyle.KerningEnabled, paragraphStyle.Alignment, runStyle.Typeface, frame.TextRotationDegrees, frame.RotationCenterX, frame.RotationCenterY, frame.TextFlipHorizontal, frame.TextFlipVertical, flowSegment.PreventCoalesce, Outline: runStyle.Outline, StrictClip: strictClip);
                        double leadingAdjustment = pendingVisibleLeadingAdjustment + segmentBoundaryAdjustment;
                        line.Add(modelRun, textRun, cursorX + segmentWidth, BuildTextAtoms(textRun, advanceEstimator), BuildGlyphSpan(textRun, advanceEstimator, leadingAdjustment));
                        pendingVisibleLeadingAdjustment = 0d;
                        noBreakAnchorSpan = null;
                        pendingNoBreakAdvanceText = string.Empty;
                    }
                    else
                    {
                        pendingVisibleLeadingAdjustment += segmentBoundaryAdjustment;
                    }

                    cursorX += segmentWidth;
                    line.AdvanceTo(cursorX);
                    previousAdvanceCodePoint = LastCodePoint(currentAdvanceText);
                    if (isNoBreakHiddenAdvance)
                    {
                        noBreakAnchorSpan = line.Spans.LastOrDefault();
                        pendingNoBreakAdvanceText = currentAdvanceText;
                    }
                    if (isDrawableTextSegment)
                    {
                        remainingDrawableSegments--;
                    }
                }
            }

            double paragraphLineFontSize = ResolveLineFontSize(maxFontSize, paragraphStyle.FontSize);
            if (afterManualLineBreak && line.Spans.Count == 0)
            {
                paragraphLineFontSize = paragraph.EndParagraphProperties is null
                    ? paragraphLineFontSize
                    : paragraph.EndParagraphStyle.FontSize;
            }

            if (double.IsNaN(cursorY))
            {
                cursorY = cursorLineTop - (afterManualLineBreak
                    ? ManualBreakBaselineOffset(paragraphLineFontSize, paragraphStyle.LineSpacing, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset)
                    : LineBaselineOffset(paragraphLineFontSize, paragraphStyle.LineSpacing, frame.UseOfficeBaselineFloor, useExplicitMultipleBaselineOffset));
            }

            AddAlignedParagraphLine(lineLayouts, line, CreateLineBox(cursorLineTop, cursorY, paragraphStyle.LineSpacing, paragraphLineFontSize, line, advanceEstimator, frame.UseOfficeBaselineFloor), paragraphStyle.Alignment, columnStartX, effectiveTextWidth, justify: false, distribute: paragraphStyle.Alignment == TextAlignment.Distributed, advanceEstimator);
            double paragraphAdvance = ReadParagraphAdvance(paragraphStyle.LineSpacing, paragraphLineFontSize);
            cursorLineTop -= paragraphAdvance + paragraphStyle.SpacingAfter;
            MoveToNextColumnIfNeeded(ref cursorLineTop, ref columnIndex, ref columnStartX, ref linesInCurrentColumn, flowFrame.Box.CursorTop, frame.TextX, columnWidth, frame.ColumnSpacing, frame.ColumnCount, flowFrame.Box, frame.BodyProperties.VerticalOverflow, columnBreakMode, paragraphAdvance, lineBalanceTarget, lineBalanceStartColumn, linePlaced: true);
            hasPlacedParagraph = true;
            paragraphLayouts.Add(new PptxTextParagraphLayout(paragraph, lineLayouts));
        }

        return new PptxTextFrameLayout(frame, paragraphLayouts);
    }

    private static bool ShouldUseOfficeOverflowColumnBalance(PptxTextFrameLayout layout)
    {
        if (!TryReadOverflowColumnLineCounts(layout, out PptxTextLineLayout[] lines, out int[] counts))
        {
            return false;
        }

        PptxTextFrameModel frame = layout.Model;
        if (lines.Length % frame.ColumnCount != 0)
        {
            return false;
        }

        int target = lines.Length / frame.ColumnCount;
        return counts.Take(frame.ColumnCount - 1).All(count => count == target - 1) &&
            counts[^1] == target + frame.ColumnCount - 1;
    }

    private static bool TryResolveOfficeOverflowColumnLineBalance(PptxTextFrameLayout layout, out int lineBalanceTarget, out int lineBalanceStartColumn)
    {
        lineBalanceTarget = 0;
        lineBalanceStartColumn = 0;
        if (!TryReadOverflowColumnLineCounts(layout, out PptxTextLineLayout[] lines, out int[] counts))
        {
            return false;
        }

        PptxTextFrameModel frame = layout.Model;
        int balancedFloor = lines.Length / frame.ColumnCount;
        int balancedCeiling = (int)Math.Ceiling((double)lines.Length / frame.ColumnCount);
        if (counts.Take(frame.ColumnCount - 1).All(count => count < balancedFloor) &&
            counts[^1] >= balancedCeiling + frame.ColumnCount)
        {
            lineBalanceTarget = balancedCeiling;
            return true;
        }

        if (ShouldResolveTrailingOverflowColumnBalance(layout, counts, balancedFloor, balancedCeiling))
        {
            lineBalanceTarget = balancedFloor;
            lineBalanceStartColumn = 1;
            return true;
        }

        if (ShouldResolveOverflowColumnBalanceAcrossContinuedParagraph(layout, lines.Length, counts, balancedCeiling))
        {
            lineBalanceTarget = balancedCeiling;
            return true;
        }

        if (ShouldResolveEvenOverflowColumnBalanceAcrossContinuedParagraph(layout, counts, balancedFloor))
        {
            lineBalanceTarget = balancedFloor + 1;
            lineBalanceStartColumn = 1;
            return true;
        }

        return false;
    }

    private static bool ShouldResolveTrailingOverflowColumnBalance(PptxTextFrameLayout layout, int[] counts, int balancedFloor, int balancedCeiling)
    {
        PptxTextFrameModel frame = layout.Model;
        return frame.ColumnCount >= 3 &&
            counts[0] >= balancedFloor &&
            counts.Skip(1).Take(frame.ColumnCount - 2).Any(count => count < balancedFloor) &&
            counts[^1] > balancedFloor;
    }

    private static bool ShouldResolveOverflowColumnBalanceAcrossContinuedParagraph(PptxTextFrameLayout layout, int lineCount, int[] counts, int balancedCeiling)
    {
        PptxTextFrameModel frame = layout.Model;
        if (lineCount % frame.ColumnCount != frame.ColumnCount - 1 ||
            counts[^1] != balancedCeiling ||
            counts.Take(frame.ColumnCount - 1).Sum() != balancedCeiling * (frame.ColumnCount - 1) - 1 ||
            counts[^2] != balancedCeiling - 1)
        {
            return false;
        }

        var columns = layout.Paragraphs
            .SelectMany((paragraph, paragraphIndex) => paragraph.Lines.Select(line => new
            {
                Line = line,
                ParagraphIndex = paragraphIndex
            }))
            .GroupBy(item => Math.Round(item.Line.StartX, 2))
            .OrderBy(group => group.Key)
            .Select(group => group.ToArray())
            .ToArray();
        if (columns.Length != frame.ColumnCount || columns[^2].Length == 0 || columns[^1].Length == 0)
        {
            return false;
        }

        return columns[^2][^1].ParagraphIndex == columns[^1][0].ParagraphIndex;
    }

    private static bool ShouldResolveEvenOverflowColumnBalanceAcrossContinuedParagraph(PptxTextFrameLayout layout, int[] counts, int balancedFloor)
    {
        PptxTextFrameModel frame = layout.Model;
        if (frame.ColumnCount < 3 ||
            counts.Any(count => count != balancedFloor))
        {
            return false;
        }

        var columns = layout.Paragraphs
            .SelectMany((paragraph, paragraphIndex) => paragraph.Lines.Select(line => new
            {
                Line = line,
                ParagraphIndex = paragraphIndex
            }))
            .GroupBy(item => Math.Round(item.Line.StartX, 2))
            .OrderBy(group => group.Key)
            .Select(group => group.ToArray())
            .ToArray();
        if (columns.Length != frame.ColumnCount || columns[^2].Length == 0 || columns[^1].Length == 0)
        {
            return false;
        }

        return columns[^2][^1].ParagraphIndex == columns[^1][0].ParagraphIndex;
    }

    private static bool TryReadOverflowColumnLineCounts(PptxTextFrameLayout layout, out PptxTextLineLayout[] lines, out int[] counts)
    {
        PptxTextFrameModel frame = layout.Model;
        lines = [];
        counts = [];
        if (frame.ColumnCount <= 1 || frame.BodyProperties.VerticalOverflow != PptxTextVerticalOverflow.Overflow)
        {
            return false;
        }

        lines = layout.Paragraphs.SelectMany(paragraph => paragraph.Lines).ToArray();
        if (lines.Length == 0)
        {
            return false;
        }

        counts = lines
            .GroupBy(line => Math.Round(line.StartX, 2))
            .OrderBy(group => group.Key)
            .Select(group => group.Count())
            .ToArray();
        if (counts.Length != frame.ColumnCount)
        {
            return false;
        }

        return true;
    }

    private static int CountLeadingSpaces(string text)
    {
        int count = 0;
        while (count < text.Length && text[count] == ' ')
        {
            count++;
        }

        return count;
    }

    private static int CountDrawableTextSegments(PptxTextFlowParagraph paragraph)
    {
        return paragraph.Runs
            .SelectMany(run => run.Segments)
            .Count(segment => segment.Draw && segment.AdvanceText.TrimStart().Length > 0);
    }

    private static bool UsesRotatedFrameAutoFit(PptxTextOrientation orientation)
    {
        return orientation is PptxTextOrientation.Vertical or PptxTextOrientation.Vertical270;
    }

    private static bool ClipsTextVerticalOverflow(PptxTextVerticalOverflow overflow)
    {
        return overflow is PptxTextVerticalOverflow.Clip or PptxTextVerticalOverflow.Ellipsis;
    }

    private static bool IsShortWordSegment(string text)
    {
        string trimmed = text.Trim();
        return trimmed.Length is > 0 and <= 16 && !trimmed.Any(char.IsWhiteSpace);
    }

    private static double ResolveLineFontSize(double visibleMaxFontSize, double fallbackFontSize)
    {
        return visibleMaxFontSize > PptxTextMetricRules.TextStateTolerance
            ? visibleMaxFontSize
            : fallbackFontSize;
    }

    private static string[] SplitTextIntoFittingChunks(
        string text,
        double maxWidth,
        double fontSize,
        ResolvedRunTextStyle style,
        TextAdvanceEstimator advanceEstimator)
    {
        var chunks = new List<string>();
        var chunk = new StringBuilder();
        foreach (Rune rune in text.EnumerateRunes())
        {
            string candidate = chunk.ToString() + rune;
            double candidateWidth = advanceEstimator.Measure(candidate, fontSize, style.Typeface, style.Bold, style.Italic, style.CharacterSpacing, style.KerningEnabled);
            if (chunk.Length > 0 && candidateWidth > maxWidth)
            {
                chunks.Add(chunk.ToString());
                chunk.Clear();
            }

            chunk.Append(rune);
        }

        if (chunk.Length > 0)
        {
            chunks.Add(chunk.ToString());
        }

        return chunks.Count == 0 ? [text] : chunks.ToArray();
    }

    private enum PptxTextColumnBreakMode
    {
        StrictFit,
        OverflowBalance,
        LineCountBalance
    }

    private static void MoveToNextColumnIfNeeded(
        ref double cursorLineTop,
        ref int columnIndex,
        ref double columnStartX,
        ref int linesInCurrentColumn,
        double firstColumnTop,
        double frameTextX,
        double columnWidth,
        double columnSpacing,
        int columnCount,
        PptxTextFlowBox box,
        PptxTextVerticalOverflow verticalOverflow,
        PptxTextColumnBreakMode columnBreakMode,
        double lineAdvance,
        int lineBalanceTarget,
        int lineBalanceStartColumn,
        bool linePlaced)
    {
        if (columnCount <= 1 || columnIndex >= columnCount - 1)
        {
            return;
        }

        if (columnBreakMode == PptxTextColumnBreakMode.LineCountBalance &&
            columnIndex >= lineBalanceStartColumn)
        {
            if (linePlaced)
            {
                linesInCurrentColumn++;
            }

            if (lineBalanceTarget <= 0 || linesInCurrentColumn < lineBalanceTarget)
            {
                return;
            }

            columnIndex++;
            columnStartX = frameTextX + columnIndex * (columnWidth + columnSpacing);
            cursorLineTop = firstColumnTop;
            linesInCurrentColumn = 0;
            return;
        }

        double bottom = box.CursorTop - box.TextHeight;
        double nextLineThreshold = verticalOverflow == PptxTextVerticalOverflow.Overflow &&
            columnBreakMode == PptxTextColumnBreakMode.OverflowBalance
            ? cursorLineTop
            : cursorLineTop - lineAdvance;
        if (nextLineThreshold >= bottom - PptxTextMetricRules.TextStateTolerance)
        {
            return;
        }

        columnIndex++;
        columnStartX = frameTextX + columnIndex * (columnWidth + columnSpacing);
        cursorLineTop = firstColumnTop;
        linesInCurrentColumn = 0;
    }

    private static GroupTransform ReadAncestorGroupTransform(XElement shape)
    {
        GroupTransform transform = GroupTransform.Identity;
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp").Reverse())
        {
            transform = transform.Combine(ReadGroupTransform(group));
        }

        return transform;
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForSceneNode(PptxSceneNode node, PptxRenderContext context, PptxColorMap colorMap, bool includePlaceholders)
    {
        return node.TextBody is null
            ? []
            : ReadTextSpansForShape(node.Source, context.Document, context.Theme, colorMap, context.SlideNumber, includePlaceholders, context.InheritedXml, context.FontResolver);
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForShape(
        XElement shape,
        PptxRenderContext context,
        bool includePlaceholders)
    {
        return ReadTextSpansForShape(shape, context.Document, context.Theme, context.SlideColorMap, context.SlideNumber, includePlaceholders, context.InheritedXml, context.FontResolver);
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForShape(
        XElement shape,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PresentationFontResolver? fontResolver = null)
    {
        XElement current = new(shape);
        foreach (XElement group in shape.Ancestors(PresentationNamespace + "grpSp"))
        {
            var groupCopy = new XElement(PresentationNamespace + "grpSp");
            if (group.Element(PresentationNamespace + "grpSpPr") is { } properties)
            {
                groupCopy.Add(new XElement(properties));
            }

            groupCopy.Add(current);
            current = groupCopy;
        }

        var slide = new XDocument(
            new XElement(PresentationNamespace + "sld",
                new XElement(PresentationNamespace + "cSld",
                    new XElement(PresentationNamespace + "spTree", current))));
        return FlattenTextLayoutToSpans(BuildTextLayoutModel(slide, document, theme, colorMap, slideNumber, includePlaceholders, placeholderSources, fontResolver), fontResolver);
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTextSpansForTableCellTextFrame(PptxTableCellTextFrame tableFrame, PptxRenderContext context)
    {
        PptxTextFrameModel frameModel = BuildTextFrameModel(tableFrame, context.Document, context.Theme, context.SlideNumber, context.InheritedXml);
        PptxTextFrameLayout layout = BuildTextFrameLayout(frameModel, context.Document, new TextAdvanceEstimator(context.FontResolver));
        return FlattenTextLayoutToSpans(new PptxTextLayoutModel([layout]), context.FontResolver);
    }

    private static IReadOnlyList<XElement> FindInheritedPlaceholderShapes(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        return PptxPlaceholderMatcher.FindInheritedPlaceholderShapes(shape, placeholderSources);
    }

    private static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        return PptxTextStyleInheritance.FindInheritedTextStyle(shape, placeholderSources, levelName);
    }

    private static XElement? FindDefaultTextStyle(IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        return PptxTextStyleInheritance.FindDefaultTextStyle(placeholderSources, levelName);
    }

    private static IEnumerable<PptxTextFlowSegment> SplitFlowSegments(string text, bool attachSpacesToFollowingWord)
    {
        if (!attachSpacesToFollowingWord)
        {
            foreach (PptxTextFlowSegment segment in SplitFlowSegmentsWithTrailingSpaces(text))
            {
                yield return segment;
            }

            yield break;
        }

        int index = 0;
        while (index < text.Length)
        {
            int start = index;
            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            if (index >= text.Length)
            {
                if (index > start)
                {
                    foreach (PptxTextFlowSegment segment in SplitControlSegments(text[start..index]))
                    {
                        yield return segment;
                    }
                }

                yield break;
            }

            while (index < text.Length && text[index] != ' ')
            {
                index++;
                if (text[index - 1] == '-' && index < text.Length && text[index] != ' ')
                {
                    break;
                }
            }

            if (index > start)
            {
                foreach (PptxTextFlowSegment segment in SplitControlSegments(text[start..index]))
                {
                    yield return segment;
                }
            }
        }
    }

    private static IEnumerable<PptxTextFlowSegment> SplitFlowSegmentsWithTrailingSpaces(string text)
    {
        int index = 0;
        while (index < text.Length)
        {
            int start = index;
            while (index < text.Length && text[index] != ' ')
            {
                index++;
                if (text[index - 1] == '-' && index < text.Length && text[index] != ' ')
                {
                    break;
                }
            }

            while (index < text.Length && text[index] == ' ')
            {
                index++;
            }

            if (index > start)
            {
                foreach (PptxTextFlowSegment segment in SplitControlSegments(text[start..index]))
                {
                    yield return segment;
                }
            }
        }
    }

    private static IEnumerable<PptxTextFlowSegment> SplitControlSegments(string text)
    {
        var builder = new StringBuilder();
        bool nextPreventsCoalesce = false;
        bool hideLeadingSpaces = false;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '\r' || c == '\n')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
                    builder.Clear();
                }

                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                {
                    i++;
                }

                yield return new PptxTextFlowSegment("\n", "\n", PptxTextFlowSegmentKind.Break, Draw: false, PreventCoalesce: true);
                nextPreventsCoalesce = false;
                hideLeadingSpaces = false;
                continue;
            }

            if (hideLeadingSpaces && c == ' ')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
                    builder.Clear();
                }

                yield return new PptxTextFlowSegment(string.Empty, c.ToString(), PptxTextFlowSegmentKind.HiddenAdvance, Draw: false, PreventCoalesce: true);
                nextPreventsCoalesce = true;
                continue;
            }

            hideLeadingSpaces = false;
            if (c == '\u00A0' || c == '\u202F')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
                    builder.Clear();
                }

                yield return new PptxTextFlowSegment(string.Empty, c.ToString(), PptxTextFlowSegmentKind.NoBreakHiddenAdvance, Draw: false, PreventCoalesce: true);
                nextPreventsCoalesce = true;
                continue;
            }

            if (c == '\u00AD')
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
                    builder.Clear();
                }

                nextPreventsCoalesce = true;
                continue;
            }

            if (IsOfficeTextOperationBoundaryPunctuation(c))
            {
                if (builder.Length > 0)
                {
                    yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
                    builder.Clear();
                }

                yield return new PptxTextFlowSegment(c.ToString(), c.ToString(), PptxTextFlowSegmentKind.BoundaryPunctuation, Draw: true, PreventCoalesce: true);
                nextPreventsCoalesce = true;
                hideLeadingSpaces = true;
                continue;
            }

            builder.Append(c);
        }

        if (builder.Length > 0)
        {
            yield return new PptxTextFlowSegment(builder.ToString(), builder.ToString(), PptxTextFlowSegmentKind.Text, Draw: true, PreventCoalesce: nextPreventsCoalesce);
        }
    }

    private static bool IsOfficeTextOperationBoundaryPunctuation(char value)
    {
        return CharUnicodeInfo.GetUnicodeCategory(value) == UnicodeCategory.DashPunctuation;
    }

    private static double MeasureFlowSegmentBoundaryAdjustment(TextAdvanceEstimator advanceEstimator, string advanceText, int? previousCodePoint, double fontSize, string? typeface, bool bold, bool italic, double characterSpacing, bool kerningEnabled)
    {
        if (previousCodePoint is int previous && FirstCodePoint(advanceText) is int first)
        {
            return advanceEstimator.MeasureBoundaryAdvance(previous, first, fontSize, typeface, bold, italic, characterSpacing, kerningEnabled);
        }

        return 0d;
    }

    private static int? FirstCodePoint(string text)
    {
        foreach (Rune rune in text.EnumerateRunes())
        {
            return rune.Value;
        }

        return null;
    }

    private static int? LastCodePoint(string text)
    {
        int? last = null;
        foreach (Rune rune in text.EnumerateRunes())
        {
            last = rune.Value;
        }

        return last;
    }

    private static IReadOnlyList<PptxTextAtomLayout> BuildTextAtoms(TextRun run, TextAdvanceEstimator advanceEstimator, PptxTextAtomKind? forcedKind = null)
    {
        if (run.Text.Length == 0)
        {
            return [];
        }

        if (forcedKind is { } kind)
        {
            return [new PptxTextAtomLayout(kind, run.Text, run.X, run.Width, Draw: kind != PptxTextAtomKind.HiddenAdvance)];
        }

        var atoms = new List<PptxTextAtomLayout>();
        double cursorX = run.X;
        int index = 0;
        while (index < run.Text.Length)
        {
            int start = index;
            bool isSpace = run.Text[index] == ' ';
            while (index < run.Text.Length && (run.Text[index] == ' ') == isSpace)
            {
                index++;
            }

            string text = run.Text[start..index];
            double width = advanceEstimator.Measure(text, run.FontSize, run.FontFamily, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled);
            atoms.Add(new PptxTextAtomLayout(isSpace ? PptxTextAtomKind.Space : PptxTextAtomKind.Word, text, cursorX, width, Draw: true));
            cursorX += width;
        }

        if (atoms.Count > 0)
        {
            PptxTextAtomLayout last = atoms[^1];
            double delta = run.X + run.Width - (last.X + last.Width);
            if (Math.Abs(delta) > PptxTextMetricRules.TextStateTolerance)
            {
                atoms[^1] = last with { Width = Math.Max(0d, last.Width + delta) };
            }
        }

        return atoms;
    }

    private static PptxTextGlyphSpanLayout BuildGlyphSpan(TextRun run, TextAdvanceEstimator advanceEstimator, double leadingAdjustment = 0d)
    {
        var glyphs = new List<PptxTextGlyphLayout>();
        OpenTypeFont? previousFont = null;
        ushort previousGlyph = 0;
        foreach (Rune rune in run.Text.EnumerateRunes())
        {
            ResolvedGlyphFont? resolved = advanceEstimator.ResolveGlyphFont(run.FontFamily, run.Bold, run.Italic, rune.Value);
            if (resolved is null || resolved.Font.UnitsPerEm == 0)
            {
                continue;
            }

            OpenTypeFont font = resolved.Font;
            ushort glyph = font.MapCodePoint(rune.Value);
            if (glyph == 0)
            {
                continue;
            }

            double adjustmentBefore = 0d;
            if (glyphs.Count > 0)
            {
                adjustmentBefore += run.CharacterSpacing;
                if (run.KerningEnabled && previousFont == font && previousGlyph != 0)
                {
                    adjustmentBefore += font.GetKerning(previousGlyph, glyph) * run.FontSize / font.UnitsPerEm;
                }

                if (previousFont == font && previousGlyph != 0 && resolved.SyntheticBold)
                {
                    adjustmentBefore -= PptxTextMetricRules.OfficeSyntheticBoldAdvanceTightening(run.FontSize);
                }
            }

            double advance = font.GetAdvanceWidth(glyph) * run.FontSize / font.UnitsPerEm;
            glyphs.Add(new PptxTextGlyphLayout(rune.Value, resolved.Typeface, resolved.Source, glyph, advance, adjustmentBefore));
            previousFont = font;
            previousGlyph = glyph;
        }

        if (glyphs.Count == 0)
        {
            return PptxTextGlyphSpanLayout.Empty(run);
        }

        double naturalWidth = glyphs.Sum(glyph => glyph.Advance) + glyphs.Sum(glyph => glyph.AdjustmentBefore);
        return new PptxTextGlyphSpanLayout(
            run.Text,
            run.FontFamily,
            run.Bold,
            run.Italic,
            run.FontSize,
            run.CharacterSpacing,
            run.KerningEnabled,
            leadingAdjustment,
            Math.Max(0d, naturalWidth),
            run.Width,
            glyphs);
    }

    private static PptxTextLineBoxLayout CreateLineBox(
        double lineTopY,
        double baselineY,
        LineSpacing lineSpacing,
        double maxFontSize,
        TextLayoutLine line,
        TextAdvanceEstimator advanceEstimator,
        bool useOfficeBaselineFloor)
    {
        double advance = ReadLineAdvance(lineSpacing, maxFontSize);
        PptxTextLineMetrics lineMetrics = ResolvePositionedLineMetrics(lineTopY, baselineY, advance);
        PptxTextSpanLayout? baselineSpan = line.Spans.FirstOrDefault();
        ResolvedRunTextStyle? baselineStyle = baselineSpan?.SourceRun?.Style;
        double baselineFontSize = baselineSpan?.Run.FontSize ?? maxFontSize;
        PptxTextBaselineMetricLayout baselineMetric = ReadBaselineMetric(baselineFontSize, baselineStyle, advanceEstimator, useOfficeBaselineFloor);
        return new PptxTextLineBoxLayout(
            lineTopY,
            baselineY,
            lineMetrics.LineAdvance,
            lineMetrics.BaselineOffset,
            maxFontSize,
            lineSpacing,
            baselineMetric);
    }

    private static PptxTextLineMetrics ResolvePositionedLineMetrics(double lineTopY, double baselineY, double lineAdvance)
    {
        return new PptxTextLineMetrics(lineTopY - baselineY, lineAdvance, "positioned-line");
    }

    private static void AddAlignedParagraphLine(
        List<PptxTextLineLayout> lines,
        TextLayoutLine line,
        PptxTextLineBoxLayout box,
        TextAlignment alignment,
        double textX,
        double textWidth,
        bool justify,
        bool distribute,
        TextAdvanceEstimator advanceEstimator)
    {
        if (line.Spans.Count == 0)
        {
            return;
        }

        double alignmentEndX = ReadAlignmentEndX(line);
        double paragraphWidth = Math.Max(0d, alignmentEndX - textX);
        bool justifyLine = justify && paragraphWidth > 0d && paragraphWidth < textWidth;
        double offset = alignment switch
        {
            TextAlignment.Center => Math.Max(0d, textWidth - paragraphWidth) / 2d,
            TextAlignment.Right => Math.Max(0d, textWidth - paragraphWidth),
            _ => 0d
        };

        if (justifyLine)
        {
            PptxTextLineLayout? justified = TryJustifyLine(line, box, textX, textWidth, advanceEstimator);
            if (justified is not null)
            {
                lines.Add(justified);
                return;
            }
        }

        if (distribute && paragraphWidth > 0d && paragraphWidth < textWidth)
        {
            PptxTextLineLayout? distributed = TryDistributeLine(line, box, textX, textWidth);
            if (distributed is not null)
            {
                lines.Add(distributed);
                return;
            }
        }

        IEnumerable<PptxTextSpanLayout> spans = line.Spans
            .Select(span => span with
            {
                Run = span.Run with
                {
                    X = span.Run.X + offset,
                    Width = span.Run.Width,
                    Alignment = TextAlignment.Left
                },
                EndX = span.EndX + offset,
                Atoms = OffsetAtoms(span.Atoms, offset)
            });
        if (IsWordJustifiedAlignment(alignment))
        {
            spans = spans.SelectMany(span => SplitJustifiedWordSpans(span, advanceEstimator));
        }

        lines.Add(new PptxTextLineLayout(box, textX + offset, line.EndX + offset, line.EndX + offset, alignment, spans.ToArray()));
    }

    private static double ReadAlignmentEndX(TextLayoutLine line)
    {
        for (int spanIndex = line.Spans.Count - 1; spanIndex >= 0; spanIndex--)
        {
            PptxTextSpanLayout span = line.Spans[spanIndex];
            for (int atomIndex = span.Atoms.Count - 1; atomIndex >= 0; atomIndex--)
            {
                PptxTextAtomLayout atom = span.Atoms[atomIndex];
                if (!atom.Draw || atom.Kind == PptxTextAtomKind.HiddenAdvance || atom.Kind == PptxTextAtomKind.Space)
                {
                    continue;
                }

                return atom.X + atom.Width;
            }
        }

        return line.EndX;
    }

    private static IReadOnlyList<PptxTextAtomLayout> OffsetAtoms(IReadOnlyList<PptxTextAtomLayout> atoms, double offset)
    {
        if (Math.Abs(offset) <= PptxTextMetricRules.TextStateTolerance)
        {
            return atoms;
        }

        return atoms.Select(atom => atom with { X = atom.X + offset }).ToArray();
    }

    private static bool IsWordJustifiedAlignment(TextAlignment alignment)
    {
        return alignment is TextAlignment.Justify or TextAlignment.JustLow or TextAlignment.ThaiDistributed;
    }

    private static PptxTextLineLayout? TryJustifyLine(TextLayoutLine line, PptxTextLineBoxLayout box, double textX, double textWidth, TextAdvanceEstimator advanceEstimator)
    {
        double drawableEndX = ReadAlignmentEndX(line);
        int spaceCount = CountStretchableJustificationSpaces(line, drawableEndX);
        if (spaceCount == 0)
        {
            return null;
        }

        double extraWidth = textWidth - Math.Max(0d, drawableEndX - textX);
        if (extraWidth <= PptxTextMetricRules.TextStateTolerance)
        {
            return null;
        }

        double extraPerSpace = extraWidth / spaceCount;
        double shift = 0d;
        var spans = new List<PptxTextSpanLayout>(line.Spans.Count);
        foreach (PptxTextSpanLayout span in line.Spans)
        {
            int spanSpaces = span.Run.Text.Count(static c => c == ' ');
            double spanExtra = spanSpaces * extraPerSpace;
            PptxTextSpanLayout justifiedSpan = span with
            {
                Run = span.Run with
                {
                    X = span.Run.X + shift,
                    Width = span.Run.Width + spanExtra,
                    Alignment = TextAlignment.Left,
                    PreventCoalesce = true
                },
                EndX = span.EndX + shift + spanExtra,
                Atoms = JustifyAtoms(span.Atoms, shift, extraPerSpace),
                GlyphSpan = span.GlyphSpan with { LayoutWidth = span.GlyphSpan.LayoutWidth + spanExtra }
            };
            foreach (PptxTextSpanLayout wordSpan in SplitJustifiedWordSpans(justifiedSpan, advanceEstimator))
            {
                spans.Add(wordSpan);
            }

            shift += spanExtra;
        }

        return new PptxTextLineLayout(box, textX, textX + textWidth, line.EndX, TextAlignment.Justify, spans);
    }

    private static int CountStretchableJustificationSpaces(TextLayoutLine line, double drawableEndX)
    {
        int count = 0;
        foreach (PptxTextSpanLayout span in line.Spans)
        {
            foreach (PptxTextAtomLayout atom in span.Atoms)
            {
                if (!atom.Draw ||
                    atom.Kind != PptxTextAtomKind.Space ||
                    atom.X >= drawableEndX - PptxTextMetricRules.TextStateTolerance)
                {
                    continue;
                }

                count += atom.Text.Count(static c => c == ' ');
            }
        }

        return count;
    }

    private static PptxTextLineLayout? TryDistributeLine(TextLayoutLine line, PptxTextLineBoxLayout box, double textX, double textWidth)
    {
        int glyphCount = line.Spans.Sum(span => span.GlyphSpan.Glyphs.Count);
        if (glyphCount <= 1)
        {
            return null;
        }

        double extraWidth = textWidth - Math.Max(0d, line.EndX - textX);
        if (extraWidth <= PptxTextMetricRules.TextStateTolerance)
        {
            return null;
        }

        double extraPerGlyphGap = extraWidth / (glyphCount - 1);
        double cursor = textX;
        int globalGlyphIndex = 0;
        var spans = new List<PptxTextSpanLayout>(glyphCount);
        foreach (PptxTextSpanLayout span in line.Spans)
        {
            TextElementEnumerator elements = StringInfo.GetTextElementEnumerator(span.Run.Text);
            foreach (PptxTextGlyphLayout glyph in span.GlyphSpan.Glyphs)
            {
                if (!elements.MoveNext())
                {
                    return null;
                }

                string text = elements.GetTextElement();
                cursor += glyph.AdjustmentBefore;
                TextRun glyphRun = span.Run with
                {
                    Text = text,
                    X = cursor,
                    Width = glyph.Advance,
                    CharacterSpacing = 0d,
                    Alignment = TextAlignment.Left,
                    PreventCoalesce = true
                };
                var glyphSpan = new PptxTextGlyphSpanLayout(
                    text,
                    span.GlyphSpan.Typeface,
                    span.GlyphSpan.Bold,
                    span.GlyphSpan.Italic,
                    span.GlyphSpan.FontSize,
                    0d,
                    span.GlyphSpan.KerningEnabled,
                    0d,
                    glyph.Advance,
                    glyph.Advance,
                    [glyph with { AdjustmentBefore = 0d }]);
                spans.Add(span with
                {
                    Run = glyphRun,
                    EndX = cursor + glyph.Advance,
                    Atoms = [new PptxTextAtomLayout(char.IsWhiteSpace(text, 0) ? PptxTextAtomKind.Space : PptxTextAtomKind.Word, text, cursor, glyph.Advance, Draw: true)],
                    GlyphSpan = glyphSpan
                });

                cursor += glyph.Advance;
                if (++globalGlyphIndex < glyphCount)
                {
                    cursor += extraPerGlyphGap;
                }
            }
        }

        return new PptxTextLineLayout(box, textX, textX + textWidth, line.EndX, TextAlignment.Distributed, spans);
    }

    private static IEnumerable<PptxTextSpanLayout> SplitJustifiedWordSpans(PptxTextSpanLayout span, TextAdvanceEstimator advanceEstimator)
    {
        PptxTextAtomLayout[] words = span.Atoms
            .Where(static atom => atom.Kind == PptxTextAtomKind.Word && atom.Draw && atom.Text.Length != 0)
            .ToArray();
        if (words.Length == 0)
        {
            if (span.Atoms.Any(static atom => atom.Kind is not PptxTextAtomKind.Space and not PptxTextAtomKind.HiddenAdvance))
            {
                yield return span;
            }

            yield break;
        }

        foreach (PptxTextAtomLayout word in words
                     .SelectMany(word => SplitWordAtomOnSpaces(span.Run, word, advanceEstimator))
                     .SelectMany(word => SplitJustifiedWordAtomOnSentencePeriod(span.Run, word, advanceEstimator)))
        {
            TextRun wordRun = span.Run with
            {
                Text = word.Text,
                X = word.X,
                Width = word.Width,
                PreventCoalesce = true
            };
            yield return span with
            {
                Run = wordRun,
                EndX = word.X + word.Width,
                Atoms = [word],
                GlyphSpan = BuildGlyphSpan(wordRun, advanceEstimator)
            };
        }
    }

    private static IEnumerable<PptxTextAtomLayout> SplitJustifiedWordAtomOnSentencePeriod(TextRun run, PptxTextAtomLayout atom, TextAdvanceEstimator advanceEstimator)
    {
        if (atom.Text.Length <= 1 || atom.Text[^1] != '.')
        {
            yield return atom;
            yield break;
        }

        string wordText = atom.Text[..^1];
        double wordWidth = advanceEstimator.Measure(wordText, run.FontSize, run.FontFamily, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled);
        double periodBoundaryAdjustment = MeasureFlowSegmentBoundaryAdjustment(
            advanceEstimator,
            ".",
            LastCodePoint(wordText),
            run.FontSize,
            run.FontFamily,
            run.Bold,
            run.Italic,
            run.CharacterSpacing,
            run.KerningEnabled);
        double periodX = atom.X + wordWidth + periodBoundaryAdjustment;
        yield return atom with { Text = wordText, Width = Math.Max(0d, periodX - atom.X) };
        yield return atom with { Text = ".", X = periodX, Width = Math.Max(0d, atom.X + atom.Width - periodX) };
    }

    private static IEnumerable<PptxTextAtomLayout> SplitWordAtomOnSpaces(TextRun run, PptxTextAtomLayout atom, TextAdvanceEstimator advanceEstimator)
    {
        if (atom.Text.IndexOf(' ') < 0)
        {
            yield return atom;
            yield break;
        }

        double cursorX = atom.X;
        int index = 0;
        while (index < atom.Text.Length)
        {
            int start = index;
            bool isSpace = atom.Text[index] == ' ';
            while (index < atom.Text.Length && (atom.Text[index] == ' ') == isSpace)
            {
                index++;
            }

            string text = atom.Text[start..index];
            double width = advanceEstimator.Measure(text, run.FontSize, run.FontFamily, run.Bold, run.Italic, run.CharacterSpacing, run.KerningEnabled);
            if (!isSpace)
            {
                yield return atom with
                {
                    Kind = PptxTextAtomKind.Word,
                    Text = text,
                    X = cursorX,
                    Width = width,
                    Draw = true
                };
            }

            cursorX += width;
        }
    }

    private static IReadOnlyList<PptxTextAtomLayout> JustifyAtoms(IReadOnlyList<PptxTextAtomLayout> atoms, double initialShift, double extraPerSpace)
    {
        double shift = initialShift;
        var justified = new PptxTextAtomLayout[atoms.Count];
        for (int i = 0; i < atoms.Count; i++)
        {
            PptxTextAtomLayout atom = atoms[i];
            double extra = atom.Kind == PptxTextAtomKind.Space
                ? atom.Text.Count(static c => c == ' ') * extraPerSpace
                : 0d;
            justified[i] = atom with
            {
                X = atom.X + shift,
                Width = atom.Width + extra
            };
            shift += extra;
        }

        return justified;
    }

    private static XElement? MergeParagraphProperties(params XElement?[] sources)
    {
        return PptxParagraphPropertyMerger.MergeRendererDefaultProperties(DrawingNamespace + "defRPr", sources);
    }

    private static ResolvedParagraphTextStyle ResolveParagraphTextStyle(
        XElement paragraph,
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        double fontScale,
        double lineSpacingScale,
        bool compatibleLineSpacing,
        double compatibleDefaultLineSpacingFactor)
    {
        XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
        double fontSize = ReadFirstParagraphFontSize(paragraph, defaultRunProperties) * fontScale;
        string? alignmentValue = ReadAlignmentValue(paragraph, defaultParagraphProperties);
        return new ResolvedParagraphTextStyle(
            ParseAlignment(alignmentValue),
            alignmentValue,
            paragraphProperties,
            defaultRunProperties,
            fontSize,
            ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", fontSize),
            ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", fontSize),
            ApplyCompatibleLineSpacing(
                ReadLineSpacing(paragraphProperties, defaultParagraphProperties),
                compatibleLineSpacing && HasManualLineBreak(paragraph) && !HasExplicitParagraphSpacing(paragraphProperties),
                compatibleDefaultLineSpacingFactor).ScaleExplicit(lineSpacingScale),
            ReadParagraphIndent(paragraphProperties, defaultParagraphProperties),
            ReadTabStops(paragraphProperties, defaultParagraphProperties));
    }

    private static bool HasManualLineBreak(XElement paragraph)
    {
        return paragraph.Elements(DrawingNamespace + "br").Any() ||
            paragraph.Elements(DrawingNamespace + "r")
                .Elements(DrawingNamespace + "t")
                .Any(text => TextContainsManualLineBreak(text.Value));
    }

    private static bool HasExplicitParagraphSpacing(XElement? paragraphProperties)
    {
        return paragraphProperties?.Element(DrawingNamespace + "spcBef") is not null ||
            paragraphProperties?.Element(DrawingNamespace + "spcAft") is not null;
    }

    private static ResolvedRunTextStyle ResolveRunTextStyle(
        PptxRunStyleCascade cascade,
        RgbColor? shapeFontColor,
        PptxTheme theme,
        PptxColorMap colorMap,
        double fontScale,
        PptxSceneTableCellTextStyle tableStyleTextStyle = default)
    {
        return ResolveRunTextStyle(
            cascade.DirectProperties,
            cascade.ResolvedDefaultProperties,
            shapeFontColor,
            theme,
            colorMap,
            fontScale,
            tableStyleTextStyle);
    }

    private static ResolvedRunTextStyle ResolveRunTextStyle(
        XElement? runProperties,
        XElement? defaultRunProperties,
        RgbColor? shapeFontColor,
        PptxTheme theme,
        PptxColorMap colorMap,
        double fontScale,
        PptxSceneTableCellTextStyle tableStyleTextStyle = default)
    {
        double nominalFontSize = ReadFontSize(runProperties, defaultRunProperties) * fontScale;
        double baselineOffset = ReadBaselineOffset(runProperties, defaultRunProperties, nominalFontSize);
        double fontSize = PptxTextMetricRules.ShouldScaleSuperscriptSubscript(baselineOffset, nominalFontSize)
            ? PptxTextMetricRules.SuperscriptSubscriptFontSize(nominalFontSize)
            : nominalFontSize;
        double alpha = 1d;
        RgbColor color;
        PptxRunTextColorSource colorSource;
        if (HasTextNoFill(runProperties))
        {
            color = new RgbColor(0, 0, 0);
            colorSource = PptxRunTextColorSource.RunNoFill;
            alpha = 0d;
        }
        else if (HasHyperlinkClick(runProperties) && theme.TryResolveColor("hlink", colorMap, out RgbColor hyperlinkColor))
        {
            color = hyperlinkColor;
            colorSource = PptxRunTextColorSource.ThemeHyperlink;
        }
        else if (TryReadSolidColorWithAlpha(runProperties, theme, colorMap, out RgbColor runColor, out double runAlpha))
        {
            color = runColor;
            colorSource = PptxRunTextColorSource.RunSolidFill;
            alpha = runAlpha;
        }
        else if (tableStyleTextStyle.Color is { } tableTextColor && !HasTextFill(runProperties))
        {
            color = tableTextColor;
            colorSource = PptxRunTextColorSource.TableTextStyle;
        }
        else if (shapeFontColor is { } fontRefColor)
        {
            color = fontRefColor;
            colorSource = PptxRunTextColorSource.ShapeFontRef;
        }
        else if (HasTextNoFill(defaultRunProperties))
        {
            color = new RgbColor(0, 0, 0);
            colorSource = PptxRunTextColorSource.DefaultNoFill;
            alpha = 0d;
        }
        else if (TryReadSolidColorWithAlpha(defaultRunProperties, theme, colorMap, out RgbColor defaultColor, out double defaultAlpha))
        {
            color = defaultColor;
            colorSource = PptxRunTextColorSource.DefaultSolidFill;
            alpha = defaultAlpha;
        }
        else
        {
            color = new RgbColor(0, 0, 0);
            colorSource = PptxRunTextColorSource.FallbackBlack;
        }

        PptxThemeTypefaceResolution typeface = ReadRunTypeface(runProperties, defaultRunProperties, theme);
        bool bold = ParseOptionalBoolAttribute(runProperties, "b") ||
            (runProperties?.Attribute("b") is null && tableStyleTextStyle.Bold) ||
            (runProperties?.Attribute("b") is null && ParseOptionalBoolAttribute(defaultRunProperties, "b"));
        bool italic = ParseOptionalBoolAttribute(runProperties, "i") ||
            (runProperties?.Attribute("i") is null && ParseOptionalBoolAttribute(defaultRunProperties, "i"));
        bool hasHyperlinkClick = HasHyperlinkClick(runProperties);
        string? underlineValue = ReadUnderlineValue(runProperties, defaultRunProperties);
        string? strikeValue = ReadStrikeValue(runProperties, defaultRunProperties);
        string? capsValue = ReadTextCapsValue(runProperties, defaultRunProperties);
        bool underline = underlineValue is null
            ? hasHyperlinkClick
            : !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);

        return new ResolvedRunTextStyle(
            nominalFontSize,
            fontSize,
            ReadCharacterSpacing(runProperties, defaultRunProperties),
            baselineOffset,
            color,
            colorSource,
            alpha,
            TryReadTextOutline(runProperties, defaultRunProperties, theme, colorMap, out TextOutline outline) ? outline : null,
            TryReadHighlightColor(runProperties, out RgbColor highlightColor) ? highlightColor : null,
            hasHyperlinkClick,
            ReadHyperlinkClickId(runProperties),
            bold,
            italic,
            underline,
            underlineValue ?? (hasHyperlinkClick ? "sng" : null),
            IsStrikeEnabled(strikeValue),
            strikeValue,
            capsValue,
            IsKerningEnabled(runProperties, defaultRunProperties, fontSize),
            typeface.Source,
            typeface.Typeface);
    }

    private static bool HasTextNoFill(XElement? runProperties)
    {
        return runProperties?.Element(DrawingNamespace + "noFill") is not null;
    }

    private static bool HasTextFill(XElement? runProperties)
    {
        return runProperties?.Element(DrawingNamespace + "solidFill") is not null ||
            runProperties?.Element(DrawingNamespace + "noFill") is not null ||
            runProperties?.Element(DrawingNamespace + "gradFill") is not null;
    }

    private static bool TryReadTextOutline(XElement? runProperties, XElement? defaultRunProperties, PptxTheme theme, PptxColorMap colorMap, out TextOutline outline)
    {
        return TryReadTextOutline(runProperties, theme, colorMap, out outline) ||
            TryReadTextOutline(defaultRunProperties, theme, colorMap, out outline);
    }

    private static bool TryReadTextOutline(XElement? runProperties, PptxTheme theme, PptxColorMap colorMap, out TextOutline outline)
    {
        outline = default;
        XElement? line = runProperties?.Element(DrawingNamespace + "ln");
        if (line is null || line.Element(DrawingNamespace + "noFill") is not null)
        {
            return false;
        }

        double? width = line.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : null;
        if (!TryReadSolidColorWithAlpha(line, theme, colorMap, out RgbColor color, out double alpha))
        {
            return false;
        }

        outline = new TextOutline(color, alpha, PptxTextMetricRules.TextOutlineWidth(width));
        return true;
    }

    private static bool HasHyperlinkClick(XElement? runProperties)
    {
        return runProperties?.Element(DrawingNamespace + "hlinkClick") is not null;
    }

    private static string? ReadHyperlinkClickId(XElement? runProperties)
    {
        return (string?)runProperties
            ?.Element(DrawingNamespace + "hlinkClick")
            ?.Attribute(RelationshipsNamespace + "id");
    }

    private static PptxThemeTypefaceResolution ReadRunTypeface(XElement? runProperties, XElement? defaultRunProperties, PptxTheme theme)
    {
        return theme.ResolveTypefaceWithSource(ReadTypeface(runProperties) ?? ReadTypeface(defaultRunProperties));
    }

    private static string? ReadTypeface(XElement? runProperties)
    {
        return (string?)(runProperties?.Element(DrawingNamespace + "latin") ??
            runProperties?.Element(DrawingNamespace + "ea") ??
            runProperties?.Element(DrawingNamespace + "cs"))
            ?.Attribute("typeface");
    }

    private static bool IsKerningEnabled(XElement? runProperties, XElement? defaultRunProperties, double fontSize)
    {
        XAttribute? threshold = runProperties?.Attribute("kern") ?? defaultRunProperties?.Attribute("kern");
        if (threshold is null)
        {
            return false;
        }

        double minimumFontSize = int.Parse(threshold.Value, CultureInfo.InvariantCulture) / 100d;
        return minimumFontSize <= 0d || fontSize >= minimumFontSize;
    }

    private static bool IsTextRunElement(XElement element)
    {
        return element.Name == DrawingNamespace + "r" ||
            element.Name == DrawingNamespace + "fld";
    }

    private static string ReadTextElementText(XElement element, int slideNumber)
    {
        if (slideNumber > 0 &&
            element.Name == DrawingNamespace + "fld" &&
            string.Equals((string?)element.Attribute("type"), "slidenum", StringComparison.OrdinalIgnoreCase))
        {
            return slideNumber.ToString(CultureInfo.InvariantCulture);
        }

        return NormalizeText((string?)element.Element(DrawingNamespace + "t") ?? string.Empty);
    }

    private static string NormalizeText(string text)
    {
        return text;
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
        return IsStrikeEnabled(ReadStrikeValue(runProperties, defaultRunProperties));
    }

    private static bool IsStrikeEnabled(string? value)
    {
        return value is not null && !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static string? ReadUnderlineValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("u") ?? defaultRunProperties?.Attribute("u"));
    }

    private static string? ReadStrikeValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("strike") ?? defaultRunProperties?.Attribute("strike"));
    }

    private static string? ReadTextCapsValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("cap") ?? defaultRunProperties?.Attribute("cap"));
    }

    private static IReadOnlyList<TextCapsFragment> ApplyTextCaps(string text, XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = ReadTextCapsValue(runProperties, defaultRunProperties);
        if (text.Length == 0)
        {
            return [];
        }

        if (value is "all")
        {
            return [new TextCapsFragment(text.ToUpperInvariant(), 1d)];
        }

        if (value is not "small")
        {
            return [new TextCapsFragment(text, 1d)];
        }

        var fragments = new List<TextCapsFragment>();
        var builder = new StringBuilder();
        bool? currentSmall = null;
        foreach (char character in text)
        {
            bool isSmall = char.IsLetter(character) && char.IsLower(character);
            if (currentSmall is not null && currentSmall != isSmall)
            {
                fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? PptxTextMetricRules.SmallCapsFontScale() : 1d));
                builder.Clear();
            }

            currentSmall = isSmall;
            builder.Append(char.ToUpperInvariant(character));
        }

        if (builder.Length > 0 && currentSmall is not null)
        {
            fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? PptxTextMetricRules.SmallCapsFontScale() : 1d));
        }

        return fragments;
    }

    private static (TextInsets Insets, TextInsetSources Sources, TextInsetValues Values) ReadTextInsets(
        XElement textBody,
        XElement? inheritedTextBody)
    {
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        XElement? inheritedBodyProperties = inheritedTextBody?.Element(DrawingNamespace + "bodyPr");
        (double left, PptxTextBodyPropertySource leftSource, string? leftValue) = ReadInset(
            bodyProperties, inheritedBodyProperties, "lIns", 91440);
        (double right, PptxTextBodyPropertySource rightSource, string? rightValue) = ReadInset(
            bodyProperties, inheritedBodyProperties, "rIns", 91440);
        (double top, PptxTextBodyPropertySource topSource, string? topValue) = ReadInset(
            bodyProperties, inheritedBodyProperties, "tIns", 45720);
        (double bottom, PptxTextBodyPropertySource bottomSource, string? bottomValue) = ReadInset(
            bodyProperties, inheritedBodyProperties, "bIns", 45720);

        return (
            new TextInsets(left, right, top, bottom),
            new TextInsetSources(leftSource, rightSource, topSource, bottomSource),
            new TextInsetValues(leftValue, rightValue, topValue, bottomValue));
    }

    private static PptxTextOrientation ParseTextOrientation(string? orientation)
    {
        return orientation switch
        {
            null or "" => PptxTextOrientation.Horizontal,
            "horz" => PptxTextOrientation.Horizontal,
            "vert" => PptxTextOrientation.Vertical,
            "vert270" => PptxTextOrientation.Vertical270,
            "eaVert" => PptxTextOrientation.EastAsianVertical,
            "mongolianVert" => PptxTextOrientation.MongolianVertical,
            "wordArtVert" => PptxTextOrientation.WordArtVertical,
            "wordArtVertRtl" => PptxTextOrientation.WordArtVerticalRightToLeft,
            _ when orientation.Equals("horz", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.Horizontal,
            _ when orientation.Equals("vert", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.Vertical,
            _ when orientation.Equals("vert270", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.Vertical270,
            _ when orientation.Equals("eaVert", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.EastAsianVertical,
            _ when orientation.Equals("mongolianVert", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.MongolianVertical,
            _ when orientation.Equals("wordArtVert", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.WordArtVertical,
            _ when orientation.Equals("wordArtVertRtl", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.WordArtVerticalRightToLeft,
            _ => PptxTextOrientation.Unknown
        };
    }

    private static double TextOrientationRotationDegrees(PptxTextOrientation orientation)
    {
        return orientation switch
        {
            PptxTextOrientation.Vertical270 => 270d,
            PptxTextOrientation.Vertical or
            PptxTextOrientation.EastAsianVertical or
            PptxTextOrientation.MongolianVertical or
            PptxTextOrientation.WordArtVertical or
            PptxTextOrientation.WordArtVerticalRightToLeft => 90d,
            _ => 0d
        };
    }

    private static double? ParseTextBodyRotationDegrees(string? rotation)
    {
        return rotation is not null &&
            long.TryParse(rotation, NumberStyles.Integer, CultureInfo.InvariantCulture, out long rotationValue)
                ? rotationValue / 60000d
                : null;
    }

    private static (
        int Count,
        double Spacing,
        PptxTextBodyPropertySource CountSource,
        PptxTextBodyPropertySource SpacingSource,
        string? CountValue,
        string? SpacingValue) ReadTextColumns(XElement textBody, XElement? inheritedTextBody)
    {
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        XElement? inheritedBodyProperties = inheritedTextBody?.Element(DrawingNamespace + "bodyPr");
        (int count, PptxTextBodyPropertySource countSource, string? countValue) = ReadTextColumnCount(bodyProperties, inheritedBodyProperties);
        (double spacing, PptxTextBodyPropertySource spacingSource, string? spacingValue) = ReadTextColumnSpacing(bodyProperties, inheritedBodyProperties);
        return (count, spacing, countSource, spacingSource, countValue, spacingValue);
    }

    private static (int Count, PptxTextBodyPropertySource Source, string? Value) ReadTextColumnCount(
        XElement? bodyProperties,
        XElement? inheritedBodyProperties)
    {
        if (bodyProperties?.Attribute("numCol") is { } directAttribute)
        {
            if (int.TryParse(directAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int directCount))
            {
                return (Math.Clamp(directCount, 1, 16), PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
            }

            return (1, PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
        }

        if (inheritedBodyProperties?.Attribute("numCol") is { } inheritedAttribute)
        {
            if (int.TryParse(inheritedAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int inheritedCount))
            {
                return (Math.Clamp(inheritedCount, 1, 16), PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
            }

            return (1, PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
        }

        return (1, PptxTextBodyPropertySource.DefaultValue, null);
    }

    private static (double Spacing, PptxTextBodyPropertySource Source, string? Value) ReadTextColumnSpacing(
        XElement? bodyProperties,
        XElement? inheritedBodyProperties)
    {
        if (bodyProperties?.Attribute("spcCol") is { } directAttribute)
        {
            if (long.TryParse(directAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long directSpacing))
            {
                return (Math.Max(0d, OoxUnits.EmuToPoints(directSpacing)), PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
            }

            return (0d, PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
        }

        if (inheritedBodyProperties?.Attribute("spcCol") is { } inheritedAttribute)
        {
            if (long.TryParse(inheritedAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long inheritedSpacing))
            {
                return (Math.Max(0d, OoxUnits.EmuToPoints(inheritedSpacing)), PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
            }

            return (0d, PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
        }

        return (0d, PptxTextBodyPropertySource.DefaultValue, null);
    }

    private static double NormalizeRotationDegrees(double rotationDegrees)
    {
        double normalized = rotationDegrees % 360d;
        return normalized < 0d ? normalized + 360d : normalized;
    }

    private static (double Scale, PptxTextBodyPropertySource Source, string? Value) ReadNormAutofitFontScale(
        XElement? autofit,
        PptxTextBodyPropertySource autofitSource)
    {
        if (autofit?.Name.LocalName != "normAutofit" ||
            autofit.Attribute("fontScale") is not { } fontScale)
        {
            return (1d, PptxTextBodyPropertySource.DefaultValue, null);
        }

        if (!int.TryParse(fontScale.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fontScaleValue))
        {
            return (1d, autofitSource, fontScale.Value);
        }

        double scale = Math.Clamp(
            fontScaleValue / 100000d,
            PptxTextMetricRules.MinimumAutofitScale,
            PptxTextMetricRules.MaximumAutofitScale);
        return (scale, autofitSource, fontScale.Value);
    }

    private static (double Scale, PptxTextBodyPropertySource Source, string? Value) ReadNormAutofitLineSpacingScale(
        XElement? autofit,
        PptxTextBodyPropertySource autofitSource)
    {
        if (autofit?.Name.LocalName != "normAutofit" ||
            autofit.Attribute("lnSpcReduction") is not { } reduction)
        {
            return (1d, PptxTextBodyPropertySource.DefaultValue, null);
        }

        if (!int.TryParse(reduction.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int reductionValue))
        {
            return (1d, autofitSource, reduction.Value);
        }

        double reductionRatio = Math.Clamp(reductionValue / 100000d, 0d, PptxTextMetricRules.MaximumLineSpacingReduction);
        return (1d - reductionRatio, autofitSource, reduction.Value);
    }

    private static PptxTextWrapMode ParseTextWrapMode(string? wrap)
    {
        return wrap switch
        {
            null or "" => PptxTextWrapMode.Square,
            "none" => PptxTextWrapMode.None,
            "square" => PptxTextWrapMode.Square,
            _ when wrap.Equals("none", StringComparison.OrdinalIgnoreCase) => PptxTextWrapMode.None,
            _ when wrap.Equals("square", StringComparison.OrdinalIgnoreCase) => PptxTextWrapMode.Square,
            _ => PptxTextWrapMode.Unknown
        };
    }

    private static PptxTextVerticalOverflow ParseTextVerticalOverflow(string? overflow)
    {
        return overflow switch
        {
            "clip" => PptxTextVerticalOverflow.Clip,
            "ellipsis" => PptxTextVerticalOverflow.Ellipsis,
            _ when overflow?.Equals("clip", StringComparison.OrdinalIgnoreCase) == true => PptxTextVerticalOverflow.Clip,
            _ when overflow?.Equals("ellipsis", StringComparison.OrdinalIgnoreCase) == true => PptxTextVerticalOverflow.Ellipsis,
            _ when overflow?.Equals("overflow", StringComparison.OrdinalIgnoreCase) == true => PptxTextVerticalOverflow.Overflow,
            _ when !string.IsNullOrEmpty(overflow) => PptxTextVerticalOverflow.Unknown,
            _ => PptxTextVerticalOverflow.Overflow
        };
    }

    private static (double Value, PptxTextBodyPropertySource Source, string? RawValue) ReadInset(
        XElement? bodyProperties,
        XElement? inheritedBodyProperties,
        string attributeName,
        long defaultEmu)
    {
        if (bodyProperties?.Attribute(attributeName) is { } directAttribute)
        {
            if (long.TryParse(directAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu))
            {
                return (OoxUnits.EmuToPoints(emu), PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
            }

            return (OoxUnits.EmuToPoints(defaultEmu), PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
        }

        if (inheritedBodyProperties?.Attribute(attributeName) is { } inheritedAttribute)
        {
            if (long.TryParse(inheritedAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu))
            {
                return (OoxUnits.EmuToPoints(emu), PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
            }

            return (OoxUnits.EmuToPoints(defaultEmu), PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
        }

        return (OoxUnits.EmuToPoints(defaultEmu), PptxTextBodyPropertySource.DefaultValue, null);
    }

    private static double ReadParagraphSpacing(
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        string elementName,
        double referenceFontSize)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + elementName) ??
            defaultParagraphProperties?.Element(DrawingNamespace + elementName);
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d;
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return referenceFontSize * PptxTextMetricRules.ClampNonNegative(int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d);
        }

        return 0d;
    }

    private static ParagraphIndent ReadParagraphIndent(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        return new ParagraphIndent(
            ReadParagraphEmuAttribute(paragraphProperties, defaultParagraphProperties, "marL"),
            ReadParagraphEmuAttribute(paragraphProperties, defaultParagraphProperties, "indent"));
    }

    private static double ReadParagraphEmuAttribute(XElement? paragraphProperties, XElement? defaultParagraphProperties, string attributeName)
    {
        return paragraphProperties?.Attribute(attributeName) is { } attribute
            ? OoxUnits.EmuToPoints(long.Parse(attribute.Value, CultureInfo.InvariantCulture))
            : defaultParagraphProperties?.Attribute(attributeName) is { } defaultAttribute
                ? OoxUnits.EmuToPoints(long.Parse(defaultAttribute.Value, CultureInfo.InvariantCulture))
                : 0d;
    }

    private static IReadOnlyList<double> ReadTabStops(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        XElement? tabList = paragraphProperties?.Element(DrawingNamespace + "tabLst") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "tabLst");
        if (tabList is null)
        {
            return Array.Empty<double>();
        }

        return tabList
            .Elements(DrawingNamespace + "tab")
            .Select(tab => tab.Attribute("pos") is { } position
                ? OoxUnits.EmuToPoints(long.Parse(position.Value, CultureInfo.InvariantCulture))
                : double.NaN)
            .Where(position => !double.IsNaN(position))
            .Order()
            .ToArray();
    }

    private static double ResolveNextTabX(double cursorX, double paragraphTextX, IReadOnlyList<double> tabStops)
    {
        double current = cursorX - paragraphTextX;
        foreach (double tabStop in tabStops)
        {
            if (tabStop > current + PptxTextMetricRules.CoordinateTolerance)
            {
                return paragraphTextX + tabStop;
            }
        }

        const long defaultTabStopEmus = 914400;
        double defaultTabStop = OoxUnits.EmuToPoints(defaultTabStopEmus);
        return paragraphTextX + Math.Ceiling((current + PptxTextMetricRules.CoordinateTolerance) / defaultTabStop) * defaultTabStop;
    }

    private static LineSpacing ReadLineSpacing(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + "lnSpc") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "lnSpc");
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return LineSpacing.Absolute(Math.Max(PptxTextMetricRules.MinimumLineSpacing, int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d));
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return LineSpacing.Multiple(Math.Max(PptxTextMetricRules.MinimumLineSpacing, int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d), true);
        }

        return LineSpacing.Multiple(1d, false);
    }

    private static LineSpacing ApplyCompatibleLineSpacing(LineSpacing lineSpacing, bool compatibleLineSpacing, double defaultLineSpacingFactor)
    {
        return compatibleLineSpacing && !lineSpacing.IsExplicit
            ? LineSpacing.Multiple(defaultLineSpacingFactor, isExplicit: true, useNormalLineAdvance: false)
            : lineSpacing;
    }

    private static double ReadParagraphAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return ReadLineAdvance(lineSpacing, fontSize);
    }

    private static double ReadLineAdvance(LineSpacing lineSpacing, double fontSize)
    {
        if (lineSpacing.IsAbsolute)
        {
            return lineSpacing.Resolve(fontSize);
        }

        double normalAdvance = fontSize * PptxTextMetricRules.CssNormalLineHeightFallback;
        return lineSpacing.IsExplicit
            ? (lineSpacing.UseNormalLineAdvance ? normalAdvance : fontSize) * lineSpacing.Value
            : normalAdvance;
    }

    private static double ReadFirstLineBaselineOffset(
        PptxTextParagraphModel paragraph,
        LineSpacing lineSpacing,
        TextAdvanceEstimator advanceEstimator,
        bool useOfficeBaselineFloor,
        bool shapeAutoFit,
        bool useExplicitMultipleBaselineOffset = true)
    {
        bool startsWithManualLineBreak = paragraph.Runs.FirstOrDefault()?.Kind == PptxTextRunKind.Break;
        PptxTextRunModel? firstRun = paragraph.Runs.FirstOrDefault(run => run.Kind != PptxTextRunKind.Break);
        double fontSize = firstRun?.Style.NominalFontSize ?? paragraph.FirstLineFallbackFontSize;
        return startsWithManualLineBreak || (paragraph.HasManualLineBreak && !shapeAutoFit)
            ? ManualBreakBaselineOffset(fontSize, lineSpacing, useOfficeBaselineFloor, useExplicitMultipleBaselineOffset)
            : LineBaselineOffset(fontSize, lineSpacing, firstRun?.Style, advanceEstimator, useOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
    }

    private static double ReadManualBreakLineAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return lineSpacing.IsExplicit ? ReadLineAdvance(lineSpacing, fontSize) : fontSize * PptxTextMetricRules.OfficeManualBreakDefaultLineHeightFallback;
    }

    private static double ReadFirstParagraphFontSize(XElement paragraph, XElement? defaultRunProperties)
    {
        const double defaultFontSize = 18d;
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == DrawingNamespace + "br")
            {
                return defaultFontSize;
            }

            if (!IsTextRunElement(child))
            {
                continue;
            }

            XElement? runProperties = child.Element(DrawingNamespace + "rPr");
            if (runProperties?.Attribute("sz") is { } size)
            {
                return int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d;
            }

            if (defaultRunProperties?.Attribute("sz") is { } defaultSize)
            {
                return int.Parse(defaultSize.Value, CultureInfo.InvariantCulture) / 100d;
            }

            return defaultFontSize;
        }

        return defaultFontSize;
    }

    private static double LineBaselineOffset(double fontSize, LineSpacing lineSpacing, bool useOfficeBaselineFloor, bool useExplicitMultipleBaselineOffset = true)
    {
        if (lineSpacing.IsAbsolute)
        {
            return Math.Max(BaselineOffset(fontSize), lineSpacing.Value - fontSize * PptxTextMetricRules.AbsoluteLineBaselineGapFallback);
        }

        return lineSpacing.IsExplicit && useExplicitMultipleBaselineOffset
            ? ReadExplicitMultipleBaselineOffset(lineSpacing, fontSize)
            : BaselineOffset(fontSize);
    }

    private static double LineBaselineOffset(
        double fontSize,
        LineSpacing lineSpacing,
        ResolvedRunTextStyle? style,
        TextAdvanceEstimator advanceEstimator,
        bool useOfficeBaselineFloor,
        bool useExplicitMultipleBaselineOffset = true)
    {
        if (lineSpacing.IsAbsolute)
        {
            return Math.Max(BaselineOffset(fontSize, style, advanceEstimator, useOfficeBaselineFloor: false), lineSpacing.Value - fontSize * PptxTextMetricRules.AbsoluteLineBaselineGapFallback);
        }

        return lineSpacing.IsExplicit && useExplicitMultipleBaselineOffset
            ? ReadExplicitMultipleBaselineOffset(lineSpacing, fontSize)
            : BaselineOffset(fontSize, style, advanceEstimator, useOfficeBaselineFloor);
    }

    private static double ManualBreakBaselineOffset(double fontSize, LineSpacing lineSpacing, bool useOfficeBaselineFloor, bool useExplicitMultipleBaselineOffset = true)
    {
        return lineSpacing.IsExplicit ? LineBaselineOffset(fontSize, lineSpacing, useOfficeBaselineFloor, useExplicitMultipleBaselineOffset) : fontSize * PptxTextMetricRules.OfficeManualBreakBaselineFallback;
    }

    private static bool ShouldUseExplicitMultipleBaselineOffset(PptxTextFrameModel frame, LineSpacing lineSpacing)
    {
        if (!lineSpacing.IsExplicit ||
            lineSpacing.IsAbsolute ||
            lineSpacing.Value <= 1d + PptxTextMetricRules.CoordinateTolerance)
        {
            return true;
        }

        return frame.ColumnCount <= 1 ||
            frame.BodyProperties.VerticalOverflow != PptxTextVerticalOverflow.Overflow ||
            !HasNoAutoFit(frame.BodyProperties);
    }

    private static double ReadExplicitMultipleBaselineOffset(LineSpacing lineSpacing, double fontSize)
    {
        double baseline = BaselineOffset(fontSize);
        if (lineSpacing.UseNormalLineAdvance &&
            lineSpacing.Value < 1d - PptxTextMetricRules.CoordinateTolerance)
        {
            double normalAdvance = fontSize * PptxTextMetricRules.CssNormalLineHeightFallback;
            double compressedAdvance = normalAdvance * lineSpacing.Value;
            double compressedBaseline = baseline - (normalAdvance - compressedAdvance);
            return Math.Max(fontSize * PptxTextMetricRules.MinimumBaselineMetricRatio, compressedBaseline);
        }

        return baseline * lineSpacing.Value;
    }

    private static double BaselineOffset(double fontSize)
    {
        return fontSize * PptxTextMetricRules.OfficeBaselineFallback;
    }

    private static double BaselineOffset(double fontSize, ResolvedRunTextStyle? style, TextAdvanceEstimator advanceEstimator, bool useOfficeBaselineFloor)
    {
        if (style is null)
        {
            return BaselineOffset(fontSize);
        }

        ResolvedRunTextStyle runStyle = style.Value;
        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(runStyle.Typeface, runStyle.Bold, runStyle.Italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return BaselineOffset(fontSize);
        }

        double ascenderRatio = font.Os2.WindowsAscender / (double)font.UnitsPerEm;
        if (ascenderRatio <= 0d || ascenderRatio > PptxTextMetricRules.MaximumBaselineMetricRatio)
        {
            return BaselineOffset(fontSize);
        }

        double metricRatio = ResolveOfficeBaselineMetricRatio(font, ascenderRatio, fontSize, out _);
        if (useOfficeBaselineFloor && TextMetricUsesOfficeBaselineFloor(font, runStyle, advanceEstimator, ascenderRatio))
        {
            metricRatio = Math.Max(PptxTextMetricRules.OfficeBaselineFallback, metricRatio);
        }

        return fontSize * metricRatio;
    }

    private static PptxTextBaselineMetricLayout ReadBaselineMetric(double fontSize, ResolvedRunTextStyle? style, TextAdvanceEstimator advanceEstimator, bool useOfficeBaselineFloor)
    {
        const double fallbackRatio = PptxTextMetricRules.OfficeBaselineFallback;
        if (style is null)
        {
            return new PptxTextBaselineMetricLayout("Fallback", null, false, false, fontSize, fallbackRatio, 0, 0, 0, 0, 0, 0);
        }

        ResolvedRunTextStyle runStyle = style.Value;
        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(runStyle.Typeface, runStyle.Bold, runStyle.Italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return new PptxTextBaselineMetricLayout("Fallback", runStyle.Typeface, runStyle.Bold, runStyle.Italic, fontSize, fallbackRatio, 0, 0, 0, 0, 0, 0);
        }

        double ascenderRatio = font.Os2.WindowsAscender / (double)font.UnitsPerEm;
        string source;
        double ratio;
        if (ascenderRatio > 0d && ascenderRatio <= PptxTextMetricRules.MaximumBaselineMetricRatio)
        {
            ratio = ResolveOfficeBaselineMetricRatio(font, ascenderRatio, fontSize, out source);
        }
        else
        {
            ratio = fallbackRatio;
            source = "Fallback";
        }

        if (useOfficeBaselineFloor && TextMetricUsesOfficeBaselineFloor(font, runStyle, advanceEstimator, ascenderRatio))
        {
            ratio = Math.Max(fallbackRatio, ratio);
        }
        return new PptxTextBaselineMetricLayout(
            source,
            runStyle.Typeface,
            runStyle.Bold,
            runStyle.Italic,
            fontSize,
            ratio,
            font.UnitsPerEm,
            font.Os2.WindowsAscender,
            font.Os2.WindowsDescender,
            font.Os2.TypographicAscender,
            font.Os2.TypographicDescender,
            font.Os2.TypographicLineGap);
    }

    private static double ResolveOfficeBaselineMetricRatio(OpenTypeFont font, double windowsAscenderRatio, double fontSize, out string source)
    {
        double ratio = windowsAscenderRatio;
        source = "OS/2 usWinAscent";
        double typographicAscenderRatio = font.Os2.TypographicAscender / (double)font.UnitsPerEm;
        if (windowsAscenderRatio > PptxTextMetricRules.MaximumOfficeBaselineWindowsAscenderRatio &&
            fontSize <= PptxTextMetricRules.MaximumOfficeTypographicBaselineFontSize &&
            typographicAscenderRatio > 0d &&
            typographicAscenderRatio >= PptxTextMetricRules.MinimumOfficeTypographicBaselineAscenderRatio &&
            typographicAscenderRatio <= PptxTextMetricRules.MaximumBaselineMetricRatio)
        {
            ratio = typographicAscenderRatio;
            source = "OS/2 sTypoAscender";
        }

        return Math.Max(ratio, PptxTextMetricRules.MinimumBaselineMetricRatio);
    }

    private static bool TextMetricUsesOfficeBaselineFloor(OpenTypeFont font, ResolvedRunTextStyle runStyle, TextAdvanceEstimator advanceEstimator, double ascenderRatio)
    {
        return font.TableTags.Contains("MATH") ||
            advanceEstimator.RequestedTypefaceHasMathTable(runStyle.Typeface, runStyle.Bold, runStyle.Italic) ||
            ascenderRatio <= 0d ||
            ascenderRatio > PptxTextMetricRules.MaximumBaselineMetricRatio ||
            ascenderRatio < PptxTextMetricRules.OfficeBaselineFloorMetricThreshold ||
            font.Os2.WindowsDescender / (double)font.UnitsPerEm <= PptxTextMetricRules.OfficeBaselineFloorMaximumWindowsDescenderRatio;
    }

    private static string? ReadBulletText(PptxParagraphBulletModel bullet, ref int autoNumberValue)
    {
        if (bullet.Kind == PptxParagraphBulletKind.None || bullet.Kind == PptxParagraphBulletKind.Blip)
        {
            return null;
        }

        if (bullet.Kind == PptxParagraphBulletKind.Character)
        {
            return bullet.ResolvedCharacter;
        }

        if (bullet.Kind != PptxParagraphBulletKind.AutoNumber)
        {
            return null;
        }

        if (bullet.AutoNumberStartAt is { } start)
        {
            autoNumberValue = start;
        }

        string result = FormatAutoNumber(autoNumberValue, bullet.AutoNumberType);
        autoNumberValue++;
        return result;
    }

    private static bool IsSymbolBulletFont(XElement? bulletFont)
    {
        string? charset = (string?)bulletFont?.Attribute("charset");
        return charset is "2" or "-2";
    }

    private static string MapSymbolBulletText(string bullet)
    {
        Span<char> mapped = bullet.Length <= 256
            ? stackalloc char[bullet.Length]
            : new char[bullet.Length];
        for (int i = 0; i < bullet.Length; i++)
        {
            char ch = bullet[i];
            mapped[i] = ch <= 0x00FF
                ? (char)(0xF000 + ch)
                : ch;
        }

        return new string(mapped);
    }

    private static string FormatAutoNumber(int value, string? type)
    {
        return type switch
        {
            "arabicParenBoth" => $"({value})",
            "arabicParenR" => $"{value})",
            "alphaLcPeriod" => $"{FormatAlphaNumber(value, upper: false)}.",
            "alphaUcPeriod" => $"{FormatAlphaNumber(value, upper: true)}.",
            "alphaLcParenR" => $"{FormatAlphaNumber(value, upper: false)})",
            "alphaUcParenR" => $"{FormatAlphaNumber(value, upper: true)})",
            "romanLcPeriod" => $"{FormatRomanNumber(value, upper: false)}.",
            "romanUcPeriod" => $"{FormatRomanNumber(value, upper: true)}.",
            "romanLcParenR" => $"{FormatRomanNumber(value, upper: false)})",
            "romanUcParenR" => $"{FormatRomanNumber(value, upper: true)})",
            _ => $"{value}."
        };
    }

    private static string FormatAlphaNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        int current = value;
        while (current > 0)
        {
            current--;
            char letter = (char)((upper ? 'A' : 'a') + current % 26);
            builder.Insert(0, letter);
            current /= 26;
        }

        return builder.ToString();
    }

    private static string FormatRomanNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        (int Value, string Numeral)[] numerals =
        [
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        ];

        var builder = new StringBuilder();
        int current = value;
        foreach ((int numeralValue, string numeral) in numerals)
        {
            while (current >= numeralValue)
            {
                builder.Append(numeral);
                current -= numeralValue;
            }
        }

        string result = builder.ToString();
        return upper ? result : result.ToLowerInvariant();
    }

    private static BulletStyle ReadBulletStyle(PptxParagraphBulletModel bullet, double textFontSize, RgbColor textColor, string? textTypeface)
    {
        RgbColor color = bullet.Color ?? textColor;
        double fontSize = textFontSize;
        if (bullet.SizeKind == PptxParagraphBulletSizeKind.Percent &&
            int.TryParse(bullet.SizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizePercent))
        {
            fontSize = textFontSize * Math.Max(0.1d, sizePercent / 100000d);
        }
        else if (bullet.SizeKind == PptxParagraphBulletSizeKind.Points &&
            int.TryParse(bullet.SizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizePoints))
        {
            fontSize = Math.Max(0.1d, sizePoints / 100d);
        }

        return new BulletStyle(fontSize, color, bullet.ResolvedFontTypeface ?? textTypeface);
    }

    private static XElement? FindBulletProperty(XElement? paragraphProperties, string localName)
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

    private static bool TryReadHighlightColor(XElement? runProperties, out RgbColor color)
    {
        XElement? highlight = runProperties?.Element(DrawingNamespace + "highlight");
        string? hex = (string?)highlight?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        return RgbColor.TryParse(hex, out color);
    }

    private static TextVerticalAnchor ParseTextVerticalAnchor(string? anchor)
    {
        return anchor switch
        {
            null or "" => TextVerticalAnchor.Top,
            "t" => TextVerticalAnchor.Top,
            "ctr" => TextVerticalAnchor.Middle,
            "b" => TextVerticalAnchor.Bottom,
            _ when anchor.Equals("t", StringComparison.OrdinalIgnoreCase) => TextVerticalAnchor.Top,
            _ when anchor?.Equals("ctr", StringComparison.OrdinalIgnoreCase) == true => TextVerticalAnchor.Middle,
            _ when anchor?.Equals("b", StringComparison.OrdinalIgnoreCase) == true => TextVerticalAnchor.Bottom,
            _ when !string.IsNullOrEmpty(anchor) => TextVerticalAnchor.Unknown,
            _ => TextVerticalAnchor.Top
        };
    }

    private static double EstimateTextHeight(
        IReadOnlyList<PptxTextParagraphModel> paragraphs,
        double textWidth,
        PptxTextBodyProperties bodyProperties)
    {
        double height = 0d;
        var advanceEstimator = new TextAdvanceEstimator();
        bool allowWrapping = TextBodyAllowsWrapping(bodyProperties);
        bool attachSpacesToFollowingWord = HasNoAutoFit(bodyProperties);
        bool useWindowsFontBoxForDefaultLineSpacing = !IsTableCellVerticalAnchorSource(bodyProperties.VerticalAnchorSource);
        bool hasEstimatedParagraph = false;
        double pendingSpacingAfter = 0d;
        foreach (PptxTextParagraphModel paragraph in paragraphs)
        {
            ResolvedParagraphTextStyle paragraphStyle = paragraph.Style;
            LineSpacing lineSpacing = paragraphStyle.LineSpacing;
            if (!paragraph.HasVisibleContent)
            {
                if (paragraph.HasLayoutContent)
                {
                    double emptyFontSize = paragraph.EndParagraphProperties is null
                        ? paragraphStyle.FontSize
                        : paragraph.EndParagraphStyle.FontSize;
                    if (hasEstimatedParagraph)
                    {
                        height += pendingSpacingAfter + paragraph.EmptySpacingBefore;
                    }

                    string? emptyTypeface = paragraph.EndParagraphStyle.Typeface;
                    bool emptyBold = paragraph.EndParagraphStyle.Bold;
                    bool emptyItalic = paragraph.EndParagraphStyle.Italic;
                    height += ReadEstimatedAnchorEmptyLineAdvance(lineSpacing, emptyFontSize, emptyTypeface, emptyBold, emptyItalic, useWindowsFontBoxForDefaultLineSpacing, advanceEstimator);
                    pendingSpacingAfter = paragraph.EmptySpacingAfter;
                    hasEstimatedParagraph = true;
                }

                continue;
            }

            double paragraphFontSize = paragraphStyle.FontSize;
            if (hasEstimatedParagraph)
            {
                height += pendingSpacingAfter + paragraphStyle.SpacingBefore;
            }

            PptxTextFlowParagraph flowParagraph = BuildTextFlowParagraph(paragraph, attachSpacesToFollowingWord, advanceEstimator);
            double maxFontSize = 0d;
            string? lineTypeface = null;
            bool lineBold = false;
            bool lineItalic = false;
            double lineWidth = 0d;
            bool hasLineContent = false;
            foreach (PptxTextFlowRun flowRun in flowParagraph.Runs)
            {
                ResolvedRunTextStyle runStyle = flowRun.Style;
                if (flowRun.Source.Kind == PptxTextRunKind.Break)
                {
                    height += ReadEstimatedAnchorLineAdvance(lineSpacing, ResolveLineFontSize(maxFontSize, paragraphFontSize), lineTypeface, lineBold, lineItalic, useWindowsFontBoxForDefaultLineSpacing, advanceEstimator);
                    maxFontSize = 0d;
                    lineTypeface = null;
                    lineBold = false;
                    lineItalic = false;
                    lineWidth = 0d;
                    hasLineContent = false;
                    continue;
                }

                foreach (PptxTextFlowSegment segment in flowRun.Segments)
                {
                    if (segment.Kind == PptxTextFlowSegmentKind.Break)
                    {
                        continue;
                    }

                    double fontSize = runStyle.FontSize * segment.FontScale;
                    double advance = advanceEstimator.Measure(
                        segment.AdvanceText,
                        fontSize,
                        runStyle.Typeface,
                        runStyle.Bold,
                        runStyle.Italic,
                        runStyle.CharacterSpacing,
                        runStyle.KerningEnabled);
                    if (allowWrapping &&
                        !string.IsNullOrWhiteSpace(segment.AdvanceText) &&
                        lineWidth > PptxTextMetricRules.TextStateTolerance &&
                        lineWidth + advance > textWidth)
                    {
                        height += ReadEstimatedAnchorLineAdvance(lineSpacing, ResolveLineFontSize(maxFontSize, paragraphFontSize), lineTypeface, lineBold, lineItalic, useWindowsFontBoxForDefaultLineSpacing, advanceEstimator);
                        maxFontSize = fontSize;
                        lineTypeface = runStyle.Typeface;
                        lineBold = runStyle.Bold;
                        lineItalic = runStyle.Italic;
                        lineWidth = 0d;
                        hasLineContent = false;
                    }

                    if (string.IsNullOrWhiteSpace(segment.AdvanceText) && lineWidth <= PptxTextMetricRules.TextStateTolerance)
                    {
                        continue;
                    }

                    if (fontSize >= maxFontSize)
                    {
                        maxFontSize = fontSize;
                        lineTypeface = runStyle.Typeface;
                        lineBold = runStyle.Bold;
                        lineItalic = runStyle.Italic;
                    }

                    lineWidth += advance;
                    hasLineContent |= segment.Draw && segment.Text.Length > 0;
                }
            }

            if (hasLineContent || maxFontSize > PptxTextMetricRules.TextStateTolerance)
            {
                height += ReadEstimatedAnchorLineAdvance(lineSpacing, ResolveLineFontSize(maxFontSize, paragraphFontSize), lineTypeface, lineBold, lineItalic, useWindowsFontBoxForDefaultLineSpacing, advanceEstimator);
            }

            pendingSpacingAfter = paragraphStyle.SpacingAfter;
            hasEstimatedParagraph = true;
        }

        return height;
    }

    private static bool IsTableCellVerticalAnchorSource(PptxTextBodyPropertySource source)
    {
        return source is PptxTextBodyPropertySource.TableCellProperties or PptxTextBodyPropertySource.TableCellStyle;
    }

    private static double ReadEstimatedAnchorLineAdvance(
        LineSpacing lineSpacing,
        double fontSize,
        string? typeface,
        bool bold,
        bool italic,
        bool useWindowsFontBoxForDefaultLineSpacing,
        TextAdvanceEstimator advanceEstimator)
    {
        if (lineSpacing.IsExplicit)
        {
            return ReadLineAdvance(lineSpacing, fontSize);
        }

        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(typeface, bold, italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return ReadLineAdvance(lineSpacing, fontSize);
        }

        double metricUnits = useWindowsFontBoxForDefaultLineSpacing
            ? font.Os2.WindowsAscender + font.Os2.WindowsDescender
            : font.Os2.TypographicAscender - font.Os2.TypographicDescender + font.Os2.TypographicLineGap;
        double metricRatio = metricUnits / font.UnitsPerEm;
        if (metricRatio <= PptxTextMetricRules.MinimumFontLineBoxMetricRatio ||
            metricRatio > PptxTextMetricRules.MaximumFontLineBoxMetricRatio)
        {
            return ReadLineAdvance(lineSpacing, fontSize);
        }

        return fontSize * metricRatio;
    }

    private static double ReadEstimatedAnchorEmptyLineAdvance(
        LineSpacing lineSpacing,
        double fontSize,
        string? typeface,
        bool bold,
        bool italic,
        bool useWindowsFontBoxForDefaultLineSpacing,
        TextAdvanceEstimator advanceEstimator)
    {
        double paragraphAdvance = ReadLineAdvance(lineSpacing, fontSize);
        if (lineSpacing.IsExplicit || !useWindowsFontBoxForDefaultLineSpacing)
        {
            return paragraphAdvance;
        }

        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(typeface, bold, italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return paragraphAdvance;
        }

        double metricUnits = font.Os2.WindowsAscender + font.Os2.WindowsDescender;
        double metricRatio = metricUnits / font.UnitsPerEm;
        if (metricRatio <= PptxTextMetricRules.MinimumFontLineBoxMetricRatio ||
            metricRatio > PptxTextMetricRules.MaximumFontLineBoxMetricRatio)
        {
            return paragraphAdvance;
        }

        double fontBoxAdvance = fontSize * metricRatio;
        return (paragraphAdvance + fontBoxAdvance) / 2d;
    }

}
