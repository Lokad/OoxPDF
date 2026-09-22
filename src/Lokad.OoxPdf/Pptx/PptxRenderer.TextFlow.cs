using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<PptxPositionedTextSpan> ReadSceneShapeTextSpans(PptxRenderContext context, bool includeMasterNodes = true)
    {
        var textSpans = new List<PptxPositionedTextSpan>();
        // font preflight traversed master nodes even when the slide suppresses
        // their paint, laying out invisible shapes and embedding orphan fonts. Honor
        // visibility consistently with painting.
        if (includeMasterNodes)
        {
            AddSceneShapeTextSpans(context.SceneSlide.MasterNodes, context, textSpans, context.MasterColorMap, renderPlaceholders: false);
        }
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
        context.CancellationToken.ThrowIfCancellationRequested();
        return FlattenTextLayoutToSpans(BuildTextLayoutModel(context, source, includePlaceholders, placeholderSources), context.FontResolver);
    }

    private static PptxTextLayoutModel BuildTextLayoutModel(
        PptxRenderContext context,
        PptxRenderSource source,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources)
    {
        return BuildTextLayoutModel(source.Xml, context.Document, context.Theme, source.ColorMap, context.SlideNumber, includePlaceholders, placeholderSources, context.FontResolver, context.CancellationToken);
    }

    private static PptxTextLayoutModel BuildTextLayoutModel(
        XDocument slideXml,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        int slideNumber,
        bool includePlaceholders,
        IReadOnlyList<XDocument> placeholderSources,
        PresentationFontResolver? fontResolver,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var advanceEstimator = new TextAdvanceEstimator(fontResolver, cancellationToken);
        var frames = new List<PptxTextFrameLayout>();
        IReadOnlyList<PptxTextFrameModel> frameModels = BuildTextFrameModels(slideXml, document, theme, colorMap, slideNumber, includePlaceholders, placeholderSources, fontResolver, cancellationToken);
        foreach (PptxTextFrameModel frameModel in frameModels)
        {
            cancellationToken.ThrowIfCancellationRequested();
            frames.Add(BuildTextFrameLayout(frameModel, document, advanceEstimator));
        }

        return new PptxTextLayoutModel(frames);
    }

    private static PptxTextFrameLayout BuildTextFrameLayout(PptxTextFrameModel frameModel, PptxDocument document, TextAdvanceEstimator advanceEstimator)
    {
        frameModel = ResetEstimatedVerticalAnchorOffset(frameModel);
        PptxTextFlowFrame flowFrame = BuildTextFlowFrame(frameModel, document, advanceEstimator);
        PptxTextFrameLayout layout = BuildTextFrameLayout(flowFrame, document, advanceEstimator, true);
        if (HasShapeAutoFit(frameModel.BodyProperties) && UsesRotatedFrameAutoFit(frameModel.Orientation))
        {
            // spAutoFit grows the shape around overflowing text, so under the non-clipping
            // overflow mode the wrapped layout stands at full size: unwrapping plus font shrink
            // models shrink-to-fit, which dropped vertical-text-port row 2 to 4.08pt against
            // Office 24pt and the vert-oriented 270 case to 8.04pt against Office 20.04pt.
            // Clipping modes and true vert270 keep the legacy shrink path below (vert270 behavior
            // is unobserved; PptxSyntheticVerticalShapeAutoFitPrefersSingleLine pins it).
            if (frameModel.Orientation == PptxTextOrientation.Vertical &&
                !ClipsTextVerticalOverflow(frameModel.BodyProperties.VerticalOverflow))
            {
                return ApplyActualVerticalAnchorOffsetIfNeeded(layout, document, advanceEstimator, allowWrapping: true);
            }

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
                BuildTextFrameLayout(BuildTextFlowFrame(fitted, document, advanceEstimator), document, advanceEstimator, true),
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
        var advanceEstimator = new TextAdvanceEstimator(null, CancellationToken.None);
        return new PptxTextFlowModel(frames.Select(frame => BuildTextFlowFrame(frame, document, advanceEstimator)).ToArray());
    }

    private static PptxTextFlowFrame BuildTextFlowFrame(PptxTextFrameModel frame, PptxDocument document, TextAdvanceEstimator advanceEstimator)
    {
        // For vertical text the flow stacking axis maps onto the shape horizontal axis with
        // the first line at the shape-right side, so the stack origin consumes the right
        // inset, not the top one. Three asymmetric-inset Office probes prove it (ignored
        // vprobe-* artifacts, all 20pt Cambria): zero insets put the first token at 328.61
        // (ours 328.52), killing the Cambria-baseline rival; l14.4/r0 keeps 328.61, killing
        // Top; l14.4/r7.2 keeps 321.41 unchanged, killing Left. A uniform -0.09 start
        // residual stands across all five vertical samples. Vertical270 keeps legacy Top
        // (unobserved).
        double stackOriginInset = frame.Orientation == PptxTextOrientation.Vertical
            ? frame.Insets.Right
            : frame.Insets.Top;
        var box = new PptxTextFlowBox(
            frame.FlowYTop,
            document.SlideHeightPoints - frame.FlowYTop - stackOriginInset - frame.VerticalOffset,
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
        bool attachSpacesToFollowingWord = UsesOfficeFollowingSpaceFlow(frame);
        return new PptxTextFlowFrame(frame, box, frame.Paragraphs.Select(paragraph => BuildTextFlowParagraph(paragraph, attachSpacesToFollowingWord, advanceEstimator)).ToArray());
    }

    private static bool UsesOfficeFollowingSpaceFlow(PptxTextFrameModel frame)
    {
        return HasNoAutoFit(frame.BodyProperties) || frame.TableRowIndex.HasValue;
    }

    private static bool HasShapeAutoFit(PptxTextBodyProperties bodyProperties)
    {
        return bodyProperties.AutofitModeValue == "spAutoFit";
    }

    private static bool HasNoAutoFit(PptxTextBodyProperties bodyProperties)
    {
        return bodyProperties.AutofitModeValue == "noAutofit";
    }

    // Office default when bodyPr carries no autofit element (North-clone probes 2026-09-06: absent-autofit content centers by line advances).
    private static bool HasAbsentAutofit(PptxTextBodyProperties bodyProperties)
    {
        return bodyProperties.AutofitModeValue == string.Empty;
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
        // Vertical middle/bottom anchors center the laid-out stack in the left/right-based
        // text height (the stacking extent spans the shape width) with unclamped negatives
        // (Office middle-anchor probe sits +6.77 off the top baseline, bottom +13.56).
        bool verticalAnchorPath = frame.Orientation == PptxTextOrientation.Vertical;
        if ((frame.Orientation != PptxTextOrientation.Horizontal && !verticalAnchorPath) ||
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
        double anchorBoxHeight = verticalAnchorPath
            ? OoxUnits.EmuToPoints(frame.Bounds.Width) - frame.Insets.Left - frame.Insets.Right
            : frame.TextHeight;
        double slack = anchorBoxHeight - occupiedHeight;
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
            bool StartsWithDrawableRegularSpace(IReadOnlyList<PptxTextFlowSegment> segments)
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

            PptxTextFlowRun flowRun = BuildTextFlowRun(run, paragraph.Style.DefaultRunProperties);
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

        IReadOnlyList<PptxTextFlowSegment> HideSpacesAfterBoundaryPunctuation(IReadOnlyList<PptxTextFlowSegment> segments, ref bool hideLeadingSpaces)
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

            IReadOnlyList<PptxTextFlowSegment> HideLeadingSpacesIfNeeded(PptxTextFlowSegment segment, bool hideLeadingSpaces)
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
        }

        PptxTextFlowRun BuildTextFlowRun(PptxTextRunModel run, XElement? defaultRunProperties)
        {
            if (run.Kind == PptxTextRunKind.Break)
            {
                return new PptxTextFlowRun(run, run.Style, [new PptxTextFlowSegment("\n", "\n", PptxTextFlowSegmentKind.Break, Draw: false, PreventCoalesce: true, FontScale: 1d)]);
            }

            var segments = new List<PptxTextFlowSegment>();
            string[] tabParts = run.Text.Split('\t');
            for (int tabPartIndex = 0; tabPartIndex < tabParts.Length; tabPartIndex++)
            {
                if (tabPartIndex > 0)
                {
                    segments.Add(new PptxTextFlowSegment(" ", " ", PptxTextFlowSegmentKind.Tab, Draw: true, PreventCoalesce: true, FontScale: 1d));
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

        bool CanCoalesceFlowRunStyles(ResolvedRunTextStyle left, ResolvedRunTextStyle right)
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

    private static bool HasDrawableText(IReadOnlyList<PptxTextFlowSegment> segments)
    {
        return segments.Any(static segment => segment.Draw && segment.Kind == PptxTextFlowSegmentKind.Text && segment.Text.Length != 0);
    }

    private static IReadOnlyList<TextRun> FlattenTextLayout(PptxTextLayoutModel layout)
    {
        return FlattenTextLayoutToSpans(layout, null).Select(span => span.Run).ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> FlattenTextLayoutToSpans(PptxTextLayoutModel layout, PresentationFontResolver? fontResolver)
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
                span.SourceRun?.RunIndex,
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
        return AddEllipsisOverflowMarkers();

        IReadOnlyList<PptxPositionedTextSpan> AddEllipsisOverflowMarkers()
        {
            if (spans.Length == 0)
            {
                return spans;
            }

            var result = new List<PptxPositionedTextSpan>(spans.Length);
            var advanceEstimator = new TextAdvanceEstimator(fontResolver, CancellationToken.None);
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
                result.Add(CreateEllipsisOverflowMarker(last));
            }

            return result
                .OrderBy(span => span.FrameIndex)
                .ThenBy(span => span.ParagraphIndex)
                .ThenBy(span => span.LineIndex)
                .ThenBy(span => span.SpanIndex)
                .ToArray();

            PptxPositionedTextSpan CreateEllipsisOverflowMarker(PptxPositionedTextSpan anchor)
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
                    GlyphSpan = BuildGlyphSpan(run, advanceEstimator, 0d)
                };
            }
        }
    }

}
