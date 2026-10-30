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
    private static IReadOnlyList<PptxPositionedTextSpan> ReadSlideTextSpansForInspection(PptxDocument document, OoxPackage package, int slideIndex, PptxScene? sharedScene = null, PresentationFontResolver? inspectionFonts = null)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null, cancellationToken: CancellationToken.None, sharedScene: sharedScene, inspectionFonts: inspectionFonts);
        if (context is null)
        {
            return [];
        }

        var spans = new List<PptxPositionedTextSpan>();
        int nextFrameIndex = 0;
        // Painting coalesces each node separately. Cached node spans have local
        // frame indexes, so assign distinct inspection indexes before combining
        // them. Keep shape/table ordering aligned with font preflight.
        foreach (bool shapes in new[] { true, false })
        {
            if (context.SceneSlide.ShowMasterShapes)
                Add(context.SceneSlide.MasterNodes, context.MasterColorMap, false, shapes, GroupTransform.Identity);
            Add(context.SceneSlide.LayoutNodes, context.LayoutColorMap, false, shapes, GroupTransform.Identity);
            Add(context.SceneSlide.SlideNodes, context.SlideColorMap, true, shapes, GroupTransform.Identity);
        }
        return spans;

        void Add(IReadOnlyList<PptxSceneNode> nodes, PptxColorMap colorMap, bool placeholders, bool shapes, GroupTransform transform)
        {
            foreach (PptxSceneNode node in nodes)
            {
                if (node.Kind == PptxSceneNodeKind.Group)
                {
                    Add(node.Children, colorMap, placeholders, shapes, transform.Combine(ToGroupTransform(node.GroupTransform)));
                    continue;
                }
                if (node.Kind != (shapes ? PptxSceneNodeKind.Shape : PptxSceneNodeKind.Table)) continue;
                var frame = new List<PptxPositionedTextSpan>();
                if (shapes) AddSceneShapeTextSpansAsync([node], context, frame, colorMap, placeholders).GetAwaiter().GetResult();
                else AddSceneTableTextSpansAsync([node], context, frame, transform, colorMap).GetAwaiter().GetResult();
                int frameCount = 0;
                foreach (PptxPositionedTextSpan span in frame)
                {
                    spans.Add(span with { FrameIndex = nextFrameIndex + span.FrameIndex });
                    frameCount = Math.Max(frameCount, span.FrameIndex + 1);
                }
                nextFrameIndex += frameCount;
            }
        }
    }

    internal static PptxTextLayoutSnapshot InspectTextLayout(PptxDocument document, OoxPackage package, int slideIndex, PptxScene? sharedScene = null, PresentationFontResolver? inspectionFonts = null)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null, cancellationToken: CancellationToken.None, sharedScene: sharedScene, inspectionFonts: inspectionFonts);
        if (context is null)
        {
            return new PptxTextLayoutSnapshot([]);
        }

        PptxTextLayoutModel inheritedLayout = BuildTextLayoutModelForSources(context.InheritedSources, context);
        PptxTextLayoutModel slideLayout = (BuildTextLayoutModelAsync(context, context.SlideSource, includePlaceholders: true, context.InheritedXml).GetAwaiter().GetResult());
        return ToSnapshot(new PptxTextLayoutModel(inheritedLayout.Frames.Concat(slideLayout.Frames).ToArray()));

        PptxTextLayoutModel BuildTextLayoutModelForSources(IReadOnlyList<PptxRenderSource> sources, PptxRenderContext context)
        {
            var frames = new List<PptxTextFrameLayout>();
            foreach (PptxRenderSource source in sources)
            {
                frames.AddRange((BuildTextLayoutModelAsync(context, source, includePlaceholders: false, placeholderSources: []).GetAwaiter().GetResult()).Frames);
            }

            return new PptxTextLayoutModel(frames);
        }
    }

    internal static PptxTextFlowSnapshot InspectTextFlow(PptxDocument document, OoxPackage package, int slideIndex, PptxScene? sharedScene = null, PresentationFontResolver? inspectionFonts = null)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null, cancellationToken: CancellationToken.None, sharedScene: sharedScene, inspectionFonts: inspectionFonts);
        if (context is null)
        {
            return new PptxTextFlowSnapshot([]);
        }

        PptxTextFlowModel inheritedFlow = BuildTextFlowModelForSources(context.InheritedSources, context);
        PptxTextFlowModel slideFlow = (BuildTextFlowModelAsync(context, context.SlideSource, includePlaceholders: true, context.InheritedXml).GetAwaiter().GetResult());
        return ToSnapshot(new PptxTextFlowModel(inheritedFlow.Frames.Concat(slideFlow.Frames).ToArray()));

        PptxTextFlowModel BuildTextFlowModelForSources(IReadOnlyList<PptxRenderSource> sources, PptxRenderContext context)
        {
            var frames = new List<PptxTextFlowFrame>();
            foreach (PptxRenderSource source in sources)
            {
                frames.AddRange((BuildTextFlowModelAsync(context, source, includePlaceholders: false, placeholderSources: []).GetAwaiter().GetResult()).Frames);
            }

            return new PptxTextFlowModel(frames);
        }
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
        PptxInterGlyphAdjustmentSummary adjustments = SummarizeInterGlyphAdjustments(span.Glyphs);
        return new PptxTextGlyphSpanLayoutSnapshot(
            span.Text,
            span.Typeface,
            span.FontSize,
            span.LeadingAdjustment,
            span.NaturalWidth,
            span.LayoutWidth,
            span.Glyphs.Count,
            span.Glyphs.Skip(1).FirstOrDefault()?.AdjustmentBefore ?? 0d,
            adjustments.Count,
            adjustments.Sum,
            adjustments.Min,
            adjustments.Max,
            adjustments.Average,
            span.Glyphs.Select(ToSnapshot).ToArray());
    }

    private static PptxInterGlyphAdjustmentSummary SummarizeInterGlyphAdjustments(IReadOnlyList<PptxTextGlyphLayout> glyphs)
    {
        int count = 0;
        double sum = 0d;
        double min = 0d;
        double max = 0d;

        for (int index = 1; index < glyphs.Count; index++)
        {
            double adjustment = glyphs[index].AdjustmentBefore;
            if (count == 0)
            {
                min = adjustment;
                max = adjustment;
            }
            else
            {
                min = Math.Min(min, adjustment);
                max = Math.Max(max, adjustment);
            }

            sum += adjustment;
            count++;
        }

        return count == 0
            ? default
            : new PptxInterGlyphAdjustmentSummary(count, sum, min, max, sum / count);
    }

    private readonly record struct PptxInterGlyphAdjustmentSummary(
        int Count,
        double Sum,
        double Min,
        double Max,
        double Average);

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

    internal static IReadOnlyList<PptxTextGlyphRunSnapshot> InspectTextGlyphRuns(PptxDocument document, OoxPackage package, int slideIndex, PptxScene? sharedScene = null, PresentationFontResolver? inspectionFonts = null)
    {
        inspectionFonts ??= new PresentationFontResolver(null);
        IReadOnlyList<PptxPositionedTextSpan> textSpans = ReadSlideTextSpansForInspection(document, package, slideIndex, sharedScene, inspectionFonts);
        RenderedFonts renderedFonts = CreateRenderedFonts(textSpans, [], inspectionFonts, "F", CancellationToken.None);
        textSpans = SplitLeadingSpacesAtHighlightBoundaries(textSpans);
        textSpans = CoalesceAdjacentTextSpans(textSpans, compareHighlight: true);
        textSpans = CoalesceUnderlineSpans(textSpans);
        textSpans = ApplyOfficePdfCharacterSpacing(textSpans);
        var glyphRuns = new List<PptxTextGlyphRunSnapshot>();
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            foreach (PptxPositionedTextSpan emissionSpan in SplitSpanByGlyphTypeface(span))
            {
                TextRun run = emissionSpan.Run;
                if (!renderedFonts.Fonts.TryGetValue(FontRequestForRun(run), out RenderedFont rendered))
                {
                    continue;
                }

                // RV01: inspection mirrors emission misses for fallback spans; the
                // conversion path paints them with the fallback face.
                if (rendered.Font is not { } snapshotEmbedded)
                {
                    continue;
                }

                TextGlyphRun? glyphRun = BuildTextGlyphRun(rendered.ResourceName, snapshotEmbedded, emissionSpan, rendered.SyntheticBold, rendered.SyntheticItalic);
                if (glyphRun is null)
                {
                    continue;
                }

                TextHighlightRectangle? highlightRectangle = null;
                if (run.HighlightColor is not null)
                {
                    double highlightBaselineY = emissionSpan.LineBox?.BaselineY ?? run.Y + run.BaselineOffset;
                    if (TryGetHighlightRectangle(snapshotEmbedded, run, highlightBaselineY, emissionSpan.GlyphSpan.NaturalWidth, out TextHighlightRectangle rectangle))
                    {
                        highlightRectangle = rectangle;
                    }
                }

                TextDecorationRectangle? underlineRectangle = run.Underline && TryGetUnderlineRectangle(snapshotEmbedded, glyphRun, out TextDecorationRectangle underline)
                    ? underline
                    : null;
                TextDecorationRectangle? strikeRectangle = run.Strike && TryGetStrikeRectangle(snapshotEmbedded, glyphRun, out TextDecorationRectangle strike)
                    ? strike
                    : null;
                PptxInterGlyphAdjustmentSummary adjustments = SummarizeTextInterGlyphAdjustments(glyphRun.Glyphs);

                glyphRuns.Add(new PptxTextGlyphRunSnapshot(
                    run.Text,
                    glyphRun.X,
                    glyphRun.BaselineY,
                    glyphRun.Width,
                    emissionSpan.GlyphSpan.NaturalWidth,
                    emissionSpan.GlyphSpan.LayoutWidth,
                    run.FontFamily,
                    rendered.Resolution.FamilyName,
                    rendered.Resolution.IsFallback,
                    rendered.Resolution.Bold,
                    rendered.Resolution.Italic,
                    rendered.Resolution.WeightClass,
                    rendered.Resolution.FontFaceIndex,
                    rendered.Resolution.HasMathTable,
                    snapshotEmbedded.Font.TableTags.Contains("GPOS", StringComparer.Ordinal),
                    snapshotEmbedded.Font.TableTags.Contains("kern", StringComparer.Ordinal),
                    snapshotEmbedded.Font.UnitsPerEm,
                    snapshotEmbedded.Font.GlyphCount,
                    snapshotEmbedded.Font.Os2.Version,
                    snapshotEmbedded.Font.Os2.WidthClass,
                    snapshotEmbedded.Font.Os2.TypographicAscender,
                    snapshotEmbedded.Font.Os2.TypographicDescender,
                    snapshotEmbedded.Font.Os2.TypographicLineGap,
                    snapshotEmbedded.Font.Os2.WindowsAscender,
                    snapshotEmbedded.Font.Os2.WindowsDescender,
                    snapshotEmbedded.Font.Post.ItalicAngle,
                    snapshotEmbedded.Font.Post.UnderlinePosition,
                    snapshotEmbedded.Font.Post.UnderlineThickness,
                    snapshotEmbedded.Font.Post.IsFixedPitch,
                    run.Bold,
                    run.Italic,
                    run.Underline,
                    run.Strike,
                    glyphRun.SyntheticBold,
                    glyphRun.SyntheticItalic,
                    run.HighlightColor,
                    highlightRectangle?.X,
                    highlightRectangle?.Y,
                    highlightRectangle?.Width,
                    highlightRectangle?.Height,
                    underlineRectangle?.X,
                    underlineRectangle?.Y,
                    underlineRectangle?.Width,
                    underlineRectangle?.Height,
                    strikeRectangle?.X,
                    strikeRectangle?.Y,
                    strikeRectangle?.Width,
                    strikeRectangle?.Height,
                    emissionSpan.FrameIndex,
                    emissionSpan.ParagraphIndex,
                    emissionSpan.SourceRunIndex,
                    emissionSpan.ParagraphBulletKind.ToString(),
                    emissionSpan.ParagraphAutoNumberType,
                    emissionSpan.ParagraphAutoNumberStartAt,
                    emissionSpan.LineIndex,
                    emissionSpan.SpanIndex,
                    emissionSpan.LineSpanCount,
                    emissionSpan.FrameFontScale,
                    emissionSpan.FrameShapeX,
                    emissionSpan.FrameShapeTopY,
                    emissionSpan.FrameShapeWidth,
                    emissionSpan.FrameShapeHeight,
                    emissionSpan.TableRowIndex,
                    emissionSpan.TableColumnIndex,
                    emissionSpan.TableRowSpan,
                    emissionSpan.TableColumnSpan,
                    emissionSpan.FrameInsetLeft,
                    emissionSpan.FrameInsetRight,
                    emissionSpan.FrameInsetTop,
                    emissionSpan.FrameInsetBottom,
                    emissionSpan.FrameWrapMode.ToString(),
                    emissionSpan.FrameWrapValue,
                    emissionSpan.FrameVerticalOverflowMode.ToString(),
                    emissionSpan.FrameVerticalOverflowValue,
                    emissionSpan.FrameVerticalOverflowSource.ToString(),
                    emissionSpan.FrameAutofitModeValue,
                    emissionSpan.FrameTextX,
                    emissionSpan.FrameTextWidth,
                    emissionSpan.FrameTextWrapWidth,
                    emissionSpan.FrameTextHeight,
                    emissionSpan.FrameClipX,
                    emissionSpan.FrameClipWidth,
                    emissionSpan.FrameClipY,
                    emissionSpan.FrameClipHeight,
                    emissionSpan.FrameColumnCount,
                    emissionSpan.FrameColumnSpacing,
                    emissionSpan.LineBox?.TopY ?? run.Y,
                    emissionSpan.LineBox?.Advance ?? 0d,
                    emissionSpan.LineBox?.MaxFontSize ?? run.FontSize,
                    run.FontSize,
                    glyphRun.PdfFontSize,
                    emissionSpan.GlyphSpan.CharacterSpacing,
                    glyphRun.PdfCharacterSpacing,
                    glyphRun.Glyphs.Count,
                    glyphRun.Glyphs.Skip(1).FirstOrDefault()?.AdjustmentBefore ?? 0d,
                    adjustments.Count,
                    adjustments.Sum,
                    adjustments.Min,
                    adjustments.Max,
                    adjustments.Average,
                    glyphRun.Glyphs
                        .Select(glyph => new PptxTextGlyphRunAtomSnapshot(
                            glyph.CodePoint,
                            glyph.Typeface,
                            glyph.TypefaceResolutionSource.ToString(),
                            glyphRun.ResourceName,
                            glyph.GlyphId,
                            glyph.Advance,
                            glyph.AdjustmentBefore))
                        .ToArray()));
            }
        }

        return glyphRuns;
    }

    private static PptxInterGlyphAdjustmentSummary SummarizeTextInterGlyphAdjustments(IReadOnlyList<TextGlyphAtom> glyphs)
    {
        int count = 0;
        double sum = 0d;
        double min = 0d;
        double max = 0d;

        for (int index = 1; index < glyphs.Count; index++)
        {
            double adjustment = glyphs[index].AdjustmentBefore;
            if (count == 0)
            {
                min = adjustment;
                max = adjustment;
            }
            else
            {
                min = Math.Min(min, adjustment);
                max = Math.Max(max, adjustment);
            }

            sum += adjustment;
            count++;
        }

        return count == 0
            ? default
            : new PptxInterGlyphAdjustmentSummary(count, sum, min, max, sum / count);
    }
}
