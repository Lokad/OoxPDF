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
    internal static IReadOnlyList<PptxTextRunSnapshot> InspectTextRuns(PptxDocument document, OoxPackage package, int slideIndex)
    {
        return ReadSceneTextRunsForInspection(document, package, slideIndex)
            .Select(ToSnapshot)
            .ToArray();
    }

    private static IReadOnlyList<TextRun> ReadSceneTextRunsForInspection(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null, cancellationToken: CancellationToken.None);
        if (context is null)
        {
            return [];
        }

        return CoalesceSourceTextRunsForInspection(
                ReadSceneShapeTextSpans(context)
                    .Concat(ReadSceneTableTextSpans(context))
                    .ToArray())
            .Select(span => span.Run)
            .ToArray();
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadSlideTextSpansForInspection(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null, cancellationToken: CancellationToken.None);
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

    private static IReadOnlyList<PptxPositionedTextSpan> CoalesceSourceTextRunsForInspection(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        var coalesced = new List<PptxPositionedTextSpan>(textSpans.Count);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            if (coalesced.Count != 0 && CanCoalesceSourceTextRunForInspection(coalesced[^1], span))
            {
                PptxPositionedTextSpan previous = coalesced[^1];
                TextRun mergedRun = previous.Run with
                {
                    Text = previous.Run.Text + span.Run.Text,
                    Width = Math.Max(previous.Run.Width, span.Run.X + span.Run.Width - previous.Run.X)
                };
                coalesced[^1] = previous with
                {
                    Run = mergedRun,
                    EndX = Math.Max(previous.EndX, span.EndX),
                    Atoms = previous.Atoms.Concat(span.Atoms).ToArray()
                };
                continue;
            }

            coalesced.Add(span);
        }

        return coalesced;
    }

    private static bool CanCoalesceSourceTextRunForInspection(PptxPositionedTextSpan left, PptxPositionedTextSpan right)
    {
        return left.SourceRun is not null &&
            ReferenceEquals(left.SourceRun, right.SourceRun) &&
            left.FrameIndex == right.FrameIndex &&
            left.ParagraphIndex == right.ParagraphIndex &&
            left.SourceRunIndex == right.SourceRunIndex &&
            left.Run.Color.Equals(right.Run.Color) &&
            Math.Abs(left.Run.Alpha - right.Run.Alpha) < PptxTextMetricRules.TextStateTolerance &&
            left.Run.HighlightColor.Equals(right.Run.HighlightColor) &&
            left.Run.Bold == right.Run.Bold &&
            left.Run.Italic == right.Run.Italic &&
            left.Run.Underline == right.Run.Underline &&
            left.Run.Strike == right.Run.Strike &&
            left.Run.KerningEnabled == right.Run.KerningEnabled &&
            string.Equals(left.Run.FontFamily, right.Run.FontFamily, StringComparison.OrdinalIgnoreCase);
    }

    internal static PptxTextLayoutSnapshot InspectTextLayout(PptxDocument document, OoxPackage package, int slideIndex)
    {
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null, cancellationToken: CancellationToken.None);
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
        PptxRenderContext? context = TryLoadRenderContext(document, package, slideIndex, new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase), diagnosticSink: null, cancellationToken: CancellationToken.None);
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
}
