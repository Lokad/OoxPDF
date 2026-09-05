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

        IReadOnlyList<PptxPositionedTextSpan> CoalesceSourceTextRunsForInspection(IReadOnlyList<PptxPositionedTextSpan> textSpans)
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

            bool CanCoalesceSourceTextRunForInspection(PptxPositionedTextSpan left, PptxPositionedTextSpan right)
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
        }
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

        PptxTextLayoutModel BuildTextLayoutModelForSources(IReadOnlyList<PptxRenderSource> sources, PptxRenderContext context)
        {
            var frames = new List<PptxTextFrameLayout>();
            foreach (PptxRenderSource source in sources)
            {
                frames.AddRange(BuildTextLayoutModel(context, source, includePlaceholders: false, placeholderSources: []).Frames);
            }

            return new PptxTextLayoutModel(frames);
        }
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

        PptxTextFlowModel BuildTextFlowModelForSources(IReadOnlyList<PptxRenderSource> sources, PptxRenderContext context)
        {
            var frames = new List<PptxTextFlowFrame>();
            foreach (PptxRenderSource source in sources)
            {
                frames.AddRange(BuildTextFlowModel(context, source, includePlaceholders: false, placeholderSources: []).Frames);
            }

            return new PptxTextFlowModel(frames);
        }
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

    internal static IReadOnlyList<PptxTextGlyphRunSnapshot> InspectTextGlyphRuns(PptxDocument document, OoxPackage package, int slideIndex)
    {
        IReadOnlyList<PptxPositionedTextSpan> textSpans = ReadSlideTextSpansForInspection(document, package, slideIndex);
        RenderedFonts renderedFonts = CreateRenderedFonts(textSpans, [], new PresentationFontResolver(null), "F", CancellationToken.None);
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
                if (!renderedFonts.Fonts.TryGetValue(FontKey(run), out RenderedFont rendered))
                {
                    continue;
                }

                TextGlyphRun? glyphRun = BuildTextGlyphRun(rendered.ResourceName, rendered.Font, emissionSpan, rendered.SyntheticBold, rendered.SyntheticItalic);
                if (glyphRun is null)
                {
                    continue;
                }

                TextHighlightRectangle? highlightRectangle = null;
                if (run.HighlightColor is not null)
                {
                    double highlightBaselineY = emissionSpan.LineBox?.BaselineY ?? run.Y + run.BaselineOffset;
                    if (TryGetHighlightRectangle(rendered.Font, run, highlightBaselineY, emissionSpan.GlyphSpan.NaturalWidth, out TextHighlightRectangle rectangle))
                    {
                        highlightRectangle = rectangle;
                    }
                }

                TextDecorationRectangle? underlineRectangle = run.Underline && TryGetUnderlineRectangle(rendered.Font, glyphRun, out TextDecorationRectangle underline)
                    ? underline
                    : null;
                TextDecorationRectangle? strikeRectangle = run.Strike && TryGetStrikeRectangle(rendered.Font, glyphRun, out TextDecorationRectangle strike)
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
                    rendered.Font.Font.TableTags.Contains("GPOS", StringComparer.Ordinal),
                    rendered.Font.Font.TableTags.Contains("kern", StringComparer.Ordinal),
                    rendered.Font.Font.UnitsPerEm,
                    rendered.Font.Font.GlyphCount,
                    rendered.Font.Font.Os2.Version,
                    rendered.Font.Font.Os2.WidthClass,
                    rendered.Font.Font.Os2.TypographicAscender,
                    rendered.Font.Font.Os2.TypographicDescender,
                    rendered.Font.Font.Os2.TypographicLineGap,
                    rendered.Font.Font.Os2.WindowsAscender,
                    rendered.Font.Font.Os2.WindowsDescender,
                    rendered.Font.Font.Post.ItalicAngle,
                    rendered.Font.Font.Post.UnderlinePosition,
                    rendered.Font.Font.Post.UnderlineThickness,
                    rendered.Font.Font.Post.IsFixedPitch,
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
                    emissionSpan.ParagraphBulletKind,
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
                    emissionSpan.FrameWrapMode,
                    emissionSpan.FrameWrapValue,
                    emissionSpan.FrameVerticalOverflowMode,
                    emissionSpan.FrameVerticalOverflowValue,
                    emissionSpan.FrameVerticalOverflowSource,
                    emissionSpan.FrameAutofitMode,
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
