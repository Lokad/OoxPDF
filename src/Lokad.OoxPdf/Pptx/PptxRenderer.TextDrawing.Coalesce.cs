using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<TextRun> CoalesceAdjacentTextRuns(IReadOnlyList<TextRun> textRuns, bool compareHighlight)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (run.Text.Length == 0)
            {
                continue;
            }

            if (coalesced.Count != 0 && CanCoalesceTextRun(coalesced[^1], run, compareHighlight, true))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = run.X + run.Width - previous.X
                };
            }
            else
            {
                coalesced.Add(run);
            }
        }

        return coalesced;
    }

    private static IReadOnlyList<PptxPositionedTextSpan> CoalesceAdjacentTextSpans(IReadOnlyList<PptxPositionedTextSpan> textSpans, bool compareHighlight)
    {
        var coalesced = new List<PptxPositionedTextSpan>(textSpans.Count);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            TextRun run = span.Run;
            if (run.Text.Length == 0)
            {
                continue;
            }

            if (coalesced.Count != 0 && CanCoalesceTextSpan(coalesced[^1], span))
            {
                PptxPositionedTextSpan previous = coalesced[^1];
                TextRun mergedRun = previous.Run with
                {
                    Text = previous.Run.Text + run.Text,
                    Width = run.X + run.Width - previous.Run.X
                };
                coalesced[^1] = previous with
                {
                    Run = mergedRun,
                    EndX = span.EndX,
                    Atoms = previous.Atoms.Concat(span.Atoms).ToArray(),
                    GlyphSpan = MergeGlyphSpans(mergedRun, previous.Run, previous.GlyphSpan, run, span.GlyphSpan)
                };
            }
            else
            {
                coalesced.Add(span);
            }
        }

        return coalesced;

        bool CanCoalesceTextSpan(PptxPositionedTextSpan left, PptxPositionedTextSpan right)
        {
            if (PreservesHyperlinkTextOperationBoundary(left, right))
            {
                return false;
            }

            if (left.SourceRun is not null &&
                right.SourceRun is not null &&
                !ReferenceEquals(left.SourceRun, right.SourceRun) &&
                PreservesSourceRunTextOperationBoundary(left, right))
            {
                return false;
            }

            return CanCoalesceTextRun(left.Run, right.Run, compareHighlight, PreservesHighlightTextOperationBoundaries(left, right));
        }
    }

    private static bool CanCoalesceTextRun(TextRun left, TextRun right, bool compareHighlight, bool preserveHighlightBoundary)
    {
        return Math.Abs(left.Y - right.Y) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.FontSize - right.FontSize) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.CharacterSpacing - right.CharacterSpacing) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.BaselineOffset - right.BaselineOffset) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.RotationDegrees - right.RotationDegrees) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.RotationCenterX - right.RotationCenterX) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.RotationCenterY - right.RotationCenterY) < PptxTextMetricRules.CoordinateTolerance &&
            left.FlipHorizontal == right.FlipHorizontal &&
            left.FlipVertical == right.FlipVertical &&
            Math.Abs(left.ClipX - right.ClipX) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.ClipY - right.ClipY) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.ClipWidth - right.ClipWidth) < PptxTextMetricRules.CoordinateTolerance &&
            Math.Abs(left.ClipHeight - right.ClipHeight) < PptxTextMetricRules.CoordinateTolerance &&
            left.Color.Equals(right.Color) &&
            Math.Abs(left.Alpha - right.Alpha) < PptxTextMetricRules.TextStateTolerance &&
            TextOutlinesEqual(left.Outline, right.Outline) &&
            (!compareHighlight || !preserveHighlightBoundary || left.HighlightColor.Equals(right.HighlightColor)) &&
            !left.PreventCoalesce &&
            !right.PreventCoalesce &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.Underline == right.Underline &&
            left.Strike == right.Strike &&
            left.KerningEnabled == right.KerningEnabled &&
            left.Alignment == right.Alignment &&
            string.Equals(left.FontFamily, right.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            right.X >= left.X &&
            Math.Abs(right.X - (left.X + left.Width)) < PptxTextMetricRules.TextCoalesceGap(left.FontSize);
    }

    private static bool PreservesHighlightTextOperationBoundaries(PptxPositionedTextSpan left, PptxPositionedTextSpan right)
    {
        return left.SourceAlignment != TextAlignment.Left || right.SourceAlignment != TextAlignment.Left;
    }

    // Spans on opposite sides of a hyperlink boundary must stay separate text
    // operations so each emitted span carries exactly one link identity (S08).
    private static bool PreservesHyperlinkTextOperationBoundary(PptxPositionedTextSpan left, PptxPositionedTextSpan right)
    {
        return !string.Equals(left.SourceRun?.Style.HyperlinkClickId, right.SourceRun?.Style.HyperlinkClickId, StringComparison.Ordinal);
    }

    private static bool PreservesSourceRunTextOperationBoundary(PptxPositionedTextSpan left, PptxPositionedTextSpan right)
    {
        if (left.SourceAlignment != TextAlignment.Left || right.SourceAlignment != TextAlignment.Left)
        {
            return true;
        }

        if (UsesMathTypeface(left.Run.FontFamily) ||
            UsesMathTypeface(right.Run.FontFamily) ||
            string.Equals(left.FrameAutofitMode, "spAutoFit", StringComparison.Ordinal) ||
            string.Equals(right.FrameAutofitMode, "spAutoFit", StringComparison.Ordinal))
        {
            return true;
        }

        // Office merges adjacent same-style source runs into one text operation.
        // The previous trailing-whitespace veto kept runs like Left/Tab/Spaces as
        // three operations (whitespace-controls probe: 15 vs 13 ops with identical
        // decoded text and pixels). Merge whenever styles match; hyperlink,
        // math/autofit, and alignment guards elsewhere stay intact, as do the
        // PreventCoalesce vetoes from control segments and glyph-typeface splits.
        return false;
    }

    private static bool UsesMathTypeface(string? typeface)
    {
        return typeface?.IndexOf("Math", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static IReadOnlyList<TextRun> CoalesceUnderlineRuns(IReadOnlyList<TextRun> textRuns)
    {
        var coalesced = new List<TextRun>(textRuns.Count);
        foreach (TextRun run in textRuns)
        {
            if (coalesced.Count > 0 && CanCoalesceUnderlineRun(coalesced[^1], run))
            {
                TextRun previous = coalesced[^1];
                coalesced[^1] = previous with
                {
                    Text = previous.Text + run.Text,
                    Width = Math.Max(previous.Width, run.X + run.Width - previous.X)
                };
                continue;
            }

            coalesced.Add(run);
        }

        return coalesced;
    }

    private static IReadOnlyList<PptxPositionedTextSpan> CoalesceUnderlineSpans(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        var coalesced = new List<PptxPositionedTextSpan>(textSpans.Count);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            if (coalesced.Count > 0 && !PreservesHyperlinkTextOperationBoundary(coalesced[^1], span) && CanCoalesceUnderlineRun(coalesced[^1].Run, span.Run))
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
                    EndX = span.EndX,
                    Atoms = previous.Atoms.Concat(span.Atoms).ToArray(),
                    GlyphSpan = MergeGlyphSpans(mergedRun, previous.Run, previous.GlyphSpan, span.Run, span.GlyphSpan)
                };
                continue;
            }

            coalesced.Add(span);
        }

        return coalesced;
    }

    private static PptxTextGlyphSpanLayout MergeGlyphSpans(TextRun mergedRun, TextRun leftRun, PptxTextGlyphSpanLayout left, TextRun rightRun, PptxTextGlyphSpanLayout right)
    {
        if (left.Glyphs.Count == 0 && right.Glyphs.Count == 0)
        {
            return PptxTextGlyphSpanLayout.Empty(mergedRun);
        }

        if (left.Glyphs.Count == 0)
        {
            return right with
            {
                Text = mergedRun.Text,
                LayoutWidth = mergedRun.Width
            };
        }

        if (right.Glyphs.Count == 0)
        {
            return left with
            {
                Text = mergedRun.Text,
                LayoutWidth = mergedRun.Width
            };
        }

        double leftVisualEnd = leftRun.X + left.NaturalWidth;
        double interSpanGap = rightRun.X - leftVisualEnd;
        PptxTextGlyphLayout[] glyphs = left.Glyphs
            .Concat(right.Glyphs.Select((glyph, index) => index == 0
                ? glyph with { AdjustmentBefore = glyph.AdjustmentBefore + interSpanGap }
                : glyph))
            .ToArray();
        double naturalWidth = glyphs.Sum(glyph => glyph.Advance) + glyphs.Sum(glyph => glyph.AdjustmentBefore);
        return left with
        {
            Text = mergedRun.Text,
            NaturalWidth = Math.Max(0d, naturalWidth),
            LayoutWidth = mergedRun.Width,
            Glyphs = glyphs
        };
    }

    private static bool CanCoalesceUnderlineRun(TextRun left, TextRun right)
    {
        return left.Underline &&
            right.Underline &&
            !left.Strike &&
            !right.Strike &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            left.KerningEnabled == right.KerningEnabled &&
            left.Alignment == right.Alignment &&
            string.Equals(left.FontFamily, right.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            left.Color.Equals(right.Color) &&
            NearlyEqual(left.Alpha, right.Alpha) &&
            TextOutlinesEqual(left.Outline, right.Outline) &&
            left.HighlightColor.Equals(right.HighlightColor) &&
            !left.PreventCoalesce &&
            !right.PreventCoalesce &&
            NearlyEqual(left.Y, right.Y) &&
            NearlyEqual(left.Height, right.Height) &&
            NearlyEqual(left.ClipX, right.ClipX) &&
            NearlyEqual(left.ClipY, right.ClipY) &&
            NearlyEqual(left.ClipWidth, right.ClipWidth) &&
            NearlyEqual(left.ClipHeight, right.ClipHeight) &&
            NearlyEqual(left.FontSize, right.FontSize) &&
            NearlyEqual(left.CharacterSpacing, right.CharacterSpacing) &&
            NearlyEqual(left.BaselineOffset, right.BaselineOffset) &&
            NearlyEqual(left.RotationDegrees, right.RotationDegrees) &&
            NearlyEqual(left.RotationCenterX, right.RotationCenterX) &&
            NearlyEqual(left.RotationCenterY, right.RotationCenterY) &&
            left.FlipHorizontal == right.FlipHorizontal &&
            left.FlipVertical == right.FlipVertical &&
            Math.Abs((left.X + left.Width) - right.X) <= PptxTextMetricRules.UnderlineCoalesceGap(left.FontSize);
    }

    private static bool NearlyEqual(double left, double right)
    {
        return Math.Abs(left - right) <= PptxTextMetricRules.TextStateTolerance;
    }

    private static bool TextOutlinesEqual(TextOutline? left, TextOutline? right)
    {
        if (left is null || right is null)
        {
            return left is null && right is null;
        }

        TextOutline leftOutline = left.Value;
        TextOutline rightOutline = right.Value;
        return leftOutline.Color.Equals(rightOutline.Color) &&
            NearlyEqual(leftOutline.Alpha, rightOutline.Alpha) &&
            NearlyEqual(leftOutline.Width, rightOutline.Width);
    }

    private static IEnumerable<PptxPositionedTextSpan> SplitSpanByGlyphTypeface(PptxPositionedTextSpan span)
    {
        if (span.GlyphSpan.Glyphs.Count == 0)
        {
            yield return span;
            yield break;
        }

        double cursor = span.Run.X;
        int index = 0;
        while (index < span.GlyphSpan.Glyphs.Count)
        {
            PptxTextGlyphLayout first = span.GlyphSpan.Glyphs[index];
            string typeface = string.IsNullOrWhiteSpace(first.Typeface) ? PptxFontFallbackRules.ResolveDefaultLatinTypeface(span.Run.FontFamily) : first.Typeface;
            int start = index;
            index++;
            while (index < span.GlyphSpan.Glyphs.Count &&
                string.Equals(span.GlyphSpan.Glyphs[index].Typeface ?? PptxFontFallbackRules.ResolveDefaultLatinTypeface(span.Run.FontFamily), typeface, StringComparison.OrdinalIgnoreCase))
            {
                index++;
            }

            PptxTextGlyphLayout[] glyphs = span.GlyphSpan.Glyphs.Skip(start).Take(index - start).ToArray();
            double leadingOffset = start == 0 ? 0d : glyphs[0].AdjustmentBefore;
            if (start != 0 && Math.Abs(leadingOffset) > PptxTextMetricRules.TextStateTolerance)
            {
                glyphs[0] = glyphs[0] with { AdjustmentBefore = 0d };
            }

            double naturalWidth = glyphs.Sum(glyph => glyph.Advance) + glyphs.Sum(glyph => glyph.AdjustmentBefore);
            double x = cursor + leadingOffset;
            string text = string.Concat(glyphs.Select(glyph => char.ConvertFromUtf32(glyph.CodePoint)));
            TextRun run = span.Run with
            {
                Text = text,
                X = x,
                Width = naturalWidth,
                FontFamily = typeface,
                Alignment = TextAlignment.Left,
                PreventCoalesce = true
            };
            yield return span with
            {
                Run = run,
                EndX = x + naturalWidth,
                GlyphSpan = new PptxTextGlyphSpanLayout(
                    text,
                    typeface,
                    span.GlyphSpan.Bold,
                    span.GlyphSpan.Italic,
                    span.GlyphSpan.FontSize,
                    span.GlyphSpan.CharacterSpacing,
                    span.GlyphSpan.KerningEnabled,
                    start == 0 ? span.GlyphSpan.LeadingAdjustment : 0d,
                    naturalWidth,
                    naturalWidth,
                    glyphs)
            };

            cursor = x + naturalWidth;
        }
    }
}
