using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<PptxPositionedTextSpan> ApplyOfficePdfCharacterSpacing(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        if (textSpans.Count == 0)
        {
            return textSpans;
        }

        var adjusted = new PptxPositionedTextSpan[textSpans.Count];
        Dictionary<int, (double SpacingEm, int FirstNumberedParagraphIndex)> numberedFrameSpacing = ReadOfficeNumberedFrameCharacterSpacing(textSpans);
        int currentFrame = -1;
        bool useHighlightContinuationSpacing = false;
        for (int i = 0; i < textSpans.Count; i++)
        {
            PptxPositionedTextSpan span = textSpans[i];
            if (span.FrameIndex != currentFrame)
            {
                currentFrame = span.FrameIndex;
                useHighlightContinuationSpacing = false;
            }

            bool hasZeroAuthoredCharacterSpacing =
                Math.Abs(span.GlyphSpan.CharacterSpacing) < PptxTextMetricRules.TextStateTolerance;
            if (hasZeroAuthoredCharacterSpacing && StartsOfficeHighlightContinuationTextState(span))
            {
                useHighlightContinuationSpacing = true;
            }

            bool eligibleNumberedProfileSpan =
                UsesNumberedTextStateProfile(span) &&
                hasZeroAuthoredCharacterSpacing;
            double? pdfCharacterSpacingOverride = null;
            if (hasZeroAuthoredCharacterSpacing && useHighlightContinuationSpacing)
            {
                pdfCharacterSpacingOverride = PptxTextMetricRules.OfficeHighlightContinuationCharacterSpacing(span.Run.FontSize);
            }
            else if (eligibleNumberedProfileSpan &&
                numberedFrameSpacing.TryGetValue(span.FrameIndex, out (double SpacingEm, int FirstNumberedParagraphIndex) frameSpacing) &&
                span.ParagraphIndex >= frameSpacing.FirstNumberedParagraphIndex)
            {
                pdfCharacterSpacingOverride = span.Run.FontSize * frameSpacing.SpacingEm;
            }

            adjusted[i] = pdfCharacterSpacingOverride is not null
                ? span with { PdfCharacterSpacingOverride = pdfCharacterSpacingOverride }
                : span;
        }

        return adjusted;
    }

    private static bool StartsOfficeHighlightContinuationTextState(PptxPositionedTextSpan span) =>
        span.Run.HighlightColor is not null &&
        span.FrameAutofitMode != PptxTextAutofitMode.None &&
        span.ParagraphIndex > 0;

    private static Dictionary<int, (double SpacingEm, int FirstNumberedParagraphIndex)> ReadOfficeNumberedFrameCharacterSpacing(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        var spacingByFrame = new Dictionary<int, (double SpacingEm, int FirstNumberedParagraphIndex)>();
        foreach (IGrouping<int, PptxPositionedTextSpan> frameGroup in textSpans.GroupBy(span => span.FrameIndex))
        {
            PptxPositionedTextSpan[] frame = frameGroup.ToArray();
            if (!frame.Any(UsesNumberedTextStateProfile) ||
                !frame.Any(span => span.ParagraphBulletKind == PptxParagraphBulletKind.AutoNumber) ||
                frame.Any(span => Math.Abs(span.GlyphSpan.CharacterSpacing) >= PptxTextMetricRules.TextStateTolerance))
            {
                continue;
            }

            bool hasExplicitAutoNumberStart = frame.Any(span => span.ParagraphAutoNumberStartAt is not null);
            bool hasBodyContinuation = frame.Any(span =>
                span.ParagraphBulletKind != PptxParagraphBulletKind.AutoNumber);
            double spacingEm = hasExplicitAutoNumberStart && hasBodyContinuation
                ? PptxTextMetricRules.OfficeAutofitNumberedDenseCharacterSpacingEm
                : PptxTextMetricRules.OfficeAutofitNumberedDefaultCharacterSpacingEm;
            int firstNumberedParagraphIndex = frame
                .Where(span => span.ParagraphBulletKind == PptxParagraphBulletKind.AutoNumber)
                .Min(span => span.ParagraphIndex);
            spacingByFrame[frameGroup.Key] = (spacingEm, firstNumberedParagraphIndex);
        }

        return spacingByFrame;
    }

    private static bool UsesNumberedTextStateProfile(PptxPositionedTextSpan span) =>
        span.FrameAutofitMode == PptxTextAutofitMode.Shape ||
        span.FrameAutofitMode == PptxTextAutofitMode.None;

    private static IReadOnlyList<PptxPositionedTextSpan> SplitLeadingSpacesAtHighlightBoundaries(IReadOnlyList<PptxPositionedTextSpan> textSpans)
    {
        if (textSpans.Count < 2)
        {
            return textSpans;
        }

        var split = new List<PptxPositionedTextSpan>(textSpans.Count);
        foreach (PptxPositionedTextSpan span in textSpans)
        {
            int leadingSpaceCount = CountLeadingSpaces(span.Run.Text);
            if (split.Count == 0 ||
                leadingSpaceCount == 0 ||
                leadingSpaceCount >= span.Run.Text.Length ||
                !PreservesHighlightTextOperationBoundaries(split[^1], span) ||
                split[^1].Run.HighlightColor.Equals(span.Run.HighlightColor))
            {
                split.Add(span);
                continue;
            }

            split.AddRange(SplitLeadingSpaceSpan(span, leadingSpaceCount));
        }

        return split;
    }

    private static IEnumerable<PptxPositionedTextSpan> SplitLeadingSpaceSpan(PptxPositionedTextSpan span, int leadingSpaceCount)
    {
        string leadingText = span.Run.Text[..leadingSpaceCount];
        string remainingText = span.Run.Text[leadingSpaceCount..];
        PptxTextGlyphLayout[] leadingGlyphs = span.GlyphSpan.Glyphs.Take(leadingSpaceCount).ToArray();
        PptxTextGlyphLayout[] remainingGlyphs = span.GlyphSpan.Glyphs.Skip(leadingSpaceCount).ToArray();
        double leadingWidth = leadingGlyphs.Length == 0
            ? Math.Min(span.Run.Width, span.GlyphSpan.LayoutWidth * leadingSpaceCount / Math.Max(1, span.Run.Text.Length))
            : GlyphNaturalWidth(leadingGlyphs);
        double remainingLeadingAdjustment = remainingGlyphs.Length == 0 ? 0d : remainingGlyphs[0].AdjustmentBefore;
        if (remainingGlyphs.Length > 0 && Math.Abs(remainingLeadingAdjustment) > PptxTextMetricRules.TextStateTolerance)
        {
            remainingGlyphs[0] = remainingGlyphs[0] with { AdjustmentBefore = 0d };
        }

        TextRun leadingRun = span.Run with
        {
            Text = leadingText,
            Width = Math.Max(0d, leadingWidth),
            PreventCoalesce = true
        };
        yield return span with
        {
            Run = leadingRun,
            EndX = leadingRun.X + leadingWidth,
            Atoms = [new PptxTextAtomLayout(PptxTextAtomKind.Space, leadingText, leadingRun.X, leadingWidth, Draw: true)],
            GlyphSpan = SliceGlyphSpan(span.GlyphSpan, leadingText, leadingWidth, leadingGlyphs)
        };

        double remainingX = span.Run.X + leadingWidth + remainingLeadingAdjustment;
        double remainingWidth = Math.Max(0d, span.EndX - remainingX);
        TextRun remainingRun = span.Run with
        {
            Text = remainingText,
            X = remainingX,
            Width = remainingWidth
        };
        yield return span with
        {
            Run = remainingRun,
            EndX = span.EndX,
            Atoms = [new PptxTextAtomLayout(PptxTextAtomKind.Word, remainingText, remainingRun.X, remainingWidth, Draw: true)],
            GlyphSpan = SliceGlyphSpan(span.GlyphSpan, remainingText, remainingWidth, remainingGlyphs)
        };
    }

    private static PptxTextGlyphSpanLayout SliceGlyphSpan(PptxTextGlyphSpanLayout source, string text, double width, IReadOnlyList<PptxTextGlyphLayout> glyphs)
    {
        double naturalWidth = GlyphNaturalWidth(glyphs);
        return source with
        {
            Text = text,
            NaturalWidth = naturalWidth,
            LayoutWidth = width,
            Glyphs = glyphs
        };
    }

    private static double GlyphNaturalWidth(IReadOnlyList<PptxTextGlyphLayout> glyphs)
    {
        return Math.Max(0d, glyphs.Sum(glyph => glyph.Advance) + glyphs.Sum(glyph => glyph.AdjustmentBefore));
    }
}
