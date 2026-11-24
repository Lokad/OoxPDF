using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private sealed record DocxMarkupWrappedTerminalRows(
        DocxUniformBalloonRow[] TerminalRows, double TerminalPitch, double FinalPitch, bool ClosingOwnRow,
        string? PartialClosingPrefix, string? PartialClosingSuffix, double MixedTerminalPitch)
    {
        public bool HasClosingRow => ClosingOwnRow || PartialClosingSuffix is not null;

    }

    private static DocxMarkupWrappedTerminalRows? ResolveWordCompatibleWrappedTerminalRows(
        DocxMarkupMixedTerminalParagraph paragraphs, double size,
        double continuationWidth, CancellationToken cancellationToken, out double firstStep)
    {
        firstStep = 0d;
        int leadingCount = paragraphs.Leading.Count;
        var first = paragraphs.Parts[0];
        var closing = paragraphs.Parts[1];
        if (!TryMeasureBalloonParagraphBody(first, size, continuationWidth, cancellationToken, out double terminalPitch) ||
            !TryMeasureBalloonParagraphBody(closing, size, continuationWidth, cancellationToken, out _)) { return null; }
        // Trimming or wrapping the closing run would change its authored separation.
        if (first.Text.EndsWith(' ')) { return null; }
        var terminalRows = WrapUniformBalloonWords(first.Text, first.Resource.Embedded, size,
            continuationWidth, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
        if (terminalRows.Length < 2 || terminalRows.Any(row => row.Text.Length == 0)) { return null; }
        double lastWidth = first.Resource.Embedded.MeasureTextPoints(terminalRows[^1].Text, size) +
            closing.Resource.Embedded.MeasureTextPoints(closing.Text, size);
        if (!double.IsFinite(lastWidth)) { return null; }
        bool closingOwnRow = false;
        string? partialClosingPrefix = null;
        string? partialClosingSuffix = null;
        if (lastWidth > continuationWidth && leadingCount is (1 or 2 or 3 or 4) && closing.Text.StartsWith(' '))
        {
            // Office preserves breakable closing separators on the preceding row.
            // A qualified fitting prefix stays there; the remaining closing text moves together.
            string closingWord = closing.Text[1..];
            double closingWidth = closing.Resource.Embedded.MeasureTextPoints(closingWord, size);
            double precedingWidth = first.Resource.Embedded.MeasureTextPoints(terminalRows[^1].Text, size) +
                closing.Resource.Embedded.MeasureTextPoints(" ", size);
            int wordSeparator = closingWord.IndexOf(' ');
            bool hasTwoWords = leadingCount == 4 && wordSeparator > 0 && wordSeparator < closingWord.Length - 1 &&
                closingWord.LastIndexOf(' ') == wordSeparator;
            int secondSeparator = wordSeparator < 0 ? -1 : closingWord.IndexOf(' ', wordSeparator + 1);
            bool hasThreeWords = leadingCount == 4 && wordSeparator > 0 && secondSeparator > wordSeparator + 1 &&
                secondSeparator < closingWord.Length - 1 && closingWord.LastIndexOf(' ') == secondSeparator;
            bool hasFittingClosingWords = false;
            if (hasTwoWords || hasThreeWords)
            {
                double firstClosingWordWidth = first.Resource.Embedded.MeasureTextPoints(terminalRows[^1].Text, size) +
                    closing.Resource.Embedded.MeasureTextPoints(closing.Text[..(wordSeparator + 1)], size);
                hasFittingClosingWords = double.IsFinite(firstClosingWordWidth) && firstClosingWordWidth > continuationWidth;
                double secondClosingWordWidth = hasThreeWords ? first.Resource.Embedded.MeasureTextPoints(terminalRows[^1].Text, size) +
                    closing.Resource.Embedded.MeasureTextPoints(closing.Text[..(secondSeparator + 1)], size) : double.PositiveInfinity;
                if (double.IsFinite(firstClosingWordWidth) && firstClosingWordWidth <= continuationWidth &&
                    (!hasThreeWords || double.IsFinite(secondClosingWordWidth)) &&
                    double.IsFinite(closingWidth) && closingWidth <= continuationWidth &&
                    double.IsFinite(precedingWidth) && precedingWidth <= continuationWidth)
                {
                    int splitSeparator = hasThreeWords && secondClosingWordWidth <= continuationWidth ? secondSeparator : wordSeparator;
                    partialClosingPrefix = closing.Text[..(splitSeparator + 2)];
                    partialClosingSuffix = closingWord[(splitSeparator + 1)..];
                }
            }
            bool canMoveClosingWord = closingWord.Length != 0 && (!closingWord.Contains(' ') || hasFittingClosingWords) && double.IsFinite(closingWidth) &&
                closingWidth <= continuationWidth && double.IsFinite(precedingWidth) && precedingWidth <= continuationWidth;
            if (!canMoveClosingWord && partialClosingSuffix is null && leadingCount is (1 or 2 or 4)) { return null; }
            // Fourth-paragraph inputs outside the closing-only scope retain their prior reflow.
            closingOwnRow = canMoveClosingWord;
        }
        if (lastWidth > continuationWidth && !closingOwnRow && partialClosingSuffix is null)
        {
            string last = terminalRows[^1].Text;
            int boundary = last.LastIndexOf(' ');
            if (boundary <= 0) { return null; }
            string word = last[(boundary + 1)..];
            double joinedWidth = first.Resource.Embedded.MeasureTextPoints(word, size) +
                closing.Resource.Embedded.MeasureTextPoints(closing.Text, size);
            if (!double.IsFinite(joinedWidth) || joinedWidth > continuationWidth) { return null; }
            terminalRows = [.. terminalRows[..^1], new(last[..boundary], true), new(word, true)];
        }
        var firstFont = first.Resource.Embedded.Font;
        var closingFont = closing.Resource.Embedded.Font;
        var previous = paragraphs.Leading[leadingCount - 1].Body.Resource.Embedded.Font;
        double firstAscent = (firstFont.Hhea.HorizontalAscender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm;
        double closingAscent = (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm;
        firstStep = (-previous.Hhea.HorizontalDescender / (double)previous.UnitsPerEm + firstAscent) * size;
        double finalPitch = (-firstFont.Hhea.HorizontalDescender / (double)firstFont.UnitsPerEm + (closingOwnRow || partialClosingSuffix is not null ? closingAscent : Math.Max(firstAscent, closingAscent))) * size;
        if (!double.IsFinite(firstStep) || firstStep <= 0d || !double.IsFinite(finalPitch) || finalPitch <= 0d) { return null; }
        double mixedTerminalPitch = (-firstFont.Hhea.HorizontalDescender / (double)firstFont.UnitsPerEm +
            Math.Max(firstAscent, closingAscent)) * size;
        if (!double.IsFinite(mixedTerminalPitch) || mixedTerminalPitch <= 0d) { return null; }
        return new DocxMarkupWrappedTerminalRows(terminalRows, terminalPitch, finalPitch, closingOwnRow,
            partialClosingPrefix, partialClosingSuffix, mixedTerminalPitch);
    }

    private static void RenderWordCompatibleWrappedTerminalRows(
        DocxMarkupMixedTerminalParagraph paragraphs, DocxMarkupWrappedTerminalRows rows,
        DocxMarkupBalloonPlacement placement, PdfGraphicsBuilder graphics,
        double continuationX, double y, double size, CancellationToken cancellationToken)
    {
        var first = paragraphs.Parts[0];
        var closing = paragraphs.Parts[1];
        for (int index = 0; index < rows.TerminalRows.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            bool last = index == rows.TerminalRows.Length - 1;
            if (index > 0)
            {
                y -= last && rows.PartialClosingSuffix is not null ? rows.MixedTerminalPitch :
                    last && !rows.ClosingOwnRow ? rows.FinalPitch : rows.TerminalPitch;
            }
            string text = rows.TerminalRows[index].Text + (!last && rows.TerminalRows[index].SpaceAfter ? " " : string.Empty);
            DrawBalloonText(graphics, first.Resource, text, continuationX, y, size,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            if (last)
            {
                double x = continuationX + first.Resource.Embedded.MeasureTextPoints(text, size);
                if (rows.HasClosingRow)
                {
                    DrawBalloonText(graphics, closing.Resource, rows.PartialClosingPrefix ?? " ", x, y, size,
                        placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
                    y -= rows.FinalPitch;
                    x = continuationX;
                }
                string closingText = rows.PartialClosingSuffix ?? (rows.ClosingOwnRow ? closing.Text[1..] : closing.Text);
                DrawBalloonText(graphics, closing.Resource, closingText, x, y, size,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
                x += closing.Resource.Embedded.MeasureTextPoints(closingText, size);
                DrawBalloonText(graphics, paragraphs.Mark, " ", x, y, size,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            }
        }
    }
}
