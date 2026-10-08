using System.Text;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private sealed record DocxMarkupWrappedMixedTerminalRows(
        DocxMarkupMixedTerminalParagraph Paragraphs,
        DocxUniformBalloonRow[][] LeadingRows, double[] Pitches, double[] Transitions,
        DocxUniformBalloonRow[] TerminalRows, double TerminalPitch, double FinalPitch, bool ClosingOwnRow)
    {
        public double ContinuationsHeight => Enumerable.Range(0, Paragraphs.Leading.Count)
            .Sum(index => (LeadingRows[index].Length - 1) * Pitches[index]) + Transitions.Sum() +
            (TerminalRows.Length - (ClosingOwnRow ? 1 : 2)) * TerminalPitch + FinalPitch;
    }

    private static DocxMarkupWrappedMixedTerminalRows? ResolveWordCompatibleWrappedMixedTerminalRows(
        DocxMarkupMixedTerminalParagraph? paragraphs, double size, double firstWidth,
        double continuationWidth, CancellationToken cancellationToken)
    {
        if (paragraphs is null || paragraphs.Leading.Count is not (1 or 2 or 3) || !double.IsFinite(firstWidth) || firstWidth <= 0d ||
            !double.IsFinite(continuationWidth) || continuationWidth <= 0d) { return null; }
        int leadingCount = paragraphs.Leading.Count;
        var rows = new DocxUniformBalloonRow[leadingCount][];
        var pitches = new double[leadingCount];
        var transitions = new double[leadingCount];
        for (int index = 0; index < leadingCount; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var body = paragraphs.Leading[index].Body;
            if (!HasValidWrappedTerminalBody(body, size, continuationWidth, cancellationToken)) { return null; }
            rows[index] = WrapUniformBalloonWords(body.Text, body.Resource.Embedded, size,
                index == 0 ? firstWidth : continuationWidth, continuationWidth, cancellationToken,
                reserveFirstRowBreakSpace: false);
            if (rows[index].Length == 0 || rows[index][0].Text.Length == 0) { return null; }
            pitches[index] = DocxLineMetrics.MeasureHheaLineHeight(body.Resource.Embedded.Font, size);
            if (index > 0)
            {
                double? step = MeasureWordCompatibleParagraphTransition(paragraphs.Leading[index - 1], paragraphs.Leading[index], size);
                if (step is null) { return null; }
                transitions[index - 1] = step.Value;
            }
        }
        var first = paragraphs.Parts[0];
        var closing = paragraphs.Parts[1];
        if (!HasValidWrappedTerminalBody(first, size, continuationWidth, cancellationToken) ||
            !HasValidWrappedTerminalBody(closing, size, continuationWidth, cancellationToken)) { return null; }
        // Trimming or wrapping the closing run would change its authored separation.
        if (first.Text.EndsWith(' ')) { return null; }
        var terminalRows = WrapUniformBalloonWords(first.Text, first.Resource.Embedded, size,
            continuationWidth, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
        if (terminalRows.Length < 2 || terminalRows.Any(row => row.Text.Length == 0)) { return null; }
        double lastWidth = first.Resource.Embedded.MeasureTextPoints(terminalRows[^1].Text, size) +
            closing.Resource.Embedded.MeasureTextPoints(closing.Text, size);
        if (!double.IsFinite(lastWidth)) { return null; }
        bool closingOwnRow = false;
        if (lastWidth > continuationWidth && leadingCount is (1 or 2 or 3) && closing.Text.StartsWith(' '))
        {
            // Office keeps the breakable separator in its closing face on the
            // preceding row, then moves the complete closing word to its own row.
            string closingWord = closing.Text[1..];
            double closingWidth = closing.Resource.Embedded.MeasureTextPoints(closingWord, size);
            double precedingWidth = first.Resource.Embedded.MeasureTextPoints(terminalRows[^1].Text, size) +
                closing.Resource.Embedded.MeasureTextPoints(" ", size);
            bool canMoveClosingWord = closingWord.Length != 0 && !closingWord.Contains(' ') && double.IsFinite(closingWidth) &&
                closingWidth <= continuationWidth && double.IsFinite(precedingWidth) && precedingWidth <= continuationWidth;
            if (!canMoveClosingWord && leadingCount is (1 or 2)) { return null; }
            // Fourth-paragraph inputs outside the closing-only scope retain their prior reflow.
            closingOwnRow = canMoveClosingWord;
        }
        if (lastWidth > continuationWidth && !closingOwnRow)
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
        double firstStep = (-previous.Hhea.HorizontalDescender / (double)previous.UnitsPerEm + firstAscent) * size;
        double finalPitch = (-firstFont.Hhea.HorizontalDescender / (double)firstFont.UnitsPerEm + (closingOwnRow ? closingAscent : Math.Max(firstAscent, closingAscent))) * size;
        double terminalPitch = DocxLineMetrics.MeasureHheaLineHeight(firstFont, size);
        if (!double.IsFinite(firstStep) || firstStep <= 0d || !double.IsFinite(finalPitch) || finalPitch <= 0d) { return null; }
        transitions[leadingCount - 1] = firstStep;
        var result = new DocxMarkupWrappedMixedTerminalRows(paragraphs, rows, pitches, transitions, terminalRows, terminalPitch, finalPitch, closingOwnRow);
        return double.IsFinite(result.ContinuationsHeight) ? result : null;
    }

    private static bool HasValidWrappedTerminalBody(DocxMarkupBalloonBodyPart body, double size,
        double continuationWidth, CancellationToken cancellationToken)
    {
        var font = body.Resource.Embedded.Font;
        double width = body.Resource.Embedded.MeasureTextPoints(body.Text, size);
        double pitch = DocxLineMetrics.MeasureHheaLineHeight(font, size);
        if (font.UnitsPerEm <= 0 || !double.IsFinite(width) || width <= 0d || !double.IsFinite(pitch) || pitch <= 0d) { return false; }
        int index = 0;
        foreach (Rune rune in body.Text.EnumerateRunes())
        {
            if ((index++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            double scalarWidth = font.GetAdvanceWidth(font.MapCodePoint(rune.Value)) * size / font.UnitsPerEm;
            if (!double.IsFinite(scalarWidth) || scalarWidth > continuationWidth) { return false; }
        }
        return true;
    }

    private static void RenderWordCompatibleWrappedMixedTerminalRows(
        DocxMarkupWrappedMixedTerminalRows rows, DocxMarkupBalloonPlacement placement, PdfGraphicsBuilder graphics,
        double firstX, double continuationX, double firstY, double size, CancellationToken cancellationToken)
    {
        double y = firstY;
        for (int index = 0; index < rows.Paragraphs.Leading.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var paragraph = rows.Paragraphs.Leading[index];
            RenderUniformBalloonRows(rows.LeadingRows[index], placement, graphics, paragraph.Body.Resource, paragraph.Mark,
                index == 0 ? firstX : continuationX, continuationX, y, rows.Pitches[index], size, cancellationToken);
            y -= (rows.LeadingRows[index].Length - 1) * rows.Pitches[index] + rows.Transitions[index];
        }
        var first = rows.Paragraphs.Parts[0];
        var closing = rows.Paragraphs.Parts[1];
        for (int index = 0; index < rows.TerminalRows.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            bool last = index == rows.TerminalRows.Length - 1;
            if (index > 0) { y -= last && !rows.ClosingOwnRow ? rows.FinalPitch : rows.TerminalPitch; }
            string text = rows.TerminalRows[index].Text + (!last && rows.TerminalRows[index].SpaceAfter ? " " : string.Empty);
            DrawBalloonText(graphics, first.Resource, text, continuationX, y, size,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            if (last)
            {
                double x = continuationX + first.Resource.Embedded.MeasureTextPoints(text, size);
                if (rows.ClosingOwnRow)
                {
                    DrawBalloonText(graphics, closing.Resource, " ", x, y, size,
                        placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
                    y -= rows.FinalPitch;
                    x = continuationX;
                }
                string closingText = rows.ClosingOwnRow ? closing.Text[1..] : closing.Text;
                DrawBalloonText(graphics, closing.Resource, closingText, x, y, size,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
                x += closing.Resource.Embedded.MeasureTextPoints(closingText, size);
                DrawBalloonText(graphics, rows.Paragraphs.Mark, " ", x, y, size,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            }
        }
    }
}
