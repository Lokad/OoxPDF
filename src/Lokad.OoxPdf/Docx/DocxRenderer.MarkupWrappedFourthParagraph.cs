using System.Text;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private sealed record DocxMarkupWrappedMixedFourthRows(
        DocxMarkupFittingMixedFourthParagraph Paragraphs,
        DocxUniformBalloonRow[][] LeadingRows, double[] Pitches, double[] Transitions,
        DocxUniformBalloonRow[] FourthRows, double FourthPitch, double FinalPitch)
    {
        public double ContinuationsHeight => Enumerable.Range(0, 3)
            .Sum(index => (LeadingRows[index].Length - 1) * Pitches[index]) + Transitions.Sum() +
            (FourthRows.Length - 2) * FourthPitch + FinalPitch;
    }

    private static DocxMarkupWrappedMixedFourthRows? ResolveWordCompatibleWrappedMixedFourthRows(
        DocxMarkupFittingMixedFourthParagraph? paragraphs, double size, double firstWidth,
        double continuationWidth, CancellationToken cancellationToken)
    {
        if (paragraphs is null || !double.IsFinite(firstWidth) || firstWidth <= 0d ||
            !double.IsFinite(continuationWidth) || continuationWidth <= 0d) { return null; }
        var rows = new DocxUniformBalloonRow[3][];
        var pitches = new double[3];
        var transitions = new double[3];
        for (int index = 0; index < 3; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var body = paragraphs.Leading[index].Body;
            if (!HasValidWrappedFourthBody(body, size, continuationWidth, cancellationToken)) { return null; }
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
        if (!HasValidWrappedFourthBody(first, size, continuationWidth, cancellationToken) ||
            !HasValidWrappedFourthBody(closing, size, continuationWidth, cancellationToken)) { return null; }
        // Trimming or wrapping the closing run would change its authored separation.
        if (first.Text.EndsWith(' ')) { return null; }
        var fourthRows = WrapUniformBalloonWords(first.Text, first.Resource.Embedded, size,
            continuationWidth, continuationWidth, cancellationToken, reserveFirstRowBreakSpace: false);
        if (fourthRows.Length < 2 || fourthRows.Any(row => row.Text.Length == 0)) { return null; }
        double lastWidth = first.Resource.Embedded.MeasureTextPoints(fourthRows[^1].Text, size) +
            closing.Resource.Embedded.MeasureTextPoints(closing.Text, size);
        if (!double.IsFinite(lastWidth)) { return null; }
        if (lastWidth > continuationWidth)
        {
            string last = fourthRows[^1].Text;
            int boundary = last.LastIndexOf(' ');
            if (boundary <= 0) { return null; }
            string word = last[(boundary + 1)..];
            double joinedWidth = first.Resource.Embedded.MeasureTextPoints(word, size) +
                closing.Resource.Embedded.MeasureTextPoints(closing.Text, size);
            if (!double.IsFinite(joinedWidth) || joinedWidth > continuationWidth) { return null; }
            fourthRows = [.. fourthRows[..^1], new(last[..boundary], true), new(word, true)];
        }
        var firstFont = first.Resource.Embedded.Font;
        var closingFont = closing.Resource.Embedded.Font;
        var previous = paragraphs.Leading[2].Body.Resource.Embedded.Font;
        double firstAscent = (firstFont.Hhea.HorizontalAscender + firstFont.Hhea.HorizontalLineGap) / (double)firstFont.UnitsPerEm;
        double closingAscent = (closingFont.Hhea.HorizontalAscender + closingFont.Hhea.HorizontalLineGap) / (double)closingFont.UnitsPerEm;
        double firstStep = (-previous.Hhea.HorizontalDescender / (double)previous.UnitsPerEm + firstAscent) * size;
        double finalPitch = (-firstFont.Hhea.HorizontalDescender / (double)firstFont.UnitsPerEm + Math.Max(firstAscent, closingAscent)) * size;
        double fourthPitch = DocxLineMetrics.MeasureHheaLineHeight(firstFont, size);
        if (!double.IsFinite(firstStep) || firstStep <= 0d || !double.IsFinite(finalPitch) || finalPitch <= 0d) { return null; }
        transitions[2] = firstStep;
        var result = new DocxMarkupWrappedMixedFourthRows(paragraphs, rows, pitches, transitions, fourthRows, fourthPitch, finalPitch);
        return double.IsFinite(result.ContinuationsHeight) ? result : null;
    }

    private static bool HasValidWrappedFourthBody(DocxMarkupBalloonBodyPart body, double size,
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

    private static void RenderWordCompatibleWrappedMixedFourthRows(
        DocxMarkupWrappedMixedFourthRows rows, DocxMarkupBalloonPlacement placement, PdfGraphicsBuilder graphics,
        double firstX, double continuationX, double firstY, double size, CancellationToken cancellationToken)
    {
        double y = firstY;
        for (int index = 0; index < 3; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var paragraph = rows.Paragraphs.Leading[index];
            RenderUniformBalloonRows(rows.LeadingRows[index], placement, graphics, paragraph.Body.Resource, paragraph.Mark,
                index == 0 ? firstX : continuationX, continuationX, y, rows.Pitches[index], size, cancellationToken);
            y -= (rows.LeadingRows[index].Length - 1) * rows.Pitches[index] + rows.Transitions[index];
        }
        var first = rows.Paragraphs.Parts[0];
        var closing = rows.Paragraphs.Parts[1];
        for (int index = 0; index < rows.FourthRows.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            bool last = index == rows.FourthRows.Length - 1;
            if (index > 0) { y -= last ? rows.FinalPitch : rows.FourthPitch; }
            string text = rows.FourthRows[index].Text + (!last && rows.FourthRows[index].SpaceAfter ? " " : string.Empty);
            DrawBalloonText(graphics, first.Resource, text, continuationX, y, size,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            if (last)
            {
                double x = continuationX + first.Resource.Embedded.MeasureTextPoints(text, size);
                DrawBalloonText(graphics, closing.Resource, closing.Text, x, y, size,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
                x += closing.Resource.Embedded.MeasureTextPoints(closing.Text, size);
                DrawBalloonText(graphics, rows.Paragraphs.Mark, " ", x, y, size,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            }
        }
    }
}
