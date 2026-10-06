using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static DocxMarkupTwoFaceRows? ResolveWordCompatibleClosingRunRows(
        IReadOnlyList<DocxMarkupBalloonBodyPart> parts, double fontSize,
        double firstLineWidth, double continuationWidth, CancellationToken cancellationToken)
    {
        DocxMarkupBalloonBodyPart prefix = parts[0], tail = parts[1], closing = parts[2];
        if (!ReferenceEquals(prefix.Resource.Embedded.Font, closing.Resource.Embedded.Font) ||
            prefix.Text.Length == 0 || tail.Text.Length == 0 || closing.Text.Length == 0 ||
            tail.Text.StartsWith(' ') || tail.Text.EndsWith(' ') || tail.Text.Contains("  ", StringComparison.Ordinal)) { return null; }
        int split = prefix.Text.IndexOf(' ');
        if (split <= 0 || split != prefix.Text.LastIndexOf(' ') || split == prefix.Text.Length - 1) { return null; }
        for (int index = 0; index < prefix.Text.Length; index++)
        {
            if ((index & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            if (prefix.Text[index] != ' ' && !char.IsAsciiLetterOrDigit(prefix.Text[index])) { return null; }
        }
        for (int index = 0; index < closing.Text.Length; index++)
        {
            if ((index & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
            char value = closing.Text[index];
            if (!char.IsAsciiLetterOrDigit(value) && !(value == '.' && index == closing.Text.Length - 1 && index > 0)) { return null; }
        }
        double available = firstLineWidth - prefix.Resource.Embedded.MeasureTextPoints(prefix.Text, fontSize);
        if (!double.IsFinite(available) || available <= 0d || !double.IsFinite(continuationWidth) || continuationWidth <= 0d) { return null; }
        string[] words = tail.Text.Split(' ');
        if (words.Length < 3) { return null; }
        for (int index = 0; index < words.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            double width = tail.Resource.Embedded.MeasureTextPoints(words[index], fontSize);
            if (!double.IsFinite(width) || width <= 0d || width > continuationWidth) { return null; }
        }
        double firstWord = tail.Resource.Embedded.MeasureTextPoints(words[0] + " ", fontSize);
        if (!double.IsFinite(firstWord) || firstWord > available) { return null; }
        DocxUniformBalloonRow[] lines = WrapUniformBalloonWords(tail.Text, tail.Resource.Embedded, fontSize,
            available, continuationWidth, cancellationToken);
        if (lines.Length < 3 || lines[0].Text.Length == 0 || lines.Any(line => !line.SpaceAfter)) { return null; }
        double lastWidth = tail.Resource.Embedded.MeasureTextPoints(lines[^1].Text, fontSize) +
            closing.Resource.Embedded.MeasureTextPoints(closing.Text, fontSize);
        if (!double.IsFinite(lastWidth)) { return null; }
        if (lastWidth > continuationWidth)
        {
            // The closing run belongs to the last tail word. When their joined
            // advance no longer fits, move that whole word to a new final row;
            // its authored preceding separator remains in the tail face.
            string last = lines[^1].Text;
            int boundary = last.LastIndexOf(' ');
            if (boundary <= 0) { return null; }
            string word = last[(boundary + 1)..];
            double joinedWidth = tail.Resource.Embedded.MeasureTextPoints(word, fontSize) +
                closing.Resource.Embedded.MeasureTextPoints(closing.Text, fontSize);
            if (!double.IsFinite(joinedWidth) || joinedWidth > continuationWidth) { return null; }
            lines = [.. lines[..^1], new(last[..boundary], true), new(word, true)];
        }
        var prefixFont = prefix.Resource.Embedded.Font;
        var tailFont = tail.Resource.Embedded.Font;
        if (prefixFont.UnitsPerEm <= 0 || tailFont.UnitsPerEm <= 0) { return null; }
        double prefixDescent = -prefixFont.Hhea.HorizontalDescender / (double)prefixFont.UnitsPerEm;
        double tailDescent = -tailFont.Hhea.HorizontalDescender / (double)tailFont.UnitsPerEm;
        double prefixAscent = (prefixFont.Hhea.HorizontalAscender + prefixFont.Hhea.HorizontalLineGap) / (double)prefixFont.UnitsPerEm;
        double tailAscent = (tailFont.Hhea.HorizontalAscender + tailFont.Hhea.HorizontalLineGap) / (double)tailFont.UnitsPerEm;
        double firstGap = Math.Max(prefixDescent, tailDescent) + tailAscent;
        double tailGap = tailDescent + tailAscent;
        double lastGap = tailDescent + Math.Max(prefixAscent, tailAscent);
        if (!double.IsFinite(firstGap) || firstGap <= 0d || !double.IsFinite(tailGap) || tailGap <= 0d ||
            !double.IsFinite(lastGap) || lastGap <= 0d) { return null; }
        return new(prefix, tail.Resource, lines, tailGap, firstGap, ClosingPart: closing, FinalContinuationGapEm: lastGap);
    }

    private static void RenderWordCompatibleClosingRunRows(
        DocxMarkupTwoFaceRows rows, DocxMarkupBalloonPlacement placement, PdfGraphicsBuilder graphics,
        DocxRunFontResource terminalResource, double textX, double firstX, double firstY,
        double fontSize, CancellationToken cancellationToken)
    {
        DrawBalloonText(graphics, rows.Prefix.Resource, rows.Prefix.Text, firstX, firstY, fontSize,
            placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
        double tailX = firstX + rows.Prefix.Resource.Embedded.MeasureTextPoints(rows.Prefix.Text, fontSize);
        double pureGap = rows.TailLineHeightEm!.Value * fontSize;
        for (int index = 0; index < rows.TailLines.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            bool last = index == rows.TailLines.Length - 1;
            double x = index == 0 ? tailX : textX;
            double y = index == 0 ? firstY : firstY - rows.FirstContinuationGapEm!.Value * fontSize - (index - 1) * pureGap;
            if (last) { y -= rows.FinalContinuationGapEm!.Value * fontSize - pureGap; }
            string text = rows.TailLines[index].Text + (last ? string.Empty : " ");
            DrawBalloonText(graphics, rows.TailResource, text, x, y, fontSize,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            if (last)
            {
                DocxMarkupBalloonBodyPart closing = rows.ClosingPart!;
                double closingX = x + rows.TailResource.Embedded.MeasureTextPoints(text, fontSize);
                DrawBalloonText(graphics, closing.Resource, closing.Text, closingX, y, fontSize,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
                DrawBalloonText(graphics, terminalResource, " ", closingX + closing.Resource.Embedded.MeasureTextPoints(closing.Text, fontSize),
                    y, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            }
        }
    }
}
