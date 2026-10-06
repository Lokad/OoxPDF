using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private readonly record struct DocxUniformBalloonRow(string Text, bool SpaceAfter);

    private static DocxUniformBalloonRow[] WrapUniformBalloonWords(
        string text, PdfEmbeddedFont embedded, double fontSize,
        double firstLineWidth, double continuationWidth, CancellationToken cancellationToken,
        bool reserveFirstRowBreakSpace = true)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string normalized = Regex.Replace(text.Trim(), @"\s+", " ");
        if (normalized.Length == 0) { return []; }
        string[] words = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var rows = new List<DocxUniformBalloonRow>();
        int wordIndex = 0, wordOffset = 0;
        ushort space = embedded.Font.MapCodePoint(' ');
        double pointScale = fontSize / embedded.Font.UnitsPerEm;
        while (wordIndex < words.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            bool firstRow = rows.Count == 0;
            double maxWidth = firstRow ? firstLineWidth : continuationWidth;
            var line = new StringBuilder();
            double lineUnits = 0d;
            ushort previous = 0;
            bool spaceAfter = true;
            while (wordIndex < words.Length)
            {
                string word = words[wordIndex];
                bool hasEarlierWord = line.Length != 0;
                double initialUnits = lineUnits;
                ushort initialPrevious = previous;
                if (hasEarlierWord)
                {
                    initialUnits += embedded.Font.GetAdvanceWidth(space) + embedded.Font.GetKerning(previous, space);
                    initialPrevious = space;
                }
                int length = FitPrefix(word, wordOffset, initialUnits, initialPrevious, maxWidth,
                    out double fittedUnits, out ushort lastGlyph);
                bool completeWord = wordOffset + length == word.Length;
                double withBreakSpace = fittedUnits + embedded.Font.GetAdvanceWidth(space) + embedded.Font.GetKerning(lastGlyph, space);
                if (hasEarlierWord && (!completeWord || reserveFirstRowBreakSpace && firstRow && wordIndex + 1 < words.Length && withBreakSpace * pointScale > maxWidth))
                {
                    break;
                }
                if (!completeWord)
                {
                    if (length == 0 && firstRow)
                    {
                        // A title may consume the entire first row. Continue below it.
                        spaceAfter = false;
                        break;
                    }
                    if (length == 0)
                    {
                        // Preserve progress/content even if one glyph is wider than the lane.
                        length = Rune.GetRuneAt(word, wordOffset).Utf16SequenceLength;
                    }
                    line.Append(word.AsSpan(wordOffset, length));
                    wordOffset += length;
                    spaceAfter = false;
                    if (wordOffset == word.Length)
                    {
                        wordIndex++;
                        wordOffset = 0;
                        spaceAfter = true;
                    }
                    break;
                }
                if (hasEarlierWord) { line.Append(' '); }
                line.Append(word.AsSpan(wordOffset));
                lineUnits = fittedUnits;
                previous = lastGlyph;
                wordOffset = 0;
                wordIndex++;
            }
            rows.Add(new(line.ToString(), spaceAfter));
        }
        return rows.ToArray();

        int FitPrefix(string word, int offset, double units, ushort prior, double width,
            out double fittedUnits, out ushort lastGlyph)
        {
            int length = 0, count = 0;
            foreach (Rune rune in word.AsSpan(offset).EnumerateRunes())
            {
                if ((count++ & 255) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                ushort glyph = embedded.Font.MapCodePoint(rune.Value);
                double next = units + embedded.Font.GetAdvanceWidth(glyph) +
                    (prior == 0 || glyph == 0 ? 0d : embedded.Font.GetKerning(prior, glyph));
                if (next * pointScale > width) { break; }
                units = next;
                prior = glyph;
                length += rune.Utf16SequenceLength;
            }
            fittedUnits = units;
            lastGlyph = prior;
            return length;
        }
    }

    private static void RenderUniformBalloonRows(
        DocxUniformBalloonRow[] rows, DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics, DocxRunFontResource body, DocxRunFontResource terminal,
        double firstX, double continuationX, double firstY, double gap, double fontSize,
        CancellationToken cancellationToken, double? firstContinuationGap = null)
    {
        for (int i = 0; i < rows.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            double x = i == 0 ? firstX : continuationX;
            double y = i > 0 && firstContinuationGap is double firstGap
                ? firstY - firstGap - (i - 1) * gap : firstY - i * gap;
            string text = i == 0 && rows.Length > 1 && rows[i].SpaceAfter && rows[i].Text.Length != 0
                ? rows[i].Text + " " : rows[i].Text;
            DrawBalloonText(graphics, body, text, x, y, fontSize,
                placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            if ((i > 0 || rows.Length == 1) && rows[i].SpaceAfter)
            {
                DrawBalloonText(graphics, i == rows.Length - 1 ? terminal : body, " ",
                    x + body.Embedded.MeasureTextPoints(text, fontSize), y, fontSize,
                    placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            }
        }
    }
}
