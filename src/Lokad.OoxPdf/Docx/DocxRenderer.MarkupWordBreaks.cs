using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private readonly record struct DocxUniformBalloonRow(string Text, bool SpaceAfter);

    private static DocxUniformBalloonRow[] WrapUniformBalloonWords(
        string text, PdfEmbeddedFont? embedded, double fontSize,
        double firstLineWidth, double continuationWidth, CancellationToken cancellationToken,
        bool reserveFirstRowBreakSpace = true, bool preserveAsciiSpaces = false, PdfFallbackFont? fallback = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string normalized = preserveAsciiSpaces ? text.Trim(' ') : Regex.Replace(text.Trim(), @"\s+", " ");
        if (normalized.Length == 0) { return []; }
        string[] words = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        int[]? separators = null;
        if (preserveAsciiSpaces)
        {
            // Keep each authored separator's width and emitted space count.
            separators = new int[words.Length];
            int offset = 0;
            for (int index = 0; index < words.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                offset += words[index].Length;
                int start = offset;
                while (offset < normalized.Length && normalized[offset] == ' ') { offset++; }
                separators[index] = index == words.Length - 1 ? 1 : offset - start;
            }
        }
        var rows = new List<DocxUniformBalloonRow>();
        int wordIndex = 0, wordOffset = 0;
        var font = embedded?.Font;
        ushort space = font?.MapCodePoint(' ') ?? 0;
        double spaceAdvance = font is null ? fallback!.MeasureAdvanceEm(' ') : font.GetAdvanceWidth(space);
        double pointScale = fontSize / (font?.UnitsPerEm ?? PdfFallbackFont.UnitsPerEm);
        while (wordIndex < words.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            bool firstRow = rows.Count == 0;
            double maxWidth = firstRow ? firstLineWidth : continuationWidth;
            var line = new StringBuilder();
            double lineUnits = 0d;
            ushort previous = 0;
            bool spaceAfter = true;
            int trailingSpaces = 1;
            while (wordIndex < words.Length)
            {
                string word = words[wordIndex];
                bool hasEarlierWord = line.Length != 0;
                double initialUnits = lineUnits;
                ushort initialPrevious = previous;
                if (hasEarlierWord)
                {
                    int count = separators?[wordIndex - 1] ?? 1;
                    for (int index = 0; index < count; index++)
                    {
                        initialUnits += spaceAdvance + (font?.GetKerning(initialPrevious, space) ?? 0);
                        initialPrevious = space;
                    }
                }
                int length = FitPrefix(word, wordOffset, initialUnits, initialPrevious, maxWidth,
                    out double fittedUnits, out ushort lastGlyph);
                bool completeWord = wordOffset + length == word.Length;
                double withBreakSpace = fittedUnits + spaceAdvance + (font?.GetKerning(lastGlyph, space) ?? 0);
                int breakSpaces = separators?[wordIndex] ?? 1;
                for (int index = 1; index < breakSpaces; index++)
                {
                    withBreakSpace += spaceAdvance + (font?.GetKerning(space, space) ?? 0);
                }
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
                        trailingSpaces = breakSpaces;
                        wordIndex++;
                        wordOffset = 0;
                        spaceAfter = true;
                    }
                    break;
                }
                if (hasEarlierWord) { line.Append(' ', separators?[wordIndex - 1] ?? 1); }
                line.Append(word.AsSpan(wordOffset));
                lineUnits = fittedUnits;
                previous = lastGlyph;
                trailingSpaces = breakSpaces;
                wordOffset = 0;
                wordIndex++;
            }
            // SpaceAfter emits one separator; retain the rest in the body row.
            if (spaceAfter && line.Length != 0 && trailingSpaces > 1) { line.Append(' ', trailingSpaces - 1); }
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
                ushort glyph = font?.MapCodePoint(rune.Value) ?? 0;
                double advance = font is null ? fallback!.MeasureAdvanceEm(rune) : font.GetAdvanceWidth(glyph);
                double next = units + advance + (prior == 0 || glyph == 0 ? 0d : font!.GetKerning(prior, glyph));
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
