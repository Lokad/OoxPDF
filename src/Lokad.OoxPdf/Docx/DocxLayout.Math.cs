namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private sealed record DocxMidLineMath(DocxPreparedMathParagraph Math, int LineIndex, int LineCharOffset, double Width);

    private static (double Above, double Below) ResolveInlineMathExpansion(
        DocxPreparedMathParagraph math, double fontSize, double baselineLineHeight, double lineHeight, double scale, bool exact)
    {
        if (exact) { return (0d, 0d); }
        double inset = DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, baselineLineHeight, false);
        return (Math.Max(0d, math.BaselineInset * scale - inset),
            Math.Max(0d, (math.LineHeight - math.BaselineInset) * scale - (lineHeight - inset)));
    }

    private static DocxMidLineMath? CreateMidLineMath(DocxParagraph paragraph, DocxPreparedMathParagraph math,
        IReadOnlyList<DocxTextSpan> spans, DocxWrappedTextLine[] lines, double firstWidth, double continuationWidth,
        double fontSize, IDocxTextMeasurer measurer, double scale, double defaultTabStop, int pageNumber)
    {
        int position = spans.Where(span => span.SourceTextRunIndex < math.Equation.SourceRunIndex).Sum(span => span.Text.Length);
        int start = 0;
        for (int index = 0; index < lines.Length; index++)
        {
            DocxWrappedTextLine line = lines[index];
            int length = line.Spans.Sum(span => span.Text.Length);
            bool last = index == lines.Length - 1;
            if (position < start + length || last && position == start + length)
            {
                double maximum = index == 0 ? firstWidth : continuationWidth;
                double width = math.Box.Width * scale;
                double textWidth = MeasureTextSpansForLayout(line.Spans, fontSize, measurer,
                    ScaleTabStopPositions(paragraph.TabStops, scale), defaultTabStop * scale, pageNumber);
                // A fitting row shares the existing inline-object placement contract.
                // Atomic-object driven wrapping remains an explicit fallback here.
                return textWidth + width <= maximum
                    ? new DocxMidLineMath(math, index, position - start, width) : null;
            }
            start += length;
        }
        return null;
    }
}
