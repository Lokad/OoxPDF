namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private sealed record DocxMidLineMath(DocxPreparedMathParagraph Math, int LineIndex, int LineCharOffset, double Width);

    private static DocxMidLineMath? WrapInlineMath(DocxParagraph paragraph, DocxPreparedMathParagraph math,
        IReadOnlyList<DocxTextSpan> spans, ref DocxWrappedTextLine[] lines, double firstWidth, double continuationWidth,
        double fontSize, IDocxTextMeasurer measurer, double scale, double defaultTabStop, int pageNumber,
        CancellationToken cancellationToken)
    {
        DocxMidLineMath? fitting = CreateMidLineMath(paragraph, math, spans, lines, firstWidth, continuationWidth,
            fontSize, measurer, scale, defaultTabStop, pageNumber);
        if (fitting is not null) { return fitting; }
        double width = math.Box.Width * scale;
        // A box that fits no line retains the whole-equation fallback. Never
        // shrink mathematical content to force it into the body frame.
        if (width > Math.Max(firstWidth, continuationWidth)) { return null; }

        // Word can break a literal formula at mathematical operators. Until
        // that contract is supported, only wrap following prose: the complete
        // formula must still fit at its original source-owned line position.
        var source = FindMathSourceLine(math.Equation, spans, lines);
        if (source is null) { return null; }
        int sourceLine = source.Value.LineIndex;
        double before = MeasureTextSpansForLayout(SliceTextSpans(lines[sourceLine].Spans, 0, source.Value.CharOffset),
            fontSize, measurer, ScaleTabStopPositions(paragraph.TabStops, scale), defaultTabStop * scale, pageNumber);
        if (before + width > (sourceLine == 0 ? firstWidth : continuationWidth)) { return null; }
        DocxTextSpan? following = spans.FirstOrDefault(span => span.SourceTextRunIndex >= math.Equation.SourceRunIndex && span.Text.Length != 0);
        if (following is null || !char.IsWhiteSpace(following.Text[0])) { return null; }

        var marked = new List<DocxTextSpan>(spans.Count + 1);
        int position = 0;
        bool inserted = false;
        foreach (DocxTextSpan span in spans)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!inserted && span.SourceTextRunIndex >= math.Equation.SourceRunIndex)
            {
                marked.Add(Marker(span.StyleRun));
                inserted = true;
            }
            marked.Add(span);
            if (!inserted) { position += span.Text.Length; }
        }
        if (!inserted) { marked.Add(Marker(spans[^1].StyleRun)); }
        DocxWrappedTextLine[] wrapped = WrapTextLines(marked, firstWidth, continuationWidth, fontSize, measurer,
            ScaleTabStopPositions(paragraph.TabStops, scale), defaultTabStop * scale,
            ShouldAllowCharacterLevelWordWrap(paragraph), pageNumber, cancellationToken,
            inlineImageWidths: [(position, width)]).ToArray();
        for (int index = 0; index < wrapped.Length; index++)
        {
            if (index == sourceLine && wrapped[index].InlineObjectCharOffset is { } offset)
            {
                lines = wrapped;
                return new DocxMidLineMath(math, index, offset, width);
            }
        }
        return null;

        static DocxTextSpan Marker(DocxTextRun style) => new("\u200B", style, -1, 0) { IsInlineObjectMarker = true };
    }

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
        var source = FindMathSourceLine(math.Equation, spans, lines);
        if (source is null) { return null; }
        int index = source.Value.LineIndex;
        double maximum = index == 0 ? firstWidth : continuationWidth;
        double width = math.Box.Width * scale;
        double textWidth = MeasureTextSpansForLayout(lines[index].Spans, fontSize, measurer,
            ScaleTabStopPositions(paragraph.TabStops, scale), defaultTabStop * scale, pageNumber);
        return textWidth + width <= maximum
            ? new DocxMidLineMath(math, index, source.Value.CharOffset, width) : null;
    }

    private static (int LineIndex, int CharOffset)? FindMathSourceLine(DocxEquation equation,
        IReadOnlyList<DocxTextSpan> spans, DocxWrappedTextLine[] lines)
    {
        int position = spans.Where(span => span.SourceTextRunIndex < equation.SourceRunIndex).Sum(span => span.Text.Length);
        int start = 0;
        for (int index = 0; index < lines.Length; index++)
        {
            DocxWrappedTextLine line = lines[index];
            int length = line.Spans.Sum(span => span.Text.Length);
            bool last = index == lines.Length - 1;
            if (position < start + length || last && position == start + length)
            {
                return (index, position - start);
            }
            start += length;
        }
        return null;
    }
}
