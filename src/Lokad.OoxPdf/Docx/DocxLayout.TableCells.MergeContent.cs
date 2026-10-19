namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static bool HasAdmittedMergeSourceFrame(DocxTable table, DocxTableLayoutContext context,
        DocxVerticalMergeOwner source, IReadOnlyList<double> rowHeights, IReadOnlyList<DocxTextLineLayout> lines,
        IEnumerable<DocxInlineImageLayout> images, double top)
    {
        double capacity = context.PageContentHeight;
        for (int index = 0; index < table.Rows.Count && table.Rows[index].IsHeader; index++)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (index == source.RowIndex) return false;
            capacity -= rowHeights[index];
        }
        double height = context.AdmittedMergeSpanHeights.TryGetValue(source.Cell, out double admittedHeight)
            ? admittedHeight : rowHeights[source.RowIndex];
        return double.IsFinite(capacity) && double.IsFinite(height) && height > 0d && height <= capacity &&
            HasContainedInlineFrames(lines, images, top - height, top, context.CancellationToken);
    }

    private static bool IsStaticTopAlignedMergeSource(DocxTableCell cell) =>
        !HasPageDynamicFields(GetTableCellLayoutBodyElements(cell)) &&
        GetTableCellLayoutBodyElements(cell).All(element => element is DocxParagraphElement) &&
        (string.IsNullOrEmpty(cell.VerticalAlignmentValue) ||
            string.Equals(cell.VerticalAlignmentValue, "top", StringComparison.OrdinalIgnoreCase));

    private sealed record MeasuredMergeSpan(DocxTableCell Cell, int FirstRow, int LastRow)
    {
        public double RequiredHeight { get; set; }
    }

    private static Dictionary<DocxTableCell, MeasuredMergeSpan> FindCompleteMergeSpans(
        DocxTable table, IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? math, CancellationToken token)
    {
        var result = new Dictionary<DocxTableCell, MeasuredMergeSpan>(ReferenceEqualityComparer.Instance);
        if (!table.Rows.Any(row => row.Cells.Any(cell => cell.IsVerticalMergeRestart))) return result;
        // Fixed rows and sources whose page capacity can change during placement
        // retain the existing staged path, rather than lending unadmitted space.
        foreach (var row in table.Rows)
        {
            token.ThrowIfCancellationRequested();
            if (row.IsHeader || string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase)) return result;
            foreach (var cell in row.Cells)
                if (GetTableCellLayoutBodyElements(cell).Any(element => element is not DocxParagraphElement) ||
                    HasPageDynamicFields(GetTableCellLayoutBodyElements(cell)) ||
                    cell.Paragraphs.Any(paragraph => paragraph.InlineReferences.Count != 0)) return result;
        }
        for (int first = 0; first < table.Rows.Count; first++)
        {
            token.ThrowIfCancellationRequested();
            int gridColumn = 0;
            foreach (var cell in table.Rows[first].Cells)
            {
                if (cell.IsVerticalMergeRestart && IsStaticTopAlignedMergeSource(cell) &&
                    DocxMathParagraphPolicy.HasSupportedCellContentProperties(cell) &&
                    cell.Paragraphs.All(paragraph => paragraph.InlineTextBoxes.Count == 0) &&
                    cell.Paragraphs.SelectMany(p => p.Equations).All(equation => math?.ContainsKey(equation) == true))
                {
                    int last = first;
                    for (int next = first + 1; next < table.Rows.Count; next++)
                    {
                        token.ThrowIfCancellationRequested();
                        if (!TryGetCellAtGridColumn(table.Rows[next], gridColumn, out var continuation) ||
                            continuation is null || !IsVerticalMergeContinuation(continuation) ||
                            continuation.GridSpan != cell.GridSpan) break;
                        if (continuation.Paragraphs.Any(p => p.Runs.Any(r => r.Text.Length != 0) ||
                            p.Equations.Count != 0 || p.Images.Count != 0 || p.InlineTextBoxes.Count != 0))
                        { last = first; break; }
                        last = next;
                    }
                    if (last > first) result.Add(cell, new(cell, first, last));
                }
                gridColumn += Math.Max(1, cell.GridSpan);
            }
        }
        return result;
    }

    private static bool TryAdmitMergeSpans(IReadOnlyList<double> baseHeights,
        Dictionary<DocxTableCell, MeasuredMergeSpan> spans, double available,
        CancellationToken token, out double[] rowHeights, out IReadOnlyDictionary<DocxTableCell, double> admitted)
    {
        rowHeights = [];
        admitted = new Dictionary<DocxTableCell, double>(ReferenceEqualityComparer.Instance);
        if (spans.Count == 0 || !double.IsFinite(available) || available <= 0d) return false;
        var ending = new List<MeasuredMergeSpan>?[baseHeights.Count];
        foreach (var span in spans.Values)
        {
            token.ThrowIfCancellationRequested();
            if (!double.IsFinite(span.RequiredHeight) || span.RequiredHeight <= 0d) return false;
            (ending[span.LastRow] ??= []).Add(span);
        }
        var prefix = new double[baseHeights.Count + 1];
        var candidate = new double[baseHeights.Count];
        for (int row = 0; row < candidate.Length; row++)
        {
            token.ThrowIfCancellationRequested();
            if (!double.IsFinite(baseHeights[row]) || baseHeights[row] <= 0d) return false;
            prefix[row + 1] = prefix[row] + baseHeights[row];
            if (ending[row] is { } owners)
                foreach (var owner in owners)
                    prefix[row + 1] = Math.Max(prefix[row + 1], prefix[owner.FirstRow] + owner.RequiredHeight);
            if (!double.IsFinite(prefix[row + 1]) || prefix[row + 1] > available) return false;
            candidate[row] = prefix[row + 1] - prefix[row];
        }
        var heights = new Dictionary<DocxTableCell, double>(ReferenceEqualityComparer.Instance);
        foreach (var span in spans.Values)
        {
            token.ThrowIfCancellationRequested();
            heights.Add(span.Cell, prefix[span.LastRow + 1] - prefix[span.FirstRow]);
        }
        rowHeights = candidate;
        admitted = heights;
        return true;
    }
}
