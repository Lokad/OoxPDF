namespace Lokad.OoxPdf.Docx;

// Publish one replacement paragraph through both ordered body elements and
// fallback views. Unchanged containers retain their existing identities.
internal sealed class DocxParagraphReplacements(
    IReadOnlyDictionary<DocxParagraph, DocxParagraph> paragraphs, CancellationToken token)
{
    private readonly Dictionary<DocxTable, DocxTable> tables = new(ReferenceEqualityComparer.Instance);

    internal DocxParagraph Replace(DocxParagraph paragraph) => paragraphs.GetValueOrDefault(paragraph, paragraph);

    internal IReadOnlyList<DocxBodyElement> ReplaceElements(IReadOnlyList<DocxBodyElement> elements)
    {
        List<DocxBodyElement>? changed = null;
        for (int index = 0; index < elements.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            var original = elements[index];
            var updated = original switch {
                DocxParagraphElement element when !ReferenceEquals(element.Paragraph, Replace(element.Paragraph)) =>
                    element with { Paragraph = Replace(element.Paragraph) },
                DocxTableElement element when !ReferenceEquals(element.Table, ReplaceTable(element.Table)) =>
                    element with { Table = ReplaceTable(element.Table) },
                _ => original
            };
            if (!ReferenceEquals(original, updated) && changed is null) changed = [.. elements.Take(index)];
            changed?.Add(updated);
        }
        return changed is null ? elements : changed;
    }

    internal DocxTable ReplaceTable(DocxTable table)
    {
        token.ThrowIfCancellationRequested();
        if (tables.TryGetValue(table, out var cached)) return cached;
        List<DocxTableRow>? rows = null;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            token.ThrowIfCancellationRequested();
            var row = table.Rows[rowIndex];
            List<DocxTableCell>? cells = null;
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                token.ThrowIfCancellationRequested();
                var cell = row.Cells[cellIndex];
                var source = DocxTableCellContent.GetBodyElements(cell);
                var elements = ReplaceElements(source);
                var updated = cell;
                if (!ReferenceEquals(source, elements))
                {
                    var direct = DocxBlockTraversal.EnumerateDirectParagraphs(elements).ToArray();
                    updated = cell with {
                        BodyElements = elements, Paragraphs = direct,
                        Text = string.Join(" ", direct.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))
                            .Where(text => text.Length != 0))
                    };
                    cells ??= [.. row.Cells.Take(cellIndex)];
                }
                cells?.Add(updated);
            }
            if (cells is not null) rows ??= [.. table.Rows.Take(rowIndex)];
            rows?.Add(cells is null ? row : row with { Cells = cells });
        }
        var result = rows is null ? table : table with { Rows = rows };
        tables.Add(table, result);
        return result;
    }
}
