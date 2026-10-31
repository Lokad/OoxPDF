namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Source-family capacity for repeated frames, not per-object receipts. The
    // largest flowing table header can recur in every column of every page.
    private static long EstimateRepeatedTableHeaderCapacity(DocxDocument document, CancellationToken token)
    {
        long largest = 0;
        foreach (var table in DocxBlockTraversal.EnumerateBodyTables(document))
        {
            token.ThrowIfCancellationRequested();
            long capacity = 0;
            foreach (var row in table.Rows.TakeWhile(row => row.IsHeader))
            {
                token.ThrowIfCancellationRequested();
                capacity = checked(capacity + 512L * (1 + row.Cells.Count));
                foreach (var cell in row.Cells)
                    capacity = checked(capacity + EstimateBlockPlacementCapacity(DocxTableCellContent.GetBodyElements(cell), token));
            }
            largest = Math.Max(largest, capacity);
        }
        return largest;
    }

    private static long EstimatePagePlacementCapacity(DocxEffectiveSectionSettings section, long repeatedTableHeaderCapacity, CancellationToken token)
    {
        DocxPageSettings settings = section.PageSettings;
        // An upper estimate of frame multiplicity, without repeating geometry's
        // choice between declared equal columns and custom definitions.
        int columns = Math.Max(1, Math.Max(section.SectionProperties.ColumnCount ?? 1, section.SectionProperties.ColumnDefinitions.Count));
        long capacity = checked(4096 + columns * repeatedTableHeaderCapacity);
        foreach (var drawing in settings.HeaderFloatingDrawingsByType.Values.Concat(settings.FooterFloatingDrawingsByType.Values).SelectMany(value => value))
        {
            token.ThrowIfCancellationRequested();
            capacity = checked(capacity + 1024 + EstimateBlockPlacementCapacity(drawing.TextBoxBodyElements, token));
        }
        foreach (var paragraph in DocxBlockTraversal.EnumerateStaticStoryParagraphs(settings))
            capacity = checked(capacity + EstimateParagraphPlacementCapacity(paragraph, token));
        foreach (var elements in settings.HeaderBodyElementsByType.Values.Concat(settings.FooterBodyElementsByType.Values))
            capacity = checked(capacity + EstimateTableStructureCapacity(elements, token));
        return capacity;
    }

    private static long EstimateBlockPlacementCapacity(IReadOnlyList<DocxBodyElement> elements, CancellationToken token)
    {
        long capacity = EstimateTableStructureCapacity(elements, token);
        foreach (var paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(elements))
            capacity = checked(capacity + EstimateParagraphPlacementCapacity(paragraph, token));
        return capacity;
    }

    private static long EstimateTableStructureCapacity(IReadOnlyList<DocxBodyElement> elements, CancellationToken token)
    {
        long capacity = 0;
        foreach (var table in DocxBlockTraversal.EnumerateBodyTables(elements))
            foreach (var row in table.Rows)
            {
                token.ThrowIfCancellationRequested();
                capacity = checked(capacity + 512L * (1 + row.Cells.Count));
            }
        return capacity;
    }

    private static long EstimateParagraphPlacementCapacity(DocxParagraph paragraph, CancellationToken token)
    {
        long capacity = Content(paragraph);
        foreach (var child in DocxBlockTraversal.EnumerateInlineTextBoxParagraphs(paragraph))
            capacity = checked(capacity + Content(child));
        return capacity;

        long Content(DocxParagraph source)
        {
            token.ThrowIfCancellationRequested();
            long bytes = checked(512L + 512L * (source.Images.Count + source.Equations.Count + source.InlineTextBoxes.Count));
            foreach (var box in source.InlineTextBoxes)
                bytes = checked(bytes + EstimateTableStructureCapacity(box.BodyElements, token));
            foreach (var run in source.Runs)
            {
                token.ThrowIfCancellationRequested();
                bytes = checked(bytes + 128L + 64L * run.Text.Length);
            }
            return bytes;
        }
    }
}
