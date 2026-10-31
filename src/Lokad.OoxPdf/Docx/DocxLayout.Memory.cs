namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Source-family capacity for repeated page frames, not per-object receipts.
    // Ordinary body placement already has model headroom; headers/footers and
    // their tables/textboxes can be copied on every page, independent of body size.
    private static long EstimatePagePlacementCapacity(DocxPageSettings settings, CancellationToken token)
    {
        long capacity = 4096;
        IEnumerable<DocxFloatingDrawing> drawings = settings.HeaderFloatingDrawingsByType.Values
            .Concat(settings.FooterFloatingDrawingsByType.Values).SelectMany(value => value);
        foreach (var drawing in drawings)
        {
            token.ThrowIfCancellationRequested();
            capacity = checked(capacity + 1024);
            foreach (var paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(drawing.TextBoxBodyElements)) AddParagraph(paragraph);
        }
        foreach (var paragraph in DocxBlockTraversal.EnumerateStaticStoryParagraphs(settings)) AddParagraph(paragraph);
        foreach (var elements in settings.HeaderBodyElementsByType.Values.Concat(settings.FooterBodyElementsByType.Values))
            foreach (var table in DocxBlockTraversal.EnumerateBodyTables(elements))
                foreach (var row in table.Rows)
                {
                    token.ThrowIfCancellationRequested();
                    capacity = checked(capacity + 512L * (1 + row.Cells.Count));
                }
        return capacity;

        void AddParagraph(DocxParagraph paragraph)
        {
            AddContent(paragraph);
            foreach (var child in DocxBlockTraversal.EnumerateInlineTextBoxParagraphs(paragraph)) AddContent(child);
        }
        void AddContent(DocxParagraph paragraph)
        {
            token.ThrowIfCancellationRequested();
            capacity = checked(capacity + 512L + 512L * (paragraph.Images.Count +
                paragraph.Equations.Count + paragraph.InlineTextBoxes.Count));
            foreach (var run in paragraph.Runs)
            {
                token.ThrowIfCancellationRequested();
                capacity = checked(capacity + 128L + 64L * run.Text.Length);
            }
        }
    }
}
