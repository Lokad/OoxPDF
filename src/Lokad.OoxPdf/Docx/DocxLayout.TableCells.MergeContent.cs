namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static bool HasAdmittedMergeSourceFrame(DocxTable table, DocxTableLayoutContext context,
        int sourceRowIndex, IReadOnlyList<double> rowHeights, IReadOnlyList<DocxTextLineLayout> lines,
        IEnumerable<DocxInlineImageLayout> images, double top)
    {
        double capacity = context.PageContentHeight;
        for (int index = 0; index < table.Rows.Count && table.Rows[index].IsHeader; index++)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            if (index == sourceRowIndex) return false;
            capacity -= rowHeights[index];
        }
        double height = rowHeights[sourceRowIndex];
        return double.IsFinite(capacity) && double.IsFinite(height) && height > 0d && height <= capacity &&
            HasContainedInlineFrames(lines, images, top - height, top, context.CancellationToken);
    }
}
