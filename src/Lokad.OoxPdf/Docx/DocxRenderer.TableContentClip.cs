namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    internal static DocxTableCellClipSnapshot ResolveTableCellContentClip(
        DocxTableRowLayout row, DocxTableCellLayout cell) => DocxTableCellClipGeometry.Resolve(row, cell);
}
