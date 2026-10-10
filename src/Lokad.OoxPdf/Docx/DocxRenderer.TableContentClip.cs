namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    internal static DocxTableCellClipSnapshot ResolveTableCellContentClip(
        DocxTableRowLayout row, DocxTableCellLayout cell)
    {
        var fallback = new DocxTableCellClipSnapshot(false, cell.X, cell.Y, cell.Width, cell.Height);
        bool declaredMarginClip = row.RowIndex >= 0 && row.RowIndex < row.Table.AlignedMinimumRowBottomMargins.Count &&
            row.Table.AlignedMinimumRowBottomMargins[row.RowIndex] is > 0d;
        bool automaticClip = row.Table.AutomaticMinimumRowPlacement && row.RowIndex == 0 && row.Cells.Count == 1;
        if (row.FragmentCount != 1 || row.ReviewBorderPrintScale != 1d || row.RowIndex < 0 ||
            !(declaredMarginClip || automaticClip) ||
            row.Cells.Any(c => c.VisualCell.Borders.Any(b =>
                DocxTableBorderGeometry.ResolveVisibleWidth(b) > 0d &&
                !string.Equals(b.Value, "single", StringComparison.OrdinalIgnoreCase) &&
                !(b.UseDoubleBandWidth && string.Equals(b.Value, "double", StringComparison.OrdinalIgnoreCase)))))
            return fallback;

        double Width(string edge) => DocxTableBorderGeometry.ResolveVisibleWidth(
            DocxTableBorderGeometry.Find(cell.VisualCell.Borders, edge));
        double left = Width("left") / 2d;
        double right = Width("right") / 2d;
        double top = cell.ContentPaddingTop + Width("top");
        double bottom = (automaticClip ? cell.ContentPaddingBottom : row.Table.AlignedMinimumRowBottomMargins[row.RowIndex]) +
            (row.RowIndex == row.Table.RowCount - 1 ? Width("bottom") : 0d);
        double width = cell.Width - left - right;
        double height = cell.Height - top - bottom;
        if (!(width > 0d) || !(height > 0d)) return fallback;
        return new DocxTableCellClipSnapshot(true, cell.X + left, cell.Y + bottom, width, height);
    }
}
