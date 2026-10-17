namespace Lokad.OoxPdf.Docx;

internal static class DocxTableCellClipGeometry
{
    // Numbering can use a cell only through its complete measured, painted line
    // frame. Missing measurements, merged ownership and clipped lines keep the
    // existing fallback; typography callers do not acquire this broader scope.
    internal static IEnumerable<(DocxTextLineLayout Line, DocxTableCellClipSnapshot? Clip)> EnumerateBodyLines(
        DocxLayoutPage page, CancellationToken token)
    {
        foreach (var item in page.Items)
        {
            token.ThrowIfCancellationRequested();
            if (item is DocxTextLineLayout line) yield return (line, null);
            else if (item is DocxTableRowLayout row)
                foreach (var owner in EnumerateCellLines(row, null, token)) yield return owner;
        }
    }

    internal static IEnumerable<(DocxTextLineLayout Line, DocxTableCellClipSnapshot? Clip)> EnumerateCellLines(
        DocxTableRowLayout row, DocxTableCellClipSnapshot? parent, CancellationToken token)
    {
        const double roundoff = .000001d;
        foreach (var cell in row.Cells)
        {
            token.ThrowIfCancellationRequested();
            if (cell.VisualOwnership != DocxTableCellVisualOwnership.OwnCell) continue;
            var clip = Resolve(row, cell);
            if (parent is not null)
            {
                double x = Math.Max(clip.X, parent.X), y = Math.Max(clip.Y, parent.Y);
                clip = new(true, x, y, Math.Min(clip.X + clip.Width, parent.X + parent.Width) - x,
                    Math.Min(clip.Y + clip.Height, parent.Y + parent.Height) - y);
            }
            if (!double.IsFinite(clip.X) || !double.IsFinite(clip.Y) || !double.IsFinite(clip.Width) ||
                !double.IsFinite(clip.Height) || clip.Width <= 0d || clip.Height <= 0d) continue;
            foreach (var line in cell.TextLines)
            {
                token.ThrowIfCancellationRequested();
                if (DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double bottom, out double top) &&
                    bottom >= clip.Y - roundoff && top <= clip.Y + clip.Height + roundoff)
                    yield return (line, clip);
            }
            foreach (var nested in cell.NestedRows)
                foreach (var owner in EnumerateCellLines(nested, clip, token)) yield return owner;
        }
    }

    internal static bool ContainsSegment(DocxTableCellClipSnapshot? clip, DocxTextSegmentLayout segment)
    {
        const double roundoff = .000001d;
        return clip is null || double.IsFinite(segment.X) && double.IsFinite(segment.Width) && segment.Width > 0d &&
            segment.X >= clip.X - roundoff && segment.X + segment.Width <= clip.X + clip.Width + roundoff;
    }

    internal static DocxTableCellClipSnapshot Resolve(
        DocxTableRowLayout row, DocxTableCellLayout cell)
    {
        var fallback = new DocxTableCellClipSnapshot(false, cell.X, cell.Y, cell.Width, cell.Height);
        bool declaredMarginClip = row.RowIndex >= 0 && row.RowIndex < row.Table.AlignedMinimumRowBottomMargins.Count &&
            row.Table.AlignedMinimumRowBottomMargins[row.RowIndex] is > 0d;
        bool automaticClip = (row.Table.AutomaticMinimumRowPlacement || row.Table.AutomaticTopRowContentClip) && row.Cells.Count > 0 &&
            DocxLayoutEngine.HasAutomaticTableRowHeight(row.DeclaredHeightPoints, row.HeightRuleValue);
        bool declaredTopClip = row.RowIndex >= 0 && row.RowIndex < row.Table.DeclaredTopRowContentClips.Count &&
            row.Table.DeclaredTopRowContentClips[row.RowIndex];
        if (row.FragmentCount != 1 || row.ReviewBorderPrintScale != 1d || row.RowIndex < 0 ||
            !(declaredMarginClip || automaticClip || declaredTopClip) ||
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
        double bottom = (automaticClip ? row.Cells.Max(c => c.ContentPaddingBottom) : row.Table.AlignedMinimumRowBottomMargins[row.RowIndex]) +
            (row.RowIndex == row.Table.RowCount - 1 ? Width("bottom") : 0d);
        double width = cell.Width - left - right;
        double height = cell.Height - top - bottom;
        if (!(width > 0d) || !(height > 0d)) return fallback;
        return new DocxTableCellClipSnapshot(true, cell.X + left, cell.Y + bottom, width, height);
    }
}
