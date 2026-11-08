using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static class DocxTableCellClipGeometry
{
    internal static bool HasCompleteMeasuredRows(IReadOnlyList<DocxTableRowLayout> rows, CancellationToken token,
        Func<DocxTextLineLayout, DocxTableCellClipSnapshot?, bool>? inspect = null)
    {
        long expected = 0, visible = 0;
        bool Count(IReadOnlyList<DocxTableRowLayout> current)
        {
            foreach (var row in current)
            {
                token.ThrowIfCancellationRequested();
                if (row.FragmentCount != 1) return false;
                foreach (var cell in row.Cells)
                {
                    token.ThrowIfCancellationRequested();
                    expected += cell.TextLines.Count(line => !DocxInlineLineGeometry.IsNonPaintingSourceCarrier(line));
                    if (!Count(cell.NestedRows)) return false;
                }
            }
            return true;
        }
        if (!Count(rows)) return false;
        foreach (var row in rows)
        foreach (var owner in EnumerateCellLines(row, null, token))
        {
            token.ThrowIfCancellationRequested();
            visible++;
            if (inspect is not null && !inspect(owner.Line, owner.Clip)) return false;
        }
        return visible == expected;
    }

    // A complete inline frame may paint its stroke around the cell's content
    // boundary. Children still use the unchanged content clip. A partial drawing
    // viewport or a frame outside the cell retains the conservative clip.
    internal static DocxTableCellClipSnapshot ResolveInlineOutline(
        DocxTableCellClipSnapshot contentClip, PdfRectangle frame, double strokeWidth, double? viewportBottom)
    {
        const double roundoff = .000001d;
        double right = frame.X + frame.Width, top = frame.Y + frame.Height;
        if (!(strokeWidth > 0d) || !double.IsFinite(strokeWidth) ||
            !(frame.Width > 0d) || !(frame.Height > 0d) ||
            !double.IsFinite(frame.X) || !double.IsFinite(frame.Y) || !double.IsFinite(right) || !double.IsFinite(top) ||
            !double.IsFinite(contentClip.X) || !double.IsFinite(contentClip.Y) ||
            !double.IsFinite(contentClip.Width) || !double.IsFinite(contentClip.Height) ||
            !(frame.X >= contentClip.X - roundoff && frame.Y >= contentClip.Y - roundoff &&
              right <= contentClip.X + contentClip.Width + roundoff && top <= contentClip.Y + contentClip.Height + roundoff) ||
            viewportBottom is { } bottom && !(bottom <= frame.Y + roundoff))
            return contentClip;

        return contentClip with { X = contentClip.X - strokeWidth / 2d, Y = contentClip.Y - strokeWidth / 2d,
            Width = contentClip.Width + strokeWidth, Height = contentClip.Height + strokeWidth };
    }

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
        foreach (var cell in row.Cells)
        {
            token.ThrowIfCancellationRequested();
            if (cell.VisualOwnership != DocxTableCellVisualOwnership.OwnCell) continue;
            var clip = Resolve(row, cell, parent);
            if (!double.IsFinite(clip.X) || !double.IsFinite(clip.Y) || !double.IsFinite(clip.Width) ||
                !double.IsFinite(clip.Height) || clip.Width <= 0d || clip.Height <= 0d) continue;
            foreach (var line in cell.TextLines)
            {
                token.ThrowIfCancellationRequested();
                if (DocxInlineLineGeometry.IsNonPaintingSourceCarrier(line)) continue;
                if (DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double bottom, out double top) &&
                    ContainsVerticalRange(clip, bottom, top))
                    yield return (line, clip);
            }
            foreach (var nested in cell.NestedRows)
                foreach (var owner in EnumerateCellLines(nested, clip, token)) yield return owner;
        }
    }

    internal static bool ContainsSegment(DocxTableCellClipSnapshot? clip, DocxTextSegmentLayout segment)
        => clip is null || segment.Width > 0d && ContainsHorizontalRange(clip, segment.X, segment.Width);

    internal static bool ContainsHorizontalRange(DocxTableCellClipSnapshot? clip, double x, double width)
    {
        const double roundoff = .000001d;
        return clip is null || double.IsFinite(x) && double.IsFinite(width) && width >= 0d &&
            x >= clip.X - roundoff && x + width <= clip.X + clip.Width + roundoff;
    }

    internal static bool ContainsVerticalRange(DocxTableCellClipSnapshot? clip, double bottom, double top)
    {
        const double roundoff = .000001d;
        return clip is null || double.IsFinite(bottom) && double.IsFinite(top) && top >= bottom &&
            bottom >= clip.Y - roundoff && top <= clip.Y + clip.Height + roundoff;
    }

    internal static DocxTableCellClipSnapshot Resolve(DocxTableRowLayout row, DocxTableCellLayout cell,
        DocxTableCellClipSnapshot? parent)
    {
        var clip = Resolve(row, cell);
        if (parent is null) return clip;
        double x = Math.Max(clip.X, parent.X), y = Math.Max(clip.Y, parent.Y);
        return new(true, x, y, Math.Min(clip.X + clip.Width, parent.X + parent.Width) - x,
            Math.Min(clip.Y + clip.Height, parent.Y + parent.Height) - y);
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
