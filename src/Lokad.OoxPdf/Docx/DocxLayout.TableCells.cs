using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Automatic rows measure their occupied height with printed line metrics.
    // Fragment admission and painting must use that same first-baseline inset,
    // independently of paragraph keep/widow policy or page-capacity eligibility.
    private static double ResolveTableRowFragmentBaselineScale(DocxTableRow row, double scale, int fragmentCount) =>
        fragmentCount > 1 && HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue) &&
        row.Cells.All(HasNaturalReviewCellLineMetrics) ? scale : 1d;

    // These font-inset metrics also own prose when unsupported mathematics
    // falls back. Capacity eligibility separately excludes object-height owners.
    private static bool HasNaturalReviewCellLineMetrics(DocxTableCell cell)
    {
        IReadOnlyList<DocxBodyElement> elements = DocxTableCellContent.GetBodyElements(cell);
        return elements.Count == 1 && elements[0] is DocxParagraphElement { Paragraph: var paragraph } &&
            !cell.HasVerticalMerge && paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0 &&
            paragraph.InlineReferences.Count == 0 && paragraph.EffectiveProperties.LineSpacingPoints is null &&
            (string.IsNullOrEmpty(cell.VerticalAlignmentValue) ||
                string.Equals(cell.VerticalAlignmentValue, "top", StringComparison.OrdinalIgnoreCase));
    }

    private static int CountTableRowSourceRevisions(DocxTableRow row)
    {
        return CollectTableRowSourceRevisions(row).Count;
    }

    private static IReadOnlyList<DocxRevisionInfo> CollectTableRowSourceRevisions(DocxTableRow row)
    {
        var revisions = new List<DocxRevisionInfo>(row.Revisions);
        foreach (DocxTableCell cell in row.Cells)
        {
            revisions.AddRange(cell.Revisions);
            foreach (DocxTable nestedTable in cell.BodyElements.OfType<DocxTableElement>().Select(element => element.Table))
            {
                foreach (DocxTableRow nestedRow in nestedTable.Rows)
                {
                    revisions.AddRange(CollectTableRowSourceRevisions(nestedRow));
                }
            }
        }

        return revisions;
    }

    private static bool IsTextLineVisibleInCellFragmentGeometry(
        DocxTextLineLayout line,
        double cellY,
        double cellHeight,
        int fragmentIndex,
        int fragmentCount)
    {
        if (fragmentCount <= 1)
        {
            return true;
        }

        double bottom = fragmentIndex == 0 ? cellY - 0.001d : cellY + 0.001d;
        return line.BaselineY >= bottom && line.BaselineY <= cellY + cellHeight + 0.001d;
    }

    private static double VerticalOverlap(double firstY, double firstHeight, double secondY, double secondHeight)
    {
        return Math.Min(firstY + firstHeight, secondY + secondHeight) - Math.Max(firstY, secondY);
    }

    private static bool IsVerticalMergeContinuation(DocxTableCell cell)
    {
        return cell.HasVerticalMerge && !cell.IsVerticalMergeRestart;
    }

    private static DocxVerticalMergeOwner? FindVerticalMergeRestartOwner(
        DocxTable table,
        int rowIndex,
        int gridColumnIndex)
    {
        for (int previousRowIndex = rowIndex - 1; previousRowIndex >= 0; previousRowIndex--)
        {
            if (!TryGetCellAtGridColumn(table.Rows[previousRowIndex], gridColumnIndex, out DocxTableCell? previousCell) ||
                previousCell is null)
            {
                return null;
            }

            if (previousCell.IsVerticalMergeRestart)
            {
                return new DocxVerticalMergeOwner(previousCell, previousRowIndex, gridColumnIndex);
            }

            if (!IsVerticalMergeContinuation(previousCell))
            {
                return null;
            }
        }

        return null;
    }

    private static double GetVerticalMergeSpanHeight(
        DocxTable table,
        int rowIndex,
        int gridColumnIndex,
        IReadOnlyList<double> rowHeights)
    {
        double height = rowHeights[rowIndex];
        for (int nextRowIndex = rowIndex + 1; nextRowIndex < table.Rows.Count; nextRowIndex++)
        {
            if (!TryGetCellAtGridColumn(table.Rows[nextRowIndex], gridColumnIndex, out DocxTableCell? nextCell) ||
                nextCell is null ||
                !IsVerticalMergeContinuation(nextCell))
            {
                break;
            }

            height += rowHeights[nextRowIndex];
        }

        return height;
    }

    private static bool TryGetCellAtGridColumn(DocxTableRow row, int gridColumnIndex, out DocxTableCell? cell)
    {
        int currentGridColumnIndex = row.GridBefore;
        foreach (DocxTableCell candidate in row.Cells)
        {
            int span = Math.Max(1, candidate.GridSpan);
            if (gridColumnIndex >= currentGridColumnIndex && gridColumnIndex < currentGridColumnIndex + span)
            {
                cell = candidate;
                return true;
            }

            currentGridColumnIndex += span;
        }

        cell = null;
        return false;
    }

    private static double[] GetTableRowCellWidths(DocxTableRow row, IReadOnlyList<double> effectiveColumns, double scale)
    {
        var widths = new double[row.Cells.Count];
        int gridColumnIndex = row.GridBefore;
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCell cell = row.Cells[cellIndex];
            int span = Math.Max(1, cell.GridSpan);
            double width = 0d;
            for (int spanIndex = 0; spanIndex < span; spanIndex++)
            {
                width += effectiveColumns[Math.Min(gridColumnIndex + spanIndex, effectiveColumns.Count - 1)] * scale;
            }

            widths[cellIndex] = width;
            gridColumnIndex += span;
        }

        return widths;
    }

}
