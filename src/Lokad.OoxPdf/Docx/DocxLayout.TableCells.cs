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

    private static bool IsVerticalMergeRestart(DocxTableCell cell)
    {
        return cell.HasVerticalMerge &&
            string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase);
    }


    private static bool IsTextLineVisibleInCellFragmentGeometry(
        bool cellPageBreakAlignsWithFragmentBoundary,
        DocxTextLineLayout line,
        double cellY,
        double cellHeight,
        int fragmentIndex,
        int fragmentCount)
    {
        bool IsTextLineVisibleInCellFragment()
        {
            if (fragmentCount <= 1)
            {
                return true;
            }
    
            double bottom = fragmentIndex == 0 ? cellY - 0.001d : cellY + 0.001d;
            return line.BaselineY >= bottom && line.BaselineY <= cellY + cellHeight + 0.001d;
        }

        return cellPageBreakAlignsWithFragmentBoundary
            ? true
            : IsTextLineVisibleInCellFragment();
    }

    private static bool IsCellFragmentGeometryVisible(
        bool cellPageBreakAlignsWithFragmentBoundary,
        double y,
        double height,
        double cellY,
        double cellHeight)
    {
        return cellPageBreakAlignsWithFragmentBoundary
            ? true
            : VerticalOverlap(y, height, cellY, cellHeight) > 0.001d;
    }

    private static bool IsInlineImageVisibleInCellFragmentGeometry(
        bool cellPageBreakAlignsWithFragmentBoundary,
        DocxInlineImageLayout image,
        double cellY,
        double cellHeight,
        int fragmentIndex,
        int fragmentCount)
    {
        return IsCellFragmentGeometryVisible(
            cellPageBreakAlignsWithFragmentBoundary,
            image.Y,
            image.Height,
            cellY,
            cellHeight);
    }

    private static bool IsInlineTextBoxOnVisibleSideOfCellPageBreak(
        bool useCellPageBreakBoundaryPartition,
        int lowerParagraphBoundaryIndex,
        int? upperParagraphBoundaryIndex,
        DocxTableCell cell,
        DocxInlineTextBoxLayout box,
        int fragmentIndex,
        int fragmentCount)
    {
        return IsSourceParagraphIndexOnVisibleSideOfCellPageBreak(
            useCellPageBreakBoundaryPartition,
            lowerParagraphBoundaryIndex,
            upperParagraphBoundaryIndex,
            box.SourceParagraphIndex,
            fragmentCount);
    }

    private static bool IsInlineTextBoxVisibleInCellFragmentGeometry(
        bool cellPageBreakAlignsWithFragmentBoundary,
        DocxInlineTextBoxLayout box,
        double cellY,
        double cellHeight,
        int fragmentIndex,
        int fragmentCount)
    {
        return IsCellFragmentGeometryVisible(
            cellPageBreakAlignsWithFragmentBoundary,
            box.BoxTop - box.BoxHeight,
            box.BoxHeight,
            cellY,
            cellHeight);
    }

    private static bool IsNestedTableRowVisibleInCellFragmentGeometry(
        bool cellPageBreakAlignsWithFragmentBoundary,
        DocxTableRowLayout row,
        double cellY,
        double cellHeight,
        int fragmentIndex,
        int fragmentCount)
    {
        return IsCellFragmentGeometryVisible(
            cellPageBreakAlignsWithFragmentBoundary,
            row.Y,
            row.Height,
            cellY,
            cellHeight);
    }

    private static bool TryResolveTableCellParagraphBoundaryIndex(
        DocxTableCell cell,
        double cellWidth,
        double rowTopPadding,
        double fragmentBoundaryFromRowTop,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        out int paragraphBoundaryIndex,
        double paragraphSpacingScale)
    {
        paragraphBoundaryIndex = 0;
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return false;
        }

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double consumedHeight = rowTopPadding;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                consumedHeight += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                consumedHeight += MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, paragraphSpacingScale);
                if (consumedHeight >= fragmentBoundaryFromRowTop - 0.001d)
                {
                    return true;
                }

                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            consumedHeight += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            consumedHeight += MeasureTableCellParagraphContentHeight(cell, paragraph, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, paragraphSpacingScale);
            paragraphBoundaryIndex++;
            if (consumedHeight >= fragmentBoundaryFromRowTop - 0.001d)
            {
                return true;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        return paragraphBoundaryIndex > 0;
    }

    private sealed record DocxNestedTableBoundary(int BoundaryIndex, bool IsInsideNestedTable);

    private static bool TryResolveTableCellNestedTableBoundary(
        DocxTableCell cell,
        double cellWidth,
        double rowTopPadding,
        double fragmentBoundaryFromRowTop,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        out DocxNestedTableBoundary boundary,
        double paragraphSpacingScale)
    {
        boundary = new DocxNestedTableBoundary(0, IsInsideNestedTable: false);
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return false;
        }

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double consumedHeight = rowTopPadding;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int nestedTableIndex = 0;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                consumedHeight += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                double tableTop = consumedHeight;
                double tableHeight = MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, paragraphSpacingScale);
                consumedHeight += tableHeight;
                if (consumedHeight >= fragmentBoundaryFromRowTop - 0.001d)
                {
                    bool boundaryInsideTable =
                        fragmentBoundaryFromRowTop > tableTop + 0.001d &&
                        fragmentBoundaryFromRowTop < consumedHeight - 0.001d;
                    boundary = new DocxNestedTableBoundary(
                        boundaryInsideTable ? nestedTableIndex : nestedTableIndex + 1,
                        boundaryInsideTable);
                    return true;
                }

                nestedTableIndex++;
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            consumedHeight += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            consumedHeight += MeasureTableCellParagraphContentHeight(cell, paragraph, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, paragraphSpacingScale);
            if (consumedHeight >= fragmentBoundaryFromRowTop - 0.001d)
            {
                boundary = new DocxNestedTableBoundary(nestedTableIndex, IsInsideNestedTable: false);
                return true;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        boundary = new DocxNestedTableBoundary(nestedTableIndex, IsInsideNestedTable: false);
        return nestedTableIndex > 0;
    }

    private static bool IsSourceParagraphIndexOnVisibleSideOfCellPageBreak(
        bool useCellPageBreakBoundaryPartition,
        int lowerParagraphBoundaryIndex,
        int? upperParagraphBoundaryIndex,
        int? paragraphIndex,
        int fragmentCount)
    {
        if (!useCellPageBreakBoundaryPartition ||
            fragmentCount <= 1 ||
            paragraphIndex is not { } resolvedParagraphIndex)
        {
            return true;
        }

        return resolvedParagraphIndex >= lowerParagraphBoundaryIndex &&
            (upperParagraphBoundaryIndex is not { } upper || resolvedParagraphIndex < upper);
    }

    private static bool IsTextLineOnVisibleSideOfCellPageBreak(
        bool useCellPageBreakBoundaryPartition,
        int lowerParagraphBoundaryIndex,
        int? upperParagraphBoundaryIndex,
        DocxTableCell cell,
        DocxTextLineLayout line,
        int fragmentIndex,
        int fragmentCount)
    {
        return IsSourceParagraphIndexOnVisibleSideOfCellPageBreak(
            useCellPageBreakBoundaryPartition,
            lowerParagraphBoundaryIndex,
            upperParagraphBoundaryIndex,
            line.SourceParagraphIndex,
            fragmentCount);
    }

    private static bool IsInlineImageOnVisibleSideOfCellPageBreak(
        bool useCellPageBreakBoundaryPartition,
        int lowerParagraphBoundaryIndex,
        int? upperParagraphBoundaryIndex,
        DocxTableCell cell,
        DocxInlineImageLayout image,
        int fragmentIndex,
        int fragmentCount)
    {
        return IsSourceParagraphIndexOnVisibleSideOfCellPageBreak(
            useCellPageBreakBoundaryPartition,
            lowerParagraphBoundaryIndex,
            upperParagraphBoundaryIndex,
            image.SourceParagraphIndex,
            fragmentCount);
    }

    private static bool IsNestedTableRowOnVisibleSideOfCellPageBreak(
        bool useCellPageBreakBoundaryPartition,
        int lowerNestedTableBoundaryIndex,
        int? upperNestedTableBoundaryIndex,
        DocxTableRowLayout row,
        int fragmentCount)
    {
        if (!useCellPageBreakBoundaryPartition || fragmentCount <= 1)
        {
            return true;
        }

        return row.Table.TableIndex >= lowerNestedTableBoundaryIndex &&
            (upperNestedTableBoundaryIndex is not { } upper || row.Table.TableIndex < upper);
    }

    private static double VerticalOverlap(double firstY, double firstHeight, double secondY, double secondHeight)
    {
        return Math.Min(firstY + firstHeight, secondY + secondHeight) - Math.Max(firstY, secondY);
    }

    private static bool IsVerticalMergeContinuation(DocxTableCell cell)
    {
        return cell.HasVerticalMerge && !IsVerticalMergeRestart(cell);
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

            if (IsVerticalMergeRestart(previousCell))
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
        int currentGridColumnIndex = 0;
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
        int gridColumnIndex = 0;
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

    private static IReadOnlyList<DocxBodyElement> GetTableCellLayoutBodyElements(DocxTableCell cell)
    {
        IReadOnlyList<DocxBodyElement> authoredElements = DocxTableCellContent.GetBodyElements(cell);
        if (!authoredElements.Any(IsTableCellColumnBreakElement))
        {
            return authoredElements;
        }

        var layoutElements = new List<DocxBodyElement>(authoredElements.Count);
        for (int index = 0; index < authoredElements.Count; index++)
        {
            if (authoredElements[index] is not DocxParagraphElement paragraphElement)
            {
                if (!IsTableCellColumnBreakElement(authoredElements[index]))
                {
                    layoutElements.Add(authoredElements[index]);
                }

                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraph MergeTableCellColumnBreakParagraphs(DocxParagraph first, DocxParagraph second)
            {
                return first with
                {
                    Runs = first.Runs.Concat(second.Runs).ToArray(),
                    Images = first.Images.Concat(second.Images).ToArray(),
                    SpacingAfterPoints = second.EffectiveProperties.SpacingAfterPoints,
                    Spacing = second.EffectiveProperties.Spacing
                };
            }

            while (index + 2 < authoredElements.Count &&
                IsTableCellColumnBreakElement(authoredElements[index + 1]) &&
                authoredElements[index + 2] is DocxParagraphElement continuationElement)
            {
                paragraph = MergeTableCellColumnBreakParagraphs(paragraph, continuationElement.Paragraph);
                index += 2;
            }

            layoutElements.Add(DocxBodyElementFactory.CreateParagraph(paragraph));
        }

        return layoutElements;
    }

    private static IReadOnlyList<DocxParagraph> GetParagraphsFromBodyElements(IReadOnlyList<DocxBodyElement> bodyElements)
    {
        return DocxBlockTraversal.EnumerateDirectParagraphs(bodyElements).ToArray();
    }

    private static bool IsTableCellColumnBreakElement(DocxBodyElement element)
    {
        return element is DocxManualBreakElement manualBreak &&
            manualBreak.Value?.Equals("column", StringComparison.OrdinalIgnoreCase) == true;
    }
}
