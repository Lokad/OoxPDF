using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static void RenderTableRow(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        CancellationToken cancellationToken,
        ref int imageIndex)
    {
        // Office A/B (comment-table and w6-celltuckborders probes, Word-COM rendered):
        // the first-pin emission shift moves text but table geometry was painted raw, so
        // cell text rendered a full offset below its rows and clips. Geometry follows the
        // same text offset (a no-op outside the Word-compatible profile, where it is 0).
        double geometryXOffset = ResolveTextEmissionXOffset(markupContext);
        double geometryYOffset = ResolveTextEmissionBaselineOffset(markupContext);
        DocxTableRowLayout geometryRow = geometryXOffset == 0d && geometryYOffset == 0d ? row : ShiftTableRowGeometry(row, geometryXOffset, geometryYOffset);
        DocxTableRowLayout? geometryPreviousRow = previousRow is null || (geometryXOffset == 0d && geometryYOffset == 0d) ? previousRow : ShiftTableRowGeometry(previousRow, geometryXOffset, geometryYOffset);
        DocxTableRowLayout? geometryNextRow = nextRow is null || (geometryXOffset == 0d && geometryYOffset == 0d) ? nextRow : ShiftTableRowGeometry(nextRow, geometryXOffset, geometryYOffset);
        foreach (DocxTableCellLayout cellLayout in geometryRow.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ShouldRenderTableCellVisualFragment(cellLayout, geometryPreviousRow))
            {
                continue;
            }

            DocxTableCell cell = cellLayout.VisualCell;
            RenderShadingFill(cell.FillHex, cell.ShadingValue, cell.ShadingColor, graphics, cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
        }

        RowPairBorderPlan? sharedBorderPlan = RowPairBorderPlan.TryBuild(geometryRow, geometryNextRow, cancellationToken);
        RenderTableRowBorders(geometryRow, geometryPreviousRow, geometryNextRow, graphics, cancellationToken, sharedBorderPlan);
        RenderTableBorderJunctions(geometryRow, geometryPreviousRow, geometryNextRow, graphics, cancellationToken, sharedBorderPlan);
        RenderTableRowMarkupIndicators(row, graphics, markupContext);

        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxTableCellLayout cellLayout = row.Cells[cellIndex];
            if (!ShouldRenderTableCellContentFragment(cellLayout, previousRow))
            {
                continue;
            }

            if (cellLayout.TextLines.Count != 0 || cellLayout.InlineImages.Count != 0 || cellLayout.InlineTextBoxes.Count != 0 || cellLayout.NestedRows.Count != 0)
            {
                DocxTableCellLayout geometryCell = geometryRow.Cells[cellIndex];
                graphics.SaveState();
                graphics.ClipRectangle(geometryCell.X, geometryCell.Y, geometryCell.Width, geometryCell.Height);
                foreach (DocxTextLineLayout line in cellLayout.TextLines)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    RenderTextLine(line, graphics, fontResources, markupContext, pageNumber, pageCount, cancellationToken);
                }

                foreach (DocxInlineImageLayout image in cellLayout.InlineImages)
                {
                    RenderInlineImage(geometryXOffset == 0d && geometryYOffset == 0d ? image : image with { X = image.X + geometryXOffset, Y = image.Y - geometryYOffset }, graphics, pageImages, diagnosticSink, cancellationToken, ref imageIndex);
                }

                foreach (DocxInlineTextBoxLayout textBox in cellLayout.InlineTextBoxes)
                {
                    RenderInlineTextBox(textBox, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, ref imageIndex);
                }

                for (int nestedRowIndex = 0; nestedRowIndex < cellLayout.NestedRows.Count; nestedRowIndex++)
                {
                    DocxTableRowLayout nestedRow = cellLayout.NestedRows[nestedRowIndex];
                    DocxTableRowLayout? previousNestedRow = nestedRowIndex > 0 ? cellLayout.NestedRows[nestedRowIndex - 1] : null;
                    DocxTableRowLayout? nextNestedRow = nestedRowIndex + 1 < cellLayout.NestedRows.Count ? cellLayout.NestedRows[nestedRowIndex + 1] : null;
                    RenderTableRow(
                        nestedRow,
                        IsAdjacentTableRow(previousNestedRow, nestedRow) ? previousNestedRow : null,
                        IsAdjacentTableRow(nestedRow, nextNestedRow) ? nextNestedRow : null,
                        graphics,
                        pageImages,
                        fontResources,
                        markupContext,
                        diagnosticSink,
                        pageNumber,
                        pageCount,
                        cancellationToken,
                        ref imageIndex);
                }

                graphics.RestoreState();
            }
        }
    }

    private static DocxTableRowLayout ShiftTableRowGeometry(DocxTableRowLayout row, double xOffset, double yOffset)
    {
        return row with
        {
            Y = row.Y - yOffset,
            Cells = row.Cells.Select(cell => cell with { X = cell.X + xOffset, Y = cell.Y - yOffset }).ToArray(),
        };
    }
    private static void RenderTableRowMarkupIndicators(
        DocxTableRowLayout row,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        if (!markupContext.DrawsChangeBars ||
            row.RevisionCount == 0 ||
            UsesWordCompatibleAllMarkupTextProfile(markupContext))
        {
            return;
        }

        double height = Math.Max(6d, row.Height);
        DocxMarkupBalloonRgb color = ResolveRevisionAuthorColor(row.Revisions ?? []);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        graphics.FillRectangle(Math.Max(0d, row.Table.TableX - 7d), row.Y, 1.5d, height);
    }

    private static void RenderTableBorderJunctions(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics,
        CancellationToken cancellationToken,
        RowPairBorderPlan? sharedPlan)
    {
        DocxTableBorderBoundary[] rowBoundaries = ResolveVisibleVerticalBoundaries(row, previousRow, cancellationToken);
        if (rowBoundaries.Length == 0)
        {
            return;
        }

        OrderedBoundaryIndex rowBoundaryIndex = OrderedBoundaryIndex.Build(rowBoundaries);
        var emittedJunctions = new HashSet<(double X, double Y)>();
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow))
            {
                continue;
            }

            if (previousRow is null && IsFirstTableRowFragment(row))
            {
                RenderHorizontalBorderJunctions(cellLayout.X, cellLayout.X + cellLayout.Width, cellLayout.Y + cellLayout.Height, cellLayout.VisualCell, "top", rowBoundaryIndex.SliceRaw(cellLayout.X, cellLayout.X + cellLayout.Width, cancellationToken), graphics, emittedJunctions);
            }

            if (nextRow is null && IsLastTableRowFragment(row))
            {
                RenderHorizontalBorderJunctions(cellLayout.X, cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.VisualCell, "bottom", rowBoundaryIndex.SliceRaw(cellLayout.X, cellLayout.X + cellLayout.Width, cancellationToken), graphics, emittedJunctions);
            }
        }

        if (previousRow is null && IsFirstTableRowFragment(row))
        {
            RenderOuterHorizontalFragmentCornerJunctions(row, previousRow, rowBoundaries, "top", graphics);
        }

        if (nextRow is null && IsLastTableRowFragment(row))
        {
            RenderOuterHorizontalFragmentCornerJunctions(row, previousRow, rowBoundaries, "bottom", graphics);
        }

        if (sharedPlan is not null)
        {
            RenderSharedHorizontalBorderJunctions(sharedPlan, rowBoundaries, rowBoundaryIndex, graphics, emittedJunctions, cancellationToken);
        }
    }

    private static void RenderOuterHorizontalFragmentCornerJunctions(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        IReadOnlyList<DocxTableBorderBoundary> boundaries,
        string edge,
        PdfGraphicsBuilder graphics)
    {
        DocxTableCellLayout? firstCell = row.Cells.FirstOrDefault(cell => ShouldRenderTableCellVisualFragment(cell, previousRow));
        DocxTableCellLayout? lastCell = row.Cells.LastOrDefault(cell => ShouldRenderTableCellVisualFragment(cell, previousRow));
        if (firstCell is null || lastCell is null)
        {
            return;
        }

        DocxTableCellBorder? firstHorizontal = DocxTableBorderGeometry.Find(firstCell.VisualCell.Borders, edge);
        DocxTableCellBorder? lastHorizontal = DocxTableBorderGeometry.Find(lastCell.VisualCell.Borders, edge);
        DocxTableBorderBoundary? firstBoundary = boundaries.OrderBy(boundary => boundary.X).FirstOrDefault();
        DocxTableBorderBoundary? lastBoundary = boundaries.OrderByDescending(boundary => boundary.X).FirstOrDefault();
        double y = string.Equals(edge, "top", StringComparison.Ordinal)
            ? firstCell.Y + firstCell.Height
            : firstCell.Y;

        if (firstBoundary is not null && firstHorizontal is not null && !DocxTableBorderGeometry.IsSuppressed(firstHorizontal))
        {
            RenderBorderJunctions([firstBoundary], y, firstHorizontal, graphics, []);
        }

        if (lastBoundary is not null && lastHorizontal is not null && !DocxTableBorderGeometry.IsSuppressed(lastHorizontal))
        {
            RenderBorderJunctions([lastBoundary], y, lastHorizontal, graphics, []);
        }
    }

    private static void RenderSharedHorizontalBorderJunctions(
        RowPairBorderPlan plan,
        IReadOnlyList<DocxTableBorderBoundary> rowBoundaries,
        OrderedBoundaryIndex rowBoundaryIndex,
        PdfGraphicsBuilder graphics,
        HashSet<(double X, double Y)> emittedJunctions,
        CancellationToken cancellationToken)
    {
        DocxTableBorderBoundary[] nextRowBoundaries = ResolveVisibleVerticalBoundaries(plan.NextRow, plan.Row, cancellationToken);
        OrderedBoundaryIndex mergedBoundaries = OrderedBoundaryIndex.Build(rowBoundaries, nextRowBoundaries);
        int pairIndex = 0;
        for (int cellIndex = 0; cellIndex < plan.Row.Cells.Count; cellIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!plan.CurrentVisible[cellIndex])
            {
                continue;
            }

            DocxTableCellLayout cellLayout = plan.Row.Cells[cellIndex];
            if (plan.OverlapCounts[cellIndex] == 0)
            {
                RenderHorizontalBorderJunctions(cellLayout.X, cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.VisualCell, "bottom", rowBoundaryIndex.SliceRaw(cellLayout.X, cellLayout.X + cellLayout.Width, cancellationToken), graphics, emittedJunctions);
                continue;
            }

            for (int k = 0; k < plan.OverlapCounts[cellIndex]; k++, pairIndex++)
            {
                (DocxTableCellLayout current, DocxTableCellLayout nextRowCell, double x, double right) = plan.Overlaps[pairIndex];
                DocxTableCellBorder? horizontal = ResolveSharedHorizontalBorder(current, nextRowCell);
                if (horizontal is null)
                {
                    continue;
                }

                if (right <= x)
                {
                    continue;
                }

                DocxTableBorderBoundary[] boundaries = mergedBoundaries.SliceGrouped(x, right, cancellationToken);
                RenderBorderJunctions(boundaries, current.Y - DocxTableBorderGeometry.ResolveVisibleWidth(horizontal) / 2d, horizontal, graphics, emittedJunctions);
            }
        }
    }

    private static void RenderHorizontalBorderJunctions(
        double x,
        double right,
        double y,
        DocxTableCell cell,
        string edge,
        IReadOnlyList<DocxTableBorderBoundary> boundaries,
        PdfGraphicsBuilder graphics,
        HashSet<(double X, double Y)> emittedJunctions)
    {
        DocxTableCellBorder? horizontal = DocxTableBorderGeometry.Find(cell.Borders, edge);
        if (horizontal is null || DocxTableBorderGeometry.IsSuppressed(horizontal))
        {
            return;
        }

        DocxTableBorderBoundary[] crossingBoundaries = boundaries
            .Where(boundary => boundary.X >= x - 0.001d && boundary.X <= right + 0.001d)
            .ToArray();
        RenderBorderJunctions(crossingBoundaries, y, horizontal, graphics, emittedJunctions);
    }

    private static void RenderBorderJunctions(
        IReadOnlyList<DocxTableBorderBoundary> boundaries,
        double y,
        DocxTableCellBorder horizontal,
        PdfGraphicsBuilder graphics,
        HashSet<(double X, double Y)> emittedJunctions)
    {
        double horizontalWidth = DocxTableBorderGeometry.ResolveVisibleWidth(horizontal);
        if (horizontalWidth <= 0d)
        {
            return;
        }

        foreach (DocxTableBorderBoundary boundary in boundaries)
        {
            if (!emittedJunctions.Add((Math.Round(boundary.X, 3), Math.Round(y, 3))))
            {
                continue;
            }

            DocxTableCellBorder? border = DocxTableBorderGeometry.SelectStronger(horizontal, boundary.Border);
            if (border is null)
            {
                continue;
            }

            RgbColor color = ReadColor(border.Color);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            RenderTableBorderStrip(graphics, border, boundary.X, y, boundary.Width, horizontalWidth, DocxTableBorderOrientation.Horizontal);
        }
    }

    private static DocxTableBorderBoundary[] ResolveVisibleVerticalBoundaries(DocxTableRowLayout row, DocxTableRowLayout? previousRow, CancellationToken cancellationToken)
    {
        var boundaries = new List<DocxTableBorderBoundary>();
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxTableCellLayout cellLayout = row.Cells[cellIndex];
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow))
            {
                continue;
            }

            DocxTableCell visualCell = cellLayout.VisualCell;
            DocxTableCellBorder? left = DocxTableBorderGeometry.Find(visualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "start");
            if (cellIndex == 0)
            {
                AddVerticalBoundary(boundaries, cellLayout.X, left);
            }

            DocxTableCellBorder? right = DocxTableBorderGeometry.Find(visualCell.Borders, "right") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "end");
            if (cellIndex == row.Cells.Count - 1)
            {
                AddVerticalBoundary(boundaries, cellLayout.X + cellLayout.Width, right);
                continue;
            }

            DocxTableCellLayout nextCell = row.Cells[cellIndex + 1];
            DocxTableCell nextVisualCell = nextCell.VisualCell;
            DocxTableCellBorder? nextLeft = DocxTableBorderGeometry.Find(nextVisualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(nextVisualCell.Borders, "start");
            if (!DocxTableBorderGeometry.IsSuppressed(right) && !DocxTableBorderGeometry.IsSuppressed(nextLeft))
            {
                AddVerticalBoundary(boundaries, cellLayout.X + cellLayout.Width, DocxTableBorderGeometry.SelectStronger(right, nextLeft));
            }
        }

        return boundaries.ToArray();
    }

    private static void AddVerticalBoundary(List<DocxTableBorderBoundary> boundaries, double x, DocxTableCellBorder? border)
    {
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        if (width <= 0d || border is null)
        {
            return;
        }

        // Office A/B (w63-w68 probes): junctions ride the centered vertical bands.
        boundaries.Add(new DocxTableBorderBoundary(x - width / 2d, width, border));
    }
}
