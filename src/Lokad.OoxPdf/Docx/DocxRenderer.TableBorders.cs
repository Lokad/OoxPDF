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

// Collapsed table border rendering. Split from the renderer file: row/cell
// border strips (solid/double/triple/compound/3D/wave/segmented), junctions,
// and merge-visibility predicates.
internal sealed partial class DocxRenderer
{
    private static void RenderTableRowBorders(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics,
        CancellationToken cancellationToken,
        RowPairBorderPlan? sharedPlan)
    {
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxTableCellLayout cellLayout = row.Cells[cellIndex];
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow))
            {
                continue;
            }

            if (previousRow is null && IsFirstTableRowFragment(row))
            {
                RenderHorizontalTableCellBorder(cellLayout, "top", graphics);
            }

            if (nextRow is null && IsLastTableRowFragment(row))
            {
                RenderHorizontalTableCellBorder(cellLayout, "bottom", graphics);
            }

            DocxTableCell visualCell = cellLayout.VisualCell;
            DocxTableCellBorder? left = DocxTableBorderGeometry.Find(visualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "start");
            if (cellIndex == 0)
            {
                RenderVerticalTableCellBorder(cellLayout.X, cellLayout.Y, cellLayout.Height, left, graphics);
            }

            DocxTableCellBorder? right = DocxTableBorderGeometry.Find(visualCell.Borders, "right") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "end");
            if (cellIndex == row.Cells.Count - 1)
            {
                RenderVerticalTableCellBorder(cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.Height, right, graphics);
                continue;
            }

            DocxTableCellLayout nextCell = row.Cells[cellIndex + 1];
            DocxTableCell nextVisualCell = nextCell.VisualCell;
            DocxTableCellBorder? nextLeft = DocxTableBorderGeometry.Find(nextVisualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(nextVisualCell.Borders, "start");
            RenderSharedVerticalTableBorder(cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.Height, right, nextLeft, graphics);
        }

        if (sharedPlan is not null)
        {
            RenderSharedHorizontalTableBorders(sharedPlan, graphics, cancellationToken);
        }
    }

    private static bool IsFirstTableRowFragment(DocxTableRowLayout row)
    {
        return row.FragmentIndex <= 0;
    }

    private static bool IsLastTableRowFragment(DocxTableRowLayout row)
    {
        return row.FragmentIndex >= Math.Max(0, row.FragmentCount - 1);
    }

    private static void RenderSharedHorizontalTableBorders(
        RowPairBorderPlan plan,
        PdfGraphicsBuilder graphics,
        CancellationToken cancellationToken)
    {
        int pairIndex = 0;
        for (int cellIndex = 0; cellIndex < plan.Row.Cells.Count; cellIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!plan.CurrentVisible[cellIndex])
            {
                continue;
            }

            if (plan.OverlapCounts[cellIndex] == 0)
            {
                RenderHorizontalTableCellBorder(plan.Row.Cells[cellIndex], "bottom", graphics);
                continue;
            }

            for (int k = 0; k < plan.OverlapCounts[cellIndex]; k++, pairIndex++)
            {
                RenderSharedHorizontalTableBorderSegment(plan.Overlaps[pairIndex].Current, plan.Overlaps[pairIndex].Next, graphics);
            }
        }
    }

    private static void RenderSharedHorizontalTableBorderSegment(
        DocxTableCellLayout cellLayout,
        DocxTableCellLayout nextRowCell,
        PdfGraphicsBuilder graphics)
    {
        DocxTableCell cell = cellLayout.VisualCell;
        DocxTableCell nextCell = nextRowCell.VisualCell;
        DocxTableCellBorder? border = ResolveSharedHorizontalBorder(cellLayout, nextRowCell);
        if (border is null)
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        double x = Math.Max(cellLayout.X, nextRowCell.X);
        double right = Math.Min(cellLayout.X + cellLayout.Width, nextRowCell.X + nextRowCell.Width);
        if (right <= x)
        {
            return;
        }

        double leftBorderWidth = ResolveLeftVerticalBorderWidth(cellLayout);
        double segmentX = Math.Min(right, x + leftBorderWidth / 2d);
        if (right <= segmentX)
        {
            return;
        }

        // Office A/B (w9 exact/atLeast bordered row pairs, Word-COM rendered plus
        // PdfInspect, re-rendered digit-identical): interior horizontal border bands hang
        // below the row boundary by the full nominal width (atLeast mid band at 683.02
        // for boundary 683.52, exact mid band at 683.50 for boundary 684.0), matching
        // the inside-hang convention of the outer top/bottom strips; centering them on
        // the boundary sat 0.26 high with identical text (pitch-neutral paint).
        RenderTableBorderStrip(graphics, border, segmentX, cellLayout.Y - DocxTableBorderGeometry.ResolveNominalWidth(border), right - segmentX, width, DocxTableBorderOrientation.Horizontal);
    }

    private static DocxTableCellBorder? ResolveSharedHorizontalBorder(DocxTableCellLayout cellLayout, DocxTableCellLayout nextRowCell)
    {
        DocxTableCellBorder? bottom = DocxTableBorderGeometry.Find(cellLayout.VisualCell.Borders, "bottom");
        DocxTableCellBorder? nextTop = DocxTableBorderGeometry.Find(nextRowCell.VisualCell.Borders, "top");
        return DocxTableBorderGeometry.IsSuppressed(bottom) || DocxTableBorderGeometry.IsSuppressed(nextTop)
            ? null
            : DocxTableBorderGeometry.SelectStronger(bottom, nextTop);
    }

    private static double HorizontalOverlap(DocxTableCellLayout first, DocxTableCellLayout second)
    {
        return Math.Min(first.X + first.Width, second.X + second.Width) - Math.Max(first.X, second.X);
    }

    private static void RenderHorizontalTableCellBorder(DocxTableCellLayout cellLayout, string edge, PdfGraphicsBuilder graphics)
    {
        DocxTableCell visualCell = cellLayout.VisualCell;
        DocxTableCellBorder? border = DocxTableBorderGeometry.Find(visualCell.Borders, edge);
        if (border is null || DocxTableBorderGeometry.IsSuppressed(border))
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        switch (edge)
        {
            case "top":
                double topX = cellLayout.X + ResolveLeftVerticalBorderWidth(cellLayout) / 2d;
                double topWidth = cellLayout.Width - (topX - cellLayout.X);
                if (topWidth > 0d)
                {
                    RenderTableBorderStrip(graphics, border, topX, cellLayout.Y + cellLayout.Height - width, topWidth, width, DocxTableBorderOrientation.Horizontal);
                }
                break;
            case "bottom":
                double bottomX = cellLayout.X + ResolveLeftVerticalBorderWidth(cellLayout) / 2d;
                double bottomWidth = cellLayout.Width - (bottomX - cellLayout.X);
                if (bottomWidth > 0d)
                {
                    RenderTableBorderStrip(graphics, border, bottomX, cellLayout.Y, bottomWidth, width, DocxTableBorderOrientation.Horizontal);
                }
                break;
        }
    }

    private static double ResolveLeftVerticalBorderWidth(DocxTableCellLayout cellLayout)
    {
        DocxTableCell visualCell = cellLayout.VisualCell;
        DocxTableCellBorder? left = DocxTableBorderGeometry.Find(visualCell.Borders, "left") ??
            DocxTableBorderGeometry.Find(visualCell.Borders, "start");
        return DocxTableBorderGeometry.IsSuppressed(left)
            ? 0d
            : DocxTableBorderGeometry.ResolveVisibleWidth(left);
    }

    private static void RenderSharedVerticalTableBorder(
        double boundaryX,
        double y,
        double height,
        DocxTableCellBorder? leftCellRight,
        DocxTableCellBorder? rightCellLeft,
        PdfGraphicsBuilder graphics)
    {
        if (DocxTableBorderGeometry.IsSuppressed(leftCellRight) || DocxTableBorderGeometry.IsSuppressed(rightCellLeft))
        {
            return;
        }

        DocxTableCellBorder? border = DocxTableBorderGeometry.SelectStronger(leftCellRight, rightCellLeft);
        if (border is null)
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        // Office A/B (w63-w68 probes): vertical bands center on grid lines.
        RenderTableBorderStrip(graphics, border, boundaryX - width / 2d, y, width, height, DocxTableBorderOrientation.Vertical);
    }

    private static void RenderVerticalTableCellBorder(
        double boundaryX,
        double y,
        double height,
        DocxTableCellBorder? border,
        PdfGraphicsBuilder graphics)
    {
        if (border is null || DocxTableBorderGeometry.IsSuppressed(border))
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        // Office A/B (w63-w68 probes): vertical bands center on grid lines.
        RenderTableBorderStrip(graphics, border, boundaryX - width / 2d, y, width, height, DocxTableBorderOrientation.Vertical);
    }

    private static void RenderTableBorderStrip(
        PdfGraphicsBuilder graphics,
        DocxTableCellBorder border,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        string value = border.Value ?? "single";
        if (value.Equals("double", StringComparison.OrdinalIgnoreCase))
        {
            RenderDoubleTableBorderStrip(graphics, x, y, width, height, orientation);
            return;
        }

        if (value.Equals("triple", StringComparison.OrdinalIgnoreCase))
        {
            RenderTripleTableBorderStrip(graphics, x, y, width, height, orientation);
            return;
        }

        if (TryGetCompoundTableBorderProfile(value, out double[] stripeWeights, out double gapWeight))
        {
            RenderCompoundTableBorderStrip(graphics, x, y, width, height, orientation, stripeWeights, gapWeight);
            return;
        }

        if (IsThreeDTableBorderStyle(value))
        {
            RenderThreeDTableBorderStrip(graphics, border, value, x, y, width, height, orientation);
            return;
        }

        if (IsWaveTableBorderStyle(value))
        {
            RenderWaveTableBorderStrip(graphics, border, value, x, y, width, height, orientation);
            return;
        }

        if (value.Equals("dashDotStroked", StringComparison.OrdinalIgnoreCase))
        {
            RenderDashDotStrokedTableBorderStrip(graphics, x, y, width, height, orientation);
            return;
        }

        if (IsSegmentedTableBorderStyle())
        {
            RenderSegmentedTableBorderStrip(graphics, value, x, y, width, height, orientation);
            return;
        }

        graphics.FillRectangle(x, y, width, height);

        bool IsSegmentedTableBorderStyle()
        {
            return value.Equals("dotted", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("dashed", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("dashSmallGap", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("dotDash", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase);
        }
    }

    private static void RenderDoubleTableBorderStrip(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        if (orientation == DocxTableBorderOrientation.Horizontal)
        {
            double stripeHeight = Math.Max(0.12d, height / 3d);
            graphics.FillRectangle(x, y, width, Math.Min(stripeHeight, height));
            graphics.FillRectangle(x, y + Math.Max(0d, height - stripeHeight), width, Math.Min(stripeHeight, height));
            return;
        }

        double stripeWidth = Math.Max(0.12d, width / 3d);
        graphics.FillRectangle(x, y, Math.Min(stripeWidth, width), height);
        graphics.FillRectangle(x + Math.Max(0d, width - stripeWidth), y, Math.Min(stripeWidth, width), height);
    }

    private static void RenderTripleTableBorderStrip(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        double stripe = Math.Min(thickness, Math.Max(0.08d, thickness / 5d));
        double gap = Math.Max(0d, (thickness - 3d * stripe) / 2d);
        for (int index = 0; index < 3; index++)
        {
            double offset = index * (stripe + gap);
            if (orientation == DocxTableBorderOrientation.Horizontal)
            {
                graphics.FillRectangle(x, y + offset, width, Math.Min(stripe, Math.Max(0d, height - offset)));
            }
            else
            {
                graphics.FillRectangle(x + offset, y, Math.Min(stripe, Math.Max(0d, width - offset)), height);
            }
        }
    }

    private static bool TryGetCompoundTableBorderProfile(string value, out double[] stripeWeights, out double gapWeight)
    {
        gapWeight = value.Contains("LargeGap", StringComparison.OrdinalIgnoreCase)
            ? 2d
            : value.Contains("MediumGap", StringComparison.OrdinalIgnoreCase)
                ? 1.25d
                : 0.65d;

        if (value.StartsWith("thinThickThin", StringComparison.OrdinalIgnoreCase))
        {
            stripeWeights = [1d, 2d, 1d];
            return true;
        }

        if (value.StartsWith("thinThick", StringComparison.OrdinalIgnoreCase))
        {
            stripeWeights = [1d, 2d];
            return true;
        }

        if (value.StartsWith("thickThin", StringComparison.OrdinalIgnoreCase))
        {
            stripeWeights = [2d, 1d];
            return true;
        }

        stripeWeights = [];
        return false;
    }

    private static void RenderCompoundTableBorderStrip(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation,
        IReadOnlyList<double> stripeWeights,
        double gapWeight)
    {
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        if (thickness <= 0d || stripeWeights.Count == 0)
        {
            return;
        }

        double totalWeight = stripeWeights.Sum() + gapWeight * Math.Max(0, stripeWeights.Count - 1);
        if (totalWeight <= 0d)
        {
            return;
        }

        double offset = 0d;
        for (int index = 0; index < stripeWeights.Count; index++)
        {
            double stripe = Math.Min(thickness - offset, thickness * stripeWeights[index] / totalWeight);
            if (stripe > 0.001d)
            {
                if (orientation == DocxTableBorderOrientation.Horizontal)
                {
                    graphics.FillRectangle(x, y + offset, width, stripe);
                }
                else
                {
                    graphics.FillRectangle(x + offset, y, stripe, height);
                }
            }

            offset += stripe;
            if (index < stripeWeights.Count - 1)
            {
                offset += thickness * gapWeight / totalWeight;
            }
        }
    }

    private static bool IsThreeDTableBorderStyle(string value)
    {
        return value.Equals("threeDEmboss", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("threeDEngrave", StringComparison.OrdinalIgnoreCase);
    }

    private static void RenderThreeDTableBorderStrip(
        PdfGraphicsBuilder graphics,
        DocxTableCellBorder border,
        string value,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        RgbColor baseColor = ReadColor(border.Color);
        RgbColor light = MixRgbColor(baseColor, new RgbColor(255, 255, 255), 0.55d);
        RgbColor dark = MixRgbColor(baseColor, new RgbColor(0, 0, 0), 0.45d);
        bool emboss = value.Equals("threeDEmboss", StringComparison.OrdinalIgnoreCase);
        RgbColor first = emboss ? light : dark;
        RgbColor second = emboss ? dark : light;

        if (orientation == DocxTableBorderOrientation.Horizontal)
        {
            double firstHeight = Math.Max(0.05d, height * 0.5d);
            double secondHeight = Math.Max(0d, height - firstHeight);
            graphics.SetFillRgb(first.Red, first.Green, first.Blue);
            graphics.FillRectangle(x, y, width, Math.Min(firstHeight, height));
            if (secondHeight > 0.001d)
            {
                graphics.SetFillRgb(second.Red, second.Green, second.Blue);
                graphics.FillRectangle(x, y + firstHeight, width, secondHeight);
            }

            return;
        }

        double firstWidth = Math.Max(0.05d, width * 0.5d);
        double secondWidth = Math.Max(0d, width - firstWidth);
        graphics.SetFillRgb(first.Red, first.Green, first.Blue);
        graphics.FillRectangle(x, y, Math.Min(firstWidth, width), height);
        if (secondWidth > 0.001d)
        {
            graphics.SetFillRgb(second.Red, second.Green, second.Blue);
            graphics.FillRectangle(x + firstWidth, y, secondWidth, height);
        }
    }

    private static bool IsWaveTableBorderStyle(string value)
    {
        return value.Equals("wave", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("doubleWave", StringComparison.OrdinalIgnoreCase);
    }

    private static void RenderWaveTableBorderStrip(
        PdfGraphicsBuilder graphics,
        DocxTableCellBorder border,
        string value,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        RgbColor color = ReadColor(border.Color);
        graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        double lineWidth = Math.Min(thickness, Math.Max(0.08d, thickness / 3d));
        double amplitude = Math.Max(0.08d, Math.Max(0d, thickness - lineWidth) / 2.5d);
        double halfPeriod = Math.Max(1d, thickness * 1.5d);
        graphics.SetLineWidth(lineWidth);

        if (!value.Equals("doubleWave", StringComparison.OrdinalIgnoreCase))
        {
            RenderWaveTableBorderLine(graphics, x, y, width, height, orientation, amplitude, halfPeriod, 0d);
            return;
        }

        double offset = Math.Max(lineWidth, amplitude);
        RenderWaveTableBorderLine(graphics, x, y, width, height, orientation, amplitude * 0.65d, halfPeriod, -offset * 0.45d);
        RenderWaveTableBorderLine(graphics, x, y, width, height, orientation, amplitude * 0.65d, halfPeriod, offset * 0.45d);
    }

    private static void RenderWaveTableBorderLine(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation,
        double amplitude,
        double halfPeriod,
        double minorOffset)
    {
        double majorLength = orientation == DocxTableBorderOrientation.Horizontal ? width : height;
        if (majorLength <= 0.001d)
        {
            return;
        }

        double major = 0d;
        double previousMinor = -amplitude;
        bool high = true;
        while (major < majorLength - 0.001d)
        {
            double nextMajor = Math.Min(major + halfPeriod, majorLength);
            double nextMinor = high ? amplitude : -amplitude;
            if (orientation == DocxTableBorderOrientation.Horizontal)
            {
                double centerY = y + height / 2d + minorOffset;
                graphics.StrokeLine(x + major, centerY + previousMinor, x + nextMajor, centerY + nextMinor);
            }
            else
            {
                double centerX = x + width / 2d + minorOffset;
                graphics.StrokeLine(centerX + previousMinor, y + major, centerX + nextMinor, y + nextMajor);
            }

            major = nextMajor;
            previousMinor = nextMinor;
            high = !high;
        }
    }

    private static void RenderDashDotStrokedTableBorderStrip(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        double dotLength = Math.Max(thickness, 0.35d);
        double dashLength = Math.Max(thickness * 3d, 1d);
        double gapLength = Math.Max(thickness * 1.5d, 0.5d);
        RenderPatternedTableBorderStrip(graphics, [dashLength, dotLength], gapLength, x, y, width, height, orientation);

        double strokeThickness = Math.Min(thickness, Math.Max(0.05d, thickness / 5d));
        if (orientation == DocxTableBorderOrientation.Horizontal)
        {
            graphics.FillRectangle(x, y + Math.Max(0d, (height - strokeThickness) / 2d), width, strokeThickness);
        }
        else
        {
            graphics.FillRectangle(x + Math.Max(0d, (width - strokeThickness) / 2d), y, strokeThickness, height);
        }
    }

    private static void RenderSegmentedTableBorderStrip(
        PdfGraphicsBuilder graphics,
        string value,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        double dotLength = Math.Max(thickness, 0.35d);
        double dashLength = Math.Max(thickness * 3d, 1d);
        double gapLength = value.Equals("dashSmallGap", StringComparison.OrdinalIgnoreCase)
            ? Math.Max(thickness, 0.35d)
            : Math.Max(thickness * 1.5d, 0.5d);

        if (value.Equals("dotDash", StringComparison.OrdinalIgnoreCase))
        {
            RenderPatternedTableBorderStrip(graphics, [dashLength, dotLength], gapLength, x, y, width, height, orientation);
            return;
        }

        if (value.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase))
        {
            RenderPatternedTableBorderStrip(graphics, [dashLength, dotLength, dotLength], gapLength, x, y, width, height, orientation);
            return;
        }

        double segmentLength = value.Equals("dotted", StringComparison.OrdinalIgnoreCase)
            ? dotLength
            : dashLength;
        RenderPatternedTableBorderStrip(graphics, [segmentLength], gapLength, x, y, width, height, orientation);
    }

    private static void RenderPatternedTableBorderStrip(
        PdfGraphicsBuilder graphics,
        IReadOnlyList<double> segmentLengths,
        double gapLength,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        double majorLength = orientation == DocxTableBorderOrientation.Horizontal ? width : height;
        double offset = 0d;
        while (offset < majorLength - 0.001d)
        {
            foreach (double segmentLength in segmentLengths)
            {
                double drawLength = Math.Min(segmentLength, majorLength - offset);
                if (drawLength <= 0.001d)
                {
                    return;
                }

                if (orientation == DocxTableBorderOrientation.Horizontal)
                {
                    graphics.FillRectangle(x + offset, y, drawLength, height);
                }
                else
                {
                    graphics.FillRectangle(x, y + offset, width, drawLength);
                }

                offset += drawLength;
                if (offset >= majorLength - 0.001d)
                {
                    return;
                }

                offset += gapLength;
            }
        }
    }

    private enum DocxTableBorderOrientation
    {
        Horizontal,
        Vertical
    }

    private static bool ShouldRenderTableCellContentFragment(
        DocxTableCellLayout cellLayout,
        DocxTableRowLayout? previousRow)
    {
        return cellLayout.VisualOwnership == DocxTableCellVisualOwnership.OwnCell ||
            cellLayout.VisualOwnership == DocxTableCellVisualOwnership.VerticalMergeOwner &&
            ShouldRenderTableCellVisualFragment(cellLayout, previousRow);
    }

    private static bool ShouldRenderTableCellVisualFragment(
        DocxTableCellLayout cellLayout,
        DocxTableRowLayout? previousRow)
    {
        if (cellLayout.VisualOwnership == DocxTableCellVisualOwnership.OwnCell)
        {
            return true;
        }

        if (cellLayout.VisualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner)
        {
            return false;
        }

        return previousRow is null || !previousRow.Cells.Any(previousCell =>
            !previousCell.IsVerticalMergeContinuation &&
            HorizontalOverlap(previousCell, cellLayout) > 0d &&
            previousCell.Y <= cellLayout.Y + 0.001d &&
            previousCell.Y + previousCell.Height >= cellLayout.Y + cellLayout.Height - 0.001d);
    }

    // R11: one ordered row-pair overlap plan shared by the horizontal stroke and
    // junction painters. Both painters previously rescanned next-row cells per
    // current-row cell (O(C^2) overlap searches plus a per-cell overlap allocation),
    // re-evaluating merge-visibility scans inside every pair test. The sweep evaluates
    // each cell visibility once with the exact predicates both painters use, discovers
    // every overlapping pair with binary searches, and emits pairs in nested-loop
    // order so shared output is unchanged.
    internal sealed class RowPairBorderPlan
    {
        public DocxTableRowLayout Row { get; }

        public DocxTableRowLayout NextRow { get; }

        public bool[] CurrentVisible { get; }

        public int[] OverlapCounts { get; }

        public (DocxTableCellLayout Current, DocxTableCellLayout Next, double X, double Right)[] Overlaps { get; }

        private RowPairBorderPlan(
            DocxTableRowLayout row,
            DocxTableRowLayout nextRow,
            bool[] currentVisible,
            int[] overlapCounts,
            (DocxTableCellLayout Current, DocxTableCellLayout Next, double X, double Right)[] overlaps)
        {
            Row = row;
            NextRow = nextRow;
            CurrentVisible = currentVisible;
            OverlapCounts = overlapCounts;
            Overlaps = overlaps;
        }

        public static RowPairBorderPlan? TryBuild(DocxTableRowLayout row, DocxTableRowLayout? nextRow, CancellationToken cancellationToken)
        {
            if (nextRow is null || nextRow.RowIndex == row.RowIndex)
            {
                return null;
            }

            bool[] currentVisible = new bool[row.Cells.Count];
            for (int i = 0; i < row.Cells.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                currentVisible[i] = ShouldRenderTableCellVisualFragment(row.Cells[i], previousRow: null);
            }

            bool[] nextVisible = new bool[nextRow.Cells.Count];
            for (int j = 0; j < nextRow.Cells.Count; j++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                nextVisible[j] = ShouldRenderTableCellVisualFragment(nextRow.Cells[j], row);
            }

            double[] nextLefts = new double[nextRow.Cells.Count];
            double[] nextRights = new double[nextRow.Cells.Count];
            bool boundsFinite = true;
            for (int j = 0; j < nextRow.Cells.Count; j++)
            {
                double left = nextRow.Cells[j].X;
                double right = left + nextRow.Cells[j].Width;
                nextLefts[j] = left;
                nextRights[j] = right;
                boundsFinite &= double.IsFinite(left) && double.IsFinite(right);
            }

            var pairs = new List<(int CurrentIndex, int NextIndex, double X, double Right)>();
            if (boundsFinite)
            {
                int[] nextOrder = new int[nextRow.Cells.Count];
                for (int j = 0; j < nextOrder.Length; j++)
                {
                    nextOrder[j] = j;
                }

                Array.Sort(nextOrder, (a, b) => nextLefts[a].CompareTo(nextLefts[b]));
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!currentVisible[i])
                    {
                        continue;
                    }

                    double left = row.Cells[i].X;
                    double right = left + row.Cells[i].Width;
                    if (!double.IsFinite(left) || !double.IsFinite(right))
                    {
                        AddNestedPairs(row, nextRow, currentVisible, nextVisible, i, pairs);
                        continue;
                    }

                    int lo = 0;
                    int hi = nextOrder.Length;
                    while (lo < hi)
                    {
                        int mid = lo + ((hi - lo) >> 1);
                        if (nextRights[nextOrder[mid]] > left)
                        {
                            hi = mid;
                        }
                        else
                        {
                            lo = mid + 1;
                        }
                    }

                    for (int k = lo; k < nextOrder.Length && nextLefts[nextOrder[k]] < right; k++)
                    {
                        int j = nextOrder[k];
                        if (!nextVisible[j])
                        {
                            continue;
                        }

                        double overlapLeft = Math.Max(left, nextLefts[j]);
                        double overlapRight = Math.Min(right, nextRights[j]);
                        if (overlapRight > overlapLeft)
                        {
                            pairs.Add((i, j, overlapLeft, overlapRight));
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddNestedPairs(row, nextRow, currentVisible, nextVisible, i, pairs);
                }
            }

            pairs.Sort((a, b) =>
            {
                int order = a.CurrentIndex.CompareTo(b.CurrentIndex);
                return order != 0 ? order : a.NextIndex.CompareTo(b.NextIndex);
            });

            int[] overlapCounts = new int[row.Cells.Count];
            var overlaps = new (DocxTableCellLayout Current, DocxTableCellLayout Next, double X, double Right)[pairs.Count];
            for (int k = 0; k < pairs.Count; k++)
            {
                overlapCounts[pairs[k].CurrentIndex]++;
                overlaps[k] = (row.Cells[pairs[k].CurrentIndex], nextRow.Cells[pairs[k].NextIndex], pairs[k].X, pairs[k].Right);
            }

            return new RowPairBorderPlan(row, nextRow, currentVisible, overlapCounts, overlaps);
        }

        private static void AddNestedPairs(
            DocxTableRowLayout row,
            DocxTableRowLayout nextRow,
            bool[] currentVisible,
            bool[] nextVisible,
            int currentIndex,
            List<(int CurrentIndex, int NextIndex, double X, double Right)> pairs)
        {
            if (!currentVisible[currentIndex])
            {
                return;
            }

            DocxTableCellLayout current = row.Cells[currentIndex];
            for (int j = 0; j < nextRow.Cells.Count; j++)
            {
                if (!nextVisible[j])
                {
                    continue;
                }

                double overlapLeft = Math.Max(current.X, nextRow.Cells[j].X);
                double overlapRight = Math.Min(current.X + current.Width, nextRow.Cells[j].X + nextRow.Cells[j].Width);
                if (overlapRight > overlapLeft)
                {
                    pairs.Add((currentIndex, j, overlapLeft, overlapRight));
                }
            }
        }
    }

    // R11: X-ordered index over concatenated vertical-boundary arrays. Grouping and
    // junction emission keep concat order, so per-pair slices reproduce the original
    // filter/group output exactly with binary searches instead of full scans.
    private sealed class OrderedBoundaryIndex
    {
        private readonly DocxTableBorderBoundary[] ordered;
        private readonly int[] byX;

        private OrderedBoundaryIndex(DocxTableBorderBoundary[] ordered, int[] byX)
        {
            this.ordered = ordered;
            this.byX = byX;
        }

        public static OrderedBoundaryIndex Build(IReadOnlyList<DocxTableBorderBoundary> first, IReadOnlyList<DocxTableBorderBoundary>? second = null)
        {
            int firstCount = first.Count;
            int secondCount = second is null ? 0 : second.Count;
            var ordered = new DocxTableBorderBoundary[firstCount + secondCount];
            for (int i = 0; i < firstCount; i++)
            {
                ordered[i] = first[i];
            }

            if (second is not null)
            {
                for (int j = 0; j < secondCount; j++)
                {
                    ordered[firstCount + j] = second[j];
                }
            }

            int[] byX = new int[ordered.Length];
            for (int i = 0; i < byX.Length; i++)
            {
                byX[i] = i;
            }

            Array.Sort(byX, (a, b) => ordered[a].X.CompareTo(ordered[b].X));
            return new OrderedBoundaryIndex(ordered, byX);
        }

        public DocxTableBorderBoundary[] SliceRaw(double x, double right, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int lo = LowerBound(x - 0.001d);
            int hi = UpperBound(right + 0.001d);
            if (hi <= lo)
            {
                return [];
            }

            int[] positions = new int[hi - lo];
            for (int k = lo; k < hi; k++)
            {
                positions[k - lo] = byX[k];
            }

            Array.Sort(positions);
            var slice = new DocxTableBorderBoundary[positions.Length];
            for (int k = 0; k < positions.Length; k++)
            {
                slice[k] = ordered[positions[k]];
            }

            return slice;
        }

        public DocxTableBorderBoundary[] SliceGrouped(double x, double right, CancellationToken cancellationToken)
        {
            return SliceRaw(x, right, cancellationToken)
                .Where(boundary => boundary.X >= x - 0.001d && boundary.X <= right + 0.001d)
                .GroupBy(boundary => Math.Round(boundary.X, 3))
                .Select(group => group.OrderByDescending(boundary => boundary.Width).First())
                .ToArray();
        }

        private int LowerBound(double value)
        {
            int lo = 0;
            int hi = byX.Length;
            while (lo < hi)
            {
                int mid = lo + ((hi - lo) >> 1);
                if (ordered[byX[mid]].X >= value)
                {
                    hi = mid;
                }
                else
                {
                    lo = mid + 1;
                }
            }

            return lo;
        }

        private int UpperBound(double value)
        {
            int lo = 0;
            int hi = byX.Length;
            while (lo < hi)
            {
                int mid = lo + ((hi - lo) >> 1);
                if (ordered[byX[mid]].X > value)
                {
                    hi = mid;
                }
                else
                {
                    lo = mid + 1;
                }
            }

            return lo;
        }
    }

    private sealed record DocxTableBorderBoundary(double X, double Width, DocxTableCellBorder Border);
}
