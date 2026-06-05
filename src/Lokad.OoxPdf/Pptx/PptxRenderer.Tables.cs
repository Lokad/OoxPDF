using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private const double OfficeTableRowContentExpansionSlackFactor = 1.05d;
    private const double OfficeTableRowSmallPositiveSlackFactor = 1.05d;
    private const double OfficeTableRowDeclaredHeightSlackFactor = 1.10d;
    private const double OfficeMiddleAnchoredTableCellDefaultTopInsetAdjustment = 0.54d;
    private const double OfficeBottomAnchoredTableCellDefaultBottomInsetAdjustment = 0.6d;

    private sealed record TableFrameLayout(
        IReadOnlyList<PptxPositionedTextSpan> TextSpans,
        IReadOnlyList<PptxTableCellTextFrame> TextFrames,
        IReadOnlyList<TableCellFill> CellFills,
        double X,
        double YTop,
        double Width,
        double Height,
        TableDefaultGrid? DefaultGrid,
        IReadOnlyList<TableBorderLine> ExplicitBorders);

    private readonly record struct TableCellFill(double X, double Y, double Width, double Height, RgbColor Color, double Alpha);

    private sealed record TableDefaultGrid(
        double X,
        double YTop,
        double Width,
        double Height,
        IReadOnlyList<double> ColumnWidths,
        IReadOnlyList<double> RowTops,
        PptxSceneTableStyle TableStyle,
        int RowCount,
        bool[,] SkippedVerticalSegments,
        bool[,] SkippedHorizontalSegments);

    private static IReadOnlyList<PptxPositionedTextSpan> ReadSceneTableTextSpans(PptxRenderContext context)
    {
        var textSpans = new List<PptxPositionedTextSpan>();
        AddSceneTableTextSpans(context.SceneSlide.MasterNodes, context, textSpans, GroupTransform.Identity, context.MasterColorMap);
        AddSceneTableTextSpans(context.SceneSlide.LayoutNodes, context, textSpans, GroupTransform.Identity, context.LayoutColorMap);
        AddSceneTableTextSpans(context.SceneSlide.SlideNodes, context, textSpans, GroupTransform.Identity, context.SlideColorMap);

        return textSpans;
    }

    private static IReadOnlyList<PptxTableCellTextFrame> ReadSceneTableTextFrames(PptxRenderContext context)
    {
        var textFrames = new List<PptxTableCellTextFrame>();
        AddSceneTableTextFrames(context.SceneSlide.MasterNodes, context, textFrames, GroupTransform.Identity, context.MasterColorMap);
        AddSceneTableTextFrames(context.SceneSlide.LayoutNodes, context, textFrames, GroupTransform.Identity, context.LayoutColorMap);
        AddSceneTableTextFrames(context.SceneSlide.SlideNodes, context, textFrames, GroupTransform.Identity, context.SlideColorMap);

        return textFrames;
    }

    private static void AddSceneTableTextSpans(
        IReadOnlyList<PptxSceneNode> nodes,
        PptxRenderContext context,
        List<PptxPositionedTextSpan> textSpans,
        GroupTransform transform,
        PptxColorMap colorMap)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Kind == PptxSceneNodeKind.Table)
            {
                textSpans.AddRange(ReadTableFrameTextSpans(context, node, transform, colorMap));
                continue;
            }

            if (node.Kind == PptxSceneNodeKind.Group)
            {
                AddSceneTableTextSpans(node.Children, context, textSpans, transform.Combine(ToGroupTransform(node.GroupTransform)), colorMap);
            }
        }
    }

    private static void AddSceneTableTextFrames(
        IReadOnlyList<PptxSceneNode> nodes,
        PptxRenderContext context,
        List<PptxTableCellTextFrame> textFrames,
        GroupTransform transform,
        PptxColorMap colorMap)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Kind == PptxSceneNodeKind.Table)
            {
                ShapeBounds? bounds = node.Bounds is { } rawBounds
                    ? transform.Apply(ToShapeBounds(rawBounds))
                    : null;
                IReadOnlyList<PptxTableCellTextFrame> tableTextFrames = BuildTableFrameLayout(context, bounds, node.Table, emitUnsupportedStyleDiagnostic: false, colorMap)?.TextFrames ?? [];
                textFrames.AddRange(tableTextFrames);
                continue;
            }

            if (node.Kind == PptxSceneNodeKind.Group)
            {
                AddSceneTableTextFrames(node.Children, context, textFrames, transform.Combine(ToGroupTransform(node.GroupTransform)), colorMap);
            }
        }
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTableFrameTextSpans(PptxRenderContext context, PptxSceneNode node, GroupTransform transform, PptxColorMap colorMap)
    {
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        return BuildTableFrameLayout(context, bounds, node.Table, emitUnsupportedStyleDiagnostic: false, colorMap)?.TextSpans ?? [];
    }

    private static IReadOnlyList<PptxPositionedTextSpan> RenderTableFrame(PptxRenderContext context, PptxSceneNode node, PdfGraphicsBuilder graphics, GroupTransform transform, PptxColorMap colorMap)
    {
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        TableFrameLayout? layout = BuildTableFrameLayout(context, bounds, node.Table, emitUnsupportedStyleDiagnostic: true, colorMap);
        if (layout is null)
        {
            return [];
        }

        RenderTableFrameLayout(graphics, layout);
        return layout.TextSpans;
    }

    private static TableFrameLayout? BuildTableFrameLayout(PptxRenderContext context, ShapeBounds? bounds, PptxSceneTable? sceneTable, bool emitUnsupportedStyleDiagnostic, PptxColorMap colorMap)
    {
        var textSpans = new List<PptxPositionedTextSpan>();
        var textFrames = new List<PptxTableCellTextFrame>();
        var cellFills = new List<TableCellFill>();
        if (bounds is null || sceneTable is null)
        {
            return null;
        }

        IReadOnlyList<double> rawColumnWidths = sceneTable.ColumnWidths;
        PptxSceneTableStyle tableStyle = sceneTable.Style;
        if (emitUnsupportedStyleDiagnostic &&
            tableStyle.HasStyle &&
            !tableStyle.IsSupported)
        {
            context.DiagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "PPTX_UNSUPPORTED_TABLE_STYLE",
                OoxPdfSeverity.Warning,
                "Table style is not in the supported built-in style subset and was rendered without Office table-style cascade formatting.",
                context.SlidePartName,
                context.SlideNumber,
                null,
                "table style",
                "DefaultStyle"));
        }

        IReadOnlyList<PptxSceneTableRow> rows = sceneTable.Rows;
        if (rawColumnWidths.Count == 0 || rows.Count == 0)
        {
            return null;
        }

        double frameX = OoxUnits.EmuToPoints(bounds.Value.X);
        double frameYTop = OoxUnits.EmuToPoints(bounds.Value.Y);
        double frameWidth = OoxUnits.EmuToPoints(bounds.Value.Width);
        double frameHeight = OoxUnits.EmuToPoints(bounds.Value.Height);
        double frameTop = context.Document.SlideHeightPoints - frameYTop;
        double columnScale = frameWidth / rawColumnWidths.Sum();

        IReadOnlyList<double> rawRowHeights = sceneTable.RowHeights;
        if (rawRowHeights.Count != rows.Count)
        {
            return null;
        }

        double declaredTableHeight = rawRowHeights.Sum() * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch;
        double tableHeightSlackFactor = frameHeight / Math.Max(PptxTextMetricRules.TextStateTolerance, declaredTableHeight);
        double rowScale = frameHeight / rawRowHeights.Sum();
        double[] rowHeights = ResolveTableRowHeights(context, sceneTable, rawColumnWidths, rawRowHeights, columnScale, rowScale, frameHeight, colorMap);

        double yTop = frameTop;
        var rowTops = new double[rows.Count + 1];
        var skippedVerticalGridSegments = new bool[rawColumnWidths.Count + 1, rows.Count];
        var skippedHorizontalGridSegments = new bool[rows.Count + 1, rawColumnWidths.Count];
        var explicitBorders = new List<TableBorderLine>();
        rowTops[0] = yTop;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            double rowHeight = rowHeights[rowIndex];
            double cellY = yTop - rowHeight;
            IReadOnlyList<PptxSceneTableCell> cells = rows[rowIndex].Cells;

            double cellX = frameX;
            int columnIndex = 0;
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
            {
                if (columnIndex >= rawColumnWidths.Count)
                {
                    break;
                }

                PptxSceneTableCell sceneCell = cells[cellIndex];
                if (sceneCell.IsMergedContinuation)
                {
                    cellX += rawColumnWidths[columnIndex] * columnScale;
                    columnIndex++;
                    continue;
                }

                int columnSpan = Math.Min(sceneCell.ColumnSpan, rawColumnWidths.Count - columnIndex);
                int rowSpan = Math.Min(sceneCell.RowSpan, rows.Count - rowIndex);
                for (int boundary = columnIndex + 1; boundary < columnIndex + columnSpan; boundary++)
                {
                    skippedVerticalGridSegments[boundary, rowIndex] = true;
                }

                for (int boundary = rowIndex + 1; boundary < rowIndex + rowSpan; boundary++)
                {
                    for (int skippedColumn = columnIndex; skippedColumn < columnIndex + columnSpan; skippedColumn++)
                    {
                        skippedHorizontalGridSegments[boundary, skippedColumn] = true;
                    }
                }

                double columnWidth = rawColumnWidths
                        .Skip(columnIndex)
                        .Take(columnSpan)
                        .Sum() * columnScale;
                double cellHeight = rowHeights
                        .Skip(rowIndex)
                        .Take(rowSpan)
                        .Sum();
                double cellTop = yTop;
                double cellBottom = cellTop - cellHeight;

                bool hasCellFill = sceneCell.Fill.HasFill;
                RgbColor fill = sceneCell.Fill.Color;
                double fillAlpha = sceneCell.Fill.Alpha;
                if (!hasCellFill)
                {
                    hasCellFill = sceneCell.StyleFill.HasFill;
                    fill = sceneCell.StyleFill.Color;
                    fillAlpha = sceneCell.StyleFill.Alpha;
                }
                if (hasCellFill)
                {
                    cellFills.Add(new TableCellFill(cellX, cellBottom, columnWidth, cellHeight, fill, fillAlpha));
                }

                AddTableCellBorders(explicitBorders, sceneCell.Borders, cellX, cellBottom, columnWidth, cellHeight);
                double declaredRowHeight = rawRowHeights[rowIndex] * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch;
                double declaredRowSpanHeight = rawRowHeights
                    .Skip(rowIndex)
                    .Take(rowSpan)
                    .Sum() * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch;
                AddTableCellTextSpans(context, sceneCell, rowIndex, columnIndex, rowSpan, columnSpan, declaredRowHeight, declaredRowSpanHeight, declaredTableHeight, tableHeightSlackFactor, cellX, cellBottom, columnWidth, cellHeight, textSpans, textFrames, colorMap, sceneCell.StyleText);
                cellX += columnWidth;
                columnIndex += columnSpan;
            }

            yTop -= rowHeight;
            rowTops[rowIndex + 1] = yTop;
        }

        if (!TableHasExplicitBorders(sceneTable))
        {
            var defaultGrid = new TableDefaultGrid(
                frameX,
                frameTop,
                frameWidth,
                frameHeight,
                rawColumnWidths.Select(width => width * columnScale).ToArray(),
                rowTops,
                tableStyle,
                rows.Count,
                skippedVerticalGridSegments,
                skippedHorizontalGridSegments);
            return new TableFrameLayout(textSpans, textFrames, cellFills, frameX, frameTop, frameWidth, frameHeight, defaultGrid, []);
        }

        return new TableFrameLayout(textSpans, textFrames, cellFills, frameX, frameTop, frameWidth, frameHeight, DefaultGrid: null, explicitBorders);
    }

    private static double[] ResolveTableRowHeights(
        PptxRenderContext context,
        PptxSceneTable sceneTable,
        IReadOnlyList<double> rawColumnWidths,
        IReadOnlyList<double> rawRowHeights,
        double columnScale,
        double rowScale,
        double frameHeight,
        PptxColorMap colorMap)
    {
        double[] rowHeights = rawRowHeights.Select(height => height * rowScale).ToArray();
        double rowHeightSlackFactor = frameHeight / Math.Max(PptxTextMetricRules.TextStateTolerance, rawRowHeights.Sum() * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch);
        if (rowHeightSlackFactor <= OfficeTableRowContentExpansionSlackFactor)
        {
            if (TryDistributeSmallPositiveTableRowSlack(rawRowHeights, frameHeight, rowHeightSlackFactor, out double[] slackAdjustedRows))
            {
                return slackAdjustedRows;
            }

            if (rowHeightSlackFactor > 1d + PptxTextMetricRules.TextStateTolerance &&
                ShouldKeepDeclaredRowsForSparsePositiveFillSlack(sceneTable))
            {
                return rawRowHeights
                    .Select(height => height * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch)
                    .ToArray();
            }

            return rowHeights;
        }

        if (rowHeightSlackFactor <= OfficeTableRowDeclaredHeightSlackFactor)
        {
            rowHeights = rawRowHeights
                .Select(height => height * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch)
                .ToArray();
        }

        double[] minimumHeights = new double[rowHeights.Length];
        for (int rowIndex = 0; rowIndex < sceneTable.Rows.Count; rowIndex++)
        {
            IReadOnlyList<PptxSceneTableCell> cells = sceneTable.Rows[rowIndex].Cells;
            int columnIndex = 0;
            for (int cellIndex = 0; cellIndex < cells.Count && columnIndex < rawColumnWidths.Count; cellIndex++)
            {
                PptxSceneTableCell sceneCell = cells[cellIndex];
                if (sceneCell.IsMergedContinuation)
                {
                    columnIndex++;
                    continue;
                }

                int columnSpan = Math.Min(sceneCell.ColumnSpan, rawColumnWidths.Count - columnIndex);
                int rowSpan = Math.Min(sceneCell.RowSpan, rowHeights.Length - rowIndex);
                double columnWidth = rawColumnWidths
                    .Skip(columnIndex)
                    .Take(columnSpan)
                    .Sum() * columnScale;
                double minimumHeight = EstimateTableCellMinimumHeight(context, sceneCell, columnWidth, colorMap, sceneCell.StyleText);
                if (minimumHeight > PptxTextMetricRules.TextStateTolerance)
                {
                    if (rowSpan <= 1)
                    {
                        minimumHeights[rowIndex] = Math.Max(minimumHeights[rowIndex], minimumHeight);
                    }
                    else
                    {
                        double currentSpannedHeight = rowHeights.Skip(rowIndex).Take(rowSpan).Sum();
                        double deficit = minimumHeight - currentSpannedHeight;
                        if (deficit > PptxTextMetricRules.TextStateTolerance)
                        {
                            minimumHeights[rowIndex + rowSpan - 1] = Math.Max(
                                minimumHeights[rowIndex + rowSpan - 1],
                                rowHeights[rowIndex + rowSpan - 1] + deficit);
                        }
                    }
                }

                columnIndex += columnSpan;
            }
        }

        for (int i = 0; i < rowHeights.Length; i++)
        {
            rowHeights[i] = Math.Max(rowHeights[i], minimumHeights[i]);
        }

        double minimumTotal = minimumHeights.Sum();
        if (minimumTotal > frameHeight + PptxTextMetricRules.TextStateTolerance)
        {
            return minimumHeights;
        }

        double overflow = rowHeights.Sum() - frameHeight;
        if (overflow > PptxTextMetricRules.TextStateTolerance)
        {
            double shrinkCapacity = rowHeights
                .Select((height, index) => Math.Max(0d, height - minimumHeights[index]))
                .Sum();
            if (shrinkCapacity > PptxTextMetricRules.TextStateTolerance)
            {
                double shrink = Math.Min(overflow, shrinkCapacity);
                for (int i = 0; i < rowHeights.Length; i++)
                {
                    double capacity = Math.Max(0d, rowHeights[i] - minimumHeights[i]);
                    rowHeights[i] -= shrink * capacity / shrinkCapacity;
                }
            }
        }

        return rowHeights;
    }

    private static bool ShouldKeepDeclaredRowsForSparsePositiveFillSlack(PptxSceneTable sceneTable)
    {
        int filledRowCount = 0;
        for (int rowIndex = 0; rowIndex < sceneTable.Rows.Count; rowIndex++)
        {
            bool rowHasFill = sceneTable.Rows[rowIndex].Cells.Any(cell =>
                !cell.IsMergedContinuation &&
                (cell.Fill.HasFill || cell.StyleFill.HasFill));
            if (!rowHasFill)
            {
                continue;
            }

            filledRowCount++;
            if (filledRowCount > 1)
            {
                return false;
            }
        }

        return filledRowCount == 1;
    }

    private static bool TryDistributeSmallPositiveTableRowSlack(
        IReadOnlyList<double> rawRowHeights,
        double frameHeight,
        double rowHeightSlackFactor,
        out double[] rowHeights)
    {
        rowHeights = [];
        if (rawRowHeights.Count <= 1 ||
            rowHeightSlackFactor <= 1d + PptxTextMetricRules.TextStateTolerance ||
            rowHeightSlackFactor > OfficeTableRowSmallPositiveSlackFactor)
        {
            return false;
        }

        double[] declaredRows = rawRowHeights
            .Select(height => height * OoxUnits.PointsPerInch / OoxUnits.EmusPerInch)
            .ToArray();
        double declaredTotal = declaredRows.Sum();
        double slack = frameHeight - declaredTotal;
        if (slack <= PptxTextMetricRules.CoordinateTolerance)
        {
            return false;
        }

        double shortestRow = declaredRows.Min();
        int shortestCount = declaredRows.Count(height => Math.Abs(height - shortestRow) <= PptxTextMetricRules.CoordinateTolerance);
        if (shortestCount <= 1 || shortestCount >= declaredRows.Length)
        {
            return false;
        }

        double slackPerShortestRow = slack / shortestCount;
        for (int i = 0; i < declaredRows.Length; i++)
        {
            if (Math.Abs(declaredRows[i] - shortestRow) <= PptxTextMetricRules.CoordinateTolerance)
            {
                declaredRows[i] += slackPerShortestRow;
            }
        }

        rowHeights = declaredRows;
        return true;
    }

    private static double EstimateTableCellMinimumHeight(
        PptxRenderContext context,
        PptxSceneTableCell sceneCell,
        double width,
        PptxColorMap colorMap,
        PptxSceneTableCellTextStyle tableStyleTextStyle)
    {
        PptxTableCellTextFrame? tableTextFrame = BuildTableCellTextFrame(sceneCell, -1, -1, 1, 1, 1d, 1d, 1d, 1d, 0d, 0d, width, 1d, colorMap, tableStyleTextStyle);
        if (tableTextFrame is null)
        {
            return 0d;
        }

        PptxTextFrameModel frame = BuildTextFrameModel(tableTextFrame, context.Document, context.Theme, context.SlideNumber, context.InheritedXml);
        double textHeight = EstimateTextHeight(frame.Paragraphs, frame.TextWrapWidth, frame.BodyProperties);
        if (textHeight <= PptxTextMetricRules.TextStateTolerance)
        {
            return 0d;
        }

        return frame.Insets.Top + textHeight + frame.Insets.Bottom;
    }

    private static void RenderTableFrameLayout(PdfGraphicsBuilder graphics, TableFrameLayout layout)
    {
        foreach (TableCellFill fill in layout.CellFills)
        {
            bool transparentFill = fill.Alpha < 0.999d;
            if (transparentFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            graphics.FillRectangleEvenOdd(fill.X, fill.Y, fill.Width, fill.Height);
            if (transparentFill)
            {
                graphics.RestoreState();
            }
        }

        if (layout.DefaultGrid is { } defaultGrid)
        {
            StrokeDefaultTableGrid(
                graphics,
                defaultGrid.X,
                defaultGrid.YTop,
                defaultGrid.Width,
                defaultGrid.Height,
                defaultGrid.ColumnWidths,
                defaultGrid.RowTops,
                defaultGrid.TableStyle,
                defaultGrid.RowCount,
                defaultGrid.SkippedVerticalSegments,
                defaultGrid.SkippedHorizontalSegments);
        }
        else
        {
            StrokeTableBorders(
                graphics,
                layout.ExplicitBorders,
                layout.X,
                layout.YTop - layout.Height,
                layout.X + layout.Width,
                layout.YTop);
        }
    }

    private static bool TableHasExplicitBorders(PptxSceneTable sceneTable)
    {
        return sceneTable.Rows
            .SelectMany(row => row.Cells)
            .Any(cell => cell.Borders.HasExplicitBorder);
    }

    private static void StrokeDefaultTableGrid(PdfGraphicsBuilder graphics, double x, double yTop, double width, double height, IReadOnlyList<double> columnWidths, IReadOnlyList<double> rowTops, PptxSceneTableStyle tableStyle, int rowCount, bool[,] skippedVerticalGridSegments, bool[,] skippedHorizontalGridSegments)
    {
        if (tableStyle.HasStyle)
        {
            graphics.SetStrokeRgb(255, 255, 255);
        }
        else
        {
            graphics.SetStrokeRgb(0, 0, 0);
        }

        double cursorX = x;
        graphics.SetLineWidth(1d);
        StrokeDefaultVerticalGridLine(graphics, cursorX, yTop, height, rowTops, skippedVerticalGridSegments, 0);
        for (int columnIndex = 0; columnIndex < columnWidths.Count; columnIndex++)
        {
            cursorX += columnWidths[columnIndex];
            StrokeDefaultVerticalGridLine(graphics, cursorX, yTop, height, rowTops, skippedVerticalGridSegments, columnIndex + 1);
        }

        for (int i = 0; i < rowTops.Count; i++)
        {
            bool firstRowBoundary = i == 1 &&
                tableStyle.FirstRow &&
                rowCount > 1;
            graphics.SetLineWidth(firstRowBoundary ? 3d : 1d);
            StrokeDefaultHorizontalGridLine(graphics, x, width, rowTops[i], columnWidths, skippedHorizontalGridSegments, i);
        }
    }

    private static void StrokeDefaultVerticalGridLine(PdfGraphicsBuilder graphics, double x, double yTop, double height, IReadOnlyList<double> rowTops, bool[,] skippedSegments, int boundaryIndex)
    {
        if (boundaryIndex == 0 || boundaryIndex == skippedSegments.GetLength(0) - 1)
        {
            graphics.StrokeLine(x, yTop + 0.5d, x, yTop - height - 0.5d);
            return;
        }

        int rowCount = skippedSegments.GetLength(1);
        int rowIndex = 0;
        while (rowIndex < rowCount)
        {
            while (rowIndex < rowCount && skippedSegments[boundaryIndex, rowIndex])
            {
                rowIndex++;
            }

            if (rowIndex >= rowCount)
            {
                break;
            }

            int startRow = rowIndex;
            while (rowIndex < rowCount && !skippedSegments[boundaryIndex, rowIndex])
            {
                rowIndex++;
            }

            double y1 = rowTops[startRow] + 0.5d;
            double y2 = rowIndex == rowCount
                ? rowTops[rowIndex]
                : rowTops[rowIndex] - 0.5d;
            graphics.StrokeLine(x, y1, x, y2);
        }
    }

    private static void StrokeDefaultHorizontalGridLine(PdfGraphicsBuilder graphics, double x, double width, double y, IReadOnlyList<double> columnWidths, bool[,] skippedSegments, int boundaryIndex)
    {
        if (boundaryIndex == 0 || boundaryIndex == skippedSegments.GetLength(0) - 1)
        {
            graphics.StrokeLine(x - 0.5d, y, x + width + 0.5d, y);
            return;
        }

        int columnCount = skippedSegments.GetLength(1);
        var columnLefts = new double[columnCount + 1];
        columnLefts[0] = x;
        for (int i = 0; i < columnCount; i++)
        {
            columnLefts[i + 1] = columnLefts[i] + columnWidths[i];
        }

        int columnIndex = 0;
        while (columnIndex < columnCount)
        {
            while (columnIndex < columnCount && skippedSegments[boundaryIndex, columnIndex])
            {
                columnIndex++;
            }

            if (columnIndex >= columnCount)
            {
                break;
            }

            int startColumn = columnIndex;
            while (columnIndex < columnCount && !skippedSegments[boundaryIndex, columnIndex])
            {
                columnIndex++;
            }

            double x1 = startColumn == 0
                ? columnLefts[startColumn]
                : columnLefts[startColumn] - 0.5d;
            double x2 = columnIndex == columnCount
                ? columnLefts[columnIndex]
                : columnLefts[columnIndex] + 0.5d;
            graphics.StrokeLine(x1, y, x2, y);
        }
    }

    private static void AddTableCellBorders(List<TableBorderLine> borders, PptxSceneTableCellBorders cellBorders, double x, double y, double width, double height)
    {
        AddTableBorder(borders, cellBorders.Left, x, y, x, y + height);
        AddTableBorder(borders, cellBorders.Right, x + width, y, x + width, y + height);
        AddTableBorder(borders, cellBorders.Top, x, y + height, x + width, y + height);
        AddTableBorder(borders, cellBorders.Bottom, x, y, x + width, y);
    }

    private static void AddTableBorder(List<TableBorderLine> borders, PptxSceneTableCellBorder border, double x1, double y1, double x2, double y2)
    {
        if (!border.Line.HasLine)
        {
            return;
        }

        borders.Add(new TableBorderLine(x1, y1, x2, y2, border.Line.Width, border.Line.Color, border.Line.Alpha));
    }

    private static void StrokeTableBorders(PdfGraphicsBuilder graphics, IReadOnlyList<TableBorderLine> borders, double tableLeft, double tableBottom, double tableRight, double tableTop)
    {
        foreach (IGrouping<TableBorderKey, TableBorderLine> group in borders.GroupBy(TableBorderKey.From))
        {
            IReadOnlyList<TableBorderLine> ordered = group
                .OrderBy(border => group.Key.Vertical ? Math.Min(border.Y1, border.Y2) : Math.Min(border.X1, border.X2))
                .ToArray();
            double start = group.Key.Vertical
                ? ordered.Min(border => Math.Min(border.Y1, border.Y2))
                : ordered.Min(border => Math.Min(border.X1, border.X2));
            double end = group.Key.Vertical
                ? ordered.Max(border => Math.Max(border.Y1, border.Y2))
                : ordered.Max(border => Math.Max(border.X1, border.X2));
            double halfWidth = group.Key.LineWidth / 2d;
            double x1 = group.Key.Vertical
                ? group.Key.FixedCoordinate
                : ExtendTableBorderEndpoint(start, -halfWidth, tableLeft);
            double y1 = group.Key.Vertical
                ? ExtendTableBorderEndpoint(start, -halfWidth, tableBottom)
                : group.Key.FixedCoordinate;
            double x2 = group.Key.Vertical
                ? group.Key.FixedCoordinate
                : ExtendTableBorderEndpoint(end, halfWidth, tableRight);
            double y2 = group.Key.Vertical
                ? ExtendTableBorderEndpoint(end, halfWidth, tableTop)
                : group.Key.FixedCoordinate;

            bool transparentStroke = group.Key.Alpha < 0.999d;
            if (transparentStroke)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, group.Key.Alpha);
            }

            graphics.SetStrokeRgb(group.Key.Color.Red, group.Key.Color.Green, group.Key.Color.Blue);
            graphics.SetLineWidth(group.Key.LineWidth);
            graphics.StrokeLine(x1, y1, x2, y2);
            if (transparentStroke)
            {
                graphics.RestoreState();
            }
        }
    }

    private static double ExtendTableBorderEndpoint(double coordinate, double extension, double outerBoundary)
    {
        if (Math.Abs(coordinate - outerBoundary) <= PptxTextMetricRules.CoordinateTolerance)
        {
            return coordinate;
        }

        return coordinate + extension;
    }

    private static void AddTableCellTextSpans(
        PptxRenderContext context,
        PptxSceneTableCell sceneCell,
        int rowIndex,
        int columnIndex,
        int rowSpan,
        int columnSpan,
        double declaredRowHeight,
        double declaredRowSpanHeight,
        double declaredTableHeight,
        double tableHeightSlackFactor,
        double x,
        double y,
        double width,
        double height,
        List<PptxPositionedTextSpan> spans,
        List<PptxTableCellTextFrame> textFrames,
        PptxColorMap colorMap,
        PptxSceneTableCellTextStyle tableStyleTextStyle = default)
    {
        PptxTableCellTextFrame? tableTextFrame = BuildTableCellTextFrame(sceneCell, rowIndex, columnIndex, rowSpan, columnSpan, declaredRowHeight, declaredRowSpanHeight, declaredTableHeight, tableHeightSlackFactor, x, y, width, height, colorMap, tableStyleTextStyle);
        if (tableTextFrame is null)
        {
            return;
        }

        textFrames.Add(tableTextFrame);
        spans.AddRange(ReadTextSpansForTableCellTextFrame(tableTextFrame, context));
    }

    private static PptxTableCellTextFrame? BuildTableCellTextFrame(PptxSceneTableCell sceneCell, int rowIndex, int columnIndex, int rowSpan, int columnSpan, double declaredRowHeight, double declaredRowSpanHeight, double declaredTableHeight, double tableHeightSlackFactor, double x, double y, double width, double height, PptxColorMap colorMap, PptxSceneTableCellTextStyle tableStyleTextStyle = default)
    {
        XElement? textBody = sceneCell.LayoutTextBody;
        if (textBody is null)
        {
            return null;
        }

        TextInsets insets = ResolveTableCellTextInsets(sceneCell);
        return new PptxTableCellTextFrame(
            textBody,
            x,
            y,
            width,
            height,
            rowIndex,
            columnIndex,
            rowSpan,
            columnSpan,
            declaredRowHeight,
            declaredRowSpanHeight,
            declaredTableHeight,
            tableHeightSlackFactor,
            insets,
            ToTextInsetSources(sceneCell.TextInsetSources),
            new TextInsetValues(
                sceneCell.TextInsetValues.Left,
                sceneCell.TextInsetValues.Right,
                sceneCell.TextInsetValues.Top,
                sceneCell.TextInsetValues.Bottom),
            ToTextVerticalAnchor(sceneCell.VerticalAnchor),
            ReadTableCellVerticalAnchorValue(sceneCell),
            ToTextBodyPropertySource(sceneCell.VerticalAnchorSource),
            colorMap,
            tableStyleTextStyle);
    }

    private static TextInsets ToTextInsets(PptxSceneTextInsets insets)
    {
        return new TextInsets(insets.Left, insets.Right, insets.Top, insets.Bottom);
    }

    private static TextInsets ResolveTableCellTextInsets(PptxSceneTableCell cell)
    {
        TextInsets insets = ToTextInsets(cell.TextInsets);
        if (cell.VerticalAnchor == PptxSceneTableCellVerticalAnchor.Middle &&
            cell.TextInsetSources.Top == PptxSceneTableCellTextInsetSource.Default)
        {
            return insets with { Top = insets.Top + OfficeMiddleAnchoredTableCellDefaultTopInsetAdjustment };
        }

        if (cell.VerticalAnchor == PptxSceneTableCellVerticalAnchor.Bottom &&
            cell.TextInsetSources.Bottom == PptxSceneTableCellTextInsetSource.Default)
        {
            return insets with { Bottom = insets.Bottom + OfficeBottomAnchoredTableCellDefaultBottomInsetAdjustment };
        }

        return insets;
    }

    private static TextInsetSources ToTextInsetSources(PptxSceneTableCellTextInsetSources sources)
    {
        return new TextInsetSources(
            ToTextBodyPropertySource(sources.Left),
            ToTextBodyPropertySource(sources.Right),
            ToTextBodyPropertySource(sources.Top),
            ToTextBodyPropertySource(sources.Bottom));
    }

    private static PptxTextBodyPropertySource ToTextBodyPropertySource(PptxSceneTableCellTextInsetSource source)
    {
        return source switch
        {
            PptxSceneTableCellTextInsetSource.CellProperties => PptxTextBodyPropertySource.TableCellProperties,
            PptxSceneTableCellTextInsetSource.BodyProperties => PptxTextBodyPropertySource.DirectBodyPr,
            _ => PptxTextBodyPropertySource.DefaultValue
        };
    }

    private static TextVerticalAnchor ToTextVerticalAnchor(PptxSceneTableCellVerticalAnchor anchor)
    {
        return anchor switch
        {
            PptxSceneTableCellVerticalAnchor.Middle => TextVerticalAnchor.Middle,
            PptxSceneTableCellVerticalAnchor.Bottom => TextVerticalAnchor.Bottom,
            _ => TextVerticalAnchor.Top
        };
    }

    private static string ReadTableCellVerticalAnchorValue(PptxSceneTableCell cell)
    {
        if (cell.VerticalAnchorValue is { } value)
        {
            return value;
        }

        return cell.VerticalAnchor switch
        {
            PptxSceneTableCellVerticalAnchor.Middle => "ctr",
            PptxSceneTableCellVerticalAnchor.Bottom => "b",
            _ => "t"
        };
    }

    private static PptxTextBodyPropertySource ToTextBodyPropertySource(PptxSceneTableCellVerticalAnchorSource source)
    {
        return source switch
        {
            PptxSceneTableCellVerticalAnchorSource.CellProperties => PptxTextBodyPropertySource.TableCellProperties,
            _ => PptxTextBodyPropertySource.DefaultValue
        };
    }

    private static long PointsToEmu(double points)
    {
        return (long)Math.Round(points / OoxUnits.PointsPerInch * OoxUnits.EmusPerInch);
    }

    private readonly record struct TableBorderLine(double X1, double Y1, double X2, double Y2, double LineWidth, RgbColor Color, double Alpha);

    private readonly record struct TableBorderKey(bool Vertical, double FixedCoordinate, double LineWidth, RgbColor Color, double Alpha)
    {
        public static TableBorderKey From(TableBorderLine border)
        {
            bool vertical = Math.Abs(border.X1 - border.X2) < 0.001d;
            double fixedCoordinate = vertical ? border.X1 : border.Y1;
            return new TableBorderKey(vertical, Math.Round(fixedCoordinate, 3), Math.Round(border.LineWidth, 3), border.Color, Math.Round(border.Alpha, 5));
        }
    }
}
