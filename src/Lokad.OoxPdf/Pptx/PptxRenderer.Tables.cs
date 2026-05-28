using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private sealed record TableFrameLayout(
        IReadOnlyList<PptxPositionedTextSpan> TextSpans,
        IReadOnlyList<TableCellFill> CellFills,
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
        AddSceneTableTextSpans(context.SceneSlide.MasterNodes, context, textSpans, GroupTransform.Identity);
        AddSceneTableTextSpans(context.SceneSlide.LayoutNodes, context, textSpans, GroupTransform.Identity);
        AddSceneTableTextSpans(context.SceneSlide.SlideNodes, context, textSpans, GroupTransform.Identity);

        return textSpans;
    }

    private static void AddSceneTableTextSpans(
        IReadOnlyList<PptxSceneNode> nodes,
        PptxRenderContext context,
        List<PptxPositionedTextSpan> textSpans,
        GroupTransform transform)
    {
        foreach (PptxSceneNode node in nodes)
        {
            if (node.Kind == PptxSceneNodeKind.Table)
            {
                textSpans.AddRange(ReadTableFrameTextSpans(context, node, transform));
                continue;
            }

            if (node.Kind == PptxSceneNodeKind.Group)
            {
                AddSceneTableTextSpans(node.Children, context, textSpans, transform.Combine(ToGroupTransform(node.GroupTransform)));
            }
        }
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ReadTableFrameTextSpans(PptxRenderContext context, PptxSceneNode node, GroupTransform transform)
    {
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        return BuildTableFrameLayout(context, bounds, node.Table, emitUnsupportedStyleDiagnostic: false)?.TextSpans ?? [];
    }

    private static IReadOnlyList<PptxPositionedTextSpan> RenderTableFrame(PptxRenderContext context, PptxSceneNode node, PdfGraphicsBuilder graphics, GroupTransform transform)
    {
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        TableFrameLayout? layout = BuildTableFrameLayout(context, bounds, node.Table, emitUnsupportedStyleDiagnostic: true);
        if (layout is null)
        {
            return [];
        }

        RenderTableFrameLayout(graphics, layout);
        return layout.TextSpans;
    }

    private static TableFrameLayout? BuildTableFrameLayout(PptxRenderContext context, ShapeBounds? bounds, PptxSceneTable? sceneTable, bool emitUnsupportedStyleDiagnostic)
    {
        var textSpans = new List<PptxPositionedTextSpan>();
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
                context.Slide.PartName,
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

        double rowScale = frameHeight / rawRowHeights.Sum();

        double yTop = frameTop;
        var rowTops = new double[rows.Count + 1];
        var skippedVerticalGridSegments = new bool[rawColumnWidths.Count + 1, rows.Count];
        var skippedHorizontalGridSegments = new bool[rows.Count + 1, rawColumnWidths.Count];
        var explicitBorders = new List<TableBorderLine>();
        rowTops[0] = yTop;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            double rowHeight = rawRowHeights[rowIndex] * rowScale;
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
                double cellHeight = rawRowHeights
                        .Skip(rowIndex)
                        .Take(rowSpan)
                        .Sum() * rowScale;
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
                AddTableCellTextSpans(context, sceneCell, cellX, cellBottom, columnWidth, cellHeight, textSpans, sceneCell.StyleText);
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
            return new TableFrameLayout(textSpans, cellFills, defaultGrid, []);
        }

        return new TableFrameLayout(textSpans, cellFills, DefaultGrid: null, explicitBorders);
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
            graphics.FillRectangle(fill.X, fill.Y, fill.Width, fill.Height);
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
            StrokeTableBorders(graphics, layout.ExplicitBorders);
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

            graphics.StrokeLine(x, rowTops[startRow] + 0.5d, x, rowTops[rowIndex] - 0.5d);
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

            graphics.StrokeLine(columnLefts[startColumn] - 0.5d, y, columnLefts[columnIndex] + 0.5d, y);
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

    private static void StrokeTableBorders(PdfGraphicsBuilder graphics, IReadOnlyList<TableBorderLine> borders)
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
            double x1 = group.Key.Vertical ? group.Key.FixedCoordinate : start - halfWidth;
            double y1 = group.Key.Vertical ? start - halfWidth : group.Key.FixedCoordinate;
            double x2 = group.Key.Vertical ? group.Key.FixedCoordinate : end + halfWidth;
            double y2 = group.Key.Vertical ? end + halfWidth : group.Key.FixedCoordinate;

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

    private static void AddTableCellTextSpans(PptxRenderContext context, PptxSceneTableCell sceneCell, double x, double y, double width, double height, List<PptxPositionedTextSpan> spans, PptxSceneTableCellTextStyle tableStyleTextStyle = default)
    {
        XElement? textBody = sceneCell.LayoutTextBody;
        if (textBody is null)
        {
            return;
        }

        TextInsets insets = ToTextInsets(sceneCell.TextInsets);
        var tableTextFrame = new PptxTableCellTextFrame(
            textBody,
            x,
            y,
            width,
            height,
            insets,
            ToTextVerticalAnchor(sceneCell.VerticalAnchor),
            ReadTableCellVerticalAnchorValue(sceneCell),
            ToTextBodyPropertySource(sceneCell.VerticalAnchorSource),
            tableStyleTextStyle);
        spans.AddRange(ReadTextSpansForTableCellTextFrame(tableTextFrame, context));
    }

    private static TextInsets ToTextInsets(PptxSceneTextInsets insets)
    {
        return new TextInsets(insets.Left, insets.Right, insets.Top, insets.Bottom);
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
            _ => PptxTextBodyPropertySource.TableCellStyle
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
