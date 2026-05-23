using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<TextRun> RenderTables(PptxRenderContext context, XDocument slideXml, PdfGraphicsBuilder graphics)
    {
        var textRuns = new List<TextRun>();
        foreach (XElement frame in slideXml.Descendants(PresentationNamespace + "graphicFrame"))
        {
            textRuns.AddRange(RenderTableFrame(context, frame, graphics));
        }

        return textRuns;
    }

    private static bool IsTableGraphicFrame(XElement frame)
    {
        return frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData")
            ?.Element(DrawingNamespace + "tbl") is not null;
    }

    private static IReadOnlyList<TextRun> RenderTableFrame(PptxRenderContext context, XElement frame, PdfGraphicsBuilder graphics)
    {
        var textRuns = new List<TextRun>();
        ShapeBounds? bounds = ReadGraphicFrameBounds(frame);
        XElement? table = frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData")
            ?.Element(DrawingNamespace + "tbl");
        if (bounds is null || table is null)
        {
            return textRuns;
        }

        IReadOnlyList<double> rawColumnWidths = table
                .Element(DrawingNamespace + "tblGrid")
                ?.Elements(DrawingNamespace + "gridCol")
                .Select(column => Math.Max(1d, ParseOptionalLongAttribute(column, "w", 1)))
                .ToArray() ?? [];
        IReadOnlyList<XElement> rows = table.Elements(DrawingNamespace + "tr").ToArray();
        if (rawColumnWidths.Count == 0 || rows.Count == 0)
        {
            return textRuns;
        }

        double frameX = OoxUnits.EmuToPoints(bounds.Value.X);
        double frameYTop = OoxUnits.EmuToPoints(bounds.Value.Y);
        double frameWidth = OoxUnits.EmuToPoints(bounds.Value.Width);
        double frameHeight = OoxUnits.EmuToPoints(bounds.Value.Height);
        double frameTop = context.Document.SlideHeightPoints - frameYTop;
        double columnScale = frameWidth / rawColumnWidths.Sum();

        IReadOnlyList<double> rawRowHeights = rows
                .Select(row => Math.Max(1d, ParseOptionalLongAttribute(row, "h", 1)))
                .ToArray();
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
            IReadOnlyList<XElement> cells = rows[rowIndex].Elements(DrawingNamespace + "tc").ToArray();

            double cellX = frameX;
            int columnIndex = 0;
            foreach (XElement cell in cells)
            {
                if (columnIndex >= rawColumnWidths.Count)
                {
                    break;
                }

                if (IsMergedTableCellContinuation(cell))
                {
                    cellX += rawColumnWidths[columnIndex] * columnScale;
                    columnIndex++;
                    continue;
                }

                int columnSpan = Math.Min(ReadTableCellColumnSpan(cell), rawColumnWidths.Count - columnIndex);
                int rowSpan = Math.Min(ReadTableCellRowSpan(cell), rows.Count - rowIndex);
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
                XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");

                bool hasCellFill = TryReadSolidColorWithAlpha(cellProperties, context.Theme, out RgbColor fill, out double fillAlpha) ||
                    TryReadBuiltInTableStyleCellFill(table, rowIndex, columnIndex, rows.Count, rawColumnWidths.Count, context.Theme, out fill, out fillAlpha);
                if (hasCellFill)
                {
                    bool transparentFill = fillAlpha < 0.999d;
                    if (transparentFill)
                    {
                        graphics.SaveState();
                        graphics.SetAlpha(fillAlpha, 1d);
                    }

                    graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                    graphics.FillRectangle(cellX, cellBottom, columnWidth, cellHeight);
                    if (transparentFill)
                    {
                        graphics.RestoreState();
                    }
                }

                AddTableCellBorders(explicitBorders, cellProperties, context.Theme, cellX, cellBottom, columnWidth, cellHeight);
                TableCellTextStyle tableStyleTextStyle = ReadBuiltInTableStyleTextStyle(table, rowIndex, columnIndex, context.Theme);
                AddTableCellTextRuns(cell, cellX, cellBottom, columnWidth, cellHeight, context.Theme, textRuns, tableStyleTextStyle);
                cellX += columnWidth;
                columnIndex += columnSpan;
            }

            yTop -= rowHeight;
            rowTops[rowIndex + 1] = yTop;
        }

        if (!TableHasExplicitBorders(table))
        {
            StrokeDefaultTableGrid(graphics, frameX, frameTop, frameWidth, frameHeight, rawColumnWidths.Select(width => width * columnScale).ToArray(), rowTops, table, skippedVerticalGridSegments, skippedHorizontalGridSegments);
        }
        else
        {
            StrokeTableBorders(graphics, explicitBorders);
        }

        return textRuns;
    }

    private static bool IsMergedTableCellContinuation(XElement cell)
    {
        return ParseOptionalBoolAttribute(cell, "hMerge") ||
            ParseOptionalBoolAttribute(cell, "vMerge");
    }

    private static int ReadTableCellColumnSpan(XElement cell)
    {
        return cell.Attribute("gridSpan") is { } spanAttribute &&
            int.TryParse(spanAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            ? Math.Max(1, span)
            : 1;
    }

    private static int ReadTableCellRowSpan(XElement cell)
    {
        return cell.Attribute("rowSpan") is { } spanAttribute &&
            int.TryParse(spanAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            ? Math.Max(1, span)
            : 1;
    }

    private static bool TableHasExplicitBorders(XElement table)
    {
        return table
            .Descendants(DrawingNamespace + "tcPr")
            .Any(cellProperties =>
                cellProperties.Element(DrawingNamespace + "lnL") is not null ||
                cellProperties.Element(DrawingNamespace + "lnR") is not null ||
                cellProperties.Element(DrawingNamespace + "lnT") is not null ||
                cellProperties.Element(DrawingNamespace + "lnB") is not null);
    }

    private static void StrokeDefaultTableGrid(PdfGraphicsBuilder graphics, double x, double yTop, double width, double height, IReadOnlyList<double> columnWidths, IReadOnlyList<double> rowTops, XElement table, bool[,] skippedVerticalGridSegments, bool[,] skippedHorizontalGridSegments)
    {
        bool hasTableStyle = table
            .Element(DrawingNamespace + "tblPr")
            ?.Element(DrawingNamespace + "tableStyleId") is not null;
        if (hasTableStyle)
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

        IReadOnlyList<XElement> rows = table.Elements(DrawingNamespace + "tr").ToArray();
        for (int i = 0; i < rowTops.Count; i++)
        {
            bool firstRowBoundary = i == 1 &&
                table.Element(DrawingNamespace + "tblPr")?.Attribute("firstRow")?.Value == "1" &&
                rows.Count > 1;
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

    private static void AddTableCellBorders(List<TableBorderLine> borders, XElement? cellProperties, PptxTheme theme, double x, double y, double width, double height)
    {
        AddTableBorder(borders, cellProperties?.Element(DrawingNamespace + "lnL"), theme, x, y, x, y + height);
        AddTableBorder(borders, cellProperties?.Element(DrawingNamespace + "lnR"), theme, x + width, y, x + width, y + height);
        AddTableBorder(borders, cellProperties?.Element(DrawingNamespace + "lnT"), theme, x, y + height, x + width, y + height);
        AddTableBorder(borders, cellProperties?.Element(DrawingNamespace + "lnB"), theme, x, y, x + width, y);
    }

    private static void AddTableBorder(List<TableBorderLine> borders, XElement? line, PptxTheme theme, double x1, double y1, double x2, double y2)
    {
        if (line is null || line.Element(DrawingNamespace + "noFill") is not null || !TryReadSolidColorWithAlpha(line, theme, out RgbColor color, out double alpha))
        {
            return;
        }

        double lineWidth = line.Attribute("w") is { } widthAttribute
            ? Math.Max(1d, OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture)) / 2d)
            : 0.75d;
        borders.Add(new TableBorderLine(x1, y1, x2, y2, lineWidth, color, alpha));
    }

    private static void StrokeTableBorders(PdfGraphicsBuilder graphics, List<TableBorderLine> borders)
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

    private static ShapeBounds? ReadGraphicFrameBounds(XElement frame)
    {
        XElement? transform = frame.Element(PresentationNamespace + "xfrm");
        return transform is null ? null : ReadBoundsFromTransform(transform);
    }

    private static void AddTableCellTextRuns(XElement cell, double x, double y, double width, double height, PptxTheme theme, List<TextRun> runs, TableCellTextStyle tableStyleTextStyle = default)
    {
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        if (textBody is null)
        {
            return;
        }

        TextInsets insets = ReadTextInsets(textBody);
        double textAreaHeight = Math.Max(0d, height - insets.Top - insets.Bottom);
        double verticalOffset = ReadTableCellVerticalAnchor(cell) switch
        {
            TextVerticalAnchor.Middle => textAreaHeight / 2d,
            TextVerticalAnchor.Bottom => textAreaHeight,
            _ => 0d
        };
        double firstFontSize = ReadFirstTableCellFontSize(textBody);
        double cursorY = y + height - insets.Top - firstFontSize + 0.54d - verticalOffset;
        var advanceEstimator = new TextAdvanceEstimator();
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            TextAlignment alignment = ReadAlignment(paragraph, null);
            double textX = x + insets.Left;
            double textWidth = Math.Max(1d, width - insets.Left - insets.Right);
            double cursorX = textX;
            double maxFontSize = 12d;
            foreach (XElement run in paragraph.Elements().Where(IsTextRunElement))
            {
                XElement? runProperties = run.Element(DrawingNamespace + "rPr");
                double fontSize = runProperties?.Attribute("sz") is { } size
                    ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
                    : 12d;
                double alpha = 1d;
                RgbColor color;
                if (TryReadSolidColorWithAlpha(runProperties, theme, out RgbColor runColor, out double runAlpha))
                {
                    color = runColor;
                    alpha = runAlpha;
                }
                else
                {
                    color = tableStyleTextStyle.Color ?? new RgbColor(0, 0, 0);
                }
                string? typeface = theme.ResolveTypeface((string?)runProperties?
                    .Element(DrawingNamespace + "latin")
                    ?.Attribute("typeface"));
                bool bold = tableStyleTextStyle.Bold || ParseOptionalBoolAttribute(runProperties, "b");
                bool italic = ParseOptionalBoolAttribute(runProperties, "i");
                bool underline = ((string?)runProperties?.Attribute("u")) is { } underlineValue
                    && !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
                bool strike = IsStrikeEnabled(runProperties, null);
                foreach (TextCapsFragment fragment in ApplyTextCaps(ReadTextElementText(run, slideNumber: 0), runProperties, null))
                {
                    foreach (string token in SplitTableTextWrapTokens(fragment.Text))
                    {
                        if (token.Length == 0)
                        {
                            continue;
                        }

                        double fragmentFontSize = fontSize * fragment.FontScale;
                        maxFontSize = Math.Max(maxFontSize, fragmentFontSize);
                        double advance = advanceEstimator.Measure(token, fragmentFontSize, typeface, bold, italic, characterSpacing: 0d);
                        if (!string.IsNullOrWhiteSpace(token) &&
                            cursorX > textX + PptxTextMetricRules.TextStateTolerance &&
                            cursorX + advance > textX + textWidth)
                        {
                            cursorY -= maxFontSize * 1.2d;
                            cursorX = textX;
                            maxFontSize = fragmentFontSize;
                        }

                        if (string.IsNullOrWhiteSpace(token) && cursorX <= textX + PptxTextMetricRules.TextStateTolerance)
                        {
                            continue;
                        }

                        runs.Add(new TextRun(token, cursorX, cursorY, Math.Max(1d, advance), textAreaHeight, x, y - height * 0.75d, Math.Max(1d, width), Math.Max(1d, height * 2.1d), fragmentFontSize, 0d, 0d, color, alpha, null, bold, italic, underline, strike, true, alignment, typeface, 0d, 0d, 0d, false, false));
                        cursorX += advance;
                    }
                }
            }

            cursorY -= maxFontSize * 1.2d;
        }
    }

    private static IEnumerable<string> SplitTableTextWrapTokens(string text)
    {
        int start = 0;
        bool? whitespace = null;
        for (int i = 0; i < text.Length; i++)
        {
            bool currentWhitespace = char.IsWhiteSpace(text[i]);
            if (whitespace is null)
            {
                whitespace = currentWhitespace;
                continue;
            }

            if (currentWhitespace != whitespace.Value)
            {
                yield return text[start..i];
                start = i;
                whitespace = currentWhitespace;
            }
        }

        if (start < text.Length)
        {
            yield return text[start..];
        }
    }

    private static double ReadFirstTableCellFontSize(XElement textBody)
    {
        foreach (XElement runProperties in textBody
            .Elements(DrawingNamespace + "p")
            .Elements()
            .Where(IsTextRunElement)
            .Elements(DrawingNamespace + "rPr"))
        {
            if (runProperties.Attribute("sz") is { } size &&
                int.TryParse(size.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int centipoints))
            {
                return Math.Max(1d, centipoints / 100d);
            }
        }

        return 12d;
    }

    private static TextVerticalAnchor ReadTableCellVerticalAnchor(XElement cell)
    {
        string? anchor = (string?)cell
            .Element(DrawingNamespace + "tcPr")
            ?.Attribute("anchor");
        return anchor switch
        {
            "ctr" => TextVerticalAnchor.Middle,
            "b" => TextVerticalAnchor.Bottom,
            _ => TextVerticalAnchor.Top
        };
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
