using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<PptxPositionedTextSpan> ReadSceneTableTextSpans(PptxRenderContext context)
    {
        var textSpans = new List<PptxPositionedTextSpan>();
        AddSceneTableTextSpans(context.SceneSlide?.MasterNodes ?? [], context, textSpans, GroupTransform.Identity);
        AddSceneTableTextSpans(context.SceneSlide?.LayoutNodes ?? [], context, textSpans, GroupTransform.Identity);
        AddSceneTableTextSpans(context.SceneSlide?.SlideNodes ?? [], context, textSpans, GroupTransform.Identity);

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
        return ProcessTableFrame(context, bounds, node.Table?.Source, node.Table, graphics: null);
    }

    private static IReadOnlyList<PptxPositionedTextSpan> RenderTableFrame(PptxRenderContext context, PptxSceneNode node, PdfGraphicsBuilder graphics, GroupTransform transform)
    {
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        return ProcessTableFrame(context, bounds, node.Table?.Source, node.Table, graphics);
    }

    private static IReadOnlyList<PptxPositionedTextSpan> ProcessTableFrame(PptxRenderContext context, ShapeBounds? bounds, XElement? table, PptxSceneTable? sceneTable, PdfGraphicsBuilder? graphics)
    {
        var textSpans = new List<PptxPositionedTextSpan>();
        if (bounds is null || table is null)
        {
            return textSpans;
        }

        IReadOnlyList<double> rawColumnWidths = sceneTable?.ColumnWidths ?? PptxSceneBuilder.ReadTableColumnWidths(table);
        PptxSceneTableStyle tableStyle = sceneTable?.Style ?? PptxSceneBuilder.ReadTableStyle(table);
        IReadOnlyList<XElement> rows = table.Elements(DrawingNamespace + "tr").ToArray();
        if (rawColumnWidths.Count == 0 || rows.Count == 0)
        {
            return textSpans;
        }

        double frameX = OoxUnits.EmuToPoints(bounds.Value.X);
        double frameYTop = OoxUnits.EmuToPoints(bounds.Value.Y);
        double frameWidth = OoxUnits.EmuToPoints(bounds.Value.Width);
        double frameHeight = OoxUnits.EmuToPoints(bounds.Value.Height);
        double frameTop = context.Document.SlideHeightPoints - frameYTop;
        double columnScale = frameWidth / rawColumnWidths.Sum();

        IReadOnlyList<double> rawRowHeights = sceneTable?.RowHeights ?? PptxSceneBuilder.ReadTableRowHeights(table);
        if (rawRowHeights.Count != rows.Count)
        {
            rawRowHeights = PptxSceneBuilder.ReadTableRowHeights(table);
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
            IReadOnlyList<XElement> cells = rows[rowIndex].Elements(DrawingNamespace + "tc").ToArray();

            double cellX = frameX;
            int columnIndex = 0;
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
            {
                XElement cell = cells[cellIndex];
                if (columnIndex >= rawColumnWidths.Count)
                {
                    break;
                }

                PptxSceneTableCell sceneCell = ReadSceneTableCell(sceneTable, rowIndex, cellIndex, columnIndex, cell, context.Theme, tableStyle, rows.Count, rawColumnWidths.Count);
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
                if (hasCellFill && graphics is not null)
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

                AddTableCellBorders(explicitBorders, sceneCell.Borders, cellX, cellBottom, columnWidth, cellHeight);
                AddTableCellTextSpans(context, sceneCell, cellX, cellBottom, columnWidth, cellHeight, textSpans, sceneCell.StyleText);
                cellX += columnWidth;
                columnIndex += columnSpan;
            }

            yTop -= rowHeight;
            rowTops[rowIndex + 1] = yTop;
        }

        if (graphics is null)
        {
            return textSpans;
        }

        if (!TableHasExplicitBorders(sceneTable, table))
        {
            StrokeDefaultTableGrid(graphics, frameX, frameTop, frameWidth, frameHeight, rawColumnWidths.Select(width => width * columnScale).ToArray(), rowTops, tableStyle, rows.Count, skippedVerticalGridSegments, skippedHorizontalGridSegments);
        }
        else
        {
            StrokeTableBorders(graphics, explicitBorders);
        }

        return textSpans;
    }

    private static PptxSceneTableCell ReadSceneTableCell(
        PptxSceneTable? sceneTable,
        int rowIndex,
        int cellIndex,
        int columnIndex,
        XElement cell,
        PptxTheme theme,
        PptxSceneTableStyle tableStyle,
        int rowCount,
        int columnCount)
    {
        if (sceneTable is not null &&
            rowIndex < sceneTable.Rows.Count &&
            cellIndex < sceneTable.Rows[rowIndex].Cells.Count)
        {
            return sceneTable.Rows[rowIndex].Cells[cellIndex];
        }

        return PptxSceneBuilder.ReadTableCell(cell, theme, tableStyle, rowIndex, columnIndex, rowCount, columnCount);
    }

    private static bool TableHasExplicitBorders(PptxSceneTable? sceneTable, XElement table)
    {
        if (sceneTable is not null)
        {
            return sceneTable.Rows
                .SelectMany(row => row.Cells)
                .Any(cell => cell.Borders.HasExplicitBorder);
        }

        return table
            .Descendants(DrawingNamespace + "tcPr")
            .Any(cellProperties =>
                cellProperties.Element(DrawingNamespace + "lnL") is not null ||
                cellProperties.Element(DrawingNamespace + "lnR") is not null ||
                cellProperties.Element(DrawingNamespace + "lnT") is not null ||
                cellProperties.Element(DrawingNamespace + "lnB") is not null);
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

    private static void AddTableCellTextSpans(PptxRenderContext context, PptxSceneTableCell sceneCell, double x, double y, double width, double height, List<PptxPositionedTextSpan> spans, PptxSceneTableCellTextStyle tableStyleTextStyle = default)
    {
        XElement? textBody = sceneCell.TextBody;
        if (textBody is null)
        {
            return;
        }

        XElement tableTextShape = BuildTableCellTextShape(context, textBody, x, y, width, height, ToTextInsets(sceneCell.TextInsets), ToTextVerticalAnchor(sceneCell.VerticalAnchor), tableStyleTextStyle);
        spans.AddRange(ReadTextSpansForShape(tableTextShape, context, includePlaceholders: false));
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

    private static XElement BuildTableCellTextShape(PptxRenderContext context, XElement textBody, double x, double y, double width, double height, TextInsets insets, TextVerticalAnchor anchor, PptxSceneTableCellTextStyle tableStyleTextStyle)
    {
        var textBodyCopy = new XElement(PresentationNamespace + "txBody", textBody.Elements().Select(element => new XElement(element)));
        XElement bodyProperties = textBodyCopy.Element(DrawingNamespace + "bodyPr") ?? new XElement(DrawingNamespace + "bodyPr");
        bodyProperties.SetAttributeValue("lIns", PointsToEmu(insets.Left).ToString(CultureInfo.InvariantCulture));
        bodyProperties.SetAttributeValue("rIns", PointsToEmu(insets.Right).ToString(CultureInfo.InvariantCulture));
        bodyProperties.SetAttributeValue("tIns", PointsToEmu(insets.Top).ToString(CultureInfo.InvariantCulture));
        bodyProperties.SetAttributeValue("bIns", PointsToEmu(insets.Bottom).ToString(CultureInfo.InvariantCulture));
        bodyProperties.SetAttributeValue("vertOverflow", "clip");
        bodyProperties.SetAttributeValue(OoxPdfInternalNamespace + "wrapWidth", Math.Max(1d, width - insets.Left).ToString("R", CultureInfo.InvariantCulture));
        bodyProperties.SetAttributeValue("anchor", anchor switch
        {
            TextVerticalAnchor.Middle => "ctr",
            TextVerticalAnchor.Bottom => "b",
            _ => "t"
        });
        if (bodyProperties.Parent is null)
        {
            textBodyCopy.AddFirst(bodyProperties);
        }

        PruneLeadingEmptyTableParagraphs(textBodyCopy);
        ApplyTableStyleTextDefaults(textBodyCopy, tableStyleTextStyle);
        long shapeX = PointsToEmu(x);
        long shapeY = PointsToEmu(context.Document.SlideHeightPoints - y - height);
        long shapeWidth = PointsToEmu(width);
        long shapeHeight = PointsToEmu(height);
        return new XElement(PresentationNamespace + "sp",
            new XElement(PresentationNamespace + "nvSpPr",
                new XElement(PresentationNamespace + "cNvPr",
                    new XAttribute("id", "0"),
                    new XAttribute("name", "TableCellText")),
                new XElement(PresentationNamespace + "cNvSpPr",
                    new XAttribute("txBox", "1")),
                new XElement(PresentationNamespace + "nvPr")),
            new XElement(PresentationNamespace + "spPr",
                new XElement(DrawingNamespace + "xfrm",
                    new XElement(DrawingNamespace + "off",
                        new XAttribute("x", shapeX.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("y", shapeY.ToString(CultureInfo.InvariantCulture))),
                    new XElement(DrawingNamespace + "ext",
                        new XAttribute("cx", shapeWidth.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("cy", shapeHeight.ToString(CultureInfo.InvariantCulture)))),
                new XElement(DrawingNamespace + "prstGeom",
                    new XAttribute("prst", "rect"),
                    new XElement(DrawingNamespace + "avLst"))),
            textBodyCopy);
    }

    private static void PruneLeadingEmptyTableParagraphs(XElement textBody)
    {
        XElement[] paragraphs = textBody.Elements(DrawingNamespace + "p").ToArray();
        if (!paragraphs.Any(ParagraphHasVisibleContent))
        {
            return;
        }

        foreach (XElement paragraph in paragraphs)
        {
            if (ParagraphHasVisibleContent(paragraph))
            {
                return;
            }

            paragraph.Remove();
        }
    }

    private static void ApplyTableStyleTextDefaults(XElement textBody, PptxSceneTableCellTextStyle style)
    {
        if (style == default)
        {
            return;
        }

        foreach (XElement runProperties in textBody.Descendants(DrawingNamespace + "rPr").Concat(textBody.Descendants(DrawingNamespace + "endParaRPr")))
        {
            if (style.Bold && runProperties.Attribute("b") is null)
            {
                runProperties.SetAttributeValue("b", "1");
            }

            if (style.Color is { } color && !HasTextFill(runProperties))
            {
                runProperties.AddFirst(
                        new XElement(DrawingNamespace + "solidFill",
                            new XElement(DrawingNamespace + "srgbClr",
                            new XAttribute("val", string.Create(CultureInfo.InvariantCulture, $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}")))));
            }
        }
    }

    private static bool HasTextFill(XElement runProperties)
    {
        return runProperties.Element(DrawingNamespace + "solidFill") is not null ||
            runProperties.Element(DrawingNamespace + "noFill") is not null ||
            runProperties.Element(DrawingNamespace + "gradFill") is not null;
    }

    private static long PointsToEmu(double points)
    {
        return (long)Math.Round(points / OoxUnits.PointsPerInch * OoxUnits.EmusPerInch);
    }

    private static double EstimateTableCellTextHeight(XElement textBody, double textWidth, PptxTheme theme, PptxSceneTableCellTextStyle tableStyleTextStyle)
    {
        double height = 0d;
        var advanceEstimator = new TextAdvanceEstimator();
        foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
        {
            double lineWidth = 0d;
            double maxFontSize = 12d;
            bool hasLineContent = false;
            foreach (XElement run in paragraph.Elements().Where(IsTextRunElement))
            {
                XElement? runProperties = run.Element(DrawingNamespace + "rPr");
                double fontSize = runProperties?.Attribute("sz") is { } size
                    ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
                    : 12d;
                string? typeface = theme.ResolveTypeface((string?)runProperties?
                    .Element(DrawingNamespace + "latin")
                    ?.Attribute("typeface"));
                bool bold = tableStyleTextStyle.Bold || ParseOptionalBoolAttribute(runProperties, "b");
                bool italic = ParseOptionalBoolAttribute(runProperties, "i");
                foreach (TextCapsFragment fragment in ApplyTextCaps(ReadTextElementText(run, slideNumber: 0), runProperties, null))
                {
                    foreach (string token in SplitTableTextWrapTokens(fragment.Text))
                    {
                        if (token.Length == 0)
                        {
                            continue;
                        }

                        double fragmentFontSize = fontSize * fragment.FontScale;
                        double advance = advanceEstimator.Measure(token, fragmentFontSize, typeface, bold, italic, characterSpacing: 0d);
                        if (!string.IsNullOrWhiteSpace(token) &&
                            lineWidth > PptxTextMetricRules.TextStateTolerance &&
                            lineWidth + advance > textWidth)
                        {
                            height += maxFontSize * 1.2d;
                            maxFontSize = fragmentFontSize;
                            lineWidth = 0d;
                            hasLineContent = false;
                        }

                        if (string.IsNullOrWhiteSpace(token) && lineWidth <= PptxTextMetricRules.TextStateTolerance)
                        {
                            continue;
                        }

                        maxFontSize = Math.Max(maxFontSize, fragmentFontSize);
                        lineWidth += advance;
                        hasLineContent = true;
                    }
                }
            }

            if (hasLineContent)
            {
                height += maxFontSize * 1.2d;
            }
        }

        return height;
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
