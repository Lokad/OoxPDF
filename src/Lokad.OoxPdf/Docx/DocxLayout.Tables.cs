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
    private static void LayoutTable(
        DocxTable table,
        double marginBottom,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        Func<DocxTableLayoutFrame> resolveFrame,
        Action finishPage,
        Func<bool> hasPageContent,
        Action markBoundaryContent,
        CancellationToken cancellationToken,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null)
    {
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows = table.Rows
            .Select((row, rowIndex) => (row, rowIndex))
            .TakeWhile(entry => entry.row.IsHeader)
            .Select(entry => (entry.row, entry.rowIndex))
            .ToArray();
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxTableLayoutFrame frame = resolveFrame();
            DocxTableRow row = table.Rows[rowIndex];
            IReadOnlyList<double> rowHeights = frame.RowHeights;
            double rowHeight = rowHeights[rowIndex];
            double remainingPageHeight = Math.Max(0d, cursorY - marginBottom);
            if (!row.CantSplit &&
                TryResolveExplicitTableCellPageBreakBoundaries(row, frame.EffectiveColumns, frame.Scale, rowHeight, textMeasurer, defaultTabStopPoints, getPageNumber(), pageCount: null, out IReadOnlyList<double> explicitBreakBoundaries, paragraphSpacingScale: paragraphSpacingScale))
            {
                double explicitBreakFragmentHeight = explicitBreakBoundaries[0];
                if (explicitBreakFragmentHeight > remainingPageHeight && hasPageContent())
                {
                    finishPage();
                    if (!row.IsHeader)
                    {
                        frame = resolveFrame();
                        rowHeights = frame.RowHeights;
                        AddRepeatedTableHeaderRows(table, frame.Context, rowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, frame.TableX, paragraphSpacingScale, cellMemo);
                        markBoundaryContent();
                    }
                }

                AddSplitTableRowLayout(table, row, rowIndex, headerRows, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, resolveFrame, explicitBreakBoundaries, "CellPageBreak", finishPage, paragraphSpacingScale, 0d, cellMemo);
                markBoundaryContent();
                continue;
            }

            if (!row.CantSplit &&
                rowHeight > remainingPageHeight &&
                remainingPageHeight > 0.001d &&
                CanSplitTableRowAtPageBoundary(row, frame.EffectiveColumns, frame.Scale, rowHeight, remainingPageHeight, out double splitPitch))
            {
                // RV06 pagination probe (edge-page-auto, Word 16.0): split fragments keep
                // whole lines that fit (floor capacity); the boundary is floored to whole
                // line pitches so a fractional remainder never squeezes an extra line in.
                double firstFragmentHeight = FloorTableRowFragmentHeightToPitch(remainingPageHeight, splitPitch);
                AddSplitTableRowLayout(table, row, rowIndex, headerRows, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, resolveFrame, firstFragmentHeight, "PageBoundary", finishPage, paragraphSpacingScale, splitPitch, cellMemo);
                markBoundaryContent();
                continue;
            }

            if (cursorY - rowHeight < marginBottom && hasPageContent())
            {
                finishPage();
                if (!row.IsHeader)
                {
                    frame = resolveFrame();
                    rowHeights = frame.RowHeights;
                    AddRepeatedTableHeaderRows(table, frame.Context, rowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, frame.TableX, paragraphSpacingScale, cellMemo);
                    markBoundaryContent();
                }
            }

            frame = resolveFrame();
            rowHeights = frame.RowHeights;
            AddTableRowLayout(table, frame.Context, row, rowIndex, rowHeights, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, frame.TableX, paragraphSpacingScale, cellMemo);
            markBoundaryContent();
        }

        // RV06 pagination probe (edge-page-auto, Word 16.0): split fragments pack whole
        // lines by floor capacity. The pitch is the median positive baseline gap across
        // the row's laid-out lines (robust to images shifting individual lines); unknown
        // when fewer than two lines laid out, which never needs flooring anyway.
        static double ResolveTableRowSplitPitch(List<double> baselines)
        {
            baselines.Sort();
            var gaps = new List<double>(baselines.Count);
            for (int gapIndex = 1; gapIndex < baselines.Count; gapIndex++)
            {
                double gap = baselines[gapIndex] - baselines[gapIndex - 1];
                if (gap > 0.001d)
                {
                    gaps.Add(gap);
                }
            }

            if (gaps.Count == 0)
            {
                return 0d;
            }

            gaps.Sort();
            return gaps[gaps.Count / 2];
        }


        bool CanSplitTableRowAtPageBoundary(DocxTableRow row, IReadOnlyList<double> effectiveColumns, double scale, double rowHeight, double firstFragmentHeight, out double splitPitch)
        {
            splitPitch = 0d;
            if (textMeasurer is null)
            {
                return false;
            }

            double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
            double rowTopPadding = ResolveTableRowTopPadding(row, paragraphSpacingScale);
            var laidOutCells = new List<(IReadOnlyList<DocxParagraph> Paragraphs, IReadOnlyList<DocxTextLineLayout> TextLines)>();
            var splitLineBaselines = new List<double>();
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                DocxTableCell cell = row.Cells[cellIndex];
                if (IsVerticalMergeContinuation(cell))
                {
                    continue;
                }

                IReadOnlyList<DocxTextLineLayout> textLines = LayoutTableCellTextLines(cell, 0d, 0d, cellWidths[cellIndex], rowHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, null, null, paragraphSpacingScale: paragraphSpacingScale, cellMemo: cellMemo).Lines;
                laidOutCells.Add((GetParagraphsFromBodyElements(GetTableCellLayoutBodyElements(cell)), textLines));
                splitLineBaselines.AddRange(textLines.Select(line => line.BaselineY));
            }

            // RV06 pagination probe (edge-page-ex48-r550, Word 16.0): test the floored
            // capacity the execution packs, not the raw remainder. A boundary landing
            // below the last baseline but above the row bottom still keeps whole lines.
            splitPitch = ResolveTableRowSplitPitch(splitLineBaselines);
            double flooredFirstFragmentHeight = FloorTableRowFragmentHeightToPitch(firstFragmentHeight, splitPitch);
            double fragmentBottomY = rowHeight - flooredFirstFragmentHeight;

            foreach ((IReadOnlyList<DocxParagraph> paragraphs, IReadOnlyList<DocxTextLineLayout> textLines) in laidOutCells)
            {
                bool HasTableCellKeepRuleBoundaryViolation()
                {
                    if (textLines.Count == 0)
                    {
                        return false;
                    }
            
                    foreach (IGrouping<int?, DocxTextLineLayout> group in textLines.GroupBy(line => line.SourceParagraphIndex))
                    {
                        if (group.Key is not { } paragraphIndex ||
                            paragraphIndex < 0 ||
                            paragraphIndex >= paragraphs.Count)
                        {
                            continue;
                        }
            
                        DocxParagraphKeepRules keepRules = paragraphs[paragraphIndex].EffectiveProperties.KeepRules;
                        int firstFragmentLineCount = group.Count(line => line.BaselineY >= fragmentBottomY);
                        int continuationLineCount = group.Count(line => line.BaselineY < fragmentBottomY);
                        bool splitsParagraph = firstFragmentLineCount != 0 && continuationLineCount != 0;
                        if (splitsParagraph &&
                            (keepRules.KeepNext == true ||
                                keepRules.KeepLines == true ||
                                (keepRules.WidowControl != false &&
                                    (firstFragmentLineCount == 1 || continuationLineCount == 1))))
                        {
                            return true;
                        }
            
                        bool IsNextTableCellParagraphInContinuation()
                        {
                            int nextParagraphIndex = paragraphIndex + 1;
                            return textLines.Any(line => line.SourceParagraphIndex == nextParagraphIndex) &&
                                textLines
                                    .Where(line => line.SourceParagraphIndex == nextParagraphIndex)
                                    .All(line => line.BaselineY < fragmentBottomY);
                        }

                        if (keepRules.KeepNext == true &&
                            firstFragmentLineCount != 0 &&
                            continuationLineCount == 0 &&
                            IsNextTableCellParagraphInContinuation())
                        {
                            return true;
                        }
                    }
            
                    return false;
                }

                if (HasTableCellKeepRuleBoundaryViolation())
                {
                    return false;
                }

                bool hasLineInFirstFragment = textLines.Any(line => flooredFirstFragmentHeight >= line.LineHeight && line.BaselineY >= fragmentBottomY);
                bool hasLineInContinuation = textLines.Any(line => line.BaselineY < fragmentBottomY);
                if (hasLineInFirstFragment && hasLineInContinuation)
                {
                    return true;
                }
            }

            return false;
        }
    }

    private static DocxTableLayoutFrame CreateTableLayoutFrame(
        DocxTable table,
        int tableIndex,
        int sourceBlockIndex,
        double x,
        double availableWidth,
        double pageContentHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        CancellationToken cancellationToken,
        int? pageNumber,
        int? pageCount,
        double paragraphSpacingScale)
    {
        DocxResolvedTableGrid grid = ResolveTableGrid(table, x, availableWidth, paragraphSpacingScale, textMeasurer, defaultTabStopPoints, pageNumber, pageCount);
        var tableContext = new DocxTableLayoutContext(
            tableIndex,
            sourceBlockIndex,
            table.Rows.Count,
            table.ColumnWidthsPoints.Count,
            table.ColumnWidthsPoints.Sum() * paragraphSpacingScale,
            table.HasExplicitGrid,
            grid.ResolvedColumnWidths,
            grid.TargetTableWidth,
            grid.TableX,
            table.PreferredWidthPoints,
            table.PreferredWidthValue,
            table.PreferredWidthType,
            table.IndentPoints,
            table.CellSpacingPoints,
            table.LayoutValue,
            table.Revisions);
        var rowHeights = new double[table.Rows.Count];
        double MeasureTableRowHeight(DocxTableRow row, IReadOnlyList<double> effectiveColumns, double scale, bool isLastRow)
        {
            double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
            double rowTopPadding = ResolveTableRowTopPadding(row, paragraphSpacingScale);
            double contentHeight = textMeasurer is null
                ? 0d
                : row.Cells
                    .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer, defaultTabStopPoints, rowTopPadding, pageNumber, pageCount, paragraphSpacingScale))
                    .DefaultIfEmpty(0d)
                    .Max();
            return ResolveTableRowHeight(row, contentHeight, paragraphSpacingScale, isLastRow);
        }

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rowHeights[rowIndex] = MeasureTableRowHeight(table.Rows[rowIndex], grid.EffectiveColumns, grid.Scale, isLastRow: rowIndex == table.Rows.Count - 1);
        }

        return new DocxTableLayoutFrame(tableContext, grid.EffectiveColumns, grid.Scale, rowHeights, pageContentHeight, grid.TableX);
    }

    private static DocxResolvedTableGrid ResolveTableGrid(DocxTable table, double x, double availableWidth, double fixedScale, IDocxTextMeasurer? textMeasurer, double defaultTabStopPoints, int? pageNumber, int? pageCount)
    {
        // W6-a1: fixed table geometry joins scaled space (fixed lengths times the layout
        // scale) so the uniform shift-composition maps frames like body text. Percent and
        // auto widths key off the already-scaled available width and need no change.
        // fixedScale reuses the layout spacing scale: identical in every mode (WC: s, else 1).
        double indentPoints = Math.Max(0d, table.IndentPoints ?? 0d) * fixedScale;
        double tableX = x + indentPoints;
        // Grid origin by regime (Office A/B w62-w72/m1-m14 probes; reader notes on
        // UseLegacyTableGrid/PinTextToMargin): compat aligns the outer border edge at
        // the origin (m13/ladder, indent or not per M14); modern pins text at an indent
        // origin (w66); pin-regime pins text at the margin via style margins (w72/m12).
        // Bare tables sit at the origin (w63). Table-level shift keeps columns aligned
        // across rows (ragged-border rows unprobed); explicit-0 indent counts as absent
        // (unprobed, assumed same).
        double gridX = tableX;
        if (table.Rows.FirstOrDefault()?.Cells.FirstOrDefault() is { } firstCell)
        {
            if (table.UseLegacyTableGrid)
            {
                DocxTableCellBorder? outerLeft = DocxTableBorderGeometry.Find(firstCell.Borders, "left") ?? DocxTableBorderGeometry.Find(firstCell.Borders, "start");
                gridX = tableX + DocxTableBorderGeometry.ResolveVisibleWidth(outerLeft) / 2d * fixedScale;
            }
            else if (indentPoints > 0d)
            {
                gridX = tableX - ResolveTableCellHorizontalEdgeInset(firstCell, "left", firstCell.Margins.LeftPoints, fixedScale);
            }
            else if ((firstCell.StyleMargins?.LeftPoints ?? 0d) > 0d)
            {
                gridX = tableX - (firstCell.StyleMargins?.LeftPoints ?? 0d) * fixedScale;
            }
        }
        double tableAvailableWidth = Math.Max(1d, availableWidth - indentPoints);
        IReadOnlyList<double> gridPoints = table.ColumnWidthsPoints.Select(width => width * fixedScale).ToArray();
        double gridTableWidth = gridPoints.Sum();
        double fallbackTableWidth = table.HasExplicitGrid && gridTableWidth > 0d ? gridTableWidth : tableAvailableWidth;
        double ResolveTargetTableWidth()
        {
            double? ResolvePreferredTableWidth()
            {
                if (table.PreferredWidthPoints is { } points)
                {
                    return points * fixedScale;
                }
        
                double ResolveOuterTableCellContentInset()
                {
                    DocxTableRow? firstRow = table.Rows.FirstOrDefault();
                    if (firstRow is null || firstRow.Cells.Count == 0)
                    {
                        return 0d;
                    }
            
                    DocxTableCell firstCell = firstRow.Cells[0];
                    DocxTableCell lastCell = firstRow.Cells[^1];
                    return ResolveTableCellHorizontalPadding(firstCell.Margins.LeftPoints, fixedScale) +
                        ResolveTableCellBorderContentInset(firstCell, "left", fixedScale) +
                        ResolveTableCellHorizontalPadding(lastCell.Margins.RightPoints, fixedScale) +
                        ResolveTableCellBorderContentInset(lastCell, "right", fixedScale);
                }
        
                if (table.PreferredWidthKind == DocxTableWidthKind.Percent &&
                    int.TryParse(table.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
                {
                    double normalPercentageWidth = tableAvailableWidth * fiftiethsPercent / 5000d;
                    double explicitGridWidth = table.HasExplicitGrid ? gridTableWidth : 0d;
                    double percentageBasis = explicitGridWidth > 0d && explicitGridWidth < normalPercentageWidth - 0.001d
                        ? tableAvailableWidth + ResolveOuterTableCellContentInset()
                        : tableAvailableWidth;
                    return Math.Max(0d, percentageBasis * fiftiethsPercent / 5000d);
                }
        
                return null;
            }

            double preferredWidth = ResolvePreferredTableWidth() ?? fallbackTableWidth;
            return table.PreferredWidthKind is DocxTableWidthKind.Dxa or DocxTableWidthKind.Percent
                ? Math.Max(1d, preferredWidth)
                : Math.Min(tableAvailableWidth, preferredWidth);
        }

        double targetTableWidth = ResolveTargetTableWidth();
        IReadOnlyList<double> effectiveColumns = GetEffectiveTableColumnWidths(targetTableWidth);
        double rawTableWidth = effectiveColumns.Sum();
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        return new DocxResolvedTableGrid(
            gridX,
            tableAvailableWidth,
            targetTableWidth,
            effectiveColumns,
            scale,
            effectiveColumns.Select(width => width * scale).ToArray());

        IReadOnlyList<double> GetEffectiveTableColumnWidths(double preferredTableWidth)
        {
            int columnCount = table.ColumnWidthsPoints.Count;
            if (columnCount == 0)
            {
                int inferredColumnCount = GetMaxGridColumnCount(table);
                return inferredColumnCount > 0
                    ? Enumerable.Repeat(preferredTableWidth / inferredColumnCount, inferredColumnCount).ToArray()
                    : gridPoints;
            }

            double?[] preferredWidths = new double?[columnCount];
            foreach (DocxTableRow row in table.Rows)
            {
                int gridColumnIndex = 0;
                foreach (DocxTableCell cell in row.Cells)
                {
                    int span = Math.Max(1, cell.GridSpan);
                    double? ResolvePreferredCellWidth()
                    {
                        if (cell.PreferredWidthPoints is { } points)
                        {
                            return points * fixedScale;
                        }
                
                        if (cell.PreferredWidthKind == DocxTableWidthKind.Percent &&
                            int.TryParse(cell.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
                        {
                            return Math.Max(0d, preferredTableWidth * fiftiethsPercent / 5000d);
                        }
                
                        return null;
                    }

                    double? cellPreferredWidth = ResolvePreferredCellWidth();
                    if (span == 1 &&
                        gridColumnIndex < columnCount &&
                        cellPreferredWidth is > 0d)
                    {
                        preferredWidths[gridColumnIndex] = cellPreferredWidth.Value;
                    }

                    gridColumnIndex += span;
                }

                if (preferredWidths.All(width => width is > 0d))
                {
                    return preferredWidths.Select(width => width ?? 0d).ToArray();
                }

                IReadOnlyList<double>? autoContentColumns = TryResolveAutoLayoutContentColumns(
                    table,
                    preferredTableWidth,
                    textMeasurer,
                    defaultTabStopPoints,
                    pageNumber,
                    pageCount,
                    fixedScale);
                if (autoContentColumns is not null)
                {
                    return autoContentColumns;
                }
            }

            if (!table.HasExplicitGrid)
            {
                int inferredColumnCount = columnCount == 0 ? GetMaxGridColumnCount(table) : columnCount;
                if (inferredColumnCount > 0)
                {
                    return Enumerable.Repeat(preferredTableWidth / inferredColumnCount, inferredColumnCount).ToArray();
                }
            }

            return gridPoints;
        }
    }

    // Office A/B (comment-table autofit probes: baseline/long/short/words plus comment
    // permutation renders, Word-COM rendered): tables without a fixed layout distribute
    // width by column content instead of the grid, so skewed content yields skewed columns.
    // Explicit per-column preferred widths, spans, and differentiated grids keep the
    // legacy path; content measurement mirrors cell layout inputs (single source with
    // LayoutTableCellTextLines for text, design extents for drawings).
    private static IReadOnlyList<double>? TryResolveAutoLayoutContentColumns(
        DocxTable table,
        double targetTableWidth,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        double fixedScale)
    {
        if (textMeasurer is null)
        {
            return null;
        }

        if (table.LayoutValue is not null &&
            !string.Equals(table.LayoutValue, "autofit", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        foreach (DocxTableRow row in table.Rows)
        {
            foreach (DocxTableCell cell in row.Cells)
            {
                if (Math.Max(1, cell.GridSpan) != 1)
                {
                    return null;
                }

                if (HasPreferredCellWidth(cell))
                {
                    return null;
                }
            }
        }

        // RV06: Office autofit ignores the grid (even a differentiated one) and sizes
        // by content, so grid uniformity never selects the legacy path here. Fixed layout
        // keeps grid widths through the LayoutValue guard above.

        int columnCount = table.ColumnWidthsPoints.Count != 0
            ? table.ColumnWidthsPoints.Count
            : GetMaxGridColumnCount(table);
        if (columnCount <= 0 || table.Rows.Count == 0)
        {
            return null;
        }

        double[] pads = new double[columnCount];
        double[]? bestRowTotals = null;
        double bestRowTotal = 0d;
        foreach (DocxTableRow row in table.Rows)
        {
            var rowTotals = new double[columnCount];
            for (int cellIndex = 0; cellIndex < row.Cells.Count && cellIndex < columnCount; cellIndex++)
            {
                DocxTableCell cell = row.Cells[cellIndex];
                double cellVisibleMax = MeasureTableCellMaxContentWidth(cell, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, fixedScale);
                double cellDeletedTotal = MeasureTableCellDeletedTextWidth(cell, textMeasurer);
                rowTotals[cellIndex] = cellVisibleMax + cellDeletedTotal;

                double cellPad = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints, fixedScale) +
                    ResolveTableCellHorizontalPadding(cell.Margins.RightPoints, fixedScale) +
                    ResolveTableCellBorderContentInset(cell, "left", fixedScale) +
                    ResolveTableCellBorderContentInset(cell, "right", fixedScale);
                if (cellPad > pads[cellIndex])
                {
                    pads[cellIndex] = cellPad;
                }
            }

            double rowTotal = rowTotals.Sum();
            if (bestRowTotals is null || rowTotal > bestRowTotal)
            {
                bestRowTotals = rowTotals;
                bestRowTotal = rowTotal;
            }
        }

        if (bestRowTotals is null || bestRowTotal <= 0d)
        {
            return null;
        }

        double totalPad = pads.Sum();
        double contentTarget = targetTableWidth - totalPad;
        if (contentTarget <= 0d)
        {
            return null;
        }

        return bestRowTotals
            .Select((rowTotal, index) => pads[index] + contentTarget * rowTotal / bestRowTotal)
            .ToArray();
    }

    private static bool HasPreferredCellWidth(DocxTableCell cell)
    {
        if (cell.PreferredWidthPoints is > 0d)
        {
            return true;
        }

        return cell.PreferredWidthKind == DocxTableWidthKind.Percent &&
            int.TryParse(cell.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent) &&
            fiftiethsPercent > 0;
    }

    private static double MeasureTableCellMaxContentWidth(
        DocxTableCell cell,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        double fixedScale)
    {
        double maxWidth = 0d;
        foreach (DocxBodyElement bodyElement in GetTableCellLayoutBodyElements(cell))
        {
            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs, pageNumber, pageCount);
            if (textSpans.Count != 0)
            {
                double fontSize = GetParagraphFontSize(paragraph);
                maxWidth = Math.Max(maxWidth, MeasureTextSpansForLayout(
                    textSpans,
                    fontSize,
                    textMeasurer,
                    ScaleTabStopPositions(paragraph.EffectiveProperties.TabStops, fixedScale),
                    defaultTabStopPoints * fixedScale,
                    pageNumber));
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                maxWidth = Math.Max(maxWidth, Math.Max(0d, image.WidthPoints) * fixedScale);
            }

            foreach (DocxInlineTextBox textBox in paragraph.InlineTextBoxes)
            {
                double boxWidth = ReadEmuPoints(textBox.ExtentCxValue) ?? 0d;
                maxWidth = Math.Max(maxWidth, Math.Max(0d, boxWidth) * fixedScale);
            }
        }

        return maxWidth;
    }
    // RV06: deleted/moveFrom text excluded from Runs by the markup view still sizes
    // columns in Office autofit, so the reader preserves it per paragraph and it is
    // measured here at the paragraph max font size.
    private static double MeasureTableCellDeletedTextWidth(DocxTableCell cell, IDocxTextMeasurer textMeasurer)
    {
        double total = 0d;
        foreach (DocxBodyElement bodyElement in GetTableCellLayoutBodyElements(cell))
        {
            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;

            if (paragraph.DeletedText.Length == 0)
            {
                continue;
            }

            DocxTextRun? contextRun = null;
            foreach (DocxTextRun candidate in paragraph.Runs)
            {
                if (candidate.Text.Length != 0)
                {
                    contextRun = candidate;
                    break;
                }
            }
            total += textMeasurer.MeasureText(contextRun ?? paragraph.Runs.FirstOrDefault(), paragraph.DeletedText, GetParagraphFontSize(paragraph));
        }

        return total;
    }

    internal const int MaxLayoutGridColumns = 1024;

    private static int GetMaxGridColumnCount(DocxTable table)
    {
        int max = 0;
        foreach (DocxTableRow row in table.Rows)
        {
            long total = 0;
            foreach (DocxTableCell cell in row.Cells)
            {
                total = checked(total + Math.Max(1, cell.GridSpan));
                if (total > MaxLayoutGridColumns)
                {
                    throw new OoxPdfLimitExceededException(
                        "DOCX table layout grid exceeds the maximum supported column count of " + MaxLayoutGridColumns + ".");
                }
            }

            max = Math.Max(max, (int)total);
        }

        return max;
    }

    // W6-a1: fixed table geometry joins scaled space; fixedScale reuses the layout
    // spacing scale (identical in every mode).
    private static double ResolveTableRowHeight(DocxTableRow row, double contentHeight, double fixedScale, bool isLastRow)
    {
        double MaxBottomBorderWidth()
        {
            return row.Cells
                .Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "bottom")))
                .DefaultIfEmpty(0d)
                .Max();
        }

        if (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) &&
            row.HeightPoints is { } exactHeight)
        {
            double exact = Math.Max(1d, exactHeight * fixedScale);
            if (isLastRow)
            {
                // Office A/B (w8 exact-row probe, Word-COM rendered plus PdfInspect
                // border rects): an exact-36 bordered row renders 36.5 tall with cell
                // content top-anchored exactly like atLeast, so the exact height pins
                // the content box and the last-row bottom border still hangs below it.
                exact += MaxBottomBorderWidth();
            }

            return exact;
        }

        double declaredHeight = string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase)
            ? 0d
            : row.HeightPoints ?? 0d;
        if (row.HeightPoints is not null &&
            !string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase))
        {
            declaredHeight += ResolveTableRowTopPadding(row, fixedScale);
        }

        double height = Math.Max(declaredHeight, contentHeight);
        double ResolveTableRowCollapsedHorizontalBorderAdvance()
        {
            double maxBottom = MaxBottomBorderWidth();
            if (maxBottom > 0d)
            {
                return maxBottom;
            }
    
            return row.Cells
                .Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "top")))
                .DefaultIfEmpty(0d)
                .Max();
        }

        height += ResolveTableRowCollapsedHorizontalBorderAdvance();
        if (isLastRow)
        {
            // Office A/B (w7 doc-start table probe, Word-COM rendered plus PdfInspect
            // border rects): Word hangs the last row bottom border BELOW cell content
            // (outside: 689.14 to 689.62 under content ending 689.62), while shared
            // interior borders straddle boundaries pitch-neutrally, so the table
            // terminus consumes one extra bottom width and following flow rides lower.
            // Unbordered last rows resolve zero width and stay bit-identical.
            height += MaxBottomBorderWidth();
        }

        return Math.Max(1d, height);
    }

    private static void AddTableRowLayout(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null)
    {
        double rowHeight = rowHeights[rowIndex];
        currentItems.Add(CreateTableRowLayout(
            table,
            tableContext,
            row,
            rowIndex,
            rowHeights,
            effectiveColumns,
            scale,
            textMeasurer,
            defaultTabStopPoints,
            getPageNumber,
            cursorY,
            rowHeight,
            logicalRowTopY: cursorY,
            FragmentIndex: 0,
            FragmentCount: 1,
            FragmentReason: "None",
            Story: null,
            pageCount: null,
            paragraphSpacingScale: paragraphSpacingScale,
            cellMemo: cellMemo));
        cursorY -= rowHeight;
    }

    private static void AddSplitTableRowLayout(
        DocxTable table,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        Func<DocxTableLayoutFrame> resolveFrame,
        double firstFragmentHeight,
        string fragmentReason,
        Action finishPage,
        double paragraphSpacingScale,
        double splitPitch,
        DocxTableCellTextLinesMemo? cellMemo = null)
    {
        AddSplitTableRowLayout(
            table,
            row,
            rowIndex,
            headerRows,
            textMeasurer,
            defaultTabStopPoints,
            getPageNumber,
            ref currentItems,
            ref cursorY,
            resolveFrame,
            [firstFragmentHeight],
            fragmentReason,
            finishPage,
            paragraphSpacingScale,
            splitPitch,
            cellMemo);
    }

    private static void AddSplitTableRowLayout(
        DocxTable table,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        Func<DocxTableLayoutFrame> resolveFrame,
        IReadOnlyList<double> fragmentBoundariesFromRowTop,
        string fragmentReason,
        Action finishPage,
        double paragraphSpacingScale,
        double splitPitch,
        DocxTableCellTextLinesMemo? cellMemo = null)
    {
        DocxTableLayoutFrame initialFrame = resolveFrame();
        IReadOnlyList<double> initialRowHeights = initialFrame.RowHeights;
        double rowHeight = initialRowHeights[rowIndex];
        double SumRepeatedTableHeaderRowsHeight()
        {
            return headerRows.Sum(entry => entry.RowIndex >= 0 && entry.RowIndex < initialRowHeights.Count ? initialRowHeights[entry.RowIndex] : 0d);
        }

        double continuationContentHeight = row.IsHeader
            ? initialFrame.PageContentHeight
            : Math.Max(1d, initialFrame.PageContentHeight - SumRepeatedTableHeaderRowsHeight());
        IReadOnlyList<double> fragmentHeights = ComputeTableRowFragmentHeights(rowHeight, fragmentBoundariesFromRowTop, continuationContentHeight, splitPitch);
        double consumedHeight = 0d;
        for (int fragmentIndex = 0; fragmentIndex < fragmentHeights.Count; fragmentIndex++)
        {
            DocxTableLayoutFrame frame = resolveFrame();
            double fragmentHeight = fragmentHeights[fragmentIndex];
            currentItems.Add(CreateTableRowLayout(
                table,
                frame.Context,
                row,
                rowIndex,
                frame.RowHeights,
                frame.EffectiveColumns,
                frame.Scale,
                textMeasurer,
                defaultTabStopPoints,
                getPageNumber,
                cursorY,
                fragmentHeight,
                logicalRowTopY: cursorY + consumedHeight,
                FragmentIndex: fragmentIndex,
                FragmentCount: fragmentHeights.Count,
                FragmentReason: fragmentReason,
                Story: null,
                pageCount: null,
                paragraphSpacingScale: paragraphSpacingScale,
                cellMemo: cellMemo));
            cursorY -= fragmentHeight;
            consumedHeight += fragmentHeight;

            if (fragmentIndex + 1 < fragmentHeights.Count)
            {
                finishPage();
                if (!row.IsHeader)
                {
                    frame = resolveFrame();
                    AddRepeatedTableHeaderRows(table, frame.Context, frame.RowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, frame.TableX, paragraphSpacingScale);
                }
            }
        }
    }

    private static void AddRepeatedTableHeaderRows(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null)
    {
        foreach ((DocxTableRow headerRow, int headerRowIndex) in headerRows)
        {
            AddTableRowLayout(table, tableContext, headerRow, headerRowIndex, rowHeights, effectiveColumns, scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, x, paragraphSpacingScale);
        }
    }

    private static bool TryResolveExplicitTableCellPageBreakBoundaries(
        DocxTableRow row,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        double rowHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int? pageNumber,
        int? pageCount,
        out IReadOnlyList<double> breakBoundariesFromRowTop,
        double paragraphSpacingScale)
    {
        breakBoundariesFromRowTop = [];
        if (textMeasurer is null || rowHeight <= 1.001d)
        {
            return false;
        }

        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row, paragraphSpacingScale);
        var breakHeights = new List<double>();
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCell cell = row.Cells[cellIndex];
            if (IsVerticalMergeContinuation(cell) ||
                !TryMeasureTableCellHeightBeforePageBreak(cell, cellWidths[cellIndex], textMeasurer, defaultTabStopPoints, rowTopPadding, pageNumber, pageCount, out double heightBeforeBreak, paragraphSpacingScale))
            {
                continue;
            }

            double breakHeight = Math.Min(rowHeight - 1d, Math.Max(1d, heightBeforeBreak));
            if (breakHeight < rowHeight - 0.001d)
            {
                breakHeights.Add(breakHeight);
            }
        }

        if (breakHeights.Count == 0)
        {
            return false;
        }

        breakBoundariesFromRowTop = breakHeights
            .Order()
            .Aggregate(new List<double>(), (boundaries, breakHeight) =>
            {
                if (boundaries.Count == 0 || Math.Abs(boundaries[^1] - breakHeight) > 0.001d)
                {
                    boundaries.Add(breakHeight);
                }

                return boundaries;
            });
        return breakBoundariesFromRowTop.Count != 0;
    }

    private static bool TryMeasureTableCellHeightBeforePageBreak(
        DocxTableCell cell,
        double cellWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        double rowTopPadding,
        int? pageNumber,
        int? pageCount,
        out double heightBeforeBreak,
        double paragraphSpacingScale)
    {
        heightBeforeBreak = 0d;
        var context = new DocxTableCellLayoutContext(
            TextMeasurer: textMeasurer,
            DefaultTabStopPoints: defaultTabStopPoints,
            ParagraphSpacingScale: paragraphSpacingScale,
            PageNumber: pageNumber,
            PageCount: pageCount,
            CellMemo: null);
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return false;
        }

        int pageBreakIndex = -1;
        for (int index = 0; index < bodyElements.Count; index++)
        {
            // RV06 cellbreak probe: explicit run breaks inside table cells do not turn
            // pages in Office; only other break provenances may fragment rows here.
            if (bodyElements[index] is DocxPageBreakElement cellPageBreak &&
                cellPageBreak.SourceKind != DocxBreakSourceKind.RunBreak)
            {
                pageBreakIndex = index;
                break;
            }
        }

        if (pageBreakIndex <= 0 || !bodyElements.Skip(pageBreakIndex + 1).Any(IsRenderableTableCellBodyElement))
        {
            return false;
        }

        double paddingLeft = ResolveTableCellHorizontalEdgeInset(cell, "left", cell.Margins.LeftPoints, paragraphSpacingScale);
        double paddingRight = ResolveTableCellHorizontalEdgeInset(cell, "right", cell.Margins.RightPoints, paragraphSpacingScale);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        heightBeforeBreak = rowTopPadding;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        for (int index = 0; index < pageBreakIndex; index++)
        {
            DocxBodyElement bodyElement = bodyElements[index];
            if (bodyElement is DocxTableElement tableElement)
            {
                heightBeforeBreak += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                heightBeforeBreak += MeasureNestedTableHeight(tableElement.Table, textWidth, context);
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
            heightBeforeBreak += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            heightBeforeBreak += MeasureTableCellParagraphContentHeight(cell, paragraph, textWidth, context);
            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        heightBeforeBreak += pendingSpacingAfter;
        return heightBeforeBreak > rowTopPadding + 0.001d;
    }

    private static bool IsRenderableTableCellBodyElement(DocxBodyElement bodyElement)
    {
        return bodyElement switch
        {
            DocxParagraphElement paragraphElement => paragraphElement.Paragraph.Runs.Count != 0 ||
                paragraphElement.Paragraph.Images.Count != 0 ||
                paragraphElement.Paragraph.ListLabel is not null,
            DocxTableElement => true,
            _ => false
        };
    }

    private static IReadOnlyList<double> ComputeTableRowFragmentHeights(double rowHeight, double firstFragmentHeight, double pageContentHeight)
    {
        return ComputeTableRowFragmentHeights(rowHeight, [firstFragmentHeight], pageContentHeight);
    }

    internal const int MaxTableRowFragments = 1000;

    // RV06 pagination probe (edge-page-auto, Word 16.0): split fragments pack whole
    // lines by floor capacity. Flooring a fragment height to whole line pitches keeps
    // exact fits bit-identical while fractional remainders stop squeezing an extra
    // line in. Non-positive pitches (unknown) leave heights untouched.
    private static double FloorTableRowFragmentHeightToPitch(double height, double pitch)
    {
        if (!(pitch > 0d) || !double.IsFinite(height) || !double.IsFinite(pitch))
        {
            return height;
        }

        double floored = Math.Floor(height / pitch) * pitch;
        return floored >= pitch ? floored : height;
    }

    private static IReadOnlyList<double> ComputeTableRowFragmentHeights(double rowHeight, IReadOnlyList<double> fragmentBoundariesFromRowTop, double pageContentHeight, double splitPitch = 0d)
    {
        // validate geometry before fragment expansion and require numeric
        // progress so extreme authored heights cannot append unbounded fragments.
        if (!double.IsFinite(rowHeight) || !double.IsFinite(pageContentHeight))
        {
            throw new OoxPdfLimitExceededException("DOCX table row geometry is not finite.");
        }

        if (rowHeight <= 0d || rowHeight > 100000d)
        {
            throw new OoxPdfLimitExceededException("DOCX table row height exceeds the maximum supported height.");
        }

        var fragments = new List<double>();
        double consumedHeight = 0d;
        double fullPageHeight = Math.Max(1d, pageContentHeight);
        if (!double.IsFinite(fullPageHeight) || fullPageHeight < 1d)
        {
            throw new OoxPdfLimitExceededException("DOCX table continuation height is invalid.");
        }

        foreach (double boundary in fragmentBoundariesFromRowTop.Order())
        {
            if (!double.IsFinite(boundary))
            {
                throw new OoxPdfLimitExceededException("DOCX table fragment boundary is not finite.");
            }

            double clampedBoundary = Math.Min(rowHeight, Math.Max(0d, boundary));
            if (clampedBoundary <= consumedHeight + 0.001d)
            {
                continue;
            }

            AddTableRowFragmentSegmentHeights(fragments, clampedBoundary - consumedHeight, fullPageHeight, splitPitch);
            consumedHeight = clampedBoundary;
        }

        AddTableRowFragmentSegmentHeights(fragments, rowHeight - consumedHeight, fullPageHeight, splitPitch);
        return fragments.Count == 0 ? [Math.Max(1d, rowHeight)] : fragments;
    }

    private static void AddTableRowFragmentSegmentHeights(List<double> fragments, double segmentHeight, double fullPageHeight, double splitPitch = 0d)
    {
        if (!double.IsFinite(segmentHeight) || !double.IsFinite(fullPageHeight))
        {
            throw new OoxPdfLimitExceededException("DOCX table fragment geometry is not finite.");
        }

        if (fullPageHeight < 1d)
        {
            throw new OoxPdfLimitExceededException("DOCX table continuation height is invalid.");
        }

        // Require numeric progress: extreme doubles can make subtraction stop
        // changing the value and loop without bound (M06).
        double remainingHeight = segmentHeight;
        double pageChunkHeight = FloorTableRowFragmentHeightToPitch(fullPageHeight, splitPitch);
        double lastRemaining = double.PositiveInfinity;
        while (remainingHeight > pageChunkHeight + 0.001d)
        {
            if (fragments.Count >= MaxTableRowFragments)
            {
                throw new OoxPdfLimitExceededException(
                    "DOCX table row exceeds the maximum supported fragment count of " + MaxTableRowFragments + ".");
            }

            if (!(remainingHeight < lastRemaining))
            {
                throw new OoxPdfLimitExceededException("DOCX table row fragmentation cannot progress.");
            }

            lastRemaining = remainingHeight;
            fragments.Add(pageChunkHeight);
            // conversion-wide cumulative charge per constructed fragment.
            OoxConversionBudget.Current?.ChargeTableFragments(1);
            double next = remainingHeight - pageChunkHeight;
            if (!(next < remainingHeight))
            {
                throw new OoxPdfLimitExceededException("DOCX table row fragmentation cannot progress.");
            }

            remainingHeight = next;
        }

        if (remainingHeight > 0.001d)
        {
            if (fragments.Count >= MaxTableRowFragments)
            {
                throw new OoxPdfLimitExceededException(
                    "DOCX table row exceeds the maximum supported fragment count of " + MaxTableRowFragments + ".");
            }

            fragments.Add(remainingHeight);
            // conversion-wide cumulative charge per constructed fragment.
            OoxConversionBudget.Current?.ChargeTableFragments(1);
        }
    }

    private static DocxTableRowLayout CreateTableRowLayout(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageNumber,
        double cursorY,
        double rowHeight,
        double logicalRowTopY,
        int FragmentIndex,
        int FragmentCount,
        string FragmentReason,
        DocxStoryId? Story,
        int? pageCount,
        double paragraphSpacingScale,
        DocxTableCellTextLinesMemo? cellMemo = null)
    {
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row, paragraphSpacingScale);
        double fullRowHeight = rowHeights[rowIndex];
        double fragmentOffsetFromRowTop = logicalRowTopY - cursorY;
        double cellX = tableContext.TableX;
        double cellY = cursorY - rowHeight;
        double fullCellY = logicalRowTopY - fullRowHeight;
        var cells = new List<DocxTableCellLayout>(row.Cells.Count);
        int gridColumnIndex = 0;
        for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
        {
            double cellWidth = cellWidths[columnIndex];
            DocxTableCell cell = row.Cells[columnIndex];
            bool isVerticalMergeContinuation = IsVerticalMergeContinuation(cell);
            DocxVerticalMergeOwner? verticalMergeOwner = isVerticalMergeContinuation
                ? FindVerticalMergeRestartOwner(table, rowIndex, gridColumnIndex)
                : null;
            DocxTableCell? verticalMergeOwnerCell = verticalMergeOwner?.Cell;
            DocxTableCellVisualOwnership visualOwnership = isVerticalMergeContinuation
                ? verticalMergeOwnerCell is null
                    ? DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                    : DocxTableCellVisualOwnership.VerticalMergeOwner
                : DocxTableCellVisualOwnership.OwnCell;
            double visualHeight = rowHeight;
            double visualY = cellY;
            double fullVisualHeight = fullRowHeight;
            double fullVisualY = fullCellY;
            if (IsVerticalMergeRestart(cell))
            {
                fullVisualHeight = GetVerticalMergeSpanHeight(table, rowIndex, gridColumnIndex, rowHeights);
                fullVisualY = cursorY - fullVisualHeight;
                if (FragmentCount == 1)
                {
                    visualHeight = fullVisualHeight;
                    visualY = fullVisualY;
                }
            }

            bool useCellPageBreakBoundaryPartition = FragmentReason == "CellPageBreak" && textMeasurer is not null;
            int cellPageBreakLowerParagraphBoundaryIndex = 0;
            int? cellPageBreakUpperParagraphBoundaryIndex = null;
            int cellPageBreakLowerNestedTableBoundaryIndex = 0;
            int? cellPageBreakUpperNestedTableBoundaryIndex = null;
            bool cellPageBreakLowerBoundaryInsideNestedTable = false;
            bool cellPageBreakUpperBoundaryInsideNestedTable = false;
            int currentPageNumber = getPageNumber();
            int currentPageIndex = Math.Max(0, currentPageNumber - 1);
            if (useCellPageBreakBoundaryPartition && textMeasurer is not null)
            {
                if (fragmentOffsetFromRowTop > 0.001d &&
                    TryResolveTableCellParagraphBoundaryIndex(cell, cellWidth, rowTopPadding, fragmentOffsetFromRowTop, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, out int lowerParagraphBoundaryIndex, paragraphSpacingScale))
                {
                    cellPageBreakLowerParagraphBoundaryIndex = lowerParagraphBoundaryIndex;
                }

                if (fragmentOffsetFromRowTop > 0.001d &&
                    TryResolveTableCellNestedTableBoundary(cell, cellWidth, rowTopPadding, fragmentOffsetFromRowTop, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, out DocxNestedTableBoundary lowerNestedTableBoundary, paragraphSpacingScale))
                {
                    cellPageBreakLowerNestedTableBoundaryIndex = lowerNestedTableBoundary.BoundaryIndex;
                    cellPageBreakLowerBoundaryInsideNestedTable = lowerNestedTableBoundary.IsInsideNestedTable;
                }

                double fragmentEndFromRowTop = fragmentOffsetFromRowTop + rowHeight;
                if (fragmentEndFromRowTop < fullRowHeight - 0.001d &&
                    TryResolveTableCellParagraphBoundaryIndex(cell, cellWidth, rowTopPadding, fragmentEndFromRowTop, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, out int upperParagraphBoundaryIndex, paragraphSpacingScale))
                {
                    cellPageBreakUpperParagraphBoundaryIndex = upperParagraphBoundaryIndex;
                }

                if (fragmentEndFromRowTop < fullRowHeight - 0.001d &&
                    TryResolveTableCellNestedTableBoundary(cell, cellWidth, rowTopPadding, fragmentEndFromRowTop, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, out DocxNestedTableBoundary upperNestedTableBoundary, paragraphSpacingScale))
                {
                    cellPageBreakUpperNestedTableBoundaryIndex = upperNestedTableBoundary.BoundaryIndex + (upperNestedTableBoundary.IsInsideNestedTable ? 1 : 0);
                    cellPageBreakUpperBoundaryInsideNestedTable = upperNestedTableBoundary.IsInsideNestedTable;
                }
            }

            bool cellPageBreakBoundaryInsideNestedTable =
                cellPageBreakLowerBoundaryInsideNestedTable ||
                cellPageBreakUpperBoundaryInsideNestedTable;
            bool cellPageBreakAlignsWithNestedTableBlock =
                useCellPageBreakBoundaryPartition &&
                !cellPageBreakBoundaryInsideNestedTable;

            DocxTableCell contentCell = visualOwnership == DocxTableCellVisualOwnership.VerticalMergeOwner && verticalMergeOwnerCell is not null
                ? verticalMergeOwnerCell
                : cell;
            double contentY = isVerticalMergeContinuation ? visualY : fullVisualY;
            double contentHeight = isVerticalMergeContinuation ? visualHeight : fullVisualHeight;
            double contentPaddingLeft = ResolveTableCellHorizontalEdgeInset(contentCell, "left", contentCell.Margins.LeftPoints, paragraphSpacingScale);
            double contentPaddingTop = rowTopPadding;
            double contentPaddingRight = ResolveTableCellHorizontalEdgeInset(contentCell, "right", contentCell.Margins.RightPoints, paragraphSpacingScale);
            double contentPaddingBottom = ResolveTableCellVerticalPadding(contentCell.Margins.BottomPoints, paragraphSpacingScale);
            (IReadOnlyList<DocxTextLineLayout> cellTextLines, IReadOnlyList<DocxInlineImageLayout> cellPlacedImages) = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? (Array.Empty<DocxTextLineLayout>(), Array.Empty<DocxInlineImageLayout>())
                : LayoutTableCellTextLines(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, paragraphSpacingScale, cellMemo, currentPageIndex);
            IReadOnlyList<DocxTextLineLayout> textLines = cellTextLines
                    .Where(line => IsTextLineOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, cell, line, FragmentIndex, FragmentCount))
                    .Where(line => IsTextLineVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, line, visualY, visualHeight, FragmentIndex, FragmentCount))
                    .ToArray();
            (IReadOnlyList<DocxInlineImageLayout> cellInlineImages, IReadOnlyList<DocxInlineTextBoxLayout> cellInlineTextBoxes) = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? (Array.Empty<DocxInlineImageLayout>(), Array.Empty<DocxInlineTextBoxLayout>())
                : LayoutTableCellInlineImages(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageIndex, currentPageNumber, pageCount, paragraphSpacingScale);
            IReadOnlyList<DocxInlineImageLayout> inlineImages = cellInlineImages
                    .Concat(cellPlacedImages)
                    .Where(image => IsInlineImageOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, cell, image, FragmentIndex, FragmentCount))
                    .Where(image => IsInlineImageVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, image, visualY, visualHeight, FragmentIndex, FragmentCount))
                    .ToArray();
            IReadOnlyList<DocxInlineTextBoxLayout> inlineTextBoxes = cellInlineTextBoxes
                    .Where(box => IsInlineTextBoxOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, cell, box, FragmentIndex, FragmentCount))
                    .Where(box => IsInlineTextBoxVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, box, visualY, visualHeight, FragmentIndex, FragmentCount))
                    .ToArray();
            IReadOnlyList<DocxTableRowLayout> nestedTableRows = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? []
                : LayoutTableCellNestedTables(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageIndex, currentPageNumber, pageCount, paragraphSpacingScale, cellMemo)
                    .Where(rowLayout => IsNestedTableRowOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerNestedTableBoundaryIndex, cellPageBreakUpperNestedTableBoundaryIndex, rowLayout, FragmentCount))
                    .Where(rowLayout => IsNestedTableRowVisibleInCellFragmentGeometry(cellPageBreakAlignsWithNestedTableBlock, rowLayout, visualY, visualHeight, FragmentIndex, FragmentCount))
                    .ToArray();
            cells.Add(new DocxTableCellLayout(
                cell,
                cellX,
                visualY,
                cellWidth,
                visualHeight,
                contentPaddingLeft,
                contentPaddingTop,
                contentPaddingRight,
                contentPaddingBottom,
                textLines,
                inlineImages,
                isVerticalMergeContinuation,
                verticalMergeOwnerCell,
                verticalMergeOwner,
                visualOwnership,
                nestedTableRows) with { InlineTextBoxes = inlineTextBoxes });
            cellX += cellWidth + (table.CellSpacingPoints ?? 0d) * paragraphSpacingScale;
            gridColumnIndex += Math.Max(1, cell.GridSpan);
        }

        return new DocxTableRowLayout(
            tableContext,
            rowIndex,
            FragmentIndex,
            FragmentCount,
            FragmentReason,
            fullRowHeight,
            fragmentOffsetFromRowTop,
            cells.ToArray(),
            cellY,
            rowHeight,
            row.HeightPoints,
            row.HeightValue,
            row.HeightRuleValue,
            row.IsHeader,
            row.HeaderValue,
            row.TablePropertyExceptionCellMargins is not null,
            row.CantSplit,
            row.CantSplitValue,
            RevisionCount: CountTableRowSourceRevisions(row),
            Revisions: CollectTableRowSourceRevisions(row),
            Story: Story);
    }
}
