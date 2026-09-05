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
        double paragraphSpacingScale)
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
                        AddRepeatedTableHeaderRows(table, frame.Context, rowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, frame.TableX, paragraphSpacingScale);
                        markBoundaryContent();
                    }
                }

                AddSplitTableRowLayout(table, row, rowIndex, headerRows, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, resolveFrame, explicitBreakBoundaries, "CellPageBreak", finishPage, paragraphSpacingScale);
                markBoundaryContent();
                continue;
            }

            if (!row.CantSplit &&
                rowHeight > remainingPageHeight &&
                remainingPageHeight > 0.001d &&
                CanSplitTableRowAtPageBoundary(row, frame.EffectiveColumns, frame.Scale, rowHeight, remainingPageHeight))
            {
                AddSplitTableRowLayout(table, row, rowIndex, headerRows, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, resolveFrame, remainingPageHeight, "PageBoundary", finishPage, paragraphSpacingScale);
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
                    AddRepeatedTableHeaderRows(table, frame.Context, rowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, frame.TableX, paragraphSpacingScale);
                    markBoundaryContent();
                }
            }

            frame = resolveFrame();
            rowHeights = frame.RowHeights;
            AddTableRowLayout(table, frame.Context, row, rowIndex, rowHeights, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageNumber, ref currentItems, ref cursorY, frame.TableX, paragraphSpacingScale);
            markBoundaryContent();
        }

        bool CanSplitTableRowAtPageBoundary(DocxTableRow row, IReadOnlyList<double> effectiveColumns, double scale, double rowHeight, double firstFragmentHeight)
        {
            if (textMeasurer is null)
            {
                return false;
            }

            double fragmentBottomY = rowHeight - firstFragmentHeight;
            double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
            double rowTopPadding = ResolveTableRowTopPadding(row);
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
            {
                DocxTableCell cell = row.Cells[cellIndex];
                if (IsVerticalMergeContinuation(cell))
                {
                    continue;
                }

                IReadOnlyList<DocxTextLineLayout> textLines = LayoutTableCellTextLines(cell, 0d, 0d, cellWidths[cellIndex], rowHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, null, null, paragraphSpacingScale: paragraphSpacingScale);
                bool HasTableCellKeepRuleBoundaryViolation()
                {
                    if (textLines.Count == 0)
                    {
                        return false;
                    }
            
                    IReadOnlyList<DocxParagraph> paragraphs = GetParagraphsFromBodyElements(GetTableCellLayoutBodyElements(cell));
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

                bool hasLineInFirstFragment = textLines.Any(line => firstFragmentHeight >= line.LineHeight && line.BaselineY >= fragmentBottomY);
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
        DocxResolvedTableGrid grid = ResolveTableGrid(table, x, availableWidth);
        var tableContext = new DocxTableLayoutContext(
            tableIndex,
            sourceBlockIndex,
            table.Rows.Count,
            table.ColumnWidthsPoints.Count,
            table.ColumnWidthsPoints.Sum(),
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
        double MeasureTableRowHeight(DocxTableRow row, IReadOnlyList<double> effectiveColumns, double scale)
        {
            double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
            double rowTopPadding = ResolveTableRowTopPadding(row);
            double contentHeight = textMeasurer is null
                ? 0d
                : row.Cells
                    .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer, defaultTabStopPoints, rowTopPadding, pageNumber, pageCount, paragraphSpacingScale))
                    .DefaultIfEmpty(0d)
                    .Max();
            return ResolveTableRowHeight(row, contentHeight);
        }

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rowHeights[rowIndex] = MeasureTableRowHeight(table.Rows[rowIndex], grid.EffectiveColumns, grid.Scale);
        }

        return new DocxTableLayoutFrame(tableContext, grid.EffectiveColumns, grid.Scale, rowHeights, pageContentHeight, grid.TableX);
    }

    private static DocxResolvedTableGrid ResolveTableGrid(DocxTable table, double x, double availableWidth)
    {
        double tableX = x + Math.Max(0d, table.IndentPoints ?? 0d);
        double tableAvailableWidth = Math.Max(1d, availableWidth - Math.Max(0d, table.IndentPoints ?? 0d));
        double gridTableWidth = table.ColumnWidthsPoints.Sum();
        double fallbackTableWidth = table.HasExplicitGrid && gridTableWidth > 0d ? gridTableWidth : tableAvailableWidth;
        double ResolveTargetTableWidth()
        {
            double? ResolvePreferredTableWidth()
            {
                if (table.PreferredWidthPoints is { } points)
                {
                    return points;
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
                    return ResolveTableCellHorizontalPadding(firstCell.Margins.LeftPoints) +
                        ResolveTableCellBorderContentInset(firstCell, "left") +
                        ResolveTableCellHorizontalPadding(lastCell.Margins.RightPoints) +
                        ResolveTableCellBorderContentInset(lastCell, "right");
                }
        
                if (table.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true &&
                    int.TryParse(table.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
                {
                    double normalPercentageWidth = tableAvailableWidth * fiftiethsPercent / 5000d;
                    double explicitGridWidth = table.HasExplicitGrid ? table.ColumnWidthsPoints.Sum() : 0d;
                    double percentageBasis = explicitGridWidth > 0d && explicitGridWidth < normalPercentageWidth - 0.001d
                        ? tableAvailableWidth + ResolveOuterTableCellContentInset()
                        : tableAvailableWidth;
                    return Math.Max(0d, percentageBasis * fiftiethsPercent / 5000d);
                }
        
                return null;
            }

            double preferredWidth = ResolvePreferredTableWidth() ?? fallbackTableWidth;
            return table.PreferredWidthType?.Equals("dxa", StringComparison.OrdinalIgnoreCase) == true ||
                table.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true
                ? Math.Max(1d, preferredWidth)
                : Math.Min(tableAvailableWidth, preferredWidth);
        }

        double targetTableWidth = ResolveTargetTableWidth();
        IReadOnlyList<double> effectiveColumns = GetEffectiveTableColumnWidths(targetTableWidth);
        double rawTableWidth = effectiveColumns.Sum();
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        return new DocxResolvedTableGrid(
            tableX,
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
                    : table.ColumnWidthsPoints;
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
                            return points;
                        }
                
                        if (cell.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true &&
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
            }

            if (!table.HasExplicitGrid)
            {
                int inferredColumnCount = columnCount == 0 ? GetMaxGridColumnCount(table) : columnCount;
                if (inferredColumnCount > 0)
                {
                    return Enumerable.Repeat(preferredTableWidth / inferredColumnCount, inferredColumnCount).ToArray();
                }
            }

            return table.ColumnWidthsPoints;
        }
    }

    private static int GetMaxGridColumnCount(DocxTable table)
    {
        return table.Rows
            .Select(row => row.Cells.Sum(cell => Math.Max(1, cell.GridSpan)))
            .DefaultIfEmpty(0)
            .Max();
    }

    private static double ResolveTableRowHeight(DocxTableRow row, double contentHeight)
    {
        if (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) &&
            row.HeightPoints is { } exactHeight)
        {
            return Math.Max(1d, exactHeight);
        }

        double declaredHeight = string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase)
            ? 0d
            : row.HeightPoints ?? 0d;
        if (row.HeightPoints is not null &&
            !string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase))
        {
            declaredHeight += ResolveTableRowTopPadding(row);
        }

        double height = Math.Max(declaredHeight, contentHeight);
        double ResolveTableRowCollapsedHorizontalBorderAdvance()
        {
            double maxBottom = row.Cells
                .Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "bottom")))
                .DefaultIfEmpty(0d)
                .Max();
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
        double paragraphSpacingScale)
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
            StoryKind: null,
            StoryVariantType: null,
            pageCount: null,
            paragraphSpacingScale: paragraphSpacingScale));
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
        double paragraphSpacingScale)
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
            paragraphSpacingScale);
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
        double paragraphSpacingScale)
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
        IReadOnlyList<double> fragmentHeights = ComputeTableRowFragmentHeights(rowHeight, fragmentBoundariesFromRowTop, continuationContentHeight);
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
                StoryKind: null,
                StoryVariantType: null,
                pageCount: null,
                paragraphSpacingScale: paragraphSpacingScale));
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
        double paragraphSpacingScale)
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
        double rowTopPadding = ResolveTableRowTopPadding(row);
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
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return false;
        }

        int pageBreakIndex = -1;
        for (int index = 0; index < bodyElements.Count; index++)
        {
            if (bodyElements[index] is DocxPageBreakElement)
            {
                pageBreakIndex = index;
                break;
            }
        }

        if (pageBreakIndex <= 0 || !bodyElements.Skip(pageBreakIndex + 1).Any(IsRenderableTableCellBodyElement))
        {
            return false;
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
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
                heightBeforeBreak += MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount, paragraphSpacingScale);
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
            heightBeforeBreak += MeasureTableCellParagraphContentHeight(cell, paragraph, textWidth, textMeasurer, defaultTabStopPoints, pageNumber, pageCount);
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

    private static IReadOnlyList<double> ComputeTableRowFragmentHeights(double rowHeight, IReadOnlyList<double> fragmentBoundariesFromRowTop, double pageContentHeight)
    {
        var fragments = new List<double>();
        double consumedHeight = 0d;
        double fullPageHeight = Math.Max(1d, pageContentHeight);
        foreach (double boundary in fragmentBoundariesFromRowTop.Order())
        {
            double clampedBoundary = Math.Min(rowHeight, Math.Max(0d, boundary));
            if (clampedBoundary <= consumedHeight + 0.001d)
            {
                continue;
            }

            AddTableRowFragmentSegmentHeights(fragments, clampedBoundary - consumedHeight, fullPageHeight);
            consumedHeight = clampedBoundary;
        }

        AddTableRowFragmentSegmentHeights(fragments, rowHeight - consumedHeight, fullPageHeight);
        return fragments.Count == 0 ? [Math.Max(1d, rowHeight)] : fragments;
    }

    private static void AddTableRowFragmentSegmentHeights(List<double> fragments, double segmentHeight, double fullPageHeight)
    {
        double remainingHeight = segmentHeight;
        while (remainingHeight > fullPageHeight + 0.001d)
        {
            fragments.Add(fullPageHeight);
            remainingHeight -= fullPageHeight;
        }

        if (remainingHeight > 0.001d)
        {
            fragments.Add(remainingHeight);
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
        string? StoryKind,
        string? StoryVariantType,
        int? pageCount,
        double paragraphSpacingScale)
    {
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
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
            double contentPaddingLeft = ResolveTableCellHorizontalPadding(contentCell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(contentCell, "left");
            double contentPaddingTop = rowTopPadding;
            double contentPaddingRight = ResolveTableCellHorizontalPadding(contentCell.Margins.RightPoints) + ResolveTableCellBorderContentInset(contentCell, "right");
            double contentPaddingBottom = ResolveTableCellVerticalPadding(contentCell.Margins.BottomPoints);
            IReadOnlyList<DocxTextLineLayout> textLines = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? []
                : LayoutTableCellTextLines(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageNumber, pageCount, paragraphSpacingScale)
                    .Where(line => IsTextLineOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, cell, line, FragmentIndex, FragmentCount))
                    .Where(line => IsTextLineVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, line, visualY, visualHeight, FragmentIndex, FragmentCount))
                    .ToArray();
            IReadOnlyList<DocxInlineImageLayout> inlineImages = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? []
                : LayoutTableCellInlineImages(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageIndex, currentPageNumber, pageCount, paragraphSpacingScale)
                    .Where(image => IsInlineImageOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, cell, image, FragmentIndex, FragmentCount))
                    .Where(image => IsInlineImageVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, image, visualY, visualHeight, FragmentIndex, FragmentCount))
                    .ToArray();
            IReadOnlyList<DocxTableRowLayout> nestedTableRows = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? []
                : LayoutTableCellNestedTables(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, currentPageIndex, currentPageNumber, pageCount, paragraphSpacingScale)
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
                nestedTableRows));
            cellX += cellWidth + (table.CellSpacingPoints ?? 0d);
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
            StoryKind: StoryKind,
            StoryVariantType: StoryVariantType);
    }
}
