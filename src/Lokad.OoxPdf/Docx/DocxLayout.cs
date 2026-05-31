using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxLayout(IReadOnlyList<DocxLayoutPage> Pages);

internal sealed record DocxLayoutSnapshot(IReadOnlyList<DocxLayoutPageSnapshot> Pages)
{
    public static DocxLayoutSnapshot FromLayout(DocxLayout layout)
    {
        return new DocxLayoutSnapshot(layout.Pages.Select(ToSnapshot).ToArray());
    }

    private static DocxLayoutPageSnapshot ToSnapshot(DocxLayoutPage page)
    {
        IReadOnlyList<DocxLayoutItemSnapshot> items = page.Items.Select(ToSnapshot).ToArray();
        double verticalTop = items.Count == 0 ? 0d : items.Max(item => item.Y + item.Height);
        double verticalBottom = items.Count == 0 ? 0d : items.Min(item => item.Y);
        return new DocxLayoutPageSnapshot(
            page.Width,
            page.Height,
            items.Count,
            items.Count(item => item.Kind == "TextLine"),
            items.Count(item => item.Kind == "InlineImage"),
            items.Count(item => item.Kind == "TableRow"),
            Math.Max(0d, verticalTop - verticalBottom),
            items.Where(item => item.Kind == "TextLine").Sum(item => item.Height),
            items.Where(item => item.Kind == "InlineImage").Sum(item => item.Height),
            items.Where(item => item.Kind == "TableRow").Sum(item => item.Height),
            items);
    }

    private static DocxLayoutItemSnapshot ToSnapshot(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => new DocxLayoutItemSnapshot(
                "TextLine",
                text.X,
                text.BaselineY,
                text.Width,
                text.FontSize,
                TextLength: text.Text.Length,
                CellCount: 0),
            DocxInlineImageLayout image => new DocxLayoutItemSnapshot(
                "InlineImage",
                image.X,
                image.Y,
                image.Width,
                image.Height,
                TextLength: 0,
                CellCount: 0),
            DocxTableRowLayout row => new DocxLayoutItemSnapshot(
                "TableRow",
                row.Cells.Count == 0 ? 0d : row.Cells.Min(cell => cell.X),
                row.Y,
                row.Cells.Sum(cell => cell.Width),
                row.Height,
                TextLength: row.Cells.Sum(cell => cell.TextLines.Sum(line => line.Text.Length)),
                CellCount: row.Cells.Count),
            _ => new DocxLayoutItemSnapshot("Unknown", 0d, 0d, 0d, 0d, 0, 0)
        };
    }
}

internal sealed record DocxLayoutPageSnapshot(
    double Width,
    double Height,
    int ItemCount,
    int TextLineCount,
    int InlineImageCount,
    int TableRowCount,
    double VerticalUsed,
    double TextLineHeightSum,
    double InlineImageHeightSum,
    double TableRowHeightSum,
    IReadOnlyList<DocxLayoutItemSnapshot> Items);

internal sealed record DocxLayoutItemSnapshot(
    string Kind,
    double X,
    double Y,
    double Width,
    double Height,
    int TextLength,
    int CellCount);

internal sealed record DocxLayoutPage(
    double Width,
    double Height,
    IReadOnlyList<DocxLayoutItem> Items);

internal abstract record DocxLayoutItem;

internal sealed record DocxTextLineLayout(
    string Text,
    DocxTextRun StyleRun,
    double FontSize,
    double X,
    double BaselineY,
    double Width,
    IReadOnlyList<DocxTextSegmentLayout> Segments) : DocxLayoutItem;

internal sealed record DocxTextSegmentLayout(
    string Text,
    DocxTextRun StyleRun,
    double X,
    double Width);

internal sealed record DocxInlineImageLayout(
    DocxInlineImage Image,
    double X,
    double Y,
    double Width,
    double Height,
    int PageIndex) : DocxLayoutItem;

internal sealed record DocxTableRowLayout(
    IReadOnlyList<DocxTableCellLayout> Cells,
    double Y,
    double Height) : DocxLayoutItem;

internal sealed record DocxTableCellLayout(
    DocxTableCell Cell,
    double X,
    double Y,
    double Width,
    double Height,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages);

internal readonly record struct DocxFontResources(PdfEmbeddedFont? Embedded, PdfFontResource? Resource);

internal interface IDocxTextMeasurer
{
    double MeasureText(DocxTextRun? run, string text, double fontSize);
}

internal sealed class DocxEmbeddedTextMeasurer(PdfEmbeddedFont embedded) : IDocxTextMeasurer
{
    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        return embedded.MeasureTextPoints(text, fontSize);
    }
}

internal sealed class DocxLayoutEngine
{
    private const double BaselineOffsetFactor = 0.94d;
    private const double DefaultTableRowHeight = 16d;

    public DocxLayout Create(DocxDocument document, PdfEmbeddedFont? embedded)
    {
        IDocxTextMeasurer? textMeasurer = embedded is null ? null : new DocxEmbeddedTextMeasurer(embedded);
        return Create(document, textMeasurer);
    }

    internal DocxLayout Create(DocxDocument document, IDocxTextMeasurer? textMeasurer)
    {
        var pages = new List<DocxLayoutPage>();
        var currentItems = new List<DocxLayoutItem>();
        double x = document.MarginLeftPoints;
        double width = Math.Max(1d, document.PageWidthPoints - document.MarginLeftPoints - document.MarginRightPoints);
        double cursorY = document.PageHeightPoints - document.MarginTopPoints;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;

        void FinishPage()
        {
            pages.Add(new DocxLayoutPage(document.PageWidthPoints, document.PageHeightPoints, currentItems.ToArray()));
            currentItems = [];
            cursorY = document.PageHeightPoints - document.MarginTopPoints;
            pendingSpacingAfter = 0d;
            previousParagraph = null;
        }

        bool HasPageContent() => currentItems.Count > 0;

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            DocxBodyElement element = document.BodyElements[elementIndex];
            if (element is DocxPageBreakElement)
            {
                if (HasPageContent())
                {
                    FinishPage();
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                LayoutTable(tableElement.Table, document, textMeasurer, () => pages.Count + 1, ref currentItems, ref cursorY, x, width, FinishPage, HasPageContent);
                continue;
            }

            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            cursorY -= ShouldSuppressContextualSpacing(previousParagraph, paragraph)
                ? 0d
                : Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            pendingSpacingAfter = 0d;
            double paragraphFontSize = paragraph.Runs.Count == 0 ? 11d : paragraph.Runs.Max(r => r.FontSize);
            double lineHeight = paragraph.LineSpacingPoints ?? paragraphFontSize * paragraph.LineSpacingFactor;
            if (textMeasurer is not null &&
                HasPageContent() &&
                ShouldKeepParagraphBlockTogether(paragraph) &&
                cursorY - EstimateKeptParagraphBlockHeight(document.BodyElements, elementIndex, width, textMeasurer) < document.MarginBottomPoints)
            {
                FinishPage();
            }

            if (textMeasurer is not null && paragraph.Runs.Count > 0)
            {
                string text = string.Concat(paragraph.Runs.Select(r => r.Text));
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, paragraphFontSize, textMeasurer);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
                double paragraphX = x + textStartOffset;
                double paragraphWidth = Math.Max(1d, width - textStartOffset - GetParagraphRightInset(paragraph));
                DocxTextRun firstRun = paragraph.Runs[0];
                bool firstLine = true;
                double continuationParagraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                string[] lines = WrapTextLines(text, paragraphWidth, continuationParagraphWidth, paragraphFontSize, firstRun, textMeasurer).ToArray();
                if (ShouldMoveParagraphForWidowControl(paragraph, lines.Length, cursorY, lineHeight, document.MarginBottomPoints, HasPageContent()))
                {
                    FinishPage();
                }

                foreach (string line in lines)
                {
                    if (cursorY - lineHeight < document.MarginBottomPoints && HasPageContent())
                    {
                        FinishPage();
                    }

                    double lineWidth = firstLine && paragraph.ListLabel is not null
                        ? textMeasurer.MeasureText(firstRun, line, paragraphFontSize)
                        : MeasureTextSegments(line, paragraph.Runs, paragraphFontSize, textMeasurer);
                    double lineX = paragraph.Alignment switch
                    {
                        DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                        DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                        _ => paragraphX
                    };
                    double baselineOffset = paragraph.LineSpacingPoints is null
                        ? paragraphFontSize * BaselineOffsetFactor
                        : Math.Max(0d, lineHeight - paragraphFontSize * 0.299d);
                    IReadOnlyList<DocxTextSegmentLayout> segments = firstLine && paragraph.ListLabel is not null
                        ? CreateNumberedLineSegments(paragraph.ListLabel, line, firstRun, x + labelStartOffset, lineX, paragraphFontSize, textMeasurer)
                        : CreateTextSegments(line, paragraph.Runs, lineX, paragraphFontSize, textMeasurer);
                    double effectiveX = firstLine && paragraph.ListLabel is not null ? x + labelStartOffset : lineX;
                    double effectiveWidth = firstLine && paragraph.ListLabel is not null
                        ? Math.Max(lineX + lineWidth, x + labelStartOffset + textMeasurer.MeasureText(firstRun, paragraph.ListLabel.Text, paragraphFontSize)) - (x + labelStartOffset)
                        : lineWidth;
                    currentItems.Add(new DocxTextLineLayout(
                        firstLine && paragraph.ListLabel is not null ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line : line,
                        firstRun,
                        paragraphFontSize,
                        effectiveX,
                        cursorY - baselineOffset,
                        effectiveWidth,
                        segments));
                    firstLine = false;
                    paragraphX = x + continuationTextStartOffset;
                    paragraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                    cursorY -= lineHeight;
                }
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(width, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                if (cursorY - imageHeight < document.MarginBottomPoints && HasPageContent())
                {
                    FinishPage();
                }

                double imageX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => x + Math.Max(0, width - imageWidth) / 2d,
                    DocxTextAlignment.Right => x + Math.Max(0, width - imageWidth),
                    _ => x
                };
                currentItems.Add(new DocxInlineImageLayout(image, imageX, cursorY - imageHeight, imageWidth, imageHeight, pages.Count + 1));
                cursorY -= imageHeight + 6d;
            }

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
            previousParagraph = paragraph;
        }

        if (HasPageContent() || pages.Count == 0)
        {
            FinishPage();
        }

        return new DocxLayout(pages.ToArray());
    }

    private static bool ShouldKeepParagraphBlockTogether(DocxParagraph paragraph)
    {
        return paragraph.KeepRules.KeepLines == true || paragraph.KeepRules.KeepNext == true;
    }

    private static bool ShouldMoveParagraphForWidowControl(
        DocxParagraph paragraph,
        int lineCount,
        double cursorY,
        double lineHeight,
        double marginBottom,
        bool hasPageContent)
    {
        if (paragraph.KeepRules.WidowControl != true ||
            lineCount <= 1 ||
            !hasPageContent)
        {
            return false;
        }

        int fittingLineCount = (int)Math.Floor(Math.Max(0d, cursorY - marginBottom) / lineHeight);
        return fittingLineCount > 0 &&
            fittingLineCount < lineCount &&
            (fittingLineCount == 1 || lineCount - fittingLineCount == 1);
    }

    private static bool ShouldSuppressContextualSpacing(DocxParagraph? previousParagraph, DocxParagraph paragraph)
    {
        return paragraph.Spacing.ContextualSpacing == true &&
            previousParagraph?.StyleId is not null &&
            paragraph.StyleId is not null &&
            string.Equals(previousParagraph.StyleId, paragraph.StyleId, StringComparison.Ordinal);
    }

    private static double EstimateKeptParagraphBlockHeight(
        IReadOnlyList<DocxBodyElement> elements,
        int elementIndex,
        double availableWidth,
        IDocxTextMeasurer textMeasurer)
    {
        if (elements[elementIndex] is not DocxParagraphElement paragraphElement)
        {
            return 0d;
        }

        DocxParagraph paragraph = paragraphElement.Paragraph;
        double height = EstimateParagraphContentHeight(paragraph, availableWidth, textMeasurer);
        int nextSearchIndex = elementIndex + 1;
        while (paragraph.KeepRules.KeepNext == true &&
            TryFindNextKeepTarget(elements, nextSearchIndex, out int nextIndex, out DocxBodyElement? next))
        {
            if (next is DocxParagraphElement nextParagraph)
            {
                height += Math.Max(paragraph.SpacingAfterPoints, nextParagraph.Paragraph.SpacingBeforePoints);
                height += EstimateParagraphContentHeight(nextParagraph.Paragraph, availableWidth, textMeasurer);
                paragraph = nextParagraph.Paragraph;
                nextSearchIndex = nextIndex + 1;
                continue;
            }

            if (next is DocxTableElement nextTable)
            {
                height += paragraph.SpacingAfterPoints;
                height += EstimateFirstTableRowHeight(nextTable.Table, availableWidth, textMeasurer);
            }

            break;
        }

        return height;
    }

    private static bool TryFindNextKeepTarget(IReadOnlyList<DocxBodyElement> elements, int startIndex, out DocxBodyElement? target)
    {
        bool found = TryFindNextKeepTarget(elements, startIndex, out _, out DocxBodyElement? indexedTarget);
        target = indexedTarget;
        return found;
    }

    private static bool TryFindNextKeepTarget(IReadOnlyList<DocxBodyElement> elements, int startIndex, out int targetIndex, out DocxBodyElement? target)
    {
        for (int i = startIndex; i < elements.Count; i++)
        {
            if (elements[i] is DocxParagraphElement or DocxTableElement)
            {
                targetIndex = i;
                target = elements[i];
                return true;
            }

            if (elements[i] is DocxPageBreakElement or DocxSectionBreakElement)
            {
                break;
            }
        }

        targetIndex = -1;
        target = null;
        return false;
    }

    private static double EstimateParagraphContentHeight(DocxParagraph paragraph, double availableWidth, IDocxTextMeasurer textMeasurer)
    {
        double height = 0d;
        if (paragraph.Runs.Count != 0)
        {
            double fontSize = paragraph.Runs.Max(run => run.FontSize);
            double lineHeight = paragraph.LineSpacingPoints ?? fontSize * paragraph.LineSpacingFactor;
            string text = string.Concat(paragraph.Runs.Select(run => run.Text));
            DocxTextRun firstRun = paragraph.Runs[0];
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double firstParagraphWidth = Math.Max(1d, availableWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = Math.Max(1d, availableWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            height += WrapTextLines(text, firstParagraphWidth, continuationParagraphWidth, fontSize, firstRun, textMeasurer).Count() * lineHeight;
        }

        foreach (DocxInlineImage image in paragraph.Images)
        {
            double imageWidth = Math.Min(availableWidth, image.WidthPoints);
            double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
            height += imageHeight + 6d;
        }

        return height;
    }

    private static double EstimateFirstTableRowHeight(DocxTable table, double availableWidth, IDocxTextMeasurer textMeasurer)
    {
        DocxTableRow? row = table.Rows.FirstOrDefault();
        if (row is null)
        {
            return 0d;
        }

        double gridTableWidth = table.ColumnWidthsPoints.Sum();
        double targetTableWidth = Math.Min(availableWidth, ResolvePreferredTableWidth(table, availableWidth) ?? gridTableWidth);
        IReadOnlyList<double> effectiveColumns = GetEffectiveTableColumnWidths(table, targetTableWidth);
        double rawTableWidth = effectiveColumns.Sum();
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double contentHeight = row.Cells
            .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer))
            .DefaultIfEmpty(0d)
            .Max();
        return Math.Max(row.HeightPoints ?? DefaultTableRowHeight, contentHeight);
    }

    private static double GetParagraphTextStartOffset(DocxParagraph paragraph)
    {
        if (paragraph.ListLabel is null)
        {
            return 0d;
        }

        DocxNumberingIndent indent = paragraph.ListLabel.Indent;
        double left = indent.LeftPoints ?? 0d;
        double firstLine = indent.FirstLinePoints ?? 0d;
        return Math.Max(0d, left + firstLine);
    }

    private static double GetParagraphFirstLineTextStartOffset(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer textMeasurer)
    {
        if (paragraph.ListLabel is null)
        {
            return 0d;
        }

        if (IsNumberingTabSuffix(paragraph.ListLabel))
        {
            return GetParagraphTextStartOffset(paragraph);
        }

        double gap = IsNumberingSpaceSuffix(paragraph.ListLabel)
            ? textMeasurer.MeasureText(paragraph.Runs.FirstOrDefault(), " ", fontSize)
            : 0d;
        return Math.Max(
            0d,
            GetParagraphLabelStartOffset(paragraph) + textMeasurer.MeasureText(paragraph.Runs.FirstOrDefault(), paragraph.ListLabel.Text, fontSize) + gap);
    }

    private static double GetParagraphLabelStartOffset(DocxParagraph paragraph)
    {
        if (paragraph.ListLabel is null)
        {
            return 0d;
        }

        DocxNumberingIndent indent = paragraph.ListLabel.Indent;
        double left = indent.LeftPoints ?? 0d;
        double hanging = indent.HangingPoints ?? 0d;
        double firstLine = indent.FirstLinePoints ?? 0d;
        return Math.Max(0d, left - hanging + firstLine);
    }

    private static double GetParagraphStartOffset(DocxParagraph paragraph)
    {
        return GetParagraphLabelStartOffset(paragraph);
    }

    private static double GetParagraphRightInset(DocxParagraph paragraph)
    {
        return paragraph.ListLabel?.Indent.RightPoints ?? 0d;
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateNumberedLineSegments(
        DocxListLabel label,
        string line,
        DocxTextRun styleRun,
        double labelX,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        double labelWidth = textMeasurer.MeasureText(styleRun, label.Text, fontSize);
        double lineWidth = textMeasurer.MeasureText(styleRun, line, fontSize);
        return
        [
            new DocxTextSegmentLayout(label.Text, styleRun, labelX, labelWidth),
            new DocxTextSegmentLayout(line, styleRun, lineX, lineWidth)
        ];
    }

    private static string GetListLabelTextSeparator(DocxListLabel label)
    {
        return label.SuffixValue switch
        {
            "nothing" => string.Empty,
            "space" => " ",
            _ => "\t"
        };
    }

    private static bool IsNumberingTabSuffix(DocxListLabel label)
    {
        return string.IsNullOrEmpty(label.SuffixValue) ||
            label.SuffixValue.Equals("tab", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNumberingSpaceSuffix(DocxListLabel label)
    {
        return label.SuffixValue.Equals("space", StringComparison.OrdinalIgnoreCase);
    }

    private static void LayoutTable(
        DocxTable table,
        DocxDocument document,
        IDocxTextMeasurer? textMeasurer,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x,
        double availableWidth,
        Action finishPage,
        Func<bool> hasPageContent)
    {
        double tableX = x + Math.Max(0d, table.IndentPoints ?? 0d);
        double tableAvailableWidth = Math.Max(1d, availableWidth - Math.Max(0d, table.IndentPoints ?? 0d));
        double gridTableWidth = table.ColumnWidthsPoints.Sum();
        double preferredTableWidth = ResolvePreferredTableWidth(table, tableAvailableWidth) ?? gridTableWidth;
        double targetTableWidth = Math.Min(tableAvailableWidth, preferredTableWidth);
        IReadOnlyList<double> effectiveColumns = GetEffectiveTableColumnWidths(table, targetTableWidth);
        double rawTableWidth = effectiveColumns.Sum();
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        double tableHeight = table.Rows.Sum(row => row.HeightPoints ?? DefaultTableRowHeight);
        if (cursorY - tableHeight < document.MarginBottomPoints && hasPageContent())
        {
            finishPage();
        }

        IReadOnlyList<DocxTableRow> headerRows = table.Rows.TakeWhile(row => row.IsHeader).ToArray();
        foreach (DocxTableRow row in table.Rows)
        {
            double rowHeight = MeasureTableRowHeight(table, row, effectiveColumns, scale, textMeasurer);
            if (cursorY - rowHeight < document.MarginBottomPoints && hasPageContent())
            {
                finishPage();
                if (!row.IsHeader)
                {
                    foreach (DocxTableRow headerRow in headerRows)
                    {
                        AddTableRowLayout(table, headerRow, effectiveColumns, scale, textMeasurer, getPageIndex, ref currentItems, ref cursorY, tableX);
                    }
                }
            }

            AddTableRowLayout(table, row, effectiveColumns, scale, textMeasurer, getPageIndex, ref currentItems, ref cursorY, tableX);
        }

        cursorY -= 6d;
    }

    private static IReadOnlyList<double> GetEffectiveTableColumnWidths(DocxTable table, double preferredTableWidth)
    {
        int columnCount = table.ColumnWidthsPoints.Count;
        if (columnCount == 0)
        {
            return table.ColumnWidthsPoints;
        }

        double?[] preferredWidths = new double?[columnCount];
        foreach (DocxTableRow row in table.Rows)
        {
            int gridColumnIndex = 0;
            foreach (DocxTableCell cell in row.Cells)
            {
                int span = Math.Max(1, cell.GridSpan);
                double? preferredWidth = ResolvePreferredCellWidth(cell, preferredTableWidth);
                if (span == 1 &&
                    gridColumnIndex < columnCount &&
                    preferredWidth is > 0d)
                {
                    preferredWidths[gridColumnIndex] = preferredWidth.Value;
                }

                gridColumnIndex += span;
            }

            if (preferredWidths.All(width => width is > 0d))
            {
                return preferredWidths.Select(width => width!.Value).ToArray();
            }
        }

        return table.ColumnWidthsPoints;
    }

    private static double? ResolvePreferredCellWidth(DocxTableCell cell, double preferredTableWidth)
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

    private static double? ResolvePreferredTableWidth(DocxTable table, double availableWidth)
    {
        if (table.PreferredWidthPoints is { } points)
        {
            return points;
        }

        if (table.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true &&
            int.TryParse(table.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
        {
            return Math.Max(0d, availableWidth * fiftiethsPercent / 5000d);
        }

        return null;
    }

    private static double MeasureTableRowHeight(
        DocxTable table,
        DocxTableRow row,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer)
    {
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double contentHeight = textMeasurer is null
            ? 0d
            : row.Cells
                .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer))
                .DefaultIfEmpty(0d)
                .Max();
        return Math.Max(row.HeightPoints ?? DefaultTableRowHeight, contentHeight);
    }

    private static void AddTableRowLayout(
        DocxTable table,
        DocxTableRow row,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x)
    {
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowHeight = MeasureTableRowHeight(table, row, effectiveColumns, scale, textMeasurer);
        double cellX = x;
        double cellY = cursorY - rowHeight;
        var cells = new List<DocxTableCellLayout>(row.Cells.Count);
        for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
        {
            double cellWidth = cellWidths[columnIndex];
            DocxTableCell cell = row.Cells[columnIndex];
            IReadOnlyList<DocxTextLineLayout> textLines = LayoutTableCellTextLines(cell, cellX, cellY, cellWidth, rowHeight, textMeasurer);
            IReadOnlyList<DocxInlineImageLayout> inlineImages = LayoutTableCellInlineImages(cell, cellX, cellY, cellWidth, rowHeight, textMeasurer, getPageIndex());
            cells.Add(new DocxTableCellLayout(cell, cellX, cellY, cellWidth, rowHeight, textLines, inlineImages));
            cellX += cellWidth + (table.CellSpacingPoints ?? 0d);
        }

        currentItems.Add(new DocxTableRowLayout(cells.ToArray(), cellY, rowHeight));
        cursorY -= rowHeight;
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

    private static double MeasureTableCellContentHeight(DocxTableCell cell, double cellWidth, IDocxTextMeasurer textMeasurer)
    {
        IReadOnlyList<DocxParagraph> paragraphs = cell.Paragraphs.Count == 0 && cell.Text.Length != 0
            ? [new DocxParagraph([new DocxTextRun(cell.Text, 11d, null, false, false, false, null, null)], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null)]
            : cell.Paragraphs;
        if (paragraphs.Count == 0)
        {
            return 0d;
        }

        double paddingLeft = cell.Margins.LeftPoints ?? 4d;
        double paddingRight = cell.Margins.RightPoints ?? 4d;
        double paddingTop = cell.Margins.TopPoints ?? 0d;
        double paddingBottom = cell.Margins.BottomPoints ?? 0d;
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double contentHeight = paddingTop + paddingBottom;
        foreach (DocxParagraph paragraph in paragraphs)
        {
            if (paragraph.Runs.Count == 0)
            {
                continue;
            }

            double fontSize = paragraph.Runs.Max(run => run.FontSize);
            double lineHeight = paragraph.LineSpacingPoints ?? fontSize * paragraph.LineSpacingFactor;
            string text = string.Concat(paragraph.Runs.Select(run => run.Text));
            DocxTextRun firstRun = paragraph.Runs[0];
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double firstParagraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = Math.Max(1d, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            int lineCount = WrapTextLines(text, firstParagraphWidth, continuationParagraphWidth, fontSize, firstRun, textMeasurer).Count();
            contentHeight += lineCount * lineHeight + paragraph.SpacingAfterPoints;
            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(textWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                contentHeight += imageHeight + 6d;
            }
        }

        return contentHeight;
    }

    private static IReadOnlyList<DocxTextLineLayout> LayoutTableCellTextLines(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        IDocxTextMeasurer? textMeasurer)
    {
        if (textMeasurer is null)
        {
            return [];
        }

        IReadOnlyList<DocxParagraph> paragraphs = cell.Paragraphs.Count == 0 && cell.Text.Length != 0
            ? [new DocxParagraph([new DocxTextRun(cell.Text, 11d, null, false, false, false, null, null)], [], null, DocxTextAlignment.Left, null, 0d, 0d, 1d, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, null)]
            : cell.Paragraphs;
        if (paragraphs.Count == 0)
        {
            return [];
        }

        double paddingLeft = cell.Margins.LeftPoints ?? 4d;
        double paddingRight = cell.Margins.RightPoints ?? 4d;
        double paddingTop = cell.Margins.TopPoints ?? 0d;
        double paddingBottom = cell.Margins.BottomPoints ?? 0d;
        const double legacyBaselineInset = 17d;
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - legacyBaselineInset - paddingTop;
        double cursorY = startBaselineY;
        var lines = new List<DocxTextLineLayout>();
        foreach (DocxParagraph paragraph in paragraphs)
        {
            if (paragraph.Runs.Count == 0)
            {
                continue;
            }

            double fontSize = paragraph.Runs.Max(run => run.FontSize);
            double lineHeight = paragraph.LineSpacingPoints ?? fontSize * paragraph.LineSpacingFactor;
            DocxTextRun firstRun = paragraph.Runs[0];
            string text = string.Concat(paragraph.Runs.Select(run => run.Text));
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
            double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
            double paragraphX = cellX + paddingLeft + textStartOffset;
            double paragraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
            bool firstLine = true;
            foreach (string line in WrapTextLines(text, paragraphWidth, continuationParagraphWidth, fontSize, firstRun, textMeasurer))
            {
                double lineWidth = firstLine && paragraph.ListLabel is not null
                    ? textMeasurer.MeasureText(firstRun, line, fontSize)
                    : MeasureTextSegments(line, paragraph.Runs, fontSize, textMeasurer);
                double lineX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                    DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                    _ => paragraphX
                };
                IReadOnlyList<DocxTextSegmentLayout> segments = firstLine && paragraph.ListLabel is not null
                    ? CreateNumberedLineSegments(paragraph.ListLabel, line, firstRun, cellX + paddingLeft + labelStartOffset, lineX, fontSize, textMeasurer)
                    : CreateTextSegments(line, paragraph.Runs, lineX, fontSize, textMeasurer);
                double effectiveX = firstLine && paragraph.ListLabel is not null ? cellX + paddingLeft + labelStartOffset : lineX;
                double effectiveWidth = firstLine && paragraph.ListLabel is not null
                    ? Math.Max(lineX + lineWidth, cellX + paddingLeft + labelStartOffset + textMeasurer.MeasureText(firstRun, paragraph.ListLabel.Text, fontSize)) - (cellX + paddingLeft + labelStartOffset)
                    : lineWidth;
                lines.Add(new DocxTextLineLayout(
                    firstLine && paragraph.ListLabel is not null ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line : line,
                    firstRun,
                    fontSize,
                    effectiveX,
                    cursorY,
                    effectiveWidth,
                    segments));
                firstLine = false;
                paragraphX = cellX + paddingLeft + continuationTextStartOffset;
                paragraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                cursorY -= lineHeight;
            }

            cursorY -= paragraph.SpacingAfterPoints;
        }

        if (lines.Count == 0)
        {
            return lines;
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - legacyBaselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = cell.VerticalAlignmentValue?.Equals("bottom", StringComparison.OrdinalIgnoreCase) == true
            ? extra
            : cell.VerticalAlignmentValue?.Equals("center", StringComparison.OrdinalIgnoreCase) == true
                ? extra / 2d
                : 0d;
        return verticalOffset == 0d ? lines : ShiftTextLines(lines, -verticalOffset);
    }

    private static IReadOnlyList<DocxInlineImageLayout> LayoutTableCellInlineImages(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        IDocxTextMeasurer? textMeasurer,
        int pageIndex)
    {
        IReadOnlyList<DocxParagraph> paragraphs = cell.Paragraphs;
        if (paragraphs.Count == 0 || !paragraphs.Any(paragraph => paragraph.Images.Count != 0))
        {
            return [];
        }

        double paddingLeft = cell.Margins.LeftPoints ?? 4d;
        double paddingRight = cell.Margins.RightPoints ?? 4d;
        double paddingTop = cell.Margins.TopPoints ?? 0d;
        double paddingBottom = cell.Margins.BottomPoints ?? 0d;
        const double legacyBaselineInset = 17d;
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - legacyBaselineInset - paddingTop;
        double cursorY = startBaselineY;
        var images = new List<DocxInlineImageLayout>();
        foreach (DocxParagraph paragraph in paragraphs)
        {
            if (textMeasurer is not null && paragraph.Runs.Count != 0)
            {
                double fontSize = paragraph.Runs.Max(run => run.FontSize);
                double lineHeight = paragraph.LineSpacingPoints ?? fontSize * paragraph.LineSpacingFactor;
                string text = string.Concat(paragraph.Runs.Select(run => run.Text));
                DocxTextRun firstRun = paragraph.Runs[0];
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                double firstParagraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                double continuationParagraphWidth = Math.Max(1d, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                cursorY -= WrapTextLines(text, firstParagraphWidth, continuationParagraphWidth, fontSize, firstRun, textMeasurer).Count() * lineHeight;
                cursorY -= paragraph.SpacingAfterPoints;
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double paragraphX = cellX + paddingLeft + GetParagraphStartOffset(paragraph);
                double paragraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                double imageWidth = Math.Min(paragraphWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                double imageX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - imageWidth) / 2d,
                    DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - imageWidth),
                    _ => paragraphX
                };
                images.Add(new DocxInlineImageLayout(image, imageX, cursorY - imageHeight, imageWidth, imageHeight, pageIndex));
                cursorY -= imageHeight + 6d;
            }
        }

        if (images.Count == 0)
        {
            return images;
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - legacyBaselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = cell.VerticalAlignmentValue?.Equals("bottom", StringComparison.OrdinalIgnoreCase) == true
            ? extra
            : cell.VerticalAlignmentValue?.Equals("center", StringComparison.OrdinalIgnoreCase) == true
                ? extra / 2d
                : 0d;
        return verticalOffset == 0d
            ? images
            : images.Select(image => image with { Y = image.Y - verticalOffset }).ToArray();
    }

    private static IReadOnlyList<DocxTextLineLayout> ShiftTextLines(IReadOnlyList<DocxTextLineLayout> lines, double deltaY)
    {
        return lines
            .Select(line => line with { BaselineY = line.BaselineY + deltaY })
            .ToArray();
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateTextSegments(
        string line,
        IReadOnlyList<DocxTextRun> runs,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        if (line.Length == 0 || runs.Count == 0)
        {
            return [];
        }

        var segments = new List<DocxTextSegmentLayout>();
        int lineOffset = 0;
        double segmentX = lineX;
        foreach (DocxTextRun run in runs)
        {
            if (lineOffset >= line.Length)
            {
                break;
            }

            string runText = run.Text;
            int runOffset = 0;
            while (runOffset < runText.Length && lineOffset < line.Length)
            {
                if (line[lineOffset] != runText[runOffset])
                {
                    runOffset++;
                    continue;
                }

                int start = runOffset;
                while (runOffset < runText.Length && lineOffset < line.Length && line[lineOffset] == runText[runOffset])
                {
                    runOffset++;
                    lineOffset++;
                }

                string segmentText = runText[start..runOffset];
                double width = textMeasurer.MeasureText(run, segmentText, fontSize);
                segments.Add(new DocxTextSegmentLayout(segmentText, run, segmentX, width));
                segmentX += width;
            }
        }

        return segments.Count == 0
            ? [new DocxTextSegmentLayout(line, runs[0], lineX, textMeasurer.MeasureText(runs[0], line, fontSize))]
            : segments;
    }

    private static double MeasureTextSegments(
        string line,
        IReadOnlyList<DocxTextRun> runs,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        return CreateTextSegments(line, runs, 0d, fontSize, textMeasurer).Sum(segment => segment.Width);
    }

    private static IEnumerable<string> WrapTextLines(string text, double maxWidth, double fontSize, DocxTextRun? run, IDocxTextMeasurer textMeasurer)
    {
        return WrapTextLines(text, maxWidth, maxWidth, fontSize, run, textMeasurer);
    }

    private static IEnumerable<string> WrapTextLines(string text, double firstLineMaxWidth, double continuationLineMaxWidth, double fontSize, DocxTextRun? run, IDocxTextMeasurer textMeasurer)
    {
        int lineIndex = 0;
        foreach (string segment in text.Split('\n'))
        {
            bool yielded = false;
            foreach (string line in WrapWords(segment, index => index == 0 && lineIndex == 0 ? firstLineMaxWidth : continuationLineMaxWidth, fontSize, run, textMeasurer))
            {
                yielded = true;
                yield return line;
                lineIndex++;
            }

            if (!yielded && segment.Length == 0)
            {
                yield return string.Empty;
                lineIndex++;
            }
        }
    }

    private static IEnumerable<string> WrapWords(string text, double maxWidth, double fontSize, DocxTextRun? run, IDocxTextMeasurer textMeasurer)
    {
        return WrapWords(text, _ => maxWidth, fontSize, run, textMeasurer);
    }

    private static IEnumerable<string> WrapWords(string text, Func<int, double> maxWidth, double fontSize, DocxTextRun? run, IDocxTextMeasurer textMeasurer)
    {
        IReadOnlyList<string> tokens = TokenizeSpaces(text);
        if (tokens.Count == 0)
        {
            yield break;
        }

        var line = new System.Text.StringBuilder();
        int lineIndex = 0;
        foreach (string token in tokens)
        {
            string candidate = line + token;
            if (line.Length > 0 &&
                line.ToString().Any(c => !char.IsWhiteSpace(c)) &&
                !string.IsNullOrWhiteSpace(token) &&
                textMeasurer.MeasureText(run, candidate, fontSize) > maxWidth(lineIndex))
            {
                yield return line.ToString();
                lineIndex++;
                line.Clear();
                line.Append(token);
            }
            else
            {
                line.Clear();
                line.Append(candidate);
            }
        }

        if (line.Length > 0)
        {
            yield return line.ToString();
        }
    }

    private static IReadOnlyList<string> TokenizeSpaces(string text)
    {
        if (text.Length == 0)
        {
            return [];
        }

        var tokens = new List<string>();
        int start = 0;
        bool inWhitespace = char.IsWhiteSpace(text[0]);
        for (int i = 1; i < text.Length; i++)
        {
            bool whitespace = char.IsWhiteSpace(text[i]);
            if (whitespace == inWhitespace)
            {
                continue;
            }

            tokens.Add(text[start..i]);
            start = i;
            inWhitespace = whitespace;
        }

        tokens.Add(text[start..]);
        return tokens;
    }
}
