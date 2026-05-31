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
        return new DocxLayoutPageSnapshot(
            page.Width,
            page.Height,
            items.Count,
            items.Count(item => item.Kind == "TextLine"),
            items.Count(item => item.Kind == "InlineImage"),
            items.Count(item => item.Kind == "TableRow"),
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

internal sealed class DocxLayoutEngine
{
    private const double BaselineOffsetFactor = 0.94d;
    private const double DefaultTableRowHeight = 16d;

    public DocxLayout Create(DocxDocument document, PdfEmbeddedFont? embedded)
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
                LayoutTable(tableElement.Table, document, embedded, () => pages.Count + 1, ref currentItems, ref cursorY, x, width, FinishPage, HasPageContent);
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
            if (embedded is not null &&
                HasPageContent() &&
                ShouldKeepParagraphBlockTogether(paragraph) &&
                cursorY - EstimateKeptParagraphBlockHeight(document.BodyElements, elementIndex, width, embedded) < document.MarginBottomPoints)
            {
                FinishPage();
            }

            if (embedded is not null && paragraph.Runs.Count > 0)
            {
                string text = paragraph.ListLabel is null
                    ? string.Concat(paragraph.Runs.Select(r => r.Text))
                    : paragraph.ListLabel.Text + " " + string.Concat(paragraph.Runs.Select(r => r.Text));
                double paragraphX = x + GetParagraphStartOffset(paragraph);
                double paragraphWidth = Math.Max(1d, width - GetParagraphStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                DocxTextRun firstRun = paragraph.Runs[0];
                foreach (string line in WrapWords(text, paragraphWidth, paragraphFontSize, embedded))
                {
                    if (cursorY - lineHeight < document.MarginBottomPoints && HasPageContent())
                    {
                        FinishPage();
                    }

                    double lineWidth = embedded.MeasureTextPoints(line, paragraphFontSize);
                    double lineX = paragraph.Alignment switch
                    {
                        DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                        DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                        _ => paragraphX
                    };
                    double baselineOffset = paragraph.LineSpacingPoints is null
                        ? paragraphFontSize * BaselineOffsetFactor
                        : Math.Max(0d, lineHeight - paragraphFontSize * 0.299d);
                    currentItems.Add(new DocxTextLineLayout(
                        line,
                        firstRun,
                        paragraphFontSize,
                        lineX,
                        cursorY - baselineOffset,
                        lineWidth,
                        [new DocxTextSegmentLayout(line, firstRun, lineX, lineWidth)]));
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
        PdfEmbeddedFont embedded)
    {
        if (elements[elementIndex] is not DocxParagraphElement paragraphElement)
        {
            return 0d;
        }

        DocxParagraph paragraph = paragraphElement.Paragraph;
        double height = EstimateParagraphContentHeight(paragraph, availableWidth, embedded);
        if (paragraph.KeepRules.KeepNext == true && TryFindNextKeepTarget(elements, elementIndex + 1, out DocxBodyElement? next))
        {
            if (next is DocxParagraphElement nextParagraph)
            {
                height += Math.Max(paragraph.SpacingAfterPoints, nextParagraph.Paragraph.SpacingBeforePoints);
                height += EstimateParagraphContentHeight(nextParagraph.Paragraph, availableWidth, embedded);
            }
            else if (next is DocxTableElement nextTable)
            {
                height += paragraph.SpacingAfterPoints;
                height += EstimateFirstTableRowHeight(nextTable.Table, availableWidth, embedded);
            }
        }

        return height;
    }

    private static bool TryFindNextKeepTarget(IReadOnlyList<DocxBodyElement> elements, int startIndex, out DocxBodyElement? target)
    {
        for (int i = startIndex; i < elements.Count; i++)
        {
            if (elements[i] is DocxParagraphElement or DocxTableElement)
            {
                target = elements[i];
                return true;
            }

            if (elements[i] is DocxPageBreakElement or DocxSectionBreakElement)
            {
                break;
            }
        }

        target = null;
        return false;
    }

    private static double EstimateParagraphContentHeight(DocxParagraph paragraph, double availableWidth, PdfEmbeddedFont embedded)
    {
        double height = 0d;
        if (paragraph.Runs.Count != 0)
        {
            double fontSize = paragraph.Runs.Max(run => run.FontSize);
            double lineHeight = paragraph.LineSpacingPoints ?? fontSize * paragraph.LineSpacingFactor;
            string text = paragraph.ListLabel is null
                ? string.Concat(paragraph.Runs.Select(run => run.Text))
                : paragraph.ListLabel.Text + " " + string.Concat(paragraph.Runs.Select(run => run.Text));
            double paragraphWidth = Math.Max(1d, availableWidth - GetParagraphStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            height += WrapWords(text, paragraphWidth, fontSize, embedded).Count() * lineHeight;
        }

        foreach (DocxInlineImage image in paragraph.Images)
        {
            double imageWidth = Math.Min(availableWidth, image.WidthPoints);
            double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
            height += imageHeight + 6d;
        }

        return height;
    }

    private static double EstimateFirstTableRowHeight(DocxTable table, double availableWidth, PdfEmbeddedFont embedded)
    {
        DocxTableRow? row = table.Rows.FirstOrDefault();
        if (row is null)
        {
            return 0d;
        }

        double rawTableWidth = table.ColumnWidthsPoints.Sum();
        double targetTableWidth = Math.Min(availableWidth, table.PreferredWidthPoints ?? rawTableWidth);
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        double[] cellWidths = row.Cells
            .Select((_, columnIndex) => table.ColumnWidthsPoints[Math.Min(columnIndex, table.ColumnWidthsPoints.Count - 1)] * scale)
            .ToArray();
        double contentHeight = row.Cells
            .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], embedded))
            .DefaultIfEmpty(0d)
            .Max();
        return Math.Max(row.HeightPoints ?? DefaultTableRowHeight, contentHeight);
    }

    private static double GetParagraphStartOffset(DocxParagraph paragraph)
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

    private static double GetParagraphRightInset(DocxParagraph paragraph)
    {
        return paragraph.ListLabel?.Indent.RightPoints ?? 0d;
    }

    private static void LayoutTable(
        DocxTable table,
        DocxDocument document,
        PdfEmbeddedFont? embedded,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x,
        double availableWidth,
        Action finishPage,
        Func<bool> hasPageContent)
    {
        double rawTableWidth = table.ColumnWidthsPoints.Sum();
        double targetTableWidth = Math.Min(availableWidth, table.PreferredWidthPoints ?? rawTableWidth);
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        double tableHeight = table.Rows.Sum(row => row.HeightPoints ?? DefaultTableRowHeight);
        if (cursorY - tableHeight < document.MarginBottomPoints && hasPageContent())
        {
            finishPage();
        }

        IReadOnlyList<DocxTableRow> headerRows = table.Rows.TakeWhile(row => row.IsHeader).ToArray();
        foreach (DocxTableRow row in table.Rows)
        {
            double rowHeight = MeasureTableRowHeight(table, row, scale, embedded);
            if (cursorY - rowHeight < document.MarginBottomPoints && hasPageContent())
            {
                finishPage();
                if (!row.IsHeader)
                {
                    foreach (DocxTableRow headerRow in headerRows)
                    {
                        AddTableRowLayout(table, headerRow, scale, embedded, getPageIndex, ref currentItems, ref cursorY, x);
                    }
                }
            }

            AddTableRowLayout(table, row, scale, embedded, getPageIndex, ref currentItems, ref cursorY, x);
        }

        cursorY -= 6d;
    }

    private static double MeasureTableRowHeight(DocxTable table, DocxTableRow row, double scale, PdfEmbeddedFont? embedded)
    {
        double[] cellWidths = row.Cells
            .Select((_, columnIndex) => table.ColumnWidthsPoints[Math.Min(columnIndex, table.ColumnWidthsPoints.Count - 1)] * scale)
            .ToArray();
        double contentHeight = embedded is null
            ? 0d
            : row.Cells
                .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], embedded))
                .DefaultIfEmpty(0d)
                .Max();
        return Math.Max(row.HeightPoints ?? DefaultTableRowHeight, contentHeight);
    }

    private static void AddTableRowLayout(
        DocxTable table,
        DocxTableRow row,
        double scale,
        PdfEmbeddedFont? embedded,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x)
    {
        double[] cellWidths = row.Cells
            .Select((_, columnIndex) => table.ColumnWidthsPoints[Math.Min(columnIndex, table.ColumnWidthsPoints.Count - 1)] * scale)
            .ToArray();
        double rowHeight = MeasureTableRowHeight(table, row, scale, embedded);
        double cellX = x;
        double cellY = cursorY - rowHeight;
        var cells = new List<DocxTableCellLayout>(row.Cells.Count);
        for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
        {
            double cellWidth = cellWidths[columnIndex];
            DocxTableCell cell = row.Cells[columnIndex];
            IReadOnlyList<DocxTextLineLayout> textLines = LayoutTableCellTextLines(cell, cellX, cellY, cellWidth, rowHeight, embedded);
            IReadOnlyList<DocxInlineImageLayout> inlineImages = LayoutTableCellInlineImages(cell, cellX, cellY, cellWidth, rowHeight, embedded, getPageIndex());
            cells.Add(new DocxTableCellLayout(cell, cellX, cellY, cellWidth, rowHeight, textLines, inlineImages));
            cellX += cellWidth;
        }

        currentItems.Add(new DocxTableRowLayout(cells.ToArray(), cellY, rowHeight));
        cursorY -= rowHeight;
    }

    private static double MeasureTableCellContentHeight(DocxTableCell cell, double cellWidth, PdfEmbeddedFont embedded)
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
            string text = paragraph.ListLabel is null
                ? string.Concat(paragraph.Runs.Select(run => run.Text))
                : paragraph.ListLabel.Text + " " + string.Concat(paragraph.Runs.Select(run => run.Text));
            double paragraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            int lineCount = WrapWords(text, paragraphWidth, fontSize, embedded).Count();
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
        PdfEmbeddedFont? embedded)
    {
        if (embedded is null)
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
            string text = paragraph.ListLabel is null
                ? string.Concat(paragraph.Runs.Select(run => run.Text))
                : paragraph.ListLabel.Text + " " + string.Concat(paragraph.Runs.Select(run => run.Text));
            double paragraphX = cellX + paddingLeft + GetParagraphStartOffset(paragraph);
            double paragraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            foreach (string line in WrapWords(text, paragraphWidth, fontSize, embedded))
            {
                double lineWidth = embedded.MeasureTextPoints(line, fontSize);
                double lineX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                    DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                    _ => paragraphX
                };
                IReadOnlyList<DocxTextSegmentLayout> segments = CreateTextSegments(line, paragraph.Runs, lineX, fontSize, embedded);
                lines.Add(new DocxTextLineLayout(line, firstRun, fontSize, lineX, cursorY, lineWidth, segments));
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
        PdfEmbeddedFont? embedded,
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
            if (embedded is not null && paragraph.Runs.Count != 0)
            {
                double fontSize = paragraph.Runs.Max(run => run.FontSize);
                double lineHeight = paragraph.LineSpacingPoints ?? fontSize * paragraph.LineSpacingFactor;
                string text = paragraph.ListLabel is null
                    ? string.Concat(paragraph.Runs.Select(run => run.Text))
                    : paragraph.ListLabel.Text + " " + string.Concat(paragraph.Runs.Select(run => run.Text));
                double paragraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                cursorY -= WrapWords(text, paragraphWidth, fontSize, embedded).Count() * lineHeight;
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
        PdfEmbeddedFont embedded)
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
                double width = embedded.MeasureTextPoints(segmentText, fontSize);
                segments.Add(new DocxTextSegmentLayout(segmentText, run, segmentX, width));
                segmentX += width;
            }
        }

        return segments.Count == 0
            ? [new DocxTextSegmentLayout(line, runs[0], lineX, embedded.MeasureTextPoints(line, fontSize))]
            : segments;
    }

    private static IEnumerable<string> WrapWords(string text, double maxWidth, double fontSize, PdfEmbeddedFont embedded)
    {
        string[] words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0)
        {
            yield break;
        }

        var line = new System.Text.StringBuilder();
        foreach (string word in words)
        {
            string candidate = line.Length == 0 ? word : line + " " + word;
            if (line.Length > 0 && embedded.MeasureTextPoints(candidate, fontSize) > maxWidth)
            {
                yield return line.ToString();
                line.Clear();
                line.Append(word);
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
}
