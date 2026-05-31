using System.Globalization;
using Lokad.OoxPdf.Fonts;
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

internal sealed record DocxTextSpan(
    string Text,
    DocxTextRun StyleRun);

internal sealed record DocxWrappedTextLine(
    string Text,
    IReadOnlyList<DocxTextSpan> Spans);

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

internal interface IDocxLineMetricsProvider
{
    double MeasureSingleLineHeight(DocxTextRun? run, double fontSize);
}

internal static class DocxLineMetrics
{
    public static double MeasureOpenTypeSingleLineHeight(OpenTypeFont font, double fontSize)
    {
        if (font.UnitsPerEm == 0)
        {
            return fontSize;
        }

        double units = font.Os2.TypographicAscender - font.Os2.TypographicDescender + font.Os2.TypographicLineGap;
        if (units <= 0d)
        {
            units = font.Os2.WindowsAscender + font.Os2.WindowsDescender;
        }

        return Math.Max(fontSize, units * fontSize / font.UnitsPerEm);
    }
}

internal sealed class DocxEmbeddedTextMeasurer(PdfEmbeddedFont embedded) : IDocxTextMeasurer, IDocxLineMetricsProvider
{
    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        return embedded.MeasureTextPoints(text, fontSize);
    }

    public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureOpenTypeSingleLineHeight(embedded.Font, fontSize);
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
            double lineHeight = ResolveLineHeight(paragraph, paragraphFontSize, textMeasurer);
            if (textMeasurer is not null &&
                HasPageContent() &&
                ShouldKeepParagraphBlockTogether(paragraph) &&
                cursorY - EstimateKeptParagraphBlockHeight(document.BodyElements, elementIndex, width, textMeasurer) < document.MarginBottomPoints)
            {
                FinishPage();
            }

            if (textMeasurer is not null && paragraph.Runs.Count > 0)
            {
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, paragraphFontSize, textMeasurer);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
                double paragraphX = x + textStartOffset;
                double paragraphWidth = Math.Max(1d, width - textStartOffset - GetParagraphRightInset(paragraph));
                DocxTextRun firstRun = paragraph.Runs[0];
                IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
                bool firstLine = true;
                double continuationParagraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, paragraphFontSize, textMeasurer).ToArray();
                if (ShouldMoveParagraphForWidowControl(paragraph, lines.Length, cursorY, lineHeight, document.MarginBottomPoints, HasPageContent()))
                {
                    FinishPage();
                }

                foreach (DocxWrappedTextLine line in lines)
                {
                    if (cursorY - lineHeight < document.MarginBottomPoints && HasPageContent())
                    {
                        FinishPage();
                    }

                    double lineWidth = MeasureTextSpans(line.Spans, paragraphFontSize, textMeasurer);
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
                        ? CreateNumberedLineSegments(paragraph.ListLabel, line.Spans, firstRun, x + labelStartOffset, lineX, paragraphFontSize, textMeasurer)
                        : CreateTextSegments(line.Spans, lineX, paragraphFontSize, textMeasurer);
                    double effectiveX = firstLine && paragraph.ListLabel is not null ? x + labelStartOffset : lineX;
                    double effectiveWidth = firstLine && paragraph.ListLabel is not null
                        ? Math.Max(lineX + lineWidth, x + labelStartOffset + textMeasurer.MeasureText(firstRun, paragraph.ListLabel.Text, paragraphFontSize)) - (x + labelStartOffset)
                        : lineWidth;
                    currentItems.Add(new DocxTextLineLayout(
                        firstLine && paragraph.ListLabel is not null ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line.Text : line.Text,
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

    private static double ResolveLineHeight(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer? textMeasurer)
    {
        if (paragraph.LineSpacingPoints is { } exactLineHeight)
        {
            return exactLineHeight;
        }

        double singleLineHeight = textMeasurer is IDocxLineMetricsProvider metricsProvider
            ? metricsProvider.MeasureSingleLineHeight(paragraph.Runs.FirstOrDefault(), fontSize)
            : fontSize;
        return singleLineHeight * paragraph.LineSpacingFactor;
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
            double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double firstParagraphWidth = Math.Max(1d, availableWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = Math.Max(1d, availableWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            height += WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer).Count() * lineHeight;
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
        IReadOnlyList<DocxTextSpan> lineSpans,
        DocxTextRun styleRun,
        double labelX,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        double labelWidth = textMeasurer.MeasureText(styleRun, label.Text, fontSize);
        return
        [
            new DocxTextSegmentLayout(label.Text, styleRun, labelX, labelWidth),
            .. CreateTextSegments(lineSpans, lineX, fontSize, textMeasurer)
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
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxParagraph paragraph in paragraphs)
        {
            if (paragraph.Runs.Count == 0)
            {
                continue;
            }

            contentHeight += ShouldSuppressContextualSpacing(previousParagraph, paragraph)
                ? 0d
                : Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            pendingSpacingAfter = 0d;
            double fontSize = paragraph.Runs.Max(run => run.FontSize);
            double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double firstParagraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = Math.Max(1d, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            int lineCount = WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer).Count();
            contentHeight += lineCount * lineHeight;
            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(textWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                contentHeight += imageHeight + 6d;
            }

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
            previousParagraph = paragraph;
        }

        contentHeight += pendingSpacingAfter;
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
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxParagraph paragraph in paragraphs)
        {
            if (paragraph.Runs.Count == 0)
            {
                continue;
            }

            cursorY -= ShouldSuppressContextualSpacing(previousParagraph, paragraph)
                ? 0d
                : Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            pendingSpacingAfter = 0d;
            double fontSize = paragraph.Runs.Max(run => run.FontSize);
            double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
            DocxTextRun firstRun = paragraph.Runs[0];
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
            double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
            double paragraphX = cellX + paddingLeft + textStartOffset;
            double paragraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
            bool firstLine = true;
            foreach (DocxWrappedTextLine line in WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, fontSize, textMeasurer))
            {
                double lineWidth = MeasureTextSpans(line.Spans, fontSize, textMeasurer);
                double lineX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                    DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                    _ => paragraphX
                };
                IReadOnlyList<DocxTextSegmentLayout> segments = firstLine && paragraph.ListLabel is not null
                    ? CreateNumberedLineSegments(paragraph.ListLabel, line.Spans, firstRun, cellX + paddingLeft + labelStartOffset, lineX, fontSize, textMeasurer)
                    : CreateTextSegments(line.Spans, lineX, fontSize, textMeasurer);
                double effectiveX = firstLine && paragraph.ListLabel is not null ? cellX + paddingLeft + labelStartOffset : lineX;
                double effectiveWidth = firstLine && paragraph.ListLabel is not null
                    ? Math.Max(lineX + lineWidth, cellX + paddingLeft + labelStartOffset + textMeasurer.MeasureText(firstRun, paragraph.ListLabel.Text, fontSize)) - (cellX + paddingLeft + labelStartOffset)
                    : lineWidth;
                lines.Add(new DocxTextLineLayout(
                    firstLine && paragraph.ListLabel is not null ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line.Text : line.Text,
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

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
            previousParagraph = paragraph;
        }

        cursorY -= pendingSpacingAfter;
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
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxParagraph paragraph in paragraphs)
        {
            cursorY -= ShouldSuppressContextualSpacing(previousParagraph, paragraph)
                ? 0d
                : Math.Max(pendingSpacingAfter, paragraph.SpacingBeforePoints);
            pendingSpacingAfter = 0d;
            if (textMeasurer is not null && paragraph.Runs.Count != 0)
            {
                double fontSize = paragraph.Runs.Max(run => run.FontSize);
                double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
                IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                double firstParagraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                double continuationParagraphWidth = Math.Max(1d, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                cursorY -= WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer).Count() * lineHeight;
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

            pendingSpacingAfter = paragraph.SpacingAfterPoints;
            previousParagraph = paragraph;
        }

        cursorY -= pendingSpacingAfter;
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

    private static IReadOnlyList<DocxTextSpan> CreateTextSpans(IReadOnlyList<DocxTextRun> runs)
    {
        return runs
            .Where(run => run.Text.Length != 0)
            .Select(run => new DocxTextSpan(run.Text, run))
            .ToArray();
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        var segments = new List<DocxTextSegmentLayout>(spans.Count);
        double segmentX = lineX;
        foreach (DocxTextSpan span in spans)
        {
            double width = textMeasurer.MeasureText(span.StyleRun, span.Text, fontSize);
            segments.Add(new DocxTextSegmentLayout(span.Text, span.StyleRun, segmentX, width));
            segmentX += width;
        }

        return segments;
    }

    private static double MeasureTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        return spans.Sum(span => textMeasurer.MeasureText(span.StyleRun, span.Text, fontSize));
    }

    private static IEnumerable<DocxWrappedTextLine> WrapTextLines(
        IReadOnlyList<DocxTextSpan> spans,
        double firstLineMaxWidth,
        double continuationLineMaxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        string text = string.Concat(spans.Select(span => span.Text));
        int lineIndex = 0;
        int segmentStart = 0;
        while (segmentStart <= text.Length)
        {
            int breakIndex = text.IndexOf('\n', segmentStart);
            int segmentLength = breakIndex < 0 ? text.Length - segmentStart : breakIndex - segmentStart;
            bool yielded = false;
            foreach (DocxWrappedTextLine line in WrapWords(text, spans, segmentStart, segmentLength, index => index == 0 && lineIndex == 0 ? firstLineMaxWidth : continuationLineMaxWidth, fontSize, textMeasurer))
            {
                yielded = true;
                yield return line;
                lineIndex++;
            }

            if (!yielded && segmentLength == 0)
            {
                yield return new DocxWrappedTextLine(string.Empty, []);
                lineIndex++;
            }

            if (breakIndex < 0)
            {
                yield break;
            }

            segmentStart = breakIndex + 1;
        }
    }

    private static IEnumerable<DocxWrappedTextLine> WrapWords(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int segmentStart,
        int segmentLength,
        Func<int, double> maxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer)
    {
        IReadOnlyList<TextToken> tokens = TokenizeSpaces(text, segmentStart, segmentLength);
        if (tokens.Count == 0)
        {
            yield break;
        }

        int lineStart = tokens[0].Start;
        int lineLength = 0;
        int lineIndex = 0;
        foreach (TextToken token in tokens)
        {
            int candidateLength = token.Start + token.Length - lineStart;
            bool lineHasNonWhitespace = HasNonWhitespace(text, lineStart, lineLength);
            if (lineLength > 0 &&
                lineHasNonWhitespace &&
                !string.IsNullOrWhiteSpace(token.Text) &&
                MeasureTextSpans(SliceTextSpans(spans, lineStart, candidateLength), fontSize, textMeasurer) > maxWidth(lineIndex))
            {
                yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
                lineIndex++;
                lineStart = token.Start;
                lineLength = token.Length;
            }
            else
            {
                lineLength = candidateLength;
            }
        }

        if (lineLength > 0)
        {
            yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
        }
    }

    private static DocxWrappedTextLine CreateWrappedTextLine(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length)
    {
        return new DocxWrappedTextLine(text.Substring(start, length), SliceTextSpans(spans, start, length));
    }

    private static IReadOnlyList<DocxTextSpan> SliceTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length)
    {
        if (length == 0)
        {
            return [];
        }

        var sliced = new List<DocxTextSpan>();
        int spanStart = 0;
        int end = start + length;
        foreach (DocxTextSpan span in spans)
        {
            int spanEnd = spanStart + span.Text.Length;
            int sliceStart = Math.Max(start, spanStart);
            int sliceEnd = Math.Min(end, spanEnd);
            if (sliceStart < sliceEnd)
            {
                sliced.Add(new DocxTextSpan(span.Text[(sliceStart - spanStart)..(sliceEnd - spanStart)], span.StyleRun));
            }

            if (spanEnd >= end)
            {
                break;
            }

            spanStart = spanEnd;
        }

        return sliced;
    }

    private static IReadOnlyList<TextToken> TokenizeSpaces(string text, int start, int length)
    {
        if (length == 0)
        {
            return [];
        }

        string segment = text.Substring(start, length);
        return TokenizeSpaces(segment)
            .Select(token => new TextToken(token.Text, start + token.Start, token.Length))
            .ToArray();
    }

    private static bool HasNonWhitespace(string text, int start, int length)
    {
        for (int i = start; i < start + length; i++)
        {
            if (!char.IsWhiteSpace(text[i]))
            {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<TextToken> TokenizeSpaces(string text)
    {
        if (text.Length == 0)
        {
            return [];
        }

        var tokens = new List<TextToken>();
        int start = 0;
        bool inWhitespace = char.IsWhiteSpace(text[0]);
        for (int i = 1; i < text.Length; i++)
        {
            bool whitespace = char.IsWhiteSpace(text[i]);
            if (whitespace == inWhitespace)
            {
                continue;
            }

            tokens.Add(new TextToken(text[start..i], start, i - start));
            start = i;
            inWhitespace = whitespace;
        }

        tokens.Add(new TextToken(text[start..], start, text.Length - start));
        return tokens;
    }

    private readonly record struct TextToken(string Text, int Start, int Length);
}
