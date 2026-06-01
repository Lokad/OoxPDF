using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxStructureSnapshot(
    int BlockCount,
    int ParagraphBlockCount,
    int TableBlockCount,
    int PageBreakBlockCount,
    int SectionBreakBlockCount,
    int BodyTextLength,
    int InlineImageCount,
    int FloatingDrawingCount,
    IReadOnlyList<DocxStructureBlockSnapshot> Blocks,
    IReadOnlyList<DocxStructureStorySnapshot> Stories,
    IReadOnlyList<DocxStructureFloatingDrawingSnapshot> FloatingDrawings,
    IReadOnlyList<DocxStructureStyleUsageSnapshot> StyleUsages,
    IReadOnlyList<DocxStructureListUsageSnapshot> ListUsages,
    IReadOnlyList<DocxStructureTableSnapshot> Tables,
    IReadOnlyList<DocxStructureTableAdjacencySnapshot> TableAdjacency)
{
    public static DocxStructureSnapshot FromDocument(DocxDocument document)
    {
        var blocks = new List<DocxStructureBlockSnapshot>(document.BodyElements.Count);
        var tables = new List<DocxStructureTableSnapshot>();
        var adjacency = new List<DocxStructureTableAdjacencySnapshot>();
        int tableIndex = 0;
        for (int blockIndex = 0; blockIndex < document.BodyElements.Count; blockIndex++)
        {
            DocxBodyElement element = document.BodyElements[blockIndex];
            string? previousKind = blockIndex == 0 ? null : GetBlockKind(document.BodyElements[blockIndex - 1]);
            string? nextKind = blockIndex + 1 >= document.BodyElements.Count ? null : GetBlockKind(document.BodyElements[blockIndex + 1]);
            switch (element)
            {
                case DocxParagraphElement paragraph:
                    blocks.Add(FromParagraph(blockIndex, previousKind, nextKind, paragraph.Paragraph));
                    break;
                case DocxTableElement table:
                    blocks.Add(FromTable(blockIndex, previousKind, nextKind, table.Table, tableIndex));
                    tables.Add(ToTableSnapshot(table.Table, tableIndex, blockIndex));
                    adjacency.Add(ToTableAdjacencySnapshot(document.BodyElements, table.Table, tableIndex, blockIndex, previousKind, nextKind));
                    tableIndex++;
                    break;
                case DocxPageBreakElement pageBreak:
                    blocks.Add(FromPageBreak(blockIndex, previousKind, nextKind, pageBreak));
                    break;
                case DocxSectionBreakElement sectionBreak:
                    blocks.Add(FromSectionBreak(blockIndex, previousKind, nextKind, sectionBreak));
                    break;
                default:
                    blocks.Add(new DocxStructureBlockSnapshot(blockIndex, "Unknown", previousKind, nextKind));
                    break;
            }
        }

        return new DocxStructureSnapshot(
            document.BodyElements.Count,
            blocks.Count(block => block.Kind == "Paragraph"),
            tables.Count,
            blocks.Count(block => block.Kind == "PageBreak"),
            blocks.Count(block => block.Kind == "SectionBreak"),
            blocks.Sum(block => block.TextLength),
            blocks.Sum(block => block.InlineImageCount),
            document.FloatingDrawings.Count,
            blocks,
            ToStorySnapshots(document, blocks),
            document.FloatingDrawings.Select((drawing, index) => ToFloatingDrawingSnapshot(drawing, index)).ToArray(),
            ToStyleUsages(document),
            ToListUsages(document),
            tables,
            adjacency);
    }

    private static IReadOnlyList<DocxStructureStorySnapshot> ToStorySnapshots(
        DocxDocument document,
        IReadOnlyList<DocxStructureBlockSnapshot> bodyBlocks)
    {
        var stories = new List<DocxStructureStorySnapshot>
        {
            new(
                "Body",
                "document",
                null,
                null,
                bodyBlocks.Count,
                bodyBlocks.Count(block => block.Kind == "Paragraph"),
                bodyBlocks.Count(block => block.Kind == "Table"),
                bodyBlocks.Sum(block => block.TextLength),
                bodyBlocks.Sum(block => block.InlineImageCount))
        };

        AddStaticStories(stories, "Header", "document", null, document.HeaderParagraphsByType);
        AddStaticStories(stories, "Footer", "document", null, document.FooterParagraphsByType);
        for (int blockIndex = 0; blockIndex < document.BodyElements.Count; blockIndex++)
        {
            if (document.BodyElements[blockIndex] is not DocxSectionBreakElement sectionBreak)
            {
                continue;
            }

            string scope = "section@" + blockIndex.ToString(CultureInfo.InvariantCulture);
            AddStaticStories(stories, "Header", scope, blockIndex, sectionBreak.PageSettings.HeaderParagraphsByType);
            AddStaticStories(stories, "Footer", scope, blockIndex, sectionBreak.PageSettings.FooterParagraphsByType);
        }

        return stories;
    }

    private static void AddStaticStories(
        List<DocxStructureStorySnapshot> stories,
        string kind,
        string scope,
        int? sectionBreakBlockIndex,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType)
    {
        foreach (KeyValuePair<string, IReadOnlyList<DocxParagraph>> entry in paragraphsByType.OrderBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase))
        {
            stories.Add(new DocxStructureStorySnapshot(
                kind,
                scope,
                sectionBreakBlockIndex,
                entry.Key,
                0,
                entry.Value.Count,
                0,
                entry.Value.Sum(TextLength),
                entry.Value.Sum(paragraph => paragraph.Images.Count)));
        }
    }

    private static DocxStructureBlockSnapshot FromParagraph(int blockIndex, string? previousKind, string? nextKind, DocxParagraph paragraph)
    {
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "Paragraph",
            previousKind,
            nextKind,
            ParagraphStyleId: paragraph.StyleId,
            RunCount: paragraph.Runs.Count,
            TextLength: TextLength(paragraph),
            HasVisibleText: HasVisibleText(paragraph),
            ListFormatValue: paragraph.ListLabel?.FormatValue,
            InlineImageCount: paragraph.Images.Count,
            SpacingBeforePoints: paragraph.SpacingBeforePoints,
            SpacingAfterPoints: paragraph.SpacingAfterPoints,
            LineSpacingPoints: paragraph.LineSpacingPoints,
            LineSpacingFactor: paragraph.LineSpacingFactor,
            HasBeforeSpacingToken: HasBeforeSpacingToken(paragraph.Spacing),
            HasAfterSpacingToken: HasAfterSpacingToken(paragraph.Spacing),
            ContextualSpacing: paragraph.Spacing.ContextualSpacing,
            KeepNext: paragraph.KeepRules.KeepNext,
            KeepLines: paragraph.KeepRules.KeepLines,
            WidowControl: paragraph.KeepRules.WidowControl,
            ParagraphIndentLeftPoints: paragraph.Indent.LeftPoints,
            ParagraphIndentRightPoints: paragraph.Indent.RightPoints,
            ParagraphIndentFirstLinePoints: paragraph.Indent.FirstLinePoints,
            ParagraphIndentHangingPoints: paragraph.Indent.HangingPoints,
            WhitespaceDelimitedTokenCount: CountWhitespaceDelimitedTokens(paragraph),
            LongestWhitespaceDelimitedTokenLength: LongestWhitespaceDelimitedTokenLength(paragraph),
            SpaceCharacterCount: CountCharacters(paragraph, static c => c == ' '),
            NonAsciiCharacterCount: CountCharacters(paragraph, static c => c > 127),
            PunctuationCharacterCount: CountCharacters(paragraph, char.IsPunctuation),
            DigitCharacterCount: CountCharacters(paragraph, char.IsDigit),
            UppercaseCharacterCount: CountCharacters(paragraph, char.IsUpper),
            LowercaseCharacterCount: CountCharacters(paragraph, char.IsLower),
            TabStopCount: paragraph.TabStops.Count,
            SnapToGrid: paragraph.SnapToGrid,
            SnapToGridValue: paragraph.SnapToGridValue);
    }

    private static DocxStructureBlockSnapshot FromTable(int blockIndex, string? previousKind, string? nextKind, DocxTable table, int tableIndex)
    {
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "Table",
            previousKind,
            nextKind,
            TableRowCount: table.Rows.Count,
            TableIndex: tableIndex,
            TableMaxColumnCount: MaxColumnCount(table),
            TablePreferredWidthPoints: table.PreferredWidthPoints,
            TablePreferredWidthType: table.PreferredWidthType,
            TableIndentPoints: table.IndentPoints,
            TableCellSpacingPoints: table.CellSpacingPoints,
            TableLayoutValue: table.LayoutValue);
    }

    private static DocxStructureBlockSnapshot FromPageBreak(int blockIndex, string? previousKind, string? nextKind, DocxPageBreakElement pageBreak)
    {
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "PageBreak",
            previousKind,
            nextKind,
            PageBreakSourceKind: pageBreak.SourceKind,
            PageBreakValue: pageBreak.Value,
            PageBreakConsumesParagraphLine: pageBreak.BreakParagraph is not null,
            PageBreakLineSpacingPoints: pageBreak.BreakParagraph?.LineSpacingPoints,
            PageBreakLineSpacingFactor: pageBreak.BreakParagraph?.LineSpacingFactor);
    }

    private static DocxStructureBlockSnapshot FromSectionBreak(int blockIndex, string? previousKind, string? nextKind, DocxSectionBreakElement sectionBreak)
    {
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "SectionBreak",
            previousKind,
            nextKind,
            SectionBreakTypeValue: sectionBreak.TypeValue,
            SectionColumnCountValue: sectionBreak.ColumnCountValue,
            SectionColumnSpaceValue: sectionBreak.ColumnSpaceValue);
    }

    private static DocxStructureTableSnapshot ToTableSnapshot(DocxTable table, int tableIndex, int blockIndex)
    {
        int cellCount = table.Rows.Sum(row => row.Cells.Count);
        int paragraphCount = table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Count));
        return new DocxStructureTableSnapshot(
            tableIndex,
            blockIndex,
            table.StyleId,
            table.Rows.Count,
            MaxColumnCount(table),
            table.ColumnWidthsPoints.Count,
            table.ColumnWidthsPoints.Sum(),
            table.HasExplicitGrid,
            table.PreferredWidthPoints,
            table.PreferredWidthValue,
            table.PreferredWidthType,
            table.IndentPoints,
            table.IndentValue,
            table.IndentType,
            table.CellSpacingPoints,
            table.CellSpacingValue,
            table.CellSpacingType,
            table.LayoutValue,
            table.Rows.Count(row => row.IsHeader),
            table.Rows.Count(row => row.CantSplit),
            table.Rows.Count(row => row.HeightPoints is not null),
            table.Rows.Count(row => string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase)),
            table.Rows.Count(row => string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase)),
            table.Rows.Count(row => row.TablePropertyExceptionCellMargins is not null),
            cellCount,
            table.Rows.Sum(row => row.Cells.Count(cell => cell.GridSpan > 1)),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.HasVerticalMerge)),
            table.Rows.Sum(row => row.Cells.Count(cell => string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase))),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.FillHex is not null || cell.ShadingValue is not null)),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.VerticalAlignmentValue is not null)),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.PreferredWidthPoints is not null || cell.PreferredWidthValue is not null)),
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Borders.Count(border => !string.Equals(border.Value, "nil", StringComparison.OrdinalIgnoreCase) && !string.Equals(border.Value, "none", StringComparison.OrdinalIgnoreCase)))),
            paragraphCount,
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Sum(paragraph => paragraph.Runs.Count))),
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Sum(TextLength))),
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Sum(CountWhitespaceDelimitedTokens))),
            table.Rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs).Select(LongestWhitespaceDelimitedTokenLength).DefaultIfEmpty(0).Max(),
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Sum(paragraph => paragraph.Images.Count))),
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Count(paragraph => paragraph.ListLabel is not null))),
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Count(paragraph => paragraph.KeepRules.KeepNext == true || paragraph.KeepRules.KeepLines == true))),
            table.Look?.FirstRow,
            table.Look?.FirstColumn,
            table.Look?.NoHorizontalBand,
            table.Look?.NoVerticalBand,
            table.Rows.Select((row, rowIndex) => ToTableRowSnapshot(row, rowIndex)).ToArray());
    }

    private static DocxStructureTableRowSnapshot ToTableRowSnapshot(DocxTableRow row, int rowIndex)
    {
        DocxStructureTableCellSnapshot[] cells = row.Cells.Select(ToTableCellSnapshot).ToArray();
        return new DocxStructureTableRowSnapshot(
            rowIndex,
            row.Cells.Count,
            row.Cells.Sum(cell => Math.Max(1, cell.GridSpan)),
            row.IsHeader,
            row.HeaderValue,
            row.CantSplit,
            row.CantSplitValue,
            row.HeightPoints,
            row.HeightValue,
            row.HeightRuleValue,
            row.TablePropertyExceptionCellMargins is not null,
            cells.Count(cell => cell.GridSpan > 1),
            cells.Count(cell => cell.HasVerticalMerge),
            cells.Count(cell => string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase)),
            cells.Count(cell => cell.HasShading),
            cells.Count(cell => cell.VerticalAlignmentValue is not null),
            cells.Count(cell => cell.HasPreferredWidth),
            cells.Sum(cell => cell.VisibleBorderCount),
            cells.Sum(cell => cell.ParagraphCount),
            cells.Sum(cell => cell.RunCount),
            cells.Sum(cell => cell.TextLength),
            cells.Sum(cell => cell.WhitespaceDelimitedTokenCount),
            cells.Select(cell => cell.LongestWhitespaceDelimitedTokenLength).DefaultIfEmpty(0).Max(),
            cells.Sum(cell => cell.InlineImageCount),
            cells.Sum(cell => cell.NumberedParagraphCount),
            cells.Sum(cell => cell.KeepRuleParagraphCount),
            cells.Sum(cell => cell.BeforeSpacingTokenParagraphCount),
            cells.Sum(cell => cell.AfterSpacingTokenParagraphCount),
            cells.Select(cell => cell.MaxFontSize).DefaultIfEmpty(0d).Max(),
            cells);
    }

    private static DocxStructureTableCellSnapshot ToTableCellSnapshot(DocxTableCell cell, int cellIndex)
    {
        return new DocxStructureTableCellSnapshot(
            cellIndex,
            Math.Max(1, cell.GridSpan),
            cell.GridSpanValue,
            cell.HasVerticalMerge,
            cell.VerticalMergeValue,
            cell.FillHex is not null || cell.ShadingValue is not null,
            cell.ShadingValue,
            cell.VerticalAlignmentValue,
            cell.PreferredWidthPoints is not null || cell.PreferredWidthValue is not null,
            cell.PreferredWidthPoints,
            cell.PreferredWidthValue,
            cell.PreferredWidthType,
            cell.Borders.Count(border => !string.Equals(border.Value, "nil", StringComparison.OrdinalIgnoreCase) && !string.Equals(border.Value, "none", StringComparison.OrdinalIgnoreCase)),
            cell.Paragraphs.Count,
            cell.Paragraphs.Sum(paragraph => paragraph.Runs.Count),
            cell.Paragraphs.Sum(TextLength),
            cell.Paragraphs.Sum(CountWhitespaceDelimitedTokens),
            cell.Paragraphs.Select(LongestWhitespaceDelimitedTokenLength).DefaultIfEmpty(0).Max(),
            cell.Paragraphs.Sum(paragraph => paragraph.Images.Count),
            cell.Paragraphs.Count(paragraph => paragraph.ListLabel is not null),
            cell.Paragraphs.Count(paragraph => paragraph.KeepRules.KeepNext == true || paragraph.KeepRules.KeepLines == true),
            cell.Paragraphs.Count(paragraph => HasBeforeSpacingToken(paragraph.Spacing)),
            cell.Paragraphs.Count(paragraph => HasAfterSpacingToken(paragraph.Spacing)),
            cell.Paragraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.FontSize).DefaultIfEmpty(0d).Max(),
            CountCharacters(cell.Paragraphs, static c => c == ' '),
            CountCharacters(cell.Paragraphs, static c => c > 127),
            CountCharacters(cell.Paragraphs, char.IsPunctuation),
            CountCharacters(cell.Paragraphs, char.IsDigit),
            CountCharacters(cell.Paragraphs, char.IsUpper),
            CountCharacters(cell.Paragraphs, char.IsLower));
    }

    private static DocxStructureFloatingDrawingSnapshot ToFloatingDrawingSnapshot(DocxFloatingDrawing drawing, int index)
    {
        return new DocxStructureFloatingDrawingSnapshot(
            index,
            drawing.WrapKind,
            drawing.WrapTextValue,
            drawing.BehindDocumentValue,
            drawing.LayoutInCellValue,
            drawing.AllowOverlapValue,
            drawing.HorizontalRelativeFromValue,
            drawing.HorizontalAlignValue,
            drawing.HorizontalOffsetValue,
            drawing.VerticalRelativeFromValue,
            drawing.VerticalAlignValue,
            drawing.VerticalOffsetValue,
            drawing.ExtentCxValue,
            drawing.ExtentCyValue,
            drawing.DistanceTopValue,
            drawing.DistanceBottomValue,
            drawing.DistanceLeftValue,
            drawing.DistanceRightValue,
            drawing.SimplePositionValue,
            drawing.RelativeHeightValue,
            drawing.LockedValue);
    }

    private static IReadOnlyList<DocxStructureStyleUsageSnapshot> ToStyleUsages(DocxDocument document)
    {
        IEnumerable<DocxParagraph> bodyParagraphs = document.BodyElements.OfType<DocxParagraphElement>().Select(element => element.Paragraph);
        IEnumerable<DocxParagraph> tableParagraphs = document.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs);
        IEnumerable<DocxParagraph> staticParagraphs = document.HeaderParagraphsByType.Values
            .Concat(document.FooterParagraphsByType.Values)
            .SelectMany(paragraphs => paragraphs);
        DocxStructureStyleUsageSnapshot[] paragraphStyles = bodyParagraphs
            .Concat(tableParagraphs)
            .Concat(staticParagraphs)
            .GroupBy(paragraph => paragraph.StyleId, StringComparer.Ordinal)
            .Select(group => new DocxStructureStyleUsageSnapshot(
                "Paragraph",
                group.Key,
                group.Count(),
                group.Count(),
                0,
                group.Sum(TextLength)))
            .ToArray();
        DocxStructureStyleUsageSnapshot[] tableStyles = document.Tables
            .GroupBy(table => table.StyleId, StringComparer.Ordinal)
            .Select(group => new DocxStructureStyleUsageSnapshot(
                "Table",
                group.Key,
                group.Count(),
                0,
                group.Count(),
                group.Sum(table => table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Sum(TextLength))))))
            .ToArray();
        return paragraphStyles.Concat(tableStyles).ToArray();
    }

    private static IReadOnlyList<DocxStructureListUsageSnapshot> ToListUsages(DocxDocument document)
    {
        IEnumerable<DocxParagraph> paragraphs = document.BodyElements.OfType<DocxParagraphElement>().Select(element => element.Paragraph)
            .Concat(document.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs));
        return paragraphs
            .Where(paragraph => paragraph.ListLabel is not null)
            .GroupBy(paragraph => new
            {
                paragraph.ListLabel!.NumberId,
                paragraph.ListLabel.Level,
                paragraph.ListLabel.FormatValue,
                paragraph.ListLabel.SuffixValue
            })
            .Select(group => new DocxStructureListUsageSnapshot(
                group.Key.NumberId,
                group.Key.Level,
                group.Key.FormatValue,
                group.Key.SuffixValue,
                group.Count(),
                group.Sum(TextLength)))
            .ToArray();
    }

    private static DocxStructureTableAdjacencySnapshot ToTableAdjacencySnapshot(
        IReadOnlyList<DocxBodyElement> elements,
        DocxTable table,
        int tableIndex,
        int blockIndex,
        string? previousKind,
        string? nextKind)
    {
        DocxParagraph? previousParagraph = TryGetAdjacentParagraph(elements, blockIndex - 1);
        DocxParagraph? nextParagraph = TryGetAdjacentParagraph(elements, blockIndex + 1);
        return new DocxStructureTableAdjacencySnapshot(
            tableIndex,
            blockIndex,
            previousKind,
            nextKind,
            table.Rows.Count,
            MaxColumnCount(table),
            previousParagraph?.StyleId,
            previousParagraph?.SpacingAfterPoints,
            previousParagraph is null ? null : HasAfterSpacingToken(previousParagraph.Spacing),
            nextParagraph?.StyleId,
            nextParagraph?.SpacingBeforePoints,
            nextParagraph is null ? null : HasBeforeSpacingToken(nextParagraph.Spacing),
            nextParagraph is null ? null : TextLength(nextParagraph),
            nextParagraph?.ListLabel is not null,
            nextParagraph?.KeepRules.KeepNext,
            nextParagraph?.KeepRules.KeepLines);
    }

    private static string GetBlockKind(DocxBodyElement element)
    {
        return element switch
        {
            DocxParagraphElement => "Paragraph",
            DocxTableElement => "Table",
            DocxPageBreakElement => "PageBreak",
            DocxSectionBreakElement => "SectionBreak",
            _ => "Unknown"
        };
    }

    private static DocxParagraph? TryGetAdjacentParagraph(IReadOnlyList<DocxBodyElement> elements, int index)
    {
        return index >= 0 &&
            index < elements.Count &&
            elements[index] is DocxParagraphElement paragraph
                ? paragraph.Paragraph
                : null;
    }

    private static bool HasBeforeSpacingToken(DocxParagraphSpacing spacing)
    {
        return spacing.BeforeValue is not null ||
            spacing.BeforeLinesValue is not null ||
            spacing.BeforeAutoSpacingValue is not null;
    }

    private static bool HasAfterSpacingToken(DocxParagraphSpacing spacing)
    {
        return spacing.AfterValue is not null ||
            spacing.AfterLinesValue is not null ||
            spacing.AfterAutoSpacingValue is not null;
    }

    private static int TextLength(DocxParagraph paragraph)
    {
        return paragraph.Runs.Sum(run => run.Text.Length);
    }

    private static int CountCharacters(DocxParagraph paragraph, Func<char, bool> predicate)
    {
        return paragraph.Runs.Sum(run => run.Text.Count(predicate));
    }

    private static int CountWhitespaceDelimitedTokens(DocxParagraph paragraph)
    {
        return EnumerateWhitespaceDelimitedTokenLengths(paragraph).Count();
    }

    private static int LongestWhitespaceDelimitedTokenLength(DocxParagraph paragraph)
    {
        return EnumerateWhitespaceDelimitedTokenLengths(paragraph).DefaultIfEmpty(0).Max();
    }

    private static IEnumerable<int> EnumerateWhitespaceDelimitedTokenLengths(DocxParagraph paragraph)
    {
        int currentLength = 0;
        foreach (DocxTextRun run in paragraph.Runs)
        {
            foreach (char c in run.Text)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (currentLength > 0)
                    {
                        yield return currentLength;
                        currentLength = 0;
                    }

                    continue;
                }

                currentLength++;
            }
        }

        if (currentLength > 0)
        {
            yield return currentLength;
        }
    }

    private static int CountCharacters(IEnumerable<DocxParagraph> paragraphs, Func<char, bool> predicate)
    {
        return paragraphs.Sum(paragraph => CountCharacters(paragraph, predicate));
    }

    private static bool HasVisibleText(DocxParagraph paragraph)
    {
        return paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text));
    }

    private static int MaxColumnCount(DocxTable table)
    {
        return table.Rows.Select(row => row.Cells.Sum(cell => Math.Max(1, cell.GridSpan))).DefaultIfEmpty(0).Max();
    }
}

internal sealed record DocxStructureBlockSnapshot(
    int Index,
    string Kind,
    string? PreviousKind,
    string? NextKind,
    string? ParagraphStyleId = null,
    int RunCount = 0,
    int TextLength = 0,
    bool HasVisibleText = false,
    string? ListFormatValue = null,
    int InlineImageCount = 0,
    int TableRowCount = 0,
    int? TableIndex = null,
    double? SpacingBeforePoints = null,
    double? SpacingAfterPoints = null,
    double? LineSpacingPoints = null,
    double? LineSpacingFactor = null,
    bool HasBeforeSpacingToken = false,
    bool HasAfterSpacingToken = false,
    bool? ContextualSpacing = null,
    bool? KeepNext = null,
    bool? KeepLines = null,
    bool? WidowControl = null,
    int? TableMaxColumnCount = null,
    double? TablePreferredWidthPoints = null,
    string? TablePreferredWidthType = null,
    double? TableIndentPoints = null,
    double? TableCellSpacingPoints = null,
    string? TableLayoutValue = null,
    string? PageBreakSourceKind = null,
    string? PageBreakValue = null,
    string? SectionBreakTypeValue = null,
    string? SectionColumnCountValue = null,
    string? SectionColumnSpaceValue = null,
    double? ParagraphIndentLeftPoints = null,
    double? ParagraphIndentRightPoints = null,
    double? ParagraphIndentFirstLinePoints = null,
    double? ParagraphIndentHangingPoints = null,
    int? WhitespaceDelimitedTokenCount = null,
    int? LongestWhitespaceDelimitedTokenLength = null,
    int? SpaceCharacterCount = null,
    int? NonAsciiCharacterCount = null,
    int? PunctuationCharacterCount = null,
    int? DigitCharacterCount = null,
    int? UppercaseCharacterCount = null,
    int? LowercaseCharacterCount = null,
    int? TabStopCount = null,
    bool? SnapToGrid = null,
    string? SnapToGridValue = null,
    bool? PageBreakConsumesParagraphLine = null,
    double? PageBreakLineSpacingPoints = null,
    double? PageBreakLineSpacingFactor = null);

internal sealed record DocxStructureStorySnapshot(
    string Kind,
    string Scope,
    int? SectionBreakBlockIndex,
    string? VariantType,
    int BlockCount,
    int ParagraphCount,
    int TableCount,
    int TextLength,
    int InlineImageCount);

internal sealed record DocxStructureFloatingDrawingSnapshot(
    int Index,
    string? WrapKind,
    string? WrapTextValue,
    string? BehindDocumentValue,
    string? LayoutInCellValue,
    string? AllowOverlapValue,
    string? HorizontalRelativeFromValue,
    string? HorizontalAlignValue,
    string? HorizontalOffsetValue,
    string? VerticalRelativeFromValue,
    string? VerticalAlignValue,
    string? VerticalOffsetValue,
    string? ExtentCxValue,
    string? ExtentCyValue,
    string? DistanceTopValue,
    string? DistanceBottomValue,
    string? DistanceLeftValue,
    string? DistanceRightValue,
    string? SimplePositionValue,
    string? RelativeHeightValue,
    string? LockedValue);

internal sealed record DocxStructureStyleUsageSnapshot(
    string Kind,
    string? StyleId,
    int Count,
    int ParagraphCount,
    int TableCount,
    int TextLength);

internal sealed record DocxStructureListUsageSnapshot(
    string NumberId,
    int Level,
    string FormatValue,
    string SuffixValue,
    int ParagraphCount,
    int TextLength);

internal sealed record DocxStructureTableSnapshot(
    int TableIndex,
    int BlockIndex,
    string? StyleId,
    int RowCount,
    int MaxColumnCount,
    int GridColumnCount,
    double GridColumnsWidthSum,
    bool HasExplicitGrid,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    double? IndentPoints,
    string? IndentValue,
    string? IndentType,
    double? CellSpacingPoints,
    string? CellSpacingValue,
    string? CellSpacingType,
    string? LayoutValue,
    int HeaderRowCount,
    int CantSplitRowCount,
    int DeclaredHeightRowCount,
    int ExactHeightRowCount,
    int AtLeastHeightRowCount,
    int RowPropertyExceptionCount,
    int CellCount,
    int GridSpanCellCount,
    int VerticalMergeCellCount,
    int VerticalMergeRestartCellCount,
    int ShadedCellCount,
    int VerticalAlignmentCellCount,
    int PreferredWidthCellCount,
    int VisibleBorderCount,
    int ParagraphCount,
    int RunCount,
    int TextLength,
    int WhitespaceDelimitedTokenCount,
    int LongestWhitespaceDelimitedTokenLength,
    int InlineImageCount,
    int NumberedParagraphCount,
    int KeepRuleParagraphCount,
    bool? LookFirstRow,
    bool? LookFirstColumn,
    bool? LookNoHorizontalBand,
    bool? LookNoVerticalBand,
    IReadOnlyList<DocxStructureTableRowSnapshot> Rows);

internal sealed record DocxStructureTableRowSnapshot(
    int RowIndex,
    int CellCount,
    int LogicalGridSpan,
    bool IsHeader,
    string? HeaderValue,
    bool CantSplit,
    string? CantSplitValue,
    double? HeightPoints,
    string? HeightValue,
    string? HeightRuleValue,
    bool HasTablePropertyExceptionCellMargins,
    int GridSpanCellCount,
    int VerticalMergeCellCount,
    int VerticalMergeRestartCellCount,
    int ShadedCellCount,
    int VerticalAlignmentCellCount,
    int PreferredWidthCellCount,
    int VisibleBorderCount,
    int ParagraphCount,
    int RunCount,
    int TextLength,
    int WhitespaceDelimitedTokenCount,
    int LongestWhitespaceDelimitedTokenLength,
    int InlineImageCount,
    int NumberedParagraphCount,
    int KeepRuleParagraphCount,
    int BeforeSpacingTokenParagraphCount,
    int AfterSpacingTokenParagraphCount,
    double MaxFontSize,
    IReadOnlyList<DocxStructureTableCellSnapshot> Cells);

internal sealed record DocxStructureTableCellSnapshot(
    int CellIndex,
    int GridSpan,
    string? GridSpanValue,
    bool HasVerticalMerge,
    string? VerticalMergeValue,
    bool HasShading,
    string? ShadingValue,
    string? VerticalAlignmentValue,
    bool HasPreferredWidth,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    int VisibleBorderCount,
    int ParagraphCount,
    int RunCount,
    int TextLength,
    int WhitespaceDelimitedTokenCount,
    int LongestWhitespaceDelimitedTokenLength,
    int InlineImageCount,
    int NumberedParagraphCount,
    int KeepRuleParagraphCount,
    int BeforeSpacingTokenParagraphCount,
    int AfterSpacingTokenParagraphCount,
    double MaxFontSize,
    int SpaceCharacterCount,
    int NonAsciiCharacterCount,
    int PunctuationCharacterCount,
    int DigitCharacterCount,
    int UppercaseCharacterCount,
    int LowercaseCharacterCount);

internal sealed record DocxStructureTableAdjacencySnapshot(
    int TableIndex,
    int BlockIndex,
    string? PreviousKind,
    string? NextKind,
    int RowCount,
    int MaxColumnCount,
    string? PreviousParagraphStyleId,
    double? PreviousParagraphSpacingAfterPoints,
    bool? PreviousParagraphHasAfterSpacingToken,
    string? NextParagraphStyleId,
    double? NextParagraphSpacingBeforePoints,
    bool? NextParagraphHasBeforeSpacingToken,
    int? NextParagraphTextLength,
    bool? NextParagraphHasListLabel,
    bool? NextParagraphKeepNext,
    bool? NextParagraphKeepLines);
