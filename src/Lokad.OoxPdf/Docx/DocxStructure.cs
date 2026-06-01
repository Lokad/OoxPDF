namespace Lokad.OoxPdf.Docx;

internal sealed record DocxStructureSnapshot(
    int BlockCount,
    int ParagraphBlockCount,
    int TableBlockCount,
    int PageBreakBlockCount,
    int SectionBreakBlockCount,
    int BodyTextLength,
    int InlineImageCount,
    IReadOnlyList<DocxStructureBlockSnapshot> Blocks,
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
            blocks,
            tables,
            adjacency);
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
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Sum(paragraph => paragraph.Images.Count))),
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Count(paragraph => paragraph.ListLabel is not null))),
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Paragraphs.Count(paragraph => paragraph.KeepRules.KeepNext == true || paragraph.KeepRules.KeepLines == true))),
            table.Look?.FirstRow,
            table.Look?.FirstColumn,
            table.Look?.NoHorizontalBand,
            table.Look?.NoVerticalBand);
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
    int InlineImageCount,
    int NumberedParagraphCount,
    int KeepRuleParagraphCount,
    bool? LookFirstRow,
    bool? LookFirstColumn,
    bool? LookNoHorizontalBand,
    bool? LookNoVerticalBand);

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
