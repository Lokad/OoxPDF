using System.Text.Encodings.Web;
using System.Text.Json;

using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.DocxInspect <input.docx> <output-directory>");
    Environment.Exit(2);
}

string inputPath = Path.GetFullPath(args[0]);
string outputDirectory = Path.GetFullPath(args[1]);
Directory.CreateDirectory(outputDirectory);

using FileStream stream = File.OpenRead(inputPath);
OoxPackage package = OoxPackage.Open(stream);
DocxDocument document = new DocxReader().Read(package);
var renderer = new DocxRenderer();

var options = new JsonSerializerOptions
{
    WriteIndented = true,
    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};

DocxLayoutSnapshot layout = renderer.InspectLayout(document);
DocxFontPlanSnapshot fontPlan = renderer.InspectFontPlan(document);
File.WriteAllText(
    Path.Combine(outputDirectory, "layout-snapshot.json"),
    JsonSerializer.Serialize(layout, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "font-plan-snapshot.json"),
    JsonSerializer.Serialize(fontPlan, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "document-settings.json"),
    JsonSerializer.Serialize(document.Settings, options));
File.WriteAllText(
    Path.Combine(outputDirectory, "page-summary.json"),
    JsonSerializer.Serialize(layout.Pages.Select((page, index) => new
    {
        Page = index + 1,
        page.Width,
        page.Height,
        page.ItemCount,
        page.TextLineCount,
        page.InlineImageCount,
        page.TableRowCount
    }), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "block-sequence.json"),
    JsonSerializer.Serialize(ToBlockSequence(document), options));
File.WriteAllText(
    Path.Combine(outputDirectory, "table-adjacency-summary.json"),
    JsonSerializer.Serialize(ToTableAdjacencySummary(document), options));

static IReadOnlyList<DocxBlockSequenceEntry> ToBlockSequence(DocxDocument document)
{
    var entries = new List<DocxBlockSequenceEntry>();
    int tableIndex = 0;
    for (int index = 0; index < document.BodyElements.Count; index++)
    {
        DocxBodyElement element = document.BodyElements[index];
        string? previousKind = index == 0 ? null : GetBlockKind(document.BodyElements[index - 1]);
        string? nextKind = index + 1 >= document.BodyElements.Count ? null : GetBlockKind(document.BodyElements[index + 1]);
        entries.Add(element switch
        {
            DocxParagraphElement paragraph => FromParagraph(index, previousKind, nextKind, paragraph.Paragraph),
            DocxTableElement table => FromTable(index, previousKind, nextKind, table.Table, tableIndex++),
            DocxPageBreakElement pageBreak => new DocxBlockSequenceEntry(
                index,
                "PageBreak",
                previousKind,
                nextKind,
                null,
                0,
                0,
                false,
                null,
                0,
                0,
                null,
                null,
                null,
                null,
                false,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                pageBreak.SourceKind,
                pageBreak.Value,
                null,
                null,
                null),
            DocxSectionBreakElement sectionBreak => new DocxBlockSequenceEntry(
                index,
                "SectionBreak",
                previousKind,
                nextKind,
                null,
                0,
                0,
                false,
                null,
                0,
                0,
                null,
                null,
                null,
                null,
                false,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                sectionBreak.TypeValue,
                sectionBreak.ColumnCountValue,
                sectionBreak.ColumnSpaceValue),
            _ => new DocxBlockSequenceEntry(
                index,
                "Unknown",
                previousKind,
                nextKind,
                null,
                0,
                0,
                false,
                null,
                0,
                0,
                null,
                null,
                null,
                null,
                false,
                false,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null)
        });
    }

    return entries;
}

static IReadOnlyList<DocxTableAdjacencyEntry> ToTableAdjacencySummary(DocxDocument document)
{
    var entries = new List<DocxTableAdjacencyEntry>();
    int tableIndex = 0;
    for (int index = 0; index < document.BodyElements.Count; index++)
    {
        if (document.BodyElements[index] is not DocxTableElement tableElement)
        {
            continue;
        }

        DocxParagraph? previousParagraph = TryGetAdjacentParagraph(document.BodyElements, index - 1);
        DocxParagraph? nextParagraph = TryGetAdjacentParagraph(document.BodyElements, index + 1);
        string? previousKind = index == 0 ? null : GetBlockKind(document.BodyElements[index - 1]);
        string? nextKind = index + 1 >= document.BodyElements.Count ? null : GetBlockKind(document.BodyElements[index + 1]);
        entries.Add(new DocxTableAdjacencyEntry(
            tableIndex++,
            index,
            previousKind,
            nextKind,
            tableElement.Table.Rows.Count,
            MaxColumnCount(tableElement.Table),
            previousParagraph?.StyleId,
            previousParagraph?.SpacingAfterPoints,
            previousParagraph is null ? null : HasAfterSpacingToken(previousParagraph.Spacing),
            nextParagraph?.StyleId,
            nextParagraph?.SpacingBeforePoints,
            nextParagraph is null ? null : HasBeforeSpacingToken(nextParagraph.Spacing),
            nextParagraph is null ? null : TextLength(nextParagraph),
            nextParagraph?.ListLabel is not null,
            nextParagraph?.KeepRules.KeepNext,
            nextParagraph?.KeepRules.KeepLines));
    }

    return entries;
}

static DocxBlockSequenceEntry FromParagraph(int index, string? previousKind, string? nextKind, DocxParagraph paragraph)
{
    return new DocxBlockSequenceEntry(
        index,
        "Paragraph",
        previousKind,
        nextKind,
        paragraph.StyleId,
        paragraph.Runs.Count,
        TextLength(paragraph),
        HasVisibleText(paragraph),
        paragraph.ListLabel?.FormatValue,
        paragraph.Images.Count,
        0,
        null,
        paragraph.SpacingBeforePoints,
        paragraph.SpacingAfterPoints,
        paragraph.LineSpacingPoints,
        HasBeforeSpacingToken(paragraph.Spacing),
        HasAfterSpacingToken(paragraph.Spacing),
        paragraph.Spacing.ContextualSpacing,
        paragraph.KeepRules.KeepNext,
        paragraph.KeepRules.KeepLines,
        paragraph.KeepRules.WidowControl,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        paragraph.Indent?.LeftPoints,
        paragraph.Indent?.RightPoints,
        paragraph.Indent?.FirstLinePoints,
        paragraph.Indent?.HangingPoints,
        CountCharacters(paragraph, static c => c == ' '),
        CountCharacters(paragraph, static c => c > 127),
        CountCharacters(paragraph, char.IsPunctuation),
        CountCharacters(paragraph, char.IsDigit),
        CountCharacters(paragraph, char.IsUpper),
        CountCharacters(paragraph, char.IsLower));
}

static DocxBlockSequenceEntry FromTable(int index, string? previousKind, string? nextKind, DocxTable table, int tableIndex)
{
    return new DocxBlockSequenceEntry(
        index,
        "Table",
        previousKind,
        nextKind,
        null,
        0,
        0,
        false,
        null,
        0,
        table.Rows.Count,
        tableIndex,
        null,
        null,
        null,
        false,
        false,
        null,
        null,
        null,
        null,
        MaxColumnCount(table),
        table.PreferredWidthPoints,
        table.PreferredWidthType,
        table.IndentPoints,
        table.CellSpacingPoints,
        table.LayoutValue,
        null,
        null,
        null,
        null,
        null);
}

static string GetBlockKind(DocxBodyElement element)
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

static DocxParagraph? TryGetAdjacentParagraph(IReadOnlyList<DocxBodyElement> elements, int index)
{
    if (index < 0 || index >= elements.Count)
    {
        return null;
    }

    return elements[index] is DocxParagraphElement paragraph ? paragraph.Paragraph : null;
}

static bool HasBeforeSpacingToken(DocxParagraphSpacing spacing)
{
    return spacing.BeforeValue is not null ||
        spacing.BeforeLinesValue is not null ||
        spacing.BeforeAutoSpacingValue is not null;
}

static bool HasAfterSpacingToken(DocxParagraphSpacing spacing)
{
    return spacing.AfterValue is not null ||
        spacing.AfterLinesValue is not null ||
        spacing.AfterAutoSpacingValue is not null;
}

static int TextLength(DocxParagraph paragraph)
{
    return paragraph.Runs.Sum(run => run.Text.Length);
}

static int CountCharacters(DocxParagraph paragraph, Func<char, bool> predicate)
{
    return paragraph.Runs.Sum(run => run.Text.Count(predicate));
}

static bool HasVisibleText(DocxParagraph paragraph)
{
    return paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text));
}

static int MaxColumnCount(DocxTable table)
{
    return table.Rows.Select(row => row.Cells.Sum(cell => Math.Max(1, cell.GridSpan))).DefaultIfEmpty(0).Max();
}

internal sealed record DocxBlockSequenceEntry(
    int Index,
    string Kind,
    string? PreviousKind,
    string? NextKind,
    string? ParagraphStyleId,
    int RunCount,
    int TextLength,
    bool HasVisibleText,
    string? ListFormatValue,
    int InlineImageCount,
    int TableRowCount,
    int? TableIndex,
    double? SpacingBeforePoints,
    double? SpacingAfterPoints,
    double? LineSpacingPoints,
    bool HasBeforeSpacingToken,
    bool HasAfterSpacingToken,
    bool? ContextualSpacing,
    bool? KeepNext,
    bool? KeepLines,
    bool? WidowControl,
    int? TableMaxColumnCount,
    double? TablePreferredWidthPoints,
    string? TablePreferredWidthType,
    double? TableIndentPoints,
    double? TableCellSpacingPoints,
    string? TableLayoutValue,
    string? PageBreakSourceKind,
    string? PageBreakValue,
    string? SectionBreakTypeValue,
    string? SectionColumnCountValue,
    string? SectionColumnSpaceValue,
    double? ParagraphIndentLeftPoints = null,
    double? ParagraphIndentRightPoints = null,
    double? ParagraphIndentFirstLinePoints = null,
    double? ParagraphIndentHangingPoints = null,
    int? SpaceCharacterCount = null,
    int? NonAsciiCharacterCount = null,
    int? PunctuationCharacterCount = null,
    int? DigitCharacterCount = null,
    int? UppercaseCharacterCount = null,
    int? LowercaseCharacterCount = null);

internal sealed record DocxTableAdjacencyEntry(
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
