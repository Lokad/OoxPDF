namespace Lokad.OoxPdf.Docx;

internal sealed record DocxTable(
    string? LayoutValue,
    IReadOnlyList<double> ColumnWidthsPoints,
    IReadOnlyList<DocxTableRow> Rows,
    string? StyleId,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    double? IndentPoints,
    string? IndentValue,
    string? IndentType,
    double? CellSpacingPoints,
    string? CellSpacingValue,
    string? CellSpacingType,
    DocxTableLook? Look,
    bool HasExplicitGrid,
    bool UseLegacyTableGrid = false)
{

    public DocxTable(
        string? LayoutValue,
        IReadOnlyList<double> ColumnWidthsPoints,
        IReadOnlyList<DocxTableRow> Rows)
        : this(LayoutValue, ColumnWidthsPoints, Rows, null, null, null, null, null, null, null, null, null, null, null, true, false)
    {
    }
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal sealed record DocxTableLook(
    string? Value,
    bool? FirstRow,
    string? FirstRowValue,
    bool? LastRow,
    string? LastRowValue,
    bool? FirstColumn,
    string? FirstColumnValue,
    bool? LastColumn,
    string? LastColumnValue,
    bool? NoHorizontalBand,
    string? NoHorizontalBandValue,
    bool? NoVerticalBand,
    string? NoVerticalBandValue)
{
    public static DocxTableLook Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null, null);
}

internal sealed record DocxTableRow(
    IReadOnlyList<DocxTableCell> Cells,
    double? HeightPoints,
    bool IsHeader,
    string? HeaderValue,
    string? HeightValue,
    string? HeightRuleValue,
    DocxTableCellMargins? TablePropertyExceptionCellMargins,
    bool CantSplit,
    string? CantSplitValue)
{

    public DocxTableRow(
        IReadOnlyList<DocxTableCell> Cells,
        double? HeightPoints)
        : this(Cells, HeightPoints, false, null, null, null, null, false, null)
    {
    }
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

// R17: parsed cell vertical alignment. The raw w:val spelling stays on
// VerticalAlignmentValue as provenance (snapshots and inspection read it); execution
// switches on the parsed enum instead of re-comparing strings per layout. Unknown
// spellings (including "both") fall back to Top, matching the legacy comparisons,
// which only special-cased bottom and center.
internal enum DocxTableCellVerticalAlignment
{
    Top,
    Center,
    Bottom
}

internal sealed record DocxTableCell(
    string Text,
    IReadOnlyList<DocxParagraph> Paragraphs,
    string? FillHex,
    string? ShadingValue,
    string? ShadingColor,
    string? VerticalAlignmentValue,
    IReadOnlyList<DocxTableCellBorder> Borders,
    DocxTableCellMargins Margins,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    int GridSpan,
    string? GridSpanValue,
    DocxTableCellConditionalFormat? ConditionalFormat,
    bool HasVerticalMerge,
    string? VerticalMergeValue,
    bool NoWrap,
    string? NoWrapValue,
    bool FitText,
    string? FitTextValue,
    string? TextDirectionValue,
    DocxTableCellMargins? StyleMargins = null,
    bool PinTextToMargin = false)
{

    public DocxTableCell(
        string Text,
        IReadOnlyList<DocxParagraph> Paragraphs,
        string? FillHex,
        string? ShadingValue,
        string? ShadingColor,
        string? VerticalAlignmentValue,
        IReadOnlyList<DocxTableCellBorder> Borders,
        DocxTableCellMargins Margins)
        : this(Text, Paragraphs, FillHex, ShadingValue, ShadingColor, VerticalAlignmentValue, Borders, Margins, null, null, null, 1, null, null, false, null, false, null, false, null, null, null, false)
    {
    }
    public IReadOnlyList<DocxBodyElement> BodyElements { get; init; } = [];
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];

    public DocxTableCellVerticalAlignment VerticalAlignment => ParseVerticalAlignment(VerticalAlignmentValue);

    public static DocxTableCellVerticalAlignment ParseVerticalAlignment(string? value)
    {
        if (string.Equals(value, "bottom", StringComparison.OrdinalIgnoreCase))
        {
            return DocxTableCellVerticalAlignment.Bottom;
        }

        if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase))
        {
            return DocxTableCellVerticalAlignment.Center;
        }

        return DocxTableCellVerticalAlignment.Top;
    }
}

internal sealed record DocxTableCellConditionalFormat(
    string? Value,
    bool? FirstRow,
    string? FirstRowValue,
    bool? LastRow,
    string? LastRowValue,
    bool? FirstColumn,
    string? FirstColumnValue,
    bool? LastColumn,
    string? LastColumnValue,
    bool? OddHorizontalBand,
    string? OddHorizontalBandValue,
    bool? EvenHorizontalBand,
    string? EvenHorizontalBandValue,
    bool? OddVerticalBand,
    string? OddVerticalBandValue,
    bool? EvenVerticalBand,
    string? EvenVerticalBandValue,
    bool? FirstRowFirstColumn,
    string? FirstRowFirstColumnValue,
    bool? FirstRowLastColumn,
    string? FirstRowLastColumnValue,
    bool? LastRowFirstColumn,
    string? LastRowFirstColumnValue,
    bool? LastRowLastColumn,
    string? LastRowLastColumnValue)
{
    public bool IsDefined =>
        FirstRow is not null ||
        LastRow is not null ||
        FirstColumn is not null ||
        LastColumn is not null ||
        OddHorizontalBand is not null ||
        EvenHorizontalBand is not null ||
        OddVerticalBand is not null ||
        EvenVerticalBand is not null ||
        FirstRowFirstColumn is not null ||
        FirstRowLastColumn is not null ||
        LastRowFirstColumn is not null ||
        LastRowLastColumn is not null;
}

internal sealed record DocxTableCellBorder(string Edge, string? Value, string? Color, string? SizeValue);

internal sealed record DocxTableCellMargins(
    double? TopPoints,
    double? RightPoints,
    double? BottomPoints,
    double? LeftPoints,
    string? TopValue,
    string? RightValue,
    string? BottomValue,
    string? LeftValue)
{
    public static DocxTableCellMargins Empty { get; } = new(null, null, null, null, null, null, null, null);

    public DocxTableCellMargins Merge(DocxTableCellMargins inherited)
    {
        return new DocxTableCellMargins(
            TopPoints ?? inherited.TopPoints,
            RightPoints ?? inherited.RightPoints,
            BottomPoints ?? inherited.BottomPoints,
            LeftPoints ?? inherited.LeftPoints,
            TopValue ?? inherited.TopValue,
            RightValue ?? inherited.RightValue,
            BottomValue ?? inherited.BottomValue,
            LeftValue ?? inherited.LeftValue);
    }
}
