using System.Globalization;

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
    public DocxTextAlignment Alignment { get; init; }

    public DocxTableWidthKind PreferredWidthKind => ParseTableWidthKind(PreferredWidthType, PreferredWidthValue);

    public static DocxTableWidthKind ParseTableWidthKind(string? type, string? value = null)
    {
        // A percent literal carries its unit even when the declared type differs.
        if (string.Equals(type, "pct", StringComparison.OrdinalIgnoreCase) ||
            TryReadPercentWidthLiteral(value.AsSpan().Trim(), out _))
        {
            return DocxTableWidthKind.Percent;
        }

        return string.Equals(type, "dxa", StringComparison.OrdinalIgnoreCase)
            ? DocxTableWidthKind.Dxa : DocxTableWidthKind.Auto;
    }

    public static bool TryReadPercentWidth(string? value, out int fiftiethsOfPercent)
    {
        ReadOnlySpan<char> text = value.AsSpan().Trim();
        if (text.IsEmpty || text[^1] != '%')
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out fiftiethsOfPercent);

        return TryReadPercentWidthLiteral(text, out fiftiethsOfPercent);
    }

    private static bool TryReadPercentWidthLiteral(ReadOnlySpan<char> text, out int fiftiethsOfPercent)
    {
        if (!text.IsEmpty && text[^1] == '%' &&
            decimal.TryParse(text[..^1], NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
            CultureInfo.InvariantCulture, out decimal percent) &&
            percent >= int.MinValue / 50m && percent <= int.MaxValue / 50m)
        {
            fiftiethsOfPercent = (int)decimal.Round(percent * 50m, 0, MidpointRounding.AwayFromZero);
            return true;
        }

        fiftiethsOfPercent = 0;
        return false;
    }
}

// Keep the authored width and type for inspection alongside the effective unit.
internal enum DocxTableWidthKind
{
    Auto,
    Dxa,
    Percent
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

    public int GridBefore { get; init; }
    public IReadOnlyList<DocxBookmarkAnchor>? BookmarkAnchors { get; init; }

    public bool HasSupportedHeight => HeightRuleValue is null or "auto" or "exact" or "atLeast" &&
        (HeightPoints is { } height ? double.IsFinite(height) && height >= 0d :
            HeightValue is null && HeightRuleValue is null or "auto");
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
    IReadOnlyList<DocxParagraph> Paragraphs,
    string? FillHex,
    string? ShadingValue,
    string? ShadingColor,
    string? VerticalAlignmentValue,
    IReadOnlyList<DocxTableCellBorder> Borders,
    DocxTableCellMargins Margins,
    double? PreferredWidthPoints = null,
    string? PreferredWidthValue = null,
    string? PreferredWidthType = null,
    int GridSpan = 1,
    string? GridSpanValue = null,
    DocxTableCellConditionalFormat? ConditionalFormat = null,
    bool HasVerticalMerge = false,
    string? VerticalMergeValue = null,
    bool NoWrap = false,
    string? NoWrapValue = null,
    bool FitText = false,
    string? FitTextValue = null,
    string? TextDirectionValue = null,
    // Resolved style origin: zero anchors the grid; null retains the existing
    // unstyled/direct-margin policy. Positive values retain TableNormal anchoring.
    DocxTableCellMargins? GridAnchorMargins = null,
    bool PinTextToMargin = false)
{

    public IReadOnlyList<DocxBodyElement> BodyElements { get; init; } = [];
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
    public bool IsVerticalMergeRestart => HasVerticalMerge &&
        string.Equals(VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase);

    public DocxTableCellVerticalAlignment VerticalAlignment => ParseVerticalAlignment(VerticalAlignmentValue);

    public DocxTableWidthKind PreferredWidthKind => DocxTable.ParseTableWidthKind(PreferredWidthType, PreferredWidthValue);

    public bool HasHorizontalTextDirection => IsHorizontalTextDirection(TextDirectionValue);

    // ISO/IEC 29500 defines the Transitional lrTb spelling as equivalent to tb.
    public static bool IsHorizontalTextDirection(string? value) => value is null or "lrTb" or "tb";

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
    bool? LastRow,
    bool? FirstColumn,
    bool? LastColumn,
    bool? OddHorizontalBand,
    bool? EvenHorizontalBand,
    bool? OddVerticalBand,
    bool? EvenVerticalBand,
    bool? FirstRowFirstColumn,
    bool? FirstRowLastColumn,
    bool? LastRowFirstColumn,
    bool? LastRowLastColumn)
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

internal sealed record DocxTableCellBorder(string Edge, string? Value, string? Color, string? SizeValue)
{
    public bool UseDoubleBandWidth { get; init; }
}

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
