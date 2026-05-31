namespace Lokad.OoxPdf.Docx;

internal sealed record DocxDocument(
    double PageWidthPoints,
    double PageHeightPoints,
    double MarginLeftPoints,
    double MarginRightPoints,
    double MarginTopPoints,
    double MarginBottomPoints,
    DocxPageSettings PageSettings,
    IReadOnlyList<DocxFloatingDrawing> FloatingDrawings,
    IReadOnlyList<DocxParagraph> HeaderParagraphs,
    IReadOnlyList<DocxParagraph> FooterParagraphs,
    IReadOnlyList<DocxBodyElement> BodyElements,
    IReadOnlyList<DocxParagraph> Paragraphs,
    IReadOnlyList<DocxTable> Tables)
{
    public DocxDocument(double pageWidthPoints, double pageHeightPoints)
        : this(pageWidthPoints, pageHeightPoints, 72d, 72d, 72d, 72d, DocxPageSettings.Empty, [], [], [], [], [], [])
    {
    }
}

internal sealed record DocxPageSettings(
    string? WidthValue,
    string? HeightValue,
    string? OrientationValue,
    string? MarginTopValue,
    string? MarginRightValue,
    string? MarginBottomValue,
    string? MarginLeftValue)
{
    public static DocxPageSettings Empty { get; } = new(null, null, null, null, null, null, null);
}

internal sealed record DocxFloatingDrawing(
    string? DistanceTopValue,
    string? DistanceBottomValue,
    string? DistanceLeftValue,
    string? DistanceRightValue,
    string? SimplePositionValue,
    string? RelativeHeightValue,
    string? BehindDocumentValue,
    string? LockedValue,
    string? LayoutInCellValue,
    string? AllowOverlapValue,
    string? ExtentCxValue,
    string? ExtentCyValue,
    string? HorizontalRelativeFromValue,
    string? HorizontalAlignValue,
    string? HorizontalOffsetValue,
    string? VerticalRelativeFromValue,
    string? VerticalAlignValue,
    string? VerticalOffsetValue,
    string? WrapKind,
    string? WrapTextValue);

internal abstract record DocxBodyElement;

internal sealed record DocxParagraphElement(DocxParagraph Paragraph) : DocxBodyElement;

internal sealed record DocxTableElement(DocxTable Table) : DocxBodyElement;

internal sealed record DocxPageBreakElement(string SourceKind, string? Value) : DocxBodyElement;

internal sealed record DocxSectionBreakElement(
    DocxPageSettings PageSettings,
    string? TypeValue,
    string? ColumnCountValue,
    string? ColumnEqualWidthValue,
    string? ColumnSpaceValue) : DocxBodyElement;

internal sealed record DocxParagraph(
    IReadOnlyList<DocxTextRun> Runs,
    IReadOnlyList<DocxInlineImage> Images,
    string? StyleId,
    DocxTextAlignment Alignment,
    string? AlignmentValue,
    double SpacingBeforePoints,
    double SpacingAfterPoints,
    double LineSpacingFactor,
    double? LineSpacingPoints,
    DocxParagraphSpacing Spacing,
    DocxParagraphKeepRules KeepRules,
    DocxListLabel? ListLabel);

internal sealed record DocxParagraphSpacing(
    string? BeforeValue,
    string? AfterValue,
    string? BeforeLinesValue,
    string? AfterLinesValue,
    string? BeforeAutoSpacingValue,
    string? AfterAutoSpacingValue,
    string? LineValue,
    string? LineRuleValue,
    bool? ContextualSpacing)
{
    public static DocxParagraphSpacing Empty { get; } = new(null, null, null, null, null, null, null, null, null);
}

internal sealed record DocxParagraphKeepRules(
    bool? KeepNext,
    string? KeepNextValue,
    bool? KeepLines,
    string? KeepLinesValue,
    bool? WidowControl,
    string? WidowControlValue)
{
    public static DocxParagraphKeepRules Empty { get; } = new(null, null, null, null, null, null);
}

internal sealed record DocxListLabel(
    string Text,
    string FormatValue,
    string LevelTextValue,
    string SuffixValue,
    string NumberId,
    int Level,
    DocxNumberingIndent Indent);

internal sealed record DocxNumberingIndent(
    double? LeftPoints,
    double? RightPoints,
    double? FirstLinePoints,
    double? HangingPoints,
    string? LeftValue,
    string? RightValue,
    string? FirstLineValue,
    string? HangingValue)
{
    public static DocxNumberingIndent Empty { get; } = new(null, null, null, null, null, null, null, null);
}

internal sealed record DocxTextRun(
    string Text,
    double FontSize,
    string? ColorHex,
    bool Bold,
    bool Italic,
    bool Underline,
    string? UnderlineValue,
    string? FontFamily);

internal sealed record DocxInlineImage(double WidthPoints, double HeightPoints, string ContentType, byte[] Bytes, string? PartName);

internal sealed record DocxTable(
    string? LayoutValue,
    IReadOnlyList<double> ColumnWidthsPoints,
    IReadOnlyList<DocxTableRow> Rows,
    string? StyleId = null,
    double? PreferredWidthPoints = null,
    string? PreferredWidthValue = null,
    string? PreferredWidthType = null,
    double? IndentPoints = null,
    string? IndentValue = null,
    string? IndentType = null,
    DocxTableLook? Look = null);

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
    bool IsHeader = false,
    string? HeaderValue = null);

internal sealed record DocxTableCell(
    string Text,
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
    string? GridSpanValue = null);

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
}

internal enum DocxTextAlignment
{
    Left,
    Center,
    Right
}
