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

internal sealed record DocxParagraph(
    IReadOnlyList<DocxTextRun> Runs,
    IReadOnlyList<DocxInlineImage> Images,
    DocxTextAlignment Alignment,
    string? AlignmentValue,
    double SpacingBeforePoints,
    double SpacingAfterPoints,
    double LineSpacingFactor,
    double? LineSpacingPoints,
    DocxListLabel? ListLabel);

internal sealed record DocxListLabel(
    string Text,
    string FormatValue,
    string LevelTextValue,
    string NumberId,
    int Level);

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

internal sealed record DocxTable(string? LayoutValue, IReadOnlyList<double> ColumnWidthsPoints, IReadOnlyList<DocxTableRow> Rows);

internal sealed record DocxTableRow(IReadOnlyList<DocxTableCell> Cells, double? HeightPoints);

internal sealed record DocxTableCell(
    string Text,
    string? FillHex,
    string? ShadingValue,
    string? ShadingColor,
    string? VerticalAlignmentValue,
    IReadOnlyList<DocxTableCellBorder> Borders);

internal sealed record DocxTableCellBorder(string Edge, string? Value, string? Color, string? SizeValue);

internal enum DocxTextAlignment
{
    Left,
    Center,
    Right
}
