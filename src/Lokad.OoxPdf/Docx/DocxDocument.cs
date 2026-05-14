namespace Lokad.OoxPdf.Docx;

internal sealed record DocxDocument(
    double PageWidthPoints,
    double PageHeightPoints,
    double MarginLeftPoints,
    double MarginRightPoints,
    double MarginTopPoints,
    double MarginBottomPoints,
    IReadOnlyList<DocxParagraph> Paragraphs)
{
    public DocxDocument(double pageWidthPoints, double pageHeightPoints)
        : this(pageWidthPoints, pageHeightPoints, 72d, 72d, 72d, 72d, [])
    {
    }
}

internal sealed record DocxParagraph(
    IReadOnlyList<DocxTextRun> Runs,
    DocxTextAlignment Alignment,
    double SpacingBeforePoints,
    double SpacingAfterPoints,
    double LineSpacingFactor);

internal sealed record DocxTextRun(
    string Text,
    double FontSize,
    string? ColorHex,
    bool Bold,
    bool Italic,
    bool Underline,
    string? FontFamily);

internal enum DocxTextAlignment
{
    Left,
    Center,
    Right
}
