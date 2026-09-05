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
    IReadOnlyList<DocxParagraph> FallbackParagraphs,
    IReadOnlyList<DocxTable> FallbackTables)
{
    public DocxFontCatalog FontCatalog { get; init; } = DocxFontCatalog.Empty;
    public DocxStyleCatalog StyleCatalog { get; init; } = DocxStyleCatalog.Empty;
    public IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> HeaderParagraphsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> FooterParagraphsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> HeaderBodyElementsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> FooterBodyElementsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> HeaderFloatingDrawingsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> FooterFloatingDrawingsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyList<DocxRelatedStory> RelatedStories { get; init; } = [];
    public IReadOnlyList<string> PackageCommentAnchorIds { get; init; } = [];
    public IReadOnlyList<string> HiddenCommentAnchorIds { get; init; } = [];
    public DocxDocumentSettings Settings { get; init; } = DocxDocumentSettings.Empty;
    public OoxPdfDocxMarkupMode MarkupMode { get; init; } = OoxPdfDocxMarkupMode.Final;
    public DocxSectionBreakElement? FinalSectionBreak { get; init; }
    public IReadOnlyList<DocxParagraph> Paragraphs => BodyElements.Count == 0
        ? FallbackParagraphs
        : DocxBlockTraversal.EnumerateDirectParagraphs(BodyElements).ToArray();
    public IReadOnlyList<DocxTable> Tables => BodyElements.Count == 0
        ? FallbackTables
        : DocxBlockTraversal.EnumerateBodyTables(BodyElements).ToArray();

    public DocxDocument(double pageWidthPoints, double pageHeightPoints)
        : this(pageWidthPoints, pageHeightPoints, 72d, 72d, 72d, 72d, DocxPageSettings.Empty, [], [], [], [], [], [])
    {
    }
}

internal sealed record DocxRelatedStory(
    DocxRelatedStoryKind Kind,
    string PartName,
    string? Id,
    IReadOnlyList<DocxBodyElement> BodyElements,
    IReadOnlyList<DocxParagraph> FallbackParagraphs,
    IReadOnlyList<DocxTable> FallbackTables,
    DocxRelatedStoryType? Type)
{
    public IReadOnlyList<DocxFloatingDrawing> FloatingDrawings { get; init; } = [];
    public bool IsNormalStoryType => Type is null || Type == DocxRelatedStoryType.Normal;
    public DocxCommentMetadata? CommentMetadata { get; init; }
    public IReadOnlyList<DocxParagraph> Paragraphs => BodyElements.Count == 0
        ? FallbackParagraphs
        : DocxBlockTraversal.EnumerateDirectParagraphs(BodyElements).ToArray();
    public IReadOnlyList<DocxTable> Tables => BodyElements.Count == 0
        ? FallbackTables
        : DocxBlockTraversal.EnumerateBodyTables(BodyElements).ToArray();
}

internal sealed record DocxCommentMetadata(
    string? Author,
    string? Initials,
    string? Date,
    string? ParagraphId,
    string? ParentParagraphId,
    string? ParentCommentId,
    bool? IsResolved);