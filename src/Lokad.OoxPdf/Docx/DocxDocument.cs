using System.Collections.ObjectModel;

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
    IReadOnlyList<DocxBodyElement> BodyElements)
{
    public string MainPartName { get; init; } = "/word/document.xml";
    public IReadOnlyList<DocxEquationPartSource> EquationParts { get; init; } = [];
    public IReadOnlyList<Diagnostics.OoxPdfDiagnostic> EquationDiagnostics { get; init; } = [];
    public DocxFontCatalog FontCatalog { get; init; } = DocxFontCatalog.Empty;
    public DocxStyleCatalog StyleCatalog { get; init; } = DocxStyleCatalog.Empty;
    public IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> HeaderBodyElementsByType { get; init; } =
        ReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>>.Empty;
    public IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> FooterBodyElementsByType { get; init; } =
        ReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>>.Empty;
    public IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> HeaderFloatingDrawingsByType { get; init; } =
        ReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>>.Empty;
    public IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> FooterFloatingDrawingsByType { get; init; } =
        ReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>>.Empty;
    public IReadOnlyList<DocxRelatedStory> RelatedStories { get; init; } = [];
    public IReadOnlyList<string> PackageCommentAnchorIds { get; init; } = [];
    public IReadOnlyList<string> HiddenCommentAnchorIds { get; init; } = [];
    public DocxDocumentSettings Settings { get; init; } = DocxDocumentSettings.Empty;
    public OoxPdfDocxMarkupMode MarkupMode { get; init; } = OoxPdfDocxMarkupMode.Final;
    public DocxSectionBreakElement? FinalSectionBreak { get; init; }
    public DocxParagraph? EmptyFinalSectionParagraph { get; init; }
    public IReadOnlyList<DocxParagraph> Paragraphs => DocxBlockTraversal.EnumerateDirectParagraphs(BodyElements).ToArray();
    public IReadOnlyList<DocxTable> Tables => DocxBlockTraversal.EnumerateBodyTables(BodyElements).ToArray();

    public DocxDocument(double pageWidthPoints, double pageHeightPoints)
        : this(pageWidthPoints, pageHeightPoints, 72d, 72d, 72d, 72d, DocxPageSettings.Empty, [], [], [], [])
    {
    }
}

internal sealed record DocxRelatedStory(
    DocxRelatedStoryKind Kind,
    string PartName,
    string? Id,
    IReadOnlyList<DocxBodyElement> BodyElements,
    DocxRelatedStoryType? Type)
{
    public IReadOnlyList<DocxFloatingDrawing> FloatingDrawings { get; init; } = [];
    // The reader supplies default separator content for unregistered footnote
    // and endnote entries and retains their IDs. It also creates an ID 0
    // separator when an endnote part lacks one. Layout uses this flag to
    // distinguish default separator geometry from authored content.
    public bool IsSynthetic { get; init; }
    public int SourceEquationCount { get; init; }
    public bool IsNormalStoryType => Type is null || Type == DocxRelatedStoryType.Normal;
    public DocxCommentMetadata? CommentMetadata { get; init; }
    public IReadOnlyList<DocxParagraph> Paragraphs => DocxBlockTraversal.EnumerateDirectParagraphs(BodyElements).ToArray();
    public IReadOnlyList<DocxTable> Tables => DocxBlockTraversal.EnumerateBodyTables(BodyElements).ToArray();
}

internal sealed record DocxCommentMetadata(
    string? Author,
    string? Initials,
    string? Date,
    string? ParagraphId,
    string? ParentParagraphId,
    string? ParentCommentId,
    bool? IsResolved);
