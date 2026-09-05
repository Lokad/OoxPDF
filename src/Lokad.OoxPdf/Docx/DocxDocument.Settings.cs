namespace Lokad.OoxPdf.Docx;

internal sealed record DocxDocumentSettings(
    string? CharacterSpacingControlValue,
    string? DefaultTabStopValue,
    double? DefaultTabStopPoints,
    bool? UseFELayout,
    string? UseFELayoutValue,
    DocxRevisionViewSettings RevisionViewSettings,
    DocxTrackChangesSettings TrackChangesSettings,
    DocxNoteReferenceSettings FootnoteReferenceSettings,
    DocxNoteReferenceSettings EndnoteReferenceSettings,
    string? MirrorMarginsValue,
    bool? MirrorMargins,
    IReadOnlyList<DocxCompatSetting> CompatSettings)
{
    public static DocxDocumentSettings Empty { get; } = new(null, null, null, null, null, DocxRevisionViewSettings.Empty, DocxTrackChangesSettings.Empty, DocxNoteReferenceSettings.Empty, DocxNoteReferenceSettings.Empty, null, null, []);
}

internal sealed record DocxRevisionViewSettings(
    string? MarkupValue,
    bool? ShowMarkup,
    string? CommentsValue,
    bool? ShowComments,
    string? InsertionsAndDeletionsValue,
    bool? ShowInsertionsAndDeletions,
    string? FormattingValue,
    bool? ShowFormatting,
    string? InkAnnotationsValue,
    bool? ShowInkAnnotations)
{
    public static DocxRevisionViewSettings Empty { get; } = new(null, null, null, null, null, null, null, null, null, null);
}

internal sealed record DocxTrackChangesSettings(
    string? TrackRevisionsValue,
    bool? TrackRevisions,
    string? DoNotTrackMovesValue,
    bool? DoNotTrackMoves,
    string? DoNotTrackFormattingValue,
    bool? DoNotTrackFormatting)
{
    public static DocxTrackChangesSettings Empty { get; } = new(null, null, null, null, null, null);
}

internal sealed record DocxNoteReferenceSettings(
    string? PositionValue,
    string? NumberFormatValue,
    string? NumberStartValue,
    int? NumberStart,
    string? NumberRestartValue)
{
    public static DocxNoteReferenceSettings Empty { get; } = new(null, null, null, null, null);
}

internal sealed record DocxCompatSetting(
    string? Name,
    string? Uri,
    string? Value);

internal sealed record DocxPageSettings(
    string? WidthValue,
    string? HeightValue,
    string? OrientationValue,
    string? MarginTopValue,
    string? MarginRightValue,
    string? MarginBottomValue,
    string? MarginLeftValue,
    double? HeaderDistancePoints,
    double? FooterDistancePoints,
    string? HeaderDistanceValue,
    string? FooterDistanceValue,
    bool? TitlePage,
    string? TitlePageValue,
    bool? EvenAndOddHeaders,
    string? EvenAndOddHeadersValue)
{
    public static DocxPageSettings Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

    public double? DocGridLinePitchPoints { get; init; }
    public string? DocGridLinePitchValue { get; init; }
    public double? GutterDistancePoints { get; init; }
    public string? GutterDistanceValue { get; init; }
    public DocxNoteReferenceSettings FootnoteReferenceSettings { get; init; } = DocxNoteReferenceSettings.Empty;
    public DocxNoteReferenceSettings EndnoteReferenceSettings { get; init; } = DocxNoteReferenceSettings.Empty;

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
}
