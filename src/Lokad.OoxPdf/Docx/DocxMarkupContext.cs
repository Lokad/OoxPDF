namespace Lokad.OoxPdf.Docx;

using Lokad.OoxPdf.Fonts;

internal sealed record DocxMarkupContext(
    OoxPdfDocxMarkupMode Mode,
    OoxPdfDocxMarkupGeometryMode GeometryMode,
    bool IncludesInsertions,
    bool IncludesDeletions,
    bool IncludesMoveFrom,
    bool IncludesMoveTo,
    bool AppliesInlineRevisionStyle,
    bool DrawsChangeBars,
    bool DrawsCommentMarkers,
    bool ApproximatesComments,
    bool ApproximatesTrackedChanges,
    bool ApproximatesFormattingRevisions,
    bool RendersCommentBalloons,
    bool RendersRevisionBalloons,
    bool ExpandsMarkupMargin,
    double WordCompatiblePrintScale = 1d,
    double WordCompatibleTextXOffset = 0d,
    double WordCompatibleTextYOffset = 0d)
{
    // RV06: first-seen comment-author palette slots by comment w:id, built once
    // per conversion from the laid-out related stories. Empty until the renderer
    // attaches document state; readers fall back to slot 0 (the review color).
    public IReadOnlyDictionary<string, int> CommentAuthorPaletteSlots { get; init; } = new Dictionary<string, int>();
    // RV06: resolved Segoe UI Bold face for balloon titles (Office sets titles in Segoe
    // UI Bold, bodies in the document face). Null keeps the label face everywhere.
    public FontFaceResolution? BalloonTitleFaceResolution { get; init; }
    // RV16: comment/reply story lookup built once per conversion from the laid-out
    // related stories; per-page candidate collection reads it instead of rebuilding
    // group dictionaries for every page. Null keeps entry paths without it on the
    // per-page fallback.
    public MarkupCommentStoryIndex? CommentStoryIndex { get; init; }
    public IReadOnlyDictionary<string, string>? CommentMarkerLabels { get; init; }
    // Word prints comments anchored in the main story, including its table cells.
    // A reference set distinguishes identical paragraph values in other stories.
    // Null keeps the established policy for other geometry modes and entry paths.
    public IReadOnlySet<DocxParagraph>? WordCompatibleMainStoryCommentParagraphs { get; init; }
    // The review lane belongs to printed balloons, not the presence of a comment part.
    public bool? WordCompatibleHasPrintedBalloons { get; init; }
    // Plain headers retain their established emission origin when body text
    // selects its own first-baseline anchor on a scaled review page.
    public double WordCompatibleStaticTextYAdjustmentPoints { get; init; }
    public bool IsCommentMarkupVisible(DocxParagraph paragraph) =>
        WordCompatibleMainStoryCommentParagraphs?.Contains(paragraph) ?? true;
    public DocxMarkupContext ApplyDocumentSettings(DocxDocumentSettings settings)
    {
        DocxRevisionViewSettings revisionView = settings.RevisionViewSettings;
        DocxMarkupContext context = this;
        if (revisionView.ShowMarkup == false)
        {
            context = context with
            {
                DrawsChangeBars = false,
                DrawsCommentMarkers = false,
                ApproximatesComments = false,
                ApproximatesTrackedChanges = false,
                ApproximatesFormattingRevisions = false,
                RendersCommentBalloons = false,
                RendersRevisionBalloons = false,
                ExpandsMarkupMargin = false
            };
        }

        if (revisionView.ShowComments == false)
        {
            context = context with
            {
                DrawsCommentMarkers = false,
                ApproximatesComments = false,
                RendersCommentBalloons = false
            };
        }

        if (revisionView.ShowInsertionsAndDeletions == false)
        {
            context = context with
            {
                DrawsChangeBars = false,
                ApproximatesTrackedChanges = false,
                RendersRevisionBalloons = false
            };
        }

        if (revisionView.ShowFormatting == false)
        {
            context = context with
            {
                ApproximatesFormattingRevisions = false
            };
        }

        return context;
    }

    public static DocxMarkupContext FromMode(OoxPdfDocxMarkupMode mode)
    {
        return FromMode(mode, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
    }

    public static DocxMarkupContext FromMode(
        OoxPdfDocxMarkupMode mode,
        OoxPdfDocxMarkupGeometryMode geometryMode)
    {
        bool expandsMarkupMargin = mode == OoxPdfDocxMarkupMode.AllMarkup &&
            geometryMode is OoxPdfDocxMarkupGeometryMode.ReserveMarkupMargin or OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup;
        return mode switch
        {
            OoxPdfDocxMarkupMode.Original => new(
                mode,
                geometryMode,
                IncludesInsertions: false,
                IncludesDeletions: true,
                IncludesMoveFrom: true,
                IncludesMoveTo: false,
                AppliesInlineRevisionStyle: false,
                DrawsChangeBars: false,
                DrawsCommentMarkers: false,
                ApproximatesComments: false,
                ApproximatesTrackedChanges: true,
                ApproximatesFormattingRevisions: true,
                RendersCommentBalloons: false,
                RendersRevisionBalloons: false,
                ExpandsMarkupMargin: false),
            OoxPdfDocxMarkupMode.SimpleMarkup => new(
                mode,
                geometryMode,
                IncludesInsertions: true,
                IncludesDeletions: false,
                IncludesMoveFrom: false,
                IncludesMoveTo: true,
                AppliesInlineRevisionStyle: false,
                DrawsChangeBars: true,
                DrawsCommentMarkers: true,
                ApproximatesComments: true,
                ApproximatesTrackedChanges: true,
                ApproximatesFormattingRevisions: true,
                RendersCommentBalloons: false,
                RendersRevisionBalloons: false,
                ExpandsMarkupMargin: false),
            OoxPdfDocxMarkupMode.AllMarkup => new(
                mode,
                geometryMode,
                IncludesInsertions: true,
                IncludesDeletions: true,
                IncludesMoveFrom: true,
                IncludesMoveTo: true,
                AppliesInlineRevisionStyle: true,
                DrawsChangeBars: true,
                DrawsCommentMarkers: true,
                ApproximatesComments: true,
                ApproximatesTrackedChanges: true,
                ApproximatesFormattingRevisions: true,
                RendersCommentBalloons: true,
                RendersRevisionBalloons: true,
                ExpandsMarkupMargin: expandsMarkupMargin),
            _ => new(
                OoxPdfDocxMarkupMode.Final,
                geometryMode,
                IncludesInsertions: true,
                IncludesDeletions: false,
                IncludesMoveFrom: false,
                IncludesMoveTo: true,
                AppliesInlineRevisionStyle: false,
                DrawsChangeBars: false,
                DrawsCommentMarkers: false,
                ApproximatesComments: false,
                ApproximatesTrackedChanges: false,
                ApproximatesFormattingRevisions: false,
                RendersCommentBalloons: false,
                RendersRevisionBalloons: false,
                ExpandsMarkupMargin: false)
        };
    }
}
