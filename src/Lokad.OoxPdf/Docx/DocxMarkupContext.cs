namespace Lokad.OoxPdf.Docx;

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
    bool ExpandsMarkupMargin)
{
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

    public static DocxMarkupContext FromMode(
        OoxPdfDocxMarkupMode mode,
        OoxPdfDocxMarkupGeometryMode geometryMode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
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
