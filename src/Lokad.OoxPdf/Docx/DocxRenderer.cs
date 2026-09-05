using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxMarkupBalloonPlacementSnapshot(
    int PageIndex,
    string Kind,
    string Side,
    double X,
    double Y,
    double Width,
    double Height,
    double AnchorY,
    bool IsOverflowSummary,
    double AnchorConnectorX,
    double BalloonConnectorX,
    bool AnchorConnectorClamped,
    int CandidateCount,
    int CommentCandidateCount,
    int RevisionCandidateCount,
    int CommentWithDateCount,
    int CommentResolvedCount,
    int CommentOpenCount,
    int CommentReplyCount,
    int BodySummaryPartCount,
    int WordCompatibleBodySummaryPartCount,
    int CommentSeparatorLineCount,
    int? OverflowStartIndex,
    int? OverflowEndIndex,
    int LaneBandIndex,
    int LaneBandCandidateCount);

internal sealed partial class DocxRenderer
{
    internal const string DefaultDocumentTypefaceRequest = DocxFontFallbackRules.DefaultDocumentTypefaceRequest;
    private const double WordCompatibleAllMarkupPrintScale = 0.842391d;
    private const double WordCompatibleAllMarkupLineMetricScale = 0.79359971328d;
    private const double WordCompatibleAllMarkupTextXOffsetPoints = -16.15d;
    private const double WordCompatibleAllMarkupTextBaselineYOffsetPoints = 69.58d;
    private const double WordCompatibleAllMarkupMaxBodyTextFontSizePoints = 11.625d;
    private const double WordCompatibleAllMarkupTerminalLineSpaceFontSizePoints = DocxDefaults.FontSizePoints * WordCompatibleAllMarkupPrintScale;
    private const double WordCompatibleAllMarkupBodyPositioningCharacterSpacingPoints = 0.071d;
    private const double WordCompatibleAllMarkupHeadingPositioningCharacterSpacingPoints = 0.043d;
    private const double WordCompatibleAllMarkupPunctuationPositioningCharacterSpacingPoints = 0.102d;
    private const double WordCompatibleAllMarkupInsertionPositioningCharacterSpacingPoints = 0.126d;
    private const double WordCompatibleAllMarkupDeletionPositioningCharacterSpacingPoints = 0.060d;
    private const double WordCompatibleAllMarkupShortWordPositioningCharacterSpacingPoints = 0.024d;
    private const double WordCompatibleAllMarkupBodyXOffsetAsymptotePoints = -3.0d;
    private const double WordCompatibleAllMarkupBodyXOffsetDecayPoints = 55.0d;
    private const double WordCompatibleAllMarkupDeletionXOffsetPoints = 2.707d;
    private const double WordCompatibleAllMarkupInsertionXOffsetPoints = 2.140d;
    private const double WordCompatibleAllMarkupBalloonAnchorYOffsetPoints = 84.84d;
    private const double WordCompatibleAllMarkupBalloonHeightPoints = 20.48d;
    private const double WordCompatibleAllMarkupBalloonTopInsetPoints = 11.15d;
    private const double WordCompatibleAllMarkupConnectorStrokeWidthPoints = 0.475d;
    private const double WordCompatibleAllMarkupConnectorBodyAnchorInsetPoints = 3.18d;
    private const double WordCompatibleAllMarkupBalloonTextFontSizePoints = 6.975d;
    private const double WordCompatibleAllMarkupBalloonTextInsetXPoints = 3.25d;
    private const double WordCompatibleAllMarkupBalloonTitlePositioningCharacterSpacingPoints = 0.03357d;
    private const double WordCompatibleAllMarkupBalloonBodyFirstLineXOffsetPoints = 2.541d;
    private const double WordCompatibleAllMarkupBalloonContinuationPositioningCharacterSpacingPoints = -0.02186d;
    private const double WordCompatibleAllMarkupBalloonContinuationTerminalSpaceXOffsetPoints = -2.968d;
    private const double WordCompatibleAllMarkupBalloonFirstBaselineOffsetPoints = 11.27d;
    private const double WordCompatibleAllMarkupBalloonFirstBaselineTopInsetPoints = WordCompatibleAllMarkupBalloonHeightPoints - WordCompatibleAllMarkupBalloonFirstBaselineOffsetPoints;
    private const double WordCompatibleAllMarkupCommentThreadReplyHeightPoints = WordCompatibleAllMarkupBalloonTextFontSizePoints * 1.2d;
    private const double WordCompatibleAllMarkupCommentThreadSeparatorYOffsetPoints = 1.9d;
    private const int WordCompatibleAllMarkupCommentThreadMaxSeparatorLineCount = 2;
    private const double WordCompatibleAllMarkupLaneBackgroundRightBleedPoints = 0.37d;
    private const double WordCompatibleAllMarkupLaneBackgroundWidthPoints = 199.70d;
    private const double WordCompatibleAllMarkupLaneBackgroundBottomInsetPoints = 17.475d;
    private const double WordCompatibleAllMarkupLaneBackgroundTopInsetPoints = 16.275d;
    private const double WordCompatibleAllMarkupRevisionBarXPoints = 27.925d;
    private const double WordCompatibleAllMarkupRevisionBarWidthPoints = 0.475d;
    private const double WordCompatibleAllMarkupRevisionBarBottomOutsetPoints = 7.47d;
    private const double WordCompatibleAllMarkupRevisionBarTopInsetPoints = 0.39d;
    private const double WordCompatibleAllMarkupRevisionDecorationThicknessPoints = 0.475d;
    private const double WordCompatibleAllMarkupShortInsertionDecorationWidthInsetPoints = 1.434d;
    private const double WordCompatibleAllMarkupCommentRangeStrokeWidthPoints = 0.475d;
    private const double MarkupBalloonConnectorCollisionAnchorYThresholdPoints = 9d;
    private const double MarkupBalloonMinimumSpacingPoints = 3d;
    private const double MarkupBalloonLaneBandSeparationPoints = 18d;
    private const int MarkupBalloonMaxNearbyRevisionGroupSize = 3;
    private const double WordCompatibleAllMarkupCommentRangeFillXInsetPoints = 0.30d;
    private const double WordCompatibleAllMarkupCommentRangeFillBaselineYOffsetPoints = -2.80d;
    private const double WordCompatibleAllMarkupCommentRangeFillHeightPoints = WordCompatibleAllMarkupMaxBodyTextFontSizePoints;
    private const double WordCompatibleAllMarkupCommentRangeStartInsetPoints = 2.025d;
    private const double WordCompatibleAllMarkupCommentRangeEndInsetPoints = 0.564d;
    private const double WordCompatibleAllMarkupCommentRangeBottomTickYOffsetPoints = -2.75d;
    private const double WordCompatibleAllMarkupCommentRangeVerticalBottomYOffsetPoints = -2.07d;
    private const double WordCompatibleAllMarkupCommentRangeVerticalTopYOffsetPoints = 6.78d;
    private const double WordCompatibleAllMarkupCommentRangeTopTickYOffsetPoints = 7.71d;
    private const double WordCompatibleAllMarkupCommentRangeTickLengthPoints = 0.475d;
    private const double WordCompatibleAllMarkupCommentRangeFarTickLengthPoints = 0.95d;
    private const double WordCompatibleAllMarkupCommentReferenceMarkerWidthPoints = 1.9d;
    private const double WordMarkupBalloonStemWidthPoints = 20.63d;
    private const double MinimumMarkupBalloonBodyWidthPoints = 24d;
    private static readonly DocxMarkupBalloonRgb WordCompatibleAllMarkupReviewStrokeRgb = new(209, 52, 56);
    private static readonly DocxMarkupBalloonRgb WordCompatibleAllMarkupReviewFillRgb = new(248, 220, 221);

    private static readonly DocxMarkupBalloonRgb[] RevisionAuthorColorPalette =
    [
        new(0, 114, 189),
        new(192, 80, 77),
        new(112, 173, 71),
        new(128, 100, 162),
        new(75, 172, 198),
        new(247, 150, 70),
        new(155, 89, 182),
        new(121, 85, 72)
    ];

    private readonly IFontResolver fontResolver;
    private readonly DocxMarkupContext markupContext;

    public DocxRenderer(
        IFontResolver? fontResolver,
        OoxPdfDocxMarkupMode markupMode,
        OoxPdfDocxMarkupGeometryMode markupGeometryMode)
    {
        this.fontResolver = fontResolver ?? new WindowsFontResolver();
        markupContext = DocxMarkupContext.FromMode(markupMode, markupGeometryMode);
    }

    public IReadOnlyList<PdfPage> RenderBlankPages(DocxDocument document, Action<OoxPdfDiagnostic>? diagnosticSink, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!HasRenderableContent())
        {
            return [new PdfPage(document.PageWidthPoints, document.PageHeightPoints)];
        }

        return RenderParagraphs(document, fontResolver, ResolveEffectiveMarkupContext(document), diagnosticSink, cancellationToken);

        bool HasRenderableContent()
        {
            return document.BodyElements.Count != 0
                || HasRenderableDrawings(document.FloatingDrawings)
                || document.HeaderParagraphs.Count != 0
                || document.FooterParagraphs.Count != 0
                || HasBodyElements(document.HeaderBodyElementsByType)
                || HasBodyElements(document.FooterBodyElementsByType)
                || HasBodyElements(document.PageSettings.HeaderBodyElementsByType)
                || HasBodyElements(document.PageSettings.FooterBodyElementsByType)
                || HasRenderableDrawings(document.HeaderFloatingDrawingsByType)
                || HasRenderableDrawings(document.FooterFloatingDrawingsByType)
                || HasRenderableDrawings(document.PageSettings.HeaderFloatingDrawingsByType)
                || HasRenderableDrawings(document.PageSettings.FooterFloatingDrawingsByType);
        }
    }

    internal DocxLayoutSnapshot InspectLayout(DocxDocument document)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver, CancellationToken.None);
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        OoxPdfDocxMarkupGeometryMode effectiveGeometryMode = ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext);
        DocxLayout layout = new DocxLayoutEngine(effectiveGeometryMode).Create(document, ResolveLayoutTextMeasurer(fontResources, effectiveMarkupContext), CancellationToken.None);
        return DocxLayoutSnapshot.FromLayout(layout, document.MarkupMode, effectiveGeometryMode);
    }

    internal IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> InspectMarkupBalloons(DocxDocument document)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver, CancellationToken.None);
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        DocxLayout layout = new DocxLayoutEngine(ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext)).Create(document, ResolveLayoutTextMeasurer(fontResources, effectiveMarkupContext), CancellationToken.None);
        var snapshots = new List<DocxMarkupBalloonPlacementSnapshot>();
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            foreach (DocxMarkupBalloonPlacement placement in BuildMarkupBalloonPlacements(
                layout.Pages[pageIndex],
                layout.RelatedStories,
                EnumeratePageFloatingDrawings(layout, pageIndex).ToArray(),
                effectiveMarkupContext))
            {
                snapshots.Add(placement.ToSnapshot(pageIndex));
            }
        }

        return snapshots;
    }

    internal DocxFontPlanSnapshot InspectFontPlan(DocxDocument document)
    {
        return DocxFontPlanSnapshot.FromPlan(DocxFontPlan.Create(document, fontResolver, CancellationToken.None));
    }

    internal DocxTextEmissionSnapshot InspectTextEmission(DocxDocument document)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver, CancellationToken.None);
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        DocxLayout layout = new DocxLayoutEngine(ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext)).Create(document, ResolveLayoutTextMeasurer(fontResources, effectiveMarkupContext), CancellationToken.None);
        double textEmissionFontScale = ResolveTextEmissionFontScale(effectiveMarkupContext);
        double textEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(effectiveMarkupContext);
        double textEmissionXOffset = ResolveTextEmissionXOffset(effectiveMarkupContext);
        bool useWordCompatibleTextProfile = UsesWordCompatibleAllMarkupTextProfile(effectiveMarkupContext);
        bool suppressCommentReferenceSpacer = ShouldSuppressWordCompatibleCommentReferenceSpacer(effectiveMarkupContext);
        var lines = new List<DocxTextEmissionLineSnapshot>();
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            DocxLayoutPage page = layout.Pages[pageIndex];
            int pageNumber = pageIndex + 1;
            void AddLine(
                DocxTextLineLayout line,
                bool isStaticStory,
                string fallbackStoryKind,
                string? containerStoryKind,
                string? containerStoryVariantType)
            {
                string ResolveTextEmissionStoryKind()
                {
                    return string.IsNullOrWhiteSpace(line.StoryKind) ? fallbackStoryKind : line.StoryKind;
                }

                lines.Add(ToTextEmissionLineSnapshot(
                    pageIndex,
                    isStaticStory,
                    ResolveTextEmissionStoryKind(),
                    line.StoryVariantType,
                    containerStoryKind,
                    containerStoryVariantType,
                    line,
                    fontResources,
                    pageNumber,
                    layout.Pages.Count,
                    textEmissionFontScale,
                    textEmissionBaselineOffset,
                    textEmissionXOffset,
                    suppressCommentReferenceSpacer,
                    useWordCompatibleTextProfile));
            }

            foreach (DocxTextLineLayout line in EnumerateStaticTextLines(page))
            {
                AddLine(line, isStaticStory: true, "Static", line.StoryKind, line.StoryVariantType);
            }

            foreach (DocxTextLineLayout line in EnumerateBodyTextLines(page))
            {
                AddLine(line, isStaticStory: false, "Body", "Body", null);
            }

            foreach (DocxTextLineLayout line in EnumeratePlacedRelatedStoryTextLines(page))
            {
                AddLine(line, isStaticStory: false, "RelatedStory", line.StoryKind, line.StoryVariantType);
            }

            foreach (DocxTextEmissionLineSource source in EnumerateRenderedFloatingDrawingTextBoxTextLines(layout, page, pageIndex))
            {
                lines.Add(ToTextEmissionLineSnapshot(
                    pageIndex,
                    source.IsStaticStory,
                    source.StoryKind,
                    source.StoryVariantType,
                    source.ContainerStoryKind,
                    source.ContainerStoryVariantType,
                    source.Line,
                    fontResources,
                    pageNumber,
                    layout.Pages.Count,
                    textEmissionFontScale,
                    textEmissionBaselineOffset,
                    textEmissionXOffset,
                    suppressCommentReferenceSpacer,
                    useWordCompatibleTextProfile));
            }
        }

        return new DocxTextEmissionSnapshot(
            document.MarkupMode.ToString(),
            lines.Count,
            lines.Sum(line => line.SegmentCount),
            lines.Sum(line => line.TerminalSpaceSegmentCount),
            lines.Sum(line => line.NonzeroPdfCharacterSpacingSegmentCount),
            lines.Sum(line => line.Segments.Count(segment => segment.CompensatePdfCharacterSpacing)),
            lines.Sum(line => line.RevisionSegmentCount),
            lines.Sum(line => line.InsertionRevisionSegmentCount),
            lines.Sum(line => line.DeletionRevisionSegmentCount),
            lines.Sum(line => line.MoveFromRevisionSegmentCount),
            lines.Sum(line => line.MoveToRevisionSegmentCount),
            lines.Sum(line => line.OtherRevisionSegmentCount),
            lines.Count(line => line.CommentReferenceCount != 0),
            lines.Sum(line => line.CommentReferenceCount),
            lines);
    }

    internal DocxStructureSnapshot InspectStructure(DocxDocument document)
    {
        return DocxStructureSnapshot.FromDocument(document);
    }

    private DocxMarkupContext ResolveEffectiveMarkupContext(DocxDocument document)
    {
        return markupContext.ApplyDocumentSettings(document.Settings);
    }

    private static IDocxTextMeasurer? ResolveLayoutTextMeasurer(DocxFontResources fontResources, DocxMarkupContext markupContext)
    {
        double textScale = ResolveTextEmissionFontScale(markupContext);
        double lineMetricScale = ResolveLayoutLineMetricScale(markupContext);
        return fontResources.TextMeasurer is null ||
            (Math.Abs(textScale - 1d) < 0.000001d && Math.Abs(lineMetricScale - 1d) < 0.000001d)
            ? fontResources.TextMeasurer
            : new ScaledDocxTextMeasurer(fontResources.TextMeasurer, textScale, lineMetricScale);
    }

    private static double ResolveTextEmissionFontScale(DocxMarkupContext markupContext)
    {
        return markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin
            ? WordCompatibleAllMarkupPrintScale
            : 1d;
    }

    private static double ResolveTextEmissionBaselineOffset(DocxMarkupContext markupContext)
    {
        return markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin
            ? WordCompatibleAllMarkupTextBaselineYOffsetPoints
            : 0d;
    }

    private static double ResolveTextEmissionXOffset(DocxMarkupContext markupContext)
    {
        return markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin
            ? WordCompatibleAllMarkupTextXOffsetPoints
            : 0d;
    }

    private static bool UsesWordCompatibleAllMarkupTextProfile(DocxMarkupContext markupContext)
    {
        return markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin;
    }

    private static bool ShouldSuppressWordCompatibleCommentReferenceSpacer(DocxMarkupContext markupContext)
    {
        return UsesWordCompatibleAllMarkupTextProfile(markupContext);
    }

    private static double ResolveLayoutLineMetricScale(DocxMarkupContext markupContext)
    {
        return markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin
            ? WordCompatibleAllMarkupLineMetricScale
            : ResolveTextEmissionFontScale(markupContext);
    }

    private static bool HasBodyElements(IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> elementsByType)
    {
        return elementsByType.Values.Any(elements => elements.Count != 0);
    }

    private static bool HasRenderableDrawings(IReadOnlyList<DocxFloatingDrawing> drawings)
    {
        return drawings.Any(drawing => drawing.Image is not null || drawing.TextBoxBodyElements.Count != 0);
    }

    private static bool HasRenderableDrawings(IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> drawingsByType)
    {
        return drawingsByType.Values.Any(HasRenderableDrawings);
    }

    private static IReadOnlyList<PdfPage> RenderParagraphs(
        DocxDocument document,
        IFontResolver fontResolver,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver, cancellationToken);

        DocxLayout layout = new DocxLayoutEngine(ResolveEffectiveMarkupGeometryMode(markupContext)).Create(document, ResolveLayoutTextMeasurer(fontResources, markupContext), cancellationToken);
        double textEmissionFontScale = ResolveTextEmissionFontScale(markupContext);
        double textEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(markupContext);
        double textEmissionXOffset = ResolveTextEmissionXOffset(markupContext);
        bool useWordCompatibleTextProfile = UsesWordCompatibleAllMarkupTextProfile(markupContext);
        bool suppressCommentReferenceSpacer = ShouldSuppressWordCompatibleCommentReferenceSpacer(markupContext);
        IReadOnlyDictionary<string, PdfLinkDestination> bookmarkDestinations = CreateBookmarkDestinations();
        var pages = new List<PdfPage>(layout.Pages.Count);
        int imageIndex = 1;

        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutPage layoutPage = layout.Pages[pageIndex];
            var graphics = new PdfGraphicsBuilder();
            var pageImages = new List<PdfImageResource>();
            int pageNumber = pageIndex + 1;
            RenderWordCompatibleMarkupLaneBackground(layoutPage, graphics, markupContext);
            RenderFloatingDrawings(
                layout.FloatingDrawings,
                pageIndex,
                behindDocument: true,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                ref imageIndex);
            RenderFloatingDrawings(
                layout.StaticFloatingDrawings,
                pageIndex,
                behindDocument: true,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                ref imageIndex);

            IReadOnlyList<DocxLayoutItem> staticItems = EnumerateStaticLayoutItems(layoutPage).ToArray();
            for (int itemIndex = 0; itemIndex < staticItems.Count; itemIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxLayoutItem staticItem = staticItems[itemIndex];
                DocxTableRowLayout? previousRow = itemIndex > 0 ? staticItems[itemIndex - 1] as DocxTableRowLayout : null;
                DocxTableRowLayout? nextRow = itemIndex + 1 < staticItems.Count ? staticItems[itemIndex + 1] as DocxTableRowLayout : null;
                RenderLayoutItem(staticItem, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, layout.Pages.Count, ref imageIndex);
            }

            for (int itemIndex = 0; itemIndex < layoutPage.Items.Count; itemIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxLayoutItem item = layoutPage.Items[itemIndex];
                DocxTableRowLayout? previousRow = itemIndex > 0 ? layoutPage.Items[itemIndex - 1] as DocxTableRowLayout : null;
                DocxTableRowLayout? nextRow = itemIndex + 1 < layoutPage.Items.Count ? layoutPage.Items[itemIndex + 1] as DocxTableRowLayout : null;
                RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageIndex + 1, layout.Pages.Count, ref imageIndex);
            }

            RenderWordCompatibleRevisionBar(layout, layoutPage, pageIndex, graphics, markupContext);
            RenderMarkupBalloons(
                layoutPage,
                layout.RelatedStories,
                EnumeratePageFloatingDrawings(layout, pageIndex).ToArray(),
                graphics,
                fontResources,
                markupContext);

            foreach (DocxPlacedRelatedStoryLayout story in layoutPage.PlacedRelatedStories)
            {
                cancellationToken.ThrowIfCancellationRequested();
                RenderPlacedRelatedStoryDrawings(story, behindDocument: true, graphics, pageImages, fontResources, markupContext, pageNumber, layout.Pages.Count, diagnosticSink, ref imageIndex);
                RenderPlacedRelatedStory(story, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, layout.Pages.Count, ref imageIndex);
                RenderPlacedRelatedStoryDrawings(story, behindDocument: false, graphics, pageImages, fontResources, markupContext, pageNumber, layout.Pages.Count, diagnosticSink, ref imageIndex);
            }

            RenderFloatingDrawings(
                layout.FloatingDrawings,
                pageIndex,
                behindDocument: false,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                ref imageIndex);
            RenderFloatingDrawings(
                layout.StaticFloatingDrawings,
                pageIndex,
                behindDocument: false,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                ref imageIndex);

            IReadOnlyList<PdfLinkAnnotation> annotations = CreateHyperlinkAnnotations(layoutPage, pageIndex, pageNumber, layout.Pages.Count);
            pages.Add(new PdfPage(
                layoutPage.Width,
                layoutPage.Height,
                graphics.ToString(),
                fontResources.Resources,
                pageImages.ToArray(),
                graphics.ExtGStates,
                graphics.Shadings,
                graphics.Patterns,
                annotations));
        }

        return pages;

        IReadOnlyDictionary<string, PdfLinkDestination> CreateBookmarkDestinations()
        {
            var destinations = new Dictionary<string, PdfLinkDestination>(StringComparer.Ordinal);
            for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxLayoutPage page = layout.Pages[pageIndex];
                int pageNumber = pageIndex + 1;
                foreach (DocxTextLineLayout line in EnumerateRenderedPageTextLines(layout, page, pageIndex))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (line.SourceParagraph is not { } paragraph ||
                        paragraph.BookmarkAnchors.Count == 0)
                    {
                        continue;
                    }

                    IReadOnlyList<DocxTextEmissionSegment> segments = CreateTextEmissionSegments(line, fontResources, pageNumber, layout.Pages.Count, textEmissionFontScale, textEmissionBaselineOffset, textEmissionXOffset, suppressCommentReferenceSpacer, useWordCompatibleTextProfile)
                        .Where(segment => !segment.IsTerminalLineSpace && segment.SourceTextRunIndex >= 0 && segment.Width > 0d)
                        .ToArray();
                    if (segments.Count == 0)
                    {
                        continue;
                    }

                    foreach (DocxBookmarkAnchor bookmark in paragraph.BookmarkAnchors)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (string.IsNullOrEmpty(bookmark.Name) || destinations.ContainsKey(bookmark.Name))
                        {
                            continue;
                        }

                        if (!TryResolveBookmarkDestinationSegment(segments, bookmark, out DocxTextEmissionSegment? target, out double targetX))
                        {
                            continue;
                        }

                        double ascender = target.Resource.Embedded.Font.Os2.WindowsAscender * target.FontSize / target.Resource.Embedded.Font.UnitsPerEm;
                        destinations[bookmark.Name] = new PdfLinkDestination(
                            pageIndex,
                            targetX,
                            target.BaselineY + ascender,
                            Zoom: null);
                    }
                }
            }

            return destinations;
        }

        IReadOnlyList<PdfLinkAnnotation> CreateHyperlinkAnnotations(DocxLayoutPage page, int pageIndex, int pageNumber, int pageCount)
        {
            var annotations = new List<PdfLinkAnnotation>();
            foreach (DocxTextLineLayout line in EnumerateRenderedPageTextLines(layout, page, pageIndex))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (line.SourceParagraph is not { } paragraph ||
                    paragraph.Hyperlinks.Count == 0)
                {
                    continue;
                }

                IReadOnlyList<DocxHyperlinkSpan> links = paragraph.Hyperlinks;
                foreach (DocxTextEmissionSegment segment in CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount, textEmissionFontScale, textEmissionBaselineOffset, textEmissionXOffset, suppressCommentReferenceSpacer, useWordCompatibleTextProfile))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (segment.IsTerminalLineSpace || segment.SourceTextRunIndex < 0 || segment.Width <= 0d)
                    {
                        continue;
                    }

                    DocxHyperlinkSpan? link = links.FirstOrDefault(item => IsHyperlinkSegment(item, segment.SourceTextRunIndex));
                    if (link is null)
                    {
                        continue;
                    }

                    double ascender = segment.Resource.Embedded.Font.Os2.WindowsAscender * segment.FontSize / segment.Resource.Embedded.Font.UnitsPerEm;
                    double descender = segment.Resource.Embedded.Font.Os2.WindowsDescender * segment.FontSize / segment.Resource.Embedded.Font.UnitsPerEm;
                    double annotationWidth = ResolveHyperlinkAnnotationWidth(segment, useWordCompatibleTextProfile);
                    if (IsExternalHyperlink(link))
                    {
                        annotations.Add(PdfLinkAnnotation.ToUri(
                            segment.X,
                            segment.BaselineY - descender,
                            annotationWidth,
                            ascender + descender,
                            link.Target ?? string.Empty));
                    }
                    else if (!string.IsNullOrEmpty(link.Anchor) &&
                        bookmarkDestinations.TryGetValue(link.Anchor, out PdfLinkDestination destination))
                    {
                        annotations.Add(PdfLinkAnnotation.ToDestination(
                            segment.X,
                            segment.BaselineY - descender,
                            annotationWidth,
                            ascender + descender,
                            destination));
                    }
                }
            }

            return annotations;

        bool IsHyperlinkSegment(DocxHyperlinkSpan link, int sourceTextRunIndex)
        {
            return sourceTextRunIndex >= link.SourceRunStartIndex &&
                sourceTextRunIndex < link.SourceRunStartIndex + link.SourceRunCount;
        }

        bool IsExternalHyperlink(DocxHyperlinkSpan link)
        {
            return !string.IsNullOrEmpty(link.Target) &&
                string.Equals(link.TargetMode, "External", StringComparison.OrdinalIgnoreCase);
        }
        }
    }
    private static double ResolveHyperlinkAnnotationWidth(
        DocxTextEmissionSegment segment,
        bool useWordCompatibleTextProfile)
    {
        if (!useWordCompatibleTextProfile)
        {
            return segment.Width;
        }

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            segment.StyleRun,
            segment.FontSize,
            segment.PdfCharacterSpacing,
            segment.PdfCharacterSpacingSource,
            segment.CompensatePdfCharacterSpacing,
            segment.IsTerminalLineSpace);
        double emittedAdvance = DocxTextEmissionPlanner.MeasureAdvanceProfile(
            segment.Text,
            segment.Resource.Embedded,
            segment.Width,
            plan).PlannedEmittedAdvance;
        return emittedAdvance > 0d ? emittedAdvance : segment.Width;
    }

    private static bool TryResolveBookmarkDestinationSegment(
        IReadOnlyList<DocxTextEmissionSegment> segments,
        DocxBookmarkAnchor bookmark,
        [NotNullWhen(true)] out DocxTextEmissionSegment? target,
        out double targetX)
    {
        target = null;
        targetX = 0d;
        int sourceRunIndex = bookmark.SourceRunIndex;
        int sourceOffset = Math.Max(0, bookmark.TextOffset);
        foreach (DocxTextEmissionSegment segment in segments)
        {
            if (segment.SourceTextRunIndex != sourceRunIndex)
            {
                continue;
            }

            int segmentStart = Math.Max(0, segment.SourceTextOffsetInRun);
            int segmentEnd = segmentStart + Math.Max(0, segment.Text.Length);
            if (sourceOffset < segmentStart || sourceOffset > segmentEnd)
            {
                continue;
            }

            target = segment;
            targetX = segment.Text.Length == 0
                ? segment.X
                : segment.X + Math.Max(0d, segment.Width) * Math.Clamp((sourceOffset - segmentStart) / (double)segment.Text.Length, 0d, 1d);
            return true;
        }

        DocxTextEmissionSegment? fallback = segments.FirstOrDefault(segment => segment.SourceTextRunIndex >= sourceRunIndex);
        if (fallback is null)
        {
            return false;
        }

        target = fallback;
        targetX = fallback.X;
        return true;
    }


    private static DocxFontResources PrepareFontResources(DocxDocument document, IFontResolver fontResolver, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxFontPlan plan = DocxFontPlan.Create(document, fontResolver, cancellationToken);
        var resources = new List<PdfFontResource>();
        var runResources = new Dictionary<DocxTextRun, DocxRunFontResource>();
        var fontCache = new Dictionary<(string StableId, int FaceIndex), OpenTypeFont?>();
        PrepareResolvedRunFontResources(plan, resources, runResources, fontCache, cancellationToken);
        DocxRunFontResource? fallback = PrepareFallbackFontResource(plan, fontResolver, resources, runResources, fontCache, cancellationToken);
        IDocxTextMeasurer? measurer = plan.Runs.Any(run => LoadFont(run.Resolution, fontCache, cancellationToken) is not null) || fallback is not null
            ? new DocxFontPlanTextMeasurer(plan, fallback?.Resolution, cancellationToken)
            : null;
        return new DocxFontResources(plan, measurer, resources, runResources, fallback);
    }

    private sealed class ScaledDocxTextMeasurer(IDocxTextMeasurer inner, double textScale, double lineMetricScale) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
    {
        public double MeasureText(DocxTextRun? run, string text, double fontSize)
        {
            return inner.MeasureText(run, text, fontSize) * textScale;
        }

        public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxLineMetricsProvider lineMetrics
                ? lineMetrics.MeasureSingleLineHeight(run, fontSize) * lineMetricScale
                : fontSize * 1.2d * lineMetricScale;
        }

        public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxStaticTextMetricsProvider staticMetrics
                ? staticMetrics.MeasureWindowsAscender(run, fontSize) * lineMetricScale
                : fontSize * lineMetricScale;
        }

        public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
        {
            return inner is IDocxStaticTextMetricsProvider staticMetrics
                ? staticMetrics.MeasureWindowsDescender(run, fontSize) * lineMetricScale
                : fontSize * 0.2d * lineMetricScale;
        }
    }

    private static void PrepareResolvedRunFontResources(
        DocxFontPlan plan,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken)
    {
        var resolvedRuns = new List<(DocxResolvedRunTypeface Run, FontFaceResolution Resolution)>();
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (run.Resolution is not { } resolution || LoadFont(resolution, fontCache, cancellationToken) is null)
            {
                continue;
            }
            resolvedRuns.Add((run, resolution));
        }

        foreach (IGrouping<(string StableId, int FaceIndex), (DocxResolvedRunTypeface Run, FontFaceResolution Resolution)> group in resolvedRuns.GroupBy(item => (item.Resolution.Source.StableId, item.Resolution.FontFaceIndex)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            FontFaceResolution resolution = group.First().Resolution;
            IReadOnlyList<int> glyphs = CollectRunGlyphs(group.Select(item => item.Run), cancellationToken);
            if (glyphs.Count == 0)
            {
                continue;
            }

            OpenTypeFont? font = LoadFont(resolution, fontCache, cancellationToken);
            if (font is null)
            {
                continue;
            }

            PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs, cancellationToken);
            string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
            var runResource = new DocxRunFontResource(name, embedded, resolution);
            resources.Add(new PdfFontResource(name, embedded));
            foreach (DocxResolvedRunTypeface run in group.Select(item => item.Run))
            {
                cancellationToken.ThrowIfCancellationRequested();
                runResources[run.Run] = runResource;
            }
        }
    }

    private static DocxRunFontResource? PrepareFallbackFontResource(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        List<PdfFontResource> resources,
        Dictionary<DocxTextRun, DocxRunFontResource> runResources,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken)
    {
        FontFaceResolution resolution = ResolveDocumentBaseFont(plan, fontResolver, fontCache, cancellationToken);
        OpenTypeFont? font = LoadFont(resolution, fontCache, cancellationToken);
        if (font is null)
        {
            return null;
        }

        DocxResolvedRunTypeface[] fallbackRuns = plan.Runs
            .Where(run => !runResources.ContainsKey(run.Run))
            .ToArray();
        if (fallbackRuns.Length == 0)
        {
            return null;
        }

        IReadOnlyList<int> glyphs = CollectRunGlyphs(fallbackRuns, cancellationToken);
        if (glyphs.Count == 0)
        {
            return null;
        }

        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, glyphs, cancellationToken);
        string name = "F" + (resources.Count + 1).ToString(CultureInfo.InvariantCulture);
        var runResource = new DocxRunFontResource(name, embedded, resolution);
        resources.Add(new PdfFontResource(name, embedded));
        foreach (DocxResolvedRunTypeface run in fallbackRuns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            runResources[run.Run] = runResource;
        }

        return runResource;
    }

    private static IReadOnlyList<int> CollectRunGlyphs(IEnumerable<DocxResolvedRunTypeface> runs, CancellationToken cancellationToken)
    {
        var glyphs = new HashSet<int>();
        foreach (DocxResolvedRunTypeface run in runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (Rune rune in run.Run.Text.EnumerateRunes())
            {
                cancellationToken.ThrowIfCancellationRequested();
                glyphs.Add(rune.Value);
            }
        }

        foreach (Rune rune in " 0123456789".EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            glyphs.Add(rune.Value);
        }

        return glyphs.ToArray();
    }

    private static FontFaceResolution ResolveDocumentBaseFont(
        DocxFontPlan plan,
        IFontResolver fontResolver,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken)
    {
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (run.Resolution is { } fontResolution && LoadFont(fontResolution, fontCache, cancellationToken) is not null)
            {
                return fontResolution;
            }
        }

        return DocxFontFallbackRules.ResolveDefaultDocumentTypeface(fontResolver, false, false);
    }

    private static OpenTypeFont? LoadFont(
        FontFaceResolution? resolution,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (resolution is null)
        {
            return null;
        }

        var key = (resolution.Source.StableId, resolution.FontFaceIndex);
        if (fontCache.TryGetValue(key, out OpenTypeFont? cached))
        {
            return cached;
        }

        cached = FontProgramLoader.Load(resolution, cancellationToken);
        fontCache[key] = cached;
        return cached;
    }

    private static void RenderLayoutItem(
        DocxLayoutItem item,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        ref int imageIndex)
    {
        switch (item)
        {
            case DocxTextLineLayout textLine:
                RenderTextLine(textLine, graphics, fontResources, markupContext, pageNumber, pageCount);
                break;
            case DocxInlineImageLayout image:
                RenderInlineImage(image, graphics, pageImages, diagnosticSink, ref imageIndex);
                break;
            case DocxTableRowLayout row:
                RenderTableRow(row, IsAdjacentTableRow(previousRow, row) ? previousRow : null, IsAdjacentTableRow(row, nextRow) ? nextRow : null, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, ref imageIndex);
                break;
        }
    }

    private static OoxPdfDocxMarkupGeometryMode ResolveEffectiveMarkupGeometryMode(DocxMarkupContext context)
    {
        return context.ExpandsMarkupMargin
            ? context.GeometryMode
            : OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;
    }

    private static bool IsAdjacentTableRow(DocxTableRowLayout? first, DocxTableRowLayout? second)
    {
        return first is not null &&
            second is not null &&
            first.Table.TableIndex == second.Table.TableIndex &&
            (first.RowIndex + 1 == second.RowIndex ||
                (first.RowIndex == second.RowIndex && first.FragmentIndex + 1 == second.FragmentIndex));
    }

    private static void RenderWordCompatibleMarkupLaneBackground(
        DocxLayoutPage page,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        if (!UsesWordCompatibleAllMarkupTextProfile(markupContext))
        {
            return;
        }

        double bottom = page.MarginBottom + WordCompatibleAllMarkupLaneBackgroundBottomInsetPoints;
        double top = page.Height - page.MarginTop - WordCompatibleAllMarkupLaneBackgroundTopInsetPoints;
        double height = Math.Max(0d, top - bottom);
        if (height <= 0d)
        {
            return;
        }

        double width = WordCompatibleAllMarkupLaneBackgroundWidthPoints;
        double x = ShouldUseLeftMarkupLane(page)
            ? WordCompatibleAllMarkupLaneBackgroundRightBleedPoints
            : page.Width - width - WordCompatibleAllMarkupLaneBackgroundRightBleedPoints;
        graphics.SetFillRgb(242, 242, 242);
        graphics.FillRectangle(x, bottom, width, height);
    }

    private static bool ShouldUseLeftMarkupLane(DocxLayoutPage page)
    {
        double leftAvailable = Math.Max(0d, page.MarginLeft - 8d);
        double rightAvailable = Math.Max(0d, page.MarginRight - 8d);
        return leftAvailable > rightAvailable && leftAvailable >= MinimumMarkupBalloonBodyWidthPoints;
    }

    private static void RenderWordCompatibleRevisionBar(
        DocxLayout layout,
        DocxLayoutPage page,
        int pageIndex,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        if (!UsesWordCompatibleAllMarkupTextProfile(markupContext) ||
            !markupContext.DrawsChangeBars)
        {
            return;
        }

        double? bottom = null;
        double? top = null;
        void IncludeRevisionBounds(double y, double height)
        {
            bottom = bottom is null ? y : Math.Min(bottom.Value, y);
            top = top is null ? y + height : Math.Max(top.Value, y + height);
        }

        foreach (DocxTextLineLayout line in EnumerateRenderedPageTextLines(layout, page, pageIndex))
        {
            if (!HasTextLineRevision(line))
            {
                continue;
            }

            double scaledFontSize = line.FontSize * ResolveTextEmissionFontScale(markupContext);
            double height = Math.Max(6d, line.LineHeight ?? scaledFontSize * 1.2d);
            double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
            double y = baselineY - height * 0.25d;
            IncludeRevisionBounds(y, height);
        }

        foreach (DocxTableRowLayout row in EnumerateMarkupBalloonTableRows(page, EnumeratePageFloatingDrawings(layout, pageIndex).ToArray()))
        {
            if (!HasTableRowRevision(row))
            {
                continue;
            }

            IncludeRevisionBounds(row.Y, Math.Max(6d, row.Height));
        }

        if (bottom is null || top is null)
        {
            return;
        }

        double barBottom = Math.Max(0d, bottom.Value - WordCompatibleAllMarkupRevisionBarBottomOutsetPoints);
        double barTop = Math.Min(page.Height, top.Value - WordCompatibleAllMarkupRevisionBarTopInsetPoints);
        double barHeight = Math.Max(0d, barTop - barBottom);
        if (barHeight <= 0d)
        {
            return;
        }

        graphics.SetFillRgb(0, 0, 0);
        graphics.FillRectangle(
            WordCompatibleAllMarkupRevisionBarXPoints,
            barBottom,
            WordCompatibleAllMarkupRevisionBarWidthPoints,
            barHeight);

        bool HasTableRowRevision(DocxTableRowLayout row)
        {
            return row.RevisionCount != 0 ||
                row.Table.Revisions?.Count > 0;
        }

        bool HasTextLineRevision(DocxTextLineLayout line)
        {
            return CollectTextLineRevisions(line).Count != 0;
        }
    }

    private static IReadOnlyList<DocxRevisionInfo> CollectTextLineRevisions(DocxTextLineLayout line)
    {
        var revisions = new List<DocxRevisionInfo>();
        if (line.SourceParagraph is not null)
        {
            revisions.AddRange(line.SourceParagraph.Revisions);
        }

        AddTextRunRevisions(line.StyleRun, revisions);
        foreach (DocxTextSegmentLayout segment in line.Segments)
        {
            AddTextRunRevisions(segment.StyleRun, revisions);
        }

        return revisions;
    }

    private static void AddTextRunRevisions(DocxTextRun run, List<DocxRevisionInfo> revisions)
    {
        if (run.Revision is not null)
        {
            revisions.Add(run.Revision);
        }

        revisions.AddRange(run.Revisions);
    }

    private static void RenderTextLine(
        DocxTextLineLayout line,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount)
    {
        RenderMarkupIndicators(line, graphics, fontResources, markupContext);
        foreach (DocxTextEmissionSegment segment in CreateTextEmissionSegments(
            line,
            fontResources,
            pageNumber,
            pageCount,
            ResolveTextEmissionFontScale(markupContext),
            ResolveTextEmissionBaselineOffset(markupContext),
            ResolveTextEmissionXOffset(markupContext),
            ShouldSuppressWordCompatibleCommentReferenceSpacer(markupContext),
            UsesWordCompatibleAllMarkupTextProfile(markupContext)))
        {
            RenderTextEmissionSegment(segment, graphics, markupContext);
        }
    }

    private static void RenderMarkupIndicators(
        DocxTextLineLayout line,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext)
    {
        if (!markupContext.DrawsChangeBars && !markupContext.DrawsCommentMarkers)
        {
            return;
        }

        IReadOnlyList<DocxRevisionInfo> lineRevisions = markupContext.DrawsChangeBars && !UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? CollectTextLineRevisions(line)
            : [];
        if (markupContext.DrawsChangeBars &&
            lineRevisions.Count != 0 &&
            !UsesWordCompatibleAllMarkupTextProfile(markupContext))
        {
            double scaledFontSize = line.FontSize * ResolveTextEmissionFontScale(markupContext);
            double height = Math.Max(6d, line.LineHeight ?? scaledFontSize * 1.2d);
            double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
            double y = baselineY - height * 0.25d;
            DocxMarkupBalloonRgb color = ResolveRevisionAuthorColor(lineRevisions);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(Math.Max(0d, line.X - 7d), y, 1.5d, height);
        }

        if (line.SourceParagraph is not { } paragraph)
        {
            return;
        }

        if (markupContext.DrawsCommentMarkers && paragraph.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment))
        {
            if (UsesWordCompatibleAllMarkupTextProfile(markupContext))
            {
                RenderWordCompatibleCommentRangeMarkers(line, paragraph, graphics, markupContext);
                return;
            }

            double scaledFontSize = line.FontSize * ResolveTextEmissionFontScale(markupContext);
            string label = ResolveCommentMarkerLabel(paragraph);
            double labelFontSize = Math.Max(4.5d, Math.Min(7d, scaledFontSize * 0.55d));
            double markerHeight = Math.Max(6d, labelFontSize + 2d);
            double markerWidth = Math.Max(markerHeight, label.Length * labelFontSize * 0.55d + 3d);
            double markerX = line.X + Math.Max(0d, line.Width) + 2d;
            double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
            double markerY = baselineY + markerHeight * 0.15d;
            graphics.SetFillRgb(255, 192, 0);
            graphics.FillRectangle(markerX, markerY, markerWidth, markerHeight);
            graphics.SetStrokeRgb(217, 151, 0);
            graphics.SetLineWidth(0.5d);
            graphics.StrokeRectangle(markerX, markerY, markerWidth, markerHeight);
            if (!ShouldDrawCommentMarkerLabel(markupContext))
            {
                return;
            }

            DocxRunFontResource? labelResource = fontResources.Fallback ??
                line.Segments
                    .Select(segment => ResolveFontResource(segment.StyleRun, fontResources))
                    .FirstOrDefault(resource => resource is not null);
            string glyphHex = labelResource?.Embedded.EncodeGlyphHex(label) ?? string.Empty;
            if (labelResource is not null && glyphHex.Length != 0)
            {
                graphics.DrawGlyphText(labelResource.Name, labelFontSize, markerX + 1.5d, markerY + 1.4d, 0, 0, 0, glyphHex, italic: false, characterSpacing: 0d, textRenderingMode: 0, strokeRed: 0, strokeGreen: 0, strokeBlue: 0, strokeWidth: 0d);
            }
        }
    }

    private static void RenderWordCompatibleCommentRangeMarkers(
        DocxTextLineLayout line,
        DocxParagraph paragraph,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        foreach (DocxInlineReference reference in paragraph.InlineReferences.Where(reference => reference.Kind == DocxRelatedStoryKind.Comment))
        {
            DocxCommentRange? range = paragraph.CommentRanges.FirstOrDefault(range =>
                string.Equals(range.Id, reference.Id, StringComparison.Ordinal));
            if (range is not null &&
                TryResolveWordCompatibleCommentRangeBounds(line, range, out double startX, out double endX))
            {
                double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
                RenderWordCompatibleCommentRangeMarker(startX, endX, baselineY, graphics);
            }
            else if ((range is null || !HasCommentRangeBounds(range)) &&
                TryResolveWordCompatibleCommentReferenceMarkerBounds(line, reference, out startX, out endX))
            {
                double baselineY = line.BaselineY - ResolveTextEmissionBaselineOffset(markupContext);
                RenderWordCompatibleCommentRangeMarker(startX, endX, baselineY, graphics);
            }
        }

        bool HasCommentRangeBounds(DocxCommentRange range)
        {
            return range.StartSourceRunIndex is not null ||
                range.StartTextOffset is not null ||
                range.EndSourceRunIndex is not null ||
                range.EndTextOffset is not null;
        }
    }

    private static bool TryResolveWordCompatibleCommentReferenceMarkerBounds(
        DocxTextLineLayout line,
        DocxInlineReference reference,
        out double startX,
        out double endX)
    {
        if (!TryResolveSourceOffsetAnchorX(line, reference.SourceRunIndex, reference.TextOffsetInRun, out double anchorX) &&
            !TryResolvePreviousSourceRunEndAnchorX(line, reference.SourceRunIndex, minimumSourceRunIndex: null, out anchorX))
        {
            startX = 0d;
            endX = 0d;
            return false;
        }

        double centerX = anchorX + WordCompatibleAllMarkupTextXOffsetPoints;
        startX = centerX - WordCompatibleAllMarkupCommentReferenceMarkerWidthPoints / 2d;
        endX = centerX + WordCompatibleAllMarkupCommentReferenceMarkerWidthPoints / 2d;
        return endX > startX;
    }

    private static bool TryResolveWordCompatibleCommentRangeBounds(
        DocxTextLineLayout line,
        DocxCommentRange range,
        out double startX,
        out double endX)
    {
        startX = 0d;
        endX = 0d;
        DocxTextSegmentLayout[] rangeSegments = line.Segments
            .Where(segment => SegmentOverlapsCommentRange(segment, range))
            .ToArray();
        if (rangeSegments.Length == 0)
        {
            return false;
        }

        if (!TryResolveSourceOffsetAnchorX(line, range.StartSourceRunIndex, range.StartTextOffset, out double layoutStartX))
        {
            layoutStartX = rangeSegments.Min(segment => segment.X);
        }

        if (!TryResolveSourceOffsetAnchorX(line, range.EndSourceRunIndex, range.EndTextOffset, out double layoutEndX) &&
            !TryResolvePreviousSourceRunEndAnchorX(line, range.EndSourceRunIndex, range.StartSourceRunIndex, out layoutEndX))
        {
            layoutEndX = rangeSegments.Max(segment => segment.X + Math.Max(0d, segment.Width));
        }

        if (layoutEndX <= layoutStartX)
        {
            return false;
        }

        startX = layoutStartX +
            WordCompatibleAllMarkupTextXOffsetPoints -
            WordCompatibleAllMarkupCommentRangeStartInsetPoints;
        endX = layoutEndX +
            WordCompatibleAllMarkupTextXOffsetPoints -
            WordCompatibleAllMarkupCommentRangeEndInsetPoints;
        return endX > startX;
    }

    private static bool SegmentOverlapsCommentRange(
        DocxTextSegmentLayout segment,
        DocxCommentRange range)
    {
        if (segment.SourceTextRunIndex < 0 ||
            range.StartSourceRunIndex is not { } startRunIndex)
        {
            return false;
        }

        int endRunIndex = range.EndSourceRunIndex ?? range.ReferenceSourceRunIndex ?? startRunIndex;
        int segmentRunIndex = segment.SourceTextRunIndex;
        if (segmentRunIndex < startRunIndex || segmentRunIndex > endRunIndex)
        {
            return false;
        }

        int segmentStartOffset = Math.Max(0, segment.SourceTextOffsetInRun);
        int segmentEndOffset = segmentStartOffset + Math.Max(0, segment.Text.Length);
        if (segmentRunIndex == startRunIndex &&
            segmentEndOffset <= Math.Max(0, range.StartTextOffset ?? 0))
        {
            return false;
        }

        if (segmentRunIndex == endRunIndex &&
            segmentStartOffset >= Math.Max(0, range.EndTextOffset ?? int.MaxValue))
        {
            return false;
        }

        return segmentEndOffset > segmentStartOffset || segment.Width > 0d;
    }

    private static void RenderWordCompatibleCommentRangeMarker(
        double startX,
        double endX,
        double baselineY,
        PdfGraphicsBuilder graphics)
    {
        DocxMarkupBalloonRgb color = WordCompatibleAllMarkupReviewStrokeRgb;
        double bottomTickY = baselineY + WordCompatibleAllMarkupCommentRangeBottomTickYOffsetPoints;
        double verticalBottomY = baselineY + WordCompatibleAllMarkupCommentRangeVerticalBottomYOffsetPoints;
        double verticalTopY = baselineY + WordCompatibleAllMarkupCommentRangeVerticalTopYOffsetPoints;
        double topTickY = baselineY + WordCompatibleAllMarkupCommentRangeTopTickYOffsetPoints;
        graphics.SetFillRgb(WordCompatibleAllMarkupReviewFillRgb.Red, WordCompatibleAllMarkupReviewFillRgb.Green, WordCompatibleAllMarkupReviewFillRgb.Blue);
        graphics.FillRectangle(
            Math.Max(0d, startX - WordCompatibleAllMarkupCommentRangeFillXInsetPoints),
            baselineY + WordCompatibleAllMarkupCommentRangeFillBaselineYOffsetPoints,
            Math.Max(0d, endX - startX),
            WordCompatibleAllMarkupCommentRangeFillHeightPoints);
        graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        graphics.SetLineWidth(WordCompatibleAllMarkupCommentRangeStrokeWidthPoints);
        graphics.StrokeLine(
            startX + WordCompatibleAllMarkupCommentRangeTickLengthPoints,
            bottomTickY,
            startX + WordCompatibleAllMarkupCommentRangeFarTickLengthPoints,
            bottomTickY);
        graphics.StrokeLine(startX, verticalBottomY, startX, verticalTopY);
        graphics.StrokeLine(
            startX,
            topTickY,
            startX + WordCompatibleAllMarkupCommentRangeTickLengthPoints,
            topTickY);
        graphics.StrokeLine(
            endX - WordCompatibleAllMarkupCommentRangeFarTickLengthPoints,
            bottomTickY,
            endX - WordCompatibleAllMarkupCommentRangeTickLengthPoints,
            bottomTickY);
        graphics.StrokeLine(endX, verticalBottomY, endX, verticalTopY);
        graphics.StrokeLine(
            endX - WordCompatibleAllMarkupCommentRangeFarTickLengthPoints,
            topTickY,
            endX,
            topTickY);
    }

    private static bool ShouldDrawCommentMarkerLabel(DocxMarkupContext markupContext)
    {
        return markupContext.Mode != OoxPdfDocxMarkupMode.AllMarkup ||
            markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
            !markupContext.ExpandsMarkupMargin;
    }

    private static string ResolveCommentMarkerLabel(DocxParagraph paragraph)
    {
        string? id = paragraph
            .InlineReferences
            .Where(reference => reference.Kind == DocxRelatedStoryKind.Comment)
            .Select(reference => reference.Id)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        return id is null ? "?" : id;
    }

    private static DocxRevisionMarkupPalette ResolveRevisionMarkupPalette(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        DocxMarkupBalloonRgb stroke = ResolveRevisionAuthorColor(revisions);
        return new DocxRevisionMarkupPalette(
            FillRgb: MixRgb(stroke, new DocxMarkupBalloonRgb(255, 255, 255), 0.88d),
            StrokeRgb: stroke,
            TitleRgb: MixRgb(stroke, new DocxMarkupBalloonRgb(0, 0, 0), 0.35d));
    }

    private static DocxMarkupBalloonRgb ResolveRevisionAuthorColor(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        int bucket = ResolveRevisionAuthorBucket(ResolveRevisionAuthorBucketKey(revisions));
        return RevisionAuthorColorPalette[bucket];
    }

    internal static (byte Red, byte Green, byte Blue) ResolveRevisionAuthorColorSnapshot(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        DocxMarkupBalloonRgb color = ResolveRevisionAuthorColor(revisions);
        return (color.Red, color.Green, color.Blue);
    }

    private static string ResolveRevisionAuthorBucketKey(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        return revisions
            .Select(revision => NormalizeRevisionAuthorBucketKey(revision.Author))
            .Where(author => author.Length != 0)
            .GroupBy(author => author, StringComparer.Ordinal)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key, StringComparer.Ordinal)
            .Select(group => group.Key)
            .FirstOrDefault() ?? string.Empty;
    }

    private static int ResolveRevisionAuthorBucket(string? author)
    {
        string normalized = NormalizeRevisionAuthorBucketKey(author);
        unchecked
        {
            uint hash = 2166136261u;
            foreach (char character in normalized)
            {
                hash ^= character;
                hash *= 16777619u;
            }

            return (int)(hash % (uint)RevisionAuthorColorPalette.Length);
        }
    }

    private static string NormalizeRevisionAuthorBucketKey(string? author)
    {
        string? value = FirstNonEmpty(author);
        if (value is null)
        {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length);
        bool pendingSpace = false;
        foreach (char character in value.Trim().Normalize(NormalizationForm.FormC))
        {
            if (char.IsWhiteSpace(character))
            {
                pendingSpace = builder.Length != 0;
                continue;
            }

            if (pendingSpace)
            {
                builder.Append(' ');
                pendingSpace = false;
            }

            builder.Append(char.ToUpperInvariant(character));
        }

        return builder.ToString();
    }

    private static DocxMarkupBalloonRgb MixRgb(DocxMarkupBalloonRgb source, DocxMarkupBalloonRgb target, double targetWeight)
    {
        double sourceWeight = 1d - targetWeight;
        return new DocxMarkupBalloonRgb(
            (byte)Math.Round(source.Red * sourceWeight + target.Red * targetWeight),
            (byte)Math.Round(source.Green * sourceWeight + target.Green * targetWeight),
            (byte)Math.Round(source.Blue * sourceWeight + target.Blue * targetWeight));
    }


}
