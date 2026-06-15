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
    double AnchorConnectorX = 0d,
    double BalloonConnectorX = 0d,
    bool AnchorConnectorClamped = false,
    int CandidateCount = 1,
    int CommentCandidateCount = 0,
    int RevisionCandidateCount = 0,
    int CommentWithDateCount = 0,
    int CommentResolvedCount = 0,
    int CommentOpenCount = 0,
    int CommentReplyCount = 0,
    int BodySummaryPartCount = 0,
    int WordCompatibleBodySummaryPartCount = 0,
    int CommentSeparatorLineCount = 0,
    int? OverflowStartIndex = null,
    int? OverflowEndIndex = null,
    int LaneBandIndex = 0,
    int LaneBandCandidateCount = 0);

internal sealed class DocxRenderer
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
        IFontResolver? fontResolver = null,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        OoxPdfDocxMarkupGeometryMode markupGeometryMode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout)
    {
        this.fontResolver = fontResolver ?? new WindowsFontResolver();
        markupContext = DocxMarkupContext.FromMode(markupMode, markupGeometryMode);
    }

    public IReadOnlyList<PdfPage> RenderBlankPages(DocxDocument document, Action<OoxPdfDiagnostic>? diagnosticSink = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!HasRenderableContent(document))
        {
            return [new PdfPage(document.PageWidthPoints, document.PageHeightPoints)];
        }

        return RenderParagraphs(document, fontResolver, ResolveEffectiveMarkupContext(document), diagnosticSink, cancellationToken);
    }

    internal DocxLayoutSnapshot InspectLayout(DocxDocument document)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver);
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        OoxPdfDocxMarkupGeometryMode effectiveGeometryMode = ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext);
        DocxLayout layout = new DocxLayoutEngine(effectiveGeometryMode).Create(document, ResolveLayoutTextMeasurer(fontResources, effectiveMarkupContext));
        return DocxLayoutSnapshot.FromLayout(layout, document.MarkupMode, effectiveGeometryMode);
    }

    internal IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> InspectMarkupBalloons(DocxDocument document)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver);
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        DocxLayout layout = new DocxLayoutEngine(ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext)).Create(document, ResolveLayoutTextMeasurer(fontResources, effectiveMarkupContext));
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
        return DocxFontPlanSnapshot.FromPlan(DocxFontPlan.Create(document, fontResolver));
    }

    internal DocxTextEmissionSnapshot InspectTextEmission(DocxDocument document)
    {
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver);
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        DocxLayout layout = new DocxLayoutEngine(ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext)).Create(document, ResolveLayoutTextMeasurer(fontResources, effectiveMarkupContext));
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
                lines.Add(ToTextEmissionLineSnapshot(
                    pageIndex,
                    isStaticStory,
                    ResolveTextEmissionStoryKind(line, fallbackStoryKind),
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

    private static bool HasRenderableContent(DocxDocument document)
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
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        DocxFontResources fontResources = PrepareFontResources(document, fontResolver, cancellationToken);

        DocxLayout layout = new DocxLayoutEngine(ResolveEffectiveMarkupGeometryMode(markupContext)).Create(document, ResolveLayoutTextMeasurer(fontResources, markupContext), cancellationToken);
        double textEmissionFontScale = ResolveTextEmissionFontScale(markupContext);
        double textEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(markupContext);
        double textEmissionXOffset = ResolveTextEmissionXOffset(markupContext);
        bool useWordCompatibleTextProfile = UsesWordCompatibleAllMarkupTextProfile(markupContext);
        bool suppressCommentReferenceSpacer = ShouldSuppressWordCompatibleCommentReferenceSpacer(markupContext);
        IReadOnlyDictionary<string, PdfLinkDestination> bookmarkDestinations = CreateBookmarkDestinations(layout, fontResources, textEmissionFontScale, textEmissionBaselineOffset, textEmissionXOffset, suppressCommentReferenceSpacer, useWordCompatibleTextProfile, cancellationToken);
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

            IReadOnlyList<PdfLinkAnnotation> annotations = CreateHyperlinkAnnotations(layout, layoutPage, pageIndex, fontResources, pageNumber, layout.Pages.Count, bookmarkDestinations, textEmissionFontScale, textEmissionBaselineOffset, textEmissionXOffset, suppressCommentReferenceSpacer, useWordCompatibleTextProfile, cancellationToken);
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
    }

    private static IReadOnlyList<PdfLinkAnnotation> CreateHyperlinkAnnotations(
        DocxLayout layout,
        DocxLayoutPage page,
        int pageIndex,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount,
        IReadOnlyDictionary<string, PdfLinkDestination> bookmarkDestinations,
        double textEmissionFontScale,
        double textEmissionBaselineOffset,
        double textEmissionXOffset,
        bool suppressCommentReferenceSpacer,
        bool useWordCompatibleTextProfile,
        CancellationToken cancellationToken = default)
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
                        link.Target!));
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

    private static IReadOnlyDictionary<string, PdfLinkDestination> CreateBookmarkDestinations(
        DocxLayout layout,
        DocxFontResources fontResources,
        double textEmissionFontScale,
        double textEmissionBaselineOffset,
        double textEmissionXOffset,
        bool suppressCommentReferenceSpacer,
        bool useWordCompatibleTextProfile,
        CancellationToken cancellationToken = default)
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

    private static bool TryResolveBookmarkDestinationSegment(
        IReadOnlyList<DocxTextEmissionSegment> segments,
        DocxBookmarkAnchor bookmark,
        out DocxTextEmissionSegment target,
        out double targetX)
    {
        target = null!;
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

    private static bool IsHyperlinkSegment(DocxHyperlinkSpan link, int sourceTextRunIndex)
    {
        return sourceTextRunIndex >= link.SourceRunStartIndex &&
            sourceTextRunIndex < link.SourceRunStartIndex + link.SourceRunCount;
    }

    private static bool IsExternalHyperlink(DocxHyperlinkSpan link)
    {
        return !string.IsNullOrEmpty(link.Target) &&
            string.Equals(link.TargetMode, "External", StringComparison.OrdinalIgnoreCase);
    }

    private static DocxFontResources PrepareFontResources(DocxDocument document, IFontResolver fontResolver, CancellationToken cancellationToken = default)
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
        CancellationToken cancellationToken = default)
    {
        var resolvedRuns = new List<DocxResolvedRunTypeface>();
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (LoadFont(run.Resolution, fontCache, cancellationToken) is not null)
            {
                resolvedRuns.Add(run);
            }
        }

        foreach (IGrouping<(string StableId, int FaceIndex), DocxResolvedRunTypeface> group in resolvedRuns.GroupBy(run => (run.Resolution!.Source.StableId, run.Resolution.FontFaceIndex)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            FontFaceResolution resolution = group.First().Resolution!;
            IReadOnlyList<int> glyphs = CollectRunGlyphs(group, cancellationToken);
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
            foreach (DocxResolvedRunTypeface run in group)
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
        CancellationToken cancellationToken = default)
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

    private static IReadOnlyList<int> CollectRunGlyphs(IEnumerable<DocxResolvedRunTypeface> runs, CancellationToken cancellationToken = default)
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
        CancellationToken cancellationToken = default)
    {
        foreach (DocxResolvedRunTypeface run in plan.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (LoadFont(run.Resolution, fontCache, cancellationToken) is not null)
            {
                return run.Resolution!;
            }
        }

        return DocxFontFallbackRules.ResolveDefaultDocumentTypeface(fontResolver);
    }

    private static OpenTypeFont? LoadFont(
        FontFaceResolution? resolution,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache,
        CancellationToken cancellationToken = default)
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
    }

    private static bool HasTextLineRevision(DocxTextLineLayout line)
    {
        return CollectTextLineRevisions(line).Count != 0;
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

    private static bool HasTableRowRevision(DocxTableRowLayout row)
    {
        return row.RevisionCount != 0 ||
            row.Table.Revisions?.Count > 0;
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

        if (markupContext.DrawsCommentMarkers && paragraph.InlineReferences.Any(reference => reference.Kind == "Comment"))
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
                graphics.DrawGlyphText(labelResource.Name, labelFontSize, markerX + 1.5d, markerY + 1.4d, 0, 0, 0, glyphHex);
            }
        }
    }

    private static void RenderWordCompatibleCommentRangeMarkers(
        DocxTextLineLayout line,
        DocxParagraph paragraph,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        foreach (DocxInlineReference reference in paragraph.InlineReferences.Where(reference => reference.Kind == "Comment"))
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
    }

    private static bool HasCommentRangeBounds(DocxCommentRange range)
    {
        return range.StartSourceRunIndex is not null ||
            range.StartTextOffset is not null ||
            range.EndSourceRunIndex is not null ||
            range.EndTextOffset is not null;
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
            .Where(reference => reference.Kind == "Comment")
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

    private static void RenderMarkupBalloons(
        DocxLayoutPage page,
        IReadOnlyList<DocxRelatedStoryLayout> relatedStories,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        PdfGraphicsBuilder graphics,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext)
    {
        DocxRunFontResource? labelResource = ResolveMarkupLabelFontResource(fontResources);
        if (labelResource is null)
        {
            return;
        }

        DocxRunFontResource bodyResource = ResolveMarkupBodyFontResource(fontResources) ?? labelResource;
        foreach (DocxMarkupBalloonPlacement placement in BuildMarkupBalloonPlacements(page, relatedStories, floatingDrawings, markupContext))
        {
            RenderMarkupBalloonPlacement(placement, graphics, labelResource, bodyResource, markupContext);
        }
    }

    private static DocxRunFontResource? ResolveMarkupLabelFontResource(DocxFontResources fontResources)
    {
        return fontResources.Fallback ?? fontResources.RunResources.Values.FirstOrDefault();
    }

    private static DocxRunFontResource? ResolveMarkupBodyFontResource(DocxFontResources fontResources)
    {
        return fontResources.RunResources.Values.FirstOrDefault(resource => !resource.Resolution.Bold) ??
            fontResources.Fallback;
    }

    private static IEnumerable<DocxFloatingDrawingLayout> EnumeratePageFloatingDrawings(DocxLayout layout, int pageIndex)
    {
        return layout.FloatingDrawings
            .Concat(layout.StaticFloatingDrawings)
            .Where(drawing => drawing.AnchorPageIndex == pageIndex);
    }

    private static IReadOnlyList<DocxMarkupBalloonPlacement> BuildMarkupBalloonPlacements(
        DocxLayoutPage page,
        IReadOnlyList<DocxRelatedStoryLayout> relatedStories,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        DocxMarkupContext markupContext)
    {
        if (!markupContext.RendersCommentBalloons && !markupContext.RendersRevisionBalloons)
        {
            return [];
        }

        DocxMarkupBalloonArea area = ResolveMarkupBalloonArea(page);
        DocxMarkupBalloonCandidate[] candidates = OrderMarkupBalloonCandidatesForPlacement(
                GroupNearbyMarkupBalloonCandidates(
                    OrderMarkupBalloonCandidatesForPlacement(CollectMarkupBalloonCandidates(page, relatedStories, floatingDrawings, markupContext, area.Width)),
                    area.Width))
            .ToArray();
        if (candidates.Length == 0)
        {
            return [];
        }

        IReadOnlyList<DocxMarkupBalloonLaneBand> laneBands = BuildMarkupBalloonLaneBands(candidates, page, markupContext);
        var placements = new List<DocxMarkupBalloonPlacement>();
        double nextTop = page.Height - page.MarginTop;
        int nextOverflowStartIndex = 1;
        foreach (DocxMarkupBalloonLaneBand laneBand in laneBands)
        {
            double bandTopLimit = Math.Max(laneBand.TopLimit, page.MarginBottom + laneBand.MaxBalloonHeight);
            nextTop = Math.Min(nextTop, bandTopLimit);
            var overflowCandidates = new List<DocxMarkupBalloonCandidate>();
            foreach (DocxMarkupBalloonCandidate candidate in laneBand.Candidates)
            {
                double anchorY = ResolveMarkupBalloonAnchorY(candidate.AnchorY, page, markupContext);
                double height = ResolveMarkupBalloonHeight(candidate, markupContext);
                double topInset = ResolveMarkupBalloonTopInset(markupContext);
                double desiredTop = Math.Min(nextTop, anchorY + topInset);
                double y = desiredTop - height;
                if (y < page.MarginBottom)
                {
                    if (nextTop - height < page.MarginBottom)
                    {
                        overflowCandidates.Add(candidate);
                        continue;
                    }

                    y = page.MarginBottom;
                }

                int nearbyAnchorConnectorCount = placements.Count(placement =>
                    !placement.IsOverflowSummary &&
                    placement.LaneBandIndex == laneBand.Index &&
                    Math.Abs(placement.AnchorY - anchorY) < MarkupBalloonConnectorCollisionAnchorYThresholdPoints);
                double connectorAnchorX = candidate.AnchorConnectorX;
                connectorAnchorX += area.Side == "Left"
                    ? -nearbyAnchorConnectorCount * 1.5d
                    : nearbyAnchorConnectorCount * 1.5d;
                double unclampedConnectorAnchorX = connectorAnchorX;
                connectorAnchorX = ClampMarkupBalloonConnectorAnchorX(connectorAnchorX, page);
                bool anchorConnectorClamped = Math.Abs(connectorAnchorX - unclampedConnectorAnchorX) > 0.001d;
                placements.Add(new DocxMarkupBalloonPlacement(
                    candidate.Kind,
                    area.Side,
                    candidate.Title,
                    candidate.Body,
                    candidate.WordCompatibleTitle,
                    candidate.WordCompatibleBody,
                    area.X,
                    y,
                    area.Width,
                    height,
                    anchorY,
                    connectorAnchorX,
                    area.ConnectorX,
                    anchorConnectorClamped,
                    candidate.FillRgb,
                    candidate.StrokeRgb,
                    candidate.TitleRgb,
                    candidate.BodyRgb,
                    IsOverflowSummary: false,
                    candidate.CandidateCount,
                    candidate.CommentCandidateCount,
                    candidate.RevisionCandidateCount,
                    candidate.CommentWithDateCount,
                    candidate.CommentResolvedCount,
                    candidate.CommentOpenCount,
                    candidate.CommentReplyCount,
                    candidate.BodySummaryPartCount,
                    candidate.WordCompatibleBodySummaryPartCount,
                    LaneBandIndex: laneBand.Index,
                    LaneBandCandidateCount: laneBand.CandidateCount));
                nextTop = y - MarkupBalloonMinimumSpacingPoints;
            }

            (nextTop, nextOverflowStartIndex) = AddOverflowContinuationPlacements(
                overflowCandidates,
                placements,
                area,
                page,
                nextTop,
                laneBand.Index,
                laneBand.CandidateCount,
                nextOverflowStartIndex);
        }

        return placements;
    }

    private static IReadOnlyList<DocxMarkupBalloonLaneBand> BuildMarkupBalloonLaneBands(
        IReadOnlyList<DocxMarkupBalloonCandidate> candidates,
        DocxLayoutPage page,
        DocxMarkupContext markupContext)
    {
        var bands = new List<DocxMarkupBalloonLaneBand>();
        var current = new List<DocxMarkupBalloonCandidate>();
        double currentTopLimit = 0d;
        double currentBottom = 0d;
        double currentMaxBalloonHeight = 0d;

        foreach (DocxMarkupBalloonCandidate candidate in candidates)
        {
            double anchorY = ResolveMarkupBalloonAnchorY(candidate.AnchorY, page, markupContext);
            double height = ResolveMarkupBalloonHeight(candidate, markupContext);
            double preferredTop = Math.Min(page.Height - page.MarginTop, anchorY + ResolveMarkupBalloonTopInset(markupContext));
            double preferredBottom = preferredTop - height;
            if (current.Count != 0 &&
                currentBottom - preferredTop > MarkupBalloonLaneBandSeparationPoints)
            {
                AddCurrentBand();
                current.Clear();
            }

            if (current.Count == 0)
            {
                currentTopLimit = preferredTop;
                currentBottom = preferredBottom;
                currentMaxBalloonHeight = height;
            }
            else
            {
                currentTopLimit = Math.Max(currentTopLimit, preferredTop);
                currentBottom = Math.Min(currentBottom, preferredBottom);
                currentMaxBalloonHeight = Math.Max(currentMaxBalloonHeight, height);
            }

            current.Add(candidate);
        }

        AddCurrentBand();
        return bands;

        void AddCurrentBand()
        {
            if (current.Count == 0)
            {
                return;
            }

            int index = bands.Count;
            bands.Add(new DocxMarkupBalloonLaneBand(
                index,
                current.ToArray(),
                currentTopLimit,
                currentMaxBalloonHeight,
                current.Sum(candidate => candidate.CandidateCount)));
        }
    }

    private static IEnumerable<DocxMarkupBalloonCandidate> OrderMarkupBalloonCandidatesForPlacement(
        IEnumerable<DocxMarkupBalloonCandidate> candidates)
    {
        DocxMarkupBalloonCandidate[] byAnchor = candidates
            .OrderByDescending(candidate => candidate.AnchorY)
            .ThenBy(candidate => candidate.Sequence)
            .ToArray();
        var band = new List<DocxMarkupBalloonCandidate>();
        double bandTopAnchorY = 0d;
        foreach (DocxMarkupBalloonCandidate candidate in byAnchor)
        {
            if (band.Count == 0)
            {
                bandTopAnchorY = candidate.AnchorY;
                band.Add(candidate);
                continue;
            }

            if (bandTopAnchorY - candidate.AnchorY <= MarkupBalloonConnectorCollisionAnchorYThresholdPoints)
            {
                band.Add(candidate);
                continue;
            }

            foreach (DocxMarkupBalloonCandidate orderedCandidate in OrderMarkupBalloonCandidateBand(band))
            {
                yield return orderedCandidate;
            }

            band.Clear();
            bandTopAnchorY = candidate.AnchorY;
            band.Add(candidate);
        }

        foreach (DocxMarkupBalloonCandidate orderedCandidate in OrderMarkupBalloonCandidateBand(band))
        {
            yield return orderedCandidate;
        }
    }

    private static IEnumerable<DocxMarkupBalloonCandidate> OrderMarkupBalloonCandidateBand(
        IReadOnlyList<DocxMarkupBalloonCandidate> band)
    {
        return band
            .OrderBy(candidate => MarkupBalloonKindPriority(candidate.Kind))
            .ThenByDescending(candidate => candidate.AnchorY)
            .ThenBy(candidate => candidate.Sequence);
    }

    private static int MarkupBalloonKindPriority(string kind)
    {
        return kind switch
        {
            "Comment" => 0,
            "Markup" => 1,
            "Revision" => 2,
            "Overflow" => 3,
            _ => 4
        };
    }

    private static double ClampMarkupBalloonConnectorAnchorX(double connectorAnchorX, DocxLayoutPage page)
    {
        return Math.Min(page.Width - 0.5d, Math.Max(0.5d, connectorAnchorX));
    }

    private static double ResolveMarkupBalloonAnchorY(
        double anchorY,
        DocxLayoutPage page,
        DocxMarkupContext markupContext)
    {
        if (markupContext.Mode != OoxPdfDocxMarkupMode.AllMarkup ||
            markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
            !markupContext.ExpandsMarkupMargin)
        {
            return anchorY;
        }

        double resolved = anchorY - WordCompatibleAllMarkupBalloonAnchorYOffsetPoints;
        return Math.Min(page.Height - page.MarginTop, Math.Max(page.MarginBottom, resolved));
    }

    private static double ResolveMarkupBalloonHeight(
        DocxMarkupBalloonCandidate candidate,
        DocxMarkupContext markupContext)
    {
        if (markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin)
        {
            return WordCompatibleAllMarkupBalloonHeightPoints + ResolveWordCompatibleCommentThreadExtraHeight(candidate.CommentReplyCount);
        }

        return string.IsNullOrWhiteSpace(candidate.Body) ? 16d : 26d;
    }

    private static double ResolveWordCompatibleCommentThreadExtraHeight(int replyCount)
    {
        return ResolveCommentThreadSeparatorLineCount(replyCount) * WordCompatibleAllMarkupCommentThreadReplyHeightPoints;
    }

    private static int ResolveCommentThreadSeparatorLineCount(int replyCount)
    {
        return Math.Min(
            WordCompatibleAllMarkupCommentThreadMaxSeparatorLineCount,
            Math.Max(0, replyCount));
    }

    private static double ResolveMarkupBalloonTopInset(DocxMarkupContext markupContext)
    {
        return markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin
            ? WordCompatibleAllMarkupBalloonTopInsetPoints
            : 10d;
    }

    private static IReadOnlyList<DocxMarkupBalloonCandidate> GroupNearbyMarkupBalloonCandidates(
        IEnumerable<DocxMarkupBalloonCandidate> orderedCandidates,
        double textWidth)
    {
        var grouped = new List<DocxMarkupBalloonCandidate>();
        var group = new List<DocxMarkupBalloonCandidate>();
        foreach (DocxMarkupBalloonCandidate candidate in orderedCandidates)
        {
            if (group.Count == 0 ||
                CanGroupMarkupBalloonCandidates(group, candidate))
            {
                group.Add(candidate);
                continue;
            }

            grouped.Add(MergeMarkupBalloonGroup(group, textWidth));
            group.Clear();
            group.Add(candidate);
        }

        if (group.Count != 0)
        {
            grouped.Add(MergeMarkupBalloonGroup(group, textWidth));
        }

        return grouped;
    }

    private static bool CanGroupMarkupBalloonCandidates(IReadOnlyList<DocxMarkupBalloonCandidate> group, DocxMarkupBalloonCandidate candidate)
    {
        DocxMarkupBalloonCandidate previous = group[^1];
        if (Math.Abs(previous.AnchorY - candidate.AnchorY) <= 0.5d)
        {
            return true;
        }

        return group.Count < MarkupBalloonMaxNearbyRevisionGroupSize &&
            group.All(item => item.Kind == "Revision") &&
            previous.Kind == "Revision" &&
            candidate.Kind == "Revision" &&
            Math.Abs(previous.AnchorY - candidate.AnchorY) <= 9d;
    }

    private static DocxMarkupBalloonCandidate MergeMarkupBalloonGroup(
        IReadOnlyList<DocxMarkupBalloonCandidate> group,
        double textWidth)
    {
        if (group.Count == 1)
        {
            return group[0];
        }

        bool allRevisions = group.All(candidate => candidate.Kind == "Revision");
        string title = allRevisions
            ? group.Count.ToString(CultureInfo.InvariantCulture) + " tracked changes"
            : group.Count.ToString(CultureInfo.InvariantCulture) + " markup items";
        string[] bodyParts = BuildMarkupBalloonGroupBodyParts(group);
        string body = TrimBalloonText(string.Join("; ", bodyParts), textWidth);
        string[] wordCompatibleBodyParts = allRevisions
            ? bodyParts
            : BuildWordCompatibleMarkupBalloonGroupBodyParts(group);
        string? wordCompatibleTitle = allRevisions
            ? TrimBalloonText(title, textWidth)
            : group
                .Select(candidate => candidate.WordCompatibleTitle)
                .FirstOrDefault(title => !string.IsNullOrWhiteSpace(title));
        string? wordCompatibleBody = allRevisions
            ? body
            : wordCompatibleBodyParts.Length == 0 ? null : TrimBalloonText(string.Join("; ", wordCompatibleBodyParts), textWidth);
        return group[0] with
        {
            Kind = group.Select(candidate => candidate.Kind).Distinct(StringComparer.Ordinal).Count() == 1
                ? group[0].Kind
                : "Markup",
            Title = TrimBalloonText(title, textWidth),
            Body = body,
            WordCompatibleTitle = wordCompatibleTitle,
            WordCompatibleBody = wordCompatibleBody,
            AnchorY = group.Max(candidate => candidate.AnchorY),
            AnchorConnectorX = allRevisions
                ? group.Select(candidate => candidate.AnchorConnectorX).DefaultIfEmpty(group[0].AnchorConnectorX).Average()
                : group[0].AnchorConnectorX,
            AnchorLeftX = group.Min(candidate => candidate.AnchorLeftX),
            AnchorRightX = group.Max(candidate => candidate.AnchorRightX),
            CandidateCount = group.Sum(candidate => candidate.CandidateCount),
            CommentCandidateCount = group.Sum(candidate => candidate.CommentCandidateCount),
            RevisionCandidateCount = group.Sum(candidate => candidate.RevisionCandidateCount),
            CommentWithDateCount = group.Sum(candidate => candidate.CommentWithDateCount),
            CommentResolvedCount = group.Sum(candidate => candidate.CommentResolvedCount),
            CommentOpenCount = group.Sum(candidate => candidate.CommentOpenCount),
            CommentReplyCount = group.Sum(candidate => candidate.CommentReplyCount),
            BodySummaryPartCount = bodyParts.Length,
            WordCompatibleBodySummaryPartCount = allRevisions ? bodyParts.Length : wordCompatibleBodyParts.Length
        };
    }

    private static string[] BuildMarkupBalloonGroupBodyParts(IReadOnlyList<DocxMarkupBalloonCandidate> group)
    {
        return group
            .Select(BuildMarkupBalloonGroupBodyPart)
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .Distinct(StringComparer.Ordinal)
            .ToArray()!;
    }

    private static string? BuildMarkupBalloonGroupBodyPart(DocxMarkupBalloonCandidate candidate)
    {
        string? body = FirstNonEmpty(candidate.Body);
        string? title = FirstNonEmpty(candidate.Title);
        if (body is null)
        {
            return title;
        }

        return title is null ? body : title + ": " + body;
    }

    private static string[] BuildWordCompatibleMarkupBalloonGroupBodyParts(IReadOnlyList<DocxMarkupBalloonCandidate> group)
    {
        return group
            .Select(BuildWordCompatibleMarkupBalloonGroupBodyPart)
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .Distinct(StringComparer.Ordinal)
            .ToArray()!;
    }

    private static string? BuildWordCompatibleMarkupBalloonGroupBodyPart(DocxMarkupBalloonCandidate candidate)
    {
        string? body = FirstNonEmpty(candidate.WordCompatibleBody, candidate.Body);
        if (body is not null)
        {
            return body;
        }

        return string.Equals(candidate.Kind, "Comment", StringComparison.Ordinal)
            ? null
            : FirstNonEmpty(candidate.WordCompatibleTitle, candidate.Title);
    }

    private static int CountBalloonSummaryPart(string? body)
    {
        return string.IsNullOrWhiteSpace(body) ? 0 : 1;
    }

    private static (double NextTop, int NextOverflowStartIndex) AddOverflowContinuationPlacements(
        IReadOnlyList<DocxMarkupBalloonCandidate> overflowCandidates,
        List<DocxMarkupBalloonPlacement> placements,
        DocxMarkupBalloonArea area,
        DocxLayoutPage page,
        double nextTop,
        int laneBandIndex,
        int laneBandCandidateCount,
        int nextOverflowStartIndex)
    {
        if (overflowCandidates.Count == 0)
        {
            return (nextTop, nextOverflowStartIndex);
        }

        const double height = 12d;
        const double gap = 2d;
        int slotCount = Math.Max(0, (int)Math.Floor((nextTop - page.MarginBottom + gap) / (height + gap)));
        if (slotCount == 0)
        {
            return (nextTop, nextOverflowStartIndex);
        }

        int continuationCount = Math.Min(slotCount, overflowCandidates.Count);
        int chunkSize = (int)Math.Ceiling(overflowCandidates.Count / (double)continuationCount);
        int overflowIndex = 0;
        for (int continuationIndex = 0; continuationIndex < continuationCount && overflowIndex < overflowCandidates.Count; continuationIndex++)
        {
            DocxMarkupBalloonCandidate[] chunk = overflowCandidates
                .Skip(overflowIndex)
                .Take(chunkSize)
                .ToArray();
            if (chunk.Length == 0)
            {
                break;
            }

            double y = nextTop - height;
            if (y < page.MarginBottom)
            {
                break;
            }

            string title = continuationCount == 1
                ? "More markup"
                : "More markup " + (continuationIndex + 1).ToString(CultureInfo.InvariantCulture) + "/" + continuationCount.ToString(CultureInfo.InvariantCulture);
            string body = TrimBalloonText(BuildOverflowContinuationBody(chunk), area.Width);
            placements.Add(new DocxMarkupBalloonPlacement(
                "Overflow",
                area.Side,
                title,
                body,
                null,
                null,
                area.X,
                y,
                area.Width,
                height,
                chunk[0].AnchorY,
                area.ConnectorX,
                area.ConnectorX,
                false,
                new DocxMarkupBalloonRgb(245, 245, 245),
                new DocxMarkupBalloonRgb(120, 120, 120),
                new DocxMarkupBalloonRgb(70, 70, 70),
                new DocxMarkupBalloonRgb(0, 0, 0),
                IsOverflowSummary: true,
                CandidateCount: chunk.Sum(candidate => candidate.CandidateCount),
                CommentCandidateCount: chunk.Sum(candidate => candidate.CommentCandidateCount),
                RevisionCandidateCount: chunk.Sum(candidate => candidate.RevisionCandidateCount),
                CommentWithDateCount: chunk.Sum(candidate => candidate.CommentWithDateCount),
                CommentResolvedCount: chunk.Sum(candidate => candidate.CommentResolvedCount),
                CommentOpenCount: chunk.Sum(candidate => candidate.CommentOpenCount),
                CommentReplyCount: chunk.Sum(candidate => candidate.CommentReplyCount),
                OverflowStartIndex: nextOverflowStartIndex + overflowIndex,
                OverflowEndIndex: nextOverflowStartIndex + overflowIndex + chunk.Length - 1,
                LaneBandIndex: laneBandIndex,
                LaneBandCandidateCount: laneBandCandidateCount));
            overflowIndex += chunk.Length;
            nextTop = y - gap;
        }

        return (nextTop, nextOverflowStartIndex + overflowIndex);
    }

    private static string BuildOverflowContinuationBody(IReadOnlyList<DocxMarkupBalloonCandidate> candidates)
    {
        int count = candidates.Sum(candidate => candidate.CandidateCount);
        string sample = candidates
            .Select(candidate => candidate.Body)
            .FirstOrDefault(body => !string.IsNullOrWhiteSpace(body)) ?? string.Empty;
        return string.IsNullOrWhiteSpace(sample)
            ? count.ToString(CultureInfo.InvariantCulture)
            : count.ToString(CultureInfo.InvariantCulture) + ": " + sample;
    }

    private static DocxCommentThreadBalloonMetrics CountCommentThreadBalloonMetrics(
        DocxRelatedStoryLayout? storyLayout,
        IReadOnlyList<DocxRelatedStoryLayout> replies)
    {
        int withDateCount = 0;
        int resolvedCount = 0;
        int openCount = 0;

        Count(storyLayout?.Story.CommentMetadata);
        foreach (DocxRelatedStoryLayout reply in replies)
        {
            Count(reply.Story.CommentMetadata);
        }

        return new DocxCommentThreadBalloonMetrics(withDateCount, resolvedCount, openCount, replies.Count);

        void Count(DocxCommentMetadata? metadata)
        {
            if (metadata is null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(metadata.Date))
            {
                withDateCount++;
            }

            if (metadata.IsResolved == true)
            {
                resolvedCount++;
            }
            else if (metadata.IsResolved == false)
            {
                openCount++;
            }
        }
    }

    private static IReadOnlyList<DocxMarkupBalloonCandidate> CollectMarkupBalloonCandidates(
        DocxLayoutPage page,
        IReadOnlyList<DocxRelatedStoryLayout> relatedStories,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        DocxMarkupContext markupContext,
        double textWidth)
    {
        var candidates = new List<DocxMarkupBalloonCandidate>();
        Dictionary<string, DocxRelatedStoryLayout> commentStories = relatedStories
            .Where(story => story.Story.Kind == "Comment" && story.Story.Id is not null)
            .GroupBy(story => story.Story.Id!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        Dictionary<string, DocxRelatedStoryLayout[]> commentRepliesByParentId = relatedStories
            .Where(story => story.Story.Kind == "Comment" && story.Story.CommentMetadata?.ParentCommentId is not null)
            .GroupBy(story => story.Story.CommentMetadata!.ParentCommentId!, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => group
                    .OrderBy(story => FormatCommentDate(story.Story.CommentMetadata?.Date), StringComparer.Ordinal)
                    .ThenBy(story => story.Story.Id, StringComparer.Ordinal)
                    .ToArray(),
                StringComparer.Ordinal);
        var renderedComments = new HashSet<string>(StringComparer.Ordinal);
        var renderedRevisions = new HashSet<int>();
        var renderedTableRevisions = new HashSet<string>(StringComparer.Ordinal);
        var renderedTableRowRevisions = new HashSet<string>(StringComparer.Ordinal);
        var renderedTableCellRevisions = new HashSet<string>(StringComparer.Ordinal);
        DocxTextLineLayout[] anchorTextLines = EnumerateMarkupBalloonAnchorTextLines(page, floatingDrawings).ToArray();
        int sequence = 0;
        foreach (DocxTextLineLayout line in anchorTextLines)
        {
            if (line.SourceParagraph is not { } paragraph)
            {
                continue;
            }

            if (markupContext.RendersCommentBalloons)
            {
                foreach (DocxInlineReference reference in paragraph.InlineReferences.Where(reference => reference.Kind == "Comment"))
                {
                    string key = RuntimeHelpers.GetHashCode(paragraph).ToString(CultureInfo.InvariantCulture) + ":" + (reference.Id ?? string.Empty);
                    if (!renderedComments.Add(key))
                    {
                        continue;
                    }

                    commentStories.TryGetValue(reference.Id ?? string.Empty, out DocxRelatedStoryLayout? storyLayout);
                    commentRepliesByParentId.TryGetValue(reference.Id ?? string.Empty, out DocxRelatedStoryLayout[]? replies);
                    string commentBody = TrimBalloonText(BuildCommentBalloonPreview(storyLayout, replies ?? []), textWidth);
                    string wordCompatibleCommentBody = TrimBalloonText(BuildWordCompatibleCommentBalloonPreview(storyLayout, replies ?? []), textWidth);
                    DocxCommentThreadBalloonMetrics commentMetrics = CountCommentThreadBalloonMetrics(storyLayout, replies ?? []);
                    DocxTextLineLayout anchorLine = ResolveCommentAnchorLine(line, anchorTextLines, paragraph, reference);
                    candidates.Add(new DocxMarkupBalloonCandidate(
                        "Comment",
                        TrimBalloonText(BuildCommentBalloonTitle(storyLayout?.Story, reference.Id), textWidth),
                        commentBody,
                        BuildWordCompatibleCommentBalloonTitle(storyLayout?.Story, reference.Id),
                        wordCompatibleCommentBody,
                        anchorLine.BaselineY,
                        ResolveCommentAnchorX(anchorLine, anchorTextLines, paragraph, reference, markupContext),
                        anchorLine.X - 2d,
                        anchorLine.X + Math.Max(0d, anchorLine.Width) + 2d,
                        sequence++,
                        new DocxMarkupBalloonRgb(255, 250, 220),
                        new DocxMarkupBalloonRgb(217, 151, 0),
                        new DocxMarkupBalloonRgb(70, 70, 70),
                        new DocxMarkupBalloonRgb(0, 0, 0),
                        CommentCandidateCount: 1,
                        CommentWithDateCount: commentMetrics.WithDateCount,
                        CommentResolvedCount: commentMetrics.ResolvedCount,
                        CommentOpenCount: commentMetrics.OpenCount,
                        CommentReplyCount: commentMetrics.ReplyCount,
                        BodySummaryPartCount: CountBalloonSummaryPart(commentBody),
                        WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(commentBody)));
                }
            }

            if (markupContext.RendersRevisionBalloons && paragraph.Revisions.Count != 0)
            {
                int key = RuntimeHelpers.GetHashCode(paragraph);
                if (!renderedRevisions.Add(key))
                {
                    continue;
                }

                DocxRevisionMarkupPalette paragraphRevisionPalette = ResolveRevisionMarkupPalette(paragraph.Revisions);
                string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(paragraph.Revisions), textWidth);
                string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(paragraph), textWidth);
                candidates.Add(new DocxMarkupBalloonCandidate(
                    "Revision",
                    revisionTitle,
                    revisionBody,
                    revisionTitle,
                    revisionBody,
                    line.BaselineY,
                    line.X + Math.Max(0d, line.Width) * 0.5d,
                    line.X - 2d,
                    line.X + Math.Max(0d, line.Width) + 2d,
                    sequence++,
                    paragraphRevisionPalette.FillRgb,
                    paragraphRevisionPalette.StrokeRgb,
                    paragraphRevisionPalette.TitleRgb,
                    new DocxMarkupBalloonRgb(0, 0, 0),
                    RevisionCandidateCount: 1,
                    BodySummaryPartCount: CountBalloonSummaryPart(revisionBody),
                    WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(revisionBody)));
            }
        }

        if (markupContext.RendersRevisionBalloons)
        {
            foreach (DocxTableRowLayout row in EnumerateMarkupBalloonTableRows(page, floatingDrawings))
            {
                IReadOnlyList<DocxRevisionInfo> tableRevisions = row.Table.Revisions ?? [];
                if (tableRevisions.Count != 0)
                {
                    string tableKey = TableBalloonKey(row);
                    if (renderedTableRevisions.Add(tableKey))
                    {
                        DocxRevisionMarkupPalette tableRevisionPalette = ResolveRevisionMarkupPalette(tableRevisions);
                        string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(tableRevisions), textWidth);
                        string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(tableRevisions), textWidth);
                        candidates.Add(new DocxMarkupBalloonCandidate(
                            "Revision",
                            revisionTitle,
                            revisionBody,
                            revisionTitle,
                            revisionBody,
                            row.Y + Math.Max(0d, row.Height) * 0.5d,
                            row.Table.TableX + Math.Max(0d, row.Table.ResolvedTableWidth) * 0.5d,
                            row.Table.TableX - 2d,
                            row.Table.TableX + Math.Max(0d, row.Table.ResolvedTableWidth) + 2d,
                            sequence++,
                            tableRevisionPalette.FillRgb,
                            tableRevisionPalette.StrokeRgb,
                            tableRevisionPalette.TitleRgb,
                            new DocxMarkupBalloonRgb(0, 0, 0),
                            RevisionCandidateCount: 1,
                            BodySummaryPartCount: CountBalloonSummaryPart(revisionBody),
                            WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(revisionBody)));
                    }
                }

                IReadOnlyList<DocxRevisionInfo> rowRevisions = row.Revisions ?? [];
                if (rowRevisions.Count != 0)
                {
                    string rowKey = TableRowBalloonKey(row);
                    if (renderedTableRowRevisions.Add(rowKey))
                    {
                        DocxRevisionMarkupPalette rowRevisionPalette = ResolveRevisionMarkupPalette(rowRevisions);
                        string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(rowRevisions), textWidth);
                        string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(rowRevisions), textWidth);
                        candidates.Add(new DocxMarkupBalloonCandidate(
                            "Revision",
                            revisionTitle,
                            revisionBody,
                            revisionTitle,
                            revisionBody,
                            row.Y + Math.Max(0d, row.Height) * 0.5d,
                            row.Table.TableX + Math.Max(0d, row.Table.ResolvedTableWidth) * 0.5d,
                            row.Table.TableX - 2d,
                            row.Table.TableX + Math.Max(0d, row.Table.ResolvedTableWidth) + 2d,
                            sequence++,
                            rowRevisionPalette.FillRgb,
                            rowRevisionPalette.StrokeRgb,
                            rowRevisionPalette.TitleRgb,
                            new DocxMarkupBalloonRgb(0, 0, 0),
                            RevisionCandidateCount: 1,
                            BodySummaryPartCount: CountBalloonSummaryPart(revisionBody),
                            WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(revisionBody)));
                    }
                }

                for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
                {
                    DocxTableCellLayout cell = row.Cells[cellIndex];
                    IReadOnlyList<DocxRevisionInfo> cellRevisions = cell.Cell.Revisions;
                    if (cellRevisions.Count == 0)
                    {
                        continue;
                    }

                    string cellKey = TableCellBalloonKey(row, cellIndex);
                    if (!renderedTableCellRevisions.Add(cellKey))
                    {
                        continue;
                    }

                    DocxRevisionMarkupPalette cellRevisionPalette = ResolveRevisionMarkupPalette(cellRevisions);
                    string revisionTitle = TrimBalloonText(BuildRevisionBalloonTitle(cellRevisions), textWidth);
                    string revisionBody = TrimBalloonText(BuildRevisionBalloonPreview(cellRevisions), textWidth);
                    candidates.Add(new DocxMarkupBalloonCandidate(
                        "Revision",
                        revisionTitle,
                        revisionBody,
                        revisionTitle,
                        revisionBody,
                        cell.Y + Math.Max(0d, cell.Height) * 0.5d,
                        cell.X + Math.Max(0d, cell.Width) * 0.5d,
                        cell.X - 2d,
                        cell.X + Math.Max(0d, cell.Width) + 2d,
                        sequence++,
                        cellRevisionPalette.FillRgb,
                        cellRevisionPalette.StrokeRgb,
                        cellRevisionPalette.TitleRgb,
                        new DocxMarkupBalloonRgb(0, 0, 0),
                        RevisionCandidateCount: 1,
                        BodySummaryPartCount: CountBalloonSummaryPart(revisionBody),
                        WordCompatibleBodySummaryPartCount: CountBalloonSummaryPart(revisionBody)));
                }
            }
        }

        return candidates;
    }

    private static DocxTextLineLayout ResolveCommentAnchorLine(
        DocxTextLineLayout fallbackLine,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph,
        DocxInlineReference reference)
    {
        DocxCommentRange? range = paragraph.CommentRanges.FirstOrDefault(range =>
            string.Equals(range.Id, reference.Id, StringComparison.Ordinal));
        if (range is not null &&
            TryResolveCommentRangeEndAnchor(fallbackLine, anchorTextLines, paragraph, range, out DocxTextLineLayout anchorLine, out _))
        {
            return anchorLine;
        }

        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(fallbackLine, anchorTextLines, paragraph))
        {
            if (TryResolveSourceOffsetAnchorX(candidateLine, reference.SourceRunIndex, reference.TextOffsetInRun, out _) ||
                TryResolvePreviousSourceRunEndAnchorX(candidateLine, reference.SourceRunIndex, minimumSourceRunIndex: null, out _))
            {
                return candidateLine;
            }
        }

        return fallbackLine;
    }

    private static bool TryResolveCommentRangeEndAnchor(
        DocxTextLineLayout line,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph,
        DocxCommentRange range,
        out DocxTextLineLayout anchorLine,
        out double anchorX)
    {
        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(line, anchorTextLines, paragraph))
        {
            if (TryResolveSourceOffsetAnchorX(candidateLine, range.EndSourceRunIndex, range.EndTextOffset, out anchorX))
            {
                anchorLine = candidateLine;
                return true;
            }
        }

        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(line, anchorTextLines, paragraph))
        {
            if (TryResolvePreviousSourceRunEndAnchorX(candidateLine, range.EndSourceRunIndex, range.StartSourceRunIndex, out anchorX))
            {
                anchorLine = candidateLine;
                return true;
            }
        }

        foreach (DocxTextLineLayout candidateLine in EnumerateCommentAnchorSearchLines(line, anchorTextLines, paragraph))
        {
            if (TryResolveSourceRunEndAnchorX(candidateLine, range.StartSourceRunIndex, out anchorX))
            {
                anchorLine = candidateLine;
                return true;
            }
        }

        anchorLine = line;
        anchorX = 0d;
        return false;
    }

    private static double ResolveCommentAnchorX(
        DocxTextLineLayout line,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph,
        DocxInlineReference reference,
        DocxMarkupContext markupContext)
    {
        DocxCommentRange? range = paragraph.CommentRanges.FirstOrDefault(range =>
            string.Equals(range.Id, reference.Id, StringComparison.Ordinal));
        if (range is not null)
        {
            if (UsesWordCompatibleAllMarkupTextProfile(markupContext) &&
                TryResolveCommentRangeEndAnchor(line, anchorTextLines, paragraph, range, out _, out double wordCompatibleRangeEndX))
            {
                return wordCompatibleRangeEndX +
                    WordCompatibleAllMarkupTextXOffsetPoints -
                    WordCompatibleAllMarkupConnectorBodyAnchorInsetPoints;
            }

            if (TryResolveSourceOffsetAnchorX(line, range.StartSourceRunIndex, range.StartTextOffset, out double startX))
            {
                return startX;
            }

            if (TryResolveSourceOffsetAnchorX(line, range.EndSourceRunIndex, range.EndTextOffset, out double endX))
            {
                return endX;
            }
        }

        return ResolveInlineReferenceAnchorX(line, reference);
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateCommentAnchorSearchLines(
        DocxTextLineLayout line,
        IReadOnlyList<DocxTextLineLayout> anchorTextLines,
        DocxParagraph paragraph)
    {
        yield return line;
        foreach (DocxTextLineLayout candidateLine in anchorTextLines)
        {
            if (!ReferenceEquals(candidateLine, line) &&
                ReferenceEquals(candidateLine.SourceParagraph, paragraph))
            {
                yield return candidateLine;
            }
        }
    }

    private static double ResolveInlineReferenceAnchorX(DocxTextLineLayout line, DocxInlineReference reference)
    {
        if (reference.SourceRunIndex < 0)
        {
            return line.X + Math.Max(0d, line.Width) * 0.5d;
        }

        if (TryResolveSourceOffsetAnchorX(line, reference.SourceRunIndex, reference.TextOffsetInRun, out double anchorX))
        {
            return anchorX;
        }

        return line.X + Math.Max(0d, line.Width) * 0.5d;
    }

    private static bool TryResolveSourceOffsetAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        int? textOffsetInRun,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex || runIndex < 0)
        {
            return false;
        }

        int offset = Math.Max(0, textOffsetInRun ?? 0);
        foreach (DocxTextSegmentLayout segment in line.Segments)
        {
            if (segment.SourceTextRunIndex != runIndex)
            {
                continue;
            }

            int segmentStart = Math.Max(0, segment.SourceTextOffsetInRun);
            int segmentEnd = segmentStart + segment.Text.Length;
            if (offset < segmentStart || offset > segmentEnd)
            {
                continue;
            }

            if (segment.Text.Length == 0)
            {
                anchorX = segment.X;
                return true;
            }

            double ratio = Math.Clamp((offset - segmentStart) / (double)segment.Text.Length, 0d, 1d);
            anchorX = segment.X + Math.Max(0d, segment.Width) * ratio;
            return true;
        }

        return false;
    }

    private static bool TryResolveSourceRunStartAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex)
        {
            return false;
        }

        DocxTextSegmentLayout[] segments = line.Segments
            .Where(segment => segment.SourceTextRunIndex == runIndex)
            .ToArray();
        if (segments.Length == 0)
        {
            return false;
        }

        anchorX = segments.Min(segment => segment.X);
        return true;
    }

    private static bool TryResolvePreviousSourceRunEndAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        int? minimumSourceRunIndex,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex)
        {
            return false;
        }

        int previousRunIndex = line.Segments
            .Where(segment =>
                segment.SourceTextRunIndex < runIndex &&
                (minimumSourceRunIndex is null || segment.SourceTextRunIndex >= minimumSourceRunIndex.Value))
            .Select(segment => segment.SourceTextRunIndex)
            .DefaultIfEmpty(-1)
            .Max();
        return previousRunIndex >= 0 &&
            TryResolveSourceRunEndAnchorX(line, previousRunIndex, out anchorX);
    }

    private static bool TryResolveSourceRunEndAnchorX(
        DocxTextLineLayout line,
        int? sourceRunIndex,
        out double anchorX)
    {
        anchorX = 0d;
        if (sourceRunIndex is not { } runIndex)
        {
            return false;
        }

        DocxTextSegmentLayout[] segments = line.Segments
            .Where(segment => segment.SourceTextRunIndex == runIndex)
            .ToArray();
        if (segments.Length == 0)
        {
            return false;
        }

        anchorX = segments.Max(segment => segment.X + Math.Max(0d, segment.Width));
        return true;
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateMarkupBalloonTableRows(
        DocxLayoutPage page,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings)
    {
        foreach (DocxTableRowLayout row in EnumerateTableRows(page.StaticTableRows))
        {
            yield return row;
        }

        foreach (DocxTableRowLayout row in EnumerateTableRows(page.Items))
        {
            yield return row;
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTableRowLayout row in EnumerateTableRows(story.TableRows))
            {
                yield return row;
            }
        }

        foreach (DocxTableRowLayout row in EnumerateFloatingDrawingTextBoxTableRows(floatingDrawings))
        {
            yield return row;
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTableRowLayout row in EnumerateFloatingDrawingTextBoxTableRows(story.FloatingDrawings))
            {
                yield return row;
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateTableRows(IEnumerable<DocxLayoutItem> items)
    {
        foreach (DocxLayoutItem item in items)
        {
            if (item is DocxTableRowLayout row)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(row))
                {
                    yield return nested;
                }
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateTableRows(IEnumerable<DocxTableRowLayout> rows)
    {
        foreach (DocxTableRowLayout row in rows)
        {
            foreach (DocxTableRowLayout nested in EnumerateTableRows(row))
            {
                yield return nested;
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateTableRows(DocxTableRowLayout row)
    {
        yield return row;
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(nestedRow))
                {
                    yield return nested;
                }
            }
        }
    }

    private static string TableBalloonKey(DocxTableRowLayout row)
    {
        return (row.StoryKind ?? string.Empty) +
            ":" + (row.StoryVariantType ?? string.Empty) +
            ":" + row.Table.SourceBlockIndex.ToString(CultureInfo.InvariantCulture) +
            ":" + row.Table.TableIndex.ToString(CultureInfo.InvariantCulture);
    }

    private static string TableRowBalloonKey(DocxTableRowLayout row)
    {
        return TableBalloonKey(row) +
            ":" + row.RowIndex.ToString(CultureInfo.InvariantCulture);
    }

    private static string TableCellBalloonKey(DocxTableRowLayout row, int cellIndex)
    {
        return TableRowBalloonKey(row) +
            ":" + cellIndex.ToString(CultureInfo.InvariantCulture);
    }

    private static DocxMarkupBalloonArea ResolveMarkupBalloonArea(DocxLayoutPage page)
    {
        const double mediaInset = 2d;
        double leftAvailable = Math.Max(0d, page.MarginLeft - 8d);
        double rightAvailable = Math.Max(0d, page.MarginRight - 8d);
        bool useLeft = leftAvailable > rightAvailable && leftAvailable >= 24d;
        double laneWidth = Math.Max(MinimumMarkupBalloonBodyWidthPoints, useLeft ? leftAvailable : rightAvailable);
        double stemWidth = Math.Min(WordMarkupBalloonStemWidthPoints, Math.Max(0d, laneWidth - MinimumMarkupBalloonBodyWidthPoints));
        double mediaBodyWidth = Math.Max(1d, page.Width - mediaInset * 2d);
        double bodyWidth = Math.Min(mediaBodyWidth, Math.Max(MinimumMarkupBalloonBodyWidthPoints, laneWidth - stemWidth));
        if (useLeft)
        {
            double leftLaneX = Math.Max(mediaInset, page.MarginLeft - laneWidth - 4d);
            double bodyX = ClampMarkupBalloonBodyX(leftLaneX, bodyWidth, page.Width, mediaInset);
            double connectorX = Math.Min(page.Width - mediaInset, bodyX + bodyWidth + stemWidth);
            return new DocxMarkupBalloonArea("Left", bodyX, bodyWidth, connectorX);
        }

        double rightLaneX = Math.Min(page.Width - laneWidth - 2d, page.Width - page.MarginRight + 4d);
        double rightBodyX = ClampMarkupBalloonBodyX(rightLaneX + stemWidth, bodyWidth, page.Width, mediaInset);
        double rightConnectorX = Math.Max(mediaInset, rightBodyX - stemWidth);
        return new DocxMarkupBalloonArea("Right", rightBodyX, bodyWidth, rightConnectorX);
    }

    private static double ClampMarkupBalloonBodyX(double x, double width, double pageWidth, double mediaInset)
    {
        double maxX = Math.Max(mediaInset, pageWidth - mediaInset - width);
        return Math.Min(maxX, Math.Max(mediaInset, x));
    }

    private static void RenderMarkupBalloonPlacement(
        DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics,
        DocxRunFontResource labelResource,
        DocxRunFontResource bodyResource,
        DocxMarkupContext markupContext)
    {
        DocxMarkupBalloonRgb fillRgb = ResolveMarkupBalloonBodyFillRgb(placement, markupContext);
        DocxMarkupBalloonRgb strokeRgb = ResolveMarkupBalloonBodyStrokeRgb(placement, markupContext);
        graphics.SetFillRgb(fillRgb.Red, fillRgb.Green, fillRgb.Blue);
        graphics.SetStrokeRgb(strokeRgb.Red, strokeRgb.Green, strokeRgb.Blue);
        graphics.SetLineWidth(ResolveMarkupBalloonBodyStrokeWidth(markupContext));
        graphics.FillStrokeRectangleEvenOdd(placement.X, placement.Y, placement.Width, placement.Height);
        if (!placement.IsOverflowSummary)
        {
            RenderMarkupBalloonConnector(placement, graphics, markupContext);
        }

        if (ShouldRenderWordCompatibleBalloonText(placement, markupContext))
        {
            RenderWordCompatibleCommentThreadSeparators(placement, graphics);
            RenderWordCompatibleBalloonText(placement, graphics, labelResource, bodyResource);
            return;
        }

        DrawBalloonText(graphics, labelResource, placement.Title, placement.X + 3d, placement.Y + placement.Height - 7d, 5.5d, placement.TitleRgb.Red, placement.TitleRgb.Green, placement.TitleRgb.Blue);
        if (!string.IsNullOrWhiteSpace(placement.Body))
        {
            DrawBalloonText(graphics, labelResource, placement.Body, placement.X + 3d, placement.Y + 4d, 5d, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
        }
    }

    private static DocxMarkupBalloonRgb ResolveMarkupBalloonBodyFillRgb(
        DocxMarkupBalloonPlacement placement,
        DocxMarkupContext markupContext)
    {
        return !placement.IsOverflowSummary && UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? WordCompatibleAllMarkupReviewFillRgb
            : placement.FillRgb;
    }

    private static DocxMarkupBalloonRgb ResolveMarkupBalloonBodyStrokeRgb(
        DocxMarkupBalloonPlacement placement,
        DocxMarkupContext markupContext)
    {
        return !placement.IsOverflowSummary && UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? WordCompatibleAllMarkupReviewStrokeRgb
            : placement.StrokeRgb;
    }

    private static double ResolveMarkupBalloonBodyStrokeWidth(DocxMarkupContext markupContext)
    {
        return UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? WordCompatibleAllMarkupConnectorStrokeWidthPoints
            : 0.5d;
    }

    private static bool ShouldRenderWordCompatibleBalloonText(
        DocxMarkupBalloonPlacement placement,
        DocxMarkupContext markupContext)
    {
        return !placement.IsOverflowSummary &&
            markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin &&
            !string.IsNullOrWhiteSpace(placement.WordCompatibleTitle);
    }

    private static void RenderWordCompatibleBalloonText(
        DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics,
        DocxRunFontResource labelResource,
        DocxRunFontResource bodyResource)
    {
        string title = placement.WordCompatibleTitle ?? placement.Title;
        string body = placement.WordCompatibleBody ?? string.Empty;
        double fontSize = WordCompatibleAllMarkupBalloonTextFontSizePoints;
        double textX = placement.X + WordCompatibleAllMarkupBalloonTextInsetXPoints;
        double firstBaselineY = ResolveWordCompatibleBalloonFirstBaselineY(placement);
        double lineGap = fontSize * 1.2d;
        const byte titleRgb = 0;

        DrawBalloonText(
            graphics,
            labelResource,
            title,
            textX,
            firstBaselineY,
            fontSize,
            titleRgb,
            titleRgb,
            titleRgb,
            WordCompatibleAllMarkupBalloonTitlePositioningCharacterSpacingPoints);
        if (string.IsNullOrWhiteSpace(body))
        {
            return;
        }

        double titleWidth = labelResource.Embedded.MeasureTextPoints(title, fontSize);
        double bodyFirstLineX = textX + titleWidth + WordCompatibleAllMarkupBalloonBodyFirstLineXOffsetPoints;
        double rightEdge = placement.X + placement.Width - 0.5d;
        double firstLineWidth = Math.Max(0d, rightEdge - bodyFirstLineX);
        double continuationWidth = Math.Max(0d, rightEdge - textX);
        string[] lines = WrapWordCompatibleBalloonBody(body, bodyResource.Embedded, fontSize, firstLineWidth, continuationWidth);
        if (lines.Length == 0)
        {
            return;
        }

        DrawBalloonText(graphics, bodyResource, lines[0], bodyFirstLineX, firstBaselineY, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
        if (lines.Length == 1)
        {
            DrawBalloonText(graphics, bodyResource, " ", bodyFirstLineX + labelResource.Embedded.MeasureTextPoints(lines[0], fontSize), firstBaselineY, fontSize, placement.BodyRgb.Red, placement.BodyRgb.Green, placement.BodyRgb.Blue);
            return;
        }

        double secondBaselineY = firstBaselineY - lineGap;
        DrawBalloonText(
            graphics,
            bodyResource,
            lines[1],
            textX,
            secondBaselineY,
            fontSize,
            placement.BodyRgb.Red,
            placement.BodyRgb.Green,
            placement.BodyRgb.Blue,
            WordCompatibleAllMarkupBalloonContinuationPositioningCharacterSpacingPoints);
        DrawBalloonText(
            graphics,
            bodyResource,
            " ",
            textX + labelResource.Embedded.MeasureTextPoints(lines[1], fontSize) + WordCompatibleAllMarkupBalloonContinuationTerminalSpaceXOffsetPoints,
            secondBaselineY,
            fontSize,
            placement.BodyRgb.Red,
            placement.BodyRgb.Green,
            placement.BodyRgb.Blue);
    }

    private static double ResolveWordCompatibleBalloonFirstBaselineY(DocxMarkupBalloonPlacement placement)
    {
        return placement.Y + placement.Height - WordCompatibleAllMarkupBalloonFirstBaselineTopInsetPoints;
    }

    private static void RenderWordCompatibleCommentThreadSeparators(
        DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics)
    {
        int separatorCount = ResolveCommentThreadSeparatorLineCount(placement.CommentReplyCount);
        if (separatorCount == 0)
        {
            return;
        }

        double leftX = placement.X + WordCompatibleAllMarkupBalloonTextInsetXPoints;
        double rightX = placement.X + placement.Width - WordCompatibleAllMarkupBalloonTextInsetXPoints;
        double lineGap = WordCompatibleAllMarkupCommentThreadReplyHeightPoints;
        double firstSeparatorY = ResolveWordCompatibleBalloonFirstBaselineY(placement) -
            lineGap -
            WordCompatibleAllMarkupCommentThreadSeparatorYOffsetPoints;
        graphics.ClearLineDash();
        graphics.SetLineWidth(WordCompatibleAllMarkupConnectorStrokeWidthPoints);
        graphics.SetStrokeRgb(WordCompatibleAllMarkupReviewStrokeRgb.Red, WordCompatibleAllMarkupReviewStrokeRgb.Green, WordCompatibleAllMarkupReviewStrokeRgb.Blue);
        for (int i = 0; i < separatorCount; i++)
        {
            double separatorY = firstSeparatorY - i * lineGap;
            if (separatorY <= placement.Y + 1d)
            {
                break;
            }

            graphics.StrokeLine(leftX, separatorY, rightX, separatorY);
        }
    }

    private static void RenderMarkupBalloonConnector(
        DocxMarkupBalloonPlacement placement,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        if (markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin)
        {
            double bodyConnectorX = placement.Side == "Left"
                ? placement.X + placement.Width
                : placement.X;
            double bodyConnectorY = placement.Y + placement.Height * 0.775d;
            graphics.SetLineWidth(WordCompatibleAllMarkupConnectorStrokeWidthPoints);
            graphics.SetLineDash(WordCompatibleAllMarkupConnectorStrokeWidthPoints, WordCompatibleAllMarkupConnectorStrokeWidthPoints);
            graphics.StrokeLine(placement.AnchorConnectorX, placement.AnchorY, placement.BalloonConnectorX, placement.AnchorY);
            graphics.StrokeLine(placement.BalloonConnectorX, placement.AnchorY, bodyConnectorX, bodyConnectorY);
            graphics.ClearLineDash();
            return;
        }

        graphics.StrokeLine(placement.AnchorConnectorX, placement.AnchorY, placement.BalloonConnectorX, placement.Y + placement.Height * 0.5d);
    }

    private readonly record struct DocxMarkupBalloonRgb(byte Red, byte Green, byte Blue);

    private readonly record struct DocxRevisionMarkupPalette(
        DocxMarkupBalloonRgb FillRgb,
        DocxMarkupBalloonRgb StrokeRgb,
        DocxMarkupBalloonRgb TitleRgb);

    private readonly record struct DocxCommentThreadBalloonMetrics(
        int WithDateCount,
        int ResolvedCount,
        int OpenCount,
        int ReplyCount);

    private sealed record DocxMarkupBalloonArea(
        string Side,
        double X,
        double Width,
        double ConnectorX);

    private sealed record DocxMarkupBalloonCandidate(
        string Kind,
        string Title,
        string Body,
        string? WordCompatibleTitle,
        string? WordCompatibleBody,
        double AnchorY,
        double AnchorConnectorX,
        double AnchorLeftX,
        double AnchorRightX,
        int Sequence,
        DocxMarkupBalloonRgb FillRgb,
        DocxMarkupBalloonRgb StrokeRgb,
        DocxMarkupBalloonRgb TitleRgb,
        DocxMarkupBalloonRgb BodyRgb,
        int CandidateCount = 1,
        int CommentCandidateCount = 0,
        int RevisionCandidateCount = 0,
        int CommentWithDateCount = 0,
        int CommentResolvedCount = 0,
        int CommentOpenCount = 0,
        int CommentReplyCount = 0,
        int BodySummaryPartCount = 0,
        int WordCompatibleBodySummaryPartCount = 0);

    private sealed record DocxMarkupBalloonLaneBand(
        int Index,
        IReadOnlyList<DocxMarkupBalloonCandidate> Candidates,
        double TopLimit,
        double MaxBalloonHeight,
        int CandidateCount);

    private sealed record DocxMarkupBalloonPlacement(
        string Kind,
        string Side,
        string Title,
        string Body,
        string? WordCompatibleTitle,
        string? WordCompatibleBody,
        double X,
        double Y,
        double Width,
        double Height,
        double AnchorY,
        double AnchorConnectorX,
        double BalloonConnectorX,
        bool AnchorConnectorClamped,
        DocxMarkupBalloonRgb FillRgb,
        DocxMarkupBalloonRgb StrokeRgb,
        DocxMarkupBalloonRgb TitleRgb,
        DocxMarkupBalloonRgb BodyRgb,
        bool IsOverflowSummary,
        int CandidateCount = 1,
        int CommentCandidateCount = 0,
        int RevisionCandidateCount = 0,
        int CommentWithDateCount = 0,
        int CommentResolvedCount = 0,
        int CommentOpenCount = 0,
        int CommentReplyCount = 0,
        int BodySummaryPartCount = 0,
        int WordCompatibleBodySummaryPartCount = 0,
        int? OverflowStartIndex = null,
        int? OverflowEndIndex = null,
        int LaneBandIndex = 0,
        int LaneBandCandidateCount = 0)
    {
        public DocxMarkupBalloonPlacementSnapshot ToSnapshot(int pageIndex)
        {
            return new DocxMarkupBalloonPlacementSnapshot(
                pageIndex,
                Kind,
                Side,
                X,
                Y,
                Width,
                Height,
                AnchorY,
                IsOverflowSummary,
                AnchorConnectorX,
                BalloonConnectorX,
                AnchorConnectorClamped,
                CandidateCount,
                CommentCandidateCount,
                RevisionCandidateCount,
                CommentWithDateCount,
                CommentResolvedCount,
                CommentOpenCount,
                CommentReplyCount,
                BodySummaryPartCount,
                WordCompatibleBodySummaryPartCount,
                ResolveCommentThreadSeparatorLineCount(CommentReplyCount),
                OverflowStartIndex,
                OverflowEndIndex,
                LaneBandIndex,
                LaneBandCandidateCount);
        }
    }

    internal static string BuildRevisionBalloonPreview(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        return BuildRevisionBalloonPreview(revisions, []);
    }

    internal static string BuildRevisionBalloonPreview(DocxParagraph paragraph)
    {
        return BuildRevisionBalloonPreview(paragraph.Revisions, BuildRevisionTextPreviewLabels(paragraph));
    }

    internal static string BuildRevisionBalloonTitle(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        if (revisions.Count == 0)
        {
            return "Tracked change";
        }

        string[] authors = revisions
            .Select(revision => FirstNonEmpty(revision.Author))
            .Where(author => author is not null)
            .Select(author => author!)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(author => author, StringComparer.Ordinal)
            .ToArray();
        string[] dates = revisions
            .Select(revision => FormatCommentDate(revision.Date))
            .Where(date => date is not null)
            .Select(date => date!)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(date => date, StringComparer.Ordinal)
            .ToArray();
        string owner = authors.Length switch
        {
            0 => "Tracked change",
            1 => authors[0],
            _ => authors.Length.ToString(CultureInfo.InvariantCulture) + " reviewers"
        };
        string? dateLabel = dates.Length switch
        {
            0 => null,
            1 => dates[0],
            _ => dates.Length.ToString(CultureInfo.InvariantCulture) + " dates"
        };
        return dateLabel is null ? owner : owner + " " + dateLabel;
    }

    private static string BuildRevisionBalloonPreview(
        IReadOnlyList<DocxRevisionInfo> revisions,
        IReadOnlyList<string> previewLabels)
    {
        if (revisions.Count == 0)
        {
            return string.Empty;
        }

        string[] labels = BuildRevisionBalloonLabels(revisions)
            .Concat(previewLabels)
            .Where(label => label.Length != 0)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(RevisionBalloonLabelPriority)
            .ThenBy(label => label, StringComparer.Ordinal)
            .ToArray();
        return labels.Length == 0 ? "Revision" : string.Join(", ", labels);
    }

    private static IReadOnlyList<string> BuildRevisionBalloonLabels(IReadOnlyList<DocxRevisionInfo> revisions)
    {
        var labels = new List<string>(revisions.Count);
        labels.AddRange(revisions
            .Where(revision => !IsFormattingRevision(revision))
            .Select(RevisionBalloonLabel));
        labels.AddRange(revisions
            .Where(IsFormattingRevision)
            .GroupBy(ResolveFormattingRevisionFamily, StringComparer.Ordinal)
            .Select(group => BuildFormattingRevisionBalloonLabel(
                group.Key,
                group.SelectMany(revision => revision.PropertyElementNames))));
        return labels;
    }

    private static IReadOnlyList<string> BuildRevisionTextPreviewLabels(DocxParagraph paragraph)
    {
        var labels = new List<string>(3);
        string? deletedText = BuildRevisionTextPreview(paragraph.Runs, "Deletion");
        if (deletedText is not null)
        {
            labels.Add("Deleted: \"" + deletedText + "\"");
        }

        string? movedFromText = BuildRevisionTextPreview(paragraph.Runs, "MoveFrom");
        if (movedFromText is not null)
        {
            labels.Add("Moved from: \"" + movedFromText + "\"");
        }

        string? movedToText = BuildRevisionTextPreview(paragraph.Runs, "MoveTo");
        if (movedToText is not null)
        {
            labels.Add("Moved to: \"" + movedToText + "\"");
        }

        return labels;
    }

    private static string? BuildRevisionTextPreview(IEnumerable<DocxTextRun> runs, string kind)
    {
        string text = string.Concat(runs
            .Where(run => HasRevisionKind(run, kind))
            .Select(run => run.Text));
        string normalized = NormalizeRevisionPreviewText(text);
        return normalized.Length == 0 || normalized.Length > 42 ? null : normalized;
    }

    private static string NormalizeRevisionPreviewText(string value)
    {
        if (value.Any(c => char.IsControl(c) && !char.IsWhiteSpace(c)))
        {
            return string.Empty;
        }

        return Regex.Replace(value, @"\s+", " ").Trim();
    }

    private static string RevisionBalloonLabel(DocxRevisionInfo revision)
    {
        return revision.Kind switch
        {
            "Insertion" => "Inserted text",
            "Deletion" => "Deleted text",
            "MoveFrom" => "Moved from",
            "MoveTo" => "Moved to",
            "RunPropertiesChange" => BuildFormattingRevisionBalloonLabel(revision),
            "ParagraphPropertiesChange" => BuildFormattingRevisionBalloonLabel(revision),
            "TablePropertiesChange" => BuildFormattingRevisionBalloonLabel(revision),
            "TableRowPropertiesChange" => BuildFormattingRevisionBalloonLabel(revision),
            "TableCellPropertiesChange" => BuildFormattingRevisionBalloonLabel(revision),
            "SectionPropertiesChange" => BuildFormattingRevisionBalloonLabel(revision),
            _ => string.IsNullOrWhiteSpace(revision.Kind) ? string.Empty : revision.Kind
        };
    }

    private static int RevisionBalloonLabelPriority(string label)
    {
        if (string.Equals(label, "Inserted text", StringComparison.Ordinal))
        {
            return 10;
        }

        if (string.Equals(label, "Deleted text", StringComparison.Ordinal))
        {
            return 20;
        }

        if (string.Equals(label, "Moved from", StringComparison.Ordinal))
        {
            return 30;
        }

        if (string.Equals(label, "Moved to", StringComparison.Ordinal))
        {
            return 40;
        }

        int? formattingPriority = FormattingRevisionLabelPriority(label);
        if (formattingPriority is not null)
        {
            return formattingPriority.Value;
        }

        if (label.StartsWith("Deleted: ", StringComparison.Ordinal))
        {
            return 60;
        }

        if (label.StartsWith("Moved from: ", StringComparison.Ordinal))
        {
            return 70;
        }

        if (label.StartsWith("Moved to: ", StringComparison.Ordinal))
        {
            return 80;
        }

        return 100;
    }

    private static int? FormattingRevisionLabelPriority(string label)
    {
        if (label.StartsWith("Formatted section", StringComparison.Ordinal))
        {
            return 50;
        }

        if (label.StartsWith("Formatted table", StringComparison.Ordinal))
        {
            return 51;
        }

        if (label.StartsWith("Formatted row", StringComparison.Ordinal))
        {
            return 52;
        }

        if (label.StartsWith("Formatted cell", StringComparison.Ordinal))
        {
            return 53;
        }

        if (label.StartsWith("Formatted paragraph", StringComparison.Ordinal))
        {
            return 54;
        }

        if (label.StartsWith("Formatted run", StringComparison.Ordinal))
        {
            return 55;
        }

        if (string.Equals(label, "Formatting change", StringComparison.Ordinal) ||
            label.StartsWith("Formatted ", StringComparison.Ordinal))
        {
            return 56;
        }

        return null;
    }

    private static bool IsFormattingRevision(DocxRevisionInfo revision)
    {
        return revision.Kind is
            "RunPropertiesChange" or
            "ParagraphPropertiesChange" or
            "TablePropertiesChange" or
            "TableRowPropertiesChange" or
            "TableCellPropertiesChange" or
            "SectionPropertiesChange";
    }

    private static string ResolveFormattingRevisionFamily(DocxRevisionInfo revision)
    {
        return revision.PropertyChangeFamily ?? revision.Kind switch
        {
            "RunPropertiesChange" => "Run",
            "ParagraphPropertiesChange" => "Paragraph",
            "TablePropertiesChange" => "Table",
            "TableRowPropertiesChange" => "Row",
            "TableCellPropertiesChange" => "Cell",
            "SectionPropertiesChange" => "Section",
            _ => "formatting"
        };
    }

    private static string BuildFormattingRevisionBalloonLabel(DocxRevisionInfo revision)
    {
        return BuildFormattingRevisionBalloonLabel(ResolveFormattingRevisionFamily(revision), revision.PropertyElementNames);
    }

    private static string BuildFormattingRevisionBalloonLabel(string family, IEnumerable<string> propertyElementNames)
    {
        string prefix = family switch
        {
            "Run" => "Formatted run",
            "Paragraph" => "Formatted paragraph",
            "Table" => "Formatted table",
            "Row" => "Formatted row",
            "Cell" => "Formatted cell",
            "Section" => "Formatted section",
            _ => "Formatting change"
        };
        string[] names = propertyElementNames
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        string[] properties = names
            .OrderBy(name => FormattingRevisionPropertyPriority(family, name))
            .ThenBy(name => FormatFormattingRevisionPropertyName(family, name), StringComparer.Ordinal)
            .Select(name => FormatFormattingRevisionPropertyName(family, name))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .Take(4)
            .ToArray();
        int hiddenPropertyCount = names
            .Select(name => FormatFormattingRevisionPropertyName(family, name))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .Count() - properties.Length;
        string suffix = hiddenPropertyCount <= 0
            ? string.Empty
            : ", +" + hiddenPropertyCount.ToString(CultureInfo.InvariantCulture) + " more";
        return properties.Length == 0 ? prefix : prefix + ": " + string.Join(", ", properties) + suffix;
    }

    private static int FormattingRevisionPropertyPriority(string family, string value)
    {
        return family switch
        {
            "Run" => value switch
            {
                "rStyle" => 0,
                "rFonts" => 10,
                "sz" => 20,
                "color" => 30,
                "highlight" => 40,
                "b" => 50,
                "i" => 60,
                "u" => 70,
                "vertAlign" => 80,
                "shd" => 90,
                "strike" => 100,
                "dstrike" => 110,
                "caps" => 120,
                "smallCaps" => 130,
                "vanish" => 140,
                "rtl" => 150,
                "lang" => 160,
                "eastAsianLayout" => 170,
                "position" => 180,
                "kern" => 190,
                "outline" => 200,
                "shadow" => 210,
                "emboss" => 220,
                "imprint" => 230,
                "webHidden" => 240,
                "szCs" => 250,
                "bCs" => 260,
                "iCs" => 270,
                "spacing" => 280,
                "w" => 290,
                "bdr" => 300,
                "effect" => 310,
                "em" => 320,
                "fitText" => 330,
                "snapToGrid" => 340,
                "noProof" => 350,
                "specVanish" => 360,
                "cs" => 370,
                "oMath" => 380,
                _ => 500
            },
            "Paragraph" => value switch
            {
                "pStyle" => 0,
                "numPr" => 10,
                "ilvl" => 20,
                "numId" => 30,
                "numberingChange" => 40,
                "jc" => 50,
                "ind" => 60,
                "spacing" => 70,
                "tabs" => 80,
                "keepNext" => 90,
                "keepLines" => 100,
                "widowControl" => 110,
                "pageBreakBefore" => 120,
                "outlineLvl" => 130,
                "pBdr" => 140,
                "shd" => 150,
                "contextualSpacing" => 160,
                "wordWrap" => 170,
                "bidi" => 180,
                "textAlignment" => 190,
                "framePr" => 200,
                "suppressLineNumbers" => 210,
                "adjustRightInd" => 220,
                "autoSpaceDE" => 230,
                "autoSpaceDN" => 240,
                "snapToGrid" => 250,
                "kinsoku" => 260,
                "overflowPunct" => 270,
                "topLinePunct" => 280,
                "suppressAutoHyphens" => 290,
                "mirrorIndents" => 300,
                "suppressOverlap" => 310,
                "cnfStyle" => 320,
                "divId" => 330,
                _ => 500
            },
            "Table" => value switch
            {
                "tblStyle" => 0,
                "tblBorders" => 10,
                "tblW" => 20,
                "tblLayout" => 30,
                "tblInd" => 40,
                "tblCellMar" => 50,
                "tblLook" => 60,
                "jc" => 70,
                "tblCellSpacing" => 80,
                "shd" => 90,
                "bidiVisual" => 100,
                "tblpPr" => 110,
                "tblOverlap" => 120,
                "tblCaption" => 130,
                "tblDescription" => 140,
                "tblStyleRowBandSize" => 150,
                "tblStyleColBandSize" => 160,
                "cnfStyle" => 170,
                _ => 500
            },
            "Row" => value switch
            {
                "tblHeader" => 0,
                "trHeight" => 10,
                "cantSplit" => 20,
                "jc" => 30,
                "tblCellSpacing" => 40,
                "gridBefore" => 50,
                "gridAfter" => 60,
                "wBefore" => 70,
                "wAfter" => 80,
                "hidden" => 90,
                "tblPrEx" => 100,
                "cnfStyle" => 110,
                "divId" => 120,
                _ => 500
            },
            "Cell" => value switch
            {
                "tcW" => 0,
                "gridSpan" => 10,
                "vMerge" => 20,
                "vAlign" => 30,
                "tcMar" => 40,
                "tcBorders" => 50,
                "shd" => 60,
                "textDirection" => 70,
                "tcFitText" => 80,
                "noWrap" => 90,
                "hideMark" => 100,
                "hMerge" => 110,
                "cellIns" => 120,
                "cellDel" => 130,
                "cellMerge" => 140,
                "cnfStyle" => 150,
                _ => 500
            },
            "Section" => value switch
            {
                "type" => 0,
                "pgSz" => 10,
                "pgMar" => 20,
                "cols" => 30,
                "headerReference" => 40,
                "footerReference" => 50,
                "pgNumType" => 60,
                "docGrid" => 70,
                "lnNumType" => 80,
                "footnotePr" => 90,
                "endnotePr" => 100,
                "pgBorders" => 110,
                "titlePg" => 120,
                "vAlign" => 130,
                "paperSrc" => 140,
                "textDirection" => 150,
                "bidi" => 160,
                "rtlGutter" => 170,
                "printerSettings" => 180,
                "formProt" => 190,
                _ => 500
            },
            _ => 500
        };
    }

    private static string FormatFormattingRevisionPropertyName(string family, string value)
    {
        if (family == "Run")
        {
            string? runName = value switch
            {
                "spacing" => "character spacing",
                "w" => "character scale",
                "bdr" => "border",
                "effect" => "text effect",
                "em" => "emphasis mark",
                "fitText" => "fit text",
                "snapToGrid" => "snap to grid",
                "noProof" => "proofing",
                "specVanish" => "special hidden text",
                "cs" => "complex script",
                "oMath" => "math",
                _ => null
            };
            if (runName is not null)
            {
                return runName;
            }
        }

        return value switch
        {
            "rStyle" => "character style",
            "b" => "bold",
            "bCs" => "complex-script bold",
            "i" => "italic",
            "iCs" => "complex-script italic",
            "u" => "underline",
            "sz" => "font size",
            "szCs" => "complex-script font size",
            "rFonts" => "font",
            "color" => "color",
            "highlight" => "highlight",
            "shd" => "shading",
            "strike" => "strike",
            "dstrike" => "double strike",
            "caps" => "all caps",
            "smallCaps" => "small caps",
            "vanish" => "hidden text",
            "vertAlign" => "vertical position",
            "rtl" => "right-to-left",
            "lang" => "language",
            "eastAsianLayout" => "East Asian layout",
            "position" => "position",
            "kern" => "kerning",
            "outline" => "outline",
            "shadow" => "shadow",
            "emboss" => "emboss",
            "imprint" => "engrave",
            "webHidden" => "web hidden",
            "jc" => "alignment",
            "ind" => "indent",
            "spacing" => "spacing",
            "tabs" => "tab stops",
            "numPr" => "numbering",
            "ilvl" => "list level",
            "numId" => "numbering id",
            "numberingChange" => "numbering change",
            "pStyle" => "paragraph style",
            "keepNext" => "keep next",
            "keepLines" => "keep lines",
            "widowControl" => "widow control",
            "pageBreakBefore" => "page break before",
            "outlineLvl" => "outline level",
            "pBdr" => "borders",
            "contextualSpacing" => "contextual spacing",
            "wordWrap" => "word wrap",
            "bidi" => "right-to-left",
            "textAlignment" => "text alignment",
            "framePr" => "frame",
            "suppressLineNumbers" => "line numbering",
            "adjustRightInd" => "adjust right indent",
            "autoSpaceDE" => "East Asian auto spacing",
            "autoSpaceDN" => "number auto spacing",
            "snapToGrid" => "snap to grid",
            "kinsoku" => "kinsoku",
            "overflowPunct" => "overflow punctuation",
            "topLinePunct" => "top-line punctuation",
            "suppressAutoHyphens" => "auto hyphenation",
            "mirrorIndents" => "mirror indents",
            "suppressOverlap" => "overlap suppression",
            "cnfStyle" => "conditional formatting",
            "divId" => "division",
            "tblW" => "table width",
            "tblBorders" => "borders",
            "tblCellMar" => "cell margins",
            "tblCellSpacing" => "cell spacing",
            "tblInd" => "table indent",
            "tblLayout" => "table layout",
            "tblLook" => "table look",
            "bidiVisual" => "right-to-left table",
            "tblpPr" => "floating table position",
            "tblOverlap" => "table overlap",
            "tblCaption" => "table caption",
            "tblDescription" => "table description",
            "tblStyleRowBandSize" => "row band size",
            "tblStyleColBandSize" => "column band size",
            "trHeight" => "row height",
            "cantSplit" => "row split",
            "tblHeader" => "header row",
            "gridBefore" => "grid before",
            "gridAfter" => "grid after",
            "wBefore" => "width before",
            "wAfter" => "width after",
            "hidden" => "hidden row",
            "tblPrEx" => "table property exceptions",
            "tcW" => "cell width",
            "gridSpan" => "grid span",
            "vMerge" => "vertical merge",
            "hMerge" => "horizontal merge",
            "tcBorders" => "cell borders",
            "tcMar" => "cell margins",
            "vAlign" => "vertical alignment",
            "textDirection" => "text direction",
            "tcFitText" => "fit text",
            "noWrap" => "no wrap",
            "hideMark" => "end mark",
            "cellIns" => "inserted cell",
            "cellDel" => "deleted cell",
            "cellMerge" => "cell merge",
            "pgSz" => "page size",
            "pgMar" => "page margins",
            "cols" => "columns",
            "type" => "section type",
            "headerReference" => "header",
            "footerReference" => "footer",
            "pgNumType" => "page numbering",
            "docGrid" => "document grid",
            "lnNumType" => "line numbering",
            "footnotePr" => "footnotes",
            "endnotePr" => "endnotes",
            "pgBorders" => "page borders",
            "titlePg" => "different first page",
            "paperSrc" => "paper source",
            "rtlGutter" => "right-to-left gutter",
            "printerSettings" => "printer settings",
            "formProt" => "form protection",
            _ => HumanizeFormattingRevisionPropertyName(value)
        };
    }

    private static string HumanizeFormattingRevisionPropertyName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        string spaced = Regex.Replace(value.Trim(), "([a-z0-9])([A-Z])", "$1 $2");
        return spaced.Replace(" Pr", " properties", StringComparison.Ordinal).ToLowerInvariant();
    }

    internal static string BuildCommentBalloonTitle(DocxRelatedStory? story, string? fallbackId)
    {
        DocxCommentMetadata? metadata = story?.CommentMetadata;
        string? author = FirstNonEmpty(metadata?.Author);
        string? initials = FirstNonEmpty(metadata?.Initials);
        string? owner = author is not null && initials is not null &&
            !string.Equals(author, initials, StringComparison.Ordinal)
            ? author + " (" + initials + ")"
            : author ?? initials;
        string label = owner is null
            ? "Comment " + (fallbackId ?? "?")
            : fallbackId is null ? owner : owner + " #" + fallbackId;
        string? date = FormatCommentDate(metadata?.Date);
        if (metadata?.ParentCommentId is not null)
        {
            label = "Reply " + label;
        }

        if (date is not null)
        {
            label += " " + date;
        }

        if (metadata?.IsResolved is not null)
        {
            label += metadata.IsResolved == true ? " resolved" : " open";
        }

        return label;
    }

    private static string BuildWordCompatibleCommentBalloonTitle(DocxRelatedStory? story, string? fallbackId)
    {
        DocxCommentMetadata? metadata = story?.CommentMetadata;
        string? initials = FirstNonEmpty(metadata?.Initials, metadata?.Author);
        string? id = FirstNonEmpty(fallbackId);
        if (initials is not null && id is not null)
        {
            return "Commented [" + initials + id + "]: ";
        }

        if (initials is not null)
        {
            return "Commented [" + initials + "]: ";
        }

        return id is null
            ? "Commented: "
            : "Commented [Comment " + id + "]: ";
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        foreach (string? value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return null;
    }

    private static string? FormatCommentDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out DateTimeOffset date)
            ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : value.Trim();
    }

    internal static string BuildCommentBalloonPreview(DocxRelatedStoryLayout? storyLayout)
    {
        return BuildCommentBalloonPreview(storyLayout, []);
    }

    internal static string BuildCommentBalloonPreview(
        DocxRelatedStoryLayout? storyLayout,
        IReadOnlyList<DocxRelatedStoryLayout> replies)
    {
        if (storyLayout is null)
        {
            return string.Empty;
        }

        List<string> parts = BuildCommentStoryPreviewParts(storyLayout);
        if (replies.Count != 0)
        {
            parts.Add(replies.Count == 1 ? "1 reply" : replies.Count.ToString(CultureInfo.InvariantCulture) + " replies");
            foreach (string replyPreview in replies
                .Select(reply => string.Join(" ", BuildCommentStoryPreviewParts(reply)))
                .Where(preview => !string.IsNullOrWhiteSpace(preview)))
            {
                parts.Add("Reply: " + replyPreview);
            }
        }

        return string.Join(" ", parts);
    }

    internal static string BuildWordCompatibleCommentBalloonPreview(
        DocxRelatedStoryLayout? storyLayout,
        IReadOnlyList<DocxRelatedStoryLayout> replies)
    {
        if (storyLayout is null)
        {
            return string.Empty;
        }

        var parts = new List<string>();
        string parentPreview = BuildWordCompatibleCommentStoryPreview(storyLayout, prefix: null);
        if (!string.IsNullOrWhiteSpace(parentPreview))
        {
            parts.Add(parentPreview);
        }

        if (replies.Count != 0)
        {
            parts.Add(replies.Count == 1 ? "1 reply" : replies.Count.ToString(CultureInfo.InvariantCulture) + " replies");
            foreach (string replyPreview in replies
                .Select(reply => BuildWordCompatibleCommentStoryPreview(reply, "Reply"))
                .Where(preview => !string.IsNullOrWhiteSpace(preview)))
            {
                parts.Add(replyPreview);
            }
        }

        return string.Join(" ", parts);
    }

    private static string BuildWordCompatibleCommentStoryPreview(DocxRelatedStoryLayout storyLayout, string? prefix)
    {
        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(prefix))
        {
            parts.Add(prefix);
        }

        DocxCommentMetadata? metadata = storyLayout.Story.CommentMetadata;
        string? date = FormatCommentDate(metadata?.Date);
        if (date is not null)
        {
            parts.Add(date);
        }

        if (metadata?.IsResolved is { } isResolved)
        {
            parts.Add(isResolved ? "resolved" : "open");
        }

        parts.AddRange(BuildCommentStoryPreviewParts(storyLayout));
        return string.Join(" ", parts);
    }

    private static List<string> BuildCommentStoryPreviewParts(DocxRelatedStoryLayout storyLayout)
    {
        var parts = DocxBlockTraversal
            .EnumerateBodyParagraphs(storyLayout.Story)
            .Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)).Trim())
            .Where(text => text.Length != 0)
            .ToList();
        int tableCount = DocxBlockTraversal.EnumerateBodyTables(storyLayout.Story).Count();
        int inlineImageCount = storyLayout.InlineImages.Count + CountTableInlineImages(storyLayout.TableRows);
        int floatingDrawingCount = storyLayout.FloatingDrawings.Count(drawing => drawing.Drawing.Image is not null || !string.IsNullOrWhiteSpace(drawing.Drawing.ImageRelationshipId));

        if (tableCount != 0)
        {
            parts.Add(tableCount == 1 ? "[table]" : "[" + tableCount.ToString(CultureInfo.InvariantCulture) + " tables]");
        }

        int visualImageCount = inlineImageCount + floatingDrawingCount;
        if (visualImageCount != 0)
        {
            parts.Add(visualImageCount == 1 ? "[image]" : "[" + visualImageCount.ToString(CultureInfo.InvariantCulture) + " images]");
        }

        return parts;
    }

    private static int CountTableInlineImages(IReadOnlyList<DocxTableRowLayout> rows)
    {
        int count = 0;
        foreach (DocxTableRowLayout row in rows)
        {
            foreach (DocxTableCellLayout cell in row.Cells)
            {
                count += cell.InlineImages.Count;
                count += CountTableInlineImages(cell.NestedRows);
            }
        }

        return count;
    }

    private static string TrimBalloonText(string text, double marginRight)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        int maxLength = Math.Max(8, (int)Math.Floor(Math.Max(42d, marginRight - 8d) / 2.8d));
        string normalized = text.Trim();
        return normalized.Length <= maxLength
            ? normalized
            : normalized[..Math.Max(0, maxLength - 3)] + "...";
    }

    private static string[] WrapWordCompatibleBalloonBody(
        string text,
        PdfEmbeddedFont embedded,
        double fontSize,
        double firstLineWidth,
        double continuationWidth)
    {
        string normalized = Regex.Replace(text.Trim(), @"\s+", " ");
        if (normalized.Length == 0)
        {
            return [];
        }

        string[] words = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        string firstLine = ConsumeBalloonWords(words, 0, firstLineWidth, embedded, fontSize, out int nextWordIndex);
        if (nextWordIndex >= words.Length)
        {
            return [firstLine];
        }

        string secondLine = string.Join(" ", words.Skip(nextWordIndex));
        secondLine = FitWordCompatibleBalloonLine(secondLine, continuationWidth, embedded, fontSize);
        return [firstLine + " ", secondLine];
    }

    private static string ConsumeBalloonWords(
        string[] words,
        int startIndex,
        double maxWidth,
        PdfEmbeddedFont embedded,
        double fontSize,
        out int nextWordIndex)
    {
        string line = string.Empty;
        int index = startIndex;
        for (; index < words.Length; index++)
        {
            string candidate = line.Length == 0
                ? words[index]
                : line + " " + words[index];
            if (line.Length != 0 &&
                embedded.MeasureTextPoints(candidate, fontSize) > maxWidth)
            {
                break;
            }

            line = candidate;
        }

        if (line.Length == 0 && startIndex < words.Length)
        {
            line = FitWordCompatibleBalloonLine(words[startIndex], maxWidth, embedded, fontSize);
            index = startIndex + 1;
        }

        nextWordIndex = index;
        return line;
    }

    private static string FitWordCompatibleBalloonLine(
        string text,
        double maxWidth,
        PdfEmbeddedFont embedded,
        double fontSize)
    {
        if (text.Length == 0 ||
            maxWidth <= 0d ||
            embedded.MeasureTextPoints(text, fontSize) <= maxWidth)
        {
            return text;
        }

        const string suffix = "...";
        for (int length = Math.Max(0, text.Length - 1); length >= 0; length--)
        {
            string candidate = text[..length].TrimEnd() + suffix;
            if (embedded.MeasureTextPoints(candidate, fontSize) <= maxWidth)
            {
                return candidate;
            }
        }

        return suffix;
    }

    private static void DrawBalloonText(
        PdfGraphicsBuilder graphics,
        DocxRunFontResource resource,
        string text,
        double x,
        double baselineY,
        double fontSize,
        byte red,
        byte green,
        byte blue,
        double positioningCharacterSpacing = 0d)
    {
        if (Math.Abs(positioningCharacterSpacing) > 0.001d)
        {
            string? positioningArray = resource.Embedded.EncodeGlyphPositioningArray(
                text,
                positioningCharacterSpacing,
                fontSize,
                forcePositioningArray: true,
                kerningEnabled: false);
            if (positioningArray is not null)
            {
                graphics.DrawGlyphPositionedText(resource.Name, fontSize, x, baselineY, red, green, blue, positioningArray);
            }

            return;
        }

        string glyphHex = resource.Embedded.EncodeGlyphHex(text);
        if (glyphHex.Length != 0)
        {
            graphics.DrawGlyphText(resource.Name, fontSize, x, baselineY, red, green, blue, glyphHex);
        }
    }

    private static double GetSegmentFontSize(DocxTextSegmentLayout segment, double lineFontSize)
    {
        return segment.FontSize ?? lineFontSize;
    }

    private static double GetSegmentBaselineY(DocxTextSegmentLayout segment, double lineBaselineY)
    {
        return lineBaselineY + segment.BaselineOffsetY;
    }

    private static void RenderTextEmissionSegment(
        DocxTextEmissionSegment segment,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        DocxTextRun style = segment.StyleRun;
        RgbColor color = segment.Color;
        if (!segment.IsTerminalLineSpace)
        {
            RenderRunBackground(style, segment.Resource.Embedded.Font, segment.X, segment.Width, segment.FontSize, segment.BaselineY, graphics);
        }

        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            style,
            segment.FontSize,
            segment.PdfCharacterSpacing,
            segment.PdfCharacterSpacingSource,
            segment.CompensatePdfCharacterSpacing,
            segment.IsTerminalLineSpace);
        DrawRunGlyphText(graphics, segment.Resource, segment.Text, segment.X, segment.BaselineY, color, plan, segment.SyntheticItalic);
        if (!segment.IsTerminalLineSpace && segment.SyntheticBold)
        {
            DrawRunGlyphText(graphics, segment.Resource, segment.Text, segment.X + 0.35d, segment.BaselineY, color, plan, segment.SyntheticItalic);
        }

        if (!segment.IsTerminalLineSpace)
        {
            bool useWordCompatibleRevisionDecorationProfile = ShouldUseWordCompatibleRevisionDecorationProfile(segment, markupContext);
            RgbColor decorationColor = useWordCompatibleRevisionDecorationProfile
                ? new RgbColor(WordCompatibleAllMarkupReviewStrokeRgb.Red, WordCompatibleAllMarkupReviewStrokeRgb.Green, WordCompatibleAllMarkupReviewStrokeRgb.Blue)
                : color;
            double decorationWidth = ResolveRevisionDecorationWidth(segment, useWordCompatibleRevisionDecorationProfile);
            RenderTextDecorations(
                style,
                segment.Resource.Embedded,
                segment.Text,
                segment.X,
                decorationWidth,
                segment.FontSize,
                segment.BaselineY,
                decorationColor,
                plan,
                useWordCompatibleRevisionDecorationProfile,
                graphics);
        }
    }

    private static bool ShouldUseWordCompatibleRevisionDecorationProfile(
        DocxTextEmissionSegment segment,
        DocxMarkupContext markupContext)
    {
        return UsesWordCompatibleAllMarkupTextProfile(markupContext) &&
            (segment.StyleRun.Revision is not null || segment.StyleRun.Revisions.Count != 0);
    }

    private static double ResolveRevisionDecorationWidth(
        DocxTextEmissionSegment segment,
        bool useWordCompatibleRevisionDecorationProfile)
    {
        if (!useWordCompatibleRevisionDecorationProfile ||
            !HasInsertionLikeRevisionKind(segment.StyleRun) ||
            segment.Text.EnumerateRunes().Count() > 8)
        {
            return segment.Width;
        }

        return Math.Max(0d, segment.Width - WordCompatibleAllMarkupShortInsertionDecorationWidthInsetPoints);
    }

    private static DocxRunFontResource? ResolveFontResource(DocxTextRun run, DocxFontResources fontResources)
    {
        return fontResources.RunResources.TryGetValue(run, out DocxRunFontResource? resource)
            ? resource
            : fontResources.Fallback;
    }

    private static void RenderInlineImage(
        DocxInlineImageLayout image,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        PdfImageXObject? xObject = CreateImage(image.Image, diagnosticSink, image.PageIndex);
        if (xObject is null)
        {
            return;
        }

        string imageName = "Im" + imageIndex++;
        graphics.DrawImage(imageName, image.X, image.Y, image.Width, image.Height);
        pageImages.Add(new PdfImageResource(imageName, xObject));
    }

    private static void RenderPlacedRelatedStory(
        DocxPlacedRelatedStoryLayout story,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        ref int imageIndex)
    {
        if (story.SeparatorY is { } separatorY)
        {
            graphics.SetFillRgb(0, 0, 0);
            graphics.FillRectangle(story.X, separatorY, Math.Min(story.SeparatorWidth, story.Width), story.SeparatorThickness);
        }

        graphics.SaveState();
        graphics.ClipRectangle(story.X, story.TopY - story.Height, story.Width, story.Height);
        IReadOnlyList<DocxLayoutItem> items = story.TextLines
            .Cast<DocxLayoutItem>()
            .Concat(story.InlineImages)
            .Concat(story.TableRows)
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, ref imageIndex);
        }

        graphics.RestoreState();
    }

    private static void RenderFloatingDrawings(
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        int pageIndex,
        bool behindDocument,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        foreach (DocxFloatingDrawingLayout drawing in floatingDrawings
            .Where(drawing => drawing.AnchorPageIndex == pageIndex && IsBehindDocument(drawing.Drawing) == behindDocument)
            .OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue)))
        {
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, ref imageIndex);
        }
    }

    private static void RenderPlacedRelatedStoryDrawings(
        DocxPlacedRelatedStoryLayout story,
        bool behindDocument,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        DocxFloatingDrawingLayout[] drawings = story.FloatingDrawings
            .Where(drawing => IsBehindDocument(drawing.Drawing) == behindDocument)
            .OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue))
            .ToArray();
        if (drawings.Length == 0)
        {
            return;
        }

        graphics.SaveState();
        graphics.ClipRectangle(story.X, story.TopY - story.Height, story.Width, story.Height);
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, ref imageIndex);
        }

        graphics.RestoreState();
    }

    private static void RenderFloatingDrawing(
        DocxFloatingDrawingLayout drawing,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        if (drawing.PlacedX is not { } placedX ||
            drawing.PlacedTop is not { } placedTop ||
            drawing.ExtentWidthPoints is not { } width ||
            drawing.ExtentHeightPoints is not { } height)
        {
            return;
        }

        if (drawing.Drawing.Image is { } image)
        {
            PdfImageXObject? xObject = CreateImage(image, diagnosticSink, drawing.AnchorPageIndex ?? 0);
            if (xObject is not null)
            {
                string imageName = "Im" + imageIndex++;
                graphics.DrawImage(imageName, placedX, placedTop - height, width, height);
                pageImages.Add(new PdfImageResource(imageName, xObject));
            }
        }

        if (drawing.TextBoxLayout is { } textBoxLayout)
        {
            RenderFloatingTextBox(drawing, textBoxLayout, placedX, placedTop, width, height, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, ref imageIndex);
        }
    }

    private static void RenderFloatingTextBox(
        DocxFloatingDrawingLayout drawing,
        DocxRelatedStoryLayout textBoxLayout,
        double placedX,
        double placedTop,
        double width,
        double height,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        ref int imageIndex)
    {
        if (textBoxLayout.TextLines.Count == 0 &&
            textBoxLayout.InlineImages.Count == 0 &&
            textBoxLayout.TableRows.Count == 0)
        {
            return;
        }

        graphics.SaveState();
        graphics.ClipRectangle(placedX, placedTop - height, width, height);
        IReadOnlyList<DocxLayoutItem> items = textBoxLayout.TextLines
            .Select(line => TranslateTextLine(line, placedX, placedTop))
            .Cast<DocxLayoutItem>()
            .Concat(textBoxLayout.InlineImages.Select(image => image with
            {
                X = placedX + image.X,
                Y = placedTop + image.Y,
                PageIndex = drawing.AnchorPageIndex ?? image.PageIndex
            }))
            .Concat(textBoxLayout.TableRows.Select(row => TranslateTableRow(row, placedX, placedTop)))
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, ref imageIndex);
        }

        graphics.RestoreState();
    }

    private static double ResolveLayoutItemTop(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout textLine => textLine.BaselineY,
            DocxInlineImageLayout image => image.Y + image.Height,
            DocxTableRowLayout row => row.Y + row.Height,
            _ => 0d
        };
    }

    private static DocxTextLineLayout TranslateTextLine(DocxTextLineLayout line, double deltaX, double deltaY)
    {
        return line with
        {
            X = line.X + deltaX,
            BaselineY = line.BaselineY + deltaY,
            Segments = line.Segments
                .Select(segment => segment with { X = segment.X + deltaX })
                .ToArray()
        };
    }

    private static bool IsBehindDocument(DocxFloatingDrawing drawing)
    {
        return IsOnOffTrue(drawing.BehindDocumentValue);
    }

    private static long ReadZOrder(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long zOrder)
            ? zOrder
            : 0L;
    }

    private static bool IsOnOffTrue(string? value)
    {
        return value is not null &&
            (value.Length == 0 ||
            value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("on", StringComparison.OrdinalIgnoreCase));
    }

    private static void RenderTableRow(
        DocxTableRowLayout row,
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
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow))
            {
                continue;
            }

            DocxTableCell cell = cellLayout.VisualCell;
            RenderShadingFill(cell.FillHex, cell.ShadingValue, cell.ShadingColor, graphics, cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
        }

        RenderTableRowBorders(row, previousRow, nextRow, graphics);
        RenderTableBorderJunctions(row, previousRow, nextRow, graphics);
        RenderTableRowMarkupIndicators(row, graphics, markupContext);

        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (!ShouldRenderTableCellContentFragment(cellLayout, previousRow))
            {
                continue;
            }

            if (cellLayout.TextLines.Count != 0 || cellLayout.InlineImages.Count != 0 || cellLayout.NestedRows.Count != 0)
            {
                graphics.SaveState();
                graphics.ClipRectangle(cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
                foreach (DocxTextLineLayout line in cellLayout.TextLines)
                {
                    RenderTextLine(line, graphics, fontResources, markupContext, pageNumber, pageCount);
                }

                foreach (DocxInlineImageLayout image in cellLayout.InlineImages)
                {
                    RenderInlineImage(image, graphics, pageImages, diagnosticSink, ref imageIndex);
                }

                for (int nestedRowIndex = 0; nestedRowIndex < cellLayout.NestedRows.Count; nestedRowIndex++)
                {
                    DocxTableRowLayout nestedRow = cellLayout.NestedRows[nestedRowIndex];
                    DocxTableRowLayout? previousNestedRow = nestedRowIndex > 0 ? cellLayout.NestedRows[nestedRowIndex - 1] : null;
                    DocxTableRowLayout? nextNestedRow = nestedRowIndex + 1 < cellLayout.NestedRows.Count ? cellLayout.NestedRows[nestedRowIndex + 1] : null;
                    RenderTableRow(
                        nestedRow,
                        IsAdjacentTableRow(previousNestedRow, nestedRow) ? previousNestedRow : null,
                        IsAdjacentTableRow(nestedRow, nextNestedRow) ? nextNestedRow : null,
                        graphics,
                        pageImages,
                        fontResources,
                        markupContext,
                        diagnosticSink,
                        pageNumber,
                        pageCount,
                        ref imageIndex);
                }

                graphics.RestoreState();
            }
        }
    }

    private static void RenderTableRowMarkupIndicators(
        DocxTableRowLayout row,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        if (!markupContext.DrawsChangeBars ||
            row.RevisionCount == 0 ||
            UsesWordCompatibleAllMarkupTextProfile(markupContext))
        {
            return;
        }

        double height = Math.Max(6d, row.Height);
        DocxMarkupBalloonRgb color = ResolveRevisionAuthorColor(row.Revisions ?? []);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        graphics.FillRectangle(Math.Max(0d, row.Table.TableX - 7d), row.Y, 1.5d, height);
    }

    private static void RenderTableBorderJunctions(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics)
    {
        DocxTableBorderBoundary[] rowBoundaries = ResolveVisibleVerticalBoundaries(row, previousRow);
        if (rowBoundaries.Length == 0)
        {
            return;
        }

        var emittedJunctions = new HashSet<(double X, double Y)>();
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow))
            {
                continue;
            }

            if (previousRow is null && IsFirstTableRowFragment(row))
            {
                RenderHorizontalBorderJunctions(cellLayout.X, cellLayout.X + cellLayout.Width, cellLayout.Y + cellLayout.Height, cellLayout.VisualCell, "top", rowBoundaries, graphics, emittedJunctions);
            }

            if (nextRow is null && IsLastTableRowFragment(row))
            {
                RenderHorizontalBorderJunctions(cellLayout.X, cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.VisualCell, "bottom", rowBoundaries, graphics, emittedJunctions);
            }
        }

        if (previousRow is null && IsFirstTableRowFragment(row))
        {
            RenderOuterHorizontalFragmentCornerJunctions(row, previousRow, rowBoundaries, "top", graphics);
        }

        if (nextRow is null && IsLastTableRowFragment(row))
        {
            RenderOuterHorizontalFragmentCornerJunctions(row, previousRow, rowBoundaries, "bottom", graphics);
        }

        if (nextRow is not null && nextRow.RowIndex != row.RowIndex)
        {
            RenderSharedHorizontalBorderJunctions(row, nextRow, rowBoundaries, graphics, emittedJunctions);
        }
    }

    private static void RenderOuterHorizontalFragmentCornerJunctions(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        IReadOnlyList<DocxTableBorderBoundary> boundaries,
        string edge,
        PdfGraphicsBuilder graphics)
    {
        DocxTableCellLayout? firstCell = row.Cells.FirstOrDefault(cell => ShouldRenderTableCellVisualFragment(cell, previousRow));
        DocxTableCellLayout? lastCell = row.Cells.LastOrDefault(cell => ShouldRenderTableCellVisualFragment(cell, previousRow));
        if (firstCell is null || lastCell is null)
        {
            return;
        }

        DocxTableCellBorder? firstHorizontal = DocxTableBorderGeometry.Find(firstCell.VisualCell.Borders, edge);
        DocxTableCellBorder? lastHorizontal = DocxTableBorderGeometry.Find(lastCell.VisualCell.Borders, edge);
        DocxTableBorderBoundary? firstBoundary = boundaries.OrderBy(boundary => boundary.X).FirstOrDefault();
        DocxTableBorderBoundary? lastBoundary = boundaries.OrderByDescending(boundary => boundary.X).FirstOrDefault();
        double y = string.Equals(edge, "top", StringComparison.Ordinal)
            ? firstCell.Y + firstCell.Height
            : firstCell.Y;

        if (firstBoundary is not null && firstHorizontal is not null && !DocxTableBorderGeometry.IsSuppressed(firstHorizontal))
        {
            RenderBorderJunctions([firstBoundary], y, firstHorizontal, graphics, []);
        }

        if (lastBoundary is not null && lastHorizontal is not null && !DocxTableBorderGeometry.IsSuppressed(lastHorizontal))
        {
            RenderBorderJunctions([lastBoundary], y, lastHorizontal, graphics, []);
        }
    }

    private static void RenderSharedHorizontalBorderJunctions(
        DocxTableRowLayout row,
        DocxTableRowLayout nextRow,
        IReadOnlyList<DocxTableBorderBoundary> rowBoundaries,
        PdfGraphicsBuilder graphics,
        HashSet<(double X, double Y)> emittedJunctions)
    {
        DocxTableBorderBoundary[] nextRowBoundaries = ResolveVisibleVerticalBoundaries(nextRow, row);
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow: null))
            {
                continue;
            }

            DocxTableCellLayout[] overlappingNextCells = nextRow.Cells
                .Where(nextCell => ShouldRenderTableCellVisualFragment(nextCell, row) && HorizontalOverlap(cellLayout, nextCell) > 0d)
                .ToArray();
            if (overlappingNextCells.Length == 0)
            {
                RenderHorizontalBorderJunctions(cellLayout.X, cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.VisualCell, "bottom", rowBoundaries, graphics, emittedJunctions);
                continue;
            }

            foreach (DocxTableCellLayout nextRowCell in overlappingNextCells)
            {
                DocxTableCellBorder? horizontal = ResolveSharedHorizontalBorder(cellLayout, nextRowCell);
                if (horizontal is null)
                {
                    continue;
                }

                double x = Math.Max(cellLayout.X, nextRowCell.X);
                double right = Math.Min(cellLayout.X + cellLayout.Width, nextRowCell.X + nextRowCell.Width);
                if (right <= x)
                {
                    continue;
                }

                DocxTableBorderBoundary[] boundaries = rowBoundaries
                    .Concat(nextRowBoundaries)
                    .Where(boundary => boundary.X >= x - 0.001d && boundary.X <= right + 0.001d)
                    .GroupBy(boundary => Math.Round(boundary.X, 3))
                    .Select(group => group.OrderByDescending(boundary => boundary.Width).First())
                    .ToArray();
                RenderBorderJunctions(boundaries, cellLayout.Y - DocxTableBorderGeometry.ResolveVisibleWidth(horizontal) / 2d, horizontal, graphics, emittedJunctions);
            }
        }
    }

    private static void RenderHorizontalBorderJunctions(
        double x,
        double right,
        double y,
        DocxTableCell cell,
        string edge,
        IReadOnlyList<DocxTableBorderBoundary> boundaries,
        PdfGraphicsBuilder graphics,
        HashSet<(double X, double Y)> emittedJunctions)
    {
        DocxTableCellBorder? horizontal = DocxTableBorderGeometry.Find(cell.Borders, edge);
        if (horizontal is null || DocxTableBorderGeometry.IsSuppressed(horizontal))
        {
            return;
        }

        DocxTableBorderBoundary[] crossingBoundaries = boundaries
            .Where(boundary => boundary.X >= x - 0.001d && boundary.X <= right + 0.001d)
            .ToArray();
        RenderBorderJunctions(crossingBoundaries, y, horizontal, graphics, emittedJunctions);
    }

    private static void RenderBorderJunctions(
        IReadOnlyList<DocxTableBorderBoundary> boundaries,
        double y,
        DocxTableCellBorder horizontal,
        PdfGraphicsBuilder graphics,
        HashSet<(double X, double Y)> emittedJunctions)
    {
        double horizontalWidth = DocxTableBorderGeometry.ResolveVisibleWidth(horizontal);
        if (horizontalWidth <= 0d)
        {
            return;
        }

        foreach (DocxTableBorderBoundary boundary in boundaries)
        {
            if (!emittedJunctions.Add((Math.Round(boundary.X, 3), Math.Round(y, 3))))
            {
                continue;
            }

            DocxTableCellBorder? border = DocxTableBorderGeometry.SelectStronger(horizontal, boundary.Border);
            if (border is null)
            {
                continue;
            }

            RgbColor color = ReadColor(border.Color);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            RenderTableBorderStrip(graphics, border, boundary.X, y, boundary.Width, horizontalWidth, DocxTableBorderOrientation.Horizontal);
        }
    }

    private static DocxTableBorderBoundary[] ResolveVisibleVerticalBoundaries(DocxTableRowLayout row, DocxTableRowLayout? previousRow)
    {
        var boundaries = new List<DocxTableBorderBoundary>();
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCellLayout cellLayout = row.Cells[cellIndex];
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow))
            {
                continue;
            }

            DocxTableCell visualCell = cellLayout.VisualCell;
            DocxTableCellBorder? left = DocxTableBorderGeometry.Find(visualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "start");
            if (cellIndex == 0)
            {
                AddVerticalBoundary(boundaries, cellLayout.X, left);
            }

            DocxTableCellBorder? right = DocxTableBorderGeometry.Find(visualCell.Borders, "right") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "end");
            if (cellIndex == row.Cells.Count - 1)
            {
                AddVerticalBoundary(boundaries, cellLayout.X + cellLayout.Width, right);
                continue;
            }

            DocxTableCellLayout nextCell = row.Cells[cellIndex + 1];
            DocxTableCell nextVisualCell = nextCell.VisualCell;
            DocxTableCellBorder? nextLeft = DocxTableBorderGeometry.Find(nextVisualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(nextVisualCell.Borders, "start");
            if (!DocxTableBorderGeometry.IsSuppressed(right) && !DocxTableBorderGeometry.IsSuppressed(nextLeft))
            {
                AddVerticalBoundary(boundaries, cellLayout.X + cellLayout.Width, DocxTableBorderGeometry.SelectStronger(right, nextLeft));
            }
        }

        return boundaries.ToArray();
    }

    private static void AddVerticalBoundary(List<DocxTableBorderBoundary> boundaries, double x, DocxTableCellBorder? border)
    {
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        if (width <= 0d || border is null)
        {
            return;
        }

        boundaries.Add(new DocxTableBorderBoundary(x, width, border));
    }

    private static IReadOnlyList<DocxTextEmissionSegment> CreateTextEmissionSegments(
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount,
        double fontScale = 1d,
        double baselineOffsetY = 0d,
        double xOffset = 0d,
        bool suppressCommentReferenceSpacer = false,
        bool useWordCompatibleTextProfile = false)
    {
        IReadOnlyList<DocxTextSegmentLayout> segments = line.Segments.Count == 0
            ? [new DocxTextSegmentLayout(line.Text, line.StyleRun, line.X, line.Width)]
            : line.Segments;
        var emissionSegments = new List<DocxTextEmissionSegment>(segments.Count + 1);
        double substitutedFieldXAdjustment = 0d;
        foreach (DocxTextSegmentLayout segment in segments)
        {
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            if (resource is null)
            {
                continue;
            }

            double fontSize = ResolveTextEmissionSegmentFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
            int partSourceTextOffset = segment.SourceTextOffsetInRun;
            foreach (DocxTextEmissionPart part in DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, fontSize, fontResources.TextMeasurer))
            {
                int currentPartSourceTextOffset = partSourceTextOffset;
                partSourceTextOffset += part.Text.Length;
                if (ShouldSuppressCommentReferenceSpacerPart(line, part, suppressCommentReferenceSpacer))
                {
                    continue;
                }

                DocxTextRun emissionStyleRun = ResolveWordCompatibleAllMarkupEmissionStyleRun(
                    segment.StyleRun,
                    segment.Role,
                    part.Text,
                    fontSize,
                    useWordCompatibleTextProfile);
                DocxEffectiveRunProperties emissionEffective = emissionStyleRun.EffectiveProperties;
                RgbColor textColor = ResolveTextEmissionColor(
                    emissionStyleRun,
                    useWordCompatibleTextProfile,
                    ReadColor(emissionEffective.ColorHex));
                string emittedText = ResolveStaticFieldPlaceholders(part.Text, pageNumber, pageCount);
                double emittedWidth = ResolveSubstitutedFieldEmissionWidth(part.Text, emittedText, emissionStyleRun, fontSize, fontResources.TextMeasurer, part.Width);
                emissionSegments.Add(new DocxTextEmissionSegment(
                    emittedText,
                    emissionStyleRun,
                    resource,
                    textColor,
                    part.X + xOffset + substitutedFieldXAdjustment + ResolveWordCompatibleAllMarkupBodyXOffset(segment, part, emissionStyleRun, fontSize, line.X, useWordCompatibleTextProfile),
                    baselineY,
                    emittedWidth,
                    fontSize,
                    segment.PdfCharacterSpacing,
                    segment.PdfCharacterSpacingSource,
                    segment.CompensatePdfCharacterSpacing,
                    ShouldApplySyntheticBold(emissionStyleRun, resource),
                    emissionEffective.Italic && !resource.Resolution.Italic,
                    IsTerminalLineSpace: false,
                    segment.SourceTextRunIndex,
                    currentPartSourceTextOffset,
                    segment.Role));
                substitutedFieldXAdjustment += emittedWidth - part.Width;
            }
        }

        if (line.EndsWithIntraTokenBreak)
        {
            return emissionSegments;
        }

        if (line.EmitsTerminalParagraphMark)
        {
            AddTerminalLineSpace(emissionSegments, segments, line, fontResources, fontScale, baselineOffsetY, xOffset, useWordCompatibleTextProfile);
            return emissionSegments;
        }

        for (int i = segments.Count - 1; i >= 0; i--)
        {
            DocxTextSegmentLayout segment = segments[i];
            if (string.IsNullOrEmpty(segment.Text) || string.IsNullOrWhiteSpace(segment.Text))
            {
                continue;
            }

            if (char.IsWhiteSpace(segment.Text[^1]))
            {
                return emissionSegments;
            }

            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            if (resource is null)
            {
                return emissionSegments;
            }

            double fontSize = ResolveTerminalLineSpaceFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
            double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
            DocxEffectiveRunProperties effective = segment.StyleRun.EffectiveProperties;
            RgbColor color = ResolveTextEmissionColor(
                segment.StyleRun,
                useWordCompatibleTextProfile,
                ReadColor(effective.ColorHex));
            emissionSegments.Add(new DocxTextEmissionSegment(
                " ",
                segment.StyleRun,
                resource,
                color,
                ResolveTerminalLineSpaceX(segments, line, fontResources, fontScale, xOffset, useWordCompatibleTextProfile),
                baselineY,
                0d,
                fontSize,
                PdfCharacterSpacing: 0d,
                PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.TerminalLineSpace,
                CompensatePdfCharacterSpacing: true,
                SyntheticBold: false,
                SyntheticItalic: effective.Italic && !resource.Resolution.Italic,
                IsTerminalLineSpace: true,
                segment.SourceTextRunIndex,
                segment.SourceTextOffsetInRun + segment.Text.Length,
                segment.Role));
            break;
        }

        return emissionSegments;
    }

    private static RgbColor ResolveTextEmissionColor(
        DocxTextRun styleRun,
        bool useWordCompatibleTextProfile,
        RgbColor fallbackColor)
    {
        if (!useWordCompatibleTextProfile ||
            (styleRun.Revision is null && styleRun.Revisions.Count == 0))
        {
            return fallbackColor;
        }

        return new RgbColor(
            WordCompatibleAllMarkupReviewStrokeRgb.Red,
            WordCompatibleAllMarkupReviewStrokeRgb.Green,
            WordCompatibleAllMarkupReviewStrokeRgb.Blue);
    }

    private static DocxTextRun ResolveWordCompatibleAllMarkupEmissionStyleRun(
        DocxTextRun styleRun,
        DocxTextSegmentRole role,
        string text,
        double fontSize,
        bool useWordCompatibleTextProfile)
    {
        if (!ShouldApplyWordCompatibleAllMarkupBodyPositioningSpacing(role, text, useWordCompatibleTextProfile))
        {
            return styleRun;
        }

        return styleRun with
        {
            CharacterSpacingPoints = styleRun.CharacterSpacingPoints +
                ResolveWordCompatibleAllMarkupBodyPositioningCharacterSpacing(styleRun, text, fontSize)
        };
    }

    private static double ResolveWordCompatibleAllMarkupBodyPositioningCharacterSpacing(
        DocxTextRun styleRun,
        string text,
        double fontSize)
    {
        int runeCount = text.EnumerateRunes().Count();
        if (runeCount <= 8)
        {
            return WordCompatibleAllMarkupShortWordPositioningCharacterSpacingPoints;
        }

        if (HasInsertionLikeRevisionKind(styleRun))
        {
            return WordCompatibleAllMarkupInsertionPositioningCharacterSpacingPoints;
        }

        if (HasDeletionLikeRevisionKind(styleRun))
        {
            return WordCompatibleAllMarkupDeletionPositioningCharacterSpacingPoints;
        }

        if (fontSize >= WordCompatibleAllMarkupMaxBodyTextFontSizePoints - 0.001d)
        {
            return WordCompatibleAllMarkupHeadingPositioningCharacterSpacingPoints;
        }

        if (text.Any(char.IsPunctuation))
        {
            return WordCompatibleAllMarkupPunctuationPositioningCharacterSpacingPoints;
        }

        return WordCompatibleAllMarkupBodyPositioningCharacterSpacingPoints;
    }

    private static bool HasRevisionKind(DocxTextRun styleRun, string kind)
    {
        return IsRevisionKind(styleRun.Revision, kind) ||
            styleRun.Revisions.Any(revision => IsRevisionKind(revision, kind));
    }

    private static bool HasInsertionLikeRevisionKind(DocxTextRun styleRun)
    {
        return HasRevisionKind(styleRun, "Insertion") ||
            HasRevisionKind(styleRun, "MoveTo");
    }

    private static bool HasDeletionLikeRevisionKind(DocxTextRun styleRun)
    {
        return HasRevisionKind(styleRun, "Deletion") ||
            HasRevisionKind(styleRun, "MoveFrom");
    }

    private static bool IsRevisionKind(DocxRevisionInfo? revision, string kind)
    {
        return revision is not null &&
            string.Equals(revision.Kind, kind, StringComparison.OrdinalIgnoreCase);
    }

    private static bool ShouldApplyWordCompatibleAllMarkupBodyPositioningSpacing(
        DocxTextSegmentRole role,
        string text,
        bool useWordCompatibleTextProfile)
    {
        return useWordCompatibleTextProfile &&
            role == DocxTextSegmentRole.Text &&
            text.EnumerateRunes().Take(2).Count() > 1 &&
            text.Any(character => !char.IsWhiteSpace(character));
    }

    private static double ResolveWordCompatibleAllMarkupBodyXOffset(
        DocxTextSegmentLayout sourceSegment,
        DocxTextEmissionPart part,
        DocxTextRun emissionStyleRun,
        double fontSize,
        double lineX,
        bool useWordCompatibleTextProfile)
    {
        if (!ShouldApplyWordCompatibleAllMarkupBodyPositioningSpacing(sourceSegment.Role, part.Text, useWordCompatibleTextProfile) ||
            fontSize >= WordCompatibleAllMarkupMaxBodyTextFontSizePoints - 0.001d)
        {
            return 0d;
        }

        double lineRelativeX = Math.Max(0d, part.X - lineX);
        if (lineRelativeX < 0.001d)
        {
            return 0d;
        }

        if (HasDeletionLikeRevisionKind(emissionStyleRun))
        {
            return WordCompatibleAllMarkupDeletionXOffsetPoints;
        }

        if (HasInsertionLikeRevisionKind(emissionStyleRun))
        {
            return WordCompatibleAllMarkupInsertionXOffsetPoints;
        }

        return WordCompatibleAllMarkupBodyXOffsetAsymptotePoints *
            (1d - Math.Exp(-lineRelativeX / WordCompatibleAllMarkupBodyXOffsetDecayPoints));
    }

    private static double ResolveTextEmissionSegmentFontSize(
        DocxTextSegmentLayout segment,
        DocxTextLineLayout line,
        double fontScale,
        bool useWordCompatibleTextProfile)
    {
        double fontSize = GetSegmentFontSize(segment, line.FontSize) * fontScale;
        return ShouldCapWordCompatibleAllMarkupTextFontSize(fontSize, useWordCompatibleTextProfile)
            ? WordCompatibleAllMarkupMaxBodyTextFontSizePoints
            : fontSize;
    }

    private static double ResolveTerminalLineSpaceFontSize(
        DocxTextSegmentLayout segment,
        DocxTextLineLayout line,
        double fontScale,
        bool useWordCompatibleTextProfile)
    {
        double fontSize = GetSegmentFontSize(segment, line.FontSize) * fontScale;
        return ShouldCapWordCompatibleAllMarkupTextFontSize(fontSize, useWordCompatibleTextProfile)
            ? WordCompatibleAllMarkupTerminalLineSpaceFontSizePoints
            : fontSize;
    }

    private static double ResolveTerminalLineSpaceX(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale,
        double xOffset,
        bool useWordCompatibleTextProfile)
    {
        if (segments.Count == 0)
        {
            return line.X + xOffset;
        }

        DocxTextSegmentLayout terminalSegment = segments[^1];
        if (!ShouldUseWordCompatibleAllMarkupEmittedTerminalAdvance(segments, line, fontScale, useWordCompatibleTextProfile))
        {
            return terminalSegment.X + terminalSegment.Width + xOffset;
        }

        double? emittedEndX = TryResolveWordCompatibleAllMarkupTextEmissionEndX(
            segments,
            line,
            fontResources,
            fontScale,
            xOffset,
            useWordCompatibleTextProfile);
        return emittedEndX ?? terminalSegment.X + terminalSegment.Width + xOffset;
    }

    private static double? TryResolveWordCompatibleAllMarkupTextEmissionEndX(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale,
        double xOffset,
        bool useWordCompatibleTextProfile)
    {
        double emittedEndX = double.NegativeInfinity;
        foreach (DocxTextSegmentLayout segment in segments)
        {
            DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
            if (resource is null)
            {
                continue;
            }

            double fontSize = ResolveTextEmissionSegmentFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
            foreach (DocxTextEmissionPart part in DocxTextEmissionPlanner.SplitOfficeTextOperationParts(segment, fontSize, fontResources.TextMeasurer))
            {
                DocxTextRun emissionStyleRun = ResolveWordCompatibleAllMarkupEmissionStyleRun(
                    segment.StyleRun,
                    segment.Role,
                    part.Text,
                    fontSize,
                    useWordCompatibleTextProfile);
                DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
                    emissionStyleRun,
                    fontSize,
                    segment.PdfCharacterSpacing,
                    segment.PdfCharacterSpacingSource,
                    segment.CompensatePdfCharacterSpacing,
                    isTerminalLineSpace: false);
                double partX = part.X + xOffset + ResolveWordCompatibleAllMarkupBodyXOffset(
                    segment,
                    part,
                    emissionStyleRun,
                    fontSize,
                    line.X,
                    useWordCompatibleTextProfile);
                double partAdvance = DocxTextEmissionPlanner.MeasureAdvanceProfile(part.Text, resource.Embedded, part.Width, plan)
                    .PlannedEmittedAdvance;
                emittedEndX = Math.Max(emittedEndX, partX + partAdvance);
            }
        }

        return double.IsNegativeInfinity(emittedEndX) ? null : emittedEndX;
    }

    private static bool ShouldUseWordCompatibleAllMarkupEmittedTerminalAdvance(
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        double fontScale,
        bool useWordCompatibleTextProfile)
    {
        return useWordCompatibleTextProfile &&
            segments.Any(segment => ShouldCapWordCompatibleAllMarkupTextFontSize(GetSegmentFontSize(segment, line.FontSize) * fontScale, useWordCompatibleTextProfile));
    }

    private static bool ShouldCapWordCompatibleAllMarkupTextFontSize(double fontSize, bool useWordCompatibleTextProfile)
    {
        return useWordCompatibleTextProfile && fontSize > WordCompatibleAllMarkupMaxBodyTextFontSizePoints;
    }

    private static bool ShouldSuppressCommentReferenceSpacerPart(
        DocxTextLineLayout line,
        DocxTextEmissionPart part,
        bool suppressCommentReferenceSpacer)
    {
        return suppressCommentReferenceSpacer &&
            part.Width > 0d &&
            !string.IsNullOrEmpty(part.Text) &&
            string.IsNullOrWhiteSpace(part.Text) &&
            line.SourceParagraph?.InlineReferences.Any(reference => reference.Kind == "Comment") == true;
    }

    private static void AddTerminalLineSpace(
        List<DocxTextEmissionSegment> emissionSegments,
        IReadOnlyList<DocxTextSegmentLayout> segments,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        double fontScale = 1d,
        double baselineOffsetY = 0d,
        double xOffset = 0d,
        bool useWordCompatibleTextProfile = false)
    {
        if (segments.Count == 0)
        {
            return;
        }

        DocxTextSegmentLayout segment = segments[^1];
        DocxRunFontResource? resource = ResolveFontResource(segment.StyleRun, fontResources);
        if (resource is null)
        {
            return;
        }

        double fontSize = ResolveTerminalLineSpaceFontSize(segment, line, fontScale, useWordCompatibleTextProfile);
        double baselineY = GetSegmentBaselineY(segment, line.BaselineY) - baselineOffsetY;
        DocxEffectiveRunProperties effective = segment.StyleRun.EffectiveProperties;
        emissionSegments.Add(new DocxTextEmissionSegment(
            " ",
            segment.StyleRun,
            resource,
            ReadColor(effective.ColorHex),
            ResolveTerminalLineSpaceX(segments, line, fontResources, fontScale, xOffset, useWordCompatibleTextProfile),
            baselineY,
            0d,
            fontSize,
            PdfCharacterSpacing: 0d,
            PdfCharacterSpacingSource: DocxTextStateCharacterSpacingSource.TerminalLineSpace,
            CompensatePdfCharacterSpacing: true,
            SyntheticBold: false,
            SyntheticItalic: effective.Italic && !resource.Resolution.Italic,
            IsTerminalLineSpace: true,
            segment.SourceTextRunIndex,
            segment.SourceTextOffsetInRun + segment.Text.Length,
            segment.Role));
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateBodyTextLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            switch (item)
            {
                case DocxTextLineLayout line:
                    yield return line;
                    break;
                case DocxTableRowLayout row:
                    foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                    {
                        yield return cellLine;
                    }

                    break;
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateRenderedPageTextLines(
        DocxLayout layout,
        DocxLayoutPage page,
        int pageIndex)
    {
        return EnumerateStaticTextLines(page)
            .Concat(EnumerateBodyTextLines(page))
            .Concat(EnumeratePlacedRelatedStoryTextLines(page))
            .Concat(EnumerateFloatingDrawingTextBoxTextLines(EnumeratePageFloatingDrawings(layout, pageIndex)))
            .Concat(page.PlacedRelatedStories.SelectMany(story => EnumerateFloatingDrawingTextBoxTextLines(story.FloatingDrawings)));
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateMarkupBalloonAnchorTextLines(
        DocxLayoutPage page,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings)
    {
        return EnumerateStaticTextLines(page)
            .Concat(EnumerateBodyTextLines(page))
            .Concat(EnumeratePlacedRelatedStoryTextLines(page))
            .Concat(EnumerateFloatingDrawingTextBoxTextLines(floatingDrawings))
            .Concat(page.PlacedRelatedStories.SelectMany(story => EnumerateFloatingDrawingTextBoxTextLines(story.FloatingDrawings)));
    }

    private sealed record DocxTextEmissionLineSource(
        DocxTextLineLayout Line,
        bool IsStaticStory,
        string StoryKind,
        string? StoryVariantType,
        string? ContainerStoryKind,
        string? ContainerStoryVariantType);

    private static IEnumerable<DocxTextEmissionLineSource> EnumerateRenderedFloatingDrawingTextBoxTextLines(
        DocxLayout layout,
        DocxLayoutPage page,
        int pageIndex)
    {
        foreach (DocxFloatingDrawingLayout drawing in EnumeratePageFloatingDrawings(layout, pageIndex))
        {
            bool isStaticStory = string.Equals(drawing.StoryKind, "Header", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(drawing.StoryKind, "Footer", StringComparison.OrdinalIgnoreCase);
            foreach (DocxTextLineLayout line in EnumerateFloatingDrawingTextBoxTextLines(drawing))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    isStaticStory,
                    "TextBox",
                    line.StoryVariantType,
                    drawing.StoryKind ?? "Body",
                    drawing.StoryVariantType);
            }
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTextLineLayout line in EnumerateFloatingDrawingTextBoxTextLines(story.FloatingDrawings))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    IsStaticStory: false,
                    "TextBox",
                    line.StoryVariantType,
                    story.StoryLayout.Story.Kind,
                    story.StoryLayout.Story.Id);
            }
        }
    }

    private static string ResolveTextEmissionStoryKind(DocxTextLineLayout line, string fallback)
    {
        return string.IsNullOrWhiteSpace(line.StoryKind) ? fallback : line.StoryKind;
    }

    private static IEnumerable<DocxLayoutItem> EnumerateStaticLayoutItems(DocxLayoutPage page)
    {
        return page.StaticTextLines
            .Cast<DocxLayoutItem>()
            .Concat(page.StaticInlineImages)
            .Concat(page.StaticTableRows)
            .OrderByDescending(item => item switch
            {
                DocxTextLineLayout textLine => textLine.BaselineY,
                DocxInlineImageLayout image => image.Y + image.Height,
                DocxTableRowLayout row => row.Y + row.Height,
                _ => 0d
            });
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextLines(DocxLayoutPage page)
    {
        foreach (DocxTextLineLayout line in page.StaticTextLines)
        {
            yield return line;
        }

        foreach (DocxTableRowLayout row in page.StaticTableRows)
        {
            foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
            {
                yield return cellLine;
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumeratePlacedRelatedStoryTextLines(DocxLayoutPage page)
    {
        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (DocxTextLineLayout line in story.TextLines)
            {
                yield return line;
            }

            foreach (DocxTableRowLayout row in story.TableRows)
            {
                foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                {
                    yield return cellLine;
                }
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateFloatingDrawingTextBoxTextLines(IEnumerable<DocxFloatingDrawingLayout> drawings)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            foreach (DocxTextLineLayout line in EnumerateFloatingDrawingTextBoxTextLines(drawing))
            {
                yield return line;
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateFloatingDrawingTextBoxTextLines(DocxFloatingDrawingLayout drawing)
    {
        if (drawing.PlacedX is not { } placedX ||
            drawing.PlacedTop is not { } placedTop ||
            drawing.TextBoxLayout is not { } textBoxLayout)
        {
            yield break;
        }

        foreach (DocxTextLineLayout line in textBoxLayout.TextLines)
        {
            yield return TranslateTextLine(line, placedX, placedTop);
        }

        foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
        {
            foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
            {
                yield return TranslateTextLine(cellLine, placedX, placedTop);
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateFloatingDrawingTextBoxTableRows(IEnumerable<DocxFloatingDrawingLayout> drawings)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            if (drawing.PlacedX is not { } placedX ||
                drawing.PlacedTop is not { } placedTop ||
                drawing.TextBoxLayout is not { } textBoxLayout)
            {
                continue;
            }

            foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(TranslateTableRow(row, placedX, placedTop)))
                {
                    yield return nested;
                }
            }
        }
    }

    private static DocxTableRowLayout TranslateTableRow(DocxTableRowLayout row, double deltaX, double deltaY)
    {
        return row with
        {
            Table = row.Table with
            {
                TableX = row.Table.TableX + deltaX
            },
            Y = row.Y + deltaY,
            Cells = row.Cells
                .Select(cell => TranslateTableCell(cell, deltaX, deltaY))
                .ToArray()
        };
    }

    private static DocxTableCellLayout TranslateTableCell(DocxTableCellLayout cell, double deltaX, double deltaY)
    {
        return cell with
        {
            X = cell.X + deltaX,
            Y = cell.Y + deltaY,
            TextLines = cell.TextLines
                .Select(line => TranslateTextLine(line, deltaX, deltaY))
                .ToArray(),
            InlineImages = cell.InlineImages
                .Select(image => image with { X = image.X + deltaX, Y = image.Y + deltaY })
                .ToArray(),
            NestedTableRows = cell.NestedRows
                .Select(row => TranslateTableRow(row, deltaX, deltaY))
                .ToArray()
        };
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextLines(DocxTableRowLayout row)
    {
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxTextLineLayout line in cell.TextLines)
            {
                yield return line;
            }

            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTextLineLayout nestedLine in EnumerateTableRowTextLines(nestedRow))
                {
                    yield return nestedLine;
                }
            }
        }
    }

    private static DocxTextEmissionLineSnapshot ToTextEmissionLineSnapshot(
        int pageIndex,
        bool isStaticStory,
        string storyKind,
        string? storyVariantType,
        string? containerStoryKind,
        string? containerStoryVariantType,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        int pageNumber,
        int pageCount,
        double fontScale,
        double baselineOffsetY,
        double xOffset,
        bool suppressCommentReferenceSpacer,
        bool useWordCompatibleTextProfile)
    {
        DocxTextEmissionSegmentSnapshot[] segments = CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount, fontScale, baselineOffsetY, xOffset, suppressCommentReferenceSpacer, useWordCompatibleTextProfile)
            .Select(segment => ToTextEmissionSegmentSnapshot(segment, line))
            .ToArray();
        return new DocxTextEmissionLineSnapshot(
            pageIndex,
            isStaticStory,
            storyKind,
            storyVariantType,
            containerStoryKind,
            containerStoryVariantType,
            line.SourceBlockIndex,
            line.SourceParagraphIndex,
            line.SourceLineIndex,
            line.EndsWithIntraTokenBreak,
            segments.Length,
            segments.Sum(segment => segment.TextLength),
            segments.Count(segment => segment.RevisionKind is not null),
            segments.Count(segment => segment.RevisionKind == "Insertion"),
            segments.Count(segment => segment.RevisionKind == "Deletion"),
            segments.Count(segment => segment.RevisionKind == "MoveFrom"),
            segments.Count(segment => segment.RevisionKind == "MoveTo"),
            segments.Count(segment => segment.RevisionKind is not null &&
                segment.RevisionKind != "Insertion" &&
                segment.RevisionKind != "Deletion" &&
                segment.RevisionKind != "MoveFrom" &&
                segment.RevisionKind != "MoveTo"),
            line.SourceParagraph?.InlineReferences.Count(reference => reference.Kind == "Comment") ?? 0,
            segments.Count(segment => segment.IsTerminalLineSpace),
            segments.Count(segment => Math.Abs(segment.PdfCharacterSpacing) > 0.0001d),
            segments);
    }

    private static DocxTextEmissionSegmentSnapshot ToTextEmissionSegmentSnapshot(
        DocxTextEmissionSegment segment,
        DocxTextLineLayout line)
    {
        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(
            segment.StyleRun,
            segment.FontSize,
            segment.PdfCharacterSpacing,
            segment.PdfCharacterSpacingSource,
            segment.CompensatePdfCharacterSpacing,
            segment.IsTerminalLineSpace);
        return new DocxTextEmissionSegmentSnapshot(
            segment.Text.Length,
            line.SourceBlockIndex,
            line.SourceParagraphIndex,
            line.SourceLineIndex,
            segment.Role.ToString(),
            segment.X,
            segment.BaselineY,
            segment.Width,
            segment.FontSize,
            plan.PdfFontSize,
            segment.StyleRun.EffectiveProperties.CharacterSpacingPoints,
            plan.PdfCharacterSpacing,
            plan.PdfCharacterSpacingSource.ToString(),
            plan.PositioningCharacterSpacing,
            plan.CompensatePdfCharacterSpacing,
            DocxTextEmissionPlanner.ClassifyText(segment.Text),
            DocxTextEmissionPlanner.MeasureAdvanceProfile(segment.Text, segment.Resource.Embedded, segment.Width, plan),
            DocxTextEmissionPlanner.CreateGlyphAdvanceSignature(segment.Text, segment.Resource.Embedded),
            segment.IsTerminalLineSpace,
            segment.Resource.Name,
            segment.SyntheticBold,
            segment.SyntheticItalic,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleId,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleFound,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleDepth,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasDocumentDefaultRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasParagraphStyleRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasCharacterStyleRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasDirectRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasTableStyleRunProperties,
            segment.StyleRun.Revision?.Kind,
            segment.StyleRun.Revision?.SourceElement);
    }

    private static void DrawRunGlyphText(
        PdfGraphicsBuilder graphics,
        DocxRunFontResource resource,
        string text,
        double x,
        double baselineY,
        RgbColor color,
        DocxTextEmissionPlan plan,
        bool syntheticItalic)
    {
        string? positioningArray = resource.Embedded.EncodeGlyphPositioningArray(text, plan.PositioningCharacterSpacing, plan.PdfFontSize, forcePositioningArray: true);
        if (positioningArray is not null)
        {
            graphics.DrawGlyphPositionedText(resource.Name, plan.PdfFontSize, x, baselineY, color.Red, color.Green, color.Blue, positioningArray, syntheticItalic, plan.PdfCharacterSpacing);
            return;
        }

        string glyphHex = resource.Embedded.EncodeGlyphHex(text);
        if (glyphHex.Length == 0)
        {
            return;
        }

        graphics.DrawGlyphText(resource.Name, plan.PdfFontSize, x, baselineY, color.Red, color.Green, color.Blue, glyphHex, syntheticItalic, plan.PdfCharacterSpacing);
    }

    private static bool ShouldApplySyntheticBold(DocxTextRun style, DocxRunFontResource resource)
    {
        return style.EffectiveProperties.Bold && !resource.Resolution.Bold;
    }

    private static void RenderRunBackground(
        DocxTextRun style,
        OpenTypeFont font,
        double x,
        double width,
        double fontSize,
        double baselineY,
        PdfGraphicsBuilder graphics)
    {
        if (width <= 0d)
        {
            return;
        }

        double ascender = DocxLineMetrics.MeasureWindowsAscender(font, fontSize);
        double descender = DocxLineMetrics.MeasureWindowsDescender(font, fontSize);
        double fillY = baselineY - descender;
        double fillHeight = ascender + descender;
        DocxEffectiveRunProperties effective = style.EffectiveProperties;
        if (TryResolveHighlightColor(effective.HighlightValue, out RgbColor highlight))
        {
            graphics.SetFillRgb(highlight.Red, highlight.Green, highlight.Blue);
            graphics.FillRectangle(x, fillY, width, fillHeight);
            return;
        }

        RenderShadingFill(effective.ShadingFillHex, effective.ShadingValue, effective.ShadingColor, graphics, x, fillY, width, fillHeight);
    }

    private static bool TryResolveShadingColor(string? fillHex, string? value, string? foregroundHex, out RgbColor color)
    {
        if (value is null || value.Equals("clear", StringComparison.OrdinalIgnoreCase))
        {
            return RgbColor.TryParse(fillHex, out color);
        }

        if (TryResolvePercentageShadingColor(fillHex, value, foregroundHex, out color))
        {
            return true;
        }

        color = default;
        return false;
    }

    private static void RenderShadingFill(
        string? fillHex,
        string? value,
        string? foregroundHex,
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height)
    {
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        if (TryResolveShadingColor(fillHex, value, foregroundHex, out RgbColor solid))
        {
            graphics.SetFillRgb(solid.Red, solid.Green, solid.Blue);
            graphics.FillRectangle(x, y, width, height);
            return;
        }

        if (TryResolveShadingPattern(fillHex, value, foregroundHex, out PdfTilingPattern? pattern) && pattern is not null)
        {
            graphics.FillRectangleWithTilingPattern(x, y, width, height, pattern);
        }
    }

    private static bool TryResolveShadingPattern(string? fillHex, string? value, string? foregroundHex, out PdfTilingPattern? pattern)
    {
        pattern = null;
        if (value is null ||
            !RgbColor.TryParse(fillHex, out RgbColor background) ||
            !RgbColor.TryParse(foregroundHex, out RgbColor foreground))
        {
            return false;
        }

        if (!TryResolveShadingStripeKind(value, out PdfStripePatternKind kind, out bool thin))
        {
            return false;
        }

        pattern = PdfTilingPattern.OfficeBitmapStripeLines(
            kind,
            thin,
            foreground.Red,
            foreground.Green,
            foreground.Blue,
            background.Red,
            background.Green,
            background.Blue);
        return true;
    }

    private static bool TryResolveShadingStripeKind(string value, out PdfStripePatternKind kind, out bool thin)
    {
        switch (value)
        {
            case "horzStripe":
                kind = PdfStripePatternKind.Horizontal;
                thin = false;
                return true;
            case "thinHorzStripe":
                kind = PdfStripePatternKind.Horizontal;
                thin = true;
                return true;
            case "vertStripe":
                kind = PdfStripePatternKind.Vertical;
                thin = false;
                return true;
            case "thinVertStripe":
                kind = PdfStripePatternKind.Vertical;
                thin = true;
                return true;
            case "diagStripe":
                kind = PdfStripePatternKind.DownDiagonal;
                thin = false;
                return true;
            case "thinDiagStripe":
                kind = PdfStripePatternKind.DownDiagonal;
                thin = true;
                return true;
            case "reverseDiagStripe":
                kind = PdfStripePatternKind.UpDiagonal;
                thin = false;
                return true;
            case "thinReverseDiagStripe":
                kind = PdfStripePatternKind.UpDiagonal;
                thin = true;
                return true;
            default:
                kind = PdfStripePatternKind.Horizontal;
                thin = false;
                return false;
        }
    }

    private static bool TryResolvePercentageShadingColor(string? fillHex, string? value, string? foregroundHex, out RgbColor color)
    {
        color = default;
        if (value is null ||
            !value.StartsWith("pct", StringComparison.OrdinalIgnoreCase) ||
            !int.TryParse(value.AsSpan(3), NumberStyles.Integer, CultureInfo.InvariantCulture, out int percent) ||
            !RgbColor.TryParse(fillHex, out RgbColor background) ||
            !RgbColor.TryParse(foregroundHex, out RgbColor foreground))
        {
            return false;
        }

        double weight = Math.Clamp(percent, 0, 100) / 100d;
        color = new RgbColor(
            BlendByte(background.Red, foreground.Red, weight),
            BlendByte(background.Green, foreground.Green, weight),
            BlendByte(background.Blue, foreground.Blue, weight));
        return true;
    }

    private static byte BlendByte(byte background, byte foreground, double foregroundWeight)
    {
        return (byte)Math.Round(background * (1d - foregroundWeight) + foreground * foregroundWeight);
    }

    private static bool TryResolveHighlightColor(string? value, out RgbColor color)
    {
        switch (value)
        {
            case "black":
                color = new RgbColor(0x00, 0x00, 0x00);
                return true;
            case "blue":
                color = new RgbColor(0x00, 0x00, 0xFF);
                return true;
            case "cyan":
                color = new RgbColor(0x00, 0xFF, 0xFF);
                return true;
            case "green":
                color = new RgbColor(0x00, 0xFF, 0x00);
                return true;
            case "magenta":
                color = new RgbColor(0xFF, 0x00, 0xFF);
                return true;
            case "red":
                color = new RgbColor(0xFF, 0x00, 0x00);
                return true;
            case "yellow":
                color = new RgbColor(0xFF, 0xFF, 0x00);
                return true;
            case "white":
                color = new RgbColor(0xFF, 0xFF, 0xFF);
                return true;
            case "darkBlue":
                color = new RgbColor(0x00, 0x00, 0x80);
                return true;
            case "darkCyan":
                color = new RgbColor(0x00, 0x80, 0x80);
                return true;
            case "darkGreen":
                color = new RgbColor(0x00, 0x80, 0x00);
                return true;
            case "darkMagenta":
                color = new RgbColor(0x80, 0x00, 0x80);
                return true;
            case "darkRed":
                color = new RgbColor(0x80, 0x00, 0x00);
                return true;
            case "darkYellow":
                color = new RgbColor(0x80, 0x80, 0x00);
                return true;
            case "darkGray":
                color = new RgbColor(0x80, 0x80, 0x80);
                return true;
            case "lightGray":
                color = new RgbColor(0xC0, 0xC0, 0xC0);
                return true;
            default:
                color = default;
                return false;
        }
    }

    private static void RenderTextDecorations(
        DocxTextRun style,
        PdfEmbeddedFont embedded,
        string text,
        double x,
        double width,
        double fontSize,
        double baselineY,
        RgbColor color,
        DocxTextEmissionPlan plan,
        bool useWordCompatibleRevisionDecorationProfile,
        PdfGraphicsBuilder graphics)
    {
        if (width <= 0d)
        {
            return;
        }

        OpenTypeFont font = embedded.Font;
        DocxEffectiveRunProperties effective = style.EffectiveProperties;
        if (effective.Underline)
        {
            RgbColor underlineColor = ResolveUnderlineDecorationColor(effective, color);
            graphics.SetFillRgb(underlineColor.Red, underlineColor.Green, underlineColor.Blue);
            double thickness = ResolveDecorationThickness(font.Post.UnderlineThickness, font, fontSize, useWordCompatibleRevisionDecorationProfile);
            double y = baselineY + font.Post.UnderlinePosition * fontSize / font.UnitsPerEm;
            RenderUnderlineDecoration(graphics, embedded, text, x, y, width, thickness, fontSize, underlineColor, plan, effective.UnderlineValue);
        }

        if (effective.Strike || effective.DoubleStrike)
        {
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            double thickness = ResolveDecorationThickness(font.Os2.StrikeoutSize, font, fontSize, useWordCompatibleRevisionDecorationProfile);
            double y = baselineY + font.Os2.StrikeoutPosition * fontSize / font.UnitsPerEm;
            if (effective.DoubleStrike)
            {
                double offset = Math.Max(thickness, fontSize / 18d);
                graphics.FillRectangle(x, y - offset - thickness / 2d, width, thickness);
                graphics.FillRectangle(x, y + offset - thickness / 2d, width, thickness);
            }
            else
            {
                graphics.FillRectangle(x, y - thickness / 2d, width, thickness);
            }
        }
    }

    private static RgbColor ResolveUnderlineDecorationColor(DocxEffectiveRunProperties effective, RgbColor textColor)
    {
        return RgbColor.TryParse(effective.UnderlineColorHex, out RgbColor underlineColor)
            ? underlineColor
            : textColor;
    }

    private static void RenderUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        PdfEmbeddedFont embedded,
        string text,
        double x,
        double y,
        double width,
        double thickness,
        double fontSize,
        RgbColor color,
        DocxTextEmissionPlan plan,
        string? underlineValue)
    {
        if (IsWordsUnderlineValue(underlineValue))
        {
            RenderWordsUnderlineDecoration(graphics, embedded, text, x, y, width, thickness, plan);
            return;
        }

        if (IsWaveUnderlineValue(underlineValue))
        {
            RenderWaveUnderlineDecoration(graphics, x, y, width, thickness, fontSize, color, underlineValue);
            return;
        }

        if (IsSegmentedUnderlineValue(underlineValue))
        {
            RenderSegmentedUnderlineDecoration(graphics, x, y, width, thickness, fontSize, underlineValue);
            return;
        }

        if (IsDoubleUnderlineValue(underlineValue))
        {
            double offset = Math.Max(thickness, fontSize / 18d);
            graphics.FillRectangle(x, y - thickness / 2d, width, thickness);
            graphics.FillRectangle(x, y - offset - thickness / 2d, width, thickness);
            return;
        }

        double solidThickness = IsHeavyUnderlineValue(underlineValue)
            ? Math.Max(thickness * 1.35d, 0.3d)
            : thickness;
        graphics.FillRectangle(x, y - solidThickness / 2d, width, solidThickness);
    }

    private static bool IsWordsUnderlineValue(string? underlineValue)
    {
        return underlineValue?.Equals("words", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static void RenderWordsUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        PdfEmbeddedFont embedded,
        string text,
        double x,
        double y,
        double width,
        double thickness,
        DocxTextEmissionPlan plan)
    {
        int index = 0;
        while (index < text.Length)
        {
            while (index < text.Length && char.IsWhiteSpace(text[index]))
            {
                index++;
            }

            int wordStart = index;
            while (index < text.Length && !char.IsWhiteSpace(text[index]))
            {
                index++;
            }

            if (wordStart == index)
            {
                continue;
            }

            double startAdvance = MeasureDecorationTextAdvance(text[..wordStart], embedded, plan);
            double endAdvance = MeasureDecorationTextAdvance(text[..index], embedded, plan);
            double segmentX = x + Math.Min(startAdvance, width);
            double segmentWidth = Math.Min(endAdvance, width) - Math.Min(startAdvance, width);
            if (segmentWidth > 0.001d)
            {
                graphics.FillRectangle(segmentX, y - thickness / 2d, segmentWidth, thickness);
            }
        }
    }

    private static double MeasureDecorationTextAdvance(
        string text,
        PdfEmbeddedFont embedded,
        DocxTextEmissionPlan plan)
    {
        if (text.Length == 0)
        {
            return 0d;
        }

        int glyphGapCount = Math.Max(0, CountMappedGlyphs(text, embedded) - 1);
        return embedded.MeasureTextPoints(text, plan.PdfFontSize, kerningEnabled: true) +
            (plan.PositioningCharacterSpacing + plan.PdfCharacterSpacing) * glyphGapCount;
    }

    private static int CountMappedGlyphs(string text, PdfEmbeddedFont embedded)
    {
        int count = 0;
        foreach (Rune rune in text.EnumerateRunes())
        {
            if (embedded.Font.MapCodePoint(rune.Value) != 0)
            {
                count++;
            }
        }

        return count;
    }

    private static bool IsDoubleUnderlineValue(string? underlineValue)
    {
        return underlineValue is not null &&
            (underlineValue.Equals("double", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dbl", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsWaveUnderlineValue(string? underlineValue)
    {
        return underlineValue is not null &&
            (underlineValue.Equals("wave", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("wavyHeavy", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("wavyDouble", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsSegmentedUnderlineValue(string? underlineValue)
    {
        return underlineValue is not null &&
            (underlineValue.Equals("dash", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dashed", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dashHeavy", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dashedHeavy", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dashLong", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dashLongHeavy", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dotted", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dottedHeavy", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dotDash", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dashDotHeavy", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dotDashHeavy", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dashDotDotHeavy", StringComparison.OrdinalIgnoreCase) ||
            underlineValue.Equals("dotDotDashHeavy", StringComparison.OrdinalIgnoreCase));
    }

    private static void RenderSegmentedUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double thickness,
        double fontSize,
        string? underlineValue)
    {
        double segmentThickness = IsHeavyUnderlineValue(underlineValue)
            ? Math.Max(thickness * 1.35d, 0.3d)
            : thickness;
        double dotLength = Math.Max(segmentThickness, 0.35d);
        double dashLength = IsLongDashUnderlineValue(underlineValue)
            ? Math.Max(fontSize / 2d, segmentThickness * 5d)
            : Math.Max(fontSize / 4d, segmentThickness * 3d);
        double gapLength = Math.Max(segmentThickness * 1.5d, 0.5d);

        if (underlineValue?.Equals("dotDash", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dashDotHeavy", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dotDashHeavy", StringComparison.OrdinalIgnoreCase) == true)
        {
            RenderPatternedUnderlineDecoration(graphics, [dashLength, dotLength], gapLength, x, y, width, segmentThickness);
            return;
        }

        if (underlineValue?.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dashDotDotHeavy", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dotDotDashHeavy", StringComparison.OrdinalIgnoreCase) == true)
        {
            RenderPatternedUnderlineDecoration(graphics, [dashLength, dotLength, dotLength], gapLength, x, y, width, segmentThickness);
            return;
        }

        double segmentLength = underlineValue?.Equals("dotted", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dottedHeavy", StringComparison.OrdinalIgnoreCase) == true
                ? dotLength
                : dashLength;
        RenderPatternedUnderlineDecoration(graphics, [segmentLength], gapLength, x, y, width, segmentThickness);
    }

    private static bool IsHeavyUnderlineValue(string? underlineValue)
    {
        return underlineValue?.Contains("Heavy", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("thick", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool IsLongDashUnderlineValue(string? underlineValue)
    {
        return underlineValue?.Equals("dashLong", StringComparison.OrdinalIgnoreCase) == true ||
            underlineValue?.Equals("dashLongHeavy", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static void RenderPatternedUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        IReadOnlyList<double> segmentLengths,
        double gapLength,
        double x,
        double y,
        double width,
        double thickness)
    {
        double offset = 0d;
        while (offset < width - 0.001d)
        {
            foreach (double segmentLength in segmentLengths)
            {
                double drawLength = Math.Min(segmentLength, width - offset);
                if (drawLength <= 0.001d)
                {
                    return;
                }

                graphics.FillRectangle(x + offset, y - thickness / 2d, drawLength, thickness);
                offset += drawLength + gapLength;
                if (offset >= width - 0.001d)
                {
                    return;
                }
            }
        }
    }

    private static void RenderWaveUnderlineDecoration(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double thickness,
        double fontSize,
        RgbColor color,
        string? underlineValue)
    {
        double lineWidth = underlineValue?.Equals("wavyHeavy", StringComparison.OrdinalIgnoreCase) == true
            ? Math.Max(thickness * 1.35d, 0.3d)
            : Math.Max(thickness, 0.25d);
        double amplitude = Math.Max(lineWidth * 0.85d, fontSize / 28d);
        double halfPeriod = Math.Max(fontSize / 5d, 2d);

        graphics.SaveState();
        graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
        graphics.SetLineWidth(lineWidth);
        graphics.SetLineCap(1);
        graphics.SetLineJoin(1);
        if (underlineValue?.Equals("wavyDouble", StringComparison.OrdinalIgnoreCase) == true)
        {
            double offset = Math.Max(lineWidth + amplitude, fontSize / 16d);
            RenderWaveUnderlineLine(graphics, x, y, width, amplitude * 0.75d, halfPeriod);
            RenderWaveUnderlineLine(graphics, x, y - offset, width, amplitude * 0.75d, halfPeriod);
        }
        else
        {
            RenderWaveUnderlineLine(graphics, x, y, width, amplitude, halfPeriod);
        }

        graphics.RestoreState();
    }

    private static void RenderWaveUnderlineLine(
        PdfGraphicsBuilder graphics,
        double x,
        double centerY,
        double width,
        double amplitude,
        double halfPeriod)
    {
        double major = 0d;
        double previousMinor = -amplitude;
        bool high = true;
        while (major < width - 0.001d)
        {
            double nextMajor = Math.Min(major + halfPeriod, width);
            double nextMinor = high ? amplitude : -amplitude;
            graphics.StrokeLine(x + major, centerY + previousMinor, x + nextMajor, centerY + nextMinor);
            major = nextMajor;
            previousMinor = nextMinor;
            high = !high;
        }
    }

    private static double ResolveDecorationThickness(
        short metricValue,
        OpenTypeFont font,
        double fontSize,
        bool useWordCompatibleRevisionDecorationProfile = false)
    {
        double thickness = Math.Max(0.25d, Math.Abs(metricValue) * fontSize / font.UnitsPerEm);
        return useWordCompatibleRevisionDecorationProfile
            ? Math.Max(WordCompatibleAllMarkupRevisionDecorationThicknessPoints, thickness)
            : thickness;
    }

    private static void RenderTableRowBorders(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        PdfGraphicsBuilder graphics)
    {
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCellLayout cellLayout = row.Cells[cellIndex];
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow))
            {
                continue;
            }

            if (previousRow is null && IsFirstTableRowFragment(row))
            {
                RenderHorizontalTableCellBorder(cellLayout, "top", graphics);
            }

            if (nextRow is null && IsLastTableRowFragment(row))
            {
                RenderHorizontalTableCellBorder(cellLayout, "bottom", graphics);
            }

            DocxTableCell visualCell = cellLayout.VisualCell;
            DocxTableCellBorder? left = DocxTableBorderGeometry.Find(visualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "start");
            if (cellIndex == 0)
            {
                RenderVerticalTableCellBorder(cellLayout.X, cellLayout.Y, cellLayout.Height, left, graphics);
            }

            DocxTableCellBorder? right = DocxTableBorderGeometry.Find(visualCell.Borders, "right") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "end");
            if (cellIndex == row.Cells.Count - 1)
            {
                RenderVerticalTableCellBorder(cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.Height, right, graphics);
                continue;
            }

            DocxTableCellLayout nextCell = row.Cells[cellIndex + 1];
            DocxTableCell nextVisualCell = nextCell.VisualCell;
            DocxTableCellBorder? nextLeft = DocxTableBorderGeometry.Find(nextVisualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(nextVisualCell.Borders, "start");
            RenderSharedVerticalTableBorder(cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.Height, right, nextLeft, graphics);
        }

        if (nextRow is not null && nextRow.RowIndex != row.RowIndex)
        {
            RenderSharedHorizontalTableBorders(row, nextRow, graphics);
        }
    }

    private static bool IsFirstTableRowFragment(DocxTableRowLayout row)
    {
        return row.FragmentIndex <= 0;
    }

    private static bool IsLastTableRowFragment(DocxTableRowLayout row)
    {
        return row.FragmentIndex >= Math.Max(0, row.FragmentCount - 1);
    }

    private static void RenderSharedHorizontalTableBorders(
        DocxTableRowLayout row,
        DocxTableRowLayout nextRow,
        PdfGraphicsBuilder graphics)
    {
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            if (!ShouldRenderTableCellVisualFragment(cellLayout, previousRow: null))
            {
                continue;
            }

            DocxTableCellLayout[] overlappingNextCells = nextRow.Cells
                .Where(nextCell => ShouldRenderTableCellVisualFragment(nextCell, row) && HorizontalOverlap(cellLayout, nextCell) > 0d)
                .ToArray();
            if (overlappingNextCells.Length == 0)
            {
                RenderHorizontalTableCellBorder(cellLayout, "bottom", graphics);
                continue;
            }

            foreach (DocxTableCellLayout nextRowCell in overlappingNextCells)
            {
                RenderSharedHorizontalTableBorderSegment(cellLayout, nextRowCell, graphics);
            }
        }
    }

    private static void RenderSharedHorizontalTableBorderSegment(
        DocxTableCellLayout cellLayout,
        DocxTableCellLayout nextRowCell,
        PdfGraphicsBuilder graphics)
    {
        DocxTableCell cell = cellLayout.VisualCell;
        DocxTableCell nextCell = nextRowCell.VisualCell;
        DocxTableCellBorder? border = ResolveSharedHorizontalBorder(cellLayout, nextRowCell);
        if (border is null)
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        double x = Math.Max(cellLayout.X, nextRowCell.X);
        double right = Math.Min(cellLayout.X + cellLayout.Width, nextRowCell.X + nextRowCell.Width);
        if (right <= x)
        {
            return;
        }

        double leftBorderWidth = ResolveLeftVerticalBorderWidth(cellLayout);
        double segmentX = Math.Min(right, x + leftBorderWidth);
        if (right <= segmentX)
        {
            return;
        }

        RenderTableBorderStrip(graphics, border, segmentX, cellLayout.Y - width / 2d, right - segmentX, width, DocxTableBorderOrientation.Horizontal);
    }

    private static DocxTableCellBorder? ResolveSharedHorizontalBorder(DocxTableCellLayout cellLayout, DocxTableCellLayout nextRowCell)
    {
        DocxTableCellBorder? bottom = DocxTableBorderGeometry.Find(cellLayout.VisualCell.Borders, "bottom");
        DocxTableCellBorder? nextTop = DocxTableBorderGeometry.Find(nextRowCell.VisualCell.Borders, "top");
        return DocxTableBorderGeometry.IsSuppressed(bottom) || DocxTableBorderGeometry.IsSuppressed(nextTop)
            ? null
            : DocxTableBorderGeometry.SelectStronger(bottom, nextTop);
    }

    private static double HorizontalOverlap(DocxTableCellLayout first, DocxTableCellLayout second)
    {
        return Math.Min(first.X + first.Width, second.X + second.Width) - Math.Max(first.X, second.X);
    }

    private static void RenderHorizontalTableCellBorder(DocxTableCellLayout cellLayout, string edge, PdfGraphicsBuilder graphics)
    {
        DocxTableCell visualCell = cellLayout.VisualCell;
        DocxTableCellBorder? border = DocxTableBorderGeometry.Find(visualCell.Borders, edge);
        if (border is null || DocxTableBorderGeometry.IsSuppressed(border))
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        switch (edge)
        {
            case "top":
                double topX = cellLayout.X + ResolveLeftVerticalBorderWidth(cellLayout);
                double topWidth = cellLayout.Width - (topX - cellLayout.X);
                if (topWidth > 0d)
                {
                    RenderTableBorderStrip(graphics, border, topX, cellLayout.Y + cellLayout.Height - width, topWidth, width, DocxTableBorderOrientation.Horizontal);
                }
                break;
            case "bottom":
                double bottomX = cellLayout.X + ResolveLeftVerticalBorderWidth(cellLayout);
                double bottomWidth = cellLayout.Width - (bottomX - cellLayout.X);
                if (bottomWidth > 0d)
                {
                    RenderTableBorderStrip(graphics, border, bottomX, cellLayout.Y, bottomWidth, width, DocxTableBorderOrientation.Horizontal);
                }
                break;
        }
    }

    private static double ResolveLeftVerticalBorderWidth(DocxTableCellLayout cellLayout)
    {
        DocxTableCell visualCell = cellLayout.VisualCell;
        DocxTableCellBorder? left = DocxTableBorderGeometry.Find(visualCell.Borders, "left") ??
            DocxTableBorderGeometry.Find(visualCell.Borders, "start");
        return DocxTableBorderGeometry.IsSuppressed(left)
            ? 0d
            : DocxTableBorderGeometry.ResolveVisibleWidth(left);
    }

    private static void RenderSharedVerticalTableBorder(
        double boundaryX,
        double y,
        double height,
        DocxTableCellBorder? leftCellRight,
        DocxTableCellBorder? rightCellLeft,
        PdfGraphicsBuilder graphics)
    {
        if (DocxTableBorderGeometry.IsSuppressed(leftCellRight) || DocxTableBorderGeometry.IsSuppressed(rightCellLeft))
        {
            return;
        }

        DocxTableCellBorder? border = DocxTableBorderGeometry.SelectStronger(leftCellRight, rightCellLeft);
        if (border is null)
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        RenderTableBorderStrip(graphics, border, boundaryX, y, width, height, DocxTableBorderOrientation.Vertical);
    }

    private static void RenderVerticalTableCellBorder(
        double boundaryX,
        double y,
        double height,
        DocxTableCellBorder? border,
        PdfGraphicsBuilder graphics)
    {
        if (border is null || DocxTableBorderGeometry.IsSuppressed(border))
        {
            return;
        }

        RgbColor color = ReadColor(border.Color);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        RenderTableBorderStrip(graphics, border, boundaryX, y, width, height, DocxTableBorderOrientation.Vertical);
    }

    private static void RenderTableBorderStrip(
        PdfGraphicsBuilder graphics,
        DocxTableCellBorder border,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        string value = border.Value ?? "single";
        if (value.Equals("double", StringComparison.OrdinalIgnoreCase))
        {
            RenderDoubleTableBorderStrip(graphics, x, y, width, height, orientation);
            return;
        }

        if (value.Equals("triple", StringComparison.OrdinalIgnoreCase))
        {
            RenderTripleTableBorderStrip(graphics, x, y, width, height, orientation);
            return;
        }

        if (TryGetCompoundTableBorderProfile(value, out double[] stripeWeights, out double gapWeight))
        {
            RenderCompoundTableBorderStrip(graphics, x, y, width, height, orientation, stripeWeights, gapWeight);
            return;
        }

        if (IsThreeDTableBorderStyle(value))
        {
            RenderThreeDTableBorderStrip(graphics, border, value, x, y, width, height, orientation);
            return;
        }

        if (IsWaveTableBorderStyle(value))
        {
            RenderWaveTableBorderStrip(graphics, border, value, x, y, width, height, orientation);
            return;
        }

        if (value.Equals("dashDotStroked", StringComparison.OrdinalIgnoreCase))
        {
            RenderDashDotStrokedTableBorderStrip(graphics, x, y, width, height, orientation);
            return;
        }

        if (IsSegmentedTableBorderStyle(value))
        {
            RenderSegmentedTableBorderStrip(graphics, value, x, y, width, height, orientation);
            return;
        }

        graphics.FillRectangle(x, y, width, height);
    }

    private static void RenderDoubleTableBorderStrip(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        if (orientation == DocxTableBorderOrientation.Horizontal)
        {
            double stripeHeight = Math.Max(0.12d, height / 3d);
            graphics.FillRectangle(x, y, width, Math.Min(stripeHeight, height));
            graphics.FillRectangle(x, y + Math.Max(0d, height - stripeHeight), width, Math.Min(stripeHeight, height));
            return;
        }

        double stripeWidth = Math.Max(0.12d, width / 3d);
        graphics.FillRectangle(x, y, Math.Min(stripeWidth, width), height);
        graphics.FillRectangle(x + Math.Max(0d, width - stripeWidth), y, Math.Min(stripeWidth, width), height);
    }

    private static void RenderTripleTableBorderStrip(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        double stripe = Math.Min(thickness, Math.Max(0.08d, thickness / 5d));
        double gap = Math.Max(0d, (thickness - 3d * stripe) / 2d);
        for (int index = 0; index < 3; index++)
        {
            double offset = index * (stripe + gap);
            if (orientation == DocxTableBorderOrientation.Horizontal)
            {
                graphics.FillRectangle(x, y + offset, width, Math.Min(stripe, Math.Max(0d, height - offset)));
            }
            else
            {
                graphics.FillRectangle(x + offset, y, Math.Min(stripe, Math.Max(0d, width - offset)), height);
            }
        }
    }

    private static bool TryGetCompoundTableBorderProfile(string value, out double[] stripeWeights, out double gapWeight)
    {
        gapWeight = value.Contains("LargeGap", StringComparison.OrdinalIgnoreCase)
            ? 2d
            : value.Contains("MediumGap", StringComparison.OrdinalIgnoreCase)
                ? 1.25d
                : 0.65d;

        if (value.StartsWith("thinThickThin", StringComparison.OrdinalIgnoreCase))
        {
            stripeWeights = [1d, 2d, 1d];
            return true;
        }

        if (value.StartsWith("thinThick", StringComparison.OrdinalIgnoreCase))
        {
            stripeWeights = [1d, 2d];
            return true;
        }

        if (value.StartsWith("thickThin", StringComparison.OrdinalIgnoreCase))
        {
            stripeWeights = [2d, 1d];
            return true;
        }

        stripeWeights = [];
        return false;
    }

    private static void RenderCompoundTableBorderStrip(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation,
        IReadOnlyList<double> stripeWeights,
        double gapWeight)
    {
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        if (thickness <= 0d || stripeWeights.Count == 0)
        {
            return;
        }

        double totalWeight = stripeWeights.Sum() + gapWeight * Math.Max(0, stripeWeights.Count - 1);
        if (totalWeight <= 0d)
        {
            return;
        }

        double offset = 0d;
        for (int index = 0; index < stripeWeights.Count; index++)
        {
            double stripe = Math.Min(thickness - offset, thickness * stripeWeights[index] / totalWeight);
            if (stripe > 0.001d)
            {
                if (orientation == DocxTableBorderOrientation.Horizontal)
                {
                    graphics.FillRectangle(x, y + offset, width, stripe);
                }
                else
                {
                    graphics.FillRectangle(x + offset, y, stripe, height);
                }
            }

            offset += stripe;
            if (index < stripeWeights.Count - 1)
            {
                offset += thickness * gapWeight / totalWeight;
            }
        }
    }

    private static bool IsThreeDTableBorderStyle(string value)
    {
        return value.Equals("threeDEmboss", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("threeDEngrave", StringComparison.OrdinalIgnoreCase);
    }

    private static void RenderThreeDTableBorderStrip(
        PdfGraphicsBuilder graphics,
        DocxTableCellBorder border,
        string value,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        RgbColor baseColor = ReadColor(border.Color);
        RgbColor light = MixRgbColor(baseColor, new RgbColor(255, 255, 255), 0.55d);
        RgbColor dark = MixRgbColor(baseColor, new RgbColor(0, 0, 0), 0.45d);
        bool emboss = value.Equals("threeDEmboss", StringComparison.OrdinalIgnoreCase);
        RgbColor first = emboss ? light : dark;
        RgbColor second = emboss ? dark : light;

        if (orientation == DocxTableBorderOrientation.Horizontal)
        {
            double firstHeight = Math.Max(0.05d, height * 0.5d);
            double secondHeight = Math.Max(0d, height - firstHeight);
            graphics.SetFillRgb(first.Red, first.Green, first.Blue);
            graphics.FillRectangle(x, y, width, Math.Min(firstHeight, height));
            if (secondHeight > 0.001d)
            {
                graphics.SetFillRgb(second.Red, second.Green, second.Blue);
                graphics.FillRectangle(x, y + firstHeight, width, secondHeight);
            }

            return;
        }

        double firstWidth = Math.Max(0.05d, width * 0.5d);
        double secondWidth = Math.Max(0d, width - firstWidth);
        graphics.SetFillRgb(first.Red, first.Green, first.Blue);
        graphics.FillRectangle(x, y, Math.Min(firstWidth, width), height);
        if (secondWidth > 0.001d)
        {
            graphics.SetFillRgb(second.Red, second.Green, second.Blue);
            graphics.FillRectangle(x + firstWidth, y, secondWidth, height);
        }
    }

    private static bool IsWaveTableBorderStyle(string value)
    {
        return value.Equals("wave", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("doubleWave", StringComparison.OrdinalIgnoreCase);
    }

    private static void RenderWaveTableBorderStrip(
        PdfGraphicsBuilder graphics,
        DocxTableCellBorder border,
        string value,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        RgbColor color = ReadColor(border.Color);
        graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        double lineWidth = Math.Min(thickness, Math.Max(0.08d, thickness / 3d));
        double amplitude = Math.Max(0.08d, Math.Max(0d, thickness - lineWidth) / 2.5d);
        double halfPeriod = Math.Max(1d, thickness * 1.5d);
        graphics.SetLineWidth(lineWidth);

        if (!value.Equals("doubleWave", StringComparison.OrdinalIgnoreCase))
        {
            RenderWaveTableBorderLine(graphics, x, y, width, height, orientation, amplitude, halfPeriod, 0d);
            return;
        }

        double offset = Math.Max(lineWidth, amplitude);
        RenderWaveTableBorderLine(graphics, x, y, width, height, orientation, amplitude * 0.65d, halfPeriod, -offset * 0.45d);
        RenderWaveTableBorderLine(graphics, x, y, width, height, orientation, amplitude * 0.65d, halfPeriod, offset * 0.45d);
    }

    private static void RenderWaveTableBorderLine(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation,
        double amplitude,
        double halfPeriod,
        double minorOffset)
    {
        double majorLength = orientation == DocxTableBorderOrientation.Horizontal ? width : height;
        if (majorLength <= 0.001d)
        {
            return;
        }

        double major = 0d;
        double previousMinor = -amplitude;
        bool high = true;
        while (major < majorLength - 0.001d)
        {
            double nextMajor = Math.Min(major + halfPeriod, majorLength);
            double nextMinor = high ? amplitude : -amplitude;
            if (orientation == DocxTableBorderOrientation.Horizontal)
            {
                double centerY = y + height / 2d + minorOffset;
                graphics.StrokeLine(x + major, centerY + previousMinor, x + nextMajor, centerY + nextMinor);
            }
            else
            {
                double centerX = x + width / 2d + minorOffset;
                graphics.StrokeLine(centerX + previousMinor, y + major, centerX + nextMinor, y + nextMajor);
            }

            major = nextMajor;
            previousMinor = nextMinor;
            high = !high;
        }
    }

    private static bool IsSegmentedTableBorderStyle(string value)
    {
        return value.Equals("dotted", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dashed", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dashSmallGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dotDash", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase);
    }

    private static void RenderDashDotStrokedTableBorderStrip(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        double dotLength = Math.Max(thickness, 0.35d);
        double dashLength = Math.Max(thickness * 3d, 1d);
        double gapLength = Math.Max(thickness * 1.5d, 0.5d);
        RenderPatternedTableBorderStrip(graphics, [dashLength, dotLength], gapLength, x, y, width, height, orientation);

        double strokeThickness = Math.Min(thickness, Math.Max(0.05d, thickness / 5d));
        if (orientation == DocxTableBorderOrientation.Horizontal)
        {
            graphics.FillRectangle(x, y + Math.Max(0d, (height - strokeThickness) / 2d), width, strokeThickness);
        }
        else
        {
            graphics.FillRectangle(x + Math.Max(0d, (width - strokeThickness) / 2d), y, strokeThickness, height);
        }
    }

    private static void RenderSegmentedTableBorderStrip(
        PdfGraphicsBuilder graphics,
        string value,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        double thickness = orientation == DocxTableBorderOrientation.Horizontal ? height : width;
        double dotLength = Math.Max(thickness, 0.35d);
        double dashLength = Math.Max(thickness * 3d, 1d);
        double gapLength = value.Equals("dashSmallGap", StringComparison.OrdinalIgnoreCase)
            ? Math.Max(thickness, 0.35d)
            : Math.Max(thickness * 1.5d, 0.5d);

        if (value.Equals("dotDash", StringComparison.OrdinalIgnoreCase))
        {
            RenderPatternedTableBorderStrip(graphics, [dashLength, dotLength], gapLength, x, y, width, height, orientation);
            return;
        }

        if (value.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase))
        {
            RenderPatternedTableBorderStrip(graphics, [dashLength, dotLength, dotLength], gapLength, x, y, width, height, orientation);
            return;
        }

        double segmentLength = value.Equals("dotted", StringComparison.OrdinalIgnoreCase)
            ? dotLength
            : dashLength;
        RenderPatternedTableBorderStrip(graphics, [segmentLength], gapLength, x, y, width, height, orientation);
    }

    private static void RenderPatternedTableBorderStrip(
        PdfGraphicsBuilder graphics,
        IReadOnlyList<double> segmentLengths,
        double gapLength,
        double x,
        double y,
        double width,
        double height,
        DocxTableBorderOrientation orientation)
    {
        double majorLength = orientation == DocxTableBorderOrientation.Horizontal ? width : height;
        double offset = 0d;
        while (offset < majorLength - 0.001d)
        {
            foreach (double segmentLength in segmentLengths)
            {
                double drawLength = Math.Min(segmentLength, majorLength - offset);
                if (drawLength <= 0.001d)
                {
                    return;
                }

                if (orientation == DocxTableBorderOrientation.Horizontal)
                {
                    graphics.FillRectangle(x + offset, y, drawLength, height);
                }
                else
                {
                    graphics.FillRectangle(x, y + offset, width, drawLength);
                }

                offset += drawLength;
                if (offset >= majorLength - 0.001d)
                {
                    return;
                }

                offset += gapLength;
            }
        }
    }

    private enum DocxTableBorderOrientation
    {
        Horizontal,
        Vertical
    }

    private static bool ShouldRenderTableCellContentFragment(
        DocxTableCellLayout cellLayout,
        DocxTableRowLayout? previousRow)
    {
        return cellLayout.VisualOwnership == DocxTableCellVisualOwnership.OwnCell ||
            cellLayout.VisualOwnership == DocxTableCellVisualOwnership.VerticalMergeOwner &&
            ShouldRenderTableCellVisualFragment(cellLayout, previousRow);
    }

    private static bool ShouldRenderTableCellVisualFragment(
        DocxTableCellLayout cellLayout,
        DocxTableRowLayout? previousRow)
    {
        if (cellLayout.VisualOwnership == DocxTableCellVisualOwnership.OwnCell)
        {
            return true;
        }

        if (cellLayout.VisualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner)
        {
            return false;
        }

        return previousRow is null || !previousRow.Cells.Any(previousCell =>
            !previousCell.IsVerticalMergeContinuation &&
            HorizontalOverlap(previousCell, cellLayout) > 0d &&
            previousCell.Y <= cellLayout.Y + 0.001d &&
            previousCell.Y + previousCell.Height >= cellLayout.Y + cellLayout.Height - 0.001d);
    }

    private sealed record DocxTableBorderBoundary(double X, double Width, DocxTableCellBorder Border);

    private static string ResolveStaticFieldPlaceholders(string text, int pageNumber, int pageCount)
    {
        return text
            .Replace("{NUMPAGES}", pageCount.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
            .Replace("{PAGE}", pageNumber.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    private static double ResolveSubstitutedFieldEmissionWidth(
        string sourceText,
        string emittedText,
        DocxTextRun styleRun,
        double fontSize,
        IDocxTextMeasurer? textMeasurer,
        double fallbackWidth)
    {
        if (string.Equals(sourceText, emittedText, StringComparison.Ordinal) ||
            textMeasurer is null ||
            (!sourceText.Contains("{PAGE}", StringComparison.Ordinal) &&
            !sourceText.Contains("{NUMPAGES}", StringComparison.Ordinal)))
        {
            return fallbackWidth;
        }

        double measured = textMeasurer.MeasureText(styleRun, emittedText, fontSize);
        return measured > 0d ? measured : fallbackWidth;
    }

    private static RgbColor ReadColor(string? hex)
    {
        return RgbColor.TryParse(hex, out RgbColor color) ? color : new RgbColor(0, 0, 0);
    }

    private static RgbColor MixRgbColor(RgbColor source, RgbColor target, double targetWeight)
    {
        double clampedTargetWeight = Math.Clamp(targetWeight, 0d, 1d);
        double sourceWeight = 1d - clampedTargetWeight;
        return new RgbColor(
            (byte)Math.Round(source.Red * sourceWeight + target.Red * clampedTargetWeight),
            (byte)Math.Round(source.Green * sourceWeight + target.Green * clampedTargetWeight),
            (byte)Math.Round(source.Blue * sourceWeight + target.Blue * clampedTargetWeight));
    }

    private static PdfImageXObject? CreateImage(DocxInlineImage image, Action<OoxPdfDiagnostic>? diagnosticSink, int pageIndex)
    {
        try
        {
            if (image.ContentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                image.ContentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
            {
                JpegInfo info = JpegInfo.Read(image.Bytes);
                return PdfImageXObject.Jpeg(info.Width, info.Height, image.Bytes, info.ComponentCount, info.BitsPerComponent);
            }

            if (image.ContentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
            {
                PngImage png = PngImage.Read(image.Bytes);
                return PdfImageXObject.RgbPng(png.Width, png.Height, png.Rgb, png.Alpha);
            }

            if (image.ContentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
                image.ContentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase))
            {
                BmpImage bmp = BmpImage.Read(image.Bytes);
                return PdfImageXObject.RgbPng(bmp.Width, bmp.Height, bmp.Rgb, bmp.Alpha);
            }
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            EmitImageDiagnostic(diagnosticSink, image, pageIndex, ex.Message);
            return null;
        }

        EmitImageDiagnostic(diagnosticSink, image, pageIndex, "Unsupported image content type.");
        return null;
    }

    private static void EmitImageDiagnostic(Action<OoxPdfDiagnostic>? diagnosticSink, DocxInlineImage image, int pageIndex, string reason)
    {
        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            "IMAGE_UNSUPPORTED_FORMAT",
            OoxPdfSeverity.Error,
            $"Image '{image.ContentType}' could not be rendered and was ignored: {reason}",
            image.PartName,
            PageIndex: pageIndex,
            Feature: image.ContentType,
            Fallback: "Ignored"));
    }

}
