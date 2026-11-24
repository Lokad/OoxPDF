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

    private const double WordCompatibleAllMarkupMaxBodyTextFontSizePoints = 11.625d;
    // Unmeasured balloon titles land on the anchor row minus ~0.5 (Office: dense/c14/mirrored/threaded refs).
    // = doc-adaptive baseline shift + title rule (top inset 11.15 - height 20.48 + first baseline 11.27 + 0.5).
    private const double WordCompatibleAllMarkupBalloonAnchorRowInsetPoints = 2.44d;
    private const double WordCompatibleAllMarkupBalloonHeightPoints = 20.48d;
    private const double WordCompatibleAllMarkupBalloonTopInsetPoints = 11.15d;
    private const double WordCompatibleAllMarkupConnectorStrokeWidthPoints = 0.475d;
    private const double WordCompatibleAllMarkupConnectorBodyAnchorInsetPoints = 3.18d;
    // Measured uniform balloon rows use font boxes plus review-design padding.
    private const double WordCompatibleAllMarkupBalloonTextPaddingDesignPoints = 2.25d;
    // Office A/B (W5-X1 right-margin probes w5-xr72/w5-xr207 plus dense/landscape refs):
    // Word balloon text starts at (bodyEndDesign + gap) times the lane-fit scale
    // (implied gap 35.1/35.6/36.0/35.8 across the four datasets).
    private const double WordCompatibleAllMarkupBalloonLaneGapPoints = 35.6d;
    // Office A/B (margin-variant R72/R207 probes plus mirrored/dense/landscape/author references,
    // Word-COM rendered across five print scales): balloon bodies are 230.2pt design wide
    // (emission widths 170.73/174.58/178.80/184.92/209.88 at scales 0.7423/0.7575/0.7762/0.8028/0.9114).
    // The prior 233pt value fitted our own candidate widths, never an Office measurement.
    private const double WordCompatibleAllMarkupBalloonBodyWidthPoints = 230.2d;
    private const double WordCompatibleAllMarkupBalloonLaneBackgroundBalloonInsetPoints = 22.56d;
    // Office A/B (unresolved b2 wrap failure plus all/original marginal first rows,
    // Word-COM rendered): balloon body rows stop 1.02..4.01pt before the body rect
    // edge, so the first line keeps a ~2.5pt trailing pad past the 0.5pt wrap edge.
    // This preserves the pre-fix wrap widths bit-for-bit while the emission origin
    // above moves to the Office-true title end.
    private const double WordCompatibleAllMarkupBalloonBodyFirstLineTrailingPadPoints = 2.541d;
    // RV06-L37: trusted font/margin boundary controls give an overlapping 221.33
    // to 221.405pt design-width interval when the continuation's final blank is
    // excluded. Office text insets measure 4.42..4.44 design points; using 4.42
    // on each side gives 221.36pt and keeps the 41-character Courier guard out.
    private const double WordCompatibleAllMarkupBalloonTextInsetDesignPoints = 4.42d;
    // Legacy insets remain for mixed rows and unmeasured/fontless fallbacks.
    private const double WordCompatibleAllMarkupBalloonFirstBaselineOffsetPoints = 11.27d;
    private const double WordCompatibleAllMarkupBalloonFirstBaselineTopInsetPoints = WordCompatibleAllMarkupBalloonHeightPoints - WordCompatibleAllMarkupBalloonFirstBaselineOffsetPoints;
    // Office A/B (tbxrev/dense): mixed/unmeasured bodies retain the established
    // first-baseline inset and 3.4pt last-baseline padding.
    private const double WordCompatibleAllMarkupBalloonBottomInsetPoints = 3.4d;
    private const double WordCompatibleAllMarkupLaneBackgroundRightBleedPoints = 0.37d;
    private const double WordCompatibleAllMarkupLaneBackgroundWidthPoints = 199.70d;
    private const double WordCompatibleAllMarkupRevisionBarMaximumTextDistancePoints = 36d;
    private const double WordCompatibleAllMarkupRevisionBarWidthPoints = 0.72d;
    private const double WordCompatibleAllMarkupRevisionBarBottomOutsetPoints = 7.47d;
    private const double WordCompatibleAllMarkupRevisionBarTopInsetPoints = 0.39d;
    private const double WordCompatibleAllMarkupRevisionDecorationThicknessPoints = 0.475d;
    private const double WordCompatibleAllMarkupShortInsertionDecorationWidthInsetPoints = 1.434d;
    private const double WordCompatibleAllMarkupCommentRangeStrokeWidthPoints = 0.475d;
    private const double MarkupBalloonConnectorCollisionAnchorYThresholdPoints = 9d;
    // Office A/B (dense/landscape/unresolved references, Word-COM rendered): packed
    // consecutive balloons sit ~0.87pt apart (dense 0.81-0.94, landscape 0.87-0.96,
    // unresolved 0.82); the unevidenced June 3pt minimum pushed packed balloons ~2pt low.
    private const double MarkupBalloonMinimumSpacingPoints = 0.87d;
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

    // R06.3: yields pages progressively so the writer can spill content and retain
    // only descriptors; callers must drain inside the conversion budget scope.
    public async IAsyncEnumerable<PdfPage> RenderPagesAsync(DocxDocument document, Action<OoxPdfDiagnostic>? diagnosticSink, [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!HasRenderableContent())
        {
            ReportEquationDiagnostics(document, [], diagnosticSink, cancellationToken);
            // R06.2: admit the single blank page like any produced page.
            OoxConversionBudget.Current?.ChargePdfPages(1);
            yield return new PdfPage(document.PageWidthPoints, document.PageHeightPoints);
            yield break;
        }

        await foreach (PdfPage page in RenderParagraphsAsync(document, fontResolver, ResolveEffectiveMarkupContext(document, cancellationToken), diagnosticSink, cancellationToken))
        {
            yield return page;
        }

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

    internal async ValueTask<(DocxLayoutSnapshot Layout, DocxFontPlanSnapshot FontPlan,
        DocxTextEmissionSnapshot TextEmission, IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> MarkupBalloons)>
        InspectAsync(DocxDocument document, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var context = ResolveEffectiveMarkupContext(document, cancellationToken);
        var geometry = ResolveEffectiveMarkupGeometryMode(context);
        var prepared = await PreparePageNumberedLayoutAsync(document, fontResolver, context, geometry,
            null, cancellationToken).ConfigureAwait(false);
        var layout = DocxLayoutSnapshot.FromLayout(prepared.Layout, prepared.Document.MarkupMode, geometry);
        cancellationToken.ThrowIfCancellationRequested();
        // Ordinary font-plan inspection describes source requests before font
        // substitution. Page-numbered notes instead describe the final plan.
        var fontPlan = DocxPageFootnoteNumbering.IsEligibleDocument(document)
            ? DocxFontPlanSnapshot.FromPlan(prepared.FontResources.Plan, cancellationToken)
            : InspectSourceFontPlan(document, context, geometry, cancellationToken);
        var textEmission = CreateTextEmissionSnapshot(prepared.Document, prepared.FontResources,
            prepared.Layout, context, cancellationToken);
        var balloons = await CreateMarkupBalloonSnapshotsAsync(prepared.Document, prepared.FontResources,
            prepared.Layout, context, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return (layout, fontPlan, textEmission, balloons);
    }

    internal DocxLayoutSnapshot InspectLayout(DocxDocument document)
    {
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        OoxPdfDocxMarkupGeometryMode effectiveGeometryMode = ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext);
        var prepared = PreparePageNumberedLayoutAsync(document, fontResolver, effectiveMarkupContext, effectiveGeometryMode, null, CancellationToken.None).GetAwaiter().GetResult();
        document = prepared.Document;
        DocxLayout layout = prepared.Layout;
        return DocxLayoutSnapshot.FromLayout(layout, document.MarkupMode, effectiveGeometryMode);
    }

    internal IReadOnlyList<DocxMarkupBalloonPlacementSnapshot> InspectMarkupBalloons(DocxDocument document)
    {
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        var prepared = PreparePageNumberedLayoutAsync(document, fontResolver, effectiveMarkupContext, ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext), null, CancellationToken.None).GetAwaiter().GetResult();
        return CreateMarkupBalloonSnapshotsAsync(prepared.Document, prepared.FontResources, prepared.Layout,
            effectiveMarkupContext, CancellationToken.None).GetAwaiter().GetResult();
    }

    private async ValueTask<IReadOnlyList<DocxMarkupBalloonPlacementSnapshot>> CreateMarkupBalloonSnapshotsAsync(
        DocxDocument document, DocxFontResources fontResources, DocxLayout layout,
        DocxMarkupContext effectiveMarkupContext, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        bool separatePlainHeaderOrigins = CanSeparatePlainHeaderOrigins(layout);
        if (UsesWordCompatibleAllMarkupTextProfile(effectiveMarkupContext) && HasWordCompatibleBalloonContent(document, effectiveMarkupContext))
        {
            effectiveMarkupContext = effectiveMarkupContext with { BalloonTitleFace = await ResolveBalloonTitleFaceAsync(fontResolver, effectiveMarkupContext, cancellationToken).ConfigureAwait(false) };
        }
        (DocxRunFontResource? _, DocxRunFontResource? balloonTitleResource, DocxRunFontResource? balloonBodyResource,
            DocxRunFontResource? revisionBodyResource) = await EnsureMarkupBalloonTextResourcesAsync(
                layout, fontResources, effectiveMarkupContext, fontResolver, cancellationToken).ConfigureAwait(false);
        effectiveMarkupContext = effectiveMarkupContext with { RevisionBalloonBodyResource = revisionBodyResource };
        var snapshots = new List<DocxMarkupBalloonPlacementSnapshot>();
        DocxRunFontResource? balloonLabelResource = ResolveMarkupLabelFontResource(fontResources);
        PdfEmbeddedFont? balloonLabelEmbedded = balloonTitleResource?.Embedded ?? balloonLabelResource?.Embedded;
        PdfEmbeddedFont? balloonBodyEmbedded = (balloonBodyResource ?? ResolveMarkupBodyFontResource(fontResources) ?? balloonLabelResource)?.Embedded;
        FloatingDrawingPageIndex.PageIndexPair drawingPages = FloatingDrawingPageIndex.BuildPair(layout, cancellationToken);
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (DocxMarkupBalloonPlacement placement in BuildMarkupBalloonPlacements(
                layout.Pages[pageIndex],
                layout.RelatedStories,
                drawingPages.PageAll(pageIndex),
                WithPageTextEmissionOffsets(effectiveMarkupContext, layout, pageIndex, separatePlainHeaderOrigins),
                balloonLabelEmbedded,
                balloonBodyEmbedded,
                fontResources,
                cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();
                snapshots.Add(placement.ToSnapshot(pageIndex));
            }
        }

        return snapshots;
    }

    internal DocxFontPlanSnapshot InspectFontPlan(DocxDocument document)
    {
        var context = ResolveEffectiveMarkupContext(document);
        var geometry = ResolveEffectiveMarkupGeometryMode(context);
        if (DocxPageFootnoteNumbering.IsEligibleDocument(document))
        {
            var prepared = PreparePageNumberedLayoutAsync(document, fontResolver, context, geometry, null, CancellationToken.None).GetAwaiter().GetResult();
            return DocxFontPlanSnapshot.FromPlan(prepared.FontResources.Plan);
        }
        return InspectSourceFontPlan(document, context, geometry, CancellationToken.None);
    }

    private DocxFontPlanSnapshot InspectSourceFontPlan(DocxDocument document, DocxMarkupContext context,
        OoxPdfDocxMarkupGeometryMode geometry, CancellationToken cancellationToken)
    {
        document = DocxFinalSectionParagraph.Materialize(document, cancellationToken,
            geometry == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ? context.WordCompatiblePrintScale : 1d);
        return DocxFontPlanSnapshot.FromPlan(DocxFontPlan.Create(document, fontResolver, cancellationToken,
            context.CommentMarkerLabels is not null), cancellationToken);
    }

    internal DocxTextEmissionSnapshot InspectTextEmission(DocxDocument document)
    {
        DocxMarkupContext effectiveMarkupContext = ResolveEffectiveMarkupContext(document);
        var prepared = PreparePageNumberedLayoutAsync(document, fontResolver, effectiveMarkupContext, ResolveEffectiveMarkupGeometryMode(effectiveMarkupContext), null, CancellationToken.None).GetAwaiter().GetResult();
        return CreateTextEmissionSnapshot(prepared.Document, prepared.FontResources, prepared.Layout,
            effectiveMarkupContext, CancellationToken.None);
    }

    private static DocxTextEmissionSnapshot CreateTextEmissionSnapshot(DocxDocument document,
        DocxFontResources fontResources, DocxLayout layout, DocxMarkupContext effectiveMarkupContext,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        bool separatePlainHeaderOrigins = CanSeparatePlainHeaderOrigins(layout);
        double textEmissionFontScale = ResolveTextEmissionFontScale(effectiveMarkupContext);
        double textEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(effectiveMarkupContext);
        double textEmissionXOffset = ResolveTextEmissionXOffset(effectiveMarkupContext);
        bool useWordCompatibleTextProfile = UsesWordCompatibleAllMarkupTextProfile(effectiveMarkupContext);
        var lines = new List<DocxTextEmissionLineSnapshot>();
        FloatingDrawingPageIndex.PageIndexPair drawingPages = FloatingDrawingPageIndex.BuildPair(layout, cancellationToken);
        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutPage page = layout.Pages[pageIndex];
            var pageNumber = page.GetPageNumber(pageIndex + 1);
            effectiveMarkupContext = WithPageTextEmissionOffsets(effectiveMarkupContext, layout, pageIndex, separatePlainHeaderOrigins);
            textEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(effectiveMarkupContext);
            textEmissionXOffset = ResolveTextEmissionXOffset(effectiveMarkupContext);
            void AddLine(
                DocxTextLineLayout line,
                bool isStaticStory,
                string fallbackStoryKind,
                string? containerStoryKind,
                string? containerStoryVariantType)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string ResolveTextEmissionStoryKind()
                {
                    // T03: unstoried lines take the enumeration-group fallback.
                    return line.Story?.ToKindString() ?? fallbackStoryKind;
                }

                lines.Add(ToTextEmissionLineSnapshot(
                    pageIndex,
                    isStaticStory,
                    ResolveTextEmissionStoryKind(),
                    line.Story?.VariantType,
                    containerStoryKind,
                    containerStoryVariantType,
                    line,
                    fontResources,
                    pageNumber,
                    layout.Pages.Count,
                    textEmissionFontScale,
                    textEmissionBaselineOffset,
                    textEmissionXOffset,
                    useWordCompatibleTextProfile));
            }

            foreach (DocxTextLineLayout line in EnumerateStaticTextLines(page, effectiveMarkupContext))
            {
                AddLine(line, isStaticStory: true, "Static", line.Story?.ToKindString(), line.Story?.VariantType);
            }

            foreach (DocxTextLineLayout line in EnumerateBodyTextLines(page))
            {
                AddLine(line, isStaticStory: false, "Body", "Body", null);
            }

            // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote and endnote
            // story lines map uniformly like their emission so snapshots agree with
            // balloon anchors; other stories keep legacy coordinates.
            FloatingTextBoxEmissionMap? placedStoryMap = TryCreateFloatingTextBoxEmissionMap(effectiveMarkupContext, page.Height);
            foreach (DocxTextLineLayout line in EnumerateMappedPlacedNoteStoryTextLines(page, placedStoryMap))
            {
                AddLine(line, isStaticStory: false, "RelatedStory", line.Story?.ToKindString(), line.Story?.VariantType);
            }

            foreach (DocxTextEmissionLineSource source in EnumerateRenderedFloatingDrawingTextBoxTextLines(drawingPages, page, pageIndex, effectiveMarkupContext, page.Height))
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
                    useWordCompatibleTextProfile));
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
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

    private static async ValueTask<(DocxDocument Document, DocxFontResources FontResources, DocxLayout Layout)> PreparePageNumberedLayoutAsync(
        DocxDocument document, IFontResolver resolver, DocxMarkupContext context, OoxPdfDocxMarkupGeometryMode geometry,
        Action<OoxPdfDiagnostic>? diagnostics, CancellationToken cancellationToken)
    {
        document = DocxFinalSectionParagraph.Materialize(document, cancellationToken,
            geometry == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ? context.WordCompatiblePrintScale : 1d);
        List<OoxPdfDiagnostic>? messages = diagnostics is null ? null : [];
        // Source identities/face indices retain immutable programs throughout
        // this preparation. Changed numbering/kerning needs new subsets; a count
        // retry can retain unchanged preparation. Ownership ends with
        // this conversion; nothing is shared across converters or resolvers.
        var fontCache = new Dictionary<(string StableId, int FaceIndex), OpenTypeFont?>();
        var first = await PrepareNoteAdjustedLayoutAsync(document, resolver, context, geometry, messages is null ? null : messages.Add, cancellationToken,
            fontCache).ConfigureAwait(false);
        if (first.FontResources.DeferredPageCountEquations.Count != 0 ||
            (DocxFieldText.GetDependencies(document.BodyElements, cancellationToken) & DocxCounterDependencies.PageCount) != 0 ||
            document.RelatedStories.Any(story => story.IsNormalStoryType &&
                story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
                (DocxFieldText.GetDependencies(story.BodyElements, cancellationToken) & DocxCounterDependencies.PageCount) != 0))
        {
            List<OoxPdfDiagnostic>? retryMessages = diagnostics is null ? null : [];
            // Decimal field glyphs are already in every subset. Reuse only the
            // same source with complete math and no warnings needing replay;
            // changed note owners or deferred math keep full preparation.
            var reusableFonts = ReferenceEquals(first.Document, document) &&
                first.FontResources.DeferredPageCountEquations.Count == 0 &&
                (messages is null || messages.Count == 0) ? first.FontResources : null;
            // Recompute note-dependent purposes too: accepting a stable count
            // must not leave numbering or kerning derived from the omitted math.
            int expectedPageCount = first.Layout.Pages.Count;
            var retry = await PrepareNoteAdjustedLayoutAsync(document, resolver, context, geometry, retryMessages is null ? null : retryMessages.Add,
                cancellationToken, fontCache, expectedPageCount, reusableFonts).ConfigureAwait(false);
            if (retry.Layout.Pages.Count != expectedPageCount)
            {
                // Restored deferred content can change pagination. Confirm that
                // new count once before rejecting complete math in favor of the
                // initial fallback. A further change retains the existing fallback;
                // this is a bounded confirmation, not an open-ended fixed point.
                cancellationToken.ThrowIfCancellationRequested();
                expectedPageCount = retry.Layout.Pages.Count;
                retryMessages = diagnostics is null ? null : [];
                retry = await PrepareNoteAdjustedLayoutAsync(document, resolver, context, geometry,
                    retryMessages is null ? null : retryMessages.Add, cancellationToken, fontCache, expectedPageCount).ConfigureAwait(false);
            }
            if (retry.Layout.Pages.Count == expectedPageCount)
            {
                var painted = new HashSet<DocxEquation>(EnumeratePaintedEquations(retry.Layout, cancellationToken), ReferenceEqualityComparer.Instance);
                if (first.FontResources.DeferredPageCountEquations.All(painted.Contains))
                {
                    first = retry;
                    messages = retryMessages;
                }
            }
        }
        if (diagnostics is not null && messages is not null)
            // Bounded font/layout retries may repeat the same preparation
            // warning. Preserve distinct source/page diagnostics in first order.
            foreach (var message in messages.Distinct()) { cancellationToken.ThrowIfCancellationRequested(); diagnostics(message); }
        // SVG labels do not affect paragraph pagination. Prepare their borrowed
        // outline programs once, after count/note retries, through the same owner.
        var svgFonts = await PptxRenderer.PrepareSvgTextFontsAsync(
            DocxBlockTraversal.EnumerateResourceImages(first.Document, cancellationToken)
                .Where(image => image.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
                .Select(image => image.Bytes).OfType<byte[]>(),
            resolver, (face, token) => LoadFontAsync(face, fontCache, token), diagnostics, cancellationToken).ConfigureAwait(false);
        return (first.Document, first.FontResources with { SvgTextFonts = svgFonts },
            ReframeFloatingParagraphTops(first.Layout, context, cancellationToken));
    }

    private static async ValueTask<(DocxDocument Document, DocxFontResources FontResources, DocxLayout Layout)> PrepareNoteAdjustedLayoutAsync(
        DocxDocument document, IFontResolver resolver, DocxMarkupContext context, OoxPdfDocxMarkupGeometryMode geometry,
        Action<OoxPdfDiagnostic>? diagnostics, CancellationToken cancellationToken,
        Dictionary<(string StableId, int FaceIndex), OpenTypeFont?> fontCache, int? pageCount = null,
        DocxFontResources? initialFonts = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var fonts = initialFonts ?? await PrepareFontResourcesCoreAsync(document, resolver, diagnostics, cancellationToken, context.CommentMarkerLabels is not null, pageCount, fontCache).ConfigureAwait(false);
        var first = CreatePlainFootnoteTextLayout(document, fonts, context, geometry, cancellationToken, pageCount);
        DocxDocument adjusted = document;
        IReadOnlyDictionary<string, int> originalPages = new Dictionary<string, int>();
        IReadOnlyDictionary<string, DocxNoteKerning.PageRange> notePages = new Dictionary<string, DocxNoteKerning.PageRange>();
        DocxNoteKerning.JointPageRanges? jointNotePages = null;
        DocxRelatedStoryKind noteKind = DocxRelatedStoryKind.Endnote;
        bool plainNotes = false;
        if (fonts.TextMeasurer is not null && fonts.FallbackFaces.Count == 0)
        {
            if (DocxNoteKerning.TryCreateJointDocument(document, first, out adjusted, out jointNotePages, cancellationToken))
                plainNotes = true;
            else
            {
                foreach (DocxRelatedStoryKind kind in document.RelatedStories.Where(s => s.IsNormalStoryType &&
                    s.Kind is DocxRelatedStoryKind.Endnote or DocxRelatedStoryKind.Footnote).Select(s => s.Kind).Distinct())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!DocxNoteKerning.TryCreateDocument(document, first, out adjusted, out notePages, cancellationToken, kind)) continue;
                    noteKind = kind;
                    plainNotes = true;
                    break;
                }
            }
        }
        if (plainNotes ||
            context.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup && context.WordCompatiblePrintScale == 1d &&
            fonts.TextMeasurer is not null && fonts.FallbackFaces.Count == 0 &&
            (DocxPlainFootnoteKerning.TryCreateDocument(document, first, out adjusted, out originalPages, cancellationToken) ||
                DocxWrappedCalibriKerning.TryCreateDocument(document, first, out adjusted, out originalPages,
                    DocxWrappedCalibriPolicy.SourceThresholds, DocxFootnoteOwnerScope.AllSourceReferences, cancellationToken)))
        {
            var adjustedFonts = await PrepareFontResourcesCoreAsync(adjusted, resolver, diagnostics, cancellationToken, context.CommentMarkerLabels is not null, pageCount, fontCache).ConfigureAwait(false);
            var adjustedLayout = CreatePlainFootnoteTextLayout(adjusted, adjustedFonts, context, geometry, cancellationToken, pageCount);
            bool wrappedCalibri = !DocxPlainBodyFrame.HasSingleLineParagraphs(first) &&
                DocxWrappedFootnoteFrame.IsEligible(adjusted, DocxWrappedCalibriPolicy.NormalizedThresholds, DocxFootnoteOwnerScope.AllSourceReferences);
            bool stableOwners = plainNotes
                ? jointNotePages is not null
                    ? DocxNoteKerning.HasStableJointOwners(adjustedLayout, jointNotePages, cancellationToken)
                    : DocxNoteKerning.HasStableOwners(adjustedLayout, notePages, cancellationToken, noteKind)
                : (DocxPlainBodyFrame.HasSingleLineParagraphs(adjustedLayout) || wrappedCalibri) &&
                    (originalPages.Count == 0 || DocxPageFootnoteNumbering.HasStablePages(adjusted, adjustedLayout, originalPages,
                        allowWrappedBody: wrappedCalibri, calibriPolicy: DocxWrappedCalibriPolicy.NormalizedThresholds,
                        ownerScope: DocxFootnoteOwnerScope.AllSourceReferences));
            // Layout admission owns the plain note purpose. Corrected wrapping
            // may change source placements and note-only pages while preserving
            // complete owners. Other retries retain their total-page guard.
            if ((plainNotes || adjustedLayout.Pages.Count == first.Pages.Count) && stableOwners)
            {
                document = adjusted;
                fonts = adjustedFonts;
                first = adjustedLayout;
            }
        }
        var numberingPolicy = context.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            context.WordCompatiblePrintScale == 1d ? DocxWrappedCalibriPolicy.NormalizedThresholds :
            DocxWrappedCalibriPolicy.NormalizedNominal;
        if (fonts.TextMeasurer is null ||
            !DocxPageFootnoteNumbering.TryCreateDocument(document, first, out var numbered, out var pages,
                allowWrappedBody: context.WordCompatiblePrintScale == 1d,
                calibriPolicy: numberingPolicy, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences, cancellationToken: cancellationToken))
            return (document, fonts, first);
        // Corrected labels need their own glyph subset and measured advances.
        // One bounded retry is accepted only when body and complete note pages
        // remain stable; otherwise the original resources/layout retain fallback.
        var numberedFonts = await PrepareFontResourcesCoreAsync(numbered, resolver, diagnostics, cancellationToken, context.CommentMarkerLabels is not null, pageCount, fontCache).ConfigureAwait(false);
        var final = CreatePlainFootnoteTextLayout(numbered, numberedFonts, context, geometry, cancellationToken, pageCount);
        return final.Pages.Count == first.Pages.Count && DocxPageFootnoteNumbering.HasStableNumberingPages(numbered, final, pages,
            allowWrappedBody: context.WordCompatiblePrintScale == 1d,
            calibriPolicy: numberingPolicy, ownerScope: DocxFootnoteOwnerScope.AllSourceReferences, cancellationToken: cancellationToken)
            ? (numbered, numberedFonts, final) : (document, fonts, first);
    }

    private static DocxLayout CreatePlainFootnoteTextLayout(DocxDocument document, DocxFontResources fonts,
        DocxMarkupContext context, OoxPdfDocxMarkupGeometryMode geometry, CancellationToken cancellationToken, int? pageCount = null)
    {
        var layout = CreateCompleteNoteReservedLayout(document, fonts, context, geometry, cancellationToken, pageCount);
        bool unscaledMinimumLeading = context.WordCompatiblePrintScale == 1d &&
            (context.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
                context.Mode == OoxPdfDocxMarkupMode.Final);
        return unscaledMinimumLeading && fonts.TextMeasurer is not null && fonts.FallbackFaces.Count == 0 &&
            DocxMinimumBodyLineLeading.TryCreateLayout(document, layout, out var adjusted, cancellationToken) ? adjusted : layout;
    }

    private static DocxLayout CreateCompleteNoteReservedLayout(DocxDocument document, DocxFontResources fonts,
        DocxMarkupContext context, OoxPdfDocxMarkupGeometryMode geometry, CancellationToken cancellationToken, int? pageCount = null)
    {
        var first = CreateHeaderDisplacedLayout(document, fonts, context, geometry, cancellationToken, pageCount: pageCount);
        if (first.UsesSourceOwnedFootnoteReserve ||
            context.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup || context.WordCompatiblePrintScale != 1d ||
            fonts.TextMeasurer is null || fonts.FallbackFaces.Count != 0 ||
            !DocxPageFootnoteNumbering.IsEligibleDocument(document, includeContinuousRestart: true) ||
            document.Paragraphs.Any(p => p.KeepRules.KeepNext == true || p.KeepRules.KeepLines == true) ||
            !DocxPageFootnoteNumbering.HasCompleteNoteOwners(document, first)) return first;
        var head = first.RelatedStories.Where(s => s.Story.Kind == DocxRelatedStoryKind.Footnote &&
            s.Story.Type == DocxRelatedStoryType.Separator).ToArray();
        if (head.Length != 1 || head[0].TextLines.Count != 1 ||
            !head[0].TextLines[0].Text.All(char.IsWhiteSpace)) return first;
        // A single bounded retry can reclaim only unused reserve. A moved or
        // split note must keep the original layout and its numbering admission.
        var complete = CreateHeaderDisplacedLayout(document, fonts, context, geometry, cancellationToken, useCompleteFootnoteReserve: true, pageCount: pageCount);
        return DocxPageFootnoteNumbering.HasStableCompleteNoteOwners(document, first, complete) ? complete : first;
    }

    private static DocxLayout CreateHeaderDisplacedLayout(
        DocxDocument document,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        OoxPdfDocxMarkupGeometryMode geometryMode,
        CancellationToken cancellationToken,
        bool useCompleteFootnoteReserve = false, int? pageCount = null)
    {
        // Office A/B (w37/w39 multi-paragraph headers plus w47 cross-story spacing,
        // Word-COM rendered): body starts below overflowing header content, with the
        // header trailing after-spacing swallowed at the story boundary (w47: header
        // after-24 plus body before-12 applies 12). The first pass lays out body and
        // statics exactly as before; only pages whose header content bottom (trailing
        // cursor including trailing after-spacing) falls below the body top are laid
        // out again with a per-page start displacement. Fitting headers cost nothing
        // and change nothing (empty map returns the first layout untouched). Footer
        // content tops raise the body frame symmetrically (w48 multi-line footer plus
        // long body: body breaks before footer content instead of overlapping it).
        // R06.2 repagination control: at most two full layouts per call (first pass
        // plus one conditional displaced pass; no loop), and page/content charges apply
        // once per emitted final page in RenderParagraphs, never per pass.
        DocxLayoutEngine engine = new(geometryMode, markupContext.WordCompatiblePrintScale, markupContext.CommentMarkerLabels,
            useCompleteFootnoteReserve, requestedGeometryMode: markupContext.GeometryMode, knownPageCount: pageCount)
        {
            PlacementRelatedStoryMeasurements = new()
        };
        IDocxTextMeasurer? scaledTextMeasurer = ResolveLayoutTextMeasurer(fontResources, markupContext);
        DocxLayout first = engine.Create(document, scaledTextMeasurer, cancellationToken, fontResources.TextMeasurer, mathEquations: fontResources.MathEquations);
        // Maps stay dense: explicitly unconstrained pages record zero so the
        // beyond-map last-entry fallback in Create only fires past the first layout's
        // page count, never for a recorded storyless page (w53 even-pages probe:
        // page two has no stories, so it must start at the top with a full frame even
        // though page one displaces by 33.27 and raises by 156.73).
        Dictionary<int, double> headerDisplacementByPage = [];
        Dictionary<int, double> footerDisplacementByPage = [];
        for (int pageIndex = 0; pageIndex < first.Pages.Count; pageIndex++)
        {
            DocxLayoutPage page = first.Pages[pageIndex];
            double bodyTop = page.Height - page.MarginTop;
            double headerBottom = first.HeaderContentBottomByPage.TryGetValue(pageIndex, out double bottom)
                ? bottom
                : bodyTop;
            headerDisplacementByPage[pageIndex] = headerBottom < bodyTop ? bodyTop - headerBottom : 0d;

            double footerTop = first.FooterContentTopByPage.TryGetValue(pageIndex, out double top)
                ? top
                : double.NegativeInfinity;
            footerDisplacementByPage[pageIndex] = footerTop > page.MarginBottom ? footerTop - page.MarginBottom : 0d;
        }

        return headerDisplacementByPage.Values.All(value => value == 0d) && footerDisplacementByPage.Values.All(value => value == 0d)
            ? first
            : engine.Create(document, scaledTextMeasurer, cancellationToken, fontResources.TextMeasurer, headerDisplacementByPage, footerDisplacementByPage, fontResources.MathEquations);
    }

    private DocxMarkupContext ResolveEffectiveMarkupContext(DocxDocument document, CancellationToken cancellationToken = default)
    {
        DocxMarkupContext effective = markupContext.ApplyDocumentSettings(document.Settings);
        bool wordCompatible = UsesWordCompatibleAllMarkupTextProfile(effective);
        bool? hasPrintedBalloons = wordCompatible ? HasWordCompatibleBalloonContent(document, effective) : null;
        double printScale = ResolveWordCompatiblePrintScale(document, effective, hasPrintedBalloons);
        // Office: body X lands at margin times scale, so the uniform shift is margin times (scale - 1).
        double xOffset = wordCompatible ? -document.MarginLeftPoints * (1d - printScale) : 0d;
        // The Y shift defaults to the fitted anchor here; scaled pages refine it from the
        // owning page's laid-out baseline once a layout exists (WithPageTextEmissionOffsets).
        double yOffset = wordCompatible ? WordCompatibleTextYOffsetReferenceAnchorPoints : 0d;
        IReadOnlyDictionary<string, string>? markerLabels = wordCompatible && effective.RendersCommentBalloons
            ? BuildWordCompatibleCommentMarkerLabels(document, cancellationToken) : null;
        IReadOnlySet<DocxParagraph>? visibleComments = wordCompatible
            ? new HashSet<DocxParagraph>(DocxBlockTraversal.EnumerateBodyParagraphs(document).Where(paragraph =>
                paragraph.InlineReferences.Any(reference => reference.Kind == DocxRelatedStoryKind.Comment) ||
                paragraph.SplitComments is not null), ReferenceEqualityComparer.Instance)
            : null;
        return effective with { WordCompatiblePrintScale = printScale, WordCompatibleTextXOffset = xOffset,
            WordCompatibleTextYOffset = yOffset, CommentMarkerLabels = markerLabels,
            WordCompatibleMainStoryCommentParagraphs = visibleComments,
            WordCompatibleHasPrintedBalloons = hasPrintedBalloons };
    }

    private const double WordCompatibleBalloonLaneWidthPoints = 266.5d;

    internal static double ResolveWordCompatiblePrintScale(DocxDocument document, DocxMarkupContext markupContext,
        bool? hasPrintedBalloons = null)
    {
        if (!UsesWordCompatibleAllMarkupTextProfile(markupContext) ||
            !(hasPrintedBalloons ?? HasWordCompatibleBalloonContent(document, markupContext)))
        {
            return 1d;
        }

        // Office A/B (print-with-balloons probes across margins, orientations, and content):
        // Word reserves a fixed balloon lane beside the body and scales the page so
        // left margin + body + lane fits the paper width.
        double bodyWidth = document.PageWidthPoints - document.MarginLeftPoints - document.MarginRightPoints;
        double designWidth = document.MarginLeftPoints + bodyWidth + WordCompatibleBalloonLaneWidthPoints;
        if (designWidth <= 0d)
        {
            return 1d;
        }

        // A page that already has room for the lane retains its authored size.
        return Math.Min(1d, document.PageWidthPoints / designWidth);
    }

    private static bool HasWordCompatibleBalloonContent(DocxDocument document, DocxMarkupContext markupContext)
    {
        if (markupContext.RendersCommentBalloons &&
            HasBalloonableCommentAnchor(document))
        {
            return true;
        }

        return markupContext.RendersRevisionBalloons && HasAnyRevisionBalloon(document);
    }

    // Independent Word references omit header/footer/note comments, including
    // their table cells, and retain main-body comments in mixed documents.
    // Earlier textbox probes also omit those comments. Only matched main-story
    // anchors reserve the review lane; orphan references reserve nothing.
    private static bool HasBalloonableCommentAnchor(DocxDocument document)
    {
        HashSet<string> commentPartIds = document.RelatedStories
            .Where(story => story.Kind == DocxRelatedStoryKind.Comment && story.Id is not null)
            .Select(story => story.Id ?? string.Empty)
            .ToHashSet(StringComparer.Ordinal);
        if (commentPartIds.Count == 0)
        {
            return false;
        }

        bool HasCommentReference(DocxParagraph paragraph)
        {
            return paragraph.InlineReferences.Any(reference =>
                reference.Kind == DocxRelatedStoryKind.Comment &&
                reference.Id is not null &&
                commentPartIds.Contains(reference.Id) &&
                !DocxInlineContent.IsKnownEmptyCommentAnchor(paragraph, reference));
        }

        return DocxBlockTraversal.EnumerateBodyParagraphs(document).Any(HasCommentReference);
    }

    private static bool HasRevisionBalloon(IReadOnlyList<DocxRevisionInfo> revisions, bool isMainStory = true)
    {
        return revisions.Any(revision => IsWordCompatibleBalloonRevision(revision, isMainStory));
    }

    // Page fitting and painting use the same revision owners. A note part cannot
    // claim the page lane without a main-story reference; comments and separator
    // stories do not independently place revision captions.
    private static bool HasAnyRevisionBalloon(DocxDocument document)
    {
        if (HasContent(document.BodyElements, isMainStory: true) ||
            HasFloatingContent(document.FloatingDrawings))
        {
            return true;
        }

        foreach (var story in DocxBlockTraversal.EnumerateStaticStorySources(document, default))
        {
            if (DocxLayoutEngine.IsStaticStoryVariantEnabled(story.Settings, story.Variant) &&
                (HasContent(story.BodyElements, isMainStory: false) || HasFloatingContent(story.Drawings)))
            {
                return true;
            }
        }

        var references = DocxBlockTraversal.EnumerateBodyParagraphs(document)
            .SelectMany(paragraph => paragraph.InlineReferences)
            .Where(reference => reference.Id is not null &&
                reference.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote)
            .Select(reference => (reference.Kind, reference.Id))
            .ToHashSet();
        return document.RelatedStories.Any(story => story.IsNormalStoryType && references.Contains((story.Kind, story.Id)) &&
            (HasContent(story.BodyElements, isMainStory: false) || HasFloatingContent(story.FloatingDrawings)));

        bool HasFloatingContent(IEnumerable<DocxFloatingDrawing> drawings) =>
            drawings.Any(drawing => HasContent(drawing.TextBoxBodyElements, isMainStory: false));

        bool HasContent(IEnumerable<DocxBodyElement> elements, bool isMainStory, bool includeInlineTextBoxes = true) =>
            elements.Any(element => element switch
            {
                DocxParagraphElement paragraph => HasParagraph(paragraph.Paragraph, isMainStory, includeInlineTextBoxes),
                DocxTableElement table => HasTable(table.Table, isMainStory, includeInlineTextBoxes),
                _ => false
            });

        bool HasParagraph(DocxParagraph paragraph, bool isMainStory, bool includeInlineTextBoxes = true) =>
            HasRevisionBalloon(paragraph.Revisions, isMainStory) ||
            // Inline boxes are distinct stories; boxes nested inside them are
            // ignored by layout and cannot reserve a lane for unpainted captions.
            includeInlineTextBoxes && paragraph.InlineTextBoxes.Any(box =>
                HasContent(box.BodyElements, isMainStory: false, includeInlineTextBoxes: false));

        bool HasTable(DocxTable table, bool isMainStory, bool includeInlineTextBoxes)
        {
            if (HasRevisionBalloon(table.Revisions, isMainStory)) return true;
            foreach (DocxTableRow row in table.Rows)
            {
                if (HasRevisionBalloon(row.Revisions, isMainStory)) return true;
                foreach (DocxTableCell cell in row.Cells)
                {
                    if (HasRevisionBalloon(cell.Revisions, isMainStory) ||
                        HasContent(DocxTableCellContent.GetBodyElements(cell), isMainStory, includeInlineTextBoxes)) return true;
                }
            }

            return false;
        }
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
        return markupContext.WordCompatiblePrintScale;
    }

    // Fitted pre-layout/lineless fallback. Measured pages use their own printed
    // frame, including its identity transform at unit scale.
    private const double WordCompatibleTextYOffsetReferenceAnchorPoints = 69.58d;
    internal static double ResolveWordCompatibleTextYOffset(DocxMarkupContext markupContext, double? firstBaselineY, double pageHeight, bool hasBalloonContent, double sourceBaselineOffset = 0d)
    {
        // Office A/B (W5-Y1 top-margin probes w5-ytop54/w5-ytop144): Word scales Y about the
        // page center with the print scale. Pin the laid-out first baseline to its center-scaled
        // position with a slope-1 shift (the pre-shrunk layout keeps its own pitch), which fits
        // the probes to 0.1pt and dense to 0.2pt. Unscaled balloonless pages keep no shift.
        if (markupContext.Mode != OoxPdfDocxMarkupMode.AllMarkup ||
            markupContext.GeometryMode != OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup ||
            !markupContext.ExpandsMarkupMargin)
        {
            return 0d;
        }

        // Unit-scale pages without a body origin retain physical coordinates.
        // Balloons elsewhere cannot supply an origin for note-only continuations.
        if (Math.Abs(markupContext.WordCompatiblePrintScale - 1d) < 0.000000001d &&
            (!hasBalloonContent || firstBaselineY is null))
        {
            return 0d;
        }

        if (firstBaselineY is null)
        {
            return WordCompatibleTextYOffsetReferenceAnchorPoints;
        }

        return new DocxPrintedPageFrame(pageHeight, markupContext.WordCompatiblePrintScale)
            .BodyYOffset(firstBaselineY.Value, sourceBaselineOffset);
    }

    private static bool CanSeparatePlainHeaderOrigins(DocxLayout layout) =>
        layout.Pages.All(page => page.ColumnFrames.Count <= 1 &&
            page.StaticTextLines.Count == 1 && page.StaticTextLines[0].Story?.Kind == DocxStoryKind.Header &&
            page.StaticInlineImages.Count == 0 && page.StaticTableRows.Count == 0 &&
            page.StaticInlineTextBoxes.Count == 0);

    private static DocxMarkupContext WithPageTextEmissionOffsets(
        DocxMarkupContext markupContext, DocxLayout layout, int pageIndex, bool separatePlainHeaderOrigins)
    {
        if (!UsesWordCompatibleAllMarkupTextProfile(markupContext))
        {
            return markupContext;
        }

        // Layout already scales advances. Restore this page's source origin;
        // another page's heading or margins cannot translate its body or links.
        DocxLayoutPage page = layout.Pages[pageIndex];
        double xOffset = -page.MarginLeft * (1d - markupContext.WordCompatiblePrintScale);
        var (firstBaselineY, sourceBaselineOffset) = ResolveBodyPrintOrigin(page);
        double? firstBodyBaselineY = firstBaselineY;
        double bodySourceBaselineOffset = sourceBaselineOffset;
        foreach (DocxTextLineLayout line in EnumerateStaticTextLines(page, includeTextBoxes: false))
        {
            if (firstBaselineY is null || line.BaselineY > firstBaselineY.Value)
                (firstBaselineY, sourceBaselineOffset) = (line.BaselineY, 0d);
        }

        bool hasBalloons = markupContext.WordCompatibleHasPrintedBalloons == true;
        if ((page.StaticTextLines.Any(line => line.StaticStoryUsesPhysicalY) ||
            page.StaticTableRows.Any(row => row.StaticStoryUsesPhysicalY) ||
            page.StaticInlineTextBoxes.Any(box => box.StaticStoryUsesPhysicalY)) && firstBodyBaselineY is { } bodyBaseline)
        {
            double headerDisplacement = layout.HeaderContentBottomByPage.TryGetValue(pageIndex, out double headerBottom)
                ? Math.Max(0d, page.Height - page.MarginTop - headerBottom) : 0d;
            return markupContext with
            {
                WordCompatibleTextYOffset = ResolveWordCompatibleTextYOffset(markupContext,
                    bodyBaseline + headerDisplacement, page.Height, hasBalloons, bodySourceBaselineOffset),
                WordCompatibleTextXOffset = xOffset,
                WordCompatibleStaticTextYAdjustmentPoints = 0d
            };
        }

        double legacyOffset = ResolveWordCompatibleTextYOffset(markupContext, firstBaselineY, page.Height, hasBalloons, sourceBaselineOffset);
        // Word independently scales the first body and header origins. Their
        // continuation pitches are already scaled by layout. A higher header
        // must not move the body down by the unscaled distance between them.
        // Preserve the established static origin, and keep composed static
        // stories and table-first or multi-column bodies on their prior path.
        bool useBodyAnchor = markupContext.WordCompatiblePrintScale < 1d && firstBodyBaselineY is not null &&
            page.Items.FirstOrDefault() is DocxTextLineLayout &&
            separatePlainHeaderOrigins;
        double bodyOffset = useBodyAnchor
            ? ResolveWordCompatibleTextYOffset(markupContext, firstBodyBaselineY, page.Height, hasBalloons, bodySourceBaselineOffset)
            : legacyOffset;
        return markupContext with
        {
            WordCompatibleTextYOffset = bodyOffset,
            WordCompatibleTextXOffset = xOffset,
            WordCompatibleStaticTextYAdjustmentPoints = bodyOffset - legacyOffset
        };
    }

    private static double ResolveTextEmissionBaselineOffset(DocxMarkupContext markupContext)
    {
        return markupContext.WordCompatibleTextYOffset;
    }

    private static double ResolveTextEmissionXOffset(DocxMarkupContext markupContext)
    {
        return markupContext.WordCompatibleTextXOffset;
    }

    private static bool UsesWordCompatibleAllMarkupTextProfile(DocxMarkupContext markupContext)
    {
        return markupContext.Mode == OoxPdfDocxMarkupMode.AllMarkup &&
            markupContext.GeometryMode == OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup &&
            markupContext.ExpandsMarkupMargin;
    }

    internal static double ResolveLayoutLineMetricScale(DocxMarkupContext markupContext)
    {
        // Office dense pitch (13.18 at 12pt single spacing) equals per-doc-scale metrics,
        // not the fitted compromise: line metrics live in the same emission space as text
        // widths (the old W5-P2 veto predates the per-font evidence; R207-class re-probing
        // stays queued). Unscaled profiles keep scale 1 through the print scale itself.
        return ResolveTextEmissionFontScale(markupContext);
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

    private static async IAsyncEnumerable<PdfPage> RenderParagraphsAsync(
        DocxDocument document,
        IFontResolver fontResolver,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var prepared = await PreparePageNumberedLayoutAsync(document, fontResolver, markupContext, ResolveEffectiveMarkupGeometryMode(markupContext), diagnosticSink, cancellationToken).ConfigureAwait(false);
        document = prepared.Document;
        DocxFontResources fontResources = prepared.FontResources;
        DocxLayout layout = prepared.Layout;
        ReportEquationDiagnostics(document, EnumeratePaintedEquations(layout, cancellationToken), diagnosticSink, cancellationToken, layout);
        // RV06: comment balloons and range washes wear first-seen author colors;
        // the slots ride the markup context so static emission helpers can resolve
        // them per comment reference without signature changes.
        DocxReviewAuthorPalette? reviewAuthors = UsesWordCompatibleAllMarkupTextProfile(markupContext)
            ? BuildWordReviewAuthorPalette(document, SelectCommentAuthorsForPalette(layout.RelatedStories), cancellationToken) : null;
        markupContext = markupContext with
        {
            WordReviewAuthorPalette = reviewAuthors,
            NonRenderingRevisionParagraphs = markupContext.DrawsChangeBars && UsesWordCompatibleAllMarkupTextProfile(markupContext)
                ? new(ReferenceEqualityComparer.Instance) : null,
            CommentAuthorPaletteSlots = reviewAuthors?.CommentSlots ??
                BuildCommentAuthorPaletteSlots(SelectCommentAuthorsForPalette(layout.RelatedStories))
        };
        // RV16: comment/reply story lookup rides the markup context so per-page
        // candidate collection resolves stories without rebuilding group dictionaries
        // or re-sorting replies for every page. Skipped when comment balloons never
        // render (the placement path early-outs before consulting it).
        markupContext = markupContext with { CommentStoryIndex = markupContext.RendersCommentBalloons ? MarkupCommentStoryIndex.Build(layout.RelatedStories) : null };
        // RV06: balloon titles resolve Segoe UI Bold once per conversion when balloons exist;
        // the face rides the markup context into the balloon text subset below.
        if (UsesWordCompatibleAllMarkupTextProfile(markupContext) && HasWordCompatibleBalloonContent(document, markupContext))
        {
            markupContext = markupContext with { BalloonTitleFace = await ResolveBalloonTitleFaceAsync(fontResolver, markupContext, cancellationToken).ConfigureAwait(false) };
        }
        bool separatePlainHeaderOrigins = CanSeparatePlainHeaderOrigins(layout);
        (DocxRunFontResource? balloonTextResource, DocxRunFontResource? balloonTitleResource, DocxRunFontResource? balloonBodyResource,
            DocxRunFontResource? revisionBodyResource) = await EnsureMarkupBalloonTextResourcesAsync(
                layout, fontResources, markupContext, fontResolver, cancellationToken).ConfigureAwait(false);
        markupContext = markupContext with { RevisionBalloonBodyResource = revisionBodyResource };
        double textEmissionFontScale = ResolveTextEmissionFontScale(markupContext);
        double textEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(markupContext);
        double textEmissionXOffset = ResolveTextEmissionXOffset(markupContext);
        bool useWordCompatibleTextProfile = UsesWordCompatibleAllMarkupTextProfile(markupContext);
        FloatingDrawingPageIndex.PageIndexPair drawingPages = FloatingDrawingPageIndex.BuildPair(layout, cancellationToken);
        IReadOnlyDictionary<string, PdfLinkDestination> bookmarkDestinations = CreateBookmarkDestinations();
        IReadOnlyDictionary<(DocxRelatedStoryKind Kind, string Id), PdfLinkDestination> noteDestinations =
            CreateNoteDestinations(layout, markupContext, cancellationToken);
        IReadOnlyDictionary<DocxParagraph, NoteReferenceSpacing> noteSpacingParagraphs =
            noteDestinations.Count == 0 ? new Dictionary<DocxParagraph, NoteReferenceSpacing>() :
                CreateNoteSpacingParagraphs(document, layout, markupContext, fontResources, cancellationToken);
        int imageIndex = 1;
        var imageCache = new Dictionary<string, PdfImageXObject?>();

        for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // R06.2: admit the upcoming page before it emits; later pages, fonts,
            // and images never produce past the trip.
            OoxConversionBudget.Current?.ChargePdfPages(1);
            DocxLayoutPage layoutPage = layout.Pages[pageIndex];
            if (layoutPage.IsBlankParityPage)
            {
                yield return new PdfPage(layoutPage.Width, layoutPage.Height);
                continue;
            }
            markupContext = WithPageTextEmissionOffsets(markupContext, layout, pageIndex, separatePlainHeaderOrigins);
            textEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(markupContext);
            textEmissionXOffset = ResolveTextEmissionXOffset(markupContext);
            using IMemoryReservation? construction = ReservePageConstruction(layoutPage,
                drawingPages.Floating.Get(pageIndex).All, drawingPages.Static.Get(pageIndex).All, cancellationToken, out long svgSourceBytes);
            cancellationToken.ThrowIfCancellationRequested();
            var graphics = new PdfGraphicsBuilder(construction, svgSourceBytes);
            var pageImages = new List<PdfImageResource>();
            // Source identities record successful inline paint on this page.
            // Vectors and selected previews need the same navigation contract
            // as raster resources; no payload or decode result is copied here.
            var navigation = new PictureNavigation(new PdfRectangle(0d, 0d, layoutPage.Width, layoutPage.Height),
                new HashSet<DocxInlineImage>(ReferenceEqualityComparer.Instance), []);
            var pageNumber = layoutPage.GetPageNumber(pageIndex + 1);
            RenderWordCompatibleMarkupLaneBackground(layoutPage, graphics, markupContext);
            RenderFloatingDrawings(
                drawingPages.Floating.Get(pageIndex).Behind,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                cancellationToken,
                imageCache, ref imageIndex,
                navigation);
            RenderFloatingDrawings(
                drawingPages.Static.Get(pageIndex).Behind,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                cancellationToken,
                imageCache, ref imageIndex,
                navigation);

            IReadOnlyList<DocxLayoutItem> staticItems = EnumerateStaticLayoutItems(layoutPage, markupContext).ToArray();
            RenderParagraphBackgrounds(staticItems, graphics, markupContext, cancellationToken);
            for (int itemIndex = 0; itemIndex < staticItems.Count; itemIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxLayoutItem staticItem = staticItems[itemIndex];
                DocxTableRowLayout? previousRow = itemIndex > 0 ? staticItems[itemIndex - 1] as DocxTableRowLayout : null;
                DocxTableRowLayout? nextRow = itemIndex + 1 < staticItems.Count ? staticItems[itemIndex + 1] as DocxTableRowLayout : null;
                RenderLayoutItem(staticItem, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, layout.Pages.Count, cancellationToken, imageCache, ref imageIndex, navigation);
            }

            RenderParagraphBackgrounds(layoutPage.Items, graphics, markupContext, cancellationToken);
            for (int itemIndex = 0; itemIndex < layoutPage.Items.Count; itemIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxLayoutItem item = layoutPage.Items[itemIndex];
                DocxTableRowLayout? previousRow = itemIndex > 0 ? layoutPage.Items[itemIndex - 1] as DocxTableRowLayout : null;
                DocxTableRowLayout? nextRow = itemIndex + 1 < layoutPage.Items.Count ? layoutPage.Items[itemIndex + 1] as DocxTableRowLayout : null;
                RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, layout.Pages.Count, cancellationToken, imageCache, ref imageIndex, navigation);
            }

            RenderWordCompatibleRevisionBar(drawingPages, layoutPage, pageIndex, graphics, markupContext, cancellationToken);
            RenderMarkupBalloons(
                layoutPage,
                layout.RelatedStories,
                drawingPages.PageAll(pageIndex),
                graphics,
                fontResources,
                markupContext,
                balloonTextResource,
                balloonTitleResource,
                balloonBodyResource,
                construction,
                cancellationToken);

            PlacedNoteTableFrame?[]? noteTables = null;
            for (int storyIndex = 0; storyIndex < layoutPage.PlacedRelatedStories.Count; storyIndex++)
            {
                var story = layoutPage.PlacedRelatedStories[storyIndex];
                cancellationToken.ThrowIfCancellationRequested();
                RenderPlacedRelatedStoryDrawings(story, behindDocument: true, graphics, pageImages, fontResources, markupContext, pageNumber, layout.Pages.Count, diagnosticSink, cancellationToken, imageCache, ref imageIndex, navigation);
                if (RenderPlacedRelatedStory(story, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, layout.Pages.Count,
                    cancellationToken, imageCache, ref imageIndex, navigation) is { } noteTable)
                {
                    noteTables ??= new PlacedNoteTableFrame?[layoutPage.PlacedRelatedStories.Count];
                    noteTables[storyIndex] = noteTable;
                }
                RenderPlacedRelatedStoryDrawings(story, behindDocument: false, graphics, pageImages, fontResources, markupContext, pageNumber, layout.Pages.Count, diagnosticSink, cancellationToken, imageCache, ref imageIndex, navigation);
            }

            RenderFloatingDrawings(
                drawingPages.Floating.Get(pageIndex).Ahead,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                cancellationToken,
                imageCache, ref imageIndex,
                navigation);
            RenderFloatingDrawings(
                drawingPages.Static.Get(pageIndex).Ahead,
                graphics,
                pageImages,
                fontResources,
                markupContext,
                pageNumber,
                layout.Pages.Count,
                diagnosticSink,
                cancellationToken,
                imageCache, ref imageIndex,
                navigation);

            IReadOnlyList<PdfLinkAnnotation> annotations = CreateHyperlinkAnnotations(
                layoutPage, pageIndex, pageNumber, layout.Pages.Count, navigation.InlinePictures, staticItems, noteTables, navigation.DrawingUris);
            string content = graphics.ToString();
            // RV11: admit the emitted content bytes before yielding, so a small
            // content budget rejects before the writer encodes or stores the page.
            // DOCX layout itself stays whole-document; repagination control is a named residual.
            OoxConversionBudget.Current?.ChargePdfContentBytes(content.Length);
            yield return new PdfPage(
                layoutPage.Width,
                layoutPage.Height,
                content,
                fontResources.Resources.Where(resource => graphics.UsedFontResourceNames.Contains(resource.ResourceName)).ToArray(),
                pageImages.ToArray(),
                graphics.ExtGStates,
                graphics.Shadings,
                graphics.Patterns,
                annotations,
                fontResources.FallbackFontResources,
                graphics.Groups,
                graphics.ShadingPatterns);
        }

        IReadOnlyDictionary<string, PdfLinkDestination> CreateBookmarkDestinations()
        {
            var destinations = new Dictionary<string, PdfLinkDestination>(StringComparer.Ordinal);
            Dictionary<DocxMathBox, Dictionary<DocxMathBookmark, DocxMathBookmarkPoint>>? mathBookmarkIndexes = null;
            const double viewportContextPoints = 3d;
            for (int pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxLayoutPage page = layout.Pages[pageIndex];
                var pageNumber = page.GetPageNumber(pageIndex + 1);
                DocxMarkupContext pageMarkupContext = WithPageTextEmissionOffsets(markupContext, layout, pageIndex, separatePlainHeaderOrigins);
                double pageTextEmissionBaselineOffset = ResolveTextEmissionBaselineOffset(pageMarkupContext);
                double pageTextEmissionXOffset = ResolveTextEmissionXOffset(pageMarkupContext);
                var bodyMath = fontResources.MathEquations.Count == 0 ? null : page.Items.OfType<DocxEquationLayout>()
                    .ToLookup(item => item.Math.Equation, ReferenceEqualityComparer.Instance);
                var pageDrawings = drawingPages.PageAll(pageIndex);
                foreach (var source in EnumerateRenderedPageLineSources(page, pageDrawings, pageMarkupContext, page.Height))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    DocxTextLineLayout line = source.Line;
                    if (line.SourceParagraph is not { } paragraph ||
                        paragraph.BookmarkAnchors.Count == 0)
                    {
                        continue;
                    }

                    if (line.Math is { } math && !line.IsInlineMath && line.FrameOriginOffsetX is { } frameOffset)
                    {
                        AddMathTargets(paragraph, math, line.X, line.BaselineY, line.MathScale, line.X + frameOffset,
                            line.BaselineY + math.BaselineInset * line.MathScale, pageIndex, pageTextEmissionXOffset, pageTextEmissionBaselineOffset);
                        continue;
                    }

                    IReadOnlyList<DocxTextEmissionSegment> segments = CreateTextEmissionSegments(line, fontResources, pageNumber, layout.Pages.Count, textEmissionFontScale, pageTextEmissionBaselineOffset, pageTextEmissionXOffset, useWordCompatibleTextProfile, cancellationToken)
                        .Where(segment => !segment.IsTerminalLineSpace && segment.SourceTextRunIndex >= 0 && segment.Width > 0d)
                        .ToArray();
                    foreach (DocxBookmarkAnchor bookmark in paragraph.BookmarkAnchors)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (string.IsNullOrEmpty(bookmark.Name) || destinations.ContainsKey(bookmark.Name))
                        {
                            continue;
                        }

                        DocxInlineItem? targetItem = ResolveBookmarkInlineItem(paragraph, bookmark, cancellationToken,
                            out int itemIndex, out bool afterItem);
                        bool tab = targetItem is DocxInlineTab;
                        if (line.InlineFrame is { } inlineFrame &&
                            (tab || targetItem is DocxInlinePicture or DocxInlineBox || afterItem && targetItem is DocxInlineEquation))
                        {
                            // The measured line owns both boundaries, including tabs with no ink.
                            foreach (var extent in line.InlineSourceExtents)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                if (extent.ItemIndex != itemIndex || extent.EndOffset < extent.StartOffset ||
                                    !tab && extent.EndOffset == extent.StartOffset ||
                                    afterItem && extent.ContinuesSourceItem) continue;
                                double offset = afterItem ? extent.EndOffset :
                                    extent.StartOffset == 0d ? line.FrameOriginOffsetX ?? 0d : extent.StartOffset;
                                destinations.Add(bookmark.Name, new PdfLinkDestination(pageIndex,
                                    line.X + offset + pageTextEmissionXOffset - viewportContextPoints * textEmissionFontScale,
                                    line.BaselineY + inlineFrame.BaselineInset - pageTextEmissionBaselineOffset, Zoom: null));
                                break;
                            }
                            continue;
                        }

                        if (targetItem is DocxInlinePicture targetPicture)
                        {
                            // Legacy text wrapping still owns ordinary placement.
                            // Borrow its actual picture, including container transforms,
                            // rather than substituting a following prose segment.
                            foreach (var (image, _) in EnumeratePlacedPictureFrames(page, pageNumber.Physical,
                                EnumerateStaticLayoutItems(page, pageMarkupContext), pageMarkupContext,
                                sourceImage => ReferenceEquals(sourceImage, targetPicture.Image), cancellationToken))
                            {
                                destinations.Add(bookmark.Name, new PdfLinkDestination(pageIndex,
                                    image.X + (afterItem ? image.Width : 0d) - viewportContextPoints * textEmissionFontScale,
                                    image.Y + image.Height, Zoom: null));
                                break;
                            }
                            continue;
                        }

                        if (targetItem is DocxInlineEquation sourceEquation)
                        {
                            // Legacy body lines paint a separate item. Resolve its final
                            // placed record, rather than recomputing its geometry from prose.
                            var bodyEquation = bodyMath?[sourceEquation.Equation].FirstOrDefault(item =>
                                item.SourceBlockIndex == line.SourceBlockIndex);
                            if (line.InlineFrame is not null || line.IsInlineMath || bodyEquation is not null)
                            {
                                // A source-bound formula must wait for its own placed
                                // line, including a line with no prose segments.
                                var placement = line.InlineMath.FirstOrDefault(value =>
                                    ReferenceEquals(value.Math.Equation, sourceEquation.Equation) &&
                                    (bookmark.MathSource is null || TryGetMathBookmarkPoint(value.Math.Box, bookmark.MathSource, out _)));
                                // The qualified single-prose path retains its breaking
                                // fallback but publishes the same measured math position.
                                if (placement is null && line.IsInlineMath && line.Math is { } singleMath &&
                                    ReferenceEquals(singleMath.Equation, sourceEquation.Equation))
                                    placement = new(singleMath, line.InlineMathOffsetX);
                                if (placement is null && bodyEquation is not null)
                                    placement = new(bodyEquation.Math, bodyEquation.X - line.X);
                                if (bookmark.MathSource is { } mathSource)
                                {
                                    if (placement is not null && TryGetMathBookmarkPoint(placement.Math.Box, mathSource, out var point))
                                    {
                                        double scale = bodyEquation?.Scale ?? line.MathScale;
                                        destinations.Add(bookmark.Name, new PdfLinkDestination(pageIndex,
                                            (bodyEquation?.X ?? line.X + placement.OffsetX) + point.X * scale +
                                                pageTextEmissionXOffset - viewportContextPoints * textEmissionFontScale,
                                            (bodyEquation?.BaselineY ?? line.BaselineY) + point.Top * scale -
                                                pageTextEmissionBaselineOffset, Zoom: null));
                                    }
                                    continue;
                                }
                                double mathTop = bodyEquation is not null
                                    ? bodyEquation.BaselineY + bodyEquation.Math.BaselineInset * bodyEquation.Scale
                                    : line.BaselineY + (line.InlineFrame?.BaselineInset ?? (placement?.Math.BaselineInset ?? 0d) * line.MathScale);
                                // Word keeps line-start viewport context at the container
                                // origin; a later formula targets its actual placed position.
                                if (placement is not null && line.FrameOriginOffsetX is { } inlineOrigin)
                                {
                                    double offset = afterItem
                                        ? placement.OffsetX + placement.Math.Box.AdvanceWidth * (bodyEquation?.Scale ?? line.MathScale)
                                        : placement.OffsetX == 0d ? inlineOrigin : placement.OffsetX;
                                    destinations.Add(bookmark.Name, new PdfLinkDestination(pageIndex,
                                        line.X + offset + pageTextEmissionXOffset - viewportContextPoints * textEmissionFontScale,
                                        mathTop - pageTextEmissionBaselineOffset, Zoom: null));
                                }
                                continue;
                            }
                        }

                        if (afterItem && !tab && targetItem is not DocxInlineText) continue;

                        if (!TryResolveBookmarkDestinationSegment(segments, paragraph, bookmark,
                            afterItem ? targetItem as DocxInlineText : null, cancellationToken,
                            out DocxTextEmissionSegment? target, out double targetX))
                        {
                            continue;
                        }

                        // Word 16 bookmark controls: indentation/alignment do not move the
                        // viewport origin, while a prefix before a mid-line bookmark does.
                        // The three design-point context scales with the review print scale.
                        if (source.Owner.Kind == DocxStoryKind.Body && line.FrameOriginOffsetX is { } columnOffsetX)
                        {
                            targetX += columnOffsetX - viewportContextPoints * textEmissionFontScale;
                        }

                        double anchorInset;
                        if (DocxInlineLineGeometry.IsNonPaintingSourceCarrier(line) &&
                            line.FirstLineInsetPoints is { } sourceInset && double.IsFinite(sourceInset))
                        {
                            // Source-only marks keep their placed insertion
                            // frame; its relative inset already uses printed points.
                            anchorInset = sourceInset;
                        }
                        else if (target.FallbackFace is not null)
                        {
                            anchorInset = target.FontSize * PdfFallbackFont.AscentEm;
                        }
                        else if (target.Resource is { } bookmarkResource)
                        {
                            OpenTypeFont bookmarkFont = bookmarkResource.Embedded.Font;
                            anchorInset = line.UsesUniformBodyTypographicBaseline && (bookmarkFont.Os2.SelectionFlags & 0x80) != 0 &&
                                Math.Abs(target.FontSize - line.FontSize * textEmissionFontScale) < 0.000001d
                                ? DocxLineMetrics.MeasureTypographicBaselineInset(bookmarkFont, target.FontSize)
                                : bookmarkFont.Os2.WindowsAscender * target.FontSize / bookmarkFont.UnitsPerEm;
                        }
                        else
                        {
                            continue;
                        }
                        destinations[bookmark.Name] = new PdfLinkDestination(
                            pageIndex,
                            targetX,
                            target.BaselineY + anchorInset,
                            Zoom: null);
                    }
                }
                // Whole body equations have their own measured layout records;
                // they do not manufacture prose segments for navigation.
                if (fontResources.MathEquations.Count != 0)
                foreach (var equation in page.Items.OfType<DocxEquationLayout>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (equation.FrameOriginOffsetX is { } frameOffset)
                        AddMathTargets(equation.SourceParagraph, equation.Math, equation.X, equation.BaselineY, equation.Scale, equation.X + frameOffset,
                            equation.BaselineY + equation.Math.BaselineInset * equation.Scale,
                            pageIndex, pageTextEmissionXOffset, pageTextEmissionBaselineOffset);
                }

                AddPlacedTableBookmarkDestinations(page, pageDrawings, pageIndex, pageMarkupContext, destinations, cancellationToken);
            }

            return destinations;

            void AddMathTargets(DocxParagraph paragraph, DocxPreparedMathParagraph math, double mathX, double baseline,
                double scale, double originX, double top, int pageIndex, double xOffset, double yOffset)
            {
                foreach (var bookmark in paragraph.BookmarkAnchors)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!string.IsNullOrEmpty(bookmark.Name))
                    {
                        if (bookmark.MathSource is { } source)
                        {
                            if (TryGetMathBookmarkPoint(math.Box, source, out var point))
                                destinations.TryAdd(bookmark.Name, new PdfLinkDestination(pageIndex,
                                    mathX + point.X * scale + xOffset - viewportContextPoints * textEmissionFontScale,
                                    baseline + point.Top * scale - yOffset, Zoom: null));
                            continue;
                        }
                        destinations.TryAdd(bookmark.Name, new PdfLinkDestination(pageIndex,
                            originX + xOffset - viewportContextPoints * textEmissionFontScale,
                            top - yOffset, Zoom: null));
                    }
                }
            }

            bool TryGetMathBookmarkPoint(DocxMathBox box, DocxMathBookmark source, out DocxMathBookmarkPoint point)
            {
                if (box.BookmarkPoints.Count == 0) { point = default; return false; }
                mathBookmarkIndexes ??= new(ReferenceEqualityComparer.Instance);
                if (!mathBookmarkIndexes.TryGetValue(box, out var points))
                {
                    points = new(ReferenceEqualityComparer.Instance);
                    foreach (var value in box.BookmarkPoints)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        points.TryAdd(value.Source, value);
                    }
                    mathBookmarkIndexes.Add(box, points);
                }
                return points.TryGetValue(source, out point);
            }
        }

        IReadOnlyList<PdfLinkAnnotation> CreateHyperlinkAnnotations(DocxLayoutPage page, int pageIndex, DocxPageNumber pageNumber, int pageCount,
            IReadOnlySet<DocxInlineImage> paintedPictures, IReadOnlyList<DocxLayoutItem> staticItems,
            IReadOnlyList<PlacedNoteTableFrame?>? noteTables, IReadOnlyList<PdfLinkAnnotation> drawingUris)
        {
            var annotations = new List<PdfLinkAnnotation>();
            Dictionary<DocxInlineImage, DocxInlineImageLayout?>? linkedPictures = null;
            // Ordinary body paragraphs carry reference-qualified printed line boxes.
            // Other stories retain the earlier glyph-relative geometry until qualified.
            const double BodyHyperlinkRectHorizontalPadPoints = 2.25d;
            // RV06: measured legacy hyperlink rectangle geometry across 11/12pt:
            // horizontal pads about 0.20em per side, line tops at baseline plus 0.939em,
            // non-last bottoms tiling the next page line top, default last-line bottoms at
            // baseline minus 1.136em. (Office renders style-less runs at 12pt, matching the
            // reader default; single-spaced and exact-height page-last bottoms of 0.947em
            // and 1.271em show line-height-dependent bottoms that stay open.)
            const double HyperlinkRectHorizontalPadEm = 2.2 / 11.0;
            const double HyperlinkRectTopEm = 10.33 / 11.0;
            const double HyperlinkRectBottomEm = 12.5 / 11.0;
            List<(DocxHyperlinkSpan Link, double MinX, double MaxX, PdfRectangle? Clip)>? previousLineLinks = null;
            double previousLineTop = 0d;
            double previousLineBottomFallback = 0d;
            double previousPaintedInlineBottom = double.PositiveInfinity;
            double? previousDelta = null;
            bool previousUsesBodyLineBox = false;
            var paper = new PdfRectangle(0d, 0d, page.Width, page.Height);
            foreach (var (line, _, paintViewport) in EnumerateRenderedPageLineSources(page,
                drawingPages.PageAll(pageIndex), markupContext, page.Height, staticItems, noteTables, cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();
                LinkPaintViewport? clip = paintViewport is { } source
                    ? source with { Bounds = IntersectPaintViewports(paper, source.Bounds) } : null;
                // Break-spill glyph rows do not consume a paragraph line slot.
                // They must not shorten the preceding body's clickable line box.
                if (previousUsesBodyLineBox && line.Segments.Count > 0 &&
                    line.Segments.All(static segment => segment.Role == DocxTextSegmentRole.BreakSpill))
                {
                    continue;
                }
                // RV06: Office tiles link rectangles down to the next page line top across all
                // lines, so a held line flushes here even when the current line has no links.
                if (previousLineLinks is not null && previousDelta is not null)
                {
                    double flushNextTop = previousUsesBodyLineBox && line.BodyLineBoxBaselineInsetPoints is double nextInset
                        ? line.BaselineY - textEmissionBaselineOffset + nextInset
                        // Line coordinates are already placed; legacy font
                        // sizes still need the same print scale as emitted ink.
                        : line.BaselineY + previousDelta.Value + HyperlinkRectTopEm * line.FontSize * textEmissionFontScale;
                    // RV06: tile only into lines below the held line; side-by-side table-cell
                    // lines share one baseline, so they keep the full-height fallback instead.
                    // A table or other story's glyph top is not the body slot's
                    // next line-box boundary. Preserve the qualified body slot.
                    double flushBottom = previousUsesBodyLineBox && line.BodyLineBoxBaselineInsetPoints is null
                        ? previousLineBottomFallback
                        : flushNextTop < previousLineTop
                            ? Math.Max(flushNextTop, previousLineBottomFallback)
                            : previousLineBottomFallback;
                    EmitLineLinkRects(previousLineLinks, previousLineTop, Math.Min(flushBottom, previousPaintedInlineBottom));
                    previousLineLinks = null;
                }

                if (line.SourceParagraph is not { } paragraph || paragraph.Hyperlinks.Count == 0)
                {
                    EmitLineMathLinks(line, clip);
                    continue;
                }
                bool usesBodyLineBox = line.BodyLineBoxBaselineInsetPoints is not null && line.BodyLineBoxHeightPoints is not null;
                double lineLinkPad = usesBodyLineBox
                    ? BodyHyperlinkRectHorizontalPadPoints * textEmissionFontScale
                    : HyperlinkRectHorizontalPadEm * line.FontSize;
                var mergedLinkRects = new List<(DocxHyperlinkSpan Link, double MinX, double MaxX, double Top, double BaseBl, double BaseFs, PdfRectangle? Clip)>();
                double paintedInlineBottom = double.PositiveInfinity;
                double mathLineBottom = 0d, mathLineTop = 0d;
                Action<DocxHyperlinkSpan, double, double, double, PdfRectangle?>? rootMathLink =
                    DocxInlineLineGeometry.TryGetOccupiedBounds(line, out mathLineBottom, out mathLineTop)
                        ? AddRootMathFragment : null;

                // Root-row links share the placed line and its source target.
                // Nested argument links keep their independently measured paint.
                EmitLineMathLinks(line, clip, rootMathLink);

                IReadOnlyList<DocxHyperlinkSpan> links = paragraph.Hyperlinks;
                // RV06: Office emits one padded link rectangle per hyperlink per line and
                // tiles them top-anchored down to the next page line top.
                foreach (DocxTextEmissionSegment segment in CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount, textEmissionFontScale, textEmissionBaselineOffset, textEmissionXOffset, useWordCompatibleTextProfile, cancellationToken))
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

                    // Fragments without a measurable font resource cannot anchor a rectangle.
                    // RV01: fallback faces stay measurable through diagnosed constants downstream.
                    if (segment.FallbackFace is null && segment.Resource is null)
                    {
                        continue;
                    }

                    double annotationWidth = ResolveHyperlinkAnnotationWidth(segment, useWordCompatibleTextProfile);
                    double fragmentMinX = segment.X;
                    double fragmentMaxX = segment.X + annotationWidth;
                    PdfRectangle? fragmentClip = null;
                    if (clip is not null)
                    {
                        var font = segment.Resource?.Embedded.Font;
                        double ascent = font is null ? segment.FontSize * PdfFallbackFont.AscentEm :
                            DocxLineMetrics.MeasureWindowsAscender(font, segment.FontSize);
                        double descent = font is null ? segment.FontSize * PdfFallbackFont.DescentEm :
                            DocxLineMetrics.MeasureWindowsDescender(font, segment.FontSize);
                        if (!TryResolveLinkClip(clip, fragmentMinX, fragmentMaxX,
                            segment.BaselineY - descent, segment.BaselineY + ascent, out fragmentClip)) continue;
                        if (font is not null && fragmentClip is { } viewport &&
                            (viewport.Y > segment.BaselineY - descent ||
                             viewport.Y + viewport.Height < segment.BaselineY + ascent) &&
                            !IntersectsVerticalLinkInk(segment, font, viewport, cancellationToken)) continue;
                    }
                    double fragmentTop = usesBodyLineBox
                        ? line.BaselineY - textEmissionBaselineOffset + line.BodyLineBoxBaselineInsetPoints!.Value
                        : segment.BaselineY + HyperlinkRectTopEm * segment.FontSize;
                    double fragmentBl = usesBodyLineBox ? line.BaselineY - textEmissionBaselineOffset : segment.BaselineY;
                    double fragmentFs = segment.FontSize;
                    AddLinkFragment(link, fragmentMinX, fragmentMaxX, fragmentTop, fragmentBl, fragmentFs, fragmentClip);
                }
                // Ordered source ranges include mathematical fragments without
                // inventing a text run for them. Use their final placed bounds.
                foreach (var extent in line.InlineSourceExtents)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!extent.IsMath || extent.EndOffset <= extent.StartOffset) continue;
                    var link = FindInlineLink(paragraph, extent.ItemIndex);
                    if (link is null) continue;
                    double baseline = line.BaselineY - textEmissionBaselineOffset;
                    double left = line.X + extent.StartOffset + textEmissionXOffset;
                    double right = line.X + extent.EndOffset + textEmissionXOffset;
                    if (!TryResolveLinkClip(clip, left, right, baseline - extent.BelowBaseline,
                        baseline + extent.AboveBaseline, out var fragmentClip)) continue;
                    // A text line slot can be shallower than a fraction's final
                    // painted extent, especially after printed table placement.
                    paintedInlineBottom = Math.Min(paintedInlineBottom, baseline - extent.BelowBaseline);
                    AddLinkFragment(link, left, right,
                        baseline + Math.Max(extent.AboveBaseline, line.InlineFrame?.BaselineInset ?? 0d),
                        baseline, line.FontSize * textEmissionFontScale, fragmentClip);
                }

                // A picture extends an existing text/math link; an image-only
                // source span retains its current fallback. Drawing URI targets
                // own their separate annotations and are not absorbed here.
                foreach (var extent in line.InlineSourceExtents)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (mergedLinkRects.Count == 0 || paragraph.InlineContent is not { } items ||
                        (uint)extent.ItemIndex >= (uint)items.Count || items[extent.ItemIndex] is not DocxInlinePicture picture ||
                        picture.Image.ClickTarget is not null || FindInlineLink(paragraph, extent.ItemIndex) is not { } link)
                        continue;
                    int index = mergedLinkRects.FindIndex(item => ReferenceEquals(item.Link, link));
                    if (index < 0) continue;
                    linkedPictures ??= ReadLinkedPictureFrames();
                    if (!linkedPictures.TryGetValue(picture.Image, out var placed) || placed is null ||
                        placed.SourceLineIndex is null || placed.SourceLineIndex != line.SourceLineIndex) continue;
                    if (!TryResolveLinkClip(clip, placed.X, placed.X + placed.Width, placed.Y,
                        placed.Y + placed.Height, out var fragmentClip)) continue;
                    var merged = mergedLinkRects[index];
                    mergedLinkRects[index] = (link, Math.Min(merged.MinX, placed.X), Math.Max(merged.MaxX, placed.X + placed.Width),
                        Math.Max(merged.Top, placed.Y + placed.Height), merged.BaseBl, merged.BaseFs, merged.Clip ?? fragmentClip);
                    paintedInlineBottom = Math.Min(paintedInlineBottom, placed.Y);
                }

                void AddRootMathFragment(DocxHyperlinkSpan link, double left, double right, double scale, PdfRectangle? fragmentClip)
                {
                    // Keep the math path's qualified horizontal pad when the
                    // shared line merger applies its own pad at publication.
                    double padAdjustment = BodyHyperlinkRectHorizontalPadPoints * scale - lineLinkPad;
                    AddLinkFragment(link, left - padAdjustment, right + padAdjustment,
                        mathLineTop - textEmissionBaselineOffset, line.BaselineY - textEmissionBaselineOffset,
                        line.FontSize * textEmissionFontScale, fragmentClip);
                }

                void AddLinkFragment(DocxHyperlinkSpan link, double fragmentMinX, double fragmentMaxX,
                    double fragmentTop, double fragmentBl, double fragmentFs, PdfRectangle? fragmentClip)
                {
                    bool linkTargetEmittable = IsExternalHyperlink(link) ||
                        (!string.IsNullOrEmpty(link.Anchor) && bookmarkDestinations.ContainsKey(link.Anchor));
                    if (!linkTargetEmittable)
                    {
                        return;
                    }
                    bool fragmentAbsorbed = false;
                    for (int mergedIndex = 0; mergedIndex < mergedLinkRects.Count; mergedIndex++)
                    {
                        if (!ReferenceEquals(mergedLinkRects[mergedIndex].Link, link))
                        {
                            continue;
                        }
                        var absorbed = mergedLinkRects[mergedIndex];
                        mergedLinkRects[mergedIndex] = (
                            link,
                            Math.Min(absorbed.MinX, fragmentMinX),
                            Math.Max(absorbed.MaxX, fragmentMaxX),
                            Math.Max(absorbed.Top, fragmentTop),
                            fragmentTop > absorbed.Top ? fragmentBl : absorbed.BaseBl,
                            fragmentTop > absorbed.Top ? fragmentFs : absorbed.BaseFs,
                            absorbed.Clip ?? fragmentClip);
                        fragmentAbsorbed = true;
                        break;
                    }
                    if (!fragmentAbsorbed)
                    {
                        mergedLinkRects.Add((link, fragmentMinX, fragmentMaxX, fragmentTop, fragmentBl, fragmentFs, fragmentClip));
                    }
                }
                // Root math selections merge horizontally by source identity,
                // but their occupied frame must not change another link's
                // independently placed text-slot or argument rectangle.
                foreach (var root in mergedLinkRects)
                    if (root.Link.IsMathRootContent)
                        EmitLinkRect(root.Link, root.MinX - lineLinkPad, root.MaxX + lineLinkPad,
                            mathLineTop - textEmissionBaselineOffset,
                            mathLineBottom - textEmissionBaselineOffset - (line.ParagraphAfterSpacing ?? line.PendingAfterSpacing ?? 0d), root.Clip);
                mergedLinkRects.RemoveAll(static item => item.Link.IsMathRootContent);

                if (mergedLinkRects.Count > 0)
                {
                    double heldTop = double.NegativeInfinity;
                    double heldBl = 0d;
                    double heldFs = 0d;
                    foreach (var merged in mergedLinkRects)
                    {
                        if (merged.Top > heldTop)
                        {
                            heldTop = merged.Top;
                            heldBl = merged.BaseBl;
                            heldFs = merged.BaseFs;
                        }
                    }

                    previousLineLinks = new(mergedLinkRects.Count);
                    foreach (var merged in mergedLinkRects)
                    {
                        previousLineLinks.Add((merged.Link, merged.MinX - lineLinkPad, merged.MaxX + lineLinkPad, merged.Clip));
                    }

                    previousLineTop = heldTop;
                    previousPaintedInlineBottom = paintedInlineBottom;
                    // RV06: a page-last rectangle covers its line slot plus trailing paragraph space.
                    double heldSlotHeight = line.BodyLineBoxHeightPoints ?? line.LineHeight ?? ((HyperlinkRectTopEm + HyperlinkRectBottomEm) * heldFs);
                    double heldAfterSpacing = line.ParagraphAfterSpacing ?? line.PendingAfterSpacing ?? 0d;
                    previousLineBottomFallback = heldTop - heldSlotHeight - heldAfterSpacing;
                    previousDelta = heldBl - line.BaselineY;
                    previousUsesBodyLineBox = usesBodyLineBox;
                }
            }

            if (previousLineLinks is not null)
            {
                EmitLineLinkRects(previousLineLinks, previousLineTop, Math.Min(previousLineBottomFallback, previousPaintedInlineBottom));
            }

            foreach (var equation in page.Items.OfType<DocxEquationLayout>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                Action<DocxHyperlinkSpan, double, double, double, PdfRectangle?>? rootMathLink = null;
                if (DocxInlineLineGeometry.TryGetPreparedMathBounds(equation.Math, equation.BaselineY, equation.Scale,
                    out double bottom, out double top))
                    rootMathLink = (link, left, right, scale, clip) => EmitLinkRect(link,
                        left - BodyHyperlinkRectHorizontalPadPoints * scale, right + BodyHyperlinkRectHorizontalPadPoints * scale,
                        top - textEmissionBaselineOffset,
                        bottom - textEmissionBaselineOffset - equation.SourceParagraph.SpacingAfterPoints * scale, clip);
                EmitMathLinks(equation.Math, equation.X, equation.BaselineY, equation.Scale, rootMathLink);
            }
            annotations.AddRange(CreateInlinePictureAnnotations(page, pageNumber.Physical, staticItems, paintedPictures, markupContext, cancellationToken));
            annotations.AddRange(drawingUris);
            annotations.AddRange(CreateNoteReferenceAnnotations(page, pageNumber, pageCount, fontResources,
                markupContext, noteDestinations, noteSpacingParagraphs, cancellationToken));
            return annotations;

        void EmitLineMathLinks(DocxTextLineLayout line, LinkPaintViewport? clip,
            Action<DocxHyperlinkSpan, double, double, double, PdfRectangle?>? rootMathLink = null)
        {
            if (line.Math is { } math && !line.IsInlineMath)
                EmitMathLinks(math, line.X, line.BaselineY, line.MathScale, rootMathLink, clip);
            else
            {
                if (line.IsInlineMath && line.Math is { } inlineMath)
                    EmitMathLinks(inlineMath, line.X + line.InlineMathOffsetX, line.BaselineY, line.MathScale, rootMathLink, clip);
                foreach (var placement in line.InlineMath)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    EmitMathLinks(placement.Math, line.X + placement.OffsetX, line.BaselineY, line.MathScale, rootMathLink, clip);
                }
            }
        }

        void EmitMathLinks(DocxPreparedMathParagraph math, double x, double baseline, double scale,
            Action<DocxHyperlinkSpan, double, double, double, PdfRectangle?>? rootMathLink = null, LinkPaintViewport? clip = null)
        {
            foreach (var (link, rectangle) in ReadMathHyperlinkBounds(math.Box, cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();
                double left = x + rectangle.X * scale + textEmissionXOffset;
                double right = x + (rectangle.X + rectangle.Width) * scale + textEmissionXOffset;
                double bottom = baseline + rectangle.Y * scale - textEmissionBaselineOffset;
                double top = baseline + (rectangle.Y + rectangle.Height) * scale - textEmissionBaselineOffset;
                if (!TryResolveLinkClip(clip, left, right, bottom, top, out var fragmentClip)) continue;
                if (link.IsMathRootContent && rootMathLink is not null)
                {
                    rootMathLink(link, left, right, scale, fragmentClip);
                    continue;
                }
                double pad = BodyHyperlinkRectHorizontalPadPoints * scale;
                EmitLinkRect(link, left - pad, right + pad, top, bottom, fragmentClip);
            }
        }

        Dictionary<DocxInlineImage, DocxInlineImageLayout?> ReadLinkedPictureFrames()
        {
            var requested = new HashSet<DocxInlineImage>(ReferenceEqualityComparer.Instance);
            foreach (var line in EnumerateRenderedPageTextLines(drawingPages, page, pageIndex, markupContext, page.Height))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (line.SourceParagraph is not { InlineContent: { } items } paragraph || paragraph.Hyperlinks.Count == 0) continue;
                foreach (var extent in line.InlineSourceExtents)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if ((uint)extent.ItemIndex < (uint)items.Count && items[extent.ItemIndex] is DocxInlinePicture picture &&
                        picture.Image.ClickTarget is null && FindInlineLink(paragraph, extent.ItemIndex) is not null)
                        requested.Add(picture.Image);
                }
            }
            // These page-local identity sets fit the existing coarse construction
            // reservation per placed item. Repeated placements are ambiguous and
            // retain glyph-only bounds rather than selecting an arbitrary frame.
            var frames = new Dictionary<DocxInlineImage, DocxInlineImageLayout?>(ReferenceEqualityComparer.Instance);
            foreach (var image in EnumerateLinkedPictureFrames(page, pageNumber.Physical, staticItems, paintedPictures,
                markupContext, requested.Contains, cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!frames.TryAdd(image.Image, image)) frames[image.Image] = null;
            }
            return frames;
        }

        static DocxHyperlinkSpan? FindInlineLink(DocxParagraph paragraph, int itemIndex) =>
            paragraph.Hyperlinks.FirstOrDefault(link => !link.IsMathContent && link.StartInlineItemIndex is int start &&
                link.EndInlineItemIndex is int end && itemIndex >= start && itemIndex < end);

        void EmitLineLinkRects(
            List<(DocxHyperlinkSpan Link, double MinX, double MaxX, PdfRectangle? Clip)> emissions,
            double top,
            double bottom)
        {
            if (top - bottom <= 0d)
            {
                return;
            }

            foreach (var emission in emissions)
            {
                EmitLinkRect(emission.Link, emission.MinX, emission.MaxX, top, bottom, emission.Clip);
            }
        }

        void EmitLinkRect(DocxHyperlinkSpan link, double left, double right, double top, double bottom, PdfRectangle? clip = null)
        {
            if (clip is { } viewport)
            {
                left = Math.Max(left, viewport.X);
                right = Math.Min(right, viewport.X + viewport.Width);
                bottom = Math.Max(bottom, viewport.Y);
                top = Math.Min(top, viewport.Y + viewport.Height);
            }
            if (!(right > left) || !(top > bottom)) return;
            if (IsExternalHyperlink(link))
                annotations.Add(PdfLinkAnnotation.ToUri(left, bottom, right - left, top - bottom, link.Target ?? string.Empty));
            else if (!string.IsNullOrEmpty(link.Anchor) &&
                bookmarkDestinations.TryGetValue(link.Anchor, out PdfLinkDestination mergedDestination))
                annotations.Add(PdfLinkAnnotation.ToDestination(left, bottom, right - left, top - bottom, mergedDestination));
        }

        static bool TryResolveLinkClip(LinkPaintViewport? viewport, double left, double right, double bottom, double top,
            out PdfRectangle? clip)
        {
            clip = null;
            if (viewport is not { } paint) return true;
            var bounds = paint.Bounds;
            double boundsRight = bounds.X + bounds.Width, boundsTop = bounds.Y + bounds.Height;
            if (!(Math.Max(left, bounds.X) < Math.Min(right, boundsRight) &&
                  Math.Max(bottom, bounds.Y) < Math.Min(top, boundsTop))) return false;

            // A fully visible cell fragment retains its qualified padding.
            if (!paint.PreserveFittingPadding || left < bounds.X || right > boundsRight || bottom < bounds.Y || top > boundsTop)
                clip = bounds;
            return true;
        }

        static bool IntersectsVerticalLinkInk(DocxTextEmissionSegment segment, OpenTypeFont font,
            PdfRectangle viewport, CancellationToken token)
        {
            // The nominal font box includes blank space above and below its
            // glyphs. Keep the existing link padding, but require some ink
            // when the viewport cuts into that box vertically.
            double scale = DocxTextEmissionPlanner.CreateForEmissionSegment(segment).PdfFontSize / font.UnitsPerEm;
            double bottom = viewport.Y;
            double top = bottom + viewport.Height;
            foreach (Rune rune in segment.Text.EnumerateRunes())
            {
                token.ThrowIfCancellationRequested();
                if (Rune.IsWhiteSpace(rune)) continue;
                ushort glyph = font.MapCodePoint(rune.Value);
                if (glyph == 0) continue;
                if (!font.TryGetGlyphBounds(glyph, out var ink)) return true;
                if (ink.YMax > ink.YMin && segment.BaselineY + ink.YMax * scale > bottom &&
                    segment.BaselineY + ink.YMin * scale < top)
                    return true;
            }
            return false;
        }

        bool IsHyperlinkSegment(DocxHyperlinkSpan link, int sourceTextRunIndex)
        {
            return !link.IsMathContent && sourceTextRunIndex >= link.SourceRunStartIndex &&
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
            segment.IsTerminalLineSpace,
            segment.FontScale);
        // RV01: fallback runs have no advance profile; measured width stands in.
        double emittedAdvance = segment.Resource is { } advanceResource
            ? DocxTextEmissionPlanner.MeasureAdvanceProfile(segment.Text, advanceResource.Embedded, segment.Width, plan).PlannedEmittedAdvance
            : segment.Width;
        return emittedAdvance > 0d ? emittedAdvance : segment.Width;
    }

    private static DocxInlineItem? ResolveBookmarkInlineItem(DocxParagraph paragraph, DocxBookmarkAnchor bookmark,
        CancellationToken token, out int index, out bool afterItem)
    {
        index = bookmark.InlineItemIndex ?? -1;
        afterItem = false;
        if (paragraph.InlineContent is not { } items || index < 0 || index > items.Count)
            return null;

        while (index < items.Count && IsNonpaintingBookmarkItem(items[index]))
        {
            token.ThrowIfCancellationRequested();
            index++;
        }
        if (index < items.Count && (bookmark.MathSource is not null ||
            !IsDiscardableBookmarkSuffix(items[index], token)))
        {
            return items[index];
        }
        if (bookmark.MathSource is not null) return null;

        // Terminal breakable spaces and unused break markers have no ink.
        // Resolve their boundary on the preceding placed item.
        int last = items.Count - 1;
        while (last >= 0 && (IsNonpaintingBookmarkItem(items[last]) || IsDiscardableBookmarkSuffix(items[last], token)))
        {
            token.ThrowIfCancellationRequested();
            last--;
        }
        if (index <= last) return items[index];
        index = last;
        if (index >= 0 && items[index] is DocxInlineText or DocxInlinePicture or DocxInlineBox or DocxInlineEquation or DocxInlineTab)
        {
            afterItem = true;
            return items[index];
        }
        return null;
    }

    private static bool IsNonpaintingBookmarkItem(DocxInlineItem item) => item is
        DocxInlineText { Run.EffectiveProperties.Hidden: true } or DocxInlineText { Length: 0 } or
        DocxInlineTab { Run.EffectiveProperties.Hidden: true } or DocxInlineBreak { Run.EffectiveProperties.Hidden: true };

    private static bool IsDiscardableBookmarkSuffix(DocxInlineItem item, CancellationToken token)
    {
        if (item is not DocxInlineText { Length: > 0 } text || text.Run.FieldKind is { } field && field.IsDynamic())
            return false;
        for (int index = text.Start; index < text.Start + text.Length; index++)
        {
            token.ThrowIfCancellationRequested();
            if (!IsBookmarkBreakMarker(text.Run, index) &&
                !DocxTextBreakRules.IsBreakableWhitespaceChar(text.Run.Text[index])) return false;
        }
        return true;
    }

    private static bool IsBookmarkBreakMarker(DocxTextRun run, int index) =>
        DocxTextBreakRules.IsOptionalBreakMarker(run.Text[index]) &&
        run.LiteralNoteMarkerOffset != run.SourceTextOffsetInRun + index;

    private static bool TryResolveBookmarkDestinationSegment(
        IReadOnlyList<DocxTextEmissionSegment> segments,
        DocxParagraph paragraph,
        DocxBookmarkAnchor bookmark,
        DocxInlineText? terminalInlineText,
        CancellationToken token,
        [NotNullWhen(true)] out DocxTextEmissionSegment? target,
        out double targetX)
    {
        target = null;
        targetX = 0d;
        int sourceRunIndex = bookmark.SourceRunIndex;
        int sourceOffset = Math.Max(0, bookmark.TextOffset);
        bool terminalTextBoundary = terminalInlineText is not null;
        if (terminalInlineText is { } terminalText)
        {
            sourceRunIndex = terminalText.Run.SourceRunIndex;
            sourceOffset = terminalText.Run.SourceTextOffsetInRun + terminalText.Start + terminalText.Length;
        }
        else if (paragraph.InlineContent is null && bookmark.TextRunIndex <= paragraph.Runs.Count &&
            paragraph.Images.Count == 0 && paragraph.Equations.Count == 0 && paragraph.InlineTextBoxes.Count == 0)
        {
            // Ordinary prose keeps its run-based layout. Ignore trailing
            // nonpainting runs when deciding whether the bookmark is terminal.
            for (int index = paragraph.Runs.Count - 1; index >= 0; index--)
            {
                token.ThrowIfCancellationRequested();
                var run = paragraph.Runs[index];
                if (run.EffectiveProperties.Hidden || run.Text.Length == 0) continue;
                if (index >= bookmark.TextRunIndex) break;
                terminalTextBoundary = true;
                sourceRunIndex = run.SourceRunIndex;
                sourceOffset = run.SourceTextOffsetInRun + run.Text.Length;
                break;
            }
        }
        if (TryResolveBookmarkTextBoundary(segments, sourceRunIndex, sourceOffset, out target, out targetX))
            return true;
        if (!terminalTextBoundary) return false;

        // Wrapping can omit whitespace and unused break markers. Try the
        // painted endpoint only after the exact source boundary failed.
        if (paragraph.InlineContent is { } terminalItems)
        {
            for (int index = terminalItems.Count - 1; index >= 0; index--)
            {
                token.ThrowIfCancellationRequested();
                if (IsNonpaintingBookmarkItem(terminalItems[index])) continue;
                if (terminalItems[index] is not DocxInlineText text) return false;
                int length = TrimBookmarkTextEnd(text.Run, text.Start, text.Length, token);
                if (length == 0) continue;
                return TryResolveBookmarkTextBoundary(segments, text.Run.SourceRunIndex,
                    text.Run.SourceTextOffsetInRun + text.Start + length, out target, out targetX);
            }
        }
        else
        {
            for (int index = paragraph.Runs.Count - 1; index >= 0; index--)
            {
                token.ThrowIfCancellationRequested();
                DocxTextRun run = paragraph.Runs[index];
                if (run.EffectiveProperties.Hidden) continue;
                int length = TrimBookmarkTextEnd(run, 0, run.Text.Length, token);
                if (length == 0) continue;
                return TryResolveBookmarkTextBoundary(segments, run.SourceRunIndex,
                    run.SourceTextOffsetInRun + length, out target, out targetX);
            }
        }
        return false;
    }

    private static int TrimBookmarkTextEnd(DocxTextRun run, int start, int length, CancellationToken token)
    {
        while (length > 0)
        {
            token.ThrowIfCancellationRequested();
            int index = start + length - 1;
            char value = run.Text[index];
            if (!char.IsWhiteSpace(value) && !IsBookmarkBreakMarker(run, index))
            {
                break;
            }
            length--;
        }
        return length;
    }

    private static bool TryResolveBookmarkTextBoundary(
        IReadOnlyList<DocxTextEmissionSegment> segments, int sourceRunIndex, int sourceOffset,
        [NotNullWhen(true)] out DocxTextEmissionSegment? target, out double targetX)
    {
        target = null;
        targetX = 0d;
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

        // An earlier wrapped piece of the same run cannot resolve a later offset.
        DocxTextEmissionSegment? fallback = segments.FirstOrDefault(segment => segment.SourceTextRunIndex > sourceRunIndex);
        if (fallback is null)
        {
            return false;
        }

        target = fallback;
        targetX = fallback.X;
        return true;
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
        DocxPageNumber pageNumber,
        int pageCount,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        PictureNavigation navigation,
        RgbColor? background = null)
    {
        // emission helpers observe the conversion token so a
        // cancelled conversion fails fast inside large pages, not just between them.
        switch (item)
        {
            case DocxTextLineLayout textLine:
                RenderTextLine(textLine, graphics, fontResources, markupContext, pageNumber, pageCount, cancellationToken, background);
                break;
            case DocxEquationLayout equation:
                PaintMath(equation.Math, equation.X, equation.BaselineY, equation.Scale, graphics, markupContext, cancellationToken,
                    ResolveParagraphBackground(equation, background));
                break;
            case DocxInlineImageLayout image:
                RenderInlineImage(PlaceInlineImageForPaint(image, markupContext), graphics, pageImages, diagnosticSink, cancellationToken, imageCache, ref imageIndex, fontResources, markupContext, navigation);
                break;
            case DocxTableRowLayout row:
                RenderTableRow(row, IsAdjacentTableRow(previousRow, row) ? previousRow : null, IsAdjacentTableRow(row, nextRow) ? nextRow : null, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex, navigation, background);
                break;
            case DocxInlineTextBoxLayout textBox:
                RenderInlineTextBox(textBox, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex, navigation, background: background);
                break;
        }
    }

    private static OoxPdfDocxMarkupGeometryMode ResolveEffectiveMarkupGeometryMode(DocxMarkupContext context)
    {
        // A Word review frame is needed only when visible balloons occupy the lane.
        // Hidden story comments and orphan parts keep the authored wrap width.
        return context.ExpandsMarkupMargin && context.WordCompatibleHasPrintedBalloons != false
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
}
