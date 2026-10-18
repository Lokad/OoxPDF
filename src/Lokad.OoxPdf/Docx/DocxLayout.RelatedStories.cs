using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static IReadOnlyDictionary<int, double> CreateFootnoteReserveHeightBySourceBlock(
        DocxDocument document,
        IReadOnlyList<DocxRelatedStoryLayout> relatedStoryLayouts,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? separatorMeasurer = null)
    {
        var reserveHeightBySourceBlock = new Dictionary<int, double>();
        if (relatedStoryLayouts.Count == 0)
        {
            return reserveHeightBySourceBlock;
        }

        Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> storyByKey = CreateRelatedStoryLookup(relatedStoryLayouts);
        DocxRelatedStoryLayout? footnoteSeparatorLayout = FindSpecialRelatedStoryLayout(relatedStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.Separator);
        for (int sourceBlockIndex = 0; sourceBlockIndex < document.BodyElements.Count; sourceBlockIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            double reserveHeight = 0d;
            var reservedKeys = new HashSet<(DocxRelatedStoryKind Kind, string Id)>();
            bool reservedFootnoteSeparator = false;
            foreach (DocxInlineReferenceLocation location in EnumerateInlineReferenceLocations(document.BodyElements, sourceBlockIndex))
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxInlineReference reference = location.Reference;
                if (reference.Kind != DocxRelatedStoryKind.Footnote || reference.Id is null || !reservedKeys.Add((reference.Kind, reference.Id)))
                {
                    continue;
                }

                if (storyByKey.TryGetValue((reference.Kind, reference.Id), out DocxRelatedStoryLayout? storyLayout))
                {
                    if (!reservedFootnoteSeparator)
                    {
                        double reserveGapPoints = FootnoteSeparatorGapPoints;
                        if (footnoteSeparatorLayout is not null)
                        {
                            (DocxTextRun? reserveRun, double reserveFontSizePoints) = FindSeparatorMarkFont(footnoteSeparatorLayout.TextLines);
                            reserveGapPoints = ResolveSeparatorGapPoints(reserveRun, reserveFontSizePoints, separatorMeasurer);
                        }

                        reserveHeight += footnoteSeparatorLayout is null
                            ? reserveGapPoints
                            : Math.Max(0d, footnoteSeparatorLayout.ContentHeight) + reserveGapPoints;
                        reservedFootnoteSeparator = true;
                    }

                    reserveHeight += Math.Max(0d, storyLayout.ContentHeight);
                }
            }

            if (reserveHeight > 0d)
            {
                reserveHeightBySourceBlock[sourceBlockIndex] = reserveHeight;
            }
        }

        return reserveHeightBySourceBlock;
    }

    private static IReadOnlyList<DocxLayoutPage> AddPlacedRelatedStories(
        DocxDocument document,
        IReadOnlyList<DocxLayoutPage> pages,
        Func<double, IReadOnlyList<DocxRelatedStoryLayout>> resolveRelatedStoryLayouts,
        CancellationToken cancellationToken,
        double printScale = 1d,
        IDocxTextMeasurer? separatorMeasurer = null, 
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null,
        ISet<(DocxRelatedStoryKind Kind, string Id)>? alreadyPlacedStoryKeys = null,
        Func<DocxLayoutPage, int, int, double>? footerKeepOut = null,
        Func<DocxLayoutPage, int, int, double>? ownedFootnoteFooterKeepOut = null)
    {
        if (pages.Count == 0)
        {
            return pages;
        }

        IReadOnlyList<DocxRelatedStoryLayout> relatedStoryLayouts = resolveRelatedStoryLayouts(ResolvePageBodyWidth(pages[0]));
        if (relatedStoryLayouts.Count == 0)
        {
            return pages;
        }

        Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> storyByKey = CreateRelatedStoryLookup(relatedStoryLayouts);
        if (storyByKey.Count == 0)
        {
            return pages;
        }

        // R12: document-derived reference locations are memoized by source block (table
        // paragraphs no longer re-walk on every fragment page), story layouts and lookups
        // are memoized by page body width, and page owners/blocks come from one index.
        Dictionary<int, List<DocxInlineReferenceLocation>> locationsByBlock = BuildInlineReferenceLocationsByBlock(document, cancellationToken);
        var storyLayoutsByWidth = new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();
        var storyLookupByWidth = new Dictionary<double, Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>>();
        var footnoteLayoutsByWidth = new Dictionary<double, IReadOnlyList<DocxRelatedStoryLayout>>();
        var footnoteLookupByWidth = new Dictionary<double, Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>>();
        RelatedStoryPageIndex referenceIndex = RelatedStoryPageIndex.Build(pages, cancellationToken);
        var outputPages = pages.ToList();
        int outputShift = 0;
        var placedStoryKeys = new HashSet<(DocxRelatedStoryKind Kind, string Id)>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutPage page = pages[pageIndex];
            double bodyWidth = ResolvePageBodyWidth(page);
            if (!storyLayoutsByWidth.TryGetValue(bodyWidth, out IReadOnlyList<DocxRelatedStoryLayout>? pageRelatedStoryLayouts))
            {
                pageRelatedStoryLayouts = pageIndex == 0 ? relatedStoryLayouts : resolveRelatedStoryLayouts(bodyWidth);
                storyLayoutsByWidth[bodyWidth] = pageRelatedStoryLayouts;
                storyLookupByWidth[bodyWidth] = CreateRelatedStoryLookup(pageRelatedStoryLayouts);
            }

            storyByKey = storyLookupByWidth[bodyWidth];
            double footnoteWidth = ResolveNoteTextWidth(bodyWidth, page.MarkupMarginReservePoints);
            IReadOnlyList<DocxRelatedStoryLayout> footnoteLayouts;
            Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> footnoteByKey;
            if (page.MarkupMarginReservePoints <= 0d)
            {
                footnoteLayouts = pageRelatedStoryLayouts;
                footnoteByKey = storyByKey;
            }
            else
            {
                if (!footnoteLayoutsByWidth.TryGetValue(footnoteWidth, out IReadOnlyList<DocxRelatedStoryLayout>? cachedFootnoteLayouts))
                {
                    cachedFootnoteLayouts = resolveRelatedStoryLayouts(footnoteWidth);
                    footnoteLayoutsByWidth[footnoteWidth] = cachedFootnoteLayouts;
                    footnoteLookupByWidth[footnoteWidth] = CreateRelatedStoryLookup(cachedFootnoteLayouts);
                }

                footnoteLayouts = cachedFootnoteLayouts;
                footnoteByKey = footnoteLookupByWidth[footnoteWidth];
            }

            DocxRelatedStoryLayout? footnoteSeparatorLayout = FindSpecialRelatedStoryLayout(footnoteLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.Separator);
            List<DocxReferencedRelatedStoryLayout> pageFootnoteStories = [];
            foreach (int sourceBlockIndex in referenceIndex.SortedBlocks(pageIndex))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!locationsByBlock.TryGetValue(sourceBlockIndex, out List<DocxInlineReferenceLocation>? blockLocations))
                {
                    continue;
                }

                foreach (DocxInlineReferenceLocation location in blockLocations)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    DocxInlineReference reference = location.Reference;
                    if (reference.Kind != DocxRelatedStoryKind.Footnote || reference.Id is null)
                    {
                        continue;
                    }

                    if (alreadyPlacedStoryKeys is not null && alreadyPlacedStoryKeys.Contains((reference.Kind, reference.Id)))
                    {
                        continue;
                    }

                    if (placedStoryKeys.Contains((reference.Kind, reference.Id)))
                    {
                        continue;
                    }

                    if (!footnoteByKey.TryGetValue((reference.Kind, reference.Id), out DocxRelatedStoryLayout? storyLayout) ||
                        storyLayout.ContentHeight <= 0d)
                    {
                        continue;
                    }

                    if (!referenceIndex.IsReferenceRenderedOnPage(pageIndex, location))
                    {
                        continue;
                    }

                    placedStoryKeys.Add((reference.Kind, reference.Id));
                    pageFootnoteStories.Add(new DocxReferencedRelatedStoryLayout(location, storyLayout));
                }
            }

            int outputIndex = pageIndex + outputShift;
            int countBefore = outputPages.Count;
            DocxRelatedStoryLayout? footnoteContinuationLayout = FindSpecialRelatedStoryLayout(footnoteLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.ContinuationSeparator) ?? footnoteSeparatorLayout;
            PlaceFootnoteStories(outputPages, outputIndex, page, pageFootnoteStories, footnoteSeparatorLayout, footnoteContinuationLayout, separatorMeasurer, printScale, headerKeepOut,
                footerKeepOut: CanUsePlainFootnoteStaticFrame(document, page, separatorMeasurer, printScale, cancellationToken) ? footerKeepOut : null,
                availableFooterKeepOut: ownedFootnoteFooterKeepOut ?? footerKeepOut);
            outputShift += outputPages.Count - countBefore;
        }

        return AddPlacedEndnoteStories(outputPages);

        IReadOnlyList<DocxLayoutPage> AddPlacedEndnoteStories(IReadOnlyList<DocxLayoutPage> pages)
        {
            List<DocxInlineReferenceLocation> endnoteLocations = ResolveReferencedRelatedStoryLocations(document, DocxRelatedStoryKind.Endnote, placedStoryKeys).ToList();
            if (endnoteLocations.Count == 0)
            {
                return pages;
            }

            var outputPages = pages.ToList();
            // R12: classify with the reference scan only (the section-end page cannot
            // change the outcome: a found sectEnd reference always resolves); the page
            // itself is resolved once per section group below, after earlier groups may
            // have grown the page list.
            RelatedStoryPageIndex endnoteIndex = RelatedStoryPageIndex.Build(outputPages, cancellationToken);
            var documentEndLocations = new List<DocxInlineReferenceLocation>();
            var sectionEndLocations = new List<DocxInlineReferenceLocation>();
            foreach (DocxInlineReferenceLocation endnoteLocation in endnoteLocations)
            {
                cancellationToken.ThrowIfCancellationRequested();
                int referencePageIndex = endnoteIndex.FindFirstPageWithReference(endnoteLocation);
                if (referencePageIndex < 0 ||
                    !string.Equals(outputPages[referencePageIndex].PageSettings.EndnoteReferenceSettings.PositionValue, "sectEnd", StringComparison.OrdinalIgnoreCase))
                {
                    documentEndLocations.Add(endnoteLocation);
                    continue;
                }

                sectionEndLocations.Add(endnoteLocation);
            }

            var sectionRangeByBlock = new Dictionary<int, (int StartBlockIndex, int EndBlockIndex)>();
            foreach (IGrouping<(int StartBlockIndex, int EndBlockIndex), DocxInlineReferenceLocation> sectionGroup in sectionEndLocations
                         .GroupBy(location => ResolveMemoizedSectionBlockRange(location.SourceBlockIndex))
                         .OrderBy(group => group.Key.StartBlockIndex))
            {
                cancellationToken.ThrowIfCancellationRequested();
                endnoteIndex = RelatedStoryPageIndex.Build(outputPages, cancellationToken);
                int sectionEndPageIndex = ResolveSectionEndEndnotePageIndex(document.BodyElements, endnoteIndex, outputPages, sectionGroup.First());
                if (sectionEndPageIndex < 0)
                {
                    documentEndLocations.AddRange(sectionGroup);
                    continue;
                }

                IReadOnlyList<DocxRelatedStoryLayout> groupLayouts = resolveRelatedStoryLayouts(ResolvePageNoteTextWidth(outputPages[sectionEndPageIndex]));
                Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> endnoteStoryByKey = CreateRelatedStoryLookup(groupLayouts);
                DocxRelatedStoryLayout? endnoteSeparatorLayout = FindSpecialRelatedStoryLayout(groupLayouts, DocxRelatedStoryKind.Endnote, DocxRelatedStoryType.Separator);
                var sectionEndStories = new List<DocxReferencedRelatedStoryLayout>();
                foreach (DocxInlineReferenceLocation location in sectionGroup)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (TryResolveReferencedRelatedStoryLayout(endnoteStoryByKey, location, out DocxRelatedStoryLayout? storyLayout))
                    {
                        sectionEndStories.Add(new DocxReferencedRelatedStoryLayout(location, storyLayout));
                    }
                }

                if (sectionEndStories.Count > 0)
                {
                    DocxRelatedStoryLayout? endnoteContinuationLayout = FindSpecialRelatedStoryLayout(groupLayouts, DocxRelatedStoryKind.Endnote, DocxRelatedStoryType.ContinuationSeparator) ?? endnoteSeparatorLayout;
                    PlaceSectionEndEndnoteStories(outputPages, sectionEndPageIndex, sectionEndStories, endnoteSeparatorLayout, endnoteContinuationLayout, cancellationToken, printScale, separatorMeasurer, headerKeepOut);
                }
            }

            List<DocxRelatedStoryLayout> documentEndStories = [];
            if (documentEndLocations.Count > 0)
            {
                Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> endnoteStoryByKey = CreateRelatedStoryLookup(resolveRelatedStoryLayouts(ResolvePageNoteTextWidth(outputPages[^1])));
                foreach (DocxInlineReferenceLocation location in documentEndLocations)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (TryResolveReferencedRelatedStoryLayout(endnoteStoryByKey, location, out DocxRelatedStoryLayout? storyLayout))
                    {
                        documentEndStories.Add(storyLayout);
                    }
                }
            }

            return ReindexPageOwnedLayouts(AddDocumentEndnoteStories());

            (int StartBlockIndex, int EndBlockIndex) ResolveMemoizedSectionBlockRange(int sourceBlockIndex)
            {
                if (!sectionRangeByBlock.TryGetValue(sourceBlockIndex, out (int StartBlockIndex, int EndBlockIndex) range))
                {
                    range = ResolveSectionBlockRange(document.BodyElements, sourceBlockIndex);
                    sectionRangeByBlock[sourceBlockIndex] = range;
                }

                return range;
            }

            IReadOnlyList<DocxLayoutPage> AddDocumentEndnoteStories()
            {
                if (documentEndStories.Count == 0)
                {
                    return outputPages;
                }

                var documentEndPages = outputPages.ToList();
                DocxLayoutPage activePage = documentEndPages[^1];
                List<DocxPlacedRelatedStoryLayout> activePlacedStories = activePage.PlacedRelatedStories.ToList();
                double cursorTop = ResolveEndnoteStartTop(activePage, activePlacedStories, printScale);
                int activePageIndex = documentEndPages.Count - 1;
                DocxRelatedStoryLayout? documentEndSeparatorLayout = FindSpecialRelatedStoryLayout(resolveRelatedStoryLayouts(ResolvePageNoteTextWidth(activePage)), DocxRelatedStoryKind.Endnote, DocxRelatedStoryType.Separator);
                DocxRelatedStoryLayout? documentEndContinuationLayout = FindSpecialRelatedStoryLayout(resolveRelatedStoryLayouts(ResolvePageNoteTextWidth(activePage)), DocxRelatedStoryKind.Endnote, DocxRelatedStoryType.ContinuationSeparator) ?? documentEndSeparatorLayout;
                // Plain source/static semantics qualify the same footer edge used
                // by body flow. Scaled, dynamic and rich frames keep their fallback;
                // fitting footer edges equal the margin.
                bool hasFootnotes = document.RelatedStories.Any(story =>
                    story.IsNormalStoryType && story.Kind == DocxRelatedStoryKind.Footnote);
                Func<DocxLayoutPage, int, int, double>? endnoteFrameBottom = null;
                if (hasFootnotes)
                {
                    // Pages carrying both kinds share the physical footer edge.
                    // Empty endnote continuations retain their existing fallback.
                    if (footerKeepOut is { } footerBottom &&
                        CanUsePlainFootnoteStaticFrame(document, activePage, separatorMeasurer, printScale, cancellationToken))
                        endnoteFrameBottom = (p, number, count) => p.PlacedRelatedStories.Any(s =>
                            s.StoryLayout.Story.IsNormalStoryType && s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                            ? footerBottom(p, number, count) : p.MarginBottom;
                }
                else if (HasPlainStaticNoteStories(document, activePage, separatorMeasurer, printScale, cancellationToken) &&
                    CanUsePlainDocumentEndnoteFlow(activePage, documentEndSeparatorLayout, printScale,
                        documentEndStories, cancellationToken, out _)) endnoteFrameBottom = footerKeepOut;
                double FrameBottom() => endnoteFrameBottom?.Invoke(activePage, activePageIndex + 1, documentEndPages.Count) ?? activePage.MarginBottom;
                // A trailing paragraph owns this start boundary independently
                // of earlier tables or print scaling. Complete note-flow and
                // mixed-footnote admission retain their existing qualification.
                bool fittingEndnoteBoundaries = CanUseFittingEndnoteBoundaries(document, activePage, activePlacedStories,
                    documentEndSeparatorLayout, printScale, documentEndStories, cancellationToken);
                bool FollowsBodyParagraph() => fittingEndnoteBoundaries || CanUseTrailingParagraphEndnoteStart(activePage,
                    activePlacedStories, documentEndSeparatorLayout, printScale, documentEndStories);
                if (FollowsBodyParagraph() || CanUseNominalNoteSeparatorFlow(activePage, documentEndSeparatorLayout, printScale, documentEndStories))
                {
                    cursorTop -= ResolveTrailingBodyBoundarySpacing(activePage);
                }
                double mixedSourceEndnoteHeight = 0d;
                bool allowMixedEndnoteHead = hasFootnotes && endnoteFrameBottom is not null &&
                    CanUsePlainDocumentEndnoteFlow(activePage, documentEndSeparatorLayout, printScale,
                        documentEndStories, cancellationToken, out mixedSourceEndnoteHeight);
                DocxLayoutPage? mixedEndnoteSourceFrame = allowMixedEndnoteHead &&
                    HasPlainNoteFlowSources(document, DocxRelatedStoryKind.Footnote, cancellationToken,
                        out _, requirePlainParagraph: true) &&
                    CanUsePlainDocumentEndnoteFrame(document, activePage, activePageIndex + 1,
                        documentEndPages.Count, separatorMeasurer, printScale, cancellationToken) ? activePage : null;
                int endnoteSourcePageIndex = activePageIndex;
                bool deferredMixedEndnoteSource = false;
                cursorTop = ResolveDocumentEndnoteStartAboveFootnotes(activePage, activePlacedStories, cursorTop,
                    documentEndStories, documentEndSeparatorLayout, separatorMeasurer, printScale,
                    allowMixedEndnoteHead, out bool partialMixedEndnoteHead);
                if (partialMixedEndnoteHead && footerKeepOut is { } mixedFooterBottom)
                    endnoteFrameBottom = (p, number, count) => Math.Max(mixedFooterBottom(p, number, count),
                        ResolveFootnoteOccupiedTop(p, p.PlacedRelatedStories, separatorMeasurer, printScale));
                else if (hasFootnotes)
                {
                    // Fallback endnotes still clear occupied footnote containers.
                    // A raised footnote must not create apparent space underneath
                    // it. This placement boundary does not admit typography/reflow.
                    var qualifiedEndnoteFrameBottom = endnoteFrameBottom;
                    endnoteFrameBottom = (p, number, count) => Math.Max(
                        qualifiedEndnoteFrameBottom?.Invoke(p, number, count) ?? p.MarginBottom,
                        p.PlacedRelatedStories.Where(s => s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                            .Select(s => s.TopY).DefaultIfEmpty(p.MarginBottom).Max());
                }
                if (documentEndSeparatorLayout is not null && documentEndStories.Count > 0)
                {
                    // RV06 endnote probes: document-end endnotes draw the separator rule
                    // with a mark like section-end endnotes (no extra gap above it).
                    // RV06 height-model probes: content below the separator keeps a mark-metrics
                    // gap (the legacy constant underplaces growing marks), resolved through the
                    // same helper as footnote reserves; empty separators resolve through their
                    // pilcrow-carrying line and fall back to the legacy constant.
                    (DocxTextRun? endnoteMarkRun, double endnoteMarkSize) = FindSeparatorMarkFont(documentEndSeparatorLayout.TextLines);
                    // An overflowing separator would strand its rule at the margin while content
                    // slices onto fresh pages (Office keeps separator and content together);
                    // turn first, mirroring overflowing-item placement everywhere else.
                    if (ResolvePlainNoteSeparatorHeight(documentEndSeparatorLayout, activePage, separatorMeasurer, documentEndStories, printScale) > cursorTop - FrameBottom())
                    {
                        MoveToRelatedStoryContinuationPage(documentEndPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref cursorTop, insertContinuationAfterActivePage: false, headerKeepOut);
                    }
                    deferredMixedEndnoteSource = mixedEndnoteSourceFrame is not null && activePageIndex != endnoteSourcePageIndex;
                    double? nominalSeparatorHeight = deferredMixedEndnoteSource
                        ? ResolvePlainNoteSeparatorHeight(documentEndSeparatorLayout, mixedEndnoteSourceFrame!, separatorMeasurer, documentEndStories, printScale)
                        : FollowsBodyParagraph()
                            ? ResolveParagraphBoundarySeparatorHeight(documentEndSeparatorLayout, activePage, separatorMeasurer, documentEndStories)
                            : CanUseNominalNoteSeparatorFlow(activePage, documentEndSeparatorLayout, printScale, documentEndStories)
                            ? ResolvePlainNoteSeparatorHeight(documentEndSeparatorLayout, activePage, separatorMeasurer, documentEndStories, printScale) : null;
                    double separatorTop = cursorTop + FootnoteSeparatorGapPoints;
                    bool ownsSeparatorFirstBearing = false;
                    double separatorRuleFirstBearing = 0d;
                    if (deferredMixedEndnoteSource && separatorMeasurer is IDocxStaticTextMetricsProvider headMetrics)
                    {
                        // A fresh-page head keeps the source line's baseline floor,
                        // enlarged by the visible mark's measured ascent. Missing
                        // metrics retain the existing ride.
                        double firstInset = headMetrics.MeasureWindowsAscender(endnoteMarkRun, endnoteMarkSize);
                        if (firstInset > 0d && double.IsFinite(firstInset))
                        {
                            firstInset = Math.Max(firstInset, -documentEndSeparatorLayout.TextLines[0].BaselineY);
                            separatorTop = cursorTop + nominalSeparatorHeight!.Value - firstInset;
                            ownsSeparatorFirstBearing = true;
                        }
                    }
                    if (!deferredMixedEndnoteSource && !partialMixedEndnoteHead && mixedEndnoteSourceFrame is not null &&
                        separatorMeasurer is IDocxStaticTextMetricsProvider fittingMetrics &&
                        separatorMeasurer is IDocxLineMetricsProvider fittingLineMetrics)
                    {
                        double ascender = fittingMetrics.MeasureWindowsAscender(endnoteMarkRun, endnoteMarkSize);
                        double covered = Math.Min(
                            ResolveSeparatorHheaAscenderExcess(endnoteMarkRun, endnoteMarkSize, separatorMeasurer),
                            ResolveContentHheaAscenderExcess(documentEndStories, separatorMeasurer));
                        double sourceInset = -documentEndSeparatorLayout.TextLines[0].BaselineY;
                        if (double.IsFinite(ascender) && ascender > 0d &&
                            double.IsFinite(covered) && covered >= 0d && double.IsFinite(sourceInset) && sourceInset > 0d)
                        {
                            // The content gap shares the resolved deficit only
                            // when this origin already owns the first bearing.
                            // Uncovered bearings keep their qualified origin/gap.
                            ownsSeparatorFirstBearing = Math.Max(ascender, sourceInset) <=
                                DocxLineMetrics.WordAutoLineBaselineOffsetEm * endnoteMarkSize + covered;
                            double hheaAscender = fittingLineMetrics.MeasureHheaAscender(endnoteMarkRun, endnoteMarkSize);
                            double paintCoverage = ResolveContentHheaAscenderExcess(documentEndStories,
                                separatorMeasurer, endnoteMarkSize);
                            if (double.IsFinite(hheaAscender) && hheaAscender > 0d &&
                                double.IsFinite(paintCoverage) && paintCoverage >= 0d)
                            {
                                // Rule painting retains the source first bearing
                                // confirmed by line metrics. Content covers its
                                // font bearing at the mark size; flow keeps its
                                // existing source-size charge and content gap.
                                separatorRuleFirstBearing = Math.Max(0d,
                                    Math.Min(sourceInset, hheaAscender) -
                                    DocxLineMetrics.WordAutoLineBaselineOffsetEm * endnoteMarkSize - paintCoverage);
                            }
                        }
                    }
                    (DocxPlacedRelatedStoryLayout placedSeparator, double separatorBottom) = PlaceSeparatorStoryWithMark(activePage, activePageIndex, documentEndSeparatorLayout, sourceBlockIndex: -1, separatorTop, separatorMeasurer, useSizeDrivenPlacementHeight: true, contentStoriesForHhea: documentEndStories, placementHeightOverride: nominalSeparatorHeight, ruleFirstBearingPoints: separatorRuleFirstBearing);
                    activePlacedStories.Add(placedSeparator);
                    documentEndPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
                    double endnoteContentGapPoints = 0d;
                    foreach (DocxRelatedStoryLayout endnoteContentLayout in documentEndStories)
                    {
                        endnoteContentGapPoints = System.Math.Max(endnoteContentGapPoints, ResolveFootnoteContentGapPoints(endnoteContentLayout.TextLines, endnoteMarkRun, endnoteMarkSize, separatorMeasurer));
                    }

                    cursorTop = separatorBottom - (ownsSeparatorFirstBearing
                        ? ResolveSourceInsetNoteContentGap(endnoteContentGapPoints, endnoteMarkRun, endnoteMarkSize, separatorMeasurer)
                        : Math.Max(ResolveSeparatorGapPoints(endnoteMarkRun, endnoteMarkSize, separatorMeasurer), endnoteContentGapPoints));
                }
                double completeEndnoteHeight = 0d;
                bool plainEndnoteTypography = CanUsePlainDocumentEndnoteFrame(document, activePage,
                    activePageIndex + 1, documentEndPages.Count, separatorMeasurer, printScale, cancellationToken) &&
                    CanUsePlainDocumentEndnoteFlow(activePage,
                    documentEndSeparatorLayout, printScale, documentEndStories, cancellationToken,
                    out completeEndnoteHeight);
                if (deferredMixedEndnoteSource)
                {
                    // An empty continuation inherits the admitted original source
                    // frame; a page turn does not change its paragraph ownership.
                    plainEndnoteTypography = true;
                    completeEndnoteHeight = mixedSourceEndnoteHeight;
                }
                // Mixed static frames share typography and qualified head space;
                // their paragraph-boundary policy remains unchanged.
                bool plainEndnoteFrame = plainEndnoteTypography &&
                    (!hasFootnotes || DocxPlainBodyFrame.IsEligible(document));
                bool plainEndnoteFlow = plainEndnoteFrame && !hasFootnotes;
                bool collapseEndnoteBoundaries = fittingEndnoteBoundaries || plainEndnoteFrame &&
                    cursorTop - completeEndnoteHeight >= Math.Max(FrameBottom(),
                        activePlacedStories.Where(s => s.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
                            .Select(s => s.TopY).DefaultIfEmpty(activePage.MarginBottom).Max());
                DocxRelatedStoryLayout? previousEndnote = null;
                foreach (DocxRelatedStoryLayout endnoteStoryLayout in documentEndStories)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    double boundaryOverlap = (collapseEndnoteBoundaries || plainEndnoteFlow) && previousEndnote is not null
                        ? ResolveDocumentEndnoteBoundaryOverlap(previousEndnote, endnoteStoryLayout) : 0d;
                    // A shared paragraph boundary belongs on this page only when
                    // the next paragraph's head fits. A page turn retains its
                    // existing placement and destination policy.
                    if (!collapseEndnoteBoundaries && plainEndnoteFlow &&
                        cursorTop + boundaryOverlap - ResolveStoryTextLineBoxHeight(endnoteStoryLayout.TextLines[0],
                            endnoteStoryLayout.TextLines.Count == 1
                                ? endnoteStoryLayout.TextLines[0].ParagraphAfterSpacing ?? 0d : 0d) < FrameBottom())
                    {
                        boundaryOverlap = 0d;
                    }
                    cursorTop += boundaryOverlap;
                    PlaceRelatedStorySlices(
                        documentEndPages,
                        ref activePageIndex,
                        ref activePage,
                        ref activePlacedStories,
                        ref cursorTop,
                        endnoteStoryLayout,
                        sourceBlockIndex: -1,
                        insertContinuationAfterActivePage: false,
                        documentEndContinuationLayout,
                        separatorMeasurer,
                        printScale,
                        headerKeepOut,
                        plainDocumentEndnoteFlow: plainEndnoteFlow || partialMixedEndnoteHead,
                        // Typography uses the admitted frame even when a
                        // complementary footnote retains legacy pagination.
                        plainDocumentEndnoteTypography: plainEndnoteTypography,
                        noteDestinationTopOffset: boundaryOverlap,
                        frameBottom: endnoteFrameBottom);
                    if (collapseEndnoteBoundaries)
                    {
                        activePlacedStories[^1] = activePlacedStories[^1] with
                        {
                            UsesCompleteDocumentEndnoteFlow = true,
                        };
                        documentEndPages[activePageIndex] = activePage with
                        {
                            PlacedRelatedStories = activePlacedStories.ToArray(),
                        };
                    }
                    if (collapseEndnoteBoundaries || plainEndnoteFlow)
                    {
                        // Qualified document-end paragraphs share their boundary;
                        // the generic note separator gap belongs to legacy flow.
                        cursorTop += FootnoteSeparatorGapPoints;
                    }
                    previousEndnote = endnoteStoryLayout;
                }

                return documentEndPages;
            }
        }
    }

    private static void PlaceSectionEndEndnoteStories(
        List<DocxLayoutPage> outputPages,
        int sectionEndPageIndex,
        IReadOnlyList<DocxReferencedRelatedStoryLayout> sectionEndStories,
        DocxRelatedStoryLayout? endnoteSeparatorLayout,
        DocxRelatedStoryLayout? endnoteContinuationLayout,
        CancellationToken cancellationToken,
        double printScale,
        IDocxTextMeasurer? separatorMeasurer, 
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null)
    {
        int activePageIndex = sectionEndPageIndex;
        DocxLayoutPage activePage = outputPages[activePageIndex];
        List<DocxPlacedRelatedStoryLayout> activePlacedStories = activePage.PlacedRelatedStories.ToList();
        double cursorTop = ResolveEndnoteStartTop(activePage, activePlacedStories, printScale);
        if (endnoteSeparatorLayout is not null && sectionEndStories.Count > 0)
        {
            // RV06 endnote probes: the separator space sits exactly one body pitch below
            // the body baseline, so no extra gap applies above the separator (the start-top
            // gap is added back); content keeps its gap below the separator.
            (DocxPlacedRelatedStoryLayout placedSeparator, double separatorBottom) = PlaceSeparatorStoryWithMark(activePage, activePageIndex, endnoteSeparatorLayout, sectionEndStories[0].Location.SourceBlockIndex, cursorTop + FootnoteSeparatorGapPoints, separatorMeasurer);
            activePlacedStories.Add(placedSeparator);
            outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
            cursorTop = separatorBottom - FootnoteSeparatorGapPoints;
        }
        foreach (DocxReferencedRelatedStoryLayout story in sectionEndStories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PlaceRelatedStorySlices(
                outputPages,
                ref activePageIndex,
                ref activePage,
                ref activePlacedStories,
                ref cursorTop,
                story.StoryLayout,
                story.Location.SourceBlockIndex,
                insertContinuationAfterActivePage: true,
                endnoteContinuationLayout,
                separatorMeasurer,
                printScale,
                headerKeepOut);
        }
    }

    // RV06 continuation probes (Word 16.0, footlong/endlong): long notes split
    // between whole lines with continuation rules on later pages. Line boxes tile the
    // story from its top edge (line box plus paragraph after-spacing, with
    // contextually suppressed before-spacing excluded); a slice takes the longest
    // whole-line prefix that fits, at least one line, so degenerate capacities
    // terminate instead of looping. Stories without text lines keep legacy slicing.
    private static double[] GetStoryTextLineBoxHeights(DocxRelatedStoryLayout storyLayout,
        IReadOnlyList<double>? ownedAfterSpacings = null)
    {
        double[] heights = new double[storyLayout.TextLines.Count];
        for (int lineIndex = 0; lineIndex < heights.Length; lineIndex++)
        {
            DocxTextLineLayout line = storyLayout.TextLines[lineIndex];
            double afterSpacing = ownedAfterSpacings?[lineIndex] ?? line.ParagraphAfterSpacing ?? line.PendingAfterSpacing ?? 0d;
            heights[lineIndex] = ResolveStoryTextLineBoxHeight(line, afterSpacing);
        }

        return heights;
    }

    private static double ResolveStoryTextLineBoxHeight(DocxTextLineLayout line, double afterSpacing)
    {
        double boxHeight = line.LineHeight ?? line.SingleLineHeight ?? line.FontSize;
        double beforeSpacing = line.IsFirstParagraphLine == true && line.ContextualSpacingSuppressed != true
            ? line.ParagraphBeforeSpacing ?? line.AppliedBeforeSpacing ?? 0d : 0d;
        return Math.Max(0d, boxHeight) + Math.Max(0d, afterSpacing) + Math.Max(0d, beforeSpacing);
    }

    private static double[] GetStoryLineAfterSpacings(DocxRelatedStoryLayout storyLayout, bool deferParagraphAfter = false)
    {
        double[] afters = new double[storyLayout.TextLines.Count];
        for (int lineIndex = 0; lineIndex < afters.Length; lineIndex++)
        {
            DocxTextLineLayout line = storyLayout.TextLines[lineIndex];
            afters[lineIndex] = Math.Max(0d, line.ParagraphAfterSpacing ?? line.PendingAfterSpacing ?? 0d);
        }

        if (deferParagraphAfter && afters.Length != 0)
        {
            // Qualified plain document-end flow owns one source paragraph. Its
            // after-space metadata lives on the first line, but its charge belongs
            // at the end, both when checking a head and placing its continuations.
            double paragraphAfter = afters[0];
            afters[0] = 0d;
            afters[^1] = paragraphAfter;
        }
        return afters;
    }

    private static DocxRelatedStoryLayout NarrowStoryTextLines(DocxRelatedStoryLayout storyLayout, int startLineIndex, int lineCount)
    {
        if (startLineIndex == 0 && lineCount >= storyLayout.TextLines.Count)
        {
            return storyLayout;
        }

        DocxTextLineLayout[] sliceLines = new DocxTextLineLayout[lineCount];
        for (int lineIndex = 0; lineIndex < lineCount; lineIndex++)
        {
            sliceLines[lineIndex] = storyLayout.TextLines[startLineIndex + lineIndex];
        }

        return storyLayout with { TextLines = sliceLines };
    }

    // RV06 interleaving: exact in-flight remainder for the drain-aware body reserve.
    private static bool TryCompleteFootnoteReserve(IEnumerable<InFlightRelatedStory> notes, out double reserve,
        bool excludeFinalAfterSpacing = false)
    {
        reserve = 0d;
        double finalAfterSpacing = 0d;
        DocxRelatedStoryLayout? head = null;
        foreach (InFlightRelatedStory note in notes)
        {
            if (note.PlacedLineCount != 0 || note.RemainingLineCount != 1 || note.SeparatorPlaced ||
                note.StoryLayout.TextLines.Count != 1 || note.SeparatorLayout is not { } separator ||
                separator.TextLines.Count != 1 || !separator.TextLines[0].Text.All(char.IsWhiteSpace)) return false;
            if (head is not null && !ReferenceEquals(head, separator)) return false;
            head = separator;
            double[] boxes = GetStoryTextLineBoxHeights(note.StoryLayout);
            if (boxes.Length != 1 || !double.IsFinite(boxes[0]) || boxes[0] <= 0d) return false;
            reserve += boxes[0];
            finalAfterSpacing = Math.Max(0d, note.StoryLayout.TextLines[0].ParagraphAfterSpacing ??
                note.StoryLayout.TextLines[0].PendingAfterSpacing ?? 0d);
        }
        if (head is null) return true;
        // Complete notes share one separator line. The current page does not
        // reserve a continuation rule for each note or repeat the head's trailing
        // paragraph spacing. Source-owned admission also lets the final note's
        // trailing space collapse into the remaining body/separator capacity;
        // earlier notes retain their inter-paragraph space.
        double? headLine = head.TextLines[0].LineHeight ?? head.TextLines[0].SingleLineHeight;
        if (headLine is not > 0d || !double.IsFinite(headLine.Value)) return false;
        reserve += headLine.Value;
        if (excludeFinalAfterSpacing) reserve -= finalAfterSpacing;
        return double.IsFinite(reserve);
    }

    private static void FitCompleteFootnoteTrailingSpacing(List<DocxReferencedRelatedStoryLayout> notes, double capacity)
    {
        if (notes.Count == 0) return;
        DocxReferencedRelatedStoryLayout final = notes[^1];
        DocxTextLineLayout line = final.StoryLayout.TextLines.Single();
        double after = Math.Max(0d, line.ParagraphAfterSpacing ?? line.PendingAfterSpacing ?? 0d);
        double trim = Math.Min(after, Math.Max(0d, notes.Sum(note => note.StoryLayout.ContentHeight) - capacity));
        if (trim <= 0d) return;
        // Placement consumes the same optional final space as admission. Keep
        // canonical source/layout records intact; only this page's take changes.
        // Otherwise the plain-frame slicer can defer an already admitted line
        // merely because its full trailing paragraph space does not fit.
        notes[^1] = final with { StoryLayout = final.StoryLayout with
        {
            TextLines = [line with { ParagraphAfterSpacing = after - trim }],
            ContentHeight = final.StoryLayout.ContentHeight - trim
        }};
    }

    private static bool TrySharedFootnoteFrameReserve(IEnumerable<InFlightRelatedStory> notes,
        DocxLayoutPage frame, IDocxTextMeasurer? measurer, CancellationToken cancellationToken, out double reserve)
    {
        reserve = 0d;
        InFlightRelatedStory[] remaining = notes.Where(n => n.RemainingLineCount > 0).ToArray();
        if (remaining.Length == 0) return true;
        foreach (InFlightRelatedStory note in remaining)
        {
            cancellationToken.ThrowIfCancellationRequested();
            double height = note.RemainingContentHeight(GetStoryTextLineBoxHeights(note.StoryLayout));
            if (!double.IsFinite(height) || height < 0d) return false;
            reserve += height;
        }
        InFlightRelatedStory first = remaining[0];
        if (first.SeparatorPlaced)
        {
            reserve += ContinuationRuleOverhead(first.ContinuationLayout, measurer);
        }
        else if (first.SeparatorLayout is { } separator)
        {
            (DocxTextRun? mark, double size) = FindSeparatorMarkFont(separator.TextLines);
            double gap = Math.Max(ResolveSeparatorGapPoints(mark, size, measurer),
                remaining.Max(n => ResolveFootnoteContentGapPoints(n.StoryLayout.TextLines, mark, size, measurer)));
            reserve += ResolvePlainNoteSeparatorHeight(separator, frame, measurer,
                remaining.Select(n => n.StoryLayout).ToArray(), 1d) + gap;
        }
        else reserve += ResolveSeparatorGapPoints(null, 0d, measurer);
        return double.IsFinite(reserve);
    }

    // Content remainder sums take boxes (the same boxes takes consume); the separator
    // counts only until placed on the head page; one continuation-rule overhead covers
    // the next shared slice page. Same quantities the slice path places with.
    internal static double FootnoteRemainderHeight(
        InFlightRelatedStory inFlight,
        IReadOnlyList<double> lineBoxes,
        IDocxTextMeasurer? separatorMeasurer)
    {
        double remainder = inFlight.RemainingContentHeight(lineBoxes);
        if (!inFlight.SeparatorPlaced)
        {
            if (inFlight.SeparatorLayout is null)
            {
                remainder += ResolveSeparatorGapPoints(null, 0d, separatorMeasurer);
            }
            else
            {
                (DocxTextRun? reserveRun, double reserveFontSizePoints) = FindSeparatorMarkFont(inFlight.SeparatorLayout.TextLines);
                remainder += Math.Max(0d, inFlight.SeparatorLayout.ContentHeight) + ResolveSeparatorGapPoints(reserveRun, reserveFontSizePoints, separatorMeasurer);
            }
        }

        // A new one-line note is indivisible: its complete head is already
        // charged above, and cannot need a future continuation rule. A tail
        // still owns its current continuation rule even when one line remains.
        if (inFlight.RemainingLineCount > 1 ||
            inFlight.RemainingLineCount > 0 && inFlight.SeparatorPlaced)
        {
            remainder += ContinuationRuleOverhead(inFlight.ContinuationLayout, separatorMeasurer);
        }

        return Math.Max(0d, remainder);
    }

    private static double ContinuationRuleOverhead(
        DocxRelatedStoryLayout? continuationLayout,
        IDocxTextMeasurer? separatorMeasurer)
    {
        if (continuationLayout is null || continuationLayout.TextLines.Count == 0)
        {
            return 0d;
        }

        (DocxTextRun? gapRun, double gapFontSizePoints) = FindSeparatorMarkFont(continuationLayout.TextLines);
        double gapPoints = ResolveSeparatorGapPoints(gapRun, gapFontSizePoints, separatorMeasurer);
        (double ruleBottomOffsetPoints, double ruleThicknessPoints) = ResolveSeparatorRuleGeometry(continuationLayout.Story.Kind, gapRun, gapFontSizePoints, separatorMeasurer);
        double firstInsetPoints = Math.Max(0d, -continuationLayout.TextLines[0].BaselineY);
        return gapPoints + Math.Max(firstInsetPoints, ruleBottomOffsetPoints + ruleThicknessPoints);
    }

    // RV06 take-battery probes (Word 16.0, after-spacing and Palatino takes): Office
    // continuation takes lap the rule block into the top margin (cal9-a8 laps 0.72 and
    // cal-pal laps 0.36 above the frame top) while a4 refuses the next line at lap 4.54,
    // and the story-final line keeps its after-spacing so no constant full-block charge
    // fits the battery; the take-side charge is the emitted rule ride (gap plus rule
    // offset plus thickness; first-line insets never ride above content) net of a 2.1pt
    // overflow tolerance, floored at zero so mark-less separators keep legacy takes.
    private static double ContinuationTakeCharge(
        DocxRelatedStoryLayout? continuationLayout,
        IDocxTextMeasurer? separatorMeasurer,
        bool hasPrecedingFragment = true)
    {
        if (continuationLayout is null || continuationLayout.TextLines.Count == 0)
        {
            return 0d;
        }

        (DocxTextRun? takeGapRun, double takeGapFontSizePoints) = FindSeparatorMarkFont(continuationLayout.TextLines);
        double takeGapPoints = ResolveSeparatorGapPoints(takeGapRun, takeGapFontSizePoints, separatorMeasurer);
        (double takeRuleBottomOffsetPoints, double takeRuleThicknessPoints) = ResolveSeparatorRuleGeometry(continuationLayout.Story.Kind, takeGapRun, takeGapFontSizePoints, separatorMeasurer);
        // The measured lap belongs to a continued tail. A deferred source head
        // has no preceding fragment and must leave the complete rule ride clear.
        return Math.Max(0d, takeGapPoints + takeRuleBottomOffsetPoints + takeRuleThicknessPoints -
            (hasPrecedingFragment ? 2.1 : 0d));
    }

    private static void PlaceRelatedStorySlices(
        List<DocxLayoutPage> outputPages,
        ref int activePageIndex,
        ref DocxLayoutPage activePage,
        ref List<DocxPlacedRelatedStoryLayout> activePlacedStories,
        ref double cursorTop,
        DocxRelatedStoryLayout storyLayout,
        int sourceBlockIndex,
        bool insertContinuationAfterActivePage,
        DocxRelatedStoryLayout? continuationSeparatorLayout = null,
        IDocxTextMeasurer? separatorMeasurer = null,
        double printScale = 1d, 
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null,
        bool stopAfterCurrentPage = false,
        InFlightRelatedStory? inFlight = null,
        double headContentPushdownPoints = 0d,
        bool plainDocumentEndnoteFlow = false,
        bool plainDocumentEndnoteTypography = false,
        double noteDestinationTopOffset = 0d,
        Func<DocxLayoutPage, int, int, double>? frameBottom = null,
        bool usesPlainFootnoteFrame = true)
    {
        if (storyLayout.TextLines.Count == 0)
        {
            PlaceRelatedStorySlicesByHeight(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref cursorTop, storyLayout, sourceBlockIndex, insertContinuationAfterActivePage);
            return;
        }

        double[] lineAfters = GetStoryLineAfterSpacings(storyLayout, deferParagraphAfter: plainDocumentEndnoteFlow);
        double[] lineBoxes = GetStoryTextLineBoxHeights(storyLayout,
            plainDocumentEndnoteFlow ? lineAfters : null);
        int lineIndex = 0;
        double consumedHeight = 0d;
        bool plainParagraphFlow = plainDocumentEndnoteFlow ||
            frameBottom is not null && storyLayout.Story.Kind == DocxRelatedStoryKind.Footnote;
        while (lineIndex < storyLayout.TextLines.Count)
        {
            double FrameBottom(DocxLayoutPage page, int index) => frameBottom?.Invoke(page, index + 1, outputPages.Count) ?? page.MarginBottom;
            double availableHeight = Math.Max(0d, cursorTop - FrameBottom(activePage, activePageIndex));
            if (availableHeight <= 0.001d)
            {
                if (stopAfterCurrentPage)
                {
                    return;
                }

                MoveToRelatedStoryContinuationPage(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref cursorTop, insertContinuationAfterActivePage, headerKeepOut);
                availableHeight = Math.Max(0d, cursorTop - FrameBottom(activePage, activePageIndex));
                cursorTop = PlaceContinuationSeparatorIfNeeded(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, cursorTop, sourceBlockIndex, storyLayout, lineIndex, lineBoxes, continuationSeparatorLayout, separatorMeasurer, printScale, frameBottom: frameBottom);
                availableHeight = Math.Max(0d, cursorTop - FrameBottom(activePage, activePageIndex));
            }

            int remainingLines = storyLayout.TextLines.Count - lineIndex;
            double[]? fitAfters = printScale >= 1d ? lineAfters : null;
            bool hasPageContent = activePage.Items.Count != 0 ||
                activePlacedStories.Any(s => s.StoryLayout.Story.IsNormalStoryType);
            (int takeCount, double takeHeight) = plainParagraphFlow
                ? TakeParagraphStoryLines(storyLayout.TextLines[0].SourceParagraph!, lineBoxes, lineIndex,
                    remainingLines, availableHeight, fitAfters, hasPageContent,
                    lineIndex == 0 && (inFlight is null || inFlight.PlacedLineCount == 0))
                : TakeStoryLines(lineBoxes, lineIndex, remainingLines, availableHeight, fitAfters);
            if (plainParagraphFlow && hasPageContent && takeCount == 0)
            {
                if (stopAfterCurrentPage) return;
                // An occupied frame must fit a source line before placing it;
                // the note head also follows the body's default/on/off widow rule.
                // An empty continuation page always makes progress, even when
                // its frame can hold only one line. Preserve the source boundary
                // contribution after the turn resets the page cursor.
                MoveToRelatedStoryContinuationPage(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories,
                    ref cursorTop, insertContinuationAfterActivePage, headerKeepOut);
                cursorTop = PlaceContinuationSeparatorIfNeeded(outputPages, ref activePageIndex, ref activePage,
                    ref activePlacedStories, cursorTop, sourceBlockIndex, storyLayout, lineIndex, lineBoxes,
                    continuationSeparatorLayout, separatorMeasurer, printScale, frameBottom: frameBottom) + (lineIndex == 0 ? noteDestinationTopOffset : 0d);
                continue;
            }
            if (lineIndex + takeCount >= storyLayout.TextLines.Count)
            {
                takeHeight = Math.Min(Math.Max(0d, storyLayout.ContentHeight - consumedHeight), availableHeight);
            }

            DocxRelatedStoryLayout sliceLayout = NarrowStoryTextLines(storyLayout, lineIndex, takeCount);
            // RV06 mixlong probes (Word 16.0): overflowing head slices hang content below
            // the take-driven rule like fitting notes instead of seating at storiesTop,
            // while takes, take heights and the cursor chain keep full boxes.
            double placementTop = lineIndex == 0 ? cursorTop - Math.Max(0d, headContentPushdownPoints) : cursorTop;
            if (inFlight is not null)
            {
                inFlight.PlacedLineCount += takeCount;
            }

            activePlacedStories.Add(PlaceRelatedStoryAtTop(
                activePage,
                activePageIndex,
                sliceLayout,
                sourceBlockIndex,
                placementTop,
                consumedHeight,
                takeHeight,
                separatorY: null) with
            {
                UsesPlainDocumentEndnoteTypography = plainDocumentEndnoteTypography,
                UsesPlainFootnoteFrame = usesPlainFootnoteFrame && frameBottom is not null && storyLayout.Story.Kind == DocxRelatedStoryKind.Footnote,
                NoteDestinationTopOffset = lineIndex == 0 ? noteDestinationTopOffset : 0d,
            });
            outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
            lineIndex += takeCount;
            consumedHeight += takeHeight;
            cursorTop -= takeHeight + FootnoteSeparatorGapPoints;
            if (lineIndex < storyLayout.TextLines.Count)
            {
                if (stopAfterCurrentPage)
                {
                    return;
                }

                MoveToRelatedStoryContinuationPage(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref cursorTop, insertContinuationAfterActivePage, headerKeepOut);
                cursorTop = PlaceContinuationSeparatorIfNeeded(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, cursorTop, sourceBlockIndex, storyLayout, lineIndex, lineBoxes, continuationSeparatorLayout, separatorMeasurer, printScale, frameBottom: frameBottom);
            }
        }
    }

    private static void PlaceRelatedStorySlicesByHeight(
        List<DocxLayoutPage> outputPages,
        ref int activePageIndex,
        ref DocxLayoutPage activePage,
        ref List<DocxPlacedRelatedStoryLayout> activePlacedStories,
        ref double cursorTop,
        DocxRelatedStoryLayout storyLayout,
        int sourceBlockIndex,
        bool insertContinuationAfterActivePage, 
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null)
    {
        double remainingHeight = Math.Max(0d, storyLayout.ContentHeight);
        double storyTopOffset = 0d;
        while (remainingHeight > 0.001d)
        {
            double availableHeight = Math.Max(0d, cursorTop - activePage.MarginBottom);
            if (availableHeight <= 0.001d)
            {
                MoveToRelatedStoryContinuationPage(
                    outputPages,
                    ref activePageIndex,
                    ref activePage,
                    ref activePlacedStories,
                    ref cursorTop,
                    insertContinuationAfterActivePage, headerKeepOut);
                availableHeight = Math.Max(0d, cursorTop - activePage.MarginBottom);
            }

            double sliceHeight = Math.Min(remainingHeight, availableHeight);
            if (sliceHeight <= 0.001d)
            {
                break;
            }

            activePlacedStories.Add(PlaceRelatedStoryAtTop(
                activePage,
                activePageIndex,
                storyLayout,
                sourceBlockIndex,
                cursorTop,
                storyTopOffset,
                sliceHeight,
                separatorY: null));
            outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
            storyTopOffset += sliceHeight;
            remainingHeight -= sliceHeight;
            cursorTop -= sliceHeight + FootnoteSeparatorGapPoints;

            if (remainingHeight > 0.001d)
            {
                MoveToRelatedStoryContinuationPage(
                    outputPages,
                    ref activePageIndex,
                    ref activePage,
                    ref activePlacedStories,
                    ref cursorTop,
                    insertContinuationAfterActivePage, headerKeepOut);
            }
        }
    }

    // RV06 continuation probes (Word 16.0, footlong/endlong): continued notes open
    // each later page with a full-width continuation rule plus end mark. The story
    // starts fresh at the continuation page top, so the mark baseline lands one
    // first-line inset below the top with no extra ride; rule geometry follows the
    // shared strikeout path and content keeps the footnote gap below the rule.
    // Longest fitting whole-line prefix. Qualified callers reserve mandatory
    // one-line progress for empty frames; other callers retain their prior rule.
    // RV06 after-spacing sweep (Word 16.0, seven after values): take needs exclude the
    // take last-line trailing after-spacing while take heights keep full boxes for
    // placement, so final takes and null after-spacing callers resolve byte-identically.
    // The exclusion applies on unscaled takes only; word-compatible spacing-scale takes
    // stay legacy pending their own probe series, so callers pass null below scale 1.
    private static (int TakeCount, double TakeHeight) TakeStoryLines(double[] lineBoxes, int startLineIndex, int remainingLineCount, double capacityHeight, double[]? lineAfterSpacings = null, bool forceProgress = true)
    {
        int takeCount = 0;
        double takeHeight = 0d;
        int lastLineIndex = startLineIndex + remainingLineCount - 1;
        while (takeCount < remainingLineCount)
        {
            int lineIndex = startLineIndex + takeCount;
            double fitHeight = takeHeight + lineBoxes[lineIndex];
            if (lineAfterSpacings is not null && lineIndex < lastLineIndex)
            {
                fitHeight -= Math.Max(0d, lineAfterSpacings[lineIndex]);
            }

            if (fitHeight > capacityHeight + 0.001d)
            {
                break;
            }

            takeHeight += lineBoxes[lineIndex];
            takeCount++;
        }

        if (takeCount == 0 && forceProgress)
        {
            takeCount = 1;
            takeHeight = lineBoxes[startLineIndex];
        }

        return (takeCount, takeHeight);
    }

    // Seating and emission must choose the same paragraph fragment. Otherwise a
    // deferred head can reserve a line that the slicer never emits. Use the body's
    // default/on/off widow rules and preserve progress on an empty frame.
    private static (int TakeCount, double TakeHeight) TakeParagraphStoryLines(
        DocxParagraph paragraph, double[] lineBoxes, int startLineIndex, int remainingLineCount,
        double capacityHeight, double[]? lineAfterSpacings, bool hasPageContent, bool isFreshHead)
    {
        (int takeCount, double takeHeight) = TakeStoryLines(lineBoxes, startLineIndex, remainingLineCount,
            capacityHeight, lineAfterSpacings, forceProgress: !hasPageContent);
        if (takeCount > 1 && DocxParagraphFragmentPolicy.HasOneLineTail(paragraph, remainingLineCount, takeCount))
        {
            takeCount--;
            takeHeight -= lineBoxes[startLineIndex + takeCount];
        }
        if (hasPageContent && isFreshHead &&
            DocxParagraphFragmentPolicy.HasOneLineHead(paragraph, remainingLineCount, takeCount))
        {
            return (0, 0d);
        }
        return (takeCount, takeHeight);
    }

    // RV06 continuation probes (Word 16.0, footlong/endlong): continued notes open
    // later pages with a full-width continuation rule plus end mark. Endnote
    // continuations start fresh at the page top (mark one first-line inset below the
    // top); footnote continuations bottom-anchor (mark one footnote gap above the
    // content top, which sits one take above the margin). Rule geometry follows the
    // shared strikeout path and content keeps the footnote gap below the rule.
    private static double PlaceContinuationSeparatorIfNeeded(
        List<DocxLayoutPage> outputPages,
        ref int activePageIndex,
        ref DocxLayoutPage activePage,
        ref List<DocxPlacedRelatedStoryLayout> activePlacedStories,
        double cursorTop,
        int sourceBlockIndex,
        DocxRelatedStoryLayout bodyStoryLayout,
        int bodyLineIndex,
        double[] bodyLineBoxes,
        DocxRelatedStoryLayout? continuationSeparatorLayout,
        IDocxTextMeasurer? separatorMeasurer,
        double printScale = 1d, 
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null,
        Func<DocxLayoutPage, int, int, double>? frameBottom = null)
    {
        if (continuationSeparatorLayout is null || continuationSeparatorLayout.TextLines.Count == 0)
        {
            return cursorTop;
        }

        (DocxTextRun? gapRun, double gapFontSizePoints) = FindSeparatorMarkFont(continuationSeparatorLayout.TextLines);
        double separatorGapPoints = ResolveSeparatorGapPoints(gapRun, gapFontSizePoints, separatorMeasurer);
        double markBaselineY;
        if (continuationSeparatorLayout.Story.Kind == DocxRelatedStoryKind.Footnote)
        {
            int remainingLines = bodyStoryLayout.TextLines.Count - bodyLineIndex;
            // RV06 continuation probes: the footnote bottom-anchor query charges the rule
            // overhead net of already-reserved top space, so continuation takes do not place
            // the rule for free while header keep-out pages keep their exact takes; the endnote
            // top-anchor branch already nets first inset plus gap through cursorTop.
            double[]? bodyFitAfters = printScale >= 1d ? GetStoryLineAfterSpacings(bodyStoryLayout) : null;
            // RV06 take-battery probes (Word 16.0): the bottom-anchor take counts full
            // boxes against the ride-sized take charge above, since the story-final line
            // keeps its after-spacing under any constant full-block charge; word-compatible
            // takes stay legacy pending the balloon lane, so scaled pages keep the shared
            // full-block charge with null needs exactly as today.
            bool hasPrecedingFragment = bodyLineIndex > 0 || bodyStoryLayout.TextLines[0].SourceLineIndex > 0;
            // An admitted physical continuation frame owns the separator's first
            // inset as well as its visible rule. Charging only the rule ride lets
            // an extra source line and the separator clip cross the frame top.
            // Deferred heads and legacy frames retain their qualified ride policy.
            double continuationOverhead = printScale >= 1d && !(frameBottom is not null && hasPrecedingFragment)
                ? ContinuationTakeCharge(continuationSeparatorLayout, separatorMeasurer,
                    frameBottom is null || hasPrecedingFragment)
                : ContinuationRuleOverhead(continuationSeparatorLayout, separatorMeasurer);
            double reservedTopSpace = Math.Max(0d, activePage.Height - activePage.MarginTop - cursorTop);
            double netOverhead = Math.Max(0d, continuationOverhead - reservedTopSpace);
            double bottom = frameBottom?.Invoke(activePage, activePageIndex + 1, outputPages.Count) ?? activePage.MarginBottom;
            double capacity = Math.Max(0d, cursorTop - bottom - netOverhead);
            bool hasPageContent = activePage.Items.Count != 0 ||
                activePlacedStories.Any(s => s.StoryLayout.Story.IsNormalStoryType);
            // A qualified footnote frame anchors the actual widow-aware take,
            // rather than leaving space for a line moved into the next slice.
            (_, double takeHeight) = frameBottom is not null
                ? TakeParagraphStoryLines(bodyStoryLayout.TextLines[0].SourceParagraph!, bodyLineBoxes,
                    bodyLineIndex, remainingLines, capacity, printScale >= 1d ? null : bodyFitAfters,
                    hasPageContent, isFreshHead: bodyLineIndex == 0)
                : TakeStoryLines(bodyLineBoxes, bodyLineIndex, remainingLines, capacity, printScale >= 1d ? null : bodyFitAfters);
            double contentTop = bottom + takeHeight;
            markBaselineY = contentTop + separatorGapPoints;
        }
        else
        {
            double firstInsetPoints = Math.Max(0d, -continuationSeparatorLayout.TextLines[0].BaselineY);
            markBaselineY = cursorTop - firstInsetPoints;
        }

        (DocxPlacedRelatedStoryLayout placedSeparator, _) = PlaceContinuationSeparatorStory(
            activePage, activePageIndex, continuationSeparatorLayout, sourceBlockIndex, markBaselineY, separatorMeasurer, printScale);
        activePlacedStories.Add(placedSeparator);
        outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
        return markBaselineY - separatorGapPoints;
    }

    private static double ResolveContinuationRuleWidth(DocxLayoutPage page, double printScale)
    {
        double reservedMarginPoints = Math.Abs(printScale - 1d) < 0.000000001d ? 0d : page.MarkupMarginReservePoints;
        return Math.Max(1d, page.Width - page.MarginLeft - page.MarginRight + reservedMarginPoints);
    }

    private static (DocxPlacedRelatedStoryLayout Placed, double SeparatorBottom) PlaceContinuationSeparatorStory(
        DocxLayoutPage page,
        int pageIndex,
        DocxRelatedStoryLayout continuationLayout,
        int sourceBlockIndex,
        double markBaselineY,
        IDocxTextMeasurer? separatorMeasurer,
        double printScale = 1d)
    {
        (DocxTextRun? markRun, double markFontSizePoints) = FindSeparatorMarkFont(continuationLayout.TextLines);
        double firstInsetPoints = continuationLayout.TextLines.Count == 0 ? 0d : Math.Max(0d, -continuationLayout.TextLines[0].BaselineY);
        double topY = markBaselineY + firstInsetPoints;
        (double ruleBottomOffsetPoints, double ruleThicknessPoints) = ResolveSeparatorRuleGeometry(continuationLayout.Story.Kind, markRun, markFontSizePoints, separatorMeasurer);
        // RV06 wclong probe (Word 16.0): word-compatible continuation rules span the
        // full design body mapped once (Office 354.98 emitted). The layout body is
        // shrunk by the markup reserve for break equivalence, so the reserve joins
        // the rule width back on print-scaled pages; reserve-margin pages keep the
        // shrunk body with identity emission, and unit scales keep legacy behavior
        // bit-identically. The story widens with the rule so the emission clip
        // covers it.
        double ruleWidthPoints = ResolveContinuationRuleWidth(page, printScale);
        DocxPlacedRelatedStoryLayout placedSeparator = PlaceRelatedStoryAtTop(page, pageIndex, continuationLayout, sourceBlockIndex, topY, storyTopOffset: 0d, storyHeight: ResolvePlacedStoryHeight(continuationLayout, page), separatorY: markBaselineY + ruleBottomOffsetPoints);
        placedSeparator = placedSeparator with
        {
            SeparatorThickness = ruleThicknessPoints,
            SeparatorWidth = ruleWidthPoints,
            Width = Math.Max(placedSeparator.Width, ruleWidthPoints),
        };
        if (placedSeparator.TextLines.Count == 1 &&
            placedSeparator.TextLines[0].Segments.Count == 1 &&
            string.IsNullOrWhiteSpace(placedSeparator.TextLines[0].Segments[0].Text))
        {
            DocxTextLineLayout separatorLine = placedSeparator.TextLines[0];
            DocxTextSegmentLayout markSegment = separatorLine.Segments[0];
            // The end mark sits at the rule end like Office (wclong mark at 409.58 =
            // rule end); without a reserve the width equals the story width, so legacy
            // placement is unchanged.
            double markX = placedSeparator.X + ruleWidthPoints;
            placedSeparator = placedSeparator with
            {
                TextLines = [separatorLine with { Segments = [markSegment with { X = markX }] }],
            };
        }

        return (placedSeparator, markBaselineY);
    }

    private static void MoveToRelatedStoryContinuationPage(
        List<DocxLayoutPage> outputPages,
        ref int activePageIndex,
        ref DocxLayoutPage activePage,
        ref List<DocxPlacedRelatedStoryLayout> activePlacedStories,
        ref double cursorTop,
        bool insertContinuationAfterActivePage, 
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null)
    {
        DocxLayoutPage templatePage = activePage;
        activePage = CreateEmptyContinuationPage(templatePage);
        if (insertContinuationAfterActivePage)
        {
            activePageIndex++;
            outputPages.Insert(activePageIndex, activePage);
        }
        else
        {
            outputPages.Add(activePage);
            activePageIndex = outputPages.Count - 1;
        }

        activePlacedStories = [];
        // RV06 wclong header probes: continuation pages start below the static
        // header zone instead of the full page top (absent without headers, so
        // header-free documents keep legacy takes bit-identically).
        double pageTop = activePage.Height - activePage.MarginTop;
        double keepOutBottom = headerKeepOut?.Invoke(templatePage, activePageIndex + 1, outputPages.Count) ?? double.PositiveInfinity;
        cursorTop = Math.Min(pageTop, keepOutBottom);

        DocxLayoutPage CreateEmptyContinuationPage(DocxLayoutPage template)
        {
            return template with
            {
                StaticTextLines = [],
                StaticInlineImages = [],
                StaticTableRows = [],
                PlacedRelatedStories = [],
                Items = []
            };
        }
    }

    // R12: internal for index equivalence tests; constructed during layout only.
    internal sealed record DocxInlineReferenceLocation(
        int SourceBlockIndex,
        DocxParagraph SourceParagraph,
        DocxInlineReference Reference);

    // RV06 interleaving: in-flight footnote drain state shared by the body loop hook
    // (shared placement onto body pages) and the post-pass tail (dedicated remainder
    // drain). The placed-line counter accumulates takes; the separator flag excludes
    // the head rule from future-page remainders once placed.
    internal sealed class InFlightRelatedStory
    {
        public InFlightRelatedStory(
            DocxRelatedStoryLayout storyLayout,
            DocxRelatedStoryLayout? separatorLayout,
            DocxRelatedStoryLayout? continuationLayout,
            int sourceBlockIndex,
            DocxInlineReferenceLocation location)
        {
            StoryLayout = storyLayout;
            SeparatorLayout = separatorLayout;
            ContinuationLayout = continuationLayout;
            SourceBlockIndex = sourceBlockIndex;
            Location = location;
        }

        public DocxRelatedStoryLayout StoryLayout { get; }
        public DocxRelatedStoryLayout? SeparatorLayout { get; }
        public DocxRelatedStoryLayout? ContinuationLayout { get; }
        public int SourceBlockIndex { get; }
        public DocxInlineReferenceLocation Location { get; }
        public int PlacedLineCount { get; set; }
        public bool SeparatorPlaced { get; set; }

        public int RemainingLineCount => StoryLayout.TextLines.Count - PlacedLineCount;

        public double RemainingContentHeight(IReadOnlyList<double> lineBoxes)
        {
            double remaining = 0d;
            for (int index = PlacedLineCount; index < lineBoxes.Count; index++)
            {
                remaining += lineBoxes[index];
            }

            return Math.Max(0d, remaining);
        }
    }

    private sealed record DocxReferencedRelatedStoryLayout(
        DocxInlineReferenceLocation Location,
        DocxRelatedStoryLayout StoryLayout);

    private readonly record struct DocxPageTextLineOwner(
        DocxTextLineLayout Line,
        int? SourceBlockIndex);
    // Word lays separator stories compact around the rule and space instead of a full line box, so the
    // blank baseline rides at the story bottom (see constants in DocxLayout); shifting is rigid across lines.
    // RV06 endnote mark probes (Word COM references, ten families plus four mark sizes): Office footnote
    // and endnote marks both sit at the separator block bottom (ride 0), so the rule-mark gap
    // carries the full strikeout offset on both note kinds.
    private static DocxPlacedRelatedStoryLayout ShiftSeparatorStoryToBaseline(DocxPlacedRelatedStoryLayout placed, double bottom)
    {
        if (placed.TextLines.Count == 0)
        {
            return placed;
        }

        double ridePoints = 0d;
        double currentBaseline = placed.TextLines[^1].BaselineY;
        return placed with { TextLines = ShiftTextLines(placed.TextLines, bottom + ridePoints - currentBaseline, 0d) };
    }

    // RV06 footnote/endnote probes: both draw the separator rule with a mark space at the
    // rule end. Geometry is shared by footnote, section-end endnote and document-end
    // endnote placement; the rule follows OS/2 strikeout geometry below, with the legacy
    // footnote (2.1pt) and endnote (3.74pt) offsets kept for measurers without font metrics.
    private static (double RuleBottomOffsetPoints, double RuleThicknessPoints) ResolveSeparatorRuleGeometry(
        DocxRelatedStoryKind kind,
        DocxTextRun? markRun,
        double markFontSizePoints,
        IDocxTextMeasurer? separatorMeasurer)
    {
        double legacyOffsetPoints = kind == DocxRelatedStoryKind.Endnote
            ? EndnoteSeparatorRuleBottomOffsetPoints
            : FootnoteSeparatorRuleBottomOffsetPoints;
        // RV06 rule-thickness probes (Word 16.0, Times grids at 8 slash 10 slash 12pt plus
        // Tahoma slash Calibri 14pt grids): Office rule thickness snaps OS slash 2 strikeout size
        // to whole 600dpi pixels (0.36 slash 0.48 slash 0.60 against 0.398 slash 0.498 slash 0.598;
        // 0.72 slash 0.84 slash 0.96 against 0.697 slash 0.889 slash 0.916, all seven exact).
        const double rulePixelPoints = 72d / 600d;
        if (separatorMeasurer is not null &&
            markFontSizePoints > 0d &&
            separatorMeasurer.TryGetStrikeoutRuleMetrics(markRun, out double positionEm, out double thicknessEm) &&
            positionEm > 0d &&
            thicknessEm > 0d)
        {
            double thicknessPoints = System.Math.Round(thicknessEm * markFontSizePoints / rulePixelPoints, System.MidpointRounding.AwayFromZero) * rulePixelPoints;
            return (Math.Max(0d, positionEm * markFontSizePoints - thicknessPoints), thicknessPoints);
        }

        return (legacyOffsetPoints, FootnoteSeparatorThicknessPoints);
    }

    // RV06 separator-bottom probes (Word 16.0, Times, Aptos and Calibri at 10, 12 and 14pt,
    // exact-24 bodies, split-font runs): the Office footnote gap above the body equals one
    // single-spaced line box minus the first-baseline inset, so it derives from the mark
    // font through the measurer. Measurers without single-line metrics keep the legacy
    // 3pt footnote gap.
    private static double ResolveSeparatorGapPoints(
        DocxTextRun? markRun,
        double markFontSizePoints,
        IDocxTextMeasurer? separatorMeasurer)
    {
        if (separatorMeasurer is not null &&
            markFontSizePoints > 0d &&
            separatorMeasurer.TryGetSingleLineEm(markRun, out double singleLineEm) &&
            singleLineEm > DocxLineMetrics.WordAutoLineBaselineOffsetEm)
        {
            return (singleLineEm - DocxLineMetrics.WordAutoLineBaselineOffsetEm) * markFontSizePoints;
        }

        return FootnoteSeparatorGapPoints;
    }

    private static (DocxTextRun? MarkRun, double MarkFontSizePoints) FindSeparatorMarkFont(IReadOnlyList<DocxTextLineLayout> textLines)
    {
        foreach (DocxTextLineLayout line in textLines)
        {
            if (line.Segments.Count != 0)
            {
                DocxTextRun markRun = line.Segments[0].StyleRun;
                return (markRun, markRun.EffectiveProperties.FontSize);
            }
        }

        return (null, 0d);
    }

    // RV06 endnote height-model probes (Word COM references edge-endsepheight-* plus the
    // 10/12/14pt mark grids): Office separator placement ignores laid-out line boxes
    // (empty/exact-24/text variants place identically), growing only by the first-line
    // inset with mark size. The slope pivot preserves validated 12pt values (grid center)
    // while replacing the laid-out box slope with the inset slope.
    private const double EndnoteSeparatorSlopePivotPoints = 12d;

    // Shared by the document-end overflow turn below: the fit check must measure the
    // same size-driven height that placement will consume, not the laid-out height.
    private static double ResolveSizeDrivenSeparatorHeight(DocxRelatedStoryLayout separatorLayout, DocxLayoutPage page, IDocxTextMeasurer? separatorMeasurer, IReadOnlyList<DocxRelatedStoryLayout>? contentStoriesForHhea = null)
    {
        double laidOutHeight = ResolvePlacedStoryHeight(separatorLayout, page);
        if (separatorLayout.TextLines.Count == 0)
        {
            return laidOutHeight;
        }
        (DocxTextRun? markRun, double pivotMarkSize) = FindSeparatorMarkFont(separatorLayout.TextLines);
        if (pivotMarkSize <= 0d)
        {
            return laidOutHeight;
        }
        DocxTextLineLayout firstSeparatorLine = separatorLayout.TextLines[0];
        double firstLineBox = firstSeparatorLine.LineHeight ?? firstSeparatorLine.SingleLineHeight ?? firstSeparatorLine.FontSize;
        double firstInset = DocxLineMetrics.WordAutoLineBaselineOffsetEm * pivotMarkSize;
        double lineBoxAtPivot = firstLineBox * EndnoteSeparatorSlopePivotPoints / pivotMarkSize;
        double insetAtPivot = DocxLineMetrics.WordAutoLineBaselineOffsetEm * EndnoteSeparatorSlopePivotPoints;
        double singleLineDeficitCorrection = ResolveSeparatorSingleLineDeficitCorrection(markRun, pivotMarkSize, separatorMeasurer);
        double smallAscSupplement = ResolveSeparatorSmallAscSupplement(markRun, pivotMarkSize, separatorMeasurer);
        double hheaAscenderExcess = System.Math.Min(ResolveSeparatorHheaAscenderExcess(markRun, pivotMarkSize, separatorMeasurer), ResolveContentHheaAscenderExcess(contentStoriesForHhea, separatorMeasurer));
        return Math.Max(0d, laidOutHeight - firstLineBox + firstInset + (lineBoxAtPivot - insetAtPivot) + singleLineDeficitCorrection + smallAscSupplement + hheaAscenderExcess);
    }

    // RV06 endnote block-shift probes (Word COM references: five-family mark axis on pinned Tahoma content plus Tahoma-mark size sweep, genuine embeds both sides): Office document-end separator totals run about 0.8 taller on Tahoma/Verdana marks, flat across 8/10/14pt marks, while Calibri/Times/Georgia/Arial and the other probed families sit inside 0.25. The excess keys on hhea ascender over the auto-inset floor the layout already assumes. Gated on hhea metrics; measurers without them keep byte-identical behavior.
    private const double EndnoteSeparatorHheaAscenderExcessSlope = 1.27d;

    private static double ResolveSeparatorHheaAscenderExcess(DocxTextRun? markRun, double pivotMarkSize, IDocxTextMeasurer? separatorMeasurer)
    {
        if (pivotMarkSize <= 0d || separatorMeasurer is not IDocxLineMetricsProvider metrics)
        {
            return 0d;
        }

        double hheaAscenderEm = metrics.MeasureHheaAscender(markRun, pivotMarkSize) / pivotMarkSize;
        if (!(hheaAscenderEm > DocxLineMetrics.WordAutoLineBaselineOffsetEm))
        {
            return 0d;
        }

        return (hheaAscenderEm - DocxLineMetrics.WordAutoLineBaselineOffsetEm) * EndnoteSeparatorHheaAscenderExcessSlope * pivotMarkSize;
    }

    // The excess needs both sides Tahoma-like (Calibri on either side of the mark/content pair holds the validated level while Tahoma on both raises it about 0.8): the caller meets the mark-side term above with the content-side term below in the minimum, so single-sided Tahoma keeps the legacy level.
    // Flow uses each source run's own size. Rule painting expresses the same
    // source font bearing at its mark size, without changing the flow charge.
    private static double ResolveContentHheaAscenderExcess(IReadOnlyList<DocxRelatedStoryLayout>? contentStories,
        IDocxTextMeasurer? separatorMeasurer, double? referenceFontSizePoints = null)
    {
        if (contentStories is null || contentStories.Count == 0 || separatorMeasurer is not IDocxLineMetricsProvider metrics)
        {
            return double.PositiveInfinity;
        }

        double excessMax = 0d;
        bool hasRuns = false;
        foreach (DocxRelatedStoryLayout story in contentStories)
        {
            int firstParagraphIndex = int.MaxValue;
            foreach (DocxTextLineLayout line in story.TextLines)
            {
                if (line.SourceParagraphIndex.HasValue && line.SourceParagraphIndex.Value < firstParagraphIndex)
                {
                    firstParagraphIndex = line.SourceParagraphIndex.Value;
                }
            }
            foreach (DocxTextLineLayout line in story.TextLines)
            {
                if (line.SourceParagraphIndex.HasValue && line.SourceParagraphIndex.Value != firstParagraphIndex)
                {
                    continue;
                }
                foreach (DocxTextSegmentLayout segment in line.Segments)
                {
                    if (string.IsNullOrWhiteSpace(segment.Text))
                    {
                        continue;
                    }
                    DocxTextRun run = segment.StyleRun;
                    double runSize = run.EffectiveProperties.FontSize;
                    if (!(runSize > 0d))
                    {
                        continue;
                    }
                    double runExcessEm = metrics.MeasureHheaAscender(run, runSize) / runSize - DocxLineMetrics.WordAutoLineBaselineOffsetEm;
                    hasRuns = true;
                    if (runExcessEm > 0d)
                    {
                        excessMax = System.Math.Max(excessMax, runExcessEm * EndnoteSeparatorHheaAscenderExcessSlope * (referenceFontSizePoints ?? runSize));
                    }
                }
            }
        }

        return hasRuns ? excessMax : double.PositiveInfinity;
    }
    // RV06 four-family separator grids (Word COM references edge-endsepgrid-cal/tmr/tah/vdn
    // at 10/12/14pt): Office separator height carries a singleLine-deficit slope term against
    // the validated Calibri anchor that no laid-out box spread explains (Times, Tahoma and
    // Verdana share singleLineEm 1.15 yet place about 0.7pt lower at 12pt with matching content
    // boxes, rule thickness, rule offsets and below-gaps). The term vanishes for the validated
    // family and for measurers without single-line metrics, so those paths stay byte-identical.
    private const double EndnoteSeparatorReferenceSingleLineEm = 1.2207d;
    private const double EndnoteSeparatorSingleLineDeficitSlope = 0.84d;

    private static double ResolveSeparatorSingleLineDeficitCorrection(DocxTextRun? markRun, double pivotMarkSize, IDocxTextMeasurer? separatorMeasurer)
    {
        if (pivotMarkSize <= 0d || separatorMeasurer is null || !separatorMeasurer.TryGetSingleLineEm(markRun, out double singleLineEm) || singleLineEm <= 0d)
        {
            return 0d;
        }

        return Math.Max(0d, EndnoteSeparatorReferenceSingleLineEm - singleLineEm) * EndnoteSeparatorSingleLineDeficitSlope * pivotMarkSize;
    }

    private const double EndnoteSeparatorReferenceWindowsAscenderEm = 0.9521d;
    private const double EndnoteSeparatorSmallAscSupplementSlope = 2.0d;

    private static double ResolveSeparatorSmallAscSupplement(DocxTextRun? markRun, double pivotMarkSize, IDocxTextMeasurer? separatorMeasurer)
    {
        if (pivotMarkSize <= 0d || separatorMeasurer is not IDocxStaticTextMetricsProvider staticMetrics)
        {
            return 0d;
        }

        double windowsAscenderPoints = staticMetrics.MeasureWindowsAscender(markRun, pivotMarkSize);
        if (windowsAscenderPoints <= 0d)
        {
            return 0d;
        }

        double windowsAscenderEm = windowsAscenderPoints / pivotMarkSize;
        double ascenderDeficitEm = EndnoteSeparatorReferenceWindowsAscenderEm - windowsAscenderEm;
        if (ascenderDeficitEm <= 0d)
        {
            return 0d;
        }

        double smallSizePoints = EndnoteSeparatorSlopePivotPoints - pivotMarkSize;
        if (smallSizePoints <= 0d)
        {
            return 0d;
        }

        return ascenderDeficitEm * EndnoteSeparatorSmallAscSupplementSlope * smallSizePoints;
    }

    // Source-owned heads consume the separator-height deficit once. The mark
    // gap remains the floor for matching content metrics and metric fallbacks.
    private static double ResolveSourceInsetNoteContentGap(double contentGapPoints,
        DocxTextRun? markRun, double markFontSizePoints, IDocxTextMeasurer? separatorMeasurer) =>
        Math.Max(ResolveSeparatorGapPoints(markRun, markFontSizePoints, separatorMeasurer),
            contentGapPoints - ResolveSeparatorSingleLineDeficitCorrection(markRun, markFontSizePoints, separatorMeasurer));

    private static double ResolveFootnoteContentGapPoints(
        IReadOnlyList<DocxTextLineLayout> contentTextLines,
        DocxTextRun? markRun,
        double markFontSizePoints,
        IDocxTextMeasurer? separatorMeasurer)
    {
        // RV06 mixed-gap probes (Word 16.0, Pal/Cal orders with r2f diff minus 1.20):
        // Office content gap follows the first paragraph, so later paragraphs do not
        // widen the gap; single-paragraph notes resolve identically.
        int firstParagraphIndex = int.MaxValue;
        foreach (DocxTextLineLayout firstParaLine in contentTextLines)
        {
            if (firstParaLine.SourceParagraphIndex.HasValue && firstParaLine.SourceParagraphIndex.Value < firstParagraphIndex)
            {
                firstParagraphIndex = firstParaLine.SourceParagraphIndex.Value;
            }
        }

        double firstInsetPoints = 0d;
        foreach (DocxTextLineLayout firstLine in contentTextLines)
        {
            if (firstLine.SourceParagraphIndex.HasValue && firstLine.SourceParagraphIndex.Value != firstParagraphIndex)
            {
                continue;
            }

            if (firstLine.Segments.Count != 0)
            {
                firstInsetPoints = System.Math.Max(0d, -firstLine.BaselineY);
                break;
            }
        }

        double contentGapMax = 0d;
        foreach (DocxTextLineLayout line in contentTextLines)
        {
            if (line.SourceParagraphIndex.HasValue && line.SourceParagraphIndex.Value != firstParagraphIndex)
            {
                continue;
            }

            foreach (DocxTextSegmentLayout segment in line.Segments)
            {
                if (string.IsNullOrWhiteSpace(segment.Text))
                {
                    continue;
                }

                DocxTextRun run = segment.StyleRun;
                // RV06 endnote mark-size probes (m28/m28c): the inset shelter covers content-proportionally, so it scales with the run size while the mark-size form releases the supplement early on large marks.
                double runFontSizePoints = run.EffectiveProperties.FontSize;
                double runCoveredInsetExcessPoints = runFontSizePoints > 0d
                    ? System.Math.Max(0d, firstInsetPoints - DocxLineMetrics.WordAutoLineBaselineOffsetEm * runFontSizePoints)
                    : System.Math.Max(0d, firstInsetPoints - DocxLineMetrics.WordAutoLineBaselineOffsetEm * markFontSizePoints);
                contentGapMax = System.Math.Max(contentGapMax, ResolveSeparatorGapPoints(run, markFontSizePoints, separatorMeasurer));
                contentGapMax = System.Math.Max(contentGapMax, ResolveWascContentGapSupplement(run, markRun, markFontSizePoints, separatorMeasurer, runCoveredInsetExcessPoints));
            }
        }

        return contentGapMax;
    }

    private static double ResolveFootnoteContentGapPoints(
        IReadOnlyList<DocxReferencedRelatedStoryLayout> footnoteStories,
        DocxTextRun? markRun,
        double markFontSizePoints,
        IDocxTextMeasurer? separatorMeasurer)
    {
        double contentGapMax = 0d;
        foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
        {
            contentGapMax = System.Math.Max(contentGapMax, ResolveFootnoteContentGapPoints(story.StoryLayout.TextLines, markRun, markFontSizePoints, separatorMeasurer));
        }

        return contentGapMax;
    }

    // RV06 fnwrap content-gap probes (Word 16.0): Office hangs big-ascender content
    // lower by the Windows-ascender excess over the mark, so the supplement competes
    // with the single-line gap through the caller max and vanishes without static
    // metrics, keeping those paths byte-identical. The excess yields to first-inset
    // space the layout already placed (Tah-Verdana mixes), so covered insets do not
    // double-count Office single excess.
    // RV06 endnote mark-size probes (m28/m28c): ascents and the inset shelter scale with
    // the run size; the mark-size form released the supplement early on large marks.
    private static double ResolveWascContentGapSupplement(
        DocxTextRun run,
        DocxTextRun? markRun,
        double markFontSizePoints,
        IDocxTextMeasurer? separatorMeasurer,
        double coveredInsetExcessPoints = 0d)
    {
        if (markRun is null || markFontSizePoints <= 0d || separatorMeasurer is not IDocxStaticTextMetricsProvider staticMetrics)
        {
            return 0d;
        }

        // RV06 Aptos probes (Word 16.0, uniform Aptos lacks the excess while Segoe keeps
        // 1.524 of it with Verdana present at 0.64): only Aptos sets OS/2 USE_TYPO_METRICS
        // among probed families, so runs whose resolved face requests typographic metrics
        // skip the Windows-ascender supplement instead of keying it by magnitude.
        if (separatorMeasurer is IDocxTypographicMetricsProvider typographic &&
            typographic.UseTypographicMetrics(run))
        {
            return 0d;
        }

        double runFontSizePoints = run.EffectiveProperties.FontSize > 0d ? run.EffectiveProperties.FontSize : markFontSizePoints;
        double runAscenderPoints = staticMetrics.MeasureWindowsAscender(run, runFontSizePoints);
        double markAscenderPoints = staticMetrics.MeasureWindowsAscender(markRun, runFontSizePoints);
        if (runAscenderPoints <= 0d || markAscenderPoints <= 0d)
        {
            return 0d;
        }

        double excessPoints = System.Math.Max(0d, runAscenderPoints - markAscenderPoints);
        double netExcessPoints = System.Math.Max(0d, excessPoints - coveredInsetExcessPoints);
        if (netExcessPoints <= 0d)
        {
            return 0d;
        }

        return ResolveSeparatorGapPoints(markRun, markFontSizePoints, separatorMeasurer) + netExcessPoints;
    }

    private static (DocxPlacedRelatedStoryLayout Placed, double SeparatorBottom) PlaceSeparatorStoryWithMark(
        DocxLayoutPage page,
        int pageIndex,
        DocxRelatedStoryLayout separatorLayout,
        int sourceBlockIndex,
        double separatorTop,
        IDocxTextMeasurer? separatorMeasurer,
        bool useSizeDrivenPlacementHeight = false,
        IReadOnlyList<DocxRelatedStoryLayout>? contentStoriesForHhea = null,
        double? placementHeightOverride = null,
        double ruleFirstBearingPoints = 0d)
    {
        double separatorHeight = ResolvePlacedStoryHeight(separatorLayout, page);
        if (useSizeDrivenPlacementHeight)
        {
            separatorHeight = ResolveSizeDrivenSeparatorHeight(separatorLayout, page, separatorMeasurer, contentStoriesForHhea);
        }
        separatorHeight = placementHeightOverride ?? separatorHeight;
        double separatorBottom = separatorTop - separatorHeight;
        (DocxTextRun? markRun, double markFontSizePoints) = FindSeparatorMarkFont(separatorLayout.TextLines);
        (double ruleBottomOffsetPoints, double ruleThicknessPoints) = ResolveSeparatorRuleGeometry(separatorLayout.Story.Kind, markRun, markFontSizePoints, separatorMeasurer);
        DocxPlacedRelatedStoryLayout placedSeparator = PlaceRelatedStoryAtTop(page, pageIndex, separatorLayout, sourceBlockIndex, separatorTop, separatorY: separatorBottom + ruleBottomOffsetPoints - ruleFirstBearingPoints);
        if (placementHeightOverride.HasValue)
        {
            placedSeparator = placedSeparator with { Height = separatorHeight };
        }
        placedSeparator = placedSeparator with { SeparatorThickness = ruleThicknessPoints };
        if (useSizeDrivenPlacementHeight)
        {
            placedSeparator = placedSeparator with { Height = separatorHeight };
        }
        placedSeparator = ShiftSeparatorStoryToBaseline(placedSeparator, separatorBottom);
        if (placedSeparator.TextLines.Count == 1 &&
            placedSeparator.TextLines[0].Segments.Count == 1 &&
            string.IsNullOrWhiteSpace(placedSeparator.TextLines[0].Segments[0].Text))
        {
            DocxTextLineLayout separatorLine = placedSeparator.TextLines[0];
            DocxTextSegmentLayout markSegment = separatorLine.Segments[0];
            double markX = placedSeparator.X + Math.Min(FootnoteSeparatorWidthPoints, placedSeparator.Width);
            placedSeparator = placedSeparator with
            {
                TextLines = [separatorLine with { Segments = [markSegment with { X = markX }] }],
            };
        }

        return (placedSeparator, separatorBottom);
    }

    private static DocxPlacedRelatedStoryLayout PlaceRelatedStory(DocxLayoutPage page, int pageIndex, DocxRelatedStoryLayout storyLayout, int sourceBlockIndex)
    {
        double storyHeight = Math.Min(Math.Max(0d, storyLayout.ContentHeight), Math.Max(0d, page.Height - page.MarginTop - page.MarginBottom));
        double topY = page.MarginBottom + storyHeight;
        return PlaceRelatedStoryAtTop(page, pageIndex, storyLayout, sourceBlockIndex, topY, topY + FootnoteSeparatorGapPoints);
    }

    // RV06 footlong probe (Word 16.0): footnote blocks taller than the body-anchored
    // area spill off-page top because the block stacks unbounded from the margin. Clamp
    // the stories top to the body bottom edge minus the separator height; short blocks
    // keep the legacy bottom-anchored stacking bit-identically (the reserve holds body
    // clear otherwise, so the clamp only bites on overflow). Stories that fit place
    // whole as before; the first overflowing story and all later ones slice through
    // the shared line-partition machinery with footnote continuation separators.
    // Absent-separator stories keep legacy whole placement (synthetic rule preserved).
    // RV06 p1-clamp probes (Word 16.0, b/ba/2para/bodyb sweeps): Office storiesTop is
    // independent of body size/after/length/before and follows body start with
    // footnote-side moves (fn0s/fna24). The clamp bottom hangs one first-footnote
    // inset below body start (itself below the header zone) instead of one body em
    // below the lowest body baseline.
    private static double ResolveFootnoteClampBodyBottom(
        DocxLayoutPage page,
        int pageNumber,
        int pageCount,
        DocxRelatedStoryLayout firstStoryLayout,
        Func<DocxLayoutPage, int, int, double>? headerKeepOut)
    {
        double fullTop = page.Height - page.MarginTop;
        double keepOut = headerKeepOut?.Invoke(page, pageNumber, pageCount) ?? double.PositiveInfinity;
        // Mirror the header-displacement rule: body starts at the full top unless
        // header content dips below it, in which case it starts at the header bottom.
        double bodyStart = Math.Min(fullTop, keepOut);
        double firstInset = 0d;
        if (firstStoryLayout.TextLines.Count != 0)
        {
            firstInset = Math.Max(0d, -firstStoryLayout.TextLines[0].BaselineY);
        }

        // RV06 take-battery probes (Word 16.0, after-spacing, size and Palatino takes):
        // Office head takes fit one more line than clamp-minus-separator allows on Calibri
        // uniform notes, and the 9pt boundary needs total trim above 3.73 while the after-12
        // boundary tolerates total trim up to 4.22 with cal-pal to 4.44 and 10pt to 5.62, so the
        // clamp carries a uniform 4.0pt Office-fitted reserve shaved from the take side only;
        // fitting stories keep cursor-driven placement bit-identically.
        return bodyStart - firstInset - 4.0;
    }

    // RV06 interleaving: narrowed remainder view for in-flight continuation. The take
    // loop consumes design line boxes, so the narrowed content height sums the same
    // boxes the takes fit, keeping whole-fit decisions exact. Kept lines re-base at
    // the story origin by the consumed box height, so later takes place from a zeroed
    // offset exactly as if the head lines were never laid out.
    internal static DocxRelatedStoryLayout NarrowStoryTextLinesForOffset(DocxRelatedStoryLayout storyLayout, int startLineIndex)
    {
        if (startLineIndex <= 0)
        {
            return storyLayout;
        }

        DocxTextLineLayout[] sliceLines = new DocxTextLineLayout[storyLayout.TextLines.Count - startLineIndex];
        for (int lineIndex = 0; lineIndex < sliceLines.Length; lineIndex++)
        {
            sliceLines[lineIndex] = storyLayout.TextLines[startLineIndex + lineIndex];
        }

        double[] lineBoxes = GetStoryTextLineBoxHeights(storyLayout);
        double consumedHeight = 0d;
        double remainingHeight = 0d;
        for (int lineIndex = 0; lineIndex < lineBoxes.Length; lineIndex++)
        {
            if (lineIndex < startLineIndex)
            {
                consumedHeight += lineBoxes[lineIndex];
            }
            else
            {
                remainingHeight += lineBoxes[lineIndex];
            }
        }

        return storyLayout with
        {
            TextLines = ShiftTextLines(sliceLines, consumedHeight, 0d),
            InlineImages = ShiftInlineImages(storyLayout.InlineImages, consumedHeight, 0d),
            TableRows = ShiftTableRows(storyLayout.TableRows, consumedHeight, 0d),
            ContentHeight = Math.Max(0d, remainingHeight)
        };
    }

    private static double ResolveSeatedFootnoteTop(DocxLayoutPage page,
        IReadOnlyList<DocxReferencedRelatedStoryLayout> notes, double top, double bottom,
        Dictionary<(DocxRelatedStoryKind Kind, string Id), InFlightRelatedStory>? inFlightNotes,
        double printScale, bool usesParagraphFrames)
    {
        double seated = 0d;
        foreach (var head in usesParagraphFrames ? notes : notes.Take(1))
        {
            var layout = head.StoryLayout;
            if (layout.TextLines.Count == 0) break;
            double[] boxes = GetStoryTextLineBoxHeights(layout);
            double[]? afters = printScale >= 1d ? GetStoryLineAfterSpacings(layout) : null;
            double capacity = Math.Max(0d, top - bottom - seated);
            bool fresh = inFlightNotes is null || head.Location.Reference.Id is null ||
                !inFlightNotes.TryGetValue((head.Location.Reference.Kind, head.Location.Reference.Id), out var flight) ||
                flight.PlacedLineCount == 0;
            (int take, double height) = usesParagraphFrames
                ? TakeParagraphStoryLines(layout.TextLines[0].SourceParagraph!, boxes, 0, layout.TextLines.Count,
                    capacity, afters, page.Items.Count != 0 || seated != 0d, fresh)
                : TakeStoryLines(boxes, 0, layout.TextLines.Count, capacity, afters, forceProgress: seated == 0d);
            // A continued tail and a fresh head share one bottom-anchored batch.
            if (take < layout.TextLines.Count) return bottom + seated + height;
            seated += ResolvePlacedStoryHeight(layout, page);
        }
        return top;
    }

    private static void PlaceFootnoteStories(
        List<DocxLayoutPage> outputPages,
        int outputIndex,
        DocxLayoutPage page,
        IReadOnlyList<DocxReferencedRelatedStoryLayout> footnoteStories,
        DocxRelatedStoryLayout? separatorLayout,
        DocxRelatedStoryLayout? continuationSeparatorLayout,
        IDocxTextMeasurer? separatorMeasurer,
        double printScale = 1d, 
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null,
        bool sharedSlicePlacement = false,
        Dictionary<(DocxRelatedStoryKind Kind, string Id), InFlightRelatedStory>? inFlightNotes = null,
        int? pageNumberOverride = null,
        int? pageCountOverride = null,
        double? sharedContentTop = null,
        Func<DocxLayoutPage, int, int, double>? footerKeepOut = null,
        Func<DocxLayoutPage, int, int, double>? availableFooterKeepOut = null)
    {
        if (footnoteStories.Count == 0)
        {
            return;
        }

        if (separatorLayout is null)
        {
            PlaceFootnoteStoriesWithoutSeparator(outputPages, outputIndex, page, footnoteStories, continuationSeparatorLayout, separatorMeasurer, printScale, headerKeepOut, sharedSlicePlacement, inFlightNotes, pageNumberOverride, pageCountOverride,
                sharedContentTop: sharedContentTop, footerKeepOut: footerKeepOut, availableFooterKeepOut: availableFooterKeepOut);
            return;
        }

        // Normalize once, including when the separator-less path dispatches.
        ResolveOwnedInlineFootnoteFrame(page, footnoteStories, printScale,
            ref sharedContentTop, ref footerKeepOut, availableFooterKeepOut);

        DocxLayoutPage activePage = outputPages[outputIndex];
        int activePageIndex = outputIndex;
        List<DocxPlacedRelatedStoryLayout> activePlacedStories = activePage.PlacedRelatedStories.ToList();
        double bodyHeight = footnoteStories.Sum(story => ResolvePlacedStoryHeight(story.StoryLayout, activePage));
        var placementFooter = ResolveFootnotePlacementFooter(activePage, pageNumberOverride ?? outputIndex + 1,
            pageCountOverride ?? outputPages.Count, footerKeepOut, availableFooterKeepOut);
        double frameBottom = placementFooter?.Invoke(activePage, pageNumberOverride ?? (outputIndex + 1), pageCountOverride ?? outputPages.Count) ?? activePage.MarginBottom;
        double cursorTop = frameBottom + bodyHeight;
        bool continuesHead = inFlightNotes is not null && footnoteStories[0].Location.Reference.Id is { } firstId &&
            inFlightNotes.TryGetValue((DocxRelatedStoryKind.Footnote, firstId), out InFlightRelatedStory? firstFlight) && firstFlight.SeparatorPlaced;
        bool plainSharedHead = footerKeepOut is not null && continuesHead;
        double separatorHeight = frameBottom > activePage.MarginBottom || plainSharedHead
            ? ResolvePlainNoteSeparatorHeight(separatorLayout, activePage, separatorMeasurer,
                footnoteStories.Select(s => s.StoryLayout).ToArray(), printScale)
            : ResolvePlacedStoryHeight(separatorLayout, activePage);
        (DocxTextRun? gapRun, double gapFontSizePoints) = FindSeparatorMarkFont(separatorLayout.TextLines);
        double separatorGapPoints = ResolveSeparatorGapPoints(gapRun, gapFontSizePoints, separatorMeasurer);
        double footnoteContentGapPoints = ResolveFootnoteContentGapPoints(footnoteStories, gapRun, gapFontSizePoints, separatorMeasurer);
        // The physical frame checks the original document sources. An occupied
        // continuation uses the same nominal separator gap as a fresh head;
        // empty pages and unqualified sources retain their existing fallback.
        if (footerKeepOut is not null && CanUseNominalNoteSeparatorFlow(activePage,
            separatorLayout, printScale, footnoteStories.Select(s => s.StoryLayout).ToArray()))
            footnoteContentGapPoints = ResolveSourceInsetNoteContentGap(footnoteContentGapPoints,
                gapRun, gapFontSizePoints, separatorMeasurer);
        double footnoteContentPushdownPoints = System.Math.Max(0d, footnoteContentGapPoints - separatorGapPoints);
        double storiesTop = cursorTop;
        double clampBodyBottom = ResolveFootnoteClampBodyBottom(page, pageNumberOverride ?? (outputIndex + 1), pageCountOverride ?? outputPages.Count, footnoteStories[0].StoryLayout, headerKeepOut);
        // The admitted shared frame already reserves the separator's flow charge.
        // Its physical clip can retain latent spacing without charging it again.
        double separatorCharge = footerKeepOut is not null
            ? ResolvePlainNoteSeparatorHeight(separatorLayout, activePage, separatorMeasurer,
                footnoteStories.Select(s => s.StoryLayout).ToArray(), printScale) : separatorHeight;
        storiesTop = Math.Min(cursorTop, clampBodyBottom - separatorCharge);
        if (sharedContentTop is not null)
        {
            storiesTop = Math.Min(storiesTop, sharedContentTop.Value);
        }
        // RV06 p1-position probes (Word 16.0): Office seats the overflowing head block
        // bottom-up (storiesTop = margin + takeHeight), so absolute positions match instead
        // of leaving capacity slack above the margin. The head take still comes from the
        // same clamped capacity, so take counts are unchanged; fitting stories keep the
        // legacy whole placement bit-identically.
        storiesTop = ResolveSeatedFootnoteTop(activePage, footnoteStories, storiesTop, frameBottom,
            inFlightNotes, printScale, footerKeepOut is not null);
        bool sharesContinuationSeparator = plainSharedHead && continuationSeparatorLayout is { TextLines.Count: > 0 };
        bool drawHeadSeparator = !sharesContinuationSeparator;
        if (inFlightNotes is not null && !sharesContinuationSeparator)
        {
            drawHeadSeparator = false;
            foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
            {
                if (story.Location.Reference.Id is null || !inFlightNotes.TryGetValue((story.Location.Reference.Kind, story.Location.Reference.Id), out InFlightRelatedStory? tracked) || !tracked.SeparatorPlaced)
                {
                    drawHeadSeparator = true;
                    break;
                }
            }

        }

        if (drawHeadSeparator)
        {
        double separatorTop = storiesTop + separatorGapPoints + separatorHeight;
        (DocxPlacedRelatedStoryLayout placedSeparator, _) = PlaceSeparatorStoryWithMark(activePage, activePageIndex, separatorLayout,
            footnoteStories[0].Location.SourceBlockIndex, separatorTop, separatorMeasurer,
            placementHeightOverride: placementFooter is not null ? separatorHeight : null);
        activePlacedStories.Add(placedSeparator);
        outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
        }
        // A plain footer-frame batch owns one physical separator, matching its
        // shared reserve. A continued first note supplies the full-width rule.
        bool sharedSeparatorPlaced = drawHeadSeparator;
        double contentTop = storiesTop;
        foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
        {
            InFlightRelatedStory? trackedStory = inFlightNotes is not null && story.Location.Reference.Id is not null && inFlightNotes.TryGetValue((story.Location.Reference.Kind, story.Location.Reference.Id), out InFlightRelatedStory? tracked) ? tracked : null;
            double storyHeight = ResolvePlacedStoryHeight(story.StoryLayout, activePage);
            if (trackedStory is not null && trackedStory.SeparatorPlaced && continuationSeparatorLayout is not null && continuationSeparatorLayout.TextLines.Count != 0 &&
                (footerKeepOut is null || !sharedSeparatorPlaced))
            {
                (DocxTextRun? continuationGapRun, double continuationGapFontSizePoints) = FindSeparatorMarkFont(continuationSeparatorLayout.TextLines);
                double continuationGapPoints = ResolveSeparatorGapPoints(continuationGapRun, continuationGapFontSizePoints, separatorMeasurer);
                (DocxPlacedRelatedStoryLayout placedContinuation, _) = PlaceContinuationSeparatorStory(activePage, activePageIndex, continuationSeparatorLayout, story.Location.SourceBlockIndex, contentTop + continuationGapPoints, separatorMeasurer, printScale);
                activePlacedStories.Add(placedContinuation);
                sharedSeparatorPlaced = true;
                outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
            }

            if (contentTop - storyHeight >= frameBottom - 0.001d)
            {
                DocxPlacedRelatedStoryLayout placedStory = PlaceRelatedStoryAtTop(activePage, activePageIndex, story.StoryLayout, story.Location.SourceBlockIndex, contentTop - footnoteContentPushdownPoints, separatorY: null)
                    with { UsesPlainFootnoteFrame = footerKeepOut is not null };
                activePlacedStories.Add(placedStory);
                outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
                contentTop -= storyHeight;
                if (trackedStory is not null)
                {
                    trackedStory.PlacedLineCount += story.StoryLayout.TextLines.Count;
                }
            }
            else
            {
                PlaceRelatedStorySlices(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref contentTop, story.StoryLayout, story.Location.SourceBlockIndex, insertContinuationAfterActivePage: true, continuationSeparatorLayout, separatorMeasurer, printScale, headerKeepOut, stopAfterCurrentPage: sharedSlicePlacement, inFlight: trackedStory, headContentPushdownPoints: footnoteContentPushdownPoints, frameBottom: placementFooter, usesPlainFootnoteFrame: footerKeepOut is not null);
            }
        }

        if (inFlightNotes is not null)
        {
            foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
            {
                if (story.Location.Reference.Id is not null && inFlightNotes.TryGetValue((story.Location.Reference.Kind, story.Location.Reference.Id), out InFlightRelatedStory? fresh) && !fresh.SeparatorPlaced)
                {
                    fresh.SeparatorPlaced = true;
                }
            }
        }
    }

    // RV06 footlong probes: separator-less footnote blocks keep legacy whole
    // placement when they fit (synthetic rule preserved), but overflowing stories
    // slice through the shared line-partition machinery with the synthetic rule
    // patched onto the head slice. The stories top clamps like the present path,
    // using the first body line box for the missing separator height.
    private static void PlaceFootnoteStoriesWithoutSeparator(
        List<DocxLayoutPage> outputPages,
        int outputIndex,
        DocxLayoutPage page,
        IReadOnlyList<DocxReferencedRelatedStoryLayout> footnoteStories,
        DocxRelatedStoryLayout? continuationSeparatorLayout,
        IDocxTextMeasurer? separatorMeasurer,
        double printScale = 1d, 
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null,
        bool sharedSlicePlacement = false,
        Dictionary<(DocxRelatedStoryKind Kind, string Id), InFlightRelatedStory>? inFlightNotes = null,
        int? pageNumberOverride = null,
        int? pageCountOverride = null,
        double? sharedContentTop = null,
        Func<DocxLayoutPage, int, int, double>? footerKeepOut = null,
        Func<DocxLayoutPage, int, int, double>? availableFooterKeepOut = null)
    {
        bool ownedFrame = ResolveOwnedInlineFootnoteFrame(page, footnoteStories, printScale,
            ref sharedContentTop, ref footerKeepOut, availableFooterKeepOut);
        DocxLayoutPage activePage = outputPages[outputIndex];
        int activePageIndex = outputIndex;
        List<DocxPlacedRelatedStoryLayout> activePlacedStories = activePage.PlacedRelatedStories.ToList();
        double bodyHeight = footnoteStories.Sum(story => ResolvePlacedStoryHeight(story.StoryLayout, activePage));
        var placementFooter = ResolveFootnotePlacementFooter(activePage, pageNumberOverride ?? outputIndex + 1,
            pageCountOverride ?? outputPages.Count, footerKeepOut, availableFooterKeepOut);
        double frameBottom = placementFooter?.Invoke(activePage, pageNumberOverride ?? (outputIndex + 1), pageCountOverride ?? outputPages.Count) ?? activePage.MarginBottom;
        double cursorTop = frameBottom + bodyHeight;
        double storiesTop = cursorTop;
        DocxRelatedStoryLayout firstLayout = footnoteStories[0].StoryLayout;
        double clampBodyBottom = ownedFrame
            ? Math.Min(page.Height - page.MarginTop,
                headerKeepOut?.Invoke(page, pageNumberOverride ?? (outputIndex + 1), pageCountOverride ?? outputPages.Count) ?? double.PositiveInfinity)
            : ResolveFootnoteClampBodyBottom(page, pageNumberOverride ?? (outputIndex + 1), pageCountOverride ?? outputPages.Count, firstLayout, headerKeepOut);
        if (firstLayout.TextLines.Count != 0)
        {
            double firstBoxHeight = GetStoryTextLineBoxHeights(firstLayout)[0];
            if (ownedFrame && separatorMeasurer is IDocxLineMetricsProvider && firstLayout.TextLines[0].SourceParagraph is { } paragraph)
            {
                // A synthetic separator consumes a natural font line, independent
                // of authored paragraph minima and inline mathematical extents.
                var natural = paragraph with
                {
                    LineSpacingPoints = null,
                    LineSpacingFactor = 1d,
                    Spacing = paragraph.Spacing with { LineRuleValue = "auto", LineValue = "240" }
                };
                double height = ResolveNoteContentLineHeightProfile(natural, GetParagraphFontSize(natural), separatorMeasurer).LineHeight;
                if (double.IsFinite(height) && height > 0d) firstBoxHeight = height;
            }
            storiesTop = Math.Min(cursorTop, clampBodyBottom - firstBoxHeight);
            if (sharedContentTop is not null)
            {
                storiesTop = Math.Min(storiesTop, sharedContentTop.Value - (ownedFrame ? firstBoxHeight : 0d));
            }
        }
        if (ownedFrame)
        {
            storiesTop = ResolveSeatedFootnoteTop(activePage, footnoteStories, storiesTop, frameBottom,
                inFlightNotes, printScale, usesParagraphFrames: true);
        }
        double contentTop = storiesTop;
        bool firstStory = true;
        foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
        {
            InFlightRelatedStory? trackedStory = inFlightNotes is not null && story.Location.Reference.Id is not null && inFlightNotes.TryGetValue((story.Location.Reference.Kind, story.Location.Reference.Id), out InFlightRelatedStory? tracked) ? tracked : null;
            double storyHeight = ResolvePlacedStoryHeight(story.StoryLayout, activePage);
            // Omitted continuation stories still have a default full-width rule.
            // Source progress distinguishes a true continuation from an unplaced head.
            double? defaultContinuationWidth = ownedFrame && continuationSeparatorLayout is null &&
                trackedStory is { PlacedLineCount: > 0 }
                    ? ResolveContinuationRuleWidth(activePage, printScale) : null;
            bool needsContinuationRule = false;
            if (trackedStory is not null && trackedStory.SeparatorPlaced && continuationSeparatorLayout is not null && continuationSeparatorLayout.TextLines.Count != 0)
            {
                needsContinuationRule = true;
                (DocxTextRun? continuationGapRun, double continuationGapFontSizePoints) = FindSeparatorMarkFont(continuationSeparatorLayout.TextLines);
                double continuationGapPoints = ResolveSeparatorGapPoints(continuationGapRun, continuationGapFontSizePoints, separatorMeasurer);
                (DocxPlacedRelatedStoryLayout placedContinuation, _) = PlaceContinuationSeparatorStory(activePage, activePageIndex, continuationSeparatorLayout, story.Location.SourceBlockIndex, contentTop + continuationGapPoints, separatorMeasurer, printScale);
                activePlacedStories.Add(placedContinuation);
                outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
            }

            if (contentTop - storyHeight >= frameBottom - 0.001d)
            {
                double? separatorY = null;
                double separatorThickness = FootnoteSeparatorThicknessPoints;
                if (firstStory && !needsContinuationRule)
                {
                    (DocxTextRun? bodyRun, double bodyFontSizePoints) = FindSeparatorMarkFont(story.StoryLayout.TextLines);
                    (double syntheticOffsetPoints, double syntheticThicknessPoints) = ResolveSeparatorRuleGeometry(DocxRelatedStoryKind.Footnote, bodyRun, bodyFontSizePoints, separatorMeasurer);
                    separatorY = contentTop + ResolveSeparatorGapPoints(bodyRun, bodyFontSizePoints, separatorMeasurer) + syntheticOffsetPoints;
                    separatorThickness = syntheticThicknessPoints;
                }

                DocxPlacedRelatedStoryLayout placedStory = PlaceRelatedStoryAtTop(activePage, activePageIndex, story.StoryLayout, story.Location.SourceBlockIndex, contentTop, separatorY)
                    with { UsesPlainFootnoteFrame = footerKeepOut is not null };
                if (separatorY is not null)
                {
                    placedStory = placedStory with
                    {
                        SeparatorThickness = separatorThickness,
                        SeparatorWidth = defaultContinuationWidth ?? placedStory.SeparatorWidth
                    };
                }

                activePlacedStories.Add(placedStory);
                outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
                contentTop -= storyHeight;
                if (trackedStory is not null)
                {
                    trackedStory.PlacedLineCount += story.StoryLayout.TextLines.Count;
                }
            }
            else
            {
                int headPageIndex = activePageIndex;
                int headCount = activePlacedStories.Count;
                PlaceRelatedStorySlices(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref contentTop, story.StoryLayout, story.Location.SourceBlockIndex, insertContinuationAfterActivePage: true, continuationSeparatorLayout, separatorMeasurer, printScale, headerKeepOut, stopAfterCurrentPage: sharedSlicePlacement, inFlight: trackedStory, frameBottom: placementFooter, usesPlainFootnoteFrame: footerKeepOut is not null);
                if (firstStory && !needsContinuationRule)
                {
                    PatchSyntheticRuleOntoHeadSlice(outputPages, headPageIndex, headCount, separatorMeasurer, defaultContinuationWidth);
                }
            }

            firstStory = false;
        }

        if (inFlightNotes is not null)
        {
            foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
            {
                if (story.Location.Reference.Id is not null && inFlightNotes.TryGetValue((story.Location.Reference.Kind, story.Location.Reference.Id), out InFlightRelatedStory? fresh) && !fresh.SeparatorPlaced)
                {
                    fresh.SeparatorPlaced = true;
                }
            }
        }
    }

    private static void PatchSyntheticRuleOntoHeadSlice(
        List<DocxLayoutPage> outputPages,
        int headPageIndex,
        int headCount,
        IDocxTextMeasurer? separatorMeasurer,
        double? separatorWidth = null)
    {
        DocxLayoutPage headPage = outputPages[headPageIndex];
        for (int storyIndex = headCount; storyIndex < headPage.PlacedRelatedStories.Count; storyIndex++)
        {
            DocxPlacedRelatedStoryLayout placed = headPage.PlacedRelatedStories[storyIndex];
            if (placed.StoryLayout.Story.Kind != DocxRelatedStoryKind.Footnote ||
                (placed.StoryLayout.Story.Type is not null && placed.StoryLayout.Story.Type != DocxRelatedStoryType.Normal))
            {
                continue;
            }

            (DocxTextRun? bodyRun, double bodyFontSizePoints) = FindSeparatorMarkFont(placed.StoryLayout.TextLines);
            (double syntheticOffsetPoints, double syntheticThicknessPoints) = ResolveSeparatorRuleGeometry(DocxRelatedStoryKind.Footnote, bodyRun, bodyFontSizePoints, separatorMeasurer);
            double syntheticGapPoints = ResolveSeparatorGapPoints(bodyRun, bodyFontSizePoints, separatorMeasurer);
            DocxPlacedRelatedStoryLayout[] patchedStories = new DocxPlacedRelatedStoryLayout[headPage.PlacedRelatedStories.Count];
            for (int copyIndex = 0; copyIndex < patchedStories.Length; copyIndex++)
            {
                patchedStories[copyIndex] = headPage.PlacedRelatedStories[copyIndex];
            }

            patchedStories[storyIndex] = placed with
            {
                SeparatorY = placed.TopY + syntheticGapPoints + syntheticOffsetPoints,
                SeparatorThickness = syntheticThicknessPoints,
                SeparatorWidth = separatorWidth ?? placed.SeparatorWidth,
            };
            outputPages[headPageIndex] = headPage with { PlacedRelatedStories = patchedStories };
            return;
        }
    }

    private static DocxPlacedRelatedStoryLayout PlaceRelatedStoryAtTop(
        DocxLayoutPage page,
        int pageIndex,
        DocxRelatedStoryLayout storyLayout,
        int sourceBlockIndex,
        double topY,
        double? separatorY)
    {
        return PlaceRelatedStoryAtTop(
            page,
            pageIndex,
            storyLayout,
            sourceBlockIndex,
            topY,
            storyTopOffset: 0d,
            storyHeight: ResolvePlacedStoryHeight(storyLayout, page),
            separatorY);
    }

    private static DocxPlacedRelatedStoryLayout PlaceRelatedStoryAtTop(
        DocxLayoutPage page,
        int pageIndex,
        DocxRelatedStoryLayout storyLayout,
        int sourceBlockIndex,
        double topY,
        double storyTopOffset,
        double storyHeight,
        double? separatorY)
    {
        double clampedStoryTopOffset = Math.Max(0d, storyTopOffset);
        double clampedStoryHeight = Math.Min(
            Math.Max(0d, storyHeight),
            Math.Max(0d, storyLayout.ContentHeight - clampedStoryTopOffset));
        double deltaY = topY + clampedStoryTopOffset;
        double storyWidthPoints = storyLayout.Story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote
            ? ResolvePageNoteTextWidth(page) : ResolvePageBodyWidth(page);

        return new DocxPlacedRelatedStoryLayout(
            storyLayout,
            storyLayout.StoryIndex,
            sourceBlockIndex,
            page.MarginLeft,
            topY,
            storyWidthPoints,
            clampedStoryHeight,
            clampedStoryTopOffset,
            storyLayout.ContentHeight,
            separatorY,
            FootnoteSeparatorWidthPoints,
            FootnoteSeparatorThicknessPoints,
            ShiftTextLines(storyLayout.TextLines, deltaY, page.MarginLeft),
            ShiftInlineImages(storyLayout.InlineImages, deltaY, page.MarginLeft),
            ShiftFloatingDrawings(storyLayout.FloatingDrawings, pageIndex, deltaY, page.MarginLeft),
            ShiftTableRows(storyLayout.TableRows, deltaY, page.MarginLeft));
    }

    private static Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> CreateRelatedStoryLookup(IReadOnlyList<DocxRelatedStoryLayout> relatedStoryLayouts)
    {
        var storyByKey = new Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout>();
        foreach (DocxRelatedStoryLayout story in relatedStoryLayouts)
        {
            if (story.Story.Id is not null && story.Story.IsNormalStoryType)
            {
                storyByKey.TryAdd((story.Story.Kind, story.Story.Id), story);
            }
        }

        return storyByKey;
    }

    private static DocxRelatedStoryLayout? FindSpecialRelatedStoryLayout(
        IReadOnlyList<DocxRelatedStoryLayout> relatedStoryLayouts,
        DocxRelatedStoryKind kind,
        DocxRelatedStoryType type)
    {
        return relatedStoryLayouts.FirstOrDefault(story =>
            story.Story.Id is not null &&
            story.ContentHeight > 0d &&
            story.Story.Kind == kind &&
            story.Story.Type == type);
    }

    private static IReadOnlyList<DocxRelatedStoryLayout> CreateRelatedStoryLayouts(
        IReadOnlyList<DocxRelatedStory> stories,
        double bodyWidth,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? unscaledTextMeasurer = null,
        IReadOnlyDictionary<DocxRelatedStory, DocxTextRun>? customNoteTargets = null,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (textMeasurer is null || stories.Count == 0)
        {
            var emptyLayouts = new DocxRelatedStoryLayout[stories.Count];
            for (int index = 0; index < stories.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                emptyLayouts[index] = new DocxRelatedStoryLayout(stories[index], index, [], [], [], [], 0d);
            }

            return emptyLayouts;
        }

        var layouts = new DocxRelatedStoryLayout[stories.Count];
        for (int index = 0; index < stories.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            layouts[index] = CreateRelatedStoryLayout(
                stories[index],
                index,
                bodyWidth,
                textMeasurer,
                defaultTabStopPoints,
                paragraphSpacingScale,
                pageNumber: null,
                pageCount: null,
                cancellationToken: cancellationToken,
                unscaledTextMeasurer: unscaledTextMeasurer, customNoteMarker: customNoteTargets?.GetValueOrDefault(stories[index]),
                mathEquations: mathEquations);
        }

        return layouts;
    }

    private static DocxRelatedStoryLayout CreateRelatedStoryLayout(
        DocxRelatedStory story,
        int storyIndex,
        double bodyWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        int? pageNumber,
        int? pageCount,
        CancellationToken cancellationToken,
        IDocxTextMeasurer? unscaledTextMeasurer = null,
        DocxTextRun? customNoteMarker = null,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? mathEquations = null)
    {
        double sourceStoryScale = paragraphSpacingScale;
        // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote and endnote
        // stories lay out unscaled in design space while emission maps uniformly
        // (WC first baseline 236.62 = affine-mapped design). Comments keep scaled
        // layout for the balloon path and textboxes keep their own design switch.
        if (story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
            Math.Abs(paragraphSpacingScale - 1d) >= 0.000000001d)
        {
            textMeasurer = unscaledTextMeasurer ?? textMeasurer;
            paragraphSpacingScale = 1d;
        }

        var textLines = new List<DocxTextLineLayout>();
        var inlineImages = new List<DocxInlineImageLayout>();
        var tableRows = new List<DocxTableRowLayout>();
        double cursorY = 0d;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        double? firstRelatedLineBaselineOffset = null;
        int paragraphIndex = 0;
        int tableIndex = 0;

        for (int elementIndex = 0; elementIndex < story.BodyElements.Count; elementIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxBodyElement element = story.BodyElements[elementIndex];
            if (element is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                firstRelatedLineBaselineOffset = null;
                var cellMemo = new DocxTableCellTextLinesMemo();
                DocxTableLayoutFrame frame = CreateTableLayoutFrame(
                    tableElement.Table,
                    tableIndex++,
                    elementIndex,
                    0d,
                    bodyWidth,
                    UnpagedRelatedStoryCanvasHeightPoints,
                    textMeasurer,
                    defaultTabStopPoints,
                    cancellationToken,
                    pageNumber: pageNumber,
                    pageCount: pageCount,
                    paragraphSpacingScale: paragraphSpacingScale,
                    mathEquations: mathEquations);
                int relatedStoryPageNumber = pageNumber ?? 1;
                for (int rowIndex = 0; rowIndex < tableElement.Table.Rows.Count; rowIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    double rowHeight = frame.RowHeights[rowIndex];
                    tableRows.Add(CreateTableRowLayout(
                        tableElement.Table,
                        frame.Context,
                        tableElement.Table.Rows[rowIndex],
                        rowIndex,
                        frame.RowHeights,
                        frame.EffectiveColumns,
                        frame.Scale,
                        textMeasurer,
                        defaultTabStopPoints,
                        () => relatedStoryPageNumber,
                        cursorY,
                        rowHeight,
                        cursorY,
                        FragmentIndex: 0,
                        FragmentCount: 1,
                        FragmentReason: "None",
                        Story: null,
                        pageCount: pageCount,
                        paragraphSpacingScale: paragraphSpacingScale,
                        cellMemo: cellMemo));
                    cursorY -= rowHeight;
                }

                continue;
            }

            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            if (story.Type is DocxRelatedStoryType.Separator or DocxRelatedStoryType.ContinuationSeparator &&
                IsExactLineSpacing(paragraph.EffectiveProperties))
            {
                // RV06 exact-line probe (Word 16.0, edge-endsepex): Office shifts
                // endnotes identically with exact-24 and auto separator lines, so
                // separator stories ignore exact line rules while the renderer
                // clamped the mark line box to 24pt (10pt content error). At-least
                // rules stay flowing (unprobed).
                paragraph = paragraph with
                {
                    LineSpacingPoints = null,
                    Spacing = paragraph.Spacing with { LineValue = null, LineRuleValue = null },
                };
            }

            // Note-story placement/reserve heights need their own qualification;
            // keep their existing spacing fallback while body contributions improve.
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale, resolveContextualContributions: false);
            if (story.Type is DocxRelatedStoryType.Separator or DocxRelatedStoryType.ContinuationSeparator)
            {
                // RV06 endnote-spacing probes (Word 16.0, edge-endsepsp/endsepspb):
                // Office holds separator-story output byte-identical across direct
                // before/after 0 vs 24pt, so separator stories lay out with
                // latent-default spacing while the renderer flowed direct values
                // into the separator height (24pt content error on endnotes).
                // Substitute latent defaults; docs without direct separator
                // spacing resolve identically and stay byte-identical.
                spacingProfile = spacingProfile with
                {
                    ParagraphBeforeSpacing = 0d,
                    ParagraphAfterSpacing = DocxDefaults.DefaultParagraphAfterSpacingPoints * paragraphSpacingScale,
                    AppliedBeforeSpacing = 0d,
                };
            }

            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            // Direct normal-note paragraphs share ordered measurement and retain
            // their design-space frame until the common note emission mapping.
            // Tables and special separator stories keep their existing path.
            if (story.IsNormalStoryType && story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
                paragraph.Equations.Count != 0)
            {
                double size = GetParagraphFontSize(paragraph);
                double lineHeight = ResolveLineHeightProfile(paragraph, size, textMeasurer, selectMaxAcrossRuns: true).LineHeight;
                double baseline = ResolveRelatedStoryBaselineOffset(paragraph, size, lineHeight, textMeasurer, DocxStoryId.Related(story.Kind).Kind);
                double firstInset = GetParagraphFirstLineTextStartOffset(paragraph, size, textMeasurer, paragraphSpacingScale, defaultTabStopPoints);
                double nextInset = GetParagraphTextStartOffset(paragraph, paragraphSpacingScale);
                double rightInset = GetParagraphRightInset(paragraph, paragraphSpacingScale);
                if (TryMeasureOrderedInlineContent(paragraph, mathEquations, textMeasurer, size,
                    Math.Max(1d, bodyWidth - firstInset - rightInset), Math.Max(1d, bodyWidth - nextInset - rightInset),
                    baseline, lineHeight, paragraphSpacingScale, cancellationToken, out var measured,
                    defaultTabStopPoints: defaultTabStopPoints))
                {
                    for (int index = 0; index < measured.Count; index++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var line = measured[index];
                        double inset = index == 0 ? firstInset : nextInset;
                        double available = Math.Max(1d, bodyWidth - inset - rightInset);
                        double align = ResolveOrderedInlineAlignment(line, paragraph.Alignment, available);
                        var placed = PlaceOrderedInlineLine(paragraph, line, inset + align, 0d, cursorY,
                            size, paragraphSpacingScale, 0, elementIndex, paragraphIndex, index, DocxStoryId.Related(story.Kind));
                        textLines.Add(placed.Line with
                        {
                            AppliedBeforeSpacing = index == 0 ? spacingProfile.AppliedBeforeSpacing : 0d,
                            ParagraphBeforeSpacing = index == 0 ? spacingProfile.ParagraphBeforeSpacing : 0d,
                            ParagraphAfterSpacing = index == measured.Count - 1 ? spacingProfile.ParagraphAfterSpacing : 0d,
                            PendingAfterSpacing = null
                        });
                        inlineImages.AddRange(placed.Images);
                        cursorY -= line.OccupiedHeight;
                    }
                    if (measured.Count != 0) firstRelatedLineBaselineOffset ??= measured[0].BaselineInset;
                    pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
                    previousParagraph = paragraph;
                    paragraphIndex++;
                    continue;
                }
            }
            (IReadOnlyList<DocxTextLineLayout> paragraphLines, IReadOnlyList<DocxInlineImageLayout> placedStoryImages, double paragraphUsedHeight, double paragraphBaselineOffset) = LayoutRelatedStoryParagraphTextLines(
                paragraph,
                paragraphSpacingScale,
                elementIndex,
                paragraphIndex,
                DocxStoryId.Related(story.Kind),
                bodyWidth,
                cursorY,
                spacingProfile,
                textMeasurer,
                defaultTabStopPoints,
                pageNumber,
                pageCount,
                firstRelatedLineBaselineOffset,
                paragraphSpacingScale,
                sourceStoryScale,
                customNoteMarker);
            if (paragraphLines.Count != 0)
            {
                firstRelatedLineBaselineOffset ??= paragraphBaselineOffset;
            }
            textLines.AddRange(paragraphLines);
            inlineImages.AddRange(placedStoryImages);
            cursorY -= paragraphUsedHeight;
            if (paragraphLines.Count == 0 && paragraph.Images.Count == 0)
            {
                double fontSize = GetParagraphFontSize(paragraph);
                cursorY -= ResolveLineHeight(paragraph, fontSize, textMeasurer);
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (placedStoryImages.Any(placed => ReferenceEquals(placed.Image, image)))
                {
                    continue;
                }

                double imageWidth = Math.Min(bodyWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                double imageX = paragraph.EffectiveProperties.Alignment switch
                {
                    DocxTextAlignment.Center => Math.Max(0, bodyWidth - imageWidth) / 2d,
                    DocxTextAlignment.Right => Math.Max(0, bodyWidth - imageWidth),
                    _ => 0d
                };
                inlineImages.Add(new DocxInlineImageLayout(
                    image,
                    imageX,
                    cursorY - imageHeight,
                    imageWidth,
                    imageHeight,
                    PageIndex: 0,
                    SourceBlockIndex: elementIndex,
                    SourceParagraphIndex: paragraphIndex, Story: null));
                cursorY -= imageHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        cursorY -= pendingSpacingAfter;
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings = CreateRelatedStoryFloatingDrawingLayouts(
            story,
            bodyWidth,
            textLines,
            inlineImages,
            tableRows,
            textMeasurer,
            defaultTabStopPoints,
            paragraphSpacingScale,
            cancellationToken,
            pageNumber,
            pageCount,
            unscaledTextMeasurer);
        return new DocxRelatedStoryLayout(story, storyIndex, textLines.ToArray(), inlineImages.ToArray(), floatingDrawings, tableRows.ToArray(), Math.Abs(cursorY));
    }
}
