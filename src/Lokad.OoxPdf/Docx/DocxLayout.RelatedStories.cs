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
        ISet<(DocxRelatedStoryKind Kind, string Id)>? alreadyPlacedStoryKeys = null)
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
            double footnoteWidth = Math.Max(1d, bodyWidth + page.MarkupMarginReservePoints);
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
            PlaceFootnoteStories(outputPages, outputIndex, page, pageFootnoteStories, footnoteSeparatorLayout, footnoteContinuationLayout, separatorMeasurer, printScale, headerKeepOut);
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

                IReadOnlyList<DocxRelatedStoryLayout> groupLayouts = resolveRelatedStoryLayouts(ResolvePageBodyWidth(outputPages[sectionEndPageIndex]));
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
                Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> endnoteStoryByKey = CreateRelatedStoryLookup(resolveRelatedStoryLayouts(ResolvePageBodyWidth(outputPages[^1])));
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
                DocxRelatedStoryLayout? documentEndSeparatorLayout = FindSpecialRelatedStoryLayout(resolveRelatedStoryLayouts(ResolvePageBodyWidth(activePage)), DocxRelatedStoryKind.Endnote, DocxRelatedStoryType.Separator);
                DocxRelatedStoryLayout? documentEndContinuationLayout = FindSpecialRelatedStoryLayout(resolveRelatedStoryLayouts(ResolvePageBodyWidth(activePage)), DocxRelatedStoryKind.Endnote, DocxRelatedStoryType.ContinuationSeparator) ?? documentEndSeparatorLayout;
                if (CanUseNominalDocumentEndnoteFlow(activePage, documentEndSeparatorLayout, printScale, documentEndStories))
                {
                    cursorTop -= ResolveTrailingBodyAfterSpacing(activePage);
                }
                cursorTop = ResolveFittingDocumentEndnoteStartTop(activePage, activePlacedStories, cursorTop,
                    documentEndStories, documentEndSeparatorLayout, separatorMeasurer, printScale);
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
                    if (ResolveDocumentEndnoteSeparatorHeight(documentEndSeparatorLayout, activePage, separatorMeasurer, documentEndStories, printScale) > cursorTop - activePage.MarginBottom)
                    {
                        MoveToRelatedStoryContinuationPage(documentEndPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref cursorTop, insertContinuationAfterActivePage: false, headerKeepOut);
                    }
                    double? nominalSeparatorHeight = CanUseNominalDocumentEndnoteFlow(activePage, documentEndSeparatorLayout, printScale, documentEndStories)
                        ? ResolveDocumentEndnoteSeparatorHeight(documentEndSeparatorLayout, activePage, separatorMeasurer, documentEndStories, printScale) : null;
                    (DocxPlacedRelatedStoryLayout placedSeparator, double separatorBottom) = PlaceSeparatorStoryWithMark(activePage, activePageIndex, documentEndSeparatorLayout, sourceBlockIndex: -1, cursorTop + FootnoteSeparatorGapPoints, separatorMeasurer, useSizeDrivenPlacementHeight: true, contentStoriesForHhea: documentEndStories, placementHeightOverride: nominalSeparatorHeight);
                    activePlacedStories.Add(placedSeparator);
                    documentEndPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
                    double endnoteContentGapPoints = 0d;
                    foreach (DocxRelatedStoryLayout endnoteContentLayout in documentEndStories)
                    {
                        endnoteContentGapPoints = System.Math.Max(endnoteContentGapPoints, ResolveFootnoteContentGapPoints(endnoteContentLayout.TextLines, endnoteMarkRun, endnoteMarkSize, separatorMeasurer));
                    }

                    cursorTop = separatorBottom - System.Math.Max(ResolveSeparatorGapPoints(endnoteMarkRun, endnoteMarkSize, separatorMeasurer), endnoteContentGapPoints);
                }
                foreach (DocxRelatedStoryLayout endnoteStoryLayout in documentEndStories)
                {
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
                        headerKeepOut);
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
    private static double[] GetStoryTextLineBoxHeights(DocxRelatedStoryLayout storyLayout)
    {
        double[] heights = new double[storyLayout.TextLines.Count];
        for (int lineIndex = 0; lineIndex < heights.Length; lineIndex++)
        {
            DocxTextLineLayout line = storyLayout.TextLines[lineIndex];
            double boxHeight = line.LineHeight ?? line.SingleLineHeight ?? line.FontSize;
            double afterSpacing = line.ParagraphAfterSpacing ?? line.PendingAfterSpacing ?? 0d;
            double beforeSpacing = line.IsFirstParagraphLine == true && line.ContextualSpacingSuppressed != true
                ? line.ParagraphBeforeSpacing ?? line.AppliedBeforeSpacing ?? 0d
                : 0d;
            heights[lineIndex] = Math.Max(0d, boxHeight) + Math.Max(0d, afterSpacing) + Math.Max(0d, beforeSpacing);
        }

        return heights;
    }

    private static double[] GetStoryLineAfterSpacings(DocxRelatedStoryLayout storyLayout)
    {
        double[] afters = new double[storyLayout.TextLines.Count];
        for (int lineIndex = 0; lineIndex < afters.Length; lineIndex++)
        {
            DocxTextLineLayout line = storyLayout.TextLines[lineIndex];
            afters[lineIndex] = Math.Max(0d, line.ParagraphAfterSpacing ?? line.PendingAfterSpacing ?? 0d);
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

        if (inFlight.RemainingLineCount > 0)
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
        IDocxTextMeasurer? separatorMeasurer)
    {
        if (continuationLayout is null || continuationLayout.TextLines.Count == 0)
        {
            return 0d;
        }

        (DocxTextRun? takeGapRun, double takeGapFontSizePoints) = FindSeparatorMarkFont(continuationLayout.TextLines);
        double takeGapPoints = ResolveSeparatorGapPoints(takeGapRun, takeGapFontSizePoints, separatorMeasurer);
        (double takeRuleBottomOffsetPoints, double takeRuleThicknessPoints) = ResolveSeparatorRuleGeometry(continuationLayout.Story.Kind, takeGapRun, takeGapFontSizePoints, separatorMeasurer);
        return Math.Max(0d, takeGapPoints + takeRuleBottomOffsetPoints + takeRuleThicknessPoints - 2.1);
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
        double headContentPushdownPoints = 0d)
    {
        if (storyLayout.TextLines.Count == 0)
        {
            PlaceRelatedStorySlicesByHeight(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref cursorTop, storyLayout, sourceBlockIndex, insertContinuationAfterActivePage);
            return;
        }

        double[] lineBoxes = GetStoryTextLineBoxHeights(storyLayout);
        double[] lineAfters = GetStoryLineAfterSpacings(storyLayout);
        int lineIndex = 0;
        double consumedHeight = 0d;
        while (lineIndex < storyLayout.TextLines.Count)
        {
            double availableHeight = Math.Max(0d, cursorTop - activePage.MarginBottom);
            if (availableHeight <= 0.001d)
            {
                if (stopAfterCurrentPage)
                {
                    return;
                }

                MoveToRelatedStoryContinuationPage(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref cursorTop, insertContinuationAfterActivePage, headerKeepOut);
                availableHeight = Math.Max(0d, cursorTop - activePage.MarginBottom);
                cursorTop = PlaceContinuationSeparatorIfNeeded(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, cursorTop, sourceBlockIndex, storyLayout, lineIndex, lineBoxes, continuationSeparatorLayout, separatorMeasurer, printScale);
                availableHeight = Math.Max(0d, cursorTop - activePage.MarginBottom);
            }

            int remainingLines = storyLayout.TextLines.Count - lineIndex;
            double[]? fitAfters = printScale >= 1d ? lineAfters : null;
            (int takeCount, double takeHeight) = TakeStoryLines(lineBoxes, lineIndex, remainingLines, availableHeight, fitAfters);
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
                separatorY: null));
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
                cursorTop = PlaceContinuationSeparatorIfNeeded(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, cursorTop, sourceBlockIndex, storyLayout, lineIndex, lineBoxes, continuationSeparatorLayout, separatorMeasurer, printScale);
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
    // Longest whole-line prefix (at least one line) fitting the capacity.
    // RV06 after-spacing sweep (Word 16.0, seven after values): take needs exclude the
    // take last-line trailing after-spacing while take heights keep full boxes for
    // placement, so final takes and null after-spacing callers resolve byte-identically.
    // The exclusion applies on unscaled takes only; word-compatible spacing-scale takes
    // stay legacy pending their own probe series, so callers pass null below scale 1.
    private static (int TakeCount, double TakeHeight) TakeStoryLines(double[] lineBoxes, int startLineIndex, int remainingLineCount, double capacityHeight, double[]? lineAfterSpacings = null)
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

        if (takeCount == 0)
        {
            takeCount = 1;
            takeHeight = lineBoxes[startLineIndex];
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
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null)
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
            double continuationOverhead = printScale >= 1d
                ? ContinuationTakeCharge(continuationSeparatorLayout, separatorMeasurer)
                : ContinuationRuleOverhead(continuationSeparatorLayout, separatorMeasurer);
            double reservedTopSpace = Math.Max(0d, activePage.Height - activePage.MarginTop - cursorTop);
            double netOverhead = Math.Max(0d, continuationOverhead - reservedTopSpace);
            (_, double takeHeight) = TakeStoryLines(bodyLineBoxes, bodyLineIndex, remainingLines, Math.Max(0d, cursorTop - activePage.MarginBottom - netOverhead), printScale >= 1d ? null : bodyFitAfters);
            double contentTop = activePage.MarginBottom + takeHeight;
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
        double reservedMarginPoints = Math.Abs(printScale - 1d) < 0.000000001d ? 0d : page.MarkupMarginReservePoints;
        double ruleWidthPoints = Math.Max(1d, page.Width - page.MarginLeft - page.MarginRight + reservedMarginPoints);
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
    private static double ResolveContentHheaAscenderExcess(IReadOnlyList<DocxRelatedStoryLayout>? contentStories, IDocxTextMeasurer? separatorMeasurer)
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
                        excessMax = System.Math.Max(excessMax, runExcessEm * EndnoteSeparatorHheaAscenderExcessSlope * runSize);
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
        double? placementHeightOverride = null)
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
        DocxPlacedRelatedStoryLayout placedSeparator = PlaceRelatedStoryAtTop(page, pageIndex, separatorLayout, sourceBlockIndex, separatorTop, separatorY: separatorBottom + ruleBottomOffsetPoints);
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
        double? sharedContentTop = null)
    {
        if (footnoteStories.Count == 0)
        {
            return;
        }

        if (separatorLayout is null)
        {
            PlaceFootnoteStoriesWithoutSeparator(outputPages, outputIndex, page, footnoteStories, continuationSeparatorLayout, separatorMeasurer, printScale, headerKeepOut, sharedSlicePlacement, inFlightNotes, pageNumberOverride, pageCountOverride);
            return;
        }

        DocxLayoutPage activePage = outputPages[outputIndex];
        int activePageIndex = outputIndex;
        List<DocxPlacedRelatedStoryLayout> activePlacedStories = activePage.PlacedRelatedStories.ToList();
        double bodyHeight = footnoteStories.Sum(story => ResolvePlacedStoryHeight(story.StoryLayout, activePage));
        double cursorTop = activePage.MarginBottom + bodyHeight;
        double separatorHeight = ResolvePlacedStoryHeight(separatorLayout, activePage);
        (DocxTextRun? gapRun, double gapFontSizePoints) = FindSeparatorMarkFont(separatorLayout.TextLines);
        double separatorGapPoints = ResolveSeparatorGapPoints(gapRun, gapFontSizePoints, separatorMeasurer);
        double storiesTop = cursorTop;
        double clampBodyBottom = ResolveFootnoteClampBodyBottom(page, pageNumberOverride ?? (outputIndex + 1), pageCountOverride ?? outputPages.Count, footnoteStories[0].StoryLayout, headerKeepOut);
        storiesTop = Math.Min(cursorTop, clampBodyBottom - separatorHeight);
        if (sharedContentTop is not null)
        {
            storiesTop = Math.Min(storiesTop, sharedContentTop.Value);
        }
        // RV06 p1-position probes (Word 16.0): Office seats the overflowing head block
        // bottom-up (storiesTop = margin + takeHeight), so absolute positions match instead
        // of leaving capacity slack above the margin. The head take still comes from the
        // same clamped capacity, so take counts are unchanged; fitting stories keep the
        // legacy whole placement bit-identically.
        DocxRelatedStoryLayout firstStoryLayout = footnoteStories[0].StoryLayout;
        if (firstStoryLayout.TextLines.Count != 0)
        {
            double[] headLineBoxes = GetStoryTextLineBoxHeights(firstStoryLayout);
            double[]? headFitAfters = printScale >= 1d ? GetStoryLineAfterSpacings(firstStoryLayout) : null;
            (int headTakeCount, double headTakeHeight) = TakeStoryLines(headLineBoxes, 0, firstStoryLayout.TextLines.Count, Math.Max(0d, storiesTop - activePage.MarginBottom), headFitAfters);
            if (headTakeCount < firstStoryLayout.TextLines.Count)
            {
                storiesTop = activePage.MarginBottom + headTakeHeight;
            }
        }
        bool drawHeadSeparator = true;
        if (inFlightNotes is not null)
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
        (DocxPlacedRelatedStoryLayout placedSeparator, _) = PlaceSeparatorStoryWithMark(activePage, activePageIndex, separatorLayout, footnoteStories[0].Location.SourceBlockIndex, separatorTop, separatorMeasurer);
        activePlacedStories.Add(placedSeparator);
        outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
        }
        double footnoteContentGapPoints = ResolveFootnoteContentGapPoints(footnoteStories, gapRun, gapFontSizePoints, separatorMeasurer);
        double footnoteContentPushdownPoints = System.Math.Max(0d, footnoteContentGapPoints - separatorGapPoints);
        double contentTop = storiesTop;
        foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
        {
            InFlightRelatedStory? trackedStory = inFlightNotes is not null && story.Location.Reference.Id is not null && inFlightNotes.TryGetValue((story.Location.Reference.Kind, story.Location.Reference.Id), out InFlightRelatedStory? tracked) ? tracked : null;
            double storyHeight = ResolvePlacedStoryHeight(story.StoryLayout, activePage);
            if (trackedStory is not null && trackedStory.SeparatorPlaced && continuationSeparatorLayout is not null && continuationSeparatorLayout.TextLines.Count != 0)
            {
                (DocxTextRun? continuationGapRun, double continuationGapFontSizePoints) = FindSeparatorMarkFont(continuationSeparatorLayout.TextLines);
                double continuationGapPoints = ResolveSeparatorGapPoints(continuationGapRun, continuationGapFontSizePoints, separatorMeasurer);
                (DocxPlacedRelatedStoryLayout placedContinuation, _) = PlaceContinuationSeparatorStory(activePage, activePageIndex, continuationSeparatorLayout, story.Location.SourceBlockIndex, contentTop + continuationGapPoints, separatorMeasurer, printScale);
                activePlacedStories.Add(placedContinuation);
                outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
            }

            if (contentTop - storyHeight >= activePage.MarginBottom - 0.001d)
            {
                DocxPlacedRelatedStoryLayout placedStory = PlaceRelatedStoryAtTop(activePage, activePageIndex, story.StoryLayout, story.Location.SourceBlockIndex, contentTop - footnoteContentPushdownPoints, separatorY: null);
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
                PlaceRelatedStorySlices(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref contentTop, story.StoryLayout, story.Location.SourceBlockIndex, insertContinuationAfterActivePage: true, continuationSeparatorLayout, separatorMeasurer, printScale, headerKeepOut, stopAfterCurrentPage: sharedSlicePlacement, inFlight: trackedStory, headContentPushdownPoints: footnoteContentPushdownPoints);
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
        double? sharedContentTop = null)
    {
        DocxLayoutPage activePage = outputPages[outputIndex];
        int activePageIndex = outputIndex;
        List<DocxPlacedRelatedStoryLayout> activePlacedStories = activePage.PlacedRelatedStories.ToList();
        double bodyHeight = footnoteStories.Sum(story => ResolvePlacedStoryHeight(story.StoryLayout, activePage));
        double cursorTop = activePage.MarginBottom + bodyHeight;
        double storiesTop = cursorTop;
        DocxRelatedStoryLayout firstLayout = footnoteStories[0].StoryLayout;
        double clampBodyBottom = ResolveFootnoteClampBodyBottom(page, pageNumberOverride ?? (outputIndex + 1), pageCountOverride ?? outputPages.Count, firstLayout, headerKeepOut);
        if (firstLayout.TextLines.Count != 0)
        {
            double firstBoxHeight = GetStoryTextLineBoxHeights(firstLayout)[0];
            storiesTop = Math.Min(cursorTop, clampBodyBottom - firstBoxHeight);
        if (sharedContentTop is not null)
        {
            storiesTop = Math.Min(storiesTop, sharedContentTop.Value);
        }
        }
        double contentTop = storiesTop;
        bool firstStory = true;
        foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
        {
            InFlightRelatedStory? trackedStory = inFlightNotes is not null && story.Location.Reference.Id is not null && inFlightNotes.TryGetValue((story.Location.Reference.Kind, story.Location.Reference.Id), out InFlightRelatedStory? tracked) ? tracked : null;
            double storyHeight = ResolvePlacedStoryHeight(story.StoryLayout, activePage);
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

            if (contentTop - storyHeight >= activePage.MarginBottom - 0.001d)
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

                DocxPlacedRelatedStoryLayout placedStory = PlaceRelatedStoryAtTop(activePage, activePageIndex, story.StoryLayout, story.Location.SourceBlockIndex, contentTop, separatorY);
                if (separatorY is not null)
                {
                    placedStory = placedStory with { SeparatorThickness = separatorThickness };
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
                PlaceRelatedStorySlices(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref contentTop, story.StoryLayout, story.Location.SourceBlockIndex, insertContinuationAfterActivePage: true, continuationSeparatorLayout, separatorMeasurer, printScale, headerKeepOut, stopAfterCurrentPage: sharedSlicePlacement, inFlight: trackedStory);
                if (firstStory && !needsContinuationRule)
                {
                    PatchSyntheticRuleOntoHeadSlice(outputPages, headPageIndex, headCount, separatorMeasurer);
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
        IDocxTextMeasurer? separatorMeasurer)
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
        double storyWidthPoints = Math.Max(1d, page.Width - page.MarginLeft - page.MarginRight);
        if (storyLayout.Story.Kind == DocxRelatedStoryKind.Footnote && page.MarkupMarginReservePoints > 0d)
        {
            storyWidthPoints += page.MarkupMarginReservePoints;
        }

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
        IDocxTextMeasurer? unscaledTextMeasurer = null)
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
                unscaledTextMeasurer: unscaledTextMeasurer);
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
        IDocxTextMeasurer? unscaledTextMeasurer = null)
    {
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
                    paragraphSpacingScale: paragraphSpacingScale);
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
                paragraphSpacingScale);
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
