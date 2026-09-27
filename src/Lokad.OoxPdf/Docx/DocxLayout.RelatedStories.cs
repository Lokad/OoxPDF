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
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null)
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
            DocxRelatedStoryLayout? footnoteSeparatorLayout = FindSpecialRelatedStoryLayout(pageRelatedStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.Separator);
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

                    if (placedStoryKeys.Contains((reference.Kind, reference.Id)))
                    {
                        continue;
                    }

                    if (!storyByKey.TryGetValue((reference.Kind, reference.Id), out DocxRelatedStoryLayout? storyLayout) ||
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
            DocxRelatedStoryLayout? footnoteContinuationLayout = FindSpecialRelatedStoryLayout(pageRelatedStoryLayouts, DocxRelatedStoryKind.Footnote, DocxRelatedStoryType.ContinuationSeparator) ?? footnoteSeparatorLayout;
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
                if (documentEndSeparatorLayout is not null && documentEndStories.Count > 0)
                {
                    // RV06 endnote probes: document-end endnotes draw the separator rule
                    // with a mark like section-end endnotes (no extra gap above it).
                    (DocxPlacedRelatedStoryLayout placedSeparator, double separatorBottom) = PlaceSeparatorStoryWithMark(activePage, activePageIndex, documentEndSeparatorLayout, sourceBlockIndex: -1, cursorTop + FootnoteSeparatorGapPoints, separatorMeasurer);
                    activePlacedStories.Add(placedSeparator);
                    documentEndPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
                    cursorTop = separatorBottom - FootnoteSeparatorGapPoints;
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
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null)
    {
        if (storyLayout.TextLines.Count == 0)
        {
            PlaceRelatedStorySlicesByHeight(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref cursorTop, storyLayout, sourceBlockIndex, insertContinuationAfterActivePage);
            return;
        }

        double[] lineBoxes = GetStoryTextLineBoxHeights(storyLayout);
        int lineIndex = 0;
        double consumedHeight = 0d;
        while (lineIndex < storyLayout.TextLines.Count)
        {
            double availableHeight = Math.Max(0d, cursorTop - activePage.MarginBottom);
            if (availableHeight <= 0.001d)
            {
                MoveToRelatedStoryContinuationPage(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref cursorTop, insertContinuationAfterActivePage, headerKeepOut);
                availableHeight = Math.Max(0d, cursorTop - activePage.MarginBottom);
                cursorTop = PlaceContinuationSeparatorIfNeeded(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, cursorTop, sourceBlockIndex, storyLayout, lineIndex, lineBoxes, continuationSeparatorLayout, separatorMeasurer, printScale);
                availableHeight = Math.Max(0d, cursorTop - activePage.MarginBottom);
            }

            int remainingLines = storyLayout.TextLines.Count - lineIndex;
            (int takeCount, double takeHeight) = TakeStoryLines(lineBoxes, lineIndex, remainingLines, availableHeight);
            if (lineIndex + takeCount >= storyLayout.TextLines.Count)
            {
                takeHeight = Math.Min(Math.Max(0d, storyLayout.ContentHeight - consumedHeight), availableHeight);
            }

            DocxRelatedStoryLayout sliceLayout = NarrowStoryTextLines(storyLayout, lineIndex, takeCount);
            activePlacedStories.Add(PlaceRelatedStoryAtTop(
                activePage,
                activePageIndex,
                sliceLayout,
                sourceBlockIndex,
                cursorTop,
                consumedHeight,
                takeHeight,
                separatorY: null));
            outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
            lineIndex += takeCount;
            consumedHeight += takeHeight;
            cursorTop -= takeHeight + FootnoteSeparatorGapPoints;
            if (lineIndex < storyLayout.TextLines.Count)
            {
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
    private static (int TakeCount, double TakeHeight) TakeStoryLines(double[] lineBoxes, int startLineIndex, int remainingLineCount, double capacityHeight)
    {
        int takeCount = 0;
        double takeHeight = 0d;
        while (takeCount < remainingLineCount &&
            takeHeight + lineBoxes[startLineIndex + takeCount] <= capacityHeight + 0.001d)
        {
            takeHeight += lineBoxes[startLineIndex + takeCount];
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
            (_, double takeHeight) = TakeStoryLines(bodyLineBoxes, bodyLineIndex, remainingLines, Math.Max(0d, cursorTop - activePage.MarginBottom));
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

    private sealed record DocxReferencedRelatedStoryLayout(
        DocxInlineReferenceLocation Location,
        DocxRelatedStoryLayout StoryLayout);

    private readonly record struct DocxPageTextLineOwner(
        DocxTextLineLayout Line,
        int? SourceBlockIndex);


    // Word lays separator stories compact around the rule and space instead of a full line box, so the
    // blank baseline rides at the story bottom (see constants in DocxLayout); shifting is rigid across lines.
    // RV06 separator-bottom probes (Word 16.0, Times/Aptos/Calibri footnote points):
    // Office footnote marks sit at the separator block bottom (ride ~0), so the
    // rule-mark gap carries the full strikeout offset. Endnote marks keep the legacy
    // 0.15pt ride that matches the Office endnote probes exactly.
    private static DocxPlacedRelatedStoryLayout ShiftSeparatorStoryToBaseline(DocxPlacedRelatedStoryLayout placed, double bottom)
    {
        if (placed.TextLines.Count == 0)
        {
            return placed;
        }

        double ridePoints = placed.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote
            ? FootnoteSeparatorBaselineOffsetPoints
            : 0d;
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
        if (separatorMeasurer is not null &&
            markFontSizePoints > 0d &&
            separatorMeasurer.TryGetStrikeoutRuleMetrics(markRun, out double positionEm, out double thicknessEm) &&
            positionEm > 0d &&
            thicknessEm > 0d)
        {
            double thicknessPoints = thicknessEm * markFontSizePoints;
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

    private static (DocxPlacedRelatedStoryLayout Placed, double SeparatorBottom) PlaceSeparatorStoryWithMark(
        DocxLayoutPage page,
        int pageIndex,
        DocxRelatedStoryLayout separatorLayout,
        int sourceBlockIndex,
        double separatorTop,
        IDocxTextMeasurer? separatorMeasurer)
    {
        double separatorHeight = ResolvePlacedStoryHeight(separatorLayout, page);
        double separatorBottom = separatorTop - separatorHeight;
        (DocxTextRun? markRun, double markFontSizePoints) = FindSeparatorMarkFont(separatorLayout.TextLines);
        (double ruleBottomOffsetPoints, double ruleThicknessPoints) = ResolveSeparatorRuleGeometry(separatorLayout.Story.Kind, markRun, markFontSizePoints, separatorMeasurer);
        DocxPlacedRelatedStoryLayout placedSeparator = PlaceRelatedStoryAtTop(page, pageIndex, separatorLayout, sourceBlockIndex, separatorTop, separatorY: separatorBottom + ruleBottomOffsetPoints);
        placedSeparator = placedSeparator with { SeparatorThickness = ruleThicknessPoints };
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
    private static void PlaceFootnoteStories(
        List<DocxLayoutPage> outputPages,
        int outputIndex,
        DocxLayoutPage page,
        IReadOnlyList<DocxReferencedRelatedStoryLayout> footnoteStories,
        DocxRelatedStoryLayout? separatorLayout,
        DocxRelatedStoryLayout? continuationSeparatorLayout,
        IDocxTextMeasurer? separatorMeasurer,
        double printScale = 1d, 
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null)
    {
        if (footnoteStories.Count == 0)
        {
            return;
        }

        if (separatorLayout is null)
        {
            PlaceFootnoteStoriesWithoutSeparator(outputPages, outputIndex, page, footnoteStories, continuationSeparatorLayout, separatorMeasurer, printScale, headerKeepOut);
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
        double? bodyBottomEdge = BodyBottomEdge(page);
        if (bodyBottomEdge.HasValue)
        {
            storiesTop = Math.Min(cursorTop, bodyBottomEdge.Value - separatorHeight);
        }
        double separatorTop = storiesTop + separatorGapPoints + separatorHeight;
        (DocxPlacedRelatedStoryLayout placedSeparator, _) = PlaceSeparatorStoryWithMark(activePage, activePageIndex, separatorLayout, footnoteStories[0].Location.SourceBlockIndex, separatorTop, separatorMeasurer);
        activePlacedStories.Add(placedSeparator);
        outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
        double contentTop = storiesTop;
        foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
        {
            double storyHeight = ResolvePlacedStoryHeight(story.StoryLayout, activePage);
            if (contentTop - storyHeight >= activePage.MarginBottom - 0.001d)
            {
                DocxPlacedRelatedStoryLayout placedStory = PlaceRelatedStoryAtTop(activePage, activePageIndex, story.StoryLayout, story.Location.SourceBlockIndex, contentTop, separatorY: null);
                activePlacedStories.Add(placedStory);
                outputPages[activePageIndex] = activePage with { PlacedRelatedStories = activePlacedStories.ToArray() };
                contentTop -= storyHeight;
            }
            else
            {
                PlaceRelatedStorySlices(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref contentTop, story.StoryLayout, story.Location.SourceBlockIndex, insertContinuationAfterActivePage: true, continuationSeparatorLayout, separatorMeasurer, printScale, headerKeepOut);
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
        Func<DocxLayoutPage, int, int, double>? headerKeepOut = null)
    {
        DocxLayoutPage activePage = outputPages[outputIndex];
        int activePageIndex = outputIndex;
        List<DocxPlacedRelatedStoryLayout> activePlacedStories = activePage.PlacedRelatedStories.ToList();
        double bodyHeight = footnoteStories.Sum(story => ResolvePlacedStoryHeight(story.StoryLayout, activePage));
        double cursorTop = activePage.MarginBottom + bodyHeight;
        double storiesTop = cursorTop;
        DocxRelatedStoryLayout firstLayout = footnoteStories[0].StoryLayout;
        double? bodyBottomEdge = BodyBottomEdge(page);
        if (firstLayout.TextLines.Count != 0 && bodyBottomEdge.HasValue)
        {
            double firstBoxHeight = GetStoryTextLineBoxHeights(firstLayout)[0];
            storiesTop = Math.Min(cursorTop, bodyBottomEdge.Value - firstBoxHeight);
        }
        double contentTop = storiesTop;
        bool firstStory = true;
        foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
        {
            double storyHeight = ResolvePlacedStoryHeight(story.StoryLayout, activePage);
            if (contentTop - storyHeight >= activePage.MarginBottom - 0.001d)
            {
                double? separatorY = null;
                double separatorThickness = FootnoteSeparatorThicknessPoints;
                if (firstStory)
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
            }
            else
            {
                int headPageIndex = activePageIndex;
                int headCount = activePlacedStories.Count;
                PlaceRelatedStorySlices(outputPages, ref activePageIndex, ref activePage, ref activePlacedStories, ref contentTop, story.StoryLayout, story.Location.SourceBlockIndex, insertContinuationAfterActivePage: true, continuationSeparatorLayout, separatorMeasurer, printScale, headerKeepOut);
                if (firstStory)
                {
                    PatchSyntheticRuleOntoHeadSlice(outputPages, headPageIndex, headCount, separatorMeasurer);
                }
            }

            firstStory = false;
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

    // Body-anchored footnote area top: lowest body text baseline minus one em
    // (GetVerticalBounds convention), so overflowing blocks clamp below body content.
    private static double? BodyBottomEdge(DocxLayoutPage page)
    {
        double? bottomEdge = null;
        foreach (DocxLayoutItem item in page.Items)
        {
            if (item is DocxTextLineLayout line)
            {
                double lineBottom = line.BaselineY - line.FontSize;
                bottomEdge = bottomEdge.HasValue ? Math.Min(bottomEdge.Value, lineBottom) : lineBottom;
            }
        }

        return bottomEdge;
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
        return new DocxPlacedRelatedStoryLayout(
            storyLayout,
            storyLayout.StoryIndex,
            sourceBlockIndex,
            page.MarginLeft,
            topY,
            Math.Max(1d, page.Width - page.MarginLeft - page.MarginRight),
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
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter, paragraphSpacingScale);
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
