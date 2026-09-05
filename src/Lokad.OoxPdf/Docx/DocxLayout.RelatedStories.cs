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
        CancellationToken cancellationToken)
    {
        var reserveHeightBySourceBlock = new Dictionary<int, double>();
        if (relatedStoryLayouts.Count == 0)
        {
            return reserveHeightBySourceBlock;
        }

        Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> storyByKey = CreateRelatedStoryLookup(relatedStoryLayouts);
        DocxRelatedStoryLayout? footnoteSeparatorLayout = FindSpecialRelatedStoryLayout(relatedStoryLayouts, DocxRelatedStoryKind.Footnote, "separator");
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
                        reserveHeight += footnoteSeparatorLayout is null
                            ? FootnoteSeparatorGapPoints
                            : Math.Max(0d, footnoteSeparatorLayout.ContentHeight) + FootnoteSeparatorGapPoints;
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
        CancellationToken cancellationToken)
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

        var pagesWithStories = new DocxLayoutPage[pages.Count];
        var placedStoryKeys = new HashSet<(DocxRelatedStoryKind Kind, string Id)>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutPage page = pages[pageIndex];
            IReadOnlyList<DocxRelatedStoryLayout> pageRelatedStoryLayouts = pageIndex == 0
                ? relatedStoryLayouts
                : resolveRelatedStoryLayouts(ResolvePageBodyWidth(page));
            storyByKey = CreateRelatedStoryLookup(pageRelatedStoryLayouts);
            DocxRelatedStoryLayout? footnoteSeparatorLayout = FindSpecialRelatedStoryLayout(pageRelatedStoryLayouts, DocxRelatedStoryKind.Footnote, "separator");
            List<DocxReferencedRelatedStoryLayout> pageFootnoteStories = [];
            foreach (int sourceBlockIndex in EnumeratePageSourceBlockIndexes(page))
            {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (DocxInlineReferenceLocation location in EnumerateInlineReferenceLocations(document.BodyElements, sourceBlockIndex))
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

                    if (!IsInlineReferenceRenderedOnPage(pages, pageIndex, location))
                    {
                        continue;
                    }

                    placedStoryKeys.Add((reference.Kind, reference.Id));
                    pageFootnoteStories.Add(new DocxReferencedRelatedStoryLayout(location, storyLayout));
                }
            }

            IReadOnlyList<DocxPlacedRelatedStoryLayout> placedStories = PlaceFootnoteStories(page, pageIndex, pageFootnoteStories, footnoteSeparatorLayout);
            pagesWithStories[pageIndex] = placedStories.Count == 0
                ? page
                : page with { PlacedRelatedStories = placedStories };
        }

        return AddPlacedEndnoteStories(pagesWithStories);

        IReadOnlyList<DocxLayoutPage> AddPlacedEndnoteStories(IReadOnlyList<DocxLayoutPage> pages)
        {
            List<DocxInlineReferenceLocation> endnoteLocations = ResolveReferencedRelatedStoryLocations(document, DocxRelatedStoryKind.Endnote, placedStoryKeys).ToList();
            if (endnoteLocations.Count == 0)
            {
                return pages;
            }

            var outputPages = pages.ToList();
            var documentEndLocations = new List<DocxInlineReferenceLocation>();
            var sectionEndLocations = new List<DocxInlineReferenceLocation>();
            foreach (DocxInlineReferenceLocation endnoteLocation in endnoteLocations)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (ResolveSectionEndEndnotePageIndex(document.BodyElements, outputPages, endnoteLocation) < 0)
                {
                    documentEndLocations.Add(endnoteLocation);
                    continue;
                }

                sectionEndLocations.Add(endnoteLocation);
            }

            foreach (IGrouping<(int StartBlockIndex, int EndBlockIndex), DocxInlineReferenceLocation> sectionGroup in sectionEndLocations
                         .GroupBy(location => ResolveSectionBlockRange(document.BodyElements, location.SourceBlockIndex))
                         .OrderBy(group => group.Key.StartBlockIndex))
            {
                cancellationToken.ThrowIfCancellationRequested();
                int sectionEndPageIndex = ResolveSectionEndEndnotePageIndex(document.BodyElements, outputPages, sectionGroup.First());
                if (sectionEndPageIndex < 0)
                {
                    documentEndLocations.AddRange(sectionGroup);
                    continue;
                }

                Dictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> endnoteStoryByKey = CreateRelatedStoryLookup(resolveRelatedStoryLayouts(ResolvePageBodyWidth(outputPages[sectionEndPageIndex])));
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
                    PlaceSectionEndEndnoteStories(outputPages, sectionEndPageIndex, sectionEndStories, cancellationToken);
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

            IReadOnlyList<DocxLayoutPage> AddDocumentEndnoteStories()
            {
                if (documentEndStories.Count == 0)
                {
                    return outputPages;
                }

                var documentEndPages = outputPages.ToList();
                DocxLayoutPage activePage = documentEndPages[^1];
                List<DocxPlacedRelatedStoryLayout> activePlacedStories = activePage.PlacedRelatedStories.ToList();
                double cursorTop = ResolveEndnoteStartTop(activePage, activePlacedStories);
                int activePageIndex = documentEndPages.Count - 1;
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
                        insertContinuationAfterActivePage: false);
                }

                return documentEndPages;
            }
        }
    }

    private static void PlaceSectionEndEndnoteStories(
        List<DocxLayoutPage> outputPages,
        int sectionEndPageIndex,
        IReadOnlyList<DocxReferencedRelatedStoryLayout> sectionEndStories,
        CancellationToken cancellationToken)
    {
        int activePageIndex = sectionEndPageIndex;
        DocxLayoutPage activePage = outputPages[activePageIndex];
        List<DocxPlacedRelatedStoryLayout> activePlacedStories = activePage.PlacedRelatedStories.ToList();
        double cursorTop = ResolveEndnoteStartTop(activePage, activePlacedStories);
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
                insertContinuationAfterActivePage: true);
        }
    }

    private static void PlaceRelatedStorySlices(
        List<DocxLayoutPage> outputPages,
        ref int activePageIndex,
        ref DocxLayoutPage activePage,
        ref List<DocxPlacedRelatedStoryLayout> activePlacedStories,
        ref double cursorTop,
        DocxRelatedStoryLayout storyLayout,
        int sourceBlockIndex,
        bool insertContinuationAfterActivePage)
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
                    insertContinuationAfterActivePage);
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
                    insertContinuationAfterActivePage);
            }
        }
    }

    private static void MoveToRelatedStoryContinuationPage(
        List<DocxLayoutPage> outputPages,
        ref int activePageIndex,
        ref DocxLayoutPage activePage,
        ref List<DocxPlacedRelatedStoryLayout> activePlacedStories,
        ref double cursorTop,
        bool insertContinuationAfterActivePage)
    {
        activePage = CreateEmptyContinuationPage(activePage);
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
        cursorTop = activePage.Height - activePage.MarginTop;

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

    private static IEnumerable<DocxInlineReferenceLocation> ResolveReferencedRelatedStoryLocations(
        DocxDocument document,
        DocxRelatedStoryKind kind,
        HashSet<(DocxRelatedStoryKind Kind, string Id)> placedStoryKeys)
    {
        for (int sourceBlockIndex = 0; sourceBlockIndex < document.BodyElements.Count; sourceBlockIndex++)
        {
            foreach (DocxInlineReferenceLocation location in EnumerateInlineReferenceLocations(document.BodyElements, sourceBlockIndex))
            {
                DocxInlineReference reference = location.Reference;
                if (reference.Kind != kind ||
                    reference.Id is null ||
                    !placedStoryKeys.Add((reference.Kind, reference.Id)))
                {
                    continue;
                }

                yield return location;
            }
        }
    }

    private static bool TryResolveReferencedRelatedStoryLayout(
        IReadOnlyDictionary<(DocxRelatedStoryKind Kind, string Id), DocxRelatedStoryLayout> storyByKey,
        DocxInlineReferenceLocation location,
        [NotNullWhen(true)] out DocxRelatedStoryLayout? storyLayout)
    {
        storyLayout = null;
        if (location.Reference.Id is null ||
            !storyByKey.TryGetValue((location.Reference.Kind, location.Reference.Id), out DocxRelatedStoryLayout? resolvedLayout) ||
            resolvedLayout is null)
        {
            return false;
        }

        storyLayout = resolvedLayout;
        return true;
    }

    private static int ResolveSectionEndEndnotePageIndex(
        IReadOnlyList<DocxBodyElement> elements,
        IReadOnlyList<DocxLayoutPage> pages,
        DocxInlineReferenceLocation location)
    {
        int referencePageIndex = FindInlineReferenceRenderedPageIndex();
        if (referencePageIndex < 0 ||
            !string.Equals(pages[referencePageIndex].PageSettings.EndnoteReferenceSettings.PositionValue, "sectEnd", StringComparison.OrdinalIgnoreCase))
        {
            return -1;
        }

        (int startBlockIndex, int endBlockIndex) = ResolveSectionBlockRange(elements, location.SourceBlockIndex);
        int sectionEndPageIndex = -1;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            if (EnumeratePageSourceBlockIndexes(pages[pageIndex]).Any(index => index >= startBlockIndex && index <= endBlockIndex))
            {
                sectionEndPageIndex = pageIndex;
            }
        }

        return sectionEndPageIndex >= 0 ? sectionEndPageIndex : referencePageIndex;

        int FindInlineReferenceRenderedPageIndex()
        {
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
            {
                if (IsInlineReferenceRenderedOnPage(pages, pageIndex, location))
                {
                    return pageIndex;
                }
            }

            return -1;
        }
    }

    private static (int StartBlockIndex, int EndBlockIndex) ResolveSectionBlockRange(IReadOnlyList<DocxBodyElement> elements, int sourceBlockIndex)
    {
        int startBlockIndex = 0;
        for (int index = Math.Min(sourceBlockIndex - 1, elements.Count - 1); index >= 0; index--)
        {
            if (elements[index] is DocxSectionBreakElement)
            {
                startBlockIndex = index + 1;
                break;
            }
        }

        int endBlockIndex = Math.Max(0, elements.Count - 1);
        for (int index = Math.Max(0, sourceBlockIndex); index < elements.Count; index++)
        {
            if (elements[index] is DocxSectionBreakElement)
            {
                endBlockIndex = index;
                break;
            }
        }

        return (startBlockIndex, endBlockIndex);
    }

    private static double ResolvePageBodyWidth(DocxLayoutPage page)
    {
        return Math.Max(1d, page.Width - page.MarginLeft - page.MarginRight);
    }

    private static double ResolvePlacedStoryHeight(DocxRelatedStoryLayout storyLayout, DocxLayoutPage page)
    {
        return Math.Min(Math.Max(0d, storyLayout.ContentHeight), Math.Max(0d, page.Height - page.MarginTop - page.MarginBottom));
    }

    private static double ResolveEndnoteStartTop(DocxLayoutPage page, IReadOnlyList<DocxPlacedRelatedStoryLayout> placedStories)
    {
        double bodyBottom = page.Items.Count == 0
            ? page.Height - page.MarginTop
            : page.Items.Min(item => GetVerticalBounds(item).Y);
        double placedStoryBottom = placedStories
            .Select(story => Math.Max(page.MarginBottom, story.TopY - story.Height))
            .DefaultIfEmpty(page.Height - page.MarginTop)
            .Min();
        return Math.Min(bodyBottom, placedStoryBottom) - FootnoteSeparatorGapPoints;
    }

    private static IReadOnlyList<DocxLayoutPage> ReindexPageOwnedLayouts(IReadOnlyList<DocxLayoutPage> pages)
    {
        return pages
            .Select((page, pageIndex) => page with
            {
                StaticInlineImages = ReindexInlineImages(page.StaticInlineImages, pageIndex),
                StaticTableRows = ReindexTableRows(page.StaticTableRows, pageIndex),
                PlacedRelatedStories = page.PlacedRelatedStories
                    .Select(story => ReindexPlacedRelatedStory(story, pageIndex))
                    .ToArray(),
                Items = page.Items
                    .Select(item => ReindexLayoutItem(item, pageIndex))
                    .ToArray()
            })
            .ToArray();
    }

    private static DocxLayoutItem ReindexLayoutItem(DocxLayoutItem item, int pageIndex)
    {
        return item switch
        {
            DocxInlineImageLayout image => image with { PageIndex = pageIndex },
            DocxTableRowLayout row => ReindexTableRow(row, pageIndex),
            _ => item
        };
    }

    private static DocxPlacedRelatedStoryLayout ReindexPlacedRelatedStory(DocxPlacedRelatedStoryLayout story, int pageIndex)
    {
        return story with
        {
            InlineImages = ReindexInlineImages(story.InlineImages, pageIndex),
            FloatingDrawings = story.FloatingDrawings
                .Select(drawing => drawing with
                {
                    PageStartIndex = pageIndex,
                    PageEndIndex = pageIndex,
                    AnchorPageIndex = pageIndex
                })
                .ToArray(),
            TableRows = ReindexTableRows(story.TableRows, pageIndex)
        };
    }

    private static IReadOnlyList<DocxTableRowLayout> ReindexTableRows(IReadOnlyList<DocxTableRowLayout> rows, int pageIndex)
    {
        return rows
            .Select(row => ReindexTableRow(row, pageIndex))
            .ToArray();
    }

    private static DocxTableRowLayout ReindexTableRow(DocxTableRowLayout row, int pageIndex)
    {
        return row with
        {
            Cells = row.Cells
                .Select(cell => ReindexTableCell(cell, pageIndex))
                .ToArray()
        };
    }

    private static DocxTableCellLayout ReindexTableCell(DocxTableCellLayout cell, int pageIndex)
    {
        return cell with
        {
            InlineImages = ReindexInlineImages(cell.InlineImages, pageIndex),
            NestedTableRows = ReindexTableRows(cell.NestedRows, pageIndex)
        };
    }

    private static IReadOnlyList<DocxInlineImageLayout> ReindexInlineImages(IReadOnlyList<DocxInlineImageLayout> images, int pageIndex)
    {
        return images
            .Select(image => image with { PageIndex = pageIndex })
            .ToArray();
    }

    private sealed record DocxInlineReferenceLocation(
        int SourceBlockIndex,
        DocxParagraph SourceParagraph,
        DocxInlineReference Reference);

    private sealed record DocxReferencedRelatedStoryLayout(
        DocxInlineReferenceLocation Location,
        DocxRelatedStoryLayout StoryLayout);

    private readonly record struct DocxPageTextLineOwner(
        DocxTextLineLayout Line,
        int? SourceBlockIndex);

    private static IEnumerable<int> EnumeratePageSourceBlockIndexes(DocxLayoutPage page)
    {
        return page.Items
            .Select(GetSourceBlockIndex)
            .Where(index => index is not null)
            .OfType<int>()
            .Distinct()
            .OrderBy(index => index);
    }

    private static IEnumerable<DocxInlineReferenceLocation> EnumerateInlineReferenceLocations(IReadOnlyList<DocxBodyElement> elements, int sourceBlockIndex)
    {
        if (sourceBlockIndex < 0 || sourceBlockIndex >= elements.Count)
        {
            yield break;
        }

        foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateDirectParagraphs([elements[sourceBlockIndex]]))
        {
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                yield return new DocxInlineReferenceLocation(sourceBlockIndex, paragraph, reference);
            }
        }

        if (elements[sourceBlockIndex] is DocxTableElement tableElement)
        {
            foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateTableParagraphs(tableElement.Table))
            {
                foreach (DocxInlineReference reference in paragraph.InlineReferences)
                {
                    yield return new DocxInlineReferenceLocation(sourceBlockIndex, paragraph, reference);
                }
            }
        }
    }

    private static bool IsInlineReferenceRenderedOnPage(
        IReadOnlyList<DocxLayoutPage> pages,
        int pageIndex,
        DocxInlineReferenceLocation location)
    {
        int sourceBlockIndex = location.SourceBlockIndex;
        DocxInlineReference reference = location.Reference;
        if (reference.SourceRunIndex < 0 ||
            !IsInlineReferenceRunRenderedAnywhere(reference.SourceRunIndex))
        {
            return true;
        }

        if (IsInlineReferenceOffsetRenderedAnywhere(reference.SourceRunIndex, reference.TextOffsetInRun))
        {
            return EnumeratePageTextLineOwners(pages[pageIndex])
                .Any(owner => TextLineMatchesInlineReferenceOwner(owner, sourceBlockIndex, location.SourceParagraph) &&
                    owner.Line.Segments.Any(segment => SegmentContainsSourceTextOffset(segment, reference.SourceRunIndex, reference.TextOffsetInRun)));
        }

        return EnumeratePageTextLineOwners(pages[pageIndex])
            .Any(owner => TextLineMatchesInlineReferenceOwner(owner, sourceBlockIndex, location.SourceParagraph) &&
                owner.Line.Segments.Any(segment => segment.SourceTextRunIndex == reference.SourceRunIndex));

        bool IsInlineReferenceRunRenderedAnywhere(int sourceRunIndex)
        {
            foreach (DocxLayoutPage page in pages)
            {
                foreach (DocxPageTextLineOwner owner in EnumeratePageTextLineOwners(page))
                {
                    if (TextLineMatchesInlineReferenceOwner(owner, location.SourceBlockIndex, location.SourceParagraph) &&
                        owner.Line.Segments.Any(segment => segment.SourceTextRunIndex == sourceRunIndex))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        bool IsInlineReferenceOffsetRenderedAnywhere(int sourceRunIndex, int textOffsetInRun)
        {
            foreach (DocxLayoutPage page in pages)
            {
                foreach (DocxPageTextLineOwner owner in EnumeratePageTextLineOwners(page))
                {
                    if (TextLineMatchesInlineReferenceOwner(owner, location.SourceBlockIndex, location.SourceParagraph) &&
                        owner.Line.Segments.Any(segment => SegmentContainsSourceTextOffset(segment, sourceRunIndex, textOffsetInRun)))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }

    private static bool SegmentContainsSourceTextOffset(DocxTextSegmentLayout segment, int sourceRunIndex, int textOffsetInRun)
    {
        if (segment.SourceTextRunIndex != sourceRunIndex)
        {
            return false;
        }

        int start = Math.Max(0, segment.SourceTextOffsetInRun);
        int end = start + segment.Text.Length;
        return textOffsetInRun >= start && textOffsetInRun < end;
    }

    private static bool TextLineMatchesInlineReferenceOwner(DocxPageTextLineOwner owner, int sourceBlockIndex, DocxParagraph sourceParagraph)
    {
        if (owner.SourceBlockIndex != sourceBlockIndex)
        {
            return false;
        }

        return owner.Line.SourceParagraph is null || ReferenceEquals(owner.Line.SourceParagraph, sourceParagraph);
    }

    private static IEnumerable<DocxPageTextLineOwner> EnumeratePageTextLineOwners(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            foreach (DocxPageTextLineOwner owner in EnumerateTextLineOwners(item, inheritedSourceBlockIndex: null))
            {
                yield return owner;
            }
        }
    }

    private static IEnumerable<DocxPageTextLineOwner> EnumerateTextLineOwners(DocxLayoutItem item, int? inheritedSourceBlockIndex)
    {
        switch (item)
        {
            case DocxTextLineLayout line:
                yield return new DocxPageTextLineOwner(line, line.SourceBlockIndex ?? inheritedSourceBlockIndex);
                break;
            case DocxTableRowLayout row:
                int? rowSourceBlockIndex = row.Table.SourceBlockIndex;
                foreach (DocxTableCellLayout cell in row.Cells)
                {
                    foreach (DocxTextLineLayout line in cell.TextLines)
                    {
                        yield return new DocxPageTextLineOwner(line, line.SourceBlockIndex ?? rowSourceBlockIndex);
                    }

                    foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
                    {
                        foreach (DocxPageTextLineOwner owner in EnumerateTextLineOwners(nestedRow, rowSourceBlockIndex))
                        {
                            yield return owner;
                        }
                    }
                }

                break;
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumeratePageTextLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            foreach (DocxTextLineLayout line in EnumerateTextLines(item))
            {
                yield return line;
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateTextLines(DocxLayoutItem item)
    {
        switch (item)
        {
            case DocxTextLineLayout line:
                yield return line;
                break;
            case DocxTableRowLayout row:
                foreach (DocxTableCellLayout cell in row.Cells)
                {
                    foreach (DocxTextLineLayout line in cell.TextLines)
                    {
                        yield return line;
                    }

                    foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
                    {
                        foreach (DocxTextLineLayout line in EnumerateTextLines(nestedRow))
                        {
                            yield return line;
                        }
                    }
                }

                break;
        }
    }

    private static DocxPlacedRelatedStoryLayout PlaceRelatedStory(DocxLayoutPage page, int pageIndex, DocxRelatedStoryLayout storyLayout, int sourceBlockIndex)
    {
        double storyHeight = Math.Min(Math.Max(0d, storyLayout.ContentHeight), Math.Max(0d, page.Height - page.MarginTop - page.MarginBottom));
        double topY = page.MarginBottom + storyHeight;
        return PlaceRelatedStoryAtTop(page, pageIndex, storyLayout, sourceBlockIndex, topY, topY + FootnoteSeparatorGapPoints);
    }

    private static IReadOnlyList<DocxPlacedRelatedStoryLayout> PlaceFootnoteStories(
        DocxLayoutPage page,
        int pageIndex,
        IReadOnlyList<DocxReferencedRelatedStoryLayout> footnoteStories,
        DocxRelatedStoryLayout? separatorLayout)
    {
        if (footnoteStories.Count == 0)
        {
            return [];
        }

        double bodyHeight = footnoteStories.Sum(story => ResolvePlacedStoryHeight(story.StoryLayout, page));
        double cursorTop = page.MarginBottom + bodyHeight;
        var placedStories = new List<DocxPlacedRelatedStoryLayout>(footnoteStories.Count + (separatorLayout is null ? 0 : 1));
        if (separatorLayout is not null)
        {
            double separatorHeight = ResolvePlacedStoryHeight(separatorLayout, page);
            double separatorTop = cursorTop + FootnoteSeparatorGapPoints + separatorHeight;
            placedStories.Add(PlaceRelatedStoryAtTop(page, pageIndex, separatorLayout, footnoteStories[0].Location.SourceBlockIndex, separatorTop, separatorY: null));
        }

        bool firstStory = true;
        foreach (DocxReferencedRelatedStoryLayout story in footnoteStories)
        {
            double? separatorY = separatorLayout is null && firstStory
                ? cursorTop + FootnoteSeparatorGapPoints
                : null;
            placedStories.Add(PlaceRelatedStoryAtTop(page, pageIndex, story.StoryLayout, story.Location.SourceBlockIndex, cursorTop, separatorY));
            cursorTop -= ResolvePlacedStoryHeight(story.StoryLayout, page);
            firstStory = false;
        }

        return placedStories.ToArray();
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
            if (story.Story.Id is not null && IsNormalRelatedStory(story.Story))
            {
                storyByKey.TryAdd((story.Story.Kind, story.Story.Id), story);
            }
        }

        return storyByKey;
    }

    private static bool IsNormalRelatedStory(DocxRelatedStory story)
    {
        return string.IsNullOrEmpty(story.Type) ||
            string.Equals(story.Type, "normal", StringComparison.OrdinalIgnoreCase);
    }

    private static DocxRelatedStoryLayout? FindSpecialRelatedStoryLayout(
        IReadOnlyList<DocxRelatedStoryLayout> relatedStoryLayouts,
        DocxRelatedStoryKind kind,
        string type)
    {
        return relatedStoryLayouts.FirstOrDefault(story =>
            story.Story.Id is not null &&
            story.ContentHeight > 0d &&
            story.Story.Kind == kind &&
            string.Equals(story.Story.Type, type, StringComparison.OrdinalIgnoreCase));
    }

    private static IReadOnlyList<DocxRelatedStoryLayout> CreateRelatedStoryLayouts(
        IReadOnlyList<DocxRelatedStory> stories,
        double bodyWidth,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        double paragraphSpacingScale,
        CancellationToken cancellationToken)
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
                cancellationToken: cancellationToken);
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
        CancellationToken cancellationToken)
    {
        var textLines = new List<DocxTextLineLayout>();
        var inlineImages = new List<DocxInlineImageLayout>();
        var tableRows = new List<DocxTableRowLayout>();
        double cursorY = 0d;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
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
                        StoryKind: null,
                        StoryVariantType: null,
                        pageCount: pageCount,
                        paragraphSpacingScale: paragraphSpacingScale));
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
            IReadOnlyList<DocxTextLineLayout> paragraphLines = LayoutRelatedStoryParagraphTextLines(
                paragraph,
                elementIndex,
                paragraphIndex,
                story.Kind.ToValueString(),
                bodyWidth,
                cursorY,
                spacingProfile,
                textMeasurer,
                defaultTabStopPoints,
                pageNumber,
                pageCount);
            textLines.AddRange(paragraphLines);
            cursorY -= paragraphLines.Sum(line => line.LineHeight ?? 0d);
            if (paragraphLines.Count == 0 && paragraph.Images.Count == 0)
            {
                double fontSize = GetParagraphFontSize(paragraph);
                cursorY -= ResolveLineHeight(paragraph, fontSize, textMeasurer);
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                cancellationToken.ThrowIfCancellationRequested();
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
                    SourceParagraphIndex: paragraphIndex, StoryKind: null, StoryVariantType: null));
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
            pageCount);
        return new DocxRelatedStoryLayout(story, storyIndex, textLines.ToArray(), inlineImages.ToArray(), floatingDrawings, tableRows.ToArray(), Math.Abs(cursorY));
    }
}
