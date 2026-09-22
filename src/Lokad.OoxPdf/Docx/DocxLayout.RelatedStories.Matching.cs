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

    // R12: reference and section-end page searches run against the once-per-group
    // page index instead of re-walking page item trees per location.
    private static int ResolveSectionEndEndnotePageIndex(
        IReadOnlyList<DocxBodyElement> elements,
        RelatedStoryPageIndex referenceIndex,
        IReadOnlyList<DocxLayoutPage> pages,
        DocxInlineReferenceLocation location)
    {
        int referencePageIndex = referenceIndex.FindFirstPageWithReference(location);
        if (referencePageIndex < 0 ||
            !string.Equals(pages[referencePageIndex].PageSettings.EndnoteReferenceSettings.PositionValue, "sectEnd", StringComparison.OrdinalIgnoreCase))
        {
            return -1;
        }

        (int startBlockIndex, int endBlockIndex) = ResolveSectionBlockRange(elements, location.SourceBlockIndex);
        int sectionEndPageIndex = referenceIndex.FindLastPageWithBlockInRange(startBlockIndex, endBlockIndex);
        return sectionEndPageIndex >= 0 ? sectionEndPageIndex : referencePageIndex;
    }

    private static Dictionary<int, List<DocxInlineReferenceLocation>> BuildInlineReferenceLocationsByBlock(
        DocxDocument document,
        CancellationToken cancellationToken)
    {
        // R12: reference locations derive from the immutable document, so memoize them
        // by source block instead of re-walking (table) paragraphs on every fragment page.
        var locationsByBlock = new Dictionary<int, List<DocxInlineReferenceLocation>>();
        for (int sourceBlockIndex = 0; sourceBlockIndex < document.BodyElements.Count; sourceBlockIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            List<DocxInlineReferenceLocation> blockLocations = EnumerateInlineReferenceLocations(document.BodyElements, sourceBlockIndex).ToList();
            if (blockLocations.Count != 0)
            {
                locationsByBlock[sourceBlockIndex] = blockLocations;
            }
        }

        return locationsByBlock;
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
}
