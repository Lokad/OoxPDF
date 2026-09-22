namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // R12: reference-to-rendered-page index built once per placement pass. Matching
    // used to re-walk every page item tree for each (page, reference) check, and the
    // section-end search rebuilt sorted block sets per page per endnote. The index
    // materializes per-page text-line owners and sorted source-block sets a single
    // time; queries reuse them with the exact legacy predicates. Callers rebuild the
    // index whenever the page list changes (endnote placement mutates pages between
    // section groups), so no stale repagination data applies.
    internal sealed class RelatedStoryPageIndex
    {
        private readonly IReadOnlyList<IReadOnlyList<DocxPageTextLineOwner>> ownersByPage;
        private readonly int[][] sortedBlocksByPage;

        private RelatedStoryPageIndex(
            IReadOnlyList<IReadOnlyList<DocxPageTextLineOwner>> ownersByPage,
            int[][] sortedBlocksByPage)
        {
            this.ownersByPage = ownersByPage;
            this.sortedBlocksByPage = sortedBlocksByPage;
        }

        public int PageCount => ownersByPage.Count;

        public static RelatedStoryPageIndex Build(IReadOnlyList<DocxLayoutPage> pages, CancellationToken cancellationToken)
        {
            var owners = new List<DocxPageTextLineOwner>[pages.Count];
            var blocks = new int[pages.Count][];
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                owners[pageIndex] = EnumeratePageTextLineOwners(pages[pageIndex]).ToList();
                blocks[pageIndex] = EnumeratePageSourceBlockIndexes(pages[pageIndex]).ToArray();
            }

            return new RelatedStoryPageIndex(owners, blocks);
        }

        public int[] SortedBlocks(int pageIndex)
        {
            return sortedBlocksByPage[pageIndex];
        }

        public bool IsReferenceRenderedOnPage(int pageIndex, DocxInlineReferenceLocation location)
        {
            return IsInlineReferenceRenderedOnPage(ownersByPage, pageIndex, location);
        }

        public int FindFirstPageWithReference(DocxInlineReferenceLocation location)
        {
            for (int pageIndex = 0; pageIndex < ownersByPage.Count; pageIndex++)
            {
                if (IsInlineReferenceRenderedOnPage(ownersByPage, pageIndex, location))
                {
                    return pageIndex;
                }
            }

            return -1;
        }

        public int FindLastPageWithBlockInRange(int startBlockIndex, int endBlockIndex)
        {
            int found = -1;
            for (int pageIndex = 0; pageIndex < sortedBlocksByPage.Length; pageIndex++)
            {
                foreach (int blockIndex in sortedBlocksByPage[pageIndex])
                {
                    if (blockIndex >= startBlockIndex && blockIndex <= endBlockIndex)
                    {
                        found = pageIndex;
                        break;
                    }
                }
            }

            return found;
        }
    }
}
