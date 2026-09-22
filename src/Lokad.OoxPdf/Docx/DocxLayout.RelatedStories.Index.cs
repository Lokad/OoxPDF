namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // R12: reference-to-rendered-page index built once per placement pass. Matching
    // used to re-walk every page item tree for each (page, reference) check (and the
    // per-check "anywhere" helpers re-scanned all pages on top), and the section-end
    // search rebuilt sorted block sets per page per endnote. The index materializes
    // per-block segment match records plus sorted source-block sets a single time;
    // queries scan one block record list with the exact legacy predicates instead of
    // every page owner list. Callers rebuild the index whenever the page list changes
    // (endnote placement mutates pages between section groups), so no stale
    // repagination data applies.
    internal sealed class RelatedStoryPageIndex
    {
        private readonly Dictionary<int, List<BlockMatchRecord>> recordsByBlock;
        private readonly int[][] sortedBlocksByPage;

        private readonly record struct BlockMatchRecord(
            DocxParagraph? LineParagraph,
            int RunIndex,
            int SegmentStart,
            int SegmentEnd,
            int PageIndex);

        private RelatedStoryPageIndex(
            Dictionary<int, List<BlockMatchRecord>> recordsByBlock,
            int[][] sortedBlocksByPage)
        {
            this.recordsByBlock = recordsByBlock;
            this.sortedBlocksByPage = sortedBlocksByPage;
        }

        public int PageCount => sortedBlocksByPage.Length;

        public static RelatedStoryPageIndex Build(IReadOnlyList<DocxLayoutPage> pages, CancellationToken cancellationToken)
        {
            var recordsByBlock = new Dictionary<int, List<BlockMatchRecord>>();
            var blocks = new int[pages.Count][];
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                blocks[pageIndex] = EnumeratePageSourceBlockIndexes(pages[pageIndex]).ToArray();
                foreach (DocxPageTextLineOwner owner in EnumeratePageTextLineOwners(pages[pageIndex]))
                {
                    if (owner.SourceBlockIndex is not int sourceBlockIndex)
                    {
                        continue;
                    }

                    foreach (DocxTextSegmentLayout segment in owner.Line.Segments)
                    {
                        int start = Math.Max(0, segment.SourceTextOffsetInRun);
                        if (!recordsByBlock.TryGetValue(sourceBlockIndex, out List<BlockMatchRecord>? records))
                        {
                            records = new List<BlockMatchRecord>();
                            recordsByBlock[sourceBlockIndex] = records;
                        }

                        records.Add(new BlockMatchRecord(
                            owner.Line.SourceParagraph,
                            segment.SourceTextRunIndex,
                            start,
                            start + segment.Text.Length,
                            pageIndex));
                    }
                }
            }

            return new RelatedStoryPageIndex(recordsByBlock, blocks);
        }

        public int[] SortedBlocks(int pageIndex)
        {
            return sortedBlocksByPage[pageIndex];
        }

        public bool IsReferenceRenderedOnPage(int pageIndex, DocxInlineReferenceLocation location)
        {
            int sourceRunIndex = location.Reference.SourceRunIndex;
            if (sourceRunIndex < 0 || !IsRunRenderedAnywhere(location, sourceRunIndex))
            {
                return true;
            }

            int textOffsetInRun = location.Reference.TextOffsetInRun;
            bool requireOffset = IsOffsetRenderedAnywhere(location, sourceRunIndex, textOffsetInRun);
            return MatchRecords(location, record =>
                record.PageIndex == pageIndex &&
                record.RunIndex == sourceRunIndex &&
                (!requireOffset || (record.SegmentStart <= textOffsetInRun && textOffsetInRun < record.SegmentEnd)));
        }

        public int FindFirstPageWithReference(DocxInlineReferenceLocation location)
        {
            if (sortedBlocksByPage.Length == 0)
            {
                return -1;
            }

            int sourceRunIndex = location.Reference.SourceRunIndex;
            if (sourceRunIndex < 0 || !IsRunRenderedAnywhere(location, sourceRunIndex))
            {
                return 0;
            }

            int textOffsetInRun = location.Reference.TextOffsetInRun;
            bool requireOffset = IsOffsetRenderedAnywhere(location, sourceRunIndex, textOffsetInRun);
            int found = -1;
            if (recordsByBlock.TryGetValue(location.SourceBlockIndex, out List<BlockMatchRecord>? records))
            {
                foreach (BlockMatchRecord record in records)
                {
                    if ((record.LineParagraph is null || ReferenceEquals(record.LineParagraph, location.SourceParagraph)) &&
                        record.RunIndex == sourceRunIndex &&
                        (!requireOffset || (record.SegmentStart <= textOffsetInRun && textOffsetInRun < record.SegmentEnd)) &&
                        (found < 0 || record.PageIndex < found))
                    {
                        found = record.PageIndex;
                    }
                }
            }

            return found;
        }

        private bool IsRunRenderedAnywhere(DocxInlineReferenceLocation location, int sourceRunIndex)
        {
            return MatchRecords(location, record => record.RunIndex == sourceRunIndex);
        }

        private bool IsOffsetRenderedAnywhere(DocxInlineReferenceLocation location, int sourceRunIndex, int textOffsetInRun)
        {
            return MatchRecords(location, record =>
                record.RunIndex == sourceRunIndex &&
                record.SegmentStart <= textOffsetInRun && textOffsetInRun < record.SegmentEnd);
        }

        private bool MatchRecords(DocxInlineReferenceLocation location, Func<BlockMatchRecord, bool> match)
        {
            if (!recordsByBlock.TryGetValue(location.SourceBlockIndex, out List<BlockMatchRecord>? records))
            {
                return false;
            }

            foreach (BlockMatchRecord record in records)
            {
                if ((record.LineParagraph is null || ReferenceEquals(record.LineParagraph, location.SourceParagraph)) &&
                    match(record))
                {
                    return true;
                }
            }

            return false;
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
