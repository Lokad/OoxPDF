namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Index body items by source block and page. Placing notes on an existing
    // page leaves the body items intact. Inserting continuation pages changes
    // page indexes and requires a rebuild.
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

        public int FindFirstPageWithReference(DocxInlineReferenceLocation location, bool requireRendered = false)
        {
            if (sortedBlocksByPage.Length == 0)
            {
                return -1;
            }

            int sourceRunIndex = location.Reference.SourceRunIndex;
            if (sourceRunIndex < 0 || !IsRunRenderedAnywhere(location, sourceRunIndex))
            {
                // Legacy placement can fall back to the first page. A field
                // binding needs an observed reference rather than that guess.
                return requireRendered ? -1 : 0;
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
