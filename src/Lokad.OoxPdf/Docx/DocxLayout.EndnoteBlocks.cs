namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static bool TryPlaceCompleteEndnoteBlocks(
        List<DocxLayoutPage> outputPages, ref int pageIndex, ref DocxLayoutPage page,
        ref List<DocxPlacedRelatedStoryLayout> placed, ref double top,
        DocxRelatedStoryLayout story, int sourceBlockIndex, bool insertAfterPage,
        DocxRelatedStoryLayout? continuation, IDocxTextMeasurer? measurer, double printScale,
        Func<DocxLayoutPage, int, int, double>? headerKeepOut,
        Func<DocxLayoutPage, int, int, double>? frameBottom, CancellationToken token)
    {
        if (story.Story.Kind != DocxRelatedStoryKind.Endnote || !story.Story.IsNormalStoryType ||
            story.TableRows.Count == 0 || printScale != 1d ||
            story.ContentHeight <= top - (frameBottom?.Invoke(page, pageIndex + 1, outputPages.Count) ?? page.MarginBottom) ||
            !HasCompleteMeasuredNoteBlocks(story, token)) return false;

        double continuationCharge = 0d;
        if (continuation is { TextLines.Count: > 0 })
        {
            var (mark, size) = FindSeparatorMarkFont(continuation.TextLines);
            continuationCharge = Math.Max(0d, -continuation.TextLines[0].BaselineY) +
                ResolveSeparatorGapPoints(mark, size, measurer);
        }
        // Complete rows are atomic. Rows larger than a fresh note frame retain
        // their existing diagnosed/approximate path; never cut through glyphs.
        double fullHeight = page.Height - page.MarginTop - page.MarginBottom - continuationCharge;
        var blocks = new List<(DocxRelatedStoryLayout Layout, bool Row)>();
        int lineIndex = 0;
        double offset = 0d;
        foreach (var row in story.TableRows)
        {
            token.ThrowIfCancellationRequested();
            double rowTop = -row.Y - row.Height;
            if (!double.IsFinite(rowTop) || !double.IsFinite(row.Height) || row.Height <= 0d ||
                row.Height > fullHeight || rowTop < offset - .001d) return false;
            AddTextBefore(row.Table.SourceBlockIndex, rowTop);
            blocks.Add((story with { TextLines = [], InlineImages = [],
                TableRows = ShiftTableRows([row], rowTop, 0d), ContentHeight = row.Height }, true));
            offset = rowTop + row.Height;
        }
        AddTextBefore(int.MaxValue, story.ContentHeight);
        if (lineIndex != story.TextLines.Count || offset > story.ContentHeight + .001d) return false;

        // Preflight every source block before changing pages. Existing line
        // slicing handles paragraph continuations; only complete rows defer as
        // units. The canonical story and measured row/cell geometry remain owned.
        foreach (var block in blocks)
        {
            token.ThrowIfCancellationRequested();
            if (!block.Row)
            {
                PlaceRelatedStorySlices(outputPages, ref pageIndex, ref page, ref placed, ref top,
                    block.Layout, sourceBlockIndex, insertAfterPage, continuation, measurer, printScale,
                    headerKeepOut, frameBottom: frameBottom, cancellationToken: token);
                top += FootnoteSeparatorGapPoints; // This is one story, not a new note boundary.
                continue;
            }
            double bottom = frameBottom?.Invoke(page, pageIndex + 1, outputPages.Count) ?? page.MarginBottom;
            if (block.Layout.ContentHeight > top - bottom + .001d)
            {
                MoveToRelatedStoryContinuationPage(outputPages, ref pageIndex, ref page, ref placed,
                    ref top, insertAfterPage, headerKeepOut);
                top = PlaceContinuationSeparatorIfNeeded(outputPages, ref pageIndex, ref page, ref placed,
                    top, sourceBlockIndex, story, 0, [], continuation, measurer, printScale,
                    frameBottom: frameBottom, hasPlacedSource: true);
            }
            bottom = frameBottom?.Invoke(page, pageIndex + 1, outputPages.Count) ?? page.MarginBottom;
            // A changed header/footer can make the actual fresh frame smaller
            // than the measured row. Keep that unfit row's bounded clip fallback.
            placed.Add(PlaceRelatedStoryAtTop(page, pageIndex, block.Layout, sourceBlockIndex, top,
                storyTopOffset: 0d, storyHeight: Math.Max(0d, Math.Min(block.Layout.ContentHeight, top - bottom)),
                separatorY: null));
            outputPages[pageIndex] = page with { PlacedRelatedStories = placed.ToArray() };
            top -= block.Layout.ContentHeight;
        }
        top -= FootnoteSeparatorGapPoints;
        return true;

        void AddTextBefore(int nextBlock, double end)
        {
            int first = lineIndex;
            while (lineIndex < story.TextLines.Count && story.TextLines[lineIndex].SourceBlockIndex < nextBlock)
            { token.ThrowIfCancellationRequested(); lineIndex++; }
            if (lineIndex == first && end <= offset + .001d) return;
            var text = NarrowStoryTextLines(story with { TableRows = [] }, first, lineIndex - first);
            blocks.Add((text with { TextLines = ShiftTextLines(text.TextLines, offset, 0d),
                InlineImages = ShiftInlineImages(text.InlineImages, offset, 0d), ContentHeight = end - offset }, false));
        }
    }
}
