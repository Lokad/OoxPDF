namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Counter inputs belong to the rendered reference, independently of whether
    // this container ends a section or the document. Typography corrections do
    // not own PAGE. Each admitted source block must fit before the selected
    // footer edge; an unrelated later overflow cannot revoke its context.
    private static int CountReferenceBoundEndnotes(DocxLayoutPage page,
        IReadOnlyList<DocxPlacedRelatedStoryLayout> placed, DocxRelatedStoryLayout? separator,
        double printScale, IReadOnlyList<DocxRelatedStoryLayout> stories, CancellationToken token,
        double frameBottom)
    {
        token.ThrowIfCancellationRequested();
        if (CanUsePlainDocumentEndnoteFlow(page, separator, printScale, stories, token, out _)) return stories.Count;
        // The fitting proof already maps scaled body coordinates back into note
        // design space. Print scale must not revoke an otherwise complete result.
        return page.ColumnFrames.Count <= 1
            ? CountCompleteFittingEndnoteBlocks(page, placed, separator, printScale, stories, token, frameBottom) : 0;
    }

    // A complete measured block owns paragraph boundaries even when its middle
    // contains supported inline objects or complete nested tables. This admission does
    // not change typography or the separately qualified continuation policies.
    private static bool CanUseFittingEndnoteBoundaries(DocxDocument document, DocxLayoutPage page,
        IReadOnlyList<DocxPlacedRelatedStoryLayout> placed, DocxRelatedStoryLayout? separator,
        double printScale, IReadOnlyList<DocxRelatedStoryLayout> stories, CancellationToken token,
        bool hasCompleteStaticFrame, double frameBottom)
    {
        token.ThrowIfCancellationRequested();
        if (!CanUseUnscaledEndnoteParagraphBoundaries(document, page, separator, printScale, hasCompleteStaticFrame) ||
            !HasNumericTrailingBodyAfterSpacing(page)) return false;
        if (!HasCompleteFittingEndnoteBlocks(page, placed, separator, printScale, stories, token, frameBottom)) return false;
        foreach (var story in stories)
        {
            token.ThrowIfCancellationRequested();
            if (story.Story.BodyElements.Count == 0 || story.TextLines.Count == 0 ||
                story.Story.BodyElements[0] is not DocxParagraphElement first ||
                story.Story.BodyElements[^1] is not DocxParagraphElement last ||
                !ReferenceEquals(story.TextLines[0].SourceParagraph, first.Paragraph) ||
                !ReferenceEquals(story.TextLines[^1].SourceParagraph, last.Paragraph)) return false;
        }
        return true;
    }

    private static bool HasCompleteFittingEndnoteBlocks(DocxLayoutPage page,
        IReadOnlyList<DocxPlacedRelatedStoryLayout> placed, DocxRelatedStoryLayout? separator,
        double printScale, IReadOnlyList<DocxRelatedStoryLayout> stories, CancellationToken token, double frameBottom) =>
        CountCompleteFittingEndnoteBlocks(page, placed, separator, printScale, stories, token, frameBottom) == stories.Count &&
        stories.Count != 0;

    // Counter binding consumes the complete fitting prefix. Whole-group layout
    // keeps the stricter all-blocks result through the predicate above. One
    // source-order pass uses the same conservative charges and ownership proof.
    private static int CountCompleteFittingEndnoteBlocks(DocxLayoutPage page,
        IReadOnlyList<DocxPlacedRelatedStoryLayout> placed, DocxRelatedStoryLayout? separator,
        double printScale, IReadOnlyList<DocxRelatedStoryLayout> stories, CancellationToken token, double frameBottom)
    {
        token.ThrowIfCancellationRequested();
        frameBottom = Math.Max(page.MarginBottom, frameBottom);
        if (stories.Count == 0 || !double.IsFinite(frameBottom) ||
            separator?.FloatingDrawings.Count > 0 ||
            placed.Any(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)) return 0;
        double available = ResolveEndnoteStartTop(page, placed, printScale) - frameBottom;
        double height = separator?.ContentHeight ?? 0d;
        if (!double.IsFinite(height) || height < 0d || !double.IsFinite(available)) return 0;
        int count = 0;
        foreach (var story in stories)
        {
            token.ThrowIfCancellationRequested();
            if (!double.IsFinite(story.ContentHeight) || story.ContentHeight <= 0d || story.FloatingDrawings.Count != 0) break;
            // Keep the conservative legacy charges for admission. Placement's
            // separately qualified paragraph boundaries may remove them later.
            height += story.ContentHeight + FootnoteSeparatorGapPoints;
            if (!double.IsFinite(height) || height > available || !HasCompleteMeasuredNoteBlocks(story, token)) break;
            count++;
        }
        return count;
    }

    private static bool CanUseUnscaledEndnoteParagraphBoundaries(DocxDocument document,
        DocxLayoutPage page, DocxRelatedStoryLayout? separator, double printScale, bool hasCompleteStaticFrame) =>
        printScale == 1d && page.ColumnFrames.Count <= 1 &&
        DocxPlainBodyFrame.HasKnownBodyFrame(document) && hasCompleteStaticFrame &&
        HasEmptyNoteSeparator(separator);

    // The first complete source block owns the note-group start boundary.
    // Later notes may need continuation pages without adding the separator's
    // latent paragraph spacing back above this already measured block.
    // Counter binding and complete footer flow retain separate admission.
    private static bool CanUseMeasuredEndnoteStart(DocxDocument document, DocxLayoutPage page,
        IReadOnlyList<DocxPlacedRelatedStoryLayout> placed, DocxRelatedStoryLayout? separator,
        double printScale, IReadOnlyList<DocxRelatedStoryLayout> stories, CancellationToken token,
        bool hasCompleteStaticFrame)
    {
        token.ThrowIfCancellationRequested();
        if (!CanUseUnscaledEndnoteParagraphBoundaries(document, page, separator, printScale, hasCompleteStaticFrame) ||
            !HasNumericTrailingBodyAfterSpacing(page) || stories.Count == 0 ||
            placed.Any(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)) return false;
        var first = stories[0];
        return first.Story.BodyElements.Count != 0 && first.TextLines.Count != 0 &&
            first.Story.BodyElements[0] is DocxParagraphElement paragraph &&
            ReferenceEquals(first.TextLines[0].SourceParagraph, paragraph.Paragraph) &&
            HasCompleteMeasuredNoteBlocks(first, token);
    }

    // Adjacent source paragraphs own their shared boundary even when a
    // complete table or inline object precedes the trailing paragraph. The
    // placement caller still admits the next head against the actual page.
    private static bool HasMeasuredEndnoteParagraphBoundary(DocxRelatedStoryLayout previous,
        DocxRelatedStoryLayout current, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (HasPlainNoteFlowStory(previous) && HasPlainNoteFlowStory(current)) return true;
        return previous.Story.BodyElements.Count != 0 && current.Story.BodyElements.Count != 0 &&
            previous.TextLines.Count != 0 && current.TextLines.Count != 0 &&
            previous.Story.BodyElements[^1] is DocxParagraphElement tail &&
            current.Story.BodyElements[0] is DocxParagraphElement head &&
            ReferenceEquals(previous.TextLines[^1].SourceParagraph, tail.Paragraph) &&
            ReferenceEquals(current.TextLines[0].SourceParagraph, head.Paragraph) &&
            HasCompleteMeasuredNoteBlocks(previous, token) && HasCompleteMeasuredNoteBlocks(current, token);
    }

    // The container supplies its actual cursor and lower edge. Source ownership
    // qualifies spacing; only an admitted next head can keep the overlap here.
    private static (double Gap, double Overlap) ResolveEndnoteParagraphBoundary(
        DocxRelatedStoryLayout? previous, DocxRelatedStoryLayout current, double cursorTop, double frameBottom,
        bool completeBoundaries, bool plainFlow, bool allowLocalBoundary, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        bool local = !completeBoundaries && !plainFlow && allowLocalBoundary && previous is not null &&
            HasMeasuredEndnoteParagraphBoundary(previous, current, token);
        double overlap = (completeBoundaries || plainFlow || local) && previous is not null
            ? ResolveEndnoteBoundaryOverlap(previous, current) : 0d;
        double gap = local ? FootnoteSeparatorGapPoints : 0d;
        if (!completeBoundaries && (plainFlow || local) &&
            cursorTop + gap + overlap - ResolveStoryTextLineBoxHeight(current.TextLines[0],
                current.TextLines.Count == 1 ? current.TextLines[0].ParagraphAfterSpacing ?? 0d : 0d) < frameBottom)
            overlap = 0d;
        return (gap, overlap);
    }

    // Fitting-boundary collapse and complete-block admission use the same source
    // and occupied-frame proof. A measured first line alone cannot admit a cell
    // whose remaining content is clipped or whose source owner is unsupported.
    private static bool HasCompleteMeasuredNoteBlocks(DocxRelatedStoryLayout story, CancellationToken token)
    {
        if (!story.HasResolvedCounterInputs || story.Story.BodyElements.Count == 0 || story.FloatingDrawings.Count != 0 ||
            !double.IsFinite(story.ContentHeight) || story.ContentHeight <= 0d) return false;
        var measured = new Dictionary<DocxParagraph, bool>(ReferenceEqualityComparer.Instance);
        var sources = new HashSet<DocxParagraph>(ReferenceEqualityComparer.Instance);
        bool Index(DocxTextLineLayout line)
        {
            token.ThrowIfCancellationRequested();
            if (line.SourceParagraph is not { } paragraph || !double.IsFinite(line.BaselineY)) return false;
            bool frame = DocxInlineLineGeometry.TryGetOccupiedBounds(line, out _, out _);
            measured[paragraph] = frame && (!measured.TryGetValue(paragraph, out bool previous) || previous);
            return true;
        }
        bool Paragraph(DocxParagraph paragraph)
        {
            token.ThrowIfCancellationRequested();
            return sources.Add(paragraph) && measured.TryGetValue(paragraph, out bool frame) &&
                HasNoteParagraphFlowProperties(paragraph, allowSourceAffinedImages: true) &&
                (IsPlainNoteFlowParagraph(paragraph) || DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph)) &&
                (paragraph.Images.Count == 0 && paragraph.Equations.Count == 0 || frame);
        }
        foreach (var line in story.TextLines) if (!Index(line)) return false;
        // A surviving first line cannot stand in for a partly clipped
        // paragraph. Every measured line must survive all ancestor clips.
        if (!DocxTableCellClipGeometry.HasCompleteMeasuredRows(story.TableRows, token, (line, _) => Index(line))) return false;
        bool Elements(IReadOnlyList<DocxBodyElement> elements)
        {
            foreach (var element in elements)
            {
                token.ThrowIfCancellationRequested();
                if (element.Revisions.Count != 0) return false;
                if (element is DocxParagraphElement paragraph)
                {
                    if (!Paragraph(paragraph.Paragraph)) return false;
                }
                else if (element is DocxTableElement table)
                {
                    if (table.Table.Revisions.Count != 0) return false;
                    foreach (var row in table.Table.Rows)
                    {
                        token.ThrowIfCancellationRequested();
                        if (!HasOnlyPropertyHistory(row.Revisions, DocxRevisionKind.TableRowPropertiesChange,
                            DocxRevisionPropertyFamily.Row, "trPrChange") || !DocxMathParagraphPolicy.IsSupportedSimpleRow(row)) return false;
                        foreach (var cell in row.Cells)
                        {
                            token.ThrowIfCancellationRequested();
                            if (!HasOnlyPropertyHistory(cell.Revisions, DocxRevisionKind.TableCellPropertiesChange,
                                DocxRevisionPropertyFamily.Cell, "tcPrChange") || !DocxMathParagraphPolicy.IsSupportedCellProperties(cell) ||
                                !Elements(DocxTableCellContent.GetBodyElements(cell))) return false;
                        }
                    }
                }
                else return false;
            }
            return true;
        }
        return Elements(story.Story.BodyElements);

        // The reader has already resolved current properties. One known snapshot
        // does not replace the measured row/cell or its source content. Ambiguous
        // history and content/structural revisions retain the existing fallback.
        static bool HasOnlyPropertyHistory(IReadOnlyList<DocxRevisionInfo> revisions,
            DocxRevisionKind kind, DocxRevisionPropertyFamily family, string element) =>
            revisions.Count == 0 || revisions is [var revision] && revision.Kind == kind &&
                revision.PropertyChangeFamily == family && revision.SourceElement == element;
    }

}
