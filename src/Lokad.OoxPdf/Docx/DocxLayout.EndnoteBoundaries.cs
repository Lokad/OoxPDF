namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Counter inputs belong to the rendered reference, independently of whether
    // this container ends a section or the document. Typography corrections do
    // not own PAGE. Mixed notes still require a complete measured group inside
    // the available frame, including the actual selected footer edge.
    private static bool CanUseReferenceBoundEndnoteGroup(DocxLayoutPage page,
        IReadOnlyList<DocxPlacedRelatedStoryLayout> placed, DocxRelatedStoryLayout? separator,
        double printScale, IReadOnlyList<DocxRelatedStoryLayout> stories, CancellationToken token,
        double frameBottom) =>
        CanUsePlainDocumentEndnoteFlow(page, separator, printScale, stories, token, out _) ||
        printScale == 1d && page.ColumnFrames.Count <= 1 &&
        HasCompleteFittingEndnoteBlocks(page, placed, separator, printScale, stories, token, frameBottom);

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
        double printScale, IReadOnlyList<DocxRelatedStoryLayout> stories, CancellationToken token, double frameBottom)
    {
        token.ThrowIfCancellationRequested();
        frameBottom = Math.Max(page.MarginBottom, frameBottom);
        if (stories.Count == 0 || !double.IsFinite(frameBottom) ||
            separator?.FloatingDrawings.Count > 0 ||
            placed.Any(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)) return false;
        double available = ResolveEndnoteStartTop(page, placed, printScale) - frameBottom;
        double height = separator?.ContentHeight ?? 0d;
        if (!double.IsFinite(height) || height < 0d || !double.IsFinite(available)) return false;
        foreach (var story in stories)
        {
            token.ThrowIfCancellationRequested();
            if (!double.IsFinite(story.ContentHeight) || story.ContentHeight <= 0d || story.FloatingDrawings.Count != 0) return false;
            // Keep the conservative legacy charges for admission. Placement's
            // separately qualified paragraph boundaries may remove them later.
            height += story.ContentHeight + FootnoteSeparatorGapPoints;
            if (!double.IsFinite(height) || height > available) return false;
        }
        foreach (var story in stories)
            if (!HasCompleteMeasuredNoteBlocks(story, token)) return false;
        return true;
    }

    private static bool CanUseUnscaledEndnoteParagraphBoundaries(DocxDocument document,
        DocxLayoutPage page, DocxRelatedStoryLayout? separator, double printScale, bool hasCompleteStaticFrame) =>
        printScale == 1d && page.ColumnFrames.Count <= 1 &&
        DocxPlainBodyFrame.HasKnownBodyFrame(document) && hasCompleteStaticFrame &&
        HasEmptyNoteSeparator(separator);

    // The first complete source block owns the document-end start boundary.
    // Later notes may need continuation pages without adding the separator's
    // latent paragraph spacing back above this already measured block. Whole-
    // group fitting still owns counter admission, footer flow and overlaps.
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
                        if (row.Revisions.Count != 0 || !DocxMathParagraphPolicy.IsSupportedSimpleRow(row)) return false;
                        foreach (var cell in row.Cells)
                        {
                            token.ThrowIfCancellationRequested();
                            if (cell.Revisions.Count != 0 || !DocxMathParagraphPolicy.IsSupportedCellProperties(cell) ||
                                !Elements(DocxTableCellContent.GetBodyElements(cell))) return false;
                        }
                    }
                }
                else return false;
            }
            return true;
        }
        return Elements(story.Story.BodyElements);
    }

}
