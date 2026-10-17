namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // A complete measured block owns paragraph boundaries even when its middle
    // contains supported inline objects or complete nested tables. This admission does
    // not change typography or the separately qualified continuation policies.
    private static bool CanUseFittingEndnoteBoundaries(DocxDocument document, DocxLayoutPage page,
        IReadOnlyList<DocxPlacedRelatedStoryLayout> placed, DocxRelatedStoryLayout? separator,
        double printScale, IReadOnlyList<DocxRelatedStoryLayout> stories, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (printScale != 1d || page.ColumnFrames.Count > 1 || stories.Count == 0 ||
            !DocxPlainBodyFrame.HasKnownBodyFrame(document) || !DocxPlainBodyFrame.HasNoStaticStories(document) ||
            !HasEmptyNoteSeparator(separator) || !HasNumericTrailingBodyAfterSpacing(page) ||
            placed.Any(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote)) return false;
        double available = ResolveEndnoteStartTop(page, placed, printScale) - page.MarginBottom;
        double height = separator!.ContentHeight;
        if (!double.IsFinite(height) || height < 0d || !double.IsFinite(available)) return false;
        foreach (var story in stories)
        {
            token.ThrowIfCancellationRequested();
            if (!double.IsFinite(story.ContentHeight) || story.ContentHeight <= 0d || story.FloatingDrawings.Count != 0) return false;
            // Retain the legacy charges in this conservative fit proof. Removing
            // them for placement cannot turn a complete fitting block into a slice.
            height += story.ContentHeight + FootnoteSeparatorGapPoints;
            if (!double.IsFinite(height) || height > available) return false;
        }
        foreach (var story in stories)
        {
            token.ThrowIfCancellationRequested();
            if (story.Story.BodyElements.Count == 0 || story.TextLines.Count == 0 ||
                story.Story.BodyElements[0] is not DocxParagraphElement first ||
                story.Story.BodyElements[^1] is not DocxParagraphElement last ||
                !ReferenceEquals(story.TextLines[0].SourceParagraph, first.Paragraph) ||
                !ReferenceEquals(story.TextLines[^1].SourceParagraph, last.Paragraph)) return false;
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
            long expectedCellLines = 0;
            bool CompleteRows(IReadOnlyList<DocxTableRowLayout> rows)
            {
                foreach (var row in rows)
                {
                    token.ThrowIfCancellationRequested();
                    if (row.FragmentCount != 1) return false;
                    foreach (var cell in row.Cells)
                    {
                        token.ThrowIfCancellationRequested();
                        expectedCellLines += cell.TextLines.Count;
                        if (!CompleteRows(cell.NestedRows)) return false;
                    }
                }
                return true;
            }
            if (!CompleteRows(story.TableRows)) return false;
            long visibleCellLines = 0;
            foreach (var row in story.TableRows)
            foreach (var owner in DocxTableCellClipGeometry.EnumerateCellLines(row, null, token))
            {
                token.ThrowIfCancellationRequested();
                visibleCellLines++;
                if (!Index(owner.Line)) return false;
            }
            // A surviving first line cannot stand in for a partly clipped
            // paragraph. Every measured line must survive all ancestor clips.
            if (visibleCellLines != expectedCellLines) return false;
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
            if (!Elements(story.Story.BodyElements)) return false;
        }
        return true;
    }
}
