namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // A complete measured block owns paragraph boundaries even when its middle
    // contains supported inline objects or a simple table. This admission does
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
            foreach (var row in story.TableRows)
            {
                token.ThrowIfCancellationRequested();
                if (row.FragmentCount != 1) return false;
                foreach (var cell in row.Cells)
                {
                    token.ThrowIfCancellationRequested();
                    if (cell.NestedRows.Count != 0) return false;
                    var clip = DocxTableCellClipGeometry.Resolve(row, cell);
                    foreach (var line in cell.TextLines)
                    {
                        token.ThrowIfCancellationRequested();
                        if (!DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double bottom, out double top) ||
                            bottom < clip.Y - .000001d || top > clip.Y + clip.Height + .000001d || !Index(line)) return false;
                    }
                }
            }
            foreach (var element in story.Story.BodyElements)
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
                            if (cell.Revisions.Count != 0 || !DocxMathParagraphPolicy.IsSupportedSimpleCell(cell)) return false;
                            foreach (var source in cell.Paragraphs) if (!Paragraph(source)) return false;
                        }
                    }
                }
                else return false;
            }
        }
        return true;
    }
}
