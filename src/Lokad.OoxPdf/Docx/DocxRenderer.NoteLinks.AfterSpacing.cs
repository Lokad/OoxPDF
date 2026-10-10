namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private readonly record struct NoteReferenceSpacing(double Before, double After);

    private static IReadOnlyDictionary<DocxParagraph, NoteReferenceSpacing> CreateNoteSpacingParagraphs(
        DocxDocument document, DocxLayout layout, DocxMarkupContext context, CancellationToken cancellationToken)
    {
        var candidates = new Dictionary<DocxParagraph, NoteReferenceSpacing>(ReferenceEqualityComparer.Instance);
        if (Math.Abs(ResolveTextEmissionFontScale(context) - 1d) > .000001d ||
            document.FloatingDrawings.Count != 0 || document.HeaderParagraphs.Count != 0 || document.FooterParagraphs.Count != 0 ||
            document.HeaderParagraphsByType.Count != 0 || document.FooterParagraphsByType.Count != 0 ||
            document.HeaderBodyElementsByType.Count != 0 || document.FooterBodyElementsByType.Count != 0 ||
            document.HeaderFloatingDrawingsByType.Count != 0 || document.FooterFloatingDrawingsByType.Count != 0 ||
            document.Settings.MirrorMargins == true ||
            document.BodyElements.Any(e => e.Revisions.Count != 0 ||
                e is not (DocxParagraphElement or DocxSectionBreakElement)) ||
            layout.Pages.Any(p => p.ColumnFrames.Count != 1 || p.StaticTextLines.Count != 0 ||
                p.StaticInlineImages.Count != 0 || p.StaticTableRows.Count != 0 || p.StaticInlineTextBoxes.Count != 0))
        {
            return candidates;
        }

        var counts = new Dictionary<DocxParagraph, int>(ReferenceEqualityComparer.Instance);
        var firstLines = new Dictionary<DocxParagraph, DocxTextLineLayout>(ReferenceEqualityComparer.Instance);
        foreach (DocxLayoutPage page in layout.Pages)
        foreach (DocxTextLineLayout line in EnumerateBodyTextLines(page))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (line.SourceParagraph is not { } paragraph) continue;
            counts[paragraph] = counts.GetValueOrDefault(paragraph) + 1;
            if ((line.LineHeightSource is DocxLineHeightSource.BodySingleLineAuto or DocxLineHeightSource.ExactLineSpacing) &&
                line.BodyLineBoxHeightPoints is > 0d && line.BodyLineBoxBaselineInsetPoints is not null &&
                (line.ParagraphAfterSpacing is > 0d || line.ParagraphBeforeSpacing is > 0d) &&
                paragraph.SpacingBeforePoints >= 0d && paragraph.SpacingAfterPoints >= 0d &&
                paragraph.Spacing.AfterAutoSpacingValue is null &&
                paragraph.Spacing.AfterLinesValue is null && paragraph.Spacing.BeforeAutoSpacingValue is null &&
                paragraph.Spacing.BeforeLinesValue is null && paragraph.Images.Count == 0 &&
                paragraph.InlineTextBoxes.Count == 0 && paragraph.FieldReferences.Count == 0 &&
                paragraph.CommentRanges.Count == 0 && paragraph.Revisions.Count == 0 && paragraph.RevisionRanges.Count == 0 &&
                paragraph.ListLabel is null && paragraph.Hyperlinks.Count == 0 &&
                paragraph.Runs.All(r => !r.Text.Contains('\n') && !r.Text.Contains('\r') && !r.Text.Contains('\t')))
            {
                firstLines[paragraph] = line;
                candidates[paragraph] = new NoteReferenceSpacing(
                    Math.Max(0d, (line.AppliedBeforeSpacing ?? 0d) - (line.PendingAfterSpacing ?? 0d)),
                    line.ParagraphAfterSpacing ?? 0d);
            }
        }
        DocxParagraph? previous = null;
        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxBodyElement element = document.BodyElements[elementIndex];
            if (element is DocxSectionBreakElement) { previous = null; continue; }
            if (element is not DocxParagraphElement current) continue;
            DocxParagraph paragraph = current.Paragraph;
            if (paragraph.SpacingBeforePoints > 0d && previous is not null &&
                (counts.GetValueOrDefault(previous) > 1 ||
                 previous.Spacing.ContextualSpacing == true && counts.GetValueOrDefault(paragraph) != 1 &&
                     (counts.GetValueOrDefault(previous) != 1 ||
                      previous.Runs.All(r => string.IsNullOrWhiteSpace(r.Text))) ||
                 previous.SpacingBeforePoints < 0d || previous.SpacingAfterPoints < 0d ||
                 previous.Spacing.BeforeAutoSpacingValue is not null || previous.Spacing.BeforeLinesValue is not null ||
                 previous.Spacing.AfterAutoSpacingValue is not null || previous.Spacing.AfterLinesValue is not null ||
                 previous.Images.Count != 0 || previous.InlineTextBoxes.Count != 0 || previous.FieldReferences.Count != 0 ||
                 previous.CommentRanges.Count != 0 || previous.Revisions.Count != 0 || previous.RevisionRanges.Count != 0 ||
                 previous.ListLabel is not null || previous.Hyperlinks.Count != 0 ||
                 previous.Runs.Any(r => r.Text.Contains('\n') || r.Text.Contains('\r') || r.Text.Contains('\t'))))
            {
                candidates.Remove(paragraph);
            }
            if (candidates.TryGetValue(paragraph, out NoteReferenceSpacing spacing))
            {
                DocxTextLineLayout line = firstLines[paragraph];
                bool suppressPredecessorAfter = previous?.Spacing.ContextualSpacing == true &&
                    DocxLayoutEngine.HasSameContextualSpacingStyle(previous, paragraph) &&
                    DocxLayoutEngine.IsPlainContextualSpacingParagraph(previous) &&
                    DocxLayoutEngine.IsPlainContextualSpacingParagraph(paragraph);
                DocxParagraph? next = elementIndex + 1 < document.BodyElements.Count &&
                    document.BodyElements[elementIndex + 1] is DocxParagraphElement following ? following.Paragraph : null;
                if (paragraph.Spacing.ContextualSpacing == true &&
                    (!DocxLayoutEngine.IsPlainContextualSpacingParagraph(paragraph) ||
                     next is null || !DocxLayoutEngine.IsPlainContextualSpacingParagraph(next)))
                {
                    candidates.Remove(paragraph);
                    previous = paragraph;
                    continue;
                }
                bool suppressAfter = paragraph.Spacing.ContextualSpacing == true &&
                    next is not null && DocxLayoutEngine.HasSameContextualSpacingStyle(paragraph, next);
                candidates[paragraph] = spacing with
                {
                    Before = Math.Max(0d, (line.AppliedBeforeSpacing ?? 0d) -
                        (suppressPredecessorAfter ? 0d : line.PendingAfterSpacing ?? 0d)),
                    After = suppressAfter ? 0d : spacing.After
                };
            }
            previous = paragraph;
        }
        // Resolve the predecessor's contribution before deciding whether a
        // multiline paragraph owns a positive first-line before-gap.
        foreach ((DocxParagraph paragraph, NoteReferenceSpacing spacing) in candidates.ToArray())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (counts[paragraph] == 1) continue;
            if (spacing.Before <= 0d) candidates.Remove(paragraph);
            else candidates[paragraph] = spacing with { After = 0d };
        }
        return candidates;
    }
}
