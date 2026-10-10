namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static IReadOnlySet<DocxParagraph> CreateSingleLineNoteSpacingParagraphs(
        DocxDocument document, DocxLayout layout, DocxMarkupContext context, CancellationToken cancellationToken)
    {
        var candidates = new HashSet<DocxParagraph>(ReferenceEqualityComparer.Instance);
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
        foreach (DocxLayoutPage page in layout.Pages)
        foreach (DocxTextLineLayout line in EnumerateBodyTextLines(page))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (line.SourceParagraph is not { } paragraph) continue;
            counts[paragraph] = counts.GetValueOrDefault(paragraph) + 1;
            if ((line.LineHeightSource is DocxLineHeightSource.BodySingleLineAuto or DocxLineHeightSource.ExactLineSpacing) &&
                line.BodyLineBoxHeightPoints is > 0d && line.BodyLineBoxBaselineInsetPoints is not null &&
                line.ParagraphAfterSpacing is > 0d && paragraph.SpacingBeforePoints == 0d &&
                paragraph.Spacing.ContextualSpacing != true && paragraph.Spacing.AfterAutoSpacingValue is null &&
                paragraph.Spacing.AfterLinesValue is null && paragraph.Images.Count == 0 &&
                paragraph.InlineTextBoxes.Count == 0 && paragraph.FieldReferences.Count == 0 &&
                paragraph.CommentRanges.Count == 0 && paragraph.Revisions.Count == 0 && paragraph.RevisionRanges.Count == 0 &&
                paragraph.ListLabel is null && paragraph.Hyperlinks.Count == 0 &&
                paragraph.Runs.All(r => !r.Text.Contains('\n') && !r.Text.Contains('\r') && !r.Text.Contains('\t')))
            {
                candidates.Add(paragraph);
            }
        }
        candidates.RemoveWhere(p => counts[p] != 1);
        return candidates;
    }
}
