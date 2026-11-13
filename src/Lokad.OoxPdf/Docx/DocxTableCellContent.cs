namespace Lokad.OoxPdf.Docx;

internal static class DocxTableCellContent
{
    // An empty terminal paragraph after a nested table is a cell-end source
    // owner, not another occupied line. Ordinary characters, references and
    // painted/revised content retain their paragraph frame. Bookmarks remain
    // attached to the source and can use its nonpainting layout carrier.
    public static bool IsTerminalNestedTableMark(IReadOnlyList<DocxBodyElement> elements,
        int index, CancellationToken token)
    {
        if (index != elements.Count - 1 || index < 1 || elements[index - 1] is not DocxTableElement ||
            elements[index] is not DocxParagraphElement { Paragraph: var p } ||
            p.ListLabel is not null || p.Images.Count != 0 || p.Equations.Count != 0 ||
            p.InlineTextBoxes.Count != 0 || p.InlineReferences.Count != 0 || p.FieldReferences.Count != 0 ||
            p.Hyperlinks.Count != 0 || p.CommentRanges.Count != 0 || p.Shading is not null ||
            p.Revisions.Count != 0 || DocxInlineContent.HasContentRevisionRanges(p) || p.HasDeletedParagraphMark)
            return false;
        foreach (var run in p.Runs)
        {
            token.ThrowIfCancellationRequested();
            if (run.Text.Length != 0 || run.FieldKind is not null || run.Revision is not null || run.Revisions.Count != 0)
                return false;
        }
        foreach (var owner in p.NonRenderingRunOwners)
        {
            token.ThrowIfCancellationRequested();
            if (owner.Revisions.Count != 0) return false;
        }
        if (p.InlineContent is { } items)
            foreach (var item in items)
            {
                token.ThrowIfCancellationRequested();
                if (item is not DocxInlineText { Length: 0 }) return false;
            }
        return true;
    }

    public static IReadOnlyList<DocxBodyElement> GetBodyElements(DocxTableCell cell)
    {
        return cell.BodyElements.Count != 0
            ? cell.BodyElements
            : GetParagraphs(cell)
                .Select(DocxBodyElementFactory.CreateParagraph)
                .Cast<DocxBodyElement>()
                .ToArray();
    }

    public static IReadOnlyList<DocxParagraph> GetParagraphs(DocxTableCell cell)
    {
        if (cell.BodyElements.Count != 0)
        {
            return DocxBlockTraversal.EnumerateDirectParagraphs(cell.BodyElements).ToArray();
        }

        return cell.Paragraphs.Count == 0 && cell.Text.Length != 0
            ? [CreatePlainTextParagraph(cell.Text)]
            : cell.Paragraphs;

        DocxParagraph CreatePlainTextParagraph(string text)
        {
            return new DocxParagraph(
                [new DocxTextRun(text, DocxDefaults.FontSizePoints, null, false, false, false, null, null)],
                [],
                null,
                DocxTextAlignment.Left,
                null,
                0d,
                0d,
                1d,
                null,
                DocxParagraphSpacing.Empty,
                DocxParagraphKeepRules.Empty,
                null);
        }
    }

}
