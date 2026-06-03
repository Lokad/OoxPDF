namespace Lokad.OoxPdf.Docx;

internal static class DocxTableCellContent
{
    public static IReadOnlyList<DocxBodyElement> GetBodyElements(DocxTableCell cell)
    {
        return cell.BodyElements.Count != 0
            ? cell.BodyElements
            : GetParagraphs(cell)
                .Select(paragraph => new DocxParagraphElement(paragraph))
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
    }

    private static DocxParagraph CreatePlainTextParagraph(string text)
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
