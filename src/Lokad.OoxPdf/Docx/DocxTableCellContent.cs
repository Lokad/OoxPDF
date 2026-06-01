namespace Lokad.OoxPdf.Docx;

internal static class DocxTableCellContent
{
    private const double PlainTextDefaultFontSize = 11d;

    public static IReadOnlyList<DocxParagraph> GetParagraphs(DocxTableCell cell)
    {
        return cell.Paragraphs.Count == 0 && cell.Text.Length != 0
            ? [CreatePlainTextParagraph(cell.Text)]
            : cell.Paragraphs;
    }

    private static DocxParagraph CreatePlainTextParagraph(string text)
    {
        return new DocxParagraph(
            [new DocxTextRun(text, PlainTextDefaultFontSize, null, false, false, false, null, null)],
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
