namespace Lokad.OoxPdf.Docx;

internal sealed record DocxPreparedMathParagraph(DocxEquation Equation, DocxMathBox Box, double BaselineInset, double LineHeight);

internal static class DocxMathParagraphPolicy
{
    public static IEnumerable<DocxParagraph> EnumeratePreparationCandidates(DocxDocument document, CancellationToken token)
    {
        foreach (DocxBodyElement element in document.BodyElements)
        {
            token.ThrowIfCancellationRequested();
            if (element is DocxParagraphElement paragraph) { yield return paragraph.Paragraph; }
            else if (element is DocxTableElement table)
            {
                foreach (DocxTableRow row in table.Table.Rows)
                {
                    token.ThrowIfCancellationRequested();
                    if (!row.HasSupportedHeight || row.Revisions.Count != 0)
                        continue;
                    foreach (DocxTableCell cell in row.Cells)
                    {
                        token.ThrowIfCancellationRequested();
                        if (cell.HasVerticalMerge || cell.FitText || cell.Revisions.Count != 0 ||
                            cell.TextDirectionValue is not (null or "lrTb") ||
                            cell.BodyElements.Any(child => child is not DocxParagraphElement)) { continue; }
                        foreach (DocxParagraph source in cell.Paragraphs)
                        {
                            token.ThrowIfCancellationRequested();
                            if (GetStandaloneRow(source) is not null || GetInlineRow(source) is not null) { yield return source; }
                        }
                    }
                }
            }
        }
    }

    public static IEnumerable<DocxParagraph> EnumerateStaticPreparationCandidates(DocxDocument document, CancellationToken token)
    {
        var seen = new HashSet<DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (DocxEquationPartSource part in document.EquationParts)
        {
            token.ThrowIfCancellationRequested();
            if (part.PartName == document.MainPartName) { continue; }
            foreach (DocxParagraphElement element in part.BodyElements.OfType<DocxParagraphElement>())
            {
                token.ThrowIfCancellationRequested();
                if (GetStandaloneRow(element.Paragraph) is not null && seen.Add(element.Paragraph))
                    yield return element.Paragraph;
            }
        }
    }

    // Office promotes a sole visible equation to centered display layout, even
    // when its source is inline. Preserve the source record and select that
    // style only for measurement; mixed prose retains its inline ownership.
    public static DocxEquation? GetStandaloneRow(DocxParagraph paragraph)
    {
        if (!HasSupportedOwner(paragraph) ||
            paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text)) ||
            !TryGetDisplayLineSpacingFactor(paragraph, out _))
        { return null; }
        DocxEquation equation = paragraph.Equations[0];
        return equation.AlignmentValue is null or "center" or "centerGroup" or "left" or "right" ? equation : null;
    }

    public static bool TryGetDisplayLineSpacingFactor(DocxParagraph paragraph, out double factor)
    {
        factor = 1d;
        if (paragraph.Spacing.LineRuleValue is not (null or "auto") || paragraph.LineSpacingPoints is not null)
            return false;
        if (paragraph.Spacing.LineValue is null) { return true; }
        if (!int.TryParse(paragraph.Spacing.LineValue, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out int line) || line <= 0) { return false; }
        factor = paragraph.LineSpacingFactor;
        // Compressed automatic slots change Office's first math baseline and
        // need a separate placement contract. Expanded slots retain the box.
        return double.IsFinite(factor) && factor >= 1d;
    }

    public static DocxEquation? GetInlineRow(DocxParagraph paragraph) =>
        HasSupportedOwner(paragraph) && !paragraph.Equations[0].Display &&
        paragraph.Spacing.LineRuleValue is null or "auto" &&
        paragraph.Runs.All(run => !run.Text.Contains('\t') && !run.Text.Contains('\n')) &&
        paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text))
            ? paragraph.Equations[0] : null;

    private static bool HasSupportedOwner(DocxParagraph paragraph) =>
        paragraph.Equations.Count == 1 && paragraph.Equations[0].Revision is null && paragraph.Revisions.Count == 0 &&
        paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0 &&
        paragraph.InlineReferences.Count == 0 && paragraph.CommentRanges.Count == 0 &&
        paragraph.FieldReferences.Count == 0 && paragraph.Hyperlinks.Count == 0 && paragraph.BookmarkAnchors.Count == 0 &&
        paragraph.RevisionRanges.Count == 0 && paragraph.ListLabel is null;
}
