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
                            if (GetPreparationEquations(source).Count != 0) { yield return source; }
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
                if (GetPreparationEquations(element.Paragraph).Count != 0 && seen.Add(element.Paragraph))
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
            HasVisibleProse(paragraph) ||
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
        HasVisibleProse(paragraph)
            ? paragraph.Equations[0] : null;

    private static bool HasVisibleProse(DocxParagraph paragraph) => paragraph.InlineContent is { } items
        ? DocxInlineContent.HasVisibleText(items) : paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text));

    // Multiple adjacent formulas form one display group when no other visible
    // inline object owns the row. Mixed rows retain their inline measurements.
    public static bool IsMathOnlyDisplayGroup(DocxParagraph paragraph) =>
        HasSupportedSourceOwner(paragraph) && paragraph.Equations.Count > 1 &&
        paragraph.Images.Count == 0 && paragraph.Hyperlinks.Count == 0 &&
        paragraph.InlineContent is { } items &&
        items.All(item => item is DocxInlineEquation ||
            item is DocxInlineText text && text.Run.Text.AsSpan(text.Start, text.Length).IsWhiteSpace()) &&
        paragraph.Equations.All(equation => !equation.Display) &&
        TryGetDisplayLineSpacingFactor(paragraph, out _);

    public static bool NeedsOrderedInlineLayout(DocxParagraph paragraph) =>
        paragraph.Equations.Count > 1 || paragraph.Images.Count != 0 || paragraph.Hyperlinks.Count != 0;

    public static IReadOnlyList<DocxEquation> GetPreparationEquations(DocxParagraph paragraph)
    {
        if (GetStandaloneRow(paragraph) is { } standalone) return [standalone];
        if (GetInlineRow(paragraph) is { } inline) return [inline];
        return NeedsOrderedInlineLayout(paragraph) && HasSupportedSourceOwner(paragraph) && paragraph.InlineContent is not null &&
            paragraph.Equations.All(equation => !equation.Display) &&
            paragraph.Alignment is DocxTextAlignment.Left or DocxTextAlignment.Center or DocxTextAlignment.Right &&
            paragraph.Spacing.LineRuleValue is null or "auto" && paragraph.LineSpacingPoints is null
            ? paragraph.Equations : [];
    }

    private static bool HasSupportedOwner(DocxParagraph paragraph) =>
        HasSupportedSourceOwner(paragraph) && paragraph.Equations.Count == 1 &&
        paragraph.Images.Count == 0 && paragraph.Hyperlinks.Count == 0;

    private static bool HasSupportedSourceOwner(DocxParagraph paragraph) =>
        paragraph.Equations.Count != 0 && paragraph.Equations.All(equation => equation.Revision is null) && paragraph.Revisions.Count == 0 &&
        paragraph.InlineTextBoxes.Count == 0 &&
        paragraph.InlineReferences.Count == 0 && paragraph.CommentRanges.Count == 0 &&
        // Bookmarks are source metadata with no measured or painted extent.
        // Their presence cannot change ownership of an otherwise supported box.
        paragraph.FieldReferences.Count == 0 &&
        paragraph.RevisionRanges.Count == 0 && paragraph.ListLabel is null;
}
