namespace Lokad.OoxPdf.Docx;

internal sealed record DocxPreparedMathParagraph(DocxEquation Equation, DocxMathBox Box, double BaselineInset, double LineHeight)
{
    public bool UsesDisplayStyle { get; init; }
}

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
        if (!HasSupportedOwner(paragraph) || paragraph.ListLabel is not null || paragraph.Hyperlinks.Count != 0 ||
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
        HasSupportedOwner(paragraph) && paragraph.ListLabel is null && !paragraph.Equations[0].Display &&
        paragraph.Spacing.LineRuleValue is null or "auto" &&
        paragraph.Runs.All(run => !run.Text.Contains('\t') && !run.Text.Contains('\n')) &&
        HasVisibleProse(paragraph)
            ? paragraph.Equations[0] : null;

    private static bool HasVisibleProse(DocxParagraph paragraph) => paragraph.InlineContent is { } items
        ? DocxInlineContent.HasVisibleText(items) : paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text));

    // Authored breaks delimit source rows. Only rows without another visible
    // inline owner promote their formulas; soft wrapping never changes style.
    public static IReadOnlySet<DocxEquation> GetDisplaySourceEquations(DocxParagraph paragraph,
        CancellationToken token)
    {
        var display = new HashSet<DocxEquation>(ReferenceEqualityComparer.Instance);
        if (!HasSupportedSourceOwner(paragraph) || paragraph.ListLabel is not null || paragraph.Hyperlinks.Count != 0 ||
            paragraph.InlineContent is not { } items || !TryGetDisplayLineSpacingFactor(paragraph, out _))
            return display;
        var row = new List<DocxEquation>();
        bool hasOtherContent = false;
        foreach (var item in items)
        {
            token.ThrowIfCancellationRequested();
            switch (item)
            {
                case DocxInlineText text when text.Run.EffectiveProperties.Hidden ||
                    text.Run.Text.AsSpan(text.Start, text.Length).IsWhiteSpace():
                case DocxInlineTab tab when tab.Run.EffectiveProperties.Hidden:
                case DocxInlineBreak lineBreak when lineBreak.Run.EffectiveProperties.Hidden:
                    break;
                case DocxInlineEquation math:
                    row.Add(math.Equation);
                    break;
                case DocxInlineBreak:
                    CompleteRow();
                    break;
                default:
                    hasOtherContent = true;
                    break;
            }
        }
        CompleteRow();
        return display;

        void CompleteRow()
        {
            if (!hasOtherContent)
                foreach (var equation in row)
                {
                    token.ThrowIfCancellationRequested();
                    display.Add(equation);
                }
            row.Clear();
            hasOtherContent = false;
        }
    }

    // Retain qualified single-prose breaking until mathematical operator breaks
    // migrate. An unchanged annotation cannot choose a different layout engine.
    public static bool NeedsOrderedInlineLayout(DocxParagraph paragraph) =>
        paragraph.Equations.Count > 1 || paragraph.Images.Count != 0 ||
        paragraph.Equations.Count == 1 && GetInlineRow(paragraph) is null && GetStandaloneRow(paragraph) is null;

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
        paragraph.Images.Count == 0;

    private static bool HasSupportedSourceOwner(DocxParagraph paragraph) =>
        paragraph.Equations.Count != 0 && paragraph.Equations.All(equation => equation.Revision is null) && paragraph.Revisions.Count == 0 &&
        paragraph.InlineTextBoxes.Count == 0 &&
        paragraph.InlineReferences.Count == 0 && paragraph.CommentRanges.Count == 0 &&
        // Bookmarks are source metadata with no measured or painted extent.
        // Their presence cannot change ownership of an otherwise supported box.
        paragraph.FieldReferences.Count == 0 &&
        paragraph.RevisionRanges.Count == 0;
}
