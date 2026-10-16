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
                foreach (DocxParagraph source in EnumerateSupportedTableParagraphs(table.Table, token))
                    if (GetPreparationEquations(source, UsesOnlyPageNumberCounters(source)).Count != 0) { yield return source; }
            }
        }
    }

    // The same simple cell contract applies in body and page-owned static stories.
    // Nested/rich cells retain their existing whole-owner fallback.
    private static IEnumerable<DocxParagraph> EnumerateSupportedTableParagraphs(DocxTable table, CancellationToken token)
    {
        foreach (DocxTableRow row in table.Rows)
        {
            token.ThrowIfCancellationRequested();
            if (!IsSupportedSimpleRow(row)) { continue; }
            foreach (DocxTableCell cell in row.Cells)
            {
                token.ThrowIfCancellationRequested();
                if (!IsSupportedSimpleCell(cell)) { continue; }
                foreach (DocxParagraph paragraph in cell.Paragraphs)
                { token.ThrowIfCancellationRequested(); yield return paragraph; }
            }
        }
    }

    public static bool IsSupportedSimpleRow(DocxTableRow row) => row.HasSupportedHeight && row.Revisions.Count == 0;

    public static bool IsSupportedSimpleCell(DocxTableCell cell) =>
        !cell.HasVerticalMerge && !cell.FitText && cell.Revisions.Count == 0 &&
        cell.TextDirectionValue is null or "lrTb" &&
        cell.BodyElements.All(child => child is DocxParagraphElement);

    public static IEnumerable<(DocxParagraph Paragraph, bool HasPageInputs)> EnumerateStaticPreparationCandidates(DocxDocument document, CancellationToken token)
    {
        var pageOwners = new HashSet<DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var paragraph in DocxBlockTraversal.EnumerateReferencedStaticStoryParagraphs(
                document.HeaderBodyElementsByType, document.HeaderParagraphsByType, document.HeaderParagraphs)
            .Concat(DocxBlockTraversal.EnumerateReferencedStaticStoryParagraphs(
                document.FooterBodyElementsByType, document.FooterParagraphsByType, document.FooterParagraphs))
            .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.PageSettings))
            .Concat(document.BodyElements.OfType<DocxSectionBreakElement>().SelectMany(section =>
                DocxBlockTraversal.EnumerateStaticStoryParagraphs(section.PageSettings))))
        {
            token.ThrowIfCancellationRequested();
            pageOwners.Add(paragraph);
        }
        var seen = new HashSet<DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (DocxEquationPartSource part in document.EquationParts)
        {
            token.ThrowIfCancellationRequested();
            if (part.PartName == document.MainPartName) { continue; }
            foreach (DocxBodyElement element in part.BodyElements)
            {
                token.ThrowIfCancellationRequested();
                IEnumerable<DocxParagraph> paragraphs = element switch
                {
                    DocxParagraphElement direct => [direct.Paragraph],
                    DocxTableElement table => EnumerateSupportedTableParagraphs(table.Table, token).Where(pageOwners.Contains),
                    _ => []
                };
                foreach (DocxParagraph paragraph in paragraphs)
                {
                    token.ThrowIfCancellationRequested();
                    bool hasPageInputs = pageOwners.Contains(paragraph);
                    if (GetPreparationEquations(paragraph, hasPageInputs).Count != 0 && seen.Add(paragraph))
                        yield return (paragraph, hasPageInputs);
                }
            }
        }
    }

    // Office promotes a sole visible equation to centered display layout, even
    // when its source is inline. Preserve the source record and select that
    // style only for measurement; mixed prose retains its inline ownership.
    public static DocxEquation? GetStandaloneRow(DocxParagraph paragraph)
    {
        if (!HasSupportedOwner(paragraph) || paragraph.ListLabel is not null || paragraph.Hyperlinks.Count != 0 ||
            paragraph.CommentRanges.Count != 0 ||
            HasVisibleProse(paragraph) || (paragraph.InlineContent is { } items &&
            items.Any(item => item is DocxInlineTab tab && !tab.Run.EffectiveProperties.Hidden)) ||
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

    // Qualified complete inline boxes share ordered placement; sole display
    // rows and unsupported alignment retain their existing container paths.
    public static bool NeedsOrderedInlineLayout(DocxParagraph paragraph,
        IReadOnlyDictionary<DocxEquation, DocxPreparedMathParagraph>? prepared = null) =>
        paragraph.Equations.Count > 1 || paragraph.Images.Count != 0 ||
        paragraph.Equations.Count != 0 && paragraph.CommentRanges.Count != 0 ||
        paragraph.Equations.Count != 0 && paragraph.InlineReferences.Any(reference => DocxInlineContent.TryGetAutomaticNoteRun(reference, out _)) ||
        paragraph.Equations.Count != 0 && paragraph.Runs.Any(run => run.FieldKind is { } kind && kind.IsDynamic()) ||
        paragraph.Equations.Count == 1 && (GetInlineRow(paragraph) is null && GetStandaloneRow(paragraph) is null ||
            GetInlineRow(paragraph) is { } inline && prepared is not null &&
            prepared.TryGetValue(inline, out var math) && (math.Box.InlineBreakAfterGlyphs.Count != 0 ||
                math.Box.AllowsInlineRelocation && HasSupportedOrderedAlignment(paragraph)));

    private static bool HasSupportedOrderedAlignment(DocxParagraph paragraph) =>
        paragraph.Alignment is DocxTextAlignment.Left or DocxTextAlignment.Center or DocxTextAlignment.Right;

    // Preparation needs the dependency kind, not a guessed counter value.
    // Body/cell placement supplies PAGE; total counts remain unavailable there.
    public static bool UsesOnlyPageNumberCounters(DocxParagraph paragraph) =>
        paragraph.Runs.Any(run => run.FieldKind == DocxFieldKind.Page) &&
        paragraph.Runs.All(run => run.FieldKind != DocxFieldKind.NumPages) &&
        paragraph.FieldReferences.All(field => field.Kind != DocxFieldKind.NumPages);

    public static IReadOnlyList<DocxEquation> GetPreparationEquations(DocxParagraph paragraph, bool hasPageInputs = false)
    {
        if (GetStandaloneRow(paragraph) is { } standalone) return [standalone];
        if (GetInlineRow(paragraph) is { } inline) return [inline];
        return NeedsOrderedInlineLayout(paragraph) && HasSupportedSourceOwner(paragraph, hasPageInputs) && paragraph.InlineContent is not null &&
            paragraph.Equations.All(equation => !equation.Display) &&
            HasSupportedOrderedAlignment(paragraph) &&
            paragraph.Spacing.LineRuleValue is null or "auto" && paragraph.LineSpacingPoints is null
            ? paragraph.Equations : [];
    }

    private static bool HasSupportedOwner(DocxParagraph paragraph) =>
        HasSupportedSourceOwner(paragraph) && paragraph.Equations.Count == 1 &&
        paragraph.Images.Count == 0;

    private static bool HasSupportedSourceOwner(DocxParagraph paragraph, bool hasPageInputs = false) =>
        paragraph.Equations.Count != 0 && paragraph.Equations.All(equation => equation.Revision is null) && paragraph.Revisions.Count == 0 &&
        paragraph.InlineTextBoxes.Count == 0 &&
        DocxInlineContent.HasSupportedReferences(paragraph.InlineReferences, paragraph.Runs) &&
        // Bookmarks are source metadata with no measured or painted extent.
        // Their presence cannot change ownership of an otherwise supported box.
        (hasPageInputs ? DocxInlineContent.HasSourceFieldResults(paragraph.FieldReferences, paragraph.Runs)
            : DocxInlineContent.HasMaterializedFieldResults(paragraph.FieldReferences, paragraph.Runs)) &&
        paragraph.RevisionRanges.Count == 0;
}
