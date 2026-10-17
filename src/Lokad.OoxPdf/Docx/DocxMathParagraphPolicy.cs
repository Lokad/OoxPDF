namespace Lokad.OoxPdf.Docx;

internal sealed record DocxPreparedMathParagraph(DocxEquation Equation, DocxMathBox Box, double BaselineInset, double LineHeight)
{
    public bool UsesDisplayStyle { get; init; } = Equation.Display;
}

internal static class DocxMathParagraphPolicy
{
    public static IEnumerable<DocxParagraph> EnumeratePreparationCandidates(DocxDocument document, CancellationToken token)
        => EnumeratePreparationCandidates(document, token, includePageCountDependencies: false);

    public static IEnumerable<DocxParagraph> EnumeratePreparationCandidates(DocxDocument document, CancellationToken token,
        bool includePageCountDependencies)
    {
        foreach (DocxBodyElement element in document.BodyElements)
        {
            token.ThrowIfCancellationRequested();
            if (element is DocxParagraphElement paragraph) { yield return paragraph.Paragraph; }
            else if (element is DocxTableElement table)
            {
                foreach (DocxParagraph source in EnumerateSupportedTableParagraphs(table.Table, token))
                    if (GetPreparationEquations(source, includePageCountDependencies || UsesOnlyPageNumberCounters(source), token).Count != 0) { yield return source; }
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

    public static bool IsSupportedSimpleRow(DocxTableRow row) => row.HasSupportedHeight &&
        HasSupportedTableContentRevisions(row.Revisions, DocxRevisionKind.TableRowPropertiesChange);

    public static bool IsSupportedSimpleCell(DocxTableCell cell) =>
        !cell.HasVerticalMerge && !cell.FitText &&
        HasSupportedTableContentRevisions(cell.Revisions, DocxRevisionKind.TableCellPropertiesChange) &&
        cell.TextDirectionValue is null or "lrTb" &&
        cell.BodyElements.All(child => child is DocxParagraphElement);

    // An insertion or a change to the resolved container properties does not
    // replace the child paragraph's source owner. Paragraph eligibility still
    // checks its actual revisions; deleting/moving table owners stay staged.
    private static bool HasSupportedTableContentRevisions(IReadOnlyList<DocxRevisionInfo> revisions,
        DocxRevisionKind propertiesKind) => revisions.All(revision =>
            revision.Kind == DocxRevisionKind.Insertion || revision.Kind == propertiesKind);

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
                    if (GetPreparationEquations(paragraph, hasPageInputs, token).Count != 0 && seen.Add(paragraph))
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
        CancellationToken token) => paragraph.Hyperlinks.Count == 0
            ? GetMathOnlySourceEquations(paragraph, token)
            : new HashSet<DocxEquation>(ReferenceEqualityComparer.Instance);

    private static IReadOnlySet<DocxEquation> GetMathOnlySourceEquations(DocxParagraph paragraph,
        CancellationToken token) => GetSourceRowEquations(paragraph, token, withInlineOwner: false);

    public static IReadOnlySet<DocxEquation> GetMixedSourceEquations(DocxParagraph paragraph,
        CancellationToken token) => GetSourceRowEquations(paragraph, token, withInlineOwner: true);

    private static IReadOnlySet<DocxEquation> GetSourceRowEquations(DocxParagraph paragraph,
        CancellationToken token, bool withInlineOwner)
    {
        var display = new HashSet<DocxEquation>(ReferenceEqualityComparer.Instance);
        if (!HasSupportedSourceOwner(paragraph, hasPageInputs: true, token) ||
            paragraph.InlineContent is not { } items || !TryGetDisplayLineSpacingFactor(paragraph, out _))
            return display;
        var row = new List<DocxEquation>();
        // A list label is visible content on the first authored row. A hard
        // break ends that ownership just like prose; later math-only rows may
        // own their display frame independently.
        bool hasInlineOwner = paragraph.ListLabel is not null;
        bool hasOtherContent = hasInlineOwner;
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
                case DocxInlineText:
                case DocxInlinePicture:
                    hasInlineOwner = true;
                    hasOtherContent = true;
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
            if (withInlineOwner ? hasInlineOwner : !hasOtherContent)
                foreach (var equation in row)
                {
                    token.ThrowIfCancellationRequested();
                    display.Add(equation);
                }
            row.Clear();
            hasOtherContent = false;
            hasInlineOwner = false;
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

    public static DocxTextAlignment? ResolveDisplayAlignment(IEnumerable<DocxEquation> equations, CancellationToken token)
    {
        DocxTextAlignment? alignment = null;
        foreach (var equation in equations)
        {
            token.ThrowIfCancellationRequested();
            DocxTextAlignment? current = equation.AlignmentValue switch
            {
                null or "center" or "centerGroup" => DocxTextAlignment.Center,
                "left" => DocxTextAlignment.Left,
                "right" => DocxTextAlignment.Right,
                _ => null
            };
            if (current is null || alignment is not null && alignment != current) return null;
            alignment = current;
        }
        return alignment;
    }

    private static bool DisplayRowsOwnParagraphAlignment(DocxParagraph paragraph,
        IReadOnlySet<DocxEquation> displayRows, CancellationToken token)
    {
        if (paragraph.Hyperlinks.Count != 0 || paragraph.InlineContent is not { } items) return false;
        foreach (var item in items)
        {
            token.ThrowIfCancellationRequested();
            if (item is DocxInlineEquation math && displayRows.Contains(math.Equation) || item is DocxInlineBreak ||
                item is DocxInlineText text && (text.Run.EffectiveProperties.Hidden || text.Run.Text.AsSpan(text.Start, text.Length).IsWhiteSpace())) continue;
            return false;
        }
        return true;
    }

    public static IReadOnlyList<DocxEquation> GetPreparationEquations(DocxParagraph paragraph, bool hasPageInputs = false) =>
        GetPreparationEquations(paragraph, hasPageInputs, default);

    public static IReadOnlyList<DocxEquation> GetPreparationEquations(DocxParagraph paragraph, bool hasPageInputs, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (GetStandaloneRow(paragraph) is { } standalone) return [standalone];
        if (GetInlineRow(paragraph) is { } inline) return [inline];
        if (!NeedsOrderedInlineLayout(paragraph) || !HasSupportedSourceOwner(paragraph, hasPageInputs, token) ||
            paragraph.InlineContent is null ||
            paragraph.Spacing.LineRuleValue is not (null or "auto") || paragraph.LineSpacingPoints is not null) return [];
        // Source rows own resolved style: Word measures explicit formulas inline
        // beside visible text/pictures; math-only rows retain display measurement/alignment.
        var displayRows = paragraph.Equations.Any(equation => equation.Display)
            ? GetMathOnlySourceEquations(paragraph, token) : null;
        var mixedRows = displayRows is not null ? GetMixedSourceEquations(paragraph, token) : null;
        if (displayRows is not null && ResolveDisplayAlignment(paragraph.Equations, token) is null) return [];
        if (!HasSupportedOrderedAlignment(paragraph) &&
            (displayRows is null || !DisplayRowsOwnParagraphAlignment(paragraph, displayRows, token))) return [];
        foreach (var equation in paragraph.Equations)
        {
            token.ThrowIfCancellationRequested();
            if (equation.Display && (displayRows is null || !displayRows.Contains(equation)) &&
                (mixedRows is null || !mixedRows.Contains(equation))) return [];
        }
        return paragraph.Equations;
    }

    private static bool HasSupportedOwner(DocxParagraph paragraph) =>
        HasSupportedSourceOwner(paragraph) && paragraph.Equations.Count == 1 &&
        paragraph.Images.Count == 0;

    private static bool HasSupportedSourceOwner(DocxParagraph paragraph, bool hasPageInputs = false,
        CancellationToken token = default) =>
        paragraph.Equations.Count != 0 && DocxInlineContent.HasSupportedSourceOwner(paragraph, hasPageInputs, token);
}
