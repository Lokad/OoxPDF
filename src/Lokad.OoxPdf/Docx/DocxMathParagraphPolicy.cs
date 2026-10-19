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

    // Follow source block order through ordinary cells. Every ancestor keeps
    // the same property guards; a nested table never substitutes a paragraph
    // from the cell's derived collection for its actual block owner.
    private static IEnumerable<DocxParagraph> EnumerateSupportedTableParagraphs(DocxTable table, CancellationToken token)
    {
        foreach (DocxTableRow row in table.Rows)
        {
            token.ThrowIfCancellationRequested();
            if (!IsSupportedSimpleRow(row)) { continue; }
            foreach (DocxTableCell cell in row.Cells)
            {
                token.ThrowIfCancellationRequested();
                var children = DocxTableCellContent.GetBodyElements(cell);
                // A restart cell owns visible content across its merge span.
                // Preparation preserves that source; placement still applies
                // the existing merge geometry and continuation ownership.
                if (!HasSupportedCellContentProperties(cell) || cell.HasVerticalMerge && !cell.IsVerticalMergeRestart ||
                    !children.All(child => child is DocxParagraphElement or DocxTableElement)) { continue; }
                foreach (var child in children)
                {
                    token.ThrowIfCancellationRequested();
                    if (child is DocxParagraphElement paragraph) yield return paragraph.Paragraph;
                    else if (child is DocxTableElement nested)
                        foreach (var descendant in EnumerateSupportedTableParagraphs(nested.Table, token)) yield return descendant;
                }
            }
        }
    }

    public static bool IsSupportedSimpleRow(DocxTableRow row) => row.HasSupportedHeight &&
        HasSupportedTableContentRevisions(row.Revisions, DocxRevisionKind.TableRowPropertiesChange);

    // Inline boxes own independent child stories. Prepare their direct children
    // independently of whether the enclosing paragraph contains mathematics.
    // Notes/floating/nested boxes retain their existing staged placement paths.
    public static IEnumerable<DocxParagraph> EnumerateInlineBoxPreparationCandidates(DocxDocument document, CancellationToken token)
    {
        var parents = DocxBlockTraversal.EnumerateBodyParagraphs(document)
            .Concat(DocxBlockTraversal.EnumerateReferencedStaticStoryParagraphs(document.HeaderBodyElementsByType,
                document.HeaderParagraphsByType, document.HeaderParagraphs))
            .Concat(DocxBlockTraversal.EnumerateReferencedStaticStoryParagraphs(document.FooterBodyElementsByType,
                document.FooterParagraphsByType, document.FooterParagraphs))
            .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.PageSettings))
            .Concat(document.BodyElements.OfType<DocxSectionBreakElement>().SelectMany(section =>
                DocxBlockTraversal.EnumerateStaticStoryParagraphs(section.PageSettings)));
        var seen = new HashSet<DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var parent in parents)
        {
            token.ThrowIfCancellationRequested();
            foreach (var box in parent.InlineTextBoxes)
            {
                token.ThrowIfCancellationRequested();
                if (box.Revisions.Count != 0) continue;
                foreach (var element in box.BodyElements)
                {
                    token.ThrowIfCancellationRequested();
                    IEnumerable<DocxParagraph> children = element switch
                    {
                        DocxParagraphElement paragraph => [paragraph.Paragraph],
                        DocxTableElement table => EnumerateSupportedTableParagraphs(table.Table, token),
                        _ => []
                    };
                    foreach (var child in children)
                    {
                        token.ThrowIfCancellationRequested();
                        if (seen.Add(child)) yield return child;
                    }
                }
            }
        }
    }

    public static bool IsSupportedSimpleCell(DocxTableCell cell) =>
        IsSupportedCellProperties(cell) && cell.BodyElements.All(child => child is DocxParagraphElement);

    public static bool IsSupportedCellProperties(DocxTableCell cell) =>
        !cell.HasVerticalMerge && HasSupportedCellContentProperties(cell);

    internal static bool HasSupportedCellContentProperties(DocxTableCell cell) =>
        !cell.FitText &&
        HasSupportedTableContentRevisions(cell.Revisions, DocxRevisionKind.TableCellPropertiesChange) &&
        cell.TextDirectionValue is null or "lrTb";

    // The reader selects visible rows before layout. Row insertion/deletion
    // markers and resolved property changes do not replace child source owners.
    // Paragraph eligibility still checks its actual revisions; moving/unknown
    // owners and deleting cell wrappers retain their staged fallback.
    private static bool HasSupportedTableContentRevisions(IReadOnlyList<DocxRevisionInfo> revisions,
        DocxRevisionKind propertiesKind) => revisions.All(revision =>
            revision.Kind == DocxRevisionKind.Insertion || revision.Kind == propertiesKind ||
            propertiesKind == DocxRevisionKind.TableRowPropertiesChange && revision.Kind == DocxRevisionKind.Deletion ||
            propertiesKind == DocxRevisionKind.TableCellPropertiesChange && revision.Kind.IsCellStructureRevision());

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
        var noteTables = new HashSet<DocxTable>(ReferenceEqualityComparer.Instance);
        foreach (var story in document.RelatedStories)
        {
            token.ThrowIfCancellationRequested();
            if (story.Kind is not (DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote) || !story.IsNormalStoryType) continue;
            foreach (var table in story.Tables)
            { token.ThrowIfCancellationRequested(); noteTables.Add(table); }
        }
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
                    DocxTableElement table when noteTables.Contains(table.Table) => EnumerateSupportedTableParagraphs(table.Table, token),
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
        if (!HasSupportedOwner(paragraph) || DocxPlainBodyFrame.HasFixedLineFlow(paragraph) ||
            paragraph.ListLabel is not null || paragraph.Hyperlinks.Count != 0 ||
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
        if (DocxPlainBodyFrame.HasFixedLineFlow(paragraph)) { return true; }
        if (!DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph))
            return false;
        if (!DocxPlainBodyFrame.HasAutomaticLineFlow(paragraph)) { return true; }
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
        var display = GetMathOnlySourceEquations(paragraph, token);
        // A link on another authored row does not own these equations.
        // Equations inside the actual link keep their qualified inline style.
        if (DocxInlineContent.GetHyperlinkedEquations(paragraph, token) is { } linked)
            display.ExceptWith(linked);
        return display;
    }

    private static HashSet<DocxEquation> GetMathOnlySourceEquations(DocxParagraph paragraph,
        CancellationToken token) => GetSourceRowEquations(paragraph, token, withInlineOwner: false);

    public static IReadOnlySet<DocxEquation> GetMixedSourceEquations(DocxParagraph paragraph,
        CancellationToken token) => GetSourceRowEquations(paragraph, token, withInlineOwner: true);

    private static HashSet<DocxEquation> GetSourceRowEquations(DocxParagraph paragraph,
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
                case DocxInlineBox:
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
        paragraph.InlineTextBoxes.Count != 0 ||
        paragraph.Equations.Count != 0 && paragraph.LineSpacingPoints is not null &&
            (DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph) || DocxPlainBodyFrame.HasFixedLineFlow(paragraph)) ||
        paragraph.Equations.Count != 0 && paragraph.CommentRanges.Count != 0 ||
        paragraph.Equations.Count != 0 && paragraph.InlineReferences.Any(reference =>
            reference.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote) ||
        paragraph.Equations.Count != 0 && paragraph.Runs.Any(run => run.FieldKind is { } kind && kind.IsDynamic()) ||
        paragraph.Equations.Count == 1 && (GetInlineRow(paragraph) is null && GetStandaloneRow(paragraph) is null ||
            GetInlineRow(paragraph) is { } inline && prepared is not null &&
            prepared.TryGetValue(inline, out var math) && (math.Box.InlineBreakAfterGlyphs.Count != 0 ||
                math.Box.AllowsInlineRelocation && HasSimpleRelocationAlignment(paragraph)));

    // Preserve the qualified legacy single-equation placement. Mixed ordered
    // content additionally owns whitespace justification after line breaking.
    private static bool HasSimpleRelocationAlignment(DocxParagraph paragraph) =>
        paragraph.Alignment is DocxTextAlignment.Left or DocxTextAlignment.Center or DocxTextAlignment.Right;

    private static bool HasSupportedOrderedAlignment(DocxParagraph paragraph) =>
        HasSimpleRelocationAlignment(paragraph) || paragraph.Alignment == DocxTextAlignment.Justified;

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
            !(DocxPlainBodyFrame.HasContentDrivenLineFlow(paragraph) || DocxPlainBodyFrame.HasFixedLineFlow(paragraph))) return [];
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
        paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0;

    private static bool HasSupportedSourceOwner(DocxParagraph paragraph, bool hasPageInputs = false,
        CancellationToken token = default) =>
        paragraph.Equations.Count != 0 && DocxInlineContent.HasSupportedSourceOwner(paragraph, hasPageInputs, token);
}
