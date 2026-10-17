using System.Diagnostics.CodeAnalysis;

namespace Lokad.OoxPdf.Docx;

// A source view for the migrating paragraph path. Items retain the original
// resolved objects; bookmarks, links and revision ranges remain paragraph metadata.
// Null on the paragraph means its source owner has not migrated yet.
internal abstract record DocxInlineItem;
internal sealed record DocxInlineText(DocxTextRun Run, int Start, int Length) : DocxInlineItem;
internal sealed record DocxInlineTab(DocxTextRun Run, int Offset) : DocxInlineItem;
internal sealed record DocxInlineBreak(DocxTextRun Run, int Offset) : DocxInlineItem;
internal sealed record DocxInlinePicture(DocxInlineImage Image) : DocxInlineItem;
internal sealed record DocxInlineEquation(DocxEquation Equation) : DocxInlineItem;

internal static class DocxInlineContent
{
    // Source ownership is independent of whether a paragraph contains math.
    // Containers retain their own supported measurement and line-flow rules.
    public static bool HasSupportedSourceOwner(DocxParagraph paragraph, bool hasPageInputs = false,
        CancellationToken token = default) =>
        HasSupportedParagraphRevisionOwners(paragraph, token) &&
        paragraph.InlineTextBoxes.Count == 0 &&
        HasSupportedReferences(paragraph.InlineReferences, paragraph.Runs) &&
        (hasPageInputs ? HasSourceFieldResults(paragraph.FieldReferences, paragraph.Runs)
            : HasMaterializedFieldResults(paragraph.FieldReferences, paragraph.Runs)) &&
        paragraph.RevisionRanges.Count == 0;

    public static bool RequiresMeasuredSourceRanges(DocxParagraph paragraph) =>
        paragraph.CommentRanges.Count != 0 || paragraph.Hyperlinks.Count != 0;

    public static IReadOnlySet<DocxEquation>? GetHyperlinkedEquations(DocxParagraph paragraph, CancellationToken token)
    {
        if (paragraph.Hyperlinks.Count == 0 || paragraph.InlineContent is not { } items) return null;
        var linked = new HashSet<DocxEquation>(ReferenceEqualityComparer.Instance);
        int visitedThrough = 0;
        foreach (var link in paragraph.Hyperlinks.Where(link => link.StartInlineItemIndex is not null && link.EndInlineItemIndex is not null)
            .OrderBy(link => link.StartInlineItemIndex))
        {
            token.ThrowIfCancellationRequested();
            int start = Math.Max(visitedThrough, Math.Clamp(link.StartInlineItemIndex!.Value, 0, items.Count));
            int end = Math.Clamp(link.EndInlineItemIndex!.Value, 0, items.Count);
            for (int index = start; index < end; index++)
            {
                token.ThrowIfCancellationRequested();
                if (items[index] is DocxInlineEquation math) linked.Add(math.Equation);
            }
            visitedThrough = Math.Max(visitedThrough, end);
        }
        return linked;
    }

    public static bool TryGetAutomaticNoteRun(DocxInlineReference reference, [NotNullWhen(true)] out DocxTextRun? run)
    {
        run = reference.DisplayRun;
        return reference.Kind is (DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote) &&
            reference.NoteNumberRestartValue is null or "continuous" &&
            run is { IsAutomaticNoteReferenceMarker: true } && run.Text == reference.DisplayText;
    }

    public static bool HasSourceFieldResults(IReadOnlyList<DocxFieldReference> fields,
        IReadOnlyList<DocxTextRun> runs) => fields.All(field =>
            field.HasCachedResult && field.RendersCachedResult && !field.UsesPlaceholder ||
            field.Kind.IsDynamic() && field.UsesPlaceholder && !field.RendersCachedResult &&
                field.Placeholder == DocxFieldText.Placeholder(field.Kind)) &&
        runs.All(run => run.FieldKind is null || run.FieldKind is { } kind && kind.IsDynamic() &&
            run.Text == DocxFieldText.Placeholder(kind));

    // Cached field results already own ordinary resolved runs. Instructions and
    // page/count placeholders still need their separate dynamic layout path.
    public static bool HasMaterializedFieldResults(IReadOnlyList<DocxFieldReference> fields,
        IReadOnlyList<DocxTextRun> runs) => fields.All(field =>
            field.HasCachedResult && field.RendersCachedResult && !field.UsesPlaceholder) &&
        runs.All(run => run.FieldKind is null);

    public static bool HasSupportedReferences(IReadOnlyList<DocxInlineReference> references,
        IReadOnlyList<DocxTextRun> runs)
    {
        if (references.All(reference => reference.Kind == DocxRelatedStoryKind.Comment)) return true;
        var owners = new HashSet<DocxTextRun>(runs, ReferenceEqualityComparer.Instance);
        return references.All(reference => reference.Kind == DocxRelatedStoryKind.Comment ||
            TryGetAutomaticNoteRun(reference, out var run) && owners.Contains(run));
    }

    public static List<DocxInlineItem> FromPrefix(IReadOnlyList<DocxTextRun> runs,
        IReadOnlyList<DocxInlineImage> images, CancellationToken token)
    {
        var items = new List<DocxInlineItem>();
        var textGroups = runs.GroupBy(run => run.SourceRunIndex).ToDictionary(group => group.Key, group => group.ToArray());
        var imageGroups = images.GroupBy(image => image.SourceRunIndex).ToDictionary(group => group.Key, group => group.ToArray());
        foreach (int source in textGroups.Keys.Concat(imageGroups.Keys).Distinct().Order())
        {
            token.ThrowIfCancellationRequested();
            AppendRun(items, textGroups.GetValueOrDefault(source) ?? [], imageGroups.GetValueOrDefault(source) ?? [], token);
        }
        return items;
    }

    public static void AppendRun(List<DocxInlineItem> items, IEnumerable<DocxTextRun> runs,
        IEnumerable<DocxInlineImage> images, CancellationToken token)
    {
        var pictures = images.OrderBy(image => image.SourceTextOffsetInRun).ThenBy(image => image.SourceOrderInRun).ToArray();
        int picture = 0;
        foreach (DocxTextRun run in runs)
        {
            token.ThrowIfCancellationRequested();
            int start = 0;
            while (picture < pictures.Length && pictures[picture].SourceTextOffsetInRun <= run.SourceTextOffsetInRun + run.Text.Length)
            {
                int end = Math.Clamp(pictures[picture].SourceTextOffsetInRun - run.SourceTextOffsetInRun, start, run.Text.Length);
                AppendText(run, start, end);
                items.Add(new DocxInlinePicture(pictures[picture++]));
                start = end;
            }
            AppendText(run, start, run.Text.Length);
        }
        while (picture < pictures.Length)
        {
            token.ThrowIfCancellationRequested();
            items.Add(new DocxInlinePicture(pictures[picture++]));
        }

        void AppendText(DocxTextRun run, int start, int end)
        {
            int textStart = start;
            for (int offset = start; offset < end; offset++)
            {
                if ((offset & 255) == 0) token.ThrowIfCancellationRequested();
                char value = run.Text[offset];
                if (value is not ('\t' or '\n')) continue;
                if (offset > textStart) items.Add(new DocxInlineText(run, textStart, offset - textStart));
                items.Add(value == '\t' ? new DocxInlineTab(run, offset) : new DocxInlineBreak(run, offset));
                textStart = offset + 1;
            }
            if (end > textStart) items.Add(new DocxInlineText(run, textStart, end - textStart));
        }
    }

    public static bool HasVisibleText(IReadOnlyList<DocxInlineItem> items) => items.Any(item =>
        item is DocxInlineText text && !text.Run.EffectiveProperties.Hidden &&
        !text.Run.Text.AsSpan(text.Start, text.Length).IsWhiteSpace());

    public static bool TryGetEquationTextOffset(IReadOnlyList<DocxInlineItem> items, DocxEquation equation,
        out int offset)
    {
        offset = 0;
        foreach (DocxInlineItem item in items)
        {
            if (item is DocxInlineEquation math && ReferenceEquals(math.Equation, equation)) return true;
            offset += item switch { DocxInlineText text => text.Length, DocxInlineTab or DocxInlineBreak => 1, _ => 0 };
        }
        return false;
    }

    internal static bool HasOnlyNonRenderingRevisionMetadata(DocxParagraph paragraph,
        IReadOnlyList<DocxRevisionInfo>? revisions = null, CancellationToken token = default)
    {
        token.ThrowIfCancellationRequested();
        var source = revisions ?? paragraph.Revisions;
        if (source.Count == 0) return true;
        if (paragraph.RevisionRanges.Count != 0 ||
            paragraph.HasDeletedParagraphMark ||
            paragraph.Runs.Any(run => run.Revision is not null || run.Revisions.Count != 0) ||
            paragraph.Images.Any(image => image.Revisions.Count != 0) ||
            paragraph.InlineTextBoxes.Any(box => box.Revisions.Count != 0) ||
            paragraph.ParagraphMarkRun is { } mark && (mark.Revision is not null || mark.Revisions.Count != 0)) return false;
        foreach (var equation in paragraph.Equations)
        {
            token.ThrowIfCancellationRequested();
            // Structural paint retains authored ink. Only actual revised literal
            // sources can make this equation contribute a visible revision band.
            if ((equation.Revision is not null || equation.HasRunRevisions) &&
                (!DocxMathLayout.TryGetRuns(equation, token, out var runs) || runs.Count != 0)) return false;
        }
        return HasSupportedRunRevisionOwners(source, [], token, nonRenderingRunOwners: paragraph.NonRenderingRunOwners,
            equations: paragraph.Equations);
    }

    public static bool HasSupportedParagraphRevisionOwners(DocxParagraph paragraph, CancellationToken token,
        IReadOnlyList<DocxRevisionInfo>? revisions = null) =>
        HasSupportedRunRevisionOwners(revisions ?? paragraph.Revisions, paragraph.Runs, token,
            paragraph.ListLabel is null ? null : paragraph.ParagraphMarkRun, paragraph.Images, paragraph.NonRenderingRunOwners, paragraph.Equations);

    // Reference identity and measured source position establish actual owners.
    // The hint avoids rediscovering unchanged math; it cannot invent ownership.
    internal static bool TryGetEquationRevisionOwners(DocxEquation equation, CancellationToken token,
        out IReadOnlyList<DocxRevisionInfo> owners)
    {
        owners = [];
        token.ThrowIfCancellationRequested();
        if (equation.Revision is null && !equation.HasRunRevisions) return true;
        if (equation.SourceRunIndex < 0 || !DocxMathLayout.TryGetRuns(equation, token, out var runs)) return false;
        var actual = new List<DocxRevisionInfo>();
        var seen = new HashSet<DocxRevisionInfo>(ReferenceEqualityComparer.Instance);
        if (runs.Count == 0)
        {
            var controls = DocxMathLayout.GetControlSources(equation, token);
            if (controls.Count == 0) return false;
            foreach (var control in controls)
                if (control.ControlText is not { } source || !AddOwner(source)) return false;
        }
        else foreach (var source in runs)
        {
            token.ThrowIfCancellationRequested();
            if (!AddOwner(source.Run.Text)) return false;
        }
        if (equation.Revision is { } primary && !seen.Contains(primary)) return false;
        owners = actual;
        return true;

        bool AddOwner(DocxTextRun run)
        {
            token.ThrowIfCancellationRequested();
            if (run.SourceRunIndex != equation.SourceRunIndex ||
                equation.Revision is { } wrapper && !ReferenceEquals(run.Revision, wrapper)) return false;
            if (run.Revision is not { } revision) return run.Revisions.Count == 0;
            if (revision.Kind is not (DocxRevisionKind.Insertion or DocxRevisionKind.Deletion) ||
                run.Revisions.Count == 0 || run.Revisions.Any(owner => !ReferenceEquals(owner, revision))) return false;
            if (seen.Add(revision)) actual.Add(revision);
            return true;
        }
    }

    public static bool HasSupportedRunRevisionOwners(IReadOnlyList<DocxRevisionInfo> revisions,
        IReadOnlyList<DocxTextRun> runs, CancellationToken token, DocxTextRun? numberingMark = null, IReadOnlyList<DocxInlineImage>? images = null,
        IReadOnlyList<DocxNonRenderingRunOwner>? nonRenderingRunOwners = null, IReadOnlyList<DocxEquation>? equations = null)
    {
        List<DocxRevisionInfo>? equationOwners = null;
        if (equations is not null)
            foreach (var equation in equations)
            {
                if (!TryGetEquationRevisionOwners(equation, token, out var owners)) return false;
                if (owners.Count != 0) (equationOwners ??= []).AddRange(owners);
            }
        if (revisions.Count == 0) return true;
        var remaining = new HashSet<DocxRevisionInfo>(ReferenceEqualityComparer.Instance);
        foreach (var revision in revisions)
        {
            token.ThrowIfCancellationRequested();
            if (revision.Kind is not (DocxRevisionKind.Insertion or DocxRevisionKind.Deletion or DocxRevisionKind.RunPropertiesChange)) return false;
            remaining.Add(revision);
        }
        // Resolved text owns its inserted/deleted content and run formatting.
        // Formatting balloons consume the same run's metadata independently of
        // equation measurement. A numbered inserted mark has its separate owner
        // below; other paragraph/mark, unsupported equation and nontext revisions retain
        // their existing fallbacks.
        // Resolve actual source ownership in linear work.
        foreach (var run in runs)
        {
            token.ThrowIfCancellationRequested();
            if (run.Text.Length == 0) continue;
            if (run.Revision is { } revision) remaining.Remove(revision);
            foreach (var item in run.Revisions)
            { token.ThrowIfCancellationRequested(); remaining.Remove(item); }
        }
        if (nonRenderingRunOwners is not null)
            foreach (var owner in nonRenderingRunOwners)
            {
                token.ThrowIfCancellationRequested();
                if (owner.SourceRunIndex < 0) continue;
                foreach (var item in owner.Revisions)
                { token.ThrowIfCancellationRequested(); remaining.Remove(item); }
            }
        // A picture-only run retains its resolved source owner even without
        // printable text. Consume only that actual image's insertion metadata;
        // aggregate or value-equal revisions cannot invent ownership.
        if (images is not null)
            foreach (var image in images)
            {
                token.ThrowIfCancellationRequested();
                if (image.SourceRunStyle is not { } owner || owner.SourceRunIndex != image.SourceRunIndex) continue;
                foreach (var revision in image.Revisions)
                {
                    token.ThrowIfCancellationRequested();
                    if (revision.Kind == DocxRevisionKind.Insertion &&
                        (ReferenceEquals(owner.Revision, revision) || owner.Revisions.Any(item => ReferenceEquals(item, revision))))
                        remaining.Remove(revision);
                }
            }
        if (equationOwners is not null)
            foreach (var owner in equationOwners)
            { token.ThrowIfCancellationRequested(); remaining.Remove(owner); }
        // A parsed paragraph mark owns an inserted numbering prefix even when
        // none of its visible text runs is revised. Keep other mark kinds and
        // unrelated/value-equal metadata on their existing fallback paths.
        if (numberingMark is { SourceRunIndex: -1 } && string.IsNullOrWhiteSpace(numberingMark.Text))
            foreach (var revision in numberingMark.Revisions)
            {
                token.ThrowIfCancellationRequested();
                if (revision.Kind == DocxRevisionKind.Insertion && revision.SourceElement == "ins")
                    remaining.Remove(revision);
            }
        return remaining.Count == 0;
    }
}
