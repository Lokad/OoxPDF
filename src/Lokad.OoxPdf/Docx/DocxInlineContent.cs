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
        paragraph.Equations.All(equation => equation.Revision is null) &&
        HasSupportedRunRevisionOwners(paragraph.Revisions, paragraph.Runs, token) &&
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

    public static bool HasSupportedRunRevisionOwners(IReadOnlyList<DocxRevisionInfo> revisions,
        IReadOnlyList<DocxTextRun> runs, CancellationToken token)
    {
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
        // equation measurement. Paragraph/mark, equation and nontext revisions
        // retain their existing fallbacks.
        // Resolve actual source ownership in linear work.
        foreach (var run in runs)
        {
            token.ThrowIfCancellationRequested();
            if (run.Text.Length == 0) continue;
            if (run.Revision is { } revision) remaining.Remove(revision);
            foreach (var item in run.Revisions)
            { token.ThrowIfCancellationRequested(); remaining.Remove(item); }
        }
        return remaining.Count == 0;
    }
}
