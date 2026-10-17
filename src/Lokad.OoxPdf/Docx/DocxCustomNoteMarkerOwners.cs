using System.Diagnostics.CodeAnalysis;

namespace Lokad.OoxPdf.Docx;

// Exact body character ownership is shared by note advances and line metrics.
internal static class DocxCustomNoteMarkerOwners
{
    internal static Dictionary<(int Run, int Offset), DocxTextRun?> IndexSourceRuns(
        IReadOnlyList<DocxTextRun> runs, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var owners = new Dictionary<(int Run, int Offset), DocxTextRun?>();
        foreach (var run in runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var key = (run.SourceRunIndex, run.SourceTextOffsetInRun);
            if (!owners.TryAdd(key, run)) owners[key] = null;
        }
        return owners;
    }

    internal static bool TryGetSourceRun(DocxInlineReference reference,
        IReadOnlyDictionary<(int Run, int Offset), DocxTextRun?> owners,
        [NotNullWhen(true)] out DocxTextRun? run)
    {
        run = null;
        if (reference.Kind is not (DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote) ||
            reference.CustomMarkText is not { Length: 1 } custom || char.IsWhiteSpace(custom[0]) ||
            char.IsSurrogate(custom[0]) || reference.Revision is not null || reference.Revisions.Count != 0 ||
            reference.SourceRunIndex < 0 || reference.TextOffsetInRun < 0 ||
            !owners.TryGetValue((reference.SourceRunIndex, reference.TextOffsetInRun), out var owner) ||
            owner is null || owner.Revision is not null || owner.Revisions.Count != 0 ||
            !owner.Text.StartsWith(custom, StringComparison.Ordinal)) return false;
        run = owner;
        return true;
    }

    internal static IReadOnlyDictionary<DocxRelatedStory, DocxTextRun> ResolveTargets(
        DocxDocument source, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var result = new Dictionary<DocxRelatedStory, DocxTextRun>(ReferenceEqualityComparer.Instance);
        var counts = new Dictionary<(DocxRelatedStoryKind Kind, string Id), int>();
        foreach (DocxRelatedStory story in source.RelatedStories.Where(s => s.IsNormalStoryType && s.Id is not null))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var key = (story.Kind, story.Id!);
            counts[key] = counts.GetValueOrDefault(key) + 1;
        }
        var marksByKind = new Dictionary<DocxRelatedStoryKind, Dictionary<string, string?>>();
        foreach (DocxRelatedStory story in source.RelatedStories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (story.Kind is not (DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote) || story.Id is null ||
                !story.IsNormalStoryType || counts[(story.Kind, story.Id)] != 1 || story.Paragraphs.Count != 1 ||
                story.Paragraphs[0].Runs.Count == 0 || story.Paragraphs[0].Runs[0].IsAutomaticNoteStoryMarker ||
                story.Paragraphs[0].Runs[0].VerticalAlignment != DocxRunVerticalAlignment.Baseline) continue;
            if (!marksByKind.TryGetValue(story.Kind, out var marks))
            {
                marks = Resolve(source, story.Kind, cancellationToken);
                marksByKind.Add(story.Kind, marks);
            }
            if (ResolveBaselineTarget(story, marks) is { } marker) result.Add(story, marker);
        }
        return result;
    }

    private static Dictionary<string, string?> Resolve(DocxDocument source,
        DocxRelatedStoryKind kind, CancellationToken cancellationToken)
    {
        var marks = new Dictionary<string, string?>(StringComparer.Ordinal);
        foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(source))
        {
            cancellationToken.ThrowIfCancellationRequested();
            Dictionary<(int Run, int Offset), DocxTextRun?>? runOwners = null;
            bool plain = paragraph.InlineReferences.Any(r => r.CustomMarkText is not null) &&
                DocxPlainBodyFrame.IsPlainParagraph(paragraph);
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reference.Kind != kind || reference.Id is null) continue;
                string? mark = null;
                if (reference.CustomMarkText is not null && plain &&
                    TryGetSourceRun(reference, runOwners ??= IndexSourceRuns(paragraph.Runs, cancellationToken), out _))
                    mark = reference.CustomMarkText;
                if (!marks.TryAdd(reference.Id, mark)) marks[reference.Id] = null;
            }
        }
        return marks;
    }

    internal static bool TryGetTargetRun(DocxRelatedStory story, string mark,
        [NotNullWhen(true)] out DocxTextRun? run)
    {
        run = null;
        if (!story.IsNormalStoryType || story.Id is null || story.BodyElements.Count != 1 ||
            story.Paragraphs.Count != 1 || story.Paragraphs[0].Runs.Count == 0) return false;
        var first = story.Paragraphs[0].Runs[0];
        if (first.IsAutomaticNoteStoryMarker || first.SourceRunIndex < 0 || first.SourceTextOffsetInRun != 0 ||
            !first.Text.StartsWith(mark, StringComparison.Ordinal)) return false;
        run = first;
        return true;
    }

    private static DocxTextRun? ResolveBaselineTarget(DocxRelatedStory story, IReadOnlyDictionary<string, string?> marks) =>
        story.Id is not null && marks.TryGetValue(story.Id, out string? mark) && mark is not null &&
        TryGetTargetRun(story, mark, out var first) && first.VerticalAlignment == DocxRunVerticalAlignment.Baseline
            ? first : null;
}
