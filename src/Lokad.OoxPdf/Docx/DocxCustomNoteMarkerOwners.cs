namespace Lokad.OoxPdf.Docx;

// Exact body character ownership is shared by note advances and line metrics.
internal static class DocxCustomNoteMarkerOwners
{
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
                if (reference.CustomMarkText is { Length: 1 } custom && !char.IsWhiteSpace(custom[0]) &&
                    !char.IsSurrogate(custom[0]) && reference.Revision is null && reference.Revisions.Count == 0 &&
                    reference.SourceRunIndex >= 0 && reference.TextOffsetInRun >= 0 && plain)
                {
                    if (runOwners is null)
                    {
                        runOwners = new();
                        foreach (DocxTextRun run in paragraph.Runs)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            var key = (run.SourceRunIndex, run.SourceTextOffsetInRun);
                            if (!runOwners.TryAdd(key, run)) runOwners[key] = null;
                        }
                    }
                    if (runOwners.TryGetValue((reference.SourceRunIndex, reference.TextOffsetInRun), out DocxTextRun? owner) &&
                        owner is not null && owner.Text.StartsWith(custom, StringComparison.Ordinal)) mark = custom;
                }
                if (!marks.TryAdd(reference.Id, mark)) marks[reference.Id] = null;
            }
        }
        return marks;
    }

    private static DocxTextRun? ResolveBaselineTarget(DocxRelatedStory story, IReadOnlyDictionary<string, string?> marks)
    {
        if (!story.IsNormalStoryType || story.Id is null || story.BodyElements.Count != 1 ||
            story.Paragraphs.Count != 1 || story.Paragraphs[0].Runs.Count == 0) return null;
        DocxTextRun first = story.Paragraphs[0].Runs[0];
        return !first.IsAutomaticNoteStoryMarker && first.VerticalAlignment == DocxRunVerticalAlignment.Baseline &&
            first.SourceRunIndex >= 0 && first.SourceTextOffsetInRun == 0 && marks.TryGetValue(story.Id, out string? mark) &&
            mark is not null && first.Text.StartsWith(mark, StringComparison.Ordinal) ? first : null;
    }
}
