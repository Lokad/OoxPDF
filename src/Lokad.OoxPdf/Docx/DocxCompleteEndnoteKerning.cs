using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static class DocxCompleteEndnoteKerning
{
    internal static bool TryCreateDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument adjusted, out IReadOnlyDictionary<string, int> pages, CancellationToken cancellationToken)
    {
        adjusted = source;
        var pageById = new Dictionary<string, int>(StringComparer.Ordinal);
        pages = pageById;
        cancellationToken.ThrowIfCancellationRequested();
        if (!source.RelatedStories.Any(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote)) return false;
        if (source.RelatedStories.Any(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Footnote)) return false;
        var owners = new Dictionary<DocxRelatedStory, int>(ReferenceEqualityComparer.Instance);
        for (int page = 0; page < layout.Pages.Count; page++)
        foreach (DocxPlacedRelatedStoryLayout placed in layout.Pages[page].PlacedRelatedStories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (placed.UsesCompleteDocumentEndnoteFlow && !owners.TryAdd(placed.StoryLayout.Story, page)) return false;
        }
        var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        Dictionary<string, string?>? customMarks = null;
        foreach (DocxRelatedStory story in source.RelatedStories.Where(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (story.Id is null || story.Paragraphs.Count != 1 || !owners.TryGetValue(story, out int page) ||
                !pageById.TryAdd(story.Id, page)) return false;
            DocxParagraph paragraph = story.Paragraphs[0];
            if (!DocxPlainBodyFrame.IsPlainParagraph(paragraph) || paragraph.Runs.Count == 0) return false;
            DocxTextRun first = paragraph.Runs[0];
            int protectedPrefix = 0;
            if (!first.IsAutomaticNoteStoryMarker && first.VerticalAlignment != DocxRunVerticalAlignment.Superscript)
            {
                // Custom note text has no XML automatic-mark element. Resolve its
                // first character through one unique, source-owned body reference.
                // Other alignments and ambiguous/missing owners retain the fallback.
                if (first.VerticalAlignment != DocxRunVerticalAlignment.Baseline || first.SourceRunIndex < 0 ||
                    first.SourceTextOffsetInRun != 0) return false;
                customMarks ??= ResolveCustomMarks(source, cancellationToken);
                if (!customMarks.TryGetValue(story.Id, out string? mark) || mark is null ||
                    !first.Text.StartsWith(mark, StringComparison.Ordinal)) return false;
                protectedPrefix = mark.Length;
            }
            var runs = new List<DocxTextRun>(paragraph.Runs);
            bool changed = false;
            if (protectedPrefix > 0 && first.Text.Length > protectedPrefix)
            {
                runs[0] = first with { Text = first.Text[..protectedPrefix] };
                runs.Insert(1, first with { Text = first.Text[protectedPrefix..],
                    SourceTextOffsetInRun = first.SourceTextOffsetInRun + protectedPrefix });
                changed = true;
            }
            for (int i = 0; i < runs.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxTextRun run = runs[i];
                if (i == 0 && protectedPrefix > 0 || run.IsAutomaticNoteStoryMarker ||
                    run.VerticalAlignment != DocxRunVerticalAlignment.Baseline || run.Text.Length == 0) continue;
                if (!DocxKerningThreshold.TryResolve(run, out bool enabled)) return false;
                bool nominalAdvances = !enabled || OfficePdfTextEmissionProfile.FontSize(run.FontSize) != run.FontSize;
                if (enabled == run.PairKerningEnabled && nominalAdvances == run.UseLogicalSizeForPdfAdvances) continue;
                runs[i] = run with { PairKerningEnabled = enabled, UseLogicalSizeForPdfAdvances = nominalAdvances };
                changed = true;
            }
            if (changed) replacements.Add(paragraph, paragraph with { Runs = runs.ToArray() });
        }
        if (pageById.Count == 0 || replacements.Count == 0) return false;
        DocxParagraph Replace(DocxParagraph p) => replacements.GetValueOrDefault(p, p);
        adjusted = source with
        {
            RelatedStories = source.RelatedStories.Select(s => s with
            {
                BodyElements = s.BodyElements.Select(e => e is DocxParagraphElement p
                    ? p with { Paragraph = Replace(p.Paragraph) } : e).ToArray(),
                FallbackParagraphs = s.FallbackParagraphs.Select(Replace).ToArray(),
            }).ToArray(),
        };
        return true;
    }

    private static Dictionary<string, string?> ResolveCustomMarks(DocxDocument source, CancellationToken cancellationToken)
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
                if (reference.Kind != DocxRelatedStoryKind.Endnote || reference.Id is null) continue;
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

    internal static bool HasStablePages(DocxLayout layout, IReadOnlyDictionary<string, int> expected, CancellationToken cancellationToken)
    {
        var actual = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int page = 0; page < layout.Pages.Count; page++)
        foreach (DocxPlacedRelatedStoryLayout placed in layout.Pages[page].PlacedRelatedStories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRelatedStory story = placed.StoryLayout.Story;
            if (story.Kind != DocxRelatedStoryKind.Endnote || !story.IsNormalStoryType) continue;
            if (!placed.UsesCompleteDocumentEndnoteFlow || story.Id is null || !actual.TryAdd(story.Id, page)) return false;
        }
        return actual.Count == expected.Count && actual.All(p => expected.TryGetValue(p.Key, out int page) && page == p.Value);
    }
}
