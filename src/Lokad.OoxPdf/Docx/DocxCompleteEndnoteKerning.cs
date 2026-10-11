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
        foreach (DocxRelatedStory story in source.RelatedStories.Where(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (story.Id is null || story.Paragraphs.Count != 1 || !owners.TryGetValue(story, out int page) ||
                !pageById.TryAdd(story.Id, page)) return false;
            DocxParagraph paragraph = story.Paragraphs[0];
            // Explicit reader ownership admits baseline automatic marks without
            // inferring a marker from its font, text or size. Keep the previously
            // qualified superscript frame and its existing custom-mark fallback.
            if (!DocxPlainBodyFrame.IsPlainParagraph(paragraph) || paragraph.Runs.Count == 0 ||
                !paragraph.Runs[0].IsAutomaticNoteStoryMarker &&
                paragraph.Runs[0].VerticalAlignment != DocxRunVerticalAlignment.Superscript) return false;
            var runs = paragraph.Runs.ToArray();
            bool changed = false;
            for (int i = 0; i < runs.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                DocxTextRun run = runs[i];
                if (run.IsAutomaticNoteStoryMarker || run.VerticalAlignment != DocxRunVerticalAlignment.Baseline || run.Text.Length == 0) continue;
                if (!DocxKerningThreshold.TryResolve(run, out bool enabled)) return false;
                bool nominalAdvances = !enabled || OfficePdfTextEmissionProfile.FontSize(run.FontSize) != run.FontSize;
                if (enabled == run.PairKerningEnabled && nominalAdvances == run.UseLogicalSizeForPdfAdvances) continue;
                runs[i] = run with { PairKerningEnabled = enabled, UseLogicalSizeForPdfAdvances = nominalAdvances };
                changed = true;
            }
            if (changed) replacements.Add(paragraph, paragraph with { Runs = runs });
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
