using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static class DocxNoteKerning
{
    internal sealed record PageRange(int First, int Last, int Slices);

    internal static bool TryCreateDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument adjusted, out IReadOnlyDictionary<string, PageRange> pages, CancellationToken cancellationToken,
        DocxRelatedStoryKind kind = DocxRelatedStoryKind.Endnote)
    {
        adjusted = source;
        var pageById = new Dictionary<string, PageRange>(StringComparer.Ordinal);
        pages = pageById;
        cancellationToken.ThrowIfCancellationRequested();
        if (kind is not (DocxRelatedStoryKind.Endnote or DocxRelatedStoryKind.Footnote) ||
            !source.RelatedStories.Any(s => s.IsNormalStoryType && s.Kind == kind) ||
            source.RelatedStories.Any(s => s.IsNormalStoryType && s.Kind != kind &&
                s.Kind is DocxRelatedStoryKind.Endnote or DocxRelatedStoryKind.Footnote)) return false;
        if (!TryResolvePages(layout, kind, out var owners, cancellationToken)) return false;
        var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        Dictionary<string, string?>? customMarks = null;
        foreach (DocxRelatedStory story in source.RelatedStories.Where(s => s.IsNormalStoryType && s.Kind == kind))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (story.Id is null || story.Paragraphs.Count != 1 || !owners.TryGetValue(story, out PageRange? range) ||
                !pageById.TryAdd(story.Id, range)) return false;
            DocxParagraph paragraph = story.Paragraphs[0];
            if (!DocxPlainBodyFrame.IsPlainParagraph(paragraph) || paragraph.Runs.Count == 0) return false;
            DocxTextRun first = paragraph.Runs[0];
            // Custom baseline marks share exact source ownership. Other custom
            // footnote alignments retain their independently qualified fallback.
            if (kind == DocxRelatedStoryKind.Footnote && !first.IsAutomaticNoteStoryMarker &&
                first.VerticalAlignment != DocxRunVerticalAlignment.Baseline) return false;
            int protectedPrefix = 0;
            if (!first.IsAutomaticNoteStoryMarker && first.VerticalAlignment != DocxRunVerticalAlignment.Superscript)
            {
                // Custom note text has no XML automatic-mark element. Resolve its
                // first character through one unique, source-owned body reference.
                // Other alignments and ambiguous/missing owners retain the fallback.
                if (first.VerticalAlignment != DocxRunVerticalAlignment.Baseline || first.SourceRunIndex < 0 ||
                    first.SourceTextOffsetInRun != 0) return false;
                customMarks ??= ResolveCustomMarks(source, kind, cancellationToken);
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

    private static Dictionary<string, string?> ResolveCustomMarks(DocxDocument source,
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

    internal static bool HasStableOwners(DocxLayout layout, IReadOnlyDictionary<string, PageRange> expected, CancellationToken cancellationToken,
        DocxRelatedStoryKind kind = DocxRelatedStoryKind.Endnote)
    {
        if (!TryResolvePages(layout, kind, out var owners, cancellationToken)) return false;
        var actual = new Dictionary<string, PageRange>(StringComparer.Ordinal);
        foreach (var owner in owners)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (owner.Key.Id is null || !actual.TryAdd(owner.Key.Id, owner.Value)) return false;
        }
        // Corrected advances may change first placements and continuation ranges.
        // Qualified slices still require the same complete, unique source owners;
        // authored body references and numbering are unchanged by the retry.
        return actual.Count == expected.Count && actual.All(p => expected.TryGetValue(p.Key, out PageRange? range) &&
            range.First >= 0 && range.Last >= range.First && range.Slices > 0 &&
            range.Slices <= (long)range.Last - range.First + 1);
    }

    private static bool TryResolvePages(DocxLayout layout, DocxRelatedStoryKind kind,
        out Dictionary<DocxRelatedStory, PageRange> owners, CancellationToken cancellationToken)
    {
        owners = new(ReferenceEqualityComparer.Instance);
        var sourceLines = new Dictionary<DocxRelatedStory, IReadOnlyList<DocxTextLineLayout>>(ReferenceEqualityComparer.Instance);
        var nextLines = new Dictionary<DocxRelatedStory, int>(ReferenceEqualityComparer.Instance);
        foreach (DocxRelatedStoryLayout source in layout.RelatedStories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (source.Story.Kind != kind || !source.Story.IsNormalStoryType) continue;
            if (source.TextLines.Count == 0 || !sourceLines.TryAdd(source.Story, source.TextLines)) return false;
            nextLines.Add(source.Story, 0);
        }
        for (int page = 0; page < layout.Pages.Count; page++)
        foreach (DocxPlacedRelatedStoryLayout placed in layout.Pages[page].PlacedRelatedStories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRelatedStory story = placed.StoryLayout.Story;
            if (story.Kind != kind || !story.IsNormalStoryType) continue;
            bool admitted = kind == DocxRelatedStoryKind.Footnote
                ? placed.UsesPlainFootnoteFrame : placed.UsesPlainDocumentEndnoteTypography;
            if (!admitted || !sourceLines.TryGetValue(story, out var canonical) || placed.TextLines.Count == 0) return false;
            int next = nextLines[story];
            foreach (DocxTextLineLayout line in placed.StoryLayout.TextLines)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (next >= canonical.Count || line.SourceLineIndex != canonical[next].SourceLineIndex ||
                    !ReferenceEquals(line.SourceParagraph, canonical[next].SourceParagraph)) return false;
                next++;
            }
            nextLines[story] = next;
            if (owners.TryGetValue(story, out PageRange? range))
            {
                if (page <= range.Last || kind == DocxRelatedStoryKind.Endnote && placed.ContentTopOffset <= 0d) return false;
                owners[story] = range with { Last = page, Slices = range.Slices + 1 };
            }
            else
            {
                if (kind == DocxRelatedStoryKind.Endnote && placed.ContentTopOffset != 0d) return false;
                owners.Add(story, new(page, page, 1));
            }
        }
        return owners.Count > 0 && owners.Count == sourceLines.Count &&
            nextLines.All(p => p.Value == sourceLines[p.Key].Count);
    }
}
