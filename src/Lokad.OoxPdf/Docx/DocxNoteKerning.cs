using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal static class DocxNoteKerning
{
    internal sealed record PageRange(int First, int Last, int Slices);

    private sealed class OwnedPageRanges(IReadOnlyList<DocxRelatedStory> sourceStories) : Dictionary<string, PageRange>(StringComparer.Ordinal)
    {
        internal IReadOnlyList<DocxRelatedStory> SourceStories { get; } = sourceStories;
    }

    internal sealed record JointPageRanges(
        IReadOnlyDictionary<string, PageRange> Footnotes,
        IReadOnlyDictionary<string, PageRange> Endnotes);

    internal static bool TryCreateJointDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument adjusted, out JointPageRanges? pages, CancellationToken cancellationToken)
    {
        adjusted = source;
        pages = null;
        cancellationToken.ThrowIfCancellationRequested();
        if (!source.RelatedStories.Any(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Footnote) ||
            !source.RelatedStories.Any(s => s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote)) return false;
        // Both independent plans inspect the same complete original owners.
        // The renderer measures their combined document in one bounded retry.
        // Check the complementary admission before cloning the endnote plan,
        // so the established single-kind fallback incurs no duplicate clone.
        if (!TryPlanDocument(source, layout, out DocxDocument footnotes, out var footnotePages,
                cancellationToken, DocxRelatedStoryKind.Footnote, allowMixedFootnote: true) ||
            !TryPlanDocument(source, layout, out DocxDocument endnotes, out var endnotePages,
                cancellationToken, DocxRelatedStoryKind.Endnote, allowMixedFootnote: false) ||
            ReferenceEquals(source, footnotes) && ReferenceEquals(source, endnotes)) return false;
        adjusted = source with
        {
            RelatedStories = source.RelatedStories.Select((story, index) => story.Kind switch
            {
                DocxRelatedStoryKind.Footnote => footnotes.RelatedStories[index],
                DocxRelatedStoryKind.Endnote => endnotes.RelatedStories[index],
                _ => story,
            }).ToArray(),
        };
        IReadOnlyList<DocxRelatedStory> adjustedSources = adjusted.RelatedStories;
        IReadOnlyDictionary<string, PageRange> Rebind(IReadOnlyDictionary<string, PageRange> expected)
        {
            var rebound = new OwnedPageRanges(adjustedSources);
            foreach (var entry in expected) rebound.Add(entry.Key, entry.Value);
            return rebound;
        }
        pages = new(Rebind(footnotePages), Rebind(endnotePages));
        return true;
    }

    internal static bool HasStableJointOwners(DocxLayout layout, JointPageRanges expected, CancellationToken cancellationToken) =>
        HasStableOwners(layout, expected.Footnotes, cancellationToken, DocxRelatedStoryKind.Footnote) &&
        HasStableOwners(layout, expected.Endnotes, cancellationToken, DocxRelatedStoryKind.Endnote);

    internal static bool TryCreateDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument adjusted, out IReadOnlyDictionary<string, PageRange> pages, CancellationToken cancellationToken,
        DocxRelatedStoryKind kind = DocxRelatedStoryKind.Endnote) =>
        TryPlanDocument(source, layout, out adjusted, out pages, cancellationToken, kind, allowMixedFootnote: false) &&
        !ReferenceEquals(source, adjusted);

    private static bool TryPlanDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument adjusted, out IReadOnlyDictionary<string, PageRange> pages, CancellationToken cancellationToken,
        DocxRelatedStoryKind kind, bool allowMixedFootnote)
    {
        adjusted = source;
        var pageById = new OwnedPageRanges(source.RelatedStories);
        pages = pageById;
        cancellationToken.ThrowIfCancellationRequested();
        if (kind is not (DocxRelatedStoryKind.Endnote or DocxRelatedStoryKind.Footnote) ||
            !source.RelatedStories.Any(s => s.IsNormalStoryType && s.Kind == kind)) return false;
        // A standalone mixed-footnote retry keeps the qualified endnote path
        // independent. The joint path separately verifies both admissions.
        if (!allowMixedFootnote && kind == DocxRelatedStoryKind.Footnote && source.RelatedStories.Any(s =>
            s.IsNormalStoryType && s.Kind == DocxRelatedStoryKind.Endnote)) return false;
        if (!HasCompleteOtherNoteOwners(layout, kind, cancellationToken, source.RelatedStories)) return false;
        if (!TryResolvePages(layout, kind, out var owners, cancellationToken)) return false;
        var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        IReadOnlyDictionary<DocxRelatedStory, DocxTextRun>? customTargets = null;
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
                customTargets ??= DocxCustomNoteMarkerOwners.ResolveTargets(source, cancellationToken);
                if (!ReferenceEquals(first, customTargets.GetValueOrDefault(story))) return false;
                protectedPrefix = 1;
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
        if (pageById.Count == 0) return false;
        if (replacements.Count == 0) return true;
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

    internal static bool HasStableOwners(DocxLayout layout, IReadOnlyDictionary<string, PageRange> expected, CancellationToken cancellationToken,
        DocxRelatedStoryKind kind = DocxRelatedStoryKind.Endnote)
    {
        if (!HasCompleteOtherNoteOwners(layout, kind, cancellationToken, (expected as OwnedPageRanges)?.SourceStories) ||
            !TryResolvePages(layout, kind, out var owners, cancellationToken)) return false;
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

    private static DocxRelatedStoryKind OtherKind(DocxRelatedStoryKind kind) =>
        kind == DocxRelatedStoryKind.Footnote ? DocxRelatedStoryKind.Endnote : DocxRelatedStoryKind.Footnote;

    private static bool HasCompleteOtherNoteOwners(DocxLayout layout, DocxRelatedStoryKind kind,
        CancellationToken cancellationToken, IReadOnlyList<DocxRelatedStory>? sourceStories = null)
    {
        DocxRelatedStoryKind other = OtherKind(kind);
        Dictionary<string, DocxRelatedStory>? expected = null;
        if (sourceStories is not null)
        {
            expected = new(StringComparer.Ordinal);
            foreach (DocxRelatedStory story in sourceStories.Where(s => s.IsNormalStoryType && s.Kind == other))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (story.Id is null || !expected.TryAdd(story.Id, story) || story.Paragraphs.Count != 1) return false;
            }
        }
        if (!layout.RelatedStories.Any(s => s.Story.IsNormalStoryType && s.Story.Kind == other)) return expected is null || expected.Count == 0;
        // The complementary kind keeps its established placement/typography.
        // It must still have complete, unique, plain source owners on every slice.
        if (!TryResolvePages(layout, other, out var owners, cancellationToken, requireAdmission: false)) return false;
        if (expected is not null && expected.Count != owners.Count) return false;
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (DocxRelatedStory story in owners.Keys)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (story.Id is null || !ids.Add(story.Id) || story.BodyElements.Count != 1 ||
                story.Paragraphs.Count != 1 || !DocxPlainBodyFrame.IsPlainParagraph(story.Paragraphs[0]) ||
                expected is not null && (!expected.TryGetValue(story.Id, out var original) ||
                    !ReferenceEquals(original.Paragraphs[0], story.Paragraphs[0]))) return false;
        }
        return true;
    }

    private static bool TryResolvePages(DocxLayout layout, DocxRelatedStoryKind kind,
        out Dictionary<DocxRelatedStory, PageRange> owners, CancellationToken cancellationToken,
        bool requireAdmission = true)
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
            if (requireAdmission && !admitted || !sourceLines.TryGetValue(story, out var canonical) || placed.TextLines.Count == 0) return false;
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
                if (page <= range.Last || requireAdmission && kind == DocxRelatedStoryKind.Endnote && placed.ContentTopOffset <= 0d) return false;
                owners[story] = range with { Last = page, Slices = range.Slices + 1 };
            }
            else
            {
                if (requireAdmission && kind == DocxRelatedStoryKind.Endnote && placed.ContentTopOffset != 0d) return false;
                owners.Add(story, new(page, page, 1));
            }
        }
        return owners.Count > 0 && owners.Count == sourceLines.Count &&
            nextLines.All(p => p.Value == sourceLines[p.Key].Count);
    }
}
