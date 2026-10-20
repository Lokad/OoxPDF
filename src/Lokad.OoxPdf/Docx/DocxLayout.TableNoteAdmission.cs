namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private sealed record DocxTableNoteFragmentPlan(double FirstHeight, IReadOnlyList<double> PageHeights,
        IReadOnlyList<double> SeparatorHeights, IReadOnlyDictionary<DocxInlineReference, int> Owners);

    // Index complete, placement-independent heads once. A bounded refinement
    // consumes measured source lines; unknown or continued owners keep the
    // existing conservative capacity. This state belongs to one table layout.
    private sealed class DocxTableNoteAdmission
    {
        private readonly InFlightRelatedStory[] notes;
        private readonly Dictionary<DocxParagraph, Dictionary<int, List<InFlightRelatedStory>>> bySource =
            new(ReferenceEqualityComparer.Instance);

        private DocxTableNoteAdmission(InFlightRelatedStory[] notes, CancellationToken token)
        {
            this.notes = notes;
            foreach (var note in notes)
            {
                token.ThrowIfCancellationRequested();
                var location = note.Location;
                if (!bySource.TryGetValue(location.SourceParagraph, out var runs))
                    bySource.Add(location.SourceParagraph, runs = []);
                if (!runs.TryGetValue(location.Reference.SourceRunIndex, out var references))
                    runs.Add(location.Reference.SourceRunIndex, references = []);
                references.Add(note);
            }
        }

        internal static DocxTableNoteAdmission? TryCreate(int sourceBlock,
            IEnumerable<InFlightRelatedStory> pending, IReadOnlyList<DocxInlineReferenceLocation>? locations,
            CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var notes = pending.ToArray();
            if (notes.Length == 0 || locations is null || !TryCompleteFootnoteReserve(notes, out _)) return null;
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var location in locations.Where(location => location.Reference.Kind == DocxRelatedStoryKind.Footnote))
            {
                token.ThrowIfCancellationRequested();
                if (location.Reference.Id is not { } id || !ids.Add(id)) return null;
            }
            if (ids.Count != notes.Length) return null;
            foreach (var note in notes)
            {
                token.ThrowIfCancellationRequested();
                if (note.SourceBlockIndex != sourceBlock || note.Location.Reference.SourceRunIndex < 0 ||
                    note.Location.Reference.TextOffsetInRun < 0 ||
                    (DocxFieldText.GetDependencies(note.StoryLayout.Story.BodyElements, token) & DocxCounterDependencies.PageNumber) != 0)
                    return null;
            }
            return new(notes, token);
        }

        internal DocxTableNoteFragmentPlan? TryPlan(IReadOnlyList<IReadOnlyList<DocxTextLineLayout>> fragments,
            Func<double, double> frameBottom, double firstTop, double pageTop, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (!TryCompleteFootnoteReserve(notes, out _)) return null;
            var owners = new Dictionary<DocxInlineReference, int>(ReferenceEqualityComparer.Instance);
            var grouped = Enumerable.Range(0, fragments.Count).Select(_ => new List<InFlightRelatedStory>()).ToArray();
            for (int index = 0; index < fragments.Count; index++)
            foreach (var line in fragments[index])
            {
                token.ThrowIfCancellationRequested();
                if (line.SourceParagraph is null || !bySource.TryGetValue(line.SourceParagraph, out var runs)) continue;
                foreach (var segment in line.Segments)
                {
                    if (!runs.TryGetValue(segment.SourceTextRunIndex, out var references)) continue;
                    foreach (var note in references)
                    {
                        int offset = note.Location.Reference.TextOffsetInRun;
                        if (offset < Math.Max(0, segment.SourceTextOffsetInRun) ||
                            offset >= (long)Math.Max(0, segment.SourceTextOffsetInRun) + segment.Text.Length) continue;
                        if (owners.TryGetValue(note.Location.Reference, out int previous))
                        { if (previous != index) return null; }
                        else
                        {
                            owners.Add(note.Location.Reference, index);
                            grouped[index].Add(note);
                        }
                    }
                }
            }
            if (owners.Count != notes.Length) return null;
            var heights = new double[fragments.Count];
            var separatorHeights = new double[fragments.Count];
            double firstHeight = 0d;
            for (int index = 0; index < grouped.Length; index++)
            {
                token.ThrowIfCancellationRequested();
                if (!TryCompleteFootnoteReserve(grouped[index], out double reserve)) return null;
                separatorHeights[index] = grouped[index].Count == 0 ? 0d :
                    grouped[index][0].SeparatorLayout!.TextLines[0].LineHeight!.Value;
                double bottom = frameBottom(reserve);
                heights[index] = pageTop - bottom;
                if (index == 0) firstHeight = firstTop - bottom;
                if (!double.IsFinite(heights[index]) || heights[index] <= 0d) return null;
            }
            return double.IsFinite(firstHeight) && firstHeight > 0d ? new(firstHeight, heights, separatorHeights, owners) : null;
        }
    }
}
