namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Complete one-line notes have no continuation dependency. Their source
    // fragments can own admission directly instead of reserving an entire
    // paragraph's future notes before its first line arrives.
    private sealed class CompleteFootnoteAdmission(
        IReadOnlyDictionary<int, Dictionary<int, DocxInlineReferenceLocation>> owners)
    {
        internal IReadOnlyList<DocxInlineReferenceLocation> Select(int block,
            IEnumerable<DocxTextSpan> spans, CancellationToken token)
        {
            if (!owners.TryGetValue(block, out var runs)) return [];
            List<DocxInlineReferenceLocation>? selected = null;
            foreach (var span in spans)
            {
                token.ThrowIfCancellationRequested();
                if (runs.TryGetValue(span.SourceTextRunIndex, out var location) &&
                    span.SourceTextOffsetInRun <= location.Reference.TextOffsetInRun &&
                    location.Reference.TextOffsetInRun < (long)span.SourceTextOffsetInRun + span.Text.Length)
                    (selected ??= []).Add(location);
            }
            return selected is null ? [] : selected;
        }
    }

    private static CompleteFootnoteAdmission? CreateCompleteFootnoteAdmission(DocxDocument document,
        IReadOnlyList<DocxRelatedStoryLayout> stories,
        IReadOnlyDictionary<int, List<DocxInlineReferenceLocation>> locations, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (!stories.Any(story => story.Story.Kind == DocxRelatedStoryKind.Footnote && story.Story.IsNormalStoryType)) return null;
        // Column/section changes, selected static frames and keep chains retain
        // their existing reservation/continuation contracts. Unrelated tables or
        // mixed paragraph content do not change complete body-marker ownership.
        if (!DocxPlainBodyFrame.HasKnownBodyFrame(document) || !DocxPlainBodyFrame.HasNoStaticStories(document) ||
            document.FinalSectionBreak!.ColumnCountValue is not (null or "1") ||
            document.FinalSectionBreak.ColumnDefinitions.Count > 1 ||
            document.FinalSectionBreak.PageSettings.FootnoteReferenceSettings.PositionValue is not (null or "pageBottom")) return null;
        foreach (var element in document.BodyElements)
        {
            token.ThrowIfCancellationRequested();
            if (element is DocxSectionBreakElement) return null;
        }
        var separators = stories.Where(story => story.Story.Kind == DocxRelatedStoryKind.Footnote &&
            story.Story.Type == DocxRelatedStoryType.Separator).ToArray();
        if (separators.Length != 1 || separators[0].TextLines.Count != 1 ||
            !separators[0].TextLines[0].Text.All(char.IsWhiteSpace) ||
            separators[0].TextLines[0].LineHeight is not > 0d ||
            !double.IsFinite(separators[0].TextLines[0].LineHeight!.Value)) return null;
        var notes = new Dictionary<string, DocxRelatedStoryLayout>(StringComparer.Ordinal);
        foreach (var story in stories.Where(story => story.Story.Kind == DocxRelatedStoryKind.Footnote && story.Story.IsNormalStoryType))
        {
            token.ThrowIfCancellationRequested();
            if (story.Story.Id is not { } id || !notes.TryAdd(id, story) || story.TextLines.Count != 1 ||
                story.Story.FloatingDrawings.Count != 0 || story.Story.Tables.Count != 0 ||
                story.Story.Paragraphs.Count != 1 || !DocxPlainBodyFrame.IsPlainParagraph(story.Story.Paragraphs[0])) return null;
            double[] boxes = GetStoryTextLineBoxHeights(story);
            if (boxes.Length != 1 || !double.IsFinite(boxes[0]) || boxes[0] <= 0d) return null;
        }
        if (notes.Count == 0) return null;
        var owners = new Dictionary<int, Dictionary<int, DocxInlineReferenceLocation>>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (block, references) in locations)
        {
            token.ThrowIfCancellationRequested();
            Dictionary<(int Run, int Offset), DocxTextRun?>? sourceRuns = null;
            foreach (var location in references.Where(location => location.Reference.Kind == DocxRelatedStoryKind.Footnote))
            {
                token.ThrowIfCancellationRequested();
                var reference = location.Reference;
                var paragraph = location.SourceParagraph;
                if (document.BodyElements[block] is not DocxParagraphElement element || !ReferenceEquals(element.Paragraph, paragraph) ||
                    paragraph.KeepRules.KeepNext == true || paragraph.KeepRules.KeepLines == true ||
                    block > 0 && document.BodyElements[block - 1] is DocxParagraphElement { Paragraph.KeepRules.KeepNext: true } ||
                    !DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph, element.Revisions) ||
                    reference.Id is not { } id || !ids.Add(id) || !notes.TryGetValue(id, out var note) ||
                    reference.Revision is not null || reference.Revisions.Count != 0) return null;
                sourceRuns ??= DocxCustomNoteMarkerOwners.IndexSourceRuns(paragraph.Runs, token);
                if (!sourceRuns.TryGetValue((reference.SourceRunIndex, reference.TextOffsetInRun), out var run) ||
                    run is null || run.EffectiveProperties.Hidden) return null;
                if (reference.CustomMarkText is { } custom)
                {
                    if (!DocxCustomNoteMarkerOwners.TryGetSourceRun(reference, sourceRuns, out _) ||
                        !DocxCustomNoteMarkerOwners.TryGetTargetRun(note.Story, custom, out var target) ||
                        target.EffectiveProperties.Hidden) return null;
                }
                else if (!DocxInlineContent.TryGetAutomaticNoteRun(reference, out var marker) || !ReferenceEquals(run, marker) ||
                    note.Story.Paragraphs[0].Runs.Count == 0 || note.Story.Paragraphs[0].Runs[0].Text != reference.DisplayText) return null;
                if (!owners.TryGetValue(block, out var indexed)) owners.Add(block, indexed = []);
                if (!indexed.TryAdd(reference.SourceRunIndex, location)) return null;
            }
        }
        return ids.Count == notes.Count ? new(owners) : null;
    }
}
