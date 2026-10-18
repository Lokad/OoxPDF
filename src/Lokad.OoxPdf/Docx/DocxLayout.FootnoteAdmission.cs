namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    // Complete one-line notes have no continuation dependency. Their source
    // fragments can own admission directly instead of reserving an entire
    // paragraph's future notes before its first line arrives.
    private sealed class CompleteFootnoteAdmission(
        IReadOnlyDictionary<int, Dictionary<int, DocxInlineReferenceLocation>> owners,
        IReadOnlyDictionary<int, List<DocxInlineReferenceLocation>> tableOwners,
        double separatorHeight)
    {
        internal double SeparatorHeight => separatorHeight;

        internal IReadOnlyList<DocxInlineReferenceLocation> Select(int block,
            IEnumerable<DocxTextSpan>? spans, CancellationToken token)
        {
            // Tables retain their existing block admission. Only body paragraphs
            // defer registration until the actual source line arrives.
            if (spans is null) return tableOwners.TryGetValue(block, out var references) ? references : [];
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
        // Column/section changes and keep chains retain their existing
        // reservation/continuation contracts. Selected static stories and mixed
        // paragraph content do not change complete body-marker ownership.
        if (!DocxPlainBodyFrame.HasKnownBodyFrame(document) ||
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
        var tableOwners = new Dictionary<int, List<DocxInlineReferenceLocation>>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (block, references) in locations)
        {
            token.ThrowIfCancellationRequested();
            var paragraphs = IndexCompleteNoteParagraphs(document.BodyElements[block], token);
            var sourceRunsByParagraph = new Dictionary<DocxParagraph, Dictionary<(int Run, int Offset), DocxTextRun?>>(ReferenceEqualityComparer.Instance);
            foreach (var location in references.Where(location => location.Reference.Kind == DocxRelatedStoryKind.Footnote))
            {
                token.ThrowIfCancellationRequested();
                var reference = location.Reference;
                var paragraph = location.SourceParagraph;
                if (paragraphs is null || !paragraphs.Contains(paragraph) ||
                    paragraph.KeepRules.KeepNext == true || paragraph.KeepRules.KeepLines == true ||
                    block > 0 && document.BodyElements[block - 1] is DocxParagraphElement { Paragraph.KeepRules.KeepNext: true } ||
                    reference.Id is not { } id || !ids.Add(id) || !notes.TryGetValue(id, out var note) ||
                    reference.Revision is not null || reference.Revisions.Count != 0) return null;
                if (!sourceRunsByParagraph.TryGetValue(paragraph, out var sourceRuns))
                    sourceRunsByParagraph.Add(paragraph, sourceRuns = DocxCustomNoteMarkerOwners.IndexSourceRuns(paragraph.Runs, token));
                if (!sourceRuns.TryGetValue((reference.SourceRunIndex, reference.TextOffsetInRun), out var run) ||
                    run is null || run.EffectiveProperties.Hidden) return null;
                if (reference.CustomMarkText is { } custom)
                {
                    if (!DocxCustomNoteMarkerOwners.TryGetSourceRun(reference, sourceRuns, out _) ||
                        !DocxCustomNoteMarkerOwners.TryGetTargetRun(note.Story, custom, out var target) ||
                        target.EffectiveProperties.Hidden) return null;
                }
                // Physical admission needs the owned body reference and unique
                // target story, even when that story has no painted number.
                // Numbering and custom-mark rewriting keep their own checks.
                else if (!DocxInlineContent.TryGetAutomaticNoteRun(reference, out var marker) || !ReferenceEquals(run, marker)) return null;
                if (document.BodyElements[block] is DocxTableElement)
                {
                    if (!tableOwners.TryGetValue(block, out var indexed)) tableOwners.Add(block, indexed = []);
                    indexed.Add(location);
                }
                else
                {
                    if (!owners.TryGetValue(block, out var indexed)) owners.Add(block, indexed = []);
                    if (!indexed.TryAdd(reference.SourceRunIndex, location)) return null;
                }
            }
        }
        // Pure tables continue to use their existing reservation contract. This
        // migration concerns paragraphs whose notes arrive on measured lines.
        return owners.Count != 0 && ids.Count == notes.Count
            ? new(owners, tableOwners, separators[0].TextLines[0].LineHeight!.Value) : null;
    }

    private static HashSet<DocxParagraph>? IndexCompleteNoteParagraphs(DocxBodyElement source, CancellationToken token)
    {
        var paragraphs = new HashSet<DocxParagraph>(ReferenceEqualityComparer.Instance);
        bool Add(DocxParagraphElement element) =>
            DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(element.Paragraph, element.Revisions, token) &&
            paragraphs.Add(element.Paragraph);
        if (source is DocxParagraphElement paragraph) return Add(paragraph) ? paragraphs : null;
        if (source is not DocxTableElement table || source.Revisions.Count != 0 || table.Table.Revisions.Count != 0) return null;
        foreach (var row in table.Table.Rows)
        {
            token.ThrowIfCancellationRequested();
            if (row.Revisions.Count != 0) return null;
            foreach (var cell in row.Cells)
            {
                token.ThrowIfCancellationRequested();
                if (cell.Revisions.Count != 0) return null;
                foreach (var element in DocxTableCellContent.GetBodyElements(cell))
                {
                    token.ThrowIfCancellationRequested();
                    if (element is DocxParagraphElement cellParagraph)
                    {
                        if (!Add(cellParagraph)) return null;
                    }
                    else if (element is not (DocxPageBreakElement or DocxImplicitParagraphElement) || element.Revisions.Count != 0)
                        return null; // Nested and revised containers retain their existing fallback.
                }
            }
        }
        return paragraphs;
    }
}
