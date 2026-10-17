using System.Globalization;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Docx;

internal enum DocxFootnoteOwnerScope { SingleReferenceParagraphs, AllSourceReferences }

internal static class DocxPageFootnoteNumbering
{
    internal static bool IsEligibleDocument(DocxDocument document, bool includeContinuousRestart = false)
    {
        return DocxPlainBodyFrame.IsEligible(document) &&
            HasEligibleNoteSettingsAndStories(document, includeContinuousRestart);
    }

    // Numbering owns measured marker/note pages. Font and leading purposes keep
    // their existing plain-body admission through IsEligibleDocument above.
    private static bool IsEligibleInlineNumberingDocument(DocxDocument document) =>
        DocxPlainBodyFrame.HasKnownBodyFrame(document) && DocxPlainBodyFrame.HasNoStaticStories(document) &&
        HasEligibleNoteSettingsAndStories(document, includeContinuousRestart: false) &&
        document.BodyElements.Count != 0 && document.BodyElements.All(element =>
            element is DocxParagraphElement paragraph
                ? DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph.Paragraph, element.Revisions)
                : element is DocxPageBreakElement && element.Revisions.Count == 0) &&
        document.Paragraphs.All(paragraph => DocxPlainBodyFrame.IsPlainParagraph(paragraph) ||
            paragraph.InlineContent is not null && DocxInlineContent.HasSupportedSourceOwner(paragraph));

    private static bool HasEligibleNoteSettingsAndStories(DocxDocument document, bool includeContinuousRestart)
    {
        if (document.FinalSectionBreak is null || !document.Settings.FootnoteReferenceSettings.HasKnownNumberRestart) return false;
        var settings = document.FinalSectionBreak.PageSettings.FootnoteReferenceSettings;
        return (settings.NumberRestartValue == "eachPage" || includeContinuousRestart &&
                settings.NumberRestartValue is "continuous" or "eachSect") &&
            settings.NumberFormatValue is null or "decimal" or "decimalZero" &&
            (settings.NumberStartValue is null || settings.NumberStart is > 0 and <= 32767) &&
            settings.PositionValue is null or "pageBottom" &&
            document.RelatedStories.All(s => s.Kind == DocxRelatedStoryKind.Footnote &&
                s.FloatingDrawings.Count == 0 && s.Tables.Count == 0 && s.Paragraphs.All(DocxPlainBodyFrame.IsPlainParagraph));
    }

    internal static bool TryCreateDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument numbered, out IReadOnlyDictionary<string, int> referencePages, bool allowWrappedBody = false,
        DocxWrappedCalibriPolicy calibriPolicy = DocxWrappedCalibriPolicy.NormalizedNominal,
        DocxFootnoteOwnerScope ownerScope = DocxFootnoteOwnerScope.SingleReferenceParagraphs)
    {
        numbered = source;
        referencePages = new Dictionary<string, int>();
        if ((!IsEligibleDocument(source) && !IsEligibleInlineNumberingDocument(source)) ||
            !TryMapReferencePages(source, layout, out var pages, allowWrappedBody,
            calibriPolicy: calibriPolicy, ownerScope: ownerScope)) return false;
        var settings = source.FinalSectionBreak!.PageSettings.FootnoteReferenceSettings;
        var counts = new Dictionary<int, int>();
        var labels = new Dictionary<string, (string Before, string After)>(StringComparer.Ordinal);
        var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var paragraph in source.Paragraphs)
        {
            if (paragraph.InlineReferences.Count == 0) continue;
            var runLabels = new Dictionary<int, string>();
            var references = new List<DocxInlineReference>();
            foreach (var reference in paragraph.InlineReferences)
            {
                int page = pages[reference.Id!];
                counts.TryGetValue(page, out int count);
                int value = checked((settings.NumberStart ?? 1) + count);
                counts[page] = count + 1;
                string label = value.ToString(settings.NumberFormatValue == "decimalZero" ? "D2" : "D", CultureInfo.InvariantCulture);
                runLabels.Add(reference.SourceRunIndex, label);
                references.Add(reference with { DisplayText = label });
                labels.Add(reference.Id!, (reference.DisplayText!, label));
            }
            // Only the generated marker in its own source run changes. Authored
            // digits and original source offsets retain their identities.
            if (!TryReplaceMarkerRuns(paragraph, runLabels, references, out var replacement)) return false;
            replacements.Add(paragraph, replacement);
        }
        // Complete ownership was validated, but unchanged labels need neither
        // a replacement document nor the coordinator's font/layout retry.
        if (labels.Values.All(label => label.Before == label.After)) return false;
        foreach (var story in source.RelatedStories.Where(s => s.IsNormalStoryType))
        {
            if (story.Id is null || !labels.TryGetValue(story.Id, out var label) || story.Paragraphs.Count != 1 ||
                story.Paragraphs[0].Runs.Count == 0 || story.Paragraphs[0].Runs[0].Text != label.Before) return false;
            var p = story.Paragraphs[0];
            replacements.Add(p, p with { Runs = [p.Runs[0] with { Text = label.After }, .. p.Runs.Skip(1)] });
        }
        DocxParagraph Replace(DocxParagraph p) => replacements.GetValueOrDefault(p, p);
        IReadOnlyList<DocxBodyElement> ReplaceElements(IReadOnlyList<DocxBodyElement> elements) =>
            elements.Select(e => e is DocxParagraphElement p ? p with { Paragraph = Replace(p.Paragraph) } : e).ToArray();
        numbered = source with {
            BodyElements = ReplaceElements(source.BodyElements),
            FallbackParagraphs = source.FallbackParagraphs.Select(Replace).ToArray(),
            RelatedStories = source.RelatedStories.Select(s => s with {
                BodyElements = ReplaceElements(s.BodyElements), FallbackParagraphs = s.FallbackParagraphs.Select(Replace).ToArray()
            }).ToArray()
        };
        referencePages = pages;
        return true;
    }

    private static bool TryReplaceMarkerRuns(DocxParagraph paragraph, IReadOnlyDictionary<int, string> labels,
        IReadOnlyList<DocxInlineReference> references, out DocxParagraph replacement)
    {
        replacement = paragraph;
        var changed = new Dictionary<DocxTextRun, DocxTextRun>(ReferenceEqualityComparer.Instance);
        var runs = paragraph.Runs.Select(run =>
        {
            if (!labels.TryGetValue(run.SourceRunIndex, out string? label) || run.Text == label) return run;
            var updated = run with { Text = label };
            changed.Add(run, updated);
            return updated;
        }).ToArray();
        if (changed.Count == 0) return true;
        List<DocxInlineItem>? items = paragraph.InlineContent is null ? null : [];
        if (items is not null)
            foreach (var item in paragraph.InlineContent!)
            {
                if (item is DocxInlineText text && changed.TryGetValue(text.Run, out var updated))
                {
                    // The complete measured marker owns one numeric source run.
                    // A partial source slice retains the original numbering.
                    if (text.Start != 0 || text.Length != text.Run.Text.Length) return false;
                    items.Add(text with { Run = updated, Length = updated.Text.Length });
                }
                else items.Add(item);
            }
        replacement = paragraph with {
            Runs = runs,
            InlineReferences = references.Select(reference => reference.DisplayRun is { } run &&
                changed.TryGetValue(run, out var updated) ? reference with { DisplayRun = updated } : reference).ToArray(),
            InlineContent = items?.ToArray(),
        };
        return true;
    }

    internal static bool HasStablePages(DocxDocument numbered, DocxLayout layout,
        IReadOnlyDictionary<string, int> originalPages, bool allowWrappedBody = false, DocxWrappedCalibriPolicy calibriPolicy = DocxWrappedCalibriPolicy.NormalizedNominal,
        DocxFootnoteOwnerScope ownerScope = DocxFootnoteOwnerScope.SingleReferenceParagraphs) =>
        TryMapReferencePages(numbered, layout, out var actual, allowWrappedBody,
            calibriPolicy: calibriPolicy, ownerScope: ownerScope) && actual.Count == originalPages.Count &&
        actual.All(p => originalPages.TryGetValue(p.Key, out int before) && before == p.Value);

    internal static bool HasCompleteNoteOwners(DocxDocument document, DocxLayout layout) =>
        TryMapReferencePages(document, layout, out _);

    // Layout and font adjustments need complete source ownership even when
    // numbering itself would leave every label unchanged. Do not clone a
    // numbered document merely to answer this independent question.
    internal static bool TryGetReferencePages(DocxDocument document, DocxLayout layout,
        out IReadOnlyDictionary<string, int> pages)
    {
        pages = new Dictionary<string, int>();
        if (!IsEligibleDocument(document) || !TryMapReferencePages(document, layout, out var mapped)) return false;
        var labels = document.Paragraphs.SelectMany(p => p.InlineReferences)
            .ToDictionary(reference => reference.Id!, reference => reference.DisplayText!, StringComparer.Ordinal);
        foreach (var story in document.RelatedStories.Where(story => story.IsNormalStoryType))
            if (story.Id is null || !labels.TryGetValue(story.Id, out var label) || story.Paragraphs.Count != 1 ||
                story.Paragraphs[0].Runs.Count == 0 || story.Paragraphs[0].Runs[0].Text != label) return false;
        pages = mapped;
        return true;
    }

    internal static bool HasCompleteWrappedNoteOwners(DocxDocument document, DocxLayout layout) =>
        IsEligibleDocument(document) &&
        DocxWrappedFootnoteFrame.IsEligible(document, DocxWrappedCalibriPolicy.NormalizedThresholds) &&
        TryMapReferencePages(document, layout, out _, allowWrappedBody: true,
            calibriPolicy: DocxWrappedCalibriPolicy.NormalizedThresholds);

    internal static bool HasStableCompleteNoteOwners(DocxDocument document, DocxLayout before, DocxLayout after) =>
        before.Pages.Count == after.Pages.Count && TryMapReferencePages(document, before, out var pages) &&
        HasStablePages(document, after, pages);

    private static bool TryMapReferencePages(DocxDocument document, DocxLayout layout, out Dictionary<string, int> pages,
        bool allowWrappedBody = false, DocxWrappedCalibriPolicy calibriPolicy = DocxWrappedCalibriPolicy.NormalizedNominal,
        DocxFootnoteOwnerScope ownerScope = DocxFootnoteOwnerScope.SingleReferenceParagraphs)
    {
        pages = new(StringComparer.Ordinal);
        var linesByParagraph = new Dictionary<DocxParagraph, List<(int Page, DocxTextLineLayout Line)>>(ReferenceEqualityComparer.Instance);
        var notes = new Dictionary<string, List<(int Page, DocxPlacedRelatedStoryLayout Note)>>(StringComparer.Ordinal);
        for (int i = 0; i < layout.Pages.Count; i++)
        {
            var page = layout.Pages[i];
            if (page.ColumnFrames.Count > 1 || page.StaticTextLines.Count != 0 || page.StaticInlineImages.Count != 0 ||
                page.StaticTableRows.Count != 0 || page.StaticInlineTextBoxes.Count != 0) return false;
            foreach (var line in page.Items.OfType<DocxTextLineLayout>())
            {
                if (line.SourceParagraph is not { } p) continue;
                // A run page break can emit a spacing row owned by the preceding
                // paragraph. Keep that row in the layout, but do not count it as
                // another text line containing the paragraph's note reference.
                if (line.Segments.Count > 0 &&
                    line.Segments.All(s => s.Role == DocxTextSegmentRole.BreakSpill)) continue;
                if (!linesByParagraph.TryGetValue(p, out var lines)) linesByParagraph.Add(p, lines = []);
                lines.Add((i + 1, line));
            }
            foreach (var note in page.PlacedRelatedStories.Where(n => n.StoryLayout.Story.IsNormalStoryType))
            {
                if (note.StoryLayout.Story.Kind != DocxRelatedStoryKind.Footnote || note.StoryLayout.Story.Id is not { } id) return false;
                if (!notes.TryGetValue(id, out var places)) notes.Add(id, places = []);
                places.Add((i + 1, note));
            }
        }
        bool? wrappedBodyEligible = null;
        foreach (var p in document.Paragraphs.Where(p => p.InlineReferences.Count != 0))
        {
            if ((ownerScope != DocxFootnoteOwnerScope.AllSourceReferences && p.InlineReferences.Count != 1) ||
                !linesByParagraph.TryGetValue(p, out var lines) || lines.Count == 0) return false;
            bool wrapped = lines.Count != 1;
            // Numbering can use complete measured marker/note ownership below.
            // Font adjustment purposes keep their separately qualified frames.
            if (wrapped && (!allowWrappedBody || ownerScope != DocxFootnoteOwnerScope.AllSourceReferences &&
                (wrappedBodyEligible ??= DocxWrappedFootnoteFrame.IsEligible(document, calibriPolicy, ownerScope)) != true ||
                lines.Any(l => l.Page != lines[0].Page))) return false;
            // Index source runs and actual text segments once per paragraph.
            // Numbering can own every marker; layout/kerning purposes retain
            // their independently qualified single-reference paragraph scope.
            var markerRuns = p.Runs.GroupBy(r => r.SourceRunIndex).ToDictionary(g => g.Key, g => g.ToArray());
            var markerLines = lines.SelectMany(l => l.Line.Segments.Where(s => s.Role == DocxTextSegmentRole.Text)
                .Select(s => (l.Page, Segment: s)))
                .GroupBy(l => (l.Segment.SourceTextRunIndex, l.Segment.SourceTextOffsetInRun, l.Segment.Text))
                .ToDictionary(g => g.Key, g => g.ToArray());
            var ownedRuns = new HashSet<int>();
            foreach (var r in p.InlineReferences)
            {
                if (r.Kind != DocxRelatedStoryKind.Footnote || string.IsNullOrEmpty(r.Id) || string.IsNullOrEmpty(r.DisplayText) ||
                    r.CustomMarkText is not null || !string.IsNullOrEmpty(r.CustomMarkFollowsValue) && !OoxBoolean.IsOff(r.CustomMarkFollowsValue) ||
                    r.Revision is not null || r.Revisions.Count != 0 || !r.DisplayText.All(char.IsAsciiDigit)) return false;
                if (!markerRuns.TryGetValue(r.SourceRunIndex, out var marker) || !ownedRuns.Add(r.SourceRunIndex)) return false;
                bool completeSegmentOwner = wrapped || ownerScope == DocxFootnoteOwnerScope.AllSourceReferences;
                markerLines.TryGetValue((r.SourceRunIndex, r.TextOffsetInRun, r.DisplayText), out var ownedLines);
                if (completeSegmentOwner && ownedLines?.Length != 1) return false;
                int ownerPage = completeSegmentOwner ? ownedLines![0].Page : lines[0].Page;
                if (marker.Length != 1 || marker[0].SourceTextOffsetInRun != r.TextOffsetInRun || marker[0].Text != r.DisplayText ||
                    !notes.TryGetValue(r.Id, out var places) || places.Count != 1 || places[0].Page != ownerPage ||
                    places[0].Note.ContentTopOffset != 0d || places[0].Note.TextLines.Count != 1 ||
                    places[0].Note.StoryLayout.TextLines.Count != 1 || !pages.TryAdd(r.Id, ownerPage)) return false;
            }
        }
        return pages.Count > 0 && pages.Count == notes.Count &&
            pages.Count == document.RelatedStories.Count(s => s.IsNormalStoryType);
    }

    internal static bool TryGetWrappedCalibriPages(DocxDocument document, DocxLayout layout,
        out IReadOnlyDictionary<string, int> pages, DocxWrappedCalibriPolicy policy,
        DocxFootnoteOwnerScope ownerScope = DocxFootnoteOwnerScope.SingleReferenceParagraphs)
    {
        pages = new Dictionary<string, int>();
        string? family = document.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => r.Text.Length != 0)?.FontFamily;
        if (!string.Equals(family, "Calibri", StringComparison.OrdinalIgnoreCase) ||
            !IsEligibleDocument(document) || DocxPlainFootnoteKerning.HasSingleLineBody(layout) ||
            !DocxWrappedFootnoteFrame.IsEligible(document, policy, ownerScope) ||
            !TryMapReferencePages(document, layout, out var mapped, allowWrappedBody: true, calibriPolicy: policy, ownerScope: ownerScope)) return false;
        pages = mapped;
        return true;
    }
}
