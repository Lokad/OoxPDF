using System.Globalization;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Docx;

internal static class DocxPageFootnoteNumbering
{
    internal static bool IsEligibleDocument(DocxDocument document, bool includeContinuousRestart = false)
    {
        if (!DocxPlainBodyFrame.IsEligible(document) ||
            document.Settings.FootnoteReferenceSettings.NumberRestartValue is not (null or "continuous")) return false;
        var settings = document.FinalSectionBreak!.PageSettings.FootnoteReferenceSettings;
        return (settings.NumberRestartValue == "eachPage" || includeContinuousRestart &&
                settings.NumberRestartValue is "continuous" or "eachSect") &&
            settings.NumberFormatValue is null or "decimal" or "decimalZero" &&
            (settings.NumberStartValue is null || settings.NumberStart is > 0 and <= 32767) &&
            settings.PositionValue is null or "pageBottom" &&
            document.RelatedStories.All(s => s.Kind == DocxRelatedStoryKind.Footnote &&
                s.FloatingDrawings.Count == 0 && s.Tables.Count == 0 && s.Paragraphs.All(DocxPlainBodyFrame.IsPlainParagraph));
    }

    internal static bool TryCreateDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument numbered, out IReadOnlyDictionary<string, int> referencePages, bool allowWrappedBody = false)
    {
        numbered = source;
        referencePages = new Dictionary<string, int>();
        if (!IsEligibleDocument(source) || !TryMapReferencePages(source, layout, out var pages, allowWrappedBody)) return false;
        var settings = source.FinalSectionBreak!.PageSettings.FootnoteReferenceSettings;
        var counts = new Dictionary<int, int>();
        var labels = new Dictionary<string, (string Before, string After)>(StringComparer.Ordinal);
        var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var paragraph in source.Paragraphs)
        {
            if (paragraph.InlineReferences.Count == 0) continue;
            var reference = paragraph.InlineReferences.Single();
            int page = pages[reference.Id!];
            counts.TryGetValue(page, out int count);
            int value = checked((settings.NumberStart ?? 1) + count);
            counts[page] = count + 1;
            string label = value.ToString(settings.NumberFormatValue == "decimalZero" ? "D2" : "D", CultureInfo.InvariantCulture);
            // Only the generated marker in its own source run changes. Authored
            // digits and original source offsets retain their identities.
            var runs = paragraph.Runs.Select(r => r.SourceRunIndex == reference.SourceRunIndex
                ? r with { Text = label } : r).ToArray();
            replacements.Add(paragraph, paragraph with { Runs = runs,
                InlineReferences = [reference with { DisplayText = label }] });
            labels.Add(reference.Id!, (reference.DisplayText!, label));
        }
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

    internal static bool HasStablePages(DocxDocument numbered, DocxLayout layout,
        IReadOnlyDictionary<string, int> originalPages, bool allowWrappedBody = false) =>
        TryMapReferencePages(numbered, layout, out var actual, allowWrappedBody) && actual.Count == originalPages.Count &&
        actual.All(p => originalPages.TryGetValue(p.Key, out int before) && before == p.Value);

    internal static bool HasCompleteNoteOwners(DocxDocument document, DocxLayout layout) =>
        TryMapReferencePages(document, layout, out _);

    internal static bool HasStableCompleteNoteOwners(DocxDocument document, DocxLayout before, DocxLayout after) =>
        before.Pages.Count == after.Pages.Count && TryMapReferencePages(document, before, out var pages) &&
        HasStablePages(document, after, pages);

    private static bool TryMapReferencePages(DocxDocument document, DocxLayout layout, out Dictionary<string, int> pages,
        bool allowWrappedBody = false, bool allowNominalCalibri = false)
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
            if (p.InlineReferences.Count != 1 || !linesByParagraph.TryGetValue(p, out var lines) || lines.Count == 0) return false;
            bool wrapped = lines.Count != 1;
            if (wrapped && (!allowWrappedBody || (wrappedBodyEligible ??= IsQualifiedWrappedBody(document, allowNominalCalibri)) != true ||
                lines.Any(l => l.Page != lines[0].Page))) return false;
            var r = p.InlineReferences[0];
            if (r.Kind != DocxRelatedStoryKind.Footnote || string.IsNullOrEmpty(r.Id) || string.IsNullOrEmpty(r.DisplayText) ||
                r.CustomMarkText is not null || !string.IsNullOrEmpty(r.CustomMarkFollowsValue) && !OoxBoolean.IsOff(r.CustomMarkFollowsValue) ||
                r.Revision is not null || r.Revisions.Count != 0 || !r.DisplayText.All(char.IsAsciiDigit)) return false;
            var marker = p.Runs.Where(run => run.SourceRunIndex == r.SourceRunIndex).ToArray();
            var ownedLines = lines.Where(l => l.Line.Segments.Any(s => s.Role == DocxTextSegmentRole.Text &&
                s.SourceTextRunIndex == r.SourceRunIndex && s.SourceTextOffsetInRun == r.TextOffsetInRun &&
                s.Text == r.DisplayText)).ToArray();
            if (wrapped && (ownedLines.Length != 1 || ownedLines[0].Line.Segments.Count(s => s.Role == DocxTextSegmentRole.Text &&
                s.SourceTextRunIndex == r.SourceRunIndex && s.SourceTextOffsetInRun == r.TextOffsetInRun &&
                s.Text == r.DisplayText) != 1)) return false;
            int ownerPage = wrapped ? ownedLines[0].Page : lines[0].Page;
            if (marker.Length != 1 || marker[0].SourceTextOffsetInRun != r.TextOffsetInRun || marker[0].Text != r.DisplayText ||
                !notes.TryGetValue(r.Id, out var places) || places.Count != 1 || places[0].Page != ownerPage ||
                places[0].Note.ContentTopOffset != 0d || places[0].Note.TextLines.Count != 1 ||
                places[0].Note.StoryLayout.TextLines.Count != 1 || !pages.TryAdd(r.Id, ownerPage)) return false;
        }
        return pages.Count > 0 && pages.Count == notes.Count &&
            pages.Count == document.RelatedStories.Count(s => s.IsNormalStoryType);
    }

    internal static bool TryGetNominalWrappedCalibriPages(DocxDocument document, DocxLayout layout,
        out IReadOnlyDictionary<string, int> pages)
    {
        pages = new Dictionary<string, int>();
        if (!IsEligibleDocument(document) || DocxPlainFootnoteKerning.HasSingleLineBody(layout) ||
            !IsNominalWrappedCalibriBody(document) || !TryMapReferencePages(document, layout, out var mapped,
                allowWrappedBody: true, allowNominalCalibri: true)) return false;
        pages = mapped;
        return true;
    }

    internal static bool IsNominalWrappedCalibriBody(DocxDocument document) =>
        string.Equals(document.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => r.Text.Length != 0)?.FontFamily,
            "Calibri", StringComparison.OrdinalIgnoreCase) && HasNominalBodyThresholds(document) &&
        IsQualifiedWrappedBody(document, allowNominalCalibri: true);

    private static bool HasNominalBodyThresholds(DocxDocument document) => document.Paragraphs.SelectMany(p => p.Runs)
        .Where(r => r.Text.Length != 0 && r.VerticalAlignment == DocxRunVerticalAlignment.Baseline).All(r =>
            r.KerningMinimumHalfPointsValue is null ||
            ushort.TryParse(r.KerningMinimumHalfPointsValue, NumberStyles.None, CultureInfo.InvariantCulture, out ushort minimum) &&
            minimum <= 32767 && (minimum == 0 || minimum > 24));

    private static bool IsQualifiedWrappedBody(DocxDocument document, bool allowNominalCalibri = false)
    {
        string? family = document.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => r.Text.Length != 0)?.FontFamily;
        // Calibri numbering follows a separately qualified nominal-advance
        // retry. Raw source frames still retain their original fallback.
        bool nominalCalibri = string.Equals(family, "Calibri", StringComparison.OrdinalIgnoreCase) &&
            HasNominalBodyThresholds(document) && (allowNominalCalibri || document.Paragraphs.SelectMany(p => p.Runs)
                .Where(r => r.Text.Length != 0 && r.VerticalAlignment == DocxRunVerticalAlignment.Baseline).All(r => !r.PairKerningEnabled));
        return (new[] { "Arial", "Times New Roman" }.Contains(family, StringComparer.OrdinalIgnoreCase) || nominalCalibri) && document.Paragraphs.All(p =>
            p.SpacingBeforePoints == 0d && p.SpacingAfterPoints == 0d && p.LineSpacingPoints is null &&
            p.LineSpacingFactor == 1d && DocxMinimumBodyLineLeading.IsQualifiedExpansionSpacing(p) &&
            p.KeepRules.KeepNext != true && p.KeepRules.KeepLines != true &&
            p.Runs.Where(r => r.Text.Length != 0).All(r => DocxMinimumBodyLineLeading.IsQualifiedParagraphRun(p, r) &&
                string.Equals(r.FontFamily, family, StringComparison.OrdinalIgnoreCase))) &&
            document.RelatedStories.SelectMany(s => s.Paragraphs).All(p => p.Runs.Where(r => r.Text.Any(char.IsLetter)).All(r =>
                DocxMinimumBodyLineLeading.IsQualifiedParagraphRun(p, r) &&
                string.Equals(r.FontFamily, family, StringComparison.OrdinalIgnoreCase)));
    }
}
