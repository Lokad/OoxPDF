namespace Lokad.OoxPdf.Docx;

internal static class DocxPlainFootnoteKerning
{
    internal static bool TryCreateDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument adjusted, out IReadOnlyDictionary<string, int> pages)
    {
        adjusted = source;
        pages = new Dictionary<string, int>();
        bool ordinary = source.RelatedStories.Count == 0 && source.Paragraphs.All(p => p.InlineReferences.Count == 0);
        if (!(ordinary ? DocxPlainBodyFrame.IsEligible(source) : DocxPageFootnoteNumbering.IsEligibleDocument(source)) ||
            !HasSingleLineBody(layout)) return false;
        if (ordinary && (source.Paragraphs.Any(p => p.KeepRules.KeepNext == true || p.KeepRules.KeepLines == true) ||
            layout.Pages.Any(p => p.ColumnFrames.Count > 1 || p.StaticTextLines.Count != 0 || p.StaticInlineImages.Count != 0))) return false;
        bool elevenPointBody = DocxPlainElevenPointBody.IsEligible(source, layout);
        var elevenPointParagraphs = elevenPointBody
            ? new HashSet<DocxParagraph>(source.Paragraphs, ReferenceEqualityComparer.Instance) : null;
        var baselineRuns = source.Paragraphs.SelectMany(p => p.Runs)
            .Where(r => r.Text.Length != 0 && r.VerticalAlignment == DocxRunVerticalAlignment.Baseline).ToArray();
        bool nominalElevenPointFrame = elevenPointBody && baselineRuns.All(r =>
            DocxKerningThreshold.TryGetMinimum(r, out ushort minimum) && !DocxKerningThreshold.IsMet(minimum, 11d));
        // Independent pair controls admit ordinary Times and complete Calibri/
        // Arial note flow. Other met or mixed-threshold frames keep fallback.
        bool metElevenPointFrame = elevenPointBody && baselineRuns.Length > 0 &&
            (ordinary ? string.Equals(baselineRuns[0].FontFamily, "Times New Roman", StringComparison.OrdinalIgnoreCase) :
                new[] { "Calibri", "Arial" }.Contains(baselineRuns[0].FontFamily, StringComparer.OrdinalIgnoreCase)) &&
            DocxKerningThreshold.TryGetMinimum(baselineRuns[0], out ushort firstMinimum) &&
            DocxKerningThreshold.IsMet(firstMinimum, 11d) && baselineRuns.All(r =>
                DocxKerningThreshold.TryGetMinimum(r, out ushort minimum) && minimum == firstMinimum);
        var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (DocxParagraph p in source.Paragraphs.Concat(source.RelatedStories.SelectMany(s => s.Paragraphs)))
        {
            var runs = p.Runs.ToArray();
            bool changed = false;
            for (int i = 0; i < runs.Length; i++)
            {
                var run = runs[i];
                if (!DocxKerningThreshold.TryGetMinimum(run, out ushort minimum)) return false;
                // Independent Word controls keep automatic superscript note
                // numbers unkerned, including positive marker thresholds.
                if ((nominalElevenPointFrame || metElevenPointFrame) && run.FontSize == 11d && elevenPointParagraphs?.Contains(p) == true &&
                    run.VerticalAlignment == DocxRunVerticalAlignment.Superscript &&
                    p.InlineReferences.Any(r => r.SourceRunIndex == run.SourceRunIndex &&
                        r.TextOffsetInRun == run.SourceTextOffsetInRun && r.DisplayText == run.Text))
                {
                    if (run.PairKerningEnabled)
                    {
                        runs[i] = run with { PairKerningEnabled = false };
                        changed = true;
                    }
                    continue;
                }
                bool nominalBody = run.FontSize == 11d && elevenPointParagraphs?.Contains(p) == true &&
                    run.VerticalAlignment == DocxRunVerticalAlignment.Baseline;
                if (!run.Text.Any(char.IsLetter) && !nominalBody) continue;
                // Independently qualified 11pt body frames retain 12pt note
                // stories. Other sizes keep their existing fallback policy.
                if (run.FontSize != 12d && !(run.FontSize == 11d && elevenPointParagraphs?.Contains(p) == true) ||
                    run.Bold || run.Italic || run.Underline || run.Strike || run.DoubleStrike ||
                    run.AllCaps || run.SmallCaps || run.CharacterSpacingPoints != 0d ||
                    run.VerticalAlignment != DocxRunVerticalAlignment.Baseline ||
                    ordinary && (run.Hidden || run.HighlightValue is not null || run.ShadingFillHex is not null) ||
                    !new[] { "Calibri", "Arial", "Times New Roman" }.Contains(run.FontFamily, StringComparer.OrdinalIgnoreCase)) return false;
                // Measured Word output treats absent/zero as disabled and a
                // positive threshold as enabled only when the nominal size meets it.
                bool enabled = DocxKerningThreshold.IsMet(minimum, run.FontSize);
                bool logicalAdvances = nominalBody && (!enabled || metElevenPointFrame);
                if (enabled == run.PairKerningEnabled && logicalAdvances == run.UseLogicalSizeForPdfAdvances) continue;
                runs[i] = run with { PairKerningEnabled = enabled, UseLogicalSizeForPdfAdvances = logicalAdvances };
                changed = true;
            }
            if (changed) replacements.Add(p, p with { Runs = runs });
        }
        if (replacements.Count == 0 ||
            !ordinary && !DocxPageFootnoteNumbering.TryGetReferencePages(source, layout, out pages)) return false;
        DocxParagraph Replace(DocxParagraph p) => replacements.GetValueOrDefault(p, p);
        IReadOnlyList<DocxBodyElement> ReplaceElements(IReadOnlyList<DocxBodyElement> elements) =>
            elements.Select(e => e is DocxParagraphElement p ? p with { Paragraph = Replace(p.Paragraph) } : e).ToArray();
        adjusted = source with {
            BodyElements = ReplaceElements(source.BodyElements),
            FallbackParagraphs = source.FallbackParagraphs.Select(Replace).ToArray(),
            RelatedStories = source.RelatedStories.Select(s => s with {
                BodyElements = ReplaceElements(s.BodyElements), FallbackParagraphs = s.FallbackParagraphs.Select(Replace).ToArray()
            }).ToArray()
        };
        return true;
    }

    internal static bool HasSingleLineBody(DocxLayout layout) =>
        layout.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>()
            .Where(l => l.SourceParagraph is not null && l.Segments.Any(s => s.Role == DocxTextSegmentRole.Text))
            .GroupBy(l => l.SourceParagraph!, ReferenceEqualityComparer.Instance).All(g => g.Count() == 1);
}
