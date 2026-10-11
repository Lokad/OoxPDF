namespace Lokad.OoxPdf.Docx;

internal static class DocxMinimumBodyLineLeading
{
    internal static bool TryCreateLayout(DocxDocument document, DocxLayout source, out DocxLayout adjusted)
    {
        adjusted = source;
        bool ordinaryBody = document.RelatedStories.Count == 0 &&
            document.Paragraphs.All(p => p.InlineReferences.Count == 0);
        if (!(ordinaryBody ? DocxPlainBodyFrame.IsEligible(document) : DocxPageFootnoteNumbering.IsEligibleDocument(document)) ||
            !DocxPlainFootnoteKerning.HasSingleLineBody(source) ||
            source.Pages.Any(p => p.ColumnFrames.Count > 1 || p.StaticTextLines.Count != 0 || p.StaticInlineImages.Count != 0) ||
            document.Paragraphs.Any(p => p.KeepRules.KeepNext == true || p.KeepRules.KeepLines == true) ||
            (!ordinaryBody && !DocxPageFootnoteNumbering.TryCreateDocument(document, source, out _, out _))) return false;

        var replacements = new Dictionary<DocxTextLineLayout, DocxTextLineLayout>(ReferenceEqualityComparer.Instance);
        foreach (var line in source.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>())
        {
            if (line.LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing ||
                string.IsNullOrWhiteSpace(line.Text)) continue;
            if (line.SourceParagraph is not { } paragraph || paragraph.InlineReferences.Count != 0 ||
                paragraph.SpacingBeforePoints != 0d || paragraph.SpacingAfterPoints != 0d ||
                line.LineHeight is not (>= 14d and <= 24d) ||
                line.SingleLineHeight is not > 0d || !double.IsFinite(line.SingleLineHeight.Value) ||
                line.FirstLineInsetPoints is not > 0d ||
                paragraph.Runs.Any(r => !IsQualifiedRun(r) ||
                    !string.Equals(r.FontFamily, line.StyleRun.FontFamily, StringComparison.OrdinalIgnoreCase)) ||
                string.Equals(line.StyleRun.FontFamily, "Times New Roman", StringComparison.OrdinalIgnoreCase) &&
                    line.LineHeight < 16d) return false;

            // Independent Word font/slot probes place positive minimum-height
            // excess above the baseline. Exact/auto pitches keep their own rule.
            double excess = Math.Max(0d, line.LineHeight.Value - line.SingleLineHeight.Value);
            if (excess == 0d) continue;
            replacements.Add(line, line with {
                BaselineY = line.BaselineY - excess,
                FirstLineInsetPoints = line.FirstLineInsetPoints + excess
            });
        }
        if (replacements.Count == 0) return false;
        adjusted = source with {
            Pages = source.Pages.Select(p => p with {
                Items = p.Items.Select(i => i is DocxTextLineLayout line && replacements.TryGetValue(line, out var replacement)
                    ? replacement : i).ToArray()
            }).ToArray()
        };
        return true;
    }

    private static bool IsQualifiedRun(DocxTextRun run) =>
        run.FontSize == 12d && !run.Bold && !run.Italic && !run.Underline && !run.Strike && !run.DoubleStrike &&
        !run.AllCaps && !run.SmallCaps && !run.Hidden && run.CharacterSpacingPoints == 0d &&
        run.HighlightValue is null && run.ShadingFillHex is null &&
        run.VerticalAlignment == DocxRunVerticalAlignment.Baseline &&
        new[] { "Calibri", "Arial", "Times New Roman" }.Contains(run.FontFamily, StringComparer.OrdinalIgnoreCase);
}
