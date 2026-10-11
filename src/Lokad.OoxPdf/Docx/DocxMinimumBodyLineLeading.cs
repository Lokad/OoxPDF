using System.Globalization;

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
            if (line.SourceParagraph is not { } paragraph) return false;
            bool marked = paragraph.InlineReferences.Count != 0;
            bool extended = marked || paragraph.SpacingAfterPoints != 0d;
            if (paragraph.SpacingBeforePoints != 0d ||
                extended && (ordinaryBody || line.LineHeight < 16d || !IsQualifiedExpansionSpacing(paragraph)) ||
                line.LineHeight is not (>= 14d and <= 24d) ||
                line.SingleLineHeight is not > 0d || !double.IsFinite(line.SingleLineHeight.Value) ||
                line.FirstLineInsetPoints is not > 0d ||
                paragraph.Runs.Any(r => !IsQualifiedParagraphRun(paragraph, r) ||
                    !string.Equals(r.FontFamily, line.StyleRun.FontFamily, StringComparison.OrdinalIgnoreCase)) ||
                string.Equals(line.StyleRun.FontFamily, "Times New Roman", StringComparison.OrdinalIgnoreCase) &&
                    line.LineHeight < 16d) return false;

            // Independent Word font/slot probes place positive minimum-height
            // excess above the baseline. Exact/auto pitches keep their own rule.
            double excess = Math.Max(0d, line.LineHeight.Value - line.SingleLineHeight.Value);
            if (excess == 0d) continue;
            replacements.Add(line, line with {
                BaselineY = line.BaselineY - excess,
                FirstLineInsetPoints = line.FirstLineInsetPoints + excess,
                // A generated marker still owns the authored paragraph slot;
                // moving its text baseline must preserve the link's slot top.
                BodyLineBoxBaselineInsetPoints = marked ? line.FirstLineInsetPoints + excess : line.BodyLineBoxBaselineInsetPoints,
                BodyLineBoxHeightPoints = marked ? line.LineHeight : line.BodyLineBoxHeightPoints
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

    private static bool IsQualifiedParagraphRun(DocxParagraph paragraph, DocxTextRun run)
    {
        if (paragraph.InlineReferences.Count > 1) return false;
        var reference = paragraph.InlineReferences.FirstOrDefault();
        bool marker = reference is not null && run.SourceRunIndex == reference.SourceRunIndex;
        if (marker && (run.SourceTextOffsetInRun != reference!.TextOffsetInRun ||
            run.Text != reference.DisplayText || run.Text.Length == 0 || !run.Text.All(char.IsAsciiDigit))) return false;
        return IsQualifiedRun(run, marker);
    }

    private static bool IsQualifiedExpansionSpacing(DocxParagraph paragraph) =>
        paragraph.SpacingAfterPoints is >= 0d and <= 8d && paragraph.Spacing.ContextualSpacing != true &&
        paragraph.Spacing.BeforeLinesValue is null && paragraph.Spacing.AfterLinesValue is null &&
        paragraph.Spacing.BeforeAutoSpacingValue is null && paragraph.Spacing.AfterAutoSpacingValue is null &&
        IsTwipsToken(paragraph.Spacing.BeforeValue, 0) && IsTwipsToken(paragraph.Spacing.AfterValue, 160);

    private static bool IsTwipsToken(string? value, int maximum) => value is null ||
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int twips) && twips >= 0 && twips <= maximum;

    private static bool IsQualifiedRun(DocxTextRun run, bool marker) =>
        run.FontSize == 12d && !run.Bold && !run.Italic && !run.Underline && !run.Strike && !run.DoubleStrike &&
        !run.AllCaps && !run.SmallCaps && !run.Hidden && run.CharacterSpacingPoints == 0d &&
        run.HighlightValue is null && run.ShadingFillHex is null &&
        run.VerticalAlignment == (marker ? DocxRunVerticalAlignment.Superscript : DocxRunVerticalAlignment.Baseline) &&
        new[] { "Calibri", "Arial", "Times New Roman" }.Contains(run.FontFamily, StringComparer.OrdinalIgnoreCase);
}
