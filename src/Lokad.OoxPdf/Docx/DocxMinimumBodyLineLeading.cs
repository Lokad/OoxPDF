using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxMinimumBodyLineLeading
{
    internal static bool TryCreateLayout(DocxDocument document, DocxLayout source, out DocxLayout adjusted)
    {
        adjusted = source;
        if (!DocxPlainBodyFrame.HasKnownBodyFrame(document)) return false;
        var counts = new Dictionary<DocxParagraph, int>(ReferenceEqualityComparer.Instance);
        foreach (var line in source.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>())
            if (line.SourceParagraph is { } paragraph)
                counts[paragraph] = counts.GetValueOrDefault(paragraph) + 1;

        var replacements = new Dictionary<DocxTextLineLayout, DocxTextLineLayout>(ReferenceEqualityComparer.Instance);
        foreach (var page in source.Pages)
        {
            // Column placement retains its independently qualified fallback.
            if (page.ColumnFrames.Count > 1) continue;
            foreach (var line in page.Items.OfType<DocxTextLineLayout>())
            {
                if (line.LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing ||
                    string.IsNullOrWhiteSpace(line.Text) ||
                    line.SourceParagraph is not { } paragraph || counts[paragraph] != 1 ||
                    !DocxPlainBodyFrame.IsPlainParagraph(paragraph) ||
                    paragraph.Runs.Any(run => run.Hidden) || paragraph.InlineReferences.Count > 1 ||
                    line.LineHeight is not > 0d || !double.IsFinite(line.LineHeight.Value) ||
                    line.LineHeight > page.Height - page.MarginTop - page.MarginBottom ||
                    line.SingleLineHeight is not > 0d || !double.IsFinite(line.SingleLineHeight.Value) ||
                    line.FirstLineInsetPoints is not > 0d || !double.IsFinite(line.FirstLineInsetPoints.Value)) continue;

                // The measured natural line owns its bottom bearing. A larger
                // minimum slot grows above that baseline, independently of other
                // paragraphs, font size, surrounding tables or run paint.
                double excess = Math.Max(0d, line.LineHeight.Value - line.SingleLineHeight.Value);
                if (excess == 0d) continue;
                bool marked = paragraph.InlineReferences.Count != 0;
                replacements.Add(line, line with {
                    BaselineY = line.BaselineY - excess,
                    FirstLineInsetPoints = line.FirstLineInsetPoints + excess,
                    // Generated markers retain the original source slot top.
                    BodyLineBoxBaselineInsetPoints = marked ? line.FirstLineInsetPoints + excess : line.BodyLineBoxBaselineInsetPoints,
                    BodyLineBoxHeightPoints = marked ? line.LineHeight : line.BodyLineBoxHeightPoints
                });
            }
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

    internal static bool IsQualifiedParagraphRun(DocxParagraph paragraph, DocxTextRun run, double fontSize = 12d)
    {
        if (paragraph.InlineReferences.Count > 1) return false;
        var reference = paragraph.InlineReferences.FirstOrDefault();
        return IsQualifiedOwnedRun(run, reference?.SourceRunIndex == run.SourceRunIndex ? reference : null, fontSize);
    }

    internal static bool AreQualifiedParagraphRuns(DocxParagraph paragraph, string? family, DocxFootnoteOwnerScope ownerScope)
    {
        if (ownerScope != DocxFootnoteOwnerScope.AllSourceReferences)
            return paragraph.Runs.Where(r => r.Text.Length != 0).All(r => IsQualifiedParagraphRun(paragraph, r) &&
                string.Equals(r.FontFamily, family, StringComparison.OrdinalIgnoreCase));

        var references = new Dictionary<int, DocxInlineReference>();
        foreach (var reference in paragraph.InlineReferences)
            if (!references.TryAdd(reference.SourceRunIndex, reference)) return false;
        foreach (var run in paragraph.Runs.Where(r => r.Text.Length != 0))
        {
            references.TryGetValue(run.SourceRunIndex, out var reference);
            if (!IsQualifiedOwnedRun(run, reference, 12d) ||
                !string.Equals(run.FontFamily, family, StringComparison.OrdinalIgnoreCase)) return false;
        }
        return true;
    }

    private static bool IsQualifiedOwnedRun(DocxTextRun run, DocxInlineReference? reference, double fontSize)
    {
        bool marker = reference is not null;
        if (marker && (run.SourceTextOffsetInRun != reference!.TextOffsetInRun ||
            run.Text != reference.DisplayText || run.Text.Length == 0 || !run.Text.All(char.IsAsciiDigit))) return false;
        return IsQualifiedRun(run, marker, fontSize);
    }

    private static bool IsQualifiedOrdinaryExpansion(DocxDocument document, string? family, bool contextual, double fontSize) =>
        document.Paragraphs.All(p => IsQualifiedExpansionSpacing(p, contextual) && p.Runs.Where(r => r.Text.Length != 0).All(r =>
            IsQualifiedRun(r, marker: false, fontSize) && string.Equals(r.FontFamily, family, StringComparison.OrdinalIgnoreCase)));

    private static bool IsQualifiedContextualExpansion(DocxDocument document, string? family, double fontSize) =>
        document.Paragraphs.All(p => IsQualifiedExpansionSpacing(p, allowContextual: true) &&
            (p.StyleId is not null ? p.StyleResolution.StyleFound : !string.IsNullOrWhiteSpace(p.ContextualSpacingDefaultStyleId)) &&
            p.Runs.Where(r => r.Text.Length != 0).All(r => IsQualifiedParagraphRun(p, r, fontSize) &&
                string.Equals(r.FontFamily, family, StringComparison.OrdinalIgnoreCase)));

    internal static bool IsQualifiedExpansionSpacing(DocxParagraph paragraph, bool allowContextual = false) =>
        paragraph.SpacingBeforePoints is >= 0d and <= 8d &&
        paragraph.SpacingAfterPoints is >= 0d and <= 8d && (allowContextual || paragraph.Spacing.ContextualSpacing != true) &&
        paragraph.Spacing.BeforeLinesValue is null && paragraph.Spacing.AfterLinesValue is null &&
        paragraph.Spacing.BeforeAutoSpacingValue is null && paragraph.Spacing.AfterAutoSpacingValue is null &&
        IsTwipsToken(paragraph.Spacing.BeforeValue, 160) && IsTwipsToken(paragraph.Spacing.AfterValue, 160);

    private static bool IsTwipsToken(string? value, int maximum) => value is null ||
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int twips) && twips >= 0 && twips <= maximum;

    private static bool IsQualifiedRun(DocxTextRun run, bool marker, double fontSize) =>
        run.FontSize == fontSize && !run.Bold && !run.Italic && !run.Underline && !run.Strike && !run.DoubleStrike &&
        !run.AllCaps && !run.SmallCaps && !run.Hidden && run.CharacterSpacingPoints == 0d &&
        run.VerticalAlignment == (marker ? DocxRunVerticalAlignment.Superscript : DocxRunVerticalAlignment.Baseline) &&
        new[] { "Calibri", "Arial", "Times New Roman" }.Contains(run.FontFamily, StringComparer.OrdinalIgnoreCase);
}
