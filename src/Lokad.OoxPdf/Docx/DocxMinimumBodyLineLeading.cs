using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxMinimumBodyLineLeading
{
    internal static bool TryCreateLayout(DocxDocument document, DocxLayout source, out DocxLayout adjusted, CancellationToken cancellationToken = default)
    {
        adjusted = source;
        cancellationToken.ThrowIfCancellationRequested();
        if (!DocxPlainBodyFrame.HasKnownBodyFrame(document)) return false;
        var counts = new Dictionary<DocxParagraph, int>(ReferenceEqualityComparer.Instance);
        foreach (var line in source.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (line.SourceParagraph is { } paragraph)
                counts[paragraph] = counts.GetValueOrDefault(paragraph) + 1;
        }

        var replacements = new Dictionary<DocxTextLineLayout, DocxTextLineLayout>(ReferenceEqualityComparer.Instance);
        foreach (var page in source.Pages)
        {
            // Column placement retains its independently qualified fallback.
            if (page.ColumnFrames.Count > 1) continue;
            foreach (var line in page.Items.OfType<DocxTextLineLayout>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (line.LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing ||
                    string.IsNullOrWhiteSpace(line.Text) ||
                    line.SourceParagraph is not { } paragraph || counts[paragraph] != 1 ||
                    !DocxPlainBodyFrame.IsPlainParagraph(paragraph) ||
                    !HasSupportedLineSource(paragraph) ||
                    line.LineHeight is not > 0d || !double.IsFinite(line.LineHeight.Value) ||
                    line.LineHeight > page.Height - page.MarginTop - page.MarginBottom ||
                    line.SingleLineHeight is not > 0d || !double.IsFinite(line.SingleLineHeight.Value) ||
                    line.FirstLineInsetPoints is not > 0d || !double.IsFinite(line.FirstLineInsetPoints.Value)) continue;
                // Compact marked slots retain their separately qualified link policy.
                if (paragraph.InlineReferences.Count != 0 && line.LineHeight < 16d) continue;

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
        cancellationToken.ThrowIfCancellationRequested();
        adjusted = source with {
            Pages = source.Pages.Select(p => p with {
                Items = p.Items.Select(i => i is DocxTextLineLayout line && replacements.TryGetValue(line, out var replacement)
                    ? replacement : i).ToArray()
            }).ToArray()
        };
        return true;
    }

    private static bool HasSupportedLineSource(DocxParagraph paragraph)
    {
        if (!IsQualifiedExpansionSpacing(paragraph, allowContextual: true) || paragraph.InlineReferences.Count > 1 ||
            paragraph.Spacing.ContextualSpacing == true && !(paragraph.StyleId is not null
                ? paragraph.StyleResolution.StyleFound : !string.IsNullOrWhiteSpace(paragraph.ContextualSpacingDefaultStyleId))) return false;
        var reference = paragraph.InlineReferences.FirstOrDefault();
        bool foundMarker = false;
        foreach (var run in paragraph.Runs)
        {
            if (run.Hidden) return false;
            if (reference?.SourceRunIndex == run.SourceRunIndex)
            {
                if (!IsQualifiedOwnedRun(run, reference, run.FontSize)) return false;
                foundMarker = true;
            }
            else if (run.VerticalAlignment != DocxRunVerticalAlignment.Baseline) return false;
        }
        return reference is null || foundMarker;
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
