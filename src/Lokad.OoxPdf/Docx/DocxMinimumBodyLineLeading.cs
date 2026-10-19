using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxMinimumBodyLineLeading
{
    internal static bool TryCreateLayout(DocxDocument document, DocxLayout source, out DocxLayout adjusted)
    {
        adjusted = source;
        bool ordinaryBody = DocxPlainBodyFrame.IsOrdinaryBody(document);
        if (!(ordinaryBody ? DocxPlainBodyFrame.IsEligible(document) : DocxPageFootnoteNumbering.IsEligibleDocument(document)) ||
            !DocxPlainBodyFrame.HasSingleLineLayout(document, source) ||
            (!ordinaryBody && !DocxPageFootnoteNumbering.TryGetReferencePages(document, source, out _))) return false;

        var replacements = new Dictionary<DocxTextLineLayout, DocxTextLineLayout>(ReferenceEqualityComparer.Instance);
        double bodyFontSize = DocxPlainElevenPointBody.IsEligible(document, source) ? 11d : 12d;
        bool? ordinaryExpansionQualified = null;
        bool contextualBody = document.Paragraphs.Any(p => p.Spacing.ContextualSpacing == true);
        bool? contextualExpansionQualified = null;
        foreach (var line in source.Pages.SelectMany(p => p.Items).OfType<DocxTextLineLayout>())
        {
            if (line.LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing ||
                string.IsNullOrWhiteSpace(line.Text)) continue;
            if (line.SourceParagraph is not { } paragraph) return false;
            bool marked = paragraph.InlineReferences.Count != 0;
            bool extended = marked || paragraph.SpacingBeforePoints != 0d || paragraph.SpacingAfterPoints != 0d;
            if (extended && contextualBody)
            {
                // New contextual admissions resolve every boundary from trusted
                // styles and qualify the complete visible frame once.
                contextualExpansionQualified ??= IsQualifiedContextualExpansion(document, line.StyleRun.FontFamily, bodyFontSize);
                if (contextualExpansionQualified != true) return false;
            }
            if (ordinaryBody && extended)
            {
                // New spaced ordinary admissions qualify the complete visible
                // frame once; legacy zero-spacing paragraphs keep their gates.
                ordinaryExpansionQualified ??= IsQualifiedOrdinaryExpansion(document, line.StyleRun.FontFamily, contextualBody, bodyFontSize);
                if (ordinaryExpansionQualified != true) return false;
            }
            if (extended && (line.LineHeight < 16d || !IsQualifiedExpansionSpacing(paragraph, contextualBody)) ||
                line.LineHeight is not (>= 14d and <= 24d) ||
                line.SingleLineHeight is not > 0d || !double.IsFinite(line.SingleLineHeight.Value) ||
                line.FirstLineInsetPoints is not > 0d ||
                paragraph.Runs.Any(r => !IsQualifiedParagraphRun(paragraph, r, bodyFontSize) ||
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
        run.HighlightValue is null && run.ShadingFillHex is null &&
        run.VerticalAlignment == (marker ? DocxRunVerticalAlignment.Superscript : DocxRunVerticalAlignment.Baseline) &&
        new[] { "Calibri", "Arial", "Times New Roman" }.Contains(run.FontFamily, StringComparer.OrdinalIgnoreCase);
}
