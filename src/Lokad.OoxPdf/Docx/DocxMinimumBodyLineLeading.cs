using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxMinimumBodyLineLeading
{
    internal static bool HasSupportedPrintedSource(DocxParagraph paragraph, CancellationToken cancellationToken) =>
        paragraph.LineSpacingPoints is > 0d &&
        string.Equals(paragraph.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase) &&
        DocxPlainBodyFrame.IsPlainParagraph(paragraph) &&
        // Printed advances retain their independently qualified spacing range.
        IsQualifiedExpansionSpacing(paragraph, allowContextual: true) &&
        HasSupportedLineSource(paragraph, out bool marked, cancellationToken) &&
        (!marked || paragraph.LineSpacingPoints >= 16d);

    internal static bool TryCreateLayout(DocxDocument document, DocxLayout source, out DocxLayout adjusted, CancellationToken cancellationToken = default)
    {
        adjusted = source;
        cancellationToken.ThrowIfCancellationRequested();
        if (!DocxPlainBodyFrame.HasKnownBodyFrame(document) ||
            !DocxPlainBodyFrame.HasSupportedStaticLayout(document, source)) return false;
        var replacements = new Dictionary<DocxTextLineLayout, DocxTextLineLayout>(ReferenceEqualityComparer.Instance);
        foreach (var page in source.Pages)
        {
            for (int itemIndex = 0; itemIndex < page.Items.Count; itemIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (page.Items[itemIndex] is not DocxTextLineLayout line) continue;
                if (line.LineHeightSource != DocxLineHeightSource.AtLeastLineSpacing ||
                    string.IsNullOrWhiteSpace(line.Text) ||
                    line.SourceParagraph is not { } paragraph ||
                    !DocxPlainBodyFrame.HasPlainTextContent(paragraph) ||
                    !HasExplicitPointSpacing(paragraph, allowContextual: true) ||
                    !HasSupportedLineSource(paragraph, out bool marked, cancellationToken) ||
                    line.LineHeight is not > 0d || !double.IsFinite(line.LineHeight.Value) ||
                    line.LineHeight > page.Height - page.MarginTop - page.MarginBottom ||
                    line.SingleLineHeight is not > 0d || !double.IsFinite(line.SingleLineHeight.Value) ||
                    line.IsFirstParagraphLine is not { } firstLine ||
                    firstLine && (line.FirstLineInsetPoints is not > 0d || !double.IsFinite(line.FirstLineInsetPoints.Value))) continue;
                // Compact marked slots retain their separately qualified link policy.
                if (marked && line.LineHeight < 16d) continue;

                // Each measured natural line owns its bottom bearing. A larger
                // minimum slot grows above that baseline, independently of other
                // paragraphs, font size, horizontal alignment or run paint.
                double excess = Math.Max(0d, line.LineHeight.Value - line.SingleLineHeight.Value);
                if (excess == 0d) continue;
                bool isolatedFirstPaint = firstLine && !HasJoinedShadingBefore(page.Items, itemIndex, line);
                replacements.Add(line, line with {
                    BaselineY = line.BaselineY - excess,
                    FirstLineInsetPoints = line.FirstLineInsetPoints + excess,
                    // Isolated first lines leave minimum leading unshaded;
                    // connected backgrounds keep their measured occupied frames.
                    ParagraphShadingFrame = line.ParagraphShadingFrame is { } paint
                        ? isolatedFirstPaint ? paint with { Height = Math.Max(0d, paint.Height - excess) }
                            : paint with { BaselineInset = paint.BaselineInset + excess }
                        : null,
                    // Published first-line marker boxes retain their source top;
                    // continuation marker links keep their existing fallback.
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

    private static bool HasJoinedShadingBefore(IReadOnlyList<DocxLayoutItem> items, int index, DocxTextLineLayout line)
    {
        if (index == 0 || items[index - 1] is not DocxTextLineLayout { ParagraphShadingFrame: { } beforePaint } previous ||
            line.ParagraphShadingFrame is not { } paint || !DocxLayoutEngine.ParagraphShadingEdgesAlign(previous, line)) return false;
        double previousBottom = previous.BaselineY + beforePaint.BaselineInset - beforePaint.Height;
        double top = line.BaselineY + paint.BaselineInset;
        return Math.Abs(previousBottom - top) <= .000001d;
    }

    private static bool HasSupportedLineSource(DocxParagraph paragraph, out bool marked, CancellationToken token)
    {
        marked = false;
        if (paragraph.Spacing.ContextualSpacing == true && !(paragraph.StyleId is not null
                ? paragraph.StyleResolution.StyleFound : !string.IsNullOrWhiteSpace(paragraph.ContextualSpacingDefaultStyleId))) return false;
        DocxInlineReference? reference = null;
        foreach (var candidate in paragraph.InlineReferences)
        {
            token.ThrowIfCancellationRequested();
            // Comment references carry source annotations, not generated
            // inline marker glyphs or the note marker's reserved line box.
            if (candidate.Kind == DocxRelatedStoryKind.Comment) continue;
            if (reference is not null) return false;
            reference = candidate;
        }
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
        marked = reference is not null;
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
        HasExplicitPointSpacing(paragraph, allowContextual) &&
        paragraph.SpacingBeforePoints <= 8d && paragraph.SpacingAfterPoints <= 8d &&
        IsTwipsToken(paragraph.Spacing.BeforeValue, 160) && IsTwipsToken(paragraph.Spacing.AfterValue, 160);

    private static bool HasExplicitPointSpacing(DocxParagraph paragraph, bool allowContextual) =>
        paragraph.SpacingBeforePoints >= 0d && double.IsFinite(paragraph.SpacingBeforePoints) &&
        paragraph.SpacingAfterPoints >= 0d && double.IsFinite(paragraph.SpacingAfterPoints) &&
        (allowContextual || paragraph.Spacing.ContextualSpacing != true) &&
        paragraph.Spacing.BeforeLinesValue is null && paragraph.Spacing.AfterLinesValue is null &&
        paragraph.Spacing.BeforeAutoSpacingValue is null && paragraph.Spacing.AfterAutoSpacingValue is null &&
        IsTwipsToken(paragraph.Spacing.BeforeValue, int.MaxValue) && IsTwipsToken(paragraph.Spacing.AfterValue, int.MaxValue);

    private static bool IsTwipsToken(string? value, int maximum) => value is null ||
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int twips) && twips >= 0 && twips <= maximum;

    private static bool IsQualifiedRun(DocxTextRun run, bool marker, double fontSize) =>
        run.FontSize == fontSize && !run.Bold && !run.Italic && !run.Underline && !run.Strike && !run.DoubleStrike &&
        !run.AllCaps && !run.SmallCaps && !run.Hidden && run.CharacterSpacingPoints == 0d &&
        run.VerticalAlignment == (marker ? DocxRunVerticalAlignment.Superscript : DocxRunVerticalAlignment.Baseline) &&
        new[] { "Calibri", "Arial", "Times New Roman" }.Contains(run.FontFamily, StringComparer.OrdinalIgnoreCase);
}
