using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxPlainFootnoteKerning
{
    internal static bool TryCreateDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument adjusted, out IReadOnlyDictionary<string, int> pages)
    {
        adjusted = source;
        pages = new Dictionary<string, int>();
        if (!DocxPageFootnoteNumbering.IsEligibleDocument(source) || !HasSingleLineBody(layout)) return false;
        var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (DocxParagraph p in source.Paragraphs.Concat(source.RelatedStories.SelectMany(s => s.Paragraphs)))
        {
            var runs = p.Runs.ToArray();
            bool changed = false;
            for (int i = 0; i < runs.Length; i++)
            {
                var run = runs[i];
                ushort minimum = 0;
                if (run.KerningMinimumHalfPointsValue is { } value &&
                    (!ushort.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out minimum) || minimum > 32767)) return false;
                if (!run.Text.Any(char.IsLetter)) continue;
                // Word controls qualify nominal advances at 12pt in these regular
                // faces. Other sizes retain independent export-grid residuals.
                if (run.FontSize != 12d || run.Bold || run.Italic || run.Underline || run.Strike || run.DoubleStrike ||
                    run.AllCaps || run.SmallCaps || run.CharacterSpacingPoints != 0d ||
                    run.VerticalAlignment != DocxRunVerticalAlignment.Baseline ||
                    !new[] { "Calibri", "Arial", "Times New Roman" }.Contains(run.FontFamily, StringComparer.OrdinalIgnoreCase)) return false;
                // Measured Word output treats absent/zero as disabled and a
                // positive threshold as enabled only when the nominal size meets it.
                bool enabled = minimum > 0 && run.FontSize * 2d >= minimum;
                if (enabled == run.PairKerningEnabled) continue;
                runs[i] = run with { PairKerningEnabled = enabled };
                changed = true;
            }
            if (changed) replacements.Add(p, p with { Runs = runs });
        }
        if (replacements.Count == 0 ||
            !DocxPageFootnoteNumbering.TryCreateDocument(source, layout, out _, out pages)) return false;
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
