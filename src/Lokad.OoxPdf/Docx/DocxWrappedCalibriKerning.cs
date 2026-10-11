using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxWrappedCalibriKerning
{
    internal static bool TryCreateDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument adjusted, out IReadOnlyDictionary<string, int> pages)
    {
        adjusted = source;
        if (!DocxPageFootnoteNumbering.TryGetNominalWrappedCalibriPages(source, layout, out pages)) return false;
        var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var paragraph in source.Paragraphs.Concat(source.RelatedStories.SelectMany(s => s.Paragraphs)))
        {
            var runs = paragraph.Runs.ToArray();bool changed = false;
            for (int i = 0; i < runs.Length; i++)
            {
                var run = runs[i];ushort minimum = 0;
                if (run.KerningMinimumHalfPointsValue is { } token &&
                    (!ushort.TryParse(token, NumberStyles.None, CultureInfo.InvariantCulture, out minimum) || minimum > 32767)) return false;
                if (run.Text.Length == 0 || run.FontSize != 12d || run.VerticalAlignment != DocxRunVerticalAlignment.Baseline) continue;
                bool enabled = minimum is > 0 and <= 24;
                if (run.PairKerningEnabled == enabled) continue;
                runs[i] = run with { PairKerningEnabled = enabled };changed = true;
            }
            if (changed) replacements.Add(paragraph, paragraph with { Runs = runs });
        }
        if (replacements.Count == 0) return false;
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
}
