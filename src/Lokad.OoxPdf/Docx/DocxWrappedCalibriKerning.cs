namespace Lokad.OoxPdf.Docx;

internal static class DocxWrappedCalibriKerning
{
    internal static bool TryCreateDocument(DocxDocument source, DocxLayout layout,
        out DocxDocument adjusted, out IReadOnlyDictionary<string, int> pages,
        DocxWrappedCalibriPolicy policy = DocxWrappedCalibriPolicy.SourceNominal,
        DocxFootnoteOwnerScope ownerScope = DocxFootnoteOwnerScope.SingleReferenceParagraphs,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        adjusted = source;
        if (!DocxPageFootnoteNumbering.TryGetWrappedCalibriPages(source, layout, out pages, policy, ownerScope)) return false;
        var replacements = new Dictionary<DocxParagraph, DocxParagraph>(ReferenceEqualityComparer.Instance);
        foreach (var paragraph in source.Paragraphs.Concat(source.RelatedStories.SelectMany(s => s.Paragraphs)))
        {
            var runs = paragraph.Runs.ToArray();
            bool changed = false;
            for (int i = 0; i < runs.Length; i++)
            {
                var run = runs[i];
                if (!DocxWrappedFootnoteFrame.TryGetMinimum(run, out ushort minimum)) return false;
                if (run.Text.Length == 0 || run.FontSize != 12d || run.VerticalAlignment != DocxRunVerticalAlignment.Baseline) continue;
                bool enabled = DocxKerningThreshold.IsMet(minimum, run.FontSize);
                if (run.PairKerningEnabled == enabled) continue;
                runs[i] = run with { PairKerningEnabled = enabled };
                changed = true;
            }
            if (changed) replacements.Add(paragraph, paragraph with { Runs = runs });
        }
        if (replacements.Count == 0) return policy == DocxWrappedCalibriPolicy.SourceThresholds;
        var publication = new DocxParagraphReplacements(replacements, cancellationToken);
        adjusted = publication.Replace(source);
        return true;
    }
}
