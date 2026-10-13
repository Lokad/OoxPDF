namespace Lokad.OoxPdf.Docx;

// These purposes preserve the existing source/derived policy distinction. The
// mixed extension uses the same regular-12pt frame and exact owner mapping.
internal enum DocxWrappedCalibriPolicy
{
    NormalizedNominal,
    SourceNominal,
    NormalizedUniformMet,
    SourceThresholds,
    NormalizedThresholds
}

internal enum DocxBodyKerningKind { Invalid, Nominal, UniformMet, Mixed }

internal static class DocxWrappedFootnoteFrame
{
    internal readonly record struct KerningPolicy(DocxBodyKerningKind Kind, bool PairsMatchThresholds, bool PairsDisabled);

    internal static bool TryGetMinimum(DocxTextRun run, out ushort minimum)
        => DocxKerningThreshold.TryGetMinimum(run, out minimum);

    internal static KerningPolicy ClassifyKerning(DocxDocument document)
    {
        ushort? firstMet = null;
        bool nominal = false, differingMet = false, matches = true, disabled = true;
        foreach (var run in document.Paragraphs.SelectMany(p => p.Runs)
            .Where(r => r.Text.Length != 0 && r.VerticalAlignment == DocxRunVerticalAlignment.Baseline))
        {
            if (!TryGetMinimum(run, out ushort minimum)) return new(DocxBodyKerningKind.Invalid, false, false);
            bool met = minimum is > 0 and <= 24;
            matches &= run.PairKerningEnabled == met;
            disabled &= !run.PairKerningEnabled;
            if (!met) nominal = true;
            else if (firstMet is null) firstMet = minimum;
            else differingMet |= firstMet != minimum;
        }
        return new(firstMet is null ? DocxBodyKerningKind.Nominal :
            nominal || differingMet ? DocxBodyKerningKind.Mixed : DocxBodyKerningKind.UniformMet, matches, disabled);
    }

    internal static bool IsEligible(DocxDocument document, DocxWrappedCalibriPolicy purpose,
        DocxFootnoteOwnerScope ownerScope = DocxFootnoteOwnerScope.SingleReferenceParagraphs)
    {
        string? family = document.Paragraphs.SelectMany(p => p.Runs).FirstOrDefault(r => r.Text.Length != 0)?.FontFamily;
        bool admitted = string.Equals(family, "Arial", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(family, "Times New Roman", StringComparison.OrdinalIgnoreCase);
        if (string.Equals(family, "Calibri", StringComparison.OrdinalIgnoreCase))
        {
            var policy = ClassifyKerning(document);
            admitted = purpose switch
            {
                DocxWrappedCalibriPolicy.SourceNominal => policy.Kind == DocxBodyKerningKind.Nominal,
                DocxWrappedCalibriPolicy.SourceThresholds => policy.Kind is DocxBodyKerningKind.Nominal or DocxBodyKerningKind.Mixed,
                DocxWrappedCalibriPolicy.NormalizedUniformMet =>
                    policy.Kind == DocxBodyKerningKind.Nominal && policy.PairsDisabled ||
                    policy.Kind == DocxBodyKerningKind.UniformMet && policy.PairsMatchThresholds,
                DocxWrappedCalibriPolicy.NormalizedThresholds =>
                    policy.Kind != DocxBodyKerningKind.Invalid && policy.PairsMatchThresholds,
                _ => policy.Kind == DocxBodyKerningKind.Nominal && policy.PairsDisabled
            };
        }
        return admitted && document.Paragraphs.All(p =>
            (purpose is DocxWrappedCalibriPolicy.SourceThresholds or DocxWrappedCalibriPolicy.NormalizedThresholds ||
                p.SpacingBeforePoints == 0d && p.SpacingAfterPoints == 0d) && p.LineSpacingPoints is null &&
            p.LineSpacingFactor == 1d && DocxMinimumBodyLineLeading.IsQualifiedExpansionSpacing(p) &&
            p.KeepRules.KeepNext != true && p.KeepRules.KeepLines != true &&
            DocxMinimumBodyLineLeading.AreQualifiedParagraphRuns(p, family, ownerScope)) &&
            document.RelatedStories.SelectMany(s => s.Paragraphs).All(p => p.Runs.Where(r => r.Text.Any(char.IsLetter)).All(r =>
                DocxMinimumBodyLineLeading.IsQualifiedParagraphRun(p, r) &&
                string.Equals(r.FontFamily, family, StringComparison.OrdinalIgnoreCase)));
    }
}
