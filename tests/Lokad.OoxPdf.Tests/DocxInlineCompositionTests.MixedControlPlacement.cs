using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMixedControlPlacementMatchesIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent mixed placement geometry uses installed Windows fonts.");
        CheckEquationRevisions(null, true, "mixed-control-placement-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionMixedControlPlacementKeepsActualOwners()
    {
        foreach (string story in new[] { "cell", "header", "footer" })
        foreach (string order in new[] { "script-first", "fraction-first" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            string? part = story is "header" or "footer" ? "word/" + story + "1.xml" : null;
            var document = ReadMutation("CompositionProbes/mixed-placement-" + story + "-" + order,
                mode, false, false, partOverride: part);
            var paragraph = DocxBlockTraversal.EnumerateSourceParagraphs(document).First(paragraph =>
                paragraph.Equations.Any(equation => DocxMathLayout.GetControlSources(equation, default).Count == 2));
            var equation = paragraph.Equations.Single();
            bool hasRuns = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            var metrics = DocxMathLayout.GetMetricOnlyControls(equation, runs, default);
            bool hasOwners = DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners);
            TestAssert.True(hasRuns && runs.Count == 4 && metrics.Count == 1 && metrics[0] is DocxMathFraction &&
                hasOwners && owners.Count == 1 && owners[0].Kind == DocxRevisionKind.Insertion &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default) &&
                DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph) &&
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true)
                    .Any(item => ReferenceEquals(item, equation)),
                "Container placement preserves actual metrics and neutral control revisions beside visible neighbors.");
        }
    }
}
