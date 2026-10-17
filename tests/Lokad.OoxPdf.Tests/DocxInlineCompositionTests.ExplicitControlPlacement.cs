using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionExplicitControlPlacementKeepsIndependentWordContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent control placement uses installed Windows fonts.");
        CheckEquationRevisions(null, false, "explicit-control-placement-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionExplicitControlPlacementKeepsActualPreparationOwners()
    {
        foreach (string story in new[] { "cell", "header", "footer" })
        foreach (string role in new[] { "script", "fraction" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation($"CompositionProbes/explicit-placement-{story}-{role}", mode, false, false);
            var paragraph = DocxBlockTraversal.EnumerateSourceParagraphs(document).First(paragraph =>
                paragraph.Equations.Any(equation => DocxMathLayout.GetControlSources(equation, default).Count == 2));
            var equation = paragraph.Equations.Single();
            bool supported = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            var owners = DocxMathLayout.GetControlSources(equation, default);
            TestAssert.True(supported && owners.Count == 2 && owners[1].ControlText?.FontSize == (role == "script" ? 11 : 12) &&
                owners[1].ControlFontFamily == "Cambria Math" &&
                DocxMathLayout.GetMetricOnlyControls(equation, runs, default).Count == (role == "script" ? 1 : 0) &&
                DocxMathParagraphPolicy.GetPreparationEquations(paragraph, hasPageInputs: true).Any(item => ReferenceEquals(item, equation)),
                "The same authored child control keeps its sources and preparation eligibility across containers.");
        }
    }
}
