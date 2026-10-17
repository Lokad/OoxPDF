using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionChildMetricControlsKeepIndependentWordContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent child metric owners use installed Windows fonts.");
        CheckEquationRevisions(null, false, "child-metric-controls-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionChildMetricControlsKeepActualOrderedOwners()
    {
        foreach (string role in new[] { "delimiter-script", "function-script", "root-script", "operator-script",
            "matrix-script", "script-fraction", "delimiter-fraction", "function-fraction", "root-fraction", "matrix-fraction" })
        {
            var document = ReadMutation("CompositionProbes/child-metric-" + role, OoxPdfDocxMarkupMode.Final, false, false);
            var equation = document.Paragraphs.First().Equations.Single();
            bool supported = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            var controls = DocxMathLayout.GetControlSources(equation, default);
            var metrics = DocxMathLayout.GetMetricOnlyControls(equation, runs, default, controls);
            TestAssert.True(supported && controls.Count == 2 &&
                ReferenceEquals(controls[0], equation.Content.Children.Single()) &&
                metrics.Count == 1 && ReferenceEquals(metrics[0], controls[1]) &&
                (role.EndsWith("-script", StringComparison.Ordinal) ? controls[1] is DocxMathScripts : controls[1] is DocxMathFraction),
                "Every admitted parent/child pair retains ordered actual controls and an independent unpainted metric owner.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.GetControlSources(equation, canceled.Token));
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.GetMetricOnlyControls(equation, runs, canceled.Token, controls));
        }
    }
}
