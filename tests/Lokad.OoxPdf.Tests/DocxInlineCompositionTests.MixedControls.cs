using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMixedControlsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent mixed control geometry uses installed Windows fonts.");
        CheckEquationRevisions(null, true, "mixed-controls-office.json", splitMathRuns: true,
            includeRow: row => !VisibleDeletion(row));
        // The already supported visible Review controls retain their raw width
        // residuals; retain content and ink checks without relaxing any gate.
        CheckEquationRevisions(null, false, "mixed-controls-office.json", splitMathRuns: true,
            includeRow: VisibleDeletion);

        static bool VisibleDeletion(System.Text.Json.JsonElement row) =>
            row.GetProperty("Id").GetString()!.EndsWith("-fraction-del", StringComparison.Ordinal) &&
            row.GetProperty("Mode").GetString() == "all-word-compatible";
    }

    public static void DocxInlineCompositionMixedControlsKeepIndependentMetricAndRevisionOwners()
    {
        foreach (string order in new[] { "script-first", "fraction-first" })
        foreach (string kind in new[] { "authored-empty", "fraction-ins", "fraction-del" })
        foreach (var mode in new[] { OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupMode.AllMarkup })
        {
            var document = ReadMutation("CompositionProbes/mixed-control-" + order + "-" + kind, mode, false, false);
            var paragraph = document.Paragraphs.First();
            var equation = paragraph.Equations.Single();
            bool visible = kind == "fraction-del" && mode == OoxPdfDocxMarkupMode.AllMarkup;
            bool hasRuns = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            var controls = DocxMathLayout.GetControlSources(equation, default);
            var metrics = DocxMathLayout.GetMetricOnlyControls(equation, runs, default);
            bool hasOwners = DocxInlineContent.TryGetEquationRevisionOwners(equation, default, out var owners);
            TestAssert.True(hasRuns && runs.Count == (visible ? 6 : 4) && controls.Count == 2 &&
                metrics.Count == (visible ? 0 : 1) && (visible || metrics[0] is DocxMathFraction) &&
                hasOwners && owners.Count == (visible || kind == "fraction-ins" ? 1 : 0) &&
                DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph, default) &&
                DocxInlineContent.HasOnlyNonRenderingRevisionMetadata(paragraph) != visible,
                "Painted neighbors do not hide an empty control's metrics or invent revised literal ink.");
            if (kind == "fraction-ins")
            {
                var script = controls.OfType<DocxMathScripts>().Single();
                TestAssert.True(!DocxInlineContent.HasSupportedParagraphRevisionOwners(paragraph with
                    { Equations = [equation with { Content = new([script]) }] }, default),
                    "A visible unrevised neighbor cannot own an absent revised control's metadata.");
            }
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.GetMetricOnlyControls(equation, runs, canceled.Token));
            TestAssert.Throws<OperationCanceledException>(() => DocxInlineContent.TryGetEquationRevisionOwners(equation, canceled.Token, out _));
        }
    }
}
