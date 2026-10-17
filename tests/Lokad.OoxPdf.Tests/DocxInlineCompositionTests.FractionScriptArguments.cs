using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionFractionScriptArgumentsRetainIndependentWordContentAndDisplayFrames()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent fraction/script geometry uses installed Windows fonts.");
        CheckEquationRevisions(null, true, "fraction-script-arguments-office.json", splitMathRuns: true,
            includeCase: id => id.Contains("-display-", StringComparison.Ordinal));
        // Raw inline gates retain the 1.057pt baseline and 0.108pt width
        // residuals. Content, rule count, ink and source transformations still
        // compare independently with Word in both styles; no gate is relaxed.
        CheckEquationRevisions(null, false, "fraction-script-arguments-office.json", splitMathRuns: true,
            includeCase: id => id.Contains("-inline-", StringComparison.Ordinal));
    }

    public static void DocxInlineCompositionFractionScriptArgumentsKeepRolesAndDeeperFallback()
    {
        foreach (string style in new[] { "inline", "display" })
        foreach (string role in new[] { "numerator", "denominator" })
        {
            var document = ReadMutation("CompositionProbes/fraction-script-" + style + "-" + role,
                OoxPdfDocxMarkupMode.Final, false, false);
            var paragraph = document.Paragraphs.First();
            var authored = paragraph.Equations.Single();
            // Source display markup can occur alongside prose. Preparation
            // selects the same semantic source-row style used by rendering.
            var equation = authored with { Display = authored.Display &&
                !DocxMathParagraphPolicy.GetMixedSourceEquations(paragraph, default).Contains(authored) };
            bool discovered = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            int basis = style == "display" ? 0 : 1, argument = basis + 1;
            int[] expected = role == "numerator" ? [basis, argument, argument, basis] : [basis, basis, argument, argument];
            TestAssert.True(discovered && runs.Select(run => run.ScriptLevel).SequenceEqual(expected),
                "Fraction arguments retain text/script/script-script source roles before font publication.");
            var fraction = (DocxMathFraction)equation.Content.Children.Single();
            var script = (DocxMathScripts)(role == "numerator" ? fraction.Numerator : fraction.Denominator).Children.Single();
            var nested = new DocxMathRow([script with { Base = new([fraction]) }]);
            var deeper = equation with { Content = new([role == "numerator" ? fraction with { Numerator = nested } : fraction with { Denominator = nested }]) };
            TestAssert.True(!DocxMathLayout.TryGetRuns(deeper, default, out _) &&
                !DocxMathLayout.TryMeasure(deeper, _ => null, default, out var box, out _) && box is null,
                "A fraction argument script requires literal roles; deeper content cannot publish a partial box.");
            TestAssert.True(DocxMathLayout.RemoveDeletedLiteralMathText(equation, default) is null,
                "Unverified whole-deletion retention of structured arguments keeps its explicit fallback.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, canceled.Token, out _));
        }
    }
}
