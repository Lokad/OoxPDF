using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionDelimiterScriptsKeepIndependentWordContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent delimited scripts use installed Windows fonts.");
        CheckEquationRevisions(null, false, "delimiter-scripts-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionDelimiterScriptsKeepActualControlAndDeeperFallback()
    {
        foreach (string role in new[] { "sub", "sup", "combined", "fraction-base" })
        {
            var document = ReadMutation("CompositionProbes/delimiter-script-" + role, OoxPdfDocxMarkupMode.Final, false, false);
            var paragraph = document.Paragraphs.First();
            var source = paragraph.Equations.Single();
            var equation = source with { Display = source.Display &&
                !DocxMathParagraphPolicy.GetMixedSourceEquations(paragraph, default).Contains(source) };
            var delimiter = (DocxMathDelimiter)equation.Content.Children.Single();
            bool supported = DocxMathLayout.TryGetRuns(equation, default, out var runs);
            int count = role == "combined" ? 3 : role == "fraction-base" ? 4 : 2;
            TestAssert.True(supported && runs.Count == count &&
                runs[0].ControlOwners.Any(control => ReferenceEquals(control, delimiter)),
                "The actual delimiter owns structural paint while script roles retain their literal sources.");
            var deeper = equation with { Content = new([delimiter with { Arguments = [new([delimiter])] }]) };
            TestAssert.True(!DocxMathLayout.TryGetRuns(deeper, default, out _) &&
                !DocxMathLayout.TryMeasure(deeper, _ => null, default, out var box, out _) && box is null,
                "Nested delimiter admission remains explicit and cannot publish partial mathematical text.");
            using var canceled = new CancellationTokenSource(); canceled.Cancel();
            TestAssert.Throws<OperationCanceledException>(() => DocxMathLayout.TryGetRuns(equation, canceled.Token, out _));
        }
    }
}
