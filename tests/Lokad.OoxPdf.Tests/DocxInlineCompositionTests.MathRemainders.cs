namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMathRemaindersMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent remaining-argument geometry uses installed Windows fonts.");
        CheckEquationRevisions(null, true, "math-remaining-office.json", splitMathRuns: true);
    }
}
