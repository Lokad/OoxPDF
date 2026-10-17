namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionEmptyMathBodiesMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent empty body geometry uses installed Windows fonts.");
        CheckEquationRevisions(null, true, "math-body-office.json", splitMathRuns: true);
    }
}
