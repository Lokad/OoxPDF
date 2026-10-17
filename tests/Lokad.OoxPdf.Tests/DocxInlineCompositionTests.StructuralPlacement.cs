namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionStructuralPaintKeepsMixedStoryContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native mixed structural ink uses installed Windows fonts.");
        // Independent group geometry retains the known operator-cell Review
        // one-point failure. Content/ink checks do not claim that gate passed.
        CheckEquationRevisions(null, false, "structural-placement-office.json", splitMathRuns: true);
    }
}
