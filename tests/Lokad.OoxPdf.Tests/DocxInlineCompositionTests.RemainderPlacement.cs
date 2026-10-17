namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMathRemaindersKeepMixedStoryContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native remainder ink uses installed Windows fonts.");
        // Word group geometry is retained as independent evidence. Operator-cell
        // outcomes still fail its one-point gate; do not turn that into a pass.
        CheckEquationRevisions(null, false, "remainder-placement-office.json", splitMathRuns: true);
    }
}
