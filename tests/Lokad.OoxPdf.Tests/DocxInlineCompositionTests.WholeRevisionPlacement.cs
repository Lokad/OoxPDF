namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionWholeRevisionsKeepCellAndStaticContentAndInk()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent revised story controls use installed Windows fonts.");
        // The operator-cell Review case retains its independently recorded 1.163pt
        // X failure. This observer qualifies content, source ink and transformations.
        CheckEquationRevisions(null, false, "whole-math-placement-office.json", splitMathRuns: true);
    }
}
