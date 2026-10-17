namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionRootDegreeKernsMatchIndependentHorizontalWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent root degree kern geometry uses installed Windows fonts.");
        // Vertical-clearance failures remain separate, with unchanged gates.
        CheckEquationRevisions(null, true, "root-kern-office.json", splitMathRuns: true, horizontalGeometryOnly: true);
    }
}
