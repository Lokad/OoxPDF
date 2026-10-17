namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionEmptyRolesMatchIndependentStoryPlacement()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent mixed-story frames use installed Windows fonts.");
        CheckEquationRevisions(null, true, "empty-role-placement-office.json", splitMathRuns: true);
    }
}
