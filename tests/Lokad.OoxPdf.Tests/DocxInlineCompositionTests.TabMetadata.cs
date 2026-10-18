namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionUnusedTabMetadataKeepsIndependentFieldOffsets()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent tab offsets require installed Windows fonts.");
        CheckTabOffsets(null, "unused-tab-math-office.json");
    }
}
