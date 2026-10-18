using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMixedMatrixCellsKeepIndependentWordContentAndSourceOrder()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent mixed-cell content uses installed Windows fonts.");
        CheckMixedMatrixCells(null);
    }

    internal static void CheckMixedMatrixCells(IFontResolver? resolver)
        // The retained Word comparison still fails the one-point gate for two
        // multi-row outcomes. Content/source invariance must not relabel them.
        => CheckContainerMatrix(resolver, "mixed-matrix-cell-office.json", validateWordGeometry: false);
}
