using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionMixedBoxArgumentsKeepIndependentWordContentAndOrigins()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent mixed-argument origins use installed Windows fonts.");
        CheckMixedBoxArguments(null);
    }

    internal static void CheckMixedBoxArguments(IFontResolver? resolver)
        => CheckContainerMatrix(resolver, "mixed-box-argument-office.json");
}
