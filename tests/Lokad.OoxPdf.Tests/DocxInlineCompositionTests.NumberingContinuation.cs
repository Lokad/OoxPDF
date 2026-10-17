namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionNumberingContinuationRetainsDefinitionAndTimelineOwners() =>
        CheckNumberingCounterTimelines("numbering-continuations-office.json", checkUnsupported: false);

    public static void DocxInlineCompositionNumberingContinuationMatchesIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native numbering frames use installed Word-compatible fonts.");
        CheckNumberingCounterContent(null, "numbering-continuations-office.json");
    }
}
