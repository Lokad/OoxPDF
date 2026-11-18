namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineBoxReviewFramesPreserveIndependentSourceContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native review-frame content uses installed Windows fonts.");
        CheckInlineBoxMath(null, "inline-box-review-frames-office.json");
    }

    public static void DocxInlineBoxReviewStaticGlyphOriginsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent review-frame geometry uses installed Windows fonts.");
        CheckInlineBoxReference("inline-box-review-frames-office.json", InlineBoxChecks.MathOrigins);
    }
}
