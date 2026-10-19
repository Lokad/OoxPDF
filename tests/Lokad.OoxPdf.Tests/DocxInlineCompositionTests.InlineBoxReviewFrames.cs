using System.Text;
using Lokad.OoxPdf.Docx;

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
        CheckInlineBoxWordOrigins("inline-box-review-frames-office.json");
    }

    private static void CheckInlineBoxWordOrigins(string goldenFile)
    {
        using var gold = ReadInlineBoxMathGold(goldenFile);
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string id = row.GetProperty("Id").GetString()!;
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = ReadMutation("InlineBoxMath/" + id, mode, false, false,
                partOverride: row.GetProperty("MathPart").GetString());
            var pages = new DocxRenderer(null, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, default).ToArray();
            var math = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                usePageCoordinates: true).Where(show => page.Fonts.Any(font => font.ResourceName.StartsWith("Eq", StringComparison.Ordinal) &&
                    ReferenceEquals(font.Font, show.Font))).Select(show => (Page: index + 1, Height: page.Height, Show: show))).ToArray();
            var expected = row.GetProperty("WordMathAnchors").EnumerateArray().ToArray();
            TestAssert.Equal(expected.Length, math.Length);
            for (int index = 0; index < expected.Length; index++)
            {
                var actual = math[index]; var anchor = expected[index];
                TestAssert.Equal(anchor.GetProperty("Text").GetString(), actual.Show.Text.Normalize(NormalizationForm.FormKC));
                TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                foreach (var axis in row.GetProperty("GeometryAxes").EnumerateArray().Select(axis => axis.GetString()))
                {
                    double position = axis == "X" ? actual.Show.X : actual.Height - actual.Show.Y;
                    TestAssert.True(Math.Abs(position - anchor.GetProperty(axis!).GetDouble()) < 1d,
                        id + "/" + row.GetProperty("Mode").GetString() + ": child glyph " + axis + " retains the one-point independent Word gate.");
                }
            }
        }
    }
}
