using System.Text;
using Lokad.OoxPdf.Docx;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxPlainBoxChildInsetScalesWithItsOwningFrame()
    {
        var source = ReadOwnedBoxParagraph("inline-box-boxed-text-valid");
        var natural = Measure(1d);
        var printed = Measure(.75d);
        TestAssert.Equal(1, natural.TextLines.Count);
        TestAssert.Equal(1, printed.TextLines.Count);
        TestAssert.True(Math.Abs(natural.TextLines[0].BaselineY * .75d - printed.TextLines[0].BaselineY) < 1e-9,
            "The ordinary child baseline and its inset scale together with the complete owning frame.");
        TestAssert.True(printed.TextLines[0].InlineFrame is not null, "The shared line retains its physical geometry contract.");
        TestAssert.Equal(60d, natural.BoxHeight);
        TestAssert.Equal(45d, printed.BoxHeight);

        DocxInlineTextBoxLayout Measure(double scale)
        {
            TestAssert.True(DocxLayoutEngine.TryMeasureOrderedInlineAtoms(source, null, new SiblingFrameMeasurer(),
                11d, 10d * scale, 16d * scale, scale, default, out var atoms), "The ordinary parsed box measures.");
            return TestAssert.NotNull(TestAssert.NotNull(atoms).Fragments.Single(fragment => fragment.TextBox is not null).TextBox);
        }
    }

    public static void DocxPlainBoxChildFramesPreserveIndependentContentAndGeometry()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent child-frame geometry uses installed Windows fonts.");
        CheckInlineBoxMath(null, "inline-box-plain-child-frames-office.json");
        CheckInlineBoxWordOrigins("inline-box-plain-child-frames-office.json");
        CheckInlineBoxFollowingOrigins("inline-box-plain-child-frames-office.json");
    }

    public static void DocxPlainBoxChildTextOriginsMatchIndependentWord()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent ordinary child origins use installed Windows fonts.");
        using var gold = ReadInlineBoxMathGold("inline-box-plain-child-frames-office.json");
        foreach (var row in gold.RootElement.GetProperty("Rows").EnumerateArray())
        {
            if (!row.TryGetProperty("WordProseAnchors", out var expected)) continue;
            string id = row.GetProperty("Id").GetString()!;
            bool final = row.GetProperty("Mode").GetString() == "final-preserve";
            var mode = final ? OoxPdfDocxMarkupMode.Final : OoxPdfDocxMarkupMode.AllMarkup;
            var source = ReadMutation("InlineBoxMath/" + id, mode, false, false,
                partOverride: row.GetProperty("MathPart").GetString());
            var pages = new DocxRenderer(null, mode, final ? OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout :
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(source, null, default).ToArray();
            var lines = pages.SelectMany((page, index) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page,
                usePageCoordinates: true).GroupBy(show => show.Y).Select(line =>
                    (Page: index + 1, Height: page.Height, Y: line.Key, X: line.Min(show => show.X),
                    Text: string.Concat(line.OrderBy(show => show.X).Select(show => show.Text)).Normalize(NormalizationForm.FormKC).Trim()))).ToArray();
            foreach (var anchor in expected.EnumerateArray())
            {
                var actual = lines.Single(line => line.Text.StartsWith(anchor.GetProperty("Text").GetString()!, StringComparison.Ordinal));
                TestAssert.Equal(anchor.GetProperty("Page").GetInt32(), actual.Page);
                foreach (var axis in row.GetProperty("ProseGeometryAxes").EnumerateArray().Select(axis => axis.GetString()))
                {
                    double position = axis == "X" ? actual.X : actual.Height - actual.Y;
                    TestAssert.True(Math.Abs(position - anchor.GetProperty(axis!).GetDouble()) < 1d,
                        id + "/" + row.GetProperty("Mode").GetString() + ": ordinary child " + axis + " retains the one-point independent Word gate.");
                }
            }
        }
    }
}
