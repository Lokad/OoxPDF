using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class DocxInlineCompositionTests
{
    public static void DocxInlineCompositionCellPlacementKeepsIndependentWordContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Independent cell placement uses installed Windows fonts.");
        CheckEquationRevisions(null, false, "cell-structure-placement-office.json", splitMathRuns: true);
    }

    public static void DocxInlineCompositionCellPlacementLabelsPrecedeChangedContent()
    {
        if (!OperatingSystem.IsWindows()) TestAssert.Skip("Native cell labels use installed Windows fonts.");
        CheckCellPlacement(null);
    }

    internal static void CheckCellPlacement(IFontResolver? resolver)
    {
        foreach (string kind in new[] { "ins", "del" })
        foreach (string shape in new[] { "original", "tall", "following" })
        {
            string input = "CompositionProbes/" + (shape == "original" ? "table-math-revision-cell-" + kind :
                "cell-structure-placement-" + shape + "-" + kind);
            var document = ReadMutation(input, OoxPdfDocxMarkupMode.AllMarkup, false, false);
            var page = new DocxRenderer(resolver, OoxPdfDocxMarkupMode.AllMarkup,
                OoxPdfDocxMarkupGeometryMode.WordCompatibleAllMarkup).RenderBlankPages(document, null, default).Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray();
            string label = kind == "ins" ? "Inserted Cells" : "Deleted Cells";
            var title = shows.Single(show => show.Text.Contains(label, StringComparison.Ordinal));
            var summary = shows.First(show => show.Text.Contains("Summary:", StringComparison.Ordinal));
            // Word keeps the structural label above the first changed paragraph,
            // even when the cell grows or follows an unchanged row. A center
            // anchor moves it below the content in all six controls.
            TestAssert.True(title.Y > summary.Y,
                $"The structural label must precede its changed content: {input}/{title.Y}/{summary.Y}.");
            TestAssert.Contains("share the same input series.", DocxPageFootnoteNumberingTests.DecodeText(page));
            if (shape == "following") TestAssert.Contains("Retained preceding row.", DocxPageFootnoteNumberingTests.DecodeText(page));
        }
    }
}
