using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxWrappedSeriesAndCategoryFillLegendsShareMeasuredCells()
    {
        string[] expected = Enumerable.Range(1, 6).Select(i => $"North West Region {i}").ToArray();
        foreach (var (width, size) in new[] { (288, 8), (288, 12), (576, 12), (576, 18) })
        {
            var series = Render(width, size, expected.Length, overlay: true, kind: "column").Page;
            var categories = Render(width, size, expected.Length, overlay: true, kind: "column",
                categoryLegend: true, categoryNames: expected).Page;
            var seriesNames = DocxCommentsTests.ReadEmbeddedGlyphTextShows(series, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
            var categoryNames = DocxCommentsTests.ReadEmbeddedGlyphTextShows(categories, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
            TestAssert.True(seriesNames.Select(show => show.Text).SequenceEqual(expected) &&
                    categoryNames.Select(show => show.Text).SequenceEqual(expected),
                "Changing legend identity must preserve every supported visible name and source order.");
            TestAssert.True(seriesNames.Select(show => (show.X, show.Y))
                    .SequenceEqual(categoryNames.Select(show => (show.X, show.Y))),
                "The same filled keys and names share wrapped geometry regardless of series/category identity.");
            TestAssert.True(seriesNames.Select(show => show.Y).Distinct().Count() > 1,
                "The counterpart exercises wrapped cells rather than single-row spacing.");
        }
    }

    public static void PptxWrappedFillLegendsShareColumnsAndPreserveSourceOrder()
    {
        foreach (int width in new[] { 576, 590 })
        foreach (bool overlay in new[] { false, true })
        {
            PdfPage page = Render(width, 12, 19, overlay, kind: "column", categoryLegend: true).Page;
            var names = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(19, names.Length);
            for (int i = 0; i < names.Length; i++)
                TestAssert.Equal($"Category {i + 1}", names[i].Text);
            var rows = names.GroupBy(name => name.Y).Select(row => row.ToArray()).ToArray();
            TestAssert.True(rows.Length > 1 && rows[^1].Length < rows[0].Length,
                "The fixture exercises a wrapped legend with an incomplete final row.");
            foreach (var row in rows.Skip(1))
            for (int column = 0; column < row.Length; column++)
                TestAssert.True(Math.Abs(row[column].X - rows[0][column].X) < .002,
                    "Wrapped fill keys share columns, including their incomplete final row.");
        }
    }

    public static void PptxNaturallyFittingFillLegendKeepsUnequalNamesOnOneRow()
    {
        // Portable counterpart of the registered Office transition: a natural
        // row fits although the old preferred inter-entry gaps would wrap it.
        string[] expected = Enumerable.Repeat("Category 20", 2).Concat(Enumerable.Repeat("C", 18)).ToArray();
        foreach (bool overlay in new[] { false, true })
        {
            PdfPage page = Render(576, 11, expected.Length, overlay, kind: "column",
                categoryLegend: true, categoryNames: expected).Page;
            var names = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: true,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
            TestAssert.True(names.Select(name => name.Text).SequenceEqual(expected),
                "Every source name survives in order, including repeated short names.");
            TestAssert.Equal(1, names.Select(name => name.Y).Distinct().Count());
            double previousRight = 72d;
            foreach (var show in names)
            {
                var font = show.Font.Font;
                double advance = show.Text.Sum(c => font.GetAdvanceWidth(font.MapCodePoint(c))) * show.Size / font.UnitsPerEm;
                TestAssert.True(show.X >= previousRight && show.X + advance <= 648.01d,
                    "A fitting row preserves complete, nonoverlapping text inside the chart.");
                previousRight = show.X + advance;
            }
        }
    }
}
