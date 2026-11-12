using System.Globalization;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxPieFactorOffsetsFollowMeasuredAutomaticFramesInBothAxes()
    {
        // Independent paired Office sources establish the factor offset from
        // automatic placement. These transformations test that ownership across
        // dimensions/fonts; they do not qualify new Office geometry.
        foreach (string position in new[] { "t", "b", "l", "r" })
        foreach (var (height, size, count, width, extent) in new[]
        {
            (180, 8, 3, .4d, .5d), (360, 18, 9, .8d, .7d)
        })
        {
            var automatic = Render(576, size, count, false, kind: "pie", height: height,
                categoryLegend: true, legendPosition: position).Page;
            var basis = ReadManualPieSectors(automatic, count)[0];
            foreach (string xMode in new[] { "factor", "edge" })
            foreach (double offset in new[] { 0d, .2d })
            {
                string layout = "<c:layout><c:manualLayout><c:xMode val=\"" + xMode + "\"/><c:yMode val=\"factor\"/>" +
                    "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".1\"/>" +
                    $"<c:y val=\"{offset.ToString(CultureInfo.InvariantCulture)}\"/><c:w val=\"{width.ToString(CultureInfo.InvariantCulture)}\"/>" +
                    $"<c:h val=\"{extent.ToString(CultureInfo.InvariantCulture)}\"/></c:manualLayout></c:layout>";
                var page = Render(576, size, count, false, kind: "pie", height: height,
                    categoryLegend: true, legendPosition: position, plotLayout: layout).Page;
                var actual = ReadManualPieSectors(page, count)[0];
                double x = basis.X + 57.6d;
                if (xMode == "edge")
                {
                    var edge = Render(576, size, count, false, kind: "pie", height: height,
                        categoryLegend: true, legendPosition: position,
                        plotLayout: layout.Replace("<c:yMode val=\"factor\"/>", "<c:yMode val=\"edge\"/>", StringComparison.Ordinal)).Page;
                    x = ReadManualPieSectors(edge, count)[0].X;
                }
                TestAssert.True(Math.Abs(actual.X - x) < .002d && Math.Abs(actual.Y - (basis.Y - offset * height)) < .002d,
                    $"Factor offsets follow the measured automatic frame: {position}/{height}/{size}/{count}/{xMode}/{offset}; " +
                    $"actual {actual.X},{actual.Y}; expected {x},{basis.Y - offset * height}; basis {basis.X},{basis.Y}.");
                TestAssert.True(Math.Abs(actual.Radius - Math.Min(576d * width, height * extent) / 2d) < .002d,
                    "Offsets do not replace the authored viewport extent.");
                string names = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                    fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).Select(show => show.Text));
                foreach (int index in Enumerable.Range(1, count)) TestAssert.Contains("Category " + index, names);
            }
        }
    }

    public static void PptxPieFactorYReusesBoundedOffsetsAndRetainsEdgeX()
    {
        var automatic = ReadManualPieSectors(Render(576, 12, 6, false, kind: "pie", categoryLegend: true).Page, 6)[0];
        foreach (var (offset, expectedNames) in new[] { ("-1e306", 6), (".1", 6), ("1e306", 0) })
        {
            string layout = "<c:layout><c:manualLayout><c:xMode val=\"edge\"/><c:yMode val=\"factor\"/>" +
                "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".1\"/>" +
                $"<c:y val=\"{offset}\"/><c:w val=\".6\"/><c:h val=\".7\"/></c:manualLayout></c:layout>";
            var result = Render(576, 12, 6, false, kind: "pie", categoryLegend: true, plotLayout: layout);
            var page = result.Page;
            var actual = ReadManualPieSectors(page, 6)[0];
            double expected = automatic.Y - Math.Clamp(double.Parse(offset, CultureInfo.InvariantCulture), 0d, 1d) * 360d;
            TestAssert.True(Math.Abs(actual.X - 302.4d) < .002d && Math.Abs(actual.Y - expected) < .002d,
                "Factor Y uses the resolved bounded source offset; it cannot overflow and drop the chart or rewrite edge X.");
            string names = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).Select(show => show.Text));
            foreach (int index in Enumerable.Range(1, expectedNames)) TestAssert.True(names.Contains("Category " + index, StringComparison.Ordinal),
                $"Bounded Y {offset} retains Category {index}; decoded legend: {names}.");
            if (expectedNames == 0)
            {
                // The existing out-of-frame authored fallback is not a content
                // completeness pass. Both parent and candidate retain its warning.
                TestAssert.Equal("", names);
                TestAssert.True(result.Diagnostics.Any(diagnostic => diagnostic.Id == "PPTX_UNSUPPORTED_CHART" &&
                    diagnostic.Message.Contains("Legend exceeds its available frame", StringComparison.Ordinal)),
                    "An unfit authored frame remains diagnosed instead of being reported complete.");
            }
        }
    }
}
