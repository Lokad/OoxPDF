using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxManualPiesKeepIndependentOfficeCentersAndRadii()
    {
        // Independent cached Office PDFs vary factor offset, width, target and
        // edge mode. Their independent radius/X evidence does not qualify Y.
        foreach (var (id, hash, expected, expectedRadius) in new[]
        {
            ("manual-x0", "B402457E45C4FA06757589DCF834FF3286D66A6135C888B4573D7A6C99A118F0", 360d, 126d),
            ("manual-x02", "57E22F9E7E7A222CDDA7B551D94EBF2CF7B89819502F4D8EBE4B46893B1E96C3", 475.2d, 126d),
            ("manual-width04", "CE30CF2BA733384C88BD8B07E06A3002D12E6D8AC858E2D4962B45C9EDF1868E", 417.6d, 115.2d),
            ("manual-width08", "C8088FC83A7D398651D0E5D9658515C670F1B22C3E62B955AAB997E931494917", 417.6d, 126d),
            ("manual-edge", "814D9D9D23F66FBD5A4C7232C2F400E0A0BB926C1ABA1A69039D0349370D8756", 302.4d, 126d),
            ("manual-outer", "276F43081DDD6074EBF9E21E69959F74E743B787D17610BB25D7376FF5623D11", 417.6d, 126d)
        })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "PolarPlots", id + ".pptx"));
            TestAssert.Equal(hash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            using var package = OoxPackage.Open(path, default);
            var document = new PptxReader().Read(package, default);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            var geometry = ReadManualPieSectors(page, 6)[0];
            TestAssert.True(Math.Abs(geometry.X - expected) <= 1d,
                $"{id} must preserve Office's independent horizontal center; changing width cannot reinterpret a factor offset.");
            TestAssert.True(Math.Abs(geometry.Radius - expectedRadius) <= 1d,
                $"{id}: authored viewport radius must match the independent Office circle.");
            string names = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).Select(show => show.Text));
            foreach (int index in Enumerable.Range(1, 6)) TestAssert.Contains("Category " + index, names);
        }
    }

    public static void PptxPieFactorAndEdgePlacementKeepTheirOwnersAcrossViewportWidths()
    {
        foreach (string mode in new[] { "factor", "edge" })
        foreach (string position in new[] { "t", "b" })
        foreach (int frameWidth in new[] { 420, 576 })
        foreach (int frameHeight in new[] { 180, 360 })
        foreach (double viewportWidth in new[] { .4d, .8d })
        foreach (int count in new[] { 2, 5, 9 })
        {
            string layout = $"<c:layout><c:manualLayout><c:xMode val=\"{mode}\"/><c:yMode val=\"factor\"/>" +
                "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".1\"/><c:y val=\".1\"/>" +
                $"<c:w val=\"{viewportWidth.ToString(CultureInfo.InvariantCulture)}\"/><c:h val=\".7\"/></c:manualLayout></c:layout>";
            var page = Render(frameWidth, 12, count, overlay: false, kind: "pie", categoryLegend: true,
                plotLayout: layout, legendPosition: position, height: frameHeight).Page;
            double expected = 72d + .1d * frameWidth + frameWidth * (mode == "factor" ? .5d : viewportWidth / 2d);
            TestAssert.True(Math.Abs(ReadManualPieSectors(page, count)[0].X - expected) < .002d,
                "Factor offsets own the automatic reference; edge mode owns the source rectangle, independently of legend text and point count.");
            TestAssert.True(Math.Abs(ReadManualPieSectors(page, count)[0].Radius - Math.Min(frameWidth * viewportWidth, frameHeight * .7d) / 2d) < .002d,
                "Authored extent follows both viewport dimensions; these transformations do not claim new Office qualification.");
        }
    }

    public static void PptxPieFactorOffsetsRetainBoundedSourceAndCompleteLegendContent()
    {
        foreach (string offset in new[] { "-1e306", ".1", "1e306" })
        {
            string layout = "<c:layout><c:manualLayout><c:xMode val=\"factor\"/><c:yMode val=\"factor\"/>" +
                "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/>" +
                $"<c:x val=\"{offset}\"/><c:y val=\".1\"/><c:w val=\".6\"/><c:h val=\".7\"/></c:manualLayout></c:layout>";
            var page = Render(576, 12, 6, overlay: false, kind: "pie", categoryLegend: true, plotLayout: layout).Page;
            string names = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).Select(show => show.Text));
            TestAssert.Equal(string.Concat(Enumerable.Range(1, 6).Select(index => "Category " + index)), names);
            double expected = 360d + Math.Clamp(double.Parse(offset, CultureInfo.InvariantCulture), 0d, 1d) * 576d;
            TestAssert.True(Math.Abs(ReadManualPieSectors(page, 6)[0].X - expected) < .002d,
                "Factor placement must reuse the existing resolved source bounds before applying the automatic reference.");
        }
    }

    public static void PptxManualPieViewportsRetainLabelsAndExplosionClearance()
    {
        const string layout = "<c:layout><c:manualLayout><c:xMode val=\"factor\"/><c:yMode val=\"factor\"/>" +
            "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".1\"/><c:y val=\".1\"/>" +
            "<c:w val=\".6\"/><c:h val=\".7\"/></c:manualLayout></c:layout>";
        foreach (int width in new[] { 230, 576 })
        foreach (bool labeled in new[] { false, true })
        {
            string labels = labeled ? "<c:dLbls><c:showVal val=\"1\"/></c:dLbls>" : "";
            var ordinary = Render(width, 12, 6, overlay: false, kind: "pie", categoryLegend: true,
                plotLayout: layout, dataLabels: labels).Page;
            var exploded = Render(width, 12, 6, overlay: false, kind: "pie", categoryLegend: true,
                plotLayout: layout, dataLabels: labels, seriesOptions: "<c:explosion val=\"20\"/>").Page;
            var baseline = ReadManualPieSectors(ordinary, 6);
            var displaced = ReadManualPieSectors(exploded, 6);
            for (int index = 0; index < 6; index++)
            {
                var sector = displaced[index];
                TestAssert.True(Math.Abs(sector.Radius * 1.2d - baseline[index].Radius) < .002d,
                    "Explosion reservation must scale the shared radius, including retained label clearance.");
                double offset = Math.Sqrt(Math.Pow(sector.X - baseline[index].X, 2d) + Math.Pow(sector.Y - baseline[index].Y, 2d));
                TestAssert.True(Math.Abs(offset - sector.Radius * .2d) < .002d,
                    "Every exploded sector stays within the reserved circular extent.");
            }
            foreach (PdfPage page in new[] { ordinary, exploded })
            {
                var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false).ToArray();
                string names = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                    fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).Select(show => show.Text));
                foreach (int index in Enumerable.Range(1, 6)) TestAssert.Contains("Category " + index, names);
                if (labeled) TestAssert.Equal(6, shows.Count(show => show.Text == "4"));
            }
        }
    }

    private static (double X, double Y, double Radius)[] ReadManualPieSectors(PdfPage page, int count)
    {
        const string number = @"(-?\d+(?:\.\d+)?)";
        var sectors = Regex.Matches(page.Content, $@"(?m)^{number} {number} m\r?\n(?:[^\r\n]+ c\r?\n)+{number} {number} l\r?\nh\r?\nf\*");
        TestAssert.Equal(count, sectors.Count);
        double Value(Match match, int group) => double.Parse(match.Groups[group].Value, CultureInfo.InvariantCulture);
        return sectors.Select(match =>
        {
            double x = Value(match, 3), y = Value(match, 4);
            return (x, y, Math.Sqrt(Math.Pow(Value(match, 1) - x, 2d) + Math.Pow(Value(match, 2) - y, 2d)));
        }).ToArray();
    }
}
