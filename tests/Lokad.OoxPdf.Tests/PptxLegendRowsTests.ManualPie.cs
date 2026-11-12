using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxPieFactorOffsetsKeepTheIndependentOfficeHorizontalReference()
    {
        // Independent cached Office PDFs vary factor offset, width, target and
        // edge mode. X is qualified here; their Y/radius failures stay separate.
        foreach (var (id, hash, expected) in new[]
        {
            ("manual-x0", "B402457E45C4FA06757589DCF834FF3286D66A6135C888B4573D7A6C99A118F0", 360d),
            ("manual-x02", "57E22F9E7E7A222CDDA7B551D94EBF2CF7B89819502F4D8EBE4B46893B1E96C3", 475.2d),
            ("manual-width04", "CE30CF2BA733384C88BD8B07E06A3002D12E6D8AC858E2D4962B45C9EDF1868E", 417.6d),
            ("manual-width08", "C8088FC83A7D398651D0E5D9658515C670F1B22C3E62B955AAB997E931494917", 417.6d),
            ("manual-edge", "814D9D9D23F66FBD5A4C7232C2F400E0A0BB926C1ABA1A69039D0349370D8756", 302.4d),
            ("manual-outer", "276F43081DDD6074EBF9E21E69959F74E743B787D17610BB25D7376FF5623D11", 417.6d)
        })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "PolarPlots", id + ".pptx"));
            TestAssert.Equal(hash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            using var package = OoxPackage.Open(path, default);
            var document = new PptxReader().Read(package, default);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            TestAssert.True(Math.Abs(ReadManualPieCenterX(page, 6) - expected) <= 1d,
                $"{id} must preserve Office's independent horizontal center; changing width cannot reinterpret a factor offset.");
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
        foreach (double viewportWidth in new[] { .4d, .8d })
        foreach (int count in new[] { 2, 5, 9 })
        {
            string layout = $"<c:layout><c:manualLayout><c:xMode val=\"{mode}\"/><c:yMode val=\"factor\"/>" +
                "<c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".1\"/><c:y val=\".1\"/>" +
                $"<c:w val=\"{viewportWidth.ToString(CultureInfo.InvariantCulture)}\"/><c:h val=\".7\"/></c:manualLayout></c:layout>";
            var page = Render(frameWidth, 12, count, overlay: false, kind: "pie", categoryLegend: true,
                plotLayout: layout, legendPosition: position).Page;
            double expected = 72d + .1d * frameWidth + frameWidth * (mode == "factor" ? .5d : viewportWidth / 2d);
            TestAssert.True(Math.Abs(ReadManualPieCenterX(page, count) - expected) < .002d,
                "Factor offsets own the automatic reference; edge mode owns the source rectangle, independently of legend text and point count.");
        }
    }

    private static double ReadManualPieCenterX(PdfPage page, int count)
    {
        var centers = Regex.Matches(page.Content, @"(?m)^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) l\r?\nh\r?\nf\*");
        TestAssert.Equal(count, centers.Count);
        double center = double.Parse(centers[0].Groups[1].Value, CultureInfo.InvariantCulture);
        TestAssert.True(centers.All(match => Math.Abs(double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) - center) < .001d),
            "Every sector closes through the same horizontal center.");
        return center;
    }
}
