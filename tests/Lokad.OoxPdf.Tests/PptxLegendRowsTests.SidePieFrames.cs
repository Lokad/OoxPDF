using System.Security.Cryptography;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxPieSideFramesRetainIndependentOfficeExtentsAndPairedOffsets()
    {
        // Office's source-paired Y/radius controls are font-independent in these
        // wide frames. X uses measured fonts, so portable assertions check the
        // opposite slots and paired source offsets instead of assuming Aptos.
        var geometry = new Dictionary<string, (double X, double Y, double Radius)>();
        foreach (var (id, hash, manual) in new[]
        {
            ("left", "197D1BDB00A2508951DE8CD3DE631E1AB3D96918BC16BA383C9EB7CB04D63942", false),
            ("right", "0A6285925E3F41E2E4ADCDA173A3C381101CE9759CC6E9A97839CF3274FAA6DB", false),
            ("left-manual", "758419CDDF4802976BD7F35BF7E0423E5CD564E5D0D141F3E96287A7942CEDD6", true),
            ("right-manual", "E203AEAE0B96EC101439ED3BCB9ACF623DB22DB2370820EC7D2B788C1744A007", true)
        })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "PolarPlots", id + ".pptx"));
            TestAssert.Equal(hash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            using var package = OoxPackage.Open(path, default);
            var document = new PptxReader().Read(package, default);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            var actual = ReadManualPieSectors(page, 6)[0];
            geometry.Add(id, actual);
            TestAssert.True(Math.Abs(actual.Y - (page.Height - (manual ? 324d : 288d))) <= 1d &&
                Math.Abs(actual.Radius - (manual ? 126d : 169.0015d)) <= 1d,
                $"{id}: the source-paired Office Y and extent must survive font-dependent side allocation.");
            string names = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false).Select(show => show.Text));
            foreach (int index in Enumerable.Range(1, 6)) TestAssert.Contains("Category " + index, names);
        }
        TestAssert.True(geometry["left"].X > 360d && geometry["right"].X < 360d &&
            Math.Abs(geometry["left"].X + geometry["right"].X - 720d) < .002d,
            "Equal left/right legend content reserves opposite slots around the same frame.");
        foreach (string side in new[] { "left", "right" })
            TestAssert.True(Math.Abs(geometry[side + "-manual"].X - geometry[side].X - 57.6d) < .002d,
                "Manual factor offsets retain the measured automatic side-frame owner.");
    }

    public static void PptxPieSideSlotsFollowContentAndRetainCompleteLabels()
    {
        foreach (string side in new[] { "l", "r" })
        foreach (int width in new[] { 288, 576 })
        foreach (int size in new[] { 8, 18 })
        {
            var shortPage = Render(width, size, 3, false, kind: "pie", categoryLegend: true,
                legendPosition: side).Page;
            string[] names = Enumerable.Range(1, 3).Select(i => "North Western Region " + i).ToArray();
            var longPage = Render(width, size, 3, false, kind: "pie", categoryLegend: true,
                legendPosition: side, categoryNames: names).Page;
            var a = ReadManualPieSectors(shortPage, 3)[0];
            var b = ReadManualPieSectors(longPage, 3)[0];
            TestAssert.True(side == "l" ? b.X > a.X : b.X < a.X,
                "More side content consumes its own side rather than a count-specific plot ratio.");
            TestAssert.True(Math.Abs(a.Y - b.Y) < .002d && b.Radius <= a.Radius + .002d,
                "A horizontal side slot keeps vertical placement and fits the remaining extent.");
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(longPage, false).Select(show => show.Text));
            foreach (string name in names) TestAssert.Contains(name, text);
            const string flags = "<c:dLbls><c:showVal val=\"1\"/><c:showCatName val=\"1\"/>" +
                "<c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/><c:showLegendKey val=\"0\"/></c:dLbls>";
            var labeled = Render(width, size, 3, false, kind: "pie", categoryLegend: true,
                legendPosition: side, dataLabels: flags, seriesOptions: "<c:explosion val=\"20\"/>").Page;
            TestAssert.True(AssertPolarLabelGlyphInk(labeled, $"side/{side}/{width}/{size}") > 15,
                "Fitting the side plot preserves actual compound-label glyph ink.");
        }
    }

    public static void PptxUnfitPieSideSlotsRetainDiagnosedPlacementFallback()
    {
        foreach (string side in new[] { "l", "r" })
        {
            var result = Render(72, 18, 3, false, kind: "pie", categoryLegend: true, legendPosition: side,
                categoryNames: Enumerable.Range(1, 3).Select(i => new string('W', 40) + i).ToArray());
            TestAssert.Equal(3, ReadManualPieSectors(result.Page, 3).Length);
            TestAssert.True(result.Diagnostics.Any(item => item.Id == "PPTX_UNSUPPORTED_CHART" &&
                item.Message.Contains("inherited plot placement is retained", StringComparison.Ordinal)),
                "An unfit side slot keeps its explicit fallback rather than collapsing the plot or claiming completeness.");
        }
    }
}
