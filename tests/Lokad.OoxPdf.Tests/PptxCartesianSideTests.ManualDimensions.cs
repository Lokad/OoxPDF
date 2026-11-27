using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxCartesianSideTests
{
    private static string ManualCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "ManualCaptionFrames"));

    public static void PptxManualFactorAndEdgeSizesShareIndependentOfficeRectangles()
    {
        using var controls = ReadManualControls();
        string[] passive = ["inner-factor-zero", "inner-factor-down", "inner-factor-up", "inner-factor-h45", "inner-factor-h80",
            "inner-factor-edge-height", "inner-edge-edge-height", "inner-factor-edge-height-up", "line-valid-edge-y16"];
        var rows = controls.RootElement.EnumerateArray().ToDictionary(row => row.GetProperty("Id").GetString()!, StringComparer.Ordinal);
        foreach (string id in passive)
        {
            TestAssert.True(rows.TryGetValue(id, out var row), $"The manual chart reference is missing '{id}'.");
            PdfPage page = RenderManual(row);
            CheckManualRectangle(page, row);
            if (id == "inner-factor-edge-height")
            {
                // The outline paths remain in the content stream, but the
                // separate axis strokes are discarded. Grid strokes remain.
                var discardedAxes = page with
                {
                    Content = page.Content.ReplaceLineEndings("\n").Replace(" l S\n", " l n\n", StringComparison.Ordinal)
                };
                var missing = TestAssert.Throws<InvalidOperationException>(() => CheckManualRectangle(discardedAxes, row));
                TestAssert.Contains("Both painted plot axes", missing.Message);
            }
        }
    }

    public static void PptxManualScatterEdgeHeightKeepsIndependentOfficeRectangle()
    {
        using var controls = ReadManualControls();
        var row = controls.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "scatter-valid-edge-height");
        var actual = ManualPlotBounds(RenderManual(row));
        var native = row.GetProperty("NativePlot");
        double bottom = native.GetProperty("Bottom").GetDouble(), height = native.GetProperty("Height").GetDouble();
        TestAssert.True(Math.Abs(actual.Bottom - bottom) <= .1d && Math.Abs(actual.Top - bottom - height) <= .1d,
            "Scatter edge heights retain the independent Office vertical rectangle.");
        // Scatter already measures its horizontal numeric captions. Keep that
        // font-dependent owner, and compare equivalent source representations
        // with the same pinned face instead of asserting a Calibri X origin.
        using var stream = new MemoryStream(); stream.Write(ReadManualBytes(row)); stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
            XDocument xml; using (var input = entry.Open()) xml = XDocument.Load(input);
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var manual = xml.Descendants(c + "plotArea").Single().Element(c + "layout")!.Element(c + "manualLayout")!;
            manual.Element(c + "hMode")!.SetAttributeValue("val", "factor");
            manual.Element(c + "h")!.SetAttributeValue("val", (height / row.GetProperty("Frame").GetProperty("Height").GetDouble()).ToString("R", CultureInfo.InvariantCulture));
            entry.Delete(); using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open(); xml.Save(output);
        }
        TestAssert.Equal(actual, ManualPlotBounds(RenderManualBytes(stream.ToArray())));
    }

    public static void PptxManualEdgeHeightsPreserveTheirRectangleAcrossValueDirection()
    {
        using var controls = ReadManualControls();
        foreach (string id in new[] { "inner-factor-edge-height", "inner-factor-edge-height-up" })
        {
            var row = controls.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == id);
            byte[] bytes = ReadManualBytes(row);
            using var stream = new MemoryStream(); stream.Write(bytes); stream.Position = 0;
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
            {
                var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
                XDocument xml; using (var input = entry.Open()) xml = XDocument.Load(input);
                XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                xml.Descendants(c + "valAx").First().Element(c + "scaling")!.Element(c + "orientation")!.SetAttributeValue("val", "maxMin");
                entry.Delete(); using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open(); xml.Save(output);
            }
            CheckManualRectangle(RenderManualBytes(stream.ToArray()), row);
        }
    }

    private static void CheckManualRectangle(PdfPage page, JsonElement row)
    {
        var actual = ManualPlotBounds(page, requireAuthoredAxisColors: true);
        var expected = row.GetProperty("NativePlot");
        double left = expected.GetProperty("Left").GetDouble(), bottom = expected.GetProperty("Bottom").GetDouble();
        TestAssert.True(Math.Abs(actual.Left - left) <= .1d && Math.Abs(actual.Bottom - bottom) <= .1d &&
            Math.Abs(actual.Right - left - expected.GetProperty("Width").GetDouble()) <= .1d &&
            Math.Abs(actual.Top - bottom - expected.GetProperty("Height").GetDouble()) <= .1d,
            row.GetProperty("Id").GetString() + ": source factor/edge geometry matches the independent Office rectangle.");
    }

    private static (double Left, double Right, double Bottom, double Top) ManualPlotBounds(PdfPage page, bool requireAuthoredAxisColors = false)
    {
        // The fixtures mark the horizontal axis green and the vertical axis
        // red. Keep batched grid strokes in the rectangle calculation.
        var lines = PptxColumnIntervalTests.Lines(page);
        var horizontal = lines.Where(line => Math.Abs(line.V[1] - line.V[3]) < .001d && Math.Abs(line.V[0] - line.V[2]) > 100d).ToArray();
        var vertical = lines.Where(line => Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) > 100d).ToArray();
        TestAssert.True(horizontal.Length > 0 && vertical.Length > 0 &&
            (!requireAuthoredAxisColors || horizontal.Any(line => line.Color == new RgbColor(0, 170, 0)) &&
                vertical.Any(line => line.Color == new RgbColor(255, 0, 0))), "Both painted plot axes must be present.");
        return (horizontal.Min(line => Math.Min(line.V[0], line.V[2])), horizontal.Max(line => Math.Max(line.V[0], line.V[2])),
            vertical.Min(line => Math.Min(line.V[1], line.V[3])), vertical.Max(line => Math.Max(line.V[1], line.V[3])));
    }

    private static JsonDocument ReadManualControls() => JsonDocument.Parse(File.ReadAllText(Path.Combine(ManualCases, "office-frames.json")));
    private static byte[] ReadManualBytes(JsonElement row)
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine(ManualCases, row.GetProperty("Id").GetString() + ".pptx"));
        TestAssert.Equal(row.GetProperty("Sha256").GetString()!, Convert.ToHexString(SHA256.HashData(bytes)));
        return bytes;
    }
    private static PdfPage RenderManual(JsonElement row) => RenderManualBytes(ReadManualBytes(row));
    private static PdfPage RenderManualBytes(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
