using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
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
        foreach (var row in controls.RootElement.EnumerateArray().Where(row => passive.Contains(row.GetProperty("Id").GetString())))
            CheckManualRectangle(RenderManual(row), row);
    }

    public static void PptxManualScatterEdgeHeightKeepsIndependentOfficeRectangle()
    {
        using var controls = ReadManualControls();
        var row = controls.RootElement.EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "scatter-valid-edge-height");
        CheckManualRectangle(RenderManual(row), row);
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
        var actual = ManualPlotBounds(page);
        var expected = row.GetProperty("NativePlot");
        double left = expected.GetProperty("Left").GetDouble(), bottom = expected.GetProperty("Bottom").GetDouble();
        TestAssert.True(Math.Abs(actual.Left - left) <= .1d && Math.Abs(actual.Bottom - bottom) <= .1d &&
            Math.Abs(actual.Right - left - expected.GetProperty("Width").GetDouble()) <= .1d &&
            Math.Abs(actual.Top - bottom - expected.GetProperty("Height").GetDouble()) <= .1d,
            row.GetProperty("Id").GetString() + ": source factor/edge geometry matches the independent Office rectangle.");
    }

    private static (double Left, double Right, double Bottom, double Top) ManualPlotBounds(PdfPage page)
    {
        // Include batched paths and both axes: an interior category-axis line
        // alone does not define the top/bottom of its plot.
        var lines = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) m\s+([-\d.]+) ([-\d.]+) l")
            .Select(match => match.Groups.Cast<Group>().Skip(1).Select(group => double.Parse(group.Value, CultureInfo.InvariantCulture)).ToArray()).ToArray();
        var horizontal = lines.Where(v => Math.Abs(v[1] - v[3]) < .001d && Math.Abs(v[0] - v[2]) > 100d).ToArray();
        var vertical = lines.Where(v => Math.Abs(v[0] - v[2]) < .001d && Math.Abs(v[1] - v[3]) > 100d).ToArray();
        TestAssert.True(horizontal.Length > 0 && vertical.Length > 0, "Both independent plot axes must be present.");
        return (horizontal.Min(v => Math.Min(v[0], v[2])), horizontal.Max(v => Math.Max(v[0], v[2])),
            vertical.Min(v => Math.Min(v[1], v[3])), vertical.Max(v => Math.Max(v[1], v[3])));
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
