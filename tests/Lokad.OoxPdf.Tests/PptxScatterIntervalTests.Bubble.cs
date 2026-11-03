using System.Globalization;
using System.IO.Compression;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxBubbleMeasuredScalesPreserveNativeIntervalsAndContent()
    {
        foreach (BubbleCase sample in BubbleCases())
        {
            PdfPage page = RenderBubble(sample);
            TestAssert.Equal("0.5,1.75,3,4.25,5.5,6.75,8," + string.Join(",", sample.ExpectedYLabels), string.Join(",", AutoLabels(page)));
            TestAssert.Equal(sample.ExpectedCircles, BubbleCircles(page).Count(c => c.Radius > 8d));
            TestAssert.True(BubbleText(page, "CSD").Contains(sample.ExpectedValueLabel), sample.Id + ": source value label must survive.");
        }
        var maxima = BubbleCases().Where(c => c.YFactor == 1d && c.YShift == 0d && c.Input == "bubble-measured-source.pptx" && c.Scale > 0)
            .ToDictionary(c => c.Scale, c => BubbleCircles(RenderBubble(c)).Max(circle => circle.Radius));
        // Ratios are independent native observations on an unchanged container.
        TestAssert.True(Math.Abs(maxima[50] / maxima[100] - 13d / 23d) < .001d, "Native50% circle ratio.");
        TestAssert.True(Math.Abs(maxima[125] / maxima[100] - 13d / 11d) < .001d, "Native125% circle ratio.");
        TestAssert.True(Math.Abs(maxima[200] / maxima[100] - 13d / 8d) < .001d, "Native200% circle ratio.");
    }

    public static void PptxBubbleSubunitMeasuresKeepPhysicalCircles()
    {
        BubbleCase original = BubbleCases().Single(c => c.Id == "bubble-query-base");
        BubbleCase scaled = BubbleCases().Single(c => c.Id == "bubble-heldout-subunit");
        var before = BubbleCircles(RenderBubble(original)).Where(c => c.Radius > 8d).ToArray();
        var after = BubbleCircles(RenderBubble(scaled)).Where(c => c.Radius > 8d).ToArray();
        TestAssert.Equal(4, after.Length);
        for (int i = 0; i < before.Length; i++)
        {
            TestAssert.True(Math.Abs(before[i].Radius - after[i].Radius) < .002d, "Changing Y/Z units must not shrink a circle.");
            TestAssert.True(Math.Abs(before[i].Y - after[i].Y) < .002d, "Source rescaling must retain physical Y placement.");
        }
    }

    public static void PptxBubbleCompoundSeriesShareGeometryAndCircularKeys()
    {
        BubbleCase compound = BubbleCases().Single(c => c.Id == "bubble-heldout-two-series");
        BubbleCase original = BubbleCases().Single(c => c.Id == "bubble-query-base");
        PdfPage single = RenderBubble(original, xml => xml.Descendants(CurveC + "autoTitleDeleted").Single().SetAttributeValue("val", "1"));
        PdfPage actual = RenderBubble(compound);
        TestAssert.Equal("Markets,Supply", string.Join(",", BubbleText(actual, "CL")));
        TestAssert.True(BubbleText(actual, "CSD").Contains("79"), "Splitting a series must preserve its authored value label.");
        TestAssert.Equal(2, BubbleCircles(actual).Count(c => c.Radius <= 8d));
        var before = BubbleCircles(single).Where(c => c.Radius > 8d).ToArray();
        var after = BubbleCircles(actual).Where(c => c.Radius > 8d).ToArray();
        TestAssert.Equal(4, after.Length);
        for (int i = 0; i < before.Length; i++)
            TestAssert.True(Math.Abs(before[i].Radius - after[i].Radius) < .002d &&
                Math.Abs(before[i].X - after[i].X) < .002d && Math.Abs(before[i].Y - after[i].Y) < .002d,
                "Series composition must keep the shared source geometry.");
    }

    public static void PptxBubbleSceneScaleAndRejectedDeclarationsKeepOwnership()
    {
        BubbleCase sample = BubbleCases().Single(c => c.Id == "bubble-query-base");
        var parts = BubbleParts(sample.Input);
        string name = parts.Keys.Single(p => p.StartsWith("ppt/charts/chart", StringComparison.Ordinal) && p.EndsWith(".xml", StringComparison.Ordinal));
        var chart = PptxTests.BuildSingleChartScene(System.Text.Encoding.UTF8.GetString(parts[name]))!;
        var plot = chart.Plots.Single();
        TestAssert.Equal(100d, plot.BubbleScale);
        TestAssert.Equal("100", plot.BubbleScaleValue);
        plot.Source.Element(CurveC + "bubbleScale")!.SetAttributeValue("val", "0");
        TestAssert.Equal(100d, plot.BubbleScale);
        TestAssert.Equal("100", plot.BubbleScaleValue);
        foreach (string rejected in new[] { "-1", "301", "0.5", "NaN", "Infinity" })
        {
            PdfPage page = RenderBubble(sample, xml => xml.Descendants(CurveC + "bubbleScale").Single().SetAttributeValue("val", rejected));
            TestAssert.True(BubbleText(page, "CSD").Contains("79"), "Rejected source options keep supported neighboring content.");
            TestAssert.Equal(4, BubbleCircles(page).Count(c => c.Radius > 8d));
        }
    }

    public static void PptxBubbleFixedBoundsAndReversalSharePreparedPlacement()
    {
        BubbleCase sample = BubbleCases().Single(c => c.Id == "bubble-query-base");
        PdfPage Render(bool reversed) => RenderBubble(sample, xml => {
            XElement axis = xml.Descendants(CurveC + "valAx").Last();
            axis.Element(CurveC + "scaling")!.Add(new XElement(CurveC + "min", new XAttribute("val", "-7")),
                new XElement(CurveC + "max", new XAttribute("val", "103")));
            axis.Add(new XElement(CurveC + "majorUnit", new XAttribute("val", "25")));
            if (reversed) foreach (var orientation in xml.Descendants(CurveC + "orientation")) orientation.SetAttributeValue("val", "maxMin");
        });
        PdfPage normal = Render(false), reverse = Render(true);
        TestAssert.Equal("0.5,1.75,3,4.25,5.5,6.75,8,-7,18,43,68,93", string.Join(",", AutoLabels(reverse)));
        var before = BubbleCircles(normal).Where(c => c.Radius > 8d).ToArray();
        var after = BubbleCircles(reverse).Where(c => c.Radius > 8d).ToArray();
        var first = BubblePlot(normal); var second = BubblePlot(reverse);
        for (int i = 0; i < before.Length; i++)
        {
            TestAssert.True(Math.Abs(before[i].Radius - after[i].Radius) < .002d, "Reversal does not change circle size.");
            TestAssert.True(Math.Abs((before[i].X - first.Left) / first.Width + (after[i].X - second.Left) / second.Width - 1d) < .00005d &&
                Math.Abs((before[i].Y - first.Bottom) / first.Height + (after[i].Y - second.Bottom) / second.Height - 1d) < .00005d,
                "Source reversal reflects centers through their own admitted plot.");
        }
    }

    private sealed record BubbleCase(string Id, string Input, int Scale, double YFactor, double SizeFactor,
        double YShift, string[] ExpectedYLabels, int ExpectedCircles, string ExpectedValueLabel);

    private static BubbleCase[] BubbleCases()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(BubbleCasePath("bubble-measured-office.json")));
        return JsonSerializer.Deserialize<BubbleCase[]>(json.RootElement.GetProperty("Cases").GetRawText())!;
    }
    private static string BubbleCasePath(string name) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name));
    private static Dictionary<string, byte[]> BubbleParts(string name)
    {
        using var input = File.OpenRead(BubbleCasePath(name));
        using var archive = new ZipArchive(input);
        return archive.Entries.Where(e => !string.IsNullOrEmpty(e.Name)).ToDictionary(e => e.FullName, e => {
            using var source = e.Open(); using var buffer = new MemoryStream(); source.CopyTo(buffer); return buffer.ToArray();
        });
    }
    private static PdfPage RenderBubble(BubbleCase sample, Action<XDocument>? transform = null)
    {
        var parts = BubbleParts(sample.Input);
        string name = parts.Keys.Single(p => p.StartsWith("ppt/charts/chart", StringComparison.Ordinal) && p.EndsWith(".xml", StringComparison.Ordinal));
        XDocument xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts[name]));
        xml.Descendants(CurveC + "bubbleScale").Single().SetAttributeValue("val", sample.Scale.ToString(CultureInfo.InvariantCulture));
        foreach (var value in xml.Descendants(CurveC + "yVal").Descendants(CurveC + "v"))
            value.Value = (double.Parse(value.Value, CultureInfo.InvariantCulture) * sample.YFactor + sample.YShift).ToString("R", CultureInfo.InvariantCulture);
        foreach (var value in xml.Descendants(CurveC + "bubbleSize").Descendants(CurveC + "v"))
            value.Value = (double.Parse(value.Value, CultureInfo.InvariantCulture) * sample.SizeFactor).ToString("R", CultureInfo.InvariantCulture);
        transform?.Invoke(xml);
        parts[name] = TestFixtures.Utf8(xml.ToString(SaveOptions.DisableFormatting));
        using var stream = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
    private static string[] BubbleText(PdfPage page, string role) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
        fontResourceFilter: name => name.StartsWith(role, StringComparison.Ordinal)).Select(s => s.Text).ToArray();

    // Only the fixture's four-arc filled circles are inspected. This is not a
    // general PDF parser; coordinates include the full path before plot clipping.
    private static (double X, double Y, double Radius)[] BubbleCircles(PdfPage page) =>
        Regex.Matches(page.Content, @"(?<path>[-\d.]+ [-\d.]+ m(?:\s+[-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ c){4})\s+(?:h\s+)?f\*")
            .Select(m => Regex.Matches(m.Groups["path"].Value, @"-?\d+(?:\.\d+)?").Select(n => double.Parse(n.Value, CultureInfo.InvariantCulture)).ToArray())
            .Select(v => { var x = v.Where((_, i) => i % 2 == 0).ToArray(); var y = v.Where((_, i) => i % 2 != 0).ToArray();
                return ((x.Min() + x.Max()) / 2d, (y.Min() + y.Max()) / 2d, (x.Max() - x.Min()) / 2d); }).ToArray();
    private static (double Left, double Bottom, double Width, double Height) BubblePlot(PdfPage page)
    {
        var lines = PptxColumnIntervalTests.Lines(page).Where(p => p.Color == new RgbColor(90, 90, 90)).ToArray();
        var x = lines.Single(p => Math.Abs(p.V[1] - p.V[3]) < .001d && Math.Abs(p.V[0] - p.V[2]) > 100d).V;
        var y = lines.Single(p => Math.Abs(p.V[0] - p.V[2]) < .001d && Math.Abs(p.V[1] - p.V[3]) > 100d).V;
        return (x[0], y[1], x[2] - x[0], y[3] - y[1]);
    }
}
