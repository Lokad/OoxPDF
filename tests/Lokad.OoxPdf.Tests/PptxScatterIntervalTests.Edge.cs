using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxScatterEdgeWidthSharesNativeAxesAndIntervals()
    {
        foreach (var sample in EdgeCases())
        {
            PdfPage page = RenderEdge(EdgeChart(sample.XFactor, sample.RightEdge));
            var axes = EdgeAxes(page);
            double[] expectedX = sample.ExpectedXAxis, expectedY = sample.ExpectedYAxis;
            double[] x = [axes.X[0], page.Height - axes.X[1], axes.X[2], page.Height - axes.X[3]];
            double[] y = [axes.Y[0], page.Height - axes.Y[1], axes.Y[2], page.Height - axes.Y[3]];
            TestAssert.True(x.Zip(expectedX).Concat(y.Zip(expectedY)).All(p => Math.Abs(p.First - p.Second) < .04d),
                sample.Id + ": use independent Office axes, not an equivalent-looking source ratio.");
            string expected = string.Join(",", Enumerable.Range(0, (int)(8d / sample.ExpectedMajorUnit) + 1)
                .Select(i => (i * sample.ExpectedMajorUnit).ToString(CultureInfo.InvariantCulture))) + ",7,32,57,82";
            TestAssert.Equal(expected, string.Join(",", AutoLabels(page)));
            TestAssert.True(EdgeValueLabel(page).Text == "79", "Preserve the authored value label.");
        }
    }

    public static void PptxScatterEquivalentWidthDeclarationsPreserveContent()
    {
        foreach (double width in new[] { .35d, .78d, .8d })
        {
            XDocument factor = EdgeChart(.04d, width, "factor");
            factor.Descendants(CurveC + "valAx").First().Element(CurveC + "scaling")!.Element(CurveC + "max")!.SetAttributeValue("val", "9");
            XDocument edge = new(factor);
            edge.Descendants(CurveC + "wMode").Single().SetAttributeValue("val", "edge");
            edge.Descendants(CurveC + "w").Single().SetAttributeValue("val", (.5d + .04d + width / 2d).ToString("R", CultureInfo.InvariantCulture));
            string expected = RenderEdge(factor).Content;
            TestAssert.Equal(expected, RenderEdge(edge).Content);
            XDocument absolute = new(edge);
            absolute.Descendants(CurveC + "xMode").Single().SetAttributeValue("val", "edge");
            absolute.Descendants(CurveC + "x").Single().SetAttributeValue("val", (.5d + .04d - width / 2d).ToString("R", CultureInfo.InvariantCulture));
            TestAssert.Equal(expected, RenderEdge(absolute).Content);
        }
    }

    public static void PptxScatterEdgeWidthReversalPreservesPointLabels()
    {
        XDocument source = EdgeChart(.10d, .82d);
        source.Descendants(CurveC + "valAx").First().Element(CurveC + "scaling")!
            .Element(CurveC + "max")!.SetAttributeValue("val", "9");
        PdfPage normal = RenderEdge(source);
        var before = EdgeValueLabel(normal);
        foreach (var orientation in source.Descendants(CurveC + "orientation")) orientation.SetAttributeValue("val", "maxMin");
        PdfPage reverse = RenderEdge(source);
        var after = EdgeValueLabel(reverse);
        // Source point X=3,Y=79 in0..9 and7..103 reflects in the same
        // native220pt by311.04pt rectangle; its label offset remains occupied.
        Near((1d - 2d * 3d / 9d) * 220d, after.X - before.X);
        Near(-.5d * 311.04d, after.Y - before.Y);
        TestAssert.Equal(string.Join(",", AutoLabels(normal)), string.Join(",", AutoLabels(reverse)));
    }

    public static void PptxScatterEdgeWidthKeepsSceneAuthorityAndFallbacks()
    {
        XDocument xml = EdgeChart(.10d, .82d);
        var scene = PptxTests.BuildSingleChartScene(xml.ToString())!;
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.NonPublic;
        Type axisType = typeof(PptxRenderer).GetNestedType("ChartAxisSource", flags)!;
        Type frameType = typeof(PptxRenderer).GetNestedType("ChartFrameBox", flags)!;
        object frame = Activator.CreateInstance(frameType, [72d, 72d, 500d, 432d])!;
        double? Read(PptxSceneChart chart, XDocument document) => (double?)CurveInvoke("ReadScatterSourceIntervalTarget", PptxTheme.Empty,
            chart, document, frame, Activator.CreateInstance(axisType, [chart.Axes.First(), document.Descendants(CurveC + "valAx").First()])!);
        double? original = Read(scene, xml);
        TestAssert.True(original is not null, "A resolved inner edge supplies interval geometry.");
        xml.Descendants(CurveC + "wMode").Single().SetAttributeValue("val", "factor");
        xml.Descendants(CurveC + "w").Single().SetAttributeValue("val", ".99");
        TestAssert.Equal(original, Read(scene, xml));
        foreach (string fallback in new[] { "outer", "unknown-width", "missing-position", "degenerate", "Infinity" })
        {
            XDocument source = EdgeChart(.10d, .82d);
            if (fallback == "outer") source.Descendants(CurveC + "layoutTarget").Single().SetAttributeValue("val", "outer");
            if (fallback == "unknown-width") source.Descendants(CurveC + "wMode").Single().SetAttributeValue("val", "unknown");
            if (fallback == "missing-position") source.Descendants(CurveC + "x").Single().Remove();
            if (fallback == "degenerate") source.Descendants(CurveC + "w").Single().SetAttributeValue("val", ".3");
            if (fallback == "Infinity") source.Descendants(CurveC + "w").Single().SetAttributeValue("val", "Infinity");
            TestAssert.True(Read(PptxTests.BuildSingleChartScene(source.ToString())!, source) is null,
                "Unresolved declarations keep their scale fallback: " + fallback);
        }
    }

    private sealed record EdgeCase(string Id, double XFactor, double RightEdge, double ExpectedMajorUnit,
        double[] ExpectedXAxis, double[] ExpectedYAxis);
    private static EdgeCase[] EdgeCases()
    {
        using var document = JsonDocument.Parse(File.ReadAllText(ChartCasePath("scatter-edge-width-office.json")));
        return JsonSerializer.Deserialize<EdgeCase[]>(document.RootElement.GetProperty("Cases").GetRawText())!;
    }
    private static XDocument EdgeChart(double xFactor, double width, string widthMode = "edge")
    {
        var parts = ReadChartParts("scatter-edge-width-source.pptx");
        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["ppt/charts/chart1.xml"]));
        xml.Descendants(CurveC + "x").Single().SetAttributeValue("val", xFactor.ToString("R", CultureInfo.InvariantCulture));
        xml.Descendants(CurveC + "w").Single().SetAttributeValue("val", width.ToString("R", CultureInfo.InvariantCulture));
        xml.Descendants(CurveC + "wMode").Single().SetAttributeValue("val", widthMode);
        return xml;
    }
    private static PdfPage RenderEdge(XDocument source)
    {
        var parts = ReadChartParts("scatter-edge-width-source.pptx");
        parts["ppt/charts/chart1.xml"] = TestFixtures.Utf8(source.ToString(SaveOptions.DisableFormatting));
        using var stream = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(new PptxReader().Read(package, default),
            package, null, default).ToBlockingEnumerable().Single();
    }
    private static (double[] X, double[] Y) EdgeAxes(PdfPage page)
    {
        var lines = PptxColumnIntervalTests.Lines(page).Where(p => p.Color == default).ToArray();
        return (lines.Single(p => Math.Abs(p.V[1] - p.V[3]) < .001d && Math.Abs(p.V[0] - p.V[2]) > 100d).V,
            lines.Single(p => Math.Abs(p.V[0] - p.V[2]) < .001d && Math.Abs(p.V[1] - p.V[3]) > 100d).V);
    }
    private static (string Text, double X, double Y) EdgeValueLabel(PdfPage page)
    {
        var show = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
            fontResourceFilter: name => name.StartsWith("CSD", StringComparison.Ordinal)).Single(p => p.Text == "79");
        return (show.Text, show.X, show.Y);
    }
}
