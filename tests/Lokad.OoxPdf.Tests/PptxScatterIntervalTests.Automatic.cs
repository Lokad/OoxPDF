using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxScatterRequiredIntegerDigitsReachNumericCaptions()
    {
        foreach (AutoCase sample in AutoCases().Where(c => c.NumberFormat == "0000000000"))
        {
            string[] actual = AutoLabels(RenderAuto(sample, AutoChart(sample)));
            string[] expected = sample.Id == "auto-x-width-360-long-format"
                ? ["0000000000", "0000000002", "0000000004", "0000000006", "0000000008", "0000000010", "7", "32", "57", "82"]
                : ["0000000000", "0000000001", "0000000002", "0000000003", "0000000004", "0000000005", "0000000006", "0000000007", "0000000008", "0000000009", "7", "32", "57", "82"];
            TestAssert.Equal(string.Join(",", expected), string.Join(",", actual));
        }
    }

    public static void PptxScatterAutomaticIntervalsMatchIndependentOfficeGeometry()
    {
        foreach (AutoCase sample in AutoCases())
        {
            PdfPage page = RenderAuto(sample, AutoChart(sample));
            double[] expected = Enumerable.Range(0, (int)Math.Round((sample.ExpectedMaximum - sample.ExpectedMinimum) / sample.ExpectedUnit) + 1)
                .Select(i => sample.ExpectedMinimum + i * sample.ExpectedUnit).ToArray();
            string[] labels = AutoLabels(page);
            TestAssert.True(labels.Length == expected.Length + 4, sample.Id + ": label count differs: " + labels.Length);
            TestAssert.True(labels.Take(expected.Length).Select(text => double.Parse(text, CultureInfo.InvariantCulture))
                .Zip(expected).All(pair => Math.Abs(pair.First - pair.Second) < 1e-9d), sample.Id + ": numeric labels must use the independent Office scale.");
            var lines = PptxColumnIntervalTests.Lines(page);
            double[] axis = lines.Single(line => line.Color == default && Math.Abs(line.V[1] - line.V[3]) < .001d &&
                Math.Abs(line.V[0] - line.V[2]) > 100d).V;
            double[] ticks = lines.Where(line => line.Color == new RgbColor(0, 255, 0) &&
                Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) < 20d).Select(line => line.V[0]).ToArray();
            TestAssert.True(ticks.Length == expected.Length, sample.Id + ": tick count differs: " + ticks.Length);
            TestAssert.True(ticks.Zip(expected).All(pair => Math.Abs(pair.First - (axis[0] + (axis[2] - axis[0]) *
                (pair.Second - sample.ExpectedMinimum) / (sample.ExpectedMaximum - sample.ExpectedMinimum))) < .002d),
                sample.Id + ": marks and labels must share the resolved scale.");
        }
    }

    public static void PptxScatterAutomaticIntervalsPreserveSourceUnitsAndRejectedDeclarations()
    {
        AutoCase sample = AutoCases().Single(c => c.Id == "auto-x-width-360-font-20");
        string[] explicitLabels = AutoLabels(RenderAuto(sample, AutoChart(sample, "1")));
        TestAssert.Equal("0,1,2,3,4,5,6,7,8,9,7,32,57,82", string.Join(",", explicitLabels));
        foreach (string rejected in new[] { "NaN", "-1", "Infinity" })
        {
            XDocument xml = XDocument.Parse(AutoChart(sample, rejected));
            string[] source = AutoLabels(RenderAuto(sample, xml.ToString()));
            xml.Descendants(CurveC + "layout").Single().Remove();
            TestAssert.Equal(string.Join(",", AutoLabels(RenderAuto(sample, xml.ToString()))), string.Join(",", source));
        }
        foreach (string fallback in new[] { "outer", "edge", "NaN", "log", "scaling" })
        {
            XDocument xml = XDocument.Parse(AutoChart(sample));
            XElement xAxis = xml.Descendants(CurveC + "valAx").First();
            XElement layout = xml.Descendants(CurveC + "manualLayout").Single();
            if (fallback == "outer") layout.Element(CurveC + "layoutTarget")!.SetAttributeValue("val", "outer");
            if (fallback == "edge") layout.Element(CurveC + "wMode")!.SetAttributeValue("val", "edge");
            if (fallback == "NaN") layout.Element(CurveC + "w")!.SetAttributeValue("val", "NaN");
            if (fallback == "log") xAxis.Element(CurveC + "scaling")!.Add(new XElement(CurveC + "logBase", new XAttribute("val", "10")));
            if (fallback == "scaling") xAxis.Element(CurveC + "scaling")!.Remove();
            string[] source = AutoLabels(RenderAuto(sample, xml.ToString()));
            xml.Descendants(CurveC + "layout").Single().Remove();
            TestAssert.Equal(string.Join(",", AutoLabels(RenderAuto(sample, xml.ToString()))), string.Join(",", source));
        }
    }

    public static void PptxScatterAutomaticIntervalGeometryRemainsSceneOwned()
    {
        AutoCase sample = AutoCases().Single(c => c.Id == "auto-x-width-360-font-20");
        XDocument xml = XDocument.Parse(AutoChart(sample));
        PptxSceneChart scene = PptxTests.BuildSingleChartScene(xml.ToString())!;
        XElement xAxis = xml.Descendants(CurveC + "valAx").First();
        Type sourceType = typeof(PptxRenderer).GetNestedType("ChartAxisSource", System.Reflection.BindingFlags.NonPublic)!;
        Type frameType = typeof(PptxRenderer).GetNestedType("ChartFrameBox", System.Reflection.BindingFlags.NonPublic)!;
        object source = Activator.CreateInstance(sourceType, [scene.Axes.First(), xAxis])!;
        object frame = Activator.CreateInstance(frameType, [72d, 72d, sample.FrameWidth, 360d])!;
        double Read() => (double)CurveInvoke("ReadScatterSourceIntervalTarget", PptxTheme.Empty, scene, xml, frame, source);
        double before = Read();
        xml.Descendants(CurveC + "w").Single().SetAttributeValue("val", ".95");
        xml.Descendants(XName.Get("defRPr", "http://schemas.openxmlformats.org/drawingml/2006/main")).Single().SetAttributeValue("sz", "500");
        TestAssert.Equal(before, Read());
        TestAssert.Equal(2d, before);
    }

    private sealed record AutoCase(string Id, double FrameWidth, double WidthFactor, double FontSize, double[] X,
        double? Minimum, double? Maximum, string NumberFormat, double ExpectedMinimum, double ExpectedMaximum, double ExpectedUnit);
    private static AutoCase[] AutoCases()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "scatter-auto-intervals-office.json"));
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return JsonSerializer.Deserialize<AutoCase[]>(document.RootElement.GetProperty("Cases").GetRawText())!;
    }
    private static string[] AutoLabels(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
        fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).Select(label => label.Text).ToArray();

    private static string AutoChart(AutoCase sample, string? majorUnit = null)
    {
        var data = new RangeCase(sample.Id, "X", sample.X, [1d, 2d, 79d, 4d, 5d, 6d, 7d, 8d], sample.Minimum, sample.Maximum,
            1d, null, false, sample.ExpectedMinimum, sample.ExpectedMaximum);
        XDocument xml = XDocument.Parse(RangeChart(data));
        XElement xAxis = xml.Descendants(CurveC + "valAx").First();
        XElement yAxis = xml.Descendants(CurveC + "valAx").Last();
        yAxis.Element(CurveC + "scaling")!.Element(CurveC + "min")!.SetAttributeValue("val", "7");
        if (majorUnit is null) xAxis.Element(CurveC + "majorUnit")!.Remove();
        else xAxis.Element(CurveC + "majorUnit")!.SetAttributeValue("val", majorUnit);
        xAxis.Add(new XElement(CurveC + "numFmt", new XAttribute("formatCode", sample.NumberFormat), new XAttribute("sourceLinked", "0")));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        xAxis.Add(new XElement(CurveC + "txPr", new XElement(a + "bodyPr"), new XElement(a + "p", new XElement(a + "pPr",
            new XElement(a + "defRPr", new XAttribute("sz", Math.Round(sample.FontSize * 100d)))))));
        xml.Descendants(CurveC + "plotArea").Single().AddFirst(XElement.Parse($$"""
            <c:layout xmlns:c="{{CurveC}}"><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="factor"/><c:yMode val="factor"/>
              <c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0"/><c:y val="0"/>
              <c:w val="{{sample.WidthFactor.ToString("R", CultureInfo.InvariantCulture)}}"/><c:h val=".6"/></c:manualLayout></c:layout>
            """));
        return xml.ToString(SaveOptions.DisableFormatting);
    }
    private static PdfPage RenderAuto(AutoCase sample, string chart)
    {
        string slide = $$"""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="{{CurveC}}">
              <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="{{Math.Round(sample.FrameWidth * 12700d)}}" cy="4572000"/></p:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
            """;
        const string presentation = """
            <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart, new Dictionary<string, byte[]> {
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide), ["ppt/presentation.xml"] = TestFixtures.Utf8(presentation) });
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
