using System.Globalization;
using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxScatterAndBubbleIntervalsUseAuthoredMinima()
    {
        // Pinned independent Office exports establish source-relative intervals
        // on both value axes; a non-interval maximum is not another tick.
        foreach (bool bubble in new[] { false, true })
        foreach (var test in new[]
        {
            (-7d, 103d, "-7,18,43,68,93"),
            (7d, 103d, "7,32,57,82"),
            (7d, 107d, "7,32,57,82,107"),
            (0d, 103d, "0,25,50,75,100")
        })
        {
            PdfPage page = Render(bubble, .5d, test.Item1, test.Item2);
            string[] labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).Select(l => l.Text).ToArray();
            TestAssert.Equal("0.5,1.75,3,4.25,5.5,6.75,8," + test.Item3, string.Join(",", labels));
        }
    }

    public static void PptxScatterAndBubbleMarksAndGridsShareSourceIntervals()
    {
        foreach (bool bubble in new[] { false, true })
        foreach (double xMinimum in new[] { -.5d, .5d })
        {
            var lines = PptxColumnIntervalTests.Lines(Render(bubble, xMinimum, 7d, 103d));
            double[] xAxis = lines.Single(p => Math.Abs(p.V[1] - p.V[3]) < .001d &&
                Math.Abs(p.V[0] - p.V[2]) > 100d && p.Color == default).V;
            double[] yAxis = lines.Single(p => Math.Abs(p.V[0] - p.V[2]) < .001d &&
                Math.Abs(p.V[1] - p.V[3]) > 100d && p.Color == default).V;
            double XValue(double x) => xMinimum + (x - xAxis[0]) / (xAxis[2] - xAxis[0]) * (8.5d - xMinimum);
            double YValue(double y) => 7d + (y - yAxis[1]) / (yAxis[3] - yAxis[1]) * 96d;
            double[] xTicks = lines.Where(p => p.Color == new RgbColor(0, 255, 0) &&
                Math.Abs(p.V[0] - p.V[2]) < .001d && Math.Abs(p.V[1] - p.V[3]) < 20d).Select(p => XValue(p.V[0])).ToArray();
            var yLines = lines.Where(p => p.Color == new RgbColor(0, 0, 255) && Math.Abs(p.V[1] - p.V[3]) < .001d).ToArray();
            Check(xTicks, Enumerable.Range(0, xMinimum < 0d ? 8 : 7).Select(i => xMinimum + 1.25d * i).ToArray());
            Check(yLines.Where(p => Math.Abs(p.V[0] - p.V[2]) < 20d).Select(p => YValue(p.V[1])).ToArray(), [7d, 32d, 57d, 82d]);
            Check(yLines.Where(p => Math.Abs(p.V[0] - p.V[2]) > 100d).Select(p => YValue(p.V[1])).ToArray(), [32d, 57d, 82d]);
        }

        static void Check(double[] actual, double[] expected) => TestAssert.True(actual.Length == expected.Length &&
            actual.Zip(expected).All(p => Math.Abs(p.First - p.Second) < .002d),
            "Intervals must agree in painted plot coordinates: " + string.Join(",", actual));
    }

    public static void PptxScatterIntervalPolicyKeepsSceneAuthorityAndFallbacks()
    {
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Static;
        Type extentsType = typeof(PptxRenderer).GetNestedType("ChartValueExtents", BindingFlags.NonPublic)!;
        var reader = typeof(PptxRenderer).GetMethod("ReadSceneChartBubbleValueAxisOptions", flags)!;
        XElement xml = XElement.Parse(Axis("1", "b", "2", "7", "103", "25", "00FF00"));
        string Policy(PptxSceneChartAxis? scene, double min = 7d)
        {
            object options = reader.Invoke(null, [scene,
                Activator.CreateInstance(extentsType, [min, 103d]), 8d])!;
            return options.GetType().GetProperty("TickPolicy")!.GetValue(options)!.ToString()!;
        }
        PptxSceneChartAxis parsed = PptxTests.BuildSingleChartScene(
            $"<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea>{xml}</c:plotArea></c:chart></c:chartSpace>")!.Axes.Single();
        TestAssert.Equal("SourceMinimumIntervals", Policy(parsed));
        TestAssert.Equal("ZeroAlignedWithEndpoints", Policy(null));
        TestAssert.Equal("ZeroAlignedWithEndpoints", Policy(parsed, 0d));
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (string minimum in new[] { "", "NaN", "7" })
        foreach (bool log in new[] { false, true })
        {
            XElement source = new(xml);
            XElement scaling = source.Element(c + "scaling")!;
            scaling.Element(c + "min")!.Remove();
            if (minimum.Length != 0) scaling.Add(new XElement(c + "min", new XAttribute("val", minimum)));
            if (log) scaling.Add(new XElement(c + "logBase", new XAttribute("val", "10")));
            string chart = $"<c:chartSpace xmlns:c=\"{c}\"><c:chart><c:plotArea>{source}</c:plotArea></c:chart></c:chartSpace>";
            PptxSceneChartAxis scene = PptxTests.BuildSingleChartScene(chart)!.Axes.Single();
            string expected = minimum == "7" && !log ? "SourceMinimumIntervals" : "ZeroAlignedWithEndpoints";
            TestAssert.Equal(expected, Policy(scene));
        }
    }

    public static void PptxScatterMeasurementsRetainTranslatedAndSubunitDataExtrema()
    {
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Static;
        var read = typeof(PptxRenderer).GetMethod("ReadSceneScatterSeries", flags)!;
        foreach (bool bubble in new[] { false, true })
        foreach (double scale in new[] { 1d, .01d })
        foreach (double shift in new[] { -100d, 0d, 100d })
        {
            string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
            string kind = bubble ? "bubbleChart" : "scatterChart";
            string Values(string name, double first, double last) => $"<c:{name}><c:numLit><c:pt idx=\"0\"><c:v>{Number(first)}</c:v></c:pt><c:pt idx=\"1\"><c:v>{Number(last)}</c:v></c:pt></c:numLit></c:{name}>";
            string source = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:{{kind}}>
                  <c:ser><c:idx val="0"/><c:order val="0"/>
                    {{Values("xVal", shift+scale, shift+8d*scale)}}{{Values("yVal", shift+2d*scale, shift+79d*scale)}}
                    {{(bubble ? Values("bubbleSize", 1d, 4d) : "")}}
                  </c:ser></c:{{kind}}></c:plotArea></c:chart></c:chartSpace>
                """;
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var scene = PptxTests.BuildSingleChartScene(source)!;
            object series = read.Invoke(null, [scene.Plots.Single(), bubble, null, true])!;
            foreach (var axis in new[] { ("X", 1d, 8d), ("Y", 2d, 79d) })
            {
                object measured = typeof(PptxRenderer).GetMethod("MeasureScatterChartExtents", flags)!.Invoke(null, [series, null, null, CancellationToken.None])!;
                var extents = (PptxRenderer.ChartValueExtents)measured.GetType().GetField(axis.Item1 == "X" ? "Item1" : "Item2")!.GetValue(measured)!;
                TestAssert.True(Math.Abs(extents.Min-(shift+axis.Item2*scale))<1e-10d && Math.Abs(extents.Max-(shift+axis.Item3*scale))<1e-10d,
                    "Measured extrema follow source translation/scaling without inserting zero or a unit-sized range.");
            }
        }
    }

    public static void PptxPointAxisFallbackOwnsLegacyZeroAndUnitBounds()
    {
        var reader = typeof(PptxRenderer).GetMethod("ReadSceneBubbleChartValueAxisExtents", BindingFlags.NonPublic|BindingFlags.Static)!;
        foreach (var test in new[] { (90d, 99d, 0d, 99d), (.01d, .08d, 0d, 1d), (-8d, -1d, -8d, 1d), (-3d, 2d, -3d, 2d) })
        {
            var actual = (PptxRenderer.ChartValueExtents)reader.Invoke(null, [null, new PptxRenderer.ChartValueExtents(test.Item1,test.Item2), 5d, true])!;
            TestAssert.Equal(new PptxRenderer.ChartValueExtents(test.Item3,test.Item4), actual);
        }
    }

    private static PdfPage Render(bool bubble, double xMin, double yMin, double yMax, bool reverseX = false, bool reverseY = false,
        string? labelPosition = null, double? horizontalCrossing = null, double? verticalCrossing = null, bool labels = false,
        string? crossBetween = null, bool verticalMaximum = false)
    {
        string kind = bubble ? "bubbleChart" : "scatterChart";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
              <c:{{kind}}><c:scatterStyle val="marker"/><c:ser><c:idx val="0"/><c:order val="0"/>
                <c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt></c:numLit></c:xVal>
                <c:yVal><c:numLit><c:pt idx="0"><c:v>20</c:v></c:pt><c:pt idx="1"><c:v>79</c:v></c:pt></c:numLit></c:yVal>
                {{(bubble ? "<c:bubbleSize><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt><c:pt idx=\"1\"><c:v>3</c:v></c:pt></c:numLit></c:bubbleSize>" : "")}}
                {{(labels ? "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>" : "")}}
              </c:ser><c:axId val="1"/><c:axId val="2"/></c:{{kind}}>
              {{Axis("1", "b", "2", xMin.ToString(CultureInfo.InvariantCulture), "8.5", "1.25", "00FF00", reverseX, labelPosition, horizontalCrossing, crossBetween)}}
              {{Axis("2", "l", "1", yMin.ToString(CultureInfo.InvariantCulture), yMax.ToString(CultureInfo.InvariantCulture), "25", "0000FF", reverseY, labelPosition, verticalCrossing, crossBetween, verticalMaximum)}}
            </c:plotArea></c:chart></c:chartSpace>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }

    private static string Axis(string id, string side, string cross, string min, string max, string major, string color,
        bool reversed = false, string? labelPosition = null, double? crossing = null, string? crossBetween = null, bool maximum = false) => $$"""
        <c:valAx xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:axId val="{{id}}"/><c:scaling><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/><c:min val="{{min}}"/><c:max val="{{max}}"/></c:scaling>
          <c:axPos val="{{side}}"/><c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="{{color}}"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
          <c:majorTickMark val="out"/>{{(labelPosition is null ? "" : $"<c:tickLblPos val=\"{labelPosition}\"/>")}}<c:crossAx val="{{cross}}"/>
          {{(crossing is { } value ? $"<c:crossesAt val=\"{value.ToString(CultureInfo.InvariantCulture)}\"/>" : $"<c:crosses val=\"{(maximum ? "max" : "autoZero")}\"/>")}}
          {{(crossBetween is null ? "" : $"<c:crossBetween val=\"{crossBetween}\"/>")}}<c:majorUnit val="{{major}}"/>
        </c:valAx>
        """;
}
