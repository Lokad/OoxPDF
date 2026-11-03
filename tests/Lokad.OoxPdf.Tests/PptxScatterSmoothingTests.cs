using System.Collections;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxScatterSmoothingTests
{
    public static void PptxScatterSmoothingKeepsSourceAuthorityAndExplicitFalse()
    {
        var read = typeof(PptxRenderer).GetMethod("ReadSceneOrXmlChartScatterOptions", BindingFlags.NonPublic | BindingFlags.Static)!;
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (string style in new[] { "lineMarker", "smoothMarker" })
        foreach (string? smooth in new string?[] { null, "0", "1" })
        {
            string source = Chart(style, smooth);
            var scene = PptxTests.BuildSingleChartScene(source)!;
            var xml = XDocument.Parse(source).Descendants(c + "scatterChart").Single();
            var mismatch = XDocument.Parse(Chart(style, smooth == "0" ? "1" : "0")).Descendants(c + "scatterChart").Single();
            foreach (var sourcePair in new[] { (Plot: (PptxSceneChartPlot?)null, Xml: xml), (Plot: scene.Plots[0], Xml: mismatch) })
            {
                object options = read.Invoke(null, [sourcePair.Plot, sourcePair.Xml])!;
                object option = ((IEnumerable)options.GetType().GetProperty("SmoothSeries")!.GetValue(options)!).Cast<object>().Single();
                TestAssert.Equal(smooth != "0", PptxTests.ChartBooleanOptionValue(option));
                TestAssert.Equal(smooth is not null, PptxTests.ChartBooleanOptionIsDefined(option));
            }
            TestAssert.Equal(smooth is null, scene.Plots[0].Series[0].Smooth is null);
        }
    }

    public static void PptxScatterSmoothingDefaultsPaintCurvesAndKeepStraightAndHiddenLines()
    {
        string absent = Render(Chart("lineMarker", null));
        string enabled = Render(Chart("lineMarker", "1"));
        string straight = Render(Chart("lineMarker", "0"));
        string straightStyle = Render(Chart("smoothMarker", "0"));
        string hidden = Render(Chart("lineMarker", null, hidden: true));
        TestAssert.Equal(enabled, absent);
        TestAssert.Equal(straight, straightStyle);
        TestAssert.True(Curves(absent) == 2, "An omitted smoothing element paints the same connected curves as explicit true.");
        TestAssert.True(Curves(straight) == 0 && Curves(hidden) == 0,
            "Explicit false keeps straight segments and series noFill suppresses the connection.");

        static int Curves(string content) => content.Split('\n').Count(line => line.TrimEnd().EndsWith(" c", StringComparison.Ordinal));
    }

    public static void PptxScatterConnectionDefaultsKeepSceneOwnershipAndRawStyles()
    {
        var read = typeof(PptxRenderer).GetMethod("ReadSceneOrXmlChartScatterOptions", BindingFlags.NonPublic | BindingFlags.Static)!;
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        foreach (string? style in new string?[] { null, "marker", "line", "smooth", "none", "bogus" })
        {
            string source = Chart(style, null);
            var scene = PptxTests.BuildSingleChartScene(source)!;
            XElement xml = XDocument.Parse(source).Descendants(c + "scatterChart").Single();
            XElement mismatch = XDocument.Parse(Chart(style is "none" or "bogus" ? "lineMarker" : "bogus", null))
                .Descendants(c + "scatterChart").Single();
            foreach (var pair in new[] { (Plot: (PptxSceneChartPlot?)null, Xml: xml), (Plot: scene.Plots[0], Xml: mismatch) })
            {
                object options = read.Invoke(null, [pair.Plot, pair.Xml])!;
                TestAssert.Equal(style is not ("none" or "bogus"), (bool)options.GetType().GetProperty("ConnectLines")!.GetValue(options)!);
            }
            TestAssert.Equal(style ?? string.Empty, scene.Plots[0].ScatterStyle);
            if (style is null)
                TestAssert.Equal(PptxSceneChartScatterStyle.Unknown, scene.Plots[0].ScatterStyleKind);
        }
    }

    public static void PptxScatterConnectionsFollowSeriesSmoothingAndLineVisibility()
    {
        foreach (string? smooth in new string?[] { null, "0", "1" })
        {
            string expected = Render(Chart("lineMarker", smooth));
            foreach (string? style in new string?[] { null, "marker", "line", "lineMarker", "smooth", "smoothMarker" })
            {
                string actual = Render(Chart(style, smooth));
                TestAssert.Equal(expected, actual);
                TestAssert.Equal(1, Connections(actual));
                TestAssert.Equal(0, Connections(Render(Chart(style, smooth, hidden: true))));
            }
        }
        foreach (string style in new[] { "none", "bogus" })
            TestAssert.Equal(0, Connections(Render(Chart(style, null))));

        static int Connections(string content) => Regex.Matches(content,
            @"(?m)^[^\r\n]+ m\r?\n(?:[^\r\n]+ [lc]\r?\n){2,}S(?:\r?\n|$)").Count;
    }

    private static string Render(string source)
    {
        using var stream = PptxTests.CreateSingleChartPackageStream(source);
        using var package = OoxPackage.Open(stream, default);
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
        return page.Content;
    }

    private static string Chart(string? style, string? smooth, bool hidden = false) => $$"""
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
          <c:scatterChart>{{(style is null ? "" : $"<c:scatterStyle val=\"{style}\"/>")}}<c:ser><c:idx val="0"/><c:order val="0"/>
            {{(hidden ? "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>" : "")}}
            <c:marker><c:symbol val="none"/></c:marker>
            <c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:xVal>
            <c:yVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numLit></c:yVal>
            {{(smooth is null ? "" : $"<c:smooth val=\"{smooth}\"/>")}}
          </c:ser><c:axId val="1"/><c:axId val="2"/></c:scatterChart>
          <c:valAx><c:axId val="1"/><c:scaling><c:min val="0"/><c:max val="4"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/><c:majorUnit val="1"/></c:valAx>
          <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="8"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/><c:majorUnit val="2"/></c:valAx>
        </c:plotArea></c:chart></c:chartSpace>
        """;
}
