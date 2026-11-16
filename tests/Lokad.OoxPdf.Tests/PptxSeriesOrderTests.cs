using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxSeriesOrderTests
{
    private static readonly XNamespace C = OoxNamespaces.ChartNamespace;
    private static readonly string[] Names = ["Alpha", "Beta", "Gamma"];
    private static readonly int[] Ids = [41, 8, 99];
    private static readonly string[] Automatic = ["diamond", "square", "triangle"];
    private static readonly string[] Explicit = ["triangle", "diamond", "square"];
    private static readonly RgbColor[] Colors = [new(255, 0, 0), new(0, 255, 0), new(0, 0, 255)];

    public static void PptxSeriesOrderAutomaticDefaultsFollowResolvedOrder()
    {
        foreach (string kind in new[] { "lineChart", "scatterChart" })
        foreach (int[] permutation in new[] { new[] { 2, 1, 0 }, new[] { 1, 2, 0 } })
        {
            var chart = Chart(kind, explicitMarkers: false);
            Reorder(chart, permutation);
            var series = PptxTests.BuildSingleChartScene(chart.ToString())!.Plots.Single().Series;
            for (int index = 0; index < series.Count; index++)
            {
                CheckSource(series[index], index, kind);
                TestAssert.Equal(10 * (index + 1), series[index].Order!.Value);
                TestAssert.Equal(Automatic[index], series[index].Marker.Symbol);
                TestAssert.Equal(Automatic[index], series[index].PointStyles.Single().Marker!.Symbol);
            }
        }
    }

    public static void PptxSeriesOrderPermutationsKeepAutomaticPaintAndVisibleNames()
    {
        foreach (string kind in new[] { "lineChart", "scatterChart" })
        {
            XDocument chart = Chart(kind, explicitMarkers: false);
            PdfPage baseline = Render(chart);
            Reorder(chart, [2, 1, 0]);
            PdfPage permuted = Render(chart);
            TestAssert.True(baseline.Content == permuted.Content,
                "Equivalent source order must preserve marker/key paint: " + kind);
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(permuted).Select(show => show.Text));
            foreach (string name in Names) TestAssert.Contains(name, text);
            PdfContentValidator.ValidatePage(permuted, 0, default);
        }
    }

    public static void PptxSeriesOrderExplicitAndInvalidDeclarationsRetainOwners()
    {
        XDocument explicitChart = Chart("lineChart", explicitMarkers: true);
        Reorder(explicitChart, [2, 1, 0]);
        var explicitSeries = PptxTests.BuildSingleChartScene(explicitChart.ToString())!.Plots.Single().Series;
        for (int index = 0; index < explicitSeries.Count; index++)
        {
            CheckSource(explicitSeries[index], index, "lineChart");
            TestAssert.Equal(Explicit[index], explicitSeries[index].Marker.Symbol);
            TestAssert.Equal(Explicit[index], explicitSeries[index].PointStyles.Single().Marker!.Symbol);
        }
        foreach (string? invalid in new string?[] { null, "-1", "20", "invalid", "2147483648" })
        {
            XDocument chart = Chart("lineChart", explicitMarkers: false);
            XElement order = chart.Descendants(C + "ser").First().Element(C + "order")!;
            if (invalid is null) order.Remove(); else order.SetAttributeValue("val", invalid);
            Reorder(chart, [2, 1, 0]);
            var series = PptxTests.BuildSingleChartScene(chart.ToString())!.Plots.Single().Series;
            for (int index = 0; index < series.Count; index++)
            {
                CheckSource(series[index], 2 - index, "lineChart");
                TestAssert.Equal(Automatic[index], series[index].Marker.Symbol);
            }
            TestAssert.Equal(invalid ?? "", series[2].OrderValue);
        }
    }

    private static void CheckSource(PptxSceneChartSeries series, int index, string kind)
    {
        TestAssert.Equal(Names[index], series.Name);
        TestAssert.Equal(Ids[index], series.Index!.Value);
        TestAssert.Equal(Colors[index], series.Fill.Color);
        TestAssert.Equal((double)index + 1, (kind == "scatterChart" ? series.YValues : series.Values).Single());
    }

    private static void Reorder(XDocument chart, int[] permutation)
    {
        var series = chart.Descendants(C + "ser").ToArray();
        var anchor = new XComment("Series source position");
        series[0].AddBeforeSelf(anchor);
        foreach (var item in series) item.Remove();
        anchor.AddBeforeSelf(permutation.Select(index => series[index]));
        anchor.Remove();
    }

    private static PdfPage Render(XDocument chart)
    {
        using var input = PptxTests.CreateSingleChartPackageStream(chart.ToString());
        using var package = OoxPackage.Open(input, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }

    private static XDocument Chart(string kind, bool explicitMarkers)
    {
        string Data(int index) => kind == "scatterChart"
            ? $"<c:xVal><c:numLit><c:pt idx=\"0\"><c:v>1</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx=\"0\"><c:v>{index + 1}</c:v></c:pt></c:numLit></c:yVal>"
            : $"<c:cat><c:strLit><c:pt idx=\"0\"><c:v>Category</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>{index + 1}</c:v></c:pt></c:numLit></c:val>";
        string series = string.Concat(Enumerable.Range(0, 3).Select(index => $$"""
            <c:ser><c:idx val="{{Ids[index]}}"/><c:order val="{{10 * (index + 1)}}"/><c:tx><c:v>{{Names[index]}}</c:v></c:tx>
            <c:spPr><a:solidFill><a:srgbClr val="{{new[] { "FF0000", "00FF00", "0000FF" }[index]}}"/></a:solidFill></c:spPr>
            {{(explicitMarkers ? "<c:marker><c:symbol val=\"" + Explicit[index] + "\"/><c:size val=\"7\"/></c:marker>" : "")}}
            <c:dPt><c:idx val="0"/><c:marker><c:size val="12"/></c:marker></c:dPt>
            {{Data(index)}}</c:ser>
            """));
        string xAxis = kind == "scatterChart" ? "valAx" : "catAx";
        return XDocument.Parse($$"""
            <c:chartSpace xmlns:c="{{C}}" xmlns:a="{{OoxNamespaces.DrawingNamespace}}"><c:chart><c:plotArea><c:{{kind}}>
            {{(kind == "scatterChart" ? "<c:scatterStyle val=\"marker\"/>" : "")}}
            {{series}}{{(kind == "lineChart" ? "<c:marker val=\"1\"/>" : "")}}<c:axId val="1"/><c:axId val="2"/></c:{{kind}}>
            <c:{{xAxis}}><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="2"/></c:{{xAxis}}>
            <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="4"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="1"/></c:valAx>
            </c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>
            """);
    }
}
