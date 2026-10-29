using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartNegativeInversionTests
{
    public static void PptxNegativeBarScenePreservesSeriesAndPointInversion()
    {
        foreach ((string declaration, bool expected) in new[]
        {
            ("", true), ("<c:invertIfNegative/>", true),
            ("<c:invertIfNegative val=\"1\"/>", true),
            ("<c:invertIfNegative val=\"true\"/>", true),
            ("<c:invertIfNegative val=\"0\"/>", false),
            ("<c:invertIfNegative val=\"false\"/>", false)
        })
        {
            PptxSceneChart chart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(Chart(declaration)));
            var policy = chart.Plots[0].Series[0].NegativeInversion;
            TestAssert.Equal(expected, policy.ForPoint(0));
            TestAssert.Equal(expected, policy.ForPoint(7));
        }
        string overrides = """
            <c:dPt><c:idx val="0"/><c:invertIfNegative val="0"/></c:dPt>
            <c:dPt><c:idx val="1"/><c:invertIfNegative val="0"/></c:dPt>
            <c:dPt><c:idx val="1"/><c:invertIfNegative/></c:dPt>
            <c:dPt><c:idx val="-1"/><c:invertIfNegative val="0"/></c:dPt>
            """;
        var series = TestAssert.NotNull(PptxTests.BuildSingleChartScene(Chart("<c:invertIfNegative val=\"1\"/>", overrides))).Plots[0].Series[0];
        TestAssert.Equal(false, series.NegativeInversion.ForPoint(0));
        TestAssert.Equal(true, series.NegativeInversion.ForPoint(1));
        TestAssert.Equal(true, series.NegativeInversion.ForPoint(2));
        TestAssert.Equal(2, TestAssert.NotNull(series.NegativeInversion.PointValues).Count);
        var pointDefault = TestAssert.NotNull(PptxTests.BuildSingleChartScene(Chart(
            "<c:invertIfNegative val=\"0\"/>", "<c:dPt><c:idx val=\"0\"/></c:dPt>"))).Plots[0].Series[0];
        TestAssert.Equal(true, pointDefault.NegativeInversion.ForPoint(0));
        TestAssert.Equal(false, pointDefault.NegativeInversion.ForPoint(1));
    }

    public static void PptxNegativeBarPaintHonorsInversionAcrossLayouts()
    {
        foreach (string direction in new[] { "col", "bar" })
        foreach (string grouping in new[] { "clustered", "stacked" })
        foreach ((string declaration, bool blue) in new[]
        {
            ("<c:invertIfNegative val=\"0\"/>", true),
            ("<c:invertIfNegative val=\"1\"/>", false),
            ("", false)
        })
        {
            PdfPage page = Render(Chart(declaration, direction: direction, grouping: grouping));
            TestAssert.Equal(blue ? 1 : 0, CountSeriesColorPaint(page.Content));
        }
    }

    public static void PptxNegativeBarPointOverridesAlsoControlLabelSwatches()
    {
        foreach ((bool series, bool point) in new[] { (true, false), (false, true), (false, false), (true, true) })
        {
            string declaration = "<c:invertIfNegative val=\"" + (series ? "1" : "0") + "\"/>";
            string pointStyle = "<c:dPt><c:idx val=\"0\"/><c:invertIfNegative val=\"" + (point ? "1" : "0") + "\"/></c:dPt>";
            PdfPage page = Render(Chart(declaration, pointStyle, labels: true));
            TestAssert.Equal(point ? 0 : 2, CountSeriesColorPaint(page.Content));
        }
    }

    public static void PptxNegativeBarPointDefaultsAndExplicitPaintMatchOffice()
    {
        const string fill = "<c:spPr><a:solidFill><a:srgbClr val=\"3366CC\"/></a:solidFill></c:spPr>";
        foreach ((string flag, string paint, bool negative, int expected) in new[]
        {
            ("", "", true, 0),
            ("", fill, true, 0),
            ("<c:invertIfNegative val=\"0\"/>", fill, true, 1),
            ("<c:invertIfNegative val=\"1\"/>", fill, true, 0),
            ("", fill, false, 1)
        })
        {
            string point = "<c:dPt><c:idx val=\"0\"/>" + flag + paint + "</c:dPt>";
            PdfPage page = Render(Chart("<c:invertIfNegative val=\"0\"/>", point, negative: negative));
            TestAssert.Equal(expected, CountSeriesColorPaint(page.Content));
        }
    }

    private static string Chart(string inversion, string points = "", string direction = "col", string grouping = "clustered", bool labels = false, bool negative = true) => $$"""
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart><c:plotArea><c:barChart>
            <c:barDir val="{{direction}}"/><c:grouping val="{{grouping}}"/><c:varyColors val="0"/>
            <c:ser><c:idx val="0"/><c:order val="0"/>
              <c:spPr><a:solidFill><a:srgbClr val="3366CC"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>
              {{inversion}}{{points}}
              <c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Negative</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>{{(negative ? "-2" : "2")}}</c:v></c:pt></c:numLit></c:val>
              {{(labels ? "<c:dLbls><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"1\"/><c:showVal val=\"1\"/></c:dLbls>" : "")}}
            </c:ser>
          </c:barChart></c:plotArea><c:legend><c:delete val="1"/></c:legend></c:chart>
        </c:chartSpace>
        """;

    private static PdfPage Render(string chart)
    {
        var (document, package) = PptxTests.BuildSingleChartPackage(chart);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPages(document, package, null, default).Single();
    }

    // Assert painted source color, independent of rectangle/path grouping and IDs.
    private static int CountSeriesColorPaint(string content)
    {
        string color = "";
        var saved = new Stack<string>();
        int count = 0;
        foreach (string raw in content.Split('\n'))
        {
            string line = raw.Trim();
            if (line == "q") { saved.Push(color); }
            else if (line == "Q") { color = saved.Pop(); }
            else if (line.EndsWith(" rg", StringComparison.Ordinal) || line.EndsWith(" g", StringComparison.Ordinal)) { color = line; }
            else if (color == "0.2 0.4 0.8 rg" &&
                (line is "f" or "f*" || line.EndsWith(" f", StringComparison.Ordinal) || line.EndsWith(" f*", StringComparison.Ordinal)))
            { count++; }
        }
        return count;
    }
}
