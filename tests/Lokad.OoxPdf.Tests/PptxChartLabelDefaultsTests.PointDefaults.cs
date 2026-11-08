using System.Xml.Linq;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxActivePlotPointsOwnDefaultsWithoutChangingRawFlags()
    {
        string fourOff = string.Concat(Components.Where(c => c.Property is not ("ShowPercent" or "ShowBubbleSize"))
            .Select(c => $"<c:{c.Element} val=\"0\"/>"));
        foreach (string parent in new[] { "<c:showVal val=\"0\"/>", fourOff,
            fourOff + "<c:showPercent val=\"0\"/>", fourOff + "<c:showBubbleSize val=\"0\"/>" })
        {
            var (chart, plot) = PointDefaultsSource("lineChart", parent, "");
            TestAssert.Equal(null, chart.Plots[0].DataLabels.Overrides[0].ShowValue);
            object options = Read(chart);
            object point = Invoke("ResolveChartDataLabelOptions", options, 0);
            foreach (var component in Components)
            {
                TestAssert.True(Flag(point, component.Property), "An admitted point owns its absent-component defaults.");
                object raw = PptxTests.ChartDataLabelFlagOption(point, component.Element);
                TestAssert.True(!PptxTests.ChartBooleanOptionIsDefined(raw), "Default resolution preserves missing source flags.");
            }
        }
        foreach (var component in Components)
        foreach (string value in new[] { "0", "false", "1", "" })
        {
            string attribute = value.Length == 0 ? "" : $" val=\"{value}\"";
            var (chart, plot) = PointDefaultsSource("lineChart", "<c:showVal val=\"0\"/>", $"<c:{component.Element}{attribute}/>");
            object options = Read(chart);
            object point = Invoke("ResolveChartDataLabelOptions", options, 0);
            foreach (var other in Components)
                TestAssert.Equal(other == component ? value is not ("0" or "false") : true, Flag(point, other.Property));
            object raw = PptxTests.ChartDataLabelFlagOption(point, component.Element);
            TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(raw), "An explicit point flag keeps its own source presence.");
            TestAssert.Equal(value, PptxTests.ChartBooleanOptionRawValue(raw));
        }
    }

    public static void PptxDisabledCartesianPlotsSuppressPointRequestsButKeepSource()
    {
        foreach (string kind in new[] { "barChart", "lineChart", "areaChart" })
        foreach (string point in new[] { "", "<c:showVal val=\"1\"/>", PointDefaultsCustomText })
        {
            var (chart, plot) = PointDefaultsSource(kind, PointDefaultsDisabledFlags, point);
            TestAssert.Equal(1, chart.Plots[0].DataLabels.Overrides.Count);
            object options = Read(chart);
            TestAssert.True(!Flag(options, "HasVisibleContent"), "A fully disabled Cartesian plot does not admit its point requests.");
            TestAssert.True(!Flag(Invoke("ResolveChartDataLabelOptions", options, 0), "HasVisibleContent"),
                "The ignored point does not re-enable the resolved rendering options.");
        }
    }

    public static void PptxSectorPointsRemainIndependentOfPlotDisables()
    {
        foreach (string kind in new[] { "pieChart", "doughnutChart" })
        {
            var (chart, plot) = PointDefaultsSource(kind, PointDefaultsDisabledFlags, "<c:showVal val=\"0\"/>");
            object options = Read(chart);
            TestAssert.True(Flag(options, "HasVisibleContent"), "Sector points have independent label scope.");
            object point = Invoke("ResolveChartDataLabelOptions", options, 0);
            foreach (var component in Components)
                TestAssert.Equal(component.Property != "ShowValue", Flag(point, component.Property));
        }
    }

    public static void PptxDefinedSeriesOwnPointDeclarationsWithoutLeakingPlotPoints()
    {
        string parent = "<c:dLbls><c:dLbl><c:idx val=\"0\"/><c:delete val=\"1\"/></c:dLbl>" +
            "<c:dLbl><c:idx val=\"2\"/><c:showVal val=\"0\"/></c:dLbl><c:showVal val=\"0\"/></c:dLbls>";
        foreach (string series in new[] { "<c:dLbls/>", "<c:dLbls>" + PointDefaultsDisabledFlags + "</c:dLbls>",
            "<c:dLbls><c:dLbl><c:idx val=\"1\"/><c:showVal val=\"0\"/></c:dLbl>" + PointDefaultsDisabledFlags + "</c:dLbls>" })
        {
            var (chart, plot) = Source(parent, series);
            object options = Read(chart);
            object perSeries = Invoke("ReadSceneSeriesDataLabelOptions", chart,
                chart.Plots[0]);
            object merged = Invoke("ResolveChartDataLabelOptionsForSeries", options, perSeries, 0);
            if (series == "<c:dLbls/>")
            {
                TestAssert.True(Flag(Invoke("ResolveChartDataLabelOptions", merged, 0), "HasVisibleContent"),
                    "A defined empty series does not inherit the plot's point deletion.");
                TestAssert.True(Flag(Invoke("ResolveChartDataLabelOptions", merged, 2), "ShowValue"),
                    "The plot point's explicit false does not leak into the new series scope.");
            }
            else
            {
                TestAssert.True(!Flag(Invoke("ResolveChartDataLabelOptions", merged, 2), "HasVisibleContent"),
                    "A disabled series does not inherit another plot point's default components.");
                if (series.Contains("<c:dLbl>", StringComparison.Ordinal))
                {
                    object point = Invoke("ResolveChartDataLabelOptions", merged, 1);
                    TestAssert.True(!Flag(point, "ShowValue") && Flag(point, "ShowCategoryName"),
                        "The series point keeps explicit false and its own other defaults.");
                }
            }
        }
        foreach (string point in new[] { PointDefaultsCustomText,
            "<c:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></c:spPr>" })
        {
            XDocument xml = PointDefaultsRenderSource(false, point);
            var original = RenderChartPage(xml);
            if (point == PointDefaultsCustomText)
                TestAssert.Equal("Pinned", string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(original, false,
                    fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text)));
            else
                TestAssert.Equal(1, MarkerRectangles(original, new RgbColor(0, 255, 0)).Length);
            xml.Descendants(C + "ser").Single().Element(C + "cat")!.AddBeforeSelf(new XElement(C + "dLbls"));
            var page = RenderChartPage(xml);
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text));
            TestAssert.Equal("Actual, W1, 2Actual, W2, 4Actual, W3, 3", text);
            TestAssert.Equal(0, MarkerRectangles(page, new RgbColor(0, 255, 0)).Length);
            TestAssert.Equal(2, MarkerRectangles(page, new RgbColor(255, 0, 0)).Length);
        }
    }

    public static void PptxPlotPointDefaultRenderingRestoresCompleteContentAndKey()
    {
        var page = RenderChartPage(PointDefaultsRenderSource(false, ""));
        string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
            fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text));
        TestAssert.Equal("Actual, W3, 3", text);
        TestAssert.Equal(2, MarkerRectangles(page, new RgbColor(255, 0, 0)).Length);
    }

    public static void PptxDisabledPlotDefaultRenderingOmitsPointContentAndKeys()
    {
        foreach (string point in new[] { "<c:showVal val=\"1\"/>", PointDefaultsCustomText })
        {
            var page = RenderChartPage(PointDefaultsRenderSource(true, point));
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text));
            TestAssert.Equal("", text);
            TestAssert.Equal(1, MarkerRectangles(page, new RgbColor(255, 0, 0)).Length);
        }
    }

    private static string PointDefaultsDisabledFlags => string.Concat(Components.Select(c => $"<c:{c.Element} val=\"0\"/>"));
    private const string PointDefaultsCustomText = "<c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Pinned</a:t></a:r></a:p></c:rich></c:tx>";

    private static (PptxSceneChart Chart, XElement Plot) PointDefaultsSource(string kind, string parentFlags, string point)
    {
        var (_, plot) = Source("<c:dLbls><c:dLbl><c:idx val=\"0\"/>" + point + "</c:dLbl>" + parentFlags + "</c:dLbls>");
        plot.Name = C + kind;
        return (PptxTests.BuildSingleChartScene(plot.Document!.ToString())!, plot);
    }

    private static XDocument PointDefaultsRenderSource(bool disabled, string point)
    {
        XDocument xml = MarkerSource([PointMarker(2, RedMarker)]);
        string flags = disabled ? PointDefaultsDisabledFlags : string.Concat(Components
            .Where(c => c.Property is not ("ShowPercent" or "ShowBubbleSize"))
            .Select(c => $"<c:{c.Element} val=\"0\"/>"));
        XElement labels = XElement.Parse("<c:dLbls xmlns:c=\"" + C + "\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<c:dLbl><c:idx val=\"0\"/><c:delete val=\"1\"/></c:dLbl>" +
            "<c:dLbl><c:idx val=\"1\"/><c:delete val=\"1\"/></c:dLbl>" +
            "<c:dLbl><c:idx val=\"2\"/>" + point + "</c:dLbl>" + flags + "</c:dLbls>");
        xml.Descendants(C + "dLbls").Single().ReplaceWith(labels);
        return xml;
    }
}
