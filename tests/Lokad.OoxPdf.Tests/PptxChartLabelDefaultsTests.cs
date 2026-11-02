using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartLabelDefaultsTests
{
    private static readonly (string Element, string Property)[] Components =
    [
        ("showVal", "ShowValue"), ("showPercent", "ShowPercent"),
        ("showCatName", "ShowCategoryName"), ("showSerName", "ShowSeriesName"),
        ("showLegendKey", "ShowLegendKey"), ("showBubbleSize", "ShowBubbleSize")
    ];

    public static void PptxDefinedLabelsUseIndependentOfficeComponentDefaults()
    {
        // Native controls cover line, column, area, scatter, pie/doughnut and
        // bubble. Missing containers stay distinct from missing component flags.
        foreach (bool defined in new[] { false, true })
        {
            var (chart, plot) = Source(defined ? "<c:dLbls/>" : "");
            foreach (object options in ReadBoth(chart, plot))
            {
                foreach (var component in Components)
                    TestAssert.Equal(defined, Flag(options, component.Property));
                TestAssert.True(!Flag(options, "ShowLeaderLines"), "Unproven leader-line fallback remains unchanged.");
            }
            TestAssert.Equal(null, chart.Plots[0].DataLabels.ShowSeriesName);
            TestAssert.Equal(defined, chart.Plots[0].DataLabels.IsDefined);
        }
        foreach (var component in Components)
        foreach (string value in new[] { "0", "false", "1", "" })
        {
            string attribute = value.Length == 0 ? "" : $" val=\"{value}\"";
            var (chart, plot) = Source($"<c:dLbls><c:{component.Element}{attribute}/></c:dLbls>");
            foreach (object options in ReadBoth(chart, plot))
            {
                foreach (var other in Components)
                    TestAssert.Equal(other == component ? value is not ("0" or "false") : true, Flag(options, other.Property));
                object raw = PptxTests.ChartDataLabelFlagOption(options, component.Element);
                TestAssert.True(PptxTests.ChartBooleanOptionIsDefined(raw), "Explicit flags retain source presence.");
                TestAssert.Equal(value, PptxTests.ChartBooleanOptionRawValue(raw));
            }
        }
    }

    public static void PptxLabelDefaultsFollowPlotSeriesAndPointPrecedence()
    {
        string disabled = "<c:dLbls>" + string.Concat(Components.Select(c => $"<c:{c.Element} val=\"0\"/>")) + "</c:dLbls>";
        const string series = """
            <c:dLbls><c:dLbl><c:idx val="0"/><c:showSerName val="1"/></c:dLbl>
              <c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl><c:showVal val="1"/></c:dLbls>
            """;
        foreach (string plotLabels in new[] { "", disabled })
        foreach (string seriesLabels in new[] { "<c:dLbls/>", series })
        {
            var (chart, plot) = Source(plotLabels, seriesLabels);
            foreach (bool scene in new[] { false, true })
            {
                object options = Read(chart, plot, scene);
                object perSeries = Invoke("ReadSceneOrXmlSeriesDataLabelOptions", scene ? chart : null,
                    scene ? chart.Plots[0] : null, plot, PptxTheme.Empty, PptxColorMap.Default);
                object merged = Invoke("ResolveChartDataLabelOptionsForSeries", options, perSeries, 0);
                foreach (var component in Components)
                    TestAssert.Equal(plotLabels.Length == 0 || (seriesLabels == series && component.Property == "ShowValue"), Flag(merged, component.Property));
                if (seriesLabels == series)
                {
                    object point = Invoke("ResolveChartDataLabelOptions", merged, 0);
                    TestAssert.True(Flag(point, "ShowSeriesName"), "The point's explicit true overrides an inherited false.");
                    object deleted = Invoke("ResolveChartDataLabelOptions", merged, 1);
                    TestAssert.True(!Flag(deleted, "HasVisibleContent"), "Deleted labels remain absent.");
                }
            }
        }
    }

    public static void PptxLineLabelDefaultsPreserveCompleteSourceTextWithoutKeys()
    {
        foreach (bool keys in new[] { false, true })
        {
            const string name = "A longer source series";
            string labels = $"<c:dLbls><c:showLegendKey val=\"{(keys ? 1 : 0)}\"/></c:dLbls>";
            var (_, plot) = Source(labels);
            plot.Name = C + "lineChart";
            plot.Descendants(C + "tx").First().Element(C + "v")!.Value = name;
            var xml = new XDocument(plot.Document!);
            // Use the authored 720pt frame from the independent origin portfolio.
            const string slide = """
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="457200"/><a:ext cx="9144000" cy="5486400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
                """;
            using var stream = PptxTests.CreateSingleChartPackageStream(xml.ToString(), new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
            using var package = OoxPackage.Open(stream, default);
            var document = new PptxReader().Read(package, default);
            var page = new PptxRenderer(new TestFaceFontResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(show => show.Text));
            TestAssert.Equal($"{name}, W1, 2{name}, W2, 4{name}, W3, 3", text);
        }
    }

    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    private static (PptxSceneChart Chart, XElement Plot) Source(string plotLabels, string seriesLabels = "")
    {
        var xml = XDocument.Parse($$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="800"/></a:pPr></a:p></c:txPr><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Actual</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>W1</c:v></c:pt><c:pt idx="1"><c:v>W2</c:v></c:pt><c:pt idx="2"><c:v>W3</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>{{seriesLabels}}</c:ser>{{plotLabels}}</c:barChart></c:plotArea></c:chart></c:chartSpace>
            """);
        return (PptxTests.BuildSingleChartScene(xml.ToString())!, xml.Descendants(C + "barChart").Single());
    }

    private static IEnumerable<object> ReadBoth(PptxSceneChart chart, XElement plot)
        => new[] { Read(chart, plot, false), Read(chart, plot, true) };

    private static object Read(PptxSceneChart chart, XElement plot, bool scene)
        => Invoke("ReadSceneOrXmlDataLabelOptions", scene ? chart : null, scene ? chart.Plots[0] : null,
            plot, PptxTheme.Empty, PptxColorMap.Default);

    private static object Invoke(string name, params object?[] arguments)
        => typeof(PptxRenderer).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, arguments)!;

    private static bool Flag(object value, string name) => (bool)value.GetType().GetProperty(name)!.GetValue(value)!;
}
