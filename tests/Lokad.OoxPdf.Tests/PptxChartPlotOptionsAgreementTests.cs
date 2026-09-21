using System.IO.Compression;
using System.Reflection;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

// D01 branch-equivalence evidence: plot/chart option readers must agree whether they
// interpret the scene model or re-parse the plot XML. The scene builder and the render
// XML arms share parsers, but defaults and clamps are applied at different layers, so
// absent/unknown/out-of-range spellings are the divergence surface. Any failure here is
// a dual-interpretation bug and blocks deleting the XML arms.
internal static class PptxChartPlotOptionsAgreementTests
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public static void BarScalarOptionsAgreeBetweenSceneAndXml()
    {
        // grouping x barDir x varyColors x gapWidth x overlap, incl. absent/unknown/out-of-range.
        string[] groupings = new[] { "<c:grouping val=\"clustered\"/>", "<c:grouping val=\"stacked\"/>", "<c:grouping val=\"bogus\"/>", "" };
        string[] barDirs = new[] { "<c:barDir val=\"bar\"/>", "<c:barDir val=\"col\"/>", "<c:barDir val=\"bogus\"/>", "" };
        string[] varyColors = new[] { "<c:varyColors/>", "<c:varyColors val=\"0\"/>", "<c:varyColors val=\"1\"/>", "" };
        string[] gapWidths = new[] { "<c:gapWidth val=\"150\"/>", "<c:gapWidth val=\"999\"/>", "<c:gapWidth val=\"-5\"/>", "" };
        string[] overlaps = new[] { "<c:overlap val=\"0\"/>", "<c:overlap val=\"250\"/>", "<c:overlap val=\"-150\"/>", "" };
        foreach (string grouping in groupings)
        {
            foreach (string barDir in barDirs)
            {
                foreach (string vary in varyColors)
                {
                    foreach (string gap in gapWidths)
                    {
                        foreach (string overlap in overlaps)
                        {
                            AssertBarOptionsAgree(grouping + barDir + vary + gap + overlap);
                        }
                    }
                }
            }
        }
    }

    public static void ScatterStyleAgreesBetweenSceneAndXml()
    {
        string[] styles = new[] { "<c:scatterStyle val=\"marker\"/>", "<c:scatterStyle val=\"smoothMarker\"/>", "<c:scatterStyle val=\"bogus\"/>", "" };
        foreach (string style in styles)
        {
            (object? plot, XElement element) = LoadPlot("scatterChart", style);
            object? xml = Invoke("ReadSceneOrXmlChartScatterStyle", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { null, element });
            object? scene = Invoke("ReadSceneOrXmlChartScatterStyle", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { plot, element });
            TestAssert.True(Equals(xml, scene), "Scatter style must agree for plot XML: " + style);
        }
    }

    public static void RadarStyleAgreesBetweenSceneAndXml()
    {
        string[] styles = new[] { "<c:radarStyle val=\"standard\"/>", "<c:radarStyle val=\"marker\"/>", "<c:radarStyle val=\"filled\"/>", "<c:radarStyle val=\"bogus\"/>", "" };
        foreach (string style in styles)
        {
            (object? plot, XElement element) = LoadPlot("radarChart", style);
            object? xml = Invoke("ReadSceneOrXmlChartRadarStyle", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { null, element });
            object? scene = Invoke("ReadSceneOrXmlChartRadarStyle", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { plot, element });
            TestAssert.True(Equals(xml, scene), "Radar style must agree for plot XML: " + style);
        }
    }

    public static void ChartLevelOptionsAgreeBetweenSceneAndXml()
    {
        string[] blanks = new[] { "<c:dispBlanksAs val=\"span\"/>", "<c:dispBlanksAs val=\"zero\"/>", "<c:dispBlanksAs val=\"bogus\"/>", "" };
        string[] visible = new[] { "<c:plotVisOnly val=\"0\"/>", "<c:plotVisOnly val=\"1\"/>", "" };
        foreach (string blank in blanks)
        {
            foreach (string vis in visible)
            {
                string xml = ChartSpace("<c:barChart><c:ser><c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:barChart>", blank + vis);
                PptxSceneChart? sceneChart = PptxTests.BuildSingleChartScene(xml);
                TestAssert.NotNull(sceneChart);
                XDocument xmlDoc = XDocument.Parse(xml);
                object? xmlBlanks = Invoke("ReadSceneOrXmlChartDisplayBlanksAs", new[] { typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { null, xmlDoc });
                object? sceneBlanks = Invoke("ReadSceneOrXmlChartDisplayBlanksAs", new[] { typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { sceneChart, xmlDoc });
                TestAssert.True(Equals(xmlBlanks, sceneBlanks), "DisplayBlanksAs must agree for: " + blank + vis);
                object? xmlVisible = Invoke("ReadSceneOrXmlChartPlotVisibleOnly", new[] { typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { null, xmlDoc });
                object? sceneVisible = Invoke("ReadSceneOrXmlChartPlotVisibleOnly", new[] { typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { sceneChart, xmlDoc });
                TestAssert.True(Equals(xmlVisible, sceneVisible), "PlotVisibleOnly must agree for: " + blank + vis);
            }
        }
    }

    private static void AssertBarOptionsAgree(string plotInner)
    {
        (object? plot, XElement element) = LoadPlot("barChart", plotInner);
        object? xmlGrouping = Invoke("ReadSceneOrXmlChartGrouping", new[] { typeof(PptxSceneChartPlot), typeof(XElement), typeof(PptxSceneChartGrouping) }, new object?[] { null, element, PptxSceneChartGrouping.Standard });
        object? sceneGrouping = Invoke("ReadSceneOrXmlChartGrouping", new[] { typeof(PptxSceneChartPlot), typeof(XElement), typeof(PptxSceneChartGrouping) }, new object?[] { plot, element, PptxSceneChartGrouping.Standard });
        TestAssert.True(Equals(xmlGrouping, sceneGrouping), "Grouping must agree for plot XML: " + plotInner);
        object? xmlDir = Invoke("ReadSceneOrXmlChartBarDirection", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { null, element });
        object? sceneDir = Invoke("ReadSceneOrXmlChartBarDirection", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { plot, element });
        TestAssert.True(Equals(xmlDir, sceneDir), "BarDirection must agree for plot XML: " + plotInner);
        object? xmlOptions = Invoke("ReadSceneOrXmlChartBarOptions", new[] { typeof(PptxSceneChartPlot), typeof(XElement), typeof(PptxSceneChartGrouping) }, new object?[] { null, element, PptxSceneChartGrouping.Standard });
        object? sceneOptions = Invoke("ReadSceneOrXmlChartBarOptions", new[] { typeof(PptxSceneChartPlot), typeof(XElement), typeof(PptxSceneChartGrouping) }, new object?[] { plot, element, PptxSceneChartGrouping.Standard });
        TestAssert.True(Equals(xmlOptions, sceneOptions), "Bar options record must agree for plot XML: " + plotInner);
    }

    public static void ScenePlotsCorrespondToXmlPlotElementsAcrossFixtures()
    {
        // D01 pairing proof: every plotArea *Chart child must have exactly one scene
        // plot with the same element (reference identity), name order, and per-name
        // index. Renderers pair scene plots with XML elements by (kind, index), so any
        // gap here is a latent mis-pairing between the two interpretations.
        string cases = Path.Combine(Directory.GetCurrentDirectory(), "tests", "Lokad.OoxPdf.Tests", "Cases");
        int charts = 0;
        int plots = 0;
        foreach (string fixture in Directory.GetFiles(cases, "*.pptx"))
        {
            if (!ZipHasChartPart(fixture))
            {
                continue;
            }

            using FileStream stream = File.OpenRead(fixture);
            OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
            PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
            PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
            foreach (PptxSceneSlide slide in scene.Slides)
            {
                foreach (PptxSceneNode node in AllNodes(slide))
                {
                    if (node.Chart is not { ChartXml: { } xml } chart)
                    {
                        continue;
                    }

                    charts++;
                    List<XElement> elements = xml.Descendants(C + "plotArea").FirstOrDefault()?
                        .Elements()
                        .Where(e => e.Name.Namespace == C && e.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal))
                        .ToList() ?? new List<XElement>();
                    TestAssert.Equal(elements.Count, chart.Plots.Count);
                    for (int i = 0; i < elements.Count; i++)
                    {
                        TestAssert.Equal(elements[i].Name.LocalName, chart.Plots[i].Kind);
                        TestAssert.True(ReferenceEquals(elements[i], chart.Plots[i].Source), "Scene plot must wrap the plotArea child element itself.");
                    }

                    var seen = new Dictionary<string, int>(StringComparer.Ordinal);
                    foreach (PptxSceneChartPlot plot in chart.Plots)
                    {
                        int expected = seen.TryGetValue(plot.Kind, out int n) ? n : 0;
                        seen[plot.Kind] = expected + 1;
                        TestAssert.Equal(expected, plot.KindIndex);
                    }

                    plots += elements.Count;
                }
            }
        }

        TestAssert.True(charts > 0 && plots > 0, "Chart corpus must not be empty.");
    }

    private static bool ZipHasChartPart(string path)
    {
        using ZipArchive archive = ZipFile.OpenRead(path);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            if (entry.FullName.Contains("/charts/", StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<PptxSceneNode> AllNodes(PptxSceneSlide slide)
    {
        foreach (PptxSceneNode node in slide.MasterNodes.Concat(slide.LayoutNodes).Concat(slide.SlideNodes))
        {
            foreach (PptxSceneNode nested in Walk(node))
            {
                yield return nested;
            }
        }
    }

    private static IEnumerable<PptxSceneNode> Walk(PptxSceneNode node)
    {
        yield return node;
        foreach (PptxSceneNode child in node.Children)
        {
            foreach (PptxSceneNode nested in Walk(child))
            {
                yield return nested;
            }
        }
    }

    private static (object? Plot, XElement Element) LoadPlot(string plotKind, string plotInner)
    {
        string xml = ChartSpace("<c:" + plotKind + ">" + plotInner + "<c:ser><c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:" + plotKind + ">", "");
        PptxSceneChart? sceneChart = PptxTests.BuildSingleChartScene(xml);
        PptxSceneChart chart = TestAssert.NotNull(sceneChart);
        TestAssert.Equal(1, chart.Plots.Count);
        XElement element = XDocument.Parse(xml).Descendants(C + plotKind).First();
        return (chart.Plots[0], element);
    }

    private static string ChartSpace(string plotAreaInner, string options)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
            + "<c:chart><c:plotArea>" + plotAreaInner + "</c:plotArea>"
            + options
            + "</c:chart></c:chartSpace>";
    }

    private static object? Invoke(string name, Type[] types, object?[] args)
    {
        MethodInfo method = typeof(PptxRenderer).GetMethod(
            name,
            BindingFlags.NonPublic | BindingFlags.Static,
            null,
            types,
            null) ?? throw new InvalidOperationException("Expected renderer bridge: " + name);
        try
        {
            return method.Invoke(null, args);
        }
        catch (TargetInvocationException ex)
        {
            throw ex.InnerException ?? ex;
        }
    }
}
