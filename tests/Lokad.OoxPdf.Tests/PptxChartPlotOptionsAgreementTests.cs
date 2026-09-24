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
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

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

    public static void SeriesSmoothAgreesBetweenSceneAndXml()
    {
        string[] smooths = new[] { "<c:smooth/>", "<c:smooth val=\"0\"/>", "<c:smooth val=\"1\"/>", "<c:smooth val=\"bogus\"/>", "" };
        foreach (string smooth in smooths)
        {
            (object? plot, XElement element) = LoadPlot("barChart", smooth);
            object? xml = Invoke("ReadSceneOrXmlSmoothSeries", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { null, element });
            object? scene = Invoke("ReadSceneOrXmlSmoothSeries", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { plot, element });
            TestAssert.True(SequenceEqual(xml, scene), "Smooth must agree for ser XML: " + smooth);
        }
    }

    public static void SeriesLineHiddenAgreesBetweenSceneAndXml()
    {
        string[] lineForms = new[] { "", "<c:spPr><a:ln w=\"12700\"/></c:spPr>", "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>" };
        foreach (string lineForm in lineForms)
        {
            (object? plot, XElement element) = LoadPlot("barChart", lineForm);
            object? xml = Invoke("ReadSceneOrXmlSeriesLineHidden", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { null, element });
            object? scene = Invoke("ReadSceneOrXmlSeriesLineHidden", new[] { typeof(PptxSceneChartPlot), typeof(XElement) }, new object?[] { plot, element });
            TestAssert.True(SequenceEqual(xml, scene), "LineHidden must agree for ser XML: " + lineForm);
        }
    }

    public static void SeriesNamesAgreeBetweenSceneAndXml()
    {
        string[] names = new[]
        {
            "",
            "<c:tx><c:v>Solo</c:v></c:tx>",
            "<c:tx><c:v>  Padded  </c:v></c:tx>",
            "<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>Cached</c:v></c:pt></c:strCache></c:strRef></c:tx>",
            "<c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx>",
        };
        Type workbookType = TestAssert.NotNull(typeof(PptxRenderer).Assembly.GetType("Lokad.OoxPdf.Pptx.PptxRenderer+ChartWorkbookData"));
        
        foreach (string name in names)
        {
            (object? plot, XElement element) = LoadPlot("barChart", name);
            object? xml = Invoke("ReadSceneOrXmlChartSeriesNameRecords", new[] { typeof(PptxSceneChartPlot), typeof(XElement), workbookType }, new object?[] { null, element, null });
            object? scene = Invoke("ReadSceneOrXmlChartSeriesNameRecords", new[] { typeof(PptxSceneChartPlot), typeof(XElement), workbookType }, new object?[] { plot, element, null });
            TestAssert.True(SeriesNamesEqual(xml, scene), "Series names must agree for ser XML: " + name);
        }
    }

    public static void SeriesExplosionsAgreeBetweenSceneAndXml()
    {
        string[] explosions = new[]
        {
            "",
            "<c:explosion val=\"25\"/>",
            "<c:explosion val=\"150\"/>",
            "<c:explosion val=\"bogus\"/>",
            "<c:dPt><c:idx val=\"0\"/><c:explosion val=\"50\"/></c:dPt>",
            "<c:dPt><c:idx val=\"bogus\"/><c:explosion val=\"50\"/></c:dPt>",
        };
        Type workbookType = TestAssert.NotNull(typeof(PptxRenderer).Assembly.GetType("Lokad.OoxPdf.Pptx.PptxRenderer+ChartWorkbookData"));
        
        foreach (string explosion in explosions)
        {
            (object? plot, XElement element) = LoadPlot("pieChart", explosion);
            object? xml = Invoke("ReadSceneOrXmlChartPointExplosions", new[] { typeof(PptxSceneChartPlot), typeof(XElement), workbookType }, new object?[] { null, element, null });
            object? scene = Invoke("ReadSceneOrXmlChartPointExplosions", new[] { typeof(PptxSceneChartPlot), typeof(XElement), workbookType }, new object?[] { plot, element, null });
            TestAssert.True(DictionariesEqual(xml, scene), "Explosions must agree for ser XML: " + explosion);
        }
    }

    public static void AxisOptionsAgreeBetweenSceneAndXml()
    {
        string[] scalings = new[] { "", "<c:scaling><c:orientation val=\"minMax\"/></c:scaling>", "<c:scaling><c:orientation val=\"maxMin\"/></c:scaling>", "<c:scaling><c:orientation val=\"bogus\"/></c:scaling>" };
        string[] gridlines = new[] { "", "<c:majorGridlines/>", "<c:minorGridlines/>", "<c:majorGridlines/><c:minorGridlines/>" };
        string[] units = new[] { "", "<c:majorUnit val=\"10\"/>", "<c:minorUnit val=\"5\"/>", "<c:majorUnit val=\"bogus\"/>" };
        string[] formats = new[] { "", "<c:numFmt formatCode=\"0.0\" sourceLinked=\"1\"/>", "<c:numFmt formatCode=\"General\" sourceLinked=\"0\"/>" };
        string[] ticks = new[] { "", "<c:majorTickMark val=\"out\"/>", "<c:majorTickMark val=\"bogus\"/>" };
        foreach (string axisInner in AllAxisForms(scalings, gridlines, units, formats, ticks))
        {
            (object? axis, XElement element) = LoadAxis(axisInner);
            AssertAxisReaderAgrees("ReadSceneOrXmlValueAxisReversed", axis, element, axisInner);
            AssertAxisReaderAgrees("ReadSceneOrXmlMajorGridlines", axis, element, axisInner);
            AssertAxisReaderAgrees("ReadSceneOrXmlMinorGridlines", axis, element, axisInner);
            AssertAxisReaderAgrees("ReadSceneOrXmlChartValueAxisUnits", axis, element, axisInner);
            AssertAxisReaderAgrees("ReadSceneOrXmlChartAxisMajorTickMark", axis, element, axisInner);
            AssertAxisReaderAgrees("ReadSceneOrXmlChartAxisNumberFormat", axis, element, axisInner);
        }
    }

    private static IEnumerable<string> AllAxisForms(string[] scalings, string[] gridlines, string[] units, string[] formats, string[] ticks)
    {
        foreach (string scaling in scalings)
        {
            yield return scaling;
        }

        foreach (string gridline in gridlines)
        {
            if (gridline.Length != 0)
            {
                yield return gridline;
            }
        }

        foreach (string unit in units)
        {
            if (unit.Length != 0)
            {
                yield return unit;
            }
        }

        foreach (string format in formats)
        {
            if (format.Length != 0)
            {
                yield return format;
            }
        }

        foreach (string tick in ticks)
        {
            if (tick.Length != 0)
            {
                yield return tick;
            }
        }
    }

    private static void AssertAxisReaderAgrees(string name, object? axis, XElement element, string axisInner)
    {
        object? xml = Invoke(name, new[] { typeof(PptxSceneChartAxis), typeof(XElement) }, new object?[] { null, element });
        object? scene = Invoke(name, new[] { typeof(PptxSceneChartAxis), typeof(XElement) }, new object?[] { axis, element });
        TestAssert.True(Equals(xml, scene), name + " must agree for axis XML: " + axisInner);
    }

    private static (object? Axis, XElement Element) LoadAxis(string axisInner)
    {
        string xml = ChartSpace("<c:barChart><c:ser><c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:barChart><c:valAx><c:axId val=\"10\"/>" + axisInner + "</c:valAx>", "");
        PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));

        TestAssert.Equal(1, sceneChart.Axes.Count);
        XElement element = XDocument.Parse(xml).Descendants(C + "valAx").First();
        return (sceneChart.Axes[0], element);
    }

    public static void TitleOptionsAgreeBetweenSceneAndXml()
    {
        // Shape-style arms carry nested reference payloads (list identity never
        // agrees across arms), so the battery pins text, body properties, text
        // style, and text runs; style equality is covered by renderer families.
        string[] titles = new[]
        {
            "",
            "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val=\"0\"/></c:title>",
            "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr b=\"1\" sz=\"1400\"/><a:t>Q1 </a:t></a:r><a:r><a:rPr i=\"1\"/><a:t>rev</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val=\"1\"/></c:title>",
            "<c:title><c:txPr><a:bodyPr rot=\"5400000\"/></c:txPr><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Spun</a:t></a:r></a:p></c:rich></c:tx><c:layout/></c:title>",
        };
        foreach (string title in titles)
        {
            string xml = ChartSpace("<c:barChart><c:ser><c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:barChart>", title);
            PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));

            XDocument xmlDoc = XDocument.Parse(xml);
            object? xmlText = Invoke("ReadSceneOrXmlChartTitleText", new[] { typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { null, xmlDoc });
            object? sceneText = Invoke("ReadSceneOrXmlChartTitleText", new[] { typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { sceneChart, xmlDoc });
            TestAssert.True(Equals(xmlText, sceneText), "Title text must agree for: " + title);
            object? xmlBody = Invoke("ReadSceneOrXmlChartTitleTextBodyProperties", new[] { typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { null, xmlDoc });
            object? sceneBody = Invoke("ReadSceneOrXmlChartTitleTextBodyProperties", new[] { typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { sceneChart, xmlDoc });
            TestAssert.True(Equals(xmlBody, sceneBody), "Title body properties must agree for: " + title);
            object? xmlStyle = Invoke("ReadSceneOrXmlChartTitleTextStyle", new[] { typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument), typeof(bool) }, new object?[] { PptxTheme.Empty, PptxColorMap.Default, null, xmlDoc, false });
            object? sceneStyle = Invoke("ReadSceneOrXmlChartTitleTextStyle", new[] { typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument), typeof(bool) }, new object?[] { PptxTheme.Empty, PptxColorMap.Default, sceneChart, xmlDoc, false });
            TestAssert.True(Equals(xmlStyle, sceneStyle), "Title text style must agree for: " + title);
            object? xmlRuns = Invoke("ReadSceneOrXmlChartTitleTextRuns", new[] { typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { PptxTheme.Empty, PptxColorMap.Default, null, xmlDoc });
            object? sceneRuns = Invoke("ReadSceneOrXmlChartTitleTextRuns", new[] { typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { PptxTheme.Empty, PptxColorMap.Default, sceneChart, xmlDoc });
            TestAssert.True(SequenceEqual(xmlRuns, sceneRuns), "Title text runs must agree for: " + title);
        }
    }

    public static void LegendOptionsAgreeBetweenSceneAndXml()
    {
        // Absent legends flow through the shared legend builder on both arms, so
        // the absent spelling must agree exactly (this pins the no-divergence claim).
        // Shape-style variants stay out: nested reference payloads never agree by
        // identity across arms (same documented scope as title shape style).
        string[] legends = new[]
        {
            "",
            "<c:legend><c:legendPos val=\"r\"/><c:layout/><c:overlay val=\"0\"/></c:legend>",
            "<c:legend><c:legendPos val=\"b\"/><c:layout/><c:overlay val=\"1\"/></c:legend>",
            "<c:legend><c:legendPos val=\"bogus\"/><c:layout/></c:legend>",
            "<c:legend><c:legendPos val=\"r\"/><c:layout/><c:overlay val=\"0\"/><c:delete val=\"1\"/></c:legend>",
        };
        foreach (string legend in legends)
        {
            string xml = ChartSpace("<c:barChart><c:ser><c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:barChart>", legend);
            PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));

            XDocument xmlDoc = XDocument.Parse(xml);
            object? xmlLayout = Invoke("ReadSceneOrXmlChartLegendLayout", new[] { typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { PptxTheme.Empty, PptxColorMap.Default, null, xmlDoc });
            object? sceneLayout = Invoke("ReadSceneOrXmlChartLegendLayout", new[] { typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { PptxTheme.Empty, PptxColorMap.Default, sceneChart, xmlDoc });
            TestAssert.True(Equals(xmlLayout, sceneLayout), "Legend layout must agree for: " + legend);
            object? xmlStyle = Invoke("ReadSceneOrXmlChartLegendTextStyle", new[] { typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { PptxTheme.Empty, PptxColorMap.Default, null, xmlDoc });
            object? sceneStyle = Invoke("ReadSceneOrXmlChartLegendTextStyle", new[] { typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument) }, new object?[] { PptxTheme.Empty, PptxColorMap.Default, sceneChart, xmlDoc });
            TestAssert.True(Equals(xmlStyle, sceneStyle), "Legend text style must agree for: " + legend);
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
                PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));

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

    private static bool SequenceEqual(object? xml, object? scene)
    {
        if (xml is System.Collections.IEnumerable xmlItems && scene is System.Collections.IEnumerable sceneItems)
        {
            var xmlList = xmlItems.Cast<object?>().ToList();
            var sceneList = sceneItems.Cast<object?>().ToList();
            return xmlList.Count == sceneList.Count && xmlList.SequenceEqual(sceneList);
        }

        return Equals(xml, scene);
    }

    private static bool DictionariesEqual(object? xml, object? scene)
    {
        if (xml is System.Collections.IDictionary xmlMap && scene is System.Collections.IDictionary sceneMap)
        {
            if (xmlMap.Count != sceneMap.Count)
            {
                return false;
            }

            foreach (object? key in xmlMap.Keys)
            {
                if (!sceneMap.Contains(key) || !Equals(xmlMap[key], sceneMap[key]))
                {
                    return false;
                }
            }

            return true;
        }

        return Equals(xml, scene);
    }

    private static bool SeriesNamesEqual(object? xml, object? scene)
    {
        if (xml is System.Collections.IEnumerable xmlItems && scene is System.Collections.IEnumerable sceneItems)
        {
            List<object?> xmlList = xmlItems.Cast<object?>().ToList();
            List<object?> sceneList = sceneItems.Cast<object?>().ToList();
            return xmlList.Count == sceneList.Count && xmlList.Zip(sceneList).All(pair => SeriesNameEqual(pair.First, pair.Second));
        }

        return Equals(xml, scene);
    }

    private static bool SeriesNameEqual(object? xml, object? scene)
    {
        if (xml is null || scene is null)
        {
            return xml is null && scene is null;
        }

        // WorkbookPoints is a fresh list per arm, so compare its contents rather
        // than the record (list identity would never agree across arms).
        return Prop(xml, "ActiveName") == Prop(scene, "ActiveName") &&
            Prop(xml, "CacheName") == Prop(scene, "CacheName") &&
            Prop(xml, "ActiveNameSource") == Prop(scene, "ActiveNameSource") &&
            PointsText(xml, "WorkbookPoints").SequenceEqual(PointsText(scene, "WorkbookPoints"));
    }

    private static string? Prop(object record, string name)
    {
        return record.GetType().GetProperty(name)?.GetValue(record)?.ToString();
    }

    private static List<string?> PointsText(object record, string name)
    {
        if (record.GetType().GetProperty(name)?.GetValue(record) is System.Collections.IEnumerable points)
        {
            return points.Cast<object?>().Select(point => point?.ToString()).ToList();
        }

        return new List<string?>();
    }

    private static (object? Plot, XElement Element) LoadPlot(string plotKind, string plotInner)
    {
        string xml = ChartSpace("<c:" + plotKind + ">" + plotInner + "<c:ser><c:cat><c:strLit><c:pt idx=\"0\"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val=\"10\"/><c:axId val=\"20\"/></c:" + plotKind + ">", "");
        PptxSceneChart sceneChart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml));
        PptxSceneChart chart = sceneChart;
        TestAssert.Equal(1, chart.Plots.Count);
        XElement element = XDocument.Parse(xml).Descendants(C + plotKind).First();
        return (chart.Plots[0], element);
    }

    private static string ChartSpace(string plotAreaInner, string options)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
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
