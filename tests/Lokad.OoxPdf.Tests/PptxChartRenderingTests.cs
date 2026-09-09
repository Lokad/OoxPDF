using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartRenderingTests
{
    public static void PptxSceneLineChartMarkerDefaultsUsePlotMarkerState()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:marker val="1"/>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:ser><c:marker><c:symbol val="plus"/><c:size val="9"/></c:marker><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:ser><c:marker><c:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr></c:marker><c:val><c:numLit><c:pt idx="0"><c:v>5</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartPlot plot = chart?.Plots[0] ?? throw new InvalidOperationException("Expected line-chart plot.");
        TestAssert.True(plot.MarkersEnabled == true, "Expected chart-level marker enablement to survive in the scene model.");
        TestAssert.Equal("1", plot.MarkersEnabledValue);
        TestAssert.True(plot.Series[0].Marker.IsDefined == false, "Expected an automatically enabled marker to remain distinct from series marker XML.");
        TestAssert.Equal("diamond", plot.Series[0].Marker.Symbol);
        TestAssert.Equal(7d, plot.Series[0].Marker.Size);
        TestAssert.True(plot.Series[1].Marker.IsDefined == true, "Expected explicit series marker XML to remain explicit.");
        TestAssert.Equal("plus", plot.Series[1].Marker.Symbol);
        TestAssert.Equal("9", plot.Series[1].Marker.SizeValue ?? string.Empty);
        TestAssert.True(plot.Series[2].Marker.IsDefined == false, "Expected missing sibling markers to stay missing even when their effective style is modeled.");
        TestAssert.Equal("triangle", plot.Series[2].Marker.Symbol);
        TestAssert.Equal(7d, plot.Series[2].Marker.Size);
        TestAssert.True(plot.Series[3].Marker.IsDefined == true, "Expected styled series marker XML to stay explicit.");
        TestAssert.Equal("x", plot.Series[3].Marker.Symbol);
        TestAssert.Equal(9d, plot.Series[3].Marker.Size);
    }

    public static void PptxChartMarkerStylesPreserveDefinitionState()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:marker val="1"/>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:ser><c:marker><c:symbol val="plus"/><c:size val="9"/><c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></c:spPr></c:marker><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots[0];

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedXmlFallback = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:lineChart>
                <c:marker val="0"/>
                <c:ser><c:marker><c:symbol val="none"/><c:size val="4"/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill></a:ln></c:spPr></c:marker></c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "lineChart").Single();
        System.Reflection.MethodInfo readStyles = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlMarkerStyles",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected marker-style resolver.");
        object[] sceneStyles = (((System.Collections.IEnumerable?)readStyles.Invoke(null, [plot, mismatchedXmlFallback, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected scene marker styles.")).Cast<object>().ToArray();
        TestAssert.Equal(2, sceneStyles.Length);
        TestAssert.Equal("diamond", PptxTests.ChartMarkerStyleSymbol(sceneStyles[0]));
        TestAssert.True(!PptxTests.ChartMarkerStyleIsDefined(sceneStyles[0]), "Expected auto line marker to remain distinct from explicit series marker XML.");
        TestAssert.Equal(string.Empty, PptxTests.ChartMarkerStyleSizeValue(sceneStyles[0]) ?? string.Empty);
        TestAssert.Equal(7d, PptxTests.ChartMarkerStyleSize(sceneStyles[0]));
        TestAssert.Equal("plus", PptxTests.ChartMarkerStyleSymbol(sceneStyles[1]));
        TestAssert.True(PptxTests.ChartMarkerStyleIsDefined(sceneStyles[1]), "Expected explicit series marker XML to remain explicit at the renderer option boundary.");
        TestAssert.Equal("9", PptxTests.ChartMarkerStyleSizeValue(sceneStyles[1]) ?? string.Empty);
        TestAssert.Equal(new RgbColor(17, 34, 51), PptxTests.ChartSeriesFillColor(PptxTests.ChartMarkerStyleFill(sceneStyles[1])));
        TestAssert.Equal(new RgbColor(68, 85, 102), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartMarkerStyleStroke(sceneStyles[1])));
        TestAssert.Equal(3d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartMarkerStyleStroke(sceneStyles[1])));

        XElement xmlOnly = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:lineChart>
                <c:marker val="1"/>
                <c:ser/>
                <c:ser><c:marker><c:symbol val="square"/><c:size val="8"/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill></a:ln></c:spPr></c:marker></c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "lineChart").Single();
        object[] xmlStyles = (((System.Collections.IEnumerable?)readStyles.Invoke(null, [null, xmlOnly, PptxTheme.Empty, PptxColorMap.Default])) ?? throw new InvalidOperationException("Expected XML marker styles.")).Cast<object>().ToArray();
        TestAssert.Equal(2, xmlStyles.Length);
        TestAssert.Equal("diamond", PptxTests.ChartMarkerStyleSymbol(xmlStyles[0]));
        TestAssert.True(!PptxTests.ChartMarkerStyleIsDefined(xmlStyles[0]), "Expected XML-only auto marker to remain distinct from explicit series marker XML.");
        TestAssert.Equal("square", PptxTests.ChartMarkerStyleSymbol(xmlStyles[1]));
        TestAssert.True(PptxTests.ChartMarkerStyleIsDefined(xmlStyles[1]), "Expected XML-only explicit marker to preserve definition state.");
        TestAssert.Equal("8", PptxTests.ChartMarkerStyleSizeValue(xmlStyles[1]) ?? string.Empty);
        TestAssert.Equal(new RgbColor(171, 205, 239), PptxTests.ChartSeriesFillColor(PptxTests.ChartMarkerStyleFill(xmlStyles[1])));
        TestAssert.Equal(new RgbColor(254, 220, 186), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartMarkerStyleStroke(xmlStyles[1])));
        TestAssert.Equal(1d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartMarkerStyleStroke(xmlStyles[1])));
    }

    public static void PptxSyntheticLineAndPieChartsRenderNativeCharts()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
                </Relationships>
                """),
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="457200" y="457200"/><a:ext cx="3657600" cy="2743200"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="5029200" y="457200"/><a:ext cx="3657600" cy="2743200"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="F0F0F0"/></a:gs><a:gs pos="100000"><a:srgbClr val="D0E0F0"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill><a:ln><a:solidFill><a:srgbClr val="444444"/></a:solidFill></a:ln></c:spPr>
                  <c:chart>
                  <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1150"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:rPr><a:t>Rev</a:t></a:r><a:r><a:rPr i="1"><a:latin typeface="Arial"/></a:rPr><a:t>enue</a:t></a:r></a:p></c:rich></c:tx><c:spPr><a:solidFill><a:srgbClr val="E6F8D0"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="301020"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1300" b="1" i="0"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:title>
                  <c:plotArea>
                  <c:layout><c:manualLayout><c:x val="0.2"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.5"/></c:manualLayout></c:layout>
                  <c:spPr><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
                  <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                  <c:valAx><c:axId val="20"/><c:delete/><c:scaling><c:min val="0"/><c:max val="10"/></c:scaling><c:majorUnit val="2"/><c:minorUnit val="1"/><c:majorGridlines/><c:minorGridlines/><c:spPr><a:ln><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln></c:spPr><c:crossAx val="10"/></c:valAx><c:lineChart>
                    <c:axId val="10"/>
                    <c:axId val="20"/>
                    <c:ser><c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="AA00AA"/></a:solidFill></a:ln></c:spPr>
                    <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Forecast</c:v></c:pt></c:strCache></c:strRef></c:tx>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit>
                      <c:pt idx="0"><c:v>2</c:v></c:pt>
                      <c:pt idx="1"><c:v>5</c:v></c:pt>
                      <c:pt idx="2"><c:v>4</c:v></c:pt>
                    </c:numLit></c:val><c:smooth val="1"/><c:marker><c:symbol val="square"/><c:size val="8"/><c:spPr><a:solidFill><a:srgbClr val="0000AA"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00AAAA"/></a:solidFill></a:ln></c:spPr></c:marker></c:ser>
                    <c:ser><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="AA5500"/></a:solidFill></a:ln></c:spPr><c:val><c:numLit>
                      <c:pt idx="0"><c:v>1</c:v></c:pt>
                      <c:pt idx="1"><c:v>3</c:v></c:pt>
                      <c:pt idx="2"><c:v>2</c:v></c:pt>
                    </c:numLit></c:val><c:marker><c:symbol val="plus"/><c:size val="9"/></c:marker></c:ser>
                    <c:ser><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="555555"/></a:solidFill></a:ln></c:spPr><c:val><c:numLit>
                      <c:pt idx="0"><c:v>3</c:v></c:pt>
                      <c:pt idx="1"><c:v>2</c:v></c:pt>
                      <c:pt idx="2"><c:v>1</c:v></c:pt>
                    </c:numLit></c:val><c:dLbls><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:dLblPos val="t"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800" b="1" i="1"><a:solidFill><a:srgbClr val="0066AA"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbls><c:marker><c:symbol val="star"/><c:size val="9"/></c:marker></c:ser>
                    <c:dLbls><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/><c:dLblPos val="b"/><c:separator> | </c:separator><c:spPr><a:solidFill><a:srgbClr val="FFEACC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="0A0B0C"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr><c:dLbl><c:idx val="1"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="850" b="1"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:rPr><a:t>ZX</a:t></a:r><a:r><a:rPr i="1"><a:latin typeface="Arial"/></a:rPr><a:t>Q</a:t></a:r></a:p></c:rich></c:tx><c:dLblPos val="ctr"/><c:layout><c:manualLayout><c:x val="0.44"/><c:y val="0.28"/><c:w val="0.1"/><c:h val="0.08"/></c:manualLayout></c:layout><c:spPr><a:solidFill><a:srgbClr val="CCEEFF"/></a:solidFill></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="6600CC"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbl></c:dLbls>
                  </c:lineChart></c:plotArea><c:legend><c:legendPos val="b"/><c:spPr><a:solidFill><a:srgbClr val="F8E6D0"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="102030"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="700" b="0" i="1"><a:solidFill><a:srgbClr val="654321"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:legend></c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/chart2.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea><c:pieChart>
                    <c:ser>
                    <c:dPt><c:idx val="1"/><c:explosion val="25"/><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill></a:ln></c:spPr></c:dPt>
                    <c:val><c:numLit>
                      <c:pt idx="0"><c:v>35</c:v></c:pt>
                      <c:pt idx="1"><c:v>25</c:v></c:pt>
                      <c:pt idx="2"><c:v>40</c:v></c:pt>
                    </c:numLit></c:val><c:dLbls><c:showVal val="1"/><c:showPercent val="1"/><c:showLeaderLines val="1"/><c:leaderLines><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="44CC88"/></a:solidFill></a:ln></c:spPr></c:leaderLines><c:dLbl><c:idx val="1"/><c:leaderLines><c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="CC11AA"/></a:solidFill></a:ln></c:spPr></c:leaderLines></c:dLbl></c:dLbls></c:ser>
                  </c:pieChart></c:plotArea></c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/style1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="42">
                  <cs:gridlineMajor>
                    <cs:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="224466"><a:alpha val="85000"/></a:srgbClr></a:solidFill></a:ln></cs:spPr>
                  </cs:gridlineMajor>
                </cs:style>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        OoxPackage package = OoxPackage.Open(input, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneChart lineChartScene = scene.Slides[0].SlideNodes
            .Select(node => node.Chart)
            .First(chart => chart?.TargetPartName == "/ppt/charts/chart1.xml")!;
        PptxSceneChartAxis valueAxisScene = lineChartScene.Axes.First(axis => axis.Id == "20");

        TestAssert.Equal("/ppt/charts/style1.xml", lineChartScene.StylePart.PartName ?? string.Empty);
        TestAssert.Equal("10,20", string.Join(",", lineChartScene.Plots.First(plot => plot.PlotKind == PptxSceneChartPlotKind.Line).AxisIds));
        TestAssert.True(!valueAxisScene.MajorGridlineLine.HasLine, "Expected the direct empty majorGridlines node to leave styling to chart-style fallback.");
        TestAssert.True(valueAxisScene.MajorGridlineStyleLine.HasLine, "Expected chart-style major gridline fallback in the scene axis.");
        TestAssert.Equal(new RgbColor(34, 68, 102), valueAxisScene.MajorGridlineStyleLine.Color);
        TestAssert.Equal(0.85d, valueAxisScene.MajorGridlineStyleLine.Alpha);

        TestAssert.Contains("/ShadingType 2", pdf);
        TestAssert.Contains(" sh", pdf);
        TestAssert.Contains("0.902 0.973 0.816 rg", pdf);
        TestAssert.Contains("0.188 0.063 0.125 RG", pdf);
        TestAssert.Contains("0.973 0.902 0.816 rg", pdf);
        TestAssert.Contains("0.063 0.125 0.188 RG", pdf);
        TestAssert.Contains("0.267 G", pdf);
        TestAssert.Contains("0 1 1 rg", pdf);
        TestAssert.Contains("93.6 352.8 144 108 re f", pdf);
        TestAssert.Contains("93.6 374.4 m", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"(?:[0-9.]+ [0-9.]+ m\s+[0-9.]+ [0-9.]+ l\s+){5}S"),
            "Line-chart major gridlines should share one stroked path for the five non-baseline tick marks.");
        TestAssert.Contains("0.133 0.267 0.4 RG", pdf);
        TestAssert.Contains("/CA 0.85", pdf);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("BT", pdf);
        TestAssert.True(pdf.Split("BT", StringSplitOptions.None).Length >= 7, "Chart title, axes, legend, and data labels should emit chart text objects.");
        TestAssert.Contains("0.922 G", pdf);
        TestAssert.DoesNotContain("0 0.667 0 RG", pdf);
        TestAssert.Contains("0.667 0 0.667 RG", pdf);
        TestAssert.Contains("8 8 re f", pdf);
        TestAssert.Contains("0.667 0.333 0 RG", pdf);
        TestAssert.Contains("0.333 G", pdf);
        TestAssert.Contains("0 0 0.667 rg", pdf);
        TestAssert.Contains("0 0.667 0.667 RG", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("0 0.667 0 rg", pdf);
        TestAssert.Contains("0.667 0 0 RG", pdf);
        TestAssert.Contains(" l S", pdf);
        TestAssert.Contains(" f", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <007C>"), "Expected line-chart data labels to include the custom separator between series, category, and value text.");
        TestAssert.Contains("1 0.918 0.8 rg", pdf);
        TestAssert.Contains("0.067 0.133 0.2 RG", pdf);
        TestAssert.Contains("79.2 357.525 28.8 13.5 re f", pdf);
        TestAssert.Contains("0.8 0.933 1 rg", pdf);
        TestAssert.Contains("214.56 373.995 14.4 8.64 re f", pdf);
        TestAssert.Contains("0.4 0 0.8 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CLD[0-9]+ 9 Tf"), "Expected per-label custom data label text style to drive font size.");
        TestAssert.Contains("0.071 0.204 0.337 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CT[0-9]+ 12\.96 Tf"), "Expected chart title txPr font size to drive title rendering.");
        TestAssert.Contains("0.671 0.804 0.937 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CT[0-9]+ 11\.52 Tf"), "Expected first rich-text title run to use its run font size.");
        TestAssert.Contains("0.396 0.263 0.129 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CL[0-9]+ 6\.96 Tf"), "Expected chart legend txPr font size to drive legend rendering.");
        TestAssert.True(Regex.IsMatch(pdf, @"1 0 0 1 [0-9.]+ 373\.995 Tm"), "Expected explicit data-label manualLayout to drive the custom label box baseline.");
        TestAssert.Contains("0.671 0.804 0.937 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CLD[0-9]+ 8\.52 Tf"), "Expected first custom rich-text data-label run to use its run font size.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <005A>") &&
            Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0058>") &&
            Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0051>"),
            "Expected per-label custom rich text to be emitted in the chart data label ToUnicode map.");
        TestAssert.Contains("0 0.4 0.667 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CLD[0-9]+ 8\.04 Tf"), "Expected second-series data labels to consume the series-level font size.");
        TestAssert.Contains("/ItalicAngle -12", pdf);
        TestAssert.Contains("0.039 0.043 0.047 rg", pdf);
        TestAssert.Contains("/CLD1 9.96 Tf", pdf);
        TestAssert.Contains("1 0 0 1 79.2 357.525 Tm", pdf);
        TestAssert.Contains("0.267 0.8 0.533 RG", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"0\.267 0\.8 0\.533 RG\s+1\.5 w[\s\S]+? m\s+[\s\S]+? l\s+[\s\S]+? l\s+S"),
            "Expected polar data-label leader lines to render as two-segment stroked paths using the preserved leader-line style.");
        TestAssert.True(Regex.IsMatch(pdf, @"0\.8 0\.067 0\.667 RG\s+2\.25 w[\s\S]+? m\s+[\s\S]+? l\s+[\s\S]+? l\s+S"),
            "Expected per-label polar leader-line style to override the chart-wide leader-line style.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0025>"), "Expected percentage labels to include a percent glyph in the ToUnicode map.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <002C>"), "Expected combined value/percentage labels to include the default comma separator.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Supported chart rendering should not emit static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Line and pie charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticPieChartUsesSparsePointIndicesForSliceStyles()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2743200"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea><c:pieChart>
                    <c:ser>
                      <c:dPt><c:idx val="4"/><c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></c:spPr></c:dPt>
                      <c:cat><c:strLit>
                        <c:ptCount val="5"/>
                        <c:pt idx="2"><c:v>Source two</c:v></c:pt>
                        <c:pt idx="4"><c:v>Source four</c:v></c:pt>
                      </c:strLit></c:cat>
                      <c:val><c:numLit>
                        <c:ptCount val="5"/>
                        <c:pt idx="0"><c:v></c:v></c:pt>
                        <c:pt idx="2"><c:v>30</c:v></c:pt>
                        <c:pt idx="4"><c:v>70</c:v></c:pt>
                      </c:numLit></c:val>
                    </c:ser>
                  </c:pieChart></c:plotArea><c:legend><c:legendPos val="r"/></c:legend></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.071 0.204 0.337 rg", pdf);
        TestAssert.True(Regex.Matches(pdf, "0\\.071 0\\.204 0\\.337 rg").Count >= 2, "Expected sparse point index 4 to drive both the slice fill and its category legend marker fill.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Sparse pie charts should render through the native chart path.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Sparse pie charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticAreaScatterRadarAndDoughnutChartsRenderNativeCharts()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
                  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
                  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart4.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="457200" y="457200"/><a:ext cx="1828800" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                    <p:graphicFrame><p:xfrm><a:off x="2743200" y="457200"/><a:ext cx="1828800" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame>
                    <p:graphicFrame><p:xfrm><a:off x="5029200" y="457200"/><a:ext cx="1828800" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId3"/></a:graphicData></a:graphic></p:graphicFrame>
                    <p:graphicFrame><p:xfrm><a:off x="7315200" y="457200"/><a:ext cx="1371600" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId4"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:areaChart>
                  <c:grouping val="stacked"/>
                  <c:ser><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:areaChart></c:plotArea></c:chart></c:chartSpace>
                """),
            ["ppt/charts/chart2.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:scatterChart>
                  <c:scatterStyle val="lineMarker"/>
                  <c:dLbls><c:showVal val="1"/><c:dLblPos val="t"/><c:numFmt formatCode="0"/></c:dLbls>
                  <c:ser><c:spPr><a:ln><a:solidFill><a:srgbClr val="AA00AA"/></a:solidFill></a:ln></c:spPr><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:yVal></c:ser>
                </c:scatterChart></c:plotArea></c:chart></c:chartSpace>
                """),
            ["ppt/charts/chart3.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:radarChart>
                  <c:radarStyle val="filled"/>
                  <c:ser><c:spPr><a:solidFill><a:srgbClr val="AAAA00"/></a:solidFill></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:radarChart></c:plotArea></c:chart></c:chartSpace>
                """),
            ["ppt/charts/chart4.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:doughnutChart>
                  <c:holeSize val="75"/>
                  <c:ser><c:explosion val="25"/><c:val><c:numLit><c:pt idx="0"><c:v>30</c:v></c:pt><c:pt idx="1"><c:v>70</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:doughnutChart></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0.667 0 rg", pdf);
        TestAssert.Contains("0.667 0 0.667 RG", pdf);
        TestAssert.Contains("0.667 0.667 0 rg", pdf);
        TestAssert.Contains("/CSD1", pdf);
        TestAssert.True(Regex.Matches(pdf, @"[0-9.]+ [0-9.]+ m").Count >= 4, "Expected native chart paths for area, scatter, radar, and doughnut charts.");
        TestAssert.DoesNotContain("/GS62000F100000S gs", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("f*", pdf);
        TestAssert.Contains(" l S", pdf);
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Supported chart rendering should not emit static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Area, scatter, radar, and doughnut charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxChartEmbeddedWorkbookFormulaOnlyReferencesReportMissingCachedData()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2743200"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/>
                </Relationships>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:chart><c:autoTitleDeleted val="1"/><c:plotVisOnly val="0"/><c:plotArea><c:doughnutChart>
                    <c:varyColors val="1"/>
                    <c:ser><c:idx val="0"/><c:order val="0"/><c:explosion val="25"/>
                      <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
                      <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
                      <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val>
                    </c:ser>
                    <c:firstSliceAng val="0"/><c:holeSize val="50"/>
                  </c:doughnutChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="1"/></c:legend></c:chart>
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr></a:p></c:txPr>
                  <c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>
                </c:chartSpace>
                """),
            ["ppt/embeddings/Microsoft_Excel_Worksheet.xlsx"] = PptxTests.EmbeddedChartWorkbook()
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int pathStartCount = Regex.Matches(pdf, @"[0-9.]+ [0-9.]+ m").Count;
        TestAssert.True(pathStartCount == 0, $"Expected formula-only embedded workbook references to stay out of active chart rendering without chart-side caches; saw {pathStartCount} path starts and diagnostics: {string.Join(", ", collector.Diagnostics.Select(d => d.Id))}.");
        TestAssert.True(PptxTests.CountTextMatrices(pdf) == 0, "Expected formula-only embedded workbook category references to stay out of native chart text without chart-side caches.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Formula-only charts should not fall back to static chart rendering.");
        TestAssert.True(collector.Diagnostics.Any(d => d.Id == "PPTX_CHART_MISSING_CACHED_DATA"), "Formula-only embedded workbook references should report missing chart-side caches while preserving workbook values as provenance.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Formula-only supported chart families should report missing cached data rather than a generic unsupported chart.");
    }

    public static void PptxChartRawVectorsPreserveWorkbookSidecarPoints()
    {
        Type workbookType = typeof(PptxRenderer).GetNestedType(
            "ChartWorkbookData",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart workbook data type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["A2"] = "North",
                ["A4"] = "West",
                ["B2"] = "8.2",
                ["B4"] = "1.4"
            }
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets]) ?? throw new InvalidOperationException("Expected workbook instance.");
        var readNumberVector = typeof(PptxRenderer)
            .GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(method => method.Name == "ReadChartNumberVector" && method.GetParameters().Length == 3);
        var readCategoryVector = typeof(PptxRenderer)
            .GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(method => method.Name == "ReadChartCategoryLabelVector" && method.GetParameters().Length == 3);

        XElement values = XElement.Parse("""
            <c:val xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:numRef>
                <c:f>Sheet1!$B$2:$B$4</c:f>
                <c:numCache>
                  <c:formatCode>General</c:formatCode>
                  <c:ptCount val="4"/>
                  <c:pt idx="2"><c:v>100</c:v></c:pt>
                  <c:pt idx="futureNumber"><c:v>200</c:v></c:pt>
                </c:numCache>
              </c:numRef>
            </c:val>
            """);
        object numberVector = readNumberVector.Invoke(null, [values, workbook, true]) ?? throw new InvalidOperationException("Expected raw number vector.");
        object[] cacheNumberPoints = ((System.Collections.IEnumerable?)numberVector.GetType().GetProperty("Points")?.GetValue(numberVector))?.Cast<object>().ToArray() ?? [];
        object[] workbookNumberPoints = ((System.Collections.IEnumerable?)numberVector.GetType().GetProperty("WorkbookPoints")?.GetValue(numberVector))?.Cast<object>().ToArray() ?? [];
        TestAssert.True(cacheNumberPoints.Length == 2, "Expected raw vector rendering points to stay bound to the chart cache.");
        TestAssert.True((int?)cacheNumberPoints[0].GetType().GetProperty("Index")?.GetValue(cacheNumberPoints[0]) == 2, "Expected raw numeric cache point index to survive.");
        TestAssert.Equal("OoxmlIndex", cacheNumberPoints[0].GetType().GetProperty("IndexSource")?.GetValue(cacheNumberPoints[0])?.ToString() ?? string.Empty);
        TestAssert.Equal(100d, (double?)cacheNumberPoints[0].GetType().GetProperty("Value")?.GetValue(cacheNumberPoints[0]) ?? 0d);
        TestAssert.True((int?)cacheNumberPoints[1].GetType().GetProperty("Index")?.GetValue(cacheNumberPoints[1]) == 1, "Expected malformed raw numeric cache point index to keep the ordinal fallback.");
        TestAssert.Equal("OrdinalFallback", cacheNumberPoints[1].GetType().GetProperty("IndexSource")?.GetValue(cacheNumberPoints[1])?.ToString() ?? string.Empty);
        TestAssert.True(workbookNumberPoints.Length == 3, "Expected raw numeric vector to keep workbook sidecar points without promoting them to renderable cache points.");
        TestAssert.Equal("WorkbookRange", workbookNumberPoints[0].GetType().GetProperty("IndexSource")?.GetValue(workbookNumberPoints[0])?.ToString() ?? string.Empty);
        TestAssert.Equal("1.4", (string?)workbookNumberPoints[2].GetType().GetProperty("Text")?.GetValue(workbookNumberPoints[2]) ?? string.Empty);

        XElement chart = XElement.Parse("""
            <c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:ser>
                <c:cat>
                  <c:strRef>
                    <c:f>Sheet1!$A$2:$A$4</c:f>
                    <c:strCache>
                      <c:ptCount val="4"/>
                      <c:pt idx="2"><c:v>Cached</c:v></c:pt>
                      <c:pt idx="futureCategory"><c:v>Fallback</c:v></c:pt>
                    </c:strCache>
                  </c:strRef>
                </c:cat>
              </c:ser>
            </c:barChart>
            """);
        object categoryVector = readCategoryVector.Invoke(null, [chart, workbook, true]) ?? throw new InvalidOperationException("Expected raw category vector.");
        object[] cacheTextPoints = ((System.Collections.IEnumerable?)categoryVector.GetType().GetProperty("Points")?.GetValue(categoryVector))?.Cast<object>().ToArray() ?? [];
        object[] workbookTextPoints = ((System.Collections.IEnumerable?)categoryVector.GetType().GetProperty("WorkbookPoints")?.GetValue(categoryVector))?.Cast<object>().ToArray() ?? [];
        TestAssert.True(cacheTextPoints.Length == 2, "Expected raw category rendering points to stay bound to the chart cache.");
        TestAssert.Equal("Cached", (string?)cacheTextPoints[0].GetType().GetProperty("Text")?.GetValue(cacheTextPoints[0]) ?? string.Empty);
        TestAssert.Equal("OoxmlIndex", cacheTextPoints[0].GetType().GetProperty("IndexSource")?.GetValue(cacheTextPoints[0])?.ToString() ?? string.Empty);
        TestAssert.True((int?)cacheTextPoints[1].GetType().GetProperty("Index")?.GetValue(cacheTextPoints[1]) == 1, "Expected malformed raw text cache point index to keep the ordinal fallback.");
        TestAssert.Equal("OrdinalFallback", cacheTextPoints[1].GetType().GetProperty("IndexSource")?.GetValue(cacheTextPoints[1])?.ToString() ?? string.Empty);
        TestAssert.True(workbookTextPoints.Length == 3, "Expected raw category vector to keep workbook sidecar text points.");
        TestAssert.Equal("WorkbookRange", workbookTextPoints[0].GetType().GetProperty("IndexSource")?.GetValue(workbookTextPoints[0])?.ToString() ?? string.Empty);
        TestAssert.Equal("West", (string?)workbookTextPoints[2].GetType().GetProperty("Text")?.GetValue(workbookTextPoints[2]) ?? string.Empty);
    }

    public static void PptxChartWorkbookHydrationPreservesRangePointIndices()
    {
        Type workbookType = typeof(PptxRenderer).GetNestedType(
            "ChartWorkbookData",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart workbook data type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["A2"] = "North",
                ["A4"] = "West",
                ["B2"] = "8.2",
                ["B4"] = "1.4"
            }
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets]) ?? throw new InvalidOperationException("Expected workbook instance.");
        System.Reflection.PropertyInfo date1904Property = workbookType.GetProperty("Date1904") ?? throw new InvalidOperationException("Expected workbook date1904 property.");
        TestAssert.True((bool?)date1904Property.GetValue(workbook) == false, "Expected workbook date1904 metadata to default to false.");
        object date1904Workbook = Activator.CreateInstance(workbookType, [sheets, true]) ?? throw new InvalidOperationException("Expected date1904 workbook instance.");
        TestAssert.True((bool?)date1904Property.GetValue(date1904Workbook) == true, "Expected workbook date1904 metadata to survive construction.");
        using MemoryStream embeddedWorkbookStream = new(PptxTests.EmbeddedChartWorkbook());
        OoxPackage embeddedWorkbookPackage = OoxPackage.Open(embeddedWorkbookStream, CancellationToken.None);
        var readWorkbookData = typeof(PptxRenderer).GetMethod(
            "ReadWorkbookData",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected workbook reader helper.");
        object parsedWorkbook = readWorkbookData.Invoke(null, [embeddedWorkbookPackage]) ?? throw new InvalidOperationException("Expected parsed workbook data.");
        TestAssert.True((bool?)date1904Property.GetValue(parsedWorkbook) == true, "Expected parsed workbookPr/date1904 to survive workbook parsing.");
        System.Reflection.PropertyInfo sheetsProperty = workbookType.GetProperty("Sheets") ?? throw new InvalidOperationException("Expected workbook sheet records.");
        object[] parsedSheets = (((System.Collections.IEnumerable?)sheetsProperty.GetValue(parsedWorkbook)) ?? throw new InvalidOperationException("Expected parsed workbook sheet records.")).Cast<object>().ToArray();
        TestAssert.True(parsedSheets.Length == 1, "Expected workbook sheet catalog to preserve the declared sheet.");
        object parsedSheet = parsedSheets[0];
        TestAssert.Equal("Sheet1", (string?)parsedSheet.GetType().GetProperty("Name")?.GetValue(parsedSheet) ?? string.Empty);
        TestAssert.Equal("1", (string?)parsedSheet.GetType().GetProperty("SheetId")?.GetValue(parsedSheet) ?? string.Empty);
        TestAssert.Equal("rId1", (string?)parsedSheet.GetType().GetProperty("RelationshipId")?.GetValue(parsedSheet) ?? string.Empty);
        TestAssert.Equal("hidden", (string?)parsedSheet.GetType().GetProperty("State")?.GetValue(parsedSheet) ?? string.Empty);
        TestAssert.True((int?)parsedSheet.GetType().GetProperty("Index")?.GetValue(parsedSheet) == 0, "Expected workbook sheet order to survive parsing.");
        TestAssert.Equal("/xl/worksheets/sheet1.xml", (string?)parsedSheet.GetType().GetProperty("TargetPartName")?.GetValue(parsedSheet) ?? string.Empty);
        var readRangeCells = workbookType.GetMethod("ReadRangeCells") ?? throw new InvalidOperationException("Expected range-cell reader.");
        Array parsedCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["Sheet1!$B$2:$B$4"]) ?? throw new InvalidOperationException("Expected parsed workbook range cells."));
        object firstParsedCell = parsedCells.GetValue(0) ?? throw new InvalidOperationException("Expected first parsed workbook range cell.");
        System.Reflection.PropertyInfo sheetNameProperty = firstParsedCell.GetType().GetProperty("SheetName") ?? throw new InvalidOperationException("Expected range-cell sheet name.");
        TestAssert.Equal("Sheet1", (string?)sheetNameProperty.GetValue(firstParsedCell) ?? string.Empty);
        TestAssert.Equal("Sheet1!$B$2:$B$4", (string?)firstParsedCell.GetType().GetProperty("SourceFormula")?.GetValue(firstParsedCell) ?? string.Empty);
        TestAssert.Equal("Sheet1!$B$2:$B$4", (string?)firstParsedCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedCell) ?? string.Empty);
        TestAssert.Equal("DirectRange", firstParsedCell.GetType().GetProperty("SourceKind")?.GetValue(firstParsedCell)?.ToString() ?? string.Empty);
        TestAssert.Equal(string.Empty, (string?)firstParsedCell.GetType().GetProperty("DefinedName")?.GetValue(firstParsedCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedCell.GetType().GetProperty("RangeRowIndex")?.GetValue(firstParsedCell) == 0, "Expected range-local row index to survive range-cell projection.");
        TestAssert.True((int?)firstParsedCell.GetType().GetProperty("RangeColumnIndex")?.GetValue(firstParsedCell) == 0, "Expected range-local column index to survive range-cell projection.");
        TestAssert.True((int?)firstParsedCell.GetType().GetProperty("RangeRowCount")?.GetValue(firstParsedCell) == 3, "Expected range row count to survive range-cell projection.");
        TestAssert.True((int?)firstParsedCell.GetType().GetProperty("RangeColumnCount")?.GetValue(firstParsedCell) == 1, "Expected range column count to survive range-cell projection.");
        TestAssert.True((int?)firstParsedCell.GetType().GetProperty("SheetRow")?.GetValue(firstParsedCell) == 2, "Expected absolute worksheet row to survive range-cell projection.");
        TestAssert.True((int?)firstParsedCell.GetType().GetProperty("SheetColumn")?.GetValue(firstParsedCell) == 2, "Expected absolute worksheet column to survive range-cell projection.");
        TestAssert.Equal("Number", firstParsedCell.GetType().GetProperty("ValueKind")?.GetValue(firstParsedCell)?.ToString() ?? string.Empty);
        TestAssert.Equal("8.2", (string?)firstParsedCell.GetType().GetProperty("RawValue")?.GetValue(firstParsedCell) ?? string.Empty);
        TestAssert.True((bool?)firstParsedCell.GetType().GetProperty("HasValueElement")?.GetValue(firstParsedCell) == true, "Expected numeric worksheet cell to preserve its cached value element.");
        System.Reflection.PropertyInfo styleIndexProperty = firstParsedCell.GetType().GetProperty("StyleIndex") ?? throw new InvalidOperationException("Expected range-cell style index.");
        TestAssert.True((int?)styleIndexProperty.GetValue(firstParsedCell) == 5, "Expected worksheet cell style index to survive workbook parsing.");
        System.Reflection.PropertyInfo styleNumberFormatIdProperty = firstParsedCell.GetType().GetProperty("StyleNumberFormatId") ?? throw new InvalidOperationException("Expected range-cell style number-format ID.");
        TestAssert.True((int?)styleNumberFormatIdProperty.GetValue(firstParsedCell) == 165, "Expected worksheet cell style to resolve its number-format ID.");
        System.Reflection.PropertyInfo styleNumberFormatCodeProperty = firstParsedCell.GetType().GetProperty("StyleNumberFormatCode") ?? throw new InvalidOperationException("Expected range-cell style number-format code.");
        TestAssert.Equal("m/d/yy", (string?)styleNumberFormatCodeProperty.GetValue(firstParsedCell) ?? string.Empty);
        System.Reflection.PropertyInfo styleAppliesNumberFormatProperty = firstParsedCell.GetType().GetProperty("StyleAppliesNumberFormat") ?? throw new InvalidOperationException("Expected range-cell style number-format apply flag.");
        TestAssert.True((bool?)styleAppliesNumberFormatProperty.GetValue(firstParsedCell) == true, "Expected worksheet cell style to preserve applyNumberFormat.");
        System.Reflection.PropertyInfo styleNumberFormatIsDateLikeProperty = firstParsedCell.GetType().GetProperty("StyleNumberFormatIsDateLike") ?? throw new InvalidOperationException("Expected range-cell style date-like number-format flag.");
        TestAssert.True((bool?)styleNumberFormatIsDateLikeProperty.GetValue(firstParsedCell) == true, "Expected worksheet cell style to classify its date-like number format before chart formatting consumes it.");
        System.Reflection.PropertyInfo rowHiddenProperty = firstParsedCell.GetType().GetProperty("RowHidden") ?? throw new InvalidOperationException("Expected range-cell row-hidden flag.");
        System.Reflection.PropertyInfo columnHiddenProperty = firstParsedCell.GetType().GetProperty("ColumnHidden") ?? throw new InvalidOperationException("Expected range-cell column-hidden flag.");
        TestAssert.True((bool?)rowHiddenProperty.GetValue(firstParsedCell) == false, "Expected visible worksheet row to remain visible in range metadata.");
        TestAssert.True((bool?)columnHiddenProperty.GetValue(firstParsedCell) == true, "Expected hidden worksheet column to survive workbook parsing.");
        Array parsedExternalQuotedCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["'[Book.xlsx]Sheet1'!$B$2:$B$4"]) ?? throw new InvalidOperationException("Expected parsed external-workbook-qualified range cells."));
        object firstParsedExternalQuotedCell = parsedExternalQuotedCells.GetValue(0) ?? throw new InvalidOperationException("Expected first external-workbook-qualified range cell.");
        TestAssert.Equal("Sheet1", (string?)sheetNameProperty.GetValue(firstParsedExternalQuotedCell) ?? string.Empty);
        TestAssert.Equal("'[Book.xlsx]Sheet1'!$B$2:$B$4", (string?)firstParsedExternalQuotedCell.GetType().GetProperty("SourceFormula")?.GetValue(firstParsedExternalQuotedCell) ?? string.Empty);
        TestAssert.Equal("'[Book.xlsx]Sheet1'!$B$2:$B$4", (string?)firstParsedExternalQuotedCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedExternalQuotedCell) ?? string.Empty);
        Array parsedUnionCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["Sheet1!$B$2:$B$2,$B$4:$B$4"]) ?? throw new InvalidOperationException("Expected parsed multi-area workbook range cells."));
        object firstParsedUnionCell = parsedUnionCells.GetValue(0) ?? throw new InvalidOperationException("Expected first multi-area workbook range cell.");
        object secondParsedUnionCell = parsedUnionCells.GetValue(1) ?? throw new InvalidOperationException("Expected second multi-area workbook range cell.");
        TestAssert.True(parsedUnionCells.Length == 2, "Expected multi-area workbook range to preserve cells from both areas.");
        TestAssert.True((int?)firstParsedUnionCell.GetType().GetProperty("Index")?.GetValue(firstParsedUnionCell) == 0, "Expected multi-area range to preserve global point order.");
        TestAssert.True((int?)firstParsedUnionCell.GetType().GetProperty("RangeAreaIndex")?.GetValue(firstParsedUnionCell) == 0, "Expected first multi-area cell to preserve its range-area index.");
        TestAssert.True((int?)secondParsedUnionCell.GetType().GetProperty("RangeAreaIndex")?.GetValue(secondParsedUnionCell) == 1, "Expected second multi-area cell to preserve its range-area index.");
        TestAssert.True((int?)secondParsedUnionCell.GetType().GetProperty("RangeAreaCount")?.GetValue(secondParsedUnionCell) == 2, "Expected multi-area cells to preserve the total range-area count.");
        TestAssert.True((int?)secondParsedUnionCell.GetType().GetProperty("Index")?.GetValue(secondParsedUnionCell) == 1, "Expected multi-area range to keep point indices continuous across areas.");
        TestAssert.Equal("Sheet1!$B$2:$B$2,$B$4:$B$4", (string?)secondParsedUnionCell.GetType().GetProperty("SourceFormula")?.GetValue(secondParsedUnionCell) ?? string.Empty);
        TestAssert.Equal("Sheet1!$B$4:$B$4", (string?)secondParsedUnionCell.GetType().GetProperty("ResolvedFormula")?.GetValue(secondParsedUnionCell) ?? string.Empty);
        TestAssert.Equal("1.4", (string?)secondParsedUnionCell.GetType().GetProperty("RawValue")?.GetValue(secondParsedUnionCell) ?? string.Empty);
        var commaSheetSheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Q,1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["A1"] = "3",
                ["A2"] = "5"
            }
        };
        object commaSheetWorkbook = Activator.CreateInstance(workbookType, [commaSheetSheets]) ?? throw new InvalidOperationException("Expected quoted comma-sheet workbook instance.");
        Array commaSheetUnionCells = (Array)(readRangeCells.Invoke(commaSheetWorkbook, ["'Q,1'!$A$1:$A$1,$A$2:$A$2"]) ?? throw new InvalidOperationException("Expected quoted comma-sheet multi-area range cells."));
        object firstCommaSheetUnionCell = commaSheetUnionCells.GetValue(0) ?? throw new InvalidOperationException("Expected first quoted comma-sheet multi-area range cell.");
        object secondCommaSheetUnionCell = commaSheetUnionCells.GetValue(1) ?? throw new InvalidOperationException("Expected second quoted comma-sheet multi-area range cell.");
        TestAssert.True(commaSheetUnionCells.Length == 2, "Expected multi-area range parsing not to split commas inside quoted sheet names.");
        TestAssert.Equal("Q,1", (string?)sheetNameProperty.GetValue(firstCommaSheetUnionCell) ?? string.Empty);
        TestAssert.Equal("'Q,1'!$A$1:$A$1,$A$2:$A$2", (string?)secondCommaSheetUnionCell.GetType().GetProperty("SourceFormula")?.GetValue(secondCommaSheetUnionCell) ?? string.Empty);
        TestAssert.Equal("'Q,1'!$A$2:$A$2", (string?)secondCommaSheetUnionCell.GetType().GetProperty("ResolvedFormula")?.GetValue(secondCommaSheetUnionCell) ?? string.Empty);
        TestAssert.True((int?)secondCommaSheetUnionCell.GetType().GetProperty("RangeAreaIndex")?.GetValue(secondCommaSheetUnionCell) == 1, "Expected quoted comma-sheet range unions to preserve area ownership.");
        var readNumericRange = workbookType.GetMethod("ReadNumericRange") ?? throw new InvalidOperationException("Expected typed numeric range reader.");
        Array parsedNumericValues = (Array)(readNumericRange.Invoke(parsedWorkbook, ["Sheet1!$B$2:$B$4"]) ?? throw new InvalidOperationException("Expected typed workbook numeric values."));
        object firstParsedNumericValue = parsedNumericValues.GetValue(0) ?? throw new InvalidOperationException("Expected first typed numeric value.");
        System.Reflection.PropertyInfo numericValueProperty = firstParsedNumericValue.GetType().GetProperty("Value") ?? throw new InvalidOperationException("Expected typed numeric value.");
        TestAssert.Equal(8.2d, (double?)numericValueProperty.GetValue(firstParsedNumericValue) ?? 0d);
        System.Reflection.PropertyInfo numericCellProperty = firstParsedNumericValue.GetType().GetProperty("Cell") ?? throw new InvalidOperationException("Expected typed numeric range-cell ownership.");
        object firstParsedNumericCell = numericCellProperty.GetValue(firstParsedNumericValue) ?? throw new InvalidOperationException("Expected typed numeric range cell.");
        TestAssert.True((bool?)columnHiddenProperty.GetValue(firstParsedNumericCell) == true, "Expected typed numeric value to preserve its source range-cell metadata.");
        Array parsedUnionNumericValues = (Array)(readNumericRange.Invoke(parsedWorkbook, ["Sheet1!$B$2:$B$2,$B$4:$B$4"]) ?? throw new InvalidOperationException("Expected typed multi-area workbook numeric values."));
        TestAssert.True(parsedUnionNumericValues.Length == 2, "Expected typed numeric range reader to preserve multi-area workbook values.");
        object secondParsedUnionNumericValue = parsedUnionNumericValues.GetValue(1) ?? throw new InvalidOperationException("Expected second typed multi-area numeric value.");
        TestAssert.Equal(1.4d, (double?)numericValueProperty.GetValue(secondParsedUnionNumericValue) ?? 0d);
        object secondParsedUnionNumericCell = numericCellProperty.GetValue(secondParsedUnionNumericValue) ?? throw new InvalidOperationException("Expected typed multi-area numeric range cell.");
        TestAssert.True((int?)secondParsedUnionNumericCell.GetType().GetProperty("RangeAreaIndex")?.GetValue(secondParsedUnionNumericCell) == 1, "Expected typed numeric value to preserve its multi-area source metadata.");
        var readTextRange = workbookType.GetMethod("ReadTextRange") ?? throw new InvalidOperationException("Expected typed text range reader.");
        Array parsedTextValues = (Array)(readTextRange.Invoke(parsedWorkbook, ["SalesLabels"]) ?? throw new InvalidOperationException("Expected typed workbook text values."));
        object firstParsedTextValue = parsedTextValues.GetValue(0) ?? throw new InvalidOperationException("Expected first typed text value.");
        System.Reflection.PropertyInfo textValueProperty = firstParsedTextValue.GetType().GetProperty("Text") ?? throw new InvalidOperationException("Expected typed text value.");
        TestAssert.Equal("North", (string?)textValueProperty.GetValue(firstParsedTextValue) ?? string.Empty);
        System.Reflection.PropertyInfo textCellProperty = firstParsedTextValue.GetType().GetProperty("Cell") ?? throw new InvalidOperationException("Expected typed text range-cell ownership.");
        object firstParsedTextCell = textCellProperty.GetValue(firstParsedTextValue) ?? throw new InvalidOperationException("Expected typed text range cell.");
        TestAssert.Equal("SalesLabels", (string?)firstParsedTextCell.GetType().GetProperty("SourceFormula")?.GetValue(firstParsedTextCell) ?? string.Empty);
        TestAssert.Equal("Sheet1!$A$2:$A$4", (string?)firstParsedTextCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedTextCell) ?? string.Empty);
        TestAssert.Equal("DefinedName", firstParsedTextCell.GetType().GetProperty("SourceKind")?.GetValue(firstParsedTextCell)?.ToString() ?? string.Empty);
        TestAssert.Equal("SalesLabels", (string?)firstParsedTextCell.GetType().GetProperty("DefinedName")?.GetValue(firstParsedTextCell) ?? string.Empty);
        TestAssert.Equal("SharedString", firstParsedTextCell.GetType().GetProperty("ValueKind")?.GetValue(firstParsedTextCell)?.ToString() ?? string.Empty);
        TestAssert.Equal("0", (string?)firstParsedTextCell.GetType().GetProperty("RawValue")?.GetValue(firstParsedTextCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedTextCell.GetType().GetProperty("SharedStringIndex")?.GetValue(firstParsedTextCell) == 0, "Expected shared-string index to survive workbook parsing.");
        TestAssert.True((int?)firstParsedTextCell.GetType().GetProperty("SharedStringRunCount")?.GetValue(firstParsedTextCell) == 2, "Expected shared-string rich-text run count to survive workbook parsing.");
        TestAssert.True((bool?)firstParsedTextCell.GetType().GetProperty("SharedStringHasRichText")?.GetValue(firstParsedTextCell) == true, "Expected shared-string rich-text presence to survive workbook parsing.");
        TestAssert.True((bool?)firstParsedTextCell.GetType().GetProperty("SharedStringPreserveSpace")?.GetValue(firstParsedTextCell) == true, "Expected shared-string xml:space metadata to survive workbook parsing.");
        object lastParsedTextValue = parsedTextValues.GetValue(2) ?? throw new InvalidOperationException("Expected inline-string typed text value.");
        object lastParsedTextCell = textCellProperty.GetValue(lastParsedTextValue) ?? throw new InvalidOperationException("Expected inline-string range cell.");
        TestAssert.Equal("InlineString", lastParsedTextCell.GetType().GetProperty("ValueKind")?.GetValue(lastParsedTextCell)?.ToString() ?? string.Empty);
        TestAssert.True((bool?)lastParsedTextCell.GetType().GetProperty("HasValueElement")?.GetValue(lastParsedTextCell) == false, "Expected inline string cells to remain distinguishable from cached <v> cells.");
        Array parsedStructuredCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["SalesTable[Amount]"]) ?? throw new InvalidOperationException("Expected parsed structured-reference range cells."));
        object firstParsedStructuredCell = parsedStructuredCells.GetValue(0) ?? throw new InvalidOperationException("Expected first structured-reference range cell.");
        TestAssert.Equal("SalesTable[Amount]", (string?)firstParsedStructuredCell.GetType().GetProperty("SourceFormula")?.GetValue(firstParsedStructuredCell) ?? string.Empty);
        TestAssert.Equal("Sheet1!B2:B4", (string?)firstParsedStructuredCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedStructuredCell) ?? string.Empty);
        TestAssert.Equal("StructuredReference", firstParsedStructuredCell.GetType().GetProperty("SourceKind")?.GetValue(firstParsedStructuredCell)?.ToString() ?? string.Empty);
        TestAssert.Equal("SalesTable", (string?)firstParsedStructuredCell.GetType().GetProperty("TableName")?.GetValue(firstParsedStructuredCell) ?? string.Empty);
        TestAssert.Equal("Amount", (string?)firstParsedStructuredCell.GetType().GetProperty("TableColumnName")?.GetValue(firstParsedStructuredCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedStructuredCell.GetType().GetProperty("TableColumnId")?.GetValue(firstParsedStructuredCell) == 2, "Expected structured-reference source column id to survive resolution.");
        TestAssert.Equal("Amount", (string?)firstParsedStructuredCell.GetType().GetProperty("TableFirstColumnName")?.GetValue(firstParsedStructuredCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedStructuredCell.GetType().GetProperty("TableFirstColumnId")?.GetValue(firstParsedStructuredCell) == 2, "Expected structured-reference first source column id to survive resolution.");
        TestAssert.Equal("Amount", (string?)firstParsedStructuredCell.GetType().GetProperty("TableLastColumnName")?.GetValue(firstParsedStructuredCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedStructuredCell.GetType().GetProperty("TableLastColumnId")?.GetValue(firstParsedStructuredCell) == 2, "Expected structured-reference last source column id to survive resolution.");
        TestAssert.True((int?)firstParsedStructuredCell.GetType().GetProperty("TableRowIndex")?.GetValue(firstParsedStructuredCell) == 1, "Expected structured-reference table row index to survive resolution.");
        TestAssert.True((bool?)firstParsedStructuredCell.GetType().GetProperty("TableDataRow")?.GetValue(firstParsedStructuredCell) == true, "Expected structured-reference data row role to survive resolution.");
        TestAssert.True((bool?)firstParsedStructuredCell.GetType().GetProperty("TableHeaderRow")?.GetValue(firstParsedStructuredCell) == false, "Expected structured-reference data row not to be classified as a header row.");
        TestAssert.True((bool?)firstParsedStructuredCell.GetType().GetProperty("TableTotalsRow")?.GetValue(firstParsedStructuredCell) == false, "Expected structured-reference data row not to be classified as a totals row.");
        TestAssert.True((int?)firstParsedStructuredCell.GetType().GetProperty("TableCellColumnIndex")?.GetValue(firstParsedStructuredCell) == 1, "Expected structured-reference cell column index to survive resolution.");
        TestAssert.Equal("Amount", (string?)firstParsedStructuredCell.GetType().GetProperty("TableCellColumnName")?.GetValue(firstParsedStructuredCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedStructuredCell.GetType().GetProperty("TableCellColumnId")?.GetValue(firstParsedStructuredCell) == 2, "Expected structured-reference cell column id to survive resolution.");
        TestAssert.Equal("B2", (string?)firstParsedStructuredCell.GetType().GetProperty("TableCellCalculatedColumnFormula")?.GetValue(firstParsedStructuredCell) ?? string.Empty);
        Array parsedWholeTableCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["SalesTable[#Data]"]) ?? throw new InvalidOperationException("Expected parsed whole-table structured-reference range cells."));
        object firstParsedWholeTableCell = parsedWholeTableCells.GetValue(0) ?? throw new InvalidOperationException("Expected first whole-table structured-reference range cell.");
        TestAssert.True(parsedWholeTableCells.Length == 9, "Expected whole-table data-body structured reference to hydrate the complete data-body rectangle.");
        TestAssert.Equal("Sheet1!A2:C4", (string?)firstParsedWholeTableCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedWholeTableCell) ?? string.Empty);
        TestAssert.Equal("SalesTable", (string?)firstParsedWholeTableCell.GetType().GetProperty("TableName")?.GetValue(firstParsedWholeTableCell) ?? string.Empty);
        TestAssert.Equal(string.Empty, (string?)firstParsedWholeTableCell.GetType().GetProperty("TableColumnName")?.GetValue(firstParsedWholeTableCell) ?? string.Empty);
        TestAssert.True(firstParsedWholeTableCell.GetType().GetProperty("TableColumnId")?.GetValue(firstParsedWholeTableCell) is null, "Expected whole-table structured references to keep column id empty.");
        TestAssert.Equal("Region", (string?)firstParsedWholeTableCell.GetType().GetProperty("TableFirstColumnName")?.GetValue(firstParsedWholeTableCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedWholeTableCell.GetType().GetProperty("TableFirstColumnId")?.GetValue(firstParsedWholeTableCell) == 1, "Expected whole-table first source column id to survive resolution.");
        TestAssert.Equal("Quoted ] Amount", (string?)firstParsedWholeTableCell.GetType().GetProperty("TableLastColumnName")?.GetValue(firstParsedWholeTableCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedWholeTableCell.GetType().GetProperty("TableLastColumnId")?.GetValue(firstParsedWholeTableCell) == 3, "Expected whole-table last source column id to survive resolution.");
        Array parsedHeaderStructuredCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["SalesTable[#Headers]"]) ?? throw new InvalidOperationException("Expected parsed header structured-reference range cells."));
        object firstParsedHeaderStructuredCell = parsedHeaderStructuredCells.GetValue(0) ?? throw new InvalidOperationException("Expected first header structured-reference range cell.");
        TestAssert.True(parsedHeaderStructuredCells.Length == 3, "Expected header-only structured reference to hydrate the complete header row.");
        TestAssert.Equal("Sheet1!A1:C1", (string?)firstParsedHeaderStructuredCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedHeaderStructuredCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedHeaderStructuredCell.GetType().GetProperty("RangeRowCount")?.GetValue(firstParsedHeaderStructuredCell) == 1, "Expected header-only structured reference row count to survive range projection.");
        TestAssert.True((int?)firstParsedHeaderStructuredCell.GetType().GetProperty("RangeColumnCount")?.GetValue(firstParsedHeaderStructuredCell) == 3, "Expected header-only structured reference column count to survive range projection.");
        TestAssert.True((int?)firstParsedHeaderStructuredCell.GetType().GetProperty("TableRowIndex")?.GetValue(firstParsedHeaderStructuredCell) == 0, "Expected header-only structured reference table row index to survive resolution.");
        TestAssert.True((bool?)firstParsedHeaderStructuredCell.GetType().GetProperty("TableHeaderRow")?.GetValue(firstParsedHeaderStructuredCell) == true, "Expected header-only structured reference row role to survive resolution.");
        TestAssert.True((bool?)firstParsedHeaderStructuredCell.GetType().GetProperty("TableDataRow")?.GetValue(firstParsedHeaderStructuredCell) == false, "Expected header-only structured reference not to be classified as a data row.");
        TestAssert.True((bool?)firstParsedHeaderStructuredCell.GetType().GetProperty("TableTotalsRow")?.GetValue(firstParsedHeaderStructuredCell) == false, "Expected header-only structured reference not to be classified as a totals row.");
        TestAssert.Equal("Region", (string?)firstParsedHeaderStructuredCell.GetType().GetProperty("TableCellColumnName")?.GetValue(firstParsedHeaderStructuredCell) ?? string.Empty);
        Array parsedAllStructuredCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["SalesTotalsTable[#All]"]) ?? throw new InvalidOperationException("Expected parsed all-row structured-reference range cells."));
        object firstParsedAllStructuredCell = parsedAllStructuredCells.GetValue(0) ?? throw new InvalidOperationException("Expected first all-row structured-reference range cell.");
        object lastParsedAllStructuredCell = parsedAllStructuredCells.GetValue(parsedAllStructuredCells.Length - 1) ?? throw new InvalidOperationException("Expected last all-row structured-reference range cell.");
        TestAssert.True(parsedAllStructuredCells.Length == 8, "Expected all-row structured reference to hydrate header, data, and totals rows.");
        TestAssert.Equal("Sheet1!A1:B4", (string?)firstParsedAllStructuredCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedAllStructuredCell) ?? string.Empty);
        TestAssert.True((bool?)firstParsedAllStructuredCell.GetType().GetProperty("TableHeaderRow")?.GetValue(firstParsedAllStructuredCell) == true, "Expected all-row structured reference to preserve header row role.");
        TestAssert.True((bool?)lastParsedAllStructuredCell.GetType().GetProperty("TableTotalsRow")?.GetValue(lastParsedAllStructuredCell) == true, "Expected all-row structured reference to preserve totals row role.");
        TestAssert.True((int?)lastParsedAllStructuredCell.GetType().GetProperty("TableRowIndex")?.GetValue(lastParsedAllStructuredCell) == 3, "Expected all-row structured reference totals row index to survive resolution.");
        TestAssert.Equal("Amount", (string?)lastParsedAllStructuredCell.GetType().GetProperty("TableCellColumnName")?.GetValue(lastParsedAllStructuredCell) ?? string.Empty);
        Array parsedColumnSpanCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["SalesTable[[Region]:[Amount]]"]) ?? throw new InvalidOperationException("Expected parsed structured-reference column-span range cells."));
        object firstParsedColumnSpanCell = parsedColumnSpanCells.GetValue(0) ?? throw new InvalidOperationException("Expected first structured-reference column-span range cell.");
        object secondParsedColumnSpanCell = parsedColumnSpanCells.GetValue(1) ?? throw new InvalidOperationException("Expected second structured-reference column-span range cell.");
        TestAssert.True(parsedColumnSpanCells.Length == 6, "Expected structured-reference column span to hydrate the selected data-body rectangle.");
        TestAssert.Equal("Sheet1!A2:B4", (string?)firstParsedColumnSpanCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedColumnSpanCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedColumnSpanCell.GetType().GetProperty("RangeRowCount")?.GetValue(firstParsedColumnSpanCell) == 3, "Expected structured-reference column span row count to survive range projection.");
        TestAssert.True((int?)firstParsedColumnSpanCell.GetType().GetProperty("RangeColumnCount")?.GetValue(firstParsedColumnSpanCell) == 2, "Expected structured-reference column span column count to survive range projection.");
        TestAssert.True((int?)secondParsedColumnSpanCell.GetType().GetProperty("RangeColumnIndex")?.GetValue(secondParsedColumnSpanCell) == 1, "Expected structured-reference column span to preserve range-local column coordinates.");
        TestAssert.True((int?)secondParsedColumnSpanCell.GetType().GetProperty("SheetColumn")?.GetValue(secondParsedColumnSpanCell) == 2, "Expected structured-reference column span to preserve absolute worksheet column coordinates.");
        TestAssert.Equal("Region", (string?)firstParsedColumnSpanCell.GetType().GetProperty("TableCellColumnName")?.GetValue(firstParsedColumnSpanCell) ?? string.Empty);
        TestAssert.Equal("Amount", (string?)secondParsedColumnSpanCell.GetType().GetProperty("TableCellColumnName")?.GetValue(secondParsedColumnSpanCell) ?? string.Empty);
        TestAssert.True((int?)secondParsedColumnSpanCell.GetType().GetProperty("TableCellColumnId")?.GetValue(secondParsedColumnSpanCell) == 2, "Expected structured-reference column span to preserve each cell's table column id.");
        TestAssert.Equal("SalesTable", (string?)firstParsedColumnSpanCell.GetType().GetProperty("TableName")?.GetValue(firstParsedColumnSpanCell) ?? string.Empty);
        TestAssert.Equal(string.Empty, (string?)firstParsedColumnSpanCell.GetType().GetProperty("TableColumnName")?.GetValue(firstParsedColumnSpanCell) ?? string.Empty);
        TestAssert.True(firstParsedColumnSpanCell.GetType().GetProperty("TableColumnId")?.GetValue(firstParsedColumnSpanCell) is null, "Expected multi-column structured references to keep single column id empty.");
        TestAssert.Equal("Region", (string?)firstParsedColumnSpanCell.GetType().GetProperty("TableFirstColumnName")?.GetValue(firstParsedColumnSpanCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedColumnSpanCell.GetType().GetProperty("TableFirstColumnId")?.GetValue(firstParsedColumnSpanCell) == 1, "Expected structured-reference column span first id to survive resolution.");
        TestAssert.Equal("Amount", (string?)firstParsedColumnSpanCell.GetType().GetProperty("TableLastColumnName")?.GetValue(firstParsedColumnSpanCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedColumnSpanCell.GetType().GetProperty("TableLastColumnId")?.GetValue(firstParsedColumnSpanCell) == 2, "Expected structured-reference column span last id to survive resolution.");
        Array parsedEscapedStructuredCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["SalesTable[Quoted ']' Amount]"]) ?? throw new InvalidOperationException("Expected parsed escaped structured-reference range cells."));
        object firstParsedEscapedStructuredCell = parsedEscapedStructuredCells.GetValue(0) ?? throw new InvalidOperationException("Expected first escaped structured-reference range cell.");
        TestAssert.Equal("Sheet1!C2:C4", (string?)firstParsedEscapedStructuredCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedEscapedStructuredCell) ?? string.Empty);
        TestAssert.Equal("Quoted ] Amount", (string?)firstParsedEscapedStructuredCell.GetType().GetProperty("TableColumnName")?.GetValue(firstParsedEscapedStructuredCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedEscapedStructuredCell.GetType().GetProperty("TableColumnId")?.GetValue(firstParsedEscapedStructuredCell) == 3, "Expected escaped structured-reference source column id to survive resolution.");
        Array parsedTotalsStructuredCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["SalesTotalsTable[[#Totals],[Amount]]"]) ?? throw new InvalidOperationException("Expected parsed totals structured-reference range cells."));
        object firstParsedTotalsStructuredCell = parsedTotalsStructuredCells.GetValue(0) ?? throw new InvalidOperationException("Expected first totals structured-reference range cell.");
        TestAssert.True(parsedTotalsStructuredCells.Length == 1, "Expected totals-only structured reference to hydrate only the declared totals row.");
        TestAssert.Equal("Sheet1!B4:B4", (string?)firstParsedTotalsStructuredCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedTotalsStructuredCell) ?? string.Empty);
        TestAssert.Equal("Amount", (string?)firstParsedTotalsStructuredCell.GetType().GetProperty("TableColumnName")?.GetValue(firstParsedTotalsStructuredCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedTotalsStructuredCell.GetType().GetProperty("TableColumnId")?.GetValue(firstParsedTotalsStructuredCell) == 2, "Expected totals structured-reference source column id to survive resolution.");
        TestAssert.True((int?)firstParsedTotalsStructuredCell.GetType().GetProperty("TableRowIndex")?.GetValue(firstParsedTotalsStructuredCell) == 3, "Expected totals structured-reference table row index to survive resolution.");
        TestAssert.True((bool?)firstParsedTotalsStructuredCell.GetType().GetProperty("TableTotalsRow")?.GetValue(firstParsedTotalsStructuredCell) == true, "Expected totals structured-reference row role to survive resolution.");
        TestAssert.True((bool?)firstParsedTotalsStructuredCell.GetType().GetProperty("TableDataRow")?.GetValue(firstParsedTotalsStructuredCell) == false, "Expected totals structured-reference not to be classified as a data row.");
        TestAssert.Equal("sum", (string?)firstParsedTotalsStructuredCell.GetType().GetProperty("TableCellTotalsRowFunction")?.GetValue(firstParsedTotalsStructuredCell) ?? string.Empty);
        TestAssert.Equal("SUBTOTAL(109,[Amount])", (string?)firstParsedTotalsStructuredCell.GetType().GetProperty("TableCellTotalsRowFormula")?.GetValue(firstParsedTotalsStructuredCell) ?? string.Empty);
        Array parsedBlankCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["Sheet1!$C$2:$C$2"]) ?? throw new InvalidOperationException("Expected parsed blank workbook range cell."));
        object parsedBlankCell = parsedBlankCells.GetValue(0) ?? throw new InvalidOperationException("Expected blank workbook range cell.");
        TestAssert.True((bool?)parsedBlankCell.GetType().GetProperty("HasCell")?.GetValue(parsedBlankCell) == true, "Expected styled formula blank cell to remain a physical cell.");
        TestAssert.True((bool?)parsedBlankCell.GetType().GetProperty("HasValue")?.GetValue(parsedBlankCell) == false, "Expected formula blank cell without cached value to preserve missing value state.");
        TestAssert.True((bool?)parsedBlankCell.GetType().GetProperty("HasValueElement")?.GetValue(parsedBlankCell) == false, "Expected formula blank cell without cached <v> to preserve value-element absence.");
        TestAssert.Equal("Blank", parsedBlankCell.GetType().GetProperty("ValueKind")?.GetValue(parsedBlankCell)?.ToString() ?? string.Empty);
        TestAssert.Equal("NA()", (string?)parsedBlankCell.GetType().GetProperty("Formula")?.GetValue(parsedBlankCell) ?? string.Empty);
        TestAssert.Equal("array", (string?)parsedBlankCell.GetType().GetProperty("FormulaType")?.GetValue(parsedBlankCell) ?? string.Empty);
        var parsedBlankFormulaAttributes = (IReadOnlyDictionary<string, string>?)parsedBlankCell.GetType().GetProperty("FormulaAttributes")?.GetValue(parsedBlankCell) ?? throw new InvalidOperationException("Expected blank formula attributes.");
        TestAssert.True(parsedBlankFormulaAttributes.TryGetValue("ref", out string? blankFormulaReference) && blankFormulaReference == "C2:C2", "Expected array formula reference to survive workbook parsing.");
        TestAssert.True(parsedBlankFormulaAttributes.TryGetValue("ca", out string? blankFormulaCalculateAlways) && blankFormulaCalculateAlways == "1", "Expected formula calculation attribute to survive workbook parsing.");
        object secondParsedCell = parsedCells.GetValue(1) ?? throw new InvalidOperationException("Expected second parsed workbook range cell.");
        TestAssert.True((bool?)rowHiddenProperty.GetValue(secondParsedCell) == true, "Expected hidden worksheet row to survive workbook parsing.");
        object thirdParsedCell = parsedCells.GetValue(2) ?? throw new InvalidOperationException("Expected third parsed workbook range cell.");
        System.Reflection.PropertyInfo formulaProperty = thirdParsedCell.GetType().GetProperty("Formula") ?? throw new InvalidOperationException("Expected range-cell formula metadata.");
        TestAssert.Equal("B2-B3", (string?)formulaProperty.GetValue(thirdParsedCell) ?? string.Empty);
        TestAssert.Equal("shared", (string?)thirdParsedCell.GetType().GetProperty("FormulaType")?.GetValue(thirdParsedCell) ?? string.Empty);
        var thirdFormulaAttributes = (IReadOnlyDictionary<string, string>?)thirdParsedCell.GetType().GetProperty("FormulaAttributes")?.GetValue(thirdParsedCell) ?? throw new InvalidOperationException("Expected cached formula attributes.");
        TestAssert.True(thirdFormulaAttributes.TryGetValue("si", out string? sharedFormulaIndex) && sharedFormulaIndex == "7", "Expected shared formula index to survive workbook parsing.");
        TestAssert.True(thirdFormulaAttributes.TryGetValue("ref", out string? sharedFormulaReference) && sharedFormulaReference == "B4:B4", "Expected shared formula reference to survive workbook parsing.");
        System.Reflection.PropertyInfo definedNamesProperty = workbookType.GetProperty("DefinedNames") ?? throw new InvalidOperationException("Expected workbook defined names.");
        var definedNames = (System.Collections.Generic.IReadOnlyDictionary<string, string>?)definedNamesProperty.GetValue(parsedWorkbook) ?? throw new InvalidOperationException("Expected parsed defined names.");
        TestAssert.True(definedNames.TryGetValue("SalesValues", out string? salesValuesFormula) && salesValuesFormula == "Sheet1!$B$2:$B$4", "Expected workbook-level defined name to survive parsing.");
        TestAssert.True(definedNames.TryGetValue("SalesUnionValues", out string? salesUnionValuesFormula) && salesUnionValuesFormula == "Sheet1!$B$2:$B$2,$B$4:$B$4", "Expected workbook-level multi-area defined name to survive parsing.");
        TestAssert.True(!definedNames.ContainsKey("SheetLocalValues"), "Expected sheet-local defined names to stay out of workbook-level name resolution.");
        System.Reflection.PropertyInfo definedNameRecordsProperty = workbookType.GetProperty("DefinedNameRecords") ?? throw new InvalidOperationException("Expected workbook defined-name records.");
        var definedNameRecords = ((System.Collections.IEnumerable?)definedNameRecordsProperty.GetValue(parsedWorkbook) ?? throw new InvalidOperationException("Expected parsed defined-name records.")).Cast<object>().ToArray();
        object localDefinedName = definedNameRecords.Single(record => (string?)record.GetType().GetProperty("Name")?.GetValue(record) == "SheetLocalValues");
        TestAssert.True((int?)localDefinedName.GetType().GetProperty("LocalSheetId")?.GetValue(localDefinedName) == 0, "Expected sheet-local defined-name scope to survive workbook parsing.");
        TestAssert.Equal("Sheet1", (string?)localDefinedName.GetType().GetProperty("SheetName")?.GetValue(localDefinedName) ?? string.Empty);
        Array parsedLocalDefinedNameCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["Sheet1!SheetLocalValues"]) ?? throw new InvalidOperationException("Expected parsed sheet-local defined-name range cells."));
        object firstParsedLocalDefinedNameCell = parsedLocalDefinedNameCells.GetValue(0) ?? throw new InvalidOperationException("Expected first sheet-local defined-name range cell.");
        TestAssert.Equal("Sheet1!SheetLocalValues", (string?)firstParsedLocalDefinedNameCell.GetType().GetProperty("SourceFormula")?.GetValue(firstParsedLocalDefinedNameCell) ?? string.Empty);
        TestAssert.Equal("Sheet1!$B$2:$B$4", (string?)firstParsedLocalDefinedNameCell.GetType().GetProperty("ResolvedFormula")?.GetValue(firstParsedLocalDefinedNameCell) ?? string.Empty);
        TestAssert.Equal("DefinedName", firstParsedLocalDefinedNameCell.GetType().GetProperty("SourceKind")?.GetValue(firstParsedLocalDefinedNameCell)?.ToString() ?? string.Empty);
        TestAssert.Equal("SheetLocalValues", (string?)firstParsedLocalDefinedNameCell.GetType().GetProperty("DefinedName")?.GetValue(firstParsedLocalDefinedNameCell) ?? string.Empty);
        TestAssert.Equal("Sheet1", (string?)firstParsedLocalDefinedNameCell.GetType().GetProperty("DefinedNameSheetName")?.GetValue(firstParsedLocalDefinedNameCell) ?? string.Empty);
        TestAssert.True((int?)firstParsedLocalDefinedNameCell.GetType().GetProperty("DefinedNameLocalSheetId")?.GetValue(firstParsedLocalDefinedNameCell) == 0, "Expected sheet-local defined-name source scope to survive range resolution.");
        Array parsedBareLocalDefinedNameCells = (Array)(readRangeCells.Invoke(parsedWorkbook, ["SheetLocalValues"]) ?? throw new InvalidOperationException("Expected bare sheet-local defined-name lookup to return an empty range."));
        TestAssert.True(parsedBareLocalDefinedNameCells.Length == 0, "Expected bare sheet-local names to remain unresolved without an explicit sheet scope.");
        Array parsedUnionDefinedNameValues = (Array)(readNumericRange.Invoke(parsedWorkbook, ["SalesUnionValues"]) ?? throw new InvalidOperationException("Expected typed multi-area defined-name numeric values."));
        TestAssert.True(parsedUnionDefinedNameValues.Length == 2, "Expected multi-area defined-name formula to hydrate both workbook areas.");
        object secondParsedUnionDefinedNameValue = parsedUnionDefinedNameValues.GetValue(1) ?? throw new InvalidOperationException("Expected second typed multi-area defined-name numeric value.");
        TestAssert.Equal(1.4d, (double?)numericValueProperty.GetValue(secondParsedUnionDefinedNameValue) ?? 0d);
        object secondParsedUnionDefinedNameCell = numericCellProperty.GetValue(secondParsedUnionDefinedNameValue) ?? throw new InvalidOperationException("Expected typed multi-area defined-name range cell.");
        TestAssert.Equal("DefinedName", secondParsedUnionDefinedNameCell.GetType().GetProperty("SourceKind")?.GetValue(secondParsedUnionDefinedNameCell)?.ToString() ?? string.Empty);
        TestAssert.Equal("SalesUnionValues", (string?)secondParsedUnionDefinedNameCell.GetType().GetProperty("DefinedName")?.GetValue(secondParsedUnionDefinedNameCell) ?? string.Empty);
        TestAssert.True((int?)secondParsedUnionDefinedNameCell.GetType().GetProperty("RangeAreaIndex")?.GetValue(secondParsedUnionDefinedNameCell) == 1, "Expected multi-area defined-name cell to preserve its range-area index.");
        System.Reflection.PropertyInfo calculationProperty = workbookType.GetProperty("Calculation") ?? throw new InvalidOperationException("Expected workbook calculation metadata.");
        object calculation = calculationProperty.GetValue(parsedWorkbook) ?? throw new InvalidOperationException("Expected parsed workbook calculation metadata.");
        System.Reflection.PropertyInfo calculationModeProperty = calculation.GetType().GetProperty("CalculationMode") ?? throw new InvalidOperationException("Expected workbook calculation mode.");
        TestAssert.Equal("auto", (string?)calculationModeProperty.GetValue(calculation) ?? string.Empty);
        System.Reflection.PropertyInfo fullCalculationOnLoadProperty = calculation.GetType().GetProperty("FullCalculationOnLoad") ?? throw new InvalidOperationException("Expected full-calc-on-load metadata.");
        TestAssert.True((bool?)fullCalculationOnLoadProperty.GetValue(calculation) == true, "Expected workbook fullCalcOnLoad metadata to survive parsing.");
        System.Reflection.PropertyInfo tablesProperty = workbookType.GetProperty("Tables") ?? throw new InvalidOperationException("Expected workbook tables.");
        object tables = tablesProperty.GetValue(parsedWorkbook) ?? throw new InvalidOperationException("Expected parsed workbook tables.");
        System.Reflection.PropertyInfo tableCountProperty = tables.GetType().GetProperty("Count") ?? throw new InvalidOperationException("Expected workbook table count.");
        TestAssert.True((int?)tableCountProperty.GetValue(tables) == 4, "Expected table names and display names to index the parsed workbook tables.");
        System.Reflection.PropertyInfo tableValuesProperty = tables.GetType().GetProperty("Values") ?? throw new InvalidOperationException("Expected workbook table values.");
        object salesTable = (((System.Collections.IEnumerable?)tableValuesProperty.GetValue(tables)) ?? throw new InvalidOperationException("Expected workbook table values."))
            .Cast<object>()
            .First(table => (string?)table.GetType().GetProperty("Name")?.GetValue(table) == "SalesTable");
        TestAssert.Equal(1, (int?)salesTable.GetType().GetProperty("HeaderRowCount")?.GetValue(salesTable) ?? 0);
        TestAssert.Equal("A1:C4", (string?)salesTable.GetType().GetProperty("AutoFilterReference")?.GetValue(salesTable) ?? string.Empty);
        object filterColumnIds = salesTable.GetType().GetProperty("FilterColumnIds")?.GetValue(salesTable) ?? throw new InvalidOperationException("Expected table filter column ids.");
        TestAssert.True(((System.Collections.IEnumerable)filterColumnIds).Cast<int>().SequenceEqual([0, 1]), "Expected table filter column ids to survive workbook parsing.");
        object[] filterColumns = (((System.Collections.IEnumerable?)salesTable.GetType().GetProperty("FilterColumns")?.GetValue(salesTable)) ?? throw new InvalidOperationException("Expected table filter column records.")).Cast<object>().ToArray();
        TestAssert.True(filterColumns.Length == 2, "Expected table filter column records to survive workbook parsing.");
        TestAssert.Equal("filters", (string?)filterColumns[0].GetType().GetProperty("FilterKind")?.GetValue(filterColumns[0]) ?? string.Empty);
        object filterValues = filterColumns[0].GetType().GetProperty("FilterValues")?.GetValue(filterColumns[0]) ?? throw new InvalidOperationException("Expected filter values.");
        TestAssert.True(((System.Collections.IEnumerable)filterValues).Cast<string>().Single() == "North", "Expected table filter value to survive workbook parsing.");
        TestAssert.Equal("customFilters", (string?)filterColumns[1].GetType().GetProperty("FilterKind")?.GetValue(filterColumns[1]) ?? string.Empty);
        object customFilters = filterColumns[1].GetType().GetProperty("CustomFilters")?.GetValue(filterColumns[1]) ?? throw new InvalidOperationException("Expected custom filters.");
        object customFilter = ((System.Collections.IEnumerable)customFilters).Cast<object>().Single();
        TestAssert.Equal("greaterThan", (string?)customFilter.GetType().GetProperty("Operator")?.GetValue(customFilter) ?? string.Empty);
        TestAssert.Equal("3", (string?)customFilter.GetType().GetProperty("Value")?.GetValue(customFilter) ?? string.Empty);
        object[] salesTableColumns = (((System.Collections.IEnumerable?)salesTable.GetType().GetProperty("Columns")?.GetValue(salesTable)) ?? throw new InvalidOperationException("Expected table column records.")).Cast<object>().ToArray();
        TestAssert.True(salesTableColumns.Length == 3, "Expected table column records to survive workbook parsing.");
        object amountColumn = salesTableColumns[1];
        TestAssert.True((int?)amountColumn.GetType().GetProperty("Id")?.GetValue(amountColumn) == 2, "Expected table column id metadata to survive workbook parsing.");
        TestAssert.Equal("Amount", (string?)amountColumn.GetType().GetProperty("Name")?.GetValue(amountColumn) ?? string.Empty);
        TestAssert.Equal("B2", (string?)amountColumn.GetType().GetProperty("CalculatedColumnFormula")?.GetValue(amountColumn) ?? string.Empty);
        object escapedColumn = salesTableColumns[2];
        TestAssert.True((int?)escapedColumn.GetType().GetProperty("Id")?.GetValue(escapedColumn) == 3, "Expected escaped table column id metadata to survive workbook parsing.");
        TestAssert.Equal("Quoted ] Amount", (string?)escapedColumn.GetType().GetProperty("Name")?.GetValue(escapedColumn) ?? string.Empty);
        var chartXml = XDocument.Parse("""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:doughnutChart><c:ser>
                <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
                <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val="0"/></c:numCache></c:numRef></c:val>
              </c:ser></c:doughnutChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);
        var hydrate = typeof(PptxRenderer).GetMethod(
            "HydrateChartReferenceCaches",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected cache hydration helper.");

        hydrate.Invoke(null, [workbook, chartXml]);

        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement numCache = chartXml.Descendants(c + "numCache").Single();
        TestAssert.Equal("0.0", numCache.Element(c + "formatCode")?.Value ?? string.Empty);
        TestAssert.Equal("3", numCache.Element(c + "ptCount")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("0", numCache.Elements(c + "pt").First().Attribute("idx")?.Value ?? string.Empty);
        TestAssert.Equal("2", numCache.Elements(c + "pt").Last().Attribute("idx")?.Value ?? string.Empty);
        XElement strCache = chartXml.Descendants(c + "strCache").Single();
        TestAssert.Equal("3", strCache.Element(c + "ptCount")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("0", strCache.Elements(c + "pt").First().Attribute("idx")?.Value ?? string.Empty);
        TestAssert.Equal("2", strCache.Elements(c + "pt").Last().Attribute("idx")?.Value ?? string.Empty);
        var definedNameChartXml = XDocument.Parse("""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:doughnutChart><c:ser>
                <c:cat><c:strRef><c:f>SalesLabels</c:f></c:strRef></c:cat>
                <c:val><c:numRef><c:f>SalesValues</c:f></c:numRef></c:val>
              </c:ser></c:doughnutChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        hydrate.Invoke(null, [parsedWorkbook, definedNameChartXml]);

        XElement definedNameNumCache = definedNameChartXml.Descendants(c + "numCache").Single();
        TestAssert.Equal("3", definedNameNumCache.Element(c + "ptCount")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("8.2", definedNameNumCache.Elements(c + "pt").First().Element(c + "v")?.Value ?? string.Empty);
        XElement definedNameStrCache = definedNameChartXml.Descendants(c + "strCache").Single();
        TestAssert.Equal("3", definedNameStrCache.Element(c + "ptCount")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("North", definedNameStrCache.Elements(c + "pt").First().Element(c + "v")?.Value ?? string.Empty);
        var tableChartXml = XDocument.Parse("""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:doughnutChart><c:ser>
                <c:cat><c:strRef><c:f>SalesTable[Region]</c:f></c:strRef></c:cat>
                <c:val><c:numRef><c:f>SalesTable[Amount]</c:f></c:numRef></c:val>
              </c:ser></c:doughnutChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        hydrate.Invoke(null, [parsedWorkbook, tableChartXml]);

        XElement tableNumCache = tableChartXml.Descendants(c + "numCache").Single();
        TestAssert.Equal("3", tableNumCache.Element(c + "ptCount")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("8.2", tableNumCache.Elements(c + "pt").First().Element(c + "v")?.Value ?? string.Empty);
        XElement tableStrCache = tableChartXml.Descendants(c + "strCache").Single();
        TestAssert.Equal("3", tableStrCache.Element(c + "ptCount")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("North", tableStrCache.Elements(c + "pt").First().Element(c + "v")?.Value ?? string.Empty);
        var totalsTableChartXml = XDocument.Parse("""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:doughnutChart><c:ser>
                <c:cat><c:strRef><c:f>SalesTotalsTable[Region]</c:f></c:strRef></c:cat>
                <c:val><c:numRef><c:f>SalesTotalsTable[Amount]</c:f></c:numRef></c:val>
              </c:ser></c:doughnutChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        hydrate.Invoke(null, [parsedWorkbook, totalsTableChartXml]);

        XElement totalsTableNumCache = totalsTableChartXml.Descendants(c + "numCache").Single();
        TestAssert.Equal("2", totalsTableNumCache.Element(c + "ptCount")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("3.2", totalsTableNumCache.Elements(c + "pt").Last().Element(c + "v")?.Value ?? string.Empty);
    }

    public static void PptxChartIndexedVectorsPreserveWorkbookSidecarPoints()
    {
        Type workbookType = typeof(PptxRenderer).GetNestedType(
            "ChartWorkbookData",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart workbook data type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["B2"] = "8.2",
                ["B3"] = "3.4"
            }
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets]) ?? throw new InvalidOperationException("Expected workbook instance.");
        var source = new PptxSceneChartDataSource(
            "Sheet1!$B$2:$B$3",
            PptxSceneChartDataSourceReferenceKind.NumberReference,
            "numRef",
            PptxSceneChartDataSourceCacheKind.NumberCache,
            "numCache",
            HasCachedPoints: true);
        System.Reflection.MethodInfo buildVector = typeof(PptxRenderer)
            .GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(method => method.Name == "BuildChartIndexedNumberVector" && method.GetParameters().Length == 6);

        object vector = buildVector.Invoke(
            null,
            [new[] { 99d }, Array.Empty<PptxSceneChartNumberPoint>(), 1, "General", source, workbook]) ?? throw new InvalidOperationException("Expected indexed chart vector.");

        object[] activePoints = (((System.Collections.IEnumerable?)vector.GetType().GetProperty("Points")?.GetValue(vector)) ?? throw new InvalidOperationException("Expected active chart points.")).Cast<object>().ToArray();
        object[] workbookPoints = (((System.Collections.IEnumerable?)vector.GetType().GetProperty("WorkbookPoints")?.GetValue(vector)) ?? throw new InvalidOperationException("Expected workbook sidecar points.")).Cast<object>().ToArray();
        TestAssert.True(activePoints.Length == 1, "Expected existing chart-cache point projection to remain active.");
        TestAssert.True((double?)activePoints[0].GetType().GetProperty("Value")?.GetValue(activePoints[0]) == 99d, "Expected chart-cache value to remain the rendered point value.");
        object[] densePoints = (((System.Collections.IEnumerable?)vector.GetType().GetMethod("DensePoints")?.Invoke(vector, [])) ?? throw new InvalidOperationException("Expected dense active point projection.")).Cast<object>().ToArray();
        TestAssert.True(densePoints.Length == 1, "Expected dense point projection to preserve rendered point count.");
        TestAssert.True((double?)densePoints[0].GetType().GetProperty("Value")?.GetValue(densePoints[0]) == 99d, "Expected dense point projection to preserve rendered cache point records.");
        TestAssert.True(workbookPoints.Length == 2, "Expected workbook sidecar points to preserve workbook source values next to cached chart points.");
        TestAssert.True((int?)workbookPoints[0].GetType().GetProperty("Index")?.GetValue(workbookPoints[0]) == 0, "Expected workbook sidecar point index to preserve range order.");
        TestAssert.True((double?)workbookPoints[0].GetType().GetProperty("Value")?.GetValue(workbookPoints[0]) == 8.2d, "Expected workbook sidecar point to preserve workbook numeric value.");
        object workbookCell = workbookPoints[0].GetType().GetProperty("WorkbookCell")?.GetValue(workbookPoints[0]) ?? throw new InvalidOperationException("Expected workbook sidecar point cell provenance.");
        TestAssert.Equal("B2", (string?)workbookCell.GetType().GetProperty("Reference")?.GetValue(workbookCell) ?? string.Empty);
        var noCacheSource = source with { HasCachedPoints = false };
        object noCacheVector = buildVector.Invoke(
            null,
            [Array.Empty<double>(), Array.Empty<PptxSceneChartNumberPoint>(), null, "General", noCacheSource, workbook]) ?? throw new InvalidOperationException("Expected formula-only indexed chart vector.");
        object[] noCacheActivePoints = (((System.Collections.IEnumerable?)noCacheVector.GetType().GetProperty("Points")?.GetValue(noCacheVector)) ?? throw new InvalidOperationException("Expected formula-only active chart points.")).Cast<object>().ToArray();
        object[] noCacheWorkbookPoints = (((System.Collections.IEnumerable?)noCacheVector.GetType().GetProperty("WorkbookPoints")?.GetValue(noCacheVector)) ?? throw new InvalidOperationException("Expected formula-only workbook sidecar points.")).Cast<object>().ToArray();
        object[] noCacheDensePoints = (((System.Collections.IEnumerable?)noCacheVector.GetType().GetMethod("DensePoints")?.Invoke(noCacheVector, [])) ?? throw new InvalidOperationException("Expected formula-only dense active point projection.")).Cast<object>().ToArray();
        TestAssert.True(noCacheActivePoints.Length == 0, "Expected formula-only workbook numeric values to remain absent from cache-owned active points.");
        TestAssert.True(noCacheDensePoints.Length == 0, "Expected formula-only workbook numeric values to stay out of dense rendered point slots when chart caches are absent.");
        TestAssert.True(noCacheWorkbookPoints.Length == 2, "Expected formula-only workbook numeric values to remain available as sidecar points.");
        System.Reflection.MethodInfo buildPieSlices = typeof(PptxRenderer).GetMethod(
            "BuildChartIndexedPieSlices",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected indexed pie-slice builder.");
        object[] pieSlices = (((System.Collections.IEnumerable?)buildPieSlices.Invoke(null, [vector])) ?? throw new InvalidOperationException("Expected indexed pie slices.")).Cast<object>().ToArray();
        TestAssert.True(pieSlices.Length == 1, "Expected stale positive cache value to remain an active pie slice.");
        object slicePoint = pieSlices[0].GetType().GetProperty("Point")?.GetValue(pieSlices[0]) ?? throw new InvalidOperationException("Expected pie slice to preserve its active point record.");
        object sliceWorkbookPoint = pieSlices[0].GetType().GetProperty("WorkbookPoint")?.GetValue(pieSlices[0]) ?? throw new InvalidOperationException("Expected pie slice to preserve matching workbook sidecar point.");
        TestAssert.True((double?)slicePoint.GetType().GetProperty("Value")?.GetValue(slicePoint) == 99d, "Expected pie slice active point to preserve the rendered cache value.");
        TestAssert.True((double?)sliceWorkbookPoint.GetType().GetProperty("Value")?.GetValue(sliceWorkbookPoint) == 8.2d, "Expected pie slice workbook sidecar point to preserve the workbook source value.");
        System.Reflection.MethodInfo buildRadarSeries = typeof(PptxRenderer).GetMethod(
            "BuildRadarSeries",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected radar-series projection.");
        Array typedNumberVectors = Array.CreateInstance(vector.GetType(), 1);
        typedNumberVectors.SetValue(vector, 0);
        object[] radarSeries = (((System.Collections.IEnumerable?)buildRadarSeries.Invoke(null, [typedNumberVectors])) ?? throw new InvalidOperationException("Expected radar series.")).Cast<object>().ToArray();
        TestAssert.True(radarSeries.Length == 1, "Expected positive cache value to remain an active radar series.");
        object[] radarPoints = (((System.Collections.IEnumerable?)radarSeries[0].GetType().GetProperty("Points")?.GetValue(radarSeries[0])) ?? throw new InvalidOperationException("Expected radar series to preserve active point records.")).Cast<object>().ToArray();
        TestAssert.True(radarPoints.Length == 1, "Expected radar series to preserve one active point record.");
        TestAssert.True((int?)radarPoints[0].GetType().GetProperty("Index")?.GetValue(radarPoints[0]) == 0, "Expected radar active point to preserve the source point index.");
        TestAssert.True((double?)radarPoints[0].GetType().GetProperty("Value")?.GetValue(radarPoints[0]) == 99d, "Expected radar active point to preserve the rendered cache value.");
        object radarSource = radarSeries[0].GetType().GetProperty("Source")?.GetValue(radarSeries[0]) ?? throw new InvalidOperationException("Expected radar series to preserve its source vector.");
        object[] radarWorkbookPoints = (((System.Collections.IEnumerable?)radarSource.GetType().GetProperty("WorkbookPoints")?.GetValue(radarSource)) ?? throw new InvalidOperationException("Expected radar source vector workbook sidecar points.")).Cast<object>().ToArray();
        TestAssert.True((double?)radarWorkbookPoints[0].GetType().GetProperty("Value")?.GetValue(radarWorkbookPoints[0]) == 8.2d, "Expected radar series source to preserve workbook sidecar values.");
        object sparseRadarVector = buildVector.Invoke(
            null,
            [Array.Empty<double>(), new[] { new PptxSceneChartNumberPoint(2, "2", true, 44d, "44", true) }, 3, "General", source, null]) ?? throw new InvalidOperationException("Expected sparse indexed chart vector.");
        Array sparseRadarVectors = Array.CreateInstance(vector.GetType(), 1);
        sparseRadarVectors.SetValue(sparseRadarVector, 0);
        object[] sparseRadarSeries = (((System.Collections.IEnumerable?)buildRadarSeries.Invoke(null, [sparseRadarVectors])) ?? throw new InvalidOperationException("Expected sparse radar series.")).Cast<object>().ToArray();
        object[] sparseRadarPoints = (((System.Collections.IEnumerable?)sparseRadarSeries[0].GetType().GetProperty("Points")?.GetValue(sparseRadarSeries[0])) ?? throw new InvalidOperationException("Expected sparse radar points.")).Cast<object>().ToArray();
        TestAssert.True(sparseRadarPoints.Length == 3, "Expected radar geometry projection to preserve sparse point slots.");
        TestAssert.True(sparseRadarPoints[0] is null && sparseRadarPoints[1] is null, "Expected missing radar point slots to remain explicit gaps.");
        TestAssert.True((int?)sparseRadarPoints[2].GetType().GetProperty("Index")?.GetValue(sparseRadarPoints[2]) == 2, "Expected sparse radar point to keep its source index.");
        Type scatterSeriesType = typeof(PptxRenderer).GetNestedType(
            "ChartIndexedScatterSeries",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected indexed scatter-series type.");
        object scatterSeries = Activator.CreateInstance(scatterSeriesType, [vector, vector, vector, true]) ?? throw new InvalidOperationException("Expected indexed scatter-series instance.");
        System.Reflection.MethodInfo buildScatterSeries = typeof(PptxRenderer).GetMethod(
            "BuildScatterSeries",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected scatter-series projection.");
        object renderedScatter = buildScatterSeries.Invoke(null, [scatterSeries]) ?? throw new InvalidOperationException("Expected rendered scatter-series.");
        object renderedScatterSource = renderedScatter.GetType().GetProperty("Source")?.GetValue(renderedScatter) ?? throw new InvalidOperationException("Expected rendered scatter-series to preserve its indexed source vectors.");
        TestAssert.True((bool?)renderedScatterSource.GetType().GetProperty("ReadBubbleSize")?.GetValue(renderedScatterSource) == true, "Expected rendered scatter-series source to preserve bubble-size channel ownership.");
        object[] scatterPoints = (((System.Collections.IEnumerable?)renderedScatter.GetType().GetProperty("Points")?.GetValue(renderedScatter)) ?? throw new InvalidOperationException("Expected rendered scatter points.")).Cast<object>().ToArray();
        TestAssert.True(scatterPoints.Length == 1, "Expected stale positive cache value to remain an active scatter point.");
        object scatterXPoint = scatterPoints[0].GetType().GetProperty("XPoint")?.GetValue(scatterPoints[0]) ?? throw new InvalidOperationException("Expected scatter point to preserve its active X point.");
        object scatterXWorkbookPoint = scatterPoints[0].GetType().GetProperty("XWorkbookPoint")?.GetValue(scatterPoints[0]) ?? throw new InvalidOperationException("Expected scatter point to preserve matching workbook X point.");
        TestAssert.True((double?)scatterXPoint.GetType().GetProperty("Value")?.GetValue(scatterXPoint) == 99d, "Expected scatter active point to preserve the rendered cache value.");
        TestAssert.True((double?)scatterXWorkbookPoint.GetType().GetProperty("Value")?.GetValue(scatterXWorkbookPoint) == 8.2d, "Expected scatter workbook sidecar point to preserve the workbook source value.");
        object sparseScatterXVector = buildVector.Invoke(
            null,
            [Array.Empty<double>(), new[] { new PptxSceneChartNumberPoint(2, "2", true, 12d, "12", true) }, 3, "General", source, null]) ?? throw new InvalidOperationException("Expected sparse scatter X vector.");
        object sparseScatterYVector = buildVector.Invoke(
            null,
            [Array.Empty<double>(), new[] { new PptxSceneChartNumberPoint(2, "2", true, 34d, "34", true) }, 3, "General", source, null]) ?? throw new InvalidOperationException("Expected sparse scatter Y vector.");
        object sparseScatterSizeVector = buildVector.Invoke(
            null,
            [Array.Empty<double>(), new[] { new PptxSceneChartNumberPoint(2, "2", true, 5d, "5", true) }, 3, "General", source, null]) ?? throw new InvalidOperationException("Expected sparse scatter bubble-size vector.");
        object sparseScatterSeriesRecord = Activator.CreateInstance(scatterSeriesType, [sparseScatterXVector, sparseScatterYVector, sparseScatterSizeVector, true]) ?? throw new InvalidOperationException("Expected sparse indexed scatter-series instance.");
        object sparseRenderedScatter = buildScatterSeries.Invoke(null, [sparseScatterSeriesRecord]) ?? throw new InvalidOperationException("Expected sparse rendered scatter-series.");
        object[] sparseScatterPoints = (((System.Collections.IEnumerable?)sparseRenderedScatter.GetType().GetProperty("Points")?.GetValue(sparseRenderedScatter)) ?? throw new InvalidOperationException("Expected sparse rendered scatter points.")).Cast<object>().ToArray();
        TestAssert.True(sparseScatterPoints.Length == 1, "Expected sparse scatter channels to pair by source index.");
        TestAssert.True((int?)sparseScatterPoints[0].GetType().GetProperty("Index")?.GetValue(sparseScatterPoints[0]) == 2, "Expected sparse scatter point to preserve the paired source index.");
        TestAssert.True((double?)sparseScatterPoints[0].GetType().GetProperty("Size")?.GetValue(sparseScatterPoints[0]) == 5d, "Expected sparse bubble-size channel to pair by source index.");
        var blankSource = source with { Formula = "Sheet1!$B$2:$B$4" };
        object blankVector = buildVector.Invoke(
            null,
            [new[] { 99d }, Array.Empty<PptxSceneChartNumberPoint>(), 3, "General", blankSource, workbook]) ?? throw new InvalidOperationException("Expected blank-aware indexed chart vector.");
        object[] blankAwareWorkbookPoints = (((System.Collections.IEnumerable?)blankVector.GetType().GetProperty("WorkbookPoints")?.GetValue(blankVector)) ?? throw new InvalidOperationException("Expected blank-aware workbook sidecar points.")).Cast<object>().ToArray();
        TestAssert.True(blankAwareWorkbookPoints.Length == 3, "Expected workbook sidecar points to preserve missing cells inside the source range.");
        TestAssert.True(blankAwareWorkbookPoints[2].GetType().GetProperty("Value")?.GetValue(blankAwareWorkbookPoints[2]) is null, "Expected missing workbook source cells to remain nullable sidecar points.");
        object blankWorkbookCell = blankAwareWorkbookPoints[2].GetType().GetProperty("WorkbookCell")?.GetValue(blankAwareWorkbookPoints[2]) ?? throw new InvalidOperationException("Expected blank workbook cell provenance.");
        TestAssert.True((bool?)blankWorkbookCell.GetType().GetProperty("HasCell")?.GetValue(blankWorkbookCell) == false, "Expected missing workbook source cells to preserve HasCell=false.");

        using MemoryStream embeddedWorkbookStream = new(PptxTests.EmbeddedChartWorkbook());
        OoxPackage embeddedWorkbookPackage = OoxPackage.Open(embeddedWorkbookStream, CancellationToken.None);
        var readWorkbookData = typeof(PptxRenderer).GetMethod(
            "ReadWorkbookData",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected workbook reader helper.");
        object parsedWorkbook = readWorkbookData.Invoke(null, [embeddedWorkbookPackage]) ?? throw new InvalidOperationException("Expected parsed workbook data.");
        var hiddenColumnSource = source with { Formula = "Sheet1!$B$2:$B$4" };
        object hiddenColumnVector = buildVector.Invoke(
            null,
            [new[] { 99d }, Array.Empty<PptxSceneChartNumberPoint>(), 3, "General", hiddenColumnSource, parsedWorkbook]) ?? throw new InvalidOperationException("Expected hidden-column indexed chart vector.");
        object[] allHiddenColumnWorkbookPoints = (((System.Collections.IEnumerable?)hiddenColumnVector.GetType().GetMethod("WorkbookPointsForPlotVisibility")?.Invoke(hiddenColumnVector, [false])) ?? throw new InvalidOperationException("Expected unfiltered workbook points.")).Cast<object>().ToArray();
        object[] visibleHiddenColumnWorkbookPoints = (((System.Collections.IEnumerable?)hiddenColumnVector.GetType().GetMethod("WorkbookPointsForPlotVisibility")?.Invoke(hiddenColumnVector, [true])) ?? throw new InvalidOperationException("Expected visibility-filtered workbook points.")).Cast<object>().ToArray();
        TestAssert.True(allHiddenColumnWorkbookPoints.Length == 3, "Expected plot visibility projection to start from the full workbook source vector.");
        TestAssert.True(visibleHiddenColumnWorkbookPoints.Length == 0, "Expected plot-visible-only projection to exclude workbook points from a hidden source column.");
        object? hiddenWorkbookPoint = hiddenColumnVector.GetType().GetMethod("WorkbookPointForIndex")?.Invoke(hiddenColumnVector, [0]);
        TestAssert.True(hiddenWorkbookPoint is null, "Expected indexed workbook source lookup to honor the vector's plot-visible-only policy.");
        object noCacheHiddenColumnVector = buildVector.Invoke(
            null,
            [Array.Empty<double>(), Array.Empty<PptxSceneChartNumberPoint>(), null, "General", hiddenColumnSource, parsedWorkbook]) ?? throw new InvalidOperationException("Expected no-cache hidden-column indexed chart vector.");
        object[] noCacheHiddenDensePoints = (((System.Collections.IEnumerable?)noCacheHiddenColumnVector.GetType().GetMethod("DensePoints")?.Invoke(noCacheHiddenColumnVector, [])) ?? throw new InvalidOperationException("Expected visibility-filtered no-cache dense points.")).Cast<object>().ToArray();
        TestAssert.True(noCacheHiddenDensePoints.Length == 0, "Expected no-cache dense projection to honor plot-visible-only when the workbook source column is hidden.");

        var labelSource = new PptxSceneChartDataSource(
            "Sheet1!$A$2:$A$4",
            PptxSceneChartDataSourceReferenceKind.StringReference,
            "strRef",
            PptxSceneChartDataSourceCacheKind.StringCache,
            "strCache",
            HasCachedPoints: true);
        System.Reflection.MethodInfo buildTextVector = typeof(PptxRenderer)
            .GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(method => method.Name == "BuildChartIndexedTextVector" && method.GetParameters().Length == 6);
        object textVector = buildTextVector.Invoke(
            null,
            [
                new[] { "Cached" },
                Array.Empty<PptxSceneChartStringPoint>(),
                3,
                new List<IReadOnlyList<PptxSceneChartStringPoint>>(),
                labelSource,
                parsedWorkbook
            ]) ?? throw new InvalidOperationException("Expected indexed chart text vector.");
        object[] visibleLabelWorkbookPoints = (((System.Collections.IEnumerable?)textVector.GetType().GetMethod("WorkbookPointsForPlotVisibility")?.Invoke(textVector, [true])) ?? throw new InvalidOperationException("Expected visibility-filtered text workbook points.")).Cast<object>().ToArray();
        TestAssert.True(visibleLabelWorkbookPoints.Length == 2, "Expected plot-visible-only projection to exclude the hidden worksheet row while preserving visible label points.");
        object visibleLabelCell = visibleLabelWorkbookPoints[1].GetType().GetProperty("WorkbookCell")?.GetValue(visibleLabelWorkbookPoints[1]) ?? throw new InvalidOperationException("Expected visible text point cell provenance.");
        TestAssert.Equal("A4", (string?)visibleLabelCell.GetType().GetProperty("Reference")?.GetValue(visibleLabelCell) ?? string.Empty);
        var noCacheLabelSource = labelSource with { HasCachedPoints = false };
        object noCacheTextVector = buildTextVector.Invoke(
            null,
            [
                Array.Empty<string>(),
                Array.Empty<PptxSceneChartStringPoint>(),
                null,
                new List<IReadOnlyList<PptxSceneChartStringPoint>>(),
                noCacheLabelSource,
                parsedWorkbook
            ]) ?? throw new InvalidOperationException("Expected formula-only indexed chart text vector.");
        object[] noCacheTextActivePoints = (((System.Collections.IEnumerable?)noCacheTextVector.GetType().GetProperty("Points")?.GetValue(noCacheTextVector)) ?? throw new InvalidOperationException("Expected formula-only active text points.")).Cast<object>().ToArray();
        object[] noCacheTextWorkbookPoints = (((System.Collections.IEnumerable?)noCacheTextVector.GetType().GetProperty("WorkbookPoints")?.GetValue(noCacheTextVector)) ?? throw new InvalidOperationException("Expected formula-only text workbook sidecar points.")).Cast<object>().ToArray();
        TestAssert.True(noCacheTextActivePoints.Length == 0, "Expected formula-only workbook text values to remain inactive when chart caches are absent.");
        TestAssert.True(noCacheTextWorkbookPoints.Length == 3, "Expected formula-only workbook text values to remain available as sidecar points.");
    }

    public static void PptxChartSeriesNamesPreserveWorkbookSidecarPoints()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:lineChart>
                <c:ser>
                  <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:pt idx="0"><c:v>Cached Name</c:v></c:pt></c:strCache></c:strRef></c:tx>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;
        PptxSceneChartPlot plot = PptxTests.BuildSingleChartScene(chartXml)?.Plots.Single()
            ?? throw new InvalidOperationException("Expected one chart plot.");
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement chartElement = XDocument.Parse(chartXml).Descendants(chartNamespace + "lineChart").Single();
        Type workbookType = typeof(PptxRenderer).GetNestedType(
            "ChartWorkbookData",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart workbook data type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["B1"] = "Workbook Name"
            }
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets]) ?? throw new InvalidOperationException("Expected workbook instance.");
        System.Reflection.MethodInfo readNameRecords = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartSeriesNameRecords",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series-name record reader.");

        object[] records = (((System.Collections.IEnumerable?)readNameRecords.Invoke(null, [plot, chartElement, workbook])) ?? throw new InvalidOperationException("Expected series-name records.")).Cast<object>().ToArray();

        TestAssert.True(records.Length == 1, "Expected one series-name record.");
        TestAssert.Equal("Cached Name", (string?)records[0].GetType().GetProperty("ActiveName")?.GetValue(records[0]) ?? string.Empty);
        TestAssert.Equal("Cache", records[0].GetType().GetProperty("ActiveNameSource")?.GetValue(records[0])?.ToString() ?? string.Empty);
        object[] workbookPoints = (((System.Collections.IEnumerable?)records[0].GetType().GetProperty("WorkbookPoints")?.GetValue(records[0])) ?? throw new InvalidOperationException("Expected workbook series-name sidecar points.")).Cast<object>().ToArray();
        TestAssert.True(workbookPoints.Length == 1, "Expected series-name sidecar to preserve the workbook source name.");
        TestAssert.Equal("Workbook Name", (string?)workbookPoints[0].GetType().GetProperty("Text")?.GetValue(workbookPoints[0]) ?? string.Empty);
        object workbookCell = workbookPoints[0].GetType().GetProperty("WorkbookCell")?.GetValue(workbookPoints[0]) ?? throw new InvalidOperationException("Expected series-name workbook cell provenance.");
        TestAssert.Equal("B1", (string?)workbookCell.GetType().GetProperty("Reference")?.GetValue(workbookCell) ?? string.Empty);
        Array typedRecords = Array.CreateInstance(records[0].GetType(), records.Length);
        typedRecords.SetValue(records[0], 0);
        System.Reflection.MethodInfo getActiveSeriesName = typeof(PptxRenderer).GetMethod(
            "GetActiveSeriesName",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected active series-name helper.");
        string activeSeriesName = (string?)getActiveSeriesName.Invoke(null, [typedRecords, 0]) ?? string.Empty;
        TestAssert.Equal("Cached Name", activeSeriesName);

        object[] rawRecords = (((System.Collections.IEnumerable?)readNameRecords.Invoke(null, [null, chartElement, workbook])) ?? throw new InvalidOperationException("Expected raw fallback series-name records.")).Cast<object>().ToArray();
        TestAssert.True(rawRecords.Length == 1, "Expected raw XML fallback to preserve one series-name record.");
        TestAssert.Equal("Cached Name", (string?)rawRecords[0].GetType().GetProperty("ActiveName")?.GetValue(rawRecords[0]) ?? string.Empty);
        TestAssert.Equal("Cache", rawRecords[0].GetType().GetProperty("ActiveNameSource")?.GetValue(rawRecords[0])?.ToString() ?? string.Empty);
        object rawSource = rawRecords[0].GetType().GetProperty("Source")?.GetValue(rawRecords[0]) ?? throw new InvalidOperationException("Expected raw series-name source metadata.");
        TestAssert.Equal("Sheet1!$B$1", (string?)rawSource.GetType().GetProperty("Formula")?.GetValue(rawSource) ?? string.Empty);
        object[] rawWorkbookPoints = (((System.Collections.IEnumerable?)rawRecords[0].GetType().GetProperty("WorkbookPoints")?.GetValue(rawRecords[0])) ?? throw new InvalidOperationException("Expected raw series-name workbook sidecar points.")).Cast<object>().ToArray();
        TestAssert.True(rawWorkbookPoints.Length == 1, "Expected raw XML series-name fallback to preserve workbook sidecar points.");
        TestAssert.Equal("Workbook Name", (string?)rawWorkbookPoints[0].GetType().GetProperty("Text")?.GetValue(rawWorkbookPoints[0]) ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.StringReference, (PptxSceneChartDataSourceReferenceKind?)rawSource.GetType().GetProperty("ReferenceKindValue")?.GetValue(rawSource) ?? default);
        TestAssert.Equal("strRef", (string?)rawSource.GetType().GetProperty("ReferenceKind")?.GetValue(rawSource) ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceCacheKind.StringCache, (PptxSceneChartDataSourceCacheKind?)rawSource.GetType().GetProperty("CacheKindValue")?.GetValue(rawSource) ?? default);
        TestAssert.Equal("strCache", (string?)rawSource.GetType().GetProperty("CacheKind")?.GetValue(rawSource) ?? string.Empty);
        TestAssert.True((bool?)rawSource.GetType().GetProperty("HasCachedPoints")?.GetValue(rawSource) == true, "Expected raw series-name cache point presence to survive fallback parsing.");

        const string noCacheChartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:lineChart>
                <c:ser>
                  <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;
        PptxSceneChartPlot noCachePlot = PptxTests.BuildSingleChartScene(noCacheChartXml)?.Plots.Single()
            ?? throw new InvalidOperationException("Expected one no-cache chart plot.");
        XElement noCacheChartElement = XDocument.Parse(noCacheChartXml).Descendants(chartNamespace + "lineChart").Single();
        object[] noCacheSceneRecords = (((System.Collections.IEnumerable?)readNameRecords.Invoke(null, [noCachePlot, noCacheChartElement, workbook])) ?? throw new InvalidOperationException("Expected no-cache scene series-name records.")).Cast<object>().ToArray();
        TestAssert.Equal("Workbook Name", (string?)noCacheSceneRecords[0].GetType().GetProperty("ActiveName")?.GetValue(noCacheSceneRecords[0]) ?? string.Empty);
        TestAssert.Equal("Workbook", noCacheSceneRecords[0].GetType().GetProperty("ActiveNameSource")?.GetValue(noCacheSceneRecords[0])?.ToString() ?? string.Empty);

        object[] noCacheRawRecords = (((System.Collections.IEnumerable?)readNameRecords.Invoke(null, [null, noCacheChartElement, workbook])) ?? throw new InvalidOperationException("Expected no-cache raw series-name records.")).Cast<object>().ToArray();
        TestAssert.Equal("Series 1", (string?)noCacheRawRecords[0].GetType().GetProperty("ActiveName")?.GetValue(noCacheRawRecords[0]) ?? string.Empty);
        TestAssert.Equal("Default", noCacheRawRecords[0].GetType().GetProperty("ActiveNameSource")?.GetValue(noCacheRawRecords[0])?.ToString() ?? string.Empty);
    }

    public static void PptxChartCategoryAxisSourceAcceptsDateAxes()
    {
        const string chartXml = """
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser>
                    <c:idx val="0"/><c:order val="0"/>
                    <c:cat><c:numRef><c:f>Sheet1!$A$2:$A$4</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>44927</c:v></c:pt><c:pt idx="1"><c:v>44928</c:v></c:pt><c:pt idx="2"><c:v>44929</c:v></c:pt></c:numCache></c:numRef></c:cat>
                    <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:dateAx>
                  <c:axId val="10"/><c:axPos val="b"/>
                  <c:scaling><c:orientation val="minMax"/></c:scaling>
                  <c:numFmt formatCode="m/d/yy" sourceLinked="1"/>
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100"><a:solidFill><a:srgbClr val="1155AA"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                  <c:crossAx val="20"/><c:tickLblPos val="low"/>
                </c:dateAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """;
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        PptxSceneChart chart = PptxTests.BuildSingleChartScene(chartXml) ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartPlot plot = chart.Plots.Single();
        XElement lineChart = chart.ChartXml?.Descendants(c + "lineChart").Single() ?? throw new InvalidOperationException("Expected line chart XML.");
        System.Reflection.MethodInfo readCategoryAxis = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartCategoryAxisForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected category-axis source helper.");

        object source = readCategoryAxis.Invoke(null, [chart, plot, chart.ChartXml!, lineChart]) ?? throw new InvalidOperationException("Expected category-axis source.");
        PptxSceneChartAxis sceneAxis = (PptxSceneChartAxis?)source.GetType().GetProperty("SceneAxis")?.GetValue(source) ?? throw new InvalidOperationException("Expected scene date axis source.");
        XElement xmlAxis = (XElement?)source.GetType().GetProperty("XmlAxis")?.GetValue(source) ?? throw new InvalidOperationException("Expected XML date axis source.");
        object xmlOnlySource = readCategoryAxis.Invoke(null, [null, null, chart.ChartXml!, lineChart]) ?? throw new InvalidOperationException("Expected XML-only category-axis source.");
        XElement xmlOnlyAxis = (XElement?)xmlOnlySource.GetType().GetProperty("XmlAxis")?.GetValue(xmlOnlySource) ?? throw new InvalidOperationException("Expected XML-only date axis source.");

        TestAssert.Equal(PptxSceneChartAxisKind.Date, sceneAxis.AxisKind);
        TestAssert.Equal("dateAx", sceneAxis.Kind);
        TestAssert.Equal("10", sceneAxis.Id);
        TestAssert.Equal("dateAx", xmlAxis.Name.LocalName);
        TestAssert.Equal("10", xmlAxis.Element(c + "axId")?.Attribute("val")?.Value ?? string.Empty);
        TestAssert.Equal("dateAx", xmlOnlyAxis.Name.LocalName);
    }

    public static void PptxSyntheticScatterChartRendersLegendKeyOnlyDataLabels()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:scatterChart>
                  <c:scatterStyle val="marker"/>
                  <c:dLbls><c:showVal val="0"/><c:showLegendKey val="1"/><c:dLblPos val="t"/><c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></c:spPr></c:dLbls>
                  <c:ser>
                    <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
                    <c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:xVal>
                    <c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:yVal>
                  </c:ser>
                </c:scatterChart></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.071 0.204 0.337 rg", pdf);
        TestAssert.Contains(" re W* n", pdf);
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Scatter charts with legend-key-only labels should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Scatter charts with legend-key-only labels should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticScatterChartConsumesChartWideDataLabelManualBox()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:scatterChart>
                  <c:scatterStyle val="marker"/>
                  <c:dLbls>
                    <c:layout><c:manualLayout><c:x val="0.10"/><c:y val="0.15"/><c:w val="0.55"/><c:h val="0.35"/></c:manualLayout></c:layout>
                    <c:showVal val="1"/>
                    <c:dLblPos val="t"/>
                    <c:spPr><a:solidFill><a:srgbClr val="19334D"/></a:solidFill></c:spPr>
                  </c:dLbls>
                  <c:ser>
                    <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
                    <c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal>
                    <c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal>
                  </c:ser>
                </c:scatterChart></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/CSD1", pdf);
        TestAssert.True(
            Regex.IsMatch(pdf, @"0\.098 0\.2 0\.302 rg\s+[0-9.]+ [0-9.]+ (?:1[5-9][0-9]|[2-9][0-9]{2})\.[0-9]{2,3} (?:[5-9][0-9]|[1-9][0-9]{2})\.[0-9]{2,3} re f"),
            "Expected chart-wide data-label manual layout to drive a large label shape box.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Scatter charts with chart-wide manual data-label layout should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Scatter charts with chart-wide manual data-label layout should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticBarChartRendersLegendKeyOnlyDataLabels()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart>
                  <c:barDir val="col"/>
                  <c:grouping val="clustered"/>
                  <c:varyColors val="1"/>
                  <c:dLbls><c:showVal val="0"/><c:showLegendKey val="1"/><c:dLblPos val="outEnd"/><c:spPr><a:solidFill><a:srgbClr val="654321"/></a:solidFill></c:spPr></c:dLbls>
                  <c:ser>
                    <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.396 0.263 0.129 rg", pdf);
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Bar charts with legend-key-only labels should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Bar charts with legend-key-only labels should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticBarChartLegendSwatchesUseSeriesStroke()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart>
                  <c:barDir val="col"/>
                  <c:grouping val="clustered"/>
                  <c:varyColors val="0"/>
                  <c:ser>
                    <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:ser>
                    <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="654321"/></a:solidFill></a:ln></c:spPr>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.267 0.447 0.769 rg", pdf);
        TestAssert.Contains("0.071 0.204 0.337 RG", pdf);
        TestAssert.Contains("0.929 0.49 0.192 rg", pdf);
        TestAssert.Contains("0.396 0.263 0.129 RG", pdf);
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Bar chart legends with stroked swatches should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Bar chart legends with stroked swatches should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticLineChartRendersLegendKeyOnlyDataLabels()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
                  <c:grouping val="standard"/>
                  <c:dLbls><c:showVal val="0"/><c:showLegendKey val="1"/><c:dLblPos val="t"/><c:spPr><a:solidFill><a:srgbClr val="2468AC"/></a:solidFill></c:spPr></c:dLbls>
                  <c:ser>
                    <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                    <c:spPr><a:ln w="38100"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></c:spPr>
                    <c:marker><c:symbol val="circle"/><c:size val="9"/><c:spPr><a:solidFill><a:srgbClr val="00B050"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></a:ln></c:spPr></c:marker>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.141 0.408 0.675 rg", pdf);
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Line charts with legend-key-only labels should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Line charts with legend-key-only labels should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticLineChartSeriesLineWithoutWidthUsesOfficeDefault()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
                  <c:grouping val="standard"/>
                  <c:marker val="0"/>
                  <c:ser>
                    <c:spPr><a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="4"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"0\.071 0\.204 0\.337 RG\s+3 w"), "Expected chart series <a:ln> without @w to inherit Office's 3 pt series line width.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Line chart series stroke defaults should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Line chart series stroke defaults should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticBarChartSeriesLineWithoutWidthUsesFilledSeriesDefault()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotArea><c:barChart>
                  <c:barDir val="col"/><c:grouping val="clustered"/>
                  <c:ser>
                    <c:tx><c:v>Actual</c:v></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="4"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"0\.071 0\.204 0\.337 RG\s+0\.75 w"), "Expected filled chart series <a:ln> without @w to inherit Office's 0.75 pt filled-series border width.");
        TestAssert.True(!Regex.IsMatch(pdf, @"0\.071 0\.204 0\.337 RG\s+3 w"), "Filled chart series <a:ln> without @w should not use the 3 pt line-chart inherited width.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Bar chart series stroke defaults should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Bar chart series stroke defaults should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticLineChartMarkerLineWithoutWidthUsesMarkerDefault()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotArea><c:lineChart>
                  <c:grouping val="standard"/><c:marker val="1"/>
                  <c:ser>
                    <c:tx><c:v>Actual</c:v></c:tx>
                    <c:spPr><a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr>
                    <c:marker><c:symbol val="square"/><c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:ln></c:spPr></c:marker>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="4"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"0\.671 0\.804 0\.937 RG\s+0\.75 w"), "Expected line-chart marker <a:ln> without @w to inherit Office's 0.75 pt marker-outline width.");
        TestAssert.True(!Regex.IsMatch(pdf, @"0\.671 0\.804 0\.937 RG\s+3 w"), "Marker <a:ln> without @w should not use the 3 pt series-line inherited width.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Line chart marker stroke defaults should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Line chart marker stroke defaults should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticLineChartConsumesDisplayBlanksAs()
    {
        string gapPdf = PptxTests.RenderLineChartWithDisplayBlanksAs("gap");
        string spanPdf = PptxTests.RenderLineChartWithDisplayBlanksAs("span");
        string zeroPdf = PptxTests.RenderLineChartWithDisplayBlanksAs("zero");

        TestAssert.Equal(0, PptxTests.CountLineChartSeriesSegments(gapPdf, "0.071 0.671 0.204 RG"));
        TestAssert.Equal(1, PptxTests.CountLineChartSeriesSegments(spanPdf, "0.071 0.671 0.204 RG"));
        TestAssert.Equal(2, PptxTests.CountLineChartSeriesSegments(zeroPdf, "0.071 0.671 0.204 RG"));
    }

    public static void PptxSyntheticLineChartPreservesSeriesStylesAfterBlankSeries()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
                  <c:grouping val="standard"/>
                  <c:marker val="0"/>
                  <c:ser>
                    <c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></c:spPr>
                    <c:marker><c:symbol val="none"/></c:marker>
                    <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v></c:v></c:pt><c:pt idx="1"><c:v></c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:ser>
                    <c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="12AB34"/></a:solidFill></a:ln></c:spPr>
                    <c:marker><c:symbol val="none"/></c:marker>
                    <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="20"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:legend><c:delete val="1"/></c:legend></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(1, PptxTests.CountLineChartSeriesSegments(pdf, "0.071 0.671 0.204 RG"));
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Line charts with leading blank series should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Line charts with leading blank series should not emit unsupported chart diagnostics.");
    }

    public static void PptxChartExplosionsUseCachedPointCount()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:doughnutChart>
                <c:ser><c:explosion val="25"/>
                  <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="2"><c:v>West</c:v></c:pt></c:strCache></c:strRef></c:cat>
                  <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>8.2</c:v></c:pt><c:pt idx="2"><c:v>1.4</c:v></c:pt></c:numCache></c:numRef></c:val>
                </c:ser>
              </c:doughnutChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;
        PptxSceneChartPlot plot = PptxTests.BuildSingleChartScene(chartXml)?.Plots.Single()
            ?? throw new InvalidOperationException("Expected one chart plot.");
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement chartElement = XDocument.Parse(chartXml).Descendants(chartNamespace + "doughnutChart").Single();
        Type workbookType = typeof(PptxRenderer).GetNestedType(
            "ChartWorkbookData",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart workbook data type.");
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["A2"] = "North",
                ["A4"] = "West",
                ["B2"] = "8.2",
                ["B4"] = "1.4"
            }
        };
        object workbook = Activator.CreateInstance(workbookType, [sheets]) ?? throw new InvalidOperationException("Expected workbook instance.");
        System.Reflection.MethodInfo readExplosions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartPointExplosions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected point-explosion helper.");

        var explosions = (IReadOnlyDictionary<int, double>)(readExplosions.Invoke(null, [plot, chartElement, workbook])
            ?? throw new InvalidOperationException("Expected point explosions."));

        TestAssert.Equal(3, explosions.Count);
        TestAssert.Equal(0.25d, explosions[0]);
        TestAssert.Equal(0.25d, explosions[1]);
        TestAssert.Equal(0.25d, explosions[2]);
    }

    public static void PptxChartPointExplosionReaderRejectsNegativeIndices()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:doughnutChart>
                <c:ser>
                  <c:dPt><c:idx val="-1"/><c:explosion val="80"/></c:dPt>
                  <c:dPt><c:idx val="0"/><c:explosion val="25"/></c:dPt>
                </c:ser>
              </c:doughnutChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement chartElement = XDocument.Parse(chartXml).Descendants(chartNamespace + "doughnutChart").Single();
        System.Reflection.MethodInfo readExplosions = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartPointExplosions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected point-explosion helper.");

        var explosions = (IReadOnlyDictionary<int, double>)(readExplosions.Invoke(null, [null, chartElement, null])
            ?? throw new InvalidOperationException("Expected point explosions."));

        TestAssert.Equal(1, explosions.Count);
        TestAssert.True(!explosions.ContainsKey(-1), "Expected negative keyed point explosions to be rejected.");
        TestAssert.Equal(0.25d, explosions[0]);
    }

    public static void PptxScenePreservesRejectedChartKeyedOverrideIndices()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea>
                <c:pieChart>
                  <c:ser>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                    <c:marker><c:symbol val="circle"/><c:size val="7"/></c:marker>
                    <c:dPt><c:idx val="0"/><c:explosion val="25"/><c:spPr><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill></c:spPr></c:dPt>
                    <c:dPt><c:idx val="-1"/><c:spPr><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></c:spPr></c:dPt>
                    <c:dPt><c:idx val="futurePoint"/><c:spPr><a:solidFill><a:srgbClr val="0000AA"/></a:solidFill></c:spPr></c:dPt>
                    <c:dPt><c:spPr><a:solidFill><a:srgbClr val="999999"/></a:solidFill></c:spPr></c:dPt>
                  </c:ser>
                  <c:dLbls>
                    <c:dLbl><c:idx val="0"/><c:layout><c:manualLayout><c:x val="0.1"/><c:y val="0.2"/></c:manualLayout></c:layout><c:showVal val="1"/></c:dLbl>
                    <c:dLbl><c:idx val="-1"/><c:showVal val="1"/></c:dLbl>
                    <c:dLbl><c:idx val="futureLabel"/><c:showVal val="1"/></c:dLbl>
                    <c:dLbl><c:showVal val="1"/></c:dLbl>
                  </c:dLbls>
                </c:pieChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """;
        PptxSceneChart chart = PptxTests.BuildSingleChartScene(chartXml) ?? throw new InvalidOperationException("Expected chart scene.");

        PptxSceneChartSeries series = chart.Plots[0].Series[0];
        TestAssert.True(series.Marker.IsDefined, "Expected explicit chart marker to remain observable.");
        TestAssert.Equal(1, series.PointStyles.Count);
        TestAssert.Equal(25d, series.PointStyles[0].Explosion ?? -1d);
        TestAssert.Equal(3, series.RejectedPointStyleIndexValues.Count);
        TestAssert.Equal("-1", series.RejectedPointStyleIndexValues[0]);
        TestAssert.Equal("futurePoint", series.RejectedPointStyleIndexValues[1]);
        TestAssert.Equal(string.Empty, series.RejectedPointStyleIndexValues[2]);

        TestAssert.Equal(1, chart.Plots[0].DataLabels.Overrides.Count);
        TestAssert.True(chart.Plots[0].DataLabels.Overrides[0].Layout.HasLayout, "Expected data-label manual layout to remain observable.");
        TestAssert.Equal(3, chart.Plots[0].DataLabels.RejectedOverrideIndexValues.Count);
        TestAssert.Equal("-1", chart.Plots[0].DataLabels.RejectedOverrideIndexValues[0]);
        TestAssert.Equal("futureLabel", chart.Plots[0].DataLabels.RejectedOverrideIndexValues[1]);
        TestAssert.Equal(string.Empty, chart.Plots[0].DataLabels.RejectedOverrideIndexValues[2]);

        PptxSceneNodeSnapshot snapshot = PptxTests.BuildSingleChartSceneSnapshot(chartXml);
        TestAssert.Equal(1, snapshot.ChartSeriesCount);
        TestAssert.Equal(1, snapshot.ChartSeriesMarkerCount);
        TestAssert.Equal(1, snapshot.ChartSeriesPointStyleCount);
        TestAssert.Equal(1, snapshot.ChartSeriesPointExplosionCount);
        TestAssert.Equal(1, snapshot.ChartDataLabelsDefinedCount);
        TestAssert.Equal(1, snapshot.ChartDataLabelOverrideCount);
        TestAssert.Equal(1, snapshot.ChartDataLabelManualLayoutCount);
        TestAssert.Equal(3, snapshot.ChartRejectedPointStyleIndexCount);
        TestAssert.Equal("-1", snapshot.ChartRejectedPointStyleIndexValues[0]);
        TestAssert.Equal("futurePoint", snapshot.ChartRejectedPointStyleIndexValues[1]);
        TestAssert.Equal(string.Empty, snapshot.ChartRejectedPointStyleIndexValues[2]);
        TestAssert.Equal(3, snapshot.ChartRejectedDataLabelOverrideIndexCount);
        TestAssert.Equal("-1", snapshot.ChartRejectedDataLabelOverrideIndexValues[0]);
        TestAssert.Equal("futureLabel", snapshot.ChartRejectedDataLabelOverrideIndexValues[1]);
        TestAssert.Equal(string.Empty, snapshot.ChartRejectedDataLabelOverrideIndexValues[2]);
    }

    public static void PptxScenePreservesChartColorStyleDeclarations()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:barChart>
                <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;
        (PptxDocument document, OoxPackage package) = PptxTests.BuildSingleChartPackage(chartXml, new Dictionary<string, byte[]>
        {
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
                </Relationships>
                """),
            ["ppt/charts/colors1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                               meth="cycle" id="77">
                  <a:srgbClr val="FF00CC"/>
                  <a:schemeClr val="accent1"/>
                  <cs:variation>
                    <a:srgbClr val="112233"/>
                  </cs:variation>
                </cs:colorStyle>
                """)
        });
        PptxSceneChart chart = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0].SlideNodes[0].Chart
            ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneNodeSnapshot snapshot = PptxRenderer.InspectScene(document, package).Slides[0].SlideNodes[0];

        TestAssert.Equal(1, chart.ColorStyle.VariationCount);
        TestAssert.Equal(3, chart.ColorStyle.Declarations.Count);
        TestAssert.Equal(2, chart.ColorStyle.RootDeclarations.Count);
        TestAssert.Equal("srgbClr", chart.ColorStyle.Declarations[0].Kind);
        TestAssert.Equal("FF00CC", chart.ColorStyle.Declarations[0].Value);
        TestAssert.Equal(null, chart.ColorStyle.Declarations[0].VariationIndex);
        TestAssert.Equal(new RgbColor(255, 0, 204), chart.ColorStyle.Declarations[0].Color ?? default);
        TestAssert.Equal("schemeClr", chart.ColorStyle.Declarations[1].Kind);
        TestAssert.Equal("accent1", chart.ColorStyle.Declarations[1].Value);
        TestAssert.Equal(null, chart.ColorStyle.Declarations[1].VariationIndex);
        TestAssert.Equal("srgbClr", chart.ColorStyle.RootDeclarations[0].Kind);
        TestAssert.Equal("schemeClr", chart.ColorStyle.RootDeclarations[1].Kind);
        TestAssert.Equal("srgbClr", chart.ColorStyle.Declarations[2].Kind);
        TestAssert.Equal("112233", chart.ColorStyle.Declarations[2].Value);
        TestAssert.Equal(0, chart.ColorStyle.Declarations[2].VariationIndex ?? -1);
        TestAssert.Equal(1, chart.ColorStyle.Variations.Count);
        TestAssert.Equal(0, chart.ColorStyle.Variations[0].Index);
        TestAssert.Equal(1, chart.ColorStyle.Variations[0].Declarations.Count);
        TestAssert.Equal("112233", chart.ColorStyle.Variations[0].Declarations[0].Value);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.ColorStyle.Variations[0].Colors[0]);
        TestAssert.True(chart.ColorStyle.Declarations[0].IsResolved, "Expected direct RGB color-style declaration to resolve.");
        TestAssert.True(!chart.ColorStyle.Declarations[1].IsResolved, "Expected unresolved scheme color-style declaration to remain observable.");
        TestAssert.Equal(1, chart.ColorStyle.Colors.Count);
        TestAssert.Equal(1, snapshot.ChartColorStyleVariationCount);
        TestAssert.Equal(3, snapshot.ChartColorStyleDeclarationCount);
        TestAssert.Equal(2, snapshot.ChartColorStyleRootDeclarationCount);
        TestAssert.Equal(2, snapshot.ChartColorStyleResolvedDeclarationCount);
        TestAssert.Equal("srgbClr", snapshot.ChartColorStyleDeclarationKinds[0]);
        TestAssert.Equal("schemeClr", snapshot.ChartColorStyleDeclarationKinds[1]);
        TestAssert.Equal("srgbClr", snapshot.ChartColorStyleDeclarationKinds[2]);
    }

    public static void PptxScenePreservesChartStyleUnderlineStrikeOnlyTextRoles()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:lineChart>
                <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;
        (PptxDocument document, OoxPackage package) = PptxTests.BuildSingleChartPackage(chartXml, new Dictionary<string, byte[]>
        {
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
                </Relationships>
                """),
            ["ppt/charts/style1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="91">
                  <cs:legend>
                    <cs:defRPr u="sng" strike="sngStrike"/>
                  </cs:legend>
                </cs:style>
                """)
        });

        PptxSceneChart chart = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0].SlideNodes[0].Chart
            ?? throw new InvalidOperationException("Expected chart scene.");
        PptxSceneChartStyleEntry legendStyle = chart.StylePart.Entries.FirstOrDefault(entry => entry.Role == "legend");

        TestAssert.True(chart.StylePart.IsDefined, "Expected chart style-part ownership in the scene model.");
        TestAssert.Equal("legend", legendStyle.Role ?? string.Empty);
        TestAssert.True(legendStyle.TextStyle.Underline == true, "Expected underline-only chart-style text role to remain scene-visible.");
        TestAssert.True(legendStyle.TextStyle.Strike == true, "Expected strike-only chart-style text role to remain scene-visible.");
        TestAssert.Equal(null, legendStyle.TextStyle.FontFamily);
        TestAssert.Equal(null, legendStyle.TextStyle.Color);
    }

    public static void PptxPercentStackedColumnChartUsesPercentValueAxis()
    {
        string input = Path.Combine(
            Directory.GetCurrentDirectory(),
            "tests",
            "Lokad.OoxPdf.Tests",
            "Cases",
            "pptx-ladder-11-chart-column-100-stacked-port.pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = PptxTests.ReadPdfDecodedAscii(output);
        TestAssert.True(
            Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0025>"),
            "Expected percent-stacked value-axis labels to include a percent sign by default.");
        TestAssert.True(
            !Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <002E>"),
            "Percent-stacked value-axis labels should not use the generic 0..1.2 decimal scale.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Percent-stacked column charts should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Percent-stacked column charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxUnsupportedEffectDiagnosticsUseSceneChartEffects()
    {
        string contentTypes = PptxTests.BasicContentTypes().Replace(
            "</Types>",
            """
              <Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
            </Types>
            """);
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = contentTypes,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:nvPr/></p:nvGraphicFramePr>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/charts/chart1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:spPr><a:effectDag/></c:spPr>
                  <c:chart><c:plotArea><c:lineChart>
                    <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                  </c:lineChart></c:plotArea></c:chart>
                </c:chartSpace>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_EFFECT"), "Unsupported chart effects should be diagnostic-covered from scene-owned chart effect provenance.");
    }

    public static void PptxUnsupportedEffectDiagnosticsUseSceneChartStyleEffectReferences()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:lineChart>
                <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;
        (PptxDocument document, OoxPackage package) = PptxTests.BuildSingleChartPackage(chartXml, new Dictionary<string, byte[]>
        {
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """),
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
                </Relationships>
                """),
            ["ppt/charts/style1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="1">
                  <cs:plotArea><cs:effectRef idx="1"/></cs:plotArea>
                </cs:style>
                """),
            ["ppt/theme/theme1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">
                  <a:themeElements>
                    <a:clrScheme name="Test"/>
                    <a:fontScheme name="Test"/>
                    <a:fmtScheme name="Test">
                      <a:fillStyleLst/>
                      <a:lnStyleLst/>
                      <a:effectStyleLst>
                        <a:effectStyle><a:effectLst><a:blur rad="6350"/></a:effectLst></a:effectStyle>
                      </a:effectStyleLst>
                      <a:bgFillStyleLst/>
                    </a:fmtScheme>
                  </a:themeElements>
                </a:theme>
                """)
        });
        PptxSceneSlide sceneSlide = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0];
        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutEffects = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutEffects, "/ppt/slides/slide1.xml", 1, sink]);

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_EFFECT"), "Unsupported chart style effect references should be diagnostic-covered from scene-owned effect-reference provenance.");
    }

    public static void PptxUnsupportedGradientDiagnosticsUseSceneChartShapeStyleGradient()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:spPr>
                <a:gradFill>
                  <a:gsLst>
                    <a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="35000"/></a:srgbClr></a:gs>
                    <a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="85000"/></a:srgbClr></a:gs>
                  </a:gsLst>
                  <a:lin ang="0"/>
                </a:gradFill>
              </c:spPr>
              <c:chart><c:plotArea>
                <c:spPr>
                  <a:gradFill>
                    <a:gsLst>
                      <a:gs pos="0"><a:srgbClr val="00FF00"><a:alpha val="100000"/></a:srgbClr></a:gs>
                      <a:gs pos="100000"><a:srgbClr val="000000"><a:alpha val="25000"/></a:srgbClr></a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000"/>
                  </a:gradFill>
                </c:spPr>
                <c:lineChart>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """;
        (PptxDocument document, OoxPackage package) = PptxTests.BuildSingleChartPackage(chartXml);
        PptxSceneSlide sceneSlide = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0];
        PptxSceneChart chart = sceneSlide.SlideNodes[0].Chart
            ?? throw new InvalidOperationException("Expected chart scene node.");

        TestAssert.True(chart.ChartAreaStyle.GradientFill?.HasUnsupportedGradient == true, "Expected chart-area variable-alpha gradient state to remain scene-owned.");
        TestAssert.True(chart.PlotAreaStyle.GradientFill?.HasUnsupportedGradient == true, "Expected plot-area variable-alpha gradient state to remain scene-owned.");

        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutGradients = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutGradients, "/ppt/slides/slide1.xml", 1, sink]);

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Unsupported chart shape-style gradients should be diagnostic-covered from scene-owned gradient provenance.");
    }

    public static void PptxUnsupportedPictureFillDiagnosticsUseSceneChartShapeStylePictureFill()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <c:spPr>
                <a:blipFill>
                  <a:blip r:embed="rIdImage"><a:alphaModFix amt="42000"/></a:blip>
                  <a:srcRect l="10000" t="20000"/>
                  <a:stretch><a:fillRect r="30000" b="40000"/></a:stretch>
                </a:blipFill>
              </c:spPr>
              <c:chart><c:plotArea>
                <c:spPr>
                  <a:blipFill>
                    <a:blip><a:alphaModFix amt="62500"/></a:blip>
                    <a:tile algn="ctr" flip="x"/>
                  </a:blipFill>
                </c:spPr>
                <c:lineChart>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """;
        (PptxDocument document, OoxPackage package) = PptxTests.BuildSingleChartPackage(chartXml);
        PptxSceneSlide sceneSlide = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0];
        PptxSceneChart chart = sceneSlide.SlideNodes[0].Chart
            ?? throw new InvalidOperationException("Expected chart scene node.");

        TestAssert.True(chart.ChartAreaStyle.PictureFill.HasPicture, "Expected chart-area picture-fill state to remain scene-owned.");
        TestAssert.True(chart.PlotAreaStyle.PictureFill.HasPicture, "Expected plot-area picture-fill state to remain scene-owned.");

        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutPictureFills = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutPictureFills, "/ppt/slides/slide1.xml", 1, sink]);

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_PICTURE_FILL"), "Unsupported chart shape-style picture fills should be diagnostic-covered from scene-owned picture-fill provenance.");
    }

    public static void PptxUnsupportedTextDiagnosticsUseSceneChartTextBodyProperties()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>
                <c:title>
                  <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Title</a:t></a:r></a:p></c:rich></c:tx>
                  <c:txPr><a:bodyPr vert="futureVert" vertOverflow="ellipsis"/></c:txPr>
                </c:title>
                <c:plotArea><c:lineChart>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart></c:plotArea>
              </c:chart>
            </c:chartSpace>
            """;
        (PptxDocument document, OoxPackage package) = PptxTests.BuildSingleChartPackage(chartXml);
        PptxSceneSlide sceneSlide = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0];
        TestAssert.Equal("futureVert", sceneSlide.SlideNodes[0].Chart?.Title.TextBodyProperties.OrientationValue ?? string.Empty);
        TestAssert.Equal("ellipsis", sceneSlide.SlideNodes[0].Chart?.Title.TextBodyProperties.VerticalOverflowValue ?? string.Empty);
        PptxSceneNodeSnapshot snapshot = PptxRenderer.InspectScene(document, package).Slides[0].SlideNodes[0];
        TestAssert.Equal(1, snapshot.ChartTextBodyOrientationCount);
        TestAssert.Equal(1, snapshot.ChartTextBodyVerticalOverflowCount);

        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutBodyProperties = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutBodyProperties, "/ppt/slides/slide1.xml", 1, sink]);

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_TEXT_ORIENTATION"), "Unsupported chart text orientation should be diagnostic-covered from scene-owned text-body provenance.");
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_TEXT_OVERFLOW"), "Unsupported chart text overflow should be diagnostic-covered from scene-owned text-body provenance.");
    }

    public static void PptxChartLegendAndDataLabelEllipsisDoesNotWarnWithoutObservedOverflow()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>
                <c:plotArea><c:lineChart>
                  <c:dLbls>
                    <c:txPr><a:bodyPr vertOverflow="ellipsis"/><a:lstStyle/><a:p/></c:txPr>
                    <c:showVal val="1"/>
                  </c:dLbls>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart></c:plotArea>
                <c:legend>
                  <c:legendPos val="r"/>
                  <c:overlay val="0"/>
                  <c:txPr><a:bodyPr vertOverflow="ellipsis"/><a:lstStyle/><a:p/></c:txPr>
                </c:legend>
              </c:chart>
            </c:chartSpace>
            """;
        (PptxDocument document, OoxPackage package) = PptxTests.BuildSingleChartPackage(chartXml);
        PptxSceneNodeSnapshot snapshot = PptxRenderer.InspectScene(document, package).Slides[0].SlideNodes[0];
        TestAssert.Equal(2, snapshot.ChartTextBodyVerticalOverflowCount);

        var diagnostics = new List<OoxPdfDiagnostic>();
        XDocument slideXmlWithoutBodyProperties = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree/></p:cSld>
            </p:sld>
            """);
        System.Reflection.MethodInfo emitDiagnostics = typeof(PptxRenderer).GetMethod(
            "EmitUnsupportedFeatureDiagnostics",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected unsupported feature diagnostic emitter.");
        Action<OoxPdfDiagnostic> sink = diagnostics.Add;
        emitDiagnostics.Invoke(null, [new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0], slideXmlWithoutBodyProperties, "/ppt/slides/slide1.xml", 1, sink]);

        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TEXT_OVERFLOW"), "Chart legend/data-label ellipsis defaults should not warn unless an overflowing chart text-frame surface is unsupported.");
    }

    public static void PptxUnsupportedEffectDiagnosticsUseSceneChartSeriesEffects()
    {
        IReadOnlyList<OoxPdfDiagnostic> diagnostics = PptxTests.ConvertSingleChartAndCollectDiagnostics("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:lineChart>
                <c:ser>
                  <c:spPr><a:effectLst><a:reflection/></a:effectLst></c:spPr>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_EFFECT"), "Unsupported chart series effects should be diagnostic-covered from scene-owned series effect provenance.");
    }

    public static void PptxUnsupportedEffectDiagnosticsUseSceneChartPointEffects()
    {
        IReadOnlyList<OoxPdfDiagnostic> diagnostics = PptxTests.ConvertSingleChartAndCollectDiagnostics("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:barChart>
                <c:ser>
                  <c:dPt><c:idx val="0"/><c:spPr><a:effectDag/></c:spPr></c:dPt>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:barChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_EFFECT"), "Unsupported chart point effects should be diagnostic-covered from scene-owned point effect provenance.");
    }
}
