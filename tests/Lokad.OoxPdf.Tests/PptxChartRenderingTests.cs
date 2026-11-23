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

    public static void PptxSceneScatterChartMarkerDefaultsUseDensityAndAutoSymbols()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:scatterChart>
                  <c:ser><c:xVal><c:numLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:ptCount val="5"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt></c:numLit></c:yVal></c:ser>
                  <c:ser><c:xVal><c:numLit><c:ptCount val="6"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt><c:pt idx="5"><c:v>6</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:ptCount val="6"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>5</c:v></c:pt><c:pt idx="5"><c:v>6</c:v></c:pt></c:numLit></c:yVal></c:ser>
                  <c:ser><c:marker><c:symbol val="star"/><c:size val="8"/></c:marker><c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser>
                </c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartPlot plot = chart?.Plots[0] ?? throw new InvalidOperationException("Expected scatter-chart plot.");
        TestAssert.Equal(3, plot.Series.Count);
        TestAssert.True(plot.Series[0].Marker.IsDefined == false, "Expected the sparse auto marker to stay implicit.");
        TestAssert.Equal("diamond", plot.Series[0].Marker.Symbol);
        TestAssert.Equal(9.9d, plot.Series[0].Marker.Size);
        TestAssert.True(plot.Series[1].Marker.IsDefined == false, "Expected the dense auto marker to stay implicit.");
        TestAssert.Equal("square", plot.Series[1].Marker.Symbol);
        TestAssert.Equal(7d, plot.Series[1].Marker.Size);
        TestAssert.True(plot.Series[2].Marker.IsDefined == true, "Expected explicit series marker XML to stay explicit.");
        TestAssert.Equal("star", plot.Series[2].Marker.Symbol);
        TestAssert.Equal(8d, plot.Series[2].Marker.Size);

        PptxSceneChart? bubble = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:bubbleChart>
                  <c:ser><c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:yVal><c:bubbleSize><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:bubbleSize></c:ser>
                </c:bubbleChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartPlot bubblePlot = bubble?.Plots[0] ?? throw new InvalidOperationException("Expected bubble-chart plot.");
        TestAssert.Equal("circle", bubblePlot.Series[0].Marker.Symbol);
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
        System.Reflection.MethodInfo readStyles = typeof(PptxRenderer).GetMethod(
            "ReadSceneMarkerStyles",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected marker-style resolver.");
        object[] sceneStyles = (((System.Collections.IEnumerable?)readStyles.Invoke(null, [plot])) ?? throw new InvalidOperationException("Expected scene marker styles.")).Cast<object>().ToArray();
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

        XElement alternatePlotXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea><c:lineChart>
                <c:marker val="1"/>
                <c:ser/>
                <c:ser><c:marker><c:symbol val="square"/><c:size val="8"/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill></a:ln></c:spPr></c:marker></c:ser>
              </c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(c + "lineChart").Single();
        PptxSceneChartPlot alternatePlot = (PptxTests.BuildSingleChartScene(alternatePlotXml.Document!.ToString())
            ?? throw new InvalidOperationException("Expected alternate marker scene.")).Plots[0];
        object[] alternateStyles = (((System.Collections.IEnumerable?)readStyles.Invoke(null, [alternatePlot])) ?? throw new InvalidOperationException("Expected alternate scene marker styles.")).Cast<object>().ToArray();
        TestAssert.Equal(2, alternateStyles.Length);
        TestAssert.Equal("diamond", PptxTests.ChartMarkerStyleSymbol(alternateStyles[0]));
        TestAssert.True(!PptxTests.ChartMarkerStyleIsDefined(alternateStyles[0]), "Expected alternate scene auto marker to remain distinct from explicit series marker XML.");
        TestAssert.Equal("square", PptxTests.ChartMarkerStyleSymbol(alternateStyles[1]));
        TestAssert.True(PptxTests.ChartMarkerStyleIsDefined(alternateStyles[1]), "Expected alternate scene explicit marker to preserve definition state.");
        TestAssert.Equal("8", PptxTests.ChartMarkerStyleSizeValue(alternateStyles[1]) ?? string.Empty);
        TestAssert.Equal(new RgbColor(171, 205, 239), PptxTests.ChartSeriesFillColor(PptxTests.ChartMarkerStyleFill(alternateStyles[1])));
        TestAssert.Equal(new RgbColor(254, 220, 186), PptxTests.ChartSeriesStrokeColor(PptxTests.ChartMarkerStyleStroke(alternateStyles[1])));
        TestAssert.Equal(1d, PptxTests.ChartSeriesStrokeWidth(PptxTests.ChartMarkerStyleStroke(alternateStyles[1])));
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
                    </c:numLit></c:val><c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:dLblPos val="t"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800" b="1" i="1"><a:solidFill><a:srgbClr val="0066AA"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbls><c:marker><c:symbol val="star"/><c:size val="9"/></c:marker></c:ser>
                    <c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/><c:dLblPos val="b"/><c:separator> | </c:separator><c:spPr><a:solidFill><a:srgbClr val="FFEACC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="0A0B0C"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr><c:dLbl><c:idx val="1"/><c:showLegendKey val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="850" b="1"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:rPr><a:t>ZX</a:t></a:r><a:r><a:rPr i="1"><a:latin typeface="Arial"/></a:rPr><a:t>Q</a:t></a:r></a:p></c:rich></c:tx><c:dLblPos val="ctr"/><c:layout><c:manualLayout><c:x val="0.44"/><c:y val="0.28"/><c:w val="0.1"/><c:h val="0.08"/></c:manualLayout></c:layout><c:spPr><a:solidFill><a:srgbClr val="CCEEFF"/></a:solidFill></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="6600CC"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbl></c:dLbls>
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
                    </c:numLit></c:val><c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showLeaderLines val="1"/><c:leaderLines><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="44CC88"/></a:solidFill></a:ln></c:spPr></c:leaderLines><c:dLbl><c:idx val="1"/><c:showLegendKey val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:layout><c:manualLayout><c:xMode val="factor"/><c:yMode val="factor"/><c:x val="0.1"/><c:y val="0.1"/></c:manualLayout></c:layout><c:leaderLines><c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="CC11AA"/></a:solidFill></a:ln></c:spPr></c:leaderLines></c:dLbl></c:dLbls></c:ser>
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

        var requestedFonts = new List<FontRequest>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions
        {
            DiagnosticSink = collector.Add,
            FontResolver = new TestFaceFontResolver(requestedFonts.Add)
        });

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
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <007C>"), "Expected line-chart data labels to include the custom separator between series, category, and value text.");
        TestAssert.Contains("1 0.918 0.8 rg", pdf);
        TestAssert.Contains("0.067 0.133 0.2 RG", pdf);
        Match labelShape = Regex.Match(pdf, @"1 0\.918 0\.8 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) (?<width>[\d.]+) (?<height>[\d.]+) re f");
        Match firstMarker = Regex.Match(pdf, @"0 0 0\.667 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) 8 8 re f");
        TestAssert.True(labelShape.Success && firstMarker.Success,
            "Expected the styled label background and its first source point marker.");
        double labelWidth = double.Parse(labelShape.Groups["width"].Value, CultureInfo.InvariantCulture);
        TestAssert.True(labelWidth >= 28.8d && labelWidth <= 144d,
            "The automatic box admits its composed text inside the authored plot width; the old span minimum is not a text-width limit.");
        double labelTop = double.Parse(labelShape.Groups["y"].Value, CultureInfo.InvariantCulture) +
            double.Parse(labelShape.Groups["height"].Value, CultureInfo.InvariantCulture);
        TestAssert.True(labelTop < double.Parse(firstMarker.Groups["y"].Value, CultureInfo.InvariantCulture),
            "The source-authored bottom position reserves the first marker extent and places its occupied text frame below it.");
        TestAssert.True(Math.Abs(double.Parse(labelShape.Groups["x"].Value, CultureInfo.InvariantCulture) + labelWidth / 2d -
            (double.Parse(firstMarker.Groups["x"].Value, CultureInfo.InvariantCulture) + 4d)) < .001d,
            "The automatic label box stays centered on its plotted category point inside the authored plot frame.");
        TestAssert.Contains("0.8 0.933 1 rg", pdf);
        Match sizedShape = Regex.Match(pdf, @"0\.8 0\.933 1 rg\s+(?<x>[-\d.]+) (?<y>[-\d.]+) 28\.8 17\.28 re f");
        TestAssert.True(sizedShape.Success, "The authored0.1Ãƒâ€”0.08 label scales with its288Ãƒâ€”216pt chart frame.");
        TestAssert.Contains("0.4 0 0.8 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CLD[0-9]+ 9 Tf"), "Expected per-label custom data label text style to drive font size.");
        TestAssert.Contains("0.071 0.204 0.337 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CT[0-9]+ 12\.96 Tf"), "Expected chart title txPr font size to drive title rendering.");
        TestAssert.Contains("0.671 0.804 0.937 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CT[0-9]+ 11\.52 Tf"), "Expected first rich-text title run to use its run font size.");
        TestAssert.Contains("0.396 0.263 0.129 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CL[0-9]+ 6\.96 Tf"), "Expected chart legend txPr font size to drive legend rendering.");
        Match sizedText = Regex.Match(pdf, @"/CLD[0-9]+ 9 Tf\s+\S+ Tc\s+1 0 0 1 (?<x>[-\d.]+) (?<y>[-\d.]+) Tm");
        TestAssert.True(sizedText.Success &&
            double.Parse(sizedText.Groups["y"].Value, CultureInfo.InvariantCulture) > double.Parse(sizedShape.Groups["y"].Value, CultureInfo.InvariantCulture) &&
            double.Parse(sizedText.Groups["y"].Value, CultureInfo.InvariantCulture) < double.Parse(sizedShape.Groups["y"].Value, CultureInfo.InvariantCulture) + 17.28d,
            "The authored sized label contains its measured text baseline.");
        TestAssert.Contains("0.671 0.804 0.937 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CLD[0-9]+ 8\.52 Tf"), "Expected first custom rich-text data-label run to use its run font size.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <005A>") &&
            Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0058>") &&
            Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0051>"),
            "Expected per-label custom rich text to be emitted in the chart data label ToUnicode map.");
        TestAssert.Contains("0 0.4 0.667 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CLD[0-9]+ 8\.04 Tf"), "Expected second-series data labels to consume the series-level font size.");
        // The portable face has regular outlines. Check the chart's style requests
        // rather than an installed Arial file's particular post-table angle.
        TestAssert.True(requestedFonts.Any(request => request.FamilyName == "Arial" && request.Italic && !request.Bold),
            "The italic legend and custom label run must request the italic face.");
        TestAssert.True(requestedFonts.Any(request => request.FamilyName == "Arial" && request.Italic && request.Bold),
            "The bold-italic series labels must request both styles together.");
        TestAssert.Contains("0.039 0.043 0.047 rg", pdf);
        TestAssert.Contains("/CLD1 9.96 Tf", pdf);
        Match labelText = Regex.Match(pdf, @"/CLD1 9\.96 Tf\s+\S+ Tc\s+1 0 0 1 (?<x>[-\d.]+) (?<y>[-\d.]+) Tm");
        TestAssert.True(labelText.Success &&
            double.Parse(labelText.Groups["x"].Value, CultureInfo.InvariantCulture) > double.Parse(labelShape.Groups["x"].Value, CultureInfo.InvariantCulture) &&
            double.Parse(labelText.Groups["x"].Value, CultureInfo.InvariantCulture) < double.Parse(labelShape.Groups["x"].Value, CultureInfo.InvariantCulture) + labelWidth &&
            double.Parse(labelText.Groups["y"].Value, CultureInfo.InvariantCulture) > double.Parse(labelShape.Groups["y"].Value, CultureInfo.InvariantCulture) &&
            double.Parse(labelText.Groups["y"].Value, CultureInfo.InvariantCulture) < labelTop,
            "The composed source label baseline lies inside its padded text background.");
        // Auto pie labels emit no leaders (Office draws none for the styled auto labels
        // on the auto-radius reference); only the factor-manual label below keeps one.
        TestAssert.DoesNotContain("0.267 0.8 0.533 RG", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"0\.8 0\.067 0\.667 RG\s+2\.25 w[\s\S]+? m\s+[\s\S]+? l\s+[\s\S]+? l\s+S"),
            "Expected per-label polar leader-line style to override the chart-wide leader-line style.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0025>"), "Expected percentage labels to include a percent glyph in the ToUnicode map.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <002C>"), "Expected combined value/percentage labels to include the default comma separator.");
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
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Sparse pie charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticPieAutoLabelsEmitNoLeaders()
    {
        // Office draws no leaders for styled auto pie labels (auto-radius reference
        // carries a red 2.25pt leader style yet zero leader strokes); only manual
        // labels keep them.
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
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>30</c:v></c:pt><c:pt idx="1"><c:v>70</c:v></c:pt></c:numLit></c:val>
                      <c:dLbls><c:showVal val="1"/><c:showLeaderLines val="1"/><c:leaderLines><c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:leaderLines></c:dLbls>
                    </c:ser>
                  </c:pieChart></c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.DoesNotContain("1 0 0 RG", pdf);
        TestAssert.Contains("BT", pdf);
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

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add, FontResolver = new TestFaceFontResolver() });

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
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Area, scatter, radar, and doughnut charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticRadarMarkerOutlineUsesOfficeDefaultWidth()
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
                      <p:xfrm><a:off x="457200" y="457200"/><a:ext cx="1828800" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:radarChart>
                  <c:radarStyle val="marker"/>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:radarChart></c:plotArea></c:chart></c:chartSpace>
                """),
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("3.75 w", pdf);
        TestAssert.DoesNotContain("1.2 w", pdf);
    }

    // RV04: radar grid spokes must paint over the series fill.
    public static void RadarSpokesPaintOverSeriesFill()
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
                      <p:xfrm><a:off x="457200" y="457200"/><a:ext cx="1828800" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:radarChart>
                  <c:radarStyle val="filled"/>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:radarChart></c:plotArea></c:chart></c:chartSpace>
                """),
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var ops = new System.Text.StringBuilder();
        foreach (string line in pdf.Split((char)10))
        {
            string t = line.Trim();
            if (t.EndsWith(" m") || t.EndsWith(" l") || t == "S" || t == "f")
            {
                ops.Append(t.Length > 24 ? t.Substring(t.Length - 24) : t).Append((char)124);
            }
        }
        string content = pdf.Replace("\r\n", "\n");
        int fillAt = content.LastIndexOf("\nf\n");
        int lastStrokeAt = content.LastIndexOf("\nS\n");
        TestAssert.True(fillAt >= 0 && lastStrokeAt >= 0, "Filled radar must emit grid, spokes and series fill.");
        TestAssert.True(fillAt < lastStrokeAt, "Series fill must paint before grid strokes so spokes stay visible.");
    }


    // RV04: style-118 filled-radar series must paint a vertical gradient, not a flat fill.
    public static void RadarStyle118FilledSeriesPaintsVerticalGradient()
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
                      <p:xfrm><a:off x="457200" y="457200"/><a:ext cx="1828800" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><mc:AlternateContent><mc:Choice Requires="c14"><c14:style val="118"/></mc:Choice><mc:Fallback><c:style val="18"/></mc:Fallback></mc:AlternateContent><c:chart><c:plotArea><c:radarChart>
                  <c:radarStyle val="filled"/>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:radarChart></c:plotArea></c:chart></c:chartSpace>
                """),
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Sh1 sh", pdf);
    }
    // RV04: the derived radar dark stop paints Office-measured bytes (series 1 of the
    // filled ladder deck resolves 4F81BD outside any sampled table).
    public static void RadarDerivedDarkStopPaintsOfficeDarkBase()
    {
        string input = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "pptx-ladder-11-chart-radar-filled-port.pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"/Sh\d+ sh"), "Radar gradient must paint shading resources.");
        TestAssert.Contains("/C1 [0.247 0.502 0.804]", pdf);
    }

    // RV04: the derived radar dark stop is byte-exact on the 16-face Office corpus.
    public static void RadarDerivedDarkStopsMatchOfficeExactly()
    {
        foreach (string row in RadarDarkStopCorpus)
        {
            string[] parts = row.Split((char)58);
            TestAssert.True(RgbColor.TryParse(parts[0], out RgbColor baseColor), row);
            TestAssert.True(RgbColor.TryParse(parts[1], out RgbColor expected), row);
            TestAssert.Equal(expected, PptxRenderer.DeriveRadarDarkStop(baseColor));
        }
    }

    private static readonly string[] RadarDarkStopCorpus =
    [
        "4F81BD:3F80CD",
        "C0504D:D1403C",
        "FF0000:FF0000",
        "00FF00:00FF00",
        "0000FF:0000FF",
        "FFFF00:FFFF00",
        "00FFFF:00FFFF",
        "FF00FF:FF00FF",
        "000000:000000",
        "FFFFFF:FFFFFF",
        "808080:808080",
        "404040:404040",
        "C0C0C0:C0C0C0",
        "00C000:00DD00",
        "FF4000:FF2D00",
        "FFC800:FFDE00",
    ];

    // RV04: the derived radar light stop matches Office-measured bytes on the
    // theme-recolor corpus (artifacts/rv04radar); the legacy linear top misses
    // these faces by up to 23 levels.
    public static void RadarDerivedLightStopsMatchOfficeExactly()
    {
        foreach (string row in RadarLightStopCorpus)
        {
            string[] parts = row.Split((char)58);
            TestAssert.True(RgbColor.TryParse(parts[0], out RgbColor baseColor), row);
            TestAssert.True(RgbColor.TryParse(parts[1], out RgbColor expected), row);
            TestAssert.Equal(expected, PptxRenderer.DeriveRadarLightStop(baseColor));
        }
    }

    private static readonly string[] RadarLightStopCorpus =
    [
        "000000:BCBCBC",
        "808080:CDCDCD",
        "0000FF:6767FF",
        "00FF00:67FF67",
        "FF0000:FF6767",
        "208050:A9E0BC",
        "208080:A9E0E0",
        "206040:B2CFBD",
        "607060:C1CDC1",
        "60A080:AFECCA",
        "107010:ACD8AC",
        "156B15:AED5AE",
        "176917:AFD4AF",
        "404040:C0C0C0",
        "488848:AFE0AF",
        "589858:AFE8AF",
        "608860:B9DBB9",
        "202020:BDBDBD",
    ];
    // RV04: style-18 radar fills paint a 5-stop forward stitch whose interior
    // knots match Office sampled-shading quarter samples: q1/mid within 2 levels,
    // q3 within 4 (the min-channel lower knot is the documented weak joint).
    public static void RadarSeriesGradientKnotsMatchOfficeSamples()
    {
        AssertKnots("208050", (158, 214, 177), (123, 188, 146), (70, 157, 104));
        AssertKnots("406040", (173, 193, 173), (139, 162, 139), (91, 122, 91));
        AssertKnots("309060", (155, 224, 181), (123, 199, 152), (73, 171, 115));
    }

    private static void AssertKnots(string baseHex, (int, int, int) q1, (int, int, int) mid, (int, int, int) q3)
    {
        TestAssert.True(RgbColor.TryParse(baseHex, out RgbColor baseColor), baseHex);
        IReadOnlyList<PdfShadingStop> stops = PptxRenderer.RadarSeriesGradientStops(baseColor);
        TestAssert.Equal(5, stops.Count);
        AssertKnot(stops[1], q1, baseHex, 2);
        AssertKnot(stops[2], mid, baseHex, 2);
        AssertKnot(stops[3], q3, baseHex, 4);
    }

    private static void AssertKnot(PdfShadingStop stop, (int, int, int) expected, string face, int tolerance)
    {
        TestAssert.True(Math.Abs(stop.Red - expected.Item1) <= tolerance, face);
        TestAssert.True(Math.Abs(stop.Green - expected.Item2) <= tolerance, face);
        TestAssert.True(Math.Abs(stop.Blue - expected.Item3) <= tolerance, face);
    }
    public static void PptxSyntheticRadarFilledOutlineNeedsExplicitLine()
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
                      <p:xfrm><a:off x="457200" y="457200"/><a:ext cx="1828800" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:radarChart>
                  <c:radarStyle val="filled"/>
                  <c:ser><c:spPr><a:ln w="47625"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:radarChart></c:plotArea></c:chart></c:chartSpace>
                """),
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("3.75 w", pdf);
        TestAssert.DoesNotContain("1.2 w", pdf);
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
        TestAssert.True(collector.Diagnostics.Any(d => d.Id == "PPTX_CHART_MISSING_CACHED_DATA"), "Formula-only embedded workbook references should report missing chart-side caches while preserving workbook values as provenance.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Formula-only supported chart families should report missing cached data rather than a generic unsupported chart.");
    }

    public static void PptxChartParsedVectorsPreserveWorkbookSidecarPoints()
    {
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
        object workbook = PptxChartRangeTests.ReadWorkbookCells(sheets);

        PptxSceneChartPlot ParsePlot(XElement element) => TestAssert.NotNull(PptxTests.BuildSingleChartScene(
            $"""<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea>{element}</c:plotArea></c:chart></c:chartSpace>""")).Plots.Single();

        var readNumberVector = typeof(PptxRenderer)
            .GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(method => method.Name == "ReadSceneChartSeriesVectors" && method.GetParameters().Length == 3);
        var readCategoryVector = typeof(PptxRenderer)
            .GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(method => method.Name == "ReadSceneCategoryLabelVector" && method.GetParameters().Length == 3);

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
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        PptxSceneChartPlot numberPlot = ParsePlot(new XElement(c + "barChart", new XElement(c + "ser", values)));
        object numberVector = ((System.Collections.IEnumerable)(readNumberVector.Invoke(null, [numberPlot, workbook, true])
            ?? throw new InvalidOperationException("Expected parsed number vectors."))).Cast<object>().Single();
        object[] cacheNumberPoints = ((System.Collections.IEnumerable?)numberVector.GetType().GetProperty("Points")?.GetValue(numberVector))?.Cast<object>().ToArray() ?? [];
        object[] workbookNumberPoints = ((System.Collections.IEnumerable?)numberVector.GetType().GetProperty("WorkbookPoints")?.GetValue(numberVector))?.Cast<object>().ToArray() ?? [];
        TestAssert.True(cacheNumberPoints.Length == 2, "Expected parsed vector rendering points to stay bound to the chart cache.");
        TestAssert.True((int?)cacheNumberPoints[0].GetType().GetProperty("Index")?.GetValue(cacheNumberPoints[0]) == 2, "Expected parsed numeric cache point index to survive.");
        TestAssert.Equal("OoxmlIndex", cacheNumberPoints[0].GetType().GetProperty("IndexSource")?.GetValue(cacheNumberPoints[0])?.ToString() ?? string.Empty);
        TestAssert.Equal(100d, (double?)cacheNumberPoints[0].GetType().GetProperty("Value")?.GetValue(cacheNumberPoints[0]) ?? 0d);
        TestAssert.True((int?)cacheNumberPoints[1].GetType().GetProperty("Index")?.GetValue(cacheNumberPoints[1]) == 1, "Expected malformed raw numeric cache point index to keep the ordinal fallback.");
        TestAssert.Equal("OrdinalFallback", cacheNumberPoints[1].GetType().GetProperty("IndexSource")?.GetValue(cacheNumberPoints[1])?.ToString() ?? string.Empty);
        TestAssert.True(workbookNumberPoints.Length == 3, "Expected parsed numeric vector to keep workbook sidecar points without promoting them to renderable cache points.");
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
        object categoryVector = readCategoryVector.Invoke(null, [ParsePlot(chart), workbook, true]) ?? throw new InvalidOperationException("Expected parsed category vector.");
        object[] cacheTextPoints = ((System.Collections.IEnumerable?)categoryVector.GetType().GetProperty("Points")?.GetValue(categoryVector))?.Cast<object>().ToArray() ?? [];
        object[] workbookTextPoints = ((System.Collections.IEnumerable?)categoryVector.GetType().GetProperty("WorkbookPoints")?.GetValue(categoryVector))?.Cast<object>().ToArray() ?? [];
        TestAssert.True(cacheTextPoints.Length == 2, "Expected parsed category rendering points to stay bound to the chart cache.");
        TestAssert.Equal("Cached", (string?)cacheTextPoints[0].GetType().GetProperty("Text")?.GetValue(cacheTextPoints[0]) ?? string.Empty);
        TestAssert.Equal("OoxmlIndex", cacheTextPoints[0].GetType().GetProperty("IndexSource")?.GetValue(cacheTextPoints[0])?.ToString() ?? string.Empty);
        TestAssert.True((int?)cacheTextPoints[1].GetType().GetProperty("Index")?.GetValue(cacheTextPoints[1]) == 1, "Expected malformed raw text cache point index to keep the ordinal fallback.");
        TestAssert.Equal("OrdinalFallback", cacheTextPoints[1].GetType().GetProperty("IndexSource")?.GetValue(cacheTextPoints[1])?.ToString() ?? string.Empty);
        TestAssert.True(workbookTextPoints.Length == 3, "Expected parsed category vector to keep workbook sidecar text points.");
        TestAssert.Equal("WorkbookRange", workbookTextPoints[0].GetType().GetProperty("IndexSource")?.GetValue(workbookTextPoints[0])?.ToString() ?? string.Empty);
        TestAssert.Equal("West", (string?)workbookTextPoints[2].GetType().GetProperty("Text")?.GetValue(workbookTextPoints[2]) ?? string.Empty);
    }

    public static void PptxChartWorkbookRangesPreserveValuesVisibilityAndFormats()
    {
        using var input = new MemoryStream(PptxTests.EmbeddedChartWorkbook());
        using var package = OoxPackage.Open(input, default);
        var workbook = PptxChartRangeTests.ReadWorkbook(package);
        string[] amounts = ["8.2", "3.2", "1.4"];
        foreach (var (formula, expected) in new (string, string[])[] {
            ("Sheet1!$B$2:$B$4", amounts), ("'[Book.xlsx]Sheet1'!$B$2:$B$4", amounts),
            ("Sheet1!$B$2:$B$2,$B$4:$B$4", ["8.2", "1.4"]),
            ("SalesValues", amounts), ("SalesUnionValues", ["8.2", "1.4"]),
            ("SalesLabels", ["North", "South", "West"]),
            ("Sheet1!SheetLocalValues", amounts), ("SheetLocalValues", []),
            ("SalesTable[Amount]", amounts), ("Sales_Table[Amount]", amounts),
            ("SalesTable[#Data]", ["North", "8.2", "", "South", "3.2", "7.5", "West", "1.4", "9.1"]),
            ("SalesTable[#Headers]", ["", "Share", "Quoted ] Amount"]),
            ("SalesTotalsTable[#All]", ["", "Share", "North", "8.2", "South", "3.2", "West", "1.4"]),
            ("SalesTable[[Region]:[Amount]]", ["North", "8.2", "South", "3.2", "West", "1.4"]),
            ("SalesTable[Quoted ']' Amount]", ["", "7.5", "9.1"]),
            ("SalesTotalsTable[[#Totals],[Amount]]", ["1.4"]),
            ("Sheet1!$C$2:$C$2", [""]), ("UnknownTable[Amount]", []), ("", []) })
        {
            var cells = workbook.ReadRangeCells(formula);
            TestAssert.Equal(string.Join("|", expected), string.Join("|", cells.Select(cell => cell.Text)));
            TestAssert.Equal(expected.Length, cells.Length);
            for (int index = 0; index < cells.Length; index++)
            {
                TestAssert.Equal(index, cells[index].Index);
                TestAssert.Equal(expected[index].Length > 0, cells[index].HasValue);
            }
        }
        var values = workbook.ReadRangeCells("SalesValues");
        TestAssert.True(values.All(cell => cell.ColumnHidden), "Column B is authored hidden.");
        TestAssert.True(!values[0].RowHidden && values[1].RowHidden && !values[2].RowHidden,
            "Only the authored third row is hidden.");
        TestAssert.Equal("m/d/yy", values[0].StyleNumberFormatCode);
        TestAssert.Equal(true, values[0].StyleAppliesNumberFormat);
        var labels = workbook.ReadRangeCells("SalesLabels");
        TestAssert.True(labels.All(cell => !cell.ColumnHidden), "Column A stays visible beside the hidden column.");
        var blank = workbook.ReadRangeCells("Sheet1!C2").Single();
        TestAssert.True(!blank.HasValue && blank.Text == "", "A formula without a stored result is not recalculated.");
        TestAssert.True(ReferenceEquals(values, workbook.ReadRangeCells("SalesValues")), "Range memoization retains the owned immutable result.");
        var empty = PptxChartRangeTests.ReadWorkbookCells(
            new Dictionary<string, Dictionary<string, string>> { ["Sheet1"] = new() { ["A1"] = "" } });
        TestAssert.True(empty.ReadRangeCells("Sheet1!A1").Single().HasValue,
            "An authored empty stored value differs from a missing result.");
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
        object workbook = PptxChartRangeTests.ReadWorkbookCells(sheets);
        var cachedPlot = ReadPointSourcePlot("""
            <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode>
            <c:ptCount val="1"/><c:pt idx="0"><c:v>99</c:v></c:pt></c:numCache></c:numRef></c:val>
            """);
        var readNumbers = typeof(PptxRenderer).GetMethod("ReadSceneChartSeriesVectors",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        object NumberVector(PptxSceneChartPlot plot, object? book) =>
            ((System.Collections.IEnumerable)readNumbers.Invoke(null, [plot, book, true])!).Cast<object>().Single();
        object vector = NumberVector(cachedPlot, workbook);

        object[] activePoints = (((System.Collections.IEnumerable?)vector.GetType().GetProperty("Points")?.GetValue(vector)) ?? throw new InvalidOperationException("Expected active chart points.")).Cast<object>().ToArray();
        object[] workbookPoints = (((System.Collections.IEnumerable?)vector.GetType().GetProperty("WorkbookPoints")?.GetValue(vector)) ?? throw new InvalidOperationException("Expected workbook sidecar points.")).Cast<object>().ToArray();
        TestAssert.True(activePoints.Length == 1, "Expected existing chart-cache point projection to remain active.");
        TestAssert.True((double?)activePoints[0].GetType().GetProperty("Value")?.GetValue(activePoints[0]) == 99d, "Expected chart-cache value to remain the rendered point value.");
        object[] densePoints = (((System.Collections.IEnumerable?)vector.GetType().GetMethod("DenseValues")?.Invoke(vector, [])) ?? throw new InvalidOperationException("Expected dense active point projection.")).Cast<object>().ToArray();
        TestAssert.True(densePoints.Length == 1, "Expected dense point projection to preserve rendered point count.");
        TestAssert.True((double?)densePoints[0] == 99d, "Expected compact geometry slots to preserve rendered cache values.");
        System.Reflection.MethodInfo buildPointIndex = typeof(PptxRenderer).GetMethod(
            "BuildDensePointIndex", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            ?? throw new InvalidOperationException("Expected rendering provenance index.");
        var pointIndex = (System.Collections.IDictionary)buildPointIndex.Invoke(null, [vector])!;
        TestAssert.True(activePoints[0].Equals(pointIndex[0]), "Expected rendering index to preserve the full active source point.");
        TestAssert.True(workbookPoints.Length == 2, "Expected workbook sidecar points to preserve workbook source values next to cached chart points.");
        TestAssert.True((int?)workbookPoints[0].GetType().GetProperty("Index")?.GetValue(workbookPoints[0]) == 0, "Expected workbook sidecar point index to preserve range order.");
        TestAssert.True((double?)workbookPoints[0].GetType().GetProperty("Value")?.GetValue(workbookPoints[0]) == 8.2d, "Expected workbook sidecar point to preserve workbook numeric value.");
        object workbookCell = workbookPoints[0].GetType().GetProperty("WorkbookCell")?.GetValue(workbookPoints[0]) ?? throw new InvalidOperationException("Expected workbook sidecar point cell provenance.");
        TestAssert.Equal("8.2", ((PptxRenderer.ChartWorkbookRangeCell)workbookCell).Text);
        TestAssert.Equal(0, ((PptxRenderer.ChartWorkbookRangeCell)workbookCell).Index);
        var noCachePlot = ReadPointSourcePlot("""
            <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val>
            """);
        object noCacheVector = NumberVector(noCachePlot, workbook);
        object[] noCacheActivePoints = (((System.Collections.IEnumerable?)noCacheVector.GetType().GetProperty("Points")?.GetValue(noCacheVector)) ?? throw new InvalidOperationException("Expected formula-only active chart points.")).Cast<object>().ToArray();
        object[] noCacheWorkbookPoints = (((System.Collections.IEnumerable?)noCacheVector.GetType().GetProperty("WorkbookPoints")?.GetValue(noCacheVector)) ?? throw new InvalidOperationException("Expected formula-only workbook sidecar points.")).Cast<object>().ToArray();
        object[] noCacheDensePoints = (((System.Collections.IEnumerable?)noCacheVector.GetType().GetMethod("DenseValues")?.Invoke(noCacheVector, [])) ?? throw new InvalidOperationException("Expected formula-only dense active point projection.")).Cast<object>().ToArray();
        TestAssert.True(noCacheActivePoints.Length == 0, "Expected formula-only workbook numeric values to remain absent from cache-owned active points.");
        TestAssert.True(noCacheDensePoints.Length == 0, "Expected formula-only workbook numeric values to stay out of dense rendered point slots when chart caches are absent.");
        TestAssert.True(noCacheWorkbookPoints.Length == 2, "Expected formula-only workbook numeric values to remain available as sidecar points.");
        System.Reflection.MethodInfo buildPieSlices = typeof(PptxRenderer).GetMethod(
            "BuildChartIndexedPieSlices",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected indexed pie-slice builder.");
        object?[] pieArguments = [vector, null];
        object[] pieSlices = (((System.Collections.IEnumerable?)buildPieSlices.Invoke(null, pieArguments)) ?? throw new InvalidOperationException("Expected indexed pie slices.")).Cast<object>().ToArray();
        TestAssert.Equal(1, (int)pieArguments[1]!);
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
        var radarSeries = (double?[][])buildRadarSeries.Invoke(null, [typedNumberVectors])!;
        TestAssert.Equal(99d, radarSeries.Single().Single());
        var sparseRadarPlot = ReadPointSourcePlot("""
            <c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/>
            <c:pt idx="2"><c:v>44</c:v></c:pt></c:numLit></c:val>
            """);
        object sparseRadarVector = NumberVector(sparseRadarPlot, null);
        Array sparseRadarVectors = Array.CreateInstance(vector.GetType(), 1);
        sparseRadarVectors.SetValue(sparseRadarVector, 0);
        var sparseRadarSeries = (double?[][])buildRadarSeries.Invoke(null, [sparseRadarVectors])!;
        double?[] sparseRadarValues = sparseRadarSeries.Single();
        TestAssert.True(sparseRadarValues.Length == 3, "Expected radar geometry projection to preserve sparse point slots.");
        TestAssert.True(sparseRadarValues[0] is null && sparseRadarValues[1] is null, "Expected missing radar point slots to remain explicit gaps.");
        TestAssert.True(sparseRadarValues[2] == 44d, "Expected sparse radar value to keep its source value.");
        var readScatter = typeof(PptxRenderer).GetMethod("ReadSceneScatterSeriesVectors",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        object ScatterVector(PptxSceneChartPlot plot, object? book) =>
            ((System.Collections.IEnumerable)readScatter.Invoke(null, [plot, true, book, true])!).Cast<object>().Single();
        var bubblePlot = ReadPointSourcePlot("""
            <c:xVal><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="1"/>
            <c:pt idx="0"><c:v>99</c:v></c:pt></c:numCache></c:numRef></c:xVal>
            <c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="1"/>
            <c:pt idx="0"><c:v>99</c:v></c:pt></c:numCache></c:numRef></c:yVal>
            <c:bubbleSize><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="1"/>
            <c:pt idx="0"><c:v>99</c:v></c:pt></c:numCache></c:numRef></c:bubbleSize>
            """, "bubbleChart");
        object scatterSeries = ScatterVector(bubblePlot, workbook);
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
        var sparseBubblePlot = ReadPointSourcePlot("""
            <c:xVal><c:numLit><c:ptCount val="3"/><c:pt idx="2"><c:v>12</c:v></c:pt></c:numLit></c:xVal>
            <c:yVal><c:numLit><c:ptCount val="3"/><c:pt idx="2"><c:v>34</c:v></c:pt></c:numLit></c:yVal>
            <c:bubbleSize><c:numLit><c:ptCount val="3"/><c:pt idx="2"><c:v>5</c:v></c:pt></c:numLit></c:bubbleSize>
            """, "bubbleChart");
        object sparseScatterSeriesRecord = ScatterVector(sparseBubblePlot, null);
        object sparseRenderedScatter = buildScatterSeries.Invoke(null, [sparseScatterSeriesRecord]) ?? throw new InvalidOperationException("Expected sparse rendered scatter-series.");
        object[] sparseScatterPoints = (((System.Collections.IEnumerable?)sparseRenderedScatter.GetType().GetProperty("Points")?.GetValue(sparseRenderedScatter)) ?? throw new InvalidOperationException("Expected sparse rendered scatter points.")).Cast<object>().ToArray();
        TestAssert.True(sparseScatterPoints.Length == 1, "Expected sparse scatter channels to pair by source index.");
        TestAssert.True((int?)sparseScatterPoints[0].GetType().GetProperty("Index")?.GetValue(sparseScatterPoints[0]) == 2, "Expected sparse scatter point to preserve the paired source index.");
        TestAssert.True((double?)sparseScatterPoints[0].GetType().GetProperty("Size")?.GetValue(sparseScatterPoints[0]) == 5d, "Expected sparse bubble-size channel to pair by source index.");
        var threeSlotPlot = ReadPointSourcePlot("""
            <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode>
            <c:ptCount val="3"/><c:pt idx="0"><c:v>99</c:v></c:pt></c:numCache></c:numRef></c:val>
            """);
        object blankVector = NumberVector(threeSlotPlot, workbook);
        object[] blankAwareWorkbookPoints = (((System.Collections.IEnumerable?)blankVector.GetType().GetProperty("WorkbookPoints")?.GetValue(blankVector)) ?? throw new InvalidOperationException("Expected blank-aware workbook sidecar points.")).Cast<object>().ToArray();
        TestAssert.True(blankAwareWorkbookPoints.Length == 3, "Expected workbook sidecar points to preserve missing cells inside the source range.");
        TestAssert.True(blankAwareWorkbookPoints[2].GetType().GetProperty("Value")?.GetValue(blankAwareWorkbookPoints[2]) is null, "Expected missing workbook source cells to remain nullable sidecar points.");
        object blankWorkbookCell = blankAwareWorkbookPoints[2].GetType().GetProperty("WorkbookCell")?.GetValue(blankAwareWorkbookPoints[2]) ?? throw new InvalidOperationException("Expected blank workbook cell provenance.");
        TestAssert.True(!((PptxRenderer.ChartWorkbookRangeCell)blankWorkbookCell).HasValue,
            "Missing workbook source cells cannot supply a stored value.");

        using MemoryStream embeddedWorkbookStream = new(PptxTests.EmbeddedChartWorkbook());
        using OoxPackage embeddedWorkbookPackage = OoxPackage.Open(embeddedWorkbookStream, CancellationToken.None);
        var parsedWorkbook = PptxChartRangeTests.ReadWorkbook(embeddedWorkbookPackage);
        object hiddenColumnVector = NumberVector(threeSlotPlot, parsedWorkbook);
        object[] allHiddenColumnWorkbookPoints = (((System.Collections.IEnumerable?)hiddenColumnVector.GetType().GetMethod("WorkbookPointsForPlotVisibility")?.Invoke(hiddenColumnVector, [false])) ?? throw new InvalidOperationException("Expected unfiltered workbook points.")).Cast<object>().ToArray();
        object[] visibleHiddenColumnWorkbookPoints = (((System.Collections.IEnumerable?)hiddenColumnVector.GetType().GetMethod("WorkbookPointsForPlotVisibility")?.Invoke(hiddenColumnVector, [true])) ?? throw new InvalidOperationException("Expected visibility-filtered workbook points.")).Cast<object>().ToArray();
        TestAssert.True(allHiddenColumnWorkbookPoints.Length == 3, "Expected plot visibility projection to start from the full workbook source vector.");
        TestAssert.True(visibleHiddenColumnWorkbookPoints.Length == 0, "Expected plot-visible-only projection to exclude workbook points from a hidden source column.");
        object? hiddenWorkbookPoint = hiddenColumnVector.GetType().GetMethod("WorkbookPointForIndex")?.Invoke(hiddenColumnVector, [0]);
        TestAssert.True(hiddenWorkbookPoint is null, "Expected indexed workbook source lookup to honor the vector's plot-visible-only policy.");
        var threeSlotNoCachePlot = ReadPointSourcePlot("""
            <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef></c:val>
            """);
        object noCacheHiddenColumnVector = NumberVector(threeSlotNoCachePlot, parsedWorkbook);
        object[] noCacheHiddenDensePoints = (((System.Collections.IEnumerable?)noCacheHiddenColumnVector.GetType().GetMethod("DenseValues")?.Invoke(noCacheHiddenColumnVector, [])) ?? throw new InvalidOperationException("Expected visibility-filtered no-cache dense points.")).Cast<object>().ToArray();
        TestAssert.True(noCacheHiddenDensePoints.Length == 0, "Expected no-cache dense projection to honor plot-visible-only when the workbook source column is hidden.");

        var labelPlot = ReadPointSourcePlot("""
            <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>
            <c:pt idx="0"><c:v>Cached</c:v></c:pt></c:strCache></c:strRef></c:cat>
            """);
        var readText = typeof(PptxRenderer).GetMethod("BuildChartIndexedTextVector",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        object TextVector(PptxSceneChartPlot plot)
        {
            var series = plot.Series.Single();
            return readText.Invoke(null, [series.CategoryPoints, series.CategoryPointCount, series.CategoryLevels,
                series.DataSources.Categories, parsedWorkbook, true])!;
        }
        object textVector = TextVector(labelPlot);
        object[] visibleLabelWorkbookPoints = (((System.Collections.IEnumerable?)textVector.GetType().GetMethod("WorkbookPointsForPlotVisibility")?.Invoke(textVector, [true])) ?? throw new InvalidOperationException("Expected visibility-filtered text workbook points.")).Cast<object>().ToArray();
        TestAssert.True(visibleLabelWorkbookPoints.Length == 2, "Expected plot-visible-only projection to exclude the hidden worksheet row while preserving visible label points.");
        object visibleLabelCell = visibleLabelWorkbookPoints[1].GetType().GetProperty("WorkbookCell")?.GetValue(visibleLabelWorkbookPoints[1]) ?? throw new InvalidOperationException("Expected visible text point cell provenance.");
        TestAssert.Equal("West", ((PptxRenderer.ChartWorkbookRangeCell)visibleLabelCell).Text);
        TestAssert.Equal(2, ((PptxRenderer.ChartWorkbookRangeCell)visibleLabelCell).Index);
        var noCacheLabelPlot = ReadPointSourcePlot("""
            <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
            """);
        object noCacheTextVector = TextVector(noCacheLabelPlot);
        object[] noCacheTextActivePoints = (((System.Collections.IEnumerable?)noCacheTextVector.GetType().GetProperty("Points")?.GetValue(noCacheTextVector)) ?? throw new InvalidOperationException("Expected formula-only active text points.")).Cast<object>().ToArray();
        object[] noCacheTextWorkbookPoints = (((System.Collections.IEnumerable?)noCacheTextVector.GetType().GetProperty("WorkbookPoints")?.GetValue(noCacheTextVector)) ?? throw new InvalidOperationException("Expected formula-only text workbook sidecar points.")).Cast<object>().ToArray();
        TestAssert.True(noCacheTextActivePoints.Length == 0, "Expected formula-only workbook text values to remain inactive when chart caches are absent.");
        TestAssert.True(noCacheTextWorkbookPoints.Length == 3, "Expected formula-only workbook text values to remain available as sidecar points.");
    }

    private static PptxSceneChartPlot ReadPointSourcePlot(string channels, string kind = "barChart") =>
        TestAssert.NotNull(PptxTests.BuildSingleChartScene($$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:{{kind}}><c:ser>{{channels}}</c:ser></c:{{kind}}></c:plotArea></c:chart>
            </c:chartSpace>
            """)).Plots.Single();

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
        var sheets = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Sheet1"] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["B1"] = "Workbook Name"
            }
        };
        object workbook = PptxChartRangeTests.ReadWorkbookCells(sheets);
        System.Reflection.MethodInfo readNameRecords = typeof(PptxRenderer).GetMethod(
            "ReadSceneChartSeriesNameRecords",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected series-name record reader.");

        object[] records = (((System.Collections.IEnumerable?)readNameRecords.Invoke(null, [plot, workbook])) ?? throw new InvalidOperationException("Expected series-name records.")).Cast<object>().ToArray();

        TestAssert.True(records.Length == 1, "Expected one series-name record.");
        TestAssert.Equal("Cached Name", (string?)records[0].GetType().GetProperty("ActiveName")?.GetValue(records[0]) ?? string.Empty);
        TestAssert.Equal("Cache", records[0].GetType().GetProperty("ActiveNameSource")?.GetValue(records[0])?.ToString() ?? string.Empty);
        object[] workbookPoints = (((System.Collections.IEnumerable?)records[0].GetType().GetProperty("WorkbookPoints")?.GetValue(records[0])) ?? throw new InvalidOperationException("Expected workbook series-name sidecar points.")).Cast<object>().ToArray();
        TestAssert.True(workbookPoints.Length == 1, "Expected series-name sidecar to preserve the workbook source name.");
        TestAssert.Equal("Workbook Name", (string?)workbookPoints[0].GetType().GetProperty("Text")?.GetValue(workbookPoints[0]) ?? string.Empty);
        object workbookCell = workbookPoints[0].GetType().GetProperty("WorkbookCell")?.GetValue(workbookPoints[0]) ?? throw new InvalidOperationException("Expected series-name workbook cell provenance.");
        TestAssert.Equal("Workbook Name", ((PptxRenderer.ChartWorkbookRangeCell)workbookCell).Text);
        TestAssert.Equal(0, ((PptxRenderer.ChartWorkbookRangeCell)workbookCell).Index);
        Array typedRecords = Array.CreateInstance(records[0].GetType(), records.Length);
        typedRecords.SetValue(records[0], 0);
        System.Reflection.MethodInfo getActiveSeriesName = typeof(PptxRenderer).GetMethod(
            "GetActiveSeriesName",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected active series-name helper.");
        string activeSeriesName = (string?)getActiveSeriesName.Invoke(null, [typedRecords, 0]) ?? string.Empty;
        TestAssert.Equal("Cached Name", activeSeriesName);

        object source = records[0].GetType().GetProperty("Source")?.GetValue(records[0]) ?? throw new InvalidOperationException("Expected parsed series-name source metadata.");
        TestAssert.Equal("Sheet1!$B$1", (string?)source.GetType().GetProperty("Formula")?.GetValue(source) ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.StringReference, (PptxSceneChartDataSourceReferenceKind?)source.GetType().GetProperty("ReferenceKindValue")?.GetValue(source) ?? default);
        TestAssert.Equal("strRef", (string?)source.GetType().GetProperty("ReferenceKind")?.GetValue(source) ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceCacheKind.StringCache, (PptxSceneChartDataSourceCacheKind?)source.GetType().GetProperty("CacheKindValue")?.GetValue(source) ?? default);
        TestAssert.Equal("strCache", (string?)source.GetType().GetProperty("CacheKind")?.GetValue(source) ?? string.Empty);
        TestAssert.True((bool?)source.GetType().GetProperty("HasCachedPoints")?.GetValue(source) == true, "Expected parsed series-name cache point presence to survive scene parsing.");

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
        object[] noCacheSceneRecords = (((System.Collections.IEnumerable?)readNameRecords.Invoke(null, [noCachePlot, workbook])) ?? throw new InvalidOperationException("Expected no-cache scene series-name records.")).Cast<object>().ToArray();
        TestAssert.Equal("Workbook Name", (string?)noCacheSceneRecords[0].GetType().GetProperty("ActiveName")?.GetValue(noCacheSceneRecords[0]) ?? string.Empty);
        TestAssert.Equal("Workbook", noCacheSceneRecords[0].GetType().GetProperty("ActiveNameSource")?.GetValue(noCacheSceneRecords[0])?.ToString() ?? string.Empty);
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
            "ReadChartCategoryAxisForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected category-axis source helper.");

        object source = readCategoryAxis.Invoke(null, [chart, plot]) ?? throw new InvalidOperationException("Expected category-axis source.");
        PptxSceneChartAxis sceneAxis = (PptxSceneChartAxis?)source.GetType().GetProperty("SceneAxis")?.GetValue(source) ?? throw new InvalidOperationException("Expected scene date axis source.");
        object rebuiltSource = readCategoryAxis.Invoke(null, [PptxTests.BuildSingleChartScene(chart.ChartXml!.ToString())!, null]) ?? throw new InvalidOperationException("Expected Rebuilt-scene category-axis source.");
        PptxSceneChartAxis rebuiltAxis = (PptxSceneChartAxis)TestAssert.NotNull(rebuiltSource.GetType().GetProperty("SceneAxis")!.GetValue(rebuiltSource));

        TestAssert.Equal(PptxSceneChartAxisKind.Date, sceneAxis.AxisKind);
        TestAssert.Equal("dateAx", sceneAxis.Kind);
        TestAssert.Equal("10", sceneAxis.Id);
        TestAssert.Equal(PptxSceneChartAxisKind.Date, rebuiltAxis.AxisKind);
    }

    public static void PptxBubbleChartRendersNativeAxesGridlinesLegendAndBubbles()
    {
        string input = TestFixtures.PublicCasePath("pptx-ladder-11-chart-bubble-port.pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add, FontResolver = new TestFaceFontResolver() });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.Matches(pdf, @" c\s").Count >= 16, "Expected native bubble ellipses to be emitted as Bezier curves.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0035>"), "Expected bubble value-axis labels to include the Office 0..5 scale.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0036>"), "Expected bubble value-axis labels to include the Office 0..6 scale.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Bubble charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticBubbleChartRendersBubbleSizeDataLabels()
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
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:bubbleChart>
                  <c:dLbls><c:showVal val="0"/><c:showBubbleSize val="1"/><c:dLblPos val="t"/><c:numFmt formatCode="0"/></c:dLbls>
                  <c:ser>
                    <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
                    <c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt></c:numLit></c:xVal>
                    <c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>5</c:v></c:pt></c:numLit></c:yVal>
                    <c:bubbleSize><c:numLit><c:formatCode>0</c:formatCode><c:pt idx="0"><c:v>9</c:v></c:pt><c:pt idx="1"><c:v>16</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:bubbleSize>
                  </c:ser>
                </c:bubbleChart></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add, FontResolver = new TestFaceFontResolver() });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/CSD1", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0039>"), "Expected bubble-size data label 9 to be emitted from c:bubbleSize.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0036>"), "Expected bubble-size data label 16 to be emitted from c:bubbleSize.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Bubble charts with bubble-size labels should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticScatterChartRendersLegendKeyOnlyDataLabels()
    {
        AssertLegendKeyOnlyDataLabels("""
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
            """, new RgbColor(0x12, 0x34, 0x56), expectPlotClip: true);
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

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add, FontResolver = new TestFaceFontResolver() });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/CSD1", pdf);
        TestAssert.True(
            Regex.IsMatch(pdf, @"0\.098 0\.2 0\.302 rg\s+[0-9.]+ [0-9.]+ (?:1[5-9][0-9]|[2-9][0-9]{2})\.[0-9]{2,3} (?:[5-9][0-9]|[1-9][0-9]{2})\.[0-9]{2,3} re f"),
            "Expected chart-wide data-label manual layout to drive a large label shape box.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Scatter charts with chart-wide manual data-label layout should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticBarChartRendersLegendKeyOnlyDataLabels()
    {
        AssertLegendKeyOnlyDataLabels("""
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
            """, new RgbColor(0x65, 0x43, 0x21));
    }

    private static void AssertLegendKeyOnlyDataLabels(string chartXml, RgbColor labelColor, bool expectPlotClip = false)
    {
        // Keep the frame of the original bar and scatter probes.
        const string slide = """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                   xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:cSld><p:spTree>
                <p:graphicFrame>
                  <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm>
                  <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                </p:graphicFrame>
              </p:spTree></p:cSld>
            </p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chartXml,
            new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        var diagnostics = new DiagnosticCollector();
        PdfPage page = new PptxRenderer(fontResolver: null).RenderPagesAsync(new PptxReader().Read(package, default),
            package, diagnostics.Add, default).ToBlockingEnumerable().Single();

        var rectangles = PptxComposedPaletteTests.ReadFilledRectangles(page).ToArray();
        var labels = rectangles.Where(rect => rect.Color == labelColor).ToArray();
        TestAssert.Equal(2, labels.Length);
        RgbColor seriesColor = new(0x44, 0x72, 0xC4);
        var swatches = rectangles.Where(rect => rect.Color == seriesColor &&
            rect.Width > 0d && Math.Abs(rect.Width - rect.Height) <= .001d).ToArray();
        TestAssert.Equal(2, swatches.Length);
        var enclosed = labels.Select(label => swatches.Single(swatch => swatch.X >= label.X && swatch.Y >= label.Y &&
            swatch.X + swatch.Width <= label.X + label.Width && swatch.Y + swatch.Height <= label.Y + label.Height));
        TestAssert.Equal(2, enclosed.Distinct().Count());
        if (expectPlotClip)
            TestAssert.Contains(" re W* n", page.Content);
        TestAssert.True(diagnostics.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"),
            "Legend-key-only labels are supported for these chart families.");
        PdfContentValidator.ValidatePage(page, 0, default);
    }

    private static Dictionary<string, byte[]> ThemedBarChartPackage(string chartXml)
    {
        return new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/theme/theme1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="StyleProbe">
                  <a:themeElements>
                    <a:clrScheme name="StyleProbe">
                      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
                      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
                      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
                      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
                      <a:accent1><a:srgbClr val="156082"/></a:accent1>
                      <a:accent2><a:srgbClr val="E97132"/></a:accent2>
                      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
                      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
                      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
                      <a:accent6><a:srgbClr val="ED7D31"/></a:accent6>
                      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
                      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
                    </a:clrScheme>
                  </a:themeElements>
                </a:theme>
                """),
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
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8(chartXml)
        };
    }

    // RV04: style-18 clustered columns paint the sampled vertical gradient instead of
    // the flat theme fill (fails: flat solid, no shading resource).
    public static void BarStyle18ColumnPaintsSampledGradient()
    {
        string pdf = RenderStyle18BarChart("col", "clustered", includeSupply: false);
        TestAssert.Contains("/Sh1 sh", pdf);
        TestAssert.Contains("ShadingType 2", pdf);
        TestAssert.Contains("/Bounds [", pdf);
    }

    // RV04: the second style-18 series resolves its own theme base and gradient stops.
    public static void BarStyle18SecondSeriesPaintsOwnGradient()
    {
        string pdf = RenderStyle18BarChart("col", "clustered", includeSupply: true);
        TestAssert.Contains("/Sh4 sh", pdf);
    }
    // Office uses the raw theme base and a 5pt stroke for style-18 series without markers.
    public static void LineStyle18UsesRawFivePointSeriesStroke()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.082 0.376 0.51 RG", pdf);
        TestAssert.DoesNotContain("0.078 0.369 0.498 RG", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"^5 w\r?$", RegexOptions.Multiline), "Expected a 5pt style-18 series stroke.");
        TestAssert.DoesNotContain("2.25 w", pdf);
    }

    // RV04: an explicit series line width wins over the style-18 5pt default.
    public static void LineStyle18KeepsExplicitSeriesStrokeWidth()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"^2 w\r?$", RegexOptions.Multiline), "Expected the explicit 2pt series stroke to survive style 18.");
        TestAssert.True(!Regex.IsMatch(pdf, @"^5 w\r?$", RegexOptions.Multiline), "The explicit width must win over the style-18 5pt default.");
    }
    // Office paints style-18 automatic gallery markers with 12.96pt geometry and axial gradient fills.
    public static void LineStyle18DrawsGradientGalleryMarkers()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("12.96 12.96 re", pdf);
        TestAssert.Equal(4, Regex.Matches(pdf, "/Sh\\d+ sh").Count);
        TestAssert.Contains("ShadingType 2", pdf);
        TestAssert.DoesNotContain("0.082 0.376 0.51 rg", pdf);
        TestAssert.DoesNotContain("0.914 0.443 0.196 rg", pdf);
    }
    // RV04: style-18 line series draw smoothed beziers instead of straight segments
    // (Office emits curve operators on the style-18 deck while default-style lines
    // stay straight; markers stay off so the polyline is the only curved path).
    public static void LineStyle18SmoothsSeriesCurves()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @" c\r?$", RegexOptions.Multiline), "Expected smoothed style-18 series curves.");
    }

    // RV04: an explicit smooth=false keeps straight segments under style 18.
    public static void LineStyle18KeepsExplicitStraightSegments()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>
                <c:smooth val="0"/>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(!Regex.IsMatch(pdf, @" c\r?$", RegexOptions.Multiline), "Explicit smooth=false must keep straight segments under style 18.");
    }
    // RV04: style-26 unstyled line series stroke the raw base at 7pt with smoothing
    // (Office emits raw strokes at 7 w with beziers and no markers on the style-26
    // deck; markers stay off so the polyline is the only wide curved operator).
    public static void LineStyle26StrokesRawSevenPointCurves()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="26"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.082 0.376 0.51 RG", pdf);
        TestAssert.DoesNotContain("0.078 0.369 0.498 RG", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"^7 w\r?$", RegexOptions.Multiline), "Expected a 7pt style-26 series stroke.");
        TestAssert.True(Regex.IsMatch(pdf, @" c\r?$", RegexOptions.Multiline), "Expected smoothed style-26 series curves.");
    }
    // RV04: series-level symbol=none suppresses forced style-18 markers for that
    // series only (Office draws no series-0 markers on the symbol=none probe while
    // series 1 keeps its forced markers; plot-level marker markup never suppresses
    // gallery forcing).
    public static void LineStyle18ExplicitSeriesNoneDrawsNoMarkers()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:marker><c:symbol val="none"/></c:marker>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(1, Regex.Matches(pdf, @"0\.082 0\.376 0\.51 RG").Count);
        TestAssert.Equal(3, Regex.Matches(pdf, @"0\.914 0\.443 0\.196 RG").Count);
    }
    // RV04: style-26 draws no gallery markers even with plot-level marker=1
    // (Office renders no markers on the marker=1 style-26 probe; only the raw
    // polyline emits series color).
    public static void LineStyle26PlotMarkerOneDrawsNoMarkers()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="26"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="1"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(1, Regex.Matches(pdf, @"0\.082 0\.376 0\.51 RG").Count);
        TestAssert.DoesNotContain("0.078 0.369 0.498 RG", pdf);
    }
    // RV04: style-18 markers with plot-level marker=1 but no series markup also draw
    // at the 12.96pt gallery size (Office draws 12.96 on the marker=1 style-18 probe;
    // the square series pins the geometry through its outline rectangle).
    public static void LineStyle18MarkerOneUsesGallerySize()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="1"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("12.96 12.96 re", pdf);
    }
    // RV04: style-26 ignores even series-explicit markers (Office draws no markers
    // on the explicit-diamond style-26 probe; only the raw polyline emits color).
    public static void LineStyle26ExplicitMarkersDrawNothing()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="26"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:marker><c:symbol val="diamond"/></c:marker>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(1, Regex.Matches(pdf, @"0\.082 0\.376 0\.51 RG").Count);
        TestAssert.DoesNotContain(" rg", pdf);
    }

    // RV04: style-18 gradient-fills explicit circle markers at their explicit size
    // (Office gradient-fills the size-9 circles on the explicit-marker probe).
    public static void LineStyle18ExplicitCirclePaintsGradient()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:marker><c:symbol val="circle"/><c:size val="9"/></c:marker>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(2, Regex.Matches(pdf, @"/Sh\d+ sh").Count);
        TestAssert.Contains("ShadingType 2", pdf);
        TestAssert.DoesNotContain("0.082 0.376 0.51 rg", pdf);
    }
    // RV04: explicit smooth series render the smoothed appearance package (Office
    // draws raw strokes at 3pt with smoothed beziers and forced flat markers on
    // the default-style smooth probe; this pins the raw/width operators while the
    // marker test below pins the forced-marker geometry).
    public static void LineSmoothSeriesStrokesRawThreePoint()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>
                <c:smooth val="1"/>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.082 0.376 0.51 RG", pdf);
        TestAssert.DoesNotContain("0.078 0.369 0.498 RG", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"^3 w\r?$", RegexOptions.Multiline), "Expected a 3pt smoothed series stroke.");
    }

    // RV04: explicit smooth series force flat 9pt markers past marker=0 (Office
    // draws 9pt flat markers with raw rims on the marker=0 smooth probe; the square
    // series pins the 9pt geometry through its rectangle operator).
    public static void LineSmoothSeriesForcesNinePointMarkers()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:smooth val="1"/>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
                <c:smooth val="1"/>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("9 9 re", pdf);
        TestAssert.DoesNotContain("0.078 0.369 0.498 RG", pdf);
    }

    public static void LinePlotSmoothedAppliesPackagePlotWide() =>
        AssertLinePlotSmoothing();

    public static void LinePlotSmoothedKeepsExplicitStraightSeries() =>
        AssertLinePlotSmoothing(straightSeries: true);

    private static void AssertLinePlotSmoothing(bool straightSeries = false, bool straightPlot = false)
    {
        const string disabledSmoothing = """<c:smooth val="0"/>""";
        string seriesSmooth = straightSeries ? disabledSmoothing : string.Empty;
        string plotSmooth = straightPlot ? disabledSmoothing : string.Empty;
        using var input = TestFixtures.CreateZipPackage(ThemedBarChartPackage($$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:smooth val="1"/>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>
                {{seriesSmooth}}
              </c:ser>
              {{plotSmooth}}
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="10"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        using var package = OoxPackage.Open(input, default);
        var document = new PptxReader().Read(package, default);
        string paint = new PptxRenderer(new TestFaceFontResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single().Content;

        // The explicit 0–10 axis scale fixes these PDF endpoints at the three category centers.
        AssertSeries("0.082 0.376 0.51", smooth: true, diamond: true, heights: ["272.264", "410.816", "318.448"]);
        AssertSeries(straightSeries ? "0.91 0.427 0.173" : "0.914 0.443 0.196",
            smooth: !straightSeries, diamond: false, heights: ["249.172", "341.54", "295.356"]);
        TestAssert.DoesNotContain("0.078 0.369 0.498 RG", paint);
        if (!straightSeries) TestAssert.DoesNotContain("0.91 0.427 0.173 RG", paint);

        void AssertSeries(string rgb, bool smooth, bool diamond, string[] heights)
        {
            const string coordinate = @"[-0-9.]+";
            string color = Regex.Escape(rgb);
            string width = smooth ? "3" : @"2\.25";
            var paths = Regex.Matches(paint, $@"(?m)^{color} RG\s+{width} w\s+\[\] 0 d\s+1 J\s+1 j\s+q\s+"
                + $@"83 226\.08 410 230\.92 re W\* n\s+(?<path>{coordinate} {coordinate} m\s+(?:(?:{coordinate} )+[cl]\s+){{2}})S\s+Q");
            TestAssert.True(paths.Count == 1, $"Series {rgb} must paint one path in its plot clip.");
            string[] commands = paths[0].Groups["path"].Value.Trim().Split('\n', StringSplitOptions.TrimEntries);
            string[] positions = ["151.333", "288", "424.667"];
            TestAssert.Equal($"{positions[0]} {heights[0]} m", commands[0]);
            for (int index = 1; index < positions.Length; index++)
            {
                string point = Regex.Escape($"{positions[index]} {heights[index]}");
                string segment = smooth ? $@"^(?:{coordinate} ){{4}}{point} c$" : $@"^{point} l$";
                TestAssert.True(Regex.IsMatch(commands[index], segment),
                    $"Series {rgb} must reach point {index + 1} with the authored smoothing precedence.");
            }

            string markerWidth = smooth ? "1" : @"0\.75";
            string markerSize = smooth ? "9" : "7";
            string marker = diamond ? $@"(?:{coordinate} {coordinate} [ml]\s+){{4}}h\s+S"
                : $@"{coordinate} {coordinate} {markerSize} {markerSize} re S";
            TestAssert.Equal(3, Regex.Matches(paint,
                $@"(?m)^{color} RG\s+{markerWidth} w\s+\[\] 0 d\s+0 J\s+1 j\s+{marker}\s+Q").Count);
        }
    }

    // RV04: style-18 legend keys match the series appearance (Office draws raw 5pt
    // swatches with gradient gallery markers; the two plot series pin four marker
    // shadings and the two legend keys pin two more).
    public static void LineStyle18LegendKeysMatchSeriesAppearance()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(6, Regex.Matches(pdf, @"/Sh\d+ sh").Count);
        TestAssert.DoesNotContain("0.078 0.369 0.498 RG", pdf);
        TestAssert.DoesNotContain("0.91 0.427 0.173 RG", pdf);
    }

    // RV04: explicit marker outlines win under the smooth package and in gallery
    // legend keys (Office strokes red rims with default fills; characterization pins of
    // already-correct behavior).
    public static void LineSmoothExplicitMarkerOutlineStrokesRed()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:marker><c:symbol val="circle"/><c:size val="9"/><c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:marker>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:smooth val="1"/>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        MatchCollection rims = Regex.Matches(pdf, @"1 0 0 RG\r?\n([\d.]+) w", RegexOptions.Multiline);
        TestAssert.Equal(2, rims.Count);
        foreach (Match rim in rims)
        {
            TestAssert.Equal("1", rim.Groups[1].Value);
        }
    }

    public static void LineStyle26LegendExplicitMarkerOutlineStrokesRed() =>
        AssertStyle26LegendExplicitMarkerPaint(stroke: true);

    public static void LineStyle26LegendExplicitMarkerFillPaintsFlat() =>
        AssertStyle26LegendExplicitMarkerPaint(stroke: false);

    private static void AssertStyle26LegendExplicitMarkerPaint(bool stroke)
    {
        const string red = """<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>""";
        string markerPaint = stroke ? "<a:ln>" + red + "</a:ln>" : red;
        var parts = ThemedBarChartPackage($$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="26"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:marker><c:symbol val="circle"/><c:size val="9"/><c:spPr>{{markerPaint}}</c:spPr></c:marker>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend></c:chart></c:chartSpace>
            """);
        const string slidePart = "ppt/slides/slide1.xml";
        // The unrelated red frame makes page-wide color checks insufficient.
        const string frame = """
            <p:sp><p:spPr>
              <a:xfrm><a:off x="127000" y="127000"/><a:ext cx="381000" cy="254000"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
              <a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
            </p:spPr></p:sp>
            """;
        parts[slidePart] = TestFixtures.Utf8(Encoding.UTF8.GetString(parts[slidePart])
            .Replace("</p:spTree>", frame + "</p:spTree>", StringComparison.Ordinal));
        using var stream = TestFixtures.CreateZipPackage(parts);
        using var package = OoxPackage.Open(stream, default);
        var renderer = new PptxRenderer(new PortableRenderingTests.PinnedResolver());
        PdfPage page = renderer.RenderPagesAsync(new PptxReader().Read(package, default), package, null, default)
            .ToBlockingEnumerable().Single();

        var demand = PortableRenderingTests.ReadPaintedSourceLabel(page, "Demand", fontPrefix: "CL");
        var supply = PortableRenderingTests.ReadPaintedSourceLabel(page, "Supply", fontPrefix: "CL");
        string content = page.Content.ReplaceLineEndings("\n");
        const string number = @"[-\d.]+";
        string pair = number + " " + number;
        string circle = $@"(?<x>{number}) (?<y>{number}) m\n(?:(?:{pair} ){{3}}c\n){{4}}h\n";
        string color = stroke ? @"RG\n(?:[^\n]+ [wdJj]\n)*" : @"rg\n";
        string operation = stroke ? @"S" : @"f";
        MatchCollection painted = Regex.Matches(content, $@"(?m)^1 0 0 {color}{circle}{operation}$");
        TestAssert.Equal(1, painted.Count);
        double x = double.Parse(painted[0].Groups["x"].Value, CultureInfo.InvariantCulture);
        double y = double.Parse(painted[0].Groups["y"].Value, CultureInfo.InvariantCulture);
        TestAssert.True(demand[0].Origin.X - x is > 0d and < 9d &&
            y > demand[0].Origin.Y && y < demand[0].Origin.Y + demand[0].Ink.Height &&
            y > supply[0].Origin.Y + supply[0].Ink.Height,
            "The red circle paints immediately beside Demand on its own legend row.");

        var paths = new List<string>();
        var paints = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add, strokes: true).ToArray();
        var supplyKey = paints.Zip(paths).Single(paint => paint.First == new RgbColor(232, 109, 44) &&
            paint.Second.EndsWith(" re S", StringComparison.Ordinal)).Second.Split(' ');
        double right = double.Parse(supplyKey[0], CultureInfo.InvariantCulture) + double.Parse(supplyKey[2], CultureInfo.InvariantCulture);
        double bottom = double.Parse(supplyKey[1], CultureInfo.InvariantCulture);
        double top = bottom + double.Parse(supplyKey[3], CultureInfo.InvariantCulture);
        TestAssert.True(supply[0].Origin.X - right is > 0d and < 9d &&
            bottom < supply[0].Origin.Y && top > supply[0].Origin.Y &&
            top < demand[0].Origin.Y,
            "The independently styled Supply key stays on its own row.");
    }

    // RV04: explicit plus/square markers without a size element draw at the nominal
    // size-5 geometry like dots and stars (Office paints no-size plus diameters and
    // square boxes at 5.04; the 0.04 residual is the shared geometry-rule domain).
    public static void LineExplicitPlusSquareWithoutSizeDrawNominalFive()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:marker><c:symbol val="plus"/></c:marker>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:marker><c:symbol val="square"/></c:marker>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("5 5 re", pdf);
        bool hasFivePointVertical = false;
        foreach (Match m in Regex.Matches(pdf, @"^([\d.]+) ([\d.]+) m \1 ([\d.]+) l S\r?$", RegexOptions.Multiline))
        {
            double y1 = double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            double y2 = double.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
            if (Math.Abs(Math.Abs(y2 - y1) - 5d) < 0.001d)
            {
                hasFivePointVertical = true;
            }
        }
        TestAssert.True(hasFivePointVertical, "Expected a 5pt vertical plus diameter.");
    }
    public static void LinePlotSmoothZeroYieldsToExplicitSeriesSmooth() =>
        AssertLinePlotSmoothing(straightPlot: true);

    // RV04: default-style legend keys keep the legacy tinted swatches.
    public static void DefaultLegendKeysKeepTintedSwatches()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="1"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.078 0.369 0.498 RG", pdf);
    }
    // RV04: post-gap fragment starts ease in with a halved control offset (Office
    // halves the gap-adjacent control while the chart-end control keeps full thirds;
    // markers stay off so the fragment curve is the only bezier).
    public static void LineStyle18GapFragmentEasesPostGapStart()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        MatchCollection curves = Regex.Matches(pdf, @"^([\d.]+) ([\d.]+) m\r?$\s*^([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) c\r?$", RegexOptions.Multiline);
        TestAssert.Equal(1, curves.Count);
        double startX = double.Parse(curves[0].Groups[1].Value, CultureInfo.InvariantCulture);
        double firstControlX = double.Parse(curves[0].Groups[3].Value, CultureInfo.InvariantCulture);
        double secondControlX = double.Parse(curves[0].Groups[5].Value, CultureInfo.InvariantCulture);
        double endX = double.Parse(curves[0].Groups[7].Value, CultureInfo.InvariantCulture);
        TestAssert.True(Math.Abs((firstControlX - startX) / (endX - startX) - 1d / 6d) < 0.02, "Expected the post-gap start control at one sixth of the segment.");
        TestAssert.True(Math.Abs((endX - secondControlX) / (endX - startX) - 1d / 3d) < 0.02, "Expected the chart-end control at one third of the segment.");
    }
    // RV04: span-mode smoothing uses category-based offsets (Office spans the hole
    // in one subpath: the chart-edge start eases at one sixth of its double-width
    // segment while span-adjacent interior controls sit at half a category width).
    public static void LineStyle18SpanFragmentUsesCategoryOffsets()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:dispBlanksAs val="span"/></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        MatchCollection paths = Regex.Matches(pdf, @"^([\d.]+) ([\d.]+) m\r?$\s*^([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) c\r?$\s*^([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) c\r?$", RegexOptions.Multiline);
        TestAssert.Equal(1, paths.Count);
        double spanStartX = double.Parse(paths[0].Groups[1].Value, CultureInfo.InvariantCulture);
        double spanFirstX = double.Parse(paths[0].Groups[3].Value, CultureInfo.InvariantCulture);
        double spanJointSecondX = double.Parse(paths[0].Groups[5].Value, CultureInfo.InvariantCulture);
        double spanJointX = double.Parse(paths[0].Groups[7].Value, CultureInfo.InvariantCulture);
        double trailFirstX = double.Parse(paths[0].Groups[9].Value, CultureInfo.InvariantCulture);
        double trailSecondX = double.Parse(paths[0].Groups[11].Value, CultureInfo.InvariantCulture);
        double trailEndX = double.Parse(paths[0].Groups[13].Value, CultureInfo.InvariantCulture);
        double spanDx = spanJointX - spanStartX;
        double trailDx = trailEndX - spanJointX;
        TestAssert.True(Math.Abs((spanFirstX - spanStartX) / spanDx - 1d / 6d) < 0.02, "Expected the span start control at one sixth of the double-width segment.");
        TestAssert.True(Math.Abs((spanJointX - spanJointSecondX) / spanDx - 1d / 4d) < 0.02, "Expected the span joint control at one quarter of the double-width segment.");
        TestAssert.True(Math.Abs((trailFirstX - spanJointX) / trailDx - 1d / 2d) < 0.02, "Expected the trailing joint control at half the single-width segment.");
        TestAssert.True(Math.Abs((trailEndX - trailSecondX) / trailDx - 1d / 3d) < 0.02, "Expected the chart-end control at one third of the segment.");
    }
    // RV04: explicit series lines without a width inherit the gallery width (Office
    // keeps 7pt under style-26 when only a color is explicit; an explicit width wins).
    public static void LineStyle26ExplicitLineWithoutWidthKeepsGalleryWidth()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="26"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Equal(2, Regex.Matches(pdf, @"^7 w\r?$", RegexOptions.Multiline).Count);
    }

    // RV04: explicit marker fills win under the smooth package (Office paints red
    // flat circles with no gradient shading for a red-filled marker on a smoothed
    // default-style series; characterization pin of already-correct behavior).
    public static void LineSmoothExplicitMarkerFillPaintsFlat()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:marker><c:symbol val="circle"/><c:size val="9"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:marker>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:smooth val="1"/>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
    }

    // RV04: zero-mode blanks plot as zero with contiguous full-thirds smoothing (Office
    // draws the hole slot at zero in one subpath; zero-filled points are real plotted
    // points, so no control halves apply; characterization pin of confirmed behavior).
    public static void LineStyle18ZeroFragmentSmoothsThroughZero()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt><c:pt idx="3"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:dispBlanksAs val="zero"/></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        MatchCollection paths = Regex.Matches(pdf, @"^([\d.]+) ([\d.]+) m\r?$\s*^([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) c\r?$\s*^([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) c\r?$\s*^([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) c\r?$", RegexOptions.Multiline);
        TestAssert.Equal(1, paths.Count);
        double x0 = double.Parse(paths[0].Groups[1].Value, CultureInfo.InvariantCulture);
        double c1 = double.Parse(paths[0].Groups[3].Value, CultureInfo.InvariantCulture);
        double c2 = double.Parse(paths[0].Groups[5].Value, CultureInfo.InvariantCulture);
        double x1 = double.Parse(paths[0].Groups[7].Value, CultureInfo.InvariantCulture);
        double c3 = double.Parse(paths[0].Groups[9].Value, CultureInfo.InvariantCulture);
        double c4 = double.Parse(paths[0].Groups[11].Value, CultureInfo.InvariantCulture);
        double x2 = double.Parse(paths[0].Groups[13].Value, CultureInfo.InvariantCulture);
        double c5 = double.Parse(paths[0].Groups[15].Value, CultureInfo.InvariantCulture);
        double c6 = double.Parse(paths[0].Groups[17].Value, CultureInfo.InvariantCulture);
        double x3 = double.Parse(paths[0].Groups[19].Value, CultureInfo.InvariantCulture);
                TestAssert.True(Math.Abs((c1 - x0) / (x1 - x0) - 1d / 3d) < 0.02, "seg1 start third");
        TestAssert.True(Math.Abs((x1 - c2) / (x1 - x0) - 1d / 3d) < 0.02, "seg1 end third");
        TestAssert.True(Math.Abs((c3 - x1) / (x2 - x1) - 1d / 3d) < 0.02, "seg2 start third");
        TestAssert.True(Math.Abs((x2 - c4) / (x2 - x1) - 1d / 3d) < 0.02, "seg2 end third");
        TestAssert.True(Math.Abs((c5 - x2) / (x3 - x2) - 1d / 3d) < 0.02, "seg3 start third");
        TestAssert.True(Math.Abs((x3 - c6) / (x3 - x2) - 1d / 3d) < 0.02, "seg3 end third");
    }

    // RV04: consecutive-hole spans ease with mean-thirds joints (Office spans
    // two holes in ratios one sixth, two ninths, two thirds with full-thirds tails;
    // the single rendered subpath is C1-identical to the Office hole-per-subpath split).
    public static void LineStyle18MultiHoleSpanUsesMeanThirds()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:marker val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt><c:pt idx="4"><c:v>E</c:v></c:pt><c:pt idx="5"><c:v>F</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>4</c:v></c:pt><c:pt idx="4"><c:v>3</c:v></c:pt><c:pt idx="5"><c:v>5</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:dispBlanksAs val="span"/></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        MatchCollection paths = Regex.Matches(pdf, @"^([\d.]+) ([\d.]+) m\r?$\s*^([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) c\r?$\s*^([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) c\r?$\s*^([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) c\r?$", RegexOptions.Multiline);
        TestAssert.Equal(1, paths.Count);
        double sx = double.Parse(paths[0].Groups[1].Value, CultureInfo.InvariantCulture);
        double s1 = double.Parse(paths[0].Groups[3].Value, CultureInfo.InvariantCulture);
        double e1 = double.Parse(paths[0].Groups[5].Value, CultureInfo.InvariantCulture);
        double jx = double.Parse(paths[0].Groups[7].Value, CultureInfo.InvariantCulture);
        double s2 = double.Parse(paths[0].Groups[9].Value, CultureInfo.InvariantCulture);
        double tx = double.Parse(paths[0].Groups[13].Value, CultureInfo.InvariantCulture);
        double dx1 = jx - sx;
        double dx2 = tx - jx;
        TestAssert.True(Math.Abs((s1 - sx) / dx1 - 1d / 6d) < 0.02, "double-span start sixth");
        TestAssert.True(Math.Abs((jx - e1) / dx1 - 2d / 9d) < 0.02, "double-span joint two ninths");
        TestAssert.True(Math.Abs((s2 - jx) / dx2 - 2d / 3d) < 0.02, "trailing joint two thirds");
    }
    // Office paints explicit style-18 dots as short gradient-filled rectangles.
    public static void LineStyle18PaintsGradientFilledDotRectangle()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:marker><c:symbol val="dot"/><c:size val="9"/></c:marker>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(2, Regex.Matches(pdf, @"/Sh\d+ sh").Count);
        TestAssert.Contains("ShadingType 2", pdf);
        TestAssert.DoesNotContain("0.082 0.376 0.51 rg", pdf);
        TestAssert.Contains("4.44 1.8 re", pdf);
    }
    // RV04: style-18 legend circles gradient-fill like plot circles (Office
    // gradient-fills the explicit size-9 legend circle key; one series pins two
    // plot shadings plus one legend shading).
    public static void LineStyle18LegendCirclePaintsGradient()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:marker><c:symbol val="circle"/><c:size val="9"/></c:marker>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(3, Regex.Matches(pdf, @"/Sh\d+ sh").Count);
        TestAssert.Contains("ShadingType 2", pdf);
        TestAssert.DoesNotContain("0.082 0.376 0.51 rg", pdf);
    }
    // RV04: plot-level smooth applies the smoothed package plot-wide (Office smooths
    // every series on the plot-smooth-only probe; the two series pin raw polylines,
    // 3pt widths, curves and forced markers with no series smooth markup).
    public static void LinePlotLevelSmoothAppliesPackagePlotWide()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Supply</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:smooth val="1"/>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(6, Regex.Matches(pdf, @"0\.082 0\.376 0\.51 RG|0\.914 0\.443 0\.196 RG").Count);
        TestAssert.DoesNotContain("0.078 0.369 0.498 RG", pdf);
        TestAssert.DoesNotContain("0.91 0.427 0.173 RG", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"^3 w\r?$", RegexOptions.Multiline), "Expected 3pt plot-smoothed series strokes.");
        TestAssert.True(Regex.IsMatch(pdf, @" c\r?$", RegexOptions.Multiline), "Expected plot-smoothed series curves.");
    }
    // RV04: style-18 stars stroke a 3-diameter asterisk with no fill (Office strokes
    // vertical plus two diagonal diameters with a meaningless gradient fill; the
    // size-9 vertical diameter pins the asterisk geometry).
    public static void LineStyle18ExplicitStarStrokesAsterisk()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:marker><c:symbol val="star"/><c:size val="9"/></c:marker>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.DoesNotContain("0.082 0.376 0.51 rg", pdf);
        bool hasNinePointVertical = false;
        foreach (Match m in Regex.Matches(pdf, @"^([\d.]+) ([\d.]+) m \1 ([\d.]+) l S\r?$", RegexOptions.Multiline))
        {
            double y1 = double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            double y2 = double.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
            if (Math.Abs(Math.Abs(y2 - y1) - 9d) < 0.001d)
            {
                hasNinePointVertical = true;
            }
        }
        TestAssert.True(hasNinePointVertical, "Expected a 9pt vertical asterisk diameter.");
    }

    // RV04: style-18 dashes render as filled gradient bars (Office paints size-wide
    // bars with dot-rule heights through the shared gradient table, not stroked lines).
    public static void LineStyle18ExplicitDashPaintsGradientBar()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:marker><c:symbol val="dash"/><c:size val="9"/></c:marker>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("9 1.8 re", pdf);
        TestAssert.Equal(2, Regex.Matches(pdf, @"/Sh\d+ sh").Count);
    }

    // RV04: default-style line charts without marker markup still draw no markers.
    public static void DefaultStyleDrawsNoMarkersWithoutMarkup()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.DoesNotContain("12.96 12.96 re", pdf);
    }

    // RV04: explicit dot/star markers without a size element draw at the nominal
    // size-5 geometry (Office paints no-size dots as 2.52 by 0.96 rects and no-size
    // star diameters at 5.04, matching explicit size 5 digit-for-digit).
    public static void LineStyle18MissingDotStarSizesDrawNominalFive()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:lineChart>
              <c:grouping val="standard"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Dots</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                <c:marker><c:symbol val="dot"/></c:marker>
              </c:ser>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Stars</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numLit></c:val>
                <c:marker><c:symbol val="star"/></c:marker>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("2.52 0.96 re", pdf);
        bool hasFivePointVertical = false;
        foreach (Match m in Regex.Matches(pdf, @"^([\d.]+) ([\d.]+) m \1 ([\d.]+) l S\r?$", RegexOptions.Multiline))
        {
            double y1 = double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            double y2 = double.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture);
            if (Math.Abs(Math.Abs(y2 - y1) - 5d) < 0.001d)
            {
                hasFivePointVertical = true;
            }
        }
        TestAssert.True(hasFivePointVertical, "Expected a 5pt vertical asterisk diameter.");
    }

    // RV04: style-18 columns paint the derived Office gradient when the theme base has
    // no sampled recipe (default theme accent resolves outside the measured table).
    public static void BarStyle18DerivedStopsPaintGradientForUnmeasuredBase()
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
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:barChart>
                  <c:barDir val="col"/>
                  <c:grouping val="clustered"/>
                  <c:varyColors val="0"/>
                  <c:ser>
                    <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"/Sh\d+ sh"), "Derived stops must shade unmeasured theme bases.");
        TestAssert.Contains("ShadingType 2", pdf);
        TestAssert.DoesNotContain("0.267 0.447 0.769 rg", pdf);
    }

    // RV04: derived style-18/26 mid stops are byte-exact on the 37-face Office corpus.
    public static void BarStyle18DerivedMidStopsMatchOfficeExactly()
    {
        foreach (string row in StyleBarStopCorpus)
        {
            string[] parts = row.Split((char)58);
            TestAssert.True(RgbColor.TryParse(parts[0], out RgbColor baseColor), row);
            TestAssert.True(RgbColor.TryParse(parts[2], out RgbColor expected), row);
            (_, RgbColor middle, _) = PptxRenderer.DeriveStyleBarGradientStops(baseColor);
            TestAssert.Equal(expected, middle);
        }
    }

    // RV04: derived dark stops stay within 1 level of Office on the 37-face corpus.
    public static void BarStyle18DerivedDarkStopsStayWithinOneLevelOfOffice()
    {
        foreach (string row in StyleBarStopCorpus)
        {
            string[] parts = row.Split((char)58);
            TestAssert.True(RgbColor.TryParse(parts[0], out RgbColor baseColor), row);
            TestAssert.True(RgbColor.TryParse(parts[1], out RgbColor expected), row);
            (RgbColor dark, _, _) = PptxRenderer.DeriveStyleBarGradientStops(baseColor);
            TestAssert.True(System.Math.Abs(dark.Red - expected.Red) <= 1 && System.Math.Abs(dark.Green - expected.Green) <= 1 && System.Math.Abs(dark.Blue - expected.Blue) <= 1, row);
        }
    }

    // RV04: derived light stops stay within 2 levels of Office on the 37-face corpus.
    public static void BarStyle18DerivedLightStopsStayWithinTwoLevelsOfOffice()
    {
        foreach (string row in StyleBarStopCorpus)
        {
            string[] parts = row.Split((char)58);
            TestAssert.True(RgbColor.TryParse(parts[0], out RgbColor baseColor), row);
            TestAssert.True(RgbColor.TryParse(parts[3], out RgbColor expected), row);
            (_, _, RgbColor light) = PptxRenderer.DeriveStyleBarGradientStops(baseColor);
            TestAssert.True(System.Math.Abs(light.Red - expected.Red) <= 2 && System.Math.Abs(light.Green - expected.Green) <= 2 && System.Math.Abs(light.Blue - expected.Blue) <= 2, row);
        }
    }

    private static readonly string[] StyleBarStopCorpus =
    [
        "156082:08587C:106287:497491",
        "E97132:E15E19:F26E29:ED8256",
        "C00000:BB0000:CA0000:CB4545",
        "7030A0:641F97:712AA6:8153AC",
        "0E2841:07223D:0B2844:48505E",
        "FFC000:E5B600:FFC600:FFC746",
        "000000:000000:000000:454545",
        "000040:00003D:000043:45455E",
        "000080:00007C:000086:454590",
        "0000FF:0000E5:0000FF:4646FF",
        "004000:003D00:004300:455E45",
        "008000:007C00:008600:459045",
        "00C000:00BB00:00CA00:45CB45",
        "00FF00:00E500:00FF00:46FF46",
        "00FFFF:00E5E5:00FFFF:46FFFF",
        "0E2850:06214C:0B2753:475069",
        "101010:0D0D0D:101010:484848",
        "142850:0B214B:112753:495069",
        "202020:1B1B1B:202020:4D4D4D",
        "400000:3D0000:430000:5E4545",
        "404040:383838:404040:5D5D5D",
        "606060:545454:606060:747474",
        "800000:7C0000:860000:904545",
        "808000:7C7C00:868600:909045",
        "808080:717171:808080:8E8E8E",
        "A0A0A0:8D8D8D:A0A0A0:ABABAB",
        "C0C0C0:AAAAAA:C0C0C0:C8C8C8",
        "E0E0E0:C7C7C7:E0E0E0:E6E6E6",
        "F0F0F0:D5D5D5:F0F0F0:F5F5F5",
        "FF0000:E50000:FF0000:FF4646",
        "FF00FF:E500E5:FF00FF:FF46FF",
        "FF4000:E52C00:FF3A00:FF5E46",
        "FF8000:E57100:FF8000:FF8E46",
        "FFC800:E5BE00:FFCF00:FFCF46",
        "FFE000:E5D800:FFEA00:FFE546",
        "FFFF00:E5E500:FFFF00:FFFF46",
        "FFFFFF:E2E2E2:FFFFFF:FFFFFF",
    ];
    // RV04: explicit series fills win over the style gradient even when the explicit
    // color equals a sampled base.
    public static void BarStyle18ExplicitFillKeepsFlatFill()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", ThemedBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="18"/><c:chart><c:plotArea><c:barChart>
              <c:barDir val="col"/>
              <c:grouping val="clustered"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:spPr><a:solidFill><a:srgbClr val="156082"/></a:solidFill></c:spPr>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(!Regex.IsMatch(pdf, @"/Sh\d+ sh"), "Explicit fills must not gain style shading resources.");
        TestAssert.Contains("0.082 0.376 0.51 rg", pdf);
    }
    private static Dictionary<string, byte[]> PlainBarChartPackage(string chartXml)
    {
        return new Dictionary<string, byte[]>
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
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8(chartXml)
        };
    }

    // RV04: style-10 clustered columns stroke white 1pt round-join bar outlines
    // (fails: no white stroke on the bar rects).
    public static void BarStyle10ColumnStrokesWhiteOutline()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", PlainBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="10"/><c:chart><c:plotArea><c:barChart>
              <c:barDir val="col"/>
              <c:grouping val="clustered"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 G", pdf);
        TestAssert.True(Regex.Matches(pdf, @" re S").Count >= 2, "Both style-10 bars must carry a stroked outline rect.");
    }

    // RV04: other styles keep bar outlines off (style-2 paints no white stroke).
    public static void BarStyle2ColumnKeepsOutlineOff()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", PlainBarChartPackage("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:style val="2"/><c:chart><c:plotArea><c:barChart>
              <c:barDir val="col"/>
              <c:grouping val="clustered"/>
              <c:varyColors val="0"/>
              <c:ser>
                <c:tx><c:strLit><c:pt idx="0"><c:v>Demand</c:v></c:pt></c:strLit></c:tx>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
              </c:ser>
              <c:axId val="10"/><c:axId val="20"/>
            </c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
            """));
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.DoesNotContain("1 G", pdf);
    }
    // RV04: style-18 horizontal bars paint the same vertical base-relative gradient
    // (fails: flat solid, no shading resource).
    public static void BarStyle18HorizontalBarPaintsGradient()
    {
        PdfRectangle[] bars = ReadStyle18BarGradientClips(
            RenderStyle18BarChart("bar", "clustered", includeSupply: false)).OrderBy(clip => clip.Y).ToArray();
        TestAssert.Equal(2, bars.Length);
        TestAssert.True(bars[0].X == bars[1].X && bars[0].Height == bars[1].Height &&
            bars[0].Y + bars[0].Height < bars[1].Y,
            "Each category must paint a separate horizontal bar on the common value axis.");
        TestAssert.True(Math.Abs(2d * bars[0].Width - bars[1].Width) < 0.005d,
            "The gradient clips must cover the authored values 2 and 4.");
    }

    // RV04: style-18 stacked columns paint each segment with its own bar-relative
    // gradient instead of one flat compound path (fails: flat solid, no shading).
    public static void BarStyle18StackedColumnPaintsPerSegmentGradient()
    {
        PdfRectangle[] clips = ReadStyle18BarGradientClips(
            RenderStyle18BarChart("col", "stacked", includeSupply: true));
        TestAssert.Equal(4, clips.Length);
        PdfRectangle[][] columns = clips.GroupBy(clip => clip.X).OrderBy(column => column.Key)
            .Select(column => column.OrderBy(clip => clip.Y).ToArray()).ToArray();
        TestAssert.Equal(2, columns.Length);
        TestAssert.Equal(2, columns[0].Length);
        TestAssert.Equal(2, columns[1].Length);
        PdfRectangle baseline = columns[0][0];
        double valueUnit = baseline.Height / 2d;
        TestAssert.True(baseline.X + baseline.Width < columns[1][0].X,
            "The two stacked categories must paint separate columns.");
        foreach (var (column, lowerValue, upperValue) in new[]
        {
            (columns[0], 2d, 1d),
            (columns[1], 4d, 3d)
        })
        {
            PdfRectangle lower = column[0], upper = column[1];
            TestAssert.True(lower.Y == baseline.Y && lower.Width == baseline.Width && upper.Width == lower.Width &&
                Math.Abs(lower.Y + lower.Height - upper.Y) < 0.005d,
                "Every stacked pair must share the value-axis baseline and meet at its segment boundary.");
            TestAssert.True(Math.Abs(lower.Height - valueUnit * lowerValue) < 0.005d &&
                Math.Abs(upper.Height - valueUnit * upperValue) < 0.005d,
                $"The separate gradient clips must cover the authored stack {lowerValue} + {upperValue}.");
        }
    }

    private static string RenderStyle18BarChart(string direction, string grouping, bool includeSupply)
    {
        static string Series(string name, int categoryA, int categoryB) => FormattableString.Invariant($$"""
            <c:ser>
              <c:tx><c:strLit><c:pt idx="0"><c:v>{{name}}</c:v></c:pt></c:strLit></c:tx>
              <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>{{categoryA}}</c:v></c:pt><c:pt idx="1"><c:v>{{categoryB}}</c:v></c:pt></c:numLit></c:val>
            </c:ser>
            """);
        string series = Series("Demand", 2, 4);
        if (includeSupply) series += Series("Supply", 1, 3);
        using var input = TestFixtures.CreateZipPackage(ThemedBarChartPackage($$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:style val="18"/><c:chart><c:plotArea><c:barChart>
                <c:barDir val="{{direction}}"/><c:grouping val="{{grouping}}"/><c:varyColors val="0"/>
                {{series}}
                <c:axId val="10"/><c:axId val="20"/>
              </c:barChart>
              <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx>
              <c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """));
        using var output = new MemoryStream();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx });
        string pdf = Encoding.ASCII.GetString(output.ToArray());
        TestAssert.DoesNotContain("0.082 0.376 0.51 rg", pdf);
        return pdf;
    }

    private static PdfRectangle[] ReadStyle18BarGradientClips(string pdf)
    {
        // Each complete bar paint is a rectangle clip followed by its shading operation.
        var clips = Regex.Matches(pdf, @"(?m)^([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re W n\r?\n/Sh\d+ sh\r?$")
            .Select(match => new PdfRectangle(ReadCoordinate(match, 1), ReadCoordinate(match, 2),
                ReadCoordinate(match, 3), ReadCoordinate(match, 4)))
            .ToArray();
        TestAssert.True(clips.All(clip => clip.Width > 0d && clip.Height > 0d),
            "Every painted bar or segment requires a nonempty clip.");
        return clips;

        static double ReadCoordinate(Match match, int index) =>
            double.Parse(match.Groups[index].Value, CultureInfo.InvariantCulture);
    }

    public static void PptxSyntheticBarChartLegendSwatchesUseSeriesStroke()
    {
        const string legendMarkup = "<c:legend><c:legendPos val=\"r\"/><c:overlay val=\"0\"/></c:legend>";
        const string chartPart = "ppt/charts/chart1.xml";
        var parts = new Dictionary<string, byte[]>
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
            [chartPart] = TestFixtures.Utf8($$"""
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
                </c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea>{{legendMarkup}}</c:chart></c:chartSpace>
                """)
        };
        PdfPage page = Render();
        AssertSwatches(page);

        const string demandStroke = "0.071 0.204 0.337 RG";
        const string supplyStroke = "0.396 0.263 0.129 RG";
        string swappedContent = string.Join('\n', page.Content.Split('\n').Select(line => line.TrimEnd('\r') switch
        {
            demandStroke => supplyStroke,
            supplyStroke => demandStroke,
            _ => line
        }));
        PdfPage swapped = page with { Content = swappedContent };
        var failure = TestAssert.Throws<InvalidOperationException>(() => AssertSwatches(swapped));
        TestAssert.Contains("Demand legend swatch stroke", failure.Message);

        parts[chartPart] = TestFixtures.Utf8(Encoding.UTF8.GetString(parts[chartPart]).Replace(legendMarkup, "", StringComparison.Ordinal));
        PdfPage hiddenLegend = Render();
        TestAssert.Equal(0, DocxCommentsTests.ReadEmbeddedGlyphTextShows(hiddenLegend).Count());
        TestAssert.Equal(0, ReadRectangles(hiddenLegend, strokes: true).Length);
        TestAssert.Equal(4, ReadRectangles(hiddenLegend, strokes: false).Length);

        PdfPage Render()
        {
            using var input = TestFixtures.CreateZipPackage(parts);
            using var package = OoxPackage.Open(input, default);
            var diagnostics = new DiagnosticCollector();
            var document = new PptxReader().Read(package, default);
            PdfPage rendered = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(document, package, diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.True(diagnostics.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"),
                "Both bar chart controls should render without unsupported chart diagnostics.");
            return rendered;
        }

        static void AssertSwatches(PdfPage rendered)
        {
            var fills = ReadRectangles(rendered, strokes: false);
            var strokes = ReadRectangles(rendered, strokes: true);
            TestAssert.Equal(2, strokes.Length);
            AssertSwatch("Demand", new RgbColor(68, 114, 196), new RgbColor(18, 52, 86));
            AssertSwatch("Supply", new RgbColor(237, 125, 49), new RgbColor(101, 67, 33));

            void AssertSwatch(string name, RgbColor expectedFill, RgbColor expectedStroke)
            {
                var firstGlyph = PortableRenderingTests.ReadPaintedSourceLabel(rendered, name, fontPrefix: "CL")[0];
                var key = strokes.Single(paint =>
                {
                    PdfRectangle rect = paint.Rectangle;
                    double gap = firstGlyph.Origin.X - (rect.X + rect.Width);
                    return gap >= 0 && gap <= 2d * rect.Width && Math.Abs(firstGlyph.Origin.Y - rect.Y) <= rect.Height;
                });
                TestAssert.Equal(expectedFill, fills.Single(paint => paint.Rectangle == key.Rectangle).Color);
                TestAssert.True(expectedStroke == key.Color, $"{name} legend swatch stroke must match its source series.");
            }
        }

        static (RgbColor Color, PdfRectangle Rectangle)[] ReadRectangles(PdfPage rendered, bool strokes)
        {
            var operations = new List<string>();
            RgbColor[] colors = PptxChartPaletteTests.ReadColoredPaints(rendered.Content, operations.Add, strokes).ToArray();
            return colors.Zip(operations).Where(paint => paint.Second.EndsWith(strokes ? " re S" : " re f", StringComparison.Ordinal))
                .Select(paint =>
                {
                    double[] bounds = paint.Second[..^5].Split(' ', StringSplitOptions.RemoveEmptyEntries)
                        .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
                    return (paint.First, new PdfRectangle(bounds[0], bounds[1], bounds[2], bounds[3]));
                }).ToArray();
        }
    }

    public static void PptxSyntheticLineChartOmitsLegendKeyOnlyDataLabels()
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
                  <c:dLbls><c:showVal val="0"/><c:showLegendKey val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:dLblPos val="t"/><c:spPr><a:solidFill><a:srgbClr val="2468AC"/></a:solidFill></c:spPr></c:dLbls>
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
        TestAssert.True(!pdf.Contains("0.141 0.408 0.675 rg", StringComparison.Ordinal), "A legend key alone does not create a label background.");
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
        TestAssert.True(Regex.IsMatch(pdf, @"0\.671 0\.804 0\.937 RG\s+1 w"), "Expected line-chart marker <a:ln> without @w to inherit the 1 pt explicit-outline width (COM square deck falsified the 0.75 pt extension, which covers unstyled outlines only).");
        TestAssert.True(!Regex.IsMatch(pdf, @"0\.671 0\.804 0\.937 RG\s+3 w"), "Marker <a:ln> without @w should not use the 3 pt series-line inherited width.");
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
        object workbook = PptxChartRangeTests.ReadWorkbookCells(sheets);
        System.Reflection.MethodInfo readExplosions = typeof(PptxRenderer).GetMethod(
            "ReadSceneChartPointExplosions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected point-explosion helper.");

        var explosions = (IReadOnlyDictionary<int, double>)(readExplosions.Invoke(null, [plot, workbook])
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
        PptxSceneChartPlot plot = PptxTests.BuildSingleChartScene(chartXml)?.Plots.Single()
            ?? throw new InvalidOperationException("Expected one chart plot.");
        System.Reflection.MethodInfo readExplosions = typeof(PptxRenderer).GetMethod(
            "ReadSceneChartPointExplosions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected point-explosion helper.");

        var explosions = (IReadOnlyDictionary<int, double>)(readExplosions.Invoke(null, [plot, null])
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
        TestAssert.True(chart.Plots.Single().DataLabels.IsDefined, "Expected a plot-level data-label definition.");
        TestAssert.Equal(1, chart.Plots.Single().Series.Count);
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
                  <Relationship Id="rId0" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="https://example.com/colors.xml" TargetMode="External"/>
                  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
                  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="missing-colors.xml"/>
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
                    <a:srgbClr val="112233"><a:alpha val="50000"/><a:alphaMod val="50000"/><a:alphaOff val="2500"/></a:srgbClr>
                  </cs:variation>
                </cs:colorStyle>
                """)
        });
        PptxSceneChart chart = new PptxSceneBuilder().Build(document, package, CancellationToken.None).Slides[0].SlideNodes[0].Chart
            ?? throw new InvalidOperationException("Expected chart scene.");

        TestAssert.Equal("/ppt/charts/colors1.xml", chart.ColorStyle.PartName);
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
        TestAssert.True(Math.Abs(chart.ColorStyle.Declarations[2].Alpha - 0.275d) < 0.000001d,
            "Direct color declarations must apply alpha transforms in source order.");
        TestAssert.True(chart.ColorStyle.Declarations[0].IsResolved, "Expected direct RGB color-style declaration to resolve.");
        TestAssert.True(!chart.ColorStyle.Declarations[1].IsResolved, "Expected unresolved scheme color-style declaration to remain observable.");
        TestAssert.Equal(1, chart.ColorStyle.Colors.Count);
        TestAssert.Equal(2, chart.ColorStyle.Declarations.Count(declaration => declaration.IsResolved));
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
        string fixture = TestFixtures.PublicCasePath("pptx-ladder-11-chart-column-100-stacked-port.pptx");
        using var original = File.OpenRead(fixture);
        PdfPage page = Render(original);
        TestAssert.Equal("0%|10%|20%|30%|40%|50%|60%|70%|80%|90%|100%", Labels(page, "CVA"));

        // A percent sign painted by a category cannot stand in for the value-axis labels.
        using var hiddenAxis = new MemoryStream();
        original.Position = 0;
        original.CopyTo(hiddenAxis);
        hiddenAxis.Position = 0;
        using (var archive = new ZipArchive(hiddenAxis, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("ppt/charts/chart1.xml")!;
            XDocument chart;
            using (var input = entry.Open()) chart = XDocument.Load(input);
            XNamespace c = chart.Root!.Name.Namespace;
            chart.Descendants(c + "valAx").Single().Element(c + "tickLblPos")!.SetAttributeValue("val", "none");
            foreach (var category in chart.Descendants(c + "cat"))
                category.Descendants(c + "v").First().Value = "A%";
            entry.Delete();
            using var output = archive.CreateEntry("ppt/charts/chart1.xml").Open();
            chart.Save(output);
        }
        hiddenAxis.Position = 0;
        PdfPage control = Render(hiddenAxis);
        TestAssert.Equal("A%|B|C", Labels(control, "CCA"));
        TestAssert.Equal(string.Empty, Labels(control, "CVA"));

        static PdfPage Render(Stream input)
        {
            using var package = OoxPackage.Open(input, default);
            var diagnostics = new DiagnosticCollector();
            var document = new PptxReader().Read(package, default, diagnostics.Add);
            PdfPage page = new PptxRenderer(new TestFaceFontResolver())
                .RenderPagesAsync(document, package, diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.True(diagnostics.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"),
                "Percent-stacked column charts should not emit unsupported chart diagnostics.");
            return page;
        }

        static string Labels(PdfPage page, string axis) => string.Join('|', DocxCommentsTests.ReadEmbeddedGlyphTextShows(
            page, usePageCoordinates: false, fontResourceFilter: name => name.StartsWith(axis, StringComparison.Ordinal))
            .Select(show => show.Text));
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
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutEffects, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal), null]);

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
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutGradients, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal), null]);

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
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutPictureFills, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal), null]);

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
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutBodyProperties, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal), null]);

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
        PptxSceneSlide sceneSlide = new PptxSceneBuilder().Build(document, package, default).Slides.Single();
        PptxSceneChart chart = TestAssert.NotNull(sceneSlide.SlideNodes.Single().Chart);
        TestAssert.Equal("ellipsis", chart.Legend.TextBodyProperties.VerticalOverflowValue);
        TestAssert.Equal("ellipsis", chart.Plots.Single().DataLabels.TextBodyProperties.VerticalOverflowValue);

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
        emitDiagnostics.Invoke(null, [sceneSlide, slideXmlWithoutBodyProperties, "/ppt/slides/slide1.xml", 1, sink, new HashSet<string>(StringComparer.Ordinal), null]);

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

    public static void PptxSceneScatterSeriesLineNoFillSurvivesInSceneModel()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea>
                <c:scatterChart>
                  <c:scatterStyle val="lineMarker"/>
                  <c:ser><c:spPr><a:ln w="47625"><a:noFill/></a:ln></c:spPr><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser>
                  <c:ser><c:xVal><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:yVal></c:ser>
                </c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartPlot plot = chart?.Plots[0] ?? throw new InvalidOperationException("Expected scatter-chart plot.");
        TestAssert.Equal(2, plot.Series.Count);
        TestAssert.True(plot.Series[0].Line.NoFill, "Expected explicit series line noFill to survive in the scene model.");
        TestAssert.True(!plot.Series[1].Line.NoFill, "Expected an absent series line to remain distinct from explicit noFill.");
    }

    public static void PptxScatterLineConnectionCoversSmoothStyles()
    {
        System.Reflection.MethodInfo read = typeof(PptxRenderer).GetMethod(
            "ReadSceneChartScatterOptions",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected scatter option reader.");
        bool Invoke(PptxSceneChartScatterStyle style)
        {
            XElement source = XElement.Parse($"<c:scatterChart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:scatterStyle val=\"{style}\"/></c:scatterChart>");
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            var xml = new XElement(c + "chartSpace", new XElement(c + "chart", new XElement(c + "plotArea", source)));
            var plot = TestAssert.NotNull(PptxTests.BuildSingleChartScene(xml.ToString())).Plots[0];
            object options = read.Invoke(null, [plot])!;
            return (bool)options.GetType().GetProperty("ConnectLines")!.GetValue(options)!;
        }
        TestAssert.True(Invoke(PptxSceneChartScatterStyle.Line), "Line scatter style should connect points.");
        TestAssert.True(Invoke(PptxSceneChartScatterStyle.LineMarker), "LineMarker scatter style should connect points.");
        TestAssert.True(Invoke(PptxSceneChartScatterStyle.Smooth), "Smooth scatter style should connect points.");
        TestAssert.True(Invoke(PptxSceneChartScatterStyle.SmoothMarker), "SmoothMarker scatter style should connect points.");
        TestAssert.True(Invoke(PptxSceneChartScatterStyle.Marker), "Office's marker-only style keeps the series connection unless its line is hidden.");
        TestAssert.True(!Invoke(PptxSceneChartScatterStyle.None), "None scatter style should not connect points.");
        TestAssert.True(!Invoke(PptxSceneChartScatterStyle.Unknown), "Unknown scatter style should not connect points.");
    }

    public static void PptxSyntheticScatterSmoothConnectUsesLineWidth()
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
                  <c:scatterStyle val="smoothMarker"/>
                  <c:ser><c:tx><c:strLit><c:pt idx="0"><c:v>S</c:v></c:pt></c:strLit></c:tx><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:yVal></c:ser>
                </c:scatterChart></c:plotArea></c:chart></c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("2.25 w", pdf);
    }
}
