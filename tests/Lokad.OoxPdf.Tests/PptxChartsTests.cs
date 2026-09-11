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

internal static class PptxChartsTests
{
    public static void PptxChartTextConversionUsesCustomFontResolver()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
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
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld>
                    <p:spTree>
                      <p:graphicFrame>
                        <p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:nvPr/></p:nvGraphicFramePr>
                        <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm>
                        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic>
                      </p:graphicFrame>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """,
            ["ppt/charts/chart1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart>
                    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Chart resolver probe</a:t></a:r></a:p></c:rich></c:tx></c:title>
                    <c:plotArea>
                      <c:barChart>
                        <c:barDir val="col"/>
                        <c:grouping val="clustered"/>
                        <c:ser>
                          <c:idx val="0"/><c:order val="0"/>
                          <c:tx><c:strLit><c:pt idx="0"><c:v>Series</c:v></c:pt></c:strLit></c:tx>
                          <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                          <c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                        <c:axId val="10"/><c:axId val="20"/>
                      </c:barChart>
                      <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/><c:tickLblPos val="none"/></c:catAx>
                      <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:crossAx val="10"/><c:majorUnit val="5"/><c:tickLblPos val="none"/></c:valAx>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var resolver = new CountingFontResolver();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver });

        TestAssert.True(resolver.ResolveCalls > 0, "PPTX chart text conversion should use the supplied font resolver.");
    }

    public static void PptxChartSeriesLineNoFillSuppressesBarOutlines()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>
                <c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:grouping val="stacked"/>
                    <c:ser>
                      <c:idx val="0"/><c:order val="0"/>
                      <c:spPr>
                        <a:solidFill><a:srgbClr val="00B050"/></a:solidFill>
                        <a:ln><a:noFill/></a:ln>
                      </c:spPr>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:axId val="10"/><c:axId val="20"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/>
                    <c:spPr><a:ln><a:noFill/></a:ln></c:spPr>
                    <c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="20"/><c:axPos val="l"/>
                    <c:scaling><c:min val="0"/><c:max val="5"/></c:scaling>
                    <c:majorUnit val="5"/><c:crossAx val="10"/>
                    <c:spPr><a:ln><a:noFill/></a:ln></c:spPr>
                    <c:tickLblPos val="none"/>
                  </c:valAx>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        PptxSceneChartSeries series = chart.Plots[0].Series[0];
        TestAssert.True(series.Fill.HasFill, "Expected chart series fill to remain visible when only the line is noFill.");
        TestAssert.True(series.Line.HasLine == false, "Expected explicit chart series line noFill to suppress bar outlines.");
    }

    public static void PptxChartDataLabelOverridePreservesDelete()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:dLbls>
                      <c:showVal val="1"/>
                      <c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl>
                    </c:dLbls>
                    <c:ser>
                      <c:idx val="0"/>
                      <c:order val="0"/>
                      <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:axId val="1"/>
                    <c:axId val="2"/>
                  </c:barChart>
                  <c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx>
                  <c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        PptxSceneChartDataLabelOverride labelOverride = chart.Plots[0].DataLabels.Overrides.Single();
        TestAssert.Equal(1, labelOverride.Index);
        TestAssert.True(labelOverride.IsDeleted == true, "Expected point-level data label delete to be preserved.");
        TestAssert.Equal("1", labelOverride.IsDeletedValue);
    }

    public static void PptxSyntheticBarChartsRenderNativeCharts()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes().Replace(
                "</Types>",
                "  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\r\n</Types>",
                StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/theme/theme1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ChartTheme">
                  <a:themeElements>
                    <a:clrScheme name="ChartTheme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent1><a:srgbClr val="00AA00"/></a:accent1>
                    </a:clrScheme>
                    <a:fontScheme name="ChartTheme"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
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
                    <p:graphicFrame>
                      <p:xfrm><a:off x="5029200" y="914400"/><a:ext cx="2286000" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100"><a:latin typeface="+mn-lt"/></a:defRPr></a:pPr></a:p></c:txPr>
                  <c:chart><c:plotArea><c:barChart>
                    <c:ser><c:tx><c:v>Primary A</c:v></c:tx><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr><c:val><c:numLit>
                      <c:pt idx="0"><c:v>2</c:v></c:pt>
                      <c:pt idx="1"><c:v>4</c:v></c:pt>
                    </c:numLit></c:val></c:ser>
                    <c:ser><c:tx><c:v>Primary B</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill></c:spPr><c:val><c:numLit>
                      <c:pt idx="0"><c:v>1</c:v></c:pt>
                      <c:pt idx="1"><c:v>3</c:v></c:pt>
                    </c:numLit></c:val></c:ser>
                    <c:dLbls><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/><c:dLblPos val="b"/><c:separator> | </c:separator><c:spPr><a:solidFill><a:srgbClr val="FFEACC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="0A0B0C"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr><c:dLbl><c:idx val="1"/><c:dLblPos val="ctr"/><c:spPr><a:solidFill><a:srgbClr val="CCEEFF"/></a:solidFill></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="6600CC"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbl></c:dLbls>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:barChart>
                    <c:barDir val="col"/><c:grouping val="stacked"/>
                    <c:ser><c:tx><c:v>Secondary</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="0000AA"/></a:solidFill></c:spPr><c:val><c:numLit>
                      <c:pt idx="0"><c:v>0</c:v></c:pt>
                      <c:pt idx="1"><c:v>20</c:v></c:pt>
                    </c:numLit></c:val></c:ser>
                    <c:dLbls><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/><c:separator> | </c:separator></c:dLbls>
                    <c:axId val="3"/><c:axId val="4"/>
                  </c:barChart>
                  <c:valAx>
                    <c:axId val="2"/>
                    <c:delete val="1"/>
                    <c:axPos val="l"/>
                    <c:scaling><c:min val="0"/><c:max val="4"/></c:scaling>
                    <c:majorUnit val="2"/>
                    <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:pPr></a:p></c:txPr>
                  </c:valAx>
                  <c:valAx>
                    <c:axId val="4"/>
                    <c:axPos val="r"/>
                    <c:scaling><c:min val="0"/><c:max val="40"/></c:scaling>
                    <c:majorUnit val="20"/>
                    <c:spPr><a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr>
                    <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:pPr></a:p></c:txPr>
                  </c:valAx>
                  </c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/chart2.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea><c:barChart>
                    <c:varyColors/>
                    <c:ser>
                    <c:dPt><c:idx val="1"/><c:spPr><a:pattFill prst="ltUpDiag"><a:fgClr><a:srgbClr val="2F856A"/></a:fgClr><a:bgClr><a:srgbClr val="EEEEEE"/></a:bgClr></a:pattFill></c:spPr></c:dPt>
                    <c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="AA00AA"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00AAAA"/></a:solidFill></a:ln></c:spPr></c:dPt>
                    <c:val><c:numLit>
                      <c:pt idx="0"><c:v>2</c:v></c:pt>
                      <c:pt idx="1"><c:v>4</c:v></c:pt>
                      <c:pt idx="2"><c:v>3</c:v></c:pt>
                    </c:numLit></c:val></c:ser>
                  </c:barChart></c:plotArea></c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/_rels/chart2.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors2.xml"/>
                </Relationships>
                """),
            ["ppt/charts/colors2.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                               meth="cycle" id="10">
                  <a:srgbClr val="FF00CC"/>
                  <a:schemeClr val="accent1"/>
                </cs:colorStyle>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0.8 rg", pdf);
        TestAssert.Contains("0 0.667 0 rg", pdf);
        TestAssert.Contains("0.667 0 0 rg", pdf);
        TestAssert.Contains("0 0 0.667 rg", pdf);
        TestAssert.Contains("0.667 0 0.667 rg", pdf);
        TestAssert.Contains("0 0.667 0.667 RG", pdf);
        TestAssert.Contains("0.071 0.204 0.337 RG", pdf);
        TestAssert.Contains("0.067 0.133 0.2 rg", pdf);
        TestAssert.Contains("/Pattern cs", pdf);
        TestAssert.Contains("/ImPattern Do", pdf);
        TestAssert.Contains(" re W* n", pdf);
        TestAssert.Contains("BT", pdf);
        TestAssert.Contains(" re f", pdf);
        TestAssert.True(pdf.Split("BT", StringSplitOptions.None).Length >= 17, "Combo bar charts should emit text for primary and secondary labels, axes, and legend entries.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Supported chart rendering should not emit static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Supported bar charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticChartAxisTickLabelOptionsRender()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>Hidden category</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1000</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:tickLblPos val="none"/>
                    <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/>
                    <c:tickLblPos val="high"/>
                    <c:scaling><c:min val="0"/><c:max val="2000"/></c:scaling><c:majorUnit val="1000"/>
                    <c:numFmt formatCode="$#,##0"/>
                    <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0.667 0 rg", pdf);
        TestAssert.DoesNotContain("1 0 0 rg", pdf);
        TestAssert.Contains("<0024>", pdf);
        // Value labels moved with the plot edge, which now rests on the preset floor since
        // the recalibrated label reserve (frame indent plus font-relative gap) fits inside it.
        TestAssert.Contains("323.014", pdf);
    }

    public static void PptxSyntheticChartCategoryAxisLabelOffsetRender()
    {
        static string RenderPdf(int labelOffset)
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
                ["ppt/charts/chart1.xml"] = TestFixtures.Utf8($$"""
                    <?xml version="1.0" encoding="UTF-8"?>
                    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <c:chart><c:plotArea>
                      <c:barChart>
                        <c:barDir val="col"/>
                        <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>Offset category</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1000</c:v></c:pt></c:numLit></c:val></c:ser>
                        <c:axId val="1"/><c:axId val="2"/>
                      </c:barChart>
                      <c:catAx>
                        <c:axId val="1"/><c:axPos val="b"/><c:lblOffset val="{{labelOffset}}"/>
                        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                      </c:catAx>
                      <c:valAx>
                        <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                        <c:scaling><c:min val="0"/><c:max val="2000"/></c:scaling><c:majorUnit val="1000"/>
                      </c:valAx>
                      </c:plotArea></c:chart>
                    </c:chartSpace>
                    """)
            });
            string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

            OoxPdfConverter.Convert(input, output);

            return File.ReadAllText(output, Encoding.ASCII);
        }

        string defaultPdf = RenderPdf(100);
        string offsetPdf = RenderPdf(200);
        double defaultBaseline = PptxTests.ReadOnlyTextBaseline(defaultPdf);
        double offsetBaseline = PptxTests.ReadOnlyTextBaseline(offsetPdf);

        TestAssert.Contains("1 0 0 rg", defaultPdf);
        TestAssert.Contains("1 0 0 rg", offsetPdf);
        TestAssert.True(offsetBaseline < defaultBaseline - 5d, $"Expected lblOffset=200 to move the category label farther from the plot area. Default: {defaultBaseline.ToString(CultureInfo.InvariantCulture)}; offset: {offsetBaseline.ToString(CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticChartCategoryAxisTickLabelSkipRender()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser>
                      <c:cat><c:strLit>
                        <c:pt idx="0"><c:v>North</c:v></c:pt>
                        <c:pt idx="1"><c:v>South</c:v></c:pt>
                        <c:pt idx="2"><c:v>East</c:v></c:pt>
                        <c:pt idx="3"><c:v>West</c:v></c:pt>
                      </c:strLit></c:cat>
                      <c:val><c:numLit>
                        <c:pt idx="0"><c:v>10</c:v></c:pt>
                        <c:pt idx="1"><c:v>20</c:v></c:pt>
                        <c:pt idx="2"><c:v>30</c:v></c:pt>
                        <c:pt idx="3"><c:v>40</c:v></c:pt>
                      </c:numLit></c:val>
                    </c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblSkip val="2"/>
                    <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="50"/></c:scaling><c:majorUnit val="10"/>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Equal(2, PptxTests.CountTextMatrices(pdf));
    }

    public static void PptxSyntheticChartCategoryGridlinesDoNotEnableValueGridlines()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:majorGridlines/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling><c:majorUnit val="10"/>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(!Regex.IsMatch(pdf, @"0\.75 w\s+(?:[0-9.]+ [0-9.]+ m\s+[0-9.]+ [0-9.]+ l\s+){2,}S"),
            "Category-axis gridlines must not enable value-axis gridline rendering.");
    }

    public static void PptxSyntheticChartValueGridlinesUseOfficeLikeMultiSegmentPath()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling><c:majorUnit val="10"/><c:majorGridlines/>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 G", pdf);
        TestAssert.Contains("0.75 w", pdf);
        TestAssert.DoesNotContain("0.851 G", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"(?:[0-9.]+ [0-9.]+ m\s+[0-9.]+ [0-9.]+ l\s+){3}S"),
            "Major value gridlines should be emitted as one multi-segment stroked path that includes the maximum tick and excludes the axis baseline.");
    }

    public static void PptxSyntheticChartHorizontalValueAxisAutoUnitUsesOfficeDenseTicks()
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
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="9144000" cy="5486400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="bar"/>
                    <c:grouping val="clustered"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>45</c:v></c:pt><c:pt idx="1"><c:v>32</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="l"/><c:crossAx val="2"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="b"/><c:scaling/><c:majorGridlines/><c:tickLblPos val="none"/><c:crossAx val="1"/><c:crosses val="autoZero"/>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"(?:[0-9.]+ [0-9.]+ m\s+[0-9.]+ [0-9.]+ l\s+){10}S"),
            "Horizontal bar value-axis auto units should use Office-like 5-unit ticks for a 0..50 axis, excluding the zero crossing gridline.");
    }

    public static void PptxSyntheticChartHorizontalValueAxisVisibleLabelsWithoutManualBoxUseOfficeDenseTicks()
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
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="9144000" cy="5486400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="bar"/>
                    <c:grouping val="clustered"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>45</c:v></c:pt><c:pt idx="1"><c:v>32</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="l"/><c:crossAx val="2"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="b"/><c:scaling/><c:majorGridlines/><c:crossAx val="1"/><c:crosses val="autoZero"/>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"(?:[0-9.]+ [0-9.]+ m\s+[0-9.]+ [0-9.]+ l\s+){10}S"),
            "Visible horizontal bar value-axis labels without manual plot layout should use Office-like 5-unit ticks for a 0..50 axis, excluding the zero crossing gridline.");
    }

    public static void PptxSyntheticChartHorizontalValueAxisVisibleLabelsWithManualBoxUseOfficeSparseTicks()
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
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="9144000" cy="5486400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea>
                  <c:layout><c:manualLayout><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0.18"/><c:y val="0.14"/><c:w val="0.62"/><c:h val="0.68"/></c:manualLayout></c:layout>
                  <c:barChart>
                    <c:barDir val="bar"/>
                    <c:grouping val="clustered"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>45</c:v></c:pt><c:pt idx="1"><c:v>32</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="l"/><c:crossAx val="2"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="b"/><c:scaling/><c:majorGridlines/><c:crossAx val="1"/><c:crosses val="autoZero"/>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"(?:[0-9.]+ [0-9.]+ m\s+[0-9.]+ [0-9.]+ l\s+){5}S"),
            "Visible horizontal bar value-axis labels should use Office-like 10-unit ticks for a 0..50 axis, excluding the zero crossing gridline.");
    }

    public static void PptxSyntheticChartLineAxisNearMaximumKeepsOfficeHeadroom()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "GetNiceChartAxisMax",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart axis maximum helper to remain inspectable by the Office evidence guard.");

        double lineMarkerMax = (double)method!.Invoke(null, [96d, 0d, 8d, true, 0.96d, false])!;
        double ordinaryMax = (double)method.Invoke(null, [96d, 0d, 8d, false, 0.96d, false])!;
        double lineThreeSeriesMax = (double)method.Invoke(null, [1520d, 0d, 8d, true, 0.96d, false])!;

        TestAssert.Equal(120d, lineMarkerMax);
        TestAssert.Equal(100d, ordinaryMax);
        TestAssert.Equal(1600d, lineThreeSeriesMax);
    }

    public static void PptxSyntheticChartScatterAxisNearMaximumKeepsOfficeHeadroom()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "GetNiceChartAxisMax",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart axis maximum helper to remain inspectable by the Office evidence guard.");

        double scatterClustersMax = (double)method!.Invoke(null, [4.8d, 0d, 5d, true, 0.96d, false])!;
        double scatterClustersNoHeadroom = (double)method.Invoke(null, [4.8d, 0d, 5d, false, 0.96d, false])!;
        double bubbleMax = (double)method.Invoke(null, [4.0d, 0d, 5d, true, 0.96d, false])!;
        double bubbleNoHeadroom = (double)method.Invoke(null, [4.0d, 0d, 5d, false, 0.96d, false])!;

        TestAssert.Equal(6d, scatterClustersMax);
        TestAssert.Equal(5d, scatterClustersNoHeadroom);
        TestAssert.Equal(5d, bubbleMax);
        TestAssert.Equal(5d, bubbleNoHeadroom);
    }

    public static void PptxSyntheticScatterXAxisUnitHalvesEarlierThanY()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ChooseChartAxisMajorUnit",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart axis unit helper to remain inspectable by the Office evidence guard.");

        TestAssert.Equal(1d, (double)method!.Invoke(null, [8d, 8d])!);
        TestAssert.Equal(2d, (double)method.Invoke(null, [9d, 8d])!);
        TestAssert.Equal(2d, (double)method.Invoke(null, [9.5d, 8d])!);
        TestAssert.Equal(2d, (double)method.Invoke(null, [12d, 8d])!);
        TestAssert.Equal(1d, (double)method.Invoke(null, [9.5d, 10d])!);
        TestAssert.Equal(1d, (double)method.Invoke(null, [9d, 10d])!);
    }

    public static void PptxSyntheticScatterAxisMaxPrefersUnitOneOverTwo()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "GetNiceChartAxisMax",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart axis maximum helper to remain inspectable by the Office evidence guard.");

        TestAssert.Equal(6d, (double)method!.Invoke(null, [5.5d, 0d, 5d, true, 0.96d, true])!);
        TestAssert.Equal(7d, (double)method.Invoke(null, [6.5d, 0d, 5d, true, 0.96d, true])!);
        TestAssert.Equal(9d, (double)method.Invoke(null, [7.992d, 0d, 5d, true, 0.96d, true])!);
        TestAssert.Equal(9d, (double)method.Invoke(null, [8.5d, 0d, 5d, true, 0.96d, true])!);
        TestAssert.Equal(10d, (double)method.Invoke(null, [9.5d, 0d, 5d, true, 0.96d, true])!);
        TestAssert.Equal(6d, (double)method.Invoke(null, [4.8d, 0d, 5d, true, 0.96d, true])!);
        TestAssert.Equal(5d, (double)method.Invoke(null, [4.0d, 0d, 5d, true, 0.96d, true])!);
        TestAssert.Equal(10d, (double)method.Invoke(null, [7.992d, 0d, 5d, true, 0.96d, false])!);
        TestAssert.Equal(8d, (double)method.Invoke(null, [6.5d, 0d, 5d, true, 0.96d, false])!);
        TestAssert.Equal(10d, (double)method.Invoke(null, [9.5d, 0d, 8d, true, 0.96d, false])!);
        TestAssert.Equal(14d, (double)method.Invoke(null, [12d, 0d, 8d, true, 0.96d, false])!);
        TestAssert.Equal(8d, (double)method.Invoke(null, [7.5d, 0d, 8d, true, 0.96d, false])!);
        TestAssert.Equal(5d, (double)method.Invoke(null, [4.5d, 0d, 8d, true, 0.96d, false])!);
        TestAssert.Equal(3.5d, (double)method.Invoke(null, [3d, 0d, 8d, true, 0.96d, false])!);
        TestAssert.Equal(3d, (double)method.Invoke(null, [2.5d, 0d, 8d, true, 0.96d, false])!);
        TestAssert.Equal(3.5d, (double)method.Invoke(null, [3.1d, 0d, 8d, true, 0.96d, false])!);

    }

    public static void PptxSyntheticChartNoTitleRightLegendLeftInsetKeepsOfficePadding()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeNoTitleRightLegendLeftInset",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart left-inset helper to remain inspectable by the Office evidence guard.");

        double singleDigit = (double)method!.Invoke(null, [9.12d, 720d])!;
        double threeDigit = (double)method.Invoke(null, [27.37d, 720d])!;
        double fourDigit = (double)method.Invoke(null, [36.49d, 720d])!;
        double narrowFrame = (double)method.Invoke(null, [9.12d, 400d])!;

        TestAssert.True(Math.Abs(singleDigit - 32.32d) < 0.0001d, "Single-digit inset should be label plus 23.2pt Office padding. Got " + singleDigit);
        TestAssert.True(Math.Abs(threeDigit - 50.57d) < 0.0001d, "Three-digit inset should use the same padding. Got " + threeDigit);
        TestAssert.True(Math.Abs(fourDigit - 59.69d) < 0.0001d, "Four-digit inset must not add a character-count extra. Got " + fourDigit);
        TestAssert.True(Math.Abs(narrowFrame - 29.12d) < 0.0001d, "Narrow frames keep the width-ratio padding cap. Got " + narrowFrame);
    }

    public static void PptxSyntheticChartRightLegendReserveKeepsOfficeTail()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeRightLegendReservePadding",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart right-reserve helper to remain inspectable by the Office evidence guard.");

        double lineScatter = (double)method!.Invoke(null, [18d, 9, 720d, false, 0d, 0d])!;
        double shortName = (double)method.Invoke(null, [18d, 5, 720d, false, 0d, 0d])!;
        double scatterWideX = (double)method.Invoke(null, [18d, 9, 720d, false, 0d, 18.25d])!;
        double areaWideCats = (double)method.Invoke(null, [18d, 9, 720d, true, 36.56d, 0d])!;
        double areaNarrowCats = (double)method.Invoke(null, [18d, 9, 720d, true, 21.25d, 0d])!;
        double areaNoCats = (double)method.Invoke(null, [18d, 9, 720d, true, 0d, 0d])!;

        TestAssert.True(Math.Abs(lineScatter - 47.1d) < 0.001d, "Line/scatter tail should be marker block plus 10.8pt Office padding. Got " + lineScatter);
        TestAssert.True(Math.Abs(shortName - 47.1d) < 0.001d, "Short names must not shrink the tail: the widest name already spans the text. Got " + shortName);
        TestAssert.True(Math.Abs(scatterWideX - 55.425d) < 0.001d, "Scatter padding should add half the overhanging edge label with its own tail. Got " + scatterWideX);
        TestAssert.True(Math.Abs(areaWideCats - 60.33d) < 0.001d, "Area padding should be the 42.05pt block plus half the last category label. Got " + areaWideCats);
        TestAssert.True(Math.Abs(areaNarrowCats - 52.675d) < 0.001d, "Narrow edge labels should shrink the area padding. Got " + areaNarrowCats);
        TestAssert.True(Math.Abs(areaNoCats - 42.05d) < 0.001d, "Missing edge labels should leave the bare block. Got " + areaNoCats);
    }

    public static void PptxSyntheticChartHorizontalBarCategoryReserveKeepsOfficeGap()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeHorizontalBarCategoryLeftReserve",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart category-reserve helper to remain inspectable by the Office evidence guard.");

        double shortLabels = (double)method!.Invoke(null, [49.9d])!;
        double longLabels = (double)method.Invoke(null, [86.5d])!;
        double noLabels = (double)method.Invoke(null, [0d])!;

        TestAssert.True(Math.Abs(shortLabels - 73.1d) < 0.0001d, "Short-label reserve should be label plus 23.2pt Office gap. Got " + shortLabels);
        TestAssert.True(Math.Abs(longLabels - 109.7d) < 0.0001d, "Long-label reserve should use the same gap. Got " + longLabels);
        TestAssert.True(Math.Abs(noLabels - 23.2d) < 0.0001d, "Empty labels keep the bare gap below every preset. Got " + noLabels);
    }

    public static void PptxSyntheticChartBarValueAxisReserveKeepsOfficeGap()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeBarValueAxisLeftReserve",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart value-reserve helper to remain inspectable by the Office evidence guard.");

        double threeDigit = (double)method!.Invoke(null, [27.37d, 18d])!;
        double twoDigit = (double)method.Invoke(null, [18.25d, 18d])!;
        double compositeNarrow = (double)method.Invoke(null, [6.41d, 12d])!;

        TestAssert.True(Math.Abs(threeDigit - 50.43d) < 0.01d, "Three-digit reserve should be label plus indent plus font-relative gap. Got " + threeDigit);
        TestAssert.True(Math.Abs(twoDigit - 41.31d) < 0.01d, "Two-digit reserve should meet narrow presets exactly. Got " + twoDigit);
        TestAssert.True(Math.Abs(compositeNarrow - 23.95d) < 0.01d, "Narrow-tick titled columns should not overshoot the preset. Got " + compositeNarrow);
    }

    public static void PptxSyntheticStackedValueAxisReserveKeepsSinglePlotRegime()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "UseMeasuredStackedValueAxisReserve",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected stacked value-reserve regime gate to remain inspectable by the Office evidence guard.");

        TestAssert.True((bool)method!.Invoke(null, [1])!, "Single-plot stacked charts should use the measured indent-plus-font-gap rule (column-stacked Office origin within 0.1pt).");
        TestAssert.True(!(bool)method.Invoke(null, [2])!, "Multi-plot stacked charts should keep the shared estimator (compact probe Office strip 17.6pt vs 24.1pt measured).");
    }

    public static void PptxSyntheticPieManualLeaderPicksSmallestXFactor()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "SelectPieManualLeaderLabelIndex",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie manual-leader pick helper to remain inspectable by the Office evidence guard.");

        int? probePick = (int?)method!.Invoke(null, [new List<(int, double)> { (0, -0.77024), (1, 0.66015), (2, -0.07016), (3, 0.61922) }]);
        TestAssert.True(probePick == 2, "Leader probe should pick Gamma (smallest |x-factor|). Got " + probePick);
        int? offsetPick = (int?)method.Invoke(null, [new List<(int, double)> { (0, -0.77342), (1, 0.43820), (2, 0.07619), (3, 0.63844) }]);
        TestAssert.True(offsetPick == 2, "Offset probe should pick West (smallest |x-factor|). Got " + offsetPick);
        int? legacyPick = (int?)method.Invoke(null, [new List<(int, double)>()]);
        TestAssert.True(legacyPick is null, "Charts without valid manual labels should keep legacy emission. Got " + legacyPick);
    }

    public static void PptxSyntheticPieManualLeaderRejectsOutOfRangeFactors()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "TryGetPieManualLeaderFactorX",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie manual-leader validity helper to remain inspectable by the Office evidence guard.");

        PptxSceneChartManualLayout FactorLayout(double? x, double? y) => new(
            HasLayout: true, X: x, XValue: "", Y: y, YValue: "",
            Width: null, WidthValue: "", Height: null, HeightValue: "",
            LayoutTargetKind: PptxSceneChartManualLayoutTarget.Unknown, LayoutTarget: "",
            XModeKind: PptxSceneChartManualLayoutMode.Unknown, XMode: "",
            YModeKind: PptxSceneChartManualLayoutMode.Unknown, YMode: "",
            WidthModeKind: PptxSceneChartManualLayoutMode.Unknown, WidthMode: "",
            HeightModeKind: PptxSceneChartManualLayoutMode.Unknown, HeightMode: "");

        object?[] inRange = [FactorLayout(0.07, 0.54), 0d];
        TestAssert.True(((bool?)method!.Invoke(null, inRange) ?? false) && Math.Abs((double)(inRange[1] ?? 0d) - 0.07) < 0.000001d, "In-range factor manuals should stay eligible with their x factor.");
        object?[] wideY = [FactorLayout(0.03, 1.93), 0d];
        TestAssert.True(!(bool)method.Invoke(null, wideY)!, "Out-of-range y factor should drop the manual layout (ladder Delta falls back to auto).");
        object?[] wideX = [FactorLayout(1.5, 0.2), 0d];
        TestAssert.True(!(bool)method.Invoke(null, wideX)!, "Out-of-range x factor should drop the manual layout.");
        PptxSceneChartManualLayout edge = FactorLayout(0.07, 0.54) with { XModeKind = PptxSceneChartManualLayoutMode.Edge };
        object?[] edgeArgs = [edge, 0d];
        TestAssert.True(!(bool)method.Invoke(null, edgeArgs)!, "Edge-mode manuals carry points, not factors, and stay out of the pick.");
    }

    public static void PptxSyntheticPieManualAnchorSitsOutsideRimMidpoint()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var layoutType = renderer.Assembly.GetType("Lokad.OoxPdf.Pptx.PptxSceneChartManualLayout") ?? throw new InvalidOperationException("Expected manual layout.");
        var method = renderer.GetMethod("ResolvePieManualDataLabelBox", flags);
        TestAssert.True(method is not null, "Expected pie manual-anchor helper to remain inspectable by the Office evidence guard.");
        var targetType = renderer.Assembly.GetType("Lokad.OoxPdf.Pptx.PptxSceneChartManualLayoutTarget")!;
        var modeType = renderer.Assembly.GetType("Lokad.OoxPdf.Pptx.PptxSceneChartManualLayoutMode")!;
        object Layout(double? x, double? y)
        {
            return System.Activator.CreateInstance(layoutType, [true, x, "", y, "", null, "", null, "", System.Enum.ToObject(targetType, 0), "", System.Enum.ToObject(modeType, 0), "", System.Enum.ToObject(modeType, 0), "", System.Enum.ToObject(modeType, 0), "", System.Enum.ToObject(modeType, 0), ""])!;
        }
        var plotBoxType = renderer.GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected plot box.");
        object plotBox = System.Activator.CreateInstance(plotBoxType, [144d, 120d, 520d, 360d])!;
        object right = method!.Invoke(null, [plotBox, Layout(-0.77024, -0.36990), 549.7d, 309.2d, 404d, 80d, 24.3d, 18d])!;
        object left = method.Invoke(null, [plotBox, Layout(0.66015, -0.73512), 325.4d, 177d, 404d, 80d, 24.3d, 18d])!;
        (double rx, double ry) = ReadLayoutBoxXY(right);
        (double lx, double ly) = ReadLayoutBoxXY(left);
        TestAssert.True(Math.Abs(rx - 158.2d) < 0.5d && Math.Abs(ry - 430.2d) < 0.5d, "Right-side anchor reads away from the slice. Got " + rx + "/" + ry);
        TestAssert.True(Math.Abs(lx - 579.7d) < 0.5d && Math.Abs(ly - 429.5d) < 0.5d, "Left-side anchor reads toward the slice past the full width. Got " + lx + "/" + ly);
    }

    private static (double X, double Y) ReadLayoutBoxXY(object box)
    {
        double x = (double)box.GetType().GetProperty("X")!.GetValue(box)!;
        double y = (double)box.GetType().GetProperty("Y")!.GetValue(box)!;
        return (x, y);
    }

    public static void PptxSyntheticPieLabeledLayoutShrinksAndCenters()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var method = renderer.GetMethod("ResolvePieOrDoughnutLayout", flags);
        TestAssert.True(method is not null, "Expected polar layout helper to remain inspectable by the Office evidence guard.");
        var kindType = renderer.GetNestedType("ChartPolarKind", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected polar kind.");
        var plotBoxType = renderer.GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected plot box.");
        var legendType = renderer.GetNestedType("ChartLegendLayout", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected legend layout.");
        object pie = System.Enum.ToObject(kindType, 0);
        object plotBox = System.Activator.CreateInstance(plotBoxType, [144d, 120d, 520d, 360d])!;
        object hidden = legendType.GetProperty("Hidden", flags | System.Reflection.BindingFlags.Public)!.GetValue(null)!;
        var empty = new Dictionary<int, double>();
        object unlabeled = method!.Invoke(null, [pie, plotBox, empty, hidden, false])!;
        object labeled = method.Invoke(null, [pie, plotBox, empty, hidden, true])!;
        (double ux, double uy, double ur) = ReadPolarGeometry(unlabeled);
        (double lx, double ly, double lr) = ReadPolarGeometry(labeled);
        TestAssert.True(Math.Abs(ux - 404d) < 0.01d && Math.Abs(uy - 284.88d) < 0.01d && Math.Abs(ur - 156.24d) < 0.01d, "Unlabeled pies keep the tall radius (5-categories port). Got " + ux + "/" + uy + "/" + ur);
        TestAssert.True(Math.Abs(lx - 404d) < 0.01d && Math.Abs(ly - 300d) < 0.01d && Math.Abs(lr - 145.44d) < 0.01d, "Labeled pies center with the smaller radius (leader probes). Got " + lx + "/" + ly + "/" + lr);
    }

    public static void PptxSyntheticPieAutoLabelConstantsKeepOfficeCalibration()
    {
        var rules = typeof(PptxRenderer).GetNestedType("PptxChartMetricRules", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected metric rules.");
        double radius = (double)rules.GetField("PieAutoDataLabelRadiusRatio")!.GetValue(null)!;
        double baseline = (double)rules.GetField("PieDataLabelBaselineFactor")!.GetValue(null)!;
        TestAssert.True(Math.Abs(radius - 0.74d) < 0.000001d, "Pie auto radius should stay at the 9-sample Office mean. Got " + radius);
        TestAssert.True(Math.Abs(baseline - 0.88d) < 0.000001d, "Pie baseline inset should stay at the measured ascent. Got " + baseline);
    }

    private static (double CenterX, double CenterY, double Radius) ReadPolarGeometry(object layout)
    {
        object geometry = layout.GetType().GetProperty("Geometry")!.GetValue(layout)!;
        double cx = (double)geometry.GetType().GetProperty("CenterX")!.GetValue(geometry)!;
        double cy = (double)geometry.GetType().GetProperty("CenterY")!.GetValue(geometry)!;
        double radius = (double)geometry.GetType().GetProperty("Radius")!.GetValue(geometry)!;
        return (cx, cy, radius);
    }

    public static void PptxSyntheticChartMeasuredRightReserveKeepsPresetFloor()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ResolveMeasuredRightReserve",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart right-reserve helper to remain inspectable by the Office evidence guard.");

        double wideLabels = (double)method!.Invoke(null, [10.94d, 27.37d])!;
        double narrowLabels = (double)method.Invoke(null, [20.16d, 18.25d])!;

        TestAssert.True(Math.Abs(wideLabels - 24.69d) < 0.01d, "Wide labels should beat the preset by half width plus tail. Got " + wideLabels);
        TestAssert.True(Math.Abs(narrowLabels - 20.16d) < 0.01d, "Fitting labels must keep the preset floor. Got " + narrowLabels);
    }

    public static void PptxSyntheticChartPieLabelAnglesMirrorSlices()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var startMethod = renderer.GetMethod("GetPieDataLabelStartAngle", flags);
        var midMethod = renderer.GetMethod("GetPieDataLabelMidpointAngle", flags);
        TestAssert.True(startMethod is not null && midMethod is not null, "Expected pie label angle helpers to remain inspectable by the Office evidence guard.");

        double start = (double)startMethod!.Invoke(null, [0d])!;
        double mid = (double)midMethod!.Invoke(null, [start, 48d, 100d])!;

        TestAssert.True(Math.Abs(start - Math.PI / 2d) < 0.000001d, "Labels should start at the top like slices. Got " + start);
        TestAssert.True(Math.Abs(mid - 0.02d * Math.PI) < 0.000001d, "48-share label mid should follow start-minus-half-share. Got " + mid);
    }

    public static void PptxSyntheticChartPieRadiusBaseUsesPlotHeight()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var method = renderer.GetMethod("GetPieOrDoughnutRadiusBase", flags);
        TestAssert.True(method is not null, "Expected polar radius helper to remain inspectable by the Office evidence guard.");
        var kindType = renderer.GetNestedType("ChartPolarKind", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected polar kind.");
        object pie = System.Enum.ToObject(kindType, 0);
        object doughnut = System.Enum.ToObject(kindType, 1);

        double piePortrait = (double)method!.Invoke(null, [pie, 396d, 432d])!;
        double pieLandscape = (double)method.Invoke(null, [pie, 720d, 432d])!;
        double doughnutPortrait = (double)method.Invoke(null, [doughnut, 396d, 432d])!;

        TestAssert.True(Math.Abs(piePortrait - 432d) < 0.000001d, "Portrait pies should keep the height-based radius. Got " + piePortrait);
        TestAssert.True(Math.Abs(pieLandscape - 432d) < 0.000001d, "Landscape pies are unchanged. Got " + pieLandscape);
        TestAssert.True(Math.Abs(doughnutPortrait - 396d) < 0.000001d, "Doughnuts keep the smaller side without portrait evidence. Got " + doughnutPortrait);
    }

    public static void PptxSyntheticChartMeasuredLeftInsetKeepsPresetFloor()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ResolveMeasuredLeftInset",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart left-inset helper to remain inspectable by the Office evidence guard.");

        double measured = (double)method!.Invoke(null, [25.3d, 27.37d])!;
        double presetWins = (double)method.Invoke(null, [50.61d, 27.4d])!;

        TestAssert.True(Math.Abs(measured - 50.57d) < 0.0001d, "Wide labels should beat the preset by the Office gap. Got " + measured);
        TestAssert.True(Math.Abs(presetWins - 50.61d) < 0.0001d, "Fitting labels must keep the preset floor. Got " + presetWins);
    }

    public static void PptxSyntheticChartPieArcKeepsMathConvention()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "AppendCircularArc",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie arc helper to remain inspectable by the Office evidence guard.");

        var graphics = new PdfGraphicsBuilder();
        method!.Invoke(null, [graphics, 100d, 200d, 50d, System.Math.PI / 2d, -System.Math.PI, true]);
        string path = graphics.ToString();
        var points = new System.Collections.Generic.List<(double X, double Y)>();
        var pending = new System.Collections.Generic.List<double>();
        foreach (string token in path.Split(new[] { (char)32, (char)10, (char)13, (char)9 }, System.StringSplitOptions.RemoveEmptyEntries))
        {
            if (double.TryParse(token, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double value))
            {
                pending.Add(value);
                continue;
            }
            if ((token == "m" || token == "l") && pending.Count >= 2)
            {
                points.Add((pending[pending.Count - 2], pending[pending.Count - 1]));
            }
            if (token == "c" && pending.Count >= 6)
            {
                points.Add((pending[pending.Count - 2], pending[pending.Count - 1]));
            }
            pending.Clear();
        }

        TestAssert.True(points.Count > 0, "Expected emitted arc points, got: " + path);
        foreach ((double X, double Y) in points)
        {
            double radius = System.Math.Sqrt((X - 100d) * (X - 100d) + (Y - 200d) * (Y - 200d));
            TestAssert.True(System.Math.Abs(radius - 50d) < 0.6d, "Arc endpoint should stay on the circle, got radius " + radius);
        }
        (double firstX, double firstY) = points[0];
        TestAssert.True(System.Math.Abs(firstX - 100d) < 0.6d && System.Math.Abs(firstY - 250d) < 0.6d, "Semicircle should start at the top, got " + firstX + "," + firstY);
        (double lastX, double lastY) = points[points.Count - 1];
        TestAssert.True(System.Math.Abs(lastX - 100d) < 0.6d && System.Math.Abs(lastY - 150d) < 0.6d, "Clockwise semicircle should end at the bottom, got " + lastX + "," + lastY);
        double maxX = points.Max(p => p.X);
        TestAssert.True(System.Math.Abs(maxX - 150d) < 0.6d, "Clockwise semicircle should pass through the right point, got maxX " + maxX);
    }

    public static void PptxSyntheticChartValueGridlinesExcludeCrossingTick()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling><c:crosses val="max"/><c:majorUnit val="10"/><c:majorGridlines/>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        Match group = Regex.Match(pdf, @"(?:[0-9.]+ [0-9.]+ m\s+[0-9.]+ [0-9.]+ l\s+){3}S");
        TestAssert.True(group.Success, "Expected three value gridline segments after excluding the max crossing tick.");
        double[] groupY = Regex.Matches(group.Value, @"[0-9.]+ (?<y>[0-9.]+) m\s+[0-9.]+ [0-9.]+ l")
            .Select(match => double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture))
            .ToArray();
        double categoryAxisY = Regex.Matches(pdf, @"[0-9.]+ (?<y>[0-9.]+) m [0-9.]+ (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Where(match => match.Groups["y"].Value == match.Groups["y2"].Value)
            .Select(match => double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture))
            .Max();
        TestAssert.True(groupY.All(y => Math.Abs(y - categoryAxisY) > 0.001d),
            "When c:crosses is max, the grouped gridlines should exclude the top crossing tick.");
        TestAssert.True(groupY.Min() < categoryAxisY,
            "When c:crosses is max, lower value ticks remain visible as gridlines.");
    }

    public static void PptxSyntheticChartCategoryAxisLineUsesCrossingTick()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling><c:crosses val="max"/>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var horizontalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.Y1 - line.Y2) < 0.001d && Math.Abs(line.X1 - line.X2) > 100d)
            .ToArray();
        var verticalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.X1 - line.X2) < 0.001d && Math.Abs(line.Y1 - line.Y2) > 100d)
            .ToArray();

        TestAssert.True(horizontalLines.Length >= 1, "Expected a category-axis horizontal stroke.");
        TestAssert.True(verticalLines.Length >= 1, "Expected a value-axis vertical stroke.");
        double categoryAxisY = horizontalLines.Max(line => line.Y1);
        double valueAxisTopY = verticalLines.Max(line => Math.Max(line.Y1, line.Y2));
        TestAssert.True(Math.Abs(categoryAxisY - valueAxisTopY) < 0.001d,
            "When c:crosses is max, the category-axis line should be placed at the value-axis maximum.");
    }

    public static void PptxSyntheticChartValueAxisStrokeUsesAxisPosition()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="r"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var horizontalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.Y1 - line.Y2) < 0.001d && Math.Abs(line.X1 - line.X2) > 100d)
            .ToArray();
        var verticalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.X1 - line.X2) < 0.001d && Math.Abs(line.Y1 - line.Y2) > 100d)
            .ToArray();

        TestAssert.True(horizontalLines.Length >= 1, "Expected a category-axis horizontal stroke.");
        TestAssert.True(verticalLines.Length >= 1, "Expected a value-axis vertical stroke.");
        double categoryAxisLeftX = horizontalLines.Min(line => Math.Min(line.X1, line.X2));
        double categoryAxisRightX = horizontalLines.Max(line => Math.Max(line.X1, line.X2));
        double valueAxisX = verticalLines.Max(line => line.X1);
        TestAssert.True(Math.Abs(valueAxisX - categoryAxisRightX) < 0.001d,
            "A value axis with c:axPos val=\"r\" should emit its vertical stroke on the plot area's right edge.");
        TestAssert.True(Math.Abs(valueAxisX - categoryAxisLeftX) > 100d,
            "The value-axis stroke should not remain on the left edge when c:axPos requests the right side.");
    }

    public static void PptxSyntheticChartSecondaryValueAxisStrokeUsesOwnAxisPosition()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>20</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="3"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="r"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  <c:valAx>
                    <c:axId val="3"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling>
                    <c:spPr><a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln></c:spPr>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var horizontalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.Y1 - line.Y2) < 0.001d && Math.Abs(line.X1 - line.X2) > 100d)
            .ToArray();
        var verticalLineXs = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.X1 - line.X2) < 0.001d && Math.Abs(line.Y1 - line.Y2) > 100d)
            .Select(line => line.X1)
            .Distinct()
            .OrderBy(x => x)
            .ToArray();

        TestAssert.True(horizontalLines.Length >= 1, "Expected a category-axis horizontal stroke.");
        double categoryAxisLeftX = horizontalLines.Min(line => Math.Min(line.X1, line.X2));
        double categoryAxisRightX = horizontalLines.Max(line => Math.Max(line.X1, line.X2));
        TestAssert.True(verticalLineXs.Any(x => Math.Abs(x - categoryAxisLeftX) < 0.001d),
            "The secondary value axis should use its own c:axPos val=\"l\" instead of being forced to the right side.");
        TestAssert.True(verticalLineXs.Any(x => Math.Abs(x - categoryAxisRightX) < 0.001d),
            "The primary value axis should still use c:axPos val=\"r\".");
    }

    public static void PptxSyntheticChartVisibleSecondaryAxisReservesInnerPlotStrips()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/><c:grouping val="stacked"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>130</c:v></c:pt><c:pt idx="1"><c:v>0</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:barChart>
                    <c:barDir val="col"/><c:grouping val="stacked"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="3"/><c:axId val="4"/>
                  </c:barChart>
                  <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="2"/></c:catAx>
                  <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="low"/><c:scaling><c:min val="0"/><c:max val="200"/></c:scaling><c:majorUnit val="20"/><c:crossAx val="1"/></c:valAx>
                  <c:catAx><c:axId val="3"/><c:delete val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="4"/></c:catAx>
                  <c:valAx><c:axId val="4"/><c:axPos val="r"/><c:tickLblPos val="low"/><c:scaling><c:min val="0"/><c:max val="20"/></c:scaling><c:majorUnit val="2"/><c:crossAx val="3"/></c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var horizontalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.Y1 - line.Y2) < 0.001d && Math.Abs(line.X1 - line.X2) > 100d)
            .ToArray();

        TestAssert.True(horizontalLines.Length >= 1, "Expected a category-axis horizontal stroke.");
        double categoryAxisLeftX = horizontalLines.Min(line => Math.Min(line.X1, line.X2));
        TestAssert.True(categoryAxisLeftX > 105d,
            "Visible primary value-axis labels should reserve an inner plot strip instead of starting at the outer overlay plot box.");
    }

    public static void PptxSyntheticChartComboLineUsesOwnValueAxisScale()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:lineChart>
                    <c:ser>
                      <c:spPr><a:ln><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>50</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:axId val="1"/><c:axId val="3"/>
                  </c:lineChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="10"/></c:scaling>
                  </c:valAx>
                  <c:valAx>
                    <c:axId val="3"/><c:axPos val="r"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="100"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var allLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .ToArray();
        var verticalLines = allLines
            .Where(line => Math.Abs(line.X1 - line.X2) < 0.001d && Math.Abs(line.Y1 - line.Y2) > 100d)
            .ToArray();
        MatchCollection blueLineMatches = Regex.Matches(
            pdf,
            @"0 0 1 RG\s+[0-9.]+ w\s+\[\] 0 d\s+0 J\s+0 j[\s\S]*?(?<x1>[0-9.]+) (?<y1>[0-9.]+) m\s+(?<x2>[0-9.]+) (?<y2>[0-9.]+) l\s+S");
        var blueLine = blueLineMatches
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .FirstOrDefault(line => Math.Abs(line.X1 - line.X2) > 100d);

        TestAssert.True(verticalLines.Length >= 1, "Expected value-axis strokes to establish the plot height.");
        TestAssert.True(blueLine is not null, "Expected the combo line chart to emit its blue line series.");
        double plotBottomY = verticalLines.Min(line => Math.Min(line.Y1, line.Y2));
        double plotTopY = verticalLines.Max(line => Math.Max(line.Y1, line.Y2));
        double plotMidY = (plotBottomY + plotTopY) / 2d;
        double lineSecondY = blueLine!.X2 > blueLine.X1 ? blueLine.Y2 : blueLine.Y1;
        TestAssert.True(Math.Abs(lineSecondY - plotMidY) < 1d,
            "The combo line point with value 50 should use the line chart's 0..100 secondary value-axis scale.");
        TestAssert.True(Math.Abs(lineSecondY - plotTopY) > 20d,
            "The combo line point should not be clamped to the top of the primary 0..10 value-axis scale.");
    }

    public static void PptxSyntheticChartHorizontalBarCategoryAxisStrokeUsesAxisPosition()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="bar"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="r"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var horizontalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.Y1 - line.Y2) < 0.001d && Math.Abs(line.X1 - line.X2) > 100d)
            .ToArray();
        var verticalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.X1 - line.X2) < 0.001d && Math.Abs(line.Y1 - line.Y2) > 100d)
            .ToArray();

        TestAssert.True(horizontalLines.Length >= 1, "Expected a horizontal value-axis stroke.");
        TestAssert.True(verticalLines.Length >= 1, "Expected a vertical category-axis stroke.");
        double valueAxisLeftX = horizontalLines.Min(line => Math.Min(line.X1, line.X2));
        double valueAxisRightX = horizontalLines.Max(line => Math.Max(line.X1, line.X2));
        double categoryAxisX = verticalLines.Max(line => line.X1);
        TestAssert.True(Math.Abs(categoryAxisX - valueAxisRightX) < 0.001d,
            "A horizontal bar category axis with c:axPos val=\"r\" should emit its vertical stroke on the plot area's right edge.");
        TestAssert.True(Math.Abs(categoryAxisX - valueAxisLeftX) > 100d,
            "The horizontal bar category-axis stroke should not remain on the left edge when c:axPos requests the right side.");
    }

    public static void PptxSyntheticChartHorizontalBarValueAxisStrokeUsesAxisPosition()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="bar"/>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var horizontalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.Y1 - line.Y2) < 0.001d && Math.Abs(line.X1 - line.X2) > 100d)
            .ToArray();
        var verticalLines = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S")
            .Cast<Match>()
            .Select(match => new
            {
                X1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)
            })
            .Where(line => Math.Abs(line.X1 - line.X2) < 0.001d && Math.Abs(line.Y1 - line.Y2) > 100d)
            .ToArray();

        TestAssert.True(horizontalLines.Length >= 1, "Expected a horizontal value-axis stroke.");
        TestAssert.True(verticalLines.Length >= 1, "Expected a vertical category-axis stroke.");
        double categoryAxisTopY = verticalLines.Max(line => Math.Max(line.Y1, line.Y2));
        double categoryAxisBottomY = verticalLines.Min(line => Math.Min(line.Y1, line.Y2));
        double valueAxisY = horizontalLines.Min(line => line.Y1);
        TestAssert.True(Math.Abs(valueAxisY - categoryAxisBottomY) < 0.001d,
            "A horizontal bar value axis with c:axPos val=\"b\" should emit its horizontal stroke on the plot area's bottom edge.");
        TestAssert.True(Math.Abs(valueAxisY - categoryAxisTopY) > 100d,
            "The horizontal bar value-axis stroke should not remain on the top edge when c:axPos requests the bottom side.");
    }

    public static void PptxSyntheticChartValueAxisReversedOrientationInvertsLineGeometry()
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
                  <c:chart><c:plotArea>
                  <c:lineChart>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:lineChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:orientation val="maxMin"/><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        Match diagonalLine = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m\s+(?<x2>[0-9.]+) (?<y2>[0-9.]+) l\s+S")
            .Cast<Match>()
            .FirstOrDefault(match =>
            {
                double x1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture);
                double y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture);
                double x2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture);
                double y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture);
                return Math.Abs(x2 - x1) > 100d && Math.Abs(y2 - y1) > 100d;
            }) ?? Match.Empty;
        TestAssert.True(diagonalLine.Success, "Expected a diagonal two-point line-series stroke.");
        double firstY = double.Parse(diagonalLine.Groups["y1"].Value, CultureInfo.InvariantCulture);
        double secondY = double.Parse(diagonalLine.Groups["y2"].Value, CultureInfo.InvariantCulture);
        TestAssert.True(firstY > secondY,
            "With c:scaling/c:orientation maxMin, the lower value should plot above the higher value.");
    }

    public static void PptxSyntheticChartValueAxisReversedOrientationMovesLineDataLabels()
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
                  <c:chart><c:plotArea>
                  <c:lineChart>
                    <c:ser>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:dLbls><c:showVal val="1"/><c:dLblPos val="t"/></c:dLbls>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:lineChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:orientation val="maxMin"/><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var labelMatrices = Regex.Matches(pdf, @"1 0 0 1 (?<x>[0-9.]+) (?<y>[0-9.]+) Tm")
            .Cast<Match>()
            .Select(match => new
            {
                X = double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Y = double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture)
            })
            .OrderBy(matrix => matrix.X)
            .ToArray();
        TestAssert.True(labelMatrices.Length >= 2, "Expected two line data-label text matrices.");
        TestAssert.True(labelMatrices[0].Y > labelMatrices[1].Y,
            "Line data-label anchors should follow maxMin value-axis orientation.");
    }

    public static void PptxSyntheticChartValueAxisReversedOrientationMovesBarDataLabels()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:ser>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:dLbls><c:showVal val="1"/><c:dLblPos val="ctr"/></c:dLbls>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:orientation val="maxMin"/><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var labelMatrices = Regex.Matches(pdf, @"1 0 0 1 (?<x>[0-9.]+) (?<y>[0-9.]+) Tm")
            .Cast<Match>()
            .Select(match => new
            {
                X = double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Y = double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture)
            })
            .OrderBy(matrix => matrix.X)
            .ToArray();
        TestAssert.True(labelMatrices.Length >= 2, "Expected two bar data-label text matrices.");
        TestAssert.True(labelMatrices[0].Y > labelMatrices[1].Y,
            "Bar data-label anchors should follow maxMin value-axis orientation.");
    }

    public static void PptxSyntheticChartValueAxisReversedOrientationMovesHorizontalBarsAndLabels()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="bar"/>
                    <c:ser>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:dLbls><c:showVal val="1"/><c:dLblPos val="ctr"/></c:dLbls>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                    <c:scaling><c:orientation val="maxMin"/><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var barRectangles = Regex.Matches(pdf, @"(?<x>[0-9.]+) (?<y>[0-9.]+) (?<w>[0-9.]+) (?<h>[0-9.]+) re\s+f")
            .Cast<Match>()
            .Select(match => new
            {
                X = double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Width = double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture),
                Height = double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture)
            })
            .Where(rectangle => rectangle.Width > 50d && rectangle.Height > 5d && rectangle.Height < 80d)
            .OrderBy(rectangle => rectangle.Width)
            .ToArray();
        TestAssert.True(barRectangles.Length >= 2, "Expected two filled horizontal bar rectangles.");
        TestAssert.True(barRectangles[^1].X < barRectangles[0].X,
            "The max value horizontal bar should extend left of the smaller value when the value axis is maxMin.");

        var labelMatrices = Regex.Matches(pdf, @"1 0 0 1 (?<x>[0-9.]+) (?<y>[0-9.]+) Tm")
            .Cast<Match>()
            .Select(match => new
            {
                X = double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Y = double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture)
            })
            .ToArray();
        TestAssert.True(labelMatrices.Length >= 2, "Expected two horizontal bar data-label text matrices.");
        TestAssert.True(labelMatrices[0].X > labelMatrices[1].X,
            "Horizontal bar geometry and data-label anchors should follow maxMin value-axis orientation.");
    }

    public static void PptxSyntheticChartValueAxisReversedOrientationStacksColumnSegmentsByValue()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:grouping val="stacked"/>
                    <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>20</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:orientation val="maxMin"/><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var segmentRectangles = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m\s+(?<x2>[0-9.]+) (?<y2>[0-9.]+) l\s+(?<x3>[0-9.]+) (?<y3>[0-9.]+) l\s+(?<x4>[0-9.]+) (?<y4>[0-9.]+) l\s+h\s+f")
            .Cast<Match>()
            .Select(match =>
            {
                double x1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture);
                double x2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture);
                double x3 = double.Parse(match.Groups["x3"].Value, CultureInfo.InvariantCulture);
                double x4 = double.Parse(match.Groups["x4"].Value, CultureInfo.InvariantCulture);
                double y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture);
                double y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture);
                double y3 = double.Parse(match.Groups["y3"].Value, CultureInfo.InvariantCulture);
                double y4 = double.Parse(match.Groups["y4"].Value, CultureInfo.InvariantCulture);
                return new
                {
                    Y = new[] { y1, y2, y3, y4 }.Min(),
                    Width = new[] { x1, x2, x3, x4 }.Max() - new[] { x1, x2, x3, x4 }.Min(),
                    Height = new[] { y1, y2, y3, y4 }.Max() - new[] { y1, y2, y3, y4 }.Min()
                };
            })
            .Where(rectangle => rectangle.Width > 10d && rectangle.Width < 220d && rectangle.Height > 20d && rectangle.Height < 220d)
            .ToArray();
        TestAssert.True(segmentRectangles.Length >= 2, "Expected two filled stacked column segment paths.");
        TestAssert.True(segmentRectangles[1].Y < segmentRectangles[0].Y,
            "Stacked column segments should accumulate through mapped value intervals when the value axis is maxMin.");
    }

    public static void PptxSyntheticChartStackedColumnDataLabelsUseStackSegmentGeometry()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="col"/>
                    <c:grouping val="stacked"/>
                    <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>20</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:dLbls><c:showVal val="1"/></c:dLbls>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                    <c:scaling><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var labelMatrices = Regex.Matches(pdf, @"1 0 0 1 (?<x>[0-9.]+) (?<y>[0-9.]+) Tm")
            .Cast<Match>()
            .Select(match => new
            {
                X = double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Y = double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture)
            })
            .OrderBy(matrix => matrix.Y)
            .ToArray();
        TestAssert.Equal(2, labelMatrices.Length);
        TestAssert.True(Math.Abs(labelMatrices[0].X - labelMatrices[1].X) < 1d,
            "Stacked column data labels should share the stacked segment column instead of using clustered series slots.");
        TestAssert.True(labelMatrices[1].Y > labelMatrices[0].Y,
            "Stacked column data labels should follow cumulative segment geometry.");
    }

    public static void PptxSyntheticChartValueAxisReversedOrientationStacksHorizontalBarSegmentsByValue()
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
                  <c:chart><c:plotArea>
                  <c:barChart>
                    <c:barDir val="bar"/>
                    <c:grouping val="stacked"/>
                    <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>20</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:axId val="1"/><c:axId val="2"/>
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1"/><c:axPos val="l"/><c:tickLblPos val="none"/>
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2"/><c:axPos val="b"/><c:tickLblPos val="none"/>
                    <c:scaling><c:orientation val="maxMin"/><c:min val="0"/><c:max val="30"/></c:scaling>
                  </c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var segmentRectangles = Regex.Matches(pdf, @"(?<x1>[0-9.]+) (?<y1>[0-9.]+) m\s+(?<x2>[0-9.]+) (?<y2>[0-9.]+) l\s+(?<x3>[0-9.]+) (?<y3>[0-9.]+) l\s+(?<x4>[0-9.]+) (?<y4>[0-9.]+) l\s+h\s+f")
            .Cast<Match>()
            .Select(match =>
            {
                double x1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture);
                double x2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture);
                double x3 = double.Parse(match.Groups["x3"].Value, CultureInfo.InvariantCulture);
                double x4 = double.Parse(match.Groups["x4"].Value, CultureInfo.InvariantCulture);
                double y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture);
                double y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture);
                double y3 = double.Parse(match.Groups["y3"].Value, CultureInfo.InvariantCulture);
                double y4 = double.Parse(match.Groups["y4"].Value, CultureInfo.InvariantCulture);
                return new
                {
                    X = new[] { x1, x2, x3, x4 }.Min(),
                    Width = new[] { x1, x2, x3, x4 }.Max() - new[] { x1, x2, x3, x4 }.Min(),
                    Height = new[] { y1, y2, y3, y4 }.Max() - new[] { y1, y2, y3, y4 }.Min()
                };
            })
            .Where(rectangle => rectangle.Width > 20d && rectangle.Width < 220d && rectangle.Height > 10d && rectangle.Height < 220d)
            .ToArray();
        TestAssert.True(segmentRectangles.Length >= 2, "Expected two filled stacked horizontal bar segment paths.");
        TestAssert.True(segmentRectangles[1].X < segmentRectangles[0].X,
            "Stacked horizontal bar segments should accumulate through mapped value intervals when the value axis is maxMin.");
    }

    public static void PptxSyntheticChartManualLayoutEdgeModesRender()
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
                  <c:chart><c:plotArea>
                    <c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="edge"/><c:hMode val="edge"/><c:x val="0.2"/><c:y val="0.1"/><c:w val="0.7"/><c:h val="0.6"/></c:manualLayout></c:layout>
                    <c:spPr><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill></c:spPr>
                    <c:barChart>
                      <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                    </c:barChart>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 1 1 rg", pdf);
        TestAssert.Contains("129.6 338.4 144 108 re f", pdf);
        TestAssert.DoesNotContain("129.6 252 144 216 re f", pdf);
    }

    public static void PptxSyntheticChartManualLayoutFactorPositionUsesDefaultPlotBox()
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
                  <c:chart><c:plotArea>
                    <c:layout><c:manualLayout><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0"/><c:y val="0"/><c:w val="0.5"/><c:h val="0.5"/></c:manualLayout></c:layout>
                    <c:spPr><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill></c:spPr>
                    <c:lineChart>
                      <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                    </c:lineChart>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("106.56 325.44 144 108 re f", pdf);
        TestAssert.DoesNotContain("72 360 144 108 re f", pdf);
    }

    public static void PptxSyntheticChartManualLayoutUnknownTargetUsesDefaultPlotBox()
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
                  <c:chart><c:plotArea>
                    <c:layout><c:manualLayout><c:layoutTarget val="bogus"/><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0"/><c:y val="0"/><c:w val="0.5"/><c:h val="0.5"/></c:manualLayout></c:layout>
                    <c:spPr><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill></c:spPr>
                    <c:barChart>
                      <c:barDir val="bar"/>
                      <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                    </c:barChart>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("88.589 352.008 144 108 re f", pdf);
        TestAssert.DoesNotContain("123.552", pdf);
    }

    public static void PptxSyntheticChartTitleManualBoxDrivesPlacement()
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
                  <c:chart>
                    <c:title>
                      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Manual Title</a:t></a:r></a:p></c:rich></c:tx>
                      <c:layout><c:manualLayout><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0.25"/><c:y val="0.5"/><c:w val="0.5"/><c:h val="0.2"/></c:manualLayout></c:layout>
                        <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
                    </c:title>
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                      </c:lineChart>
                      <c:valAx><c:axId val="2"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Axis Manual</a:t></a:r></a:p></c:rich></c:tx><c:layout><c:manualLayout><c:x val="0.05"/><c:y val="0.15"/><c:w val="0.35"/><c:h val="0.12"/></c:manualLayout></c:layout><c:spPr><a:solidFill><a:srgbClr val="BADA55"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="5533AA"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="442288"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:title></c:valAx>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.729 0.855 0.333 rg", pdf);
        TestAssert.Contains("0.333 0.2 0.667 RG", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CAT[0-9]+ 9 Tf"), "Expected explicit manual-layout axis title txPr to drive axis-title rendering.");
        TestAssert.True(Regex.IsMatch(pdf, @"1 0 0 1 [0-9.]+ 354\.816 Tm"), "Expected explicit chart title manualLayout to drive the title text baseline.");
        TestAssert.True(!Regex.IsMatch(pdf, @"1 0 0 1 [0-9.]+ 442\.[0-9]+ Tm"), "Expected chart title rendering not to fall back to the full-frame title box.");
    }

    public static void PptxSyntheticChartDefaultAxisTitleRendersWithoutUnsupportedDiagnostic()
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
                  <c:chart>
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                      </c:lineChart>
                      <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Default Axis</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_CHART_AXIS_TITLE_LAYOUT"),
            "Supported native chart branches should render default-placement axis titles instead of reporting the old unsupported-layout diagnostic.");
        TestAssert.Contains("/CAT", pdf);
    }

    public static void PptxSyntheticChartDefaultAxisTitleMissingPositionEmitsTargetedDiagnostic()
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
                  <c:chart>
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                      </c:lineChart>
                      <c:valAx><c:axId val="2"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Default Axis</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(diagnostics.Any(d =>
            d.Id == "PPTX_UNSUPPORTED_CHART_AXIS_TITLE_AXIS_POSITION" &&
            d.PartName == "/ppt/charts/chart1.xml" &&
            d.Fallback == "Ignored"),
            "A default axis title with an unknown axis position should be targeted diagnostic-covered.");
        TestAssert.True(!diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_CHART_AXIS_TITLE_LAYOUT"),
            "The old blanket default-layout diagnostic should not return for supported chart branches.");
        TestAssert.DoesNotContain("/CAT", pdf);
    }

    public static void PptxSyntheticTransparentChartTitleUsesGlyphOutlinePaths()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

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
                  <c:chart>
                    <c:title>
                      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Alpha Title</a:t></a:r></a:p></c:rich></c:tx>
                      <c:layout><c:manualLayout><c:x val="0.15"/><c:y val="0.15"/><c:w val="0.7"/><c:h val="0.2"/></c:manualLayout></c:layout>
                      <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1600"><a:solidFill><a:srgbClr val="336699"><a:alpha val="45000"/></a:srgbClr></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
                    </c:title>
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                      </c:lineChart>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("GS45000F100000S", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("f", pdf);
    }

    public static void PptxSyntheticTransparentChartTitleRichRunUsesGlyphOutlinePaths()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

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
                  <c:chart>
                    <c:title>
                      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1600"><a:solidFill><a:srgbClr val="336699"><a:alpha val="35000"/></a:srgbClr></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Run Alpha</a:t></a:r></a:p></c:rich></c:tx>
                      <c:layout><c:manualLayout><c:x val="0.15"/><c:y val="0.15"/><c:w val="0.7"/><c:h val="0.2"/></c:manualLayout></c:layout>
                      <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1600"><a:solidFill><a:srgbClr val="111111"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
                    </c:title>
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                      </c:lineChart>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        OoxPackage package = OoxPackage.Open(input, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneChart chart = scene.Slides[0].SlideNodes.Select(node => node.Chart).First(chartNode => chartNode is not null)!;
        TestAssert.Equal(0.35d, chart.Title.TextRuns[0].TextStyle.Alpha ?? 0d);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("GS35000F100000S", pdf);
        TestAssert.Contains("0.2 0.4 0.6 rg", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("f", pdf);
    }

    public static void PptxSyntheticChartStyleTitleDefaultsDriveChartTitleRendering()
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
                  <c:chart>
                    <c:title>
                      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Styled Title</a:t></a:r></a:p></c:rich></c:tx>
                    </c:title>
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                      </c:lineChart>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/style1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="77">
                  <cs:title>
                    <cs:defRPr sz="1400" b="1" u="sng" strike="sngStrike">
                      <a:solidFill><a:srgbClr val="C86432"/></a:solidFill>
                      <a:latin typeface="Arial"/>
                    </cs:defRPr>
                  </cs:title>
                </cs:style>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.784 0.392 0.196 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CT[0-9]+ 14\.04 Tf"), "Expected chart-style title role font size to drive title rendering when c:title has no direct txPr.");
        int decorationPathCount = PptxTests.CountFilledDecorationPaths(pdf, @"0\.784 0\.392 0\.196");
        TestAssert.True(decorationPathCount >= 2, "Expected chart-style title role underline and strike to emit filled decoration paths through the common text renderer.");
    }

    public static void PptxSyntheticChartStyleTitleAlphaUsesGlyphOutlinePaths()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

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
                  <c:chart>
                    <c:title>
                      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Style Alpha</a:t></a:r></a:p></c:rich></c:tx>
                      <c:layout><c:manualLayout><c:x val="0.15"/><c:y val="0.15"/><c:w val="0.7"/><c:h val="0.2"/></c:manualLayout></c:layout>
                    </c:title>
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                      </c:lineChart>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/style1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="77">
                  <cs:title>
                    <cs:defRPr sz="1600">
                      <a:solidFill><a:srgbClr val="336699"><a:alpha val="45000"/></a:srgbClr></a:solidFill>
                      <a:latin typeface="Arial"/>
                    </cs:defRPr>
                  </cs:title>
                </cs:style>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("GS45000F100000S", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("f", pdf);
    }
}
