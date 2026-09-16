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

    public static void PptxSyntheticClusteredBarWidthHonorsNegativeOverlap()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "GetClusteredBarWidth",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected clustered bar width helper to remain inspectable by the Office evidence guard.");

        // Office composite vectors: barW 23.64 with step 30.0 on band 135.45, gap 219,
        // overlap -27 (573 divisor); zero overlap keeps the legacy divisor exactly.
        double separated = (double)method!.Invoke(null, [135.45d, 3, 219d, -27d])!;
        TestAssert.True(System.Math.Abs(separated - 23.64d) < 0.01d, "Expected separated width 23.64, got " + separated.ToString(System.Globalization.CultureInfo.InvariantCulture));
        double plain = (double)method.Invoke(null, [135.27d, 3, 219d, 0d])!;
        TestAssert.True(System.Math.Abs(plain - 26.0636d) < 0.01d, "Expected legacy width 26.0636, got " + plain.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }
    public static void PptxSyntheticSingleSeriesVaryColorsShadeGate()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ShouldShadeSingleSeriesVaryColors",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected vary-colors shade gate to remain inspectable by the Office evidence guard.");

        // Office count ladder: 4/5pt raw (neg4/neg5/dash5), 6/7pt dark (neg6/dash6/dash7);
        // horizontals keep legacy output (no 6pt horizontal sample).
        TestAssert.Equal(false, (bool)method!.Invoke(null, [4, true])!);
        TestAssert.Equal(false, (bool)method.Invoke(null, [5, true])!);
        TestAssert.Equal(true, (bool)method.Invoke(null, [6, true])!);
        TestAssert.Equal(true, (bool)method.Invoke(null, [7, true])!);
        TestAssert.Equal(false, (bool)method.Invoke(null, [6, false])!);
        TestAssert.Equal(false, (bool)method.Invoke(null, [0, true])!);
    }

    public static void PptxSyntheticFullFrameFillLegendLineHeight()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeFullFrameFillLegendLineHeight",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected fill line height to remain inspectable by the Office evidence guard.");

        // Office overlay pitches 20.51/27.77/35.00 at 12/18/24pt (residuals under 0.025);
        // non-full-frame paths keep their factors.
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [12d, true, false, false])! - 20.51d) < 0.05d, "Expected fs12 fill pitch.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, true, false, false])! - 27.77d) < 0.05d, "Expected fs18 fill pitch.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [24d, true, false, false])! - 35.00d) < 0.05d, "Expected fs24 fill pitch.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, false, false, true])! - 27.78d) < 0.05d, "Expected side-fill legacy pitch.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, false, true, false])! - 27.78d) < 0.05d, "Expected stroke legacy pitch.");
    }

    public static void PptxSyntheticOverlayLegendTextGap()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ResolveSideLegendTextGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected side text gap to remain inspectable by the Office evidence guard.");

        // Overlay side keys take the measured 4.66pt fill-swatch gap like doughnut-right
        // (Office overlay swatch-to-text 4.66 exact); everything else keeps legacy routing.
        var placementType = typeof(PptxRenderer).GetNestedType("ChartLegendPlacement", System.Reflection.BindingFlags.NonPublic);
        object? placement = System.Enum.Parse(placementType!, "Default");
        TestAssert.Equal(4.66d, (double)method!.Invoke(null, [placement, false, false, false, 18d, true])!);
        TestAssert.Equal(3.0d, (double)method.Invoke(null, [placement, false, false, false, 18d, false])!);
    }

    public static void PptxSyntheticOverlayLegendVerticalShift()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeOverlayLegendVerticalShift",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected overlay shift to remain inspectable by the Office evidence guard.");

        // Office middles sit markerSize/2 below frame middle (5.01/5.03 at fs18, 6.64 at fs24).
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [9.9d, true])! - 4.95d) < 0.01d, "Expected fs18 overlay shift.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [13.2d, true])! - 6.6d) < 0.01d, "Expected fs24 overlay shift.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [9.9d, false])!) < 0.01d, "Expected no shift off-overlay.");
    }

    public static void PptxSyntheticHorizontalLegendPackingGaps()
    {
        var rendererType = typeof(PptxRenderer);
        var textGap = rendererType.GetMethod(
            "ComputeHorizontalLegendTextGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        var interGap = rendererType.GetMethod(
            "ComputeHorizontalLegendInterEntryGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(textGap is not null && interGap is not null, "Expected horizontal packing helpers to remain inspectable by the Office evidence guard.");

        // Swatch-text gap: legacy 3.0 off-bottom; content-independent line 0.275fs-0.35
        // (2.96/4.60/6.25 at 12/18/24pt) on bottom legends.
        TestAssert.True(System.Math.Abs((double)textGap!.Invoke(null, [12d, false])! - 3.0d) < 0.01d, "Expected legacy text gap.");
        TestAssert.True(System.Math.Abs((double)textGap.Invoke(null, [12.02d, true])! - 2.96d) < 0.05d, "Expected leg12 text gap.");
        TestAssert.True(System.Math.Abs((double)textGap.Invoke(null, [18d, true])! - 4.60d) < 0.05d, "Expected leg18 text gap.");
        TestAssert.True(System.Math.Abs((double)textGap.Invoke(null, [24d, true])! - 6.25d) < 0.05d, "Expected leg24 text gap.");

        // Inter-entry gap: legacy 8.0 off-bottom; 0.813fs-0.0276range+0.285 on bottom
        // legends (six Office knots within 0.031).
        TestAssert.True(System.Math.Abs((double)interGap!.Invoke(null, [18d, 24.19d, false])! - 8.0d) < 0.01d, "Expected legacy inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [12.02d, 16.18d, true])! - 9.60d) < 0.05d, "Expected leg12 inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [18d, 24.19d, true])! - 14.27d) < 0.05d, "Expected leg18 inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [24d, 32.30d, true])! - 18.90d) < 0.05d, "Expected leg24 inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [18d, 50.56d, true])! - 13.54d) < 0.05d, "Expected longnames inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [18d, 5.80d, true])! - 14.79d) < 0.05d, "Expected nsw inter gap.");
    }

    public static void PptxSyntheticDefaultAxisTitleVerticalBottomReserve()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeDefaultAxisTitleVerticalBottomReserve",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected vertical-bottom composition to remain inspectable by the Office evidence guard.");

        // Additive composition: topright base 46.79, cat-title-18 knot 54.12,
        // cat-label-14 knot 56.03 (Office 46.77/54.09/56.00, all within 0.03).
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [9d, 12d])! - 46.79d) < 0.05d, "Expected base 46.79.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [9d, 18d])! - 54.12d) < 0.05d, "Expected title18 54.12.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [14.04d, 12d])! - 56.03d) < 0.05d, "Expected label14 56.03.");
    }

    public static void PptxSyntheticDefaultAxisTitleHorizontalBarReserves()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeDefaultAxisTitleHorizontalBarReserves",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected horizontal-bar title reserves to remain inspectable by the Office evidence guard.");

        // Office three-knot ladders: top 46.77/52.31/56.00 over 9/12/14.04pt value ticks,
        // bottom 46.37/41.10/37.57 over 21.6/17.28/14.4pt chart titles (residuals under 0.03).
        (double top9, double bottom216) = ((double, double))method!.Invoke(null, [9d, 21.6d])!;
        TestAssert.True(System.Math.Abs(top9 - 46.77d) < 0.05d, "Expected top 46.77, got " + top9.ToString(System.Globalization.CultureInfo.InvariantCulture));
        TestAssert.True(System.Math.Abs(bottom216 - 46.37d) < 0.05d, "Expected bottom 46.37, got " + bottom216.ToString(System.Globalization.CultureInfo.InvariantCulture));
        (double top12, double bottom1728) = ((double, double))method.Invoke(null, [12d, 17.28d])!;
        TestAssert.True(System.Math.Abs(top12 - 52.31d) < 0.05d, "Expected top 52.31, got " + top12.ToString(System.Globalization.CultureInfo.InvariantCulture));
        TestAssert.True(System.Math.Abs(bottom1728 - 41.10d) < 0.05d, "Expected bottom 41.10, got " + bottom1728.ToString(System.Globalization.CultureInfo.InvariantCulture));
        (double top14, double bottom144) = ((double, double))method.Invoke(null, [14.04d, 14.4d])!;
        TestAssert.True(System.Math.Abs(top14 - 56.00d) < 0.05d, "Expected top 56.00, got " + top14.ToString(System.Globalization.CultureInfo.InvariantCulture));
        TestAssert.True(System.Math.Abs(bottom144 - 37.57d) < 0.05d, "Expected bottom 37.57, got " + bottom144.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    public static void PptxSyntheticSingleSeriesVaryColorsShadeFactor()
    {
        var rendererType = typeof(PptxRenderer);
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the shade pin.");
        var method = rendererType.GetMethod(
            "ShadeSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected vary-colors shade helper to remain inspectable by the Office evidence guard.");

        // Office allpositive vectors: accent1 (79,129,189) renders (69,114,167);
        // AwayFromZero rounding lands (70,114,166), inside PDF-quantum noise.
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? shaded = method!.Invoke(null, [accent]);
        TestAssert.Equal((byte)70, (byte)rgbType!.GetProperty("Red")!.GetValue(shaded)!);
        TestAssert.Equal((byte)114, (byte)rgbType.GetProperty("Green")!.GetValue(shaded)!);
        TestAssert.Equal((byte)166, (byte)rgbType.GetProperty("Blue")!.GetValue(shaded)!);
    }

    public static void PptxSyntheticSingleSeriesVaryColorsOverflowSlots7Plus()
    {
        var rendererType = typeof(PptxRenderer);
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the overflow pin.");
        var method = rendererType.GetMethod(
            "TryResolveSingleSeriesVaryColorsOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected vary-colors overflow helper to remain inspectable by the Office evidence guard.");

        // Office dash7 through dash53 fills agree: slot7 light steel
        // (0.576/0.663/0.812 to 147,169,207), slot8 dusty rose (0.82/0.576/0.573
        // to 209,147,146), slot9 light green (0.725/0.804/0.588 to 185,205,150),
        // slot10 lavender (0.663/0.608/0.741 to 169,155,189).
        object?[] slot7 = [6, null];
        TestAssert.Equal(true, (bool)method!.Invoke(null, slot7)!);
        TestAssert.Equal((byte)147, (byte)rgbType!.GetProperty("Red")!.GetValue(slot7[1])!);
        TestAssert.Equal((byte)169, (byte)rgbType.GetProperty("Green")!.GetValue(slot7[1])!);
        TestAssert.Equal((byte)207, (byte)rgbType.GetProperty("Blue")!.GetValue(slot7[1])!);
        object?[] first = [0, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, first)!);
        object?[] sixth = [5, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, sixth)!);
        object?[] eighth = [7, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, eighth)!);
        TestAssert.Equal((byte)209, (byte)rgbType.GetProperty("Red")!.GetValue(eighth[1])!);
        TestAssert.Equal((byte)147, (byte)rgbType.GetProperty("Green")!.GetValue(eighth[1])!);
        TestAssert.Equal((byte)146, (byte)rgbType.GetProperty("Blue")!.GetValue(eighth[1])!);
        object?[] ninth = [8, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, ninth)!);
        TestAssert.Equal((byte)185, (byte)rgbType.GetProperty("Red")!.GetValue(ninth[1])!);
        TestAssert.Equal((byte)205, (byte)rgbType.GetProperty("Green")!.GetValue(ninth[1])!);
        TestAssert.Equal((byte)150, (byte)rgbType.GetProperty("Blue")!.GetValue(ninth[1])!);
        object?[] tenth = [9, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, tenth)!);
        TestAssert.Equal((byte)169, (byte)rgbType.GetProperty("Red")!.GetValue(tenth[1])!);
        TestAssert.Equal((byte)155, (byte)rgbType.GetProperty("Green")!.GetValue(tenth[1])!);
        TestAssert.Equal((byte)189, (byte)rgbType.GetProperty("Blue")!.GetValue(tenth[1])!);
        // slot13 pale teal (0.667/0.729/0.843 to 170,186,215), slot14 dusty pink
        // (0.851/0.667/0.663 to 217,170,169), slot15 mint (0.776/0.839/0.675 to
        // 198,214,172), slot16 periwinkle (0.729/0.69/0.788 to 186,176,201),
        // slot19 pale sky (0.714/0.765/0.863 to 182,195,220),
        // slot20 dusty mauve (0.867/0.714/0.71 to 221,182,181),
        // slot21 pale pistachio (0.804/0.859/0.722 to 205,219,184),
        // slot22 pale lilac (0.765/0.729/0.816 to 195,186,208),
        // slot25 periwinkle-blue (0.737/0.784/0.875 to 188,200,223),
        // slot26 pale rose (0.878/0.737/0.737 to 224,188,188),
        // slot27 pale mint (0.82/0.871/0.745 to 209,222,190),
        // slot28 pale grape (0.784/0.753/0.831 to 200,192,212),
        // slot31 lavender-blue (0.761/0.804/0.882 to 194,205,225),
        // slot32 pale coral (0.886/0.761/0.761 to 226,194,194),
        // slot33 pale spring (0.835/0.878/0.769 to 213,224,196),
        // slot34 pale violet (0.804/0.776/0.843 to 205,198,215),
        // slot37 periwinkle (0.773/0.812/0.886 to 197,207,226),
        // slot38 pale terracotta (0.894/0.773/0.773 to 228,197,197),
        // slot39 pale lime (0.843/0.886/0.78 to 215,226,199),
        // slot40 pale heather (0.812/0.788/0.851 to 207,201,217),
        // slot43 sky-blue (0.784/0.82/0.894 to 200,209,228),
        // slot44 pale salmon (0.898/0.784/0.78 to 229,200,199),
        // slot45 pale apple (0.851/0.89/0.788 to 217,227,201),
        // slot46 pale lilac-gray (0.82/0.796/0.859 to 209,203,219),
        // slot49 steel-blue (0.796/0.831/0.898 to 203,212,229),
        // slot50 pale terracotta-rose (0.902/0.796/0.792 to 230,203,202),
        // slot51 pale honeydew (0.855/0.894/0.8 to 218,228,204),
        // slot52 pale lavender-gray (0.831/0.808/0.863 to 212,206,220);
        // slots 1-6 keep the 0.88 shade, slot-53-plus keeps shaded cycling (slot53 sky single sample).
        object?[] eleventh = [10, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, eleventh)!);
        object?[] thirteenth = [12, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, thirteenth)!);
        TestAssert.Equal((byte)170, (byte)rgbType.GetProperty("Red")!.GetValue(thirteenth[1])!);
        TestAssert.Equal((byte)186, (byte)rgbType.GetProperty("Green")!.GetValue(thirteenth[1])!);
        TestAssert.Equal((byte)215, (byte)rgbType.GetProperty("Blue")!.GetValue(thirteenth[1])!);
        object?[] fourteenth = [13, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fourteenth)!);
        TestAssert.Equal((byte)217, (byte)rgbType.GetProperty("Red")!.GetValue(fourteenth[1])!);
        TestAssert.Equal((byte)170, (byte)rgbType.GetProperty("Green")!.GetValue(fourteenth[1])!);
        TestAssert.Equal((byte)169, (byte)rgbType.GetProperty("Blue")!.GetValue(fourteenth[1])!);
        object?[] fifteenth = [14, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fifteenth)!);
        TestAssert.Equal((byte)198, (byte)rgbType.GetProperty("Red")!.GetValue(fifteenth[1])!);
        TestAssert.Equal((byte)214, (byte)rgbType.GetProperty("Green")!.GetValue(fifteenth[1])!);
        TestAssert.Equal((byte)172, (byte)rgbType.GetProperty("Blue")!.GetValue(fifteenth[1])!);
        object?[] sixteenth = [15, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, sixteenth)!);
        TestAssert.Equal((byte)186, (byte)rgbType.GetProperty("Red")!.GetValue(sixteenth[1])!);
        TestAssert.Equal((byte)176, (byte)rgbType.GetProperty("Green")!.GetValue(sixteenth[1])!);
        TestAssert.Equal((byte)201, (byte)rgbType.GetProperty("Blue")!.GetValue(sixteenth[1])!);
        object?[] seventeenth = [16, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, seventeenth)!);
        object?[] nineteenth = [18, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, nineteenth)!);
        TestAssert.Equal((byte)182, (byte)rgbType.GetProperty("Red")!.GetValue(nineteenth[1])!);
        TestAssert.Equal((byte)195, (byte)rgbType.GetProperty("Green")!.GetValue(nineteenth[1])!);
        TestAssert.Equal((byte)220, (byte)rgbType.GetProperty("Blue")!.GetValue(nineteenth[1])!);
        object?[] twentieth = [19, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, twentieth)!);
        TestAssert.Equal((byte)221, (byte)rgbType.GetProperty("Red")!.GetValue(twentieth[1])!);
        TestAssert.Equal((byte)182, (byte)rgbType.GetProperty("Green")!.GetValue(twentieth[1])!);
        TestAssert.Equal((byte)181, (byte)rgbType.GetProperty("Blue")!.GetValue(twentieth[1])!);
        object?[] twentyFirst = [20, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, twentyFirst)!);
        TestAssert.Equal((byte)205, (byte)rgbType.GetProperty("Red")!.GetValue(twentyFirst[1])!);
        TestAssert.Equal((byte)219, (byte)rgbType.GetProperty("Green")!.GetValue(twentyFirst[1])!);
        TestAssert.Equal((byte)184, (byte)rgbType.GetProperty("Blue")!.GetValue(twentyFirst[1])!);
        object?[] twentySecond = [21, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, twentySecond)!);
        TestAssert.Equal((byte)195, (byte)rgbType.GetProperty("Red")!.GetValue(twentySecond[1])!);
        TestAssert.Equal((byte)186, (byte)rgbType.GetProperty("Green")!.GetValue(twentySecond[1])!);
        TestAssert.Equal((byte)208, (byte)rgbType.GetProperty("Blue")!.GetValue(twentySecond[1])!);
        object?[] twentyThird = [22, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, twentyThird)!);
        object?[] twentyFifth = [24, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, twentyFifth)!);
        TestAssert.Equal((byte)188, (byte)rgbType.GetProperty("Red")!.GetValue(twentyFifth[1])!);
        TestAssert.Equal((byte)200, (byte)rgbType.GetProperty("Green")!.GetValue(twentyFifth[1])!);
        TestAssert.Equal((byte)223, (byte)rgbType.GetProperty("Blue")!.GetValue(twentyFifth[1])!);
        object?[] twentySixth = [25, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, twentySixth)!);
        TestAssert.Equal((byte)224, (byte)rgbType.GetProperty("Red")!.GetValue(twentySixth[1])!);
        TestAssert.Equal((byte)188, (byte)rgbType.GetProperty("Green")!.GetValue(twentySixth[1])!);
        TestAssert.Equal((byte)188, (byte)rgbType.GetProperty("Blue")!.GetValue(twentySixth[1])!);
        object?[] twentySeventh = [26, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, twentySeventh)!);
        TestAssert.Equal((byte)209, (byte)rgbType.GetProperty("Red")!.GetValue(twentySeventh[1])!);
        TestAssert.Equal((byte)222, (byte)rgbType.GetProperty("Green")!.GetValue(twentySeventh[1])!);
        TestAssert.Equal((byte)190, (byte)rgbType.GetProperty("Blue")!.GetValue(twentySeventh[1])!);
        object?[] twentyEighth = [27, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, twentyEighth)!);
        TestAssert.Equal((byte)200, (byte)rgbType.GetProperty("Red")!.GetValue(twentyEighth[1])!);
        TestAssert.Equal((byte)192, (byte)rgbType.GetProperty("Green")!.GetValue(twentyEighth[1])!);
        TestAssert.Equal((byte)212, (byte)rgbType.GetProperty("Blue")!.GetValue(twentyEighth[1])!);
        object?[] twentyNinth = [28, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, twentyNinth)!);
        object?[] thirtyFirst = [30, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, thirtyFirst)!);
        TestAssert.Equal((byte)194, (byte)rgbType.GetProperty("Red")!.GetValue(thirtyFirst[1])!);
        TestAssert.Equal((byte)205, (byte)rgbType.GetProperty("Green")!.GetValue(thirtyFirst[1])!);
        TestAssert.Equal((byte)225, (byte)rgbType.GetProperty("Blue")!.GetValue(thirtyFirst[1])!);
        object?[] thirtySecond = [31, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, thirtySecond)!);
        TestAssert.Equal((byte)226, (byte)rgbType.GetProperty("Red")!.GetValue(thirtySecond[1])!);
        TestAssert.Equal((byte)194, (byte)rgbType.GetProperty("Green")!.GetValue(thirtySecond[1])!);
        TestAssert.Equal((byte)194, (byte)rgbType.GetProperty("Blue")!.GetValue(thirtySecond[1])!);
        object?[] thirtyThird = [32, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, thirtyThird)!);
        TestAssert.Equal((byte)213, (byte)rgbType.GetProperty("Red")!.GetValue(thirtyThird[1])!);
        TestAssert.Equal((byte)224, (byte)rgbType.GetProperty("Green")!.GetValue(thirtyThird[1])!);
        TestAssert.Equal((byte)196, (byte)rgbType.GetProperty("Blue")!.GetValue(thirtyThird[1])!);
        object?[] thirtyFourth = [33, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, thirtyFourth)!);
        TestAssert.Equal((byte)205, (byte)rgbType.GetProperty("Red")!.GetValue(thirtyFourth[1])!);
        TestAssert.Equal((byte)198, (byte)rgbType.GetProperty("Green")!.GetValue(thirtyFourth[1])!);
        TestAssert.Equal((byte)215, (byte)rgbType.GetProperty("Blue")!.GetValue(thirtyFourth[1])!);
        object?[] thirtyFifth = [34, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, thirtyFifth)!);
        object?[] thirtySeventh = [36, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, thirtySeventh)!);
        TestAssert.Equal((byte)197, (byte)rgbType.GetProperty("Red")!.GetValue(thirtySeventh[1])!);
        TestAssert.Equal((byte)207, (byte)rgbType.GetProperty("Green")!.GetValue(thirtySeventh[1])!);
        TestAssert.Equal((byte)226, (byte)rgbType.GetProperty("Blue")!.GetValue(thirtySeventh[1])!);
        object?[] thirtyEighth = [37, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, thirtyEighth)!);
        TestAssert.Equal((byte)228, (byte)rgbType.GetProperty("Red")!.GetValue(thirtyEighth[1])!);
        TestAssert.Equal((byte)197, (byte)rgbType.GetProperty("Green")!.GetValue(thirtyEighth[1])!);
        TestAssert.Equal((byte)197, (byte)rgbType.GetProperty("Blue")!.GetValue(thirtyEighth[1])!);
        object?[] thirtyNinth = [38, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, thirtyNinth)!);
        TestAssert.Equal((byte)215, (byte)rgbType.GetProperty("Red")!.GetValue(thirtyNinth[1])!);
        TestAssert.Equal((byte)226, (byte)rgbType.GetProperty("Green")!.GetValue(thirtyNinth[1])!);
        TestAssert.Equal((byte)199, (byte)rgbType.GetProperty("Blue")!.GetValue(thirtyNinth[1])!);
        object?[] fortieth = [39, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fortieth)!);
        TestAssert.Equal((byte)207, (byte)rgbType.GetProperty("Red")!.GetValue(fortieth[1])!);
        TestAssert.Equal((byte)201, (byte)rgbType.GetProperty("Green")!.GetValue(fortieth[1])!);
        TestAssert.Equal((byte)217, (byte)rgbType.GetProperty("Blue")!.GetValue(fortieth[1])!);
        object?[] fortyFirst = [40, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, fortyFirst)!);
        object?[] fortyThird = [42, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fortyThird)!);
        TestAssert.Equal((byte)200, (byte)rgbType.GetProperty("Red")!.GetValue(fortyThird[1])!);
        TestAssert.Equal((byte)209, (byte)rgbType.GetProperty("Green")!.GetValue(fortyThird[1])!);
        TestAssert.Equal((byte)228, (byte)rgbType.GetProperty("Blue")!.GetValue(fortyThird[1])!);
        object?[] fortyFourth = [43, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fortyFourth)!);
        TestAssert.Equal((byte)229, (byte)rgbType.GetProperty("Red")!.GetValue(fortyFourth[1])!);
        TestAssert.Equal((byte)200, (byte)rgbType.GetProperty("Green")!.GetValue(fortyFourth[1])!);
        TestAssert.Equal((byte)199, (byte)rgbType.GetProperty("Blue")!.GetValue(fortyFourth[1])!);
        object?[] fortyFifth = [44, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fortyFifth)!);
        TestAssert.Equal((byte)217, (byte)rgbType.GetProperty("Red")!.GetValue(fortyFifth[1])!);
        TestAssert.Equal((byte)227, (byte)rgbType.GetProperty("Green")!.GetValue(fortyFifth[1])!);
        TestAssert.Equal((byte)201, (byte)rgbType.GetProperty("Blue")!.GetValue(fortyFifth[1])!);
        object?[] fortySixth = [45, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fortySixth)!);
        TestAssert.Equal((byte)209, (byte)rgbType.GetProperty("Red")!.GetValue(fortySixth[1])!);
        TestAssert.Equal((byte)203, (byte)rgbType.GetProperty("Green")!.GetValue(fortySixth[1])!);
        TestAssert.Equal((byte)219, (byte)rgbType.GetProperty("Blue")!.GetValue(fortySixth[1])!);
        object?[] fortySeventh = [46, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, fortySeventh)!);
        object?[] fortyNinth = [48, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fortyNinth)!);
        TestAssert.Equal((byte)203, (byte)rgbType.GetProperty("Red")!.GetValue(fortyNinth[1])!);
        TestAssert.Equal((byte)212, (byte)rgbType.GetProperty("Green")!.GetValue(fortyNinth[1])!);
        TestAssert.Equal((byte)229, (byte)rgbType.GetProperty("Blue")!.GetValue(fortyNinth[1])!);
        object?[] fiftieth = [49, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fiftieth)!);
        TestAssert.Equal((byte)230, (byte)rgbType.GetProperty("Red")!.GetValue(fiftieth[1])!);
        TestAssert.Equal((byte)203, (byte)rgbType.GetProperty("Green")!.GetValue(fiftieth[1])!);
        TestAssert.Equal((byte)202, (byte)rgbType.GetProperty("Blue")!.GetValue(fiftieth[1])!);
        object?[] fiftyFirst = [50, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fiftyFirst)!);
        TestAssert.Equal((byte)218, (byte)rgbType.GetProperty("Red")!.GetValue(fiftyFirst[1])!);
        TestAssert.Equal((byte)228, (byte)rgbType.GetProperty("Green")!.GetValue(fiftyFirst[1])!);
        TestAssert.Equal((byte)204, (byte)rgbType.GetProperty("Blue")!.GetValue(fiftyFirst[1])!);
        object?[] fiftySecond = [51, null];
        TestAssert.Equal(true, (bool)method.Invoke(null, fiftySecond)!);
        TestAssert.Equal((byte)212, (byte)rgbType.GetProperty("Red")!.GetValue(fiftySecond[1])!);
        TestAssert.Equal((byte)206, (byte)rgbType.GetProperty("Green")!.GetValue(fiftySecond[1])!);
        TestAssert.Equal((byte)220, (byte)rgbType.GetProperty("Blue")!.GetValue(fiftySecond[1])!);
        object?[] fiftyThird = [52, null];
        TestAssert.Equal(false, (bool)method.Invoke(null, fiftyThird)!);
    }

    public static void PptxSyntheticThirdVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseThirdVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected third-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var dark = rendererType.GetMethod(
            "ShadeThirdRegimeDarkSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected third-regime dark shade to remain inspectable by the Office evidence guard.");
        var mid = rendererType.GetMethod(
            "ShadeThirdRegimeMidSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(mid is not null, "Expected third-regime mid shade to remain inspectable by the Office evidence guard.");
        var light = rendererType.GetMethod(
            "TryResolveThirdRegimeLightFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(light is not null, "Expected third-regime light table to remain inspectable by the Office evidence guard.");

        // Seventeen points and fewer keep earlier regimes; eighteen-plus take dark/mid/light rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [17])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [18])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [19])!);
        // Office dash18/dash19 dark row at 0.78 (maxabs 2): accent1 (79,129,189) renders (62,101,147).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? darked = dark!.Invoke(null, [accent]);
        TestAssert.Equal((byte)62, (byte)rgbType!.GetProperty("Red")!.GetValue(darked)!);
        TestAssert.Equal((byte)101, (byte)rgbType.GetProperty("Green")!.GetValue(darked)!);
        TestAssert.Equal((byte)147, (byte)rgbType.GetProperty("Blue")!.GetValue(darked)!);
        // Mid row at 0.93 (maxabs 1): accent1 renders (73,120,176).
        object? midded = mid!.Invoke(null, [accent]);
        TestAssert.Equal((byte)73, (byte)rgbType.GetProperty("Red")!.GetValue(midded)!);
        TestAssert.Equal((byte)120, (byte)rgbType.GetProperty("Green")!.GetValue(midded)!);
        TestAssert.Equal((byte)176, (byte)rgbType.GetProperty("Blue")!.GetValue(midded)!);
        // Light row fixed table, both renders agree on all six.
        int[] wantR = [126, 202, 174, 155, 124, 248];
        int[] wantG = [155, 126, 198, 137, 187, 170];
        int[] wantB = [200, 125, 131, 179, 207, 121];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [12 + slot, null];
            TestAssert.Equal(true, (bool)light!.Invoke(null, args)!);
            TestAssert.Equal((byte)wantR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)wantG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)wantB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        object?[] past = [18, null];
        TestAssert.Equal(false, (bool)light!.Invoke(null, past)!);
    }
    public static void PptxSyntheticFourthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseFourthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected fourth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var dark = rendererType.GetMethod(
            "TryResolveFourthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected fourth-regime dark row to remain inspectable by the Office evidence guard.");
        var replay = rendererType.GetMethod(
            "TryResolveFourthRegimeReplayFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(replay is not null, "Expected fourth-regime replay row to remain inspectable by the Office evidence guard.");

        // Twenty-three points and fewer keep earlier regimes; twenty-four-plus take dark/0.88/raw/replay rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [23])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [24])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [25])!);
        // Office dash24/dash25 dark row, both renders agree on all six vectors.
        int[] darkR = [57, 144, 116, 95, 54, 186];
        int[] darkG = [96, 58, 140, 73, 129, 112];
        int[] darkB = [142, 56, 65, 121, 149, 50];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType!.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        object?[] seventh = [6, null];
        TestAssert.Equal(false, (bool)dark.Invoke(null, seventh)!);
        // Replay row: first-regime fixed tail byte-exact across both renders, slots 19-24.
        int[] reR = [147, 209, 185, 169, 145, 249];
        int[] reG = [169, 147, 205, 155, 195, 181];
        int[] reB = [207, 146, 150, 189, 213, 144];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [18 + slot, null];
            TestAssert.Equal(true, (bool)replay!.Invoke(null, args)!);
            TestAssert.Equal((byte)reR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)reG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)reB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        object?[] past = [24, null];
        TestAssert.Equal(false, (bool)replay.Invoke(null, past)!);
    }
    public static void PptxSyntheticFifthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseFifthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected fifth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var mid = rendererType.GetMethod(
            "ShadeFifthRegimeMidSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(mid is not null, "Expected fifth-regime mid shade to remain inspectable by the Office evidence guard.");
        var light = rendererType.GetMethod(
            "ShadeFifthRegimeLightSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(light is not null, "Expected fifth-regime light shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveFifthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected fifth-regime dark table to remain inspectable by the Office evidence guard.");
        var palest = rendererType.GetMethod(
            "TryResolveFifthRegimePalestFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(palest is not null, "Expected fifth-regime palest table to remain inspectable by the Office evidence guard.");

        // Twenty-nine points and fewer keep earlier regimes; thirty-plus take dark/0.85/0.95/fixed/palest rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [29])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [30])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [31])!);
        // Office dash30/dash31 mid row at 0.85 (maxabs 2): accent1 (79,129,189) renders (67,110,161).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? midded = mid!.Invoke(null, [accent]);
        TestAssert.Equal((byte)67, (byte)rgbType!.GetProperty("Red")!.GetValue(midded)!);
        TestAssert.Equal((byte)110, (byte)rgbType.GetProperty("Green")!.GetValue(midded)!);
        TestAssert.Equal((byte)161, (byte)rgbType.GetProperty("Blue")!.GetValue(midded)!);
        // Light row at 0.95 (maxabs 1): accent1 renders (75,123,180).
        object? lighted = light!.Invoke(null, [accent]);
        TestAssert.Equal((byte)75, (byte)rgbType.GetProperty("Red")!.GetValue(lighted)!);
        TestAssert.Equal((byte)123, (byte)rgbType.GetProperty("Green")!.GetValue(lighted)!);
        TestAssert.Equal((byte)180, (byte)rgbType.GetProperty("Blue")!.GetValue(lighted)!);
        // Dark row fixed table, both renders agree on all six.
        int[] darkR = [56, 140, 113, 92, 53, 182];
        int[] darkG = [93, 56, 137, 71, 125, 109];
        int[] darkB = [138, 54, 63, 118, 145, 49];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Palest row fixed table, both renders agree on all six (slots 25-30).
        int[] paleR = [161, 214, 192, 179, 160, 249];
        int[] paleG = [180, 161, 210, 168, 202, 190];
        int[] paleB = [212, 160, 164, 196, 217, 158];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [24 + slot, null];
            TestAssert.Equal(true, (bool)palest!.Invoke(null, args)!);
            TestAssert.Equal((byte)paleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)paleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)paleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        object?[] past = [30, null];
        TestAssert.Equal(false, (bool)palest.Invoke(null, past)!);
    }
    public static void PptxSyntheticSixthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseSixthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected sixth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var light = rendererType.GetMethod(
            "ShadeSixthRegimeLightSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(light is not null, "Expected sixth-regime light shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveSixthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected sixth-regime dark table to remain inspectable by the Office evidence guard.");
        var replay = rendererType.GetMethod(
            "TryResolveSixthRegimeReplayFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(replay is not null, "Expected sixth-regime replay table to remain inspectable by the Office evidence guard.");

        // Thirty-five points and fewer keep earlier regimes; thirty-six-plus take dark/0.82/0.91/raw/fixed/replay rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [35])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [36])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [37])!);
        // Office dash36/dash37 light row at 0.91 (maxabs 1): accent1 (79,129,189) renders (72,117,172).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? lighted = light!.Invoke(null, [accent]);
        TestAssert.Equal((byte)72, (byte)rgbType!.GetProperty("Red")!.GetValue(lighted)!);
        TestAssert.Equal((byte)117, (byte)rgbType.GetProperty("Green")!.GetValue(lighted)!);
        TestAssert.Equal((byte)172, (byte)rgbType.GetProperty("Blue")!.GetValue(lighted)!);
        // Dark row fixed table, both renders agree on all six.
        int[] darkR = [54, 136, 109, 90, 51, 177];
        int[] darkG = [90, 55, 133, 69, 122, 106];
        int[] darkB = [134, 52, 61, 114, 141, 47];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Replay row: regime-2 fixed tail byte-exact across both renders, slots 31-36.
        int[] reR = [170, 217, 198, 186, 169, 250];
        int[] reG = [186, 170, 214, 176, 206, 195];
        int[] reB = [215, 169, 172, 201, 220, 168];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [30 + slot, null];
            TestAssert.Equal(true, (bool)replay!.Invoke(null, args)!);
            TestAssert.Equal((byte)reR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)reG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)reB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        object?[] past = [36, null];
        TestAssert.Equal(false, (bool)replay.Invoke(null, past)!);
    }
    public static void PptxSyntheticSeventhVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseSeventhVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected seventh-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var light = rendererType.GetMethod(
            "ShadeSeventhRegimeLightSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(light is not null, "Expected seventh-regime light shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveSeventhRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected seventh-regime dark table to remain inspectable by the Office evidence guard.");
        var mid = rendererType.GetMethod(
            "TryResolveSeventhRegimeMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(mid is not null, "Expected seventh-regime mid table to remain inspectable by the Office evidence guard.");
        var palest = rendererType.GetMethod(
            "TryResolveSeventhRegimePalestFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(palest is not null, "Expected seventh-regime palest table to remain inspectable by the Office evidence guard.");

        // Forty-one points and fewer keep earlier regimes; forty-two-plus take dark/mid/0.88/0.96/fixed/replay/palest rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [41])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [42])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [43])!);
        // Seventh-regime light shade at 0.96 (maxabs 1 over Office dash42/dash43): accent1 (79,129,189) computes (76,124,181) against observed (76,124,182).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? lighted = light!.Invoke(null, [accent]);
        TestAssert.Equal((byte)76, (byte)rgbType!.GetProperty("Red")!.GetValue(lighted)!);
        TestAssert.Equal((byte)124, (byte)rgbType.GetProperty("Green")!.GetValue(lighted)!);
        TestAssert.Equal((byte)181, (byte)rgbType.GetProperty("Blue")!.GetValue(lighted)!);
        // Dark row fixed table, both renders agree on all six.
        int[] darkR = [53, 134, 107, 88, 50, 173];
        int[] darkG = [89, 53, 130, 68, 119, 104];
        int[] darkB = [132, 51, 60, 112, 138, 46];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Mid row fixed table, both renders agree on all six (slots 7-12).
        int[] midR = [62, 154, 124, 102, 58, 198];
        int[] midG = [102, 62, 150, 79, 137, 119];
        int[] midB = [151, 60, 70, 129, 159, 54];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [6 + slot, null];
            TestAssert.Equal(true, (bool)mid!.Invoke(null, args)!);
            TestAssert.Equal((byte)midR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)midG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)midB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Palest row fixed table, both renders agree on all six (slots 37-42).
        int[] paleR = [175, 219, 201, 190, 174, 250];
        int[] paleG = [190, 175, 216, 180, 209, 199];
        int[] paleB = [217, 175, 177, 204, 222, 173];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [36 + slot, null];
            TestAssert.Equal(true, (bool)palest!.Invoke(null, args)!);
            TestAssert.Equal((byte)paleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)paleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)paleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        object?[] past = [42, null];
        TestAssert.Equal(false, (bool)palest.Invoke(null, past)!);
    }
    public static void PptxSyntheticEighthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseEighthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected eighth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var mid = rendererType.GetMethod(
            "ShadeEighthRegimeMidSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(mid is not null, "Expected eighth-regime mid shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveEighthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected eighth-regime dark table to remain inspectable by the Office evidence guard.");
        var palest = rendererType.GetMethod(
            "TryResolveEighthRegimePalestFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(palest is not null, "Expected eighth-regime palest table to remain inspectable by the Office evidence guard.");

        // Forty-seven points and fewer keep earlier regimes; forty-eight-plus take dark/0.78/0.86/0.93/raw/light/fixed/palest rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [47])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [48])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [49])!);
        // Office dash48/dash49 mid row at 0.86 (maxabs 2): accent1 (79,129,189) computes (68,111,163) against observed (67,111,163).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? midded = mid!.Invoke(null, [accent]);
        TestAssert.Equal((byte)68, (byte)rgbType!.GetProperty("Red")!.GetValue(midded)!);
        TestAssert.Equal((byte)111, (byte)rgbType.GetProperty("Green")!.GetValue(midded)!);
        TestAssert.Equal((byte)163, (byte)rgbType.GetProperty("Blue")!.GetValue(midded)!);
        // Dark row fixed table, both renders agree on all six.
        int[] darkR = [52, 132, 106, 87, 49, 171];
        int[] darkG = [88, 53, 129, 67, 118, 102];
        int[] darkB = [130, 51, 59, 111, 137, 46];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Palest row fixed table, both renders agree on all six (slots 43-48).
        int[] paleR = [182, 221, 205, 195, 181, 250];
        int[] paleG = [195, 182, 219, 186, 212, 203];
        int[] paleB = [220, 181, 184, 208, 224, 180];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [42 + slot, null];
            TestAssert.Equal(true, (bool)palest!.Invoke(null, args)!);
            TestAssert.Equal((byte)paleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)paleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)paleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        object?[] past = [48, null];
        TestAssert.Equal(false, (bool)palest.Invoke(null, past)!);
    }

    public static void PptxSyntheticNinthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseNinthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected ninth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var second = rendererType.GetMethod(
            "ShadeNinthRegimeSecondRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(second is not null, "Expected ninth-regime second-row shade to remain inspectable by the Office evidence guard.");
        var third = rendererType.GetMethod(
            "ShadeNinthRegimeThirdRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(third is not null, "Expected ninth-regime third-row shade to remain inspectable by the Office evidence guard.");
        var fourth = rendererType.GetMethod(
            "ShadeNinthRegimeFourthRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourth is not null, "Expected ninth-regime fourth-row shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveNinthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected ninth-regime dark table to remain inspectable by the Office evidence guard.");
        var ninth = rendererType.GetMethod(
            "TryResolveNinthRegimeNinthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninth is not null, "Expected ninth-regime ninth table to remain inspectable by the Office evidence guard.");

        // Fifty-three points and fewer keep earlier regimes; fifty-four-plus take dark/0.76/0.84/0.90/0.96/fixed/fixed/palest/ninth rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [53])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [54])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [55])!);
        // Computed shade pins on accent1 (79,129,189): 0.76 gives (60,98,144), 0.84 gives (66,108,159), 0.90 gives (71,116,170).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? shadedSecond = second!.Invoke(null, [accent]);
        TestAssert.Equal((byte)60, (byte)rgbType!.GetProperty("Red")!.GetValue(shadedSecond)!);
        TestAssert.Equal((byte)98, (byte)rgbType.GetProperty("Green")!.GetValue(shadedSecond)!);
        TestAssert.Equal((byte)144, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedSecond)!);
        object? shadedThird = third!.Invoke(null, [accent]);
        TestAssert.Equal((byte)66, (byte)rgbType.GetProperty("Red")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)108, (byte)rgbType.GetProperty("Green")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)159, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedThird)!);
        object? shadedFourth = fourth!.Invoke(null, [accent]);
        TestAssert.Equal((byte)71, (byte)rgbType.GetProperty("Red")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)116, (byte)rgbType.GetProperty("Green")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)170, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedFourth)!);
        // Dark row fixed table, both renders agree on all six.
        int[] darkR = [51, 130, 104, 85, 48, 168];
        int[] darkG = [86, 51, 126, 65, 116, 100];
        int[] darkB = [127, 49, 58, 109, 134, 45];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Ninth row fixed table, both renders agree on all six (slots 49-54).
        int[] ninthR = [185, 223, 207, 197, 184, 251];
        int[] ninthG = [198, 185, 220, 189, 214, 205];
        int[] ninthB = [221, 184, 187, 210, 225, 183];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [48 + slot, null];
            TestAssert.Equal(true, (bool)ninth!.Invoke(null, args)!);
            TestAssert.Equal((byte)ninthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)ninthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)ninthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // +6 replay row fixed table: slots 55-58 replay eighth-regime overflow singles at idx48-51 within maxabs 2 (5/4/3/2 samples, bit-identical across dash55-dash59).
        var replay = rendererType.GetMethod(
            "TryResolveNinthRegimeReplayFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(replay is not null, "Expected ninth-regime replay table to remain inspectable by the Office evidence guard.");
        int[] replayR = [204, 231, 219, 213];
        int[] replayG = [213, 204, 229, 207];
        int[] replayB = [230, 204, 205, 221];
        for (int slot = 0; slot < 4; slot++)
        {
            object?[] replayArgs = [54 + slot, null];
            TestAssert.Equal(true, (bool)replay!.Invoke(null, replayArgs)!);
            TestAssert.Equal((byte)replayR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(replayArgs[1])!);
            TestAssert.Equal((byte)replayG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(replayArgs[1])!);
            TestAssert.Equal((byte)replayB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(replayArgs[1])!);
        }
        object?[] replayPast = [58, null];
        TestAssert.Equal(false, (bool)replay.Invoke(null, replayPast)!);
        object?[] past = [54, null];
        TestAssert.Equal(false, (bool)ninth.Invoke(null, past)!);
    }
    public static void PptxSyntheticTenthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTenthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected tenth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var third = rendererType.GetMethod(
            "ShadeSecondRegimeSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(third is not null, "Expected tenth-regime third-row 0.82 shade to remain inspectable by the Office evidence guard.");
        var fourth = rendererType.GetMethod(
            "ShadeSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourth is not null, "Expected tenth-regime fourth-row 0.88 shade to remain inspectable by the Office evidence guard.");
        var fifth = rendererType.GetMethod(
            "ShadeTenthRegimeFifthRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifth is not null, "Expected tenth-regime fifth-row 0.94 shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveTenthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected tenth-regime dark table to remain inspectable by the Office evidence guard.");
        var fixed37 = rendererType.GetMethod(
            "TryResolveTenthRegimeFixedFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed37 is not null, "Expected tenth-regime fixed table to remain inspectable by the Office evidence guard.");
        var fourthReplay = rendererType.GetMethod(
            "TryResolveTenthRegimeFourthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourthReplay is not null, "Expected tenth-regime fourth replay to remain inspectable by the Office evidence guard.");
        var overflow = rendererType.GetMethod(
            "TryResolveTenthRegimeOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflow is not null, "Expected tenth-regime overflow replay to remain inspectable by the Office evidence guard.");

        // Fifty-nine points and fewer keep earlier regimes; sixty-plus take dark/fourth-replay/0.82/0.88/0.94/raw/fixed/seventh-replay/sixth-replay/overflow-replay rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [59])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [60])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [61])!);
        // Computed shade pins on accent1 (79,129,189): 0.82 gives (65,106,155), 0.88 gives (70,114,166), 0.94 gives (74,121,178).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? shadedThird = third!.Invoke(null, [accent]);
        TestAssert.Equal((byte)65, (byte)rgbType!.GetProperty("Red")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)106, (byte)rgbType.GetProperty("Green")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)155, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedThird)!);
        object? shadedFourth = fourth!.Invoke(null, [accent]);
        TestAssert.Equal((byte)70, (byte)rgbType.GetProperty("Red")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)114, (byte)rgbType.GetProperty("Green")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)166, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedFourth)!);
        object? shadedFifth = fifth!.Invoke(null, [accent]);
        TestAssert.Equal((byte)74, (byte)rgbType.GetProperty("Red")!.GetValue(shadedFifth)!);
        TestAssert.Equal((byte)121, (byte)rgbType.GetProperty("Green")!.GetValue(shadedFifth)!);
        TestAssert.Equal((byte)178, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedFifth)!);
        // Dark row fixed table, both renders agree on all six.
        int[] darkR = [50, 128, 103, 84, 47, 166];
        int[] darkG = [85, 51, 125, 65, 114, 99];
        int[] darkB = [126, 49, 57, 107, 132, 44];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Fixed row (slots 37-42), both renders agree on all six.
        int[] fixedR = [118, 200, 170, 149, 115, 248];
        int[] fixedG = [150, 118, 196, 130, 184, 166];
        int[] fixedB = [198, 116, 123, 176, 205, 113];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [36 + slot, null];
            TestAssert.Equal(true, (bool)fixed37!.Invoke(null, args)!);
            TestAssert.Equal((byte)fixedR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)fixedG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)fixedB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Fourth-regime replay (slots 7-12), both renders agree on all six.
        int[] fourthR = [57, 144, 116, 95, 54, 186];
        int[] fourthG = [96, 58, 140, 73, 129, 112];
        int[] fourthB = [142, 56, 65, 121, 149, 50];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [6 + slot, null];
            TestAssert.Equal(true, (bool)fourthReplay!.Invoke(null, args)!);
            TestAssert.Equal((byte)fourthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)fourthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)fourthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Overflow replay covers slots 55-58; fixed tail covers slots 59-64 (3/2/2/2/2/2 samples, bit-identical across dash60-dash65).
        int[] tailR = [187, 251, 205, 231, 220, 214];
        int[] tailG = [215, 207, 214, 205, 230, 208];
        int[] tailB = [227, 186, 230, 205, 207, 222];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] tailArgs = [58 + slot, null];
            TestAssert.Equal(true, (bool)overflow!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] past = [64, null];
        TestAssert.Equal(false, (bool)overflow!.Invoke(null, past)!);
    }
    public static void PptxSyntheticEleventhVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseEleventhVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected eleventh-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var third = rendererType.GetMethod(
            "ShadeEleventhRegimeThirdRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(third is not null, "Expected eleventh-regime third-row 0.80 shade to remain inspectable by the Office evidence guard.");
        var fourth = rendererType.GetMethod(
            "ShadeEleventhRegimeFourthRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourth is not null, "Expected eleventh-regime fourth-row 0.87 shade to remain inspectable by the Office evidence guard.");
        var fifth = rendererType.GetMethod(
            "ShadeEleventhRegimeFifthRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifth is not null, "Expected eleventh-regime fifth-row 0.92 shade to remain inspectable by the Office evidence guard.");
        var sixth = rendererType.GetMethod(
            "ShadeEleventhRegimeSixthRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixth is not null, "Expected eleventh-regime sixth-row 0.97 shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveEleventhRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected eleventh-regime dark table to remain inspectable by the Office evidence guard.");
        var second = rendererType.GetMethod(
            "TryResolveEleventhRegimeSecondFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(second is not null, "Expected eleventh-regime second table to remain inspectable by the Office evidence guard.");
        var tail = rendererType.GetMethod(
            "TryResolveEleventhRegimeFixed61Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected eleventh-regime last fixed table to remain inspectable by the Office evidence guard.");

        // Sixty-five points and fewer keep earlier regimes; sixty-six-plus take dark/second/0.80/0.87/0.92/0.97/fixed rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [65])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [66])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [67])!);
        // Computed shade pins on accent1 (79,129,189): 0.80 gives (63,103,151), 0.87 gives (69,112,164), 0.92 gives (73,119,174), 0.97 gives (77,125,183).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? shadedThird = third!.Invoke(null, [accent]);
        TestAssert.Equal((byte)63, (byte)rgbType!.GetProperty("Red")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)103, (byte)rgbType.GetProperty("Green")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)151, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedThird)!);
        object? shadedFourth = fourth!.Invoke(null, [accent]);
        TestAssert.Equal((byte)69, (byte)rgbType.GetProperty("Red")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)112, (byte)rgbType.GetProperty("Green")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)164, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedFourth)!);
        object? shadedFifth = fifth!.Invoke(null, [accent]);
        TestAssert.Equal((byte)73, (byte)rgbType.GetProperty("Red")!.GetValue(shadedFifth)!);
        TestAssert.Equal((byte)119, (byte)rgbType.GetProperty("Green")!.GetValue(shadedFifth)!);
        TestAssert.Equal((byte)174, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedFifth)!);
        object? shadedSixth = sixth!.Invoke(null, [accent]);
        TestAssert.Equal((byte)77, (byte)rgbType.GetProperty("Red")!.GetValue(shadedSixth)!);
        TestAssert.Equal((byte)125, (byte)rgbType.GetProperty("Green")!.GetValue(shadedSixth)!);
        TestAssert.Equal((byte)183, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedSixth)!);
        // Dark row fixed table, both renders agree on all six.
        int[] darkR = [49, 127, 101, 83, 47, 164];
        int[] darkG = [84, 50, 123, 64, 113, 98];
        int[] darkB = [125, 48, 56, 106, 131, 43];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Second row fixed table (slots 7-12), both renders agree on all six.
        int[] secondR = [56, 142, 114, 93, 53, 183];
        int[] secondG = [94, 57, 138, 72, 126, 110];
        int[] secondB = [139, 55, 64, 119, 146, 49];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] secondArgs = [6 + slot, null];
            TestAssert.Equal(true, (bool)second!.Invoke(null, secondArgs)!);
            TestAssert.Equal((byte)secondR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(secondArgs[1])!);
            TestAssert.Equal((byte)secondG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(secondArgs[1])!);
            TestAssert.Equal((byte)secondB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(secondArgs[1])!);
        }
        // Last fixed row (slots 61-66), both renders agree on all six.
        int[] tailR = [191, 225, 211, 202, 190, 251];
        int[] tailG = [203, 191, 223, 195, 217, 209];
        int[] tailB = [224, 191, 193, 213, 228, 189];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] tailArgs = [60 + slot, null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        // Fixed tail (slots 67-70), second samples bit-identical across dash67-dash71.
        var tail67 = rendererType.GetMethod(
            "TryResolveEleventhRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail67 is not null, "Expected eleventh-regime tail table to remain inspectable by the Office evidence guard.");
        object?[] tail67Args = [66, null];
        TestAssert.Equal(true, (bool)tail67!.Invoke(null, tail67Args)!);
        TestAssert.Equal((byte)207, (byte)rgbType.GetProperty("Red")!.GetValue(tail67Args[1])!);
        TestAssert.Equal((byte)215, (byte)rgbType.GetProperty("Green")!.GetValue(tail67Args[1])!);
        TestAssert.Equal((byte)231, (byte)rgbType.GetProperty("Blue")!.GetValue(tail67Args[1])!);
        object?[] tail68Args = [67, null];
        TestAssert.Equal(true, (bool)tail67!.Invoke(null, tail68Args)!);
        TestAssert.Equal((byte)232, (byte)rgbType.GetProperty("Red")!.GetValue(tail68Args[1])!);
        TestAssert.Equal((byte)207, (byte)rgbType.GetProperty("Green")!.GetValue(tail68Args[1])!);
        TestAssert.Equal((byte)206, (byte)rgbType.GetProperty("Blue")!.GetValue(tail68Args[1])!);
        object?[] tail69Args = [68, null];
        TestAssert.Equal(true, (bool)tail67!.Invoke(null, tail69Args)!);
        TestAssert.Equal((byte)221, (byte)rgbType.GetProperty("Red")!.GetValue(tail69Args[1])!);
        TestAssert.Equal((byte)230, (byte)rgbType.GetProperty("Green")!.GetValue(tail69Args[1])!);
        TestAssert.Equal((byte)208, (byte)rgbType.GetProperty("Blue")!.GetValue(tail69Args[1])!);
        object?[] tail70Args = [69, null];
        TestAssert.Equal(true, (bool)tail67!.Invoke(null, tail70Args)!);
        TestAssert.Equal((byte)215, (byte)rgbType.GetProperty("Red")!.GetValue(tail70Args[1])!);
        TestAssert.Equal((byte)210, (byte)rgbType.GetProperty("Green")!.GetValue(tail70Args[1])!);
        TestAssert.Equal((byte)223, (byte)rgbType.GetProperty("Blue")!.GetValue(tail70Args[1])!);
        object?[] past70 = [70, null];
        TestAssert.Equal(false, (bool)tail67.Invoke(null, past70)!);
        object?[] past = [66, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, past)!);
    }
    public static void PptxSyntheticTwelfthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwelfthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twelfth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var second = rendererType.GetMethod(
            "ShadeTwelfthRegimeSecondRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(second is not null, "Expected twelfth-regime second-row 0.73 shade to remain inspectable by the Office evidence guard.");
        var third = rendererType.GetMethod(
            "ShadeTwelfthRegimeThirdRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(third is not null, "Expected twelfth-regime third-row 0.79 shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveTwelfthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected twelfth-regime dark table to remain inspectable by the Office evidence guard.");
        var fixed49 = rendererType.GetMethod(
            "TryResolveTwelfthRegimeFixed49Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed49 is not null, "Expected twelfth-regime fixed49 table to remain inspectable by the Office evidence guard.");
                var fixed61 = rendererType.GetMethod(
            "TryResolveTwelfthRegimeFixed61Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed61 is not null, "Expected twelfth-regime fixed61 table to remain inspectable by the Office evidence guard.");
var tail = rendererType.GetMethod(
            "TryResolveTwelfthRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected twelfth-regime tail table to remain inspectable by the Office evidence guard.");

        // Seventy-one points and fewer keep earlier regimes; seventy-two-plus take dark/0.73/0.79/0.85/0.90/0.95/raw/fifth/fixed/fifth/fixed/overflow/tail rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [71])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [72])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [73])!);
        // Computed shade pins on accent1 (79,129,189): 0.73 gives (58,94,138), 0.79 gives (62,102,149).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? shadedSecond = second!.Invoke(null, [accent]);
        TestAssert.Equal((byte)58, (byte)rgbType!.GetProperty("Red")!.GetValue(shadedSecond)!);
        TestAssert.Equal((byte)94, (byte)rgbType.GetProperty("Green")!.GetValue(shadedSecond)!);
        TestAssert.Equal((byte)138, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedSecond)!);
        object? shadedThird = third!.Invoke(null, [accent]);
        TestAssert.Equal((byte)62, (byte)rgbType.GetProperty("Red")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)102, (byte)rgbType.GetProperty("Green")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)149, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedThird)!);
        // Dark row replays eleventh-regime slots 1-6 byte-exact across both renders.
        int[] darkR = [49, 127, 101, 83, 47, 164];
        int[] darkG = [84, 50, 123, 64, 113, 98];
        int[] darkB = [125, 48, 56, 106, 131, 43];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Fixed row (slots 49-54), both renders agree on all six.
        int[] fixedR = [140, 207, 181, 164, 139, 249];
        int[] fixedG = [165, 140, 203, 149, 192, 178];
        int[] fixedB = [204, 139, 144, 186, 211, 137];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] fixedArgs = [48 + slot, null];
            TestAssert.Equal(true, (bool)fixed49!.Invoke(null, fixedArgs)!);
            TestAssert.Equal((byte)fixedR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fixedArgs[1])!);
        }
        // Fixed row (slots 61-66), both renders agree on all six.
        int[] fixed61R = [178, 220, 203, 192, 177, 250];
        int[] fixed61G = [193, 179, 218, 184, 210, 201];
        int[] fixed61B = [219, 178, 181, 206, 223, 176];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] fixed61Args = [60 + slot, null];
            TestAssert.Equal(true, (bool)fixed61!.Invoke(null, fixed61Args)!);
            TestAssert.Equal((byte)fixed61R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fixed61Args[1])!);
            TestAssert.Equal((byte)fixed61G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fixed61Args[1])!);
            TestAssert.Equal((byte)fixed61B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fixed61Args[1])!);
        }
        // Tail singles (slots 71-72), second-sampled bit-identical in dash73.
        int[] tailR = [193, 251, 208, 232, 222, 216];
        int[] tailG = [219, 211, 216, 208, 231, 211];
        int[] tailB = [229, 193, 232, 208, 209, 224];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] tailArgs = [70 + slot, null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] past = [76, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, past)!);
    }
    public static void PptxSyntheticThirteenthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseThirteenthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected thirteenth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var third = rendererType.GetMethod(
            "ShadeThirteenthRegimeThirdRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(third is not null, "Expected thirteenth-regime third-row 0.78 shade to remain inspectable by the Office evidence guard.");
        var fourth = rendererType.GetMethod(
            "ShadeThirteenthRegimeFourthRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourth is not null, "Expected thirteenth-regime fourth-row 0.83 shade to remain inspectable by the Office evidence guard.");
        var sixth = rendererType.GetMethod(
            "ShadeThirteenthRegimeSixthRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixth is not null, "Expected thirteenth-regime sixth-row 0.93 shade to remain inspectable by the Office evidence guard.");
        var seventh = rendererType.GetMethod(
            "ShadeThirteenthRegimeSeventhRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventh is not null, "Expected thirteenth-regime seventh-row 0.98 shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveThirteenthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected thirteenth-regime dark table to remain inspectable by the Office evidence guard.");
        var second = rendererType.GetMethod(
            "TryResolveThirteenthRegimeSecondFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(second is not null, "Expected thirteenth-regime second table to remain inspectable by the Office evidence guard.");
        var fixed43 = rendererType.GetMethod(
            "TryResolveThirteenthRegimeFixed43Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed43 is not null, "Expected thirteenth-regime fixed43 table to remain inspectable by the Office evidence guard.");
        var fixed73 = rendererType.GetMethod(
            "TryResolveThirteenthRegimeFixed73Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed73 is not null, "Expected thirteenth-regime fixed73 table to remain inspectable by the Office evidence guard.");

        // Seventy-seven points and fewer keep earlier regimes; seventy-eight-plus take dark/fixed/0.78/0.83/0.88/0.93/0.98/fixed/eighth/seventh/fixed/palest/fixed rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [77])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [78])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [79])!);
        // Computed shade pins on accent1 (79,129,189): 0.78 gives (62,101,147), 0.83 gives (66,107,157), 0.93 gives (73,120,176), 0.98 gives (77,126,185).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? shadedThird = third!.Invoke(null, [accent]);
        TestAssert.Equal((byte)62, (byte)rgbType!.GetProperty("Red")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)101, (byte)rgbType.GetProperty("Green")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)147, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedThird)!);
        object? shadedFourth = fourth!.Invoke(null, [accent]);
        TestAssert.Equal((byte)66, (byte)rgbType.GetProperty("Red")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)107, (byte)rgbType.GetProperty("Green")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)157, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedFourth)!);
        object? shadedSixth = sixth!.Invoke(null, [accent]);
        TestAssert.Equal((byte)73, (byte)rgbType.GetProperty("Red")!.GetValue(shadedSixth)!);
        TestAssert.Equal((byte)120, (byte)rgbType.GetProperty("Green")!.GetValue(shadedSixth)!);
        TestAssert.Equal((byte)176, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedSixth)!);
        object? shadedSeventh = seventh!.Invoke(null, [accent]);
        TestAssert.Equal((byte)77, (byte)rgbType.GetProperty("Red")!.GetValue(shadedSeventh)!);
        TestAssert.Equal((byte)126, (byte)rgbType.GetProperty("Green")!.GetValue(shadedSeventh)!);
        TestAssert.Equal((byte)185, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedSeventh)!);
        // Dark row replays eleventh-regime slots 1-6 byte-exact across both renders.
        int[] darkR = [49, 127, 101, 83, 47, 164];
        int[] darkG = [84, 50, 123, 64, 113, 98];
        int[] darkB = [125, 48, 56, 106, 131, 43];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Second row fixed table (slots 7-12), both renders agree on all six.
        int[] secondR = [54, 138, 110, 91, 52, 178];
        int[] secondG = [91, 55, 134, 70, 123, 107];
        int[] secondB = [136, 53, 62, 116, 142, 48];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] secondArgs = [6 + slot, null];
            TestAssert.Equal(true, (bool)second!.Invoke(null, secondArgs)!);
            TestAssert.Equal((byte)secondR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(secondArgs[1])!);
            TestAssert.Equal((byte)secondG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(secondArgs[1])!);
            TestAssert.Equal((byte)secondB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(secondArgs[1])!);
        }
        // Fixed row (slots 43-48), both renders agree on all six.
        int[] fixedR = [95, 195, 161, 136, 92, 247];
        int[] fixedG = [137, 96, 190, 112, 176, 156];
        int[] fixedB = [192, 94, 103, 167, 201, 89];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] fixedArgs = [42 + slot, null];
            TestAssert.Equal(true, (bool)fixed43!.Invoke(null, fixedArgs)!);
            TestAssert.Equal((byte)fixedR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fixedArgs[1])!);
        }
        // Fixed row (slots 73-78), both renders agree on all six.
        int[] tailR = [195, 227, 214, 206, 195, 251];
        int[] tailG = [206, 196, 225, 199, 220, 212];
        int[] tailB = [226, 195, 197, 216, 229, 194];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] tailArgs = [72 + slot, null];
            TestAssert.Equal(true, (bool)fixed73!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        // Fixed tail (slots 79-82), second samples bit-identical across dash79-dash83.
        var tail79 = rendererType.GetMethod(
            "TryResolveThirteenthRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail79 is not null, "Expected thirteenth-regime tail table to remain inspectable by the Office evidence guard.");
        object?[] tail79Args = [78, null];
        TestAssert.Equal(true, (bool)tail79!.Invoke(null, tail79Args)!);
        TestAssert.Equal((byte)208, (byte)rgbType.GetProperty("Red")!.GetValue(tail79Args[1])!);
        TestAssert.Equal((byte)216, (byte)rgbType.GetProperty("Green")!.GetValue(tail79Args[1])!);
        TestAssert.Equal((byte)232, (byte)rgbType.GetProperty("Blue")!.GetValue(tail79Args[1])!);
        object?[] tail80Args = [79, null];
        TestAssert.Equal(true, (bool)tail79!.Invoke(null, tail80Args)!);
        TestAssert.Equal((byte)232, (byte)rgbType.GetProperty("Red")!.GetValue(tail80Args[1])!);
        TestAssert.Equal((byte)208, (byte)rgbType.GetProperty("Green")!.GetValue(tail80Args[1])!);
        TestAssert.Equal((byte)208, (byte)rgbType.GetProperty("Blue")!.GetValue(tail80Args[1])!);
        object?[] tail81Args = [80, null];
        TestAssert.Equal(true, (bool)tail79!.Invoke(null, tail81Args)!);
        TestAssert.Equal((byte)222, (byte)rgbType.GetProperty("Red")!.GetValue(tail81Args[1])!);
        TestAssert.Equal((byte)231, (byte)rgbType.GetProperty("Green")!.GetValue(tail81Args[1])!);
        TestAssert.Equal((byte)209, (byte)rgbType.GetProperty("Blue")!.GetValue(tail81Args[1])!);
        object?[] tail82Args = [81, null];
        TestAssert.Equal(true, (bool)tail79!.Invoke(null, tail82Args)!);
        TestAssert.Equal((byte)216, (byte)rgbType.GetProperty("Red")!.GetValue(tail82Args[1])!);
        TestAssert.Equal((byte)211, (byte)rgbType.GetProperty("Green")!.GetValue(tail82Args[1])!);
        TestAssert.Equal((byte)224, (byte)rgbType.GetProperty("Blue")!.GetValue(tail82Args[1])!);
        object?[] past82 = [82, null];
        TestAssert.Equal(false, (bool)tail79.Invoke(null, past82)!);
        object?[] past = [78, null];
        TestAssert.Equal(false, (bool)fixed73.Invoke(null, past)!);
    }
    public static void PptxSyntheticFourteenthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseFourteenthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected fourteenth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var third = rendererType.GetMethod(
            "ShadeFourteenthRegimeThirdRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(third is not null, "Expected fourteenth-regime third-row 0.77 shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveFourteenthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected fourteenth-regime dark table to remain inspectable by the Office evidence guard.");
        var sixth = rendererType.GetMethod(
            "TryResolveFourteenthRegimeSixthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixth is not null, "Expected fourteenth-regime sixth replay to remain inspectable by the Office evidence guard.");
        var fixed49 = rendererType.GetMethod(
            "TryResolveFourteenthRegimeFixed49Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed49 is not null, "Expected fourteenth-regime fixed49 table to remain inspectable by the Office evidence guard.");
        var fixed73 = rendererType.GetMethod(
            "TryResolveFourteenthRegimeFixed73Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed73 is not null, "Expected fourteenth-regime fixed73 table to remain inspectable by the Office evidence guard.");
        var tail = rendererType.GetMethod(
            "TryResolveFourteenthRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected fourteenth-regime tail table to remain inspectable by the Office evidence guard.");

        // Eighty-three points and fewer keep earlier regimes; eighty-four-plus take dark/sixth/0.77/0.82/0.87/0.91/0.96/raw/fixed/sixth/eleventh/sixth/fixed/overflow/tail rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [83])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [84])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [85])!);
        // Computed shade pin on accent1 (79,129,189): 0.77 gives (61,99,146).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? shadedThird = third!.Invoke(null, [accent]);
        TestAssert.Equal((byte)61, (byte)rgbType!.GetProperty("Red")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)99, (byte)rgbType.GetProperty("Green")!.GetValue(shadedThird)!);
        TestAssert.Equal((byte)146, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedThird)!);
        // Dark row fixed table, both renders agree on all six.
        int[] darkR = [48, 124, 99, 81, 45, 160];
        int[] darkG = [82, 49, 120, 62, 110, 96];
        int[] darkB = [122, 47, 55, 104, 128, 42];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Sixth-regime replay (slots 7-12), both renders agree on all six.
        int[] sixthR = [54, 136, 109, 90, 51, 177];
        int[] sixthG = [90, 55, 133, 69, 122, 106];
        int[] sixthB = [134, 52, 61, 114, 141, 47];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] sixthArgs = [6 + slot, null];
            TestAssert.Equal(true, (bool)sixth!.Invoke(null, sixthArgs)!);
            TestAssert.Equal((byte)sixthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixthArgs[1])!);
            TestAssert.Equal((byte)sixthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixthArgs[1])!);
            TestAssert.Equal((byte)sixthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixthArgs[1])!);
        }
        // Fixed row (slots 49-54), both renders agree on all six.
        int[] fixedR = [109, 198, 166, 144, 106, 248];
        int[] fixedG = [145, 109, 194, 123, 181, 162];
        int[] fixedB = [195, 107, 115, 172, 203, 103];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] fixedArgs = [48 + slot, null];
            TestAssert.Equal(true, (bool)fixed49!.Invoke(null, fixedArgs)!);
            TestAssert.Equal((byte)fixedR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fixedArgs[1])!);
        }
        // Fixed row (slots 73-78), both renders agree on all six.
        int[] tailR = [183, 222, 206, 196, 182, 250];
        int[] tailG = [196, 183, 220, 188, 213, 204];
        int[] tailB = [221, 183, 185, 209, 225, 181];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] tailArgs = [72 + slot, null];
            TestAssert.Equal(true, (bool)fixed73!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        // Tail singles (slots 83-84), second-sampled bit-identical in dash85.
        int[] endR = [196, 251, 209, 233, 223, 217];
        int[] endG = [220, 213, 217, 210, 232, 212];
        int[] endB = [230, 196, 232, 209, 211, 225];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] endArgs = [82 + slot, null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, endArgs)!);
            TestAssert.Equal((byte)endR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(endArgs[1])!);
            TestAssert.Equal((byte)endG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(endArgs[1])!);
            TestAssert.Equal((byte)endB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(endArgs[1])!);
        }
        object?[] past = [88, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, past)!);
    }
    public static void PptxSyntheticFifteenthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseFifteenthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected fifteenth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var fourth = rendererType.GetMethod(
            "ShadeFifteenthRegimeFourthRowSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourth is not null, "Expected fifteenth-regime fourth-row 0.80 shade to remain inspectable by the Office evidence guard.");
        var dark = rendererType.GetMethod(
            "TryResolveFifteenthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected fifteenth-regime dark table to remain inspectable by the Office evidence guard.");
        var second = rendererType.GetMethod(
            "TryResolveFifteenthRegimeSecondFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(second is not null, "Expected fifteenth-regime second table to remain inspectable by the Office evidence guard.");
        var fixed55 = rendererType.GetMethod(
            "TryResolveFifteenthRegimeFixed55Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed55 is not null, "Expected fifteenth-regime fixed55 table to remain inspectable by the Office evidence guard.");
        var fixed85 = rendererType.GetMethod(
            "TryResolveFifteenthRegimeFixed85Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed85 is not null, "Expected fifteenth-regime fixed85 table to remain inspectable by the Office evidence guard.");

        // Eighty-nine points and fewer keep earlier regimes; ninety-plus take dark/fixed/fixed/0.80/0.85/0.90/0.94/0.98/thirteenth/fixed/twelfth/eighth/eleventh/fixed/fixed rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [89])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [90])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [91])!);
        // Computed shade pin on accent1 (79,129,189): 0.80 gives (63,103,151).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? shadedFourth = fourth!.Invoke(null, [accent]);
        TestAssert.Equal((byte)63, (byte)rgbType!.GetProperty("Red")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)103, (byte)rgbType.GetProperty("Green")!.GetValue(shadedFourth)!);
        TestAssert.Equal((byte)151, (byte)rgbType.GetProperty("Blue")!.GetValue(shadedFourth)!);
        // Dark row replays fourteenth-regime slots 1-6 byte-exact across both renders.
        int[] darkR = [48, 124, 99, 81, 45, 160];
        int[] darkG = [82, 49, 120, 62, 110, 96];
        int[] darkB = [122, 47, 55, 104, 128, 42];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Second row fixed table (slots 7-12), both renders agree on all six.
        int[] secondR = [53, 135, 108, 89, 50, 175];
        int[] secondG = [89, 54, 131, 68, 121, 105];
        int[] secondB = [133, 52, 61, 113, 139, 47];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] secondArgs = [6 + slot, null];
            TestAssert.Equal(true, (bool)second!.Invoke(null, secondArgs)!);
            TestAssert.Equal((byte)secondR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(secondArgs[1])!);
            TestAssert.Equal((byte)secondG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(secondArgs[1])!);
            TestAssert.Equal((byte)secondB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(secondArgs[1])!);
        }
        // Fixed row (slots 55-60), both renders agree on all six.
        int[] fixedR = [120, 201, 171, 151, 118, 248];
        int[] fixedG = [152, 121, 197, 132, 185, 167];
        int[] fixedB = [198, 119, 126, 177, 206, 116];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] fixedArgs = [54 + slot, null];
            TestAssert.Equal(true, (bool)fixed55!.Invoke(null, fixedArgs)!);
            TestAssert.Equal((byte)fixedR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fixedArgs[1])!);
        }
        // Fixed row (slots 85-90), both renders agree on all six.
        int[] tailR = [198, 228, 216, 208, 198, 251];
        int[] tailG = [208, 198, 226, 202, 221, 214];
        int[] tailB = [227, 198, 200, 218, 231, 197];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] tailArgs = [84 + slot, null];
            TestAssert.Equal(true, (bool)fixed85!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        // Fixed tail (slots 91-94), second samples bit-identical across dash91-dash95.
        var tail91 = rendererType.GetMethod(
            "TryResolveFifteenthRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail91 is not null, "Expected fifteenth-regime tail table to remain inspectable by the Office evidence guard.");
        object?[] tail91Args = [90, null];
        TestAssert.Equal(true, (bool)tail91!.Invoke(null, tail91Args)!);
        TestAssert.Equal((byte)209, (byte)rgbType.GetProperty("Red")!.GetValue(tail91Args[1])!);
        TestAssert.Equal((byte)217, (byte)rgbType.GetProperty("Green")!.GetValue(tail91Args[1])!);
        TestAssert.Equal((byte)232, (byte)rgbType.GetProperty("Blue")!.GetValue(tail91Args[1])!);
        object?[] tail92Args = [91, null];
        TestAssert.Equal(true, (bool)tail91!.Invoke(null, tail92Args)!);
        TestAssert.Equal((byte)233, (byte)rgbType.GetProperty("Red")!.GetValue(tail92Args[1])!);
        TestAssert.Equal((byte)210, (byte)rgbType.GetProperty("Green")!.GetValue(tail92Args[1])!);
        TestAssert.Equal((byte)209, (byte)rgbType.GetProperty("Blue")!.GetValue(tail92Args[1])!);
        object?[] tail93Args = [92, null];
        TestAssert.Equal(true, (bool)tail91!.Invoke(null, tail93Args)!);
        TestAssert.Equal((byte)223, (byte)rgbType.GetProperty("Red")!.GetValue(tail93Args[1])!);
        TestAssert.Equal((byte)232, (byte)rgbType.GetProperty("Green")!.GetValue(tail93Args[1])!);
        TestAssert.Equal((byte)211, (byte)rgbType.GetProperty("Blue")!.GetValue(tail93Args[1])!);
        object?[] tail94Args = [93, null];
        TestAssert.Equal(true, (bool)tail91!.Invoke(null, tail94Args)!);
        TestAssert.Equal((byte)217, (byte)rgbType.GetProperty("Red")!.GetValue(tail94Args[1])!);
        TestAssert.Equal((byte)212, (byte)rgbType.GetProperty("Green")!.GetValue(tail94Args[1])!);
        TestAssert.Equal((byte)225, (byte)rgbType.GetProperty("Blue")!.GetValue(tail94Args[1])!);
        object?[] past94 = [94, null];
        TestAssert.Equal(false, (bool)tail91.Invoke(null, past94)!);
        object?[] past = [90, null];
        TestAssert.Equal(false, (bool)fixed85.Invoke(null, past)!);
    }
    public static void PptxSyntheticSixteenthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseSixteenthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected sixteenth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var dark = rendererType.GetMethod(
            "TryResolveSixteenthRegimeDarkFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(dark is not null, "Expected sixteenth-regime dark table to remain inspectable by the Office evidence guard.");
        var seventh = rendererType.GetMethod(
            "TryResolveSixteenthRegimeSeventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventh is not null, "Expected sixteenth-regime seventh replay to remain inspectable by the Office evidence guard.");
        var fixed61 = rendererType.GetMethod(
            "TryResolveSixteenthRegimeFixed61Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed61 is not null, "Expected sixteenth-regime fixed61 table to remain inspectable by the Office evidence guard.");
        var tail = rendererType.GetMethod(
            "TryResolveSixteenthRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected sixteenth-regime tail table to remain inspectable by the Office evidence guard.");

        // Ninety-five points and fewer keep earlier regimes; ninety-six-plus take dark/seventh/fourth/seventh/0.85/0.88/0.93/0.96/raw/seventh/fixed/seventh/ninth/seventh/overflow/tenth/overflow/tail rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [95])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [96])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [97])!);
        // Dark row fixed table, both renders agree on all six.
        int[] darkR = [47, 122, 98, 80, 45, 158];
        int[] darkG = [80, 48, 119, 61, 109, 94];
        int[] darkB = [120, 46, 54, 102, 126, 41];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] args = [slot, null];
            TestAssert.Equal(true, (bool)dark!.Invoke(null, args)!);
            TestAssert.Equal((byte)darkR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(args[1])!);
            TestAssert.Equal((byte)darkB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(args[1])!);
        }
        // Seventh-regime replay (slots 7-12), both renders agree on all six.
        int[] seventhR = [53, 134, 107, 88, 50, 173];
        int[] seventhG = [89, 53, 130, 68, 119, 104];
        int[] seventhB = [132, 51, 60, 112, 138, 46];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] seventhArgs = [6 + slot, null];
            TestAssert.Equal(true, (bool)seventh!.Invoke(null, seventhArgs)!);
            TestAssert.Equal((byte)seventhR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventhArgs[1])!);
            TestAssert.Equal((byte)seventhG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventhArgs[1])!);
            TestAssert.Equal((byte)seventhB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventhArgs[1])!);
        }
        // Fixed row (slots 61-66), both renders agree on all six.
        int[] fixedR = [128, 203, 175, 156, 126, 248];
        int[] fixedG = [157, 129, 199, 139, 188, 171];
        int[] fixedB = [201, 127, 133, 180, 208, 124];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] fixedArgs = [60 + slot, null];
            TestAssert.Equal(true, (bool)fixed61!.Invoke(null, fixedArgs)!);
            TestAssert.Equal((byte)fixedR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fixedArgs[1])!);
            TestAssert.Equal((byte)fixedB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fixedArgs[1])!);
        }
        // Tail singles (slots 95-96), second-sampled bit-identical in dash97.
        int[] tailR = [199, 251, 211, 234, 224, 218];
        int[] tailG = [222, 215, 218, 211, 232, 213];
        int[] tailB = [231, 199, 233, 211, 212, 226];
        for (int slot = 0; slot < 6; slot++)
        {
            object?[] tailArgs = [94 + slot, null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] past = [100, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, past)!);
    }
    public static void PptxSyntheticSeventeenthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseSeventeenthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected seventeenth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var seventhMid = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeSeventhMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventhMid is not null, "Expected seventeenth-regime seventhmid table to remain inspectable by the Office evidence guard.");
        var fifthMid = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeFifthMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifthMid is not null, "Expected seventeenth-regime fifthmid table to remain inspectable by the Office evidence guard.");
        var shade = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeShadeFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(shade is not null, "Expected seventeenth-regime shade table to remain inspectable by the Office evidence guard.");
        var thirteenthSixth = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeThirteenthSixthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(thirteenthSixth is not null, "Expected seventeenth-regime thirteenthsixth table to remain inspectable by the Office evidence guard.");
        var seventhLight = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeSeventhLightFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventhLight is not null, "Expected seventeenth-regime seventhlight table to remain inspectable by the Office evidence guard.");
        var raw = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeRawFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(raw is not null, "Expected seventeenth-regime raw table to remain inspectable by the Office evidence guard.");
        var fixed55 = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeFixed55Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fixed55 is not null, "Expected seventeenth-regime fixed55 table to remain inspectable by the Office evidence guard.");
        var tenth = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeTenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tenth is not null, "Expected seventeenth-regime tenth table to remain inspectable by the Office evidence guard.");
        var sixth25 = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeSixth25Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixth25 is not null, "Expected seventeenth-regime sixth25 table to remain inspectable by the Office evidence guard.");
        var ninth = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeNinthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninth is not null, "Expected seventeenth-regime ninth table to remain inspectable by the Office evidence guard.");
        var eleventh = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeEleventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventh is not null, "Expected seventeenth-regime eleventh table to remain inspectable by the Office evidence guard.");
        var thirteenth = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeThirteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(thirteenth is not null, "Expected seventeenth-regime thirteenth table to remain inspectable by the Office evidence guard.");
        var twelfth = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeTwelfthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfth is not null, "Expected seventeenth-regime twelfth table to remain inspectable by the Office evidence guard.");
        var overflow = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflow is not null, "Expected seventeenth-regime overflow table to remain inspectable by the Office evidence guard.");
        var eleventhLate = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeEleventhLateFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventhLate is not null, "Expected seventeenth-regime eleventhlate table to remain inspectable by the Office evidence guard.");
        var sixteenthTailReplay = rendererType.GetMethod(
            "TryResolveSeventeenthRegimeSixteenthTailReplayFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixteenthTailReplay is not null, "Expected seventeenth-regime sixteenthtailreplay table to remain inspectable by the Office evidence guard.");
        // One-hundred-one points and fewer keep earlier regimes; one-hundred-two-plus take the seventeenth rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [101])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [102])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [103])!);
        // Seventeenth-regime seventh-mid row (slots 20, 21, 23, 24), dash102/dash103 agree bit-identical.
        int[] seventhMidIdx = [19, 20, 22, 23];
        int[] seventhMidR = [151, 122, 57, 195];
        int[] seventhMidG = [61, 147, 135, 118];
        int[] seventhMidB = [59, 69, 156, 53];
        for (int slot = 0; slot < seventhMidIdx.Length; slot++)
        {
            object?[] seventhMidArgs = [seventhMidIdx[slot], null];
            TestAssert.Equal(true, (bool)seventhMid!.Invoke(null, seventhMidArgs)!);
            TestAssert.Equal((byte)seventhMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventhMidArgs[1])!);
            TestAssert.Equal((byte)seventhMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventhMidArgs[1])!);
            TestAssert.Equal((byte)seventhMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventhMidArgs[1])!);
        }
        object?[] seventhMidPast = [21, null];
        TestAssert.Equal(false, (bool)seventhMid.Invoke(null, seventhMidPast)!);
        // Seventeenth-regime fifth-mid row (slots 25-30), dash102/dash103 agree bit-identical.
        int[] fifthMidIdx = [24, 25, 26, 27, 28, 29];
        int[] fifthMidR = [64, 159, 128, 105, 61, 206];
        int[] fifthMidG = [106, 65, 155, 82, 142, 124];
        int[] fifthMidB = [157, 63, 73, 134, 164, 57];
        for (int slot = 0; slot < fifthMidIdx.Length; slot++)
        {
            object?[] fifthMidArgs = [fifthMidIdx[slot], null];
            TestAssert.Equal(true, (bool)fifthMid!.Invoke(null, fifthMidArgs)!);
            TestAssert.Equal((byte)fifthMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifthMidArgs[1])!);
            TestAssert.Equal((byte)fifthMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifthMidArgs[1])!);
            TestAssert.Equal((byte)fifthMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifthMidArgs[1])!);
        }
        object?[] fifthMidPast = [30, null];
        TestAssert.Equal(false, (bool)fifthMid.Invoke(null, fifthMidPast)!);
        // Seventeenth-regime shade single (slot 36), dash102/dash103 agree bit-identical.
        int[] shadeIdx = [35];
        int[] shadeR = [216];
        int[] shadeG = [131];
        int[] shadeB = [60];
        for (int slot = 0; slot < shadeIdx.Length; slot++)
        {
            object?[] shadeArgs = [shadeIdx[slot], null];
            TestAssert.Equal(true, (bool)shade!.Invoke(null, shadeArgs)!);
            TestAssert.Equal((byte)shadeR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(shadeArgs[1])!);
            TestAssert.Equal((byte)shadeG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(shadeArgs[1])!);
            TestAssert.Equal((byte)shadeB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(shadeArgs[1])!);
        }
        object?[] shadePast = [34, null];
        TestAssert.Equal(false, (bool)shade.Invoke(null, shadePast)!);
        // Seventeenth-regime thirteenth-sixth replacement row (slots 37-42), dash102/dash103 agree bit-identical.
        int[] thirteenthSixthIdx = [36, 37, 38, 39, 40, 41];
        int[] thirteenthSixthR = [71, 175, 141, 116, 68, 225];
        int[] thirteenthSixthG = [117, 72, 170, 90, 156, 136];
        int[] thirteenthSixthB = [172, 69, 80, 147, 180, 63];
        for (int slot = 0; slot < thirteenthSixthIdx.Length; slot++)
        {
            object?[] thirteenthSixthArgs = [thirteenthSixthIdx[slot], null];
            TestAssert.Equal(true, (bool)thirteenthSixth!.Invoke(null, thirteenthSixthArgs)!);
            TestAssert.Equal((byte)thirteenthSixthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(thirteenthSixthArgs[1])!);
            TestAssert.Equal((byte)thirteenthSixthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(thirteenthSixthArgs[1])!);
            TestAssert.Equal((byte)thirteenthSixthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(thirteenthSixthArgs[1])!);
        }
        object?[] thirteenthSixthPast = [35, null];
        TestAssert.Equal(false, (bool)thirteenthSixth.Invoke(null, thirteenthSixthPast)!);
        // Seventeenth-regime seventh-light replacement row (slots 43-48), dash102/dash103 agree bit-identical.
        int[] seventhLightIdx = [42, 43, 44, 45, 46, 47];
        int[] seventhLightR = [74, 181, 146, 121, 70, 233];
        int[] seventhLightG = [122, 75, 177, 94, 162, 141];
        int[] seventhLightB = [178, 72, 84, 153, 187, 66];
        for (int slot = 0; slot < seventhLightIdx.Length; slot++)
        {
            object?[] seventhLightArgs = [seventhLightIdx[slot], null];
            TestAssert.Equal(true, (bool)seventhLight!.Invoke(null, seventhLightArgs)!);
            TestAssert.Equal((byte)seventhLightR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventhLightArgs[1])!);
            TestAssert.Equal((byte)seventhLightG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventhLightArgs[1])!);
            TestAssert.Equal((byte)seventhLightB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventhLightArgs[1])!);
        }
        object?[] seventhLightPast = [41, null];
        TestAssert.Equal(false, (bool)seventhLight.Invoke(null, seventhLightPast)!);
        // Seventeenth-regime raw replacement row (slots 49-54), dash102/dash103 agree bit-identical.
        int[] rawIdx = [48, 49, 50, 51, 52, 53];
        int[] rawR = [77, 189, 152, 126, 74, 243];
        int[] rawG = [127, 78, 184, 98, 169, 147];
        int[] rawB = [186, 75, 87, 159, 194, 69];
        for (int slot = 0; slot < rawIdx.Length; slot++)
        {
            object?[] rawArgs = [rawIdx[slot], null];
            TestAssert.Equal(true, (bool)raw!.Invoke(null, rawArgs)!);
            TestAssert.Equal((byte)rawR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(rawArgs[1])!);
            TestAssert.Equal((byte)rawG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(rawArgs[1])!);
            TestAssert.Equal((byte)rawB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(rawArgs[1])!);
        }
        object?[] rawPast = [47, null];
        TestAssert.Equal(false, (bool)raw.Invoke(null, rawPast)!);
        // Seventeenth-regime fixed55 replacement row (slots 55-60), dash102/dash103 agree bit-identical.
        int[] fixed55Idx = [54, 55, 56, 57, 58, 59];
        int[] fixed55R = [92, 194, 159, 134, 88, 247];
        int[] fixed55G = [135, 92, 190, 109, 175, 155];
        int[] fixed55B = [191, 90, 100, 166, 200, 84];
        for (int slot = 0; slot < fixed55Idx.Length; slot++)
        {
            object?[] fixed55Args = [fixed55Idx[slot], null];
            TestAssert.Equal(true, (bool)fixed55!.Invoke(null, fixed55Args)!);
            TestAssert.Equal((byte)fixed55R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fixed55Args[1])!);
            TestAssert.Equal((byte)fixed55G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fixed55Args[1])!);
            TestAssert.Equal((byte)fixed55B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fixed55Args[1])!);
        }
        object?[] fixed55Past = [53, null];
        TestAssert.Equal(false, (bool)fixed55.Invoke(null, fixed55Past)!);
        // Seventeenth-regime tenth replay (slots 61-66 take tenth slots 37-42), dash102/dash103 agree bit-identical.
        int[] tenthIdx = [60, 61, 62, 63, 64, 65];
        int[] tenthR = [118, 200, 170, 149, 115, 248];
        int[] tenthG = [150, 118, 196, 130, 184, 166];
        int[] tenthB = [198, 116, 123, 176, 205, 113];
        for (int slot = 0; slot < tenthIdx.Length; slot++)
        {
            object?[] tenthArgs = [tenthIdx[slot], null];
            TestAssert.Equal(true, (bool)tenth!.Invoke(null, tenthArgs)!);
            TestAssert.Equal((byte)tenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tenthArgs[1])!);
            TestAssert.Equal((byte)tenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tenthArgs[1])!);
            TestAssert.Equal((byte)tenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tenthArgs[1])!);
        }
        object?[] tenthPast = [59, null];
        TestAssert.Equal(false, (bool)tenth.Invoke(null, tenthPast)!);
        // Seventeenth-regime sixth25 replay (slot 67), dash102/dash103 agree bit-identical.
        int[] sixth25Idx = [66];
        int[] sixth25R = [133];
        int[] sixth25G = [160];
        int[] sixth25B = [202];
        for (int slot = 0; slot < sixth25Idx.Length; slot++)
        {
            object?[] sixth25Args = [sixth25Idx[slot], null];
            TestAssert.Equal(true, (bool)sixth25!.Invoke(null, sixth25Args)!);
            TestAssert.Equal((byte)sixth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixth25Args[1])!);
            TestAssert.Equal((byte)sixth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixth25Args[1])!);
            TestAssert.Equal((byte)sixth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixth25Args[1])!);
        }
        object?[] sixth25Past = [65, null];
        TestAssert.Equal(false, (bool)sixth25.Invoke(null, sixth25Past)!);
        // Seventeenth-regime ninth replay (slots 68-72 take ninth slots 38-42), dash102/dash103 agree bit-identical.
        int[] ninthIdx = [67, 68, 69, 70, 71];
        int[] ninthR = [206, 180, 163, 136, 249];
        int[] ninthG = [138, 202, 147, 192, 177];
        int[] ninthB = [137, 142, 185, 210, 134];
        for (int slot = 0; slot < ninthIdx.Length; slot++)
        {
            object?[] ninthArgs = [ninthIdx[slot], null];
            TestAssert.Equal(true, (bool)ninth!.Invoke(null, ninthArgs)!);
            TestAssert.Equal((byte)ninthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninthArgs[1])!);
            TestAssert.Equal((byte)ninthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninthArgs[1])!);
            TestAssert.Equal((byte)ninthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninthArgs[1])!);
        }
        object?[] ninthPast = [66, null];
        TestAssert.Equal(false, (bool)ninth.Invoke(null, ninthPast)!);
        // Seventeenth-regime eleventh replay (slots 73-78 take eleventh slots 49-54), dash102/dash103 agree bit-identical.
        int[] eleventhIdx = [72, 73, 74, 75, 76, 77];
        int[] eleventhR = [153, 211, 188, 173, 152, 249];
        int[] eleventhG = [174, 153, 208, 161, 198, 185];
        int[] eleventhB = [209, 152, 156, 192, 215, 150];
        for (int slot = 0; slot < eleventhIdx.Length; slot++)
        {
            object?[] eleventhArgs = [eleventhIdx[slot], null];
            TestAssert.Equal(true, (bool)eleventh!.Invoke(null, eleventhArgs)!);
            TestAssert.Equal((byte)eleventhR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventhArgs[1])!);
            TestAssert.Equal((byte)eleventhG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventhArgs[1])!);
            TestAssert.Equal((byte)eleventhB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventhArgs[1])!);
        }
        object?[] eleventhPast = [71, null];
        TestAssert.Equal(false, (bool)eleventh.Invoke(null, eleventhPast)!);
        // Seventeenth-regime thirteenth replay (slots 79-84 take thirteenth slots 61-66), dash102/dash103 agree bit-identical.
        int[] thirteenthIdx = [78, 79, 80, 81, 82, 83];
        int[] thirteenthR = [164, 215, 195, 182, 163, 250];
        int[] thirteenthG = [182, 165, 212, 171, 204, 192];
        int[] thirteenthB = [213, 164, 167, 198, 218, 162];
        for (int slot = 0; slot < thirteenthIdx.Length; slot++)
        {
            object?[] thirteenthArgs = [thirteenthIdx[slot], null];
            TestAssert.Equal(true, (bool)thirteenth!.Invoke(null, thirteenthArgs)!);
            TestAssert.Equal((byte)thirteenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(thirteenthArgs[1])!);
            TestAssert.Equal((byte)thirteenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(thirteenthArgs[1])!);
            TestAssert.Equal((byte)thirteenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(thirteenthArgs[1])!);
        }
        object?[] thirteenthPast = [77, null];
        TestAssert.Equal(false, (bool)thirteenth.Invoke(null, thirteenthPast)!);
        // Seventeenth-regime twelfth replay (slots 85-90 take twelfth slots 61-66), dash102/dash103 agree bit-identical.
        int[] twelfthIdx = [84, 85, 86, 87, 88, 89];
        int[] twelfthR = [178, 220, 203, 192, 177, 250];
        int[] twelfthG = [193, 179, 218, 184, 210, 201];
        int[] twelfthB = [219, 178, 181, 206, 223, 176];
        for (int slot = 0; slot < twelfthIdx.Length; slot++)
        {
            object?[] twelfthArgs = [twelfthIdx[slot], null];
            TestAssert.Equal(true, (bool)twelfth!.Invoke(null, twelfthArgs)!);
            TestAssert.Equal((byte)twelfthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfthArgs[1])!);
            TestAssert.Equal((byte)twelfthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfthArgs[1])!);
            TestAssert.Equal((byte)twelfthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfthArgs[1])!);
        }
        object?[] twelfthPast = [83, null];
        TestAssert.Equal(false, (bool)twelfth.Invoke(null, twelfthPast)!);
        // Seventeenth-regime overflow replay (slots 91/93/97-100), dash102/dash103 agree bit-identical.
        int[] overflowIdx = [90, 92, 96, 97, 98, 99];
        int[] overflowR = [188, 209, 200, 229, 218, 209];
        int[] overflowG = [200, 222, 209, 200, 228, 203];
        int[] overflowB = [223, 190, 228, 199, 204, 219];
        for (int slot = 0; slot < overflowIdx.Length; slot++)
        {
            object?[] overflowArgs = [overflowIdx[slot], null];
            TestAssert.Equal(true, (bool)overflow!.Invoke(null, overflowArgs)!);
            TestAssert.Equal((byte)overflowR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflowArgs[1])!);
            TestAssert.Equal((byte)overflowG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflowArgs[1])!);
            TestAssert.Equal((byte)overflowB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflowArgs[1])!);
        }
        object?[] overflowPast = [91, null];
        TestAssert.Equal(false, (bool)overflow.Invoke(null, overflowPast)!);
        // Seventeenth-regime eleventh-late replay (slots 92/94-96), dash102/dash103 agree bit-identical.
        int[] eleventhLateIdx = [91, 93, 94, 95];
        int[] eleventhLateR = [225, 202, 190, 251];
        int[] eleventhLateG = [191, 195, 217, 209];
        int[] eleventhLateB = [191, 213, 228, 189];
        for (int slot = 0; slot < eleventhLateIdx.Length; slot++)
        {
            object?[] eleventhLateArgs = [eleventhLateIdx[slot], null];
            TestAssert.Equal(true, (bool)eleventhLate!.Invoke(null, eleventhLateArgs)!);
            TestAssert.Equal((byte)eleventhLateR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventhLateArgs[1])!);
            TestAssert.Equal((byte)eleventhLateG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventhLateArgs[1])!);
            TestAssert.Equal((byte)eleventhLateB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventhLateArgs[1])!);
        }
        object?[] eleventhLatePast = [92, null];
        TestAssert.Equal(false, (bool)eleventhLate.Invoke(null, eleventhLatePast)!);
        // Seventeenth-regime sixteenth-tail replay (slots 101-106), dash102/dash103 agree bit-identical.
        int[] sixteenthTailReplayIdx = [100, 101, 102, 103, 104, 105];
        int[] sixteenthTailReplayR = [199, 251, 211, 234, 224, 218];
        int[] sixteenthTailReplayG = [222, 215, 218, 211, 232, 213];
        int[] sixteenthTailReplayB = [231, 199, 233, 211, 212, 226];
        for (int slot = 0; slot < sixteenthTailReplayIdx.Length; slot++)
        {
            object?[] sixteenthTailReplayArgs = [sixteenthTailReplayIdx[slot], null];
            TestAssert.Equal(true, (bool)sixteenthTailReplay!.Invoke(null, sixteenthTailReplayArgs)!);
            TestAssert.Equal((byte)sixteenthTailReplayR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixteenthTailReplayArgs[1])!);
            TestAssert.Equal((byte)sixteenthTailReplayG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixteenthTailReplayArgs[1])!);
            TestAssert.Equal((byte)sixteenthTailReplayB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixteenthTailReplayArgs[1])!);
        }
        object?[] sixteenthTailReplayPast = [106, null];
        TestAssert.Equal(false, (bool)sixteenthTailReplay.Invoke(null, sixteenthTailReplayPast)!);
    }
    public static void PptxSyntheticEighteenthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseEighteenthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected eighteenth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var earlySingles = rendererType.GetMethod(
            "TryResolveEighteenthRegimeEarlySinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(earlySingles is not null, "Expected eighteenth-regime earlysingles table to remain inspectable by the Office evidence guard.");
        var shadeSingles = rendererType.GetMethod(
            "TryResolveEighteenthRegimeShadeSinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(shadeSingles is not null, "Expected eighteenth-regime shadesingles table to remain inspectable by the Office evidence guard.");
        var rawDrift = rendererType.GetMethod(
            "TryResolveEighteenthRegimeRawDriftFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(rawDrift is not null, "Expected eighteenth-regime rawdrift table to remain inspectable by the Office evidence guard.");
        var raw = rendererType.GetMethod(
            "TryResolveEighteenthRegimeRawFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(raw is not null, "Expected eighteenth-regime raw table to remain inspectable by the Office evidence guard.");
        var twelfthSingle = rendererType.GetMethod(
            "TryResolveEighteenthRegimeTwelfthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfthSingle is not null, "Expected eighteenth-regime twelfthsingle table to remain inspectable by the Office evidence guard.");
        var tailSingle = rendererType.GetMethod(
            "TryResolveEighteenthRegimeTailSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tailSingle is not null, "Expected eighteenth-regime tailsingle table to remain inspectable by the Office evidence guard.");
        var seventh = rendererType.GetMethod(
            "TryResolveEighteenthRegimeSeventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventh is not null, "Expected eighteenth-regime seventh table to remain inspectable by the Office evidence guard.");
        var eighth = rendererType.GetMethod(
            "TryResolveEighteenthRegimeEighthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighth is not null, "Expected eighteenth-regime eighth table to remain inspectable by the Office evidence guard.");
        var twelfthEarly = rendererType.GetMethod(
            "TryResolveEighteenthRegimeTwelfthEarlyFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfthEarly is not null, "Expected eighteenth-regime twelfthearly table to remain inspectable by the Office evidence guard.");
        var eighthMid = rendererType.GetMethod(
            "TryResolveEighteenthRegimeEighthMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighthMid is not null, "Expected eighteenth-regime eighthmid table to remain inspectable by the Office evidence guard.");
        var overflowEarly = rendererType.GetMethod(
            "TryResolveEighteenthRegimeOverflowEarlyFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflowEarly is not null, "Expected eighteenth-regime overflowearly table to remain inspectable by the Office evidence guard.");
        var sixth = rendererType.GetMethod(
            "TryResolveEighteenthRegimeSixthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixth is not null, "Expected eighteenth-regime sixth table to remain inspectable by the Office evidence guard.");
        var eighthLate = rendererType.GetMethod(
            "TryResolveEighteenthRegimeEighthLateFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighthLate is not null, "Expected eighteenth-regime eighthlate table to remain inspectable by the Office evidence guard.");
        var eleventhLate = rendererType.GetMethod(
            "TryResolveEighteenthRegimeEleventhLateFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventhLate is not null, "Expected eighteenth-regime eleventhlate table to remain inspectable by the Office evidence guard.");
        var twelfth71 = rendererType.GetMethod(
            "TryResolveEighteenthRegimeTwelfth71Fill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfth71 is not null, "Expected eighteenth-regime twelfth71 table to remain inspectable by the Office evidence guard.");
        var eleventhSingle = rendererType.GetMethod(
            "TryResolveEighteenthRegimeEleventhSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventhSingle is not null, "Expected eighteenth-regime eleventhsingle table to remain inspectable by the Office evidence guard.");
        var overflowLate = rendererType.GetMethod(
            "TryResolveEighteenthRegimeOverflowLateFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflowLate is not null, "Expected eighteenth-regime overflowlate table to remain inspectable by the Office evidence guard.");
        var sixteenthSingle = rendererType.GetMethod(
            "TryResolveEighteenthRegimeSixteenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixteenthSingle is not null, "Expected eighteenth-regime sixteenthsingle table to remain inspectable by the Office evidence guard.");
        // One-hundred-seven points and fewer keep earlier regimes; one-hundred-eight-plus take the eighteenth rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [107])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [108])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [109])!);
        // Eighteenth-regime early singles (slots 13/17-19/22), dash108/dash109 agree bit-identical.
        int[] earlySinglesIdx = [12, 16, 17, 18, 21];
        int[] earlySinglesR = [56, 53, 183, 60, 99];
        int[] earlySinglesG = [94, 126, 110, 100, 77];
        int[] earlySinglesB = [139, 146, 49, 148, 126];
        for (int slot = 0; slot < earlySinglesIdx.Length; slot++)
        {
            object?[] earlySinglesArgs = [earlySinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)earlySingles!.Invoke(null, earlySinglesArgs)!);
            TestAssert.Equal((byte)earlySinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(earlySinglesArgs[1])!);
            TestAssert.Equal((byte)earlySinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(earlySinglesArgs[1])!);
            TestAssert.Equal((byte)earlySinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(earlySinglesArgs[1])!);
        }
        object?[] earlySinglesPast = [20, null];
        TestAssert.Equal(false, (bool)earlySingles.Invoke(null, earlySinglesPast)!);
        // Eighteenth-regime shade singles (slots 31-35), dash108/dash109 agree bit-identical.
        int[] shadeSinglesIdx = [30, 31, 32, 33, 34];
        int[] shadeSinglesR = [67, 166, 133, 110, 64];
        int[] shadeSinglesG = [111, 68, 161, 85, 148];
        int[] shadeSinglesB = [163, 65, 76, 139, 171];
        for (int slot = 0; slot < shadeSinglesIdx.Length; slot++)
        {
            object?[] shadeSinglesArgs = [shadeSinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)shadeSingles!.Invoke(null, shadeSinglesArgs)!);
            TestAssert.Equal((byte)shadeSinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(shadeSinglesArgs[1])!);
            TestAssert.Equal((byte)shadeSinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(shadeSinglesArgs[1])!);
            TestAssert.Equal((byte)shadeSinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(shadeSinglesArgs[1])!);
        }
        object?[] shadeSinglesPast = [29, null];
        TestAssert.Equal(false, (bool)shadeSingles.Invoke(null, shadeSinglesPast)!);
        // Eighteenth-regime raw-drift singles (slots 49-51/53-54), dash108/dash109 agree bit-identical.
        int[] rawDriftIdx = [48, 49, 50, 52, 53];
        int[] rawDriftR = [76, 186, 150, 72, 239];
        int[] rawDriftG = [125, 77, 181, 166, 145];
        int[] rawDriftB = [183, 74, 86, 192, 67];
        for (int slot = 0; slot < rawDriftIdx.Length; slot++)
        {
            object?[] rawDriftArgs = [rawDriftIdx[slot], null];
            TestAssert.Equal(true, (bool)rawDrift!.Invoke(null, rawDriftArgs)!);
            TestAssert.Equal((byte)rawDriftR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(rawDriftArgs[1])!);
            TestAssert.Equal((byte)rawDriftG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(rawDriftArgs[1])!);
            TestAssert.Equal((byte)rawDriftB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(rawDriftArgs[1])!);
        }
        object?[] rawDriftPast = [51, null];
        TestAssert.Equal(false, (bool)rawDrift.Invoke(null, rawDriftPast)!);
        // Eighteenth-regime raw row (slots 55-60), dash108/dash109 agree bit-identical.
        int[] rawIdx = [54, 55, 56, 57, 58, 59];
        int[] rawR = [79, 192, 155, 128, 75, 247];
        int[] rawG = [129, 80, 187, 100, 172, 150];
        int[] rawB = [189, 77, 89, 162, 198, 70];
        for (int slot = 0; slot < rawIdx.Length; slot++)
        {
            object?[] rawArgs = [rawIdx[slot], null];
            TestAssert.Equal(true, (bool)raw!.Invoke(null, rawArgs)!);
            TestAssert.Equal((byte)rawR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(rawArgs[1])!);
            TestAssert.Equal((byte)rawG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(rawArgs[1])!);
            TestAssert.Equal((byte)rawB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(rawArgs[1])!);
        }
        object?[] rawPast = [53, null];
        TestAssert.Equal(false, (bool)raw.Invoke(null, rawPast)!);
        // Eighteenth-regime twelfth single (slot 74), dash108/dash109 agree bit-identical.
        int[] twelfthSingleIdx = [73];
        int[] twelfthSingleR = [207];
        int[] twelfthSingleG = [143];
        int[] twelfthSingleB = [142];
        for (int slot = 0; slot < twelfthSingleIdx.Length; slot++)
        {
            object?[] twelfthSingleArgs = [twelfthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)twelfthSingle!.Invoke(null, twelfthSingleArgs)!);
            TestAssert.Equal((byte)twelfthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfthSingleArgs[1])!);
            TestAssert.Equal((byte)twelfthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfthSingleArgs[1])!);
            TestAssert.Equal((byte)twelfthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfthSingleArgs[1])!);
        }
        object?[] twelfthSinglePast = [72, null];
        TestAssert.Equal(false, (bool)twelfthSingle.Invoke(null, twelfthSinglePast)!);
        // Eighteenth-regime tail singles (slots 107/109-112), dash108/dash109 agree bit-identical.
        int[] tailSingleIdx = [106, 108, 109, 110, 111];
        int[] tailSingleR = [202, 211, 234, 224, 218];
        int[] tailSingleG = [223, 218, 211, 232, 213];
        int[] tailSingleB = [232, 233, 211, 212, 226];
        for (int slot = 0; slot < tailSingleIdx.Length; slot++)
        {
            object?[] tailSingleArgs = [tailSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)tailSingle!.Invoke(null, tailSingleArgs)!);
            TestAssert.Equal((byte)tailSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailSingleArgs[1])!);
            TestAssert.Equal((byte)tailSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailSingleArgs[1])!);
            TestAssert.Equal((byte)tailSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailSingleArgs[1])!);
        }
        object?[] tailSinglePast = [105, null];
        TestAssert.Equal(false, (bool)tailSingle.Invoke(null, tailSinglePast)!);
        // Eighteenth-regime seventh replay (slots 61-66 take seventh slots 25-30), dash108/dash109 agree bit-identical.
        int[] seventhIdx = [60, 61, 62, 63, 64, 65];
        int[] seventhR = [106, 197, 165, 142, 103, 248];
        int[] seventhG = [143, 106, 193, 120, 180, 160];
        int[] seventhB = [195, 104, 112, 171, 203, 100];
        for (int slot = 0; slot < seventhIdx.Length; slot++)
        {
            object?[] seventhArgs = [seventhIdx[slot], null];
            TestAssert.Equal(true, (bool)seventh!.Invoke(null, seventhArgs)!);
            TestAssert.Equal((byte)seventhR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventhArgs[1])!);
            TestAssert.Equal((byte)seventhG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventhArgs[1])!);
            TestAssert.Equal((byte)seventhB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventhArgs[1])!);
        }
        object?[] seventhPast = [59, null];
        TestAssert.Equal(false, (bool)seventh.Invoke(null, seventhPast)!);
        // Eighteenth-regime eighth replay (slots 67-72 take eighth slots 31-36), dash108/dash109 agree bit-identical.
        int[] eighthIdx = [66, 67, 68, 69, 70, 71];
        int[] eighthR = [126, 202, 174, 155, 124, 248];
        int[] eighthG = [155, 126, 198, 137, 187, 170];
        int[] eighthB = [200, 125, 131, 179, 207, 121];
        for (int slot = 0; slot < eighthIdx.Length; slot++)
        {
            object?[] eighthArgs = [eighthIdx[slot], null];
            TestAssert.Equal(true, (bool)eighth!.Invoke(null, eighthArgs)!);
            TestAssert.Equal((byte)eighthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighthArgs[1])!);
            TestAssert.Equal((byte)eighthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighthArgs[1])!);
            TestAssert.Equal((byte)eighthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighthArgs[1])!);
        }
        object?[] eighthPast = [65, null];
        TestAssert.Equal(false, (bool)eighth.Invoke(null, eighthPast)!);
        // Eighteenth-regime twelfth-early replay (slots 73/75-78), dash108/dash109 agree bit-identical.
        int[] twelfthEarlyIdx = [72, 74, 75, 76, 77];
        int[] twelfthEarlyR = [140, 181, 164, 139, 249];
        int[] twelfthEarlyG = [165, 203, 149, 192, 178];
        int[] twelfthEarlyB = [204, 144, 186, 211, 137];
        for (int slot = 0; slot < twelfthEarlyIdx.Length; slot++)
        {
            object?[] twelfthEarlyArgs = [twelfthEarlyIdx[slot], null];
            TestAssert.Equal(true, (bool)twelfthEarly!.Invoke(null, twelfthEarlyArgs)!);
            TestAssert.Equal((byte)twelfthEarlyR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfthEarlyArgs[1])!);
            TestAssert.Equal((byte)twelfthEarlyG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfthEarlyArgs[1])!);
            TestAssert.Equal((byte)twelfthEarlyB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfthEarlyArgs[1])!);
        }
        object?[] twelfthEarlyPast = [73, null];
        TestAssert.Equal(false, (bool)twelfthEarly.Invoke(null, twelfthEarlyPast)!);
        // Eighteenth-regime eighth-mid replay (slots 79-84 take eighth slots 37-42), dash108/dash109 agree bit-identical.
        int[] eighthMidIdx = [78, 79, 80, 81, 82, 83];
        int[] eighthMidR = [157, 212, 190, 176, 156, 249];
        int[] eighthMidG = [177, 157, 209, 164, 200, 187];
        int[] eighthMidB = [210, 156, 160, 194, 216, 154];
        for (int slot = 0; slot < eighthMidIdx.Length; slot++)
        {
            object?[] eighthMidArgs = [eighthMidIdx[slot], null];
            TestAssert.Equal(true, (bool)eighthMid!.Invoke(null, eighthMidArgs)!);
            TestAssert.Equal((byte)eighthMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighthMidArgs[1])!);
            TestAssert.Equal((byte)eighthMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighthMidArgs[1])!);
            TestAssert.Equal((byte)eighthMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighthMidArgs[1])!);
        }
        object?[] eighthMidPast = [77, null];
        TestAssert.Equal(false, (bool)eighthMid.Invoke(null, eighthMidPast)!);
        // Eighteenth-regime overflow-early replay (slots 85-88), dash108/dash109 agree bit-identical.
        int[] overflowEarlyIdx = [84, 85, 86, 87];
        int[] overflowEarlyR = [170, 217, 198, 186];
        int[] overflowEarlyG = [186, 170, 214, 176];
        int[] overflowEarlyB = [215, 169, 172, 201];
        for (int slot = 0; slot < overflowEarlyIdx.Length; slot++)
        {
            object?[] overflowEarlyArgs = [overflowEarlyIdx[slot], null];
            TestAssert.Equal(true, (bool)overflowEarly!.Invoke(null, overflowEarlyArgs)!);
            TestAssert.Equal((byte)overflowEarlyR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflowEarlyArgs[1])!);
            TestAssert.Equal((byte)overflowEarlyG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflowEarlyArgs[1])!);
            TestAssert.Equal((byte)overflowEarlyB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflowEarlyArgs[1])!);
        }
        object?[] overflowEarlyPast = [83, null];
        TestAssert.Equal(false, (bool)overflowEarly.Invoke(null, overflowEarlyPast)!);
        // Eighteenth-regime sixth replay (slots 89-90 take sixth slots 35-36), dash108/dash109 agree bit-identical.
        int[] sixthIdx = [88, 89];
        int[] sixthR = [169, 250];
        int[] sixthG = [206, 195];
        int[] sixthB = [220, 168];
        for (int slot = 0; slot < sixthIdx.Length; slot++)
        {
            object?[] sixthArgs = [sixthIdx[slot], null];
            TestAssert.Equal(true, (bool)sixth!.Invoke(null, sixthArgs)!);
            TestAssert.Equal((byte)sixthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixthArgs[1])!);
            TestAssert.Equal((byte)sixthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixthArgs[1])!);
            TestAssert.Equal((byte)sixthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixthArgs[1])!);
        }
        object?[] sixthPast = [87, null];
        TestAssert.Equal(false, (bool)sixth.Invoke(null, sixthPast)!);
        // Eighteenth-regime eighth-late replay (slots 91-96 take eighth slots 43-48), dash108/dash109 agree bit-identical.
        int[] eighthLateIdx = [90, 91, 92, 93, 94, 95];
        int[] eighthLateR = [182, 221, 205, 195, 181, 250];
        int[] eighthLateG = [195, 182, 219, 186, 212, 203];
        int[] eighthLateB = [220, 181, 184, 208, 224, 180];
        for (int slot = 0; slot < eighthLateIdx.Length; slot++)
        {
            object?[] eighthLateArgs = [eighthLateIdx[slot], null];
            TestAssert.Equal(true, (bool)eighthLate!.Invoke(null, eighthLateArgs)!);
            TestAssert.Equal((byte)eighthLateR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighthLateArgs[1])!);
            TestAssert.Equal((byte)eighthLateG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighthLateArgs[1])!);
            TestAssert.Equal((byte)eighthLateB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighthLateArgs[1])!);
        }
        object?[] eighthLatePast = [89, null];
        TestAssert.Equal(false, (bool)eighthLate.Invoke(null, eighthLatePast)!);
        // Eighteenth-regime eleventh-late replay (slots 97-100 take eleventh slots 61-64), dash108/dash109 agree bit-identical.
        int[] eleventhLateIdx = [96, 97, 98, 99];
        int[] eleventhLateR = [191, 225, 211, 202];
        int[] eleventhLateG = [203, 191, 223, 195];
        int[] eleventhLateB = [224, 191, 193, 213];
        for (int slot = 0; slot < eleventhLateIdx.Length; slot++)
        {
            object?[] eleventhLateArgs = [eleventhLateIdx[slot], null];
            TestAssert.Equal(true, (bool)eleventhLate!.Invoke(null, eleventhLateArgs)!);
            TestAssert.Equal((byte)eleventhLateR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventhLateArgs[1])!);
            TestAssert.Equal((byte)eleventhLateG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventhLateArgs[1])!);
            TestAssert.Equal((byte)eleventhLateB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventhLateArgs[1])!);
        }
        object?[] eleventhLatePast = [95, null];
        TestAssert.Equal(false, (bool)eleventhLate.Invoke(null, eleventhLatePast)!);
        // Eighteenth-regime twelfth71 replay (slot 101), dash108/dash109 agree bit-identical.
        int[] twelfth71Idx = [100];
        int[] twelfth71R = [193];
        int[] twelfth71G = [219];
        int[] twelfth71B = [229];
        for (int slot = 0; slot < twelfth71Idx.Length; slot++)
        {
            object?[] twelfth71Args = [twelfth71Idx[slot], null];
            TestAssert.Equal(true, (bool)twelfth71!.Invoke(null, twelfth71Args)!);
            TestAssert.Equal((byte)twelfth71R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfth71Args[1])!);
            TestAssert.Equal((byte)twelfth71G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfth71Args[1])!);
            TestAssert.Equal((byte)twelfth71B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfth71Args[1])!);
        }
        object?[] twelfth71Past = [99, null];
        TestAssert.Equal(false, (bool)twelfth71.Invoke(null, twelfth71Past)!);
        // Eighteenth-regime eleventh replay (slot 102), dash108/dash109 agree bit-identical.
        int[] eleventhSingleIdx = [101];
        int[] eleventhSingleR = [251];
        int[] eleventhSingleG = [209];
        int[] eleventhSingleB = [189];
        for (int slot = 0; slot < eleventhSingleIdx.Length; slot++)
        {
            object?[] eleventhSingleArgs = [eleventhSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)eleventhSingle!.Invoke(null, eleventhSingleArgs)!);
            TestAssert.Equal((byte)eleventhSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventhSingleArgs[1])!);
            TestAssert.Equal((byte)eleventhSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventhSingleArgs[1])!);
            TestAssert.Equal((byte)eleventhSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventhSingleArgs[1])!);
        }
        object?[] eleventhSinglePast = [100, null];
        TestAssert.Equal(false, (bool)eleventhSingle.Invoke(null, eleventhSinglePast)!);
        // Eighteenth-regime overflow-late replay (slots 103-106), dash108/dash109 agree bit-identical.
        int[] overflowLateIdx = [102, 103, 104, 105];
        int[] overflowLateR = [203, 230, 218, 212];
        int[] overflowLateG = [212, 203, 228, 206];
        int[] overflowLateB = [229, 202, 204, 220];
        for (int slot = 0; slot < overflowLateIdx.Length; slot++)
        {
            object?[] overflowLateArgs = [overflowLateIdx[slot], null];
            TestAssert.Equal(true, (bool)overflowLate!.Invoke(null, overflowLateArgs)!);
            TestAssert.Equal((byte)overflowLateR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflowLateArgs[1])!);
            TestAssert.Equal((byte)overflowLateG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflowLateArgs[1])!);
            TestAssert.Equal((byte)overflowLateB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflowLateArgs[1])!);
        }
        object?[] overflowLatePast = [101, null];
        TestAssert.Equal(false, (bool)overflowLate.Invoke(null, overflowLatePast)!);
        // Eighteenth-regime sixteenth replay (slot 108), dash108/dash109 agree bit-identical.
        int[] sixteenthSingleIdx = [107];
        int[] sixteenthSingleR = [251];
        int[] sixteenthSingleG = [215];
        int[] sixteenthSingleB = [199];
        for (int slot = 0; slot < sixteenthSingleIdx.Length; slot++)
        {
            object?[] sixteenthSingleArgs = [sixteenthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)sixteenthSingle!.Invoke(null, sixteenthSingleArgs)!);
            TestAssert.Equal((byte)sixteenthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixteenthSingleArgs[1])!);
            TestAssert.Equal((byte)sixteenthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixteenthSingleArgs[1])!);
            TestAssert.Equal((byte)sixteenthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixteenthSingleArgs[1])!);
        }
        object?[] sixteenthSinglePast = [106, null];
        TestAssert.Equal(false, (bool)sixteenthSingle.Invoke(null, sixteenthSinglePast)!);
    }
    public static void PptxSyntheticNineteenthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseNineteenthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected nineteenth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var earlySingles = rendererType.GetMethod(
            "TryResolveNineteenthRegimeEarlySinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(earlySingles is not null, "Expected nineteenth-regime earlysingles table to remain inspectable by the Office evidence guard.");
        var midSingles = rendererType.GetMethod(
            "TryResolveNineteenthRegimeMidSinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(midSingles is not null, "Expected nineteenth-regime midsingles table to remain inspectable by the Office evidence guard.");
        var shadeRow = rendererType.GetMethod(
            "TryResolveNineteenthRegimeShadeRowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(shadeRow is not null, "Expected nineteenth-regime shaderow table to remain inspectable by the Office evidence guard.");
        var lateSingles = rendererType.GetMethod(
            "TryResolveNineteenthRegimeLateSinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(lateSingles is not null, "Expected nineteenth-regime latesingles table to remain inspectable by the Office evidence guard.");
        var fifteenthEarly = rendererType.GetMethod(
            "TryResolveNineteenthRegimeFifteenthEarlyFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifteenthEarly is not null, "Expected nineteenth-regime fifteenthearly table to remain inspectable by the Office evidence guard.");
        var seventhSingle = rendererType.GetMethod(
            "TryResolveNineteenthRegimeSeventhSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventhSingle is not null, "Expected nineteenth-regime seventhsingle table to remain inspectable by the Office evidence guard.");
        var seventeenthMid = rendererType.GetMethod(
            "TryResolveNineteenthRegimeSeventeenthMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenthMid is not null, "Expected nineteenth-regime seventeenthmid table to remain inspectable by the Office evidence guard.");
        var seventeenthRaw = rendererType.GetMethod(
            "TryResolveNineteenthRegimeSeventeenthRawFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenthRaw is not null, "Expected nineteenth-regime seventeenthraw table to remain inspectable by the Office evidence guard.");
        var seventeenthFixed = rendererType.GetMethod(
            "TryResolveNineteenthRegimeSeventeenthFixedFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenthFixed is not null, "Expected nineteenth-regime seventeenthfixed table to remain inspectable by the Office evidence guard.");
        var fifth = rendererType.GetMethod(
            "TryResolveNineteenthRegimeFifthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifth is not null, "Expected nineteenth-regime fifth table to remain inspectable by the Office evidence guard.");
        var eleventhEarly = rendererType.GetMethod(
            "TryResolveNineteenthRegimeEleventhEarlyFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventhEarly is not null, "Expected nineteenth-regime eleventhearly table to remain inspectable by the Office evidence guard.");
        var fourth = rendererType.GetMethod(
            "TryResolveNineteenthRegimeFourthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourth is not null, "Expected nineteenth-regime fourth table to remain inspectable by the Office evidence guard.");
        var fifthMid = rendererType.GetMethod(
            "TryResolveNineteenthRegimeFifthMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifthMid is not null, "Expected nineteenth-regime fifthmid table to remain inspectable by the Office evidence guard.");
        var eleventhMid = rendererType.GetMethod(
            "TryResolveNineteenthRegimeEleventhMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventhMid is not null, "Expected nineteenth-regime eleventhmid table to remain inspectable by the Office evidence guard.");
        var overflowSingle = rendererType.GetMethod(
            "TryResolveNineteenthRegimeOverflowSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflowSingle is not null, "Expected nineteenth-regime overflowsingle table to remain inspectable by the Office evidence guard.");
        var fourteenth = rendererType.GetMethod(
            "TryResolveNineteenthRegimeFourteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourteenth is not null, "Expected nineteenth-regime fourteenth table to remain inspectable by the Office evidence guard.");
        var overflowMid = rendererType.GetMethod(
            "TryResolveNineteenthRegimeOverflowMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflowMid is not null, "Expected nineteenth-regime overflowmid table to remain inspectable by the Office evidence guard.");
        var twelfthLate = rendererType.GetMethod(
            "TryResolveNineteenthRegimeTwelfthLateFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfthLate is not null, "Expected nineteenth-regime twelfthlate table to remain inspectable by the Office evidence guard.");
        var overflowLate = rendererType.GetMethod(
            "TryResolveNineteenthRegimeOverflowLateFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflowLate is not null, "Expected nineteenth-regime overflowlate table to remain inspectable by the Office evidence guard.");
        var eighteenthTailReplay = rendererType.GetMethod(
            "TryResolveNineteenthRegimeEighteenthTailReplayFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenthTailReplay is not null, "Expected nineteenth-regime eighteenthtailreplay table to remain inspectable by the Office evidence guard.");
        var sixteenthSingle = rendererType.GetMethod(
            "TryResolveNineteenthRegimeSixteenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixteenthSingle is not null, "Expected nineteenth-regime sixteenthsingle table to remain inspectable by the Office evidence guard.");
        var tail = rendererType.GetMethod(
            "TryResolveNineteenthRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected nineteenth-regime tail table to remain inspectable by the Office evidence guard.");
        // One-hundred-thirteen points and fewer keep earlier regimes; one-hundred-fourteen-plus take the nineteenth rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [113])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [114])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [115])!);
        // Nineteenth-regime early singles (slots 7-9/11-12), dash114/dash115 agree bit-identical.
        int[] earlySinglesIdx = [6, 7, 8, 10, 11];
        int[] earlySinglesR = [51, 131, 105, 49, 170];
        int[] earlySinglesG = [87, 52, 127, 117, 101];
        int[] earlySinglesB = [129, 50, 58, 135, 45];
        for (int slot = 0; slot < earlySinglesIdx.Length; slot++)
        {
            object?[] earlySinglesArgs = [earlySinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)earlySingles!.Invoke(null, earlySinglesArgs)!);
            TestAssert.Equal((byte)earlySinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(earlySinglesArgs[1])!);
            TestAssert.Equal((byte)earlySinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(earlySinglesArgs[1])!);
            TestAssert.Equal((byte)earlySinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(earlySinglesArgs[1])!);
        }
        object?[] earlySinglesPast = [9, null];
        TestAssert.Equal(false, (bool)earlySingles.Invoke(null, earlySinglesPast)!);
        // Nineteenth-regime mid singles (slots 14-16/20-21/24/26-30), dash114/dash115 agree bit-identical.
        int[] midSinglesIdx = [13, 14, 15, 19, 20, 23, 25, 26, 27, 28, 29];
        int[] midSinglesR = [140, 113, 92, 148, 119, 191, 156, 125, 103, 59, 201];
        int[] midSinglesG = [56, 137, 71, 60, 144, 115, 64, 152, 80, 139, 121];
        int[] midSinglesB = [54, 63, 118, 57, 67, 52, 61, 71, 131, 161, 55];
        for (int slot = 0; slot < midSinglesIdx.Length; slot++)
        {
            object?[] midSinglesArgs = [midSinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)midSingles!.Invoke(null, midSinglesArgs)!);
            TestAssert.Equal((byte)midSinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(midSinglesArgs[1])!);
            TestAssert.Equal((byte)midSinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(midSinglesArgs[1])!);
            TestAssert.Equal((byte)midSinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(midSinglesArgs[1])!);
        }
        object?[] midSinglesPast = [24, null];
        TestAssert.Equal(false, (bool)midSingles.Invoke(null, midSinglesPast)!);
        // Nineteenth-regime shade row (slots 36-42), dash114/dash115 agree bit-identical.
        int[] shadeRowIdx = [35, 36, 37, 38, 39, 40, 41];
        int[] shadeRowR = [211, 69, 170, 137, 113, 65, 219];
        int[] shadeRowG = [127, 114, 70, 165, 88, 152, 132];
        int[] shadeRowB = [58, 167, 67, 78, 143, 175, 61];
        for (int slot = 0; slot < shadeRowIdx.Length; slot++)
        {
            object?[] shadeRowArgs = [shadeRowIdx[slot], null];
            TestAssert.Equal(true, (bool)shadeRow!.Invoke(null, shadeRowArgs)!);
            TestAssert.Equal((byte)shadeRowR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(shadeRowArgs[1])!);
            TestAssert.Equal((byte)shadeRowG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(shadeRowArgs[1])!);
            TestAssert.Equal((byte)shadeRowB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(shadeRowArgs[1])!);
        }
        object?[] shadeRowPast = [34, null];
        TestAssert.Equal(false, (bool)shadeRow.Invoke(null, shadeRowPast)!);
        // Nineteenth-regime late singles (slots 44/46/48/54), dash114/dash115 agree bit-identical.
        int[] lateSinglesIdx = [43, 45, 47, 53];
        int[] lateSinglesR = [177, 117, 228, 236];
        int[] lateSinglesG = [73, 91, 138, 143];
        int[] lateSinglesB = [70, 149, 64, 66];
        for (int slot = 0; slot < lateSinglesIdx.Length; slot++)
        {
            object?[] lateSinglesArgs = [lateSinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)lateSingles!.Invoke(null, lateSinglesArgs)!);
            TestAssert.Equal((byte)lateSinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(lateSinglesArgs[1])!);
            TestAssert.Equal((byte)lateSinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(lateSinglesArgs[1])!);
            TestAssert.Equal((byte)lateSinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(lateSinglesArgs[1])!);
        }
        object?[] lateSinglesPast = [52, null];
        TestAssert.Equal(false, (bool)lateSingles.Invoke(null, lateSinglesPast)!);
        // Nineteenth-regime fifteenth-early replay (slots 19/23), dash114/dash115 agree bit-identical.
        int[] fifteenthEarlyIdx = [18, 22];
        int[] fifteenthEarlyR = [58, 55];
        int[] fifteenthEarlyG = [97, 130];
        int[] fifteenthEarlyB = [143, 150];
        for (int slot = 0; slot < fifteenthEarlyIdx.Length; slot++)
        {
            object?[] fifteenthEarlyArgs = [fifteenthEarlyIdx[slot], null];
            TestAssert.Equal(true, (bool)fifteenthEarly!.Invoke(null, fifteenthEarlyArgs)!);
            TestAssert.Equal((byte)fifteenthEarlyR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifteenthEarlyArgs[1])!);
            TestAssert.Equal((byte)fifteenthEarlyG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifteenthEarlyArgs[1])!);
            TestAssert.Equal((byte)fifteenthEarlyB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifteenthEarlyArgs[1])!);
        }
        object?[] fifteenthEarlyPast = [21, null];
        TestAssert.Equal(false, (bool)fifteenthEarly.Invoke(null, fifteenthEarlyPast)!);
        // Nineteenth-regime seventh replay (slot 25), dash114/dash115 agree bit-identical.
        int[] seventhSingleIdx = [24];
        int[] seventhSingleR = [62];
        int[] seventhSingleG = [102];
        int[] seventhSingleB = [151];
        for (int slot = 0; slot < seventhSingleIdx.Length; slot++)
        {
            object?[] seventhSingleArgs = [seventhSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)seventhSingle!.Invoke(null, seventhSingleArgs)!);
            TestAssert.Equal((byte)seventhSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventhSingleArgs[1])!);
            TestAssert.Equal((byte)seventhSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventhSingleArgs[1])!);
            TestAssert.Equal((byte)seventhSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventhSingleArgs[1])!);
        }
        object?[] seventhSinglePast = [23, null];
        TestAssert.Equal(false, (bool)seventhSingle.Invoke(null, seventhSinglePast)!);
        // Nineteenth-regime seventeenth-mid replay (slots 43/45/47-53), dash114/dash115 agree bit-identical.
        int[] seventeenthMidIdx = [42, 44, 46, 48, 49, 50, 51, 52];
        int[] seventeenthMidR = [71, 141, 68, 74, 181, 146, 121, 70];
        int[] seventeenthMidG = [117, 170, 156, 122, 75, 177, 94, 162];
        int[] seventeenthMidB = [172, 80, 180, 178, 72, 84, 153, 187];
        for (int slot = 0; slot < seventeenthMidIdx.Length; slot++)
        {
            object?[] seventeenthMidArgs = [seventeenthMidIdx[slot], null];
            TestAssert.Equal(true, (bool)seventeenthMid!.Invoke(null, seventeenthMidArgs)!);
            TestAssert.Equal((byte)seventeenthMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenthMidArgs[1])!);
            TestAssert.Equal((byte)seventeenthMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenthMidArgs[1])!);
            TestAssert.Equal((byte)seventeenthMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenthMidArgs[1])!);
        }
        object?[] seventeenthMidPast = [43, null];
        TestAssert.Equal(false, (bool)seventeenthMid.Invoke(null, seventeenthMidPast)!);
        // Nineteenth-regime seventeenth-raw replay (slots 55-60), dash114/dash115 agree bit-identical.
        int[] seventeenthRawIdx = [54, 55, 56, 57, 58, 59];
        int[] seventeenthRawR = [77, 189, 152, 126, 74, 243];
        int[] seventeenthRawG = [127, 78, 184, 98, 169, 147];
        int[] seventeenthRawB = [186, 75, 87, 159, 194, 69];
        for (int slot = 0; slot < seventeenthRawIdx.Length; slot++)
        {
            object?[] seventeenthRawArgs = [seventeenthRawIdx[slot], null];
            TestAssert.Equal(true, (bool)seventeenthRaw!.Invoke(null, seventeenthRawArgs)!);
            TestAssert.Equal((byte)seventeenthRawR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenthRawArgs[1])!);
            TestAssert.Equal((byte)seventeenthRawG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenthRawArgs[1])!);
            TestAssert.Equal((byte)seventeenthRawB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenthRawArgs[1])!);
        }
        object?[] seventeenthRawPast = [53, null];
        TestAssert.Equal(false, (bool)seventeenthRaw.Invoke(null, seventeenthRawPast)!);
        // Nineteenth-regime seventeenth-fixed replay (slots 61-66), dash114/dash115 agree bit-identical.
        int[] seventeenthFixedIdx = [60, 61, 62, 63, 64, 65];
        int[] seventeenthFixedR = [92, 194, 159, 134, 88, 247];
        int[] seventeenthFixedG = [135, 92, 190, 109, 175, 155];
        int[] seventeenthFixedB = [191, 90, 100, 166, 200, 84];
        for (int slot = 0; slot < seventeenthFixedIdx.Length; slot++)
        {
            object?[] seventeenthFixedArgs = [seventeenthFixedIdx[slot], null];
            TestAssert.Equal(true, (bool)seventeenthFixed!.Invoke(null, seventeenthFixedArgs)!);
            TestAssert.Equal((byte)seventeenthFixedR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenthFixedArgs[1])!);
            TestAssert.Equal((byte)seventeenthFixedG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenthFixedArgs[1])!);
            TestAssert.Equal((byte)seventeenthFixedB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenthFixedArgs[1])!);
        }
        object?[] seventeenthFixedPast = [59, null];
        TestAssert.Equal(false, (bool)seventeenthFixed.Invoke(null, seventeenthFixedPast)!);
        // Nineteenth-regime fifth replay (slots 67-72 take fifth slots 19-24), dash114/dash115 agree bit-identical.
        int[] fifthIdx = [66, 67, 68, 69, 70, 71];
        int[] fifthR = [115, 200, 169, 148, 112, 248];
        int[] fifthG = [148, 115, 195, 128, 183, 165];
        int[] fifthB = [197, 114, 121, 174, 205, 110];
        for (int slot = 0; slot < fifthIdx.Length; slot++)
        {
            object?[] fifthArgs = [fifthIdx[slot], null];
            TestAssert.Equal(true, (bool)fifth!.Invoke(null, fifthArgs)!);
            TestAssert.Equal((byte)fifthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifthArgs[1])!);
            TestAssert.Equal((byte)fifthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifthArgs[1])!);
            TestAssert.Equal((byte)fifthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifthArgs[1])!);
        }
        object?[] fifthPast = [65, null];
        TestAssert.Equal(false, (bool)fifth.Invoke(null, fifthPast)!);
        // Nineteenth-regime eleventh-early replay (slots 73-78 take eleventh slots 43-48), dash114/dash115 agree bit-identical.
        int[] eleventhEarlyIdx = [72, 73, 74, 75, 76, 77];
        int[] eleventhEarlyR = [131, 204, 176, 158, 129, 248];
        int[] eleventhEarlyG = [159, 131, 200, 141, 189, 173];
        int[] eleventhEarlyB = [202, 130, 135, 181, 209, 127];
        for (int slot = 0; slot < eleventhEarlyIdx.Length; slot++)
        {
            object?[] eleventhEarlyArgs = [eleventhEarlyIdx[slot], null];
            TestAssert.Equal(true, (bool)eleventhEarly!.Invoke(null, eleventhEarlyArgs)!);
            TestAssert.Equal((byte)eleventhEarlyR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventhEarlyArgs[1])!);
            TestAssert.Equal((byte)eleventhEarlyG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventhEarlyArgs[1])!);
            TestAssert.Equal((byte)eleventhEarlyB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventhEarlyArgs[1])!);
        }
        object?[] eleventhEarlyPast = [71, null];
        TestAssert.Equal(false, (bool)eleventhEarly.Invoke(null, eleventhEarlyPast)!);
        // Nineteenth-regime fourth replay (slots 79-84 take fourth slots 19-24), dash114/dash115 agree bit-identical.
        int[] fourthIdx = [78, 79, 80, 81, 82, 83];
        int[] fourthR = [147, 209, 185, 169, 145, 249];
        int[] fourthG = [169, 147, 205, 155, 195, 181];
        int[] fourthB = [207, 146, 150, 189, 213, 144];
        for (int slot = 0; slot < fourthIdx.Length; slot++)
        {
            object?[] fourthArgs = [fourthIdx[slot], null];
            TestAssert.Equal(true, (bool)fourth!.Invoke(null, fourthArgs)!);
            TestAssert.Equal((byte)fourthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourthArgs[1])!);
            TestAssert.Equal((byte)fourthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourthArgs[1])!);
            TestAssert.Equal((byte)fourthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourthArgs[1])!);
        }
        object?[] fourthPast = [77, null];
        TestAssert.Equal(false, (bool)fourth.Invoke(null, fourthPast)!);
        // Nineteenth-regime fifth-mid replay (slots 85-90 take fifth slots 25-30), dash114/dash115 agree bit-identical.
        int[] fifthMidIdx = [84, 85, 86, 87, 88, 89];
        int[] fifthMidR = [161, 214, 192, 179, 160, 249];
        int[] fifthMidG = [180, 161, 210, 168, 202, 190];
        int[] fifthMidB = [212, 160, 164, 196, 217, 158];
        for (int slot = 0; slot < fifthMidIdx.Length; slot++)
        {
            object?[] fifthMidArgs = [fifthMidIdx[slot], null];
            TestAssert.Equal(true, (bool)fifthMid!.Invoke(null, fifthMidArgs)!);
            TestAssert.Equal((byte)fifthMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifthMidArgs[1])!);
            TestAssert.Equal((byte)fifthMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifthMidArgs[1])!);
            TestAssert.Equal((byte)fifthMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifthMidArgs[1])!);
        }
        object?[] fifthMidPast = [83, null];
        TestAssert.Equal(false, (bool)fifthMid.Invoke(null, fifthMidPast)!);
        // Nineteenth-regime eleventh-mid replay (slots 91-93/95-96), dash114/dash115 agree bit-identical.
        int[] eleventhMidIdx = [90, 91, 92, 94, 95];
        int[] eleventhMidR = [173, 218, 200, 172, 250];
        int[] eleventhMidG = [189, 173, 215, 208, 198];
        int[] eleventhMidB = [217, 173, 176, 221, 171];
        for (int slot = 0; slot < eleventhMidIdx.Length; slot++)
        {
            object?[] eleventhMidArgs = [eleventhMidIdx[slot], null];
            TestAssert.Equal(true, (bool)eleventhMid!.Invoke(null, eleventhMidArgs)!);
            TestAssert.Equal((byte)eleventhMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventhMidArgs[1])!);
            TestAssert.Equal((byte)eleventhMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventhMidArgs[1])!);
            TestAssert.Equal((byte)eleventhMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventhMidArgs[1])!);
        }
        object?[] eleventhMidPast = [93, null];
        TestAssert.Equal(false, (bool)eleventhMid.Invoke(null, eleventhMidPast)!);
        // Nineteenth-regime overflow replay (slot 94), dash114/dash115 agree bit-identical.
        int[] overflowSingleIdx = [93];
        int[] overflowSingleR = [186];
        int[] overflowSingleG = [176];
        int[] overflowSingleB = [201];
        for (int slot = 0; slot < overflowSingleIdx.Length; slot++)
        {
            object?[] overflowSingleArgs = [overflowSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)overflowSingle!.Invoke(null, overflowSingleArgs)!);
            TestAssert.Equal((byte)overflowSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflowSingleArgs[1])!);
            TestAssert.Equal((byte)overflowSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflowSingleArgs[1])!);
            TestAssert.Equal((byte)overflowSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflowSingleArgs[1])!);
        }
        object?[] overflowSinglePast = [92, null];
        TestAssert.Equal(false, (bool)overflowSingle.Invoke(null, overflowSinglePast)!);
        // Nineteenth-regime fourteenth replay (slots 97-102 take fourteenth slots 73-78), dash114/dash115 agree bit-identical.
        int[] fourteenthIdx = [96, 97, 98, 99, 100, 101];
        int[] fourteenthR = [183, 222, 206, 196, 182, 250];
        int[] fourteenthG = [196, 183, 220, 188, 213, 204];
        int[] fourteenthB = [221, 183, 185, 209, 225, 181];
        for (int slot = 0; slot < fourteenthIdx.Length; slot++)
        {
            object?[] fourteenthArgs = [fourteenthIdx[slot], null];
            TestAssert.Equal(true, (bool)fourteenth!.Invoke(null, fourteenthArgs)!);
            TestAssert.Equal((byte)fourteenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourteenthArgs[1])!);
            TestAssert.Equal((byte)fourteenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourteenthArgs[1])!);
            TestAssert.Equal((byte)fourteenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourteenthArgs[1])!);
        }
        object?[] fourteenthPast = [95, null];
        TestAssert.Equal(false, (bool)fourteenth.Invoke(null, fourteenthPast)!);
        // Nineteenth-regime overflow-mid replay (slots 103-106), dash114/dash115 agree bit-identical.
        int[] overflowMidIdx = [102, 103, 104, 105];
        int[] overflowMidR = [194, 226, 213, 205];
        int[] overflowMidG = [205, 194, 224, 198];
        int[] overflowMidB = [225, 194, 196, 215];
        for (int slot = 0; slot < overflowMidIdx.Length; slot++)
        {
            object?[] overflowMidArgs = [overflowMidIdx[slot], null];
            TestAssert.Equal(true, (bool)overflowMid!.Invoke(null, overflowMidArgs)!);
            TestAssert.Equal((byte)overflowMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflowMidArgs[1])!);
            TestAssert.Equal((byte)overflowMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflowMidArgs[1])!);
            TestAssert.Equal((byte)overflowMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflowMidArgs[1])!);
        }
        object?[] overflowMidPast = [101, null];
        TestAssert.Equal(false, (bool)overflowMid.Invoke(null, overflowMidPast)!);
        // Nineteenth-regime twelfth-late replay (slots 107-108), dash114/dash115 agree bit-identical.
        int[] twelfthLateIdx = [106, 107];
        int[] twelfthLateR = [193, 251];
        int[] twelfthLateG = [219, 211];
        int[] twelfthLateB = [229, 193];
        for (int slot = 0; slot < twelfthLateIdx.Length; slot++)
        {
            object?[] twelfthLateArgs = [twelfthLateIdx[slot], null];
            TestAssert.Equal(true, (bool)twelfthLate!.Invoke(null, twelfthLateArgs)!);
            TestAssert.Equal((byte)twelfthLateR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfthLateArgs[1])!);
            TestAssert.Equal((byte)twelfthLateG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfthLateArgs[1])!);
            TestAssert.Equal((byte)twelfthLateB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfthLateArgs[1])!);
        }
        object?[] twelfthLatePast = [105, null];
        TestAssert.Equal(false, (bool)twelfthLate.Invoke(null, twelfthLatePast)!);
        // Nineteenth-regime overflow-late replay (slots 109-112), dash114/dash115 agree bit-identical.
        int[] overflowLateIdx = [108, 109, 110, 111];
        int[] overflowLateR = [203, 230, 218, 212];
        int[] overflowLateG = [212, 203, 228, 206];
        int[] overflowLateB = [229, 202, 204, 220];
        for (int slot = 0; slot < overflowLateIdx.Length; slot++)
        {
            object?[] overflowLateArgs = [overflowLateIdx[slot], null];
            TestAssert.Equal(true, (bool)overflowLate!.Invoke(null, overflowLateArgs)!);
            TestAssert.Equal((byte)overflowLateR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflowLateArgs[1])!);
            TestAssert.Equal((byte)overflowLateG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflowLateArgs[1])!);
            TestAssert.Equal((byte)overflowLateB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflowLateArgs[1])!);
        }
        object?[] overflowLatePast = [107, null];
        TestAssert.Equal(false, (bool)overflowLate.Invoke(null, overflowLatePast)!);
        // Nineteenth-regime eighteenth-tail replay (slot 113), dash114/dash115 agree bit-identical.
        int[] eighteenthTailReplayIdx = [112];
        int[] eighteenthTailReplayR = [202];
        int[] eighteenthTailReplayG = [223];
        int[] eighteenthTailReplayB = [232];
        for (int slot = 0; slot < eighteenthTailReplayIdx.Length; slot++)
        {
            object?[] eighteenthTailReplayArgs = [eighteenthTailReplayIdx[slot], null];
            TestAssert.Equal(true, (bool)eighteenthTailReplay!.Invoke(null, eighteenthTailReplayArgs)!);
            TestAssert.Equal((byte)eighteenthTailReplayR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenthTailReplayArgs[1])!);
            TestAssert.Equal((byte)eighteenthTailReplayG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenthTailReplayArgs[1])!);
            TestAssert.Equal((byte)eighteenthTailReplayB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenthTailReplayArgs[1])!);
        }
        object?[] eighteenthTailReplayPast = [111, null];
        TestAssert.Equal(false, (bool)eighteenthTailReplay.Invoke(null, eighteenthTailReplayPast)!);
        // Nineteenth-regime sixteenth replay (slot 114), dash114/dash115 agree bit-identical.
        int[] sixteenthSingleIdx = [113];
        int[] sixteenthSingleR = [251];
        int[] sixteenthSingleG = [215];
        int[] sixteenthSingleB = [199];
        for (int slot = 0; slot < sixteenthSingleIdx.Length; slot++)
        {
            object?[] sixteenthSingleArgs = [sixteenthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)sixteenthSingle!.Invoke(null, sixteenthSingleArgs)!);
            TestAssert.Equal((byte)sixteenthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixteenthSingleArgs[1])!);
            TestAssert.Equal((byte)sixteenthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixteenthSingleArgs[1])!);
            TestAssert.Equal((byte)sixteenthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixteenthSingleArgs[1])!);
        }
        object?[] sixteenthSinglePast = [112, null];
        TestAssert.Equal(false, (bool)sixteenthSingle.Invoke(null, sixteenthSinglePast)!);
        // Nineteenth-regime tail (slots 115-118), dash115/dash116 agree bit-identical.
        int[] tailIdx = [114, 115, 116, 117];
        int[] tailR = [211, 234, 224, 218];
        int[] tailG = [218, 211, 232, 213];
        int[] tailB = [233, 211, 212, 226];
        for (int slot = 0; slot < tailIdx.Length; slot++)
        {
            object?[] tailArgs = [tailIdx[slot], null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] tailPast = [113, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailPast)!);
    }
    public static void PptxSyntheticTwentiethVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwentiethVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twentieth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var earlySingles = rendererType.GetMethod(
            "TryResolveTwentiethRegimeEarlySinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(earlySingles is not null, "Expected twentieth-regime earlysingles table to remain inspectable by the Office evidence guard.");
        var eighteenthFixed = rendererType.GetMethod(
            "TryResolveTwentiethRegimeEighteenthFixedFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenthFixed is not null, "Expected twentieth-regime eighteenthfixed table to remain inspectable by the Office evidence guard.");
        var eighteenthMid = rendererType.GetMethod(
            "TryResolveTwentiethRegimeEighteenthMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenthMid is not null, "Expected twentieth-regime eighteenthmid table to remain inspectable by the Office evidence guard.");
        var eighteenthTailReplay = rendererType.GetMethod(
            "TryResolveTwentiethRegimeEighteenthTailReplayFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenthTailReplay is not null, "Expected twentieth-regime eighteenthtailreplay table to remain inspectable by the Office evidence guard.");
        var eleventh = rendererType.GetMethod(
            "TryResolveTwentiethRegimeEleventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventh is not null, "Expected twentieth-regime eleventh table to remain inspectable by the Office evidence guard.");
        var fifteenth = rendererType.GetMethod(
            "TryResolveTwentiethRegimeFifteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifteenth is not null, "Expected twentieth-regime fifteenth table to remain inspectable by the Office evidence guard.");
        var midSingles = rendererType.GetMethod(
            "TryResolveTwentiethRegimeMidSinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(midSingles is not null, "Expected twentieth-regime midsingles table to remain inspectable by the Office evidence guard.");
        var ninth = rendererType.GetMethod(
            "TryResolveTwentiethRegimeNinthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninth is not null, "Expected twentieth-regime ninth table to remain inspectable by the Office evidence guard.");
        var ninthLate = rendererType.GetMethod(
            "TryResolveTwentiethRegimeNinthLateFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninthLate is not null, "Expected twentieth-regime ninthlate table to remain inspectable by the Office evidence guard.");
        var ninthMid = rendererType.GetMethod(
            "TryResolveTwentiethRegimeNinthMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninthMid is not null, "Expected twentieth-regime ninthmid table to remain inspectable by the Office evidence guard.");
        var ninthOverflow = rendererType.GetMethod(
            "TryResolveTwentiethRegimeNinthOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninthOverflow is not null, "Expected twentieth-regime ninthoverflow table to remain inspectable by the Office evidence guard.");
        var ninthReplay = rendererType.GetMethod(
            "TryResolveTwentiethRegimeNinthReplayFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninthReplay is not null, "Expected twentieth-regime ninthreplay table to remain inspectable by the Office evidence guard.");
        var overflow = rendererType.GetMethod(
            "TryResolveTwentiethRegimeOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflow is not null, "Expected twentieth-regime overflow table to remain inspectable by the Office evidence guard.");
        var seventeenthLate = rendererType.GetMethod(
            "TryResolveTwentiethRegimeSeventeenthLateFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenthLate is not null, "Expected twentieth-regime seventeenthlate table to remain inspectable by the Office evidence guard.");
        var seventeenthMid = rendererType.GetMethod(
            "TryResolveTwentiethRegimeSeventeenthMidFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenthMid is not null, "Expected twentieth-regime seventeenthmid table to remain inspectable by the Office evidence guard.");
        var seventeenthSingle = rendererType.GetMethod(
            "TryResolveTwentiethRegimeSeventeenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenthSingle is not null, "Expected twentieth-regime seventeenthsingle table to remain inspectable by the Office evidence guard.");
        var seventh = rendererType.GetMethod(
            "TryResolveTwentiethRegimeSeventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventh is not null, "Expected twentieth-regime seventh table to remain inspectable by the Office evidence guard.");
        var seventhSingle = rendererType.GetMethod(
            "TryResolveTwentiethRegimeSeventhSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventhSingle is not null, "Expected twentieth-regime seventhsingle table to remain inspectable by the Office evidence guard.");
        var tailSingle = rendererType.GetMethod(
            "TryResolveTwentiethRegimeTailSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tailSingle is not null, "Expected twentieth-regime tailsingle table to remain inspectable by the Office evidence guard.");
        var twelfth = rendererType.GetMethod(
            "TryResolveTwentiethRegimeTwelfthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfth is not null, "Expected twentieth-regime twelfth table to remain inspectable by the Office evidence guard.");
        // One-hundred-nineteen points and fewer keep earlier regimes; one-hundred-twenty-plus take the twentieth rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [119])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [120])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [121])!);
        // Twentieth-regime EarlySingles (slots 10,17,18,22), dash120/dash121 agree bit-identical.
        int[] earlySinglesIdx = [9, 16, 17, 21];
        int[] earlySinglesR = [85, 52, 180, 97];
        int[] earlySinglesG = [65, 124, 108, 75];
        int[] earlySinglesB = [109, 143, 48, 123];
        for (int slot = 0; slot < earlySinglesIdx.Length; slot++)
        {
            object?[] earlySinglesArgs = [earlySinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)earlySingles!.Invoke(null, earlySinglesArgs)!);
            TestAssert.Equal((byte)earlySinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(earlySinglesArgs[1])!);
            TestAssert.Equal((byte)earlySinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(earlySinglesArgs[1])!);
            TestAssert.Equal((byte)earlySinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(earlySinglesArgs[1])!);
        }
        object?[] earlySinglesPast = [22, null];
        TestAssert.Equal(false, (bool)earlySingles.Invoke(null, earlySinglesPast)!);
        // Twentieth-regime EighteenthFixed (slots 61,62,63,64,65,66), dash120/dash121 agree bit-identical.
        int[] eighteenthFixedIdx = [60, 61, 62, 63, 64, 65];
        int[] eighteenthFixedR = [79, 192, 155, 128, 75, 247];
        int[] eighteenthFixedG = [129, 80, 187, 100, 172, 150];
        int[] eighteenthFixedB = [189, 77, 89, 162, 198, 70];
        for (int slot = 0; slot < eighteenthFixedIdx.Length; slot++)
        {
            object?[] eighteenthFixedArgs = [eighteenthFixedIdx[slot], null];
            TestAssert.Equal(true, (bool)eighteenthFixed!.Invoke(null, eighteenthFixedArgs)!);
            TestAssert.Equal((byte)eighteenthFixedR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenthFixedArgs[1])!);
            TestAssert.Equal((byte)eighteenthFixedG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenthFixedArgs[1])!);
            TestAssert.Equal((byte)eighteenthFixedB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenthFixedArgs[1])!);
        }
        object?[] eighteenthFixedPast = [59, null];
        TestAssert.Equal(false, (bool)eighteenthFixed.Invoke(null, eighteenthFixedPast)!);
        // Twentieth-regime EighteenthMid (slots 55,56,57,59,60), dash120/dash121 agree bit-identical.
        int[] eighteenthMidIdx = [54, 55, 56, 58, 59];
        int[] eighteenthMidR = [76, 186, 150, 72, 239];
        int[] eighteenthMidG = [125, 77, 181, 166, 145];
        int[] eighteenthMidB = [183, 74, 86, 192, 67];
        for (int slot = 0; slot < eighteenthMidIdx.Length; slot++)
        {
            object?[] eighteenthMidArgs = [eighteenthMidIdx[slot], null];
            TestAssert.Equal(true, (bool)eighteenthMid!.Invoke(null, eighteenthMidArgs)!);
            TestAssert.Equal((byte)eighteenthMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenthMidArgs[1])!);
            TestAssert.Equal((byte)eighteenthMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenthMidArgs[1])!);
            TestAssert.Equal((byte)eighteenthMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenthMidArgs[1])!);
        }
        object?[] eighteenthMidPast = [57, null];
        TestAssert.Equal(false, (bool)eighteenthMid.Invoke(null, eighteenthMidPast)!);
        // Twentieth-regime EighteenthTailReplay (slots 119), dash120/dash121 agree bit-identical.
        int[] eighteenthTailReplayIdx = [118];
        int[] eighteenthTailReplayR = [202];
        int[] eighteenthTailReplayG = [223];
        int[] eighteenthTailReplayB = [232];
        for (int slot = 0; slot < eighteenthTailReplayIdx.Length; slot++)
        {
            object?[] eighteenthTailReplayArgs = [eighteenthTailReplayIdx[slot], null];
            TestAssert.Equal(true, (bool)eighteenthTailReplay!.Invoke(null, eighteenthTailReplayArgs)!);
            TestAssert.Equal((byte)eighteenthTailReplayR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenthTailReplayArgs[1])!);
            TestAssert.Equal((byte)eighteenthTailReplayG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenthTailReplayArgs[1])!);
            TestAssert.Equal((byte)eighteenthTailReplayB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenthTailReplayArgs[1])!);
        }
        object?[] eighteenthTailReplayPast = [117, null];
        TestAssert.Equal(false, (bool)eighteenthTailReplay.Invoke(null, eighteenthTailReplayPast)!);
        // Twentieth-regime Eleventh (slots 85,86,87,88,89,90), dash120/dash121 agree bit-identical.
        int[] eleventhIdx = [84, 85, 86, 87, 88, 89];
        int[] eleventhR = [153, 211, 188, 173, 152, 249];
        int[] eleventhG = [174, 153, 208, 161, 198, 185];
        int[] eleventhB = [209, 152, 156, 192, 215, 150];
        for (int slot = 0; slot < eleventhIdx.Length; slot++)
        {
            object?[] eleventhArgs = [eleventhIdx[slot], null];
            TestAssert.Equal(true, (bool)eleventh!.Invoke(null, eleventhArgs)!);
            TestAssert.Equal((byte)eleventhR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventhArgs[1])!);
            TestAssert.Equal((byte)eleventhG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventhArgs[1])!);
            TestAssert.Equal((byte)eleventhB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventhArgs[1])!);
        }
        object?[] eleventhPast = [83, null];
        TestAssert.Equal(false, (bool)eleventh.Invoke(null, eleventhPast)!);
        // Twentieth-regime Fifteenth (slots 73,74,75,76,77,78), dash120/dash121 agree bit-identical.
        int[] fifteenthIdx = [72, 73, 74, 75, 76, 77];
        int[] fifteenthR = [120, 201, 171, 151, 118, 248];
        int[] fifteenthG = [152, 121, 197, 132, 185, 167];
        int[] fifteenthB = [198, 119, 126, 177, 206, 116];
        for (int slot = 0; slot < fifteenthIdx.Length; slot++)
        {
            object?[] fifteenthArgs = [fifteenthIdx[slot], null];
            TestAssert.Equal(true, (bool)fifteenth!.Invoke(null, fifteenthArgs)!);
            TestAssert.Equal((byte)fifteenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifteenthArgs[1])!);
            TestAssert.Equal((byte)fifteenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifteenthArgs[1])!);
            TestAssert.Equal((byte)fifteenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifteenthArgs[1])!);
        }
        object?[] fifteenthPast = [71, null];
        TestAssert.Equal(false, (bool)fifteenth.Invoke(null, fifteenthPast)!);
        // Twentieth-regime MidSingles (slots 31,33,34,35), dash120/dash121 agree bit-identical.
        int[] midSinglesIdx = [30, 32, 33, 34];
        int[] midSinglesR = [65, 130, 107, 62];
        int[] midSinglesG = [108, 157, 83, 144];
        int[] midSinglesB = [159, 74, 136, 167];
        for (int slot = 0; slot < midSinglesIdx.Length; slot++)
        {
            object?[] midSinglesArgs = [midSinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)midSingles!.Invoke(null, midSinglesArgs)!);
            TestAssert.Equal((byte)midSinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(midSinglesArgs[1])!);
            TestAssert.Equal((byte)midSinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(midSinglesArgs[1])!);
            TestAssert.Equal((byte)midSinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(midSinglesArgs[1])!);
        }
        object?[] midSinglesPast = [31, null];
        TestAssert.Equal(false, (bool)midSingles.Invoke(null, midSinglesPast)!);
        // Twentieth-regime Ninth (slots 67,68,69,70,71,72), dash120/dash121 agree bit-identical.
        int[] ninthIdx = [66, 67, 68, 69, 70, 71];
        int[] ninthR = [102, 197, 163, 140, 100, 247];
        int[] ninthG = [141, 103, 192, 118, 178, 159];
        int[] ninthB = [194, 101, 109, 170, 202, 96];
        for (int slot = 0; slot < ninthIdx.Length; slot++)
        {
            object?[] ninthArgs = [ninthIdx[slot], null];
            TestAssert.Equal(true, (bool)ninth!.Invoke(null, ninthArgs)!);
            TestAssert.Equal((byte)ninthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninthArgs[1])!);
            TestAssert.Equal((byte)ninthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninthArgs[1])!);
            TestAssert.Equal((byte)ninthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninthArgs[1])!);
        }
        object?[] ninthPast = [65, null];
        TestAssert.Equal(false, (bool)ninth.Invoke(null, ninthPast)!);
        // Twentieth-regime NinthLate (slots 91,92,93,94,95,96), dash120/dash121 agree bit-identical.
        int[] ninthLateIdx = [90, 91, 92, 93, 94, 95];
        int[] ninthLateR = [163, 214, 193, 180, 161, 250];
        int[] ninthLateG = [181, 163, 211, 169, 203, 191];
        int[] ninthLateB = [212, 162, 166, 197, 218, 160];
        for (int slot = 0; slot < ninthLateIdx.Length; slot++)
        {
            object?[] ninthLateArgs = [ninthLateIdx[slot], null];
            TestAssert.Equal(true, (bool)ninthLate!.Invoke(null, ninthLateArgs)!);
            TestAssert.Equal((byte)ninthLateR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninthLateArgs[1])!);
            TestAssert.Equal((byte)ninthLateG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninthLateArgs[1])!);
            TestAssert.Equal((byte)ninthLateB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninthLateArgs[1])!);
        }
        object?[] ninthLatePast = [89, null];
        TestAssert.Equal(false, (bool)ninthLate.Invoke(null, ninthLatePast)!);
        // Twentieth-regime NinthMid (slots 79,80,81,82,83,84), dash120/dash121 agree bit-identical.
        int[] ninthMidIdx = [78, 79, 80, 81, 82, 83];
        int[] ninthMidR = [138, 206, 180, 163, 136, 249];
        int[] ninthMidG = [163, 138, 202, 147, 192, 177];
        int[] ninthMidB = [204, 137, 142, 185, 210, 134];
        for (int slot = 0; slot < ninthMidIdx.Length; slot++)
        {
            object?[] ninthMidArgs = [ninthMidIdx[slot], null];
            TestAssert.Equal(true, (bool)ninthMid!.Invoke(null, ninthMidArgs)!);
            TestAssert.Equal((byte)ninthMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninthMidArgs[1])!);
            TestAssert.Equal((byte)ninthMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninthMidArgs[1])!);
            TestAssert.Equal((byte)ninthMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninthMidArgs[1])!);
        }
        object?[] ninthMidPast = [77, null];
        TestAssert.Equal(false, (bool)ninthMid.Invoke(null, ninthMidPast)!);
        // Twentieth-regime NinthOverflow (slots 103,104,105,106,107,108), dash120/dash121 agree bit-identical.
        int[] ninthOverflowIdx = [102, 103, 104, 105, 106, 107];
        int[] ninthOverflowR = [185, 223, 207, 197, 184, 251];
        int[] ninthOverflowG = [198, 185, 220, 189, 214, 205];
        int[] ninthOverflowB = [221, 184, 187, 210, 225, 183];
        for (int slot = 0; slot < ninthOverflowIdx.Length; slot++)
        {
            object?[] ninthOverflowArgs = [ninthOverflowIdx[slot], null];
            TestAssert.Equal(true, (bool)ninthOverflow!.Invoke(null, ninthOverflowArgs)!);
            TestAssert.Equal((byte)ninthOverflowR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninthOverflowArgs[1])!);
            TestAssert.Equal((byte)ninthOverflowG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninthOverflowArgs[1])!);
            TestAssert.Equal((byte)ninthOverflowB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninthOverflowArgs[1])!);
        }
        object?[] ninthOverflowPast = [101, null];
        TestAssert.Equal(false, (bool)ninthOverflow.Invoke(null, ninthOverflowPast)!);
        // Twentieth-regime NinthReplay (slots 115,116,117,118), dash120/dash121 agree bit-identical.
        int[] ninthReplayIdx = [114, 115, 116, 117];
        int[] ninthReplayR = [204, 231, 219, 213];
        int[] ninthReplayG = [213, 204, 229, 207];
        int[] ninthReplayB = [230, 204, 205, 221];
        for (int slot = 0; slot < ninthReplayIdx.Length; slot++)
        {
            object?[] ninthReplayArgs = [ninthReplayIdx[slot], null];
            TestAssert.Equal(true, (bool)ninthReplay!.Invoke(null, ninthReplayArgs)!);
            TestAssert.Equal((byte)ninthReplayR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninthReplayArgs[1])!);
            TestAssert.Equal((byte)ninthReplayG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninthReplayArgs[1])!);
            TestAssert.Equal((byte)ninthReplayB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninthReplayArgs[1])!);
        }
        object?[] ninthReplayPast = [113, null];
        TestAssert.Equal(false, (bool)ninthReplay.Invoke(null, ninthReplayPast)!);
        // Twentieth-regime Overflow (slots 109,110,111,112), dash120/dash121 agree bit-identical.
        int[] overflowIdx = [108, 109, 110, 111];
        int[] overflowR = [194, 226, 213, 205];
        int[] overflowG = [205, 194, 224, 198];
        int[] overflowB = [225, 194, 196, 215];
        for (int slot = 0; slot < overflowIdx.Length; slot++)
        {
            object?[] overflowArgs = [overflowIdx[slot], null];
            TestAssert.Equal(true, (bool)overflow!.Invoke(null, overflowArgs)!);
            TestAssert.Equal((byte)overflowR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflowArgs[1])!);
            TestAssert.Equal((byte)overflowG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflowArgs[1])!);
            TestAssert.Equal((byte)overflowB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflowArgs[1])!);
        }
        object?[] overflowPast = [107, null];
        TestAssert.Equal(false, (bool)overflow.Invoke(null, overflowPast)!);
        // Twentieth-regime SeventeenthLate (slots 50,51,53,54), dash120/dash121 agree bit-identical.
        int[] seventeenthLateIdx = [49, 50, 52, 53];
        int[] seventeenthLateR = [181, 146, 70, 233];
        int[] seventeenthLateG = [75, 177, 162, 141];
        int[] seventeenthLateB = [72, 84, 187, 66];
        for (int slot = 0; slot < seventeenthLateIdx.Length; slot++)
        {
            object?[] seventeenthLateArgs = [seventeenthLateIdx[slot], null];
            TestAssert.Equal(true, (bool)seventeenthLate!.Invoke(null, seventeenthLateArgs)!);
            TestAssert.Equal((byte)seventeenthLateR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenthLateArgs[1])!);
            TestAssert.Equal((byte)seventeenthLateG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenthLateArgs[1])!);
            TestAssert.Equal((byte)seventeenthLateB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenthLateArgs[1])!);
        }
        object?[] seventeenthLatePast = [51, null];
        TestAssert.Equal(false, (bool)seventeenthLate.Invoke(null, seventeenthLatePast)!);
        // Twentieth-regime SeventeenthMid (slots 42,43,44,45,46,47,48), dash120/dash121 agree bit-identical.
        int[] seventeenthMidIdx = [41, 42, 43, 44, 45, 46, 47];
        int[] seventeenthMidR = [216, 71, 175, 141, 116, 68, 225];
        int[] seventeenthMidG = [131, 117, 72, 170, 90, 156, 136];
        int[] seventeenthMidB = [60, 172, 69, 80, 147, 180, 63];
        for (int slot = 0; slot < seventeenthMidIdx.Length; slot++)
        {
            object?[] seventeenthMidArgs = [seventeenthMidIdx[slot], null];
            TestAssert.Equal(true, (bool)seventeenthMid!.Invoke(null, seventeenthMidArgs)!);
            TestAssert.Equal((byte)seventeenthMidR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenthMidArgs[1])!);
            TestAssert.Equal((byte)seventeenthMidG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenthMidArgs[1])!);
            TestAssert.Equal((byte)seventeenthMidB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenthMidArgs[1])!);
        }
        object?[] seventeenthMidPast = [40, null];
        TestAssert.Equal(false, (bool)seventeenthMid.Invoke(null, seventeenthMidPast)!);
        // Twentieth-regime SeventeenthSingle (slots 32,36), dash120/dash121 agree bit-identical.
        int[] seventeenthSingleIdx = [31, 35];
        int[] seventeenthSingleR = [159, 206];
        int[] seventeenthSingleG = [65, 124];
        int[] seventeenthSingleB = [63, 57];
        for (int slot = 0; slot < seventeenthSingleIdx.Length; slot++)
        {
            object?[] seventeenthSingleArgs = [seventeenthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)seventeenthSingle!.Invoke(null, seventeenthSingleArgs)!);
            TestAssert.Equal((byte)seventeenthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenthSingleArgs[1])!);
            TestAssert.Equal((byte)seventeenthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenthSingleArgs[1])!);
            TestAssert.Equal((byte)seventeenthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenthSingleArgs[1])!);
        }
        object?[] seventeenthSinglePast = [30, null];
        TestAssert.Equal(false, (bool)seventeenthSingle.Invoke(null, seventeenthSinglePast)!);
        // Twentieth-regime Seventh (slots 97,98,99,100,101,102), dash120/dash121 agree bit-identical.
        int[] seventhIdx = [96, 97, 98, 99, 100, 101];
        int[] seventhR = [175, 219, 201, 190, 174, 250];
        int[] seventhG = [190, 175, 216, 180, 209, 199];
        int[] seventhB = [217, 175, 177, 204, 222, 173];
        for (int slot = 0; slot < seventhIdx.Length; slot++)
        {
            object?[] seventhArgs = [seventhIdx[slot], null];
            TestAssert.Equal(true, (bool)seventh!.Invoke(null, seventhArgs)!);
            TestAssert.Equal((byte)seventhR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventhArgs[1])!);
            TestAssert.Equal((byte)seventhG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventhArgs[1])!);
            TestAssert.Equal((byte)seventhB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventhArgs[1])!);
        }
        object?[] seventhPast = [95, null];
        TestAssert.Equal(false, (bool)seventh.Invoke(null, seventhPast)!);
        // Twentieth-regime SeventhSingle (slots 30), dash120/dash121 agree bit-identical.
        int[] seventhSingleIdx = [29];
        int[] seventhSingleR = [198];
        int[] seventhSingleG = [119];
        int[] seventhSingleB = [54];
        for (int slot = 0; slot < seventhSingleIdx.Length; slot++)
        {
            object?[] seventhSingleArgs = [seventhSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)seventhSingle!.Invoke(null, seventhSingleArgs)!);
            TestAssert.Equal((byte)seventhSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventhSingleArgs[1])!);
            TestAssert.Equal((byte)seventhSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventhSingleArgs[1])!);
            TestAssert.Equal((byte)seventhSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventhSingleArgs[1])!);
        }
        object?[] seventhSinglePast = [28, null];
        TestAssert.Equal(false, (bool)seventhSingle.Invoke(null, seventhSinglePast)!);
        // Twentieth-regime TailSingle (slots 120-124), dash120/dash121/dash122/dash123/dash124 agree bit-identical.
        int[] tailSingleIdx = [119, 120, 121, 122, 123];
        int[] tailSingleR = [252, 211, 234, 224, 218];
        int[] tailSingleG = [218, 218, 211, 232, 213];
        int[] tailSingleB = [203, 233, 211, 212, 226];
        for (int slot = 0; slot < tailSingleIdx.Length; slot++)
        {
            object?[] tailSingleArgs = [tailSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)tailSingle!.Invoke(null, tailSingleArgs)!);
            TestAssert.Equal((byte)tailSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailSingleArgs[1])!);
            TestAssert.Equal((byte)tailSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailSingleArgs[1])!);
            TestAssert.Equal((byte)tailSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailSingleArgs[1])!);
        }
        object?[] tailSinglePast = [118, null];
        TestAssert.Equal(false, (bool)tailSingle.Invoke(null, tailSinglePast)!);
        // Twentieth-regime Twelfth (slots 113,114), dash120/dash121 agree bit-identical.
        int[] twelfthIdx = [112, 113];
        int[] twelfthR = [193, 251];
        int[] twelfthG = [219, 211];
        int[] twelfthB = [229, 193];
        for (int slot = 0; slot < twelfthIdx.Length; slot++)
        {
            object?[] twelfthArgs = [twelfthIdx[slot], null];
            TestAssert.Equal(true, (bool)twelfth!.Invoke(null, twelfthArgs)!);
            TestAssert.Equal((byte)twelfthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfthArgs[1])!);
            TestAssert.Equal((byte)twelfthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfthArgs[1])!);
            TestAssert.Equal((byte)twelfthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfthArgs[1])!);
        }
        object?[] twelfthPast = [111, null];
        TestAssert.Equal(false, (bool)twelfth.Invoke(null, twelfthPast)!);
    }
    public static void PptxSyntheticTwentyFirstVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwentyFirstVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twenty-first-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var earlySingles = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeEarlySinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(earlySingles is not null, "Expected twenty-first-regime earlysingles table to remain inspectable by the Office evidence guard.");
        var eighteenth = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeEighteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenth is not null, "Expected twenty-first-regime eighteenth table to remain inspectable by the Office evidence guard.");
        var eighth = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeEighthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighth is not null, "Expected twenty-first-regime eighth table to remain inspectable by the Office evidence guard.");
        var fifteenth = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeFifteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifteenth is not null, "Expected twenty-first-regime fifteenth table to remain inspectable by the Office evidence guard.");
        var fourteenthSingle = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeFourteenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourteenthSingle is not null, "Expected twenty-first-regime fourteenthsingle table to remain inspectable by the Office evidence guard.");
        var lateSingles = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeLateSinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(lateSingles is not null, "Expected twenty-first-regime latesingles table to remain inspectable by the Office evidence guard.");
        var nineteenth = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeNineteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(nineteenth is not null, "Expected twenty-first-regime nineteenth table to remain inspectable by the Office evidence guard.");
        var ninthReplay = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeNinthReplayFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninthReplay is not null, "Expected twenty-first-regime ninthreplay table to remain inspectable by the Office evidence guard.");
        var seventeenth = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeSeventeenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenth is not null, "Expected twenty-first-regime seventeenth table to remain inspectable by the Office evidence guard.");
        var seventh = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeSeventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventh is not null, "Expected twenty-first-regime seventh table to remain inspectable by the Office evidence guard.");
        var sixteenth = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeSixteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixteenth is not null, "Expected twenty-first-regime sixteenth table to remain inspectable by the Office evidence guard.");
        var thirteenth = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeThirteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(thirteenth is not null, "Expected twenty-first-regime thirteenth table to remain inspectable by the Office evidence guard.");
        var twelfth = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeTwelfthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfth is not null, "Expected twenty-first-regime twelfth table to remain inspectable by the Office evidence guard.");
        var twentieth = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeTwentiethFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentieth is not null, "Expected twenty-first-regime twentieth table to remain inspectable by the Office evidence guard.");
        // One-hundred-twenty-five points and fewer keep earlier regimes; one-hundred-twenty-six-plus take the twenty-first rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [125])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [126])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [127])!);
        // Twenty-first-regime EarlySingles (slots 49,51,53), dash126/dash127 agree bit-identical.
        int[] earlySinglesIdx = [48, 50, 52];
        int[] earlySinglesR = [72, 143, 69];
        int[] earlySinglesG = [119, 173, 159];
        int[] earlySinglesB = [175, 82, 183];
        for (int slot = 0; slot < earlySinglesIdx.Length; slot++)
        {
            object?[] earlySinglesArgs = [earlySinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)earlySingles!.Invoke(null, earlySinglesArgs)!);
            TestAssert.Equal((byte)earlySinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(earlySinglesArgs[1])!);
            TestAssert.Equal((byte)earlySinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(earlySinglesArgs[1])!);
            TestAssert.Equal((byte)earlySinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(earlySinglesArgs[1])!);
        }
        object?[] earlySinglesPast = [22, null];
        TestAssert.Equal(false, (bool)earlySingles.Invoke(null, earlySinglesPast)!);
        // Twenty-first-regime Eighteenth (slots 37,38,39,40,41,86,125), dash126/dash127 agree bit-identical.
        int[] eighteenthIdx = [36, 37, 38, 39, 40, 85, 124];
        int[] eighteenthR = [67, 166, 133, 110, 64, 207, 202];
        int[] eighteenthG = [111, 68, 161, 85, 148, 143, 223];
        int[] eighteenthB = [163, 65, 76, 139, 171, 142, 232];
        for (int slot = 0; slot < eighteenthIdx.Length; slot++)
        {
            object?[] eighteenthArgs = [eighteenthIdx[slot], null];
            TestAssert.Equal(true, (bool)eighteenth!.Invoke(null, eighteenthArgs)!);
            TestAssert.Equal((byte)eighteenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenthArgs[1])!);
            TestAssert.Equal((byte)eighteenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenthArgs[1])!);
            TestAssert.Equal((byte)eighteenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenthArgs[1])!);
        }
        object?[] eighteenthPast = [35, null];
        TestAssert.Equal(false, (bool)eighteenth.Invoke(null, eighteenthPast)!);
        // Twenty-first-regime Eighth (slots 91,92,93,94,95,96), dash126/dash127 agree bit-identical.
        int[] eighthIdx = [90, 91, 92, 93, 94, 95];
        int[] eighthR = [157, 212, 190, 176, 156, 249];
        int[] eighthG = [177, 157, 209, 164, 200, 187];
        int[] eighthB = [210, 156, 160, 194, 216, 154];
        for (int slot = 0; slot < eighthIdx.Length; slot++)
        {
            object?[] eighthArgs = [eighthIdx[slot], null];
            TestAssert.Equal(true, (bool)eighth!.Invoke(null, eighthArgs)!);
            TestAssert.Equal((byte)eighthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighthArgs[1])!);
            TestAssert.Equal((byte)eighthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighthArgs[1])!);
            TestAssert.Equal((byte)eighthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighthArgs[1])!);
        }
        object?[] eighthPast = [89, null];
        TestAssert.Equal(false, (bool)eighth.Invoke(null, eighthPast)!);
        // Twenty-first-regime Fifteenth (slots 20,21,24,109,110,111,112,113,114), dash126/dash127 agree bit-identical.
        int[] fifteenthIdx = [19, 20, 23, 108, 109, 110, 111, 112, 113];
        int[] fifteenthR = [145, 117, 188, 186, 223, 208, 199, 186, 251];
        int[] fifteenthG = [59, 141, 113, 199, 187, 221, 191, 215, 206];
        int[] fifteenthB = [56, 66, 51, 222, 186, 188, 211, 226, 185];
        for (int slot = 0; slot < fifteenthIdx.Length; slot++)
        {
            object?[] fifteenthArgs = [fifteenthIdx[slot], null];
            TestAssert.Equal(true, (bool)fifteenth!.Invoke(null, fifteenthArgs)!);
            TestAssert.Equal((byte)fifteenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifteenthArgs[1])!);
            TestAssert.Equal((byte)fifteenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifteenthArgs[1])!);
            TestAssert.Equal((byte)fifteenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifteenthArgs[1])!);
        }
        object?[] fifteenthPast = [18, null];
        TestAssert.Equal(false, (bool)fifteenth.Invoke(null, fifteenthPast)!);
        // Twenty-first-regime FourteenthSingle (slots 76), dash126/dash127 agree bit-identical.
        int[] fourteenthSingleIdx = [75];
        int[] fourteenthSingleR = [144];
        int[] fourteenthSingleG = [123];
        int[] fourteenthSingleB = [172];
        for (int slot = 0; slot < fourteenthSingleIdx.Length; slot++)
        {
            object?[] fourteenthSingleArgs = [fourteenthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)fourteenthSingle!.Invoke(null, fourteenthSingleArgs)!);
            TestAssert.Equal((byte)fourteenthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourteenthSingleArgs[1])!);
            TestAssert.Equal((byte)fourteenthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourteenthSingleArgs[1])!);
            TestAssert.Equal((byte)fourteenthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourteenthSingleArgs[1])!);
        }
        object?[] fourteenthSinglePast = [74, null];
        TestAssert.Equal(false, (bool)fourteenthSingle.Invoke(null, fourteenthSinglePast)!);
        // Twenty-first-regime LateSingles (slots 73,74,75,77,78), dash126/dash127 agree bit-identical.
        int[] lateSinglesIdx = [72, 73, 74, 76, 77];
        int[] lateSinglesR = [112, 199, 167, 109, 248];
        int[] lateSinglesG = [147, 112, 194, 182, 163];
        int[] lateSinglesB = [196, 111, 118, 204, 107];
        for (int slot = 0; slot < lateSinglesIdx.Length; slot++)
        {
            object?[] lateSinglesArgs = [lateSinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)lateSingles!.Invoke(null, lateSinglesArgs)!);
            TestAssert.Equal((byte)lateSinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(lateSinglesArgs[1])!);
            TestAssert.Equal((byte)lateSinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(lateSinglesArgs[1])!);
            TestAssert.Equal((byte)lateSinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(lateSinglesArgs[1])!);
        }
        object?[] lateSinglesPast = [75, null];
        TestAssert.Equal(false, (bool)lateSingles.Invoke(null, lateSinglesPast)!);
        // Twenty-first-regime Nineteenth (slots 43,44,45,47,48,50,52,54,60), dash126/dash127 agree bit-identical.
        int[] nineteenthIdx = [42, 43, 44, 46, 47, 49, 51, 53, 59];
        int[] nineteenthR = [69, 170, 137, 65, 219, 177, 117, 228, 236];
        int[] nineteenthG = [114, 70, 165, 152, 132, 73, 91, 138, 143];
        int[] nineteenthB = [167, 67, 78, 175, 61, 70, 149, 64, 66];
        for (int slot = 0; slot < nineteenthIdx.Length; slot++)
        {
            object?[] nineteenthArgs = [nineteenthIdx[slot], null];
            TestAssert.Equal(true, (bool)nineteenth!.Invoke(null, nineteenthArgs)!);
            TestAssert.Equal((byte)nineteenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(nineteenthArgs[1])!);
            TestAssert.Equal((byte)nineteenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(nineteenthArgs[1])!);
            TestAssert.Equal((byte)nineteenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(nineteenthArgs[1])!);
        }
        object?[] nineteenthPast = [41, null];
        TestAssert.Equal(false, (bool)nineteenth.Invoke(null, nineteenthPast)!);
        // Twenty-first-regime NinthReplay (slots 121,122,123,124), dash126/dash127 agree bit-identical.
        int[] ninthReplayIdx = [120, 121, 122, 123];
        int[] ninthReplayR = [204, 231, 219, 213];
        int[] ninthReplayG = [213, 204, 229, 207];
        int[] ninthReplayB = [230, 204, 205, 221];
        for (int slot = 0; slot < ninthReplayIdx.Length; slot++)
        {
            object?[] ninthReplayArgs = [ninthReplayIdx[slot], null];
            TestAssert.Equal(true, (bool)ninthReplay!.Invoke(null, ninthReplayArgs)!);
            TestAssert.Equal((byte)ninthReplayR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninthReplayArgs[1])!);
            TestAssert.Equal((byte)ninthReplayG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninthReplayArgs[1])!);
            TestAssert.Equal((byte)ninthReplayB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninthReplayArgs[1])!);
        }
        object?[] ninthReplayPast = [119, null];
        TestAssert.Equal(false, (bool)ninthReplay.Invoke(null, ninthReplayPast)!);
        // Twenty-first-regime Seventeenth (slots 27,29,35,55,56,57,58,59,61,62,63,64,65,66,67,68,69,70,71,72), dash126/dash127 agree bit-identical.
        int[] seventeenthIdx = [26, 28, 34, 54, 55, 56, 57, 58, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71];
        int[] seventeenthR = [122, 57, 61, 74, 181, 146, 121, 70, 77, 189, 152, 126, 74, 243, 92, 194, 159, 134, 88, 247];
        int[] seventeenthG = [147, 135, 142, 122, 75, 177, 94, 162, 127, 78, 184, 98, 169, 147, 135, 92, 190, 109, 175, 155];
        int[] seventeenthB = [69, 156, 164, 178, 72, 84, 153, 187, 186, 75, 87, 159, 194, 69, 191, 90, 100, 166, 200, 84];
        for (int slot = 0; slot < seventeenthIdx.Length; slot++)
        {
            object?[] seventeenthArgs = [seventeenthIdx[slot], null];
            TestAssert.Equal(true, (bool)seventeenth!.Invoke(null, seventeenthArgs)!);
            TestAssert.Equal((byte)seventeenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenthArgs[1])!);
            TestAssert.Equal((byte)seventeenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenthArgs[1])!);
            TestAssert.Equal((byte)seventeenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenthArgs[1])!);
        }
        object?[] seventeenthPast = [33, null];
        TestAssert.Equal(false, (bool)seventeenth.Invoke(null, seventeenthPast)!);
        // Twenty-first-regime Seventh (slots 26,28,103,104,105,106), dash126/dash127 agree bit-identical.
        int[] seventhIdx = [25, 27, 102, 103, 104, 105];
        int[] seventhR = [154, 102, 175, 219, 201, 190];
        int[] seventhG = [62, 79, 190, 175, 216, 180];
        int[] seventhB = [60, 129, 217, 175, 177, 204];
        for (int slot = 0; slot < seventhIdx.Length; slot++)
        {
            object?[] seventhArgs = [seventhIdx[slot], null];
            TestAssert.Equal(true, (bool)seventh!.Invoke(null, seventhArgs)!);
            TestAssert.Equal((byte)seventhR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventhArgs[1])!);
            TestAssert.Equal((byte)seventhG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventhArgs[1])!);
            TestAssert.Equal((byte)seventhB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventhArgs[1])!);
        }
        object?[] seventhPast = [26, null];
        TestAssert.Equal(false, (bool)seventh.Invoke(null, seventhPast)!);
        // Twenty-first-regime Sixteenth (slots 79,80,81,82,83,84), dash126/dash127 agree bit-identical.
        int[] sixteenthIdx = [78, 79, 80, 81, 82, 83];
        int[] sixteenthR = [128, 203, 175, 156, 126, 248];
        int[] sixteenthG = [157, 129, 199, 139, 188, 171];
        int[] sixteenthB = [201, 127, 133, 180, 208, 124];
        for (int slot = 0; slot < sixteenthIdx.Length; slot++)
        {
            object?[] sixteenthArgs = [sixteenthIdx[slot], null];
            TestAssert.Equal(true, (bool)sixteenth!.Invoke(null, sixteenthArgs)!);
            TestAssert.Equal((byte)sixteenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixteenthArgs[1])!);
            TestAssert.Equal((byte)sixteenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixteenthArgs[1])!);
            TestAssert.Equal((byte)sixteenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixteenthArgs[1])!);
        }
        object?[] sixteenthPast = [77, null];
        TestAssert.Equal(false, (bool)sixteenth.Invoke(null, sixteenthPast)!);
        // Twenty-first-regime Thirteenth (slots 13,15,97,98,99,100,101,102,115,116,117,118,119,120), dash126/dash127 agree bit-identical.
        int[] thirteenthIdx = [12, 14, 96, 97, 98, 99, 100, 101, 114, 115, 116, 117, 118, 119];
        int[] thirteenthR = [54, 110, 164, 215, 195, 182, 163, 250, 195, 227, 214, 206, 195, 251];
        int[] thirteenthG = [91, 134, 182, 165, 212, 171, 204, 192, 206, 196, 225, 199, 220, 212];
        int[] thirteenthB = [136, 62, 213, 164, 167, 198, 218, 162, 226, 195, 197, 216, 229, 194];
        for (int slot = 0; slot < thirteenthIdx.Length; slot++)
        {
            object?[] thirteenthArgs = [thirteenthIdx[slot], null];
            TestAssert.Equal(true, (bool)thirteenth!.Invoke(null, thirteenthArgs)!);
            TestAssert.Equal((byte)thirteenthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(thirteenthArgs[1])!);
            TestAssert.Equal((byte)thirteenthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(thirteenthArgs[1])!);
            TestAssert.Equal((byte)thirteenthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(thirteenthArgs[1])!);
        }
        object?[] thirteenthPast = [11, null];
        TestAssert.Equal(false, (bool)thirteenth.Invoke(null, thirteenthPast)!);
        // Twenty-first-regime Twelfth (slots 85,87,88,89,90,107,108), dash126/dash127 agree bit-identical.
        int[] twelfthIdx = [84, 86, 87, 88, 89, 106, 107];
        int[] twelfthR = [140, 181, 164, 139, 249, 177, 250];
        int[] twelfthG = [165, 203, 149, 192, 178, 210, 201];
        int[] twelfthB = [204, 144, 186, 211, 137, 223, 176];
        for (int slot = 0; slot < twelfthIdx.Length; slot++)
        {
            object?[] twelfthArgs = [twelfthIdx[slot], null];
            TestAssert.Equal(true, (bool)twelfth!.Invoke(null, twelfthArgs)!);
            TestAssert.Equal((byte)twelfthR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfthArgs[1])!);
            TestAssert.Equal((byte)twelfthG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfthArgs[1])!);
            TestAssert.Equal((byte)twelfthB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfthArgs[1])!);
        }
        object?[] twelfthPast = [83, null];
        TestAssert.Equal(false, (bool)twelfth.Invoke(null, twelfthPast)!);
        // Twenty-first-regime Twentieth (slots 126), dash126/dash127 agree bit-identical.
        int[] twentiethIdx = [125];
        int[] twentiethR = [252];
        int[] twentiethG = [218];
        int[] twentiethB = [203];
        for (int slot = 0; slot < twentiethIdx.Length; slot++)
        {
            object?[] twentiethArgs = [twentiethIdx[slot], null];
            TestAssert.Equal(true, (bool)twentieth!.Invoke(null, twentiethArgs)!);
            TestAssert.Equal((byte)twentiethR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentiethArgs[1])!);
            TestAssert.Equal((byte)twentiethG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentiethArgs[1])!);
            TestAssert.Equal((byte)twentiethB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentiethArgs[1])!);
        }
        object?[] twentiethPast = [124, null];
        TestAssert.Equal(false, (bool)twentieth.Invoke(null, twentiethPast)!);
        var tail = rendererType.GetMethod(
            "TryResolveTwentyFirstRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected twenty-first-regime tail table to remain inspectable by the Office evidence guard.");
        // Twenty-first-regime Tail (slots 127-130), dash127/dash128, dash128/dash129, dash129/dash130 and dash130/dash131 agree bit-identical.
        int[] tailIdx = [126, 127, 128, 129];
        int[] tailR = [211, 234, 224, 218];
        int[] tailG = [218, 211, 232, 213];
        int[] tailB = [233, 211, 212, 226];
        for (int slot = 0; slot < tailIdx.Length; slot++)
        {
            object?[] tailArgs = [tailIdx[slot], null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] tailPast = [125, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailPast)!);
        object?[] tailFuture = [130, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailFuture)!);
    }
    public static void PptxSyntheticTwentySecondVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwentySecondVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twenty-second-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var nineteenthSingle = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeNineteenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(nineteenthSingle is not null, "Expected twenty-second-regime nineteenthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime NineteenthSingle (slot 42), dash132/dash133 agree bit-identical.
        int[] nineteenthSingleIdx = [41];
        int[] nineteenthSingleR = [211];
        int[] nineteenthSingleG = [127];
        int[] nineteenthSingleB = [58];
        for (int slot = 0; slot < nineteenthSingleIdx.Length; slot++)
        {
            object?[] nineteenthSingleArgs = [nineteenthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)nineteenthSingle!.Invoke(null, nineteenthSingleArgs)!);
            TestAssert.Equal((byte)nineteenthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(nineteenthSingleArgs[1])!);
            TestAssert.Equal((byte)nineteenthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(nineteenthSingleArgs[1])!);
            TestAssert.Equal((byte)nineteenthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(nineteenthSingleArgs[1])!);
        }
        object?[] nineteenthSinglePast = [40, null];
        TestAssert.Equal(false, (bool)nineteenthSingle!.Invoke(null, nineteenthSinglePast)!);
        var seventeenth22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeSeventeenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenth22 is not null, "Expected twenty-second-regime seventeenth table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Seventeenth (slots 54, 60), dash132/dash133 agree bit-identical replays within 1.
        int[] seventeenth22Idx = [53, 59];
        int[] seventeenth22R = [225, 233];
        int[] seventeenth22G = [136, 141];
        int[] seventeenth22B = [63, 66];
        for (int slot = 0; slot < seventeenth22Idx.Length; slot++)
        {
            object?[] seventeenth22Args = [seventeenth22Idx[slot], null];
            TestAssert.Equal(true, (bool)seventeenth22!.Invoke(null, seventeenth22Args)!);
            TestAssert.Equal((byte)seventeenth22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenth22Args[1])!);
            TestAssert.Equal((byte)seventeenth22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenth22Args[1])!);
            TestAssert.Equal((byte)seventeenth22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenth22Args[1])!);
        }
        object?[] seventeenth22Past = [54, null];
        TestAssert.Equal(false, (bool)seventeenth22!.Invoke(null, seventeenth22Past)!);
        var eighteenth22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeEighteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenth22 is not null, "Expected twenty-second-regime eighteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Eighteenth (slots 66-72), dash132/dash133 agree bit-identical replays within 1.
        int[] eighteenth22Idx = [65, 66, 67, 68, 69, 70, 71];
        int[] eighteenth22R = [239, 79, 192, 155, 128, 75, 247];
        int[] eighteenth22G = [145, 129, 80, 187, 100, 172, 150];
        int[] eighteenth22B = [67, 189, 77, 89, 162, 198, 70];
        for (int slot = 0; slot < eighteenth22Idx.Length; slot++)
        {
            object?[] eighteenth22Args = [eighteenth22Idx[slot], null];
            TestAssert.Equal(true, (bool)eighteenth22!.Invoke(null, eighteenth22Args)!);
            TestAssert.Equal((byte)eighteenth22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenth22Args[1])!);
            TestAssert.Equal((byte)eighteenth22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenth22Args[1])!);
            TestAssert.Equal((byte)eighteenth22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenth22Args[1])!);
        }
        object?[] eighteenth22Past = [64, null];
        TestAssert.Equal(false, (bool)eighteenth22!.Invoke(null, eighteenth22Past)!);
        var eleventh22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeEleventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventh22 is not null, "Expected twenty-second-regime eleventh table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Eleventh (slots 73-78), dash132/dash133 agree bit-identical.
        int[] eleventh22Idx = [72, 73, 74, 75, 76, 77];
        int[] eleventh22R = [99, 196, 162, 138, 96, 247];
        int[] eleventh22G = [139, 100, 191, 115, 177, 158];
        int[] eleventh22B = [193, 97, 106, 168, 201, 92];
        for (int slot = 0; slot < eleventh22Idx.Length; slot++)
        {
            object?[] eleventh22Args = [eleventh22Idx[slot], null];
            TestAssert.Equal(true, (bool)eleventh22!.Invoke(null, eleventh22Args)!);
            TestAssert.Equal((byte)eleventh22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventh22Args[1])!);
            TestAssert.Equal((byte)eleventh22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventh22Args[1])!);
            TestAssert.Equal((byte)eleventh22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventh22Args[1])!);
        }
        object?[] eleventh22Past = [60, null];
        TestAssert.Equal(false, (bool)eleventh22!.Invoke(null, eleventh22Past)!);
        var tenth22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeTenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tenth22 is not null, "Expected twenty-second-regime tenth table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Tenth (slots 79-84, 119-120, 127-130), dash132/dash133 agree bit-identical.
        int[] tenth22Idx = [78, 79, 80, 81, 82, 83, 118, 119, 126, 127, 128, 129];
        int[] tenth22R = [118, 200, 170, 149, 115, 248, 187, 251, 205, 231, 220, 214];
        int[] tenth22G = [150, 118, 196, 130, 184, 166, 215, 207, 214, 205, 230, 208];
        int[] tenth22B = [198, 116, 123, 176, 205, 113, 227, 186, 230, 205, 207, 222];
        for (int slot = 0; slot < tenth22Idx.Length; slot++)
        {
            object?[] tenth22Args = [tenth22Idx[slot], null];
            TestAssert.Equal(true, (bool)tenth22!.Invoke(null, tenth22Args)!);
            TestAssert.Equal((byte)tenth22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tenth22Args[1])!);
            TestAssert.Equal((byte)tenth22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tenth22Args[1])!);
            TestAssert.Equal((byte)tenth22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tenth22Args[1])!);
        }
        object?[] tenth22Past = [61, null];
        TestAssert.Equal(false, (bool)tenth22!.Invoke(null, tenth22Past)!);
        var sixth22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeSixthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixth22 is not null, "Expected twenty-second-regime sixth table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Sixth (slots 85-90, 107-108), dash132/dash133 agree bit-identical.
        int[] sixth22Idx = [84, 85, 86, 87, 88, 89, 106, 107];
        int[] sixth22R = [133, 205, 177, 160, 131, 248, 169, 250];
        int[] sixth22G = [160, 134, 201, 143, 190, 174, 206, 195];
        int[] sixth22B = [202, 132, 138, 182, 209, 129, 220, 168];
        for (int slot = 0; slot < sixth22Idx.Length; slot++)
        {
            object?[] sixth22Args = [sixth22Idx[slot], null];
            TestAssert.Equal(true, (bool)sixth22!.Invoke(null, sixth22Args)!);
            TestAssert.Equal((byte)sixth22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixth22Args[1])!);
            TestAssert.Equal((byte)sixth22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixth22Args[1])!);
            TestAssert.Equal((byte)sixth22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixth22Args[1])!);
        }
        object?[] sixth22Past = [62, null];
        TestAssert.Equal(false, (bool)sixth22!.Invoke(null, sixth22Past)!);
        var overflow22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflow22 is not null, "Expected twenty-second-regime overflow table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Overflow (slots 91-94, 103-106, 115-118, 121-124), dash132/dash133 agree bit-identical.
        int[] overflow22Idx = [90, 91, 92, 93, 102, 103, 104, 105, 114, 115, 116, 117, 120, 121, 122, 123];
        int[] overflow22R = [147, 209, 185, 169, 170, 217, 198, 186, 188, 224, 209, 200, 197, 228, 215, 207];
        int[] overflow22G = [169, 147, 205, 155, 186, 170, 214, 176, 200, 188, 222, 192, 207, 197, 226, 201];
        int[] overflow22B = [207, 146, 150, 189, 215, 169, 172, 201, 223, 188, 190, 212, 226, 197, 199, 217];
        for (int slot = 0; slot < overflow22Idx.Length; slot++)
        {
            object?[] overflow22Args = [overflow22Idx[slot], null];
            TestAssert.Equal(true, (bool)overflow22!.Invoke(null, overflow22Args)!);
            TestAssert.Equal((byte)overflow22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflow22Args[1])!);
            TestAssert.Equal((byte)overflow22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflow22Args[1])!);
            TestAssert.Equal((byte)overflow22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflow22Args[1])!);
        }
        object?[] overflow22Past = [63, null];
        TestAssert.Equal(false, (bool)overflow22!.Invoke(null, overflow22Past)!);
        var fourth22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeFourthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourth22 is not null, "Expected twenty-second-regime fourth table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Fourth (slots 95-96), dash132/dash133 agree bit-identical.
        int[] fourth22Idx = [94, 95];
        int[] fourth22R = [145, 249];
        int[] fourth22G = [195, 181];
        int[] fourth22B = [213, 144];
        for (int slot = 0; slot < fourth22Idx.Length; slot++)
        {
            object?[] fourth22Args = [fourth22Idx[slot], null];
            TestAssert.Equal(true, (bool)fourth22!.Invoke(null, fourth22Args)!);
            TestAssert.Equal((byte)fourth22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourth22Args[1])!);
            TestAssert.Equal((byte)fourth22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourth22Args[1])!);
            TestAssert.Equal((byte)fourth22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourth22Args[1])!);
        }
        object?[] fourth22Past = [58, null];
        TestAssert.Equal(false, (bool)fourth22!.Invoke(null, fourth22Past)!);
        var fifth22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeFifthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifth22 is not null, "Expected twenty-second-regime fifth table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Fifth (slots 97-102), dash132/dash133 agree bit-identical replays within 2.
        int[] fifth22Idx = [96, 97, 98, 99, 100, 101];
        int[] fifth22R = [161, 214, 192, 179, 160, 249];
        int[] fifth22G = [180, 161, 210, 168, 202, 190];
        int[] fifth22B = [212, 160, 164, 196, 217, 158];
        for (int slot = 0; slot < fifth22Idx.Length; slot++)
        {
            object?[] fifth22Args = [fifth22Idx[slot], null];
            TestAssert.Equal(true, (bool)fifth22!.Invoke(null, fifth22Args)!);
            TestAssert.Equal((byte)fifth22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifth22Args[1])!);
            TestAssert.Equal((byte)fifth22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifth22Args[1])!);
            TestAssert.Equal((byte)fifth22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifth22Args[1])!);
        }
        object?[] fifth22Past = [57, null];
        TestAssert.Equal(false, (bool)fifth22!.Invoke(null, fifth22Past)!);
        var twelfth22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeTwelfthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfth22 is not null, "Expected twenty-second-regime twelfth table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Twelfth (slots 109-114), dash132/dash133 agree bit-identical.
        int[] twelfth22Idx = [108, 109, 110, 111, 112, 113];
        int[] twelfth22R = [178, 220, 203, 192, 177, 250];
        int[] twelfth22G = [193, 179, 218, 184, 210, 201];
        int[] twelfth22B = [219, 178, 181, 206, 223, 176];
        for (int slot = 0; slot < twelfth22Idx.Length; slot++)
        {
            object?[] twelfth22Args = [twelfth22Idx[slot], null];
            TestAssert.Equal(true, (bool)twelfth22!.Invoke(null, twelfth22Args)!);
            TestAssert.Equal((byte)twelfth22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfth22Args[1])!);
            TestAssert.Equal((byte)twelfth22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfth22Args[1])!);
            TestAssert.Equal((byte)twelfth22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfth22Args[1])!);
        }
        object?[] twelfth22Past = [56, null];
        TestAssert.Equal(false, (bool)twelfth22!.Invoke(null, twelfth22Past)!);
        var fourteenth22 = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeFourteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourteenth22 is not null, "Expected twenty-second-regime fourteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Fourteenth (slots 125-126), dash132/dash133 agree bit-identical.
        int[] fourteenth22Idx = [124, 125];
        int[] fourteenth22R = [196, 251];
        int[] fourteenth22G = [220, 213];
        int[] fourteenth22B = [230, 196];
        for (int slot = 0; slot < fourteenth22Idx.Length; slot++)
        {
            object?[] fourteenth22Args = [fourteenth22Idx[slot], null];
            TestAssert.Equal(true, (bool)fourteenth22!.Invoke(null, fourteenth22Args)!);
            TestAssert.Equal((byte)fourteenth22R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourteenth22Args[1])!);
            TestAssert.Equal((byte)fourteenth22G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourteenth22Args[1])!);
            TestAssert.Equal((byte)fourteenth22B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourteenth22Args[1])!);
        }
        object?[] fourteenth22Past = [55, null];
        TestAssert.Equal(false, (bool)fourteenth22!.Invoke(null, fourteenth22Past)!);
        var lateSingle = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeLateSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(lateSingle is not null, "Expected twenty-second-regime latesingle table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime LateSingles (slot 131 dash132/dash133 agree, slot 132 dash132-136 agree quintuple, slot 136 dash136/dash137 agree).
        int[] lateSingleIdx = [130, 131, 135];
        int[] lateSingleR = [205, 252, 220];
        int[] lateSingleG = [225, 219, 216];
        int[] lateSingleB = [233, 204, 227];
        for (int slot = 0; slot < lateSingleIdx.Length; slot++)
        {
            object?[] lateSingleArgs = [lateSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)lateSingle!.Invoke(null, lateSingleArgs)!);
            TestAssert.Equal((byte)lateSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(lateSingleArgs[1])!);
            TestAssert.Equal((byte)lateSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(lateSingleArgs[1])!);
            TestAssert.Equal((byte)lateSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(lateSingleArgs[1])!);
        }
        object?[] lateSinglePast = [129, null];
        TestAssert.Equal(false, (bool)lateSingle!.Invoke(null, lateSinglePast)!);
        var tail = rendererType.GetMethod(
            "TryResolveTwentySecondRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected twenty-second-regime tail table to remain inspectable by the Office evidence guard.");
        // Twenty-second-regime Tail (slots 134-136), dash133/dash134, dash134/dash135 and dash135/dash136 agree bit-identical.
        int[] tailIdx = [132, 133, 134];
        int[] tailR = [211, 234, 224];
        int[] tailG = [218, 211, 232];
        int[] tailB = [233, 211, 212];
        for (int slot = 0; slot < tailIdx.Length; slot++)
        {
            object?[] tailArgs = [tailIdx[slot], null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] tailPast = [131, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailPast)!);
        object?[] tailFuture = [135, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailFuture)!);
        // One-hundred-thirty-one points and fewer keep earlier regimes; one-hundred-thirty-two-plus take the twenty-second rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [131])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [132])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [133])!);
    }
    public static void PptxSyntheticTwentyThirdVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwentyThirdVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twenty-third-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var seventhSingle = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeSeventhSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventhSingle is not null, "Expected twenty-third-regime seventhsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime SeventhSingle (slot 31), dash138/dash139 agree bit-identical replays within 2.
        int[] seventhSingleIdx = [30];
        int[] seventhSingleR = [62];
        int[] seventhSingleG = [102];
        int[] seventhSingleB = [151];
        for (int slot = 0; slot < seventhSingleIdx.Length; slot++)
        {
            object?[] seventhSingleArgs = [seventhSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)seventhSingle!.Invoke(null, seventhSingleArgs)!);
            TestAssert.Equal((byte)seventhSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventhSingleArgs[1])!);
            TestAssert.Equal((byte)seventhSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventhSingleArgs[1])!);
            TestAssert.Equal((byte)seventhSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventhSingleArgs[1])!);
        }
        object?[] seventhSinglePast = [31, null];
        TestAssert.Equal(false, (bool)seventhSingle!.Invoke(null, seventhSinglePast)!);
        var nineteenth23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeNineteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(nineteenth23 is not null, "Expected twenty-third-regime nineteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Nineteenth (slots 36, 66), dash138/dash139 agree bit-identical replays within 1.
        int[] nineteenth23Idx = [35, 65];
        int[] nineteenth23R = [201, 236];
        int[] nineteenth23G = [121, 143];
        int[] nineteenth23B = [55, 66];
        for (int slot = 0; slot < nineteenth23Idx.Length; slot++)
        {
            object?[] nineteenth23Args = [nineteenth23Idx[slot], null];
            TestAssert.Equal(true, (bool)nineteenth23!.Invoke(null, nineteenth23Args)!);
            TestAssert.Equal((byte)nineteenth23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(nineteenth23Args[1])!);
            TestAssert.Equal((byte)nineteenth23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(nineteenth23Args[1])!);
            TestAssert.Equal((byte)nineteenth23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(nineteenth23Args[1])!);
        }
        object?[] nineteenth23Past = [34, null];
        TestAssert.Equal(false, (bool)nineteenth23!.Invoke(null, nineteenth23Past)!);
        var seventeenth23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeSeventeenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenth23 is not null, "Expected twenty-third-regime seventeenth table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Seventeenth (slots 48, 63-64, 67-69, 71-72), dash138/dash139 agree bit-identical replays within 2.
        int[] seventeenth23Idx = [47, 62, 63, 66, 67, 68, 70, 71];
        int[] seventeenth23R = [216, 146, 121, 77, 189, 152, 74, 243];
        int[] seventeenth23G = [131, 177, 94, 127, 78, 184, 169, 147];
        int[] seventeenth23B = [60, 84, 153, 186, 75, 87, 194, 69];
        for (int slot = 0; slot < seventeenth23Idx.Length; slot++)
        {
            object?[] seventeenth23Args = [seventeenth23Idx[slot], null];
            TestAssert.Equal(true, (bool)seventeenth23!.Invoke(null, seventeenth23Args)!);
            TestAssert.Equal((byte)seventeenth23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenth23Args[1])!);
            TestAssert.Equal((byte)seventeenth23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenth23Args[1])!);
            TestAssert.Equal((byte)seventeenth23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenth23Args[1])!);
        }
        object?[] seventeenth23Past = [46, null];
        TestAssert.Equal(false, (bool)seventeenth23!.Invoke(null, seventeenth23Past)!);
        var eighteenth23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeEighteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenth23 is not null, "Expected twenty-third-regime eighteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Eighteenth (slots 61-62, 65), dash138/dash139 agree bit-identical replays within 2.
        int[] eighteenth23Idx = [60, 61, 64];
        int[] eighteenth23R = [76, 186, 72];
        int[] eighteenth23G = [125, 77, 166];
        int[] eighteenth23B = [183, 74, 192];
        for (int slot = 0; slot < eighteenth23Idx.Length; slot++)
        {
            object?[] eighteenth23Args = [eighteenth23Idx[slot], null];
            TestAssert.Equal(true, (bool)eighteenth23!.Invoke(null, eighteenth23Args)!);
            TestAssert.Equal((byte)eighteenth23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenth23Args[1])!);
            TestAssert.Equal((byte)eighteenth23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenth23Args[1])!);
            TestAssert.Equal((byte)eighteenth23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenth23Args[1])!);
        }
        object?[] eighteenth23Past = [59, null];
        TestAssert.Equal(false, (bool)eighteenth23!.Invoke(null, eighteenth23Past)!);
        var fourteenth23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeFourteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourteenth23 is not null, "Expected twenty-third-regime fourteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Fourteenth (slots 79-84), dash138/dash139 agree bit-identical.
        int[] fourteenth23Idx = [78, 79, 80, 81, 82, 83];
        int[] fourteenth23R = [109, 198, 166, 144, 106, 248];
        int[] fourteenth23G = [145, 109, 194, 123, 181, 162];
        int[] fourteenth23B = [195, 107, 115, 172, 203, 103];
        for (int slot = 0; slot < fourteenth23Idx.Length; slot++)
        {
            object?[] fourteenth23Args = [fourteenth23Idx[slot], null];
            TestAssert.Equal(true, (bool)fourteenth23!.Invoke(null, fourteenth23Args)!);
            TestAssert.Equal((byte)fourteenth23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourteenth23Args[1])!);
            TestAssert.Equal((byte)fourteenth23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourteenth23Args[1])!);
            TestAssert.Equal((byte)fourteenth23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourteenth23Args[1])!);
        }
        object?[] fourteenth23Past = [77, null];
        TestAssert.Equal(false, (bool)fourteenth23!.Invoke(null, fourteenth23Past)!);
        var third23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeThirdFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(third23 is not null, "Expected twenty-third-regime third table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Third (slots 85-90), dash138/dash139 agree bit-identical.
        int[] third23Idx = [84, 85, 86, 87, 88, 89];
        int[] third23R = [126, 202, 174, 155, 124, 248];
        int[] third23G = [155, 126, 198, 137, 187, 170];
        int[] third23B = [200, 125, 131, 179, 207, 121];
        for (int slot = 0; slot < third23Idx.Length; slot++)
        {
            object?[] third23Args = [third23Idx[slot], null];
            TestAssert.Equal(true, (bool)third23!.Invoke(null, third23Args)!);
            TestAssert.Equal((byte)third23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(third23Args[1])!);
            TestAssert.Equal((byte)third23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(third23Args[1])!);
            TestAssert.Equal((byte)third23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(third23Args[1])!);
        }
        object?[] third23Past = [83, null];
        TestAssert.Equal(false, (bool)third23!.Invoke(null, third23Past)!);
        var ninth23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeNinthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninth23 is not null, "Expected twenty-third-regime ninth table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Ninth (slots 91-96), dash138/dash139 agree bit-identical.
        int[] ninth23Idx = [90, 91, 92, 93, 94, 95];
        int[] ninth23R = [138, 206, 180, 163, 136, 249];
        int[] ninth23G = [163, 138, 202, 147, 192, 177];
        int[] ninth23B = [204, 137, 142, 185, 210, 134];
        for (int slot = 0; slot < ninth23Idx.Length; slot++)
        {
            object?[] ninth23Args = [ninth23Idx[slot], null];
            TestAssert.Equal(true, (bool)ninth23!.Invoke(null, ninth23Args)!);
            TestAssert.Equal((byte)ninth23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninth23Args[1])!);
            TestAssert.Equal((byte)ninth23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninth23Args[1])!);
            TestAssert.Equal((byte)ninth23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninth23Args[1])!);
        }
        object?[] ninth23Past = [89, null];
        TestAssert.Equal(false, (bool)ninth23!.Invoke(null, ninth23Past)!);
        var eleventh23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeEleventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventh23 is not null, "Expected twenty-third-regime eleventh table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Eleventh (slots 97-102, 109, 113, 124-126), dash138/dash139 agree bit-identical replays within 2.
        int[] eleventh23Idx = [96, 97, 98, 99, 100, 101, 108, 112, 123, 124, 125];
        int[] eleventh23R = [153, 211, 188, 173, 152, 249, 173, 172, 202, 190, 251];
        int[] eleventh23G = [174, 153, 208, 161, 198, 185, 189, 208, 195, 217, 209];
        int[] eleventh23B = [209, 152, 156, 192, 215, 150, 217, 221, 213, 228, 189];
        for (int slot = 0; slot < eleventh23Idx.Length; slot++)
        {
            object?[] eleventh23Args = [eleventh23Idx[slot], null];
            TestAssert.Equal(true, (bool)eleventh23!.Invoke(null, eleventh23Args)!);
            TestAssert.Equal((byte)eleventh23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventh23Args[1])!);
            TestAssert.Equal((byte)eleventh23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventh23Args[1])!);
            TestAssert.Equal((byte)eleventh23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventh23Args[1])!);
        }
        object?[] eleventh23Past = [95, null];
        TestAssert.Equal(false, (bool)eleventh23!.Invoke(null, eleventh23Past)!);
        var fifth23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeFifthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifth23 is not null, "Expected twenty-third-regime fifth table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Fifth (slots 103-108), dash138/dash139 agree bit-identical.
        int[] fifth23Idx = [102, 103, 104, 105, 106, 107];
        int[] fifth23R = [161, 214, 192, 179, 160, 249];
        int[] fifth23G = [180, 161, 210, 168, 202, 190];
        int[] fifth23B = [212, 160, 164, 196, 217, 158];
        for (int slot = 0; slot < fifth23Idx.Length; slot++)
        {
            object?[] fifth23Args = [fifth23Idx[slot], null];
            TestAssert.Equal(true, (bool)fifth23!.Invoke(null, fifth23Args)!);
            TestAssert.Equal((byte)fifth23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifth23Args[1])!);
            TestAssert.Equal((byte)fifth23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifth23Args[1])!);
            TestAssert.Equal((byte)fifth23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifth23Args[1])!);
        }
        object?[] fifth23Past = [101, null];
        TestAssert.Equal(false, (bool)fifth23!.Invoke(null, fifth23Past)!);
        var overflow23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflow23 is not null, "Expected twenty-third-regime overflow table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Overflow (slots 110-112, 115-118, 121-123), dash138/dash139 agree bit-identical replays within 2.
        int[] overflow23Idx = [109, 110, 111, 114, 115, 116, 117, 120, 121, 122];
        int[] overflow23R = [217, 198, 186, 182, 221, 205, 195, 188, 224, 209];
        int[] overflow23G = [170, 214, 176, 195, 182, 219, 186, 200, 188, 222];
        int[] overflow23B = [169, 172, 201, 220, 181, 184, 208, 223, 188, 190];
        for (int slot = 0; slot < overflow23Idx.Length; slot++)
        {
            object?[] overflow23Args = [overflow23Idx[slot], null];
            TestAssert.Equal(true, (bool)overflow23!.Invoke(null, overflow23Args)!);
            TestAssert.Equal((byte)overflow23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflow23Args[1])!);
            TestAssert.Equal((byte)overflow23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflow23Args[1])!);
            TestAssert.Equal((byte)overflow23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflow23Args[1])!);
        }
        object?[] overflow23Past = [108, null];
        TestAssert.Equal(false, (bool)overflow23!.Invoke(null, overflow23Past)!);
        var eighth23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeEighthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighth23 is not null, "Expected twenty-third-regime eighth table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Eighth (slots 119-120), dash138/dash139 agree bit-identical.
        int[] eighth23Idx = [118, 119];
        int[] eighth23R = [181, 250];
        int[] eighth23G = [212, 203];
        int[] eighth23B = [224, 180];
        for (int slot = 0; slot < eighth23Idx.Length; slot++)
        {
            object?[] eighth23Args = [eighth23Idx[slot], null];
            TestAssert.Equal(true, (bool)eighth23!.Invoke(null, eighth23Args)!);
            TestAssert.Equal((byte)eighth23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighth23Args[1])!);
            TestAssert.Equal((byte)eighth23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighth23Args[1])!);
            TestAssert.Equal((byte)eighth23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighth23Args[1])!);
        }
        object?[] eighth23Past = [117, null];
        TestAssert.Equal(false, (bool)eighth23!.Invoke(null, eighth23Past)!);
        var sixthSingle = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeSixthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixthSingle is not null, "Expected twenty-third-regime sixthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime SixthSingle (slot 114), dash138/dash139 agree bit-identical replays within 2.
        int[] sixthSingleIdx = [113];
        int[] sixthSingleR = [250];
        int[] sixthSingleG = [195];
        int[] sixthSingleB = [168];
        for (int slot = 0; slot < sixthSingleIdx.Length; slot++)
        {
            object?[] sixthSingleArgs = [sixthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)sixthSingle!.Invoke(null, sixthSingleArgs)!);
            TestAssert.Equal((byte)sixthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixthSingleArgs[1])!);
            TestAssert.Equal((byte)sixthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixthSingleArgs[1])!);
            TestAssert.Equal((byte)sixthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixthSingleArgs[1])!);
        }
        object?[] sixthSinglePast = [112, null];
        TestAssert.Equal(false, (bool)sixthSingle!.Invoke(null, sixthSinglePast)!);
        var fifteenth23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeFifteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifteenth23 is not null, "Expected twenty-third-regime fifteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Fifteenth (slots 127-132), dash138/dash139 agree bit-identical.
        int[] fifteenth23Idx = [126, 127, 128, 129, 130, 131];
        int[] fifteenth23R = [198, 228, 216, 208, 198, 251];
        int[] fifteenth23G = [208, 198, 226, 202, 221, 214];
        int[] fifteenth23B = [227, 198, 200, 218, 231, 197];
        for (int slot = 0; slot < fifteenth23Idx.Length; slot++)
        {
            object?[] fifteenth23Args = [fifteenth23Idx[slot], null];
            TestAssert.Equal(true, (bool)fifteenth23!.Invoke(null, fifteenth23Args)!);
            TestAssert.Equal((byte)fifteenth23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifteenth23Args[1])!);
            TestAssert.Equal((byte)fifteenth23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifteenth23Args[1])!);
            TestAssert.Equal((byte)fifteenth23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifteenth23Args[1])!);
        }
        object?[] fifteenth23Past = [125, null];
        TestAssert.Equal(false, (bool)fifteenth23!.Invoke(null, fifteenth23Past)!);
        var tenth23 = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeTenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tenth23 is not null, "Expected twenty-third-regime tenth table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Tenth (slots 133-136), dash138/dash139 agree bit-identical.
        int[] tenth23Idx = [132, 133, 134, 135];
        int[] tenth23R = [205, 231, 220, 214];
        int[] tenth23G = [214, 205, 230, 208];
        int[] tenth23B = [230, 205, 207, 222];
        for (int slot = 0; slot < tenth23Idx.Length; slot++)
        {
            object?[] tenth23Args = [tenth23Idx[slot], null];
            TestAssert.Equal(true, (bool)tenth23!.Invoke(null, tenth23Args)!);
            TestAssert.Equal((byte)tenth23R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tenth23Args[1])!);
            TestAssert.Equal((byte)tenth23G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tenth23Args[1])!);
            TestAssert.Equal((byte)tenth23B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tenth23Args[1])!);
        }
        object?[] tenth23Past = [131, null];
        TestAssert.Equal(false, (bool)tenth23!.Invoke(null, tenth23Past)!);
        var twentySecondSingle = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeTwentySecondSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentySecondSingle is not null, "Expected twenty-third-regime twentysecondsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime TwentySecondSingle (slot 137), dash138/dash139 agree bit-identical.
        int[] twentySecondSingleIdx = [136];
        int[] twentySecondSingleR = [205];
        int[] twentySecondSingleG = [225];
        int[] twentySecondSingleB = [233];
        for (int slot = 0; slot < twentySecondSingleIdx.Length; slot++)
        {
            object?[] twentySecondSingleArgs = [twentySecondSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)twentySecondSingle!.Invoke(null, twentySecondSingleArgs)!);
            TestAssert.Equal((byte)twentySecondSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentySecondSingleArgs[1])!);
            TestAssert.Equal((byte)twentySecondSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentySecondSingleArgs[1])!);
            TestAssert.Equal((byte)twentySecondSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentySecondSingleArgs[1])!);
        }
        object?[] twentySecondSinglePast = [135, null];
        TestAssert.Equal(false, (bool)twentySecondSingle!.Invoke(null, twentySecondSinglePast)!);
        var midSingles = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeMidSinglesFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(midSingles is not null, "Expected twenty-third-regime midsingles table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime MidSingles (slots 73-78), dash138/dash139 agree bit-identical.
        int[] midSinglesIdx = [72, 73, 74, 75, 76, 77];
        int[] midSinglesR = [88, 194, 158, 132, 84, 247];
        int[] midSinglesG = [133, 88, 189, 106, 174, 153];
        int[] midSinglesB = [191, 86, 96, 165, 199, 80];
        for (int slot = 0; slot < midSinglesIdx.Length; slot++)
        {
            object?[] midSinglesArgs = [midSinglesIdx[slot], null];
            TestAssert.Equal(true, (bool)midSingles!.Invoke(null, midSinglesArgs)!);
            TestAssert.Equal((byte)midSinglesR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(midSinglesArgs[1])!);
            TestAssert.Equal((byte)midSinglesG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(midSinglesArgs[1])!);
            TestAssert.Equal((byte)midSinglesB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(midSinglesArgs[1])!);
        }
        object?[] midSinglesPast = [71, null];
        TestAssert.Equal(false, (bool)midSingles!.Invoke(null, midSinglesPast)!);
        var tail = rendererType.GetMethod(
            "TryResolveTwentyThirdRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected twenty-third-regime tail table to remain inspectable by the Office evidence guard.");
        // Twenty-third-regime Tail (slots 140-143), dash139/dash140, dash140/dash141, dash141/dash142 and dash142/dash143 agree bit-identical.
        int[] tailIdx = [138, 139, 140, 141];
        int[] tailR = [211, 234, 224, 220];
        int[] tailG = [218, 211, 232, 216];
        int[] tailB = [233, 211, 212, 227];
        for (int slot = 0; slot < tailIdx.Length; slot++)
        {
            object?[] tailArgs = [tailIdx[slot], null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] tailPast = [137, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailPast)!);
        object?[] tailFuture = [142, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailFuture)!);
        // One-hundred-thirty-seven points and fewer keep earlier regimes; one-hundred-thirty-eight-plus take the twenty-third rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [137])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [138])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [139])!);
    }
    public static void PptxSyntheticTwentyFourthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwentyFourthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twenty-fourth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var fifteenth24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeFifteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifteenth24 is not null, "Expected twenty-fourth-regime fifteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Fifteenth (slots 25, 29, 133-138), dash144/dash145 agree bit-identical replays within 2.
        int[] fifteenth24Idx = [24, 28, 132, 133, 134, 135, 136, 137];
        int[] fifteenth24R = [58, 55, 198, 228, 216, 208, 198, 251];
        int[] fifteenth24G = [97, 130, 208, 198, 226, 202, 221, 214];
        int[] fifteenth24B = [143, 150, 227, 198, 200, 218, 231, 197];
        for (int slot = 0; slot < fifteenth24Idx.Length; slot++)
        {
            object?[] fifteenth24Args = [fifteenth24Idx[slot], null];
            TestAssert.Equal(true, (bool)fifteenth24!.Invoke(null, fifteenth24Args)!);
            TestAssert.Equal((byte)fifteenth24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifteenth24Args[1])!);
            TestAssert.Equal((byte)fifteenth24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifteenth24Args[1])!);
            TestAssert.Equal((byte)fifteenth24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifteenth24Args[1])!);
        }
        object?[] fifteenth24Past = [25, null];
        TestAssert.Equal(false, (bool)fifteenth24!.Invoke(null, fifteenth24Past)!);
        var nineteenth24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeNineteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(nineteenth24 is not null, "Expected twenty-fourth-regime nineteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Nineteenth (slots 30, 54, 60), dash144/dash145 agree bit-identical replays within 2.
        int[] nineteenth24Idx = [29, 53, 59];
        int[] nineteenth24R = [191, 219, 228];
        int[] nineteenth24G = [115, 132, 138];
        int[] nineteenth24B = [52, 61, 64];
        for (int slot = 0; slot < nineteenth24Idx.Length; slot++)
        {
            object?[] nineteenth24Args = [nineteenth24Idx[slot], null];
            TestAssert.Equal(true, (bool)nineteenth24!.Invoke(null, nineteenth24Args)!);
            TestAssert.Equal((byte)nineteenth24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(nineteenth24Args[1])!);
            TestAssert.Equal((byte)nineteenth24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(nineteenth24Args[1])!);
            TestAssert.Equal((byte)nineteenth24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(nineteenth24Args[1])!);
        }
        object?[] nineteenth24Past = [28, null];
        TestAssert.Equal(false, (bool)nineteenth24!.Invoke(null, nineteenth24Past)!);
        var seventeenth24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeSeventeenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenth24 is not null, "Expected twenty-fourth-regime seventeenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Seventeenth (slots 41-42), dash144/dash145 agree bit-identical replays within 1.
        int[] seventeenth24Idx = [40, 41];
        int[] seventeenth24R = [61, 206];
        int[] seventeenth24G = [142, 124];
        int[] seventeenth24B = [164, 57];
        for (int slot = 0; slot < seventeenth24Idx.Length; slot++)
        {
            object?[] seventeenth24Args = [seventeenth24Idx[slot], null];
            TestAssert.Equal(true, (bool)seventeenth24!.Invoke(null, seventeenth24Args)!);
            TestAssert.Equal((byte)seventeenth24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenth24Args[1])!);
            TestAssert.Equal((byte)seventeenth24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenth24Args[1])!);
            TestAssert.Equal((byte)seventeenth24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenth24Args[1])!);
        }
        object?[] seventeenth24Past = [39, null];
        TestAssert.Equal(false, (bool)seventeenth24!.Invoke(null, seventeenth24Past)!);
        var twentyFirst24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeTwentyFirstFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentyFirst24 is not null, "Expected twenty-fourth-regime twentyfirst table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime TwentyFirst (slots 55, 57, 59), dash144/dash145 agree bit-identical replays within 1.
        int[] twentyFirst24Idx = [54, 56, 58];
        int[] twentyFirst24R = [72, 143, 69];
        int[] twentyFirst24G = [119, 173, 159];
        int[] twentyFirst24B = [175, 82, 183];
        for (int slot = 0; slot < twentyFirst24Idx.Length; slot++)
        {
            object?[] twentyFirst24Args = [twentyFirst24Idx[slot], null];
            TestAssert.Equal(true, (bool)twentyFirst24!.Invoke(null, twentyFirst24Args)!);
            TestAssert.Equal((byte)twentyFirst24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentyFirst24Args[1])!);
            TestAssert.Equal((byte)twentyFirst24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentyFirst24Args[1])!);
            TestAssert.Equal((byte)twentyFirst24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentyFirst24Args[1])!);
        }
        object?[] twentyFirst24Past = [55, null];
        TestAssert.Equal(false, (bool)twentyFirst24!.Invoke(null, twentyFirst24Past)!);
        var eighteenth24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeEighteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenth24 is not null, "Expected twenty-fourth-regime eighteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Eighteenth (slots 71-78, 98), dash144/dash145 agree bit-identical replays within 1.
        int[] eighteenth24Idx = [70, 71, 72, 73, 74, 75, 76, 77, 97];
        int[] eighteenth24R = [72, 239, 79, 192, 155, 128, 75, 247, 207];
        int[] eighteenth24G = [166, 145, 129, 80, 187, 100, 172, 150, 143];
        int[] eighteenth24B = [192, 67, 189, 77, 89, 162, 198, 70, 142];
        for (int slot = 0; slot < eighteenth24Idx.Length; slot++)
        {
            object?[] eighteenth24Args = [eighteenth24Idx[slot], null];
            TestAssert.Equal(true, (bool)eighteenth24!.Invoke(null, eighteenth24Args)!);
            TestAssert.Equal((byte)eighteenth24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenth24Args[1])!);
            TestAssert.Equal((byte)eighteenth24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenth24Args[1])!);
            TestAssert.Equal((byte)eighteenth24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenth24Args[1])!);
        }
        object?[] eighteenth24Past = [69, null];
        TestAssert.Equal(false, (bool)eighteenth24!.Invoke(null, eighteenth24Past)!);
        var eleventh24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeEleventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventh24 is not null, "Expected twenty-fourth-regime eleventh table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Eleventh (slots 79-84, 91-96, 103-108, 115-120, 127-132, 139-142), dash144/dash145 agree bit-identical.
        int[] eleventh24Idx = [78, 79, 80, 81, 82, 83, 90, 91, 92, 93, 94, 95, 102, 103, 104, 105, 106, 107, 114, 115, 116, 117, 118, 119, 126, 127, 128, 129, 130, 131, 138, 139, 140, 141];
        int[] eleventh24R = [99, 196, 162, 138, 96, 247, 131, 204, 176, 158, 129, 248, 153, 211, 188, 173, 152, 249, 173, 218, 200, 189, 172, 250, 191, 225, 211, 202, 190, 251, 207, 232, 221, 215];
        int[] eleventh24G = [139, 100, 191, 115, 177, 158, 159, 131, 200, 141, 189, 173, 174, 153, 208, 161, 198, 185, 189, 173, 215, 179, 208, 198, 203, 191, 223, 195, 217, 209, 215, 207, 230, 210];
        int[] eleventh24B = [193, 97, 106, 168, 201, 92, 202, 130, 135, 181, 209, 127, 209, 152, 156, 192, 215, 150, 217, 173, 176, 203, 221, 171, 224, 191, 193, 213, 228, 189, 231, 206, 208, 223];
        for (int slot = 0; slot < eleventh24Idx.Length; slot++)
        {
            object?[] eleventh24Args = [eleventh24Idx[slot], null];
            TestAssert.Equal(true, (bool)eleventh24!.Invoke(null, eleventh24Args)!);
            TestAssert.Equal((byte)eleventh24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventh24Args[1])!);
            TestAssert.Equal((byte)eleventh24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventh24Args[1])!);
            TestAssert.Equal((byte)eleventh24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventh24Args[1])!);
        }
        object?[] eleventh24Past = [77, null];
        TestAssert.Equal(false, (bool)eleventh24!.Invoke(null, eleventh24Past)!);
        var fifth24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeFifthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifth24 is not null, "Expected twenty-fourth-regime fifth table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Fifth (slots 85-90), dash144/dash145 agree bit-identical.
        int[] fifth24Idx = [84, 85, 86, 87, 88, 89];
        int[] fifth24R = [115, 200, 169, 148, 112, 248];
        int[] fifth24G = [148, 115, 195, 128, 183, 165];
        int[] fifth24B = [197, 114, 121, 174, 205, 110];
        for (int slot = 0; slot < fifth24Idx.Length; slot++)
        {
            object?[] fifth24Args = [fifth24Idx[slot], null];
            TestAssert.Equal(true, (bool)fifth24!.Invoke(null, fifth24Args)!);
            TestAssert.Equal((byte)fifth24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifth24Args[1])!);
            TestAssert.Equal((byte)fifth24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifth24Args[1])!);
            TestAssert.Equal((byte)fifth24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifth24Args[1])!);
        }
        object?[] fifth24Past = [83, null];
        TestAssert.Equal(false, (bool)fifth24!.Invoke(null, fifth24Past)!);
        var twelfth24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeTwelfthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfth24 is not null, "Expected twenty-fourth-regime twelfth table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Twelfth (slots 97, 99-102), dash144/dash145 agree bit-identical replays within 2.
        int[] twelfth24Idx = [96, 98, 99, 100, 101];
        int[] twelfth24R = [140, 181, 164, 139, 249];
        int[] twelfth24G = [165, 203, 149, 192, 178];
        int[] twelfth24B = [204, 144, 186, 211, 137];
        for (int slot = 0; slot < twelfth24Idx.Length; slot++)
        {
            object?[] twelfth24Args = [twelfth24Idx[slot], null];
            TestAssert.Equal(true, (bool)twelfth24!.Invoke(null, twelfth24Args)!);
            TestAssert.Equal((byte)twelfth24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfth24Args[1])!);
            TestAssert.Equal((byte)twelfth24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfth24Args[1])!);
            TestAssert.Equal((byte)twelfth24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfth24Args[1])!);
        }
        object?[] twelfth24Past = [95, null];
        TestAssert.Equal(false, (bool)twelfth24!.Invoke(null, twelfth24Past)!);
        var thirteenth24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeThirteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(thirteenth24 is not null, "Expected twenty-fourth-regime thirteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Thirteenth (slots 109-114), dash144/dash145 agree bit-identical.
        int[] thirteenth24Idx = [108, 109, 110, 111, 112, 113];
        int[] thirteenth24R = [164, 215, 195, 182, 163, 250];
        int[] thirteenth24G = [182, 165, 212, 171, 204, 192];
        int[] thirteenth24B = [213, 164, 167, 198, 218, 162];
        for (int slot = 0; slot < thirteenth24Idx.Length; slot++)
        {
            object?[] thirteenth24Args = [thirteenth24Idx[slot], null];
            TestAssert.Equal(true, (bool)thirteenth24!.Invoke(null, thirteenth24Args)!);
            TestAssert.Equal((byte)thirteenth24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(thirteenth24Args[1])!);
            TestAssert.Equal((byte)thirteenth24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(thirteenth24Args[1])!);
            TestAssert.Equal((byte)thirteenth24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(thirteenth24Args[1])!);
        }
        object?[] thirteenth24Past = [107, null];
        TestAssert.Equal(false, (bool)thirteenth24!.Invoke(null, thirteenth24Past)!);
        var fourteenth24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeFourteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourteenth24 is not null, "Expected twenty-fourth-regime fourteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Fourteenth (slots 121-126), dash144/dash145 agree bit-identical.
        int[] fourteenth24Idx = [120, 121, 122, 123, 124, 125];
        int[] fourteenth24R = [183, 222, 206, 196, 182, 250];
        int[] fourteenth24G = [196, 183, 220, 188, 213, 204];
        int[] fourteenth24B = [221, 183, 185, 209, 225, 181];
        for (int slot = 0; slot < fourteenth24Idx.Length; slot++)
        {
            object?[] fourteenth24Args = [fourteenth24Idx[slot], null];
            TestAssert.Equal(true, (bool)fourteenth24!.Invoke(null, fourteenth24Args)!);
            TestAssert.Equal((byte)fourteenth24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourteenth24Args[1])!);
            TestAssert.Equal((byte)fourteenth24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourteenth24Args[1])!);
            TestAssert.Equal((byte)fourteenth24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourteenth24Args[1])!);
        }
        object?[] fourteenth24Past = [119, null];
        TestAssert.Equal(false, (bool)fourteenth24!.Invoke(null, fourteenth24Past)!);
        var twentySecondSingle24 = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeTwentySecondSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentySecondSingle24 is not null, "Expected twenty-fourth-regime twentysecondsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime TwentySecondSingle (slot 143), dash144/dash145 agree bit-identical replays within 1.
        int[] twentySecondSingle24Idx = [142];
        int[] twentySecondSingle24R = [205];
        int[] twentySecondSingle24G = [225];
        int[] twentySecondSingle24B = [233];
        for (int slot = 0; slot < twentySecondSingle24Idx.Length; slot++)
        {
            object?[] twentySecondSingle24Args = [twentySecondSingle24Idx[slot], null];
            TestAssert.Equal(true, (bool)twentySecondSingle24!.Invoke(null, twentySecondSingle24Args)!);
            TestAssert.Equal((byte)twentySecondSingle24R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentySecondSingle24Args[1])!);
            TestAssert.Equal((byte)twentySecondSingle24G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentySecondSingle24Args[1])!);
            TestAssert.Equal((byte)twentySecondSingle24B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentySecondSingle24Args[1])!);
        }
        object?[] twentySecondSingle24Past = [137, null];
        TestAssert.Equal(false, (bool)twentySecondSingle24!.Invoke(null, twentySecondSingle24Past)!);
        var tail = rendererType.GetMethod(
            "TryResolveTwentyFourthRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected twenty-fourth-regime tail table to remain inspectable by the Office evidence guard.");
        // Twenty-fourth-regime Tail (slots 146-149), dash145/dash146, dash146/dash147, dash147/dash148 and dash148/dash149 agree bit-identical.
        int[] tailIdx = [144, 145, 146, 147];
        int[] tailR = [211, 234, 224, 220];
        int[] tailG = [218, 211, 232, 216];
        int[] tailB = [233, 211, 212, 227];
        for (int slot = 0; slot < tailIdx.Length; slot++)
        {
            object?[] tailArgs = [tailIdx[slot], null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] tailPast = [143, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailPast)!);
        object?[] tailFuture = [148, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailFuture)!);
        // One-hundred-forty-three points and fewer keep earlier regimes; one-hundred-forty-four-plus take the twenty-fourth rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [143])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [144])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [145])!);
    }
    public static void PptxSyntheticTwentyFifthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwentyFifthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twenty-fifth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var nineteenth25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeNineteenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(nineteenth25 is not null, "Expected twenty-fifth-regime nineteenthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime NineteenthSingle (slot 48), dash150/dash151 agree bit-identical replays within 1.
        int[] nineteenth25Idx = [47];
        int[] nineteenth25R = [211];
        int[] nineteenth25G = [127];
        int[] nineteenth25B = [58];
        for (int slot = 0; slot < nineteenth25Idx.Length; slot++)
        {
            object?[] nineteenth25Args = [nineteenth25Idx[slot], null];
            TestAssert.Equal(true, (bool)nineteenth25!.Invoke(null, nineteenth25Args)!);
            TestAssert.Equal((byte)nineteenth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(nineteenth25Args[1])!);
            TestAssert.Equal((byte)nineteenth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(nineteenth25Args[1])!);
            TestAssert.Equal((byte)nineteenth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(nineteenth25Args[1])!);
        }
        object?[] nineteenth25Past = [46, null];
        TestAssert.Equal(false, (bool)nineteenth25!.Invoke(null, nineteenth25Past)!);
        var seventeenth25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeSeventeenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenth25 is not null, "Expected twenty-fifth-regime seventeenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Seventeenth (slots 60, 66, 73-75, 77-78), dash150/dash151 agree bit-identical replays within 1.
        int[] seventeenth25Idx = [59, 65, 72, 73, 74, 76, 77];
        int[] seventeenth25R = [225, 233, 77, 189, 152, 74, 243];
        int[] seventeenth25G = [136, 141, 127, 78, 184, 169, 147];
        int[] seventeenth25B = [63, 66, 186, 75, 87, 194, 69];
        for (int slot = 0; slot < seventeenth25Idx.Length; slot++)
        {
            object?[] seventeenth25Args = [seventeenth25Idx[slot], null];
            TestAssert.Equal(true, (bool)seventeenth25!.Invoke(null, seventeenth25Args)!);
            TestAssert.Equal((byte)seventeenth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenth25Args[1])!);
            TestAssert.Equal((byte)seventeenth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenth25Args[1])!);
            TestAssert.Equal((byte)seventeenth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenth25Args[1])!);
        }
        object?[] seventeenth25Past = [58, null];
        TestAssert.Equal(false, (bool)seventeenth25!.Invoke(null, seventeenth25Past)!);
        var twentyThird25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeTwentyThirdFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentyThird25 is not null, "Expected twenty-fifth-regime twentythird table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime TwentyThird (slots 79-84), dash150/dash151 agree bit-identical.
        int[] twentyThird25Idx = [78, 79, 80, 81, 82, 83];
        int[] twentyThird25R = [88, 194, 158, 132, 84, 247];
        int[] twentyThird25G = [133, 88, 189, 106, 174, 153];
        int[] twentyThird25B = [191, 86, 96, 165, 199, 80];
        for (int slot = 0; slot < twentyThird25Idx.Length; slot++)
        {
            object?[] twentyThird25Args = [twentyThird25Idx[slot], null];
            TestAssert.Equal(true, (bool)twentyThird25!.Invoke(null, twentyThird25Args)!);
            TestAssert.Equal((byte)twentyThird25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentyThird25Args[1])!);
            TestAssert.Equal((byte)twentyThird25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentyThird25Args[1])!);
            TestAssert.Equal((byte)twentyThird25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentyThird25Args[1])!);
        }
        object?[] twentyThird25Past = [77, null];
        TestAssert.Equal(false, (bool)twentyThird25!.Invoke(null, twentyThird25Past)!);
        var seventh25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeSeventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventh25 is not null, "Expected twenty-fifth-regime seventh table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Seventh (slots 85-90, 121-126), dash150/dash151 agree bit-identical.
        int[] seventh25Idx = [84, 85, 86, 87, 88, 89, 120, 121, 122, 123, 124, 125];
        int[] seventh25R = [106, 197, 165, 142, 103, 248, 175, 219, 201, 190, 174, 250];
        int[] seventh25G = [143, 106, 193, 120, 180, 160, 190, 175, 216, 180, 209, 199];
        int[] seventh25B = [195, 104, 112, 171, 203, 100, 217, 175, 177, 204, 222, 173];
        for (int slot = 0; slot < seventh25Idx.Length; slot++)
        {
            object?[] seventh25Args = [seventh25Idx[slot], null];
            TestAssert.Equal(true, (bool)seventh25!.Invoke(null, seventh25Args)!);
            TestAssert.Equal((byte)seventh25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventh25Args[1])!);
            TestAssert.Equal((byte)seventh25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventh25Args[1])!);
            TestAssert.Equal((byte)seventh25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventh25Args[1])!);
        }
        object?[] seventh25Past = [83, null];
        TestAssert.Equal(false, (bool)seventh25!.Invoke(null, seventh25Past)!);
        var fifteenth25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeFifteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifteenth25 is not null, "Expected twenty-fifth-regime fifteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Fifteenth (slots 91-96), dash150/dash151 agree bit-identical.
        int[] fifteenth25Idx = [90, 91, 92, 93, 94, 95];
        int[] fifteenth25R = [120, 201, 171, 151, 118, 248];
        int[] fifteenth25G = [152, 121, 197, 132, 185, 167];
        int[] fifteenth25B = [198, 119, 126, 177, 206, 116];
        for (int slot = 0; slot < fifteenth25Idx.Length; slot++)
        {
            object?[] fifteenth25Args = [fifteenth25Idx[slot], null];
            TestAssert.Equal(true, (bool)fifteenth25!.Invoke(null, fifteenth25Args)!);
            TestAssert.Equal((byte)fifteenth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifteenth25Args[1])!);
            TestAssert.Equal((byte)fifteenth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifteenth25Args[1])!);
            TestAssert.Equal((byte)fifteenth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifteenth25Args[1])!);
        }
        object?[] fifteenth25Past = [89, null];
        TestAssert.Equal(false, (bool)fifteenth25!.Invoke(null, fifteenth25Past)!);
        var sixth25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeSixthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixth25 is not null, "Expected twenty-fifth-regime sixth table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Sixth (slots 97, 99-100), dash150/dash151 agree bit-identical replays within 2.
        int[] sixth25Idx = [96, 98, 99];
        int[] sixth25R = [133, 177, 160];
        int[] sixth25G = [160, 201, 143];
        int[] sixth25B = [202, 138, 182];
        for (int slot = 0; slot < sixth25Idx.Length; slot++)
        {
            object?[] sixth25Args = [sixth25Idx[slot], null];
            TestAssert.Equal(true, (bool)sixth25!.Invoke(null, sixth25Args)!);
            TestAssert.Equal((byte)sixth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixth25Args[1])!);
            TestAssert.Equal((byte)sixth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixth25Args[1])!);
            TestAssert.Equal((byte)sixth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixth25Args[1])!);
        }
        object?[] sixth25Past = [95, null];
        TestAssert.Equal(false, (bool)sixth25!.Invoke(null, sixth25Past)!);
        var ninth25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeNinthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninth25 is not null, "Expected twenty-fifth-regime ninth table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Ninth (slots 98, 101-102, 127-132), dash150/dash151 agree bit-identical replays within 2.
        int[] ninth25Idx = [97, 100, 101, 126, 127, 128, 129, 130, 131];
        int[] ninth25R = [206, 136, 249, 185, 223, 207, 197, 184, 251];
        int[] ninth25G = [138, 192, 177, 198, 185, 220, 189, 214, 205];
        int[] ninth25B = [137, 210, 134, 221, 184, 187, 210, 225, 183];
        for (int slot = 0; slot < ninth25Idx.Length; slot++)
        {
            object?[] ninth25Args = [ninth25Idx[slot], null];
            TestAssert.Equal(true, (bool)ninth25!.Invoke(null, ninth25Args)!);
            TestAssert.Equal((byte)ninth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninth25Args[1])!);
            TestAssert.Equal((byte)ninth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninth25Args[1])!);
            TestAssert.Equal((byte)ninth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninth25Args[1])!);
        }
        object?[] ninth25Past = [96, null];
        TestAssert.Equal(false, (bool)ninth25!.Invoke(null, ninth25Past)!);
        var overflow25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflow25 is not null, "Expected twenty-fifth-regime overflow table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Overflow (slots 103-106, 134, 139-142), dash150/dash151 agree bit-identical replays within 2.
        int[] overflow25Idx = [102, 103, 104, 105, 133, 138, 139, 140, 141];
        int[] overflow25R = [147, 209, 185, 169, 226, 200, 229, 217, 209];
        int[] overflow25G = [169, 147, 205, 155, 194, 209, 200, 227, 203];
        int[] overflow25B = [207, 146, 150, 189, 194, 228, 199, 201, 219];
        for (int slot = 0; slot < overflow25Idx.Length; slot++)
        {
            object?[] overflow25Args = [overflow25Idx[slot], null];
            TestAssert.Equal(true, (bool)overflow25!.Invoke(null, overflow25Args)!);
            TestAssert.Equal((byte)overflow25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflow25Args[1])!);
            TestAssert.Equal((byte)overflow25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflow25Args[1])!);
            TestAssert.Equal((byte)overflow25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflow25Args[1])!);
        }
        object?[] overflow25Past = [101, null];
        TestAssert.Equal(false, (bool)overflow25!.Invoke(null, overflow25Past)!);
        var fourth25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeFourthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourth25 is not null, "Expected twenty-fifth-regime fourth table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Fourth (slots 107-108), dash150/dash151 agree bit-identical.
        int[] fourth25Idx = [106, 107];
        int[] fourth25R = [145, 249];
        int[] fourth25G = [195, 181];
        int[] fourth25B = [213, 144];
        for (int slot = 0; slot < fourth25Idx.Length; slot++)
        {
            object?[] fourth25Args = [fourth25Idx[slot], null];
            TestAssert.Equal(true, (bool)fourth25!.Invoke(null, fourth25Args)!);
            TestAssert.Equal((byte)fourth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourth25Args[1])!);
            TestAssert.Equal((byte)fourth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourth25Args[1])!);
            TestAssert.Equal((byte)fourth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourth25Args[1])!);
        }
        object?[] fourth25Past = [105, null];
        TestAssert.Equal(false, (bool)fourth25!.Invoke(null, fourth25Past)!);
        var eighth25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeEighthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighth25 is not null, "Expected twenty-fifth-regime eighth table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Eighth (slots 109-114), dash150/dash151 agree bit-identical.
        int[] eighth25Idx = [108, 109, 110, 111, 112, 113];
        int[] eighth25R = [157, 212, 190, 176, 156, 249];
        int[] eighth25G = [177, 157, 209, 164, 200, 187];
        int[] eighth25B = [210, 156, 160, 194, 216, 154];
        for (int slot = 0; slot < eighth25Idx.Length; slot++)
        {
            object?[] eighth25Args = [eighth25Idx[slot], null];
            TestAssert.Equal(true, (bool)eighth25!.Invoke(null, eighth25Args)!);
            TestAssert.Equal((byte)eighth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighth25Args[1])!);
            TestAssert.Equal((byte)eighth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighth25Args[1])!);
            TestAssert.Equal((byte)eighth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighth25Args[1])!);
        }
        object?[] eighth25Past = [107, null];
        TestAssert.Equal(false, (bool)eighth25!.Invoke(null, eighth25Past)!);
        var thirteenth25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeThirteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(thirteenth25 is not null, "Expected twenty-fifth-regime thirteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Thirteenth (slots 115-120), dash150/dash151 agree bit-identical replays within 2.
        int[] thirteenth25Idx = [114, 115, 116, 117, 118, 119];
        int[] thirteenth25R = [164, 215, 195, 182, 163, 250];
        int[] thirteenth25G = [182, 165, 212, 171, 204, 192];
        int[] thirteenth25B = [213, 164, 167, 198, 218, 162];
        for (int slot = 0; slot < thirteenth25Idx.Length; slot++)
        {
            object?[] thirteenth25Args = [thirteenth25Idx[slot], null];
            TestAssert.Equal(true, (bool)thirteenth25!.Invoke(null, thirteenth25Args)!);
            TestAssert.Equal((byte)thirteenth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(thirteenth25Args[1])!);
            TestAssert.Equal((byte)thirteenth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(thirteenth25Args[1])!);
            TestAssert.Equal((byte)thirteenth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(thirteenth25Args[1])!);
        }
        object?[] thirteenth25Past = [113, null];
        TestAssert.Equal(false, (bool)thirteenth25!.Invoke(null, thirteenth25Past)!);
        var eleventh25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeEleventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventh25 is not null, "Expected twenty-fifth-regime eleventh table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Eleventh (slots 133, 135-136, 138, 145-148), dash150/dash151 agree bit-identical replays within 2.
        int[] eleventh25Idx = [132, 134, 135, 137, 144, 145, 146, 147];
        int[] eleventh25R = [191, 211, 202, 251, 207, 232, 221, 215];
        int[] eleventh25G = [203, 223, 195, 209, 215, 207, 230, 210];
        int[] eleventh25B = [224, 193, 213, 189, 231, 206, 208, 223];
        for (int slot = 0; slot < eleventh25Idx.Length; slot++)
        {
            object?[] eleventh25Args = [eleventh25Idx[slot], null];
            TestAssert.Equal(true, (bool)eleventh25!.Invoke(null, eleventh25Args)!);
            TestAssert.Equal((byte)eleventh25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventh25Args[1])!);
            TestAssert.Equal((byte)eleventh25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventh25Args[1])!);
            TestAssert.Equal((byte)eleventh25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventh25Args[1])!);
        }
        object?[] eleventh25Past = [131, null];
        TestAssert.Equal(false, (bool)eleventh25!.Invoke(null, eleventh25Past)!);
        var twelfthSingle = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeTwelfthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfthSingle is not null, "Expected twenty-fifth-regime twelfthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime TwelfthSingle (slot 137), dash150/dash151 agree bit-identical replays within 1.
        int[] twelfthSingleIdx = [136];
        int[] twelfthSingleR = [193];
        int[] twelfthSingleG = [219];
        int[] twelfthSingleB = [229];
        for (int slot = 0; slot < twelfthSingleIdx.Length; slot++)
        {
            object?[] twelfthSingleArgs = [twelfthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)twelfthSingle!.Invoke(null, twelfthSingleArgs)!);
            TestAssert.Equal((byte)twelfthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfthSingleArgs[1])!);
            TestAssert.Equal((byte)twelfthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfthSingleArgs[1])!);
            TestAssert.Equal((byte)twelfthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfthSingleArgs[1])!);
        }
        object?[] twelfthSinglePast = [135, null];
        TestAssert.Equal(false, (bool)twelfthSingle!.Invoke(null, twelfthSinglePast)!);
        var sixteenth25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeSixteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixteenth25 is not null, "Expected twenty-fifth-regime sixteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Sixteenth (slots 143-144), dash150/dash151 agree bit-identical.
        int[] sixteenth25Idx = [142, 143];
        int[] sixteenth25R = [199, 251];
        int[] sixteenth25G = [222, 215];
        int[] sixteenth25B = [231, 199];
        for (int slot = 0; slot < sixteenth25Idx.Length; slot++)
        {
            object?[] sixteenth25Args = [sixteenth25Idx[slot], null];
            TestAssert.Equal(true, (bool)sixteenth25!.Invoke(null, sixteenth25Args)!);
            TestAssert.Equal((byte)sixteenth25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixteenth25Args[1])!);
            TestAssert.Equal((byte)sixteenth25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixteenth25Args[1])!);
            TestAssert.Equal((byte)sixteenth25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixteenth25Args[1])!);
        }
        object?[] sixteenth25Past = [141, null];
        TestAssert.Equal(false, (bool)sixteenth25!.Invoke(null, sixteenth25Past)!);
        var twentySecondSingle25 = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeTwentySecondSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentySecondSingle25 is not null, "Expected twenty-fifth-regime twentysecondsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime TwentySecondSingle (slot 149), dash150/dash151 agree bit-identical replays within 1.
        int[] twentySecondSingle25Idx = [148];
        int[] twentySecondSingle25R = [205];
        int[] twentySecondSingle25G = [225];
        int[] twentySecondSingle25B = [233];
        for (int slot = 0; slot < twentySecondSingle25Idx.Length; slot++)
        {
            object?[] twentySecondSingle25Args = [twentySecondSingle25Idx[slot], null];
            TestAssert.Equal(true, (bool)twentySecondSingle25!.Invoke(null, twentySecondSingle25Args)!);
            TestAssert.Equal((byte)twentySecondSingle25R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentySecondSingle25Args[1])!);
            TestAssert.Equal((byte)twentySecondSingle25G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentySecondSingle25Args[1])!);
            TestAssert.Equal((byte)twentySecondSingle25B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentySecondSingle25Args[1])!);
        }
        object?[] twentySecondSingle25Past = [147, null];
        TestAssert.Equal(false, (bool)twentySecondSingle25!.Invoke(null, twentySecondSingle25Past)!);
        var tail = rendererType.GetMethod(
            "TryResolveTwentyFifthRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected twenty-fifth-regime tail table to remain inspectable by the Office evidence guard.");
        // Twenty-fifth-regime Tail (slots 152-155), dash151/dash152, dash152/dash153, dash153/dash154 and dash154/dash155 agree bit-identical.
        int[] tailIdx = [150, 151, 152, 153];
        int[] tailR = [211, 234, 224, 220];
        int[] tailG = [218, 211, 232, 216];
        int[] tailB = [233, 211, 212, 227];
        for (int slot = 0; slot < tailIdx.Length; slot++)
        {
            object?[] tailArgs = [tailIdx[slot], null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] tailPast = [149, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailPast)!);
        object?[] tailFuture = [154, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailFuture)!);
        // One-hundred-forty-nine points and fewer keep earlier regimes; one-hundred-fifty-plus take the twenty-fifth rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [149])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [150])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [151])!);
    }
    public static void PptxSyntheticTwentySixthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwentySixthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twenty-sixth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var seventeenth26 = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeSeventeenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenth26 is not null, "Expected twenty-sixth-regime seventeenthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime SeventeenthSingle (slot 78), dash156/dash157 agree bit-identical replays within 2.
        int[] seventeenth26Idx = [77];
        int[] seventeenth26R = [243];
        int[] seventeenth26G = [147];
        int[] seventeenth26B = [69];
        for (int slot = 0; slot < seventeenth26Idx.Length; slot++)
        {
            object?[] seventeenth26Args = [seventeenth26Idx[slot], null];
            TestAssert.Equal(true, (bool)seventeenth26!.Invoke(null, seventeenth26Args)!);
            TestAssert.Equal((byte)seventeenth26R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenth26Args[1])!);
            TestAssert.Equal((byte)seventeenth26G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenth26Args[1])!);
            TestAssert.Equal((byte)seventeenth26B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenth26Args[1])!);
        }
        object?[] seventeenth26Past = [76, null];
        TestAssert.Equal(false, (bool)seventeenth26!.Invoke(null, seventeenth26Past)!);
        var eighteenth26 = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeEighteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenth26 is not null, "Expected twenty-sixth-regime eighteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime Eighteenth (slots 79-84, 149), dash156/dash157 agree bit-identical replays within 1.
        int[] eighteenth26Idx = [78, 79, 80, 81, 82, 83, 148];
        int[] eighteenth26R = [79, 192, 155, 128, 75, 247, 202];
        int[] eighteenth26G = [129, 80, 187, 100, 172, 150, 223];
        int[] eighteenth26B = [189, 77, 89, 162, 198, 70, 232];
        for (int slot = 0; slot < eighteenth26Idx.Length; slot++)
        {
            object?[] eighteenth26Args = [eighteenth26Idx[slot], null];
            TestAssert.Equal(true, (bool)eighteenth26!.Invoke(null, eighteenth26Args)!);
            TestAssert.Equal((byte)eighteenth26R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenth26Args[1])!);
            TestAssert.Equal((byte)eighteenth26G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenth26Args[1])!);
            TestAssert.Equal((byte)eighteenth26B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenth26Args[1])!);
        }
        object?[] eighteenth26Past = [77, null];
        TestAssert.Equal(false, (bool)eighteenth26!.Invoke(null, eighteenth26Past)!);
        var eleventh26 = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeEleventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventh26 is not null, "Expected twenty-sixth-regime eleventh table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime Eleventh (slots 85-90, 109-114), dash156/dash157 agree bit-identical replays within 2.
        int[] eleventh26Idx = [84, 85, 86, 87, 88, 89, 108, 109, 110, 111, 112, 113];
        int[] eleventh26R = [99, 196, 162, 138, 96, 247, 153, 211, 188, 173, 152, 249];
        int[] eleventh26G = [139, 100, 191, 115, 177, 158, 174, 153, 208, 161, 198, 185];
        int[] eleventh26B = [193, 97, 106, 168, 201, 92, 209, 152, 156, 192, 215, 150];
        for (int slot = 0; slot < eleventh26Idx.Length; slot++)
        {
            object?[] eleventh26Args = [eleventh26Idx[slot], null];
            TestAssert.Equal(true, (bool)eleventh26!.Invoke(null, eleventh26Args)!);
            TestAssert.Equal((byte)eleventh26R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventh26Args[1])!);
            TestAssert.Equal((byte)eleventh26G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventh26Args[1])!);
            TestAssert.Equal((byte)eleventh26B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventh26Args[1])!);
        }
        object?[] eleventh26Past = [83, null];
        TestAssert.Equal(false, (bool)eleventh26!.Invoke(null, eleventh26Past)!);
        var fifth26 = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeFifthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifth26 is not null, "Expected twenty-sixth-regime fifth table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime Fifth (slots 91-96, 115-120), dash156/dash157 agree bit-identical.
        int[] fifth26Idx = [90, 91, 92, 93, 94, 95, 114, 115, 116, 117, 118, 119];
        int[] fifth26R = [115, 200, 169, 148, 112, 248, 161, 214, 192, 179, 160, 249];
        int[] fifth26G = [148, 115, 195, 128, 183, 165, 180, 161, 210, 168, 202, 190];
        int[] fifth26B = [197, 114, 121, 174, 205, 110, 212, 160, 164, 196, 217, 158];
        for (int slot = 0; slot < fifth26Idx.Length; slot++)
        {
            object?[] fifth26Args = [fifth26Idx[slot], null];
            TestAssert.Equal(true, (bool)fifth26!.Invoke(null, fifth26Args)!);
            TestAssert.Equal((byte)fifth26R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifth26Args[1])!);
            TestAssert.Equal((byte)fifth26G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifth26Args[1])!);
            TestAssert.Equal((byte)fifth26B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifth26Args[1])!);
        }
        object?[] fifth26Past = [89, null];
        TestAssert.Equal(false, (bool)fifth26!.Invoke(null, fifth26Past)!);
        var sixteenth26 = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeSixteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixteenth26 is not null, "Expected twenty-sixth-regime sixteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime Sixteenth (slots 97-102, 150), dash156/dash157 agree bit-identical replays within 1.
        int[] sixteenth26Idx = [96, 97, 98, 99, 100, 101, 149];
        int[] sixteenth26R = [128, 203, 175, 156, 126, 248, 251];
        int[] sixteenth26G = [157, 129, 199, 139, 188, 171, 215];
        int[] sixteenth26B = [201, 127, 133, 180, 208, 124, 199];
        for (int slot = 0; slot < sixteenth26Idx.Length; slot++)
        {
            object?[] sixteenth26Args = [sixteenth26Idx[slot], null];
            TestAssert.Equal(true, (bool)sixteenth26!.Invoke(null, sixteenth26Args)!);
            TestAssert.Equal((byte)sixteenth26R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixteenth26Args[1])!);
            TestAssert.Equal((byte)sixteenth26G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixteenth26Args[1])!);
            TestAssert.Equal((byte)sixteenth26B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixteenth26Args[1])!);
        }
        object?[] sixteenth26Past = [95, null];
        TestAssert.Equal(false, (bool)sixteenth26!.Invoke(null, sixteenth26Past)!);
        var twelfth26 = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeTwelfthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfth26 is not null, "Expected twenty-sixth-regime twelfth table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime Twelfth (slots 103-108, 127-132, 143-144, 151-154), dash156/dash157 agree bit-identical.
        int[] twelfth26Idx = [102, 103, 104, 105, 106, 107, 126, 127, 128, 129, 130, 131, 142, 143, 150, 151, 152, 153];
        int[] twelfth26R = [140, 207, 181, 164, 139, 249, 178, 220, 203, 192, 177, 250, 193, 251, 208, 232, 222, 216];
        int[] twelfth26G = [165, 140, 203, 149, 192, 178, 193, 179, 218, 184, 210, 201, 219, 211, 216, 208, 231, 211];
        int[] twelfth26B = [204, 139, 144, 186, 211, 137, 219, 178, 181, 206, 223, 176, 229, 193, 232, 208, 209, 224];
        for (int slot = 0; slot < twelfth26Idx.Length; slot++)
        {
            object?[] twelfth26Args = [twelfth26Idx[slot], null];
            TestAssert.Equal(true, (bool)twelfth26!.Invoke(null, twelfth26Args)!);
            TestAssert.Equal((byte)twelfth26R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfth26Args[1])!);
            TestAssert.Equal((byte)twelfth26G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfth26Args[1])!);
            TestAssert.Equal((byte)twelfth26B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfth26Args[1])!);
        }
        object?[] twelfth26Past = [101, null];
        TestAssert.Equal(false, (bool)twelfth26!.Invoke(null, twelfth26Past)!);
        var overflow26 = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflow26 is not null, "Expected twenty-sixth-regime overflow table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime Overflow (slots 121-124, 139-142, 145-148), dash156/dash157 agree bit-identical replays within 2.
        int[] overflow26Idx = [120, 121, 122, 123, 138, 139, 140, 141, 144, 145, 146, 147];
        int[] overflow26R = [170, 217, 198, 186, 194, 226, 213, 205, 200, 229, 218, 209];
        int[] overflow26G = [186, 170, 214, 176, 205, 194, 224, 198, 209, 200, 228, 203];
        int[] overflow26B = [215, 169, 172, 201, 225, 194, 196, 215, 228, 199, 204, 219];
        for (int slot = 0; slot < overflow26Idx.Length; slot++)
        {
            object?[] overflow26Args = [overflow26Idx[slot], null];
            TestAssert.Equal(true, (bool)overflow26!.Invoke(null, overflow26Args)!);
            TestAssert.Equal((byte)overflow26R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflow26Args[1])!);
            TestAssert.Equal((byte)overflow26G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflow26Args[1])!);
            TestAssert.Equal((byte)overflow26B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflow26Args[1])!);
        }
        object?[] overflow26Past = [119, null];
        TestAssert.Equal(false, (bool)overflow26!.Invoke(null, overflow26Past)!);
        var sixth26 = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeSixthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixth26 is not null, "Expected twenty-sixth-regime sixth table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime Sixth (slots 125-126), dash156/dash157 agree bit-identical.
        int[] sixth26Idx = [124, 125];
        int[] sixth26R = [169, 250];
        int[] sixth26G = [206, 195];
        int[] sixth26B = [220, 168];
        for (int slot = 0; slot < sixth26Idx.Length; slot++)
        {
            object?[] sixth26Args = [sixth26Idx[slot], null];
            TestAssert.Equal(true, (bool)sixth26!.Invoke(null, sixth26Args)!);
            TestAssert.Equal((byte)sixth26R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixth26Args[1])!);
            TestAssert.Equal((byte)sixth26G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixth26Args[1])!);
            TestAssert.Equal((byte)sixth26B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixth26Args[1])!);
        }
        object?[] sixth26Past = [123, null];
        TestAssert.Equal(false, (bool)sixth26!.Invoke(null, sixth26Past)!);
        var fifteenth26 = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeFifteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifteenth26 is not null, "Expected twenty-sixth-regime fifteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime Fifteenth (slots 133-138), dash156/dash157 agree bit-identical.
        int[] fifteenth26Idx = [132, 133, 134, 135, 136, 137];
        int[] fifteenth26R = [186, 223, 208, 199, 186, 251];
        int[] fifteenth26G = [199, 187, 221, 191, 215, 206];
        int[] fifteenth26B = [222, 186, 188, 211, 226, 185];
        for (int slot = 0; slot < fifteenth26Idx.Length; slot++)
        {
            object?[] fifteenth26Args = [fifteenth26Idx[slot], null];
            TestAssert.Equal(true, (bool)fifteenth26!.Invoke(null, fifteenth26Args)!);
            TestAssert.Equal((byte)fifteenth26R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifteenth26Args[1])!);
            TestAssert.Equal((byte)fifteenth26G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifteenth26Args[1])!);
            TestAssert.Equal((byte)fifteenth26B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifteenth26Args[1])!);
        }
        object?[] fifteenth26Past = [131, null];
        TestAssert.Equal(false, (bool)fifteenth26!.Invoke(null, fifteenth26Past)!);
        var lateSingle = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeLateSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(lateSingle is not null, "Expected twenty-sixth-regime latesingle table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime LateSingles (slot 155 dash156/dash157 agree, slot 157 dash157/dash158 agree, slot 158 dash158/dash159 agree, slot 159 dash159/dash160 agree).
        int[] lateSingleIdx = [154, 156, 157, 158];
        int[] lateSingleR = [208, 215, 235, 227];
        int[] lateSingleG = [227, 222, 215, 234];
        int[] lateSingleB = [234, 235, 214, 216];
        for (int slot = 0; slot < lateSingleIdx.Length; slot++)
        {
            object?[] lateSingleArgs = [lateSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)lateSingle!.Invoke(null, lateSingleArgs)!);
            TestAssert.Equal((byte)lateSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(lateSingleArgs[1])!);
            TestAssert.Equal((byte)lateSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(lateSingleArgs[1])!);
            TestAssert.Equal((byte)lateSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(lateSingleArgs[1])!);
        }
        object?[] lateSinglePast = [153, null];
        TestAssert.Equal(false, (bool)lateSingle!.Invoke(null, lateSinglePast)!);
        var tail = rendererType.GetMethod(
            "TryResolveTwentySixthRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected twenty-sixth-regime tail table to remain inspectable by the Office evidence guard.");
        // Twenty-sixth-regime Tail (slot 161), dash160/dash161 agree bit-identical.
        int[] tailIdx = [159];
        int[] tailR = [220];
        int[] tailG = [216];
        int[] tailB = [227];
        for (int slot = 0; slot < tailIdx.Length; slot++)
        {
            object?[] tailArgs = [tailIdx[slot], null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] tailPast = [158, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailPast)!);
        // One-hundred-fifty-five points and fewer keep earlier regimes; one-hundred-fifty-six-plus take the twenty-sixth rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [155])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [156])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [157])!);
    }
    public static void PptxSyntheticTwentySeventhVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwentySeventhVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twenty-seventh-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var seventeenth27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeSeventeenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenth27 is not null, "Expected twenty-seventh-regime seventeenth table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Seventeenth (slots 32, 36, 44, 48, 54, 68-69, 71-72, 79-81, 83-84), dash162/dash163 agree bit-identical replays within 2.
        int[] seventeenth27Idx = [31, 35, 43, 47, 53, 67, 68, 70, 71, 78, 79, 80, 82, 83];
        int[] seventeenth27R = [151, 195, 159, 206, 216, 181, 146, 70, 233, 77, 189, 152, 74, 243];
        int[] seventeenth27G = [61, 118, 65, 124, 131, 75, 177, 162, 141, 127, 78, 184, 169, 147];
        int[] seventeenth27B = [59, 53, 63, 57, 60, 72, 84, 187, 66, 186, 75, 87, 194, 69];
        for (int slot = 0; slot < seventeenth27Idx.Length; slot++)
        {
            object?[] seventeenth27Args = [seventeenth27Idx[slot], null];
            TestAssert.Equal(true, (bool)seventeenth27!.Invoke(null, seventeenth27Args)!);
            TestAssert.Equal((byte)seventeenth27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenth27Args[1])!);
            TestAssert.Equal((byte)seventeenth27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenth27Args[1])!);
            TestAssert.Equal((byte)seventeenth27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenth27Args[1])!);
        }
        object?[] seventeenth27Past = [30, null];
        TestAssert.Equal(false, (bool)seventeenth27!.Invoke(null, seventeenth27Past)!);
        var nineteenth27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeNineteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(nineteenth27 is not null, "Expected twenty-seventh-regime nineteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Nineteenth (slots 42, 60), dash162/dash163 agree bit-identical replays within 2.
        int[] nineteenth27Idx = [41, 59];
        int[] nineteenth27R = [201, 219];
        int[] nineteenth27G = [121, 132];
        int[] nineteenth27B = [55, 61];
        for (int slot = 0; slot < nineteenth27Idx.Length; slot++)
        {
            object?[] nineteenth27Args = [nineteenth27Idx[slot], null];
            TestAssert.Equal(true, (bool)nineteenth27!.Invoke(null, nineteenth27Args)!);
            TestAssert.Equal((byte)nineteenth27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(nineteenth27Args[1])!);
            TestAssert.Equal((byte)nineteenth27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(nineteenth27Args[1])!);
            TestAssert.Equal((byte)nineteenth27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(nineteenth27Args[1])!);
        }
        object?[] nineteenth27Past = [40, null];
        TestAssert.Equal(false, (bool)nineteenth27!.Invoke(null, nineteenth27Past)!);
        var seventh27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeSeventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventh27 is not null, "Expected twenty-seventh-regime seventh table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Seventh (slots 37, 91-96), dash162/dash163 agree bit-identical replays within 2.
        int[] seventh27Idx = [36, 90, 91, 92, 93, 94, 95];
        int[] seventh27R = [62, 106, 197, 165, 142, 103, 248];
        int[] seventh27G = [102, 143, 106, 193, 120, 180, 160];
        int[] seventh27B = [151, 195, 104, 112, 171, 203, 100];
        for (int slot = 0; slot < seventh27Idx.Length; slot++)
        {
            object?[] seventh27Args = [seventh27Idx[slot], null];
            TestAssert.Equal(true, (bool)seventh27!.Invoke(null, seventh27Args)!);
            TestAssert.Equal((byte)seventh27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventh27Args[1])!);
            TestAssert.Equal((byte)seventh27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventh27Args[1])!);
            TestAssert.Equal((byte)seventh27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventh27Args[1])!);
        }
        object?[] seventh27Past = [34, null];
        TestAssert.Equal(false, (bool)seventh27!.Invoke(null, seventh27Past)!);
        var eighteenth27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeEighteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenth27 is not null, "Expected twenty-seventh-regime eighteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Eighteenth (slots 73-75, 77-78, 110, 155), dash162/dash163 agree bit-identical replays within 1.
        int[] eighteenth27Idx = [72, 73, 74, 76, 77, 109, 154];
        int[] eighteenth27R = [76, 186, 150, 72, 239, 207, 202];
        int[] eighteenth27G = [125, 77, 181, 166, 145, 143, 223];
        int[] eighteenth27B = [183, 74, 86, 192, 67, 142, 232];
        for (int slot = 0; slot < eighteenth27Idx.Length; slot++)
        {
            object?[] eighteenth27Args = [eighteenth27Idx[slot], null];
            TestAssert.Equal(true, (bool)eighteenth27!.Invoke(null, eighteenth27Args)!);
            TestAssert.Equal((byte)eighteenth27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenth27Args[1])!);
            TestAssert.Equal((byte)eighteenth27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenth27Args[1])!);
            TestAssert.Equal((byte)eighteenth27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenth27Args[1])!);
        }
        object?[] eighteenth27Past = [71, null];
        TestAssert.Equal(false, (bool)eighteenth27!.Invoke(null, eighteenth27Past)!);
        var eleventh27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeEleventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventh27 is not null, "Expected twenty-seventh-regime eleventh table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Eleventh (slots 103-108, 115-120, 127, 131), dash162/dash163 agree bit-identical replays within 1.
        int[] eleventh27Idx = [102, 103, 104, 105, 106, 107, 114, 115, 116, 117, 118, 119, 126, 130];
        int[] eleventh27R = [131, 204, 176, 158, 129, 248, 153, 211, 188, 173, 152, 249, 173, 172];
        int[] eleventh27G = [159, 131, 200, 141, 189, 173, 174, 153, 208, 161, 198, 185, 189, 208];
        int[] eleventh27B = [202, 130, 135, 181, 209, 127, 209, 152, 156, 192, 215, 150, 217, 221];
        for (int slot = 0; slot < eleventh27Idx.Length; slot++)
        {
            object?[] eleventh27Args = [eleventh27Idx[slot], null];
            TestAssert.Equal(true, (bool)eleventh27!.Invoke(null, eleventh27Args)!);
            TestAssert.Equal((byte)eleventh27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventh27Args[1])!);
            TestAssert.Equal((byte)eleventh27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventh27Args[1])!);
            TestAssert.Equal((byte)eleventh27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventh27Args[1])!);
        }
        object?[] eleventh27Past = [83, null];
        TestAssert.Equal(false, (bool)eleventh27!.Invoke(null, eleventh27Past)!);
        var twentyThird27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeTwentyThirdFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentyThird27 is not null, "Expected twenty-seventh-regime twentythird table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime TwentyThird (slots 85-90), dash162/dash163 agree bit-identical.
        int[] twentyThird27Idx = [84, 85, 86, 87, 88, 89];
        int[] twentyThird27R = [88, 194, 158, 132, 84, 247];
        int[] twentyThird27G = [133, 88, 189, 106, 174, 153];
        int[] twentyThird27B = [191, 86, 96, 165, 199, 80];
        for (int slot = 0; slot < twentyThird27Idx.Length; slot++)
        {
            object?[] twentyThird27Args = [twentyThird27Idx[slot], null];
            TestAssert.Equal(true, (bool)twentyThird27!.Invoke(null, twentyThird27Args)!);
            TestAssert.Equal((byte)twentyThird27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentyThird27Args[1])!);
            TestAssert.Equal((byte)twentyThird27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentyThird27Args[1])!);
            TestAssert.Equal((byte)twentyThird27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentyThird27Args[1])!);
        }
        object?[] twentyThird27Past = [83, null];
        TestAssert.Equal(false, (bool)twentyThird27!.Invoke(null, twentyThird27Past)!);
        var fifteenth27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeFifteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fifteenth27 is not null, "Expected twenty-seventh-regime fifteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Fifteenth (slots 97-102, 139-144), dash162/dash163 agree bit-identical.
        int[] fifteenth27Idx = [96, 97, 98, 99, 100, 101, 138, 139, 140, 141, 142, 143];
        int[] fifteenth27R = [120, 201, 171, 151, 118, 248, 186, 223, 208, 199, 186, 251];
        int[] fifteenth27G = [152, 121, 197, 132, 185, 167, 199, 187, 221, 191, 215, 206];
        int[] fifteenth27B = [198, 119, 126, 177, 206, 116, 222, 186, 188, 211, 226, 185];
        for (int slot = 0; slot < fifteenth27Idx.Length; slot++)
        {
            object?[] fifteenth27Args = [fifteenth27Idx[slot], null];
            TestAssert.Equal(true, (bool)fifteenth27!.Invoke(null, fifteenth27Args)!);
            TestAssert.Equal((byte)fifteenth27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fifteenth27Args[1])!);
            TestAssert.Equal((byte)fifteenth27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fifteenth27Args[1])!);
            TestAssert.Equal((byte)fifteenth27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fifteenth27Args[1])!);
        }
        object?[] fifteenth27Past = [89, null];
        TestAssert.Equal(false, (bool)fifteenth27!.Invoke(null, fifteenth27Past)!);
        var twelfth27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeTwelfthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfth27 is not null, "Expected twenty-seventh-regime twelfth table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Twelfth (slots 109, 111-114, 135, 149-150, 157-160), dash162/dash163 agree bit-identical replays within 2.
        int[] twelfth27Idx = [108, 110, 111, 112, 113, 134, 148, 149, 156, 157, 158, 159];
        int[] twelfth27R = [140, 181, 164, 139, 249, 203, 193, 251, 208, 232, 222, 216];
        int[] twelfth27G = [165, 203, 149, 192, 178, 218, 219, 211, 216, 208, 231, 211];
        int[] twelfth27B = [204, 144, 186, 211, 137, 181, 229, 193, 232, 208, 209, 224];
        for (int slot = 0; slot < twelfth27Idx.Length; slot++)
        {
            object?[] twelfth27Args = [twelfth27Idx[slot], null];
            TestAssert.Equal(true, (bool)twelfth27!.Invoke(null, twelfth27Args)!);
            TestAssert.Equal((byte)twelfth27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfth27Args[1])!);
            TestAssert.Equal((byte)twelfth27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfth27Args[1])!);
            TestAssert.Equal((byte)twelfth27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfth27Args[1])!);
        }
        object?[] twelfth27Past = [101, null];
        TestAssert.Equal(false, (bool)twelfth27!.Invoke(null, twelfth27Past)!);
        var sixthSingle = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeSixthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixthSingle is not null, "Expected twenty-seventh-regime sixthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime SixthSingle (slot 132), dash162/dash163 agree bit-identical replays within 2.
        int[] sixthSingleIdx = [131];
        int[] sixthSingleR = [250];
        int[] sixthSingleG = [195];
        int[] sixthSingleB = [168];
        for (int slot = 0; slot < sixthSingleIdx.Length; slot++)
        {
            object?[] sixthSingleArgs = [sixthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)sixthSingle!.Invoke(null, sixthSingleArgs)!);
            TestAssert.Equal((byte)sixthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixthSingleArgs[1])!);
            TestAssert.Equal((byte)sixthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixthSingleArgs[1])!);
            TestAssert.Equal((byte)sixthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixthSingleArgs[1])!);
        }
        object?[] sixthSinglePast = [130, null];
        TestAssert.Equal(false, (bool)sixthSingle!.Invoke(null, sixthSinglePast)!);
        var ninth27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeNinthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninth27 is not null, "Expected twenty-seventh-regime ninth table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Ninth (slots 121-126), dash162/dash163 agree bit-identical.
        int[] ninth27Idx = [120, 121, 122, 123, 124, 125];
        int[] ninth27R = [163, 214, 193, 180, 161, 250];
        int[] ninth27G = [181, 163, 211, 169, 203, 191];
        int[] ninth27B = [212, 162, 166, 197, 218, 160];
        for (int slot = 0; slot < ninth27Idx.Length; slot++)
        {
            object?[] ninth27Args = [ninth27Idx[slot], null];
            TestAssert.Equal(true, (bool)ninth27!.Invoke(null, ninth27Args)!);
            TestAssert.Equal((byte)ninth27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninth27Args[1])!);
            TestAssert.Equal((byte)ninth27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninth27Args[1])!);
            TestAssert.Equal((byte)ninth27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninth27Args[1])!);
        }
        object?[] ninth27Past = [119, null];
        TestAssert.Equal(false, (bool)ninth27!.Invoke(null, ninth27Past)!);
        var overflow27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflow27 is not null, "Expected twenty-seventh-regime overflow table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Overflow (slots 128-130, 133-134, 136, 145-148, 151-154), dash162/dash163 agree bit-identical replays within 2.
        int[] overflow27Idx = [127, 128, 129, 132, 133, 135, 144, 145, 146, 147, 150, 151, 152, 153];
        int[] overflow27R = [217, 198, 186, 182, 221, 195, 194, 226, 213, 205, 200, 229, 218, 209];
        int[] overflow27G = [170, 214, 176, 195, 182, 186, 205, 194, 224, 198, 209, 200, 228, 203];
        int[] overflow27B = [169, 172, 201, 220, 181, 208, 225, 194, 196, 215, 228, 199, 204, 219];
        for (int slot = 0; slot < overflow27Idx.Length; slot++)
        {
            object?[] overflow27Args = [overflow27Idx[slot], null];
            TestAssert.Equal(true, (bool)overflow27!.Invoke(null, overflow27Args)!);
            TestAssert.Equal((byte)overflow27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflow27Args[1])!);
            TestAssert.Equal((byte)overflow27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflow27Args[1])!);
            TestAssert.Equal((byte)overflow27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflow27Args[1])!);
        }
        object?[] overflow27Past = [125, null];
        TestAssert.Equal(false, (bool)overflow27!.Invoke(null, overflow27Past)!);
        var eighth27 = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeEighthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighth27 is not null, "Expected twenty-seventh-regime eighth table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Eighth (slots 137-138), dash162/dash163 agree bit-identical replays within 2.
        int[] eighth27Idx = [136, 137];
        int[] eighth27R = [181, 250];
        int[] eighth27G = [212, 203];
        int[] eighth27B = [224, 180];
        for (int slot = 0; slot < eighth27Idx.Length; slot++)
        {
            object?[] eighth27Args = [eighth27Idx[slot], null];
            TestAssert.Equal(true, (bool)eighth27!.Invoke(null, eighth27Args)!);
            TestAssert.Equal((byte)eighth27R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighth27Args[1])!);
            TestAssert.Equal((byte)eighth27G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighth27Args[1])!);
            TestAssert.Equal((byte)eighth27B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighth27Args[1])!);
        }
        object?[] eighth27Past = [135, null];
        TestAssert.Equal(false, (bool)eighth27!.Invoke(null, eighth27Past)!);
        var sixteenthSingle = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeSixteenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixteenthSingle is not null, "Expected twenty-seventh-regime sixteenthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime SixteenthSingle (slot 156), dash162/dash163 agree bit-identical replays within 1.
        int[] sixteenthSingleIdx = [155];
        int[] sixteenthSingleR = [251];
        int[] sixteenthSingleG = [215];
        int[] sixteenthSingleB = [199];
        for (int slot = 0; slot < sixteenthSingleIdx.Length; slot++)
        {
            object?[] sixteenthSingleArgs = [sixteenthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)sixteenthSingle!.Invoke(null, sixteenthSingleArgs)!);
            TestAssert.Equal((byte)sixteenthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixteenthSingleArgs[1])!);
            TestAssert.Equal((byte)sixteenthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixteenthSingleArgs[1])!);
            TestAssert.Equal((byte)sixteenthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixteenthSingleArgs[1])!);
        }
        object?[] sixteenthSinglePast = [147, null];
        TestAssert.Equal(false, (bool)sixteenthSingle!.Invoke(null, sixteenthSinglePast)!);
        var twentySixthSingle = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeTwentySixthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentySixthSingle is not null, "Expected twenty-seventh-regime twentysixthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime TwentySixthSingle (slot 161), dash162/dash163 agree bit-identical.
        int[] twentySixthSingleIdx = [160];
        int[] twentySixthSingleR = [208];
        int[] twentySixthSingleG = [227];
        int[] twentySixthSingleB = [234];
        for (int slot = 0; slot < twentySixthSingleIdx.Length; slot++)
        {
            object?[] twentySixthSingleArgs = [twentySixthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)twentySixthSingle!.Invoke(null, twentySixthSingleArgs)!);
            TestAssert.Equal((byte)twentySixthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentySixthSingleArgs[1])!);
            TestAssert.Equal((byte)twentySixthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentySixthSingleArgs[1])!);
            TestAssert.Equal((byte)twentySixthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentySixthSingleArgs[1])!);
        }
        object?[] twentySixthSinglePast = [159, null];
        TestAssert.Equal(false, (bool)twentySixthSingle!.Invoke(null, twentySixthSinglePast)!);
        var earlySingle = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeEarlySingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(earlySingle is not null, "Expected twenty-seventh-regime earlysingle table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime EarlySingle (slot 6), dash162/dash163 agree bit-identical.
        int[] earlySingleIdx = [5];
        int[] earlySingleR = [152];
        int[] earlySingleG = [91];
        int[] earlySingleB = [40];
        for (int slot = 0; slot < earlySingleIdx.Length; slot++)
        {
            object?[] earlySingleArgs = [earlySingleIdx[slot], null];
            TestAssert.Equal(true, (bool)earlySingle!.Invoke(null, earlySingleArgs)!);
            TestAssert.Equal((byte)earlySingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(earlySingleArgs[1])!);
            TestAssert.Equal((byte)earlySingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(earlySingleArgs[1])!);
            TestAssert.Equal((byte)earlySingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(earlySingleArgs[1])!);
        }
        object?[] earlySinglePast = [4, null];
        TestAssert.Equal(false, (bool)earlySingle!.Invoke(null, earlySinglePast)!);
        var lateSingle = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeLateSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(lateSingle is not null, "Expected twenty-seventh-regime latesingle table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime LateSingles (slot 163 dash163/dash164 agree, slot 164 dash164/dash165 agree, slot 165 dash165/dash166 agree).
        int[] lateSingleIdx = [162, 163, 164];
        int[] lateSingleR = [215, 235, 227];
        int[] lateSingleG = [222, 215, 234];
        int[] lateSingleB = [235, 214, 216];
        for (int slot = 0; slot < lateSingleIdx.Length; slot++)
        {
            object?[] lateSingleArgs = [lateSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)lateSingle!.Invoke(null, lateSingleArgs)!);
            TestAssert.Equal((byte)lateSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(lateSingleArgs[1])!);
            TestAssert.Equal((byte)lateSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(lateSingleArgs[1])!);
            TestAssert.Equal((byte)lateSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(lateSingleArgs[1])!);
        }
        object?[] lateSinglePast = [161, null];
        TestAssert.Equal(false, (bool)lateSingle!.Invoke(null, lateSinglePast)!);
        var tail = rendererType.GetMethod(
            "TryResolveTwentySeventhRegimeTailFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tail is not null, "Expected twenty-seventh-regime tail table to remain inspectable by the Office evidence guard.");
        // Twenty-seventh-regime Tail (slot 167), dash166/dash167 agree bit-identical.
        int[] tailIdx = [165];
        int[] tailR = [220];
        int[] tailG = [216];
        int[] tailB = [227];
        for (int slot = 0; slot < tailIdx.Length; slot++)
        {
            object?[] tailArgs = [tailIdx[slot], null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] tailPast = [164, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, tailPast)!);
        // One-hundred-sixty-one points and fewer keep earlier regimes; one-hundred-sixty-two-plus take the twenty-seventh rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [161])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [162])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [163])!);
    }
    public static void PptxSyntheticTwentyEighthVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseTwentyEighthVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected twenty-eighth-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var seventeenth28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeSeventeenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(seventeenth28 is not null, "Expected twenty-eighth-regime seventeenth table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Seventeenth (slots 66, 84), dash168/dash169 agree bit-identical replays within 2.
        int[] seventeenth28Idx = [65, 83];
        int[] seventeenth28R = [225, 243];
        int[] seventeenth28G = [136, 147];
        int[] seventeenth28B = [63, 69];
        for (int slot = 0; slot < seventeenth28Idx.Length; slot++)
        {
            object?[] seventeenth28Args = [seventeenth28Idx[slot], null];
            TestAssert.Equal(true, (bool)seventeenth28!.Invoke(null, seventeenth28Args)!);
            TestAssert.Equal((byte)seventeenth28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(seventeenth28Args[1])!);
            TestAssert.Equal((byte)seventeenth28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(seventeenth28Args[1])!);
            TestAssert.Equal((byte)seventeenth28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(seventeenth28Args[1])!);
        }
        object?[] seventeenth28Past = [64, null];
        TestAssert.Equal(false, (bool)seventeenth28!.Invoke(null, seventeenth28Past)!);
        var eighteenth28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeEighteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighteenth28 is not null, "Expected twenty-eighth-regime eighteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Eighteenth (slots 85-90, 161), dash168/dash169 agree bit-identical.
        int[] eighteenth28Idx = [84, 85, 86, 87, 88, 89, 160];
        int[] eighteenth28R = [79, 192, 155, 128, 75, 247, 202];
        int[] eighteenth28G = [129, 80, 187, 100, 172, 150, 223];
        int[] eighteenth28B = [189, 77, 89, 162, 198, 70, 232];
        for (int slot = 0; slot < eighteenth28Idx.Length; slot++)
        {
            object?[] eighteenth28Args = [eighteenth28Idx[slot], null];
            TestAssert.Equal(true, (bool)eighteenth28!.Invoke(null, eighteenth28Args)!);
            TestAssert.Equal((byte)eighteenth28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighteenth28Args[1])!);
            TestAssert.Equal((byte)eighteenth28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighteenth28Args[1])!);
            TestAssert.Equal((byte)eighteenth28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighteenth28Args[1])!);
        }
        object?[] eighteenth28Past = [82, null];
        TestAssert.Equal(false, (bool)eighteenth28!.Invoke(null, eighteenth28Past)!);
        var thirteenth28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeThirteenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(thirteenth28 is not null, "Expected twenty-eighth-regime thirteenth table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Thirteenth (slots 91-96, 127-132, 151-156), dash168/dash169 agree bit-identical.
        int[] thirteenth28Idx = [90, 91, 92, 93, 94, 95, 126, 127, 128, 129, 130, 131, 150, 151, 152, 153, 154, 155];
        int[] thirteenth28R = [95, 195, 161, 136, 92, 247, 164, 215, 195, 182, 163, 250, 195, 227, 214, 206, 195, 251];
        int[] thirteenth28G = [137, 96, 190, 112, 176, 156, 182, 165, 212, 171, 204, 192, 206, 196, 225, 199, 220, 212];
        int[] thirteenth28B = [192, 94, 103, 167, 201, 89, 213, 164, 167, 198, 218, 162, 226, 195, 197, 216, 229, 194];
        for (int slot = 0; slot < thirteenth28Idx.Length; slot++)
        {
            object?[] thirteenth28Args = [thirteenth28Idx[slot], null];
            TestAssert.Equal(true, (bool)thirteenth28!.Invoke(null, thirteenth28Args)!);
            TestAssert.Equal((byte)thirteenth28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(thirteenth28Args[1])!);
            TestAssert.Equal((byte)thirteenth28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(thirteenth28Args[1])!);
            TestAssert.Equal((byte)thirteenth28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(thirteenth28Args[1])!);
        }
        object?[] thirteenth28Past = [89, null];
        TestAssert.Equal(false, (bool)thirteenth28!.Invoke(null, thirteenth28Past)!);
        var twentyFirst28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeTwentyFirstFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentyFirst28 is not null, "Expected twenty-eighth-regime twentyfirst table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime TwentyFirst (slots 97-99, 101-102), dash168/dash169 agree bit-identical.
        int[] twentyFirst28Idx = [96, 97, 98, 100, 101];
        int[] twentyFirst28R = [112, 199, 167, 109, 248];
        int[] twentyFirst28G = [147, 112, 194, 182, 163];
        int[] twentyFirst28B = [196, 111, 118, 204, 107];
        for (int slot = 0; slot < twentyFirst28Idx.Length; slot++)
        {
            object?[] twentyFirst28Args = [twentyFirst28Idx[slot], null];
            TestAssert.Equal(true, (bool)twentyFirst28!.Invoke(null, twentyFirst28Args)!);
            TestAssert.Equal((byte)twentyFirst28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentyFirst28Args[1])!);
            TestAssert.Equal((byte)twentyFirst28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentyFirst28Args[1])!);
            TestAssert.Equal((byte)twentyFirst28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentyFirst28Args[1])!);
        }
        object?[] twentyFirst28Past = [95, null];
        TestAssert.Equal(false, (bool)twentyFirst28!.Invoke(null, twentyFirst28Past)!);
        var fourteenthSingle = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeFourteenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourteenthSingle is not null, "Expected twenty-eighth-regime fourteenthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime FourteenthSingle (slot 100), dash168/dash169 agree bit-identical replays within 2.
        int[] fourteenthSingleIdx = [99];
        int[] fourteenthSingleR = [144];
        int[] fourteenthSingleG = [123];
        int[] fourteenthSingleB = [172];
        for (int slot = 0; slot < fourteenthSingleIdx.Length; slot++)
        {
            object?[] fourteenthSingleArgs = [fourteenthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)fourteenthSingle!.Invoke(null, fourteenthSingleArgs)!);
            TestAssert.Equal((byte)fourteenthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourteenthSingleArgs[1])!);
            TestAssert.Equal((byte)fourteenthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourteenthSingleArgs[1])!);
            TestAssert.Equal((byte)fourteenthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourteenthSingleArgs[1])!);
        }
        object?[] fourteenthSinglePast = [98, null];
        TestAssert.Equal(false, (bool)fourteenthSingle!.Invoke(null, fourteenthSinglePast)!);
        var third28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeThirdFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(third28 is not null, "Expected twenty-eighth-regime third table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Third (slots 103-108), dash168/dash169 agree bit-identical.
        int[] third28Idx = [102, 103, 104, 105, 106, 107];
        int[] third28R = [126, 202, 174, 155, 124, 248];
        int[] third28G = [155, 126, 198, 137, 187, 170];
        int[] third28B = [200, 125, 131, 179, 207, 121];
        for (int slot = 0; slot < third28Idx.Length; slot++)
        {
            object?[] third28Args = [third28Idx[slot], null];
            TestAssert.Equal(true, (bool)third28!.Invoke(null, third28Args)!);
            TestAssert.Equal((byte)third28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(third28Args[1])!);
            TestAssert.Equal((byte)third28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(third28Args[1])!);
            TestAssert.Equal((byte)third28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(third28Args[1])!);
        }
        object?[] third28Past = [101, null];
        TestAssert.Equal(false, (bool)third28!.Invoke(null, third28Past)!);
        var sixth28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeSixthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixth28 is not null, "Expected twenty-eighth-regime sixth table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Sixth (slots 109, 111-112), dash168/dash169 agree bit-identical replays within 2.
        int[] sixth28Idx = [108, 110, 111];
        int[] sixth28R = [133, 177, 160];
        int[] sixth28G = [160, 201, 143];
        int[] sixth28B = [202, 138, 182];
        for (int slot = 0; slot < sixth28Idx.Length; slot++)
        {
            object?[] sixth28Args = [sixth28Idx[slot], null];
            TestAssert.Equal(true, (bool)sixth28!.Invoke(null, sixth28Args)!);
            TestAssert.Equal((byte)sixth28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixth28Args[1])!);
            TestAssert.Equal((byte)sixth28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixth28Args[1])!);
            TestAssert.Equal((byte)sixth28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixth28Args[1])!);
        }
        object?[] sixth28Past = [107, null];
        TestAssert.Equal(false, (bool)sixth28!.Invoke(null, sixth28Past)!);
        var ninth28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeNinthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(ninth28 is not null, "Expected twenty-eighth-regime ninth table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Ninth (slots 110, 113-114), dash168/dash169 agree bit-identical replays within 2.
        int[] ninth28Idx = [109, 112, 113];
        int[] ninth28R = [206, 136, 249];
        int[] ninth28G = [138, 192, 177];
        int[] ninth28B = [137, 210, 134];
        for (int slot = 0; slot < ninth28Idx.Length; slot++)
        {
            object?[] ninth28Args = [ninth28Idx[slot], null];
            TestAssert.Equal(true, (bool)ninth28!.Invoke(null, ninth28Args)!);
            TestAssert.Equal((byte)ninth28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(ninth28Args[1])!);
            TestAssert.Equal((byte)ninth28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(ninth28Args[1])!);
            TestAssert.Equal((byte)ninth28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(ninth28Args[1])!);
        }
        object?[] ninth28Past = [108, null];
        TestAssert.Equal(false, (bool)ninth28!.Invoke(null, ninth28Past)!);
        var overflow28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeOverflowFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(overflow28 is not null, "Expected twenty-eighth-regime overflow table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Overflow (slots 115-118, 139-142, 145-148, 157-160), dash168/dash169 agree bit-identical.
        int[] overflow28Idx = [114, 115, 116, 117, 138, 139, 140, 141, 144, 145, 146, 147, 156, 157, 158, 159];
        int[] overflow28R = [147, 209, 185, 169, 182, 221, 205, 195, 188, 224, 209, 200, 203, 230, 218, 212];
        int[] overflow28G = [169, 147, 205, 155, 195, 182, 219, 186, 200, 188, 222, 192, 212, 203, 228, 206];
        int[] overflow28B = [207, 146, 150, 189, 220, 181, 184, 208, 223, 188, 190, 212, 229, 202, 204, 220];
        for (int slot = 0; slot < overflow28Idx.Length; slot++)
        {
            object?[] overflow28Args = [overflow28Idx[slot], null];
            TestAssert.Equal(true, (bool)overflow28!.Invoke(null, overflow28Args)!);
            TestAssert.Equal((byte)overflow28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(overflow28Args[1])!);
            TestAssert.Equal((byte)overflow28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(overflow28Args[1])!);
            TestAssert.Equal((byte)overflow28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(overflow28Args[1])!);
        }
        object?[] overflow28Past = [113, null];
        TestAssert.Equal(false, (bool)overflow28!.Invoke(null, overflow28Past)!);
        var fourth28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeFourthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(fourth28 is not null, "Expected twenty-eighth-regime fourth table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Fourth (slots 119-120), dash168/dash169 agree bit-identical.
        int[] fourth28Idx = [118, 119];
        int[] fourth28R = [145, 249];
        int[] fourth28G = [195, 181];
        int[] fourth28B = [213, 144];
        for (int slot = 0; slot < fourth28Idx.Length; slot++)
        {
            object?[] fourth28Args = [fourth28Idx[slot], null];
            TestAssert.Equal(true, (bool)fourth28!.Invoke(null, fourth28Args)!);
            TestAssert.Equal((byte)fourth28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(fourth28Args[1])!);
            TestAssert.Equal((byte)fourth28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(fourth28Args[1])!);
            TestAssert.Equal((byte)fourth28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(fourth28Args[1])!);
        }
        object?[] fourth28Past = [117, null];
        TestAssert.Equal(false, (bool)fourth28!.Invoke(null, fourth28Past)!);
        var eighth28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeEighthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eighth28 is not null, "Expected twenty-eighth-regime eighth table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Eighth (slots 121-126, 143-144), dash168/dash169 agree bit-identical.
        int[] eighth28Idx = [120, 121, 122, 123, 124, 125, 142, 143];
        int[] eighth28R = [157, 212, 190, 176, 156, 249, 181, 250];
        int[] eighth28G = [177, 157, 209, 164, 200, 187, 212, 203];
        int[] eighth28B = [210, 156, 160, 194, 216, 154, 224, 180];
        for (int slot = 0; slot < eighth28Idx.Length; slot++)
        {
            object?[] eighth28Args = [eighth28Idx[slot], null];
            TestAssert.Equal(true, (bool)eighth28!.Invoke(null, eighth28Args)!);
            TestAssert.Equal((byte)eighth28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eighth28Args[1])!);
            TestAssert.Equal((byte)eighth28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eighth28Args[1])!);
            TestAssert.Equal((byte)eighth28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eighth28Args[1])!);
        }
        object?[] eighth28Past = [119, null];
        TestAssert.Equal(false, (bool)eighth28!.Invoke(null, eighth28Past)!);
        var eleventh28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeEleventhFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(eleventh28 is not null, "Expected twenty-eighth-regime eleventh table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Eleventh (slots 133-138), dash168/dash169 agree bit-identical.
        int[] eleventh28Idx = [132, 133, 134, 135, 136, 137];
        int[] eleventh28R = [173, 218, 200, 189, 172, 250];
        int[] eleventh28G = [189, 173, 215, 179, 208, 198];
        int[] eleventh28B = [217, 173, 176, 203, 221, 171];
        for (int slot = 0; slot < eleventh28Idx.Length; slot++)
        {
            object?[] eleventh28Args = [eleventh28Idx[slot], null];
            TestAssert.Equal(true, (bool)eleventh28!.Invoke(null, eleventh28Args)!);
            TestAssert.Equal((byte)eleventh28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(eleventh28Args[1])!);
            TestAssert.Equal((byte)eleventh28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(eleventh28Args[1])!);
            TestAssert.Equal((byte)eleventh28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(eleventh28Args[1])!);
        }
        object?[] eleventh28Past = [125, null];
        TestAssert.Equal(false, (bool)eleventh28!.Invoke(null, eleventh28Past)!);
        var twelfth28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeTwelfthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twelfth28 is not null, "Expected twenty-eighth-regime twelfth table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Twelfth (slots 163-166), dash168/dash169 agree bit-identical.
        int[] twelfth28Idx = [162, 163, 164, 165];
        int[] twelfth28R = [208, 232, 222, 216];
        int[] twelfth28G = [216, 208, 231, 211];
        int[] twelfth28B = [232, 208, 209, 224];
        for (int slot = 0; slot < twelfth28Idx.Length; slot++)
        {
            object?[] twelfth28Args = [twelfth28Idx[slot], null];
            TestAssert.Equal(true, (bool)twelfth28!.Invoke(null, twelfth28Args)!);
            TestAssert.Equal((byte)twelfth28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twelfth28Args[1])!);
            TestAssert.Equal((byte)twelfth28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twelfth28Args[1])!);
            TestAssert.Equal((byte)twelfth28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twelfth28Args[1])!);
        }
        object?[] twelfth28Past = [131, null];
        TestAssert.Equal(false, (bool)twelfth28!.Invoke(null, twelfth28Past)!);
        var sixteenthSingle = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeSixteenthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(sixteenthSingle is not null, "Expected twenty-eighth-regime sixteenthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime SixteenthSingle (slot 162), dash168/dash169 agree bit-identical replays within 2.
        int[] sixteenthSingleIdx = [161];
        int[] sixteenthSingleR = [251];
        int[] sixteenthSingleG = [215];
        int[] sixteenthSingleB = [199];
        for (int slot = 0; slot < sixteenthSingleIdx.Length; slot++)
        {
            object?[] sixteenthSingleArgs = [sixteenthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)sixteenthSingle!.Invoke(null, sixteenthSingleArgs)!);
            TestAssert.Equal((byte)sixteenthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(sixteenthSingleArgs[1])!);
            TestAssert.Equal((byte)sixteenthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(sixteenthSingleArgs[1])!);
            TestAssert.Equal((byte)sixteenthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(sixteenthSingleArgs[1])!);
        }
        object?[] sixteenthSinglePast = [147, null];
        TestAssert.Equal(false, (bool)sixteenthSingle!.Invoke(null, sixteenthSinglePast)!);
        var tenth28 = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeTenthFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(tenth28 is not null, "Expected twenty-eighth-regime tenth table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime Tenth (slots 149-150), dash168/dash169 agree bit-identical.
        int[] tenth28Idx = [148, 149];
        int[] tenth28R = [187, 251];
        int[] tenth28G = [215, 207];
        int[] tenth28B = [227, 186];
        for (int slot = 0; slot < tenth28Idx.Length; slot++)
        {
            object?[] tenth28Args = [tenth28Idx[slot], null];
            TestAssert.Equal(true, (bool)tenth28!.Invoke(null, tenth28Args)!);
            TestAssert.Equal((byte)tenth28R[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tenth28Args[1])!);
            TestAssert.Equal((byte)tenth28G[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tenth28Args[1])!);
            TestAssert.Equal((byte)tenth28B[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tenth28Args[1])!);
        }
        object?[] tenth28Past = [147, null];
        TestAssert.Equal(false, (bool)tenth28!.Invoke(null, tenth28Past)!);
        var twentySixthSingle = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeTwentySixthSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(twentySixthSingle is not null, "Expected twenty-eighth-regime twentysixthsingle table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime TwentySixthSingle (slot 167), dash168/dash169 agree bit-identical.
        int[] twentySixthSingleIdx = [166];
        int[] twentySixthSingleR = [208];
        int[] twentySixthSingleG = [227];
        int[] twentySixthSingleB = [234];
        for (int slot = 0; slot < twentySixthSingleIdx.Length; slot++)
        {
            object?[] twentySixthSingleArgs = [twentySixthSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)twentySixthSingle!.Invoke(null, twentySixthSingleArgs)!);
            TestAssert.Equal((byte)twentySixthSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(twentySixthSingleArgs[1])!);
            TestAssert.Equal((byte)twentySixthSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(twentySixthSingleArgs[1])!);
            TestAssert.Equal((byte)twentySixthSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(twentySixthSingleArgs[1])!);
        }
        object?[] twentySixthSinglePast = [165, null];
        TestAssert.Equal(false, (bool)twentySixthSingle!.Invoke(null, twentySixthSinglePast)!);
        var lateSingle = rendererType.GetMethod(
            "TryResolveTwentyEighthRegimeLateSingleFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(lateSingle is not null, "Expected twenty-eighth-regime latesingle table to remain inspectable by the Office evidence guard.");
        // Twenty-eighth-regime LateSingles (slot 169 dash169/dash170 agree, slot 170 dash170/dash171 agree).
        int[] lateSingleIdx = [168, 169];
        int[] lateSingleR = [215, 235];
        int[] lateSingleG = [222, 215];
        int[] lateSingleB = [235, 214];
        for (int slot = 0; slot < lateSingleIdx.Length; slot++)
        {
            object?[] lateSingleArgs = [lateSingleIdx[slot], null];
            TestAssert.Equal(true, (bool)lateSingle!.Invoke(null, lateSingleArgs)!);
            TestAssert.Equal((byte)lateSingleR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(lateSingleArgs[1])!);
            TestAssert.Equal((byte)lateSingleG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(lateSingleArgs[1])!);
            TestAssert.Equal((byte)lateSingleB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(lateSingleArgs[1])!);
        }
        object?[] lateSinglePast = [167, null];
        TestAssert.Equal(false, (bool)lateSingle!.Invoke(null, lateSinglePast)!);
        // One-hundred-sixty-seven points and fewer keep earlier regimes; one-hundred-sixty-eight-plus take the twenty-eighth rows.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [167])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [168])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [169])!);
    }
    public static void PptxSyntheticSecondVaryColorsRegime()
    {
        var rendererType = typeof(PptxRenderer);
        var gate = rendererType.GetMethod(
            "UseSecondVaryColorsRegime",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gate is not null, "Expected second-regime gate to remain inspectable by the Office evidence guard.");
        var rgbType = rendererType.Assembly.GetType("Lokad.OoxPdf.Pptx.RgbColor");
        TestAssert.True(rgbType is not null, "Expected RgbColor to remain resolvable for the regime pin.");
        var shade = rendererType.GetMethod(
            "ShadeSecondRegimeSingleSeriesVaryColorsFill",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(shade is not null, "Expected second-regime shade helper to remain inspectable by the Office evidence guard.");

        // Eleven points and fewer keep the first regime; twelve-plus take dark-plus-raw.
        TestAssert.Equal(false, (bool)gate!.Invoke(null, [11])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [12])!);
        TestAssert.Equal(true, (bool)gate.Invoke(null, [13])!);
        // Office dash12 slots 1-6 shade at 0.82 (maxabs 1 over 18 channels): accent1
        // (79,129,189) renders (65,106,155).
        object? accent = System.Activator.CreateInstance(rgbType!, (byte)79, (byte)129, (byte)189);
        object? shaded = shade!.Invoke(null, [accent]);
        TestAssert.Equal((byte)65, (byte)rgbType!.GetProperty("Red")!.GetValue(shaded)!);
        TestAssert.Equal((byte)106, (byte)rgbType.GetProperty("Green")!.GetValue(shaded)!);
        TestAssert.Equal((byte)155, (byte)rgbType.GetProperty("Blue")!.GetValue(shaded)!);
    }

    public static void PptxSyntheticNoTitleBottomLegendAnchoredTop()
    {
        var rendererType = typeof(PptxRenderer);
        var method = rendererType.GetMethod(
            "ResolveNoTitleBottomLegendAnchoredTop",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected bottom-legend anchored top to remain inspectable by the Office evidence guard.");
        var boxType = rendererType.GetNestedType("ChartFrameBox", System.Reflection.BindingFlags.NonPublic);
        TestAssert.True(boxType is not null, "Expected chart frame box to remain resolvable for the anchor pin.");

        // Office composite clips agree bit-identically base+tall (plot top 419.76
        // under plus-100H downward growth); the 0.949H ratio top falls 5.1 per 100H.
        object? baseFrame = System.Activator.CreateInstance(boxType!, 72d, 72d, 576d, 360d);
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [baseFrame])! - 419.76d) < 0.01d, "Expected base anchored top.");
        object? tallFrame = System.Activator.CreateInstance(boxType!, 72d, -28d, 576d, 460d);
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [tallFrame])! - 419.76d) < 0.01d, "Expected tall anchored top.");
    }

    public static void PptxSyntheticPerSeriesHorizontalLegendInterEntryGap()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputePerSeriesHorizontalLegendInterEntryGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected per-series inter law to remain inspectable by the Office evidence guard.");

        // Office per-series knots (fs, mean advance, inter): composite base plus
        // leg18 plus longnames plus shortnames plus NSW plus rangednames fit 0.45fs
        // plus 0.148avgW plus 0.195 within 0.062 with three spare DOF (wide ranges to 42 absorbed).
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [12d, 40.99d])! - 11.66d) < 0.1d, "Expected composite-base inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, 61.57d])! - 17.39d) < 0.1d, "Expected leg18 inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [12d, 85.40d])! - 18.24d) < 0.1d, "Expected longnames inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, 37.77d])! - 13.85d) < 0.1d, "Expected shortnames inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, 43.47d])! - 14.79d) < 0.1d, "Expected nsw inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [12d, 40.40d])! - 11.55d) < 0.1d, "Expected rangednames inter (wide range absorbed).");
    }

    public static void PptxSyntheticBottomLegendFrameAnchorX()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ResolveBottomLegendFrameAnchorX",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected bottom-legend frame anchor to remain inspectable by the Office evidence guard.");

        // Office bottom-legend blocks sit at frame-center plus markerSize over 4 plus
        // 0.65 (twelve renders within 0.05); the val14 splitter proves frame anchoring
        // with the plot moved plus-6 and the block unmoved.
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [72d, 576d, 174.82d, 6.57d])! - 274.87d) < 0.1d, "Expected composite anchored X.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [120d, 360d, 128.05d, 6.59d])! - 238.23d) < 0.1d, "Expected botleg anchored X.");
    }

    public static void PptxSyntheticMajorTickSegmentsScaleWithLabelSize()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "StrokeMajorTickSegments",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected major-tick segment helper to remain inspectable by the Office evidence guard.");

        // Office tick ink runs about 0.315 label sizes (5.71 at 18pt, 2.82 at 9pt).
        var graphics = new PdfGraphicsBuilder();
        method!.Invoke(null, [graphics, new double[] { 100d, 200d }, 50d, true, true, PptxSceneChartAxisTickMark.Outside, 18d * 0.315d]);
        string content = graphics.ToString();
        System.Text.RegularExpressions.MatchCollection segments = System.Text.RegularExpressions.Regex.Matches(content, @"([0-9.]+) ([0-9.]+) m ([0-9.]+) ([0-9.]+) l");
        TestAssert.Equal(2, segments.Count);
        foreach (System.Text.RegularExpressions.Match segment in segments)
        {
            double y1 = double.Parse(segment.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
            double y2 = double.Parse(segment.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture);
            TestAssert.True(System.Math.Abs((y2 - y1) - 18d * 0.315d) < 0.01d, "Expected tick length " + (18d * 0.315d).ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
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

    public static void PptxSyntheticChartAreaDefaultBorderKeepsOfficeCalibration()
    {
        var rules = typeof(PptxRenderer).GetNestedType("PptxChartMetricRules", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected metric rules.");
        double border = (double)rules.GetField("ChartAreaDefaultBorderWidth")!.GetValue(null)!;
        TestAssert.True(Math.Abs(border - 0.14d) < 0.000001d, "Chart-area default border should stay at the Office hairline (10 kind ports). Got " + border);
    }

    public static void PptxSyntheticBarLegendKeyUnitLeftCentersOnBar()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeyUnitLeft",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected bar legend-key unit helper to remain inspectable by the Office evidence guard.");
        double twoDigit = (double)method!.Invoke(null, [500.65d, 4.4d, 8.63d])!;
        double singleDigit = (double)method.Invoke(null, [200.25d, 4.4d, 4.2934d])!;
        TestAssert.True(Math.Abs(twoDigit - 493.385d) < 0.1d, "Two-digit swatches should start at the unit left (Office Delta 493.4). Got " + twoDigit);
        TestAssert.True(Math.Abs(singleDigit - 195.157d) < 0.1d, "Single-digit swatches should start at the unit left (Office Alpha 195.1). Got " + singleDigit);
    }

    public static void PptxSyntheticBarCategoryBottomReserveMeasuresOfficeStrip()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeBarLabelStripBottomReserve",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected bar category bottom helper to remain inspectable by the Office evidence guard.");
        double seven = (double)method!.Invoke(null, [6.96d])!;
        double fourteen = (double)method!.Invoke(null, [14.04d])!;
        double eighteen = (double)method!.Invoke(null, [18d])!;
        double nine = (double)method!.Invoke(null, [9d])!;
        TestAssert.True(Math.Abs(seven - 19.62d) < 0.15d, "Seven-point cats should reserve the measured bottom strip (Office 18.72). Got " + seven);
        TestAssert.True(Math.Abs(fourteen - 32.45d) < 0.15d, "Fourteen-point cats should reserve the measured bottom strip (Office 31.68). Got " + fourteen);
        TestAssert.True(Math.Abs(eighteen - 39.63d) < 0.15d, "Eighteen-point cats should reserve the measured bottom strip (Office 39.24). Got " + eighteen);
        TestAssert.True(Math.Abs(nine - 23.32d) < 0.15d, "Nine-point value labels should reserve the measured bottom strip (Office 23.2). Got " + nine);
    }

    public static void PptxSyntheticBarLegendKeyOutEndGapLiftsWithFontSize()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeyOutEndGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected bar out-end gap helper to remain inspectable by the Office evidence guard.");
        double eight = (double)method!.Invoke(null, [8d])!;
        double twelve = (double)method!.Invoke(null, [12d])!;
        double sixteen = (double)method!.Invoke(null, [16d])!;
        TestAssert.True(Math.Abs(eight - 6.74d) < 0.05d, "Eight-point labels should lift 6.74 above bar tops (Office 6.66). Got " + eight);
        TestAssert.True(Math.Abs(twelve - 7.86d) < 0.05d, "Twelve-point labels should lift 7.86 above bar tops (Office 7.83). Got " + twelve);
        TestAssert.True(Math.Abs(sixteen - 8.98d) < 0.05d, "Sixteen-point labels should lift 8.98 above bar tops (Office 8.99). Got " + sixteen);
    }

    public static void PptxSyntheticBarLegendKeySwatchAnchorsToText()
    {
        var gapMethod = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeySwatchGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        var xMethod = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeySwatchX",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        var yMethod = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeySwatchY",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gapMethod is not null && xMethod is not null && yMethod is not null, "Expected bar swatch helpers to remain inspectable by the Office evidence guard.");
        double gapEight = (double)gapMethod!.Invoke(null, [8d])!;
        double gapSixteen = (double)gapMethod!.Invoke(null, [16d])!;
        TestAssert.True(Math.Abs(gapEight - 4.5d) < 0.05d, "Eight-point swatch gap should read 4.50 (Office 4.47). Got " + gapEight);
        TestAssert.True(Math.Abs(gapSixteen - 6.63d) < 0.05d, "Sixteen-point swatch gap should read 6.63 (Office 6.64, kills G equals S). Got " + gapSixteen);
        double swatchXEight = (double)xMethod!.Invoke(null, [203.98d, 8.04d])!;
        double swatchYEight = (double)yMethod!.Invoke(null, [247.94d, 8.04d])!;
        TestAssert.True(Math.Abs(swatchXEight - 195.12d) < 0.15d, "Eight-point swatch should start at the Office left (195.12). Got " + swatchXEight);
        TestAssert.True(Math.Abs(swatchYEight - 248.54d) < 0.15d, "Eight-point swatch should sit at the Office top (248.54). Got " + swatchYEight);
        double swatchXSixteen = (double)xMethod!.Invoke(null, [334.97d, 15.96d])!;
        double swatchYSixteen = (double)yMethod!.Invoke(null, [271.25d, 15.96d])!;
        TestAssert.True(Math.Abs(swatchXSixteen - 319.54d) < 0.15d, "Sixteen-point swatch should start at the Office left (319.54). Got " + swatchXSixteen);
        TestAssert.True(Math.Abs(swatchYSixteen - 272.25d) < 0.15d, "Sixteen-point swatch should sit at the Office top (272.25). Got " + swatchYSixteen);
    }

    public static void PptxSyntheticChartAxisLabelsSkipPairKerning()
    {
        var measurerType = typeof(PptxRenderer).GetNestedType(
            "ChartTextMeasurer",
            System.Reflection.BindingFlags.NonPublic);
        TestAssert.True(measurerType is not null, "Expected chart text measurer to remain inspectable by the Office evidence guard.");
        object kerned = System.Activator.CreateInstance(measurerType!, [null, true])!;
        object plain = System.Activator.CreateInstance(measurerType!, [null, false])!;
        System.Reflection.MethodInfo? measure = null;
        foreach (var candidate in measurerType!.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
        {
            if (candidate.Name == "Measure" && candidate.GetParameters().Length == 6)
            {
                measure = candidate;
                break;
            }
        }
        double kernedWidth = (double)measure!.Invoke(kerned, ["Marketing", 18d, "Calibri", false, false, 0d])!;
        double plainWidth = (double)measure.Invoke(plain, ["Marketing", 18d, "Calibri", false, false, 0d])!;
        TestAssert.True(Math.Abs(plainWidth - 75.55d) < 0.5d, "Unkerned Marketing should match the Office advance (75.55). Got " + plainWidth);
        TestAssert.True(Math.Abs((plainWidth - kernedWidth) - 0.68d) < 0.15d, "Legacy pair kerning should tighten Marketing by 0.68 (ke plus et); axis labels must not carry it. Got " + (plainWidth - kernedWidth));
    }

    public static void PptxSyntheticPieLongWordSplitWidthReservesSeparator()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputePieLongWordSplitWidth",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie long-word split helper to remain inspectable by the Office evidence guard.");
        double split = (double)method!.Invoke(null, [98.8d, 3.77d])!;
        TestAssert.True(Math.Abs(split - 95.03d) < 0.01d, "Pre-split chunks should reserve the trailing separator (E2 breaks at 7 chars). Got " + split);
    }

    public static void PptxSyntheticPieLeaderFootYKeepsOfficeBaselines()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputePieLeaderFootY",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie leader-foot helper to remain inspectable by the Office evidence guard.");
        double single = (double)method!.Invoke(null, [239.28d, 24.97d, 18d, 1])!;
        double wrapped = (double)method.Invoke(null, [120d, 46.94d, 18d, 2])!;
        TestAssert.True(Math.Abs(single - 245.82d) < 0.1d, "Single-line feet should sit on the text baseline (Office Beta 245.78). Got " + single);
        TestAssert.True(Math.Abs(wrapped - 150.86d) < 0.1d, "Wrapped feet should sit 2.35pt above the first baseline (Office Gamma 150.79). Got " + wrapped);
    }

    public static void PptxSyntheticPieManualLeaderNeedsSameSideNarrowText()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ShouldDrawPieManualLeaderLabel",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie manual-leader pick helper to remain inspectable by the Office evidence guard.");
        object?[] gamma = [182.3d, 404d, -0.9767d, 100.3d, 97.67d, 520d, 360d];
        object?[] west = [157.66d, 294d, -1d, 77.3d, 79.8d, 420d, 300d];
        object?[] wide = [187.7d, 404d, -0.9767d, 119.07d, 97.67d, 520d, 360d];
        object?[] narrowWrapped = [182.73d, 404d, -0.9767d, 259.34d, 97.67d, 520d, 360d];
        object?[] squareNarrow = [172.11d, 324d, -0.9767d, 50.15d, 68.4d, 360d, 360d];
        object?[] alpha = [202.32d, 404d, 0.998d, 82.6d, 97.67d, 520d, 360d];
        object?[] axial = [478.04d, 294d, 0d, 83.8d, 79.8d, 420d, 300d];
        TestAssert.True((bool)method!.Invoke(null, gamma)!, "Same-side Gamma at 102.7% of cap should keep its leader (Office draws one).");
        TestAssert.True((bool)method.Invoke(null, west)!, "Same-side West at 96.9% of cap should keep its leader (Office draws one).");
        TestAssert.True(!(bool)method.Invoke(null, wide)!, "Same-side wide labels should stay leaderless (Office graded Gamma draws none).");
        TestAssert.True(!(bool)method.Invoke(null, narrowWrapped)!, "Narrow-line wrapped boxes over the text cap should stay leaderless (Office G2 draws none).");
        TestAssert.True(!(bool)method.Invoke(null, squareNarrow)!, "Square plots should stay leaderless (Office H2 Gamma draws none).");
        TestAssert.True(!(bool)method.Invoke(null, alpha)!, "Opposite-side Alpha should stay leaderless (Office draws none).");
        TestAssert.True(!(bool)method.Invoke(null, axial)!, "On-axis rims should stay leaderless (Office South draws none).");
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

    public static void PptxSyntheticPieManualAnchorRidesCirclePastRim()
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
        object right = method!.Invoke(null, [plotBox, Layout(-0.77024, -0.36990), 558.46d, 309.72d, 0.06279d, 0.99803d, 404d, 88.64d, 24.97d, 24.97d])!;
        object left = method.Invoke(null, [plotBox, Layout(0.66015, -0.73512), 320.98d, 169.40d, -0.84385d, -0.53644d, 404d, 79.43d, 24.97d, 24.97d])!;
        (double rx, double ry) = ReadLayoutBoxXY(right);
        (double lx, double ly) = ReadLayoutBoxXY(left);
        TestAssert.True(Math.Abs(rx - 157.94d) < 0.5d && Math.Abs(ry - 431.18d) < 0.5d, "Right-side anchor reads away from the circle (Office Alpha 158.0/431.0). Got " + rx + "/" + ry);
        TestAssert.True(Math.Abs(lx - 584.83d) < 0.5d && Math.Abs(ly - 411.02d) < 0.5d, "Left-side anchor reads toward the circle past the full width (Office Beta 584.6/411.0). Got " + lx + "/" + ly);
        object straight = method!.Invoke(null, [plotBox, Layout(0.63844, 0.87143), 294d, 174d, -1d, 0d, 294d, 40d, 24.3d, 24.3d])!;
        (double sx, double sy) = ReadLayoutBoxXY(straight);
        TestAssert.True(Math.Abs(sx - 606.0d) < 0.5d, "Straight up/down slices should center the box on the rim midpoint. Got " + sx + "/" + sy);
    }

    public static void PptxSyntheticPieLabelClipExpandsForOverflow()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var method = renderer.GetMethod("ExpandPieLabelClipToText", flags);
        TestAssert.True(method is not null, "Expected pie clip helper to remain inspectable by the Office evidence guard.");
        var boxType = renderer.GetNestedType("ChartLayoutBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected layout box.");
        object box = System.Activator.CreateInstance(boxType, [136d, 100d, 80d, 24d])!;
        object fitting = method!.Invoke(null, [box, 136d, 100d, 80d, 24d, 62d])!;
        (double fx, double fy) = ReadLayoutBoxXY(fitting);
        TestAssert.True(Math.Abs(fx - 136d) < 0.0001d, "Fitting text should keep the clip. Got " + fx);
        object wide = method.Invoke(null, [box, 136d, 100d, 80d, 24d, 100d])!;
        (double wx, double wy) = ReadLayoutBoxXY(wide);
        TestAssert.True(Math.Abs(wx - 126d) < 0.0001d, "Overflowing text should expand the clip symmetrically. Got " + wx);
    }

    public static void PptxSyntheticPieLabelWordWrapPacksWords()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "SplitPieLabelWordLines",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie word-wrap helper to remain inspectable by the Office evidence guard.");
        System.Collections.Generic.List<int[]> Gamma() => (System.Collections.Generic.List<int[]>)method!.Invoke(null, [new double[] { 62.6d, 34.1d }, new double[] { 5d, 0d }, 98.8d])!;
        System.Collections.Generic.List<int[]> West() => (System.Collections.Generic.List<int[]>)method!.Invoke(null, [new double[] { 40.1d, 34.1d }, new double[] { 5d, 0d }, 79.8d])!;
        System.Collections.Generic.List<int[]> North() => (System.Collections.Generic.List<int[]>)method!.Invoke(null, [new double[] { 44.4d, 34.1d }, new double[] { 5d, 0d }, 79.8d])!;
        TestAssert.True(Gamma().Count == 2 && Gamma()[0].Length == 1 && Gamma()[1].Length == 1, "Gamma total over the cap should break between words.");
        TestAssert.True(West().Count == 1 && West()[0].Length == 2, "West total inside the cap should stay single-line.");
        TestAssert.True(North().Count == 2, "North total over the cap should wrap despite near-equal single-line widths.");
    }

    public static void PptxSyntheticPieLabelLongWordSplitsGreedily()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "SplitPieLabelLongWord",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie long-word helper to remain inspectable by the Office evidence guard.");
        System.Collections.Generic.List<int> Tens() => (System.Collections.Generic.List<int>)method!.Invoke(null, [new double[] { 10d, 10d, 10d, 10d, 10d, 10d, 10d, 10d, 10d, 10d, 10d, 10d }, 95d])!;
        System.Collections.Generic.List<int> Exact() => (System.Collections.Generic.List<int>)method!.Invoke(null, [new double[] { 40d, 40d }, 80d])!;
        TestAssert.True(string.Join(",", Tens()) == "9,3", "Twelve 10pt chars at cap 95 should split 9 and 3.");
        TestAssert.True(string.Join(",", Exact()) == "2", "An exact-fit word should stay whole.");
    }

    private static (double X, double Y) ReadLayoutBoxXY(object box)
    {
        double x = (double)box.GetType().GetProperty("X")!.GetValue(box)!;
        double y = (double)box.GetType().GetProperty("Y")!.GetValue(box)!;
        return (x, y);
    }

    public static void PptxSyntheticPolarLabeledLayoutShrinksAndCenters()
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
        object unlabeled = method!.Invoke(null, [pie, plotBox, empty, hidden, false, false, 0d, 0d, false])!;
        object labeled = method.Invoke(null, [pie, plotBox, empty, hidden, true, false, 0d, 0d, false])!;
        (double ux, double uy, double ur) = ReadPolarGeometry(unlabeled);
        (double lx, double ly, double lr) = ReadPolarGeometry(labeled);
        TestAssert.True(Math.Abs(ux - 404d) < 0.01d && Math.Abs(uy - 284.88d) < 0.01d && Math.Abs(ur - 156.24d) < 0.01d, "Unlabeled pies keep the tall radius (5-categories port). Got " + ux + "/" + uy + "/" + ur);
        TestAssert.True(Math.Abs(lx - 404d) < 0.01d && Math.Abs(ly - 300d) < 0.01d && Math.Abs(lr - 146.5d) < 0.01d, "Labeled pies center with the tall radius (leader probes 146.50). Got " + lx + "/" + ly + "/" + lr);
        object shortPlotBox = System.Activator.CreateInstance(plotBoxType, [84d, 144d, 420d, 300d])!;
        object shortLabeled = method.Invoke(null, [pie, shortPlotBox, empty, hidden, true, false, 0d, 0d, false])!;
        (double qx, double qy, double qr) = ReadPolarGeometry(shortLabeled);
        TestAssert.True(Math.Abs(qx - 294d) < 0.01d && Math.Abs(qy - 294d) < 0.01d && Math.Abs(qr - 120.25d) < 0.01d, "Short labeled pies use the short radius (offset probe 120.25). Got " + qx + "/" + qy + "/" + qr);
        object doughnut = System.Enum.ToObject(kindType, 1);
        object narrowPlot = System.Activator.CreateInstance(plotBoxType, [144d, 72d, 246d, 432d])!;
        object narrowDoughnut = method.Invoke(null, [doughnut, narrowPlot, empty, hidden, false, false, 0d, 0d, false])!;
        (double nx, double ny, double nr) = ReadPolarGeometry(narrowDoughnut);
        TestAssert.True(Math.Abs(nx - 267d) < 0.01d && Math.Abs(ny - 288d) < 0.01d && Math.Abs(nr - 112.03d) < 0.01d, "Narrow doughnuts bind the width margin, not min-side (portrait probe). Got " + nx + "/" + ny + "/" + nr);
    }
    public static void PptxSyntheticExplodedDoughnutReserveFollowsRingGap()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeExplodedDoughnutRightLegendReserve",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected exploded doughnut reserve helper to remain inspectable by the Office evidence guard.");

        double narrow = (double)method!.Invoke(null, [685.03d, 598.53d])!;
        double mid = (double)method.Invoke(null, [642.11d, 598.53d])!;
        double wide = (double)method.Invoke(null, [575.55d, 598.53d])!;
        double shortPlot = (double)method.Invoke(null, [642.11d, 468.93d])!;

        TestAssert.True(Math.Abs(narrow - 0d) < 0.01d, "Narrow keeps full frame at the clearance boundary (exploded port). Got " + narrow);
        TestAssert.True(Math.Abs(mid - 42.72d) < 0.05d, "Mid legend crowds the ring by 42.72. Got " + mid);
        TestAssert.True(Math.Abs(wide - 109.28d) < 0.05d, "Wide legend crowds the ring by 109.28. Got " + wide);
        TestAssert.True(Math.Abs(shortPlot - 0d) < 0.000001d, "Short plots keep full frame at identical content (gap clears). Got " + shortPlot);
    }


    public static void PptxSyntheticExplodedDoughnutLayoutTranslatesRigidly()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var method = renderer.GetMethod("ResolvePieOrDoughnutLayout", flags);
        TestAssert.True(method is not null, "Expected polar layout helper to remain inspectable by the Office evidence guard.");
        var kindType = renderer.GetNestedType("ChartPolarKind", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected polar kind.");
        var plotBoxType = renderer.GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected plot box.");
        var legendType = renderer.GetNestedType("ChartLegendLayout", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected legend layout.");
        var instanceFlags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance;
        object doughnut = System.Enum.ToObject(kindType, 1);
        object plotBox = System.Activator.CreateInstance(plotBoxType, [144d, 72d, 576d, 432d])!;
        object hidden = legendType.GetProperty("Hidden", flags | System.Reflection.BindingFlags.Public)!.GetValue(null)!;
        object manual = legendType.GetProperty("Layout", instanceFlags)!.GetValue(hidden)!;
        object textBody = legendType.GetProperty("TextBodyProperties", instanceFlags)!.GetValue(hidden)!;
        object shape = legendType.GetProperty("ShapeStyle", instanceFlags)!.GetValue(hidden)!;
        object rightVisible = System.Activator.CreateInstance(legendType, [PptxSceneChartLegendPosition.Right, "r", false, true, manual, textBody, shape])!;
        var explosions = new Dictionary<int, double> { [0] = 0.1d };
        object legacy = method!.Invoke(null, [doughnut, plotBox, explosions, rightVisible, false, false, 0d, 0d, false])!;
        object shifted = method.Invoke(null, [doughnut, plotBox, explosions, rightVisible, false, false, 42.83d, 0d, false])!;
        (double lx, double ly, double lr) = ReadPolarGeometry(legacy);
        (double sx, double sy, double sr) = ReadPolarGeometry(shifted);
        TestAssert.True(Math.Abs(sr - lr) < 0.000001d && Math.Abs(sy - ly) < 0.000001d, "The reserve must not resize the ring, only translate it. Got " + sr + "/" + sy);
        TestAssert.True(Math.Abs((lx - sx) - 42.83d / 2d) < 0.01d, "The ring translates by half the reserve (1:2 ring:legend rule). Got " + (lx - sx));
    }


    public static void PptxSyntheticDoughnutLegendConstantsKeepOfficeCalibration()
    {
        var rules = typeof(PptxRenderer).GetNestedType("PptxChartMetricRules", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected metric rules.");
        double clearance = (double)rules.GetField("DoughnutExplodedLegendClearance")!.GetValue(null)!;
        double shift = (double)rules.GetField("DoughnutRightLegendVerticalShift")!.GetValue(null)!;
        double leftShift = (double)rules.GetField("DoughnutLeftLegendVerticalShift")!.GetValue(null)!;
        TestAssert.True(Math.Abs(clearance - 86.3d) < 0.000001d, "Exploded doughnuts should keep the 86.3pt ring clearance (narrow/mid/wide/short renders). Got " + clearance);
        TestAssert.True(Math.Abs(shift - 22.92d) < 0.000001d, "Doughnut right legends should keep the 22.92pt vertical shift (8 Office renders). Got " + shift);
        TestAssert.True(Math.Abs(leftShift - 5.0d) < 0.000001d, "Doughnut left legends should keep the 5.0pt vertical shift (2 Office renders). Got " + leftShift);
        double headInset = (double)rules.GetField("DoughnutLeftLegendHeadInset")!.GetValue(null)!;
        TestAssert.True(Math.Abs(headInset - 13.19d) < 0.000001d, "Doughnut left legends should keep the 13.19pt head inset (4 Office renders). Got " + headInset);
        double lead = (double)rules.GetField("DoughnutLeftRingCenterLead")!.GetValue(null)!;
        TestAssert.True(Math.Abs(lead - 2.05d) < 0.000001d, "Left rings should keep the 2.05pt center lead (4 Office box widths). Got " + lead);
        double band = (double)rules.GetField("DoughnutTitledCenterYOffset")!.GetValue(null)!;
        TestAssert.True(Math.Abs(band - 17.68d) < 0.000001d, "Titled doughnuts should keep the 17.68pt center band (8 Office renders). Got " + band);
    }

    public static void PptxSyntheticPolarTitleTopOffsetKeepsOfficeCalibration()
    {
        var rules = typeof(PptxRenderer).GetNestedType("PptxChartMetricRules", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected metric rules.");
        double offset = (double)rules.GetField("PolarTitleTopOffset")!.GetValue(null)!;
        TestAssert.True(Math.Abs(offset - 28.0d) < 0.000001d, "Polar auto titles should keep the 28pt top inset (full/short/moved frames). Got " + offset);
    }

    public static void PptxSyntheticDoughnutLeftGeometryFollowsLegendBox()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var method = renderer.GetMethod("ResolvePieOrDoughnutLayout", flags);
        TestAssert.True(method is not null, "Expected polar layout helper to remain inspectable by the Office evidence guard.");
        var kindType = renderer.GetNestedType("ChartPolarKind", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected polar kind.");
        var plotBoxType = renderer.GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected plot box.");
        var legendType = renderer.GetNestedType("ChartLegendLayout", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected legend layout.");
        var instanceFlags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance;
        object doughnut = System.Enum.ToObject(kindType, 1);
        object plotBox = System.Activator.CreateInstance(plotBoxType, [144d, 72d, 576d, 432d])!;
        object hidden = legendType.GetProperty("Hidden", flags | System.Reflection.BindingFlags.Public)!.GetValue(null)!;
        object manual = legendType.GetProperty("Layout", instanceFlags)!.GetValue(hidden)!;
        object textBody = legendType.GetProperty("TextBodyProperties", instanceFlags)!.GetValue(hidden)!;
        object shape = legendType.GetProperty("ShapeStyle", instanceFlags)!.GetValue(hidden)!;
        object leftVisible = System.Activator.CreateInstance(legendType, [PptxSceneChartLegendPosition.Left, "l", false, true, manual, textBody, shape])!;
        var empty = new Dictionary<int, double>();
        object narrow = method!.Invoke(null, [doughnut, plotBox, empty, leftVisible, false, false, 0d, 217.67d, false])!;
        object wide = method.Invoke(null, [doughnut, plotBox, empty, leftVisible, false, false, 0d, 296.57d, false])!;
        (double nx, double ny, double nr) = ReadPolarGeometry(narrow);
        (double wx, double wy, double wr) = ReadPolarGeometry(wide);
        TestAssert.True(Math.Abs((wx - nx) - (296.57d - 217.67d) / 2d) < 0.01d, "The ring must translate by half the box growth (1:2 rule). Got " + (wx - nx));
        TestAssert.True(Math.Abs(nr - 205.027d) < 0.05d, "Narrow left rings keep the baseline radius. Got " + nr);
        TestAssert.True(wr < nr, "Wide left rings must bind the frame margin, not keep baseline. Got " + wr + " vs " + nr);
    }

    public static void PptxSyntheticDoughnutTitledBandLowersRingCenter()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var method = renderer.GetMethod("ResolvePieOrDoughnutLayout", flags);
        TestAssert.True(method is not null, "Expected polar layout helper to remain inspectable by the Office evidence guard.");
        var kindType = renderer.GetNestedType("ChartPolarKind", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected polar kind.");
        var plotBoxType = renderer.GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected plot box.");
        var legendType = renderer.GetNestedType("ChartLegendLayout", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected legend layout.");
        var instanceFlags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance;
        object doughnut = System.Enum.ToObject(kindType, 1);
        object plotBox = System.Activator.CreateInstance(plotBoxType, [144d, 72d, 576d, 432d])!;
        object hidden = legendType.GetProperty("Hidden", flags | System.Reflection.BindingFlags.Public)!.GetValue(null)!;
        object manual = legendType.GetProperty("Layout", instanceFlags)!.GetValue(hidden)!;
        object textBody = legendType.GetProperty("TextBodyProperties", instanceFlags)!.GetValue(hidden)!;
        object shape = legendType.GetProperty("ShapeStyle", instanceFlags)!.GetValue(hidden)!;
        object rightVisible = System.Activator.CreateInstance(legendType, [PptxSceneChartLegendPosition.Right, "r", false, true, manual, textBody, shape])!;
        var empty = new Dictionary<int, double>();
        object untitled = method!.Invoke(null, [doughnut, plotBox, empty, rightVisible, false, false, 0d, 0d, false])!;
        object titled = method.Invoke(null, [doughnut, plotBox, empty, rightVisible, false, false, 0d, 0d, true])!;
        (double ux, double uy, double ur) = ReadPolarGeometry(untitled);
        (double tx, double ty, double tr) = ReadPolarGeometry(titled);
        TestAssert.True(Math.Abs(ty - 270.32d) < 0.01d, "Titled rings must sit on the Office band center (270.32). Got " + ty);
        TestAssert.True(Math.Abs(uy - 269.86d) < 0.01d, "Untitled right rings keep the legacy ratio center (unobserved Office case frozen). Got " + uy);
        TestAssert.True(Math.Abs(tx - ux) < 0.000001d, "The band must not move the ring horizontally. Got " + tx + " vs " + ux);
        TestAssert.True(Math.Abs(tr - 187.35d) < 0.01d, "Titled rings must fit between band and margin (Office 187.32). Got " + tr);
    }
    public static void PptxSyntheticDoughnutShortPlotRadiusFitsBand()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var method = renderer.GetMethod("ResolvePieOrDoughnutLayout", flags);
        TestAssert.True(method is not null, "Expected polar layout helper to remain inspectable by the Office evidence guard.");
        var kindType = renderer.GetNestedType("ChartPolarKind", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected polar kind.");
        var plotBoxType = renderer.GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected plot box.");
        var legendType = renderer.GetNestedType("ChartLegendLayout", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected legend layout.");
        var instanceFlags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance;
        object doughnut = System.Enum.ToObject(kindType, 1);
        object shortPlot = System.Activator.CreateInstance(plotBoxType, [144d, 288d, 576d, 216d])!;
        object hidden = legendType.GetProperty("Hidden", flags | System.Reflection.BindingFlags.Public)!.GetValue(null)!;
        object manual = legendType.GetProperty("Layout", instanceFlags)!.GetValue(hidden)!;
        object textBody = legendType.GetProperty("TextBodyProperties", instanceFlags)!.GetValue(hidden)!;
        object shape = legendType.GetProperty("ShapeStyle", instanceFlags)!.GetValue(hidden)!;
        object rightVisible = System.Activator.CreateInstance(legendType, [PptxSceneChartLegendPosition.Right, "r", false, true, manual, textBody, shape])!;
        var empty = new Dictionary<int, double>();
        object shortTitled = method!.Invoke(null, [doughnut, shortPlot, empty, rightVisible, false, false, 0d, 0d, true])!;
        (double sx, double sy, double sr) = ReadPolarGeometry(shortTitled);
        TestAssert.True(Math.Abs(sr - 79.35d) < 0.01d, "Short titled rings must fit between band and margin (Office 79.32). Got " + sr);
        TestAssert.True(Math.Abs(sy - 378.32d) < 0.01d, "Short titled rings must center below the band (Office 378.32). Got " + sy);
    }


    public static void PptxSyntheticPieAutoLabelConstantsKeepOfficeCalibration()
    {
        var rules = typeof(PptxRenderer).GetNestedType("PptxChartMetricRules", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected metric rules.");
        double radius = (double)rules.GetField("PieAutoDataLabelRadiusRatio")!.GetValue(null)!;
        double baseline = (double)rules.GetField("PieDataLabelBaselineFactor")!.GetValue(null)!;
        TestAssert.True(Math.Abs(radius - 0.74d) < 0.000001d, "Pie auto radius should stay at the 9-sample Office mean. Got " + radius);
        TestAssert.True(Math.Abs(baseline - 0.88d) < 0.000001d, "Pie baseline inset should stay at the measured ascent. Got " + baseline);
        double pitch = (double)rules.GetField("PieDataLabelLinePitchFactor")!.GetValue(null)!;
        TestAssert.True(Math.Abs(pitch - 1.22d) < 0.000001d, "Pie wrap pitch should stay at the measured line step. Got " + pitch);
    }

    public static void PptxSyntheticPieLabeledRadiusGateKeepsPlotHeight()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "SelectPieLabeledRadiusRatio",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie labeled-radius helper to remain inspectable by the Office evidence guard.");
        double shortPlot = (double)method!.Invoke(null, [300d])!;
        double tallPlot = (double)method.Invoke(null, [360d])!;
        TestAssert.True(Math.Abs(shortPlot - 0.40083d) < 0.00001d, "300H labeled pies should use the short radius ratio (Office 0.40083). Got " + shortPlot);
        TestAssert.True(Math.Abs(tallPlot - 0.40694d) < 0.00001d, "360H labeled pies should use the tall radius ratio (Office 0.40694). Got " + tallPlot);
    }

    public static void PptxSyntheticPieManualCircleGapKeepsOfficeCalibration()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputePieManualLabelCircleGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie manual circle-gap helper to remain inspectable by the Office evidence guard.");

        double landscape = (double)method!.Invoke(null, [146.5d])!;
        double small = (double)method.Invoke(null, [120.25d])!;
        TestAssert.True(Math.Abs(landscape - 8.292d) < 0.01d, "Landscape manual labels should anchor 8.29pt past the rim (joint ladder fit). Got " + landscape);
        TestAssert.True(Math.Abs(small - 6.806d) < 0.01d, "Small-plot manual labels should anchor 6.81pt past the rim (joint ladder fit). Got " + small);
    }

    public static void PptxSyntheticPieManualBoxHeightKeepsOfficeCalibration()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputePieManualLabelBoxHeight",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie manual box-height helper to remain inspectable by the Office evidence guard.");

        double single = (double)method!.Invoke(null, [18d, 1])!;
        double wrapped = (double)method.Invoke(null, [18d, 2])!;
        double large = (double)method.Invoke(null, [24d, 2])!;
        TestAssert.True(Math.Abs(single - 24.96d) < 0.01d, "Single-line manual boxes should be pitch plus pad (Office 24.97). Got " + single);
        TestAssert.True(Math.Abs(wrapped - 46.92d) < 0.01d, "Two-line manual boxes should be two pitches plus pad (Office 46.94). Got " + wrapped);
        TestAssert.True(Math.Abs(large - 61.56d) < 0.01d, "Large-font wrapped boxes should scale by pitch (Office 61.59). Got " + large);
    }

    public static void PptxSyntheticPieManualBaselinePadKeepsOfficeCalibration()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputePieManualLabelBaselinePad",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie manual baseline-pad helper to remain inspectable by the Office evidence guard.");

        double basePad = (double)method!.Invoke(null, [18d])!;
        double largePad = (double)method.Invoke(null, [24d])!;
        TestAssert.True(Math.Abs(basePad - 6.54d) < 0.01d, "Manual baselines should sit 6.5pt above the box bottom at 18pt. Got " + basePad);
        TestAssert.True(Math.Abs(largePad - 8.22d) < 0.01d, "Manual baselines should sit 8.2pt above the box bottom at 24pt. Got " + largePad);
    }

    public static void PptxSyntheticPieManualEdgeClampKeepsPlot()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ClampPieManualLabelEdge",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected pie manual edge-clamp helper to remain inspectable by the Office evidence guard.");

        double inside = (double)method!.Invoke(null, [158d, 88.64d, 144d, 520d])!;
        double pastRight = (double)method.Invoke(null, [700d, 80d, 144d, 520d])!;
        double pastLeft = (double)method.Invoke(null, [0d, 80d, 144d, 520d])!;
        double oversize = (double)method.Invoke(null, [0d, 600d, 144d, 520d])!;
        TestAssert.True(Math.Abs(inside - 158d) < 0.01d, "Inside edges should keep the anchored position. Got " + inside);
        TestAssert.True(Math.Abs(pastRight - 584d) < 0.01d, "Right-overflow edges should freeze at the plot edge minus width. Got " + pastRight);
        TestAssert.True(Math.Abs(pastLeft - 144d) < 0.01d, "Left-overflow edges should freeze at the plot edge (Office Gamma 144.0). Got " + pastLeft);
        TestAssert.True(Math.Abs(oversize - 144d) < 0.01d, "Oversize boxes should pin to the plot origin. Got " + oversize);
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
        TestAssert.True(Math.Abs(doughnutPortrait - 432d) < 0.000001d, "Portrait doughnut probes use the height base with the width-margin min. Got " + doughnutPortrait);
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
        // Unknown-target manuals fall back to the default plot box, which now carries
        // the measured horizontal top and bottom strips (axis-less charts still render
        // default value ticks, so the strips are legitimate; X and size are untouched,
        // bars and labels fit inside as verified by inspection).
        TestAssert.Contains("88.589 349 144 108 re f", pdf);
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
    }    public static void PptxSyntheticChartColumnAxisNearMaximumKeepsOfficeHeadroom()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "GetNiceChartAxisMax",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart axis maximum helper to remain inspectable by the Office evidence guard.");
        // Office ceilings dataMax 68 to 80 on the ladder column axis-titles probe (unit 10 kept).
        TestAssert.Equal(80d, (double)method!.Invoke(null, [68d, 0d, 9d, true, 0.96d, false])!);

        var resolveHeadroom = typeof(PptxRenderer).GetMethod(
            "ResolveBarValueAxisHeadroom",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected bar value-axis headroom bridge.");
        // Only vertical non-percent bars take the shared headroom rule; horizontal bars
        // keep the frozen behavior their tuned right-margin laws rely on.
        TestAssert.True((bool)resolveHeadroom.Invoke(null, [false, false])!, "Expected vertical non-percent columns to take Office headroom.");
        TestAssert.True(!(bool)resolveHeadroom.Invoke(null, [true, false])!, "Expected horizontal bars to keep frozen behavior.");
        TestAssert.True(!(bool)resolveHeadroom.Invoke(null, [false, true])!, "Expected percent stacks to keep frozen behavior.");
        TestAssert.True(!(bool)resolveHeadroom.Invoke(null, [true, true])!, "Expected horizontal percent stacks to keep frozen behavior.");
    }    public static void PptxDualAxisStripFallbackUsesOwningPlotData()
    {
        string chartXmlText = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:ser><c:tx><c:v>Primary</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>30</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:barChart>
                <c:barChart>
                  <c:ser><c:tx><c:v>Secondary</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>50</c:v></c:pt><c:pt idx="1"><c:v>100</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="30"/><c:axId val="40"/>
                </c:barChart>
                <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
                <c:catAx><c:axId val="30"/><c:axPos val="b"/><c:crossAx val="40"/></c:catAx>
                <c:valAx><c:axId val="40"/><c:axPos val="l"/><c:crossAx val="30"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """;
        PptxSceneChart chart = PptxTests.BuildSingleChartScene(chartXmlText) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(2, chart.Plots.Count);
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XDocument chartXml = XDocument.Parse(chartXmlText);
        System.Collections.Generic.List<XElement> barCharts = chartXml.Descendants(c + "barChart").ToList();
        TestAssert.Equal(2, barCharts.Count);
        System.Reflection.MethodInfo fallback = typeof(PptxRenderer).GetMethod(
            "GetDualAxisStripFallbackExtents",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected dual-axis strip fallback bridge.");
        // The secondary axis must resolve its owning plot data (clustered max 100), not
        // the legacy (0,1) dummy that froze auto-max dual reserves on the Office probes.
        object secondary = fallback.Invoke(null, [chart, chartXml, chart.Plots, barCharts, "40", null, false]) ?? throw new InvalidOperationException("Expected secondary fallback extents.");
        TestAssert.Equal(100d, (double)secondary.GetType().GetProperty("Max")?.GetValue(secondary)!);
        TestAssert.Equal(0d, (double)secondary.GetType().GetProperty("Min")?.GetValue(secondary)!);
        object unknown = fallback.Invoke(null, [chart, chartXml, chart.Plots, barCharts, "99", null, false]) ?? throw new InvalidOperationException("Expected unknown-axis fallback extents.");
        TestAssert.Equal(1d, (double)unknown.GetType().GetProperty("Max")?.GetValue(unknown)!);
    }    public static void PptxNoTitleBottomLegendReserveMatchesOfficePlane()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeNoTitleBottomLegendReserve",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected bottom-legend reserve bridge.");
        // Four same-frame Office probes pin the additive plane (all within 0.06); the
        // legacy single-knob formula misses by plus-1.1 to minus-1.9 across the same set.
        TestAssert.True(Math.Abs((double)method.Invoke(null, [18d, 8.04d])! - 55.02d) < 0.1, "Expected ladder reserve near 55.02.");
        TestAssert.True(Math.Abs((double)method.Invoke(null, [12d, 8.04d])! - 47.77d) < 0.1, "Expected leg12 reserve near 47.77.");
        TestAssert.True(Math.Abs((double)method.Invoke(null, [18d, 14.04d])! - 66.10d) < 0.1, "Expected cat14 reserve near 66.10.");
        TestAssert.True(Math.Abs((double)method.Invoke(null, [12d, 14.04d])! - 58.86d) < 0.1, "Expected leg12cat14 reserve near 58.86.");
    }    public static void PptxSyntheticChartSeriesLineUsesRoundCapsAndJoins()
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
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart>
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser>
                          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Wave</c:v></c:pt></c:strCache></c:strRef></c:tx>
                          <c:marker><c:symbol val="none"/></c:marker>
                          <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                        <c:axId val="10"/><c:axId val="20"/>
                      </c:lineChart>
                      <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx>
                      <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Contains("1 J"), "Expected round line caps on chart series lines.");
        TestAssert.True(pdf.Contains("1 j"), "Expected round line joins on chart series lines.");
    }

    public static void PptxBarPlotClipPadsBottomAndRight()
    {
        // Office bar/column plot clips extend past the axis-bounded plot rect on the
        // bottom and right edges only (horizontal stacked -0.68/+0.68, clustered
        // -0.68/+0.72, axis-titles -0.71/+0.69, shifted-frame -0.71/+0.65; vertical
        // column-stacked -0.68/+0.68, column-clustered -0.68/+0.68).
        Type plotBoxType = typeof(PptxRenderer).GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart plot box.");
        System.Reflection.MethodInfo clip = typeof(PptxRenderer).GetMethod("GetBarPlotClipBox", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected bar clip helper.");
        object plot = Activator.CreateInstance(plotBoxType, [145.2d, 111.92d, 622.12d, 381.08d]) ?? throw new InvalidOperationException("Expected chart plot box.");
        object padded = clip.Invoke(null, [plot]) ?? throw new InvalidOperationException("Expected padded clip box.");
        Type boxType = padded.GetType();
        TestAssert.True(Math.Abs((double)(boxType.GetProperty("X")?.GetValue(padded) ?? 0d) - 145.2d) < 1e-9, "Clip pad must not move the left edge.");
        TestAssert.True(Math.Abs((double)(boxType.GetProperty("Y")?.GetValue(padded) ?? 0d) - 111.23d) < 1e-9, "Clip pad must extend 0.69 below the axis edge.");
        TestAssert.True(Math.Abs((double)(boxType.GetProperty("Width")?.GetValue(padded) ?? 0d) - 622.81d) < 1e-9, "Clip pad must extend 0.69 past the right edge.");
        TestAssert.True(Math.Abs((double)(boxType.GetProperty("Height")?.GetValue(padded) ?? 0d) - 381.77d) < 1e-9, "Clip height must grow with the bottom pad.");
    }

    public static void PptxUnstyledLineStrokeTintMatchesOfficeLuminance()
    {
        // Office unstyled line-chart series strokes carry the raw theme accent through
        // 97.5% HSL luminance (3 cached Office refs: blue/red/green series, 9/9 bytes
        // exact); fills and explicitly styled strokes keep raw colors.
        System.Reflection.MethodInfo tint = typeof(PptxRenderer).GetMethod(
            "ApplyUnstyledLineStrokeTint",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected line-stroke tint helper.");
        TestAssert.Equal(new RgbColor(74, 126, 187), (RgbColor)tint.Invoke(null, [new RgbColor(79, 129, 189)])!);
        TestAssert.Equal(new RgbColor(190, 75, 72), (RgbColor)tint.Invoke(null, [new RgbColor(192, 80, 77)])!);
        TestAssert.Equal(new RgbColor(152, 185, 84), (RgbColor)tint.Invoke(null, [new RgbColor(155, 187, 89)])!);
    }

    public static void PptxChartAxisDefaultStrokeIsBlack()
    {
        // Office draws unstyled axes and ticks black across bar, column, line,
        // scatter, and area refs; the 90-gray default is killed.
        System.Reflection.PropertyInfo axisDefault = typeof(PptxRenderer).GetProperty(
            "ChartAxisDefaultStroke",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected axis default stroke.");
        object stroke = axisDefault.GetValue(null) ?? throw new InvalidOperationException("Expected axis default stroke value.");
        TestAssert.Equal(new RgbColor(0, 0, 0), (RgbColor)stroke.GetType().GetProperty("Color")?.GetValue(stroke)!);
    }

    public static void PptxGalleryAxisFamilyDefaultNeedsStyleTwoWithoutPart()
    {
        // Gallery style 2 without a style part renders unstyled axis-family strokes
        // gray-0.537 at width 1.0; any other gallery id, missing id, or style-part
        // chart keeps the null default (legacy black fallback downstream).
        System.Reflection.MethodInfo gallery = typeof(PptxRenderer).GetMethod(
            "ResolveGalleryAxisFamilyDefault",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected gallery axis default.");
        object styled = gallery.Invoke(null, ["2", false]) ?? throw new InvalidOperationException("Expected gallery default stroke.");
        TestAssert.Equal(new RgbColor(137, 137, 137), (RgbColor)styled.GetType().GetProperty("Color")?.GetValue(styled)!);
        TestAssert.Equal(1d, (double)styled.GetType().GetProperty("Width")?.GetValue(styled)!);
        TestAssert.True(gallery.Invoke(null, ["18", false]) is null, "Gallery default must not touch other style ids.");
        TestAssert.True(gallery.Invoke(null, [null, false]) is null, "Gallery default must not touch style-less charts.");
        TestAssert.True(gallery.Invoke(null, ["2", true]) is null, "Gallery default must not touch style-part charts.");
    }

    public static void PptxBubbleMaxRadiusMatchesOfficePlotRatio()
    {
        // Office renders the largest bubble at 0.131 of the smaller plot dimension
        // (bubble port: Office max diameter 89.22 over plot min 341.52; relative sqrt
        // sizing already exact across all four bubbles).
        System.Type renderer = typeof(PptxRenderer);
        System.Type plotBoxType = renderer.GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart plot box.");
        System.Type extentsType = renderer.GetNestedType("ChartValueExtents", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart value extents.");
        System.Type pointType = renderer.GetNestedType("ScatterPoint", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected scatter point.");
        System.Type indexedPointType = renderer.GetNestedType("ChartIndexedNumberPoint", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected indexed number point.");
        System.Type cellType = renderer.GetNestedType("ChartWorkbookRangeCell", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected workbook range cell.");
        System.Type sourceType = renderer.GetNestedType("ChartPointIndexSource", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected point index source.");
        System.Reflection.MethodInfo geometry = renderer.GetMethod("ResolveScatterPointGeometry", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected scatter geometry helper.");
        object indexSource = System.Enum.GetValues(sourceType).GetValue(0) ?? throw new InvalidOperationException("Expected index source value.");
        object cell = System.Activator.CreateInstance(cellType) ?? throw new InvalidOperationException("Expected default cell.");
        object indexed = System.Activator.CreateInstance(indexedPointType, [0, indexSource, null, "", false, cell]) ?? throw new InvalidOperationException("Expected indexed point.");
        object plot = System.Activator.CreateInstance(plotBoxType, [0d, 0d, 200d, 100d]) ?? throw new InvalidOperationException("Expected plot box.");
        object extents = System.Activator.CreateInstance(extentsType, [0d, 10d]) ?? throw new InvalidOperationException("Expected extents.");
        double RadiusOf(object point)
        {
            object result = geometry.Invoke(null, [plot, point, true, extents, extents, 16d]) ?? throw new InvalidOperationException("Expected geometry result.");
            return (double)(result.GetType().GetField("Item3")?.GetValue(result) ?? 0d);
        }
        object maxPoint = System.Activator.CreateInstance(pointType, [1d, 2d, 16d, 0, indexed, indexed, null, null, null, null, null, null]) ?? throw new InvalidOperationException("Expected max point.");
        TestAssert.True(System.Math.Abs(RadiusOf(maxPoint) - 13.1d) < 1e-9, "Max bubble radius drifts from the Office 0.131 plot ratio.");
        object quarterPoint = System.Activator.CreateInstance(pointType, [1d, 2d, 4d, 0, indexed, indexed, null, null, null, null, null, null]) ?? throw new InvalidOperationException("Expected quarter point.");
        TestAssert.True(System.Math.Abs(RadiusOf(quarterPoint) - 6.55d) < 1e-9, "Bubble sqrt sizing must halve radius at quarter size.");
    }
}
