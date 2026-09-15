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
        int[] tailR = [199, 251];
        int[] tailG = [222, 215];
        int[] tailB = [231, 199];
        for (int slot = 0; slot < 2; slot++)
        {
            object?[] tailArgs = [94 + slot, null];
            TestAssert.Equal(true, (bool)tail!.Invoke(null, tailArgs)!);
            TestAssert.Equal((byte)tailR[slot], (byte)rgbType.GetProperty("Red")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailG[slot], (byte)rgbType.GetProperty("Green")!.GetValue(tailArgs[1])!);
            TestAssert.Equal((byte)tailB[slot], (byte)rgbType.GetProperty("Blue")!.GetValue(tailArgs[1])!);
        }
        object?[] past = [96, null];
        TestAssert.Equal(false, (bool)tail.Invoke(null, past)!);
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
