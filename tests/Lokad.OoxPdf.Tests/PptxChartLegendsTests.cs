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

internal static class PptxChartLegendsTests
{
    public static void PptxSyntheticChartLegendManualBoxDrivesPlacement()
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
                        <c:ser>
                          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Manual Legend</c:v></c:pt></c:strCache></c:strRef></c:tx>
                          <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                      </c:lineChart>
                    </c:plotArea>
                    <c:legend>
                      <c:legendPos val="b"/>
                      <c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="edge"/><c:hMode val="edge"/><c:x val="0.5"/><c:y val="0.5"/><c:w val="0.8"/><c:h val="0.6"/></c:manualLayout></c:layout>
                    </c:legend>
                  </c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"1 0 0 1 [0-9.]+ 338\.4 Tm"), "Expected explicit chart legend manualLayout to drive the legend text baseline.");
    }

    public static void PptxSyntheticChartStyleLegendDefaultsDriveChartLegendRendering()
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
                    <c:legend><c:legendPos val="r"/></c:legend>
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser>
                          <c:tx><c:v>Styled Legend</c:v></c:tx>
                          <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                      </c:lineChart>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/style1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="78">
                  <cs:legend>
                    <cs:defRPr sz="1150" u="sng" strike="sngStrike">
                      <a:solidFill><a:srgbClr val="2468AC"/></a:solidFill>
                      <a:latin typeface="Arial"/>
                    </cs:defRPr>
                  </cs:legend>
                </cs:style>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.141 0.408 0.675 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CL[0-9]+ 11\.52 Tf"), "Expected chart-style legend role font size to drive legend rendering when c:legend has no direct txPr.");
        int decorationPathCount = PptxTests.CountFilledDecorationPaths(pdf, @"0\.141 0\.408 0\.675");
        TestAssert.True(decorationPathCount >= 2, "Expected chart-style legend role underline and strike to emit filled decoration paths through the common text renderer.");
    }

    public static void PptxSyntheticChartStyleDataLabelDefaultsDriveDataLabelRendering()
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
                    <c:plotArea>
                      <c:lineChart>
                        <c:dLbls><c:showVal val="1"/><c:dLblPos val="t"/></c:dLbls>
                        <c:ser>
                          <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                      </c:lineChart>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/style1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="79">
                  <cs:dataLabel>
                    <cs:defRPr sz="1400" u="sng" strike="sngStrike">
                      <a:solidFill><a:srgbClr val="33AA66"/></a:solidFill>
                      <a:latin typeface="Arial"/>
                    </cs:defRPr>
                  </cs:dataLabel>
                </cs:style>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.2 0.667 0.4 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CLD[0-9]+ 14\.04 Tf"), "Expected chart-style dataLabel role font size to drive data-label rendering when c:dLbls has no direct txPr.");
        int decorationPathCount = PptxTests.CountFilledDecorationPaths(pdf, @"0\.2 0\.667 0\.4");
        TestAssert.True(decorationPathCount >= 2, "Expected chart-style dataLabel role underline and strike to emit filled decoration paths through the common text renderer.");
    }

    public static void PptxSyntheticChartStyleAxisDefaultsDriveTickLabelRendering()
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
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser>
                          <c:cat><c:strLit><c:pt idx="0"><c:v>Alpha</c:v></c:pt><c:pt idx="1"><c:v>Beta</c:v></c:pt></c:strLit></c:cat>
                          <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                        <c:axId val="10"/>
                        <c:axId val="20"/>
                      </c:lineChart>
                      <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Category Axis</a:t></a:r></a:p></c:rich></c:tx></c:title><c:tickLblPos val="nextTo"/><c:crossAx val="20"/></c:catAx>
                      <c:valAx>
                        <c:axId val="20"/>
                        <c:axPos val="l"/>
                        <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Value Axis</a:t></a:r></a:p></c:rich></c:tx></c:title>
                        <c:scaling><c:min val="0"/><c:max val="10"/></c:scaling>
                        <c:majorUnit val="5"/>
                        <c:tickLblPos val="nextTo"/>
                        <c:crossAx val="10"/>
                      </c:valAx>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/style1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="80">
                  <cs:categoryAxis>
                    <cs:defRPr sz="1300" u="sng" strike="sngStrike">
                      <a:solidFill><a:srgbClr val="AA5500"/></a:solidFill>
                      <a:latin typeface="Arial"/>
                    </cs:defRPr>
                  </cs:categoryAxis>
                  <cs:valueAxis>
                    <cs:defRPr sz="1400" u="sng" strike="sngStrike">
                      <a:solidFill><a:srgbClr val="0055AA"/></a:solidFill>
                      <a:latin typeface="Arial"/>
                    </cs:defRPr>
                  </cs:valueAxis>
                </cs:style>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.667 0.333 0 rg", pdf);
        TestAssert.Contains("0 0.333 0.667 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CCA[0-9]+ 12\.96 Tf"), "Expected chart-style categoryAxis role font size to drive category tick labels when c:catAx has no direct txPr.");
        TestAssert.True(Regex.IsMatch(pdf, @"/CVA[0-9]+ 14\.04 Tf"), "Expected chart-style valueAxis role font size to drive value tick labels when c:valAx has no direct txPr.");
        TestAssert.True(Regex.IsMatch(pdf, @"/CAT[0-9]+ 12\.96 Tf"), "Expected chart-style categoryAxis role font size to drive default category axis-title rendering.");
        TestAssert.True(Regex.IsMatch(pdf, @"/CAT[0-9]+ 14\.04 Tf"), "Expected chart-style valueAxis role font size to drive default value axis-title rendering.");
        int categoryDecorationPathCount = PptxTests.CountFilledDecorationPaths(pdf, @"0\.667 0\.333 0");
        int valueDecorationPathCount = PptxTests.CountFilledDecorationPaths(pdf, @"0 0\.333 0\.667");
        TestAssert.True(categoryDecorationPathCount >= 2, "Expected chart-style categoryAxis role underline and strike to emit filled decoration paths.");
        TestAssert.True(valueDecorationPathCount >= 2, "Expected chart-style valueAxis role underline and strike to emit filled decoration paths.");
    }

    public static void PptxSyntheticChartStyleAxisDefaultsDriveManualAxisTitleRendering()
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
                    <c:plotArea>
                      <c:lineChart>
                        <c:ser>
                          <c:cat><c:strLit><c:pt idx="0"><c:v>Alpha</c:v></c:pt><c:pt idx="1"><c:v>Beta</c:v></c:pt></c:strLit></c:cat>
                          <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                        <c:axId val="10"/>
                        <c:axId val="20"/>
                      </c:lineChart>
                      <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="20"/></c:catAx>
                      <c:valAx>
                        <c:axId val="20"/>
                        <c:axPos val="l"/>
                        <c:title>
                          <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Manual Value Axis</a:t></a:r></a:p></c:rich></c:tx>
                          <c:layout><c:manualLayout><c:x val="0.05"/><c:y val="0.15"/><c:w val="0.35"/><c:h val="0.12"/></c:manualLayout></c:layout>
                        </c:title>
                        <c:scaling><c:min val="0"/><c:max val="10"/></c:scaling>
                        <c:majorUnit val="5"/>
                        <c:tickLblPos val="nextTo"/>
                        <c:crossAx val="10"/>
                      </c:valAx>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """),
            ["ppt/charts/style1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="81">
                  <cs:valueAxis>
                    <cs:defRPr sz="1400" u="sng" strike="sngStrike">
                      <a:solidFill><a:srgbClr val="0055AA"/></a:solidFill>
                      <a:latin typeface="Arial"/>
                    </cs:defRPr>
                  </cs:valueAxis>
                </cs:style>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0.333 0.667 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CAT[0-9]+ 14\.04 Tf"), "Expected chart-style valueAxis role font size to drive manual-layout value axis-title rendering.");
        int decorationPathCount = PptxTests.CountFilledDecorationPaths(pdf, @"0 0\.333 0\.667");
        TestAssert.True(decorationPathCount >= 2, "Expected chart-style valueAxis role underline and strike to emit filled decoration paths for manual-layout axis titles.");
    }

    public static void PptxSyntheticChartManualLayoutMissingPositionUsesDefaultPlotBox()
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
                    <c:layout><c:manualLayout><c:wMode val="factor"/><c:hMode val="factor"/><c:w val="0.5"/><c:h val="0.5"/></c:manualLayout></c:layout>
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
        TestAssert.DoesNotContain("106.56 286.56 218.88 146.88 re f", pdf);
    }

    public static void PptxSyntheticAreaChartManualPlotBoxUsesSharedPath()
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
                    <c:areaChart>
                      <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                    </c:areaChart>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("129.6 338.4 m 273.6 338.4 l S", pdf);
        TestAssert.DoesNotContain("106.56 321.12 m 325.44 321.12 l S", pdf);
    }

    public static void PptxSyntheticRadarChartManualPlotBoxUsesSharedPath()
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
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart><c:plotArea>
                    <c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="edge"/><c:hMode val="edge"/><c:x val="0.2"/><c:y val="0.1"/><c:w val="0.7"/><c:h val="0.6"/></c:manualLayout></c:layout>
                    <c:radarChart>
                      <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                    </c:radarChart>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("201.6 392.4 m", pdf);
        TestAssert.Contains("201.6 438.041 l", pdf);
        TestAssert.DoesNotContain("201.6 392.4 m 201.6 438.041 l S", pdf);
        TestAssert.DoesNotContain("216 364.32 m 216 295.2 l S", pdf);
    }

    public static void PptxSyntheticPieChartManualPlotBoxUsesPolarPath()
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
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <c:chart><c:plotArea>
                    <c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/><c:wMode val="edge"/><c:hMode val="edge"/><c:x val="0.2"/><c:y val="0.1"/><c:w val="0.7"/><c:h val="0.6"/></c:manualLayout></c:layout>
                    <c:pieChart>
                      <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>100</c:v></c:pt></c:numLit></c:val></c:ser>
                    </c:pieChart>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" f", pdf);
        TestAssert.DoesNotContain("204.48 364.32 m", pdf);
        TestAssert.DoesNotContain("204.48 290.88 l", pdf);
    }

    public static void PptxSyntheticBarChartOverlayLegendDoesNotReservePlotSpace()
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
                  <c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
                    <c:spPr><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill></c:spPr>
                    <c:barChart>
                      <c:ser><c:tx><c:v>Series</c:v></c:tx><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                    </c:barChart>
                  </c:plotArea><c:legend><c:legendPos val="b"/><c:overlay/></c:legend></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 1 1 rg", pdf);
        // The overlay legend still reserves nothing; the plot box now also fits the rendered
        // value tick labels under the recalibrated label reserve (frame indent plus
        // font-relative gap).
        TestAssert.Contains("98.136 271.958 257.486 188.05 re f", pdf);
        TestAssert.DoesNotContain("104.256 259.56 247.68 191.16 re f", pdf);
        TestAssert.DoesNotContain("100.8 282.24 236.16 174.96 re f", pdf);
    }

    public static void PptxSyntheticChartAutoTitleUsesOfficeScaledSize()
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
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800"><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
                  <c:chart><c:autoTitleDeleted val="0"/><c:plotArea>
                    <c:barChart><c:barDir val="bar"/><c:ser><c:tx><c:v>Series</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneChartTitle? sceneTitle = scene.Slides[0].SlideNodes[0].Chart?.Title;
        TestAssert.Equal("Series", sceneTitle?.Text ?? string.Empty);
        TestAssert.True(sceneTitle?.IsAutoGenerated == true, "Expected the typed scene chart model to own auto-title inference.");
        TestAssert.True(sceneTitle?.IsAutoDeleted == false, "Expected explicit autoTitleDeleted=false metadata to remain distinct from a missing element.");
        TestAssert.Equal("0", sceneTitle?.IsAutoDeletedValue ?? string.Empty);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"/CT[0-9]+ 21\.6 Tf"), "Expected absent auto chart title to use Office's 120% title text scale over chart txPr.");
        TestAssert.True(pdf.Split("/FontFile2", StringSplitOptions.None).Length - 1 >= 2, "Expected absent auto chart title to request a distinct embedded face for Office's bold title default.");
        TestAssert.True(Regex.IsMatch(pdf, @"/C[AV][A-Z][0-9]+ 18 Tf"), "Expected chart axis text to keep the unscaled chart txPr size.");
    }

    public static void PptxChartAutoTitleUsesSingleSeriesNameForLineCharts()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:autoTitleDeleted val="0"/><c:plotArea>
                <c:lineChart><c:ser><c:tx><c:v>Line Series</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartTitle? title = chart?.Title;
        TestAssert.Equal("Line Series", title?.Text ?? string.Empty);
        TestAssert.True(title?.IsAutoGenerated == true, "Expected auto-title inference to use typed series metadata beyond bar charts.");
    }

    public static void PptxChartAutoTitleRequiresExplicitAutoTitleDeletedFalse()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:scatterChart><c:ser><c:tx><c:v>Scatter Series</c:v></c:tx><c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal><c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal></c:ser></c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartTitle? title = chart?.Title;
        TestAssert.True(title?.Text is null, "Expected missing autoTitleDeleted metadata to remain titleless.");
        TestAssert.True(title?.IsAutoGenerated == false, "Expected auto-title inference only when Office metadata explicitly enables it.");
        TestAssert.True(title?.IsAutoDeleted is null, "Expected missing autoTitleDeleted metadata to remain distinguishable from false.");
    }

    public static void PptxChartShapeStylePreservesPictureFillTokens()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
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
                <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneShapePictureFill chartArea = chart?.ChartAreaStyle.PictureFill
            ?? throw new InvalidOperationException("Expected chart area picture fill.");
        TestAssert.True(chartArea.HasPicture, "Expected chart area picture fill to be scene-visible.");
        TestAssert.Equal("rIdImage", chartArea.RelationshipId);
        TestAssert.Equal("10000", chartArea.Crop.LeftValue ?? string.Empty);
        TestAssert.Equal("20000", chartArea.Crop.TopValue ?? string.Empty);
        TestAssert.Equal("30000", chartArea.Fill.RightValue ?? string.Empty);
        TestAssert.Equal("40000", chartArea.Fill.BottomValue ?? string.Empty);
        TestAssert.Equal(0.42d, chartArea.Alpha);
        TestAssert.Equal("42000", chartArea.AlphaValue ?? string.Empty);

        PptxSceneShapePictureFill plotArea = chart?.PlotAreaStyle.PictureFill
            ?? throw new InvalidOperationException("Expected plot area picture fill.");
        TestAssert.True(plotArea.HasPicture, "Expected plot area picture fill to be scene-visible even without a resolvable relationship.");
        TestAssert.Equal(string.Empty, plotArea.RelationshipId);
        TestAssert.Equal(0.625d, plotArea.Alpha);
        TestAssert.Equal("62500", plotArea.AlphaValue ?? string.Empty);
        TestAssert.True(plotArea.Tile.HasTile, "Expected chart plot area picture tile mode to be preserved.");
        TestAssert.Equal("ctr", plotArea.Tile.AlignmentValue ?? string.Empty);
        TestAssert.Equal("x", plotArea.Tile.FlipValue ?? string.Empty);
    }

    public static void PptxChartShapeStylePreservesUnsupportedEffectTokens()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:spPr>
                <a:effectLst>
                  <a:outerShdw dist="12700" dir="0"><a:srgbClr val="010203"/></a:outerShdw>
                  <a:reflection blurRad="6350"/>
                  <a:softEdge rad="12700"/>
                </a:effectLst>
              </c:spPr>
              <c:chart><c:plotArea>
                <c:spPr><a:effectDag/></c:spPr>
                <c:lineChart><c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartEffectFamily chartArea = chart?.ChartAreaStyle.Effects
            ?? throw new InvalidOperationException("Expected chart area effects.");
        TestAssert.True(chartArea.HasEffectList, "Expected chart area effect list presence to be preserved.");
        TestAssert.True(!chartArea.HasEffectDag, "Expected chart area effectDag absence to remain distinct.");
        TestAssert.Equal("reflection,softEdge", string.Join(",", chartArea.UnsupportedEffectNames));
        TestAssert.True(chart?.ChartAreaStyle.OuterShadow.HasShadow == true, "Expected supported outer shadow to remain parsed beside unsupported effects.");

        PptxSceneChartEffectFamily plotArea = chart?.PlotAreaStyle.Effects
            ?? throw new InvalidOperationException("Expected plot area effects.");
        TestAssert.True(!plotArea.HasEffectList, "Expected plot area effect list absence to remain distinct.");
        TestAssert.True(plotArea.HasEffectDag, "Expected plot area effectDag presence to be preserved.");
        TestAssert.Equal(string.Empty, string.Join(",", plotArea.UnsupportedEffectNames));
    }

    public static void PptxScenePreservesChartSeriesDataSourceReferences()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:ser>
                  <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache/></c:strRef></c:tx>
                  <c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$A$4</c:f><c:multiLvlStrCache/></c:multiLvlStrRef></c:cat>
                  <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache/></c:numRef></c:val>
                </c:ser></c:lineChart>
                <c:scatterChart><c:ser>
                  <c:xVal><c:numRef><c:f>Sheet2!$A$2:$A$4</c:f><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:xVal>
                  <c:yVal><c:numRef><c:f>Sheet2!$B$2:$B$4</c:f><c:numCache><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:yVal>
                  <c:bubbleSize><c:numRef><c:f>Sheet2!$C$2:$C$4</c:f><c:numCache/></c:numRef></c:bubbleSize>
                </c:ser></c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartSeries lineSeries = chart?.Plots[0].Series[0] ?? throw new InvalidOperationException("Expected line chart series.");
        TestAssert.Equal("Sheet1!$B$1", lineSeries.DataSources.Name.Formula ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.StringReference, lineSeries.DataSources.Name.ReferenceKindValue);
        TestAssert.Equal("strRef", lineSeries.DataSources.Name.ReferenceKind);
        TestAssert.Equal(PptxSceneChartDataSourceCacheKind.StringCache, lineSeries.DataSources.Name.CacheKindValue);
        TestAssert.Equal("strCache", lineSeries.DataSources.Name.CacheKind);
        TestAssert.True(!lineSeries.DataSources.Name.HasCachedPoints, "Expected empty string cache to preserve the formula without inventing cached points.");
        TestAssert.Equal("Sheet1!$A$2:$A$4", lineSeries.DataSources.Categories.Formula ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.MultiLevelStringReference, lineSeries.DataSources.Categories.ReferenceKindValue);
        TestAssert.Equal("multiLvlStrRef", lineSeries.DataSources.Categories.ReferenceKind);
        TestAssert.Equal(PptxSceneChartDataSourceCacheKind.MultiLevelStringCache, lineSeries.DataSources.Categories.CacheKindValue);
        TestAssert.Equal("multiLvlStrCache", lineSeries.DataSources.Categories.CacheKind);
        TestAssert.Equal("Sheet1!$B$2:$B$4", lineSeries.DataSources.Values.Formula ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.NumberReference, lineSeries.DataSources.Values.ReferenceKindValue);
        TestAssert.Equal("numRef", lineSeries.DataSources.Values.ReferenceKind);
        TestAssert.Equal(PptxSceneChartDataSourceCacheKind.NumberCache, lineSeries.DataSources.Values.CacheKindValue);
        TestAssert.Equal("numCache", lineSeries.DataSources.Values.CacheKind);
        TestAssert.True(!lineSeries.DataSources.Values.HasCachedPoints, "Expected formula-only numeric references to remain distinguishable from cached values.");

        PptxSceneChartSeries scatterSeries = chart?.Plots[1].Series[0] ?? throw new InvalidOperationException("Expected scatter chart series.");
        TestAssert.Equal("Sheet2!$A$2:$A$4", scatterSeries.DataSources.XValues.Formula ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.NumberReference, scatterSeries.DataSources.XValues.ReferenceKindValue);
        TestAssert.Equal(PptxSceneChartDataSourceCacheKind.NumberCache, scatterSeries.DataSources.XValues.CacheKindValue);
        TestAssert.True(scatterSeries.DataSources.XValues.HasCachedPoints, "Expected x-value cache point presence in the scene model.");
        TestAssert.Equal("Sheet2!$B$2:$B$4", scatterSeries.DataSources.YValues.Formula ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.NumberReference, scatterSeries.DataSources.YValues.ReferenceKindValue);
        TestAssert.Equal(PptxSceneChartDataSourceCacheKind.NumberCache, scatterSeries.DataSources.YValues.CacheKindValue);
        TestAssert.True(scatterSeries.DataSources.YValues.HasCachedPoints, "Expected y-value cache point presence in the scene model.");
        TestAssert.Equal("Sheet2!$C$2:$C$4", scatterSeries.DataSources.BubbleSizes.Formula ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.NumberReference, scatterSeries.DataSources.BubbleSizes.ReferenceKindValue);
        TestAssert.Equal(PptxSceneChartDataSourceCacheKind.NumberCache, scatterSeries.DataSources.BubbleSizes.CacheKindValue);
        TestAssert.True(!scatterSeries.DataSources.BubbleSizes.HasCachedPoints, "Expected empty bubble-size cache to remain explicit.");
    }

    public static void PptxSceneInspectionSummarizesChartDataSourceStructure()
    {
        PptxSceneNodeSnapshot snapshot = PptxTests.BuildSingleChartSceneSnapshot("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:ser>
                  <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache/></c:strRef></c:tx>
                  <c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$A$4</c:f><c:multiLvlStrCache/></c:multiLvlStrRef></c:cat>
                  <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:pt idx="0"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val>
                </c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        TestAssert.Equal(3, snapshot.ChartDataSourceFormulaCount);
        TestAssert.Equal(1, snapshot.ChartDataSourceCachedPointSourceCount);
        TestAssert.Equal("multiLvlStrRef,numRef,strRef", string.Join(",", snapshot.ChartDataSourceReferenceKinds));
        TestAssert.Equal("multiLvlStrCache,numCache,strCache", string.Join(",", snapshot.ChartDataSourceCacheKinds));
    }

    public static void PptxScenePreservesChartPlotNumericOptionTokens()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:barChart>
                  <c:gapWidth val="futureGap"/>
                  <c:overlap val="25"/>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:barChart>
                <c:doughnutChart>
                  <c:holeSize val="futureHole"/>
                  <c:firstSliceAng val="450"/>
                  <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:doughnutChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartPlot barPlot = chart?.Plots[0] ?? throw new InvalidOperationException("Expected bar plot.");
        TestAssert.Equal(null, barPlot.GapWidth);
        TestAssert.Equal("futureGap", barPlot.GapWidthValue);
        TestAssert.Equal(25d, barPlot.Overlap ?? double.NaN);
        TestAssert.Equal("25", barPlot.OverlapValue);

        PptxSceneChartPlot doughnutPlot = chart.Plots[1];
        TestAssert.Equal(null, doughnutPlot.HoleSize);
        TestAssert.Equal("futureHole", doughnutPlot.HoleSizeValue);
        TestAssert.Equal(450d, doughnutPlot.FirstSliceAngle ?? double.NaN);
        TestAssert.Equal("450", doughnutPlot.FirstSliceAngleValue);
    }

    public static void PptxScenePreservesChartExplosionNumericOptionTokens()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:pieChart>
                <c:ser>
                  <c:idx val="futureSeriesIndex"/>
                  <c:order val="3"/>
                  <c:explosion val="futureSeriesExplosion"/>
                  <c:dPt><c:idx val="-1"/><c:explosion val="80"/></c:dPt>
                  <c:dPt><c:idx val="0"/><c:explosion val="35"/></c:dPt>
                  <c:dPt><c:idx val="1"/><c:explosion val="futurePointExplosion"/></c:dPt>
                  <c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val>
                </c:ser>
              </c:pieChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartSeries series = chart?.Plots[0].Series[0] ?? throw new InvalidOperationException("Expected pie series.");
        TestAssert.Equal(null, series.Index);
        TestAssert.Equal("futureSeriesIndex", series.IndexValue);
        TestAssert.Equal(3, series.Order ?? 0);
        TestAssert.Equal("3", series.OrderValue);
        TestAssert.Equal(null, series.Explosion);
        TestAssert.Equal("futureSeriesExplosion", series.ExplosionValue);
        TestAssert.Equal(2, series.PointStyles.Count);
        TestAssert.Equal(35d, series.PointStyles[0].Explosion ?? double.NaN);
        TestAssert.Equal("0", series.PointStyles[0].IndexValue);
        TestAssert.Equal("35", series.PointStyles[0].ExplosionValue);
        TestAssert.Equal(null, series.PointStyles[1].Explosion);
        TestAssert.Equal("1", series.PointStyles[1].IndexValue);
        TestAssert.Equal("futurePointExplosion", series.PointStyles[1].ExplosionValue);
    }

    public static void PptxScenePreservesChartAxisNumericOptionTokens()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/><c:crossAx val="20"/></c:catAx>
                <c:valAx>
                  <c:axId val="20"/>
                  <c:scaling><c:min val="futureMin"/><c:max val="42.5"/></c:scaling>
                  <c:crossAx val="10"/>
                  <c:crossesAt val="futureCross"/>
                  <c:majorUnit val="futureMajor"/>
                  <c:minorUnit val="-1"/>
                  <c:lblOffset val="futureOffset"/>
                  <c:tickLblSkip val="2"/>
                  <c:tickMarkSkip val="futureTickMarkSkip"/>
                </c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartAxis axis = chart?.Axes.First(item => item.AxisKind == PptxSceneChartAxisKind.Value)
            ?? throw new InvalidOperationException("Expected value axis.");
        TestAssert.Equal(null, axis.CrossesAt);
        TestAssert.Equal("futureCross", axis.CrossesAtValue);
        TestAssert.Equal(null, axis.Minimum);
        TestAssert.Equal("futureMin", axis.MinimumValue);
        TestAssert.Equal(42.5d, axis.Maximum ?? double.NaN);
        TestAssert.Equal("42.5", axis.MaximumValue);
        TestAssert.Equal(null, axis.MajorUnit);
        TestAssert.Equal("futureMajor", axis.MajorUnitValue);
        TestAssert.Equal(null, axis.MinorUnit);
        TestAssert.Equal("-1", axis.MinorUnitValue);
        TestAssert.Equal(null, axis.LabelOffset);
        TestAssert.Equal("futureOffset", axis.LabelOffsetValue);
        TestAssert.Equal(2, axis.TickLabelSkip ?? 0);
        TestAssert.Equal("2", axis.TickLabelSkipValue);
        TestAssert.Equal(null, axis.TickMarkSkip);
        TestAssert.Equal("futureTickMarkSkip", axis.TickMarkSkipValue);
    }

    public static void PptxScenePreservesChartManualLayoutNumericOptionTokens()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:layout><c:manualLayout>
                  <c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/>
                  <c:x val="futureX"/><c:y val="0.2"/><c:w val="futureW"/><c:h val="0.4"/>
                </c:manualLayout></c:layout>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartManualLayout layout = chart?.PlotAreaLayout ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.True(layout.HasLayout, "Expected plot-area manual layout to be preserved.");
        TestAssert.Equal(null, layout.X);
        TestAssert.Equal("futureX", layout.XValue);
        TestAssert.Equal(0.2d, layout.Y ?? double.NaN);
        TestAssert.Equal("0.2", layout.YValue);
        TestAssert.Equal(null, layout.Width);
        TestAssert.Equal("futureW", layout.WidthValue);
        TestAssert.Equal(0.4d, layout.Height ?? double.NaN);
        TestAssert.Equal("0.4", layout.HeightValue);
    }

    public static void PptxSceneAbsentChartManualLayoutHasUnknownDefaults()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """;

        PptxSceneChart? chart = PptxTests.BuildSingleChartScene(chartXml);
        PptxSceneChartManualLayout layout = chart?.PlotAreaLayout ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.True(!layout.HasLayout, "Expected absent plot-area manual layout to remain absent.");
        TestAssert.Equal(PptxSceneChartManualLayoutTarget.Unknown, layout.LayoutTargetKind);
        TestAssert.Equal(PptxSceneChartManualLayoutMode.Unknown, layout.XModeKind);
        TestAssert.Equal(PptxSceneChartManualLayoutMode.Unknown, layout.YModeKind);
        TestAssert.Equal(PptxSceneChartManualLayoutMode.Unknown, layout.WidthModeKind);
        TestAssert.Equal(PptxSceneChartManualLayoutMode.Unknown, layout.HeightModeKind);

        PptxSceneNodeSnapshot snapshot = PptxTests.BuildSingleChartSceneSnapshot(chartXml);
        TestAssert.True(!snapshot.HasChartPlotAreaManualLayout, "Expected scene inspection to keep absent plot-area manual layout distinct from a default inner target.");
        TestAssert.Equal(string.Empty, snapshot.ChartPlotAreaLayoutTarget);
        TestAssert.Equal("Unknown", snapshot.ChartPlotAreaLayoutTargetKind);
        TestAssert.Equal(string.Empty, snapshot.ChartPlotAreaLayoutXMode);
        TestAssert.Equal("Unknown", snapshot.ChartPlotAreaLayoutXModeKind);
        TestAssert.Equal(string.Empty, snapshot.ChartPlotAreaLayoutWidthMode);
        TestAssert.Equal("Unknown", snapshot.ChartPlotAreaLayoutWidthModeKind);
    }

    public static void PptxScenePreservesChartNumericPointIndicesAndBlanks()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:dispBlanksAs val="gap"/><c:plotArea>
                <c:lineChart><c:ser>
                  <c:cat><c:strLit>
                    <c:ptCount val="6"/>
                    <c:pt idx="0"><c:v>Alpha</c:v></c:pt>
                    <c:pt idx="2"><c:v></c:v></c:pt>
                    <c:pt idx="4"/>
                    <c:pt idx="5"><c:v>Omega</c:v></c:pt>
                  </c:strLit></c:cat>
                  <c:val><c:numLit>
                    <c:formatCode>0.00</c:formatCode>
                    <c:ptCount val="6"/>
                    <c:pt idx="0"><c:v>1.25</c:v></c:pt>
                    <c:pt idx="2"><c:v></c:v></c:pt>
                    <c:pt idx="3"/>
                    <c:pt idx="4"><c:v>missing</c:v></c:pt>
                    <c:pt idx="5"><c:v>3.5</c:v></c:pt>
                  </c:numLit></c:val>
                </c:ser></c:lineChart>
                <c:scatterChart><c:ser>
                  <c:xVal><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="8"/><c:pt idx="7"><c:v>9</c:v></c:pt></c:numLit></c:xVal>
                  <c:yVal><c:numLit><c:formatCode>0.0%</c:formatCode><c:ptCount val="8"/><c:pt idx="7"><c:v></c:v></c:pt></c:numLit></c:yVal>
                  <c:bubbleSize><c:numLit><c:formatCode>0</c:formatCode><c:ptCount val="8"/><c:pt idx="7"><c:v>16</c:v></c:pt></c:numLit></c:bubbleSize>
                </c:ser></c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartSeries lineSeries = chart?.Plots[0].Series[0] ?? throw new InvalidOperationException("Expected line chart series.");
        TestAssert.Equal(2, lineSeries.Categories.Count);
        TestAssert.Equal(4, lineSeries.CategoryPoints.Count);
        TestAssert.Equal(6, lineSeries.CategoryPointCount ?? 0);
        TestAssert.Equal("6", lineSeries.CategoryPointCountValue);
        TestAssert.Equal(0, lineSeries.CategoryPoints[0].Index);
        TestAssert.Equal("0", lineSeries.CategoryPoints[0].IndexValue);
        TestAssert.True(lineSeries.CategoryPoints[0].HasParsedIndex, "Expected valid category cache point index tokens to remain marked parsed.");
        TestAssert.Equal("Alpha", lineSeries.CategoryPoints[0].Text);
        TestAssert.True(lineSeries.CategoryPoints[0].HasText, "Expected category cache point text presence to be explicit.");
        TestAssert.Equal(2, lineSeries.CategoryPoints[1].Index);
        TestAssert.Equal("2", lineSeries.CategoryPoints[1].IndexValue);
        TestAssert.True(lineSeries.CategoryPoints[1].HasParsedIndex, "Expected sparse category cache point index tokens to remain marked parsed.");
        TestAssert.Equal(string.Empty, lineSeries.CategoryPoints[1].Text);
        TestAssert.True(lineSeries.CategoryPoints[1].HasText, "Expected blank category cache point to preserve its value element.");
        TestAssert.Equal(4, lineSeries.CategoryPoints[2].Index);
        TestAssert.Equal(string.Empty, lineSeries.CategoryPoints[2].Text);
        TestAssert.True(!lineSeries.CategoryPoints[2].HasText, "Expected missing category value element to remain distinguishable from an empty value.");
        TestAssert.Equal(5, lineSeries.CategoryPoints[3].Index);
        TestAssert.Equal("Omega", lineSeries.CategoryPoints[3].Text);
        TestAssert.Equal(2, lineSeries.Values.Count);
        TestAssert.Equal(5, lineSeries.ValuePoints.Count);
        TestAssert.Equal(6, lineSeries.ValuePointCount ?? 0);
        TestAssert.Equal("6", lineSeries.ValuePointCountValue);
        TestAssert.Equal("0.00", lineSeries.ValueFormatCode ?? string.Empty);
        TestAssert.Equal(0, lineSeries.ValuePoints[0].Index);
        TestAssert.Equal("0", lineSeries.ValuePoints[0].IndexValue);
        TestAssert.True(lineSeries.ValuePoints[0].HasParsedIndex, "Expected valid numeric cache point index tokens to remain marked parsed.");
        TestAssert.Equal(1.25d, lineSeries.ValuePoints[0].Value ?? 0d);
        TestAssert.Equal(2, lineSeries.ValuePoints[1].Index);
        TestAssert.Equal("2", lineSeries.ValuePoints[1].IndexValue);
        TestAssert.True(lineSeries.ValuePoints[1].HasParsedIndex, "Expected sparse numeric cache point index tokens to remain marked parsed.");
        TestAssert.True(lineSeries.ValuePoints[1].HasValueElement, "Expected blank numeric value element presence to remain explicit.");
        TestAssert.True(lineSeries.ValuePoints[1].Value is null, "Expected blank numeric point value to remain explicit.");
        TestAssert.Equal(string.Empty, lineSeries.ValuePoints[1].Text);
        TestAssert.Equal(3, lineSeries.ValuePoints[2].Index);
        TestAssert.True(!lineSeries.ValuePoints[2].HasValueElement, "Expected missing numeric value element to remain distinguishable from an empty value.");
        TestAssert.True(lineSeries.ValuePoints[2].Value is null, "Expected missing numeric point value to remain explicit.");
        TestAssert.Equal(string.Empty, lineSeries.ValuePoints[2].Text);
        TestAssert.Equal(4, lineSeries.ValuePoints[3].Index);
        TestAssert.True(lineSeries.ValuePoints[3].HasValueElement, "Expected non-numeric value element presence to remain explicit.");
        TestAssert.True(lineSeries.ValuePoints[3].Value is null, "Expected non-numeric point value to remain explicit.");
        TestAssert.Equal("missing", lineSeries.ValuePoints[3].Text);
        TestAssert.Equal(5, lineSeries.ValuePoints[4].Index);
        TestAssert.Equal(3.5d, lineSeries.ValuePoints[4].Value ?? 0d);

        PptxSceneChartSeries scatterSeries = chart?.Plots[1].Series[0] ?? throw new InvalidOperationException("Expected scatter chart series.");
        TestAssert.Equal(8, scatterSeries.XValuePointCount ?? 0);
        TestAssert.Equal("8", scatterSeries.XValuePointCountValue);
        TestAssert.Equal("General", scatterSeries.XValueFormatCode ?? string.Empty);
        TestAssert.Equal(7, scatterSeries.XValuePoints[0].Index);
        TestAssert.Equal(9d, scatterSeries.XValuePoints[0].Value ?? 0d);
        TestAssert.Equal(8, scatterSeries.YValuePointCount ?? 0);
        TestAssert.Equal("8", scatterSeries.YValuePointCountValue);
        TestAssert.Equal("0.0%", scatterSeries.YValueFormatCode ?? string.Empty);
        TestAssert.Equal(7, scatterSeries.YValuePoints[0].Index);
        TestAssert.True(scatterSeries.YValuePoints[0].Value is null, "Expected blank y-value point to remain explicit.");
        TestAssert.Equal(8, scatterSeries.BubbleSizePointCount ?? 0);
        TestAssert.Equal("8", scatterSeries.BubbleSizePointCountValue);
        TestAssert.Equal("0", scatterSeries.BubbleSizeFormatCode ?? string.Empty);
        TestAssert.Equal(7, scatterSeries.BubbleSizePoints[0].Index);
        TestAssert.Equal(16d, scatterSeries.BubbleSizePoints[0].Value ?? 0d);
    }

    public static void PptxSceneMarksFallbackChartPointIndices()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:lineChart><c:ser>
                <c:cat><c:strLit>
                  <c:pt idx="futureCategory"><c:v>Alpha</c:v></c:pt>
                  <c:pt><c:v>Beta</c:v></c:pt>
                  <c:pt idx="4"><c:v>Gamma</c:v></c:pt>
                </c:strLit></c:cat>
                <c:val><c:numLit>
                  <c:pt idx="futureValue"><c:v>1</c:v></c:pt>
                  <c:pt><c:v>2</c:v></c:pt>
                  <c:pt idx="4"><c:v>3</c:v></c:pt>
                </c:numLit></c:val>
              </c:ser></c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartSeries series = chart?.Plots[0].Series[0] ?? throw new InvalidOperationException("Expected chart series.");
        TestAssert.Equal(0, series.CategoryPoints[0].Index);
        TestAssert.Equal("futureCategory", series.CategoryPoints[0].IndexValue);
        TestAssert.True(!series.CategoryPoints[0].HasParsedIndex, "Expected malformed category cache point index to be marked as an ordinal fallback.");
        TestAssert.Equal(1, series.CategoryPoints[1].Index);
        TestAssert.Equal(string.Empty, series.CategoryPoints[1].IndexValue);
        TestAssert.True(!series.CategoryPoints[1].HasParsedIndex, "Expected missing category cache point index to be marked as an ordinal fallback.");
        TestAssert.Equal(4, series.CategoryPoints[2].Index);
        TestAssert.True(series.CategoryPoints[2].HasParsedIndex, "Expected valid category cache point index to be marked parsed.");

        TestAssert.Equal(0, series.ValuePoints[0].Index);
        TestAssert.Equal("futureValue", series.ValuePoints[0].IndexValue);
        TestAssert.True(!series.ValuePoints[0].HasParsedIndex, "Expected malformed numeric cache point index to be marked as an ordinal fallback.");
        TestAssert.Equal(1, series.ValuePoints[1].Index);
        TestAssert.Equal(string.Empty, series.ValuePoints[1].IndexValue);
        TestAssert.True(!series.ValuePoints[1].HasParsedIndex, "Expected missing numeric cache point index to be marked as an ordinal fallback.");
        TestAssert.Equal(4, series.ValuePoints[2].Index);
        TestAssert.True(series.ValuePoints[2].HasParsedIndex, "Expected valid numeric cache point index to be marked parsed.");
    }

    public static void PptxScenePreservesChartMultiLevelCategoryPoints()
    {
        string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea><c:lineChart><c:ser>
                <c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$4</c:f><c:multiLvlStrCache>
                  <c:ptCount val="3"/>
                  <c:lvl>
                    <c:pt idx="0"><c:v>FY26</c:v></c:pt>
                    <c:pt idx="2"><c:v>FY27</c:v></c:pt>
                  </c:lvl>
                  <c:lvl>
                    <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                    <c:pt idx="1"><c:v></c:v></c:pt>
                    <c:pt idx="2"><c:v>Q3</c:v></c:pt>
                  </c:lvl>
                </c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
              </c:ser></c:lineChart></c:plotArea></c:chart>
            </c:chartSpace>
            """;
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene(chartXml);

        PptxSceneChartSeries series = chart?.Plots[0].Series[0] ?? throw new InvalidOperationException("Expected chart series.");
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.MultiLevelStringReference, series.DataSources.Categories.ReferenceKindValue);
        TestAssert.True(series.DataSources.Categories.HasCachedPoints, "Expected multi-level category cache point presence to survive scene parsing.");
        TestAssert.Equal(3, series.CategoryPointCount ?? 0);
        TestAssert.Equal("3", series.CategoryPointCountValue);
        TestAssert.Equal(5, series.CategoryPoints.Count);
        TestAssert.Equal(2, series.CategoryLevels.Count);
        TestAssert.Equal(2, series.CategoryLevels[0].Count);
        TestAssert.Equal(0, series.CategoryLevels[0][0].Index);
        TestAssert.Equal("FY26", series.CategoryLevels[0][0].Text);
        TestAssert.Equal(2, series.CategoryLevels[0][1].Index);
        TestAssert.Equal("FY27", series.CategoryLevels[0][1].Text);
        TestAssert.Equal(3, series.CategoryLevels[1].Count);
        TestAssert.Equal(1, series.CategoryLevels[1][1].Index);
        TestAssert.Equal(string.Empty, series.CategoryLevels[1][1].Text);
        TestAssert.True(series.CategoryLevels[1][1].HasText, "Expected blank multi-level category value to preserve its value element.");

        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement chartElement = XDocument.Parse(chartXml).Descendants(chartNamespace + "lineChart").Single();
        System.Reflection.MethodInfo readCategoryLabelVector = typeof(PptxRenderer)
            .GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .Single(method => method.Name == "ReadChartCategoryLabelVector" && method.GetParameters().Length == 3);
        object rawVector = readCategoryLabelVector.Invoke(null, [chartElement, null, true]) ?? throw new InvalidOperationException("Expected raw category-label vector.");
        TestAssert.Equal("Sheet1!$A$2:$B$4", (string?)rawVector.GetType().GetProperty("Formula")?.GetValue(rawVector) ?? string.Empty);
        TestAssert.Equal(3, (int?)rawVector.GetType().GetProperty("PointCount")?.GetValue(rawVector) ?? 0);

        object rawSource = rawVector.GetType().GetProperty("Source")?.GetValue(rawVector) ?? throw new InvalidOperationException("Expected raw category source metadata.");
        TestAssert.Equal(PptxSceneChartDataSourceReferenceKind.MultiLevelStringReference, (PptxSceneChartDataSourceReferenceKind?)rawSource.GetType().GetProperty("ReferenceKindValue")?.GetValue(rawSource) ?? default);
        TestAssert.Equal(PptxSceneChartDataSourceCacheKind.MultiLevelStringCache, (PptxSceneChartDataSourceCacheKind?)rawSource.GetType().GetProperty("CacheKindValue")?.GetValue(rawSource) ?? default);
        TestAssert.True((bool?)rawSource.GetType().GetProperty("HasCachedPoints")?.GetValue(rawSource) == true, "Expected raw multi-level category cache point presence to survive fallback parsing.");

        object[] rawPoints = (((System.Collections.IEnumerable?)rawVector.GetType().GetProperty("Points")?.GetValue(rawVector)) ?? throw new InvalidOperationException("Expected raw category points.")).Cast<object>().ToArray();
        TestAssert.Equal(5, rawPoints.Length);
        object[] rawLevels = (((System.Collections.IEnumerable?)rawVector.GetType().GetProperty("Levels")?.GetValue(rawVector)) ?? throw new InvalidOperationException("Expected raw category levels.")).Cast<object>().ToArray();
        TestAssert.Equal(2, rawLevels.Length);
        object[] rawFirstLevel = (((System.Collections.IEnumerable?)rawLevels[0]) ?? throw new InvalidOperationException("Expected first raw category level.")).Cast<object>().ToArray();
        object[] rawSecondLevel = (((System.Collections.IEnumerable?)rawLevels[1]) ?? throw new InvalidOperationException("Expected second raw category level.")).Cast<object>().ToArray();
        TestAssert.Equal(2, rawFirstLevel.Length);
        TestAssert.Equal(3, rawSecondLevel.Length);
        TestAssert.Equal(string.Empty, (string?)rawSecondLevel[1].GetType().GetProperty("Text")?.GetValue(rawSecondLevel[1]) ?? string.Empty);
        TestAssert.True((bool?)rawSecondLevel[1].GetType().GetProperty("HasText")?.GetValue(rawSecondLevel[1]) == true, "Expected raw blank multi-level category value to preserve its value element.");
    }

    public static void PptxScenePreservesChartNumberFormatMetadata()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:dLbls>
                    <c:numFmt formatCode="#,##0.00" sourceLinked="0"/>
                    <c:dLbl>
                      <c:idx val="1"/>
                      <c:numFmt formatCode="0%" sourceLinked="1"/>
                    </c:dLbl>
                  </c:dLbls>
                  <c:ser>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                </c:lineChart>
                <c:catAx>
                  <c:axId val="10"/>
                  <c:numFmt formatCode="m/d/yy" sourceLinked="true"/>
                </c:catAx>
                <c:valAx>
                  <c:axId val="20"/>
                  <c:numFmt formatCode="0.0%" sourceLinked="0"/>
                </c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartDataLabels labels = chart?.Plots[0].DataLabels ?? throw new InvalidOperationException("Expected plot data labels.");
        TestAssert.Equal("#,##0.00", labels.NumberFormat);
        TestAssert.True(labels.NumberFormatInfo.IsDefined, "Expected plot data-label number format metadata to be explicit.");
        TestAssert.Equal("#,##0.00", labels.NumberFormatInfo.FormatCode);
        TestAssert.True(labels.NumberFormatInfo.SourceLinked == false, "Expected sourceLinked=false to survive data-label parsing.");
        TestAssert.Equal("0", labels.NumberFormatInfo.SourceLinkedValue);

        PptxSceneChartDataLabelOverride label = labels.Overrides[0];
        TestAssert.Equal(1, label.Index);
        TestAssert.Equal("1", label.IndexValue);
        TestAssert.Equal("0%", label.NumberFormat);
        TestAssert.True(label.NumberFormatInfo.IsDefined, "Expected per-label number format metadata to be explicit.");
        TestAssert.Equal("0%", label.NumberFormatInfo.FormatCode);
        TestAssert.True(label.NumberFormatInfo.SourceLinked == true, "Expected sourceLinked=true to survive data-label override parsing.");
        TestAssert.Equal("1", label.NumberFormatInfo.SourceLinkedValue);

        PptxSceneChartAxis categoryAxis = chart.Axes.First(axis => axis.Kind == "catAx");
        TestAssert.Equal("m/d/yy", categoryAxis.NumberFormat ?? string.Empty);
        TestAssert.True(categoryAxis.NumberFormatInfo.IsDefined, "Expected category-axis number format metadata to be explicit.");
        TestAssert.Equal("m/d/yy", categoryAxis.NumberFormatInfo.FormatCode);
        TestAssert.True(categoryAxis.NumberFormatInfo.SourceLinked == true, "Expected sourceLinked=true to survive category-axis parsing.");
        TestAssert.Equal("true", categoryAxis.NumberFormatInfo.SourceLinkedValue);

        PptxSceneChartAxis valueAxis = chart.Axes.First(axis => axis.Kind == "valAx");
        TestAssert.Equal("0.0%", valueAxis.NumberFormat ?? string.Empty);
        TestAssert.True(valueAxis.NumberFormatInfo.IsDefined, "Expected value-axis number format metadata to be explicit.");
        TestAssert.Equal("0.0%", valueAxis.NumberFormatInfo.FormatCode);
        TestAssert.True(valueAxis.NumberFormatInfo.SourceLinked == false, "Expected sourceLinked=false to survive value-axis parsing.");
        TestAssert.Equal("0", valueAxis.NumberFormatInfo.SourceLinkedValue);

        Type numberFormatType = typeof(PptxRenderer).GetNestedType(
            "ChartNumberFormat",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected renderer chart number-format bridge.");
        object typedNumberFormat = Activator.CreateInstance(numberFormatType, [true, "#,##0.00", false, "0"]) ?? throw new InvalidOperationException("Expected typed number format.");
        System.Reflection.MethodInfo formatDataLabelValue = typeof(PptxRenderer).GetMethod(
            "FormatChartDataLabelValue",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static,
            binder: null,
            types: [typeof(double), numberFormatType, typeof(string), typeof(bool)],
            modifiers: null) ?? throw new InvalidOperationException("Expected typed data-label formatter.");
        string typedFormatted = (string?)formatDataLabelValue.Invoke(null, [1234.5d, typedNumberFormat, string.Empty, false]) ?? string.Empty;
        TestAssert.Equal("1,234.50", typedFormatted);
        System.Reflection.MethodInfo formatDataLabelValueWithSourceFormat = typeof(PptxRenderer).GetMethod(
            "FormatChartDataLabelValue",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static,
            binder: null,
            types: [typeof(double), numberFormatType, typeof(string), typeof(string), typeof(bool)],
            modifiers: null) ?? throw new InvalidOperationException("Expected source-format data-label formatter.");
        object emptyNumberFormat = Activator.CreateInstance(numberFormatType, [false, string.Empty, null, string.Empty]) ?? throw new InvalidOperationException("Expected empty number format.");
        string sourceFormatted = (string?)formatDataLabelValueWithSourceFormat.Invoke(null, [1234.5d, emptyNumberFormat, string.Empty, "#,##0.0", false]) ?? string.Empty;
        TestAssert.Equal("1,234.5", sourceFormatted);
        string explicitFormatWins = (string?)formatDataLabelValueWithSourceFormat.Invoke(null, [1234.5d, typedNumberFormat, string.Empty, "0", false]) ?? string.Empty;
        TestAssert.Equal("1,234.50", explicitFormatWins);

        object sourceLinkedNumberFormat = Activator.CreateInstance(numberFormatType, [true, "0", true, "1"]) ?? throw new InvalidOperationException("Expected source-linked number format.");
        System.Reflection.MethodInfo resolveSourceLinkedFormat = typeof(PptxRenderer).GetMethod(
            "ResolveSourceLinkedChartNumberFormatCode",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static,
            binder: null,
            types: [numberFormatType, typeof(string), typeof(bool?)],
            modifiers: null) ?? throw new InvalidOperationException("Expected source-linked format resolver.");
        string linkedFormat = (string?)resolveSourceLinkedFormat.Invoke(null, [sourceLinkedNumberFormat, "#,##0.0", true]) ?? string.Empty;
        TestAssert.Equal("#,##0.0", linkedFormat);
        object? ignoredUnlinkedFormat = resolveSourceLinkedFormat.Invoke(null, [typedNumberFormat, "#,##0.0", true]);
        TestAssert.True(ignoredUnlinkedFormat is null, "Expected workbook number format to be ignored when the chart format is not source-linked.");
        string linkedDateFormat = (string?)resolveSourceLinkedFormat.Invoke(null, [sourceLinkedNumberFormat, "m/d/yy", true]) ?? string.Empty;
        TestAssert.Equal("m/d/yy", linkedDateFormat);
    }

    public static void PptxScenePreservesChartTitleAndLegendBooleanTokens()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>
                <c:autoTitleDeleted val="false"/>
                <c:title>
                  <c:tx><c:rich><a:p><a:r><a:t>Token Title</a:t></a:r></a:p></c:rich></c:tx>
                  <c:overlay val="true"/>
                </c:title>
                <c:plotArea>
                  <c:lineChart>
                    <c:ser>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                  </c:lineChart>
                </c:plotArea>
                <c:legend><c:overlay val="0"/><c:delete val="1"/></c:legend>
              </c:chart>
            </c:chartSpace>
            """);

        PptxSceneChartTitle title = chart?.Title ?? throw new InvalidOperationException("Expected chart title.");
        TestAssert.True(title.IsAutoDeleted == false, "Expected parsed auto-title-delete metadata to remain false.");
        TestAssert.Equal("false", title.IsAutoDeletedValue);
        TestAssert.True(title.Overlay == true, "Expected parsed chart title overlay metadata to remain true.");
        TestAssert.Equal("true", title.OverlayValue);

        PptxSceneChartLegend legend = chart.Legend;
        TestAssert.True(legend.Overlay == false, "Expected parsed chart legend overlay metadata to remain false.");
        TestAssert.Equal("0", legend.OverlayValue);
        TestAssert.True(legend.IsDeleted == true, "Expected parsed chart legend delete metadata to remain true.");
        TestAssert.Equal("1", legend.IsDeletedValue);
    }

    public static void PptxChartAutoTitleDeletedSuppressesSingleSeriesName()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
                <c:lineChart><c:ser><c:tx><c:v>Line Series</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        TestAssert.True(chart?.Title.Text is null, "Expected explicit autoTitleDeleted=true to suppress inferred chart titles.");
        TestAssert.True(chart?.Title.IsAutoGenerated == false, "Expected deleted auto title metadata not to be reported as an inferred title.");
        TestAssert.True(chart?.Title.IsAutoDeleted == true, "Expected explicit autoTitleDeleted=true metadata to be preserved.");
        TestAssert.Equal("1", chart?.Title.IsAutoDeletedValue ?? string.Empty);
    }

    public static void PptxChartAutoTitleDoesNotInventAmbiguousMultiSeriesTitle()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:autoTitleDeleted val="0"/><c:plotArea>
                <c:lineChart>
                  <c:ser><c:tx><c:v>North</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  <c:ser><c:tx><c:v>South</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        TestAssert.True(chart?.Title.Text is null, "Expected ambiguous multi-series charts to remain titleless until Office evidence defines a title rule.");
        TestAssert.True(chart?.Title.IsAutoGenerated == false, "Expected multi-series ambiguity not to be reported as an inferred title.");
        TestAssert.True(chart?.Title.IsAutoDeleted == false, "Expected explicit autoTitleDeleted=false metadata to be preserved.");
        TestAssert.Equal("0", chart?.Title.IsAutoDeletedValue ?? string.Empty);
    }

    public static void PptxChartAutoTitleFillsTextlessTitleElement()
    {
        PptxSceneChart? chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:title><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea>
                <c:barChart><c:barDir val="bar"/><c:ser><c:tx><c:v>Actual</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);

        TestAssert.Equal("Actual", chart?.Title.Text ?? string.Empty);
        TestAssert.True(chart?.Title.IsAutoGenerated == true, "Expected a textless title element with autoTitleDeleted=false to take the single-series auto title.");
        TestAssert.True(chart?.Title.IsAutoDeleted == false, "Expected explicit autoTitleDeleted=false metadata to be preserved.");
        TestAssert.Equal("0", chart?.Title.IsAutoDeletedValue ?? string.Empty);
    }

    public static void PptxSyntheticDefaultAxisTitleDefaultsToBoldFace()
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
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
                  <c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
                    <c:barChart><c:barDir val="bar"/><c:ser><c:tx><c:v>Series</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
                    <c:catAx><c:axId val="1"/><c:axPos val="l"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>Categories</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:crossAx val="2"/></c:catAx>
                    <c:valAx><c:axId val="2"/><c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>Values</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:crossAx val="1"/></c:valAx>
                  </c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Split("/FontFile2", StringSplitOptions.None).Length - 1 >= 2, "Expected unstyled default axis titles to request a distinct embedded bold face like Office.");
    }

    public static void PptxChartUnknownLegendPositionResolvesThroughExplicitDefault()
    {
        const string chartXml = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:plotArea>
                  <c:lineChart>
                    <c:ser><c:tx><c:v>Line Series</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  </c:lineChart>
                </c:plotArea>
                <c:legend><c:legendPos val="bogus"/></c:legend>
              </c:chart>
            </c:chartSpace>
            """;
        PptxSceneChart chart = PptxTests.BuildSingleChartScene(chartXml) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(PptxSceneChartLegendPosition.Unknown, chart.Legend.PositionKind);
        TestAssert.Equal("bogus", chart.Legend.Position);

        System.Reflection.MethodInfo readLegendLayout = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartLegendLayout",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer legend-layout bridge.");
        object layout = readLegendLayout.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, chart, XDocument.Parse(chartXml)]) ?? throw new InvalidOperationException("Expected chart legend layout.");
        object positionKind = layout.GetType().GetProperty("PositionKind")?.GetValue(layout) ?? throw new InvalidOperationException("Expected legend position kind.");

        TestAssert.Equal(PptxSceneChartLegendPosition.Right, (PptxSceneChartLegendPosition)positionKind);
    }

    public static void PptxChartRightLegendReserveUsesLegendTextStyle()
    {
        const string chartXmlTemplate = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>
                <c:plotArea>
                  <c:lineChart>
                    <c:ser>
                      <c:tx><c:v>Long Legend Series Name</c:v></c:tx>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                  </c:lineChart>
                </c:plotArea>
                <c:legend>
                  <c:legendPos val="r"/>
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{0}"/></a:pPr></a:p></c:txPr>
                </c:legend>
              </c:chart>
            </c:chartSpace>
            """;
        XDocument smallLegendChart = XDocument.Parse(string.Format(CultureInfo.InvariantCulture, chartXmlTemplate, "900"));
        XDocument largeLegendChart = XDocument.Parse(string.Format(CultureInfo.InvariantCulture, chartXmlTemplate, "1800"));

        System.Reflection.MethodInfo readLegendTextStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartLegendTextStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static,
            binder: null,
            [typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument)],
            modifiers: null) ?? throw new InvalidOperationException("Expected chart legend text-style bridge.");
        object smallLegendStyle = readLegendTextStyle.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, null, smallLegendChart]) ?? throw new InvalidOperationException("Expected small legend style.");
        object largeLegendStyle = readLegendTextStyle.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, null, largeLegendChart]) ?? throw new InvalidOperationException("Expected large legend style.");

        Type frameType = typeof(PptxRenderer).GetNestedType(
            "ChartFrameBox",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart frame box.");
        object frame = Activator.CreateInstance(frameType, [0d, 0d, 600d, 400d]) ?? throw new InvalidOperationException("Expected chart frame.");
        System.Reflection.MethodInfo getPlotBox = typeof(PptxRenderer).GetMethod(
            "GetCartesianNoTitleRightLegendPlotBox",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected right-legend plot-box resolver.");

        object smallPlotBox = getPlotBox.Invoke(null, [frame, PptxTheme.Empty, smallLegendChart, null, null, true, null, smallLegendStyle]) ?? throw new InvalidOperationException("Expected small legend plot box.");
        object largePlotBox = getPlotBox.Invoke(null, [frame, PptxTheme.Empty, largeLegendChart, null, null, true, null, largeLegendStyle]) ?? throw new InvalidOperationException("Expected large legend plot box.");
        double smallWidth = (double)(smallPlotBox.GetType().GetProperty("Width")?.GetValue(smallPlotBox) ?? 0d);
        double largeWidth = (double)(largePlotBox.GetType().GetProperty("Width")?.GetValue(largePlotBox) ?? 0d);

        TestAssert.True(largeWidth < smallWidth, "Expected larger legend text style to reserve more right-legend width.");
    }

    public static void PptxSyntheticNoTitleRightLegendPlotBoxSitsOnOfficeAxisLine()
    {
        XDocument chartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser>
                    <c:tx><c:v>Series</c:v></c:tx>
                    <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart>
                <c:catAx><c:axId val="10"/></c:catAx>
                <c:valAx><c:axId val="20"/></c:valAx>
              </c:plotArea>
              <c:legend><c:legendPos val="r"/></c:legend>
              </c:chart>
            </c:chartSpace>
            """);

        System.Reflection.MethodInfo readLegendTextStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartLegendTextStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static,
            binder: null,
            [typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument)],
            modifiers: null) ?? throw new InvalidOperationException("Expected chart legend text-style bridge.");
        object legendStyle = readLegendTextStyle.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, null, chartXml]) ?? throw new InvalidOperationException("Expected legend style.");
        Type frameType = typeof(PptxRenderer).GetNestedType(
            "ChartFrameBox",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart frame box.");
        object frame = Activator.CreateInstance(frameType, [72d, 36d, 720d, 432d]) ?? throw new InvalidOperationException("Expected chart frame.");
        System.Reflection.MethodInfo getPlotBox = typeof(PptxRenderer).GetMethod(
            "GetCartesianNoTitleRightLegendPlotBox",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected right-legend plot-box resolver.");
        object plotBox = getPlotBox.Invoke(null, [frame, PptxTheme.Empty, chartXml, null, null, true, null, legendStyle]) ?? throw new InvalidOperationException("Expected plot box.");
        double y = (double)(plotBox.GetType().GetProperty("Y")?.GetValue(plotBox) ?? 0d);
        double height = (double)(plotBox.GetType().GetProperty("Height")?.GetValue(plotBox) ?? 0d);
        TestAssert.True(Math.Abs(y - 75.90d) < 0.05d, "Plot bottom should sit 39.9pt below the frame top, on the Office category-axis line. Got " + y);
        TestAssert.True(Math.Abs(y + height - 452.02d) < 0.05d, "Plot top should stay fixed. Got " + (y + height));
    }

    public static void PptxChartRightLegendValueAxisReserveUsesAxisTextStyle()
    {
        const string chartXmlTemplate = """
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>
                <c:plotArea>
                  <c:lineChart>
                    <c:axId val="10"/><c:axId val="20"/>
                    <c:ser>
                      <c:tx><c:v>Series</c:v></c:tx>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>1000</c:v></c:pt><c:pt idx="1"><c:v>2000</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                  </c:lineChart>
                  <c:catAx><c:axId val="10"/></c:catAx>
                  <c:valAx><c:axId val="20"/><c:scaling><c:min val="0"/><c:max val="2000"/></c:scaling><c:majorUnit val="1000"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{0}"/></a:pPr></a:p></c:txPr></c:valAx>
                </c:plotArea>
                <c:legend><c:legendPos val="r"/></c:legend>
              </c:chart>
            </c:chartSpace>
            """;
        XDocument smallAxisChart = XDocument.Parse(string.Format(CultureInfo.InvariantCulture, chartXmlTemplate, "850"));
        XDocument largeAxisChart = XDocument.Parse(string.Format(CultureInfo.InvariantCulture, chartXmlTemplate, "1800"));

        System.Reflection.MethodInfo readLegendTextStyle = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartLegendTextStyle",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static,
            binder: null,
            [typeof(PptxTheme), typeof(PptxColorMap), typeof(PptxSceneChart), typeof(XDocument)],
            modifiers: null) ?? throw new InvalidOperationException("Expected chart legend text-style bridge.");
        object legendStyle = readLegendTextStyle.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, null, smallAxisChart]) ?? throw new InvalidOperationException("Expected legend style.");
        Type frameType = typeof(PptxRenderer).GetNestedType(
            "ChartFrameBox",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart frame box.");
        object frame = Activator.CreateInstance(frameType, [0d, 0d, 600d, 400d]) ?? throw new InvalidOperationException("Expected chart frame.");
        System.Reflection.MethodInfo getPlotBox = typeof(PptxRenderer).GetMethod(
            "GetCartesianNoTitleRightLegendPlotBox",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected right-legend plot-box resolver.");

        object smallPlotBox = getPlotBox.Invoke(null, [frame, PptxTheme.Empty, smallAxisChart, null, null, true, null, legendStyle]) ?? throw new InvalidOperationException("Expected small axis plot box.");
        object largePlotBox = getPlotBox.Invoke(null, [frame, PptxTheme.Empty, largeAxisChart, null, null, true, null, legendStyle]) ?? throw new InvalidOperationException("Expected large axis plot box.");
        double smallX = (double)(smallPlotBox.GetType().GetProperty("X")?.GetValue(smallPlotBox) ?? 0d);
        double largeX = (double)(largePlotBox.GetType().GetProperty("X")?.GetValue(largePlotBox) ?? 0d);

        TestAssert.True(largeX > smallX, "Expected larger value-axis text style to reserve more left-side plot space.");
    }

    public static void PptxChartHorizontalLegendEntryWidthIncludesOfficePadding()
    {
        Type textStyleType = typeof(PptxRenderer).GetNestedType(
            "ChartTextStyle",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart text style.");
        Type textMeasurerType = typeof(PptxRenderer).GetNestedType(
            "ChartTextMeasurer",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart text measurer.");
        object style = Activator.CreateInstance(textStyleType, ["Arial", 9d, 0d, new RgbColor(0, 0, 0), 1d, false, false, false, false, null, null]) ?? throw new InvalidOperationException("Expected chart text style.");
        object textMeasurer = Activator.CreateInstance(textMeasurerType, [null]) ?? throw new InvalidOperationException("Expected chart text measurer.");
        System.Reflection.MethodInfo measure = textMeasurerType.GetMethod(
            "Measure",
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
            binder: null,
            [typeof(string), textStyleType],
            modifiers: null) ?? throw new InvalidOperationException("Expected chart text measurement method.");
        System.Reflection.MethodInfo entryWidth = typeof(PptxRenderer).GetMethod(
            "GetPackedHorizontalLegendEntryWidth",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected horizontal legend entry width helper.");

        string label = "Series";
        double markerSize = 4.95d;
        double measuredText = (double)(measure.Invoke(textMeasurer, [label, style]) ?? 0d);
        double width = (double)(entryWidth.Invoke(null, [label, style, textMeasurer, markerSize]) ?? 0d);

        double markerTextGap = 3d;
        double officePackedLegendEntryPadding = 8d;
        TestAssert.Equal(Math.Round(markerTextGap + officePackedLegendEntryPadding, 2), Math.Round(width - markerSize - measuredText, 2));
    }

    public static void PptxChartRadarValueAxisLabelFrameMeasuresTextWidth()
    {
        Type plotBoxType = typeof(PptxRenderer).GetNestedType(
            "ChartPlotBox",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart plot box.");
        Type geometryType = typeof(PptxRenderer).GetNestedType(
            "ChartPolarGeometry",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart polar geometry.");
        Type radarStyleType = typeof(PptxRenderer).GetNestedType(
            "ChartRadarStyle",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart radar style.");
        Type labelRulesType = typeof(PptxRenderer).GetNestedType(
            "ChartRadarLabelRules",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart radar label rules.");
        Type layoutType = typeof(PptxRenderer).GetNestedType(
            "ChartRadarLayout",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart radar layout.");
        Type textStyleType = typeof(PptxRenderer).GetNestedType(
            "ChartTextStyle",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart text style.");
        Type textMeasurerType = typeof(PptxRenderer).GetNestedType(
            "ChartTextMeasurer",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart text measurer.");

        object plotBox = Activator.CreateInstance(plotBoxType, [0d, 0d, 300d, 200d]) ?? throw new InvalidOperationException("Expected chart plot box.");
        object geometry = Activator.CreateInstance(geometryType, [150d, 100d, 80d]) ?? throw new InvalidOperationException("Expected chart geometry.");
        object radarStyle = Enum.Parse(radarStyleType, "Marker");
        object labelRules = Activator.CreateInstance(labelRulesType, [0.0202d, 0.0202d, 0.28d, -0.3138d, -0.012d, 0.3951d, 1.01d, 0.255d, 3.0d]) ?? throw new InvalidOperationException("Expected radar label rules.");
        object layout = Activator.CreateInstance(layoutType, [plotBox, geometry, radarStyle, 4, labelRules]) ?? throw new InvalidOperationException("Expected radar layout.");
        object style = Activator.CreateInstance(textStyleType, ["Arial", 8.5d, 0d, new RgbColor(0, 0, 0), 1d, false, false, false, false, null, null]) ?? throw new InvalidOperationException("Expected chart text style.");
        object textMeasurer = Activator.CreateInstance(textMeasurerType, [null]) ?? throw new InvalidOperationException("Expected chart text measurer.");
        System.Reflection.MethodInfo resolveFrame = typeof(PptxRenderer).GetMethod(
            "ResolveRadarValueAxisLabelFrame",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected radar value-axis label frame resolver.");

        object shortFrame = resolveFrame.Invoke(null, [layout, "1", style, textMeasurer, 0.5d]) ?? throw new InvalidOperationException("Expected short radar frame.");
        object longFrame = resolveFrame.Invoke(null, [layout, "1000000", style, textMeasurer, 0.5d]) ?? throw new InvalidOperationException("Expected long radar frame.");
        double shortWidth = (double)(shortFrame.GetType().GetProperty("Width")?.GetValue(shortFrame) ?? 0d);
        double longWidth = (double)(longFrame.GetType().GetProperty("Width")?.GetValue(longFrame) ?? 0d);

        TestAssert.True(longWidth > shortWidth, "Expected radar value-axis label frames to expand for measured label text.");
    }

    public static void PptxSyntheticBottomLegendBaselineRidesFrame()
    {
        // Office bottom legends anchor to the frame, not the plot: baseline at
        // frame bottom plus 8.6 plus 0.35 times the legend font size.
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
                  <c:chart><c:plotArea><c:barChart><c:barDir val="col"/>
                    <c:ser><c:tx><c:v>S1</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:ser><c:tx><c:v>S2</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
                  </c:barChart></c:plotArea>
                  <c:legend><c:legendPos val="b"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200"><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:legend></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(input, output);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ 264\.8 Tm").Count == 2, "Expected both bottom-legend entries at the frame-anchored baseline.");
    }

    public static void PptxDefaultBottomAxisTitleBaselineFollowsReserve()
    {
        // Office bottom default axis titles sit at frame bottom plus 0.338 times the
        // plot reserve (108 plus 46.77 times 0.338 lands 123.81 against Office 123.8 on
        // both axis-title probes); the legacy 0.23 lands 118.76 and fails this pin.
        System.Reflection.MethodInfo compute = typeof(PptxRenderer).GetMethod(
            "ComputeDefaultBottomAxisTitleBaselineY",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected bottom axis-title baseline helper.");
        double baseline = (double)(compute.Invoke(null, [108d, 154.77d]) ?? 0d);
        TestAssert.True(Math.Abs(baseline - 123.80826d) < 1e-9, "Bottom axis-title baseline drifts from the Office reserve law.");
        double floored = (double)(compute.Invoke(null, [108d, 100d]) ?? 0d);
        TestAssert.True(Math.Abs(floored - 108d) < 1e-9, "Bottom axis-title baseline must floor at the frame bottom.");
    }

    public static void PptxVerticalBarPlotBoxTopFloorBindsShortFrames()
    {
        // Office keeps untitled legendless column tops at least 10.9pt below the frame
        // (exact on 170H/288H/130H renders); taller frames keep the ratio top.
        Type frameType = typeof(PptxRenderer).GetNestedType("ChartFrameBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart frame box.");
        Type plotBoxType = typeof(PptxRenderer).GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart plot box.");
        System.Reflection.MethodInfo adjust = typeof(PptxRenderer).GetMethod("AdjustVerticalBarPlotBoxTopFloor", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected column top-floor adjuster.");
        object frame = Activator.CreateInstance(frameType, [430d, 245d, 255d, 130d]) ?? throw new InvalidOperationException("Expected chart frame.");
        object plot = Activator.CreateInstance(plotBoxType, [468.6d, 264.7d, 212.5d, 105.5d]) ?? throw new InvalidOperationException("Expected chart plot box.");
        object floored = adjust.Invoke(null, [plot, frame, false, false, false, true]) ?? throw new InvalidOperationException("Expected floored plot box.");
        Type boxType = floored.GetType();
        double top = (double)(boxType.GetProperty("Y")?.GetValue(floored) ?? 0d) + (double)(boxType.GetProperty("Height")?.GetValue(floored) ?? 0d);
        TestAssert.True(Math.Abs(top - 364.1d) < 1e-9, "Column top floor drifts from the Office 10.9pt clearance.");
        object tallPlot = Activator.CreateInstance(plotBoxType, [122.5d, 111.9d, 659.2d, 376.1d]) ?? throw new InvalidOperationException("Expected tall plot box.");
        object tallFrame = Activator.CreateInstance(frameType, [72d, 72d, 720d, 432d]) ?? throw new InvalidOperationException("Expected tall chart frame.");
        object kept = adjust.Invoke(null, [tallPlot, tallFrame, false, false, false, true]) ?? throw new InvalidOperationException("Expected kept plot box.");
        double keptTop = (double)(boxType.GetProperty("Y")?.GetValue(kept) ?? 0d) + (double)(boxType.GetProperty("Height")?.GetValue(kept) ?? 0d);
        TestAssert.True(Math.Abs(keptTop - 488d) < 1e-9, "Column top floor must not touch ratio-driven tall tops.");
        object titled = adjust.Invoke(null, [plot, frame, false, true, false, true]) ?? throw new InvalidOperationException("Expected titled plot box.");
        double titledTop = (double)(boxType.GetProperty("Y")?.GetValue(titled) ?? 0d) + (double)(boxType.GetProperty("Height")?.GetValue(titled) ?? 0d);
        TestAssert.True(Math.Abs(titledTop - 370.2d) < 1e-9, "Column top floor must not touch titled plots.");
        object unlabeled = adjust.Invoke(null, [plot, frame, false, false, false, false]) ?? throw new InvalidOperationException("Expected unlabeled plot box.");
        double unlabeledTop = (double)(boxType.GetProperty("Y")?.GetValue(unlabeled) ?? 0d) + (double)(boxType.GetProperty("Height")?.GetValue(unlabeled) ?? 0d);
        TestAssert.True(Math.Abs(unlabeledTop - 370.2d) < 1e-9, "Column top floor must not touch label-less plots.");
    }

    public static void PptxVerticalBarPlotBoxRightReserveSetsOfficeEdge()
    {
        // Office column right edges keep at least 10.3pt inside the frame.
        Type frameType = typeof(PptxRenderer).GetNestedType("ChartFrameBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart frame box.");
        Type plotBoxType = typeof(PptxRenderer).GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart plot box.");
        Type layoutType = typeof(PptxRenderer).GetNestedType("ChartLegendLayout", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart legend layout.");
        object hiddenLegend = layoutType.GetProperty("Hidden", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)?.GetValue(null) ?? throw new InvalidOperationException("Expected hidden legend layout.");
        System.Reflection.MethodInfo adjust = typeof(PptxRenderer).GetMethod("AdjustVerticalBarPlotBoxRightReserve", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected column right-reserve adjuster.");
        object frame = Activator.CreateInstance(frameType, [72d, 72d, 720d, 432d]) ?? throw new InvalidOperationException("Expected chart frame.");
        object plot = Activator.CreateInstance(plotBoxType, [122.5d, 111.9d, 667.5d, 376.1d]) ?? throw new InvalidOperationException("Expected chart plot box.");
        object set = adjust.Invoke(null, [plot, frame, false, hiddenLegend, true]) ?? throw new InvalidOperationException("Expected set plot box.");
        Type boxType = set.GetType();
        double right = (double)(boxType.GetProperty("X")?.GetValue(set) ?? 0d) + (double)(boxType.GetProperty("Width")?.GetValue(set) ?? 0d);
        TestAssert.True(Math.Abs(right - 781.7d) < 1e-9, "Column right reserve drifts from the Office 10.3pt edge.");
        object widePlot = Activator.CreateInstance(plotBoxType, [122.5d, 111.9d, 667.5d, 376.1d]) ?? throw new InvalidOperationException("Expected wide plot box.");
        object bars = adjust.Invoke(null, [widePlot, frame, true, hiddenLegend, true]) ?? throw new InvalidOperationException("Expected bars plot box.");
        double barsRight = (double)(boxType.GetProperty("X")?.GetValue(bars) ?? 0d) + (double)(boxType.GetProperty("Width")?.GetValue(bars) ?? 0d);
        TestAssert.True(Math.Abs(barsRight - 790d) < 1e-9, "Column right reserve must not touch horizontal bars.");
    }
    public static void PptxChartRadarWebGeometryIsFrameLocked()
    {
        // Office radar webs are style-invariant frame-locked squares: 432H untitled
        // centers at (432, 288) with R 182.565, a single-line auto title shifts to
        // (432, 270.3) with R 164.865, and the 324H frame centers at (432, 342).
        Type frameType = typeof(PptxRenderer).GetNestedType(
            "ChartFrameBox",
            System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected chart frame box.");
        System.Reflection.MethodInfo compute = typeof(PptxRenderer).GetMethod(
            "ComputeRadarWebGeometry",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected radar web geometry helper.");
        object frame = Activator.CreateInstance(frameType, [144d, 72d, 576d, 432d]) ?? throw new InvalidOperationException("Expected chart frame.");
        AssertRadarWebGeometry(compute.Invoke(null, [frame, false]), 432d, 288d, 182.565d);
        AssertRadarWebGeometry(compute.Invoke(null, [frame, true]), 432d, 270.3d, 164.865d);
        object shortFrame = Activator.CreateInstance(frameType, [144d, 180d, 576d, 324d]) ?? throw new InvalidOperationException("Expected short chart frame.");
        AssertRadarWebGeometry(compute.Invoke(null, [shortFrame, false]), 432d, 342d, 128.565d);
    }

    public static void PptxChartRadarLabelRulesAreUnified()
    {
        // Office radar label gaps are style-invariant: side-proportional gaps plus
        // the sine-level baseline fit, shared by marker and filled charts.
        System.Reflection.MethodInfo resolve = typeof(PptxRenderer).GetMethod(
            "ResolveRadarLabelRules",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected radar label rule resolver.");
        object rules = resolve.Invoke(null, []) ?? throw new InvalidOperationException("Expected radar label rules.");
        Type rulesType = rules.GetType();
        double hSide = (double)(rulesType.GetProperty("CategoryHorizontalGapSideFactor")?.GetValue(rules) ?? 0d);
        double vSide = (double)(rulesType.GetProperty("CategoryVerticalGapSideFactor")?.GetValue(rules) ?? 0d);
        double vFont = (double)(rulesType.GetProperty("CategoryVerticalGapFontFactor")?.GetValue(rules) ?? 0d);
        double b0 = (double)(rulesType.GetProperty("CategoryBaselineBaseFactor")?.GetValue(rules) ?? 0d);
        double b1 = (double)(rulesType.GetProperty("CategoryBaselineSineFactor")?.GetValue(rules) ?? 0d);
        double b2 = (double)(rulesType.GetProperty("CategoryBaselineSineSquaredFactor")?.GetValue(rules) ?? 0d);
        double vOff = (double)(rulesType.GetProperty("ValueBaselineOffsetFactor")?.GetValue(rules) ?? 0d);
        TestAssert.True(Math.Abs(hSide - 0.0202d) < 1e-12, "Radar horizontal gap drifts from the Office side law.");
        TestAssert.True(Math.Abs(vSide - 0.0202d) < 1e-12, "Radar vertical gap drifts from the Office side law.");
        TestAssert.True(Math.Abs(vFont - 0.28d) < 1e-12, "Radar vertical font term drifts from the Office law.");
        TestAssert.True(Math.Abs(b0 - -0.3138d) < 1e-12, "Radar baseline base drifts from the Office sine fit.");
        TestAssert.True(Math.Abs(b1 - -0.012d) < 1e-12, "Radar baseline sine term drifts from the Office sine fit.");
        TestAssert.True(Math.Abs(b2 - 0.3951d) < 1e-12, "Radar baseline sine-squared term drifts from the Office sine fit.");
        TestAssert.True(Math.Abs(vOff - 0.255d) < 1e-12, "Radar value offset drifts from the Office micro-law.");
    }
    private static void AssertRadarWebGeometry(object? geometry, double centerX, double centerY, double radius)
    {
        TestAssert.True(geometry is not null, "Expected radar web geometry.");
        Type geometryType = geometry!.GetType();
        double actualX = (double)(geometryType.GetProperty("CenterX")?.GetValue(geometry) ?? 0d);
        double actualY = (double)(geometryType.GetProperty("CenterY")?.GetValue(geometry) ?? 0d);
        double actualR = (double)(geometryType.GetProperty("Radius")?.GetValue(geometry) ?? 0d);
        TestAssert.True(Math.Abs(actualX - centerX) < 1e-9, "Radar web center X drifts from the Office-calibrated frame lock.");
        TestAssert.True(Math.Abs(actualY - centerY) < 1e-9, "Radar web center Y drifts from the Office-calibrated frame lock.");
        TestAssert.True(Math.Abs(actualR - radius) < 1e-9, "Radar web radius drifts from the Office-calibrated frame lock.");
    }
    public static void PptxChartMissingLegendUsesSceneAuthoritativeHiddenLayout()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.True(!chart.Legend.IsDefined, "Expected absent scene legend to remain distinct from an explicit legend.");

        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea/><c:legend><c:legendPos val="b"/></c:legend></c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readLegendLayout = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartLegendLayout",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer legend-layout bridge.");
        object sceneLayout = readLegendLayout.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, chart, mismatchedChartXml]) ?? throw new InvalidOperationException("Expected scene chart legend layout.");
        object xmlOnlyLayout = readLegendLayout.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, null, mismatchedChartXml]) ?? throw new InvalidOperationException("Expected XML-only chart legend layout.");

        TestAssert.True((bool)(sceneLayout.GetType().GetProperty("Visible")?.GetValue(sceneLayout) ?? true) == false, "Expected missing scene legend not to be repaired from fallback XML.");
        TestAssert.True((bool)(xmlOnlyLayout.GetType().GetProperty("Visible")?.GetValue(xmlOnlyLayout) ?? false), "Expected XML-only legend layout to keep reading XML.");
        TestAssert.Equal(PptxSceneChartLegendPosition.Bottom, (PptxSceneChartLegendPosition)(xmlOnlyLayout.GetType().GetProperty("PositionKind")?.GetValue(xmlOnlyLayout) ?? default(PptxSceneChartLegendPosition)));
    }

    public static void PptxChartLegendLayoutPreservesTextBodyPropertiesAtRendererBoundary()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea>
              <c:legend><c:legendPos val="r"/><c:txPr><a:bodyPr rot="-5400000"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr></c:legend>
              </c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea/>
                <c:legend><c:legendPos val="b"/><c:txPr><a:bodyPr rot="1200000"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr></c:legend>
              </c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readLegendLayout = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartLegendLayout",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer legend-layout bridge.");

        object sceneLayout = readLegendLayout.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, chart, mismatchedChartXml]) ?? throw new InvalidOperationException("Expected scene chart legend layout.");
        object xmlOnlyLayout = readLegendLayout.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, null, mismatchedChartXml]) ?? throw new InvalidOperationException("Expected XML-only chart legend layout.");
        object sceneBody = PptxTests.ChartTextBodyProperties(sceneLayout);
        object xmlBody = PptxTests.ChartTextBodyProperties(xmlOnlyLayout);

        TestAssert.Equal(-90d, PptxTests.ChartTextBodyRotationDegrees(sceneBody) ?? 0d);
        TestAssert.Equal("-5400000", PptxTests.ChartTextBodyRotationValue(sceneBody));
        TestAssert.Equal(20d, PptxTests.ChartTextBodyRotationDegrees(xmlBody) ?? 0d);
        TestAssert.Equal("1200000", PptxTests.ChartTextBodyRotationValue(xmlBody));
    }

    public static void PptxChartTitlePreservesTextBodyPropertiesAtRendererBoundary()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>
                <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Scene title</a:t></a:r></a:p></c:rich></c:tx><c:txPr><a:bodyPr rot="5400000"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr></c:title>
                <c:plotArea>
                  <c:lineChart>
                    <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                  </c:lineChart>
                </c:plotArea>
              </c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");

        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>
                <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>XML title</a:t></a:r></a:p></c:rich></c:tx><c:txPr><a:bodyPr rot="-1200000"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr></c:title>
                <c:plotArea/>
              </c:chart>
            </c:chartSpace>
            """);
        System.Reflection.MethodInfo readTitleBody = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartTitleTextBodyProperties",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer title text-body bridge.");

        object sceneBody = readTitleBody.Invoke(null, [chart, mismatchedChartXml]) ?? throw new InvalidOperationException("Expected scene chart title body properties.");
        object xmlBody = readTitleBody.Invoke(null, [null, mismatchedChartXml]) ?? throw new InvalidOperationException("Expected XML-only chart title body properties.");

        TestAssert.Equal(90d, PptxTests.ChartTextBodyRotationDegrees(sceneBody) ?? 0d);
        TestAssert.Equal("5400000", PptxTests.ChartTextBodyRotationValue(sceneBody));
        TestAssert.Equal(-20d, PptxTests.ChartTextBodyRotationDegrees(xmlBody) ?? 0d);
        TestAssert.Equal("-1200000", PptxTests.ChartTextBodyRotationValue(xmlBody));
    }

    public static void PptxChartMissingSceneAxesDoNotFallBackToMismatchedXmlAxes()
    {
        PptxSceneChart chart = PptxTests.BuildSingleChartScene("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart>
                  <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                </c:lineChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """) ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal(0, chart.Axes.Count);

        XDocument mismatchedChartXml = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:plotArea>
                <c:lineChart><c:axId val="10"/><c:axId val="20"/></c:lineChart>
                <c:catAx><c:axId val="10"/></c:catAx>
                <c:valAx><c:axId val="20"/></c:valAx>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """);
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement mismatchedPlot = mismatchedChartXml.Descendants(chartNamespace + "lineChart").First();

        System.Reflection.MethodInfo readValueAxes = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartValueAxesForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer value-axis bridge.");
        System.Reflection.MethodInfo readCategoryAxis = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlChartCategoryAxisForPlot",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer category-axis bridge.");

        var sceneValueAxes = (System.Collections.ICollection)(readValueAxes.Invoke(null, [chart, chart.Plots[0], mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected scene value-axis collection."));
        var xmlOnlyValueAxes = (System.Collections.ICollection)(readValueAxes.Invoke(null, [null, null, mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected XML-only value-axis collection."));
        object sceneCategoryAxis = readCategoryAxis.Invoke(null, [chart, chart.Plots[0], mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected scene category-axis source.");
        object xmlOnlyCategoryAxis = readCategoryAxis.Invoke(null, [null, null, mismatchedChartXml, mismatchedPlot]) ?? throw new InvalidOperationException("Expected XML-only category-axis source.");

        TestAssert.Equal(0, sceneValueAxes.Count);
        TestAssert.Equal(1, xmlOnlyValueAxes.Count);
        TestAssert.True(sceneCategoryAxis.GetType().GetProperty("XmlAxis")?.GetValue(sceneCategoryAxis) is null, "Expected missing scene category axes not to be repaired from fallback XML.");
        TestAssert.True(xmlOnlyCategoryAxis.GetType().GetProperty("XmlAxis")?.GetValue(xmlOnlyCategoryAxis) is not null, "Expected XML-only category-axis lookup to keep reading XML.");
    }
    public static void PptxScatterStrokeLegendEntriesHideSuppressedLinesButKeepMarkers()
    {
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement scatterChart = XDocument.Parse("""
            <?xml version="1.0" encoding="UTF-8"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:plotArea>
                <c:scatterChart>
                  <c:ser>
                    <c:spPr><a:ln><a:noFill/></a:ln></c:spPr>
                    <c:marker><c:symbol val="diamond"/><c:size val="9"/></c:marker>
                    <c:xVal><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:xVal>
                    <c:yVal><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:yVal>
                  </c:ser>
                  <c:ser>
                    <c:xVal><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt></c:numLit></c:xVal>
                    <c:yVal><c:numLit><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:yVal>
                  </c:ser>
                </c:scatterChart>
              </c:plotArea></c:chart>
            </c:chartSpace>
            """).Descendants(chartNamespace + "scatterChart").Single();

        System.Reflection.MethodInfo readMarkerStyles = typeof(PptxRenderer).GetMethod(
            "ReadSceneOrXmlMarkerStyles",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer marker-style bridge.");
        var markerStyles = (System.Collections.IList)(readMarkerStyles.Invoke(null, [null, scatterChart, PptxTheme.Empty, PptxColorMap.Default]) ?? throw new InvalidOperationException("Expected scatter marker styles."));
        TestAssert.Equal(2, markerStyles.Count);

        System.Reflection.MethodInfo buildEntries = typeof(PptxRenderer).GetMethod(
            "BuildStrokeLegendEntries",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected renderer stroke-legend bridge.");
        Type strokeListType = typeof(List<>).MakeGenericType(buildEntries.GetParameters()[5].ParameterType.GetGenericArguments()[0]);
        object emptyStrokes = Activator.CreateInstance(strokeListType) ?? throw new InvalidOperationException("Expected stroke list.");

        object[] entries = (((System.Collections.IEnumerable?)buildEntries.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, null, null, scatterChart, emptyStrokes, markerStyles, false, null, new List<bool> { true, false }])) ?? throw new InvalidOperationException("Expected stroke legend entries.")).Cast<object>().ToArray();
        TestAssert.Equal(2, entries.Length);
        TestAssert.Equal("Series 1", (string?)entries[0].GetType().GetProperty("Name")?.GetValue(entries[0]) ?? string.Empty);
        TestAssert.True((bool?)entries[0].GetType().GetProperty("LineHidden")?.GetValue(entries[0]) == true, "Expected the explicit-noFill series to suppress its legend line sample (Office marker-only key).");
        TestAssert.True((bool?)entries[1].GetType().GetProperty("LineHidden")?.GetValue(entries[1]) == false, "Expected the default-line series to keep its legend line sample.");
        object? marker0 = entries[0].GetType().GetProperty("Marker")?.GetValue(entries[0]);
        object? marker1 = entries[1].GetType().GetProperty("Marker")?.GetValue(entries[1]);
        TestAssert.True(marker0 is not null, "Expected the suppressed-line entry to keep its marker (Office marker-only key).");
        TestAssert.True(marker1 is not null, "Expected marker styles to flow into stroke legend entries.");
        TestAssert.Equal("diamond", PptxTests.ChartMarkerStyleSymbol(marker0!));
        TestAssert.Equal(9d, PptxTests.ChartMarkerStyleSize(marker0!));

        object[] defaultEntries = (((System.Collections.IEnumerable?)buildEntries.Invoke(null, [PptxTheme.Empty, PptxColorMap.Default, null, null, scatterChart, emptyStrokes, markerStyles, false, null, null])) ?? throw new InvalidOperationException("Expected default stroke legend entries.")).Cast<object>().ToArray();
        TestAssert.True((bool?)defaultEntries[0].GetType().GetProperty("LineHidden")?.GetValue(defaultEntries[0]) == false, "Expected missing hidden-line info to keep line samples (bar/line behavior).");
        TestAssert.True((bool?)defaultEntries[1].GetType().GetProperty("LineHidden")?.GetValue(defaultEntries[1]) == false, "Expected missing hidden-line info to keep line samples (bar/line behavior).");
    }
}
