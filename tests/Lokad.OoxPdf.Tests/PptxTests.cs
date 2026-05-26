using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxTests
{
    public static void PptxSyntheticTwoSlidesProducesTwoPdfPages()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldSz cx="12192000" cy="6858000"/>
                  <p:sldIdLst>
                    <p:sldId id="256" r:id="rId1"/>
                    <p:sldId id="257" r:id="rId2"/>
                  </p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>",
            ["ppt/slides/slide2.xml"] = "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>"
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>", pdf);
        TestAssert.Contains("/MediaBox [0 0 960 540]", pdf);
    }

    public static void PptxSceneBuilderBuildsResolvedNodeLists()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                  <Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
                  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """,
            ["ppt/charts/_rels/chart1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">
                  <a:themeElements>
                    <a:clrScheme name="Test">
                      <a:dk1><a:srgbClr val="111111"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent6><a:srgbClr val="336699"/></a:accent6>
                    </a:clrScheme>
                    <a:fontScheme name="Test">
                      <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
                      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
                    </a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/charts/chart1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:style val="10"/>
                  <c:spPr><a:pattFill prst="pct25"><a:fgClr><a:srgbClr val="224466"/></a:fgClr><a:bgClr><a:srgbClr val="F1E2D3"/></a:bgClr></a:pattFill><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></c:spPr>
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1" i="1"><a:solidFill><a:srgbClr val="101112"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
                  <c:chart>
                  <c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Scene Chart</a:t></a:r></a:p></c:rich></c:tx><c:layout><c:manualLayout><c:layoutTarget val="outer"/><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0.08"/><c:y val="0.04"/><c:w val="0.55"/><c:h val="0.12"/></c:manualLayout></c:layout><c:overlay val="1"/><c:spPr><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="0F1E2D"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1300" b="1" i="0"><a:solidFill><a:srgbClr val="1122AA"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:title>
                  <c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0.12"/><c:y val="0.18"/><c:w val="0.72"/><c:h val="0.66"/></c:manualLayout></c:layout><c:spPr><a:noFill/><a:ln w="25400"><a:solidFill><a:srgbClr val="112244"><a:alpha val="60000"/></a:srgbClr></a:solidFill></a:ln></c:spPr><c:barChart>
                    <c:barDir val="bar"/>
                    <c:grouping val="stacked"/>
                    <c:varyColors val="false"/>
                    <c:gapWidth val="175"/>
                    <c:overlap val="25"/>
                    <c:dLbls><c:showVal/><c:showPercent val="false"/><c:showCatName/><c:showSerName val="0"/><c:showLeaderLines/><c:dLblPos val="outEnd"/><c:separator>; </c:separator><c:numFmt formatCode="#,##0.0"/><c:spPr><a:solidFill><a:srgbClr val="FFEACC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1" i="0"><a:solidFill><a:srgbClr val="0A0B0C"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr><c:dLbl><c:idx val="1"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>ZXQ</a:t></a:r></a:p></c:rich></c:tx><c:showVal val="0"/><c:showSerName/><c:dLblPos val="ctr"/><c:layout><c:manualLayout><c:x val="0.44"/><c:y val="0.28"/><c:w val="0.1"/><c:h val="0.08"/></c:manualLayout></c:layout><c:separator> / </c:separator><c:numFmt formatCode="0%"/><c:spPr><a:solidFill><a:srgbClr val="CCEEFF"/></a:solidFill></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="1"><a:solidFill><a:srgbClr val="334455"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbl></c:dLbls>
                    <c:ser>
                      <c:idx val="7"/>
                      <c:order val="3"/>
                      <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
                      <c:dLbls><c:showVal val="1"/><c:showCatName val="0"/><c:dLblPos val="t"/><c:separator> + </c:separator></c:dLbls>
                      <c:spPr><a:solidFill><a:srgbClr val="AA5500"><a:alpha val="70000"/></a:srgbClr></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="003366"/></a:solidFill></a:ln></c:spPr>
                      <c:marker><c:symbol val="diamond"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="00AA55"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="5500AA"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:ln></c:spPr></c:marker>
                      <c:explosion val="12"/>
                      <c:smooth val="1"/>
                      <c:dPt><c:idx val="1"/><c:explosion val="18"/><c:spPr><a:solidFill><a:srgbClr val="CC8844"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="224466"/></a:solidFill></a:ln></c:spPr></c:dPt>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>12.5</c:v></c:pt><c:pt idx="1"><c:v>14</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:axId val="10"/>
                    <c:axId val="20"/>
                  </c:barChart>
                  <c:bubbleChart>
                    <c:ser>
                      <c:idx val="8"/>
                      <c:order val="4"/>
                      <c:tx><c:strLit><c:pt idx="0"><c:v>Scatter</c:v></c:pt></c:strLit></c:tx>
                      <c:spPr><a:pattFill prst="pct25"><a:fgClr><a:srgbClr val="112233"/></a:fgClr><a:bgClr><a:srgbClr val="F0E0D0"/></a:bgClr></a:pattFill></c:spPr>
                      <c:xVal><c:numLit><c:pt idx="0"><c:v>1.5</c:v></c:pt><c:pt idx="1"><c:v>2.5</c:v></c:pt></c:numLit></c:xVal>
                      <c:yVal><c:numLit><c:pt idx="0"><c:v>3.5</c:v></c:pt><c:pt idx="1"><c:v>4.5</c:v></c:pt></c:numLit></c:yVal>
                      <c:bubbleSize><c:numLit><c:pt idx="0"><c:v>9</c:v></c:pt><c:pt idx="1"><c:v>16</c:v></c:pt></c:numLit></c:bubbleSize>
                    </c:ser>
                  </c:bubbleChart>
                  <c:lineChart>
                    <c:ser>
                      <c:idx val="9"/>
                      <c:order val="5"/>
                      <c:tx><c:strLit><c:pt idx="0"><c:v>Trend</c:v></c:pt></c:strLit></c:tx>
                      <c:smooth val="0"/>
                      <c:cat><c:strLit><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt></c:strLit></c:cat>
                      <c:val><c:numLit><c:pt idx="0"><c:v>8</c:v></c:pt><c:pt idx="1"><c:v>11</c:v></c:pt></c:numLit></c:val>
                    </c:ser>
                    <c:axId val="10"/>
                    <c:axId val="20"/>
                  </c:lineChart>
                  <c:catAx><c:axId val="10"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Categories</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"><a:solidFill><a:srgbClr val="224488"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:title><c:axPos val="b"/><c:crossAx val="20"/><c:crosses val="autoZero"/><c:majorTickMark val="out"/><c:minorTickMark val="in"/><c:lblOffset val="100"/><c:tickLblSkip val="2"/><c:tickMarkSkip val="3"/><c:noMultiLvlLbl val="1"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:catAx>
                  <c:valAx><c:axId val="20"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Value</a:t></a:r></a:p></c:rich></c:tx><c:layout><c:manualLayout><c:x val="0.02"/><c:y val="0.2"/><c:w val="0.1"/><c:h val="0.5"/></c:manualLayout></c:layout><c:spPr><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></c:spPr></c:title><c:axPos val="l"/><c:delete val="0"/><c:scaling><c:orientation val="maxMin"/><c:min val="0"/><c:max val="20"/></c:scaling><c:crossAx val="10"/><c:crosses val="max"/><c:crossesAt val="2.5"/><c:crossBetween val="between"/><c:majorUnit val="5"/><c:minorUnit val="1"/><c:majorGridlines><c:spPr><a:ln w="6350" cap="rnd"><a:solidFill><a:srgbClr val="8899AA"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/><a:bevel/></a:ln></c:spPr></c:majorGridlines><c:minorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:minorGridlines><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="336699"><a:alpha val="75000"/></a:srgbClr></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1" i="1"><a:solidFill><a:srgbClr val="654321"/></a:solidFill><a:latin typeface="Aptos"/></a:defRPr></a:pPr></a:p></c:txPr><c:tickLblPos val="high"/><c:numFmt formatCode="$#,##0"/></c:valAx>
                  </c:plotArea>
                  <c:legend><c:legendPos val="b"/><c:overlay val="1"/><c:layout><c:manualLayout><c:x val="0.7"/><c:y val="0.68"/><c:w val="0.2"/><c:h val="0.15"/></c:manualLayout></c:layout><c:spPr><a:solidFill><a:srgbClr val="EEF7FF"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="334455"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="700" b="0" i="1"><a:solidFill><a:srgbClr val="2211AA"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr></c:legend>
                  </c:chart>
                </c:chartSpace>
                """,
            ["ppt/charts/colors1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                               meth="cycle" id="10">
                  <a:srgbClr val="010203"/>
                </cs:colorStyle>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldSz cx="9144000" cy="6858000"/>
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:nvSpPr><p:cNvPr id="1" name="MasterBox"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr></p:sp>
                  </p:spTree></p:cSld>
                  <p:defaultTextStyle><a:lvl2pPr><a:defRPr strike="sng"/></a:lvl2pPr></p:defaultTextStyle>
                  <p:txStyles>
                    <p:bodyStyle><a:lvl2pPr algn="ctr"><a:defRPr sz="2800" b="1" spc="120"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl2pPr></p:bodyStyle>
                  </p:txStyles>
                </p:sldMaster>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:nvSpPr><p:cNvPr id="2" name="LayoutDecoration"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:custGeom><a:gdLst><a:gd name="xMid" fmla="*/ w 1 2"/></a:gdLst><a:pathLst><a:path w="21600" h="21600" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="xMid" y="21600"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom></p:spPr></p:sp>
                    <p:sp><p:nvSpPr><p:cNvPr id="3" name="Title Placeholder"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm></p:spPr></p:sp>
                    <p:sp><p:nvSpPr><p:cNvPr id="7" name="Body Placeholder"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl2pPr><a:defRPr sz="2600" i="1"/></a:lvl2pPr></a:lstStyle><a:p/></p:txBody></p:sp>
                  </p:spTree></p:cSld>
                </p:sldLayout>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld>
                    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="123456"><a:alpha val="80000"/></a:srgbClr></a:solidFill></p:bgPr></p:bg>
                    <p:spTree>
                    <p:sp><p:nvSpPr><p:cNvPr id="4" name="TextBox"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:solidFill><a:srgbClr val="CCDD11"><a:alpha val="75000"/></a:srgbClr></a:solidFill><a:effectLst><a:glow rad="91440"><a:srgbClr val="0000FF"><a:alpha val="25000"/></a:srgbClr></a:glow><a:outerShdw dist="91440" dir="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="1"/><a:r><a:rPr u="sng"><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr><a:t>Hello</a:t></a:r><a:br/><a:fld type="slidenum"><a:rPr sz="1200"/><a:t>1</a:t></a:fld><a:endParaRPr sz="1800"/></a:p></p:txBody></p:sp>
                    <p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"><a:alphaModFix amt="50000"/><a:grayscl/></a:blip><a:srcRect l="10000" t="20000" r="30000" b="40000"/><a:stretch><a:fillRect l="5000" r="10000"/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr></p:pic>
                    <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Table"/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="1828800"/><a:ext cx="1828800" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{93296810-A885-4BE3-A3E7-6D5BEEA58F35}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid><a:tr h="914400"><a:tc gridSpan="2"><a:txBody><a:bodyPr lIns="91440" tIns="45720"/><a:lstStyle/><a:p/></a:txBody><a:tcPr marL="182880" anchor="ctr"><a:solidFill><a:srgbClr val="445566"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:lnL w="50800"><a:solidFill><a:srgbClr val="778899"><a:alpha val="60000"/></a:srgbClr></a:solidFill></a:lnL></a:tcPr></a:tc><a:tc hMerge="1"/></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
                    <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="8" name="Connector"/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="0" y="2743200"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom><a:ln w="25400" cap="rnd"><a:solidFill><a:srgbClr val="336699"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/><a:bevel/><a:headEnd type="arrow" w="lg"/><a:tailEnd type="triangle" len="sm"/></a:ln></p:spPr></p:cxnSp>
                    <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="Chart"/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="1828800" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>
                    <p:grpSp>
                      <p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:nvPr/></p:nvGrpSpPr>
                      <p:grpSpPr><a:xfrm><a:off x="2743200" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
                      <p:sp><p:nvSpPr><p:cNvPr id="11" name="GroupedShape"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:pattFill prst="dkDnDiag"><a:fgClr><a:srgbClr val="2F856A"/></a:fgClr><a:bgClr><a:srgbClr val="EEEEEE"/></a:bgClr></a:pattFill><a:blipFill><a:blip r:embed="rIdShapeImage"/><a:srcRect l="5000" t="10000" r="15000" b="20000"/><a:stretch><a:fillRect l="2500" r="7500"/></a:stretch></a:blipFill></p:spPr></p:sp>
                    </p:grpSp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxScene scene = new PptxSceneBuilder().Build(document, package);
        PptxSceneSnapshot sceneSnapshot = PptxRenderer.InspectScene(document, package);

        TestAssert.Equal(1, scene.Slides.Count);
        TestAssert.Equal(1, sceneSnapshot.Slides.Count);
        PptxSceneSlide slide = scene.Slides[0];
        PptxSceneSlideSnapshot slideSnapshot = sceneSnapshot.Slides[0];
        TestAssert.True(slide.SlideBackground.HasFill, "Expected slide background fill in the scene model.");
        TestAssert.True(slideSnapshot.HasSlideBackground, "Expected scene inspection to expose slide background ownership.");
        TestAssert.Equal(new RgbColor(18, 52, 86), slide.SlideBackground.Color);
        TestAssert.Equal(0.8d, slide.SlideBackground.Alpha);
        TestAssert.Equal(1, slide.MasterNodes.Count);
        TestAssert.Equal(3, slide.LayoutNodes.Count);
        TestAssert.Equal(6, slide.SlideNodes.Count);
        TestAssert.Equal(1, slideSnapshot.MasterNodes.Count);
        TestAssert.Equal(3, slideSnapshot.LayoutNodes.Count);
        TestAssert.Equal(6, slideSnapshot.SlideNodes.Count);
        TestAssert.Equal("Shape", slideSnapshot.SlideNodes[0].Kind);
        TestAssert.True(slideSnapshot.SlideNodes[0].HasTextBody, "Expected private-safe scene inspection to expose text-body ownership without text content.");
        TestAssert.Equal(1, slideSnapshot.SlideNodes[0].TextParagraphCount);
        TestAssert.Equal(3, slideSnapshot.SlideNodes[0].TextRunCount);
        TestAssert.Equal("Table", slideSnapshot.SlideNodes[2].Kind);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[2].TableRowCount);
        TestAssert.Equal(2, slideSnapshot.SlideNodes[2].TableCellCount);
        TestAssert.Equal("Chart", slideSnapshot.SlideNodes[4].Kind);
        TestAssert.Equal(3, slideSnapshot.SlideNodes[4].ChartPlotCount);
        TestAssert.Equal(2, slideSnapshot.SlideNodes[4].ChartAxisCount);
        TestAssert.Equal("Group", slideSnapshot.SlideNodes[5].Kind);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[5].Children.Count);
        TestAssert.Equal(PptxSceneNodeKind.Shape, slide.SlideNodes[0].Kind);
        TestAssert.Equal("rect", slide.SlideNodes[0].Shape?.Preset ?? string.Empty);
        TestAssert.True(slide.SlideNodes[0].Shape?.HasCustomGeometry == false, "Expected preset geometry in the scene model.");
        TestAssert.True(slide.SlideNodes[0].Shape?.Fill.HasFill == true, "Expected solid fill in the scene model.");
        TestAssert.Equal(new RgbColor(204, 221, 17), slide.SlideNodes[0].Shape?.Fill.Color ?? default);
        TestAssert.Equal(0.75d, slide.SlideNodes[0].Shape?.Fill.Alpha ?? 0d);
        TestAssert.True(slide.SlideNodes[0].Shape?.Glow.HasGlow == true, "Expected glow in the scene model.");
        TestAssert.Equal(new RgbColor(0, 0, 255), slide.SlideNodes[0].Shape?.Glow.Color ?? default);
        TestAssert.Equal(0.25d, slide.SlideNodes[0].Shape?.Glow.Alpha ?? 0d);
        TestAssert.Equal(7.2d, slide.SlideNodes[0].Shape?.Glow.Radius ?? 0d);
        TestAssert.True(slide.SlideNodes[0].Shape?.OuterShadow.HasShadow == true, "Expected outer shadow in the scene model.");
        TestAssert.Equal(new RgbColor(0, 0, 0), slide.SlideNodes[0].Shape?.OuterShadow.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[0].Shape?.OuterShadow.Alpha ?? 0d);
        TestAssert.Equal(7.2d, slide.SlideNodes[0].Shape?.OuterShadow.OffsetX ?? 0d);
        TestAssert.Equal(PptxSceneNodeKind.Picture, slide.SlideNodes[1].Kind);
        TestAssert.Equal("rIdImage", slide.SlideNodes[1].Picture?.RelationshipId ?? string.Empty);
        TestAssert.Equal(0.1d, slide.SlideNodes[1].Picture?.Crop.Left ?? 0d);
        TestAssert.Equal(0.4d, slide.SlideNodes[1].Picture?.Crop.Bottom ?? 0d);
        TestAssert.Equal(0.05d, slide.SlideNodes[1].Picture?.Fill.Left ?? 0d);
        TestAssert.Equal(0.5d, slide.SlideNodes[1].Picture?.Alpha ?? 0d);
        TestAssert.Equal(PptxSceneImageRecolorKind.Grayscale, slide.SlideNodes[1].Picture?.Recolor.Kind ?? PptxSceneImageRecolorKind.None);
        TestAssert.Equal(PptxSceneNodeKind.Table, slide.SlideNodes[2].Kind);
        TestAssert.Equal(2, slide.SlideNodes[2].Table?.ColumnWidths.Count ?? 0);
        TestAssert.Equal(914400d, slide.SlideNodes[2].Table?.ColumnWidths[0] ?? 0d);
        TestAssert.Equal(1, slide.SlideNodes[2].Table?.RowHeights.Count ?? 0);
        TestAssert.Equal(914400d, slide.SlideNodes[2].Table?.RowHeights[0] ?? 0d);
        TestAssert.True(slide.SlideNodes[2].Table?.Style.IsSupported == true, "Expected built-in table style lookup in the scene model.");
        TestAssert.Equal("Medium-Style-2", slide.SlideNodes[2].Table?.Style.Name ?? string.Empty);
        TestAssert.Equal("accent6", slide.SlideNodes[2].Table?.Style.Accent ?? string.Empty);
        TestAssert.True(slide.SlideNodes[2].Table?.Style.FirstRow == true, "Expected table first-row flag in the scene model.");
        TestAssert.True(slide.SlideNodes[2].Table?.Style.BandRow == true, "Expected table band-row flag in the scene model.");
        TestAssert.True(slide.SlideNodes[2].Table?.Source is not null, "Expected table source element ownership in the scene model.");
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[0].StyleFill.HasFill == true, "Expected resolved table-style fill in the scene model.");
        TestAssert.Equal(new RgbColor(51, 102, 153), slide.SlideNodes[2].Table?.Rows[0].Cells[0].StyleFill.Color ?? default);
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[0].StyleText.Bold == true, "Expected resolved table-style text bold in the scene model.");
        TestAssert.Equal(new RgbColor(255, 255, 255), slide.SlideNodes[2].Table?.Rows[0].Cells[0].StyleText.Color ?? default);
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[0].TextBody is not null, "Expected table-cell text body ownership in the scene model.");
        TestAssert.Equal(2, slide.SlideNodes[2].Table?.Rows[0].Cells[0].ColumnSpan ?? 0);
        TestAssert.Equal(1, slide.SlideNodes[2].Table?.Rows[0].Cells[0].RowSpan ?? 0);
        TestAssert.Equal(14.4d, slide.SlideNodes[2].Table?.Rows[0].Cells[0].TextInsets.Left ?? 0d);
        TestAssert.Equal(PptxSceneTableCellVerticalAnchor.Middle, slide.SlideNodes[2].Table?.Rows[0].Cells[0].VerticalAnchor ?? PptxSceneTableCellVerticalAnchor.Top);
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[0].Fill.HasFill == true, "Expected direct table-cell fill in the scene model.");
        TestAssert.Equal(new RgbColor(68, 85, 102), slide.SlideNodes[2].Table?.Rows[0].Cells[0].Fill.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[2].Table?.Rows[0].Cells[0].Fill.Alpha ?? 0d);
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.IsSpecified == true, "Expected explicit table-cell border in the scene model.");
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.HasLine == true, "Expected resolved table-cell border line in the scene model.");
        TestAssert.Equal(new RgbColor(119, 136, 153), slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.Color ?? default);
        TestAssert.Equal(2d, slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.Width ?? 0d);
        TestAssert.Equal(0.6d, slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.Alpha ?? 0d);
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[1].IsMergedContinuation == true, "Expected merged-cell continuation in the scene model.");
        TestAssert.Equal(PptxSceneNodeKind.Connector, slide.SlideNodes[3].Kind);
        TestAssert.Equal("straightConnector1", slide.SlideNodes[3].Shape?.Preset ?? string.Empty);
        TestAssert.Equal(50000d, slide.SlideNodes[3].Shape?.PresetAdjustments["adj1"] ?? 0d);
        TestAssert.True(slide.SlideNodes[3].Shape?.Line.HasLine == true, "Expected connector line style in the scene model.");
        TestAssert.Equal(new RgbColor(51, 102, 153), slide.SlideNodes[3].Shape?.Line.Color ?? default);
        TestAssert.Equal(2d, slide.SlideNodes[3].Shape?.Line.Width ?? 0d);
        TestAssert.Equal(0.5d, slide.SlideNodes[3].Shape?.Line.Alpha ?? 0d);
        TestAssert.Equal(8d, slide.SlideNodes[3].Shape?.Line.DashPattern[0] ?? 0d);
        TestAssert.Equal(1, slide.SlideNodes[3].Shape?.Line.Cap ?? 0);
        TestAssert.Equal(2, slide.SlideNodes[3].Shape?.Line.Join ?? 0);
        TestAssert.Equal(PptxSceneLineEndKind.Arrow, slide.SlideNodes[3].Shape?.HeadEnd.Kind ?? PptxSceneLineEndKind.None);
        TestAssert.Equal(1.5d, slide.SlideNodes[3].Shape?.HeadEnd.WidthScale ?? 0d);
        TestAssert.Equal(PptxSceneLineEndKind.Triangle, slide.SlideNodes[3].Shape?.TailEnd.Kind ?? PptxSceneLineEndKind.None);
        TestAssert.Equal(0.5d, slide.SlideNodes[3].Shape?.TailEnd.LengthScale ?? 0d);
        TestAssert.Equal(PptxSceneNodeKind.Chart, slide.SlideNodes[4].Kind);
        TestAssert.Equal("rIdChart", slide.SlideNodes[4].Chart?.RelationshipId ?? string.Empty);
        TestAssert.Equal("/ppt/charts/chart1.xml", slide.SlideNodes[4].Chart?.TargetPartName ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.ChartXml is not null, "Expected chart part XML ownership in the scene model.");
        TestAssert.Equal(new RgbColor(1, 2, 3), slide.SlideNodes[4].Chart?.PaletteColors?[0] ?? default);
        TestAssert.Equal("10", slide.SlideNodes[4].Chart?.StyleId ?? string.Empty);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal(11d, slide.SlideNodes[4].Chart?.TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(16, 17, 18), slide.SlideNodes[4].Chart?.TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.TextStyle.Bold == true, "Expected chart-level default run bold style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.TextStyle.Italic == true, "Expected chart-level default run italic style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.PlotAreaLayout.HasLayout == true, "Expected chart plot-area manual layout in the scene model.");
        TestAssert.Equal(0.12d, slide.SlideNodes[4].Chart?.PlotAreaLayout.X ?? 0d);
        TestAssert.Equal(0.18d, slide.SlideNodes[4].Chart?.PlotAreaLayout.Y ?? 0d);
        TestAssert.Equal(0.72d, slide.SlideNodes[4].Chart?.PlotAreaLayout.Width ?? 0d);
        TestAssert.Equal(0.66d, slide.SlideNodes[4].Chart?.PlotAreaLayout.Height ?? 0d);
        TestAssert.Equal("inner", slide.SlideNodes[4].Chart?.PlotAreaLayout.LayoutTarget ?? string.Empty);
        TestAssert.Equal(PptxSceneChartManualLayoutTarget.Inner, slide.SlideNodes[4].Chart?.PlotAreaLayout.LayoutTargetKind);
        TestAssert.Equal("factor", slide.SlideNodes[4].Chart?.PlotAreaLayout.XMode ?? string.Empty);
        TestAssert.Equal(PptxSceneChartManualLayoutMode.Factor, slide.SlideNodes[4].Chart?.PlotAreaLayout.XModeKind);
        TestAssert.Equal("factor", slide.SlideNodes[4].Chart?.PlotAreaLayout.YMode ?? string.Empty);
        TestAssert.Equal(PptxSceneChartManualLayoutMode.Factor, slide.SlideNodes[4].Chart?.PlotAreaLayout.YModeKind);
        TestAssert.Equal("factor", slide.SlideNodes[4].Chart?.PlotAreaLayout.WidthMode ?? string.Empty);
        TestAssert.Equal(PptxSceneChartManualLayoutMode.Factor, slide.SlideNodes[4].Chart?.PlotAreaLayout.WidthModeKind);
        TestAssert.Equal("factor", slide.SlideNodes[4].Chart?.PlotAreaLayout.HeightMode ?? string.Empty);
        TestAssert.Equal(PptxSceneChartManualLayoutMode.Factor, slide.SlideNodes[4].Chart?.PlotAreaLayout.HeightModeKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.ChartAreaStyle.NoFill == false, "Expected chart area noFill to be absent in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.ChartAreaStyle.PatternFill.HasPattern == true, "Expected chart area pattern fill in the scene model.");
        TestAssert.Equal("pct25", slide.SlideNodes[4].Chart?.ChartAreaStyle.PatternFill.Preset ?? string.Empty);
        TestAssert.Equal(new RgbColor(34, 68, 102), slide.SlideNodes[4].Chart?.ChartAreaStyle.PatternFill.Foreground ?? default);
        TestAssert.Equal(new RgbColor(241, 226, 211), slide.SlideNodes[4].Chart?.ChartAreaStyle.PatternFill.Background ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.ChartAreaStyle.Line.HasLine == true, "Expected chart area line in the scene model.");
        TestAssert.Equal(new RgbColor(68, 85, 102), slide.SlideNodes[4].Chart?.ChartAreaStyle.Line.Color ?? default);
        TestAssert.Equal(1d, slide.SlideNodes[4].Chart?.ChartAreaStyle.Line.Width ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.PlotAreaStyle.NoFill == true, "Expected plot area explicit noFill in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.PlotAreaStyle.Fill.HasFill == false, "Expected plot area noFill to suppress solid fill in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.PlotAreaStyle.Line.HasLine == true, "Expected plot area line in the scene model.");
        TestAssert.Equal(new RgbColor(17, 34, 68), slide.SlideNodes[4].Chart?.PlotAreaStyle.Line.Color ?? default);
        TestAssert.Equal(2d, slide.SlideNodes[4].Chart?.PlotAreaStyle.Line.Width ?? 0d);
        TestAssert.Equal(0.6d, slide.SlideNodes[4].Chart?.PlotAreaStyle.Line.Alpha ?? 0d);
        TestAssert.Equal("barChart", slide.SlideNodes[4].Chart?.Plots[0].Kind ?? string.Empty);
        TestAssert.Equal(PptxSceneChartPlotKind.Bar, slide.SlideNodes[4].Chart?.Plots[0].PlotKind);
        TestAssert.Equal(0, slide.SlideNodes[4].Chart?.Plots[0].PlotAreaIndex ?? -1);
        TestAssert.Equal(0, slide.SlideNodes[4].Chart?.Plots[0].KindIndex ?? -1);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Plots[0].SeriesCount ?? 0);
        TestAssert.Equal("20", slide.SlideNodes[4].Chart?.Plots[0].AxisIds[1] ?? string.Empty);
        TestAssert.Equal("stacked", slide.SlideNodes[4].Chart?.Plots[0].Grouping ?? string.Empty);
        TestAssert.Equal(PptxSceneChartGrouping.Stacked, slide.SlideNodes[4].Chart?.Plots[0].GroupingKind);
        TestAssert.Equal("bar", slide.SlideNodes[4].Chart?.Plots[0].BarDirection ?? string.Empty);
        TestAssert.Equal(PptxSceneChartBarDirection.Bar, slide.SlideNodes[4].Chart?.Plots[0].BarDirectionKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].VaryColors == false, "Expected chart plot varyColors in the scene model.");
        TestAssert.Equal(175d, slide.SlideNodes[4].Chart?.Plots[0].GapWidth ?? 0d);
        TestAssert.Equal(25d, slide.SlideNodes[4].Chart?.Plots[0].Overlap ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowValue == true, "Expected chart data-label value flag in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowPercent == false, "Expected chart data-label percent flag in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowCategoryName == true, "Expected chart data-label category-name flag in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowSeriesName == false, "Expected chart data-label series-name flag in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowLeaderLines == true, "Expected chart data-label leader-lines flag in the scene model.");
        TestAssert.Equal("outEnd", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Position ?? string.Empty);
        TestAssert.Equal("; ", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Separator ?? string.Empty);
        TestAssert.Equal("#,##0.0", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.NumberFormat ?? string.Empty);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal(10d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(10, 11, 12), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.Bold == true, "Expected plot-level data-label bold style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.Italic == false, "Expected explicit plot-level data-label italic disable in the scene model.");
        TestAssert.Equal(new RgbColor(255, 234, 204), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShapeStyle.Fill.Color ?? default);
        TestAssert.Equal(new RgbColor(17, 34, 51), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShapeStyle.Line.Color ?? default);
        TestAssert.Equal(1d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShapeStyle.Line.Width ?? 0d);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides.Count ?? 0);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].Index ?? -1);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowValue == false, "Expected per-label show-value override in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowSeriesName == true, "Expected per-label show-series-name override in the scene model.");
        TestAssert.Equal("ZXQ", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomText ?? string.Empty);
        TestAssert.Equal("ctr", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].Position ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataLabelPosition.Center, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].PositionKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].Layout.HasLayout == true, "Expected per-label manual layout in the scene model.");
        TestAssert.Equal(0.44d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].Layout.X ?? 0d);
        TestAssert.Equal(0.28d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].Layout.Y ?? 0d);
        TestAssert.Equal(0.1d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].Layout.Width ?? 0d);
        TestAssert.Equal(0.08d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].Layout.Height ?? 0d);
        TestAssert.Equal(" / ", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].Separator ?? string.Empty);
        TestAssert.Equal("0%", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].NumberFormat ?? string.Empty);
        TestAssert.Equal("Calibri", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal(9d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(51, 68, 85), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].TextStyle.Bold == false, "Expected explicit per-label bold disable in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].TextStyle.Italic == true, "Expected per-label italic style in the scene model.");
        TestAssert.Equal(new RgbColor(204, 238, 255), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShapeStyle.Fill.Color ?? default);
        TestAssert.Equal(7, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Index ?? -1);
        TestAssert.Equal(3, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Order ?? -1);
        TestAssert.Equal("Revenue", slide.SlideNodes[4].Chart?.Plots[0].Series[0].Name ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].DataLabels.IsDefined == true, "Expected series-level data-label options to be preserved separately from plot-level labels.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].DataLabels.ShowValue == true, "Expected series-level show-value flag in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].DataLabels.ShowPercent is null, "Expected missing series-level show-percent flag to remain distinct from an explicit disable.");
        TestAssert.Equal("t", slide.SlideNodes[4].Chart?.Plots[0].Series[0].DataLabels.Position ?? string.Empty);
        TestAssert.Equal(PptxSceneChartDataLabelPosition.Top, slide.SlideNodes[4].Chart?.Plots[0].Series[0].DataLabels.PositionKind);
        TestAssert.Equal(" + ", slide.SlideNodes[4].Chart?.Plots[0].Series[0].DataLabels.Separator ?? string.Empty);
        TestAssert.Equal(12.5d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Values[0] ?? 0d);
        TestAssert.Equal("South", slide.SlideNodes[4].Chart?.Plots[0].Series[0].Categories[1] ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].Fill.HasFill == true, "Expected chart series fill in the scene model.");
        TestAssert.Equal(new RgbColor(170, 85, 0), slide.SlideNodes[4].Chart?.Plots[0].Series[0].Fill.Color ?? default);
        TestAssert.Equal(0.7d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Fill.Alpha ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].Line.HasLine == true, "Expected chart series line in the scene model.");
        TestAssert.Equal(new RgbColor(0, 51, 102), slide.SlideNodes[4].Chart?.Plots[0].Series[0].Line.Color ?? default);
        TestAssert.Equal(3d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Line.Width ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.IsDefined == true, "Expected explicit chart marker ownership in the scene model.");
        TestAssert.Equal("diamond", slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Symbol ?? string.Empty);
        TestAssert.Equal(PptxSceneChartMarkerSymbol.Diamond, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.SymbolKind);
        TestAssert.Equal(7d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Size ?? 0d);
        TestAssert.Equal(new RgbColor(0, 170, 85), slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Fill.Color ?? default);
        TestAssert.Equal(new RgbColor(85, 0, 170), slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Line.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Line.Alpha ?? 0d);
        TestAssert.Equal(12d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Explosion ?? 0d);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Plots[0].Series[0].PointStyles[0].Index ?? -1);
        TestAssert.Equal(new RgbColor(204, 136, 68), slide.SlideNodes[4].Chart?.Plots[0].Series[0].PointStyles[0].Fill.Color ?? default);
        TestAssert.Equal(new RgbColor(34, 68, 102), slide.SlideNodes[4].Chart?.Plots[0].Series[0].PointStyles[0].Line.Color ?? default);
        TestAssert.Equal(18d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].PointStyles[0].Explosion ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].Smooth == true, "Expected chart series smooth flag in the scene model.");
        TestAssert.Equal("bubbleChart", slide.SlideNodes[4].Chart?.Plots[1].Kind ?? string.Empty);
        TestAssert.Equal(PptxSceneChartPlotKind.Bubble, slide.SlideNodes[4].Chart?.Plots[1].PlotKind);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Plots[1].PlotAreaIndex ?? -1);
        TestAssert.Equal(0, slide.SlideNodes[4].Chart?.Plots[1].KindIndex ?? -1);
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, slide.SlideNodes[4].Chart?.Plots[1].GroupingKind);
        TestAssert.Equal(PptxSceneChartBarDirection.Unknown, slide.SlideNodes[4].Chart?.Plots[1].BarDirectionKind);
        TestAssert.Equal(8, slide.SlideNodes[4].Chart?.Plots[1].Series[0].Index ?? -1);
        TestAssert.Equal(4, slide.SlideNodes[4].Chart?.Plots[1].Series[0].Order ?? -1);
        TestAssert.Equal("Scatter", slide.SlideNodes[4].Chart?.Plots[1].Series[0].Name ?? string.Empty);
        TestAssert.Equal("pct25", slide.SlideNodes[4].Chart?.Plots[1].Series[0].PatternFill.Preset ?? string.Empty);
        TestAssert.Equal(new RgbColor(17, 34, 51), slide.SlideNodes[4].Chart?.Plots[1].Series[0].PatternFill.Foreground ?? default);
        TestAssert.Equal(new RgbColor(240, 224, 208), slide.SlideNodes[4].Chart?.Plots[1].Series[0].PatternFill.Background ?? default);
        TestAssert.Equal(2.5d, slide.SlideNodes[4].Chart?.Plots[1].Series[0].XValues[1] ?? 0d);
        TestAssert.Equal(3.5d, slide.SlideNodes[4].Chart?.Plots[1].Series[0].YValues[0] ?? 0d);
        TestAssert.Equal(16d, slide.SlideNodes[4].Chart?.Plots[1].Series[0].BubbleSizes[1] ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[1].VaryColors is null, "Expected missing varyColors metadata to remain distinct from the effective default.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[1].Series[0].Marker.IsDefined == false, "Expected a missing marker element to remain distinct from the effective marker default.");
        TestAssert.Equal("lineChart", slide.SlideNodes[4].Chart?.Plots[2].Kind ?? string.Empty);
        TestAssert.Equal(PptxSceneChartPlotKind.Line, slide.SlideNodes[4].Chart?.Plots[2].PlotKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[2].Series[0].Marker.IsDefined == false, "Expected missing line-chart marker metadata to remain distinct from explicit default markers.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[2].Series[0].Smooth == false, "Expected explicit smooth disable to remain distinct from a missing smooth element.");
        TestAssert.Equal("valAx", slide.SlideNodes[4].Chart?.Axes[1].Kind ?? string.Empty);
        TestAssert.Equal(PptxSceneChartAxisKind.Value, slide.SlideNodes[4].Chart?.Axes[1].AxisKind);
        TestAssert.Equal(PptxSceneChartAxisKind.Category, slide.SlideNodes[4].Chart?.Axes[0].AxisKind);
        TestAssert.Equal("l", slide.SlideNodes[4].Chart?.Axes[1].Position ?? string.Empty);
        TestAssert.Equal(PptxSceneChartAxisPosition.Left, slide.SlideNodes[4].Chart?.Axes[1].PositionKind);
        TestAssert.Equal("20", slide.SlideNodes[4].Chart?.Axes[0].CrossAxisId ?? string.Empty);
        TestAssert.Equal("autoZero", slide.SlideNodes[4].Chart?.Axes[0].Crosses ?? string.Empty);
        TestAssert.Equal(PptxSceneChartAxisCrosses.AutoZero, slide.SlideNodes[4].Chart?.Axes[0].CrossesKind);
        TestAssert.Equal("10", slide.SlideNodes[4].Chart?.Axes[1].CrossAxisId ?? string.Empty);
        TestAssert.Equal("max", slide.SlideNodes[4].Chart?.Axes[1].Crosses ?? string.Empty);
        TestAssert.Equal(PptxSceneChartAxisCrosses.Maximum, slide.SlideNodes[4].Chart?.Axes[1].CrossesKind);
        TestAssert.Equal(2.5d, slide.SlideNodes[4].Chart?.Axes[1].CrossesAt ?? 0d);
        TestAssert.Equal("between", slide.SlideNodes[4].Chart?.Axes[1].CrossBetween ?? string.Empty);
        TestAssert.Equal(PptxSceneChartAxisCrossBetween.Between, slide.SlideNodes[4].Chart?.Axes[1].CrossBetweenKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].IsReversed == true, "Expected chart axis orientation metadata in the scene model.");
        TestAssert.Equal("out", slide.SlideNodes[4].Chart?.Axes[0].MajorTickMark ?? string.Empty);
        TestAssert.Equal(PptxSceneChartAxisTickMark.Outside, slide.SlideNodes[4].Chart?.Axes[0].MajorTickMarkKind);
        TestAssert.Equal("in", slide.SlideNodes[4].Chart?.Axes[0].MinorTickMark ?? string.Empty);
        TestAssert.Equal(PptxSceneChartAxisTickMark.Inside, slide.SlideNodes[4].Chart?.Axes[0].MinorTickMarkKind);
        TestAssert.Equal(PptxSceneChartTickLabelPosition.Low, slide.SlideNodes[4].Chart?.Axes[0].TickLabelPositionKind);
        TestAssert.Equal(100, slide.SlideNodes[4].Chart?.Axes[0].LabelOffset ?? 0);
        TestAssert.Equal(2, slide.SlideNodes[4].Chart?.Axes[0].TickLabelSkip ?? 0);
        TestAssert.Equal(3, slide.SlideNodes[4].Chart?.Axes[0].TickMarkSkip ?? 0);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[0].IsDeleted is null, "Expected missing chart category-axis delete metadata to remain distinct from an explicit disable.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[0].NoMultiLevelLabels == true, "Expected chart axis multi-level-label flag in the scene model.");
        TestAssert.Equal("Categories", slide.SlideNodes[4].Chart?.Axes[0].Title.Text ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[0].Title.Overlay == false, "Expected chart category-axis title overlay flag in the scene model.");
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.Axes[0].Title.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal(new RgbColor(34, 68, 136), slide.SlideNodes[4].Chart?.Axes[0].Title.TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].IsDeleted == false, "Expected chart axis delete flag in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].NoMultiLevelLabels is null, "Expected missing chart value-axis multi-level-label metadata to remain distinct from an explicit disable.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].HasScaling == true, "Expected chart axis scaling ownership in the scene model.");
        TestAssert.Equal(0d, slide.SlideNodes[4].Chart?.Axes[1].Minimum ?? -1d);
        TestAssert.Equal(20d, slide.SlideNodes[4].Chart?.Axes[1].Maximum ?? 0d);
        TestAssert.Equal(5d, slide.SlideNodes[4].Chart?.Axes[1].MajorUnit ?? 0d);
        TestAssert.Equal(1d, slide.SlideNodes[4].Chart?.Axes[1].MinorUnit ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].HasMajorGridlines == true, "Expected chart major gridline visibility in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].HasMinorGridlines == false, "Expected hidden chart minor gridline visibility in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[0].Line.HasLine == true, "Expected hidden chart axis line ownership in the scene model.");
        TestAssert.Equal(0d, slide.SlideNodes[4].Chart?.Axes[0].Line.Width ?? -1d);
        TestAssert.Equal(0d, slide.SlideNodes[4].Chart?.Axes[0].Line.Alpha ?? -1d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].Line.HasLine == true, "Expected chart axis line style in the scene model.");
        TestAssert.Equal(new RgbColor(51, 102, 153), slide.SlideNodes[4].Chart?.Axes[1].Line.Color ?? default);
        TestAssert.Equal(1.5d, slide.SlideNodes[4].Chart?.Axes[1].Line.Width ?? 0d);
        TestAssert.Equal(0.75d, slide.SlideNodes[4].Chart?.Axes[1].Line.Alpha ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.HasLine == true, "Expected chart major gridline line style in the scene model.");
        TestAssert.Equal(new RgbColor(136, 153, 170), slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.Width ?? 0d);
        TestAssert.Equal(0.5d, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.Alpha ?? 0d);
        TestAssert.Equal(2d, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.DashPattern[0] ?? 0d);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.Cap ?? 0);
        TestAssert.Equal(2, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.Join ?? 0);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineLine.HasLine == true, "Expected hidden chart minor gridline line style in the scene model.");
        TestAssert.Equal(0d, slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineLine.Width ?? -1d);
        TestAssert.Equal(0d, slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineLine.Alpha ?? -1d);
        TestAssert.Equal("Aptos", slide.SlideNodes[4].Chart?.Axes[1].TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal(9d, slide.SlideNodes[4].Chart?.Axes[1].TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(101, 67, 33), slide.SlideNodes[4].Chart?.Axes[1].TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].TextStyle.Bold == true, "Expected axis bold style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].TextStyle.Italic == true, "Expected axis italic style in the scene model.");
        TestAssert.Equal("high", slide.SlideNodes[4].Chart?.Axes[1].TickLabelPosition ?? string.Empty);
        TestAssert.Equal("$#,##0", slide.SlideNodes[4].Chart?.Axes[1].NumberFormat ?? string.Empty);
        TestAssert.Equal("Value", slide.SlideNodes[4].Chart?.Axes[1].Title.Text ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].Title.Overlay is null, "Expected missing chart axis-title overlay metadata to remain distinct from an explicit disable.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].Title.Layout.HasLayout == true, "Expected chart value-axis title layout in the scene model.");
        TestAssert.Equal(0.02d, slide.SlideNodes[4].Chart?.Axes[1].Title.Layout.X ?? 0d);
        TestAssert.Equal(0.5d, slide.SlideNodes[4].Chart?.Axes[1].Title.Layout.Height ?? 0d);
        TestAssert.Equal(new RgbColor(221, 238, 255), slide.SlideNodes[4].Chart?.Axes[1].Title.ShapeStyle.Fill.Color ?? default);
        TestAssert.Equal("Scene Chart", slide.SlideNodes[4].Chart?.Title.Text ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.IsAutoDeleted is null, "Expected missing chart auto-title-delete metadata to remain distinct from an explicit disable.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.Overlay == true, "Expected chart title overlay flag in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.Layout.HasLayout == true, "Expected chart title manual layout in the scene model.");
        TestAssert.Equal(0.08d, slide.SlideNodes[4].Chart?.Title.Layout.X ?? 0d);
        TestAssert.Equal(0.04d, slide.SlideNodes[4].Chart?.Title.Layout.Y ?? 0d);
        TestAssert.Equal(0.55d, slide.SlideNodes[4].Chart?.Title.Layout.Width ?? 0d);
        TestAssert.Equal(0.12d, slide.SlideNodes[4].Chart?.Title.Layout.Height ?? 0d);
        TestAssert.Equal("outer", slide.SlideNodes[4].Chart?.Title.Layout.LayoutTarget ?? string.Empty);
        TestAssert.Equal(PptxSceneChartManualLayoutTarget.Outer, slide.SlideNodes[4].Chart?.Title.Layout.LayoutTargetKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.ShapeStyle.Fill.HasFill == true, "Expected chart title fill in the scene model.");
        TestAssert.Equal(new RgbColor(254, 220, 186), slide.SlideNodes[4].Chart?.Title.ShapeStyle.Fill.Color ?? default);
        TestAssert.Equal(new RgbColor(15, 30, 45), slide.SlideNodes[4].Chart?.Title.ShapeStyle.Line.Color);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.Title.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal(13d, slide.SlideNodes[4].Chart?.Title.TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(17, 34, 170), slide.SlideNodes[4].Chart?.Title.TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.TextStyle.Bold == true, "Expected chart title bold style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.TextStyle.Italic == false, "Expected explicit chart title italic disable in the scene model.");
        TestAssert.Equal("b", slide.SlideNodes[4].Chart?.Legend.Position ?? string.Empty);
        TestAssert.Equal(PptxSceneChartLegendPosition.Bottom, slide.SlideNodes[4].Chart?.Legend.PositionKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.IsDefined == true, "Expected chart legend presence in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.IsDeleted is null, "Expected missing chart legend delete metadata to remain distinct from an explicit disable.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.Overlay == true, "Expected chart legend overlay flag in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.Layout.HasLayout == true, "Expected chart legend manual layout in the scene model.");
        TestAssert.Equal(0.7d, slide.SlideNodes[4].Chart?.Legend.Layout.X ?? 0d);
        TestAssert.Equal(0.68d, slide.SlideNodes[4].Chart?.Legend.Layout.Y ?? 0d);
        TestAssert.Equal(0.2d, slide.SlideNodes[4].Chart?.Legend.Layout.Width ?? 0d);
        TestAssert.Equal(0.15d, slide.SlideNodes[4].Chart?.Legend.Layout.Height ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.ShapeStyle.Fill.HasFill == true, "Expected chart legend fill in the scene model.");
        TestAssert.Equal(new RgbColor(238, 247, 255), slide.SlideNodes[4].Chart?.Legend.ShapeStyle.Fill.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[4].Chart?.Legend.ShapeStyle.Line.Width ?? 0d);
        TestAssert.Equal(new RgbColor(51, 68, 85), slide.SlideNodes[4].Chart?.Legend.ShapeStyle.Line.Color);
        TestAssert.Equal("Calibri", slide.SlideNodes[4].Chart?.Legend.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal(7d, slide.SlideNodes[4].Chart?.Legend.TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(34, 17, 170), slide.SlideNodes[4].Chart?.Legend.TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.TextStyle.Bold == false, "Expected explicit chart legend bold disable in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.TextStyle.Italic == true, "Expected chart legend italic style in the scene model.");
        TestAssert.Equal(PptxSceneNodeKind.Group, slide.SlideNodes[5].Kind);
        TestAssert.Equal(2743200L, slide.SlideNodes[5].GroupTransform.OffsetX);
        TestAssert.Equal(1d, slide.SlideNodes[5].GroupTransform.ScaleX);
        TestAssert.Equal(1d, slide.SlideNodes[5].GroupTransform.ScaleY);
        TestAssert.Equal(1, slide.SlideNodes[5].Children.Count);
        TestAssert.Equal(PptxSceneNodeKind.Shape, slide.SlideNodes[5].Children[0].Kind);
        TestAssert.True(slide.SlideNodes[5].Children[0].Shape?.PatternFill.HasPattern == true, "Expected grouped shape pattern fill in the scene model.");
        TestAssert.Equal("dkDnDiag", slide.SlideNodes[5].Children[0].Shape?.PatternFill.Preset ?? string.Empty);
        TestAssert.Equal(new RgbColor(47, 133, 106), slide.SlideNodes[5].Children[0].Shape?.PatternFill.Foreground ?? default);
        TestAssert.Equal(new RgbColor(238, 238, 238), slide.SlideNodes[5].Children[0].Shape?.PatternFill.Background ?? default);
        TestAssert.True(slide.SlideNodes[5].Children[0].Shape?.PictureFill.HasPicture == true, "Expected grouped shape picture fill in the scene model.");
        TestAssert.Equal("rIdShapeImage", slide.SlideNodes[5].Children[0].Shape?.PictureFill.RelationshipId ?? string.Empty);
        TestAssert.Equal(0.05d, slide.SlideNodes[5].Children[0].Shape?.PictureFill.Crop.Left ?? 0d);
        TestAssert.Equal(0.2d, slide.SlideNodes[5].Children[0].Shape?.PictureFill.Crop.Bottom ?? 0d);
        TestAssert.Equal(0.025d, slide.SlideNodes[5].Children[0].Shape?.PictureFill.Fill.Left ?? 0d);
        TestAssert.Equal(0.075d, slide.SlideNodes[5].Children[0].Shape?.PictureFill.Fill.Right ?? 0d);
        TestAssert.True(slide.LayoutNodes[0].Shape?.CustomGeometry.HasGeometry == true, "Expected layout custom geometry in the scene model.");
        TestAssert.Equal("xMid", slide.LayoutNodes[0].Shape?.CustomGeometry.Guides[0].Name ?? string.Empty);
        TestAssert.Equal(PptxSceneCustomCommandKind.LineTo, slide.LayoutNodes[0].Shape?.CustomGeometry.Paths[0].Commands[1].Kind);
        TestAssert.Equal("xMid", slide.LayoutNodes[0].Shape?.CustomGeometry.Paths[0].Commands[1].Points[0].X);
        TestAssert.True(slide.LayoutNodes[0].Shape?.CustomGeometry.Paths[0].AllowsStroke == false, "Expected custom path stroke=\"0\" to disable stroke in the scene model.");
        TestAssert.True(slide.LayoutNodes[1].IsPlaceholder, "Expected layout placeholder metadata in the scene model.");
        TestAssert.Equal(72d, slide.SlideNodes[0].Bounds?.Width ?? 0d);
        PptxSceneTextBody textBody = TestAssert.NotNull(slide.SlideNodes[0].TextBody);
        TestAssert.Equal(1, textBody.Paragraphs.Count);
        TestAssert.Equal(1, textBody.Paragraphs[0].Level);
        TestAssert.Equal(3, textBody.Paragraphs[0].Runs.Count);
        TestAssert.Equal(PptxSceneTextRunKind.Text, textBody.Paragraphs[0].Runs[0].Kind);
        TestAssert.Equal(PptxSceneTextRunKind.Break, textBody.Paragraphs[0].Runs[1].Kind);
        TestAssert.Equal(PptxSceneTextRunKind.Field, textBody.Paragraphs[0].Runs[2].Kind);
        TestAssert.Equal(1, textBody.Paragraphs[0].ResolvedStyle.Level);
        TestAssert.Equal("ctr", textBody.Paragraphs[0].ResolvedStyle.Alignment);
        TestAssert.Equal(26d, textBody.Paragraphs[0].ResolvedStyle.FontSize);
        TestAssert.True(textBody.Paragraphs[0].ResolvedStyle.Bold, "Expected master body style bold setting to survive the scene style cascade.");
        TestAssert.True(textBody.Paragraphs[0].ResolvedStyle.Italic, "Expected layout placeholder default run style to override inherited italic setting.");
        TestAssert.Equal("Arial", textBody.Paragraphs[0].ResolvedStyle.Typeface ?? string.Empty);
        TestAssert.Equal(new RgbColor(17, 17, 17), textBody.Paragraphs[0].ResolvedStyle.Color);
        TestAssert.Equal(26d, textBody.Paragraphs[0].Runs[0].ResolvedStyle.FontSize);
        TestAssert.True(textBody.Paragraphs[0].Runs[0].ResolvedStyle.Underline, "Expected run underline in resolved scene style.");
        TestAssert.True(textBody.Paragraphs[0].Runs[0].ResolvedStyle.Strike, "Expected master defaultTextStyle to participate in resolved scene style.");
        TestAssert.Equal(new RgbColor(255, 255, 0), textBody.Paragraphs[0].Runs[0].ResolvedStyle.Highlight ?? default);

        IReadOnlyList<PptxTextRunSnapshot> directTextRuns = PptxRenderer.InspectTextRuns(document, package, 0);
        PptxTextRunSnapshot directHello = directTextRuns.First(run => run.Text == "Hello");
        TestAssert.Equal(26d, directHello.FontSize);
        TestAssert.True(directHello.Underline, "Expected direct renderer inspection to expose run underline.");
        TestAssert.Equal(new RgbColor(255, 255, 0), directHello.Highlight ?? default);

        IReadOnlyList<PptxTextFrameModelSnapshot> textFrames = PptxRenderer.InspectTextFrameModels(document, package, 0);
        PptxTextFrameModelSnapshot textFrame = textFrames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Text == "Hello")));
        TestAssert.Equal(1, textFrame.Paragraphs.Count);
        TestAssert.Equal(1, textFrame.Paragraphs[0].Level);
        TestAssert.Equal("lvl2pPr", textFrame.Paragraphs[0].CascadeLevelName);
        TestAssert.True(textFrame.Paragraphs[0].ResolvedCascadeSourceCount >= 2, "Expected text model to expose inherited cascade inputs before style resolution.");
        TestAssert.Contains("shape.lstStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
        TestAssert.Contains("inherited.txStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
        TestAssert.Contains("defaultTextStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
        TestAssert.Equal("Center", textFrame.Paragraphs[0].Alignment);
        TestAssert.Equal(26d, textFrame.Paragraphs[0].FontSize);
        TestAssert.Equal("Text", textFrame.Paragraphs[0].Runs[0].Kind);
        TestAssert.Equal("Break", textFrame.Paragraphs[0].Runs[1].Kind);
        TestAssert.Equal("Field", textFrame.Paragraphs[0].Runs[2].Kind);
        TestAssert.Equal("Hello", textFrame.Paragraphs[0].Runs[0].Text);
        TestAssert.Equal(26d, textFrame.Paragraphs[0].Runs[0].FontSize);
        TestAssert.True(textFrame.Paragraphs[0].Runs[0].Underline, "Expected text model to preserve resolved run underline before layout.");
        TestAssert.Equal(new RgbColor(255, 255, 0), textFrame.Paragraphs[0].Runs[0].Highlight ?? default);

        PptxTextFlowSnapshot textFlow = PptxRenderer.InspectTextFlow(document, package, 0);
        PptxTextFlowFrameSnapshot flowFrame = textFlow.Frames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.SourceText == "Hello")));
        TestAssert.Equal(1, flowFrame.Paragraphs.Count);
        TestAssert.Equal(1, flowFrame.Paragraphs[0].Level);
        TestAssert.Equal("Center", flowFrame.Paragraphs[0].Alignment);
        TestAssert.Equal("Hello", flowFrame.Paragraphs[0].Runs[0].SourceText);
        TestAssert.Equal("Text", flowFrame.Paragraphs[0].Runs[0].Segments[0].Kind);
        TestAssert.Equal("Hello", flowFrame.Paragraphs[0].Runs[0].Segments[0].Text);
        TestAssert.Equal("Break", flowFrame.Paragraphs[0].Runs[1].SourceKind);
        TestAssert.Equal("Break", flowFrame.Paragraphs[0].Runs[1].Segments[0].Kind);
        TestAssert.True(flowFrame.TextWidth > 0d, "Expected text flow box to own text bounds before line layout.");

        PptxTextLayoutSnapshot textLayout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextFrameLayoutSnapshot layoutFrame = textLayout.Frames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Lines.Any(line => line.Spans.Any(span => span.Text == "Hello"))));
        TestAssert.Equal(1, layoutFrame.Paragraphs.Count);
        TestAssert.Equal(2, layoutFrame.Paragraphs[0].Lines.Count);
        TestAssert.Equal("Hello", layoutFrame.Paragraphs[0].Lines[0].Spans[0].Text);
        TestAssert.Equal("Hello", layoutFrame.Paragraphs[0].Lines[0].Spans[0].SourceText);
        TestAssert.Equal("1", layoutFrame.Paragraphs[0].Lines[1].Spans[0].Text);
        TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].EndX > layoutFrame.Paragraphs[0].Lines[0].StartX, "Expected layout line to own measured advance before PDF emission.");
        TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].TopY > layoutFrame.Paragraphs[0].Lines[0].BaselineY, "Expected layout line boxes to expose top-to-baseline geometry before PDF emission.");
        TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].Advance > 0d, "Expected layout line boxes to expose line advance before paragraph stacking.");
        TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].BaselineOffset > 0d, "Expected layout line boxes to own baseline offset separately from text spans.");
        TestAssert.Equal("Default", layoutFrame.Paragraphs[0].Lines[0].LineSpacingKind);
    }

    public static void PptxSyntheticShapesProduceDrawingOperators()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
                </Relationships>
                """,
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldSz cx="9144000" cy="6858000"/>
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="ellipse"/>
                          <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                        </p:spPr>
                      </p:sp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="line"/>
                          <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="4572000" y="3657600"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="roundRect"/>
                          <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="5943600" y="3657600"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:pattFill prst="dkDnDiag">
                            <a:fgClr><a:srgbClr val="2F856A"/></a:fgClr>
                            <a:bgClr><a:srgbClr val="EEEEEE"/></a:bgClr>
                          </a:pattFill>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 1 1 rg", pdf);
        TestAssert.Contains("0 0 720 540 re f", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("72 396 144 72 re f", pdf);
        TestAssert.Contains("0 0 1 RG", pdf);
        TestAssert.Contains("0 1 0 rg", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("72 252 m 216 180 l S", pdf);
        TestAssert.Contains("371.52 180 m", pdf);
        TestAssert.Contains("0.933 0.933 0.933 rg", pdf);
        TestAssert.Contains("0.184 0.522 0.416 RG", pdf);
        TestAssert.Contains(" re W n", pdf);
    }

    public static void PptxSyntheticArrowAndConnectorShapesRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="2743200"/><a:ext cx="1828800" cy="0"/></a:xfrm>
                          <a:prstGeom prst="line"/>
                          <a:ln w="12700"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm flipV="1"><a:off x="2743200" y="914400"/><a:ext cx="0" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="straightConnector1"/>
                          <a:ln w="12700"><a:solidFill><a:srgbClr val="2F856A"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="downArrow"/>
                          <a:solidFill><a:srgbClr val="C00000"/></a:solidFill>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.133 0.133 0.133 rg", pdf);
        TestAssert.Contains("72 324.5 m", pdf);
        TestAssert.Contains("213.5 323.5 l", pdf);
        TestAssert.Contains("216 324 m", pdf);
        TestAssert.Contains("1 -0 -0 -1 0 864 cm", pdf);
        TestAssert.Contains("216 468 m 216 396 l S", pdf);
        TestAssert.Contains("0.184 0.522 0.416 rg", pdf);
        TestAssert.Contains("0.753 0 0 rg", pdf);
        TestAssert.Contains("90 468 m", pdf);
        TestAssert.Contains("126 432 l", pdf);
        TestAssert.Contains("108 396 l", pdf);
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
    }

    public static void PptxSyntheticGroupedConnectorHonorsGroupFlip()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:grpSp>
                        <p:nvGrpSpPr><p:cNvPr id="2" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                        <p:grpSpPr>
                          <a:xfrm flipV="1">
                            <a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/>
                            <a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="1828800"/>
                          </a:xfrm>
                        </p:grpSpPr>
                        <p:cxnSp>
                          <p:spPr>
                            <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="914400"/></a:xfrm>
                            <a:prstGeom prst="straightConnector1"/>
                            <a:ln w="12700"><a:solidFill><a:srgbClr val="2F856A"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
                          </p:spPr>
                        </p:cxnSp>
                      </p:grpSp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 -0 -0 -1 0 720 cm", pdf);
        TestAssert.Contains("72 396 m 72 324 l S", pdf);
        TestAssert.Contains("0.184 0.522 0.416 rg", pdf);
    }

    public static void PptxSyntheticCurvedConnectorRendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector3"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="666666"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.4 0.4 0.4 RG", pdf);
        TestAssert.Contains("72 468 m", pdf);
        TestAssert.Contains("108 468 144 432 144 396 c", pdf);
        TestAssert.Contains("144 360 180 324 216 324 c", pdf);
        TestAssert.Contains("S", pdf);
    }

    public static void PptxSyntheticCurvedConnector2RendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector2"><a:avLst/></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.867 0.867 0.867 RG", pdf);
        TestAssert.Contains("72 468 m", pdf);
        TestAssert.Contains("111.765 468 144 435.765 144 396 c", pdf);
        TestAssert.Contains("S", pdf);
        TestAssert.Contains("144 396 m", pdf);
        TestAssert.Contains("147.15 403 l", pdf);
        TestAssert.Contains("140.85 403 l", pdf);
    }

    public static void PptxSyntheticCurvedConnector2LoopUsesQuarterTurnTangents()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm rot="5400000" flipH="1" flipV="1"><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector2"><a:avLst/></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector2"><a:avLst/></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
                        </p:spPr>
                      </p:cxnSp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("-0 1 -1 -0", pdf);
        TestAssert.Contains("72 468 m", pdf);
        TestAssert.Contains("111.765 468 144 435.765 144 396 c", pdf);
        TestAssert.Contains("144 396 m", pdf);
        TestAssert.Contains("147.15 403 l", pdf);
    }

    public static void PptxSyntheticCustomGeometryCubicPathRendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:custGeom>
                            <a:pathLst>
                              <a:path w="21600" h="10800" fill="none">
                                <a:moveTo><a:pt x="0" y="5400"/></a:moveTo>
                                <a:cubicBezTo>
                                  <a:pt x="5400" y="0"/>
                                  <a:pt x="16200" y="10800"/>
                                  <a:pt x="21600" y="5400"/>
                                </a:cubicBezTo>
                              </a:path>
                            </a:pathLst>
                          </a:custGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="008000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0.502 0 RG", pdf);
        TestAssert.Contains("72 432 m", pdf);
        TestAssert.Contains("108 468 180 396 216 432 c", pdf);
        TestAssert.Contains("S", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CUSTOM_GEOMETRY"), "Renderable custom cubic geometry should not emit the unsupported diagnostic.");
    }

    public static void PptxSyntheticCustomGeometryArcPathRendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:custGeom>
                            <a:gdLst>
                              <a:gd name="rx" fmla="val 10800"/>
                              <a:gd name="ry" fmla="val 5400"/>
                              <a:gd name="start" fmla="val 0"/>
                              <a:gd name="sweep" fmla="val 10800000"/>
                            </a:gdLst>
                            <a:pathLst>
                              <a:path w="21600" h="10800" fill="none">
                                <a:moveTo><a:pt x="21600" y="5400"/></a:moveTo>
                                <a:arcTo wR="rx" hR="ry" stAng="start" swAng="sweep"/>
                              </a:path>
                            </a:pathLst>
                          </a:custGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="008000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0.502 0 RG", pdf);
        TestAssert.Contains("216 432 m", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CUSTOM_GEOMETRY"), "Renderable custom arc geometry should not emit the unsupported diagnostic.");
    }

    public static void PptxSyntheticPresetArcRendersArcInsteadOfRectangle()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 0"/><a:gd name="adj2" fmla="val 5400000"/></a:avLst></a:prstGeom>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="444444"/></a:solidFill><a:prstDash val="sysDash"/></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.267 0.267 0.267 RG", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.DoesNotContain("72 396 72 72 re S", pdf);
    }

    public static void PptxSyntheticShapeStrokeDashCapAndJoinRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="0"/></a:xfrm>
                        <a:prstGeom prst="line"/>
                        <a:ln w="25400" cap="rnd">
                          <a:solidFill><a:srgbClr val="222222"/></a:solidFill>
                          <a:prstDash val="dashDot"/>
                        </a:ln>
                      </p:spPr>
                    </p:sp>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="triangle"/>
                        <a:noFill/>
                        <a:ln w="25400">
                          <a:solidFill><a:srgbClr val="445566"/></a:solidFill>
                          <a:bevel/>
                        </a:ln>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("[8 6 2 6 ] 0 d", pdf);
        TestAssert.Contains("1 J", pdf);
        TestAssert.Contains("1 j", pdf);
        TestAssert.Contains("[] 0 d", pdf);
        TestAssert.Contains("2 j", pdf);
    }

    public static void PptxSyntheticShapeStrokeDashPresetVariantsRender()
    {
        string[] dashValues = ["dot", "sysDot", "dash", "sysDash", "lgDash", "dashDot", "sysDashDot", "lgDashDot", "lgDashDotDot", "sysDashDotDot", "solid"];
        var shapes = new StringBuilder();
        for (int i = 0; i < dashValues.Length; i++)
        {
            int y = 914400 + i * 274320;
            shapes.Append(CultureInfo.InvariantCulture, $"""
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="{y}"/><a:ext cx="1828800" cy="0"/></a:xfrm>
                        <a:prstGeom prst="line"/>
                        <a:ln w="12700"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:prstDash val="{dashValues[i]}"/></a:ln>
                      </p:spPr>
                    </p:sp>

                """);
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                {{shapes}}
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("[1 2 ] 0 d", pdf);
        TestAssert.Contains("[4 3 ] 0 d", pdf);
        TestAssert.Contains("[8 3 ] 0 d", pdf);
        TestAssert.Contains("[4 3 1 3 ] 0 d", pdf);
        TestAssert.Contains("[8 3 1 3 ] 0 d", pdf);
        TestAssert.Contains("[8 3 1 3 1 3 ] 0 d", pdf);
        TestAssert.True(CountOccurrences(pdf, "[] 0 d") == dashValues.Length - 1, "Each dashed line should reset the dash pattern; solid must not set a dash pattern.");
    }

    public static void PptxSyntheticOuterShadowRendersOffsetShape()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                        <a:effectLst><a:outerShdw dist="91440" dir="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/GS50000F100000S gs", pdf);
        TestAssert.Contains("79.2 396 72 72 re f", pdf);
        TestAssert.Contains("72 396 72 72 re f", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_EFFECT"), "Supported outer shadow should not emit an unsupported-effect diagnostic.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TRANSPARENCY"), "Rendered outer shadow alpha should not emit an unsupported-transparency diagnostic.");
    }

    public static void PptxSyntheticGlowRendersExpandedShape()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                        <a:effectLst><a:glow rad="91440"><a:srgbClr val="0000FF"><a:alpha val="25000"/></a:srgbClr></a:glow></a:effectLst>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/GS25000F100000S gs", pdf);
        TestAssert.Contains("64.8 388.8 86.4 86.4 re f", pdf);
        TestAssert.Contains("72 396 72 72 re f", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_EFFECT"), "Supported glow should not emit an unsupported-effect diagnostic.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TRANSPARENCY"), "Rendered glow alpha should not emit an unsupported-transparency diagnostic.");
    }

    public static void PptxSyntheticLinearGradientShapeUsesPdfShading()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:gradFill>
                          <a:gsLst>
                            <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
                            <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
                          </a:gsLst>
                          <a:lin ang="0"/>
                        </a:gradFill>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/ShadingType 2", pdf);
        TestAssert.Contains("/Sh1 sh", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Supported linear gradient should not emit an unsupported-gradient diagnostic.");
    }

    public static void PptxSyntheticMultiStopLinearGradientShapeUsesStitchedPdfShading()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:gradFill>
                          <a:gsLst>
                            <a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>
                            <a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>
                            <a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>
                          </a:gsLst>
                          <a:lin ang="0"/>
                        </a:gradFill>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/FunctionType 3", pdf);
        TestAssert.Contains("/Bounds [0.5]", pdf);
        TestAssert.Contains("/Sh1 sh", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Supported multi-stop linear gradient should not emit an unsupported-gradient diagnostic.");
    }

    public static void PptxSyntheticLineEndPresetVariantsRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="25400"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="stealth"/></a:ln></p:spPr></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="1371600"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="25400"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="diamond"/></a:ln></p:spPr></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="25400"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="oval"/></a:ln></p:spPr></p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("216 468 m", pdf);
        TestAssert.Contains("208 471.2 l", pdf);
        TestAssert.Contains("210.4 468 l", pdf);
        TestAssert.Contains("208 464.8 l", pdf);
        TestAssert.Contains("212 435.2 l", pdf);
        TestAssert.Contains("208 432 l", pdf);
        TestAssert.Contains("212 428.8 l", pdf);
        TestAssert.Contains(" c", pdf);
    }

    public static void PptxSyntheticLineEndSizeVariantsAffectMarkerGeometry()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="0"/></a:xfrm><a:prstGeom prst="line"/><a:ln w="25400"><a:solidFill><a:srgbClr val="222222"/></a:solidFill><a:tailEnd type="diamond" w="sm" len="lg"/></a:ln></p:spPr></p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("216 468 m", pdf);
        TestAssert.Contains("210 469.6 l", pdf);
        TestAssert.Contains("204 468 l", pdf);
        TestAssert.Contains("210 466.4 l", pdf);
    }

    public static void PptxSyntheticRotatedShapeProducesTransform()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm rot="2700000" flipH="1"><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("q", pdf);
        TestAssert.Contains("-0.707 0.707 0.707 0.707", pdf);
        TestAssert.Contains(" cm", pdf);
        TestAssert.Contains("Q", pdf);
    }

    public static void PptxSyntheticRotatedTextBoxProducesTransform()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm rot="1800000"><a:off x="2743200" y="1828800"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3200"><a:latin typeface="Arial"/></a:rPr><a:t>Rotated text</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.866 -0.5 0.5 0.866", pdf);
        TestAssert.Contains("0052", pdf);
    }

    public static void PptxSyntheticFlippedRotatedTextBoxKeepsTextReadable()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm rot="10800000" flipV="1"><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Flipped text</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(1, Regex.Matches(pdf, @"-1\s+-?0\s+-?0\s+1\s+432\s+0\s+cm").Count);
        TestAssert.Contains("0046", pdf);
    }

    public static void PptxSyntheticBodyPrRotationOverridesShapeTextTransform()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm rot="10800000" flipV="1"><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr rot="0" lIns="0" rIns="0" tIns="0" bIns="0" anchor="ctr"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Readable text</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Equal(1, Regex.Matches(pdf, @"-1\s+-?0\s+-?0\s+1\s+432\s+0\s+cm").Count);
    }

    public static void PptxSyntheticTextOrientationVariantsProduceTransforms()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="vert" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>VERT</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="2286000" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="vert270" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>V270</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="eaVert" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>EAV</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="5029200" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="wordArtVert" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>WAV</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="6400800" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="wordArtVertRtl" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>WAR</a:t></a:r></a:p></p:txBody></p:sp>
                    <p:sp><p:spPr><a:xfrm><a:off x="7772400" y="914400"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr><p:txBody><a:bodyPr vert="mongolianVert" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>MON</a:t></a:r></a:p></p:txBody></p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 -1 1 0", pdf);
        TestAssert.Contains("-0 1 -1 -0", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TEXT_ORIENTATION"), "Known text orientation variants should be routed through the orientation model.");
    }

    public static void PptxSyntheticTextBoxEmbedsFontAndDrawsGlyphs()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr/>
                          <a:lstStyle/>
                          <a:p>
                            <a:r>
                              <a:rPr sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr>
                              <a:t>Hello</a:t>
                            </a:r>
                          </a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.Contains("/ToUnicode", pdf);
        TestAssert.Contains("BT", pdf);
        TestAssert.Contains("/F1 24 Tf", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticTextBoxHonorsNoFillText()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="2400"><a:noFill/></a:rPr><a:t>Hidden</a:t></a:r>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t> visible</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/GS0F100000S gs", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticTextBoxRendersTextOutline()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r>
                          <a:rPr sz="2400">
                            <a:noFill/>
                            <a:ln w="12700"><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln>
                            <a:latin typeface="Arial"/>
                          </a:rPr>
                          <a:t>Outline</a:t>
                        </a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0.667 0 RG", pdf);
        TestAssert.Contains("1 Tr", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticTextBoxUsesThemeHyperlinkColor()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes().Replace(
                "</Types>",
                "  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\r\n</Types>",
                StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/theme/theme1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="HyperlinkTheme">
                  <a:themeElements>
                    <a:clrScheme name="HyperlinkTheme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
                    </a:clrScheme>
                    <a:fontScheme name="HyperlinkTheme"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:hlinkClick r:id="rIdHyper"/></a:rPr><a:t>Link</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticTextBoxResolvesDrawingMlColorForms()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:prstClr val="orange"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Preset</a:t></a:r>
                        <a:br/>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:sysClr val="windowText" lastClr="112233"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>System</a:t></a:r>
                        <a:br/>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:scrgbClr r="50000" g="25000" b="0"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>ScRgb</a:t></a:r>
                        <a:br/>
                        <a:r><a:rPr sz="2400"><a:solidFill><a:hslClr hue="7200000" sat="100000" lum="50000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Hsl</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0.647 0 rg", pdf);
        TestAssert.Contains("0.067 0.133 0.2 rg", pdf);
        TestAssert.Contains("0.502 0.251 0 rg", pdf);
        TestAssert.Contains("0 1 0 rg", pdf);
    }

    public static void PptxSyntheticTextBoxHonorsBodyInsets()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="914400" rIns="0" tIns="914400" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Inset</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 72d);
    }

    public static void PptxTextModelExposesTypedBodyProperties()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="vert270" anchor="b" wrap="none" vertOverflow="ellipsis" numCol="3" spcCol="914400">
                        <a:normAutofit fontScale="80000" lnSpcReduction="12000"/>
                      </a:bodyPr>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Body properties</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal("Vertical270", frame.Orientation);
        TestAssert.Equal("Bottom", frame.VerticalAnchor);
        TestAssert.Equal("None", frame.WrapMode);
        TestAssert.Equal("Ellipsis", frame.VerticalOverflow);
        TestAssert.Equal(3, frame.ColumnCount);
        TestAssert.Equal(72d, frame.ColumnSpacing);
        TestAssert.Equal(0.8d, frame.FontScale);
    }

    public static void PptxSyntheticTextBoxHonorsLineBreaks()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="1800"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 79.2 447.581 Tm", pdf);
        TestAssert.Contains("1 0 0 1 79.2 425.261 Tm", pdf);
    }

    public static void PptxSyntheticTextBoxLineBreaksUseExplicitLineSpacing()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc></a:pPr><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="1800"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextLineLayoutSnapshot[] lines = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines.All(line => line.LineSpacingKind == "Absolute"), "Expected manual-break lines to keep absolute spcPts line spacing.");
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 24d) < 0.01d), "Expected absolute spcPts line spacing to own a 24pt line box advance.");
        TestAssert.True(Math.Abs((lines[0].BaselineY - lines[1].BaselineY) - 24d) < 0.01d, "Expected manual line break baselines to step by the absolute line spacing.");
    }

    public static void PptxSyntheticTrailingBreakUsesEndParagraphFontSize()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r><a:br/><a:endParaRPr sz="7200"/></a:p>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        double baselineGap = lines[0].BaselineY - lines[1].BaselineY;
        TestAssert.True(baselineGap > 100d, $"Expected 72pt endParaRPr trailing line to push the next paragraph down, got {baselineGap.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
    }

    public static void PptxSyntheticEmptyParagraphUsesEndParagraphFontSize()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr/><a:lstStyle/>
                        <a:p><a:endParaRPr sz="7200"/></a:p>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="4572000" y="914400"/><a:ext cx="2743200" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr/><a:lstStyle/>
                        <a:p><a:endParaRPr sz="1800"/></a:p>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        TestAssert.Equal(2, layout.Frames.Count);

        PptxTextLineLayoutSnapshot largeEndParagraphLine = layout.Frames[0]
            .Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .Single();
        PptxTextLineLayoutSnapshot smallEndParagraphLine = layout.Frames[1]
            .Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .Single();

        double baselineGap = smallEndParagraphLine.BaselineY - largeEndParagraphLine.BaselineY;
        TestAssert.True(baselineGap > 50d, $"Expected empty 72pt endParaRPr paragraph to advance more than the 18pt control, got {baselineGap.ToString("0.###", CultureInfo.InvariantCulture)}pt.");
    }

    public static void PptxSyntheticVerticalAnchorUsesEndParagraphFontSizeForEmptySpacing()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r></a:p>
                        <a:p><a:pPr><a:spcBef><a:spcPct val="100000"/></a:spcBef></a:pPr><a:endParaRPr sz="7200"/></a:p>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="4572000" y="914400"/><a:ext cx="2743200" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                      <p:txBody>
                        <a:bodyPr tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r></a:p>
                        <a:p><a:pPr><a:spcBef><a:spcPts val="7200"/></a:spcBef></a:pPr><a:endParaRPr sz="7200"/></a:p>
                        <a:p><a:r><a:rPr sz="1800"/><a:t>Next</a:t></a:r></a:p>
                      </p:txBody>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        TestAssert.Equal(2, layout.Frames.Count);

        PptxTextLineLayoutSnapshot[] percentSpacingLines = layout.Frames[0]
            .Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();
        PptxTextLineLayoutSnapshot[] pointSpacingLines = layout.Frames[1]
            .Paragraphs
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, percentSpacingLines.Length);
        TestAssert.Equal(2, pointSpacingLines.Length);
        TestAssert.True(Math.Abs(percentSpacingLines[0].BaselineY - pointSpacingLines[0].BaselineY) < 0.01d, "Expected percent and equivalent point spacing before an empty endParaRPr paragraph to estimate the same anchored text height.");
        TestAssert.True(Math.Abs(percentSpacingLines[1].BaselineY - pointSpacingLines[1].BaselineY) < 0.01d, "Expected anchored follow-up paragraph baselines to stay equivalent after empty endParaRPr spacing.");
    }

    public static void PptxSyntheticTextWrapDropsBreakSpaceAtLineEnd()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="650000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha Beta</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        string firstLine = string.Concat(lines[0].Spans.Select(span => span.Text));
        string secondLine = string.Concat(lines[1].Spans.Select(span => span.Text));
        TestAssert.Equal("Alpha", firstLine);
        TestAssert.Equal("Beta", secondLine);
    }

    public static void PptxSyntheticTextWrapNoneKeepsLongTextOnOneLine()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="650000" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha Beta</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(1, lines.Length);
        TestAssert.Equal("Alpha Beta", string.Concat(lines[0].Spans.Select(span => span.Text)));
    }

    public static void PptxSyntheticTextBoxUsesCompatibleLineSpacing()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr compatLnSpc="1"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1200"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="1200"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(lines.All(line => line.LineSpacingKind == "Multiple"), "Expected compatLnSpc to keep a multiple line-spacing model.");
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 13.2d) < 0.01d), $"Expected compatLnSpc to use Office's tight default line advance, got {string.Join(",", lines.Select(line => line.Advance.ToString("0.###", CultureInfo.InvariantCulture)))}.");
        TestAssert.True(Math.Abs((lines[0].BaselineY - lines[1].BaselineY) - 13.2d) < 0.01d, $"Expected compatible line spacing to tighten default baseline steps, got {(lines[0].BaselineY - lines[1].BaselineY).ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticTextBoxRendersFieldText()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:fld id="{11111111-1111-1111-1111-111111111111}" type="slidenum"><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>7</a:t></a:fld></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticTextBoxIgnoresStandaloneTabElements()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>A</a:t></a:r><a:tab/><a:r><a:rPr sz="1800"/><a:t>B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 79.2d);
        AssertDoesNotContainTextMatrixAtX(pdf, 130.806d, "Standalone a:tab elements should not move following text.");
    }

    public static void PptxSyntheticTextBoxIgnoresStandaloneExplicitTabStops()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:pPr><a:tabLst><a:tab pos="1828800"/></a:tabLst></a:pPr><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>A</a:t></a:r><a:tab/><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 72d);
        AssertDoesNotContainTextMatrixAtX(pdf, 216d, "Standalone a:tab elements should not move following text.");
    }

    public static void PptxSyntheticTextBoxHonorsTabCharacters()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string slideXml = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:pPr><a:tabLst><a:tab pos="1828800"/></a:tabLst></a:pPr><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>A{TAB}B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """.Replace("{TAB}", "\t", StringComparison.Ordinal);

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = slideXml
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 72d);
        AssertContainsTextMatrixAtX(pdf, 216d);
    }

    public static void PptxSyntheticTextBoxUsesDefaultTabStops()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string slideXml = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>A{TAB}B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """.Replace("{TAB}", "\t", StringComparison.Ordinal);

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = slideXml
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 72d);
        AssertContainsTextMatrixAtX(pdf, 144d);
    }

    public static void PptxSyntheticTextBoxTreatsNarrowNoBreakSpaceAsHiddenAdvance()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha&#x202F;beta</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot line = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single();

        TestAssert.True(line.Spans.Any(span => span.Text == "Alpha"), "Expected narrow no-break space to split the preceding word span.");
        TestAssert.True(line.Spans.Any(span => span.Text == "beta"), "Expected narrow no-break space to split the following word span.");
        TestAssert.True(line.Spans.All(span => !span.Text.Contains('\u202F', StringComparison.Ordinal)), "Expected narrow no-break space to stay out of emitted text spans.");
        TestAssert.True(line.Spans.Zip(line.Spans.Skip(1), (left, right) => right.X - (left.X + left.Width)).Any(gap => gap > 0d), "Expected narrow no-break space to contribute hidden advance between words.");
        PptxTextFlowSegmentSnapshot hiddenAdvance = PptxRenderer.InspectTextFlow(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .SelectMany(run => run.Segments)
            .Single(segment => segment.Kind == "HiddenAdvance");
        TestAssert.Equal("\u202F", hiddenAdvance.AdvanceText);
    }

    public static void PptxSyntheticTextBoxTreatsNoBreakSpaceAsHiddenRegularSpaceAdvance()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2600"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha&#xA0;beta</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot line = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single();

        PptxTextSpanLayoutSnapshot alpha = line.Spans.Single(span => span.Text == "Alpha");
        PptxTextSpanLayoutSnapshot beta = line.Spans.Single(span => span.Text == "beta");
        double hiddenAdvance = beta.X - (alpha.X + alpha.Width);

        TestAssert.True(line.Spans.All(span => !span.Text.Contains('\u00A0', StringComparison.Ordinal)), "Expected no-break space to stay out of emitted text spans.");
        TestAssert.True(hiddenAdvance > 5d, "Expected no-break space to advance like a regular hidden space, not a narrow spacer.");
    }

    public static void PptxSyntheticTextBoxNoBreakSpaceKeepsBoundaryKerning()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400" kern="1200"><a:latin typeface="Arial"/></a:rPr><a:t>A&#xA0;B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot line = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single();

        PptxTextSpanLayoutSnapshot first = line.Spans.Single(span => span.Text == "A");
        PptxTextSpanLayoutSnapshot second = line.Spans.Single(span => span.Text == "B");
        double totalAdvance = second.X - first.X;

        TestAssert.True(totalAdvance > 21d, $"Expected no-break space to keep a visible regular-space class advance, got {totalAdvance}.");
        TestAssert.True(totalAdvance < 22d, $"Expected no-break space to preserve Office's kerning from the preceding glyph, got {totalAdvance}.");
        TestAssert.True(second.GlyphSpan.LeadingAdjustment < -0.1d, "Expected the visible glyph span after hidden no-break space to own its leading boundary kerning.");
    }

    public static void PptxSyntheticTextBoxOffsetsLargeTextByFontSize()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="6600"/><a:t>Large heading</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 79.2d);
    }

    public static void PptxSyntheticTextBoxRendersBulletCharacters()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Bullet item</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("2022", pdf);
    }

    public static void PptxSyntheticTextBoxMapsSymbolFontBulletCharacters()
    {
        string wingdings = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "wingding.ttf");
        if (!File.Exists(wingdings))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:pPr>
                          <a:buFont typeface="Wingdings" charset="2"/>
                          <a:buChar char="&#xA7;"/>
                        </a:pPr>
                        <a:r><a:rPr sz="1800"/><a:t>Symbol item</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextSpanLayoutSnapshot bullet = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .Single(span => span.Text == "\uF0A7");

        TestAssert.Equal("\uF0A7", bullet.Text);
        TestAssert.Equal("Wingdings", bullet.GlyphSpan.Typeface);
    }

    public static void PptxSyntheticTextBoxRendersAutoNumberedBullets()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>First item</a:t></a:r></a:p>
                      <a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Second item</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0031", pdf);
        TestAssert.Contains("0032", pdf);
    }

    public static void PptxSyntheticTextBoxRendersRomanAutoNumberedBullets()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:buAutoNum type="romanLcPeriod" startAt="4"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>One</a:t></a:r></a:p>
                      <a:p><a:pPr><a:buAutoNum type="romanLcPeriod"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Two</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0076", pdf);
        TestAssert.Contains("002E", pdf);
    }

    public static void PptxSyntheticTextBoxHonorsBulletHangingIndent()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr marL="914400" indent="-914400"><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr sz="1800"/><a:t>Indented item</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 79.2d);
        AssertContainsTextMatrixAtX(pdf, 151.2d);
    }

    public static void PptxSyntheticTextBoxHonorsBulletColorAndSize()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr><a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPts val="3000"/><a:buChar char="&#x2022;"/></a:pPr><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Styled bullet</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 30 Tf", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("/F1 24 Tf", pdf);
        TestAssert.Contains("0 0 0 rg", pdf);
    }

    public static void PptxSyntheticTextBoxUsesPerRunFontResources()
    {
        string fonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!File.Exists(Path.Combine(fonts, "arial.ttf")) ||
            !File.Exists(Path.Combine(fonts, "times.ttf")))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Arial</a:t></a:r>
                        <a:r><a:rPr sz="1800"><a:latin typeface="Times New Roman"/></a:rPr><a:t>Times</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 18 Tf", pdf);
        TestAssert.Contains("/F2 18 Tf", pdf);
    }

    public static void PptxSyntheticTextBoxUsesEaFontWhenLatinMissing()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:ea typeface="Courier New"/></a:rPr><a:t>East Asian fallback</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);
        PptxTextSpanLayoutSnapshot span = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames.SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .Single(candidate => candidate.Text.StartsWith("East", StringComparison.Ordinal));

        TestAssert.Equal("Courier New", span.GlyphSpan.Typeface ?? string.Empty);
    }

    public static void PptxSyntheticTextBoxSplitsMissingGlyphsToFallbackFont()
    {
        string fonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!File.Exists(Path.Combine(fonts, "arial.ttf")) ||
            !File.Exists(Path.Combine(fonts, "simsun.ttc")))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>A&#x4E2D;B</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);
        IReadOnlyList<PptxTextGlyphRunSnapshot> glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0);

        string expected = "A" + char.ConvertFromUtf32(0x4E2D) + "B";
        TestAssert.Equal(expected, string.Concat(glyphRuns.Select(run => run.Text)));
        TestAssert.True(
            glyphRuns.SelectMany(run => run.Glyphs).Select(glyph => glyph.ResourceName).Distinct(StringComparer.OrdinalIgnoreCase).Count() >= 2,
            "Expected a missing Arial glyph to emit through a separate fallback PDF font resource.");
        PptxTextGlyphRunAtomSnapshot fallbackGlyph = glyphRuns.SelectMany(run => run.Glyphs).Single(glyph => glyph.CodePoint == 0x4E2D);
        TestAssert.True(!string.Equals("Arial", fallbackGlyph.Typeface, StringComparison.OrdinalIgnoreCase), "Expected the CJK glyph to carry its fallback typeface.");
    }

    public static void PptxSyntheticTextBoxResolvesThemeEaAndCsFonts()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes().Replace(
                "</Types>",
                "  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\r\n</Types>",
                StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/theme/theme1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="FontTheme">
                  <a:themeElements>
                    <a:clrScheme name="FontTheme"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme>
                    <a:fontScheme name="FontTheme">
                      <a:majorFont><a:latin typeface="Arial"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Arial Unicode MS"/></a:majorFont>
                      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="SimSun"/><a:cs typeface="Tahoma"/></a:minorFont>
                    </a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="2400"><a:latin typeface="+mj-ea"/></a:rPr><a:t>MajorEastAsian</a:t></a:r>
                        <a:br/>
                        <a:r><a:rPr sz="2400"><a:latin typeface="+mn-cs"/></a:rPr><a:t>MinorComplex</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """)
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal("Microsoft YaHei", frame.Paragraphs[0].Runs[0].Typeface);
        TestAssert.Equal("Tahoma", frame.Paragraphs[0].Runs[2].Typeface);
    }

    public static void PptxSyntheticTextBoxUsesDistinctFontResourcesForStyles()
    {
        string fonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (!File.Exists(Path.Combine(fonts, "arial.ttf")) ||
            !File.Exists(Path.Combine(fonts, "arialbd.ttf")))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Regular</a:t></a:r>
                        <a:r><a:rPr sz="2400" b="1"><a:latin typeface="Arial"/></a:rPr><a:t>Bold</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(pdf.Split("/FontFile2", StringSplitOptions.None).Length - 1 >= 2, "Regular and bold Arial should embed distinct font files.");
    }

    public static void PptxSyntheticTextBoxUsesOfficeSyntheticBoldStroke()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400" b="1"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Synthetic bold</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.686 w", pdf);
        TestAssert.Contains("2 Tr", pdf);
        TestAssert.True(Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ [0-9.]+ Tm").Count == 1, "Synthetic bold should use fill-and-stroke text rather than duplicate offset glyph draws.");
    }

    public static void PptxSyntheticTextBoxUsesKerningWhenAvailable()
    {
        string times = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "times.ttf");
        if (!File.Exists(times))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="4000"><a:latin typeface="Times New Roman"/></a:rPr><a:t>To</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxCambriaMathDenseWrapProbeKeepsHeadingOnOneLine()
    {
        string input = Path.Combine(
            Directory.GetCurrentDirectory(),
            "tests",
            "Lokad.OoxPdf.Tests",
            "Cases",
            "pptx-ladder-04-cambria-math-dense-wrap-probe.pptx");
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();
        string[] texts = lines.Select(line => string.Concat(line.Spans.Select(span => span.Text))).ToArray();

        TestAssert.True(
            texts.Length > 0 && texts[0].EndsWith("France.", StringComparison.Ordinal),
            $"Expected Office-compatible first-line wrap. Lines: {string.Join(" | ", texts.Take(4))}. Widths: {string.Join(" | ", lines.Take(4).Select(line => (line.EndX - line.StartX).ToString("0.###", CultureInfo.InvariantCulture)))}");
    }

    public static void PptxSyntheticCenteredLogoBoxWrapsDefaultTypefaceText()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Aptos">
                  <a:themeElements>
                    <a:clrScheme name="Aptos">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                    </a:clrScheme>
                    <a:fontScheme name="Aptos"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3350712" cy="1847589"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:noFill/>
                      <a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
                    </p:spPr>
                    <p:style><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>
                    <p:txBody>
                      <a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="4000"/><a:t>Customer Logo</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();
        string[] texts = lines.Select(line => string.Concat(line.Spans.Select(span => span.Text)).Trim()).ToArray();

        TestAssert.Equal(2, texts.Length);
        TestAssert.Equal("Customer", texts[0]);
        TestAssert.Equal("Logo", texts[1]);
    }

    public static void PptxSyntheticTextBoxHonorsKerningThreshold()
    {
        string times = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "times.ttf");
        if (!File.Exists(times))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800" kern="3600"><a:latin typeface="Times New Roman"/></a:rPr><a:t>To</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("[<0037><0052>] TJ", pdf);
    }

    public static void PptxSyntheticTextBoxUsesPositioningForCharacterSpacing()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3600" spc="3000"><a:latin typeface="Arial"/></a:rPr><a:t>AB</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" TJ", pdf);
        TestAssert.True(!pdf.Contains(" Tc", StringComparison.Ordinal), "Character spacing should be encoded in the TJ positioning array.");
    }

    public static void PptxSyntheticTextBoxRendersSmallCapsAsScaledUppercaseRuns()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400" cap="small"><a:latin typeface="Arial"/></a:rPr><a:t>Abc DEF</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("24 Tf", pdf);
        TestAssert.Contains("19.2 Tf", pdf);
        TestAssert.True(CountOccurrences(pdf, " TJ") >= 3, "Expected small-caps text to split into full-size and scaled positioned text runs.");
    }

    public static void PptxSyntheticTextBoxRendersBaselineShiftWithCssScale()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Base</a:t></a:r><a:r><a:rPr sz="2400" baseline="30000"><a:latin typeface="Arial"/></a:rPr><a:t>2</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("24 Tf", pdf);
        TestAssert.Contains("15.6 Tf", pdf);
    }

    public static void PptxSyntheticTextBoxKeepsFontSizeForSmallBaselineShift()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Base</a:t></a:r><a:r><a:rPr sz="2400" baseline="10000"><a:latin typeface="Arial"/></a:rPr><a:t>nudge</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextSpanLayoutSnapshot[] spans = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .SelectMany(line => line.Spans)
            .ToArray();

        TestAssert.True(Math.Abs(spans[0].FontSize - 24d) < 0.001d, $"Expected base span to remain 24pt, got {spans[0].FontSize.ToString("0.###", CultureInfo.InvariantCulture)}.");
        TestAssert.True(Math.Abs(spans[1].FontSize - 24d) < 0.001d, $"Expected small baseline-shift span to remain 24pt, got {spans[1].FontSize.ToString("0.###", CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticTextBoxCentersMixedRunsTogether()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Left </a:t></a:r><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Right</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 177.478d);
    }

    public static void PptxSyntheticTextBoxWrapsAcrossMixedRuns()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha beta </a:t></a:r><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>gamma</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 79.2d);
        TestAssert.True(Regex.Matches(pdf, @"1 0 0 1 79\.2 [0-9.]+ Tm").Count >= 2, "Expected wrapped mixed runs to emit at least two text lines.");
    }

    public static void PptxSyntheticTextBoxUsesListStyleDefaults()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:defRPr sz="3600" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></a:lstStyle>
                      <a:p><a:r><a:rPr/><a:t>Default</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr/><a:t>Next</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 36 Tf", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("1 0 0 1 79.2 428.4 Tm", pdf);
        TestAssert.Contains("1 0 0 1 79.2 392.4 Tm", pdf);
    }

    public static void PptxSyntheticTextBoxUsesShapeFontRefColor()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:style><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Styled color</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 1 1 rg", pdf);
    }

    public static void PptxSyntheticShapeFontRefColorOverridesDefaultTextStyle()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:sldSz cx="9144000" cy="6858000"/>
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
                  <p:defaultTextStyle><a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:defaultTextStyle>
                </p:presentation>
                """,
            ["ppt/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Theme">
                  <a:themeElements>
                    <a:clrScheme name="Theme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                    </a:clrScheme>
                    <a:fontScheme name="Theme"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="808080"/></a:solidFill></p:spPr>
                    <p:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1400"><a:latin typeface="Arial"/></a:rPr><a:t>Font ref</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 1 1 rg", pdf);
    }

    public static void PptxSyntheticExplicitRunColorOverridesShapeFontRefColor()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:style><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Explicit color</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.True(!pdf.Contains("1 1 1 rg", StringComparison.Ordinal), "Explicit run color should override shape fontRef color.");
    }

    public static void PptxSyntheticTextBoxRendersTextHighlight()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:highlight><a:srgbClr val="800080"/></a:highlight><a:latin typeface="Arial"/></a:rPr><a:t>Marked</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.502 0 0.502 rg", pdf);
        TestAssert.Contains(" re f", pdf);
    }

    public static void PptxSyntheticTextBoxRendersStrikethrough()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400" strike="sngStrike"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Strike</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains(" re f", pdf);
    }

    public static void PptxSyntheticTextBoxHonorsParagraphSpacing()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:pPr><a:spcBef><a:spcPts val="1200"/></a:spcBef><a:lnSpc><a:spcPct val="100000"/></a:lnSpc></a:pPr>
                        <a:r><a:rPr sz="1800"/><a:t>Spaced</a:t></a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 1 79.2 446.4 Tm", pdf);
    }

    public static void PptxSyntheticTextBoxHonorsPercentParagraphSpacing()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:pPr><a:spcAft><a:spcPct val="50000"/></a:spcAft></a:pPr><a:r><a:rPr sz="2400"/><a:t>First</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="2400"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 72d);
    }

    public static void PptxSyntheticTextBoxSkipsEmptyParagraphs()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>First</a:t></a:r></a:p>
                      <a:p><a:endParaRPr lang="en-US"/></a:p>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 79.2d);
        TestAssert.True(Regex.Matches(pdf, @"1 0 0 1 79\.2 [0-9.]+ Tm").Count >= 2, "Expected empty paragraphs to preserve later paragraph placement.");
    }

    public static void PptxSyntheticTextBoxHonorsVerticalAnchor()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"/><a:t>Centered</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        AssertContainsTextMatrixAtX(pdf, 7.2d);
    }

    public static void PptxSyntheticTextBoxVerticalAnchorUsesWrappedHeight()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr tIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Alpha Beta Gamma Delta</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.True(lines.Length >= 3, "Expected the narrow center-anchored text frame to wrap.");
        TestAssert.True(lines[0].BaselineY > 475d, "Expected vertical centering to account for wrapped line count instead of one logical paragraph line.");
    }

    public static void PptxSyntheticTextBoxMiddleAnchorUsesVisibleRunFontSizes()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="4144300"/><a:ext cx="1634286" cy="699793"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr anchor="ctr"/><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1400"><a:latin typeface="Arial"/></a:rPr><a:t>       Data</a:t></a:r></a:p>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="900"><a:latin typeface="Arial"/></a:rPr><a:t>ERP/WMS/OMS/CRM</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, lines.Length);
        TestAssert.True(Math.Abs(lines[0].MaxFontSize - 14d) < 0.01d, "Expected the first line height to come from its visible 14pt run, not an unrelated 18pt fallback.");
        TestAssert.True(Math.Abs(lines[1].MaxFontSize - 9d) < 0.01d, "Expected the second line height to come from its visible 9pt run, not an unrelated 18pt fallback.");
        TestAssert.True(Math.Abs(lines[0].Advance - 16.8d) < 0.01d, "Expected default line advance to use the visible run font size.");
        TestAssert.True(Math.Abs(lines[1].Advance - 10.8d) < 0.01d, "Expected default line advance to use the visible run font size.");
    }

    public static void PptxSyntheticVerticalShapeAutoFitPrefersSingleLine()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="365760" cy="2057400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="vert270" wrap="square" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1100"><a:latin typeface="Arial"/></a:rPr><a:t>Results reintegration for action</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.True(lines.Length == 1, "Vertical spAutoFit text should shrink before accepting an avoidable word wrap.");
        TestAssert.True(lines[0].EndX - lines[0].StartX <= 150d, "Fitted vertical text should stay inside the rotated text width.");
    }

    public static void PptxSyntheticMongolianVerticalAutoFitDoesNotUseRotatedWidthShrink()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="365760"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="mongolianVert" wrap="square" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>VERTICAL STACKED TEXT</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.True(lines.Length > 1, "Stacked vertical text should use vertical chunk layout instead of the rotated single-line autofit path.");
        TestAssert.True(lines.All(line => Math.Abs(line.MaxFontSize - 24d) < 0.01d), "Stacked vertical spAutoFit should not shrink against the narrow rotated text width.");
    }

    public static void PptxSyntheticTransformedVerticalTextDoesNotDropPreclipBaselines()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="365760"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr vert="mongolianVert" wrap="square" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>VERTICAL STACKED TEXT</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        string emittedText = string.Concat(PptxRenderer.InspectTextGlyphRuns(document, package, 0).Select(run => run.Text));

        TestAssert.Contains("VERTIC", emittedText);
        TestAssert.Contains("TEXT", emittedText);
    }

    public static void PptxSyntheticHorizontalShapeAutoFitPreservesFontSizeWhenHeightOverflows()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-spautofit-headline-wrap-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(4, lines.Length);
        TestAssert.True(lines.All(line => Math.Abs(line.MaxFontSize - 18d) < 0.01d), "Horizontal spAutoFit should preserve Office's run font size when only vertical text height overflows.");
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 21.6d) < 0.01d), "Horizontal spAutoFit should keep the normal 1.2 line advance instead of shrinking text.");
    }

    public static void PptxSyntheticTextBoxHonorsNormAutofitFontScale()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr><a:normAutofit fontScale="80000"/></a:bodyPr><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="3000"><a:latin typeface="Arial"/></a:rPr><a:t>Scaled</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 24 Tf", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TEXT_AUTOFIT"), "normAutofit fontScale should be handled.");
    }

    public static void PptxSyntheticTextBoxHonorsNormAutofitLineSpacingReduction()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:normAutofit lnSpcReduction="25000"/></a:bodyPr><a:lstStyle/>
                      <a:p><a:pPr><a:lnSpc><a:spcPct val="120000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="3000"/><a:t>First</a:t></a:r><a:br/><a:r><a:rPr sz="3000"/><a:t>Second</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        IReadOnlyList<PptxTextLineLayoutSnapshot> lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.Equal(2, lines.Count);
        TestAssert.True(lines.All(line => Math.Abs(line.Advance - 27d) < 0.01d), "Expected normAutofit line spacing reduction to scale explicit percentage line spacing over the font size.");
        TestAssert.True(Math.Abs((lines[0].BaselineY - lines[1].BaselineY) - 27d) < 0.01d, "Expected reduced line spacing to drive manual-break baseline steps.");
    }

    public static void PptxSyntheticTextBoxInheritsParagraphIndentFromListStyle()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/>
                      <a:lstStyle><a:lvl1pPr marL="91440" indent="45720"/></a:lstStyle>
                      <a:p><a:pPr algn="l"/><a:r><a:rPr sz="1200"/><a:t>Indented</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot line = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single();
        PptxTextFrameModelSnapshot model = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();

        TestAssert.True(
            Math.Abs(line.Spans[0].X - 82.8d) < 0.01d,
            "Expected inherited marL plus first-line indent to move the emitted text span from the text body origin. Actual span: " + line.Spans[0].X.ToString("0.###", CultureInfo.InvariantCulture) +
            "; line start: " + line.StartX.ToString("0.###", CultureInfo.InvariantCulture) +
            "; cascade sources: " + model.Paragraphs[0].ResolvedCascadeSourceCount.ToString(CultureInfo.InvariantCulture) +
            "; margin: " + model.Paragraphs[0].MarginLeft.ToString("0.###", CultureInfo.InvariantCulture) +
            "; hanging: " + model.Paragraphs[0].HangingIndent.ToString("0.###", CultureInfo.InvariantCulture));
    }

    public static void PptxSyntheticEllipseTextUsesPresetTextRectangle()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/></a:xfrm><a:prstGeom prst="ellipse"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="ctr"/>
                      <a:lstStyle/>
                      <a:p><a:pPr algn="l"/><a:r><a:rPr sz="1800"/><a:t>7</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot line = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .Single();

        double expectedTextX = 72d + 144d * 0.1464466094067262d;
        TestAssert.True(
            Math.Abs(line.StartX - expectedTextX) < 0.01d,
            "Expected ellipse text to start inside the preset geometry text rectangle. Actual start: " + line.StartX.ToString("0.###", CultureInfo.InvariantCulture) +
            "; expected: " + expectedTextX.ToString("0.###", CultureInfo.InvariantCulture));

        IReadOnlyList<PptxTextGlyphRunSnapshot> glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0);
        TestAssert.Equal(1, glyphRuns.Count);
        TestAssert.True(
            Math.Abs(glyphRuns[0].X - expectedTextX) < 0.01d,
            "Expected clipped ellipse text to keep emitting a glyph run at the preset text-rectangle origin.");
    }

    public static void PptxSyntheticTextBoxFlowsAcrossColumns()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr numCol="3" spcCol="144000" lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="overflow"/>
                      <a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>One two</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Three four</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Five six</a:t></a:r></a:p>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Seven eight</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.True(lines.Any(line => Math.Abs(line.StartX - 72d) < 0.01d), "Expected first-column text.");
        TestAssert.True(lines.Any(line => Math.Abs(line.StartX - 219.78d) < 0.01d), "Expected second-column text. Starts: " + string.Join(", ", lines.Select(line => line.StartX.ToString("0.###", CultureInfo.InvariantCulture))));
    }

    public static void PptxSyntheticTextBoxClipsOverflow()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" vertOverflow="clip"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="4800"/><a:t>UnbreakableOverflowingText</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 432 72 36 re W n", pdf);
    }

    public static void PptxSyntheticTextBoxAllowsVerticalOverflowByDefault()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="4800"/><a:t>Overflow</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 0 72 540 re W n", pdf);
    }

    public static void PptxJustifiedTextLayoutDistributesWrappedLines()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-justify-port.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextLineLayoutSnapshot justifiedLine = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .First(line => line.Alignment == "Justify");

        TestAssert.True(justifiedLine.Spans.Count > 1, "Expected justified lines to keep word spans separate before PDF emission.");
        TestAssert.True(justifiedLine.EndX - justifiedLine.StartX > 500d, "Expected justified line to stretch to the text frame width.");
        TestAssert.True(justifiedLine.NaturalEndX < justifiedLine.EndX, "Expected justified line inspection to expose the natural pre-expansion end for Office wrap comparisons.");
        TestAssert.True(justifiedLine.BaselineMetric.Ratio > 0d, "Expected line boxes to expose the generic baseline metric ratio used for Office comparison.");
        TestAssert.True(!string.IsNullOrWhiteSpace(justifiedLine.BaselineMetric.Source), "Expected line boxes to expose baseline metric provenance.");
        TestAssert.True(justifiedLine.BaselineMetric.UnitsPerEm > 0, "Expected resolved-font line boxes to expose OpenType units-per-em diagnostics.");
        TestAssert.True(justifiedLine.BaselineMetric.WindowsAscender > 0, "Expected resolved-font line boxes to expose OS/2 Windows ascender diagnostics.");
        TestAssert.True(justifiedLine.Spans.All(span => span.Text.IndexOf(' ') < 0), "Expected justified emission spans to split drawable words from stretchable spaces.");
        TestAssert.True(justifiedLine.Spans.SelectMany(span => span.Atoms).Any(atom => atom.Kind == "Word"), "Expected layout atoms to preserve word fragments separately from spaces.");
        double[] wordStarts = justifiedLine.Spans.Select(span => span.X).Take(4).ToArray();
        TestAssert.Equal(4, wordStarts.Length);
        TestAssert.True(wordStarts.Zip(wordStarts.Skip(1), (left, right) => right - left).All(delta => delta > 0d), "Expected justified layout to expose monotonic word starts for Office text-op comparison.");
        double[] wordGaps = justifiedLine.Spans.Zip(justifiedLine.Spans.Skip(1), (left, right) => right.X - (left.X + left.Width)).ToArray();
        TestAssert.True(wordGaps.Any(gap => gap > 1d), "Expected justified layout to expose distributed spacing through positioned word starts.");
        PptxTextSpanLayoutSnapshot paragraphSpan = justifiedLine.Spans.First(span => span.Text.StartsWith("Paragraph", StringComparison.Ordinal));
        TestAssert.True(paragraphSpan.GlyphSpan.GlyphCount > 0, "Expected layout spans to own glyph ids before PDF emission.");
        TestAssert.True(paragraphSpan.GlyphSpan.Glyphs.All(glyph => glyph.GlyphId > 0), "Expected layout glyph spans to expose mapped glyph ids.");
        TestAssert.True(paragraphSpan.GlyphSpan.Glyphs.All(glyph => string.Equals(glyph.Typeface, paragraphSpan.GlyphSpan.Typeface, StringComparison.OrdinalIgnoreCase)), "Expected layout glyph spans to expose the resolved typeface for each glyph before font fallback splitting.");
        TestAssert.True(paragraphSpan.GlyphSpan.Glyphs.All(glyph => glyph.Advance > 0d), "Expected layout glyph spans to expose positive glyph advances.");
        TestAssert.True(Math.Abs(paragraphSpan.GlyphSpan.LayoutWidth - paragraphSpan.GlyphSpan.NaturalWidth) < 0.01d, "Expected justified word glyph spans to keep natural width while line positioning owns distributed spacing.");

        IReadOnlyList<PptxTextGlyphRunSnapshot> glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0);
        PptxTextGlyphRunSnapshot paragraphGlyphRun = glyphRuns.First(run => run.Text.StartsWith("Paragraph", StringComparison.Ordinal));
        TestAssert.True(paragraphGlyphRun.GlyphCount > 0, "Expected glyph-run inspection to expose glyph ids before PDF text emission.");
        TestAssert.Equal(paragraphGlyphRun.GlyphCount, paragraphGlyphRun.Glyphs.Count);
        TestAssert.True(paragraphGlyphRun.Glyphs.All(glyph => glyph.GlyphId > 0), "Expected glyph-run inspection to expose each emitted glyph id.");
        TestAssert.True(paragraphGlyphRun.Glyphs.All(glyph => string.Equals(glyph.Typeface, paragraphSpan.GlyphSpan.Typeface, StringComparison.OrdinalIgnoreCase)), "Expected glyph-run inspection to preserve current glyph typeface ownership before fallback splits it.");
        TestAssert.True(paragraphGlyphRun.Glyphs.All(glyph => !string.IsNullOrWhiteSpace(glyph.ResourceName)), "Expected glyph-run inspection to expose the PDF font resource that currently owns each glyph.");
        TestAssert.True(glyphRuns.Any(run => run.Text == "."), "Expected justified glyph-run inspection to preserve Office-like sentence-period operation boundaries.");
        TestAssert.True(paragraphGlyphRun.Width > 0d, "Expected glyph-run inspection to expose measured glyph advance.");
        TestAssert.Equal(paragraphSpan.GlyphSpan.GlyphCount, paragraphGlyphRun.GlyphCount);
        TestAssert.True(Math.Abs(paragraphSpan.GlyphSpan.NaturalWidth - paragraphGlyphRun.Width) < 0.01d, "Expected layout-owned glyph span width to match the emitted glyph-run width before PDF text operators are written.");
    }

    public static void PptxAlignmentValuesKeepDistinctOfficeTextDistributionModes()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-alignment-values-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextFlowSnapshot flow = PptxRenderer.InspectTextFlow(document, package, 0);
        string[] alignments = flow.Frames
            .SelectMany(frame => frame.Paragraphs)
            .Select(paragraph => paragraph.Alignment)
            .ToArray();

        string alignmentList = string.Join("|", alignments);
        TestAssert.Contains("Justify", alignmentList);
        TestAssert.Contains("Distributed", alignmentList);
        TestAssert.Contains("JustLow", alignmentList);
        TestAssert.Contains("ThaiDistributed", alignmentList);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        PptxTextLineLayoutSnapshot distributed = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .First(line => line.Alignment == "Distributed");

        TestAssert.True(distributed.Spans.Count > 10, "Expected distributed alignment to own per-glyph positioned spans.");
        TestAssert.True(distributed.Spans.All(span => span.Text.Length <= 2), "Expected distributed alignment to avoid word-level spans that hide letter spacing.");
        TestAssert.True(distributed.Spans.Last().X - distributed.StartX > distributed.Advance * 0.75d, "Expected distributed alignment to stretch glyph positions across the text frame.");
    }

    public static void PptxHighlightedRunDoesNotApplyImplicitTracking()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-spautofit-tracking-probe.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextFlowSnapshot flow = PptxRenderer.InspectTextFlow(document, package, 0);
        PptxTextFlowRunSnapshot[] runs = flow.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.SourceText.Length != 0)
            .ToArray();

        PptxTextFlowRunSnapshot highlighted = runs.First(run => run.SourceText == "AI");
        PptxTextFlowRunSnapshot following = runs.First(run => run.SourceText.StartsWith(" boundary", StringComparison.Ordinal));

        TestAssert.True(Math.Abs(highlighted.FontSize - 12d) < 0.01d, "Expected the probe highlight run to stay at 12pt.");
        TestAssert.True(highlighted.Segments.All(segment => Math.Abs(segment.FontScale - 1d) < 0.01d), "Expected highlight tracking to avoid fake font scaling.");

        TestAssert.True(following.FontSize > 0d, "Expected the following run to remain present in text flow.");
    }

    public static void PptxTypographyTextHyphenBoundariesRemainSeparateSpans()
    {
        string input = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "Cases",
            "pptx-ladder-04-typography-punctuation-boundaries.pptx"));
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        string[] firstLineTexts = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .First()
            .Spans
            .Select(span => span.Text)
            .Take(3)
            .ToArray();

        TestAssert.Equal("SKU", firstLineTexts[0]);
        TestAssert.Equal("-", firstLineTexts[1]);
        TestAssert.True(firstLineTexts[2].StartsWith("123", StringComparison.Ordinal), "Expected text after the hyphen to remain a separate positioned span for Office-style PDF text operations.");
    }

    public static void PptxTypographyTextEmDashBoundariesRemainSeparateSpans()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:nvPr/></p:nvSpPr>
                        <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm></p:spPr>
                        <p:txBody>
                          <a:bodyPr/>
                          <a:lstStyle/>
                          <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Plan &#x2014; execute</a:t></a:r></a:p>
                        </p:txBody>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);

        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, 0);
        string[] firstLineTexts = layout.Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .First()
            .Spans
            .Select(span => span.Text)
            .Take(3)
            .ToArray();

        TestAssert.Equal("Plan ", firstLineTexts[0]);
        TestAssert.Equal("\u2014", firstLineTexts[1]);
        TestAssert.Equal("execute", firstLineTexts[2]);
    }

    public static void PptxPrivateLayoutDiagnosticWhenRequested()
    {
        string? input = Environment.GetEnvironmentVariable("OOXPDF_PRIVATE_PPTX_PATH");
        string? output = Environment.GetEnvironmentVariable("OOXPDF_PRIVATE_LAYOUT_JSON");
        if (string.IsNullOrWhiteSpace(input) || string.IsNullOrWhiteSpace(output))
        {
            return;
        }

        int slideIndex = int.TryParse(Environment.GetEnvironmentVariable("OOXPDF_PRIVATE_SLIDE_INDEX"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : 0;
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);
        PptxTextLayoutSnapshot layout = PptxRenderer.InspectTextLayout(document, package, slideIndex);
        PptxTextFlowSnapshot flow = PptxRenderer.InspectTextFlow(document, package, slideIndex);
        PptxTextFrameModelSnapshot[] models = PptxRenderer.InspectTextFrameModels(document, package, slideIndex).ToArray();
        IReadOnlyList<PptxTextGlyphRunSnapshot> glyphRuns = PptxRenderer.InspectTextGlyphRuns(document, package, slideIndex);
        PptxSceneSnapshot scene = PptxRenderer.InspectScene(document, package);
        PptxSceneSlideSnapshot? sceneSlide = scene.Slides.FirstOrDefault(slide => slide.Index == slideIndex + 1);

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(output))!);
        object[] FlattenSceneNodes(IReadOnlyList<PptxSceneNodeSnapshot> nodes, string source)
        {
            return nodes
                .SelectMany(node => FlattenSceneNode(node, source))
                .ToArray();
        }

        IEnumerable<object> FlattenSceneNode(PptxSceneNodeSnapshot node, string source)
        {
            yield return new
            {
                source,
                node.Kind,
                node.IsPlaceholder,
                node.HasBounds,
                node.RotationDegrees,
                node.FlipHorizontal,
                node.FlipVertical,
                node.ShapePreset,
                node.HasTextBody,
                node.TextParagraphCount,
                node.TextRunCount,
                node.HasPicture,
                node.PictureRecolorKind,
                node.HasTable,
                node.TableRowCount,
                node.TableCellCount,
                node.HasChart,
                node.ChartPlotCount,
                node.ChartAxisCount,
                node.HasGroupTransform
            };

            foreach (PptxSceneNodeSnapshot child in node.Children)
            {
                foreach (object descendant in FlattenSceneNode(child, source))
                {
                    yield return descendant;
                }
            }
        }

        object? sceneSummary = sceneSlide is null
            ? null
            : new
            {
                sceneSlide.Index,
                sceneSlide.PartName,
                sceneSlide.MasterPartName,
                sceneSlide.LayoutPartName,
                sceneSlide.HasMasterBackground,
                sceneSlide.HasLayoutBackground,
                sceneSlide.HasSlideBackground,
                masterNodeCount = sceneSlide.MasterNodes.Count,
                layoutNodeCount = sceneSlide.LayoutNodes.Count,
                slideNodeCount = sceneSlide.SlideNodes.Count,
                nodes = FlattenSceneNodes(sceneSlide.MasterNodes, "master")
                    .Concat(FlattenSceneNodes(sceneSlide.LayoutNodes, "layout"))
                    .Concat(FlattenSceneNodes(sceneSlide.SlideNodes, "slide"))
                    .ToArray()
            };

        var frames = layout.Frames.Select((frame, frameIndex) => new
        {
            frameIndex,
            model = frameIndex < models.Length
                ? new
                {
                    models[frameIndex].TextX,
                    models[frameIndex].TextWidth,
                    models[frameIndex].FontScale,
                    paragraphs = models[frameIndex].Paragraphs.Select(paragraph => new
                    {
                        paragraph.Level,
                        paragraph.Alignment,
                        paragraph.FontSize,
                        paragraph.CascadeLevelName,
                        paragraph.ResolvedCascadeSourceCount,
                        paragraph.CascadeLayerNames,
                        runs = paragraph.Runs.Select(run => new
                        {
                            run.Kind,
                            length = run.Text.Length,
                            hash = ShortHash(run.Text),
                            run.FontSize,
                            run.CharacterSpacing,
                            run.Typeface,
                            run.Underline,
                            highlighted = run.Highlight is not null
                        })
                    })
                }
                : null,
            flow = frameIndex < flow.Frames.Count
                ? new
                {
                    flow.Frames[frameIndex].TextX,
                    flow.Frames[frameIndex].TextWidth,
                    flow.Frames[frameIndex].TextHeight,
                    paragraphs = flow.Frames[frameIndex].Paragraphs.Select(paragraph => new
                    {
                        paragraph.Level,
                        paragraph.Alignment,
                        paragraph.FontSize,
                        runs = paragraph.Runs.Select(run => new
                        {
                            run.SourceKind,
                            length = run.SourceText.Length,
                            hash = ShortHash(run.SourceText),
                            run.FontSize,
                            run.Typeface,
                            segments = run.Segments.Select(segment => new
                            {
                                segment.Kind,
                                length = segment.Text.Length,
                                advanceLength = segment.AdvanceText.Length,
                                segment.Draw,
                                segment.PreventCoalesce,
                                segment.FontScale
                            })
                        })
                    })
                }
                : null,
            paragraphs = frame.Paragraphs.Select((paragraph, paragraphIndex) => new
            {
                paragraphIndex,
                paragraph.Level,
                lines = paragraph.Lines.Select((line, lineIndex) => new
                {
                    lineIndex,
                    line.Alignment,
                    line.TopY,
                    line.BaselineY,
                    line.Advance,
                    line.BaselineOffset,
                    line.LineSpacingKind,
                    baselineMetric = new
                    {
                        line.BaselineMetric.Source,
                        line.BaselineMetric.Typeface,
                        line.BaselineMetric.FontSize,
                        line.BaselineMetric.Ratio,
                        line.BaselineMetric.UnitsPerEm,
                        line.BaselineMetric.WindowsAscender,
                        line.BaselineMetric.WindowsDescender,
                        line.BaselineMetric.TypographicAscender,
                        line.BaselineMetric.TypographicDescender,
                        line.BaselineMetric.TypographicLineGap
                    },
                    line.StartX,
                    line.EndX,
                    line.NaturalEndX,
                    naturalWidth = line.NaturalEndX - line.StartX,
                    alignmentExpansion = line.EndX - line.NaturalEndX,
                    line.MaxFontSize,
                    spanCount = line.Spans.Count,
                    glyphCount = line.Spans.Sum(span => span.GlyphSpan.GlyphCount),
                    spans = line.Spans.Select((span, spanIndex) => new
                    {
                        spanIndex,
                        length = span.Text.Length,
                        hash = ShortHash(span.Text),
                        span.X,
                        span.Width,
                        span.FontSize,
                        span.GlyphSpan.Typeface,
                        span.GlyphSpan.NaturalWidth,
                        span.GlyphSpan.LayoutWidth,
                        span.GlyphSpan.GlyphCount,
                        span.GlyphSpan.FirstAdjustmentAfterOrigin,
                        minAdjustment = span.GlyphSpan.Glyphs.Count == 0 ? 0d : span.GlyphSpan.Glyphs.Min(glyph => glyph.AdjustmentBefore),
                        maxAdjustment = span.GlyphSpan.Glyphs.Count == 0 ? 0d : span.GlyphSpan.Glyphs.Max(glyph => glyph.AdjustmentBefore),
                        avgPositiveAdjustment = AveragePositiveAdjustment(span.GlyphSpan.Glyphs),
                        avgNegativeAdjustment = AverageNegativeAdjustment(span.GlyphSpan.Glyphs)
                    })
                })
            })
        });

        File.WriteAllText(output, JsonSerializer.Serialize(new
        {
            input = Path.GetFileName(input),
            slideIndex,
            frameCount = layout.Frames.Count,
            scene = sceneSummary,
            glyphRuns = glyphRuns.Select(run => new
            {
                length = run.Text.Length,
                hash = ShortHash(run.Text),
                run.X,
                run.BaselineY,
                run.Width,
                run.GlyphCount,
                run.FirstAdjustmentAfterOrigin,
                distinctTypefaces = run.Glyphs
                    .Select(glyph => glyph.Typeface ?? string.Empty)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count(),
                distinctResources = run.Glyphs
                    .Select(glyph => glyph.ResourceName)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count()
            }),
            frames
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static string ShortHash(string text)
    {
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(text));
        return Convert.ToHexString(hash, 0, 4);
    }

    private static double AveragePositiveAdjustment(IReadOnlyList<PptxTextGlyphLayoutSnapshot> glyphs)
    {
        double[] values = glyphs.Select(glyph => glyph.AdjustmentBefore).Where(value => value > 0d).ToArray();
        return values.Length == 0 ? 0d : values.Average();
    }

    private static double AverageNegativeAdjustment(IReadOnlyList<PptxTextGlyphLayoutSnapshot> glyphs)
    {
        double[] values = glyphs.Select(glyph => glyph.AdjustmentBefore).Where(value => value < 0d).ToArray();
        return values.Length == 0 ? 0d : values.Average();
    }

    public static void PptxSyntheticStyledTextProducesStyleOperators()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr/><a:lstStyle/>
                      <a:p>
                        <a:pPr algn="ctr"/>
                        <a:r>
                          <a:rPr sz="2400" b="1" i="1" u="sng"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr>
                          <a:t>Styled text</a:t>
                        </a:r>
                      </a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains(" re f", pdf);
        TestAssert.True(CountOccurrences(pdf, " TJ") >= 1, "Expected styled text to be emitted.");
    }

    public static void PptxSyntheticThemeColorsAndFontsResolve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Theme">
                  <a:themeElements>
                    <a:clrScheme name="Theme">
                      <a:dk1><a:srgbClr val="111111"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent1><a:srgbClr val="FF0000"/></a:accent1>
                    </a:clrScheme>
                    <a:fontScheme name="Theme">
                      <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
                      <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
                    </a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
                    </p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="+mn-lt"/><a:solidFill><a:schemeClr val="dk1"/></a:solidFill></a:rPr><a:t>Theme</a:t></a:r></a:p></p:txBody>
                  </p:sp><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:schemeClr val="bg1"><a:lumMod val="65000"/></a:schemeClr></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("0.651 0.651 0.651 rg", pdf);
        TestAssert.Contains("0.067 0.067 0.067 rg", pdf);
    }

    public static void PptxSyntheticThemeCanLoadFromSlideMaster()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>
                """,
            ["ppt/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Theme">
                  <a:themeElements>
                    <a:clrScheme name="Theme">
                      <a:dk1><a:srgbClr val="111111"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                    </a:clrScheme>
                    <a:fontScheme name="Theme">
                      <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
                      <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
                    </a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>Theme</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 1 1 rg", pdf);
    }

    public static void PptxSyntheticLayoutAndMasterShapesRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = InheritedShapePart("FF0000", 914400, 914400),
            ["ppt/slideLayouts/slideLayout1.xml"] = InheritedShapePart("00FF00", 1828800, 1828800),
            ["ppt/slides/slide1.xml"] = InheritedShapePart("0000FF", 2743200, 2743200)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("0 1 0 rg", pdf);
        TestAssert.Contains("0 0 1 rg", pdf);
    }

    public static void PptxSyntheticLayoutPictureUsesLayoutRelationships()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """),
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/layout-image.png"/>
                </Relationships>
                """),
            ["ppt/slideMasters/slideMaster1.xml"] = TestFixtures.Utf8(InheritedShapePart("FF0000", 914400, 914400)),
            ["ppt/slideLayouts/slideLayout1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="1828800" y="1828800"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sldLayout>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(InheritedShapePart("0000FF", 2743200, 2743200)),
            ["ppt/media/layout-image.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticInheritedPlaceholderTextIsSkipped()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Body Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(!pdf.Contains("/Subtype /Type0", StringComparison.Ordinal), "Inherited placeholder text should not create PDF font resources.");
    }

    public static void PptxSyntheticSlidePlaceholderTextUsesInheritedBounds()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
                  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4000"/></a:lvl1pPr></p:titleStyle></p:txStyles>
                </p:sldMaster>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Layout title</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="3" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Slide title</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/F1 40 Tf", pdf);
        AssertContainsTextMatrixAtX(pdf, 79.2d);
    }

    public static void PptxSyntheticPngPictureRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/XObject", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
    }

    public static void PptxSyntheticSvgPictureRendersVectorPath()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="svg" ContentType="image/svg+xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.svg"] = TestFixtures.Utf8("""
                <svg viewBox="0 0 20 10" xmlns="http://www.w3.org/2000/svg">
                  <path d="M0 0 L20 0 L20 10 L0 10 Z" fill="#112233"/>
                </svg>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.067 0.133 0.2 rg", pdf);
        TestAssert.Contains(" m", pdf);
        TestAssert.Contains(" l", pdf);
        TestAssert.DoesNotContain("/Subtype /Image", pdf);
    }

    public static void PptxSyntheticSvgPictureSamplesCompoundPathGradient()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="svg" ContentType="image/svg+xml"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.svg"] = TestFixtures.Utf8("""
                <svg viewBox="0 0 100 20" xmlns="http://www.w3.org/2000/svg">
                  <linearGradient id="g" x1="0" y1="0" x2="100" y2="0" gradientUnits="userSpaceOnUse">
                    <stop offset="0" stop-color="#FF0000"/>
                    <stop offset="1" stop-color="#0000FF"/>
                  </linearGradient>
                  <path d="M0 0 H20 V20 H0 Z M80 0 H100 V20 H80 Z" fill="url(#g)"/>
                </svg>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.988 0 0.012 rg", pdf);
        TestAssert.Contains("0.012 0 0.988 rg", pdf);
    }

    public static void PptxSyntheticPngPictureAppliesLuminanceRecolor()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"><a:lum bright="70000" contrast="-70000"/></a:blip></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [32, 64, 96])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(CountOccurrences(pdf, "/Subtype /Image") >= 2, "The same image part with and without recolor must use distinct cached image XObjects.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_IMAGE_RECOLOR"), "Supported PNG luminance recolor should not emit unsupported diagnostics.");
    }

    public static void PptxSyntheticPngPictureAppliesDuotoneRecolor()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"><a:duotone><a:srgbClr val="000000"/><a:prstClr val="white"/></a:duotone></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [128, 128, 128])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_IMAGE_RECOLOR"), "Supported PNG duotone recolor should not emit unsupported diagnostics.");
    }

    public static void PptxSyntheticJpegPictureDuotoneRecolorEmitsDiagnostic()
    {
        byte[] jpegHeader =
        [
            0xFF, 0xD8,
            0xFF, 0xC0,
            0x00, 0x11,
            0x08,
            0x00, 0x03,
            0x00, 0x05,
            0x03,
            0x01, 0x11, 0x00,
            0x02, 0x11, 0x00,
            0x03, 0x11, 0x00,
            0xFF, 0xD9
        ];
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="jpeg" ContentType="image/jpeg"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.jpeg"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill>
                      <a:blip r:embed="rId1">
                        <a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone>
                        <a:alphaModFix amt="50000"/>
                      </a:blip>
                    </p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.jpeg"] = jpegHeader
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/DCTDecode", pdf);
        TestAssert.True(diagnostics.Any(d =>
            d.Id == "PPTX_UNSUPPORTED_IMAGE_RECOLOR" &&
            d.Feature == "image recolor" &&
            d.Fallback == "Original image"), "JPEG duotone recolor should keep a public diagnostic until a structural JPEG/PDF recolor path exists.");
    }

    public static void PptxSyntheticPngPictureAppliesGrayAndBilevelRecolor()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"><a:grayscl/></a:blip></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"><a:biLevel thresh="60000"/></a:blip></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [128, 16, 240])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(CountOccurrences(pdf, "/Subtype /Image") >= 2, "Distinct grayscale and bi-level recolors should use distinct cached image XObjects.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_IMAGE_RECOLOR"), "Supported PNG gray/bi-level recolor should not emit unsupported diagnostics.");
    }

    public static void PptxSyntheticShapePictureFillRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:blipFill><a:blip r:embed="rId1"/></a:blipFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticRgbaPngPicturePreservesSoftMask()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbaPng(2, 1, [255, 0, 0, 255, 0, 0, 255, 64])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/SMask", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticEllipsePictureFillClipsImage()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="ellipse"/>
                      <a:blipFill><a:blip r:embed="rId1"/></a:blipFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("W n", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticBmpPictureRendersImageXObject()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="bmp" ContentType="image/bmp"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.bmp"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.bmp"] = TestFixtures.CreateRgbBmp(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
        TestAssert.Contains("/Width 2 /Height 1", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "IMAGE_UNSUPPORTED_FORMAT"), "BMP should render instead of emitting unsupported image diagnostics.");
    }

    public static void PptxSyntheticSlideShapesRenderAbovePictures()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int image = pdf.IndexOf("/Im1 Do", StringComparison.Ordinal);
        int stroke = pdf.IndexOf(" re S", StringComparison.Ordinal);
        TestAssert.True(image >= 0 && stroke > image, "Expected slide shape stroke to be emitted after the image draw.");
    }

    public static void PptxSyntheticSlidePicturesRenderAboveEarlierShapes()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:sp><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>
                    <p:pic>
                      <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    </p:pic>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int fill = pdf.IndexOf("1 0 0 rg", StringComparison.Ordinal);
        int image = pdf.IndexOf("/Im1 Do", StringComparison.Ordinal);
        TestAssert.True(fill >= 0 && image > fill, "Expected a later picture to be emitted after the earlier shape fill.");
    }

    public static void PptxUnsupportedPngImageEmitsDiagnostic()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateUnsupportedHighBitDepthPng()
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        TestAssert.True(File.Exists(output), "Unsupported image should not fail the whole conversion.");
        TestAssert.True(collector.Diagnostics.Any(d => d.Id == "IMAGE_UNSUPPORTED_FORMAT" && d.Severity == OoxPdfSeverity.Error), "Unsupported image should emit a release-blocking diagnostic.");
    }

    public static void PptxSyntheticCroppedPictureUsesClipping()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:blipFill><a:blip r:embed="rId1"/><a:srcRect l="25000" r="25000"/></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(2, 1, [255, 0, 0, 0, 0, 255])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" re W n", pdf);
        TestAssert.Contains("/Im1 Do", pdf);
    }

    public static void PptxSyntheticGroupedShapeAppliesTransform()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:grpSp>
                      <p:grpSpPr>
                        <a:xfrm>
                          <a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/>
                          <a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/>
                        </a:xfrm>
                      </p:grpSpPr>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                        </p:spPr>
                      </p:sp>
                    </p:grpSp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 396 72 72 re f", pdf);
    }

    public static void PptxSyntheticRotatedGroupRotatesChildShape()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:grpSp>
                      <p:grpSpPr>
                        <a:xfrm rot="5400000">
                          <a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/>
                          <a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="1828800"/>
                        </a:xfrm>
                      </p:grpSpPr>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="685800" y="228600"/><a:ext cx="457200" cy="1371600"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill>
                        </p:spPr>
                      </p:sp>
                    </p:grpSp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 -1 1 0", pdf);
        TestAssert.Contains("36 108 re f", pdf);
    }

    public static void PptxSyntheticGroupedPictureAppliesTransform()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
                  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:grpSp>
                      <p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
                      <p:pic>
                        <p:blipFill><a:blip r:embed="rId1"/></p:blipFill>
                        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
                      </p:pic>
                    </p:grpSp>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("144 0 0 144 72 324 cm", pdf);
    }

    public static void PptxSyntheticGroupedTextAppliesTransform()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:grpSp>
                      <p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="3657600" cy="914400"/></a:xfrm></p:grpSpPr>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="0" y="0"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                        </p:spPr>
                        <p:txBody>
                          <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" anchor="t"/>
                          <a:lstStyle/>
                          <a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Grouped</a:t></a:r></a:p>
                        </p:txBody>
                      </p:sp>
                    </p:grpSp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 0 288 540 re W n", pdf);
        AssertContainsTextMatrixAtX(pdf, 72d);
    }

    public static void PptxSyntheticTextAndShapesUseSiblingOrder()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
                      <p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Covered</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int textIndex = pdf.IndexOf(" TJ", StringComparison.Ordinal);
        int coverIndex = pdf.IndexOf("72 396 288 72 re f", StringComparison.Ordinal);
        TestAssert.True(textIndex >= 0 && coverIndex > textIndex, "Expected the covering shape to be emitted after the text box.");
    }

    public static void PptxSyntheticUnknownGraphicFrameDoesNotDisableSiblingOrder()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
                      <p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Covered</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="urn:vendor:unknown-graphic"/></a:graphic>
                    </p:graphicFrame>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int textIndex = pdf.IndexOf(" TJ", StringComparison.Ordinal);
        int coverIndex = pdf.IndexOf("72 396 288 72 re f", StringComparison.Ordinal);
        TestAssert.True(textIndex >= 0 && coverIndex > textIndex, "Expected unknown graphic frames to be ignored without reverting known siblings to fallback paint order.");
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_GRAPHIC_FRAME" && d.PartName == "/ppt/slides/slide1.xml"), "Unknown graphic frame should still emit a slide-scoped diagnostic.");
    }

    public static void PptxSyntheticTableRendersGridAndText()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>
                      <a:graphic>
                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                            <a:tr h="914400">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>One</a:t></a:r></a:p></a:txBody>
                                <a:tcPr>
                                  <a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill>
                                  <a:lnL><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL>
                                  <a:lnR><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnR>
                                  <a:lnT><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnT>
                                  <a:lnB><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnB>
                                </a:tcPr>
                              </a:tc>
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Two</a:t></a:r></a:p></a:txBody>
                                <a:tcPr/>
                              </a:tc>
                            </a:tr>
                            <a:tr h="914400">
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Three</a:t></a:r></a:p></a:txBody>
                                <a:tcPr/>
                              </a:tc>
                              <a:tc>
                                <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Four</a:t></a:r></a:p></a:txBody>
                                <a:tcPr><a:solidFill><a:srgbClr val="FCE5CD"/></a:solidFill></a:tcPr>
                              </a:tc>
                            </a:tr>
                          </a:tbl>
                        </a:graphicData>
                      </a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.851 0.918 0.827 rg", pdf);
        TestAssert.Contains("72 396 144 72 re f", pdf);
        TestAssert.Contains("0 0 0 RG", pdf);
        TestAssert.Contains("/Subtype /Type0", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxSyntheticGroupedTableUsesGroupTransform()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:grpSp>
                      <p:nvGrpSpPr><p:cNvPr id="1" name="Table group"/><p:nvPr/></p:nvGrpSpPr>
                      <p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>
                      <p:graphicFrame>
                        <p:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                            <a:tr h="914400">
                              <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr></a:tc>
                            </a:tr>
                          </a:tbl>
                        </a:graphicData></a:graphic>
                      </p:graphicFrame>
                    </p:grpSp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 396 144 72 re f", pdf);
    }

    public static void PptxSyntheticGroupedTableUsesGroupTransformWithUnknownGraphicFrame()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="urn:unknown"/></a:graphic>
                    </p:graphicFrame>
                    <p:grpSp>
                      <p:nvGrpSpPr><p:cNvPr id="1" name="Table group"/><p:nvPr/></p:nvGrpSpPr>
                      <p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></p:grpSpPr>
                      <p:graphicFrame>
                        <p:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                          <a:tbl>
                            <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                            <a:tr h="914400">
                              <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr></a:tc>
                            </a:tr>
                          </a:tbl>
                        </a:graphicData></a:graphic>
                      </p:graphicFrame>
                    </p:grpSp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 396 144 72 re f", pdf);
    }

    public static void PptxSyntheticTableStyleMediumStyle2Accent6RendersFills()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes().Replace(
                "</Types>",
                "  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\r\n</Types>",
                StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
            ["ppt/theme/theme1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="TableTheme">
                  <a:themeElements>
                    <a:clrScheme name="TableTheme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent6><a:srgbClr val="336699"/></a:accent6>
                    </a:clrScheme>
                    <a:fontScheme name="TableTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblPr firstRow="1" firstCol="1" lastRow="1" lastCol="1" bandRow="1"><a:tableStyleId>{93296810-A885-4BE3-A3E7-6D5BEEA58F35}</a:tableStyleId></a:tblPr>
                          <a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>
                          <a:tr h="609600">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>H1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>H2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>H3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="609600">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>C</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="609600">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>G</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>H</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>I</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="609600">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>D</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>E</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>F</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="4114800" y="914400"/><a:ext cx="1828800" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblPr firstRow="1"><a:bandCol/><a:tableStyleId>{68D230F3-CF80-4859-8CE7-A43EE81993B5}</a:tableStyleId></a:tblPr>
                          <a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>L1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>L2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>L3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>L4</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="6400800" y="914400"/><a:ext cx="1371600" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                        <a:tbl>
                          <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{AF606853-7671-496A-8E4F-DF71F8EC918B}</a:tableStyleId></a:tblPr>
                          <a:tblGrid><a:gridCol w="1371600"/></a:tblGrid>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>D1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                          <a:tr h="914400">
                            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>D2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          </a:tr>
                        </a:tbl>
                      </a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.2 0.4 0.6 rg", pdf);
        TestAssert.Contains("0.761 0.827 0.89 rg", pdf);
        TestAssert.Contains("0.886 0.918 0.949 rg", pdf);
        TestAssert.Contains("1 1 1 rg", pdf);
        TestAssert.True(Regex.Matches(pdf, "0\\.2 0\\.4 0\\.6 rg").Count >= 3, "Expected header cells and first-column body cell to use the accent fill.");
        TestAssert.Contains("/GS40000F100000S gs", pdf);
        TestAssert.Contains("0.078 0.161 0.239 rg", pdf);
    }

    public static void PptxSyntheticTableWrapsCellTextToColumnWidth()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="914400"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Alpha Beta Gamma</a:t></a:r></a:p></a:txBody>
                            <a:tcPr/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int lineStarts = Regex.Matches(pdf, $@"1 0 0 1 {Regex.Escape(FormatPdfNumber(79.2d))} [0-9.]+ Tm").Count;
        TestAssert.True(lineStarts >= 2, "Expected narrow table-cell text to wrap onto multiple lines at the cell text inset.");
        TestAssert.Contains("79.2 399.6 57.6 64.8 re W n", pdf);
    }

    public static void PptxSyntheticTableKeepsSlide6HeaderOnOneLine()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="2453469" cy="700198"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="2453469"/></a:tblGrid>
                        <a:tr h="700198">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="1200" b="1"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/><a:ea typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr><a:t>Recurring decision Lokad automates</a:t></a:r></a:p></a:txBody>
                            <a:tcPr marL="46789" marR="46789" marT="23394" marB="23394" anchor="ctr"/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string[] matrices = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ [0-9.]+ Tm")
            .Select(match => match.Value)
            .ToArray();
        int lineBaselines = matrices
            .Select(matrix => Regex.Match(matrix, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm").Groups["y"].Value)
            .Distinct(StringComparer.Ordinal)
            .Count();
        TestAssert.True(lineBaselines == 1, $"Expected one rendered baseline for the table header; got {lineBaselines}. Matrices: {string.Join(" | ", matrices)}");
        TestAssert.True(Regex.IsMatch(pdf, @"1 0 0 1 75\.684 436\.065 Tm"), $"The centered table header should measure and render with the same table wrap width. Matrices: {string.Join(" | ", matrices)}");
    }

    public static void PptxSyntheticTableIgnoresLeadingEmptyCellParagraph()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1585169" cy="481032"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1585169"/></a:tblGrid>
                        <a:tr h="481032">
                          <a:tc>
                            <a:txBody>
                              <a:bodyPr/><a:lstStyle/>
                              <a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" sz="1200" b="0"><a:latin typeface="Cambria Math"/><a:ea typeface="Cambria Math"/></a:endParaRPr></a:p>
                              <a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1200" b="0"><a:latin typeface="Cambria Math"/><a:ea typeface="Cambria Math"/></a:rPr><a:t>Long-term Scheduling, Production</a:t></a:r></a:p>
                            </a:txBody>
                            <a:tcPr marL="35667" marR="35667" marT="17833" marB="17833" anchor="ctr"/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string[] baselines = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm")
            .Select(match => match.Groups["y"].Value)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        TestAssert.Equal(2, baselines.Length);
    }

    public static void PptxSyntheticTableCentersTextByContentHeight()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="0"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc>
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Centered</a:t></a:r></a:p></a:txBody>
                            <a:tcPr anchor="ctr" marL="0" marR="0" marT="0" marB="0"/>
                          </a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        string[] matrices = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ [0-9.]+ Tm")
            .Select(match => match.Value)
            .ToArray();
        TestAssert.True(Regex.IsMatch(pdf, @"1 0 0 1 72 499\.667 Tm"), $"Centered table-cell text should account for its line height before vertical anchoring. Matrices: {string.Join(" | ", matrices)}");
    }

    public static void PptxSyntheticTableMergedCellsSuppressInteriorGrid()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            return;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="914400">
                          <a:tc gridSpan="2">
                            <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Wide</a:t></a:r></a:p></a:txBody>
                            <a:tcPr><a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill></a:tcPr>
                          </a:tc>
                          <a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>
                        </a:tr>
                        <a:tr h="914400">
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Left</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                          <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>Right</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                        </a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 396 288 72 re f", pdf);
        TestAssert.Contains("216 396.5 m 216 323.5 l S", pdf);
        TestAssert.DoesNotContain("216 468.5 m 216 396.5 l S", pdf);
    }

    public static void PptxSyntheticBarChartsRenderNativeCharts()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes().Replace(
                "</Types>",
                "  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\r\n</Types>",
                StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        TestAssert.Contains("0.933 0.933 0.933 rg", pdf);
        TestAssert.Contains("0.184 0.522 0.416 RG", pdf);
        TestAssert.Contains(" re W n", pdf);
        TestAssert.Contains("BT", pdf);
        TestAssert.Contains(" re f", pdf);
        TestAssert.True(pdf.Split("BT", StringSplitOptions.None).Length >= 17, "Combo bar charts should emit text for primary and secondary labels, axes, and legend entries.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Supported chart rendering should not emit static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Supported bar charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticGroupedBarChartUsesGroupTransform()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
                    <p:grpSp>
                      <p:nvGrpSpPr><p:cNvPr id="2" name="Grouped chart"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                      <p:grpSpPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2743200"/><a:chOff x="0" y="0"/><a:chExt cx="3657600" cy="2743200"/></a:xfrm></p:grpSpPr>
                      <p:graphicFrame>
                        <p:nvGraphicFramePr><p:cNvPr id="3" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
                        <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                      </p:graphicFrame>
                    </p:grpSp>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:chart><c:plotArea><c:spPr><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill></c:spPr><c:barChart>
                    <c:ser><c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser>
                    <c:gapWidth val="300"/>
                  </c:barChart></c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 1 1 rg", pdf);
        TestAssert.Contains("152.294 330.653 133.517 62.683 re f", pdf);
        TestAssert.Contains("202.363 330.653 33.379 62.683 re f", pdf);
        TestAssert.DoesNotContain("160.128 326.52 123.84 63.72 re f", pdf);
        TestAssert.DoesNotContain("86.4 406.08 118.08 51.84 re f", pdf);
    }

    public static void PptxSyntheticChartAxisTickLabelOptionsRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        TestAssert.Contains("364.922", pdf);
    }

    public static void PptxSyntheticChartCategoryAxisLabelOffsetRender()
    {
        static string RenderPdf(int labelOffset)
        {
            string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
            {
                ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
                ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
                ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
                ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        double defaultBaseline = ReadOnlyTextBaseline(defaultPdf);
        double offsetBaseline = ReadOnlyTextBaseline(offsetPdf);

        TestAssert.Contains("1 0 0 rg", defaultPdf);
        TestAssert.Contains("1 0 0 rg", offsetPdf);
        TestAssert.True(offsetBaseline < defaultBaseline - 5d, $"Expected lblOffset=200 to move the category label farther from the plot area. Default: {defaultBaseline.ToString(CultureInfo.InvariantCulture)}; offset: {offsetBaseline.ToString(CultureInfo.InvariantCulture)}.");
    }

    public static void PptxSyntheticChartCategoryAxisTickLabelSkipRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        TestAssert.Equal(2, CountTextMatrices(pdf));
    }

    public static void PptxSyntheticChartCategoryGridlinesDoNotEnableValueGridlines()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        TestAssert.Contains("0 0 0 RG", pdf);
        TestAssert.Contains("0.75 w", pdf);
        TestAssert.DoesNotContain("0.851 0.851 0.851 RG", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"(?:[0-9.]+ [0-9.]+ m\s+[0-9.]+ [0-9.]+ l\s+){3}S"),
            "Major value gridlines should be emitted as one multi-segment stroked path that includes the maximum tick and excludes the axis baseline.");
    }

    public static void PptxSyntheticChartHorizontalValueAxisAutoUnitUsesOfficeDenseTicks()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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

        double lineMarkerMax = (double)method!.Invoke(null, [96d, 0d, 8d, true, 0.96d])!;
        double ordinaryMax = (double)method.Invoke(null, [96d, 0d, 8d, false, 0.96d])!;
        double lineThreeSeriesMax = (double)method.Invoke(null, [1520d, 0d, 8d, true, 0.96d])!;

        TestAssert.Equal(120d, lineMarkerMax);
        TestAssert.Equal(100d, ordinaryMax);
        TestAssert.Equal(1600d, lineThreeSeriesMax);
    }

    public static void PptxSyntheticChartValueGridlinesExcludeCrossingTick()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            @"0 0 1 RG\s+[0-9.]+ w\s+\[\] 0 d\s+0 J\s+0 j\s+(?<x1>[0-9.]+) (?<y1>[0-9.]+) m (?<x2>[0-9.]+) (?<y2>[0-9.]+) l S");
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        var segmentRectangles = Regex.Matches(pdf, @"(?<x>[0-9.]+) (?<y>[0-9.]+) (?<w>[0-9.]+) (?<h>[0-9.]+) re\s+f")
            .Cast<Match>()
            .Select(match => new
            {
                Y = double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture),
                Width = double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture),
                Height = double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture)
            })
            .Where(rectangle => rectangle.Width > 10d && rectangle.Width < 220d && rectangle.Height > 20d && rectangle.Height < 220d)
            .ToArray();
        TestAssert.True(segmentRectangles.Length >= 2, "Expected two filled stacked column segment rectangles.");
        TestAssert.True(segmentRectangles[1].Y < segmentRectangles[0].Y,
            "Stacked column segments should accumulate through mapped value intervals when the value axis is maxMin.");
    }

    public static void PptxSyntheticChartValueAxisReversedOrientationStacksHorizontalBarSegmentsByValue()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        var segmentRectangles = Regex.Matches(pdf, @"(?<x>[0-9.]+) (?<y>[0-9.]+) (?<w>[0-9.]+) (?<h>[0-9.]+) re\s+f")
            .Cast<Match>()
            .Select(match => new
            {
                X = double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Width = double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture),
                Height = double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture)
            })
            .Where(rectangle => rectangle.Width > 20d && rectangle.Width < 220d && rectangle.Height > 10d && rectangle.Height < 220d)
            .ToArray();
        TestAssert.True(segmentRectangles.Length >= 2, "Expected two filled stacked horizontal bar segment rectangles.");
        TestAssert.True(segmentRectangles[1].X < segmentRectangles[0].X,
            "Stacked horizontal bar segments should accumulate through mapped value intervals when the value axis is maxMin.");
    }

    public static void PptxSyntheticChartManualLayoutEdgeModesRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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

    public static void PptxSyntheticChartTitleManualBoxDrivesPlacement()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        TestAssert.True(Regex.IsMatch(pdf, @"1 0 0 1 [0-9.]+ 354\.816 Tm"), "Expected explicit chart title manualLayout to drive the title text baseline.");
        TestAssert.True(!Regex.IsMatch(pdf, @"1 0 0 1 [0-9.]+ 442\.[0-9]+ Tm"), "Expected chart title rendering not to fall back to the full-frame title box.");
    }

    public static void PptxSyntheticChartLegendManualBoxDrivesPlacement()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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

    public static void PptxSyntheticChartManualLayoutMissingPositionUsesDefaultPlotBox()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        TestAssert.Contains("88.589 271.958 267.034 188.05 re f", pdf);
        TestAssert.DoesNotContain("104.256 259.56 247.68 191.16 re f", pdf);
        TestAssert.DoesNotContain("100.8 282.24 236.16 174.96 re f", pdf);
    }

    public static void PptxSyntheticChartAutoTitleUsesOfficeScaledSize()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);
        PptxScene scene = new PptxSceneBuilder().Build(document, package);
        PptxSceneChartTitle? sceneTitle = scene.Slides[0].SlideNodes[0].Chart?.Title;
        TestAssert.Equal("Series", sceneTitle?.Text ?? string.Empty);
        TestAssert.True(sceneTitle?.IsAutoGenerated == true, "Expected the typed scene chart model to own auto-title inference.");
        TestAssert.True(sceneTitle?.IsAutoDeleted == false, "Expected explicit autoTitleDeleted=false metadata to remain distinct from a missing element.");

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.IsMatch(pdf, @"/CT[0-9]+ 21\.6 Tf"), "Expected absent auto chart title to use Office's 120% title text scale over chart txPr.");
        TestAssert.True(Regex.IsMatch(pdf, @"/C[AV][A-Z][0-9]+ 18 Tf"), "Expected chart axis text to keep the unscaled chart txPr size.");
    }

    public static void PptxChartAutoTitleUsesSingleSeriesNameForLineCharts()
    {
        PptxSceneChart? chart = BuildSingleChartScene("""
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

    public static void PptxScenePreservesChartSeriesDataSourceReferences()
    {
        PptxSceneChart? chart = BuildSingleChartScene("""
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
        TestAssert.Equal("strRef", lineSeries.DataSources.Name.ReferenceKind);
        TestAssert.Equal("strCache", lineSeries.DataSources.Name.CacheKind);
        TestAssert.True(!lineSeries.DataSources.Name.HasCachedPoints, "Expected empty string cache to preserve the formula without inventing cached points.");
        TestAssert.Equal("Sheet1!$A$2:$A$4", lineSeries.DataSources.Categories.Formula ?? string.Empty);
        TestAssert.Equal("multiLvlStrRef", lineSeries.DataSources.Categories.ReferenceKind);
        TestAssert.Equal("multiLvlStrCache", lineSeries.DataSources.Categories.CacheKind);
        TestAssert.Equal("Sheet1!$B$2:$B$4", lineSeries.DataSources.Values.Formula ?? string.Empty);
        TestAssert.Equal("numRef", lineSeries.DataSources.Values.ReferenceKind);
        TestAssert.Equal("numCache", lineSeries.DataSources.Values.CacheKind);
        TestAssert.True(!lineSeries.DataSources.Values.HasCachedPoints, "Expected formula-only numeric references to remain distinguishable from cached values.");

        PptxSceneChartSeries scatterSeries = chart?.Plots[1].Series[0] ?? throw new InvalidOperationException("Expected scatter chart series.");
        TestAssert.Equal("Sheet2!$A$2:$A$4", scatterSeries.DataSources.XValues.Formula ?? string.Empty);
        TestAssert.True(scatterSeries.DataSources.XValues.HasCachedPoints, "Expected x-value cache point presence in the scene model.");
        TestAssert.Equal("Sheet2!$B$2:$B$4", scatterSeries.DataSources.YValues.Formula ?? string.Empty);
        TestAssert.True(scatterSeries.DataSources.YValues.HasCachedPoints, "Expected y-value cache point presence in the scene model.");
        TestAssert.Equal("Sheet2!$C$2:$C$4", scatterSeries.DataSources.BubbleSizes.Formula ?? string.Empty);
        TestAssert.True(!scatterSeries.DataSources.BubbleSizes.HasCachedPoints, "Expected empty bubble-size cache to remain explicit.");
    }

    public static void PptxChartAutoTitleDeletedSuppressesSingleSeriesName()
    {
        PptxSceneChart? chart = BuildSingleChartScene("""
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
    }

    public static void PptxChartAutoTitleDoesNotInventAmbiguousMultiSeriesTitle()
    {
        PptxSceneChart? chart = BuildSingleChartScene("""
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
    }

    public static void PptxSyntheticLineAndPieChartsRenderNativeCharts()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
                  <c:spPr><a:solidFill><a:srgbClr val="F0F0F0"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="444444"/></a:solidFill></a:ln></c:spPr>
                  <c:chart>
                  <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1300" b="1" i="0"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:title>
                  <c:plotArea>
                  <c:layout><c:manualLayout><c:x val="0.2"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.5"/></c:manualLayout></c:layout>
                  <c:spPr><a:solidFill><a:srgbClr val="00FFFF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
                  <c:valAx><c:delete/><c:scaling><c:min val="0"/><c:max val="10"/></c:scaling><c:majorUnit val="2"/><c:minorUnit val="1"/><c:majorGridlines/><c:minorGridlines/><c:spPr><a:ln><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln></c:spPr></c:valAx><c:lineChart>
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
                    <c:dLbls><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/><c:dLblPos val="b"/><c:separator> | </c:separator><c:spPr><a:solidFill><a:srgbClr val="FFEACC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="0A0B0C"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr><c:dLbl><c:idx val="1"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>ZXQ</a:t></a:r></a:p></c:rich></c:tx><c:dLblPos val="ctr"/><c:spPr><a:solidFill><a:srgbClr val="CCEEFF"/></a:solidFill></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="6600CC"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbl></c:dLbls>
                  </c:lineChart></c:plotArea><c:legend><c:legendPos val="b"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="700" b="0" i="1"><a:solidFill><a:srgbClr val="654321"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:legend></c:chart>
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
                    </c:numLit></c:val><c:dLbls><c:showVal val="1"/><c:showPercent val="1"/></c:dLbls></c:ser>
                  </c:pieChart></c:plotArea></c:chart>
                </c:chartSpace>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0.941 0.941 0.941 rg", pdf);
        TestAssert.Contains("0.267 0.267 0.267 RG", pdf);
        TestAssert.Contains("0 1 1 rg", pdf);
        TestAssert.Contains("93.6 352.8 144 108 re f", pdf);
        TestAssert.Contains("93.6 374.4 m", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"(?:[0-9.]+ [0-9.]+ m\s+[0-9.]+ [0-9.]+ l\s+){5}S"),
            "Line-chart major gridlines should share one stroked path for the five non-baseline tick marks.");
        TestAssert.Contains("1 0 0 RG", pdf);
        TestAssert.Contains("BT", pdf);
        TestAssert.True(pdf.Split("BT", StringSplitOptions.None).Length >= 7, "Chart title, axes, legend, and data labels should emit chart text objects.");
        TestAssert.Contains("0.922 0.922 0.922 RG", pdf);
        TestAssert.Contains("0 0 0 RG", pdf);
        TestAssert.DoesNotContain("0 0.667 0 RG", pdf);
        TestAssert.Contains("0.667 0 0.667 RG", pdf);
        TestAssert.Contains("8 8 re f", pdf);
        TestAssert.Contains("0.667 0.333 0 RG", pdf);
        TestAssert.Contains("0.333 0.333 0.333 RG", pdf);
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
        TestAssert.Contains("151.2 400.725 28.8 12.15 re f", pdf);
        TestAssert.Contains("0.4 0 0.8 rg", pdf);
        TestAssert.Contains("/CLD1 9 Tf", pdf);
        TestAssert.Contains("0.071 0.204 0.337 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CT[0-9]+ 13 Tf"), "Expected chart title txPr font size to drive title rendering.");
        TestAssert.Contains("0.396 0.263 0.129 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CL[0-9]+ 7 Tf"), "Expected chart legend txPr font size to drive legend rendering.");
        TestAssert.True(Regex.IsMatch(pdf, @"1 0 0 1 [0-9.]+ 400\.725 Tm"), "Expected the centered custom data label to keep the Office-derived label-box baseline.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <005A>") &&
            Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0058>") &&
            Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0051>"),
            "Expected per-label custom rich text to be emitted in the chart data label ToUnicode map.");
        TestAssert.Contains("0 0.4 0.667 rg", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"/CLD[0-9]+ 8 Tf"), "Expected second-series data labels to consume the series-level font size.");
        TestAssert.Contains("/ItalicAngle -12", pdf);
        TestAssert.Contains("0.039 0.043 0.047 rg", pdf);
        TestAssert.Contains("/CLD1 10 Tf", pdf);
        TestAssert.Contains("1 0 0 1 79.2 357.525 Tm", pdf);
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0025>"), "Expected percentage labels to include a percent glyph in the ToUnicode map.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <002C>"), "Expected combined value/percentage labels to include the default comma separator.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Supported chart rendering should not emit static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Line and pie charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxSyntheticAreaScatterRadarAndDoughnutChartsRenderNativeCharts()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
        TestAssert.True(Regex.Matches(pdf, @"[0-9.]+ [0-9.]+ m").Count >= 4, "Expected native chart paths for area, scatter, radar, and doughnut charts.");
        TestAssert.DoesNotContain("/GS62000F100000S gs", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("f*", pdf);
        TestAssert.Contains(" l S", pdf);
        TestAssert.DoesNotContain("1 1 1 rg", pdf);
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Supported chart rendering should not emit static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Area, scatter, radar, and doughnut charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxEmbeddedWorkbookReferencesHydrateMissingChartCaches()
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
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
                  <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:doughnutChart>
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
            ["ppt/embeddings/Microsoft_Excel_Worksheet.xlsx"] = EmbeddedChartWorkbook()
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int pathStartCount = Regex.Matches(pdf, @"[0-9.]+ [0-9.]+ m").Count;
        TestAssert.True(pathStartCount >= 3, $"Expected workbook-backed chart references to produce native doughnut paths; saw {pathStartCount} path starts and diagnostics: {string.Join(", ", collector.Diagnostics.Select(d => d.Id))}.");
        TestAssert.True(CountTextMatrices(pdf) >= 3, "Expected workbook-backed category references to feed native legend text without requiring chart-side string caches.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Workbook-backed chart references should render through the native chart path.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Charts with formulas and embedded workbook data should not require chart-side caches.");
    }

    public static void PptxBubbleChartRendersNativeAxesGridlinesLegendAndBubbles()
    {
        string input = Path.Combine(
            Directory.GetCurrentDirectory(),
            "tests",
            "Lokad.OoxPdf.Tests",
            "Cases",
            "pptx-ladder-11-chart-bubble-port.pptx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var collector = new DiagnosticCollector();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = collector.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(Regex.Matches(pdf, @" c\s").Count >= 16, "Expected native bubble ellipses to be emitted as Bezier curves.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0035>"), "Expected bubble value-axis labels to include the Office 0..5 scale.");
        TestAssert.True(Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0036>"), "Expected bubble value-axis labels to include the Office 0..6 scale.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Bubble charts should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Bubble charts should not emit unsupported chart diagnostics.");
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

        string pdf = ReadPdfDecodedAscii(output);
        TestAssert.True(
            Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <0025>"),
            "Expected percent-stacked value-axis labels to include a percent sign by default.");
        TestAssert.True(
            !Regex.IsMatch(pdf, @"<[0-9A-F]{4}> <002E>"),
            "Percent-stacked value-axis labels should not use the generic 0..1.2 decimal scale.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_CHART_STATIC_FALLBACK"), "Percent-stacked column charts should render without static fallback diagnostics.");
        TestAssert.True(collector.Diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_CHART"), "Percent-stacked column charts should not emit unsupported chart diagnostics.");
    }

    public static void PptxUnsupportedFeaturesEmitDiagnostics()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PresentationRelationship(),
            ["ppt/presentation.xml"] = BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphic>
                    </p:graphicFrame>
                    <p:graphicFrame>
                      <a:graphic><a:graphicData uri="urn:vendor:unknown-graphic"/></a:graphic>
                    </p:graphicFrame>
                    <p:pic><p:nvPicPr><p:cNvPr id="2" name="Video"/><p:cNvPicPr/><p:nvPr><p:video/></p:nvPr></p:nvPicPr><p:blipFill><a:blip><a:videoFile/></a:blip></p:blipFill></p:pic>
                    <p:pic><p:nvPicPr><p:cNvPr id="3" name="Audio"/><p:cNvPicPr/><p:nvPr><p:audio/></p:nvPr></p:nvPicPr><p:blipFill><a:blip><a:audioFile/></a:blip></p:blipFill></p:pic>
                    <p:oleObj/>
                    <p:sp><p:spPr><a:gradFill/></p:spPr></p:sp>
                    <p:sp><p:spPr><a:pattFill/></p:spPr></p:sp>
                    <p:sp><p:spPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></p:spPr></p:sp>
                    <p:sp><p:spPr><a:effectLst><a:reflection/></a:effectLst></p:spPr></p:sp>
                    <p:sp><p:spPr><a:custGeom/></p:spPr></p:sp>
                    <p:sp><p:spPr><a:prstGeom prst="wedgeRoundRectCallout"/></p:spPr></p:sp>
                    <p:sp><p:spPr><a:prstGeom prst="heart"/><a:blipFill><a:blip/></a:blipFill></p:spPr></p:sp>
                    <p:pic><p:blipFill><a:blip><a:grayscl/></a:blip><a:tile/></p:blipFill></p:pic>
                    <p:sp><p:txBody><a:bodyPr vert="vert270" vertOverflow="ellipsis"/><a:lstStyle/><a:p/></p:txBody></p:sp>
                  </p:spTree></p:cSld>
                  <p:transition/>
                  <p:timing/>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string[] ids = diagnostics.Select(d => d.Id).Order(StringComparer.Ordinal).ToArray();
        TestAssert.Equal(16, ids.Length);
        TestAssert.Contains("PPTX_UNSUPPORTED_ANIMATION", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_AUDIO", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_CALLOUT", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_CHART", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_CUSTOM_GEOMETRY", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_EFFECT", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_GRADIENT_FILL", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_GRAPHIC_FRAME", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_IMAGE_TILE", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_OLE_OBJECT", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_PATTERN_FILL", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_PICTURE_FILL", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_SMARTART", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_TEXT_OVERFLOW", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_TRANSITION", string.Join("|", ids));
        TestAssert.Contains("PPTX_UNSUPPORTED_VIDEO", string.Join("|", ids));
        TestAssert.True(diagnostics.All(d => d.Severity == OoxPdfSeverity.Warning && d.SlideIndex == 1), "Unsupported PPTX diagnostics should be slide-scoped warnings.");
    }

    private static PptxSceneChart? BuildSingleChartScene(string chartXml)
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(BasicContentTypes()),
            ["_rels/.rels"] = TestFixtures.Utf8(PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(BasicPresentation()),
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
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8(chartXml)
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream);
        PptxDocument document = new PptxReader().Read(package);
        PptxScene scene = new PptxSceneBuilder().Build(document, package);
        return scene.Slides[0].SlideNodes[0].Chart;
    }

    private static string ReadPdfDecodedAscii(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        string pdf = Encoding.ASCII.GetString(bytes);
        var text = new StringBuilder(pdf);
        int searchStart = 0;
        while (true)
        {
            int streamMarker = pdf.IndexOf("stream", searchStart, StringComparison.Ordinal);
            if (streamMarker < 0)
            {
                break;
            }

            int streamStart = streamMarker + "stream".Length;
            if (streamStart < bytes.Length && bytes[streamStart] == (byte)'\r')
            {
                streamStart++;
            }

            if (streamStart < bytes.Length && bytes[streamStart] == (byte)'\n')
            {
                streamStart++;
            }

            int streamEnd = pdf.IndexOf("endstream", streamStart, StringComparison.Ordinal);
            if (streamEnd < 0)
            {
                break;
            }

            int dictionaryStart = Math.Max(0, pdf.LastIndexOf("<<", streamMarker, StringComparison.Ordinal));
            string dictionary = pdf.Substring(dictionaryStart, streamMarker - dictionaryStart);
            if (dictionary.Contains("/FlateDecode", StringComparison.Ordinal))
            {
                using var input = new MemoryStream(bytes, streamStart, streamEnd - streamStart);
                using var deflate = new ZLibStream(input, CompressionMode.Decompress);
                using var output = new MemoryStream();
                deflate.CopyTo(output);
                text.Append(Encoding.ASCII.GetString(output.ToArray()));
            }

            searchStart = streamEnd + "endstream".Length;
        }

        return text.ToString();
    }

    private static double ReadOnlyTextBaseline(string pdf)
    {
        MatchCollection matches = Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ (?<y>[0-9.]+) Tm");
        TestAssert.Equal(1, matches.Count);
        return double.Parse(matches[0].Groups["y"].Value, CultureInfo.InvariantCulture);
    }

    private static int CountTextMatrices(string pdf)
    {
        return Regex.Matches(pdf, @"1 0 0 1 [0-9.]+ [0-9.]+ Tm").Count;
    }

    private static byte[] EmbeddedChartWorkbook()
    {
        using MemoryStream stream = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """,
            ["xl/_rels/workbook.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
                </Relationships>
                """,
            ["xl/workbook.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """,
            ["xl/sharedStrings.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t>North</t></si>
                  <si><t>South</t></si>
                  <si><t>West</t></si>
                  <si><t>Share</t></si>
                </sst>
                """,
            ["xl/worksheets/sheet1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1"><c r="B1" t="s"><v>3</v></c></row>
                    <row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2"><v>8.2</v></c></row>
                    <row r="3"><c r="A3" t="s"><v>1</v></c><c r="B3"><v>3.2</v></c></row>
                    <row r="4"><c r="A4" t="s"><v>2</v></c><c r="B4"><v>1.4</v></c></row>
                  </sheetData>
                </worksheet>
                """
        });
        return stream.ToArray();
    }

    private static string BasicContentTypes()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
              <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
            </Types>
            """;
    }

    private static string PackageRelationship()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
            </Relationships>
            """;
    }

    private static string PresentationRelationship()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
            </Relationships>
            """;
    }

    private static string BasicPresentation()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:sldSz cx="9144000" cy="6858000"/>
              <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
            </p:presentation>
            """;
    }

    private static string InheritedShapePart(string color, int x, int y)
    {
        return $$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp>
                <p:spPr>
                  <a:xfrm><a:off x="{{x}}" y="{{y}}"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"/>
                  <a:solidFill><a:srgbClr val="{{color}}"/></a:solidFill>
                </p:spPr>
              </p:sp></p:spTree></p:cSld>
            </p:sld>
            """;
    }

    private static int CountOccurrences(string text, string value)
    {
        int count = 0;
        int start = 0;
        while (true)
        {
            int index = text.IndexOf(value, start, StringComparison.Ordinal);
            if (index < 0)
            {
                return count;
            }

            count++;
            start = index + value.Length;
        }
    }

    private static void AssertContainsTextMatrixAtX(string pdf, double x)
    {
        string xText = FormatPdfNumber(x);
        TestAssert.True(Regex.IsMatch(pdf, $@"1 0 0 1 {Regex.Escape(xText)} [0-9.]+ Tm"), $"Expected a text matrix at x={xText}.");
    }

    private static void AssertDoesNotContainTextMatrixAtX(string pdf, double x, string message)
    {
        string xText = FormatPdfNumber(x);
        TestAssert.True(!Regex.IsMatch(pdf, $@"1 0 0 1 {Regex.Escape(xText)} [0-9.]+ Tm"), message);
    }

    private static string FormatPdfNumber(double value)
    {
        return value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
    }
}
