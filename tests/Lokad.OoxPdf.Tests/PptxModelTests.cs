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

internal static class PptxModelTests
{
    public static void PptxSceneBuilderBuildsResolvedNodeLists()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>
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
                  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                  <Relationship Id="rIdShapeImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                  <Relationship Id="rIdShapeHyperlink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/shape" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/charts/_rels/chart1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
                  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/chart-data.xlsx"/>
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
                    <a:fmtScheme name="Test">
                      <a:fillStyleLst>
                        <a:solidFill><a:srgbClr val="010101"/></a:solidFill>
                        <a:solidFill><a:srgbClr val="BBDDEE"><a:alpha val="65000"/></a:srgbClr></a:solidFill>
                      </a:fillStyleLst>
                      <a:lnStyleLst>
                        <a:ln w="25400" cap="sq" cmpd="dbl">
                          <a:solidFill><a:srgbClr val="ABC123"><a:alpha val="60000"/></a:srgbClr></a:solidFill>
                          <a:prstDash val="dot"/>
                          <a:round/>
                        </a:ln>
                      </a:lnStyleLst>
                      <a:effectStyleLst>
                        <a:effectStyle><a:effectLst><a:blur rad="6350"/></a:effectLst></a:effectStyle>
                      </a:effectStyleLst>
                      <a:bgFillStyleLst/>
                    </a:fmtScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/charts/chart1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:date1904 val="true"/>
                  <c:roundedCorners val="0"/>
                  <c:style val="10"/>
                  <c:spPr><a:pattFill prst="pct25"><a:fgClr><a:srgbClr val="224466"/></a:fgClr><a:bgClr><a:srgbClr val="F1E2D3"/></a:bgClr></a:pattFill><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln><a:effectLst><a:outerShdw dist="12700" dir="0"><a:srgbClr val="010203"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst></c:spPr>
                  <c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1" i="1" u="sng" strike="sngStrike"><a:solidFill><a:srgbClr val="101112"><a:alpha val="80000"/></a:srgbClr></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr>
                  <c:chart>
                  <c:plotVisOnly val="false"/>
                  <c:dispBlanksAs val="span"/>
                  <c:showDLblsOverMax val="1"/>
                  <c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:rPr sz="1250" b="1"><a:solidFill><a:srgbClr val="223344"/></a:solidFill></a:rPr><a:t>Scene </a:t></a:r><a:r><a:rPr i="1"><a:latin typeface="Aptos"/></a:rPr><a:t>Chart</a:t></a:r></a:p></c:rich></c:tx><c:layout><c:manualLayout><c:layoutTarget val="outer"/><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0.08"/><c:y val="0.04"/><c:w val="0.55"/><c:h val="0.12"/></c:manualLayout></c:layout><c:overlay val="1"/><c:spPr><a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="0F1E2D"/></a:solidFill></a:ln><a:effectLst><a:glow rad="25400"><a:srgbClr val="0A0B0C"><a:alpha val="40000"/></a:srgbClr></a:glow></a:effectLst></c:spPr><c:txPr><a:bodyPr rot="5400000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1300" b="1" i="0" u="none" strike="noStrike"><a:solidFill><a:srgbClr val="1122AA"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:title>
                  <c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="factor"/><c:yMode val="factor"/><c:wMode val="factor"/><c:hMode val="factor"/><c:x val="0.12"/><c:y val="0.18"/><c:w val="0.72"/><c:h val="0.66"/></c:manualLayout></c:layout><c:spPr><a:noFill/><a:ln w="25400"><a:solidFill><a:srgbClr val="112244"><a:alpha val="60000"/></a:srgbClr></a:solidFill></a:ln></c:spPr><c:barChart>
                    <c:barDir val="bar"/>
                    <c:grouping val="stacked"/>
                    <c:varyColors val="false"/>
                    <c:gapWidth val="175"/>
                    <c:overlap val="25"/>
                    <c:dLbls><c:showVal/><c:showPercent val="false"/><c:showCatName/><c:showSerName val="0"/><c:showLeaderLines/><c:showLegendKey/><c:showBubbleSize val="0"/><c:leaderLines><c:spPr><a:ln w="6350" cap="sq"><a:solidFill><a:srgbClr val="778899"><a:alpha val="40000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr></c:leaderLines><c:dLblPos val="outEnd"/><c:separator>; </c:separator><c:numFmt formatCode="#,##0.0"/><c:spPr><a:solidFill><a:srgbClr val="FFEACC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr rot="2700000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1" i="0"><a:solidFill><a:srgbClr val="0A0B0C"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr><c:dLbl><c:idx val="1"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="850" b="1"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:rPr><a:t>ZX</a:t></a:r><a:r><a:rPr i="1"><a:latin typeface="Arial"/></a:rPr><a:t>Q</a:t></a:r></a:p></c:rich></c:tx><c:showVal val="0"/><c:showSerName/><c:showLegendKey val="0"/><c:showBubbleSize/><c:leaderLines><c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="AA3366"><a:alpha val="75000"/></a:srgbClr></a:solidFill></a:ln></c:spPr></c:leaderLines><c:dLblPos val="ctr"/><c:layout><c:manualLayout><c:x val="0.44"/><c:y val="0.28"/><c:w val="0.1"/><c:h val="0.08"/></c:manualLayout></c:layout><c:separator> / </c:separator><c:numFmt formatCode="0%"/><c:spPr><a:solidFill><a:srgbClr val="CCEEFF"/></a:solidFill></c:spPr><c:txPr><a:bodyPr rot="-1800000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="1"><a:solidFill><a:srgbClr val="334455"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr></c:dLbl></c:dLbls>
                    <c:ser>
                      <c:idx val="7"/>
                      <c:order val="3"/>
                      <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
                      <c:dLbls><c:showVal val="1"/><c:showCatName val="0"/><c:dLblPos val="t"/><c:separator> + </c:separator></c:dLbls>
                      <c:spPr><a:solidFill><a:srgbClr val="AA5500"><a:alpha val="70000"/></a:srgbClr></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="003366"/></a:solidFill></a:ln><a:effectLst><a:reflection/></a:effectLst></c:spPr>
                      <c:marker><c:symbol val="diamond"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="00AA55"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="5500AA"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:ln></c:spPr></c:marker>
                      <c:explosion val="12"/>
                      <c:smooth val="1"/>
                      <c:dPt><c:idx val="1"/><c:explosion val="18"/><c:spPr><a:solidFill><a:srgbClr val="CC8844"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="224466"/></a:solidFill></a:ln><a:effectDag/></c:spPr></c:dPt>
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
                  <c:catAx><c:axId val="10"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Categories</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/><c:txPr><a:bodyPr rot="1200000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"><a:solidFill><a:srgbClr val="224488"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:title><c:axPos val="b"/><c:crossAx val="20"/><c:crosses val="autoZero"/><c:majorTickMark val="out"/><c:minorTickMark val="in"/><c:lblOffset val="100"/><c:tickLblSkip val="2"/><c:tickMarkSkip val="3"/><c:tickLblPos val="low"/><c:noMultiLvlLbl val="1"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:catAx>
                  <c:valAx><c:axId val="20"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Value</a:t></a:r></a:p></c:rich></c:tx><c:layout><c:manualLayout><c:x val="0.02"/><c:y val="0.2"/><c:w val="0.1"/><c:h val="0.5"/></c:manualLayout></c:layout><c:spPr><a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill></c:spPr></c:title><c:axPos val="l"/><c:delete val="0"/><c:scaling><c:orientation val="maxMin"/><c:min val="0"/><c:max val="20"/></c:scaling><c:crossAx val="10"/><c:crosses val="max"/><c:crossesAt val="2.5"/><c:crossBetween val="between"/><c:majorUnit val="5"/><c:minorUnit val="1"/><c:majorGridlines><c:spPr><a:ln w="6350" cap="rnd" cmpd="thickThin"><a:solidFill><a:srgbClr val="8899AA"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/><a:bevel/></a:ln></c:spPr></c:majorGridlines><c:minorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:minorGridlines><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="336699"><a:alpha val="75000"/></a:srgbClr></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1" i="1"><a:solidFill><a:srgbClr val="654321"/></a:solidFill><a:latin typeface="Aptos"/></a:defRPr></a:pPr></a:p></c:txPr><c:tickLblPos val="high"/><c:numFmt formatCode="$#,##0"/></c:valAx>
                  </c:plotArea>
                  <c:legend><c:legendPos val="b"/><c:overlay val="1"/><c:layout><c:manualLayout><c:x val="0.7"/><c:y val="0.68"/><c:w val="0.2"/><c:h val="0.15"/></c:manualLayout></c:layout><c:spPr><a:solidFill><a:srgbClr val="EEF7FF"/></a:solidFill><a:ln w="6350"><a:solidFill><a:srgbClr val="334455"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr rot="-5400000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="700" b="0" i="1"><a:solidFill><a:srgbClr val="2211AA"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:pPr></a:p></c:txPr></c:legend>
                  </c:chart>
                  <c:externalData r:id="rId3"><c:autoUpdate val="0"/></c:externalData>
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
            ["ppt/charts/style1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="10">
                  <cs:chartStyle/>
                  <cs:gridlineMajor>
                    <cs:lnRef idx="1"/>
                    <cs:fillRef idx="2"/>
                    <cs:effectRef idx="1"/>
                    <cs:spPr><a:solidFill><a:srgbClr val="445566"><a:alpha val="75000"/></a:srgbClr></a:solidFill><a:ln w="12700" cap="rnd" cmpd="thickThin"><a:solidFill><a:srgbClr val="102030"><a:alpha val="80000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/><a:bevel/></a:ln></cs:spPr>
                  </cs:gridlineMajor>
                  <cs:gridlineMinor>
                    <cs:lnRef idx="0"/>
                    <cs:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="203040"><a:alpha val="70000"/></a:srgbClr></a:solidFill></a:ln></cs:spPr>
                  </cs:gridlineMinor>
                  <cs:title>
                    <cs:fontRef idx="major"><a:schemeClr val="accent6"><a:alpha val="50000"/></a:schemeClr></cs:fontRef>
                    <cs:defRPr sz="1400" b="1" i="0"/>
                  </cs:title>
                </cs:style>
                """,
            ["ppt/embeddings/chart-data.xlsx"] = "fake workbook package bytes for scene resource ownership",
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
                    <p:sp><p:nvSpPr><p:cNvPr id="1" name="MasterBox"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle><a:lvl2pPr/></a:lstStyle><a:p/></p:txBody></p:sp>
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
            ["ppt/media/image1.png"] = "scene image bytes",
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld>
                    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="123456"><a:alpha val="80000"/></a:srgbClr></a:solidFill></p:bgPr></p:bg>
                    <p:spTree>
                    <p:sp><p:nvSpPr><p:cNvPr id="4" name="TextBox"><a:hlinkClick r:id="rIdShapeHyperlink" action="ppaction://hlinkshowjump"/></p:cNvPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:solidFill><a:srgbClr val="CCDD11"><a:alpha val="75000"/></a:srgbClr></a:solidFill><a:effectLst><a:glow rad="91440"><a:srgbClr val="0000FF"><a:alpha val="25000"/></a:srgbClr></a:glow><a:outerShdw dist="91440" dir="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw><a:reflection blurRad="6350"/></a:effectLst><a:effectDag/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="1"/><a:r><a:rPr u="sng"><a:highlight><a:srgbClr val="FFFF00"/></a:highlight></a:rPr><a:t>Hello</a:t></a:r><a:br/><a:fld type="slidenum"><a:rPr sz="1200"/><a:t>1</a:t></a:fld><a:endParaRPr sz="1800"/></a:p></p:txBody></p:sp>
                    <p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"><a:alphaModFix amt="50000"/><a:lum bright="25000" contrast="-15000"/></a:blip><a:srcRect l="10000" t="20000" r="30000" b="40000"/><a:stretch><a:fillRect l="5000" r="10000"/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr></p:pic>
                    <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Table"/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="1828800"/><a:ext cx="1828800" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{93296810-A885-4BE3-A3E7-6D5BEEA58F35}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="914400"/><a:gridCol w="914400"/></a:tblGrid><a:tr h="914400"><a:tc gridSpan="2"><a:txBody><a:bodyPr lIns="91440" tIns="45720"/><a:lstStyle/><a:p/></a:txBody><a:tcPr marL="182880" anchor="ctr"><a:solidFill><a:srgbClr val="445566"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:lnL w="50800" cap="rnd" cmpd="dbl"><a:solidFill><a:srgbClr val="778899"><a:alpha val="60000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/><a:bevel/></a:lnL></a:tcPr></a:tc><a:tc hMerge="1"/></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
                    <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="8" name="Connector"/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="0" y="2743200"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst></a:prstGeom><a:ln w="25400" cap="rnd" cmpd="dbl"><a:solidFill><a:srgbClr val="336699"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:prstDash val="dash"/><a:bevel/><a:headEnd type="arrow" w="lg"/><a:tailEnd type="triangle" len="sm"/></a:ln></p:spPr></p:cxnSp>
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneSnapshot sceneSnapshot = PptxRenderer.InspectScene(document, package);

        TestAssert.Equal(1, scene.Slides.Count);
        TestAssert.Equal(1, sceneSnapshot.Slides.Count);
        PptxSceneSlide slide = scene.Slides[0];
        PptxSceneSlideSnapshot slideSnapshot = sceneSnapshot.Slides[0];
        TestAssert.True(slide.MasterXml is not null, "Expected master XML ownership in the scene model.");
        TestAssert.True(slide.LayoutXml is not null, "Expected layout XML ownership in the scene model.");
        TestAssert.True(slide.SlideXml.Root is not null, "Expected slide XML ownership in the scene model.");
        TestAssert.True(slideSnapshot.HasMasterXml, "Expected scene inspection to expose master XML ownership without XML content.");
        TestAssert.True(slideSnapshot.HasLayoutXml, "Expected scene inspection to expose layout XML ownership without XML content.");
        TestAssert.True(slideSnapshot.HasSlideXml, "Expected scene inspection to expose slide XML ownership without XML content.");
        TestAssert.Equal(1, slide.MasterRelationships.Count);
        TestAssert.Equal(1, slide.LayoutRelationships.Count);
        TestAssert.Equal(4, slide.SlideRelationships.Count);
        TestAssert.Equal(1, slideSnapshot.MasterRelationshipCount);
        TestAssert.Equal(1, slideSnapshot.LayoutRelationshipCount);
        TestAssert.Equal(4, slideSnapshot.SlideRelationshipCount);
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
        TestAssert.True(slide.SlideNodes[0].HyperlinkClick.IsDefined, "Expected shape-level hyperlink-click source state in the scene model.");
        TestAssert.Equal("rIdShapeHyperlink", slide.SlideNodes[0].HyperlinkClick.RelationshipId ?? string.Empty);
        TestAssert.Equal("ppaction://hlinkshowjump", slide.SlideNodes[0].HyperlinkClick.Action ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[0].HasHyperlinkClick, "Expected scene inspection to expose shape-level hyperlink-click source state.");
        TestAssert.Equal("rIdShapeHyperlink", slideSnapshot.SlideNodes[0].HyperlinkClickId ?? string.Empty);
        TestAssert.Equal("ppaction://hlinkshowjump", slideSnapshot.SlideNodes[0].HyperlinkClickAction ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[0].HasTextBody, "Expected private-safe scene inspection to expose text-body ownership without text content.");
        TestAssert.Equal(1, slideSnapshot.SlideNodes[0].TextParagraphCount);
        TestAssert.Equal(3, slideSnapshot.SlideNodes[0].TextRunCount);
        TestAssert.Equal("Table", slideSnapshot.SlideNodes[2].Kind);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[2].TableRowCount);
        TestAssert.Equal(2, slideSnapshot.SlideNodes[2].TableCellCount);
        TestAssert.Equal("{93296810-A885-4BE3-A3E7-6D5BEEA58F35}", slideSnapshot.SlideNodes[2].TableStyleId);
        TestAssert.True(slideSnapshot.SlideNodes[2].TableStyleIsSupported, "Expected scene inspection to expose table-style support state.");
        TestAssert.Equal("Medium-Style-2", slideSnapshot.SlideNodes[2].TableStyleName);
        TestAssert.Equal("MediumStyle2", slideSnapshot.SlideNodes[2].TableStyleKind);
        TestAssert.Equal("accent6", slideSnapshot.SlideNodes[2].TableStyleAccent);
        TestAssert.True(slideSnapshot.SlideNodes[2].TableStyleFirstRow, "Expected scene inspection to expose table first-row state.");
        TestAssert.Equal("1", slideSnapshot.SlideNodes[2].TableStyleFirstRowValue);
        TestAssert.True(slideSnapshot.SlideNodes[2].TableStyleBandRow, "Expected scene inspection to expose table band-row state.");
        TestAssert.Equal("1", slideSnapshot.SlideNodes[2].TableStyleBandRowValue);
        TestAssert.Equal(2, slideSnapshot.SlideNodes[2].TableStyleFillCellCount);
        TestAssert.Equal(2, slideSnapshot.SlideNodes[2].TableStyleTextColorCellCount);
        TestAssert.Equal(2, slideSnapshot.SlideNodes[2].TableStyleTextBoldCellCount);
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
        TestAssert.True(slide.SlideNodes[0].Shape?.Effects.HasEffectList == true, "Expected shape effect list provenance in the scene model.");
        TestAssert.True(slide.SlideNodes[0].Shape?.Effects.HasEffectDag == true, "Expected shape effectDag provenance in the scene model.");
        TestAssert.Equal("reflection", string.Join(",", slide.SlideNodes[0].Shape?.Effects.UnsupportedEffectNames ?? []));
        TestAssert.True(slideSnapshot.SlideNodes[0].ShapeHasEffectList, "Expected scene inspection to expose shape effect list provenance.");
        TestAssert.True(slideSnapshot.SlideNodes[0].ShapeHasEffectDag, "Expected scene inspection to expose shape effectDag provenance.");
        TestAssert.Equal(1, slideSnapshot.SlideNodes[0].ShapeUnsupportedEffectCount);
        TestAssert.Equal("reflection", string.Join(",", slideSnapshot.SlideNodes[0].ShapeUnsupportedEffectNames));
        TestAssert.Equal(PptxSceneNodeKind.Picture, slide.SlideNodes[1].Kind);
        TestAssert.Equal("rIdImage", slide.SlideNodes[1].Picture?.RelationshipId ?? string.Empty);
        TestAssert.Equal("/ppt/media/image1.png", slide.SlideNodes[1].Picture?.TargetPartName ?? string.Empty);
        TestAssert.Equal("image/png", slide.SlideNodes[1].Picture?.Resource?.ContentType ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[1].HasPictureResource, "Expected scene inspection to expose picture resource ownership without media bytes.");
        TestAssert.Equal("image/png", slideSnapshot.SlideNodes[1].PictureContentType);
        TestAssert.Equal(0.1d, slide.SlideNodes[1].Picture?.Crop.Left ?? 0d);
        TestAssert.Equal(0.4d, slide.SlideNodes[1].Picture?.Crop.Bottom ?? 0d);
        TestAssert.Equal("10000", slide.SlideNodes[1].Picture?.Crop.LeftValue ?? string.Empty);
        TestAssert.Equal("40000", slide.SlideNodes[1].Picture?.Crop.BottomValue ?? string.Empty);
        TestAssert.Equal(0.05d, slide.SlideNodes[1].Picture?.Fill.Left ?? 0d);
        TestAssert.Equal("5000", slide.SlideNodes[1].Picture?.Fill.LeftValue ?? string.Empty);
        TestAssert.Equal("10000", slide.SlideNodes[1].Picture?.Fill.RightValue ?? string.Empty);
        TestAssert.Equal(0.5d, slide.SlideNodes[1].Picture?.Alpha ?? 0d);
        TestAssert.Equal(PptxSceneImageRecolorKind.Luminance, slide.SlideNodes[1].Picture?.Recolor.Kind ?? PptxSceneImageRecolorKind.None);
        TestAssert.Equal("lum", slide.SlideNodes[1].Picture?.Recolor.KindValue ?? string.Empty);
        TestAssert.Equal("lum", slideSnapshot.SlideNodes[1].PictureRecolorKindValue);
        TestAssert.Equal(0.25d, slideSnapshot.SlideNodes[1].PictureRecolorBrightness ?? 0d);
        TestAssert.Equal(-0.15d, slideSnapshot.SlideNodes[1].PictureRecolorContrast ?? 0d);
        TestAssert.Equal("25000", slideSnapshot.SlideNodes[1].PictureRecolorBrightnessValue);
        TestAssert.Equal("-15000", slideSnapshot.SlideNodes[1].PictureRecolorContrastValue);
        TestAssert.Equal(PptxSceneNodeKind.Table, slide.SlideNodes[2].Kind);
        TestAssert.Equal(2, slide.SlideNodes[2].Table?.ColumnWidths.Count ?? 0);
        TestAssert.Equal(914400d, slide.SlideNodes[2].Table?.ColumnWidths[0] ?? 0d);
        TestAssert.Equal(1, slide.SlideNodes[2].Table?.RowHeights.Count ?? 0);
        TestAssert.Equal(914400d, slide.SlideNodes[2].Table?.RowHeights[0] ?? 0d);
        TestAssert.True(slide.SlideNodes[2].Table?.Style.IsSupported == true, "Expected built-in table style lookup in the scene model.");
        TestAssert.Equal("Medium-Style-2", slide.SlideNodes[2].Table?.Style.Name ?? string.Empty);
        TestAssert.Equal(PptxBuiltInTableStyleKind.MediumStyle2, slide.SlideNodes[2].Table?.Style.Kind ?? PptxBuiltInTableStyleKind.Unknown);
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
        TestAssert.Equal(PptxSceneTableCellTextInsetSource.CellProperties, slide.SlideNodes[2].Table?.Rows[0].Cells[0].TextInsetSources.Left ?? PptxSceneTableCellTextInsetSource.Default);
        TestAssert.Equal("182880", slide.SlideNodes[2].Table?.Rows[0].Cells[0].TextInsetValues.Left ?? string.Empty);
        TestAssert.Equal(PptxSceneTableCellTextInsetSource.BodyProperties, slide.SlideNodes[2].Table?.Rows[0].Cells[0].TextInsetSources.Top ?? PptxSceneTableCellTextInsetSource.Default);
        TestAssert.Equal("45720", slide.SlideNodes[2].Table?.Rows[0].Cells[0].TextInsetValues.Top ?? string.Empty);
        TestAssert.Equal(PptxSceneTableCellTextInsetSource.Default, slide.SlideNodes[2].Table?.Rows[0].Cells[0].TextInsetSources.Right ?? PptxSceneTableCellTextInsetSource.CellProperties);
        TestAssert.Equal(PptxSceneTableCellVerticalAnchor.Middle, slide.SlideNodes[2].Table?.Rows[0].Cells[0].VerticalAnchor ?? PptxSceneTableCellVerticalAnchor.Top);
        TestAssert.Equal("ctr", slide.SlideNodes[2].Table?.Rows[0].Cells[0].VerticalAnchorValue ?? string.Empty);
        TestAssert.Equal(PptxSceneTableCellVerticalAnchorSource.CellProperties, slide.SlideNodes[2].Table?.Rows[0].Cells[0].VerticalAnchorSource ?? PptxSceneTableCellVerticalAnchorSource.Default);
        TestAssert.Equal(PptxSceneTableCellVerticalAnchorSource.Default, slide.SlideNodes[2].Table?.Rows[0].Cells[1].VerticalAnchorSource ?? PptxSceneTableCellVerticalAnchorSource.CellProperties);
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[0].Fill.HasFill == true, "Expected direct table-cell fill in the scene model.");
        TestAssert.Equal(new RgbColor(68, 85, 102), slide.SlideNodes[2].Table?.Rows[0].Cells[0].Fill.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[2].Table?.Rows[0].Cells[0].Fill.Alpha ?? 0d);
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.IsSpecified == true, "Expected explicit table-cell border in the scene model.");
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.HasLine == true, "Expected resolved table-cell border line in the scene model.");
        TestAssert.Equal(new RgbColor(119, 136, 153), slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.Color ?? default);
        TestAssert.Equal(2d, slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.Width ?? 0d);
        TestAssert.Equal(0.6d, slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.Alpha ?? 0d);
        TestAssert.Equal("dash", slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.DashPreset ?? string.Empty);
        TestAssert.Equal("dbl", slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.CompoundValue ?? string.Empty);
        TestAssert.Equal("rnd", slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.CapValue ?? string.Empty);
        TestAssert.Equal("bevel", slide.SlideNodes[2].Table?.Rows[0].Cells[0].Borders.Left.Line.JoinValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[2].Table?.Rows[0].Cells[1].IsMergedContinuation == true, "Expected merged-cell continuation in the scene model.");
        TestAssert.Equal(PptxSceneNodeKind.Connector, slide.SlideNodes[3].Kind);
        TestAssert.Equal("straightConnector1", slide.SlideNodes[3].Shape?.Preset ?? string.Empty);
        TestAssert.Equal(50000d, slide.SlideNodes[3].Shape?.PresetAdjustments["adj1"] ?? 0d);
        TestAssert.True(slide.SlideNodes[3].Shape?.Line.HasLine == true, "Expected connector line style in the scene model.");
        TestAssert.Equal(new RgbColor(51, 102, 153), slide.SlideNodes[3].Shape?.Line.Color ?? default);
        TestAssert.Equal(2d, slide.SlideNodes[3].Shape?.Line.Width ?? 0d);
        TestAssert.Equal(0.5d, slide.SlideNodes[3].Shape?.Line.Alpha ?? 0d);
        TestAssert.Equal(8d, slide.SlideNodes[3].Shape?.Line.DashPattern[0] ?? 0d);
        TestAssert.Equal("dash", slide.SlideNodes[3].Shape?.Line.DashPreset ?? string.Empty);
        TestAssert.Equal(PptxSceneLineCompound.Double, slide.SlideNodes[3].Shape?.Line.Compound);
        TestAssert.Equal("dbl", slide.SlideNodes[3].Shape?.Line.CompoundValue ?? string.Empty);
        TestAssert.Equal(1, slide.SlideNodes[3].Shape?.Line.Cap ?? 0);
        TestAssert.Equal("rnd", slide.SlideNodes[3].Shape?.Line.CapValue ?? string.Empty);
        TestAssert.Equal(2, slide.SlideNodes[3].Shape?.Line.Join ?? 0);
        TestAssert.Equal("bevel", slide.SlideNodes[3].Shape?.Line.JoinValue ?? string.Empty);
        TestAssert.Equal(PptxSceneLineEndKind.Arrow, slide.SlideNodes[3].Shape?.HeadEnd.Kind ?? PptxSceneLineEndKind.None);
        TestAssert.Equal("arrow", slide.SlideNodes[3].Shape?.HeadEnd.TypeValue ?? string.Empty);
        TestAssert.Equal(1.5d, slide.SlideNodes[3].Shape?.HeadEnd.WidthScale ?? 0d);
        TestAssert.Equal("lg", slide.SlideNodes[3].Shape?.HeadEnd.WidthValue ?? string.Empty);
        TestAssert.Equal(PptxSceneLineEndKind.Triangle, slide.SlideNodes[3].Shape?.TailEnd.Kind ?? PptxSceneLineEndKind.None);
        TestAssert.Equal(0.5d, slide.SlideNodes[3].Shape?.TailEnd.LengthScale ?? 0d);
        TestAssert.Equal("triangle", slide.SlideNodes[3].Shape?.TailEnd.TypeValue ?? string.Empty);
        TestAssert.Equal("sm", slide.SlideNodes[3].Shape?.TailEnd.LengthValue ?? string.Empty);
        TestAssert.Equal(PptxSceneNodeKind.Chart, slide.SlideNodes[4].Kind);
        TestAssert.Equal("rIdChart", slide.SlideNodes[4].Chart?.RelationshipId ?? string.Empty);
        TestAssert.Equal("/ppt/charts/chart1.xml", slide.SlideNodes[4].Chart?.TargetPartName ?? string.Empty);
        TestAssert.Equal("rIdChart", slideSnapshot.SlideNodes[4].ChartRelationshipId);
        TestAssert.Equal("/ppt/charts/chart1.xml", slideSnapshot.SlideNodes[4].ChartTargetPartName);
        TestAssert.True(slide.SlideNodes[4].Chart?.ChartXml is not null, "Expected chart part XML ownership in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.ExternalData.IsDefined == true, "Expected chart external-data ownership in the scene model.");
        TestAssert.Equal("rId3", slide.SlideNodes[4].Chart?.ExternalData.RelationshipId ?? string.Empty);
        TestAssert.Equal("/ppt/embeddings/chart-data.xlsx", slide.SlideNodes[4].Chart?.ExternalData.TargetPartName ?? string.Empty);
        TestAssert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", slide.SlideNodes[4].Chart?.ExternalData.Resource?.ContentType ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[4].HasChartExternalData, "Expected scene inspection to expose chart external-data ownership.");
        TestAssert.Equal("rId3", slideSnapshot.SlideNodes[4].ChartExternalDataRelationshipId);
        TestAssert.Equal("/ppt/embeddings/chart-data.xlsx", slideSnapshot.SlideNodes[4].ChartExternalDataTargetPartName);
        TestAssert.True(slideSnapshot.SlideNodes[4].ChartExternalDataAutoUpdate == false, "Expected scene inspection to expose chart external-data auto-update policy.");
        TestAssert.Equal("0", slideSnapshot.SlideNodes[4].ChartExternalDataAutoUpdateValue);
        TestAssert.True(slideSnapshot.SlideNodes[4].HasChartExternalDataResource, "Expected scene inspection to expose embedded workbook resource ownership without package bytes.");
        TestAssert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", slideSnapshot.SlideNodes[4].ChartExternalDataContentType);
        TestAssert.True(slide.SlideNodes[4].Chart?.ExternalData.AutoUpdate == false, "Expected chart external-data auto-update flag in the scene model.");
        TestAssert.Equal("0", slide.SlideNodes[4].Chart?.ExternalData.AutoUpdateValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Options.Date1904 == true, "Expected chart date-system flag ownership in the scene model.");
        TestAssert.Equal("true", slide.SlideNodes[4].Chart?.Options.Date1904Value ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[4].ChartDate1904 == true, "Expected scene inspection to expose chart date-system policy.");
        TestAssert.Equal("true", slideSnapshot.SlideNodes[4].ChartDate1904Value);
        TestAssert.True(slide.SlideNodes[4].Chart?.Options.RoundedCorners == false, "Expected chart rounded-corners flag ownership in the scene model.");
        TestAssert.Equal("0", slide.SlideNodes[4].Chart?.Options.RoundedCornersValue ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[4].ChartRoundedCorners == false, "Expected scene inspection to expose chart rounded-corners policy.");
        TestAssert.Equal("0", slideSnapshot.SlideNodes[4].ChartRoundedCornersValue);
        TestAssert.True(slide.SlideNodes[4].Chart?.Options.PlotVisibleOnly == false, "Expected plot-visible-only flag ownership in the scene model.");
        TestAssert.Equal("false", slide.SlideNodes[4].Chart?.Options.PlotVisibleOnlyValue ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[4].ChartPlotVisibleOnly == false, "Expected scene inspection to expose chart plot-visible-only policy.");
        TestAssert.Equal("false", slideSnapshot.SlideNodes[4].ChartPlotVisibleOnlyValue);
        TestAssert.True(slide.SlideNodes[4].Chart?.Options.ShowDataLabelsOverMaximum == true, "Expected show-data-labels-over-maximum flag ownership in the scene model.");
        TestAssert.Equal("1", slide.SlideNodes[4].Chart?.Options.ShowDataLabelsOverMaximumValue ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[4].ChartShowDataLabelsOverMaximum == true, "Expected scene inspection to expose chart over-maximum data-label policy.");
        TestAssert.Equal("1", slideSnapshot.SlideNodes[4].ChartShowDataLabelsOverMaximumValue);
        TestAssert.Equal(PptxSceneChartDisplayBlanksAs.Span, slide.SlideNodes[4].Chart?.Options.DisplayBlanksAsKind);
        TestAssert.Equal("span", slide.SlideNodes[4].Chart?.Options.DisplayBlanksAs ?? string.Empty);
        TestAssert.Equal("span", slideSnapshot.SlideNodes[4].ChartDisplayBlanksAs);
        TestAssert.Equal(new RgbColor(1, 2, 3), slide.SlideNodes[4].Chart?.PaletteColors?[0] ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.ColorStyle.IsDefined == true, "Expected chart color-style ownership in the scene model.");
        TestAssert.Equal("/ppt/charts/colors1.xml", slide.SlideNodes[4].Chart?.ColorStyle.PartName ?? string.Empty);
        TestAssert.Equal("cycle", slide.SlideNodes[4].Chart?.ColorStyle.Method ?? string.Empty);
        TestAssert.Equal("10", slide.SlideNodes[4].Chart?.ColorStyle.Id ?? string.Empty);
        TestAssert.Equal(new RgbColor(1, 2, 3), slide.SlideNodes[4].Chart?.ColorStyle.Colors[0] ?? default);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.ColorStyle.Declarations.Count ?? 0);
        TestAssert.Equal("srgbClr", slide.SlideNodes[4].Chart?.ColorStyle.Declarations[0].Kind ?? string.Empty);
        TestAssert.Equal("010203", slide.SlideNodes[4].Chart?.ColorStyle.Declarations[0].Value ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.ColorStyle.Declarations[0].IsResolved == true, "Expected typed chart color-style declaration resolution state.");
        TestAssert.True(slide.SlideNodes[4].Chart?.ColorStyle.ColorStyleXml is not null, "Expected chart color-style XML ownership in the scene model.");
        TestAssert.Equal("10", (string?)slide.SlideNodes[4].Chart?.ColorStyle.ColorStyleXml?.Root?.Attribute("id") ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[4].HasChartColorStyle, "Expected scene inspection to expose chart color-style ownership without XML contents.");
        TestAssert.Equal("/ppt/charts/colors1.xml", slideSnapshot.SlideNodes[4].ChartColorStylePartName);
        TestAssert.Equal("cycle", slideSnapshot.SlideNodes[4].ChartColorStyleMethod);
        TestAssert.Equal("10", slideSnapshot.SlideNodes[4].ChartColorStyleId);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartColorStyleColorCount);
        TestAssert.Equal(0, slideSnapshot.SlideNodes[4].ChartColorStyleVariationCount);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartColorStyleDeclarationCount);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartColorStyleRootDeclarationCount);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartColorStyleResolvedDeclarationCount);
        TestAssert.Equal("srgbClr", slideSnapshot.SlideNodes[4].ChartColorStyleDeclarationKinds[0]);
        TestAssert.True(slide.SlideNodes[4].Chart?.StylePart.IsDefined == true, "Expected chart style-part ownership in the scene model.");
        TestAssert.Equal("/ppt/charts/style1.xml", slide.SlideNodes[4].Chart?.StylePart.PartName ?? string.Empty);
        TestAssert.Equal("10", slide.SlideNodes[4].Chart?.StylePart.Id ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.StylePart.StyleXml is not null, "Expected chart style XML ownership in the scene model.");
        TestAssert.True(slideSnapshot.SlideNodes[4].HasChartStylePart, "Expected scene inspection to expose chart style-part ownership without XML contents.");
        TestAssert.Equal("/ppt/charts/style1.xml", slideSnapshot.SlideNodes[4].ChartStylePartName);
        TestAssert.Equal("10", slideSnapshot.SlideNodes[4].ChartStylePartId);
        TestAssert.Equal(3, slideSnapshot.SlideNodes[4].ChartStyleEntryCount);
        TestAssert.Equal("gridlineMajor", slideSnapshot.SlideNodes[4].ChartStyleEntryRoles[0]);
        TestAssert.Equal("gridlineMinor", slideSnapshot.SlideNodes[4].ChartStyleEntryRoles[1]);
        TestAssert.Equal("title", slideSnapshot.SlideNodes[4].ChartStyleEntryRoles[2]);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartStyleEntrySourceIndexes[0]);
        TestAssert.Equal(2, slideSnapshot.SlideNodes[4].ChartStyleEntrySourceIndexes[1]);
        TestAssert.Equal(3, slideSnapshot.SlideNodes[4].ChartStyleEntrySourceIndexes[2]);
        TestAssert.Equal("http://schemas.microsoft.com/office/drawing/2012/chartStyle", slideSnapshot.SlideNodes[4].ChartStyleEntryNamespaceUris[0]);
        TestAssert.Equal("http://schemas.microsoft.com/office/drawing/2012/chartStyle", slideSnapshot.SlideNodes[4].ChartStyleEntryNamespaceUris[1]);
        TestAssert.Equal("http://schemas.microsoft.com/office/drawing/2012/chartStyle", slideSnapshot.SlideNodes[4].ChartStyleEntryNamespaceUris[2]);
        TestAssert.Equal(2, slideSnapshot.SlideNodes[4].ChartStyleShapeStyleCount);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartStyleShapeFillCount);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartStyleFillReferenceCount);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartStyleResolvedFillReferenceCount);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartStyleEffectReferenceCount);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartStyleResolvedEffectReferenceCount);
        TestAssert.Equal(1, slideSnapshot.SlideNodes[4].ChartStyleFontReferenceCount);
        PptxSceneChartStyleEntry majorGridlineStyle = slide.SlideNodes[4].Chart?.StylePart.Entries.FirstOrDefault(entry => entry.Role == "gridlineMajor") ?? default;
        TestAssert.Equal("gridlineMajor", majorGridlineStyle.Role ?? string.Empty);
        TestAssert.Equal(1, majorGridlineStyle.SourceIndex);
        TestAssert.Equal("http://schemas.microsoft.com/office/drawing/2012/chartStyle", majorGridlineStyle.NamespaceUri);
        TestAssert.Equal(1, majorGridlineStyle.LineReferenceIndex ?? 0);
        TestAssert.Equal("1", majorGridlineStyle.LineReferenceIndexValue);
        TestAssert.Equal(2, majorGridlineStyle.FillReferenceIndex ?? 0);
        TestAssert.Equal("2", majorGridlineStyle.FillReferenceIndexValue);
        TestAssert.True(majorGridlineStyle.FillReferenceFill.HasFill, "Expected chart style fill-reference theme resolution in the scene model.");
        TestAssert.Equal(new RgbColor(187, 221, 238), majorGridlineStyle.FillReferenceFill.Color);
        TestAssert.Equal(0.65d, majorGridlineStyle.FillReferenceFill.Alpha);
        TestAssert.Equal(1, majorGridlineStyle.EffectReferenceIndex ?? 0);
        TestAssert.Equal("1", majorGridlineStyle.EffectReferenceIndexValue);
        TestAssert.True(majorGridlineStyle.EffectReferenceEffects.HasEffectList, "Expected chart style effect-reference theme resolution in the scene model.");
        TestAssert.Equal("blur", majorGridlineStyle.EffectReferenceEffects.UnsupportedEffectNames[0]);
        TestAssert.True(majorGridlineStyle.Line.HasLine, "Expected chart style-part line-reference ownership in the scene model.");
        TestAssert.Equal(new RgbColor(171, 193, 35), majorGridlineStyle.Line.Color);
        TestAssert.Equal(2d, majorGridlineStyle.Line.Width);
        TestAssert.Equal(0.6d, majorGridlineStyle.Line.Alpha);
        TestAssert.Equal("dot", majorGridlineStyle.Line.DashPreset ?? string.Empty);
        TestAssert.Equal(PptxSceneLineCompound.Double, majorGridlineStyle.Line.Compound);
        TestAssert.Equal("sq", majorGridlineStyle.Line.CapValue ?? string.Empty);
        TestAssert.Equal(1, majorGridlineStyle.Line.Join);
        TestAssert.True(majorGridlineStyle.ShapeLine.HasLine, "Expected chart style-part role shape-line ownership in the scene model.");
        TestAssert.Equal(new RgbColor(16, 32, 48), majorGridlineStyle.ShapeLine.Color);
        TestAssert.Equal(1d, majorGridlineStyle.ShapeLine.Width);
        TestAssert.Equal(0.8d, majorGridlineStyle.ShapeLine.Alpha);
        TestAssert.Equal("dash", majorGridlineStyle.ShapeLine.DashPreset ?? string.Empty);
        TestAssert.Equal(PptxSceneLineCompound.ThickThin, majorGridlineStyle.ShapeLine.Compound);
        TestAssert.Equal("rnd", majorGridlineStyle.ShapeLine.CapValue ?? string.Empty);
        TestAssert.Equal(2, majorGridlineStyle.ShapeLine.Join);
        TestAssert.True(majorGridlineStyle.ShapeStyle.Fill.HasFill, "Expected chart style role-local shape fill ownership in the scene model.");
        TestAssert.Equal(new RgbColor(68, 85, 102), majorGridlineStyle.ShapeStyle.Fill.Color);
        TestAssert.Equal(0.75d, majorGridlineStyle.ShapeStyle.Fill.Alpha);
        PptxSceneChartStyleEntry minorGridlineStyle = slide.SlideNodes[4].Chart?.StylePart.Entries.FirstOrDefault(entry => entry.Role == "gridlineMinor") ?? default;
        TestAssert.Equal(null, minorGridlineStyle.LineReferenceIndex);
        TestAssert.Equal("0", minorGridlineStyle.LineReferenceIndexValue);
        PptxSceneChartStyleEntry titleStyle = slide.SlideNodes[4].Chart?.StylePart.Entries.FirstOrDefault(entry => entry.Role == "title") ?? default;
        TestAssert.Equal("title", titleStyle.Role ?? string.Empty);
        TestAssert.Equal("major", titleStyle.FontReferenceIndex);
        TestAssert.Equal("Arial", titleStyle.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal("+mj-lt", titleStyle.TextStyle.RequestedTypeface ?? string.Empty);
        TestAssert.Equal(PptxThemeTypefaceSource.MajorLatin, titleStyle.TextStyle.TypefaceSource ?? default);
        TestAssert.Equal(14d, titleStyle.TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(51, 102, 153), titleStyle.TextStyle.Color ?? default);
        TestAssert.Equal(0.5d, titleStyle.TextStyle.Alpha ?? 0d);
        TestAssert.True(titleStyle.TextStyle.Bold == true, "Expected chart style-part title bold default in the scene model.");
        TestAssert.True(titleStyle.TextStyle.Italic == false, "Expected chart style-part title italic default in the scene model.");
        TestAssert.Equal("10", slide.SlideNodes[4].Chart?.StyleId ?? string.Empty);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.TextStyle.RequestedTypeface ?? string.Empty);
        TestAssert.Equal(PptxThemeTypefaceSource.Direct, slide.SlideNodes[4].Chart?.TextStyle.TypefaceSource ?? default);
        TestAssert.Equal(11d, slide.SlideNodes[4].Chart?.TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(16, 17, 18), slide.SlideNodes[4].Chart?.TextStyle.Color ?? default);
        TestAssert.Equal(0.8d, slide.SlideNodes[4].Chart?.TextStyle.Alpha ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.TextStyle.Bold == true, "Expected chart-level default run bold style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.TextStyle.Italic == true, "Expected chart-level default run italic style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.TextStyle.Underline == true, "Expected chart-level default run underline style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.TextStyle.Strike == true, "Expected chart-level default run strike style in the scene model.");
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
        TestAssert.True(slideSnapshot.SlideNodes[4].HasChartPlotAreaManualLayout, "Expected scene inspection to expose chart plot-area manual layout ownership.");
        TestAssert.Equal(0.12d, slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutX ?? 0d);
        TestAssert.Equal(0.18d, slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutY ?? 0d);
        TestAssert.Equal(0.72d, slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutWidth ?? 0d);
        TestAssert.Equal(0.66d, slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutHeight ?? 0d);
        TestAssert.Equal("inner", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutTarget);
        TestAssert.Equal("Inner", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutTargetKind);
        TestAssert.Equal("factor", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutXMode);
        TestAssert.Equal("Factor", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutXModeKind);
        TestAssert.Equal("factor", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutYMode);
        TestAssert.Equal("Factor", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutYModeKind);
        TestAssert.Equal("factor", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutWidthMode);
        TestAssert.Equal("Factor", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutWidthModeKind);
        TestAssert.Equal("factor", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutHeightMode);
        TestAssert.Equal("Factor", slideSnapshot.SlideNodes[4].ChartPlotAreaLayoutHeightModeKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.ChartAreaStyle.NoFill == false, "Expected chart area noFill to be absent in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.ChartAreaStyle.PatternFill.HasPattern == true, "Expected chart area pattern fill in the scene model.");
        TestAssert.Equal("pct25", slide.SlideNodes[4].Chart?.ChartAreaStyle.PatternFill.Preset ?? string.Empty);
        TestAssert.Equal(new RgbColor(34, 68, 102), slide.SlideNodes[4].Chart?.ChartAreaStyle.PatternFill.Foreground ?? default);
        TestAssert.Equal(new RgbColor(241, 226, 211), slide.SlideNodes[4].Chart?.ChartAreaStyle.PatternFill.Background ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.ChartAreaStyle.Line.HasLine == true, "Expected chart area line in the scene model.");
        TestAssert.Equal(new RgbColor(68, 85, 102), slide.SlideNodes[4].Chart?.ChartAreaStyle.Line.Color ?? default);
        TestAssert.Equal(1d, slide.SlideNodes[4].Chart?.ChartAreaStyle.Line.Width ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.ChartAreaStyle.OuterShadow.HasShadow == true, "Expected chart area outer shadow in the scene model.");
        TestAssert.Equal(new RgbColor(1, 2, 3), slide.SlideNodes[4].Chart?.ChartAreaStyle.OuterShadow.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[4].Chart?.ChartAreaStyle.OuterShadow.Alpha ?? 0d);
        TestAssert.Equal(1d, slide.SlideNodes[4].Chart?.ChartAreaStyle.OuterShadow.OffsetX ?? 0d);
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
        TestAssert.Equal(PptxSceneChartRadarStyle.Unknown, slide.SlideNodes[4].Chart?.Plots[0].RadarStyleKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].VaryColors == false, "Expected chart plot varyColors in the scene model.");
        TestAssert.Equal("false", slide.SlideNodes[4].Chart?.Plots[0].VaryColorsValue ?? string.Empty);
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].MarkersEnabledValue ?? string.Empty);
        TestAssert.Equal(175d, slide.SlideNodes[4].Chart?.Plots[0].GapWidth ?? 0d);
        TestAssert.Equal(25d, slide.SlideNodes[4].Chart?.Plots[0].Overlap ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowValue == true, "Expected chart data-label value flag in the scene model.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowValueValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowPercent == false, "Expected chart data-label percent flag in the scene model.");
        TestAssert.Equal("false", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowPercentValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowCategoryName == true, "Expected chart data-label category-name flag in the scene model.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowCategoryNameValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowSeriesName == false, "Expected chart data-label series-name flag in the scene model.");
        TestAssert.Equal("0", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowSeriesNameValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowLeaderLines == true, "Expected chart data-label leader-lines flag in the scene model.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowLeaderLinesValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowLegendKey == true, "Expected chart data-label legend-key flag in the scene model.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowLegendKeyValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowBubbleSize == false, "Expected chart data-label bubble-size flag in the scene model.");
        TestAssert.Equal("0", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShowBubbleSizeValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.LeaderLines.IsDefined == true, "Expected chart data-label leader-line source element in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.LeaderLines.Line.HasLine == true, "Expected chart data-label leader-line style in the scene model.");
        TestAssert.Equal(new RgbColor(119, 136, 153), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.LeaderLines.Line.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.LeaderLines.Line.Width ?? 0d);
        TestAssert.Equal(0.4d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.LeaderLines.Line.Alpha ?? 0d);
        TestAssert.Equal(2d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.LeaderLines.Line.DashPattern[0] ?? 0d);
        TestAssert.Equal("dash", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.LeaderLines.Line.DashPreset ?? string.Empty);
        TestAssert.Equal(2, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.LeaderLines.Line.Cap ?? 0);
        TestAssert.Equal("sq", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.LeaderLines.Line.CapValue ?? string.Empty);
        TestAssert.Equal("outEnd", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Position ?? string.Empty);
        TestAssert.Equal("; ", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Separator ?? string.Empty);
        TestAssert.Equal("#,##0.0", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.NumberFormat ?? string.Empty);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.RequestedTypeface ?? string.Empty);
        TestAssert.Equal(PptxThemeTypefaceSource.Direct, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.TypefaceSource ?? default);
        TestAssert.Equal(10d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(10, 11, 12), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.Bold == true, "Expected plot-level data-label bold style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextStyle.Italic == false, "Expected explicit plot-level data-label italic disable in the scene model.");
        TestAssert.Equal(45d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextBodyProperties.RotationDegrees ?? 0d);
        TestAssert.Equal("2700000", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.TextBodyProperties.RotationValue ?? string.Empty);
        TestAssert.Equal(new RgbColor(255, 234, 204), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShapeStyle.Fill.Color ?? default);
        TestAssert.Equal(new RgbColor(17, 34, 51), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShapeStyle.Line.Color ?? default);
        TestAssert.Equal(1d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.ShapeStyle.Line.Width ?? 0d);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides.Count ?? 0);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].Index ?? -1);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowValue == false, "Expected per-label show-value override in the scene model.");
        TestAssert.Equal("0", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowValueValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowSeriesName == true, "Expected per-label show-series-name override in the scene model.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowSeriesNameValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowLegendKey == false, "Expected per-label show-legend-key override in the scene model.");
        TestAssert.Equal("0", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowLegendKeyValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowBubbleSize == true, "Expected per-label show-bubble-size override in the scene model.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowBubbleSizeValue ?? string.Empty);
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowPercentValue ?? string.Empty);
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowCategoryNameValue ?? string.Empty);
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].ShowLeaderLinesValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].LeaderLines.IsDefined == true, "Expected per-label leader-line source element in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].LeaderLines.Line.HasLine == true, "Expected per-label leader-line style in the scene model.");
        TestAssert.Equal(new RgbColor(170, 51, 102), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].LeaderLines.Line.Color ?? default);
        TestAssert.Equal(2d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].LeaderLines.Line.Width ?? 0d);
        TestAssert.Equal(0.75d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].LeaderLines.Line.Alpha ?? 0d);
        TestAssert.Equal("ZXQ", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomText ?? string.Empty);
        TestAssert.Equal(2, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns.Count ?? 0);
        TestAssert.Equal("ZX", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns[0].Text ?? string.Empty);
        TestAssert.Equal(8.5d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns[0].TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(171, 205, 239), slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns[0].TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns[0].TextStyle.Bold == true, "Expected first custom label rich-text run bold style in the scene model.");
        TestAssert.Equal("Q", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns[1].Text ?? string.Empty);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns[1].TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns[1].TextStyle.RequestedTypeface ?? string.Empty);
        TestAssert.Equal(PptxThemeTypefaceSource.Direct, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns[1].TextStyle.TypefaceSource ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].CustomTextRuns[1].TextStyle.Italic == true, "Expected second custom label rich-text run italic style in the scene model.");
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
        TestAssert.Equal(-30d, slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].TextBodyProperties.RotationDegrees ?? 0d);
        TestAssert.Equal("-1800000", slide.SlideNodes[4].Chart?.Plots[0].DataLabels.Overrides[0].TextBodyProperties.RotationValue ?? string.Empty);
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
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].Effects.HasEffectList == true, "Expected chart series effect-list provenance in the scene model.");
        TestAssert.Equal("reflection", slide.SlideNodes[4].Chart?.Plots[0].Series[0].Effects.UnsupportedEffectNames[0] ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.IsDefined == true, "Expected explicit chart marker ownership in the scene model.");
        TestAssert.Equal("diamond", slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Symbol ?? string.Empty);
        TestAssert.Equal(PptxSceneChartMarkerSymbol.Diamond, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.SymbolKind);
        TestAssert.Equal("7", slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.SizeValue ?? string.Empty);
        TestAssert.Equal(7d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Size ?? 0d);
        TestAssert.Equal(new RgbColor(0, 170, 85), slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Fill.Color ?? default);
        TestAssert.Equal(new RgbColor(85, 0, 170), slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Line.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Marker.Line.Alpha ?? 0d);
        TestAssert.Equal(12d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].Explosion ?? 0d);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Plots[0].Series[0].PointStyles[0].Index ?? -1);
        TestAssert.Equal(new RgbColor(204, 136, 68), slide.SlideNodes[4].Chart?.Plots[0].Series[0].PointStyles[0].Fill.Color ?? default);
        TestAssert.Equal(new RgbColor(34, 68, 102), slide.SlideNodes[4].Chart?.Plots[0].Series[0].PointStyles[0].Line.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].PointStyles[0].Effects.HasEffectDag == true, "Expected chart point effectDag provenance in the scene model.");
        TestAssert.Equal(18d, slide.SlideNodes[4].Chart?.Plots[0].Series[0].PointStyles[0].Explosion ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[0].Series[0].Smooth == true, "Expected chart series smooth flag in the scene model.");
        TestAssert.Equal("1", slide.SlideNodes[4].Chart?.Plots[0].Series[0].SmoothValue ?? string.Empty);
        TestAssert.Equal("bubbleChart", slide.SlideNodes[4].Chart?.Plots[1].Kind ?? string.Empty);
        TestAssert.Equal(PptxSceneChartPlotKind.Bubble, slide.SlideNodes[4].Chart?.Plots[1].PlotKind);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Plots[1].PlotAreaIndex ?? -1);
        TestAssert.Equal(0, slide.SlideNodes[4].Chart?.Plots[1].KindIndex ?? -1);
        TestAssert.Equal(PptxSceneChartGrouping.Unknown, slide.SlideNodes[4].Chart?.Plots[1].GroupingKind);
        TestAssert.Equal(PptxSceneChartBarDirection.Unknown, slide.SlideNodes[4].Chart?.Plots[1].BarDirectionKind);
        TestAssert.Equal(PptxSceneChartRadarStyle.Unknown, slide.SlideNodes[4].Chart?.Plots[1].RadarStyleKind);
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
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[1].VaryColorsValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[1].Series[0].Marker.IsDefined == false, "Expected a missing marker element to remain distinct from the effective marker default.");
        TestAssert.Equal("lineChart", slide.SlideNodes[4].Chart?.Plots[2].Kind ?? string.Empty);
        TestAssert.Equal(PptxSceneChartPlotKind.Line, slide.SlideNodes[4].Chart?.Plots[2].PlotKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[2].MarkersEnabled is null, "Expected missing chart-level line marker metadata to remain observable.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Plots[2].MarkersEnabledValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[2].Series[0].Marker.IsDefined == false, "Expected missing line-chart marker metadata to remain distinct from explicit default markers.");
        TestAssert.Equal("none", slide.SlideNodes[4].Chart?.Plots[2].Series[0].Marker.Symbol ?? string.Empty);
        TestAssert.Equal(PptxSceneChartMarkerSymbol.None, slide.SlideNodes[4].Chart?.Plots[2].Series[0].Marker.SymbolKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Plots[2].Series[0].Smooth == false, "Expected explicit smooth disable to remain distinct from a missing smooth element.");
        TestAssert.Equal("0", slide.SlideNodes[4].Chart?.Plots[2].Series[0].SmoothValue ?? string.Empty);
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
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Axes[0].IsDeletedValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[0].NoMultiLevelLabels == true, "Expected chart axis multi-level-label flag in the scene model.");
        TestAssert.Equal("1", slide.SlideNodes[4].Chart?.Axes[0].NoMultiLevelLabelsValue ?? string.Empty);
        TestAssert.Equal("Categories", slide.SlideNodes[4].Chart?.Axes[0].Title.Text ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[0].Title.Overlay == false, "Expected chart category-axis title overlay flag in the scene model.");
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.Axes[0].Title.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal(new RgbColor(34, 68, 136), slide.SlideNodes[4].Chart?.Axes[0].Title.TextStyle.Color ?? default);
        TestAssert.Equal(20d, slide.SlideNodes[4].Chart?.Axes[0].Title.TextBodyProperties.RotationDegrees ?? 0d);
        TestAssert.Equal("1200000", slide.SlideNodes[4].Chart?.Axes[0].Title.TextBodyProperties.RotationValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].IsDeleted == false, "Expected chart axis delete flag in the scene model.");
        TestAssert.Equal("0", slide.SlideNodes[4].Chart?.Axes[1].IsDeletedValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].NoMultiLevelLabels is null, "Expected missing chart value-axis multi-level-label metadata to remain distinct from an explicit disable.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Axes[1].NoMultiLevelLabelsValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].HasScaling == true, "Expected chart axis scaling ownership in the scene model.");
        TestAssert.Equal("maxMin", slide.SlideNodes[4].Chart?.Axes[1].Orientation ?? string.Empty);
        TestAssert.Equal(PptxSceneChartAxisOrientation.MaximumMinimum, slide.SlideNodes[4].Chart?.Axes[1].OrientationKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].IsReversed == true, "Expected reversed chart axis orientation to remain available to existing scene consumers.");
        TestAssert.Equal(0d, slide.SlideNodes[4].Chart?.Axes[1].Minimum ?? -1d);
        TestAssert.Equal(20d, slide.SlideNodes[4].Chart?.Axes[1].Maximum ?? 0d);
        TestAssert.Equal(5d, slide.SlideNodes[4].Chart?.Axes[1].MajorUnit ?? 0d);
        TestAssert.Equal(1d, slide.SlideNodes[4].Chart?.Axes[1].MinorUnit ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].HasMajorGridlines == true, "Expected chart major gridline visibility in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].HasMinorGridlines == false, "Expected hidden chart minor gridline visibility in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].HasMajorGridlineElement == true, "Expected visible chart major gridline element presence in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].HasMinorGridlineElement == true, "Expected hidden chart minor gridline element presence to remain distinct from missing gridlines.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[0].HasMajorGridlineElement == false, "Expected missing category-axis major gridline element to remain distinct from hidden gridlines.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[0].HasMinorGridlineElement == false, "Expected missing category-axis minor gridline element to remain distinct from hidden gridlines.");
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
        TestAssert.Equal(PptxSceneLineCompound.ThickThin, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.Compound);
        TestAssert.Equal(1, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.Cap ?? 0);
        TestAssert.Equal(2, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineLine.Join ?? 0);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineLine.HasLine == true, "Expected hidden chart minor gridline line style in the scene model.");
        TestAssert.Equal(0d, slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineLine.Width ?? -1d);
        TestAssert.Equal(0d, slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineLine.Alpha ?? -1d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineStyleLine.HasLine == true, "Expected chart-style major gridline line candidate in the scene model.");
        TestAssert.Equal(new RgbColor(16, 32, 48), slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineStyleLine.Color ?? default);
        TestAssert.Equal(1d, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineStyleLine.Width ?? 0d);
        TestAssert.Equal(0.8d, slide.SlideNodes[4].Chart?.Axes[1].MajorGridlineStyleLine.Alpha ?? 0d);
        TestAssert.True(slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineStyleLine.HasLine == true, "Expected chart-style minor gridline line candidate in the scene model.");
        TestAssert.Equal(new RgbColor(32, 48, 64), slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineStyleLine.Color ?? default);
        TestAssert.Equal(0.5d, slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineStyleLine.Width ?? 0d);
        TestAssert.Equal(0.7d, slide.SlideNodes[4].Chart?.Axes[1].MinorGridlineStyleLine.Alpha ?? 0d);
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
        TestAssert.Equal(2, slide.SlideNodes[4].Chart?.Title.TextRuns.Count ?? 0);
        TestAssert.Equal("Scene ", slide.SlideNodes[4].Chart?.Title.TextRuns[0].Text ?? string.Empty);
        TestAssert.Equal(12.5d, slide.SlideNodes[4].Chart?.Title.TextRuns[0].TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(34, 51, 68), slide.SlideNodes[4].Chart?.Title.TextRuns[0].TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.TextRuns[0].TextStyle.Bold == true, "Expected first chart title rich-text run bold style in the scene model.");
        TestAssert.Equal("Chart", slide.SlideNodes[4].Chart?.Title.TextRuns[1].Text ?? string.Empty);
        TestAssert.Equal("Aptos", slide.SlideNodes[4].Chart?.Title.TextRuns[1].TextStyle.FontFamily ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.TextRuns[1].TextStyle.Italic == true, "Expected second chart title rich-text run italic style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.IsAutoDeleted is null, "Expected missing chart auto-title-delete metadata to remain distinct from an explicit disable.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Title.IsAutoDeletedValue ?? string.Empty);
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.Overlay == true, "Expected chart title overlay flag in the scene model.");
        TestAssert.Equal("1", slide.SlideNodes[4].Chart?.Title.OverlayValue ?? string.Empty);
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
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.ShapeStyle.Glow.HasGlow == true, "Expected chart title glow in the scene model.");
        TestAssert.Equal(new RgbColor(10, 11, 12), slide.SlideNodes[4].Chart?.Title.ShapeStyle.Glow.Color ?? default);
        TestAssert.Equal(0.4d, slide.SlideNodes[4].Chart?.Title.ShapeStyle.Glow.Alpha ?? 0d);
        TestAssert.Equal(2d, slide.SlideNodes[4].Chart?.Title.ShapeStyle.Glow.Radius ?? 0d);
        TestAssert.Equal("Arial", slide.SlideNodes[4].Chart?.Title.TextStyle.FontFamily ?? string.Empty);
        TestAssert.Equal(13d, slide.SlideNodes[4].Chart?.Title.TextStyle.FontSize ?? 0d);
        TestAssert.Equal(new RgbColor(17, 34, 170), slide.SlideNodes[4].Chart?.Title.TextStyle.Color ?? default);
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.TextStyle.Bold == true, "Expected chart title bold style in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.TextStyle.Italic == false, "Expected explicit chart title italic disable in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.TextStyle.Underline == false, "Expected explicit chart title underline disable in the scene model.");
        TestAssert.True(slide.SlideNodes[4].Chart?.Title.TextStyle.Strike == false, "Expected explicit chart title strike disable in the scene model.");
        TestAssert.Equal(90d, slide.SlideNodes[4].Chart?.Title.TextBodyProperties.RotationDegrees ?? 0d);
        TestAssert.Equal("5400000", slide.SlideNodes[4].Chart?.Title.TextBodyProperties.RotationValue ?? string.Empty);
        TestAssert.Equal("b", slide.SlideNodes[4].Chart?.Legend.Position ?? string.Empty);
        TestAssert.Equal(PptxSceneChartLegendPosition.Bottom, slide.SlideNodes[4].Chart?.Legend.PositionKind);
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.IsDefined == true, "Expected chart legend presence in the scene model.");
        TestAssert.True(slideSnapshot.SlideNodes[4].HasChartLegend, "Expected scene inspection to expose chart legend presence.");
        TestAssert.Equal("b", slideSnapshot.SlideNodes[4].ChartLegendPosition);
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.IsDeleted is null, "Expected missing chart legend delete metadata to remain distinct from an explicit disable.");
        TestAssert.True(slideSnapshot.SlideNodes[4].ChartLegendDeleted is null, "Expected scene inspection to preserve missing chart legend delete metadata.");
        TestAssert.Equal(string.Empty, slide.SlideNodes[4].Chart?.Legend.IsDeletedValue ?? string.Empty);
        TestAssert.Equal(string.Empty, slideSnapshot.SlideNodes[4].ChartLegendDeletedValue);
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.Overlay == true, "Expected chart legend overlay flag in the scene model.");
        TestAssert.Equal("1", slide.SlideNodes[4].Chart?.Legend.OverlayValue ?? string.Empty);
        TestAssert.Equal(-90d, slide.SlideNodes[4].Chart?.Legend.TextBodyProperties.RotationDegrees ?? 0d);
        TestAssert.Equal("-5400000", slide.SlideNodes[4].Chart?.Legend.TextBodyProperties.RotationValue ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[4].ChartLegendOverlay == true, "Expected scene inspection to expose chart legend overlay policy.");
        TestAssert.Equal("1", slideSnapshot.SlideNodes[4].ChartLegendOverlayValue);
        TestAssert.True(slide.SlideNodes[4].Chart?.Legend.Layout.HasLayout == true, "Expected chart legend manual layout in the scene model.");
        TestAssert.Equal(0.7d, slide.SlideNodes[4].Chart?.Legend.Layout.X ?? 0d);
        TestAssert.Equal(0.68d, slide.SlideNodes[4].Chart?.Legend.Layout.Y ?? 0d);
        TestAssert.Equal(0.2d, slide.SlideNodes[4].Chart?.Legend.Layout.Width ?? 0d);
        TestAssert.Equal(0.15d, slide.SlideNodes[4].Chart?.Legend.Layout.Height ?? 0d);
        TestAssert.True(slideSnapshot.SlideNodes[4].HasChartLegendManualLayout, "Expected scene inspection to expose chart legend manual layout ownership.");
        TestAssert.Equal(0.7d, slideSnapshot.SlideNodes[4].ChartLegendLayoutX ?? 0d);
        TestAssert.Equal(0.68d, slideSnapshot.SlideNodes[4].ChartLegendLayoutY ?? 0d);
        TestAssert.Equal(0.2d, slideSnapshot.SlideNodes[4].ChartLegendLayoutWidth ?? 0d);
        TestAssert.Equal(0.15d, slideSnapshot.SlideNodes[4].ChartLegendLayoutHeight ?? 0d);
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
        TestAssert.Equal("/ppt/media/image1.png", slide.SlideNodes[5].Children[0].Shape?.PictureFill.TargetPartName ?? string.Empty);
        TestAssert.Equal("image/png", slide.SlideNodes[5].Children[0].Shape?.PictureFill.Resource?.ContentType ?? string.Empty);
        TestAssert.True(slideSnapshot.SlideNodes[5].Children[0].HasShapePictureFillResource, "Expected scene inspection to expose shape picture-fill resource ownership without media bytes.");
        TestAssert.Equal("image/png", slideSnapshot.SlideNodes[5].Children[0].ShapePictureFillContentType);
        TestAssert.Equal(0.05d, slide.SlideNodes[5].Children[0].Shape?.PictureFill.Crop.Left ?? 0d);
        TestAssert.Equal(0.2d, slide.SlideNodes[5].Children[0].Shape?.PictureFill.Crop.Bottom ?? 0d);
        TestAssert.Equal("5000", slide.SlideNodes[5].Children[0].Shape?.PictureFill.Crop.LeftValue ?? string.Empty);
        TestAssert.Equal("20000", slide.SlideNodes[5].Children[0].Shape?.PictureFill.Crop.BottomValue ?? string.Empty);
        TestAssert.Equal(0.025d, slide.SlideNodes[5].Children[0].Shape?.PictureFill.Fill.Left ?? 0d);
        TestAssert.Equal(0.075d, slide.SlideNodes[5].Children[0].Shape?.PictureFill.Fill.Right ?? 0d);
        TestAssert.Equal("2500", slide.SlideNodes[5].Children[0].Shape?.PictureFill.Fill.LeftValue ?? string.Empty);
        TestAssert.Equal("7500", slide.SlideNodes[5].Children[0].Shape?.PictureFill.Fill.RightValue ?? string.Empty);
        TestAssert.True(slide.LayoutNodes[0].Shape?.CustomGeometry.HasGeometry == true, "Expected layout custom geometry in the scene model.");
        TestAssert.True(slide.LayoutNodes[0].Shape?.CustomGeometry.HasUnsupportedGeometry == false, "Expected supported custom geometry to remain non-diagnostic in the scene model.");
        TestAssert.Equal("xMid", slide.LayoutNodes[0].Shape?.CustomGeometry.Guides[0].Name ?? string.Empty);
        TestAssert.Equal(PptxSceneCustomCommandKind.LineTo, slide.LayoutNodes[0].Shape?.CustomGeometry.Paths[0].Commands[1].Kind);
        TestAssert.Equal("xMid", slide.LayoutNodes[0].Shape?.CustomGeometry.Paths[0].Commands[1].Points[0].X);
        TestAssert.True(slide.LayoutNodes[0].Shape?.CustomGeometry.Paths[0].AllowsStroke == false, "Expected custom path stroke=\"0\" to disable stroke in the scene model.");
        TestAssert.True(slide.LayoutNodes[1].IsPlaceholder, "Expected layout placeholder metadata in the scene model.");
        TestAssert.Equal(144d, slide.SlideNodes[0].Bounds?.Width ?? 0d);
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
        TestAssert.True(
            directTextRuns.Any(run => run.Text == "Hello"),
            "Expected direct text inspection to include Hello. Runs: " + string.Join("|", directTextRuns.Select(run => run.Text)));
        PptxTextRunSnapshot directHello = directTextRuns.First(run => run.Text == "Hello");
        TestAssert.Equal(26d, directHello.FontSize);
        TestAssert.True(directHello.Underline, "Expected direct renderer inspection to expose run underline.");
        TestAssert.Equal(new RgbColor(255, 255, 0), directHello.Highlight ?? default);

        PptxTextGlyphRunSnapshot[] glyphHelloRuns = PptxRenderer.InspectTextGlyphRuns(document, package, 0)
            .Where(run => run.FrameIndex == 0 && run.ParagraphIndex == 0 && run.SourceRunIndex == 0)
            .ToArray();
        TestAssert.Equal("Hello", string.Concat(glyphHelloRuns.Select(run => run.Text)));
        TestAssert.True(
            glyphHelloRuns.All(run => run.UnderlineWidth is not null && run.UnderlineHeight is not null),
            "Expected glyph-run inspection to expose underline geometry for each emitted source-run segment.");
        TestAssert.True(
            glyphHelloRuns.All(run => Math.Abs(run.UnderlineWidth.GetValueOrDefault() - run.Width) < 0.01d),
            "Expected underline geometry to be owned by each emitted glyph-run segment width.");

        IReadOnlyList<PptxTextFrameModelSnapshot> textFrames = PptxRenderer.InspectTextFrameModels(document, package, 0);
        TestAssert.True(
            textFrames.Any(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Text == "Hello"))),
            "Expected text-frame model inspection to include Hello. Runs: " + string.Join("|", textFrames.SelectMany(frame => frame.Paragraphs).SelectMany(paragraph => paragraph.Runs).Select(run => run.Text)));
        PptxTextFrameModelSnapshot textFrame = textFrames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Runs.Any(run => run.Text == "Hello")));
        TestAssert.True(textFrame.InheritedPlaceholderCount >= 1, "Expected text model to expose inherited placeholder participation.");
        TestAssert.True(textFrame.HasInheritedTextBody, "Expected text model to expose inherited placeholder text body participation.");
        TestAssert.True(!textFrame.UsesInheritedShapeBounds, "Expected direct text-box geometry to stay distinct from inherited placeholder style participation.");
        TestAssert.Equal(1, textFrame.Paragraphs.Count);
        TestAssert.Equal(1, textFrame.Paragraphs[0].Level);
        TestAssert.Equal("lvl2pPr", textFrame.Paragraphs[0].CascadeLevelName);
        TestAssert.True(textFrame.Paragraphs[0].ResolvedCascadeSourceCount >= 2, "Expected text model to expose inherited cascade inputs before style resolution.");
        TestAssert.Contains("shape.lstStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
        TestAssert.Contains("layout.placeholder.lstStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
        TestAssert.Contains("master.placeholder.lstStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
        TestAssert.Contains("inherited.txStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
        TestAssert.Contains("defaultTextStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerNames));
        TestAssert.Contains("ShapeListStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
        TestAssert.Contains("LayoutPlaceholderListStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
        TestAssert.Contains("MasterPlaceholderListStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
        TestAssert.Contains("InheritedTextStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
        TestAssert.Contains("DefaultTextStyle", string.Join("|", textFrame.Paragraphs[0].CascadeLayerKinds));
        TestAssert.True(textFrame.Paragraphs[0].ResolvedStyleSourceCount >= textFrame.Paragraphs[0].ResolvedCascadeSourceCount + 1, "Expected resolved paragraph style provenance to include direct paragraph properties after inherited defaults.");
        TestAssert.Contains("paragraph.pPr", string.Join("|", textFrame.Paragraphs[0].ResolvedStyleLayerNames));
        TestAssert.Contains("ParagraphProperties", string.Join("|", textFrame.Paragraphs[0].ResolvedStyleLayerKinds));
        TestAssert.Equal("Center", textFrame.Paragraphs[0].Alignment);
        TestAssert.Equal(26d, textFrame.Paragraphs[0].FontSize);
        TestAssert.Equal("Text", textFrame.Paragraphs[0].Runs[0].Kind);
        TestAssert.Equal("Break", textFrame.Paragraphs[0].Runs[1].Kind);
        TestAssert.Equal("Field", textFrame.Paragraphs[0].Runs[2].Kind);
        TestAssert.Equal("Hello", textFrame.Paragraphs[0].Runs[0].Text);
        TestAssert.True(textFrame.Paragraphs[0].Runs[0].ResolvedCascadeSourceCount >= 2, "Expected run model to expose direct and inherited run-style inputs.");
        TestAssert.Contains("run.rPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
        TestAssert.Contains("paragraph.defRPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
        TestAssert.Contains("layout.placeholder.lstStyle.defRPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
        TestAssert.Contains("inherited.txStyle.defRPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
        TestAssert.Contains("defaultTextStyle.defRPr", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerNames));
        TestAssert.Contains("RunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
        TestAssert.Contains("ParagraphDefaultRunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
        TestAssert.Contains("LayoutPlaceholderDefaultRunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
        TestAssert.Contains("InheritedTextStyleDefaultRunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
        TestAssert.Contains("DefaultTextStyleDefaultRunProperties", string.Join("|", textFrame.Paragraphs[0].Runs[0].CascadeLayerKinds));
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
        PptxTextFrameLayoutSnapshot layoutFrame = textLayout.Frames.Single(frame => frame.Paragraphs.Any(paragraph => paragraph.Lines.Any(line => line.Spans.Any(span => span.SourceText == "Hello"))));
        TestAssert.Equal(1, layoutFrame.Paragraphs.Count);
        TestAssert.Equal(2, layoutFrame.Paragraphs[0].Lines.Count);
        PptxTextSpanLayoutSnapshot[] layoutHelloSpans = layoutFrame.Paragraphs[0].Lines[0].Spans
            .Where(span => span.SourceText == "Hello")
            .ToArray();
        TestAssert.Equal("Hello", string.Concat(layoutHelloSpans.Select(span => span.Text)));
        TestAssert.True(layoutHelloSpans.All(span => span.SourceText == "Hello"), "Expected split layout spans to retain source-run text ownership.");
        TestAssert.Equal("1", layoutFrame.Paragraphs[0].Lines[1].Spans[0].Text);
        TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].EndX > layoutFrame.Paragraphs[0].Lines[0].StartX, "Expected layout line to own measured advance before PDF emission.");
        TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].TopY > layoutFrame.Paragraphs[0].Lines[0].BaselineY, "Expected layout line boxes to expose top-to-baseline geometry before PDF emission.");
        TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].Advance > 0d, "Expected layout line boxes to expose line advance before paragraph stacking.");
        TestAssert.True(layoutFrame.Paragraphs[0].Lines[0].BaselineOffset > 0d, "Expected layout line boxes to own baseline offset separately from text spans.");
        TestAssert.Equal("Default", layoutFrame.Paragraphs[0].Lines[0].LineSpacingKind);
    }

    public static void PptxSceneTableStylePreservesConditionalFormatFlagTokens()
    {
        XElement table = XElement.Parse("""
            <a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:tblPr firstRow="false" lastRow="1" firstCol="0" lastCol="true" bandRow="false" bandCol="1">
                <a:tableStyleId>{93296810-A885-4BE3-A3E7-6D5BEEA58F35}</a:tableStyleId>
              </a:tblPr>
            </a:tbl>
            """);

        PptxSceneTableStyle style = PptxSceneBuilder.ReadTableStyle(table);

        TestAssert.True(!style.FirstRow, "Expected parsed false first-row flag.");
        TestAssert.Equal("false", style.FirstRowValue);
        TestAssert.True(style.LastRow, "Expected numeric true last-row flag.");
        TestAssert.Equal("1", style.LastRowValue);
        TestAssert.True(!style.FirstColumn, "Expected numeric false first-column flag.");
        TestAssert.Equal("0", style.FirstColumnValue);
        TestAssert.True(style.LastColumn, "Expected textual true last-column flag.");
        TestAssert.Equal("true", style.LastColumnValue);
        TestAssert.True(!style.BandRow, "Expected parsed false band-row flag.");
        TestAssert.Equal("false", style.BandRowValue);
        TestAssert.True(style.BandColumn, "Expected numeric true band-column flag.");
        TestAssert.Equal("1", style.BandColumnValue);
    }

    public static void PptxPlaceholderMatchingPrefersExactThenIndex()
    {
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XElement slideShape = XElement.Parse("""
            <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:nvSpPr><p:cNvPr id="1" name="Slide"/><p:nvPr><p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr>
            </p:sp>
            """);
        XDocument exactSource = XDocument.Parse("""
            <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="IndexOnly"/><p:nvPr><p:ph idx="7"/></p:nvPr></p:nvSpPr></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="3" name="Exact"/><p:nvPr><p:ph type="body" idx="7"/></p:nvPr></p:nvSpPr></p:sp>
              </p:spTree></p:cSld>
            </p:sldLayout>
            """);
        XDocument indexSource = XDocument.Parse("""
            <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="4" name="TypeOnly"/><p:nvPr><p:ph type="body" idx="8"/></p:nvPr></p:nvSpPr></p:sp>
                <p:sp><p:nvSpPr><p:cNvPr id="5" name="IndexOnly"/><p:nvPr><p:ph idx="7"/></p:nvPr></p:nvSpPr></p:sp>
              </p:spTree></p:cSld>
            </p:sldLayout>
            """);
        System.Reflection.MethodInfo findInheritedPlaceholders = typeof(PptxRenderer).GetMethod(
            "FindInheritedPlaceholderShapes",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static) ?? throw new InvalidOperationException("Expected placeholder matcher.");

        object exactResult = findInheritedPlaceholders.Invoke(null, [slideShape, new[] { exactSource }]) ?? throw new InvalidOperationException("Expected exact placeholder result.");
        object indexResult = findInheritedPlaceholders.Invoke(null, [slideShape, new[] { indexSource }]) ?? throw new InvalidOperationException("Expected index placeholder result.");
        IReadOnlyList<XElement> exactMatches = (IReadOnlyList<XElement>)exactResult;
        IReadOnlyList<XElement> indexMatches = (IReadOnlyList<XElement>)indexResult;

        TestAssert.Equal("Exact", exactMatches[0].Descendants(p + "cNvPr").First().Attribute("name")?.Value ?? string.Empty);
        TestAssert.Equal("IndexOnly", indexMatches[0].Descendants(p + "cNvPr").First().Attribute("name")?.Value ?? string.Empty);
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
                      <a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:hlinkClick r:id="rIdHyper"/></a:rPr><a:t>Link</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """)
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxTextRunModelSnapshot linkRun = PptxRenderer.InspectTextFrameModels(document, package, 0)
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .Single(run => run.Text == "Link");
        TestAssert.True(linkRun.HasHyperlinkClick, "Expected the text model to preserve hyperlink-click source state before color resolution.");
        TestAssert.Equal("rIdHyper", linkRun.HyperlinkClickId ?? string.Empty);
        TestAssert.Equal("ThemeHyperlink", linkRun.ColorSource);
        TestAssert.True(linkRun.Underline, "Office applies a single underline to hyperlink runs when no underline override is authored.");
        TestAssert.Equal("sng", linkRun.UnderlineValue ?? string.Empty);
        PptxTextGlyphRunSnapshot glyphRun = PptxRenderer.InspectTextGlyphRuns(document, package, 0).Single(run => run.Text == "Link");
        double underlineHeight = glyphRun.UnderlineHeight.GetValueOrDefault();
        TestAssert.True(underlineHeight > 0d && underlineHeight < 0.5d, "Expected PPTX underline fill thickness to use font metrics instead of the PDF stroke minimum.");

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 0 1 rg", pdf);
        TestAssert.Contains("re f*", pdf);
        TestAssert.Contains(" TJ", pdf);
    }

    public static void PptxTextModelInheritsPlaceholderBodyProperties()
    {
        string contentTypes = PptxTests.BasicContentTypes().Replace(
            "</Types>",
            """
              <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
              <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
            </Types>
            """,
            StringComparison.Ordinal);
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = contentTypes,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
                  <p:cSld><p:spTree/></p:cSld>
                </p:sldMaster>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Layout Body"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr vert="vert270" anchor="b" anchorCtr="1" wrap="none" vertOverflow="clip" lIns="914400" tIns="457200" bIns="0" numCol="2" compatLnSpc="1" rot="5400000"><a:normAutofit fontScale="80000" lnSpcReduction="12000"/></a:bodyPr><a:lstStyle/><a:p/></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sldLayout>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="3" name="Slide Body"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
                    <p:txBody><a:bodyPr rIns="0" spcCol="914400"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Inherited body properties</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal(72d, frame.InsetLeft);
        TestAssert.Equal(0d, frame.InsetRight);
        TestAssert.Equal(36d, frame.InsetTop);
        TestAssert.Equal(0d, frame.InsetBottom);
        TestAssert.Equal(180d, frame.TextHeight);
        TestAssert.True(frame.VerticalOffset > 162d, "Expected inherited normAutofit scale to affect bottom-anchor height estimation.");
        TestAssert.Equal("914400", frame.InsetLeftValue ?? string.Empty);
        TestAssert.Equal("0", frame.InsetRightValue ?? string.Empty);
        TestAssert.Equal("457200", frame.InsetTopValue ?? string.Empty);
        TestAssert.Equal("0", frame.InsetBottomValue ?? string.Empty);
        TestAssert.Equal("InheritedBodyPr", frame.InsetLeftSource);
        TestAssert.Equal("DirectBodyPr", frame.InsetRightSource);
        TestAssert.Equal("InheritedBodyPr", frame.InsetTopSource);
        TestAssert.Equal("InheritedBodyPr", frame.InsetBottomSource);
        TestAssert.Equal("Vertical270", frame.Orientation);
        TestAssert.Equal("InheritedBodyPr", frame.OrientationSource);
        TestAssert.Equal("Bottom", frame.VerticalAnchor);
        TestAssert.Equal("InheritedBodyPr", frame.VerticalAnchorSource);
        TestAssert.Equal(true, frame.AnchorCenter);
        TestAssert.Equal("InheritedBodyPr", frame.AnchorCenterSource);
        TestAssert.Equal("None", frame.WrapMode);
        TestAssert.Equal("InheritedBodyPr", frame.WrapSource);
        TestAssert.Equal("Clip", frame.VerticalOverflow);
        TestAssert.Equal("InheritedBodyPr", frame.VerticalOverflowSource);
        TestAssert.Equal(2, frame.ColumnCount);
        TestAssert.Equal(72d, frame.ColumnSpacing);
        TestAssert.Equal("DirectBodyPr", frame.ColumnSource);
        TestAssert.Equal("InheritedBodyPr", frame.ColumnCountSource);
        TestAssert.Equal("DirectBodyPr", frame.ColumnSpacingSource);
        TestAssert.Equal("2", frame.ColumnCountValue ?? string.Empty);
        TestAssert.Equal("914400", frame.ColumnSpacingValue ?? string.Empty);
        TestAssert.Equal("normAutofit", frame.AutofitModeValue);
        TestAssert.Equal("InheritedBodyPr", frame.AutofitModeSource);
        TestAssert.Equal(0.8d, frame.FontScale);
        TestAssert.Equal("80000", frame.FontScaleValue ?? string.Empty);
        TestAssert.Equal("InheritedBodyPr", frame.FontScaleSource);
        TestAssert.Equal(0.88d, frame.LineSpacingScale);
        TestAssert.Equal("12000", frame.LineSpacingReductionValue ?? string.Empty);
        TestAssert.Equal("InheritedBodyPr", frame.LineSpacingScaleSource);
        TestAssert.Equal(true, frame.CompatibleLineSpacing);
        TestAssert.Equal("1", frame.CompatibleLineSpacingValue);
        TestAssert.Equal("InheritedBodyPr", frame.CompatibleLineSpacingSource);
        TestAssert.Equal(90d, frame.RotationDegrees ?? double.NaN);
        TestAssert.Equal("5400000", frame.RotationValue ?? string.Empty);
        TestAssert.Equal("InheritedBodyPr", frame.RotationDegreesSource);
    }

    public static void PptxTextModelCascadesPlaceholderBodyPropertiesAcrossLayoutAndMaster()
    {
        string contentTypes = PptxTests.BasicContentTypes().Replace(
            "</Types>",
            """
              <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
              <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
            </Types>
            """,
            StringComparison.Ordinal);
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = contentTypes,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="1" name="Master Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr anchor="ctr" lIns="91440" tIns="45720" rIns="91440" bIns="45720"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p/></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sldMaster>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Layout Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sldLayout>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="3" name="Slide Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Master anchored title</a:t></a:r></a:p></p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal(2, frame.InheritedPlaceholderCount);
        TestAssert.Equal(0d, frame.InsetTop);
        TestAssert.Equal("DirectBodyPr", frame.InsetTopSource);
        TestAssert.Equal("Middle", frame.VerticalAnchor);
        TestAssert.Equal("ctr", frame.VerticalAnchorValue ?? string.Empty);
        TestAssert.Equal("InheritedBodyPr", frame.VerticalAnchorSource);
        TestAssert.True(frame.VerticalOffset > 20d, "Expected master title anchor to survive an empty layout placeholder bodyPr.");
    }

    public static void PptxSyntheticTextBoxResolvesThemeEaAndCsFonts()
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFrameModelSnapshot frame = PptxRenderer.InspectTextFrameModels(document, package, 0).Single();
        TestAssert.Equal("Microsoft YaHei", frame.Paragraphs[0].Runs[0].Typeface);
        TestAssert.Equal("MajorEastAsian", frame.Paragraphs[0].Runs[0].TypefaceSource);
        TestAssert.Equal("Tahoma", frame.Paragraphs[0].Runs[2].Typeface);
        TestAssert.Equal("MinorComplexScript", frame.Paragraphs[0].Runs[2].TypefaceSource);

        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneTextParagraph sceneParagraph = scene.Slides[0].SlideNodes[0].TextBody!.Paragraphs[0];
        TestAssert.Equal("Microsoft YaHei", sceneParagraph.Runs[0].ResolvedStyle.Typeface ?? string.Empty);
        TestAssert.Equal(PptxThemeTypefaceSource.MajorEastAsian, sceneParagraph.Runs[0].ResolvedStyle.TypefaceSource);
        TestAssert.Equal("Tahoma", sceneParagraph.Runs[2].ResolvedStyle.Typeface ?? string.Empty);
        TestAssert.Equal(PptxThemeTypefaceSource.MinorComplexScript, sceneParagraph.Runs[2].ResolvedStyle.TypefaceSource);
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

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
        TestAssert.True(Math.Abs(wordStarts[1] - 151.19d) < 0.1d, "Expected justified wrapping to ignore the carried line-end separator when distributing stretchable spaces.");
        TestAssert.True(Math.Abs(wordStarts[2] - 191.51d) < 0.1d, "Expected justified wrapping to keep Office-like word starts after excluding trailing wrap spaces.");
        TestAssert.True(Math.Abs(wordStarts[3] - 257.15d) < 0.1d, "Expected justified wrapping to keep Office-like cumulative word spacing.");
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
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
                node.ShapeHasCustomGeometry,
                node.ShapeHasUnsupportedCustomGeometry,
                node.ShapeHasGradientSource,
                node.ShapeHasUnsupportedGradient,
                node.ShapeHasPatternSource,
                node.ShapeHasUnsupportedPattern,
                node.ShapeNoFill,
                node.ShapeLineNoFill,
                node.ShapeHasEffectList,
                node.ShapeHasEffectDag,
                node.ShapeUnsupportedEffectCount,
                node.ShapeUnsupportedEffectNames,
                node.HasTextBody,
                node.TextParagraphCount,
                node.TextRunCount,
                node.TextHasUnsupportedOrientation,
                node.TextHasUnsupportedVerticalOverflow,
                node.HasPicture,
                node.PictureRecolorKind,
                node.HasTable,
                node.TableRowCount,
                node.TableCellCount,
                node.HasChart,
                node.ChartPlotCount,
                node.ChartAxisCount,
                node.ChartSeriesCount,
                node.ChartSeriesMarkerCount,
                node.ChartSeriesPointStyleCount,
                node.ChartSeriesPointExplosionCount,
                node.ChartDataLabelsDefinedCount,
                node.ChartDataLabelOverrideCount,
                node.ChartDataLabelManualLayoutCount,
                node.ChartTextBodyOrientationCount,
                node.ChartTextBodyVerticalOverflowCount,
                node.HasChartColorStyle,
                node.ChartColorStyleMethod,
                node.ChartColorStyleColorCount,
                node.ChartColorStyleVariationCount,
                node.ChartColorStyleDeclarationCount,
                node.ChartColorStyleRootDeclarationCount,
                node.ChartColorStyleResolvedDeclarationCount,
                node.ChartColorStyleDeclarationKinds,
                node.HasChartStylePart,
                node.ChartStyleEntryCount,
                node.ChartStyleEntryRoles,
                node.ChartStyleEntrySourceIndexes,
                node.ChartStyleEntryNamespaceUris,
                node.ChartStyleShapeStyleCount,
                node.ChartStyleShapeFillCount,
                node.ChartStyleFillReferenceCount,
                node.ChartStyleResolvedFillReferenceCount,
                node.ChartStyleEffectReferenceCount,
                node.ChartStyleResolvedEffectReferenceCount,
                node.ChartStyleFontReferenceCount,
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
                    models[frameIndex].TextHeight,
                    models[frameIndex].VerticalOffset,
                    models[frameIndex].InsetLeft,
                    models[frameIndex].InsetRight,
                    models[frameIndex].InsetTop,
                    models[frameIndex].InsetBottom,
                    models[frameIndex].InsetLeftValue,
                    models[frameIndex].InsetRightValue,
                    models[frameIndex].InsetTopValue,
                    models[frameIndex].InsetBottomValue,
                    models[frameIndex].FontScale,
                    models[frameIndex].FontScaleValue,
                    models[frameIndex].FontScaleSource,
                    models[frameIndex].LineSpacingScale,
                    models[frameIndex].LineSpacingReductionValue,
                    models[frameIndex].LineSpacingScaleSource,
                    models[frameIndex].CompatibleLineSpacing,
                    models[frameIndex].CompatibleLineSpacingValue,
                    models[frameIndex].CompatibleLineSpacingSource,
                    models[frameIndex].RotationDegrees,
                    models[frameIndex].RotationValue,
                    models[frameIndex].RotationDegreesSource,
                    models[frameIndex].InheritedPlaceholderCount,
                    models[frameIndex].HasInheritedTextBody,
                    models[frameIndex].UsesInheritedShapeBounds,
                    models[frameIndex].InsetLeftSource,
                    models[frameIndex].InsetRightSource,
                    models[frameIndex].InsetTopSource,
                    models[frameIndex].InsetBottomSource,
                    models[frameIndex].OrientationSource,
                    models[frameIndex].VerticalAnchorSource,
                    models[frameIndex].AnchorCenterSource,
                    models[frameIndex].WrapSource,
                    models[frameIndex].VerticalOverflowSource,
                    models[frameIndex].ColumnSource,
                    models[frameIndex].ColumnCountSource,
                    models[frameIndex].ColumnSpacingSource,
                    models[frameIndex].ColumnCountValue,
                    models[frameIndex].ColumnSpacingValue,
                    models[frameIndex].AutofitModeValue,
                    models[frameIndex].AutofitModeSource,
                    paragraphs = models[frameIndex].Paragraphs.Select(paragraph => new
                    {
                        paragraph.Level,
                        paragraph.Alignment,
                        paragraph.FontSize,
                        paragraph.CascadeLevelName,
                        paragraph.ResolvedCascadeSourceCount,
                        paragraph.CascadeLayerNames,
                        paragraph.CascadeLayerKinds,
                        paragraph.ResolvedStyleSourceCount,
                        paragraph.ResolvedStyleLayerNames,
                        paragraph.ResolvedStyleLayerKinds,
                        runs = paragraph.Runs.Select(run => new
                        {
                            run.Kind,
                            length = run.Text.Length,
                            hash = PptxTests.ShortHash(run.Text),
                            run.ResolvedCascadeSourceCount,
                            run.CascadeLayerNames,
                            run.CascadeLayerKinds,
                            run.FontSize,
                            run.CharacterSpacing,
                            run.Typeface,
                            run.ColorSource,
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
                            hash = PptxTests.ShortHash(run.SourceText),
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
                        hash = PptxTests.ShortHash(span.Text),
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
                        avgPositiveAdjustment = PptxTests.AveragePositiveAdjustment(span.GlyphSpan.Glyphs),
                        avgNegativeAdjustment = PptxTests.AverageNegativeAdjustment(span.GlyphSpan.Glyphs)
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
                hash = PptxTests.ShortHash(run.Text),
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
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
                  </p:sp><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="5029200" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:schemeClr val="bg1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("0.651 g", pdf);
        TestAssert.Contains("0.498 g", pdf);
        TestAssert.Contains("0.067 g", pdf);
    }

    public static void PptxThemeColorMapOwnsSchemeAliases()
    {
        PptxColorMap defaultMap = PptxColorMap.Default;
        TestAssert.Equal("lt1", defaultMap.ResolveSchemeColor("bg1"));
        TestAssert.Equal("dk1", defaultMap.ResolveSchemeColor("tx1"));
        TestAssert.Equal("accent1", defaultMap.ResolveSchemeColor("accent1"));

        XElement customMap = XElement.Parse("""
            <p:clrMap xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                      bg1="accent2"
                      tx1="accent3"
                      accent1="hlink"/>
            """);
        PptxColorMap map = PptxColorMap.FromElement(customMap);

        TestAssert.Equal("accent2", map.ResolveSchemeColor("bg1"));
        TestAssert.Equal("accent3", map.ResolveSchemeColor("tx1"));
        TestAssert.Equal("hlink", map.ResolveSchemeColor("accent1"));
        TestAssert.Equal("accent2", map.ResolveSchemeColor("accent2"));
        TestAssert.Equal("dk2", map.ResolveSchemeColor("dk2"));
    }

    public static void PptxThemeColorMapPreservesSceneOverrideProvenance()
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
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
                </Relationships>
                """,
            ["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdMaster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
                </Relationships>
                """,
            ["ppt/slideMasters/slideMaster1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                  <p:clrMap bg1="accent1" tx1="accent2" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
                </p:sldMaster>
                """,
            ["ppt/slideLayouts/slideLayout1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                  <p:clrMapOvr><a:overrideClrMapping bg1="accent3" tx1="accent4"/></p:clrMapOvr>
                </p:sldLayout>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree/></p:cSld>
                  <p:clrMapOvr><a:overrideClrMapping bg1="accent5" tx1="accent6"/></p:clrMapOvr>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxSceneSnapshot snapshot = PptxRenderer.InspectScene(document, package);
        PptxSceneSlideSnapshot slide = snapshot.Slides[0];

        TestAssert.Equal("accent1", slide.MasterColorMap["bg1"]);
        TestAssert.Equal("accent2", slide.MasterColorMap["tx1"]);
        TestAssert.Equal("accent3", slide.LayoutColorMap["bg1"]);
        TestAssert.Equal("accent4", slide.LayoutColorMap["tx1"]);
        TestAssert.Equal("accent2", slide.LayoutColorMap["accent2"]);
        TestAssert.Equal("accent5", slide.SlideColorMap["bg1"]);
        TestAssert.Equal("accent6", slide.SlideColorMap["tx1"]);
        TestAssert.Equal("accent2", slide.SlideColorMap["accent2"]);
    }

    public static void PptxThemeColorMapResolvesSceneTextAndShapeColors()
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
                  <Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
                  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
                  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                </Relationships>
                """,
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """,
            ["ppt/charts/_rels/chart1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
                  <Relationship Id="rIdStyle" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
                </Relationships>
                """,
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/theme/theme1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ColorMapText">
                  <a:themeElements>
                    <a:clrScheme name="ColorMapText">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent5><a:srgbClr val="112233"/></a:accent5>
                      <a:accent6><a:srgbClr val="445566"/></a:accent6>
                    </a:clrScheme>
                    <a:fontScheme name="ColorMapText"><a:majorFont/><a:minorFont/></a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                      <a:solidFill><a:schemeClr val="bg1"/></a:solidFill>
                      <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>
                    </p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>Mapped</a:t></a:r></a:p></p:txBody>
                  </p:sp>
                  <p:graphicFrame>
                    <p:xfrm><a:off x="914400" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>
                    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                      <a:tbl>
                        <a:tblPr firstRow="1"><a:tableStyleId>{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}</a:tableStyleId></a:tblPr>
                        <a:tblGrid><a:gridCol w="914400"/></a:tblGrid>
                        <a:tr h="914400"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:rPr><a:t>CellMapped</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:lnL><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:lnL></a:tcPr></a:tc></a:tr>
                      </a:tbl>
                    </a:graphicData></a:graphic>
                  </p:graphicFrame>
                  <p:pic>
                    <p:blipFill><a:blip><a:duotone><a:schemeClr val="bg1"/><a:schemeClr val="tx1"/></a:duotone></a:blip></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="1828800" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
                  </p:pic>
                  <p:graphicFrame>
                    <p:xfrm><a:off x="2743200" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>
                    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
                  </p:graphicFrame></p:spTree></p:cSld>
                  <p:clrMapOvr><a:overrideClrMapping bg1="accent5" tx1="accent6"/></p:clrMapOvr>
                </p:sld>
                """,
            ["ppt/charts/chart1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <c:spPr>
                    <a:pattFill prst="pct25">
                      <a:fgClr><a:schemeClr val="bg1"/></a:fgClr>
                      <a:bgClr><a:schemeClr val="tx1"/></a:bgClr>
                    </a:pattFill>
                    <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>
                    <a:effectLst><a:glow rad="63500"><a:schemeClr val="bg1"/></a:glow></a:effectLst>
                  </c:spPr>
                  <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                  <c:chart>
                    <c:title>
                      <c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></c:spPr>
                      <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                    </c:title>
                    <c:plotArea>
                      <c:spPr><a:ln><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln></c:spPr>
                      <c:barChart>
                        <c:barDir val="col"/>
                        <c:grouping val="clustered"/>
                        <c:ser>
                          <c:idx val="0"/>
                          <c:order val="0"/>
                          <c:spPr>
                            <a:pattFill prst="pct50">
                              <a:fgClr><a:schemeClr val="bg1"/></a:fgClr>
                              <a:bgClr><a:schemeClr val="tx1"/></a:bgClr>
                            </a:pattFill>
                            <a:solidFill><a:schemeClr val="bg1"/></a:solidFill>
                            <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>
                          </c:spPr>
                          <c:marker>
                            <c:spPr>
                              <a:solidFill><a:schemeClr val="bg1"/></a:solidFill>
                              <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>
                            </c:spPr>
                          </c:marker>
                          <c:dPt>
                            <c:idx val="0"/>
                            <c:spPr>
                              <a:pattFill prst="ltHorz">
                                <a:fgClr><a:schemeClr val="tx1"/></a:fgClr>
                                <a:bgClr><a:schemeClr val="bg1"/></a:bgClr>
                              </a:pattFill>
                              <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
                              <a:ln><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln>
                            </c:spPr>
                          </c:dPt>
                          <c:dLbls>
                            <c:spPr>
                              <a:solidFill><a:schemeClr val="tx1"/></a:solidFill>
                              <a:ln><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln>
                            </c:spPr>
                            <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                            <c:leaderLines><c:spPr><a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></c:spPr></c:leaderLines>
                            <c:dLbl>
                              <c:idx val="0"/>
                              <c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:rPr><a:t>Label</a:t></a:r></a:p></c:rich></c:tx>
                              <c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></c:spPr>
                              <c:leaderLines><c:spPr><a:ln><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln></c:spPr></c:leaderLines>
                            </c:dLbl>
                          </c:dLbls>
                          <c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                        <c:axId val="10"/>
                        <c:axId val="20"/>
                      </c:barChart>
                      <c:catAx>
                        <c:axId val="10"/>
                        <c:scaling><c:orientation val="minMax"/></c:scaling>
                        <c:axPos val="b"/>
                        <c:majorGridlines><c:spPr><a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
                        <c:title>
                          <c:spPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></c:spPr>
                          <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                        </c:title>
                        <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:pPr></a:p></c:txPr>
                        <c:spPr><a:ln><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:ln></c:spPr>
                        <c:crossAx val="20"/>
                      </c:catAx>
                      <c:valAx>
                        <c:axId val="20"/>
                        <c:scaling><c:orientation val="minMax"/></c:scaling>
                        <c:axPos val="l"/>
                        <c:crossAx val="10"/>
                      </c:valAx>
                    </c:plotArea>
                    <c:legend>
                      <c:spPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></c:spPr>
                    </c:legend>
                  </c:chart>
                </c:chartSpace>
                """,
            ["ppt/charts/colors1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                               meth="cycle" id="5">
                  <a:schemeClr val="bg1"/>
                  <a:schemeClr val="tx1"/>
                </cs:colorStyle>
                """,
            ["ppt/charts/style1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <cs:style xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          id="5">
                  <cs:title>
                    <cs:spPr>
                      <a:solidFill><a:schemeClr val="bg1"/></a:solidFill>
                      <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>
                    </cs:spPr>
                    <cs:defRPr><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></cs:defRPr>
                  </cs:title>
                  <cs:legend>
                    <cs:fontRef idx="minor"><a:schemeClr val="bg1"/></cs:fontRef>
                  </cs:legend>
                </cs:style>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode node = scene.Slides[0].SlideNodes[0];
        PptxSceneRunStyle style = node.TextBody!.Paragraphs[0].Runs[0].ResolvedStyle;

        TestAssert.Equal(new RgbColor(17, 34, 51), style.Color);
        IReadOnlyList<PptxTextRunSnapshot> textRuns = PptxRenderer.InspectTextRuns(document, package, 0);
        TestAssert.Equal(new RgbColor(17, 34, 51), textRuns.Single(run => run.Text == "Mapped").Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), textRuns.Single(run => run.Text == "CellMapped").Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), node.Shape!.Fill.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), node.Shape.Line.Color);
        PptxSceneTableCell cell = scene.Slides[0].SlideNodes[1].Table!.Rows[0].Cells[0];
        TestAssert.Equal(new RgbColor(17, 34, 51), cell.Fill.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), cell.Borders.Left.Line.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), cell.StyleFill.Color);
        TestAssert.Equal(new RgbColor(255, 255, 255), cell.StyleText.Color ?? default);
        PptxSceneImageRecolor recolor = scene.Slides[0].SlideNodes[2].Picture!.Recolor;
        TestAssert.Equal(new RgbColor(17, 34, 51), recolor.Dark);
        TestAssert.Equal(new RgbColor(68, 85, 102), recolor.Light);
        PptxSceneChart chart = scene.Slides[0].SlideNodes[3].Chart ?? throw new InvalidOperationException("Expected chart scene.");
        TestAssert.Equal("accent5", chart.ColorMap.ResolveSchemeColor("bg1"));
        TestAssert.Equal("accent6", chart.ColorMap.ResolveSchemeColor("tx1"));
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.ChartAreaStyle.PatternFill.Foreground);
        TestAssert.Equal(new RgbColor(68, 85, 102), chart.ChartAreaStyle.PatternFill.Background);
        TestAssert.Equal(new RgbColor(68, 85, 102), chart.ChartAreaStyle.Line.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.ChartAreaStyle.Glow.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.TextStyle.Color ?? default);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.Title.ShapeStyle.Fill.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), chart.Title.TextStyle.Color ?? default);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.PlotAreaStyle.Line.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), chart.Legend.ShapeStyle.Fill.Color);
        PptxSceneChartSeries series = chart.Plots[0].Series[0];
        TestAssert.Equal(new RgbColor(17, 34, 51), series.Fill.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), series.PatternFill.Foreground);
        TestAssert.Equal(new RgbColor(68, 85, 102), series.PatternFill.Background);
        TestAssert.Equal(new RgbColor(68, 85, 102), series.Line.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), series.Marker.Fill.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), series.Marker.Line.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), series.PointStyles[0].Fill.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), series.PointStyles[0].PatternFill.Foreground);
        TestAssert.Equal(new RgbColor(17, 34, 51), series.PointStyles[0].PatternFill.Background);
        TestAssert.Equal(new RgbColor(17, 34, 51), series.PointStyles[0].Line.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), series.DataLabels.TextStyle.Color ?? default);
        TestAssert.Equal(new RgbColor(68, 85, 102), series.DataLabels.ShapeStyle.Fill.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), series.DataLabels.ShapeStyle.Line.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), series.DataLabels.LeaderLines.Line.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), series.DataLabels.Overrides[0].CustomTextRuns[0].TextStyle.Color ?? default);
        TestAssert.Equal(new RgbColor(17, 34, 51), series.DataLabels.Overrides[0].ShapeStyle.Fill.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), series.DataLabels.Overrides[0].LeaderLines.Line.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.Axes[0].Line.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), chart.Axes[0].MajorGridlineLine.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.Axes[0].TextStyle.Color ?? default);
        TestAssert.Equal(new RgbColor(68, 85, 102), chart.Axes[0].Title.ShapeStyle.Fill.Color);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.Axes[0].Title.TextStyle.Color ?? default);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.ColorStyle.Colors[0]);
        TestAssert.Equal(new RgbColor(68, 85, 102), chart.ColorStyle.Colors[1]);
        TestAssert.Equal(new RgbColor(17, 34, 51), chart.ColorStyle.Declarations[0].Color ?? default);
        TestAssert.Equal(new RgbColor(68, 85, 102), chart.ColorStyle.Declarations[1].Color ?? default);
        PptxSceneChartStyleEntry titleStyle = chart.StylePart.Entries.First(entry => entry.Role == "title");
        TestAssert.Equal(new RgbColor(17, 34, 51), titleStyle.ShapeStyle.Fill.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), titleStyle.ShapeStyle.Line.Color);
        TestAssert.Equal(new RgbColor(68, 85, 102), titleStyle.TextStyle.Color ?? default);
        PptxSceneChartStyleEntry legendStyle = chart.StylePart.Entries.First(entry => entry.Role == "legend");
        TestAssert.Equal(new RgbColor(17, 34, 51), legendStyle.TextStyle.Color ?? default);
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
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
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
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        TestAssert.Contains("1 g", pdf);
    }
}
