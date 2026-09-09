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

internal static class PptxShapesTests
{
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
        TestAssert.Contains("1 g", pdf);
        TestAssert.Contains("0 0 720 540 re f", pdf);
        TestAssert.Contains("1 0 0 rg", pdf);
        TestAssert.Contains("72 396 144 72 re f*", pdf);
        TestAssert.Contains("0 0 1 RG", pdf);
        TestAssert.Contains("0 1 0 rg", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.Contains("72 252 m 216 180 l S", pdf);
        TestAssert.Contains("360 240 m", pdf);
        TestAssert.Contains("0.933 g", pdf);
        TestAssert.Contains("0.184 0.522 0.416 RG", pdf);
        TestAssert.Contains(" re W n", pdf);
    }

    public static void PptxSyntheticFilledStrokedShapeRepeatsSlideClipBeforeFill()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        int shapeClip = pdf.IndexOf("0 0 720 540 re W* n\r\n1 0 0 rg", StringComparison.Ordinal);
        int fill = pdf.IndexOf("72 396 144 72 re f*", StringComparison.Ordinal);
        int strokeColor = pdf.IndexOf("0 0 1 RG", StringComparison.Ordinal);
        TestAssert.True(shapeClip >= 0, "Expected the repeated slide clip to precede the filled shape region.");
        TestAssert.True(fill > shapeClip, "Expected the shape fill after the repeated slide clip.");
        TestAssert.True(strokeColor > fill, "Expected the stroke state after the filled region.");
        TestAssert.True(pdf.IndexOf("re f*\r\n0 0 720 540 re W* n", fill, StringComparison.Ordinal) < 0, "Filled/stroked shapes should not repeat the slide clip between fill and stroke.");
    }

    public static void PptxSyntheticShapeExplicitLineColorInheritsStyleLineWidth()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
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
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="LineWidthTheme">
                  <a:themeElements>
                    <a:clrScheme name="LineWidthTheme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
                    </a:clrScheme>
                    <a:fontScheme name="LineWidthTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
                    <a:fmtScheme name="LineWidthTheme">
                      <a:fillStyleLst/>
                      <a:lnStyleLst>
                        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                      </a:lnStyleLst>
                      <a:effectStyleLst/>
                      <a:bgFillStyleLst/>
                    </a:fmtScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:style>
                          <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
                          <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
                          <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
                          <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
                        </p:style>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:noFill/>
                          <a:ln><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneNode shapeNode = scene.Slides[0].SlideNodes[0];
        PptxSceneNodeSnapshot shapeSnapshot = PptxRenderer.InspectScene(document, package).Slides[0].SlideNodes[0];
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(shapeNode.Shape?.NoFill == true, "Expected explicit shape noFill provenance in the scene model.");
        TestAssert.True(shapeNode.Shape?.Fill.HasFill == false, "Expected explicit shape noFill to suppress scene solid fill resolution.");
        TestAssert.Equal(0, shapeNode.Shape?.FillReference.Index ?? -1);
        TestAssert.True(shapeNode.Shape?.FillReference.Reference is not null, "Expected fillRef provenance even when idx=0 does not resolve a theme style.");
        TestAssert.True(shapeNode.Shape?.FillReference.Style is null, "Expected idx=0 fillRef to remain unresolved against the theme fill style list.");
        TestAssert.Equal(2, shapeNode.Shape?.LineReference.Index ?? 0);
        TestAssert.True(shapeNode.Shape?.LineReference.Reference is not null, "Expected line reference XML ownership in the shape scene model.");
        TestAssert.True(shapeNode.Shape?.LineReference.Style is not null, "Expected line reference to resolve through the theme line style list.");
        TestAssert.True(shapeSnapshot.ShapeNoFill, "Expected private-safe inspection to expose explicit shape noFill.");
        TestAssert.Equal(0, shapeSnapshot.ShapeFillReferenceIndex);
        TestAssert.True(!shapeSnapshot.ShapeFillReferenceResolved, "Expected private-safe inspection to expose unresolved fillRef state.");
        TestAssert.Equal(2, shapeSnapshot.ShapeLineReferenceIndex);
        TestAssert.True(shapeSnapshot.ShapeLineReferenceResolved, "Expected private-safe inspection to expose resolved lnRef state.");
        TestAssert.Contains("1.5 w", pdf);
        TestAssert.Contains("0.753 0 0 RG", pdf);
        TestAssert.Contains("72 396 144 72 re S", pdf);
    }

    public static void PptxSyntheticShapeExplicitLineWithoutWidthUsesOfficeDefault()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:sp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="rect"/>
                          <a:noFill/>
                          <a:ln><a:solidFill><a:srgbClr val="336699"/></a:solidFill></a:ln>
                        </p:spPr>
                      </p:sp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(scene.Slides[0].SlideNodes[0].Shape?.Line.HasLine == true, "Expected explicit <a:ln> to produce a scene line.");
        TestAssert.Equal(0.75d, scene.Slides[0].SlideNodes[0].Shape?.Line.Width ?? 0d);
        TestAssert.Contains("0.75 w", pdf);
        TestAssert.Contains("0.2 0.4 0.6 RG", pdf);
        TestAssert.Contains("72 396 144 72 re S", pdf);
    }

    public static void PptxSyntheticConnectorExplicitLineWidthInheritsStyleLineColor()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
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
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="LineCascadeTheme">
                  <a:themeElements>
                    <a:clrScheme name="LineCascadeTheme">
                      <a:dk1><a:srgbClr val="000000"/></a:dk1>
                      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
                    </a:clrScheme>
                    <a:fontScheme name="LineCascadeTheme"><a:majorFont/><a:minorFont/></a:fontScheme>
                    <a:fmtScheme name="LineCascadeTheme">
                      <a:fillStyleLst/>
                      <a:lnStyleLst>
                        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                      </a:lnStyleLst>
                      <a:effectStyleLst/>
                      <a:bgFillStyleLst/>
                    </a:fmtScheme>
                  </a:themeElements>
                </a:theme>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:style>
                          <a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>
                          <a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
                          <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
                          <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
                        </p:style>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="0" cy="1828800"/></a:xfrm>
                          <a:prstGeom prst="line"/>
                          <a:ln w="12700"/>
                        </p:spPr>
                      </p:cxnSp>
                    </p:spTree>
                  </p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.True(scene.Slides[0].SlideNodes[0].Shape?.Line.HasLine == true, "Expected style line color to complete the explicit connector line.");
        TestAssert.Equal(1d, scene.Slides[0].SlideNodes[0].Shape?.Line.Width ?? 0d);
        TestAssert.Equal(new RgbColor(68, 114, 196), scene.Slides[0].SlideNodes[0].Shape?.Line.Color ?? default);
        TestAssert.Contains("1 w", pdf);
        TestAssert.Contains("0.267 0.447 0.769 RG", pdf);
        TestAssert.Contains("72 468 m 72 324 l S", pdf);
    }

    public static void PptxSyntheticArrowAndConnectorShapesRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        TestAssert.Contains("0.133 g", pdf);
        TestAssert.Contains("72 324.5 m", pdf);
        TestAssert.Contains("212.667 323.5 l", pdf);
        TestAssert.Contains("216 324 m", pdf);
        TestAssert.Contains("212 325.5 l", pdf);
        TestAssert.Contains("212 322.5 l", pdf);
        TestAssert.Contains("1 -0 -0 -1 0 864 cm", pdf);
        TestAssert.Contains("216.5 468 m", pdf);
        TestAssert.Contains("216.5 400 l", pdf);
        TestAssert.Contains("215.5 468 l", pdf);
        TestAssert.Contains("216 396 m", pdf);
        TestAssert.Contains("219 402 l", pdf);
        TestAssert.Contains("213 402 l", pdf);
        TestAssert.DoesNotContain("216 468 m 216 396 l S", pdf);
        TestAssert.Contains("0.184 0.522 0.416 rg", pdf);
        TestAssert.Contains("0.753 0 0 rg", pdf);
        TestAssert.Contains("90 468 m", pdf);
        TestAssert.Contains("126 432 l", pdf);
        TestAssert.Contains("108 396 l", pdf);
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
    }

    public static void PptxSyntheticVerticalTriangleConnectorUsesOfficeMarkerWidth()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="0" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="line"/>
                          <a:ln w="38100"><a:solidFill><a:srgbClr val="156082"/></a:solidFill><a:headEnd type="triangle" w="med" len="med"/></a:ln>
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
        TestAssert.Contains("288 468 m", pdf);
        TestAssert.Contains("283.5 456 l", pdf);
        TestAssert.Contains("292.5 456 l", pdf);
        TestAssert.DoesNotContain("282 456 l", pdf);
        TestAssert.DoesNotContain("294 456 l", pdf);
    }

    public static void PptxSyntheticStraightConnectorTriangleUsesOfficeMarkerWidth()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="3657600" y="914400"/><a:ext cx="0" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="straightConnector1"/>
                          <a:ln w="38100"><a:solidFill><a:srgbClr val="156082"/></a:solidFill><a:headEnd type="triangle" w="med" len="med"/></a:ln>
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
        TestAssert.Contains("288 468 m", pdf);
        TestAssert.Contains("282 456 l", pdf);
        TestAssert.Contains("294 456 l", pdf);
        TestAssert.DoesNotContain("283.5 456 l", pdf);
        TestAssert.DoesNotContain("292.5 456 l", pdf);
    }

    public static void PptxSyntheticCurvedConnectorRendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        TestAssert.Contains("0.4 g", pdf);
        TestAssert.Contains("h\r\nf", pdf);
        TestAssert.DoesNotContain("144 360 180 324 216 324 c\r\nS", pdf);
    }

    public static void PptxSyntheticCurvedConnector2RendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        TestAssert.Contains("0.867 g", pdf);
        TestAssert.Contains("h\r\nf", pdf);
        TestAssert.DoesNotContain("111.765 468 144 435.765 144 396 c\r\nS", pdf);
    }

    public static void PptxSyntheticCurvedConnector2LoopUsesQuarterTurnTangents()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        TestAssert.Contains("h\r\nf", pdf);
        TestAssert.DoesNotContain("111.765 468 144 435.765 144 396 c\r\nS", pdf);
    }

    public static void PptxSyntheticCurvedConnector2TriangleTailUsesOfficeSeparateMarkerSubpath()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
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
        TestAssert.Contains("0.867 g", pdf);
        TestAssert.True(Regex.Matches(pdf, " m\r\n").Count >= 2, "Expected separate body and triangle marker subpaths.");
        TestAssert.True(Regex.Matches(pdf, " l\r\n").Count >= 98, "Expected Office-like dense connector body sampling.");
        TestAssert.True(Regex.Matches(pdf, "h\r\n").Count >= 2, "Expected both connector body and triangle marker to be closed before filling.");
        TestAssert.Contains("h\r\nf", pdf);
        TestAssert.DoesNotContain(" c\r\nS", pdf);
    }

    public static void PptxSyntheticCurvedConnectorArrowTailUsesFilledOutline()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld>
                    <p:spTree>
                      <p:cxnSp>
                        <p:spPr>
                          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                          <a:prstGeom prst="curvedConnector3"><a:avLst/></a:prstGeom>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="606060"/></a:solidFill><a:tailEnd type="arrow"/></a:ln>
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
        TestAssert.Contains("0.376 g", pdf);
        TestAssert.Contains(" c\r\n", pdf);
        TestAssert.True(Regex.Matches(pdf, " c\r\n").Count >= 4, "Expected Office-like arrow-tail filled connector paths to retain four cubic segments.");
        TestAssert.Contains("h\r\nf", pdf);
        TestAssert.DoesNotContain(" c\r\nS", pdf);
    }

    public static void PptxSyntheticCustomGeometryCubicPathRendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
                              <a:path w="21600" h="10800">
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

    public static void PptxSyntheticCustomGeometryOpenCubicStealthTailUsesFilledOutline()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
                          <a:noFill/>
                          <a:ln w="25400"><a:solidFill><a:srgbClr val="008000"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
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
        TestAssert.Contains("0 0.502 0 rg", pdf);
        TestAssert.Contains("h" + Environment.NewLine + "f", pdf);
        TestAssert.True(Regex.Matches(pdf, " l\r\n").Count >= 120, "Expected Office-like dense stealth-tail connector body sampling.");
        TestAssert.DoesNotContain(" c\r\nS", pdf);
    }

    public static void PptxSyntheticCustomGeometryArcPathRendersCurve()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        TestAssert.Contains("0.267 G", pdf);
        TestAssert.Contains(" c", pdf);
        TestAssert.DoesNotContain("72 396 72 72 re S", pdf);
    }

    public static void PptxSyntheticPresetArcConvertsVisualAnglesOnEllipse()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 0"/><a:gd name="adj2" fmla="val 2700000"/></a:avLst></a:prstGeom>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="444444"/></a:solidFill></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("216 432 m", pdf);
        TestAssert.Contains("176.199 399.801 c", pdf);
        TestAssert.DoesNotContain("194.912", pdf);
    }

    public static void PptxSyntheticPresetArcStealthTailUsesFilledOutline()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm flipV="1"><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="arc"><a:avLst><a:gd name="adj1" fmla="val 16200000"/><a:gd name="adj2" fmla="val 19500000"/></a:avLst></a:prstGeom>
                      <a:ln w="19050"><a:solidFill><a:srgbClr val="444444"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" f", pdf);
        TestAssert.DoesNotContain(" S", pdf);
    }

    public static void PptxSyntheticPresetArcUsesOfficeDefaultAdjustments()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="arc"><a:avLst/></a:prstGeom>
                      <a:ln w="19050"><a:solidFill><a:srgbClr val="444444"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains(" f", pdf);
        TestAssert.DoesNotContain("216 432 m 216", pdf);
        TestAssert.DoesNotContain(" S", pdf);
    }

    public static void PptxSyntheticShapeRoundRectHonorsAdjustment()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/></a:avLst></a:prstGeom>
                      <a:noFill/>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="444444"/></a:solidFill></a:ln>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("72 460.8 m", pdf);
        TestAssert.DoesNotContain("72 456.48 m", pdf);
    }

    public static void PptxSyntheticShapeStrokeDashCapAndJoinRender()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        TestAssert.True(PptxTests.CountOccurrences(pdf, "[] 0 d") == dashValues.Length - 1, "Each dashed line should reset the dash pattern; solid must not set a dash pattern.");
    }

    public static void PptxSyntheticOuterShadowRendersOffsetShape()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        TestAssert.Contains("79.2 396 72 72 re f*", pdf);
        TestAssert.Contains("72 396 72 72 re f*", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_EFFECT"), "Supported outer shadow should not emit an unsupported-effect diagnostic.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TRANSPARENCY"), "Rendered outer shadow alpha should not emit an unsupported-transparency diagnostic.");
    }

    public static void PptxSyntheticShapeRectGlowUsesRasterSoftMask()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        TestAssert.True(!pdf.Contains("64.8 388.8 86.4 86.4 re f*", StringComparison.Ordinal), "Glow must not emit an expanded solid rectangle.");
        TestAssert.Contains("/Subtype /Image", pdf);
        TestAssert.Contains("/SMask ", pdf);
        TestAssert.Contains("Do", pdf);
        TestAssert.Contains("72 396 72 72 re f*", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_EFFECT"), "Rendered rectangular glow should not emit an unsupported-effect diagnostic.");
    }

    public static void PptxSyntheticLinearGradientShapeUsesPdfShading()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
                            <a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:gs>
                            <a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="50000"/></a:srgbClr></a:gs>
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
        TestAssert.Contains("/ca 0.5", pdf);
        TestAssert.Contains("72 396 144 72 re W n", pdf);
        TestAssert.Contains("/Sh1 sh", pdf);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Supported linear gradient should not emit an unsupported-gradient diagnostic.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_TRANSPARENCY"), "Uniform gradient alpha should not emit an unsupported-transparency diagnostic.");
    }

    public static void PptxSyntheticVariableAlphaGradientShapeEmitsDiagnosticUntilSoftMaskExists()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="2" name="Gradient"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:gradFill>
                          <a:gsLst>
                            <a:gs pos="0"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:gs>
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
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_GRADIENT_FILL"), "Variable gradient stop alpha needs PDF soft-mask support before it can be marked supported.");
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_TRANSPARENCY"), "Variable gradient stop alpha should keep a transparency diagnostic until the PDF backend can express it exactly.");
    }

    public static void PptxSyntheticMultiStopLinearGradientShapeUsesStitchedPdfShading()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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

    public static void PptxSyntheticLineEndSizeVariantsAffectMarkerGeometry()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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

    public static void PptxSyntheticCenteredShapeAutoFitWrapsBeforeRightOverflow()
    {
        string cambria = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "cambria.ttc");
        if (!File.Exists(cambria))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(cambria))");
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2984500" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                    <p:txBody>
                      <a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
                      <a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1600"><a:latin typeface="Cambria Math"/></a:rPr><a:t>Forecasting models align compact labels with structural intent.</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextFlowFrameSnapshot flowFrame = PptxRenderer.InspectTextFlow(document, package, 0).Frames.Single();
        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .Single()
            .Paragraphs
            .Single()
            .Lines
            .ToArray();

        TestAssert.True(lines.Length >= 2, $"Expected wrapped centered shape-autofit text, got {lines.Length} line.");
        double textRight = flowFrame.TextX + flowFrame.TextWidth;
        int[] lineLengths = lines.Select(line => line.Spans.Sum(span => span.Text.Length)).ToArray();
        TestAssert.True(lineLengths[0] < lineLengths[1],
            $"Centered shape-autofit wrapping should break before the almost-fitting word; got line lengths {string.Join(",", lineLengths)} with textRight={textRight.ToString("0.###", CultureInfo.InvariantCulture)}, ends={string.Join(",", lines.Select(line => line.EndX.ToString("0.###", CultureInfo.InvariantCulture)))}.");
        TestAssert.True(lines[0].StartX > flowFrame.TextX + 5d,
            "Expected the first centered shape-autofit line to be re-centered after wrapping instead of filling the box through tolerated overflow.");
    }

    public static void PptxSyntheticVerticalShapeAutoFitPrefersSingleLine()
    {
        string arial = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
        if (!File.Exists(arial))
        {
            TestAssert.Skip("Environmental precondition not met: (!File.Exists(arial))");
        }

        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);

        PptxTextLineLayoutSnapshot[] lines = PptxRenderer.InspectTextLayout(document, package, 0)
            .Frames
            .SelectMany(frame => frame.Paragraphs)
            .SelectMany(paragraph => paragraph.Lines)
            .ToArray();

        TestAssert.True(lines.Length == 1, "Vertical spAutoFit text should shrink before accepting an avoidable word wrap.");
        TestAssert.True(lines[0].EndX - lines[0].StartX <= 150d, "Fitted vertical text should stay inside the rotated text width.");
    }

    public static void PptxSyntheticRotatedGroupRotatesChildShape()
    {
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
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
}
