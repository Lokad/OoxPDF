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

internal static class PptxCompositionTests
{
    public static void PptxSyntheticGroupedConnectorHonorsGroupFlip()
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

    public static void PptxSyntheticUngroupedStealthConnectorUsesFilledBody()
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
                            <a:xfrm><a:off x="1200000" y="1500000"/><a:ext cx="2400000" cy="0"/></a:xfrm>
                            <a:prstGeom prst="straightConnector1"/>
                            <a:ln w="12700"><a:solidFill><a:srgbClr val="2F856A"/></a:solidFill><a:tailEnd type="stealth"/></a:ln>
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
        // Office fills stealth-tailed straight bodies (1pt and 2pt fill 6pt wide); stroked lines stay legacy for transformed shapes.
        TestAssert.DoesNotContain(" l S", pdf);
        TestAssert.Contains("0.184 0.522 0.416 rg", pdf);

    }

    public static void PptxGarbageGradientVectorDiagnosesAndKeepsSiblings()
    {
        string input = WriteGarbageGradientDeck("""
            <p:pic>
              <p:nvPicPr><p:cNvPr id="2" name="Malformed SVG"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
              <p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip></p:blipFill>
              <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
            </p:pic>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="3" name="Sibling"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr>
                <a:xfrm><a:off x="1828800" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                <a:prstGeom prst="rect"/>
                <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
              </p:spPr>
            </p:sp>
            """);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();

        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        TestAssert.Contains("0 1 0 rg", File.ReadAllText(output, Encoding.ASCII));
        AssertGarbageGradientDiagnostic(diagnostics);
    }

    public static void PptxGroupHyperlinkSurvivesDiagnosedChildFallback()
    {
        const string link = "https://example.invalid/group";
        string input = WriteGarbageGradientDeck("""
            <p:grpSp>
              <p:nvGrpSpPr><p:cNvPr id="10" name="Group"><a:hlinkClick r:id="rIdLink"/></p:cNvPr><p:nvPr/></p:nvGrpSpPr>
              <p:grpSpPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="3657600" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="4572000" cy="2286000"/></a:xfrm></p:grpSpPr>
              <p:sp>
                <p:nvSpPr><p:cNvPr id="11" name="GroupedSibling"/><p:nvPr/></p:nvSpPr>
                <p:spPr>
                  <a:xfrm><a:off x="914400" y="457200"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"/>
                  <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                </p:spPr>
              </p:sp>
              <p:pic>
                <p:nvPicPr><p:cNvPr id="12" name="Malformed SVG"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
                <p:blipFill><a:blip><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst></a:blip></p:blipFill>
                <p:spPr><a:xfrm><a:off x="2743200" y="457200"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
              </p:pic>
            </p:grpSp>
            """, $"""
            <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{link}" TargetMode="External"/>
            """);
        using FileStream stream = File.OpenRead(input);
        using OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = new PptxRenderer(null).RenderPagesAsync(document, package, diagnostics.Add, CancellationToken.None)
            .ToBlockingEnumerable().Single();

        TestAssert.Equal(link, page.Annotations.Single().Uri);
        AssertGarbageGradientDiagnostic(diagnostics);
    }

    private static void AssertGarbageGradientDiagnostic(List<OoxPdfDiagnostic> diagnostics)
    {
        TestAssert.True(
            diagnostics.Any(d => d.Id == "SVG_UNSUPPORTED_CONTENT" && d.Message.Contains("unresolvable", StringComparison.Ordinal)),
            "Garbage gradient vectors must diagnose as unresolvable.");
        TestAssert.True(
            !diagnostics.Any(d => d.Id == "PPTX_NODE_RENDER_FAILED"),
            "Graceful gradient fallback must not fail the picture node.");
    }

    private static string WriteGarbageGradientDeck(string items, string extraRelationships = "") =>
        TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes().Replace("</Types>",
                "<Default Extension=\"svg\" ContentType=\"image/svg+xml\"/></Types>", StringComparison.Ordinal),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = $"""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/>
                  {extraRelationships}
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = $"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                       xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">
                  <p:cSld><p:spTree>{items}</p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/media/image1.svg"] = """
                <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">
                  <defs><linearGradient id="g" x1="bad" x2="1">
                    <stop offset="0" stop-color="#FF0000"/><stop offset="1" stop-color="#0000FF"/>
                  </linearGradient></defs>
                  <path fill="url(#g)" d="M0 0H10V10H0Z"/>
                </svg>
                """
        });

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
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
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
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
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

    public static void PptxSyntheticGroupedShapeAppliesTransform()
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
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
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
        TestAssert.Contains("0 0 720 540 re W* n", pdf);
        TestAssert.True(pdf.Contains(" TJ", StringComparison.Ordinal) || pdf.Contains("> Tj", StringComparison.Ordinal), "Expected body text to be drawn as PDF text.");
    }

    public static void PptxSyntheticTextAndShapesUseSiblingOrder() =>
        AssertSiblingPaintOrder(includeUnknownFrame: false);

    public static void PptxSyntheticUnknownGraphicFrameDoesNotDisableSiblingOrder() =>
        AssertSiblingPaintOrder(includeUnknownFrame: true);

    private static void AssertSiblingPaintOrder(bool includeUnknownFrame)
    {
        const string unknownFrame = """
            <p:graphicFrame>
              <p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>
              <a:graphic><a:graphicData uri="urn:vendor:unknown-graphic"/></a:graphic>
            </p:graphicFrame>
            """;
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
                      <p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Arial"/></a:rPr><a:t>Covered</a:t></a:r></a:p></p:txBody>
                    </p:sp>
                    {(includeUnknownFrame ? unknownFrame : string.Empty)}
                    <p:sp>
                      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });
        using OoxPackage package = OoxPackage.Open(input, default);
        var document = new PptxReader().Read(package, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, diagnostics.Add, default).ToBlockingEnumerable().Single();

        var paths = new List<string>();
        RgbColor[] colors = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add, includeNeutral: true).ToArray();
        int cover = Enumerable.Range(0, paths.Count)
            .Single(index => paths[index] is "72 396 288 72 re f" or "72 396 288 72 re f*");
        TestAssert.Equal(new RgbColor(255, 0, 0), colors[cover]);
        TestAssert.True(!PptxChartPaletteTests.ReadColoredPaints(page.Content, strokes: true, includeNeutral: true).Any(),
            "Both shapes explicitly suppress their outlines.");

        int coverIndex = page.Content.IndexOf(paths[cover], StringComparison.Ordinal);
        var beforeCover = new PdfPage(page.Width, page.Height, page.Content[..coverIndex], page.Fonts);
        string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(beforeCover, usePageCoordinates: false,
            honorActualText: false).Select(show => show.Text));
        TestAssert.Equal("Covered", text);

        TestAssert.Equal(includeUnknownFrame ? 1 : 0, diagnostics.Count);
        if (includeUnknownFrame)
        {
            var diagnostic = diagnostics.Single();
            TestAssert.Equal("PPTX_UNSUPPORTED_GRAPHIC_FRAME", diagnostic.Id);
            TestAssert.Equal("/ppt/slides/slide1.xml", diagnostic.PartName);
            TestAssert.Equal(1, diagnostic.SlideIndex);
        }
    }

    public static void PptxSyntheticGroupedTableUsesGroupTransform() =>
        AssertGroupedTablePaint(includeUnknownFrame: false);

    public static void PptxSyntheticGroupedTableUsesGroupTransformWithUnknownGraphicFrame() =>
        AssertGroupedTablePaint(includeUnknownFrame: true);

    private static void AssertGroupedTablePaint(bool includeUnknownFrame)
    {
        const string unknownFrame = """
            <p:graphicFrame>
              <p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>
              <a:graphic><a:graphicData uri="urn:unknown"/></a:graphic>
            </p:graphicFrame>
            """;
        string input = TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree>
                    {(includeUnknownFrame ? unknownFrame : string.Empty)}
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
        string output = TestFixtures.NewTemporaryPath(".pdf");
        var diagnostics = new List<OoxPdfDiagnostic>();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions { DiagnosticSink = diagnostics.Add });

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        var paintedPaths = new List<string>();
        RgbColor[] colors = PptxChartPaletteTests.ReadColoredPaints(pdf, paintedPaths.Add).ToArray();
        RgbColor cellColor = colors.Where((_, i) => paintedPaths[i] is "72 396 144 72 re f" or "72 396 144 72 re f*").Single();
        TestAssert.Equal(new RgbColor(217, 234, 211), cellColor);

        OoxPdfDiagnostic[] unsupported = diagnostics.Where(d => d.Id == "PPTX_UNSUPPORTED_GRAPHIC_FRAME").ToArray();
        TestAssert.Equal(includeUnknownFrame ? 1 : 0, unsupported.Length);
        if (includeUnknownFrame)
        {
            TestAssert.Equal("/ppt/slides/slide1.xml", unsupported[0].PartName);
            TestAssert.Equal(1, unsupported[0].SlideIndex);
        }
    }

    public static void PptxSyntheticGroupedBarChartUsesGroupTransform()
    {
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
        };
        string input = TestFixtures.WriteTempPackage(".pptx", parts);
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        OoxPdfConverter.Convert(input, output);

        string pdf = File.ReadAllText(output, Encoding.ASCII);
        TestAssert.Contains("0 1 1 rg", pdf);
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XDocument slide = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
        XElement group = slide.Descendants(p + "grpSp").Single();
        XElement groupOrigin = group.Element(p + "grpSpPr")!.Element(a + "xfrm")!.Element(a + "off")!;
        groupOrigin.SetAttributeValue("x", "0");
        groupOrigin.SetAttributeValue("y", "0");
        parts["ppt/slides/slide1.xml"] = Encoding.UTF8.GetBytes(slide.ToString(SaveOptions.DisableFormatting));
        string referenceInput = TestFixtures.WriteTempPackage(".pptx", parts);
        string referenceOutput = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        OoxPdfConverter.Convert(referenceInput, referenceOutput);
        string referencePdf = File.ReadAllText(referenceOutput, Encoding.ASCII);
        Match[] groupedRects = Regex.Matches(pdf, @"(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f").Cast<Match>().ToArray();
        Match[] referenceRects = Regex.Matches(referencePdf, @"(?<x>[-\d.]+) (?<y>[-\d.]+) (?<w>[-\d.]+) (?<h>[-\d.]+) re f").Cast<Match>().ToArray();
        TestAssert.True(groupedRects.Length >= 2, "The control must paint its plot and bar rectangles.");
        TestAssert.Equal(groupedRects.Length, referenceRects.Length);
        TestAssert.Equal(referenceRects[0].Value, groupedRects[0].Value);
        for (int i = 1; i < groupedRects.Length; i++)
        foreach ((string field, double delta) in new[] { ("x", 72d), ("y", -72d), ("w", 0d), ("h", 0d) })
        {
            double grouped = double.Parse(groupedRects[i].Groups[field].Value, CultureInfo.InvariantCulture);
            double ungrouped = double.Parse(referenceRects[i].Groups[field].Value, CultureInfo.InvariantCulture);
            TestAssert.True(Math.Abs(grouped - ungrouped - delta) < .002d,
                $"The source group's72pt translation moves the plot and bar while preserving dimensions: rectangle{i}/{field}, {grouped} vs {ungrouped}, expected delta{delta}.");
        }
        File.Delete(input); File.Delete(output); File.Delete(referenceInput); File.Delete(referenceOutput);
    }
}
