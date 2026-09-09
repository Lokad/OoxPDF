using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxHyperlinksTests
{
    public static void PptxShapeExternalHyperlinkEmitsUriAnnotation()
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
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/shape" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Linked"><a:hlinkClick r:id="rIdLink"/></p:cNvPr><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input);

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/shape", annotation.Uri);
        TestAssert.Equal(72d, annotation.X);
        TestAssert.Equal(396d, annotation.Y);
        TestAssert.Equal(144d, annotation.Width);
        TestAssert.Equal(72d, annotation.Height);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable external shape hyperlink should not emit hyperlink diagnostics.");
    }

    public static void PptxGroupedShapeHyperlinkUsesTransformedBounds()
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
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/grouped" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:grpSp>
                    <p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:nvPr/></p:nvGrpSpPr>
                    <p:grpSpPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="3657600" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="4572000" cy="2286000"/></a:xfrm></p:grpSpPr>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="11" name="GroupedLinked"><a:hlinkClick r:id="rIdLink"/></p:cNvPr><p:nvPr/></p:nvSpPr>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="457200"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                      </p:spPr>
                    </p:sp>
                  </p:grpSp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input);

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/grouped", annotation.Uri);
        AssertNear(129.6d, annotation.X, "Grouped hyperlink X should apply the 0.8 group scale to the child offset.");
        AssertNear(417.6d, annotation.Y, "Grouped hyperlink Y should apply the 0.8 group scale and flip to PDF coordinates.");
        AssertNear(115.2d, annotation.Width, "Grouped hyperlink width should apply the 0.8 group scale.");
        AssertNear(57.6d, annotation.Height, "Grouped hyperlink height should apply the 0.8 group scale.");
    }

    public static void PptxGroupHyperlinkCoversGroupBounds()
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
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/group" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:grpSp>
                    <p:nvGrpSpPr><p:cNvPr id="10" name="LinkedGroup"><a:hlinkClick r:id="rIdLink"/></p:cNvPr><p:nvPr/></p:nvGrpSpPr>
                    <p:grpSpPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="3657600" cy="1828800"/><a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="6858000"/></a:xfrm></p:grpSpPr>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="11" name="GroupedPlain"/><p:nvPr/></p:nvSpPr>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="457200"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
                      </p:spPr>
                    </p:sp>
                  </p:grpSp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input);

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/group", annotation.Uri);
        TestAssert.Equal(72d, annotation.X);
        TestAssert.Equal(360d, annotation.Y);
        TestAssert.Equal(288d, annotation.Width);
        TestAssert.Equal(144d, annotation.Height);
    }

    public static void PptxShapeInternalSlideHyperlinkEmitsDestination()
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
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
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
                  <p:sldSz cx="9144000" cy="6858000"/>
                  <p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst>
                </p:presentation>
                """,
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide2.xml"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Jump"><a:hlinkClick r:id="rIdLink"/></p:cNvPr><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/slides/slide2.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Target"/><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input);

        TestAssert.Equal(2, pages.Count);
        PdfLinkAnnotation annotation = pages[0].Annotations.Single();
        TestAssert.True(annotation.Uri is null, "An internal slide hyperlink should not carry a URI target.");
        TestAssert.True(annotation.Destination is { PageIndex: 1 }, "An internal slide hyperlink should target the second page.");
        TestAssert.Equal(0, pages[1].Annotations.Count);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable internal slide hyperlink should not emit hyperlink diagnostics.");
    }

    public static void PptxShapeHyperlinkActionWithoutTargetEmitsDiagnostic()
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
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="ActionOnly"><a:hlinkClick action="ppaction://hlinkshowjump"/></p:cNvPr><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input);

        TestAssert.Equal(0, pages.Single().Annotations.Count);
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "An action-only hyperlink should emit the unsupported-action diagnostic.");
    }

    public static void PptxShapeHyperlinkMissingRelationshipEmitsDiagnostic()
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
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Dangling"><a:hlinkClick r:id="rIdMissing"/></p:cNvPr><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input);

        TestAssert.Equal(1, pages.Count);
        TestAssert.Equal(0, pages.Single().Annotations.Count);
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK"), "A hyperlink with an unresolvable relationship should emit the unsupported-hyperlink diagnostic.");
    }

    public static void PptxConnectorHyperlinkEmitsUriAnnotation()
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
                  <Relationship Id="rIdCxnLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/connector" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:cxnSp>
                    <p:nvCxnSpPr><p:cNvPr id="8" name="LinkedConnector"><a:hlinkClick r:id="rIdCxnLink"/></p:cNvPr><p:nvPr/></p:nvCxnSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="line"/>
                      <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                    </p:spPr>
                  </p:cxnSp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input);

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/connector", annotation.Uri);
        TestAssert.Equal(72d, annotation.X);
        TestAssert.Equal(396d, annotation.Y);
        TestAssert.Equal(144d, annotation.Width);
        TestAssert.Equal(72d, annotation.Height);
    }

    public static void PptxPictureHyperlinkEmitsUriAnnotation()
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
                  <Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
                  <Relationship Id="rIdPicLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/picture" TargetMode="External"/>
                </Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:pic>
                    <p:nvPicPr><p:cNvPr id="5" name="LinkedPicture"><a:hlinkClick r:id="rIdPicLink"/></p:cNvPr><p:nvPr/></p:nvPicPr>
                    <p:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
                    <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
                  </p:pic></p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/media/image1.png"] = TestFixtures.CreateRgbPng(1, 1, [255, 0, 0])
        });

        (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input);

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/picture", annotation.Uri);
        TestAssert.Equal(72d, annotation.X);
        TestAssert.Equal(396d, annotation.Y);
        TestAssert.Equal(144d, annotation.Width);
        TestAssert.Equal(72d, annotation.Height);
    }
    public static void PptxNodeHyperlinkMissingRelationshipAggregatesPerSlide()
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
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="2" name="First"><a:hlinkClick r:id="rIdMissing"/></p:cNvPr><p:nvPr/></p:nvSpPr>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                      </p:spPr>
                    </p:sp>
                    <p:sp>
                      <p:nvSpPr><p:cNvPr id="3" name="Second"><a:hlinkClick r:id="rIdMissing"/></p:cNvPr><p:nvPr/></p:nvSpPr>
                      <p:spPr>
                        <a:xfrm><a:off x="914400" y="2743200"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                        <a:prstGeom prst="rect"/>
                        <a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>
                      </p:spPr>
                    </p:sp>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input);

        TestAssert.Equal(0, pages.Single().Annotations.Count);
        TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK"));
    }
    private static (IReadOnlyList<PdfPage> Pages, List<OoxPdfDiagnostic> Diagnostics) RenderSlidePages(string input)
    {
        return RenderSlidePages(input, null);
    }

    private static (IReadOnlyList<PdfPage> Pages, List<OoxPdfDiagnostic> Diagnostics) RenderSlidePages(string input, IFontResolver? fontResolver)
    {
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        var diagnostics = new List<OoxPdfDiagnostic>();
        IReadOnlyList<PdfPage> pages = new PptxRenderer(fontResolver).RenderPages(document, package, diagnostics.Add, CancellationToken.None);
        return (pages, diagnostics);
    }

    private static void AssertNear(double expected, double actual, string message)
    {
        TestAssert.True(Math.Abs(expected - actual) < 1e-6d, message + $" Expected {expected}, got {actual}.");
    }
}
