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

        (List<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input);

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
        foreach (string target in new[] { "slide2.xml", "SLIDE2.XML", "./slide2.xml", "./SLIDE2.XML", "/ppt/slides/SLIDE2.XML", "missing.xml" })
        {
            string input = WriteInternalSlideLinkPackage(ShapeWithInternalLink, target);
            (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input);

            TestAssert.Equal(2, pages.Count);
            TestAssert.Equal(0, pages[1].Annotations.Count);
            if (target == "missing.xml")
            {
                TestAssert.Equal(0, pages[0].Annotations.Count);
                TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK"));
                continue;
            }
            PdfLinkAnnotation annotation = pages[0].Annotations.Single();
            TestAssert.True(annotation.Uri is null, "An internal slide hyperlink should not carry a URI target.");
            TestAssert.True(annotation.Destination is { PageIndex: 1 }, "Equivalent package targets should select the second page.");
            TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable internal slide hyperlink should not emit hyperlink diagnostics.");
        }
    }

    internal const string ShapeWithInternalLink = """
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Jump"><a:hlinkClick r:id="rIdLink"/></p:cNvPr><p:nvPr/></p:nvSpPr>
          <p:spPr>
            <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
            <a:prstGeom prst="rect"/>
            <a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>
          </p:spPr>
        </p:sp>
        """;

    internal static string WriteInternalSlideLinkPackage(string sourceShape, string target = "slide2.xml", bool hideTarget = false)
    {
        string targetVisibility = hideTarget ? " show=\"0\"" : "";
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes().Replace("</Types>", """
                <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
                </Types>
                """, StringComparison.Ordinal),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship().Replace("</Relationships>", """
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
                </Relationships>
                """, StringComparison.Ordinal),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation().Replace("</p:sldIdLst>",
                """<p:sldId id="257" r:id="rId2"/></p:sldIdLst>""", StringComparison.Ordinal),
            ["ppt/slides/_rels/slide1.xml.rels"] = $"""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="{target}"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = $"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>{sourceShape}</p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/slides/slide2.xml"] = $"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"{targetVisibility}>
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="Target"/><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                    </p:spPr>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
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
            ["ppt/slides/slide1.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    {{Connector(8, 72, 72, 144, 72)}}
                    {{Connector(9, 72, 160, 144, 0)}}
                    {{Connector(10, 290, 108, 0, 108)}}
                    {{Connector(11, 400, 180, 80, 0, rotation: 90)}}
                    {{Connector(12, 0, 250, 0, 90)}}
                    {{Connector(13, 60, 90, 60, 0, cap: "rnd")}}
                    {{Connector(14, 72, 240, 144, 0, visible: false)}}
                    {{Connector(15, -100, 310, 50, 0)}}
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input);

        IReadOnlyList<PdfLinkAnnotation> annotations = pages.Single().Annotations;
        PdfRectangle[] expected =
        [
            new(72d, 396d, 144d, 72d), new(72d, 379d, 144d, 2d), new(289d, 324d, 2d, 108d),
            new(439d, 320d, 2d, 80d), new(0d, 200d, 1d, 90d), new(59d, 449d, 62d, 2d)
        ];
        TestAssert.Equal(expected.Length, annotations.Count);
        for (int index = 0; index < expected.Length; index++)
        {
            var link = annotations[index];
            var bounds = expected[index];
            TestAssert.Equal("https://example.invalid/connector", link.Uri);
            TestAssert.True(Math.Abs(bounds.X - link.X) < .001d && Math.Abs(bounds.Y - link.Y) < .001d &&
                Math.Abs(bounds.Width - link.Width) < .001d && Math.Abs(bounds.Height - link.Height) < .001d,
                $"Connector {index + 8} has the expected clickable rectangle.");
        }

        static string Connector(int id, int x, int y, int width, int height,
            int rotation = 0, string cap = "flat", bool visible = true)
        {
            string fill = visible ? """<a:solidFill><a:srgbClr val="000000"/></a:solidFill>""" : "<a:noFill/>";
            return $$"""
                <p:cxnSp>
                  <p:nvCxnSpPr><p:cNvPr id="{{id}}" name="LinkedConnector{{id}}"><a:hlinkClick r:id="rIdCxnLink"/></p:cNvPr><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
                  <p:spPr>
                    <a:xfrm rot="{{rotation * 60000}}"><a:off x="{{x * 12700}}" y="{{y * 12700}}"/><a:ext cx="{{width * 12700}}" cy="{{height * 12700}}"/></a:xfrm>
                    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
                    <a:ln w="25400" cap="{{cap}}">{{fill}}</a:ln>
                  </p:spPr>
                </p:cxnSp>
                """;
        }
    }

    public static void PptxPictureHyperlinkEmitsUriAnnotation()
    {
        var parts = new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Default Extension="png" ContentType="image/png"/>
                  <Default Extension="svg" ContentType="image/svg+xml"/>
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
        };

        (PdfPage page, _) = Render();
        TestAssert.Equal(1, page.Images.Count);
        PdfLinkAnnotation annotation = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/picture", annotation.Uri);
        TestAssert.Equal(72d, annotation.X);
        TestAssert.Equal(396d, annotation.Y);
        TestAssert.Equal(144d, annotation.Width);
        TestAssert.Equal(72d, annotation.Height);

        const string slidePart = "ppt/slides/slide1.xml";
        string slide = Encoding.UTF8.GetString(parts[slidePart]);
        foreach (int left in new[] { 700, 740 })
        {
            var moved = new Dictionary<string, byte[]>(parts)
            {
                [slidePart] = TestFixtures.Utf8(slide.Replace("x=\"914400\"", $"x=\"{left * 12700}\"", StringComparison.Ordinal))
            };
            (PdfPage clipped, _) = Render(moved);
            if (left < page.Width)
            {
                TestAssert.Equal(1, clipped.Images.Count);
                PdfLinkAnnotation visible = clipped.Annotations.Single();
                TestAssert.Equal(annotation.Uri, visible.Uri);
                TestAssert.Equal((double)left, visible.X);
                TestAssert.Equal(page.Width - left, visible.Width);
            }
            else
            {
                TestAssert.Equal(0, clipped.Images.Count);
                TestAssert.Equal(0, clipped.Annotations.Count);
            }
        }

        var missing = new Dictionary<string, byte[]>(parts);
        missing.Remove("ppt/media/image1.png");
        (PdfPage omitted, List<OoxPdfDiagnostic> missingDiagnostics) = Render(missing);
        TestAssert.Equal("IMAGE_MISSING_PART", missingDiagnostics.Single().Id);
        TestAssert.Equal(0, omitted.Images.Count);
        TestAssert.Equal(0, omitted.Annotations.Count);

        const string relationships = "ppt/slides/_rels/slide1.xml.rels";
        var vector = new Dictionary<string, byte[]>(missing)
        {
            [relationships] = TestFixtures.Utf8(Encoding.UTF8.GetString(parts[relationships])
                .Replace("image1.png", "picture.svg", StringComparison.Ordinal)),
            ["ppt/media/picture.svg"] = TestFixtures.Utf8("""<svg xmlns="http://www.w3.org/2000/svg"/>""")
        };
        (PdfPage invalid, List<OoxPdfDiagnostic> vectorDiagnostics) = Render(vector);
        TestAssert.Equal("SVG_UNSUPPORTED_CONTENT", vectorDiagnostics.Single().Id);
        TestAssert.Equal(0, invalid.Annotations.Count);

        vector["ppt/media/picture.svg"] = TestFixtures.Utf8("""
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">
              <rect width="10" height="10" fill="#ff0000"/>
            </svg>
            """);
        (PdfPage painted, List<OoxPdfDiagnostic> paintedDiagnostics) = Render(vector);
        TestAssert.Equal(0, paintedDiagnostics.Count);
        TestAssert.Contains("1 0 0 rg", painted.Content);
        TestAssert.Equal(annotation, painted.Annotations.Single());

        (PdfPage Page, List<OoxPdfDiagnostic> Diagnostics) Render(Dictionary<string, byte[]>? edited = null)
        {
            string input = TestFixtures.WriteTempPackage(".pptx", edited ?? parts);
            var (pages, diagnostics) = RenderSlidePages(input);
            return (pages.Single(), diagnostics);
        }
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
    internal static (List<PdfPage> Pages, List<OoxPdfDiagnostic> Diagnostics) RenderSlidePages(string input, IFontResolver? fontResolver = null)
    {
        using FileStream stream = File.OpenRead(input);
        using var package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        var diagnostics = new List<OoxPdfDiagnostic>();
        List<PdfPage> pages = new PptxRenderer(fontResolver).RenderPagesAsync(document, package, diagnostics.Add, CancellationToken.None).ToBlockingEnumerable().ToList();
        return (pages, diagnostics);
    }

    private static void AssertNear(double expected, double actual, string message)
    {
        TestAssert.True(Math.Abs(expected - actual) < 1e-6d, message + $" Expected {expected}, got {actual}.");
    }
}
