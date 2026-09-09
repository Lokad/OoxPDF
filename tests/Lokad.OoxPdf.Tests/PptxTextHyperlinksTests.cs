using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxTextHyperlinksTests
{
    public static void PptxRunExternalHyperlinkEmitsUriAnnotation()
    {
        string input = WriteRunLinkPackage(
            paragraphXml: """<a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdRunLink"/></a:rPr><a:t>Link</a:t></a:r></a:p>""",
            slideRelsXml: """<Relationship Id="rIdRunLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/run" TargetMode="External"/>""");

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, new LinkTestFontResolver());

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/run", annotation.Uri);
        TestAssert.True(annotation.Width > 0d, "Run hyperlink annotations should cover the linked text width.");
        TestAssert.True(Math.Abs(19.8d - annotation.Height) < 1e-6d, $"Run hyperlink height should follow the embedded font ascent/descent box, got {annotation.Height}.");
        TestAssert.True(annotation.X >= 72d && annotation.X + annotation.Width <= 72d + 432d, "Run hyperlink X extent should stay inside the shape bounds.");
        TestAssert.True(annotation.Y >= 396d && annotation.Y + annotation.Height <= 396d + 72d, "Run hyperlink Y extent should stay inside the shape bounds.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable run hyperlink should not emit hyperlink diagnostics.");
    }

    public static void PptxRunHyperlinkBoundaryKeepsAdjacentLinksSeparate()
    {
        string input = WriteRunLinkPackage(
            paragraphXml: """<a:p>""" +
                """<a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdLink1"/></a:rPr><a:t>A </a:t></a:r>""" +
                """<a:r><a:rPr sz="1800" u="sng"><a:latin typeface="Link Sans"/></a:rPr><a:t>B</a:t></a:r>""" +
                """<a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdLink2"/></a:rPr><a:t>C</a:t></a:r></a:p>""",
            slideRelsXml: """<Relationship Id="rIdLink1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/first" TargetMode="External"/>""" +
                """<Relationship Id="rIdLink2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/second" TargetMode="External"/>""");

        (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input, new LinkTestFontResolver());

        PdfLinkAnnotation[] annotations = pages.Single().Annotations.ToArray();
        TestAssert.Equal(2, annotations.Length);
        TestAssert.Equal("https://example.invalid/first", annotations[0].Uri);
        TestAssert.Equal("https://example.invalid/second", annotations[1].Uri);
        TestAssert.True(annotations[1].X > annotations[0].X, "Adjacent link annotations should follow text order.");
    }

    public static void PptxRunInternalSlideHyperlinkEmitsDestination()
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
                  <Relationship Id="rIdRunJump" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide2.xml"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="RunJump"/><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>
                    </p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdRunJump"/></a:rPr><a:t>Jump</a:t></a:r></a:p>
                    </p:txBody>
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

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, new LinkTestFontResolver());

        TestAssert.Equal(2, pages.Count);
        PdfLinkAnnotation annotation = pages[0].Annotations.Single();
        TestAssert.True(annotation.Uri is null, "An internal run hyperlink should not carry a URI target.");
        TestAssert.True(annotation.Destination is { PageIndex: 1 }, "An internal run hyperlink should target the second page.");
        TestAssert.Equal(0, pages[1].Annotations.Count);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable internal run hyperlink should not emit hyperlink diagnostics.");
    }

    public static void PptxRunHyperlinkMissingRelationshipEmitsSingleDiagnostic()
    {
        string input = WriteRunLinkPackage(
            paragraphXml: """<a:p>""" +
                """<a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdMissing"/></a:rPr><a:t>A </a:t></a:r>""" +
                """<a:r><a:rPr sz="1800" u="sng"><a:latin typeface="Link Sans"/></a:rPr><a:t>B</a:t></a:r>""" +
                """<a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdMissing"/></a:rPr><a:t>C</a:t></a:r></a:p>""",
            slideRelsXml: string.Empty);

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, new LinkTestFontResolver());

        TestAssert.Equal(1, pages.Count);
        TestAssert.Equal(0, pages.Single().Annotations.Count);
        TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK"));
    }

    public static void PptxTableCellRunHyperlinkEmitsUriAnnotation()
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
                  <Relationship Id="rIdCellLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/cell" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
                        <a:tblGrid><a:gridCol w="1828800"/></a:tblGrid>
                        <a:tr h="914400"><a:tc>
                          <a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdCellLink"/></a:rPr><a:t>Cell</a:t></a:r></a:p></a:txBody>
                          <a:tcPr/>
                        </a:tc></a:tr>
                      </a:tbl></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, new LinkTestFontResolver());

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/cell", annotation.Uri);
        TestAssert.True(annotation.Width > 0d, "Table-cell hyperlink annotations should cover the linked text width.");
        TestAssert.True(Math.Abs(19.8d - annotation.Height) < 1e-6d, $"Table-cell hyperlink height should follow the embedded font ascent/descent box, got {annotation.Height}.");
        TestAssert.True(annotation.X >= 72d && annotation.X + annotation.Width <= 72d + 144d, "Table-cell hyperlink X extent should stay inside the table frame.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable table-cell hyperlink should not emit hyperlink diagnostics.");
    }

    public static void PptxRunHyperlinkActionWithoutTargetEmitsDiagnostic()
    {
        string input = WriteRunLinkPackage(
            paragraphXml: """<a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick action="ppaction://hlinkshowjump"/></a:rPr><a:t>Action</a:t></a:r></a:p>""",
            slideRelsXml: string.Empty);

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, new LinkTestFontResolver());

        TestAssert.Equal(0, pages.Single().Annotations.Count);
        TestAssert.True(diagnostics.Any(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "An action-only run hyperlink should emit the unsupported-action diagnostic.");
    }

    public static void PptxRotatedRunHyperlinkUsesTransformedBox()
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
                  <Relationship Id="rIdRunLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/rotated" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="RotatedLink"/><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm rot="5400000"><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>
                    </p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/>
                      <a:p><a:r><a:rPr sz="1800"><a:latin typeface="Link Sans"/><a:hlinkClick r:id="rIdRunLink"/></a:rPr><a:t>Link</a:t></a:r></a:p>
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input, new LinkTestFontResolver());

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/rotated", annotation.Uri);
        TestAssert.True(annotation.Height > annotation.Width, "A 90-degree rotated run link should cover a tall box, not the unrotated wide box.");
        TestAssert.True(annotation.X >= 0d && annotation.X + annotation.Width <= 720d, "Rotated run links should stay on the slide.");
        TestAssert.True(annotation.Y >= 0d && annotation.Y + annotation.Height <= 540d, "Rotated run links should stay on the slide.");
    }
    private static string WriteRunLinkPackage(string paragraphXml, string slideRelsXml)
    {
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  {slideRelsXml}
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = $"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree><p:sp>
                    <p:nvSpPr><p:cNvPr id="2" name="RunLink"/><p:nvPr/></p:nvSpPr>
                    <p:spPr>
                      <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
                      <a:prstGeom prst="rect"/>
                      <a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill>
                    </p:spPr>
                    <p:txBody><a:bodyPr/><a:lstStyle/>
                      {paragraphXml}
                    </p:txBody>
                  </p:sp></p:spTree></p:cSld>
                </p:sld>
                """
        });
    }

    private sealed class LinkTestFontResolver : IFontResolver
    {
        private readonly byte[] fontBytes = TestFontBuilder.CreateTestFont();

        public FontFaceResolution Resolve(FontRequest request)
        {
            return new FontFaceResolution(
                request.FamilyName,
                "LinkTestFont",
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                new MemoryFontProgramSource("link-test-font", fontBytes),
                IsFallback: false);
        }
    }

    private static (IReadOnlyList<PdfPage> Pages, List<OoxPdfDiagnostic> Diagnostics) RenderSlidePages(string input, IFontResolver fontResolver)
    {
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        var diagnostics = new List<OoxPdfDiagnostic>();
        IReadOnlyList<PdfPage> pages = new PptxRenderer(fontResolver).RenderPages(document, package, diagnostics.Add, CancellationToken.None);
        return (pages, diagnostics);
    }
}
