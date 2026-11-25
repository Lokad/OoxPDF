using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Tests.PptxHyperlinksTests;

namespace Lokad.OoxPdf.Tests;

internal static class PptxFrameHyperlinksTests
{
    public static void PptxTableFrameHyperlinkSitsBelowCellRunLinks()
    {
        string input = WriteLinkedCellTable(linkFrame: true);
        var (pages, diagnostics) = RenderSlidePages(input, new PortableRenderingTests.PinnedResolver());

        PdfPage page = pages.Single();
        TestAssert.Equal(2, page.Annotations.Count);
        var frame = page.Annotations[0];
        TestAssert.Equal("https://example.invalid/table", frame.Uri);
        TestAssert.Equal(72d, frame.X);
        TestAssert.Equal(396d, frame.Y);
        TestAssert.Equal(144d, frame.Width);
        TestAssert.Equal(72d, frame.Height);

        var cell = page.Annotations[1];
        TestAssert.Equal("https://example.invalid/cell", cell.Uri);
        TestAssert.True(cell.X >= frame.X && cell.X + cell.Width <= frame.X + frame.Width &&
            cell.Y >= frame.Y && cell.Y + cell.Height <= frame.Y + frame.Height &&
            cell.Width < frame.Width && cell.Height < frame.Height,
            "The cell link occupies its own area within the frame link.");
        var glyphs = PortableRenderingTests.ReadPaintedSourceLabel(page, "Cell");
        TestAssert.Equal(4, glyphs.Count);
        TestAssert.True(glyphs.All(glyph => glyph.Ink.X >= cell.X - .01d &&
            glyph.Ink.X + glyph.Ink.Width <= cell.X + cell.Width + .01d &&
            glyph.Ink.Y >= cell.Y - .01d && glyph.Ink.Y + glyph.Ink.Height <= cell.Y + cell.Height + .01d),
            "The cell link covers all four painted letters.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "Resolvable frame and cell links should not emit hyperlink diagnostics.");
    }

    internal static string WriteLinkedCellTable(bool linkFrame)
    {
        const string frameRelationship = """<Relationship Id="rIdFrameLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/table" TargetMode="External"/>""";
        const string frameProperties = """<p:nvGraphicFramePr><p:cNvPr id="6" name="LinkedTable"><a:hlinkClick r:id="rIdFrameLink"/></p:cNvPr><p:nvPr/></p:nvGraphicFramePr>""";
        return TestFixtures.WriteTempPackage(".pptx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  {{(linkFrame ? frameRelationship : "")}}
                  <Relationship Id="rIdCellLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/cell" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      {{(linkFrame ? frameProperties : "")}}
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
    }

    public static void PptxChartFrameHyperlinkEmitsUriAnnotation()
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
                  <Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                  <Relationship Id="rIdFrameLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/chart" TargetMode="External"/>
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
                      <p:nvGraphicFramePr><p:cNvPr id="6" name="LinkedChart"><a:hlinkClick r:id="rIdFrameLink"/></p:cNvPr><p:nvPr/></p:nvGraphicFramePr>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="3657600"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """),
            ["ppt/charts/chart1.xml"] = TestFixtures.Utf8("""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart>
                  <c:grouping val="standard"/>
                  <c:ser>
                    <c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                    <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numLit></c:val>
                  </c:ser>
                  <c:axId val="10"/><c:axId val="20"/>
                </c:lineChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="20"/></c:scaling><c:axPos val="l"/><c:tickLblPos val="none"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart></c:chartSpace>
                """)
        });

        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderSlidePages(input, new TestFaceFontResolver());

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/chart", annotation.Uri);
        TestAssert.Equal(72d, annotation.X);
        TestAssert.Equal(180d, annotation.Y);
        TestAssert.Equal(432d, annotation.Width);
        TestAssert.Equal(288d, annotation.Height);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable chart-frame link should not emit hyperlink diagnostics.");
    }

    public static void PptxUnknownFrameHyperlinkEmitsUriAnnotation()
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
                  <Relationship Id="rIdFrameLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/frame" TargetMode="External"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:nvGraphicFramePr><p:cNvPr id="6" name="LinkedUnknown"><a:hlinkClick r:id="rIdFrameLink"/></p:cNvPr><p:nvPr/></p:nvGraphicFramePr>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://example.invalid/unknown-shape" /></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """
        });

        (IReadOnlyList<PdfPage> pages, _) = RenderSlidePages(input, new TestFaceFontResolver());

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/frame", annotation.Uri);
        TestAssert.Equal(72d, annotation.X);
        TestAssert.Equal(396d, annotation.Y);
        TestAssert.Equal(144d, annotation.Width);
        TestAssert.Equal(72d, annotation.Height);
    }
}
