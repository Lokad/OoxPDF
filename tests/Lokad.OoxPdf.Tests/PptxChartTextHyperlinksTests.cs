using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxChartTextHyperlinksTests
{
    public static void PptxChartTitleExternalHyperlinkEmitsUriAnnotation()
    {
        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderChartTitlePage(
            """<a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Visit site</a:t></a:r>""",
            """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/chart-title" TargetMode="External"/>
                </Relationships>
                """);

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/chart-title", annotation.Uri);
        TestAssert.True(annotation.Width > 0d && annotation.Height > 0d, "Chart title links need a non-degenerate box.");
        TestAssert.True(annotation.Width < 288d, "Chart title links must cover the run span, not the whole frame.");
        TestAssert.True(annotation.X >= 72d && annotation.X + annotation.Width <= 360d, "Chart title links must sit inside the chart frame horizontally.");
        TestAssert.True(annotation.Y > 360d && annotation.Y + annotation.Height <= 540d, "Chart title links must sit in the title zone above the plot area.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable chart title hyperlink should not emit hyperlink diagnostics.");
    }

    public static void PptxChartTitleHyperlinkActionIsDiagnosed()
    {
        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderChartTitlePage(
            """<a:r><a:rPr><a:hlinkClick action="ppaction://hlinksldjump()"/></a:rPr><a:t>Jump</a:t></a:r>""",
            null);

        TestAssert.Equal(0, pages.Single().Annotations.Count);
        TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK_ACTION"));
    }

    public static void PptxChartTitleMissingRelationshipIsDiagnosed()
    {
        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderChartTitlePage(
            """<a:r><a:rPr><a:hlinkClick r:id="rIdMissing"/></a:rPr><a:t>Broken</a:t></a:r>""",
            null);

        TestAssert.Equal(0, pages.Single().Annotations.Count);
        TestAssert.Equal(1, diagnostics.Count(d => d.Id == "PPTX_UNSUPPORTED_HYPERLINK"));
    }

    public static void PptxChartTitleWithoutHyperlinkStaysQuiet()
    {
        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderChartTitlePage(
            """<a:r><a:rPr sz="1800"><a:latin typeface="Arial"/></a:rPr><a:t>Plain title</a:t></a:r>""",
            null);

        TestAssert.Equal(0, pages.Single().Annotations.Count);
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A plain chart title should not emit hyperlink diagnostics.");
    }

    public static void PptxChartAxisTitleExternalHyperlinkEmitsUriAnnotation()
    {
        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderChartTitlePage(
            "",
            """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/chart-axis" TargetMode="External"/>
                </Relationships>
                """,
            """
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:chart>
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
                      <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:crossAx val="10"/><c:majorUnit val="5"/><c:tickLblPos val="none"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Y axis</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title></c:valAx>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """);

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/chart-axis", annotation.Uri);
        TestAssert.True(annotation.Width > 0d && annotation.Height > 0d, "Chart axis title links need a non-degenerate box.");
        TestAssert.True(annotation.X >= 60d && annotation.X + annotation.Width <= 216d, "Chart axis title links must sit in the left title zone.");
        TestAssert.True(annotation.Y >= 252d && annotation.Y + annotation.Height <= 468d, "Chart axis title links must sit inside the chart frame vertically.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable chart axis title hyperlink should not emit hyperlink diagnostics.");
    }

    public static void PptxChartDataLabelCustomTextHyperlinkEmitsUriAnnotation()
    {
        (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderChartTitlePage(
            "",
            """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/chart-label" TargetMode="External"/>
                </Relationships>
                """,
            """
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:chart>
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
                        <c:dLbls><c:dLbl><c:idx val="0"/><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Point link</a:t></a:r></a:p></c:rich></c:tx></c:dLbl></c:dLbls>
                        <c:axId val="10"/><c:axId val="20"/>
                      </c:barChart>
                      <c:catAx><c:axId val="10"/><c:axPos val="b"/><c:crossAx val="20"/><c:tickLblPos val="none"/></c:catAx>
                      <c:valAx><c:axId val="20"/><c:axPos val="l"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:crossAx val="10"/><c:majorUnit val="5"/><c:tickLblPos val="none"/></c:valAx>
                    </c:plotArea>
                  </c:chart>
                </c:chartSpace>
                """);

        PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
        TestAssert.Equal("https://example.invalid/chart-label", annotation.Uri);
        TestAssert.True(annotation.Width > 0d && annotation.Height > 0d, "Chart data label links need a non-degenerate box.");
        TestAssert.True(annotation.Width < 288d, "Chart data label links must cover the run span, not the whole frame.");
        TestAssert.True(annotation.X >= 72d && annotation.X + annotation.Width <= 360d, "Chart data label links must sit inside the chart frame horizontally.");
        TestAssert.True(annotation.Y >= 252d && annotation.Y + annotation.Height <= 468d, "Chart data label links must sit inside the chart frame vertically.");
        TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable chart data label hyperlink should not emit hyperlink diagnostics.");
    }

    private static (IReadOnlyList<PdfPage> Pages, List<OoxPdfDiagnostic> Diagnostics) RenderChartTitlePage(string titleRunXml, string? chartRelsXml, string? chartXmlOverride = null)
    {
        string contentTypes = PptxTests.BasicContentTypes().Replace(
            "</Types>",
            """
              <Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
            </Types>
            """);
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = contentTypes,
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/_rels/slide1.xml.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                </Relationships>
                """,
            ["ppt/slides/slide1.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <p:cSld><p:spTree>
                    <p:graphicFrame>
                      <p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:nvPr/></p:nvGraphicFramePr>
                      <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2743200"/></p:xfrm>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>
                    </p:graphicFrame>
                  </p:spTree></p:cSld>
                </p:sld>
                """,
            ["ppt/charts/chart1.xml"] = chartXmlOverride ?? $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:chart>
                    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>{{titleRunXml}}</a:p></c:rich></c:tx></c:title>
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
        };
        if (chartRelsXml is not null)
        {
            parts["ppt/charts/_rels/chart1.xml.rels"] = chartRelsXml;
        }

        string input = TestFixtures.WriteTempPackage(".pptx", parts);
        using FileStream stream = File.OpenRead(input);
        OoxPackage package = OoxPackage.Open(stream, CancellationToken.None);
        PptxDocument document = new PptxReader().Read(package, CancellationToken.None);
        var diagnostics = new List<OoxPdfDiagnostic>();
        IReadOnlyList<PdfPage> pages = new PptxRenderer(null).RenderPages(document, package, diagnostics.Add, CancellationToken.None);
        return (pages, diagnostics);
    }
}
