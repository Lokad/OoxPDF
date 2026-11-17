using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingChartCachedFieldsKeepWordsAndHyperlinksAcrossTextOwners()
    {
        var resolver = new PinnedResolver();
        foreach (string owner in new[] { "title", "axis", "label" })
        {
            string? control = null;
            foreach (bool field in new[] { false, true })
            {
                string rich = "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>" +
                    PptxChartTextHyperlinksTests.CachedFieldInline(field, mixed: true) + "</a:p></c:rich></c:tx>";
                string title = "<c:title>" + rich + "<c:overlay val=\"0\"/></c:title>";
                string chart = $$"""
                    <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>
                      {{(owner == "title" ? title : "")}}
                      <c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                        <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series</c:v></c:tx><c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                        {{(owner == "label" ? "<c:dLbls><c:dLbl><c:idx val=\"0\"/>" + rich + "</c:dLbl></c:dLbls>" : "")}}
                        <c:axId val="10"/><c:axId val="20"/></c:barChart>
                        <c:catAx><c:axId val="10"/><c:scaling/><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="20"/></c:catAx>
                        <c:valAx><c:axId val="20"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="l"/>
                          {{(owner == "axis" ? title : "")}}
                          <c:tickLblPos val="none"/><c:crossAx val="10"/><c:majorUnit val="5"/></c:valAx>
                      </c:plotArea></c:chart></c:chartSpace>
                    """;
                var (pages, diagnostics) = PptxChartTextHyperlinksTests.RenderChartTitlePage("", """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdField" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/cached-field" TargetMode="External"/></Relationships>
                    """, chart, resolver);
                PdfPage page = pages.Single();
                foreach (string word in new[] { "START", "CACHED", "END" })
                    TestAssert.Equal(word.Length, ReadPaintedSourceLabel(page, word).Count);
                TestAssert.Equal("https://example.invalid/cached-field", page.Annotations.Single().Uri!);
                TestAssert.True(diagnostics.All(item => !item.Id.Contains("HYPERLINK", StringComparison.Ordinal)),
                    "A cached field keeps its resolvable link owner.");
                control ??= page.Content;
                TestAssert.Equal(control, page.Content);
            }
        }
    }
}
