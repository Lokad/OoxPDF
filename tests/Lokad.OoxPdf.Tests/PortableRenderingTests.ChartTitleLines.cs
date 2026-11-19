using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PortableRenderingTests
{
    public static void PortableRenderingRichChartTitleKeepsAuthoredLinesAndTheirLinks()
    {
        const string first = "<a:r><a:rPr sz=\"1800\"/><a:t>North</a:t></a:r>";
        const string second = """
            <a:fld id="{77777777-7777-7777-7777-777777777777}" type="datetime">
              <a:rPr sz="1600"><a:hlinkClick r:id="sales"/></a:rPr><a:t>Sales</a:t>
            </a:fld>
            """;
        const string manual = """
            <c:layout><c:manualLayout><c:x val="0.36"/><c:y val="0.12"/>
              <c:wMode val="factor"/><c:hMode val="factor"/><c:w val="0.25"/><c:h val="0.3"/>
            </c:manualLayout></c:layout>
            """;
        var resolver = new PinnedResolver();

        var adjacent = Check(Render(first + second, ""));
        TestAssert.True(Math.Abs(adjacent.First - adjacent.Second) < .002, "Adjacent styled runs share a baseline.");
        var oneBreak = CheckEquivalentBreaks("");
        CheckEquivalentBreaks(manual);
        var twoBreaks = Check(Render(first + "<a:br/><a:br/>" + second, ""));
        TestAssert.True(Math.Abs(twoBreaks.First - oneBreak.First) < .002 && twoBreaks.Second < oneBreak.Second - 10d,
            "A blank line advances subsequent content without shifting the first line.");

        (double First, double Second) CheckEquivalentBreaks(string layout)
        {
            PdfPage inlineBreak = Render(first + "<a:br/>" + second, layout);
            var baselines = Check(inlineBreak);
            TestAssert.True(baselines.First > baselines.Second + 10d, "Each authored line paints once on a separate baseline.");
            PdfPage paragraphBreak = Render(first + "</a:p><a:p>" + second, layout);
            TestAssert.Equal(inlineBreak.Content, paragraphBreak.Content);
            TestAssert.Equal(inlineBreak.Annotations.Single(), paragraphBreak.Annotations.Single());
            return baselines;
        }

        PdfPage Render(string titleRuns, string layout)
        {
            string chart = $$"""
                <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>{{titleRuns}}</a:p></c:rich></c:tx>
                    {{layout}}</c:title><c:plotArea><c:pieChart><c:ser><c:idx val="0"/><c:order val="0"/>
                      <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
                    </c:ser></c:pieChart></c:plotArea></c:chart>
                </c:chartSpace>
                """;
            const string relationships = """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="sales" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                    Target="https://example.invalid/sales" TargetMode="External"/>
                </Relationships>
                """;
            return PptxChartTextHyperlinksTests.RenderChartTitlePage("", relationships, chart, resolver).Pages.Single();
        }

        static (double First, double Second) Check(PdfPage page)
        {
            var north = ReadPaintedSourceLabel(page, "North", fontPrefix: "CT");
            var sales = ReadPaintedSourceLabel(page, "Sales", fontPrefix: "CT");
            TestAssert.True(north.All(glyph => Math.Abs(glyph.Size - 18d) < .001) &&
                sales.All(glyph => Math.Abs(glyph.Size - 16d) < .001), "Both lines preserve their own run style.");
            PdfLinkAnnotation link = page.Annotations.Single();
            TestAssert.Equal("https://example.invalid/sales", link.Uri);
            TestAssert.True(link.Y <= sales[0].Origin.Y && link.Y + link.Height > sales[0].Origin.Y,
                "The cached field link follows the second line's baseline.");
            return (north[0].Origin.Y, sales[0].Origin.Y);
        }
    }
}
