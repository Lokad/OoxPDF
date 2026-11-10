using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartTitleDefaultsTests
{
    public static void PptxRichParagraphDefaultsRespectTitleAndPointLabelScopes()
    {
        const string rich = """
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:pPr><a:defRPr sz="1800" b="1" i="1" spc="50">
              <a:latin typeface="DejaVu Sans"/><a:solidFill><a:srgbClr val="336699"/></a:solidFill>
            </a:defRPr></a:pPr>
              <a:r><a:t>Inherited</a:t></a:r>
              <a:r><a:rPr sz="1400" b="0" i="0" spc="0">
                <a:solidFill><a:srgbClr val="992233"/></a:solidFill><a:hlinkClick r:id="link"/>
              </a:rPr><a:t>Override</a:t></a:r>
            </a:p>
            <a:p><a:pPr><a:defRPr sz="1000" b="0"/></a:pPr><a:r><a:t>Next paragraph</a:t></a:r></a:p>
            """;
        PptxSceneChart chart = TestAssert.NotNull(PptxTests.BuildSingleChartScene(Chart(rich)));
        var runs = chart.Title.TextRuns;
        TestAssert.Equal(3, runs.Count);
        var inherited = runs[0].TextStyle;
        TestAssert.Equal(18d, inherited.FontSize ?? 0d);
        TestAssert.Equal("DejaVu Sans", inherited.FontFamily);
        TestAssert.True(inherited.Bold == true && inherited.Italic == true, "Paragraph defaults reach an unformatted run.");
        TestAssert.Equal(.5d, inherited.CharacterSpacing ?? 0d);
        TestAssert.Equal(new RgbColor(0x33, 0x66, 0x99), inherited.Color!.Value);
        var local = runs[1].TextStyle;
        TestAssert.Equal(14d, local.FontSize ?? 0d);
        TestAssert.Equal("DejaVu Sans", local.FontFamily);
        TestAssert.True(local.Bold == false && local.Italic == false, "Explicit false values override inherited true values.");
        TestAssert.Equal(0d, local.CharacterSpacing ?? -1d);
        TestAssert.Equal(new RgbColor(0x99, 0x22, 0x33), local.Color!.Value);
        TestAssert.Equal("link", runs[1].HyperlinkClickId);
        TestAssert.Equal(10d, runs[2].TextStyle.FontSize ?? 0d);
        TestAssert.True(runs[2].TextStyle.Bold == false && runs[2].TextStyle.Italic is null,
            "A new paragraph does not inherit defaults from its previous sibling.");
        var labels = chart.Plots.Single().DataLabels.Overrides.Single().CustomTextRuns;
        TestAssert.True(labels[0].TextStyle.FontSize is null && labels[0].TextStyle.Bold is null,
            "Native custom labels keep their point/container defaults instead of adopting rich-paragraph title defaults.");
        TestAssert.Equal(14d, labels[1].TextStyle.FontSize ?? 0d);
        TestAssert.True(labels[1].TextStyle.Bold == false && labels[1].TextStyle.Italic == false,
            "Direct label-run declarations retain their independent precedence.");
        TestAssert.Equal("link", labels[1].HyperlinkClickId);
    }

    public static void PptxRichParagraphFontSurvivesRunSplitsThroughPainting()
    {
        foreach (string[] parts in new[] { new[] { "Paragraph style" }, new[] { "Paragraph ", "style" }, new[] { "Para", "graph ", "style" } })
        {
            string rich = "<a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"1800\"/></a:pPr>" +
                string.Concat(parts.Select(part => "<a:r><a:t>" + part + "</a:t></a:r>")) + "</a:p>";
            using var bytes = PptxTests.CreateSingleChartPackageStream(Chart(rich));
            using var package = OoxPackage.Open(bytes, default);
            var document = new PptxReader().Read(package, default);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(document, package, null, default)
                .ToBlockingEnumerable().Single();
            var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CT", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal("Paragraph style", string.Concat(shows.Select(show => show.Text)));
            TestAssert.True(shows.All(show => Math.Abs(show.Size - 18d) < .001d),
                "Every painted title fragment retains its explicit paragraph font size after a run split.");
        }
    }

    private static string Chart(string rich) => $$"""
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <c:chart><c:title><c:tx><c:rich>{{rich}}</c:rich></c:tx></c:title><c:plotArea><c:pieChart><c:ser>
            <c:idx val="0"/><c:order val="0"/>
            <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
            <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
          </c:ser><c:dLbls><c:dLbl><c:idx val="0"/><c:tx><c:rich>{{rich}}</c:rich></c:tx></c:dLbl></c:dLbls>
          </c:pieChart></c:plotArea></c:chart>
        </c:chartSpace>
        """;
}
