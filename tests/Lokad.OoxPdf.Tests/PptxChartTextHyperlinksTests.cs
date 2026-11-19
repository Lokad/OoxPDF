using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartTextHyperlinksTests
{
    public static void PptxAutomaticChartTitlePlacementKeepsWholeLinkedLine()
    {
        var font = TestFontBuilder.LoadTestFont();
        foreach (IFontResolver resolver in new IFontResolver[] { new TestFaceFontResolver(), new MissingTestFontResolver() })
        {
            foreach (int size in new[] { 18, 48, 72 })
            {
                (List<PdfPage> pages, _) = RenderChartTitlePage(
                    $"""<a:r><a:rPr sz="{size * 100}"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Title</a:t></a:r>""",
                    """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/title" TargetMode="External"/>
                    </Relationships>
                    """, fontResolver: resolver);
                PdfLinkAnnotation link = pages.Single().Annotations.Single();
                double expectedHeight = resolver is MissingTestFontResolver
                    ? size * (PdfFallbackFont.AscentEm + PdfFallbackFont.DescentEm)
                    : size * (font.Os2.WindowsAscender + font.Os2.WindowsDescender) / (double)font.UnitsPerEm;
                TestAssert.True(Math.Abs(link.Height - expectedHeight) < 0.01d,
                    $"The complete {size}pt title line must survive the frame clip ({link.Height}pt).");
                TestAssert.True(link.Y + link.Height <= 468.001d, "The title must stay within its source frame.");
            }
        }
    }

    public static void PptxAuthoredChartTitleSurvivesAutomaticTitleDeletion()
    {
        foreach (IFontResolver? resolver in new IFontResolver?[] { null, new MissingTestFontResolver() })
        {
            string? baseline = null;
            foreach (bool? autoTitleDeleted in new bool?[] { null, false, true })
            {
                (List<PdfPage> pages, _) = RenderChartTitlePage(
                    """<a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Authored heading</a:t></a:r>""",
                    """
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/authored-heading" TargetMode="External"/>
                    </Relationships>
                    """, fontResolver: resolver, autoTitleDeleted: autoTitleDeleted);
                PdfPage page = pages.Single();
                TestAssert.Equal("https://example.invalid/authored-heading", page.Annotations.Single().Uri);
                string content = page.Content;
                baseline ??= content;
                TestAssert.Equal(baseline, content);
            }
        }
    }

    public static void PptxChartTitleExternalHyperlinkEmitsUriAnnotation()
    {
        foreach (IFontResolver? resolver in new IFontResolver?[] { null, new MissingTestFontResolver() })
        {
            (List<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderChartTitlePage(
                """<a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Visit site</a:t></a:r>""",
                """
                    <?xml version="1.0" encoding="UTF-8"?>
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                      <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/chart-title" TargetMode="External"/>
                    </Relationships>
                    """, fontResolver: resolver);

            PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
            TestAssert.Equal("https://example.invalid/chart-title", annotation.Uri);
            TestAssert.True(annotation.Width > 0d && annotation.Height > 0d, "Chart title links need a non-degenerate box.");
            TestAssert.True(annotation.Width < 288d, "Chart title links must cover the run span, not the whole frame.");
            TestAssert.True(annotation.X >= 72d && annotation.X + annotation.Width <= 360d, "Chart title links must sit inside the chart frame horizontally.");
            TestAssert.True(annotation.Y > 360d && annotation.Y + annotation.Height <= 540d, "Chart title links must sit in the title zone above the plot area.");
            TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable chart title hyperlink should not emit hyperlink diagnostics.");
        }
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
        foreach (IFontResolver? resolver in new IFontResolver?[] { null, new MissingTestFontResolver() })
        {
            (IReadOnlyList<PdfPage> pages, List<OoxPdfDiagnostic> diagnostics) = RenderChartTitlePage(
                "", AxisTitleRelationships, AxisTitleChart("l", """
                    <a:r><a:rPr><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Y axis</a:t></a:r>
                    """), fontResolver: resolver);

            PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
            TestAssert.Equal("https://example.invalid/chart-axis", annotation.Uri);
            TestAssert.True(annotation.Width > 0d && annotation.Height > 0d, "Chart axis title links need a non-degenerate box.");
            TestAssert.True(annotation.X >= 60d && annotation.X + annotation.Width <= 216d, "Chart axis title links must sit in the left title zone.");
            TestAssert.True(annotation.Y >= 252d && annotation.Y + annotation.Height <= 468d, "Chart axis title links must sit inside the chart frame vertically.");
            TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable chart axis title hyperlink should not emit hyperlink diagnostics.");
        }
    }

    public static void PptxRotatedAxisTitleRunsSharePaintAndNavigationFrame()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (string side in new[] { "l", "r" })
        {
            PdfPage single = RenderChartTitlePage("", null, AxisTitleChart(side, """
                <a:r><a:rPr sz="1800" b="0"/><a:t>Value Axis</a:t></a:r>
                """), resolver).Pages.Single();
            PdfPage split = RenderChartTitlePage("", AxisTitleRelationships, AxisTitleChart(side, """
                <a:r><a:rPr sz="1800" b="0"/><a:t>Value </a:t></a:r>
                <a:r><a:rPr sz="1800" b="0"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Axis</a:t></a:r>
                """), resolver).Pages.Single();

            var expected = PortableRenderingTests.ReadPaintedSourceLabel(single, "Value Axis", fontPrefix: "CAT");
            var actual = PortableRenderingTests.ReadPaintedSourceLabel(split, "Value Axis", fontPrefix: "CAT");
            TestAssert.Equal(9, actual.Count);
            foreach (var (before, after) in expected.Zip(actual))
            {
                TestAssert.True(Math.Abs(before.Origin.X - after.Origin.X) < .002d &&
                    Math.Abs(before.Origin.Y - after.Origin.Y) < .08d,
                    $"Splitting the title preserves each glyph's position along the {side} axis.");
            }

            PdfLinkAnnotation link = split.Annotations.Single();
            TestAssert.Equal("https://example.invalid/chart-axis", link.Uri);
            TestAssert.True(actual.Skip(5).All(glyph => glyph.Ink.X >= link.X - .01d &&
                glyph.Ink.X + glyph.Ink.Width <= link.X + link.Width + .01d && glyph.Ink.Y >= link.Y - .01d &&
                glyph.Ink.Y + glyph.Ink.Height <= link.Y + link.Height + .01d), "The annotation covers the rotated second run.");
            TestAssert.True(actual.Take(5).All(glyph => glyph.Ink.Y + glyph.Ink.Height < link.Y),
                "The link excludes the first run.");
        }
    }

    public static void PptxDefaultAxisTitlesKeepAuthoredLinesInTheirBands()
    {
        const string first = "Northwestern Region";
        const string second = "International Sales";
        string firstRun = $"""<a:r><a:rPr sz="1800" b="0"/><a:t>{first}</a:t></a:r>""";
        string secondRun = $"""
            <a:r><a:rPr sz="1800" b="0"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>{second}</a:t></a:r>
            """;
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (string side in new[] { "l", "r", "t", "b" })
        {
            PdfPage page = Render(side, "<a:br/>");
            var leading = PortableRenderingTests.ReadPaintedSourceLabel(page, first, fontPrefix: "CAT");
            var trailing = PortableRenderingTests.ReadPaintedSourceLabel(page, second, fontPrefix: "CAT");
            PdfLinkAnnotation link = page.Annotations.Single();
            TestAssert.Equal("https://example.invalid/chart-axis", link.Uri);
            TestAssert.True(trailing.All(glyph => glyph.Ink.X >= link.X - .02d &&
                glyph.Ink.X + glyph.Ink.Width <= link.X + link.Width + .02d &&
                glyph.Ink.Y >= link.Y - .02d && glyph.Ink.Y + glyph.Ink.Height <= link.Y + link.Height + .02d),
                "The link follows the second line's visible ink.");
            if (side is "l" or "r")
                TestAssert.True(leading.Max(glyph => glyph.Ink.X + glyph.Ink.Width) < trailing.Min(glyph => glyph.Ink.X),
                    "The rotated lines occupy separate cross-axis positions.");
            else
                TestAssert.True(trailing.Max(glyph => glyph.Ink.Y + glyph.Ink.Height) < leading.Min(glyph => glyph.Ink.Y),
                    "The horizontal lines occupy separate baselines.");

            if (side == "r")
            {
                var tick = PortableRenderingTests.ReadPaintedSourceLabel(page, "5", fontPrefix: "CVA").Single();
                TestAssert.True(tick.Ink.X + tick.Ink.Width < leading.Min(glyph => glyph.Ink.X),
                    "The value ticks stay on the plot side of the caption.");
                PdfPage paragraphs = Render(side, "</a:p><a:p>");
                TestAssert.Equal(page.Content, paragraphs.Content);
                TestAssert.Equal(link, paragraphs.Annotations.Single());

                PdfPage composed = Render(side, "<a:br/>", compose: true);
                var main = PortableRenderingTests.ReadPaintedSourceLabel(composed, "Budget and Actual", fontPrefix: "CT");
                var topFirst = PortableRenderingTests.ReadPaintedSourceLabel(composed, "Quarterly Comparison", fontPrefix: "CAT");
                var topSecond = PortableRenderingTests.ReadPaintedSourceLabel(composed, "For Sales Owners", fontPrefix: "CAT");
                var right = PortableRenderingTests.ReadPaintedSourceLabel(composed, first, fontPrefix: "CAT");
                PortableRenderingTests.ReadPaintedSourceLabel(composed, second, fontPrefix: "CAT");
                TestAssert.True(main.Min(glyph => glyph.Ink.Y) > topFirst.Max(glyph => glyph.Ink.Y + glyph.Ink.Height),
                    "The main title has its own header above the top-axis caption.");
                TestAssert.True(topSecond.Max(glyph => glyph.Ink.X + glyph.Ink.Width) < right.Min(glyph => glyph.Ink.X),
                    "The top and right captions keep separate ink while sharing the chart.");
            }
        }

        PdfPage Render(string side, string boundary, bool compose = false)
        {
            string source = AxisTitleChart(side, firstRun + boundary + secondRun).Replace(
                """<c:majorUnit val="5"/><c:tickLblPos val="none"/>""",
                """<c:majorUnit val="5"/><c:tickLblPos val="nextTo"/>""", StringComparison.Ordinal);
            if (compose)
            {
                source = source.Replace("<c:chart>", "<c:chart>" + Title("Budget and Actual"), StringComparison.Ordinal)
                    .Replace("""<c:axPos val="b"/><c:crossAx val="20"/><c:tickLblPos val="none"/>""",
                        """<c:axPos val="t"/><c:crossAx val="20"/><c:tickLblPos val="none"/>""" +
                        Title("Quarterly Comparison", "For Sales Owners"), StringComparison.Ordinal);
            }
            return RenderChartTitlePage("", AxisTitleRelationships, source, resolver).Pages.Single();
        }

        static string Title(params string[] lines) =>
            "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>" +
            string.Join("<a:br/>", lines.Select(text => $"""<a:r><a:rPr sz="1400" b="0"/><a:t>{text}</a:t></a:r>""")) +
            "</a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
    }

    public static void PptxDefaultBarTopAxisTitleUsesMeasuredInkBand()
    {
        const string first = "Above";
        const string linked = "average";
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (int size in new[] { 1200, 3200 })
        {
            string chart = AxisTitleChart("t", $"""
                <a:r><a:rPr sz="1200" b="0"/><a:t xml:space="preserve">{first} </a:t></a:r>
                <a:r><a:rPr sz="{size}" b="0"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>{linked}</a:t></a:r>
                """).Replace("<c:barDir val=\"col\"/>", "<c:barDir val=\"bar\"/>", StringComparison.Ordinal)
                .Replace("<c:axPos val=\"b\"/><c:crossAx val=\"20\"/>",
                    "<c:axPos val=\"l\"/><c:crossAx val=\"20\"/>", StringComparison.Ordinal);
            PdfPage page = RenderChartTitlePage("", AxisTitleRelationships, chart, resolver).Pages.Single();
            var firstInk = PortableRenderingTests.ReadPaintedSourceLabel(page, first, fontPrefix: "CAT");
            var linkedInk = PortableRenderingTests.ReadPaintedSourceLabel(page, linked, fontPrefix: "CAT");
            TestAssert.Equal(firstInk[0].Origin.Y, linkedInk[0].Origin.Y);
            var link = page.Annotations.Single();
            TestAssert.Equal("https://example.invalid/chart-axis", link.Uri);
            TestAssert.True(linkedInk.All(glyph => glyph.Ink.X >= link.X - .02d &&
                glyph.Ink.X + glyph.Ink.Width <= link.X + link.Width + .02d && glyph.Ink.Y >= link.Y - .02d &&
                glyph.Ink.Y + glyph.Ink.Height <= link.Y + link.Height + .02d),
                "The link contains all matching axis-title ink.");
            TestAssert.True(firstInk.All(glyph => glyph.Ink.X + glyph.Ink.Width < link.X),
                "The unlinked first run retains its own space.");
        }
    }

    public static void PptxManualAxisTitleKeepsAuthoredLinesAndLinkedInk()
    {
        const string north = "<a:r><a:rPr sz=\"1800\" b=\"0\"/><a:t>North</a:t></a:r>";
        const string sales = """
            <a:fld id="{77777777-7777-7777-7777-777777777777}" type="datetime">
              <a:rPr sz="1800" b="0"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Sales</a:t>
            </a:fld>
            """;
        const string layout = """
            <c:layout><c:manualLayout><c:x val="0.36"/><c:y val="0.12"/>
              <c:wMode val="factor"/><c:hMode val="factor"/><c:w val="0.25"/><c:h val="0.3"/>
            </c:manualLayout></c:layout>
            """;
        var resolver = new PortableRenderingTests.PinnedResolver();
        PdfPage inlineBreak = Render("<a:br/>");
        var first = PortableRenderingTests.ReadPaintedSourceLabel(inlineBreak, "North", fontPrefix: "CAT");
        var second = PortableRenderingTests.ReadPaintedSourceLabel(inlineBreak, "Sales", fontPrefix: "CAT");
        TestAssert.True(first[0].Origin.Y > second[0].Origin.Y + 10d, "Each authored line keeps its own baseline in the manual box.");
        PdfLinkAnnotation link = inlineBreak.Annotations.Single();
        TestAssert.Equal("https://example.invalid/chart-axis", link.Uri);
        TestAssert.True(link.Y <= second[0].Origin.Y && link.Y + link.Height < first[0].Origin.Y,
            "The cached field link follows the second line.");

        PdfPage paragraphBreak = Render("</a:p><a:p>");
        TestAssert.Equal(inlineBreak.Content, paragraphBreak.Content);
        TestAssert.Equal(link, paragraphBreak.Annotations.Single());

        PdfPage Render(string boundary) =>
            RenderChartTitlePage("", AxisTitleRelationships, AxisTitleChart("l", north + boundary + sales, layout), resolver).Pages.Single();
    }

    private const string AxisTitleRelationships = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/chart-axis" TargetMode="External"/>
        </Relationships>
        """;

    private static string AxisTitleChart(string side, string runs, string layout = "") => $$"""
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
              <c:valAx><c:axId val="20"/><c:axPos val="{{side}}"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling>
                <c:crossAx val="10"/><c:majorUnit val="5"/><c:tickLblPos val="none"/>
                <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>{{runs}}</a:p></c:rich></c:tx>{{layout}}<c:overlay val="0"/></c:title>
              </c:valAx>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
        """;

    public static void PptxChartDataLabelCustomTextHyperlinkEmitsUriAnnotation()
    {
        foreach (IFontResolver? resolver in new IFontResolver?[] { null, new MissingTestFontResolver() })
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
                    """, fontResolver: resolver);

            PdfLinkAnnotation annotation = pages.Single().Annotations.Single();
            TestAssert.Equal("https://example.invalid/chart-label", annotation.Uri);
            TestAssert.True(annotation.Width > 0d && annotation.Height > 0d, "Chart data label links need a non-degenerate box.");
            TestAssert.True(annotation.Width < 288d, "Chart data label links must cover the run span, not the whole frame.");
            TestAssert.True(annotation.X >= 72d && annotation.X + annotation.Width <= 360d, "Chart data label links must sit inside the chart frame horizontally.");
            TestAssert.True(annotation.Y >= 252d && annotation.Y + annotation.Height <= 468d, "Chart data label links must sit inside the chart frame vertically.");
            TestAssert.True(diagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_HYPERLINK" && d.Id != "PPTX_UNSUPPORTED_HYPERLINK_ACTION"), "A resolvable chart data label hyperlink should not emit hyperlink diagnostics.");
        }
    }

    internal static (List<PdfPage> Pages, List<OoxPdfDiagnostic> Diagnostics) RenderChartTitlePage(string titleRunXml, string? chartRelsXml, string? chartXmlOverride = null, IFontResolver? fontResolver = null, bool? autoTitleDeleted = null)
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
                    {{(autoTitleDeleted is bool deleted ? $"<c:autoTitleDeleted val=\"{(deleted ? 1 : 0)}\"/>" : "")}}
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
        List<PdfPage> pages = new PptxRenderer(fontResolver).RenderPagesAsync(document, package, diagnostics.Add, CancellationToken.None).ToBlockingEnumerable().ToList();
        if (fontResolver is MissingTestFontResolver)
            TestAssert.True(pages[0].FallbackFonts.Count > 0 && pages[0].Fonts.Count == 0, "The missing-font control must use standard PDF fonts.");
        return (pages, diagnostics);
    }
}
