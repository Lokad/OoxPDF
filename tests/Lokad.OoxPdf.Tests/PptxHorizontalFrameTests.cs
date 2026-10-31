using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxHorizontalFrameTests
{
    public static void PptxHorizontalFrameKeepsUntitledTopAcrossHeightAndLegendChanges()
    {
        foreach (int height in new[] { 300, 360, 420 })
        foreach (string legend in new[] { "none", "b", "overlay" })
        foreach (bool valueLabels in new[] { false, true })
        {
            var actual = Render(height, 12, 12, legend, valueLabels);
            TestAssert.True(Math.Abs(actual.Top - 119d) < 0.1,
                $"{height}pt frame, {legend}, labels={valueLabels}: Office's untitled top stays 11pt below the frame top.");
        }
    }

    public static void PptxHorizontalFrameAssignsFontsToTheirAxisStrips()
    {
        var baseline = Render(360, 12, 12, "b");
        var categorySmall = Render(360, 8, 12, "b");
        var categoryLarge = Render(360, 18, 12, "b");
        TestAssert.True(categoryLarge.Left > categorySmall.Left + 1,
            "Larger category labels occupy the side strip.");
        TestAssert.True(Math.Abs(categorySmall.Bottom - baseline.Bottom) < 0.01 &&
            Math.Abs(categoryLarge.Bottom - baseline.Bottom) < 0.01,
            "Category font size must not change the value-axis bottom strip.");
        var valueSmall = Render(360, 12, 8, "b");
        var valueLarge = Render(360, 12, 18, "b");
        TestAssert.True(Math.Abs(valueSmall.Left - baseline.Left) < 0.01 &&
            Math.Abs(valueLarge.Left - baseline.Left) < 0.01,
            "Value-axis font size must not change the occupied category strip.");
        TestAssert.True(valueLarge.Bottom < valueSmall.Bottom - 1,
            "Larger value labels occupy more space below the plot.");
        foreach (int categorySize in new[] { 8, 18 })
        {
            var hidden = Render(360, categorySize, 12, "b", categoryLabels: false);
            var otherHidden = Render(360, 12, 12, "b", categoryLabels: false);
            TestAssert.True(Math.Abs(hidden.Left - otherHidden.Left) < 0.01,
                "Hidden categories do not retain a font-dependent side strip.");
        }
    }

    public static void PptxHorizontalFrameOverlayLegendDoesNotConsumePlotSpace()
    {
        foreach (int categorySize in new[] { 8, 18 })
        foreach (bool valueLabels in new[] { false, true })
        {
            var plain = Render(360, categorySize, 12, "none", valueLabels);
            var overlay = Render(360, categorySize, 12, "overlay", valueLabels);
            TestAssert.True(Math.Abs(plain.Left - overlay.Left) < 0.01 &&
                Math.Abs(plain.Top - overlay.Top) < 0.01 && Math.Abs(plain.Bottom - overlay.Bottom) < 0.01,
                "An overlay legend must retain the same plot frame as the unreserved chart.");
        }
    }

    public static void PptxHorizontalLegendsReserveTheirOwnMeasuredRows()
    {
        foreach (int fontSize in new[] { 8, 12, 18 })
        foreach (bool categories in new[] { false, true })
        {
            var plain = Render(360, 12, 12, "none");
            var top = Render(360, 12, 12, "t", legendSize: fontSize, categoryLegend: categories);
            var bottom = Render(360, 12, 12, "b", legendSize: fontSize, categoryLegend: categories);
            TestAssert.True(Math.Abs(top.Bottom - plain.Bottom) < 0.01 && Math.Abs(bottom.Top - plain.Top) < 0.01,
                "A horizontal legend reserves only its selected side.");
            TestAssert.True(Math.Abs((top.Top - plain.Top) - (plain.Bottom - bottom.Bottom)) < 0.01,
                "The same measured rows occupy the same band above or below the plot.");
            TestAssert.True(top.Top > plain.Top + 1 && bottom.Bottom < plain.Bottom - 1,
                "Visible horizontal legends require an admitted band.");
        }
    }

    public static void PptxHorizontalLegendsAdmitWrappedRowsWithoutLosingNames()
    {
        foreach (string side in new[] { "t", "b" })
        {
            var shortNames = Render(360, 12, 12, side, categoryLabels: false, categoryLegend: true);
            var wrapped = Render(360, 12, 12, side, categoryLabels: false, categoryLegend: true,
                width: 288, longNames: true);
            TestAssert.True(side == "t" ? wrapped.Top > shortNames.Top + 1 : wrapped.Bottom < shortNames.Bottom - 1,
                "Wrapped legend rows reserve more space on their own side.");
            for (int i = 1; i <= 6; i++)
                TestAssert.True(wrapped.Text.Contains("North West Region " + i, StringComparison.Ordinal),
                    "Every supported wrapped legend name must remain painted.");
        }
    }

    public static void PptxHorizontalLegendAndTitleBandsCompose()
    {
        var plain = Render(360, 12, 12, "none");
        var legend = Render(360, 12, 12, "t", categoryLegend: true);
        var title = Render(360, 12, 12, "none", title: true);
        var together = Render(360, 12, 12, "t", title: true, categoryLegend: true);
        TestAssert.True(Math.Abs((together.Top - title.Top) - (legend.Top - plain.Top)) < 0.01,
            "Title and legend bands compose without either replacing the other.");
        TestAssert.True(together.Text.Contains("Supported title", StringComparison.Ordinal),
            "The supported title remains painted alongside the legend.");
    }

    public static void PptxHorizontalTitlesLeaveValueAxisOwnershipUnchanged()
    {
        foreach (int valueSize in new[] { 8, 12, 18 })
        foreach (bool visible in new[] { false, true })
        foreach (string side in new[] { "none", "b", "t" })
        {
            var plain = Render(360, 12, valueSize, side, visible);
            var titled = Render(360, 12, valueSize, side, visible, title: true);
            TestAssert.True(Math.Abs(plain.Bottom - titled.Bottom) < 0.01,
                "A chart title must not replace the measured value-axis strip.");
        }
    }

    public static void PptxHorizontalTitleMetricsFollowRichRunsAndPreserveSplits()
    {
        var small = Render(360, 12, 12, "none", title: true, titleSize: 12);
        var medium = Render(360, 12, 12, "none", title: true, titleSize: 18);
        var large = Render(360, 12, 12, "none", title: true, titleSize: 24);
        TestAssert.True(medium.Top > small.Top + 1 && large.Top > medium.Top + 1,
            "The painted rich-run size must own title occupancy.");
        TestAssert.True(medium.TitleBaseline > small.TitleBaseline + 1 && large.TitleBaseline > medium.TitleBaseline + 1,
            "The title baseline must follow its rich-run metrics, including below the old preset floor.");
        foreach (int height in new[] { 300, 360, 420 })
        {
            var whole = Render(height, 12, 12, "none", title: true);
            var split = Render(height, 12, 12, "none", title: true, splitTitle: true);
            TestAssert.True(Math.Abs(whole.Top - medium.Top) < 0.01 && Math.Abs(whole.TitleBaseline - medium.TitleBaseline) < 0.01,
                "A fixed-size title remains anchored to the frame top when frame height changes.");
            TestAssert.True(Math.Abs(whole.Top - split.Top) < 0.01 && Math.Abs(whole.TitleBaseline - split.TitleBaseline) < 0.01 &&
                split.Text.Contains("Supported title", StringComparison.Ordinal),
                "Splitting an unchanged rich run preserves title geometry and content.");
        }
        var mixed = Render(360, 12, 12, "none", title: true, mixedTitle: true);
        TestAssert.True(Math.Abs(mixed.Top - large.Top) < 0.01 && Math.Abs(mixed.TitleBaseline - large.TitleBaseline) < 0.01,
            "The largest supported run owns the shared title line, regardless of its source position.");
    }

    public static void PptxHorizontalOverlayTitlesKeepTheUnreservedPlot()
    {
        foreach (string side in new[] { "none", "b", "t" })
        {
            var plain = Render(360, 12, 12, side);
            var overlay = Render(360, 12, 12, side, title: true, overlayTitle: true);
            TestAssert.True(Math.Abs(plain.Left - overlay.Left) < 0.01 && Math.Abs(plain.Top - overlay.Top) < 0.01 &&
                Math.Abs(plain.Bottom - overlay.Bottom) < 0.01 && overlay.Text.Contains("Supported title", StringComparison.Ordinal),
                "An overlay title preserves the ordinary axis/legend frame and remains painted.");
        }
    }

    public static void PptxHorizontalTitleLinksKeepTheCompleteLineInsideItsFrame()
    {
        foreach (int size in new[] { 18, 48, 72 })
        {
            var actual = Render(360, 12, 12, "none", title: true, titleSize: size, linkedTitle: true);
            TestAssert.True(actual.LinkTop is >= 107.999 && actual.LinkHeight > size * 0.9 &&
                actual.LinkTop + actual.LinkHeight <= 468.001,
                "The supported linked title line must retain its full annotation inside the source frame.");
        }
    }

    private static (double Left, double Top, double Bottom, string Text, double TitleBaseline, double? LinkTop, double? LinkHeight) Render(int height, int categorySize,
        int valueSize, string legend, bool valueLabels = true, bool categoryLabels = true,
        int legendSize = 12, bool categoryLegend = false, int width = 576, bool longNames = false, bool title = false,
        int titleSize = 18, bool splitTitle = false, bool mixedTitle = false, bool overlayTitle = false, bool linkedTitle = false)
    {
        string legendXml = legend == "none" ? "" : $"<c:legend><c:legendPos val=\"{(legend == "t" ? "t" : "b")}\"/><c:overlay val=\"{(legend == "overlay" ? 1 : 0)}\"/><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"{legendSize*100}\"/></a:pPr></a:p></c:txPr></c:legend>";
        string values = string.Concat(Enumerable.Range(0, 6).Select(i => $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
        string categories = string.Concat(Enumerable.Range(0, 6).Select(i => $"<c:pt idx=\"{i}\"><c:v>{(longNames ? "North West Region " : "Category ")}{i+1}</c:v></c:pt>"));
        string pointStyle = categoryLegend ? "<c:dPt><c:idx val=\"0\"/><c:spPr/></c:dPt>" : "";
        string TitleRun(string text, int size) => $"<a:r><a:rPr sz=\"{size*100}\">{(linkedTitle ? "<a:hlinkClick r:id=\"titleLink\"/>" : "")}</a:rPr><a:t>{text}</a:t></a:r>";
        string titleRuns = mixedTitle ? TitleRun("Supported", 12) + TitleRun(" title", 24)
            : splitTitle ? TitleRun("Supported", titleSize) + TitleRun(" title", titleSize) : TitleRun("Supported title", titleSize);
        string titleXml = title ? "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>" + titleRuns +
            "</a:p></c:rich></c:tx><c:overlay val=\"" + (overlayTitle ? 1 : 0) + "\"/></c:title>" : "";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <c:style val="2"/><c:chart>{{titleXml}}<c:autoTitleDeleted val="{{(title ? 0 : 1)}}"/><c:plotArea>
                <c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
                  <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series</c:v></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="0F9ED5"/></a:solidFill></c:spPr>{{pointStyle}}
                    <c:cat><c:strLit><c:ptCount val="6"/>{{categories}}</c:strLit></c:cat>
                    <c:val><c:numLit><c:ptCount val="6"/>{{values}}</c:numLit></c:val>
                  </c:ser><c:axId val="1"/><c:axId val="2"/>
                </c:barChart>
                <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/>
                  <c:tickLblPos val="{{(categoryLabels ? "nextTo" : "none")}}"/><c:crossAx val="2"/>
                  <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{categorySize*100}}"/></a:pPr></a:p></c:txPr>
                </c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="5"/></c:scaling>
                  <c:axPos val="b"/><c:tickLblPos val="{{(valueLabels ? "nextTo" : "none")}}"/><c:majorUnit val="1"/><c:crossAx val="1"/>
                  <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{valueSize*100}}"/></a:pPr></a:p></c:txPr>
                </c:valAx>
              </c:plotArea>{{legendXml}}</c:chart>
            </c:chartSpace>
            """;
        string slide = $$"""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="{{width*12700}}" cy="{{height*12700}}"/></p:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            </p:sld>
            """;
        var parts = new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) };
        if (linkedTitle)
            parts["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="titleLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/title" TargetMode="External"/></Relationships>
                """);
        using var stream = PptxTests.CreateSingleChartPackageStream(chart, parts);
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
        var bars = PptxChartSpacingTests.ReadBarRectangles(page.Content).Where(rect => rect.Width > 50 && rect.Height > 2)
            .OrderBy(rect => rect.Y).ToArray();
        TestAssert.Equal(6, bars.Length);
        double first = bars[0].Y + bars[0].Height / 2;
        double last = bars[^1].Y + bars[^1].Height / 2;
        double band = (last - first) / 5;
        var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray();
        string painted = string.Concat(shows.Select(show => show.Text));
        double titleBaseline = title ? page.Height - shows.First(show => show.Text.Contains("Supported", StringComparison.Ordinal)).Y : 0d;
        var link = page.Annotations.SingleOrDefault(annotation => annotation.Uri == "https://example.invalid/title");
        return (bars[0].X, page.Height - last - band / 2, page.Height - first + band / 2, painted, titleBaseline,
            link is null ? null : page.Height - link.Y - link.Height, link?.Height);
    }
}
