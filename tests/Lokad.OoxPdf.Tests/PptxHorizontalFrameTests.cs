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

    private static (double Left, double Top, double Bottom) Render(int height, int categorySize,
        int valueSize, string legend, bool valueLabels = true, bool categoryLabels = true)
    {
        string legendXml = legend == "none" ? "" : $"<c:legend><c:legendPos val=\"b\"/><c:overlay val=\"{(legend == "overlay" ? 1 : 0)}\"/><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"1200\"/></a:pPr></a:p></c:txPr></c:legend>";
        string values = string.Concat(Enumerable.Range(0, 6).Select(i => $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
        string categories = string.Concat(Enumerable.Range(0, 6).Select(i => $"<c:pt idx=\"{i}\"><c:v>Category {i+1}</c:v></c:pt>"));
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:style val="2"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
                <c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>
                  <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series</c:v></c:tx>
                    <c:spPr><a:solidFill><a:srgbClr val="0F9ED5"/></a:solidFill></c:spPr>
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
              <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="7315200" cy="{{height*12700}}"/></p:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            </p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart,
            new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
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
        return (bars[0].X, page.Height - last - band / 2, page.Height - first + band / 2);
    }
}
