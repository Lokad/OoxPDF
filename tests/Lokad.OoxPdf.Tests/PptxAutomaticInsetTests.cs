using System.Text.Json;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxAutomaticInsetTests
{
    public static void PptxAutomaticColumnInsetsKeepTitleAndHorizontalLegendOwnership()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (int series in new[] { 1, 3 })
        foreach (int fontSize in new[] { 12, 20 })
        foreach (int width in new[] { 288, 576 })
        foreach (string legend in new[] { "none", "b", "t", "overlay" })
        {
            var baseline = Render(resolver, series, fontSize, width, "none", legend);
            foreach (string title in new[] { "automatic", "explicit" })
            {
                var actual = Render(resolver, series, fontSize, width, title, legend);
                TestAssert.True(Math.Abs(actual.Left - baseline.Left) < 0.01 && Math.Abs(actual.Right - baseline.Right) < 0.01,
                    $"{series} series, {fontSize}pt, {width}pt, {legend}, {title}: a title must not change horizontal plot reserves.");
            }
        }
    }

    public static void PptxAutomaticColumnInsetsRespectHiddenAndOppositeAxisLabels()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "AutomaticPlotInsets", "office.json"));
        using var reference = JsonDocument.Parse(File.ReadAllText(path));
        double leftPadding = reference.RootElement.GetProperty("HiddenAxisLeftInset").GetDouble();
        double rightPadding = reference.RootElement.GetProperty("HiddenAxisRightInset").GetDouble();
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (int width in new[] { 288, 576 })
        foreach (int fontSize in new[] { 12, 20 })
        foreach (string title in new[] { "none", "automatic", "explicit" })
        {
            var hidden = Render(resolver, 1, fontSize, width, title, "b", "none");
            TestAssert.True(Math.Abs(hidden.Left - 72 - leftPadding) < 0.1 &&
                Math.Abs(72 + width - hidden.Right - rightPadding) < 0.1,
                "Hidden value-axis labels retain independently measured Office frame padding.");
            var left = Render(resolver, 1, fontSize, width, title, "b");
            var right = Render(resolver, 1, fontSize, width, title, "b", "high");
            TestAssert.True(Math.Abs(left.Left - 72 - (72 + width - right.Right)) < 0.01 &&
                Math.Abs((72 + width - left.Right) - (right.Left - 72)) < 0.01,
                "Opposite value-axis labels mirror the measured strip without preserving a title preset floor.");
        }
    }

    private static (double Left, double Right) Render(PortableRenderingTests.PinnedResolver resolver,
        int seriesCount, int fontSize, int width, string title, string legend, string labels = "nextTo")
    {
        string titleXml = title == "explicit" ? """
            <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Supported title</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
            """ : "";
        string legendXml = legend == "none" ? "" : $"<c:legend><c:legendPos val=\"{(legend == "overlay" ? "b" : legend)}\"/><c:overlay val=\"{(legend == "overlay" ? 1 : 0)}\"/></c:legend>";
        string series = string.Concat(Enumerable.Range(0, seriesCount).Select(i => $$"""
            <c:ser><c:idx val="{{i}}"/><c:order val="{{i}}"/><c:tx><c:v>Series {{i + 1}}</c:v></c:tx>
              <c:spPr><a:solidFill><a:srgbClr val="0F9ED5"/></a:solidFill></c:spPr>
              <c:val><c:numLit><c:ptCount val="6"/>
              {{string.Concat(Enumerable.Range(0, 6).Select(j => $"<c:pt idx=\"{j}\"><c:v>4</c:v></c:pt>"))}}
              </c:numLit></c:val></c:ser>
            """));
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:style val="2"/><c:chart>{{titleXml}}<c:autoTitleDeleted val="{{(title == "none" ? 1 : 0)}}"/>
                <c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{{series}}<c:axId val="1"/><c:axId val="2"/></c:barChart>
                  <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
                  <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="l"/>
                    <c:tickLblPos val="{{labels}}"/><c:majorUnit val="1"/><c:crossAx val="1"/></c:valAx>
                </c:plotArea>{{legendXml}}
              </c:chart><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{fontSize * 100}}"/></a:pPr></a:p></c:txPr>
            </c:chartSpace>
            """;
        string slide = $$"""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="{{width * 12700}}" cy="4572000"/></p:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            </p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart, new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        var page = new PptxRenderer(resolver).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
        var bars = PptxChartSpacingTests.ReadBarRectangles(page.Content).Where(rect => rect.Height > 24).OrderBy(rect => rect.X).ToArray();
        TestAssert.Equal(6 * seriesCount, bars.Length);
        double first = bars.Take(seriesCount).Average(rect => rect.X + rect.Width / 2);
        double last = bars.Skip(5 * seriesCount).Average(rect => rect.X + rect.Width / 2);
        double band = (last - first) / 5;
        return (first - band / 2, last + band / 2);
    }
}
