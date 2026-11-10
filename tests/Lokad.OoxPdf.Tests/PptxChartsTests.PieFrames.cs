using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartsTests
{
    public static void PptxPieHorizontalLegendsUseSelectedPlotMidpoint()
    {
        foreach (string position in new[] { "t", "b" })
        foreach (var (x, width) in new[] { (72, 288), (36, 420), (117, 576) })
        {
            string content = RenderPieFrame(x, 72, width, 216,
                $"<c:legend><c:legendPos val=\"{position}\"/><c:overlay val=\"0\"/></c:legend>");
            // Both emitted sectors close through the shared center. Inspect the
            // placed geometry, rather than a private ratio or Bezier control box.
            var centers = Regex.Matches(content,
                @"(?m)^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) l\r?\nh\r?\nf\*");
            TestAssert.Equal(2, centers.Count);
            foreach (Match center in centers)
                TestAssert.True(Math.Abs(double.Parse(center.Groups[1].Value, CultureInfo.InvariantCulture) - (x + width / 2d)) < .001d,
                    $"A {position} legend must keep the pie centered in its authored {width} pt plot at X={x}.");
        }
    }

    public static void PptxUntitledUnreservedPiesUseVerticalMidpoint()
    {
        foreach (bool overlay in new[] { false, true })
        foreach (bool labels in new[] { false, true })
        foreach (var (y, height) in new[] { (36, 432), (108, 360), (84, 216) })
        {
            string legend = overlay
                ? "<c:legend><c:legendPos val=\"r\"/><c:overlay val=\"1\"/></c:legend>" : "";
            string content = RenderPieFrame(72, y, 576, height, legend,
                labels ? "<c:dLbls><c:showVal val=\"1\"/></c:dLbls>" : "");
            var centers = Regex.Matches(content,
                @"(?m)^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) l\r?\nh\r?\nf\*");
            TestAssert.Equal(2, centers.Count);
            // BasicPresentation is 540pt tall; PDF content uses bottom-up Y.
            double expected = 540d - y - height / 2d;
            foreach (Match center in centers)
                TestAssert.True(Math.Abs(double.Parse(center.Groups[2].Value, CultureInfo.InvariantCulture) - expected) < .001d,
                    $"Untitled pies must retain the vertical midpoint with labels={labels}, overlay={overlay}, Y={y}, H={height}.");
        }
    }

    private static string RenderPieFrame(int x, int y, int width, int height, string legend, string labels = "")
    {
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:pieChart><c:varyColors val="1"/><c:ser>
                <c:idx val="0"/><c:order val="0"/>
                <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat>
                <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val>
              </c:ser>{{labels}}</c:pieChart></c:plotArea>{{legend}}</c:chart>
            </c:chartSpace>
            """;
        string slide = $$"""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:cSld><p:spTree><p:graphicFrame><p:xfrm>
                <a:off x="{{x * 12700}}" y="{{y * 12700}}"/><a:ext cx="{{width * 12700}}" cy="{{height * 12700}}"/>
              </p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rId1"/>
              </a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            </p:sld>
            """;
        using var bytes = PptxTests.CreateSingleChartPackageStream(chart,
            new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(bytes, default);
        var source = new PptxReader().Read(package, default);
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(source, package, null, default)
            .ToBlockingEnumerable().Single();
        return page.Content;
    }
}
