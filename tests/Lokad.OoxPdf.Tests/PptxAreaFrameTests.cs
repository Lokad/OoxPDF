using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxAreaFrameTests
{
    public static void PptxAreaFrameMeasuresTitlesAndBottomLegendsTogether()
    {
        foreach (int size in new[] { 12, 24 })
        {
            var plain = Plot(Render(size));
            var bottom = Plot(Render(size, legend: true));
            var titled = Plot(Render(size, legend: true, title: true));
            TestAssert.True(Math.Abs(plain.X - bottom.X) < .002d && Math.Abs(bottom.X - titled.X) < .002d &&
                Math.Abs(plain.Width - bottom.Width) < .002d && Math.Abs(bottom.Width - titled.Width) < .002d,
                "Horizontal numeric strips are independent of the title and bottom legend.");
            TestAssert.True(bottom.Y > plain.Y + size / 2d && Math.Abs(bottom.Y + bottom.Height - plain.Y - plain.Height) < .002d,
                "A bottom legend reserves its measured band without moving the plot top.");
            TestAssert.True(Math.Abs(titled.Y - bottom.Y) < .002d && titled.Y + titled.Height < bottom.Y + bottom.Height - size / 2d,
                "A title reserves space above the plot while retaining the same legend/caption band below.");
            var overlaid = Plot(Render(size, legend: true, overlay: true));
            TestAssert.True(Math.Abs(plain.Y - overlaid.Y) < .002d && Math.Abs(plain.Height - overlaid.Height) < .002d,
                "An overlay legend does not consume plot height.");
        }
    }

    public static void PptxAreaFrameAndCaptionsShareTheSourceCrossing()
    {
        foreach (bool reversed in new[] { false, true })
        {
            PdfPage next = Render(12, legend: true, minimum: -2, reversed: reversed);
            PdfPage low = Render(12, legend: true, minimum: -2, reversed: reversed, caption: "low");
            var nextPlot = Plot(next);
            var lowPlot = Plot(low);
            double nextY = FirstCaptionY(next), lowY = FirstCaptionY(low);
            TestAssert.True(nextPlot.Height > lowPlot.Height + 1d,
                "Captions at an interior crossing release the outer caption strip for the plot.");
            TestAssert.True(reversed ? nextY < lowY - 1d : nextY > lowY + 1d,
                "Next-to captions follow the crossing and value direction; low remains at the numeric low edge.");
        }
    }

    public static void PptxAreaPrimaryValueLabelsUseTheirMeasuredSide()
    {
        foreach (bool right in new[] { false, true })
        {
            PdfPage page = Render(18, rightAxis: right);
            var plot = Plot(page);
            var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).ToArray();
            TestAssert.True(labels.Length >= 2, "Expected the primary numeric interval labels.");
            TestAssert.True(labels.All(x => right ? x.X > plot.X + plot.Width : x.X < plot.X),
                "Painting and numeric strip measurement must use the same primary axis side.");
        }
    }

    private static double FirstCaptionY(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).First().Y;

    private static (double X, double Y, double Width, double Height) Plot(PdfPage page)
    {
        var paths = new List<string>();
        var paints = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add, strokes: true).ToArray();
        string rectangle = paints.Select((color, i) => (color, path: paths[i]))
            .Single(x => x.color == new RgbColor(255, 0, 0) && x.path.EndsWith(" re S", StringComparison.Ordinal)).path;
        double[] v = rectangle[..^5].Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(x => double.Parse(x, CultureInfo.InvariantCulture)).ToArray();
        return (v[0], v[1], v[2], v[3]);
    }

    private static PdfPage Render(int size, bool legend = false, bool title = false, bool overlay = false,
        int minimum = 0, bool reversed = false, string caption = "nextTo", bool rightAxis = false)
    {
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>{{(title ? "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz=\"" + size * 100 + "\"/><a:t>Area</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>" : "<c:autoTitleDeleted val=\"1\"/>")}}
                <c:plotArea><c:areaChart><c:grouping val="standard"/><c:ser><c:tx><c:v>Series</c:v></c:tx>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>6</c:v></c:pt></c:numLit></c:val>
                </c:ser><c:axId val="1"/><c:axId val="2"/></c:areaChart>
                <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="{{caption}}"/><c:crossAx val="2"/><c:crosses val="{{(rightAxis ? "max" : "autoZero")}}"/></c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling><c:min val="{{minimum}}"/><c:max val="10"/><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/></c:scaling><c:axPos val="{{(rightAxis ? "r" : "l")}}"/><c:crossAx val="1"/><c:crossBetween val="between"/><c:majorUnit val="2"/><c:majorGridlines/></c:valAx>
                <c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
                </c:plotArea>{{(legend ? "<c:legend><c:legendPos val=\"b\"/><c:overlay val=\"" + (overlay ? "1" : "0") + "\"/></c:legend>" : "")}}
              </c:chart><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{size * 100}}"/></a:pPr></a:p></c:txPr>
            </c:chartSpace>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
