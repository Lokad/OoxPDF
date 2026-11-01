using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;
using System.Text.RegularExpressions;

namespace Lokad.OoxPdf.Tests;

internal static class PptxCartesianIntervalTests
{
    public static void PptxRightLegendLineReservesBoundsBeyondPaintedIntervals()
    {
        foreach (int size in new[] { 12, 24 })
        {
            var below = NumericLabels(Render(99, size));
            var above = NumericLabels(Render(103, size));
            var aligned = NumericLabels(Render(107, size));
            string[] labels = ["7", "27", "47", "67", "87"];
            TestAssert.True(below.Select(x => x.Text).SequenceEqual(labels) && above.Select(x => x.Text).SequenceEqual(labels),
                "An unaligned upper bound does not add a tick outside the authored minimum/major-unit lattice.");
            TestAssert.True(aligned.Select(x => x.Text).SequenceEqual(labels.Append("107")),
                "An aligned upper bound retains its actual tick label.");
            double x = below.Single(x => x.Text == "87").X;
            // Independent PowerPoint controls retain three-digit clearance for103
            // even though87 is the final painted major tick. Measurement and
            // emitted intervals deliberately have different endpoint contracts.
            double reserved = above.Single(x => x.Text == "87").X;
            TestAssert.True(reserved > x + size * .3d,
                "An authored three-digit upper bound reserves advance even without a tick at that bound.");
            TestAssert.True(Math.Abs(aligned.Single(x => x.Text == "87").X - reserved) < .002d,
                "An aligned three-digit upper tick retains the same numeric strip.");
        }
    }

    public static void PptxAreaSourceIntervalsShareLabelsGridsAndAutomaticMinimum()
    {
        foreach (int maximum in new[] { 103, 107 })
        {
            PdfPage page = Render(maximum, 18, area: true, minorGridlines: true);
            string[] expected = maximum == 103 ? ["7", "27", "47", "67", "87"] : ["7", "27", "47", "67", "87", "107"];
            TestAssert.True(NumericLabels(page).Select(x => x.Text).SequenceEqual(expected),
                "Area labels retain the independent Office source-minimum interval origin.");
            TestAssert.Equal(expected.Length - 1, PaintedGridLevels(page, "1 0 0"));
            TestAssert.Equal(20, PaintedGridLevels(page, "0 1 0"));
        }
        PdfPage automatic = Render(103, 18, area: true, minimum: null);
        TestAssert.True(NumericLabels(automatic).Select(x => x.Text).SequenceEqual(new[] { "-17", "3", "23", "43", "63", "83", "103" }),
            "An authored maximum and major unit anchor the automatic lower bound for area charts too.");
        PdfPage zero = Render(103, 18, area: true, minimum: 0, minorGridlines: true);
        TestAssert.True(NumericLabels(zero).Select(x => x.Text).SequenceEqual(new[] { "0", "20", "40", "60", "80", "100" }),
            "An unaligned upper endpoint is reserved without inventing a major tick.");
        TestAssert.Equal(5, PaintedGridLevels(zero, "1 0 0"));
        TestAssert.Equal(20, PaintedGridLevels(zero, "0 1 0"));
    }

    private static int PaintedGridLevels(PdfPage page, string color)
    {
        Match path = Regex.Match(page.Content, Regex.Escape(color) + @" RG(?<body>.*?)\r?\nS(?:\r?\n|$)", RegexOptions.Singleline);
        TestAssert.True(path.Success, "Expected independently colored grid path.");
        return Regex.Matches(path.Groups["body"].Value, @"[-\d.]+ [-\d.]+ l").Count;
    }

    public static void PptxSourceIntervalsShareAutomaticUnitsAcrossLineAndArea()
    {
        foreach (bool area in new[] { false, true })
        {
            PdfPage page = Render(103, 12, area: area, minorGridlines: true, automaticMajorUnit: true);
            TestAssert.True(NumericLabels(page).Select(x => x.Text).SequenceEqual(
                Enumerable.Range(0, 10).Select(i => (7 + 10 * i).ToString(System.Globalization.CultureInfo.InvariantCulture))),
                "Independent Office selects a10-unit major interval at these source bounds.");
            TestAssert.Equal(9, PaintedGridLevels(page, "1 0 0"));
            TestAssert.Equal(39, PaintedGridLevels(page, "0 1 0"));
        }
    }

    public static void PptxStandaloneLineAndAreaReserveUnpaintedRangeEndpoints()
    {
        foreach (bool area in new[] { false, true })
        foreach (int size in new[] { 12, 24 })
        {
            var below = NumericLabels(Render(99, size, area: area, legend: false));
            var above = NumericLabels(Render(103, size, area: area, legend: false));
            TestAssert.True(below.Select(x => x.Text).SequenceEqual(above.Select(x => x.Text)),
                "Office99/103 controls retain the same painted major labels.");
            TestAssert.True(above.Single(x => x.Text == "87").X > below.Single(x => x.Text == "87").X + size * .3d,
                "The numeric strip reserves the authored endpoint independently of the painted major labels.");
        }
    }

    private static (string Text, double X)[] NumericLabels(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal))
            .Select(show => (show.Text, show.X)).ToArray();

    private static PdfPage Render(int maximum, int size, bool area = false, int? minimum = 7, bool minorGridlines = false, bool automaticMajorUnit = false, bool legend = true)
    {
        string kind = area ? "areaChart" : "lineChart";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:{{kind}}><c:varyColors val="0"/>
                <c:ser><c:tx><c:v>Series</c:v></c:tx>{{(area ? "" : "<c:marker><c:symbol val=\"none\"/></c:marker>")}}
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>27</c:v></c:pt><c:pt idx="1"><c:v>47</c:v></c:pt><c:pt idx="2"><c:v>67</c:v></c:pt></c:numLit></c:val>
                </c:ser><c:axId val="1"/><c:axId val="2"/></c:{{kind}}>
                <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling>{{(minimum is { } lower ? "<c:min val=\"" + lower + "\"/>" : "")}}<c:max val="{{maximum}}"/></c:scaling><c:axPos val="l"/>
                  <c:crossAx val="1"/>{{(automaticMajorUnit ? "" : "<c:majorUnit val=\"20\"/>")}}
                  <c:majorGridlines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
                  {{(minorGridlines ? "<c:minorGridlines><c:spPr><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:ln></c:spPr></c:minorGridlines>" : "")}}
                  <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{size * 100}}"/></a:pPr></a:p></c:txPr>
                </c:valAx>
              </c:plotArea>{{(legend ? "<c:legend><c:legendPos val=\"r\"/><c:overlay val=\"0\"/></c:legend>" : "")}}</c:chart>
            </c:chartSpace>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
