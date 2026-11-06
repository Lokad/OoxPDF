using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxAutomaticScaleTests
{
    public static void PptxAutomaticColumnScaleRespondsToFrameAndSharesItsDataCoordinates()
    {
        var small = Read(Render(120));
        var tall = Read(Render(480));
        TestAssert.True(small.Values.Max() > tall.Values.Max(),
            "A crowded frame chooses a coarser outer scale rather than merely hiding captions on the old scale.");
        foreach (var sample in new[] { small, tall })
        {
            TestAssert.Equal(4, sample.Bars.Length);
            double plotHeight = sample.Baselines.Max() - sample.Baselines.Min();
            TestAssert.True(Math.Abs(sample.Bars[0] / plotHeight - 4d / sample.Values.Max()) < .001d,
                "Painted data and numeric coordinates consume the same resolved scale.");
        }
        TestAssert.Equal(5d, small.Values.Max());
        TestAssert.Equal(4.5d, tall.Values.Max());
    }

    public static void PptxAutomaticColumnScaleMatchesTheEquivalentAuthoredIntervals()
    {
        var automatic = Read(Render(120));
        var authored = Read(Render(120, declaration: "<c:min val=\"0\"/><c:max val=\"5\"/>", major: "5"));
        TestAssert.True(automatic.Values.SequenceEqual(authored.Values) && automatic.Values.SequenceEqual(new[] { 0d, 5d }),
            "The selected major unit governs captions as well as the bounds and grids.");
        foreach (var (first, second) in new[] { (automatic.Baselines, authored.Baselines), (automatic.Xs, authored.Xs),
            (automatic.Bars, authored.Bars), (automatic.BarXs, authored.BarXs) })
            TestAssert.True(first.Length == second.Length && first.Zip(second).All(pair => Math.Abs(pair.First - pair.Second) < .002d),
                "Equivalent automatic and authored scales have the same numeric strip and data geometry.");
    }

    public static void PptxAutomaticColumnScaleKeepsAuthoredAndUnknownDeclarations()
    {
        foreach (int height in new[] { 120, 480 })
        {
            var authored = Read(Render(height, declaration: "<c:min val=\"0\"/><c:max val=\"8\"/>", major: "2"));
            TestAssert.True(authored.Values.SequenceEqual(new[] { 0d, 2d, 4d, 6d, 8d }),
                "Authored intervals survive a change to the available frame.");
            var unknown = Read(Render(height, major: "future"));
            TestAssert.Equal(4.5d, unknown.Values.Max());
            TestAssert.Equal(4, unknown.Bars.Length);
        }
        var positive = Read(Render(120));
        var negative = Read(Render(120, value: -4d));
        TestAssert.True(positive.Values.Select(value => -value).Order().SequenceEqual(negative.Values.Order()),
            "Negative-only source values retain the reflected scale contract.");
        TestAssert.True(Math.Abs(positive.Bars[0] - negative.Bars[0]) < .002d,
            "Reflecting the source preserves the data length.");
    }

    private static (double[] Values, double[] Baselines, double[] Bars, double[] Xs, double[] BarXs) Read(PdfPage page)
    {
        var fonts = page.Fonts.Where(font => font.ResourceName.StartsWith("CVA", StringComparison.Ordinal))
            .Select(font => font.Font).ToHashSet();
        var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Where(show => fonts.Contains(show.Font)).ToArray();
        double[] values = shows.Select(show => double.Parse(show.Text.Replace('\u2212', '-'), CultureInfo.InvariantCulture)).ToArray();
        double[][] bars = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re f")
            .Select(match => match.Groups.Cast<Group>().Skip(1).Select(group => double.Parse(group.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(rect => rect[2] > 10d && rect[2] < 100d && rect[3] > 10d).ToArray();
        return (values, shows.Select(show => show.Y).ToArray(), bars.Select(rect => rect[3]).ToArray(),
            shows.Select(show => show.X).ToArray(), bars.Select(rect => rect[0]).ToArray());
    }

    private static PdfPage Render(int height, double value = 4d, string declaration = "", string? major = null)
    {
        string categories = string.Concat(Enumerable.Range(0, 4).Select(index => $"<c:pt idx=\"{index}\"><c:v>Area {index}</c:v></c:pt>"));
        string values = string.Concat(Enumerable.Range(0, 4).Select(index => $"<c:pt idx=\"{index}\"><c:v>{value.ToString(CultureInfo.InvariantCulture)}</c:v></c:pt>"));
        string slide = $$"""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="{{height * 12700}}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            </p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream($$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
                <c:ser><c:cat><c:strLit><c:ptCount val="4"/>{{categories}}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val="4"/>{{values}}</c:numLit></c:val></c:ser>
                <c:axId val="1"/><c:axId val="2"/>
              </c:barChart>
              <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/><c:crosses val="autoZero"/>
                <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr></a:p></c:txPr>
              </c:catAx>
              <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/><c:crosses val="autoZero"/>
                <c:scaling><c:orientation val="minMax"/>{{declaration}}</c:scaling>{{(major is null ? "" : $"<c:majorUnit val=\"{major}\"/>")}}
                <c:majorGridlines/><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="2400"/></a:pPr></a:p></c:txPr>
              </c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
            """, new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }
}
