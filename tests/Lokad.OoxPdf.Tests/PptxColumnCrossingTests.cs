using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxColumnCrossingTests
{
    public static void PptxColumnCategoryGeometryIgnoresNumericCrossingEdits()
    {
        foreach (int minimum in new[] { 0, 7 })
        foreach (bool reversed in new[] { false, true })
        foreach (string crossing in new[] { "autoZero", "max", "interior", "upper" })
        foreach (string labels in new[] { "nextTo", "low", "high" })
        {
            var left = Geometry(Render("autoZero", crossing, reversed, labels, minimum));
            var right = Geometry(Render("max", crossing, reversed, labels, minimum));
            TestAssert.True(left.Zip(right).All(p => Math.Abs(p.First - p.Second) < .002d),
                $"Numeric crossings own X; column category Y/strip/captions retain {crossing}, {labels}, reversal {reversed}, minimum {minimum}.");
        }
    }

    public static void PptxColumnCategorySourceModesKeepVisibleCaptionBands()
    {
        var maximum = Geometry(Render("autoZero", "max"));
        var numeric = Geometry(Render("autoZero", "upper"));
        TestAssert.True(Math.Abs(maximum[2] - maximum[1]) < .002d && maximum[3] > maximum[1],
            "An implicit maximum puts the category stroke at the upper plot edge and admits captions above it.");
        TestAssert.True(Math.Abs(numeric[2] - numeric[1]) < .002d && numeric[3] < numeric[1],
            "A numeric crossing at the same value keeps next-to captions below its stroke.");
        var reversed = Geometry(Render("autoZero", "upper", reversed: true));
        TestAssert.True(Math.Abs(reversed[2] - reversed[0]) < .002d && reversed[3] > reversed[1],
            "Reversed upper numeric crossings keep their stroke while admitting captions into the existing low-side band.");
    }

    public static void PptxColumnDataBaselinesStayIndependentOfCategoryCrossings()
    {
        foreach (string crossing in new[] { "autoZero", "interior", "max" })
        {
            PdfPage page = Render("autoZero", crossing);
            var plot = Geometry(page);
            var bars = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re f")
                .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
                .Where(r => r[2] > 20d && r[2] < 200d && r[3] > 20d).ToArray();
            TestAssert.Equal(3, bars.Length);
            TestAssert.True(bars.All(r => Math.Abs(r[1] - plot[0]) < .002d),
                "Positive columns retain their zero baseline, independently of the category stroke.");
            double[] ratios = bars.Select(r => r[3] / (plot[1] - plot[0])).Order().ToArray();
            TestAssert.True(ratios.Zip(new[] { .2d, .5d, .8d }).All(p => Math.Abs(p.First - p.Second) < .001d),
                "Category crossings preserve the source column lengths.");
        }
    }

    private static double[] Geometry(PdfPage page)
    {
        var paths = new List<string>();
        var colors = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add, strokes: true).ToArray();
        var lines = colors.Zip(paths).SelectMany(p => Regex.Matches(p.Second, @"([-\d.]+) ([-\d.]+) m ([-\d.]+) ([-\d.]+) l S")
            .Select(m => (p.First, V: m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray()))).ToArray();
        double[] axis = lines.Single(p => p.First == new RgbColor(255, 0, 0) && Math.Abs(p.V[0] - p.V[2]) < .001d && Math.Abs(p.V[1] - p.V[3]) > 100d).V;
        double[] category = lines.Single(p => p.First == new RgbColor(0, 170, 0) && Math.Abs(p.V[1] - p.V[3]) < .001d && Math.Abs(p.V[0] - p.V[2]) > 200d).V;
        var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
        TestAssert.True(labels.Select(l => l.Text).SequenceEqual(new[] { "Alfa", "Beta", "Gamma" }), "Keep all source captions.");
        TestAssert.True(labels.All(l => l.Y >= 72d && l.Y <= 504d), "Admit caption anchors inside the chart frame.");
        return [Math.Min(axis[1], axis[3]), Math.Max(axis[1], axis[3]), category[1], labels[0].Y];
    }

    private static PdfPage Render(string numericCrossing, string categoryCrossing, bool reversed = false, string labels = "nextTo", int minimum = 0)
    {
        double maximum = minimum == 0 ? 5d : 103d;
        string N(double v) => v.ToString("R", CultureInfo.InvariantCulture);
        string cross = categoryCrossing is "interior" or "upper"
            ? "<c:crossesAt val=\"" + N(categoryCrossing == "upper" ? maximum : minimum + (maximum - minimum) * .6d) + "\"/>"
            : "<c:crosses val=\"" + categoryCrossing + "\"/>";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea>
              <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser>
                <c:cat><c:strLit>{{string.Concat(new[] { "Alfa", "Beta", "Gamma" }.Select((s,i) => $"<c:pt idx=\"{i}\"><c:v>{s}</c:v></c:pt>"))}}</c:strLit></c:cat>
                <c:val><c:numLit>{{string.Concat(new[] { .2d, .5d, .8d }.Select((v,i) => $"<c:pt idx=\"{i}\"><c:v>{N(minimum + (maximum - minimum) * v)}</c:v></c:pt>"))}}</c:numLit></c:val>
              </c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart>
              <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="{{labels}}"/><c:crossAx val="2"/>{{cross}}
                <c:spPr><a:ln><a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></a:ln></c:spPr></c:catAx>
              <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/><c:min val="{{minimum}}"/><c:max val="{{N(maximum)}}"/></c:scaling>
                <c:axPos val="l"/><c:crossAx val="1"/><c:crosses val="{{numericCrossing}}"/>
                <c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:valAx>
            </c:plotArea></c:chart></c:chartSpace>
            """;
        const string slide = """
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="457200"/><a:ext cx="9144000" cy="5486400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart, new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }
}
