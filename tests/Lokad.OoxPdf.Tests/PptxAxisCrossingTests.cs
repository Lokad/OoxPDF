using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxAxisCrossingTests
{
    public static void PptxNumericCrossingOwnsPairedCategorySide()
    {
        // Independent Office matrix: numeric crosses owns the stroke/next-to
        // side; category reversal mirrors it, while low/high labels stay fixed.
        foreach (string kind in new[] { "line", "area", "column" })
        foreach (bool reversed in new[] { false, true })
        foreach (string? crossing in new string?[] { null, "autoZero", "min", "max" })
        foreach (string labels in new[] { "nextTo", "low", "high" })
        {
            PdfPage left = Render(kind, "l", crossing, reversed, labels);
            PdfPage right = Render(kind, "r", crossing, reversed, labels);
            TestAssert.True(left.Content == right.Content,
                $"Paired {kind} crossing {crossing ?? "default"}, reversal {reversed}, labels {labels} must survive an axPos-only edit.");
            Check(left, (crossing == "max") != reversed, labels);
        }
    }

    public static void PptxNumericCrossingRetainsUnpairedAndInteriorFallbacks()
    {
        foreach (string fallback in new[] { "unpaired", "date", "interior", "unknown" })
        foreach (string position in new[] { "l", "r" })
        {
            PdfPage page = Render("line", position, "max", false, "nextTo", fallback);
            Check(page, position == "r", "nextTo");
        }
    }

    private static void Check(PdfPage page, bool axisRight, string labels)
    {
        var clips = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re W\* n")
            .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(c => c[2] > 200d && c[2] < 700d && c[3] > 100d).OrderBy(c => c[2] * c[3]).ToArray();
        TestAssert.True(clips.Length > 0, "Inspect the actual data plot clip.");
        double left = clips[0][0], right = left + clips[0][2];
        var paths = new List<string>();
        RgbColor[] colors = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add, strokes: true).ToArray();
        var axes = colors.Zip(paths).Where(p => p.First == new RgbColor(255, 0, 0))
            .SelectMany(p => Regex.Matches(p.Second, @"([-\d.]+) ([-\d.]+) m ([-\d.]+) ([-\d.]+) l S"))
            .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(c => Math.Abs(c[0] - c[2]) < .001d && Math.Abs(c[1] - c[3]) > 100d).ToArray();
        TestAssert.Equal(1, axes.Length);
        TestAssert.True(Math.Abs(axes[0][0] - (axisRight ? right : left)) < .002d,
            "The colored physical numeric axis follows the paired categorical crossing.");
        var values = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).ToArray();
        TestAssert.True(values.Length >= 2, "Keep the numeric labels.");
        bool labelsRight = labels == "high" || labels == "nextTo" && axisRight;
        TestAssert.True(values.All(v => labelsRight ? v.X > right : v.X < left),
            "Measurement and painting share the numeric side; explicit low/high labels remain independent.");
    }

    private static PdfPage Render(string kind, string position, string? crossing, bool reversed, string labels, string? fallback = null)
    {
        string plot = kind == "column" ? "bar" : kind;
        string category = fallback == "date" ? "dateAx" : "catAx";
        string sourceCrossing = fallback == "interior" ? "<c:crossesAt val=\"3\"/>" :
            fallback == "unknown" ? "<c:crosses val=\"future\"/>" : crossing is null ? "" : $"<c:crosses val=\"{crossing}\"/>";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:{{plot}}Chart>
                {{(kind == "column" ? "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" : "")}}
                <c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>Alfa</c:v></c:pt><c:pt idx="1"><c:v>Beta</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
                <c:axId val="1"/><c:axId val="2"/></c:{{plot}}Chart>
                <c:{{category}}><c:axId val="1"/><c:scaling><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/></c:scaling>
                  <c:axPos val="b"/><c:crossAx val="{{(fallback == "unpaired" ? "other" : "2")}}"/><c:crosses val="autoZero"/></c:{{category}}>
                <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="{{position}}"/>
                  <c:tickLblPos val="{{labels}}"/><c:crossAx val="1"/>{{sourceCrossing}}
                  <c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
                </c:valAx></c:plotArea></c:chart>
            </c:chartSpace>
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
