using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxAxisCrossingTests
{
    public static void PptxLineAndAreaCategoriesSharePhysicalSourceDirection()
    {
        const string paint = "<c:spPr><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill></a:ln></c:spPr><c:smooth val=\"0\"/>";
        foreach (string kind in new[] { "line", "area" })
        foreach (bool reversed in new[] { false, true })
        {
            PdfPage page = Render(kind, "l", "autoZero", reversed, "low", seriesFormat: paint);
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.True(captions.Select(c => c.Text).SequenceEqual(new[] { "Alfa", "Beta" }), "Keep logical caption order and content.");
            TestAssert.True(reversed ? captions[0].X > captions[1].X : captions[0].X < captions[1].X,
                "The shared source direction determines caption placement, independently of the numeric label side.");
            var blue = Regex.Match(page.Content, @"0 0 1 r[Gg](?<body>.*?)\nQ", RegexOptions.Singleline);
            var points = Regex.Matches(blue.Groups["body"].Value, @"([-\d.]+) ([-\d.]+) [ml]")
                .Select(m => (X: double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture), Y: double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture))).ToArray();
            TestAssert.True(points.Length >= 2 && (reversed ? points[0].X > points[1].X : points[0].X < points[1].X) && points[0].Y < points[1].Y,
                $"{kind}, reversed {reversed}, blue block {blue.Success}, points {string.Join(';', points)}: source values2 and4 keep their identities while physical X coordinates mirror.");
        }
    }

    public static void PptxLineCategoryDirectionKeepsSparseLabelsAndPointPaint()
    {
        const string paint = "<c:marker><c:symbol val=\"square\"/><c:size val=\"7\"/></c:marker><c:smooth val=\"0\"/><c:dPt><c:idx val=\"3\"/><c:spPr><a:ln><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln></c:spPr></c:dPt>";
        const string labels = "<c:varyColors val=\"1\"/><c:dLbls><c:dLblPos val=\"t\"/><c:showVal val=\"0\"/><c:showCatName val=\"1\"/></c:dLbls>";
        RgbColor[]? previous = null;
        foreach (bool reversed in new[] { false, true })
        {
            PdfPage page = Render("line", "l", "autoZero", reversed, "low", seriesFormat: paint, plotFormat: labels, sparse: true);
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            var data = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CLD", StringComparison.Ordinal)).ToArray();
            TestAssert.True(captions.Select(c => c.Text).SequenceEqual(new[] { "Alfa", "Beta", "Gamma", "Delta" }), "Sparse values do not remove category slots.");
            TestAssert.True(data.Select(c => c.Text).SequenceEqual(new[] { "Alfa", "Gamma", "Delta" }), "Retain original indices around the missing second value.");
            TestAssert.True(reversed ? data[0].X > data[^1].X : data[0].X < data[^1].X, "Data labels follow the same physical direction as their source points.");
            RgbColor[] colors = PptxChartPaletteTests.ReadColoredPaints(page.Content).ToArray();
            TestAssert.True(colors.Distinct().Count() >= 3, "The three source points retain their varied marker paint around the missing value.");
            TestAssert.True(PptxChartPaletteTests.ReadColoredPaints(page.Content, strokes: true).Count(c => c == new RgbColor(255, 0, 0)) >= 2,
                "The authored incoming stroke at logical index3 remains visible alongside the red numeric axis.");
            if (previous is not null) TestAssert.True(previous.SequenceEqual(colors), "Mirroring coordinates preserves logical point paint and marker order.");
            previous = colors;
        }
    }

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
            Check(left, (crossing == "max") != reversed, labels, $"{kind}/{crossing ?? "default"}/{reversed}/{labels}");
        }
    }

    public static void PptxNumericCrossingRetainsUnpairedAndInteriorFallbacks()
    {
        foreach (string fallback in new[] { "unpaired", "date", "interior", "unknown" })
        foreach (string position in new[] { "l", "r" })
        {
            PdfPage page = Render("line", position, "max", false, "nextTo", fallback);
            Check(page, position == "r", "nextTo", fallback + "/" + position);
        }
    }

    public static void PptxNumericCrossingMovesTheMeasuredStripWithItsLabels()
    {
        foreach (string kind in new[] { "line", "area", "column" })
        {
            var left = Span(Render(kind, "l", "autoZero", false, "nextTo"));
            var right = Span(Render(kind, "l", "max", false, "nextTo"));
            TestAssert.True(left.Left > right.Left + 1d && Math.Abs((left.Right - left.Left) - (right.Right - right.Left)) < .002d &&
                Math.Abs(left.Left + right.Right - 864d) < .002d,
                kind + ": move the same measured strip across the720pt frame at X72, alongside the numeric labels.");
            var low = Span(Render(kind, "l", "max", false, "low"));
            TestAssert.True(Math.Abs(low.Left - left.Left) < .002d && Math.Abs(low.Right - left.Right) < .002d,
                "An explicit low label side keeps the measured strip when the physical axis moves right.");
        }
    }

    private static (double Left, double Right) Span(PdfPage page)
    {
        var horizontal = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) m ([-\d.]+) ([-\d.]+) l S")
            .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(c => Math.Abs(c[1] - c[3]) < .001d && Math.Abs(c[0] - c[2]) > 200d).ToArray();
        TestAssert.True(horizontal.Length > 0, "Inspect the painted plot axis/grid span; a column clip also includes stroke padding.");
        return horizontal.GroupBy(c => (Left: Math.Min(c[0], c[2]), Right: Math.Max(c[0], c[2])))
            .OrderByDescending(group => group.Count()).First().Key;
    }

    private static void Check(PdfPage page, bool axisRight, string labels, string context)
    {
        var span = Span(page);
        double left = span.Left, right = span.Right;
        var paths = new List<string>();
        RgbColor[] colors = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add, strokes: true).ToArray();
        var axes = colors.Zip(paths).Where(p => p.First == new RgbColor(255, 0, 0))
            .SelectMany(p => Regex.Matches(p.Second, @"([-\d.]+) ([-\d.]+) m ([-\d.]+) ([-\d.]+) l S"))
            .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(c => Math.Abs(c[0] - c[2]) < .001d && Math.Abs(c[1] - c[3]) > 100d).ToArray();
        TestAssert.Equal(1, axes.Length);
        TestAssert.True(Math.Abs(axes[0][0] - (axisRight ? right : left)) < .002d,
            $"{context}: colored numeric axis X={axes[0][0]}, plot {left}..{right}, expected right={axisRight}.");
        var values = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).ToArray();
        TestAssert.True(values.Length >= 2, "Keep the numeric labels.");
        bool labelsRight = labels == "high" || labels == "nextTo" && axisRight;
        TestAssert.True(values.All(v => labelsRight ? v.X > right : v.X < left),
            "Measurement and painting share the numeric side; explicit low/high labels remain independent.");
    }

    private static PdfPage Render(string kind, string position, string? crossing, bool reversed, string labels, string? fallback = null, string? seriesFormat = null, string? plotFormat = null, bool sparse = false)
    {
        string plot = kind == "column" ? "bar" : kind;
        string category = fallback == "date" ? "dateAx" : "catAx";
        string sourceCrossing = fallback == "interior" ? "<c:crossesAt val=\"3\"/>" :
            fallback == "unknown" ? "<c:crosses val=\"future\"/>" : crossing is null ? "" : $"<c:crosses val=\"{crossing}\"/>";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:{{plot}}Chart>
                {{(kind == "column" ? "<c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" : "")}}
                <c:ser>{{seriesFormat}}<c:cat><c:strLit><c:pt idx="0"><c:v>Alfa</c:v></c:pt><c:pt idx="1"><c:v>Beta</c:v></c:pt>{{(sparse ? "<c:pt idx=\"2\"><c:v>Gamma</c:v></c:pt><c:pt idx=\"3\"><c:v>Delta</c:v></c:pt>" : "")}}</c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt>{{(sparse ? "<c:pt idx=\"2\"><c:v>4</c:v></c:pt><c:pt idx=\"3\"><c:v>1</c:v></c:pt>" : "<c:pt idx=\"1\"><c:v>4</c:v></c:pt>")}}</c:numLit></c:val></c:ser>{{plotFormat}}
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
