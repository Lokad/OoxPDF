using System.Globalization;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxTitleShapeTests
{
    private const string Fill = "<c:spPr><a:solidFill><a:srgbClr val=\"FFF2CC\"/></a:solidFill></c:spPr>";
    private const string Outline = "<c:spPr><a:noFill/><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"9C6500\"/></a:solidFill></a:ln></c:spPr>";
    private static readonly RgbColor Yellow = new(255, 242, 204);
    private static readonly RgbColor Brown = new(156, 101, 0);
    private static readonly RgbColor Blue = new(15, 158, 213);

    public static void PptxAutomaticTitleShapesUseCaptionBoundsAcrossChartKinds()
    {
        foreach (string kind in new[] { "bar", "column", "line", "pie" })
        foreach (bool outline in new[] { false, true })
        {
            PdfPage page = Render(kind, shape: outline ? Outline : Fill);
            var box = Rectangles(page, outline ? Brown : Yellow, outline).Single();
            TestAssert.True(box.Width < 144 && box.Height < 72,
                "An automatic title shape must surround its caption, not the chart frame.");
            var title = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
                .Single(show => show.Text == "Supported title");
            TestAssert.True(title.X >= box.X && title.X <= box.X + box.Width &&
                title.Y >= box.Y && title.Y <= box.Y + box.Height,
                "The title shape and painted caption must use the same position.");
            if (kind is "bar" or "column")
            {
                var bars = Rectangles(page, Blue).ToArray();
                TestAssert.Equal(6, bars.Length);
                TestAssert.True(bars.All(bar => !Contains(box, bar.X + bar.Width / 2, bar.Y + bar.Height / 2)),
                    "An opaque automatic title must not cover the plotted bar centers.");
            }
        }
    }

    public static void PptxAutomaticTitleShapesFollowRichRunsAndInferredCaptions()
    {
        var whole = Rectangles(Render("bar"), Yellow).Single();
        var split = Rectangles(Render("bar", Run("Supported", 18) + Run(" title", 18)), Yellow).Single();
        TestAssert.True(Delta(whole, split) < 0.001,
            "Splitting an unchanged run must preserve its automatic title rectangle.");
        var large = Rectangles(Render("bar", Run("Supported title", 24)), Yellow).Single();
        var mixed = Rectangles(Render("bar", Run("Supported", 12) + Run(" title", 24)), Yellow).Single();
        TestAssert.True(Math.Abs(large.Height - mixed.Height) < 0.001 && mixed.Width < large.Width,
            "Mixed title sizes must use their measured height and advance.");
        PdfPage inferred = Render("bar", inferred: true);
        var inferredBox = Rectangles(inferred, Yellow).Single();
        TestAssert.True(inferredBox.Width < 144 && inferredBox.Height < 72 &&
            DocxCommentsTests.ReadEmbeddedGlyphTextShows(inferred).Any(show => show.Text == "Series"),
            "An inferred caption must be measured rather than inherit its text viewport width.");
    }

    public static void PptxAutomaticTitleShapesPreserveLinkedAndOverlayText()
    {
        foreach (string kind in new[] { "bar", "column", "line", "pie" })
        {
            PdfPage plain = Render(kind, shape: "", linked: true, overlay: true);
            PdfPage styled = Render(kind, linked: true, overlay: true);
            var box = Rectangles(styled, Yellow).Single();
            TestAssert.True(box.Width < 144 && box.Height < 72, "Overlay title shapes also use caption bounds.");
            var before = DocxCommentsTests.ReadEmbeddedGlyphTextShows(plain).Select(show => (show.Text, show.Size, show.X, show.Y));
            var after = DocxCommentsTests.ReadEmbeddedGlyphTextShows(styled).Select(show => (show.Text, show.Size, show.X, show.Y));
            TestAssert.True(before.SequenceEqual(after), "Title styling must preserve painted text and its placement.");
            TestAssert.Equal(1, styled.Annotations.Count);
            TestAssert.True(plain.Annotations.SequenceEqual(styled.Annotations),
                "The complete linked caption must retain its existing frame clipping and annotation.");
            TestAssert.Equal("https://example.invalid/title", styled.Annotations.Single().Uri);
        }
    }

    public static void PptxManualTitleShapesKeepTheirAuthoredRectangle()
    {
        PdfPage page = Render("bar", manual: true);
        var box = Rectangles(page, Yellow).Single();
        TestAssert.True(Math.Abs(box.X - 144) < 0.001 && Math.Abs(box.Width - 144) < 0.001 &&
            Math.Abs(box.Height - 43.2) < 0.001, "A manual title retains its source rectangle.");
    }

    private static string Run(string text, int size) => $"<a:r><a:rPr sz=\"{size * 100}\"/><a:t>{text}</a:t></a:r>";

    private static PdfPage Render(string kind, string? titleRuns = null, string shape = Fill,
        bool inferred = false, bool linked = false, bool manual = false, bool overlay = false)
    {
        titleRuns ??= Run("Supported title", 18);
        if (linked) titleRuns = titleRuns.Replace("/>", "><a:hlinkClick r:id=\"titleLink\"/></a:rPr>", StringComparison.Ordinal);
        string titleText = inferred ? "" : "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>" + titleRuns + "</a:p></c:rich></c:tx>";
        string layout = manual ? "<c:layout><c:manualLayout><c:xMode val=\"factor\"/><c:yMode val=\"factor\"/><c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\"0.25\"/><c:y val=\"0.5\"/><c:w val=\"0.5\"/><c:h val=\"0.2\"/></c:manualLayout></c:layout>" : "";
        string values = string.Concat(Enumerable.Range(0, 6).Select(i => $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
        string series = "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx><c:v>Series</c:v></c:tx>" +
            "<c:spPr><a:solidFill><a:srgbClr val=\"0F9ED5\"/></a:solidFill></c:spPr><c:val><c:numLit><c:ptCount val=\"6\"/>" + values + "</c:numLit></c:val></c:ser>";
        string axisIds = "<c:axId val=\"1\"/><c:axId val=\"2\"/>";
        string plot = kind switch
        {
            "line" => "<c:lineChart><c:grouping val=\"standard\"/>" + series + axisIds + "</c:lineChart>",
            "pie" => "<c:pieChart>" + series + "</c:pieChart>",
            _ => "<c:barChart><c:barDir val=\"" + (kind == "bar" ? "bar" : "col") + "\"/><c:grouping val=\"clustered\"/>" + series + axisIds + "</c:barChart>"
        };
        string axes = "<c:catAx><c:axId val=\"1\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"none\"/><c:crossAx val=\"2\"/></c:catAx>" +
            "<c:valAx><c:axId val=\"2\"/><c:scaling><c:min val=\"0\"/><c:max val=\"5\"/></c:scaling><c:axPos val=\"l\"/><c:tickLblPos val=\"none\"/><c:crossAx val=\"1\"/></c:valAx>";
        string chart = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<c:style val=\"2\"/><c:chart><c:title>" + titleText + layout + shape + "<c:overlay val=\"" + (overlay ? 1 : 0) + "\"/></c:title><c:autoTitleDeleted val=\"0\"/><c:plotArea>" + plot +
            (kind == "pie" ? "" : axes) + "</c:plotArea></c:chart></c:chartSpace>";
        var parts = new Dictionary<string, byte[]>();
        if (linked) parts["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="titleLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/title" TargetMode="External"/></Relationships>
            """);
        using var stream = PptxTests.CreateSingleChartPackageStream(chart, parts);
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }

    private static IEnumerable<(double X, double Y, double Width, double Height)> Rectangles(PdfPage page, RgbColor selected, bool stroke = false)
    {
        var paths = new List<string>();
        var colors = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add, stroke).ToArray();
        for (int i = 0; i < colors.Length; i++)
        {
            int rectangle = paths[i].IndexOf(" re", StringComparison.Ordinal);
            if (colors[i] != selected || rectangle < 0) continue;
            double[] values = paths[i][..rectangle].Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
            yield return (values[0], values[1], values[2], values[3]);
        }
    }

    private static bool Contains((double X, double Y, double Width, double Height) box, double x, double y) =>
        x >= box.X && x <= box.X + box.Width && y >= box.Y && y <= box.Y + box.Height;

    private static double Delta((double X, double Y, double Width, double Height) a,
        (double X, double Y, double Width, double Height) b) =>
        new[] { Math.Abs(a.X - b.X), Math.Abs(a.Y - b.Y), Math.Abs(a.Width - b.Width), Math.Abs(a.Height - b.Height) }.Max();
}
