using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxCartesianHeaderTests
{
    public static void PptxCartesianHeadersKeepTopAcrossFrameHeights()
    {
        foreach (string kind in new[] { "column", "line" })
        foreach (bool title in new[] { false, true })
        foreach (int size in new[] { 8, 16 })
        foreach (bool legend in new[] { false, true })
        {
            double? reference = null;
            foreach (int height in new[] { 180, 300, 420 })
            {
                var actual = Render(kind, height, size, title: title, legend: legend);
                reference ??= actual.Top;
                TestAssert.True(Math.Abs(reference.Value - actual.Top) < 0.002,
                    "The same title and axis strips keep their plot top when only frame height changes.");
                TestAssert.True(actual.Text.Contains("Category", StringComparison.Ordinal) &&
                    (!title || actual.Text.Contains("Supported title", StringComparison.Ordinal)),
                    "Reserving a header must preserve the supported caption and categories.");
            }
        }
    }

    public static void PptxCartesianHeadersComposeCaptionAndVisibleNumericBounds()
    {
        foreach (string kind in new[] { "column", "line" })
        {
            var small = Render(kind, 300, 8);
            var large = Render(kind, 300, 16);
            TestAssert.True(Math.Abs(small.Top - 119d) < 0.002 && large.Top > small.Top + 1,
                "Small labels retain minimum frame clearance; larger numeric labels reserve their own line box.");
            var titleSmall = Render(kind, 300, 12, title: true, titleSize: 12);
            var titleLarge = Render(kind, 300, 12, title: true, titleSize: 24);
            TestAssert.True(titleLarge.Top > titleSmall.Top + 6,
                "Caption height must contribute its measured band above the numeric strip.");
            var hiddenSmall = Render(kind, 300, 8, title: true, valueLabels: false);
            var hiddenLarge = Render(kind, 300, 24, title: true, valueLabels: false);
            TestAssert.True(Math.Abs(hiddenSmall.Top - hiddenLarge.Top) < 0.002,
                "Hidden numeric labels do not reserve a font-dependent header.");
        }
    }

    public static void PptxCartesianHeadersKeepOverlayCaptionsOutOfPlotReservation()
    {
        foreach (string kind in new[] { "column", "line" })
        foreach (int size in new[] { 8, 16 })
        foreach (bool legend in new[] { false, true })
        {
            var plain = Render(kind, 300, size, legend: legend);
            var overlay = Render(kind, 300, size, title: true, legend: legend, overlay: true);
            TestAssert.True(Math.Abs(plain.Top - overlay.Top) < 0.002 &&
                overlay.Text.Contains("Supported title", StringComparison.Ordinal),
                "An overlay caption remains painted without consuming a header band.");
        }
    }

    public static void PptxColumnCategoryStripsKeepTheirFrameInsetAcrossTitlesAndHeights()
    {
        foreach (int size in new[] { 8, 18 })
        foreach (bool visible in new[] { false, true })
        foreach (bool title in new[] { false, true })
        foreach (bool legend in new[] { false, true })
        {
            double? reference = null;
            foreach (int height in new[] { 180, 300, 420 })
            {
                var actual = Render("column", height, 12, title: title, legend: legend,
                    categorySize: size, categoryLabels: visible);
                double inset = 108d + height - actual.Bottom;
                reference ??= inset;
                TestAssert.True(Math.Abs(inset - reference.Value) < 0.002d,
                    "The category strip belongs to its content, independently of frame height, caption and overlay legend.");
                TestAssert.True(!visible || actual.Text.Contains("Category 6", StringComparison.Ordinal),
                    "All supported categories remain painted after footer admission.");
            }
        }
    }

    public static void PptxColumnCategoryVisibilityOwnsItsFooterClearance()
    {
        double Inset(int categorySize, bool visible, int valueSize = 12, bool valueLabels = true)
        {
            var result = Render("column", 300, valueSize, title: true, legend: true,
                categorySize: categorySize, categoryLabels: visible, valueLabels: valueLabels);
            return 408d - result.Bottom;
        }
        double small = Inset(8, true), large = Inset(18, true);
        double hiddenSmall = Inset(8, false), hiddenLarge = Inset(24, false);
        TestAssert.True(large > small + 10d && hiddenSmall < small,
            "Visible categories own their occupied band; hidden categories do not retain that band.");
        TestAssert.True(Math.Abs(hiddenSmall - hiddenLarge) < 0.002d,
            "A hidden category font size must not affect footer admission.");
        TestAssert.True(Inset(8, false, valueSize: 24) > hiddenSmall + 1d,
            "The numeric end label still owns its clearance when category labels are hidden.");
        TestAssert.True(Inset(8, false, valueLabels: false) <= hiddenSmall,
            "Hidden numeric labels do not retain a font-dependent end clearance.");
    }

    private static (double Top, double Bottom, string Text) Render(string kind, int height, int valueSize,
        bool title = false, bool legend = false, int titleSize = 18, bool overlay = false, bool valueLabels = true,
        int? categorySize = null, bool categoryLabels = true)
    {
        string titleXml = title ? $"<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz=\"{titleSize * 100}\"/><a:t>Supported title</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"{(overlay ? 1 : 0)}\"/></c:title>" : "";
        string values = string.Concat(Enumerable.Range(0, 6).Select(i => $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
        string categories = string.Concat(Enumerable.Range(0, 6).Select(i => $"<c:pt idx=\"{i}\"><c:v>Category {i + 1}</c:v></c:pt>"));
        string series = $"<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx><c:v>Series</c:v></c:tx><c:cat><c:strLit>{categories}</c:strLit></c:cat><c:val><c:numLit>{values}</c:numLit></c:val></c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/>";
        string plot = kind == "column" ? "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" + series + "</c:barChart>"
            : "<c:lineChart><c:grouping val=\"standard\"/>" + series + "</c:lineChart>";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>{{titleXml}}<c:autoTitleDeleted val="{{(title ? 0 : 1)}}"/><c:plotArea>{{plot}}
                <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/>
                  {{(categoryLabels ? "" : "<c:tickLblPos val=\"none\"/>")}}
                  {{(categorySize is { } size ? "<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"" + size * 100 + "\"/></a:pPr></a:p></c:txPr>" : "")}}
                </c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/>
                  <c:majorGridlines/><c:majorUnit val="1"/><c:tickLblPos val="{{(valueLabels ? "nextTo" : "none")}}"/>
                  <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{valueSize * 100}}"/></a:pPr></a:p></c:txPr>
                </c:valAx>
              </c:plotArea>{{(legend ? "<c:legend><c:legendPos val=\"b\"/></c:legend>" : "")}}</c:chart>
            </c:chartSpace>
            """;
        string slide = $$"""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="3657600" cy="{{height * 12700}}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            </p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart,
            new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
        var clips = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re W\* n")
            .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(v => v[2] > 144 && v[2] < 287.999 && v[3] > height / 4d).ToArray();
        TestAssert.True(clips.Length > 0, "The rendered chart retains a bounded plot clip.");
        return (page.Height - clips.Max(v => v[1] + v[3]), page.Height - clips.Min(v => v[1]),
            string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(s => s.Text)));
    }
}
