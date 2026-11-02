using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxCartesianSideTests
{
    public static void PptxInnerFactorPlotOriginsMatchIndependentOfficeControls()
    {
        string directory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "ManualInnerPlots"));
        using var expected = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "office-origins.json")));
        foreach (var row in expected.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string input = Path.Combine(directory, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString()!, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))));
            using var package = OoxPackage.Open(input, default);
            var document = new PptxReader().Read(package, default);
            var page = new PptxRenderer(new TestFaceFontResolver())
                .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
            bool horizontal = row.GetProperty("Horizontal").GetBoolean();
            var lines = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) m\s+([-\d.]+) ([-\d.]+) l(?: S)?")
                .Select(match => match.Groups.Cast<Group>().Skip(1).Select(group => double.Parse(group.Value, CultureInfo.InvariantCulture)).ToArray())
                .Where(v => horizontal
                    ? Math.Abs(v[1] - v[3]) < .001d && Math.Abs(v[0] - v[2]) > 100d
                    : Math.Abs(v[0] - v[2]) < .001d && Math.Abs(v[1] - v[3]) > 100d).ToArray();
            TestAssert.Equal(1, lines.Length);
            var line = lines[0];
            double[] actual = horizontal
                ? [page.Height - line[1], Math.Min(line[0], line[2]), Math.Max(line[0], line[2])]
                : [line[0], page.Height - Math.Max(line[1], line[3]), page.Height - Math.Min(line[1], line[3])];
            double[] office = row.GetProperty("OfficeAxis").EnumerateArray().Select(value => value.GetDouble()).ToArray();
            for (int i = 0; i < actual.Length; i++)
                TestAssert.True(Math.Abs(actual[i] - office[i]) < .002d,
                    $"{Path.GetFileName(input)} axis coordinate {i}: {actual[i]} differs from Office {office[i]}.");
        }
    }

    public static void PptxPrimaryValueAxisLabelsPaintOnceOnTheirSelectedSide()
    {
        foreach (bool right in new[] { false, true })
        foreach (int width in new[] { 288, 576 })
        {
            var page = Render(width, 12, right: right);
            var plot = Grid(page);
            var numbers = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
                .Where(show => Regex.IsMatch(show.Text, @"^\d+$")).ToArray();
            TestAssert.Equal(6, numbers.Length);
            TestAssert.True(numbers.Select(show => show.Text).Order().SequenceEqual(new[] { "0", "1", "2", "3", "4", "5" }),
                "One authored value axis paints exactly one copy of every tick label.");
            TestAssert.True(numbers.All(show => right ? show.X > plot.Right : show.X < plot.Left),
                "Primary labels follow their own side instead of a hardcoded left-side default.");
        }
    }

    public static void PptxSecondaryValueAxisOwnsItsVisibilityIndependently()
    {
        foreach (bool primaryVisible in new[] { false, true })
        foreach (bool secondaryVisible in new[] { false, true })
        {
            var page = Render(576, 12, visible: primaryVisible, secondary: true, secondaryVisible: secondaryVisible);
            var plot = Grid(page);
            var numbers = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
                .Where(show => Regex.IsMatch(show.Text, @"^\d+$")).ToArray();
            var left = numbers.Where(show => show.X < plot.Left).Select(show => show.Text).Order().ToArray();
            var right = numbers.Where(show => show.X > plot.Right).Select(show => show.Text).Order().ToArray();
            TestAssert.True(left.SequenceEqual(primaryVisible ? new[] { "0", "1", "2", "3", "4", "5" } : Array.Empty<string>()),
                "The primary axis controls only its own labels.");
            TestAssert.True(right.SequenceEqual(secondaryVisible ? new[] { "0", "10", "2", "4", "6", "8" } : Array.Empty<string>()),
                "A genuine secondary axis survives hidden primary labels and respects its own visibility.");
        }
    }

    public static void PptxLineNumericSideStripsFollowLabelsAcrossFrameWidths()
    {
        foreach (bool right in new[] { false, true })
        foreach (int size in new[] { 8, 24 })
        {
            (double Left, double Right, double Bottom, double Top)? reference = null;
            foreach (int width in new[] { 288, 576, 720 })
            {
                var page = Render(width, size, right: right, furniture: true);
                var plot = Grid(page);
                var actual = (Left: plot.Left - 72d, Right: 72d + width - plot.Right, plot.Bottom, plot.Top);
                reference ??= actual;
                TestAssert.True(Math.Abs(actual.Left - reference.Value.Left) < 0.002d &&
                    Math.Abs(actual.Right - reference.Value.Right) < 0.002d,
                    "The measured numeric strip and opposite margin remain attached to their frame sides.");
                TestAssert.True(Math.Abs(actual.Bottom - reference.Value.Bottom) < 0.002d &&
                    Math.Abs(actual.Top - reference.Value.Top) < 0.002d,
                    "Horizontal measurement preserves the existing vertical placement.");
                string text = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text));
                TestAssert.Contains("Supported title", text);
                TestAssert.Contains("Series", text);
            }
        }
        foreach (bool right in new[] { false, true })
        foreach (int size in new[] { 8, 24 })
        {
            var hidden = Grid(Render(576, size, right: right, visible: false, furniture: true));
            TestAssert.True(Math.Abs(hidden.Left - 83d) < 0.002d && Math.Abs(hidden.Right - 637d) < 0.002d,
                "Hidden numeric labels consume no font-dependent side strip; Office retains its 11pt margins.");
        }
    }

    public static void PptxManualLinePlotRetainsItsAuthoredGeometryAcrossFonts()
    {
        var small = Grid(Render(576, 8, manual: true));
        var large = Grid(Render(576, 24, manual: true));
        TestAssert.Equal(small, large);
    }

    public static void PptxLineCategoryFooterKeepsItsInsetAcrossFrameHeights()
    {
        foreach (bool right in new[] { false, true })
        foreach (bool visible in new[] { false, true })
        {
            double? inset = null;
            foreach (int height in new[] { 180, 300, 420 })
            {
                var page = Render(576, 12, right: right, visible: visible, furniture: true, height: height);
                double actual = Grid(page).Bottom - (page.Height - 108d - height);
                inset ??= actual;
                TestAssert.True(Math.Abs(actual - inset.Value) < 0.002d,
                    "The category strip keeps its measured frame inset when captions and overlay legends coexist.");
                TestAssert.Contains("Series", string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).Select(show => show.Text)));
            }
        }
    }

    private static (double Left, double Right, double Bottom, double Top) Grid(PdfPage page)
    {
        var lines = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) m ([-\d.]+) ([-\d.]+) l S")
            .Select(match => match.Groups.Cast<Group>().Skip(1).Select(group => double.Parse(group.Value, CultureInfo.InvariantCulture)).ToArray())
            .Where(v => Math.Abs(v[1] - v[3]) < 0.001d && Math.Abs(v[0] - v[2]) > 100d).ToArray();
        TestAssert.True(lines.Length > 0, "The chart retains its horizontal grid geometry.");
        return (lines.Min(v => Math.Min(v[0], v[2])), lines.Max(v => Math.Max(v[0], v[2])), lines.Min(v => v[1]), lines.Max(v => v[1]));
    }

    private static PdfPage Render(int width, int size, bool right = false, bool visible = true,
        bool secondary = false, bool secondaryVisible = true, bool furniture = false, bool manual = false, int height = 300)
    {
        const string series = "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:tx><c:v>Series</c:v></c:tx><c:cat><c:strLit><c:pt idx=\"0\"><c:v>Q1</c:v></c:pt><c:pt idx=\"1\"><c:v>Q2</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx=\"0\"><c:v>3</c:v></c:pt><c:pt idx=\"1\"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>";
        string additional = secondary ? "<c:lineChart>" + series.Replace("val=\"0\"", "val=\"1\"") + "<c:axId val=\"4\"/><c:axId val=\"3\"/></c:lineChart>" : "";
        string extraAxis = secondary ? $"<c:catAx><c:axId val=\"4\"/><c:delete val=\"1\"/><c:axPos val=\"b\"/><c:crossAx val=\"3\"/></c:catAx><c:valAx><c:axId val=\"3\"/><c:scaling><c:min val=\"0\"/><c:max val=\"10\"/></c:scaling><c:axPos val=\"r\"/><c:crossAx val=\"4\"/><c:majorUnit val=\"2\"/><c:tickLblPos val=\"{(secondaryVisible ? "nextTo" : "none")}\"/></c:valAx>" : "";
        string title = furniture ? "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Supported title</a:t></a:r></a:p></c:rich></c:tx></c:title>" : "";
        string layout = manual ? "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:x val=\"0.25\"/><c:y val=\"0.25\"/><c:w val=\"0.5\"/><c:h val=\"0.5\"/></c:manualLayout></c:layout>" : "";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>{{title}}<c:autoTitleDeleted val="{{(furniture ? 0 : 1)}}"/><c:plotArea>{{layout}}
                <c:lineChart><c:varyColors val="0"/>{{series}}<c:axId val="1"/><c:axId val="2"/></c:lineChart>{{additional}}
                <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="{{(right ? "r" : "l")}}"/><c:crossAx val="1"/>
                  <c:crosses val="{{(right ? "max" : "autoZero")}}"/><c:majorGridlines/><c:majorUnit val="1"/><c:tickLblPos val="{{(visible ? "nextTo" : "none")}}"/>
                  <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{size * 100}}"/></a:pPr></a:p></c:txPr>
                </c:valAx>{{extraAxis}}
              </c:plotArea>{{(furniture ? "<c:legend><c:legendPos val=\"b\"/></c:legend>" : "")}}</c:chart>
            </c:chartSpace>
            """;
        string slide = $$"""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="{{width * 12700}}" cy="{{height * 12700}}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            </p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart,
            new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, null, default).ToBlockingEnumerable().Single();
    }
}
