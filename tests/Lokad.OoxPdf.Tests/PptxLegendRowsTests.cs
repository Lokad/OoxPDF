using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxLegendRowsTests
{
    public static void PptxBottomLineLegendWrapKeepsNamesInsideTheChartFrame() => CheckBottomLegendRows("line");

    public static void PptxBottomColumnLegendWrapKeepsNamesInsideTheChartFrame() => CheckBottomLegendRows("column");

    public static void PptxCrowdedLegendsPreservePlotSpaceAndDiagnoseUnadmittedRows()
    {
        // Office PDF/API controls establish a row budget within half the chart
        // height. Exercise the resulting placement contract with portable fonts.
        foreach (var (width, height, size, count) in new[]
        {
            (576, 360, 12, 97), (576, 360, 12, 205),
            (288, 240, 8, 97), (840, 420, 18, 205)
        })
        {
            var overlay = Render(width, size, count, overlay: true, kind: "column", height: height, categoryLegend: true);
            var reserved = Render(width, size, count, overlay: false, kind: "column", height: height, categoryLegend: true);
            var overlayBars = ReadBars(overlay.Page, count);
            var reservedBars = ReadBars(reserved.Page, count);
            // All values are4 on an authored0..5 scale; painted height exposes
            // actual plot admission without depending on private layout helpers.
            double reservedHeight = (overlayBars[0][3] - reservedBars[0][3]) / 0.8d;
            TestAssert.True(reservedHeight > 0d,
                $"A reserved legend consumes plot space: {width}x{height}, {size}pt, {count} entries; reserved {reservedHeight}pt.");
            TestAssert.True(reservedBars.All(bar => bar[3] > size),
                "Every source bar retains useful visible height instead of collapsing to a one-point plot.");
            var overlayNames = Names(overlay.Page);
            var reservedNames = Names(reserved.Page);
            TestAssert.True(overlayNames.Length > 0 && overlayNames.Length < count,
                "The automatic frame paints an admitted prefix and diagnoses overflow.");
            double[] baselines = reservedNames.Select(name => name.Y).Distinct().Order().ToArray();
            TestAssert.True(baselines.Length > 1, "These crowded names occupy multiple admitted rows.");
            double rowAdvance = baselines[1] - baselines[0];
            double occupiedRows = baselines[^1] - baselines[0] + rowAdvance;
            // Office caps visible rows; its physical legend box also includes
            // padding/rounding, so total plot reservation need not equal that cap.
            TestAssert.True(occupiedRows <= height / 2d + 0.01d,
                $"Visible legend rows fit their height budget: {width}x{height}, {size}pt; occupied {occupiedRows}pt.");
            TestAssert.True(overlayNames.SequenceEqual(reservedNames),
                "Overlay changes plot reservation while keeping the same admitted legend frames.");
            for (int i = 0; i < reservedNames.Length; i++)
                TestAssert.Equal("Category " + (i + 1).ToString(CultureInfo.InvariantCulture), reservedNames[i].Text);
            TestAssert.True(overlay.Diagnostics.Any(d => d.Message.Contains("overflowing entries", StringComparison.Ordinal)) &&
                reserved.Diagnostics.Any(d => d.Message.Contains("overflowing entries", StringComparison.Ordinal)),
                "Unadmitted source entries must remain explicit through the existing overflow diagnostic.");
        }

        static (string Text, double X, double Y)[] Names(PdfPage page) => DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
            .SelectMany(show => Regex.Matches(show.Text, @"Category \d+")
                .Select(match => (match.Value, show.X, show.Y))).ToArray();

        static double[][] ReadBars(PdfPage page, int count)
        {
            var bars = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re f")
                .Select(match => match.Groups.Cast<Group>().Skip(1)
                    .Select(group => double.Parse(group.Value, CultureInfo.InvariantCulture)).ToArray())
                .Where(rect => rect[2] > 0d && rect[2] < 3d).ToArray();
            TestAssert.Equal(count, bars.Length);
            return bars;
        }
    }

    private static void CheckBottomLegendRows(string kind)
    {
        foreach (int width in new[] { 288, 576 })
        foreach (int size in new[] { 8, 12, 18 })
        foreach (int count in new[] { 6, 9 })
        {
            var overlay = Render(width, size, count, overlay: true, kind: kind);
            var reserved = Render(width, size, count, overlay: false, kind: kind);
            // Independent Office long6/long9 and line9 controls: the narrow
            // 18pt frame admits six rows, even when nine names fit the whole frame.
            int admittedNames = width == 288 && size == 18 ? Math.Min(count, 6) : count;
            foreach (PdfPage page in new[] { overlay.Page, reserved.Page })
            {
                int rows = CheckCompleteNamesInFrame(page, width, admittedNames);
                TestAssert.Equal(admittedNames, DocxCommentsTests.ReadEmbeddedGlyphTextShows(page)
                    .SelectMany(item => Regex.Matches(item.Text, @"North West Region \d+")).Count());
                if (width == 288 && size >= 12)
                    TestAssert.True(rows > 1, "Long series names wrap instead of continuing beyond the frame.");
            }
            TestAssert.Equal(admittedNames < count, overlay.Diagnostics.Any(item => item.Message.Contains("overflowing entries", StringComparison.Ordinal)));
            TestAssert.Equal(admittedNames < count, reserved.Diagnostics.Any(item => item.Message.Contains("overflowing entries", StringComparison.Ordinal)));
            var overlayNames = DocxCommentsTests.ReadEmbeddedGlyphTextShows(overlay.Page).Where(item => item.Text.Contains("North West Region", StringComparison.Ordinal)).ToArray();
            var reservedNames = DocxCommentsTests.ReadEmbeddedGlyphTextShows(reserved.Page).Where(item => item.Text.Contains("North West Region", StringComparison.Ordinal)).ToArray();
            TestAssert.True(overlayNames.Select(item => (item.Text, item.X, item.Y)).SequenceEqual(reservedNames.Select(item => (item.Text, item.X, item.Y))),
                "Overlay changes plot admission while the bottom legend remains attached to its frame.");
        }
    }

    public static void PptxBottomStrokeLegendKeyKeepsItsOfficeSpanAcrossTextSizes()
    {
        // Independent Office PDF controls: 8/12/18pt text, narrow/wide frames.
        foreach (int size in new[] { 8, 12, 18 })
        {
            var page = Render(288, size, 1, overlay: true).Page;
            var label = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).First(item => item.Text.Contains("North West Region 1", StringComparison.Ordinal));
            var lines = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) m ([-\d.]+) ([-\d.]+) l S")
                .Select(match => match.Groups.Cast<Group>().Skip(1).Select(group => double.Parse(group.Value, CultureInfo.InvariantCulture)).ToArray())
                .Where(v => Math.Abs(v[1] - v[3]) < 0.001d && v[0] < label.X && v[2] < label.X && Math.Abs(v[1] - label.Y) < size);
            TestAssert.True(lines.Any(v => Math.Abs(v[2] - v[0] - 19.2d) < 0.002d),
                "The bottom stroke sample retains the independently measured Office extent as text size changes.");
        }
    }

    public static void PptxEmptyLegendLayoutMetadataKeepsWrappedContent()
    {
        foreach (string kind in new[] { "line", "bar", "column" })
        foreach (bool overlay in new[] { false, true })
        {
            var ordinary = Render(288, 12, 6, overlay, kind: kind).Page;
            CheckCompleteNamesInFrame(ordinary, 288, 6);
            foreach (string metadata in new[] { "<c:layout/>", "<c:layout><c:manualLayout/></c:layout>" })
            {
                var withMetadata = Render(288, 12, 6, overlay, metadata, kind).Page;
                CheckCompleteNamesInFrame(withMetadata, 288, 6);
                TestAssert.Equal(ordinary.Content, withMetadata.Content);
            }
        }
    }

    private static int CheckCompleteNamesInFrame(PdfPage page, int width, int count)
    {
        var shows = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page).ToArray();
        var baselines = new HashSet<double>();
        for (int index = 1; index <= count; index++)
        {
            string name = "North West Region " + index;
            var show = shows.First(item => item.Text.Contains(name, StringComparison.Ordinal));
            var font = show.Font.Font;
            double Advance(string text) => text.Sum(c => font.GetAdvanceWidth(font.MapCodePoint(c))) * show.Size / font.UnitsPerEm;
            double x = show.X + Advance(show.Text[..show.Text.IndexOf(name, StringComparison.Ordinal)]);
            TestAssert.True(x >= 71.99d && x + Advance(name) <= 72d + width + 0.01d,
                "Every complete legend name must fit horizontally inside its chart frame.");
            TestAssert.True(show.Y > page.Height - 468d && show.Y + show.Size < page.Height - 108d,
                "Wrapped legend baselines stay inside the chart, rather than below the slide.");
            baselines.Add(show.Y);
        }
        return baselines.Count;
    }

    private static (PdfPage Page, List<OoxPdfDiagnostic> Diagnostics) Render(int width, double size, int count, bool overlay, string legendMetadata = "", string kind = "line", int height = 360, bool categoryLegend = false, string[]? categoryNames = null, string plotLayout = "", string legendPosition = "b", IFontResolver? resolver = null, string seriesOptions = "", string dataLabels = "")
    {
        int fontHundredths = (int)Math.Round(size * 100d);
        string series = string.Concat(Enumerable.Range(1, count).Select(i => $$"""
            <c:ser><c:idx val="{{i - 1}}"/><c:order val="{{i - 1}}"/><c:tx><c:v>North West Region {{i}}</c:v></c:tx>
              <c:cat><c:strLit><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
            """));
        if (categoryLegend)
        {
            string categories = string.Concat(Enumerable.Range(0, count).Select(i =>
                $"<c:pt idx=\"{i}\"><c:v>{System.Security.SecurityElement.Escape(categoryNames?[i] ?? $"Category {i + 1}")}</c:v></c:pt>"));
            string values = string.Concat(Enumerable.Range(0, count).Select(i =>
                $"<c:pt idx=\"{i}\"><c:v>4</c:v></c:pt>"));
            series = $"<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>{seriesOptions}<c:cat><c:strLit><c:ptCount val=\"{count}\"/>{categories}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val=\"{count}\"/>{values}</c:numLit></c:val></c:ser>";
        }
        bool polar = kind is "pie" or "doughnut";
        bool bar = !polar && kind != "line";
        bool horizontal = kind == "bar";
        string plot = polar ? kind + "Chart" : bar ? "barChart" : "lineChart";
        string axes = polar ? "" : $$"""
                <c:catAx><c:axId val="1"/><c:axPos val="{{(horizontal ? "l" : "b")}}"/><c:crossAx val="2"/>{{(categoryLegend ? "<c:tickLblPos val=\"none\"/>" : "")}}</c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="{{(horizontal ? "b" : "l")}}"/><c:crossAx val="1"/><c:majorGridlines/></c:valAx>
            """;
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea>{{plotLayout}}<c:{{plot}}>{{(bar ? "<c:barDir val=\"" + (horizontal ? "bar" : "col") + "\"/>" : "")}}{{(polar ? "" : "<c:grouping val=\"" + (bar ? "clustered" : "standard") + "\"/>")}}
                <c:varyColors val="{{(categoryLegend ? 1 : 0)}}"/>{{series}}{{dataLabels}}{{(polar ? "" : "<c:axId val=\"1\"/><c:axId val=\"2\"/>")}}</c:{{plot}}>
                {{axes}}
              </c:plotArea><c:legend>{{legendMetadata}}<c:legendPos val="{{legendPosition}}"/><c:overlay val="{{(overlay ? 1 : 0)}}"/>
                <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{fontHundredths}}"/></a:pPr></a:p></c:txPr>
              </c:legend></c:chart>
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
        var diagnostics = new List<OoxPdfDiagnostic>();
        var page = new PptxRenderer(resolver ?? new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, diagnostics.Add, default).ToBlockingEnumerable().Single();
        return (page, diagnostics);
    }
}
