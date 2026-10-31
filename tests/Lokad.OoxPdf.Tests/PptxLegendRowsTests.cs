using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxLegendRowsTests
{
    public static void PptxBottomLineLegendWrapKeepsNamesInsideTheChartFrame() => CheckBottomLegendRows("line");

    public static void PptxBottomColumnLegendWrapKeepsNamesInsideTheChartFrame() => CheckBottomLegendRows("column");

    private static void CheckBottomLegendRows(string kind)
    {
        foreach (int width in new[] { 288, 576 })
        foreach (int size in new[] { 8, 12, 18 })
        foreach (int count in new[] { 6, 9 })
        {
            var overlay = Render(width, size, count, overlay: true, kind: kind);
            var reserved = Render(width, size, count, overlay: false, kind: kind);
            foreach (PdfPage page in new[] { overlay.Page, reserved.Page })
            {
                int rows = CheckCompleteNamesInFrame(page, width, count);
                if (width == 288 && size >= 12)
                    TestAssert.True(rows > 1, "Long series names wrap instead of continuing beyond the frame.");
            }
            TestAssert.True(overlay.Diagnostics.Concat(reserved.Diagnostics).All(item => !item.Message.Contains("overflowing entries", StringComparison.Ordinal)),
                "These fitting names must not be admitted as clipped overflow.");
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

    private static (PdfPage Page, List<OoxPdfDiagnostic> Diagnostics) Render(int width, int size, int count, bool overlay, string legendMetadata = "", string kind = "line")
    {
        string series = string.Concat(Enumerable.Range(1, count).Select(i => $$"""
            <c:ser><c:idx val="{{i - 1}}"/><c:order val="{{i - 1}}"/><c:tx><c:v>North West Region {{i}}</c:v></c:tx>
              <c:cat><c:strLit><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strLit></c:cat>
              <c:val><c:numLit><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>
            """));
        bool bar = kind != "line";
        bool horizontal = kind == "bar";
        string plot = bar ? "barChart" : "lineChart";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:{{plot}}>{{(bar ? "<c:barDir val=\"" + (horizontal ? "bar" : "col") + "\"/>" : "")}}<c:grouping val="{{(bar ? "clustered" : "standard")}}"/>
                {{series}}<c:axId val="1"/><c:axId val="2"/></c:{{plot}}>
                <c:catAx><c:axId val="1"/><c:axPos val="{{(horizontal ? "l" : "b")}}"/><c:crossAx val="2"/></c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="5"/></c:scaling><c:axPos val="{{(horizontal ? "b" : "l")}}"/><c:crossAx val="1"/><c:majorGridlines/></c:valAx>
              </c:plotArea><c:legend>{{legendMetadata}}<c:legendPos val="b"/><c:overlay val="{{(overlay ? 1 : 0)}}"/>
                <c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{size * 100}}"/></a:pPr></a:p></c:txPr>
              </c:legend></c:chart>
            </c:chartSpace>
            """;
        string slide = $$"""
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <p:cSld><p:spTree><p:graphicFrame><p:xfrm><a:off x="914400" y="1371600"/><a:ext cx="{{width * 12700}}" cy="4572000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld>
            </p:sld>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart,
            new Dictionary<string, byte[]> { ["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide) });
        using var package = OoxPackage.Open(stream, default);
        var document = new PptxReader().Read(package, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(document, package, diagnostics.Add, default).ToBlockingEnumerable().Single();
        return (page, diagnostics);
    }
}
