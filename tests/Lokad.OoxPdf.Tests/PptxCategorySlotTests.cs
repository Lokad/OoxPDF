using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxCategorySlotTests
{
    public static void PptxLineAndAreaDataFollowSourceCategorySlots()
    {
        // Independent Office pairs cover both crossing modes and three/eight
        // points. Markers and labels must stay attached when the mode changes.
        foreach (int count in new[] { 3, 8 })
        foreach (bool area in new[] { false, true })
        foreach (string mode in new[] { "between", "midCat" })
        {
            PdfPage page = Render(count, area, mode);
            double[] expected = ExpectedPoints(page, count, mode);
            AssertPositions(expected, Paths(page, "0 0 1", strokes: !area).SelectMany(x => x).Distinct().Order().ToArray());
            AssertPositions(expected, LabelCenters(page, "CCA"), tolerance: .1d);
            if (!area)
            {
                AssertPositions(expected, Paths(page, "0 1 0", strokes: false).Select(x => (x.Min() + x.Max()) / 2d).Order().ToArray());
                AssertPositions(expected, LabelCenters(page, "CLD"), tolerance: .1d);
            }
        }
        PdfPage single = Render(1, area: false, "between");
        AssertPositions(ExpectedPoints(single, 1, "between"), Paths(single, "0 1 0", strokes: false).Select(x => (x.Min() + x.Max()) / 2d).ToArray());
        TestAssert.Equal(Render(3, area: false, "between").Content, Render(3, area: false, null).Content);
        TestAssert.Equal(Render(3, area: true, "between").Content, Render(3, area: true, null).Content);
    }

    public static void PptxCategoryMajorTicksFollowSourceSlotBoundaries()
    {
        foreach (int count in new[] { 1, 3, 8 })
        foreach (bool area in new[] { false, true })
        foreach (string mode in new[] { "between", "midCat" })
        {
            PdfPage page = Render(count, area, mode);
            var lines = new List<string>();
            RgbColor[] colors = PptxChartPaletteTests.ReadColoredPaints(page.Content, lines.Add, strokes: true).ToArray();
            double[] ticks = colors.Select((color, index) => (color, line: lines[index]))
                .Where(x => x.color == new RgbColor(255, 0, 0))
                .Select(x => Regex.Match(x.line, @"^(?<x>[-\d.]+) [-\d.]+ m \k<x> [-\d.]+ l S$"))
                .Where(x => x.Success).Select(x => Number(x.Groups["x"].Value)).ToArray();
            var (left, right) = PlotEdges(page);
            int tickCount = Math.Max(2, mode == "midCat" ? count : count + 1);
            // Between slots uses n+1 boundary ticks; midCat uses n data ticks.
            AssertPositions(Enumerable.Range(0, tickCount).Select(i => left + (right - left) * i / (tickCount - 1)).ToArray(), ticks);
        }
    }

    public static void PptxCategorySlotsKeepSparseSourceIndexesAndStackedNeighbors()
    {
        foreach (bool area in new[] { false, true })
        foreach (bool stacked in new[] { false, true })
        foreach (string mode in new[] { "between", "midCat" })
        {
            PdfPage page = Render(5, area, mode, gap: true, stacked: stacked);
            double[] all = ExpectedPoints(page, 5, mode);
            AssertPositions(new[] { all[0], all[1], all[3], all[4] }, Paths(page, "0 0 1", strokes: !area).SelectMany(x => x).Distinct().Order().ToArray());
            if (!area)
                AssertPositions(new[] { all[0], all[1], all[3], all[4] }, LabelCenters(page, "CLD"), tolerance: .1d);
        }
    }

    private static double[] LabelCenters(PdfPage page, string role) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith(role, StringComparison.Ordinal))
            .Select(x => x.X + x.Font.MeasureTextPoints(x.Text, x.Size) / 2d).Order().ToArray();

    // Only straight, explicitly colored fixture paths are inspected. Clips and
    // font matrices do not participate in these source-coordinate assertions.
    private static double[][] Paths(PdfPage page, string color, bool strokes) =>
        Regex.Matches(page.Content, Regex.Escape(color) + (strokes ? " RG" : " rg") + @"\r?\n(?<block>.*?)(?=\r?\n[-\d.]+(?: [-\d.]+){0,2} " + (strokes ? "(?:RG|G)" : "(?:rg|g)") + @"\r?\n|\z)", RegexOptions.Singleline)
            .SelectMany(block => Regex.Matches(block.Groups["block"].Value, @"(?<path>(?:[-\d.]+ [-\d.]+ [ml]\s+)+)(?:h\s+)?" + (strokes ? "S" : "f") + @"(?:\s|$)"))
            .Select(x => Regex.Matches(x.Groups["path"].Value, @"(?<x>[-\d.]+) [-\d.]+ [ml]")
                .Select(p => Number(p.Groups["x"].Value)).ToArray()).Where(x => x.Length > 0).ToArray();

    private static (double Left, double Right) PlotEdges(PdfPage page)
    {
        Match? edge = Regex.Matches(page.Content, @"(?<left>[-\d.]+) (?<y>[-\d.]+) m\s+(?<right>[-\d.]+) \k<y> l S")
            .OrderByDescending(x => Math.Abs(Number(x.Groups["right"].Value) - Number(x.Groups["left"].Value))).FirstOrDefault();
        TestAssert.True(edge is not null, "Expected a full-width horizontal grid or category axis.");
        double x1 = Number(edge!.Groups["left"].Value), x2 = Number(edge.Groups["right"].Value);
        return (Math.Min(x1, x2), Math.Max(x1, x2));
    }

    private static double[] ExpectedPoints(PdfPage page, int count, string mode)
    {
        var (left, right) = PlotEdges(page);
        return Enumerable.Range(0, count).Select(i => left + (right - left) *
            (mode == "midCat" && count > 1 ? i / (double)(count - 1) : (i + .5d) / count)).ToArray();
    }

    private static void AssertPositions(double[] expected, double[] actual, double tolerance = .004d)
    {
        TestAssert.Equal(expected.Length, actual.Length);
        for (int i = 0; i < expected.Length; i++)
            TestAssert.True(Math.Abs(expected[i] - actual[i]) < tolerance, $"Source category slot{i}: expected {expected[i]}, got {actual[i]}.");
    }

    private static double Number(string text) => double.Parse(text, CultureInfo.InvariantCulture);

    private static PdfPage Render(int count, bool area, string? mode, bool gap = false, bool stacked = false)
    {
        string kind = area ? "areaChart" : "lineChart";
        string labels = string.Concat(Enumerable.Range(0, count).Select(i => $"<c:pt idx=\"{i}\"><c:v>Q{i}</c:v></c:pt>"));
        string values = string.Concat(Enumerable.Range(0, count).Where(i => !gap || i != 2).Select(i => $"<c:pt idx=\"{i}\"><c:v>{2 + i % 4}</c:v></c:pt>"));
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:{{kind}}><c:grouping val="{{(stacked ? "stacked" : "standard")}}"/><c:varyColors val="0"/>
                <c:ser><c:spPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr>
                  <c:marker><c:symbol val="diamond"/><c:size val="8"/><c:spPr><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></c:spPr></c:marker>
                  <c:dLbls><c:dLblPos val="t"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showLegendKey val="0"/></c:dLbls>
                  <c:cat><c:strLit><c:ptCount val="{{count}}"/>{{labels}}</c:strLit></c:cat>
                  <c:val><c:numLit><c:ptCount val="{{count}}"/>{{values}}</c:numLit></c:val>
                </c:ser><c:axId val="1"/><c:axId val="2"/></c:{{kind}}>
                <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/><c:majorTickMark val="out"/>
                  <c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr></c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling><c:min val="0"/><c:max val="10"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/><c:majorUnit val="2"/>
                  {{(mode is null ? "" : "<c:crossBetween val=\"" + mode + "\"/>")}}
                  <c:majorGridlines><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
                </c:valAx>
              </c:plotArea><c:dispBlanksAs val="gap"/></c:chart>
            </c:chartSpace>
            """;
        using var stream = PptxTests.CreateSingleChartPackageStream(chart);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
