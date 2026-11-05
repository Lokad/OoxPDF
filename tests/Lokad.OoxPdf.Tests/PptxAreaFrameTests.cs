using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxAreaFrameTests
{
    public static void PptxManualCategoryPlacementPreservesSourceSidesAndRectangle()
    {
        foreach (bool line in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        foreach (string mode in new[] { "autoZero", "max", "middle", "numeric" })
        foreach (string position in new[] { "nextTo", "low", "high" })
        {
            PdfPage page = Render(12, line: line, reversed: reversed, caption: position,
                categoryCrossing: mode, manual: true);
            var plot = Plot(page);
            var control = Plot(Render(12, line: line, reversed: reversed, manual: true));
            TestAssert.True(Math.Abs(plot.X - control.X) < .002 && Math.Abs(plot.Y - control.Y) < .002 &&
                Math.Abs(plot.Width - control.Width) < .002 && Math.Abs(plot.Height - control.Height) < .002,
                "Changing scale placement preserves a spacious authored inner rectangle.");
            double sourceRatio = mode == "middle" ? .5d : mode is "max" or "numeric" ? 1d : 0d;
            bool above = position == "nextTo" ? (mode == "max") != reversed : (position == "high") != reversed;
            double ratio = position == "nextTo" ? sourceRatio : position == "high" ? 1d : 0d;
            // Numeric next-to upper captions on a reversed scale use the
            // independently qualified low-side endpoint fallback.
            if (position == "nextTo" && mode == "numeric" && reversed) ratio = 0d;
            double anchor = plot.Y + plot.Height * (reversed ? 1d - ratio : ratio);
            AssertCaptionSide(page, anchor, above);
        }
    }

    public static void PptxManualLineAndAreaAxesUseTheSourceCrossing()
    {
        foreach (bool line in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        foreach (string mode in new[] { "autoZero", "max", "middle", "numeric" })
        {
            PdfPage page = Render(12, line: line, reversed: reversed, categoryCrossing: mode, manual: true);
            var plot = Plot(page);
            double ratio = mode == "middle" ? .5d : mode is "max" or "numeric" ? 1d : 0d;
            double expected = plot.Y + plot.Height * (reversed ? 1d - ratio : ratio);
            var axes = PptxColumnIntervalTests.Lines(page).Where(x => x.Color == new RgbColor(0, 255, 0) &&
                Math.Abs(x.V[1] - x.V[3]) < .001 && Math.Abs(x.V[0] - x.V[2]) > plot.Width / 2d).ToArray();
            TestAssert.True(axes.Length == 1 && Math.Abs(axes[0].V[1] - expected) < .003,
                $"The authored {line}/{reversed}/{mode} category stroke follows its numeric crossing.");
        }
    }

    public static void PptxAreaFrameMeasuresTitlesAndBottomLegendsTogether()
    {
        foreach (int size in new[] { 12, 24 })
        {
            var plain = Plot(Render(size));
            var bottom = Plot(Render(size, legend: true));
            var titled = Plot(Render(size, legend: true, title: true));
            TestAssert.True(Math.Abs(plain.X - bottom.X) < .002d && Math.Abs(bottom.X - titled.X) < .002d &&
                Math.Abs(plain.Width - bottom.Width) < .002d && Math.Abs(bottom.Width - titled.Width) < .002d,
                "Horizontal numeric strips are independent of the title and bottom legend.");
            TestAssert.True(bottom.Y > plain.Y + size / 2d && Math.Abs(bottom.Y + bottom.Height - plain.Y - plain.Height) < .002d,
                "A bottom legend reserves its measured band without moving the plot top.");
            TestAssert.True(Math.Abs(titled.Y - bottom.Y) < .002d && titled.Y + titled.Height < bottom.Y + bottom.Height - size / 2d,
                "A title reserves space above the plot while retaining the same legend/caption band below.");
            var overlaid = Plot(Render(size, legend: true, overlay: true));
            TestAssert.True(Math.Abs(plain.Y - overlaid.Y) < .002d && Math.Abs(plain.Height - overlaid.Height) < .002d,
                "An overlay legend does not consume plot height.");
        }
    }

    public static void PptxAreaFrameAndCaptionsShareTheSourceCrossing()
    {
        foreach (bool reversed in new[] { false, true })
        {
            PdfPage next = Render(12, legend: true, minimum: -2, reversed: reversed);
            PdfPage low = Render(12, legend: true, minimum: -2, reversed: reversed, caption: "low");
            var nextPlot = Plot(next);
            var lowPlot = Plot(low);
            double nextY = FirstCaptionY(next), lowY = FirstCaptionY(low);
            TestAssert.True(nextPlot.Height > lowPlot.Height + 1d,
                "Captions at an interior crossing release the outer caption strip for the plot.");
            TestAssert.True(reversed ? nextY < lowY - 1d : nextY > lowY + 1d,
                "Next-to captions follow the crossing and value direction; low remains at the numeric low edge.");
        }
    }

    public static void PptxAreaPrimaryValueLabelsUseTheirMeasuredSide()
    {
        foreach (bool right in new[] { false, true })
        {
            PdfPage page = Render(18, rightAxis: right);
            var plot = Plot(page);
            var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).ToArray();
            TestAssert.True(labels.Length >= 2, "Expected the primary numeric interval labels.");
            TestAssert.True(labels.All(x => right ? x.X > plot.X + plot.Width : x.X < plot.X),
                "Painting and numeric strip measurement must use the same primary axis side.");
        }
    }

    public static void PptxCategoryMaximumCrossingUsesSourceModeAndLabelOverrides()
    {
        foreach (bool line in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        foreach (string position in new[] { "nextTo", "low", "high" })
        {
            PdfPage page = Render(12, reversed: reversed, caption: position, line: line, categoryCrossing: "max");
            var plot = Plot(page);
            bool above = (position != "low") != reversed;
            double anchor = position == "low" ? (reversed ? plot.Y + plot.Height : plot.Y) :
                reversed ? plot.Y : plot.Y + plot.Height;
            AssertCaptionSide(page, anchor, above);
        }
        foreach (bool line in new[] { false, true })
        {
            PdfPage maximum = Render(12, line: line, categoryCrossing: "max");
            PdfPage numeric = Render(12, line: line, categoryCrossing: "numeric");
            var maxPlot = Plot(maximum);
            var numericPlot = Plot(numeric);
            AssertCaptionSide(numeric, numericPlot.Y + numericPlot.Height, above: false);
            TestAssert.True(maxPlot.Height < numericPlot.Height - 1d,
                "Authored maximum captions reserve an outer strip; numeric upper-bound captions stay inside.");
            TestAssert.Equal(numeric.Content, Render(12, line: line, categoryCrossing: "both").Content);
        }
    }

    public static void PptxCategoryTickSidesDistinguishNumericAndImplicitCrossings()
    {
        foreach (bool line in new[] { false, true })
        foreach (bool reversed in new[] { false, true })
        foreach (string position in new[] { "nextTo", "low", "high" })
        foreach (string mode in new[] { "middle", "autoZero", "min", "max" })
        {
            PdfPage page = Render(12, reversed: reversed, caption: position, line: line, categoryCrossing: mode);
            var plot = Plot(page);
            double crossing = mode == "middle" ? .5d : mode == "max" ? 1d : 0d;
            double axis = plot.Y + plot.Height * (reversed ? 1d - crossing : crossing);
            bool topSide = mode != "middle" && (mode == "max") != reversed;
            var paths = new List<string>();
            var paints = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add, strokes: true).ToArray();
            var ticks = paints.Select((color, i) => (color, path: paths[i]))
                .Where(x => x.color == new RgbColor(0, 255, 0))
                .Select(x => Regex.Match(x.path, @"^(?<x>[-\d.]+) (?<from>[-\d.]+) m \k<x> (?<to>[-\d.]+) l S$"))
                .Where(x => x.Success).ToArray();
            TestAssert.Equal(4, ticks.Length);
            foreach (Match tick in ticks)
            {
                double from = double.Parse(tick.Groups["from"].Value, CultureInfo.InvariantCulture);
                double to = double.Parse(tick.Groups["to"].Value, CultureInfo.InvariantCulture);
                TestAssert.True(Math.Abs(to - axis) < .002d && (topSide ? from > axis : from < axis),
                    "Numeric crossings retain the authored bottom tick side; implicit crossings mirror with value direction independently of captions.");
            }
        }
    }

    public static void PptxReversedNumericUpperCaptionsReuseLowSideGeometry()
    {
        foreach (int size in new[] { 12, 24 })
        foreach (bool line in new[] { false, true })
        foreach ((int minimum, int maximum, double crossing) in new[] { (0, 10, 10d), (0, 10, 11d), (-10, -1, -1d), (-10, -1, 0d) })
        {
            PdfPage RenderPosition(string position, double value) => Render(size, minimum: minimum, maximum: maximum,
                reversed: true, caption: position, line: line, categoryCrossing: "numeric", numericCrossing: value);
            PdfPage next = RenderPosition("nextTo", crossing), low = RenderPosition("low", crossing), high = RenderPosition("high", crossing);
            var nextPlot = Plot(next);
            var lowPlot = Plot(low);
            var highPlot = Plot(high);
            TestAssert.True(Math.Abs(FirstCaptionY(next) - FirstCaptionY(low)) < .002d &&
                Math.Abs(nextPlot.Y - lowPlot.Y) < .002d && Math.Abs(nextPlot.Height - lowPlot.Height) < .002d,
                "Reversed upper numeric next-to captions reuse the independently validated low-side band and frame.");
            AssertCaptionSide(next, nextPlot.Y + nextPlot.Height, above: true);
            AssertCaptionSide(high, highPlot.Y, above: false);
            PdfPage inside = RenderPosition("nextTo", maximum - 1d);
            TestAssert.True(Plot(inside).Height > nextPlot.Height + 1d && FirstCaptionY(inside) < FirstCaptionY(next) - 1d,
                "An interior numeric crossing retains its interpolated label position instead of the endpoint fallback.");
        }
    }

    public static void PptxAreaFillSharesCrossingAndStackLayerOwnership()
    {
        foreach (bool reversed in new[] { false, true })
        foreach (string mode in new[] { "autoZero", "max", "numeric", "middle" })
        {
            PdfPage page = Render(12, reversed: reversed, categoryCrossing: mode);
            var plot = Plot(page);
            double value = mode == "autoZero" ? 0d : mode == "middle" ? 5d : 10d;
            double expected = plot.Y + plot.Height * (reversed ? 1d - value / 10d : value / 10d);
            TestAssert.True(PolygonY(page, "0 0 1").TakeLast(3).All(y => Math.Abs(y - expected) < .002d),
                "The standard area closes on the resolved crossing, not an unrelated zero baseline.");
        }
        foreach (bool reversed in new[] { false, true })
        foreach (bool emptyFirst in new[] { false, true })
        {
            PdfPage page = Render(12, reversed: reversed, categoryCrossing: "max", stacked: true, emptyFirst: emptyFirst, maximum: 20);
            var plot = Plot(page);
            if (!emptyFirst)
            {
                double firstBase = reversed ? plot.Y : plot.Y + plot.Height;
                TestAssert.True(PolygonY(page, "0 0 1").TakeLast(3).All(y => Math.Abs(y - firstBase) < .002d),
                    "The first stacked source layer closes on the same crossing.");
            }
            double[] laterBase = PolygonY(page, "1 0 1").TakeLast(3).Reverse().ToArray();
            for (int i = 0; i < 3; i++)
            {
                double lower = emptyFirst ? 0d : 2d + 2d * i;
                double expected = plot.Y + plot.Height * (reversed ? 1d - lower / 20d : lower / 20d);
                TestAssert.True(Math.Abs(expected - laterBase[i]) < .002d,
                    "Later stack layers retain cumulative lower values and source series identity, even after an empty first series.");
            }
        }
    }

    private static double[] PolygonY(PdfPage page, string color)
    {
        Match polygon = Regex.Match(page.Content, Regex.Escape(color) + @" rg\r?\n(?<path>.*?)\s+f(?:\r?\n|$)", RegexOptions.Singleline);
        TestAssert.True(polygon.Success, "Expected the independently colored area layer.");
        double[] y = Regex.Matches(polygon.Groups["path"].Value, @"[-\d.]+ (?<y>[-\d.]+) [ml]")
            .Select(x => double.Parse(x.Groups["y"].Value, CultureInfo.InvariantCulture)).ToArray();
        TestAssert.Equal(6, y.Length);
        return y;
    }

    private static void AssertCaptionSide(PdfPage page, double anchor, bool above)
    {
        var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
        TestAssert.Equal(3, captions.Length);
        foreach (var show in captions)
        {
            double bottom = double.PositiveInfinity, top = double.NegativeInfinity;
            foreach (char character in show.Text)
            {
                TestAssert.True(show.Font.Font.TryGetGlyphBounds(show.Font.Font.MapCodePoint(character), out var box), "Expected pinned caption glyph ink.");
                bottom = Math.Min(bottom, show.Y + box.YMin * show.Size / show.Font.Font.UnitsPerEm);
                top = Math.Max(top, show.Y + box.YMax * show.Size / show.Font.Font.UnitsPerEm);
            }
            TestAssert.True(above ? bottom > anchor : top < anchor, "Caption ink follows the authored crossing mode and explicit label position.");
        }
    }

    private static double FirstCaptionY(PdfPage page) =>
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).First().Y;

    private static (double X, double Y, double Width, double Height) Plot(PdfPage page)
    {
        var paths = new List<string>();
        var paints = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add, strokes: true).ToArray();
        string rectangle = paints.Select((color, i) => (color, path: paths[i]))
            .Single(x => x.color == new RgbColor(255, 0, 0) && x.path.EndsWith(" re S", StringComparison.Ordinal)).path;
        double[] v = rectangle[..^5].Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(x => double.Parse(x, CultureInfo.InvariantCulture)).ToArray();
        return (v[0], v[1], v[2], v[3]);
    }

    private static PdfPage Render(int size, bool legend = false, bool title = false, bool overlay = false,
        int minimum = 0, bool reversed = false, string caption = "nextTo", bool rightAxis = false,
        bool line = false, string categoryCrossing = "autoZero", bool stacked = false, bool emptyFirst = false, int maximum = 10,
        double? numericCrossing = null, bool manual = false)
    {
        string kind = line ? "lineChart" : "areaChart";
        string chart = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <c:chart>{{(title ? "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz=\"" + size * 100 + "\"/><a:t>Area</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>" : "<c:autoTitleDeleted val=\"1\"/>")}}
                <c:plotArea>{{(manual ? "<c:layout><c:manualLayout><c:layoutTarget val=\"inner\"/><c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:wMode val=\"factor\"/><c:hMode val=\"factor\"/><c:x val=\".2\"/><c:y val=\".2\"/><c:w val=\".6\"/><c:h val=\".5\"/></c:manualLayout></c:layout>" : "")}}
                <c:{{kind}}><c:grouping val="{{(stacked ? "stacked" : "standard")}}"/><c:varyColors val="0"/><c:ser><c:tx><c:v>Series</c:v></c:tx>
                  <c:spPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr>
                  <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
                  <c:val><c:numLit><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>6</c:v></c:pt></c:numLit></c:val>
                </c:ser><c:axId val="1"/><c:axId val="2"/></c:{{kind}}>
                <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:tickLblPos val="{{caption}}"/><c:crossAx val="2"/>
                  {{(categoryCrossing is "numeric" or "both" or "middle" ? "<c:crossesAt val=\"" + (categoryCrossing == "middle" ? 5d : numericCrossing ?? maximum).ToString(CultureInfo.InvariantCulture) + "\"/>" : "")}}
                  {{(categoryCrossing is "numeric" or "middle" ? "" : "<c:crosses val=\"" + (categoryCrossing == "both" ? "max" : categoryCrossing) + "\"/>")}}
                  <c:majorTickMark val="out"/><c:spPr><a:ln><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></c:spPr></c:catAx>
                <c:valAx><c:axId val="2"/><c:scaling><c:min val="{{minimum}}"/><c:max val="{{maximum}}"/><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/></c:scaling><c:axPos val="{{(rightAxis ? "r" : "l")}}"/><c:crossAx val="1"/><c:crosses val="{{(rightAxis ? "max" : "autoZero")}}"/><c:crossBetween val="between"/><c:majorUnit val="2"/><c:majorGridlines/></c:valAx>
                <c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></c:spPr>
                </c:plotArea>{{(legend ? "<c:legend><c:legendPos val=\"b\"/><c:overlay val=\"" + (overlay ? "1" : "0") + "\"/></c:legend>" : "")}}
              </c:chart><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{{size * 100}}"/></a:pPr></a:p></c:txPr>
            </c:chartSpace>
            """;
        if (stacked)
        {
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart", a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XDocument xml = XDocument.Parse(chart);
            XElement first = xml.Descendants(c + "ser").Single();
            first.AddFirst(new XElement(c + "idx", new XAttribute("val", "0")), new XElement(c + "order", new XAttribute("val", "0")));
            XElement second = new(first);
            second.Element(c + "idx")!.SetAttributeValue("val", "1");
            second.Element(c + "order")!.SetAttributeValue("val", "1");
            second.Element(c + "spPr")!.Element(a + "solidFill")!.Element(a + "srgbClr")!.SetAttributeValue("val", "FF00FF");
            first.AddAfterSelf(second);
            if (emptyFirst) first.Element(c + "val")!.ReplaceNodes(new XElement(c + "numLit", new XElement(c + "ptCount", new XAttribute("val", "0"))));
            chart = xml.ToString(SaveOptions.DisableFormatting);
        }
        using var stream = PptxTests.CreateSingleChartPackageStream(chart);
        using var package = OoxPackage.Open(stream, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
