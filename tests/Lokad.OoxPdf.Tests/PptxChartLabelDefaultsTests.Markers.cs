using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxSceneRetainsNestedPointMarkerSourceAndDefaultState()
    {
        string[] markers = ["", "<c:spPr/>", "<c:symbol val=\"square\"/><c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></c:spPr>",
            "<c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></c:spPr>", "<c:symbol val=\"none\"/><c:size val=\"12\"/>"];
        foreach (string marker in markers)
        {
            XDocument xml = MarkerSource([PointMarker(1, marker)]);
            var point = PptxTests.BuildSingleChartScene(xml.ToString())!.Plots[0].Series[0].PointStyles.Single();
            TestAssert.True(point.Marker is { IsDefined: true }, "Even an ignored empty marker retains container presence.");
            TestAssert.Equal(marker.Contains("symbol", StringComparison.Ordinal), point.Marker!.SymbolSpecified);
            TestAssert.Equal(marker is not ("" or "<c:spPr/>"), point.HasMarkerAppearance);
            TestAssert.Equal(marker.Contains("size", StringComparison.Ordinal) ? 12d : point.Marker.SymbolSpecified ? 5d : 7d, point.Marker.Size);
            TestAssert.Equal(marker.Contains("size", StringComparison.Ordinal) ? "12" : null, point.Marker.SizeValue);
            if (!marker.Contains("symbol", StringComparison.Ordinal)) TestAssert.Equal(PptxSceneChartMarkerSymbol.Square, point.Marker.SymbolKind);
        }
    }

    public static void PptxPointMarkersPaintPlotAndKeysWithoutColorVariation()
    {
        foreach (bool vary in new[] { false, true })
        foreach (int index in new[] { 0, 2 })
        {
            var page = RenderChartPage(MarkerSource([PointMarker(index, RedMarker)], vary: vary));
            var red = MarkerRectangles(page, new RgbColor(255, 0, 0));
            TestAssert.Equal(2, red.Length);
            TestAssert.True(red.All(r => Math.Abs(r.Width - 5.04d) < .002d && Math.Abs(r.Height - 5.04d) < .002d),
                "The independently observed point-container default is nominal5 for both plot and key.");
        }
    }

    public static void PptxPointMarkerDefaultsDoNotInheritSeriesSizeOrEmptyContainers()
    {
        foreach (int size in new[] { 7, 18 })
        {
            var plain = RenderChartPage(MarkerSource([""], seriesSize: size, legend: true));
            foreach (string empty in new[] { "", "<c:spPr/>" })
                TestAssert.Equal(plain.Content, RenderChartPage(MarkerSource([PointMarker(1, empty)], seriesSize: size, legend: true)).Content);
            var own = RenderChartPage(MarkerSource([PointMarker(1, RedMarker)], seriesSize: size));
            TestAssert.Equal(2, MarkerRectangles(own, new RgbColor(255, 0, 0)).Length);
            TestAssert.True(MarkerRectangles(own, new RgbColor(255, 0, 0)).All(r => Math.Abs(r.Width - 5.04d) < .002d),
                "A defined point marker owns the default size rather than borrowing the series size.");
            string formatOnly = RedMarker.Replace("<c:symbol val=\"square\"/>", "", StringComparison.Ordinal);
            var inherited = RenderChartPage(MarkerSource([PointMarker(1, formatOnly)], seriesSize: size));
            double[] dimensions = MarkerRectangles(inherited, new RgbColor(255, 0, 0)).Select(r => r.Width).Order().ToArray();
            TestAssert.Equal(2, dimensions.Length);
            TestAssert.True(Math.Abs(dimensions[0] - (size == 7 ? 6.96d : 9.96d)) < .002d && Math.Abs(dimensions[1] - (size == 7 ? 6.96d : 18d)) < .002d,
                "Formatting without a point symbol keeps series geometry; key bounds stay independent.");
        }
    }

    public static void PptxPointMarkersKeepSeriesAndSparseSourceOwnership()
    {
        const string green = "<c:symbol val=\"square\"/><c:spPr><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></c:spPr>";
        var page = RenderChartPage(MarkerSource([PointMarker(0, RedMarker), PointMarker(2, green)], sparse: true, legend: true));
        TestAssert.Equal(1, MarkerRectangles(page, new RgbColor(255, 0, 0)).Length);
        TestAssert.Equal(1, MarkerRectangles(page, new RgbColor(0, 255, 0)).Length);
        TestAssert.Equal("ActualSecond", MarkerLegendText(page));
        var single = RenderChartPage(MarkerSource([PointMarker(2, RedMarker)], sparse: true));
        TestAssert.Equal(2, MarkerRectangles(single, new RgbColor(255, 0, 0)).Length);
    }

    public static void PptxPointAppearanceControlsLegendIdentityWithoutChangingSeriesColors()
    {
        var plain = RenderChartPage(MarkerSource([""], legend: true));
        TestAssert.Equal("Actual", MarkerLegendText(plain));
        var owned = RenderChartPage(MarkerSource([PointMarker(1, RedMarker)], legend: true));
        TestAssert.Equal("W1W2W3", MarkerLegendText(owned));
        TestAssert.Equal(3, MarkerRectangles(owned, new RgbColor(255, 0, 0)).Length);
    }

    public static void PptxPointMarkerKeysKeepIndependentSizeBoundsAcrossTextSizes()
    {
        string large = RedMarker.Replace("<c:spPr>", "<c:size val=\"12\"/><c:spPr>", StringComparison.Ordinal);
        foreach (int font in new[] { 9, 18, 36 })
        {
            // Native controls vary the label font while the chart root stays18.
            // Changing the root instead changes the key's independent capacity.
            var xml = MarkerSource([PointMarker(1, large)]);
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            xml.Descendants(C + "dLbls").Single().Add(new XElement(C + "txPr",
                new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "defRPr", new XAttribute("sz", font * 100))))));
            var page = RenderChartPage(xml);
            double[] sizes = MarkerRectangles(page, new RgbColor(255, 0, 0)).Select(r => r.Width).Order().ToArray();
            TestAssert.Equal(2, sizes.Length);
            TestAssert.True(Math.Abs(sizes[0] - 9.96d) < .002d && Math.Abs(sizes[1] - 12d) < .002d,
                "Native point-size12/font9,18,36 controls retain nominal10 key bounds and authored plot bounds.");
        }
    }

    public static void PptxPointMarkerNoFillAndNoneKeepVisibleTextAndPathPaint()
    {
        int filled = PptxChartPaletteTests.ReadColoredPaints(RenderChartPage(MarkerSource([PointMarker(1, RedMarker)])).Content).Count();
        foreach (string marker in new[] { "<c:symbol val=\"none\"/>", "<c:symbol val=\"square\"/><c:spPr><a:noFill/></c:spPr>" })
        {
            var page = RenderChartPage(MarkerSource([PointMarker(1, marker)]));
            TestAssert.Equal(0, MarkerRectangles(page, new RgbColor(255, 0, 0)).Length);
            TestAssert.Equal(filled - 2, PptxChartPaletteTests.ReadColoredPaints(page.Content).Count());
            string labels = string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
                fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text));
            TestAssert.Equal("243", labels);
        }
    }

    private const string RedMarker = "<c:symbol val=\"square\"/><c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></c:spPr>";

    private static string PointMarker(int index, string marker) => $"<c:dPt><c:idx val=\"{index}\"/><c:marker>{marker}</c:marker></c:dPt>";

    private static string MarkerLegendText(PdfPage page) => string.Concat(DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
        fontResourceFilter: resource => resource.StartsWith("CL", StringComparison.Ordinal) && !resource.StartsWith("CLD", StringComparison.Ordinal)).Select(s => s.Text));

    internal static (double X, double Y, double Width, double Height)[] MarkerRectangles(PdfPage page, RgbColor paint)
    {
        var paths = new List<string>();
        var colors = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add).ToArray();
        return colors.Select((color, i) => (color, path: paths[i])).Where(x => x.color == paint && x.path.EndsWith(" re f", StringComparison.Ordinal))
            .Select(x => x.path[..^5].Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(v => double.Parse(v, CultureInfo.InvariantCulture)).ToArray())
            .Select(v => (v[0], v[1], v[2], v[3])).ToArray();
    }

    private static XDocument MarkerSource(string[] points, bool vary = false, int seriesSize = 7, bool legend = false, bool sparse = false, int fontSize = 18)
    {
        string flags = string.Concat(Components.Select(c => $"<c:{c.Element} val=\"{(c.Property is "ShowValue" or "ShowLegendKey" ? 1 : 0)}\"/>"));
        var (_, plot) = Source("<c:dLbls>" + flags + "</c:dLbls>");
        plot.Name = C + "lineChart";
        plot.AddFirst(new XElement(C + "varyColors", new XAttribute("val", vary ? 1 : 0)));
        XElement prototype = new(plot.Element(C + "ser")!);
        plot.Elements(C + "ser").Remove();
        for (int i = 0; i < points.Length; i++)
        {
            XElement series = new(prototype);
            series.Element(C + "idx")!.SetAttributeValue("val", i);
            series.Element(C + "order")!.SetAttributeValue("val", i);
            series.Element(C + "tx")!.Element(C + "v")!.Value = i == 0 ? "Actual" : "Second";
            series.AddFirst(new XElement(C + "marker", new XElement(C + "symbol", new XAttribute("val", "square")), new XElement(C + "size", new XAttribute("val", seriesSize))));
            if (points[i].Length != 0)
                series.Add(XElement.Parse("<c:holder xmlns:c=\"" + C + "\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" + points[i] + "</c:holder>").Elements());
            if (sparse && i == 0) series.Element(C + "val")!.Descendants(C + "pt").Single(p => (string?)p.Attribute("idx") == "1").Remove();
            if (i != 0)
                foreach (var value in series.Element(C + "val")!.Descendants(C + "v"))
                    value.Value = (double.Parse(value.Value, CultureInfo.InvariantCulture) * .7d).ToString(CultureInfo.InvariantCulture);
            plot.Add(series);
        }
        var chart = plot.Parent!.Parent!;
        chart.AddFirst(new XElement(C + "autoTitleDeleted", new XAttribute("val", 1)));
        if (legend) chart.Add(new XElement(C + "legend", new XElement(C + "legendPos", new XAttribute("val", "r")), new XElement(C + "overlay", new XAttribute("val", 0))));
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        plot.Document!.Root!.Element(C + "txPr")!.Descendants(a + "defRPr").Single().SetAttributeValue("sz", fontSize * 100);
        return new XDocument(plot.Document);
    }
}
