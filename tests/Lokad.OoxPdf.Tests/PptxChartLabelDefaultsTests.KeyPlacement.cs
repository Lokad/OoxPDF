using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxLineLabelDefaultsUseRightWithOneOrAllLabels()
    {
        foreach (bool all in new[] { false, true })
        foreach (bool key in new[] { false, true })
            TestAssert.Equal(RenderChartPage(KeyPlacementSource(null, 12, key, all: all)).Content,
                RenderChartPage(KeyPlacementSource("r", 12, key, all: all)).Content);
    }

    public static void PptxLineLabelsReserveAuthoredMarkerExtentInEachPosition()
    {
        // Independent native5/12 controls move an outside label by3.5pt,
        // including a none symbol; a centered label does not move.
        foreach (string position in new[] { "t", "b", "l", "r", "ctr" })
        foreach (bool key in new[] { false, true })
        foreach (bool none in new[] { false, true })
        {
            var small = KeyPlacementOrigin(RenderChartPage(KeyPlacementSource(position, 5, key, none: none)));
            var large = KeyPlacementOrigin(RenderChartPage(KeyPlacementSource(position, 12, key, none: none)));
            double dx = large.X - small.X, dy = large.Y - small.Y;
            double expectedX = position == "l" ? -3.5d : position == "r" ? 3.5d : 0d;
            double expectedY = position == "b" ? -3.5d : position == "t" ? 3.5d : 0d;
            TestAssert.True(Math.Abs(dx - expectedX) < .04d && Math.Abs(dy - expectedY) < .04d,
                "Native marker clearance moves the complete text/key unit; a key resize must not shift its text independently.");
        }
    }

    public static void PptxPointLabelKeysAlignWithMeasuredTextAcrossSizes()
    {
        foreach (int font in new[] { 9, 18, 36 })
        {
            var page = RenderChartPage(KeyPlacementSource("ctr", 12, true, font: font));
            var origin = KeyPlacementOrigin(page);
            var key = MarkerRectangles(page, new RgbColor(255, 0, 0)).Last();
            double centerAboveBaseline = key.Y + key.Height / 2d - origin.Y;
            TestAssert.True(centerAboveBaseline > 0d && centerAboveBaseline < font * .5d,
                "Independent native keys center within the font line, rather than half an unrelated1.35em label box above its baseline.");
        }
    }

    public static void PptxPointLabelBackgroundContainsTextAndExcludesKey()
    {
        foreach (string position in new[] { "t", "b", "l", "r", "ctr" })
        {
            XDocument xml = KeyPlacementSource(position, 12, true);
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            xml.Descendants(C + "dLbls").Single().Add(new XElement(C + "spPr",
                new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "00FF00")))));
            var page = RenderChartPage(xml);
            var background = MarkerRectangles(page, new RgbColor(0, 255, 0)).Single();
            var origin = KeyPlacementOrigin(page);
            var key = MarkerRectangles(page, new RgbColor(255, 0, 0)).Last();
            TestAssert.True(background.Y < origin.Y && origin.Y < background.Y + background.Height,
                "The background owns occupied line bounds on both sides of its text baseline.");
            TestAssert.True(key.X + key.Width < background.X,
                "Native label backgrounds own text and its padding; the key paints alongside that frame.");
        }
    }

    public static void PptxPointLabelManualOffsetsKeepTheirAuthoredFrame()
    {
        var plain = KeyPlacementSource("r", 12, true);
        var moved = new XDocument(plain);
        XElement Manual(double y) => new(C + "dLbl", new XElement(C + "idx", new XAttribute("val", 2)),
            new XElement(C + "layout", new XElement(C + "manualLayout",
                new XElement(C + "xMode", new XAttribute("val", "factor")),
                new XElement(C + "yMode", new XAttribute("val", "factor")),
                new XElement(C + "x", new XAttribute("val", .08d)),
                new XElement(C + "y", new XAttribute("val", y.ToString(CultureInfo.InvariantCulture))))));
        plain.Descendants(C + "dLbls").Single().AddFirst(Manual(0d));
        moved.Descendants(C + "dLbls").Single().AddFirst(Manual(.1d));
        var first = KeyPlacementOrigin(RenderChartPage(plain));
        var second = KeyPlacementOrigin(RenderChartPage(moved));
        TestAssert.True(Math.Abs(first.X - second.X) < .002d && second.Y < first.Y - 10d,
            "Manual factor offsets keep their authored vertical frame while measured keys follow the resulting baseline.");
    }

    private static (double X, double Y) KeyPlacementOrigin(PdfPage page)
    {
        Match show = Regex.Match(page.Content, @"/CLD\S+ [\d.]+ Tf\s+\S+ Tc\s+1 0 0 1 (?<x>[-\d.]+) (?<y>[-\d.]+) Tm");
        TestAssert.True(show.Success, "Expected the sole retained source label.");
        return (double.Parse(show.Groups["x"].Value, CultureInfo.InvariantCulture),
            double.Parse(show.Groups["y"].Value, CultureInfo.InvariantCulture));
    }

    private static XDocument KeyPlacementSource(string? position, int size, bool key, int font = 18, bool all = false, bool none = false)
    {
        string marker = RedMarker.Replace("<c:spPr>", $"<c:size val=\"{size}\"/><c:spPr>", StringComparison.Ordinal);
        if (none) marker = marker.Replace("val=\"square\"", "val=\"none\"", StringComparison.Ordinal);
        XDocument xml = MarkerSource([PointMarker(2, marker)], fontSize: font);
        XElement labels = xml.Descendants(C + "dLbls").Single();
        labels.Element(C + "showLegendKey")!.SetAttributeValue("val", key ? 1 : 0);
        if (position is not null) labels.AddFirst(new XElement(C + "dLblPos", new XAttribute("val", position)));
        if (!all)
            foreach (int index in new[] { 0, 1 })
                labels.AddFirst(new XElement(C + "dLbl", new XElement(C + "idx", new XAttribute("val", index)),
                    new XElement(C + "delete", new XAttribute("val", 1))));
        return xml;
    }
}
