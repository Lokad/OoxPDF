using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxAreaLabelsKeepDefaultComponentsDeletionAndCustomText()
    {
        var (_, plot) = Source("<c:dLbls/>", """
            <c:dLbls><c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl>
              <c:dLbl><c:idx val="2"/><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Pinned</a:t></a:r></a:p></c:rich></c:tx></c:dLbl></c:dLbls>
            """);
        plot.Name = C + "areaChart";
        var page = RenderChartPage(new XDocument(plot.Document!));
        TestAssert.Equal("Actual, W1, 2Pinned", AreaLabelText(page));
    }

    public static void PptxAreaLabelsCenterInStandardStackedAndPercentSourceSpans()
    {
        // Independent Office controls center default labels in each physical
        // area span, including negative values, reversed axes and crossings.
        foreach (string grouping in new[] { "standard", "stacked", "percentStacked" })
        foreach (bool reversed in new[] { false, true })
        foreach (bool categoryReversed in new[] { false, true })
        foreach (bool hideFirst in new[] { false, true })
        {
            var page = AreaPage([[2d, 4d, 3d], [3d, 2d, 4d]], grouping, reversed: reversed,
                categoryReversed: categoryReversed, hideFirst: hideFirst);
            var boxes = AreaLabelBoxes(page);
            TestAssert.Equal(hideFirst ? 3 : 6, boxes.Length);
            for (int series = hideFirst ? 1 : 0; series < 2; series++)
            {
                var polygon = AreaPolygon(page, series == 0 ? "0 0 1" : "1 0 0");
                for (int i = 0; i < 3; i++)
                {
                    var box = boxes[(series - (hideFirst ? 1 : 0)) * 3 + i];
                    double expectedY = (polygon[i].Y + polygon[5 - i].Y) / 2d;
                    TestAssert.True(Math.Abs(box.X + box.Width / 2d - polygon[i].X) < .002d &&
                        Math.Abs(box.Y + box.Height / 2d - expectedY) < .002d,
                        "Area labels share source slots and the physical center of their own layer, even when lower labels are hidden.");
                }
            }
            TestAssert.Equal(hideFirst ? "324" : "243324", AreaLabelText(page));
        }
        foreach (bool reversed in new[] { false, true })
        foreach (double crossing in new[] { 0d, 3d })
        {
            var page = AreaPage([[-2d, -4d, -3d]], reversed: reversed, minimum: -10d, crossing: crossing);
            var polygon = AreaPolygon(page, "0 0 1");
            var boxes = AreaLabelBoxes(page);
            for (int i = 0; i < 3; i++)
                TestAssert.True(Math.Abs(boxes[i].Y + boxes[i].Height / 2d - (polygon[i].Y + polygon[5 - i].Y) / 2d) < .002d,
                    "Negative area labels use the same clamped source crossing and reversed span as paint.");
        }
    }

    public static void PptxAreaLabelsKeepSparseIsolatedSourcePoints()
    {
        var page = AreaPage([[2d, null, 3d], [3d, 2d, 4d]], "stacked");
        TestAssert.Equal("23324", AreaLabelText(page));
        var boxes = AreaLabelBoxes(page);
        TestAssert.Equal(5, boxes.Length);
        var second = AreaPolygon(page, "1 0 0");
        for (int i = 0; i < 3; i++)
            TestAssert.True(Math.Abs(boxes[2 + i].Y + boxes[2 + i].Height / 2d - (second[i].Y + second[5 - i].Y) / 2d) < .002d,
                "An isolated lower source value remains part of the higher layer's label span.");
    }

    public static void PptxStackedAreaPaintKeepsIsolatedLowerContributions()
    {
        var page = AreaPage([[2d, null, 3d], [3d, 2d, 4d]], "stacked", labels: false);
        var second = AreaPolygon(page, "1 0 0");
        TestAssert.True(second[5].Y > second[4].Y && second[3].Y > second[4].Y,
            "Higher area layers stack on isolated lower source values even when that lower series has no two-point polygon.");
        TestAssert.Equal("", AreaLabelText(page));
    }

    public static void PptxAreaCustomLabelLinksFollowSharedTextGeometry()
    {
        var (_, plot) = Source("", """
            <c:dLbls><c:dLbl><c:idx val="0"/><c:tx><c:rich><a:bodyPr/><a:p><a:r>
              <a:rPr><a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdLink"/></a:rPr>
              <a:t>Point link</a:t></a:r></a:p></c:rich></c:tx></c:dLbl>
              <c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showLegendKey val="0"/></c:dLbls>
            """);
        plot.Name = C + "areaChart";
        var page = RenderChartPage(new XDocument(plot.Document!), new Dictionary<string, byte[]>
        {
            ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/area-label" TargetMode="External"/></Relationships>
                """)
        });
        TestAssert.Equal("Point link", AreaLabelText(page));
        var link = page.Annotations.Single();
        TestAssert.Equal("https://example.invalid/area-label", link.Uri);
        var show = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
            fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Single();
        TestAssert.True(link.Width > 0d && link.Height > 0d && Math.Abs(link.X - show.X) < .01d,
            "Area label links cover the same painted text origin with a non-degenerate run rectangle.");
    }

    private static string AreaLabelText(PdfPage page) => string.Concat(
        DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, false,
            fontResourceFilter: resource => resource.StartsWith("CLD", StringComparison.Ordinal)).Select(show => show.Text));

    private static (double X, double Y)[] AreaPolygon(PdfPage page, string color)
    {
        Match polygon = Regex.Match(page.Content, Regex.Escape(color) + @" rg\r?\n(?<path>.*?)\s+f(?:\r?\n|$)", RegexOptions.Singleline);
        TestAssert.True(polygon.Success, "Expected the independently colored area layer.");
        var points = Regex.Matches(polygon.Groups["path"].Value, @"(?<x>[-\d.]+) (?<y>[-\d.]+) [ml]")
            .Select(m => (Number(m.Groups["x"].Value), Number(m.Groups["y"].Value))).ToArray();
        TestAssert.Equal(6, points.Length);
        return points;
    }

    private static (double X, double Y, double Width, double Height)[] AreaLabelBoxes(PdfPage page)
    {
        var paths = new List<string>();
        var paints = PptxChartPaletteTests.ReadColoredPaints(page.Content, paths.Add).ToArray();
        return paints.Select((color, i) => (color, path: paths[i]))
            .Where(x => x.color == new RgbColor(18, 52, 86) && x.path.EndsWith(" re f", StringComparison.Ordinal))
            .Select(x => x.path[..^5].Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(Number).ToArray())
            .Select(v => (v[0], v[1], v[2], v[3])).ToArray();
    }

    private static double Number(string value) => double.Parse(value, CultureInfo.InvariantCulture);

    private static PdfPage AreaPage(double?[][] values, string grouping = "standard", bool reversed = false,
        bool categoryReversed = false, bool hideFirst = false, double minimum = 0d, double crossing = 0d, bool labels = true)
    {
        const string background = "<c:spPr><a:solidFill><a:srgbClr val=\"123456\"/></a:solidFill></c:spPr>";
        string disabled = string.Concat(Components.Select(c => $"<c:{c.Element} val=\"0\"/>"));
        string labelXml = labels ? "<c:dLbls>" + disabled.Replace("<c:showVal val=\"0\"/>", "<c:showVal val=\"1\"/>", StringComparison.Ordinal) + background + "</c:dLbls>" : "";
        var (_, plot) = Source(labelXml);
        plot.Name = C + "areaChart";
        plot.AddFirst(new XElement(C + "grouping", new XAttribute("val", grouping)));
        XElement prototype = new(plot.Element(C + "ser")!);
        plot.Elements(C + "ser").Remove();
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        for (int i = 0; i < values.Length; i++)
        {
            XElement series = new(prototype);
            series.Element(C + "idx")!.SetAttributeValue("val", i);
            series.Element(C + "order")!.SetAttributeValue("val", i);
            series.AddFirst(new XElement(C + "spPr", new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", i == 0 ? "0000FF" : "FF0000")))));
            series.Element(C + "val")!.Element(C + "numLit")!.ReplaceNodes(values[i].Select((v, index) => v is null ? null :
                new XElement(C + "pt", new XAttribute("idx", index), new XElement(C + "v", v.Value.ToString(CultureInfo.InvariantCulture)))));
            if (hideFirst && i == 0) series.Add(new XElement(C + "dLbls", XElement.Parse("<c:dLbls xmlns:c=\"" + C + "\">" + disabled + "</c:dLbls>").Elements()));
            plot.Add(series);
        }
        plot.Add(new XElement(C + "axId", new XAttribute("val", 1)), new XElement(C + "axId", new XAttribute("val", 2)));
        plot.Parent!.Add(XElement.Parse($$"""
            <c:catAx xmlns:c="{{C}}"><c:axId val="1"/><c:scaling><c:orientation val="{{(categoryReversed ? "maxMin" : "minMax")}}"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/><c:crossesAt val="{{crossing.ToString(CultureInfo.InvariantCulture)}}"/><c:tickLblPos val="none"/></c:catAx>
            """));
        plot.Parent.Add(XElement.Parse($$"""
            <c:valAx xmlns:c="{{C}}"><c:axId val="2"/><c:scaling><c:min val="{{minimum.ToString(CultureInfo.InvariantCulture)}}"/><c:max val="{{(grouping == "percentStacked" ? "1" : "10")}}"/><c:orientation val="{{(reversed ? "maxMin" : "minMax")}}"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/><c:tickLblPos val="none"/></c:valAx>
            """));
        plot.Parent.Parent!.Add(new XElement(C + "dispBlanksAs", new XAttribute("val", "gap")));
        return RenderChartPage(new XDocument(plot.Document!));
    }
}
