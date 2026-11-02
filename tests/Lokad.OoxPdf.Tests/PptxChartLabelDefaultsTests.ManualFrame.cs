using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxPointLabelZeroOffsetsPreserveMeasuredPlacement()
    {
        foreach (bool area in new[] { false, true })
        foreach (bool key in new[] { false, true })
        foreach (string? position in area ? new string?[] { null } : new[] { "t", "b", "l", "r", "ctr" })
        {
            var automatic = RenderChartPage(ManualFrameSource(null, key, area, position));
            var zero = RenderChartPage(ManualFrameSource(ManualFrameLayout("factor", 0d, 0d), key, area, position));
            TestAssert.Equal(automatic.Content, zero.Content);
        }
    }

    public static void PptxPointLabelFactorsUseOuterChartDimensions()
    {
        foreach (bool area in new[] { false, true })
        foreach (bool key in new[] { false, true })
        foreach (string? mode in new string?[] { "factor", null })
        {
            var plain = RenderChartPage(ManualFrameSource(null, key, area));
            var moved = RenderChartPage(ManualFrameSource(ManualFrameLayout(mode, .08d, .1d), key, area));
            var first = KeyPlacementOrigin(plain); var second = KeyPlacementOrigin(moved);
            TestAssert.True(Math.Abs(second.X - first.X - 57.6d) < .002d && Math.Abs(second.Y - first.Y + 43.2d) < .002d,
                "Native factor offsets use the720×432pt chart frame, including absent modes.");
            var before = MarkerRectangles(plain, new RgbColor(0, 255, 0)).Single();
            var after = MarkerRectangles(moved, new RgbColor(0, 255, 0)).Single();
            TestAssert.True(Math.Abs(after.X - before.X - 57.6d) < .002d && Math.Abs(after.Y - before.Y + 43.2d) < .002d &&
                before.Width == after.Width && before.Height == after.Height, "The occupied background translates without changing its measured size.");
        }
    }

    public static void PptxPointLabelEdgesPlaceOccupiedFrameAtChartOrigin()
    {
        foreach (bool area in new[] { false, true })
        foreach (double offset in new[] { 0d, .1d })
        {
            var page = RenderChartPage(ManualFrameSource(ManualFrameLayout("edge", offset, offset), false, area));
            var box = MarkerRectangles(page, new RgbColor(0, 255, 0)).Single();
            TestAssert.True(Math.Abs(box.X - 72d - offset * 720d) < .002d &&
                Math.Abs(box.Y + box.Height - (page.Height - 36d - offset * 432d)) < .002d,
                "Edge positions own occupied top-left bounds in the outer chart frame.");
            var origin = KeyPlacementOrigin(page);
            TestAssert.True(origin.Y > box.Y && origin.Y < box.Y + box.Height, "The baseline remains inside its translated background.");
        }
    }

    public static void PptxPointLabelManualLinksTranslateWithPaint()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var parts = new Dictionary<string, byte[]> { ["ppt/charts/_rels/chart1.xml.rels"] = TestFixtures.Utf8("""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/manual-label" TargetMode="External"/></Relationships>
            """) };
        PdfPage Render(XElement? layout)
        {
            var xml = ManualFrameSource(layout, false);
            xml.Descendants(C + "dLbl").Single(p => (int)p.Element(C + "idx")!.Attribute("val")! == 2).Add(
                new XElement(C + "tx", new XElement(C + "rich", new XElement(a + "bodyPr"),
                    new XElement(a + "p", new XElement(a + "r", new XElement(a + "rPr", new XElement(a + "hlinkClick", new XAttribute(r + "id", "rIdLink"))),
                        new XElement(a + "t", "Point link"))))));
            return RenderChartPage(xml, parts);
        }
        foreach (bool sized in new[] { false, true })
        {
            var plain = Render(sized ? SizedFrameLayout() : null);
            var moved = Render(sized ? SizedFrameLayout(.08d, .1d) : ManualFrameLayout("factor", .08d, .1d));
            var before = plain.Annotations.Single(); var after = moved.Annotations.Single();
            TestAssert.Equal("https://example.invalid/manual-label", after.Uri);
            TestAssert.True(Math.Abs(after.X - before.X - 57.6d) < .002d && Math.Abs(after.Y - before.Y + 43.2d) < .002d &&
                before.Width == after.Width && before.Height == after.Height, "Link geometry follows the same admitted translation as painted text.");
        }
    }

    private static XElement ManualFrameLayout(string? mode, double x, double y) => new(C + "layout", new XElement(C + "manualLayout",
        mode is null ? null : new XElement(C + "xMode", new XAttribute("val", mode)),
        mode is null ? null : new XElement(C + "yMode", new XAttribute("val", mode)),
        new XElement(C + "x", new XAttribute("val", x)), new XElement(C + "y", new XAttribute("val", y))));

    private static XDocument ManualFrameSource(XElement? layout, bool key, bool area = false, string? position = "r")
    {
        var xml = KeyPlacementSource(area ? null : position, 12, key);
        XElement labels = xml.Descendants(C + "dLbls").Single();
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        labels.AddFirst(new XElement(C + "dLbl", new XElement(C + "idx", new XAttribute("val", 2)), layout,
            new XElement(C + "spPr", new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "00FF00")))),
            Components.Select(c => new XElement(C + c.Element, new XAttribute("val", c.Property == "ShowValue" || key && c.Property == "ShowLegendKey" ? 1 : 0)))));
        if (area) xml.Descendants(C + "lineChart").Single().Name = C + "areaChart";
        return xml;
    }
}
