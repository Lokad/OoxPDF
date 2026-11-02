using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartLabelDefaultsTests
{
    public static void PptxPointLabelFactorSizesUseOuterFrameInEveryPosition()
    {
        foreach (bool area in new[] { false, true })
        foreach (bool key in new[] { false, true })
        foreach (string position in new[] { "t", "b", "l", "r", "ctr" })
        {
            var page = RenderChartPage(ManualFrameSource(SizedFrameLayout(), key, area, position));
            var box = MarkerRectangles(page, new RgbColor(0, 255, 0)).Single();
            TestAssert.True(Math.Abs(box.Width - 72d) < .002d && Math.Abs(box.Height - 34.56d) < .002d,
                "Authored factor extents scale with the outer chart, independently of font, key and source point count.");
            var origin = KeyPlacementOrigin(page);
            TestAssert.True(origin.X > box.X && origin.X < box.X + box.Width && origin.Y > box.Y && origin.Y < box.Y + box.Height,
                "The measured baseline and centered text occupy the authored frame.");
        }
    }

    public static void PptxPointLabelSizedFactorOffsetsPreserveDimensions()
    {
        foreach (bool key in new[] { false, true })
        {
            var first = RenderChartPage(ManualFrameSource(SizedFrameLayout(), key));
            var second = RenderChartPage(ManualFrameSource(SizedFrameLayout(.08d, .1d), key));
            var before = KeyPlacementOrigin(first); var after = KeyPlacementOrigin(second);
            TestAssert.True(Math.Abs(after.X - before.X - 57.6d) < .002d && Math.Abs(after.Y - before.Y + 43.2d) < .002d,
                "Sized label offsets translate the same measured baseline in the outer frame.");
            var a = MarkerRectangles(first, new RgbColor(0, 255, 0)).Single();
            var b = MarkerRectangles(second, new RgbColor(0, 255, 0)).Single();
            TestAssert.Equal(a.Width, b.Width); TestAssert.Equal(a.Height, b.Height);
        }
    }

    public static void PptxPointLabelLargeSizedFramesFitBeforeFactorOffsets()
    {
        var page = RenderChartPage(ManualFrameSource(SizedFrameLayout(width: .9d, height: .9d), false));
        var box = MarkerRectangles(page, new RgbColor(0, 255, 0)).Single();
        TestAssert.True(Math.Abs(box.X - 144d) < .002d && Math.Abs(box.Y + box.Height - (page.Height - 36d)) < .002d,
            "Office admits the automatic sized frame inside the chart before translating it.");
        var moved = RenderChartPage(ManualFrameSource(SizedFrameLayout(-.08d, -.1d, width: .9d, height: .9d), false));
        var translated = MarkerRectangles(moved, new RgbColor(0, 255, 0)).Single();
        TestAssert.True(Math.Abs(translated.X - box.X + 57.6d) < .002d && Math.Abs(translated.Y - box.Y - 43.2d) < .002d,
            "Negative factor offsets stay distinct from invalid edge origins.");
    }

    public static void PptxPointLabelEmptySizedBoxesOmitTextAndKeys()
    {
        foreach ((double? width, double? height) in new (double?, double?)[] { (null, .08d), (.1d, null), (0d, .08d), (.1d, 0d) })
        {
            var page = RenderChartPage(ManualFrameSource(SizedFrameLayout(width: width, height: height), true));
            TestAssert.Equal("", AreaLabelText(page));
            TestAssert.Equal(0, MarkerRectangles(page, new RgbColor(0, 255, 0)).Length);
            TestAssert.Equal(1, MarkerRectangles(page, new RgbColor(255, 0, 0)).Length);
        }
    }

    public static void PptxPointLabelRejectedSizeLayoutsKeepAutomaticSource()
    {
        var automatic = RenderChartPage(ManualFrameSource(null, false));
        foreach (XElement layout in new[] { SizedFrameLayout(width: -.1d), SizedFrameLayout(height: -.1d),
            SizedFrameLayout(-.1d, .1d, edge: true), SizedFrameLayout(.1d, 1.1d, edge: true) })
        {
            var xml = ManualFrameSource(layout, false); string source = xml.ToString();
            TestAssert.Equal(automatic.Content, RenderChartPage(xml).Content);
            TestAssert.Equal(source, xml.ToString());
            TestAssert.True(PptxTests.BuildSingleChartScene(source)!.Plots[0].DataLabels.Overrides.Single(p => p.Index == 2).Layout.HasLayout,
                "Rejected rendering geometry keeps the original source declaration inspectable.");
        }
    }

    public static void PptxPointLabelSizedEdgesPlaceOccupiedBounds()
    {
        var page = RenderChartPage(ManualFrameSource(SizedFrameLayout(.2d, .2d, .3d, .3d, edge: true), false));
        var box = MarkerRectangles(page, new RgbColor(0, 255, 0)).Single();
        TestAssert.True(Math.Abs(box.X - 216d) < .002d && Math.Abs(box.Y + box.Height - (page.Height - 122.4d)) < .002d &&
            Math.Abs(box.Width - 216d) < .002d && Math.Abs(box.Height - 129.6d) < .002d,
            "Edge positions locate the complete sized frame; factor extents are not plot-relative.");
    }

    private static XElement SizedFrameLayout(double x = 0d, double y = 0d, double? width = .1d, double? height = .08d, bool edge = false)
    {
        var layout = ManualFrameLayout(edge ? "edge" : "factor", x, y);
        var manual = layout.Element(C + "manualLayout")!;
        manual.Add(width is null ? null : new XElement(C + "w", new XAttribute("val", width.Value)),
            height is null ? null : new XElement(C + "h", new XAttribute("val", height.Value)));
        return layout;
    }
}
