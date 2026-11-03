using System.Xml.Linq;
using System.Globalization;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Pptx;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    private static readonly XNamespace CrossingC = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public static void PptxHorizontalCategoryCrossingsFollowTheNumericScale()
    {
        // Four independent Office controls establish the clamped zero crossing
        // in both directions. Fonts/plot admission may change the rectangle,
        // but cannot change its source-relative crossing.
        foreach (int minimum in new[] { 7, -7 })
        foreach (bool reversed in new[] { false, true })
        {
            var actual = HorizontalGeometry(Render(minimum, 25, reversed: reversed, horizontal: true));
            TestAssert.Equal(2, actual.Labels.Length);
            double ratio = (Math.Clamp(0d, minimum, 103d) - minimum) / (103d - minimum);
            if (reversed) ratio = 1d - ratio;
            NearHorizontal(actual.Axis, actual.Left + ratio * (actual.Right - actual.Left));
            TestAssert.True(reversed ? actual.Labels.All(x => x > actual.Axis) : actual.Labels.All(x => x < actual.Axis),
                "Next-to category labels follow the crossing and numeric direction.");
        }
    }

    public static void PptxHorizontalCategoryLabelSidesKeepTheirStrokeCrossing()
    {
        foreach (string position in new[] { "low", "high" })
        {
            var actual = HorizontalGeometry(Render(7, 25, reversed: true, horizontal: true, editChart: xml =>
                xml.Descendants(CrossingC + "catAx").Single().Add(new XElement(CrossingC + "tickLblPos", new XAttribute("val", position)))));
            NearHorizontal(actual.Axis, actual.Right);
            TestAssert.Equal(2, actual.Labels.Length);
            TestAssert.True(position == "low" ? actual.Labels.All(x => x > actual.Right) : actual.Labels.All(x => x < actual.Left),
                "Independent native low/high controls move labels to right/left without moving the source crossing to that edge.");
        }
    }

    public static void PptxHorizontalCategoryTicksShareCrossingAndDirection()
    {
        foreach (bool reversed in new[] { false, true })
        {
            var actual = HorizontalGeometry(Render(-7, 25, reversed: reversed, horizontal: true));
            TestAssert.Equal(3, actual.Ticks.Length);
            foreach (double[] tick in actual.Ticks)
            {
                NearHorizontal(reversed ? Math.Min(tick[0], tick[2]) : Math.Max(tick[0], tick[2]), actual.Axis);
                TestAssert.True(reversed ? tick[0] > tick[2] : tick[0] < tick[2], "Outward category ticks follow numeric direction.");
            }
        }
    }

    public static void PptxHorizontalEndpointAdmissionFollowsPhysicalDirection()
    {
        PdfPage RenderHidden(bool reversed) => Render(7, 25, reversed: reversed, horizontal: true, editChart: xml =>
            xml.Descendants(CrossingC + "catAx").Single().Add(new XElement(CrossingC + "tickLblPos", new XAttribute("val", "none"))));
        var normal = HorizontalGeometry(RenderHidden(false));
        var reverse = HorizontalGeometry(RenderHidden(true));
        NearHorizontal(normal.Left - 72d, 792d - reverse.Right);
        NearHorizontal(reverse.Left - 72d, 792d - normal.Right);
        TestAssert.True(reverse.Left > normal.Left + 1d && reverse.Right > normal.Right + 1d,
            "Each physical edge admits its endpoint caption; reversal moves the longer103 caption to the left.");
    }

    public static void PptxHorizontalCategoryCrossingsKeepUnpairedAndComposedFallbacks()
    {
        foreach (string fallback in new[] { "unpaired", "composed", "logarithmic" })
        {
            var actual = HorizontalGeometry(Render(7, 25, reversed: true, horizontal: true, editChart: xml =>
            {
                if (fallback == "unpaired")
                {
                    xml.Descendants(CrossingC + "crossAx").Remove();
                    xml.Descendants(CrossingC + "catAx").Single().Element(CrossingC + "axPos")!.SetAttributeValue("val", "r");
                }
                if (fallback == "composed") xml.Descendants(CrossingC + "plotArea").Single().Add(new XElement(CrossingC + "lineChart"));
                if (fallback == "logarithmic") xml.Descendants(CrossingC + "valAx").Single().Element(CrossingC + "scaling")!
                    .Add(new XElement(CrossingC + "logBase", new XAttribute("val", "10")));
            }));
            NearHorizontal(actual.Axis, fallback == "unpaired" ? actual.Right : actual.Left);
            TestAssert.Equal(2, actual.Labels.Length);
            TestAssert.True(actual.Labels.All(x => x < actual.Left), "Unresolved placement keeps its existing caption fallback: " + fallback);
        }
    }

    public static void PptxHorizontalClusteredDataAndLabelsShareTheCategoryOrigin()
    {
        foreach (bool reversed in new[] { false, true })
        {
            PdfPage page = Render(7, 25, reversed: reversed, horizontal: true, editChart: xml =>
            {
                var crossing = xml.Descendants(CrossingC + "catAx").Single().Element(CrossingC + "crosses")!;
                crossing.Name = CrossingC + "crossesAt";
                crossing.SetAttributeValue("val", "40");
                xml.Descendants(CrossingC + "barChart").Single().Add(new XElement(CrossingC + "dLbls",
                    new XElement(CrossingC + "dLblPos", new XAttribute("val", "ctr")),
                    new XElement(CrossingC + "showVal", new XAttribute("val", "1"))));
            });
            var geometry = HorizontalGeometry(page);
            double origin = geometry.Left + (reversed ? 63d / 96d : 33d / 96d) * (geometry.Right - geometry.Left);
            var bars = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([\d.]+) ([\d.]+) re f")
                .Select(m => m.Groups.Cast<Group>().Skip(1).Select(g => double.Parse(g.Value, CultureInfo.InvariantCulture)).ToArray())
                .Where(r => r[2] is > 20d and < 700d && r[3] is > 20d and < 200d).ToArray();
            var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(2, bars.Length);
            TestAssert.Equal(2, labels.Length);
            for (int i = 0; i < bars.Length; i++)
            {
                NearHorizontal(reversed ? bars[i][0] + bars[i][2] : bars[i][0], origin);
                double center = labels[i].X + labels[i].Font.MeasureTextPoints(labels[i].Text, labels[i].Size, kerningEnabled: false) / 2d;
                TestAssert.True(center >= bars[i][0] && center <= bars[i][0] + bars[i][2],
                    "Data-label placement consumes the same source origin as its clustered bar.");
            }
        }
    }

    private static void NearHorizontal(double actual, double expected) => TestAssert.True(Math.Abs(actual - expected) < .002d,
        $"Expected horizontal coordinate {expected:R}, got {actual:R}.");

    public static void PptxHorizontalOriginRelativeInversionKeepsPaintAndSwatchesTogether()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (bool reversed in new[] { false, true })
        foreach (bool invert in new[] { false, true })
        {
            PdfPage page = Render(7, 25, reversed: reversed, horizontal: true, editChart: xml =>
            {
                xml.Descendants(CrossingC + "catAx").Single().Element(CrossingC + "crosses")!.SetAttributeValue("val", "max");
                var series = xml.Descendants(CrossingC + "ser").Single();
                series.Add(new XElement(CrossingC + "spPr", new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "3366CC")))),
                    new XElement(CrossingC + "invertIfNegative", new XAttribute("val", invert ? "1" : "0")));
                xml.Descendants(CrossingC + "barChart").Single().Add(new XElement(CrossingC + "dLbls",
                    new XElement(CrossingC + "dLblPos", new XAttribute("val", "ctr")),
                    new XElement(CrossingC + "showVal", new XAttribute("val", "1")),
                    new XElement(CrossingC + "showSerName", new XAttribute("val", "0")),
                    new XElement(CrossingC + "showCatName", new XAttribute("val", "0")),
                    new XElement(CrossingC + "showLegendKey", new XAttribute("val", "1"))));
            });
            TestAssert.Equal(invert ? 0 : 4, PptxChartNegativeInversionTests.CountSeriesColorPaint(page.Content));
            TestAssert.Equal("82,96", string.Join(",", DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).Select(p => p.Text)));
        }
    }

    private static (double Left, double Right, double Axis, double[] Labels, double[][] Ticks) HorizontalGeometry(PdfPage page)
    {
        var lines = Lines(page);
        var numeric = lines.Single(p => p.Color == new RgbColor(255, 0, 0) && Math.Abs(p.V[1] - p.V[3]) < .001d && Math.Abs(p.V[0] - p.V[2]) > 100d).V;
        var category = lines.Single(p => p.Color == new RgbColor(0, 170, 0) && Math.Abs(p.V[0] - p.V[2]) < .001d && Math.Abs(p.V[1] - p.V[3]) > 100d).V;
        var ticks = lines.Where(p => p.Color == new RgbColor(0, 170, 0) && Math.Abs(p.V[1] - p.V[3]) < .001d && Math.Abs(p.V[0] - p.V[2]) is > 0d and < 20d).Select(p => p.V).ToArray();
        var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).Select(p => p.X).ToArray();
        return (Math.Min(numeric[0], numeric[2]), Math.Max(numeric[0], numeric[2]), category[0], labels, ticks);
    }
}
