using System.Globalization;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    public static void PptxOnTickColumnsKeepSourceClustersAndGaps()
    {
        using var oracle = ColumnDirectionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var before = RenderAdmission(row, caseDirectory: ColumnDirectionCases,
                editChart: chart => SetColumnCrossBetween(chart, "between"));
            var after = RenderAdmission(row, caseDirectory: ColumnDirectionCases,
                editChart: chart => SetColumnCrossBetween(chart, "midCat"));
            var a = AdmissionBounds(before); var b = AdmissionBounds(after);
            int count = row.GetProperty("Bars").EnumerateArray().Max(bar => bar.GetProperty("Index").GetInt32()) + 1;
            foreach (var expected in row.GetProperty("Bars").EnumerateArray())
            {
                double[] oldBar = ColumnDirectionBar(before, expected), bar = ColumnDirectionBar(after, expected);
                double oldRatio = (oldBar[0] + oldBar[2] / 2d - a.Left) / (a.Right - a.Left);
                double ratio = (bar[0] + bar[2] / 2d - b.Left) / (b.Right - b.Left);
                double scale = (double)count / (count - 1);
                TestAssert.True(Math.Abs(ratio - (oldRatio - .5d / count) * scale) < .00005d &&
                    Math.Abs(bar[2] / (b.Right - b.Left) - oldBar[2] / (a.Right - a.Left) * scale) < .00005d,
                    row.GetProperty("Id").GetString() + ": endpoint slots preserve source clusters, sparse indices, overlap, gap and point paint.");
            }
        }

        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray().Where(row =>
            row.GetProperty("Id").GetString() is "column-normal-three-explicit" or "column-reversed-three-explicit"))
        {
            var before = RenderAdmission(row, caseDirectory: ColumnDirectionCases, editChart: chart => SingleCategory(chart, "between"));
            var after = RenderAdmission(row, caseDirectory: ColumnDirectionCases, editChart: chart => SingleCategory(chart, "midCat"));
            var a = AdmissionBounds(before); var b = AdmissionBounds(after);
            bool reversed = row.GetProperty("Source").GetProperty("Reversed").GetBoolean();
            foreach (var expected in row.GetProperty("Bars").EnumerateArray().Where(bar => bar.GetProperty("Index").GetInt32() == 0))
            {
                double[] oldBar = ColumnDirectionBar(before, expected), bar = ColumnDirectionBar(after, expected);
                double oldRatio = (oldBar[0] + oldBar[2] / 2d - a.Left) / (a.Right - a.Left);
                double ratio = (bar[0] + bar[2] / 2d - b.Left) / (b.Right - b.Left);
                TestAssert.True(Math.Abs(ratio - oldRatio - (reversed ? .5d : -.5d)) < .00005d &&
                    Math.Abs(bar[2] / (b.Right - b.Left) - oldBar[2] / (a.Right - a.Left)) < .00005d,
                    "A sole source category keeps finite cluster widths and moves to its ordinal endpoint.");
            }
        }
    }

    public static void PptxOnTickColumnLabelsAndLinksFollowPaintedClusters()
    {
        using var oracle = ColumnDirectionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray().Where(row =>
            !row.GetProperty("Source").TryGetProperty("Keys", out _)))
        {
            var page = RenderAdmission(row, caseDirectory: ColumnDirectionCases,
                editChart: chart => SetColumnCrossBetween(chart, "midCat"));
            var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(row.GetProperty("Bars").GetArrayLength(), labels.Length);
            foreach (var expected in row.GetProperty("Bars").EnumerateArray())
            {
                double[] bar = ColumnDirectionBar(page, expected);
                string value = expected.GetProperty("Value").GetInt32().ToString(CultureInfo.InvariantCulture);
                var label = labels.Single(show => show.Text == value);
                TestAssert.True(Math.Abs(label.X + label.Font.MeasureTextPoints(label.Text, label.Size) / 2d - bar[0] - bar[2] / 2d) < .02d,
                    "On-tick labels follow the source bar including intra-cluster reflection and sparse points.");
            }
        }
        var linked = oracle.RootElement.GetProperty("Rows").EnumerateArray().Single(row =>
            row.GetProperty("Id").GetString() == "column-reversed-three-explicit");
        var linkedPage = RenderAdmission(linked, link: true, caseDirectory: ColumnDirectionCases, editChart: chart =>
        {
            SetColumnCrossBetween(chart, "midCat");
            chart.Descendants(CrossingC + "dLbl").Single().Add(new XElement(CrossingC + "showLegendKey", new XAttribute("val", "0")));
        });
        double[] target = ColumnDirectionBar(linkedPage, linked.GetProperty("Bars").EnumerateArray().Single(bar =>
            bar.GetProperty("Series").GetInt32() == 0 && bar.GetProperty("Index").GetInt32() == 0));
        var annotation = linkedPage.Annotations.Single();
        TestAssert.Equal("https://example.invalid/manual-caption", annotation.Uri);
        TestAssert.True(Math.Abs(annotation.X + annotation.Width / 2d - target[0] - target[2] / 2d) < .02d,
            "The on-tick rich-label link follows its source column.");
    }

    private static void SetColumnCrossBetween(XDocument chart, string value)
    {
        var axis = chart.Descendants(CrossingC + "valAx").Single();
        axis.Elements(CrossingC + "crossBetween").Remove();
        axis.Add(new XElement(CrossingC + "crossBetween", new XAttribute("val", value)));
    }

    private static void SingleCategory(XDocument chart, string value)
    {
        SetColumnCrossBetween(chart, value);
        foreach (var series in chart.Descendants(CrossingC + "ser"))
            foreach (string name in new[] { "cat", "val" })
            {
                var source = series.Element(CrossingC + name)!;
                source.Descendants(CrossingC + "pt").Where(point => (int)point.Attribute("idx")! != 0).Remove();
                foreach (var count in source.Descendants(CrossingC + "ptCount")) count.SetAttributeValue("val", "1");
            }
    }
}
