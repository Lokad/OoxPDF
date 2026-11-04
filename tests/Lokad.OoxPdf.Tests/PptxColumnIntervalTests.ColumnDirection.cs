using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    private static string ColumnDirectionCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "ColumnDirection"));
    private static JsonDocument ColumnDirectionOracle() => JsonDocument.Parse(File.ReadAllText(Path.Combine(ColumnDirectionCases, "office.json")));

    public static void PptxClusteredColumnDirectionMatchesIndependentOfficeSlots()
    {
        using var oracle = ColumnDirectionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var page = RenderAdmission(row, caseDirectory: ColumnDirectionCases);
            var plot = AdmissionBounds(page);
            foreach (var expected in row.GetProperty("Bars").EnumerateArray())
            {
                double[] bar = ColumnDirectionBar(page, expected);
                double ratio = (bar[0] + bar[2] / 2d - plot.Left) / (plot.Right - plot.Left);
                TestAssert.True(Math.Abs(ratio - expected.GetProperty("CenterFraction").GetDouble()) < .001d,
                    row.GetProperty("Id").GetString() + ": source index, series offset, sparse slot and point color match the independent native column.");
            }
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.True(captions.Select(show => show.Text).SequenceEqual(new[] { "N1", "N2", "N3", "N4" }),
                "Source order remains intact while physical category placement changes.");
            bool reversed = row.GetProperty("Source").GetProperty("Reversed").GetBoolean();
            TestAssert.True(reversed ? captions[0].X > captions[3].X : captions[0].X < captions[3].X,
                "Captions and painted columns share physical source direction.");
        }
    }

    public static void PptxClusteredColumnLabelsAndLinksFollowTheirPaintedSourcePoints()
    {
        using var oracle = ColumnDirectionOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray().Where(row =>
            !row.GetProperty("Source").TryGetProperty("Keys", out _)))
        {
            var page = RenderAdmission(row, caseDirectory: ColumnDirectionCases);
            var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).ToArray();
            TestAssert.Equal(row.GetProperty("Bars").GetArrayLength(), labels.Length);
            foreach (var expected in row.GetProperty("Bars").EnumerateArray())
            {
                double[] bar = ColumnDirectionBar(page, expected);
                string value = expected.GetProperty("Value").GetInt32().ToString(CultureInfo.InvariantCulture);
                var label = labels.Single(show => show.Text == value);
                double center = label.X + label.Font.MeasureTextPoints(label.Text, label.Size) / 2d;
                TestAssert.True(Math.Abs(center - bar[0] - bar[2] / 2d) < .02d,
                    "Plain labels use the actual clustered-column center, including gap/overlap and reversal.");
            }
        }
        var linked = oracle.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "column-reversed-three-explicit");
        var linkedPage = RenderAdmission(linked, link: true, caseDirectory: ColumnDirectionCases);
        double[] target = ColumnDirectionBar(linkedPage, linked.GetProperty("Bars").EnumerateArray().Single(bar =>
            bar.GetProperty("Series").GetInt32() == 0 && bar.GetProperty("Index").GetInt32() == 0));
        var annotation = linkedPage.Annotations.Single();
        TestAssert.Equal("https://example.invalid/manual-caption", annotation.Uri);
        TestAssert.True(Math.Abs(annotation.X + annotation.Width / 2d - target[0] - target[2] / 2d) < .02d,
            "The rich label link follows source point zero into the reversed cluster.");
    }

    public static void PptxClusteredColumnLegendKeyUnitsFollowTheirPaintedBars()
    {
        using var oracle = ColumnDirectionOracle();
        var row = oracle.RootElement.GetProperty("Rows").EnumerateArray().Single(row =>
            row.GetProperty("Id").GetString() == "column-reversed-keys-explicit");
        var page = RenderAdmission(row, caseDirectory: ColumnDirectionCases);
        var labels = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
            fontResourceFilter: name => name.StartsWith("CBD", StringComparison.Ordinal)).ToArray();
        TestAssert.Equal(12, labels.Length);
        foreach (var expected in row.GetProperty("Bars").EnumerateArray())
        {
            double[] bar = ColumnDirectionBar(page, expected);
            string value = expected.GetProperty("Value").GetInt32().ToString(CultureInfo.InvariantCulture);
            var label = labels.Single(show => show.Text == value);
            double[] key = ColumnDirectionRectangles(page, expected.GetProperty("Color")).Where(rect =>
                rect[2] < 20d && rect[3] < 20d && rect[0] < label.X && label.X - rect[0] < 30d &&
                Math.Abs(rect[1] + rect[3] / 2d - label.Y) < label.Size).Single();
            double end = label.X + label.Font.MeasureTextPoints(label.Text, label.Size);
            TestAssert.True(Math.Abs((key[0] + end) / 2d - bar[0] - bar[2] / 2d) <= 1d,
                "Each source-colored key and text unit follows its own reflected column.");
        }
    }

    public static void PptxClusteredColumnDirectionPreservesExplicitFallbacks()
    {
        using var oracle = ColumnDirectionOracle();
        var row = oracle.RootElement.GetProperty("Rows").EnumerateArray().Single(row => row.GetProperty("Id").GetString() == "column-reversed-three-explicit");
        foreach (string fallback in new[] { "stacked", "unpaired", "composed" })
        {
            var page = RenderAdmission(row, caseDirectory: ColumnDirectionCases, editChart: xml =>
            {
                if (fallback == "stacked") xml.Descendants(CrossingC + "grouping").Single().SetAttributeValue("val", "stacked");
                if (fallback == "unpaired") xml.Descendants(CrossingC + "crossAx").Remove();
                if (fallback == "composed") xml.Descendants(CrossingC + "plotArea").Single().Add(new XElement(CrossingC + "lineChart"));
            });
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CCA", StringComparison.Ordinal)).ToArray();
            TestAssert.True(captions[0].X < captions[3].X, "Unqualified column source direction keeps its declared fallback: " + fallback);
        }
    }

    private static double[] ColumnDirectionBar(PdfPage page, JsonElement expected)
    {
        var plot = AdmissionBounds(page);
        double value = expected.GetProperty("Value").GetDouble();
        return ColumnDirectionRectangles(page, expected.GetProperty("Color")).Where(rect => rect[2] > 3d && rect[3] > 20d &&
            Math.Abs(rect[3] / (plot.Bottom - plot.Top) * 120d - value) < .01d).Single();
    }

    private static double[][] ColumnDirectionRectangles(PdfPage page, JsonElement color)
    {
        int[] expected = color.EnumerateArray().Select(channel => channel.GetInt32()).ToArray();
        string? command = page.Content.Split('\n').Select(line => line.Trim()).Where(line => line.EndsWith(" rg", StringComparison.Ordinal))
            .FirstOrDefault(line => line[..^3].Split(' ').Select(value => (int)Math.Round(double.Parse(value, CultureInfo.InvariantCulture) * 255d)).SequenceEqual(expected));
        return command is null ? [] : LabelGroupRectangles(page, command);
    }
}
