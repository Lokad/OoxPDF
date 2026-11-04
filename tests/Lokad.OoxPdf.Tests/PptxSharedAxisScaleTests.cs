using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxSharedAxisScaleTests
{
    internal static void CheckFrames()
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "SharedCartesianFrames"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office.json")));
        foreach (var pair in oracle.RootElement.GetProperty("Rows").EnumerateArray()
            .GroupBy(row => row.GetProperty("Pair").GetString()))
        {
            var frames = new List<double[]>();
            foreach (var row in pair)
            {
                string path = Path.Combine(cases, row.GetProperty("Input").GetString()!);
                TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
                using var package = OoxPackage.Open(path, default);
                var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
                    new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
                frames.Add(PptxColumnIntervalTests.Lines(page).Single(line => line.Color == new RgbColor(255, 0, 0) &&
                    Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) > 100d).V);
            }
            TestAssert.Equal(2, frames.Count);
            for (int axis = 0; axis < 4; axis++)
                TestAssert.True(Math.Abs(frames[0][axis] - frames[1][axis]) < .002d,
                    "A shared reciprocal scale retains the single plot frame: " + pair.Key + "/" + axis);
        }
    }

    internal static void Check(bool geometry)
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "SharedCartesianAxes"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "office.json")));
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string path = Path.Combine(cases, row.GetProperty("Input").GetString()!);
            TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
            using var package = OoxPackage.Open(path, default);
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
                new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
            string id = row.GetProperty("Id").GetString()!;
            if (!geometry)
            {
                var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                    fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).Select(show => show.Text).ToArray();
                var expected = row.GetProperty("OfficeNumericLabels").EnumerateArray().Select(value => value.GetString()!).ToArray();
                TestAssert.True(expected.SequenceEqual(actual), "Shared source range/intervals: " + id + "; actual " + string.Join(",", actual));
                continue;
            }
            var axis = PptxColumnIntervalTests.Lines(page).Single(line => line.Color == new RgbColor(255, 0, 0) &&
                Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) > 100d).V;
            double bottom = Math.Min(axis[1], axis[3]), height = Math.Abs(axis[3] - axis[1]);
            double minimum = row.GetProperty("Minimum").GetDouble(), maximum = row.GetProperty("Maximum").GetDouble();
            bool reversed = row.GetProperty("Reversed").GetBoolean();
            double Y(double value)
            {
                double ratio = (value - minimum) / (maximum - minimum);
                return bottom + (reversed ? 1d - ratio : ratio) * height;
            }
            double Parse(string value) => double.Parse(value, CultureInfo.InvariantCulture);
            var rectangles = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re f\s")
                .Select(match => match.Groups.Cast<Group>().Skip(1).Select(group => Parse(group.Value)).ToArray()).ToArray();
            double[] primary = row.GetProperty("PrimaryValues").EnumerateArray().Select(value => value.GetDouble()).ToArray();
            double[] other = row.GetProperty("OtherValues").EnumerateArray().Select(value => value.GetDouble()).ToArray();
            bool columns = row.GetProperty("OtherIsColumn").GetBoolean();
            double[] bars = columns ? [.. primary, .. other] : primary;
            TestAssert.Equal(bars.Length, rectangles.Length);
            for (int index = 0; index < bars.Length; index++)
            {
                var rect = rectangles[index];
                double end = reversed ? Math.Min(rect[1], rect[1] + rect[3]) : Math.Max(rect[1], rect[1] + rect[3]);
                TestAssert.True(Math.Abs(end - Y(bars[index])) < .002d, "Column uses the shared numeric axis: " + id + "/" + index);
            }
            if (columns || other.Length == 0) continue;
            // Source data anchors may be joined by straight or cubic segments.
            // This checks their common scale, separately from smoothing and X placement.
            var paths = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) m\s+((?:(?:[-\d.]+ ){1,5}[-\d.]+ [lc]\s+){7})S");
            TestAssert.Equal(1, paths.Count);
            var pointYs = new List<double> { Parse(paths[0].Groups[2].Value) };
            foreach (var line in paths[0].Groups[3].Value.Split('\n', StringSplitOptions.RemoveEmptyEntries))
                pointYs.Add(Parse(line.Trim().Split(' ')[^2]));
            TestAssert.Equal(other.Length, pointYs.Count);
            for (int index = 0; index < other.Length; index++)
                TestAssert.True(Math.Abs(pointYs[index] - Y(other[index])) < .002d,
                    "Line uses the same numeric axis as columns: " + id + "/" + index);
        }
    }
}
