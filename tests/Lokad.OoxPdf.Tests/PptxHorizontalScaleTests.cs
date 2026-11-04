using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxHorizontalScaleTests
{
    private static string Cases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "HorizontalSourceScales"));

    public static void PptxHorizontalSourceScalesMatchIndependentOfficeLabels()
    {
        using var oracle = ReadOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(Render(row), usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal))
                .Select(show => show.Text).ToArray();
            var expected = row.GetProperty("OfficeNumericLabels").EnumerateArray().Select(value => value.GetString()!).ToArray();
            TestAssert.True(actual.SequenceEqual(expected),
                row.GetProperty("Id").GetString() + ": native numeric labels; actual " + string.Join(",", actual));
        }
    }

    public static void PptxHorizontalPlotsUseOneNativeScaleForEveryBar()
    {
        using var oracle = ReadOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            double[] values = row.GetProperty("BarValues").EnumerateArray().Select(value => value.GetDouble()).ToArray();
            if (values.Length == 0) continue;
            PdfPage page = Render(row);
            bool horizontal = row.GetProperty("Horizontal").GetBoolean();
            var axis = Axis(page, horizontal);
            double minimum = row.GetProperty("Minimum").GetDouble(), maximum = row.GetProperty("Maximum").GetDouble();
            double baseline = Math.Clamp(0d, minimum, maximum);
            double Position(double value)
            {
                double ratio = (value - minimum) / (maximum - minimum);
                return axis.Start + (row.GetProperty("Reversed").GetBoolean() ? 1d - ratio : ratio) * axis.Length;
            }
            var rectangles = Regex.Matches(page.Content, @"([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re f\s")
                .Select(match => match.Groups.Cast<Group>().Skip(1).Select(group => double.Parse(group.Value, CultureInfo.InvariantCulture)).ToArray())
                .Where(rect => horizontal ? Math.Abs(rect[2]) > 40d && Math.Abs(rect[3]) is > 1d and < 80d
                    : Math.Abs(rect[2]) is > 1d and < 80d && Math.Abs(rect[3]) > 40d).ToArray();
            TestAssert.Equal(values.Length, rectangles.Length);
            for (int index = 0; index < values.Length; index++)
            {
                double start = rectangles[index][horizontal ? 0 : 1];
                double end = start + rectangles[index][horizontal ? 2 : 3];
                double origin = Position(baseline), value = Position(values[index]);
                TestAssert.True(Math.Abs(Math.Min(start, end) - Math.Min(origin, value)) < .002d &&
                    Math.Abs(Math.Max(start, end) - Math.Max(origin, value)) < .002d,
                    row.GetProperty("Id").GetString() + ": every bar uses the native scale; source value " + values[index]);
            }
        }
    }

    public static void PptxHorizontalNumericMarksAndGridsKeepNativeIntervals()
    {
        using var oracle = ReadOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray()
            .Where(row => row.GetProperty("Horizontal").GetBoolean()))
        {
            PdfPage page = Render(row);
            var axis = Axis(page, horizontal: true);
            double minimum = row.GetProperty("Minimum").GetDouble(), maximum = row.GetProperty("Maximum").GetDouble();
            double Value(double x)
            {
                double ratio = (x - axis.Start) / axis.Length;
                return minimum + (row.GetProperty("Reversed").GetBoolean() ? 1d - ratio : ratio) * (maximum - minimum);
            }
            var lines = PptxColumnIntervalTests.Lines(page);
            double[] expected = row.GetProperty("OfficeNumericLabels").EnumerateArray()
                .Select(value => double.Parse(value.GetString()!, CultureInfo.InvariantCulture)).Order().ToArray();
            double[] marks = lines.Where(line => line.Color == new RgbColor(255, 0, 0) &&
                Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) is > 0d and < 20d)
                .Select(line => Value(line.V[0])).Order().ToArray();
            Check(marks, expected, "marks");
            double[] grids = lines.Where(line => line.Color != new RgbColor(255, 0, 0) && line.Color != new RgbColor(0, 170, 0) &&
                Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) > 100d)
                .Select(line => Value(line.V[0])).Distinct().Order().ToArray();
            double crossing = Math.Clamp(0d, minimum, maximum);
            Check(grids, expected.Where(value => Math.Abs(value - crossing) > .001d).ToArray(), "grids");

            void Check(double[] actual, double[] desired, string role) => TestAssert.True(actual.Length == desired.Length &&
                actual.Zip(desired).All(pair => Math.Abs(pair.First - pair.Second) < .002d),
                row.GetProperty("Id").GetString() + ": native " + role + "; actual " + string.Join(",", actual));
        }
    }

    private static (double Start, double Length) Axis(PdfPage page, bool horizontal)
    {
        var line = PptxColumnIntervalTests.Lines(page).Single(line => line.Color == new RgbColor(255, 0, 0) &&
            (horizontal ? Math.Abs(line.V[1] - line.V[3]) < .001d && Math.Abs(line.V[0] - line.V[2]) > 100d
                : Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) > 100d)).V;
        return horizontal ? (Math.Min(line[0], line[2]), Math.Abs(line[2] - line[0]))
            : (Math.Min(line[1], line[3]), Math.Abs(line[3] - line[1]));
    }

    private static JsonDocument ReadOracle() => JsonDocument.Parse(File.ReadAllText(Path.Combine(Cases, "office.json")));

    private static PdfPage Render(JsonElement row)
    {
        string path = Path.Combine(Cases, row.GetProperty("Input").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        using var package = OoxPackage.Open(path, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
