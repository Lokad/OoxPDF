using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxHorizontalFrameTests
{
    private static string NumericSideCases => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", "HorizontalNumericSides"));

    public static void PptxHorizontalNumericFramesFollowIndependentOfficeCaptionBands()
    {
        using var oracle = NumericSideOracle();
        var rows = oracle.RootElement.GetProperty("Rows").EnumerateArray().ToDictionary(row => row.GetProperty("Id").GetString()!);
        foreach (var row in rows.Values)
        {
            var page = RenderNumericSide(row);
            var actual = NumericSideGeometry(page);
            double[] expected = row.GetProperty("NativeFrameY").EnumerateArray().Select(value => value.GetDouble()).ToArray();
            if (!row.GetProperty("Furniture").GetBoolean())
            {
                Near(actual.Top, expected[0]); Near(actual.Bottom, expected[1]);
            }
            else if (row.GetProperty("RelativeTo").GetString() is { } normal)
            {
                var baseline = NumericSideGeometry(RenderNumericSide(rows[normal]));
                double[] nativeBaseline = rows[normal].GetProperty("NativeFrameY").EnumerateArray().Select(value => value.GetDouble()).ToArray();
                // Pinned portable faces have different title/legend metrics.
                // Check the independently witnessed band transfer while their
                // absolute Windows/Office comparisons retain the same gates.
                Near(actual.Top - baseline.Top, expected[0] - nativeBaseline[0]);
                Near(actual.Bottom - baseline.Bottom, expected[1] - nativeBaseline[1]);
            }
            void Near(double value, double target) => TestAssert.True(Math.Abs(value - target) <= 1d,
                row.GetProperty("Id").GetString() + ": native caption-band geometry; actual " + value.ToString("F3", CultureInfo.InvariantCulture));
        }
    }

    public static void PptxHorizontalNumericStrokesAndTicksFollowCategoryDirection()
    {
        using var oracle = NumericSideOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            PdfPage page = RenderNumericSide(row);
            var geometry = NumericSideGeometry(page);
            bool top = row.GetProperty("NativeAxisTop").GetBoolean();
            double target = top ? geometry.Top : geometry.Bottom;
            TestAssert.True(Math.Abs(geometry.Axis - target) < .002d,
                row.GetProperty("Id").GetString() + ": native numeric stroke side.");
            var ticks = PptxColumnIntervalTests.Lines(page).Where(line => line.Color == new RgbColor(255, 0, 0) &&
                Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) is > 0d and < 20d).ToArray();
            TestAssert.Equal(row.GetProperty("NativeNumericLabels").GetArrayLength(), ticks.Length);
            foreach (var tick in ticks)
            {
                double physicalTop = page.Height - Math.Max(tick.V[1], tick.V[3]);
                double physicalBottom = page.Height - Math.Min(tick.V[1], tick.V[3]);
                TestAssert.True(top ? Math.Abs(physicalBottom - target) < .002d && physicalTop < target
                    : Math.Abs(physicalTop - target) < .002d && physicalBottom > target,
                    row.GetProperty("Id").GetString() + ": outward numeric ticks follow the stroke side.");
            }
        }
    }

    public static void PptxHorizontalNumericCaptionsKeepNativeSideAndCompleteLabels()
    {
        using var oracle = NumericSideOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            PdfPage page = RenderNumericSide(row);
            var geometry = NumericSideGeometry(page);
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(page, usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CVA", StringComparison.Ordinal)).ToArray();
            TestAssert.True(captions.Select(show => show.Text).SequenceEqual(
                row.GetProperty("NativeNumericLabels").EnumerateArray().Select(value => value.GetString()!)),
                row.GetProperty("Id").GetString() + ": native complete numeric text.");
            bool above = row.GetProperty("NativeCaptionSide").GetString() == "above";
            TestAssert.True(captions.All(show => above ? page.Height - show.Y < geometry.Top : page.Height - show.Y > geometry.Bottom),
                row.GetProperty("Id").GetString() + ": numeric captions retain their native side independently of the stroke.");
        }
    }

    public static void PptxHorizontalComposedLegendsRetainEveryNativeSeriesName()
    {
        using var oracle = NumericSideOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            string[] expected = row.GetProperty("NativeLegendNames").EnumerateArray().Select(value => value.GetString()!).ToArray();
            if (expected.Length == 0) continue;
            var actual = DocxCommentsTests.ReadEmbeddedGlyphTextShows(RenderNumericSide(row), usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).Select(show => show.Text).ToArray();
            TestAssert.True(actual.Order(StringComparer.Ordinal).SequenceEqual(expected.Order(StringComparer.Ordinal)),
                row.GetProperty("Id").GetString() + ": every native legend series survives measurement and wrapping; actual " + string.Join(",", actual));
        }
    }

    public static void PptxWrappedFilledSeriesLegendsKeepIndependentOfficeColumns()
    {
        using var oracle = NumericSideOracle();
        foreach (var row in oracle.RootElement.GetProperty("Rows").EnumerateArray())
        {
            int nativeRows = row.GetProperty("NativeLegendRows").GetInt32();
            if (nativeRows < 2 || nativeRows != row.GetProperty("NativeLegendNames").GetArrayLength()) continue;
            var captions = DocxCommentsTests.ReadEmbeddedGlyphTextShows(RenderNumericSide(row), usePageCoordinates: false,
                fontResourceFilter: name => name.StartsWith("CL", StringComparison.Ordinal)).ToArray();
            // These native rows form one column. The pinned face is wider,
            // retaining wrapping without fitting a font-specific break width.
            TestAssert.True(captions.Select(show => show.Y).Distinct().Count() > 1 &&
                captions.All(show => Math.Abs(show.X - captions[0].X) < .002d),
                row.GetProperty("Id").GetString() + ": filled series rows preserve their native shared column.");
        }
    }

    private static (double Top, double Bottom, double Axis) NumericSideGeometry(PdfPage page)
    {
        var lines = PptxColumnIntervalTests.Lines(page);
        var red = lines.Single(line => line.Color == new RgbColor(255, 0, 0) &&
            Math.Abs(line.V[1] - line.V[3]) < .001d && Math.Abs(line.V[0] - line.V[2]) > 100d).V;
        var green = lines.Single(line => line.Color == new RgbColor(0, 170, 0) &&
            Math.Abs(line.V[0] - line.V[2]) < .001d && Math.Abs(line.V[1] - line.V[3]) > 100d).V;
        return (page.Height - Math.Max(green[1], green[3]), page.Height - Math.Min(green[1], green[3]), page.Height - red[1]);
    }

    private static JsonDocument NumericSideOracle() => JsonDocument.Parse(File.ReadAllText(Path.Combine(NumericSideCases, "office.json")));

    private static PdfPage RenderNumericSide(JsonElement row)
    {
        string path = Path.Combine(NumericSideCases, row.GetProperty("Input").GetString()!);
        TestAssert.Equal(row.GetProperty("InputSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))));
        using var package = OoxPackage.Open(path, default);
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
    }
}
