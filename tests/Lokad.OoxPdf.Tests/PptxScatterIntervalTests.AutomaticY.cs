using System.Text.Json;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxScatterAutomaticYKeepsNativeClusteredIntervals()
    {
        // The independent clustered-scatter reference has Y0..6/unit1.
        // Inferring from ten intervals instead of the shared nine chooses0.5,
        // shrinks the range and shifts the measured side strip.
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases",
            "pptx-ladder-11-chart-scatter-clusters-port.pptx"));
        using var stream = File.OpenRead(path);
        using var package = OoxPackage.Open(stream, default);
        PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver()).RenderPagesAsync(
            new PptxReader().Read(package, default), package, null, default).ToBlockingEnumerable().Single();
        TestAssert.Equal("0,1,2,3,4,5,6,7,8,0,1,2,3,4,5,6", string.Join(",", AutoLabels(page)));
    }

    public static void PptxScatterAutomaticYMatchesIndependentOfficeBoundsAndUnits()
    {
        foreach (RangeCase sample in AutomaticYCases())
        {
            var (extents, majorUnit, tickPolicy) = ResolveRange(sample);
            TestAssert.Equal(new PptxRenderer.ChartValueExtents(sample.ExpectedMinimum, sample.ExpectedMaximum), extents);
            TestAssert.Equal(sample.ExpectedUnit, majorUnit);
            TestAssert.Equal("SourceMinimumIntervals", tickPolicy);
        }
    }

    public static void PptxScatterAutomaticYLabelsMarksAndGridsShareNativeIntervals()
    {
        foreach (RangeCase sample in AutomaticYCases())
        {
            PdfPage page = RenderRange(sample);
            double unit = sample.ExpectedUnit!.Value;
            double[] values = Enumerable.Range(0, (int)Math.Round((sample.ExpectedMaximum - sample.ExpectedMinimum) / unit) + 1)
                .Select(i => sample.ExpectedMinimum + unit * i).ToArray();
            AssertPaintedYIntervals(page, sample.ExpectedMinimum, sample.ExpectedMaximum, values);
        }
    }

    private static void AssertPaintedYIntervals(PdfPage page, double minimum, double maximum, double[] values)
    {
        TestAssert.Equal("0,1,2,3,4,5,6,7,8,9," + string.Join(",", values), string.Join(",", AutoLabels(page)));
        var axes = Axes(page);
        double Y(double value) => axes.Bottom + (value - minimum) / (maximum - minimum) * (axes.Top - axes.Bottom);
        var lines = PptxColumnIntervalTests.Lines(page).Where(line => line.Color == new RgbColor(0, 0, 255) &&
            Math.Abs(line.V[1] - line.V[3]) < .001d).ToArray();
        double[] ticks = lines.Where(line => Math.Abs(line.V[0] - line.V[2]) < 20d).Select(line => line.V[1]).ToArray();
        double[] grids = lines.Where(line => Math.Abs(line.V[0] - line.V[2]) > 100d).Select(line => line.V[1]).ToArray();
        Check(ticks, values.Select(Y).ToArray());
        Check(grids, values.Where(value => value != 0d).Select(Y).ToArray());

        static void Check(double[] actual, double[] expected) => TestAssert.True(actual.Length == expected.Length &&
            actual.Zip(expected).All(pair => Math.Abs(pair.First - pair.Second) < .002d),
            "Y intervals must agree in painted coordinates: " + string.Join(",", actual));
    }

    public static void PptxScatterAutomaticYRetainsSceneFallbackAndCancellationOwnership()
    {
        RangeCase sample = AutomaticYCases().Single(c => c.Id == "range-scatter-103-auto");
        AssertFallbackRange(sample with { MajorUnit = double.NaN }, minimum: 0d, maximum: 103d, majorUnit: 20d);
        AssertFallbackRange(sample with { MajorUnit = -1d }, minimum: 0d, maximum: 103d, majorUnit: 20d);
        AssertFallbackRange(sample with { Minimum = double.NaN }, minimum: 0d, maximum: 79d, majorUnit: 10d);
        AssertFallbackRange(sample, minimum: 0d, maximum: 103d, majorUnit: 20d, logarithmic: true);
        AssertFallbackRange(sample, minimum: 0d, maximum: 103d, majorUnit: 20d, bubble: true);
        AssertFallbackRange(sample, minimum: 0d, maximum: 79d, majorUnit: 10d, includeScaling: false);

        // Emission selects an effective finite interval even when the scene
        // carries the authored infinite unit.
        RangeCase infiniteUnit = sample with { MajorUnit = double.PositiveInfinity };
        AssertFallbackRange(infiniteUnit, minimum: 0d, maximum: 103d, majorUnit: null);
        AssertPaintedYIntervals(RenderRange(infiniteUnit), 0d, 103d, [0d, 20d, 40d, 60d, 80d, 100d, 103d]);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => ResolveRange(sample, cancellationToken: cancellation.Token));
    }

    private static RangeCase[] AutomaticYCases()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "scatter-automatic-y-office.json"));
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return JsonSerializer.Deserialize<RangeCase[]>(document.RootElement.GetProperty("Cases").GetRawText())!;
    }
}
