using System.Text.Json;
using System.Xml.Linq;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxScatterIntervalTests
{
    public static void PptxScatterAutomaticYMatchesIndependentOfficeBoundsAndUnits()
    {
        foreach (RangeCase sample in AutomaticYCases())
        foreach (bool scene in new[] { false, true })
        {
            var (extents, options) = ResolveRange(sample, scene);
            TestAssert.Equal(new PptxRenderer.ChartValueExtents(sample.ExpectedMinimum, sample.ExpectedMaximum), extents);
            object units = options.GetType().GetProperty("Units")!.GetValue(options)!;
            TestAssert.Equal(sample.ExpectedUnit, (double?)units.GetType().GetProperty("MajorUnit")!.GetValue(units));
            TestAssert.Equal("SourceMinimumIntervals", options.GetType().GetProperty("TickPolicy")!.GetValue(options)!.ToString());
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
            string[] labels = AutoLabels(page);
            TestAssert.Equal("0,1,2,3,4,5,6,7,8,9," + string.Join(",", values), string.Join(",", labels));
            var axes = Axes(page);
            double Y(double value) => axes.Bottom + (value - sample.ExpectedMinimum) /
                (sample.ExpectedMaximum - sample.ExpectedMinimum) * (axes.Top - axes.Bottom);
            var lines = PptxColumnIntervalTests.Lines(page).Where(line => line.Color == new RgbColor(0, 0, 255) &&
                Math.Abs(line.V[1] - line.V[3]) < .001d).ToArray();
            double[] ticks = lines.Where(line => Math.Abs(line.V[0] - line.V[2]) < 20d).Select(line => line.V[1]).ToArray();
            double[] grids = lines.Where(line => Math.Abs(line.V[0] - line.V[2]) > 100d).Select(line => line.V[1]).ToArray();
            Check(ticks, values.Select(Y).ToArray());
            Check(grids, values.Where(value => value != 0d).Select(Y).ToArray());
        }

        static void Check(double[] actual, double[] expected) => TestAssert.True(actual.Length == expected.Length &&
            actual.Zip(expected).All(pair => Math.Abs(pair.First - pair.Second) < .002d),
            "Inferred Y intervals must agree in painted coordinates: " + string.Join(",", actual));
    }

    public static void PptxScatterAutomaticYRetainsSceneFallbackAndCancellationOwnership()
    {
        RangeCase sample = AutomaticYCases().Single(c => c.Id == "range-scatter-103-auto");
        XElement mismatch = XElement.Parse(Axis("2", "l", "1", "400", "500", "100", "0000FF"));
        TestAssert.Equal(new PptxRenderer.ChartValueExtents(-17d, 103d), ResolveRange(sample, true, mismatch).Extents);
        foreach (RangeCase rejected in new[] { sample with { MajorUnit = double.NaN }, sample with { MajorUnit = -1d },
            sample with { MajorUnit = double.PositiveInfinity }, sample with { Minimum = double.NaN } })
            CheckFallback(rejected);
        CheckFallback(sample, logarithmic: true);
        CheckFallback(sample, bubble: true);

        XElement missingScaling = XDocument.Parse(RangeChart(sample)).Descendants(CurveC + "valAx").Last();
        missingScaling.Element(CurveC + "scaling")!.Remove();
        var legacy = (PptxRenderer.ChartValueExtents)CurveInvoke("ReadSceneOrXmlBubbleChartValueAxisExtents", null, missingScaling,
            new PptxRenderer.ChartValueExtents(1d, 79d), 5d, true);
        TestAssert.Equal(legacy, ResolveRange(sample, false, missingScaling).Extents);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => ResolveRange(sample, true, cancellationToken: cancellation.Token));

        static void CheckFallback(RangeCase source, bool logarithmic = false, bool bubble = false)
        {
            string xml = RangeChart(source, logarithmic);
            var scene = PptxTests.BuildSingleChartScene(xml)!;
            var axis = scene.Axes.Single(a => a.PositionKind == PptxSceneChartAxisPosition.Left);
            XElement element = XDocument.Parse(xml).Descendants(CurveC + "valAx").Last();
            var expected = (PptxRenderer.ChartValueExtents)CurveInvoke("ReadSceneOrXmlBubbleChartValueAxisExtents", axis, element,
                new PptxRenderer.ChartValueExtents(1d, 79d), 5d, true);
            TestAssert.Equal(expected, ResolveRange(source, true, logarithmic: logarithmic, bubble: bubble).Extents);
        }
    }

    private static RangeCase[] AutomaticYCases()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", "scatter-automatic-y-office.json"));
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        return JsonSerializer.Deserialize<RangeCase[]>(document.RootElement.GetProperty("Cases").GetRawText())!;
    }
}
