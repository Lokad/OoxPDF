using System.Globalization;
using System.Reflection;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxColumnIntervalTests
{
    public static void PptxHorizontalIntervalsShareOfficeSourceMinimums()
    {
        // Pinned independent Office controls: authored origin, reversal, auto
        // unit and a maximum on the interval lattice. Unaligned maxima vanish.
        (int Minimum, int? Major, int Maximum, string Expected)[] cases =
        [
            (7, 25, 103, "7,32,57,82"), (-7, 20, 103, "-7,13,33,53,73,93"),
            (7, 25, 107, "7,32,57,82,107"), (7, null, 103, "7,17,27,37,47,57,67,77,87,97")
        ];
        foreach (var test in cases)
        foreach (bool reversed in new[] { false, true })
        {
            string actual = string.Join(",", Labels(Render(test.Minimum, test.Major, test.Maximum, reversed: reversed, horizontal: true)));
            TestAssert.Equal(test.Expected, actual);
        }
    }

    public static void PptxHorizontalIntervalsSurvivePlacementAndStacking()
    {
        foreach (string grouping in new[] { "clustered", "stacked" })
        foreach (bool manual in new[] { false, true })
            TestAssert.Equal("7,32,57,82", string.Join(",", Labels(Render(7, 25, grouping: grouping, manual: manual, horizontal: true))));
        // Fresh Office upper-only controls anchor whole intervals at103 and
        // include zero in bar data extents: the old zero-aligned fallback was
        // an implementation assertion, not a matching native range.
        TestAssert.Equal("-22,3,28,53,78,103", string.Join(",", Labels(Render(null, 25, horizontal: true))));
    }

    public static void PptxHorizontalTicksAndBothGridKindsUseTheLabelIntervals()
    {
        foreach (bool reversed in new[] { false, true })
        {
            var lines = Lines(Render(7, 25, reversed: reversed, grids: true, horizontal: true));
            double[] axis = lines.Single(p => p.Color == new RgbColor(255, 0, 0) &&
                Math.Abs(p.V[1] - p.V[3]) < .001d && Math.Abs(p.V[0] - p.V[2]) > 100d).V;
            double lower = Math.Min(axis[0], axis[2]), span = Math.Abs(axis[2] - axis[0]);
            double Value(double x) => 7d + (reversed ? 1d - (x - lower) / span : (x - lower) / span) * 96d;
            double[] Levels(RgbColor color, bool ticks) => lines.Where(p => p.Color == color && Math.Abs(p.V[0] - p.V[2]) < .001d &&
                (ticks ? Math.Abs(p.V[1] - p.V[3]) is > 0d and < 20d : Math.Abs(p.V[1] - p.V[3]) > 100d))
                .Select(p => Value(p.V[0])).Order().ToArray();
            void Check(double[] actual, double[] expected, string description) => TestAssert.True(actual.Length == expected.Length &&
                actual.Zip(expected).All(p => Math.Abs(p.First - p.Second) < .002d),
                description + ": " + string.Join(",", actual.Select(v => v.ToString("F3", CultureInfo.InvariantCulture))));
            Check(Levels(new RgbColor(255, 0, 0), true), [7d, 32d, 57d, 82d], "Marks share the source origin");
            Check(Levels(new RgbColor(0, 0, 255), false), [32d, 57d, 82d], "Major grids exclude the crossing");
            Check(Levels(new RgbColor(255, 0, 255), false), Enumerable.Range(1, 19).Where(i => i % 5 != 0).Select(i => 7d + 5d * i).ToArray(),
                "Minor grids subdivide source intervals and exclude major levels");
        }
    }

    public static void PptxAuthoredIntervalOriginKeepsRejectedAndMissingScaleFallbacks()
    {
        const BindingFlags flags = BindingFlags.NonPublic | BindingFlags.Static;
        Type extentsType = typeof(PptxRenderer).GetNestedType("ChartValueExtents", BindingFlags.NonPublic)!;
        object extents = Activator.CreateInstance(extentsType, [7d, 103d])!;
        MethodInfo read = typeof(PptxRenderer).GetMethod("ReadCartesianChartAxisTickPolicy", flags)!;
        string Policy(PptxSceneChart? scene, PptxSceneChartAxis? axis) => read.Invoke(null, [scene, axis, extents, false])!.ToString()!;
        foreach (string scaling in new[] { "", "<c:scaling/>", "<c:scaling><c:min val=\"bad\"/></c:scaling>",
            "<c:scaling><c:min val=\"Infinity\"/></c:scaling>", "<c:scaling><c:min val=\"8\"/></c:scaling>",
            "<c:scaling><c:min val=\"7\"/><c:logBase val=\"10\"/></c:scaling>" })
        {
            var scene = PptxTests.BuildSingleChartScene($"<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea><c:barChart><c:barDir val=\"bar\"/></c:barChart><c:valAx><c:axId val=\"2\"/>{scaling}</c:valAx></c:plotArea></c:chart></c:chartSpace>")!;
            TestAssert.Equal("ZeroAlignedWithEndpoints", Policy(scene, scene.Axes.Single()));
        }
        TestAssert.Equal("ZeroAlignedWithEndpoints", Policy(null, null));
    }
}
