using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private readonly record struct ChartBarPlotOptions(
        PptxSceneChartGrouping Grouping,
        PptxSceneChartBarDirection BarDirection,
        ChartBooleanOption VaryColors,
        double GapWidth,
        double Overlap,
        IReadOnlyList<PptxSceneChartNegativeInversion> NegativeInversions)
    {
        public bool CategoryAxisReversed { get; init; }
        public bool CategoryPointsOnTickMarks { get; init; }

        public double CategoryPitch(double length, int count) =>
            length / Math.Max(1, count - (CategoryPointsOnTickMarks ? 1 : 0));

        public bool InvertsNegativePoint(int seriesIndex, int pointIndex) =>
            NegativeInversions is null || seriesIndex >= NegativeInversions.Count ||
            NegativeInversions[seriesIndex].ForPoint(pointIndex);
    }

    // Parsed plot scalars share their source defaults and normalization here.
    private static ChartBarPlotOptions ReadSceneChartBarOptions(PptxSceneChartPlot plot, PptxSceneChartGrouping defaultGrouping) => new(
        ReadSceneChartGrouping(plot, defaultGrouping),
        ReadSceneChartBarDirection(plot),
        ReadSceneChartVaryColors(plot),
        NormalizeGapWidth(plot.GapWidth),
        NormalizeOverlap(plot.Overlap),
        plot.Series.Select(series => series.NegativeInversion).ToArray());

    private static ChartBooleanOption ReadSceneChartVaryColors(PptxSceneChartPlot plot) =>
        new(plot.VaryColors ?? true, plot.VaryColorsValue, plot.VaryColors is not null);

    private readonly record struct ChartLinePlotOptions(
        PptxSceneChartGrouping Grouping,
        bool Stacked,
        bool PercentStacked,
        ChartBooleanOption VaryColors,
        IReadOnlyList<ChartBooleanOption> SmoothSeries,
        PptxSceneChartDisplayBlanksAs DisplayBlanksAs);

    private static ChartLinePlotOptions ReadSceneChartLineOptions(PptxSceneChart sceneChart, PptxSceneChartPlot plot, PptxSceneChartGrouping defaultGrouping, bool defaultSmooth = false)
    {
        PptxSceneChartGrouping grouping = ReadSceneChartGrouping(plot, defaultGrouping);
        return new ChartLinePlotOptions(grouping, IsStackedChartGrouping(grouping), IsPercentStackedChartGrouping(grouping),
            ReadSceneChartVaryColors(plot), ReadSceneSmoothSeries(plot, applyPlotFallback: true, defaultValue: defaultSmooth),
            ResolveChartDisplayBlanksAs(sceneChart));
    }

    private readonly record struct ChartAreaPlotOptions(
        PptxSceneChartGrouping Grouping,
        bool Stacked,
        bool PercentStacked,
        PptxSceneChartDisplayBlanksAs DisplayBlanksAs);

    private static ChartAreaPlotOptions ReadSceneChartAreaOptions(PptxSceneChart sceneChart, PptxSceneChartPlot plot, PptxSceneChartGrouping defaultGrouping)
    {
        PptxSceneChartGrouping grouping = ReadSceneChartGrouping(plot, defaultGrouping);
        return new ChartAreaPlotOptions(grouping, IsStackedChartGrouping(grouping), IsPercentStackedChartGrouping(grouping),
            ResolveChartDisplayBlanksAs(sceneChart));
    }

    private readonly record struct ChartScatterPlotOptions(
        PptxSceneChartScatterStyle ScatterStyle,
        bool ConnectLines,
        IReadOnlyList<ChartBooleanOption> SmoothSeries);

    private static ChartScatterPlotOptions ReadSceneChartScatterOptions(PptxSceneChartPlot plot)
    {
        // Omitted/marker-only styles connect; explicit series noFill suppresses
        // the path. Unknown/rejected styles retain their existing fallback.
        PptxSceneChartScatterStyle scatterStyle = string.IsNullOrEmpty(plot.ScatterStyle)
            ? PptxSceneChartScatterStyle.LineMarker : plot.ScatterStyleKind;
        return new ChartScatterPlotOptions(scatterStyle,
            scatterStyle is not (PptxSceneChartScatterStyle.None or PptxSceneChartScatterStyle.Unknown),
            ReadSceneSmoothSeries(plot, applyPlotFallback: false, defaultValue: true));
    }

    private static double NormalizeGapWidth(double? rawGapWidth)
    {
        return rawGapWidth is { } value
            ? Math.Clamp(value, 0d, 500d)
            : 150d;
    }

    private static double NormalizeOverlap(double? rawOverlap)
    {
        return rawOverlap is { } value
            ? Math.Clamp(value, -100d, 100d)
            : 0d;
    }

    private static double NormalizeAngleDegrees(double angle)
    {
        double normalized = angle % 360d;
        return normalized < 0d ? normalized + 360d : normalized;
    }

    private readonly record struct ChartPolarPointOptions(
        IReadOnlyDictionary<int, ChartSeriesFill> PointFills,
        IReadOnlyDictionary<int, ChartSeriesStroke> PointStrokes,
        IReadOnlyDictionary<int, double> PointExplosions,
        double FirstSliceAngle);

    private static IReadOnlyList<PptxSceneChartSeries> ReadXmlChartSeries(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        PptxSceneChartPlotKind plotKind = chartElement.Name == ChartNamespace + "ser"
            ? PptxSceneChartPlotKind.Unknown
            : PptxSceneBuilder.ParseChartPlotKind(chartElement.Name.LocalName);
        bool markersEnabled = chartElement.Name != ChartNamespace + "ser" &&
            PptxSceneBuilder.IsOoxmlBooleanElementEnabled(chartElement.Element(ChartNamespace + "marker"));
        XElement plotElement = chartElement.Name == ChartNamespace + "ser"
            ? new XElement(ChartNamespace + "unknownChart", chartElement)
            : chartElement;
        return PptxSceneBuilder.ReadChartSeries(plotElement, theme, colorMap, plotKind, markersEnabled);
    }

    private static ChartPolarPointOptions ReadSceneChartPolarPointOptions(PptxSceneChartPlot plot, ChartWorkbookData? workbook)
    {
        var first = plot.Series.FirstOrDefault();
        return new ChartPolarPointOptions(
            first is null ? new Dictionary<int, ChartSeriesFill>() : ReadSceneChartPointFills(first),
            first is null ? new Dictionary<int, ChartSeriesStroke>() : ReadSceneChartPointStrokes(first),
            ReadSceneOrXmlChartPointExplosions(plot, plot.Source, workbook),
            plot.FirstSliceAngle is { } angle ? NormalizeAngleDegrees(angle) : 0d);
    }

    private readonly record struct ChartDoughnutPlotOptions(
        ChartPolarPointOptions PolarPoints,
        double HoleSize);

    private static ChartDoughnutPlotOptions ReadSceneChartDoughnutOptions(PptxSceneChartPlot plot, ChartWorkbookData? workbook) => new(
        ReadSceneChartPolarPointOptions(plot, workbook),
        plot.HoleSize is { } size
            ? Math.Clamp(size / 100d, PptxChartMetricRules.DoughnutHoleMinimumRatio, PptxChartMetricRules.DoughnutHoleMaximumRatio)
            : PptxChartMetricRules.DoughnutHoleFallbackRatio);

}
