using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool CanMeasureAutomaticColumnScale(PptxSceneChart? chart, PptxSceneChartAxis? axis) =>
        chart is { Plots.Count: 1 } && chart.Plots[0] is
            { PlotKind: PptxSceneChartPlotKind.Bar, BarDirectionKind: PptxSceneChartBarDirection.Column,
              GroupingKind: PptxSceneChartGrouping.Clustered } &&
        chart.Axes.Count(value => value.AxisKind == PptxSceneChartAxisKind.Value) == 1 &&
        axis is { OrientationKind: not PptxSceneChartAxisOrientation.Unknown } &&
        ReadChartAxisDeclaration(axis) is { Minimum: null, Maximum: null, MajorUnit: null } &&
        // An inner data rectangle follows automatic scale allocation. An
        // authored outer frame owns that allocation and keeps its fallback.
        (!chart.PlotAreaLayout.HasGeometryValues ||
            chart.PlotAreaLayout.LayoutTargetKind == PptxSceneChartManualLayoutTarget.Inner) &&
        !chart.Title.Layout.HasGeometryValues &&
        !chart.Legend.Layout.HasGeometryValues;

    private static async ValueTask<(ChartValueExtents Extents, ChartAxisUnits Units)?> MeasureAutomaticColumnScaleAsync(
        PptxTheme theme, PptxSceneChart chart, XDocument xml, ChartAxisSource axis, ChartAxisSource category,
        ChartValueExtents data, ChartLayout layout, double? legendBand, PresentationFontResolver resolver)
    {
        if (!CanMeasureAutomaticColumnScale(chart, axis.SceneAxis) || legendBand is null ||
            category.SceneAxis?.CrossesAt is not null || data.Min < 0d && data.Max > 0d ||
            !double.IsFinite(data.Min) || !double.IsFinite(data.Max) ||
            Math.Max(Math.Abs(data.Min), Math.Abs(data.Max)) <= PptxChartMetricRules.AxisValueEpsilon)
            return null;
        var plot = chart.Plots[0];
        var header = await TryMeasureCartesianHeaderAsync(theme, chart, layout.Legend, plot, plot.Source,
            xml, resolver, automaticScaleAllocation: true).ConfigureAwait(false);
        if (header is null) return null;
        // The automatic outer allocation precedes caption fitting and authored
        // placement. Reuse the admitted legend band; do not pack its rows again.
        double titleBand = header.Value.Reserve - header.Value.NumericClearance;
        double available = layout.Frame.Height - titleBand - legendBand.Value -
            2d * PptxChartMetricRules.HorizontalBarPlotTopReserve;
        ChartTextStyle style = ReadSceneChartTextStyle(theme, chart, axis.SceneAxis,
            PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
        style = style with { FontSize = Math.Round(style.FontSize * 20d, MidpointRounding.AwayFromZero) / 20d };
        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        if (await measure.TryMeasureWindowsLineHeightAsync(style).ConfigureAwait(false) is not { } line ||
            !(line > 0d) || !double.IsFinite(available) || available <= 0d)
            return null;
        double target = Math.Clamp((available - 2d * CartesianAxisEndClearance(line)) / line - 1d,
            1d, PptxChartMetricRules.AxisNiceTickTargetCount);
        double range = Math.Max(Math.Abs(data.Min), Math.Abs(data.Max));
        double unit = ChooseChartAxisMajorUnit(range, target);
        ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(axis.SceneAxis, axis.XmlAxis, data,
            useNearMaximumHeadroom: true, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio,
            boundsTickTargetCount: target);
        return (extents, new ChartAxisUnits(unit, axis.SceneAxis?.MinorUnit));
    }
}
