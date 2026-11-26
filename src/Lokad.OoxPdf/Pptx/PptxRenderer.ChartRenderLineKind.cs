using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static async ValueTask<bool> TryRenderLineChartKindAsync(
        PdfGraphicsBuilder graphics,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        ChartColorPalette chartPalette,
        ShapeBounds bounds,
        PptxSceneChart sceneChart,
        ChartWorkbookData? workbook,
        List<PdfFontResource> fonts,
        PresentationFontResolver fontResolver,
        bool plotVisibleOnly,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        PptxSceneChartPlot? linePlot = ReadFirstSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Line);
        if (linePlot is not null)
        {
            IReadOnlyList<ChartIndexedNumberVector> lineSeriesVectors = ReadSharedChartSeriesVectors(linePlot, workbook, plotVisibleOnly);
            if (CountRenderableSeries(lineSeriesVectors) != 0)
            {
                int? lineStyleId = sceneChart.SeriesStyleId;
                double? themeLineWidth = ReadChartLineThemeWidth(theme);
                ChartLinePlotOptions lineOptions = ReadSceneChartLineOptions(sceneChart, linePlot, PptxSceneChartGrouping.Standard, defaultSmooth: themeLineWidth is not null);
                IReadOnlyList<PptxSceneChartSeries> sourceSeries = linePlot.Series;
                bool varyPointColors = sceneChart.Plots.Count == 1 && lineSeriesVectors.Count == 1 && lineOptions.VaryColors;
                IReadOnlyList<ChartLineAppearance> seriesAppearance = ReadLineSeriesAppearance(sourceSeries, lineOptions,
                    theme, colorMap, chartPalette, lineStyleId, themeLineWidth,
                    varyPointColors, context.CancellationToken);
                Func<int, int, ChartLineAppearance>? pointAppearance = ReadLinePointAppearance(
                    theme, colorMap, chartPalette, lineSeriesVectors, varyPointColors, sourceSeries, seriesAppearance,
                    lineStyleId, themeLineWidth, context.CancellationToken);
                ChartAxesStyle axesStyle = ReadSceneChartAxesStyle(sceneChart, linePlot);
                ChartShapeStyle plotAreaStyle = ToChartShapeStyle(sceneChart.PlotAreaStyle);
                ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, linePlot).FirstOrDefault();
                bool categoryPointsOnTickMarks = ResolveSceneCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis);
                ChartValueExtents dataExtents = workbook is null
                    ? GetLineChartValueExtents(lineSeriesVectors, lineOptions.Stacked, lineOptions.PercentStacked)
                    : GetSharedLineChartValueExtents(linePlot, lineOptions.Stacked, lineOptions.PercentStacked, workbook, plotVisibleOnly);
                ChartValueExtents valueExtents = ReadCartesianChartValueAxisExtents(sceneChart, valueAxis.SceneAxis, dataExtents,
                    lineOptions.PercentStacked, useNearMaximumHeadroom: !lineOptions.PercentStacked,
                    nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneChartValueAxisRenderOptions(valueAxis.SceneAxis, valueExtents, lineOptions.PercentStacked) with
                { TickPolicy = ReadCartesianChartAxisTickPolicy(sceneChart, valueAxis.SceneAxis, valueExtents, lineOptions.PercentStacked) };
                ChartLayout chartLayout = (await GetCartesianChartLayoutAsync(document, theme, bounds, sceneChart, workbook, plotVisibleOnly, fontResolver, resolvedValueExtents: valueExtents, context: context).ConfigureAwait(false));
                chartLayout = await PrepareDefaultAxisTitleBandsAsync(chartLayout, theme, sceneChart, fontResolver).ConfigureAwait(false);
                bool pointLegend = sceneChart.Plots.Count == 1 && sourceSeries.Count == 1 && pointAppearance is not null;
                IReadOnlyList<ChartLegendEntry> legendEntries = pointLegend && pointAppearance is not null
                    ? BuildLinePointLegendEntries(linePlot, workbook, plotVisibleOnly, lineSeriesVectors[0].DensePointCount(), lineStyleId, pointAppearance)
                    : BuildLineSeriesLegendEntries(linePlot, workbook, seriesAppearance, sourceSeries, themeLineWidth, lineOptions.Stacked, lineStyleId);
                ChartTextStyle legendTextStyle = ReadSceneChartLegendTextStyle(theme, sceneChart);
                ChartLegendBox? preparedLegendBox;
                (chartLayout, preparedLegendBox) = await PrepareCartesianHorizontalLegendAsync(chartLayout,
                    legendEntries, legendTextStyle, fontResolver, context).ConfigureAwait(false);
                if (chartLayout.CategoryAxisPlacement is { } categoryPlacement)
                {
                    valueAxisOptions = valueAxisOptions with { CrossingValue = categoryPlacement.CrossingValue };
                }
                ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, linePlot);
                bool categoryAxisReversed = ReadSceneValueAxisReversed(categoryAxis.SceneAxis);
                var categoryLabels = ReadSharedCategoryLabels(linePlot, workbook, plotVisibleOnly);
                chartLayout = await PrepareCartesianCategoryFramesAsync(chartLayout, theme, sceneChart,
                    valueAxis, valueExtents, valueAxisOptions, GetValueAxisAutoTickTargetCount(false, true, false), axesStyle.ValueAxisRightSide,
                    categoryAxis, categoryLabels, chartLayout.CategoryAxisPlacement?.LabelPlotRatio ?? 0d,
                    chartLayout.CategoryAxisPlacement?.LabelsHighSide ?? false, fontResolver, context, preparedLegendBox,
                    lineOptions.PercentStacked ? "0%" : null).ConfigureAwait(false);
                var admitted = await AdmitCartesianRightLegendAsync(chartLayout, sceneChart, valueAxis,
                    categoryLabels, categoryPointsOnTickMarks, categoryAxisReversed, legendEntries, legendTextStyle,
                    fontResolver, ChartLegendPlacement.Default).ConfigureAwait(false);
                chartLayout = admitted.Layout;
                preparedLegendBox ??= admitted.Box;
                RenderChartAreaStyle(graphics, document, bounds, sceneChart);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                RenderLineChart(graphics, chartLayout.PlotAreaBox, plotBox, lineSeriesVectors, lineOptions, seriesAppearance, valueAxisOptions, axesStyle, plotAreaStyle, valueExtents, categoryPointsOnTickMarks,
                    ReadSceneChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis").FontSize, lineStyleId, categoryTicksTopSide: chartLayout.CategoryAxisPlacement?.AxisHighSide ?? false, pointAppearance: pointAppearance, categoryAxisReversed: categoryAxisReversed, varyPointColors: varyPointColors, context: context);
                if (axesStyle.CategoryAxisVisible && IsSceneChartAxisLabelVisible(categoryAxis.SceneAxis))
                {
                    ChartAxisPlacement? labelPlacement = chartLayout.CategoryAxisPlacement;
                    double? labelAxisY = labelPlacement is { } placement
                        ? plotBox.Y + plotBox.Height * placement.LabelPlotRatio
                        : null;
                    await RenderChartCategoryLabelsAsync(theme, graphics, chartLayout.Frame, plotBox, sceneChart, categoryAxis.SceneAxis, categoryLabels, horizontalBars: false, axisCoordinate: labelAxisY, categoryLabelsOnTickMarks: categoryPointsOnTickMarks, labelsHighSide: labelPlacement?.LabelsHighSide ?? false, fontResolver: fontResolver, chartFonts: fonts, categoryLabelOffsetAdjustment: chartLayout.CategoryLabelOffsetAdjustment, prepared: chartLayout.CategoryLabels, categoryAxisReversed: categoryAxisReversed).ConfigureAwait(false);
                }

                if (axesStyle.ValueAxisVisible && IsSceneChartAxisLabelVisible(valueAxis.SceneAxis))
                {
                    await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, sceneChart, valueAxis.SceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars: false, rightSide: axesStyle.ValueAxisRightSide, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: lineOptions.PercentStacked ? "0%" : null, fontResolver: fontResolver, chartFonts: fonts, tickPolicy: valueAxisOptions.TickPolicy, context: context).ConfigureAwait(false);
                }
                await RenderSecondaryChartValueAxisLabelsAsync(theme, graphics, plotBox, sceneChart, dataExtents,
                    fontResolver, chartFonts: fonts, primarySceneAxis: valueAxis.SceneAxis, context: context).ConfigureAwait(false);
                DrawVerticalChartValueAxisMajorTicks(graphics, theme, sceneChart, valueAxis.SceneAxis,
                    plotBox, valueExtents, valueAxisOptions, axesStyle, context);
                await RenderDefaultChartAxisTitlesAsync(theme, graphics, chartLayout, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                await RenderChartLegendAsync(graphics, chartLayout.Frame, plotBox, legendEntries, chartLayout.Legend,
                    legendTextStyle, fontResolver, ChartLegendPlacement.Default, chartFonts: fonts,
                    preparedBox: preparedLegendBox).ConfigureAwait(false);
                await RenderCategoryPointDataLabelsAsync(
                    theme,
                    colorMap,
                    chartPalette,
                    graphics,
                    chartLayout.Frame,
                    plotBox,
                    lineSeriesVectors,
                    valueExtents,
                    valueAxisOptions.Reversed,
                    categoryPointsOnTickMarks,
                    [],
                    [],
                    ReadSceneDataLabelOptions(sceneChart, linePlot),
                    ReadSceneSeriesDataLabelOptions(sceneChart, linePlot),
                    ReadSceneCategoryLabelVector(linePlot, workbook, plotVisibleOnly),
                    ReadSharedChartSeriesNames(linePlot, workbook), fontResolver, fonts, context, sceneChart, linkAnnotations, reportedHyperlinkIds, pointAppearance: pointLegend ? pointAppearance : null, seriesAppearance: seriesAppearance, categoryAxisReversed: categoryAxisReversed, chartRootFontSize: sceneChart.TextStyle.FontSize).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }
}
