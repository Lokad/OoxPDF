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
    private static async ValueTask<bool> TryRenderLineChartKindAsync(
        PdfGraphicsBuilder graphics,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        ChartColorPalette chartPalette,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart sceneChart,
        ChartWorkbookData? workbook,
        List<PdfFontResource> fonts,
        PresentationFontResolver fontResolver,
        bool plotVisibleOnly,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        PptxSceneChartPlot? linePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Line, 0);
        if (linePlot is not null)
        {
            XElement lineChart = linePlot.Source;
            IReadOnlyList<ChartIndexedNumberVector> lineSeriesVectors = ReadSharedChartSeriesVectors(linePlot, workbook, plotVisibleOnly);
            if (CountRenderableSeries(lineSeriesVectors) != 0)
            {
                double? themeLineWidth = ReadChartLineThemeWidth(theme);
                ChartLinePlotOptions lineOptions = ReadSceneChartLineOptions(sceneChart, linePlot, PptxSceneChartGrouping.Standard, defaultSmooth: themeLineWidth is not null);
                IReadOnlyList<PptxSceneChartSeries> sourceSeries = linePlot.Series;
                bool varyPointColors = sceneChart.Plots.Count == 1 && lineSeriesVectors.Count == 1 && lineOptions.VaryColors.Value;
                IReadOnlyList<ChartLineAppearance> seriesAppearance = ReadLineSeriesAppearance(sourceSeries, lineOptions,
                    theme, colorMap, chartPalette, ReadChartStyleId(chartXml), themeLineWidth,
                    varyPointColors, context.CancellationToken);
                Func<int, int, ChartLineAppearance>? pointAppearance = ReadLinePointAppearance(
                    theme, colorMap, chartPalette, lineSeriesVectors, varyPointColors, sourceSeries, seriesAppearance,
                    ReadChartStyleId(chartXml), themeLineWidth, context.CancellationToken);
                ChartAxesStyle axesStyle = ReadSceneChartAxesStyle(sceneChart, linePlot);
                ChartShapeStyle plotAreaStyle = ToChartShapeStyle(sceneChart.PlotAreaStyle);
                ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, linePlot).FirstOrDefault();
                XElement? valueAxisForScale = valueAxis.XmlAxis;
                bool categoryPointsOnTickMarks = ResolveSceneCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis);
                ChartValueExtents valueExtents = ReadCartesianChartValueAxisExtents(sceneChart, valueAxis.SceneAxis, GetSharedLineChartValueExtents(linePlot, lineOptions.Stacked, lineOptions.PercentStacked, workbook, plotVisibleOnly), lineOptions.PercentStacked, useNearMaximumHeadroom: !lineOptions.PercentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneChartValueAxisRenderOptions(valueAxis.SceneAxis, valueExtents, lineOptions.PercentStacked) with
                { TickPolicy = ReadCartesianChartAxisTickPolicy(sceneChart, valueAxis.SceneAxis, valueExtents, lineOptions.PercentStacked) };
                ChartLayout chartLayout = (await GetCartesianChartLayoutAsync(document, theme, bounds, chartXml, sceneChart, colorMap, workbook, plotVisibleOnly, fontResolver, resolvedValueExtents: valueExtents, context: context).ConfigureAwait(false));
                bool pointLegend = sceneChart.Plots.Count == 1 && sourceSeries.Count == 1 && pointAppearance is not null;
                IReadOnlyList<ChartLegendEntry> legendEntries = pointLegend && pointAppearance is not null
                    ? BuildLinePointLegendEntries(linePlot, lineChart, workbook, plotVisibleOnly, lineSeriesVectors[0].DensePointCount(), ReadChartStyleId(chartXml), pointAppearance)
                    : BuildLineSeriesLegendEntries(linePlot, lineChart, workbook, seriesAppearance, sourceSeries, themeLineWidth, lineOptions.Stacked, ReadChartStyleId(chartXml));
                ChartTextStyle legendTextStyle = ReadSceneChartLegendTextStyle(theme, sceneChart);
                ChartLegendBox? preparedLegendBox = null;
                if (chartLayout.Legend.PositionKind == PptxSceneChartLegendPosition.Bottom)
                {
                    bool automatic = !chartLayout.ManualPlotLayoutApplied && !chartLayout.Legend.Layout.HasGeometryValues;
                    var prepared = await LayoutChartLegendRowsAsync(chartLayout, legendEntries, legendTextStyle,
                        fontResolver, context, chartLayout.PlotBox.Width, measuredBand: automatic).ConfigureAwait(false);
                    chartLayout = prepared.Layout;
                    preparedLegendBox = prepared.Box;
                }
                if (chartLayout.CategoryAxisPlacement is { } categoryPlacement)
                {
                    valueAxisOptions = valueAxisOptions with { CrossingValue = categoryPlacement.CrossingValue };
                }
                ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, linePlot);
                bool categoryAxisReversed = ReadSceneValueAxisReversed(categoryAxis.SceneAxis);
                var categoryLabels = ReadSharedCategoryLabels(linePlot, workbook, plotVisibleOnly);
                chartLayout = await PrepareCartesianCategoryFramesAsync(chartLayout, theme, sceneChart, chartXml,
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
                    ReadSceneChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis").FontSize, ReadChartStyleId(chartXml), categoryTicksTopSide: chartLayout.CategoryAxisPlacement?.AxisHighSide ?? false, pointAppearance: pointAppearance, categoryAxisReversed: categoryAxisReversed, varyPointColors: varyPointColors, context: context);
                if (axesStyle.CategoryAxisVisible && IsSceneChartAxisLabelVisible(categoryAxis.SceneAxis))
                {
                    ChartAxisPlacement? labelPlacement = chartLayout.CategoryAxisPlacement;
                    double? labelAxisY = labelPlacement is { } placement
                        ? plotBox.Y + plotBox.Height * placement.LabelPlotRatio
                        : null;
                    await RenderChartCategoryLabelsAsync(theme, graphics, chartLayout.Frame, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, categoryLabels, horizontalBars: false, axisCoordinate: labelAxisY, categoryLabelsOnTickMarks: categoryPointsOnTickMarks, labelsHighSide: labelPlacement?.LabelsHighSide ?? false, fontResolver: fontResolver, chartFonts: fonts, categoryLabelOffsetAdjustment: chartLayout.CategoryLabelOffsetAdjustment, prepared: chartLayout.CategoryLabels, categoryAxisReversed: categoryAxisReversed).ConfigureAwait(false);
                }

                if (axesStyle.ValueAxisVisible && IsSceneChartAxisLabelVisible(valueAxis.SceneAxis))
                {
                    await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars: false, rightSide: axesStyle.ValueAxisRightSide, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: lineOptions.PercentStacked ? "0%" : null, fontResolver: fontResolver, chartFonts: fonts, tickPolicy: valueAxisOptions.TickPolicy, context: context).ConfigureAwait(false);
                }
                await RenderSecondaryChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, GetSharedLineChartValueExtents(linePlot, lineOptions.Stacked, lineOptions.PercentStacked, workbook, plotVisibleOnly), fontResolver, chartFonts: fonts, primarySceneAxis: valueAxis.SceneAxis, primaryXmlAxis: valueAxis.XmlAxis, context: context).ConfigureAwait(false);
                // Major value tick marks (Office honors out/in/cross at value gridlines).
                if (axesStyle.ValueAxisVisible)
                {
                    PptxSceneChartAxisTickMark lineValueTickMark = ReadSceneChartAxisMajorTickMark(valueAxis.SceneAxis);
                    if (lineValueTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke lineValueTickStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
                        if (lineValueTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle lineValueTickStyle = ReadSceneChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                            IReadOnlyList<double> lineTickValues = GetChartAxisTickValuesCore(valueExtents, valueAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: false, valueAxisLabelsVisible: IsSceneChartAxisLabelVisible(valueAxis.SceneAxis), manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied), valueAxisOptions.TickPolicy, context: context);
                            double[] lineTickEdges = new double[lineTickValues.Count];
                            for (int lineTickIndex = 0; lineTickIndex < lineTickValues.Count; lineTickIndex++)
                            {
                                lineTickEdges[lineTickIndex] = ChartValueToPlotCoordinate(valueExtents, lineTickValues[lineTickIndex], plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                            }

                            bool lineValueLabelsRightSide = ResolveSceneValueAxisLabelsRightSide(sceneChart, valueAxis.SceneAxis, axesStyle.ValueAxisRightSide);
                            double lineValueAxisX = axesStyle.ValueAxisRightSide ? plotBox.X + plotBox.Width : plotBox.X;
                            SetChartStroke(graphics, lineValueTickStroke);
                            StrokeMajorTickSegments(graphics, lineTickEdges, lineValueAxisX, verticalSegments: false, outwardIsLowSide: !lineValueLabelsRightSide, lineValueTickMark, lineValueTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                        }
                    }
                }
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
                    ReadSharedChartSeriesNames(linePlot, workbook), fontResolver, fonts, context, sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds, pointAppearance: pointLegend ? pointAppearance : null, seriesAppearance: seriesAppearance, categoryAxisReversed: categoryAxisReversed, chartRootFontSize: sceneChart.TextStyle.FontSize).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }
}
