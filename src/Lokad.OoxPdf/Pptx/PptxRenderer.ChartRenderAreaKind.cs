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
    private static async ValueTask<bool> TryRenderAreaChartKindAsync(
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
        PptxSceneChartPlot? areaPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Area, 0);
        if (areaPlot is not null)
        {
            XElement areaChart = areaPlot.Source;
            IReadOnlyList<ChartIndexedNumberVector> areaSeriesVectors = ReadSharedChartSeriesVectors(areaPlot, workbook, plotVisibleOnly);
            if (CountRenderableSeries(areaSeriesVectors) != 0)
            {
                ChartAreaPlotOptions areaOptions = ReadSceneChartAreaOptions(sceneChart, areaPlot, PptxSceneChartGrouping.Standard);
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneSeriesFills(areaPlot);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneSeriesStrokes(areaPlot, ChartSeriesInheritedStrokeWidth);
                ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, areaPlot).FirstOrDefault();
                ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, areaPlot);
                bool categoryAxisReversed = ReadSceneOrXmlValueAxisReversed(categoryAxis.SceneAxis, categoryAxis.XmlAxis);
                bool categoryPointsOnTickMarks = ResolveSceneCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis);
                ChartValueExtents valueExtents = ReadCartesianChartValueAxisExtents(sceneChart, valueAxis.SceneAxis, valueAxis.XmlAxis, GetAreaChartValueExtents(areaSeriesVectors, areaOptions.Stacked, areaOptions.PercentStacked), areaOptions.PercentStacked, useNearMaximumHeadroom: areaOptions.Stacked && !areaOptions.PercentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AreaChartStackedAxisNearMaximumHeadroomRatio);
                ChartLayout chartLayout = await GetCartesianChartLayoutAsync(document, theme, bounds, chartXml, sceneChart, colorMap, workbook, plotVisibleOnly, fontResolver,
                    resolvedValueExtents: valueAxis.SceneAxis is { HasLogarithmicScale: true } ? null : valueExtents, context: context).ConfigureAwait(false);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(valueAxis.SceneAxis, valueAxis.XmlAxis, theme, valueExtents, areaOptions.PercentStacked) with
                {
                    TickPolicy = ReadCartesianChartAxisTickPolicy(sceneChart, valueAxis.SceneAxis, valueExtents, areaOptions.PercentStacked),
                    // The category axis owns its numeric crossing independently
                    // of whether the surrounding frame can be measured.
                    CrossingValue = ReadSceneOrXmlValueAxisCrossingValue(categoryAxis.SceneAxis, categoryAxis.XmlAxis, valueExtents)
                };
                ChartAxesStyle axesStyle = ReadSceneChartAxesStyle(sceneChart, areaPlot);
                ChartShapeStyle plotAreaStyle = ToChartShapeStyle(sceneChart.PlotAreaStyle);
                IReadOnlyList<ChartLegendEntry> legendEntries = BuildFillLegendEntries(theme, colorMap, chartPalette, areaPlot, areaChart, seriesFills, seriesStrokes, paletteOffset: 0, workbook: workbook, reverseOrder: IsStackedChartGrouping(areaOptions.Grouping));
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
                var categoryLabels = ReadSharedCategoryLabels(areaPlot, workbook, plotVisibleOnly);
                chartLayout = await PrepareCartesianCategoryFramesAsync(chartLayout, theme, sceneChart, chartXml,
                    valueAxis, valueExtents, valueAxisOptions, GetValueAxisAutoTickTargetCount(false, true, false), axesStyle.ValueAxisRightSide,
                    categoryAxis, categoryLabels, chartLayout.CategoryAxisPlacement?.LabelPlotRatio ?? 0d,
                    chartLayout.CategoryAxisPlacement?.LabelsHighSide ?? false, fontResolver, context, preparedLegendBox,
                    areaOptions.PercentStacked ? "0%" : null).ConfigureAwait(false);
                var admitted = await AdmitCartesianRightLegendAsync(chartLayout, sceneChart, valueAxis,
                    categoryLabels, categoryPointsOnTickMarks, categoryAxisReversed, legendEntries, legendTextStyle,
                    fontResolver, ChartLegendPlacement.AreaRightLegend).ConfigureAwait(false);
                chartLayout = admitted.Layout;
                preparedLegendBox ??= admitted.Box;
                ChartPlotBox plotBox = chartLayout.PlotBox;
                RenderChartAreaStyle(graphics, document, bounds, sceneChart);
                RenderAreaChart(
                    graphics,
                    theme,
                    colorMap,
                    chartPalette,
                    chartLayout.PlotAreaBox,
                    plotBox,
                    areaSeriesVectors,
                    areaOptions,
                    categoryPointsOnTickMarks,
                    seriesFills,
                    seriesStrokes,
                    valueAxisOptions,
                    axesStyle,
                    plotAreaStyle,
                    valueExtents,
                    categoryAxisReversed, context: context);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    ChartAxisPlacement? labelPlacement = chartLayout.CategoryAxisPlacement;
                    double? labelAxisY = labelPlacement is { } placement
                        ? plotBox.Y + plotBox.Height * placement.LabelPlotRatio : null;
                    await RenderChartCategoryLabelsAsync(theme, graphics, chartLayout.Frame, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, categoryLabels, horizontalBars: false, axisCoordinate: labelAxisY, categoryLabelsOnTickMarks: categoryPointsOnTickMarks, labelsHighSide: labelPlacement?.LabelsHighSide ?? false, fontResolver: fontResolver, chartFonts: fonts, categoryLabelOffsetAdjustment: chartLayout.CategoryLabelOffsetAdjustment, prepared: chartLayout.CategoryLabels, categoryAxisReversed: categoryAxisReversed).ConfigureAwait(false);
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars: false, rightSide: axesStyle.ValueAxisRightSide, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: areaOptions.PercentStacked ? "0%" : null, fontResolver: fontResolver, chartFonts: fonts, tickPolicy: valueAxisOptions.TickPolicy, context: context).ConfigureAwait(false);
                }
                // Category ticks share the line chart's source slot boundaries;
                // value ticks retain the emitted interval lattice below.
                PptxSceneChartAxisTickMark areaCategoryTickMark = ReadSceneChartAxisMajorTickMark(categoryAxis.SceneAxis);
                if (axesStyle.CategoryAxisVisible && areaCategoryTickMark != PptxSceneChartAxisTickMark.None)
                {
                    ChartSeriesStroke areaCategoryTickStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
                    if (areaCategoryTickStroke.Alpha > 0.001d)
                    {
                        ChartTextStyle areaCategoryTickStyle = ReadSceneChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                        int areaPointCount = Math.Max(1, MaxDensePointCount(areaSeriesVectors));
                        double areaCategoryAxisY = ChartValueToPlotCoordinate(valueExtents, valueAxisOptions.CrossingValue, plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                        SetChartStroke(graphics, areaCategoryTickStroke);
                        DrawCartesianChartCategoryAxisMajorTicks(graphics, plotBox.X, plotBox.Width, areaPointCount, areaCategoryAxisY, areaCategoryTickMark, areaCategoryTickStyle.FontSize, categoryPointsOnTickMarks, chartLayout.CategoryAxisPlacement?.AxisHighSide ?? axesStyle.CategoryAxisTopSide, categoryAxisReversed);
                    }
                }

                PptxSceneChartAxisTickMark areaValueTickMark = ReadSceneChartAxisMajorTickMark(valueAxis.SceneAxis);
                if (axesStyle.ValueAxisVisible && areaValueTickMark != PptxSceneChartAxisTickMark.None)
                {
                    ChartSeriesStroke areaValueTickStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
                    if (areaValueTickStroke.Alpha > 0.001d)
                    {
                        ChartTextStyle areaValueTickStyle = ReadSceneChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                        bool areaValueLabelsVisible = IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis);
                        IReadOnlyList<double> areaTickValues = GetChartAxisTickValuesCore(valueExtents, valueAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: false, valueAxisLabelsVisible: areaValueLabelsVisible, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied), valueAxisOptions.TickPolicy, context: context);
                        double[] areaValueEdges = new double[areaTickValues.Count];
                        for (int areaValueTickIndex = 0; areaValueTickIndex < areaTickValues.Count; areaValueTickIndex++)
                        {
                            areaValueEdges[areaValueTickIndex] = ChartValueToPlotCoordinate(valueExtents, areaTickValues[areaValueTickIndex], plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                        }

                        bool areaValueLabelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(sceneChart, valueAxis.SceneAxis, valueAxis.XmlAxis, axesStyle.ValueAxisRightSide);
                        double areaValueAxisX = axesStyle.ValueAxisRightSide ? plotBox.X + plotBox.Width : plotBox.X;
                        SetChartStroke(graphics, areaValueTickStroke);
                        StrokeMajorTickSegments(graphics, areaValueEdges, areaValueAxisX, verticalSegments: false, outwardIsLowSide: !areaValueLabelsRightSide, areaValueTickMark, areaValueTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                    }
                }
                await RenderDefaultChartAxisTitlesAsync(theme, graphics, chartLayout, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                await RenderChartLegendAsync(graphics, chartLayout.Frame, plotBox, legendEntries, chartLayout.Legend, legendTextStyle, fontResolver, ChartLegendPlacement.AreaRightLegend, fonts, preparedBox: preparedLegendBox).ConfigureAwait(false);
                await RenderCategoryPointDataLabelsAsync(
                    theme, colorMap, chartPalette, graphics, chartLayout.Frame, plotBox, areaSeriesVectors, valueExtents,
                    valueAxisOptions.Reversed, categoryPointsOnTickMarks, seriesStrokes, [],
                    ReadSceneDataLabelOptions(sceneChart, areaPlot),
                    ReadSceneSeriesDataLabelOptions(sceneChart, areaPlot),
                    ReadSceneCategoryLabelVector(areaPlot, workbook, plotVisibleOnly),
                    ReadSharedChartSeriesNames(areaPlot, workbook), fontResolver, fonts, context,
                    sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds,
                    categoryAxisReversed: categoryAxisReversed, areaOptions: areaOptions,
                    areaSeriesFills: seriesFills, areaCrossingValue: valueAxisOptions.CrossingValue).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }
}
