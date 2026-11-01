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
        XElement? lineChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Line);
        if (lineChart is not null)
        {
            PptxSceneChartPlot? linePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Line, 0);
            IReadOnlyList<ChartIndexedNumberVector> lineSeriesVectors = ReadSharedChartSeriesVectors(linePlot, lineChart, workbook, plotVisibleOnly);
            if (CountRenderableSeries(lineSeriesVectors) != 0)
            {
                ChartLinePlotOptions lineOptions = ReadSceneOrXmlChartLineOptions(sceneChart, linePlot, chartXml, lineChart, PptxSceneChartGrouping.Standard);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(linePlot, lineChart, theme, colorMap, ResolveLineKindInheritedStrokeWidth(ReadChartStyleId(chartXml)));
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(linePlot, lineChart, theme, colorMap);
                Func<int, ChartLinePointAppearance>? pointAppearance = ReadLinePointAppearance(sceneChart, linePlot, lineChart,
                    theme, colorMap, chartPalette, lineSeriesVectors, lineOptions, seriesStrokes, markerStyles, ReadChartStyleId(chartXml), context.CancellationToken);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, linePlot, chartXml, theme, lineChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme, colorMap);
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, linePlot, chartXml, lineChart).FirstOrDefault();
                XElement? valueAxisForScale = ResolveXmlValueAxisForSource(sceneChart, valueAxis, chartXml);
                ChartValueExtents valueExtents = ReadCartesianChartValueAxisExtents(sceneChart, valueAxis.SceneAxis, valueAxisForScale, GetSharedLineChartValueExtents(linePlot, lineChart, lineOptions.Stacked, lineOptions.PercentStacked, workbook, plotVisibleOnly), lineOptions.PercentStacked, useNearMaximumHeadroom: !lineOptions.PercentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(valueAxis.SceneAxis, valueAxisForScale, theme, valueExtents, lineOptions.PercentStacked) with
                { TickPolicy = ReadCartesianChartAxisTickPolicy(sceneChart, valueAxis.SceneAxis, valueExtents, lineOptions.PercentStacked) };
                ChartLayout chartLayout = (await GetLineChartLayoutAsync(document, theme, bounds, chartXml, sceneChart, colorMap, workbook, plotVisibleOnly, fontResolver, resolvedLineExtents: valueExtents).ConfigureAwait(false));
                IReadOnlyList<ChartLegendEntry> legendEntries = pointAppearance is not null
                    ? BuildLinePointLegendEntries(linePlot, lineChart, workbook, plotVisibleOnly, lineSeriesVectors[0].DensePointCount(), ReadChartStyleId(chartXml), pointAppearance)
                    : BuildStrokeLegendEntries(theme, colorMap, chartPalette, linePlot, lineChart, seriesStrokes, markerStyles, reverseOrder: lineOptions.Stacked, workbook: workbook, chartStyleId: ReadChartStyleId(chartXml));
                ChartTextStyle legendTextStyle = ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml);
                ChartLegendBox? preparedLegendBox = null;
                if (chartLayout.Legend.PositionKind == PptxSceneChartLegendPosition.Bottom)
                {
                    bool automatic = !chartLayout.ManualPlotLayoutApplied && !chartLayout.Legend.Layout.HasGeometryValues;
                    var prepared = await LayoutChartLegendRowsAsync(chartLayout, legendEntries, legendTextStyle,
                        fontResolver, context, chartLayout.PlotBox.Width, measuredBand: automatic).ConfigureAwait(false);
                    chartLayout = prepared.Layout;
                    preparedLegendBox = prepared.Box;
                }
                if (chartLayout.AutomaticCategoryAxisPlacement is { } categoryPlacement)
                {
                    valueAxisOptions = valueAxisOptions with { CrossingValue = categoryPlacement.CrossingValue };
                }
                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, linePlot, chartXml, lineChart);
                var categoryLabels = ReadSharedCategoryLabels(linePlot, lineChart, workbook, plotVisibleOnly);
                if (axesStyle.CategoryAxisVisible)
                    chartLayout = await PrepareCategoryLabelLayoutAsync(chartLayout, theme, sceneChart, chartXml, categoryAxis, categoryLabels,
                        ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis, valueAxisForScale),
                        chartLayout.AutomaticCategoryAxisPlacement?.LabelsTopSide ?? false,
                        chartLayout.AutomaticCategoryAxisPlacement?.LabelPlotRatio ?? 0d,
                        fontResolver, preparedLegendBox).ConfigureAwait(false);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                RenderLineChart(graphics, theme, colorMap, chartPalette, chartLayout.PlotAreaBox, plotBox, lineSeriesVectors, lineOptions, seriesStrokes, markerStyles, valueAxisOptions, axesStyle, plotAreaStyle, valueExtents,
                    ReadSceneOrXmlChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, chartXml, categoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis").FontSize, ReadChartStyleId(chartXml), categoryTicksTopSide: chartLayout.AutomaticCategoryAxisPlacement?.AxisTopSide ?? false, pointAppearance: pointAppearance);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    ChartCategoryAxisPlacement? labelPlacement = chartLayout.AutomaticCategoryAxisPlacement;
                    double? labelAxisY = labelPlacement is { } placement
                        ? plotBox.Y + plotBox.Height * placement.LabelPlotRatio
                        : null;
                    await RenderChartCategoryLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, categoryLabels, horizontalBars: false, verticalAxisY: labelAxisY, categoryLabelsOnTickMarks: ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis, valueAxisForScale), categoryLabelsTopSide: labelPlacement?.LabelsTopSide ?? false, fontResolver: fontResolver, chartFonts: fonts, automaticLabelOffsetAdjustment: chartLayout.AutomaticCategoryLabelOffsetAdjustment, prepared: chartLayout.CategoryLabels).ConfigureAwait(false);
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars: false, rightSide: axesStyle.ValueAxisRightSide, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: lineOptions.PercentStacked ? "0%" : null, fontResolver: fontResolver, chartFonts: fonts, tickPolicy: valueAxisOptions.TickPolicy).ConfigureAwait(false);
                }
                await RenderSecondaryChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, GetSharedLineChartValueExtents(linePlot, lineChart, lineOptions.Stacked, lineOptions.PercentStacked, workbook, plotVisibleOnly), fontResolver, chartFonts: fonts, primarySceneAxis: valueAxis.SceneAxis, primaryXmlAxis: valueAxis.XmlAxis).ConfigureAwait(false);
                // Major value tick marks (Office honors out/in/cross at value gridlines).
                if (axesStyle.ValueAxisVisible)
                {
                    PptxSceneChartAxisTickMark lineValueTickMark = ReadSceneOrXmlChartAxisMajorTickMark(valueAxis.SceneAxis, valueAxis.XmlAxis);
                    if (lineValueTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke lineValueTickStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
                        if (lineValueTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle lineValueTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                            IReadOnlyList<double> lineTickValues = GetChartAxisTickValuesCore(valueExtents, valueAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: false, valueAxisLabelsVisible: IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis), manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied), valueAxisOptions.TickPolicy);
                            double[] lineTickEdges = new double[lineTickValues.Count];
                            for (int lineTickIndex = 0; lineTickIndex < lineTickValues.Count; lineTickIndex++)
                            {
                                lineTickEdges[lineTickIndex] = ChartValueToPlotCoordinate(valueExtents, lineTickValues[lineTickIndex], plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                            }

                            bool lineValueLabelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(valueAxis.SceneAxis, valueAxis.XmlAxis, axesStyle.ValueAxisRightSide);
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
                await RenderLineDataLabelsAsync(
                    theme,
                    colorMap,
                    chartPalette,
                    graphics,
                    plotBox,
                    lineSeriesVectors,
                    valueExtents,
                    valueAxisOptions.Reversed,
                    seriesStrokes,
                    markerStyles,
                    ReadSceneOrXmlDataLabelOptions(sceneChart, linePlot, lineChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, linePlot, lineChart, theme, colorMap),
                    ReadSceneOrXmlCategoryLabelVector(linePlot, lineChart, workbook, plotVisibleOnly),
                    ReadSharedChartSeriesNames(linePlot, lineChart, workbook), fontResolver, fonts, context, sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds, pointAppearance: pointAppearance).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }
}
