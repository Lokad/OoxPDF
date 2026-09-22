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
    private static bool TryRenderLineChartKind(
        PdfGraphicsBuilder graphics,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        IReadOnlyList<RgbColor>? chartPalette,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
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
            IReadOnlyList<ChartIndexedNumberVector> lineSeriesVectors = ReadSceneOrXmlChartSeriesVectors(linePlot, lineChart, workbook, plotVisibleOnly);
            if (CountRenderableSeries(lineSeriesVectors) != 0)
            {
                ChartLinePlotOptions lineOptions = ReadSceneOrXmlChartLineOptions(sceneChart, linePlot, chartXml, lineChart, PptxSceneChartGrouping.Standard);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(linePlot, lineChart, theme, colorMap, ChartSeriesInheritedStrokeWidth);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(linePlot, lineChart, theme, colorMap);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, linePlot, chartXml, theme, lineChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme, colorMap);
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, linePlot, chartXml, lineChart).FirstOrDefault();
                XElement? valueAxisForScale = ResolveXmlValueAxisForSource(sceneChart, valueAxis, chartXml);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, valueAxisForScale, GetLineChartValueExtents(lineSeriesVectors, lineOptions.Stacked, lineOptions.PercentStacked), lineOptions.PercentStacked, useNearMaximumHeadroom: !lineOptions.PercentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(valueAxis.SceneAxis, valueAxisForScale, theme, valueExtents, lineOptions.PercentStacked);
                ChartLayout chartLayout = GetLineChartLayout(document, theme, bounds, chartXml, sceneChart, colorMap, workbook, plotVisibleOnly, fontResolver);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, linePlot, chartXml, lineChart);
                RenderLineChart(graphics, theme, colorMap, chartPalette, chartLayout.PlotAreaBox, plotBox, lineSeriesVectors, lineOptions, seriesStrokes, markerStyles, valueAxisOptions, axesStyle, plotAreaStyle, valueExtents,
                    ReadSceneOrXmlChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, chartXml, categoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis").FontSize);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSharedCategoryLabels(linePlot, lineChart, workbook, plotVisibleOnly), horizontalBars: false, verticalAxisY: null, categoryLabelsOnTickMarks: ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis, valueAxisForScale), categoryLabelsTopSide: false, fontResolver: fontResolver, chartFonts: fonts);
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars: false, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: lineOptions.PercentStacked ? "0%" : null, fontResolver: fontResolver, chartFonts: fonts);
                    RenderSecondaryChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, GetLineChartValueExtents(lineSeriesVectors, lineOptions.Stacked, lineOptions.PercentStacked), fontResolver, chartFonts: fonts);
                }
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
                            IReadOnlyList<double> lineTickValues = GetChartAxisTickValues(valueExtents, valueAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: false, valueAxisLabelsVisible: IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis), manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied));
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
                RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds);
                RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildStrokeLegendEntries(theme, colorMap, chartPalette, linePlot, lineChart, seriesStrokes, markerStyles, reverseOrder: lineOptions.Stacked, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default, chartFonts: fonts);
                RenderLineDataLabels(
                    theme,
                    colorMap,
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
                    ReadSharedChartSeriesNames(linePlot, lineChart, workbook), fontResolver, fonts, context, sceneChart?.Relationships, linkAnnotations, reportedHyperlinkIds);
                return true;
            }
        }
        return false;
    }
}
