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
    private static bool TryRenderAreaChartKind(
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
        XElement? areaChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Area);
        if (areaChart is not null)
        {
            PptxSceneChartPlot? areaPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Area, 0);
            IReadOnlyList<ChartIndexedNumberVector> areaSeriesVectors = ReadSharedChartSeriesVectors(areaPlot, areaChart, workbook, plotVisibleOnly);
            if (CountRenderableSeries(areaSeriesVectors) != 0)
            {
                ChartAreaPlotOptions areaOptions = ReadSceneOrXmlChartAreaOptions(sceneChart, areaPlot, chartXml, areaChart, PptxSceneChartGrouping.Standard);
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(areaPlot, areaChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(areaPlot, areaChart, theme, colorMap, ChartSeriesInheritedStrokeWidth);
                ChartLayout chartLayout = GetLineChartLayout(document, theme, bounds, chartXml, sceneChart, colorMap, workbook, plotVisibleOnly, fontResolver);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, areaPlot, chartXml, areaChart).FirstOrDefault();
                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, areaPlot, chartXml, areaChart);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetAreaChartValueExtents(areaSeriesVectors, areaOptions.Stacked, areaOptions.PercentStacked), areaOptions.PercentStacked, useNearMaximumHeadroom: areaOptions.Stacked && !areaOptions.PercentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AreaChartStackedAxisNearMaximumHeadroomRatio);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(valueAxis.SceneAxis, valueAxis.XmlAxis, theme, valueExtents, areaOptions.PercentStacked);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, areaPlot, chartXml, theme, areaChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme, colorMap);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                RenderAreaChart(
                    graphics,
                    theme,
                    colorMap,
                    chartPalette,
                    chartLayout.PlotAreaBox,
                    plotBox,
                    areaSeriesVectors,
                    areaOptions,
                    seriesFills,
                    seriesStrokes,
                    valueAxisOptions,
                    axesStyle,
                    plotAreaStyle,
                    valueExtents);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSharedCategoryLabels(areaPlot, areaChart, workbook, plotVisibleOnly), horizontalBars: false, verticalAxisY: null, categoryLabelsOnTickMarks: ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis, valueAxis.XmlAxis), categoryLabelsTopSide: false, fontResolver: fontResolver, chartFonts: fonts);
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars: false, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: areaOptions.PercentStacked ? "0%" : null, fontResolver: fontResolver, chartFonts: fonts);
                }
                // Major tick marks (Office honors out/in/cross at category data points and
                // value gridlines with the same font-relative length as bars and lines).
                ChartAxisSource tickCategoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, areaPlot, chartXml, areaChart);
                PptxSceneChartAxisTickMark areaCategoryTickMark = ReadSceneOrXmlChartAxisMajorTickMark(tickCategoryAxis.SceneAxis, tickCategoryAxis.XmlAxis);
                if (axesStyle.CategoryAxisVisible && areaCategoryTickMark != PptxSceneChartAxisTickMark.None)
                {
                    ChartSeriesStroke areaCategoryTickStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
                    if (areaCategoryTickStroke.Alpha > 0.001d)
                    {
                        ChartTextStyle areaCategoryTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, tickCategoryAxis.SceneAxis, chartXml, tickCategoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                        int areaPointCount = Math.Max(1, MaxDensePointCount(areaSeriesVectors));
                        double[] areaTickEdges = new double[areaPointCount];
                        for (int areaTickIndex = 0; areaTickIndex < areaPointCount; areaTickIndex++)
                        {
                            areaTickEdges[areaTickIndex] = plotBox.X + (areaPointCount == 1 ? plotBox.Width / 2d : plotBox.Width * areaTickIndex / (areaPointCount - 1));
                        }

                        double areaCategoryAxisY = ChartValueToPlotCoordinate(valueExtents, valueAxisOptions.CrossingValue, plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                        SetChartStroke(graphics, areaCategoryTickStroke);
                        StrokeMajorTickSegments(graphics, areaTickEdges, areaCategoryAxisY, verticalSegments: true, outwardIsLowSide: !axesStyle.CategoryAxisTopSide, areaCategoryTickMark, areaCategoryTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                    }
                }

                PptxSceneChartAxisTickMark areaValueTickMark = ReadSceneOrXmlChartAxisMajorTickMark(valueAxis.SceneAxis, valueAxis.XmlAxis);
                if (axesStyle.ValueAxisVisible && areaValueTickMark != PptxSceneChartAxisTickMark.None)
                {
                    ChartSeriesStroke areaValueTickStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
                    if (areaValueTickStroke.Alpha > 0.001d)
                    {
                        ChartTextStyle areaValueTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                        bool areaValueLabelsVisible = IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis);
                        IReadOnlyList<double> areaTickValues = GetChartAxisTickValues(valueExtents, valueAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: false, valueAxisLabelsVisible: areaValueLabelsVisible, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied));
                        double[] areaValueEdges = new double[areaTickValues.Count];
                        for (int areaValueTickIndex = 0; areaValueTickIndex < areaTickValues.Count; areaValueTickIndex++)
                        {
                            areaValueEdges[areaValueTickIndex] = ChartValueToPlotCoordinate(valueExtents, areaTickValues[areaValueTickIndex], plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                        }

                        bool areaValueLabelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(valueAxis.SceneAxis, valueAxis.XmlAxis, axesStyle.ValueAxisRightSide);
                        double areaValueAxisX = axesStyle.ValueAxisRightSide ? plotBox.X + plotBox.Width : plotBox.X;
                        SetChartStroke(graphics, areaValueTickStroke);
                        StrokeMajorTickSegments(graphics, areaValueEdges, areaValueAxisX, verticalSegments: false, outwardIsLowSide: !areaValueLabelsRightSide, areaValueTickMark, areaValueTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                    }
                }
                RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds);
                RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildFillLegendEntries(theme, colorMap, chartPalette, areaPlot, areaChart, seriesFills, seriesStrokes, paletteOffset: 0, workbook: workbook, reverseOrder: IsStackedChartGrouping(areaOptions.Grouping)), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.AreaRightLegend, fonts);
                return true;
            }
        }
        return false;
    }
}
