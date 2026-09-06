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
        bool plotVisibleOnly)
    {
        XElement? areaChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Area);
        if (areaChart is not null)
        {
            PptxSceneChartPlot? areaPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Area, 0);
            IReadOnlyList<ChartIndexedNumberVector> areaSeriesVectors = ReadSceneOrXmlChartSeriesVectors(areaPlot, areaChart, workbook, plotVisibleOnly);
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
                    fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSceneOrXmlCategoryLabelVector(areaPlot, areaChart, workbook, plotVisibleOnly), horizontalBars: false, verticalAxisY: null, categoryLabelsOnTickMarks: ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis, valueAxis.XmlAxis), categoryLabelsTopSide: false, fontResolver: fontResolver));
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars: false, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: areaOptions.PercentStacked ? "0%" : null, fontResolver: fontResolver));
                }

                fonts.AddRange(RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildFillLegendEntries(theme, colorMap, chartPalette, areaPlot, areaChart, seriesFills, seriesStrokes, paletteOffset: 0, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default));
                return true;
            }
        }
        return false;
    }
}
