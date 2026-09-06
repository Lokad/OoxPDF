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
    private static bool TryRenderScatterChartKind(
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
        XElement? scatterChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Scatter);
        if (scatterChart is not null)
        {
            PptxSceneChartPlot? scatterPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Scatter, 0);
            IReadOnlyList<ScatterSeries> scatterSeries = ReadSceneOrXmlScatterSeries(scatterPlot, scatterChart, readBubbleSize: false, workbook: workbook, plotVisibleOnly: plotVisibleOnly);
            if (scatterSeries.Count != 0)
            {
                ChartScatterPlotOptions scatterOptions = ReadSceneOrXmlChartScatterOptions(scatterPlot, scatterChart);
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(scatterPlot, scatterChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(scatterPlot, scatterChart, theme, colorMap, null);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(scatterPlot, scatterChart, theme, colorMap);
                ChartLayout chartLayout = GetLineChartLayout(document, theme, bounds, chartXml, sceneChart, colorMap, workbook, plotVisibleOnly, fontResolver);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                IReadOnlyList<ChartAxisSource> valueAxes = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, scatterPlot, chartXml, scatterChart);
                ChartAxisSource xValueAxis = valueAxes.Count > 0 ? valueAxes[0] : default;
                ChartAxisSource yValueAxis = valueAxes.Count > 1 ? valueAxes[1] : xValueAxis;
                ChartValueExtents xExtents = ReadSceneOrXmlBubbleChartValueAxisExtents(xValueAxis.SceneAxis, xValueAxis.XmlAxis, GetScatterXValueExtents(scatterSeries));
                ChartValueExtents yExtents = ReadSceneOrXmlBubbleChartValueAxisExtents(yValueAxis.SceneAxis, yValueAxis.XmlAxis, GetScatterYValueExtents(scatterSeries));
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                RenderScatterChart(graphics, theme, colorMap, chartPalette, plotBox, scatterSeries, scatterOptions.ConnectLines, bubble: false, seriesFills, seriesStrokes, markerStyles, scatterOptions.SmoothSeries, xExtents, yExtents);
                fonts.AddRange(RenderScatterDataLabels(
                    theme,
                    colorMap,
                    graphics,
                    plotBox,
                    scatterSeries,
                    bubble: false,
                    xExtents,
                    yExtents,
                    seriesFills,
                    ReadSceneOrXmlDataLabelOptions(sceneChart, scatterPlot, scatterChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, scatterPlot, scatterChart, theme, colorMap),
                    ReadSceneOrXmlChartSeriesNameRecords(scatterPlot, scatterChart, workbook), fontResolver));
                fonts.AddRange(RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildStrokeLegendEntries(theme, colorMap, chartPalette, scatterPlot, scatterChart, seriesStrokes, markerStyles: null, reverseOrder: false, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default));
                return true;
            }
        }
        return false;
    }

    private static bool TryRenderBubbleChartKind(
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
        XElement? bubbleChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Bubble);
        if (bubbleChart is not null)
        {
            PptxSceneChartPlot? bubblePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bubble, 0);
            IReadOnlyList<ScatterSeries> bubbleSeries = ReadSceneOrXmlScatterSeries(bubblePlot, bubbleChart, readBubbleSize: true, workbook: workbook, plotVisibleOnly: plotVisibleOnly);
            if (bubbleSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(bubblePlot, bubbleChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(bubblePlot, bubbleChart, theme, colorMap, null);
                ChartLayout chartLayout = GetBubbleChartLayout(document, theme, bounds, chartXml, sceneChart, bubblePlot, bubbleChart, colorMap, workbook, fontResolver);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                IReadOnlyList<ChartAxisSource> valueAxes = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, bubblePlot, chartXml, bubbleChart);
                ChartAxisSource xValueAxis = valueAxes.Count > 0 ? valueAxes[0] : default;
                ChartAxisSource yValueAxis = valueAxes.Count > 1 ? valueAxes[1] : xValueAxis;
                ChartValueExtents xExtents = ReadSceneOrXmlBubbleChartValueAxisExtents(xValueAxis.SceneAxis, xValueAxis.XmlAxis, GetBubbleXValueExtents(bubbleSeries));
                ChartValueExtents yExtents = ReadSceneOrXmlBubbleChartValueAxisExtents(yValueAxis.SceneAxis, yValueAxis.XmlAxis, GetBubbleYValueExtents(bubbleSeries));
                ChartBubbleValueAxisOptions xAxisOptions = ReadSceneOrXmlChartBubbleValueAxisOptions(xValueAxis.SceneAxis, xValueAxis.XmlAxis, theme, xExtents);
                ChartBubbleValueAxisOptions yAxisOptions = ReadSceneOrXmlChartBubbleValueAxisOptions(yValueAxis.SceneAxis, yValueAxis.XmlAxis, theme, yExtents);
                DrawHorizontalChartGridlines(graphics, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, yExtents, yAxisOptions.Units.MajorUnit, crossingValue: null, reversed: false, major: true, yAxisOptions.GridlineStyle.Major);
                RenderScatterChart(graphics, theme, colorMap, chartPalette, plotBox, bubbleSeries, connectLines: false, bubble: true, seriesFills, seriesStrokes, [], [], xExtents, yExtents);
                fonts.AddRange(RenderScatterDataLabels(
                    theme,
                    colorMap,
                    graphics,
                    plotBox,
                    bubbleSeries,
                    bubble: true,
                    xExtents,
                    yExtents,
                    seriesFills,
                    ReadSceneOrXmlDataLabelOptions(sceneChart, bubblePlot, bubbleChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, bubblePlot, bubbleChart, theme, colorMap),
                    ReadSceneOrXmlChartSeriesNameRecords(bubblePlot, bubbleChart, workbook), fontResolver));
                fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, xValueAxis.XmlAxis, xValueAxis.SceneAxis, xExtents, xAxisOptions.Units, valueAxisReversed: false, horizontalBars: true, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver));
                fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, yValueAxis.XmlAxis, yValueAxis.SceneAxis, yExtents, yAxisOptions.Units, valueAxisReversed: false, horizontalBars: false, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver));
                fonts.AddRange(RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildFillLegendEntries(theme, colorMap, chartPalette, bubblePlot, bubbleChart, seriesFills, seriesStrokes, paletteOffset: 0, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.BubbleTitleRightLegend));
                return true;
            }
        }
        return false;
    }
}
