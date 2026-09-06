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
    private const string WorkbookRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    private const string SharedStringsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    private const string SpreadsheetStylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    private const string SpreadsheetTableRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    private const string SpreadsheetStylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    private const double ChartLineDefaultStrokeWidth = 2.25d;
    private const double ChartSeriesInheritedStrokeWidth = 3d;
    private const double ChartFilledSeriesInheritedStrokeWidth = 0.75d;
    private const double ChartMarkerInheritedStrokeWidth = 0.75d;

    private static bool TryRenderChart(PdfGraphicsBuilder graphics, PptxDocument document, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, ChartWorkbookData? workbook, List<PdfFontResource> fonts, PresentationFontResolver fontResolver)
    {
        bool plotVisibleOnly = ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, chartXml);
        if (TryRenderBarChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly))
        {
            return true;
        }

        if (TryRenderLineChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly))
        {
            return true;
        }

        if (TryRenderAreaChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly))
        {
            return true;
        }

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

        XElement? radarChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Radar);
        if (radarChart is not null)
        {
            PptxSceneChartPlot? radarPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Radar, 0);
            IReadOnlyList<ChartIndexedNumberVector> radarSeriesVectors = ReadSceneOrXmlChartSeriesVectors(radarPlot, radarChart, workbook, plotVisibleOnly);
            IReadOnlyList<ChartRadarSeries> radarSeries = BuildRadarSeries(radarSeriesVectors);
            if (radarSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(radarPlot, radarChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(radarPlot, radarChart, theme, colorMap, null);
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, radarPlot, chartXml, radarChart).FirstOrDefault();
                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, radarPlot, chartXml, radarChart);
                ChartValueExtents valueExtents = ReadSceneOrXmlChartValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetRadarChartValueExtents(radarSeries), false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartAxisUnits axisUnits = ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxis.XmlAxis);
                ChartPlotBox plotBox = GetPolarChartPlotBox(document, bounds, chartXml, sceneChart);
                ChartRadarPlotOptions radarOptions = ReadSceneOrXmlChartRadarOptions(radarPlot, radarChart);
                ChartRadarLayout radarLayout = ResolveRadarLayout(plotBox, radarOptions.RadarStyle, radarSeries);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                RenderRadarChart(graphics, theme, colorMap, chartPalette, radarLayout, radarSeries, seriesFills, seriesStrokes, valueExtents, axisUnits);
                if (IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    fonts.AddRange(RenderRadarCategoryLabels(theme, graphics, radarLayout, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSceneOrXmlCategoryLabelVector(radarPlot, radarChart, workbook, plotVisibleOnly), fontResolver));
                }

                if (IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    fonts.AddRange(RenderRadarValueAxisLabels(theme, graphics, radarLayout, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, axisUnits, fontResolver));
                }

                return true;
            }
        }

        XElement? pieChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Pie);
        if (pieChart is not null)
        {
            PptxSceneChartPlot? piePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Pie, 0);
            IReadOnlyList<ChartIndexedNumberVector> pieSeriesVectors = ReadSceneOrXmlChartSeriesVectors(piePlot, pieChart, workbook, plotVisibleOnly);
            IReadOnlyList<ChartIndexedPieSlice> pieSlices = pieSeriesVectors.Count == 0 ? [] : BuildChartIndexedPieSlices(pieSeriesVectors[0]);
            if (pieSlices.Count != 0)
            {
                ChartIndexedTextVector categoryLabels = ReadSceneOrXmlCategoryLabelVector(piePlot, pieChart, workbook, plotVisibleOnly);
                IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSceneOrXmlChartSeriesNameRecords(piePlot, pieChart, workbook);
                ChartDataLabelOptions labelOptions = ResolveChartDataLabelOptionsForSeries(
                    ReadSceneOrXmlDataLabelOptions(sceneChart, piePlot, pieChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, piePlot, pieChart, theme, colorMap),
                    seriesIndex: 0);
                ChartPolarPointOptions polarPoints = ReadSceneOrXmlChartPolarPointOptions(piePlot, pieChart, theme, colorMap, workbook);
                ChartFrameBox frame = GetChartFrameBox(document, bounds);
                ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
                ChartPlotBox plotBox = GetPolarChartPlotBox(document, bounds, chartXml, sceneChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Pie, plotBox, polarPoints.PointExplosions, legend);
                RenderPieChart(graphics, theme, colorMap, chartPalette, polarLayout, pieSlices, polarPoints.PointFills, polarPoints.PointStrokes, polarPoints.PointExplosions, polarPoints.FirstSliceAngle);
                fonts.AddRange(RenderPieDataLabels(theme, colorMap, graphics, chartPalette, polarLayout, pieSlices, polarPoints.PointFills, polarPoints.PointExplosions, 0d, pieSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, piePlot, pieChart, polarPoints.PointFills, workbook, plotVisibleOnly), legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default));
                return true;
            }
        }

        XElement? doughnutChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Doughnut);
        if (doughnutChart is not null)
        {
            PptxSceneChartPlot? doughnutPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Doughnut, 0);
            IReadOnlyList<ChartIndexedNumberVector> doughnutSeriesVectors = ReadSceneOrXmlChartSeriesVectors(doughnutPlot, doughnutChart, workbook, plotVisibleOnly);
            IReadOnlyList<ChartIndexedPieSlice> doughnutSlices = doughnutSeriesVectors.Count == 0 ? [] : BuildChartIndexedPieSlices(doughnutSeriesVectors[0]);
            if (doughnutSlices.Count != 0)
            {
                ChartIndexedTextVector categoryLabels = ReadSceneOrXmlCategoryLabelVector(doughnutPlot, doughnutChart, workbook, plotVisibleOnly);
                IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSceneOrXmlChartSeriesNameRecords(doughnutPlot, doughnutChart, workbook);
                ChartDataLabelOptions labelOptions = ResolveChartDataLabelOptionsForSeries(
                    ReadSceneOrXmlDataLabelOptions(sceneChart, doughnutPlot, doughnutChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, doughnutPlot, doughnutChart, theme, colorMap),
                    seriesIndex: 0);
                ChartDoughnutPlotOptions doughnutOptions = ReadSceneOrXmlChartDoughnutOptions(doughnutPlot, doughnutChart, theme, colorMap, workbook);
                ChartFrameBox frame = GetChartFrameBox(document, bounds);
                ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
                ChartPlotBox plotBox = GetPolarChartPlotBox(document, bounds, chartXml, sceneChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPolarPointOptions polarPoints = doughnutOptions.PolarPoints;
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Doughnut, plotBox, polarPoints.PointExplosions, legend);
                RenderDoughnutChart(graphics, theme, colorMap, chartPalette, polarLayout, doughnutSlices, polarPoints.PointFills, polarPoints.PointStrokes, polarPoints.PointExplosions, doughnutOptions.HoleSize, polarPoints.FirstSliceAngle);
                fonts.AddRange(RenderPieDataLabels(theme, colorMap, graphics, chartPalette, polarLayout, doughnutSlices, polarPoints.PointFills, polarPoints.PointExplosions, doughnutOptions.HoleSize, doughnutSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, doughnutPlot, doughnutChart, polarPoints.PointFills, workbook, plotVisibleOnly), legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default));
                return true;
            }
        }

        return false;
    }

    private static RgbColor ChartPalette(IReadOnlyList<RgbColor>? chartPalette, PptxTheme? theme, int index)
    {
        return ChartPalette(chartPalette, theme, PptxColorMap.Default, index);
    }

    private static RgbColor ChartPalette(IReadOnlyList<RgbColor>? chartPalette, PptxTheme? theme, PptxColorMap colorMap, int index)
    {
        if (chartPalette is { Count: > 0 })
        {
            return chartPalette[index % chartPalette.Count];
        }

        if (theme is not null && theme.TryResolveColor("accent" + (index % 6 + 1).ToString(CultureInfo.InvariantCulture), colorMap, out RgbColor themeColor))
        {
            return themeColor;
        }

        RgbColor[] palette =
        [
            new RgbColor(68, 114, 196),
            new RgbColor(237, 125, 49),
            new RgbColor(165, 165, 165),
            new RgbColor(255, 192, 0),
            new RgbColor(91, 155, 213),
            new RgbColor(112, 173, 71)
        ];
        return palette[index % palette.Length];
    }

    private static RgbColor ChartPalette(int index)
    {
        return ChartPalette(null, null, index);
    }

    private static RgbColor ChartPalette(PptxTheme? theme, int index)
    {
        return ChartPalette(null, theme, index);
    }

    private static void StrokeChartPointRectangleInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, int seriesIndex, int categoryIndex, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, double x, double y, double width, double height, ChartSeriesStroke? fallbackStroke)
    {
        bool hasExplicitStroke = seriesIndex < pointStrokes.Count && pointStrokes[seriesIndex].ContainsKey(categoryIndex);
        if (!hasExplicitStroke && fallbackStroke is null)
        {
            return;
        }

        RenderInChartPlotAreaClip(graphics, plotBox, () => StrokeChartPointRectangle(graphics, seriesIndex, categoryIndex, pointStrokes, x, y, width, height, fallbackStroke));
    }


}
