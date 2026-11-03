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
    private static async ValueTask<bool> TryRenderScatterChartKindAsync(
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
                IReadOnlyList<bool> seriesLineHidden = ReadSceneOrXmlSeriesLineHidden(scatterPlot, scatterChart);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(scatterPlot, scatterChart, theme, colorMap);
                IReadOnlyList<ChartAxisSource> valueAxes = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, scatterPlot, chartXml, scatterChart);
                ChartAxisSource xValueAxis = valueAxes.Count > 0 ? valueAxes[0] : default;
                ChartAxisSource yValueAxis = valueAxes.Count > 1 ? valueAxes[1] : xValueAxis;
                double? xTickTargetCount = ReadScatterSourceIntervalTarget(theme, sceneChart, chartXml, GetChartFrameBox(document, bounds), xValueAxis);
                ChartLayout chartLayout = (await GetCartesianChartLayoutAsync(document, theme, bounds, chartXml, sceneChart, colorMap, workbook, plotVisibleOnly, fontResolver, resolvedValueExtents: null, xTickTargetCount).ConfigureAwait(false));
                ChartPlotBox plotBox = chartLayout.PlotBox;
                var measured = MeasureScatterChartExtents(scatterSeries, scatterOptions, seriesLineHidden, context.CancellationToken);
                var (xExtents, xAxisOptions) = ReadChartPointValueAxisScale(theme, xValueAxis, measured.DataX, measured.PathX, bubble: false, horizontal: true, xTickTargetCount);
                var (yExtents, yAxisOptions) = ReadChartPointValueAxisScale(theme, yValueAxis, measured.DataY, measured.PathY, bubble: false, horizontal: false);
                ChartAxisPlacement xPlacement = ReadChartAxisPlacement(xValueAxis, yExtents, yAxisOptions.Reversed) ?? new(yExtents.Min, 0d, false, false);
                ChartAxisPlacement yPlacement = ReadChartAxisPlacement(yValueAxis, xExtents, xAxisOptions.Reversed, horizontal: false) ?? new(xExtents.Min, 0d, false, false);
                chartLayout = await AdmitScatterSourcePlotAsync(theme, sceneChart, chartXml, chartLayout,
                    xValueAxis, yValueAxis, xExtents, yExtents, xAxisOptions, yAxisOptions, yPlacement, fontResolver).ConfigureAwait(false);
                plotBox = chartLayout.PlotBox;
                // Each axis declares its crossing in the perpendicular scale.
                xAxisOptions = xAxisOptions with { CrossingValue = Math.Clamp(yPlacement.CrossingValue, xExtents.Min, xExtents.Max) };
                yAxisOptions = yAxisOptions with { CrossingValue = Math.Clamp(xPlacement.CrossingValue, yExtents.Min, yExtents.Max) };
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                DrawHorizontalChartGridlines(graphics, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, yExtents, yAxisOptions.Units, yAxisOptions.CrossingValue, yAxisOptions.Reversed, major: true, yAxisOptions.GridlineStyle.Major, hasMajorGridlineElement: true, yAxisOptions.TickPolicy);
                RenderScatterChart(graphics, theme, colorMap, chartPalette, plotBox, scatterSeries, scatterOptions.ConnectLines, bubble: false, seriesFills, seriesStrokes, markerStyles, scatterOptions.SmoothSeries, seriesLineHidden, xExtents, yExtents, xAxisOptions, yAxisOptions);
                await RenderScatterDataLabelsAsync(
                    theme,
                    colorMap,
                    graphics,
                    plotBox,
                    scatterSeries,
                    bubble: false,
                    xExtents,
                    yExtents,
                    xAxisOptions.Reversed,
                    yAxisOptions.Reversed,
                    seriesFills,
                    ReadSceneOrXmlDataLabelOptions(sceneChart, scatterPlot, scatterChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, scatterPlot, scatterChart, theme, colorMap),
                    ReadSharedChartSeriesNames(scatterPlot, scatterChart, workbook), fontResolver, fonts, context, sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, xValueAxis.XmlAxis, xValueAxis.SceneAxis, xExtents, xAxisOptions.Units, xAxisOptions.Reversed, horizontalBars: true, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver, chartFonts: fonts, tickPolicy: xAxisOptions.TickPolicy, placement: xPlacement).ConfigureAwait(false);
                await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, yValueAxis.XmlAxis, yValueAxis.SceneAxis, yExtents, yAxisOptions.Units, yAxisOptions.Reversed, horizontalBars: false, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver, chartFonts: fonts, tickPolicy: yAxisOptions.TickPolicy, placement: yPlacement).ConfigureAwait(false);
                await RenderDefaultChartAxisTitlesAsync(theme, graphics, chartLayout, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                RenderScatterBubbleMajorTicks(graphics, theme, chartXml, sceneChart, plotBox, xValueAxis, yValueAxis, xExtents, yExtents, xAxisOptions, yAxisOptions, xPlacement, yPlacement);
                double scatterLegendLeadExtra = 0.5d * (await MeasureScatterLegendEndLabelWidthAsync(theme, sceneChart, chartXml, scatterPlot, scatterChart, workbook, plotVisibleOnly, fontResolver, xTickTargetCount).ConfigureAwait(false));
                await RenderChartLegendAsync(graphics, chartLayout.Frame, plotBox, BuildStrokeLegendEntries(theme, colorMap, chartPalette, scatterPlot, scatterChart, seriesStrokes, markerStyles, reverseOrder: false, workbook: workbook, seriesLineHidden: seriesLineHidden), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default, chartFonts: fonts, legendLeadExtra: scatterLegendLeadExtra, scatterLegend: true).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }

    private static async ValueTask<ChartLayout> AdmitScatterSourcePlotAsync(PptxTheme theme, PptxSceneChart chart, XDocument xml,
        ChartLayout layout, ChartAxisSource xAxis, ChartAxisSource yAxis, ChartValueExtents xExtents, ChartValueExtents yExtents,
        ChartValueAxisRenderOptions xOptions, ChartValueAxisRenderOptions yOptions, ChartAxisPlacement yPlacement,
        PresentationFontResolver fontResolver)
    {
        PptxSceneChartManualLayout source = chart.PlotAreaLayout;
        if (source.LayoutTargetKind != PptxSceneChartManualLayoutTarget.Inner || source.X is null || source.Width is null ||
            source.XModeKind != PptxSceneChartManualLayoutMode.Factor || source.WidthModeKind != PptxSceneChartManualLayoutMode.Factor)
            return layout;
        var measurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
        double firstWidth = 0d, lastWidth = 0d, sideWidth = 0d;
        if (IsSceneOrXmlChartAxisLabelVisible(xAxis.SceneAxis, xAxis.XmlAxis))
        {
            ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, chart, xAxis.SceneAxis, xml, xAxis.XmlAxis,
                PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
            var captions = await MeasureChartValueAxisCaptionsAsync(measurer, style, xAxis, xExtents, xOptions.Units,
                PptxChartMetricRules.AxisNiceHorizontalValueTickTargetCount, xOptions.TickPolicy, null, reserveEndpoints: false).ConfigureAwait(false);
            if (captions.Values.Count > 0)
            {
                firstWidth = measurer.Measure(FormatSceneOrXmlChartAxisLabel(captions.Values[0], xAxis.SceneAxis, xAxis.XmlAxis, null), style);
                lastWidth = measurer.Measure(FormatSceneOrXmlChartAxisLabel(captions.Values[^1], xAxis.SceneAxis, xAxis.XmlAxis, null), style);
                if (xOptions.Reversed) (firstWidth, lastWidth) = (lastWidth, firstWidth);
            }
        }
        if (IsSceneOrXmlChartAxisLabelVisible(yAxis.SceneAxis, yAxis.XmlAxis))
        {
            ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, chart, yAxis.SceneAxis, xml, yAxis.XmlAxis,
                PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
            var captions = await MeasureChartValueAxisCaptionsAsync(measurer, style, yAxis, yExtents, yOptions.Units,
                PptxChartMetricRules.AxisNiceVerticalValueTickTargetCount, yOptions.TickPolicy, null, reserveEndpoints: true).ConfigureAwait(false);
            sideWidth = captions.SlotWidth;
        }
        ChartPlotBox admitted = AdmitChartPlotCaptions(layout.Frame, layout.PlotBox, firstWidth, lastWidth, sideWidth, yPlacement);
        return layout with
        {
            PlotBox = admitted,
            PlotAreaBox = layout.PlotAreaBox with { X = admitted.X, Width = admitted.Width }
        };
    }

    private static IReadOnlyList<bool> ReadSceneOrXmlSeriesLineHidden(PptxSceneChartPlot? plot, XElement chartElement)
    {
        // An explicit series line noFill suppresses the connecting path even when the
        // scatter style itself requests lines; an absent line definition still connects.
        if (plot is not null)
        {
            return plot.Series.Select(series => series.Line.NoFill).ToArray();
        }

        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => series.Element(ChartNamespace + "spPr")?.Element(DrawingNamespace + "ln")?.Element(DrawingNamespace + "noFill") is not null)
            .ToArray();
    }

    // Shared scatter/bubble major ticks (Office honors out/in/cross at both value axes
    // with the same font-relative length as bars and lines).
    private static void RenderScatterBubbleMajorTicks(PdfGraphicsBuilder graphics, PptxTheme theme, XDocument chartXml, PptxSceneChart? sceneChart, ChartPlotBox plotBox, ChartAxisSource xValueAxis, ChartAxisSource yValueAxis, ChartValueExtents xExtents, ChartValueExtents yExtents, ChartValueAxisRenderOptions xAxisOptions, ChartValueAxisRenderOptions yAxisOptions, ChartAxisPlacement xPlacement, ChartAxisPlacement yPlacement)
    {
                PptxSceneChartAxisTickMark xTickMark = ReadSceneOrXmlChartAxisMajorTickMark(xValueAxis.SceneAxis, xValueAxis.XmlAxis);
                if (xTickMark != PptxSceneChartAxisTickMark.None)
                {
                    ChartSeriesStroke xTickStroke = xAxisOptions.GridlineStyle.Major ?? DefaultChartGridlineStroke(major: true);
                    if (xTickStroke.Alpha > 0.001d)
                    {
                        ChartTextStyle xTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, xValueAxis.SceneAxis, chartXml, xValueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                        IReadOnlyList<double> xTickValues = GetChartAxisTickValuesCore(xExtents, xAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: true, valueAxisLabelsVisible: IsSceneOrXmlChartAxisLabelVisible(xValueAxis.SceneAxis, xValueAxis.XmlAxis), manualPlotLayoutApplied: false), xAxisOptions.TickPolicy);
                        double[] xTickEdges = new double[xTickValues.Count];
                        for (int xTickIndex = 0; xTickIndex < xTickValues.Count; xTickIndex++)
                        {
                            xTickEdges[xTickIndex] = ChartValueToPlotCoordinate(xExtents, xTickValues[xTickIndex], plotBox.X, plotBox.Width, xAxisOptions.Reversed);
                        }

                        SetChartStroke(graphics, xTickStroke);
                        double axisY = ChartValueToPlotCoordinate(yExtents, yAxisOptions.CrossingValue, plotBox.Y, plotBox.Height, yAxisOptions.Reversed);
                        StrokeMajorTickSegments(graphics, xTickEdges, axisY, verticalSegments: true, outwardIsLowSide: !xPlacement.AxisHighSide, xTickMark, xTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                    }
                }

                PptxSceneChartAxisTickMark yTickMark = ReadSceneOrXmlChartAxisMajorTickMark(yValueAxis.SceneAxis, yValueAxis.XmlAxis);
                if (yTickMark != PptxSceneChartAxisTickMark.None)
                {
                    ChartSeriesStroke yTickStroke = yAxisOptions.GridlineStyle.Major ?? DefaultChartGridlineStroke(major: true);
                    if (yTickStroke.Alpha > 0.001d)
                    {
                        ChartTextStyle yTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, yValueAxis.SceneAxis, chartXml, yValueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                        IReadOnlyList<double> yTickValues = GetChartAxisTickValuesCore(yExtents, yAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: false, valueAxisLabelsVisible: IsSceneOrXmlChartAxisLabelVisible(yValueAxis.SceneAxis, yValueAxis.XmlAxis), manualPlotLayoutApplied: false), yAxisOptions.TickPolicy);
                        double[] yTickEdges = new double[yTickValues.Count];
                        for (int yTickIndex = 0; yTickIndex < yTickValues.Count; yTickIndex++)
                        {
                            yTickEdges[yTickIndex] = ChartValueToPlotCoordinate(yExtents, yTickValues[yTickIndex], plotBox.Y, plotBox.Height, yAxisOptions.Reversed);
                        }

                        SetChartStroke(graphics, yTickStroke);
                        double axisX = ChartValueToPlotCoordinate(xExtents, xAxisOptions.CrossingValue, plotBox.X, plotBox.Width, xAxisOptions.Reversed);
                        StrokeMajorTickSegments(graphics, yTickEdges, axisX, verticalSegments: false, outwardIsLowSide: !yPlacement.AxisHighSide, yTickMark, yTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                    }
                }
    }

    private static async ValueTask<bool> TryRenderBubbleChartKindAsync(
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
        XElement? bubbleChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Bubble);
        if (bubbleChart is not null)
        {
            PptxSceneChartPlot? bubblePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bubble, 0);
            IReadOnlyList<ScatterSeries> bubbleSeries = ReadSceneOrXmlScatterSeries(bubblePlot, bubbleChart, readBubbleSize: true, workbook: workbook, plotVisibleOnly: plotVisibleOnly);
            if (bubbleSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(bubblePlot, bubbleChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(bubblePlot, bubbleChart, theme, colorMap, null);
                ChartLayout chartLayout = GetBubbleChartLayout(document, theme, bounds, chartXml, sceneChart, colorMap);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                IReadOnlyList<ChartAxisSource> valueAxes = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, bubblePlot, chartXml, bubbleChart);
                ChartAxisSource xValueAxis = valueAxes.Count > 0 ? valueAxes[0] : default;
                ChartAxisSource yValueAxis = valueAxes.Count > 1 ? valueAxes[1] : xValueAxis;
                var measured = MeasureScatterChartExtents(bubbleSeries, null, null, context.CancellationToken);
                var (xExtents, xAxisOptions) = ReadChartPointValueAxisScale(theme, xValueAxis, measured.DataX, measured.PathX, bubble: true, horizontal: true);
                var (yExtents, yAxisOptions) = ReadChartPointValueAxisScale(theme, yValueAxis, measured.DataY, measured.PathY, bubble: true, horizontal: false);
                IReadOnlyList<ChartLegendEntry> legendEntries = BuildFillLegendEntries(theme, colorMap, chartPalette, bubblePlot, bubbleChart,
                    seriesFills, seriesStrokes, paletteOffset: 0, workbook: workbook);
                ChartBubbleSizing? bubbleSizing = null;
                ChartLegendBox? preparedLegend = null;
                if (await TryPrepareBubbleGeometryAsync(theme, sceneChart, chartXml, chartLayout, bubblePlot, bubbleSeries,
                        legendEntries, xValueAxis, yValueAxis, measured.DataX, measured.DataY, fontResolver).ConfigureAwait(false) is { } prepared)
                {
                    chartLayout = prepared.Layout;
                    plotBox = chartLayout.PlotBox;
                    (xExtents, yExtents, xAxisOptions, yAxisOptions) = (prepared.X, prepared.Y, prepared.XOptions, prepared.YOptions);
                    bubbleSizing = prepared.Sizing;
                    preparedLegend = prepared.LegendBox;
                    legendEntries = legendEntries.Select(entry => entry with
                    {
                        Marker = new ChartMarkerStyle(PptxSceneChartMarkerSymbol.Circle, "circle", null, 0d, entry.Fill, entry.Stroke, true)
                    }).ToArray();
                }
                ChartAxisPlacement xPlacement = ReadChartAxisPlacement(xValueAxis, yExtents, yAxisOptions.Reversed) ?? new(yExtents.Min, 0d, false, false);
                ChartAxisPlacement yPlacement = ReadChartAxisPlacement(yValueAxis, xExtents, xAxisOptions.Reversed, horizontal: false) ?? new(xExtents.Min, 0d, false, false);
                xAxisOptions = xAxisOptions with { CrossingValue = Math.Clamp(yPlacement.CrossingValue, xExtents.Min, xExtents.Max) };
                yAxisOptions = yAxisOptions with { CrossingValue = Math.Clamp(xPlacement.CrossingValue, yExtents.Min, yExtents.Max) };
                DrawHorizontalChartGridlines(graphics, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, yExtents, yAxisOptions.Units, yAxisOptions.CrossingValue, yAxisOptions.Reversed, major: true, yAxisOptions.GridlineStyle.Major, hasMajorGridlineElement: true, yAxisOptions.TickPolicy);
                RenderScatterChart(graphics, theme, colorMap, chartPalette, plotBox, bubbleSeries, connectLines: false, bubble: true, seriesFills, seriesStrokes, [], [], [], xExtents, yExtents, xAxisOptions, yAxisOptions, bubbleSizing);
                await RenderScatterDataLabelsAsync(
                    theme,
                    colorMap,
                    graphics,
                    plotBox,
                    bubbleSeries,
                    bubble: true,
                    xExtents,
                    yExtents,
                    xAxisOptions.Reversed,
                    yAxisOptions.Reversed,
                    seriesFills,
                    ReadSceneOrXmlDataLabelOptions(sceneChart, bubblePlot, bubbleChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, bubblePlot, bubbleChart, theme, colorMap),
                    ReadSharedChartSeriesNames(bubblePlot, bubbleChart, workbook), fontResolver, fonts, context, sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds,
                    bubbleSizing: bubbleSizing).ConfigureAwait(false);
                await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, xValueAxis.XmlAxis, xValueAxis.SceneAxis, xExtents, xAxisOptions.Units, xAxisOptions.Reversed, horizontalBars: true, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver, chartFonts: fonts, tickPolicy: xAxisOptions.TickPolicy, placement: xPlacement).ConfigureAwait(false);
                await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, yValueAxis.XmlAxis, yValueAxis.SceneAxis, yExtents, yAxisOptions.Units, yAxisOptions.Reversed, horizontalBars: false, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver, chartFonts: fonts, tickPolicy: yAxisOptions.TickPolicy, placement: yPlacement).ConfigureAwait(false);
                await RenderDefaultChartAxisTitlesAsync(theme, graphics, chartLayout, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                RenderScatterBubbleMajorTicks(graphics, theme, chartXml, sceneChart, plotBox, xValueAxis, yValueAxis, xExtents, yExtents, xAxisOptions, yAxisOptions, xPlacement, yPlacement);
                await RenderChartLegendAsync(graphics, chartLayout.Frame, plotBox, legendEntries, chartLayout.Legend,
                    ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver,
                    ChartLegendPlacement.BubbleTitleRightLegend, chartFonts: fonts, preparedBox: preparedLegend).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }
}
