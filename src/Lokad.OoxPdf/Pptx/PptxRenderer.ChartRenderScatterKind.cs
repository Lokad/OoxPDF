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
        PptxSceneChartPlot? scatterPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Scatter, 0);
        if (scatterPlot is not null)
        {
            XElement scatterChart = scatterPlot.Source;
            IReadOnlyList<ScatterSeries> scatterSeries = ReadSceneOrXmlScatterSeries(scatterPlot, scatterChart, readBubbleSize: false, workbook: workbook, plotVisibleOnly: plotVisibleOnly);
            if (scatterSeries.Count != 0)
            {
                ChartScatterPlotOptions scatterOptions = ReadSceneChartScatterOptions(scatterPlot);
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(scatterPlot, scatterChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(scatterPlot, scatterChart, theme, colorMap, null);
                IReadOnlyList<bool> seriesLineHidden = ReadSceneSeriesLineHidden(scatterPlot);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(scatterPlot, scatterChart, theme, colorMap);
                IReadOnlyList<ChartAxisSource> valueAxes = ReadChartValueAxesForPlot(sceneChart, scatterPlot);
                ChartAxisSource xValueAxis = valueAxes.Count > 0 ? valueAxes[0] : default;
                ChartAxisSource yValueAxis = valueAxes.Count > 1 ? valueAxes[1] : xValueAxis;
                double? xTickTargetCount = ReadScatterSourceIntervalTarget(theme, sceneChart, GetChartFrameBox(document, bounds), xValueAxis);
                ChartLayout chartLayout = (await GetCartesianChartLayoutAsync(document, theme, bounds, chartXml, sceneChart, colorMap, workbook, plotVisibleOnly, fontResolver, resolvedValueExtents: null, xTickTargetCount, context: context).ConfigureAwait(false));
                ChartPlotBox plotBox = chartLayout.PlotBox;
                var measured = MeasureScatterChartExtents(scatterSeries, scatterOptions, seriesLineHidden, context.CancellationToken);
                var (xExtents, xAxisOptions) = ReadChartPointValueAxisScale(theme, xValueAxis, measured.DataX, measured.PathX, bubble: false, horizontal: true, xTickTargetCount);
                var (yExtents, yAxisOptions) = ReadChartPointValueAxisScale(theme, yValueAxis, measured.DataY, measured.PathY, bubble: false, horizontal: false);
                ChartAxisPlacement xPlacement = ReadChartAxisPlacement(xValueAxis.SceneAxis, yExtents, yAxisOptions.Reversed) ?? new(yExtents.Min, 0d, false, false);
                ChartAxisPlacement yPlacement = ReadChartAxisPlacement(yValueAxis.SceneAxis, xExtents, xAxisOptions.Reversed, horizontal: false) ?? new(xExtents.Min, 0d, false, false);
                chartLayout = await AdmitScatterSourcePlotAsync(theme, sceneChart, chartXml, chartLayout,
                    xValueAxis, yValueAxis, xExtents, yExtents, xAxisOptions, yAxisOptions, xPlacement, yPlacement, fontResolver, context: context).ConfigureAwait(false);
                plotBox = chartLayout.PlotBox;
                // Each axis declares its crossing in the perpendicular scale.
                xAxisOptions = xAxisOptions with { CrossingValue = Math.Clamp(yPlacement.CrossingValue, xExtents.Min, xExtents.Max) };
                yAxisOptions = yAxisOptions with { CrossingValue = Math.Clamp(xPlacement.CrossingValue, yExtents.Min, yExtents.Max) };
                RenderChartAreaStyle(graphics, document, bounds, sceneChart);
                DrawHorizontalChartGridlines(graphics, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, yExtents, yAxisOptions.Units, yAxisOptions.CrossingValue, yAxisOptions.Reversed, major: true, yAxisOptions.GridlineStyle.Major, hasMajorGridlineElement: true, yAxisOptions.TickPolicy, context: context);
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
                await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, xValueAxis.XmlAxis, xValueAxis.SceneAxis, xExtents, xAxisOptions.Units, xAxisOptions.Reversed, horizontalBars: true, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver, chartFonts: fonts, tickPolicy: xAxisOptions.TickPolicy, placement: xPlacement, context: context).ConfigureAwait(false);
                await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, yValueAxis.XmlAxis, yValueAxis.SceneAxis, yExtents, yAxisOptions.Units, yAxisOptions.Reversed, horizontalBars: false, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver, chartFonts: fonts, tickPolicy: yAxisOptions.TickPolicy, placement: yPlacement, context: context).ConfigureAwait(false);
                await RenderDefaultChartAxisTitlesAsync(theme, graphics, chartLayout, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                RenderScatterBubbleMajorTicks(graphics, theme, chartXml, sceneChart, plotBox, xValueAxis, yValueAxis, xExtents, yExtents, xAxisOptions, yAxisOptions, xPlacement, yPlacement, context: context);
                double scatterLegendLeadExtra = 0.5d * (await MeasureScatterLegendEndLabelWidthAsync(theme, sceneChart, chartXml, scatterPlot, scatterChart, workbook, plotVisibleOnly, fontResolver, xTickTargetCount, context: context).ConfigureAwait(false));
                await RenderChartLegendAsync(graphics, chartLayout.Frame, plotBox, BuildStrokeLegendEntries(theme, colorMap, chartPalette, scatterPlot, scatterChart, seriesStrokes, markerStyles, reverseOrder: false, workbook: workbook, seriesLineHidden: seriesLineHidden), chartLayout.Legend, ReadSceneChartLegendTextStyle(theme, sceneChart), fontResolver, ChartLegendPlacement.Default, chartFonts: fonts, legendLeadExtra: scatterLegendLeadExtra, scatterLegend: true).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }

    private static async ValueTask<ChartLayout> AdmitScatterSourcePlotAsync(PptxTheme theme, PptxSceneChart chart, XDocument xml,
        ChartLayout layout, ChartAxisSource xAxis, ChartAxisSource yAxis, ChartValueExtents xExtents, ChartValueExtents yExtents,
        ChartValueAxisRenderOptions xOptions, ChartValueAxisRenderOptions yOptions, ChartAxisPlacement xPlacement, ChartAxisPlacement yPlacement,
        PresentationFontResolver fontResolver, PptxRenderContext? context = null)
    {
        PptxSceneChartManualLayout source = chart.PlotAreaLayout;
        bool outer = source.LayoutTargetKind == PptxSceneChartManualLayoutTarget.Outer;
        if (source.X is null || (!outer && source.XModeKind != PptxSceneChartManualLayoutMode.Factor) ||
            (outer && (chart.Plots.Count != 1 || source.Y is null || source.HeightModeKind != PptxSceneChartManualLayoutMode.Factor || source.Height is null)) ||
            !TryResolveCartesianPlotWidth(source, layout.Frame, layout.PlotBox, out _))
            return layout;
        var measurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
        double firstWidth = 0d, lastWidth = 0d, sideWidth = 0d, xLine = 0d, yLine = 0d;
        if (IsSceneOrXmlChartAxisLabelVisible(xAxis.SceneAxis, xAxis.XmlAxis))
        {
            ChartTextStyle style = ReadSceneChartTextStyle(theme, chart, xAxis.SceneAxis,
                PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
            if (outer)
            {
                if (await measurer.TryMeasureNormalLineHeightAsync(style).ConfigureAwait(false) is not { } line) return layout;
                xLine = line;
            }
            var captions = await MeasureChartValueAxisCaptionsAsync(measurer, style, xAxis, xExtents, xOptions.Units,
                PptxChartMetricRules.AxisNiceHorizontalValueTickTargetCount, xOptions.TickPolicy, null, reserveEndpoints: false, context: context).ConfigureAwait(false);
            if (captions.Values.Count > 0)
            {
                firstWidth = measurer.Measure(FormatSceneOrXmlChartAxisLabel(captions.Values[0], xAxis.SceneAxis, xAxis.XmlAxis, null), style);
                lastWidth = measurer.Measure(FormatSceneOrXmlChartAxisLabel(captions.Values[^1], xAxis.SceneAxis, xAxis.XmlAxis, null), style);
                if (xOptions.Reversed) (firstWidth, lastWidth) = (lastWidth, firstWidth);
            }
        }
        if (IsSceneOrXmlChartAxisLabelVisible(yAxis.SceneAxis, yAxis.XmlAxis))
        {
            ChartTextStyle style = ReadSceneChartTextStyle(theme, chart, yAxis.SceneAxis,
                PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
            var captions = await MeasureChartValueAxisCaptionsAsync(measurer, style, yAxis, yExtents, yOptions.Units,
                PptxChartMetricRules.AxisNiceVerticalValueTickTargetCount, yOptions.TickPolicy, null, reserveEndpoints: true, context: context).ConfigureAwait(false);
            sideWidth = captions.SlotWidth;
            if (outer)
            {
                if (await measurer.TryMeasureNormalLineHeightAsync(style).ConfigureAwait(false) is not { } line) return layout;
                yLine = line;
                sideWidth = captions.MaximumWidth + PptxChartMetricRules.OuterPlotCaptionLineClearanceFactor * line +
                    PptxChartMetricRules.OuterPlotCaptionPadding;
            }
        }
        if (outer)
        {
            // The authored outer area contains caption fragments. Use the same
            // interval constraints for both dimensions, including high labels.
            ChartPlotBox area = layout.PlotBox;
            double inset = PptxChartMetricRules.ChartContentFrameInset;
            var horizontal = AdmitChartCaptionInterval(area.X, area.Width, area.X, area.Width,
                firstWidth > 0d ? firstWidth / 2d + inset : 0d, lastWidth > 0d ? lastWidth / 2d + inset : 0d, sideWidth, yPlacement);
            var vertical = AdmitChartCaptionInterval(area.Y, area.Height, area.Y, area.Height, yLine / 2d, yLine / 2d,
                xLine > 0d ? 2d * PptxChartMetricRules.OuterPlotCaptionLineClearanceFactor * xLine + PptxChartMetricRules.OuterPlotCaptionPadding : 0d, xPlacement);
            return layout with { PlotBox = new(horizontal.Start, vertical.Start, horizontal.Length, vertical.Length) };
        }
        ChartPlotBox admitted = AdmitChartPlotCaptions(layout.Frame, layout.PlotBox, firstWidth, lastWidth, sideWidth, yPlacement);
        return layout with
        {
            PlotBox = admitted,
            PlotAreaBox = layout.PlotAreaBox with { X = admitted.X, Width = admitted.Width }
        };
    }

    private static IReadOnlyList<bool> ReadSceneSeriesLineHidden(PptxSceneChartPlot plot) =>
        plot.Series.Select(series => series.Line.NoFill).ToArray();

    // Shared scatter/bubble major ticks (Office honors out/in/cross at both value axes
    // with the same font-relative length as bars and lines).
    private static void RenderScatterBubbleMajorTicks(PdfGraphicsBuilder graphics, PptxTheme theme, XDocument chartXml, PptxSceneChart sceneChart, ChartPlotBox plotBox, ChartAxisSource xValueAxis, ChartAxisSource yValueAxis, ChartValueExtents xExtents, ChartValueExtents yExtents, ChartValueAxisRenderOptions xAxisOptions, ChartValueAxisRenderOptions yAxisOptions, ChartAxisPlacement xPlacement, ChartAxisPlacement yPlacement, PptxRenderContext? context = null)
    {
                PptxSceneChartAxisTickMark xTickMark = ReadSceneChartAxisMajorTickMark(xValueAxis.SceneAxis);
                if (xTickMark != PptxSceneChartAxisTickMark.None)
                {
                    ChartSeriesStroke xTickStroke = xAxisOptions.GridlineStyle.Major ?? DefaultChartGridlineStroke(major: true);
                    if (xTickStroke.Alpha > 0.001d)
                    {
                        ChartTextStyle xTickStyle = ReadSceneChartTextStyle(theme, sceneChart, xValueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                        IReadOnlyList<double> xTickValues = GetChartAxisTickValuesCore(xExtents, xAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: true, valueAxisLabelsVisible: IsSceneOrXmlChartAxisLabelVisible(xValueAxis.SceneAxis, xValueAxis.XmlAxis), manualPlotLayoutApplied: false), xAxisOptions.TickPolicy, context: context);
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

                PptxSceneChartAxisTickMark yTickMark = ReadSceneChartAxisMajorTickMark(yValueAxis.SceneAxis);
                if (yTickMark != PptxSceneChartAxisTickMark.None)
                {
                    ChartSeriesStroke yTickStroke = yAxisOptions.GridlineStyle.Major ?? DefaultChartGridlineStroke(major: true);
                    if (yTickStroke.Alpha > 0.001d)
                    {
                        ChartTextStyle yTickStyle = ReadSceneChartTextStyle(theme, sceneChart, yValueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                        IReadOnlyList<double> yTickValues = GetChartAxisTickValuesCore(yExtents, yAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: false, valueAxisLabelsVisible: IsSceneOrXmlChartAxisLabelVisible(yValueAxis.SceneAxis, yValueAxis.XmlAxis), manualPlotLayoutApplied: false), yAxisOptions.TickPolicy, context: context);
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
        PptxSceneChartPlot? bubblePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bubble, 0);
        if (bubblePlot is not null)
        {
            XElement bubbleChart = bubblePlot.Source;
            IReadOnlyList<ScatterSeries> bubbleSeries = ReadSceneOrXmlScatterSeries(bubblePlot, bubbleChart, readBubbleSize: true, workbook: workbook, plotVisibleOnly: plotVisibleOnly);
            if (bubbleSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(bubblePlot, bubbleChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(bubblePlot, bubbleChart, theme, colorMap, null);
                ChartLayout chartLayout = GetBubbleChartLayout(document, theme, bounds, chartXml, sceneChart, colorMap);
                RenderChartAreaStyle(graphics, document, bounds, sceneChart);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                IReadOnlyList<ChartAxisSource> valueAxes = ReadChartValueAxesForPlot(sceneChart, bubblePlot);
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
                        legendEntries, xValueAxis, yValueAxis, measured.DataX, measured.DataY, fontResolver, context: context).ConfigureAwait(false) is { } prepared)
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
                ChartAxisPlacement xPlacement = ReadChartAxisPlacement(xValueAxis.SceneAxis, yExtents, yAxisOptions.Reversed) ?? new(yExtents.Min, 0d, false, false);
                ChartAxisPlacement yPlacement = ReadChartAxisPlacement(yValueAxis.SceneAxis, xExtents, xAxisOptions.Reversed, horizontal: false) ?? new(xExtents.Min, 0d, false, false);
                xAxisOptions = xAxisOptions with { CrossingValue = Math.Clamp(yPlacement.CrossingValue, xExtents.Min, xExtents.Max) };
                yAxisOptions = yAxisOptions with { CrossingValue = Math.Clamp(xPlacement.CrossingValue, yExtents.Min, yExtents.Max) };
                DrawHorizontalChartGridlines(graphics, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, yExtents, yAxisOptions.Units, yAxisOptions.CrossingValue, yAxisOptions.Reversed, major: true, yAxisOptions.GridlineStyle.Major, hasMajorGridlineElement: true, yAxisOptions.TickPolicy, context: context);
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
                await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, xValueAxis.XmlAxis, xValueAxis.SceneAxis, xExtents, xAxisOptions.Units, xAxisOptions.Reversed, horizontalBars: true, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver, chartFonts: fonts, tickPolicy: xAxisOptions.TickPolicy, placement: xPlacement, context: context).ConfigureAwait(false);
                await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, yValueAxis.XmlAxis, yValueAxis.SceneAxis, yExtents, yAxisOptions.Units, yAxisOptions.Reversed, horizontalBars: false, rightSide: false, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver, chartFonts: fonts, tickPolicy: yAxisOptions.TickPolicy, placement: yPlacement, context: context).ConfigureAwait(false);
                await RenderDefaultChartAxisTitlesAsync(theme, graphics, chartLayout, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                RenderScatterBubbleMajorTicks(graphics, theme, chartXml, sceneChart, plotBox, xValueAxis, yValueAxis, xExtents, yExtents, xAxisOptions, yAxisOptions, xPlacement, yPlacement, context: context);
                await RenderChartLegendAsync(graphics, chartLayout.Frame, plotBox, legendEntries, chartLayout.Legend,
                    ReadSceneChartLegendTextStyle(theme, sceneChart), fontResolver,
                    ChartLegendPlacement.BubbleTitleRightLegend, chartFonts: fonts, preparedBox: preparedLegend).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }
}
