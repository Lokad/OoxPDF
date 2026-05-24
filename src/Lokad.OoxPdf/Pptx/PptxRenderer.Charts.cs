using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private const string ChartColorStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

    private static void RenderChartFrame(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        PptxSceneNode node,
        GroupTransform transform,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        RenderChartFrame(context, graphics, fonts, bounds, node.Chart, relationships);
    }

    private static void RenderChartFrame(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        ShapeBounds? bounds,
        PptxSceneChart? chart,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        RenderChartFrame(
            context,
            graphics,
            fonts,
            bounds,
            chart?.RelationshipId,
            chart?.TargetPartName,
            chart?.ChartXml,
            chart?.PaletteColors,
            chart,
            relationships);
    }

    private static void RenderChartFrame(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        ShapeBounds? bounds,
        string? relationshipId,
        string? targetPartName,
        XDocument? chartXml,
        IReadOnlyList<RgbColor>? chartPalette,
        PptxSceneChart? sceneChart,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        string? chartPartName = targetPartName;
        if (chartPartName is null &&
            relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship))
        {
            chartPartName = relationship.ResolvedTarget;
        }

        if (bounds is null || chartPartName is null)
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart frame could not be resolved and was ignored.", context.Slide.PartName, context.SlideNumber, "Ignored");
            return;
        }

        XDocument? resolvedChartXml = chartXml;
        IReadOnlyList<RgbColor>? resolvedChartPalette = chartPalette;
        OoxPart? chartPart = null;
        if (resolvedChartXml is null)
        {
            chartPart = context.Package.GetPart(chartPartName);
            if (chartPart is null)
            {
                EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart part was missing and was ignored.", chartPartName, context.SlideNumber, "Ignored");
                return;
            }

            using Stream chartStream = chartPart.OpenRead();
            resolvedChartXml = SafeXml.Load(chartStream);
            resolvedChartPalette = ReadChartPaletteColors(context.Package, chartPart, context.Theme);
        }

        if (TryRenderChart(graphics, context.Document, context.Theme, resolvedChartPalette, bounds.Value, resolvedChartXml, sceneChart, fonts))
        {
            fonts.AddRange(RenderChartTitle(context.Document, context.Theme, graphics, bounds.Value, resolvedChartXml, sceneChart));
            return;
        }

        EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Only bar, line, area, scatter, bubble, radar, pie, and doughnut charts with cached numeric values are currently supported by the native chart renderer.", chartPartName, context.SlideNumber, "Ignored");
    }

    private static PptxSceneChartPlot? ReadSceneChartPlot(PptxSceneChart? chart, string kind, int index = 0)
    {
        return chart?
            .Plots
            .Where(plot => string.Equals(plot.Kind, kind, StringComparison.Ordinal))
            .Skip(index)
            .FirstOrDefault();
    }

    private static PptxSceneChartAxis? ReadSceneChartAxis(PptxSceneChart? chart, PptxSceneChartPlot? plot, string kind)
    {
        if (chart is null)
        {
            return null;
        }

        if (plot is not null && plot.AxisIds.Count != 0)
        {
            foreach (string axisId in plot.AxisIds)
            {
                PptxSceneChartAxis? axis = chart.Axes.FirstOrDefault(candidate =>
                    string.Equals(candidate.Id, axisId, StringComparison.Ordinal) &&
                    string.Equals(candidate.Kind, kind, StringComparison.Ordinal));
                if (axis is not null)
                {
                    return axis;
                }
            }
        }

        return chart.Axes.FirstOrDefault(axis => string.Equals(axis.Kind, kind, StringComparison.Ordinal));
    }

    private static string ReadSceneOrXmlChartValue(string? sceneValue, XElement element, string childName, string defaultValue = "")
    {
        return !string.IsNullOrEmpty(sceneValue)
            ? sceneValue
            : (string?)element.Element(ChartNamespace + childName)?.Attribute("val") ?? defaultValue;
    }

    private static double ReadSceneDoughnutHoleSize(PptxSceneChartPlot? plot, XElement doughnutChart)
    {
        return plot?.HoleSize is { } rawHoleSize
            ? Math.Clamp(rawHoleSize / 100d, PptxChartMetricRules.DoughnutHoleMinimumRatio, PptxChartMetricRules.DoughnutHoleMaximumRatio)
            : ReadDoughnutHoleSize(doughnutChart);
    }

    private static IReadOnlyList<IReadOnlyList<double>> ReadSceneOrXmlChartSeries(PptxSceneChartPlot? plot, XElement chartElement)
    {
        IReadOnlyList<double>[]? sceneSeries = plot?
            .Series
            .Select(series => series.Values)
            .Where(values => values.Count != 0)
            .ToArray();
        return sceneSeries is { Length: > 0 }
            ? sceneSeries
            : ReadChartSeries(chartElement);
    }

    private static IReadOnlyList<ScatterSeries> ReadSceneOrXmlScatterSeries(PptxSceneChartPlot? plot, XElement chartElement, bool readBubbleSize)
    {
        if (plot is null)
        {
            return ReadScatterSeries(chartElement, readBubbleSize);
        }

        var series = new List<ScatterSeries>();
        foreach (PptxSceneChartSeries item in plot.Series)
        {
            int count = Math.Min(item.XValues.Count, item.YValues.Count);
            if (count == 0)
            {
                continue;
            }

            var points = new ScatterPoint[count];
            for (int i = 0; i < count; i++)
            {
                double size = readBubbleSize && i < item.BubbleSizes.Count ? item.BubbleSizes[i] : 1d;
                points[i] = new ScatterPoint(item.XValues[i], item.YValues[i], size);
            }

            series.Add(new ScatterSeries(points));
        }

        return series.Count != 0 ? series : ReadScatterSeries(chartElement, readBubbleSize);
    }

    private static IReadOnlyList<string> ReadSceneOrXmlCategoryLabels(PptxSceneChartPlot? plot, XElement chartElement)
    {
        IReadOnlyList<string>? categories = plot?
            .Series
            .Select(series => series.Categories)
            .FirstOrDefault(values => values.Count != 0);
        return categories is { Count: > 0 }
            ? categories
            : ReadChartCategoryLabels(chartElement);
    }

    private static IReadOnlyList<ChartSeriesFill?> ReadSceneOrXmlSeriesFills(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme)
    {
        ChartSeriesFill?[]? fills = plot?
            .Series
            .Select(series => ToChartSeriesFill(series.Fill, series.PatternFill))
            .ToArray();
        return fills is { Length: > 0 } && fills.Any(fill => fill is not null)
            ? fills
            : ReadChartSeriesFills(chartElement, theme);
    }

    private static IReadOnlyList<ChartSeriesStroke?> ReadSceneOrXmlSeriesStrokes(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme)
    {
        ChartSeriesStroke?[]? strokes = plot?
            .Series
            .Select(series => ToChartSeriesStroke(series.Line))
            .ToArray();
        return strokes is { Length: > 0 } && strokes.Any(stroke => stroke is not null)
            ? strokes
            : ReadChartSeriesStrokes(chartElement, theme);
    }

    private static IReadOnlyList<ChartMarkerStyle> ReadSceneOrXmlMarkerStyles(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme)
    {
        ChartMarkerStyle[]? markers = plot?
            .Series
            .Select(series => new ChartMarkerStyle(
                series.Marker.Symbol,
                series.Marker.Size,
                series.Marker.Fill.HasFill ? new ChartSeriesFill(series.Marker.Fill.Color, series.Marker.Fill.Alpha) : null,
                ToChartSeriesStroke(series.Marker.Line)))
            .ToArray();
        return markers is { Length: > 0 }
            ? markers
            : ReadChartMarkerStyles(chartElement, theme);
    }

    private static IReadOnlyList<bool> ReadSceneOrXmlSmoothSeries(PptxSceneChartPlot? plot, XElement chartElement)
    {
        bool[]? smooth = plot?.Series.Select(series => series.Smooth).ToArray();
        return smooth is { Length: > 0 } && smooth.Any(value => value)
            ? smooth
            : ReadChartSeriesSmooth(chartElement);
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> ReadSceneOrXmlSeriesPointFills(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme)
    {
        IReadOnlyDictionary<int, ChartSeriesFill>[]? fills = plot?
            .Series
            .Select(ReadSceneChartPointFills)
            .ToArray();
        return fills is { Length: > 0 } && fills.Any(fill => fill.Count != 0)
            ? fills
            : ReadChartSeriesPointFills(chartElement, theme);
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> ReadSceneOrXmlSeriesPointStrokes(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme)
    {
        IReadOnlyDictionary<int, ChartSeriesStroke>[]? strokes = plot?
            .Series
            .Select(ReadSceneChartPointStrokes)
            .ToArray();
        return strokes is { Length: > 0 } && strokes.Any(stroke => stroke.Count != 0)
            ? strokes
            : ReadChartSeriesPointStrokes(chartElement, theme);
    }

    private static IReadOnlyDictionary<int, ChartSeriesFill> ReadSceneOrXmlChartPointFills(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme)
    {
        IReadOnlyDictionary<int, ChartSeriesFill>? fills = plot?.Series.Count > 0
            ? ReadSceneChartPointFills(plot.Series[0])
            : null;
        return fills is { Count: > 0 }
            ? fills
            : ReadChartPointFills(chartElement, theme);
    }

    private static IReadOnlyDictionary<int, ChartSeriesStroke> ReadSceneOrXmlChartPointStrokes(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme)
    {
        IReadOnlyDictionary<int, ChartSeriesStroke>? strokes = plot?.Series.Count > 0
            ? ReadSceneChartPointStrokes(plot.Series[0])
            : null;
        return strokes is { Count: > 0 }
            ? strokes
            : ReadChartPointStrokes(chartElement, theme);
    }

    private static IReadOnlyDictionary<int, double> ReadSceneOrXmlChartPointExplosions(PptxSceneChartPlot? plot, XElement chartElement)
    {
        IReadOnlyDictionary<int, double>? explosions = plot?.Series.Count > 0
            ? ReadSceneChartPointExplosions(plot.Series[0])
            : null;
        return explosions is { Count: > 0 }
            ? explosions
            : ReadChartPointExplosions(chartElement);
    }

    private static IReadOnlyDictionary<int, ChartSeriesFill> ReadSceneChartPointFills(PptxSceneChartSeries series)
    {
        var fills = new Dictionary<int, ChartSeriesFill>();
        foreach (PptxSceneChartPointStyle point in series.PointStyles)
        {
            if (ToChartSeriesFill(point.Fill, point.PatternFill) is { } fill)
            {
                fills[point.Index] = fill;
            }
        }

        return fills;
    }

    private static IReadOnlyDictionary<int, ChartSeriesStroke> ReadSceneChartPointStrokes(PptxSceneChartSeries series)
    {
        var strokes = new Dictionary<int, ChartSeriesStroke>();
        foreach (PptxSceneChartPointStyle point in series.PointStyles)
        {
            if (point.Line.HasLine)
            {
                strokes[point.Index] = ToChartSeriesStroke(point.Line) ?? default;
            }
        }

        return strokes;
    }

    private static IReadOnlyDictionary<int, double> ReadSceneChartPointExplosions(PptxSceneChartSeries series)
    {
        var explosions = new Dictionary<int, double>();
        foreach (PptxSceneChartPointStyle point in series.PointStyles)
        {
            if (point.Explosion is { } explosion)
            {
                explosions[point.Index] = Math.Clamp(explosion / 100d, 0d, 1d);
            }
        }

        return explosions;
    }

    private static ChartSeriesFill? ToChartSeriesFill(PptxSceneFillStyle fill, PptxScenePatternFill patternFill)
    {
        if (fill.HasFill)
        {
            return new ChartSeriesFill(fill.Color, fill.Alpha);
        }

        return patternFill.HasPattern
            ? new ChartSeriesFill(patternFill.Foreground, patternFill.Alpha, patternFill.Preset, patternFill.Background)
            : null;
    }

    private static ChartSeriesStroke? ToChartSeriesStroke(PptxSceneLineStyle line)
    {
        return line.HasLine
            ? new ChartSeriesStroke(line.Color, line.Alpha, line.Width, line.DashPattern, line.Cap, line.Join)
            : null;
    }

    private static ChartShapeStyle ToChartShapeStyle(PptxSceneChartShapeStyle style)
    {
        return new ChartShapeStyle(style.NoFill ? null : ToChartSeriesFill(style.Fill, style.PatternFill), ToChartSeriesStroke(style.Line));
    }

    private static bool TryRenderChart(PdfGraphicsBuilder graphics, PptxDocument document, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, List<PdfFontResource> fonts)
    {
        IReadOnlyList<XElement> barCharts = chartXml.Descendants(ChartNamespace + "barChart").ToArray();
        XElement? barChart = barCharts.FirstOrDefault();
        if (barChart is not null)
        {
            PptxSceneChartPlot? barPlot = ReadSceneChartPlot(sceneChart, "barChart");
            IReadOnlyList<IReadOnlyList<double>> barSeries = ReadSceneOrXmlChartSeries(barPlot, barChart);
            if (barSeries.Count != 0)
            {
                bool horizontalBars = string.Equals(ReadSceneOrXmlChartValue(barPlot?.BarDirection, barChart, "barDir"), "bar", StringComparison.Ordinal);
                string grouping = ReadSceneOrXmlChartValue(barPlot?.Grouping, barChart, "grouping", "clustered");
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(barPlot, barChart, theme);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, barPlot, chartXml, theme, barChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme);
                XElement? valueAxis = ReadChartValueAxisForChart(chartXml, barChart);
                PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, barPlot, "valAx");
                ChartGridlineStyle gridlineStyle = ReadSceneOrXmlChartGridlineStyle(valueSceneAxis, valueAxis, theme);
                ChartValueExtents valueExtents = ReadSceneOrXmlChartValueAxisExtents(valueSceneAxis, valueAxis, GetBarChartValueExtents(barSeries, grouping));
                ChartAxisUnits axisUnits = ReadSceneOrXmlChartValueAxisUnits(valueSceneAxis, valueAxis);
                bool varyColors = barPlot?.VaryColors ?? ReadChartVaryColors(barChart);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills = ReadSceneOrXmlSeriesPointFills(barPlot, barChart, theme);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes = ReadSceneOrXmlSeriesPointStrokes(barPlot, barChart, theme);
                var legendEntries = new List<ChartLegendEntry>(BuildFillLegendEntries(theme, chartPalette, barPlot, barChart, seriesFills));
                ChartLayout chartLayout = GetBarChartLayout(document, bounds, chartXml, sceneChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                RenderBarChart(graphics, theme, chartPalette, plotBox, barSeries, horizontalBars, grouping, seriesFills, pointFills, pointStrokes, ReadSceneOrXmlMajorGridlines(valueSceneAxis, chartXml), ReadSceneOrXmlMinorGridlines(valueSceneAxis, chartXml), gridlineStyle, axesStyle, plotAreaStyle, valueExtents, axisUnits, varyColors, barPlot?.GapWidth ?? ReadChartGapWidth(barChart), barPlot?.Overlap ?? ReadChartOverlap(barChart));
                XElement? secondaryValueAxis = null;
                PptxSceneChartAxis? secondaryValueSceneAxis = null;
                ChartValueExtents secondaryValueExtents = default;
                ChartAxisUnits secondaryAxisUnits = default;
                int seriesOffset = barSeries.Count;
                int barChartIndex = 1;
                foreach (XElement extraBarChart in barCharts.Skip(1))
                {
                    PptxSceneChartPlot? extraBarPlot = ReadSceneChartPlot(sceneChart, "barChart", barChartIndex);
                    IReadOnlyList<IReadOnlyList<double>> extraSeries = ReadSceneOrXmlChartSeries(extraBarPlot, extraBarChart);
                    if (extraSeries.Count == 0)
                    {
                        barChartIndex++;
                        continue;
                    }

                    string extraGrouping = ReadSceneOrXmlChartValue(extraBarPlot?.Grouping, extraBarChart, "grouping", "clustered");
                    bool extraHorizontalBars = string.Equals(ReadSceneOrXmlChartValue(extraBarPlot?.BarDirection, extraBarChart, "barDir"), "bar", StringComparison.Ordinal);
                    XElement? extraValueAxis = ReadChartValueAxisForChart(chartXml, extraBarChart);
                    PptxSceneChartAxis? extraValueSceneAxis = ReadSceneChartAxis(sceneChart, extraBarPlot, "valAx");
                    ChartValueExtents extraValueExtents = ReadSceneOrXmlChartValueAxisExtents(extraValueSceneAxis, extraValueAxis, GetBarChartValueExtents(extraSeries, extraGrouping));
                    ChartAxisUnits extraAxisUnits = ReadSceneOrXmlChartValueAxisUnits(extraValueSceneAxis, extraValueAxis);
                    IReadOnlyList<ChartSeriesFill?> extraSeriesFills = ReadSceneOrXmlSeriesFills(extraBarPlot, extraBarChart, theme);
                    IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> extraPointFills = ReadSceneOrXmlSeriesPointFills(extraBarPlot, extraBarChart, theme);
                    IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> extraPointStrokes = ReadSceneOrXmlSeriesPointStrokes(extraBarPlot, extraBarChart, theme);
                    if (!extraHorizontalBars && secondaryValueAxis is null && IsSceneOrXmlRightValueAxis(extraValueSceneAxis, extraValueAxis))
                    {
                        secondaryValueAxis = extraValueAxis;
                        secondaryValueSceneAxis = extraValueSceneAxis;
                        secondaryValueExtents = extraValueExtents;
                        secondaryAxisUnits = extraAxisUnits;
                    }

                    legendEntries.AddRange(BuildFillLegendEntries(theme, chartPalette, extraBarPlot, extraBarChart, extraSeriesFills, seriesOffset));
                    RenderBarChart(
                        graphics,
                        theme,
                        chartPalette,
                        plotBox,
                        extraSeries,
                        extraHorizontalBars,
                        extraGrouping,
                        extraSeriesFills,
                        extraPointFills,
                        extraPointStrokes,
                        majorGridlines: false,
                        minorGridlines: false,
                        ChartGridlineStyle.Empty,
                        axesStyle with { ValueAxisVisible = false, CategoryAxisVisible = false },
                        ChartShapeStyle.Empty,
                        extraValueExtents,
                        extraAxisUnits,
                        extraBarPlot?.VaryColors ?? ReadChartVaryColors(extraBarChart),
                        extraBarPlot?.GapWidth ?? ReadChartGapWidth(extraBarChart),
                        extraBarPlot?.Overlap ?? ReadChartOverlap(extraBarChart));
                    fonts.AddRange(RenderBarDataLabels(
                        theme,
                        graphics,
                        plotBox,
                        extraSeries,
                        extraValueExtents,
                        extraHorizontalBars,
                        ReadSceneOrXmlDataLabelOptions(extraBarPlot, extraBarChart, theme),
                        ReadSceneOrXmlSeriesDataLabelOptions(extraBarPlot, extraBarChart, theme),
                        ReadSceneOrXmlCategoryLabels(extraBarPlot, extraBarChart),
                        ReadSceneOrXmlChartSeriesNames(extraBarPlot, extraBarChart)));
                    seriesOffset += extraSeries.Count;
                    barChartIndex++;
                }

                XElement? categoryAxis = ReadChartCategoryAxisForChart(chartXml, barChart);
                PptxSceneChartAxis? categorySceneAxis = ReadSceneChartAxis(sceneChart, barPlot, "catAx");
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categorySceneAxis, categoryAxis))
                {
                    fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categorySceneAxis, categoryAxis, ReadSceneOrXmlCategoryLabels(barPlot, barChart), horizontalBars));
                }

                if (axesStyle.ValueAxisVisible)
                {
                    bool sameSideSecondaryValueAxis = !horizontalBars &&
                        secondaryValueAxis is not null &&
                        IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis) &&
                        GetValueAxisSideSlot(valueSceneAxis, valueAxis, secondaryValueSceneAxis, secondaryValueAxis, defaultPrimaryRightSide: false, defaultSecondaryRightSide: true) > 0;
                    if (IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis))
                    {
                        fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis, valueSceneAxis, valueExtents, axisUnits, horizontalBars, useTextSizedWidth: sameSideSecondaryValueAxis));
                    }

                    if (!horizontalBars)
                    {
                        if (secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                        {
                            int sideSlot = GetValueAxisSideSlot(valueSceneAxis, valueAxis, secondaryValueSceneAxis, secondaryValueAxis, defaultPrimaryRightSide: false, defaultSecondaryRightSide: true);
                            fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, horizontalBars: false, rightSide: true, axisSideSlot: sideSlot, useTextSizedWidth: sideSlot > 0));
                        }
                        else
                        {
                            fonts.AddRange(RenderSecondaryChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, GetBarChartValueExtents(barSeries, grouping)));
                        }
                    }
                }
                else if (!horizontalBars && secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                {
                    fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, horizontalBars: false, rightSide: true));
                }
                fonts.AddRange(RenderChartLegend(graphics, plotBox, legendEntries, chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, sceneChart, chartXml)));
                fonts.AddRange(RenderBarDataLabels(
                    theme,
                    graphics,
                    plotBox,
                    barSeries,
                    valueExtents,
                    horizontalBars,
                    ReadSceneOrXmlDataLabelOptions(barPlot, barChart, theme),
                    ReadSceneOrXmlSeriesDataLabelOptions(barPlot, barChart, theme),
                    ReadSceneOrXmlCategoryLabels(barPlot, barChart),
                    ReadSceneOrXmlChartSeriesNames(barPlot, barChart)));
                return true;
            }
        }

        XElement? lineChart = chartXml.Descendants(ChartNamespace + "lineChart").FirstOrDefault();
        if (lineChart is not null)
        {
            PptxSceneChartPlot? linePlot = ReadSceneChartPlot(sceneChart, "lineChart");
            IReadOnlyList<IReadOnlyList<double>> lineSeries = ReadSceneOrXmlChartSeries(linePlot, lineChart);
            if (lineSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(linePlot, lineChart, theme);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(linePlot, lineChart, theme);
                IReadOnlyList<bool> smoothSeries = ReadSceneOrXmlSmoothSeries(linePlot, lineChart);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, linePlot, chartXml, theme, lineChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme);
                XElement? valueAxis = ReadChartValueAxisForChart(chartXml, lineChart);
                XElement? valueAxisForScale = valueAxis ?? chartXml.Descendants(ChartNamespace + "valAx").FirstOrDefault();
                PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, linePlot, "valAx");
                ChartGridlineStyle gridlineStyle = ReadSceneOrXmlChartGridlineStyle(valueSceneAxis, valueAxisForScale, theme);
                ChartValueExtents valueExtents = ReadSceneOrXmlChartValueAxisExtents(valueSceneAxis, valueAxisForScale, GetLineChartValueExtents(lineSeries));
                ChartAxisUnits axisUnits = ReadSceneOrXmlChartValueAxisUnits(valueSceneAxis, valueAxisForScale);
                ChartLayout chartLayout = GetLineChartLayout(document, bounds, chartXml, sceneChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                RenderLineChart(graphics, plotBox, lineSeries, seriesStrokes, markerStyles, smoothSeries, ReadSceneOrXmlMajorGridlines(valueSceneAxis, chartXml), ReadSceneOrXmlMinorGridlines(valueSceneAxis, chartXml), gridlineStyle, axesStyle, plotAreaStyle, valueExtents, axisUnits);
                XElement? categoryAxis = ReadChartCategoryAxisForChart(chartXml, lineChart);
                PptxSceneChartAxis? categorySceneAxis = ReadSceneChartAxis(sceneChart, linePlot, "catAx");
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categorySceneAxis, categoryAxis))
                {
                    fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categorySceneAxis, categoryAxis, ReadSceneOrXmlCategoryLabels(linePlot, lineChart), horizontalBars: false));
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis))
                {
                    fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis, valueSceneAxis, valueExtents, axisUnits, horizontalBars: false));
                    fonts.AddRange(RenderSecondaryChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, GetLineChartValueExtents(lineSeries)));
                }
                fonts.AddRange(RenderChartLegend(graphics, plotBox, BuildStrokeLegendEntries(linePlot, lineChart, seriesStrokes), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, sceneChart, chartXml)));
                fonts.AddRange(RenderLineDataLabels(
                    theme,
                    graphics,
                    plotBox,
                    lineSeries,
                    valueExtents,
                    ReadSceneOrXmlDataLabelOptions(linePlot, lineChart, theme),
                    ReadSceneOrXmlSeriesDataLabelOptions(linePlot, lineChart, theme),
                    ReadSceneOrXmlCategoryLabels(linePlot, lineChart),
                    ReadSceneOrXmlChartSeriesNames(linePlot, lineChart)));
                return true;
            }
        }

        XElement? areaChart = chartXml.Descendants(ChartNamespace + "areaChart").FirstOrDefault();
        if (areaChart is not null)
        {
            PptxSceneChartPlot? areaPlot = ReadSceneChartPlot(sceneChart, "areaChart");
            IReadOnlyList<IReadOnlyList<double>> areaSeries = ReadSceneOrXmlChartSeries(areaPlot, areaChart);
            if (areaSeries.Count != 0)
            {
                string grouping = (string?)areaChart.Element(ChartNamespace + "grouping")?.Attribute("val") ?? "standard";
                grouping = string.IsNullOrEmpty(areaPlot?.Grouping) ? grouping : areaPlot.Grouping;
                bool stacked = string.Equals(grouping, "stacked", StringComparison.Ordinal) ||
                    string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(areaPlot, areaChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(areaPlot, areaChart, theme);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                RenderAreaChart(graphics, document, bounds, areaSeries, stacked, seriesFills, seriesStrokes);
                return true;
            }
        }

        XElement? scatterChart = chartXml.Descendants(ChartNamespace + "scatterChart").FirstOrDefault();
        if (scatterChart is not null)
        {
            PptxSceneChartPlot? scatterPlot = ReadSceneChartPlot(sceneChart, "scatterChart");
            IReadOnlyList<ScatterSeries> scatterSeries = ReadSceneOrXmlScatterSeries(scatterPlot, scatterChart, readBubbleSize: false);
            if (scatterSeries.Count != 0)
            {
                string scatterStyle = ReadSceneOrXmlChartValue(scatterPlot?.ScatterStyle, scatterChart, "scatterStyle");
                bool connectLines = scatterStyle.Contains("Line", StringComparison.OrdinalIgnoreCase);
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(scatterPlot, scatterChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(scatterPlot, scatterChart, theme);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(scatterPlot, scatterChart, theme);
                IReadOnlyList<bool> smoothSeries = ReadSceneOrXmlSmoothSeries(scatterPlot, scatterChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                RenderScatterChart(graphics, document, bounds, scatterSeries, connectLines, bubble: false, seriesFills, seriesStrokes, markerStyles, smoothSeries);
                return true;
            }
        }

        XElement? bubbleChart = chartXml.Descendants(ChartNamespace + "bubbleChart").FirstOrDefault();
        if (bubbleChart is not null)
        {
            PptxSceneChartPlot? bubblePlot = ReadSceneChartPlot(sceneChart, "bubbleChart");
            IReadOnlyList<ScatterSeries> bubbleSeries = ReadSceneOrXmlScatterSeries(bubblePlot, bubbleChart, readBubbleSize: true);
            if (bubbleSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(bubblePlot, bubbleChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(bubblePlot, bubbleChart, theme);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                RenderScatterChart(graphics, document, bounds, bubbleSeries, connectLines: false, bubble: true, seriesFills, seriesStrokes, [], []);
                return true;
            }
        }

        XElement? radarChart = chartXml.Descendants(ChartNamespace + "radarChart").FirstOrDefault();
        if (radarChart is not null)
        {
            PptxSceneChartPlot? radarPlot = ReadSceneChartPlot(sceneChart, "radarChart");
            IReadOnlyList<IReadOnlyList<double>> radarSeries = ReadSceneOrXmlChartSeries(radarPlot, radarChart);
            if (radarSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(radarPlot, radarChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(radarPlot, radarChart, theme);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                RenderRadarChart(graphics, document, bounds, radarSeries, seriesFills, seriesStrokes);
                return true;
            }
        }

        XElement? pieChart = chartXml.Descendants(ChartNamespace + "pieChart").FirstOrDefault();
        if (pieChart is not null)
        {
            PptxSceneChartPlot? piePlot = ReadSceneChartPlot(sceneChart, "pieChart");
            IReadOnlyList<IReadOnlyList<double>> pieSeries = ReadSceneOrXmlChartSeries(piePlot, pieChart);
            if (pieSeries.Count != 0)
            {
                IReadOnlyDictionary<int, ChartSeriesFill> pointFills = ReadSceneOrXmlChartPointFills(piePlot, pieChart, theme);
                IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes = ReadSceneOrXmlChartPointStrokes(piePlot, pieChart, theme);
                IReadOnlyDictionary<int, double> pointExplosions = ReadSceneOrXmlChartPointExplosions(piePlot, pieChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                RenderPieChart(graphics, document, bounds, pieSeries[0], pointFills, pointStrokes, pointExplosions);
                fonts.AddRange(RenderPieDataLabels(document, theme, graphics, bounds, pieSeries[0], pointExplosions, 0d, ReadSceneOrXmlDataLabelOptions(piePlot, pieChart, theme)));
                return true;
            }
        }

        XElement? doughnutChart = chartXml.Descendants(ChartNamespace + "doughnutChart").FirstOrDefault();
        if (doughnutChart is not null)
        {
            PptxSceneChartPlot? doughnutPlot = ReadSceneChartPlot(sceneChart, "doughnutChart");
            IReadOnlyList<IReadOnlyList<double>> doughnutSeries = ReadSceneOrXmlChartSeries(doughnutPlot, doughnutChart);
            if (doughnutSeries.Count != 0)
            {
                IReadOnlyDictionary<int, ChartSeriesFill> pointFills = ReadSceneOrXmlChartPointFills(doughnutPlot, doughnutChart, theme);
                IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes = ReadSceneOrXmlChartPointStrokes(doughnutPlot, doughnutChart, theme);
                IReadOnlyDictionary<int, double> pointExplosions = ReadSceneOrXmlChartPointExplosions(doughnutPlot, doughnutChart);
                double holeSize = ReadSceneDoughnutHoleSize(doughnutPlot, doughnutChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                RenderDoughnutChart(graphics, document, bounds, doughnutSeries[0], pointFills, pointStrokes, pointExplosions, holeSize);
                fonts.AddRange(RenderPieDataLabels(document, theme, graphics, bounds, doughnutSeries[0], pointExplosions, holeSize, ReadSceneOrXmlDataLabelOptions(doughnutPlot, doughnutChart, theme)));
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<RgbColor>? ReadChartPaletteColors(OoxPackage package, OoxPart chartPart, PptxTheme theme)
    {
        OoxRelationship? colorRelationship = package.GetRelationships(chartPart.Name)
            .FirstOrDefault(relationship => !relationship.IsExternal &&
                relationship.Type == ChartColorStyleRelationshipType &&
                relationship.ResolvedTarget is not null);
        if (colorRelationship?.ResolvedTarget is null)
        {
            return null;
        }

        OoxPart? colorPart = package.GetPart(colorRelationship.ResolvedTarget);
        if (colorPart is null)
        {
            return null;
        }

        using Stream stream = colorPart.OpenRead();
        XDocument document = SafeXml.Load(stream);
        var colors = new List<RgbColor>();
        foreach (XElement colorElement in document.Root?.Elements().Where(element => element.Name.Namespace == DrawingNamespace) ?? [])
        {
            var wrapper = new XElement(DrawingNamespace + "solidFill", new XElement(colorElement));
            if (TryReadSolidColor(wrapper, theme, out RgbColor color))
            {
                colors.Add(color);
            }
        }

        return colors.Count == 0 ? null : colors;
    }

    private static IReadOnlyList<IReadOnlyList<double>> ReadChartSeries(XDocument chartXml, string chartElementName)
    {
        XElement? chartElement = chartXml.Descendants(ChartNamespace + chartElementName).FirstOrDefault();
        return chartElement is null ? [] : ReadChartSeries(chartElement);
    }

    private static IReadOnlyList<IReadOnlyList<double>> ReadChartSeries(XElement chartElement)
    {
        var series = new List<IReadOnlyList<double>>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            double[] values = element
                .Elements(ChartNamespace + "val")
                .Descendants(ChartNamespace + "pt")
                .Select(point => (string?)point.Element(ChartNamespace + "v"))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : double.NaN)
                .Where(value => !double.IsNaN(value))
                .ToArray();
            if (values.Length > 0)
            {
                series.Add(values);
            }
        }

        return series;
    }

    private static ChartFrameBox GetChartFrameBox(PptxDocument document, ShapeBounds bounds)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        return new ChartFrameBox(x, y, width, height);
    }

    private static void RenderChartAreaStyle(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, PptxTheme theme)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        RenderChartShapeStyle(graphics, frame.X, frame.Y, frame.Width, frame.Height, ReadSceneOrXmlChartAreaStyle(sceneChart, chartXml, theme));
    }

    private static ChartShapeStyle ReadSceneOrXmlChartAreaStyle(PptxSceneChart? sceneChart, XDocument chartXml, PptxTheme theme)
    {
        return sceneChart is null
            ? ReadChartShapeStyle(chartXml.Root?.Element(ChartNamespace + "spPr"), theme)
            : ToChartShapeStyle(sceneChart.ChartAreaStyle);
    }

    private static ChartShapeStyle ReadSceneOrXmlChartPlotAreaStyle(PptxSceneChart? sceneChart, XDocument chartXml, PptxTheme theme)
    {
        return sceneChart is null
            ? ReadChartPlotAreaStyle(chartXml, theme)
            : ToChartShapeStyle(sceneChart.PlotAreaStyle);
    }

    private static ChartShapeStyle ReadChartPlotAreaStyle(XDocument chartXml, PptxTheme theme)
    {
        XElement? shapeProperties = chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "spPr");
        return ReadChartShapeStyle(shapeProperties, theme);
    }

    private static bool TryReadSceneOrXmlManualPlotBox(PptxSceneChart? sceneChart, XDocument chartXml, ChartFrameBox frame, out ChartPlotBox plotBox)
    {
        if (sceneChart is not null)
        {
            return TryBuildManualPlotBox(sceneChart.PlotAreaLayout, frame, out plotBox);
        }

        return TryReadManualPlotBox(chartXml, frame, out plotBox);
    }

    private static bool TryReadManualPlotBox(XDocument chartXml, ChartFrameBox frame, out ChartPlotBox plotBox)
    {
        plotBox = default;
        XElement? manualLayout = chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "layout")
            ?.Element(ChartNamespace + "manualLayout");
        if (manualLayout is null)
        {
            return false;
        }

        double? x = ReadManualLayoutFactor(manualLayout, "x");
        double? y = ReadManualLayoutFactor(manualLayout, "y");
        double? width = ReadManualLayoutFactor(manualLayout, "w");
        double? height = ReadManualLayoutFactor(manualLayout, "h");
        if (x is null || y is null || width is null || height is null)
        {
            return false;
        }

        return TryBuildManualPlotBox(new PptxSceneChartManualLayout(
            true,
            x.Value,
            y.Value,
            width.Value,
            height.Value,
            ReadManualLayoutValue(manualLayout, "layoutTarget"),
            ReadManualLayoutValue(manualLayout, "xMode"),
            ReadManualLayoutValue(manualLayout, "yMode"),
            ReadManualLayoutValue(manualLayout, "wMode"),
            ReadManualLayoutValue(manualLayout, "hMode")), frame, out plotBox);
    }

    private static bool TryBuildManualPlotBox(PptxSceneChartManualLayout layout, ChartFrameBox frame, out ChartPlotBox plotBox)
    {
        plotBox = default;
        if (!layout.HasLayout)
        {
            return false;
        }

        double left = Math.Clamp(layout.X, 0d, 1d);
        double top = Math.Clamp(layout.Y, 0d, 1d);
        double width = Math.Clamp(layout.Width, 0.02d, 1d);
        double height = Math.Clamp(layout.Height, 0.02d, 1d);
        double right = IsManualLayoutEdgeMode(layout.WidthMode)
            ? Math.Clamp(layout.Width, left, 1d)
            : left + width;
        double bottom = IsManualLayoutEdgeMode(layout.HeightMode)
            ? Math.Clamp(layout.Height, top, 1d)
            : top + height;
        double plotWidth = Math.Max(0d, right - left) * frame.Width;
        double plotHeight = Math.Max(0d, bottom - top) * frame.Height;
        double plotX = frame.X + left * frame.Width;
        double plotY = frame.Y + frame.Height - bottom * frame.Height;
        plotBox = new ChartPlotBox(plotX, plotY, plotWidth, plotHeight);
        return plotWidth > 0d && plotHeight > 0d;
    }

    private static bool IsManualLayoutEdgeMode(string mode)
    {
        return string.Equals(mode, "edge", StringComparison.Ordinal);
    }

    private static double? ReadManualLayoutFactor(XElement manualLayout, string elementName)
    {
        string? value = (string?)manualLayout.Element(ChartNamespace + elementName)?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? parsed
            : null;
    }

    private static string ReadManualLayoutValue(XElement manualLayout, string elementName)
    {
        return (string?)manualLayout.Element(ChartNamespace + elementName)?.Attribute("val") ?? string.Empty;
    }

    private static ChartValueExtents ReadChartValueAxisExtents(XDocument chartXml, ChartValueExtents fallback)
    {
        return ReadChartValueAxisExtents(chartXml
            .Descendants(ChartNamespace + "valAx")
            .FirstOrDefault(), fallback);
    }

    private static XElement? ReadChartValueAxisForChart(XDocument chartXml, XElement chartElement)
    {
        HashSet<string> axisIds = chartElement
            .Elements(ChartNamespace + "axId")
            .Select(axis => (string?)axis.Attribute("val"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id!)
            .ToHashSet(StringComparer.Ordinal);
        return chartXml
            .Descendants(ChartNamespace + "valAx")
            .FirstOrDefault(axis => axisIds.Contains((string?)axis.Element(ChartNamespace + "axId")?.Attribute("val") ?? string.Empty));
    }

    private static XElement? ReadChartCategoryAxisForChart(XDocument chartXml, XElement chartElement)
    {
        HashSet<string> axisIds = chartElement
            .Elements(ChartNamespace + "axId")
            .Select(axis => (string?)axis.Attribute("val"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id!)
            .ToHashSet(StringComparer.Ordinal);
        return chartXml
            .Descendants(ChartNamespace + "catAx")
            .FirstOrDefault(axis => axisIds.Contains((string?)axis.Element(ChartNamespace + "axId")?.Attribute("val") ?? string.Empty)) ??
            chartXml.Descendants(ChartNamespace + "catAx").FirstOrDefault();
    }

    private static ChartValueExtents ReadChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback)
    {
        XElement? scaling = valueAxis?.Element(ChartNamespace + "scaling");
        if (scaling is null)
        {
            return fallback;
        }

        double min = ReadAxisScalingValue(scaling, "min") ?? fallback.Min;
        double max = ReadAxisScalingValue(scaling, "max") ?? GetNiceChartAxisMax(fallback.Max, fallback.Min);
        return max > min
            ? new ChartValueExtents(min, max)
            : fallback;
    }

    private static ChartValueExtents ReadSceneOrXmlChartValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents fallback)
    {
        if (axis is null)
        {
            return ReadChartValueAxisExtents(valueAxis, fallback);
        }

        if (!axis.HasScaling)
        {
            return fallback;
        }

        double min = axis.Minimum ?? fallback.Min;
        double max = axis.Maximum ?? GetNiceChartAxisMax(fallback.Max, fallback.Min);
        return max > min
            ? new ChartValueExtents(min, max)
            : fallback;
    }

    private static double? ReadAxisScalingValue(XElement scaling, string elementName)
    {
        string? value = (string?)scaling.Element(ChartNamespace + elementName)?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? parsed
            : null;
    }

    private static ChartAxisUnits ReadChartValueAxisUnits(XDocument chartXml)
    {
        return ReadChartValueAxisUnits(chartXml
            .Descendants(ChartNamespace + "valAx")
            .FirstOrDefault());
    }

    private static ChartAxisUnits ReadChartValueAxisUnits(XElement? valueAxis)
    {
        if (valueAxis is null)
        {
            return ChartAxisUnits.Empty;
        }

        return new ChartAxisUnits(
            ReadAxisUnitValue(valueAxis, "majorUnit"),
            ReadAxisUnitValue(valueAxis, "minorUnit"));
    }

    private static ChartAxisUnits ReadSceneOrXmlChartValueAxisUnits(PptxSceneChartAxis? axis, XElement? valueAxis)
    {
        return axis is null
            ? ReadChartValueAxisUnits(valueAxis)
            : new ChartAxisUnits(axis.MajorUnit, axis.MinorUnit);
    }

    private static double? ReadAxisUnitValue(XElement valueAxis, string elementName)
    {
        string? value = (string?)valueAxis.Element(ChartNamespace + elementName)?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) && parsed > 0d
            ? parsed
            : null;
    }

    private static ChartShapeStyle ReadChartShapeStyle(XElement? shapeProperties, PptxTheme theme)
    {
        if (shapeProperties is null)
        {
            return ChartShapeStyle.Empty;
        }

        ChartSeriesFill? fill = TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
            ? new ChartSeriesFill(fillColor, fillAlpha)
            : null;
        ChartSeriesStroke? stroke = TryReadLineWithAlpha(shapeProperties, theme, out RgbColor strokeColor, out double lineWidth, out double strokeAlpha)
            ? new ChartSeriesStroke(strokeColor, strokeAlpha, lineWidth)
            : null;
        return new ChartShapeStyle(fill, stroke);
    }

    private static void RenderChartShapeStyle(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartShapeStyle style)
    {
        if (style.Fill is { } fill)
        {
            FillChartRectangle(graphics, x, y, width, height, fill);
        }

        if (style.Stroke is { } stroke)
        {
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            graphics.SetStrokeRgb(stroke.Color.Red, stroke.Color.Green, stroke.Color.Blue);
            graphics.SetLineWidth(stroke.Width);
            if (stroke.DashPattern is { Count: > 0 })
            {
                graphics.SetLineDash(stroke.DashPattern);
            }

            if (stroke.Cap is { } cap)
            {
                graphics.SetLineCap(cap);
            }

            if (stroke.Join is { } join)
            {
                graphics.SetLineJoin(join);
            }

            graphics.StrokeRectangle(x, y, width, height);
            if (stroke.DashPattern is { Count: > 0 })
            {
                graphics.ClearLineDash();
            }

            if (stroke.Cap is not null)
            {
                graphics.SetLineCap(0);
            }

            if (stroke.Join is not null)
            {
                graphics.SetLineJoin(0);
            }
            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static IReadOnlyList<PdfFontResource> RenderChartTitle(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        if (string.IsNullOrWhiteSpace(title))
        {
            return [];
        }

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        ChartTextStyle style = ReadSceneOrXmlChartTitleTextStyle(theme, sceneChart, chartXml);
        double fontSize = style.FontSize;
        var run = new TextRun(
            title.Trim(),
            x + width * 0.08d,
            y + height * 0.88d,
            width * 0.84d,
            fontSize * 1.4d,
            x,
            y,
            width,
            height,
            fontSize,
            0d,
            0d,
            style.Color,
            1d,
            null,
            Bold: style.Bold,
            Italic: style.Italic,
            Underline: false,
            Strike: false,
            KerningEnabled: true,
            TextAlignment.Center,
            FontFamily: style.FontFamily,
            RotationDegrees: 0d,
            RotationCenterX: 0d,
            RotationCenterY: 0d,
            FlipHorizontal: false,
            FlipVertical: false);
        return RenderTextRuns([run], graphics, "CT");
    }

    private static ChartTextStyle ReadSceneOrXmlChartTitleTextStyle(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        XElement? title = chartXml.Descendants(ChartNamespace + "title").FirstOrDefault();
        if (sceneChart is null)
        {
            return ReadChartTextStyle(theme, chartXml, title, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize);
        style = MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.TextStyle));
        return MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.Title.TextStyle));
    }

    private static string? ReadSceneOrXmlChartTitleText(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            if (sceneChart.Title.IsAutoDeleted)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(sceneChart.Title.Text))
            {
                return sceneChart.Title.Text;
            }
        }

        return ReadChartTitleText(chartXml);
    }

    private static string? ReadChartTitleText(XDocument chartXml)
    {
        XElement? title = chartXml.Descendants(ChartNamespace + "title").FirstOrDefault();
        if (title is null)
        {
            bool autoTitleDeleted = IsOoxmlBooleanElementEnabled(chartXml.Descendants(ChartNamespace + "autoTitleDeleted").FirstOrDefault());
            if (autoTitleDeleted)
            {
                return null;
            }

            if (!chartXml.Descendants(ChartNamespace + "barChart").Any())
            {
                return null;
            }

            IReadOnlyList<XElement> series = chartXml
                .Descendants(ChartNamespace + "ser")
                .ToArray();
            if (series.Count != 1)
            {
                return null;
            }

            return ReadChartSeriesName(series[0]);
        }

        string text = string.Concat(title
            .Descendants(DrawingNamespace + "t")
            .Select(element => (string?)element)
            .Where(value => !string.IsNullOrEmpty(value)));
        if (!string.IsNullOrWhiteSpace(text))
        {
            return text;
        }

        return (string?)title
            .Descendants(ChartNamespace + "v")
            .FirstOrDefault();
    }

    private static IReadOnlyList<string> ReadChartCategoryLabels(XElement chartElement)
    {
        XElement? categories = chartElement
            .Element(ChartNamespace + "ser")
            ?.Element(ChartNamespace + "cat");
        if (categories is null)
        {
            return [];
        }

        return categories
            .Descendants(ChartNamespace + "pt")
            .Select(point => ((string?)point.Element(ChartNamespace + "v"))?.Trim())
            .Where(value => !string.IsNullOrEmpty(value))
            .Cast<string>()
            .ToArray();
    }

    private static IReadOnlyList<ChartLegendEntry> BuildFillLegendEntries(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesFill?> seriesFills, int paletteOffset = 0)
    {
        IReadOnlyList<string> names = ReadSceneOrXmlChartSeriesNames(plot, chartElement);
        var entries = new List<ChartLegendEntry>(names.Count);
        for (int i = 0; i < names.Count; i++)
        {
            ChartSeriesFill fill = i < seriesFills.Count && seriesFills[i] is { } explicitFill
                ? explicitFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, i + paletteOffset), 1d);
            entries.Add(new ChartLegendEntry(names[i], fill, null));
        }

        return entries;
    }

    private static IReadOnlyList<ChartLegendEntry> BuildStrokeLegendEntries(PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesStroke?> seriesStrokes)
    {
        IReadOnlyList<string> names = ReadSceneOrXmlChartSeriesNames(plot, chartElement);
        var entries = new List<ChartLegendEntry>(names.Count);
        for (int i = 0; i < names.Count; i++)
        {
            entries.Add(new ChartLegendEntry(names[i], null, ChartSeriesStrokeColor(i, seriesStrokes, 1.5d)));
        }

        return entries;
    }

    private static IReadOnlyList<string> ReadSceneOrXmlChartSeriesNames(PptxSceneChartPlot? plot, XElement chartElement)
    {
        return plot?.Series.Count > 0
            ? plot.Series
                .Select((series, index) => string.IsNullOrWhiteSpace(series.Name) ? $"Series {index + 1}" : series.Name.Trim())
                .ToArray()
            : ReadChartSeriesNames(chartElement);
    }

    private static IReadOnlyList<string> ReadChartSeriesNames(XElement chartElement)
    {
        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select((series, index) => ReadChartSeriesName(series) ?? $"Series {index + 1}")
            .ToArray();
    }

    private static string? ReadChartSeriesName(XElement series)
    {
        return ReadChartText(series.Element(ChartNamespace + "tx"));
    }

    private static string? ReadChartText(XElement? text)
    {
        string? literal = text?
            .Descendants(ChartNamespace + "v")
            .Select(value => value.Value.Trim())
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (!string.IsNullOrWhiteSpace(literal))
        {
            return literal;
        }

        string? richText = text?
            .Descendants(DrawingNamespace + "t")
            .Aggregate(string.Empty, (current, textElement) => current + textElement.Value);
        return string.IsNullOrWhiteSpace(richText) ? null : richText;
    }

    private static ChartLegendLayout ReadChartLegendLayout(XDocument chartXml)
    {
        XElement? legend = chartXml
            .Descendants(ChartNamespace + "legend")
            .FirstOrDefault();
        if (legend is null)
        {
            return ChartLegendLayout.Hidden;
        }

        if (IsOoxmlBooleanElementEnabled(legend.Element(ChartNamespace + "delete")))
        {
            return ChartLegendLayout.Hidden;
        }

        string position = (string?)legend.Element(ChartNamespace + "legendPos")?.Attribute("val") ?? "r";
        bool overlay = IsOoxmlBooleanElementEnabled(legend.Element(ChartNamespace + "overlay"));
        return new ChartLegendLayout(position, overlay, Visible: true);
    }

    private static ChartLegendLayout ReadSceneOrXmlChartLegendLayout(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is null)
        {
            return ReadChartLegendLayout(chartXml);
        }

        return sceneChart.Legend.IsVisible
            ? new ChartLegendLayout(sceneChart.Legend.Position, sceneChart.Legend.Overlay, Visible: true)
            : ChartLegendLayout.Hidden;
    }

    private static ChartTextStyle ReadSceneOrXmlChartLegendTextStyle(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        XElement? legend = chartXml.Descendants(ChartNamespace + "legend").FirstOrDefault();
        if (sceneChart is null)
        {
            return ReadChartTextStyle(theme, chartXml, legend, fallbackFontSize: PptxChartMetricRules.LegendFallbackFontSize);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize: PptxChartMetricRules.LegendFallbackFontSize);
        style = MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.TextStyle));
        return MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.Legend.TextStyle));
    }

    private static IReadOnlyList<PdfFontResource> RenderChartLegend(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, IReadOnlyList<ChartLegendEntry> entries, ChartLegendLayout layout, ChartTextStyle style)
    {
        if (!layout.Visible || entries.Count == 0)
        {
            return [];
        }

        double fontSize = style.FontSize;
        double lineHeight = fontSize * 1.45d;
        double markerSize = fontSize * 0.65d;
        bool horizontal = string.Equals(layout.Position, "b", StringComparison.Ordinal) ||
            string.Equals(layout.Position, "t", StringComparison.Ordinal);
        double width = horizontal ? plotBox.Width : Math.Max(36d, plotBox.Width * 0.22d);
        double x = layout.Position switch
        {
            "l" => Math.Max(0d, plotBox.X - width - 8d),
            _ when horizontal => plotBox.X,
            _ => plotBox.X + plotBox.Width + 8d
        };
        double firstY = layout.Position switch
        {
            "b" => Math.Max(0d, plotBox.Y - lineHeight * 1.15d),
            "t" => plotBox.Y + plotBox.Height + lineHeight * 0.15d,
            _ => plotBox.Y + plotBox.Height - lineHeight
        };
        double clipHeight = horizontal ? lineHeight * 1.25d : Math.Max(lineHeight, entries.Count * lineHeight);
        var runs = new List<TextRun>(entries.Count);
        for (int i = 0; i < entries.Count; i++)
        {
            ChartLegendEntry entry = entries[i];
            double entryX = horizontal ? x + i * (width / entries.Count) : x;
            double entryWidth = horizontal ? width / entries.Count : width;
            double y = horizontal ? firstY : firstY - i * lineHeight;
            double markerY = y + lineHeight * 0.35d;
            if (entry.Fill is { } fill)
            {
                FillChartRectangle(graphics, entryX, markerY, markerSize, markerSize, fill);
            }
            else if (entry.Stroke is { } stroke)
            {
                SetChartStroke(graphics, stroke);
                graphics.StrokeLine(entryX, markerY + markerSize / 2d, entryX + markerSize, markerY + markerSize / 2d);
            }

            runs.Add(new TextRun(
                entry.Name,
                entryX + markerSize + 4d,
                y,
                Math.Max(1d, entryWidth - markerSize - 4d),
                lineHeight,
                x,
                horizontal ? firstY : Math.Max(0d, firstY - (entries.Count - 1) * lineHeight),
                width,
                clipHeight,
                fontSize,
                0d,
                0d,
                style.Color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: false,
                Strike: false,
                KerningEnabled: true,
                TextAlignment.Left,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false));
        }

        return RenderTextRuns(runs, graphics, "CL");
    }

    private static bool IsOoxmlTrue(string? value)
    {
        return OoxBoolean.IsTrue(value);
    }

    private static bool IsOoxmlBooleanElementEnabled(XElement? element)
    {
        return OoxBoolean.ParseElement(element);
    }

    private static bool IsOoxmlBooleanElementEnabled(XElement? element, bool defaultValue)
    {
        return OoxBoolean.ParseElement(element, defaultValue);
    }

    private static IReadOnlyList<PdfFontResource> RenderPieDataLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ShapeBounds bounds, IReadOnlyList<double> values, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, ChartDataLabelOptions labelOptions)
    {
        if (!labelOptions.HasVisibleText || values.Count == 0)
        {
            return [];
        }

        double total = values.Sum();
        if (total <= 0d)
        {
            return [];
        }

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double radius = Math.Min(width, height) * 0.34d;
        double centerX = x + width * 0.46d;
        double centerY = y + height * 0.52d;
        double labelRadius = radius * (holeSize > 0d ? Math.Max(0.62d, (1d + holeSize) / 2d) : 0.62d);
        double labelWidth = Math.Max(18d, radius * 0.55d);
        var runs = new List<TextRun>(values.Count);
        double angle = -90d;
        for (int i = 0; i < values.Count; i++)
        {
            ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(labelOptions, i);
            if (!effectiveOptions.HasVisibleText)
            {
                continue;
            }

            ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, effectiveOptions);
            double fontSize = style.FontSize;
            double labelHeight = fontSize * 1.35d;
            double sweep = values[i] / total * 360d;
            double mid = (angle + sweep / 2d) * Math.PI / 180d;
            double explosion = pointExplosions.TryGetValue(i, out double offset) ? Math.Clamp(offset / 100d, 0d, 1d) * radius * 0.22d : 0d;
            double labelX = centerX + Math.Cos(mid) * (labelRadius + explosion) - labelWidth / 2d;
            double labelY = centerY + Math.Sin(mid) * (labelRadius + explosion) - labelHeight / 2d;
            string label = FormatPieDataLabel(values[i], total, effectiveOptions);
            runs.Add(new TextRun(
                label,
                labelX,
                labelY,
                labelWidth,
                labelHeight,
                x,
                y,
                width,
                height,
                fontSize,
                0d,
                0d,
                style.Color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: false,
                Strike: false,
                KerningEnabled: true,
                TextAlignment.Center,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false));
            RenderChartShapeStyle(graphics, labelX, labelY, labelWidth, labelHeight, effectiveOptions.ShapeStyle);
            angle += sweep;
        }

        return RenderTextRuns(runs, graphics, "CP");
    }

    private static IReadOnlyList<PdfFontResource> RenderBarDataLabels(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<IReadOnlyList<double>> series,
        ChartValueExtents extents,
        bool horizontalBars,
        ChartDataLabelOptions labelOptions,
        IReadOnlyList<ChartDataLabelOptions> seriesLabelOptions,
        IReadOnlyList<string> categoryLabels,
        IReadOnlyList<string> seriesNames)
    {
        if ((!labelOptions.HasVisibleText && !seriesLabelOptions.Any(options => options.HasVisibleText)) || series.Count == 0)
        {
            return [];
        }

        int categoryCount = Math.Max(1, series.Max(values => values.Count));
        double range = Math.Max(1d, extents.Max - extents.Min);
        double zeroOffset = (0d - extents.Min) / range;
        double zeroX = plotBox.X + zeroOffset * plotBox.Width;
        double zeroY = plotBox.Y + zeroOffset * plotBox.Height;
        var runs = new List<TextRun>();
        if (horizontalBars)
        {
            double categoryHeight = plotBox.Height / categoryCount;
            double barSlot = categoryHeight * 0.82d / Math.Max(1, series.Count);
            double labelWidth = Math.Max(18d, plotBox.Width * 0.1d);
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryY = plotBox.Y + category * categoryHeight + categoryHeight * 0.09d;
                for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    IReadOnlyList<double> values = series[seriesIndex];
                    if (category >= values.Count)
                    {
                        continue;
                    }

                    double value = values[category];
                    double barWidth = Math.Abs(value) / range * plotBox.Width;
                    double barStart = value >= 0d ? zeroX : zeroX - barWidth;
                    double barEnd = value >= 0d ? zeroX + barWidth : zeroX;
                    ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), category);
                    ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, effectiveOptions);
                    double fontSize = style.FontSize;
                    double labelHeight = fontSize * 1.35d;
                    double x = ResolveHorizontalBarDataLabelX(effectiveOptions.Position, value, barStart, barEnd, labelWidth);
                    double y = categoryY + seriesIndex * barSlot + barSlot * 0.43d - labelHeight / 2d;
                    string label = FormatCartesianDataLabel(value, seriesIndex, category, effectiveOptions, categoryLabels, seriesNames);
                    if (!string.IsNullOrEmpty(label))
                    {
                        RenderChartShapeStyle(graphics, x, y, labelWidth, labelHeight, effectiveOptions.ShapeStyle);
                        runs.Add(CreateChartLabelRun(label, x, y, labelWidth, labelHeight, plotBox, style, TextAlignment.Left));
                    }
                }
            }
        }
        else
        {
            double categoryWidth = plotBox.Width / categoryCount;
            double barSlot = categoryWidth * 0.82d / Math.Max(1, series.Count);
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryX = plotBox.X + category * categoryWidth + categoryWidth * 0.09d;
                for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    IReadOnlyList<double> values = series[seriesIndex];
                    if (category >= values.Count)
                    {
                        continue;
                    }

                    double value = values[category];
                    double barHeight = Math.Abs(value) / range * plotBox.Height;
                    double x = categoryX + seriesIndex * barSlot;
                    double barBase = value >= 0d ? zeroY : zeroY - barHeight;
                    double barEnd = value >= 0d ? zeroY + barHeight : zeroY;
                    ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), category);
                    ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, effectiveOptions);
                    double fontSize = style.FontSize;
                    double labelHeight = fontSize * 1.35d;
                    double y = ResolveVerticalBarDataLabelY(effectiveOptions.Position, value, barBase, barEnd, labelHeight);
                    string label = FormatCartesianDataLabel(value, seriesIndex, category, effectiveOptions, categoryLabels, seriesNames);
                    if (!string.IsNullOrEmpty(label))
                    {
                        double labelWidth = Math.Max(1d, barSlot * 0.86d);
                        RenderChartShapeStyle(graphics, x, y, labelWidth, labelHeight, effectiveOptions.ShapeStyle);
                        runs.Add(CreateChartLabelRun(label, x, y, labelWidth, labelHeight, plotBox, style, TextAlignment.Center));
                    }
                }
            }
        }

        return RenderTextRuns(runs, graphics, "CBD");
    }

    private static IReadOnlyList<PdfFontResource> RenderLineDataLabels(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<IReadOnlyList<double>> series,
        ChartValueExtents extents,
        ChartDataLabelOptions labelOptions,
        IReadOnlyList<ChartDataLabelOptions> seriesLabelOptions,
        IReadOnlyList<string> categoryLabels,
        IReadOnlyList<string> seriesNames)
    {
        if ((!labelOptions.HasVisibleText && !seriesLabelOptions.Any(options => options.HasVisibleText)) || series.Count == 0)
        {
            return [];
        }

        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double range = Math.Max(1d, extents.Max - extents.Min);
        double labelWidth = Math.Max(18d, plotBox.Width / Math.Max(5d, pointCount * 1.5d));
        var runs = new List<TextRun>();
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            for (int i = 0; i < values.Count; i++)
            {
                double pointX = plotBox.X + (pointCount == 1 ? plotBox.Width / 2d : plotBox.Width * i / (pointCount - 1));
                double pointY = plotBox.Y + (values[i] - extents.Min) / range * plotBox.Height;
                ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), i);
                ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, effectiveOptions);
                double fontSize = style.FontSize;
                double labelHeight = fontSize * 1.35d;
                string label = FormatCartesianDataLabel(values[i], seriesIndex, i, effectiveOptions, categoryLabels, seriesNames);
                if (!string.IsNullOrEmpty(label))
                {
                    (double labelX, double labelY, TextAlignment alignment) = ResolveLineDataLabelPosition(
                        effectiveOptions.Position,
                        pointX,
                        pointY,
                        labelWidth,
                        labelHeight);
                    RenderChartShapeStyle(graphics, labelX, labelY, labelWidth, labelHeight, effectiveOptions.ShapeStyle);
                    runs.Add(CreateChartLabelRun(
                        label,
                        labelX,
                        labelY,
                        labelWidth,
                        labelHeight,
                        plotBox,
                        style,
                        alignment));
                }
            }
        }

        return RenderTextRuns(runs, graphics, "CLD");
    }

    private static double ResolveHorizontalBarDataLabelX(string position, double value, double barStart, double barEnd, double labelWidth)
    {
        return position switch
        {
            "ctr" or "bestFit" => (barStart + barEnd - labelWidth) / 2d,
            "inBase" => value >= 0d ? barStart + 2d : barEnd - labelWidth - 2d,
            "inEnd" or "l" or "r" => value >= 0d ? barEnd - labelWidth - 2d : barStart + 2d,
            "outEnd" or _ => value >= 0d ? barEnd + 2d : barStart - labelWidth - 2d
        };
    }

    private static double ResolveVerticalBarDataLabelY(string position, double value, double barBase, double barEnd, double labelHeight)
    {
        return position switch
        {
            "ctr" or "bestFit" => (barBase + barEnd - labelHeight) / 2d,
            "inBase" => value >= 0d ? barBase + 1d : barEnd - labelHeight - 1d,
            "inEnd" or "t" or "b" => value >= 0d ? barEnd - labelHeight - 1d : barBase + 1d,
            "outEnd" or _ => value >= 0d ? barEnd + 1d : barBase - labelHeight - 1d
        };
    }

    private static (double X, double Y, TextAlignment Alignment) ResolveLineDataLabelPosition(string position, double pointX, double pointY, double labelWidth, double labelHeight)
    {
        return position switch
        {
            "b" => (pointX - labelWidth / 2d, pointY - labelHeight * 1.25d, TextAlignment.Center),
            "l" => (pointX - labelWidth - 2d, pointY - labelHeight / 2d, TextAlignment.Right),
            "r" => (pointX + 2d, pointY - labelHeight / 2d, TextAlignment.Left),
            "ctr" or "bestFit" => (pointX - labelWidth / 2d, pointY - labelHeight / 2d, TextAlignment.Center),
            "t" or "outEnd" or _ => (pointX - labelWidth / 2d, pointY + labelHeight * 0.35d, TextAlignment.Center)
        };
    }

    private static ChartTextStyle ResolveChartDataLabelTextStyle(PptxTheme theme, ChartDataLabelOptions options)
    {
        RgbColor fallbackColor = theme.TryResolveColor("tx1", out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        ChartTextStyle style = new(ResolveChartThemeFontFamily(theme), PptxChartMetricRules.DataLabelFallbackFontSize, fallbackColor, Bold: false, Italic: false);
        return MergeChartTextStyle(style, options.TextStyle);
    }

    private static ChartDataLabelOptions ResolveChartDataLabelOptions(ChartDataLabelOptions options, int index)
    {
        if (!options.Overrides.TryGetValue(index, out ChartDataLabelOverride dataLabel))
        {
            return options;
        }

        ChartTextStyleOverride textStyle = new(
            dataLabel.TextStyle.FontFamily ?? options.TextStyle.FontFamily,
            dataLabel.TextStyle.FontSize ?? options.TextStyle.FontSize,
            dataLabel.TextStyle.Color ?? options.TextStyle.Color,
            dataLabel.TextStyle.Bold ?? options.TextStyle.Bold,
            dataLabel.TextStyle.Italic ?? options.TextStyle.Italic);
        return options with
        {
            ShowValue = dataLabel.ShowValue ?? options.ShowValue,
            ShowPercent = dataLabel.ShowPercent ?? options.ShowPercent,
            ShowCategoryName = dataLabel.ShowCategoryName ?? options.ShowCategoryName,
            ShowSeriesName = dataLabel.ShowSeriesName ?? options.ShowSeriesName,
            ShowLeaderLines = dataLabel.ShowLeaderLines ?? options.ShowLeaderLines,
            CustomText = string.IsNullOrEmpty(dataLabel.CustomText) ? options.CustomText : dataLabel.CustomText,
            Position = string.IsNullOrEmpty(dataLabel.Position) ? options.Position : dataLabel.Position,
            Separator = string.IsNullOrEmpty(dataLabel.Separator) ? options.Separator : dataLabel.Separator,
            NumberFormat = string.IsNullOrEmpty(dataLabel.NumberFormat) ? options.NumberFormat : dataLabel.NumberFormat,
            TextStyle = textStyle,
            ShapeStyle = dataLabel.ShapeStyle.IsEmpty ? options.ShapeStyle : dataLabel.ShapeStyle,
            Overrides = EmptyChartDataLabelOverrides
        };
    }

    private static ChartDataLabelOptions ResolveChartDataLabelOptionsForSeries(ChartDataLabelOptions plotOptions, IReadOnlyList<ChartDataLabelOptions> seriesOptions, int seriesIndex)
    {
        return seriesIndex < seriesOptions.Count && seriesOptions[seriesIndex].IsDefined
            ? seriesOptions[seriesIndex]
            : plotOptions;
    }

    private static TextRun CreateChartLabelRun(string text, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment)
    {
        return new TextRun(
            text,
            x,
            y,
            Math.Max(1d, width),
            height,
            plotBox.X,
            plotBox.Y,
            plotBox.Width,
            plotBox.Height,
            style.FontSize,
            0d,
            0d,
            style.Color,
            1d,
            null,
            Bold: style.Bold,
            Italic: style.Italic,
            Underline: false,
            Strike: false,
            KerningEnabled: true,
            alignment,
            FontFamily: style.FontFamily,
            RotationDegrees: 0d,
            RotationCenterX: 0d,
            RotationCenterY: 0d,
            FlipHorizontal: false,
            FlipVertical: false);
    }

    private static ChartTextStyle ReadChartTextStyle(PptxTheme theme, XDocument chartXml, XElement? element, double fallbackFontSize)
    {
        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize);
        style = MergeChartTextStyle(style, ReadChartTextStyleFromTxPr(chartXml.Root, theme));
        style = MergeChartTextStyle(style, ReadChartTextStyleFromTxPr(element, theme));
        return style;
    }

    private static ChartTextStyle ReadSceneOrXmlChartTextStyle(PptxTheme theme, PptxSceneChart? sceneChart, PptxSceneChartAxis? sceneAxis, XDocument chartXml, XElement? element, double fallbackFontSize)
    {
        if (sceneChart is null)
        {
            return ReadChartTextStyle(theme, chartXml, element, fallbackFontSize);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize);
        style = MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.TextStyle));
        style = sceneAxis is null
            ? MergeChartTextStyle(style, ReadChartTextStyleFromTxPr(element, theme))
            : MergeChartTextStyle(style, ToChartTextStyleOverride(sceneAxis.TextStyle));
        return style;
    }

    private static ChartTextStyle CreateDefaultChartTextStyle(PptxTheme theme, double fallbackFontSize)
    {
        RgbColor fallbackColor = theme.TryResolveColor("tx1", out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        return new ChartTextStyle(ResolveChartThemeFontFamily(theme), fallbackFontSize, fallbackColor, Bold: false, Italic: false);
    }

    private static ChartTextStyleOverride ToChartTextStyleOverride(PptxSceneChartTextStyleOverride style)
    {
        return new ChartTextStyleOverride(style.FontFamily, style.FontSize, style.Color, style.Bold, style.Italic);
    }

    private static string? ResolveChartThemeFontFamily(PptxTheme theme)
    {
        return theme.ResolveTypeface("+mn-lt") ??
            theme.ResolveTypeface("+mj-lt");
    }

    private static ChartTextStyleOverride ReadChartTextStyleFromTxPr(XElement? parent, PptxTheme theme)
    {
        XElement? defRunProperties = parent?
            .Element(ChartNamespace + "txPr")?
            .Elements(DrawingNamespace + "p")
            .Select(paragraph => paragraph.Element(DrawingNamespace + "pPr")?.Element(DrawingNamespace + "defRPr"))
            .FirstOrDefault(element => element is not null);
        if (defRunProperties is null)
        {
            return ChartTextStyleOverride.Empty;
        }

        string? typeface = (string?)defRunProperties.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)defRunProperties.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)defRunProperties.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        string? fontFamily = string.IsNullOrWhiteSpace(typeface)
            ? null
            : theme.ResolveTypeface(typeface);

        double? fontSize = null;
        if (defRunProperties.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0)
        {
            fontSize = sizeHundredths / 100d;
        }

        RgbColor? color = TryReadSolidColor(defRunProperties.Element(DrawingNamespace + "solidFill"), theme, out RgbColor parsedColor)
            ? parsedColor
            : null;
        bool? bold = defRunProperties.Attribute("b") is { } boldAttribute
            ? IsOoxmlTrue(boldAttribute.Value)
            : null;
        bool? italic = defRunProperties.Attribute("i") is { } italicAttribute
            ? IsOoxmlTrue(italicAttribute.Value)
            : null;
        return new ChartTextStyleOverride(fontFamily, fontSize, color, bold, italic);
    }

    private static ChartTextStyle MergeChartTextStyle(ChartTextStyle style, ChartTextStyleOverride next)
    {
        return new ChartTextStyle(
            next.FontFamily ?? style.FontFamily,
            next.FontSize ?? style.FontSize,
            next.Color ?? style.Color,
            next.Bold ?? style.Bold,
            next.Italic ?? style.Italic);
    }

    private static ChartDataLabelOptions ReadChartDataLabelOptions(XElement chartElement, PptxTheme theme)
    {
        XElement? labels = chartElement.Element(ChartNamespace + "dLbls") ??
            chartElement.Elements(ChartNamespace + "ser")
                .Select(series => series.Element(ChartNamespace + "dLbls"))
                .FirstOrDefault(element => element is not null);
        return labels is null
            ? ChartDataLabelOptions.None
            : new ChartDataLabelOptions(
                IsChartLabelFlagEnabled(labels, "showVal"),
                IsChartLabelFlagEnabled(labels, "showPercent"),
                IsChartLabelFlagEnabled(labels, "showCatName"),
                IsChartLabelFlagEnabled(labels, "showSerName"),
                IsChartLabelFlagEnabled(labels, "showLeaderLines"),
                string.Empty,
                labels.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty,
                labels.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                labels.Element(ChartNamespace + "numFmt")?.Attribute("formatCode")?.Value ?? string.Empty,
                ReadChartTextStyleFromTxPr(labels, theme),
                ReadChartShapeStyle(labels.Element(ChartNamespace + "spPr"), theme),
                ReadChartDataLabelOverrides(labels, theme),
                IsDefined: true);
    }

    private static ChartDataLabelOptions ReadSceneOrXmlDataLabelOptions(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme)
    {
        return plot is null
            ? ReadChartDataLabelOptions(chartElement, theme)
            : new ChartDataLabelOptions(
                plot.DataLabels.ShowValue,
                plot.DataLabels.ShowPercent,
                plot.DataLabels.ShowCategoryName,
                plot.DataLabels.ShowSeriesName,
                plot.DataLabels.ShowLeaderLines,
                string.Empty,
                plot.DataLabels.Position,
                plot.DataLabels.Separator,
                plot.DataLabels.NumberFormat,
                ToChartTextStyleOverride(plot.DataLabels.TextStyle),
                ToChartShapeStyle(plot.DataLabels.ShapeStyle),
                ToChartDataLabelOverrides(plot.DataLabels.Overrides),
                plot.DataLabels.IsDefined);
    }

    private static IReadOnlyList<ChartDataLabelOptions> ReadSceneOrXmlSeriesDataLabelOptions(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => ToChartDataLabelOptions(series.DataLabels))
                .ToArray();
        }

        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => ReadChartDataLabelOptions(series, theme))
            .ToArray();
    }

    private static ChartDataLabelOptions ToChartDataLabelOptions(PptxSceneChartDataLabels labels)
    {
        return new ChartDataLabelOptions(
            labels.ShowValue,
            labels.ShowPercent,
            labels.ShowCategoryName,
            labels.ShowSeriesName,
            labels.ShowLeaderLines,
            string.Empty,
            labels.Position,
            labels.Separator,
            labels.NumberFormat,
            ToChartTextStyleOverride(labels.TextStyle),
            ToChartShapeStyle(labels.ShapeStyle),
            ToChartDataLabelOverrides(labels.Overrides),
            labels.IsDefined);
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> ReadChartDataLabelOverrides(XElement labels, PptxTheme theme)
    {
        var overrides = new Dictionary<int, ChartDataLabelOverride>();
        foreach (XElement label in labels.Elements(ChartNamespace + "dLbl"))
        {
            if (!int.TryParse(label.Element(ChartNamespace + "idx")?.Attribute("val")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) ||
                index < 0)
            {
                continue;
            }

            overrides[index] = new ChartDataLabelOverride(
                ReadOptionalChartLabelFlagEnabled(label, "showVal"),
                ReadOptionalChartLabelFlagEnabled(label, "showPercent"),
                ReadOptionalChartLabelFlagEnabled(label, "showCatName"),
                ReadOptionalChartLabelFlagEnabled(label, "showSerName"),
                ReadOptionalChartLabelFlagEnabled(label, "showLeaderLines"),
                ReadChartText(label.Element(ChartNamespace + "tx")) ?? string.Empty,
                label.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty,
                label.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                label.Element(ChartNamespace + "numFmt")?.Attribute("formatCode")?.Value ?? string.Empty,
                ReadChartTextStyleFromTxPr(label, theme),
                ReadChartShapeStyle(label.Element(ChartNamespace + "spPr"), theme));
        }

        return overrides.Count == 0 ? EmptyChartDataLabelOverrides : overrides;
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> ToChartDataLabelOverrides(IReadOnlyList<PptxSceneChartDataLabelOverride> overrides)
    {
        if (overrides.Count == 0)
        {
            return EmptyChartDataLabelOverrides;
        }

        var result = new Dictionary<int, ChartDataLabelOverride>(overrides.Count);
        foreach (PptxSceneChartDataLabelOverride dataLabel in overrides)
        {
            result[dataLabel.Index] = new ChartDataLabelOverride(
                dataLabel.ShowValue,
                dataLabel.ShowPercent,
                dataLabel.ShowCategoryName,
                dataLabel.ShowSeriesName,
                dataLabel.ShowLeaderLines,
                dataLabel.CustomText,
                dataLabel.Position,
                dataLabel.Separator,
                dataLabel.NumberFormat,
                ToChartTextStyleOverride(dataLabel.TextStyle),
                ToChartShapeStyle(dataLabel.ShapeStyle));
        }

        return result;
    }

    private static bool? ReadOptionalChartLabelFlagEnabled(XElement labels, string elementName)
    {
        XElement? element = labels.Element(ChartNamespace + elementName);
        return element is null ? null : IsOoxmlBooleanElementEnabled(element);
    }

    private static bool IsChartLabelFlagEnabled(XElement labels, string elementName)
    {
        XElement? element = labels.Element(ChartNamespace + elementName);
        return IsOoxmlBooleanElementEnabled(element);
    }

    private static string FormatChartPercentageLabel(double fraction)
    {
        return (fraction * 100d).ToString("0.#", CultureInfo.InvariantCulture) + "%";
    }

    private static string FormatPieDataLabel(double value, double total, ChartDataLabelOptions options)
    {
        if (!string.IsNullOrWhiteSpace(options.CustomText))
        {
            return options.CustomText;
        }

        if (options.ShowValue && options.ShowPercent)
        {
            return FormatChartDataLabelValue(value, options) + GetChartDataLabelSeparator(options) + FormatChartPercentageLabel(value / total);
        }

        return options.ShowPercent
            ? FormatChartPercentageLabel(value / total)
            : FormatChartDataLabelValue(value, options);
    }

    private static string FormatCartesianDataLabel(
        double value,
        int seriesIndex,
        int categoryIndex,
        ChartDataLabelOptions options,
        IReadOnlyList<string> categoryLabels,
        IReadOnlyList<string> seriesNames)
    {
        if (!string.IsNullOrWhiteSpace(options.CustomText))
        {
            return options.CustomText;
        }

        var parts = new List<string>(3);
        if (options.ShowSeriesName && seriesIndex < seriesNames.Count && !string.IsNullOrWhiteSpace(seriesNames[seriesIndex]))
        {
            parts.Add(seriesNames[seriesIndex]);
        }

        if (options.ShowCategoryName && categoryIndex < categoryLabels.Count && !string.IsNullOrWhiteSpace(categoryLabels[categoryIndex]))
        {
            parts.Add(categoryLabels[categoryIndex]);
        }

        if (options.ShowValue)
        {
            parts.Add(FormatChartDataLabelValue(value, options));
        }

        return string.Join(GetChartDataLabelSeparator(options), parts);
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options)
    {
        return !string.IsNullOrWhiteSpace(options.NumberFormat) &&
            !string.Equals(options.NumberFormat, "General", StringComparison.OrdinalIgnoreCase)
            ? FormatChartNumber(value, options.NumberFormat)
            : FormatChartAxisLabel(value);
    }

    private static string GetChartDataLabelSeparator(ChartDataLabelOptions options)
    {
        return string.IsNullOrEmpty(options.Separator) ? ", " : options.Separator;
    }

    private static IReadOnlyList<PdfFontResource> RenderChartCategoryLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartAxis? sceneAxis, XElement? categoryAxis, IReadOnlyList<string> labels, bool horizontalBars)
    {
        if (labels.Count == 0)
        {
            return [];
        }

        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, categoryAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize);
        double fontSize = style.FontSize;
        RgbColor color = style.Color;
        var runs = new List<TextRun>(labels.Count);
        for (int i = 0; i < labels.Count; i++)
        {
            double x;
            double y;
            double width;
            double height = fontSize * 1.35d;
            TextAlignment alignment;
            if (horizontalBars)
            {
                double slotHeight = plotBox.Height / labels.Count;
                x = Math.Max(0d, plotBox.X - plotBox.Width * 0.18d);
                y = plotBox.Y + slotHeight * (i + 0.5d) - height * 0.45d;
                width = plotBox.Width * 0.16d;
                alignment = TextAlignment.Right;
            }
            else
            {
                double slotWidth = plotBox.Width / labels.Count;
                width = slotWidth * 1.35d;
                x = plotBox.X + slotWidth * (i + 0.5d) - width / 2d;
                y = plotBox.Y - height * 1.25d;
                alignment = TextAlignment.Center;
            }

            double labelWidth = Math.Max(1d, width);
            runs.Add(new TextRun(
                labels[i],
                x,
                y,
                labelWidth,
                height,
                x,
                y - height * 0.25d,
                labelWidth,
                height * 1.6d,
                fontSize,
                0d,
                0d,
                color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: false,
                Strike: false,
                KerningEnabled: true,
                alignment,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false));
        }

        return RenderTextRuns(runs, graphics, "CCA");
    }

    private static IReadOnlyList<PdfFontResource> RenderChartValueAxisLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, XElement? valueAxis, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits axisUnits, bool horizontalBars, bool rightSide = false, int axisSideSlot = 0, bool useTextSizedWidth = false)
    {
        double range = Math.Max(1d, extents.Max - extents.Min);
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize);
        double fontSize = style.FontSize;
        double height = fontSize * 1.35d;
        RgbColor color = style.Color;
        IReadOnlyList<double> tickValues = GetChartAxisTickValues(extents, axisUnits.MajorUnit, includeEndpoints: true);
        double maxLabelWidth = tickValues
            .Select(value => FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis))
            .DefaultIfEmpty("0")
            .Max(label => EstimateChartTextWidth(label, fontSize));
        double valueAxisLabelWidth = Math.Max(fontSize * 1.6d, maxLabelWidth + fontSize * 0.45d);
        var runs = new List<TextRun>(tickValues.Count);
        foreach (double value in tickValues)
        {
            string label = FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis);
            double offset = (value - extents.Min) / range;
            double x;
            double y;
            double width;
            TextAlignment alignment;
            if (horizontalBars)
            {
                width = plotBox.Width / 5d;
                x = plotBox.X + plotBox.Width * offset - width / 2d;
                y = plotBox.Y - height * 1.35d;
                alignment = TextAlignment.Center;
            }
            else
            {
                bool labelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(sceneAxis, valueAxis, rightSide);
                width = useTextSizedWidth ? valueAxisLabelWidth : plotBox.Width * 0.12d;
                double sideGap = Math.Max(3d, fontSize * 0.45d);
                x = labelsRightSide
                    ? plotBox.X + plotBox.Width + sideGap + axisSideSlot * (width + sideGap)
                    : Math.Max(0d, plotBox.X - (axisSideSlot + 1) * (width + sideGap));
                y = plotBox.Y + plotBox.Height * offset - height * 0.45d;
                alignment = labelsRightSide ? TextAlignment.Left : TextAlignment.Right;
            }

            double labelWidth = Math.Max(1d, width);
            double clipY = horizontalBars ? y : plotBox.Y - height;
            double clipHeight = horizontalBars ? height : plotBox.Height + height * 2d;
            runs.Add(new TextRun(
                label,
                x,
                y,
                labelWidth,
                height,
                x,
                clipY,
                labelWidth,
                clipHeight,
                fontSize,
                0d,
                0d,
                color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: false,
                Strike: false,
                KerningEnabled: true,
                alignment,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false));
        }

        return RenderTextRuns(runs, graphics, "CVA");
    }

    private static double EstimateChartTextWidth(string text, double fontSize)
    {
        double width = 0d;
        foreach (char ch in text)
        {
            width += char.IsWhiteSpace(ch)
                ? fontSize * 0.32d
                : char.IsDigit(ch)
                    ? fontSize * 0.55d
                    : fontSize * 0.58d;
        }

        return width;
    }

    private static int GetValueAxisSideSlot(PptxSceneChartAxis? primarySceneAxis, XElement? primaryAxis, PptxSceneChartAxis? secondarySceneAxis, XElement secondaryAxis, bool defaultPrimaryRightSide, bool defaultSecondaryRightSide)
    {
        bool primaryRight = ResolveSceneOrXmlValueAxisLabelsRightSide(primarySceneAxis, primaryAxis, defaultPrimaryRightSide);
        bool secondaryRight = ResolveSceneOrXmlValueAxisLabelsRightSide(secondarySceneAxis, secondaryAxis, defaultSecondaryRightSide);
        return primaryRight == secondaryRight ? 1 : 0;
    }

    private static IReadOnlyList<PdfFontResource> RenderSecondaryChartValueAxisLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, ChartValueExtents fallback)
    {
        XElement? rightValueAxis = ReadSecondaryRightValueAxis(chartXml);
        PptxSceneChartAxis? rightSceneAxis = ReadSceneSecondaryRightValueAxis(sceneChart, rightValueAxis);
        if (rightValueAxis is null && rightSceneAxis is null)
        {
            return Array.Empty<PdfFontResource>();
        }

        ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(rightSceneAxis, rightValueAxis, fallback);
        ChartAxisUnits axisUnits = ReadSceneOrXmlChartValueAxisUnits(rightSceneAxis, rightValueAxis);
        return RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, rightValueAxis, rightSceneAxis, extents, axisUnits, horizontalBars: false, rightSide: true);
    }

    private static XElement? ReadSecondaryRightValueAxis(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "valAx")
            .Where(IsRightValueAxis)
            .FirstOrDefault();
    }

    private static PptxSceneChartAxis? ReadSceneSecondaryRightValueAxis(PptxSceneChart? sceneChart, XElement? rightValueAxis)
    {
        if (sceneChart is null)
        {
            return null;
        }

        string? axisId = ReadChartAxisId(rightValueAxis);
        if (!string.IsNullOrWhiteSpace(axisId))
        {
            PptxSceneChartAxis? matchingAxis = sceneChart.Axes.FirstOrDefault(axis =>
                string.Equals(axis.Id, axisId, StringComparison.Ordinal) &&
                string.Equals(axis.Kind, "valAx", StringComparison.Ordinal));
            if (matchingAxis is not null)
            {
                return matchingAxis;
            }
        }

        return sceneChart.Axes.FirstOrDefault(axis =>
            string.Equals(axis.Kind, "valAx", StringComparison.Ordinal) &&
            string.Equals(axis.Position, "r", StringComparison.Ordinal));
    }

    private static bool IsSceneOrXmlRightValueAxis(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is not null)
        {
            return string.Equals(sceneAxis.Kind, "valAx", StringComparison.Ordinal) &&
                string.Equals(sceneAxis.Position, "r", StringComparison.Ordinal) &&
                !sceneAxis.IsDeleted;
        }

        return IsRightValueAxis(axis);
    }

    private static bool IsRightValueAxis(XElement? axis)
    {
        return axis is not null &&
            string.Equals((string?)axis.Element(ChartNamespace + "axPos")?.Attribute("val"), "r", StringComparison.Ordinal) &&
            !IsOoxmlBooleanElementEnabled(axis.Element(ChartNamespace + "delete"));
    }

    private static string? ReadChartAxisId(XElement? axis)
    {
        return (string?)axis?.Element(ChartNamespace + "axId")?.Attribute("val");
    }

    private static IReadOnlyList<double> GetChartAxisTickValues(ChartValueExtents extents, double? explicitUnit, bool includeEndpoints)
    {
        double range = Math.Max(1d, extents.Max - extents.Min);
        if (explicitUnit is not { } unit || unit <= 0d)
        {
            unit = ChooseChartAxisMajorUnit(range);
        }

        var values = new List<double>();
        if (includeEndpoints)
        {
            values.Add(extents.Min);
        }

        double first = Math.Ceiling(extents.Min / unit) * unit;
        for (double value = first; value < extents.Max - 0.0001d; value += unit)
        {
            if (value > extents.Min + 0.0001d)
            {
                values.Add(value);
            }
        }

        if (includeEndpoints)
        {
            values.Add(extents.Max);
        }

        return values;
    }

    private static double ChooseChartAxisMajorUnit(double range)
    {
        double target = Math.Max(range / 10d, double.Epsilon);
        double magnitude = Math.Pow(10d, Math.Floor(Math.Log10(target)));
        double normalized = target / magnitude;
        double nice = normalized <= 1d
            ? 1d
            : normalized <= 2d
                ? 2d
                : normalized <= 5d
                    ? 5d
                    : 10d;
        return nice * magnitude;
    }

    private static double GetNiceChartAxisMax(double dataMax, double dataMin)
    {
        double interval = GetNiceChartAxisInterval(dataMax, dataMin, desiredTicks: 5);
        double niceMax = Math.Ceiling(dataMax / interval) * interval;
        return niceMax <= dataMax ? niceMax + interval : niceMax;
    }

    private static double GetNiceChartAxisInterval(double dataMax, double dataMin, int desiredTicks)
    {
        if (Math.Abs(dataMax) < 0.0001d && Math.Abs(dataMin) < 0.0001d)
        {
            return 1d;
        }

        double range = dataMax - Math.Min(0d, dataMin);
        if (Math.Abs(range) < 0.0001d)
        {
            return dataMax > 0d ? dataMax * PptxChartMetricRules.AxisSingleValueHeadroomFactor : 1d;
        }

        double rawInterval = Math.Max(range / Math.Max(1, desiredTicks), double.Epsilon);
        double magnitude = Math.Pow(10d, Math.Floor(Math.Log10(rawInterval)));
        double normalized = rawInterval / magnitude;
        double nice = normalized <= 1d
            ? 1d
            : normalized <= 2d
                ? 2d
                : normalized <= 5d
                    ? 5d
                    : 10d;
        return nice * magnitude;
    }

    private static string FormatChartAxisLabel(double value, XElement? axis = null)
    {
        string? formatCode = (string?)axis?
            .Element(ChartNamespace + "numFmt")
            ?.Attribute("formatCode");
        if (!string.IsNullOrWhiteSpace(formatCode) &&
            !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase))
        {
            return FormatChartNumber(value, formatCode);
        }

        double rounded = Math.Round(value);
        return Math.Abs(value - rounded) < 0.0001d
            ? rounded.ToString("0", CultureInfo.InvariantCulture)
            : value.ToString("0.##", CultureInfo.InvariantCulture);
    }

    private static string FormatSceneOrXmlChartAxisLabel(double value, PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        string? formatCode = sceneAxis?.NumberFormat;
        if (!string.IsNullOrWhiteSpace(formatCode) &&
            !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase))
        {
            return FormatChartNumber(value, formatCode);
        }

        return FormatChartAxisLabel(value, axis);
    }

    private static string FormatChartNumber(double value, string formatCode)
    {
        string normalized = formatCode.Replace("\\", string.Empty, StringComparison.Ordinal);
        bool percent = normalized.Contains('%', StringComparison.Ordinal);
        double displayValue = percent ? value * 100d : value;
        int decimals = 0;
        int decimalPoint = normalized.IndexOf('.', StringComparison.Ordinal);
        if (decimalPoint >= 0)
        {
            decimals = normalized
                .Skip(decimalPoint + 1)
                .TakeWhile(ch => ch is '0' or '#')
                .Count();
        }

        bool thousands = normalized.Contains(',', StringComparison.Ordinal);
        string numberFormat = (thousands ? "#,##0" : "0") +
            (decimals > 0 ? "." + new string('0', decimals) : string.Empty);
        string text = displayValue.ToString(numberFormat, CultureInfo.InvariantCulture);
        if (normalized.Contains('$', StringComparison.Ordinal))
        {
            text = "$" + text;
        }

        if (percent)
        {
            text += "%";
        }

        return text;
    }

    private static IReadOnlyList<ChartSeriesFill?> ReadChartSeriesFills(XElement chartElement, PptxTheme theme)
    {
        var fills = new List<ChartSeriesFill?>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? shapeProperties = element.Element(ChartNamespace + "spPr");
            fills.Add(TryReadChartFill(shapeProperties, theme));
        }

        return fills;
    }

    private static IReadOnlyDictionary<int, ChartSeriesFill> ReadChartPointFills(XElement chartElement, PptxTheme theme)
    {
        var fills = new Dictionary<int, ChartSeriesFill>();
        XElement? series = chartElement.Name == ChartNamespace + "ser"
            ? chartElement
            : chartElement.Element(ChartNamespace + "ser");
        if (series is null)
        {
            return fills;
        }

        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (point.Element(ChartNamespace + "idx")?.Attribute("val") is not { } indexAttribute ||
                !int.TryParse(indexAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            if (TryReadChartFill(shapeProperties, theme) is { } fill)
            {
                fills[index] = fill;
            }
        }

        return fills;
    }

    private static ChartSeriesFill? TryReadChartFill(XElement? shapeProperties, PptxTheme theme)
    {
        if (TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor color, out double alpha))
        {
            return new ChartSeriesFill(color, alpha);
        }

        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        if (patternFill is null)
        {
            return null;
        }

        RgbColor foreground = TryReadSolidColor(patternFill.Element(DrawingNamespace + "fgClr"), theme, out RgbColor foregroundColor)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = TryReadSolidColor(patternFill.Element(DrawingNamespace + "bgClr"), theme, out RgbColor backgroundColor)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        string preset = (string?)patternFill.Attribute("prst") ?? "pct50";
        return new ChartSeriesFill(foreground, 1d, preset, background);
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> ReadChartSeriesPointFills(XElement chartElement, PptxTheme theme)
    {
        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => ReadChartPointFills(series, theme))
            .ToArray();
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> ReadChartSeriesPointStrokes(XElement chartElement, PptxTheme theme)
    {
        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => ReadChartPointStrokes(series, theme))
            .ToArray();
    }

    private static IReadOnlyDictionary<int, ChartSeriesStroke> ReadChartPointStrokes(XElement chartElement, PptxTheme theme)
    {
        var strokes = new Dictionary<int, ChartSeriesStroke>();
        XElement? series = chartElement.Name == ChartNamespace + "ser"
            ? chartElement
            : chartElement.Element(ChartNamespace + "ser");
        if (series is null)
        {
            return strokes;
        }

        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (point.Element(ChartNamespace + "idx")?.Attribute("val") is not { } indexAttribute ||
                !int.TryParse(indexAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            if (shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, out RgbColor color, out double width, out double alpha))
            {
                strokes[index] = new ChartSeriesStroke(color, alpha, width);
            }
        }

        return strokes;
    }

    private static IReadOnlyDictionary<int, double> ReadChartPointExplosions(XElement chartElement)
    {
        var explosions = new Dictionary<int, double>();
        XElement? series = chartElement.Element(ChartNamespace + "ser");
        if (series is null)
        {
            return explosions;
        }

        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (point.Element(ChartNamespace + "idx")?.Attribute("val") is not { } indexAttribute ||
                !int.TryParse(indexAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                continue;
            }

            if (point.Element(ChartNamespace + "explosion")?.Attribute("val") is { } explosionAttribute &&
                double.TryParse(explosionAttribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double explosion))
            {
                explosions[index] = Math.Clamp(explosion / 100d, 0d, 1d);
            }
        }

        return explosions;
    }

    private static double ReadDoughnutHoleSize(XElement doughnutChart)
    {
        if (doughnutChart.Element(ChartNamespace + "holeSize")?.Attribute("val") is { } value &&
            double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
        {
            return Math.Clamp(parsed / 100d, PptxChartMetricRules.DoughnutHoleMinimumRatio, PptxChartMetricRules.DoughnutHoleMaximumRatio);
        }

        return PptxChartMetricRules.DoughnutHoleFallbackRatio;
    }

    private static IReadOnlyList<ChartSeriesStroke?> ReadChartSeriesStrokes(XElement chartElement, PptxTheme theme)
    {
        var strokes = new List<ChartSeriesStroke?>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? shapeProperties = element.Element(ChartNamespace + "spPr");
            strokes.Add(shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, out RgbColor color, out double width, out double alpha)
                    ? new ChartSeriesStroke(color, alpha, width)
                    : null);
        }

        return strokes;
    }

    private static IReadOnlyList<ChartMarkerStyle> ReadChartMarkerStyles(XElement chartElement, PptxTheme theme)
    {
        var styles = new List<ChartMarkerStyle>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? marker = element.Element(ChartNamespace + "marker");
            string symbol = (string?)marker?.Element(ChartNamespace + "symbol")?.Attribute("val") ?? "circle";
            double size = marker?.Element(ChartNamespace + "size")?.Attribute("val") is { } value &&
                double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                    ? Math.Clamp(parsed, 2d, 30d)
                    : 4d;
            XElement? shapeProperties = marker?.Element(ChartNamespace + "spPr");
            ChartSeriesFill? fill = TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
                ? new ChartSeriesFill(fillColor, fillAlpha)
                : null;
            ChartSeriesStroke? stroke = shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, out RgbColor strokeColor, out double strokeWidth, out double strokeAlpha)
                    ? new ChartSeriesStroke(strokeColor, strokeAlpha, strokeWidth)
                    : null;
            styles.Add(new ChartMarkerStyle(symbol, size, fill, stroke));
        }

        return styles;
    }

    private static IReadOnlyList<bool> ReadChartSeriesSmooth(XElement chartElement)
    {
        var smooth = new List<bool>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            string? value = (string?)element.Element(ChartNamespace + "smooth")?.Attribute("val");
            smooth.Add(value is "1" or "true");
        }

        return smooth;
    }

    private static bool HasMajorGridlines(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "majorGridlines")
            .Any(IsChartGridlineVisible);
    }

    private static bool HasMinorGridlines(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "minorGridlines")
            .Any(IsChartGridlineVisible);
    }

    private static bool ReadSceneOrXmlMajorGridlines(PptxSceneChartAxis? axis, XDocument chartXml)
    {
        return axis?.HasMajorGridlines ?? HasMajorGridlines(chartXml);
    }

    private static bool ReadSceneOrXmlMinorGridlines(PptxSceneChartAxis? axis, XDocument chartXml)
    {
        return axis?.HasMinorGridlines ?? HasMinorGridlines(chartXml);
    }

    private static ChartGridlineStyle ReadSceneOrXmlChartGridlineStyle(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis, PptxTheme theme)
    {
        return new ChartGridlineStyle(
            ReadSceneOrXmlChartGridlineStroke(sceneAxis?.MajorGridlineLine ?? default, xmlAxis?.Element(ChartNamespace + "majorGridlines"), theme),
            ReadSceneOrXmlChartGridlineStroke(sceneAxis?.MinorGridlineLine ?? default, xmlAxis?.Element(ChartNamespace + "minorGridlines"), theme));
    }

    private static ChartSeriesStroke? ReadSceneOrXmlChartGridlineStroke(PptxSceneLineStyle sceneLine, XElement? gridlines, PptxTheme theme)
    {
        if (sceneLine.HasLine)
        {
            return ToChartSeriesStroke(sceneLine);
        }

        return ReadChartGridlineStroke(gridlines, theme);
    }

    private static ChartSeriesStroke? ReadChartGridlineStroke(XElement? gridlines, PptxTheme theme)
    {
        XElement? shapeProperties = gridlines?.Element(ChartNamespace + "spPr");
        if (shapeProperties?.Element(DrawingNamespace + "ln")?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new ChartSeriesStroke(new RgbColor(0, 0, 0), 0d, 0d);
        }

        return shapeProperties is not null &&
            TryReadLineWithAlpha(shapeProperties, theme, out RgbColor color, out double width, out double alpha)
                ? new ChartSeriesStroke(color, alpha, width)
                : null;
    }

    private static bool IsChartGridlineVisible(XElement gridlines)
    {
        XElement? line = gridlines
            .Element(ChartNamespace + "spPr")
            ?.Element(DrawingNamespace + "ln");
        return line?.Element(DrawingNamespace + "noFill") is null;
    }

    private static bool ReadChartVaryColors(XElement chartElement)
    {
        return IsOoxmlBooleanElementEnabled(chartElement.Element(ChartNamespace + "varyColors"), defaultValue: true);
    }

    private static double ReadChartGapWidth(XElement chartElement)
    {
        string? value = (string?)chartElement.Element(ChartNamespace + "gapWidth")?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? Math.Clamp(parsed, 0d, 500d)
            : 150d;
    }

    private static double ReadChartOverlap(XElement chartElement)
    {
        string? value = (string?)chartElement.Element(ChartNamespace + "overlap")?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? Math.Clamp(parsed, -100d, 100d)
            : 0d;
    }

    private static ChartAxesStyle ReadChartAxesStyle(XDocument chartXml, PptxTheme theme, XElement chartElement)
    {
        XElement? valueAxisElement = ReadChartValueAxisForChart(chartXml, chartElement) ??
            chartXml.Descendants(ChartNamespace + "valAx").FirstOrDefault();
        XElement? categoryAxisElement = ReadChartCategoryAxisForChart(chartXml, chartElement);
        ChartSeriesStroke? valueAxis = ReadChartAxisStroke(valueAxisElement, theme);
        ChartSeriesStroke? categoryAxis = ReadChartAxisStroke(categoryAxisElement, theme);
        return new ChartAxesStyle(
            valueAxis,
            ReadChartAxisStroke(chartXml
                .Descendants(ChartNamespace + "valAx")
                .FirstOrDefault(axis => string.Equals((string?)axis.Element(ChartNamespace + "axPos")?.Attribute("val"), "r", StringComparison.Ordinal)),
                theme),
            categoryAxis,
            !IsChartAxisDeleted(valueAxisElement),
            !IsChartAxisDeleted(categoryAxisElement));
    }

    private static ChartAxesStyle ReadSceneOrXmlChartAxesStyle(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, PptxTheme theme, XElement chartElement)
    {
        XElement? valueAxisElement = ReadChartValueAxisForChart(chartXml, chartElement) ??
            chartXml.Descendants(ChartNamespace + "valAx").FirstOrDefault();
        XElement? categoryAxisElement = ReadChartCategoryAxisForChart(chartXml, chartElement);
        XElement? secondaryValueAxisElement = chartXml
            .Descendants(ChartNamespace + "valAx")
            .FirstOrDefault(axis => string.Equals((string?)axis.Element(ChartNamespace + "axPos")?.Attribute("val"), "r", StringComparison.Ordinal));
        PptxSceneChartAxis? valueAxis = ReadSceneChartAxis(sceneChart, plot, "valAx");
        PptxSceneChartAxis? categoryAxis = ReadSceneChartAxis(sceneChart, plot, "catAx");
        PptxSceneChartAxis? secondaryValueAxis = ReadSceneSecondaryRightValueAxis(sceneChart, secondaryValueAxisElement);
        return new ChartAxesStyle(
            ReadSceneOrXmlChartAxisStroke(valueAxis, valueAxisElement, theme),
            ReadSceneOrXmlChartAxisStroke(secondaryValueAxis, secondaryValueAxisElement, theme),
            ReadSceneOrXmlChartAxisStroke(categoryAxis, categoryAxisElement, theme),
            valueAxis is null ? !IsChartAxisDeleted(valueAxisElement) : !valueAxis.IsDeleted,
            categoryAxis is null ? !IsChartAxisDeleted(categoryAxisElement) : !categoryAxis.IsDeleted);
    }

    private static ChartSeriesStroke? ReadSceneOrXmlChartAxisStroke(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis, PptxTheme theme)
    {
        return sceneAxis is not null && sceneAxis.Line.HasLine
            ? ToChartSeriesStroke(sceneAxis.Line)
            : ReadChartAxisStroke(xmlAxis, theme);
    }

    private static bool IsChartAxisDeleted(XElement? axis)
    {
        XElement? delete = axis?.Element(ChartNamespace + "delete");
        return IsOoxmlBooleanElementEnabled(delete);
    }

    private static bool IsChartAxisLabelVisible(XElement? axis)
    {
        if (IsChartAxisDeleted(axis))
        {
            return false;
        }

        string? tickLabelPosition = (string?)axis?.Element(ChartNamespace + "tickLblPos")?.Attribute("val");
        return !string.Equals(tickLabelPosition, "none", StringComparison.Ordinal);
    }

    private static bool IsSceneOrXmlChartAxisLabelVisible(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is null)
        {
            return IsChartAxisLabelVisible(axis);
        }

        return !sceneAxis.IsDeleted &&
            !string.Equals(sceneAxis.TickLabelPosition, "none", StringComparison.Ordinal);
    }

    private static bool ResolveValueAxisLabelsRightSide(XElement? axis, bool defaultRightSide)
    {
        string? tickLabelPosition = (string?)axis?.Element(ChartNamespace + "tickLblPos")?.Attribute("val");
        return tickLabelPosition switch
        {
            "high" => true,
            "low" => false,
            _ => defaultRightSide
        };
    }

    private static bool ResolveSceneOrXmlValueAxisLabelsRightSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        if (sceneAxis is null)
        {
            return ResolveValueAxisLabelsRightSide(axis, defaultRightSide);
        }

        return sceneAxis.TickLabelPosition switch
        {
            "high" => true,
            "low" => false,
            _ => defaultRightSide
        };
    }

    private static ChartSeriesStroke? ReadChartAxisStroke(XDocument chartXml, string axisName, PptxTheme theme)
    {
        return ReadChartAxisStroke(chartXml
            .Descendants(ChartNamespace + axisName)
            .FirstOrDefault(), theme);
    }

    private static ChartSeriesStroke? ReadChartAxisStroke(XElement? axis, PptxTheme theme)
    {
        XElement? shapeProperties = axis?.Element(ChartNamespace + "spPr");
        if (shapeProperties?.Element(DrawingNamespace + "ln")?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new ChartSeriesStroke(new RgbColor(0, 0, 0), 0d, 0d);
        }

        return shapeProperties is not null &&
            TryReadLineWithAlpha(shapeProperties, theme, out RgbColor color, out double width, out double alpha)
                ? new ChartSeriesStroke(color, alpha, width)
                : null;
    }

    private static IReadOnlyList<ScatterSeries> ReadScatterSeries(XElement chartElement, bool readBubbleSize)
    {
        var series = new List<ScatterSeries>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            double[] xValues = ReadCachedNumbers(element, "xVal");
            double[] yValues = ReadCachedNumbers(element, "yVal");
            double[] bubbleSizes = readBubbleSize ? ReadCachedNumbers(element, "bubbleSize") : [];
            int count = Math.Min(xValues.Length, yValues.Length);
            if (count == 0)
            {
                continue;
            }

            var points = new ScatterPoint[count];
            for (int i = 0; i < count; i++)
            {
                double size = i < bubbleSizes.Length ? bubbleSizes[i] : 1d;
                points[i] = new ScatterPoint(xValues[i], yValues[i], size);
            }

            series.Add(new ScatterSeries(points));
        }

        return series;
    }

    private static double[] ReadCachedNumbers(XElement element, string containerName)
    {
        return element
            .Elements(ChartNamespace + containerName)
            .Descendants(ChartNamespace + "pt")
            .Select(point => (string?)point.Element(ChartNamespace + "v"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : double.NaN)
            .Where(value => !double.IsNaN(value))
            .ToArray();
    }

    private static RgbColor ChartPalette(IReadOnlyList<RgbColor>? chartPalette, PptxTheme? theme, int index)
    {
        if (chartPalette is { Count: > 0 })
        {
            return chartPalette[index % chartPalette.Count];
        }

        if (theme is not null && theme.TryResolveColor("accent" + (index % 6 + 1).ToString(CultureInfo.InvariantCulture), out RgbColor themeColor))
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

    private static void RenderBarChart(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, ChartPlotBox plotBox, IReadOnlyList<IReadOnlyList<double>> series, bool horizontalBars, string grouping, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool majorGridlines, bool minorGridlines, ChartGridlineStyle gridlineStyle, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents, ChartAxisUnits axisUnits, bool varyColors, double gapWidthPercent, double overlapPercent)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        RenderChartShapeStyle(graphics, plotX, plotY, plotWidth, plotHeight, plotAreaStyle);
        int categoryCount = Math.Max(1, series.Max(values => values.Count));
        bool stacked = string.Equals(grouping, "stacked", StringComparison.Ordinal) ||
            string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
        bool percentStacked = string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
        (double minValue, double maxValue) = (valueExtents.Min, valueExtents.Max);
        double valueRange = Math.Max(1d, maxValue - minValue);

        double zeroX = plotX + (-minValue) / valueRange * plotWidth;
        double zeroY = plotY + (-minValue) / valueRange * plotHeight;
        if (minorGridlines)
        {
            if (horizontalBars)
            {
                DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, major: false, gridlineStyle.Minor);
            }
            else
            {
                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, major: false, gridlineStyle.Minor);
            }
        }

        if (majorGridlines)
        {
            if (horizontalBars)
            {
                DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, major: true, gridlineStyle.Major);
            }
            else
            {
                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, major: true, gridlineStyle.Major);
            }
        }

        ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
        ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
        if (axesStyle.CategoryAxisVisible)
        {
            ChartSeriesStroke stroke = horizontalBars ? valueAxisStroke : categoryAxisStroke;
            if (stroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, stroke);
                graphics.StrokeLine(plotX, zeroY, plotX + plotWidth, zeroY);
            }
        }

        if (axesStyle.ValueAxisVisible)
        {
            ChartSeriesStroke stroke = horizontalBars ? categoryAxisStroke : valueAxisStroke;
            if (stroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, stroke);
                graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);
            }
        }

        if (!horizontalBars && axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke)
        {
            if (secondaryValueAxisStroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, secondaryValueAxisStroke);
                graphics.StrokeLine(plotX + plotWidth, plotY, plotX + plotWidth, plotY + plotHeight);
            }
        }

        if (horizontalBars)
        {
            if (stacked)
            {
                RenderStackedHorizontalBars(graphics, theme, chartPalette, plotY, plotWidth, plotHeight, series, categoryCount, valueRange, zeroX, percentStacked, seriesFills, pointFills, pointStrokes, varyColors, gapWidthPercent);
            }
            else
            {
                RenderClusteredHorizontalBars(graphics, theme, chartPalette, plotY, plotWidth, plotHeight, series, categoryCount, valueRange, zeroX, seriesFills, pointFills, pointStrokes, varyColors, gapWidthPercent, overlapPercent);
            }

            return;
        }

        if (stacked)
        {
            RenderStackedColumns(graphics, theme, chartPalette, plotX, plotWidth, plotHeight, series, categoryCount, valueRange, zeroY, percentStacked, seriesFills, pointFills, pointStrokes, varyColors, gapWidthPercent);
            return;
        }

        double categoryWidth = plotWidth / categoryCount;
        double barWidth = GetClusteredBarWidth(categoryWidth, series.Count, gapWidthPercent);
        double step = GetClusteredBarStep(barWidth, overlapPercent);
        double clusterWidth = barWidth + Math.Max(0, series.Count - 1) * step;
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryX = plotX + category * categoryWidth + (categoryWidth - clusterWidth) / 2d;
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = values[category];
                double barHeight = Math.Abs(value) / valueRange * plotHeight;
                ChartSeriesFill fill = ChartPointCategoryOrSeriesColor(theme, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills);
                double barX = categoryX + seriesIndex * step;
                double barY = value >= 0d ? zeroY : zeroY - barHeight;
                FillChartRectangle(graphics, barX, barY, barWidth, barHeight, fill);
                StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, barX, barY, barWidth, barHeight);
            }
        }
    }

    private static ChartLayout GetBarChartLayout(PptxDocument document, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(sceneChart, chartXml);
        ChartPlotBox plotBox = GetBarChartPlotBox(frame, chartXml, sceneChart, title, legend);
        return new ChartLayout(frame, plotBox, title, legend);
    }

    private static ChartPlotBox GetBarChartPlotBox(ChartFrameBox frame, XDocument chartXml, PptxSceneChart? sceneChart, string? title, ChartLegendLayout legend)
    {
        if (TryReadSceneOrXmlManualPlotBox(sceneChart, chartXml, frame, out ChartPlotBox manualPlotBox))
        {
            return manualPlotBox;
        }

        bool hasTitle = !string.IsNullOrWhiteSpace(title);
        bool hasLegend = legend.Visible && !legend.Overlay;
        if (!hasTitle && !hasLegend)
        {
            return new ChartPlotBox(
                frame.X + frame.Width * 0.112d,
                frame.Y + frame.Height * 0.035d,
                frame.Width * 0.86d,
                frame.Height * 0.885d);
        }

        return new ChartPlotBox(
            frame.X + frame.Width * 0.1d,
            frame.Y + frame.Height * 0.14d,
            frame.Width * 0.82d,
            frame.Height * 0.81d);
    }

    private static ChartValueExtents GetBarChartValueExtents(IReadOnlyList<IReadOnlyList<double>> series, string grouping)
    {
        int categoryCount = Math.Max(1, series.Max(values => values.Count));
        bool stacked = string.Equals(grouping, "stacked", StringComparison.Ordinal) ||
            string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
        bool percentStacked = string.Equals(grouping, "percentStacked", StringComparison.Ordinal);
        (double min, double max) = stacked
            ? GetStackedValueExtents(series, categoryCount, percentStacked)
            : GetClusteredValueExtents(series);
        return new ChartValueExtents(min, max);
    }

    private static void DrawHorizontalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, ChartValueExtents extents, double? explicitUnit, bool major, ChartSeriesStroke? gridlineStroke)
    {
        ChartSeriesStroke stroke = gridlineStroke ?? DefaultChartGridlineStroke(major);
        if (stroke.Alpha <= 0.001d)
        {
            return;
        }

        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        double range = Math.Max(1d, extents.Max - extents.Min);
        foreach (double value in GetChartAxisTickValues(extents, explicitUnit, includeEndpoints: false))
        {
            double y = plotY + plotHeight * (value - extents.Min) / range;
            graphics.StrokeLine(plotX, y, plotX + plotWidth, y);
        }

        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static void DrawVerticalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, ChartValueExtents extents, double? explicitUnit, bool major, ChartSeriesStroke? gridlineStroke)
    {
        ChartSeriesStroke stroke = gridlineStroke ?? DefaultChartGridlineStroke(major);
        if (stroke.Alpha <= 0.001d)
        {
            return;
        }

        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        double range = Math.Max(1d, extents.Max - extents.Min);
        foreach (double value in GetChartAxisTickValues(extents, explicitUnit, includeEndpoints: false))
        {
            double x = plotX + plotWidth * (value - extents.Min) / range;
            graphics.StrokeLine(x, plotY, x, plotY + plotHeight);
        }

        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static ChartSeriesStroke DefaultChartGridlineStroke(bool major)
    {
        return major
            ? new ChartSeriesStroke(new RgbColor(217, 217, 217), 1d, 0.5d)
            : new ChartSeriesStroke(new RgbColor(235, 235, 235), 1d, 0.25d);
    }

    private static ChartSeriesFill ChartSeriesColor(int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha = 1d)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(seriesIndex), defaultAlpha);
    }

    private static ChartSeriesFill ChartSeriesColor(PptxTheme theme, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha = 1d)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(null, theme, seriesIndex), defaultAlpha);
    }

    private static ChartSeriesFill ChartSeriesColor(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha = 1d)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(chartPalette, theme, seriesIndex), defaultAlpha);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null)
            ? new ChartSeriesFill(ChartPalette(categoryIndex), 1d)
            : ChartSeriesColor(seriesIndex, seriesFills);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(PptxTheme theme, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null)
            ? new ChartSeriesFill(ChartPalette(null, theme, categoryIndex), 1d)
            : ChartSeriesColor(theme, seriesIndex, seriesFills);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null)
            ? new ChartSeriesFill(ChartPalette(chartPalette, theme, categoryIndex), 1d)
            : ChartSeriesColor(theme, chartPalette, seriesIndex, seriesFills);
    }

    private static ChartSeriesFill ChartPointCategoryOrSeriesColor(int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills)
    {
        if (seriesIndex < pointFills.Count && pointFills[seriesIndex].TryGetValue(categoryIndex, out ChartSeriesFill pointFill))
        {
            return pointFill;
        }

        return ChartCategoryOrSeriesColor(seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills);
    }

    private static ChartSeriesFill ChartPointCategoryOrSeriesColor(PptxTheme theme, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills)
    {
        if (seriesIndex < pointFills.Count && pointFills[seriesIndex].TryGetValue(categoryIndex, out ChartSeriesFill pointFill))
        {
            return pointFill;
        }

        return ChartCategoryOrSeriesColor(theme, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills);
    }

    private static ChartSeriesFill ChartPointCategoryOrSeriesColor(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills)
    {
        if (seriesIndex < pointFills.Count && pointFills[seriesIndex].TryGetValue(categoryIndex, out ChartSeriesFill pointFill))
        {
            return pointFill;
        }

        return ChartCategoryOrSeriesColor(theme, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills);
    }

    private static void FillChartRectangle(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartSeriesFill fill)
    {
        if (fill.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
        }

        RgbColor fillColor = fill.BackgroundColor ?? fill.Color;
        graphics.SetFillRgb(fillColor.Red, fillColor.Green, fillColor.Blue);
        graphics.FillRectangle(x, y, width, height);
        if (fill.PatternPreset is not null)
        {
            StrokeChartPatternFill(graphics, x, y, width, height, fill);
        }

        if (fill.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static void StrokeChartPatternFill(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartSeriesFill fill)
    {
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        graphics.SaveState();
        graphics.ClipRectangle(x, y, width, height);
        graphics.SetStrokeRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
        string patternPreset = fill.PatternPreset ?? "pct50";
        if (TryReadPercentageChartPattern(patternPreset, out int densityPercent))
        {
            FillChartDotPattern(graphics, x, y, width, height, fill.Color, densityPercent);
            graphics.RestoreState();
            return;
        }

        graphics.SetLineWidth(IsDarkChartPattern(patternPreset) ? 1.0d : 0.5d);
        double spacing = IsDarkChartPattern(patternPreset) ? 4d : 5d;
        bool up = patternPreset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase);
        for (double offset = -height; offset <= width + height; offset += spacing)
        {
            if (up)
            {
                graphics.StrokeLine(x + offset, y, x + offset + height, y + height);
            }
            else
            {
                graphics.StrokeLine(x + offset, y + height, x + offset + height, y);
            }
        }

        graphics.RestoreState();
    }

    private static void FillChartDotPattern(PdfGraphicsBuilder graphics, double x, double y, double width, double height, RgbColor color, int densityPercent)
    {
        double spacing = 4d;
        double dotDiameter = densityPercent >= 60
            ? 2.5d
            : densityPercent >= 30
                ? 1.5d
                : 1.0d;
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        for (double dotY = y + spacing / 2d; dotY <= y + height; dotY += spacing)
        {
            for (double dotX = x + spacing / 2d; dotX <= x + width; dotX += spacing)
            {
                graphics.FillEllipse(dotX - dotDiameter / 2d, dotY - dotDiameter / 2d, dotDiameter, dotDiameter);
            }
        }
    }

    private static bool TryReadPercentageChartPattern(string patternPreset, out int densityPercent)
    {
        densityPercent = 0;
        if (!patternPreset.StartsWith("pct", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return int.TryParse(patternPreset.AsSpan(3), NumberStyles.Integer, CultureInfo.InvariantCulture, out densityPercent);
    }

    private static bool IsDarkChartPattern(string patternPreset)
    {
        return patternPreset.StartsWith("dk", StringComparison.OrdinalIgnoreCase);
    }

    private static (double Min, double Max) GetClusteredValueExtents(IReadOnlyList<IReadOnlyList<double>> series)
    {
        double maxValue = Math.Max(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Max());
        double minValue = Math.Min(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Min());
        return (minValue, maxValue);
    }

    private static (double Min, double Max) GetStackedValueExtents(IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, bool percentStacked)
    {
        if (percentStacked)
        {
            return (0d, 1d);
        }

        double minValue = 0d;
        double maxValue = 0d;
        for (int category = 0; category < categoryCount; category++)
        {
            double positive = 0d;
            double negative = 0d;
            foreach (IReadOnlyList<double> values in series)
            {
                if (category >= values.Count)
                {
                    continue;
                }

                if (values[category] >= 0d)
                {
                    positive += values[category];
                }
                else
                {
                    negative += values[category];
                }
            }

            maxValue = Math.Max(maxValue, positive);
            minValue = Math.Min(minValue, negative);
        }

        return (minValue, maxValue);
    }

    private static void RenderClusteredHorizontalBars(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, double valueRange, double zeroX, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent, double overlapPercent)
    {
        double categoryHeight = plotHeight / categoryCount;
        double barHeight = GetClusteredBarWidth(categoryHeight, series.Count, gapWidthPercent);
        double step = GetClusteredBarStep(barHeight, overlapPercent);
        double clusterHeight = barHeight + Math.Max(0, series.Count - 1) * step;
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryY = plotY + category * categoryHeight + (categoryHeight - clusterHeight) / 2d;
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = values[category];
                double barWidth = Math.Abs(value) / valueRange * plotWidth;
                ChartSeriesFill fill = ChartPointCategoryOrSeriesColor(theme, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills);
                double barX = value >= 0d ? zeroX : zeroX - barWidth;
                double barY = categoryY + seriesIndex * step;
                FillChartRectangle(graphics, barX, barY, barWidth, barHeight, fill);
                StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, barX, barY, barWidth, barHeight);
            }
        }
    }

    private static void RenderStackedColumns(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, double plotX, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, double valueRange, double zeroY, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent)
    {
        double categoryWidth = plotWidth / categoryCount;
        double barWidth = GetStackedBarWidth(categoryWidth, gapWidthPercent);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryX = plotX + category * categoryWidth + (categoryWidth - barWidth) / 2d;
            double positiveY = zeroY;
            double negativeY = zeroY;
            double positiveTotal = GetCategoryPositiveTotal(series, category, percentStacked);
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = NormalizeStackedValue(values[category], positiveTotal, percentStacked);
                double segmentHeight = Math.Abs(value) / valueRange * plotHeight;
                ChartSeriesFill fill = ChartPointCategoryOrSeriesColor(theme, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills);
                if (value >= 0d)
                {
                    FillChartRectangle(graphics, categoryX, positiveY, barWidth, segmentHeight, fill);
                    StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, categoryX, positiveY, barWidth, segmentHeight);
                    positiveY += segmentHeight;
                }
                else
                {
                    negativeY -= segmentHeight;
                    FillChartRectangle(graphics, categoryX, negativeY, barWidth, segmentHeight, fill);
                    StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, categoryX, negativeY, barWidth, segmentHeight);
                }
            }
        }
    }

    private static void RenderStackedHorizontalBars(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, double valueRange, double zeroX, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent)
    {
        double categoryHeight = plotHeight / categoryCount;
        double barHeight = GetStackedBarWidth(categoryHeight, gapWidthPercent);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryY = plotY + category * categoryHeight + (categoryHeight - barHeight) / 2d;
            double positiveX = zeroX;
            double negativeX = zeroX;
            double positiveTotal = GetCategoryPositiveTotal(series, category, percentStacked);
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = NormalizeStackedValue(values[category], positiveTotal, percentStacked);
                double segmentWidth = Math.Abs(value) / valueRange * plotWidth;
                ChartSeriesFill fill = ChartPointCategoryOrSeriesColor(theme, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills);
                if (value >= 0d)
                {
                    FillChartRectangle(graphics, positiveX, categoryY, segmentWidth, barHeight, fill);
                    StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, positiveX, categoryY, segmentWidth, barHeight);
                    positiveX += segmentWidth;
                }
                else
                {
                    negativeX -= segmentWidth;
                    FillChartRectangle(graphics, negativeX, categoryY, segmentWidth, barHeight, fill);
                    StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, negativeX, categoryY, segmentWidth, barHeight);
                }
            }
        }
    }

    private static double GetStackedBarWidth(double categoryBand, double gapWidthPercent)
    {
        return Math.Max(0.5d, categoryBand * 100d / (100d + Math.Max(0d, gapWidthPercent)));
    }

    private static double GetClusteredBarWidth(double categoryBand, int seriesCount, double gapWidthPercent)
    {
        int count = Math.Max(1, seriesCount);
        return Math.Max(0.5d, categoryBand * 100d / (100d * count + Math.Max(0d, gapWidthPercent)));
    }

    private static double GetClusteredBarStep(double barWidth, double overlapPercent)
    {
        return Math.Max(0d, barWidth * (1d - overlapPercent / 100d));
    }

    private static double GetCategoryPositiveTotal(IReadOnlyList<IReadOnlyList<double>> series, int category, bool percentStacked)
    {
        if (!percentStacked)
        {
            return 1d;
        }

        return Math.Max(1d, series
            .Where(values => category < values.Count)
            .Select(values => Math.Max(0d, values[category]))
            .Sum());
    }

    private static double NormalizeStackedValue(double value, double positiveTotal, bool percentStacked)
    {
        return percentStacked && value > 0d ? value / positiveTotal : value;
    }

    private static void RenderLineChart(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, IReadOnlyList<IReadOnlyList<double>> series, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, IReadOnlyList<bool> smoothSeries, bool majorGridlines, bool minorGridlines, ChartGridlineStyle gridlineStyle, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents, ChartAxisUnits axisUnits)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        RenderChartShapeStyle(graphics, plotX, plotY, plotWidth, plotHeight, plotAreaStyle);
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double maxValue = valueExtents.Max;
        double minValue = valueExtents.Min;
        double valueRange = Math.Max(1d, maxValue - minValue);

        if (minorGridlines)
        {
            DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, major: false, gridlineStyle.Minor);
        }

        if (majorGridlines)
        {
            DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, major: true, gridlineStyle.Major);
        }

        ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
        ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
        if (axesStyle.CategoryAxisVisible)
        {
            if (categoryAxisStroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, categoryAxisStroke);
                graphics.StrokeLine(plotX, plotY, plotX + plotWidth, plotY);
            }
        }

        if (axesStyle.ValueAxisVisible)
        {
            if (valueAxisStroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, valueAxisStroke);
                graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);
            }

            if (axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke)
            {
                if (secondaryValueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, secondaryValueAxisStroke);
                    graphics.StrokeLine(plotX + plotWidth, plotY, plotX + plotWidth, plotY + plotHeight);
                }
            }
        }

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            if (values.Count == 0)
            {
                continue;
            }

            ChartSeriesStroke stroke = ChartSeriesStrokeColor(seriesIndex, seriesStrokes, 1.5d);
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            var points = new List<(double X, double Y)>(values.Count);
            for (int i = 0; i < values.Count; i++)
            {
                double pointX = plotX + (pointCount == 1 ? plotWidth / 2d : plotWidth * i / (pointCount - 1));
                double pointY = plotY + (values[i] - minValue) / valueRange * plotHeight;
                points.Add((pointX, pointY));
            }

            if (IsSmoothSeries(seriesIndex, smoothSeries))
            {
                StrokeSmoothChartPath(graphics, points);
            }
            else
            {
                StrokeStraightChartPath(graphics, points);
            }

            foreach ((double pointX, double pointY) in points)
            {
                graphics.SetFillRgb(stroke.Color.Red, stroke.Color.Green, stroke.Color.Blue);
                DrawChartMarker(graphics, pointX, pointY, ChartMarker(seriesIndex, markerStyles), stroke.Color, stroke.Color);
            }

            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static ChartLayout GetLineChartLayout(PptxDocument document, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(sceneChart, chartXml);
        ChartPlotBox plotBox = GetLineChartPlotBox(frame, chartXml, sceneChart);
        return new ChartLayout(frame, plotBox, title, legend);
    }

    private static ChartPlotBox GetLineChartPlotBox(ChartFrameBox frame, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        if (TryReadSceneOrXmlManualPlotBox(sceneChart, chartXml, frame, out ChartPlotBox manualPlotBox))
        {
            return manualPlotBox;
        }

        return new ChartPlotBox(
            frame.X + frame.Width * 0.12d,
            frame.Y + frame.Height * 0.16d,
            frame.Width * 0.76d,
            frame.Height * 0.68d);
    }

    private static ChartValueExtents GetLineChartValueExtents(IReadOnlyList<IReadOnlyList<double>> series)
    {
        double maxValue = Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max());
        double minValue = Math.Min(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Min());
        return new ChartValueExtents(minValue, maxValue);
    }

    private static bool IsSmoothSeries(int seriesIndex, IReadOnlyList<bool> smoothSeries)
    {
        return seriesIndex < smoothSeries.Count && smoothSeries[seriesIndex];
    }

    private static void StrokeStraightChartPath(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points)
    {
        for (int i = 1; i < points.Count; i++)
        {
            graphics.StrokeLine(points[i - 1].X, points[i - 1].Y, points[i].X, points[i].Y);
        }
    }

    private static void StrokeSmoothChartPath(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points)
    {
        if (points.Count < 2)
        {
            return;
        }

        graphics.MoveTo(points[0].X, points[0].Y);
        for (int i = 0; i < points.Count - 1; i++)
        {
            (double X, double Y) p0 = i == 0 ? points[i] : points[i - 1];
            (double X, double Y) p1 = points[i];
            (double X, double Y) p2 = points[i + 1];
            (double X, double Y) p3 = i + 2 < points.Count ? points[i + 2] : points[i + 1];
            graphics.CurveTo(
                p1.X + (p2.X - p0.X) / 6d,
                p1.Y + (p2.Y - p0.Y) / 6d,
                p2.X - (p3.X - p1.X) / 6d,
                p2.Y - (p3.Y - p1.Y) / 6d,
                p2.X,
                p2.Y);
        }

        graphics.StrokeCurrentPath();
    }

    private static ChartMarkerStyle ChartMarker(int seriesIndex, IReadOnlyList<ChartMarkerStyle> markerStyles)
    {
        return seriesIndex < markerStyles.Count ? markerStyles[seriesIndex] : ChartMarkerStyle.Default;
    }

    private static void DrawChartMarker(PdfGraphicsBuilder graphics, double x, double y, ChartMarkerStyle marker, RgbColor defaultFill, RgbColor defaultStroke)
    {
        if (string.Equals(marker.Symbol, "none", StringComparison.Ordinal))
        {
            return;
        }

        double size = marker.Size;
        ChartSeriesFill fill = marker.Fill ?? new ChartSeriesFill(defaultFill, 1d);
        ChartSeriesStroke? stroke = marker.Stroke ?? (IsLineOnlyChartMarker(marker.Symbol) ? new ChartSeriesStroke(defaultStroke, 1d, Math.Max(0.75d, size * 0.16d)) : null);
        if (!IsLineOnlyChartMarker(marker.Symbol))
        {
            if (fill.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            switch (marker.Symbol)
            {
                case "dot":
                    graphics.FillEllipse(x - size / 4d, y - size / 4d, size / 2d, size / 2d);
                    break;
                case "square":
                    graphics.FillRectangle(x - size / 2d, y - size / 2d, size, size);
                    break;
                case "diamond":
                    graphics.FillPolygon([
                        (x, y + size / 2d),
                        (x + size / 2d, y),
                        (x, y - size / 2d),
                        (x - size / 2d, y)
                    ]);
                    break;
                case "triangle":
                    graphics.FillPolygon([
                        (x, y + size / 2d),
                        (x + size / 2d, y - size / 2d),
                        (x - size / 2d, y - size / 2d)
                    ]);
                    break;
                case "star":
                    graphics.FillPolygon(BuildChartStarMarker(x, y, size));
                    break;
                default:
                    graphics.FillEllipse(x - size / 2d, y - size / 2d, size, size);
                    break;
            }

            if (fill.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }

        if (stroke is not { } markerStroke)
        {
            return;
        }

        if (markerStroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, markerStroke.Alpha);
        }

        SetChartStroke(graphics, markerStroke);
        switch (marker.Symbol)
        {
            case "dash":
                graphics.StrokeLine(x - size / 2d, y, x + size / 2d, y);
                break;
            case "plus":
                graphics.StrokeLine(x - size / 2d, y, x + size / 2d, y);
                graphics.StrokeLine(x, y - size / 2d, x, y + size / 2d);
                break;
            case "x":
                graphics.StrokeLine(x - size / 2d, y - size / 2d, x + size / 2d, y + size / 2d);
                graphics.StrokeLine(x - size / 2d, y + size / 2d, x + size / 2d, y - size / 2d);
                break;
            case "square":
                graphics.StrokeRectangle(x - size / 2d, y - size / 2d, size, size);
                break;
            case "diamond":
                graphics.StrokePolygon([
                    (x, y + size / 2d),
                    (x + size / 2d, y),
                    (x, y - size / 2d),
                    (x - size / 2d, y)
                ]);
                break;
            case "triangle":
                graphics.StrokePolygon([
                    (x, y + size / 2d),
                    (x + size / 2d, y - size / 2d),
                    (x - size / 2d, y - size / 2d)
                ]);
                break;
            case "star":
                graphics.StrokePolygon(BuildChartStarMarker(x, y, size));
                break;
            default:
                graphics.StrokeEllipse(x - size / 2d, y - size / 2d, size, size);
                break;
        }

        if (markerStroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static bool IsLineOnlyChartMarker(string symbol)
    {
        return string.Equals(symbol, "plus", StringComparison.Ordinal) ||
            string.Equals(symbol, "x", StringComparison.Ordinal) ||
            string.Equals(symbol, "dash", StringComparison.Ordinal);
    }

    private static (double X, double Y)[] BuildChartStarMarker(double x, double y, double size)
    {
        var points = new List<(double X, double Y)>(10);
        double outer = size / 2d;
        double inner = outer * 0.42d;
        for (int i = 0; i < 10; i++)
        {
            double radius = i % 2 == 0 ? outer : inner;
            double angle = -Math.PI / 2d + i * Math.PI / 5d;
            points.Add((x + Math.Cos(angle) * radius, y + Math.Sin(angle) * radius));
        }

        return points.ToArray();
    }

    private static void StrokeChartPointRectangle(PdfGraphicsBuilder graphics, int seriesIndex, int categoryIndex, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, double x, double y, double width, double height)
    {
        if (seriesIndex >= pointStrokes.Count || !pointStrokes[seriesIndex].TryGetValue(categoryIndex, out ChartSeriesStroke stroke))
        {
            return;
        }

        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        graphics.StrokeRectangle(x, y, width, height);
        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static ChartSeriesStroke ChartSeriesStrokeColor(int seriesIndex, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, double defaultWidth)
    {
        return seriesIndex < seriesStrokes.Count && seriesStrokes[seriesIndex] is { } stroke
            ? stroke
            : new ChartSeriesStroke(ChartPalette(seriesIndex), 1d, defaultWidth);
    }

    private static void SetChartStroke(PdfGraphicsBuilder graphics, ChartSeriesStroke stroke)
    {
        graphics.SetStrokeRgb(stroke.Color.Red, stroke.Color.Green, stroke.Color.Blue);
        graphics.SetLineWidth(stroke.Width);
        if (stroke.DashPattern is { Count: > 0 })
        {
            graphics.SetLineDash(stroke.DashPattern);
        }
        else
        {
            graphics.ClearLineDash();
        }

        graphics.SetLineCap(stroke.Cap ?? 0);
        graphics.SetLineJoin(stroke.Join ?? 0);
    }

    private static void RenderAreaChart(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, bool stacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double plotX = x + width * 0.12d;
        double plotY = y + height * 0.16d;
        double plotWidth = width * 0.76d;
        double plotHeight = height * 0.68d;
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double maxValue = stacked ? GetMaxStackedLineValue(series, pointCount) : Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max());
        double minValue = Math.Min(0d, series.SelectMany(values => values).DefaultIfEmpty(0d).Min());
        double valueRange = Math.Max(1d, maxValue - minValue);
        double zeroY = plotY + (-minValue) / valueRange * plotHeight;

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        graphics.StrokeLine(plotX, zeroY, plotX + plotWidth, zeroY);
        graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);

        double[] lower = new double[pointCount];
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            if (values.Count == 0)
            {
                continue;
            }

            var upperPoints = new (double X, double Y)[pointCount];
            var lowerPoints = new (double X, double Y)[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                double pointX = plotX + (pointCount == 1 ? plotWidth / 2d : plotWidth * i / (pointCount - 1));
                double value = i < values.Count ? values[i] : 0d;
                double lowerValue = stacked ? lower[i] : 0d;
                double upperValue = stacked ? lower[i] + value : value;
                upperPoints[i] = (pointX, plotY + (upperValue - minValue) / valueRange * plotHeight);
                lowerPoints[i] = (pointX, plotY + (lowerValue - minValue) / valueRange * plotHeight);
                if (stacked)
                {
                    lower[i] = upperValue;
                }
            }

            var polygon = new (double X, double Y)[pointCount * 2];
            for (int i = 0; i < pointCount; i++)
            {
                polygon[i] = upperPoints[i];
                polygon[polygon.Length - i - 1] = lowerPoints[i];
            }

            ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills, 0.62d);
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            graphics.FillPolygon(polygon);
            graphics.RestoreState();
            ChartSeriesStroke stroke = ChartSeriesStrokeColor(seriesIndex, seriesStrokes, 1.2d);
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            for (int i = 1; i < upperPoints.Length; i++)
            {
                graphics.StrokeLine(upperPoints[i - 1].X, upperPoints[i - 1].Y, upperPoints[i].X, upperPoints[i].Y);
            }

            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static double GetMaxStackedLineValue(IReadOnlyList<IReadOnlyList<double>> series, int pointCount)
    {
        double maxValue = 1d;
        for (int i = 0; i < pointCount; i++)
        {
            double sum = 0d;
            foreach (IReadOnlyList<double> values in series)
            {
                if (i < values.Count)
                {
                    sum += Math.Max(0d, values[i]);
                }
            }

            maxValue = Math.Max(maxValue, sum);
        }

        return maxValue;
    }

    private static void RenderScatterChart(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<ScatterSeries> series, bool connectLines, bool bubble, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, IReadOnlyList<bool> smoothSeries)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double plotX = x + width * 0.12d;
        double plotY = y + height * 0.16d;
        double plotWidth = width * 0.76d;
        double plotHeight = height * 0.68d;
        double minX = series.SelectMany(item => item.Points).Min(point => point.X);
        double maxX = series.SelectMany(item => item.Points).Max(point => point.X);
        double minY = Math.Min(0d, series.SelectMany(item => item.Points).Min(point => point.Y));
        double maxY = Math.Max(1d, series.SelectMany(item => item.Points).Max(point => point.Y));
        double xRange = Math.Max(1d, maxX - minX);
        double yRange = Math.Max(1d, maxY - minY);
        double maxBubbleSize = Math.Max(1d, series.SelectMany(item => item.Points).Max(point => point.Size));

        graphics.SetStrokeRgb(90, 90, 90);
        graphics.SetLineWidth(0.75d);
        graphics.StrokeLine(plotX, plotY, plotX + plotWidth, plotY);
        graphics.StrokeLine(plotX, plotY, plotX, plotY + plotHeight);

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills);
            ChartSeriesStroke stroke = ChartSeriesStrokeColor(seriesIndex, seriesStrokes, 1.2d);
            if (fill.Alpha < 1d || stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            var points = new List<(double X, double Y)>(series[seriesIndex].Points.Count);
            foreach (ScatterPoint point in series[seriesIndex].Points)
            {
                double pointX = plotX + (point.X - minX) / xRange * plotWidth;
                double pointY = plotY + (point.Y - minY) / yRange * plotHeight;
                points.Add((pointX, pointY));
            }

            if (connectLines)
            {
                if (IsSmoothSeries(seriesIndex, smoothSeries))
                {
                    StrokeSmoothChartPath(graphics, points);
                }
                else
                {
                    StrokeStraightChartPath(graphics, points);
                }
            }

            foreach (ScatterPoint point in series[seriesIndex].Points)
            {
                double pointX = plotX + (point.X - minX) / xRange * plotWidth;
                double pointY = plotY + (point.Y - minY) / yRange * plotHeight;
                double radius = bubble ? 3d + Math.Sqrt(Math.Max(0d, point.Size) / maxBubbleSize) * 8d : 3d;
                if (bubble)
                {
                    graphics.FillEllipse(pointX - radius, pointY - radius, radius * 2d, radius * 2d);
                }
                else
                {
                    DrawChartMarker(graphics, pointX, pointY, ChartMarker(seriesIndex, markerStyles), fill.Color, stroke.Color);
                }

            }

            if (fill.Alpha < 1d || stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static void RenderRadarChart(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<IReadOnlyList<double>> series, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double centerX = x + width * 0.5d;
        double centerY = y + height * 0.52d;
        double radius = Math.Min(width, height) * 0.32d;
        int pointCount = Math.Max(3, series.Max(values => values.Count));
        double maxValue = Math.Max(1d, series.SelectMany(values => values).DefaultIfEmpty(1d).Max());

        graphics.SetStrokeRgb(180, 180, 180);
        graphics.SetLineWidth(0.5d);
        for (int i = 0; i < pointCount; i++)
        {
            double angle = -Math.PI / 2d + i * Math.PI * 2d / pointCount;
            graphics.StrokeLine(centerX, centerY, centerX + Math.Cos(angle) * radius, centerY + Math.Sin(angle) * radius);
        }

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            var points = new (double X, double Y)[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                double value = i < values.Count ? Math.Max(0d, values[i]) : 0d;
                double pointRadius = value / maxValue * radius;
                double angle = -Math.PI / 2d + i * Math.PI * 2d / pointCount;
                points[i] = (centerX + Math.Cos(angle) * pointRadius, centerY + Math.Sin(angle) * pointRadius);
            }

            ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills, series.Count == 1 ? 0.40d : 0.18d);
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            graphics.FillPolygon(points);
            graphics.RestoreState();
            ChartSeriesStroke stroke = ChartSeriesStrokeColor(seriesIndex, seriesStrokes, 1.2d);
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            for (int i = 0; i < points.Length; i++)
            {
                (double X, double Y) a = points[i];
                (double X, double Y) b = points[(i + 1) % points.Length];
                graphics.StrokeLine(a.X, a.Y, b.X, b.Y);
            }

            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static void RenderPieChart(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<double> values, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions)
    {
        double total = values.Where(value => value > 0d).Sum();
        if (total <= 0d)
        {
            return;
        }

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double radius = Math.Min(width, height) * 0.34d;
        double centerX = x + width * 0.46d;
        double centerY = y + height * 0.52d;
        double angle = -Math.PI / 2d;

        for (int i = 0; i < values.Count; i++)
        {
            double value = Math.Max(0d, values[i]);
            if (value <= 0d)
            {
                continue;
            }

            double sweep = value / total * Math.PI * 2d;
            double midpointAngle = angle + sweep / 2d;
            double explosionOffset = pointExplosions.TryGetValue(i, out double explosion) ? radius * explosion : 0d;
            double sliceCenterX = centerX + Math.Cos(midpointAngle) * explosionOffset;
            double sliceCenterY = centerY + Math.Sin(midpointAngle) * explosionOffset;
            int segments = Math.Max(2, (int)Math.Ceiling(Math.Abs(sweep) / (Math.PI / 18d)));
            var points = new (double X, double Y)[segments + 2];
            points[0] = (sliceCenterX, sliceCenterY);
            for (int segment = 0; segment <= segments; segment++)
            {
                double segmentAngle = angle + sweep * segment / segments;
                points[segment + 1] = (
                    sliceCenterX + Math.Cos(segmentAngle) * radius,
                    sliceCenterY + Math.Sin(segmentAngle) * radius);
            }

            ChartSeriesFill fill = pointFills.TryGetValue(i, out ChartSeriesFill explicitFill)
                ? explicitFill
                : new ChartSeriesFill(ChartPalette(i), 1d);
            if (fill.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            graphics.FillPolygon(points);
            if (fill.Alpha < 1d)
            {
                graphics.RestoreState();
            }

            if (pointStrokes.TryGetValue(i, out ChartSeriesStroke stroke))
            {
                if (stroke.Alpha < 1d)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(1d, stroke.Alpha);
                }

                SetChartStroke(graphics, stroke);
                graphics.StrokePolygon(points);
                if (stroke.Alpha < 1d)
                {
                    graphics.RestoreState();
                }
            }

            angle += sweep;
        }
    }

    private static void RenderDoughnutChart(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, IReadOnlyList<double> values, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions, double holeSize)
    {
        RenderPieChart(graphics, document, bounds, values, pointFills, pointStrokes, pointExplosions);

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double radius = Math.Min(width, height) * 0.34d;
        double centerX = x + width * 0.46d;
        double centerY = y + height * 0.52d;
        double innerRadius = radius * holeSize;
        graphics.SetFillRgb(255, 255, 255);
        graphics.FillEllipse(centerX - innerRadius, centerY - innerRadius, innerRadius * 2d, innerRadius * 2d);
    }

    private static void EmitChartDiagnostic(Action<OoxPdfDiagnostic>? diagnosticSink, string id, OoxPdfSeverity severity, string message, string? partName, int slideIndex, string fallback)
    {
        diagnosticSink?.Invoke(new OoxPdfDiagnostic(
            id,
            severity,
            message,
            partName,
            SlideIndex: slideIndex,
            Feature: "chart",
            Fallback: fallback));
    }

    private readonly record struct ScatterSeries(IReadOnlyList<ScatterPoint> Points);

    private readonly record struct ScatterPoint(double X, double Y, double Size);

    private readonly record struct ChartSeriesFill(RgbColor Color, double Alpha, string? PatternPreset = null, RgbColor? BackgroundColor = null);

    private readonly record struct ChartSeriesStroke(
        RgbColor Color,
        double Alpha,
        double Width,
        IReadOnlyList<double>? DashPattern = null,
        int? Cap = null,
        int? Join = null);

    private static ChartSeriesStroke ChartAxisDefaultStroke { get; } = new(new RgbColor(90, 90, 90), 1d, 0.75d);

    private readonly record struct ChartAxesStyle(ChartSeriesStroke? ValueAxis, ChartSeriesStroke? SecondaryValueAxis, ChartSeriesStroke? CategoryAxis, bool ValueAxisVisible, bool CategoryAxisVisible);

    private readonly record struct ChartGridlineStyle(ChartSeriesStroke? Major, ChartSeriesStroke? Minor)
    {
        public static ChartGridlineStyle Empty { get; } = new(null, null);
    }

    private readonly record struct ChartLayout(ChartFrameBox Frame, ChartPlotBox PlotBox, string? Title, ChartLegendLayout Legend);

    private readonly record struct ChartFrameBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartPlotBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartValueExtents(double Min, double Max);

    private readonly record struct ChartAxisUnits(double? MajorUnit, double? MinorUnit)
    {
        public static ChartAxisUnits Empty { get; } = new(null, null);
    }

    private readonly record struct ChartTextStyle(string? FontFamily, double FontSize, RgbColor Color, bool Bold, bool Italic);

    private readonly record struct ChartTextStyleOverride(string? FontFamily, double? FontSize, RgbColor? Color, bool? Bold, bool? Italic)
    {
        public static ChartTextStyleOverride Empty { get; } = new(null, null, null, null, null);
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> EmptyChartDataLabelOverrides { get; } = new Dictionary<int, ChartDataLabelOverride>();

    private readonly record struct ChartDataLabelOptions(bool ShowValue, bool ShowPercent, bool ShowCategoryName, bool ShowSeriesName, bool ShowLeaderLines, string CustomText, string Position, string Separator, string NumberFormat, ChartTextStyleOverride TextStyle, ChartShapeStyle ShapeStyle, IReadOnlyDictionary<int, ChartDataLabelOverride> Overrides, bool IsDefined)
    {
        public static ChartDataLabelOptions None { get; } = new(ShowValue: false, ShowPercent: false, ShowCategoryName: false, ShowSeriesName: false, ShowLeaderLines: false, CustomText: string.Empty, Position: string.Empty, Separator: string.Empty, NumberFormat: string.Empty, TextStyle: ChartTextStyleOverride.Empty, ShapeStyle: ChartShapeStyle.Empty, Overrides: EmptyChartDataLabelOverrides, IsDefined: false);

        public bool HasVisibleText => ShowValue || ShowPercent || ShowCategoryName || ShowSeriesName ||
            !string.IsNullOrWhiteSpace(CustomText) ||
            Overrides.Values.Any(label => label.ShowValue == true || label.ShowPercent == true || label.ShowCategoryName == true || label.ShowSeriesName == true || !string.IsNullOrWhiteSpace(label.CustomText));
    }

    private readonly record struct ChartDataLabelOverride(bool? ShowValue, bool? ShowPercent, bool? ShowCategoryName, bool? ShowSeriesName, bool? ShowLeaderLines, string CustomText, string Position, string Separator, string NumberFormat, ChartTextStyleOverride TextStyle, ChartShapeStyle ShapeStyle);

    private readonly record struct ChartLegendEntry(string Name, ChartSeriesFill? Fill, ChartSeriesStroke? Stroke);

    private readonly record struct ChartLegendLayout(string Position, bool Overlay, bool Visible)
    {
        public static ChartLegendLayout Hidden { get; } = new("r", Overlay: false, Visible: false);
    }

    private readonly record struct ChartShapeStyle(ChartSeriesFill? Fill, ChartSeriesStroke? Stroke)
    {
        public static ChartShapeStyle Empty { get; } = new(null, null);

        public bool IsEmpty => Fill is null && Stroke is null;
    }

    private readonly record struct ChartMarkerStyle(string Symbol, double Size, ChartSeriesFill? Fill, ChartSeriesStroke? Stroke)
    {
        public static ChartMarkerStyle Default { get; } = new("circle", 4d, null, null);
    }
}
