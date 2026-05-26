using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private const string ChartColorStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
    private const string ChartExternalDataPackageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
    private const string WorkbookRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    private const string SharedStringsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    private const string SpreadsheetStylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    private const string SpreadsheetTableRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    private const string SpreadsheetStylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    private const double ChartLineDefaultStrokeWidth = 2.25d;
    private static readonly XNamespace SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

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
        OoxPart? chartPart = context.Package.GetPart(chartPartName);
        if (resolvedChartXml is null)
        {
            if (chartPart is null)
            {
                EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart part was missing and was ignored.", chartPartName, context.SlideNumber, "Ignored");
                return;
            }

            using Stream chartStream = chartPart.OpenRead();
            resolvedChartXml = SafeXml.Load(chartStream);
            resolvedChartPalette = ReadChartPaletteColors(context.Package, chartPart, context.Theme);
        }

        ChartWorkbookData? chartWorkbook = chartPart is not null
            ? ReadEmbeddedChartWorkbookData(context.Package, chartPart, resolvedChartXml, sceneChart?.ExternalData ?? default)
            : null;

        if (TryRenderChart(graphics, context.Document, context.Theme, resolvedChartPalette, bounds.Value, resolvedChartXml, sceneChart, chartWorkbook, fonts))
        {
            fonts.AddRange(RenderManualChartAxisTitles(context.Document, context.Theme, graphics, bounds.Value, resolvedChartXml, sceneChart));
            fonts.AddRange(RenderChartTitle(context.Document, context.Theme, graphics, bounds.Value, resolvedChartXml, sceneChart));
            return;
        }

        if (chartWorkbook is not null)
        {
            HydrateChartReferenceCaches(chartWorkbook, resolvedChartXml);
            if (TryRenderChart(graphics, context.Document, context.Theme, resolvedChartPalette, bounds.Value, resolvedChartXml, sceneChart, workbook: null, fonts))
            {
                fonts.AddRange(RenderManualChartAxisTitles(context.Document, context.Theme, graphics, bounds.Value, resolvedChartXml, sceneChart));
                fonts.AddRange(RenderChartTitle(context.Document, context.Theme, graphics, bounds.Value, resolvedChartXml, sceneChart));
                return;
            }
        }

        EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Only bar, line, area, scatter, bubble, radar, pie, and doughnut charts with cached or embedded-workbook-backed numeric values are currently supported by the native chart renderer.", chartPartName, context.SlideNumber, "Ignored");
    }

    private static PptxSceneChartPlot? ReadSceneChartPlot(PptxSceneChart? chart, PptxSceneChartPlotKind kind, int index = 0)
    {
        return chart?
            .Plots
            .Where(plot => plot.PlotKind == kind && plot.KindIndex == index)
            .FirstOrDefault();
    }

    private static IReadOnlyList<XElement> ReadChartPlotElements(XDocument chartXml, string kind)
    {
        return chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()?
            .Elements(ChartNamespace + kind)
            .ToArray() ?? [];
    }

    private static IReadOnlyList<XElement> ReadChartPlotElements(XDocument chartXml, PptxSceneChartPlotKind kind)
    {
        string? elementName = kind switch
        {
            PptxSceneChartPlotKind.Area => "areaChart",
            PptxSceneChartPlotKind.Bar => "barChart",
            PptxSceneChartPlotKind.Bubble => "bubbleChart",
            PptxSceneChartPlotKind.Doughnut => "doughnutChart",
            PptxSceneChartPlotKind.Line => "lineChart",
            PptxSceneChartPlotKind.Pie => "pieChart",
            PptxSceneChartPlotKind.Radar => "radarChart",
            PptxSceneChartPlotKind.Scatter => "scatterChart",
            _ => null
        };
        return elementName is null ? [] : ReadChartPlotElements(chartXml, elementName);
    }

    private static PptxSceneChartAxis? ReadSceneChartAxis(PptxSceneChart? chart, PptxSceneChartPlot? plot, PptxSceneChartAxisKind kind)
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
                    candidate.AxisKind == kind);
                if (axis is not null)
                {
                    return axis;
                }
            }
        }

        return chart.Axes.FirstOrDefault(axis => axis.AxisKind == kind);
    }

    private static IReadOnlyList<PptxSceneChartPlot> ReadSceneChartPlots(PptxSceneChart? chart, PptxSceneChartPlotKind kind)
    {
        return chart?.Plots
            .Where(plot => plot.PlotKind == kind)
            .ToArray() ?? [];
    }

    private static IReadOnlyList<PptxSceneChartAxis> ReadSceneChartAxes(PptxSceneChart? chart, PptxSceneChartPlot? plot, PptxSceneChartAxisKind kind)
    {
        if (chart is null)
        {
            return [];
        }

        if (plot is not null && plot.AxisIds.Count != 0)
        {
            return plot.AxisIds
                .Select(axisId => chart.Axes.FirstOrDefault(candidate =>
                    string.Equals(candidate.Id, axisId, StringComparison.Ordinal) &&
                    candidate.AxisKind == kind))
                .Where(axis => axis is not null)
                .Select(axis => axis!)
                .ToArray();
        }

        return chart.Axes
            .Where(axis => axis.AxisKind == kind)
            .ToArray();
    }

    private static IReadOnlyList<ChartAxisSource> ReadSceneOrXmlChartValueAxesForPlot(
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? scenePlot,
        XDocument chartXml,
        XElement? chartElement)
    {
        XElement[] xmlAxes = chartElement is null
            ? chartXml.Descendants(ChartNamespace + "valAx").ToArray()
            : ReadChartValueAxesForChart(chartXml, chartElement).ToArray();
        IReadOnlyList<PptxSceneChartAxis> sceneAxes = ReadSceneChartAxes(sceneChart, scenePlot, PptxSceneChartAxisKind.Value);
        if (sceneAxes.Count == 0)
        {
            return xmlAxes
                .Select(axis => new ChartAxisSource(null, axis))
                .ToArray();
        }

        var sources = new List<ChartAxisSource>(sceneAxes.Count);
        foreach (PptxSceneChartAxis sceneAxis in sceneAxes)
        {
            XElement? xmlAxis = xmlAxes.FirstOrDefault(axis => string.Equals(ReadChartAxisId(axis), sceneAxis.Id, StringComparison.Ordinal));
            sources.Add(new ChartAxisSource(sceneAxis, xmlAxis));
        }

        return sources;
    }

    private static string ReadSceneOrXmlChartValue(string? sceneValue, XElement element, string childName, string defaultValue = "")
    {
        return !string.IsNullOrEmpty(sceneValue)
            ? sceneValue
            : (string?)element.Element(ChartNamespace + childName)?.Attribute("val") ?? defaultValue;
    }

    private static PptxSceneChartGrouping ReadSceneOrXmlChartGrouping(PptxSceneChartPlot? scenePlot, XElement plotElement, PptxSceneChartGrouping defaultGrouping)
    {
        return !string.IsNullOrEmpty(scenePlot?.Grouping)
            ? scenePlot.GroupingKind
            : PptxSceneBuilder.ParseChartGrouping((string?)plotElement.Element(ChartNamespace + "grouping")?.Attribute("val")) switch
            {
                PptxSceneChartGrouping.Unknown => defaultGrouping,
                PptxSceneChartGrouping parsed => parsed
            };
    }

    private static PptxSceneChartBarDirection ReadSceneOrXmlChartBarDirection(PptxSceneChartPlot? scenePlot, XElement plotElement)
    {
        return !string.IsNullOrEmpty(scenePlot?.BarDirection)
            ? scenePlot.BarDirectionKind
            : PptxSceneBuilder.ParseChartBarDirection((string?)plotElement.Element(ChartNamespace + "barDir")?.Attribute("val"));
    }

    private static PptxSceneChartScatterStyle ReadSceneOrXmlChartScatterStyle(PptxSceneChartPlot? scenePlot, XElement plotElement)
    {
        return !string.IsNullOrEmpty(scenePlot?.ScatterStyle)
            ? scenePlot.ScatterStyleKind
            : PptxSceneBuilder.ParseChartScatterStyle((string?)plotElement.Element(ChartNamespace + "scatterStyle")?.Attribute("val"));
    }

    private static PptxSceneChartRadarStyle ReadSceneOrXmlChartRadarStyle(PptxSceneChartPlot? scenePlot, XElement plotElement)
    {
        return !string.IsNullOrEmpty(scenePlot?.RadarStyle)
            ? scenePlot.RadarStyleKind
            : PptxSceneBuilder.ParseChartRadarStyle((string?)plotElement.Element(ChartNamespace + "radarStyle")?.Attribute("val"));
    }

    private static double ReadSceneDoughnutHoleSize(PptxSceneChartPlot? plot, XElement doughnutChart)
    {
        return plot?.HoleSize is { } rawHoleSize
            ? Math.Clamp(rawHoleSize / 100d, PptxChartMetricRules.DoughnutHoleMinimumRatio, PptxChartMetricRules.DoughnutHoleMaximumRatio)
            : ReadDoughnutHoleSize(doughnutChart);
    }

    private static double ReadSceneOrXmlFirstSliceAngle(PptxSceneChartPlot? plot, XElement chartElement)
    {
        if (plot?.FirstSliceAngle is { } sceneAngle)
        {
            return NormalizeAngleDegrees(sceneAngle);
        }

        string? value = (string?)chartElement.Element(ChartNamespace + "firstSliceAng")?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? NormalizeAngleDegrees(parsed)
            : 0d;
    }

    private static double NormalizeAngleDegrees(double angle)
    {
        double normalized = angle % 360d;
        return normalized < 0d ? normalized + 360d : normalized;
    }

    private static IReadOnlyList<IReadOnlyList<double>> ReadSceneOrXmlChartSeries(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook = null)
    {
        IReadOnlyList<double>[]? sceneSeries = plot?
            .Series
            .Select(series => series.Values.Count != 0 ? series.Values : ReadWorkbookNumericValues(workbook, series.DataSources.Values))
            .Where(values => values.Count != 0)
            .ToArray();
        return sceneSeries is { Length: > 0 }
            ? sceneSeries
            : ReadChartSeries(chartElement);
    }

    private static IReadOnlyList<ScatterSeries> ReadSceneOrXmlScatterSeries(PptxSceneChartPlot? plot, XElement chartElement, bool readBubbleSize, ChartWorkbookData? workbook = null)
    {
        if (plot is null)
        {
            return ReadScatterSeries(chartElement, readBubbleSize);
        }

        var series = new List<ScatterSeries>();
        foreach (PptxSceneChartSeries item in plot.Series)
        {
            IReadOnlyList<double> xValues = item.XValues.Count != 0 ? item.XValues : ReadWorkbookNumericValues(workbook, item.DataSources.XValues);
            IReadOnlyList<double> yValues = item.YValues.Count != 0 ? item.YValues : ReadWorkbookNumericValues(workbook, item.DataSources.YValues);
            IReadOnlyList<double> bubbleSizes = item.BubbleSizes.Count != 0 ? item.BubbleSizes : ReadWorkbookNumericValues(workbook, item.DataSources.BubbleSizes);
            int count = Math.Min(xValues.Count, yValues.Count);
            if (count == 0)
            {
                continue;
            }

            var points = new ScatterPoint[count];
            for (int i = 0; i < count; i++)
            {
                double size = readBubbleSize && i < bubbleSizes.Count ? bubbleSizes[i] : 1d;
                points[i] = new ScatterPoint(xValues[i], yValues[i], size);
            }

            series.Add(new ScatterSeries(points));
        }

        return series.Count != 0 ? series : ReadScatterSeries(chartElement, readBubbleSize);
    }

    private static IReadOnlyList<string> ReadSceneOrXmlCategoryLabels(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook = null)
    {
        IReadOnlyList<string>? categories = plot?
            .Series
            .Select(series => series.Categories.Count != 0 ? series.Categories : ReadWorkbookTextValues(workbook, series.DataSources.Categories))
            .FirstOrDefault(values => values.Count != 0);
        return categories is { Count: > 0 }
            ? categories
            : ReadChartCategoryLabels(chartElement);
    }

    private static IReadOnlyList<double> ReadWorkbookNumericValues(ChartWorkbookData? workbook, PptxSceneChartDataSource source)
    {
        return workbook?
            .ReadNumericRange(source.Formula)
            .Select(value => value.Value)
            .ToArray() ?? [];
    }

    private static IReadOnlyList<string> ReadWorkbookTextValues(ChartWorkbookData? workbook, PptxSceneChartDataSource source)
    {
        return workbook?
            .ReadTextRange(source.Formula)
            .Select(value => value.Text)
            .ToArray() ?? [];
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
                series.Marker.SymbolKind,
                series.Marker.Symbol,
                series.Marker.Size,
                series.Marker.Fill.HasFill ? new ChartSeriesFill(series.Marker.Fill.Color, series.Marker.Fill.Alpha) : null,
                ToChartSeriesStroke(series.Marker.Line)))
            .ToArray();
        bool hasSceneMarkerOwnership = plot?.Series.Any(series => series.Marker.IsDefined) == true;
        return markers is { Length: > 0 } && hasSceneMarkerOwnership
            ? markers
            : ReadChartMarkerStyles(chartElement, theme, plot?.PlotKind ?? PptxSceneBuilder.ParseChartPlotKind(chartElement.Name.LocalName));
    }

    private static IReadOnlyList<bool> ReadSceneOrXmlSmoothSeries(PptxSceneChartPlot? plot, XElement chartElement)
    {
        bool?[]? smooth = plot?.Series.Select(series => series.Smooth).ToArray();
        return smooth is { Length: > 0 } && smooth.Any(value => value.HasValue)
            ? smooth.Select(value => value ?? false).ToArray()
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
        if (series.Explosion is { } seriesExplosion)
        {
            double fraction = Math.Clamp(seriesExplosion / 100d, 0d, 1d);
            int pointCount = Math.Max(series.Values.Count, series.Categories.Count);
            for (int index = 0; index < pointCount; index++)
            {
                explosions[index] = fraction;
            }
        }

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
            ? new ChartSeriesStroke(line.Color, line.Alpha, line.Width, line.DashPattern, line.Cap, line.Join, line.Compound)
            : null;
    }

    private static ChartShapeStyle ToChartShapeStyle(PptxSceneChartShapeStyle style)
    {
        return new ChartShapeStyle(style.NoFill ? null : ToChartSeriesFill(style.Fill, style.PatternFill), style.NoFill || style.GradientFill is null ? null : ToGradientFill(style.GradientFill), ToChartSeriesStroke(style.Line));
    }

    private static bool TryRenderChart(PdfGraphicsBuilder graphics, PptxDocument document, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, ChartWorkbookData? workbook, List<PdfFontResource> fonts)
    {
        IReadOnlyList<XElement> barCharts = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Bar);
        XElement? barChart = barCharts.FirstOrDefault();
        if (barChart is not null)
        {
            PptxSceneChartPlot? barPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bar);
            IReadOnlyList<IReadOnlyList<double>> barSeries = ReadSceneOrXmlChartSeries(barPlot, barChart, workbook);
            if (barSeries.Count != 0)
            {
                PptxSceneChartBarDirection barDirection = ReadSceneOrXmlChartBarDirection(barPlot, barChart);
                bool horizontalBars = barDirection == PptxSceneChartBarDirection.Bar;
                PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(barPlot, barChart, PptxSceneChartGrouping.Clustered);
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(barPlot, barChart, theme);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, barPlot, chartXml, theme, barChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme);
                XElement? valueAxis = ReadChartValueAxisForChart(chartXml, barChart);
                PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, barPlot, PptxSceneChartAxisKind.Value);
                ChartGridlineStyle gridlineStyle = ReadSceneOrXmlChartGridlineStyle(valueSceneAxis, valueAxis, theme);
                bool percentStacked = IsPercentStackedChartGrouping(grouping);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueSceneAxis, valueAxis, GetBarChartValueExtents(barSeries, grouping), percentStacked);
                ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueSceneAxis, valueAxis), percentStacked);
                bool valueAxisReversed = ReadSceneOrXmlValueAxisReversed(valueSceneAxis, valueAxis);
                bool varyColors = barPlot?.VaryColors ?? ReadChartVaryColors(barChart);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills = ReadSceneOrXmlSeriesPointFills(barPlot, barChart, theme);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes = ReadSceneOrXmlSeriesPointStrokes(barPlot, barChart, theme);
                var legendEntries = new List<ChartLegendEntry>(BuildFillLegendEntries(theme, chartPalette, barPlot, barChart, seriesFills, workbook: workbook));
                ChartLayout chartLayout = GetBarChartLayout(document, theme, bounds, chartXml, sceneChart, barPlot, barChart, horizontalBars);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                double? valueAxisCrossingValue = ReadSceneOrXmlValueAxisCrossingValue(valueSceneAxis, valueAxis, valueExtents);
                bool valueAxisLabelsVisible = IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis);
                RenderBarChart(graphics, theme, chartPalette, chartLayout.PlotAreaBox, plotBox, barSeries, horizontalBars, grouping, seriesFills, pointFills, pointStrokes, ReadSceneOrXmlMajorGridlines(valueSceneAxis, valueAxis), ReadSceneOrXmlMinorGridlines(valueSceneAxis, valueAxis), gridlineStyle, axesStyle, plotAreaStyle, valueExtents, axisUnits, valueAxisCrossingValue, valueAxisReversed, valueAxisLabelsVisible, chartLayout.ManualPlotLayoutApplied, varyColors, barPlot?.GapWidth ?? ReadChartGapWidth(barChart), barPlot?.Overlap ?? ReadChartOverlap(barChart));
                XElement? secondaryValueAxis = null;
                PptxSceneChartAxis? secondaryValueSceneAxis = null;
                ChartValueExtents secondaryValueExtents = default;
                ChartAxisUnits secondaryAxisUnits = default;
                int seriesOffset = barSeries.Count;
                int barChartIndex = 1;
                foreach (XElement extraBarChart in barCharts.Skip(1))
                {
                    PptxSceneChartPlot? extraBarPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bar, barChartIndex);
                    IReadOnlyList<IReadOnlyList<double>> extraSeries = ReadSceneOrXmlChartSeries(extraBarPlot, extraBarChart, workbook);
                    if (extraSeries.Count == 0)
                    {
                        barChartIndex++;
                        continue;
                    }

                    PptxSceneChartGrouping extraGrouping = ReadSceneOrXmlChartGrouping(extraBarPlot, extraBarChart, PptxSceneChartGrouping.Clustered);
                    PptxSceneChartBarDirection extraBarDirection = ReadSceneOrXmlChartBarDirection(extraBarPlot, extraBarChart);
                    bool extraHorizontalBars = extraBarDirection == PptxSceneChartBarDirection.Bar;
                    XElement? extraValueAxis = ReadChartValueAxisForChart(chartXml, extraBarChart);
                    PptxSceneChartAxis? extraValueSceneAxis = ReadSceneChartAxis(sceneChart, extraBarPlot, PptxSceneChartAxisKind.Value);
                    bool extraPercentStacked = IsPercentStackedChartGrouping(extraGrouping);
                    ChartValueExtents extraValueExtents = ReadPercentStackedAwareValueAxisExtents(extraValueSceneAxis, extraValueAxis, GetBarChartValueExtents(extraSeries, extraGrouping), extraPercentStacked);
                    ChartAxisUnits extraAxisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(extraValueSceneAxis, extraValueAxis), extraPercentStacked);
                    IReadOnlyList<ChartSeriesFill?> extraSeriesFills = ReadSceneOrXmlSeriesFills(extraBarPlot, extraBarChart, theme);
                    IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> extraPointFills = ReadSceneOrXmlSeriesPointFills(extraBarPlot, extraBarChart, theme);
                    IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> extraPointStrokes = ReadSceneOrXmlSeriesPointStrokes(extraBarPlot, extraBarChart, theme);
                    if (!extraHorizontalBars && secondaryValueAxis is null && IsSceneOrXmlVisibleValueAxis(extraValueSceneAxis, extraValueAxis))
                    {
                        secondaryValueAxis = extraValueAxis;
                        secondaryValueSceneAxis = extraValueSceneAxis;
                        secondaryValueExtents = extraValueExtents;
                        secondaryAxisUnits = extraAxisUnits;
                    }

                    legendEntries.AddRange(BuildFillLegendEntries(theme, chartPalette, extraBarPlot, extraBarChart, extraSeriesFills, seriesOffset, workbook));
                    RenderBarChart(
                        graphics,
                        theme,
                        chartPalette,
                        chartLayout.PlotAreaBox,
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
                        ReadSceneOrXmlValueAxisCrossingValue(extraValueSceneAxis, extraValueAxis, extraValueExtents),
                        ReadSceneOrXmlValueAxisReversed(extraValueSceneAxis, extraValueAxis),
                        valueAxisLabelsVisible: false,
                        manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied,
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
                        ReadSceneOrXmlValueAxisReversed(extraValueSceneAxis, extraValueAxis),
                        ReadSceneOrXmlDataLabelOptions(extraBarPlot, extraBarChart, theme),
                        ReadSceneOrXmlSeriesDataLabelOptions(extraBarPlot, extraBarChart, theme),
                        ReadSceneOrXmlCategoryLabels(extraBarPlot, extraBarChart, workbook),
                        ReadSceneOrXmlChartSeriesNames(extraBarPlot, extraBarChart, workbook)));
                    seriesOffset += extraSeries.Count;
                    barChartIndex++;
                }

                int lineChartIndex = 0;
                foreach (XElement comboLineChart in ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Line))
                {
                    PptxSceneChartPlot? linePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Line, lineChartIndex);
                    IReadOnlyList<IReadOnlyList<double>> lineSeries = ReadSceneOrXmlChartSeries(linePlot, comboLineChart, workbook);
                    if (lineSeries.Count == 0)
                    {
                        lineChartIndex++;
                        continue;
                    }

                    XElement? lineValueAxis = ReadChartValueAxisForChart(chartXml, comboLineChart);
                    XElement? lineValueAxisForScale = lineValueAxis ?? valueAxis;
                    PptxSceneChartAxis? lineValueSceneAxis = ReadSceneChartAxis(sceneChart, linePlot, PptxSceneChartAxisKind.Value);
                    PptxSceneChartGrouping lineGrouping = ReadSceneOrXmlChartGrouping(linePlot, comboLineChart, PptxSceneChartGrouping.Standard);
                    bool lineStacked = IsStackedChartGrouping(lineGrouping);
                    bool linePercentStacked = IsPercentStackedChartGrouping(lineGrouping);
                    ChartValueExtents lineValueExtents = ReadPercentStackedAwareValueAxisExtents(lineValueSceneAxis, lineValueAxisForScale, GetLineChartValueExtents(lineSeries, lineStacked, linePercentStacked), linePercentStacked, useNearMaximumHeadroom: !linePercentStacked);
                    ChartAxisUnits lineAxisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(lineValueSceneAxis, lineValueAxisForScale), linePercentStacked);
                    bool lineValueAxisReversed = ReadSceneOrXmlValueAxisReversed(lineValueSceneAxis, lineValueAxisForScale);
                    IReadOnlyList<ChartSeriesStroke?> lineSeriesStrokes = ReadSceneOrXmlSeriesStrokes(linePlot, comboLineChart, theme);
                    IReadOnlyList<ChartMarkerStyle> lineMarkerStyles = ReadSceneOrXmlMarkerStyles(linePlot, comboLineChart, theme);
                    IReadOnlyList<bool> lineSmoothSeries = ReadSceneOrXmlSmoothSeries(linePlot, comboLineChart);
                    if (secondaryValueAxis is null && IsSceneOrXmlVisibleValueAxis(lineValueSceneAxis, lineValueAxis))
                    {
                        secondaryValueAxis = lineValueAxis;
                        secondaryValueSceneAxis = lineValueSceneAxis;
                        secondaryValueExtents = lineValueExtents;
                        secondaryAxisUnits = lineAxisUnits;
                    }

                    legendEntries.AddRange(BuildStrokeLegendEntries(theme, chartPalette, linePlot, comboLineChart, lineSeriesStrokes, reverseOrder: lineStacked, workbook: workbook));
                    RenderLineChart(
                        graphics,
                        theme,
                        chartPalette,
                        chartLayout.PlotAreaBox,
                        plotBox,
                        lineSeries,
                        lineStacked,
                        linePercentStacked,
                        lineSeriesStrokes,
                        lineMarkerStyles,
                        lineSmoothSeries,
                        majorGridlines: false,
                        minorGridlines: false,
                        ChartGridlineStyle.Empty,
                        axesStyle with { ValueAxisVisible = false, CategoryAxisVisible = false },
                        ChartShapeStyle.Empty,
                        lineValueExtents,
                        lineAxisUnits,
                        ReadSceneOrXmlValueAxisCrossingValue(lineValueSceneAxis, lineValueAxisForScale, lineValueExtents),
                        lineValueAxisReversed);
                    fonts.AddRange(RenderLineDataLabels(
                        theme,
                        graphics,
                        plotBox,
                        lineSeries,
                        lineValueExtents,
                        lineValueAxisReversed,
                        ReadSceneOrXmlDataLabelOptions(linePlot, comboLineChart, theme),
                        ReadSceneOrXmlSeriesDataLabelOptions(linePlot, comboLineChart, theme),
                        ReadSceneOrXmlCategoryLabels(linePlot, comboLineChart, workbook),
                        ReadSceneOrXmlChartSeriesNames(linePlot, comboLineChart, workbook)));
                    lineChartIndex++;
                }

                XElement? categoryAxis = ReadChartCategoryAxisForChart(chartXml, barChart);
                    PptxSceneChartAxis? categorySceneAxis = ReadSceneChartAxis(sceneChart, barPlot, PptxSceneChartAxisKind.Category);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categorySceneAxis, categoryAxis))
                {
                    double? categoryLabelAxisY = horizontalBars
                        ? null
                        : ChartValueToPlotCoordinate(valueExtents, valueAxisCrossingValue, plotBox.Y, plotBox.Height, valueAxisReversed);
                    fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categorySceneAxis, categoryAxis, ReadSceneOrXmlCategoryLabels(barPlot, barChart, workbook), horizontalBars, categoryLabelAxisY));
                }

                if (axesStyle.ValueAxisVisible)
                {
                    bool sameSideSecondaryValueAxis = !horizontalBars &&
                        secondaryValueAxis is not null &&
                        IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis) &&
                        GetValueAxisSideSlot(
                            valueSceneAxis,
                            valueAxis,
                            secondaryValueSceneAxis,
                            secondaryValueAxis,
                            defaultPrimaryRightSide: axesStyle.ValueAxisRightSide,
                            defaultSecondaryRightSide: ResolveSceneOrXmlValueAxisRightSide(secondaryValueSceneAxis, secondaryValueAxis, axesStyle.SecondaryValueAxisRightSide)) > 0;
                    if (valueAxisLabelsVisible)
                    {
                        fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis, valueSceneAxis, valueExtents, axisUnits, valueAxisReversed, horizontalBars, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied, useTextSizedWidth: sameSideSecondaryValueAxis, defaultNumberFormat: percentStacked ? "0%" : null));
                    }

                    if (!horizontalBars)
                    {
                        if (secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                        {
                            bool secondaryValueAxisRightSide = ResolveSceneOrXmlValueAxisRightSide(secondaryValueSceneAxis, secondaryValueAxis, axesStyle.SecondaryValueAxisRightSide);
                            int sideSlot = GetValueAxisSideSlot(valueSceneAxis, valueAxis, secondaryValueSceneAxis, secondaryValueAxis, defaultPrimaryRightSide: axesStyle.ValueAxisRightSide, defaultSecondaryRightSide: secondaryValueAxisRightSide);
                            fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, ReadSceneOrXmlValueAxisReversed(secondaryValueSceneAxis, secondaryValueAxis), horizontalBars: false, rightSide: secondaryValueAxisRightSide, axisSideSlot: sideSlot, useTextSizedWidth: sideSlot > 0));
                        }
                        else
                        {
                            fonts.AddRange(RenderSecondaryChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, GetBarChartValueExtents(barSeries, grouping)));
                        }
                    }
                }
                else if (!horizontalBars && secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                {
                    fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, ReadSceneOrXmlValueAxisReversed(secondaryValueSceneAxis, secondaryValueAxis), horizontalBars: false, rightSide: ResolveSceneOrXmlValueAxisRightSide(secondaryValueSceneAxis, secondaryValueAxis, axesStyle.SecondaryValueAxisRightSide)));
                }
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, legendEntries, chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, sceneChart, chartXml)));
                fonts.AddRange(RenderBarDataLabels(
                    theme,
                    graphics,
                    plotBox,
                    barSeries,
                    valueExtents,
                    horizontalBars,
                    valueAxisReversed,
                    ReadSceneOrXmlDataLabelOptions(barPlot, barChart, theme),
                    ReadSceneOrXmlSeriesDataLabelOptions(barPlot, barChart, theme),
                    ReadSceneOrXmlCategoryLabels(barPlot, barChart, workbook),
                    ReadSceneOrXmlChartSeriesNames(barPlot, barChart, workbook)));
                return true;
            }
        }

        XElement? lineChart = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Line).FirstOrDefault();
        if (lineChart is not null)
        {
            PptxSceneChartPlot? linePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Line);
            IReadOnlyList<IReadOnlyList<double>> lineSeries = ReadSceneOrXmlChartSeries(linePlot, lineChart, workbook);
            if (lineSeries.Count != 0)
            {
                PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(linePlot, lineChart, PptxSceneChartGrouping.Standard);
                bool stacked = IsStackedChartGrouping(grouping);
                bool percentStacked = IsPercentStackedChartGrouping(grouping);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(linePlot, lineChart, theme);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(linePlot, lineChart, theme);
                IReadOnlyList<bool> smoothSeries = ReadSceneOrXmlSmoothSeries(linePlot, lineChart);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, linePlot, chartXml, theme, lineChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme);
                XElement? valueAxis = ReadChartValueAxisForChart(chartXml, lineChart);
                XElement? valueAxisForScale = valueAxis ?? chartXml.Descendants(ChartNamespace + "valAx").FirstOrDefault();
                PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, linePlot, PptxSceneChartAxisKind.Value);
                ChartGridlineStyle gridlineStyle = ReadSceneOrXmlChartGridlineStyle(valueSceneAxis, valueAxisForScale, theme);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueSceneAxis, valueAxisForScale, GetLineChartValueExtents(lineSeries, stacked, percentStacked), percentStacked, useNearMaximumHeadroom: !percentStacked);
                ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueSceneAxis, valueAxisForScale), percentStacked);
                bool valueAxisReversed = ReadSceneOrXmlValueAxisReversed(valueSceneAxis, valueAxisForScale);
                ChartLayout chartLayout = GetLineChartLayout(document, theme, bounds, chartXml, sceneChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                RenderLineChart(graphics, theme, chartPalette, chartLayout.PlotAreaBox, plotBox, lineSeries, stacked, percentStacked, seriesStrokes, markerStyles, smoothSeries, ReadSceneOrXmlMajorGridlines(valueSceneAxis, valueAxisForScale), ReadSceneOrXmlMinorGridlines(valueSceneAxis, valueAxisForScale), gridlineStyle, axesStyle, plotAreaStyle, valueExtents, axisUnits, ReadSceneOrXmlValueAxisCrossingValue(valueSceneAxis, valueAxisForScale, valueExtents), valueAxisReversed);
                XElement? categoryAxis = ReadChartCategoryAxisForChart(chartXml, lineChart);
                PptxSceneChartAxis? categorySceneAxis = ReadSceneChartAxis(sceneChart, linePlot, PptxSceneChartAxisKind.Category);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categorySceneAxis, categoryAxis))
                {
                    fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categorySceneAxis, categoryAxis, ReadSceneOrXmlCategoryLabels(linePlot, lineChart, workbook), horizontalBars: false));
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis))
                {
                    fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis, valueSceneAxis, valueExtents, axisUnits, valueAxisReversed, horizontalBars: false, defaultNumberFormat: percentStacked ? "0%" : null));
                    fonts.AddRange(RenderSecondaryChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, GetLineChartValueExtents(lineSeries, stacked, percentStacked)));
                }
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildStrokeLegendEntries(theme, chartPalette, linePlot, lineChart, seriesStrokes, reverseOrder: stacked, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, sceneChart, chartXml)));
                    fonts.AddRange(RenderLineDataLabels(
                        theme,
                        graphics,
                        plotBox,
                        lineSeries,
                        valueExtents,
                        valueAxisReversed,
                        ReadSceneOrXmlDataLabelOptions(linePlot, lineChart, theme),
                        ReadSceneOrXmlSeriesDataLabelOptions(linePlot, lineChart, theme),
                        ReadSceneOrXmlCategoryLabels(linePlot, lineChart, workbook),
                    ReadSceneOrXmlChartSeriesNames(linePlot, lineChart, workbook)));
                return true;
            }
        }

        XElement? areaChart = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Area).FirstOrDefault();
        if (areaChart is not null)
        {
            PptxSceneChartPlot? areaPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Area);
            IReadOnlyList<IReadOnlyList<double>> areaSeries = ReadSceneOrXmlChartSeries(areaPlot, areaChart, workbook);
            if (areaSeries.Count != 0)
            {
                PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(areaPlot, areaChart, PptxSceneChartGrouping.Standard);
                bool stacked = IsStackedChartGrouping(grouping);
                bool percentStacked = IsPercentStackedChartGrouping(grouping);
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(areaPlot, areaChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(areaPlot, areaChart, theme);
                ChartLayout chartLayout = GetLineChartLayout(document, theme, bounds, chartXml, sceneChart);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                XElement? valueAxis = ReadChartValueAxisForChart(chartXml, areaChart);
                XElement? categoryAxis = ReadChartCategoryAxisForChart(chartXml, areaChart);
                PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, areaPlot, PptxSceneChartAxisKind.Value);
                PptxSceneChartAxis? categorySceneAxis = ReadSceneChartAxis(sceneChart, areaPlot, PptxSceneChartAxisKind.Category);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueSceneAxis, valueAxis, GetAreaChartValueExtents(areaSeries, stacked, percentStacked), percentStacked, useNearMaximumHeadroom: stacked && !percentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AreaChartStackedAxisNearMaximumHeadroomRatio);
                ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueSceneAxis, valueAxis), percentStacked);
                bool valueAxisReversed = ReadSceneOrXmlValueAxisReversed(valueSceneAxis, valueAxis);
                ChartGridlineStyle gridlineStyle = ReadSceneOrXmlChartGridlineStyle(valueSceneAxis, valueAxis, theme);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, areaPlot, chartXml, theme, areaChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                RenderAreaChart(
                    graphics,
                    theme,
                    chartPalette,
                    chartLayout.PlotAreaBox,
                    plotBox,
                    areaSeries,
                    stacked,
                    percentStacked,
                    seriesFills,
                    seriesStrokes,
                    ReadSceneOrXmlMajorGridlines(valueSceneAxis, valueAxis),
                    ReadSceneOrXmlMinorGridlines(valueSceneAxis, valueAxis),
                    gridlineStyle,
                    axesStyle,
                    plotAreaStyle,
                    valueExtents,
                    axisUnits,
                    ReadSceneOrXmlValueAxisCrossingValue(valueSceneAxis, valueAxis, valueExtents),
                    valueAxisReversed);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categorySceneAxis, categoryAxis))
                {
                    fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categorySceneAxis, categoryAxis, ReadSceneOrXmlCategoryLabels(areaPlot, areaChart, workbook), horizontalBars: false));
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis))
                {
                    fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis, valueSceneAxis, valueExtents, axisUnits, valueAxisReversed, horizontalBars: false, defaultNumberFormat: percentStacked ? "0%" : null));
                }

                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildFillLegendEntries(theme, chartPalette, areaPlot, areaChart, seriesFills, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, sceneChart, chartXml)));
                return true;
            }
        }

        XElement? scatterChart = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Scatter).FirstOrDefault();
        if (scatterChart is not null)
        {
            PptxSceneChartPlot? scatterPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Scatter);
            IReadOnlyList<ScatterSeries> scatterSeries = ReadSceneOrXmlScatterSeries(scatterPlot, scatterChart, readBubbleSize: false, workbook: workbook);
            if (scatterSeries.Count != 0)
            {
                PptxSceneChartScatterStyle scatterStyle = ReadSceneOrXmlChartScatterStyle(scatterPlot, scatterChart);
                bool connectLines = scatterStyle is PptxSceneChartScatterStyle.Line or PptxSceneChartScatterStyle.LineMarker;
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(scatterPlot, scatterChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(scatterPlot, scatterChart, theme);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(scatterPlot, scatterChart, theme);
                IReadOnlyList<bool> smoothSeries = ReadSceneOrXmlSmoothSeries(scatterPlot, scatterChart);
                ChartLayout chartLayout = GetLineChartLayout(document, theme, bounds, chartXml, sceneChart);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                IReadOnlyList<XElement> valueAxes = ReadChartValueAxesForChart(chartXml, scatterChart);
                XElement? xValueAxis = valueAxes.Count > 0 ? valueAxes[0] : null;
                XElement? yValueAxis = valueAxes.Count > 1 ? valueAxes[1] : ReadChartValueAxisForChart(chartXml, scatterChart);
                ChartValueExtents xExtents = ReadBubbleChartValueAxisExtents(xValueAxis, GetScatterXValueExtents(scatterSeries));
                ChartValueExtents yExtents = ReadBubbleChartValueAxisExtents(yValueAxis, GetScatterYValueExtents(scatterSeries));
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                RenderScatterChart(graphics, plotBox, scatterSeries, connectLines, bubble: false, seriesFills, seriesStrokes, markerStyles, smoothSeries, xExtents, yExtents);
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildStrokeLegendEntries(theme, chartPalette, scatterPlot, scatterChart, seriesStrokes, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, sceneChart, chartXml)));
                return true;
            }
        }

        XElement? bubbleChart = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Bubble).FirstOrDefault();
        if (bubbleChart is not null)
        {
            PptxSceneChartPlot? bubblePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bubble);
            IReadOnlyList<ScatterSeries> bubbleSeries = ReadSceneOrXmlScatterSeries(bubblePlot, bubbleChart, readBubbleSize: true, workbook: workbook);
            if (bubbleSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(bubblePlot, bubbleChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(bubblePlot, bubbleChart, theme);
                ChartLayout chartLayout = GetBubbleChartLayout(document, theme, bounds, chartXml, sceneChart, bubblePlot, bubbleChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                IReadOnlyList<XElement> valueAxes = ReadChartValueAxesForChart(chartXml, bubbleChart);
                XElement? xValueAxis = valueAxes.Count > 0 ? valueAxes[0] : null;
                XElement? yValueAxis = valueAxes.Count > 1 ? valueAxes[1] : ReadChartValueAxisForChart(chartXml, bubbleChart);
                ChartValueExtents xExtents = ReadBubbleChartValueAxisExtents(xValueAxis, GetBubbleXValueExtents(bubbleSeries));
                ChartValueExtents yExtents = ReadBubbleChartValueAxisExtents(yValueAxis, GetBubbleYValueExtents(bubbleSeries));
                ChartAxisUnits xAxisUnits = ResolveBubbleAxisUnits(ReadChartValueAxisUnits(xValueAxis), xExtents);
                ChartAxisUnits yAxisUnits = ResolveBubbleAxisUnits(ReadChartValueAxisUnits(yValueAxis), yExtents);
                ChartGridlineStyle gridlineStyle = ReadSceneOrXmlChartGridlineStyle(sceneAxis: null, yValueAxis, theme);
                DrawHorizontalChartGridlines(graphics, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, yExtents, yAxisUnits.MajorUnit, crossingValue: null, reversed: false, major: true, gridlineStyle.Major);
                RenderScatterChart(graphics, plotBox, bubbleSeries, connectLines: false, bubble: true, seriesFills, seriesStrokes, [], [], xExtents, yExtents);
                fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, xValueAxis, null, xExtents, xAxisUnits, valueAxisReversed: false, horizontalBars: true));
                fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, yValueAxis, null, yExtents, yAxisUnits, valueAxisReversed: false, horizontalBars: false));
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildFillLegendEntries(theme, chartPalette, bubblePlot, bubbleChart, seriesFills, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, sceneChart, chartXml)));
                return true;
            }
        }

        XElement? radarChart = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Radar).FirstOrDefault();
        if (radarChart is not null)
        {
            PptxSceneChartPlot? radarPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Radar);
            IReadOnlyList<IReadOnlyList<double>> radarSeries = ReadSceneOrXmlChartSeries(radarPlot, radarChart, workbook);
            if (radarSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(radarPlot, radarChart, theme);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(radarPlot, radarChart, theme);
                XElement? valueAxis = ReadChartValueAxisForChart(chartXml, radarChart);
                XElement? categoryAxis = ReadChartCategoryAxisForChart(chartXml, radarChart);
                PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, radarPlot, PptxSceneChartAxisKind.Value);
                PptxSceneChartAxis? categorySceneAxis = ReadSceneChartAxis(sceneChart, radarPlot, PptxSceneChartAxisKind.Category);
                ChartValueExtents valueExtents = ReadSceneOrXmlChartValueAxisExtents(valueSceneAxis, valueAxis, GetLineChartValueExtents(radarSeries));
                ChartAxisUnits axisUnits = ReadSceneOrXmlChartValueAxisUnits(valueSceneAxis, valueAxis);
                ChartPlotBox plotBox = GetPolarChartPlotBox(document, bounds, chartXml, sceneChart);
                ChartRadarLayout radarLayout = ResolveRadarLayout(plotBox, ReadSceneOrXmlChartRadarStyle(radarPlot, radarChart), radarSeries);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                RenderRadarChart(graphics, radarLayout, radarSeries, seriesFills, seriesStrokes, valueExtents, axisUnits);
                if (IsSceneOrXmlChartAxisLabelVisible(categorySceneAxis, categoryAxis))
                {
                    fonts.AddRange(RenderRadarCategoryLabels(theme, graphics, radarLayout, chartXml, sceneChart, categorySceneAxis, categoryAxis, ReadSceneOrXmlCategoryLabels(radarPlot, radarChart, workbook)));
                }

                if (IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis))
                {
                    fonts.AddRange(RenderRadarValueAxisLabels(theme, graphics, radarLayout, chartXml, sceneChart, valueAxis, valueSceneAxis, valueExtents, axisUnits));
                }

                return true;
            }
        }

        XElement? pieChart = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Pie).FirstOrDefault();
        if (pieChart is not null)
        {
            PptxSceneChartPlot? piePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Pie);
            IReadOnlyList<IReadOnlyList<double>> pieSeries = ReadSceneOrXmlChartSeries(piePlot, pieChart, workbook);
            if (pieSeries.Count != 0)
            {
                IReadOnlyDictionary<int, ChartSeriesFill> pointFills = ReadSceneOrXmlChartPointFills(piePlot, pieChart, theme);
                IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes = ReadSceneOrXmlChartPointStrokes(piePlot, pieChart, theme);
                IReadOnlyDictionary<int, double> pointExplosions = ReadSceneOrXmlChartPointExplosions(piePlot, pieChart);
                ChartFrameBox frame = GetChartFrameBox(document, bounds);
                ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, sceneChart, chartXml);
                ChartPlotBox plotBox = GetPolarChartPlotBox(document, bounds, chartXml, sceneChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Pie, plotBox, pointExplosions, legend);
                RenderPieChart(graphics, theme, chartPalette, polarLayout, pieSeries[0], pointFills, pointStrokes, pointExplosions, ReadSceneOrXmlFirstSliceAngle(piePlot, pieChart));
                fonts.AddRange(RenderPieDataLabels(theme, graphics, polarLayout, pieSeries[0], pointExplosions, 0d, ReadSceneOrXmlDataLabelOptions(piePlot, pieChart, theme)));
                fonts.AddRange(RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, chartPalette, piePlot, pieChart, pointFills, workbook), legend, ReadSceneOrXmlChartLegendTextStyle(theme, sceneChart, chartXml)));
                return true;
            }
        }

        XElement? doughnutChart = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Doughnut).FirstOrDefault();
        if (doughnutChart is not null)
        {
            PptxSceneChartPlot? doughnutPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Doughnut);
            IReadOnlyList<IReadOnlyList<double>> doughnutSeries = ReadSceneOrXmlChartSeries(doughnutPlot, doughnutChart, workbook);
            if (doughnutSeries.Count != 0)
            {
                IReadOnlyDictionary<int, ChartSeriesFill> pointFills = ReadSceneOrXmlChartPointFills(doughnutPlot, doughnutChart, theme);
                IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes = ReadSceneOrXmlChartPointStrokes(doughnutPlot, doughnutChart, theme);
                IReadOnlyDictionary<int, double> pointExplosions = ReadSceneOrXmlChartPointExplosions(doughnutPlot, doughnutChart);
                double holeSize = ReadSceneDoughnutHoleSize(doughnutPlot, doughnutChart);
                ChartFrameBox frame = GetChartFrameBox(document, bounds);
                ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, sceneChart, chartXml);
                ChartPlotBox plotBox = GetPolarChartPlotBox(document, bounds, chartXml, sceneChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme);
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Doughnut, plotBox, pointExplosions, legend);
                RenderDoughnutChart(graphics, theme, chartPalette, polarLayout, doughnutSeries[0], pointFills, pointStrokes, pointExplosions, holeSize, ReadSceneOrXmlFirstSliceAngle(doughnutPlot, doughnutChart));
                fonts.AddRange(RenderPieDataLabels(theme, graphics, polarLayout, doughnutSeries[0], pointExplosions, holeSize, ReadSceneOrXmlDataLabelOptions(doughnutPlot, doughnutChart, theme)));
                fonts.AddRange(RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, chartPalette, doughnutPlot, doughnutChart, pointFills, workbook), legend, ReadSceneOrXmlChartLegendTextStyle(theme, sceneChart, chartXml)));
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

    private static void HydrateChartReferenceCaches(ChartWorkbookData workbook, XDocument chartXml)
    {
        foreach (XElement reference in chartXml.Descendants(ChartNamespace + "numRef").ToArray())
        {
            if (reference.Descendants(ChartNamespace + "pt").Any())
            {
                continue;
            }

            ChartWorkbookNumericValue[] values = workbook.ReadNumericRange(reference.Element(ChartNamespace + "f")?.Value);
            if (values.Length == 0)
            {
                continue;
            }

            InsertChartReferenceCache(reference, new XElement(
                ChartNamespace + "numCache",
                new XElement(ChartNamespace + "formatCode", ReadChartReferenceNumberCacheFormatCode(reference)),
                BuildChartCachePointsFromWorkbookValues(values)));
        }

        foreach (XElement reference in chartXml.Descendants(ChartNamespace + "strRef").ToArray())
        {
            if (reference.Descendants(ChartNamespace + "pt").Any())
            {
                continue;
            }

            ChartWorkbookTextValue[] values = workbook.ReadTextRange(reference.Element(ChartNamespace + "f")?.Value);
            if (values.Length == 0)
            {
                continue;
            }

            InsertChartReferenceCache(reference, new XElement(
                ChartNamespace + "strCache",
                BuildChartCachePointsFromWorkbookValues(values)));
        }

        foreach (XElement reference in chartXml.Descendants(ChartNamespace + "multiLvlStrRef").ToArray())
        {
            if (reference.Descendants(ChartNamespace + "pt").Any())
            {
                continue;
            }

            ChartWorkbookTextValue[] values = workbook.ReadTextRange(reference.Element(ChartNamespace + "f")?.Value);
            if (values.Length == 0)
            {
                continue;
            }

            InsertChartReferenceCache(reference, new XElement(
                ChartNamespace + "multiLvlStrCache",
                new XElement(ChartNamespace + "ptCount", new XAttribute("val", (values.Max(value => value.Cell.Index) + 1).ToString(CultureInfo.InvariantCulture))),
                new XElement(ChartNamespace + "lvl", values.Select(value => BuildChartCachePoint(value.Cell.Index, value.Text)))));
        }
    }

    private static void InsertChartReferenceCache(XElement reference, XElement cache)
    {
        reference.Elements(cache.Name).Remove();
        XElement? formula = reference.Element(ChartNamespace + "f");
        if (formula is not null)
        {
            formula.AddAfterSelf(cache);
            return;
        }

        reference.AddFirst(cache);
    }

    private static string ReadChartReferenceNumberCacheFormatCode(XElement reference)
    {
        return (string?)reference
            .Element(ChartNamespace + "numCache")
            ?.Element(ChartNamespace + "formatCode") ?? "General";
    }

    private static object[] BuildChartCachePointsFromWorkbookValues(IReadOnlyList<ChartWorkbookNumericValue> values)
    {
        var elements = new object[values.Count + 1];
        int pointCount = values.Count == 0 ? 0 : values.Max(value => value.Cell.Index) + 1;
        elements[0] = new XElement(ChartNamespace + "ptCount", new XAttribute("val", pointCount.ToString(CultureInfo.InvariantCulture)));
        for (int index = 0; index < values.Count; index++)
        {
            ChartWorkbookNumericValue value = values[index];
            elements[index + 1] = BuildChartCachePoint(value.Cell.Index, value.Cell.Text);
        }

        return elements;
    }

    private static object[] BuildChartCachePointsFromWorkbookValues(IReadOnlyList<ChartWorkbookTextValue> values)
    {
        var elements = new object[values.Count + 1];
        int pointCount = values.Count == 0 ? 0 : values.Max(value => value.Cell.Index) + 1;
        elements[0] = new XElement(ChartNamespace + "ptCount", new XAttribute("val", pointCount.ToString(CultureInfo.InvariantCulture)));
        for (int index = 0; index < values.Count; index++)
        {
            ChartWorkbookTextValue value = values[index];
            elements[index + 1] = BuildChartCachePoint(value.Cell.Index, value.Text);
        }

        return elements;
    }

    private static XElement BuildChartCachePoint(int index, string value)
    {
        return new XElement(
            ChartNamespace + "pt",
            new XAttribute("idx", index.ToString(CultureInfo.InvariantCulture)),
            new XElement(ChartNamespace + "v", value));
    }

    private static ChartWorkbookData? ReadEmbeddedChartWorkbookData(
        OoxPackage package,
        OoxPart chartPart,
        XDocument chartXml,
        PptxSceneChartExternalData sceneExternalData)
    {
        string? targetPartName = sceneExternalData.IsDefined &&
            !string.IsNullOrWhiteSpace(sceneExternalData.TargetPartName)
            ? sceneExternalData.TargetPartName
            : ReadEmbeddedChartWorkbookTargetPartName(package, chartPart, chartXml);
        if (targetPartName is null)
        {
            return null;
        }

        OoxPart? workbookPackagePart = package.GetPart(targetPartName);
        if (workbookPackagePart is null)
        {
            return null;
        }

        using Stream stream = workbookPackagePart.OpenRead();
        OoxPackage workbookPackage = OoxPackage.Open(stream);
        return ReadWorkbookData(workbookPackage);
    }

    private static string? ReadEmbeddedChartWorkbookTargetPartName(OoxPackage package, OoxPart chartPart, XDocument chartXml)
    {
        string? relationshipId = (string?)chartXml
            .Descendants(ChartNamespace + "externalData")
            .FirstOrDefault()?
            .Attribute(RelationshipsNamespace + "id");
        if (relationshipId is null)
        {
            return null;
        }

        return package
            .GetRelationships(chartPart.Name)
            .FirstOrDefault(item =>
                item.Id == relationshipId &&
                !item.IsExternal &&
                item.Type == ChartExternalDataPackageRelationshipType &&
                item.ResolvedTarget is not null)
            ?.ResolvedTarget;
    }

    private static ChartWorkbookData? ReadWorkbookData(OoxPackage workbookPackage)
    {
        OoxPart? workbookPart = workbookPackage
            .GetRelationships("/")
            .Where(relationship => !relationship.IsExternal && relationship.Type == WorkbookRelationshipType && relationship.ResolvedTarget is not null)
            .Select(relationship => workbookPackage.GetPart(relationship.ResolvedTarget!))
            .FirstOrDefault(part => part is not null);
        workbookPart ??= workbookPackage.Parts.FirstOrDefault(part => part.ContentType == WorkbookContentType);
        if (workbookPart is null)
        {
            return null;
        }

        using Stream workbookStream = workbookPart.OpenRead();
        XDocument workbookXml = SafeXml.Load(workbookStream);
        ChartWorkbookSharedString[] sharedStrings = ReadWorkbookSharedStrings(workbookPackage, workbookPart);
        ChartWorkbookStyles styles = ReadWorkbookStyles(workbookPackage, workbookPart);
        IReadOnlyDictionary<string, OoxRelationship> workbookRelationships = workbookPackage
            .GetRelationships(workbookPart.Name)
            .Where(relationship => !relationship.IsExternal && relationship.ResolvedTarget is not null)
            .ToDictionary(relationship => relationship.Id, StringComparer.Ordinal);
        ChartWorkbookSheet[] workbookSheets = ReadWorkbookSheetRecords(workbookXml, workbookRelationships);
        ChartWorkbookDefinedName[] definedNameRecords = ReadWorkbookDefinedNameRecords(workbookXml, workbookSheets);
        IReadOnlyDictionary<string, string> definedNames = ReadWorkbookDefinedNames(definedNameRecords);
        ChartWorkbookCalculationProperties calculation = ReadWorkbookCalculationProperties(workbookXml);
        var tables = new Dictionary<string, ChartWorkbookTable>(StringComparer.OrdinalIgnoreCase);
        var sheets = new Dictionary<string, ChartWorksheetData>(StringComparer.OrdinalIgnoreCase);

        foreach (ChartWorkbookSheet sheet in workbookSheets)
        {
            if (string.IsNullOrWhiteSpace(sheet.Name) ||
                string.IsNullOrWhiteSpace(sheet.RelationshipId) ||
                !workbookRelationships.TryGetValue(sheet.RelationshipId, out OoxRelationship? relationship) ||
                relationship.Type != WorksheetRelationshipType ||
                relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxPart? worksheetPart = workbookPackage.GetPart(relationship.ResolvedTarget);
            if (worksheetPart is null)
            {
                continue;
            }

            sheets[sheet.Name] = ReadWorksheetData(worksheetPart, sharedStrings);
            foreach (ChartWorkbookTable table in ReadWorksheetTables(workbookPackage, worksheetPart, sheet.Name))
            {
                tables[table.Name] = table;
                if (!string.Equals(table.DisplayName, table.Name, StringComparison.OrdinalIgnoreCase))
                {
                    tables[table.DisplayName] = table;
                }
            }
        }

        return sheets.Count == 0 ? null : new ChartWorkbookData(sheets, ReadWorkbookDate1904(workbookXml), styles, definedNames, tables, calculation, definedNameRecords, workbookSheets);
    }

    private static bool ReadWorkbookDate1904(XDocument workbookXml)
    {
        return IsOoxmlTrue((string?)workbookXml
            .Root?
            .Element(SpreadsheetNamespace + "workbookPr")
            ?.Attribute("date1904"));
    }

    private static ChartWorkbookCalculationProperties ReadWorkbookCalculationProperties(XDocument workbookXml)
    {
        XElement? calculation = workbookXml.Root?.Element(SpreadsheetNamespace + "calcPr");
        return calculation is null
            ? default
            : new ChartWorkbookCalculationProperties(
                (string?)calculation.Attribute("calcMode") ?? string.Empty,
                (string?)calculation.Attribute("calcId") ?? string.Empty,
                IsOoxmlTrue((string?)calculation.Attribute("fullCalcOnLoad")),
                IsOoxmlTrue((string?)calculation.Attribute("forceFullCalc")));
    }

    private static ChartWorkbookSharedString[] ReadWorkbookSharedStrings(OoxPackage workbookPackage, OoxPart workbookPart)
    {
        OoxPart? sharedStringsPart = workbookPackage
            .GetRelationships(workbookPart.Name)
            .Where(relationship => !relationship.IsExternal && relationship.Type == SharedStringsRelationshipType && relationship.ResolvedTarget is not null)
            .Select(relationship => workbookPackage.GetPart(relationship.ResolvedTarget!))
            .FirstOrDefault(part => part is not null);
        sharedStringsPart ??= workbookPackage.Parts.FirstOrDefault(part => part.ContentType == SharedStringsContentType);
        if (sharedStringsPart is null)
        {
            return [];
        }

        using Stream stream = sharedStringsPart.OpenRead();
        XDocument document = SafeXml.Load(stream);
        return document
            .Descendants(SpreadsheetNamespace + "si")
            .Select(ReadWorkbookSharedString)
            .ToArray();
    }

    private static ChartWorkbookSharedString ReadWorkbookSharedString(XElement item)
    {
        XElement[] textElements = item.Descendants(SpreadsheetNamespace + "t").ToArray();
        int runCount = item.Elements(SpreadsheetNamespace + "r").Count();
        return new ChartWorkbookSharedString(
            string.Concat(textElements.Select(text => text.Value)),
            runCount,
            runCount != 0,
            item.Descendants(SpreadsheetNamespace + "rPh").Any(),
            textElements.Any(text => string.Equals((string?)text.Attribute(XNamespace.Xml + "space"), "preserve", StringComparison.Ordinal)));
    }

    private static ChartWorkbookSheet[] ReadWorkbookSheetRecords(
        XDocument workbookXml,
        IReadOnlyDictionary<string, OoxRelationship> workbookRelationships)
    {
        return workbookXml.Root?
            .Element(SpreadsheetNamespace + "sheets")
            ?.Elements(SpreadsheetNamespace + "sheet")
            .Select((sheet, index) =>
            {
                string relationshipId = ((string?)sheet.Attribute(RelationshipsNamespace + "id") ?? string.Empty).Trim();
                workbookRelationships.TryGetValue(relationshipId, out OoxRelationship? relationship);
                return new ChartWorkbookSheet(
                    ((string?)sheet.Attribute("name") ?? string.Empty).Trim(),
                    ((string?)sheet.Attribute("sheetId") ?? string.Empty).Trim(),
                    relationshipId,
                    ((string?)sheet.Attribute("state") ?? string.Empty).Trim(),
                    index,
                    relationship?.ResolvedTarget ?? string.Empty);
            })
            .ToArray() ?? [];
    }

    private static ChartWorkbookDefinedName[] ReadWorkbookDefinedNameRecords(XDocument workbookXml, IReadOnlyList<ChartWorkbookSheet> workbookSheets)
    {
        var definedNames = new List<ChartWorkbookDefinedName>();
        foreach (XElement definedName in workbookXml.Root?.Element(SpreadsheetNamespace + "definedNames")?.Elements(SpreadsheetNamespace + "definedName") ?? [])
        {
            string? name = (string?)definedName.Attribute("name");
            string formula = definedName.Value.Trim();
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(formula))
            {
                continue;
            }

            int? localSheetId = ReadSpreadsheetIntegerAttribute(definedName, "localSheetId");
            string sheetName = localSheetId is { } index && index >= 0 && index < workbookSheets.Count
                ? workbookSheets[index].Name
                : string.Empty;
            definedNames.Add(new ChartWorkbookDefinedName(name.Trim(), formula, localSheetId, sheetName));
        }

        return definedNames.ToArray();
    }

    private static IReadOnlyDictionary<string, string> ReadWorkbookDefinedNames(IReadOnlyList<ChartWorkbookDefinedName> definedNameRecords)
    {
        var definedNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (ChartWorkbookDefinedName definedName in definedNameRecords)
        {
            if (definedName.LocalSheetId is null)
            {
                definedNames[definedName.Name] = definedName.Formula;
            }
        }

        return definedNames;
    }

    private static ChartWorkbookStyles ReadWorkbookStyles(OoxPackage workbookPackage, OoxPart workbookPart)
    {
        OoxPart? stylesPart = workbookPackage
            .GetRelationships(workbookPart.Name)
            .Where(relationship => !relationship.IsExternal && relationship.Type == SpreadsheetStylesRelationshipType && relationship.ResolvedTarget is not null)
            .Select(relationship => workbookPackage.GetPart(relationship.ResolvedTarget!))
            .FirstOrDefault(part => part is not null);
        stylesPart ??= workbookPackage.Parts.FirstOrDefault(part => part.ContentType == SpreadsheetStylesContentType);
        if (stylesPart is null)
        {
            return ChartWorkbookStyles.Empty;
        }

        using Stream stream = stylesPart.OpenRead();
        XDocument document = SafeXml.Load(stream);
        var customNumberFormats = new Dictionary<int, string>();
        foreach (XElement numberFormat in document.Root?.Element(SpreadsheetNamespace + "numFmts")?.Elements(SpreadsheetNamespace + "numFmt") ?? [])
        {
            if (numberFormat.Attribute("numFmtId") is { } idAttribute &&
                int.TryParse(idAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) &&
                id >= 0)
            {
                customNumberFormats[id] = (string?)numberFormat.Attribute("formatCode") ?? string.Empty;
            }
        }

        var cellFormats = new List<ChartWorkbookCellFormat>();
        foreach (XElement format in document.Root?.Element(SpreadsheetNamespace + "cellXfs")?.Elements(SpreadsheetNamespace + "xf") ?? [])
        {
            int? numberFormatId = ReadSpreadsheetIntegerAttribute(format, "numFmtId");
            bool? applyNumberFormat = format.Attribute("applyNumberFormat") is { } applyAttribute
                ? IsOoxmlTrue(applyAttribute.Value)
                : null;
            string numberFormatCode = numberFormatId is { } id && customNumberFormats.TryGetValue(id, out string? customCode)
                ? customCode
                : string.Empty;
            cellFormats.Add(new ChartWorkbookCellFormat(numberFormatId, numberFormatCode, applyNumberFormat));
        }

        return new ChartWorkbookStyles(customNumberFormats, cellFormats);
    }

    private static ChartWorksheetData ReadWorksheetData(OoxPart worksheetPart, IReadOnlyList<ChartWorkbookSharedString> sharedStrings)
    {
        using Stream stream = worksheetPart.OpenRead();
        XDocument document = SafeXml.Load(stream);
        var cells = new Dictionary<string, ChartWorkbookCell>(StringComparer.OrdinalIgnoreCase);
        var hiddenRows = new HashSet<int>();
        var hiddenColumns = new HashSet<int>();
        foreach (XElement row in document.Descendants(SpreadsheetNamespace + "row"))
        {
            if (IsOoxmlTrue((string?)row.Attribute("hidden")) &&
                ReadSpreadsheetIntegerAttribute(row, "r") is { } rowIndex)
            {
                hiddenRows.Add(rowIndex);
            }
        }

        foreach (XElement column in document.Descendants(SpreadsheetNamespace + "col"))
        {
            if (!IsOoxmlTrue((string?)column.Attribute("hidden")) ||
                ReadSpreadsheetIntegerAttribute(column, "min") is not { } minColumn ||
                ReadSpreadsheetIntegerAttribute(column, "max") is not { } maxColumn)
            {
                continue;
            }

            for (int columnIndex = minColumn; columnIndex <= maxColumn; columnIndex++)
            {
                hiddenColumns.Add(columnIndex);
            }
        }

        foreach (XElement cell in document.Descendants(SpreadsheetNamespace + "c"))
        {
            string? reference = (string?)cell.Attribute("r");
            if (string.IsNullOrWhiteSpace(reference))
            {
                continue;
            }

            string? cellType = (string?)cell.Attribute("t");
            XElement? valueElement = cell.Element(SpreadsheetNamespace + "v");
            XElement? formula = cell.Element(SpreadsheetNamespace + "f");
            int? styleIndex = ReadSpreadsheetCellStyleIndex(cell);
            string? value = valueElement?.Value;
            string rawValue = value ?? string.Empty;
            bool hasValue = valueElement is not null;
            if (string.Equals(cellType, "inlineStr", StringComparison.Ordinal))
            {
                value = string.Concat(cell.Descendants(SpreadsheetNamespace + "t").Select(text => text.Value));
                hasValue = cell.Element(SpreadsheetNamespace + "is") is not null;
            }

            if (!hasValue && formula is null && styleIndex is null && string.IsNullOrEmpty(cellType))
            {
                continue;
            }

            value ??= string.Empty;
            int? sharedStringIndexValue = null;
            int sharedStringRunCount = 0;
            bool sharedStringHasRichText = false;
            bool sharedStringHasPhoneticText = false;
            bool sharedStringPreserveSpace = false;
            if (string.Equals(cellType, "s", StringComparison.Ordinal) &&
                int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sharedStringIndex))
            {
                sharedStringIndexValue = sharedStringIndex;
                if (sharedStringIndex >= 0 && sharedStringIndex < sharedStrings.Count)
                {
                    ChartWorkbookSharedString sharedString = sharedStrings[sharedStringIndex];
                    value = sharedString.Text;
                    sharedStringRunCount = sharedString.RunCount;
                    sharedStringHasRichText = sharedString.HasRichText;
                    sharedStringHasPhoneticText = sharedString.HasPhoneticText;
                    sharedStringPreserveSpace = sharedString.PreserveSpace;
                }
            }

            cells[reference] = new ChartWorkbookCell(
                value,
                rawValue,
                hasValue,
                valueElement is not null,
                sharedStringIndexValue,
                sharedStringRunCount,
                sharedStringHasRichText,
                sharedStringHasPhoneticText,
                sharedStringPreserveSpace,
                styleIndex,
                cellType ?? string.Empty,
                ReadWorkbookCellValueKind(cellType, value, hasValue),
                formula?.Value ?? string.Empty,
                (string?)formula?.Attribute("t") ?? string.Empty,
                ReadWorkbookFormulaAttributes(formula));
        }

        return new ChartWorksheetData(cells, hiddenRows, hiddenColumns);
    }

    private static IReadOnlyDictionary<string, string> ReadWorkbookFormulaAttributes(XElement? formula)
    {
        return formula?.Attributes()
            .ToDictionary(attribute => attribute.Name.LocalName, attribute => attribute.Value, StringComparer.Ordinal) ??
            new Dictionary<string, string>();
    }

    private static ChartWorkbookCellValueKind ReadWorkbookCellValueKind(string? cellType, string value, bool hasValue)
    {
        if (!hasValue)
        {
            return ChartWorkbookCellValueKind.Blank;
        }

        return cellType switch
        {
            null or "" => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _)
                ? ChartWorkbookCellValueKind.Number
                : ChartWorkbookCellValueKind.Other,
            "s" => ChartWorkbookCellValueKind.SharedString,
            "inlineStr" => ChartWorkbookCellValueKind.InlineString,
            "str" => ChartWorkbookCellValueKind.FormulaString,
            "b" => ChartWorkbookCellValueKind.Boolean,
            "e" => ChartWorkbookCellValueKind.Error,
            _ => ChartWorkbookCellValueKind.Other
        };
    }

    private static IReadOnlyList<ChartWorkbookTable> ReadWorksheetTables(OoxPackage workbookPackage, OoxPart worksheetPart, string sheetName)
    {
        var tables = new List<ChartWorkbookTable>();
        foreach (OoxRelationship relationship in workbookPackage.GetRelationships(worksheetPart.Name))
        {
            if (relationship.IsExternal ||
                relationship.Type != SpreadsheetTableRelationshipType ||
                relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxPart? tablePart = workbookPackage.GetPart(relationship.ResolvedTarget);
            if (tablePart is null)
            {
                continue;
            }

            using Stream stream = tablePart.OpenRead();
            XDocument tableXml = SafeXml.Load(stream);
            XElement? tableElement = tableXml.Root;
            if (tableElement is null)
            {
                continue;
            }

            string name = (string?)tableElement.Attribute("name") ?? string.Empty;
            string displayName = (string?)tableElement.Attribute("displayName") ?? name;
            string? reference = (string?)tableElement.Attribute("ref");
            if (string.IsNullOrWhiteSpace(name) ||
                string.IsNullOrWhiteSpace(reference) ||
                !TryParseTableReference(reference, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow))
            {
                continue;
            }

            ChartWorkbookTableColumn[] columns = tableElement
                .Element(SpreadsheetNamespace + "tableColumns")?
                .Elements(SpreadsheetNamespace + "tableColumn")
                .Select(column => new ChartWorkbookTableColumn(
                    ReadSpreadsheetIntegerAttribute(column, "id"),
                    (string?)column.Attribute("name") ?? string.Empty,
                    (string?)column.Attribute("totalsRowFunction") ?? string.Empty,
                    column.Element(SpreadsheetNamespace + "totalsRowFormula")?.Value ?? string.Empty,
                    column.Element(SpreadsheetNamespace + "calculatedColumnFormula")?.Value ?? string.Empty))
                .ToArray() ?? [];
            if (columns.Length == 0)
            {
                continue;
            }

            string[] columnNames = columns.Select(column => column.Name).ToArray();

            string autoFilterReference = (string?)tableElement.Element(SpreadsheetNamespace + "autoFilter")?.Attribute("ref") ?? string.Empty;
            int[] filterColumnIds = tableElement
                .Element(SpreadsheetNamespace + "autoFilter")?
                .Elements(SpreadsheetNamespace + "filterColumn")
                .Select(column => ReadSpreadsheetIntegerAttribute(column, "colId"))
                .Where(id => id is not null)
                .Select(id => id!.Value)
                .ToArray() ?? [];
            int headerRowCount = ReadSpreadsheetIntegerAttribute(tableElement, "headerRowCount") ?? 1;
            int totalsRowCount = ReadSpreadsheetIntegerAttribute(tableElement, "totalsRowCount") ?? 0;
            bool totalsRowShown = IsOoxmlTrue((string?)tableElement.Attribute("totalsRowShown"));

            tables.Add(new ChartWorkbookTable(
                name,
                displayName,
                sheetName,
                firstColumn,
                firstRow,
                lastColumn,
                lastRow,
                columnNames,
                columns,
                headerRowCount,
                totalsRowCount,
                totalsRowShown,
                autoFilterReference,
                filterColumnIds));
        }

        return tables;
    }

    private static bool TryParseTableReference(string reference, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow)
    {
        firstColumn = 0;
        firstRow = 0;
        lastColumn = 0;
        lastRow = 0;
        string[] references = reference.Split(':', 2, StringSplitOptions.TrimEntries);
        if (!TryParseSpreadsheetCellReference(references[0], out firstColumn, out firstRow))
        {
            return false;
        }

        if (references.Length == 1)
        {
            lastColumn = firstColumn;
            lastRow = firstRow;
            return true;
        }

        return TryParseSpreadsheetCellReference(references[1], out lastColumn, out lastRow);
    }

    private static bool TryParseSpreadsheetCellReference(string reference, out int column, out int row)
    {
        column = 0;
        row = 0;
        string normalized = reference.Replace("$", string.Empty, StringComparison.Ordinal).Trim();
        int index = 0;
        while (index < normalized.Length && char.IsAsciiLetter(normalized[index]))
        {
            column = (column * 26) + (char.ToUpperInvariant(normalized[index]) - 'A' + 1);
            index++;
        }

        if (column <= 0 || index == normalized.Length)
        {
            return false;
        }

        return int.TryParse(normalized[index..], NumberStyles.Integer, CultureInfo.InvariantCulture, out row) && row > 0;
    }

    private static int? ReadSpreadsheetCellStyleIndex(XElement cell)
    {
        return ReadSpreadsheetIntegerAttribute(cell, "s");
    }

    private static int? ReadSpreadsheetIntegerAttribute(XElement element, string attributeName)
    {
        string? value = (string?)element.Attribute(attributeName);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
            ? parsed
            : null;
    }

    private readonly record struct ChartWorkbookCell(
        string Text,
        string RawValue,
        bool HasValue,
        bool HasValueElement,
        int? SharedStringIndex,
        int SharedStringRunCount,
        bool SharedStringHasRichText,
        bool SharedStringHasPhoneticText,
        bool SharedStringPreserveSpace,
        int? StyleIndex,
        string CellType,
        ChartWorkbookCellValueKind ValueKind,
        string Formula,
        string FormulaType,
        IReadOnlyDictionary<string, string> FormulaAttributes);

    private readonly record struct ChartWorkbookSharedString(
        string Text,
        int RunCount,
        bool HasRichText,
        bool HasPhoneticText,
        bool PreserveSpace);

    private enum ChartWorkbookCellValueKind
    {
        Blank,
        Number,
        SharedString,
        InlineString,
        FormulaString,
        Boolean,
        Error,
        Other
    }

    private sealed class ChartWorksheetData
    {
        public ChartWorksheetData(
            Dictionary<string, ChartWorkbookCell> cells,
            IReadOnlySet<int> hiddenRows,
            IReadOnlySet<int> hiddenColumns)
        {
            Cells = cells;
            HiddenRows = hiddenRows;
            HiddenColumns = hiddenColumns;
        }

        public Dictionary<string, ChartWorkbookCell> Cells { get; }

        public IReadOnlySet<int> HiddenRows { get; }

        public IReadOnlySet<int> HiddenColumns { get; }
    }

    private readonly record struct ChartWorkbookCellFormat(
        int? NumberFormatId,
        string NumberFormatCode,
        bool? ApplyNumberFormat);

    private readonly record struct ChartWorkbookTable(
        string Name,
        string DisplayName,
        string SheetName,
        int FirstColumn,
        int FirstRow,
        int LastColumn,
        int LastRow,
        IReadOnlyList<string> ColumnNames,
        IReadOnlyList<ChartWorkbookTableColumn> Columns,
        int HeaderRowCount,
        int TotalsRowCount,
        bool TotalsRowShown,
        string AutoFilterReference,
        IReadOnlyList<int> FilterColumnIds);

    private readonly record struct ChartWorkbookTableColumn(
        int? Id,
        string Name,
        string TotalsRowFunction,
        string TotalsRowFormula,
        string CalculatedColumnFormula);

    private readonly record struct ChartWorkbookCalculationProperties(
        string CalculationMode,
        string CalculationId,
        bool FullCalculationOnLoad,
        bool ForceFullCalculation);

    private readonly record struct ChartWorkbookDefinedName(
        string Name,
        string Formula,
        int? LocalSheetId,
        string SheetName);

    private readonly record struct ChartWorkbookSheet(
        string Name,
        string SheetId,
        string RelationshipId,
        string State,
        int Index,
        string TargetPartName);

    private sealed class ChartWorkbookStyles
    {
        public ChartWorkbookStyles(
            IReadOnlyDictionary<int, string> customNumberFormats,
            IReadOnlyList<ChartWorkbookCellFormat> cellFormats)
        {
            CustomNumberFormats = customNumberFormats;
            CellFormats = cellFormats;
        }

        public static ChartWorkbookStyles Empty { get; } = new(new Dictionary<int, string>(), []);

        public IReadOnlyDictionary<int, string> CustomNumberFormats { get; }

        public IReadOnlyList<ChartWorkbookCellFormat> CellFormats { get; }

        public ChartWorkbookCellFormat ResolveCellFormat(int? styleIndex)
        {
            return styleIndex is { } index && index >= 0 && index < CellFormats.Count
                ? CellFormats[index]
                : default;
        }
    }

    private readonly record struct ChartWorkbookRangeCell(
        int Index,
        string Reference,
        string SheetName,
        string SourceFormula,
        string ResolvedFormula,
        ChartWorkbookRangeSourceKind SourceKind,
        string DefinedName,
        string TableName,
        string TableColumnName,
        int? TableColumnId,
        string Text,
        string RawValue,
        bool HasCell,
        bool HasValue,
        bool HasValueElement,
        int? SharedStringIndex,
        int SharedStringRunCount,
        bool SharedStringHasRichText,
        bool SharedStringHasPhoneticText,
        bool SharedStringPreserveSpace,
        int? StyleIndex,
        string CellType,
        ChartWorkbookCellValueKind ValueKind,
        string Formula,
        string FormulaType,
        IReadOnlyDictionary<string, string> FormulaAttributes,
        int? StyleNumberFormatId,
        string StyleNumberFormatCode,
        bool? StyleAppliesNumberFormat,
        bool RowHidden,
        bool ColumnHidden);

    private enum ChartWorkbookRangeSourceKind
    {
        Unknown,
        DirectRange,
        DefinedName,
        StructuredReference,
        DefinedNameStructuredReference
    }

    private readonly record struct ChartWorkbookRangeResolution(
        string SourceFormula,
        string ResolvedFormula,
        ChartWorkbookRangeSourceKind SourceKind,
        string DefinedName,
        string TableName,
        string TableColumnName,
        int? TableColumnId);

    private readonly record struct ChartWorkbookNumericValue(
        ChartWorkbookRangeCell Cell,
        double Value);

    private readonly record struct ChartWorkbookTextValue(
        ChartWorkbookRangeCell Cell,
        string Text);

    private sealed class ChartWorkbookData
    {
        private readonly IReadOnlyDictionary<string, ChartWorksheetData> sheets;
        private readonly ChartWorkbookStyles styles;
        private readonly IReadOnlyDictionary<string, string> definedNames;
        private readonly IReadOnlyList<ChartWorkbookDefinedName> definedNameRecords;
        private readonly IReadOnlyList<ChartWorkbookSheet> sheetRecords;
        private readonly IReadOnlyDictionary<string, ChartWorkbookTable> tables;
        private readonly ChartWorkbookCalculationProperties calculation;

        public ChartWorkbookData(IReadOnlyDictionary<string, Dictionary<string, string>> sheets)
            : this(sheets, date1904: false)
        {
        }

        public ChartWorkbookData(IReadOnlyDictionary<string, Dictionary<string, string>> sheets, bool date1904)
        {
            this.sheets = ConvertWorkbookSheets(sheets);
            styles = ChartWorkbookStyles.Empty;
            definedNames = new Dictionary<string, string>();
            definedNameRecords = [];
            sheetRecords = ConvertWorkbookSheetRecords(this.sheets);
            tables = new Dictionary<string, ChartWorkbookTable>();
            calculation = default;
            Date1904 = date1904;
        }

        public ChartWorkbookData(IReadOnlyDictionary<string, Dictionary<string, ChartWorkbookCell>> sheets, bool date1904)
            : this(sheets, date1904, ChartWorkbookStyles.Empty)
        {
        }

        public ChartWorkbookData(IReadOnlyDictionary<string, Dictionary<string, ChartWorkbookCell>> sheets, bool date1904, ChartWorkbookStyles styles)
            : this(ConvertWorkbookSheets(sheets), date1904, styles, new Dictionary<string, string>())
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames)
            : this(sheets, date1904, styles, definedNames, new Dictionary<string, ChartWorkbookTable>())
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames,
            IReadOnlyDictionary<string, ChartWorkbookTable> tables)
            : this(sheets, date1904, styles, definedNames, tables, default)
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames,
            IReadOnlyDictionary<string, ChartWorkbookTable> tables,
            ChartWorkbookCalculationProperties calculation)
            : this(sheets, date1904, styles, definedNames, tables, calculation, [])
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames,
            IReadOnlyDictionary<string, ChartWorkbookTable> tables,
            ChartWorkbookCalculationProperties calculation,
            IReadOnlyList<ChartWorkbookDefinedName> definedNameRecords)
            : this(sheets, date1904, styles, definedNames, tables, calculation, definedNameRecords, ConvertWorkbookSheetRecords(sheets))
        {
        }

        public ChartWorkbookData(
            IReadOnlyDictionary<string, ChartWorksheetData> sheets,
            bool date1904,
            ChartWorkbookStyles styles,
            IReadOnlyDictionary<string, string> definedNames,
            IReadOnlyDictionary<string, ChartWorkbookTable> tables,
            ChartWorkbookCalculationProperties calculation,
            IReadOnlyList<ChartWorkbookDefinedName> definedNameRecords,
            IReadOnlyList<ChartWorkbookSheet> sheetRecords)
        {
            this.sheets = sheets;
            this.styles = styles;
            this.definedNames = definedNames;
            this.definedNameRecords = definedNameRecords;
            this.sheetRecords = sheetRecords;
            this.tables = tables;
            this.calculation = calculation;
            Date1904 = date1904;
        }

        public bool Date1904 { get; }

        public IReadOnlyDictionary<string, string> DefinedNames => definedNames;

        public IReadOnlyList<ChartWorkbookDefinedName> DefinedNameRecords => definedNameRecords;

        public IReadOnlyList<ChartWorkbookSheet> Sheets => sheetRecords;

        public IReadOnlyDictionary<string, ChartWorkbookTable> Tables => tables;

        public ChartWorkbookCalculationProperties Calculation => calculation;

        public string[] ReadRange(string? formula)
        {
            return ReadRangeCells(formula)
                .Where(cell => cell.HasCell)
                .Select(cell => cell.Text)
                .ToArray();
        }

        public ChartWorkbookNumericValue[] ReadNumericRange(string? formula)
        {
            return ReadRangeCells(formula)
                .Where(cell => double.TryParse(cell.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                .Select(cell =>
                {
                    double.TryParse(cell.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value);
                    return new ChartWorkbookNumericValue(cell, value);
                })
                .ToArray();
        }

        public ChartWorkbookTextValue[] ReadTextRange(string? formula)
        {
            return ReadRangeCells(formula)
                .Where(cell => !string.IsNullOrWhiteSpace(cell.Text))
                .Select(cell => new ChartWorkbookTextValue(cell, cell.Text))
                .ToArray();
        }

        public ChartWorkbookRangeCell[] ReadRangeCells(string? formula)
        {
            ChartWorkbookRangeResolution resolution = ResolveRangeFormula(formula);
            if (!TryParseRange(resolution.ResolvedFormula, out string? sheetName, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow) ||
                !sheets.TryGetValue(sheetName, out ChartWorksheetData? worksheet))
            {
                return [];
            }

            var values = new List<ChartWorkbookRangeCell>();
            int minColumn = Math.Min(firstColumn, lastColumn);
            int maxColumn = Math.Max(firstColumn, lastColumn);
            int minRow = Math.Min(firstRow, lastRow);
            int maxRow = Math.Max(firstRow, lastRow);
            int index = 0;
            for (int row = minRow; row <= maxRow; row++)
            {
                for (int column = minColumn; column <= maxColumn; column++)
                {
                    string reference = ToCellReference(column, row);
                    bool hasCell = worksheet.Cells.TryGetValue(reference, out ChartWorkbookCell cell);
                    ChartWorkbookCellFormat format = hasCell ? styles.ResolveCellFormat(cell.StyleIndex) : default;
                    values.Add(new ChartWorkbookRangeCell(
                        index,
                        reference,
                        sheetName,
                        resolution.SourceFormula,
                        resolution.ResolvedFormula,
                        resolution.SourceKind,
                        resolution.DefinedName,
                        resolution.TableName,
                        resolution.TableColumnName,
                        resolution.TableColumnId,
                        hasCell ? cell.Text : string.Empty,
                        hasCell ? cell.RawValue : string.Empty,
                        hasCell,
                        hasCell && cell.HasValue,
                        hasCell && cell.HasValueElement,
                        hasCell ? cell.SharedStringIndex : null,
                        hasCell ? cell.SharedStringRunCount : 0,
                        hasCell && cell.SharedStringHasRichText,
                        hasCell && cell.SharedStringHasPhoneticText,
                        hasCell && cell.SharedStringPreserveSpace,
                        hasCell ? cell.StyleIndex : null,
                        hasCell ? cell.CellType : string.Empty,
                        hasCell ? cell.ValueKind : ChartWorkbookCellValueKind.Blank,
                        hasCell ? cell.Formula : string.Empty,
                        hasCell ? cell.FormulaType : string.Empty,
                        hasCell ? cell.FormulaAttributes : new Dictionary<string, string>(),
                        format.NumberFormatId,
                        format.NumberFormatCode,
                        format.ApplyNumberFormat,
                        worksheet.HiddenRows.Contains(row),
                        worksheet.HiddenColumns.Contains(column)));
                    index++;
                }
            }

            return values.ToArray();
        }

        private ChartWorkbookRangeResolution ResolveRangeFormula(string? formula)
        {
            string sourceFormula = formula?.Trim() ?? string.Empty;
            if (sourceFormula.Length == 0)
            {
                return new ChartWorkbookRangeResolution(string.Empty, string.Empty, ChartWorkbookRangeSourceKind.Unknown, string.Empty, string.Empty, string.Empty, null);
            }

            string resolvedFormula = sourceFormula;
            string definedNameName = string.Empty;
            bool fromDefinedName = false;
            if (!sourceFormula.Contains('!', StringComparison.Ordinal) &&
                definedNames.TryGetValue(sourceFormula, out string? definedFormula))
            {
                resolvedFormula = definedFormula;
                definedNameName = sourceFormula;
                fromDefinedName = true;
            }

            string structuredResolvedFormula = ResolveStructuredReferenceFormula(
                resolvedFormula,
                out string tableName,
                out string tableColumnName,
                out int? tableColumnId) ?? string.Empty;
            bool fromStructuredReference = !string.Equals(structuredResolvedFormula, resolvedFormula, StringComparison.Ordinal);
            resolvedFormula = structuredResolvedFormula;

            ChartWorkbookRangeSourceKind sourceKind = (fromDefinedName, fromStructuredReference) switch
            {
                (true, true) => ChartWorkbookRangeSourceKind.DefinedNameStructuredReference,
                (true, false) => ChartWorkbookRangeSourceKind.DefinedName,
                (false, true) => ChartWorkbookRangeSourceKind.StructuredReference,
                _ => ChartWorkbookRangeSourceKind.DirectRange
            };
            return new ChartWorkbookRangeResolution(sourceFormula, resolvedFormula, sourceKind, definedNameName, tableName, tableColumnName, tableColumnId);
        }

        private string? ResolveStructuredReferenceFormula(string? formula, out string tableName, out string tableColumnName, out int? tableColumnId)
        {
            tableName = string.Empty;
            tableColumnName = string.Empty;
            tableColumnId = null;
            if (string.IsNullOrWhiteSpace(formula))
            {
                return formula;
            }

            string trimmed = formula.Trim();
            int open = trimmed.IndexOf('[', StringComparison.Ordinal);
            if (open <= 0 || trimmed[^1] != ']')
            {
                return trimmed;
            }

            string candidateTableName = trimmed[..open];
            if (!tables.TryGetValue(candidateTableName, out ChartWorkbookTable table))
            {
                return trimmed;
            }

            string body = trimmed[(open + 1)..^1];
            if (!TryParseStructuredReferenceBody(body, out string? columnName, out bool includeHeader, out bool onlyHeader) ||
                string.IsNullOrWhiteSpace(columnName))
            {
                return trimmed;
            }

            int columnOffset = -1;
            ChartWorkbookTableColumn tableColumn = default;
            for (int i = 0; i < table.Columns.Count; i++)
            {
                if (string.Equals(table.Columns[i].Name, columnName, StringComparison.OrdinalIgnoreCase))
                {
                    columnOffset = i;
                    tableColumn = table.Columns[i];
                    break;
                }
            }

            if (columnOffset < 0)
            {
                return trimmed;
            }

            int column = table.FirstColumn + columnOffset;
            int headerRowCount = Math.Max(0, table.HeaderRowCount);
            int totalsRowCount = Math.Max(0, table.TotalsRowCount);
            int firstRow = onlyHeader ? table.FirstRow : includeHeader ? table.FirstRow : table.FirstRow + headerRowCount;
            int lastRow = onlyHeader ? table.FirstRow + headerRowCount - 1 : includeHeader ? table.LastRow : table.LastRow - totalsRowCount;
            if (firstRow > lastRow)
            {
                return trimmed;
            }

            tableName = table.Name;
            tableColumnName = tableColumn.Name;
            tableColumnId = tableColumn.Id;
            return FormattableString.Invariant($"{QuoteSheetName(table.SheetName)}!{ToCellReference(column, firstRow)}:{ToCellReference(column, lastRow)}");
        }

        private static bool TryParseStructuredReferenceBody(string body, out string columnName, out bool includeHeader, out bool onlyHeader)
        {
            columnName = string.Empty;
            includeHeader = false;
            onlyHeader = false;
            string trimmed = body.Trim();
            if (trimmed.Length == 0)
            {
                return false;
            }

            if (trimmed[0] != '[')
            {
                columnName = trimmed;
                return true;
            }

            string[] segments = trimmed
                .Split("],[", StringSplitOptions.TrimEntries)
                .Select(segment => segment.Trim('[', ']'))
                .ToArray();
            foreach (string segment in segments)
            {
                if (string.Equals(segment, "#All", StringComparison.OrdinalIgnoreCase))
                {
                    includeHeader = true;
                }
                else if (string.Equals(segment, "#Headers", StringComparison.OrdinalIgnoreCase))
                {
                    includeHeader = true;
                    onlyHeader = true;
                }
                else if (!string.Equals(segment, "#Data", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(segment, "#Totals", StringComparison.OrdinalIgnoreCase))
                {
                    columnName = segment;
                }
            }

            return !string.IsNullOrWhiteSpace(columnName);
        }

        private static string QuoteSheetName(string sheetName)
        {
            return sheetName.Any(character => !char.IsAsciiLetterOrDigit(character) && character != '_')
                ? "'" + sheetName.Replace("'", "''", StringComparison.Ordinal) + "'"
                : sheetName;
        }

        private static IReadOnlyDictionary<string, ChartWorksheetData> ConvertWorkbookSheets(
            IReadOnlyDictionary<string, Dictionary<string, string>> sheets)
        {
            var converted = new Dictionary<string, ChartWorksheetData>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, Dictionary<string, string>> sheet in sheets)
            {
                var cells = new Dictionary<string, ChartWorkbookCell>(StringComparer.OrdinalIgnoreCase);
                foreach (KeyValuePair<string, string> cell in sheet.Value)
                {
                    cells[cell.Key] = new ChartWorkbookCell(cell.Value, cell.Value, true, true, null, 0, false, false, false, null, string.Empty, ReadWorkbookCellValueKind(null, cell.Value, true), string.Empty, string.Empty, new Dictionary<string, string>());
                }

                converted[sheet.Key] = new ChartWorksheetData(cells, new HashSet<int>(), new HashSet<int>());
            }

            return converted;
        }

        private static ChartWorkbookSheet[] ConvertWorkbookSheetRecords(IReadOnlyDictionary<string, ChartWorksheetData> sheets)
        {
            return sheets.Keys
                .Select((name, index) => new ChartWorkbookSheet(name, string.Empty, string.Empty, string.Empty, index, string.Empty))
                .ToArray();
        }

        private static IReadOnlyDictionary<string, ChartWorksheetData> ConvertWorkbookSheets(
            IReadOnlyDictionary<string, Dictionary<string, ChartWorkbookCell>> sheets)
        {
            var converted = new Dictionary<string, ChartWorksheetData>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, Dictionary<string, ChartWorkbookCell>> sheet in sheets)
            {
                converted[sheet.Key] = new ChartWorksheetData(sheet.Value, new HashSet<int>(), new HashSet<int>());
            }

            return converted;
        }

        private static bool TryParseRange(string? formula, out string sheetName, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow)
        {
            sheetName = string.Empty;
            firstColumn = 0;
            firstRow = 0;
            lastColumn = 0;
            lastRow = 0;
            if (string.IsNullOrWhiteSpace(formula))
            {
                return false;
            }

            string trimmed = formula.Trim();
            int separator = trimmed.LastIndexOf('!');
            if (separator <= 0 || separator == trimmed.Length - 1)
            {
                return false;
            }

            sheetName = NormalizeSheetName(trimmed[..separator]);
            string[] references = trimmed[(separator + 1)..].Split(':', 2, StringSplitOptions.TrimEntries);
            if (!TryParseCellReference(references[0], out firstColumn, out firstRow))
            {
                return false;
            }

            if (references.Length == 1)
            {
                lastColumn = firstColumn;
                lastRow = firstRow;
                return true;
            }

            return TryParseCellReference(references[1], out lastColumn, out lastRow);
        }

        private static string NormalizeSheetName(string sheetName)
        {
            int workbookEnd = sheetName.LastIndexOf(']');
            if (workbookEnd >= 0 && workbookEnd < sheetName.Length - 1)
            {
                sheetName = sheetName[(workbookEnd + 1)..];
            }

            sheetName = sheetName.Trim();
            if (sheetName.Length >= 2 && sheetName[0] == '\'' && sheetName[^1] == '\'')
            {
                sheetName = sheetName[1..^1].Replace("''", "'", StringComparison.Ordinal);
            }

            return sheetName;
        }

        private static bool TryParseCellReference(string reference, out int column, out int row)
        {
            column = 0;
            row = 0;
            string normalized = reference.Replace("$", string.Empty, StringComparison.Ordinal).Trim();
            int index = 0;
            while (index < normalized.Length && char.IsAsciiLetter(normalized[index]))
            {
                column = (column * 26) + (char.ToUpperInvariant(normalized[index]) - 'A' + 1);
                index++;
            }

            if (column <= 0 || index == normalized.Length)
            {
                return false;
            }

            return int.TryParse(normalized[index..], NumberStyles.Integer, CultureInfo.InvariantCulture, out row) && row > 0;
        }

        private static string ToCellReference(int column, int row)
        {
            Span<char> buffer = stackalloc char[16];
            int position = buffer.Length;
            int value = column;
            while (value > 0)
            {
                value--;
                buffer[--position] = (char)('A' + (value % 26));
                value /= 26;
            }

            return string.Concat(new string(buffer[position..]), row.ToString(CultureInfo.InvariantCulture));
        }
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

    private static ChartPlotBox GetDefaultChartPlotBox(ChartFrameBox frame)
    {
        return GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.DefaultCartesian);
    }

    private static ChartPlotBox GetChartPlotBoxPreset(ChartFrameBox frame, ChartPlotBoxPreset preset)
    {
        ChartPlotBoxRatios ratios = preset switch
        {
            ChartPlotBoxPreset.DefaultCartesian => new ChartPlotBoxRatios(
                PptxChartMetricRules.DefaultPlotBoxXRatio,
                PptxChartMetricRules.DefaultPlotBoxYRatio,
                PptxChartMetricRules.DefaultPlotBoxWidthRatio,
                PptxChartMetricRules.DefaultPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarDefault => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarDefaultPlotBoxXRatio,
                PptxChartMetricRules.BarDefaultPlotBoxYRatio,
                PptxChartMetricRules.BarDefaultPlotBoxWidthRatio,
                PptxChartMetricRules.BarDefaultPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarOverlayOnly => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarOverlayOnlyPlotBoxXRatio,
                PptxChartMetricRules.BarOverlayOnlyPlotBoxYRatio,
                PptxChartMetricRules.BarOverlayOnlyPlotBoxWidthRatio,
                PptxChartMetricRules.BarOverlayOnlyPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarNoTitleBottomLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxXRatio,
                PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxYRatio,
                PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxWidthRatio,
                PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarTitleNoLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarTitleNoLegendPlotBoxXRatio,
                PptxChartMetricRules.BarTitleNoLegendPlotBoxYRatio,
                PptxChartMetricRules.BarTitleNoLegendPlotBoxWidthRatio,
                PptxChartMetricRules.BarTitleNoLegendPlotBoxHeightRatio),
            ChartPlotBoxPreset.BarTitleNoLegendInsideCrossing => new ChartPlotBoxRatios(
                PptxChartMetricRules.BarTitleNoLegendInsideCrossingPlotBoxXRatio,
                PptxChartMetricRules.BarTitleNoLegendInsideCrossingPlotBoxYRatio,
                PptxChartMetricRules.BarTitleNoLegendInsideCrossingPlotBoxWidthRatio,
                PptxChartMetricRules.BarTitleNoLegendInsideCrossingPlotBoxHeightRatio),
            ChartPlotBoxPreset.HorizontalBarTitleNoLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.HorizontalBarTitleNoLegendPlotBoxXRatio,
                PptxChartMetricRules.HorizontalBarTitleNoLegendPlotBoxYRatio,
                PptxChartMetricRules.HorizontalBarTitleNoLegendPlotBoxWidthRatio,
                PptxChartMetricRules.HorizontalBarTitleNoLegendPlotBoxHeightRatio),
            ChartPlotBoxPreset.LineNoTitleRightLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.LineNoTitleRightLegendPlotBoxXRatio,
                PptxChartMetricRules.LineNoTitleRightLegendPlotBoxYRatio,
                PptxChartMetricRules.LineNoTitleRightLegendPlotBoxWidthRatio,
                PptxChartMetricRules.LineNoTitleRightLegendPlotBoxHeightRatio),
            ChartPlotBoxPreset.LineTitleRightLegend => new ChartPlotBoxRatios(
                PptxChartMetricRules.LineTitleRightLegendPlotBoxXRatio,
                PptxChartMetricRules.LineTitleRightLegendPlotBoxYRatio,
                PptxChartMetricRules.LineTitleRightLegendPlotBoxWidthRatio,
                PptxChartMetricRules.LineTitleRightLegendPlotBoxHeightRatio),
            _ => throw new ArgumentOutOfRangeException(nameof(preset), preset, null)
        };

        return GetChartPlotBox(frame, ratios);
    }

    private static ChartPlotBox GetChartPlotBox(ChartFrameBox frame, ChartPlotBoxRatios ratios)
    {
        return new ChartPlotBox(
            frame.X + frame.Width * ratios.Left,
            frame.Y + frame.Height * ratios.Top,
            frame.Width * ratios.Width,
            frame.Height * ratios.Height);
    }

    private static ChartPlotBox GetPolarChartPlotBox(PptxDocument document, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        ChartPlotBox defaultPlotBox = new(frame.X, frame.Y, frame.Width, frame.Height);
        return TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout manualPlotLayout)
            ? manualPlotLayout.PlotBox
            : defaultPlotBox;
    }

    private static ChartPolarGeometry GetPieOrDoughnutGeometry(ChartPolarKind kind, ChartPlotBox plotBox, double explosionReserve, ChartLegendLayout legend)
    {
        double radius = Math.Min(plotBox.Width, plotBox.Height) * GetPieOrDoughnutRadiusRatio(kind, legend);
        if (explosionReserve > 0d)
        {
            radius /= 1d + explosionReserve;
        }

        double centerXRatio = GetPieOrDoughnutCenterXRatio(kind, legend);
        double centerYRatio = GetPieOrDoughnutCenterYRatio(kind, legend);
        double centerXOffset = GetPieOrDoughnutCenterXOffset(kind, radius, explosionReserve, legend);
        return new ChartPolarGeometry(
            plotBox.X + plotBox.Width * centerXRatio + centerXOffset,
            plotBox.Y + plotBox.Height * centerYRatio,
            radius);
    }

    private static double GetPieOrDoughnutRadiusRatio(ChartPolarKind kind, ChartLegendLayout legend)
    {
        bool hasLegend = legend.Visible && !legend.Overlay;
        return kind switch
        {
            ChartPolarKind.Pie => PptxChartMetricRules.PieRadiusRatio,
            ChartPolarKind.Doughnut when !hasLegend => PptxChartMetricRules.DoughnutNoLegendRadiusRatio,
            ChartPolarKind.Doughnut when legend.PositionKind == PptxSceneChartLegendPosition.Left => PptxChartMetricRules.DoughnutNoLegendRadiusRatio,
            ChartPolarKind.Doughnut when IsHorizontalLegendPosition(legend.PositionKind) => PptxChartMetricRules.DoughnutHorizontalLegendRadiusRatio,
            ChartPolarKind.Doughnut => PptxChartMetricRules.PieRadiusRatio,
            _ => PptxChartMetricRules.PieRadiusRatio
        };
    }

    private static double GetPieOrDoughnutCenterXRatio(ChartPolarKind kind, ChartLegendLayout legend)
    {
        bool hasLegend = legend.Visible && !legend.Overlay;
        return kind switch
        {
            ChartPolarKind.Pie => hasLegend ? PptxChartMetricRules.PieCenterXRatio : PptxChartMetricRules.PieNoLegendCenterXRatio,
            ChartPolarKind.Doughnut when hasLegend && legend.PositionKind == PptxSceneChartLegendPosition.Right => PptxChartMetricRules.DoughnutRightLegendCenterXRatio,
            ChartPolarKind.Doughnut when hasLegend && legend.PositionKind == PptxSceneChartLegendPosition.Left => PptxChartMetricRules.DoughnutLeftLegendCenterXRatio,
            ChartPolarKind.Doughnut when hasLegend && IsHorizontalLegendPosition(legend.PositionKind) => PptxChartMetricRules.DoughnutHorizontalLegendCenterXRatio,
            ChartPolarKind.Doughnut => hasLegend ? PptxChartMetricRules.PieCenterXRatio : PptxChartMetricRules.PieNoLegendCenterXRatio,
            _ => hasLegend ? PptxChartMetricRules.PieCenterXRatio : PptxChartMetricRules.PieNoLegendCenterXRatio
        };
    }

    private static double GetPieOrDoughnutCenterYRatio(ChartPolarKind kind, ChartLegendLayout legend)
    {
        bool hasLegend = legend.Visible && !legend.Overlay;
        return kind switch
        {
            ChartPolarKind.Doughnut when !hasLegend => PptxChartMetricRules.DoughnutNoLegendCenterYRatio,
            ChartPolarKind.Doughnut when legend.PositionKind == PptxSceneChartLegendPosition.Left => PptxChartMetricRules.DoughnutNoLegendCenterYRatio,
            ChartPolarKind.Doughnut when legend.PositionKind == PptxSceneChartLegendPosition.Top => PptxChartMetricRules.DoughnutTopLegendCenterYRatio,
            ChartPolarKind.Doughnut when legend.PositionKind == PptxSceneChartLegendPosition.Bottom => PptxChartMetricRules.DoughnutBottomLegendCenterYRatio,
            _ => PptxChartMetricRules.PieCenterYRatio
        };
    }

    private static bool IsHorizontalLegendPosition(PptxSceneChartLegendPosition position)
    {
        return position is PptxSceneChartLegendPosition.Top or PptxSceneChartLegendPosition.Bottom;
    }

    private static double GetPieOrDoughnutCenterXOffset(ChartPolarKind kind, double radius, double explosionReserve, ChartLegendLayout legend)
    {
        if (kind == ChartPolarKind.Doughnut &&
            explosionReserve > 0d &&
            legend.Visible &&
            !legend.Overlay &&
            legend.PositionKind == PptxSceneChartLegendPosition.Right)
        {
            return radius * explosionReserve * PptxChartMetricRules.DoughnutExplosionCenterOffsetRatio;
        }

        return 0d;
    }

    private static ChartPolarLayout ResolvePieOrDoughnutLayout(ChartPolarKind kind, ChartPlotBox plotBox, IReadOnlyDictionary<int, double> pointExplosions, ChartLegendLayout legend)
    {
        double explosionReserve = pointExplosions.Count == 0 ? 0d : pointExplosions.Values.Max();
        bool hasLegend = legend.Visible && !legend.Overlay;
        return new ChartPolarLayout(
            kind,
            plotBox,
            GetPieOrDoughnutGeometry(kind, plotBox, explosionReserve, legend),
            explosionReserve,
            hasLegend);
    }

    private static ChartRadarLayout ResolveRadarLayout(ChartPlotBox plotBox, PptxSceneChartRadarStyle radarStyle, IReadOnlyList<IReadOnlyList<double>> series)
    {
        ChartRadarStyle style = radarStyle == PptxSceneChartRadarStyle.Filled
            ? ChartRadarStyle.Filled
            : ChartRadarStyle.Marker;
        return new ChartRadarLayout(
            plotBox,
            GetRadarChartGeometry(plotBox, style),
            style,
            Math.Max(3, series.Max(values => values.Count)),
            ResolveRadarLabelRules(style));
    }

    private static ChartPolarGeometry GetRadarChartGeometry(ChartPlotBox plotBox, ChartRadarStyle style)
    {
        ChartRadarGeometryRule rule = ResolveRadarGeometryRule(style);
        return new ChartPolarGeometry(
            plotBox.X + plotBox.Width * rule.CenterXRatio,
            plotBox.Y + plotBox.Height * rule.CenterYRatio,
            Math.Min(plotBox.Width, plotBox.Height) * rule.RadiusRatio);
    }

    private static ChartRadarGeometryRule ResolveRadarGeometryRule(ChartRadarStyle style)
    {
        return style == ChartRadarStyle.Filled
            ? new ChartRadarGeometryRule(CenterXRatio: 0.5d, CenterYRatio: 0.4583333333333333d, RadiusRatio: 0.3825d)
            : new ChartRadarGeometryRule(CenterXRatio: 0.5d, CenterYRatio: 0.5d, RadiusRatio: 0.4226d);
    }

    private static ChartRadarLabelRules ResolveRadarLabelRules(ChartRadarStyle style)
    {
        double categoryHorizontalGapFactor = style == ChartRadarStyle.Filled ? 0.35d : 0.41d;
        return new ChartRadarLabelRules(
            CategoryVerticalGapFactor: 0.65d,
            CategoryHorizontalGapFactor: categoryHorizontalGapFactor,
            CategoryBaselineBaseFactor: -0.309d,
            CategoryBaselineSineFactor: -0.005d,
            CategoryBaselineSineSquaredFactor: 0.397d,
            ValueGapFactor: 1.01d,
            ValueBaselineOffsetFactor: 0.25d,
            ValueWidthFactor: 3.0d);
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

    private static bool TryReadSceneOrXmlManualPlotLayout(PptxSceneChart? sceneChart, XDocument chartXml, ChartFrameBox frame, ChartPlotBox defaultPlotBox, out ChartPlotLayout plotLayout)
    {
        if (sceneChart is not null)
        {
            return TryBuildManualPlotLayout(sceneChart.PlotAreaLayout, frame, defaultPlotBox, out plotLayout);
        }

        return TryReadManualPlotLayout(chartXml, frame, defaultPlotBox, out plotLayout);
    }

    private static bool TryReadManualPlotLayout(XDocument chartXml, ChartFrameBox frame, ChartPlotBox defaultPlotBox, out ChartPlotLayout plotLayout)
    {
        plotLayout = default;
        XElement? plotArea = chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        if (plotArea is null)
        {
            return false;
        }

        return TryBuildManualPlotLayout(ReadManualLayout(plotArea), frame, defaultPlotBox, out plotLayout);
    }

    private static bool TryBuildManualPlotLayout(PptxSceneChartManualLayout layout, ChartFrameBox frame, ChartPlotBox defaultPlotBox, out ChartPlotLayout plotLayout)
    {
        plotLayout = default;
        if (!TryBuildManualLayoutBox(layout, frame, new ChartLayoutBox(defaultPlotBox.X, defaultPlotBox.Y, defaultPlotBox.Width, defaultPlotBox.Height), out ChartLayoutBox layoutBox))
        {
            return false;
        }

        ChartPlotBox plotBox = new(layoutBox.X, layoutBox.Y, layoutBox.Width, layoutBox.Height);
        plotLayout = new ChartPlotLayout(layoutBox, plotBox, layout.LayoutTargetKind);
        return true;
    }

    private static bool TryBuildManualLayoutBox(PptxSceneChartManualLayout layout, ChartFrameBox frame, ChartLayoutBox defaultBox, out ChartLayoutBox box)
    {
        box = default;
        if (!layout.HasLayout)
        {
            return false;
        }

        ChartPlotBoxRatios defaults = GetLayoutBoxRatios(frame, defaultBox);
        double left = layout.X is { } x
            ? Math.Clamp(ResolveManualLayoutStartRatio(x, layout.XModeKind, defaultBox.X, frame.X, frame.Width), 0d, 1d)
            : defaults.Left;
        double top = layout.Y is { } y
            ? Math.Clamp(ResolveManualLayoutStartRatio(y, layout.YModeKind, frame.Y + frame.Height - defaultBox.Y - defaultBox.Height, 0d, frame.Height), 0d, 1d)
            : defaults.Top;
        double width = layout.Width is { } layoutWidth
            ? Math.Clamp(layoutWidth, 0.02d, 1d)
            : defaults.Width;
        double height = layout.Height is { } layoutHeight
            ? Math.Clamp(layoutHeight, 0.02d, 1d)
            : defaults.Height;
        double right = IsManualLayoutEdgeMode(layout.WidthModeKind)
            ? Math.Clamp(layout.Width ?? defaults.Right, left, 1d)
            : left + width;
        double bottom = IsManualLayoutEdgeMode(layout.HeightModeKind)
            ? Math.Clamp(layout.Height ?? defaults.Bottom, top, 1d)
            : top + height;
        double boxWidth = Math.Max(0d, right - left) * frame.Width;
        double boxHeight = Math.Max(0d, bottom - top) * frame.Height;
        double boxX = frame.X + left * frame.Width;
        double boxY = frame.Y + frame.Height - bottom * frame.Height;
        box = new ChartLayoutBox(boxX, boxY, boxWidth, boxHeight);
        return boxWidth > 0d && boxHeight > 0d;
    }

    private static ChartPlotBoxRatios GetPlotBoxRatios(ChartFrameBox frame, ChartPlotBox plotBox)
    {
        return GetLayoutBoxRatios(frame, new ChartLayoutBox(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height));
    }

    private static ChartPlotBoxRatios GetLayoutBoxRatios(ChartFrameBox frame, ChartLayoutBox box)
    {
        if (frame.Width <= 0d || frame.Height <= 0d)
        {
            return new ChartPlotBoxRatios(0d, 0d, 1d, 1d);
        }

        double left = (box.X - frame.X) / frame.Width;
        double top = (frame.Y + frame.Height - box.Y - box.Height) / frame.Height;
        double width = box.Width / frame.Width;
        double height = box.Height / frame.Height;
        return new ChartPlotBoxRatios(left, top, width, height);
    }

    private static double ResolveManualLayoutStartRatio(double value, PptxSceneChartManualLayoutMode mode, double defaultStart, double frameStart, double frameLength)
    {
        if (IsManualLayoutFactorMode(mode) && frameLength > 0d)
        {
            return (defaultStart - frameStart) / frameLength + value;
        }

        return value;
    }

    private static bool IsManualLayoutEdgeMode(PptxSceneChartManualLayoutMode mode)
    {
        return mode == PptxSceneChartManualLayoutMode.Edge;
    }

    private static bool IsManualLayoutFactorMode(PptxSceneChartManualLayoutMode mode)
    {
        return mode == PptxSceneChartManualLayoutMode.Factor;
    }

    private static PptxSceneChartManualLayout ReadManualLayout(XElement parent)
    {
        XElement? manualLayout = parent
            .Element(ChartNamespace + "layout")
            ?.Element(ChartNamespace + "manualLayout");
        if (manualLayout is null)
        {
            return default;
        }

        string layoutTarget = ReadManualLayoutValue(manualLayout, "layoutTarget");
        string xMode = ReadManualLayoutValue(manualLayout, "xMode");
        string yMode = ReadManualLayoutValue(manualLayout, "yMode");
        string widthMode = ReadManualLayoutValue(manualLayout, "wMode");
        string heightMode = ReadManualLayoutValue(manualLayout, "hMode");
        return new PptxSceneChartManualLayout(
            true,
            ReadManualLayoutFactor(manualLayout, "x"),
            ReadManualLayoutFactor(manualLayout, "y"),
            ReadManualLayoutFactor(manualLayout, "w"),
            ReadManualLayoutFactor(manualLayout, "h"),
            PptxSceneBuilder.ParseChartManualLayoutTarget(layoutTarget),
            layoutTarget,
            PptxSceneBuilder.ParseChartManualLayoutMode(xMode),
            xMode,
            PptxSceneBuilder.ParseChartManualLayoutMode(yMode),
            yMode,
            PptxSceneBuilder.ParseChartManualLayoutMode(widthMode),
            widthMode,
            PptxSceneBuilder.ParseChartManualLayoutMode(heightMode),
            heightMode);
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

    private static IReadOnlyList<XElement> ReadChartValueAxesForChart(XDocument chartXml, XElement chartElement)
    {
        string[] axisIds = chartElement
            .Elements(ChartNamespace + "axId")
            .Select(axis => (string?)axis.Attribute("val"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id!)
            .ToArray();
        if (axisIds.Length == 0)
        {
            return chartXml.Descendants(ChartNamespace + "valAx").ToArray();
        }

        var axes = new List<XElement>();
        foreach (string axisId in axisIds)
        {
            XElement? axis = chartXml
                .Descendants(ChartNamespace + "valAx")
                .FirstOrDefault(candidate => string.Equals(ReadChartAxisId(candidate), axisId, StringComparison.Ordinal));
            if (axis is not null)
            {
                axes.Add(axis);
            }
        }

        return axes;
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

    private static ChartValueExtents ReadChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback, bool useNearMaximumHeadroom = false, double nearMaximumHeadroomRatio = PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio)
    {
        return ReadChartValueAxisExtents(valueAxis, fallback, PptxChartMetricRules.AxisNiceTickTargetCount, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
    }

    private static ChartValueExtents ReadBubbleChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback)
    {
        return ReadChartValueAxisExtents(valueAxis, fallback, PptxChartMetricRules.BubbleAxisBoundsTickTargetCount);
    }

    private static ChartValueExtents ReadChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback, double boundsTickTargetCount, bool useNearMaximumHeadroom = false, double nearMaximumHeadroomRatio = PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio)
    {
        XElement? scaling = valueAxis?.Element(ChartNamespace + "scaling");
        if (scaling is null)
        {
            return fallback;
        }

        double min = ReadAxisScalingValue(scaling, ChartAxisScalingBound.Minimum) ?? GetNiceChartAxisMin(fallback.Min, fallback.Max);
        double max = ReadAxisScalingValue(scaling, ChartAxisScalingBound.Maximum) ?? GetNiceChartAxisMax(fallback.Max, min, boundsTickTargetCount, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        return max > min
            ? new ChartValueExtents(min, max)
            : fallback;
    }

    private static ChartValueExtents ReadSceneOrXmlChartValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents fallback, bool useNearMaximumHeadroom = false, double nearMaximumHeadroomRatio = PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio)
    {
        if (axis is null)
        {
            return ReadChartValueAxisExtents(valueAxis, fallback, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        }

        if (!axis.HasScaling)
        {
            return fallback;
        }

        double min = axis.Minimum ?? GetNiceChartAxisMin(fallback.Min, fallback.Max);
        double max = axis.Maximum ?? GetNiceChartAxisMax(fallback.Max, min, PptxChartMetricRules.AxisNiceTickTargetCount, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        return max > min
            ? new ChartValueExtents(min, max)
            : fallback;
    }

    private static ChartValueExtents ReadPercentStackedAwareValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents fallback, bool percentStacked, bool useNearMaximumHeadroom = false, double nearMaximumHeadroomRatio = PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio)
    {
        ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(axis, valueAxis, fallback, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
        if (!percentStacked)
        {
            return extents;
        }

        double min = HasSceneOrXmlAxisScalingValue(axis, valueAxis, ChartAxisScalingBound.Minimum) ? extents.Min : 0d;
        double max = HasSceneOrXmlAxisScalingValue(axis, valueAxis, ChartAxisScalingBound.Maximum) ? extents.Max : 1d;
        return max > min
            ? new ChartValueExtents(min, max)
            : extents;
    }

    private static bool HasSceneOrXmlAxisScalingValue(PptxSceneChartAxis? axis, XElement? valueAxis, ChartAxisScalingBound bound)
    {
        if (axis is not null)
        {
            return bound == ChartAxisScalingBound.Minimum
                ? axis.Minimum is not null
                : axis.Maximum is not null;
        }

        string elementName = bound == ChartAxisScalingBound.Minimum ? "min" : "max";
        return valueAxis?
            .Element(ChartNamespace + "scaling")
            ?.Element(ChartNamespace + elementName) is not null;
    }

    private static double? ReadAxisScalingValue(XElement scaling, ChartAxisScalingBound bound)
    {
        string elementName = bound == ChartAxisScalingBound.Minimum ? "min" : "max";
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

    private static ChartAxisUnits ResolvePercentStackedAxisUnits(ChartAxisUnits axisUnits, bool percentStacked)
    {
        return percentStacked && axisUnits.MajorUnit is null
            ? axisUnits with { MajorUnit = 0.1d }
            : axisUnits;
    }

    private static ChartAxisUnits ResolveBubbleAxisUnits(ChartAxisUnits axisUnits, ChartValueExtents extents)
    {
        return axisUnits.MajorUnit is null
            ? axisUnits with { MajorUnit = ChooseChartAxisMajorUnit(Math.Max(1d, extents.Max - extents.Min), PptxChartMetricRules.BubbleAxisNiceTickTargetCount) }
            : axisUnits;
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

        bool noFill = shapeProperties.Element(DrawingNamespace + "noFill") is not null;
        ChartSeriesFill? fill = !noFill && TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
            ? new ChartSeriesFill(fillColor, fillAlpha)
            : null;
        GradientFill? gradientFill = !noFill && PptxSceneBuilder.TryReadShapeGradientFill(shapeProperties, theme, out PptxSceneGradientFill sceneGradientFill)
            ? ToGradientFill(sceneGradientFill)
            : null;
        ChartSeriesStroke? stroke = TryReadLineWithAlpha(shapeProperties, theme, out RgbColor strokeColor, out double lineWidth, out double strokeAlpha)
            ? new ChartSeriesStroke(strokeColor, strokeAlpha, lineWidth)
            : null;
        return new ChartShapeStyle(fill, gradientFill, stroke);
    }

    private static void RenderChartShapeStyle(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartShapeStyle style)
    {
        if (style.GradientFill is { } gradientFill)
        {
            DrawLinearGradientFill(graphics, gradientFill, x, y, width, height);
        }
        else if (style.Fill is { } fill)
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

    private static void ClipChartPlotArea(PdfGraphicsBuilder graphics, double x, double y, double width, double height)
    {
        graphics.ClipRectangle(x, y, width, height);
    }

    private static IReadOnlyList<PdfFontResource> RenderChartTitle(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        if (string.IsNullOrWhiteSpace(title))
        {
            return [];
        }

        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        double x = frame.X;
        double y = frame.Y;
        double width = frame.Width;
        double height = frame.Height;
        bool titleManualLayoutApplied = false;
        if (sceneChart?.Title.Layout.HasLayout == true &&
            TryBuildManualLayoutBox(sceneChart.Title.Layout, frame, new ChartLayoutBox(frame.X, frame.Y, frame.Width, frame.Height), out ChartLayoutBox titleBox))
        {
            x = titleBox.X;
            y = titleBox.Y;
            width = titleBox.Width;
            height = titleBox.Height;
            titleManualLayoutApplied = true;
        }

        bool isAutoTitle = IsAutoGeneratedChartTitle(sceneChart);
        ChartTextStyle style = ReadSceneOrXmlChartTitleTextStyle(theme, sceneChart, chartXml, isAutoTitle);
        double fontSize = style.FontSize;
        double titleBaselineRatio = !titleManualLayoutApplied && HasPolarChart(chartXml)
            ? PptxChartMetricRules.PolarTitleBaselineYRatio
            : PptxChartMetricRules.TitleBaselineYRatio;
        double fallbackBaselineY = y + height * titleBaselineRatio;
        double baselineY = titleManualLayoutApplied
            ? fallbackBaselineY
            : ResolveChartTitleBaselineY(document, theme, bounds, chartXml, sceneChart, fallbackBaselineY, fontSize);
        RenderChartShapeStyle(graphics, x, y, width, height, ReadSceneOrXmlChartTitleShapeStyle(theme, sceneChart, chartXml));
        double titleX = x + width * PptxChartMetricRules.TitleXInsetRatio;
        double titleWidth = width * PptxChartMetricRules.TitleWidthRatio;
        double titleHeight = fontSize * PptxChartMetricRules.TitleHeightFactor;
        var runs = new List<TextRun>();
        AddChartRichTextRuns(
            runs,
            ReadSceneOrXmlChartTitleTextRuns(theme, sceneChart, chartXml),
            title.Trim(),
            titleX,
            baselineY,
            titleWidth,
            titleHeight,
            x,
            y,
            width,
            height,
            style,
            TextAlignment.Center);
        return RenderTextRuns(runs, graphics, "CT");
    }

    private static ChartShapeStyle ReadSceneOrXmlChartTitleShapeStyle(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        return sceneChart is null
            ? ReadChartShapeStyle(chartXml.Descendants(ChartNamespace + "title").FirstOrDefault()?.Element(ChartNamespace + "spPr"), theme)
            : ToChartShapeStyle(sceneChart.Title.ShapeStyle);
    }

    private static IReadOnlyList<PdfFontResource> RenderManualChartAxisTitles(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        var fonts = new List<PdfFontResource>();
        if (sceneChart is not null)
        {
            foreach (PptxSceneChartAxis axis in sceneChart.Axes)
            {
                fonts.AddRange(RenderManualChartAxisTitle(
                    theme,
                    graphics,
                    frame,
                    axis.Title.Text,
                    ToChartTextRuns(axis.Title.TextRuns),
                    axis.Title.Layout,
                    ToChartShapeStyle(axis.Title.ShapeStyle),
                    ToChartTextStyleOverride(sceneChart.TextStyle),
                    ToChartTextStyleOverride(axis.Title.TextStyle)));
            }

            return fonts;
        }

        foreach (XElement axis in ReadChartAxisElements(chartXml))
        {
            XElement? title = axis.Element(ChartNamespace + "title");
            if (title is null)
            {
                continue;
            }

            string? text = ReadChartText(title.Element(ChartNamespace + "tx"));
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            fonts.AddRange(RenderManualChartAxisTitle(
                theme,
                graphics,
                frame,
                text,
                ToChartTextRuns(PptxSceneBuilder.ReadChartTextRuns(title.Element(ChartNamespace + "tx"), theme)),
                ReadManualLayout(title),
                ReadChartShapeStyle(title.Element(ChartNamespace + "spPr"), theme),
                ReadChartTextStyleFromTxPr(chartXml.Root, theme),
                ReadChartTextStyleFromTxPr(title, theme)));
        }

        return fonts;
    }

    private static IReadOnlyList<PdfFontResource> RenderManualChartAxisTitle(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartFrameBox frame,
        string? text,
        IReadOnlyList<ChartTextRunOverride> textRuns,
        PptxSceneChartManualLayout layout,
        ChartShapeStyle shapeStyle,
        ChartTextStyleOverride chartTextStyle,
        ChartTextStyleOverride titleTextStyle)
    {
        if (string.IsNullOrWhiteSpace(text) ||
            !TryBuildManualLayoutBox(layout, frame, new ChartLayoutBox(frame.X, frame.Y, frame.Width, frame.Height), out ChartLayoutBox titleBox))
        {
            return [];
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize);
        style = MergeChartTextStyle(style, chartTextStyle);
        style = MergeChartTextStyle(style, titleTextStyle);
        RenderChartShapeStyle(graphics, titleBox.X, titleBox.Y, titleBox.Width, titleBox.Height, shapeStyle);
        double titleHeight = style.FontSize * PptxChartMetricRules.TitleHeightFactor;
        double baselineY = titleBox.Y + titleBox.Height * PptxChartMetricRules.TitleBaselineYRatio;
        var runs = new List<TextRun>();
        AddChartRichTextRuns(
            runs,
            textRuns,
            text.Trim(),
            titleBox.X + titleBox.Width * PptxChartMetricRules.TitleXInsetRatio,
            baselineY,
            titleBox.Width * PptxChartMetricRules.TitleWidthRatio,
            titleHeight,
            titleBox.X,
            titleBox.Y,
            titleBox.Width,
            titleBox.Height,
            style,
            TextAlignment.Center);
        return RenderTextRuns(runs, graphics, "CAT");
    }

    private static IReadOnlyList<XElement> ReadChartAxisElements(XDocument chartXml)
    {
        XElement? plotArea = chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        return plotArea is null
            ? []
            : plotArea.Elements()
                .Where(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Ax", StringComparison.Ordinal))
                .ToArray();
    }

    private static bool HasPolarChart(XDocument chartXml)
    {
        return ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Pie).Any() ||
            ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Doughnut).Any() ||
            ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Radar).Any();
    }

    private static double ResolveChartTitleBaselineY(PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, double fallbackBaselineY, double fontSize)
    {
        XElement? barChart = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Bar).FirstOrDefault();
        if (barChart is null)
        {
            return fallbackBaselineY;
        }

        PptxSceneChartPlot? barPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bar);
        bool horizontalBars = ReadSceneOrXmlChartBarDirection(barPlot, barChart) == PptxSceneChartBarDirection.Bar;
        ChartLayout layout = GetBarChartLayout(document, theme, bounds, chartXml, sceneChart, barPlot, barChart, horizontalBars);
        ChartPlotBox titleAnchorBox = layout.PlotBox;
        if (horizontalBars && layout.ManualPlotLayoutApplied && IsAutoGeneratedChartTitle(sceneChart))
        {
            ChartFrameBox frame = GetChartFrameBox(document, bounds);
            ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, sceneChart, chartXml);
            titleAnchorBox = GetBarChartPlotLayout(
                theme,
                frame,
                chartXml,
                sceneChart,
                barPlot,
                barChart,
                ReadSceneOrXmlChartTitleText(sceneChart, chartXml),
                legend,
                horizontalBars,
                ignoreManualPlotLayout: true).PlotBox;
        }

        double offsetFactor = !horizontalBars && IsAutoGeneratedChartTitle(sceneChart)
            ? PptxChartMetricRules.AutoBarTitleAbovePlotBaselineOffsetFactor
            : PptxChartMetricRules.TitleAbovePlotBaselineOffsetFactor;
        return titleAnchorBox.Y + titleAnchorBox.Height + fontSize * offsetFactor;
    }

    private static bool IsAutoGeneratedChartTitle(PptxSceneChart? sceneChart)
    {
        return sceneChart?.Title.IsAutoGenerated == true;
    }

    private static ChartTextStyle ReadSceneOrXmlChartTitleTextStyle(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml, bool isAutoTitle)
    {
        XElement? title = chartXml.Descendants(ChartNamespace + "title").FirstOrDefault();
        if (sceneChart is null)
        {
            return ResolveAutoChartTitleTextStyle(
                ReadChartTextStyle(theme, chartXml, title, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize),
                ChartTextStyleOverride.Empty,
                isAutoTitle);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize);
        style = MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.TextStyle));
        ChartTextStyleOverride titleStyle = ToChartTextStyleOverride(sceneChart.Title.TextStyle);
        return ResolveAutoChartTitleTextStyle(MergeChartTextStyle(style, titleStyle), titleStyle, isAutoTitle);
    }

    private static ChartTextStyle ResolveAutoChartTitleTextStyle(ChartTextStyle style, ChartTextStyleOverride titleStyle, bool isAutoTitle)
    {
        if (!isAutoTitle || titleStyle.FontSize is not null)
        {
            return style;
        }

        return style with { FontSize = style.FontSize * PptxChartMetricRules.AutoTitleFontScale };
    }

    private static string? ReadSceneOrXmlChartTitleText(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            if (sceneChart.Title.IsAutoDeleted == true)
            {
                return null;
            }

            return sceneChart.Title.Text;
        }

        return ReadChartTitleText(chartXml);
    }

    private static IReadOnlyList<ChartTextRunOverride> ReadSceneOrXmlChartTitleTextRuns(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            return ToChartTextRuns(sceneChart.Title.TextRuns);
        }

        XElement? titleText = chartXml
            .Descendants(ChartNamespace + "title")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "tx");
        return ToChartTextRuns(PptxSceneBuilder.ReadChartTextRuns(titleText, theme));
    }

    private static string? ReadChartTitleText(XDocument chartXml)
    {
        XElement? title = chartXml.Descendants(ChartNamespace + "title").FirstOrDefault();
        if (title is null)
        {
            return null;
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

    private static IReadOnlyList<ChartLegendEntry> BuildFillLegendEntries(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesFill?> seriesFills, int paletteOffset = 0, ChartWorkbookData? workbook = null)
    {
        IReadOnlyList<string> names = ReadSceneOrXmlChartSeriesNames(plot, chartElement, workbook);
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

    private static IReadOnlyList<ChartLegendEntry> BuildCategoryFillLegendEntries(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, ChartWorkbookData? workbook = null)
    {
        IReadOnlyList<string> labels = ReadSceneOrXmlCategoryLabels(plot, chartElement, workbook);
        var entries = new List<ChartLegendEntry>(labels.Count);
        for (int i = 0; i < labels.Count; i++)
        {
            ChartSeriesFill fill = pointFills.TryGetValue(i, out ChartSeriesFill pointFill)
                ? pointFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, i), 1d);
            entries.Add(new ChartLegendEntry(labels[i], fill, null));
        }

        return entries;
    }

    private static IReadOnlyList<ChartLegendEntry> BuildStrokeLegendEntries(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, bool reverseOrder = false, ChartWorkbookData? workbook = null)
    {
        IReadOnlyList<string> names = ReadSceneOrXmlChartSeriesNames(plot, chartElement, workbook);
        var entries = new List<ChartLegendEntry>(names.Count);
        for (int i = 0; i < names.Count; i++)
        {
            entries.Add(new ChartLegendEntry(names[i], null, ChartSeriesStrokeColor(theme, chartPalette, i, seriesStrokes, ChartLineDefaultStrokeWidth)));
        }

        if (reverseOrder)
        {
            entries.Reverse();
        }

        return entries;
    }

    private static IReadOnlyList<string> ReadSceneOrXmlChartSeriesNames(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook = null)
    {
        return plot?.Series.Count > 0
            ? plot.Series
                .Select((series, index) =>
                {
                    if (!string.IsNullOrWhiteSpace(series.Name))
                    {
                        return series.Name.Trim();
                    }

                    string? workbookName = ReadWorkbookTextValues(workbook, series.DataSources.Name)
                        .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
                    return string.IsNullOrWhiteSpace(workbookName) ? $"Series {index + 1}" : workbookName.Trim();
                })
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

    private static ChartLegendLayout ReadChartLegendLayout(PptxTheme theme, XDocument chartXml)
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
        return new ChartLegendLayout(PptxSceneBuilder.ParseChartLegendPosition(position), position, overlay, Visible: true, ReadManualLayout(legend), ReadChartShapeStyle(legend.Element(ChartNamespace + "spPr"), theme));
    }

    private static ChartLegendLayout ReadSceneOrXmlChartLegendLayout(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is null)
        {
            return ReadChartLegendLayout(theme, chartXml);
        }

        if (!sceneChart.Legend.IsDefined)
        {
            return ReadChartLegendLayout(theme, chartXml);
        }

        return sceneChart.Legend.IsDeleted != true
            ? new ChartLegendLayout(sceneChart.Legend.PositionKind, sceneChart.Legend.Position, sceneChart.Legend.Overlay == true, Visible: true, sceneChart.Legend.Layout, ToChartShapeStyle(sceneChart.Legend.ShapeStyle))
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

    private static IReadOnlyList<PdfFontResource> RenderChartLegend(PdfGraphicsBuilder graphics, ChartFrameBox frame, ChartPlotBox plotBox, IReadOnlyList<ChartLegendEntry> entries, ChartLegendLayout layout, ChartTextStyle style)
    {
        if (!layout.Visible || entries.Count == 0)
        {
            return [];
        }

        double fontSize = style.FontSize;
        double markerSize = fontSize * PptxChartMetricRules.LegendMarkerSizeFactor;
        bool horizontal = IsHorizontalLegendPosition(layout.PositionKind);
        bool sideStrokeLegend = !horizontal && entries.All(entry => entry.Stroke is not null && entry.Fill is null);
        bool fillLegendInFullFrame = !sideStrokeLegend && IsSameChartBox(plotBox, frame);
        bool sideFillLegendInFullFrame = !horizontal && fillLegendInFullFrame;
        double lineHeight = fontSize * (sideStrokeLegend
            || sideFillLegendInFullFrame
            ? PptxChartMetricRules.LegendSideStrokeLineHeightFactor
            : PptxChartMetricRules.LegendLineHeightFactor);
        double sideGap = sideStrokeLegend
            ? fontSize * PptxChartMetricRules.LegendSideStrokeGapFactor
            : PptxChartMetricRules.LegendSideGap;
        double markerWidth = sideStrokeLegend
            ? fontSize * PptxChartMetricRules.LegendSideStrokeMarkerWidthFactor
            : markerSize;
        double textGap = sideStrokeLegend
            ? fontSize * PptxChartMetricRules.LegendSideStrokeTextGapFactor
            : PptxChartMetricRules.LegendTextGap;
        double width = horizontal
            ? Math.Min(plotBox.Width, GetPackedHorizontalLegendWidth(entries, fontSize, markerSize))
            : sideFillLegendInFullFrame
                ? GetFrameAnchoredSideFillLegendWidth(entries, fontSize, markerWidth, textGap)
            : Math.Max(PptxChartMetricRules.LegendMinimumSideWidth, plotBox.Width * PptxChartMetricRules.LegendSideWidthRatio);
        double x = layout.PositionKind switch
        {
            PptxSceneChartLegendPosition.Left when sideFillLegendInFullFrame => frame.X + frame.Width * PptxChartMetricRules.LegendFullFrameSideInsetRatio,
            PptxSceneChartLegendPosition.Left => Math.Max(0d, plotBox.X - width - sideGap),
            _ when horizontal => plotBox.X + (plotBox.Width - width) / 2d,
            _ when sideFillLegendInFullFrame => frame.X + frame.Width - width,
            _ => plotBox.X + plotBox.Width + sideGap
        };
        double firstY = layout.PositionKind switch
        {
            PptxSceneChartLegendPosition.Bottom when fillLegendInFullFrame => frame.Y + lineHeight * PptxChartMetricRules.LegendFullFrameBottomBaselineFactor,
            PptxSceneChartLegendPosition.Top when fillLegendInFullFrame => frame.Y + frame.Height - lineHeight * PptxChartMetricRules.LegendFullFrameTopBaselineFactor,
            PptxSceneChartLegendPosition.Bottom => Math.Max(0d, plotBox.Y - lineHeight * PptxChartMetricRules.LegendBottomOffsetFactor),
            PptxSceneChartLegendPosition.Top => plotBox.Y + plotBox.Height + lineHeight * PptxChartMetricRules.LegendTopOffsetFactor,
            _ when sideStrokeLegend => plotBox.Y + plotBox.Height / 2d -
                fontSize * PptxChartMetricRules.LegendSideStrokeBaselineCenterOffsetFactor +
                (entries.Count - 1) * lineHeight / 2d,
            _ when sideFillLegendInFullFrame => frame.Y + frame.Height / 2d +
                (entries.Count - 1) * lineHeight / 2d,
            _ when !horizontal => plotBox.Y + plotBox.Height / 2d +
                fontSize * PptxChartMetricRules.LegendMarkerBaselineFactor +
                (entries.Count - 1) * lineHeight / 2d,
            _ => plotBox.Y + plotBox.Height - lineHeight
        };
        double clipHeight = horizontal ? lineHeight * PptxChartMetricRules.LegendHorizontalClipHeightFactor : Math.Max(lineHeight, entries.Count * lineHeight);
        double clipY = horizontal ? firstY : Math.Max(0d, firstY - (entries.Count - 1) * lineHeight);
        if (TryBuildManualLayoutBox(layout.Layout, frame, new ChartLayoutBox(x, clipY, width, clipHeight), out ChartLayoutBox manualBox))
        {
            x = manualBox.X;
            width = manualBox.Width;
            clipY = manualBox.Y;
            clipHeight = manualBox.Height;
            firstY = horizontal
                ? manualBox.Y
                : manualBox.Y + manualBox.Height - lineHeight;
        }

        RenderChartShapeStyle(graphics, x, clipY, width, clipHeight, layout.ShapeStyle);

        var runs = new List<TextRun>(entries.Count);
        for (int i = 0; i < entries.Count; i++)
        {
            ChartLegendEntry entry = entries[i];
            double entryX = horizontal ? GetPackedHorizontalLegendEntryX(entries, fontSize, markerSize, x, i) : x;
            double entryWidth = horizontal ? GetPackedHorizontalLegendEntryWidth(entries[i].Name, fontSize, markerSize) : width;
            double y = horizontal ? firstY : firstY - i * lineHeight;
            double markerBaselineFactor = horizontal || sideStrokeLegend
                ? PptxChartMetricRules.LegendHorizontalMarkerBaselineFactor
                : PptxChartMetricRules.LegendMarkerBaselineFactor;
            double markerY = y + lineHeight * markerBaselineFactor;
            if (entry.Fill is { } fill)
            {
                FillChartRectangle(graphics, entryX, markerY, markerSize, markerSize, fill);
            }
            else if (entry.Stroke is { } stroke)
            {
                SetChartStroke(graphics, stroke);
                graphics.StrokeLine(entryX, markerY + markerSize / 2d, entryX + markerWidth, markerY + markerSize / 2d);
            }

            runs.Add(new TextRun(
                entry.Name,
                entryX + markerWidth + textGap,
                y,
                Math.Max(1d, entryWidth - markerWidth - textGap),
                lineHeight,
                x,
                clipY,
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

    private static double GetPackedHorizontalLegendWidth(IReadOnlyList<ChartLegendEntry> entries, double fontSize, double markerSize)
    {
        double width = 0d;
        foreach (ChartLegendEntry entry in entries)
        {
            width += GetPackedHorizontalLegendEntryWidth(entry.Name, fontSize, markerSize);
        }

        return Math.Max(1d, width);
    }

    private static double GetFrameAnchoredSideFillLegendWidth(IReadOnlyList<ChartLegendEntry> entries, double fontSize, double markerWidth, double textGap)
    {
        double contentWidth = entries.Count == 0
            ? 0d
            : entries.Max(entry => markerWidth + textGap + EstimateChartTextWidth(entry.Name, fontSize));
        return Math.Max(fontSize * PptxChartMetricRules.LegendSideFillMinimumWidthFactor, contentWidth);
    }

    private static bool IsSameChartBox(ChartPlotBox plotBox, ChartFrameBox frame)
    {
        const double tolerance = 0.01d;
        return Math.Abs(plotBox.X - frame.X) <= tolerance &&
            Math.Abs(plotBox.Y - frame.Y) <= tolerance &&
            Math.Abs(plotBox.Width - frame.Width) <= tolerance &&
            Math.Abs(plotBox.Height - frame.Height) <= tolerance;
    }

    private static double GetPackedHorizontalLegendEntryX(IReadOnlyList<ChartLegendEntry> entries, double fontSize, double markerSize, double legendX, int entryIndex)
    {
        double x = legendX;
        for (int i = 0; i < entryIndex; i++)
        {
            x += GetPackedHorizontalLegendEntryWidth(entries[i].Name, fontSize, markerSize);
        }

        return x;
    }

    private static double GetPackedHorizontalLegendEntryWidth(string name, double fontSize, double markerSize)
    {
        return markerSize + PptxChartMetricRules.LegendTextGap + EstimateChartTextWidth(name, fontSize);
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

    private static IReadOnlyList<PdfFontResource> RenderPieDataLabels(PptxTheme theme, PdfGraphicsBuilder graphics, ChartPolarLayout layout, IReadOnlyList<double> values, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, ChartDataLabelOptions labelOptions)
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

        ChartPlotBox plotBox = layout.PlotBox;
        ChartPolarGeometry geometry = layout.Geometry;
        double labelRadius = geometry.Radius * (holeSize > 0d ? Math.Max(PptxChartMetricRules.PieDataLabelRadiusRatio, (1d + holeSize) / 2d) : PptxChartMetricRules.PieDataLabelRadiusRatio);
        double labelWidth = Math.Max(PptxChartMetricRules.PieDataLabelMinimumWidth, geometry.Radius * PptxChartMetricRules.PieDataLabelWidthRatio);
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
            double labelHeight = fontSize * PptxChartMetricRules.PieDataLabelHeightFactor;
            double sweep = values[i] / total * 360d;
            double mid = (angle + sweep / 2d) * Math.PI / 180d;
            double explosion = pointExplosions.TryGetValue(i, out double offset) ? Math.Clamp(offset, 0d, 1d) * geometry.Radius * PptxChartMetricRules.PieExplosionLabelRadiusRatio : 0d;
            double labelX = geometry.CenterX + Math.Cos(mid) * (labelRadius + explosion) - labelWidth / 2d;
            double labelY = geometry.CenterY + Math.Sin(mid) * (labelRadius + explosion) - labelHeight / 2d;
            string label = FormatPieDataLabel(values[i], total, effectiveOptions);
            if (!string.IsNullOrEmpty(label))
            {
                ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, labelWidth, labelHeight);
                RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                AddChartLabelRuns(runs, label, effectiveOptions, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, plotBox, style, TextAlignment.Center);
            }
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
        bool valueAxisReversed,
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
        double zeroX = ChartValueToPlotCoordinate(extents, 0d, plotBox.X, plotBox.Width, valueAxisReversed);
        double zeroY = ChartValueToPlotCoordinate(extents, 0d, plotBox.Y, plotBox.Height, valueAxisReversed);
        var runs = new List<TextRun>();
        if (horizontalBars)
        {
            double categoryHeight = plotBox.Height / categoryCount;
            double barSlot = categoryHeight * PptxChartMetricRules.BarDataLabelSlotFillRatio / Math.Max(1, series.Count);
            double labelWidth = Math.Max(PptxChartMetricRules.CartesianDataLabelMinimumWidth, plotBox.Width * PptxChartMetricRules.HorizontalBarDataLabelWidthRatio);
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryY = plotBox.Y + category * categoryHeight + categoryHeight * PptxChartMetricRules.BarDataLabelCategoryInsetRatio;
                for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    IReadOnlyList<double> values = series[seriesIndex];
                    if (category >= values.Count)
                    {
                        continue;
                    }

                    double value = values[category];
                    double barBase = zeroX;
                    double barEnd = ChartValueToPlotCoordinate(extents, value, plotBox.X, plotBox.Width, valueAxisReversed);
                    ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), category);
                    ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, effectiveOptions);
                    double fontSize = style.FontSize;
                    double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                    double x = ResolveHorizontalBarDataLabelX(effectiveOptions.PositionKind, barBase, barEnd, labelWidth);
                    double y = categoryY + seriesIndex * barSlot + barSlot * PptxChartMetricRules.HorizontalBarDataLabelSlotCenterRatio - labelHeight / 2d;
                    string label = FormatCartesianDataLabel(value, seriesIndex, category, effectiveOptions, categoryLabels, seriesNames);
                    if (!string.IsNullOrEmpty(label))
                    {
                        ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, x, y, labelWidth, labelHeight);
                        RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                        AddChartLabelRuns(runs, label, effectiveOptions, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, plotBox, style, TextAlignment.Left);
                    }
                }
            }
        }
        else
        {
            double categoryWidth = plotBox.Width / categoryCount;
            double barSlot = categoryWidth * PptxChartMetricRules.BarDataLabelSlotFillRatio / Math.Max(1, series.Count);
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryX = plotBox.X + category * categoryWidth + categoryWidth * PptxChartMetricRules.BarDataLabelCategoryInsetRatio;
                for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    IReadOnlyList<double> values = series[seriesIndex];
                    if (category >= values.Count)
                    {
                        continue;
                    }

                    double value = values[category];
                    double x = categoryX + seriesIndex * barSlot;
                    double barBase = zeroY;
                    double barEnd = ChartValueToPlotCoordinate(extents, value, plotBox.Y, plotBox.Height, valueAxisReversed);
                    ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), category);
                    ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, effectiveOptions);
                    double fontSize = style.FontSize;
                    double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                    double y = ResolveVerticalBarDataLabelY(effectiveOptions.PositionKind, barBase, barEnd, labelHeight);
                    string label = FormatCartesianDataLabel(value, seriesIndex, category, effectiveOptions, categoryLabels, seriesNames);
                    if (!string.IsNullOrEmpty(label))
                    {
                        double labelWidth = Math.Max(1d, barSlot * PptxChartMetricRules.VerticalBarDataLabelWidthRatio);
                        ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, x, y, labelWidth, labelHeight);
                        RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                        AddChartLabelRuns(runs, label, effectiveOptions, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, plotBox, style, TextAlignment.Center);
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
        bool valueAxisReversed,
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
        double labelWidth = Math.Max(
            PptxChartMetricRules.CartesianDataLabelMinimumWidth,
            plotBox.Width / Math.Max(PptxChartMetricRules.LineDataLabelMinimumPointSpan, pointCount * PptxChartMetricRules.LineDataLabelPointWidthFactor));
        var runs = new List<TextRun>();
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            for (int i = 0; i < values.Count; i++)
            {
                double pointX = plotBox.X + (pointCount == 1 ? plotBox.Width / 2d : plotBox.Width * i / (pointCount - 1));
                double pointY = ChartValueToPlotCoordinate(extents, values[i], plotBox.Y, plotBox.Height, valueAxisReversed);
                ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), i);
                ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, effectiveOptions);
                double fontSize = style.FontSize;
                double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                string label = FormatCartesianDataLabel(values[i], seriesIndex, i, effectiveOptions, categoryLabels, seriesNames);
                if (!string.IsNullOrEmpty(label))
                {
                    (double labelX, double labelY, TextAlignment alignment) = ResolveLineDataLabelPosition(
                        effectiveOptions.PositionKind,
                        pointX,
                        pointY,
                        labelWidth,
                        labelHeight);
                    ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, labelWidth, labelHeight);
                    RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                    AddChartLabelRuns(
                        runs,
                        label,
                        effectiveOptions,
                        labelBox.X,
                        labelBox.Y,
                        labelBox.Width,
                        labelBox.Height,
                        plotBox,
                        style,
                        alignment);
                }
            }
        }

        return RenderTextRuns(runs, graphics, "CLD");
    }

    private static ChartLayoutBox ResolveDataLabelBox(ChartPlotBox plotBox, ChartDataLabelOptions options, double x, double y, double width, double height)
    {
        ChartLayoutBox defaultBox = new(x, y, width, height);
        if (!options.Layout.HasLayout)
        {
            return defaultBox;
        }

        ChartFrameBox frame = new(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height);
        return TryBuildManualLayoutBox(options.Layout, frame, defaultBox, out ChartLayoutBox manualBox)
            ? manualBox
            : defaultBox;
    }

    private static double ResolveHorizontalBarDataLabelX(PptxSceneChartDataLabelPosition position, double barBase, double barEnd, double labelWidth)
    {
        bool extendsRight = barEnd >= barBase;
        return position switch
        {
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (barBase + barEnd - labelWidth) / 2d,
            PptxSceneChartDataLabelPosition.InsideBase => extendsRight ? barBase + PptxChartMetricRules.BarDataLabelHorizontalGap : barBase - labelWidth - PptxChartMetricRules.BarDataLabelHorizontalGap,
            PptxSceneChartDataLabelPosition.InsideEnd or PptxSceneChartDataLabelPosition.Left or PptxSceneChartDataLabelPosition.Right => extendsRight ? barEnd - labelWidth - PptxChartMetricRules.BarDataLabelHorizontalGap : barEnd + PptxChartMetricRules.BarDataLabelHorizontalGap,
            PptxSceneChartDataLabelPosition.OutsideEnd or _ => extendsRight ? barEnd + PptxChartMetricRules.BarDataLabelHorizontalGap : barEnd - labelWidth - PptxChartMetricRules.BarDataLabelHorizontalGap
        };
    }

    private static double ResolveVerticalBarDataLabelY(PptxSceneChartDataLabelPosition position, double barBase, double barEnd, double labelHeight)
    {
        bool extendsUp = barEnd >= barBase;
        return position switch
        {
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (barBase + barEnd - labelHeight) / 2d,
            PptxSceneChartDataLabelPosition.InsideBase => extendsUp ? barBase + PptxChartMetricRules.BarDataLabelVerticalGap : barBase - labelHeight - PptxChartMetricRules.BarDataLabelVerticalGap,
            PptxSceneChartDataLabelPosition.InsideEnd or PptxSceneChartDataLabelPosition.Top or PptxSceneChartDataLabelPosition.Bottom => extendsUp ? barEnd - labelHeight - PptxChartMetricRules.BarDataLabelVerticalGap : barEnd + PptxChartMetricRules.BarDataLabelVerticalGap,
            PptxSceneChartDataLabelPosition.OutsideEnd or _ => extendsUp ? barEnd + PptxChartMetricRules.BarDataLabelVerticalGap : barEnd - labelHeight - PptxChartMetricRules.BarDataLabelVerticalGap
        };
    }

    private static (double X, double Y, TextAlignment Alignment) ResolveLineDataLabelPosition(PptxSceneChartDataLabelPosition position, double pointX, double pointY, double labelWidth, double labelHeight)
    {
        return position switch
        {
            PptxSceneChartDataLabelPosition.Bottom => (pointX - labelWidth / 2d, pointY - labelHeight * PptxChartMetricRules.LineDataLabelBelowOffsetFactor, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Left => (pointX - labelWidth - PptxChartMetricRules.LineDataLabelSideGap, pointY - labelHeight / 2d, TextAlignment.Right),
            PptxSceneChartDataLabelPosition.Right => (pointX + PptxChartMetricRules.LineDataLabelSideGap, pointY - labelHeight / 2d, TextAlignment.Left),
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (pointX - labelWidth / 2d, pointY - labelHeight / 2d, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Top or PptxSceneChartDataLabelPosition.OutsideEnd or _ => (pointX - labelWidth / 2d, pointY + labelHeight * PptxChartMetricRules.LineDataLabelAboveOffsetFactor, TextAlignment.Center)
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
            LeaderLines = dataLabel.LeaderLines.IsDefined ? dataLabel.LeaderLines : options.LeaderLines,
            CustomText = string.IsNullOrEmpty(dataLabel.CustomText) ? options.CustomText : dataLabel.CustomText,
            CustomTextRuns = dataLabel.CustomTextRuns.Count == 0 ? options.CustomTextRuns : dataLabel.CustomTextRuns,
            PositionKind = string.IsNullOrEmpty(dataLabel.Position) ? options.PositionKind : dataLabel.PositionKind,
            Position = string.IsNullOrEmpty(dataLabel.Position) ? options.Position : dataLabel.Position,
            Separator = string.IsNullOrEmpty(dataLabel.Separator) ? options.Separator : dataLabel.Separator,
            NumberFormat = string.IsNullOrEmpty(dataLabel.NumberFormat) ? options.NumberFormat : dataLabel.NumberFormat,
            NumberFormatInfo = dataLabel.NumberFormatInfo.IsDefined ? dataLabel.NumberFormatInfo : options.NumberFormatInfo,
            Layout = dataLabel.Layout.HasLayout ? dataLabel.Layout : options.Layout,
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
        return CreateChartTextRun(text, x, y, width, height, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, style, alignment);
    }

    private static void AddChartLabelRuns(List<TextRun> runs, string text, ChartDataLabelOptions options, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment)
    {
        if (options.CustomTextRuns.Count == 0)
        {
            runs.Add(CreateChartLabelRun(text, x, y, width, height, plotBox, style, alignment));
            return;
        }

        AddChartRichTextRuns(runs, options.CustomTextRuns, text, x, y, width, height, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, style, alignment);
    }

    private static void AddChartRichTextRuns(List<TextRun> runs, IReadOnlyList<ChartTextRunOverride> richTextRuns, string fallbackText, double x, double y, double width, double height, double clipX, double clipY, double clipWidth, double clipHeight, ChartTextStyle style, TextAlignment alignment)
    {
        if (richTextRuns.Count == 0)
        {
            runs.Add(CreateChartTextRun(fallbackText, x, y, width, height, clipX, clipY, clipWidth, clipHeight, style, alignment));
            return;
        }

        ChartTextRunLayout[] richRuns = richTextRuns
            .Where(run => !string.IsNullOrEmpty(run.Text))
            .Select(run =>
            {
                ChartTextStyle runStyle = MergeChartTextStyle(style, run.TextStyle);
                return new ChartTextRunLayout(run.Text, runStyle, Math.Max(0d, EstimateChartTextWidth(run.Text, runStyle.FontSize)));
            })
            .Where(run => run.Width > 0d)
            .ToArray();
        if (richRuns.Length == 0)
        {
            runs.Add(CreateChartTextRun(fallbackText, x, y, width, height, clipX, clipY, clipWidth, clipHeight, style, alignment));
            return;
        }

        double totalWidth = richRuns.Sum(run => run.Width);
        double cursor = alignment switch
        {
            TextAlignment.Right => x + Math.Max(1d, width) - totalWidth,
            TextAlignment.Center => x + (Math.Max(1d, width) - totalWidth) / 2d,
            _ => x
        };

        foreach (ChartTextRunLayout run in richRuns)
        {
            double runWidth = Math.Max(0.1d, run.Width);
            runs.Add(CreateChartTextRun(run.Text, cursor, y, runWidth, height, clipX, clipY, clipWidth, clipHeight, run.Style, TextAlignment.Left) with { PreventCoalesce = true });
            cursor += runWidth;
        }
    }

    private static TextRun CreateChartTextRun(string text, double x, double y, double width, double height, double clipX, double clipY, double clipWidth, double clipHeight, ChartTextStyle style, TextAlignment alignment)
    {
        return new TextRun(
            text,
            x,
            y,
            Math.Max(1d, width),
            height,
            clipX,
            clipY,
            clipWidth,
            clipHeight,
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
                ReadChartDataLabelLeaderLines(labels, theme),
                string.Empty,
                [],
                PptxSceneBuilder.ParseChartDataLabelPosition(labels.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty),
                labels.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty,
                labels.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                labels.Element(ChartNamespace + "numFmt")?.Attribute("formatCode")?.Value ?? string.Empty,
                ReadChartNumberFormat(labels),
                default,
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
                plot.DataLabels.ShowValue == true,
                plot.DataLabels.ShowPercent == true,
                plot.DataLabels.ShowCategoryName == true,
                plot.DataLabels.ShowSeriesName == true,
                plot.DataLabels.ShowLeaderLines == true,
                ToChartDataLabelLeaderLines(plot.DataLabels.LeaderLines),
                string.Empty,
                [],
                plot.DataLabels.PositionKind,
                plot.DataLabels.Position,
                plot.DataLabels.Separator,
                plot.DataLabels.NumberFormat,
                ToChartNumberFormat(plot.DataLabels.NumberFormatInfo),
                default,
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
            labels.ShowValue == true,
            labels.ShowPercent == true,
            labels.ShowCategoryName == true,
            labels.ShowSeriesName == true,
            labels.ShowLeaderLines == true,
            ToChartDataLabelLeaderLines(labels.LeaderLines),
            string.Empty,
            [],
            labels.PositionKind,
            labels.Position,
            labels.Separator,
            labels.NumberFormat,
            ToChartNumberFormat(labels.NumberFormatInfo),
            default,
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
                ChartDataLabelLeaderLines.Empty,
                ReadChartText(label.Element(ChartNamespace + "tx")) ?? string.Empty,
                ToChartTextRuns(PptxSceneBuilder.ReadChartTextRuns(label.Element(ChartNamespace + "tx"), theme)),
                PptxSceneBuilder.ParseChartDataLabelPosition(label.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty),
                label.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty,
                label.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                label.Element(ChartNamespace + "numFmt")?.Attribute("formatCode")?.Value ?? string.Empty,
                ReadChartNumberFormat(label),
                ReadManualLayout(label),
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
                ChartDataLabelLeaderLines.Empty,
                dataLabel.CustomText,
                ToChartTextRuns(dataLabel.CustomTextRuns),
                dataLabel.PositionKind,
                dataLabel.Position,
                dataLabel.Separator,
                dataLabel.NumberFormat,
                ToChartNumberFormat(dataLabel.NumberFormatInfo),
                dataLabel.Layout,
                ToChartTextStyleOverride(dataLabel.TextStyle),
                ToChartShapeStyle(dataLabel.ShapeStyle));
        }

        return result;
    }

    private static ChartNumberFormat ReadChartNumberFormat(XElement parent)
    {
        XElement? numberFormat = parent.Element(ChartNamespace + "numFmt");
        return numberFormat is null
            ? default
            : new ChartNumberFormat(
                IsDefined: true,
                FormatCode: (string?)numberFormat.Attribute("formatCode") ?? string.Empty,
                SourceLinked: numberFormat.Attribute("sourceLinked") is { } sourceLinked
                    ? IsOoxmlTrue(sourceLinked.Value)
                    : null);
    }

    private static ChartNumberFormat ToChartNumberFormat(PptxSceneChartNumberFormat numberFormat)
    {
        return numberFormat.IsDefined
            ? new ChartNumberFormat(numberFormat.IsDefined, numberFormat.FormatCode, numberFormat.SourceLinked)
            : default;
    }

    private static ChartDataLabelLeaderLines ReadChartDataLabelLeaderLines(XElement labels, PptxTheme theme)
    {
        XElement? leaderLines = labels.Element(ChartNamespace + "leaderLines");
        if (leaderLines is null)
        {
            return ChartDataLabelLeaderLines.Empty;
        }

        return new ChartDataLabelLeaderLines(IsDefined: true, ReadChartShapeStyle(leaderLines.Element(ChartNamespace + "spPr"), theme).Stroke);
    }

    private static ChartDataLabelLeaderLines ToChartDataLabelLeaderLines(PptxSceneChartLeaderLines leaderLines)
    {
        return leaderLines.IsDefined
            ? new ChartDataLabelLeaderLines(IsDefined: true, ToChartSeriesStroke(leaderLines.Line))
            : ChartDataLabelLeaderLines.Empty;
    }

    private static IReadOnlyList<ChartTextRunOverride> ToChartTextRuns(IReadOnlyList<PptxSceneChartTextRun> runs)
    {
        return runs.Count == 0
            ? []
            : runs.Select(run => new ChartTextRunOverride(run.Text, ToChartTextStyleOverride(run.TextStyle))).ToArray();
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

    private static IReadOnlyList<PdfFontResource> RenderChartCategoryLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartAxis? sceneAxis, XElement? categoryAxis, IReadOnlyList<string> labels, bool horizontalBars, double? verticalAxisY = null)
    {
        if (labels.Count == 0)
        {
            return [];
        }

        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, categoryAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize);
        double fontSize = style.FontSize;
        RgbColor color = style.Color;
        double labelOffsetScale = ResolveSceneOrXmlCategoryAxisLabelOffsetScale(sceneAxis, categoryAxis);
        int tickLabelSkip = ResolveSceneOrXmlCategoryAxisTickLabelSkip(sceneAxis, categoryAxis);
        var runs = new List<TextRun>(labels.Count);
        for (int i = 0; i < labels.Count; i++)
        {
            if (i % tickLabelSkip != 0)
            {
                continue;
            }

            double x;
            double y;
            double width;
            double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
            TextAlignment alignment;
            if (horizontalBars)
            {
                double slotHeight = plotBox.Height / labels.Count;
                x = Math.Max(0d, plotBox.X - plotBox.Width * PptxChartMetricRules.CategoryAxisHorizontalLeftOffsetRatio * labelOffsetScale);
                y = plotBox.Y + slotHeight * (i + 0.5d) - height * PptxChartMetricRules.CategoryAxisHorizontalBaselineRatio;
                width = plotBox.Width * PptxChartMetricRules.CategoryAxisHorizontalWidthRatio;
                alignment = TextAlignment.Right;
            }
            else
            {
                double slotWidth = plotBox.Width / labels.Count;
                width = slotWidth * PptxChartMetricRules.CategoryAxisVerticalWidthFactor;
                x = plotBox.X + slotWidth * (i + 0.5d) - width / 2d;
                double axisY = verticalAxisY ?? plotBox.Y;
                y = axisY - height * PptxChartMetricRules.CategoryAxisVerticalTopOffsetFactor * labelOffsetScale;
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
                y - height * PptxChartMetricRules.AxisLabelClipTopOffsetFactor,
                labelWidth,
                height * PptxChartMetricRules.AxisLabelClipHeightFactor,
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

    private static IReadOnlyList<PdfFontResource> RenderChartValueAxisLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, XElement? valueAxis, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits axisUnits, bool valueAxisReversed, bool horizontalBars, bool rightSide = false, int axisSideSlot = 0, bool manualPlotLayoutApplied = false, bool useTextSizedWidth = false, string? defaultNumberFormat = null)
    {
        double range = Math.Max(1d, extents.Max - extents.Min);
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize);
        double fontSize = style.FontSize;
        double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
        RgbColor color = style.Color;
        double autoTickTargetCount = GetValueAxisAutoTickTargetCount(horizontalBars, valueAxisLabelsVisible: true, manualPlotLayoutApplied);
        IReadOnlyList<double> tickValues = GetChartAxisTickValues(extents, axisUnits.MajorUnit, includeEndpoints: true, autoTickTargetCount);
        double maxLabelWidth = tickValues
            .Select(value => FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis, defaultNumberFormat))
            .DefaultIfEmpty("0")
            .Max(label => EstimateChartTextWidth(label, fontSize));
        double valueAxisLabelWidth = Math.Max(
            fontSize * PptxChartMetricRules.ValueAxisMinimumLabelWidthFactor,
            maxLabelWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
        var runs = new List<TextRun>(tickValues.Count);
        foreach (double value in tickValues)
        {
            string label = FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis, defaultNumberFormat);
            double offset = GetChartValuePlotRatio(extents, value, valueAxisReversed);
            double x;
            double y;
            double width;
            TextAlignment alignment;
            if (horizontalBars)
            {
                width = plotBox.Width / PptxChartMetricRules.HorizontalValueAxisSlotCount;
                x = plotBox.X + plotBox.Width * offset - width / 2d;
                y = plotBox.Y - height * PptxChartMetricRules.HorizontalValueAxisTopOffsetFactor;
                alignment = TextAlignment.Center;
            }
            else
            {
                bool labelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(sceneAxis, valueAxis, rightSide);
                width = useTextSizedWidth ? valueAxisLabelWidth : plotBox.Width * PptxChartMetricRules.VerticalValueAxisWidthRatio;
                double sideGap = Math.Max(3d, fontSize * PptxChartMetricRules.ValueAxisLabelSideGapFactor);
                x = labelsRightSide
                    ? plotBox.X + plotBox.Width + sideGap + axisSideSlot * (width + sideGap)
                    : Math.Max(0d, plotBox.X - (axisSideSlot + 1) * (width + sideGap));
                y = plotBox.Y + plotBox.Height * offset - height * PptxChartMetricRules.VerticalValueAxisBaselineRatio;
                alignment = labelsRightSide ? TextAlignment.Left : TextAlignment.Right;
            }

            double labelWidth = Math.Max(1d, width);
            double clipY = horizontalBars ? y : plotBox.Y - height;
            double clipHeight = horizontalBars ? height : plotBox.Height + height * PptxChartMetricRules.ValueAxisVerticalClipExtraHeightFactor;
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

    private static double GetValueAxisAutoTickTargetCount(bool horizontalBars, bool valueAxisLabelsVisible, bool manualPlotLayoutApplied)
    {
        if (!horizontalBars)
        {
            return PptxChartMetricRules.AxisNiceTickTargetCount;
        }

        return valueAxisLabelsVisible && manualPlotLayoutApplied
            ? PptxChartMetricRules.AxisNiceTickTargetCount
            : PptxChartMetricRules.AxisNiceHorizontalValueTickTargetCount;
    }

    private static double EstimateChartTextWidth(string text, double fontSize)
    {
        double width = 0d;
        foreach (char ch in text)
        {
            width += char.IsWhiteSpace(ch)
                ? fontSize * PptxChartMetricRules.ValueAxisLabelWhitespaceWidthFactor
                : char.IsDigit(ch)
                    ? fontSize * PptxChartMetricRules.ValueAxisLabelDigitWidthFactor
                    : fontSize * PptxChartMetricRules.ValueAxisLabelCharacterWidthFactor;
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
        return RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, rightValueAxis, rightSceneAxis, extents, axisUnits, ReadSceneOrXmlValueAxisReversed(rightSceneAxis, rightValueAxis), horizontalBars: false, rightSide: true);
    }

    private static XElement? ReadSecondaryRightValueAxis(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "valAx")
            .Where(IsRightValueAxis)
            .FirstOrDefault();
    }

    private static XElement? ReadSecondaryValueAxisForChart(XDocument chartXml, XElement? primaryValueAxis)
    {
        string? primaryAxisId = ReadChartAxisId(primaryValueAxis);
        if (string.IsNullOrWhiteSpace(primaryAxisId))
        {
            return ReadSecondaryRightValueAxis(chartXml);
        }

        return chartXml
            .Descendants(ChartNamespace + "valAx")
            .FirstOrDefault(axis =>
                !string.Equals(ReadChartAxisId(axis), primaryAxisId, StringComparison.Ordinal) &&
                !IsChartAxisDeleted(axis));
    }

    private static PptxSceneChartAxis? ReadSceneSecondaryValueAxis(PptxSceneChart? sceneChart, XElement? secondaryValueAxis, PptxSceneChartAxis? primaryValueAxis)
    {
        if (sceneChart is null)
        {
            return null;
        }

        string? axisId = ReadChartAxisId(secondaryValueAxis);
        if (!string.IsNullOrWhiteSpace(axisId))
        {
            PptxSceneChartAxis? matchingAxis = sceneChart.Axes.FirstOrDefault(axis =>
                string.Equals(axis.Id, axisId, StringComparison.Ordinal) &&
                axis.AxisKind == PptxSceneChartAxisKind.Value);
            if (matchingAxis is not null)
            {
                return matchingAxis;
            }
        }

        if (primaryValueAxis is not null)
        {
            return sceneChart.Axes.FirstOrDefault(axis =>
                axis.AxisKind == PptxSceneChartAxisKind.Value &&
                !string.Equals(axis.Id, primaryValueAxis.Id, StringComparison.Ordinal) &&
                axis.IsDeleted != true);
        }

        return ReadSceneSecondaryRightValueAxis(sceneChart, secondaryValueAxis);
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
                axis.AxisKind == PptxSceneChartAxisKind.Value);
            if (matchingAxis is not null)
            {
                return matchingAxis;
            }
        }

        return sceneChart.Axes.FirstOrDefault(axis =>
            axis.AxisKind == PptxSceneChartAxisKind.Value &&
            axis.PositionKind == PptxSceneChartAxisPosition.Right);
    }

    private static bool IsSceneOrXmlVisibleValueAxis(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.AxisKind == PptxSceneChartAxisKind.Value &&
                sceneAxis.IsDeleted != true;
        }

        return axis is not null &&
            !IsOoxmlBooleanElementEnabled(axis.Element(ChartNamespace + "delete"));
    }

    private static bool IsRightValueAxis(XElement? axis)
    {
        return axis is not null &&
            PptxSceneBuilder.ParseChartAxisPosition((string?)axis.Element(ChartNamespace + "axPos")?.Attribute("val")) == PptxSceneChartAxisPosition.Right &&
            !IsOoxmlBooleanElementEnabled(axis.Element(ChartNamespace + "delete"));
    }

    private static string? ReadChartAxisId(XElement? axis)
    {
        return (string?)axis?.Element(ChartNamespace + "axId")?.Attribute("val");
    }

    private static IReadOnlyList<double> GetChartAxisTickValues(ChartValueExtents extents, double? explicitUnit, bool includeEndpoints, double autoTickTargetCount = PptxChartMetricRules.AxisNiceTickTargetCount)
    {
        double range = Math.Max(1d, extents.Max - extents.Min);
        if (explicitUnit is not { } unit || unit <= 0d)
        {
            unit = ChooseChartAxisMajorUnit(range, autoTickTargetCount);
        }

        var values = new List<double>();
        if (includeEndpoints)
        {
            values.Add(extents.Min);
        }

        double first = Math.Ceiling(extents.Min / unit) * unit;
        for (double value = first; value < extents.Max - PptxChartMetricRules.AxisValueEpsilon; value += unit)
        {
            if (value > extents.Min + PptxChartMetricRules.AxisValueEpsilon)
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

    private static IReadOnlyList<double> GetChartGridlineValues(ChartValueExtents extents, double? explicitUnit, double? crossingValue, double autoTickTargetCount = PptxChartMetricRules.AxisNiceTickTargetCount)
    {
        return GetChartAxisTickValues(extents, explicitUnit, includeEndpoints: true, autoTickTargetCount)
            .Where(value => crossingValue is not { } crossing || Math.Abs(value - crossing) > PptxChartMetricRules.AxisValueEpsilon)
            .ToArray();
    }

    private static double? ReadSceneOrXmlValueAxisCrossingValue(PptxSceneChartAxis? sceneAxis, XElement? valueAxis, ChartValueExtents extents)
    {
        double? crossesAt = sceneAxis?.CrossesAt ?? ReadChartElementDouble(valueAxis, "crossesAt");
        if (crossesAt is { } explicitCrossing)
        {
            return explicitCrossing;
        }

        PptxSceneChartAxisCrosses crosses = sceneAxis is not null
            ? sceneAxis.CrossesKind
            : PptxSceneBuilder.ParseChartAxisCrosses((string?)valueAxis?.Element(ChartNamespace + "crosses")?.Attribute("val") ?? "autoZero");
        if (crosses == PptxSceneChartAxisCrosses.Maximum)
        {
            return extents.Max;
        }

        if (crosses == PptxSceneChartAxisCrosses.Minimum)
        {
            return extents.Min;
        }

        if (extents.Min <= 0d && extents.Max >= 0d)
        {
            return 0d;
        }

        return extents.Min > 0d ? extents.Min : extents.Max;
    }

    private static bool ReadSceneOrXmlValueAxisReversed(PptxSceneChartAxis? sceneAxis, XElement? valueAxis)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.OrientationKind == PptxSceneChartAxisOrientation.MaximumMinimum;
        }

        return PptxSceneBuilder.ParseChartAxisOrientation(
            (string?)valueAxis?.Element(ChartNamespace + "scaling")?.Element(ChartNamespace + "orientation")?.Attribute("val")) ==
            PptxSceneChartAxisOrientation.MaximumMinimum;
    }

    private static double GetChartValuePlotRatio(ChartValueExtents extents, double value, bool reversed)
    {
        double range = Math.Max(1d, extents.Max - extents.Min);
        double boundedValue = Math.Clamp(value, extents.Min, extents.Max);
        double ratio = (boundedValue - extents.Min) / range;
        return reversed ? 1d - ratio : ratio;
    }

    private static double ChartValueToPlotCoordinate(ChartValueExtents extents, double? value, double plotStart, double plotLength, bool reversed)
    {
        return plotStart + plotLength * GetChartValuePlotRatio(extents, value ?? 0d, reversed);
    }

    private static double ChooseChartAxisMajorUnit(double range, double tickTargetCount = PptxChartMetricRules.AxisNiceTickTargetCount)
    {
        double target = Math.Max(range / Math.Max(1d, tickTargetCount), double.Epsilon);
        double magnitude = Math.Pow(PptxChartMetricRules.AxisNiceTickStepMaximum, Math.Floor(Math.Log10(target)));
        double normalized = target / magnitude;
        double nice = normalized <= PptxChartMetricRules.AxisNiceTickStepSmall
            ? PptxChartMetricRules.AxisNiceTickStepSmall
            : normalized <= PptxChartMetricRules.AxisNiceTickStepMedium
                ? PptxChartMetricRules.AxisNiceTickStepMedium
                : normalized <= PptxChartMetricRules.AxisNiceTickStepLarge
                    ? PptxChartMetricRules.AxisNiceTickStepLarge
                    : PptxChartMetricRules.AxisNiceTickStepMaximum;
        return nice * magnitude;
    }

    private static double GetNiceChartAxisMin(double dataMin, double dataMax)
    {
        if (dataMin >= 0d)
        {
            return dataMin;
        }

        double range = dataMax - dataMin;
        if (Math.Abs(range) < PptxChartMetricRules.AxisValueEpsilon)
        {
            return dataMin;
        }

        double unit = ChooseChartAxisMajorUnit(range);
        double niceMin = Math.Floor(dataMin / unit) * unit;
        return niceMin < dataMax ? niceMin : dataMin;
    }

    private static double GetNiceChartAxisMax(double dataMax, double dataMin, double tickTargetCount = PptxChartMetricRules.AxisNiceTickTargetCount, bool useNearMaximumHeadroom = false, double nearMaximumHeadroomRatio = PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio)
    {
        if (Math.Abs(dataMax) < PptxChartMetricRules.AxisValueEpsilon && Math.Abs(dataMin) < PptxChartMetricRules.AxisValueEpsilon)
        {
            return 1d;
        }

        double range = dataMax - Math.Min(0d, dataMin);
        if (Math.Abs(range) < PptxChartMetricRules.AxisValueEpsilon)
        {
            return dataMax > 0d ? dataMax * PptxChartMetricRules.AxisSingleValueHeadroomFactor : 1d;
        }

        double unit = ChooseChartAxisMajorUnit(range, tickTargetCount);
        double niceMax = Math.Ceiling(dataMax / unit) * unit;
        if (niceMax < dataMax + PptxChartMetricRules.AxisValueEpsilon)
        {
            return niceMax + unit;
        }

        if (useNearMaximumHeadroom &&
            niceMax > 0d &&
            dataMax > 0d &&
            dataMax / niceMax >= nearMaximumHeadroomRatio)
        {
            return niceMax + unit;
        }

        return niceMax;
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
        return Math.Abs(value - rounded) < PptxChartMetricRules.AxisValueEpsilon
            ? rounded.ToString("0", CultureInfo.InvariantCulture)
            : value.ToString("0.##", CultureInfo.InvariantCulture);
    }

    private static string FormatSceneOrXmlChartAxisLabel(double value, PptxSceneChartAxis? sceneAxis, XElement? axis, string? defaultNumberFormat = null)
    {
        string? formatCode = sceneAxis?.NumberFormat;
        if (!string.IsNullOrWhiteSpace(formatCode) &&
            !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase))
        {
            return FormatChartNumber(value, formatCode);
        }

        formatCode = (string?)axis?
            .Element(ChartNamespace + "numFmt")
            ?.Attribute("formatCode");
        if (!string.IsNullOrWhiteSpace(formatCode) &&
            !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase))
        {
            return FormatChartNumber(value, formatCode);
        }

        return !string.IsNullOrWhiteSpace(defaultNumberFormat)
            ? FormatChartNumber(value, defaultNumberFormat)
            : FormatChartAxisLabel(value);
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

        if (series.Element(ChartNamespace + "explosion")?.Attribute("val") is { } seriesExplosionAttribute &&
            double.TryParse(seriesExplosionAttribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double seriesExplosion))
        {
            double fraction = Math.Clamp(seriesExplosion / 100d, 0d, 1d);
            int pointCount = series
                .Elements(ChartNamespace + "val")
                .Descendants(ChartNamespace + "pt")
                .Select(point => (string?)point.Attribute("idx"))
                .Select(value => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) ? index : -1)
                .Where(index => index >= 0)
                .DefaultIfEmpty(-1)
                .Max() + 1;
            for (int index = 0; index < pointCount; index++)
            {
                explosions[index] = fraction;
            }
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

    private static IReadOnlyList<ChartMarkerStyle> ReadChartMarkerStyles(XElement chartElement, PptxTheme theme, PptxSceneChartPlotKind plotKind)
    {
        var styles = new List<ChartMarkerStyle>();
        bool chartMarkerEnabled = IsOoxmlBooleanElementEnabled(chartElement.Element(ChartNamespace + "marker"));
        bool lineChart = plotKind == PptxSceneChartPlotKind.Line;
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? marker = element.Element(ChartNamespace + "marker");
            string symbol = (string?)marker?.Element(ChartNamespace + "symbol")?.Attribute("val") ??
                (lineChart
                    ? chartMarkerEnabled ? AutoLineChartMarkerSymbol(styles.Count) : "none"
                    : "circle");
            double size = marker?.Element(ChartNamespace + "size")?.Attribute("val") is { } value &&
                double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                    ? Math.Clamp(parsed, 2d, 30d)
                    : lineChart && chartMarkerEnabled && marker is null ? 5d : 4d;
            XElement? shapeProperties = marker?.Element(ChartNamespace + "spPr");
            ChartSeriesFill? fill = TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
                ? new ChartSeriesFill(fillColor, fillAlpha)
                : null;
            ChartSeriesStroke? stroke = shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, out RgbColor strokeColor, out double strokeWidth, out double strokeAlpha)
                    ? new ChartSeriesStroke(strokeColor, strokeAlpha, strokeWidth)
                    : null;
            styles.Add(new ChartMarkerStyle(PptxSceneBuilder.ParseChartMarkerSymbol(symbol), symbol, size, fill, stroke));
        }

        return styles;
    }

    private static string AutoLineChartMarkerSymbol(int seriesIndex)
    {
        ReadOnlySpan<string> symbols = ["diamond", "square", "triangle", "x", "star", "circle"];
        return symbols[seriesIndex % symbols.Length];
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

    private static bool HasMajorGridlines(XElement? axis)
    {
        return IsChartGridlineVisible(axis?.Element(ChartNamespace + "majorGridlines"));
    }

    private static bool HasMinorGridlines(XElement? axis)
    {
        return IsChartGridlineVisible(axis?.Element(ChartNamespace + "minorGridlines"));
    }

    private static bool ReadSceneOrXmlMajorGridlines(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        return sceneAxis?.HasMajorGridlines ?? HasMajorGridlines(axis);
    }

    private static bool ReadSceneOrXmlMinorGridlines(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        return sceneAxis?.HasMinorGridlines ?? HasMinorGridlines(axis);
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

    private static bool IsChartGridlineVisible(XElement? gridlines)
    {
        if (gridlines is null)
        {
            return false;
        }

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
        XElement? secondaryValueAxisElement = ReadSecondaryValueAxisForChart(chartXml, valueAxisElement);
        ChartSeriesStroke? valueAxis = ReadChartAxisStroke(valueAxisElement, theme);
        ChartSeriesStroke? categoryAxis = ReadChartAxisStroke(categoryAxisElement, theme);
        return new ChartAxesStyle(
            valueAxis,
            ReadChartAxisStroke(secondaryValueAxisElement, theme),
            categoryAxis,
            ResolveSceneOrXmlValueAxisRightSide(null, valueAxisElement, defaultRightSide: false),
            ResolveSceneOrXmlValueAxisRightSide(null, secondaryValueAxisElement, defaultRightSide: true),
            ResolveSceneOrXmlValueAxisBottomSide(null, valueAxisElement, defaultBottomSide: true),
            ResolveSceneOrXmlCategoryAxisRightSide(null, categoryAxisElement, defaultRightSide: false),
            !IsChartAxisDeleted(valueAxisElement),
            !IsChartAxisDeleted(categoryAxisElement),
            ReadSceneOrXmlChartAxisMajorTickMark(null, categoryAxisElement));
    }

    private static ChartAxesStyle ReadSceneOrXmlChartAxesStyle(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, PptxTheme theme, XElement chartElement)
    {
        XElement? valueAxisElement = ReadChartValueAxisForChart(chartXml, chartElement) ??
            chartXml.Descendants(ChartNamespace + "valAx").FirstOrDefault();
        XElement? categoryAxisElement = ReadChartCategoryAxisForChart(chartXml, chartElement);
        XElement? secondaryValueAxisElement = ReadSecondaryValueAxisForChart(chartXml, valueAxisElement);
        PptxSceneChartAxis? valueAxis = ReadSceneChartAxis(sceneChart, plot, PptxSceneChartAxisKind.Value);
        PptxSceneChartAxis? categoryAxis = ReadSceneChartAxis(sceneChart, plot, PptxSceneChartAxisKind.Category);
        PptxSceneChartAxis? secondaryValueAxis = ReadSceneSecondaryValueAxis(sceneChart, secondaryValueAxisElement, valueAxis);
        return new ChartAxesStyle(
            ReadSceneOrXmlChartAxisStroke(valueAxis, valueAxisElement, theme),
            ReadSceneOrXmlChartAxisStroke(secondaryValueAxis, secondaryValueAxisElement, theme),
            ReadSceneOrXmlChartAxisStroke(categoryAxis, categoryAxisElement, theme),
            ResolveSceneOrXmlValueAxisRightSide(valueAxis, valueAxisElement, defaultRightSide: false),
            ResolveSceneOrXmlValueAxisRightSide(secondaryValueAxis, secondaryValueAxisElement, defaultRightSide: true),
            ResolveSceneOrXmlValueAxisBottomSide(valueAxis, valueAxisElement, defaultBottomSide: true),
            ResolveSceneOrXmlCategoryAxisRightSide(categoryAxis, categoryAxisElement, defaultRightSide: false),
            valueAxis is null ? !IsChartAxisDeleted(valueAxisElement) : valueAxis.IsDeleted != true,
            categoryAxis is null ? !IsChartAxisDeleted(categoryAxisElement) : categoryAxis.IsDeleted != true,
            ReadSceneOrXmlChartAxisMajorTickMark(categoryAxis, categoryAxisElement));
    }

    private static ChartSeriesStroke? ReadSceneOrXmlChartAxisStroke(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis, PptxTheme theme)
    {
        return sceneAxis is not null && sceneAxis.Line.HasLine
            ? ToChartSeriesStroke(sceneAxis.Line)
            : ReadChartAxisStroke(xmlAxis, theme);
    }

    private static PptxSceneChartAxisTickMark ReadSceneOrXmlChartAxisMajorTickMark(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.MajorTickMarkKind;
        }

        return PptxSceneBuilder.ParseChartAxisTickMark((string?)xmlAxis?.Element(ChartNamespace + "majorTickMark")?.Attribute("val") ?? "none");
    }

    private static bool ResolveSceneOrXmlValueAxisRightSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Right => true,
                PptxSceneChartAxisPosition.Left => false,
                _ => defaultRightSide
            };
        }

        string? position = (string?)axis?.Element(ChartNamespace + "axPos")?.Attribute("val");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Right => true,
            PptxSceneChartAxisPosition.Left => false,
            _ => defaultRightSide
        };
    }

    private static bool ResolveSceneOrXmlValueAxisBottomSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultBottomSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Bottom => true,
                PptxSceneChartAxisPosition.Top => false,
                _ => defaultBottomSide
            };
        }

        string? position = (string?)axis?.Element(ChartNamespace + "axPos")?.Attribute("val");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Bottom => true,
            PptxSceneChartAxisPosition.Top => false,
            _ => defaultBottomSide
        };
    }

    private static bool ResolveSceneOrXmlCategoryAxisRightSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Right => true,
                PptxSceneChartAxisPosition.Left => false,
                _ => defaultRightSide
            };
        }

        string? position = (string?)axis?.Element(ChartNamespace + "axPos")?.Attribute("val");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Right => true,
            PptxSceneChartAxisPosition.Left => false,
            _ => defaultRightSide
        };
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
        return PptxSceneBuilder.ParseChartTickLabelPosition(tickLabelPosition) != PptxSceneChartTickLabelPosition.None;
    }

    private static bool IsSceneOrXmlChartAxisLabelVisible(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is null)
        {
            return IsChartAxisLabelVisible(axis);
        }

        return sceneAxis.IsDeleted != true &&
            sceneAxis.TickLabelPositionKind != PptxSceneChartTickLabelPosition.None;
    }

    private static bool ResolveValueAxisLabelsRightSide(XElement? axis, bool defaultRightSide)
    {
        string? tickLabelPosition = (string?)axis?.Element(ChartNamespace + "tickLblPos")?.Attribute("val");
        return PptxSceneBuilder.ParseChartTickLabelPosition(tickLabelPosition) switch
        {
            PptxSceneChartTickLabelPosition.High => true,
            PptxSceneChartTickLabelPosition.Low => false,
            _ => defaultRightSide
        };
    }

    private static bool ResolveSceneOrXmlValueAxisLabelsRightSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultRightSide)
    {
        if (sceneAxis is null)
        {
            return ResolveValueAxisLabelsRightSide(axis, defaultRightSide);
        }

        return sceneAxis.TickLabelPositionKind switch
        {
            PptxSceneChartTickLabelPosition.High => true,
            PptxSceneChartTickLabelPosition.Low => false,
            _ => defaultRightSide
        };
    }

    private static double ResolveSceneOrXmlCategoryAxisLabelOffsetScale(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        int offset = sceneAxis?.LabelOffset ??
            ReadChartElementInt(axis, "lblOffset") ??
            PptxChartMetricRules.CategoryAxisDefaultLabelOffset;

        offset = Math.Clamp(
            offset,
            PptxChartMetricRules.CategoryAxisMinimumLabelOffset,
            PptxChartMetricRules.CategoryAxisMaximumLabelOffset);

        return offset / (double)PptxChartMetricRules.CategoryAxisDefaultLabelOffset;
    }

    private static int ResolveSceneOrXmlCategoryAxisTickLabelSkip(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        int skip = sceneAxis?.TickLabelSkip ??
            ReadChartElementInt(axis, "tickLblSkip") ??
            PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip;

        return Math.Max(PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip, skip);
    }

    private static int? ReadChartElementInt(XElement? chartElement, string elementName)
    {
        if (chartElement?.Element(ChartNamespace + elementName)?.Attribute("val") is not { } value ||
            !int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
        {
            return null;
        }

        return parsed;
    }

    private static double? ReadChartElementDouble(XElement? chartElement, string elementName)
    {
        if (chartElement?.Element(ChartNamespace + elementName)?.Attribute("val") is not { } value ||
            !double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
        {
            return null;
        }

        return parsed;
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

    private static void RenderBarChart(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, ChartLayoutBox plotAreaBox, ChartPlotBox plotBox, IReadOnlyList<IReadOnlyList<double>> series, bool horizontalBars, PptxSceneChartGrouping grouping, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool majorGridlines, bool minorGridlines, ChartGridlineStyle gridlineStyle, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents, ChartAxisUnits axisUnits, double? valueAxisCrossingValue, bool valueAxisReversed, bool valueAxisLabelsVisible, bool manualPlotLayoutApplied, bool varyColors, double gapWidthPercent, double overlapPercent)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        graphics.SaveState();
        try
        {
            ClipChartPlotArea(graphics, plotX, plotY, plotWidth, plotHeight);
            int categoryCount = Math.Max(1, series.Max(values => values.Count));
            bool stacked = IsStackedChartGrouping(grouping);
            bool percentStacked = IsPercentStackedChartGrouping(grouping);
            double zeroX = ChartValueToPlotCoordinate(valueExtents, 0d, plotX, plotWidth, valueAxisReversed);
            double zeroY = ChartValueToPlotCoordinate(valueExtents, 0d, plotY, plotHeight, horizontalBars ? false : valueAxisReversed);
            double valueAxisCrossingY = ChartValueToPlotCoordinate(valueExtents, valueAxisCrossingValue, plotY, plotHeight, horizontalBars ? false : valueAxisReversed);
            double valueAxisAutoTickTargetCount = GetValueAxisAutoTickTargetCount(horizontalBars, valueAxisLabelsVisible, manualPlotLayoutApplied);
            if (minorGridlines)
            {
                if (horizontalBars)
                {
                    DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor);
                }
                else
                {
                    DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor);
                }
            }

            if (majorGridlines)
            {
                if (horizontalBars)
                {
                    DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major, valueAxisAutoTickTargetCount);
                }
                else
                {
                    DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major);
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
                    double axisY = horizontalBars
                        ? (axesStyle.ValueAxisBottomSide ? plotY : plotY + plotHeight)
                        : valueAxisCrossingY;
                    graphics.StrokeLine(plotX, axisY, plotX + plotWidth, axisY);
                }
            }

            if (axesStyle.ValueAxisVisible)
            {
                ChartSeriesStroke stroke = horizontalBars ? categoryAxisStroke : valueAxisStroke;
                if (stroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, stroke);
                    double axisX = horizontalBars
                        ? (axesStyle.CategoryAxisRightSide ? plotX + plotWidth : plotX)
                        : (axesStyle.ValueAxisRightSide ? plotX + plotWidth : plotX);
                    graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight);
                }
            }

            if (!horizontalBars && axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke)
            {
                if (secondaryValueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, secondaryValueAxisStroke);
                    double axisX = axesStyle.SecondaryValueAxisRightSide ? plotX + plotWidth : plotX;
                    graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight);
                }
            }

            if (horizontalBars)
            {
                if (stacked)
                {
                    RenderStackedHorizontalBars(graphics, theme, chartPalette, plotX, plotY, plotWidth, plotHeight, series, categoryCount, valueExtents, valueAxisReversed, percentStacked, seriesFills, pointFills, pointStrokes, varyColors, gapWidthPercent);
                }
                else
                {
                    RenderClusteredHorizontalBars(graphics, theme, chartPalette, plotX, plotY, plotWidth, plotHeight, series, categoryCount, valueExtents, valueAxisReversed, zeroX, seriesFills, pointFills, pointStrokes, varyColors, gapWidthPercent, overlapPercent);
                }

                return;
            }

            if (stacked)
            {
                RenderStackedColumns(graphics, theme, chartPalette, plotX, plotY, plotWidth, plotHeight, series, categoryCount, valueExtents, valueAxisReversed, percentStacked, seriesFills, pointFills, pointStrokes, varyColors, gapWidthPercent);
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
                    ChartSeriesFill fill = ResolveBarPointFill(theme, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, value);
                    double barX = categoryX + seriesIndex * step;
                    double valueY = ChartValueToPlotCoordinate(valueExtents, value, plotY, plotHeight, valueAxisReversed);
                    double barY = Math.Min(zeroY, valueY);
                    double barHeight = Math.Abs(valueY - zeroY);
                    FillChartRectangle(graphics, barX, barY, barWidth, barHeight, fill);
                    StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, barX, barY, barWidth, barHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, value));
                }
            }
        }
        finally
        {
            graphics.RestoreState();
        }
    }

    private static ChartLayout GetBarChartLayout(
        PptxDocument document,
        PptxTheme theme,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? barPlot,
        XElement barChart,
        bool horizontalBars)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, sceneChart, chartXml);
        ChartPlotLayout plotLayout = GetBarChartPlotLayout(theme, frame, chartXml, sceneChart, barPlot, barChart, title, legend, horizontalBars);
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, legend);
    }

    private static ChartPlotLayout GetBarChartPlotLayout(
        PptxTheme theme,
        ChartFrameBox frame,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? barPlot,
        XElement barChart,
        string? title,
        ChartLegendLayout legend,
        bool horizontalBars,
        bool ignoreManualPlotLayout = false)
    {
        bool hasTitle = !string.IsNullOrWhiteSpace(title);
        bool hasLegend = legend.Visible && !legend.Overlay;
        ChartPlotBox defaultPlotBox;
        if (!hasTitle && !hasLegend)
        {
            defaultPlotBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.BarOverlayOnly);
        }
        else if (!hasTitle && legend.PositionKind == PptxSceneChartLegendPosition.Bottom)
        {
            defaultPlotBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.BarNoTitleBottomLegend);
        }
        else if (horizontalBars && hasTitle && !hasLegend)
        {
            defaultPlotBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.HorizontalBarTitleNoLegend);
        }
        else if (hasTitle && !hasLegend && HasInsideValueAxisCrossing(sceneChart, barPlot, barChart, chartXml))
        {
            defaultPlotBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.BarTitleNoLegendInsideCrossing);
        }
        else if (hasTitle && !hasLegend)
        {
            defaultPlotBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.BarTitleNoLegend);
        }
        else
        {
            defaultPlotBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.BarDefault);
        }

        defaultPlotBox = AdjustBarChartPlotBoxForVisibleValueAxes(theme, defaultPlotBox, frame, chartXml, sceneChart, horizontalBars);
        defaultPlotBox = AdjustBarChartPlotBoxForStackedValueAxisLabels(theme, defaultPlotBox, frame, chartXml, sceneChart, barPlot, barChart, horizontalBars);
        if (ignoreManualPlotLayout)
        {
            return ChartPlotLayout.FromPlotBox(defaultPlotBox);
        }

        ChartPlotBox manualDefaultPlotBox = horizontalBars && HasExplicitManualPlotLayoutTarget(sceneChart, chartXml)
            ? GetHorizontalBarManualLayoutTargetDefaultPlotBox(frame, defaultPlotBox)
            : defaultPlotBox;
        if (!TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, manualDefaultPlotBox, out ChartPlotLayout manualPlotLayout))
        {
            return ChartPlotLayout.FromPlotBox(defaultPlotBox);
        }

        return ResolveBarManualPlotLayoutTarget(theme, chartXml, sceneChart, barPlot, barChart, manualPlotLayout, horizontalBars);
    }

    private static bool HasInsideValueAxisCrossing(PptxSceneChart? sceneChart, PptxSceneChartPlot? barPlot, XElement barChart, XDocument chartXml)
    {
        IReadOnlyList<IReadOnlyList<double>> series = ReadSceneOrXmlChartSeries(barPlot, barChart);
        if (series.Count == 0)
        {
            return false;
        }

        PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(barPlot, barChart, PptxSceneChartGrouping.Clustered);
        XElement? valueAxis = ReadChartValueAxisForChart(chartXml, barChart);
        PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, barPlot, PptxSceneChartAxisKind.Value);
        ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(
            valueSceneAxis,
            valueAxis,
            GetBarChartValueExtents(series, grouping),
            IsPercentStackedChartGrouping(grouping));
        double? crossing = ReadSceneOrXmlValueAxisCrossingValue(valueSceneAxis, valueAxis, valueExtents);
        return crossing > valueExtents.Min + PptxChartMetricRules.AxisValueEpsilon &&
            crossing < valueExtents.Max - PptxChartMetricRules.AxisValueEpsilon;
    }

    private static bool HasExplicitManualPlotLayoutTarget(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (!string.IsNullOrWhiteSpace(sceneChart?.PlotAreaLayout.LayoutTarget))
        {
            return true;
        }

        return chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "layout")
            ?.Element(ChartNamespace + "manualLayout")
            ?.Element(ChartNamespace + "layoutTarget")
            ?.Attribute("val") is not null;
    }

    private static ChartPlotBox GetHorizontalBarManualLayoutTargetDefaultPlotBox(ChartFrameBox frame, ChartPlotBox defaultPlotBox)
    {
        double x = frame.X + frame.Width * PptxChartMetricRules.HorizontalBarManualLayoutTargetPlotBoxXRatio;
        double y = frame.Y + frame.Height * PptxChartMetricRules.HorizontalBarManualLayoutTargetPlotBoxYRatio;
        return new ChartPlotBox(x, y, defaultPlotBox.Width, defaultPlotBox.Height);
    }

    private static ChartPlotLayout ResolveBarManualPlotLayoutTarget(
        PptxTheme theme,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? barPlot,
        XElement barChart,
        ChartPlotLayout layout,
        bool horizontalBars)
    {
        if (!horizontalBars || layout.ManualLayoutTargetKind != PptxSceneChartManualLayoutTarget.Outer)
        {
            return layout;
        }

        ChartPlotBox plotBox = DeriveHorizontalBarInnerPlotBox(theme, chartXml, sceneChart, barPlot, barChart, layout.PlotAreaBox);
        return new ChartPlotLayout(layout.PlotAreaBox, plotBox, layout.ManualLayoutTargetKind);
    }

    private static ChartPlotBox DeriveHorizontalBarInnerPlotBox(
        PptxTheme theme,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? barPlot,
        XElement barChart,
        ChartLayoutBox plotAreaBox)
    {
        double leftReserve = 0d;
        XElement? categoryAxis = ReadChartCategoryAxisForChart(chartXml, barChart);
        PptxSceneChartAxis? categorySceneAxis = ReadSceneChartAxis(sceneChart, barPlot, PptxSceneChartAxisKind.Category);
        if (IsSceneOrXmlChartAxisLabelVisible(categorySceneAxis, categoryAxis))
        {
            double labelOffsetScale = ResolveSceneOrXmlCategoryAxisLabelOffsetScale(categorySceneAxis, categoryAxis);
            double outsideFactor =
                PptxChartMetricRules.CategoryAxisHorizontalLeftOffsetRatio * labelOffsetScale +
                PptxChartMetricRules.CategoryAxisHorizontalWidthRatio;
            leftReserve = plotAreaBox.Width * outsideFactor / (1d + outsideFactor);
        }

        double rightReserve = ReadSceneOrXmlChartAxisMajorTickMark(categorySceneAxis, categoryAxis) == PptxSceneChartAxisTickMark.Outside
            ? PptxChartMetricRules.CategoryAxisMajorTickLength + Math.Max(3d, PptxChartMetricRules.ValueAxisFallbackFontSize * 0.35d)
            : 0d;
        double topReserve = 0d;
        XElement? valueAxis = ReadChartValueAxisForChart(chartXml, barChart);
        PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, barPlot, PptxSceneChartAxisKind.Value);
        if (IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis))
        {
            double labelHeight = PptxChartMetricRules.ValueAxisFallbackFontSize * PptxChartMetricRules.AxisLabelHeightFactor;
            topReserve = labelHeight * (
                PptxChartMetricRules.HorizontalValueAxisTopOffsetFactor +
                PptxChartMetricRules.AxisLabelClipTopOffsetFactor +
                PptxChartMetricRules.AxisLabelClipHeightFactor);
        }

        double x = plotAreaBox.X + Math.Min(leftReserve, plotAreaBox.Width * 0.8d);
        double y = plotAreaBox.Y;
        double width = Math.Max(1d, plotAreaBox.Width - leftReserve - rightReserve);
        double height = Math.Max(1d, plotAreaBox.Height - topReserve);
        return new ChartPlotBox(x, y, width, height);
    }

    private static ChartPlotBox AdjustBarChartPlotBoxForVisibleValueAxes(PptxTheme theme, ChartPlotBox plotBox, ChartFrameBox frame, XDocument chartXml, PptxSceneChart? sceneChart, bool horizontalBars)
    {
        if (horizontalBars)
        {
            return plotBox;
        }

        IReadOnlyList<XElement> barCharts = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Bar);
        IReadOnlyList<PptxSceneChartPlot> barPlots = ReadSceneChartPlots(sceneChart, PptxSceneChartPlotKind.Bar);
        int plotCount = Math.Max(barCharts.Count, barPlots.Count);
        if (plotCount < 2)
        {
            return plotBox;
        }

        var valueAxes = new List<ChartAxisSource>();
        var valueAxisIds = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < plotCount; i++)
        {
            PptxSceneChartPlot? currentBarPlot = i < barPlots.Count ? barPlots[i] : null;
            XElement? currentBarChart = i < barCharts.Count ? barCharts[i] : null;
            ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, currentBarPlot, chartXml, currentBarChart)
                .FirstOrDefault();
            if (!IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
            {
                continue;
            }

            string axisId = valueAxis.SceneAxis?.Id ?? ReadChartAxisId(valueAxis.XmlAxis) ?? FormattableString.Invariant($"plot-{i}");
            if (valueAxisIds.Add(axisId))
            {
                valueAxes.Add(valueAxis);
            }
        }

        if (valueAxes.Count < 2)
        {
            return plotBox;
        }

        double leftReserve = plotBox.X - frame.X;
        double rightReserve = frame.X + frame.Width - plotBox.X - plotBox.Width;
        double requiredLeftReserve = leftReserve;
        double requiredRightReserve = rightReserve;
        foreach (ChartAxisSource valueAxis in valueAxes)
        {
            ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(
                valueAxis.SceneAxis,
                valueAxis.XmlAxis,
                new ChartValueExtents(0d, 1d));
            ChartAxisUnits units = ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxis.XmlAxis);
            double stripWidth = EstimateVerticalValueAxisLabelStripWidth(
                theme,
                sceneChart,
                chartXml,
                valueAxis.XmlAxis,
                valueAxis.SceneAxis,
                extents,
                units,
                defaultNumberFormat: null);
            bool labelsRight = ResolveSceneOrXmlValueAxisLabelsRightSide(
                valueAxis.SceneAxis,
                valueAxis.XmlAxis,
                ResolveSceneOrXmlValueAxisRightSide(valueAxis.SceneAxis, valueAxis.XmlAxis, defaultRightSide: false));
            if (labelsRight)
            {
                requiredRightReserve = Math.Max(requiredRightReserve, stripWidth * PptxChartMetricRules.BarMultiValueAxisSecondaryStripFactor);
            }
            else
            {
                requiredLeftReserve = Math.Max(requiredLeftReserve, stripWidth * PptxChartMetricRules.BarMultiValueAxisPrimaryStripFactor);
            }
        }

        double x = frame.X + requiredLeftReserve;
        double right = frame.X + frame.Width - requiredRightReserve;
        double width = Math.Max(1d, right - x);
        return new ChartPlotBox(x, plotBox.Y, width, plotBox.Height);
    }

    private static ChartPlotBox AdjustBarChartPlotBoxForStackedValueAxisLabels(
        PptxTheme theme,
        ChartPlotBox plotBox,
        ChartFrameBox frame,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? barPlot,
        XElement barChart,
        bool horizontalBars)
    {
        if (horizontalBars)
        {
            return plotBox;
        }

        PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(barPlot, barChart, PptxSceneChartGrouping.Clustered);
        bool percentStacked = IsPercentStackedChartGrouping(grouping);
        if (!IsStackedChartGrouping(grouping))
        {
            return plotBox;
        }

        XElement? valueAxis = ReadChartValueAxisForChart(chartXml, barChart);
        PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, barPlot, PptxSceneChartAxisKind.Value);
        if (!IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis))
        {
            return plotBox;
        }

        IReadOnlyList<IReadOnlyList<double>> series = ReadSceneOrXmlChartSeries(barPlot, barChart);
        if (series.Count == 0)
        {
            return plotBox;
        }

        ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(
            valueSceneAxis,
            valueAxis,
            GetBarChartValueExtents(series, grouping),
            percentStacked);
        ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueSceneAxis, valueAxis), percentStacked);
        string? defaultNumberFormat = percentStacked ? "0%" : null;
        double requiredReserve = EstimateVerticalValueAxisLabelStripWidth(theme, sceneChart, chartXml, valueAxis, valueSceneAxis, valueExtents, axisUnits, defaultNumberFormat);
        double leftReserve = plotBox.X - frame.X;
        double rightReserve = frame.X + frame.Width - plotBox.X - plotBox.Width;
        bool labelsRight = ResolveSceneOrXmlValueAxisLabelsRightSide(valueSceneAxis, valueAxis, defaultRightSide: false);
        if (labelsRight)
        {
            rightReserve = Math.Max(rightReserve, requiredReserve);
        }
        else
        {
            leftReserve = Math.Max(leftReserve, requiredReserve);
        }

        double x = frame.X + leftReserve;
        double right = frame.X + frame.Width - rightReserve;
        double width = Math.Max(1d, right - x);
        return new ChartPlotBox(x, plotBox.Y, width, plotBox.Height);
    }

    private static double EstimateVerticalValueAxisLabelStripWidth(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml, XElement valueAxis)
    {
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis: null, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize);
        double fontSize = style.FontSize;
        ChartValueExtents extents = ReadChartValueAxisExtents(valueAxis, new ChartValueExtents(0d, 1d));
        ChartAxisUnits units = ReadChartValueAxisUnits(valueAxis);
        return EstimateVerticalValueAxisLabelStripWidth(theme, sceneChart, chartXml, valueAxis, sceneAxis: null, extents, units, defaultNumberFormat: null);
    }

    private static double EstimateVerticalValueAxisLabelStripWidth(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml, XElement? valueAxis, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits units, string? defaultNumberFormat)
    {
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize);
        double fontSize = style.FontSize;
        IReadOnlyList<double> tickValues = GetChartAxisTickValues(extents, units.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount);
        double maxLabelWidth = tickValues
            .Select(value => FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis, defaultNumberFormat))
            .DefaultIfEmpty("0")
            .Max(label => EstimateChartTextWidth(label, fontSize));
        double labelWidth = Math.Max(
            fontSize * PptxChartMetricRules.ValueAxisMinimumLabelWidthFactor,
            maxLabelWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
        double sideGap = Math.Max(3d, fontSize * PptxChartMetricRules.ValueAxisLabelSideGapFactor);
        return labelWidth + sideGap;
    }

    private static ChartValueExtents GetBarChartValueExtents(IReadOnlyList<IReadOnlyList<double>> series, PptxSceneChartGrouping grouping)
    {
        int categoryCount = Math.Max(1, series.Max(values => values.Count));
        bool stacked = IsStackedChartGrouping(grouping);
        bool percentStacked = IsPercentStackedChartGrouping(grouping);
        (double min, double max) = stacked
            ? GetStackedValueExtents(series, categoryCount, percentStacked)
            : GetClusteredValueExtents(series);
        return new ChartValueExtents(min, max);
    }

    private static bool IsStackedChartGrouping(PptxSceneChartGrouping grouping)
    {
        return grouping is PptxSceneChartGrouping.Stacked or PptxSceneChartGrouping.PercentStacked;
    }

    private static bool IsPercentStackedChartGrouping(PptxSceneChartGrouping grouping)
    {
        return grouping == PptxSceneChartGrouping.PercentStacked;
    }

    private static void DrawHorizontalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, ChartValueExtents extents, double? explicitUnit, double? crossingValue, bool reversed, bool major, ChartSeriesStroke? gridlineStroke)
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
        bool hasPath = false;
        foreach (double value in GetChartGridlineValues(extents, explicitUnit, crossingValue))
        {
            double y = ChartValueToPlotCoordinate(extents, value, plotY, plotHeight, reversed);
            graphics.MoveTo(plotX, y);
            graphics.LineTo(plotX + plotWidth, y);
            hasPath = true;
        }

        if (hasPath)
        {
            graphics.StrokeCurrentPath();
        }

        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static void DrawVerticalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, ChartValueExtents extents, double? explicitUnit, double? crossingValue, bool reversed, bool major, ChartSeriesStroke? gridlineStroke, double autoTickTargetCount = PptxChartMetricRules.AxisNiceTickTargetCount)
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
        bool hasPath = false;
        foreach (double value in GetChartGridlineValues(extents, explicitUnit, crossingValue, autoTickTargetCount))
        {
            double x = ChartValueToPlotCoordinate(extents, value, plotX, plotWidth, reversed);
            graphics.MoveTo(x, plotY);
            graphics.LineTo(x, plotY + plotHeight);
            hasPath = true;
        }

        if (hasPath)
        {
            graphics.StrokeCurrentPath();
        }

        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static ChartSeriesStroke DefaultChartGridlineStroke(bool major)
    {
        return major
            ? new ChartSeriesStroke(new RgbColor(0, 0, 0), 1d, 0.75d)
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

    private static ChartSeriesFill ResolveBarPointFill(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, double value)
    {
        if (value < 0d && !HasExplicitChartPointFill(pointFills, seriesIndex, categoryIndex))
        {
            return new ChartSeriesFill(new RgbColor(255, 255, 255), 1d);
        }

        return ChartPointCategoryOrSeriesColor(theme, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, pointFills);
    }

    private static bool HasExplicitChartPointFill(IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, int seriesIndex, int categoryIndex)
    {
        return seriesIndex < pointFills.Count && pointFills[seriesIndex].ContainsKey(categoryIndex);
    }

    private static bool HasExplicitChartPointStroke(IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, int seriesIndex, int categoryIndex)
    {
        return seriesIndex < pointStrokes.Count && pointStrokes[seriesIndex].ContainsKey(categoryIndex);
    }

    private static ChartSeriesStroke? ResolveNegativeBarFallbackStroke(IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, int seriesIndex, int categoryIndex, double value)
    {
        return value < 0d && !HasExplicitChartPointStroke(pointStrokes, seriesIndex, categoryIndex)
            ? ChartNegativeBarDefaultStroke
            : null;
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

    private static void RenderClusteredHorizontalBars(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, double zeroX, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent, double overlapPercent)
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
                ChartSeriesFill fill = ResolveBarPointFill(theme, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, value);
                double valueX = ChartValueToPlotCoordinate(valueExtents, value, plotX, plotWidth, valueAxisReversed);
                double barX = Math.Min(zeroX, valueX);
                double barWidth = Math.Abs(valueX - zeroX);
                double barY = categoryY + seriesIndex * step;
                FillChartRectangle(graphics, barX, barY, barWidth, barHeight, fill);
                StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, barX, barY, barWidth, barHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, value));
            }
        }
    }

    private static void RenderStackedColumns(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent)
    {
        double categoryWidth = plotWidth / categoryCount;
        double barWidth = GetStackedBarWidth(categoryWidth, gapWidthPercent);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryX = plotX + category * categoryWidth + (categoryWidth - barWidth) / 2d;
            double positiveValue = 0d;
            double negativeValue = 0d;
            double positiveTotal = GetCategoryPositiveTotal(series, category, percentStacked);
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = NormalizeStackedValue(values[category], positiveTotal, percentStacked);
                ChartSeriesFill fill = ResolveBarPointFill(theme, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, value);
                double segmentStartValue;
                double segmentEndValue;
                if (value >= 0d)
                {
                    segmentStartValue = positiveValue;
                    positiveValue += value;
                    segmentEndValue = positiveValue;
                }
                else
                {
                    segmentStartValue = negativeValue;
                    negativeValue += value;
                    segmentEndValue = negativeValue;
                }

                double startY = ChartValueToPlotCoordinate(valueExtents, segmentStartValue, plotY, plotHeight, valueAxisReversed);
                double endY = ChartValueToPlotCoordinate(valueExtents, segmentEndValue, plotY, plotHeight, valueAxisReversed);
                double segmentY = Math.Min(startY, endY);
                double segmentHeight = Math.Abs(endY - startY);
                FillChartRectangle(graphics, categoryX, segmentY, barWidth, segmentHeight, fill);
                StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, categoryX, segmentY, barWidth, segmentHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, value));
            }
        }
    }

    private static void RenderStackedHorizontalBars(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent)
    {
        double categoryHeight = plotHeight / categoryCount;
        double barHeight = GetStackedBarWidth(categoryHeight, gapWidthPercent);
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryY = plotY + category * categoryHeight + (categoryHeight - barHeight) / 2d;
            double positiveValue = 0d;
            double negativeValue = 0d;
            double positiveTotal = GetCategoryPositiveTotal(series, category, percentStacked);
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double> values = series[seriesIndex];
                if (category >= values.Count)
                {
                    continue;
                }

                double value = NormalizeStackedValue(values[category], positiveTotal, percentStacked);
                ChartSeriesFill fill = ResolveBarPointFill(theme, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, value);
                double segmentStartValue;
                double segmentEndValue;
                if (value >= 0d)
                {
                    segmentStartValue = positiveValue;
                    positiveValue += value;
                    segmentEndValue = positiveValue;
                }
                else
                {
                    segmentStartValue = negativeValue;
                    negativeValue += value;
                    segmentEndValue = negativeValue;
                }

                double startX = ChartValueToPlotCoordinate(valueExtents, segmentStartValue, plotX, plotWidth, valueAxisReversed);
                double endX = ChartValueToPlotCoordinate(valueExtents, segmentEndValue, plotX, plotWidth, valueAxisReversed);
                double segmentX = Math.Min(startX, endX);
                double segmentWidth = Math.Abs(endX - startX);
                FillChartRectangle(graphics, segmentX, categoryY, segmentWidth, barHeight, fill);
                StrokeChartPointRectangle(graphics, seriesIndex, category, pointStrokes, segmentX, categoryY, segmentWidth, barHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, value));
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

    private static void RenderLineChart(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, ChartLayoutBox plotAreaBox, ChartPlotBox plotBox, IReadOnlyList<IReadOnlyList<double>> series, bool stacked, bool percentStacked, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, IReadOnlyList<bool> smoothSeries, bool majorGridlines, bool minorGridlines, ChartGridlineStyle gridlineStyle, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents, ChartAxisUnits axisUnits, double? valueAxisCrossingValue, bool valueAxisReversed)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        graphics.SaveState();
        ClipChartPlotArea(graphics, plotX, plotY, plotWidth, plotHeight);
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double maxValue = valueExtents.Max;
        double minValue = valueExtents.Min;
        double valueRange = Math.Max(1d, maxValue - minValue);
        double valueAxisCrossingY = ChartValueToPlotCoordinate(valueExtents, valueAxisCrossingValue, plotY, plotHeight, valueAxisReversed);

        if (minorGridlines)
        {
            DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor);
        }

        if (majorGridlines)
        {
            DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major);
        }

        ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
        ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
        if (axesStyle.CategoryAxisVisible)
        {
            if (categoryAxisStroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, categoryAxisStroke);
                graphics.StrokeLine(plotX, valueAxisCrossingY, plotX + plotWidth, valueAxisCrossingY);
            }
        }

        if (axesStyle.ValueAxisVisible)
        {
            if (valueAxisStroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, valueAxisStroke);
                double axisX = axesStyle.ValueAxisRightSide ? plotX + plotWidth : plotX;
                graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight);
            }

            if (axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke)
            {
                if (secondaryValueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, secondaryValueAxisStroke);
                    double axisX = axesStyle.SecondaryValueAxisRightSide ? plotX + plotWidth : plotX;
                    graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight);
                }
            }
        }

        double[] lower = new double[pointCount];
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            if (values.Count == 0)
            {
                continue;
            }

            ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, chartPalette, seriesIndex, seriesStrokes, ChartLineDefaultStrokeWidth);
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            var points = new List<(double X, double Y)>(values.Count);
            for (int i = 0; i < values.Count; i++)
            {
                double pointX = plotX + plotWidth * (i + 0.5d) / pointCount;
                double value = values[i];
                double positiveTotal = GetCategoryPositiveTotal(series, i, percentStacked);
                double normalizedValue = NormalizeStackedValue(value, positiveTotal, percentStacked);
                double plottedValue = stacked ? lower[i] + normalizedValue : value;
                double pointY = ChartValueToPlotCoordinate(valueExtents, plottedValue, plotY, plotHeight, valueAxisReversed);
                points.Add((pointX, pointY));
                if (stacked)
                {
                    lower[i] = plottedValue;
                }
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

        graphics.RestoreState();
        if (axesStyle.CategoryAxisVisible && categoryAxisStroke.Alpha > 0.001d)
        {
            SetChartStroke(graphics, categoryAxisStroke);
            DrawLineChartCategoryAxisMajorTicks(graphics, plotX, plotWidth, pointCount, valueAxisCrossingY, axesStyle.CategoryAxisMajorTickMark);
        }
    }

    private static ChartLayout GetLineChartLayout(PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, sceneChart, chartXml);
        ChartPlotLayout plotLayout = GetLineChartPlotLayout(frame, chartXml, sceneChart, title, legend);
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, legend);
    }

    private static ChartPlotLayout GetLineChartPlotLayout(ChartFrameBox frame, XDocument chartXml, PptxSceneChart? sceneChart, string? title, ChartLegendLayout legend)
    {
        bool hasTitle = !string.IsNullOrWhiteSpace(title);
        bool hasRightLegend = legend.Visible && !legend.Overlay && legend.PositionKind == PptxSceneChartLegendPosition.Right;
        bool hasLineChart = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Line).Count != 0;
        ChartPlotBox defaultPlotBox = !hasTitle && hasRightLegend
            ? GetCartesianNoTitleRightLegendPlotBox(frame, chartXml, sceneChart)
            : hasTitle && hasRightLegend && hasLineChart
                ? GetLineTitleRightLegendPlotBox(frame)
            : GetDefaultChartPlotBox(frame);
        return TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout manualPlotLayout)
            ? manualPlotLayout
            : ChartPlotLayout.FromPlotBox(defaultPlotBox);
    }

    private static ChartPlotBox GetLineTitleRightLegendPlotBox(ChartFrameBox frame)
    {
        return GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.LineTitleRightLegend);
    }

    private static ChartPlotBox GetCartesianNoTitleRightLegendPlotBox(ChartFrameBox frame, XDocument chartXml, PptxSceneChart? sceneChart)
    {
        XElement? plotElement = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Line).FirstOrDefault();
        PptxSceneChartPlotKind plotKind = PptxSceneChartPlotKind.Line;
        if (plotElement is null)
        {
            plotElement = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Area).FirstOrDefault();
            plotKind = PptxSceneChartPlotKind.Area;
        }
        if (plotElement is null)
        {
            plotElement = ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Scatter).FirstOrDefault();
            plotKind = PptxSceneChartPlotKind.Scatter;
        }

        if (plotElement is null)
        {
            return GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.LineNoTitleRightLegend);
        }

        PptxSceneChartPlot? plot = ReadSceneChartPlot(sceneChart, plotKind);
        IReadOnlyList<string> seriesNames = ReadSceneOrXmlChartSeriesNames(plot, plotElement);
        double legendFontSize = PptxChartMetricRules.LegendFallbackFontSize;
        double maxLegendTextWidth = seriesNames.Count == 0
            ? 0d
            : seriesNames.Max(name => EstimateChartTextWidth(name, legendFontSize));
        int maxLegendTextLength = seriesNames.Count == 0
            ? 0
            : seriesNames.Max(name => name.Length);
        double rightReserve = maxLegendTextWidth +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeMarkerWidthFactor +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeTextGapFactor +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeGapFactor +
            PptxChartMetricRules.LineRightLegendReservePadding +
            Math.Max(0, maxLegendTextLength - 6) * PptxChartMetricRules.LineRightLegendExtraLegendCharacterPadding;

        double maxValueLabelWidth = 0d;
        int maxValueLabelLength = 0;
        if (plotKind == PptxSceneChartPlotKind.Scatter)
        {
            IReadOnlyList<ScatterSeries> series = ReadSceneOrXmlScatterSeries(plot, plotElement, readBubbleSize: false);
            if (series.Count > 0)
            {
                IReadOnlyList<XElement> valueAxes = ReadChartValueAxesForChart(chartXml, plotElement);
                XElement? valueAxis = valueAxes.Count > 1
                    ? valueAxes[1]
                    : ReadChartValueAxisForChart(chartXml, plotElement) ?? chartXml.Descendants(ChartNamespace + "valAx").FirstOrDefault();
                ChartValueExtents valueExtents = ReadBubbleChartValueAxisExtents(valueAxis, GetScatterYValueExtents(series));
                ChartAxisUnits axisUnits = ResolveBubbleAxisUnits(ReadChartValueAxisUnits(valueAxis), valueExtents);
                IReadOnlyList<double> tickValues = GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true);
                string[] tickLabels = tickValues
                    .Select(value => FormatSceneOrXmlChartAxisLabel(value, sceneAxis: null, valueAxis, defaultNumberFormat: null))
                    .ToArray();
                maxValueLabelWidth = tickLabels.Length == 0
                    ? 0d
                    : tickLabels.Max(label => EstimateChartTextWidth(label, PptxChartMetricRules.ValueAxisFallbackFontSize));
                maxValueLabelLength = tickLabels.Length == 0
                    ? 0
                    : tickLabels.Max(label => label.Length);
            }
        }
        else
        {
            PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(plot, plotElement, PptxSceneChartGrouping.Standard);
            bool stacked = IsStackedChartGrouping(grouping);
            bool percentStacked = IsPercentStackedChartGrouping(grouping);
            IReadOnlyList<IReadOnlyList<double>> series = ReadSceneOrXmlChartSeries(plot, plotElement);
            if (series.Count > 0)
            {
                XElement? valueAxis = ReadChartValueAxisForChart(chartXml, plotElement) ??
                    chartXml.Descendants(ChartNamespace + "valAx").FirstOrDefault();
                PptxSceneChartAxis? valueSceneAxis = ReadSceneChartAxis(sceneChart, plot, PptxSceneChartAxisKind.Value);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueSceneAxis, valueAxis, GetLineChartValueExtents(series, stacked, percentStacked), percentStacked, useNearMaximumHeadroom: !percentStacked);
                ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueSceneAxis, valueAxis), percentStacked);
                IReadOnlyList<double> tickValues = GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true);
                string[] tickLabels = tickValues
                    .Select(value => FormatSceneOrXmlChartAxisLabel(value, valueSceneAxis, valueAxis, percentStacked ? "0%" : null))
                    .ToArray();
                maxValueLabelWidth = tickLabels.Length == 0
                    ? 0d
                    : tickLabels.Max(label => EstimateChartTextWidth(label, PptxChartMetricRules.ValueAxisFallbackFontSize));
                maxValueLabelLength = tickLabels.Length == 0
                    ? 0
                    : tickLabels.Max(label => label.Length);
            }
        }

        double leftInset = maxValueLabelWidth > 0d
            ? maxValueLabelWidth +
                PptxChartMetricRules.LineRightLegendValueAxisPadding +
                Math.Max(0, maxValueLabelLength - 3) * PptxChartMetricRules.LineRightLegendExtraValueLabelCharacterPadding
            : frame.Width * PptxChartMetricRules.LineNoTitleRightLegendPlotBoxXRatio;
        double x = frame.X + leftInset;
        double y = frame.Y + frame.Height * PptxChartMetricRules.LineNoTitleRightLegendPlotBoxYRatio;
        double width = Math.Max(1d, frame.Width - leftInset - rightReserve);
        double height = frame.Height * PptxChartMetricRules.LineNoTitleRightLegendPlotBoxHeightRatio;
        return new ChartPlotBox(x, y, width, height);
    }

    private static ChartLayout GetBubbleChartLayout(PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartPlot? bubblePlot, XElement bubbleChart)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, sceneChart, chartXml);
        ChartPlotLayout plotLayout = GetBubbleChartPlotLayout(frame, chartXml, sceneChart, bubblePlot, bubbleChart, title, legend);
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, legend);
    }

    private static ChartPlotLayout GetBubbleChartPlotLayout(ChartFrameBox frame, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartPlot? bubblePlot, XElement bubbleChart, string? title, ChartLegendLayout legend)
    {
        bool hasTitle = !string.IsNullOrWhiteSpace(title);
        bool hasRightLegend = legend.Visible && !legend.Overlay && legend.PositionKind == PptxSceneChartLegendPosition.Right;
        ChartPlotBox defaultPlotBox = hasTitle && hasRightLegend
            ? GetBubbleTitleRightLegendPlotBox(frame, bubblePlot, bubbleChart)
            : GetDefaultChartPlotBox(frame);
        return TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout manualPlotLayout)
            ? manualPlotLayout
            : ChartPlotLayout.FromPlotBox(defaultPlotBox);
    }

    private static ChartPlotBox GetBubbleTitleRightLegendPlotBox(ChartFrameBox frame, PptxSceneChartPlot? bubblePlot, XElement bubbleChart)
    {
        IReadOnlyList<string> seriesNames = ReadSceneOrXmlChartSeriesNames(bubblePlot, bubbleChart);
        double legendFontSize = PptxChartMetricRules.LegendFallbackFontSize;
        double maxLegendTextWidth = seriesNames.Count == 0
            ? 0d
            : seriesNames.Max(name => EstimateChartTextWidth(name, legendFontSize));
        int maxLegendTextLength = seriesNames.Count == 0
            ? 0
            : seriesNames.Max(name => name.Length);
        double rightReserve = maxLegendTextWidth +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeMarkerWidthFactor +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeTextGapFactor +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeGapFactor +
            PptxChartMetricRules.LineRightLegendReservePadding +
            Math.Max(0, maxLegendTextLength - 6) * PptxChartMetricRules.LineRightLegendExtraLegendCharacterPadding;

        double x = frame.X + frame.Width * PptxChartMetricRules.LineTitleRightLegendPlotBoxXRatio;
        double y = frame.Y + frame.Height * PptxChartMetricRules.LineTitleRightLegendPlotBoxYRatio;
        double width = Math.Max(1d, frame.Width - (x - frame.X) - rightReserve);
        double height = frame.Height * PptxChartMetricRules.LineTitleRightLegendPlotBoxHeightRatio;
        return new ChartPlotBox(x, y, width, height);
    }

    private static ChartValueExtents GetLineChartValueExtents(IReadOnlyList<IReadOnlyList<double>> series)
    {
        return GetLineChartValueExtents(series, stacked: false, percentStacked: false);
    }

    private static ChartValueExtents GetLineChartValueExtents(IReadOnlyList<IReadOnlyList<double>> series, bool stacked, bool percentStacked)
    {
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        (double minValue, double maxValue) = stacked
            ? GetStackedValueExtents(series, pointCount, percentStacked)
            : GetClusteredValueExtents(series);
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
        if (marker.SymbolKind == PptxSceneChartMarkerSymbol.None)
        {
            return;
        }

        double size = marker.Size;
        ChartSeriesFill fill = marker.Fill ?? new ChartSeriesFill(defaultFill, 1d);
        ChartSeriesStroke? stroke = marker.Stroke ?? (IsLineOnlyChartMarker(marker.SymbolKind) ? new ChartSeriesStroke(defaultStroke, 1d, Math.Max(0.75d, size * 0.16d)) : null);
        if (!IsLineOnlyChartMarker(marker.SymbolKind))
        {
            if (fill.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            switch (marker.SymbolKind)
            {
                case PptxSceneChartMarkerSymbol.Dot:
                    graphics.FillEllipse(x - size / 4d, y - size / 4d, size / 2d, size / 2d);
                    break;
                case PptxSceneChartMarkerSymbol.Square:
                    graphics.FillRectangle(x - size / 2d, y - size / 2d, size, size);
                    break;
                case PptxSceneChartMarkerSymbol.Diamond:
                    graphics.FillPolygon([
                        (x, y + size / 2d),
                        (x + size / 2d, y),
                        (x, y - size / 2d),
                        (x - size / 2d, y)
                    ]);
                    break;
                case PptxSceneChartMarkerSymbol.Triangle:
                    graphics.FillPolygon([
                        (x, y + size / 2d),
                        (x + size / 2d, y - size / 2d),
                        (x - size / 2d, y - size / 2d)
                    ]);
                    break;
                case PptxSceneChartMarkerSymbol.Star:
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
        switch (marker.SymbolKind)
        {
            case PptxSceneChartMarkerSymbol.Dash:
                graphics.StrokeLine(x - size / 2d, y, x + size / 2d, y);
                break;
            case PptxSceneChartMarkerSymbol.Plus:
                graphics.StrokeLine(x - size / 2d, y, x + size / 2d, y);
                graphics.StrokeLine(x, y - size / 2d, x, y + size / 2d);
                break;
            case PptxSceneChartMarkerSymbol.X:
                graphics.StrokeLine(x - size / 2d, y - size / 2d, x + size / 2d, y + size / 2d);
                graphics.StrokeLine(x - size / 2d, y + size / 2d, x + size / 2d, y - size / 2d);
                break;
            case PptxSceneChartMarkerSymbol.Square:
                graphics.StrokeRectangle(x - size / 2d, y - size / 2d, size, size);
                break;
            case PptxSceneChartMarkerSymbol.Diamond:
                graphics.StrokePolygon([
                    (x, y + size / 2d),
                    (x + size / 2d, y),
                    (x, y - size / 2d),
                    (x - size / 2d, y)
                ]);
                break;
            case PptxSceneChartMarkerSymbol.Triangle:
                graphics.StrokePolygon([
                    (x, y + size / 2d),
                    (x + size / 2d, y - size / 2d),
                    (x - size / 2d, y - size / 2d)
                ]);
                break;
            case PptxSceneChartMarkerSymbol.Star:
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

    private static bool IsLineOnlyChartMarker(PptxSceneChartMarkerSymbol symbol)
    {
        return symbol is PptxSceneChartMarkerSymbol.Plus or
            PptxSceneChartMarkerSymbol.X or
            PptxSceneChartMarkerSymbol.Dash;
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

    private static void StrokeChartPointRectangle(PdfGraphicsBuilder graphics, int seriesIndex, int categoryIndex, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, double x, double y, double width, double height, ChartSeriesStroke? fallbackStroke = null)
    {
        ChartSeriesStroke stroke = default;
        bool hasExplicitStroke = seriesIndex < pointStrokes.Count && pointStrokes[seriesIndex].TryGetValue(categoryIndex, out stroke);
        if (!hasExplicitStroke)
        {
            if (fallbackStroke is not { } fallback)
            {
                return;
            }

            stroke = fallback;
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

    private static ChartSeriesStroke ChartSeriesStrokeColor(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, double defaultWidth)
    {
        return seriesIndex < seriesStrokes.Count && seriesStrokes[seriesIndex] is { } stroke
            ? stroke
            : new ChartSeriesStroke(ChartPalette(chartPalette, theme, seriesIndex), 1d, defaultWidth);
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

    private static void RenderAreaChart(
        PdfGraphicsBuilder graphics,
        PptxTheme theme,
        IReadOnlyList<RgbColor>? chartPalette,
        ChartLayoutBox plotAreaBox,
        ChartPlotBox plotBox,
        IReadOnlyList<IReadOnlyList<double>> series,
        bool stacked,
        bool percentStacked,
        IReadOnlyList<ChartSeriesFill?> seriesFills,
        IReadOnlyList<ChartSeriesStroke?> seriesStrokes,
        bool majorGridlines,
        bool minorGridlines,
        ChartGridlineStyle gridlineStyle,
        ChartAxesStyle axesStyle,
        ChartShapeStyle plotAreaStyle,
        ChartValueExtents valueExtents,
        ChartAxisUnits axisUnits,
        double? valueAxisCrossingValue,
        bool valueAxisReversed)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        graphics.SaveState();
        try
        {
            ClipChartPlotArea(graphics, plotX, plotY, plotWidth, plotHeight);
            if (minorGridlines)
            {
                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor);
            }

            if (majorGridlines)
            {
                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major);
            }

            RenderAreaChartSeries(graphics, theme, chartPalette, plotBox, series, stacked, percentStacked, seriesFills, seriesStrokes, valueExtents, valueAxisReversed);

            ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
            ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
            double valueAxisCrossingY = ChartValueToPlotCoordinate(valueExtents, valueAxisCrossingValue, plotY, plotHeight, valueAxisReversed);
            if (axesStyle.CategoryAxisVisible && categoryAxisStroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, categoryAxisStroke);
                graphics.StrokeLine(plotX, valueAxisCrossingY, plotX + plotWidth, valueAxisCrossingY);
            }

            if (axesStyle.ValueAxisVisible)
            {
                if (valueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, valueAxisStroke);
                    double axisX = axesStyle.ValueAxisRightSide ? plotX + plotWidth : plotX;
                    graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight);
                }

                if (axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke && secondaryValueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, secondaryValueAxisStroke);
                    double axisX = axesStyle.SecondaryValueAxisRightSide ? plotX + plotWidth : plotX;
                    graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight);
                }
            }
        }
        finally
        {
            graphics.RestoreState();
        }
    }

    private static void RenderAreaChartSeries(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, ChartPlotBox plotBox, IReadOnlyList<IReadOnlyList<double>> series, bool stacked, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents valueExtents, bool valueAxisReversed)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        int pointCount = Math.Max(1, series.Max(values => values.Count));
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
                double positiveTotal = GetCategoryPositiveTotal(series, i, percentStacked);
                double normalizedValue = NormalizeStackedValue(value, positiveTotal, percentStacked);
                double upperValue = stacked ? lower[i] + normalizedValue : value;
                upperPoints[i] = (pointX, ChartValueToPlotCoordinate(valueExtents, upperValue, plotY, plotHeight, valueAxisReversed));
                lowerPoints[i] = (pointX, ChartValueToPlotCoordinate(valueExtents, lowerValue, plotY, plotHeight, valueAxisReversed));
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

            ChartSeriesFill fill = ChartSeriesColor(theme, chartPalette, seriesIndex, seriesFills);
            if (fill.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            graphics.FillPolygon(polygon);
            if (fill.Alpha < 1d)
            {
                graphics.RestoreState();
            }

            ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, chartPalette, seriesIndex, seriesStrokes, ChartLineDefaultStrokeWidth);
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

    private static ChartValueExtents GetScatterXValueExtents(IReadOnlyList<ScatterSeries> series)
    {
        double maxX = Math.Max(1d, series.SelectMany(item => item.Points).Max(point => point.X));
        double minX = Math.Min(0d, series.SelectMany(item => item.Points).Min(point => point.X));
        return new ChartValueExtents(minX, maxX);
    }

    private static ChartValueExtents GetScatterYValueExtents(IReadOnlyList<ScatterSeries> series)
    {
        double maxY = Math.Max(1d, series.SelectMany(item => item.Points).Max(point => point.Y));
        double minY = Math.Min(0d, series.SelectMany(item => item.Points).Min(point => point.Y));
        return new ChartValueExtents(minY, maxY);
    }

    private static ChartValueExtents GetBubbleXValueExtents(IReadOnlyList<ScatterSeries> series)
    {
        return GetScatterXValueExtents(series);
    }

    private static ChartValueExtents GetBubbleYValueExtents(IReadOnlyList<ScatterSeries> series)
    {
        return GetScatterYValueExtents(series);
    }

    private static ChartValueExtents GetScatterDataXExtents(IReadOnlyList<ScatterSeries> series)
    {
        double maxX = series.SelectMany(item => item.Points).Max(point => point.X);
        double minX = series.SelectMany(item => item.Points).Min(point => point.X);
        return new ChartValueExtents(minX, maxX);
    }

    private static ChartValueExtents GetScatterDataYExtents(IReadOnlyList<ScatterSeries> series)
    {
        double maxY = Math.Max(1d, series.SelectMany(item => item.Points).Max(point => point.Y));
        double minY = Math.Min(0d, series.SelectMany(item => item.Points).Min(point => point.Y));
        return new ChartValueExtents(minY, maxY);
    }

    private static void RenderScatterChart(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, IReadOnlyList<ScatterSeries> series, bool connectLines, bool bubble, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, IReadOnlyList<bool> smoothSeries, ChartValueExtents? xValueExtents = null, ChartValueExtents? yValueExtents = null)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        ChartValueExtents xExtents = xValueExtents ?? GetScatterDataXExtents(series);
        ChartValueExtents yExtents = yValueExtents ?? GetScatterDataYExtents(series);
        double minX = xExtents.Min;
        double maxX = xExtents.Max;
        double minY = yExtents.Min;
        double maxY = yExtents.Max;
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
                double radius = bubble
                    ? Math.Sqrt(Math.Max(0d, point.Size) / maxBubbleSize) * Math.Min(plotWidth, plotHeight) * PptxChartMetricRules.BubbleRadiusPlotRatio
                    : 3d;
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

    private static void RenderRadarChart(PdfGraphicsBuilder graphics, ChartRadarLayout layout, IReadOnlyList<IReadOnlyList<double>> series, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents extents, ChartAxisUnits axisUnits)
    {
        ChartPolarGeometry geometry = layout.Geometry;
        int pointCount = layout.PointCount;

        SetChartStroke(graphics, RadarGridlineDefaultStroke);
        bool hasGridPath = false;
        foreach (double tickValue in GetChartAxisTickValues(extents, axisUnits.MajorUnit, includeEndpoints: true))
        {
            double tickRatio = GetChartValuePlotRatio(extents, tickValue, false);
            if (tickRatio <= PptxChartMetricRules.AxisValueEpsilon)
            {
                continue;
            }

            AppendRadarPolygonEdgeSegments(graphics, geometry, pointCount, geometry.Radius * tickRatio);
            hasGridPath = true;
        }
        if (hasGridPath)
        {
            graphics.StrokeCurrentPath();
        }

        AppendRadarRadialSegments(graphics, geometry, pointCount, repeatFirstSpoke: true);
        graphics.StrokeCurrentPath();

        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<double> values = series[seriesIndex];
            var points = new (double X, double Y)[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                double value = i < values.Count ? Math.Max(0d, values[i]) : 0d;
                double pointRadius = GetChartValuePlotRatio(extents, value, false) * geometry.Radius;
                double angle = GetRadarPointAngle(i, pointCount);
                points[i] = (geometry.CenterX + Math.Cos(angle) * pointRadius, geometry.CenterY + Math.Sin(angle) * pointRadius);
            }

            if (layout.IsFilled)
            {
                ChartSeriesFill fill = ChartSeriesColor(seriesIndex, seriesFills, series.Count == 1 ? 0.40d : 0.18d);
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
                graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
                graphics.FillPolygon(points);
                graphics.RestoreState();
            }
            ChartSeriesStroke stroke = ChartSeriesStrokeColor(seriesIndex, seriesStrokes, 1.2d);
            if (stroke.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
            }

            SetChartStroke(graphics, stroke);
            AppendClosedPolylinePath(graphics, points);
            graphics.StrokeCurrentPath();

            if (stroke.Alpha < 1d)
            {
                graphics.RestoreState();
            }
        }
    }

    private static ChartValueExtents GetAreaChartValueExtents(IReadOnlyList<IReadOnlyList<double>> series, bool stacked, bool percentStacked)
    {
        if (percentStacked)
        {
            return new ChartValueExtents(0d, 1d);
        }

        int pointCount = Math.Max(1, series.Max(values => values.Count));
        (double minValue, double maxValue) = stacked
            ? GetStackedValueExtents(series, pointCount, percentStacked)
            : GetClusteredValueExtents(series);
        return new ChartValueExtents(minValue, maxValue);
    }

    private static void AppendRadarPolygonEdgeSegments(PdfGraphicsBuilder graphics, ChartPolarGeometry geometry, int pointCount, double radius)
    {
        var points = new (double X, double Y)[pointCount];
        for (int i = 0; i < pointCount; i++)
        {
            double angle = GetRadarPointAngle(i, pointCount);
            points[i] = (geometry.CenterX + Math.Cos(angle) * radius, geometry.CenterY + Math.Sin(angle) * radius);
        }

        for (int i = 0; i < points.Length; i++)
        {
            (double X, double Y) start = points[i];
            (double X, double Y) end = points[(i + 1) % points.Length];
            graphics.MoveTo(start.X, start.Y);
            graphics.LineTo(end.X, end.Y);
        }
    }

    private static void AppendRadarRadialSegments(PdfGraphicsBuilder graphics, ChartPolarGeometry geometry, int pointCount, bool repeatFirstSpoke)
    {
        int segmentCount = repeatFirstSpoke ? pointCount + 1 : pointCount;
        for (int i = 0; i < segmentCount; i++)
        {
            double angle = GetRadarPointAngle(i % pointCount, pointCount);
            graphics.MoveTo(geometry.CenterX, geometry.CenterY);
            graphics.LineTo(
                geometry.CenterX + Math.Cos(angle) * geometry.Radius,
                geometry.CenterY + Math.Sin(angle) * geometry.Radius);
        }
    }

    private static void AppendClosedPolylinePath(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points)
    {
        if (points.Count == 0)
        {
            return;
        }

        graphics.MoveTo(points[0].X, points[0].Y);
        for (int i = 1; i < points.Count; i++)
        {
            graphics.LineTo(points[i].X, points[i].Y);
        }

        graphics.LineTo(points[0].X, points[0].Y);
    }

    private static double GetRadarPointAngle(int index, int pointCount)
    {
        return Math.PI / 2d - index * Math.PI * 2d / pointCount;
    }

    private static IReadOnlyList<PdfFontResource> RenderRadarCategoryLabels(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartRadarLayout layout,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PptxSceneChartAxis? sceneAxis,
        XElement? categoryAxis,
        IReadOnlyList<string> labels)
    {
        if (labels.Count == 0)
        {
            return [];
        }

        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, categoryAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize);
        ChartPlotBox plotBox = layout.PlotBox;
        int pointCount = Math.Max(labels.Count, layout.PointCount);
        var runs = new List<TextRun>(labels.Count);
        for (int i = 0; i < labels.Count; i++)
        {
            ChartRadarLabelFrame frame = ResolveRadarCategoryLabelFrame(layout, labels[i], style, i, pointCount);
            runs.Add(CreateChartLabelRun(labels[i], frame.X, frame.Y, frame.Width, frame.Height, plotBox, style, frame.Alignment));
        }

        return RenderTextRuns(runs, graphics, "RCA");
    }

    private static ChartRadarLabelFrame ResolveRadarCategoryLabelFrame(ChartRadarLayout layout, string label, ChartTextStyle style, int index, int pointCount)
    {
        ChartPolarGeometry geometry = layout.Geometry;
        double fontSize = style.FontSize;
        double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
        ChartRadarLabelRules labelRules = layout.LabelRules;
        double verticalGap = fontSize * labelRules.CategoryVerticalGapFactor;
        double horizontalGap = fontSize * labelRules.CategoryHorizontalGapFactor;
        double angle = GetRadarPointAngle(index, pointCount);
        double cosine = Math.Cos(angle);
        double anchorX = geometry.CenterX + cosine * (geometry.Radius + horizontalGap);
        double anchorY = geometry.CenterY + Math.Sin(angle) * (geometry.Radius + verticalGap);
        double width = Math.Max(fontSize * 2d, EstimateChartTextWidth(label, fontSize) + fontSize);
        TextAlignment alignment = cosine > 0.25d
            ? TextAlignment.Left
            : cosine < -0.25d
                ? TextAlignment.Right
                : TextAlignment.Center;
        double x = alignment switch
        {
            TextAlignment.Left => anchorX,
            TextAlignment.Right => anchorX - width,
            _ => anchorX - width / 2d
        };
        double y = ResolveRadarCategoryLabelBaselineY(anchorY, angle, height, labelRules);
        return new ChartRadarLabelFrame(x, y, width, height, alignment);
    }

    private static double ResolveRadarCategoryLabelBaselineY(double anchorY, double angle, double labelHeight, ChartRadarLabelRules labelRules)
    {
        double sine = Math.Sin(angle);
        double baselineFactor = labelRules.CategoryBaselineBaseFactor +
            labelRules.CategoryBaselineSineFactor * sine +
            labelRules.CategoryBaselineSineSquaredFactor * sine * sine;
        return anchorY + labelHeight * baselineFactor;
    }

    private static IReadOnlyList<PdfFontResource> RenderRadarValueAxisLabels(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartRadarLayout layout,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        XElement? valueAxis,
        PptxSceneChartAxis? sceneAxis,
        ChartValueExtents extents,
        ChartAxisUnits axisUnits)
    {
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize);
        ChartPlotBox plotBox = layout.PlotBox;
        var runs = new List<TextRun>();
        foreach (double tickValue in GetChartAxisTickValues(extents, axisUnits.MajorUnit, includeEndpoints: true))
        {
            double ratio = GetChartValuePlotRatio(extents, tickValue, false);
            string label = FormatSceneOrXmlChartAxisLabel(tickValue, sceneAxis, valueAxis);
            ChartRadarLabelFrame frame = ResolveRadarValueAxisLabelFrame(layout, style, ratio);
            runs.Add(CreateChartLabelRun(label, frame.X, frame.Y, frame.Width, frame.Height, plotBox, style, frame.Alignment));
        }

        return RenderTextRuns(runs, graphics, "RVA");
    }

    private static ChartRadarLabelFrame ResolveRadarValueAxisLabelFrame(ChartRadarLayout layout, ChartTextStyle style, double ratio)
    {
        ChartPolarGeometry geometry = layout.Geometry;
        double fontSize = style.FontSize;
        double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
        ChartRadarLabelRules labelRules = layout.LabelRules;
        double width = fontSize * labelRules.ValueWidthFactor;
        double x = geometry.CenterX - width - fontSize * labelRules.ValueGapFactor;
        double y = ResolveRadarValueAxisLabelBaselineY(layout, ratio, height);
        return new ChartRadarLabelFrame(x, y, width, height, TextAlignment.Right);
    }

    private static double ResolveRadarValueAxisLabelBaselineY(ChartRadarLayout layout, double ratio, double labelHeight)
    {
        return layout.Geometry.CenterY + layout.Geometry.Radius * ratio - labelHeight * layout.LabelRules.ValueBaselineOffsetFactor;
    }

    private static void RenderPieChart(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<double> values, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions, double firstSliceAngle)
    {
        RenderPieOrDoughnutSlices(graphics, theme, chartPalette, layout, values, pointFills, pointStrokes, pointExplosions, holeSize: 0d, firstSliceAngle);
    }

    private static void RenderPieOrDoughnutSlices(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<double> values, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, double firstSliceAngle)
    {
        double total = values.Where(value => value > 0d).Sum();
        if (total <= 0d)
        {
            return;
        }

        ChartPolarGeometry geometry = layout.Geometry;
        double innerRadius = geometry.Radius * Math.Clamp(holeSize, 0d, 0.95d);
        double angle = Math.PI / 2d - firstSliceAngle * Math.PI / 180d;

        for (int i = 0; i < values.Count; i++)
        {
            double value = Math.Max(0d, values[i]);
            if (value <= 0d)
            {
                continue;
            }

            double sweep = -value / total * Math.PI * 2d;
            double midpointAngle = angle + sweep / 2d;
            double explosionOffset = pointExplosions.TryGetValue(i, out double explosion) ? geometry.Radius * explosion : 0d;
            double sliceCenterX = geometry.CenterX + Math.Cos(midpointAngle) * explosionOffset;
            double sliceCenterY = geometry.CenterY + Math.Sin(midpointAngle) * explosionOffset;
            ChartSeriesFill fill = pointFills.TryGetValue(i, out ChartSeriesFill explicitFill)
                ? explicitFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, i), 1d);
            if (fill.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            AppendPieOrDoughnutSlicePath(graphics, sliceCenterX, sliceCenterY, geometry.Radius, innerRadius, angle, sweep);
            graphics.FillCurrentPathEvenOdd();
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
                AppendPieOrDoughnutSlicePath(graphics, sliceCenterX, sliceCenterY, geometry.Radius, innerRadius, angle, sweep);
                graphics.StrokeCurrentPath();
                if (stroke.Alpha < 1d)
                {
                    graphics.RestoreState();
                }
            }

            angle += sweep;
        }
    }

    private static void AppendPieOrDoughnutSlicePath(PdfGraphicsBuilder graphics, double centerX, double centerY, double outerRadius, double innerRadius, double startAngle, double sweepAngle)
    {
        double outerStartX = centerX + Math.Cos(startAngle) * outerRadius;
        double outerStartY = centerY + Math.Sin(startAngle) * outerRadius;
        if (innerRadius <= 0d)
        {
            graphics.MoveTo(outerStartX, outerStartY);
            AppendCircularArc(graphics, centerX, centerY, outerRadius, startAngle, sweepAngle, moveToStart: false);
            graphics.LineTo(centerX, centerY);
            graphics.ClosePath();
            return;
        }

        double endAngle = startAngle + sweepAngle;
        double innerEndX = centerX + Math.Cos(endAngle) * innerRadius;
        double innerEndY = centerY + Math.Sin(endAngle) * innerRadius;
        graphics.MoveTo(outerStartX, outerStartY);
        AppendCircularArc(graphics, centerX, centerY, outerRadius, startAngle, sweepAngle, moveToStart: false);
        graphics.LineTo(innerEndX, innerEndY);
        AppendCircularArc(graphics, centerX, centerY, innerRadius, endAngle, -sweepAngle, moveToStart: false);
        graphics.ClosePath();
    }

    private static void AppendCircularArc(PdfGraphicsBuilder graphics, double centerX, double centerY, double radius, double startAngle, double sweepAngle, bool moveToStart)
    {
        int segmentCount = Math.Max(1, (int)Math.Ceiling(Math.Abs(sweepAngle) / (Math.PI / 2d)));
        double segmentSweep = sweepAngle / segmentCount;
        for (int segment = 0; segment < segmentCount; segment++)
        {
            double segmentStart = startAngle + segmentSweep * segment;
            AppendEllipseArcSegment(
                graphics,
                centerX,
                centerY,
                radius,
                radius,
                segmentStart * 180d / Math.PI,
                segmentSweep * 180d / Math.PI,
                moveToStart && segment == 0);
        }
    }

    private static void RenderDoughnutChart(PdfGraphicsBuilder graphics, PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<double> values, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, double firstSliceAngle)
    {
        RenderPieOrDoughnutSlices(graphics, theme, chartPalette, layout, values, pointFills, pointStrokes, pointExplosions, holeSize, firstSliceAngle);
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
        int? Join = null,
        PptxSceneLineCompound? Compound = null);

    private static ChartSeriesStroke ChartAxisDefaultStroke { get; } = new(new RgbColor(90, 90, 90), 1d, 0.75d);

    private static ChartSeriesStroke ChartNegativeBarDefaultStroke { get; } = new(new RgbColor(0, 0, 0), 1d, 0.75d);

    private static ChartSeriesStroke RadarGridlineDefaultStroke { get; } = new(new RgbColor(134, 134, 134), 1d, 0.75d, null, 0, 1);

    private static void DrawLineChartCategoryAxisMajorTicks(PdfGraphicsBuilder graphics, double plotX, double plotWidth, int pointCount, double axisY, PptxSceneChartAxisTickMark majorTickMark)
    {
        if (pointCount <= 0 || majorTickMark == PptxSceneChartAxisTickMark.None)
        {
            return;
        }

        double outward = majorTickMark == PptxSceneChartAxisTickMark.Cross
            ? PptxChartMetricRules.CategoryAxisMajorTickLength / 2d
            : PptxChartMetricRules.CategoryAxisMajorTickLength;
        double inward = majorTickMark == PptxSceneChartAxisTickMark.Inside || majorTickMark == PptxSceneChartAxisTickMark.Cross
            ? PptxChartMetricRules.CategoryAxisMajorTickLength / 2d
            : 0d;
        double slotWidth = plotWidth / pointCount;
        for (int i = 0; i <= pointCount; i++)
        {
            double x = plotX + slotWidth * i;
            graphics.StrokeLine(x, axisY - outward, x, axisY + inward);
        }
    }

    private readonly record struct ChartAxesStyle(ChartSeriesStroke? ValueAxis, ChartSeriesStroke? SecondaryValueAxis, ChartSeriesStroke? CategoryAxis, bool ValueAxisRightSide, bool SecondaryValueAxisRightSide, bool ValueAxisBottomSide, bool CategoryAxisRightSide, bool ValueAxisVisible, bool CategoryAxisVisible, PptxSceneChartAxisTickMark CategoryAxisMajorTickMark);

    private readonly record struct ChartGridlineStyle(ChartSeriesStroke? Major, ChartSeriesStroke? Minor)
    {
        public static ChartGridlineStyle Empty { get; } = new(null, null);
    }

    private readonly record struct ChartLayout(ChartFrameBox Frame, ChartLayoutBox PlotAreaBox, ChartPlotBox PlotBox, bool ManualPlotLayoutApplied, string? Title, ChartLegendLayout Legend);

    private readonly record struct ChartFrameBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartLayoutBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartPlotBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartAxisSource(PptxSceneChartAxis? SceneAxis, XElement? XmlAxis);

    private readonly record struct ChartPlotLayout(ChartLayoutBox PlotAreaBox, ChartPlotBox PlotBox, PptxSceneChartManualLayoutTarget? ManualLayoutTargetKind)
    {
        public static ChartPlotLayout FromPlotBox(ChartPlotBox plotBox)
        {
            return new ChartPlotLayout(new ChartLayoutBox(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height), plotBox, null);
        }
    }

    private readonly record struct ChartValueExtents(double Min, double Max);

    private enum ChartAxisScalingBound
    {
        Minimum,
        Maximum
    }

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

    private readonly record struct ChartDataLabelOptions(bool ShowValue, bool ShowPercent, bool ShowCategoryName, bool ShowSeriesName, bool ShowLeaderLines, ChartDataLabelLeaderLines LeaderLines, string CustomText, IReadOnlyList<ChartTextRunOverride> CustomTextRuns, PptxSceneChartDataLabelPosition PositionKind, string Position, string Separator, string NumberFormat, ChartNumberFormat NumberFormatInfo, PptxSceneChartManualLayout Layout, ChartTextStyleOverride TextStyle, ChartShapeStyle ShapeStyle, IReadOnlyDictionary<int, ChartDataLabelOverride> Overrides, bool IsDefined)
    {
        public static ChartDataLabelOptions None { get; } = new(ShowValue: false, ShowPercent: false, ShowCategoryName: false, ShowSeriesName: false, ShowLeaderLines: false, LeaderLines: ChartDataLabelLeaderLines.Empty, CustomText: string.Empty, CustomTextRuns: [], PositionKind: PptxSceneChartDataLabelPosition.Unknown, Position: string.Empty, Separator: string.Empty, NumberFormat: string.Empty, NumberFormatInfo: default, Layout: default, TextStyle: ChartTextStyleOverride.Empty, ShapeStyle: ChartShapeStyle.Empty, Overrides: EmptyChartDataLabelOverrides, IsDefined: false);

        public bool HasVisibleText => ShowValue || ShowPercent || ShowCategoryName || ShowSeriesName ||
            !string.IsNullOrWhiteSpace(CustomText) ||
            Overrides.Values.Any(label => label.ShowValue == true || label.ShowPercent == true || label.ShowCategoryName == true || label.ShowSeriesName == true || !string.IsNullOrWhiteSpace(label.CustomText));
    }

    private readonly record struct ChartDataLabelLeaderLines(bool IsDefined, ChartSeriesStroke? Stroke)
    {
        public static ChartDataLabelLeaderLines Empty { get; } = new(IsDefined: false, Stroke: null);
    }

    private readonly record struct ChartTextRunOverride(string Text, ChartTextStyleOverride TextStyle);

    private readonly record struct ChartTextRunLayout(string Text, ChartTextStyle Style, double Width);

    private readonly record struct ChartDataLabelOverride(bool? ShowValue, bool? ShowPercent, bool? ShowCategoryName, bool? ShowSeriesName, bool? ShowLeaderLines, ChartDataLabelLeaderLines LeaderLines, string CustomText, IReadOnlyList<ChartTextRunOverride> CustomTextRuns, PptxSceneChartDataLabelPosition PositionKind, string Position, string Separator, string NumberFormat, ChartNumberFormat NumberFormatInfo, PptxSceneChartManualLayout Layout, ChartTextStyleOverride TextStyle, ChartShapeStyle ShapeStyle);

    private readonly record struct ChartNumberFormat(bool IsDefined, string FormatCode, bool? SourceLinked);

    private readonly record struct ChartLegendEntry(string Name, ChartSeriesFill? Fill, ChartSeriesStroke? Stroke);

    private readonly record struct ChartLegendLayout(PptxSceneChartLegendPosition PositionKind, string Position, bool Overlay, bool Visible, PptxSceneChartManualLayout Layout, ChartShapeStyle ShapeStyle)
    {
        public static ChartLegendLayout Hidden { get; } = new(PptxSceneChartLegendPosition.Right, "r", Overlay: false, Visible: false, default, ChartShapeStyle.Empty);
    }

    private readonly record struct ChartShapeStyle(ChartSeriesFill? Fill, GradientFill? GradientFill, ChartSeriesStroke? Stroke)
    {
        public static ChartShapeStyle Empty { get; } = new(null, null, null);

        public bool IsEmpty => Fill is null && GradientFill is null && Stroke is null;
    }

    private readonly record struct ChartMarkerStyle(PptxSceneChartMarkerSymbol SymbolKind, string Symbol, double Size, ChartSeriesFill? Fill, ChartSeriesStroke? Stroke)
    {
        public static ChartMarkerStyle Default { get; } = new(PptxSceneChartMarkerSymbol.Circle, "circle", 4d, null, null);
    }
}
