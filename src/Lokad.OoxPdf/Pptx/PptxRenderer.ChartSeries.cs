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
    private static void RenderChartFrame(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        PptxSceneNode node,
        GroupTransform transform,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        context.CancellationToken.ThrowIfCancellationRequested();
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        RenderChartFrame(context, graphics, fonts, bounds, node.Chart, linkAnnotations, reportedHyperlinkIds);
    }

    private static void RenderChartFrame(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        ShapeBounds? bounds,
        PptxSceneChart? chart,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        RenderChartFrame(
            context,
            graphics,
            fonts,
            bounds,
            chart?.TargetPartName,
            chart?.ChartXml,
            chart?.PaletteColors,
            chart,
            linkAnnotations,
            reportedHyperlinkIds);
    }

    private static void RenderChartFrame(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        ShapeBounds? bounds,
        string? targetPartName,
        XDocument? chartXml,
        IReadOnlyList<RgbColor>? chartPalette,
        PptxSceneChart? sceneChart,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        context.CancellationToken.ThrowIfCancellationRequested();
        string? chartPartName = targetPartName;

        if (bounds is null || chartPartName is null)
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart frame could not be resolved and was ignored.", context.SlidePartName, context.SlideNumber, "Ignored");
            return;
        }

        XDocument? resolvedChartXml = chartXml;
        IReadOnlyList<RgbColor>? resolvedChartPalette = chartPalette;
        if (resolvedChartXml is null)
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart part was missing from the scene model and was ignored.", chartPartName, context.SlideNumber, "Ignored");
            return;
        }

        ChartWorkbookData? chartWorkbook = ReadEmbeddedChartWorkbookData(sceneChart?.ExternalData ?? default, context.CancellationToken);
        context.CancellationToken.ThrowIfCancellationRequested();

        PptxColorMap chartColorMap = sceneChart?.ColorMap ?? context.SlideColorMap;
        if (TryRenderChart(graphics, context.Document, context.Theme, chartColorMap, resolvedChartPalette, bounds.Value, resolvedChartXml, sceneChart, chartWorkbook, fonts, context.FontResolver, context, linkAnnotations, reportedHyperlinkIds))
        {
            EmitUnrenderedDefaultChartAxisTitleDiagnostics(resolvedChartXml, sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
            EmitUnsupportedChartNumberFormatDiagnostics(resolvedChartXml, sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
            RenderManualChartAxisTitles(context.Document, context.Theme, chartColorMap, graphics, bounds.Value, resolvedChartXml, sceneChart, context.FontResolver, context.DiagnosticSink, chartPartName, context.SlideNumber, emitDefaultLayoutDiagnostics: false, fonts, context, linkAnnotations, reportedHyperlinkIds);
            RenderChartTitle(context.Document, context.Theme, chartColorMap, graphics, bounds.Value, resolvedChartXml, sceneChart, chartWorkbook, ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, resolvedChartXml), context.FontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds, context.DiagnosticSink);
            return;
        }

        if (chartWorkbook is not null && sceneChart is null)
        {
            HydrateChartReferenceCaches(chartWorkbook, resolvedChartXml);
            if (TryRenderChart(graphics, context.Document, context.Theme, chartColorMap, resolvedChartPalette, bounds.Value, resolvedChartXml, sceneChart, workbook: null, fonts, context.FontResolver, context, linkAnnotations, reportedHyperlinkIds))
            {
                EmitUnrenderedDefaultChartAxisTitleDiagnostics(resolvedChartXml, sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
                EmitUnsupportedChartNumberFormatDiagnostics(resolvedChartXml, sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
                RenderManualChartAxisTitles(context.Document, context.Theme, chartColorMap, graphics, bounds.Value, resolvedChartXml, sceneChart, context.FontResolver, context.DiagnosticSink, chartPartName, context.SlideNumber, emitDefaultLayoutDiagnostics: false, fonts, context, linkAnnotations, reportedHyperlinkIds);
                RenderChartTitle(context.Document, context.Theme, chartColorMap, graphics, bounds.Value, resolvedChartXml, sceneChart, workbook: null, ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, resolvedChartXml), context.FontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds, context.DiagnosticSink);
                return;
            }
        }

        bool HasSupportedSceneChartWithoutRenderableCachedValues()
        {
            bool IsSupportedNativeChartPlot(PptxSceneChartPlotKind kind)
            {
                return kind is PptxSceneChartPlotKind.Area or
                    PptxSceneChartPlotKind.Bar or
                    PptxSceneChartPlotKind.Bubble or
                    PptxSceneChartPlotKind.Doughnut or
                    PptxSceneChartPlotKind.Line or
                    PptxSceneChartPlotKind.Pie or
                    PptxSceneChartPlotKind.Radar or
                    PptxSceneChartPlotKind.Scatter;
            }

            bool HasRenderableCachedValues(PptxSceneChartSeries series)
            {
                return series.Values.Count != 0 ||
                    series.ValuePoints.Any(point => point.Value is not null) ||
                    (series.XValues.Count != 0 && series.YValues.Count != 0) ||
                    (series.XValuePoints.Any(point => point.Value is not null) &&
                        series.YValuePoints.Any(point => point.Value is not null));
            }

            return sceneChart?.Plots.Any(plot =>
                IsSupportedNativeChartPlot(plot.PlotKind) &&
                plot.Series.Count != 0 &&
                !plot.Series.Any(HasRenderableCachedValues)) == true;
        }

        if (HasSupportedSceneChartWithoutRenderableCachedValues())
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_CHART_MISSING_CACHED_DATA", OoxPdfSeverity.Warning, "Supported chart references formula-only data without chart-side cached numeric values. Embedded workbook values are preserved as provenance but are not used as active rendering data.", chartPartName, context.SlideNumber, "Ignored");
            return;
        }

        EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Only bar, line, area, scatter, bubble, radar, pie, and doughnut charts with cached numeric values are currently supported by the native chart renderer.", chartPartName, context.SlideNumber, "Ignored");
    }

    private static PptxSceneChartPlot? ReadSceneChartPlot(PptxSceneChart? chart, PptxSceneChartPlotKind kind, int index)
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
        string? elementName = PptxSceneBuilder.GetChartPlotElementName(kind);
        return elementName is null ? [] : ReadChartPlotElements(chartXml, elementName);
    }

    private static XElement? ReadSceneOrXmlFirstChartPlotElement(PptxSceneChart? sceneChart, XDocument chartXml, PptxSceneChartPlotKind kind)
    {
        if (sceneChart is not null)
        {
            return ReadSceneChartPlot(sceneChart, kind, 0)?.Source;
        }

        return ReadChartPlotElements(chartXml, kind).FirstOrDefault();
    }

    private static IReadOnlyList<XElement> ReadSceneOrXmlChartPlotElements(PptxSceneChart? sceneChart, XDocument chartXml, PptxSceneChartPlotKind kind)
    {
        if (sceneChart is not null)
        {
            return ReadSceneChartPlots(sceneChart, kind)
                .Select(plot => plot.Source)
                .ToArray();
        }

        return ReadChartPlotElements(chartXml, kind);
    }

    private static IReadOnlyList<PptxSceneChartPlot> ReadSceneChartPlots(PptxSceneChart? chart, PptxSceneChartPlotKind kind)
    {
        return chart?.Plots
            .Where(plot => plot.PlotKind == kind)
            .ToArray() ?? [];
    }

    private static IReadOnlyList<ChartAxisSource> ReadSceneOrXmlChartValueAxesForPlot(
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? scenePlot,
        XDocument chartXml,
        XElement? chartElement)
    {
        IReadOnlyList<PptxSceneChartAxis> ReadSceneChartAxes(PptxSceneChart? chart, PptxSceneChartPlot? plot, PptxSceneChartAxisKind kind)
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
                    .OfType<PptxSceneChartAxis>()
                    .ToArray();
            }
    
            return chart.Axes
                .Where(axis => axis.AxisKind == kind)
                .ToArray();
        }

        IReadOnlyList<PptxSceneChartAxis> sceneAxes = ReadSceneChartAxes(sceneChart, scenePlot, PptxSceneChartAxisKind.Value);
        if (sceneAxes.Count == 0)
        {
            if (sceneChart is not null)
            {
                return [];
            }

            XElement[] fallbackXmlAxes = chartElement is null
                ? chartXml.Descendants(ChartNamespace + "valAx").ToArray()
                : ReadChartValueAxesForChart(chartXml, chartElement).ToArray();
            return fallbackXmlAxes
                .Select(axis => new ChartAxisSource(null, axis))
                .ToArray();
        }

        XElement[] ReadSceneChartValueAxisElements()
        {
            if (sceneChart.ChartXml is null)
            {
                return [];
            }
    
            return scenePlot is null
                ? sceneChart.ChartXml.Descendants(ChartNamespace + "valAx").ToArray()
                : ReadChartValueAxesForChart(sceneChart.ChartXml, scenePlot.Source).ToArray();
        }

        XElement[] xmlAxes = sceneChart is not null
            ? ReadSceneChartValueAxisElements()
            : chartElement is null
                ? chartXml.Descendants(ChartNamespace + "valAx").ToArray()
                : ReadChartValueAxesForChart(chartXml, chartElement).ToArray();
        var sources = new List<ChartAxisSource>(sceneAxes.Count);
        foreach (PptxSceneChartAxis sceneAxis in sceneAxes)
        {
            XElement? xmlAxis = xmlAxes.FirstOrDefault(axis => string.Equals(ReadChartAxisId(axis), sceneAxis.Id, StringComparison.Ordinal));
            sources.Add(new ChartAxisSource(sceneAxis, xmlAxis));
        }

        return sources;
    }

    private static ChartAxisSource ReadSceneOrXmlChartCategoryAxisForPlot(
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? scenePlot,
        XDocument chartXml,
        XElement? chartElement)
    {
        IReadOnlyList<PptxSceneChartAxis> ReadSceneChartCategoryAxes(PptxSceneChart? chart, PptxSceneChartPlot? plot)
        {
            if (chart is null)
            {
                return [];
            }
    
            static bool IsCategoryLike(PptxSceneChartAxis axis)
            {
                return axis.AxisKind is PptxSceneChartAxisKind.Category or PptxSceneChartAxisKind.Date;
            }
    
            if (plot is not null && plot.AxisIds.Count != 0)
            {
                return plot.AxisIds
                    .Select(axisId => chart.Axes.FirstOrDefault(candidate =>
                        string.Equals(candidate.Id, axisId, StringComparison.Ordinal) &&
                        IsCategoryLike(candidate)))
                    .Where(axis => axis is not null)
                    .OfType<PptxSceneChartAxis>()
                    .ToArray();
            }
    
            return chart.Axes
                .Where(IsCategoryLike)
                .ToArray();
        }

        IReadOnlyList<PptxSceneChartAxis> sceneAxes = ReadSceneChartCategoryAxes(sceneChart, scenePlot);
        if (sceneAxes.Count == 0)
        {
            if (sceneChart is not null)
            {
                return new ChartAxisSource(null, null);
            }

            XElement[] fallbackXmlAxes = chartElement is null
                ? ReadChartCategoryAxes(chartXml).ToArray()
                : ReadChartCategoryAxesForChart(chartXml, chartElement).ToArray();
            XElement? xmlAxis = fallbackXmlAxes.FirstOrDefault();
            return new ChartAxisSource(null, xmlAxis);
        }

        XElement[] ReadSceneChartCategoryAxisElements()
        {
            if (sceneChart.ChartXml is null)
            {
                return [];
            }
    
            return scenePlot is null
                ? ReadChartCategoryAxes(sceneChart.ChartXml).ToArray()
                : ReadChartCategoryAxesForChart(sceneChart.ChartXml, scenePlot.Source).ToArray();
        }

        XElement[] xmlAxes = sceneChart is not null
            ? ReadSceneChartCategoryAxisElements()
            : chartElement is null
                ? ReadChartCategoryAxes(chartXml).ToArray()
                : ReadChartCategoryAxesForChart(chartXml, chartElement).ToArray();
        PptxSceneChartAxis sceneAxis = sceneAxes[0];
        XElement? matchedXmlAxis = xmlAxes.FirstOrDefault(axis => string.Equals(ReadChartAxisId(axis), sceneAxis.Id, StringComparison.Ordinal));
        return new ChartAxisSource(sceneAxis, matchedXmlAxis);
    }

    private static ChartAxisSource ReadSceneOrXmlSecondaryRightValueAxis(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is null)
        {
            return new ChartAxisSource(null, ReadSecondaryRightValueAxis(chartXml));
        }

        PptxSceneChartAxis? sceneAxis = sceneChart.Axes.FirstOrDefault(axis =>
            axis.AxisKind == PptxSceneChartAxisKind.Value &&
            axis.IsDeleted != true &&
            axis.PositionKind == PptxSceneChartAxisPosition.Right);
        XElement? sceneXmlAxis = sceneChart.ChartXml?
            .Descendants(ChartNamespace + "valAx")
            .FirstOrDefault(axis => string.Equals(ReadChartAxisId(axis), sceneAxis?.Id, StringComparison.Ordinal));
        return sceneAxis is null
            ? new ChartAxisSource(null, null)
            : new ChartAxisSource(sceneAxis, sceneXmlAxis);
    }

    private static ChartAxisSource ReadSceneOrXmlSecondaryValueAxisForChart(PptxSceneChart? sceneChart, XDocument chartXml, ChartAxisSource primaryValueAxis)
    {
        if (sceneChart is null)
        {
            return new ChartAxisSource(null, ReadSecondaryValueAxisForChart(chartXml, primaryValueAxis.XmlAxis));
        }

        string? primaryAxisId = primaryValueAxis.SceneAxis?.Id;
        PptxSceneChartAxis? sceneAxis = !string.IsNullOrWhiteSpace(primaryAxisId)
            ? sceneChart.Axes.FirstOrDefault(axis =>
                axis.AxisKind == PptxSceneChartAxisKind.Value &&
                !string.Equals(axis.Id, primaryAxisId, StringComparison.Ordinal) &&
                axis.IsDeleted != true)
            : sceneChart.Axes.FirstOrDefault(axis =>
                axis.AxisKind == PptxSceneChartAxisKind.Value &&
                axis.IsDeleted != true &&
                axis.PositionKind == PptxSceneChartAxisPosition.Right);
        XElement? sceneXmlAxis = sceneChart.ChartXml?
            .Descendants(ChartNamespace + "valAx")
            .FirstOrDefault(axis => string.Equals(ReadChartAxisId(axis), sceneAxis?.Id, StringComparison.Ordinal));
        return sceneAxis is null
            ? new ChartAxisSource(null, null)
            : new ChartAxisSource(sceneAxis, sceneXmlAxis);
    }

    private static XElement? ResolveXmlValueAxisForSource(PptxSceneChart? sceneChart, ChartAxisSource source, XDocument chartXml)
    {
        if (source.XmlAxis is not null)
        {
            return source.XmlAxis;
        }

        return sceneChart is null && source.SceneAxis is null
            ? chartXml.Descendants(ChartNamespace + "valAx").FirstOrDefault()
            : null;
    }

    private static PptxSceneChartGrouping ReadSceneOrXmlChartGrouping(PptxSceneChartPlot? scenePlot, XElement plotElement, PptxSceneChartGrouping defaultGrouping)
    {
        if (scenePlot is not null)
        {
            return ResolveGrouping(scenePlot.GroupingKind);
        }

        return ResolveGrouping(
            PptxSceneBuilder.ParseChartGrouping(PptxSceneBuilder.ReadChartElementValue(plotElement, "grouping")));

        PptxSceneChartGrouping ResolveGrouping(PptxSceneChartGrouping value)
        {
            return value == PptxSceneChartGrouping.Unknown
                ? defaultGrouping
                : value;
        }
    }

    private static PptxSceneChartBarDirection ReadSceneOrXmlChartBarDirection(PptxSceneChartPlot? scenePlot, XElement plotElement)
    {
        return scenePlot is not null
            ? ResolveChartBarDirection(scenePlot.BarDirectionKind)
            : ResolveChartBarDirection(PptxSceneBuilder.ParseChartBarDirection(PptxSceneBuilder.ReadChartElementValue(plotElement, "barDir")));

        PptxSceneChartBarDirection ResolveChartBarDirection(PptxSceneChartBarDirection value)
        {
            return value == PptxSceneChartBarDirection.Unknown
                ? PptxSceneChartBarDirection.Column
                : value;
        }
    }

    private static PptxSceneChartScatterStyle ReadSceneOrXmlChartScatterStyle(PptxSceneChartPlot? scenePlot, XElement plotElement)
    {
        return scenePlot is not null
            ? scenePlot.ScatterStyleKind
            : PptxSceneBuilder.ParseChartScatterStyle(PptxSceneBuilder.ReadChartElementValue(plotElement, "scatterStyle"));
    }

    private static bool ResolveChartScatterLineConnection(PptxSceneChartScatterStyle scatterStyle)
    {
        return scatterStyle switch
        {
            PptxSceneChartScatterStyle.Line or PptxSceneChartScatterStyle.LineMarker => true,
            _ => false
        };
    }

    private static PptxSceneChartRadarStyle ReadSceneOrXmlChartRadarStyle(PptxSceneChartPlot? scenePlot, XElement plotElement)
    {
        return scenePlot is not null
            ? ResolveChartRadarStyle(scenePlot.RadarStyleKind)
            : ResolveChartRadarStyle(PptxSceneBuilder.ParseChartRadarStyle(PptxSceneBuilder.ReadChartElementValue(plotElement, "radarStyle")));

        PptxSceneChartRadarStyle ResolveChartRadarStyle(PptxSceneChartRadarStyle value)
        {
            return value == PptxSceneChartRadarStyle.Unknown
                ? PptxSceneChartRadarStyle.Standard
                : value;
        }
    }

    private static PptxSceneChartDisplayBlanksAs ReadSceneOrXmlChartDisplayBlanksAs(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            return ResolveChartDisplayBlanksAs(sceneChart.Options.DisplayBlanksAsKind);
        }

        return ResolveChartDisplayBlanksAs(PptxSceneBuilder.ReadChartOptions(chartXml).DisplayBlanksAsKind);

        PptxSceneChartDisplayBlanksAs ResolveChartDisplayBlanksAs(PptxSceneChartDisplayBlanksAs value)
        {
            return value == PptxSceneChartDisplayBlanksAs.Unknown
                ? PptxSceneChartDisplayBlanksAs.Gap
                : value;
        }
    }

    private static bool ReadSceneOrXmlChartPlotVisibleOnly(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            return sceneChart.Options.PlotVisibleOnly ?? true;
        }

        return PptxSceneBuilder.ReadChartOptions(chartXml).PlotVisibleOnly ?? true;
    }

    private static IReadOnlyList<ChartRadarSeries> BuildRadarSeries(IEnumerable<ChartIndexedNumberVector> series)
    {
        return series
            .Select(vector => new ChartRadarSeries(vector.DensePoints(), vector))
            .Where(item => item.Points.Any(point => point?.Value is not null))
            .ToArray();
    }

    private static int CountRenderableSeries(IEnumerable<ChartIndexedNumberVector> series)
    {
        return series.Count(vector => vector.DensePoints().Any(point => point?.Value is not null));
    }

    private static IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> DensifyChartPointSeries(IEnumerable<ChartIndexedNumberVector> series)
    {
        return series
            .Select(vector => vector.DensePoints())
            .ToArray();
    }

    private static IReadOnlyList<ChartIndexedNumberVector> ReadSceneOrXmlChartSeriesVectors(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => BuildChartIndexedNumberVector(
                    series.Values,
                    series.ValuePoints,
                    series.ValuePointCount,
                    series.ValueFormatCode,
                    series.DataSources.Values,
                    workbook,
                    plotVisibleOnly))
                .ToArray();
        }

        return ReadChartSeriesVectors(chartElement, workbook, plotVisibleOnly);
    }

    private static IReadOnlyList<ScatterSeries> ReadSceneOrXmlScatterSeries(PptxSceneChartPlot? plot, XElement chartElement, bool readBubbleSize, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        return ReadSceneOrXmlScatterSeriesVectors(plot, chartElement, readBubbleSize, workbook, plotVisibleOnly)
            .Select(BuildScatterSeries)
            .Where(series => series.Points.Count != 0)
            .ToArray();
    }

    private static IReadOnlyList<ChartIndexedScatterSeries> ReadSceneOrXmlScatterSeriesVectors(PptxSceneChartPlot? plot, XElement chartElement, bool readBubbleSize, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        if (plot is null)
        {
            return ReadScatterSeriesVectors(chartElement, readBubbleSize, workbook, plotVisibleOnly);
        }

        return plot.Series
            .Select(item => new ChartIndexedScatterSeries(
                BuildChartIndexedNumberVector(
                    item.XValues,
                    item.XValuePoints,
                    item.XValuePointCount,
                    item.XValueFormatCode,
                    item.DataSources.XValues,
                    workbook,
                    plotVisibleOnly),
                BuildChartIndexedNumberVector(
                    item.YValues,
                    item.YValuePoints,
                    item.YValuePointCount,
                    item.YValueFormatCode,
                    item.DataSources.YValues,
                    workbook,
                    plotVisibleOnly),
                BuildChartIndexedNumberVector(
                    item.BubbleSizes,
                    item.BubbleSizePoints,
                    item.BubbleSizePointCount,
                    item.BubbleSizeFormatCode,
                    item.DataSources.BubbleSizes,
                    workbook,
                    plotVisibleOnly),
                readBubbleSize))
            .ToArray();
    }

    private static ChartIndexedTextVector ReadSceneOrXmlCategoryLabelVector(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => BuildChartIndexedTextVector(
                    series.Categories,
                    series.CategoryPoints,
                    series.CategoryPointCount,
                    series.CategoryLevels,
                    series.DataSources.Categories,
                    workbook,
                    plotVisibleOnly))
                .FirstOrDefault(vector => vector.Points.Count != 0 || vector.PointCount is not null || vector.DensePoints().Count != 0);
        }

        return ReadChartCategoryLabelVector(chartElement, workbook, plotVisibleOnly);
    }

    private static ScatterSeries BuildScatterSeries(ChartIndexedScatterSeries series)
    {
        IReadOnlyList<ChartIndexedNumberPoint?> xPoints = series.XValues.DensePoints();
        IReadOnlyList<ChartIndexedNumberPoint?> yPoints = series.YValues.DensePoints();
        IReadOnlyList<ChartIndexedNumberPoint?> bubbleSizePoints = series.BubbleSizes.DensePoints();
        int count = Math.Max(xPoints.Count, yPoints.Count);
        if (count == 0)
        {
            return new ScatterSeries([], series);
        }

        var points = new List<ScatterPoint>(count);
        for (int i = 0; i < count; i++)
        {
            ChartIndexedNumberPoint? xPoint = i < xPoints.Count ? xPoints[i] : null;
            ChartIndexedNumberPoint? yPoint = i < yPoints.Count ? yPoints[i] : null;
            if (xPoint?.Value is not { } xValue || yPoint?.Value is not { } yValue)
            {
                continue;
            }

            ChartIndexedNumberPoint? bubbleSizePoint = series.ReadBubbleSize && i < bubbleSizePoints.Count ? bubbleSizePoints[i] : null;
            double size = bubbleSizePoint?.Value ?? 1d;
            points.Add(new ScatterPoint(
                xValue,
                yValue,
                size,
                yPoint.Value.Index,
                xPoint.Value,
                yPoint.Value,
                bubbleSizePoint,
                series.XValues.WorkbookPointForIndex(xPoint.Value.Index),
                series.YValues.WorkbookPointForIndex(yPoint.Value.Index),
                bubbleSizePoint is { } point ? series.BubbleSizes.WorkbookPointForIndex(point.Index) : null,
                series.YValues.FormatCode,
                series.BubbleSizes.FormatCode));
        }

        return new ScatterSeries(points, series);
    }

    private static ChartIndexedNumberVector BuildChartIndexedNumberVector(
        IReadOnlyList<double> compactValues,
        IReadOnlyList<PptxSceneChartNumberPoint> scenePoints,
        int? pointCount,
        string? formatCode,
        PptxSceneChartDataSource source,
        ChartWorkbookData? workbook)
    {
        return BuildChartIndexedNumberVector(compactValues, scenePoints, pointCount, formatCode, source, workbook, plotVisibleOnly: true);
    }

    private static ChartIndexedNumberVector BuildChartIndexedNumberVector(
        IReadOnlyList<double> compactValues,
        IReadOnlyList<PptxSceneChartNumberPoint> scenePoints,
        int? pointCount,
        string? formatCode,
        PptxSceneChartDataSource source,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        IReadOnlyList<ChartIndexedNumberPoint> workbookPoints = ReadWorkbookNumberPoints(workbook, source);
        IReadOnlyList<ChartIndexedNumberPoint> points = scenePoints.Count != 0
            ? scenePoints
                .Select(ToChartIndexedNumberPoint)
                .ToArray()
            : compactValues
                .Select((value, index) => new ChartIndexedNumberPoint(index, ChartPointIndexSource.OrdinalFallback, value, value.ToString(CultureInfo.InvariantCulture), true, default))
                .ToArray();
        return new ChartIndexedNumberVector(points, pointCount ?? InferPointCount(points), source.Formula, formatCode, source, workbookPoints, plotVisibleOnly);
    }

    private static ChartIndexedNumberPoint ToChartIndexedNumberPoint(PptxSceneChartNumberPoint point)
    {
        return new ChartIndexedNumberPoint(
            point.Index,
            point.HasParsedIndex ? ChartPointIndexSource.OoxmlIndex : ChartPointIndexSource.OrdinalFallback,
            point.Value,
            point.Text,
            point.HasValueElement,
            default);
    }

    private static ChartIndexedTextVector BuildChartIndexedTextVector(
        IReadOnlyList<string> compactValues,
        IReadOnlyList<PptxSceneChartStringPoint> scenePoints,
        int? pointCount,
        IReadOnlyList<IReadOnlyList<PptxSceneChartStringPoint>> categoryLevels,
        PptxSceneChartDataSource source,
        ChartWorkbookData? workbook)
    {
        return BuildChartIndexedTextVector(compactValues, scenePoints, pointCount, categoryLevels, source, workbook, plotVisibleOnly: true);
    }

    private static ChartIndexedTextVector BuildChartIndexedTextVector(
        IReadOnlyList<string> compactValues,
        IReadOnlyList<PptxSceneChartStringPoint> scenePoints,
        int? pointCount,
        IReadOnlyList<IReadOnlyList<PptxSceneChartStringPoint>> categoryLevels,
        PptxSceneChartDataSource source,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        IReadOnlyList<ChartIndexedTextPoint> workbookPoints = ReadWorkbookTextPoints(workbook, source);
        IReadOnlyList<ChartIndexedTextPoint> points = scenePoints.Count != 0
            ? scenePoints
                .Select(point => ToChartIndexedTextPoint(point, trimText: false))
                .ToArray()
            : compactValues
                .Select((value, index) => new ChartIndexedTextPoint(index, ChartPointIndexSource.OrdinalFallback, value, true, default))
                .ToArray();
        IReadOnlyList<IReadOnlyList<ChartIndexedTextPoint>> levels = categoryLevels
            .Select(level => level.Select(point => ToChartIndexedTextPoint(point, trimText: false)).ToArray())
            .ToArray();
        return new ChartIndexedTextVector(points, pointCount ?? InferPointCount(points), levels, source.Formula, source, workbookPoints, plotVisibleOnly);
    }

    private static ChartIndexedTextPoint ToChartIndexedTextPoint(PptxSceneChartStringPoint point, bool trimText)
    {
        return new ChartIndexedTextPoint(
            point.Index,
            point.HasParsedIndex ? ChartPointIndexSource.OoxmlIndex : ChartPointIndexSource.OrdinalFallback,
            trimText ? point.Text.Trim() : point.Text,
            point.HasText,
            default);
    }

    private static IReadOnlyList<ChartIndexedNumberPoint> ReadWorkbookNumberPoints(ChartWorkbookData? workbook, PptxSceneChartDataSource source)
    {
        return workbook?
            .ReadRangeCells(source.Formula)
            .Select(cell => new ChartIndexedNumberPoint(
                cell.Index,
                ChartPointIndexSource.WorkbookRange,
                double.TryParse(cell.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : null,
                cell.Text,
                cell.HasValue,
                cell))
            .ToArray() ?? [];
    }

    private static IReadOnlyList<ChartIndexedTextPoint> ReadWorkbookTextPoints(ChartWorkbookData? workbook, PptxSceneChartDataSource source)
    {
        return workbook?
            .ReadRangeCells(source.Formula)
            .Select(cell => new ChartIndexedTextPoint(cell.Index, ChartPointIndexSource.WorkbookRange, cell.Text, cell.HasValue && !string.IsNullOrWhiteSpace(cell.Text), cell))
            .ToArray() ?? [];
    }

    private static int? InferPointCount(IReadOnlyList<ChartIndexedNumberPoint> points)
    {
        return points.Count == 0 ? null : points.Max(point => point.Index) + 1;
    }

    private static int? InferPointCount(IReadOnlyList<ChartIndexedTextPoint> points)
    {
        return points.Count == 0 ? null : points.Max(point => point.Index) + 1;
    }

    private static IReadOnlyList<ChartSeriesFill?> ReadSceneOrXmlSeriesFills(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => ToChartSeriesFill(series.Fill, series.PatternFill))
                .ToArray();
        }

        return ReadXmlChartSeries(chartElement, theme, colorMap)
            .Select(series => ToChartSeriesFill(series.Fill, series.PatternFill))
            .ToArray();
    }

    private static IReadOnlyList<ChartSeriesStroke?> ReadSceneOrXmlSeriesStrokes(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap, double? inheritedWidth)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => ToChartSeriesStroke(series.Line, inheritedWidth))
                .ToArray();
        }

        return ReadXmlChartSeries(chartElement, theme, colorMap)
            .Select(series => ToChartSeriesStroke(series.Line, inheritedWidth))
            .ToArray();
    }

    private static IReadOnlyList<ChartMarkerStyle> ReadSceneOrXmlMarkerStyles(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => new ChartMarkerStyle(
                    series.Marker.SymbolKind,
                    series.Marker.Symbol,
                    series.Marker.SizeValue,
                    series.Marker.Size,
                    series.Marker.Fill.HasFill ? new ChartSeriesFill(series.Marker.Fill.Color, series.Marker.Fill.Alpha, null, null) : null,
                    ToChartSeriesStroke(series.Marker.Line, ChartMarkerInheritedStrokeWidth),
                    series.Marker.IsDefined))
                .ToArray();
        }

        return ReadXmlChartSeries(chartElement, theme, colorMap)
            .Select(series => new ChartMarkerStyle(
                series.Marker.SymbolKind,
                series.Marker.Symbol,
                series.Marker.SizeValue,
                series.Marker.Size,
                series.Marker.Fill.HasFill ? new ChartSeriesFill(series.Marker.Fill.Color, series.Marker.Fill.Alpha, null, null) : null,
                ToChartSeriesStroke(series.Marker.Line, ChartMarkerInheritedStrokeWidth),
                series.Marker.IsDefined))
            .ToArray();
    }

    private static IReadOnlyList<ChartBooleanOption> ReadSceneOrXmlSmoothSeries(PptxSceneChartPlot? plot, XElement chartElement)
    {
        return plot is not null
            ? plot.Series.Select(series => new ChartBooleanOption(series.Smooth ?? false, series.SmoothValue, series.Smooth is not null)).ToArray()
            : chartElement
                .Elements(ChartNamespace + "ser")
                .Select(series =>
                {
                    (bool? smooth, string smoothValue) = PptxSceneBuilder.ReadChartSeriesSmooth(series);
                    return new ChartBooleanOption(smooth ?? false, smoothValue, smooth is not null);
                })
                .ToArray();
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> ReadSceneOrXmlSeriesPointFills(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return plot is not null
            ? plot.Series.Select(ReadSceneChartPointFills).ToArray()
            : ReadXmlChartSeries(chartElement, theme, colorMap).Select(ReadSceneChartPointFills).ToArray();
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> ReadSceneOrXmlSeriesPointStrokes(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return plot is not null
            ? plot.Series.Select(ReadSceneChartPointStrokes).ToArray()
            : ReadXmlChartSeries(chartElement, theme, colorMap).Select(ReadSceneChartPointStrokes).ToArray();
    }

    private static IReadOnlyDictionary<int, double> ReadSceneOrXmlChartPointExplosions(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook)
    {
        IReadOnlyDictionary<int, double> ReadSceneChartPointExplosions(PptxSceneChartSeries series, ChartWorkbookData? workbook)
        {
            var explosions = new Dictionary<int, double>();
            if (series.Explosion is { } seriesExplosion)
            {
                double fraction = Math.Clamp(seriesExplosion / 100d, 0d, 1d);
                int ResolveSceneChartSeriesPointCount()
                {
                    ChartIndexedNumberVector values = BuildChartIndexedNumberVector(
                        series.Values,
                        series.ValuePoints,
                        series.ValuePointCount,
                        series.ValueFormatCode,
                        series.DataSources.Values,
                        workbook,
                        plotVisibleOnly: true);
                    ChartIndexedTextVector categories = BuildChartIndexedTextVector(
                        series.Categories,
                        series.CategoryPoints,
                        series.CategoryPointCount,
                        series.CategoryLevels,
                        series.DataSources.Categories,
                        workbook,
                        plotVisibleOnly: true);
                    return Math.Max(values.PointCount ?? 0, categories.PointCount ?? 0);
                }
    
                int pointCount = ResolveSceneChartSeriesPointCount();
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

        if (plot is not null)
        {
            return plot.Series.Count > 0 ? ReadSceneChartPointExplosions(plot.Series[0], workbook) : new Dictionary<int, double>();
        }

        return ReadChartPointExplosions(chartElement);
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
                strokes[point.Index] = ToChartSeriesStroke(point.Line, null) ?? default;
            }
        }

        return strokes;
    }

    private static ChartSeriesFill? ToChartSeriesFill(PptxSceneFillStyle fill, PptxScenePatternFill patternFill)
    {
        if (fill.HasFill)
        {
            return new ChartSeriesFill(fill.Color, fill.Alpha, null, null);
        }

        return patternFill.HasPattern
            ? new ChartSeriesFill(patternFill.Foreground, patternFill.Alpha, patternFill.Preset, patternFill.Background)
            : null;
    }

    private static ChartSeriesStroke? ToChartSeriesStroke(PptxSceneLineStyle line, double? inheritedWidth)
    {
        return line.HasLine
            ? new ChartSeriesStroke(line.Color, line.Alpha, line.WidthSpecified ? line.Width : inheritedWidth ?? line.Width, line.DashPattern, line.Cap, line.Join, line.Compound)
            : null;
    }

    private static ChartShapeStyle ToChartShapeStyle(PptxSceneChartShapeStyle style)
    {
        return new ChartShapeStyle(
            style.NoFill ? null : ToChartSeriesFill(style.Fill, style.PatternFill),
            style.NoFill || style.GradientFill is null ? null : ToGradientFill(style.GradientFill),
            ToChartSeriesStroke(style.Line, null),
            style.Glow,
            style.OuterShadow);
    }

    private static IReadOnlyDictionary<int, double> ReadChartPointExplosions(XElement chartElement)
    {
        var explosions = new Dictionary<int, double>();
        XElement? series = chartElement.Element(ChartNamespace + "ser");
        if (series is null)
        {
            return explosions;
        }

        (double? seriesExplosion, _) = PptxSceneBuilder.ReadChartElementDoubleWithValue(series, "explosion");
        if (seriesExplosion is { } seriesExplosionValue)
        {
            double fraction = Math.Clamp(seriesExplosionValue / 100d, 0d, 1d);
            int pointCount = PptxSceneBuilder
                .ReadChartNumberPoints(
                    series
                        .Elements(ChartNamespace + "val")
                        .Descendants(ChartNamespace + "pt"),
                    requireNonNegativeIndex: true)
                .Select(point => point.HasParsedIndex ? point.Index : -1)
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
            if (!PptxSceneBuilder.TryReadChartNonNegativeIndex(point, out int index, out _))
            {
                continue;
            }

            (double? explosion, _) = PptxSceneBuilder.ReadChartElementDoubleWithValue(point, "explosion");
            if (explosion is { } explosionValue)
            {
                explosions[index] = Math.Clamp(explosionValue / 100d, 0d, 1d);
            }
        }

        return explosions;
    }

    private static IReadOnlyList<ChartIndexedScatterSeries> ReadScatterSeriesVectors(XElement chartElement, bool readBubbleSize, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        var series = new List<ChartIndexedScatterSeries>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            ChartIndexedNumberVector xValues = ReadChartNumberVector(element.Element(ChartNamespace + "xVal"), workbook, plotVisibleOnly);
            ChartIndexedNumberVector yValues = ReadChartNumberVector(element.Element(ChartNamespace + "yVal"), workbook, plotVisibleOnly);
            ChartIndexedNumberVector bubbleSizes = readBubbleSize
                ? ReadChartNumberVector(element.Element(ChartNamespace + "bubbleSize"), workbook, plotVisibleOnly)
                : default;
            if (xValues.DensePoints().Count == 0 && yValues.DensePoints().Count == 0)
            {
                continue;
            }

            series.Add(new ChartIndexedScatterSeries(xValues, yValues, bubbleSizes, readBubbleSize));
        }

        return series;
    }
}
