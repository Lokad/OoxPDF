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
    // Trendlines ride on series but are never painted; warn once per chart so the
    // omission is explicit instead of silent (unsupported content must diagnose).
    private static void EmitUnsupportedChartTrendlineDiagnostics(
        XDocument chartXml,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        string? chartPartName,
        int slideIndex)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        bool hasTrendline = chartXml
            .Descendants(ChartNamespace + "ser")
            .Any(series => series.Element(ChartNamespace + "trendline") is not null);
        if (hasTrendline)
        {
            EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART_TRENDLINE", OoxPdfSeverity.Warning, "Chart trendlines were detected and are not rendered.", chartPartName, slideIndex, "Ignored");
        }
    }

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

        // D01: the scene chart is always present below the missing-part return (both
        // derive from the same scene node payload), so render paths below resolve chart
        // options from the scene model. The guard documents the invariant for flow
        // analysis; the pairing proof test guards the builder side.
        if (sceneChart is null)
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart frame was missing from the scene model and was ignored.", chartPartName, context.SlideNumber, "Ignored");
            return;
        }

        // one bounded workbook model per embedded part, shared by every chart
        // frame of this conversion through the render context cache. The per-frame range
        // memo on the shared model is cleared below so nothing accumulates across frames.
        ChartWorkbookData? chartWorkbook = GetOrCreateChartWorkbook(context.WorkbookCache, sceneChart?.ExternalData ?? default, context.CancellationToken);
        context.CancellationToken.ThrowIfCancellationRequested();

        try
        {
            RenderChartFrameWithWorkbook(context, graphics, fonts, bounds.Value, chartPartName, resolvedChartXml, resolvedChartPalette, sceneChart, chartWorkbook, linkAnnotations, reportedHyperlinkIds);
        }
        finally
        {
            chartWorkbook?.ClearRangeMemo();
        }
    }

    // one bounded workbook model per embedded part, shared by every chart
    // frame of this conversion. The cache holds boxed models (the workbook type stays
    // private to the renderer); the per-conversion instance rides on the render context
    // (like the image cache) and dies with the conversion. Each model is already capped
    // (workbook totals, range unions) and the key space is bounded by package entries.
    private static ChartWorkbookData? GetOrCreateChartWorkbook(
        Dictionary<string, ChartWorkbookData?>? cache,
        PptxSceneChartExternalData externalData,
        CancellationToken cancellationToken)
    {
        if (!externalData.IsDefined || externalData.Resource is null)
        {
            return null;
        }

        // A null cache (inspection contexts) disables sharing without changing behavior.
        if (cache is null)
        {
            return ReadEmbeddedChartWorkbookData(externalData, cancellationToken);
        }

        string key = externalData.TargetPartName ?? externalData.Resource.PartName;
        if (cache.TryGetValue(key, out ChartWorkbookData? boxed))
        {
            return boxed;
        }

        ChartWorkbookData? model = ReadEmbeddedChartWorkbookData(externalData, cancellationToken);
        cache[key] = model;
        return model;
    }

    private static void RenderChartFrameWithWorkbook(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        ShapeBounds bounds,
        string chartPartName,
        XDocument resolvedChartXml,
        IReadOnlyList<RgbColor>? resolvedChartPalette,
        PptxSceneChart? sceneChart,
        ChartWorkbookData? chartWorkbook,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        PptxColorMap chartColorMap = sceneChart?.ColorMap ?? context.SlideColorMap;

        if (TryRenderChart(graphics, context.Document, context.Theme, chartColorMap, resolvedChartPalette, bounds, resolvedChartXml, sceneChart, chartWorkbook, fonts, context.FontResolver, context, linkAnnotations, reportedHyperlinkIds))
        {
            EmitUnrenderedDefaultChartAxisTitleDiagnostics(resolvedChartXml, sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
            EmitUnsupportedChartNumberFormatDiagnostics(resolvedChartXml, sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
            EmitUnsupportedChartTrendlineDiagnostics(resolvedChartXml, context.DiagnosticSink, chartPartName, context.SlideNumber);
            RenderManualChartAxisTitles(context.Document, context.Theme, chartColorMap, graphics, bounds, resolvedChartXml, sceneChart, context.FontResolver, context.DiagnosticSink, chartPartName, context.SlideNumber, emitDefaultLayoutDiagnostics: false, fonts, context, linkAnnotations, reportedHyperlinkIds);
            RenderChartTitle(context.Document, context.Theme, chartColorMap, graphics, bounds, resolvedChartXml, sceneChart, chartWorkbook, ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, resolvedChartXml), context.FontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds, context.DiagnosticSink);
            return;
        }

        // D01: the XML-only retry used to live here (workbook present without a scene
        // chart), but that state is unconstructible - a non-null workbook requires defined
        // external data, which requires a non-null scene chart. The branch is deleted
        // rather than kept as a second interpretation path.
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
            PptxSceneChartScatterStyle.Line or PptxSceneChartScatterStyle.LineMarker or PptxSceneChartScatterStyle.Smooth or PptxSceneChartScatterStyle.SmoothMarker => true,
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

    // D01: scene-authoritative. Render paths always carry a scene chart below the
    // frame early-return (a missing chart part returns before any option reads), so
    // the XML re-parse arm is deleted; the agreement harness pins the unified reading.
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
        // R15: filter on the sparse presence check first so valueless series never
        // materialize (or charge) dense arrays.
        return series
            .Where(vector => vector.HasAnyValue())
            .Select(vector => new ChartRadarSeries(vector.DensePoints(), vector))
            .ToArray();
    }

    private static int CountRenderableSeries(IEnumerable<ChartIndexedNumberVector> series)
    {
        return series.Count(vector => vector.HasAnyValue());
    }

    // R15: resolve per-series dense lengths sparsely so count-only passes
    // (tick edges) never materialize (or charge) dense arrays.
    private static int MaxDensePointCount(IEnumerable<ChartIndexedNumberVector> series)
    {
        int max = 0;
        foreach (ChartIndexedNumberVector vector in series)
        {
            max = Math.Max(max, vector.DensePointCount());
        }

        return max;
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
                .FirstOrDefault(vector => vector.Points.Count != 0 || vector.PointCount is not null || vector.HasAnyDenseSlot());
        }

        return ReadChartCategoryLabelVector(chartElement, workbook, plotVisibleOnly);
    }

    // R15: per-frame shared dense category labels behind one call shape. Workbook-less
    // (literal-only) charts densify directly; workbook-backed charts share one dense
    // array across the reserve, emission, and legend passes of a frame.
    private static IReadOnlyList<ChartIndexedTextPoint?> ReadSharedCategoryLabels(
        PptxSceneChartPlot? plot,
        XElement chartElement,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        if (workbook is null)
        {
            return ReadSceneOrXmlCategoryLabelVector(plot, chartElement, workbook, plotVisibleOnly).DensePoints();
        }

        object? source = plot is not null ? plot : null;
        return workbook.GetOrAddCategoryLabels(
            source,
            chartElement,
            plotVisibleOnly,
            () => ReadSceneOrXmlCategoryLabelVector(plot, chartElement, workbook, plotVisibleOnly).DensePoints());
    }

    // R15: per-frame shared series names behind one call shape, mirroring shared
    // category labels. Workbook-less charts read directly.
    private static IReadOnlyList<ChartSeriesNameRecord> ReadSharedChartSeriesNames(
        PptxSceneChartPlot? plot,
        XElement chartElement,
        ChartWorkbookData? workbook)
    {
        if (workbook is null)
        {
            return ReadSceneOrXmlChartSeriesNameRecords(plot, chartElement, workbook);
        }

        object? source = plot is not null ? plot : null;
        return workbook.GetOrAddSeriesNames(
            source,
            chartElement,
            () => ReadSceneOrXmlChartSeriesNameRecords(plot, chartElement, workbook));
    }

    // R15: per-frame shared series vectors behind one call shape, mirroring shared
    // labels and names. Workbook-less charts rebuild directly.
    private static IReadOnlyList<ChartIndexedNumberVector> ReadSharedChartSeriesVectors(
        PptxSceneChartPlot? plot,
        XElement chartElement,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        if (workbook is null)
        {
            return ReadSceneOrXmlChartSeriesVectors(plot, chartElement, workbook, plotVisibleOnly);
        }

        object? source = plot is not null ? plot : null;
        return workbook.GetOrAddSeriesVectors(
            source,
            chartElement,
            plotVisibleOnly,
            () => ReadSceneOrXmlChartSeriesVectors(plot, chartElement, workbook, plotVisibleOnly));
    }

    // R15: per-frame shared bar value extents behind one call shape, mirroring
    // shared series vectors. Extents densify every series; reserve, emission, and
    // crossing passes within one frame share one result. Workbook-less charts
    // compute directly.
    private static ChartValueExtents GetSharedBarChartValueExtents(
        PptxSceneChartPlot? plot,
        XElement chartElement,
        PptxSceneChartGrouping grouping,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        if (workbook is null)
        {
            return GetBarChartValueExtents(ReadSceneOrXmlChartSeriesVectors(plot, chartElement, workbook, plotVisibleOnly), grouping);
        }

        object? source = plot is not null ? plot : null;
        return workbook.GetOrAddValueExtents(
            source,
            chartElement,
            grouping,
            plotVisibleOnly,
            () => GetBarChartValueExtents(ReadSharedChartSeriesVectors(plot, chartElement, workbook, plotVisibleOnly), grouping));
    }

    // R15: per-frame shared line value extents behind one call shape, mirroring
    // shared bar extents. Workbook-less charts compute directly.
    private static ChartValueExtents GetSharedLineChartValueExtents(
        PptxSceneChartPlot? plot,
        XElement chartElement,
        bool stacked,
        bool percentStacked,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        if (workbook is null)
        {
            return GetLineChartValueExtents(ReadSceneOrXmlChartSeriesVectors(plot, chartElement, workbook, plotVisibleOnly), stacked, percentStacked);
        }

        object? source = plot is not null ? plot : null;
        return workbook.GetOrAddLineValueExtents(
            source,
            chartElement,
            stacked,
            percentStacked,
            plotVisibleOnly,
            () => GetLineChartValueExtents(ReadSharedChartSeriesVectors(plot, chartElement, workbook, plotVisibleOnly), stacked, percentStacked));
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

        // workbook lookups below ran a visibility filter plus linear scan per
        // point (quadratic in series length). Index once per vector instead; first-wins
        // insertion matches the linear scan exactly (range indices are unique).
        Dictionary<int, ChartIndexedNumberPoint> xWorkbook = BuildWorkbookPointIndex(series.XValues);
        Dictionary<int, ChartIndexedNumberPoint> yWorkbook = BuildWorkbookPointIndex(series.YValues);
        Dictionary<int, ChartIndexedNumberPoint> bubbleWorkbook = series.ReadBubbleSize ? BuildWorkbookPointIndex(series.BubbleSizes) : new Dictionary<int, ChartIndexedNumberPoint>();
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
                xWorkbook.TryGetValue(xPoint.Value.Index, out ChartIndexedNumberPoint xWorkbookPoint) ? xWorkbookPoint : null,
                yWorkbook.TryGetValue(yPoint.Value.Index, out ChartIndexedNumberPoint yWorkbookPoint) ? yWorkbookPoint : null,
                bubbleSizePoint is { } point && bubbleWorkbook.TryGetValue(point.Index, out ChartIndexedNumberPoint bubbleWorkbookPoint) ? bubbleWorkbookPoint : null,
                series.YValues.FormatCode,
                series.BubbleSizes.FormatCode));
        }

        return new ScatterSeries(points, series);
    }

    private static Dictionary<int, ChartIndexedNumberPoint> BuildWorkbookPointIndex(ChartIndexedNumberVector vector)
    {
        var index = new Dictionary<int, ChartIndexedNumberPoint>();
        foreach (ChartIndexedNumberPoint point in vector.WorkbookPointsForPlotVisibility(vector.PlotVisibleOnly))
        {
            if (!index.ContainsKey(point.Index))
            {
                index.Add(point.Index, point);
            }
        }

        return index;
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
        if (points.Count == 0)
        {
            return null;
        }

        long max = long.MinValue;
        foreach (ChartIndexedNumberPoint point in points)
        {
            max = Math.Max(max, (long)point.Index);
        }

        long inferred;
        try
        {
            inferred = checked(max + 1);
        }
        catch (OverflowException ex)
        {
            throw new OoxPdfLimitExceededException("Chart point index overflows.", ex);
        }

        if (inferred < 0 || inferred > int.MaxValue)
        {
            throw new OoxPdfLimitExceededException("Chart point index is out of range.");
        }

        return (int)inferred;
    }

    private static int? InferPointCount(IReadOnlyList<ChartIndexedTextPoint> points)
    {
        if (points.Count == 0)
        {
            return null;
        }

        long max = long.MinValue;
        foreach (ChartIndexedTextPoint point in points)
        {
            max = Math.Max(max, (long)point.Index);
        }

        long inferred;
        try
        {
            inferred = checked(max + 1);
        }
        catch (OverflowException ex)
        {
            throw new OoxPdfLimitExceededException("Chart point index overflows.", ex);
        }

        if (inferred < 0 || inferred > int.MaxValue)
        {
            throw new OoxPdfLimitExceededException("Chart point index is out of range.");
        }

        return (int)inferred;
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
            if (!xValues.HasAnyDenseSlot() && !yValues.HasAnyDenseSlot())
            {
                continue;
            }

            series.Add(new ChartIndexedScatterSeries(xValues, yValues, bubbleSizes, readBubbleSize));
        }

        return series;
    }
}
