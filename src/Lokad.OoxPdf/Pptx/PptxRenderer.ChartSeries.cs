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

    private static async ValueTask RenderChartFrameAsync(
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
        await RenderChartFrameAsync(context, graphics, fonts, bounds, node.Chart, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
    }

    private static async ValueTask RenderChartFrameAsync(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        ShapeBounds? bounds,
        PptxSceneChart? sceneChart,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        context.CancellationToken.ThrowIfCancellationRequested();
        string? chartPartName = sceneChart?.TargetPartName;

        if (sceneChart is null || bounds is null || chartPartName is null)
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart frame could not be resolved and was ignored.", context.SceneSlide.PartName, context.SlideNumber, "Ignored");
            return;
        }

        XDocument? resolvedChartXml = sceneChart.ChartXml;
        if (resolvedChartXml is null)
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Chart part was missing from the scene model and was ignored.", chartPartName, context.SlideNumber, "Ignored");
            return;
        }

        // one bounded workbook model per embedded part, shared by every chart
        // frame of this conversion through the render context cache. The per-frame range
        // memo on the shared model is cleared below so nothing accumulates across frames.
        ChartWorkbookData? chartWorkbook = GetOrCreateChartWorkbook(context.WorkbookCache, sceneChart.ExternalData, context.CancellationToken);
        context.CancellationToken.ThrowIfCancellationRequested();

        try
        {
            if (context.PageConstruction is { } construction && chartWorkbook is not null)
            {
                long capacity = checked(context.PageConstructionBaseBytes + chartWorkbook.EstimateRangeConstruction(sceneChart, context.CancellationToken));
                // Range memos/projections coexist within a frame, then clear below.
                // Sequential frames reuse the page's high-water construction window.
                if (capacity > construction.Bytes) OoxMemory.Resize(construction, capacity);
                context.CancellationToken.ThrowIfCancellationRequested();
                // Axis expansion coexists with this frame's range memos.
                context = context with { PageConstructionBaseBytes = capacity };
            }
            await RenderChartFrameWithWorkbookAsync(context, graphics, fonts, bounds.Value, chartPartName, resolvedChartXml, sceneChart, chartWorkbook, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
        }
        finally
        {
            chartWorkbook?.ClearRangeMemo();
        }
    }

    // One bounded model per embedded workbook, shared by chart frames for this
    // conversion. A null cache disables sharing.
    internal static ChartWorkbookData? GetOrCreateChartWorkbook(
        Dictionary<string, ChartWorkbookData?>? cache,
        PptxSceneChartExternalData externalData,
        CancellationToken cancellationToken)
    {
        if (externalData.Resource is not { } part)
        {
            return null;
        }

        if (cache is null)
        {
            return ReadEmbeddedChartWorkbookData(part, cancellationToken);
        }

        if (cache.TryGetValue(part.Name, out ChartWorkbookData? cached))
        {
            return cached;
        }

        ChartWorkbookData? model = ReadEmbeddedChartWorkbookData(part, cancellationToken);
        cache.Add(part.Name, model);
        return model;
    }

    private static async ValueTask RenderChartFrameWithWorkbookAsync(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        ShapeBounds bounds,
        string chartPartName,
        XDocument resolvedChartXml,
        PptxSceneChart sceneChart,
        ChartWorkbookData? chartWorkbook,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        PptxColorMap chartColorMap = sceneChart.ColorMap;

        if ((await TryRenderChartAsync(graphics, context.Document, context.Theme, chartColorMap, bounds, sceneChart, chartWorkbook, fonts, context.FontResolver, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false)))
        {
            EmitUnrenderedDefaultChartAxisTitleDiagnostics(sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
            EmitUnsupportedChartNumberFormatDiagnostics(resolvedChartXml, sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
            EmitUnsupportedChartTrendlineDiagnostics(resolvedChartXml, context.DiagnosticSink, chartPartName, context.SlideNumber);
            await RenderManualChartAxisTitlesAsync(context.Document, context.Theme, graphics, bounds, sceneChart, context.FontResolver, context.DiagnosticSink, fonts, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
            await RenderChartTitleAsync(context.Document, context.Theme, graphics, bounds, sceneChart, context.FontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds, context.DiagnosticSink).ConfigureAwait(false);
            return;
        }

        bool HasSupportedSceneChartWithoutRenderableCachedValues()
        {
            return sceneChart.Plots.Any(plot =>
                IsSupportedNativeChartPlot(plot.PlotKind) &&
                plot.Series.Count != 0 &&
                !plot.Series.Any(HasRenderableCachedValues));
        }

        if (HasSupportedSceneChartWithoutRenderableCachedValues())
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_CHART_MISSING_CACHED_DATA", OoxPdfSeverity.Warning, "Supported chart references formula-only data without chart-side cached numeric values. Embedded workbook values are preserved as provenance but are not used as active rendering data.", chartPartName, context.SlideNumber, "Ignored");
            return;
        }

        EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Only bar, line, area, scatter, bubble, radar, pie, and doughnut charts with cached numeric values are currently supported by the native chart renderer.", chartPartName, context.SlideNumber, "Ignored");
        }


    internal static PptxSceneChartPlot? ReadFirstSceneChartPlot(PptxSceneChart? chart, PptxSceneChartPlotKind kind) =>
        chart?.Plots.FirstOrDefault(plot => plot.PlotKind == kind);

    private static bool IsSupportedNativeChartPlot(PptxSceneChartPlotKind kind) => kind is
        PptxSceneChartPlotKind.Area or PptxSceneChartPlotKind.Bar or PptxSceneChartPlotKind.Bubble or
        PptxSceneChartPlotKind.Doughnut or PptxSceneChartPlotKind.Line or PptxSceneChartPlotKind.Pie or
        PptxSceneChartPlotKind.Radar or PptxSceneChartPlotKind.Scatter;

    private static bool HasRenderableCachedValues(PptxSceneChartSeries series) =>
        series.ValuePoints.Any(point => point.Value is not null) ||
        series.XValuePoints.Any(point => point.Value is not null) && series.YValuePoints.Any(point => point.Value is not null);

    internal static IReadOnlyList<PptxSceneChartPlot> ReadSceneChartPlots(PptxSceneChart? chart, PptxSceneChartPlotKind kind)
    {
        return chart?.Plots
            .Where(plot => plot.PlotKind == kind)
            .ToArray() ?? [];
    }

    private static IReadOnlyList<PptxSceneChartAxis> ReadChartAxesForPlot(
        PptxSceneChart chart,
        PptxSceneChartPlot? plot,
        Func<PptxSceneChartAxis, bool> matches)
    {
        if (plot is null || plot.AxisIds.Count == 0)
            return chart.Axes.Where(matches).ToArray();

        // Preserve declared order and duplicates. Missing IDs stay missing;
        // neither source XML nor another plot supplies a replacement axis.
        return plot.AxisIds.Select(id => chart.Axes.FirstOrDefault(axis =>
                PptxChartAxisIdComparer.Instance.Equals(axis.Id, id) && matches(axis)))
            .OfType<PptxSceneChartAxis>().ToArray();
    }

    private static IReadOnlyList<ChartAxisSource> ReadChartValueAxesForPlot(
        PptxSceneChart chart, PptxSceneChartPlot? plot) =>
        ReadChartAxesForPlot(chart, plot, static axis => axis.AxisKind == PptxSceneChartAxisKind.Value)
            .Select(axis => ReadChartAxisSource(chart, axis)).ToArray();

    private static ChartAxisSource ReadChartCategoryAxisForPlot(
        PptxSceneChart chart, PptxSceneChartPlot? plot) =>
        ReadChartAxisSource(chart, ReadChartAxesForPlot(chart, plot,
            static axis => axis.AxisKind is PptxSceneChartAxisKind.Category or PptxSceneChartAxisKind.Date).FirstOrDefault());

    private static ChartAxisSource ReadSceneSecondaryRightValueAxis(PptxSceneChart chart) =>
        ReadChartAxisSource(chart, chart.Axes.FirstOrDefault(axis =>
            axis.AxisKind == PptxSceneChartAxisKind.Value && axis.IsDeleted != true &&
            axis.PositionKind == PptxSceneChartAxisPosition.Right));

    private static ChartAxisSource ReadSceneSecondaryValueAxisForChart(PptxSceneChart chart, ChartAxisSource primary)
    {
        string? primaryId = primary.SceneAxis?.Id;
        PptxSceneChartAxis? selected = !string.IsNullOrWhiteSpace(primaryId)
            ? chart.Axes.FirstOrDefault(axis => axis.AxisKind == PptxSceneChartAxisKind.Value &&
                !PptxChartAxisIdComparer.Instance.Equals(axis.Id, primaryId) && axis.IsDeleted != true)
            : chart.Axes.FirstOrDefault(axis => axis.AxisKind == PptxSceneChartAxisKind.Value &&
                axis.IsDeleted != true && axis.PositionKind == PptxSceneChartAxisPosition.Right);
        return ReadChartAxisSource(chart, selected);
    }

    private static ChartAxisSource ReadChartAxisSource(PptxSceneChart chart, PptxSceneChartAxis? axis)
    {
        return axis is null ? default : new(axis, chart.Options.Date1904 == true);
    }

    private static PptxSceneChartGrouping ReadSceneChartGrouping(PptxSceneChartPlot plot, PptxSceneChartGrouping defaultGrouping) =>
        plot.GroupingKind == PptxSceneChartGrouping.Unknown ? defaultGrouping : plot.GroupingKind;

    private static PptxSceneChartBarDirection ReadSceneChartBarDirection(PptxSceneChartPlot plot) =>
        plot.BarDirectionKind == PptxSceneChartBarDirection.Unknown ? PptxSceneChartBarDirection.Column : plot.BarDirectionKind;

    private static PptxSceneChartRadarStyle ReadSceneChartRadarStyle(PptxSceneChartPlot plot) =>
        plot.RadarStyleKind == PptxSceneChartRadarStyle.Unknown ? PptxSceneChartRadarStyle.Standard : plot.RadarStyleKind;

    private static PptxSceneChartDisplayBlanksAs ResolveChartDisplayBlanksAs(PptxSceneChart chart) =>
        chart.Options.DisplayBlanksAsKind == PptxSceneChartDisplayBlanksAs.Unknown
            ? PptxSceneChartDisplayBlanksAs.Gap
            : chart.Options.DisplayBlanksAsKind;

    private static double?[][] BuildRadarSeries(IEnumerable<ChartIndexedNumberVector> series)
    {
        // Keep style indices aligned without densifying empty series.
        return series
            .Select(vector => vector.HasAnyValue() ? vector.DenseValues() : [])
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

    // RV20: compact value-only densification for geometry/extent consumers.
    // Geometry retains one nullable double per slot; provenance consumers
    // use the sparse source index below.
    private static IReadOnlyList<IReadOnlyList<double?>> DensifyChartValueSeries(IEnumerable<ChartIndexedNumberVector> series)
    {
        return series
            .Select(vector => vector.DenseValues())
            .ToArray();
    }

    // RV20: last-wins dense-slot provenance index over the sparse source,
    // sharing compact-value resolution, range and validity,
    // so provenance consumers avoid materializing the dense array.
    private static Dictionary<int, ChartIndexedNumberPoint> BuildDensePointIndex(ChartIndexedNumberVector vector)
    {
        int pointCount = vector.DensePointCount();
        var index = new Dictionary<int, ChartIndexedNumberPoint>();
        foreach (ChartIndexedNumberPoint point in vector.Points ?? [])
        {
            if (point.Index >= 0 && point.Index < pointCount && point.Value is not null)
            {
                index[point.Index] = point;
            }
        }

        return index;
    }

    private static IReadOnlyList<ChartIndexedNumberVector> ReadSceneChartSeriesVectors(PptxSceneChartPlot plot, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        return plot.Series
            .Select(series => BuildChartIndexedNumberVector(
                series.ValuePoints,
                series.ValuePointCount,
                series.ValueFormatCode,
                series.DataSources.Values,
                workbook,
                plotVisibleOnly))
            .ToArray();
    }

    private static IReadOnlyList<ScatterSeries> ReadSceneScatterSeries(PptxSceneChartPlot plot, bool readBubbleSize, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        ScatterSeries[] series = ReadSceneScatterSeriesVectors(plot, readBubbleSize, workbook, plotVisibleOnly)
            .Select(BuildScatterSeries)
            .ToArray();
        // Styles, labels and smoothing use the same source series positions.
        return series.Any(item => item.Points.Count != 0) ? series : [];
    }

    private static IReadOnlyList<ChartIndexedScatterSeries> ReadSceneScatterSeriesVectors(PptxSceneChartPlot plot, bool readBubbleSize, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        return plot.Series
            .Select(item => new ChartIndexedScatterSeries(
                BuildChartIndexedNumberVector(
                    item.XValuePoints,
                    item.XValuePointCount,
                    item.XValueFormatCode,
                    item.DataSources.XValues,
                    workbook,
                    plotVisibleOnly),
                BuildChartIndexedNumberVector(
                    item.YValuePoints,
                    item.YValuePointCount,
                    item.YValueFormatCode,
                    item.DataSources.YValues,
                    workbook,
                    plotVisibleOnly),
                BuildChartIndexedNumberVector(
                    item.BubbleSizePoints,
                    item.BubbleSizePointCount,
                    item.BubbleSizeFormatCode,
                    item.DataSources.BubbleSizes,
                    workbook,
                    plotVisibleOnly),
                readBubbleSize))
            .ToArray();
    }

    private static ChartIndexedTextVector ReadSceneCategoryLabelVector(PptxSceneChartPlot plot, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        return plot.Series
            .Select(series => BuildChartIndexedTextVector(
                series.CategoryPoints,
                series.CategoryPointCount,
                series.CategoryLevels,
                series.DataSources.Categories,
                workbook,
                plotVisibleOnly))
            .FirstOrDefault(vector => vector.Points.Count != 0 || vector.PointCount is not null);
    }

    // R15: per-frame shared dense category labels behind one call shape. Workbook-less
    // (literal-only) charts densify directly; workbook-backed charts share one dense
    // array across the reserve, emission, and legend passes of a frame.
    private static IReadOnlyList<ChartIndexedTextPoint?> ReadSharedCategoryLabels(
        PptxSceneChartPlot plot,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        if (workbook is null)
        {
            return ReadSceneCategoryLabelVector(plot, workbook, plotVisibleOnly).DensePoints();
        }

        return workbook.GetOrAddCategoryLabels(
            plot,
            plotVisibleOnly,
            () => ReadSceneCategoryLabelVector(plot, workbook, plotVisibleOnly).DensePoints());
    }

    // R15: per-frame shared series names behind one call shape, mirroring shared
    // category labels. Workbook-less charts read directly.
    internal static IReadOnlyList<string> ReadSharedChartSeriesNames(
        PptxSceneChartPlot plot,
        ChartWorkbookData? workbook)
    {
        if (workbook is null)
        {
            return ReadSceneChartSeriesNames(plot, workbook);
        }

        return workbook.GetOrAddSeriesNames(
            plot,
            () => ReadSceneChartSeriesNames(plot, workbook));
    }

    // R15: per-frame shared series vectors behind one call shape, mirroring shared
    // labels and names. Workbook-less charts rebuild directly.
    private static IReadOnlyList<ChartIndexedNumberVector> ReadSharedChartSeriesVectors(
        PptxSceneChartPlot plot,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        if (workbook is null)
        {
            return ReadSceneChartSeriesVectors(plot, workbook, plotVisibleOnly);
        }

        return workbook.GetOrAddSeriesVectors(
            plot,
            plotVisibleOnly,
            () => ReadSceneChartSeriesVectors(plot, workbook, plotVisibleOnly));
    }

    // R15: per-frame shared bar value extents behind one call shape, mirroring
    // shared series vectors. Extents densify every series; reserve, emission, and
    // crossing passes within one frame share one result. Workbook-less charts
    // compute directly.
    private static ChartValueExtents GetSharedBarChartValueExtents(
        PptxSceneChartPlot plot,
        PptxSceneChartGrouping grouping,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly,
        IReadOnlyList<ChartIndexedNumberVector>? seriesVectors = null)
    {
        return workbook is null
            ? GetExtents()
            : workbook.GetOrAddValueExtents(plot, grouping, plotVisibleOnly, GetExtents);

        ChartValueExtents GetExtents() => GetBarChartValueExtents(
            seriesVectors ?? ReadSharedChartSeriesVectors(plot, workbook, plotVisibleOnly), grouping);
    }

    // R15: per-frame shared line value extents behind one call shape, mirroring
    // shared bar extents. Workbook-less charts compute directly.
    private static ChartValueExtents GetSharedLineChartValueExtents(
        PptxSceneChartPlot plot,
        bool stacked,
        bool percentStacked,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        if (workbook is null)
        {
            return GetLineChartValueExtents(ReadSceneChartSeriesVectors(plot, workbook, plotVisibleOnly), stacked, percentStacked);
        }

        return workbook.GetOrAddLineValueExtents(
            plot,
            stacked,
            percentStacked,
            plotVisibleOnly,
            () => GetLineChartValueExtents(ReadSharedChartSeriesVectors(plot, workbook, plotVisibleOnly), stacked, percentStacked));
    }

    // Called after shared-axis ownership is established. Preserve the existing
    // series/workbook memos; only the small range value crosses stages.
    private static ChartValueExtents GetSharedCartesianChartValueExtents(PptxSceneChart chart,
        PptxSceneChartPlot primaryPlot, ChartValueExtents primaryData, ChartWorkbookData? workbook,
        bool plotVisibleOnly, CancellationToken cancellationToken)
    {
        double minimum = primaryData.Min, maximum = primaryData.Max;
        foreach (var plot in chart.Plots)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (ReferenceEquals(plot, primaryPlot)) continue;
            var vectors = ReadSharedChartSeriesVectors(plot, workbook, plotVisibleOnly);
            if (CountRenderableSeries(vectors) == 0) continue;
            var data = plot.PlotKind == PptxSceneChartPlotKind.Bar
                ? GetBarChartValueExtents(vectors, plot.GroupingKind)
                : GetLineChartValueExtents(vectors, stacked: false, percentStacked: false);
            minimum = Math.Min(minimum, data.Min);
            maximum = Math.Max(maximum, data.Max);
        }
        return new(minimum, maximum);
    }

    private static ScatterSeries BuildScatterSeries(ChartIndexedScatterSeries series)
    {
        // RV20: compact value arrays pair positions; dense-slot provenance
        // comes from last-wins sparse indexes (same slots as DensePoints).
        // Unread bubble sizes resolve their count (caps parity) without
        // materializing values.
        double?[] xValues = series.XValues.DenseValues();
        double?[] yValues = series.YValues.DenseValues();
        double?[] bubbleValues = series.ReadBubbleSize ? series.BubbleSizes.DenseValues() : [];
        if (!series.ReadBubbleSize)
        {
            series.BubbleSizes.DensePointCount();
        }

        int count = Math.Max(xValues.Length, yValues.Length);
        if (count == 0)
        {
            return new ScatterSeries([], series);
        }

        Dictionary<int, ChartIndexedNumberPoint> xPoints = BuildDensePointIndex(series.XValues);
        Dictionary<int, ChartIndexedNumberPoint> yPoints = BuildDensePointIndex(series.YValues);
        Dictionary<int, ChartIndexedNumberPoint> bubblePoints = series.ReadBubbleSize ? BuildDensePointIndex(series.BubbleSizes) : new Dictionary<int, ChartIndexedNumberPoint>();

        // workbook lookups below ran a visibility filter plus linear scan per
        // point (quadratic in series length). Index once per vector instead; first-wins
        // insertion matches the linear scan exactly (range indices are unique).
        Dictionary<int, ChartIndexedNumberPoint> xWorkbook = BuildWorkbookPointIndex(series.XValues);
        Dictionary<int, ChartIndexedNumberPoint> yWorkbook = BuildWorkbookPointIndex(series.YValues);
        Dictionary<int, ChartIndexedNumberPoint> bubbleWorkbook = series.ReadBubbleSize ? BuildWorkbookPointIndex(series.BubbleSizes) : new Dictionary<int, ChartIndexedNumberPoint>();
        var points = new List<ScatterPoint>(count);
        for (int i = 0; i < count; i++)
        {
            if (i >= xValues.Length || xValues[i] is not { } xValue || i >= yValues.Length || yValues[i] is not { } yValue)
            {
                continue;
            }

            ChartIndexedNumberPoint xPoint = xPoints[i];
            ChartIndexedNumberPoint yPoint = yPoints[i];
            double? bubbleValue = series.ReadBubbleSize && i < bubbleValues.Length ? bubbleValues[i] : null;
            ChartIndexedNumberPoint? bubbleSizePoint = bubbleValue is not null && bubblePoints.TryGetValue(i, out ChartIndexedNumberPoint bubbleSlot) ? bubbleSlot : null;
            double size = bubbleValue ?? 1d;
            points.Add(new ScatterPoint(
                xValue,
                yValue,
                size,
                yPoint.Index,
                xPoint,
                yPoint,
                bubbleSizePoint,
                xWorkbook.TryGetValue(xPoint.Index, out ChartIndexedNumberPoint xWorkbookPoint) ? xWorkbookPoint : null,
                yWorkbook.TryGetValue(yPoint.Index, out ChartIndexedNumberPoint yWorkbookPoint) ? yWorkbookPoint : null,
                bubbleSizePoint is { } point && bubbleWorkbook.TryGetValue(point.Index, out ChartIndexedNumberPoint bubbleWorkbookPoint) ? bubbleWorkbookPoint : null,
                series.XValues.FormatCode,
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
        IReadOnlyList<PptxSceneChartNumberPoint> scenePoints,
        int? pointCount,
        string? formatCode,
        PptxSceneChartDataSource source,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        IReadOnlyList<ChartIndexedNumberPoint> workbookPoints = ReadWorkbookNumberPoints(workbook, source);
        IReadOnlyList<ChartIndexedNumberPoint> points = scenePoints.Select(ToChartIndexedNumberPoint).ToArray();
        pointCount ??= InferPointCount(points.Count == 0 ? null : points.Max(static point => point.Index));
        return new ChartIndexedNumberVector(points, pointCount, source.Formula, formatCode, source, workbookPoints, plotVisibleOnly);
    }

    private static ChartIndexedNumberPoint ToChartIndexedNumberPoint(PptxSceneChartNumberPoint point)
    {
        return new ChartIndexedNumberPoint(
            point.Index,
            point.Value,
            point.Text,
            point.HasValueElement,
            default);
    }

    private static ChartIndexedTextVector BuildChartIndexedTextVector(
        IReadOnlyList<PptxSceneChartStringPoint> scenePoints,
        int? pointCount,
        IReadOnlyList<IReadOnlyList<PptxSceneChartStringPoint>> categoryLevels,
        PptxSceneChartDataSource source,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly)
    {
        IReadOnlyList<ChartIndexedTextPoint> workbookPoints = ReadWorkbookTextPoints(workbook, source);
        IReadOnlyList<ChartIndexedTextPoint> points = scenePoints.Select(ToChartIndexedTextPoint).ToArray();
        IReadOnlyList<IReadOnlyList<ChartIndexedTextPoint>> levels = categoryLevels
            .Select(level => level.Select(ToChartIndexedTextPoint).ToArray())
            .ToArray();
        pointCount ??= InferPointCount(points.Count == 0 ? null : points.Max(static point => point.Index));
        return new ChartIndexedTextVector(points, pointCount, levels, source.Formula, source, workbookPoints, plotVisibleOnly);
    }

    private static ChartIndexedTextPoint ToChartIndexedTextPoint(PptxSceneChartStringPoint point)
    {
        return new ChartIndexedTextPoint(
            point.Index,
            point.Text,
            point.HasText,
            default);
    }

    private static IReadOnlyList<ChartIndexedNumberPoint> ReadWorkbookNumberPoints(ChartWorkbookData? workbook, PptxSceneChartDataSource source)
    {
        return workbook?
            .ReadRangeCells(source.Formula)
            .Select(cell => new ChartIndexedNumberPoint(
                cell.Index,
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
            .Select(cell => new ChartIndexedTextPoint(cell.Index, cell.Text, cell.HasValue && !string.IsNullOrWhiteSpace(cell.Text), cell))
            .ToArray() ?? [];
    }

    private static int? InferPointCount(int? largestIndex)
    {
        if (largestIndex is not int index)
        {
            return null;
        }

        long inferred = (long)index + 1;
        if (inferred < 0 || inferred > int.MaxValue)
        {
            throw new OoxPdfLimitExceededException("Chart point index is out of range.");
        }

        return (int)inferred;
    }

    private static IReadOnlyList<ChartSeriesFill?> ReadSceneSeriesFills(PptxSceneChartPlot plot) =>
        plot.Series.Select(series => ToChartSeriesFill(series.Fill, series.PatternFill)).ToArray();

    private static IReadOnlyList<ChartSeriesStroke?> ReadSceneSeriesStrokes(PptxSceneChartPlot plot, double? inheritedWidth) =>
        plot.Series.Select(series => ToChartSeriesStroke(series.Line, inheritedWidth)).ToArray();

    private static IReadOnlyList<ChartMarkerStyle> ReadSceneMarkerStyles(PptxSceneChartPlot plot) =>
        plot.Series.Select(series => ToChartMarkerStyle(series.Marker)).ToArray();

    private static ChartMarkerStyle ToChartMarkerStyle(PptxSceneChartMarker marker) => new(
        marker.SymbolKind, marker.Symbol, marker.SizeValue, marker.Size,
        marker.Fill.HasFill ? new ChartSeriesFill(marker.Fill.Color, marker.Fill.Alpha, null, null) : null,
        marker.Line.NoFill ? new ChartSeriesStroke(default, 0d, ChartMarkerInheritedStrokeWidth)
            : ToChartSeriesStroke(marker.Line, ChartMarkerInheritedStrokeWidth), marker.IsDefined);

    // RV04: plot-level smooth applies to series without their own smooth element on line
    // charts (COM plot-smooth-only probe smooths every series); series-explicit values win.
    // Scatter keeps the series-only legacy path (plot-level scatter smoothing unprobed).
    private static IReadOnlyList<ChartSeriesSmoothing> ReadSceneSmoothSeries(PptxSceneChartPlot plot,
        bool applyPlotFallback, bool defaultValue = false)
    {
        bool? plotSmooth = applyPlotFallback ? PptxSceneBuilder.ReadOptionalChartBoolean(plot.Source, "smooth") : null;
        // Only series-explicit markup keeps an individual line straight in a
        // smoothed plot. A plot fallback affects the value without defining it.
        return plot.Series.Select(series => new ChartSeriesSmoothing(
            series.Smooth ?? plotSmooth ?? defaultValue, series.Smooth.HasValue)).ToArray();
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> ReadSceneSeriesPointFills(PptxSceneChartPlot plot) =>
        plot.Series.Select(ReadSceneChartPointFills).ToArray();

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> ReadSceneSeriesPointStrokes(PptxSceneChartPlot plot) =>
        plot.Series.Select(ReadSceneChartPointStrokes).ToArray();

    private static IReadOnlyDictionary<int, double> ReadSceneChartPointExplosions(PptxSceneChartPlot plot, ChartWorkbookData? workbook)
    {
        var explosions = new Dictionary<int, double>();
        if (plot.Series.FirstOrDefault() is not { } series) return explosions;
        if (series.Explosion is { } seriesExplosion)
        {
            double fraction = Math.Clamp(seriesExplosion / 100d, 0d, 1d);
            int ResolveSceneChartSeriesPointCount()
            {
                ChartIndexedNumberVector values = BuildChartIndexedNumberVector(
                    series.ValuePoints,
                    series.ValuePointCount,
                    series.ValueFormatCode,
                    series.DataSources.Values,
                    workbook,
                    plotVisibleOnly: true);
                ChartIndexedTextVector categories = BuildChartIndexedTextVector(
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
            ? new ChartSeriesFill(patternFill.Foreground, 1d, patternFill.Preset, patternFill.Background)
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

}
