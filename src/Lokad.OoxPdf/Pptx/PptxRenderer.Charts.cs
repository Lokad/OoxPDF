using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
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
    private static readonly XNamespace SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static void RenderChartFrame(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        PptxSceneNode node,
        GroupTransform transform)
    {
        ShapeBounds? bounds = node.Bounds is { } rawBounds
            ? transform.Apply(ToShapeBounds(rawBounds))
            : null;
        RenderChartFrame(context, graphics, fonts, bounds, node.Chart);
    }

    private static void RenderChartFrame(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        ShapeBounds? bounds,
        PptxSceneChart? chart)
    {
        RenderChartFrame(
            context,
            graphics,
            fonts,
            bounds,
            chart?.TargetPartName,
            chart?.ChartXml,
            chart?.PaletteColors,
            chart);
    }

    private static void RenderChartFrame(
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        List<PdfFontResource> fonts,
        ShapeBounds? bounds,
        string? targetPartName,
        XDocument? chartXml,
        IReadOnlyList<RgbColor>? chartPalette,
        PptxSceneChart? sceneChart)
    {
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

        ChartWorkbookData? chartWorkbook = ReadEmbeddedChartWorkbookData(sceneChart?.ExternalData ?? default);

        PptxColorMap chartColorMap = sceneChart?.ColorMap ?? context.SlideColorMap;
        if (TryRenderChart(graphics, context.Document, context.Theme, chartColorMap, resolvedChartPalette, bounds.Value, resolvedChartXml, sceneChart, chartWorkbook, fonts, context.FontResolver))
        {
            EmitUnrenderedDefaultChartAxisTitleDiagnostics(resolvedChartXml, sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
            fonts.AddRange(RenderManualChartAxisTitles(context.Document, context.Theme, chartColorMap, graphics, bounds.Value, resolvedChartXml, sceneChart, context.FontResolver, context.DiagnosticSink, chartPartName, context.SlideNumber, emitDefaultLayoutDiagnostics: false));
            fonts.AddRange(RenderChartTitle(context.Document, context.Theme, chartColorMap, graphics, bounds.Value, resolvedChartXml, sceneChart, chartWorkbook, ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, resolvedChartXml), context.FontResolver));
            return;
        }

        if (chartWorkbook is not null && sceneChart is null)
        {
            HydrateChartReferenceCaches(chartWorkbook, resolvedChartXml);
            if (TryRenderChart(graphics, context.Document, context.Theme, chartColorMap, resolvedChartPalette, bounds.Value, resolvedChartXml, sceneChart, workbook: null, fonts, context.FontResolver))
            {
                EmitUnrenderedDefaultChartAxisTitleDiagnostics(resolvedChartXml, sceneChart, context.DiagnosticSink, chartPartName, context.SlideNumber);
                fonts.AddRange(RenderManualChartAxisTitles(context.Document, context.Theme, chartColorMap, graphics, bounds.Value, resolvedChartXml, sceneChart, context.FontResolver, context.DiagnosticSink, chartPartName, context.SlideNumber, emitDefaultLayoutDiagnostics: false));
                fonts.AddRange(RenderChartTitle(context.Document, context.Theme, chartColorMap, graphics, bounds.Value, resolvedChartXml, sceneChart, workbook: null, ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, resolvedChartXml), context.FontResolver));
                return;
            }
        }

        if (HasSupportedSceneChartWithoutRenderableCachedValues(sceneChart))
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_CHART_MISSING_CACHED_DATA", OoxPdfSeverity.Warning, "Supported chart references formula-only data without chart-side cached numeric values. Embedded workbook values are preserved as provenance but are not used as active rendering data.", chartPartName, context.SlideNumber, "Ignored");
            return;
        }

        EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning, "Only bar, line, area, scatter, bubble, radar, pie, and doughnut charts with cached numeric values are currently supported by the native chart renderer.", chartPartName, context.SlideNumber, "Ignored");
    }

    private static bool HasSupportedSceneChartWithoutRenderableCachedValues(PptxSceneChart? chart)
    {
        return chart?.Plots.Any(plot =>
            IsSupportedNativeChartPlot(plot.PlotKind) &&
            plot.Series.Count != 0 &&
            !plot.Series.Any(HasRenderableCachedValues)) == true;
    }

    private static bool IsSupportedNativeChartPlot(PptxSceneChartPlotKind kind)
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

    private static bool HasRenderableCachedValues(PptxSceneChartSeries series)
    {
        return series.Values.Count != 0 ||
            series.ValuePoints.Any(point => point.Value is not null) ||
            (series.XValues.Count != 0 && series.YValues.Count != 0) ||
            (series.XValuePoints.Any(point => point.Value is not null) &&
                series.YValuePoints.Any(point => point.Value is not null));
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

    private static XElement? ReadSceneOrXmlFirstChartPlotElement(PptxSceneChart? sceneChart, XDocument chartXml, PptxSceneChartPlotKind kind)
    {
        if (sceneChart is not null)
        {
            return ReadSceneChartPlot(sceneChart, kind)?.Source;
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

    private static IReadOnlyList<PptxSceneChartAxis> ReadSceneChartCategoryAxes(PptxSceneChart? chart, PptxSceneChartPlot? plot)
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
                .Select(axis => axis!)
                .ToArray();
        }

        return chart.Axes
            .Where(IsCategoryLike)
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
            if (sceneChart is not null)
            {
                return [];
            }

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

    private static ChartAxisSource ReadSceneOrXmlChartCategoryAxisForPlot(
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? scenePlot,
        XDocument chartXml,
        XElement? chartElement)
    {
        XElement[] xmlAxes = chartElement is null
            ? ReadChartCategoryAxes(chartXml).ToArray()
            : ReadChartCategoryAxesForChart(chartXml, chartElement).ToArray();
        IReadOnlyList<PptxSceneChartAxis> sceneAxes = ReadSceneChartCategoryAxes(sceneChart, scenePlot);
        if (sceneAxes.Count == 0)
        {
            if (sceneChart is not null)
            {
                return new ChartAxisSource(null, null);
            }

            XElement? xmlAxis = xmlAxes.FirstOrDefault();
            return new ChartAxisSource(null, xmlAxis);
        }

        PptxSceneChartAxis sceneAxis = sceneAxes[0];
        XElement? matchedXmlAxis = xmlAxes.FirstOrDefault(axis => string.Equals(ReadChartAxisId(axis), sceneAxis.Id, StringComparison.Ordinal));
        return new ChartAxisSource(sceneAxis, matchedXmlAxis);
    }

    private static ChartAxisSource ReadSceneOrXmlSecondaryRightValueAxis(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        XElement? xmlAxis = ReadSecondaryRightValueAxis(chartXml);
        if (sceneChart is null)
        {
            return new ChartAxisSource(null, xmlAxis);
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
        XElement? xmlAxis = ReadSecondaryValueAxisForChart(chartXml, primaryValueAxis.XmlAxis);
        if (sceneChart is null)
        {
            return new ChartAxisSource(null, xmlAxis);
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
            return ResolveChartGrouping(scenePlot.GroupingKind, defaultGrouping);
        }

        return ResolveChartGrouping(
            PptxSceneBuilder.ParseChartGrouping((string?)plotElement.Element(ChartNamespace + "grouping")?.Attribute("val")),
            defaultGrouping);
    }

    private static PptxSceneChartGrouping ResolveChartGrouping(PptxSceneChartGrouping value, PptxSceneChartGrouping defaultGrouping)
    {
        return value == PptxSceneChartGrouping.Unknown
            ? defaultGrouping
            : value;
    }

    private static PptxSceneChartBarDirection ReadSceneOrXmlChartBarDirection(PptxSceneChartPlot? scenePlot, XElement plotElement)
    {
        return scenePlot is not null
            ? ResolveChartBarDirection(scenePlot.BarDirectionKind)
            : ResolveChartBarDirection(PptxSceneBuilder.ParseChartBarDirection((string?)plotElement.Element(ChartNamespace + "barDir")?.Attribute("val")));
    }

    private static PptxSceneChartBarDirection ResolveChartBarDirection(PptxSceneChartBarDirection value)
    {
        return value == PptxSceneChartBarDirection.Unknown
            ? PptxSceneChartBarDirection.Column
            : value;
    }

    private static PptxSceneChartScatterStyle ReadSceneOrXmlChartScatterStyle(PptxSceneChartPlot? scenePlot, XElement plotElement)
    {
        return scenePlot is not null
            ? scenePlot.ScatterStyleKind
            : PptxSceneBuilder.ParseChartScatterStyle((string?)plotElement.Element(ChartNamespace + "scatterStyle")?.Attribute("val"));
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
            : ResolveChartRadarStyle(PptxSceneBuilder.ParseChartRadarStyle((string?)plotElement.Element(ChartNamespace + "radarStyle")?.Attribute("val")));
    }

    private static PptxSceneChartRadarStyle ResolveChartRadarStyle(PptxSceneChartRadarStyle value)
    {
        return value == PptxSceneChartRadarStyle.Unknown
            ? PptxSceneChartRadarStyle.Standard
            : value;
    }

    private static PptxSceneChartDisplayBlanksAs ReadSceneOrXmlChartDisplayBlanksAs(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            return ResolveChartDisplayBlanksAs(sceneChart.Options.DisplayBlanksAsKind);
        }

        XElement? chart = chartXml.Root?.Element(ChartNamespace + "chart");
        return ResolveChartDisplayBlanksAs(PptxSceneBuilder.ParseChartDisplayBlanksAs((string?)chart?.Element(ChartNamespace + "dispBlanksAs")?.Attribute("val")));
    }

    private static PptxSceneChartDisplayBlanksAs ResolveChartDisplayBlanksAs(PptxSceneChartDisplayBlanksAs value)
    {
        return value == PptxSceneChartDisplayBlanksAs.Unknown
            ? PptxSceneChartDisplayBlanksAs.Gap
            : value;
    }

    private static bool ReadSceneOrXmlChartPlotVisibleOnly(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            return sceneChart.Options.PlotVisibleOnly ?? true;
        }

        XElement? chart = chartXml.Root?.Element(ChartNamespace + "chart");
        return IsOoxmlBooleanElementEnabled(chart?.Element(ChartNamespace + "plotVisOnly"), defaultValue: true);
    }

    private static double ReadSceneDoughnutHoleSize(PptxSceneChartPlot? plot, XElement doughnutChart)
    {
        if (plot is not null)
        {
            return plot.HoleSize is { } rawHoleSize
                ? Math.Clamp(rawHoleSize / 100d, PptxChartMetricRules.DoughnutHoleMinimumRatio, PptxChartMetricRules.DoughnutHoleMaximumRatio)
                : PptxChartMetricRules.DoughnutHoleFallbackRatio;
        }

        return ReadDoughnutHoleSize(doughnutChart);
    }

    private static ChartBooleanOption ReadSceneOrXmlChartVaryColors(PptxSceneChartPlot? plot, XElement chartElement)
    {
        return plot is not null
            ? new ChartBooleanOption(plot.VaryColors ?? true, plot.VaryColorsValue, plot.VaryColors is not null)
            : ReadChartVaryColorsOption(chartElement);
    }

    private readonly record struct ChartBarPlotOptions(
        PptxSceneChartGrouping Grouping,
        PptxSceneChartBarDirection BarDirection,
        ChartBooleanOption VaryColors,
        double GapWidth,
        double Overlap);

    private static ChartBarPlotOptions ReadSceneOrXmlChartBarOptions(PptxSceneChartPlot? plot, XElement chartElement, PptxSceneChartGrouping defaultGrouping)
    {
        return new ChartBarPlotOptions(
            ReadSceneOrXmlChartGrouping(plot, chartElement, defaultGrouping),
            ReadSceneOrXmlChartBarDirection(plot, chartElement),
            ReadSceneOrXmlChartVaryColors(plot, chartElement),
            ReadSceneOrXmlChartGapWidth(plot, chartElement),
            ReadSceneOrXmlChartOverlap(plot, chartElement));
    }

    private readonly record struct ChartLinePlotOptions(
        PptxSceneChartGrouping Grouping,
        bool Stacked,
        bool PercentStacked,
        IReadOnlyList<ChartBooleanOption> SmoothSeries,
        PptxSceneChartDisplayBlanksAs DisplayBlanksAs);

    private static ChartLinePlotOptions ReadSceneOrXmlChartLineOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, XElement chartElement, PptxSceneChartGrouping defaultGrouping)
    {
        PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(plot, chartElement, defaultGrouping);
        return new ChartLinePlotOptions(
            grouping,
            IsStackedChartGrouping(grouping),
            IsPercentStackedChartGrouping(grouping),
            ReadSceneOrXmlSmoothSeries(plot, chartElement),
            ReadSceneOrXmlChartDisplayBlanksAs(sceneChart, chartXml));
    }

    private readonly record struct ChartAreaPlotOptions(
        PptxSceneChartGrouping Grouping,
        bool Stacked,
        bool PercentStacked,
        PptxSceneChartDisplayBlanksAs DisplayBlanksAs);

    private static ChartAreaPlotOptions ReadSceneOrXmlChartAreaOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, XElement chartElement, PptxSceneChartGrouping defaultGrouping)
    {
        PptxSceneChartGrouping grouping = ReadSceneOrXmlChartGrouping(plot, chartElement, defaultGrouping);
        return new ChartAreaPlotOptions(
            grouping,
            IsStackedChartGrouping(grouping),
            IsPercentStackedChartGrouping(grouping),
            ReadSceneOrXmlChartDisplayBlanksAs(sceneChart, chartXml));
    }

    private readonly record struct ChartScatterPlotOptions(
        PptxSceneChartScatterStyle ScatterStyle,
        bool ConnectLines,
        IReadOnlyList<ChartBooleanOption> SmoothSeries);

    private static ChartScatterPlotOptions ReadSceneOrXmlChartScatterOptions(PptxSceneChartPlot? plot, XElement chartElement)
    {
        PptxSceneChartScatterStyle scatterStyle = ReadSceneOrXmlChartScatterStyle(plot, chartElement);
        return new ChartScatterPlotOptions(
            scatterStyle,
            ResolveChartScatterLineConnection(scatterStyle),
            ReadSceneOrXmlSmoothSeries(plot, chartElement));
    }

    private readonly record struct ChartRadarPlotOptions(PptxSceneChartRadarStyle RadarStyle);

    private static ChartRadarPlotOptions ReadSceneOrXmlChartRadarOptions(PptxSceneChartPlot? plot, XElement chartElement)
    {
        return new ChartRadarPlotOptions(ReadSceneOrXmlChartRadarStyle(plot, chartElement));
    }

    private static double ReadSceneOrXmlChartGapWidth(PptxSceneChartPlot? plot, XElement chartElement)
    {
        return plot is not null
            ? plot.GapWidth ?? 150d
            : ReadChartGapWidth(chartElement);
    }

    private static double ReadSceneOrXmlChartOverlap(PptxSceneChartPlot? plot, XElement chartElement)
    {
        return plot is not null
            ? plot.Overlap ?? 0d
            : ReadChartOverlap(chartElement);
    }

    private static double ReadSceneOrXmlFirstSliceAngle(PptxSceneChartPlot? plot, XElement chartElement)
    {
        if (plot is not null)
        {
            return plot.FirstSliceAngle is { } sceneAngle
                ? NormalizeAngleDegrees(sceneAngle)
                : 0d;
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

    private static IReadOnlyList<ChartIndexedNumberVector> ReadSceneOrXmlChartSeriesVectors(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true)
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

    private static IReadOnlyList<ScatterSeries> ReadSceneOrXmlScatterSeries(PptxSceneChartPlot? plot, XElement chartElement, bool readBubbleSize, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true)
    {
        return ReadSceneOrXmlScatterSeriesVectors(plot, chartElement, readBubbleSize, workbook, plotVisibleOnly)
            .Select(BuildScatterSeries)
            .Where(series => series.Points.Count != 0)
            .ToArray();
    }

    private static IReadOnlyList<ChartIndexedScatterSeries> ReadSceneOrXmlScatterSeriesVectors(PptxSceneChartPlot? plot, XElement chartElement, bool readBubbleSize, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true)
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

    private static ChartIndexedTextVector ReadSceneOrXmlCategoryLabelVector(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true)
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
                .Select(point => new ChartIndexedNumberPoint(
                    point.Index,
                    point.HasParsedIndex ? ChartPointIndexSource.OoxmlIndex : ChartPointIndexSource.OrdinalFallback,
                    point.Value,
                    point.Text,
                    point.HasValueElement,
                    default))
                .ToArray()
            : compactValues
                .Select((value, index) => new ChartIndexedNumberPoint(index, ChartPointIndexSource.OrdinalFallback, value, value.ToString(CultureInfo.InvariantCulture), true, default))
                .ToArray();
        return new ChartIndexedNumberVector(points, pointCount ?? InferPointCount(points), source.Formula, formatCode, source, workbookPoints, plotVisibleOnly);
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
                .Select(point => new ChartIndexedTextPoint(
                    point.Index,
                    point.HasParsedIndex ? ChartPointIndexSource.OoxmlIndex : ChartPointIndexSource.OrdinalFallback,
                    point.Text,
                    point.HasText,
                    default))
                .ToArray()
            : compactValues
                .Select((value, index) => new ChartIndexedTextPoint(index, ChartPointIndexSource.OrdinalFallback, value, true, default))
                .ToArray();
        IReadOnlyList<IReadOnlyList<ChartIndexedTextPoint>> levels = categoryLevels
            .Select(level => level.Select(point => new ChartIndexedTextPoint(
                point.Index,
                point.HasParsedIndex ? ChartPointIndexSource.OoxmlIndex : ChartPointIndexSource.OrdinalFallback,
                point.Text,
                point.HasText,
                default)).ToArray())
            .ToArray();
        return new ChartIndexedTextVector(points, pointCount ?? InferPointCount(points), levels, source.Formula, source, workbookPoints, plotVisibleOnly);
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

        return ReadChartSeriesFills(chartElement, theme, colorMap);
    }

    private static IReadOnlyList<ChartSeriesStroke?> ReadSceneOrXmlSeriesStrokes(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap, double? inheritedWidth = null)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => ToChartSeriesStroke(series.Line, inheritedWidth))
                .ToArray();
        }

        return ReadChartSeriesStrokes(chartElement, theme, colorMap, inheritedWidth);
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
                    series.Marker.Fill.HasFill ? new ChartSeriesFill(series.Marker.Fill.Color, series.Marker.Fill.Alpha) : null,
                    ToChartSeriesStroke(series.Marker.Line, ChartMarkerInheritedStrokeWidth),
                    series.Marker.IsDefined))
                .ToArray();
        }

        return ReadChartMarkerStyles(chartElement, theme, colorMap, PptxSceneBuilder.ParseChartPlotKind(chartElement.Name.LocalName));
    }

    private static IReadOnlyList<ChartBooleanOption> ReadSceneOrXmlSmoothSeries(PptxSceneChartPlot? plot, XElement chartElement)
    {
        return plot is not null
            ? plot.Series.Select(series => new ChartBooleanOption(series.Smooth ?? false, series.SmoothValue, series.Smooth is not null)).ToArray()
            : ReadChartSeriesSmoothOptions(chartElement);
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> ReadSceneOrXmlSeriesPointFills(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return plot is not null
            ? plot.Series.Select(ReadSceneChartPointFills).ToArray()
            : ReadChartSeriesPointFills(chartElement, theme, colorMap);
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> ReadSceneOrXmlSeriesPointStrokes(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return plot is not null
            ? plot.Series.Select(ReadSceneChartPointStrokes).ToArray()
            : ReadChartSeriesPointStrokes(chartElement, theme, colorMap);
    }

    private static IReadOnlyDictionary<int, ChartSeriesFill> ReadSceneOrXmlChartPointFills(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        if (plot is not null)
        {
            return plot.Series.Count > 0 ? ReadSceneChartPointFills(plot.Series[0]) : new Dictionary<int, ChartSeriesFill>();
        }

        return ReadChartPointFills(chartElement, theme, colorMap);
    }

    private static IReadOnlyDictionary<int, ChartSeriesStroke> ReadSceneOrXmlChartPointStrokes(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        if (plot is not null)
        {
            return plot.Series.Count > 0 ? ReadSceneChartPointStrokes(plot.Series[0]) : new Dictionary<int, ChartSeriesStroke>();
        }

        return ReadChartPointStrokes(chartElement, theme, colorMap);
    }

    private static IReadOnlyDictionary<int, double> ReadSceneOrXmlChartPointExplosions(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook = null)
    {
        if (plot is not null)
        {
            return plot.Series.Count > 0 ? ReadSceneChartPointExplosions(plot.Series[0], workbook) : new Dictionary<int, double>();
        }

        return ReadChartPointExplosions(chartElement);
    }

    private readonly record struct ChartPolarPointOptions(
        IReadOnlyDictionary<int, ChartSeriesFill> PointFills,
        IReadOnlyDictionary<int, ChartSeriesStroke> PointStrokes,
        IReadOnlyDictionary<int, double> PointExplosions,
        double FirstSliceAngle);

    private static ChartPolarPointOptions ReadSceneOrXmlChartPolarPointOptions(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap, ChartWorkbookData? workbook = null)
    {
        return new ChartPolarPointOptions(
            ReadSceneOrXmlChartPointFills(plot, chartElement, theme, colorMap),
            ReadSceneOrXmlChartPointStrokes(plot, chartElement, theme, colorMap),
            ReadSceneOrXmlChartPointExplosions(plot, chartElement, workbook),
            ReadSceneOrXmlFirstSliceAngle(plot, chartElement));
    }

    private readonly record struct ChartDoughnutPlotOptions(
        ChartPolarPointOptions PolarPoints,
        double HoleSize);

    private static ChartDoughnutPlotOptions ReadSceneOrXmlChartDoughnutOptions(PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap, ChartWorkbookData? workbook = null)
    {
        return new ChartDoughnutPlotOptions(
            ReadSceneOrXmlChartPolarPointOptions(plot, chartElement, theme, colorMap, workbook),
            ReadSceneDoughnutHoleSize(plot, chartElement));
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

    private static IReadOnlyDictionary<int, double> ReadSceneChartPointExplosions(PptxSceneChartSeries series, ChartWorkbookData? workbook)
    {
        var explosions = new Dictionary<int, double>();
        if (series.Explosion is { } seriesExplosion)
        {
            double fraction = Math.Clamp(seriesExplosion / 100d, 0d, 1d);
            int pointCount = ResolveSceneChartSeriesPointCount(series, workbook);
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

    private static int ResolveSceneChartSeriesPointCount(PptxSceneChartSeries series, ChartWorkbookData? workbook)
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

    private static ChartSeriesStroke? ToChartSeriesStroke(PptxSceneLineStyle line, double? inheritedWidth = null)
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
            ToChartSeriesStroke(style.Line),
            style.Glow,
            style.OuterShadow);
    }

    private static bool TryRenderChart(PdfGraphicsBuilder graphics, PptxDocument document, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, ChartWorkbookData? workbook, List<PdfFontResource> fonts, PresentationFontResolver fontResolver)
    {
        bool plotVisibleOnly = ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, chartXml);
        IReadOnlyList<XElement> barCharts = ReadSceneOrXmlChartPlotElements(sceneChart, chartXml, PptxSceneChartPlotKind.Bar);
        XElement? barChart = barCharts.FirstOrDefault();
        if (barChart is not null)
        {
            PptxSceneChartPlot? barPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bar);
            IReadOnlyList<ChartIndexedNumberVector> barSeriesVectors = ReadSceneOrXmlChartSeriesVectors(barPlot, barChart, workbook, plotVisibleOnly);
            int barSeriesCount = CountRenderableSeries(barSeriesVectors);
            if (barSeriesCount != 0)
            {
                ChartBarPlotOptions barOptions = ReadSceneOrXmlChartBarOptions(barPlot, barChart, PptxSceneChartGrouping.Clustered);
                bool horizontalBars = barOptions.BarDirection == PptxSceneChartBarDirection.Bar;
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(barPlot, barChart, theme, colorMap);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, barPlot, chartXml, theme, barChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme, colorMap);
                ChartAxisSource valueAxisSource = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, barPlot, chartXml, barChart).FirstOrDefault();
                XElement? valueAxis = valueAxisSource.XmlAxis;
                PptxSceneChartAxis? valueSceneAxis = valueAxisSource.SceneAxis;
                bool percentStacked = IsPercentStackedChartGrouping(barOptions.Grouping);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueSceneAxis, valueAxis, GetBarChartValueExtents(barSeriesVectors, barOptions.Grouping), percentStacked);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(valueSceneAxis, valueAxis, theme, valueExtents, percentStacked);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(barPlot, barChart, theme, colorMap, ChartFilledSeriesInheritedStrokeWidth);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills = ReadSceneOrXmlSeriesPointFills(barPlot, barChart, theme, colorMap);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes = ReadSceneOrXmlSeriesPointStrokes(barPlot, barChart, theme, colorMap);
                var legendEntries = new List<ChartLegendEntry>(BuildFillLegendEntries(theme, colorMap, chartPalette, barPlot, barChart, seriesFills, seriesStrokes, workbook: workbook));
                ChartLayout chartLayout = GetBarChartLayout(document, theme, bounds, chartXml, sceneChart, colorMap, barPlot, barChart, barOptions, workbook, plotVisibleOnly, fontResolver);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                bool valueAxisLabelsVisible = IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis);
                RenderBarChart(graphics, theme, colorMap, chartPalette, chartLayout.PlotAreaBox, plotBox, barSeriesVectors, horizontalBars, barOptions.Grouping, seriesFills, pointFills, pointStrokes, valueAxisOptions.MajorGridlines, valueAxisOptions.MinorGridlines, valueAxisOptions.GridlineStyle, axesStyle, plotAreaStyle, valueExtents, valueAxisOptions.Units, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, valueAxisLabelsVisible, chartLayout.ManualPlotLayoutApplied, barOptions.VaryColors.Value, barOptions.GapWidth, barOptions.Overlap);
                XElement? secondaryValueAxis = null;
                PptxSceneChartAxis? secondaryValueSceneAxis = null;
                ChartValueExtents secondaryValueExtents = default;
                ChartAxisUnits secondaryAxisUnits = default;
                bool secondaryAxisReversed = false;
                int seriesOffset = barSeriesCount;
                int barChartIndex = 1;
                foreach (XElement extraBarChart in barCharts.Skip(1))
                {
                    PptxSceneChartPlot? extraBarPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bar, barChartIndex);
                    IReadOnlyList<ChartIndexedNumberVector> extraSeriesVectors = ReadSceneOrXmlChartSeriesVectors(extraBarPlot, extraBarChart, workbook, plotVisibleOnly);
                    int extraSeriesCount = CountRenderableSeries(extraSeriesVectors);
                    if (extraSeriesCount == 0)
                    {
                        barChartIndex++;
                        continue;
                    }

                    ChartBarPlotOptions extraBarOptions = ReadSceneOrXmlChartBarOptions(extraBarPlot, extraBarChart, PptxSceneChartGrouping.Clustered);
                    bool extraHorizontalBars = extraBarOptions.BarDirection == PptxSceneChartBarDirection.Bar;
                    ChartAxisSource extraValueAxisSource = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, extraBarPlot, chartXml, extraBarChart).FirstOrDefault();
                    XElement? extraValueAxis = extraValueAxisSource.XmlAxis;
                    PptxSceneChartAxis? extraValueSceneAxis = extraValueAxisSource.SceneAxis;
                    bool extraPercentStacked = IsPercentStackedChartGrouping(extraBarOptions.Grouping);
                    ChartValueExtents extraValueExtents = ReadPercentStackedAwareValueAxisExtents(extraValueSceneAxis, extraValueAxis, GetBarChartValueExtents(extraSeriesVectors, extraBarOptions.Grouping), extraPercentStacked);
                    ChartValueAxisRenderOptions extraValueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(extraValueSceneAxis, extraValueAxis, theme, extraValueExtents, extraPercentStacked);
                    IReadOnlyList<ChartSeriesFill?> extraSeriesFills = ReadSceneOrXmlSeriesFills(extraBarPlot, extraBarChart, theme, colorMap);
                    IReadOnlyList<ChartSeriesStroke?> extraSeriesStrokes = ReadSceneOrXmlSeriesStrokes(extraBarPlot, extraBarChart, theme, colorMap, ChartFilledSeriesInheritedStrokeWidth);
                    IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> extraPointFills = ReadSceneOrXmlSeriesPointFills(extraBarPlot, extraBarChart, theme, colorMap);
                    IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> extraPointStrokes = ReadSceneOrXmlSeriesPointStrokes(extraBarPlot, extraBarChart, theme, colorMap);
                    if (!extraHorizontalBars && secondaryValueAxis is null && IsSceneOrXmlVisibleValueAxis(extraValueSceneAxis, extraValueAxis))
                    {
                        secondaryValueAxis = extraValueAxis;
                        secondaryValueSceneAxis = extraValueSceneAxis;
                        secondaryValueExtents = extraValueExtents;
                        secondaryAxisUnits = extraValueAxisOptions.Units;
                        secondaryAxisReversed = extraValueAxisOptions.Reversed;
                    }

                    legendEntries.AddRange(BuildFillLegendEntries(theme, colorMap, chartPalette, extraBarPlot, extraBarChart, extraSeriesFills, extraSeriesStrokes, seriesOffset, workbook));
                    RenderBarChart(
                        graphics,
                        theme,
                        colorMap,
                        chartPalette,
                        chartLayout.PlotAreaBox,
                        plotBox,
                        extraSeriesVectors,
                        extraHorizontalBars,
                        extraBarOptions.Grouping,
                        extraSeriesFills,
                        extraPointFills,
                        extraPointStrokes,
                        majorGridlines: false,
                        minorGridlines: false,
                        ChartGridlineStyle.Empty,
                        axesStyle with { ValueAxisVisible = false, CategoryAxisVisible = false },
                        ChartShapeStyle.Empty,
                        extraValueExtents,
                        extraValueAxisOptions.Units,
                        extraValueAxisOptions.CrossingValue,
                        extraValueAxisOptions.Reversed,
                        valueAxisLabelsVisible: false,
                        manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied,
                        extraBarOptions.VaryColors.Value,
                        extraBarOptions.GapWidth,
                        extraBarOptions.Overlap);
                    fonts.AddRange(RenderBarDataLabels(
                        theme,
                        colorMap,
                        graphics,
                        plotBox,
                        extraSeriesVectors,
                        chartPalette,
                        extraValueExtents,
                        extraHorizontalBars,
                        extraValueAxisOptions.Reversed,
                        extraSeriesFills,
                        extraPointFills,
                        extraBarOptions.VaryColors.Value,
                        ReadSceneOrXmlDataLabelOptions(sceneChart, extraBarPlot, extraBarChart, theme, colorMap),
                        ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, extraBarPlot, extraBarChart, theme, colorMap),
                        ReadSceneOrXmlCategoryLabelVector(extraBarPlot, extraBarChart, workbook, plotVisibleOnly),
                        ReadSceneOrXmlChartSeriesNameRecords(extraBarPlot, extraBarChart, workbook), fontResolver));
                    seriesOffset += extraSeriesCount;
                    barChartIndex++;
                }

                int lineChartIndex = 0;
                foreach (XElement comboLineChart in ReadSceneOrXmlChartPlotElements(sceneChart, chartXml, PptxSceneChartPlotKind.Line))
                {
                    PptxSceneChartPlot? linePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Line, lineChartIndex);
                    IReadOnlyList<ChartIndexedNumberVector> lineSeriesVectors = ReadSceneOrXmlChartSeriesVectors(linePlot, comboLineChart, workbook, plotVisibleOnly);
                    if (CountRenderableSeries(lineSeriesVectors) == 0)
                    {
                        lineChartIndex++;
                        continue;
                    }

                    ChartAxisSource lineValueAxisSource = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, linePlot, chartXml, comboLineChart).FirstOrDefault();
                    XElement? lineValueAxis = lineValueAxisSource.XmlAxis;
                    XElement? lineValueAxisForScale = lineValueAxis ?? valueAxis;
                    PptxSceneChartAxis? lineValueSceneAxis = lineValueAxisSource.SceneAxis;
                    ChartLinePlotOptions lineOptions = ReadSceneOrXmlChartLineOptions(sceneChart, linePlot, chartXml, comboLineChart, PptxSceneChartGrouping.Standard);
                    ChartValueExtents lineValueExtents = ReadPercentStackedAwareValueAxisExtents(lineValueSceneAxis, lineValueAxisForScale, GetLineChartValueExtents(lineSeriesVectors, lineOptions.Stacked, lineOptions.PercentStacked), lineOptions.PercentStacked, useNearMaximumHeadroom: !lineOptions.PercentStacked);
                    ChartValueAxisRenderOptions lineValueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(lineValueSceneAxis, lineValueAxisForScale, theme, lineValueExtents, lineOptions.PercentStacked);
                    IReadOnlyList<ChartSeriesStroke?> lineSeriesStrokes = ReadSceneOrXmlSeriesStrokes(linePlot, comboLineChart, theme, colorMap, ChartSeriesInheritedStrokeWidth);
                    IReadOnlyList<ChartMarkerStyle> lineMarkerStyles = ReadSceneOrXmlMarkerStyles(linePlot, comboLineChart, theme, colorMap);
                    if (secondaryValueAxis is null && IsSceneOrXmlVisibleValueAxis(lineValueSceneAxis, lineValueAxis))
                    {
                        secondaryValueAxis = lineValueAxis;
                        secondaryValueSceneAxis = lineValueSceneAxis;
                        secondaryValueExtents = lineValueExtents;
                        secondaryAxisUnits = lineValueAxisOptions.Units;
                        secondaryAxisReversed = lineValueAxisOptions.Reversed;
                    }

                    legendEntries.AddRange(BuildStrokeLegendEntries(theme, colorMap, chartPalette, linePlot, comboLineChart, lineSeriesStrokes, lineMarkerStyles, reverseOrder: lineOptions.Stacked, workbook: workbook));
                    RenderLineChart(
                        graphics,
                        theme,
                        colorMap,
                        chartPalette,
                        chartLayout.PlotAreaBox,
                        plotBox,
                        lineSeriesVectors,
                        lineOptions.Stacked,
                        lineOptions.PercentStacked,
                        lineSeriesStrokes,
                        lineMarkerStyles,
                        lineOptions.SmoothSeries,
                        majorGridlines: false,
                        minorGridlines: false,
                        ChartGridlineStyle.Empty,
                        axesStyle with { ValueAxisVisible = false, CategoryAxisVisible = false },
                        ChartShapeStyle.Empty,
                        lineValueExtents,
                        lineValueAxisOptions.Units,
                        lineValueAxisOptions.CrossingValue,
                        lineValueAxisOptions.Reversed,
                        lineOptions.DisplayBlanksAs);
                    fonts.AddRange(RenderLineDataLabels(
                        theme,
                        colorMap,
                        graphics,
                        plotBox,
                        lineSeriesVectors,
                        lineValueExtents,
                        lineValueAxisOptions.Reversed,
                        lineSeriesStrokes,
                        lineMarkerStyles,
                        ReadSceneOrXmlDataLabelOptions(sceneChart, linePlot, comboLineChart, theme, colorMap),
                        ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, linePlot, comboLineChart, theme, colorMap),
                        ReadSceneOrXmlCategoryLabelVector(linePlot, comboLineChart, workbook, plotVisibleOnly),
                        ReadSceneOrXmlChartSeriesNameRecords(linePlot, comboLineChart, workbook), fontResolver));
                    lineChartIndex++;
                }

                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    double? categoryLabelAxisY = horizontalBars
                        ? null
                        : ChartValueToPlotCoordinate(valueExtents, valueAxisOptions.CrossingValue, plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                    fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSceneOrXmlCategoryLabelVector(barPlot, barChart, workbook, plotVisibleOnly), horizontalBars, categoryLabelAxisY, categoryLabelsOnTickMarks: ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueSceneAxis, valueAxis), categoryLabelsTopSide: ResolveSceneOrXmlCategoryAxisTopSide(categoryAxis.SceneAxis, categoryAxis.XmlAxis, defaultTopSide: false), fontResolver: fontResolver));
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
                        fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis, valueSceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied, useTextSizedWidth: sameSideSecondaryValueAxis, defaultNumberFormat: percentStacked ? "0%" : null, fontResolver: fontResolver));
                    }

                    if (!horizontalBars)
                    {
                        if (secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                        {
                            bool secondaryValueAxisRightSide = ResolveSceneOrXmlValueAxisRightSide(secondaryValueSceneAxis, secondaryValueAxis, axesStyle.SecondaryValueAxisRightSide);
                            int sideSlot = GetValueAxisSideSlot(valueSceneAxis, valueAxis, secondaryValueSceneAxis, secondaryValueAxis, defaultPrimaryRightSide: axesStyle.ValueAxisRightSide, defaultSecondaryRightSide: secondaryValueAxisRightSide);
                            fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, secondaryAxisReversed, horizontalBars: false, rightSide: secondaryValueAxisRightSide, axisSideSlot: sideSlot, useTextSizedWidth: sideSlot > 0, fontResolver: fontResolver));
                        }
                        else
                        {
                            fonts.AddRange(RenderSecondaryChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, GetBarChartValueExtents(barSeriesVectors, barOptions.Grouping), fontResolver));
                        }
                    }
                }
                else if (!horizontalBars && secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                {
                    fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, secondaryAxisReversed, horizontalBars: false, rightSide: ResolveSceneOrXmlValueAxisRightSide(secondaryValueSceneAxis, secondaryValueAxis, axesStyle.SecondaryValueAxisRightSide), fontResolver: fontResolver));
                }
                fonts.AddRange(RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, legendEntries, chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver));
                fonts.AddRange(RenderBarDataLabels(
                    theme,
                    colorMap,
                    graphics,
                    plotBox,
                    barSeriesVectors,
                    chartPalette,
                    valueExtents,
                    horizontalBars,
                    valueAxisOptions.Reversed,
                    seriesFills,
                    pointFills,
                    barOptions.VaryColors.Value,
                    ReadSceneOrXmlDataLabelOptions(sceneChart, barPlot, barChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, barPlot, barChart, theme, colorMap),
                    ReadSceneOrXmlCategoryLabelVector(barPlot, barChart, workbook, plotVisibleOnly),
                    ReadSceneOrXmlChartSeriesNameRecords(barPlot, barChart, workbook), fontResolver));
                return true;
            }
        }

        XElement? lineChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Line);
        if (lineChart is not null)
        {
            PptxSceneChartPlot? linePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Line);
            IReadOnlyList<ChartIndexedNumberVector> lineSeriesVectors = ReadSceneOrXmlChartSeriesVectors(linePlot, lineChart, workbook, plotVisibleOnly);
            if (CountRenderableSeries(lineSeriesVectors) != 0)
            {
                ChartLinePlotOptions lineOptions = ReadSceneOrXmlChartLineOptions(sceneChart, linePlot, chartXml, lineChart, PptxSceneChartGrouping.Standard);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(linePlot, lineChart, theme, colorMap, ChartSeriesInheritedStrokeWidth);
                IReadOnlyList<ChartMarkerStyle> markerStyles = ReadSceneOrXmlMarkerStyles(linePlot, lineChart, theme, colorMap);
                ChartAxesStyle axesStyle = ReadSceneOrXmlChartAxesStyle(sceneChart, linePlot, chartXml, theme, lineChart);
                ChartShapeStyle plotAreaStyle = ReadSceneOrXmlChartPlotAreaStyle(sceneChart, chartXml, theme, colorMap);
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, linePlot, chartXml, lineChart).FirstOrDefault();
                XElement? valueAxisForScale = ResolveXmlValueAxisForSource(sceneChart, valueAxis, chartXml);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, valueAxisForScale, GetLineChartValueExtents(lineSeriesVectors, lineOptions.Stacked, lineOptions.PercentStacked), lineOptions.PercentStacked, useNearMaximumHeadroom: !lineOptions.PercentStacked);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(valueAxis.SceneAxis, valueAxisForScale, theme, valueExtents, lineOptions.PercentStacked);
                ChartLayout chartLayout = GetLineChartLayout(document, theme, bounds, chartXml, sceneChart, colorMap, workbook, plotVisibleOnly, fontResolver);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                RenderLineChart(graphics, theme, colorMap, chartPalette, chartLayout.PlotAreaBox, plotBox, lineSeriesVectors, lineOptions.Stacked, lineOptions.PercentStacked, seriesStrokes, markerStyles, lineOptions.SmoothSeries, valueAxisOptions.MajorGridlines, valueAxisOptions.MinorGridlines, valueAxisOptions.GridlineStyle, axesStyle, plotAreaStyle, valueExtents, valueAxisOptions.Units, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, lineOptions.DisplayBlanksAs);
                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, linePlot, chartXml, lineChart);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSceneOrXmlCategoryLabelVector(linePlot, lineChart, workbook, plotVisibleOnly), horizontalBars: false, verticalAxisY: null, categoryLabelsOnTickMarks: ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis, valueAxisForScale), fontResolver: fontResolver));
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars: false, defaultNumberFormat: lineOptions.PercentStacked ? "0%" : null, fontResolver: fontResolver));
                    fonts.AddRange(RenderSecondaryChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, GetLineChartValueExtents(lineSeriesVectors, lineOptions.Stacked, lineOptions.PercentStacked), fontResolver));
                }
                fonts.AddRange(RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildStrokeLegendEntries(theme, colorMap, chartPalette, linePlot, lineChart, seriesStrokes, markerStyles, reverseOrder: lineOptions.Stacked, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver));
                fonts.AddRange(RenderLineDataLabels(
                    theme,
                    colorMap,
                    graphics,
                    plotBox,
                    lineSeriesVectors,
                    valueExtents,
                    valueAxisOptions.Reversed,
                    seriesStrokes,
                    markerStyles,
                    ReadSceneOrXmlDataLabelOptions(sceneChart, linePlot, lineChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, linePlot, lineChart, theme, colorMap),
                    ReadSceneOrXmlCategoryLabelVector(linePlot, lineChart, workbook, plotVisibleOnly),
                    ReadSceneOrXmlChartSeriesNameRecords(linePlot, lineChart, workbook), fontResolver));
                return true;
            }
        }

        XElement? areaChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Area);
        if (areaChart is not null)
        {
            PptxSceneChartPlot? areaPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Area);
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
                    areaOptions.Stacked,
                    areaOptions.PercentStacked,
                    seriesFills,
                    seriesStrokes,
                    valueAxisOptions.MajorGridlines,
                    valueAxisOptions.MinorGridlines,
                    valueAxisOptions.GridlineStyle,
                    axesStyle,
                    plotAreaStyle,
                    valueExtents,
                    valueAxisOptions.Units,
                    valueAxisOptions.CrossingValue,
                    valueAxisOptions.Reversed,
                    areaOptions.DisplayBlanksAs);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    fonts.AddRange(RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSceneOrXmlCategoryLabelVector(areaPlot, areaChart, workbook, plotVisibleOnly), horizontalBars: false, verticalAxisY: null, categoryLabelsOnTickMarks: ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueAxis.SceneAxis, valueAxis.XmlAxis), fontResolver: fontResolver));
                }

                if (axesStyle.ValueAxisVisible && IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars: false, defaultNumberFormat: areaOptions.PercentStacked ? "0%" : null, fontResolver: fontResolver));
                }

                fonts.AddRange(RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildFillLegendEntries(theme, colorMap, chartPalette, areaPlot, areaChart, seriesFills, seriesStrokes, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver));
                return true;
            }
        }

        XElement? scatterChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Scatter);
        if (scatterChart is not null)
        {
            PptxSceneChartPlot? scatterPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Scatter);
            IReadOnlyList<ScatterSeries> scatterSeries = ReadSceneOrXmlScatterSeries(scatterPlot, scatterChart, readBubbleSize: false, workbook: workbook, plotVisibleOnly: plotVisibleOnly);
            if (scatterSeries.Count != 0)
            {
                ChartScatterPlotOptions scatterOptions = ReadSceneOrXmlChartScatterOptions(scatterPlot, scatterChart);
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(scatterPlot, scatterChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(scatterPlot, scatterChart, theme, colorMap);
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
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildStrokeLegendEntries(theme, colorMap, chartPalette, scatterPlot, scatterChart, seriesStrokes, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver));
                return true;
            }
        }

        XElement? bubbleChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Bubble);
        if (bubbleChart is not null)
        {
            PptxSceneChartPlot? bubblePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bubble);
            IReadOnlyList<ScatterSeries> bubbleSeries = ReadSceneOrXmlScatterSeries(bubblePlot, bubbleChart, readBubbleSize: true, workbook: workbook, plotVisibleOnly: plotVisibleOnly);
            if (bubbleSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(bubblePlot, bubbleChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(bubblePlot, bubbleChart, theme, colorMap);
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
                fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, xValueAxis.XmlAxis, xValueAxis.SceneAxis, xExtents, xAxisOptions.Units, valueAxisReversed: false, horizontalBars: true, fontResolver: fontResolver));
                fonts.AddRange(RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, yValueAxis.XmlAxis, yValueAxis.SceneAxis, yExtents, yAxisOptions.Units, valueAxisReversed: false, horizontalBars: false, fontResolver: fontResolver));
                fonts.AddRange(RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, chartLayout.Frame, plotBox, BuildFillLegendEntries(theme, colorMap, chartPalette, bubblePlot, bubbleChart, seriesFills, seriesStrokes, workbook: workbook), chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver));
                return true;
            }
        }

        XElement? radarChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Radar);
        if (radarChart is not null)
        {
            PptxSceneChartPlot? radarPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Radar);
            IReadOnlyList<ChartIndexedNumberVector> radarSeriesVectors = ReadSceneOrXmlChartSeriesVectors(radarPlot, radarChart, workbook, plotVisibleOnly);
            IReadOnlyList<ChartRadarSeries> radarSeries = BuildRadarSeries(radarSeriesVectors);
            if (radarSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(radarPlot, radarChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(radarPlot, radarChart, theme, colorMap);
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, radarPlot, chartXml, radarChart).FirstOrDefault();
                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, radarPlot, chartXml, radarChart);
                ChartValueExtents valueExtents = ReadSceneOrXmlChartValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetRadarChartValueExtents(radarSeries));
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
            PptxSceneChartPlot? piePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Pie);
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
                fonts.AddRange(RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, piePlot, pieChart, polarPoints.PointFills, workbook, plotVisibleOnly), legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver));
                return true;
            }
        }

        XElement? doughnutChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Doughnut);
        if (doughnutChart is not null)
        {
            PptxSceneChartPlot? doughnutPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Doughnut);
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
                fonts.AddRange(RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, doughnutPlot, doughnutChart, polarPoints.PointFills, workbook, plotVisibleOnly), legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver));
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<ChartIndexedNumberVector> ReadChartSeriesVectors(XElement chartElement, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true)
    {
        var series = new List<ChartIndexedNumberVector>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            ChartIndexedNumberVector values = ReadChartNumberVector(element.Element(ChartNamespace + "val"), workbook, plotVisibleOnly);
            if (values.Points.Count > 0 || values.PointCount is not null || values.DensePoints().Count != 0)
            {
                series.Add(values);
            }
        }

        return series;
    }

    private static ChartIndexedNumberVector ReadChartNumberVector(XElement? container, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true)
    {
        if (container is null)
        {
            return default;
        }

        XElement? cache = container
            .Descendants(ChartNamespace + "numCache")
            .Concat(container.Descendants(ChartNamespace + "numLit"))
            .FirstOrDefault();
        PptxSceneChartDataSource source = ReadChartDataSource(container, "numRef");
        if (cache is null)
        {
            return new ChartIndexedNumberVector(
                [],
                null,
                source.Formula,
                null,
                source,
                ReadWorkbookNumberPoints(workbook, source),
                plotVisibleOnly);
        }

        ChartIndexedNumberPoint[] points = cache
            .Elements(ChartNamespace + "pt")
            .Select((point, fallbackIndex) =>
            {
                int? parsedIndex = ReadChartCachePointIndex(point);
                XElement? valueElement = point.Element(ChartNamespace + "v");
                string text = valueElement?.Value ?? string.Empty;
                bool hasValue = valueElement is not null;
                double? value = double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                    ? parsed
                    : null;
                return new ChartIndexedNumberPoint(
                    parsedIndex ?? fallbackIndex,
                    parsedIndex is not null ? ChartPointIndexSource.OoxmlIndex : ChartPointIndexSource.OrdinalFallback,
                    value,
                    text,
                    hasValue,
                    default);
            })
            .ToArray();
        return new ChartIndexedNumberVector(
            points,
            ReadChartPointCount(cache) ?? InferPointCount(points),
            source.Formula,
            (string?)cache.Element(ChartNamespace + "formatCode"),
            source,
            ReadWorkbookNumberPoints(workbook, source),
            plotVisibleOnly);
    }

    private static int? ReadChartPointCount(XElement cache)
    {
        string? value = (string?)cache.Element(ChartNamespace + "ptCount")?.Attribute("val");
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
            ? parsed
            : null;
    }

    private static int? ReadChartCachePointIndex(XElement point)
    {
        string? value = (string?)point.Attribute("idx");
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0
            ? parsed
            : null;
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

    private static ChartWorkbookData? ReadEmbeddedChartWorkbookData(PptxSceneChartExternalData sceneExternalData)
    {
        if (!sceneExternalData.IsDefined ||
            sceneExternalData.Resource is null)
        {
            return null;
        }

        using var stream = new MemoryStream(sceneExternalData.Resource.Bytes, writable: false);
        OoxPackage workbookPackage = OoxPackage.Open(stream);
        return ReadWorkbookData(workbookPackage);
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
            cellFormats.Add(new ChartWorkbookCellFormat(
                numberFormatId,
                numberFormatCode,
                applyNumberFormat,
                IsWorkbookDateLikeNumberFormat(numberFormatId, numberFormatCode)));
        }

        return new ChartWorkbookStyles(customNumberFormats, cellFormats);
    }

    private static bool IsWorkbookDateLikeNumberFormat(int? numberFormatId, string formatCode)
    {
        if (numberFormatId is { } id && IsBuiltInWorkbookDateLikeNumberFormatId(id))
        {
            return true;
        }

        return ContainsWorkbookDateTimeFormatToken(formatCode);
    }

    private static bool IsBuiltInWorkbookDateLikeNumberFormatId(int numberFormatId)
    {
        return (numberFormatId >= 14 && numberFormatId <= 22) ||
            (numberFormatId >= 27 && numberFormatId <= 36) ||
            (numberFormatId >= 45 && numberFormatId <= 47) ||
            (numberFormatId >= 50 && numberFormatId <= 58);
    }

    private static bool ContainsWorkbookDateTimeFormatToken(string formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode) ||
            string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        for (int i = 0; i < formatCode.Length; i++)
        {
            char c = formatCode[i];
            if (c == '"')
            {
                i++;
                while (i < formatCode.Length && formatCode[i] != '"')
                {
                    i++;
                }

                continue;
            }

            if (c == '\\' || c == '_' || c == '*')
            {
                i++;
                continue;
            }

            if (c == '[')
            {
                int closingBracket = formatCode.IndexOf(']', i + 1);
                if (closingBracket < 0)
                {
                    closingBracket = formatCode.Length - 1;
                }

                string bracketToken = formatCode.Substring(i + 1, Math.Max(0, closingBracket - i - 1)).Trim();
                if (bracketToken.Length > 0 &&
                    bracketToken.All(c => c == 'h' || c == 'H' || c == 'm' || c == 'M' || c == 's' || c == 'S'))
                {
                    return true;
                }

                i = closingBracket;
                continue;
            }

            char lower = char.ToLowerInvariant(c);
            if (lower == 'y' || lower == 'd' || lower == 'h' || lower == 's')
            {
                return true;
            }
        }

        return false;
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

            XElement? autoFilter = tableElement.Element(SpreadsheetNamespace + "autoFilter");
            string autoFilterReference = (string?)autoFilter?.Attribute("ref") ?? string.Empty;
            ChartWorkbookTableFilterColumn[] filterColumns = autoFilter?
                .Elements(SpreadsheetNamespace + "filterColumn")
                .Select(ReadWorkbookTableFilterColumn)
                .ToArray() ?? [];
            int[] filterColumnIds = filterColumns
                .Where(column => column.ColumnId is not null)
                .Select(column => column.ColumnId!.Value)
                .ToArray();
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
                filterColumnIds,
                filterColumns));
        }

        return tables;
    }

    private static ChartWorkbookTableFilterColumn ReadWorkbookTableFilterColumn(XElement column)
    {
        XElement? filters = column.Element(SpreadsheetNamespace + "filters");
        XElement? customFilters = column.Element(SpreadsheetNamespace + "customFilters");
        XElement? dynamicFilter = column.Element(SpreadsheetNamespace + "dynamicFilter");
        XElement? top10 = column.Element(SpreadsheetNamespace + "top10");
        string filterKind = column.Elements().FirstOrDefault()?.Name.LocalName ?? string.Empty;
        return new ChartWorkbookTableFilterColumn(
            ReadSpreadsheetIntegerAttribute(column, "colId"),
            ReadOoxmlBooleanAttribute(column, "hiddenButton"),
            ReadOoxmlBooleanAttribute(column, "showButton"),
            filterKind,
            filters?.Elements(SpreadsheetNamespace + "filter")
                .Select(filter => (string?)filter.Attribute("val") ?? string.Empty)
                .ToArray() ?? [],
            (string?)dynamicFilter?.Attribute("type") ?? string.Empty,
            (string?)top10?.Attribute("val") ?? string.Empty,
            ReadOoxmlBooleanAttribute(top10, "percent"),
            ReadOoxmlBooleanAttribute(top10, "top"),
            customFilters?.Elements(SpreadsheetNamespace + "customFilter")
                .Select(filter => new ChartWorkbookTableCustomFilter(
                    (string?)filter.Attribute("operator") ?? string.Empty,
                    (string?)filter.Attribute("val") ?? string.Empty))
                .ToArray() ?? []);
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
        bool? ApplyNumberFormat,
        bool NumberFormatIsDateLike);

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
        IReadOnlyList<int> FilterColumnIds,
        IReadOnlyList<ChartWorkbookTableFilterColumn> FilterColumns);

    private readonly record struct ChartWorkbookTableFilterColumn(
        int? ColumnId,
        bool? HiddenButton,
        bool? ShowButton,
        string FilterKind,
        IReadOnlyList<string> FilterValues,
        string DynamicFilterType,
        string Top10Value,
        bool? Top10Percent,
        bool? Top10Top,
        IReadOnlyList<ChartWorkbookTableCustomFilter> CustomFilters);

    private readonly record struct ChartWorkbookTableCustomFilter(
        string Operator,
        string Value);

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
        int RangeAreaIndex,
        int RangeAreaCount,
        int RangeRowIndex,
        int RangeColumnIndex,
        int RangeRowCount,
        int RangeColumnCount,
        int SheetRow,
        int SheetColumn,
        string Reference,
        string SheetName,
        string SourceFormula,
        string ResolvedFormula,
        ChartWorkbookRangeSourceKind SourceKind,
        string DefinedName,
        string DefinedNameSheetName,
        int? DefinedNameLocalSheetId,
        string TableName,
        string TableColumnName,
        int? TableColumnId,
        string TableFirstColumnName,
        int? TableFirstColumnId,
        string TableLastColumnName,
        int? TableLastColumnId,
        int? TableRowIndex,
        bool TableHeaderRow,
        bool TableDataRow,
        bool TableTotalsRow,
        int? TableCellColumnIndex,
        string TableCellColumnName,
        int? TableCellColumnId,
        string TableCellTotalsRowFunction,
        string TableCellTotalsRowFormula,
        string TableCellCalculatedColumnFormula,
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
        bool StyleNumberFormatIsDateLike,
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
        string DefinedNameSheetName,
        int? DefinedNameLocalSheetId,
        string TableName,
        string TableColumnName,
        int? TableColumnId,
        string TableFirstColumnName,
        int? TableFirstColumnId,
        string TableLastColumnName,
        int? TableLastColumnId);

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
            string[] rangeAreas = SplitRangeAreas(resolution.ResolvedFormula);
            if (rangeAreas.Length == 0)
            {
                return [];
            }

            var values = new List<ChartWorkbookRangeCell>();
            int index = 0;
            for (int areaIndex = 0; areaIndex < rangeAreas.Length; areaIndex++)
            {
                AddRangeAreaCells(values, resolution with { ResolvedFormula = rangeAreas[areaIndex] }, areaIndex, rangeAreas.Length, ref index);
            }

            return values.ToArray();
        }

        private void AddRangeAreaCells(
            List<ChartWorkbookRangeCell> values,
            ChartWorkbookRangeResolution resolution,
            int rangeAreaIndex,
            int rangeAreaCount,
            ref int index)
        {
            if (!TryParseRange(resolution.ResolvedFormula, out string? sheetName, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow) ||
                !sheets.TryGetValue(sheetName, out ChartWorksheetData? worksheet))
            {
                return;
            }

            int minColumn = Math.Min(firstColumn, lastColumn);
            int maxColumn = Math.Max(firstColumn, lastColumn);
            int minRow = Math.Min(firstRow, lastRow);
            int maxRow = Math.Max(firstRow, lastRow);
            int rangeRowCount = maxRow - minRow + 1;
            int rangeColumnCount = maxColumn - minColumn + 1;
            ChartWorkbookTable sourceTable = default;
            bool hasSourceTable = !string.IsNullOrWhiteSpace(resolution.TableName) &&
                tables.TryGetValue(resolution.TableName, out sourceTable);
            for (int row = minRow; row <= maxRow; row++)
            {
                for (int column = minColumn; column <= maxColumn; column++)
                {
                    string reference = ToCellReference(column, row);
                    bool hasCell = worksheet.Cells.TryGetValue(reference, out ChartWorkbookCell cell);
                    ChartWorkbookCellFormat format = hasCell ? styles.ResolveCellFormat(cell.StyleIndex) : default;
                    int? tableRowIndex = hasSourceTable ? row - sourceTable.FirstRow : null;
                    bool tableHeaderRow = hasSourceTable && tableRowIndex >= 0 && tableRowIndex < Math.Max(0, sourceTable.HeaderRowCount);
                    bool tableTotalsRow = hasSourceTable && sourceTable.TotalsRowCount > 0 && row > sourceTable.LastRow - sourceTable.TotalsRowCount && row <= sourceTable.LastRow;
                    bool tableDataRow = hasSourceTable && !tableHeaderRow && !tableTotalsRow && row >= sourceTable.FirstRow && row <= sourceTable.LastRow;
                    int? tableCellColumnIndex = null;
                    string tableCellColumnName = string.Empty;
                    int? tableCellColumnId = null;
                    string tableCellTotalsRowFunction = string.Empty;
                    string tableCellTotalsRowFormula = string.Empty;
                    string tableCellCalculatedColumnFormula = string.Empty;
                    if (hasSourceTable)
                    {
                        int columnIndex = column - sourceTable.FirstColumn;
                        if (columnIndex >= 0 && columnIndex < sourceTable.Columns.Count)
                        {
                            ChartWorkbookTableColumn tableCellColumn = sourceTable.Columns[columnIndex];
                            tableCellColumnIndex = columnIndex;
                            tableCellColumnName = tableCellColumn.Name;
                            tableCellColumnId = tableCellColumn.Id;
                            tableCellTotalsRowFunction = tableCellColumn.TotalsRowFunction;
                            tableCellTotalsRowFormula = tableCellColumn.TotalsRowFormula;
                            tableCellCalculatedColumnFormula = tableCellColumn.CalculatedColumnFormula;
                        }
                    }

                    values.Add(new ChartWorkbookRangeCell(
                        index,
                        rangeAreaIndex,
                        rangeAreaCount,
                        row - minRow,
                        column - minColumn,
                        rangeRowCount,
                        rangeColumnCount,
                        row,
                        column,
                        reference,
                        sheetName,
                        resolution.SourceFormula,
                        resolution.ResolvedFormula,
                        resolution.SourceKind,
                        resolution.DefinedName,
                        resolution.DefinedNameSheetName,
                        resolution.DefinedNameLocalSheetId,
                        resolution.TableName,
                        resolution.TableColumnName,
                        resolution.TableColumnId,
                        resolution.TableFirstColumnName,
                        resolution.TableFirstColumnId,
                        resolution.TableLastColumnName,
                        resolution.TableLastColumnId,
                        tableRowIndex,
                        tableHeaderRow,
                        tableDataRow,
                        tableTotalsRow,
                        tableCellColumnIndex,
                        tableCellColumnName,
                        tableCellColumnId,
                        tableCellTotalsRowFunction,
                        tableCellTotalsRowFormula,
                        tableCellCalculatedColumnFormula,
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
                        format.NumberFormatIsDateLike,
                        worksheet.HiddenRows.Contains(row),
                        worksheet.HiddenColumns.Contains(column)));
                    index++;
                }
            }
        }

        private ChartWorkbookRangeResolution ResolveRangeFormula(string? formula)
        {
            string sourceFormula = formula?.Trim() ?? string.Empty;
            if (sourceFormula.Length == 0)
            {
                return new ChartWorkbookRangeResolution(string.Empty, string.Empty, ChartWorkbookRangeSourceKind.Unknown, string.Empty, string.Empty, null, string.Empty, string.Empty, null, string.Empty, null, string.Empty, null);
            }

            string resolvedFormula = sourceFormula;
            string definedNameName = string.Empty;
            string definedNameSheetName = string.Empty;
            int? definedNameLocalSheetId = null;
            bool fromDefinedName = false;
            if (!sourceFormula.Contains('!', StringComparison.Ordinal) &&
                definedNames.TryGetValue(sourceFormula, out string? definedFormula))
            {
                resolvedFormula = definedFormula;
                definedNameName = sourceFormula;
                fromDefinedName = true;
            }
            else if (TryResolveQualifiedLocalDefinedName(sourceFormula, out ChartWorkbookDefinedName localDefinedName))
            {
                resolvedFormula = localDefinedName.Formula;
                definedNameName = localDefinedName.Name;
                definedNameSheetName = localDefinedName.SheetName;
                definedNameLocalSheetId = localDefinedName.LocalSheetId;
                fromDefinedName = true;
            }

            string structuredResolvedFormula = ResolveStructuredReferenceFormula(
                resolvedFormula,
                out string tableName,
                out string tableColumnName,
                out int? tableColumnId,
                out string tableFirstColumnName,
                out int? tableFirstColumnId,
                out string tableLastColumnName,
                out int? tableLastColumnId) ?? string.Empty;
            bool fromStructuredReference = !string.Equals(structuredResolvedFormula, resolvedFormula, StringComparison.Ordinal);
            resolvedFormula = structuredResolvedFormula;

            ChartWorkbookRangeSourceKind sourceKind = (fromDefinedName, fromStructuredReference) switch
            {
                (true, true) => ChartWorkbookRangeSourceKind.DefinedNameStructuredReference,
                (true, false) => ChartWorkbookRangeSourceKind.DefinedName,
                (false, true) => ChartWorkbookRangeSourceKind.StructuredReference,
                _ => ChartWorkbookRangeSourceKind.DirectRange
            };
            return new ChartWorkbookRangeResolution(
                sourceFormula,
                resolvedFormula,
                sourceKind,
                definedNameName,
                definedNameSheetName,
                definedNameLocalSheetId,
                tableName,
                tableColumnName,
                tableColumnId,
                tableFirstColumnName,
                tableFirstColumnId,
                tableLastColumnName,
                tableLastColumnId);
        }

        private bool TryResolveQualifiedLocalDefinedName(string sourceFormula, out ChartWorkbookDefinedName definedName)
        {
            definedName = default;
            int separator = sourceFormula.LastIndexOf('!');
            if (separator <= 0 || separator == sourceFormula.Length - 1)
            {
                return false;
            }

            string sheetName = NormalizeSheetName(sourceFormula[..separator]);
            string name = sourceFormula[(separator + 1)..].Trim();
            if (sheetName.Length == 0 ||
                name.Length == 0 ||
                name.Contains('$', StringComparison.Ordinal) ||
                name.Contains(':', StringComparison.Ordinal))
            {
                return false;
            }

            foreach (ChartWorkbookDefinedName candidate in definedNameRecords)
            {
                if (candidate.LocalSheetId is not null &&
                    string.Equals(candidate.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    definedName = candidate;
                    return true;
                }
            }

            return false;
        }

        private string? ResolveStructuredReferenceFormula(
            string? formula,
            out string tableName,
            out string tableColumnName,
            out int? tableColumnId,
            out string tableFirstColumnName,
            out int? tableFirstColumnId,
            out string tableLastColumnName,
            out int? tableLastColumnId)
        {
            tableName = string.Empty;
            tableColumnName = string.Empty;
            tableColumnId = null;
            tableFirstColumnName = string.Empty;
            tableFirstColumnId = null;
            tableLastColumnName = string.Empty;
            tableLastColumnId = null;
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
            if (!TryParseStructuredReferenceBody(
                body,
                out string? firstColumnName,
                out string? lastColumnName,
                out bool includeHeader,
                out bool onlyHeader,
                out bool onlyTotals,
                out bool wholeTable))
            {
                return trimmed;
            }

            int firstColumn = table.FirstColumn;
            int lastColumn = table.LastColumn;
            ChartWorkbookTableColumn tableColumn = default;
            ChartWorkbookTableColumn firstTableColumn = table.Columns.Count > 0 ? table.Columns[0] : default;
            ChartWorkbookTableColumn lastTableColumn = table.Columns.Count > 0 ? table.Columns[^1] : default;
            if (!wholeTable)
            {
                if (string.IsNullOrWhiteSpace(firstColumnName))
                {
                    return trimmed;
                }

                lastColumnName = string.IsNullOrWhiteSpace(lastColumnName) ? firstColumnName : lastColumnName;
                int firstColumnOffset = -1;
                int lastColumnOffset = -1;
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (string.Equals(table.Columns[i].Name, firstColumnName, StringComparison.OrdinalIgnoreCase))
                    {
                        firstColumnOffset = i;
                    }

                    if (string.Equals(table.Columns[i].Name, lastColumnName, StringComparison.OrdinalIgnoreCase))
                    {
                        lastColumnOffset = i;
                    }
                }

                if (firstColumnOffset < 0 || lastColumnOffset < 0)
                {
                    return trimmed;
                }

                int minColumnOffset = Math.Min(firstColumnOffset, lastColumnOffset);
                int maxColumnOffset = Math.Max(firstColumnOffset, lastColumnOffset);
                firstColumn = table.FirstColumn + minColumnOffset;
                lastColumn = table.FirstColumn + maxColumnOffset;
                firstTableColumn = table.Columns[minColumnOffset];
                lastTableColumn = table.Columns[maxColumnOffset];
                if (minColumnOffset == maxColumnOffset)
                {
                    tableColumn = firstTableColumn;
                }
            }

            int headerRowCount = Math.Max(0, table.HeaderRowCount);
            int totalsRowCount = Math.Max(0, table.TotalsRowCount);
            int firstRow = onlyHeader
                ? table.FirstRow
                : onlyTotals
                    ? table.LastRow - totalsRowCount + 1
                    : includeHeader
                        ? table.FirstRow
                        : table.FirstRow + headerRowCount;
            int lastRow = onlyHeader
                ? table.FirstRow + headerRowCount - 1
                : onlyTotals
                    ? table.LastRow
                    : includeHeader
                        ? table.LastRow
                        : table.LastRow - totalsRowCount;
            if (firstRow > lastRow)
            {
                return trimmed;
            }

            tableName = table.Name;
            bool singleColumn = !wholeTable && firstColumn == lastColumn;
            tableColumnName = singleColumn ? tableColumn.Name : string.Empty;
            tableColumnId = singleColumn ? tableColumn.Id : null;
            tableFirstColumnName = firstTableColumn.Name;
            tableFirstColumnId = firstTableColumn.Id;
            tableLastColumnName = lastTableColumn.Name;
            tableLastColumnId = lastTableColumn.Id;
            return FormattableString.Invariant($"{QuoteSheetName(table.SheetName)}!{ToCellReference(firstColumn, firstRow)}:{ToCellReference(lastColumn, lastRow)}");
        }

        private static bool TryParseStructuredReferenceBody(
            string body,
            out string firstColumnName,
            out string lastColumnName,
            out bool includeHeader,
            out bool onlyHeader,
            out bool onlyTotals,
            out bool wholeTable)
        {
            firstColumnName = string.Empty;
            lastColumnName = string.Empty;
            includeHeader = false;
            onlyHeader = false;
            onlyTotals = false;
            wholeTable = false;
            string trimmed = body.Trim();
            if (trimmed.Length == 0)
            {
                return false;
            }

            if (trimmed[0] != '[')
            {
                string segment = UnescapeStructuredReferenceText(trimmed);
                if (ApplyStructuredReferenceItem(segment, ref includeHeader, ref onlyHeader, ref onlyTotals, ref wholeTable))
                {
                    return true;
                }

                firstColumnName = segment;
                lastColumnName = segment;
                return true;
            }

            foreach (string segment in ParseStructuredReferenceSegments(trimmed))
            {
                if (ApplyStructuredReferenceItem(segment, ref includeHeader, ref onlyHeader, ref onlyTotals, ref wholeTable))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(firstColumnName))
                {
                    firstColumnName = segment;
                }

                lastColumnName = segment;
                wholeTable = false;
            }

            return wholeTable || !string.IsNullOrWhiteSpace(firstColumnName);
        }

        private static bool ApplyStructuredReferenceItem(string segment, ref bool includeHeader, ref bool onlyHeader, ref bool onlyTotals, ref bool wholeTable)
        {
            if (string.Equals(segment, "#All", StringComparison.OrdinalIgnoreCase))
            {
                includeHeader = true;
                wholeTable = true;
                return true;
            }

            if (string.Equals(segment, "#Headers", StringComparison.OrdinalIgnoreCase))
            {
                includeHeader = true;
                onlyHeader = true;
                wholeTable = true;
                return true;
            }

            if (string.Equals(segment, "#Data", StringComparison.OrdinalIgnoreCase))
            {
                wholeTable = true;
                return true;
            }

            if (string.Equals(segment, "#Totals", StringComparison.OrdinalIgnoreCase))
            {
                onlyTotals = true;
                wholeTable = true;
                return true;
            }

            return false;
        }

        private static string[] ParseStructuredReferenceSegments(string body)
        {
            var segments = new List<string>();
            int index = 0;
            while (index < body.Length)
            {
                while (index < body.Length && (char.IsWhiteSpace(body[index]) || body[index] == ',' || body[index] == ':'))
                {
                    index++;
                }

                if (index >= body.Length)
                {
                    break;
                }

                if (body[index] != '[')
                {
                    int nextComma = body.IndexOf(',', index);
                    string segment = nextComma < 0 ? body[index..] : body[index..nextComma];
                    segments.Add(UnescapeStructuredReferenceText(segment.Trim()));
                    index = nextComma < 0 ? body.Length : nextComma + 1;
                    continue;
                }

                index++;
                var segmentBuilder = new StringBuilder();
                while (index < body.Length)
                {
                    char current = body[index];
                    if (current == '\'' && index + 1 < body.Length)
                    {
                        segmentBuilder.Append(body[index + 1]);
                        index += 2;
                        continue;
                    }

                    if (current == ']')
                    {
                        index++;
                        break;
                    }

                    segmentBuilder.Append(current);
                    index++;
                }

                segments.Add(segmentBuilder.ToString().Trim());
            }

            return segments.ToArray();
        }

        private static string UnescapeStructuredReferenceText(string text)
        {
            var builder = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\'' && i + 1 < text.Length)
                {
                    builder.Append(text[i + 1]);
                    i++;
                }
                else
                {
                    builder.Append(text[i]);
                }
            }

            return builder.ToString().Trim();
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

        private static string[] SplitRangeAreas(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
            {
                return [];
            }

            var areas = new List<string>();
            int areaStart = 0;
            bool inQuotedSheetName = false;
            for (int index = 0; index < formula.Length; index++)
            {
                char current = formula[index];
                if (current == '\'')
                {
                    if (inQuotedSheetName && index + 1 < formula.Length && formula[index + 1] == '\'')
                    {
                        index++;
                        continue;
                    }

                    inQuotedSheetName = !inQuotedSheetName;
                    continue;
                }

                if (current == ',' && !inQuotedSheetName)
                {
                    string area = formula[areaStart..index].Trim();
                    if (area.Length > 0)
                    {
                        areas.Add(area);
                    }

                    areaStart = index + 1;
                }
            }

            string lastArea = formula[areaStart..].Trim();
            if (lastArea.Length > 0)
            {
                areas.Add(lastArea);
            }

            string currentSheetPrefix = string.Empty;
            for (int index = 0; index < areas.Count; index++)
            {
                int separator = areas[index].LastIndexOf('!');
                if (separator > 0)
                {
                    currentSheetPrefix = areas[index][..(separator + 1)];
                }
                else if (currentSheetPrefix.Length > 0)
                {
                    areas[index] = currentSheetPrefix + areas[index];
                }
            }

            return areas.ToArray();
        }

        private static string NormalizeSheetName(string sheetName)
        {
            sheetName = sheetName.Trim();
            if (sheetName.Length >= 2 && sheetName[0] == '\'' && sheetName[^1] == '\'')
            {
                sheetName = sheetName[1..^1].Replace("''", "'", StringComparison.Ordinal);
            }

            int workbookEnd = sheetName.LastIndexOf(']');
            if (workbookEnd >= 0 && workbookEnd < sheetName.Length - 1)
            {
                sheetName = sheetName[(workbookEnd + 1)..];
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

    private static ChartRadarLayout ResolveRadarLayout(ChartPlotBox plotBox, PptxSceneChartRadarStyle radarStyle, IReadOnlyList<ChartRadarSeries> series)
    {
        ChartRadarStyle style = radarStyle == PptxSceneChartRadarStyle.Filled
            ? ChartRadarStyle.Filled
            : ChartRadarStyle.Marker;
        return new ChartRadarLayout(
            plotBox,
            GetRadarChartGeometry(plotBox, style),
            style,
            Math.Max(3, series.Max(item => item.Points.Count)),
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

    private static void RenderChartAreaStyle(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, PptxTheme theme, PptxColorMap colorMap)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        RenderChartShapeStyle(graphics, frame.X, frame.Y, frame.Width, frame.Height, ReadSceneOrXmlChartAreaStyle(sceneChart, chartXml, theme, colorMap));
    }

    private static ChartShapeStyle ReadSceneOrXmlChartAreaStyle(PptxSceneChart? sceneChart, XDocument chartXml, PptxTheme theme, PptxColorMap colorMap)
    {
        return sceneChart is null
            ? ReadChartShapeStyle(chartXml.Root?.Element(ChartNamespace + "spPr"), theme, colorMap)
            : ToChartShapeStyle(sceneChart.ChartAreaStyle);
    }

    private static ChartShapeStyle ReadSceneOrXmlChartPlotAreaStyle(PptxSceneChart? sceneChart, XDocument chartXml, PptxTheme theme, PptxColorMap colorMap)
    {
        return sceneChart is null
            ? ReadChartPlotAreaStyle(chartXml, theme, colorMap)
            : ToChartShapeStyle(sceneChart.PlotAreaStyle);
    }

    private static ChartShapeStyle ReadChartPlotAreaStyle(XDocument chartXml, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? shapeProperties = chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "spPr");
        return ReadChartShapeStyle(shapeProperties, theme, colorMap);
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

    private static bool TryBuildManualLayoutBox(PptxSceneChartManualLayout layout, ChartFrameBox frame, ChartLayoutBox defaultBox, out ChartLayoutBox box, bool clampToFrame = true, bool missingPositionModesAreFactor = false)
    {
        box = default;
        if (!layout.HasLayout)
        {
            return false;
        }

        ChartPlotBoxRatios defaults = GetLayoutBoxRatios(frame, defaultBox);
        double left = layout.X is { } x
            ? ClampManualLayoutRatio(ResolveManualLayoutStartRatio(x, layout.XModeKind, layout.XMode, defaultBox.X, frame.X, frame.Width, missingPositionModesAreFactor), clampToFrame)
            : defaults.Left;
        double top = layout.Y is { } y
            ? ClampManualLayoutRatio(ResolveManualLayoutStartRatio(y, layout.YModeKind, layout.YMode, frame.Y + frame.Height - defaultBox.Y - defaultBox.Height, 0d, frame.Height, missingPositionModesAreFactor), clampToFrame)
            : defaults.Top;
        double width = layout.Width is { } layoutWidth
            ? Math.Clamp(layoutWidth, 0.02d, 1d)
            : defaults.Width;
        double height = layout.Height is { } layoutHeight
            ? Math.Clamp(layoutHeight, 0.02d, 1d)
            : defaults.Height;
        double right = IsManualLayoutEdgeMode(layout.WidthModeKind)
            ? ClampManualLayoutEdgeRatio(layout.Width ?? defaults.Right, left, clampToFrame)
            : left + width;
        double bottom = IsManualLayoutEdgeMode(layout.HeightModeKind)
            ? ClampManualLayoutEdgeRatio(layout.Height ?? defaults.Bottom, top, clampToFrame)
            : top + height;
        double boxWidth = Math.Max(0d, right - left) * frame.Width;
        double boxHeight = Math.Max(0d, bottom - top) * frame.Height;
        double boxX = frame.X + left * frame.Width;
        double boxY = frame.Y + frame.Height - bottom * frame.Height;
        box = new ChartLayoutBox(boxX, boxY, boxWidth, boxHeight);
        return boxWidth > 0d && boxHeight > 0d;
    }

    private static double ClampManualLayoutRatio(double value, bool clampToFrame)
    {
        return clampToFrame ? Math.Clamp(value, 0d, 1d) : value;
    }

    private static double ClampManualLayoutEdgeRatio(double value, double minimum, bool clampToFrame)
    {
        return clampToFrame ? Math.Clamp(value, minimum, 1d) : Math.Max(minimum, value);
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

    private static double ResolveManualLayoutStartRatio(double value, PptxSceneChartManualLayoutMode mode, string modeValue, double defaultStart, double frameStart, double frameLength, bool missingModeIsFactor)
    {
        if (IsManualLayoutFactorMode(mode, modeValue, missingModeIsFactor) && frameLength > 0d)
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

    private static bool IsManualLayoutFactorMode(PptxSceneChartManualLayoutMode mode, string modeValue, bool missingModeIsFactor)
    {
        return IsManualLayoutFactorMode(mode) || (missingModeIsFactor && string.IsNullOrEmpty(modeValue));
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
        (double? x, string xValue) = ReadManualLayoutFactorWithValue(manualLayout, "x");
        (double? y, string yValue) = ReadManualLayoutFactorWithValue(manualLayout, "y");
        (double? width, string widthValue) = ReadManualLayoutFactorWithValue(manualLayout, "w");
        (double? height, string heightValue) = ReadManualLayoutFactorWithValue(manualLayout, "h");
        return new PptxSceneChartManualLayout(
            true,
            x,
            xValue,
            y,
            yValue,
            width,
            widthValue,
            height,
            heightValue,
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
        (double? parsed, _) = ReadManualLayoutFactorWithValue(manualLayout, elementName);
        return parsed;
    }

    private static (double? Value, string RawValue) ReadManualLayoutFactorWithValue(XElement manualLayout, string elementName)
    {
        string value = (string?)manualLayout.Element(ChartNamespace + elementName)?.Attribute("val") ?? string.Empty;
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? (parsed, value)
            : (null, value);
    }

    private static string ReadManualLayoutValue(XElement manualLayout, string elementName)
    {
        return (string?)manualLayout.Element(ChartNamespace + elementName)?.Attribute("val") ?? string.Empty;
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

    private static IReadOnlyList<XElement> ReadChartCategoryAxesForChart(XDocument chartXml, XElement chartElement)
    {
        string[] axisIds = chartElement
            .Elements(ChartNamespace + "axId")
            .Select(axis => (string?)axis.Attribute("val"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id!)
            .ToArray();
        if (axisIds.Length == 0)
        {
            return ReadChartCategoryAxes(chartXml).ToArray();
        }

        var axes = new List<XElement>();
        foreach (string axisId in axisIds)
        {
            XElement? axis = ReadChartCategoryAxes(chartXml)
                .FirstOrDefault(candidate => string.Equals(ReadChartAxisId(candidate), axisId, StringComparison.Ordinal));
            if (axis is not null)
            {
                axes.Add(axis);
            }
        }

        return axes.Count == 0
            ? ReadChartCategoryAxes(chartXml).Take(1).ToArray()
            : axes;
    }

    private static IEnumerable<XElement> ReadChartCategoryAxes(XDocument chartXml)
    {
        return chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()?
            .Elements()
            .Where(element => element.Name == ChartNamespace + "catAx" || element.Name == ChartNamespace + "dateAx") ??
            [];
    }

    private static ChartValueExtents ReadChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback, bool useNearMaximumHeadroom = false, double nearMaximumHeadroomRatio = PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio)
    {
        return ReadChartValueAxisExtents(valueAxis, fallback, PptxChartMetricRules.AxisNiceTickTargetCount, useNearMaximumHeadroom, nearMaximumHeadroomRatio);
    }

    private static ChartValueExtents ReadBubbleChartValueAxisExtents(XElement? valueAxis, ChartValueExtents fallback)
    {
        return ReadChartValueAxisExtents(valueAxis, fallback, PptxChartMetricRules.BubbleAxisBoundsTickTargetCount);
    }

    private static ChartValueExtents ReadSceneOrXmlBubbleChartValueAxisExtents(PptxSceneChartAxis? axis, XElement? valueAxis, ChartValueExtents fallback)
    {
        if (axis is null)
        {
            return ReadBubbleChartValueAxisExtents(valueAxis, fallback);
        }

        if (!axis.HasScaling)
        {
            return fallback;
        }

        double min = axis.Minimum ?? GetNiceChartAxisMin(fallback.Min, fallback.Max);
        double max = axis.Maximum ?? GetNiceChartAxisMax(fallback.Max, min, PptxChartMetricRules.BubbleAxisBoundsTickTargetCount);
        return max > min
            ? new ChartValueExtents(min, max)
            : fallback;
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
        return ReadChartShapeStyle(shapeProperties, theme, PptxColorMap.Default);
    }

    private static ChartShapeStyle ReadChartShapeStyle(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        if (shapeProperties is null)
        {
            return ChartShapeStyle.Empty;
        }

        bool noFill = shapeProperties.Element(DrawingNamespace + "noFill") is not null;
        ChartSeriesFill? fill = !noFill && TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor fillColor, out double fillAlpha)
            ? new ChartSeriesFill(fillColor, fillAlpha)
            : null;
        GradientFill? gradientFill = !noFill && PptxSceneBuilder.TryReadShapeGradientFill(shapeProperties, theme, colorMap, out PptxSceneGradientFill sceneGradientFill)
            ? ToGradientFill(sceneGradientFill)
            : null;
        ChartSeriesStroke? stroke = TryReadLineWithAlpha(shapeProperties, theme, colorMap, out RgbColor strokeColor, out double lineWidth, out double strokeAlpha)
            ? new ChartSeriesStroke(strokeColor, strokeAlpha, lineWidth)
            : null;
        return new ChartShapeStyle(
            fill,
            gradientFill,
            stroke,
            PptxSceneBuilder.TryReadGlow(shapeProperties, theme, colorMap, out PptxSceneGlow glow) ? glow : default,
            PptxSceneBuilder.TryReadOuterShadow(shapeProperties, theme, colorMap, out PptxSceneOuterShadow outerShadow) ? outerShadow : default);
    }

    private static void RenderChartShapeStyle(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartShapeStyle style)
    {
        if (ToGlow(style.Glow) is { } glow)
        {
            DrawGlow(graphics, "rect", x, y, width, height, glow);
        }

        if (ToOuterShadow(style.OuterShadow) is { } outerShadow)
        {
            DrawOuterShadow(graphics, "rect", x, y, width, height, outerShadow);
        }

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
        graphics.ClipRectangleEvenOdd(x, y, width, height);
    }

    private static void RenderInChartPlotAreaClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, Action render)
    {
        graphics.SaveState();
        try
        {
            ClipChartPlotArea(graphics, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height);
            render();
        }
        finally
        {
            graphics.RestoreState();
        }
    }

    private static IReadOnlyList<PdfFontResource> RenderChartTitle(PptxDocument document, PptxTheme theme, PptxColorMap colorMap, PdfGraphicsBuilder graphics, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, ChartWorkbookData? workbook, bool plotVisibleOnly, PresentationFontResolver fontResolver)
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
        ChartTextStyle style = ReadSceneOrXmlChartTitleTextStyle(theme, colorMap, sceneChart, chartXml, isAutoTitle);
        double fontSize = style.FontSize;
        double titleBaselineRatio = !titleManualLayoutApplied && HasSceneOrXmlPolarChart(sceneChart, chartXml)
            ? PptxChartMetricRules.PolarTitleBaselineYRatio
            : PptxChartMetricRules.TitleBaselineYRatio;
        double fallbackBaselineY = y + height * titleBaselineRatio;
        double baselineY = titleManualLayoutApplied
            ? fallbackBaselineY
            : ResolveChartTitleBaselineY(document, theme, colorMap, bounds, chartXml, sceneChart, workbook, plotVisibleOnly, fallbackBaselineY, fontSize, fontResolver);
        RenderChartShapeStyle(graphics, x, y, width, height, ReadSceneOrXmlChartTitleShapeStyle(theme, colorMap, sceneChart, chartXml));
        double titleX = x + width * PptxChartMetricRules.TitleXInsetRatio;
        double titleWidth = width * PptxChartMetricRules.TitleWidthRatio;
        double titleHeight = fontSize * PptxChartMetricRules.TitleHeightFactor;
        var runs = new List<TextRun>();
        AddChartRichTextRuns(
            runs,
            ReadSceneOrXmlChartTitleTextRuns(theme, colorMap, sceneChart, chartXml),
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
            TextAlignment.Center,
            fontResolver);
        return RenderTextRuns(runs, graphics, "CT", fontResolver);
    }

    private static ChartShapeStyle ReadSceneOrXmlChartTitleShapeStyle(PptxTheme theme, PptxColorMap colorMap, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        return sceneChart is null
            ? ReadChartShapeStyle(chartXml.Descendants(ChartNamespace + "title").FirstOrDefault()?.Element(ChartNamespace + "spPr"), theme, colorMap)
            : ToChartShapeStyle(sceneChart.Title.ShapeStyle);
    }

    private static IReadOnlyList<PdfFontResource> RenderManualChartAxisTitles(
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PresentationFontResolver fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        string? chartPartName,
        int slideIndex,
        bool emitDefaultLayoutDiagnostics = true)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        var fonts = new List<PdfFontResource>();
        if (sceneChart is not null)
        {
            foreach (PptxSceneChartAxis axis in sceneChart.Axes)
            {
                if (!string.IsNullOrWhiteSpace(axis.Title.Text) && !axis.Title.Layout.HasLayout)
                {
                    if (emitDefaultLayoutDiagnostics)
                    {
                        EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART_AXIS_TITLE_LAYOUT", OoxPdfSeverity.Warning, "Default-placement chart axis titles are not rendered until the Office axis-title layout model is implemented.", chartPartName, slideIndex, "Ignored");
                    }

                    continue;
                }

                fonts.AddRange(RenderManualChartAxisTitle(
                    theme,
                    graphics,
                    frame,
                    axis.Title.Text,
                    ToChartTextRuns(axis.Title.TextRuns),
                    axis.Title.Layout,
                    ToChartShapeStyle(axis.Title.ShapeStyle),
                    axis.Title.TextBodyProperties,
                    ToChartTextStyleOverride(sceneChart.TextStyle),
                    ReadChartStyleRoleTextStyle(sceneChart.StylePart, GetChartAxisStyleRole(axis.AxisKind)),
                    ToChartTextStyleOverride(axis.Title.TextStyle),
                    fontResolver));
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

            PptxSceneChartManualLayout layout = ReadManualLayout(title);
            if (!layout.HasLayout)
            {
                if (emitDefaultLayoutDiagnostics)
                {
                    EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART_AXIS_TITLE_LAYOUT", OoxPdfSeverity.Warning, "Default-placement chart axis titles are not rendered until the Office axis-title layout model is implemented.", chartPartName, slideIndex, "Ignored");
                }

                continue;
            }

            fonts.AddRange(RenderManualChartAxisTitle(
                theme,
                graphics,
                frame,
                text,
                ToChartTextRuns(PptxSceneBuilder.ReadChartTextRuns(title.Element(ChartNamespace + "tx"), theme, colorMap)),
                layout,
                ReadChartShapeStyle(title.Element(ChartNamespace + "spPr"), theme, colorMap),
                ReadChartTextBodyProperties(title),
                ReadChartTextStyleFromTxPr(chartXml.Root, theme, colorMap),
                ChartTextStyleOverride.Empty,
                ReadChartTextStyleFromTxPr(title, theme, colorMap),
                fontResolver));
        }

        return fonts;
    }

    private static IReadOnlyList<PdfFontResource> RenderDefaultChartAxisTitles(
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ChartLayout layout,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PresentationFontResolver fontResolver)
    {
        var fonts = new List<PdfFontResource>();
        if (sceneChart is not null)
        {
            ChartTextStyleOverride chartTextStyle = ToChartTextStyleOverride(sceneChart.TextStyle);
            foreach (PptxSceneChartAxis axis in sceneChart.Axes)
            {
                if (axis.IsDeleted == true ||
                    string.IsNullOrWhiteSpace(axis.Title.Text) ||
                    axis.Title.Layout.HasLayout)
                {
                    continue;
                }

                fonts.AddRange(RenderDefaultChartAxisTitle(
                    theme,
                    graphics,
                    layout,
                    axis.Title.Text,
                    ToChartTextRuns(axis.Title.TextRuns),
                    axis.AxisKind,
                    axis.PositionKind,
                    ToChartShapeStyle(axis.Title.ShapeStyle),
                    axis.Title.TextBodyProperties,
                    chartTextStyle,
                    ReadChartStyleRoleTextStyle(sceneChart.StylePart, GetChartAxisStyleRole(axis.AxisKind)),
                    ToChartTextStyleOverride(axis.Title.TextStyle),
                    fontResolver));
            }

            return fonts;
        }

        foreach (XElement axis in ReadChartAxisElements(chartXml))
        {
            XElement? title = axis.Element(ChartNamespace + "title");
            string? text = ReadChartText(title?.Element(ChartNamespace + "tx"));
            if (title is null ||
                string.IsNullOrWhiteSpace(text) ||
                ReadManualLayout(title).HasLayout)
            {
                continue;
            }

            fonts.AddRange(RenderDefaultChartAxisTitle(
                theme,
                graphics,
                layout,
                text,
                ToChartTextRuns(PptxSceneBuilder.ReadChartTextRuns(title.Element(ChartNamespace + "tx"), theme, colorMap)),
                PptxSceneBuilder.ParseChartAxisKind(axis.Name.LocalName),
                PptxSceneBuilder.ParseChartAxisPosition((string?)axis.Element(ChartNamespace + "axPos")?.Attribute("val")),
                ReadChartShapeStyle(title.Element(ChartNamespace + "spPr"), theme, colorMap),
                ReadChartTextBodyProperties(title),
                ReadChartTextStyleFromTxPr(chartXml.Root, theme, colorMap),
                ChartTextStyleOverride.Empty,
                ReadChartTextStyleFromTxPr(title, theme, colorMap),
                fontResolver));
        }

        return fonts;
    }

    private static void EmitUnrenderedDefaultChartAxisTitleDiagnostics(
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        string? chartPartName,
        int slideIndex)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        if (sceneChart is not null)
        {
            foreach (PptxSceneChartAxis axis in sceneChart.Axes)
            {
                if (axis.IsDeleted == true ||
                    string.IsNullOrWhiteSpace(axis.Title.Text) ||
                    axis.Title.Layout.HasLayout ||
                    IsRenderableDefaultChartAxisTitle(axis.AxisKind, axis.PositionKind))
                {
                    continue;
                }

                EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART_AXIS_TITLE_AXIS_POSITION", OoxPdfSeverity.Warning, "Default-placement chart axis title has an unsupported or missing axis kind/position.", chartPartName, slideIndex, "Ignored");
            }

            return;
        }

        foreach (XElement axis in ReadChartAxisElements(chartXml))
        {
            XElement? title = axis.Element(ChartNamespace + "title");
            string? text = ReadChartText(title?.Element(ChartNamespace + "tx"));
            if (title is null ||
                string.IsNullOrWhiteSpace(text) ||
                ReadManualLayout(title).HasLayout ||
                IsRenderableDefaultChartAxisTitle(
                    PptxSceneBuilder.ParseChartAxisKind(axis.Name.LocalName),
                    PptxSceneBuilder.ParseChartAxisPosition((string?)axis.Element(ChartNamespace + "axPos")?.Attribute("val"))))
            {
                continue;
            }

            EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART_AXIS_TITLE_AXIS_POSITION", OoxPdfSeverity.Warning, "Default-placement chart axis title has an unsupported or missing axis kind/position.", chartPartName, slideIndex, "Ignored");
        }
    }

    private static IReadOnlyList<PdfFontResource> RenderDefaultChartAxisTitle(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartLayout layout,
        string? text,
        IReadOnlyList<ChartTextRunOverride> textRuns,
        PptxSceneChartAxisKind axisKind,
        PptxSceneChartAxisPosition positionKind,
        ChartShapeStyle shapeStyle,
        PptxSceneChartTextBodyProperties textBodyProperties,
        ChartTextStyleOverride chartTextStyle,
        ChartTextStyleOverride chartStyleRoleTextStyle,
        ChartTextStyleOverride titleTextStyle,
        PresentationFontResolver fontResolver)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return [];
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize);
        style = MergeChartTextStyle(style, chartTextStyle);
        style = MergeChartTextStyle(style, chartStyleRoleTextStyle);
        style = MergeChartTextStyle(style, titleTextStyle);
        string trimmed = text.Trim();
        double titleHeight = style.FontSize * PptxChartMetricRules.TitleHeightFactor;
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        double textWidth = Math.Max(style.FontSize, textMeasurer.Measure(trimmed, style));
        if (!IsRenderableDefaultChartAxisTitle(axisKind, positionKind))
        {
            return [];
        }

        return positionKind switch
        {
            PptxSceneChartAxisPosition.Bottom or PptxSceneChartAxisPosition.Top =>
                RenderDefaultHorizontalChartAxisTitle(graphics, layout, trimmed, textRuns, style, shapeStyle, titleHeight, textWidth, positionKind, fontResolver),
            PptxSceneChartAxisPosition.Left or PptxSceneChartAxisPosition.Right =>
                RenderDefaultVerticalChartAxisTitle(graphics, layout, trimmed, textRuns, style, shapeStyle, titleHeight, textWidth, positionKind, fontResolver),
            _ => []
        };
    }

    private static bool IsRenderableDefaultChartAxisTitle(PptxSceneChartAxisKind axisKind, PptxSceneChartAxisPosition positionKind)
    {
        return positionKind switch
        {
            PptxSceneChartAxisPosition.Bottom or PptxSceneChartAxisPosition.Top =>
                axisKind is PptxSceneChartAxisKind.Category or PptxSceneChartAxisKind.Date or PptxSceneChartAxisKind.Series or PptxSceneChartAxisKind.Value,
            PptxSceneChartAxisPosition.Left or PptxSceneChartAxisPosition.Right =>
                axisKind is PptxSceneChartAxisKind.Value or PptxSceneChartAxisKind.Category or PptxSceneChartAxisKind.Date,
            _ => false
        };
    }

    private static string GetChartAxisStyleRole(PptxSceneChartAxisKind axisKind)
    {
        return axisKind switch
        {
            PptxSceneChartAxisKind.Value => "valueAxis",
            PptxSceneChartAxisKind.Category or PptxSceneChartAxisKind.Date or PptxSceneChartAxisKind.Series => "categoryAxis",
            _ => string.Empty
        };
    }

    private static IReadOnlyList<PdfFontResource> RenderDefaultHorizontalChartAxisTitle(
        PdfGraphicsBuilder graphics,
        ChartLayout layout,
        string text,
        IReadOnlyList<ChartTextRunOverride> textRuns,
        ChartTextStyle style,
        ChartShapeStyle shapeStyle,
        double titleHeight,
        double textWidth,
        PptxSceneChartAxisPosition positionKind,
        PresentationFontResolver? fontResolver = null)
    {
        ChartFrameBox frame = layout.Frame;
        ChartPlotBox plotBox = layout.PlotBox;
        double titleX = plotBox.X + plotBox.Width / 2d - textWidth / 2d;
        double baselineY;
        double boxY;
        if (positionKind == PptxSceneChartAxisPosition.Top)
        {
            double topReserve = Math.Max(0d, frame.Y + frame.Height - plotBox.Y - plotBox.Height);
            baselineY = plotBox.Y + plotBox.Height + topReserve * PptxChartMetricRules.DefaultAxisTitleTopBandBaselineRatio;
            boxY = Math.Max(plotBox.Y + plotBox.Height, baselineY - titleHeight);
        }
        else
        {
            double bottomReserve = Math.Max(0d, plotBox.Y - frame.Y);
            baselineY = frame.Y + bottomReserve * PptxChartMetricRules.DefaultAxisTitleBandBaselineRatio;
            boxY = Math.Max(frame.Y, baselineY - titleHeight);
        }

        double boxX = Math.Max(frame.X, titleX - style.FontSize * 0.5d);
        double boxWidth = Math.Min(frame.X + frame.Width - boxX, textWidth + style.FontSize);
        RenderChartShapeStyle(graphics, boxX, boxY, boxWidth, titleHeight, shapeStyle);
        var runs = new List<TextRun>();
        AddChartRichTextRuns(
            runs,
            textRuns,
            text,
            titleX,
            baselineY,
            textWidth,
            titleHeight,
            frame.X,
            frame.Y,
            frame.Width,
            frame.Height,
            style,
            TextAlignment.Center,
            fontResolver);
        return RenderTextRuns(runs, graphics, "CAT", fontResolver);
    }

    private static IReadOnlyList<PdfFontResource> RenderDefaultVerticalChartAxisTitle(
        PdfGraphicsBuilder graphics,
        ChartLayout layout,
        string text,
        IReadOnlyList<ChartTextRunOverride> textRuns,
        ChartTextStyle style,
        ChartShapeStyle shapeStyle,
        double titleHeight,
        double textWidth,
        PptxSceneChartAxisPosition positionKind,
        PresentationFontResolver? fontResolver = null)
    {
        ChartFrameBox frame = layout.Frame;
        ChartPlotBox plotBox = layout.PlotBox;
        bool rightSide = positionKind == PptxSceneChartAxisPosition.Right;
        double sideReserve = rightSide
            ? Math.Max(0d, frame.X + frame.Width - plotBox.X - plotBox.Width)
            : Math.Max(0d, plotBox.X - frame.X);
        double sideBaselineRatio = rightSide
            ? PptxChartMetricRules.DefaultAxisTitleRightSideBaselineRatio
            : PptxChartMetricRules.DefaultAxisTitleLeftSideBaselineRatio;
        double baselineX = rightSide
            ? plotBox.X + plotBox.Width + sideReserve * sideBaselineRatio
            : frame.X + sideReserve * sideBaselineRatio;
        double baselineY = plotBox.Y + plotBox.Height / 2d
            - textWidth * PptxChartMetricRules.DefaultAxisTitleSideBaselineRatio;
        double boxX = Math.Max(frame.X, Math.Min(frame.X + frame.Width - titleHeight, baselineX - titleHeight * 0.5d));
        double boxY = Math.Max(frame.Y, Math.Min(frame.Y + frame.Height - textWidth, baselineY));
        RenderChartShapeStyle(graphics, boxX, boxY, titleHeight, textWidth, shapeStyle);
        var runs = new List<TextRun>();
        AddChartRichTextRuns(
            runs,
            textRuns,
            text,
            baselineX,
            baselineY,
            textWidth,
            titleHeight,
            frame.X,
            frame.Y,
            frame.Width,
            frame.Height,
            style,
            TextAlignment.Left,
            fontResolver);
        double rotation = -90d;
        for (int i = 0; i < runs.Count; i++)
        {
            TextRun run = runs[i];
            runs[i] = run with
            {
                RotationDegrees = rotation,
                RotationCenterX = run.X,
                RotationCenterY = run.Y
            };
        }

        return RenderTextRuns(runs, graphics, "CAT", fontResolver);
    }

    private static IReadOnlyList<PdfFontResource> RenderManualChartAxisTitle(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartFrameBox frame,
        string? text,
        IReadOnlyList<ChartTextRunOverride> textRuns,
        PptxSceneChartManualLayout layout,
        ChartShapeStyle shapeStyle,
        PptxSceneChartTextBodyProperties textBodyProperties,
        ChartTextStyleOverride chartTextStyle,
        ChartTextStyleOverride chartStyleRoleTextStyle,
        ChartTextStyleOverride titleTextStyle,
        PresentationFontResolver? fontResolver = null)
    {
        if (string.IsNullOrWhiteSpace(text) ||
            !TryBuildManualLayoutBox(layout, frame, new ChartLayoutBox(frame.X, frame.Y, frame.Width, frame.Height), out ChartLayoutBox titleBox))
        {
            return [];
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize);
        style = MergeChartTextStyle(style, chartTextStyle);
        style = MergeChartTextStyle(style, chartStyleRoleTextStyle);
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
            TextAlignment.Center,
            fontResolver);
        return RenderTextRuns(runs, graphics, "CAT", fontResolver);
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

    private static ChartAxisTitleReserveSides ReadSceneOrXmlDefaultAxisTitleReserveSides(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        var reserveSides = new ChartAxisTitleReserveSides();
        if (sceneChart is not null)
        {
            foreach (PptxSceneChartAxis axis in sceneChart.Axes)
            {
                if (axis.IsDeleted == true ||
                    string.IsNullOrWhiteSpace(axis.Title.Text) ||
                    axis.Title.Layout.HasLayout ||
                    !IsRenderableDefaultChartAxisTitle(axis.AxisKind, axis.PositionKind))
                {
                    continue;
                }

                reserveSides = reserveSides.With(axis.PositionKind);
            }

            return reserveSides;
        }

        foreach (XElement axis in ReadChartAxisElements(chartXml))
        {
            XElement? title = axis.Element(ChartNamespace + "title");
            string? text = ReadChartText(title?.Element(ChartNamespace + "tx"));
            PptxSceneChartAxisKind axisKind = PptxSceneBuilder.ParseChartAxisKind(axis.Name.LocalName);
            PptxSceneChartAxisPosition positionKind = PptxSceneBuilder.ParseChartAxisPosition((string?)axis.Element(ChartNamespace + "axPos")?.Attribute("val"));
            if (title is null ||
                string.IsNullOrWhiteSpace(text) ||
                ReadManualLayout(title).HasLayout ||
                !IsRenderableDefaultChartAxisTitle(axisKind, positionKind))
            {
                continue;
            }

            reserveSides = reserveSides.With(positionKind);
        }

        return reserveSides;
    }

    private static bool HasSceneOrXmlPolarChart(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        return sceneChart is not null
            ? sceneChart.Plots.Any(plot =>
                plot.PlotKind is PptxSceneChartPlotKind.Pie or
                    PptxSceneChartPlotKind.Doughnut or
                    PptxSceneChartPlotKind.Radar)
            : HasPolarChart(chartXml);
    }

    private static bool HasPolarChart(XDocument chartXml)
    {
        return ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Pie).Any() ||
            ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Doughnut).Any() ||
            ReadChartPlotElements(chartXml, PptxSceneChartPlotKind.Radar).Any();
    }

    private static double ResolveChartTitleBaselineY(PptxDocument document, PptxTheme theme, PptxColorMap colorMap, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, ChartWorkbookData? workbook, bool plotVisibleOnly, double fallbackBaselineY, double fontSize, PresentationFontResolver? fontResolver)
    {
        XElement? barChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Bar);
        if (barChart is null)
        {
            return fallbackBaselineY;
        }

        PptxSceneChartPlot? barPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bar);
        ChartBarPlotOptions barOptions = ReadSceneOrXmlChartBarOptions(barPlot, barChart, PptxSceneChartGrouping.Clustered);
        bool horizontalBars = barOptions.BarDirection == PptxSceneChartBarDirection.Bar;
        ChartLayout layout = GetBarChartLayout(document, theme, bounds, chartXml, sceneChart, colorMap, barPlot, barChart, barOptions, workbook, plotVisibleOnly, fontResolver);
        ChartPlotBox titleAnchorBox = layout.PlotBox;
        if (horizontalBars && layout.ManualPlotLayoutApplied && IsAutoGeneratedChartTitle(sceneChart))
        {
            ChartFrameBox frame = GetChartFrameBox(document, bounds);
            ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
            titleAnchorBox = GetBarChartPlotLayout(
                theme,
                frame,
                chartXml,
                sceneChart,
                barPlot,
                barChart,
                ReadSceneOrXmlChartTitleText(sceneChart, chartXml),
                legend,
                barOptions,
                workbook,
                plotVisibleOnly,
                ignoreManualPlotLayout: true,
                fontResolver: fontResolver).PlotBox;
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

    private static ChartTextStyle ReadSceneOrXmlChartTitleTextStyle(PptxTheme theme, PptxColorMap colorMap, PptxSceneChart? sceneChart, XDocument chartXml, bool isAutoTitle)
    {
        if (sceneChart is null)
        {
            XElement? title = chartXml.Descendants(ChartNamespace + "title").FirstOrDefault();
            return ResolveAutoChartTitleTextStyle(
                ReadChartTextStyle(theme, colorMap, chartXml, title, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize),
                ChartTextStyleOverride.Empty,
                isAutoTitle);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, sceneChart.ColorMap, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize);
        style = MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.TextStyle));
        style = MergeChartTextStyle(style, ReadChartStyleRoleTextStyle(sceneChart.StylePart, "title"));
        ChartTextStyleOverride titleStyle = ToChartTextStyleOverride(sceneChart.Title.TextStyle);
        return ResolveAutoChartTitleTextStyle(MergeChartTextStyle(style, titleStyle), titleStyle, isAutoTitle);
    }

    private static ChartTextStyle ResolveAutoChartTitleTextStyle(ChartTextStyle style, ChartTextStyleOverride titleStyle, bool isAutoTitle)
    {
        if (!isAutoTitle)
        {
            return style;
        }

        return style with
        {
            FontSize = titleStyle.FontSize is null
                ? style.FontSize * PptxChartMetricRules.AutoTitleFontScale
                : style.FontSize,
            Bold = true
        };
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

    private static PptxSceneChartTextBodyProperties ReadSceneOrXmlChartTitleTextBodyProperties(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            return sceneChart.Title.TextBodyProperties;
        }

        XElement? title = chartXml.Descendants(ChartNamespace + "title").FirstOrDefault();
        return ReadChartTextBodyProperties(title);
    }

    private static IReadOnlyList<ChartTextRunOverride> ReadSceneOrXmlChartTitleTextRuns(PptxTheme theme, PptxColorMap colorMap, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            if (sceneChart.Title.IsAutoGenerated)
            {
                return [];
            }

            return ToChartTextRuns(sceneChart.Title.TextRuns);
        }

        XElement? titleText = chartXml
            .Descendants(ChartNamespace + "title")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "tx");
        return ToChartTextRuns(PptxSceneBuilder.ReadChartTextRuns(titleText, theme, colorMap));
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

    private static ChartIndexedTextVector ReadChartCategoryLabelVector(XElement chartElement, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true)
    {
        XElement? categories = chartElement
            .Element(ChartNamespace + "ser")
            ?.Element(ChartNamespace + "cat");
        if (categories is null)
        {
            return default;
        }

        XElement? cache = categories
            .Descendants(ChartNamespace + "strCache")
            .Concat(categories.Descendants(ChartNamespace + "strLit"))
            .Concat(categories.Descendants(ChartNamespace + "numCache"))
            .Concat(categories.Descendants(ChartNamespace + "numLit"))
            .Concat(categories.Descendants(ChartNamespace + "multiLvlStrCache"))
            .FirstOrDefault();
        PptxSceneChartDataSource source = ReadChartDataSource(categories, "strRef", "numRef", "multiLvlStrRef");
        if (cache is null)
        {
            return new ChartIndexedTextVector(
                [],
                null,
                [],
                source.Formula,
                source,
                ReadWorkbookTextPoints(workbook, source),
                plotVisibleOnly);
        }

        IEnumerable<XElement> pointElements = cache.Name.LocalName == "multiLvlStrCache"
            ? cache.Descendants(ChartNamespace + "pt")
            : cache.Elements(ChartNamespace + "pt");
        ChartIndexedTextPoint[] points = ReadChartIndexedTextPoints(pointElements);
        IReadOnlyList<IReadOnlyList<ChartIndexedTextPoint>> levels = cache.Name.LocalName == "multiLvlStrCache"
            ? cache
                .Elements(ChartNamespace + "lvl")
                .Select(level => ReadChartIndexedTextPoints(level.Elements(ChartNamespace + "pt")))
                .Where(level => level.Length != 0)
                .ToArray()
            : [];
        return new ChartIndexedTextVector(
            points,
            ReadChartPointCount(cache) ?? InferPointCount(points),
            levels,
            source.Formula,
            source,
            ReadWorkbookTextPoints(workbook, source),
            plotVisibleOnly);
    }

    private static ChartIndexedTextPoint[] ReadChartIndexedTextPoints(IEnumerable<XElement> sourcePoints)
    {
        return sourcePoints
            .Select((point, fallbackIndex) =>
            {
                XElement? valueElement = point.Element(ChartNamespace + "v");
                string text = valueElement?.Value.Trim() ?? string.Empty;
                int? parsedIndex = ReadChartCachePointIndex(point);
                return new ChartIndexedTextPoint(
                    parsedIndex ?? fallbackIndex,
                    parsedIndex is not null ? ChartPointIndexSource.OoxmlIndex : ChartPointIndexSource.OrdinalFallback,
                    text,
                    valueElement is not null,
                    default);
            })
            .ToArray();
    }

    private static IReadOnlyList<ChartLegendEntry> BuildFillLegendEntries(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?>? seriesStrokes = null, int paletteOffset = 0, ChartWorkbookData? workbook = null)
    {
        IReadOnlyList<ChartSeriesNameRecord> names = ReadSceneOrXmlChartSeriesNameRecords(plot, chartElement, workbook);
        var entries = new List<ChartLegendEntry>(names.Count);
        for (int i = 0; i < names.Count; i++)
        {
            ChartSeriesFill fill = i < seriesFills.Count && seriesFills[i] is { } explicitFill
                ? explicitFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, i + paletteOffset), 1d);
            ChartSeriesStroke? stroke = seriesStrokes is not null && i < seriesStrokes.Count
                ? seriesStrokes[i]
                : null;
            if (stroke is { } explicitStroke &&
                PptxSceneBuilder.ParseChartPlotKind(chartElement.Name.LocalName) == PptxSceneChartPlotKind.Bar &&
                Math.Abs(explicitStroke.Width - ChartSeriesInheritedStrokeWidth) <= 0.001d)
            {
                stroke = explicitStroke with { Width = ChartFilledSeriesInheritedStrokeWidth };
            }

            entries.Add(new ChartLegendEntry(names[i].ActiveName, fill, stroke, null, names[i]));
        }

        return entries;
    }

    private static IReadOnlyList<ChartLegendEntry> BuildCategoryFillLegendEntries(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true)
    {
        ChartIndexedTextVector labels = ReadSceneOrXmlCategoryLabelVector(plot, chartElement, workbook, plotVisibleOnly);
        IReadOnlyList<ChartIndexedTextPoint> points = labels.DensePoints()
            .Where(point => point is not null)
            .Select(point => point!.Value)
            .ToArray();
        var entries = new List<ChartLegendEntry>(points.Count);
        foreach (ChartIndexedTextPoint point in points.OrderBy(point => point.Index))
        {
            if (!point.HasText || string.IsNullOrWhiteSpace(point.Text))
            {
                continue;
            }

            ChartSeriesFill fill = pointFills.TryGetValue(point.Index, out ChartSeriesFill pointFill)
                ? pointFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, point.Index), 1d);
            entries.Add(new ChartLegendEntry(point.Text, fill, null, null, null));
        }

        return entries;
    }

    private static IReadOnlyList<ChartLegendEntry> BuildStrokeLegendEntries(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle>? markerStyles = null, bool reverseOrder = false, ChartWorkbookData? workbook = null)
    {
        IReadOnlyList<ChartSeriesNameRecord> names = ReadSceneOrXmlChartSeriesNameRecords(plot, chartElement, workbook);
        var entries = new List<ChartLegendEntry>(names.Count);
        for (int i = 0; i < names.Count; i++)
        {
            ChartMarkerStyle? marker = markerStyles is not null && i < markerStyles.Count
                ? markerStyles[i]
                : null;
            entries.Add(new ChartLegendEntry(names[i].ActiveName, null, ChartSeriesStrokeColor(theme, colorMap, chartPalette, i, seriesStrokes, ChartLineDefaultStrokeWidth), marker, names[i]));
        }

        if (reverseOrder)
        {
            entries.Reverse();
        }

        return entries;
    }

    private static IReadOnlyList<ChartSeriesNameRecord> ReadSceneOrXmlChartSeriesNameRecords(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook = null)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select((series, index) =>
                {
                    IReadOnlyList<ChartIndexedTextPoint> workbookPoints = ReadWorkbookTextPoints(workbook, series.DataSources.Name);
                    if (!string.IsNullOrWhiteSpace(series.Name))
                    {
                        return new ChartSeriesNameRecord(series.Name.Trim(), series.Name.Trim(), series.DataSources.Name, workbookPoints);
                    }

                    string? workbookName = workbookPoints
                        .Where(value => value.HasText)
                        .Select(value => value.Text)
                        .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
                    return new ChartSeriesNameRecord(
                        string.IsNullOrWhiteSpace(workbookName) ? $"Series {index + 1}" : workbookName.Trim(),
                        string.Empty,
                        series.DataSources.Name,
                        workbookPoints);
                })
                .ToArray();
        }

        return ReadChartSeriesNameRecords(chartElement, workbook);
    }

    private static IReadOnlyList<ChartSeriesNameRecord> ReadChartSeriesNameRecords(XElement chartElement, ChartWorkbookData? workbook = null)
    {
        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select((series, index) =>
            {
                string cacheName = ReadChartSeriesName(series)?.Trim() ?? string.Empty;
                string activeName = string.IsNullOrWhiteSpace(cacheName) ? $"Series {index + 1}" : cacheName;
                PptxSceneChartDataSource source = ReadChartSeriesNameDataSource(series);
                return new ChartSeriesNameRecord(activeName, cacheName, source, ReadWorkbookTextPoints(workbook, source));
            })
            .ToArray();
    }

    private static string? ReadChartSeriesName(XElement series)
    {
        return ReadChartText(series.Element(ChartNamespace + "tx"));
    }

    private static PptxSceneChartDataSource ReadChartSeriesNameDataSource(XElement series)
    {
        return ReadChartDataSource(series.Element(ChartNamespace + "tx"), "strRef", "numRef");
    }

    private static PptxSceneChartDataSource ReadChartDataSource(XElement? container, params string[] referenceKinds)
    {
        if (container is null)
        {
            return default;
        }

        foreach (string referenceKind in referenceKinds)
        {
            XElement? reference = container.Element(ChartNamespace + referenceKind);
            if (reference is null)
            {
                continue;
            }

            XElement? cache = reference
                .Elements()
                .FirstOrDefault(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Cache", StringComparison.Ordinal));
            bool hasCachedPoints = cache?
                .Descendants(ChartNamespace + "pt")
                .Any() == true;
            return new PptxSceneChartDataSource(
                (string?)reference.Element(ChartNamespace + "f"),
                PptxSceneBuilder.ParseChartDataSourceReferenceKind(referenceKind),
                referenceKind,
                PptxSceneBuilder.ParseChartDataSourceCacheKind(cache?.Name.LocalName),
                cache?.Name.LocalName ?? string.Empty,
                hasCachedPoints);
        }

        return default;
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

    private static ChartLegendLayout ReadChartLegendLayout(PptxTheme theme, PptxColorMap colorMap, XDocument chartXml)
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
        return new ChartLegendLayout(ResolveChartLegendPosition(PptxSceneBuilder.ParseChartLegendPosition(position)), position, overlay, Visible: true, ReadManualLayout(legend), ReadChartTextBodyProperties(legend), ReadChartShapeStyle(legend.Element(ChartNamespace + "spPr"), theme, colorMap));
    }

    private static ChartLegendLayout ReadSceneOrXmlChartLegendLayout(PptxTheme theme, PptxColorMap colorMap, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is null)
        {
            return ReadChartLegendLayout(theme, colorMap, chartXml);
        }

        if (!sceneChart.Legend.IsDefined)
        {
            return ChartLegendLayout.Hidden;
        }

        return sceneChart.Legend.IsDeleted != true
            ? new ChartLegendLayout(ResolveChartLegendPosition(sceneChart.Legend.PositionKind), sceneChart.Legend.Position, sceneChart.Legend.Overlay == true, Visible: true, sceneChart.Legend.Layout, sceneChart.Legend.TextBodyProperties, ToChartShapeStyle(sceneChart.Legend.ShapeStyle))
            : ChartLegendLayout.Hidden;
    }

    private static PptxSceneChartLegendPosition ResolveChartLegendPosition(PptxSceneChartLegendPosition position)
    {
        return position == PptxSceneChartLegendPosition.Unknown
            ? PptxSceneChartLegendPosition.Right
            : position;
    }

    private static ChartTextStyle ReadSceneOrXmlChartLegendTextStyle(PptxTheme theme, PptxColorMap colorMap, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is null)
        {
            XElement? legend = chartXml.Descendants(ChartNamespace + "legend").FirstOrDefault();
            return ReadChartTextStyle(theme, colorMap, chartXml, legend, fallbackFontSize: PptxChartMetricRules.LegendFallbackFontSize);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, sceneChart.ColorMap, fallbackFontSize: PptxChartMetricRules.LegendFallbackFontSize);
        style = MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.TextStyle));
        style = MergeChartTextStyle(style, ReadChartStyleRoleTextStyle(sceneChart.StylePart, "legend"));
        return MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.Legend.TextStyle));
    }

    private static IReadOnlyList<PdfFontResource> RenderChartLegend(PdfGraphicsBuilder graphics, ChartFrameBox frame, ChartPlotBox plotBox, IReadOnlyList<ChartLegendEntry> entries, ChartLegendLayout layout, ChartTextStyle style, PresentationFontResolver? fontResolver = null)
    {
        if (!layout.Visible || entries.Count == 0)
        {
            return [];
        }

        var textMeasurer = new ChartTextMeasurer(fontResolver);
        ChartLegendBox legendBox = ResolveChartLegendBox(frame, plotBox, entries, layout, style, textMeasurer);

        RenderChartShapeStyle(graphics, legendBox.X, legendBox.ClipY, legendBox.Width, legendBox.ClipHeight, layout.ShapeStyle);

        var runs = new List<TextRun>(entries.Count);
        for (int i = 0; i < entries.Count; i++)
        {
            ChartLegendEntry entry = entries[i];
            double entryX = legendBox.Horizontal ? GetPackedHorizontalLegendEntryX(entries, style, textMeasurer, legendBox.MarkerSize, legendBox.X, i) : legendBox.X;
            double entryWidth = legendBox.Horizontal ? GetPackedHorizontalLegendEntryWidth(entries[i].Name, style, textMeasurer, legendBox.MarkerSize) : legendBox.Width;
            double y = legendBox.Horizontal ? legendBox.FirstY : legendBox.FirstY - i * legendBox.LineHeight;
            double markerBaselineFactor = legendBox.Horizontal || legendBox.SideStrokeLegend
                ? PptxChartMetricRules.LegendHorizontalMarkerBaselineFactor
                : entry.Fill is not null
                    ? PptxChartMetricRules.LegendSideFillMarkerBaselineFactor
                    : PptxChartMetricRules.LegendMarkerBaselineFactor;
            double markerY = y + legendBox.LineHeight * markerBaselineFactor;
            if (entry.Fill is { } fill)
            {
                FillChartRectangle(graphics, entryX, markerY, legendBox.MarkerSize, legendBox.MarkerSize, fill);
                if (entry.Stroke is { } fillStroke && fillStroke.Alpha > 0d)
                {
                    if (fillStroke.Alpha < 1d)
                    {
                        graphics.SaveState();
                        graphics.SetAlpha(1d, fillStroke.Alpha);
                    }

                    SetChartStroke(graphics, ResolveFilledLegendKeyStroke(fillStroke));
                    graphics.StrokeRectangle(entryX, markerY, legendBox.MarkerSize, legendBox.MarkerSize);
                    if (fillStroke.Alpha < 1d)
                    {
                        graphics.RestoreState();
                    }
                }
            }
            else if (entry.Stroke is { } stroke)
            {
                double lineY = markerY + legendBox.MarkerSize / 2d;
                SetChartStroke(graphics, stroke);
                graphics.StrokeLine(entryX, lineY, entryX + legendBox.MarkerWidth, lineY);
                if (entry.Marker is { } marker)
                {
                    DrawChartMarker(graphics, entryX + legendBox.MarkerWidth / 2d, lineY, marker, stroke.Color, stroke.Color);
                }
            }

            runs.Add(new TextRun(
                entry.Name,
                entryX + legendBox.MarkerWidth + legendBox.TextGap,
                y,
                Math.Max(1d, entryWidth - legendBox.MarkerWidth - legendBox.TextGap),
                legendBox.LineHeight,
                legendBox.X,
                legendBox.ClipY,
                legendBox.Width,
                legendBox.ClipHeight,
                style.FontSize,
                0d,
                0d,
                style.Color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: style.Underline,
                Strike: style.Strike,
                KerningEnabled: true,
                TextAlignment.Left,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false));
        }

        return RenderTextRuns(runs, graphics, "CL", fontResolver);
    }

    private static ChartLegendBox ResolveChartLegendBox(ChartFrameBox frame, ChartPlotBox plotBox, IReadOnlyList<ChartLegendEntry> entries, ChartLegendLayout layout, ChartTextStyle style, ChartTextMeasurer textMeasurer)
    {
        double fontSize = style.FontSize;
        double markerSize = fontSize * PptxChartMetricRules.LegendMarkerSizeFactor;
        bool horizontal = IsHorizontalLegendPosition(layout.PositionKind);
        bool sideStrokeLegend = !horizontal && entries.All(entry => entry.Stroke is not null && entry.Fill is null);
        bool fillLegendInFullFrame = !sideStrokeLegend && IsSameChartBox(plotBox, frame);
        bool sideFillLegend = !horizontal && !fillLegendInFullFrame && !sideStrokeLegend && entries.Any(entry => entry.Fill is not null);
        bool sideFillLegendInFullFrame = !horizontal && fillLegendInFullFrame;
        double lineHeight = fontSize * (sideStrokeLegend
            || sideFillLegend
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
            ? Math.Min(plotBox.Width, GetPackedHorizontalLegendWidth(entries, style, textMeasurer, markerSize))
            : Math.Max(
                sideStrokeLegend || sideFillLegendInFullFrame
                    ? 0d
                    : Math.Max(PptxChartMetricRules.LegendMinimumSideWidth, plotBox.Width * PptxChartMetricRules.LegendSideWidthRatio),
                GetSideLegendContentWidth(entries, style, textMeasurer, markerWidth, textGap));
        double x = layout.PositionKind switch
        {
            PptxSceneChartLegendPosition.Left when sideFillLegendInFullFrame => frame.X + frame.Width * PptxChartMetricRules.LegendFullFrameSideInsetRatio,
            PptxSceneChartLegendPosition.Left => Math.Max(0d, plotBox.X - width - sideGap),
            _ when horizontal => plotBox.X + (plotBox.Width - width) / 2d,
            _ when sideFillLegendInFullFrame => frame.X + frame.Width - width,
            _ when sideFillLegend => plotBox.X + plotBox.Width + sideGap + frame.Width * PptxChartMetricRules.LegendSideFillContentBoxReservedBandOffsetFactor,
            _ when !sideStrokeLegend => plotBox.X + plotBox.Width + sideGap + frame.Width * PptxChartMetricRules.LegendSideFillReservedBandOffsetFactor,
            _ => plotBox.X + plotBox.Width + sideGap
        };
        double firstY = layout.PositionKind switch
        {
            PptxSceneChartLegendPosition.Bottom when fillLegendInFullFrame => frame.Y + lineHeight * PptxChartMetricRules.LegendFullFrameBottomBaselineFactor,
            PptxSceneChartLegendPosition.Top when fillLegendInFullFrame => frame.Y + frame.Height - lineHeight * PptxChartMetricRules.LegendFullFrameTopBaselineFactor,
            PptxSceneChartLegendPosition.Bottom => Math.Max(0d, plotBox.Y - lineHeight * PptxChartMetricRules.LegendBottomOffsetFactor),
            PptxSceneChartLegendPosition.Top => plotBox.Y + plotBox.Height + lineHeight * PptxChartMetricRules.LegendTopOffsetFactor,
            _ when sideStrokeLegend => plotBox.Y + plotBox.Height / 2d -
                fontSize * GetLegendSideStrokeBaselineCenterOffsetFactor(entries) +
                (entries.Count - 1) * lineHeight / 2d,
            _ when sideFillLegend => frame.Y + frame.Height / 2d -
                fontSize * PptxChartMetricRules.LegendSideFillBaselineCenterOffsetFactor +
                (entries.Count - 1) * lineHeight / 2d,
            _ when !sideStrokeLegend && !horizontal => frame.Y + frame.Height / 2d +
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

        return new ChartLegendBox(x, clipY, width, clipHeight, firstY, lineHeight, markerSize, markerWidth, textGap, horizontal, sideStrokeLegend);
    }

    private static double GetLegendSideStrokeBaselineCenterOffsetFactor(IReadOnlyList<ChartLegendEntry> entries)
    {
        return entries.Any(entry => entry.Marker is { } marker &&
            marker.Size >= PptxChartMarkerMetricRules.StyledLineChartMarkerSize - PptxChartMetricRules.AxisValueEpsilon)
            ? PptxChartMetricRules.LegendSideStrokeStyledMarkerBaselineCenterOffsetFactor
            : PptxChartMetricRules.LegendSideStrokeBaselineCenterOffsetFactor;
    }

    private static ChartSeriesStroke ResolveFilledLegendKeyStroke(ChartSeriesStroke stroke)
    {
        return stroke.Width > ChartFilledSeriesInheritedStrokeWidth
            ? stroke with { Width = ChartFilledSeriesInheritedStrokeWidth }
            : stroke;
    }

    private static double GetPackedHorizontalLegendWidth(IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle style, ChartTextMeasurer textMeasurer, double markerSize)
    {
        double width = 0d;
        foreach (ChartLegendEntry entry in entries)
        {
            width += GetPackedHorizontalLegendEntryWidth(entry.Name, style, textMeasurer, markerSize);
        }

        return Math.Max(1d, width);
    }

    private static double GetSideLegendContentWidth(IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle style, ChartTextMeasurer textMeasurer, double markerWidth, double textGap)
    {
        double contentWidth = entries.Count == 0
            ? 0d
            : entries.Max(entry => markerWidth + textGap + textMeasurer.Measure(entry.Name, style));
        return Math.Max(style.FontSize * PptxChartMetricRules.LegendSideFillMinimumWidthFactor, contentWidth);
    }

    private static bool IsSameChartBox(ChartPlotBox plotBox, ChartFrameBox frame)
    {
        const double tolerance = 0.01d;
        return Math.Abs(plotBox.X - frame.X) <= tolerance &&
            Math.Abs(plotBox.Y - frame.Y) <= tolerance &&
            Math.Abs(plotBox.Width - frame.Width) <= tolerance &&
            Math.Abs(plotBox.Height - frame.Height) <= tolerance;
    }

    private static double GetPackedHorizontalLegendEntryX(IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle style, ChartTextMeasurer textMeasurer, double markerSize, double legendX, int entryIndex)
    {
        double x = legendX;
        for (int i = 0; i < entryIndex; i++)
        {
            x += GetPackedHorizontalLegendEntryWidth(entries[i].Name, style, textMeasurer, markerSize);
        }

        return x;
    }

    private static double GetPackedHorizontalLegendEntryWidth(string name, ChartTextStyle style, ChartTextMeasurer textMeasurer, double markerSize)
    {
        return markerSize + PptxChartMetricRules.LegendTextGap + textMeasurer.Measure(name, style);
    }

    private static bool IsOoxmlTrue(string? value)
    {
        return OoxBoolean.IsTrue(value);
    }

    private static bool? ReadOoxmlBooleanAttribute(XElement? element, string name)
    {
        return element?.Attribute(name) is { } attribute
            ? IsOoxmlTrue(attribute.Value)
            : null;
    }

    private static bool IsOoxmlBooleanElementEnabled(XElement? element)
    {
        return OoxBoolean.ParseElement(element);
    }

    private static bool IsOoxmlBooleanElementEnabled(XElement? element, bool defaultValue)
    {
        return OoxBoolean.ParseElement(element, defaultValue);
    }

    private static IReadOnlyList<PdfFontResource> RenderPieDataLabels(PptxTheme theme, PptxColorMap colorMap, PdfGraphicsBuilder graphics, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<ChartIndexedPieSlice> slices, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, string? valueFormatCode, ChartDataLabelOptions labelOptions, ChartIndexedTextVector categoryLabels, IReadOnlyList<ChartSeriesNameRecord> seriesNames, PresentationFontResolver? fontResolver = null)
    {
        if (!labelOptions.HasVisibleContent || slices.Count == 0)
        {
            return [];
        }

        double total = slices.Sum(slice => slice.Value);
        if (total <= 0d)
        {
            return [];
        }

        ChartPlotBox plotBox = layout.PlotBox;
        ChartPolarGeometry geometry = layout.Geometry;
        double labelRadius = geometry.Radius * (holeSize > 0d ? Math.Max(PptxChartMetricRules.PieDataLabelRadiusRatio, (1d + holeSize) / 2d) : PptxChartMetricRules.PieDataLabelRadiusRatio);
        double labelWidth = Math.Max(PptxChartMetricRules.PieDataLabelMinimumWidth, geometry.Radius * PptxChartMetricRules.PieDataLabelWidthRatio);
        var runs = new List<TextRun>(slices.Count);
        double angle = -90d;
        foreach (ChartIndexedPieSlice slice in slices)
        {
            ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(labelOptions, slice.Index);
            if (!effectiveOptions.HasVisibleContent)
            {
                continue;
            }

            ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
            double fontSize = style.FontSize;
            double labelHeight = fontSize * PptxChartMetricRules.PieDataLabelHeightFactor;
            double sweep = slice.Value / total * 360d;
            double mid = (angle + sweep / 2d) * Math.PI / 180d;
            double explosion = pointExplosions.TryGetValue(slice.Index, out double offset) ? Math.Clamp(offset, 0d, 1d) * geometry.Radius * PptxChartMetricRules.PieExplosionLabelRadiusRatio : 0d;
            double labelX = geometry.CenterX + Math.Cos(mid) * (labelRadius + explosion) - labelWidth / 2d;
            double labelY = geometry.CenterY + Math.Sin(mid) * (labelRadius + explosion) - labelHeight / 2d;
            IReadOnlyList<string> labelParts = FormatPieDataLabelParts(slice.Value, total, slice.Index, slice.WorkbookPoint, valueFormatCode, effectiveOptions, categoryLabels, seriesNames);
            string label = JoinChartDataLabelParts(labelParts, effectiveOptions);
            if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
            {
                ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, labelWidth, labelHeight);
                RenderPieDataLabelLeaderLine(graphics, geometry, mid, explosion, labelBox, effectiveOptions);
                RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                double textX = labelBox.X;
                double textWidth = labelBox.Width;
                TextAlignment alignment = TextAlignment.Center;
                if (effectiveOptions.ShowLegendKey)
                {
                    double swatchSize = fontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
                    double swatchGap = fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
                    double swatchY = labelBox.Y + Math.Max(0d, (labelBox.Height - swatchSize) / 2d);
                    ChartSeriesFill fill = pointFills.TryGetValue(slice.Index, out ChartSeriesFill explicitFill)
                        ? explicitFill
                        : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, slice.Index), 1d);
                    FillChartRectangle(graphics, labelBox.X, swatchY, swatchSize, swatchSize, fill);
                    textX += swatchSize + swatchGap;
                    textWidth = Math.Max(1d, textWidth - swatchSize - swatchGap);
                    alignment = TextAlignment.Left;
                }

                if (!string.IsNullOrEmpty(label))
                {
                    AddPolarChartLabelRuns(runs, labelParts, label, effectiveOptions, textX, labelBox.Y, textWidth, labelBox.Height, plotBox, style, alignment, fontResolver);
                }
            }
            angle += sweep;
        }

        return RenderTextRuns(runs, graphics, "CP", fontResolver);
    }

    private static void RenderPieDataLabelLeaderLine(PdfGraphicsBuilder graphics, ChartPolarGeometry geometry, double angleRadians, double explosion, ChartLayoutBox labelBox, ChartDataLabelOptions options)
    {
        if (!options.ShowLeaderLines)
        {
            return;
        }

        ChartSeriesStroke stroke = options.LeaderLines.Stroke ?? ChartDataLabelLeaderLineDefaultStroke;
        if (stroke.Alpha <= 0.001d || stroke.Width <= 0d)
        {
            return;
        }

        double startX = geometry.CenterX + Math.Cos(angleRadians) * (geometry.Radius + explosion);
        double startY = geometry.CenterY + Math.Sin(angleRadians) * (geometry.Radius + explosion);
        double labelCenterX = labelBox.X + labelBox.Width / 2d;
        double labelCenterY = labelBox.Y + labelBox.Height / 2d;
        bool labelIsLeft = labelCenterX < geometry.CenterX;
        double labelEdgeX = labelIsLeft ? labelBox.X + labelBox.Width : labelBox.X;
        double tail = Math.Min(Math.Max(stroke.Width * 2d, 4d), Math.Max(4d, labelBox.Width * 0.18d));
        double elbowX = labelIsLeft ? labelEdgeX + tail : labelEdgeX - tail;

        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        graphics.MoveTo(startX, startY);
        graphics.LineTo(elbowX, labelCenterY);
        graphics.LineTo(labelEdgeX, labelCenterY);
        graphics.StrokeCurrentPath();

        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static IReadOnlyList<PdfFontResource> RenderBarDataLabels(
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<ChartIndexedNumberVector> series,
        IReadOnlyList<RgbColor>? chartPalette,
        ChartValueExtents extents,
        bool horizontalBars,
        bool valueAxisReversed,
        IReadOnlyList<ChartSeriesFill?> seriesFills,
        IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills,
        bool varyColors,
        ChartDataLabelOptions labelOptions,
        IReadOnlyList<ChartDataLabelOptions> seriesLabelOptions,
        ChartIndexedTextVector categoryLabels,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames,
        PresentationFontResolver? fontResolver = null)
    {
        if ((!labelOptions.HasVisibleContent && !seriesLabelOptions.Any(options => options.HasVisibleContent)) || series.Count == 0)
        {
            return [];
        }

        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> densePointSeries = DensifyChartPointSeries(series);
        if (densePointSeries.Count == 0)
        {
            return [];
        }

        int categoryCount = Math.Max(1, densePointSeries.Max(values => values.Count));
        double zeroX = ChartValueToPlotCoordinate(extents, 0d, plotBox.X, plotBox.Width, valueAxisReversed);
        double zeroY = ChartValueToPlotCoordinate(extents, 0d, plotBox.Y, plotBox.Height, valueAxisReversed);
        var runs = new List<TextRun>();
        if (horizontalBars)
        {
            double categoryHeight = plotBox.Height / categoryCount;
            double barSlot = categoryHeight * PptxChartMetricRules.BarDataLabelSlotFillRatio / Math.Max(1, densePointSeries.Count);
            double labelWidth = Math.Max(PptxChartMetricRules.CartesianDataLabelMinimumWidth, plotBox.Width * PptxChartMetricRules.HorizontalBarDataLabelWidthRatio);
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryY = plotBox.Y + category * categoryHeight + categoryHeight * PptxChartMetricRules.BarDataLabelCategoryInsetRatio;
                for (int seriesIndex = 0; seriesIndex < densePointSeries.Count; seriesIndex++)
                {
                    IReadOnlyList<ChartIndexedNumberPoint?> points = densePointSeries[seriesIndex];
                    if (category >= points.Count || points[category]?.Value is not double value)
                    {
                        continue;
                    }

                    double barBase = zeroX;
                    double barEnd = ChartValueToPlotCoordinate(extents, value, plotBox.X, plotBox.Width, valueAxisReversed);
                    ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), category);
                    if (!effectiveOptions.HasVisibleContent)
                    {
                        continue;
                    }

                    ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                    double fontSize = style.FontSize;
                    double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                    double x = ResolveHorizontalBarDataLabelX(effectiveOptions.PositionKind, barBase, barEnd, labelWidth);
                    double y = categoryY + seriesIndex * barSlot + barSlot * PptxChartMetricRules.HorizontalBarDataLabelSlotCenterRatio - labelHeight / 2d;
                    ChartIndexedNumberPoint point = points[category]!.Value;
                    string label = FormatCartesianDataLabel(value, seriesIndex, category, point, series[seriesIndex].WorkbookPointForIndex(point.Index), series[seriesIndex].FormatCode, effectiveOptions, categoryLabels, seriesNames);
                    if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
                    {
                        ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, x, y, labelWidth, labelHeight);
                        RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                        double textX = labelBox.X;
                        double textWidth = labelBox.Width;
                        if (effectiveOptions.ShowLegendKey)
                        {
                            ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, densePointSeries.Count, varyColors, seriesFills, pointFills, value);
                            double legendKeyWidth = RenderFillDataLabelLegendKey(graphics, labelBox, fontSize, fill);
                            textX += legendKeyWidth;
                            textWidth = Math.Max(1d, textWidth - legendKeyWidth);
                        }

                        if (!string.IsNullOrEmpty(label))
                        {
                            AddChartLabelRuns(runs, label, effectiveOptions, textX, labelBox.Y, textWidth, labelBox.Height, plotBox, style, TextAlignment.Left, fontResolver);
                        }
                    }
                }
            }
        }
        else
        {
            double categoryWidth = plotBox.Width / categoryCount;
            double barSlot = categoryWidth * PptxChartMetricRules.BarDataLabelSlotFillRatio / Math.Max(1, densePointSeries.Count);
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryX = plotBox.X + category * categoryWidth + categoryWidth * PptxChartMetricRules.BarDataLabelCategoryInsetRatio;
                for (int seriesIndex = 0; seriesIndex < densePointSeries.Count; seriesIndex++)
                {
                    IReadOnlyList<ChartIndexedNumberPoint?> points = densePointSeries[seriesIndex];
                    if (category >= points.Count || points[category]?.Value is not double value)
                    {
                        continue;
                    }

                    double x = categoryX + seriesIndex * barSlot;
                    double barBase = zeroY;
                    double barEnd = ChartValueToPlotCoordinate(extents, value, plotBox.Y, plotBox.Height, valueAxisReversed);
                    ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), category);
                    if (!effectiveOptions.HasVisibleContent)
                    {
                        continue;
                    }

                    ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                    double fontSize = style.FontSize;
                    double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                    double y = ResolveVerticalBarDataLabelY(effectiveOptions.PositionKind, barBase, barEnd, labelHeight);
                    ChartIndexedNumberPoint point = points[category]!.Value;
                    string label = FormatCartesianDataLabel(value, seriesIndex, category, point, series[seriesIndex].WorkbookPointForIndex(point.Index), series[seriesIndex].FormatCode, effectiveOptions, categoryLabels, seriesNames);
                    if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
                    {
                        double legendKeyWidth = effectiveOptions.ShowLegendKey
                            ? fontSize * (PptxChartMetricRules.DataLabelLegendKeySizeFactor + PptxChartMetricRules.DataLabelLegendKeyTextGapFactor)
                            : 0d;
                        double labelWidth = Math.Max(legendKeyWidth + 1d, barSlot * PptxChartMetricRules.VerticalBarDataLabelWidthRatio);
                        ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, x, y, labelWidth, labelHeight);
                        RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                        double textX = labelBox.X;
                        double textWidth = labelBox.Width;
                        TextAlignment alignment = TextAlignment.Center;
                        if (effectiveOptions.ShowLegendKey)
                        {
                            ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, densePointSeries.Count, varyColors, seriesFills, pointFills, value);
                            double consumedWidth = RenderFillDataLabelLegendKey(graphics, labelBox, fontSize, fill);
                            textX += consumedWidth;
                            textWidth = Math.Max(1d, textWidth - consumedWidth);
                            alignment = TextAlignment.Left;
                        }

                        if (!string.IsNullOrEmpty(label))
                        {
                            AddChartLabelRuns(runs, label, effectiveOptions, textX, labelBox.Y, textWidth, labelBox.Height, plotBox, style, alignment, fontResolver);
                        }
                    }
                }
            }
        }

        return RenderTextRuns(runs, graphics, "CBD", fontResolver);
    }

    private static IReadOnlyList<PdfFontResource> RenderLineDataLabels(
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<ChartIndexedNumberVector> series,
        ChartValueExtents extents,
        bool valueAxisReversed,
        IReadOnlyList<ChartSeriesStroke?> seriesStrokes,
        IReadOnlyList<ChartMarkerStyle> markerStyles,
        ChartDataLabelOptions labelOptions,
        IReadOnlyList<ChartDataLabelOptions> seriesLabelOptions,
        ChartIndexedTextVector categoryLabels,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames,
        PresentationFontResolver? fontResolver = null)
    {
        if ((!labelOptions.HasVisibleContent && !seriesLabelOptions.Any(options => options.HasVisibleContent)) || series.Count == 0)
        {
            return [];
        }

        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> densePointSeries = DensifyChartPointSeries(series);
        if (densePointSeries.Count == 0)
        {
            return [];
        }

        int pointCount = Math.Max(1, densePointSeries.Max(values => values.Count));
        double labelWidth = Math.Max(
            PptxChartMetricRules.CartesianDataLabelMinimumWidth,
            plotBox.Width / Math.Max(PptxChartMetricRules.LineDataLabelMinimumPointSpan, pointCount * PptxChartMetricRules.LineDataLabelPointWidthFactor));
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        var runs = new List<TextRun>();
        for (int seriesIndex = 0; seriesIndex < densePointSeries.Count; seriesIndex++)
        {
            IReadOnlyList<ChartIndexedNumberPoint?> points = densePointSeries[seriesIndex];
            for (int i = 0; i < points.Count; i++)
            {
                if (points[i]?.Value is not double value)
                {
                    continue;
                }

                double pointX = plotBox.X + (pointCount == 1 ? plotBox.Width / 2d : plotBox.Width * i / (pointCount - 1));
                double pointY = ChartValueToPlotCoordinate(extents, value, plotBox.Y, plotBox.Height, valueAxisReversed);
                ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), i);
                if (!effectiveOptions.HasVisibleContent)
                {
                    continue;
                }

                ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                double fontSize = style.FontSize;
                double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                ChartIndexedNumberPoint point = points[i]!.Value;
                string label = FormatCartesianDataLabel(value, seriesIndex, i, point, series[seriesIndex].WorkbookPointForIndex(point.Index), series[seriesIndex].FormatCode, effectiveOptions, categoryLabels, seriesNames);
                if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
                {
                    double legendKeyWidth = effectiveOptions.ShowLegendKey
                        ? GetStrokeMarkerDataLabelLegendKeyWidth(fontSize, ChartMarker(seriesIndex, markerStyles))
                        : 0d;
                    double effectiveLabelWidth = effectiveOptions.ShowLegendKey
                        ? Math.Max(labelWidth, textMeasurer.Measure(label, style) + legendKeyWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor)
                        : labelWidth;
                    (double labelX, double labelY, TextAlignment alignment) = ResolveLineDataLabelPosition(
                        effectiveOptions.PositionKind,
                        pointX,
                        pointY,
                        effectiveLabelWidth,
                        labelHeight);
                    ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, effectiveLabelWidth, labelHeight);
                    RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                    double textX = labelBox.X;
                    double textWidth = labelBox.Width;
                    if (effectiveOptions.ShowLegendKey)
                    {
                        ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, colorMap, null, seriesIndex, seriesStrokes, 1.2d);
                        ChartMarkerStyle marker = ChartMarker(seriesIndex, markerStyles);
                        double consumedWidth = RenderStrokeMarkerDataLabelLegendKey(graphics, labelBox, fontSize, marker, stroke);
                        textX += consumedWidth;
                        textWidth = Math.Max(1d, textWidth - consumedWidth);
                        alignment = TextAlignment.Left;
                    }

                    if (!string.IsNullOrEmpty(label))
                    {
                        AddChartLabelRuns(
                            runs,
                            label,
                            effectiveOptions,
                            textX,
                            labelBox.Y,
                            textWidth,
                            labelBox.Height,
                            plotBox,
                            style,
                            alignment,
                            fontResolver);
                    }
                }
            }
        }

        return RenderTextRuns(runs, graphics, "CLD", fontResolver);
    }

    private static double GetFillDataLabelLegendKeyWidth(double fontSize)
    {
        return fontSize * (PptxChartMetricRules.DataLabelLegendKeySizeFactor + PptxChartMetricRules.DataLabelLegendKeyTextGapFactor);
    }

    private static double RenderFillDataLabelLegendKey(PdfGraphicsBuilder graphics, ChartLayoutBox labelBox, double fontSize, ChartSeriesFill fill)
    {
        double swatchSize = fontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
        double swatchGap = fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
        double swatchY = labelBox.Y + Math.Max(0d, (labelBox.Height - swatchSize) / 2d);
        FillChartRectangle(graphics, labelBox.X, swatchY, swatchSize, swatchSize, fill);
        var stroke = new ChartSeriesStroke(fill.Color, fill.Alpha, ChartFilledSeriesInheritedStrokeWidth);
        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        graphics.StrokeRectangle(labelBox.X, swatchY, swatchSize, swatchSize);
        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }

        return swatchSize + swatchGap;
    }

    private static double GetStrokeMarkerDataLabelLegendKeyWidth(double fontSize, ChartMarkerStyle marker)
    {
        return Math.Max(marker.Size * 2d, fontSize * 1.05d) + fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
    }

    private static double RenderStrokeMarkerDataLabelLegendKey(PdfGraphicsBuilder graphics, ChartLayoutBox labelBox, double fontSize, ChartMarkerStyle marker, ChartSeriesStroke stroke)
    {
        double segmentLength = Math.Max(marker.Size * 2d, fontSize * 1.05d);
        double segmentY = labelBox.Y + labelBox.Height / 2d;
        SetChartStroke(graphics, stroke);
        graphics.StrokeLine(labelBox.X, segmentY, labelBox.X + segmentLength, segmentY);
        DrawChartMarker(graphics, labelBox.X + segmentLength / 2d, segmentY, marker, stroke.Color, stroke.Color);
        return segmentLength + fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
    }

    private static ChartLayoutBox ResolveDataLabelBox(ChartPlotBox plotBox, ChartDataLabelOptions options, double x, double y, double width, double height)
    {
        ChartLayoutBox defaultBox = new(x, y, width, height);
        if (!options.Layout.HasLayout)
        {
            return defaultBox;
        }

        ChartFrameBox frame = new(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height);
        return TryBuildManualLayoutBox(options.Layout, frame, defaultBox, out ChartLayoutBox manualBox, clampToFrame: false, missingPositionModesAreFactor: true)
            ? manualBox
            : defaultBox;
    }

    private static double ResolveHorizontalBarDataLabelX(PptxSceneChartDataLabelPosition position, double barBase, double barEnd, double labelWidth)
    {
        position = ResolveChartDataLabelPosition(position);
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
        position = ResolveChartDataLabelPosition(position);
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
        position = ResolveChartDataLabelPosition(position);
        return position switch
        {
            PptxSceneChartDataLabelPosition.Bottom => (pointX - labelWidth / 2d, pointY - labelHeight * PptxChartMetricRules.LineDataLabelBelowOffsetFactor, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Left => (pointX - labelWidth - PptxChartMetricRules.LineDataLabelSideGap, pointY - labelHeight / 2d, TextAlignment.Right),
            PptxSceneChartDataLabelPosition.Right => (pointX + PptxChartMetricRules.LineDataLabelSideGap, pointY - labelHeight / 2d, TextAlignment.Left),
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (pointX - labelWidth / 2d, pointY - labelHeight / 2d, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Top or PptxSceneChartDataLabelPosition.OutsideEnd or _ => (pointX - labelWidth / 2d, pointY + labelHeight * PptxChartMetricRules.LineDataLabelAboveOffsetFactor, TextAlignment.Center)
        };
    }

    private static PptxSceneChartDataLabelPosition ResolveChartDataLabelPosition(PptxSceneChartDataLabelPosition position)
    {
        return position == PptxSceneChartDataLabelPosition.Unknown
            ? PptxSceneChartDataLabelPosition.OutsideEnd
            : position;
    }

    private static IReadOnlyList<PdfFontResource> RenderScatterDataLabels(
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<ScatterSeries> series,
        bool bubble,
        ChartValueExtents xExtents,
        ChartValueExtents yExtents,
        IReadOnlyList<ChartSeriesFill?> seriesFills,
        ChartDataLabelOptions labelOptions,
        IReadOnlyList<ChartDataLabelOptions> seriesLabelOptions,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames,
        PresentationFontResolver? fontResolver = null)
    {
        if ((!labelOptions.HasVisibleContent && !seriesLabelOptions.Any(options => options.HasVisibleContent)) || series.Count == 0)
        {
            return [];
        }

        double maxBubbleSize = Math.Max(1d, series.SelectMany(item => item.Points).DefaultIfEmpty().Max(point => point.Size));
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        var runs = new List<TextRun>();
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            foreach (ScatterPoint point in series[seriesIndex].Points)
            {
                ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(
                    ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex),
                    point.Index);
                if (!effectiveOptions.HasVisibleContent)
                {
                    continue;
                }

                ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                double fontSize = style.FontSize;
                double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                string label = FormatScatterDataLabel(point, seriesIndex, effectiveOptions, seriesNames);
                if (string.IsNullOrEmpty(label) && !effectiveOptions.ShowLegendKey)
                {
                    continue;
                }

                double legendKeyWidth = effectiveOptions.ShowLegendKey
                    ? GetFillDataLabelLegendKeyWidth(fontSize)
                    : 0d;
                double labelWidth = Math.Max(
                    PptxChartMetricRules.CartesianDataLabelMinimumWidth,
                    textMeasurer.Measure(label, style) + legendKeyWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
                (double pointX, double pointY, double radius) = ResolveScatterPointGeometry(
                    plotBox,
                    point,
                    bubble,
                    xExtents,
                    yExtents,
                    maxBubbleSize);
                (double labelX, double labelY, TextAlignment alignment) = ResolveLineDataLabelPosition(
                    effectiveOptions.PositionKind,
                    pointX,
                    bubble ? pointY + radius : pointY,
                    labelWidth,
                    labelHeight);
                ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, labelWidth, labelHeight);
                RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                double textX = labelBox.X;
                double textWidth = labelBox.Width;
                if (effectiveOptions.ShowLegendKey)
                {
                    double consumedWidth = RenderFillDataLabelLegendKey(graphics, labelBox, fontSize, ChartSeriesColor(theme, colorMap, null, seriesIndex, seriesFills));
                    textX += consumedWidth;
                    textWidth = Math.Max(1d, textWidth - consumedWidth);
                    alignment = TextAlignment.Left;
                }

                if (!string.IsNullOrEmpty(label))
                {
                    AddChartLabelRuns(
                        runs,
                        label,
                        effectiveOptions,
                        textX,
                        labelBox.Y,
                        textWidth,
                        labelBox.Height,
                        plotBox,
                        style,
                        alignment,
                        fontResolver);
                }
            }
        }

        return RenderTextRuns(runs, graphics, "CSD", fontResolver);
    }

    private static string FormatScatterDataLabel(
        ScatterPoint point,
        int seriesIndex,
        ChartDataLabelOptions options,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames)
    {
        if (!string.IsNullOrWhiteSpace(options.CustomText))
        {
            return options.CustomText;
        }

        var parts = new List<string>(3);
        string seriesName = GetActiveSeriesName(seriesNames, seriesIndex);
        if (options.ShowSeriesName && !string.IsNullOrWhiteSpace(seriesName))
        {
            parts.Add(seriesName);
        }

        if (options.ShowValue)
        {
            parts.Add(FormatChartDataLabelValue(point.Y, options, point.YWorkbookPoint ?? point.YPoint, point.YFormatCode));
        }

        if (options.ShowBubbleSize && ResolveBubbleSizeValue(point) is { } bubbleSize)
        {
            parts.Add(FormatBubbleSizeDataLabelValue(bubbleSize, point, options));
        }

        return string.Join(GetChartDataLabelSeparator(options), parts);
    }

    private static double? ResolveBubbleSizeValue(ScatterPoint point)
    {
        return point.BubbleSizeWorkbookPoint?.Value ??
            point.BubbleSizePoint?.Value;
    }

    private static string FormatBubbleSizeDataLabelValue(double value, ScatterPoint point, ChartDataLabelOptions options)
    {
        if (options.NumberFormatInfo.IsDefined || !string.IsNullOrWhiteSpace(options.NumberFormat))
        {
            return FormatChartDataLabelValue(value, options, point.BubbleSizeWorkbookPoint ?? point.BubbleSizePoint, point.BubbleSizeFormatCode);
        }

        string formatCode = point.BubbleSizeFormatCode ?? string.Empty;
        return !string.IsNullOrWhiteSpace(formatCode) &&
            !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)
            ? FormatChartNumber(value, formatCode)
            : FormatChartAxisLabel(value);
    }

    private static (double X, double Y, double Radius) ResolveScatterPointGeometry(
        ChartPlotBox plotBox,
        ScatterPoint point,
        bool bubble,
        ChartValueExtents xExtents,
        ChartValueExtents yExtents,
        double maxBubbleSize)
    {
        double xRange = Math.Max(1d, xExtents.Max - xExtents.Min);
        double yRange = Math.Max(1d, yExtents.Max - yExtents.Min);
        double pointX = plotBox.X + (point.X - xExtents.Min) / xRange * plotBox.Width;
        double pointY = plotBox.Y + (point.Y - yExtents.Min) / yRange * plotBox.Height;
        double radius = bubble
            ? Math.Sqrt(Math.Max(0d, point.Size) / Math.Max(1d, maxBubbleSize)) * Math.Min(plotBox.Width, plotBox.Height) * PptxChartMetricRules.BubbleRadiusPlotRatio
            : 3d;
        return (pointX, pointY, radius);
    }

    private static ChartTextStyle ResolveChartDataLabelTextStyle(PptxTheme theme, PptxColorMap colorMap, ChartDataLabelOptions options)
    {
        RgbColor fallbackColor = theme.TryResolveColor("tx1", colorMap, out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        PptxThemeTypefaceResolution typeface = ResolveChartThemeTypeface(theme);
        ChartTextStyle style = new(
            typeface.Typeface,
            PptxChartMetricRules.DataLabelFallbackFontSize,
            fallbackColor,
            Alpha: 1d,
            Bold: false,
            Italic: false,
            Underline: false,
            Strike: false,
            typeface.RequestedTypeface,
            typeface.Typeface is null ? null : typeface.Source);
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
            dataLabel.TextStyle.Alpha ?? options.TextStyle.Alpha,
            dataLabel.TextStyle.Bold ?? options.TextStyle.Bold,
            dataLabel.TextStyle.Italic ?? options.TextStyle.Italic,
            dataLabel.TextStyle.Underline ?? options.TextStyle.Underline,
            dataLabel.TextStyle.Strike ?? options.TextStyle.Strike,
            dataLabel.TextStyle.FontFamily is null ? options.TextStyle.RequestedTypeface : dataLabel.TextStyle.RequestedTypeface,
            dataLabel.TextStyle.FontFamily is null ? options.TextStyle.TypefaceSource : dataLabel.TextStyle.TypefaceSource);
        return options with
        {
            ShowValue = dataLabel.ShowValue ?? options.ShowValue,
            ShowPercent = dataLabel.ShowPercent ?? options.ShowPercent,
            ShowCategoryName = dataLabel.ShowCategoryName ?? options.ShowCategoryName,
            ShowSeriesName = dataLabel.ShowSeriesName ?? options.ShowSeriesName,
            ShowLeaderLines = dataLabel.ShowLeaderLines ?? options.ShowLeaderLines,
            ShowLegendKey = dataLabel.ShowLegendKey ?? options.ShowLegendKey,
            ShowBubbleSize = dataLabel.ShowBubbleSize ?? options.ShowBubbleSize,
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
            TextBodyProperties = IsChartTextBodyPropertiesEmpty(dataLabel.TextBodyProperties) ? options.TextBodyProperties : dataLabel.TextBodyProperties,
            ShapeStyle = dataLabel.ShapeStyle.IsEmpty ? options.ShapeStyle : dataLabel.ShapeStyle,
            FlagOptions = ResolveChartDataLabelFlagOptions(options.FlagOptions, dataLabel.FlagOptions),
            Overrides = EmptyChartDataLabelOverrides
        };
    }

    private static ChartDataLabelOptions ResolveChartDataLabelOptionsForSeries(ChartDataLabelOptions plotOptions, IReadOnlyList<ChartDataLabelOptions> seriesOptions, int seriesIndex)
    {
        if (seriesIndex >= seriesOptions.Count || !seriesOptions[seriesIndex].IsDefined)
        {
            return plotOptions;
        }

        return MergeChartDataLabelOptions(plotOptions, seriesOptions[seriesIndex]);
    }

    private static ChartDataLabelOptions MergeChartDataLabelOptions(ChartDataLabelOptions baseOptions, ChartDataLabelOptions overrideOptions)
    {
        IReadOnlyDictionary<string, ChartBooleanOption> flags = ResolveChartDataLabelFlagOptions(baseOptions.FlagOptions, overrideOptions.FlagOptions);
        ChartTextStyleOverride textStyle = new(
            overrideOptions.TextStyle.FontFamily ?? baseOptions.TextStyle.FontFamily,
            overrideOptions.TextStyle.FontSize ?? baseOptions.TextStyle.FontSize,
            overrideOptions.TextStyle.Color ?? baseOptions.TextStyle.Color,
            overrideOptions.TextStyle.Alpha ?? baseOptions.TextStyle.Alpha,
            overrideOptions.TextStyle.Bold ?? baseOptions.TextStyle.Bold,
            overrideOptions.TextStyle.Italic ?? baseOptions.TextStyle.Italic,
            overrideOptions.TextStyle.Underline ?? baseOptions.TextStyle.Underline,
            overrideOptions.TextStyle.Strike ?? baseOptions.TextStyle.Strike,
            overrideOptions.TextStyle.FontFamily is null ? baseOptions.TextStyle.RequestedTypeface : overrideOptions.TextStyle.RequestedTypeface,
            overrideOptions.TextStyle.FontFamily is null ? baseOptions.TextStyle.TypefaceSource : overrideOptions.TextStyle.TypefaceSource);

        return baseOptions with
        {
            ShowValue = ChartDataLabelFlagValue(flags, "showVal"),
            ShowPercent = ChartDataLabelFlagValue(flags, "showPercent"),
            ShowCategoryName = ChartDataLabelFlagValue(flags, "showCatName"),
            ShowSeriesName = ChartDataLabelFlagValue(flags, "showSerName"),
            ShowLeaderLines = ChartDataLabelFlagValue(flags, "showLeaderLines"),
            ShowLegendKey = ChartDataLabelFlagValue(flags, "showLegendKey"),
            ShowBubbleSize = ChartDataLabelFlagValue(flags, "showBubbleSize"),
            LeaderLines = overrideOptions.LeaderLines.IsDefined ? overrideOptions.LeaderLines : baseOptions.LeaderLines,
            CustomText = string.IsNullOrEmpty(overrideOptions.CustomText) ? baseOptions.CustomText : overrideOptions.CustomText,
            CustomTextRuns = overrideOptions.CustomTextRuns.Count == 0 ? baseOptions.CustomTextRuns : overrideOptions.CustomTextRuns,
            PositionKind = string.IsNullOrEmpty(overrideOptions.Position) ? baseOptions.PositionKind : overrideOptions.PositionKind,
            Position = string.IsNullOrEmpty(overrideOptions.Position) ? baseOptions.Position : overrideOptions.Position,
            Separator = string.IsNullOrEmpty(overrideOptions.Separator) ? baseOptions.Separator : overrideOptions.Separator,
            NumberFormat = string.IsNullOrEmpty(overrideOptions.NumberFormat) ? baseOptions.NumberFormat : overrideOptions.NumberFormat,
            NumberFormatInfo = overrideOptions.NumberFormatInfo.IsDefined ? overrideOptions.NumberFormatInfo : baseOptions.NumberFormatInfo,
            Layout = overrideOptions.Layout.HasLayout ? overrideOptions.Layout : baseOptions.Layout,
            TextStyle = textStyle,
            TextBodyProperties = IsChartTextBodyPropertiesEmpty(overrideOptions.TextBodyProperties) ? baseOptions.TextBodyProperties : overrideOptions.TextBodyProperties,
            ShapeStyle = overrideOptions.ShapeStyle.IsEmpty ? baseOptions.ShapeStyle : overrideOptions.ShapeStyle,
            FlagOptions = flags,
            Overrides = MergeChartDataLabelOverrides(baseOptions.Overrides, overrideOptions.Overrides),
            IsDefined = baseOptions.IsDefined || overrideOptions.IsDefined
        };
    }

    private static bool ChartDataLabelFlagValue(IReadOnlyDictionary<string, ChartBooleanOption> flags, string name)
    {
        return flags.TryGetValue(name, out ChartBooleanOption option) && option.Value;
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> MergeChartDataLabelOverrides(IReadOnlyDictionary<int, ChartDataLabelOverride> baseOverrides, IReadOnlyDictionary<int, ChartDataLabelOverride> overrideOverrides)
    {
        if (baseOverrides.Count == 0)
        {
            return overrideOverrides;
        }

        if (overrideOverrides.Count == 0)
        {
            return baseOverrides;
        }

        var merged = new Dictionary<int, ChartDataLabelOverride>(baseOverrides);
        foreach (KeyValuePair<int, ChartDataLabelOverride> item in overrideOverrides)
        {
            merged[item.Key] = item.Value;
        }

        return merged;
    }

    private static bool IsChartTextBodyPropertiesEmpty(PptxSceneChartTextBodyProperties properties)
    {
        return properties.RotationDegrees is null && string.IsNullOrEmpty(properties.RotationValue);
    }

    private static TextRun CreateChartLabelRun(string text, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment)
    {
        return CreateChartTextRun(text, x, y, width, height, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, style, alignment);
    }

    private static void AddChartLabelRuns(List<TextRun> runs, string text, ChartDataLabelOptions options, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver = null)
    {
        ChartLayoutBox clipBox = ResolveDataLabelTextClipBox(plotBox, options, x, y, width, height);
        if (options.CustomTextRuns.Count == 0)
        {
            runs.Add(CreateChartTextRun(text, x, y, width, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, style, alignment));
            return;
        }

        AddChartRichTextRuns(runs, options.CustomTextRuns, text, x, y, width, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, style, alignment, fontResolver);
    }

    private static void AddPolarChartLabelRuns(List<TextRun> runs, IReadOnlyList<string> parts, string fallbackText, ChartDataLabelOptions options, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver = null)
    {
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        ChartLayoutBox clipBox = ResolveDataLabelTextClipBox(plotBox, options, x, y, width, height);
        if (parts.Count <= 1 || options.CustomTextRuns.Count > 0 || !ShouldSplitPolarDataLabelParts(options))
        {
            AddChartLabelRuns(runs, fallbackText, options, x, y, width, height, plotBox, style, alignment, fontResolver);
            return;
        }

        ChartTextRunLayout[] labelRuns = parts
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .Select(part => new ChartTextRunLayout(part, style, Math.Max(0d, textMeasurer.Measure(part, style))))
            .Where(run => run.Width > 0d)
            .ToArray();
        if (labelRuns.Length <= 1)
        {
            AddChartLabelRuns(runs, fallbackText, options, x, y, width, height, plotBox, style, alignment, fontResolver);
            return;
        }

        double separatorWidth = textMeasurer.Measure(GetChartDataLabelSeparator(options), style);
        double totalWidth = labelRuns.Sum(run => run.Width) + separatorWidth * Math.Max(0, labelRuns.Length - 1);
        double cursor = alignment switch
        {
            TextAlignment.Right => x + Math.Max(1d, width) - totalWidth,
            TextAlignment.Center => x + (Math.Max(1d, width) - totalWidth) / 2d,
            _ => x
        };

        foreach (ChartTextRunLayout run in labelRuns)
        {
            double runWidth = Math.Max(0.1d, run.Width);
            runs.Add(CreateChartTextRun(run.Text, cursor, y, runWidth, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, run.Style, TextAlignment.Left) with { PreventCoalesce = true });
            cursor += runWidth + separatorWidth;
        }
    }

    private static bool ShouldSplitPolarDataLabelParts(ChartDataLabelOptions options)
    {
        string separator = GetChartDataLabelSeparator(options);
        return separator.Length > 0 && separator.All(char.IsWhiteSpace);
    }

    private static ChartLayoutBox ResolveDataLabelTextClipBox(ChartPlotBox plotBox, ChartDataLabelOptions options, double x, double y, double width, double height)
    {
        if (!options.Layout.HasLayout)
        {
            return new ChartLayoutBox(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height);
        }

        double left = Math.Min(plotBox.X, x);
        double top = Math.Min(plotBox.Y, y);
        double right = Math.Max(plotBox.X + plotBox.Width, x + width);
        double bottom = Math.Max(plotBox.Y + plotBox.Height, y + height);
        return new ChartLayoutBox(left, top, Math.Max(1d, right - left), Math.Max(1d, bottom - top));
    }

    private static void AddChartRichTextRuns(List<TextRun> runs, IReadOnlyList<ChartTextRunOverride> richTextRuns, string fallbackText, double x, double y, double width, double height, double clipX, double clipY, double clipWidth, double clipHeight, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver = null)
    {
        if (richTextRuns.Count == 0)
        {
            runs.Add(CreateChartTextRun(fallbackText, x, y, width, height, clipX, clipY, clipWidth, clipHeight, style, alignment));
            return;
        }

        var textMeasurer = new ChartTextMeasurer(fontResolver);
        ChartTextRunLayout[] richRuns = richTextRuns
            .Where(run => !string.IsNullOrEmpty(run.Text))
            .Select(run =>
            {
                ChartTextStyle runStyle = MergeChartTextStyle(style, run.TextStyle);
                return new ChartTextRunLayout(run.Text, runStyle, Math.Max(0d, textMeasurer.Measure(run.Text, runStyle)));
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
            style.Alpha,
            null,
            Bold: style.Bold,
            Italic: style.Italic,
            Underline: style.Underline,
            Strike: style.Strike,
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
        return ReadChartTextStyle(theme, PptxColorMap.Default, chartXml, element, fallbackFontSize);
    }

    private static ChartTextStyle ReadChartTextStyle(PptxTheme theme, PptxColorMap colorMap, XDocument chartXml, XElement? element, double fallbackFontSize)
    {
        ChartTextStyle style = CreateDefaultChartTextStyle(theme, colorMap, fallbackFontSize);
        style = MergeChartTextStyle(style, ReadChartTextStyleFromTxPr(chartXml.Root, theme, colorMap));
        style = MergeChartTextStyle(style, ReadChartTextStyleFromTxPr(element, theme, colorMap));
        return style;
    }

    private static ChartTextStyle ReadSceneOrXmlChartTextStyle(PptxTheme theme, PptxSceneChart? sceneChart, PptxSceneChartAxis? sceneAxis, XDocument chartXml, XElement? element, double fallbackFontSize, string? chartStyleRole = null)
    {
        if (sceneChart is null)
        {
            return ReadChartTextStyle(theme, chartXml, element, fallbackFontSize);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, sceneChart.ColorMap, fallbackFontSize);
        style = MergeChartTextStyle(style, ToChartTextStyleOverride(sceneChart.TextStyle));
        if (!string.IsNullOrWhiteSpace(chartStyleRole))
        {
            style = MergeChartTextStyle(style, ReadChartStyleRoleTextStyle(sceneChart.StylePart, chartStyleRole));
        }

        if (sceneAxis is not null)
        {
            style = MergeChartTextStyle(style, ToChartTextStyleOverride(sceneAxis.TextStyle));
        }

        return style;
    }

    private static ChartTextStyle CreateDefaultChartTextStyle(PptxTheme theme, double fallbackFontSize)
    {
        return CreateDefaultChartTextStyle(theme, PptxColorMap.Default, fallbackFontSize);
    }

    private static ChartTextStyle CreateDefaultChartTextStyle(PptxTheme theme, PptxColorMap colorMap, double fallbackFontSize)
    {
        RgbColor fallbackColor = theme.TryResolveColor("tx1", colorMap, out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        PptxThemeTypefaceResolution typeface = ResolveChartThemeTypeface(theme);
        return new ChartTextStyle(
            typeface.Typeface,
            fallbackFontSize,
            fallbackColor,
            Alpha: 1d,
            Bold: false,
            Italic: false,
            Underline: false,
            Strike: false,
            typeface.RequestedTypeface,
            typeface.Typeface is null ? null : typeface.Source);
    }

    private static ChartTextStyleOverride ToChartTextStyleOverride(PptxSceneChartTextStyleOverride style)
    {
        return new ChartTextStyleOverride(
            style.FontFamily,
            style.FontSize,
            style.Color,
            style.Alpha,
            style.Bold,
            style.Italic,
            style.Underline,
            style.Strike,
            style.RequestedTypeface,
            style.TypefaceSource);
    }

    private static ChartTextStyleOverride ReadChartStyleRoleTextStyle(PptxSceneChartStyle stylePart, string role)
    {
        PptxSceneChartStyleEntry entry = stylePart.Entries.FirstOrDefault(item => item.Role == role);
        return ToChartTextStyleOverride(entry.TextStyle);
    }

    private static PptxThemeTypefaceResolution ResolveChartThemeTypeface(PptxTheme theme)
    {
        PptxThemeTypefaceResolution minorLatin = theme.ResolveTypefaceWithSource("+mn-lt");
        if (minorLatin.Typeface is not null)
        {
            return minorLatin;
        }

        PptxThemeTypefaceResolution majorLatin = theme.ResolveTypefaceWithSource("+mj-lt");
        return majorLatin.Typeface is null ? default : majorLatin;
    }

    private static ChartTextStyleOverride ReadChartTextStyleFromTxPr(XElement? parent, PptxTheme theme)
    {
        return ReadChartTextStyleFromTxPr(parent, theme, PptxColorMap.Default);
    }

    private static ChartTextStyleOverride ReadChartTextStyleFromTxPr(XElement? parent, PptxTheme theme, PptxColorMap colorMap)
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
        PptxThemeTypefaceResolution typefaceResolution = string.IsNullOrWhiteSpace(typeface)
            ? default
            : theme.ResolveTypefaceWithSource(typeface);
        string? fontFamily = string.IsNullOrWhiteSpace(typeface)
            ? null
            : typefaceResolution.Typeface;

        double? fontSize = null;
        if (defRunProperties.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0)
        {
            fontSize = sizeHundredths / 100d;
        }

        RgbColor? color = TryReadSolidColorWithAlpha(defRunProperties.Element(DrawingNamespace + "solidFill"), theme, colorMap, out RgbColor parsedColor, out double alpha)
            ? parsedColor
            : null;
        bool? bold = defRunProperties.Attribute("b") is { } boldAttribute
            ? IsOoxmlTrue(boldAttribute.Value)
            : null;
        bool? italic = defRunProperties.Attribute("i") is { } italicAttribute
            ? IsOoxmlTrue(italicAttribute.Value)
            : null;
        bool? underline = ReadChartUnderline(defRunProperties);
        bool? strike = ReadChartStrike(defRunProperties);
        return new ChartTextStyleOverride(
            fontFamily,
            fontSize,
            color,
            color is null ? null : alpha,
            bold,
            italic,
            underline,
            strike,
            typefaceResolution.RequestedTypeface,
            typefaceResolution.RequestedTypeface is null ? null : typefaceResolution.Source);
    }

    private static PptxSceneChartTextBodyProperties ReadChartTextBodyProperties(XElement? parent)
    {
        string rotation = (string?)parent?
            .Element(ChartNamespace + "txPr")?
            .Element(DrawingNamespace + "bodyPr")?
            .Attribute("rot") ?? string.Empty;
        return new PptxSceneChartTextBodyProperties(ParseChartOoxmlAngle(rotation), rotation);
    }

    private static double? ParseChartOoxmlAngle(string value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long rawAngle)
            ? rawAngle / 60000d
            : null;
    }

    private static bool? ReadChartUnderline(XElement runProperties)
    {
        string? value = (string?)runProperties.Attribute("u");
        return value is null ? null : !value.Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static bool? ReadChartStrike(XElement runProperties)
    {
        string? value = (string?)runProperties.Attribute("strike");
        return value is null ? null : !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static ChartTextStyle MergeChartTextStyle(ChartTextStyle style, ChartTextStyleOverride next)
    {
        return new ChartTextStyle(
            next.FontFamily ?? style.FontFamily,
            next.FontSize ?? style.FontSize,
            next.Color ?? style.Color,
            next.Alpha ?? style.Alpha,
            next.Bold ?? style.Bold,
            next.Italic ?? style.Italic,
            next.Underline ?? style.Underline,
            next.Strike ?? style.Strike,
            next.FontFamily is null ? style.RequestedTypeface : next.RequestedTypeface,
            next.FontFamily is null ? style.TypefaceSource : next.TypefaceSource);
    }

    private static ChartTextStyleOverride MergeChartTextStyleOverride(ChartTextStyleOverride style, ChartTextStyleOverride next)
    {
        return new ChartTextStyleOverride(
            next.FontFamily ?? style.FontFamily,
            next.FontSize ?? style.FontSize,
            next.Color ?? style.Color,
            next.Alpha ?? style.Alpha,
            next.Bold ?? style.Bold,
            next.Italic ?? style.Italic,
            next.Underline ?? style.Underline,
            next.Strike ?? style.Strike,
            next.FontFamily is null ? style.RequestedTypeface : next.RequestedTypeface,
            next.FontFamily is null ? style.TypefaceSource : next.TypefaceSource);
    }

    private static ChartDataLabelOptions ReadChartDataLabelOptions(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
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
                IsChartLabelFlagEnabled(labels, "showLegendKey"),
                IsChartLabelFlagEnabled(labels, "showBubbleSize"),
                ReadChartDataLabelLeaderLines(labels, theme, colorMap),
                string.Empty,
                [],
                PptxSceneBuilder.ParseChartDataLabelPosition(labels.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty),
                labels.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty,
                labels.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                labels.Element(ChartNamespace + "numFmt")?.Attribute("formatCode")?.Value ?? string.Empty,
                ReadChartNumberFormat(labels),
                ReadManualLayout(labels),
                ReadChartTextStyleFromTxPr(labels, theme, colorMap),
                ReadChartTextBodyProperties(labels),
                ReadChartShapeStyle(labels.Element(ChartNamespace + "spPr"), theme, colorMap),
                ReadChartDataLabelFlagOptions(labels),
                ReadChartDataLabelOverrides(labels, theme, colorMap),
                IsDefined: true);
    }

    private static ChartDataLabelOptions ReadSceneOrXmlDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return plot is null
            ? ReadChartDataLabelOptions(chartElement, theme, colorMap)
            : new ChartDataLabelOptions(
                plot.DataLabels.ShowValue == true,
                plot.DataLabels.ShowPercent == true,
                plot.DataLabels.ShowCategoryName == true,
                plot.DataLabels.ShowSeriesName == true,
                plot.DataLabels.ShowLeaderLines == true,
                plot.DataLabels.ShowLegendKey == true,
                plot.DataLabels.ShowBubbleSize == true,
                ToChartDataLabelLeaderLines(plot.DataLabels.LeaderLines),
                string.Empty,
                [],
                plot.DataLabels.PositionKind,
                plot.DataLabels.Position,
                plot.DataLabels.Separator,
                plot.DataLabels.NumberFormat,
                ToChartNumberFormat(plot.DataLabels.NumberFormatInfo),
                plot.DataLabels.Layout,
                ToChartDataLabelTextStyleOverride(sceneChart, plot.DataLabels),
                plot.DataLabels.TextBodyProperties,
                ToChartShapeStyle(plot.DataLabels.ShapeStyle),
                ToChartDataLabelFlagOptions(plot.DataLabels),
                ToChartDataLabelOverrides(plot.DataLabels.Overrides),
                plot.DataLabels.IsDefined);
    }

    private static IReadOnlyList<ChartDataLabelOptions> ReadSceneOrXmlSeriesDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => ToChartDataLabelOptions(sceneChart, series.DataLabels))
                .ToArray();
        }

        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => ReadChartDataLabelOptions(series, theme, colorMap))
            .ToArray();
    }

    private static ChartDataLabelOptions ToChartDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartDataLabels labels)
    {
        return new ChartDataLabelOptions(
            labels.ShowValue == true,
            labels.ShowPercent == true,
            labels.ShowCategoryName == true,
            labels.ShowSeriesName == true,
            labels.ShowLeaderLines == true,
            labels.ShowLegendKey == true,
            labels.ShowBubbleSize == true,
            ToChartDataLabelLeaderLines(labels.LeaderLines),
            string.Empty,
            [],
            labels.PositionKind,
            labels.Position,
            labels.Separator,
            labels.NumberFormat,
            ToChartNumberFormat(labels.NumberFormatInfo),
            labels.Layout,
            ToChartDataLabelTextStyleOverride(sceneChart, labels),
            labels.TextBodyProperties,
            ToChartShapeStyle(labels.ShapeStyle),
            ToChartDataLabelFlagOptions(labels),
            ToChartDataLabelOverrides(labels.Overrides),
            labels.IsDefined);
    }

    private static ChartTextStyleOverride ToChartDataLabelTextStyleOverride(PptxSceneChart? sceneChart, PptxSceneChartDataLabels labels)
    {
        ChartTextStyleOverride style = ChartTextStyleOverride.Empty;
        if (sceneChart is not null)
        {
            style = MergeChartTextStyleOverride(style, ToChartTextStyleOverride(sceneChart.TextStyle));
            style = MergeChartTextStyleOverride(style, ReadChartStyleRoleTextStyle(sceneChart.StylePart, "dataLabel"));
        }

        return MergeChartTextStyleOverride(style, ToChartTextStyleOverride(labels.TextStyle));
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> ReadChartDataLabelOverrides(XElement labels, PptxTheme theme, PptxColorMap colorMap)
    {
        var overrides = new Dictionary<int, ChartDataLabelOverride>();
        foreach (XElement label in labels.Elements(ChartNamespace + "dLbl"))
        {
            if (!TryReadChartNonNegativeIndex(label, out int index))
            {
                continue;
            }

            overrides[index] = new ChartDataLabelOverride(
                ReadOptionalChartLabelFlagEnabled(label, "showVal"),
                ReadOptionalChartLabelFlagEnabled(label, "showPercent"),
                ReadOptionalChartLabelFlagEnabled(label, "showCatName"),
                ReadOptionalChartLabelFlagEnabled(label, "showSerName"),
                ReadOptionalChartLabelFlagEnabled(label, "showLeaderLines"),
                ReadOptionalChartLabelFlagEnabled(label, "showLegendKey"),
                ReadOptionalChartLabelFlagEnabled(label, "showBubbleSize"),
                ReadChartDataLabelLeaderLines(label, theme, colorMap),
                ReadChartText(label.Element(ChartNamespace + "tx")) ?? string.Empty,
                ToChartTextRuns(PptxSceneBuilder.ReadChartTextRuns(label.Element(ChartNamespace + "tx"), theme, colorMap)),
                PptxSceneBuilder.ParseChartDataLabelPosition(label.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty),
                label.Element(ChartNamespace + "dLblPos")?.Attribute("val")?.Value ?? string.Empty,
                label.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                label.Element(ChartNamespace + "numFmt")?.Attribute("formatCode")?.Value ?? string.Empty,
                ReadChartNumberFormat(label),
                ReadManualLayout(label),
                ReadChartTextStyleFromTxPr(label, theme, colorMap),
                ReadChartTextBodyProperties(label),
                ReadChartShapeStyle(label.Element(ChartNamespace + "spPr"), theme, colorMap),
                ReadChartDataLabelFlagOptions(label));
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
                dataLabel.ShowLegendKey,
                dataLabel.ShowBubbleSize,
                ToChartDataLabelLeaderLines(dataLabel.LeaderLines),
                dataLabel.CustomText,
                ToChartTextRuns(dataLabel.CustomTextRuns),
                dataLabel.PositionKind,
                dataLabel.Position,
                dataLabel.Separator,
                dataLabel.NumberFormat,
                ToChartNumberFormat(dataLabel.NumberFormatInfo),
                dataLabel.Layout,
                ToChartTextStyleOverride(dataLabel.TextStyle),
                dataLabel.TextBodyProperties,
                ToChartShapeStyle(dataLabel.ShapeStyle),
                ToChartDataLabelOverrideFlagOptions(dataLabel));
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
                    : null,
                SourceLinkedValue: (string?)numberFormat.Attribute("sourceLinked") ?? string.Empty);
    }

    private static ChartNumberFormat ToChartNumberFormat(PptxSceneChartNumberFormat numberFormat)
    {
        return numberFormat.IsDefined
            ? new ChartNumberFormat(numberFormat.IsDefined, numberFormat.FormatCode, numberFormat.SourceLinked, numberFormat.SourceLinkedValue)
            : default;
    }

    private static ChartDataLabelLeaderLines ReadChartDataLabelLeaderLines(XElement labels, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? leaderLines = labels.Element(ChartNamespace + "leaderLines");
        if (leaderLines is null)
        {
            return ChartDataLabelLeaderLines.Empty;
        }

        return new ChartDataLabelLeaderLines(IsDefined: true, ReadChartShapeStyle(leaderLines.Element(ChartNamespace + "spPr"), theme, colorMap).Stroke);
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

    private static IReadOnlyDictionary<string, ChartBooleanOption> ReadChartDataLabelFlagOptions(XElement labels)
    {
        var flags = new Dictionary<string, ChartBooleanOption>(ChartDataLabelFlagNames.Length, StringComparer.Ordinal);
        foreach (string flagName in ChartDataLabelFlagNames)
        {
            XElement? element = labels.Element(ChartNamespace + flagName);
            flags[flagName] = new ChartBooleanOption(
                IsOoxmlBooleanElementEnabled(element),
                (string?)element?.Attribute("val") ?? string.Empty,
                element is not null);
        }

        return flags;
    }

    private static IReadOnlyDictionary<string, ChartBooleanOption> ToChartDataLabelFlagOptions(PptxSceneChartDataLabels labels)
    {
        return new Dictionary<string, ChartBooleanOption>(StringComparer.Ordinal)
        {
            ["showVal"] = new(labels.ShowValue == true, labels.ShowValueValue, labels.ShowValue is not null),
            ["showPercent"] = new(labels.ShowPercent == true, labels.ShowPercentValue, labels.ShowPercent is not null),
            ["showCatName"] = new(labels.ShowCategoryName == true, labels.ShowCategoryNameValue, labels.ShowCategoryName is not null),
            ["showSerName"] = new(labels.ShowSeriesName == true, labels.ShowSeriesNameValue, labels.ShowSeriesName is not null),
            ["showLeaderLines"] = new(labels.ShowLeaderLines == true, labels.ShowLeaderLinesValue, labels.ShowLeaderLines is not null),
            ["showLegendKey"] = new(labels.ShowLegendKey == true, labels.ShowLegendKeyValue, labels.ShowLegendKey is not null),
            ["showBubbleSize"] = new(labels.ShowBubbleSize == true, labels.ShowBubbleSizeValue, labels.ShowBubbleSize is not null)
        };
    }

    private static IReadOnlyDictionary<string, ChartBooleanOption> ToChartDataLabelOverrideFlagOptions(PptxSceneChartDataLabelOverride label)
    {
        return new Dictionary<string, ChartBooleanOption>(StringComparer.Ordinal)
        {
            ["showVal"] = new(label.ShowValue == true, label.ShowValueValue, label.ShowValue is not null),
            ["showPercent"] = new(label.ShowPercent == true, label.ShowPercentValue, label.ShowPercent is not null),
            ["showCatName"] = new(label.ShowCategoryName == true, label.ShowCategoryNameValue, label.ShowCategoryName is not null),
            ["showSerName"] = new(label.ShowSeriesName == true, label.ShowSeriesNameValue, label.ShowSeriesName is not null),
            ["showLeaderLines"] = new(label.ShowLeaderLines == true, label.ShowLeaderLinesValue, label.ShowLeaderLines is not null),
            ["showLegendKey"] = new(label.ShowLegendKey == true, label.ShowLegendKeyValue, label.ShowLegendKey is not null),
            ["showBubbleSize"] = new(label.ShowBubbleSize == true, label.ShowBubbleSizeValue, label.ShowBubbleSize is not null)
        };
    }

    private static IReadOnlyDictionary<string, ChartBooleanOption> ResolveChartDataLabelFlagOptions(
        IReadOnlyDictionary<string, ChartBooleanOption> baseFlags,
        IReadOnlyDictionary<string, ChartBooleanOption> overrideFlags)
    {
        if (overrideFlags.Count == 0)
        {
            return baseFlags;
        }

        var resolved = new Dictionary<string, ChartBooleanOption>(ChartDataLabelFlagNames.Length, StringComparer.Ordinal);
        foreach (string flagName in ChartDataLabelFlagNames)
        {
            if (overrideFlags.TryGetValue(flagName, out ChartBooleanOption overrideOption) && overrideOption.IsDefined)
            {
                resolved[flagName] = overrideOption;
            }
            else if (baseFlags.TryGetValue(flagName, out ChartBooleanOption baseOption))
            {
                resolved[flagName] = baseOption;
            }
            else
            {
                resolved[flagName] = new ChartBooleanOption(false, string.Empty, false);
            }
        }

        return resolved;
    }

    private static string FormatChartPercentageLabel(double fraction)
    {
        return (fraction * 100d).ToString("0.#", CultureInfo.InvariantCulture) + "%";
    }

    private static IReadOnlyList<string> FormatPieDataLabelParts(double value, double total, int categoryIndex, ChartIndexedNumberPoint? workbookPoint, string? valueFormatCode, ChartDataLabelOptions options, ChartIndexedTextVector categoryLabels, IReadOnlyList<ChartSeriesNameRecord> seriesNames)
    {
        if (!string.IsNullOrWhiteSpace(options.CustomText))
        {
            return [options.CustomText];
        }

        var parts = new List<string>(4);
        string seriesName = GetActiveSeriesName(seriesNames, 0);
        if (options.ShowSeriesName && !string.IsNullOrWhiteSpace(seriesName))
        {
            parts.Add(seriesName);
        }

        string categoryLabel = GetIndexedCategoryLabel(categoryLabels, categoryIndex);
        if (options.ShowCategoryName && !string.IsNullOrWhiteSpace(categoryLabel))
        {
            parts.Add(categoryLabel);
        }

        if (options.ShowValue)
        {
            parts.Add(FormatChartDataLabelValue(value, options, workbookPoint, valueFormatCode));
        }

        if (options.ShowPercent)
        {
            parts.Add(FormatChartPercentageLabel(value / total));
        }

        return parts;
    }

    private static string JoinChartDataLabelParts(IReadOnlyList<string> parts, ChartDataLabelOptions options)
    {
        return string.Join(GetChartDataLabelSeparator(options), parts);
    }

    private static string FormatCartesianDataLabel(
        double value,
        int seriesIndex,
        int categoryIndex,
        ChartIndexedNumberPoint point,
        ChartIndexedNumberPoint? workbookPoint,
        string? valueFormatCode,
        ChartDataLabelOptions options,
        ChartIndexedTextVector categoryLabels,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames)
    {
        if (!string.IsNullOrWhiteSpace(options.CustomText))
        {
            return options.CustomText;
        }

        var parts = new List<string>(3);
        string seriesName = GetActiveSeriesName(seriesNames, seriesIndex);
        if (options.ShowSeriesName && !string.IsNullOrWhiteSpace(seriesName))
        {
            parts.Add(seriesName);
        }

        string categoryLabel = GetIndexedCategoryLabel(categoryLabels, categoryIndex);
        if (options.ShowCategoryName && !string.IsNullOrWhiteSpace(categoryLabel))
        {
            parts.Add(categoryLabel);
        }

        if (options.ShowValue)
        {
            parts.Add(FormatChartDataLabelValue(value, options, workbookPoint ?? point, valueFormatCode));
        }

        return string.Join(GetChartDataLabelSeparator(options), parts);
    }

    private static string GetIndexedCategoryLabel(ChartIndexedTextVector categoryLabels, int categoryIndex)
    {
        return (categoryLabels.Points ?? [])
            .Where(point => point.HasText && point.Index == categoryIndex)
            .Select(point => point.Text)
            .FirstOrDefault() ?? string.Empty;
    }

    private static string GetActiveSeriesName(IReadOnlyList<ChartSeriesNameRecord> seriesNames, int seriesIndex)
    {
        return seriesIndex >= 0 && seriesIndex < seriesNames.Count
            ? seriesNames[seriesIndex].ActiveName
            : string.Empty;
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options)
    {
        return FormatChartDataLabelValue(value, options.NumberFormatInfo, options.NumberFormat);
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options, ChartIndexedNumberPoint? sourcePoint)
    {
        return FormatChartDataLabelValue(value, options, sourcePoint, sourceFormatCode: null);
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options, ChartIndexedNumberPoint? sourcePoint, string? sourceFormatCode)
    {
        if (ResolveSourceLinkedChartNumberFormatCode(options.NumberFormatInfo, sourcePoint) is { } sourceFormat)
        {
            return FormatChartNumber(value, sourceFormat);
        }

        return FormatChartDataLabelValue(value, options.NumberFormatInfo, options.NumberFormat, sourceFormatCode);
    }

    private static string FormatChartDataLabelValue(double value, ChartNumberFormat numberFormat, string legacyNumberFormat)
    {
        return FormatChartDataLabelValue(value, numberFormat, legacyNumberFormat, sourceFormatCode: null);
    }

    private static string FormatChartDataLabelValue(double value, ChartNumberFormat numberFormat, string legacyNumberFormat, string? sourceFormatCode)
    {
        if (IsRenderableChartNumberFormat(numberFormat))
        {
            return FormatChartNumber(value, numberFormat.FormatCode);
        }

        return !string.IsNullOrWhiteSpace(legacyNumberFormat) &&
            !string.Equals(legacyNumberFormat, "General", StringComparison.OrdinalIgnoreCase)
            ? FormatChartNumber(value, legacyNumberFormat)
            : IsRenderableChartFormatCode(sourceFormatCode)
                ? FormatChartNumber(value, sourceFormatCode!)
            : FormatChartAxisLabel(value);
    }

    private static string? ResolveSourceLinkedChartNumberFormatCode(ChartNumberFormat numberFormat, ChartIndexedNumberPoint? sourcePoint)
    {
        if (sourcePoint is null)
        {
            return null;
        }

        ChartWorkbookRangeCell cell = sourcePoint.Value.WorkbookCell;
        return ResolveSourceLinkedChartNumberFormatCode(
            numberFormat,
            cell.StyleNumberFormatCode,
            cell.StyleAppliesNumberFormat,
            cell.StyleNumberFormatIsDateLike);
    }

    private static string? ResolveSourceLinkedChartNumberFormatCode(
        ChartNumberFormat numberFormat,
        string? workbookFormatCode,
        bool? workbookAppliesNumberFormat,
        bool workbookFormatIsDateLike)
    {
        if (numberFormat.SourceLinked != true ||
            workbookAppliesNumberFormat == false ||
            workbookFormatIsDateLike ||
            !IsRenderableChartFormatCode(workbookFormatCode))
        {
            return null;
        }

        return workbookFormatCode;
    }

    private static string GetChartDataLabelSeparator(ChartDataLabelOptions options)
    {
        return string.IsNullOrEmpty(options.Separator) ? ", " : options.Separator;
    }

    private static IReadOnlyList<PdfFontResource> RenderChartCategoryLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartAxis? sceneAxis, XElement? categoryAxis, ChartIndexedTextVector labelVector, bool horizontalBars, double? verticalAxisY = null, bool categoryLabelsOnTickMarks = false, bool categoryLabelsTopSide = false, PresentationFontResolver? fontResolver = null)
    {
        IReadOnlyList<ChartIndexedTextPoint?> labels = labelVector.DensePoints();
        if (labels.Count == 0)
        {
            return [];
        }

        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, categoryAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
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

            string? label = labels[i]?.Text;
            if (string.IsNullOrWhiteSpace(label))
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
                double labelCenterX = categoryLabelsOnTickMarks && labels.Count > 1
                    ? plotBox.X + plotBox.Width * i / (labels.Count - 1)
                    : plotBox.X + slotWidth * (i + 0.5d);
                x = labelCenterX - width / 2d;
                double axisY = verticalAxisY ?? plotBox.Y;
                y = categoryLabelsTopSide
                    ? axisY + height * PptxChartMetricRules.CategoryAxisVerticalTopSideOffsetFactor * labelOffsetScale
                    : axisY - height * PptxChartMetricRules.CategoryAxisVerticalTopOffsetFactor * labelOffsetScale;
                alignment = TextAlignment.Center;
            }

            double labelWidth = Math.Max(1d, width);
            runs.Add(new TextRun(
                label,
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
                Underline: style.Underline,
                Strike: style.Strike,
                KerningEnabled: true,
                alignment,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false));
        }

        return RenderTextRuns(runs, graphics, "CCA", fontResolver);
    }

    private static IReadOnlyList<PdfFontResource> RenderChartValueAxisLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, XElement? valueAxis, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits axisUnits, bool valueAxisReversed, bool horizontalBars, bool rightSide = false, int axisSideSlot = 0, bool manualPlotLayoutApplied = false, bool useTextSizedWidth = false, string? defaultNumberFormat = null, PresentationFontResolver? fontResolver = null)
    {
        double range = Math.Max(1d, extents.Max - extents.Min);
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
        double fontSize = style.FontSize;
        double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
        RgbColor color = style.Color;
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        double autoTickTargetCount = GetValueAxisAutoTickTargetCount(horizontalBars, valueAxisLabelsVisible: true, manualPlotLayoutApplied);
        IReadOnlyList<double> tickValues = GetChartAxisTickValues(extents, axisUnits.MajorUnit, includeEndpoints: true, autoTickTargetCount);
        double maxLabelWidth = tickValues
            .Select(value => FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis, defaultNumberFormat))
            .DefaultIfEmpty("0")
            .Max(label => textMeasurer.Measure(label, style));
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
                Underline: style.Underline,
                Strike: style.Strike,
                KerningEnabled: true,
                alignment,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false));
        }

        return RenderTextRuns(runs, graphics, "CVA", fontResolver);
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

    private sealed class ChartTextMeasurer
    {
        private readonly TextAdvanceEstimator estimator;

        public ChartTextMeasurer(PresentationFontResolver? fontResolver)
        {
            estimator = new TextAdvanceEstimator(fontResolver);
        }

        public double Measure(string text, ChartTextStyle style)
        {
            return Measure(text, style.FontSize, style.FontFamily, style.Bold, style.Italic);
        }

        public double Measure(string text, double fontSize, string? fontFamily = null, bool bold = false, bool italic = false)
        {
            return estimator.Measure(
                text,
                fontSize,
                fontFamily,
                bold,
                italic,
                kerningEnabled: true);
        }
    }

    private static int GetValueAxisSideSlot(PptxSceneChartAxis? primarySceneAxis, XElement? primaryAxis, PptxSceneChartAxis? secondarySceneAxis, XElement secondaryAxis, bool defaultPrimaryRightSide, bool defaultSecondaryRightSide)
    {
        bool primaryRight = ResolveSceneOrXmlValueAxisLabelsRightSide(primarySceneAxis, primaryAxis, defaultPrimaryRightSide);
        bool secondaryRight = ResolveSceneOrXmlValueAxisLabelsRightSide(secondarySceneAxis, secondaryAxis, defaultSecondaryRightSide);
        return primaryRight == secondaryRight ? 1 : 0;
    }

    private static IReadOnlyList<PdfFontResource> RenderSecondaryChartValueAxisLabels(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, ChartValueExtents fallback, PresentationFontResolver fontResolver)
    {
        ChartAxisSource rightValueAxisSource = ReadSceneOrXmlSecondaryRightValueAxis(sceneChart, chartXml);
        XElement? rightValueAxis = rightValueAxisSource.XmlAxis;
        PptxSceneChartAxis? rightSceneAxis = rightValueAxisSource.SceneAxis;
        if (rightValueAxis is null && rightSceneAxis is null)
        {
            return Array.Empty<PdfFontResource>();
        }

        ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(rightSceneAxis, rightValueAxis, fallback);
        ChartValueAxisRenderOptions axisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(rightSceneAxis, rightValueAxis, theme, extents, percentStacked: false);
        return RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, rightValueAxis, rightSceneAxis, extents, axisOptions.Units, axisOptions.Reversed, horizontalBars: false, rightSide: true, fontResolver: fontResolver);
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
        double? crossesAt = sceneAxis is not null
            ? sceneAxis.CrossesAt
            : ReadChartElementDouble(valueAxis, "crossesAt");
        if (crossesAt is { } explicitCrossing)
        {
            return explicitCrossing;
        }

        PptxSceneChartAxisCrosses crosses = sceneAxis is not null
            ? ResolveChartAxisCrosses(sceneAxis.CrossesKind)
            : ResolveChartAxisCrosses(PptxSceneBuilder.ParseChartAxisCrosses((string?)valueAxis?.Element(ChartNamespace + "crosses")?.Attribute("val") ?? "autoZero"));
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

    private static PptxSceneChartAxisCrosses ResolveChartAxisCrosses(PptxSceneChartAxisCrosses crosses)
    {
        return crosses == PptxSceneChartAxisCrosses.Unknown
            ? PptxSceneChartAxisCrosses.AutoZero
            : crosses;
    }

    private static bool ReadSceneOrXmlValueAxisReversed(PptxSceneChartAxis? sceneAxis, XElement? valueAxis)
    {
        if (sceneAxis is not null)
        {
            return ResolveChartAxisReversed(sceneAxis.OrientationKind);
        }

        return ResolveChartAxisReversed(PptxSceneBuilder.ParseChartAxisOrientation(
            (string?)valueAxis?.Element(ChartNamespace + "scaling")?.Element(ChartNamespace + "orientation")?.Attribute("val")));
    }

    private static bool ResolveChartAxisReversed(PptxSceneChartAxisOrientation orientation)
    {
        return orientation == PptxSceneChartAxisOrientation.MaximumMinimum;
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
        ChartNumberFormat numberFormat = ReadSceneOrXmlChartAxisNumberFormat(sceneAxis, axis);
        if (IsRenderableChartNumberFormat(numberFormat))
        {
            return FormatChartNumber(value, numberFormat.FormatCode);
        }

        return !string.IsNullOrWhiteSpace(defaultNumberFormat)
            ? FormatChartNumber(value, defaultNumberFormat)
            : FormatChartAxisLabel(value);
    }

    private static ChartNumberFormat ReadSceneOrXmlChartAxisNumberFormat(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is not null)
        {
            return ToChartNumberFormat(sceneAxis.NumberFormatInfo);
        }

        return axis is null ? default : ReadChartNumberFormat(axis);
    }

    private static bool IsRenderableChartNumberFormat(ChartNumberFormat numberFormat)
    {
        return numberFormat.IsDefined &&
            IsRenderableChartFormatCode(numberFormat.FormatCode);
    }

    private static bool IsRenderableChartFormatCode(string? formatCode)
    {
        return !string.IsNullOrWhiteSpace(formatCode) &&
            !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase);
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

    private static IReadOnlyList<ChartSeriesFill?> ReadChartSeriesFills(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        var fills = new List<ChartSeriesFill?>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? shapeProperties = element.Element(ChartNamespace + "spPr");
            fills.Add(TryReadChartFill(shapeProperties, theme, colorMap));
        }

        return fills;
    }

    private static IReadOnlyDictionary<int, ChartSeriesFill> ReadChartPointFills(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
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
            if (!TryReadChartNonNegativeIndex(point, out int index))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            if (TryReadChartFill(shapeProperties, theme, colorMap) is { } fill)
            {
                fills[index] = fill;
            }
        }

        return fills;
    }

    private static ChartSeriesFill? TryReadChartFill(XElement? shapeProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        if (TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor color, out double alpha))
        {
            return new ChartSeriesFill(color, alpha);
        }

        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        if (patternFill is null)
        {
            return null;
        }

        RgbColor foreground = TryReadSolidColor(patternFill.Element(DrawingNamespace + "fgClr"), theme, colorMap, out RgbColor foregroundColor)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = TryReadSolidColor(patternFill.Element(DrawingNamespace + "bgClr"), theme, colorMap, out RgbColor backgroundColor)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        string preset = (string?)patternFill.Attribute("prst") ?? "pct50";
        return new ChartSeriesFill(foreground, 1d, preset, background);
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> ReadChartSeriesPointFills(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => ReadChartPointFills(series, theme, colorMap))
            .ToArray();
    }

    private static IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> ReadChartSeriesPointStrokes(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => ReadChartPointStrokes(series, theme, colorMap))
            .ToArray();
    }

    private static IReadOnlyDictionary<int, ChartSeriesStroke> ReadChartPointStrokes(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
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
            if (!TryReadChartNonNegativeIndex(point, out int index))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            if (shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, colorMap, out RgbColor color, out double width, out double alpha))
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
            if (!TryReadChartNonNegativeIndex(point, out int index))
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

    private static bool TryReadChartNonNegativeIndex(XElement element, out int index)
    {
        return int.TryParse(
                element.Element(ChartNamespace + "idx")?.Attribute("val")?.Value,
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out index) &&
            index >= 0;
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

    private static IReadOnlyList<ChartSeriesStroke?> ReadChartSeriesStrokes(XElement chartElement, PptxTheme theme, PptxColorMap colorMap, double? inheritedWidth = null)
    {
        var strokes = new List<ChartSeriesStroke?>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? shapeProperties = element.Element(ChartNamespace + "spPr");
            ChartSeriesStroke? stroke = null;
            if (shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, colorMap, out RgbColor color, out double width, out double alpha))
            {
                bool widthSpecified = shapeProperties.Element(DrawingNamespace + "ln")?.Attribute("w") is not null;
                stroke = new ChartSeriesStroke(
                    color,
                    alpha,
                    widthSpecified ? width : inheritedWidth ?? width);
            }

            strokes.Add(stroke);
        }

        return strokes;
    }

    private static IReadOnlyList<ChartMarkerStyle> ReadChartMarkerStyles(XElement chartElement, PptxTheme theme, PptxColorMap colorMap, PptxSceneChartPlotKind plotKind)
    {
        var styles = new List<ChartMarkerStyle>();
        bool chartMarkerEnabled = IsOoxmlBooleanElementEnabled(chartElement.Element(ChartNamespace + "marker"));
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? marker = element.Element(ChartNamespace + "marker");
            string symbol = (string?)marker?.Element(ChartNamespace + "symbol")?.Attribute("val") ??
                PptxChartMarkerMetricRules.ResolveDefaultSymbol(plotKind, chartMarkerEnabled, styles.Count);
            double size = PptxChartMarkerMetricRules.ResolveSize(
                (string?)marker?.Element(ChartNamespace + "size")?.Attribute("val"),
                plotKind,
                chartMarkerEnabled,
                marker is not null,
                marker?.Element(ChartNamespace + "spPr") is not null);
            XElement? shapeProperties = marker?.Element(ChartNamespace + "spPr");
            ChartSeriesFill? fill = TryReadSolidColorWithAlpha(shapeProperties, theme, colorMap, out RgbColor fillColor, out double fillAlpha)
                ? new ChartSeriesFill(fillColor, fillAlpha)
                : null;
            ChartSeriesStroke? stroke = shapeProperties is not null &&
                TryReadLineWithAlpha(shapeProperties, theme, colorMap, out RgbColor strokeColor, out double strokeWidth, out double strokeAlpha)
                    ? new ChartSeriesStroke(
                        strokeColor,
                        strokeAlpha,
                        shapeProperties.Element(DrawingNamespace + "ln")?.Attribute("w") is not null
                            ? strokeWidth
                            : ChartMarkerInheritedStrokeWidth)
                    : null;
            styles.Add(new ChartMarkerStyle(PptxSceneBuilder.ParseChartMarkerSymbol(symbol), symbol, (string?)marker?.Element(ChartNamespace + "size")?.Attribute("val"), size, fill, stroke, marker is not null));
        }

        return styles;
    }

    private static IReadOnlyList<ChartBooleanOption> ReadChartSeriesSmoothOptions(XElement chartElement)
    {
        var smooth = new List<ChartBooleanOption>();
        foreach (XElement element in chartElement.Elements(ChartNamespace + "ser"))
        {
            XElement? smoothElement = element.Element(ChartNamespace + "smooth");
            smooth.Add(new ChartBooleanOption(
                IsOoxmlBooleanElementEnabled(smoothElement),
                (string?)smoothElement?.Attribute("val") ?? string.Empty,
                smoothElement is not null));
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
            ReadSceneOrXmlChartGridlineStroke(sceneAxis, sceneAxis?.MajorGridlineLine ?? default, sceneAxis?.MajorGridlineStyleLine ?? default, xmlAxis?.Element(ChartNamespace + "majorGridlines"), theme),
            ReadSceneOrXmlChartGridlineStroke(sceneAxis, sceneAxis?.MinorGridlineLine ?? default, sceneAxis?.MinorGridlineStyleLine ?? default, xmlAxis?.Element(ChartNamespace + "minorGridlines"), theme));
    }

    private static ChartValueAxisRenderOptions ReadSceneOrXmlChartValueAxisRenderOptions(
        PptxSceneChartAxis? sceneAxis,
        XElement? xmlAxis,
        PptxTheme theme,
        ChartValueExtents extents,
        bool percentStacked)
    {
        return new ChartValueAxisRenderOptions(
            ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(sceneAxis, xmlAxis), percentStacked),
            ReadSceneOrXmlValueAxisReversed(sceneAxis, xmlAxis),
            ReadSceneOrXmlValueAxisCrossingValue(sceneAxis, xmlAxis, extents),
            ReadSceneOrXmlMajorGridlines(sceneAxis, xmlAxis),
            ReadSceneOrXmlMinorGridlines(sceneAxis, xmlAxis),
            ReadSceneOrXmlChartGridlineStyle(sceneAxis, xmlAxis, theme));
    }

    private static ChartBubbleValueAxisOptions ReadSceneOrXmlChartBubbleValueAxisOptions(
        PptxSceneChartAxis? sceneAxis,
        XElement? xmlAxis,
        PptxTheme theme,
        ChartValueExtents extents)
    {
        return new ChartBubbleValueAxisOptions(
            ResolveBubbleAxisUnits(ReadSceneOrXmlChartValueAxisUnits(sceneAxis, xmlAxis), extents),
            ReadSceneOrXmlChartGridlineStyle(sceneAxis, xmlAxis, theme));
    }

    private static ChartSeriesStroke? ReadSceneOrXmlChartGridlineStroke(PptxSceneChartAxis? sceneAxis, PptxSceneLineStyle sceneLine, PptxSceneLineStyle sceneStyleLine, XElement? gridlines, PptxTheme theme)
    {
        if (sceneAxis is not null)
        {
            return ToChartSeriesStroke(sceneLine) ?? ToChartSeriesStroke(sceneStyleLine);
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

    private static ChartBooleanOption ReadChartVaryColorsOption(XElement chartElement)
    {
        XElement? varyColors = chartElement.Element(ChartNamespace + "varyColors");
        return new ChartBooleanOption(
            IsOoxmlBooleanElementEnabled(varyColors, defaultValue: true),
            (string?)varyColors?.Attribute("val") ?? string.Empty,
            varyColors is not null);
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

    private static ChartAxesStyle ReadSceneOrXmlChartAxesStyle(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, PptxTheme theme, XElement chartElement)
    {
        ChartAxisSource valueAxisSource = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, chartElement).FirstOrDefault();
        XElement? valueAxisElement = ResolveXmlValueAxisForSource(sceneChart, valueAxisSource, chartXml);
        ChartAxisSource categoryAxisSource = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, plot, chartXml, chartElement);
        XElement? categoryAxisElement = categoryAxisSource.XmlAxis;
        ChartAxisSource secondaryValueAxisSource = ReadSceneOrXmlSecondaryValueAxisForChart(sceneChart, chartXml, valueAxisSource);
        XElement? secondaryValueAxisElement = secondaryValueAxisSource.XmlAxis;
        PptxSceneChartAxis? valueAxis = valueAxisSource.SceneAxis;
        PptxSceneChartAxis? categoryAxis = categoryAxisSource.SceneAxis;
        PptxSceneChartAxis? secondaryValueAxis = secondaryValueAxisSource.SceneAxis;
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
        return sceneAxis is not null
            ? ToChartSeriesStroke(sceneAxis.Line)
            : ReadChartAxisStroke(xmlAxis, theme);
    }

    private static PptxSceneChartAxisTickMark ReadSceneOrXmlChartAxisMajorTickMark(PptxSceneChartAxis? sceneAxis, XElement? xmlAxis)
    {
        if (sceneAxis is not null)
        {
            return ResolveChartAxisTickMark(sceneAxis.MajorTickMarkKind);
        }

        return ResolveChartAxisTickMark(PptxSceneBuilder.ParseChartAxisTickMark((string?)xmlAxis?.Element(ChartNamespace + "majorTickMark")?.Attribute("val") ?? "none"));
    }

    private static PptxSceneChartAxisTickMark ResolveChartAxisTickMark(PptxSceneChartAxisTickMark tickMark)
    {
        return tickMark == PptxSceneChartAxisTickMark.Unknown
            ? PptxSceneChartAxisTickMark.None
            : tickMark;
    }

    private static PptxSceneChartTickLabelPosition ResolveChartTickLabelPosition(PptxSceneChartTickLabelPosition position)
    {
        return position == PptxSceneChartTickLabelPosition.Unknown
            ? PptxSceneChartTickLabelPosition.NextTo
            : position;
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

    private static bool ResolveSceneOrXmlCategoryAxisTopSide(PptxSceneChartAxis? sceneAxis, XElement? axis, bool defaultTopSide)
    {
        if (sceneAxis is not null)
        {
            return sceneAxis.PositionKind switch
            {
                PptxSceneChartAxisPosition.Top => true,
                PptxSceneChartAxisPosition.Bottom => false,
                _ => defaultTopSide
            };
        }

        string? position = (string?)axis?.Element(ChartNamespace + "axPos")?.Attribute("val");
        return PptxSceneBuilder.ParseChartAxisPosition(position) switch
        {
            PptxSceneChartAxisPosition.Top => true,
            PptxSceneChartAxisPosition.Bottom => false,
            _ => defaultTopSide
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
        return ResolveChartTickLabelPosition(PptxSceneBuilder.ParseChartTickLabelPosition(tickLabelPosition)) != PptxSceneChartTickLabelPosition.None;
    }

    private static bool IsSceneOrXmlChartAxisLabelVisible(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is null)
        {
            return IsChartAxisLabelVisible(axis);
        }

        return sceneAxis.IsDeleted != true &&
            ResolveChartTickLabelPosition(sceneAxis.TickLabelPositionKind) != PptxSceneChartTickLabelPosition.None;
    }

    private static bool ResolveValueAxisLabelsRightSide(XElement? axis, bool defaultRightSide)
    {
        string? tickLabelPosition = (string?)axis?.Element(ChartNamespace + "tickLblPos")?.Attribute("val");
        return ResolveChartTickLabelPosition(PptxSceneBuilder.ParseChartTickLabelPosition(tickLabelPosition)) switch
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

        return ResolveChartTickLabelPosition(sceneAxis.TickLabelPositionKind) switch
        {
            PptxSceneChartTickLabelPosition.High => true,
            PptxSceneChartTickLabelPosition.Low => false,
            _ => defaultRightSide
        };
    }

    private static double ResolveSceneOrXmlCategoryAxisLabelOffsetScale(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        int offset = sceneAxis is not null
            ? sceneAxis.LabelOffset ?? PptxChartMetricRules.CategoryAxisDefaultLabelOffset
            : ReadChartElementInt(axis, "lblOffset") ?? PptxChartMetricRules.CategoryAxisDefaultLabelOffset;

        offset = Math.Clamp(
            offset,
            PptxChartMetricRules.CategoryAxisMinimumLabelOffset,
            PptxChartMetricRules.CategoryAxisMaximumLabelOffset);

        return offset / (double)PptxChartMetricRules.CategoryAxisDefaultLabelOffset;
    }

    private static int ResolveSceneOrXmlCategoryAxisTickLabelSkip(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        int skip = sceneAxis is not null
            ? sceneAxis.TickLabelSkip ?? PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip
            : ReadChartElementInt(axis, "tickLblSkip") ?? PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip;

        return Math.Max(PptxChartMetricRules.CategoryAxisDefaultTickLabelSkip, skip);
    }

    private static bool ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is not null)
        {
            return ResolveChartAxisCrossBetween(sceneAxis.CrossBetweenKind) == PptxSceneChartAxisCrossBetween.MidpointCategory;
        }

        string? crossBetween = (string?)axis?.Element(ChartNamespace + "crossBetween")?.Attribute("val");
        return ResolveChartAxisCrossBetween(PptxSceneBuilder.ParseChartAxisCrossBetween(crossBetween)) == PptxSceneChartAxisCrossBetween.MidpointCategory;
    }

    private static PptxSceneChartAxisCrossBetween ResolveChartAxisCrossBetween(PptxSceneChartAxisCrossBetween crossBetween)
    {
        return crossBetween == PptxSceneChartAxisCrossBetween.Unknown
            ? PptxSceneChartAxisCrossBetween.Between
            : crossBetween;
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

    private static IReadOnlyList<ChartIndexedScatterSeries> ReadScatterSeriesVectors(XElement chartElement, bool readBubbleSize, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true)
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

    private static void RenderBarChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartLayoutBox plotAreaBox, ChartPlotBox plotBox, IReadOnlyList<ChartIndexedNumberVector> series, bool horizontalBars, PptxSceneChartGrouping grouping, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool majorGridlines, bool minorGridlines, ChartGridlineStyle gridlineStyle, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents, ChartAxisUnits axisUnits, double? valueAxisCrossingValue, bool valueAxisReversed, bool valueAxisLabelsVisible, bool manualPlotLayoutApplied, bool varyColors, double gapWidthPercent, double overlapPercent)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> denseSeries = DensifyChartPointSeries(series);
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        {
            int categoryCount = Math.Max(1, denseSeries.Max(values => values.Count));
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
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotBox,
                        () => DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor));
                }
                else
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotBox,
                        () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor));
                }
            }

            if (majorGridlines)
            {
                if (horizontalBars)
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotBox,
                        () => DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major, valueAxisAutoTickTargetCount));
                }
                else
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotBox,
                        () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major));
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
                    RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(plotX, axisY, plotX + plotWidth, axisY));
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
                    RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                }
            }

            if (!horizontalBars && axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke)
            {
                if (secondaryValueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, secondaryValueAxisStroke);
                    double axisX = axesStyle.SecondaryValueAxisRightSide ? plotX + plotWidth : plotX;
                    RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                }
            }

            if (horizontalBars)
            {
                if (stacked)
                {
                    RenderStackedHorizontalBars(graphics, plotBox, theme, colorMap, chartPalette, plotX, plotY, plotWidth, plotHeight, denseSeries, categoryCount, valueExtents, valueAxisReversed, percentStacked, seriesFills, pointFills, pointStrokes, varyColors, gapWidthPercent);
                }
                else
                {
                    RenderClusteredHorizontalBars(graphics, plotBox, theme, colorMap, chartPalette, plotX, plotY, plotWidth, plotHeight, denseSeries, categoryCount, valueExtents, valueAxisReversed, zeroX, seriesFills, pointFills, pointStrokes, varyColors, gapWidthPercent, overlapPercent);
                }

                return;
            }

            if (stacked)
            {
                RenderStackedColumns(graphics, plotBox, theme, colorMap, chartPalette, plotX, plotY, plotWidth, plotHeight, denseSeries, categoryCount, valueExtents, valueAxisReversed, percentStacked, seriesFills, pointFills, pointStrokes, varyColors, gapWidthPercent);
                return;
            }

            double categoryWidth = plotWidth / categoryCount;
            double barWidth = GetClusteredBarWidth(categoryWidth, denseSeries.Count, gapWidthPercent);
            double step = GetClusteredBarStep(barWidth, overlapPercent);
            double clusterWidth = barWidth + Math.Max(0, denseSeries.Count - 1) * step;
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryX = plotX + category * categoryWidth + (categoryWidth - clusterWidth) / 2d;
                for (int seriesIndex = 0; seriesIndex < denseSeries.Count; seriesIndex++)
                {
                    IReadOnlyList<ChartIndexedNumberPoint?> values = denseSeries[seriesIndex];
                    if (category >= values.Count || values[category]?.Value is not { } value)
                    {
                        continue;
                    }

                    ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, denseSeries.Count, varyColors, seriesFills, pointFills, value);
                    double barX = categoryX + seriesIndex * step;
                    double valueY = ChartValueToPlotCoordinate(valueExtents, value, plotY, plotHeight, valueAxisReversed);
                    double barY = Math.Min(zeroY, valueY);
                    double barHeight = Math.Abs(valueY - zeroY);
                    FillChartRectangleInPlotClip(graphics, plotBox, barX, barY, barWidth, barHeight, fill);
                    StrokeChartPointRectangleInPlotClip(graphics, plotBox, seriesIndex, category, pointStrokes, barX, barY, barWidth, barHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, value));
                }
            }
        }
    }

    private static ChartLayout GetBarChartLayout(
        PptxDocument document,
        PptxTheme theme,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PptxColorMap colorMap,
        PptxSceneChartPlot? barPlot,
        XElement barChart,
        ChartBarPlotOptions barOptions,
        ChartWorkbookData? workbook = null,
        bool plotVisibleOnly = true,
        PresentationFontResolver? fontResolver = null)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        PptxSceneChartTextBodyProperties titleTextBodyProperties = ReadSceneOrXmlChartTitleTextBodyProperties(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
        ChartPlotLayout plotLayout = GetBarChartPlotLayout(theme, frame, chartXml, sceneChart, barPlot, barChart, title, legend, barOptions, workbook, plotVisibleOnly, fontResolver: fontResolver);
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, titleTextBodyProperties, legend);
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
        ChartBarPlotOptions barOptions,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly,
        bool ignoreManualPlotLayout = false,
        PresentationFontResolver? fontResolver = null)
    {
        bool hasTitle = !string.IsNullOrWhiteSpace(title);
        bool hasLegend = legend.Visible && !legend.Overlay;
        bool horizontalBars = barOptions.BarDirection == PptxSceneChartBarDirection.Bar;
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
        else if (hasTitle && !hasLegend && HasInsideValueAxisCrossing(sceneChart, barPlot, barChart, chartXml, barOptions, workbook, plotVisibleOnly))
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

        defaultPlotBox = AdjustBarChartPlotBoxForVisibleValueAxes(theme, defaultPlotBox, frame, chartXml, sceneChart, horizontalBars, fontResolver);
        defaultPlotBox = AdjustBarChartPlotBoxForStackedValueAxisLabels(theme, defaultPlotBox, frame, chartXml, sceneChart, barPlot, barChart, barOptions, workbook, plotVisibleOnly, fontResolver);
        defaultPlotBox = AdjustBarChartPlotBoxForDefaultAxisTitles(defaultPlotBox, frame, chartXml, sceneChart, horizontalBars, hasTitle, hasLegend);
        if (ignoreManualPlotLayout)
        {
            return ChartPlotLayout.FromPlotBox(defaultPlotBox);
        }

        ChartPlotBox manualDefaultPlotBox = horizontalBars && HasRecognizedManualPlotLayoutTarget(sceneChart, chartXml)
            ? GetHorizontalBarManualLayoutTargetDefaultPlotBox(frame, defaultPlotBox)
            : defaultPlotBox;
        if (!TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, manualDefaultPlotBox, out ChartPlotLayout manualPlotLayout))
        {
            return ChartPlotLayout.FromPlotBox(defaultPlotBox);
        }

        return ResolveBarManualPlotLayoutTarget(theme, chartXml, sceneChart, barPlot, barChart, manualPlotLayout, horizontalBars);
    }

    private static bool HasInsideValueAxisCrossing(PptxSceneChart? sceneChart, PptxSceneChartPlot? barPlot, XElement barChart, XDocument chartXml, ChartBarPlotOptions barOptions, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        IReadOnlyList<ChartIndexedNumberVector> seriesVectors = ReadSceneOrXmlChartSeriesVectors(barPlot, barChart, workbook, plotVisibleOnly);
        if (CountRenderableSeries(seriesVectors) == 0)
        {
            return false;
        }

        PptxSceneChartGrouping grouping = barOptions.Grouping;
        ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, barPlot, chartXml, barChart).FirstOrDefault();
        ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(
            valueAxis.SceneAxis,
            valueAxis.XmlAxis,
            GetBarChartValueExtents(seriesVectors, grouping),
            IsPercentStackedChartGrouping(grouping));
        double? crossing = ReadSceneOrXmlValueAxisCrossingValue(valueAxis.SceneAxis, valueAxis.XmlAxis, valueExtents);
        return crossing > valueExtents.Min + PptxChartMetricRules.AxisValueEpsilon &&
            crossing < valueExtents.Max - PptxChartMetricRules.AxisValueEpsilon;
    }

    private static bool HasRecognizedManualPlotLayoutTarget(PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is not null)
        {
            return sceneChart.PlotAreaLayout.LayoutTargetKind != PptxSceneChartManualLayoutTarget.Unknown;
        }

        string layoutTarget = (string?)chartXml
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "layout")
            ?.Element(ChartNamespace + "manualLayout")
            ?.Element(ChartNamespace + "layoutTarget")
            ?.Attribute("val") ?? string.Empty;
        return PptxSceneBuilder.ParseChartManualLayoutTarget(layoutTarget) != PptxSceneChartManualLayoutTarget.Unknown;
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
        ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
        if (IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
        {
            double labelOffsetScale = ResolveSceneOrXmlCategoryAxisLabelOffsetScale(categoryAxis.SceneAxis, categoryAxis.XmlAxis);
            double outsideFactor =
                PptxChartMetricRules.CategoryAxisHorizontalLeftOffsetRatio * labelOffsetScale +
                PptxChartMetricRules.CategoryAxisHorizontalWidthRatio;
            leftReserve = plotAreaBox.Width * outsideFactor / (1d + outsideFactor);
        }

        double rightReserve = ReadSceneOrXmlChartAxisMajorTickMark(categoryAxis.SceneAxis, categoryAxis.XmlAxis) == PptxSceneChartAxisTickMark.Outside
            ? PptxChartMetricRules.CategoryAxisMajorTickLength + Math.Max(3d, PptxChartMetricRules.ValueAxisFallbackFontSize * 0.35d)
            : 0d;
        double topReserve = 0d;
        ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, barPlot, chartXml, barChart).FirstOrDefault();
        if (IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
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

    private static ChartPlotBox AdjustBarChartPlotBoxForVisibleValueAxes(PptxTheme theme, ChartPlotBox plotBox, ChartFrameBox frame, XDocument chartXml, PptxSceneChart? sceneChart, bool horizontalBars, PresentationFontResolver? fontResolver)
    {
        if (horizontalBars)
        {
            return plotBox;
        }

        IReadOnlyList<PptxSceneChartPlot> barPlots = ReadSceneChartPlots(sceneChart, PptxSceneChartPlotKind.Bar);
        IReadOnlyList<XElement> barCharts = ReadSceneOrXmlChartPlotElements(sceneChart, chartXml, PptxSceneChartPlotKind.Bar);
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
                defaultNumberFormat: null,
                fontResolver: fontResolver);
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

    private static ChartPlotBox AdjustBarChartPlotBoxForDefaultAxisTitles(
        ChartPlotBox plotBox,
        ChartFrameBox frame,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        bool horizontalBars,
        bool hasChartTitle,
        bool hasLegend)
    {
        if (hasChartTitle || hasLegend)
        {
            return plotBox;
        }

        ChartAxisTitleReserveSides reserveSides = ReadSceneOrXmlDefaultAxisTitleReserveSides(sceneChart, chartXml);
        if (!reserveSides.HasHorizontalTitle || !reserveSides.HasVerticalTitle)
        {
            return plotBox;
        }

        if (horizontalBars)
        {
            double horizontalLeftReserve = frame.Width * (reserveSides.Left
                ? PptxChartMetricRules.DefaultAxisTitleHorizontalBarPlotSideReserveRatio
                : PptxChartMetricRules.DefaultAxisTitleHorizontalBarPlotOppositeSideReserveRatio);
            double horizontalRightReserve = frame.Width * (reserveSides.Right
                ? PptxChartMetricRules.DefaultAxisTitleHorizontalBarPlotSideReserveRatio
                : PptxChartMetricRules.DefaultAxisTitleHorizontalBarPlotOppositeSideReserveRatio);
            double horizontalBottomReserve = frame.Height * PptxChartMetricRules.DefaultAxisTitleHorizontalBarPlotBandReserveRatio;
            double horizontalTopReserve = frame.Height * PptxChartMetricRules.DefaultAxisTitleHorizontalBarPlotBandReserveRatio;
            double horizontalX = frame.X + horizontalLeftReserve;
            double horizontalY = frame.Y + horizontalTopReserve;
            double horizontalWidth = Math.Max(1d, frame.Width - horizontalLeftReserve - horizontalRightReserve);
            double horizontalHeight = Math.Max(1d, frame.Height - horizontalTopReserve - horizontalBottomReserve);
            return new ChartPlotBox(horizontalX, horizontalY, horizontalWidth, horizontalHeight);
        }

        double leftReserve = frame.Width * (reserveSides.Left
            ? PptxChartMetricRules.DefaultAxisTitlePlotSideReserveRatio
            : PptxChartMetricRules.DefaultAxisTitlePlotOppositeSideReserveRatio);
        double rightReserve = frame.Width * (reserveSides.Right
            ? PptxChartMetricRules.DefaultAxisTitlePlotSideReserveRatio
            : PptxChartMetricRules.DefaultAxisTitlePlotOppositeSideReserveRatio);
        double bottomReserve = frame.Height * (reserveSides.Bottom
            ? PptxChartMetricRules.DefaultAxisTitlePlotBandReserveRatio
            : PptxChartMetricRules.DefaultAxisTitlePlotOppositeBandReserveRatio);
        double topReserve = frame.Height * (reserveSides.Top
            ? PptxChartMetricRules.DefaultAxisTitlePlotBandReserveRatio
            : PptxChartMetricRules.DefaultAxisTitlePlotOppositeBandReserveRatio);
        double x = frame.X + leftReserve;
        double y = frame.Y + bottomReserve;
        double width = Math.Max(1d, frame.Width - leftReserve - rightReserve);
        double height = Math.Max(1d, frame.Height - bottomReserve - topReserve);
        return new ChartPlotBox(x, y, width, height);
    }

    private static ChartPlotBox AdjustBarChartPlotBoxForStackedValueAxisLabels(
        PptxTheme theme,
        ChartPlotBox plotBox,
        ChartFrameBox frame,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        PptxSceneChartPlot? barPlot,
        XElement barChart,
        ChartBarPlotOptions barOptions,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly,
        PresentationFontResolver? fontResolver)
    {
        bool horizontalBars = barOptions.BarDirection == PptxSceneChartBarDirection.Bar;
        if (horizontalBars)
        {
            return plotBox;
        }

        PptxSceneChartGrouping grouping = barOptions.Grouping;
        bool percentStacked = IsPercentStackedChartGrouping(grouping);
        if (!IsStackedChartGrouping(grouping))
        {
            return plotBox;
        }

        ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, barPlot, chartXml, barChart).FirstOrDefault();
        if (!IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
        {
            return plotBox;
        }

        IReadOnlyList<ChartIndexedNumberVector> seriesVectors = ReadSceneOrXmlChartSeriesVectors(barPlot, barChart, workbook, plotVisibleOnly);
        if (CountRenderableSeries(seriesVectors) == 0)
        {
            return plotBox;
        }

        ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(
            valueAxis.SceneAxis,
            valueAxis.XmlAxis,
            GetBarChartValueExtents(seriesVectors, grouping),
            percentStacked);
        ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxis.XmlAxis), percentStacked);
        string? defaultNumberFormat = percentStacked ? "0%" : null;
        double requiredReserve = EstimateVerticalValueAxisLabelStripWidth(theme, sceneChart, chartXml, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, axisUnits, defaultNumberFormat, fontResolver);
        double leftReserve = plotBox.X - frame.X;
        double rightReserve = frame.X + frame.Width - plotBox.X - plotBox.Width;
        bool labelsRight = ResolveSceneOrXmlValueAxisLabelsRightSide(valueAxis.SceneAxis, valueAxis.XmlAxis, defaultRightSide: false);
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

    private static double EstimateVerticalValueAxisLabelStripWidth(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml, XElement? valueAxis, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits units, string? defaultNumberFormat, PresentationFontResolver? fontResolver)
    {
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
        double fontSize = style.FontSize;
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        IReadOnlyList<double> tickValues = GetChartAxisTickValues(extents, units.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount);
        double maxLabelWidth = tickValues
            .Select(value => FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis, defaultNumberFormat))
            .DefaultIfEmpty("0")
            .Max(label => textMeasurer.Measure(label, style));
        double labelWidth = Math.Max(
            fontSize * PptxChartMetricRules.ValueAxisMinimumLabelWidthFactor,
            maxLabelWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
        double sideGap = Math.Max(3d, fontSize * PptxChartMetricRules.ValueAxisLabelSideGapFactor);
        return labelWidth + sideGap;
    }

    private static ChartValueExtents GetBarChartValueExtents(IReadOnlyList<ChartIndexedNumberVector> series, PptxSceneChartGrouping grouping)
    {
        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> denseSeries = DensifyChartPointSeries(series);
        int categoryCount = Math.Max(1, denseSeries.Max(values => values.Count));
        bool stacked = IsStackedChartGrouping(grouping);
        bool percentStacked = IsPercentStackedChartGrouping(grouping);
        (double min, double max) = stacked
            ? GetStackedPointValueExtents(denseSeries, categoryCount, percentStacked)
            : GetClusteredPointValueExtents(denseSeries);
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

    private static ChartSeriesFill ChartSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha = 1d)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, seriesIndex), defaultAlpha);
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

    private static ChartSeriesFill ChartCategoryOrSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null)
            ? new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, categoryIndex), 1d)
            : ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills);
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

    private static ChartSeriesFill ChartPointCategoryOrSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills)
    {
        if (seriesIndex < pointFills.Count && pointFills[seriesIndex].TryGetValue(categoryIndex, out ChartSeriesFill pointFill))
        {
            return pointFill;
        }

        return ChartCategoryOrSeriesColor(theme, colorMap, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills);
    }

    private static ChartSeriesFill ResolveBarPointFill(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, double value)
    {
        if (value < 0d && !HasExplicitChartPointFill(pointFills, seriesIndex, categoryIndex))
        {
            return new ChartSeriesFill(new RgbColor(255, 255, 255), 1d);
        }

        return ChartPointCategoryOrSeriesColor(theme, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, pointFills);
    }

    private static ChartSeriesFill ResolveBarPointFill(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, double value)
    {
        if (value < 0d && !HasExplicitChartPointFill(pointFills, seriesIndex, categoryIndex))
        {
            return new ChartSeriesFill(new RgbColor(255, 255, 255), 1d);
        }

        return ChartPointCategoryOrSeriesColor(theme, colorMap, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, pointFills);
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

    private static void FillChartRectangleInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, double x, double y, double width, double height, ChartSeriesFill fill)
    {
        RenderInChartPlotAreaClip(graphics, plotBox, () => FillChartRectangle(graphics, x, y, width, height, fill));
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

    private static (double Min, double Max) GetClusteredPointValueExtents(IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series)
    {
        double maxValue = Math.Max(0d, series.SelectMany(points => points).Select(point => point?.Value).Where(value => value is not null).Select(value => value!.Value).DefaultIfEmpty(0d).Max());
        double minValue = Math.Min(0d, series.SelectMany(points => points).Select(point => point?.Value).Where(value => value is not null).Select(value => value!.Value).DefaultIfEmpty(0d).Min());
        return (minValue, maxValue);
    }

    private static (double Min, double Max) GetStackedPointValueExtents(IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, int categoryCount, bool percentStacked)
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
            foreach (IReadOnlyList<ChartIndexedNumberPoint?> values in series)
            {
                if (category >= values.Count || values[category]?.Value is not { } value)
                {
                    continue;
                }

                if (value >= 0d)
                {
                    positive += value;
                }
                else
                {
                    negative += value;
                }
            }

            maxValue = Math.Max(maxValue, positive);
            minValue = Math.Min(minValue, negative);
        }

        return (minValue, maxValue);
    }

    private static void RenderClusteredHorizontalBars(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, double zeroX, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent, double overlapPercent)
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
                IReadOnlyList<ChartIndexedNumberPoint?> values = series[seriesIndex];
                if (category >= values.Count || values[category]?.Value is not { } value)
                {
                    continue;
                }

                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, value);
                double valueX = ChartValueToPlotCoordinate(valueExtents, value, plotX, plotWidth, valueAxisReversed);
                double barX = Math.Min(zeroX, valueX);
                double barWidth = Math.Abs(valueX - zeroX);
                double barY = categoryY + seriesIndex * step;
                FillChartRectangleInPlotClip(graphics, plotBox, barX, barY, barWidth, barHeight, fill);
                StrokeChartPointRectangleInPlotClip(graphics, plotBox, seriesIndex, category, pointStrokes, barX, barY, barWidth, barHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, value));
            }
        }
    }

    private static void RenderStackedColumns(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent)
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
                IReadOnlyList<ChartIndexedNumberPoint?> values = series[seriesIndex];
                if (category >= values.Count || values[category]?.Value is not { } rawValue)
                {
                    continue;
                }

                double value = NormalizeStackedValue(rawValue, positiveTotal, percentStacked);
                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, value);
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
                FillChartRectangleInPlotClip(graphics, plotBox, categoryX, segmentY, barWidth, segmentHeight, fill);
                StrokeChartPointRectangleInPlotClip(graphics, plotBox, seriesIndex, category, pointStrokes, categoryX, segmentY, barWidth, segmentHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, value));
            }
        }
    }

    private static void RenderStackedHorizontalBars(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent)
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
                IReadOnlyList<ChartIndexedNumberPoint?> values = series[seriesIndex];
                if (category >= values.Count || values[category]?.Value is not { } rawValue)
                {
                    continue;
                }

                double value = NormalizeStackedValue(rawValue, positiveTotal, percentStacked);
                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, value);
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
                FillChartRectangleInPlotClip(graphics, plotBox, segmentX, categoryY, segmentWidth, barHeight, fill);
                StrokeChartPointRectangleInPlotClip(graphics, plotBox, seriesIndex, category, pointStrokes, segmentX, categoryY, segmentWidth, barHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, value));
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

    private static double GetCategoryPositiveTotal(IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, int category, bool percentStacked)
    {
        if (!percentStacked)
        {
            return 1d;
        }

        double total = 0d;
        foreach (IReadOnlyList<ChartIndexedNumberPoint?> values in series)
        {
            if (category < values.Count && values[category]?.Value is { } value)
            {
                total += Math.Max(0d, value);
            }
        }

        return Math.Max(1d, total);
    }

    private static double NormalizeStackedValue(double value, double positiveTotal, bool percentStacked)
    {
        return percentStacked && value > 0d ? value / positiveTotal : value;
    }

    private static void RenderLineChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartLayoutBox plotAreaBox, ChartPlotBox plotBox, IReadOnlyList<ChartIndexedNumberVector> series, bool stacked, bool percentStacked, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, IReadOnlyList<ChartBooleanOption> smoothSeries, bool majorGridlines, bool minorGridlines, ChartGridlineStyle gridlineStyle, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents, ChartAxisUnits axisUnits, double? valueAxisCrossingValue, bool valueAxisReversed, PptxSceneChartDisplayBlanksAs displayBlanksAs)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> denseSeries = DensifyChartPointSeries(series);
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        int pointCount = 0;
        double valueAxisCrossingY = 0d;
        ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
        {
            pointCount = Math.Max(1, denseSeries.Max(values => values.Count));
            double maxValue = valueExtents.Max;
            double minValue = valueExtents.Min;
            double valueRange = Math.Max(1d, maxValue - minValue);
            valueAxisCrossingY = ChartValueToPlotCoordinate(valueExtents, valueAxisCrossingValue, plotY, plotHeight, valueAxisReversed);

            if (minorGridlines)
            {
                RenderInChartPlotAreaClip(
                    graphics,
                    plotBox,
                    () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor));
            }

            if (majorGridlines)
            {
                RenderInChartPlotAreaClip(
                    graphics,
                    plotBox,
                    () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major));
            }

            ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
            if (axesStyle.CategoryAxisVisible)
            {
                if (categoryAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, categoryAxisStroke);
                    RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(plotX, valueAxisCrossingY, plotX + plotWidth, valueAxisCrossingY));
                }
            }

            if (axesStyle.ValueAxisVisible)
            {
                if (valueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, valueAxisStroke);
                    double axisX = axesStyle.ValueAxisRightSide ? plotX + plotWidth : plotX;
                    RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                }

                if (axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke)
                {
                    if (secondaryValueAxisStroke.Alpha > 0.001d)
                    {
                        SetChartStroke(graphics, secondaryValueAxisStroke);
                        double axisX = axesStyle.SecondaryValueAxisRightSide ? plotX + plotWidth : plotX;
                        RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                    }
                }
            }

            double[] lower = new double[pointCount];
            for (int seriesIndex = 0; seriesIndex < denseSeries.Count; seriesIndex++)
            {
                IReadOnlyList<ChartIndexedNumberPoint?> values = denseSeries[seriesIndex];
                if (values.Count == 0)
                {
                    continue;
                }

                ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, colorMap, chartPalette, seriesIndex, seriesStrokes, ChartLineDefaultStrokeWidth);
                if (stroke.Alpha < 1d)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(1d, stroke.Alpha);
                }

                SetChartStroke(graphics, stroke);
                var points = new List<(double X, double Y)>(values.Count);
                var markers = new List<(double X, double Y)>(values.Count);
                for (int i = 0; i < values.Count; i++)
                {
                    if (values[i]?.Value is not { } value)
                    {
                        if (displayBlanksAs == PptxSceneChartDisplayBlanksAs.Zero)
                        {
                            value = 0d;
                        }
                        else
                        {
                            if (displayBlanksAs != PptxSceneChartDisplayBlanksAs.Span)
                            {
                                StrokeLineChartPointSegmentInPlotClip(graphics, plotBox, points, IsSmoothSeries(seriesIndex, smoothSeries));
                                points.Clear();
                            }

                            continue;
                        }
                    }

                    double pointX = plotX + plotWidth * (i + 0.5d) / pointCount;
                    double positiveTotal = GetCategoryPositiveTotal(denseSeries, i, percentStacked);
                    double normalizedValue = NormalizeStackedValue(value, positiveTotal, percentStacked);
                    double plottedValue = stacked ? lower[i] + normalizedValue : value;
                    double pointY = ChartValueToPlotCoordinate(valueExtents, plottedValue, plotY, plotHeight, valueAxisReversed);
                    points.Add((pointX, pointY));
                    markers.Add((pointX, pointY));
                    if (stacked)
                    {
                        lower[i] = plottedValue;
                    }
                }

                StrokeLineChartPointSegmentInPlotClip(graphics, plotBox, points, IsSmoothSeries(seriesIndex, smoothSeries));

                foreach ((double pointX, double pointY) in markers)
                {
                    graphics.SetFillRgb(stroke.Color.Red, stroke.Color.Green, stroke.Color.Blue);
                    DrawChartMarkerInPlotClip(graphics, plotBox, pointX, pointY, ChartMarker(seriesIndex, markerStyles), stroke.Color, stroke.Color);
                }

                if (stroke.Alpha < 1d)
                {
                    graphics.RestoreState();
                }
            }
        }
        if (axesStyle.CategoryAxisVisible && categoryAxisStroke.Alpha > 0.001d)
        {
            SetChartStroke(graphics, categoryAxisStroke);
            DrawLineChartCategoryAxisMajorTicks(graphics, plotX, plotWidth, pointCount, valueAxisCrossingY, axesStyle.CategoryAxisMajorTickMark);
        }
    }

    private static void StrokeLineChartPointSegment(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points, bool smooth)
    {
        if (points.Count < 2)
        {
            return;
        }

        if (smooth)
        {
            StrokeSmoothChartPath(graphics, points);
        }
        else
        {
            StrokeStraightChartPath(graphics, points);
        }
    }

    private static void StrokeLineChartPointSegmentInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, IReadOnlyList<(double X, double Y)> points, bool smooth)
    {
        if (points.Count < 2)
        {
            return;
        }

        RenderInChartPlotAreaClip(graphics, plotBox, () => StrokeLineChartPointSegment(graphics, points, smooth));
    }

    private static ChartLayout GetLineChartLayout(PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, PptxColorMap colorMap, ChartWorkbookData? workbook = null, bool plotVisibleOnly = true, PresentationFontResolver? fontResolver = null)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        PptxSceneChartTextBodyProperties titleTextBodyProperties = ReadSceneOrXmlChartTitleTextBodyProperties(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
        ChartTextStyle legendTextStyle = ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml);
        ChartPlotLayout plotLayout = GetLineChartPlotLayout(frame, theme, chartXml, sceneChart, title, legend, legendTextStyle, workbook, plotVisibleOnly, fontResolver);
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, titleTextBodyProperties, legend);
    }

    private static ChartPlotLayout GetLineChartPlotLayout(ChartFrameBox frame, PptxTheme theme, XDocument chartXml, PptxSceneChart? sceneChart, string? title, ChartLegendLayout legend, ChartTextStyle legendTextStyle, ChartWorkbookData? workbook, bool plotVisibleOnly, PresentationFontResolver? fontResolver)
    {
        bool hasTitle = !string.IsNullOrWhiteSpace(title);
        bool hasRightLegend = legend.Visible && !legend.Overlay && legend.PositionKind == PptxSceneChartLegendPosition.Right;
        bool hasLineChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Line) is not null;
        ChartPlotBox defaultPlotBox = !hasTitle && hasRightLegend
            ? GetCartesianNoTitleRightLegendPlotBox(frame, theme, chartXml, sceneChart, workbook, plotVisibleOnly, fontResolver, legendTextStyle)
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

    private static ChartPlotBox GetCartesianNoTitleRightLegendPlotBox(ChartFrameBox frame, PptxTheme theme, XDocument chartXml, PptxSceneChart? sceneChart, ChartWorkbookData? workbook, bool plotVisibleOnly, PresentationFontResolver? fontResolver, ChartTextStyle legendTextStyle)
    {
        XElement? plotElement = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Line);
        PptxSceneChartPlotKind plotKind = PptxSceneChartPlotKind.Line;
        if (plotElement is null)
        {
            plotElement = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Area);
            plotKind = PptxSceneChartPlotKind.Area;
        }
        if (plotElement is null)
        {
            plotElement = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Scatter);
            plotKind = PptxSceneChartPlotKind.Scatter;
        }

        if (plotElement is null)
        {
            return GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.LineNoTitleRightLegend);
        }

        PptxSceneChartPlot? plot = ReadSceneChartPlot(sceneChart, plotKind);
        IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSceneOrXmlChartSeriesNameRecords(plot, plotElement, workbook);
        ChartRightLegendReserve rightLegendReserve = ResolveRightLegendReserve(
            frame,
            seriesNames,
            legendTextStyle,
            includeAreaReserve: plotKind == PptxSceneChartPlotKind.Area,
            fontResolver: fontResolver);
        var textMeasurer = new ChartTextMeasurer(fontResolver);

        double maxValueLabelWidth = 0d;
        int maxValueLabelLength = 0;
        bool explicitValueAxisScale = false;
        if (plotKind == PptxSceneChartPlotKind.Scatter)
        {
            IReadOnlyList<ScatterSeries> series = ReadSceneOrXmlScatterSeries(plot, plotElement, readBubbleSize: false, workbook: workbook, plotVisibleOnly: plotVisibleOnly);
            if (series.Count > 0)
            {
                IReadOnlyList<ChartAxisSource> valueAxes = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, plotElement);
                ChartAxisSource valueAxis = valueAxes.Count > 1
                    ? valueAxes[1]
                    : valueAxes.Count > 0
                        ? valueAxes[0]
                        : sceneChart is null
                            ? new ChartAxisSource(null, chartXml.Descendants(ChartNamespace + "valAx").FirstOrDefault())
                            : default;
                explicitValueAxisScale = HasSceneOrXmlExplicitValueAxisScale(valueAxis.SceneAxis, valueAxis.XmlAxis);
                ChartValueExtents valueExtents = ReadSceneOrXmlBubbleChartValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetScatterYValueExtents(series));
                ChartAxisUnits axisUnits = ResolveBubbleAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxis.XmlAxis), valueExtents);
                IReadOnlyList<double> tickValues = GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true);
                ChartTextStyle valueAxisTextStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                string[] tickLabels = tickValues
                    .Select(value => FormatSceneOrXmlChartAxisLabel(value, valueAxis.SceneAxis, valueAxis.XmlAxis, defaultNumberFormat: null))
                    .ToArray();
                maxValueLabelWidth = tickLabels.Length == 0
                    ? 0d
                    : tickLabels.Max(label => textMeasurer.Measure(label, valueAxisTextStyle));
                maxValueLabelLength = tickLabels.Length == 0
                    ? 0
                    : tickLabels.Max(label => label.Length);
            }
        }
        else
        {
            PptxSceneChartGrouping grouping = ReadSceneOrXmlCartesianRightLegendGrouping(sceneChart, plot, chartXml, plotElement, plotKind);
            bool stacked = IsStackedChartGrouping(grouping);
            bool percentStacked = IsPercentStackedChartGrouping(grouping);
            IReadOnlyList<ChartIndexedNumberVector> seriesVectors = ReadSceneOrXmlChartSeriesVectors(plot, plotElement, workbook, plotVisibleOnly);
            if (CountRenderableSeries(seriesVectors) > 0)
            {
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, plotElement).FirstOrDefault();
                XElement? valueAxisForScale = ResolveXmlValueAxisForSource(sceneChart, valueAxis, chartXml);
                explicitValueAxisScale = HasSceneOrXmlExplicitValueAxisScale(valueAxis.SceneAxis, valueAxisForScale);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, valueAxisForScale, GetLineChartValueExtents(seriesVectors, stacked, percentStacked), percentStacked, useNearMaximumHeadroom: !percentStacked);
                ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxisForScale), percentStacked);
                IReadOnlyList<double> tickValues = GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true);
                ChartTextStyle valueAxisTextStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxisForScale, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                string[] tickLabels = tickValues
                    .Select(value => FormatSceneOrXmlChartAxisLabel(value, valueAxis.SceneAxis, valueAxisForScale, percentStacked ? "0%" : null))
                    .ToArray();
                maxValueLabelWidth = tickLabels.Length == 0
                    ? 0d
                    : tickLabels.Max(label => textMeasurer.Measure(label, valueAxisTextStyle));
                maxValueLabelLength = tickLabels.Length == 0
                    ? 0
                    : tickLabels.Max(label => label.Length);
            }
        }

        double leftInset = maxValueLabelWidth > 0d
            ? maxValueLabelWidth +
                Math.Min(
                    PptxChartMetricRules.LineRightLegendValueAxisPadding,
                    frame.Width * PptxChartMetricRules.LineRightLegendValueAxisFrameWidthPaddingRatio) +
                Math.Max(0, maxValueLabelLength - 3) * PptxChartMetricRules.LineRightLegendExtraValueLabelCharacterPadding
            : frame.Width * PptxChartMetricRules.LineNoTitleRightLegendPlotBoxXRatio;
        double x = frame.X + leftInset;
        double yRatio = explicitValueAxisScale
            ? PptxChartMetricRules.LineNoTitleRightLegendExplicitScalePlotBoxYRatio
            : PptxChartMetricRules.LineNoTitleRightLegendPlotBoxYRatio;
        double heightRatio = explicitValueAxisScale
            ? PptxChartMetricRules.LineNoTitleRightLegendExplicitScalePlotBoxHeightRatio
            : PptxChartMetricRules.LineNoTitleRightLegendPlotBoxHeightRatio;
        double y = frame.Y + frame.Height * yRatio;
        double width = Math.Max(1d, frame.Width - leftInset - rightLegendReserve.Width);
        double height = frame.Height * heightRatio;
        return new ChartPlotBox(x, y, width, height);
    }

    private static PptxSceneChartGrouping ReadSceneOrXmlCartesianRightLegendGrouping(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XDocument chartXml, XElement plotElement, PptxSceneChartPlotKind plotKind)
    {
        return plotKind switch
        {
            PptxSceneChartPlotKind.Area => ReadSceneOrXmlChartAreaOptions(sceneChart, plot, chartXml, plotElement, PptxSceneChartGrouping.Standard).Grouping,
            PptxSceneChartPlotKind.Line => ReadSceneOrXmlChartLineOptions(sceneChart, plot, chartXml, plotElement, PptxSceneChartGrouping.Standard).Grouping,
            _ => ReadSceneOrXmlChartGrouping(plot, plotElement, PptxSceneChartGrouping.Standard)
        };
    }

    private static bool HasSceneOrXmlExplicitValueAxisScale(PptxSceneChartAxis? axis, XElement? valueAxis)
    {
        if (axis is not null)
        {
            return axis.Minimum is not null &&
                axis.Maximum is not null &&
                axis.MajorUnit is not null;
        }

        return valueAxis?
            .Element(ChartNamespace + "scaling")
            ?.Element(ChartNamespace + "min") is not null &&
            valueAxis
                .Element(ChartNamespace + "scaling")
                ?.Element(ChartNamespace + "max") is not null &&
            valueAxis.Element(ChartNamespace + "majorUnit") is not null;
    }

    private static ChartLayout GetBubbleChartLayout(PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartPlot? bubblePlot, XElement bubbleChart, PptxColorMap colorMap, ChartWorkbookData? workbook = null, PresentationFontResolver? fontResolver = null)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        PptxSceneChartTextBodyProperties titleTextBodyProperties = ReadSceneOrXmlChartTitleTextBodyProperties(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
        ChartTextStyle legendTextStyle = ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml);
        ChartPlotLayout plotLayout = GetBubbleChartPlotLayout(frame, chartXml, sceneChart, bubblePlot, bubbleChart, title, legend, legendTextStyle, workbook, fontResolver);
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, titleTextBodyProperties, legend);
    }

    private static ChartPlotLayout GetBubbleChartPlotLayout(ChartFrameBox frame, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartPlot? bubblePlot, XElement bubbleChart, string? title, ChartLegendLayout legend, ChartTextStyle legendTextStyle, ChartWorkbookData? workbook, PresentationFontResolver? fontResolver)
    {
        bool hasTitle = !string.IsNullOrWhiteSpace(title);
        bool hasRightLegend = legend.Visible && !legend.Overlay && legend.PositionKind == PptxSceneChartLegendPosition.Right;
        ChartPlotBox defaultPlotBox = hasTitle && hasRightLegend
            ? GetBubbleTitleRightLegendPlotBox(frame, bubblePlot, bubbleChart, workbook, fontResolver, legendTextStyle)
            : GetDefaultChartPlotBox(frame);
        return TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout manualPlotLayout)
            ? manualPlotLayout
            : ChartPlotLayout.FromPlotBox(defaultPlotBox);
    }

    private static ChartPlotBox GetBubbleTitleRightLegendPlotBox(ChartFrameBox frame, PptxSceneChartPlot? bubblePlot, XElement bubbleChart, ChartWorkbookData? workbook, PresentationFontResolver? fontResolver, ChartTextStyle legendTextStyle)
    {
        IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSceneOrXmlChartSeriesNameRecords(bubblePlot, bubbleChart, workbook);
        ChartRightLegendReserve rightLegendReserve = ResolveRightLegendReserve(frame, seriesNames, legendTextStyle, includeAreaReserve: false, fontResolver: fontResolver);

        double x = frame.X + frame.Width * PptxChartMetricRules.LineTitleRightLegendPlotBoxXRatio;
        double y = frame.Y + frame.Height * PptxChartMetricRules.LineTitleRightLegendPlotBoxYRatio;
        double width = Math.Max(1d, frame.Width - (x - frame.X) - rightLegendReserve.Width);
        double height = frame.Height * PptxChartMetricRules.LineTitleRightLegendPlotBoxHeightRatio;
        return new ChartPlotBox(x, y, width, height);
    }

    private static ChartRightLegendReserve ResolveRightLegendReserve(ChartFrameBox frame, IReadOnlyList<ChartSeriesNameRecord> seriesNames, ChartTextStyle legendTextStyle, bool includeAreaReserve, PresentationFontResolver? fontResolver)
    {
        double legendFontSize = legendTextStyle.FontSize;
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        double maxLegendTextWidth = seriesNames.Count == 0
            ? 0d
            : seriesNames.Max(name => textMeasurer.Measure(name.ActiveName, legendTextStyle));
        int maxLegendTextLength = seriesNames.Count == 0
            ? 0
            : seriesNames.Max(name => name.ActiveName.Length);
        double rightReserve = maxLegendTextWidth +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeMarkerWidthFactor +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeTextGapFactor +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeGapFactor +
            PptxChartMetricRules.LineRightLegendReservePadding +
            Math.Max(0, maxLegendTextLength - 6) * PptxChartMetricRules.LineRightLegendExtraLegendCharacterPadding;
        if (includeAreaReserve)
        {
            rightReserve += frame.Width * PptxChartMetricRules.AreaRightLegendReserveFrameWidthFactor;
        }

        return new ChartRightLegendReserve(rightReserve, legendFontSize, maxLegendTextWidth, maxLegendTextLength, includeAreaReserve);
    }

    private static ChartValueExtents GetRadarChartValueExtents(IReadOnlyList<ChartRadarSeries> series)
    {
        IEnumerable<double> values = series
            .SelectMany(item => item.Points)
            .Where(point => point?.Value is not null)
            .Select(point => point!.Value!.Value!.Value);
        double maxValue = Math.Max(0d, values.DefaultIfEmpty(0d).Max());
        double minValue = Math.Min(0d, values.DefaultIfEmpty(0d).Min());
        return new ChartValueExtents(minValue, maxValue);
    }

    private static ChartValueExtents GetLineChartValueExtents(IReadOnlyList<ChartIndexedNumberVector> series, bool stacked, bool percentStacked)
    {
        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> denseSeries = DensifyChartPointSeries(series);
        int pointCount = Math.Max(1, denseSeries.Max(values => values.Count));
        (double minValue, double maxValue) = stacked
            ? GetStackedPointValueExtents(denseSeries, pointCount, percentStacked)
            : GetClusteredPointValueExtents(denseSeries);
        return new ChartValueExtents(minValue, maxValue);
    }

    private static bool IsSmoothSeries(int seriesIndex, IReadOnlyList<ChartBooleanOption> smoothSeries)
    {
        return seriesIndex < smoothSeries.Count && smoothSeries[seriesIndex].Value;
    }

    private static void StrokeStraightChartPath(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points)
    {
        if (points.Count < 2)
        {
            return;
        }

        graphics.MoveTo(points[0].X, points[0].Y);
        for (int i = 1; i < points.Count; i++)
        {
            graphics.LineTo(points[i].X, points[i].Y);
        }

        graphics.StrokeCurrentPath();
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
        ChartSeriesStroke? stroke = marker.Stroke ?? new ChartSeriesStroke(defaultStroke, 1d, Math.Max(0.75d, size * 0.16d));
        DrawChartMarkerFill(graphics, x, y, marker.SymbolKind, size, fill);
        DrawChartMarkerStroke(graphics, x, y, marker.SymbolKind, size, stroke);
    }

    private static void DrawChartMarkerInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, double x, double y, ChartMarkerStyle marker, RgbColor defaultFill, RgbColor defaultStroke)
    {
        if (marker.SymbolKind == PptxSceneChartMarkerSymbol.None)
        {
            return;
        }

        double size = marker.Size;
        ChartSeriesFill fill = marker.Fill ?? new ChartSeriesFill(defaultFill, 1d);
        ChartSeriesStroke? stroke = marker.Stroke ?? new ChartSeriesStroke(defaultStroke, 1d, Math.Max(0.75d, size * 0.16d));
        if (!IsLineOnlyChartMarker(marker.SymbolKind))
        {
            RenderInChartPlotAreaClip(graphics, plotBox, () => DrawChartMarkerFill(graphics, x, y, marker.SymbolKind, size, fill));
        }

        if (stroke is not null)
        {
            RenderInChartPlotAreaClip(graphics, plotBox, () => DrawChartMarkerStroke(graphics, x, y, marker.SymbolKind, size, stroke));
        }
    }

    private static void DrawChartMarkerFill(PdfGraphicsBuilder graphics, double x, double y, PptxSceneChartMarkerSymbol symbol, double size, ChartSeriesFill fill)
    {
        if (!IsLineOnlyChartMarker(symbol))
        {
            if (fill.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            switch (symbol)
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
    }

    private static void DrawChartMarkerStroke(PdfGraphicsBuilder graphics, double x, double y, PptxSceneChartMarkerSymbol symbol, double size, ChartSeriesStroke? stroke)
    {
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
        switch (symbol)
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

    private static ChartSeriesStroke ChartSeriesStrokeColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, double defaultWidth)
    {
        return seriesIndex < seriesStrokes.Count && seriesStrokes[seriesIndex] is { } stroke
            ? stroke
            : new ChartSeriesStroke(ChartPalette(chartPalette, theme, colorMap, seriesIndex), 1d, defaultWidth);
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
        PptxColorMap colorMap,
        IReadOnlyList<RgbColor>? chartPalette,
        ChartLayoutBox plotAreaBox,
        ChartPlotBox plotBox,
        IReadOnlyList<ChartIndexedNumberVector> series,
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
        bool valueAxisReversed,
        PptxSceneChartDisplayBlanksAs displayBlanksAs)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> denseSeries = DensifyChartPointSeries(series);
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        {
            if (minorGridlines)
            {
                RenderInChartPlotAreaClip(
                    graphics,
                    plotBox,
                    () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MinorUnit, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor));
            }

            if (majorGridlines)
            {
                RenderInChartPlotAreaClip(
                    graphics,
                    plotBox,
                    () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits.MajorUnit, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major));
            }

            RenderAreaChartSeries(graphics, theme, colorMap, chartPalette, plotBox, denseSeries, stacked, percentStacked, seriesFills, seriesStrokes, valueExtents, valueAxisReversed, displayBlanksAs);

            ChartSeriesStroke valueAxisStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
            ChartSeriesStroke categoryAxisStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
            double valueAxisCrossingY = ChartValueToPlotCoordinate(valueExtents, valueAxisCrossingValue, plotY, plotHeight, valueAxisReversed);
            if (axesStyle.CategoryAxisVisible && categoryAxisStroke.Alpha > 0.001d)
            {
                SetChartStroke(graphics, categoryAxisStroke);
                RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(plotX, valueAxisCrossingY, plotX + plotWidth, valueAxisCrossingY));
            }

            if (axesStyle.ValueAxisVisible)
            {
                if (valueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, valueAxisStroke);
                    double axisX = axesStyle.ValueAxisRightSide ? plotX + plotWidth : plotX;
                    RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                }

                if (axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke && secondaryValueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, secondaryValueAxisStroke);
                    double axisX = axesStyle.SecondaryValueAxisRightSide ? plotX + plotWidth : plotX;
                    RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                }
            }
        }
    }

    private static void RenderAreaChartSeries(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartPlotBox plotBox, IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, bool stacked, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents valueExtents, bool valueAxisReversed, PptxSceneChartDisplayBlanksAs displayBlanksAs)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        int pointCount = Math.Max(1, series.Max(values => values.Count));
        double[] lower = new double[pointCount];
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            IReadOnlyList<ChartIndexedNumberPoint?> values = series[seriesIndex];
            if (values.Count == 0)
            {
                continue;
            }

            if (displayBlanksAs == PptxSceneChartDisplayBlanksAs.Gap)
            {
                int start = 0;
                while (start < pointCount)
                {
                    while (start < pointCount && (start >= values.Count || values[start]?.Value is null))
                    {
                        start++;
                    }

                    int end = start;
                    while (end < pointCount && end < values.Count && values[end]?.Value is not null)
                    {
                        end++;
                    }

                    if (end - start >= 2)
                    {
                        RenderAreaChartSeriesSegment(graphics, theme, colorMap, chartPalette, plotBox, values, start, end, lower, stacked, percentStacked, seriesIndex, seriesFills, seriesStrokes, valueExtents, valueAxisReversed, series);
                    }

                    start = Math.Max(end, start + 1);
                }

                continue;
            }

            RenderAreaChartSeriesSegment(graphics, theme, colorMap, chartPalette, plotBox, values, 0, pointCount, lower, stacked, percentStacked, seriesIndex, seriesFills, seriesStrokes, valueExtents, valueAxisReversed, series);
        }
    }

    private static void RenderAreaChartSeriesSegment(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartPlotBox plotBox, IReadOnlyList<ChartIndexedNumberPoint?> values, int startIndex, int endIndex, double[] lower, bool stacked, bool percentStacked, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents valueExtents, bool valueAxisReversed, IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> allSeries)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        int pointCount = lower.Length;
        int segmentPointCount = Math.Max(0, endIndex - startIndex);
        if (segmentPointCount < 2)
        {
            return;
        }

        var upperPoints = new (double X, double Y)[segmentPointCount];
        var lowerPoints = new (double X, double Y)[segmentPointCount];
        for (int i = startIndex; i < endIndex; i++)
        {
            double pointX = plotX + (pointCount == 1 ? plotWidth / 2d : plotWidth * i / (pointCount - 1));
            double value = i < values.Count && values[i]?.Value is { } indexedValue ? indexedValue : 0d;
            double lowerValue = stacked ? lower[i] : 0d;
            double positiveTotal = GetCategoryPositiveTotal(allSeries, i, percentStacked);
            double normalizedValue = NormalizeStackedValue(value, positiveTotal, percentStacked);
            double upperValue = stacked ? lower[i] + normalizedValue : value;
            int segmentIndex = i - startIndex;
            upperPoints[segmentIndex] = (pointX, ChartValueToPlotCoordinate(valueExtents, upperValue, plotY, plotHeight, valueAxisReversed));
            lowerPoints[segmentIndex] = (pointX, ChartValueToPlotCoordinate(valueExtents, lowerValue, plotY, plotHeight, valueAxisReversed));
            if (stacked)
            {
                lower[i] = upperValue;
            }
        }

        var polygon = new (double X, double Y)[segmentPointCount * 2];
        for (int i = 0; i < segmentPointCount; i++)
        {
            polygon[i] = upperPoints[i];
            polygon[polygon.Length - i - 1] = lowerPoints[i];
        }

        ChartSeriesFill fill = ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills);
        if (fill.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
        }

        graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
        RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.FillPolygon(polygon));
        if (fill.Alpha < 1d)
        {
            graphics.RestoreState();
        }

        ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, colorMap, chartPalette, seriesIndex, seriesStrokes, ChartLineDefaultStrokeWidth);
        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.StrokePolygon(polygon));

        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

    private static void StrokeChartPointRectangleInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, int seriesIndex, int categoryIndex, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, double x, double y, double width, double height, ChartSeriesStroke? fallbackStroke = null)
    {
        bool hasExplicitStroke = seriesIndex < pointStrokes.Count && pointStrokes[seriesIndex].ContainsKey(categoryIndex);
        if (!hasExplicitStroke && fallbackStroke is null)
        {
            return;
        }

        RenderInChartPlotAreaClip(graphics, plotBox, () => StrokeChartPointRectangle(graphics, seriesIndex, categoryIndex, pointStrokes, x, y, width, height, fallbackStroke));
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

    private static void RenderScatterChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartPlotBox plotBox, IReadOnlyList<ScatterSeries> series, bool connectLines, bool bubble, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, IReadOnlyList<ChartBooleanOption> smoothSeries, ChartValueExtents? xValueExtents = null, ChartValueExtents? yValueExtents = null)
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

        {
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                ChartSeriesFill fill = ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills);
                ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, colorMap, chartPalette, seriesIndex, seriesStrokes, 1.2d);
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
                    (double pointX, double pointY, _) = ResolveScatterPointGeometry(plotBox, point, bubble, xExtents, yExtents, maxBubbleSize);
                    points.Add((pointX, pointY));
                }

                if (connectLines)
                {
                    StrokeScatterChartPathInPlotClip(graphics, plotBox, points, IsSmoothSeries(seriesIndex, smoothSeries));
                }

                foreach (ScatterPoint point in series[seriesIndex].Points)
                {
                    (double pointX, double pointY, double radius) = ResolveScatterPointGeometry(plotBox, point, bubble, xExtents, yExtents, maxBubbleSize);
                    if (bubble)
                    {
                        FillBubbleInPlotClip(graphics, plotBox, pointX, pointY, radius);
                    }
                    else
                    {
                        DrawChartMarkerInPlotClip(graphics, plotBox, pointX, pointY, ChartMarker(seriesIndex, markerStyles), fill.Color, stroke.Color);
                    }

                }

                if (fill.Alpha < 1d || stroke.Alpha < 1d)
                {
                    graphics.RestoreState();
                }
            }
        }
    }

    private static void StrokeScatterChartPathInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, IReadOnlyList<(double X, double Y)> points, bool smooth)
    {
        RenderInChartPlotAreaClip(graphics, plotBox, () =>
        {
            if (smooth)
            {
                StrokeSmoothChartPath(graphics, points);
            }
            else
            {
                StrokeStraightChartPath(graphics, points);
            }
        });
    }

    private static void FillBubbleInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, double pointX, double pointY, double radius)
    {
        RenderInChartPlotAreaClip(graphics, plotBox, () => graphics.FillEllipse(pointX - radius, pointY - radius, radius * 2d, radius * 2d));
    }

    private static void RenderRadarChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartRadarLayout layout, IReadOnlyList<ChartRadarSeries> series, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, ChartValueExtents extents, ChartAxisUnits axisUnits)
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
            IReadOnlyList<ChartIndexedNumberPoint?> values = series[seriesIndex].Points;
            var points = new (double X, double Y)[pointCount];
            for (int i = 0; i < pointCount; i++)
            {
                double value = i < values.Count && values[i]?.Value is { } pointValue ? Math.Max(0d, pointValue) : 0d;
                double pointRadius = GetChartValuePlotRatio(extents, value, false) * geometry.Radius;
                double angle = GetRadarPointAngle(i, pointCount);
                points[i] = (geometry.CenterX + Math.Cos(angle) * pointRadius, geometry.CenterY + Math.Sin(angle) * pointRadius);
            }

            if (layout.IsFilled)
            {
                ChartSeriesFill fill = ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills, series.Count == 1 ? 0.40d : 0.18d);
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
                graphics.SetFillRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
                graphics.FillPolygon(points);
                graphics.RestoreState();
            }
            ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, colorMap, chartPalette, seriesIndex, seriesStrokes, 1.2d);
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

    private static ChartValueExtents GetAreaChartValueExtents(IReadOnlyList<ChartIndexedNumberVector> series, bool stacked, bool percentStacked)
    {
        if (percentStacked)
        {
            return new ChartValueExtents(0d, 1d);
        }

        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> denseSeries = DensifyChartPointSeries(series);
        int pointCount = Math.Max(1, denseSeries.Max(values => values.Count));
        (double minValue, double maxValue) = stacked
            ? GetStackedPointValueExtents(denseSeries, pointCount, percentStacked)
            : GetClusteredPointValueExtents(denseSeries);
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
        ChartIndexedTextVector labelVector,
        PresentationFontResolver? fontResolver = null)
    {
        IReadOnlyList<ChartIndexedTextPoint?> labels = labelVector.DensePoints();
        if (labels.Count == 0)
        {
            return [];
        }

        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, categoryAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
        ChartPlotBox plotBox = layout.PlotBox;
        int pointCount = Math.Max(labels.Count, layout.PointCount);
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        var runs = new List<TextRun>(labels.Count);
        for (int i = 0; i < labels.Count; i++)
        {
            string? label = labels[i]?.Text;
            if (string.IsNullOrWhiteSpace(label))
            {
                continue;
            }

            ChartRadarLabelFrame frame = ResolveRadarCategoryLabelFrame(layout, label, style, textMeasurer, i, pointCount);
            runs.Add(CreateChartLabelRun(label, frame.X, frame.Y, frame.Width, frame.Height, plotBox, style, frame.Alignment));
        }

        return RenderTextRuns(runs, graphics, "RCA", fontResolver);
    }

    private static ChartRadarLabelFrame ResolveRadarCategoryLabelFrame(ChartRadarLayout layout, string label, ChartTextStyle style, ChartTextMeasurer textMeasurer, int index, int pointCount)
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
        double width = Math.Max(fontSize * 2d, textMeasurer.Measure(label, style) + fontSize);
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
        ChartAxisUnits axisUnits,
        PresentationFontResolver? fontResolver = null)
    {
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
        ChartPlotBox plotBox = layout.PlotBox;
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        var runs = new List<TextRun>();
        foreach (double tickValue in GetChartAxisTickValues(extents, axisUnits.MajorUnit, includeEndpoints: true))
        {
            double ratio = GetChartValuePlotRatio(extents, tickValue, false);
            string label = FormatSceneOrXmlChartAxisLabel(tickValue, sceneAxis, valueAxis);
            ChartRadarLabelFrame frame = ResolveRadarValueAxisLabelFrame(layout, label, style, textMeasurer, ratio);
            runs.Add(CreateChartLabelRun(label, frame.X, frame.Y, frame.Width, frame.Height, plotBox, style, frame.Alignment));
        }

        return RenderTextRuns(runs, graphics, "RVA", fontResolver);
    }

    private static ChartRadarLabelFrame ResolveRadarValueAxisLabelFrame(ChartRadarLayout layout, string label, ChartTextStyle style, ChartTextMeasurer textMeasurer, double ratio)
    {
        ChartPolarGeometry geometry = layout.Geometry;
        double fontSize = style.FontSize;
        double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
        ChartRadarLabelRules labelRules = layout.LabelRules;
        double width = Math.Max(
            fontSize * labelRules.ValueWidthFactor,
            textMeasurer.Measure(label, style) + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
        double x = geometry.CenterX - width - fontSize * labelRules.ValueGapFactor;
        double y = ResolveRadarValueAxisLabelBaselineY(layout, ratio, height);
        return new ChartRadarLabelFrame(x, y, width, height, TextAlignment.Right);
    }

    private static double ResolveRadarValueAxisLabelBaselineY(ChartRadarLayout layout, double ratio, double labelHeight)
    {
        return layout.Geometry.CenterY + layout.Geometry.Radius * ratio - labelHeight * layout.LabelRules.ValueBaselineOffsetFactor;
    }

    private static IReadOnlyList<ChartIndexedPieSlice> BuildChartIndexedPieSlices(ChartIndexedNumberVector values)
    {
        IReadOnlyDictionary<int, ChartIndexedNumberPoint> workbookPoints = values.WorkbookPointsForPlotVisibility(values.PlotVisibleOnly)
            .GroupBy(point => point.Index)
            .ToDictionary(group => group.Key, group => group.First());
        return values.DensePoints()
            .Where(point => point is not null)
            .Select(point => point!.Value)
            .Where(point => point.Value is > 0d)
            .OrderBy(point => point.Index)
            .Select(point => new ChartIndexedPieSlice(
                point.Index,
                point.Value!.Value,
                point,
                workbookPoints.TryGetValue(point.Index, out ChartIndexedNumberPoint workbookPoint) ? workbookPoint : null))
            .ToArray();
    }

    private static void RenderPieChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<ChartIndexedPieSlice> slices, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions, double firstSliceAngle)
    {
        RenderPieOrDoughnutSlices(graphics, theme, colorMap, chartPalette, layout, slices, pointFills, pointStrokes, pointExplosions, holeSize: 0d, firstSliceAngle);
    }

    private static void RenderPieOrDoughnutSlices(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<ChartIndexedPieSlice> slices, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, double firstSliceAngle)
    {
        double total = slices.Sum(slice => slice.Value);
        if (total <= 0d)
        {
            return;
        }

        ChartPolarGeometry geometry = layout.Geometry;
        double innerRadius = geometry.Radius * Math.Clamp(holeSize, 0d, 0.95d);
        double angle = Math.PI / 2d - firstSliceAngle * Math.PI / 180d;

        foreach (ChartIndexedPieSlice slice in slices)
        {
            double sweep = -slice.Value / total * Math.PI * 2d;
            double midpointAngle = angle + sweep / 2d;
            double explosionOffset = pointExplosions.TryGetValue(slice.Index, out double explosion) ? geometry.Radius * explosion : 0d;
            double sliceCenterX = geometry.CenterX + Math.Cos(midpointAngle) * explosionOffset;
            double sliceCenterY = geometry.CenterY + Math.Sin(midpointAngle) * explosionOffset;
            ChartSeriesFill fill = pointFills.TryGetValue(slice.Index, out ChartSeriesFill explicitFill)
                ? explicitFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, slice.Index), 1d);
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

            if (pointStrokes.TryGetValue(slice.Index, out ChartSeriesStroke stroke))
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

    private static void RenderDoughnutChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<ChartIndexedPieSlice> slices, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, ChartSeriesStroke> pointStrokes, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, double firstSliceAngle)
    {
        RenderPieOrDoughnutSlices(graphics, theme, colorMap, chartPalette, layout, slices, pointFills, pointStrokes, pointExplosions, holeSize, firstSliceAngle);
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

    private readonly record struct ScatterSeries(IReadOnlyList<ScatterPoint> Points, ChartIndexedScatterSeries Source);

    private readonly record struct ChartRadarSeries(IReadOnlyList<ChartIndexedNumberPoint?> Points, ChartIndexedNumberVector Source);

    private readonly record struct ScatterPoint(
        double X,
        double Y,
        double Size,
        int Index,
        ChartIndexedNumberPoint XPoint,
        ChartIndexedNumberPoint YPoint,
        ChartIndexedNumberPoint? BubbleSizePoint,
        ChartIndexedNumberPoint? XWorkbookPoint,
        ChartIndexedNumberPoint? YWorkbookPoint,
        ChartIndexedNumberPoint? BubbleSizeWorkbookPoint,
        string? YFormatCode,
        string? BubbleSizeFormatCode);

    private readonly record struct ChartIndexedScatterSeries(
        ChartIndexedNumberVector XValues,
        ChartIndexedNumberVector YValues,
        ChartIndexedNumberVector BubbleSizes,
        bool ReadBubbleSize);

    private readonly record struct ChartIndexedPieSlice(
        int Index,
        double Value,
        ChartIndexedNumberPoint Point,
        ChartIndexedNumberPoint? WorkbookPoint);

    private readonly record struct ChartIndexedNumberVector(
        IReadOnlyList<ChartIndexedNumberPoint> Points,
        int? PointCount,
        string? Formula,
        string? FormatCode,
        PptxSceneChartDataSource Source,
        IReadOnlyList<ChartIndexedNumberPoint> WorkbookPoints,
        bool PlotVisibleOnly)
    {
        public IReadOnlyList<ChartIndexedNumberPoint?> DensePoints()
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = Points ?? [];
            int pointCount = Math.Max(PointCount ?? 0, InferPointCount(points) ?? 0);
            if (pointCount <= 0)
            {
                return [];
            }

            var values = new ChartIndexedNumberPoint?[pointCount];
            foreach (ChartIndexedNumberPoint point in points)
            {
                if (point.Index >= 0 && point.Index < pointCount && point.Value is not null)
                {
                    values[point.Index] = point;
                }
            }

            return values;
        }

        public IReadOnlyList<ChartIndexedNumberPoint> WorkbookPointsForPlotVisibility(bool plotVisibleOnly)
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = WorkbookPoints ?? [];
            return plotVisibleOnly
                ? points.Where(point => IsWorkbookPointVisible(point.WorkbookCell)).ToArray()
                : points;
        }

        public ChartIndexedNumberPoint? WorkbookPointForIndex(int index)
        {
            foreach (ChartIndexedNumberPoint point in WorkbookPointsForPlotVisibility(PlotVisibleOnly))
            {
                if (point.Index == index)
                {
                    return point;
                }
            }

            return null;
        }
    }

    private readonly record struct ChartIndexedNumberPoint(
        int Index,
        ChartPointIndexSource IndexSource,
        double? Value,
        string Text,
        bool HasValue,
        ChartWorkbookRangeCell WorkbookCell);

    private readonly record struct ChartIndexedTextVector(
        IReadOnlyList<ChartIndexedTextPoint> Points,
        int? PointCount,
        IReadOnlyList<IReadOnlyList<ChartIndexedTextPoint>> Levels,
        string? Formula,
        PptxSceneChartDataSource Source,
        IReadOnlyList<ChartIndexedTextPoint> WorkbookPoints,
        bool PlotVisibleOnly)
    {
        public IReadOnlyList<ChartIndexedTextPoint?> DensePoints()
        {
            IReadOnlyList<ChartIndexedTextPoint> points = Points ?? [];
            int pointCount = Math.Max(PointCount ?? 0, InferPointCount(points) ?? 0);
            if (pointCount <= 0)
            {
                return [];
            }

            var values = new ChartIndexedTextPoint?[pointCount];
            foreach (ChartIndexedTextPoint point in points)
            {
                if (point.Index >= 0 && point.Index < pointCount && point.HasText)
                {
                    values[point.Index] = point;
                }
            }

            return values;
        }

        public IReadOnlyList<ChartIndexedTextPoint> WorkbookPointsForPlotVisibility(bool plotVisibleOnly)
        {
            IReadOnlyList<ChartIndexedTextPoint> points = WorkbookPoints ?? [];
            return plotVisibleOnly
                ? points.Where(point => IsWorkbookPointVisible(point.WorkbookCell)).ToArray()
                : points;
        }
    }

    private static bool IsWorkbookPointVisible(ChartWorkbookRangeCell cell)
    {
        return !cell.RowHidden && !cell.ColumnHidden;
    }

    private readonly record struct ChartIndexedTextPoint(
        int Index,
        ChartPointIndexSource IndexSource,
        string Text,
        bool HasText,
        ChartWorkbookRangeCell WorkbookCell);

    private enum ChartPointIndexSource
    {
        OrdinalFallback,
        OoxmlIndex,
        WorkbookRange
    }

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

    private static ChartSeriesStroke ChartDataLabelLeaderLineDefaultStroke { get; } = new(new RgbColor(89, 89, 89), 1d, 0.75d);

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

    private readonly record struct ChartValueAxisRenderOptions(ChartAxisUnits Units, bool Reversed, double? CrossingValue, bool MajorGridlines, bool MinorGridlines, ChartGridlineStyle GridlineStyle);

    private readonly record struct ChartBubbleValueAxisOptions(ChartAxisUnits Units, ChartGridlineStyle GridlineStyle);

    private readonly record struct ChartLayout(ChartFrameBox Frame, ChartLayoutBox PlotAreaBox, ChartPlotBox PlotBox, bool ManualPlotLayoutApplied, string? Title, PptxSceneChartTextBodyProperties TitleTextBodyProperties, ChartLegendLayout Legend);

    private readonly record struct ChartFrameBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartLayoutBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartPlotBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartLegendBox(
        double X,
        double ClipY,
        double Width,
        double ClipHeight,
        double FirstY,
        double LineHeight,
        double MarkerSize,
        double MarkerWidth,
        double TextGap,
        bool Horizontal,
        bool SideStrokeLegend);

    private readonly record struct ChartRightLegendReserve(
        double Width,
        double FontSize,
        double MaxTextWidth,
        int MaxTextLength,
        bool IncludesAreaReserve);

    private readonly record struct ChartAxisSource(PptxSceneChartAxis? SceneAxis, XElement? XmlAxis);

    private readonly record struct ChartPlotLayout(ChartLayoutBox PlotAreaBox, ChartPlotBox PlotBox, PptxSceneChartManualLayoutTarget? ManualLayoutTargetKind)
    {
        public static ChartPlotLayout FromPlotBox(ChartPlotBox plotBox)
        {
            return new ChartPlotLayout(new ChartLayoutBox(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height), plotBox, null);
        }
    }

    private readonly record struct ChartAxisTitleReserveSides(bool Left, bool Right, bool Bottom, bool Top)
    {
        public bool HasHorizontalTitle => Bottom || Top;

        public bool HasVerticalTitle => Left || Right;

        public ChartAxisTitleReserveSides With(PptxSceneChartAxisPosition positionKind)
        {
            return positionKind switch
            {
                PptxSceneChartAxisPosition.Left => this with { Left = true },
                PptxSceneChartAxisPosition.Right => this with { Right = true },
                PptxSceneChartAxisPosition.Bottom => this with { Bottom = true },
                PptxSceneChartAxisPosition.Top => this with { Top = true },
                _ => this
            };
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

    private readonly record struct ChartTextStyle(
        string? FontFamily,
        double FontSize,
        RgbColor Color,
        double Alpha,
        bool Bold,
        bool Italic,
        bool Underline,
        bool Strike,
        string? RequestedTypeface,
        PptxThemeTypefaceSource? TypefaceSource);

    private readonly record struct ChartTextStyleOverride(
        string? FontFamily,
        double? FontSize,
        RgbColor? Color,
        double? Alpha,
        bool? Bold,
        bool? Italic,
        bool? Underline,
        bool? Strike,
        string? RequestedTypeface,
        PptxThemeTypefaceSource? TypefaceSource)
    {
        public static ChartTextStyleOverride Empty { get; } = new(null, null, null, null, null, null, null, null, null, null);
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> EmptyChartDataLabelOverrides { get; } = new Dictionary<int, ChartDataLabelOverride>();

    private static IReadOnlyDictionary<string, ChartBooleanOption> EmptyChartDataLabelFlagOptions { get; } = new Dictionary<string, ChartBooleanOption>();

    private static readonly string[] ChartDataLabelFlagNames =
    [
        "showVal",
        "showPercent",
        "showCatName",
        "showSerName",
        "showLeaderLines",
        "showLegendKey",
        "showBubbleSize"
    ];

    private readonly record struct ChartDataLabelOptions(bool ShowValue, bool ShowPercent, bool ShowCategoryName, bool ShowSeriesName, bool ShowLeaderLines, bool ShowLegendKey, bool ShowBubbleSize, ChartDataLabelLeaderLines LeaderLines, string CustomText, IReadOnlyList<ChartTextRunOverride> CustomTextRuns, PptxSceneChartDataLabelPosition PositionKind, string Position, string Separator, string NumberFormat, ChartNumberFormat NumberFormatInfo, PptxSceneChartManualLayout Layout, ChartTextStyleOverride TextStyle, PptxSceneChartTextBodyProperties TextBodyProperties, ChartShapeStyle ShapeStyle, IReadOnlyDictionary<string, ChartBooleanOption> FlagOptions, IReadOnlyDictionary<int, ChartDataLabelOverride> Overrides, bool IsDefined)
    {
        public static ChartDataLabelOptions None { get; } = new(ShowValue: false, ShowPercent: false, ShowCategoryName: false, ShowSeriesName: false, ShowLeaderLines: false, ShowLegendKey: false, ShowBubbleSize: false, LeaderLines: ChartDataLabelLeaderLines.Empty, CustomText: string.Empty, CustomTextRuns: [], PositionKind: PptxSceneChartDataLabelPosition.Unknown, Position: string.Empty, Separator: string.Empty, NumberFormat: string.Empty, NumberFormatInfo: default, Layout: default, TextStyle: ChartTextStyleOverride.Empty, TextBodyProperties: default, ShapeStyle: ChartShapeStyle.Empty, FlagOptions: EmptyChartDataLabelFlagOptions, Overrides: EmptyChartDataLabelOverrides, IsDefined: false);

        public bool HasVisibleText => ShowValue || ShowPercent || ShowCategoryName || ShowSeriesName || ShowBubbleSize ||
            !string.IsNullOrWhiteSpace(CustomText) ||
            Overrides.Values.Any(label => label.ShowValue == true || label.ShowPercent == true || label.ShowCategoryName == true || label.ShowSeriesName == true || label.ShowBubbleSize == true || !string.IsNullOrWhiteSpace(label.CustomText));

        public bool HasVisibleContent => HasVisibleText || ShowLegendKey ||
            Overrides.Values.Any(label => label.ShowLegendKey == true);
    }

    private readonly record struct ChartDataLabelLeaderLines(bool IsDefined, ChartSeriesStroke? Stroke)
    {
        public static ChartDataLabelLeaderLines Empty { get; } = new(IsDefined: false, Stroke: null);
    }

    private readonly record struct ChartTextRunOverride(string Text, ChartTextStyleOverride TextStyle);

    private readonly record struct ChartTextRunLayout(string Text, ChartTextStyle Style, double Width);

    private readonly record struct ChartDataLabelOverride(bool? ShowValue, bool? ShowPercent, bool? ShowCategoryName, bool? ShowSeriesName, bool? ShowLeaderLines, bool? ShowLegendKey, bool? ShowBubbleSize, ChartDataLabelLeaderLines LeaderLines, string CustomText, IReadOnlyList<ChartTextRunOverride> CustomTextRuns, PptxSceneChartDataLabelPosition PositionKind, string Position, string Separator, string NumberFormat, ChartNumberFormat NumberFormatInfo, PptxSceneChartManualLayout Layout, ChartTextStyleOverride TextStyle, PptxSceneChartTextBodyProperties TextBodyProperties, ChartShapeStyle ShapeStyle, IReadOnlyDictionary<string, ChartBooleanOption> FlagOptions);

    private readonly record struct ChartNumberFormat(bool IsDefined, string FormatCode, bool? SourceLinked, string SourceLinkedValue);

    private readonly record struct ChartSeriesNameRecord(
        string ActiveName,
        string CacheName,
        PptxSceneChartDataSource Source,
        IReadOnlyList<ChartIndexedTextPoint> WorkbookPoints);

    private readonly record struct ChartLegendEntry(string Name, ChartSeriesFill? Fill, ChartSeriesStroke? Stroke, ChartMarkerStyle? Marker, ChartSeriesNameRecord? SeriesName);

    private readonly record struct ChartBooleanOption(bool Value, string RawValue, bool IsDefined);

    private readonly record struct ChartLegendLayout(PptxSceneChartLegendPosition PositionKind, string Position, bool Overlay, bool Visible, PptxSceneChartManualLayout Layout, PptxSceneChartTextBodyProperties TextBodyProperties, ChartShapeStyle ShapeStyle)
    {
        public static ChartLegendLayout Hidden { get; } = new(PptxSceneChartLegendPosition.Right, "r", Overlay: false, Visible: false, default, default, ChartShapeStyle.Empty);
    }

    private readonly record struct ChartShapeStyle(ChartSeriesFill? Fill, GradientFill? GradientFill, ChartSeriesStroke? Stroke, PptxSceneGlow Glow, PptxSceneOuterShadow OuterShadow)
    {
        public static ChartShapeStyle Empty { get; } = new(null, null, null, default, default);

        public bool IsEmpty => Fill is null && GradientFill is null && Stroke is null && !Glow.HasGlow && !OuterShadow.HasShadow;
    }

    private readonly record struct ChartMarkerStyle(PptxSceneChartMarkerSymbol SymbolKind, string Symbol, string? SizeValue, double Size, ChartSeriesFill? Fill, ChartSeriesStroke? Stroke, bool IsDefined)
    {
        public static ChartMarkerStyle Default { get; } = new(PptxSceneChartMarkerSymbol.Circle, "circle", null, PptxChartMarkerMetricRules.DefaultChartMarkerSize, null, null, false);
    }
}
