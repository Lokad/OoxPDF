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
    private static void RenderLineChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartLayoutBox plotAreaBox, ChartPlotBox plotBox, IReadOnlyList<ChartIndexedNumberVector> series, ChartLinePlotOptions lineOptions, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle> markerStyles, ChartValueAxisRenderOptions valueAxisOptions, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents)
    {
        bool stacked = lineOptions.Stacked;
        bool percentStacked = lineOptions.PercentStacked;
        IReadOnlyList<ChartBooleanOption> smoothSeries = lineOptions.SmoothSeries;
        bool majorGridlines = valueAxisOptions.MajorGridlines;
        bool minorGridlines = valueAxisOptions.MinorGridlines;
        ChartGridlineStyle gridlineStyle = valueAxisOptions.GridlineStyle;
        ChartAxisUnits axisUnits = valueAxisOptions.Units;
        double? valueAxisCrossingValue = valueAxisOptions.CrossingValue;
        bool valueAxisReversed = valueAxisOptions.Reversed;
        PptxSceneChartDisplayBlanksAs displayBlanksAs = lineOptions.DisplayBlanksAs;
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

    private static ChartLayout GetLineChartLayout(PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, PptxColorMap colorMap, ChartWorkbookData? workbook, bool plotVisibleOnly, PresentationFontResolver? fontResolver)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        PptxSceneChartTextBodyProperties titleTextBodyProperties = ReadSceneOrXmlChartTitleTextBodyProperties(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
        ChartTextStyle legendTextStyle = ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml);
        ChartPlotLayout plotLayout = GetLineChartPlotLayout();
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, titleTextBodyProperties, legend);

        ChartPlotLayout GetLineChartPlotLayout()
        {
            bool hasTitle = !string.IsNullOrWhiteSpace(title);
            bool hasRightLegend = legend.Visible && !legend.Overlay && legend.PositionKind == PptxSceneChartLegendPosition.Right;
            bool hasLineChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Line) is not null;
            ChartPlotBox defaultPlotBox = !hasTitle && hasRightLegend
                ? GetCartesianNoTitleRightLegendPlotBox(frame, theme, chartXml, sceneChart, workbook, plotVisibleOnly, fontResolver, legendTextStyle)
                : hasTitle && hasRightLegend && hasLineChart
                    ? GetLineTitleRightLegendPlotBox()
                    : GetDefaultChartPlotBox(frame);
            return TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout manualPlotLayout)
                ? manualPlotLayout
                : ChartPlotLayout.FromPlotBox(defaultPlotBox);

            ChartPlotBox GetLineTitleRightLegendPlotBox()
            {
                return GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.LineTitleRightLegend);
            }
        }
    }

    // The reserve covers the widest label plus a fixed padding; it intentionally does not
    // grow with label character count because the widest label already spans the longest text.
    private static double ComputeNoTitleRightLegendLeftInset(double maxValueLabelWidth, double frameWidth)
    {
        return maxValueLabelWidth +
            Math.Min(
                PptxChartMetricRules.LineRightLegendValueAxisPadding,
                frameWidth * PptxChartMetricRules.LineRightLegendValueAxisFrameWidthPaddingRatio);
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

        PptxSceneChartPlot? plot = ReadSceneChartPlot(sceneChart, plotKind, 0);
        IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSceneOrXmlChartSeriesNameRecords(plot, plotElement, workbook);
        ChartRightLegendReserve rightLegendReserve = ResolveRightLegendReserve(
            frame,
            seriesNames,
            legendTextStyle,
            includeAreaReserve: plotKind == PptxSceneChartPlotKind.Area,
            fontResolver: fontResolver);
        var textMeasurer = new ChartTextMeasurer(fontResolver);

        double maxValueLabelWidth = 0d;
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
                IReadOnlyList<double> tickValues = GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount);
                ChartTextStyle valueAxisTextStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                string[] tickLabels = tickValues
                    .Select(value => FormatSceneOrXmlChartAxisLabel(value, valueAxis.SceneAxis, valueAxis.XmlAxis, defaultNumberFormat: null))
                    .ToArray();
                maxValueLabelWidth = tickLabels.Length == 0
                    ? 0d
                    : tickLabels.Max(label => textMeasurer.Measure(label, valueAxisTextStyle));
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
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, valueAxisForScale, GetLineChartValueExtents(seriesVectors, stacked, percentStacked), percentStacked, useNearMaximumHeadroom: !percentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxisForScale), percentStacked);
                IReadOnlyList<double> tickValues = GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount);
                ChartTextStyle valueAxisTextStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxisForScale, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                string[] tickLabels = tickValues
                    .Select(value => FormatSceneOrXmlChartAxisLabel(value, valueAxis.SceneAxis, valueAxisForScale, percentStacked ? "0%" : null))
                    .ToArray();
                maxValueLabelWidth = tickLabels.Length == 0
                    ? 0d
                    : tickLabels.Max(label => textMeasurer.Measure(label, valueAxisTextStyle));
            }
        }

        double leftInset = maxValueLabelWidth > 0d
            ? ComputeNoTitleRightLegendLeftInset(maxValueLabelWidth, frame.Width)
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

    private static ChartLayout GetBubbleChartLayout(PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartPlot? bubblePlot, XElement bubbleChart, PptxColorMap colorMap, ChartWorkbookData? workbook, PresentationFontResolver? fontResolver)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        PptxSceneChartTextBodyProperties titleTextBodyProperties = ReadSceneOrXmlChartTitleTextBodyProperties(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
        ChartTextStyle legendTextStyle = ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml);
        ChartPlotLayout plotLayout = GetBubbleChartPlotLayout();
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, titleTextBodyProperties, legend);

        ChartPlotLayout GetBubbleChartPlotLayout()
        {
            bool hasTitle = !string.IsNullOrWhiteSpace(title);
            bool hasRightLegend = legend.Visible && !legend.Overlay && legend.PositionKind == PptxSceneChartLegendPosition.Right;
            ChartPlotBox defaultPlotBox = hasTitle && hasRightLegend
                ? GetBubbleTitleRightLegendPlotBox()
                : GetDefaultChartPlotBox(frame);
            return TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout manualPlotLayout)
                ? manualPlotLayout
                : ChartPlotLayout.FromPlotBox(defaultPlotBox);

            ChartPlotBox GetBubbleTitleRightLegendPlotBox()
            {
                double x = frame.X + frame.Width * PptxChartMetricRules.LineTitleRightLegendPlotBoxXRatio;
                double y = frame.Y + frame.Height * PptxChartMetricRules.LineTitleRightLegendPlotBoxYRatio;
                double width = frame.Width * PptxChartMetricRules.BubbleTitleRightLegendPlotBoxWidthRatio;
                double height = frame.Height * PptxChartMetricRules.LineTitleRightLegendPlotBoxHeightRatio;
                return new ChartPlotBox(x, y, width, height);
            }
        }
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
            .Select(point => point is { } unwrapped ? unwrapped.Value ?? 0d : 0d);
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
        ChartSeriesFill fill = marker.Fill ?? new ChartSeriesFill(defaultFill, 1d, null, null);
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
        ChartSeriesFill fill = marker.Fill ?? new ChartSeriesFill(defaultFill, 1d, null, null);
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

    private static void StrokeChartPointRectangle(PdfGraphicsBuilder graphics, int seriesIndex, int categoryIndex, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, double x, double y, double width, double height, ChartSeriesStroke? fallbackStroke)
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
}
