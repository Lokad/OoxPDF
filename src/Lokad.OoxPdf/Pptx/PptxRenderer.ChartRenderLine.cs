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
    private static void RenderLineChart(PdfGraphicsBuilder graphics, ChartLayoutBox plotAreaBox, ChartPlotBox plotBox, IReadOnlyList<ChartIndexedNumberVector> series, ChartLinePlotOptions lineOptions, IReadOnlyList<ChartLineAppearance> seriesAppearance, ChartValueAxisRenderOptions valueAxisOptions, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents, bool categoryPointsOnTickMarks, double categoryTickFontSize, int? chartStyleId, bool categoryTicksTopSide = false, Func<int, int, ChartLineAppearance>? pointAppearance = null, bool categoryAxisReversed = false, bool varyPointColors = false)
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
        IReadOnlyList<IReadOnlyList<double?>> denseSeries = DensifyChartValueSeries(series);
        // Preserve the gallery's plot-level smoothing fallback independently of paint.
        bool gallerySmooth = IsGalleryLineChartStyle(chartStyleId);
        bool plotSmoothed = smoothSeries.Any(static option => option.Value);
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
                    () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits, valueAxisCrossingValue, valueAxisReversed, major: false, gridlineStyle.Minor, valueAxisOptions.HasMajorGridlineElement, valueAxisOptions.TickPolicy));
            }

            if (majorGridlines)
            {
                RenderInChartPlotAreaClip(
                    graphics,
                    plotBox,
                    () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, axisUnits, valueAxisCrossingValue, valueAxisReversed, major: true, gridlineStyle.Major, valueAxisOptions.HasMajorGridlineElement, valueAxisOptions.TickPolicy));
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
                IReadOnlyList<double?> values = denseSeries[seriesIndex];
                if (values.Count == 0)
                {
                    continue;
                }

                bool explicitStraight = seriesIndex < smoothSeries.Count && smoothSeries[seriesIndex].IsDefined && !smoothSeries[seriesIndex].Value;
                bool smoothCurve = !explicitStraight && (plotSmoothed || gallerySmooth);
                ChartLineAppearance seriesStyle = seriesAppearance[seriesIndex];
                ChartSeriesStroke stroke = seriesStyle.Stroke;
                if (stroke.Alpha > .001d && stroke.Alpha < 1d)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(1d, stroke.Alpha);
                }
                SetChartStroke(graphics, stroke);
                var points = new List<(double X, double Y)>(values.Count);
                var markers = new List<(double X, double Y)>(values.Count);
                List<int>? sourceIndexes = !varyPointColors || pointAppearance is null ? null : new(values.Count);
                Func<int, ChartSeriesStroke>? strokeAtPoint = sourceIndexes is null ? null :
                    index => pointAppearance!(seriesIndex, sourceIndexes[index]).Stroke;
                // RV04: span-mode smoothing halvings (a chart-edge span start eases at a sixth
                // whether the first segment spans one hole or several; gap/zero modes leave
                // every flag false so their paths stay byte-identical).
                var halvePoints = new List<bool>(values.Count);
                bool fragmentStartsAtGap = false;
                for (int i = 0; i < values.Count; i++)
                {
                    if (values[i] is not { } value)
                    {
                        if (displayBlanksAs == PptxSceneChartDisplayBlanksAs.Zero)
                        {
                            value = 0d;
                        }
                        else
                        {
                            if (displayBlanksAs != PptxSceneChartDisplayBlanksAs.Span)
                            {
                                if (stroke.Alpha > .001d || strokeAtPoint is not null)
                                    StrokeStyleLineSegmentInPlotClip(graphics, plotBox, points, halvePoints, smoothCurve, fragmentStartsAtGap, true, strokeAtPoint);
                                points.Clear();
                                sourceIndexes?.Clear();
                                halvePoints.Clear();
                                fragmentStartsAtGap = true;
                            }

                            continue;
                        }
                    }

                    double pointX = ChartCategoryToPlotCoordinate(i, pointCount, plotX, plotWidth, categoryPointsOnTickMarks, categoryAxisReversed);
                    double positiveTotal = GetCategoryPositiveTotal(denseSeries, i, percentStacked);
                    double normalizedValue = NormalizeStackedValue(value, positiveTotal, percentStacked);
                    double plottedValue = stacked ? lower[i] + normalizedValue : value;
                    double pointY = ChartValueToPlotCoordinate(valueExtents, plottedValue, plotY, plotHeight, valueAxisReversed);
                    points.Add((pointX, pointY));
                    sourceIndexes?.Add(i);
                    markers.Add((pointX, pointY));
                    halvePoints.Add(displayBlanksAs == PptxSceneChartDisplayBlanksAs.Span && ((i > 0 && values[i - 1] is null) || (i + 1 < values.Count && values[i + 1] is null)));
                    if (stacked)
                    {
                        lower[i] = plottedValue;
                    }
                }

                if (stroke.Alpha > .001d || strokeAtPoint is not null)
                    StrokeStyleLineSegmentInPlotClip(graphics, plotBox, points, halvePoints, smoothCurve, fragmentStartsAtGap, false, strokeAtPoint);
                if (stroke.Alpha > .001d && stroke.Alpha < 1d) graphics.RestoreState();

                int markerSourceIndex = 0;
                foreach ((double pointX, double pointY) in markers)
                {
                    ChartLineAppearance? point = null;
                    if (pointAppearance is not null)
                    {
                        while (markerSourceIndex < values.Count && values[markerSourceIndex] is null && displayBlanksAs != PptxSceneChartDisplayBlanksAs.Zero)
                            markerSourceIndex++;
                        point = pointAppearance(seriesIndex, markerSourceIndex++);
                    }
                    ChartLineAppearance appearance = point ?? seriesStyle;
                    ChartMarkerStyle marker = appearance.Marker;
                    RgbColor markerColor = appearance.MarkerColor;
                    if (chartStyleId == 26 || chartStyleId == 126)
                    {
                        // RV04: style-26 draws no line markers at all (Office renders none
                        // with plot marker=1, none, or series-explicit symbols); explicit marker
                        // fills still win in legend keys, which take the flat path there.
                        continue;
                    }

                    RgbColor markerOutline = appearance.MarkerOutlineColor;
                    double? outlineWidth = appearance.MarkerOutlineWidth;
                    if (chartStyleId is 18 or 118
                        && marker.Fill is null
                        && TryGetStyleBarGradientStops(chartStyleId, new ChartSeriesFill(markerColor, 1d, null, null), out IReadOnlyList<PdfShadingStop> gradientStops)
                        && PaintStyleMarkerGradient(graphics, plotBox, marker.SymbolKind, pointX, pointY, marker.Size, gradientStops))
                    {
                        // RV04: gradient-filled markers skip the flat fill and keep only the
                        // rim, matching the Office fill-plus-outline marker passes.
                        ChartSeriesStroke? rim = ChartMarkerOutlineStroke(marker, markerOutline, outlineWidth);
                        if (rim is not null)
                        {
                            RenderInChartPlotAreaClip(graphics, plotBox, () => DrawChartMarkerStroke(graphics, pointX, pointY, marker.SymbolKind, marker.Size, rim));
                        }
                    }
                    else
                    {
                        // RV04: no loop-top fill color state (DrawChartMarkerInPlotClip sets
                        // the fill itself; Office emits one fill state per marker).
                        DrawChartMarkerInPlotClip(graphics, plotBox, pointX, pointY, marker, markerColor, markerOutline, outlineWidth);
                    }
                }

            }
        }
        if (axesStyle.CategoryAxisVisible && categoryAxisStroke.Alpha > 0.001d)
        {
            SetChartStroke(graphics, categoryAxisStroke);
            DrawCartesianChartCategoryAxisMajorTicks(graphics, plotX, plotWidth, pointCount, valueAxisCrossingY, axesStyle.CategoryAxisMajorTickMark, categoryTickFontSize, categoryPointsOnTickMarks, categoryTicksTopSide, categoryAxisReversed);
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

    // RV04: gallery polyline widths: style-18 strokes at 5pt, style-26 at 7pt,
    // other styles keep the 2.25pt default (explicit series widths keep winning
    // by construction at the call site).
    // RV04: gallery polyline widths: style-18 strokes at 5pt, style-26 at 7pt;
    // smoothed non-gallery series stroke at 3pt; other styles keep the 2.25pt
    // default (explicit series widths keep winning by construction at the call site).
    // RV04: explicit series lines without a width inherit the gallery width (Office keeps
    // 7pt under style-26 and 3pt under the smooth package when only a color is explicit;
    // an explicit width always wins by construction at the call site; the style-18
    // gallery width is the predicted equivalent, other styles keep 3pt).
    private static double ResolveLineKindInheritedStrokeWidth(int? chartStyleId)
    {
        if (chartStyleId == 26 || chartStyleId == 126)
        {
            return PptxChartMetricRules.StyleHeavyLineSeriesStrokeWidth;
        }

        if (chartStyleId == 18 || chartStyleId == 118)
        {
            return PptxChartMetricRules.StyleLineSeriesStrokeWidth;
        }

        return ChartSeriesInheritedStrokeWidth;
    }

    private static double ResolveStyleLineSeriesWidth(int? chartStyleId, bool smoothed)
    {
        if (chartStyleId == 26 || chartStyleId == 126)
        {
            return PptxChartMetricRules.StyleHeavyLineSeriesStrokeWidth;
        }

        if (chartStyleId == 18 || chartStyleId == 118)
        {
            return PptxChartMetricRules.StyleLineSeriesStrokeWidth;
        }

        return smoothed
            ? PptxChartMetricRules.SmoothLineSeriesStrokeWidth
            : ChartLineDefaultStrokeWidth;
    }

    // RV04: smoothed series: piecewise cubic Hermite through the data points with uniform-x
    // third offsets, central-difference interior tangents and one-sided (chord) end tangents
    // (COM-built style-18, style-26 and default-smooth decks: 30/30 segments predict the
    // Office bezier controls within rounding; gap-adjacent fragment ends halve their control
    // offset (COM-built gap decks); span joints use symmetric mean-thirds offsets and span
    // starts ease at a sixth (COM-built span decks, including consecutive-hole spans); the legacy
    // Catmull-Rom stays for scatter smooth paths only).
    private static void StrokeStyleLineSmoothPath(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points, IReadOnlyList<bool> halvePoints, bool startIsGap, bool endIsGap, Func<int, ChartSeriesStroke>? strokeAtPoint = null)
    {
        if (points.Count < 2)
        {
            return;
        }

        if (strokeAtPoint is null) graphics.MoveTo(points[0].X, points[0].Y);
        for (int i = 0; i < points.Count - 1; i++)
        {
            if (strokeAtPoint is not null)
            {
                ChartSeriesStroke stroke = strokeAtPoint(i + 1);
                graphics.SaveState();
                graphics.SetAlpha(1d, stroke.Alpha);
                SetChartStroke(graphics, stroke);
                graphics.MoveTo(points[i].X, points[i].Y);
            }
            double dx = points[i + 1].X - points[i].X;
            double startRun = i == 0 ? dx : points[i + 1].X - points[i - 1].X;
            double endRun = i + 2 >= points.Count ? dx : points[i + 2].X - points[i].X;
            if (dx == 0d || startRun == 0d || endRun == 0d)
            {
                graphics.LineTo(points[i + 1].X, points[i + 1].Y);
                if (strokeAtPoint is not null) { graphics.StrokeCurrentPath(); graphics.RestoreState(); }
                continue;
            }

            double startSlope = i == 0
                ? (points[1].Y - points[0].Y) / startRun
                : (points[i + 1].Y - points[i - 1].Y) / startRun;
            double endSlope = i + 2 >= points.Count
                ? (points[i + 1].Y - points[i].Y) / dx
                : (points[i + 2].Y - points[i].Y) / endRun;
            double prevDx = i == 0 ? dx : points[i].X - points[i - 1].X;
            double nextDx = i + 2 >= points.Count ? dx : points[i + 2].X - points[i + 1].X;
            double startOffset = i == 0
                ? startIsGap || halvePoints[i] ? dx / 6d : dx / 3d
                : (dx + prevDx) / 6d;
            double endOffset = i + 2 >= points.Count
                ? endIsGap ? dx / 6d : dx / 3d
                : (dx + nextDx) / 6d;
            graphics.CurveTo(
                points[i].X + startOffset,
                points[i].Y + startOffset * startSlope,
                points[i + 1].X - endOffset,
                points[i + 1].Y - endOffset * endSlope,
                points[i + 1].X,
                points[i + 1].Y);
            if (strokeAtPoint is not null) { graphics.StrokeCurrentPath(); graphics.RestoreState(); }
        }

        if (strokeAtPoint is null) graphics.StrokeCurrentPath();
    }

    // RV04: smoothed series routing: explicit, plot-wide and style-implied smoothing share
    // the end-chord Hermite path (Office draws identical beziers in all three cases);
    // an explicit smooth=false keeps the straight path.
    private static void StrokeStyleLineSegmentInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, IReadOnlyList<(double X, double Y)> points, IReadOnlyList<bool> halvePoints, bool smooth, bool startIsGap, bool endIsGap, Func<int, ChartSeriesStroke>? strokeAtPoint = null)
    {
        if (smooth)
        {
            RenderInChartPlotAreaClip(graphics, plotBox, () => StrokeStyleLineSmoothPath(graphics, points, halvePoints, startIsGap, endIsGap, strokeAtPoint));
        }
        else if (strokeAtPoint is not null)
        {
            RenderInChartPlotAreaClip(graphics, plotBox, () =>
            {
                for (int i = 1; i < points.Count; i++)
                {
                    ChartSeriesStroke stroke = strokeAtPoint(i);
                    graphics.SaveState();
                    graphics.SetAlpha(1d, stroke.Alpha);
                    SetChartStroke(graphics, stroke);
                    graphics.StrokeLine(points[i - 1].X, points[i - 1].Y, points[i].X, points[i].Y);
                    graphics.RestoreState();
                }
            });
        }
        else
        {
            StrokeLineChartPointSegmentInPlotClip(graphics, plotBox, points, false);
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

    // Titled right-legend plot right edge: the preset ratio stands, but never past the
    // measured legend reserve (composite narrow frame: preset 354.0 vs measured 347.98,
    // while ladder frames keep the preset; same min-cap shape as the column reserve).
    private static double ResolveTitledRightLegendPlotRight(double presetRight, double frameRight, double reserveWidth)
    {
        return Math.Min(presetRight, frameRight - reserveWidth);
    }

    private static async ValueTask<ChartLayout> GetCartesianChartLayoutAsync(PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, PptxSceneChart sceneChart, PptxColorMap colorMap, ChartWorkbookData? workbook, bool plotVisibleOnly, PresentationFontResolver? fontResolver, ChartValueExtents? resolvedValueExtents)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = sceneChart.Title.Text;
        PptxSceneChartTextBodyProperties titleTextBodyProperties = sceneChart.Title.TextBodyProperties;
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
        ChartTextStyle legendTextStyle = ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml);
        // Line and area callers supply their already resolved source scale.
        // Scatter and composed/unsupported layouts keep the existing fallback.
        PptxSceneChartPlotKind categoryPlotKind = sceneChart.Plots.Count == 1 && sceneChart.Plots[0].PlotKind == PptxSceneChartPlotKind.Area
            ? PptxSceneChartPlotKind.Area : PptxSceneChartPlotKind.Line;
        XElement? categoryPlotElement = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, categoryPlotKind);
        PptxSceneChartPlot? categoryPlot = ReadSceneChartPlot(sceneChart, categoryPlotKind, 0);
        bool canMeasureCategoryPlot = categoryPlotElement is not null && resolvedValueExtents is not null;
        double? categoryLabelOffsetAdjustment = null;
        ChartAxisPlacement? categoryAxisPlacement = await ReadAutomaticCategoryPlacementAsync().ConfigureAwait(false);
        ChartPlotLayout plotLayout = (await GetPlotLayoutAsync().ConfigureAwait(false));
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, titleTextBodyProperties, legend, categoryLabelOffsetAdjustment, categoryAxisPlacement);

        async ValueTask<ChartAxisPlacement?> ReadAutomaticCategoryPlacementAsync()
        {
            // Caption side/crossing belongs to the scale, independently of
            // whether the title/header can use measured plot placement.
            // Keep authored/composed frames, unsupported scales and unresolved
            // font boxes on their existing fallback.
            if (resolvedValueExtents is not { } extents || sceneChart.PlotAreaLayout.HasGeometryValues || sceneChart.Plots.Count != 1)
                return null;
            PptxSceneChartPlot? plot = categoryPlot;
            XElement? plotElement = categoryPlotElement;
            if (plotElement is null) return null;
            IReadOnlyList<ChartAxisSource> axes = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, plotElement);
            if (axes.Count != 1 || axes[0].SceneAxis is not { HasLogarithmicScale: false } valueAxis ||
                valueAxis.PositionKind is not (PptxSceneChartAxisPosition.Left or PptxSceneChartAxisPosition.Right))
                return null;
            ChartAxisSource category = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, plot, chartXml, plotElement);
            if (category.SceneAxis is not { PositionKind: PptxSceneChartAxisPosition.Bottom } ||
                !IsSceneOrXmlChartAxisLabelVisible(category.SceneAxis, category.XmlAxis))
                return null;
            var measure = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
            ChartTextStyle categoryStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, category.SceneAxis, chartXml, category.XmlAxis,
                PptxChartMetricRules.CategoryAxisFallbackFontSize, "categoryAxis");
            if (await measure.TryMeasureNormalLineHeightAsync(categoryStyle).ConfigureAwait(false) is null)
                return null;
            if (IsSceneOrXmlChartAxisLabelVisible(valueAxis, axes[0].XmlAxis))
            {
                ChartTextStyle numericStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis, chartXml, axes[0].XmlAxis,
                    PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
                if (await measure.TryMeasureNormalLineHeightAsync(numericStyle).ConfigureAwait(false) is null)
                    return null;
            }
            bool reversed = ReadSceneOrXmlValueAxisReversed(valueAxis, axes[0].XmlAxis);
            return ReadChartAxisPlacement(category, extents, reversed);
        }

        async ValueTask<ChartPlotLayout> GetPlotLayoutAsync()
        {
            bool hasTitle = !string.IsNullOrWhiteSpace(title);
            bool hasRightLegend = legend.Visible && !legend.Overlay && legend.PositionKind == PptxSceneChartLegendPosition.Right;
            bool hasLineChart = categoryPlotKind == PptxSceneChartPlotKind.Line && categoryPlotElement is not null;
            ChartPlotBox defaultPlotBox = !hasTitle && hasRightLegend
                ? (await GetCartesianNoTitleRightLegendPlotBoxAsync(frame, theme, chartXml, sceneChart, workbook, plotVisibleOnly, fontResolver, legendTextStyle).ConfigureAwait(false))
                : hasTitle && hasRightLegend && hasLineChart
                    ? (await GetLineTitleRightLegendPlotBoxAsync().ConfigureAwait(false))
                    : GetDefaultChartPlotBox(frame);
            if (!hasTitle && !legend.Visible && canMeasureCategoryPlot &&
                (await TryGetMeasuredCategoryPlotBoxAsync().ConfigureAwait(false)) is { } measuredPlotBox)
            {
                defaultPlotBox = measuredPlotBox;
            }
            else if (canMeasureCategoryPlot &&
                categoryPlotElement is { } lineElement &&
                await TryMeasureCartesianHeaderAsync(theme, sceneChart, legend,
                    categoryPlot, lineElement, chartXml, fontResolver,
                    categoryPlacement: categoryAxisPlacement).ConfigureAwait(false) is { } header)
            {
                if (await TryGetMeasuredCategoryPlotBoxAsync((header.Reserve, header.NumericClearance)).ConfigureAwait(false) is { } composedPlotBox)
                {
                    defaultPlotBox = composedPlotBox;
                }
                else
                {
                    double top = frame.Y + frame.Height - header.Reserve;
                    defaultPlotBox = defaultPlotBox with { Height = Math.Max(1d, top - defaultPlotBox.Y) };
                    // Numeric labels own their measured side strip independently
                    // of captions, bottom legends and the unresolved footer metric.
                    double? numericReserve = PptxChartMetricRules.ColumnPlotRightReserve;
                    ChartAxisSource valueAxis = header.ValueAxis;
                    if (IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                    {
                        var (labelWidth, tickStyle) = await MeasureCategoryPlotValueLabelWidthAsync().ConfigureAwait(false);
                        numericReserve = labelWidth > 0d ? ComputeCartesianAxisLabelSideReserve(labelWidth, tickStyle.FontSize) : null;
                    }
                    if (numericReserve is { } reserve)
                    {
                        bool labelsRight = ResolveSceneOrXmlValueAxisLabelsRightSide(valueAxis.SceneAxis, valueAxis.XmlAxis,
                            ResolveSceneOrXmlValueAxisRightSide(sceneChart, valueAxis.SceneAxis, valueAxis.XmlAxis, defaultRightSide: false));
                        double left = labelsRight ? PptxChartMetricRules.ColumnPlotRightReserve : reserve;
                        double right = labelsRight ? reserve : PptxChartMetricRules.ColumnPlotRightReserve;
                        double width = frame.Width - left - right;
                        if (width >= 1d) defaultPlotBox = defaultPlotBox with { X = frame.X + left, Width = width };
                    }
                }
            }
            return TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout manualPlotLayout)
                ? CenterCartesianInnerFactorPlot(sceneChart.PlotAreaLayout, frame, manualPlotLayout)
                : ChartPlotLayout.FromPlotBox(defaultPlotBox);

            async ValueTask<ChartPlotBox?> TryGetMeasuredCategoryPlotBoxAsync((double Reserve, double NumericClearance)? headerClearance = null)
            {
                // A manual layout owns its source geometry; composed plots and
                // extra value-axis strips retain their existing layout fallback.
                if (sceneChart is null || sceneChart.PlotAreaLayout.HasGeometryValues || sceneChart.Plots.Count != 1)
                {
                    return null;
                }

                PptxSceneChartPlot? plot = categoryPlot;
                XElement? plotElement = categoryPlotElement;
                if (plotElement is null)
                {
                    return null;
                }

                IReadOnlyList<ChartAxisSource> valueAxes = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, plotElement);
                if (valueAxes.Count != 1 || valueAxes[0].SceneAxis is not { } valueAxis ||
                    valueAxis.PositionKind is not (PptxSceneChartAxisPosition.Left or PptxSceneChartAxisPosition.Right))
                {
                    return null;
                }

                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, plot, chartXml, plotElement);
                if (categoryAxis.SceneAxis is not { PositionKind: PptxSceneChartAxisPosition.Bottom } ||
                    !IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    return null;
                }

                bool numericLabelsVisible = IsSceneOrXmlChartAxisLabelVisible(valueAxis, valueAxes[0].XmlAxis);
                var (maxLabelWidth, tickStyle) = await MeasureCategoryPlotValueLabelWidthAsync().ConfigureAwait(false);
                if (numericLabelsVisible && maxLabelWidth <= 0d)
                {
                    return null;
                }

                double sideReserve = numericLabelsVisible ? ComputeCartesianAxisLabelSideReserve(maxLabelWidth, tickStyle.FontSize) :
                    PptxChartMetricRules.ColumnPlotRightReserve;
                bool labelsRight = ResolveSceneOrXmlValueAxisLabelsRightSide(valueAxis, valueAxes[0].XmlAxis,
                    ResolveSceneOrXmlValueAxisRightSide(sceneChart, valueAxis, valueAxes[0].XmlAxis, defaultRightSide: false));
                double leftReserve = labelsRight ? PptxChartMetricRules.ColumnPlotRightReserve : sideReserve;
                double rightReserve = labelsRight ? sideReserve : PptxChartMetricRules.ColumnPlotRightReserve;
                double width = frame.Width - leftReserve - rightReserve;
                if (width < 1d)
                {
                    return null;
                }

                ChartTextStyle categoryStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, chartXml, categoryAxis.XmlAxis,
                    fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
                double numericLineHeight = numericLabelsVisible ? await textMeasurer.MeasureNormalLineHeightAsync(tickStyle).ConfigureAwait(false) : 0d;
                var clearance = headerClearance ?? await MeasureCartesianHeaderReserveAsync(theme, sceneChart, fontResolver, numericLineHeight).ConfigureAwait(false);
                double topReserve = clearance.Reserve;
                double bottomReserve = ComputeCartesianLabelStripBottomReserve(categoryStyle.FontSize);
                double labelOffsetScale = ResolveSceneOrXmlCategoryAxisLabelOffsetScale(categoryAxis.SceneAxis, categoryAxis.XmlAxis);
                double? measuredOffsetAdjustment = await TryMeasureCategoryLabelOffsetAdjustmentAsync(textMeasurer, categoryStyle, labelOffsetScale).ConfigureAwait(false);
                if (measuredOffsetAdjustment is { } offsetAdjustment)
                {
                    bottomReserve += offsetAdjustment;
                }

                double height = frame.Height - topReserve - bottomReserve;
                if (height < 1d)
                {
                    return null;
                }

                // Measurement and painting consume the same scale placement.
                if (categoryAxisPlacement is { } placement)
                {
                    (topReserve, bottomReserve) = ResolveCartesianCategoryReserves(frame.Height,
                        clearance.Reserve, bottomReserve, clearance.NumericClearance, placement);
                    height = frame.Height - topReserve - bottomReserve;
                    if (height < 1d)
                    {
                        return null;
                    }
                }

                categoryLabelOffsetAdjustment = measuredOffsetAdjustment;
                return new ChartPlotBox(frame.X + leftReserve, frame.Y + bottomReserve, width, height);
            }

            async ValueTask<ChartPlotBox> GetLineTitleRightLegendPlotBoxAsync()
            {
                ChartPlotBox presetBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.LineTitleRightLegend);
                double maxValueLabelWidth = (await MeasureCategoryPlotValueLabelWidthAsync().ConfigureAwait(false)).Width;
                if (maxValueLabelWidth <= 0d)
                {
                    return presetBox;
                }

                double x = frame.X + ResolveMeasuredLeftInset(presetBox.X - frame.X, maxValueLabelWidth);
                double presetRight = presetBox.X + presetBox.Width;
                double right = presetRight;
                XElement? titleLegendPlotElement = categoryPlotElement;
                if (titleLegendPlotElement is not null)
                {
                    PptxSceneChartPlot? titleLegendPlot = categoryPlot;
                    IReadOnlyList<ChartSeriesNameRecord> titleLegendSeriesNames = ReadSharedChartSeriesNames(titleLegendPlot, titleLegendPlotElement, workbook);
                    double titleLegendWidth = (await ResolveRightLegendWidthAsync(titleLegendSeriesNames, legendTextStyle, includeAreaReserve: false, fontResolver, lastCategoryLabelWidth: 0d).ConfigureAwait(false));
                    right = ResolveTitledRightLegendPlotRight(presetRight, frame.X + frame.Width, titleLegendWidth);
                }
                double width = Math.Max(1d, right - x);
                return new ChartPlotBox(x, presetBox.Y, width, presetBox.Height);
            }

            async ValueTask<(double Width, ChartTextStyle Style)> MeasureCategoryPlotValueLabelWidthAsync()
            {
                ChartTextStyle tickStyle = default;
                XElement? plotElement = categoryPlotElement;
                if (plotElement is null || resolvedValueExtents is not { } valueExtents) return (0d, tickStyle);
                PptxSceneChartPlot? plot = categoryPlot;
                var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);

                PptxSceneChartGrouping grouping = ReadSceneOrXmlCartesianRightLegendGrouping(sceneChart, plot, chartXml, plotElement, categoryPlotKind);
                bool percentStacked = IsPercentStackedChartGrouping(grouping);

                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, plotElement).FirstOrDefault();
                XElement? valueAxisForScale = ResolveXmlValueAxisForSource(sceneChart, valueAxis, chartXml);
                ChartTextStyle valueAxisTextStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxisForScale, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                tickStyle = valueAxisTextStyle;
                // Hidden labels still supply scale/crossing data, without
                // requesting their font or constructing a label measurement.
                if (!IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxisForScale)) return (0d, tickStyle);
                ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxisForScale), percentStacked);
                IReadOnlyList<double> tickValues = GetChartAxisTickValuesCore(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceVerticalValueTickTargetCount, ReadCartesianChartAxisTickPolicy(sceneChart, valueAxis.SceneAxis, valueExtents, percentStacked));
                string[] tickLabels = tickValues
                    .Select(value => FormatSceneOrXmlChartAxisLabel(value, valueAxis.SceneAxis, valueAxisForScale, percentStacked ? "0%" : null))
                    // Bounds reserve numeric advance even when the source major
                    // lattice does not paint a tick there (Office99/103 pairs).
                    .Concat(new[] { valueExtents.Min, valueExtents.Max }.Select(value =>
                        FormatSceneOrXmlChartAxisLabel(value, valueAxis.SceneAxis, valueAxisForScale, percentStacked ? "0%" : null)))
                    .ToArray();
                await textMeasurer.PrepareAsync(tickLabels, valueAxisTextStyle).ConfigureAwait(false);
                return (tickLabels.Length == 0
                    ? 0d
                    : tickLabels.Max(label => textMeasurer.Measure(label, valueAxisTextStyle)), tickStyle);
            }
        }
    }

    // The reserve covers the widest label plus a fixed padding; it intentionally does not
    // grow with label character count because the widest label already spans the longest text.
    // Left edge for preset-driven cartesian plot boxes: the preset stands unless the measured
    // tick labels need more room, in which case the widest label plus the shared Office gap wins.
    private static double ResolveMeasuredLeftInset(double presetLeftInset, double maxValueLabelWidth)
    {
        return Math.Max(presetLeftInset, maxValueLabelWidth + PptxChartMetricRules.LineRightLegendValueAxisPadding);
    }

    private static double ComputeNoTitleRightLegendLeftInset(double maxValueLabelWidth, double frameWidth)
    {
        return maxValueLabelWidth +
            Math.Min(
                PptxChartMetricRules.LineRightLegendValueAxisPadding,
                frameWidth * PptxChartMetricRules.LineRightLegendValueAxisFrameWidthPaddingRatio);
    }

    private static async ValueTask<ChartPlotBox> GetCartesianNoTitleRightLegendPlotBoxAsync(ChartFrameBox frame, PptxTheme theme, XDocument chartXml, PptxSceneChart? sceneChart, ChartWorkbookData? workbook, bool plotVisibleOnly, PresentationFontResolver? fontResolver, ChartTextStyle legendTextStyle)
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
        IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSharedChartSeriesNames(plot, plotElement, workbook);
        double lastCategoryLabelWidth = plotKind == PptxSceneChartPlotKind.Area
            ? (await MeasureLastCategoryLabelWidthAsync(theme, sceneChart, chartXml, plot, plotElement, workbook, plotVisibleOnly, fontResolver).ConfigureAwait(false))
            : 0d;
        double lastXLabelWidth = plotKind == PptxSceneChartPlotKind.Scatter
            ? (await MeasureScatterLastXLabelWidthAsync(theme, sceneChart, chartXml, plot, plotElement, workbook, plotVisibleOnly, fontResolver).ConfigureAwait(false))
            : 0d;
        double rightLegendWidth = (await ResolveRightLegendWidthAsync(
            seriesNames,
            legendTextStyle,
            includeAreaReserve: plotKind == PptxSceneChartPlotKind.Area,
            fontResolver: fontResolver,
            lastCategoryLabelWidth: lastCategoryLabelWidth,
            lastXLabelWidth: lastXLabelWidth).ConfigureAwait(false));
        var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);

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
                await textMeasurer.PrepareAsync(tickLabels, valueAxisTextStyle).ConfigureAwait(false);
                maxValueLabelWidth = tickLabels.Length == 0
                    ? 0d
                    : tickLabels.Max(label => textMeasurer.Measure(label, valueAxisTextStyle));
            }
        }
        else
        {
            PptxSceneChartGrouping grouping = ReadSceneOrXmlCartesianRightLegendGrouping(sceneChart, plot, chartXml, plotElement, PptxSceneChartPlotKind.Line);
            bool stacked = IsStackedChartGrouping(grouping);
            bool percentStacked = IsPercentStackedChartGrouping(grouping);
            IReadOnlyList<ChartIndexedNumberVector> seriesVectors = ReadSharedChartSeriesVectors(plot, plotElement, workbook, plotVisibleOnly);
            if (CountRenderableSeries(seriesVectors) > 0)
            {
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, plotElement).FirstOrDefault();
                XElement? valueAxisForScale = ResolveXmlValueAxisForSource(sceneChart, valueAxis, chartXml);
                explicitValueAxisScale = HasSceneOrXmlExplicitValueAxisScale(valueAxis.SceneAxis, valueAxisForScale);
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, valueAxisForScale, GetSharedLineChartValueExtents(plot, plotElement, stacked, percentStacked, workbook, plotVisibleOnly), percentStacked, useNearMaximumHeadroom: !percentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxisForScale), percentStacked);
                IReadOnlyList<double> tickValues = GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceVerticalValueTickTargetCount);
                ChartTextStyle valueAxisTextStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxisForScale, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                string[] tickLabels = tickValues
                    .Select(value => FormatSceneOrXmlChartAxisLabel(value, valueAxis.SceneAxis, valueAxisForScale, percentStacked ? "0%" : null))
                    .ToArray();
                await textMeasurer.PrepareAsync(tickLabels, valueAxisTextStyle).ConfigureAwait(false);
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
        double width = Math.Max(1d, frame.Width - leftInset - rightLegendWidth);
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

    private static ChartLayout GetBubbleChartLayout(PptxDocument document, PptxTheme theme, ShapeBounds bounds, XDocument chartXml, PptxSceneChart sceneChart, PptxColorMap colorMap)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = sceneChart.Title.Text;
        PptxSceneChartTextBodyProperties titleTextBodyProperties = sceneChart.Title.TextBodyProperties;
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

    // Tail reserve past the legend marker block and widest entry text. Line/scatter use the
    // Office-calibrated tail; area uses its calibrated fixed block plus half the last
    // category label (the swatch clears the last category label end by a fixed lead).
    private static double ComputeRightLegendReservePadding(double legendFontSize, bool includeAreaReserve, double lastCategoryLabelWidth, double lastXLabelWidth = 0d)
    {
        double markerBlock = legendFontSize * PptxChartMetricRules.LegendSideStrokeMarkerWidthFactor +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeTextGapFactor +
            legendFontSize * PptxChartMetricRules.LegendSideStrokeGapFactor;
        // A positive last-X width marks scatter legends with line-sample keys (line charts
        // and marker-only keys always pass zero): their swatch clears the overhanging edge
        // label by half of it on top of the marker block and their own tail.
        if (!includeAreaReserve && lastXLabelWidth <= 0d)
        {
            return markerBlock + PptxChartMetricRules.LineScatterRightLegendReservePadding;
        }

        if (!includeAreaReserve)
        {
            return markerBlock +
                PptxChartMetricRules.ScatterRightLegendReserveTail +
                0.5d * lastXLabelWidth;
        }

        return PptxChartMetricRules.AreaRightLegendFixedBlock +
            0.5d * lastCategoryLabelWidth;
    }

    // Width of the last visible category label for the area right-legend reserve: the
    // legend swatch clears the end of that edge label by a fixed lead, so the reserve
    // grows with half of it. A blank or missing last label contributes nothing.
    private static async ValueTask<double> MeasureLastCategoryLabelWidthAsync(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml, PptxSceneChartPlot? plot, XElement plotElement, ChartWorkbookData? workbook, bool plotVisibleOnly, PresentationFontResolver? fontResolver)
    {
        ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, plot, chartXml, plotElement);
        ChartTextStyle categoryStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, chartXml, categoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
        IReadOnlyList<ChartIndexedTextPoint?> labels = ReadSharedCategoryLabels(plot, plotElement, workbook, plotVisibleOnly);
        int skip = ResolveSceneOrXmlCategoryAxisTickLabelSkip(categoryAxis.SceneAxis, categoryAxis.XmlAxis);
        int last = -1;
        for (int i = labels.Count - 1; i >= 0; i--)
        {
            if (i % skip == 0)
            {
                last = i;
                break;
            }
        }

        if (last < 0 || labels[last] is not { } label || string.IsNullOrEmpty(label.Text))
        {
            return 0d;
        }

        return (await new ChartTextMeasurer(fontResolver, kerningEnabled: false).MeasureAsync(label.Text, categoryStyle).ConfigureAwait(false));
    }

    // Full width of the last visible X tick label for the scatter right-legend reserve
    // and content: the swatch clears that edge label (centered on the plot edge) by half
    // of it on top of the side gap. Mirrors the render-site X extents, units, style, and
    // formatting exactly so both callers observe identical doubles. Returns zero when X
    // labels are hidden or empty, and when any series hides its line (marker-only keys
    // keep the legacy luck-matched layout).
    private static async ValueTask<double> MeasureScatterLastXLabelWidthAsync(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml, PptxSceneChartPlot? plot, XElement plotElement, ChartWorkbookData? workbook, bool plotVisibleOnly, PresentationFontResolver? fontResolver)
    {
        IReadOnlyList<ChartAxisSource> valueAxes = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, plot, chartXml, plotElement);
        ChartAxisSource xValueAxis = valueAxes.Count > 0 ? valueAxes[0] : default;
        if (!IsSceneOrXmlChartAxisLabelVisible(xValueAxis.SceneAxis, xValueAxis.XmlAxis))
        {
            return 0d;
        }

        if (ReadSceneOrXmlSeriesLineHidden(plot, plotElement).Any(hidden => hidden))
        {
            return 0d;
        }

        ChartValueExtents xExtents = ReadSceneOrXmlBubbleChartValueAxisExtents(xValueAxis.SceneAxis, xValueAxis.XmlAxis, GetScatterXValueExtents(ReadSceneOrXmlScatterSeries(plot, plotElement, readBubbleSize: false, workbook: workbook, plotVisibleOnly: plotVisibleOnly)), PptxChartMetricRules.ScatterXAxisNiceTickTargetCount, preferUnitOneOverTwo: false);
        ChartValueAxisRenderOptions xOptions = ReadSceneOrXmlChartBubbleValueAxisOptions(xValueAxis.SceneAxis, xValueAxis.XmlAxis, theme, xExtents, PptxChartMetricRules.ScatterXAxisNiceTickTargetCount);
        IReadOnlyList<double> tickValues = GetChartAxisTickValuesCore(xExtents, xOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: true, valueAxisLabelsVisible: true, manualPlotLayoutApplied: false), xOptions.TickPolicy);
        if (tickValues.Count == 0)
        {
            return 0d;
        }

        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, xValueAxis.SceneAxis, chartXml, xValueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
        string lastLabel = FormatSceneOrXmlChartAxisLabel(tickValues[tickValues.Count - 1], xValueAxis.SceneAxis, xValueAxis.XmlAxis, defaultNumberFormat: null);
        if (string.IsNullOrEmpty(lastLabel))
        {
            return 0d;
        }

        return (await new ChartTextMeasurer(fontResolver, kerningEnabled: false).MeasureAsync(lastLabel, style).ConfigureAwait(false));
    }

    private static async ValueTask<double> ResolveRightLegendWidthAsync(IReadOnlyList<ChartSeriesNameRecord> seriesNames, ChartTextStyle legendTextStyle, bool includeAreaReserve, PresentationFontResolver? fontResolver, double lastCategoryLabelWidth, double lastXLabelWidth = 0d)
    {
        double legendFontSize = legendTextStyle.FontSize;
        // Legend layout measures unkerned advances like the emission (Office legend advances
        // equal natural widths; the kerned default under-reads kern-heavy names and narrows reserves).
        var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
        await textMeasurer.PrepareAsync(seriesNames.Select(name => name.ActiveName), legendTextStyle).ConfigureAwait(false);
        double maxLegendTextWidth = seriesNames.Count == 0
            ? 0d
            : seriesNames.Max(name => textMeasurer.Measure(name.ActiveName, legendTextStyle));
        double rightReserve = maxLegendTextWidth +
            ComputeRightLegendReservePadding(legendFontSize, includeAreaReserve, lastCategoryLabelWidth, lastXLabelWidth);

        return rightReserve;
    }

    private static ChartValueExtents GetRadarChartValueExtents(IReadOnlyList<ChartRadarSeries> series)
    {
        IEnumerable<double> values = series
            .SelectMany(item => item.Values)
            .Select(value => value ?? 0d);
        double maxValue = Math.Max(0d, values.DefaultIfEmpty(0d).Max());
        double minValue = Math.Min(0d, values.DefaultIfEmpty(0d).Min());
        return new ChartValueExtents(minValue, maxValue);
    }

    private static ChartValueExtents GetLineChartValueExtents(IReadOnlyList<ChartIndexedNumberVector> series, bool stacked, bool percentStacked)
    {
        IReadOnlyList<IReadOnlyList<double?>> denseSeries = DensifyChartValueSeries(series);
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
        // Unstyled marker outlines default to round joins (Office marker forensics); explicitly
        // styled markers keep their DrawingML cap/join defaults.
        ChartSeriesStroke? stroke = ChartMarkerOutlineStroke(marker, defaultStroke, null);
        DrawChartMarkerFill(graphics, x, y, marker.SymbolKind, size, fill);
        DrawChartMarkerStroke(graphics, x, y, marker.SymbolKind, size, stroke);
    }

    private static ChartSeriesStroke? ChartMarkerOutlineStroke(ChartMarkerStyle marker, RgbColor defaultStroke, double? markerOutlineWidth)
    {
        return marker.Stroke ?? new ChartSeriesStroke(defaultStroke, 1d, markerOutlineWidth ?? PptxChartMarkerMetricRules.DefaultMarkerOutlineWidth) with { Join = 1 };
    }

    // RV04: style-18 marker gradient fill: one axial shading per marker with Coords
    // spanning twice the marker height from its bottom edge, clipped to the marker
    // shape (Office paints PatternType-2 shadings over diamond/square/triangle/circle
    //, dot and dash markers; plus/x/star keep the flat path as stroked diameters).
    private static bool PaintStyleMarkerGradient(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, PptxSceneChartMarkerSymbol symbol, double x, double y, double size, IReadOnlyList<PdfShadingStop> stops)
    {
        if (size <= 0d)
        {
            return false;
        }

        double bottom = y - size / 2d;
        if (symbol != PptxSceneChartMarkerSymbol.Diamond &&
            symbol != PptxSceneChartMarkerSymbol.Square &&
            symbol != PptxSceneChartMarkerSymbol.Triangle &&
            symbol != PptxSceneChartMarkerSymbol.Circle &&
            symbol != PptxSceneChartMarkerSymbol.Dot &&
            symbol != PptxSceneChartMarkerSymbol.Dash)
        {
            return false;
        }

        RenderInChartPlotAreaClip(graphics, plotBox, () =>
        {
            graphics.SaveState();
            if (symbol == PptxSceneChartMarkerSymbol.Square)
            {
                graphics.ClipRectangle(x - size / 2d, bottom, size, size);
            }
            else if (symbol == PptxSceneChartMarkerSymbol.Circle)
            {
                graphics.ClipEllipse(x - size / 2d, bottom, size, size);
            }
            else if (symbol == PptxSceneChartMarkerSymbol.Dot)
            {
                graphics.ClipRectangle(x, y - PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d) / 2d, PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 2d), PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d));
            }
            else if (symbol == PptxSceneChartMarkerSymbol.Dash)
            {
                graphics.ClipRectangle(x - PptxChartMarkerMetricRules.QuantizeMarkerExtent(size) / 2d, y - PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d) / 2d, PptxChartMarkerMetricRules.QuantizeMarkerExtent(size), PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d));
            }
            else if (symbol == PptxSceneChartMarkerSymbol.Diamond)
            {
                graphics.ClipPolygon([(x, y + size / 2d), (x + size / 2d, y), (x, y - size / 2d), (x - size / 2d, y)]);
            }
            else
            {
                graphics.ClipPolygon([(x, y + size / 2d), (x + size / 2d, y - size / 2d), (x - size / 2d, y - size / 2d)]);
            }

            graphics.PaintAxialShading(x, bottom + 2d * size, x, bottom, stops);
            graphics.RestoreState();
        });
        return true;
    }

    private static void DrawChartMarkerInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, double x, double y, ChartMarkerStyle marker, RgbColor defaultFill, RgbColor defaultStroke, double? markerOutlineWidth = null)
    {
        if (marker.SymbolKind == PptxSceneChartMarkerSymbol.None)
        {
            return;
        }

        double size = marker.Size;
        ChartSeriesFill fill = marker.Fill ?? new ChartSeriesFill(defaultFill, 1d, null, null);
        // Unstyled marker outlines default to round joins (Office marker forensics); explicitly
        // styled markers keep their DrawingML cap/join defaults.
        ChartSeriesStroke? stroke = ChartMarkerOutlineStroke(marker, defaultStroke, markerOutlineWidth);
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
        if (fill.Alpha <= .001d) return;
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
                    // RV04: dots render as data-anchored quantized rectangles (Office paints
                    // s/2 by s/5 rects on a 0.12pt grid starting at the data point).
                    graphics.FillRectangle(x, y - PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d) / 2d, PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 2d), PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d));
                    break;
                case PptxSceneChartMarkerSymbol.Dash:
                    // RV04: dashes render as centered quantized bars (Office paints gradient
                    // bars size-wide with dot-rule heights, not stroked center lines).
                    graphics.FillRectangle(x - PptxChartMarkerMetricRules.QuantizeMarkerExtent(size) / 2d, y - PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d) / 2d, PptxChartMarkerMetricRules.QuantizeMarkerExtent(size), PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d));
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
        if (stroke is not { } markerStroke || markerStroke.Alpha <= .001d)
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
                graphics.StrokeRectangle(x - PptxChartMarkerMetricRules.QuantizeMarkerExtent(size) / 2d, y - PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d) / 2d, PptxChartMarkerMetricRules.QuantizeMarkerExtent(size), PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d));
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
                graphics.StrokeLine(x, y - size / 2d, x, y + size / 2d);
                graphics.StrokeLine(x - size / 2d, y - size / 2d, x + size / 2d, y + size / 2d);
                graphics.StrokeLine(x - size / 2d, y + size / 2d, x + size / 2d, y - size / 2d);
                break;
            case PptxSceneChartMarkerSymbol.Dot:
                graphics.StrokeRectangle(x, y - PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d) / 2d, PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 2d), PptxChartMarkerMetricRules.QuantizeMarkerExtent(size / 5d));
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

    // RV04: stars render as stroked asterisks with no visible fill (Office strokes a
    // vertical plus two diagonal diameters; its gradient fill paints nothing on the open
    // paths). Explicit star fills stay unprobed.
    private static bool IsLineOnlyChartMarker(PptxSceneChartMarkerSymbol symbol)
    {
        return symbol is PptxSceneChartMarkerSymbol.Plus or
            PptxSceneChartMarkerSymbol.X or
            PptxSceneChartMarkerSymbol.Star;
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

    private static ChartSeriesStroke ChartSeriesStrokeColor(PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, int seriesIndex, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, double defaultWidth)
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
