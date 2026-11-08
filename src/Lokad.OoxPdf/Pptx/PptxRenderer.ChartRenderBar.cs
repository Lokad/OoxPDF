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
    // Office bar/column plot clips extend past the axis-bounded plot rect on the bottom
    // and right edges; single- and multi-plot bar charts share the axis-coincident
    // geometry, so every bar/column clip takes the pad.
    private static ChartPlotBox GetBarPlotClipBox(ChartPlotBox plotBox)
    {
        return new ChartPlotBox(
            plotBox.X,
            plotBox.Y - PptxChartMetricRules.BarPlotClipPad,
            plotBox.Width + PptxChartMetricRules.BarPlotClipPad,
            plotBox.Height + PptxChartMetricRules.BarPlotClipPad);
    }

    private static void RenderBarChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, ChartLayoutBox plotAreaBox, ChartPlotBox plotBox, IReadOnlyList<ChartIndexedNumberVector> series, bool horizontalBars, ChartBarPlotOptions plotOptions, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, ChartValueAxisRenderOptions valueAxisOptions, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents, bool valueAxisLabelsVisible, bool manualPlotLayoutApplied, int? chartStyleId, ChartAxisPlacement? categoryPlacement = null, PptxRenderContext? context = null)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;

        // Clip-only pad: legends, titles, labels, and bar geometry keep plotBox.
        ChartPlotBox plotClipBox = GetBarPlotClipBox(plotBox);
        IReadOnlyList<IReadOnlyList<double?>> denseSeries = DensifyChartValueSeries(series);
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        {
            int categoryCount = Math.Max(1, denseSeries.Max(values => values.Count));
            bool stacked = IsStackedChartGrouping(plotOptions.Grouping);
            bool percentStacked = IsPercentStackedChartGrouping(plotOptions.Grouping);
            double horizontalOrigin = categoryPlacement?.CrossingValue ?? 0d;
            double barOriginX = ChartValueToPlotCoordinate(valueExtents, horizontalOrigin, plotX, plotWidth, valueAxisOptions.Reversed);
            double columnOrigin = valueAxisOptions.CrossingValue ?? 0d;
            double columnOriginY = ChartValueToPlotCoordinate(valueExtents, columnOrigin, plotY, plotHeight, horizontalBars ? false : valueAxisOptions.Reversed);
            double valueAxisCrossingY = ChartValueToPlotCoordinate(valueExtents, valueAxisOptions.CrossingValue, plotY, plotHeight, horizontalBars ? false : valueAxisOptions.Reversed);
            double valueAxisAutoTickTargetCount = GetValueAxisAutoTickTargetCount(horizontalBars, valueAxisLabelsVisible, manualPlotLayoutApplied);
            if (valueAxisOptions.MinorGridlines)
            {
                if (horizontalBars)
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotClipBox,
                        () => DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions, major: false, valueAxisAutoTickTargetCount, context: context));
                }
                else
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotClipBox,
                        () =>
                        {
                            if (valueAxisOptions.TickPolicy == ChartAxisTickPolicy.SourceMinimumIntervals)
                                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions.Units, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, major: false, valueAxisOptions.GridlineStyle.Minor, valueAxisOptions.HasMajorGridlineElement, valueAxisOptions.TickPolicy, context: context);
                            else
                                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions.Units.MinorUnit, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, major: false, valueAxisOptions.GridlineStyle.Minor, context: context);
                        });
                }
            }

            if (valueAxisOptions.MajorGridlines)
            {
                if (horizontalBars)
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotClipBox,
                        () => DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions, major: true, valueAxisAutoTickTargetCount, context: context));
                }
                else
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotClipBox,
                        () =>
                        {
                            if (valueAxisOptions.TickPolicy == ChartAxisTickPolicy.SourceMinimumIntervals)
                                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions.Units, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, major: true, valueAxisOptions.GridlineStyle.Major, valueAxisOptions.HasMajorGridlineElement, valueAxisOptions.TickPolicy, context: context);
                            else
                                DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions.Units.MajorUnit, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, major: true, valueAxisOptions.GridlineStyle.Major, context: context);
                        });
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
                    RenderInChartPlotAreaClip(graphics, plotClipBox, () => graphics.StrokeLine(plotX, axisY, plotX + plotWidth, axisY));
                }
            }

            if (axesStyle.ValueAxisVisible)
            {
                ChartSeriesStroke stroke = horizontalBars ? categoryAxisStroke : valueAxisStroke;
                if (stroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, stroke);
                    double axisX = horizontalBars
                        ? categoryPlacement is { } placement
                            ? ChartValueToPlotCoordinate(valueExtents, placement.CrossingValue, plotX, plotWidth, valueAxisOptions.Reversed)
                            : (axesStyle.CategoryAxisRightSide ? plotX + plotWidth : plotX)
                        : (axesStyle.ValueAxisRightSide ? plotX + plotWidth : plotX);
                    RenderInChartPlotAreaClip(graphics, plotClipBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                }
            }

            if (!horizontalBars && axesStyle.SecondaryValueAxis is { } secondaryValueAxisStroke)
            {
                if (secondaryValueAxisStroke.Alpha > 0.001d)
                {
                    SetChartStroke(graphics, secondaryValueAxisStroke);
                    double axisX = axesStyle.SecondaryValueAxisRightSide ? plotX + plotWidth : plotX;
                    RenderInChartPlotAreaClip(graphics, plotClipBox, () => graphics.StrokeLine(axisX, plotY, axisX, plotY + plotHeight));
                }
            }

            if (horizontalBars)
            {
                if (stacked)
                {
                    RenderStackedHorizontalBars(graphics, plotClipBox, theme, colorMap, chartPalette, plotX, plotY, plotWidth, plotHeight, denseSeries, categoryCount, valueExtents, valueAxisOptions.Reversed, percentStacked, seriesFills, pointFills, pointStrokes, plotOptions, chartStyleId);
                }
                else
                {
                    RenderClusteredHorizontalBars(graphics, plotClipBox, theme, colorMap, chartPalette, plotX, plotY, plotWidth, plotHeight, denseSeries, categoryCount, valueExtents, valueAxisOptions.Reversed, horizontalOrigin, barOriginX, seriesFills, pointFills, pointStrokes, plotOptions, chartStyleId);
                }

                return;
            }

            if (stacked)
            {
                RenderStackedColumns(graphics, plotClipBox, theme, colorMap, chartPalette, plotX, plotY, plotWidth, plotHeight, denseSeries, categoryCount, valueExtents, valueAxisOptions.Reversed, percentStacked, seriesFills, pointFills, pointStrokes, plotOptions, chartStyleId);
                return;
            }

            double categoryWidth = plotOptions.CategoryPitch(plotWidth, categoryCount);
            double barWidth = GetClusteredBarWidth(categoryWidth, denseSeries.Count, plotOptions.GapWidth, plotOptions.Overlap);
            double step = GetClusteredBarStep(barWidth, plotOptions.Overlap);
            double clusterWidth = barWidth + Math.Max(0, denseSeries.Count - 1) * step;
            for (int category = 0; category < categoryCount; category++)
            {
                for (int seriesIndex = 0; seriesIndex < denseSeries.Count; seriesIndex++)
                {
                    IReadOnlyList<double?> values = denseSeries[seriesIndex];
                    if (category >= values.Count || values[category] is not { } value)
                    {
                        continue;
                    }

                    ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, denseSeries.Count, plotOptions.VaryColors.Value, seriesFills, pointFills, values.Count, value - columnOrigin, plotOptions.InvertsNegativePoint(seriesIndex, category));
                    double centerX = ChartCategoryToPlotCoordinate(category, categoryCount, plotX, plotWidth,
                        plotOptions.CategoryPointsOnTickMarks, plotOptions.CategoryAxisReversed,
                        slotOffset: -clusterWidth / 2d + seriesIndex * step + barWidth / 2d);
                    double barX = centerX - barWidth / 2d;
                    double valueY = ChartValueToPlotCoordinate(valueExtents, value, plotY, plotHeight, valueAxisOptions.Reversed);
                    double barY = Math.Min(columnOriginY, valueY);
                    double barHeight = Math.Abs(valueY - columnOriginY);
                    bool isStyleDriven = (seriesIndex >= seriesFills.Count || seriesFills[seriesIndex] is null) && !HasExplicitChartPointFill(pointFills, seriesIndex, category);
                    if (!(isStyleDriven && TryGetStyleBarGradientStops(chartStyleId, fill, out IReadOnlyList<PdfShadingStop> stops) && PaintBarGradientRect(graphics, plotClipBox, barX, barY, barWidth, barHeight, stops)))
                    {
                        FillChartRectangleInPlotClip(graphics, plotClipBox, barX, barY, barWidth, barHeight, fill);
                    }
                    PaintStyleColumnOutline(graphics, plotClipBox, chartStyleId, !isStyleDriven, fill, barX, barY, barWidth, barHeight);
                    StrokeChartPointRectangleInPlotClip(graphics, plotClipBox, seriesIndex, category, pointStrokes, barX, barY, barWidth, barHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, value - columnOrigin, plotOptions.InvertsNegativePoint(seriesIndex, category)));
                }
            }
        }
    }

    private static async ValueTask<ChartLayout> GetBarChartLayoutAsync(
        PptxDocument document,
        PptxTheme theme,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart sceneChart,
        PptxColorMap colorMap,
        PptxSceneChartPlot barPlot,
        XElement barChart,
        ChartBarPlotOptions barOptions,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly,
        PresentationFontResolver? fontResolver,
        ChartAxisPlacement? categoryPlacement, ChartLegendLayout legend, ChartAxisPlacement? numericPlacement,
        PptxRenderContext? context = null,
        (ChartValueExtents Extents, ChartAxisUnits Units)? resolvedValueScale = null)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = sceneChart.Title.Text;
        PptxSceneChartTextBodyProperties titleTextBodyProperties = sceneChart.Title.TextBodyProperties;
        // Horizontal legend rows are admitted by the legend owner. The baseline
        // plot retains title/axis space; its legend band is applied exactly once.
        ChartLegendLayout plotLegend = UsesMeasuredBarLegend(barOptions, sceneChart, legend)
            ? legend with { Visible = false } : legend;
        (double Ascent, double Height)? titleLine = null;
        ChartAxisTitleReserveSides axisTitles = ReadDefaultAxisTitleReserveSides(sceneChart);
        if (!string.IsNullOrWhiteSpace(title) && !sceneChart.Title.Layout.HasLayout &&
            barOptions.BarDirection == PptxSceneChartBarDirection.Bar && UsesAutomaticBarFrame(barOptions, sceneChart, legend) &&
            !axisTitles.HasHorizontalTitle && !axisTitles.HasVerticalTitle)
            titleLine = await MeasureChartTitleLineAsync(theme, sceneChart, fontResolver).ConfigureAwait(false);
        ChartPlotLayout plotLayout = (await GetBarChartPlotLayoutAsync(theme, colorMap, frame, chartXml, sceneChart, barPlot, barChart, title, plotLegend, barOptions, workbook, plotVisibleOnly, fontResolver: fontResolver, ignoreManualPlotLayout: false, measuredTitle: titleLine, categoryPlacement: categoryPlacement, context: context, resolvedValueScale: resolvedValueScale, numericPlacement: numericPlacement).ConfigureAwait(false));
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, titleTextBodyProperties, legend,
            plotLayout.CategoryLabelOffsetAdjustment,
            // Resolved ranges serve admitted single/shared horizontal plots and
            // shared columns. Retain their automatic category placement too.
            plotLayout.ManualLayoutTargetKind is null && (sceneChart.Plots.Count == 1 || resolvedValueScale is not null)
                ? categoryPlacement : null)
        { AutomaticCaptionFittingClearance = plotLayout.AutomaticCaptionFittingClearance };
    }

    private static double ReadCategoryTitleFontSizeAtPosition(PptxSceneChart sceneChart, PptxSceneChartAxisPosition titlePosition)
    {
        // Horizontal axis titles render on the opposite side of their recorded position
        // (Top hangs below the plot, Bottom sits above it). The caller sizes each reserve var
        // from the title recorded on its own side: Bottom-positioned titles size bottomReserve
        // (which anchors the plot top edge y), Top-positioned titles size topReserve (which sets
        // the height and hence the bottom edge). Category-kind only: value-kind horizontal
        // titles keep legacy output, unobserved. Run sizes win over paragraph defaults,
        // mirroring emission.
        foreach (PptxSceneChartAxis axis in sceneChart.Axes)
        {
            if (axis.PositionKind != titlePosition || !IsBottomPlacedCategoryTitleCandidate(axis.AxisKind, axis.Title.Text, axis.Title.Layout.HasLayout))
            {
                continue;
            }

            foreach (PptxSceneChartTextRun run in axis.Title.TextRuns)
            {
                if (run.TextStyle.FontSize is { } runSize && runSize > 0d)
                {
                    return runSize;
                }
            }

            ChartTextStyleOverride titleStyle = ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartElementTextStyleOverride(sceneChart, axis.Title.TextStyle, GetChartAxisStyleRole(axis.AxisKind)));
            if (titleStyle.FontSize is { } sceneSize && sceneSize > 0d)
            {
                return sceneSize;
            }

            return 0d;
        }

        return 0d;
    }

    private static bool IsBottomPlacedCategoryTitleCandidate(PptxSceneChartAxisKind axisKind, string? titleText, bool hasManualLayout)
    {
        return (axisKind is PptxSceneChartAxisKind.Category or PptxSceneChartAxisKind.Date or PptxSceneChartAxisKind.Series)
            && !string.IsNullOrWhiteSpace(titleText)
            && !hasManualLayout;
    }

    private static double ComputeDefaultAxisTitleVerticalBottomReserve(double categoryLabelFontSize, double categoryTitleFontSize)
    {
        // Additive composition confirmed by two single-variable topright knots (title 12 to 18
        // lands 54.09 vs 54.12 predicted, labels 9 to 14.04 land 56.00 vs 56.03): the horizontal-top
        // law evaluated at the category-label size, re-based from its baked-in 12pt title to the
        // actual category-title size via the horizontal-bottom slope.
        (double topAtLabelSize, _) = ComputeDefaultAxisTitleHorizontalBarReserves(categoryLabelFontSize, PptxChartMetricRules.DefaultAxisTitleHorizontalBarReferenceTitleFontSize);
        return topAtLabelSize + PptxChartMetricRules.DefaultAxisTitleHorizontalBarBottomReservePerTitleFontSize * (categoryTitleFontSize - PptxChartMetricRules.DefaultAxisTitleHorizontalBarReferenceTitleFontSize);
    }

    private static (double TopReserve, double BottomReserve) ComputeDefaultAxisTitleHorizontalBarReserves(double valueTickFontSize, double chartTitleFontSize)
    {
        return (
            PptxChartMetricRules.DefaultAxisTitleHorizontalBarTopReserveBase + PptxChartMetricRules.DefaultAxisTitleHorizontalBarTopReservePerTickFontSize * valueTickFontSize,
            PptxChartMetricRules.DefaultAxisTitleHorizontalBarBottomReserveBase + PptxChartMetricRules.DefaultAxisTitleHorizontalBarBottomReservePerTitleFontSize * chartTitleFontSize);
    }

    private static async ValueTask<ChartPlotLayout> GetBarChartPlotLayoutAsync(
        PptxTheme theme,
        PptxColorMap colorMap,
        ChartFrameBox frame,
        XDocument chartXml,
        PptxSceneChart sceneChart,
        PptxSceneChartPlot barPlot,
        XElement barChart,
        string? title,
        ChartLegendLayout legend,
        ChartBarPlotOptions barOptions,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly,
        bool ignoreManualPlotLayout,
        PresentationFontResolver? fontResolver,
        (double Ascent, double Height)? measuredTitle = null,
        ChartAxisPlacement? categoryPlacement = null, PptxRenderContext? context = null,
        (ChartValueExtents Extents, ChartAxisUnits Units)? resolvedValueScale = null, ChartAxisPlacement? numericPlacement = null)
    {
        bool hasTitle = !string.IsNullOrWhiteSpace(title);
        bool hasLegend = legend.Visible && !legend.Overlay;
        bool horizontalBars = barOptions.BarDirection == PptxSceneChartBarDirection.Bar;
        double? categoryLabelOffsetAdjustment = null;
        double? categoryFittingClearance = null;
        double? categoryTopReserve = null;
        bool usesMeasuredAxisInsets = !sceneChart.PlotAreaLayout.HasLayout &&
            (!hasLegend || legend.PositionKind is PptxSceneChartLegendPosition.Top or PptxSceneChartLegendPosition.Bottom);
        // Match the bar painter's order: columns first, then line overlays.
        // Retain each owning plot instead of rediscovering it from an ID.
        var visibleValueAxes = new List<(ChartAxisSource Axis, PptxSceneChartPlot Owner)>();
        var visibleValueAxisIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var owner in ReadSceneChartPlots(sceneChart, PptxSceneChartPlotKind.Bar)
            .Concat(ReadSceneChartPlots(sceneChart, PptxSceneChartPlotKind.Line)))
        {
            context?.CancellationToken.ThrowIfCancellationRequested();
            if (CountRenderableSeries(ReadSharedChartSeriesVectors(owner, workbook, plotVisibleOnly)) == 0)
                continue;
            ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, owner)
                .FirstOrDefault();
            if (!IsSceneVisibleValueAxis(valueAxis.SceneAxis) ||
                !IsSceneChartAxisLabelVisible(valueAxis.SceneAxis))
            {
                continue;
            }

            string axisId = valueAxis.SceneAxis?.Id ?? string.Empty;
            if (visibleValueAxisIds.Add(axisId))
            {
                visibleValueAxes.Add((valueAxis, owner));
            }
        }

        ChartPlotBox defaultPlotBox;
        if (!hasTitle && !hasLegend)
        {
            defaultPlotBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.BarOverlayOnly);
        }
        else if (!hasTitle && legend.PositionKind == PptxSceneChartLegendPosition.Bottom)
        {
            defaultPlotBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.BarNoTitleBottomLegend);
            defaultPlotBox = AdjustNoTitleBottomLegendPlotBoxForMeasuredContent(defaultPlotBox, frame, theme, colorMap, sceneChart, chartXml, barPlot, barChart, horizontalBars, usesMeasuredAxisInsets);
        }
        else if (horizontalBars && hasTitle && !hasLegend)
        {
            defaultPlotBox = GetChartPlotBoxPreset(frame, ChartPlotBoxPreset.HorizontalBarTitleNoLegend);
        }
        else if (hasTitle && !hasLegend && HasInsideValueAxisCrossing())
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

        if (usesMeasuredAxisInsets)
        {
            // Titles and horizontal legends reserve vertical space. In both
            // orientations, horizontal edges belong to the measured axis strips; retain
            // the explicit manual and side-legend placement contracts.
            double inset = PptxChartMetricRules.ColumnPlotRightReserve;
            defaultPlotBox = defaultPlotBox with { X = frame.X + inset, Width = Math.Max(1d, frame.Width - 2d * inset) };
        }

        (double Reserve, ChartAxisSource ValueAxis, double NumericClearance)? cartesianHeader = horizontalBars ? null :
            await TryMeasureCartesianHeaderAsync(theme, sceneChart, legend, barPlot, barChart, chartXml, fontResolver, categoryPlacement).ConfigureAwait(false);
        defaultPlotBox = (await AdjustBarChartPlotBoxForStackedValueAxisLabelsAsync(defaultPlotBox).ConfigureAwait(false));
        defaultPlotBox = (await AdjustBarChartPlotBoxForSingleValueAxisLabelsAsync(defaultPlotBox).ConfigureAwait(false));
        // The combined reserve is final: a single-axis preset may shrink its own
        // band, but cannot reclaim space occupied by another painted axis.
        defaultPlotBox = (await AdjustBarChartPlotBoxForVisibleValueAxesAsync(defaultPlotBox).ConfigureAwait(false));
        defaultPlotBox = await AdjustBarChartPlotBoxForCategoryLabelsAsync(defaultPlotBox).ConfigureAwait(false);
        defaultPlotBox = AdjustStackedColumnBottomLegendPlotBox(defaultPlotBox, frame, horizontalBars, barOptions.Grouping, hasTitle, legend);
        defaultPlotBox = AdjustBarChartPlotBoxForDefaultAxisTitles(defaultPlotBox, horizontalBars, hasLegend);
        defaultPlotBox = (await AdjustHorizontalBarPlotBoxForCaptionsAsync(defaultPlotBox).ConfigureAwait(false));
        defaultPlotBox = (await AdjustHorizontalBarPlotBoxForValueLabelsAsync(defaultPlotBox).ConfigureAwait(false));
        defaultPlotBox = AdjustHorizontalBarPlotBoxForTopAndBottom(defaultPlotBox);
        if (cartesianHeader is null)
            defaultPlotBox = AdjustVerticalBarPlotBoxTopFloor(defaultPlotBox, frame, horizontalBars, hasTitle, hasLegend, HasRenderableCategoryLabels(barPlot, barChart, workbook, plotVisibleOnly));
        ChartAxisSource rightReserveCategoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
        bool rightReserveLabelsVisible = IsSceneChartAxisLabelVisible(rightReserveCategoryAxis.SceneAxis);
        defaultPlotBox = AdjustVerticalBarPlotBoxRightReserve(defaultPlotBox, frame, horizontalBars, legend, HasRenderableCategoryLabels(barPlot, barChart, workbook, plotVisibleOnly) && rightReserveLabelsVisible);
        if (cartesianHeader is { } measuredHeader)
        {
            double top = frame.Y + frame.Height - Math.Max(measuredHeader.Reserve, categoryTopReserve ?? 0d);
            defaultPlotBox = defaultPlotBox with { Height = Math.Max(1d, top - defaultPlotBox.Y) };
        }
        else if (measuredTitle is { } measuredLine)
        {
            // Title box and plot gap share the frame with axes and legend rows.
            // Overlay titles use the same painted baseline and consume no plot space.
            if (sceneChart.Title.Overlay != true)
            {
                double band = PptxChartMetricRules.ChartContentFrameInset + 2d * PptxChartMetricRules.ChartTitleLinePadding +
                    measuredLine.Height + PptxChartMetricRules.HorizontalBarPlotTopReserve;
                double top = frame.Y + frame.Height - band;
                defaultPlotBox = defaultPlotBox with { Height = Math.Max(1d, top - defaultPlotBox.Y) };
            }
        }
        else if (hasTitle && !sceneChart.Title.Layout.HasLayout && sceneChart.Title.Overlay != true)
        {
            // Default placement anchors the title above the plot. The same font
            // metrics used for chart text must fit in the remaining frame band;
            // frame-ratio presets alone can put the entire title above the clip.
            ChartTextStyle titleStyle = ReadChartTitleTextStyle(theme, sceneChart);
            var titleMeasurer = new ChartTextMeasurer(fontResolver);
            double lineHeight = (await titleMeasurer.MeasureNormalLineHeightAsync(titleStyle).ConfigureAwait(false));
            foreach (ChartTextRunOverride run in ReadChartTitleTextRuns(sceneChart))
            {
                lineHeight = Math.Max(lineHeight, (await titleMeasurer.MeasureNormalLineHeightAsync(titleStyle.Merge(run.TextStyle)).ConfigureAwait(false)));
            }
            double titleOffset = titleStyle.FontSize * (!horizontalBars && sceneChart.Title.IsAutoGenerated
                ? PptxChartMetricRules.AutoBarTitleAbovePlotBaselineOffsetFactor
                : PptxChartMetricRules.TitleAbovePlotBaselineOffsetFactor);
            double top = Math.Min(defaultPlotBox.Y + defaultPlotBox.Height, frame.Y + frame.Height - lineHeight - titleOffset);
            defaultPlotBox = defaultPlotBox with { Height = Math.Max(1d, top - defaultPlotBox.Y) };
        }
        if (numericPlacement is { LabelsHighSide: true })
        {
            ChartAxisSource numericAxis = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
            if (IsSceneChartAxisLabelVisible(numericAxis.SceneAxis))
            {
                // Transfer only the numeric band. Title and legend owners keep
                // their measured space; plot height and category geometry stay.
                ChartTextStyle numericStyle = ReadSceneChartTextStyle(theme, sceneChart, numericAxis.SceneAxis, PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
                double transfer = ComputeCartesianLabelStripBottomReserve(numericStyle.FontSize) -
                    PptxChartMetricRules.HorizontalBarPlotTopReserve;
                defaultPlotBox = defaultPlotBox with { Y = defaultPlotBox.Y - transfer };
            }
        }
        if (ignoreManualPlotLayout)
        {
            return ChartPlotLayout.FromPlotBox(defaultPlotBox) with
            { CategoryLabelOffsetAdjustment = categoryLabelOffsetAdjustment, AutomaticCaptionFittingClearance = categoryFittingClearance };
        }

        ChartPlotBox manualDefaultPlotBox = horizontalBars && sceneChart.PlotAreaLayout is
            { HasLayout: true, LayoutTargetKind: not PptxSceneChartManualLayoutTarget.Unknown }
            ? GetHorizontalBarManualLayoutTargetDefaultPlotBox(defaultPlotBox)
            : defaultPlotBox;
        if (!TryBuildManualPlotLayout(sceneChart.PlotAreaLayout, frame, manualDefaultPlotBox, out ChartPlotLayout manualPlotLayout))
        {
            return ChartPlotLayout.FromPlotBox(defaultPlotBox) with
            { CategoryLabelOffsetAdjustment = categoryLabelOffsetAdjustment, AutomaticCaptionFittingClearance = categoryFittingClearance };
        }

        // Bars retain their dedicated outer-area placement and caption owner.
        if (sceneChart.PlotAreaLayout.LayoutTargetKind == PptxSceneChartManualLayoutTarget.Inner)
            manualPlotLayout = ResolveCartesianPlotPosition(sceneChart.PlotAreaLayout, frame, manualPlotLayout);
        return (await ResolveBarManualPlotLayoutTargetAsync(theme, chartXml, sceneChart, barPlot, barChart, manualPlotLayout, horizontalBars, frame, workbook, plotVisibleOnly, fontResolver).ConfigureAwait(false));

        ChartPlotBox GetHorizontalBarManualLayoutTargetDefaultPlotBox(ChartPlotBox defaultPlotBox)
        {
            double x = frame.X + frame.Width * PptxChartMetricRules.HorizontalBarManualLayoutTargetPlotBoxXRatio;
            double y = frame.Y + frame.Height * PptxChartMetricRules.HorizontalBarManualLayoutTargetPlotBoxYRatio;
            return new ChartPlotBox(x, y, defaultPlotBox.Width, defaultPlotBox.Height);
        }

        ChartPlotBox AdjustBarChartPlotBoxForDefaultAxisTitles(ChartPlotBox plotBox, bool horizontalBars, bool hasLegend)
        {
            if (hasLegend)
            {
                return plotBox;
            }

            ChartAxisTitleReserveSides reserveSides = ReadDefaultAxisTitleReserveSides(sceneChart);
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
                (double horizontalTopReserve, double horizontalBottomReserve) = ResolveDefaultAxisTitleHorizontalBarReserves(frame.Height * PptxChartMetricRules.DefaultAxisTitleHorizontalBarPlotBandReserveRatio);
                double horizontalX = frame.X + horizontalLeftReserve;
                double horizontalY = frame.Y + horizontalTopReserve;
                double horizontalWidth = Math.Max(1d, frame.Width - horizontalLeftReserve - horizontalRightReserve);
                double horizontalHeight = Math.Max(1d, frame.Height - horizontalTopReserve - horizontalBottomReserve);
                return new ChartPlotBox(horizontalX, horizontalY, horizontalWidth, horizontalHeight);
            }
            (double TopReserve, double BottomReserve) ResolveDefaultAxisTitleHorizontalBarReserves(double legacyBandReserve)
            {
                double topReserve = legacyBandReserve;
                double bottomReserve = legacyBandReserve;
                ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
                if (IsSceneChartAxisLabelVisible(valueAxis.SceneAxis))
                {
                    ChartTextStyle valueTickStyle = ReadSceneChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                    topReserve = ComputeDefaultAxisTitleHorizontalBarReserves(valueTickStyle.FontSize, 0d).TopReserve;
                }

                if (hasTitle)
                {
                    ChartTextStyle titleStyle = ReadChartTitleTextStyle(theme, sceneChart);
                    bottomReserve = ComputeDefaultAxisTitleHorizontalBarReserves(0d, titleStyle.FontSize).BottomReserve;
                }

                return (topReserve, bottomReserve);
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
            if (!horizontalBars)
            {
                bottomReserve = ResolveDefaultAxisTitleVerticalSideReserve(bottomReserve, PptxSceneChartAxisPosition.Bottom);
                topReserve = ResolveDefaultAxisTitleVerticalSideReserve(topReserve, PptxSceneChartAxisPosition.Top);
            }

            double ResolveDefaultAxisTitleVerticalSideReserve(double legacyReserve, PptxSceneChartAxisPosition titlePosition)
            {
                ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
                if (!IsSceneChartAxisLabelVisible(categoryAxis.SceneAxis))
                {
                    return legacyReserve;
                }

                ChartTextStyle categoryTickStyle = ReadSceneChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                double categoryTitleFontSize = ReadCategoryTitleFontSizeAtPosition(sceneChart, titlePosition);
                if (categoryTitleFontSize <= 0d)
                {
                    return legacyReserve;
                }

                return ComputeDefaultAxisTitleVerticalBottomReserve(categoryTickStyle.FontSize, categoryTitleFontSize);
            }

            double x = frame.X + leftReserve;
            double y = frame.Y + bottomReserve;
            double width = Math.Max(1d, frame.Width - leftReserve - rightReserve);
            double height = Math.Max(1d, frame.Height - bottomReserve - topReserve);
            return new ChartPlotBox(x, y, width, height);
        }

        async ValueTask<ChartPlotBox> AdjustHorizontalBarPlotBoxForValueLabelsAsync(ChartPlotBox plotBox)
        {
            if (!horizontalBars || usesMeasuredAxisInsets && categoryPlacement is not null)
            {
                return plotBox;
            }

            ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
            if (!IsSceneChartAxisLabelVisible(valueAxis.SceneAxis))
            {
                return plotBox;
            }

            IReadOnlyList<ChartIndexedNumberVector> seriesVectors = ReadSharedChartSeriesVectors(barPlot, workbook, plotVisibleOnly);
            if (CountRenderableSeries(seriesVectors) == 0)
            {
                return plotBox;
            }

            bool percentStacked = IsPercentStackedChartGrouping(barOptions.Grouping);
            ChartValueExtents valueExtents = resolvedValueScale?.Extents ?? ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, GetSharedBarChartValueExtents(barPlot, barOptions.Grouping, workbook, plotVisibleOnly), percentStacked, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            ChartValueAxisRenderOptions valueOptions = ReadSceneChartValueAxisRenderOptions(valueAxis.SceneAxis, valueExtents, percentStacked);
            ChartAxisUnits axisUnits = resolvedValueScale?.Units ?? valueOptions.Units;
            ChartTextStyle tickStyle = ReadSceneChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
            var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
            double maxLabelWidth = 0d;
            foreach (double tickValue in GetChartAxisTickValuesCore(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceHorizontalValueTickTargetCount,
                ReadCartesianChartAxisTickPolicy(sceneChart, valueAxis.SceneAxis, valueExtents, percentStacked), context: context).Append(valueExtents.Min).Append(valueExtents.Max))
            {
                string label = FormatSceneChartAxisLabel(tickValue, valueAxis.SceneAxis, valueAxis.Date1904, percentStacked ? "0%" : null);
                if (!string.IsNullOrWhiteSpace(label))
                {
                    maxLabelWidth = Math.Max(maxLabelWidth, (await textMeasurer.MeasureAsync(label, tickStyle).ConfigureAwait(false)));
                }
            }

            if (maxLabelWidth <= 0d)
            {
                return plotBox;
            }

            double leftReserve = plotBox.X - frame.X;
            if (usesMeasuredAxisInsets)
            {
                double endpoint = valueOptions.Reversed ? valueExtents.Max : valueExtents.Min;
                string label = FormatSceneChartAxisLabel(endpoint, valueAxis.SceneAxis, valueAxis.Date1904, percentStacked ? "0%" : null);
                double endpointWidth = string.IsNullOrWhiteSpace(label) ? 0d : await textMeasurer.MeasureAsync(label, tickStyle).ConfigureAwait(false);
                leftReserve = Math.Max(leftReserve, endpointWidth / 2d + PptxChartMetricRules.HorizontalBarValueAxisRightPadding);
            }

            double rightReserve = ResolveMeasuredRightReserve(frame.X + frame.Width - plotBox.X - plotBox.Width, maxLabelWidth);
            double right = frame.X + frame.Width - rightReserve;
            double x = frame.X + leftReserve;
            double width = Math.Max(1d, right - x);
            return new ChartPlotBox(x, plotBox.Y, width, plotBox.Height);
        }

        ChartPlotBox AdjustHorizontalBarPlotBoxForTopAndBottom(ChartPlotBox plotBox)
        {
            if (!horizontalBars)
            {
                return plotBox;
            }

            bool usesAutomaticTop = usesMeasuredAxisInsets && !hasTitle &&
                (!hasLegend || legend.PositionKind == PptxSceneChartLegendPosition.Bottom);
            double automaticTop = frame.Y + frame.Height - PptxChartMetricRules.HorizontalBarPlotTopReserve;
            bool measuredFrame = usesAutomaticTop || measuredTitle is not null;
            if ((hasTitle || hasLegend || (legend.Visible && !usesMeasuredAxisInsets)) && measuredTitle is null)
            {
                return usesAutomaticTop ? plotBox with { Height = Math.Max(1d, automaticTop - plotBox.Y) } : plotBox;
            }

            ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
            if (!IsSceneChartAxisLabelVisible(valueAxis.SceneAxis))
            {
                double bottom = frame.Y + PptxChartMetricRules.HorizontalBarPlotTopReserve;
                return measuredFrame ? new ChartPlotBox(plotBox.X, bottom, plotBox.Width, Math.Max(1d, automaticTop - bottom)) : plotBox;
            }

            IReadOnlyList<ChartIndexedNumberVector> topBottomSeriesVectors = ReadSharedChartSeriesVectors(barPlot, workbook, plotVisibleOnly);
            if (CountRenderableSeries(topBottomSeriesVectors) == 0)
            {
                return plotBox;
            }

            bool percentBottomStacked = IsPercentStackedChartGrouping(barOptions.Grouping);
            ChartValueExtents bottomValueExtents = resolvedValueScale?.Extents ?? ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, GetSharedBarChartValueExtents(barPlot, barOptions.Grouping, workbook, plotVisibleOnly), percentBottomStacked, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            ChartAxisUnits bottomAxisUnits = resolvedValueScale?.Units ?? ReadSceneChartValueAxisRenderOptions(valueAxis.SceneAxis, bottomValueExtents, percentBottomStacked).Units;
            ChartTextStyle bottomTickStyle = ReadSceneChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
            bool hasBottomValueLabel = false;
            foreach (double bottomTickValue in GetChartAxisTickValuesCore(bottomValueExtents, bottomAxisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceHorizontalValueTickTargetCount,
                ReadCartesianChartAxisTickPolicy(sceneChart, valueAxis.SceneAxis, bottomValueExtents, percentBottomStacked), context: context))
            {
                string bottomLabel = FormatSceneChartAxisLabel(bottomTickValue, valueAxis.SceneAxis, valueAxis.Date1904, percentBottomStacked ? "0%" : null);
                if (!string.IsNullOrWhiteSpace(bottomLabel))
                {
                    hasBottomValueLabel = true;
                    break;
                }
            }

            if (!hasBottomValueLabel)
            {
                return plotBox;
            }

            double requiredBottomReserve = ComputeCartesianLabelStripBottomReserve(bottomTickStyle.FontSize);
            double presetBottomReserve = plotBox.Y - frame.Y;
            double newY = measuredFrame ? frame.Y + requiredBottomReserve : Math.Abs(presetBottomReserve - requiredBottomReserve) < PptxChartMetricRules.HorizontalBarPlotFloorSlop
                ? plotBox.Y
                : frame.Y + requiredBottomReserve;
            double presetTopReserve = frame.Y + frame.Height - plotBox.Y - plotBox.Height;
            double newTop = measuredFrame ? automaticTop : Math.Abs(presetTopReserve - PptxChartMetricRules.HorizontalBarPlotTopReserve) < PptxChartMetricRules.HorizontalBarPlotFloorSlop
                ? plotBox.Y + plotBox.Height
                : frame.Y + frame.Height - PptxChartMetricRules.HorizontalBarPlotTopReserve;
            return new ChartPlotBox(plotBox.X, newY, plotBox.Width, Math.Max(1d, newTop - newY));
        }

        async ValueTask<ChartPlotBox> AdjustHorizontalBarPlotBoxForCaptionsAsync(ChartPlotBox plotBox)
        {
            if (!horizontalBars)
            {
                return plotBox;
            }

            ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
            ChartTextStyle tickStyle = ReadSceneChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
            var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
            double maxCategoryWidth = 0d;
            if (IsSceneChartAxisLabelVisible(categoryAxis.SceneAxis))
            foreach (ChartIndexedTextPoint? label in ReadSharedCategoryLabels(barPlot, workbook, plotVisibleOnly))
            {
                string? labelText = label?.Text;
                if (!string.IsNullOrWhiteSpace(labelText))
                {
                    maxCategoryWidth = Math.Max(maxCategoryWidth, (await textMeasurer.MeasureAsync(labelText, tickStyle).ConfigureAwait(false)));
                }
            }

            if (usesMeasuredAxisInsets && categoryPlacement is { } placement)
            {
                double endpointLeft = PptxChartMetricRules.ColumnPlotRightReserve, endpointRight = endpointLeft;
                ChartAxisSource value = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
                if (IsSceneChartAxisLabelVisible(value.SceneAxis))
                {
                    bool percent = IsPercentStackedChartGrouping(barOptions.Grouping);
                    ChartValueExtents extents = resolvedValueScale?.Extents ?? ReadCartesianChartValueAxisExtents(sceneChart, value.SceneAxis,
                        GetSharedBarChartValueExtents(barPlot, barOptions.Grouping, workbook, plotVisibleOnly), percent, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                    ChartTextStyle style = ReadSceneChartTextStyle(theme, sceneChart, value.SceneAxis,
                        PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
                    double first = await textMeasurer.MeasureAsync(FormatSceneChartAxisLabel(extents.Min, value.SceneAxis, value.Date1904, percent ? "0%" : null), style).ConfigureAwait(false);
                    double last = await textMeasurer.MeasureAsync(FormatSceneChartAxisLabel(extents.Max, value.SceneAxis, value.Date1904, percent ? "0%" : null), style).ConfigureAwait(false);
                    if (ReadSceneValueAxisReversed(value.SceneAxis)) (first, last) = (last, first);
                    endpointLeft = Math.Max(endpointLeft, first / 2d + PptxChartMetricRules.HorizontalBarValueAxisRightPadding);
                    endpointRight = Math.Max(endpointRight, last / 2d + PptxChartMetricRules.HorizontalBarValueAxisRightPadding);
                }
                double side = maxCategoryWidth > 0d ? ComputeCartesianAxisLabelSideReserve(maxCategoryWidth, tickStyle.FontSize) : 0d;
                return AdmitChartPlotCaptionReserves(frame, plotBox, endpointLeft, endpointRight, side, placement);
            }

            if (maxCategoryWidth <= 0d)
            {
                return plotBox;
            }

            double requiredReserve = usesMeasuredAxisInsets
                ? ComputeCartesianAxisLabelSideReserve(maxCategoryWidth, tickStyle.FontSize)
                : maxCategoryWidth + PptxChartMetricRules.LineRightLegendValueAxisPadding;
            double leftReserve = plotBox.X - frame.X;
            double rightReserve = frame.X + frame.Width - plotBox.X - plotBox.Width;
            bool labelsRight = ResolveSceneCategoryAxisRightSide(categoryAxis.SceneAxis, defaultRightSide: false);
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

        async ValueTask<ChartPlotBox> AdjustBarChartPlotBoxForSingleValueAxisLabelsAsync(ChartPlotBox plotBox)
        {
            if (horizontalBars || visibleValueAxes.Count > 1)
            {
                return plotBox;
            }

            if (IsStackedChartGrouping(barOptions.Grouping))
            {
                return plotBox;
            }

            ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
            if (!IsSceneChartAxisLabelVisible(valueAxis.SceneAxis))
            {
                return plotBox;
            }

            IReadOnlyList<ChartIndexedNumberVector> seriesVectors = ReadSharedChartSeriesVectors(barPlot, workbook, plotVisibleOnly);
            if (CountRenderableSeries(seriesVectors) == 0)
            {
                return plotBox;
            }

            ChartValueExtents valueExtents = resolvedValueScale?.Extents ?? ReadCartesianChartValueAxisExtents(sceneChart, valueAxis.SceneAxis, GetSharedBarChartValueExtents(barPlot, barOptions.Grouping, workbook, plotVisibleOnly), false, ResolveBarValueAxisHeadroom(horizontalBars, IsPercentStackedChartGrouping(barOptions.Grouping)), PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            ChartAxisUnits axisUnits = resolvedValueScale?.Units ?? ResolvePercentStackedAxisUnits(ReadSceneChartValueAxisUnits(valueAxis.SceneAxis), false);
            ChartTextStyle tickStyle = ReadSceneChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
            var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
            double maxLabelWidth = 0d;
            // Bounds still reserve their advance when no source interval lands there.
            foreach (double tickValue in GetChartAxisTickValuesCore(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceVerticalValueTickTargetCount, ReadCartesianChartAxisTickPolicy(sceneChart, valueAxis.SceneAxis, valueExtents), context: context)
                .Concat(new[] { valueExtents.Min, valueExtents.Max }))
            {
                string label = FormatSceneChartAxisLabel(tickValue, valueAxis.SceneAxis, valueAxis.Date1904, defaultNumberFormat: null);
                if (!string.IsNullOrWhiteSpace(label))
                {
                    maxLabelWidth = Math.Max(maxLabelWidth, (await textMeasurer.MeasureAsync(label, tickStyle).ConfigureAwait(false)));
                }
            }

            if (maxLabelWidth <= 0d)
            {
                return plotBox;
            }

            double requiredReserve = ComputeCartesianAxisLabelSideReserve(maxLabelWidth, tickStyle.FontSize);
            double leftReserve = plotBox.X - frame.X;
            double rightReserve = frame.X + frame.Width - plotBox.X - plotBox.Width;
            bool labelsRight = ResolveSceneValueAxisLabelsRightSide(sceneChart, valueAxis.SceneAxis,
                ResolveSceneValueAxisRightSide(sceneChart, valueAxis.SceneAxis, defaultRightSide: false));
            if (labelsRight)
            {
                rightReserve = Math.Max(rightReserve, requiredReserve);
            }
            else if (hasTitle || hasLegend || legend.Visible || leftReserve - requiredReserve < PptxChartMetricRules.BarValueAxisPresetFloorSlop)
            {
                leftReserve = Math.Max(leftReserve, requiredReserve);
            }
            else
            {
                // Office measures the strip (frame + 6.5 indent + tick width + 0.925 fs)
                // rather than flooring at the preset (legend-keys probe: preset exceeds measured
                // by 4.45 against Office 20.44); layouts with titles or any legend, and presets
                // within slop, keep the legacy floor (column-clustered agrees within 0.06).
                double presetLeftReserve = leftReserve;
                leftReserve = requiredReserve;
                rightReserve = Math.Max(rightReserve, rightReserve + (presetLeftReserve - requiredReserve));
            }

            double x = frame.X + leftReserve;
            double right = frame.X + frame.Width - rightReserve;
            double width = Math.Max(1d, right - x);
            return new ChartPlotBox(x, plotBox.Y, width, plotBox.Height);
        }

        async ValueTask<ChartPlotBox> AdjustBarChartPlotBoxForCategoryLabelsAsync(ChartPlotBox plotBox)
        {
            if (horizontalBars)
            {
                return plotBox;
            }

            ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
            bool labelsVisible = IsSceneChartAxisLabelVisible(categoryAxis.SceneAxis);
            // Reserve and painting share the category offset measurement, under
            // the existing frame/source-furniture eligibility.
            bool measuredFrame = cartesianHeader is not null;
            if ((hasTitle || hasLegend || legend.Visible) && !measuredFrame)
            {
                return plotBox;
            }

            double requiredBottomReserve;
            double? measuredOffsetAdjustment = null;
            if (!labelsVisible)
            {
                if (!measuredFrame) return plotBox;
                // Hidden categories contribute no text band; the numeric end
                // label keeps the same clearance used at the opposite end.
                requiredBottomReserve = cartesianHeader!.Value.NumericClearance;
            }
            else
            {
                IReadOnlyList<ChartIndexedNumberVector> categorySeriesVectors = ReadSharedChartSeriesVectors(barPlot, workbook, plotVisibleOnly);
                if (CountRenderableSeries(categorySeriesVectors) == 0 || !HasRenderableCategoryLabels(barPlot, barChart, workbook, plotVisibleOnly))
                {
                    return plotBox;
                }

                ChartTextStyle categoryTickStyle = ReadSceneChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                requiredBottomReserve = ComputeCartesianLabelStripBottomReserve(categoryTickStyle.FontSize);
                double offsetScale = ResolveSceneCategoryAxisLabelOffsetScale(categoryAxis.SceneAxis);
                if (measuredFrame && offsetScale != 1d)
                {
                    var measurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
                    measuredOffsetAdjustment = await TryMeasureCategoryLabelOffsetAdjustmentAsync(measurer, categoryTickStyle, offsetScale).ConfigureAwait(false);
                    if (measuredOffsetAdjustment is { } adjustment)
                        requiredBottomReserve += adjustment;
                    else
                    {
                        measuredFrame = false;
                        if (hasTitle || hasLegend || legend.Visible) return plotBox;
                    }
                }
            }
            if (measuredFrame && labelsVisible && categoryPlacement is { } placement)
            {
                var headerMetrics = cartesianHeader!.Value;
                var reserves = ResolveCartesianCategoryReserves(frame.Height, headerMetrics.Reserve, requiredBottomReserve, headerMetrics.NumericClearance, placement);
                double height = frame.Height - reserves.Top - reserves.Bottom;
                if (height < 1d) return plotBox;
                categoryTopReserve = reserves.Top;
                categoryLabelOffsetAdjustment = measuredOffsetAdjustment;
                categoryFittingClearance = await MeasureCategoryFittingClearanceAsync(theme, sceneChart, chartXml, fontResolver, headerMetrics.NumericClearance).ConfigureAwait(false);
                return new(plotBox.X, frame.Y + reserves.Bottom, plotBox.Width, height);
            }
            double presetBottomReserve = plotBox.Y - frame.Y;
            if (!measuredFrame && Math.Abs(presetBottomReserve - requiredBottomReserve) < PptxChartMetricRules.BarValueAxisPresetFloorSlop)
            {
                return plotBox;
            }

            double newY = frame.Y + requiredBottomReserve;
            double plotTop = plotBox.Y + plotBox.Height;
            double admittedTop = cartesianHeader is { } header ? frame.Y + frame.Height - header.Reserve : plotTop;
            if (measuredOffsetAdjustment is not null && admittedTop - newY < 1d) return plotBox;
            categoryLabelOffsetAdjustment = measuredOffsetAdjustment;
            categoryFittingClearance = measuredFrame ? await MeasureCategoryFittingClearanceAsync(theme, sceneChart, chartXml, fontResolver, cartesianHeader!.Value.NumericClearance).ConfigureAwait(false) : null;
            return new ChartPlotBox(plotBox.X, newY, plotBox.Width, Math.Max(1d, plotTop - newY));
        }

        async ValueTask<ChartPlotBox> AdjustBarChartPlotBoxForStackedValueAxisLabelsAsync(ChartPlotBox plotBox)
        {
            bool stackedHorizontalBars = barOptions.BarDirection == PptxSceneChartBarDirection.Bar;
            if (stackedHorizontalBars)
            {
                return plotBox;
            }

            PptxSceneChartGrouping grouping = barOptions.Grouping;
            bool percentStacked = IsPercentStackedChartGrouping(grouping);
            if (!IsStackedChartGrouping(grouping))
            {
                return plotBox;
            }

            ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
            if (!IsSceneChartAxisLabelVisible(valueAxis.SceneAxis))
            {
                return plotBox;
            }

            IReadOnlyList<ChartIndexedNumberVector> seriesVectors = ReadSharedChartSeriesVectors(barPlot, workbook, plotVisibleOnly);
            if (CountRenderableSeries(seriesVectors) == 0)
            {
                return plotBox;
            }

            ChartValueExtents valueExtents = resolvedValueScale?.Extents ?? ReadCartesianChartValueAxisExtents(
                sceneChart,
                valueAxis.SceneAxis,
                GetSharedBarChartValueExtents(barPlot, grouping, workbook, plotVisibleOnly),
                percentStacked, ResolveBarValueAxisHeadroom(barOptions.BarDirection == PptxSceneChartBarDirection.Bar, percentStacked), PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            ChartAxisUnits axisUnits = resolvedValueScale?.Units ?? ResolvePercentStackedAxisUnits(ReadSceneChartValueAxisUnits(valueAxis.SceneAxis), percentStacked);
            string? defaultNumberFormat = percentStacked ? "0%" : null;
            // Office single-value-axis origins decompose to frame + 6.5pt indent + tick width
            // + 0.92 * tick font size (column-stacked/clustered, dashboard, composite ports),
            // while multi-plot dual-axis charts pack their strips tighter (compact probe:
            // Office keeps a 17.6pt strip where the measured rule wants 24.1pt), so the
            // measured rule applies to single-plot stacked charts while multi-plot charts
            // keep the shared estimator.
            IReadOnlyList<PptxSceneChartPlot> stackedBarPlots = ReadSceneChartPlots(sceneChart, PptxSceneChartPlotKind.Bar);
            double requiredReserve;
            if (UseMeasuredStackedValueAxisReserve(stackedBarPlots.Count))
            {
                ChartTextStyle stackedTickStyle = ReadSceneChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                var stackedTextMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
                double stackedMaxLabelWidth = 0d;
                foreach (double tickValue in GetChartAxisTickValuesCore(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceVerticalValueTickTargetCount, ReadCartesianChartAxisTickPolicy(sceneChart, valueAxis.SceneAxis, valueExtents, percentStacked), context: context)
                    .Concat(new[] { valueExtents.Min, valueExtents.Max }))
                {
                    string stackedLabel = FormatSceneChartAxisLabel(tickValue, valueAxis.SceneAxis, valueAxis.Date1904, defaultNumberFormat);
                    if (!string.IsNullOrWhiteSpace(stackedLabel))
                    {
                        stackedMaxLabelWidth = Math.Max(stackedMaxLabelWidth, (await stackedTextMeasurer.MeasureAsync(stackedLabel, stackedTickStyle).ConfigureAwait(false)));
                    }
                }

                if (stackedMaxLabelWidth <= 0d)
                {
                    return plotBox;
                }

                requiredReserve = ComputeCartesianAxisLabelSideReserve(stackedMaxLabelWidth, stackedTickStyle.FontSize);
            }
            else
            {
                requiredReserve = (await EstimateVerticalValueAxisLabelStripWidthAsync(theme, sceneChart, chartXml, valueAxis.SceneAxis, valueExtents, axisUnits, defaultNumberFormat, fontResolver, context: context).ConfigureAwait(false));
            }
            double leftReserve = plotBox.X - frame.X;
            double rightReserve = frame.X + frame.Width - plotBox.X - plotBox.Width;
            bool labelsRight = ResolveSceneValueAxisLabelsRightSide(sceneChart, valueAxis.SceneAxis,
                ResolveSceneValueAxisRightSide(sceneChart, valueAxis.SceneAxis, defaultRightSide: false));
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

        async ValueTask<ChartPlotBox> AdjustBarChartPlotBoxForVisibleValueAxesAsync(ChartPlotBox plotBox)
        {
            if (horizontalBars)
            {
                return plotBox;
            }

            if (visibleValueAxes.Count < 2)
            {
                return plotBox;
            }

            double leftReserve = plotBox.X - frame.X;
            double rightReserve = frame.X + frame.Width - plotBox.X - plotBox.Width;
            double rightStripWidth = 0d;
            double rightSingleReserve = 0d;
            int rightStripCount = 0;
            var leftStrips = new List<ChartValueAxisStripMeasure>();
            foreach (var (valueAxis, owner) in visibleValueAxes)
            {
                bool percentStacked = IsPercentStackedChartGrouping(owner.GroupingKind);
                ChartValueExtents data;
                if (owner.PlotKind == PptxSceneChartPlotKind.Bar)
                    data = GetSharedBarChartValueExtents(owner, owner.GroupingKind, workbook, plotVisibleOnly);
                else
                {
                    var line = ReadSceneChartLineOptions(sceneChart, owner, PptxSceneChartGrouping.Standard);
                    data = GetSharedLineChartValueExtents(owner, line.Stacked, line.PercentStacked, workbook, plotVisibleOnly);
                }
                ChartValueExtents extents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, data, percentStacked, useNearMaximumHeadroom: !percentStacked,
                    nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartAxisUnits units = ResolvePercentStackedAxisUnits(
                    ReadSceneChartValueAxisUnits(valueAxis.SceneAxis), percentStacked);
                ChartValueAxisStripMeasure strip = (await MeasureVerticalValueAxisLabelStripAsync(
                    theme,
                    sceneChart,
                    chartXml,
                    valueAxis.SceneAxis,
                    extents,
                    units,
                    defaultNumberFormat: percentStacked ? "0%" : null,
                    fontResolver: fontResolver, context: context).ConfigureAwait(false));
                double stripWidth = strip.StripWidth;
                bool labelsRight = ResolveSceneValueAxisLabelsRightSide(sceneChart,
                    valueAxis.SceneAxis,
                    ResolveSceneValueAxisRightSide(sceneChart, valueAxis.SceneAxis, defaultRightSide: false));
                if (labelsRight)
                {
                    rightStripWidth = Math.Max(rightStripWidth, stripWidth);
                    rightSingleReserve = Math.Max(rightSingleReserve,
                        ComputeCartesianAxisLabelSideReserve(strip.MaxLabelWidth, strip.FontSize));
                    rightStripCount++;
                }
                else
                {
                    leftStrips.Add(strip);
                }
            }

            // Same-side axis pairs chain their strips (Office lays dual left axes side by side)
            // instead of maxing one; right-side pairs keep the legacy factor (unobserved).
            double requiredLeftReserve = Math.Max(leftReserve, leftStrips.Count switch
            {
                0 => 0d,
                1 => ComputeCartesianAxisLabelSideReserve(leftStrips[0].MaxLabelWidth, leftStrips[0].FontSize),
                _ => ComputeChainedLeftValueAxisReserve(leftStrips)
            });
            double requiredRightReserve = Math.Max(rightReserve, rightStripCount == 1
                ? rightSingleReserve
                : rightStripWidth * PptxChartMetricRules.BarMultiValueAxisSecondaryStripFactor);
            double x = frame.X + requiredLeftReserve;
            double right = frame.X + frame.Width - requiredRightReserve;
            double width = Math.Max(1d, right - x);
            return new ChartPlotBox(x, plotBox.Y, width, plotBox.Height);
        }

        bool HasInsideValueAxisCrossing()
        {
            IReadOnlyList<ChartIndexedNumberVector> seriesVectors = ReadSharedChartSeriesVectors(barPlot, workbook, plotVisibleOnly);
            if (CountRenderableSeries(seriesVectors) == 0)
            {
                return false;
            }

            PptxSceneChartGrouping grouping = barOptions.Grouping;
            ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
            ChartValueExtents valueExtents = resolvedValueScale?.Extents ?? ReadPercentStackedAwareValueAxisExtents(
                valueAxis.SceneAxis,
                GetSharedBarChartValueExtents(barPlot, grouping, workbook, plotVisibleOnly),
                IsPercentStackedChartGrouping(grouping), false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            ChartAxisSource category = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
            double? crossing = ReadSceneValueAxisCrossingValue(category.SceneAxis, valueExtents);
            return crossing > valueExtents.Min + PptxChartMetricRules.AxisValueEpsilon &&
                crossing < valueExtents.Max - PptxChartMetricRules.AxisValueEpsilon;
        }
    }

    private static async ValueTask<ChartPlotLayout> ResolveBarManualPlotLayoutTargetAsync(
        PptxTheme theme,
        XDocument chartXml,
        PptxSceneChart sceneChart,
        PptxSceneChartPlot barPlot,
        XElement barChart,
        ChartPlotLayout layout,
        bool horizontalBars,
        ChartFrameBox frame,
        ChartWorkbookData? workbook,
        bool plotVisibleOnly,
        PresentationFontResolver? fontResolver)
    {
        if (!horizontalBars || layout.ManualLayoutTargetKind != PptxSceneChartManualLayoutTarget.Outer)
        {
            return layout;
        }

        ChartLayoutBox centeredArea = ResolveCenteredOuterAreaBox(layout.PlotAreaBox);
        ChartPlotBox plotBox = (await DeriveHorizontalBarOuterPlotBoxAsync(centeredArea).ConfigureAwait(false));
        return new ChartPlotLayout(centeredArea, plotBox, layout.ManualLayoutTargetKind);

        // Office centers outer manual areas: six-probe calibration (x10/x18/x24/x30/w50
        // plus an x14 prediction verified within 0.1pt) shows area center = frame center
        // plus x-factor times frame width, clamped to fit the frame. Left-anchored factor
        // math (shared with the inner target) misses x18 by 7.9pt and diverges without
        // bound on larger x, while centering fits all six exactly.
        ChartLayoutBox ResolveCenteredOuterAreaBox(ChartLayoutBox areaBox)
        {
            PptxSceneChartManualLayout manual = sceneChart.PlotAreaLayout;
            if (manual.XModeKind != PptxSceneChartManualLayoutMode.Factor || manual.X is not { } xFactor)
            {
                return areaBox;
            }

            double centerX = frame.X + frame.Width / 2d + xFactor * frame.Width;
            double minCenterX = frame.X + areaBox.Width / 2d;
            double maxCenterX = frame.X + frame.Width - areaBox.Width / 2d;
            double clampedCenterX = Math.Clamp(centerX, Math.Min(minCenterX, maxCenterX), Math.Max(minCenterX, maxCenterX));
            return new ChartLayoutBox(clampedCenterX - areaBox.Width / 2d, areaBox.Y, areaBox.Width, areaBox.Height);
        }

        // Outer carve: the manual area minus a measured label strip (widest label plus
        // the font-relative gap plus the area pad) on the left and the constant right
        // reserve. The ratio carve shared with the inner derivation overshoots by a
        // constant 10.5pt here; the measured carve lands within 0.1pt on all six probes.
        async ValueTask<ChartPlotBox> DeriveHorizontalBarOuterPlotBoxAsync(ChartLayoutBox plotAreaBox)
        {
            ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
            double leftReserve = 0d;
            if (IsSceneChartAxisLabelVisible(categoryAxis.SceneAxis))
            {
                ChartTextStyle tickStyle = ReadSceneChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                double labelOffsetScale = ResolveSceneCategoryAxisLabelOffsetScale(categoryAxis.SceneAxis);
                int tickLabelSkip = ResolveSceneCategoryAxisTickLabelSkip(categoryAxis.SceneAxis);
                var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
                double maxCategoryWidth = 0d;
                int labelIndex = 0;
                foreach (ChartIndexedTextPoint? label in ReadSharedCategoryLabels(barPlot, workbook, plotVisibleOnly))
                {
                    // Mirror the emission skip logic in RenderChartCategoryLabels so the
                    // reserve covers rendered labels only.
                    string? reserveLabel = label?.Text;
                    if (labelIndex % tickLabelSkip == 0 && !string.IsNullOrWhiteSpace(reserveLabel))
                    {
                        maxCategoryWidth = Math.Max(maxCategoryWidth, (await textMeasurer.MeasureAsync(reserveLabel, tickStyle).ConfigureAwait(false)));
                    }

                    labelIndex++;
                }

                leftReserve = maxCategoryWidth +
                    PptxChartMetricRules.HorizontalBarCategoryLabelPlotGapFactor * tickStyle.FontSize * labelOffsetScale +
                    PptxChartMetricRules.HorizontalBarOuterAreaLabelPad;
            }

            double plotX = plotAreaBox.X + leftReserve;
            double plotRight = plotAreaBox.X + plotAreaBox.Width - PptxChartMetricRules.HorizontalBarOuterPlotRightReserve;
            ChartPlotBox derived = DeriveHorizontalBarInnerPlotBox(plotAreaBox);
            return new ChartPlotBox(plotX, derived.Y, Math.Max(1d, plotRight - plotX), derived.Height);
        }
        ChartPlotBox DeriveHorizontalBarInnerPlotBox(ChartLayoutBox plotAreaBox)
        {
            double leftReserve = 0d;
            ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
            if (IsSceneChartAxisLabelVisible(categoryAxis.SceneAxis))
            {
                double labelOffsetScale = ResolveSceneCategoryAxisLabelOffsetScale(categoryAxis.SceneAxis);
                double outsideFactor =
                    PptxChartMetricRules.CategoryAxisHorizontalLeftOffsetRatio * labelOffsetScale +
                    PptxChartMetricRules.CategoryAxisHorizontalWidthRatio;
                leftReserve = plotAreaBox.Width * outsideFactor / (1d + outsideFactor);
            }

            double rightReserve = ReadSceneChartAxisMajorTickMark(categoryAxis.SceneAxis) == PptxSceneChartAxisTickMark.Outside
                ? PptxChartMetricRules.CategoryAxisMajorTickLength + Math.Max(3d, PptxChartMetricRules.ValueAxisFallbackFontSize * 0.35d)
                : 0d;
            double topReserve = 0d;
            ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
            if (IsSceneChartAxisLabelVisible(valueAxis.SceneAxis))
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
    }

    // Chained same-side value-axis reserve: Office lays dual left axes side by side, each
    // with the single-axis indent/gap idiom (6.5pt frame indent, 0.92fs plot gap) and a
    // sideGap-plus-1.46 inter-axis gap (three Office renders agree within 0.1pt). Strips
    // run outer-to-inner in plot order; beyond-two and mixed-font configs unobserved.
    private static double ComputeChainedLeftValueAxisReserve(IReadOnlyList<ChartValueAxisStripMeasure> strips)
    {
        double reserve = PptxChartMetricRules.BarValueAxisLabelFrameIndent + strips[0].MaxLabelWidth;
        for (int stripIndex = 1; stripIndex < strips.Count; stripIndex++)
        {
            reserve += strips[stripIndex].SideGap + PptxChartMetricRules.BarDualValueAxisInterAxisGap + strips[stripIndex].MaxLabelWidth;
        }

        return reserve + strips[strips.Count - 1].FontSize * PptxChartMetricRules.BarValueAxisLabelGapFactor;
    }

    // Minimum top clearance for untitled legendless columns: Office keeps the plot top
    // at least 10.9pt below the frame top (exact on 170H, 288H and 130H Office renders
    // where the frame ratio alone wants 6.3/10.7/4.8); taller frames keep the ratio.
    // Titled, legend, horizontal and label-less paths keep legacy tops (the law comes
    // from labeled charts only; the single-point confound stays open).
    private static ChartPlotBox AdjustVerticalBarPlotBoxTopFloor(ChartPlotBox plotBox, ChartFrameBox frame, bool horizontalBars, bool hasTitle, bool hasLegend, bool hasCategoryLabels)
    {
        if (horizontalBars || hasTitle || hasLegend || !hasCategoryLabels)
        {
            return plotBox;
        }

        double flooredTop = Math.Min(plotBox.Y + plotBox.Height, frame.Y + frame.Height - PptxChartMetricRules.ColumnPlotTopMarginFloor);
        return new ChartPlotBox(plotBox.X, plotBox.Y, plotBox.Width, Math.Max(1d, flooredTop - plotBox.Y));
    }

    // Right reserve for columns: Office keeps the plot AXIS right edge 11.0pt inside
    // the frame (axis truth after the clip-vs-axis finding; an exact set trips tight tick gates on green
    // ports and shoves right-axis charts into their axes, so the edge is floored, not set).
    // Horizontal bars keep the value-axis tail law; side legends, hidden-label and
    // label-less charts keep legacy edges.
    private static ChartPlotBox AdjustVerticalBarPlotBoxRightReserve(ChartPlotBox plotBox, ChartFrameBox frame, bool horizontalBars, ChartLegendLayout legend, bool hasCategoryLabels)
    {
        if (horizontalBars || !hasCategoryLabels)
        {
            return plotBox;
        }

        if (legend.Visible && !legend.Overlay &&
            (legend.PositionKind == PptxSceneChartLegendPosition.Left || legend.PositionKind == PptxSceneChartLegendPosition.Right))
        {
            return plotBox;
        }

        double flooredRight = frame.X + frame.Width - PptxChartMetricRules.ColumnPlotRightReserve;
        double right = Math.Min(plotBox.X + plotBox.Width, flooredRight);
        return new ChartPlotBox(plotBox.X, plotBox.Y, Math.Max(1d, right - plotBox.X), plotBox.Height);
    }
    // R15: existence probe over the sparse points; the dense array is only built
    // when labels actually render.
    private static bool HasRenderableCategoryLabels(PptxSceneChartPlot barPlot, XElement barChart, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        return ReadSceneCategoryLabelVector(barPlot, workbook, plotVisibleOnly).HasAnyVisibleText();
    }
    private static ChartPlotBox AdjustStackedColumnBottomLegendPlotBox(
        ChartPlotBox plotBox,
        ChartFrameBox frame,
        bool horizontalBars,
        PptxSceneChartGrouping grouping,
        bool hasTitle,
        ChartLegendLayout legend)
    {
        if (horizontalBars ||
            hasTitle ||
            !legend.Visible ||
            legend.Overlay ||
            legend.PositionKind != PptxSceneChartLegendPosition.Bottom ||
            !IsStackedChartGrouping(grouping))
        {
            return plotBox;
        }

        // No left pad: the measured value-axis reserve already lands the left edge within
        // 0.1pt (the 1.8pt pad predates it and overshoots); the right pad stays measured.
        double left = Math.Min(frame.X + frame.Width, plotBox.X);
        double right = Math.Max(
            left + 1d,
            plotBox.X + plotBox.Width - PptxChartMetricRules.StackedColumnBottomLegendPlotBoxRightPadding);
        return new ChartPlotBox(left, plotBox.Y, right - left, plotBox.Height);
    }

    // Bottom reserve for untitled bottom-legend columns as a calibrated additive plane
    // over legend and category-tick font sizes (see call-site probe table).
    private static double ComputeNoTitleBottomLegendReserve(double legendFontSize, double tickFontSize)
    {
        return legendFontSize * PptxChartMetricRules.BarNoTitleBottomLegendReserveLegendFactor +
            tickFontSize * PptxChartMetricRules.BarNoTitleBottomLegendReserveTickFactor +
            PptxChartMetricRules.BarNoTitleBottomLegendReserveBase;
    }

    private static ChartPlotBox AdjustNoTitleBottomLegendPlotBoxForMeasuredContent(
        ChartPlotBox plotBox,
        ChartFrameBox frame,
        PptxTheme theme,
        PptxColorMap colorMap,
        PptxSceneChart sceneChart,
        XDocument chartXml,
        PptxSceneChartPlot barPlot,
        XElement barChart,
        bool horizontalBars,
        bool usesMeasuredAxisInsets)
    {
        // Office keeps a content-sized bottom margin for untitled bottom-legend columns.
        // Four same-frame probes with varied fonts pin an additive plane (all within 0.06):
        // 55.02 over 18pt legend plus 8pt cats, 47.77 over 12pt plus 8pt, 66.10 over 18pt
        // plus 14pt, 58.86 over 12pt plus 14pt. The single-knob forms are dead (the error
        // flips sign across probes), and bottom-minus-cat-baseline scales cleanly at 1.566
        // catFs inside the total without its own usable law, so the plane ships as such.
        ChartTextStyle legendStyle = ReadSceneChartLegendTextStyle(theme, sceneChart);
        bool horizontalStrip = horizontalBars && usesMeasuredAxisInsets;
        ChartAxisSource tickAxis = horizontalStrip
            ? ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault()
            : ReadChartCategoryAxisForPlot(sceneChart, barPlot);
        ChartTextStyle tickStyle = ReadSceneChartTextStyle(theme, sceneChart, tickAxis.SceneAxis,
            fallbackFontSize: horizontalStrip ? PptxChartMetricRules.ValueAxisFallbackFontSize : PptxChartMetricRules.CategoryAxisFallbackFontSize,
            chartStyleRole: horizontalStrip ? "valueAxis" : "categoryAxis");
        double bottomReserve = ComputeNoTitleBottomLegendReserve(legendStyle.FontSize, tickStyle.FontSize);
        if (horizontalStrip && !IsSceneChartAxisLabelVisible(tickAxis.SceneAxis))
        {
            bottomReserve -= ComputeCartesianLabelStripBottomReserve(tickStyle.FontSize) - PptxChartMetricRules.HorizontalBarPlotTopReserve;
        }
        double presetTop = frame.Y + frame.Height * (PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxYRatio + PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxHeightRatio);
        double y = frame.Y + bottomReserve;
        double top = presetTop;
        if (!horizontalBars)
        {
            top = System.Math.Max(top, ResolveNoTitleBottomLegendAnchoredTop(frame));
        }

        return new ChartPlotBox(plotBox.X, y, plotBox.Width, Math.Max(1d, top - y));
    }

    // Frame-top-anchored plot top for untitled bottom-legend columns: Office clips
    // agree bit-identically base+tall, so the top answers only to the frame top edge.
    private static double ResolveNoTitleBottomLegendAnchoredTop(ChartFrameBox frame)
    {
        return frame.Y + frame.Height - PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxTopInset;
    }

    // Regime gate for stacked value-axis reserves: the measured indent-plus-font-gap rule
    // is calibrated on single-plot single-axis charts only; multi-plot charts (dual-axis
    // compact/overlay probes in corpus) keep the shared estimator.
    private static bool UseMeasuredStackedValueAxisReserve(int barPlotCount)
    {
        return barPlotCount == 1;
    }
    // Side strips share label width, frame indent and a font-relative gap:
    // vertical value labels and horizontal category labels have the same owner.
    // Manual and side-legend defaults retain their placement fallback above.
    private static double ComputeCartesianAxisLabelSideReserve(double maxLabelWidth, double tickFontSize)
    {
        return maxLabelWidth +
            PptxChartMetricRules.BarValueAxisLabelFrameIndent +
            tickFontSize * PptxChartMetricRules.BarValueAxisLabelGapFactor;
    }

    // Bottom reserve for the label strip below a bar plot: frame margin plus descent
    // plus our emitted label gap (Office baselines frame plus 7.0 plus descent on
    // vertical category and horizontal value labels alike).
    private static double ComputeCartesianLabelStripBottomReserve(double labelFontSize)
    {
        return PptxChartMetricRules.BarCategoryBottomMargin +
            labelFontSize * PptxChartMetricRules.BarCategoryDescentFactor +
            labelFontSize * PptxChartMetricRules.AxisLabelHeightFactor *
            PptxChartMetricRules.CategoryAxisVerticalTopOffsetFactor;
    }

    // Right reserve for horizontal-bar value labels: the edge tick centers on the plot
    // edge, so half the widest label plus the Office tail fits it (bar margins decompose
    // to half-width plus 11.0pt on stacked and clustered references).
    private static double ResolveMeasuredRightReserve(double presetRightReserve, double maxValueLabelWidth)
    {
        return Math.Max(presetRightReserve, maxValueLabelWidth / 2d + PptxChartMetricRules.HorizontalBarValueAxisRightPadding);
    }

    private readonly record struct ChartValueAxisStripMeasure(double MaxLabelWidth, double FontSize, double StripWidth, double SideGap);

    private static async ValueTask<double> EstimateVerticalValueAxisLabelStripWidthAsync(PptxTheme theme, PptxSceneChart sceneChart, XDocument chartXml, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits units, string? defaultNumberFormat, PresentationFontResolver? fontResolver, PptxRenderContext? context = null)
    {
        return (await MeasureVerticalValueAxisLabelStripAsync(theme, sceneChart, chartXml, sceneAxis, extents, units, defaultNumberFormat, fontResolver, context: context).ConfigureAwait(false)).StripWidth;
    }

    private static async ValueTask<ChartValueAxisStripMeasure> MeasureVerticalValueAxisLabelStripAsync(PptxTheme theme, PptxSceneChart sceneChart, XDocument chartXml, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits units, string? defaultNumberFormat, PresentationFontResolver? fontResolver, PptxRenderContext? context = null)
    {
        ChartTextStyle style = ReadSceneChartTextStyle(theme, sceneChart, sceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
        double fontSize = style.FontSize;
        var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
        bool percentStacked = sceneChart is { Plots.Count: 1 } && IsPercentStackedChartGrouping(sceneChart.Plots[0].GroupingKind);
        var captions = await MeasureChartValueAxisCaptionsAsync(textMeasurer, style, new(sceneAxis, sceneChart.Options.Date1904 == true),
            extents, units, PptxChartMetricRules.AxisNiceVerticalValueTickTargetCount,
            ReadCartesianChartAxisTickPolicy(sceneChart, sceneAxis, extents, percentStacked),
            defaultNumberFormat, reserveEndpoints: true, context).ConfigureAwait(false);
        return new ChartValueAxisStripMeasure(captions.MaximumWidth, fontSize,
            captions.SlotWidth + captions.SideGap, captions.SideGap);
    }

    // Near-maximum headroom (the shared 0.96 rule) applies to vertical-bar value axes:
    // Office ceilings dataMax 68 to 80 on the ladder column probe. Horizontal bars keep
    // the legacy frozen behavior their tuned right-margin laws were calibrated against,
    // and percent stacks keep theirs (their max is pinned downstream anyway).
    private static bool ResolveBarValueAxisHeadroom(bool horizontalBars, bool percentStacked)
    {
        return !horizontalBars && !percentStacked;
    }

    private static ChartValueExtents GetBarChartValueExtents(IReadOnlyList<ChartIndexedNumberVector> series, PptxSceneChartGrouping grouping)
    {
        IReadOnlyList<IReadOnlyList<double?>> denseSeries = DensifyChartValueSeries(series);
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

    private static void DrawHorizontalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, ChartValueExtents extents, double? explicitUnit, double? crossingValue, bool reversed, bool major, ChartSeriesStroke? gridlineStroke, PptxRenderContext? context = null)
    {
        DrawHorizontalChartGridlinePaths(graphics, plotX, plotY, plotWidth, plotHeight, extents, reversed, major, gridlineStroke,
            () => GetChartGridlineValues(extents, explicitUnit, crossingValue, PptxChartMetricRules.AxisNiceTickTargetCount, context: context));
    }

    private static void DrawHorizontalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, ChartValueExtents extents, ChartAxisUnits units, double? crossingValue, bool reversed, bool major, ChartSeriesStroke? gridlineStroke, bool hasMajorGridlineElement, ChartAxisTickPolicy tickPolicy = ChartAxisTickPolicy.ZeroAlignedWithEndpoints, PptxRenderContext? context = null)
    {
        DrawHorizontalChartGridlinePaths(graphics, plotX, plotY, plotWidth, plotHeight, extents, reversed, major, gridlineStroke,
            () => GetChartGridlineValuesCore(extents, units, major, crossingValue,
                tickPolicy == ChartAxisTickPolicy.SourceMinimumIntervals
                    ? PptxChartMetricRules.AxisNiceVerticalValueTickTargetCount
                    : PptxChartMetricRules.AxisNiceTickTargetCount,
                hasMajorGridlineElement, tickPolicy, context: context));
    }

    private static void DrawHorizontalChartGridlinePaths(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, ChartValueExtents extents, bool reversed, bool major, ChartSeriesStroke? gridlineStroke, Func<IReadOnlyList<double>> readValues)
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

        // Office defaults gridlines to round joins (caps stay butt).
        SetChartStroke(graphics, stroke with { Join = stroke.Join ?? 1 });
        bool hasPath = false;
        foreach (double value in readValues())
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

    private static void DrawVerticalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, ChartValueExtents extents, ChartValueAxisRenderOptions options, bool major, double autoTickTargetCount, PptxRenderContext? context = null)
    {
        ChartSeriesStroke stroke = (major ? options.GridlineStyle.Major : options.GridlineStyle.Minor) ?? DefaultChartGridlineStroke(major);
        if (stroke.Alpha <= 0.001d)
        {
            return;
        }

        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        // Office defaults gridlines to round joins (caps stay butt).
        SetChartStroke(graphics, stroke with { Join = stroke.Join ?? 1 });
        bool hasPath = false;
        IReadOnlyList<double> values = options.TickPolicy == ChartAxisTickPolicy.SourceMinimumIntervals
            ? GetChartGridlineValuesCore(extents, options.Units, major, options.CrossingValue, autoTickTargetCount, options.HasMajorGridlineElement, options.TickPolicy, context: context)
            : GetChartGridlineValues(extents, major ? options.Units.MajorUnit : options.Units.MinorUnit, options.CrossingValue,
                major ? autoTickTargetCount : PptxChartMetricRules.AxisNiceTickTargetCount, context: context);
        foreach (double value in values)
        {
            double x = ChartValueToPlotCoordinate(extents, value, plotX, plotWidth, options.Reversed);
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

    // RV04: effective-style-18/26 bars paint a vertical base-relative gradient resolved
    // from the series base: every painted series rect (clustered bar, stack segment,
    // horizontal bar) carries bar-relative Coords spanning twice its height from its
    // bottom edge, so the visible bottom half runs light-middle-dark. The six sampled
    // theme bases below carry byte-exact Office triples; any other opaque base resolves
    // through DeriveStyleBarGradientStops (mid exact, dark within 1 level, light within
    // 2 over the 37-face PowerPoint 16.0 corpus in artifacts/rv04dark). Other styles,
    // transparent or patterned fills keep the flat path.
    private static bool TryGetStyleBarGradientStops(int? chartStyleId, ChartSeriesFill fill, out IReadOnlyList<PdfShadingStop> stops)
    {
        stops = Array.Empty<PdfShadingStop>();
        if ((chartStyleId != 18 && chartStyleId != 26 && chartStyleId != 118 && chartStyleId != 126) ||
            fill.PatternPreset is not null ||
            fill.Alpha < 1d)
        {
            return false;
        }

        if (fill.Color.Equals(new RgbColor(21, 96, 130)))
        {
            stops = StyleBarGradientStops(new RgbColor(73, 116, 145), new RgbColor(16, 98, 135), new RgbColor(8, 88, 124));
            return true;
        }

        if (fill.Color.Equals(new RgbColor(233, 113, 50)))
        {
            stops = StyleBarGradientStops(new RgbColor(237, 130, 86), new RgbColor(242, 110, 41), new RgbColor(225, 94, 25));
            return true;
        }

        if (fill.Color.Equals(new RgbColor(192, 0, 0)))
        {
            stops = StyleBarGradientStops(new RgbColor(203, 69, 69), new RgbColor(202, 0, 0), new RgbColor(187, 0, 0));
            return true;
        }

        if (fill.Color.Equals(new RgbColor(112, 48, 160)))
        {
            stops = StyleBarGradientStops(new RgbColor(129, 83, 172), new RgbColor(113, 42, 166), new RgbColor(100, 31, 151));
            return true;
        }

        if (fill.Color.Equals(new RgbColor(14, 40, 65)))
        {
            stops = StyleBarGradientStops(new RgbColor(72, 80, 94), new RgbColor(11, 40, 68), new RgbColor(7, 34, 61));
            return true;
        }

        if (fill.Color.Equals(new RgbColor(255, 192, 0)))
        {
            stops = StyleBarGradientStops(new RgbColor(255, 199, 70), new RgbColor(255, 198, 0), new RgbColor(229, 182, 0));
            return true;
        }

        (RgbColor derivedDark, RgbColor derivedMiddle, RgbColor derivedLight) = DeriveStyleBarGradientStops(fill.Color);
        stops = StyleBarGradientStops(derivedLight, derivedMiddle, derivedDark);
        return true;
    }

    // RV04: derived style-18/26 gradient stops for bases outside the sampled table.
    // Corpus: 37 PowerPoint 16.0 (base, dark, mid, light) triples extracted from Office
    // axial-shading stitching functions (artifacts/rv04dark). Mid is an exact sRGB
    // saturation stretch about the face lightness; dark tracks a 0.763 linear-light
    // factor plus a 0.184 saturation coupling, clipped to 229; light tracks a 0.98
    // linear-light factor plus a 0.06 lift with a 0.008 coupling. Measured agreement:
    // mid byte-exact on every face, dark within 1 level, light within 2.
    internal static (RgbColor Dark, RgbColor Middle, RgbColor Light) DeriveStyleBarGradientStops(RgbColor baseColor)
    {
        double lightness = (System.Math.Max(baseColor.Red, System.Math.Max(baseColor.Green, baseColor.Blue)) +
            System.Math.Min(baseColor.Red, System.Math.Min(baseColor.Green, baseColor.Blue))) / 2d;
        return (
            DeriveStyleDarkStop(baseColor, lightness),
            DeriveStyleMiddleStop(baseColor, lightness),
            DeriveStyleLightStop(baseColor, lightness));
    }

    private static RgbColor DeriveStyleMiddleStop(RgbColor baseColor, double lightness)
    {
        return new RgbColor(
            StyleGradientChannel(lightness + 1.1d * (baseColor.Red - lightness), 0, 255),
            StyleGradientChannel(lightness + 1.1d * (baseColor.Green - lightness), 0, 255),
            StyleGradientChannel(lightness + 1.1d * (baseColor.Blue - lightness), 0, 255));
    }

    private static RgbColor DeriveStyleDarkStop(RgbColor baseColor, double lightness)
    {
        return new RgbColor(
            DeriveStyleDarkChannel(baseColor.Red, lightness),
            DeriveStyleDarkChannel(baseColor.Green, lightness),
            DeriveStyleDarkChannel(baseColor.Blue, lightness));
    }

    private static byte DeriveStyleDarkChannel(byte channel, double lightness)
    {
        double staged = LinearLightColor.ToSrgb(0.763d * LinearLightColor.ToLinear(channel)) * 255d +
            0.184d * (channel - lightness);
        return StyleGradientChannel(staged, 0, 229);
    }

    private static RgbColor DeriveStyleLightStop(RgbColor baseColor, double lightness)
    {
        return new RgbColor(
            DeriveStyleLightChannel(baseColor.Red, lightness),
            DeriveStyleLightChannel(baseColor.Green, lightness),
            DeriveStyleLightChannel(baseColor.Blue, lightness));
    }

    private static byte DeriveStyleLightChannel(byte channel, double lightness)
    {
        double staged = LinearLightColor.ToSrgb(0.98d * LinearLightColor.ToLinear(channel) + 0.06d) * 255d +
            0.008d * (channel - lightness);
        return StyleGradientChannel(staged, 0, 255);
    }

    private static byte StyleGradientChannel(double staged, int min, int max)
    {
        return (byte)System.Math.Clamp((int)System.Math.Round(staged, System.MidpointRounding.AwayFromZero), min, max);
    }

    // RV04: five-knot forward stitch matching the visible bottom half of the Office
    // symmetric profile (dark, middle, light, middle, dark at quarter offsets).
    private static IReadOnlyList<PdfShadingStop> StyleBarGradientStops(RgbColor light, RgbColor middle, RgbColor dark)
    {
        return
        [
            new PdfShadingStop(0d, dark.Red, dark.Green, dark.Blue),
            new PdfShadingStop(0.25d, middle.Red, middle.Green, middle.Blue),
            new PdfShadingStop(0.5d, light.Red, light.Green, light.Blue),
            new PdfShadingStop(0.75d, middle.Red, middle.Green, middle.Blue),
            new PdfShadingStop(1d, dark.Red, dark.Green, dark.Blue),
        ];
    }

    private static bool PaintBarGradientRect(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, double barX, double barY, double barWidth, double barHeight, IReadOnlyList<PdfShadingStop> stops)
    {
        if (barWidth <= 0d || barHeight <= 0d)
        {
            return false;
        }

        RenderInChartPlotAreaClip(graphics, plotBox, () =>
        {
            graphics.SaveState();
            graphics.ClipRectangle(barX, barY, barWidth, barHeight);
            graphics.PaintAxialShading(barX, barY + 2d * barHeight, barX, barY, stops);
            graphics.RestoreState();
        });
        return true;
    }

    // RV04: effective-style-10 clustered columns stroke white 1pt round-join bar
    // outlines sampled from Office bar references; other styles keep current strokes.
    // Stacked and horizontal bars keep the current path.
    private static void PaintStyleColumnOutline(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, int? chartStyleId, bool hasExplicitFill, ChartSeriesFill fill, double barX, double barY, double barWidth, double barHeight)
    {
        if ((chartStyleId != 10 && chartStyleId != 110) ||
            hasExplicitFill ||
            fill.PatternPreset is not null ||
            fill.Alpha < 1d ||
            barWidth <= 0d ||
            barHeight <= 0d)
        {
            return;
        }

        RenderInChartPlotAreaClip(graphics, plotBox, () =>
        {
            graphics.SaveState();
            graphics.SetLineWidth(1d);
            graphics.SetStrokeRgb(255, 255, 255);
            graphics.SetLineJoin(1);
            graphics.StrokeRectangle(barX, barY, barWidth, barHeight);
            graphics.RestoreState();
        });
    }

    private static ChartSeriesStroke DefaultChartGridlineStroke(bool major)
    {
        return major
            ? new ChartSeriesStroke(new RgbColor(0, 0, 0), 1d, 0.75d)
            : new ChartSeriesStroke(new RgbColor(235, 235, 235), 1d, 0.25d);
    }

    private static void FillChartRectangle(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartSeriesFill fill)
    {
        if (fill.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(fill.Alpha, 1d);
        }

        if (fill.PatternPreset is not null && !ChartPatternRequiresBackgroundPaint(fill.PatternPreset))
        {
            StrokeChartPatternFill(graphics, x, y, width, height, fill);
        }
        else
        {
            RgbColor fillColor = fill.BackgroundColor ?? fill.Color;
            graphics.SetFillRgb(fillColor.Red, fillColor.Green, fillColor.Blue);
            graphics.FillRectangle(x, y, width, height);
            if (fill.PatternPreset is not null)
            {
                StrokeChartPatternFill(graphics, x, y, width, height, fill);
            }
        }

        if (fill.Alpha < 1d)
        {
            graphics.RestoreState();
        }

        bool ChartPatternRequiresBackgroundPaint(string patternPreset)
        {
            return TryReadPercentageChartPattern(patternPreset, out _);
        }
    }

    private static void StrokeChartPatternFill(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ChartSeriesFill fill)
    {
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        string patternPreset = fill.PatternPreset ?? "pct50";
        if (TryReadPercentageChartPattern(patternPreset, out int densityPercent))
        {
            if (densityPercent == 70)
            {
                RgbColor weaveBackground = fill.BackgroundColor ?? new RgbColor(255, 255, 255);
                var weave = PdfTilingPattern.OfficeBitmapWeaveBlocks(fill.Color.Red, fill.Color.Green, fill.Color.Blue, weaveBackground.Red, weaveBackground.Green, weaveBackground.Blue);
                graphics.FillRectangleWithTilingPattern(x, y, width, height, weave);
                return;
            }

            graphics.SaveState();
            graphics.ClipRectangle(x, y, width, height);
            graphics.SetStrokeRgb(fill.Color.Red, fill.Color.Green, fill.Color.Blue);
            FillChartDotPattern(fill.Color, densityPercent);
            graphics.RestoreState();
            return;
        }

        RgbColor backgroundColor = fill.BackgroundColor ?? new RgbColor(255, 255, 255);
        var pattern = PdfTilingPattern.OfficeBitmapDiagonalLines(
            patternPreset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase),
            fill.Color.Red,
            fill.Color.Green,
            fill.Color.Blue,
            backgroundColor.Red,
            backgroundColor.Green,
            backgroundColor.Blue);
        graphics.FillRectangleWithTilingPattern(x, y, width, height, pattern);

        void FillChartDotPattern(RgbColor color, int densityPercent)
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
    }

    private static void FillChartRectangleInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, double x, double y, double width, double height, ChartSeriesFill fill)
    {
        RenderInChartPlotAreaClip(graphics, plotBox, () => FillChartRectangle(graphics, x, y, width, height, fill));
    }

    private static void FillChartRectanglesAsCompoundPathInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, IReadOnlyList<ChartRectangle> rectangles, ChartSeriesFill fill)
    {
        RenderInChartPlotAreaClip(graphics, plotBox, () => FillChartRectanglesAsCompoundPath());

        void FillChartRectanglesAsCompoundPath()
        {
            if (rectangles.Count == 0)
            {
                return;
            }

            if (fill.PatternPreset is not null)
            {
                foreach (ChartRectangle rectangle in rectangles)
                {
                    FillChartRectangle(graphics, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height, fill);
                }

                return;
            }

            if (fill.Alpha < 1d)
            {
                graphics.SaveState();
                graphics.SetAlpha(fill.Alpha, 1d);
            }

            RgbColor fillColor = fill.BackgroundColor ?? fill.Color;
            graphics.SetFillRgb(fillColor.Red, fillColor.Green, fillColor.Blue);
            foreach (ChartRectangle rectangle in rectangles)
            {
                AppendChartRectanglePath(rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
            }

            graphics.FillCurrentPath();
            if (fill.Alpha < 1d)
            {
                graphics.RestoreState();
            }

            void AppendChartRectanglePath(double x, double y, double width, double height)
            {
                graphics.MoveTo(x, y);
                graphics.LineTo(x + width, y);
                graphics.LineTo(x + width, y + height);
                graphics.LineTo(x, y + height);
                graphics.ClosePath();
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

    private static void RenderClusteredHorizontalBars(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double?>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, double originValue, double originX, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, ChartBarPlotOptions barOptions, int? chartStyleId)
    {
        bool varyColors = barOptions.VaryColors.Value;
        double gapWidthPercent = barOptions.GapWidth;
        double overlapPercent = barOptions.Overlap;
        double categoryHeight = plotHeight / categoryCount;
        double barHeight = GetClusteredBarWidth(categoryHeight, series.Count, gapWidthPercent, overlapPercent);
        double step = GetClusteredBarStep(barHeight, overlapPercent);
        double clusterHeight = barHeight + Math.Max(0, series.Count - 1) * step;
        for (int category = 0; category < categoryCount; category++)
        {
            double categoryY = plotY + category * categoryHeight + (categoryHeight - clusterHeight) / 2d;
            for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                IReadOnlyList<double?> values = series[seriesIndex];
                if (category >= values.Count || values[category] is not { } value)
                {
                    continue;
                }

                double relativeValue = value - originValue;
                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, values.Count, relativeValue, barOptions.InvertsNegativePoint(seriesIndex, category));
                double valueX = ChartValueToPlotCoordinate(valueExtents, value, plotX, plotWidth, valueAxisReversed);
                double barX = Math.Min(originX, valueX);
                double barWidth = Math.Abs(valueX - originX);
                double barY = categoryY + seriesIndex * step;
                if (barOptions.CategoryAxisReversed)
                    barY = ChartCategoryToPlotCoordinate(category, categoryCount, plotY, plotHeight,
                        onTickMarks: false, reversed: true,
                        slotOffset: seriesIndex * step + (barHeight - clusterHeight) / 2d) - barHeight / 2d;
                bool isStyleDriven = (seriesIndex >= seriesFills.Count || seriesFills[seriesIndex] is null) && !HasExplicitChartPointFill(pointFills, seriesIndex, category);
                if (isStyleDriven && TryGetStyleBarGradientStops(chartStyleId, fill, out IReadOnlyList<PdfShadingStop> stops))
                {
                    PaintBarGradientRect(graphics, plotBox, barX, barY, barWidth, barHeight, stops);
                }
                else
                {
                    FillChartRectangleInPlotClip(graphics, plotBox, barX, barY, barWidth, barHeight, fill);
                }
                StrokeChartPointRectangleInPlotClip(graphics, plotBox, seriesIndex, category, pointStrokes, barX, barY, barWidth, barHeight, ResolveNegativeBarFallbackStroke(pointStrokes, seriesIndex, category, relativeValue, barOptions.InvertsNegativePoint(seriesIndex, category)));
            }
        }
    }

    private static void RenderStackedColumns(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double?>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, ChartBarPlotOptions barOptions, int? chartStyleId)
    {
        bool varyColors = barOptions.VaryColors.Value;
        double gapWidthPercent = barOptions.GapWidth;
        double categoryWidth = plotWidth / categoryCount;
        double barWidth = GetStackedBarWidth(categoryWidth, gapWidthPercent);
        double[] positiveValues = new double[categoryCount];
        double[] negativeValues = new double[categoryCount];
        double[] positiveTotals = GetCategoryPositiveTotals(series, categoryCount, percentStacked);
        var strokeSegments = new List<ChartStackedBarSegment>();
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            var fillRuns = new List<ChartStackedBarFillRun>();
            for (int category = 0; category < categoryCount; category++)
            {
                IReadOnlyList<double?> values = series[seriesIndex];
                if (category >= values.Count || values[category] is not { } rawValue)
                {
                    continue;
                }

                double categoryX = plotX + category * categoryWidth + (categoryWidth - barWidth) / 2d;
                double value = NormalizeStackedValue(rawValue, positiveTotals[category], percentStacked);
                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, values.Count, value, barOptions.InvertsNegativePoint(seriesIndex, category));
                double segmentStartValue;
                double segmentEndValue;
                if (value >= 0d)
                {
                    segmentStartValue = positiveValues[category];
                    positiveValues[category] += value;
                    segmentEndValue = positiveValues[category];
                }
                else
                {
                    segmentStartValue = negativeValues[category];
                    negativeValues[category] += value;
                    segmentEndValue = negativeValues[category];
                }

                double startY = ChartValueToPlotCoordinate(valueExtents, segmentStartValue, plotY, plotHeight, valueAxisReversed);
                double endY = ChartValueToPlotCoordinate(valueExtents, segmentEndValue, plotY, plotHeight, valueAxisReversed);
                double segmentY = Math.Min(startY, endY);
                double segmentHeight = Math.Abs(endY - startY);
                AddStackedBarFillRun(fillRuns, fill, new ChartRectangle(categoryX, segmentY, barWidth, segmentHeight));
                strokeSegments.Add(new ChartStackedBarSegment(seriesIndex, category, categoryX, segmentY, barWidth, segmentHeight, value));
            }

            bool seriesDriven = (seriesIndex >= seriesFills.Count || seriesFills[seriesIndex] is null) &&
                (seriesIndex >= pointFills.Count || pointFills[seriesIndex].Count == 0);
            foreach (ChartStackedBarFillRun run in fillRuns)
            {
                if (seriesDriven && TryGetStyleBarGradientStops(chartStyleId, run.Fill, out IReadOnlyList<PdfShadingStop> stops))
                {
                    foreach (ChartRectangle rectangle in run.Rectangles)
                    {
                        PaintBarGradientRect(graphics, plotBox, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height, stops);
                    }
                }
                else
                {
                    FillChartRectanglesAsCompoundPathInPlotClip(graphics, plotBox, run.Rectangles, run.Fill);
                }
            }
        }

        foreach (ChartStackedBarSegment segment in strokeSegments)
        {
            StrokeChartPointRectangleInPlotClip(
                graphics,
                plotBox,
                segment.SeriesIndex,
                segment.CategoryIndex,
                pointStrokes,
                segment.X,
                segment.Y,
                segment.Width,
                segment.Height,
                ResolveNegativeBarFallbackStroke(pointStrokes, segment.SeriesIndex, segment.CategoryIndex, segment.Value, barOptions.InvertsNegativePoint(segment.SeriesIndex, segment.CategoryIndex)));
        }
    }

    private static void AddStackedBarFillRun(List<ChartStackedBarFillRun> fillRuns, ChartSeriesFill fill, ChartRectangle rectangle)
    {
        for (int i = 0; i < fillRuns.Count; i++)
        {
            ChartStackedBarFillRun run = fillRuns[i];
            if (run.Fill == fill)
            {
                run.Rectangles.Add(rectangle);
                return;
            }
        }

        var rectangles = new List<ChartRectangle> { rectangle };
        fillRuns.Add(new ChartStackedBarFillRun(fill, rectangles));
    }

    private static void RenderStackedHorizontalBars(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, PptxTheme theme, PptxColorMap colorMap, ChartColorPalette chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<double?>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, ChartBarPlotOptions barOptions, int? chartStyleId)
    {
        bool varyColors = barOptions.VaryColors.Value;
        double gapWidthPercent = barOptions.GapWidth;
        double categoryHeight = plotHeight / categoryCount;
        double barHeight = GetStackedBarWidth(categoryHeight, gapWidthPercent);
        double[] positiveValues = new double[categoryCount];
        double[] negativeValues = new double[categoryCount];
        double[] positiveTotals = GetCategoryPositiveTotals(series, categoryCount, percentStacked);
        var strokeSegments = new List<ChartStackedBarSegment>();
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            var fillRuns = new List<ChartStackedBarFillRun>();
            for (int category = 0; category < categoryCount; category++)
            {
                IReadOnlyList<double?> values = series[seriesIndex];
                if (category >= values.Count || values[category] is not { } rawValue)
                {
                    continue;
                }

                double categoryY = plotY + category * categoryHeight + (categoryHeight - barHeight) / 2d;
                double value = NormalizeStackedValue(rawValue, positiveTotals[category], percentStacked);
                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, values.Count, value, barOptions.InvertsNegativePoint(seriesIndex, category));
                double segmentStartValue;
                double segmentEndValue;
                if (value >= 0d)
                {
                    segmentStartValue = positiveValues[category];
                    positiveValues[category] += value;
                    segmentEndValue = positiveValues[category];
                }
                else
                {
                    segmentStartValue = negativeValues[category];
                    negativeValues[category] += value;
                    segmentEndValue = negativeValues[category];
                }

                double startX = ChartValueToPlotCoordinate(valueExtents, segmentStartValue, plotX, plotWidth, valueAxisReversed);
                double endX = ChartValueToPlotCoordinate(valueExtents, segmentEndValue, plotX, plotWidth, valueAxisReversed);
                double segmentX = Math.Min(startX, endX);
                double segmentWidth = Math.Abs(endX - startX);
                AddStackedBarFillRun(fillRuns, fill, new ChartRectangle(segmentX, categoryY, segmentWidth, barHeight));
                strokeSegments.Add(new ChartStackedBarSegment(seriesIndex, category, segmentX, categoryY, segmentWidth, barHeight, value));
            }

            bool seriesDriven = (seriesIndex >= seriesFills.Count || seriesFills[seriesIndex] is null) &&
                (seriesIndex >= pointFills.Count || pointFills[seriesIndex].Count == 0);
            foreach (ChartStackedBarFillRun run in fillRuns)
            {
                if (seriesDriven && TryGetStyleBarGradientStops(chartStyleId, run.Fill, out IReadOnlyList<PdfShadingStop> stops))
                {
                    foreach (ChartRectangle rectangle in run.Rectangles)
                    {
                        PaintBarGradientRect(graphics, plotBox, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height, stops);
                    }
                }
                else
                {
                    FillChartRectanglesAsCompoundPathInPlotClip(graphics, plotBox, run.Rectangles, run.Fill);
                }
            }
        }

        foreach (ChartStackedBarSegment segment in strokeSegments)
        {
            StrokeChartPointRectangleInPlotClip(
                graphics,
                plotBox,
                segment.SeriesIndex,
                segment.CategoryIndex,
                pointStrokes,
                segment.X,
                segment.Y,
                segment.Width,
                segment.Height,
                ResolveNegativeBarFallbackStroke(pointStrokes, segment.SeriesIndex, segment.CategoryIndex, segment.Value, barOptions.InvertsNegativePoint(segment.SeriesIndex, segment.CategoryIndex)));
        }
    }

}
