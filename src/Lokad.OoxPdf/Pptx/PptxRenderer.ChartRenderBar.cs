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
    private static void RenderBarChart(PdfGraphicsBuilder graphics, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ChartLayoutBox plotAreaBox, ChartPlotBox plotBox, IReadOnlyList<ChartIndexedNumberVector> series, bool horizontalBars, ChartBarPlotOptions plotOptions, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, ChartValueAxisRenderOptions valueAxisOptions, ChartAxesStyle axesStyle, ChartShapeStyle plotAreaStyle, ChartValueExtents valueExtents, bool valueAxisLabelsVisible, bool manualPlotLayoutApplied)
    {
        double plotX = plotBox.X;
        double plotY = plotBox.Y;
        double plotWidth = plotBox.Width;
        double plotHeight = plotBox.Height;
        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> denseSeries = DensifyChartPointSeries(series);
        RenderChartShapeStyle(graphics, plotAreaBox.X, plotAreaBox.Y, plotAreaBox.Width, plotAreaBox.Height, plotAreaStyle);
        {
            int categoryCount = Math.Max(1, denseSeries.Max(values => values.Count));
            bool stacked = IsStackedChartGrouping(plotOptions.Grouping);
            bool percentStacked = IsPercentStackedChartGrouping(plotOptions.Grouping);
            double zeroX = ChartValueToPlotCoordinate(valueExtents, 0d, plotX, plotWidth, valueAxisOptions.Reversed);
            double zeroY = ChartValueToPlotCoordinate(valueExtents, 0d, plotY, plotHeight, horizontalBars ? false : valueAxisOptions.Reversed);
            double valueAxisCrossingY = ChartValueToPlotCoordinate(valueExtents, valueAxisOptions.CrossingValue, plotY, plotHeight, horizontalBars ? false : valueAxisOptions.Reversed);
            double valueAxisAutoTickTargetCount = GetValueAxisAutoTickTargetCount(horizontalBars, valueAxisLabelsVisible, manualPlotLayoutApplied);
            if (valueAxisOptions.MinorGridlines)
            {
                if (horizontalBars)
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotBox,
                        () => DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions.Units.MinorUnit, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, major: false, valueAxisOptions.GridlineStyle.Minor, PptxChartMetricRules.AxisNiceTickTargetCount));
                }
                else
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotBox,
                        () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions.Units.MinorUnit, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, major: false, valueAxisOptions.GridlineStyle.Minor));
                }
            }

            if (valueAxisOptions.MajorGridlines)
            {
                if (horizontalBars)
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotBox,
                        () => DrawVerticalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions.Units.MajorUnit, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, major: true, valueAxisOptions.GridlineStyle.Major, valueAxisAutoTickTargetCount));
                }
                else
                {
                    RenderInChartPlotAreaClip(
                        graphics,
                        plotBox,
                        () => DrawHorizontalChartGridlines(graphics, plotX, plotY, plotWidth, plotHeight, valueExtents, valueAxisOptions.Units.MajorUnit, valueAxisOptions.CrossingValue, valueAxisOptions.Reversed, major: true, valueAxisOptions.GridlineStyle.Major));
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
                    RenderStackedHorizontalBars(graphics, plotBox, theme, colorMap, chartPalette, plotX, plotY, plotWidth, plotHeight, denseSeries, categoryCount, valueExtents, valueAxisOptions.Reversed, percentStacked, seriesFills, pointFills, pointStrokes, plotOptions.VaryColors.Value, plotOptions.GapWidth);
                }
                else
                {
                    RenderClusteredHorizontalBars(graphics, plotBox, theme, colorMap, chartPalette, plotX, plotY, plotWidth, plotHeight, denseSeries, categoryCount, valueExtents, valueAxisOptions.Reversed, zeroX, seriesFills, pointFills, pointStrokes, plotOptions.VaryColors.Value, plotOptions.GapWidth, plotOptions.Overlap);
                }

                return;
            }

            if (stacked)
            {
                RenderStackedColumns(graphics, plotBox, theme, colorMap, chartPalette, plotX, plotY, plotWidth, plotHeight, denseSeries, categoryCount, valueExtents, valueAxisOptions.Reversed, percentStacked, seriesFills, pointFills, pointStrokes, plotOptions.VaryColors.Value, plotOptions.GapWidth);
                return;
            }

            double categoryWidth = plotWidth / categoryCount;
            double barWidth = GetClusteredBarWidth(categoryWidth, denseSeries.Count, plotOptions.GapWidth);
            double step = GetClusteredBarStep(barWidth, plotOptions.Overlap);
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

                    ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, denseSeries.Count, plotOptions.VaryColors.Value, seriesFills, pointFills, value);
                    double barX = categoryX + seriesIndex * step;
                    double valueY = ChartValueToPlotCoordinate(valueExtents, value, plotY, plotHeight, valueAxisOptions.Reversed);
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
        ChartWorkbookData? workbook,
        bool plotVisibleOnly,
        PresentationFontResolver? fontResolver)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        string? title = ReadSceneOrXmlChartTitleText(sceneChart, chartXml);
        PptxSceneChartTextBodyProperties titleTextBodyProperties = ReadSceneOrXmlChartTitleTextBodyProperties(sceneChart, chartXml);
        ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
        ChartPlotLayout plotLayout = GetBarChartPlotLayout(theme, frame, chartXml, sceneChart, barPlot, barChart, title, legend, barOptions, workbook, plotVisibleOnly, fontResolver: fontResolver, ignoreManualPlotLayout: false);
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
        bool ignoreManualPlotLayout,
        PresentationFontResolver? fontResolver)
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

        defaultPlotBox = AdjustBarChartPlotBoxForVisibleValueAxes(defaultPlotBox);
        defaultPlotBox = AdjustBarChartPlotBoxForStackedValueAxisLabels(defaultPlotBox);
        defaultPlotBox = AdjustStackedColumnBottomLegendPlotBox(defaultPlotBox, frame, horizontalBars, barOptions.Grouping, hasTitle, legend);
        defaultPlotBox = AdjustBarChartPlotBoxForDefaultAxisTitles(defaultPlotBox, horizontalBars, hasTitle, hasLegend);
        if (ignoreManualPlotLayout)
        {
            return ChartPlotLayout.FromPlotBox(defaultPlotBox);
        }

        ChartPlotBox manualDefaultPlotBox = horizontalBars && HasRecognizedManualPlotLayoutTarget()
            ? GetHorizontalBarManualLayoutTargetDefaultPlotBox(defaultPlotBox)
            : defaultPlotBox;
        if (!TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, manualDefaultPlotBox, out ChartPlotLayout manualPlotLayout))
        {
            return ChartPlotLayout.FromPlotBox(defaultPlotBox);
        }

        return ResolveBarManualPlotLayoutTarget(theme, chartXml, sceneChart, barPlot, barChart, manualPlotLayout, horizontalBars);

        ChartPlotBox GetHorizontalBarManualLayoutTargetDefaultPlotBox(ChartPlotBox defaultPlotBox)
        {
            double x = frame.X + frame.Width * PptxChartMetricRules.HorizontalBarManualLayoutTargetPlotBoxXRatio;
            double y = frame.Y + frame.Height * PptxChartMetricRules.HorizontalBarManualLayoutTargetPlotBoxYRatio;
            return new ChartPlotBox(x, y, defaultPlotBox.Width, defaultPlotBox.Height);
        }

        ChartPlotBox AdjustBarChartPlotBoxForDefaultAxisTitles(ChartPlotBox plotBox, bool horizontalBars, bool hasChartTitle, bool hasLegend)
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

        ChartPlotBox AdjustBarChartPlotBoxForStackedValueAxisLabels(ChartPlotBox plotBox)
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
                percentStacked, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
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

        ChartPlotBox AdjustBarChartPlotBoxForVisibleValueAxes(ChartPlotBox plotBox)
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
            double leftStripWidth = 0d;
            double rightStripWidth = 0d;
            int leftStripCount = 0;
            int rightStripCount = 0;
            foreach (ChartAxisSource valueAxis in valueAxes)
            {
                ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(
                    valueAxis.SceneAxis,
                    valueAxis.XmlAxis,
                    new ChartValueExtents(0d, 1d), false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
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
                    rightStripWidth = Math.Max(rightStripWidth, stripWidth);
                    rightStripCount++;
                }
                else
                {
                    leftStripWidth = Math.Max(leftStripWidth, stripWidth);
                    leftStripCount++;
                }
            }

            double requiredLeftReserve = Math.Max(leftReserve, leftStripWidth * GetMultiValueAxisStripFactor(leftStripCount, labelsRight: false));
            double requiredRightReserve = Math.Max(rightReserve, rightStripWidth * GetMultiValueAxisStripFactor(rightStripCount, labelsRight: true));
            double x = frame.X + requiredLeftReserve;
            double right = frame.X + frame.Width - requiredRightReserve;
            double width = Math.Max(1d, right - x);
            return new ChartPlotBox(x, plotBox.Y, width, plotBox.Height);
        }

        bool HasInsideValueAxisCrossing()
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
                IsPercentStackedChartGrouping(grouping), false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            double? crossing = ReadSceneOrXmlValueAxisCrossingValue(valueAxis.SceneAxis, valueAxis.XmlAxis, valueExtents);
            return crossing > valueExtents.Min + PptxChartMetricRules.AxisValueEpsilon &&
                crossing < valueExtents.Max - PptxChartMetricRules.AxisValueEpsilon;
        }

        bool HasRecognizedManualPlotLayoutTarget()
        {
            if (sceneChart is not null)
            {
                return sceneChart.PlotAreaLayout.HasLayout &&
                    sceneChart.PlotAreaLayout.LayoutTargetKind != PptxSceneChartManualLayoutTarget.Unknown;
            }

            PptxSceneChartManualLayout layout = PptxSceneBuilder.ReadChartPlotAreaManualLayout(chartXml);
            return layout.HasLayout &&
                layout.LayoutTargetKind != PptxSceneChartManualLayoutTarget.Unknown;
        }
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

        ChartPlotBox plotBox = DeriveHorizontalBarInnerPlotBox(layout.PlotAreaBox);
        return new ChartPlotLayout(layout.PlotAreaBox, plotBox, layout.ManualLayoutTargetKind);

        ChartPlotBox DeriveHorizontalBarInnerPlotBox(ChartLayoutBox plotAreaBox)
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
    }

    private static double GetMultiValueAxisStripFactor(int sameSideAxisCount, bool labelsRight)
    {
        if (sameSideAxisCount <= 1)
        {
            return 1d;
        }

        return labelsRight
            ? PptxChartMetricRules.BarMultiValueAxisSecondaryStripFactor
            : PptxChartMetricRules.BarMultiValueAxisPrimaryStripFactor;
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

        double left = Math.Min(
            frame.X + frame.Width,
            plotBox.X + PptxChartMetricRules.StackedColumnBottomLegendPlotBoxLeftPadding);
        double right = Math.Max(
            left + 1d,
            plotBox.X + plotBox.Width - PptxChartMetricRules.StackedColumnBottomLegendPlotBoxRightPadding);
        return new ChartPlotBox(left, plotBox.Y, right - left, plotBox.Height);
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
        foreach (double value in GetChartGridlineValues(extents, explicitUnit, crossingValue, PptxChartMetricRules.AxisNiceTickTargetCount))
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

    private static void DrawVerticalChartGridlines(PdfGraphicsBuilder graphics, double plotX, double plotY, double plotWidth, double plotHeight, ChartValueExtents extents, double? explicitUnit, double? crossingValue, bool reversed, bool major, ChartSeriesStroke? gridlineStroke, double autoTickTargetCount)
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

    private static ChartSeriesFill ChartSeriesColor(int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(seriesIndex), defaultAlpha, null, null);
    }

    private static ChartSeriesFill ChartSeriesColor(PptxTheme theme, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(null, theme, seriesIndex), defaultAlpha, null, null);
    }

    private static ChartSeriesFill ChartSeriesColor(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(chartPalette, theme, seriesIndex), defaultAlpha, null, null);
    }

    private static ChartSeriesFill ChartSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, IReadOnlyList<ChartSeriesFill?> seriesFills, double defaultAlpha)
    {
        return seriesIndex < seriesFills.Count && seriesFills[seriesIndex] is { } fill
            ? fill
            : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, seriesIndex), defaultAlpha, null, null);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null)
            ? new ChartSeriesFill(ChartPalette(categoryIndex), 1d, null, null)
            : ChartSeriesColor(seriesIndex, seriesFills, 1d);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(PptxTheme theme, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null)
            ? new ChartSeriesFill(ChartPalette(null, theme, categoryIndex), 1d, null, null)
            : ChartSeriesColor(theme, seriesIndex, seriesFills, 1d);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(PptxTheme theme, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null)
            ? new ChartSeriesFill(ChartPalette(chartPalette, theme, categoryIndex), 1d, null, null)
            : ChartSeriesColor(theme, chartPalette, seriesIndex, seriesFills, 1d);
    }

    private static ChartSeriesFill ChartCategoryOrSeriesColor(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills)
    {
        return varyColors && seriesCount == 1 && (seriesFills.Count == 0 || seriesFills[0] is null)
            ? new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, categoryIndex), 1d, null, null)
            : ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, seriesFills, 1d);
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
            return new ChartSeriesFill(new RgbColor(255, 255, 255), 1d, null, null);
        }

        return ChartPointCategoryOrSeriesColor(theme, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, pointFills);
    }

    private static ChartSeriesFill ResolveBarPointFill(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, int seriesIndex, int categoryIndex, int seriesCount, bool varyColors, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, double value)
    {
        if (value < 0d && !HasExplicitChartPointFill(pointFills, seriesIndex, categoryIndex))
        {
            return new ChartSeriesFill(new RgbColor(255, 255, 255), 1d, null, null);
        }

        return ChartPointCategoryOrSeriesColor(theme, colorMap, chartPalette, seriesIndex, categoryIndex, seriesCount, varyColors, seriesFills, pointFills);
    }

    private static bool HasExplicitChartPointFill(IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, int seriesIndex, int categoryIndex)
    {
        return seriesIndex < pointFills.Count && pointFills[seriesIndex].ContainsKey(categoryIndex);
    }

    private static ChartSeriesStroke? ResolveNegativeBarFallbackStroke(IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, int seriesIndex, int categoryIndex, double value)
    {
        return value < 0d && !HasExplicitChartPointStroke()
            ? ChartNegativeBarDefaultStroke
            : null;

        bool HasExplicitChartPointStroke()
        {
            return seriesIndex < pointStrokes.Count && pointStrokes[seriesIndex].ContainsKey(categoryIndex);
        }
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

    private static bool IsDarkChartPattern(string patternPreset)
    {
        return patternPreset.StartsWith("dk", StringComparison.OrdinalIgnoreCase);
    }

    private static (double Min, double Max) GetClusteredPointValueExtents(IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series)
    {
        double maxValue = Math.Max(0d, series.SelectMany(points => points).Select(point => point?.Value).OfType<double>().DefaultIfEmpty(0d).Max());
        double minValue = Math.Min(0d, series.SelectMany(points => points).Select(point => point?.Value).OfType<double>().DefaultIfEmpty(0d).Min());
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
        double[] positiveValues = new double[categoryCount];
        double[] negativeValues = new double[categoryCount];
        double[] positiveTotals = GetCategoryPositiveTotals(series, categoryCount, percentStacked);
        var strokeSegments = new List<ChartStackedBarSegment>();
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            var fillRuns = new List<ChartStackedBarFillRun>();
            for (int category = 0; category < categoryCount; category++)
            {
                IReadOnlyList<ChartIndexedNumberPoint?> values = series[seriesIndex];
                if (category >= values.Count || values[category]?.Value is not { } rawValue)
                {
                    continue;
                }

                double categoryX = plotX + category * categoryWidth + (categoryWidth - barWidth) / 2d;
                double value = NormalizeStackedValue(rawValue, positiveTotals[category], percentStacked);
                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, value);
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

            foreach (ChartStackedBarFillRun run in fillRuns)
            {
                FillChartRectanglesAsCompoundPathInPlotClip(graphics, plotBox, run.Rectangles, run.Fill);
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
                ResolveNegativeBarFallbackStroke(pointStrokes, segment.SeriesIndex, segment.CategoryIndex, segment.Value));
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

    private static void RenderStackedHorizontalBars(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, bool percentStacked, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent)
    {
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
                IReadOnlyList<ChartIndexedNumberPoint?> values = series[seriesIndex];
                if (category >= values.Count || values[category]?.Value is not { } rawValue)
                {
                    continue;
                }

                double categoryY = plotY + category * categoryHeight + (categoryHeight - barHeight) / 2d;
                double value = NormalizeStackedValue(rawValue, positiveTotals[category], percentStacked);
                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, value);
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

            foreach (ChartStackedBarFillRun run in fillRuns)
            {
                FillChartRectanglesAsCompoundPathInPlotClip(graphics, plotBox, run.Rectangles, run.Fill);
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
                ResolveNegativeBarFallbackStroke(pointStrokes, segment.SeriesIndex, segment.CategoryIndex, segment.Value));
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

    private static double[] GetCategoryPositiveTotals(IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, int categoryCount, bool percentStacked)
    {
        var totals = new double[categoryCount];
        if (!percentStacked)
        {
            Array.Fill(totals, 1d);
            return totals;
        }

        for (int category = 0; category < categoryCount; category++)
        {
            totals[category] = GetCategoryPositiveTotal(series, category, percentStacked);
        }

        return totals;
    }

    private static double NormalizeStackedValue(double value, double positiveTotal, bool percentStacked)
    {
        return percentStacked && value > 0d ? value / positiveTotal : value;
    }
}