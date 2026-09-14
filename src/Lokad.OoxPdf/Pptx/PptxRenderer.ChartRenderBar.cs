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
            double columnBaseY = !horizontalBars && axesStyle.CategoryAxisTopSide
                ? ChartValueToPlotCoordinate(valueExtents, valueExtents.Max, plotY, plotHeight, valueAxisOptions.Reversed)
                : zeroY;
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
                        : (axesStyle.CategoryAxisTopSide ? plotY + plotHeight : valueAxisCrossingY);
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
            double barWidth = GetClusteredBarWidth(categoryWidth, denseSeries.Count, plotOptions.GapWidth, plotOptions.Overlap);
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

                    ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, denseSeries.Count, plotOptions.VaryColors.Value, seriesFills, pointFills, values.Count, true, value);
                    double barX = categoryX + seriesIndex * step;
                    double valueY = ChartValueToPlotCoordinate(valueExtents, value, plotY, plotHeight, valueAxisOptions.Reversed);
                    double barY = Math.Min(columnBaseY, valueY);
                    double barHeight = Math.Abs(valueY - columnBaseY);
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
        ChartPlotLayout plotLayout = GetBarChartPlotLayout(theme, colorMap, frame, chartXml, sceneChart, barPlot, barChart, title, legend, barOptions, workbook, plotVisibleOnly, fontResolver: fontResolver, ignoreManualPlotLayout: false);
        return new ChartLayout(frame, plotLayout.PlotAreaBox, plotLayout.PlotBox, plotLayout.ManualLayoutTargetKind is not null, title, titleTextBodyProperties, legend);
    }

    private static double ReadSceneOrXmlCategoryTitleFontSizeAtPosition(PptxTheme theme, PptxColorMap colorMap, PptxSceneChart? sceneChart, XDocument chartXml, PptxSceneChartAxisPosition titlePosition)
    {
        // Horizontal axis titles render on the opposite side of their recorded position
        // (Top hangs below the plot, Bottom sits above it). The caller sizes each reserve var
        // from the title recorded on its own side: Bottom-positioned titles size bottomReserve
        // (which anchors the plot top edge y), Top-positioned titles size topReserve (which sets
        // the height and hence the bottom edge). Category-kind only: value-kind horizontal
        // titles keep legacy output, unobserved. Run sizes win over paragraph defaults,
        // mirroring emission.
        if (sceneChart is not null)
        {
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

        foreach (XElement axis in ReadChartAxisElements(chartXml))
        {
            PptxSceneChartAxisKind axisKind = PptxSceneBuilder.ParseChartAxisKind(axis.Name.LocalName);
            PptxSceneChartAxisPosition positionKind = PptxSceneBuilder.ParseChartAxisPosition(PptxSceneBuilder.ReadChartElementValue(axis, "axPos"));
            XElement? title = axis.Element(ChartNamespace + "title");
            if (positionKind != titlePosition || title is null || !IsBottomPlacedCategoryTitleCandidate(axisKind, PptxSceneBuilder.ReadChartText(title.Element(ChartNamespace + "tx"), trimLiteral: true), PptxSceneBuilder.ReadChartManualLayout(title).HasLayout))
            {
                continue;
            }

            foreach (XElement runProperties in title.Descendants(OoxNamespaces.DrawingNamespace + "rPr"))
            {
                if (runProperties.Attribute("sz") is { } sizeAttribute && int.TryParse(sizeAttribute.Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int sizeHundredths) && sizeHundredths > 0)
                {
                    return sizeHundredths / 100d;
                }
            }

            ChartTextStyleOverride titleStyle = ToChartTextStyleOverride(PptxSceneBuilder.ReadChartTextStyleOverride(title, theme, colorMap));
            if (titleStyle.FontSize is { } xmlSize && xmlSize > 0d)
            {
                return xmlSize;
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

    private static ChartPlotLayout GetBarChartPlotLayout(
        PptxTheme theme,
        PptxColorMap colorMap,
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
            defaultPlotBox = AdjustNoTitleBottomLegendPlotBoxForMeasuredContent(defaultPlotBox, frame, theme, colorMap, sceneChart, chartXml, barPlot, barChart);
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
        defaultPlotBox = AdjustBarChartPlotBoxForSingleValueAxisLabels(defaultPlotBox);
        defaultPlotBox = AdjustBarChartPlotBoxForCategoryLabels(defaultPlotBox);
        defaultPlotBox = AdjustStackedColumnBottomLegendPlotBox(defaultPlotBox, frame, horizontalBars, barOptions.Grouping, hasTitle, legend);
        defaultPlotBox = AdjustBarChartPlotBoxForDefaultAxisTitles(defaultPlotBox, horizontalBars, hasLegend);
        defaultPlotBox = AdjustHorizontalBarPlotBoxForCategoryLabels(defaultPlotBox);
        defaultPlotBox = AdjustHorizontalBarPlotBoxForValueLabels(defaultPlotBox);
        defaultPlotBox = AdjustHorizontalBarPlotBoxForTopAndBottom(defaultPlotBox);
        defaultPlotBox = AdjustVerticalBarPlotBoxTopFloor(defaultPlotBox, frame, horizontalBars, hasTitle, hasLegend, HasRenderableCategoryLabels(barPlot, barChart, workbook, plotVisibleOnly));
        ChartAxisSource rightReserveCategoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
        bool rightReserveLabelsVisible = IsSceneOrXmlChartAxisLabelVisible(rightReserveCategoryAxis.SceneAxis, rightReserveCategoryAxis.XmlAxis);
        defaultPlotBox = AdjustVerticalBarPlotBoxRightReserve(defaultPlotBox, frame, horizontalBars, legend, HasRenderableCategoryLabels(barPlot, barChart, workbook, plotVisibleOnly) && rightReserveLabelsVisible);
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

        return ResolveBarManualPlotLayoutTarget(theme, chartXml, sceneChart, barPlot, barChart, manualPlotLayout, horizontalBars, frame, workbook, plotVisibleOnly, fontResolver);

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
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, barPlot, chartXml, barChart).FirstOrDefault();
                if (IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    ChartTextStyle valueTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                    topReserve = ComputeDefaultAxisTitleHorizontalBarReserves(valueTickStyle.FontSize, 0d).TopReserve;
                }

                if (hasTitle)
                {
                    ChartTextStyle titleStyle = ReadSceneOrXmlChartTitleTextStyle(theme, colorMap, sceneChart, chartXml, IsAutoGeneratedChartTitle(sceneChart));
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
                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
                if (!IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    return legacyReserve;
                }

                ChartTextStyle categoryTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, chartXml, categoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                double categoryTitleFontSize = ReadSceneOrXmlCategoryTitleFontSizeAtPosition(theme, colorMap, sceneChart, chartXml, titlePosition);
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

        ChartPlotBox AdjustHorizontalBarPlotBoxForValueLabels(ChartPlotBox plotBox)
        {
            if (!horizontalBars)
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

            bool percentStacked = IsPercentStackedChartGrouping(barOptions.Grouping);
            ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetBarChartValueExtents(seriesVectors, barOptions.Grouping), percentStacked, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            ChartAxisUnits axisUnits = ReadSceneOrXmlChartValueAxisRenderOptions(valueAxis.SceneAxis, valueAxis.XmlAxis, theme, valueExtents, percentStacked).Units;
            ChartTextStyle tickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
            var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
            double maxLabelWidth = 0d;
            foreach (double tickValue in GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceHorizontalValueTickTargetCount))
            {
                string label = FormatSceneOrXmlChartAxisLabel(tickValue, valueAxis.SceneAxis, valueAxis.XmlAxis, percentStacked ? "0%" : null);
                if (!string.IsNullOrWhiteSpace(label))
                {
                    maxLabelWidth = Math.Max(maxLabelWidth, textMeasurer.Measure(label, tickStyle));
                }
            }

            if (maxLabelWidth <= 0d)
            {
                return plotBox;
            }

            double rightReserve = ResolveMeasuredRightReserve(frame.X + frame.Width - plotBox.X - plotBox.Width, maxLabelWidth);
            double right = frame.X + frame.Width - rightReserve;
            double width = Math.Max(1d, right - plotBox.X);
            return new ChartPlotBox(plotBox.X, plotBox.Y, width, plotBox.Height);
        }

        ChartPlotBox AdjustHorizontalBarPlotBoxForTopAndBottom(ChartPlotBox plotBox)
        {
            if (!horizontalBars)
            {
                return plotBox;
            }

            if (hasTitle || hasLegend || legend.Visible)
            {
                return plotBox;
            }

            ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, barPlot, chartXml, barChart).FirstOrDefault();
            if (!IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
            {
                return plotBox;
            }

            IReadOnlyList<ChartIndexedNumberVector> topBottomSeriesVectors = ReadSceneOrXmlChartSeriesVectors(barPlot, barChart, workbook, plotVisibleOnly);
            if (CountRenderableSeries(topBottomSeriesVectors) == 0)
            {
                return plotBox;
            }

            bool percentBottomStacked = IsPercentStackedChartGrouping(barOptions.Grouping);
            ChartValueExtents bottomValueExtents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetBarChartValueExtents(topBottomSeriesVectors, barOptions.Grouping), percentBottomStacked, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            ChartAxisUnits bottomAxisUnits = ReadSceneOrXmlChartValueAxisRenderOptions(valueAxis.SceneAxis, valueAxis.XmlAxis, theme, bottomValueExtents, percentBottomStacked).Units;
            ChartTextStyle bottomTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
            bool hasBottomValueLabel = false;
            foreach (double bottomTickValue in GetChartAxisTickValues(bottomValueExtents, bottomAxisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceHorizontalValueTickTargetCount))
            {
                string bottomLabel = FormatSceneOrXmlChartAxisLabel(bottomTickValue, valueAxis.SceneAxis, valueAxis.XmlAxis, percentBottomStacked ? "0%" : null);
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

            double requiredBottomReserve = ComputeBarLabelStripBottomReserve(bottomTickStyle.FontSize);
            double presetBottomReserve = plotBox.Y - frame.Y;
            double newY = Math.Abs(presetBottomReserve - requiredBottomReserve) < PptxChartMetricRules.HorizontalBarPlotFloorSlop
                ? plotBox.Y
                : frame.Y + requiredBottomReserve;
            double presetTopReserve = frame.Y + frame.Height - plotBox.Y - plotBox.Height;
            double newTop = Math.Abs(presetTopReserve - PptxChartMetricRules.HorizontalBarPlotTopReserve) < PptxChartMetricRules.HorizontalBarPlotFloorSlop
                ? plotBox.Y + plotBox.Height
                : frame.Y + frame.Height - PptxChartMetricRules.HorizontalBarPlotTopReserve;
            return new ChartPlotBox(plotBox.X, newY, plotBox.Width, Math.Max(1d, newTop - newY));
        }

        ChartPlotBox AdjustHorizontalBarPlotBoxForCategoryLabels(ChartPlotBox plotBox)
        {
            if (!horizontalBars)
            {
                return plotBox;
            }

            ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
            if (!IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
            {
                return plotBox;
            }

            ChartTextStyle tickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, chartXml, categoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
            var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
            double maxCategoryWidth = 0d;
            foreach (ChartIndexedTextPoint? label in ReadSceneOrXmlCategoryLabelVector(barPlot, barChart, workbook, plotVisibleOnly).DensePoints())
            {
                string? labelText = label?.Text;
                if (!string.IsNullOrWhiteSpace(labelText))
                {
                    maxCategoryWidth = Math.Max(maxCategoryWidth, textMeasurer.Measure(labelText, tickStyle));
                }
            }

            if (maxCategoryWidth <= 0d)
            {
                return plotBox;
            }

            double requiredReserve = ComputeHorizontalBarCategoryLeftReserve(maxCategoryWidth);
            double leftReserve = plotBox.X - frame.X;
            double rightReserve = frame.X + frame.Width - plotBox.X - plotBox.Width;
            bool labelsRight = ResolveSceneOrXmlCategoryAxisRightSide(categoryAxis.SceneAxis, categoryAxis.XmlAxis, defaultRightSide: false);
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

        ChartPlotBox AdjustBarChartPlotBoxForSingleValueAxisLabels(ChartPlotBox plotBox)
        {
            if (horizontalBars)
            {
                return plotBox;
            }

            if (IsStackedChartGrouping(barOptions.Grouping))
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

            ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetBarChartValueExtents(seriesVectors, barOptions.Grouping), false, ResolveBarValueAxisHeadroom(horizontalBars, IsPercentStackedChartGrouping(barOptions.Grouping)), PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxis.XmlAxis), false);
            ChartTextStyle tickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
            var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
            double maxLabelWidth = 0d;
            foreach (double tickValue in GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount))
            {
                string label = FormatSceneOrXmlChartAxisLabel(tickValue, valueAxis.SceneAxis, valueAxis.XmlAxis, defaultNumberFormat: null);
                if (!string.IsNullOrWhiteSpace(label))
                {
                    maxLabelWidth = Math.Max(maxLabelWidth, textMeasurer.Measure(label, tickStyle));
                }
            }

            if (maxLabelWidth <= 0d)
            {
                return plotBox;
            }

            double requiredReserve = ComputeBarValueAxisLeftReserve(maxLabelWidth, tickStyle.FontSize);
            double leftReserve = plotBox.X - frame.X;
            double rightReserve = frame.X + frame.Width - plotBox.X - plotBox.Width;
            bool labelsRight = ResolveSceneOrXmlValueAxisLabelsRightSide(valueAxis.SceneAxis, valueAxis.XmlAxis, defaultRightSide: false);
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

        ChartPlotBox AdjustBarChartPlotBoxForCategoryLabels(ChartPlotBox plotBox)
        {
            if (horizontalBars)
            {
                return plotBox;
            }

            // Stacked columns share the measured bottom strip (Office bottom margins fit
            // the same 7.0 plus 1.813fs rule within 1pt on ladder, compact and overlay
            // probes); titled/legend/label-less charts keep legacy paths below.
            if (hasTitle || hasLegend || legend.Visible)
            {
                return plotBox;
            }

            ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
            if (!IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
            {
                return plotBox;
            }

            IReadOnlyList<ChartIndexedNumberVector> categorySeriesVectors = ReadSceneOrXmlChartSeriesVectors(barPlot, barChart, workbook, plotVisibleOnly);
            if (CountRenderableSeries(categorySeriesVectors) == 0)
            {
                return plotBox;
            }

            if (!HasRenderableCategoryLabels(barPlot, barChart, workbook, plotVisibleOnly))
            {
                return plotBox;
            }

            ChartTextStyle categoryTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, chartXml, categoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
            double requiredBottomReserve = ComputeBarLabelStripBottomReserve(categoryTickStyle.FontSize);
            double presetBottomReserve = plotBox.Y - frame.Y;
            if (Math.Abs(presetBottomReserve - requiredBottomReserve) < PptxChartMetricRules.BarValueAxisPresetFloorSlop)
            {
                return plotBox;
            }

            double newY = frame.Y + requiredBottomReserve;
            double plotTop = plotBox.Y + plotBox.Height;
            return new ChartPlotBox(plotBox.X, newY, plotBox.Width, Math.Max(1d, plotTop - newY));
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
                percentStacked, ResolveBarValueAxisHeadroom(barOptions.BarDirection == PptxSceneChartBarDirection.Bar, percentStacked), PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
            ChartAxisUnits axisUnits = ResolvePercentStackedAxisUnits(ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxis.XmlAxis), percentStacked);
            string? defaultNumberFormat = percentStacked ? "0%" : null;
            // Office single-value-axis origins decompose to frame + 6.5pt indent + tick width
            // + 0.92 * tick font size (column-stacked/clustered, dashboard, composite ports),
            // while multi-plot dual-axis charts pack their strips tighter (compact probe:
            // Office keeps a 17.6pt strip where the measured rule wants 24.1pt), so the
            // measured rule applies to single-plot stacked charts while multi-plot charts
            // keep the shared estimator.
            IReadOnlyList<PptxSceneChartPlot> stackedBarPlots = ReadSceneChartPlots(sceneChart, PptxSceneChartPlotKind.Bar);
            IReadOnlyList<XElement> stackedBarCharts = ReadSceneOrXmlChartPlotElements(sceneChart, chartXml, PptxSceneChartPlotKind.Bar);
            double requiredReserve;
            if (UseMeasuredStackedValueAxisReserve(Math.Max(stackedBarCharts.Count, stackedBarPlots.Count)))
            {
                ChartTextStyle stackedTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueAxis.SceneAxis, chartXml, valueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                var stackedTextMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
                double stackedMaxLabelWidth = 0d;
                foreach (double tickValue in GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount))
                {
                    string stackedLabel = FormatSceneOrXmlChartAxisLabel(tickValue, valueAxis.SceneAxis, valueAxis.XmlAxis, defaultNumberFormat);
                    if (!string.IsNullOrWhiteSpace(stackedLabel))
                    {
                        stackedMaxLabelWidth = Math.Max(stackedMaxLabelWidth, stackedTextMeasurer.Measure(stackedLabel, stackedTickStyle));
                    }
                }

                if (stackedMaxLabelWidth <= 0d)
                {
                    return plotBox;
                }

                requiredReserve = ComputeBarValueAxisLeftReserve(stackedMaxLabelWidth, stackedTickStyle.FontSize);
            }
            else
            {
                requiredReserve = EstimateVerticalValueAxisLabelStripWidth(theme, sceneChart, chartXml, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, axisUnits, defaultNumberFormat, fontResolver);
            }
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
            var leftStrips = new List<ChartValueAxisStripMeasure>();
            var rightStrips = new List<ChartValueAxisStripMeasure>();
            foreach (ChartAxisSource valueAxis in valueAxes)
            {
                string stripAxisId = valueAxis.SceneAxis?.Id ?? ReadChartAxisId(valueAxis.XmlAxis) ?? string.Empty;
                ChartValueExtents stripFallback = GetDualAxisStripFallbackExtents(sceneChart, chartXml, barPlots, barCharts, stripAxisId, workbook, plotVisibleOnly);
                ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(
                    valueAxis.SceneAxis,
                    valueAxis.XmlAxis,
                    stripFallback, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartAxisUnits units = ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxis.XmlAxis);
                ChartValueAxisStripMeasure strip = MeasureVerticalValueAxisLabelStrip(
                    theme,
                    sceneChart,
                    chartXml,
                    valueAxis.XmlAxis,
                    valueAxis.SceneAxis,
                    extents,
                    units,
                    defaultNumberFormat: null,
                    fontResolver: fontResolver);
                double stripWidth = strip.StripWidth;
                bool labelsRight = ResolveSceneOrXmlValueAxisLabelsRightSide(
                    valueAxis.SceneAxis,
                    valueAxis.XmlAxis,
                    ResolveSceneOrXmlValueAxisRightSide(valueAxis.SceneAxis, valueAxis.XmlAxis, defaultRightSide: false));
                if (labelsRight)
                {
                    rightStripWidth = Math.Max(rightStripWidth, stripWidth);
                    rightStrips.Add(strip);
                    rightStripCount++;
                }
                else
                {
                    leftStripWidth = Math.Max(leftStripWidth, stripWidth);
                    leftStrips.Add(strip);
                    leftStripCount++;
                }
            }

            // Same-side axis pairs chain their strips (Office lays dual left axes side by side)
            // instead of maxing one; right-side pairs keep the legacy factor (unobserved).
            double requiredLeftReserve = leftStrips.Count >= 2
                ? Math.Max(leftReserve, ComputeChainedLeftValueAxisReserve(leftStrips))
                : Math.Max(leftReserve, leftStripWidth * GetMultiValueAxisStripFactor(leftStripCount, labelsRight: false));
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
        ChartPlotBox plotBox = DeriveHorizontalBarOuterPlotBox(centeredArea);
        return new ChartPlotLayout(centeredArea, plotBox, layout.ManualLayoutTargetKind);

        // Office centers outer manual areas: six-probe calibration (x10/x18/x24/x30/w50
        // plus an x14 prediction verified within 0.1pt) shows area center = frame center
        // plus x-factor times frame width, clamped to fit the frame. Left-anchored factor
        // math (shared with the inner target) misses x18 by 7.9pt and diverges without
        // bound on larger x, while centering fits all six exactly.
        ChartLayoutBox ResolveCenteredOuterAreaBox(ChartLayoutBox areaBox)
        {
            PptxSceneChartManualLayout manual = sceneChart?.PlotAreaLayout ?? PptxSceneBuilder.ReadChartPlotAreaManualLayout(chartXml);
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
        ChartPlotBox DeriveHorizontalBarOuterPlotBox(ChartLayoutBox plotAreaBox)
        {
            ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
            double leftReserve = 0d;
            if (IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
            {
                ChartTextStyle tickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, chartXml, categoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                double labelOffsetScale = ResolveSceneOrXmlCategoryAxisLabelOffsetScale(categoryAxis.SceneAxis, categoryAxis.XmlAxis);
                int tickLabelSkip = ResolveSceneOrXmlCategoryAxisTickLabelSkip(categoryAxis.SceneAxis, categoryAxis.XmlAxis);
                var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
                double maxCategoryWidth = 0d;
                int labelIndex = 0;
                foreach (ChartIndexedTextPoint? label in ReadSceneOrXmlCategoryLabelVector(barPlot, barChart, workbook, plotVisibleOnly).DensePoints())
                {
                    // Mirror the emission skip logic in RenderChartCategoryLabels so the
                    // reserve covers rendered labels only.
                    string? reserveLabel = label?.Text;
                    if (labelIndex % tickLabelSkip == 0 && !string.IsNullOrWhiteSpace(reserveLabel))
                    {
                        maxCategoryWidth = Math.Max(maxCategoryWidth, textMeasurer.Measure(reserveLabel, tickStyle));
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

    // Chained same-side value-axis reserve: Office lays dual left axes side by side, each
    // with the single-axis indent/gap idiom (6.5pt frame indent, 0.92fs plot gap) and a
    // sideGap-plus-1.46 inter-axis gap (three Office renders agree within 0.1pt). Strips
    // run outer-to-inner in plot order; beyond-two and mixed-font configs unobserved.
    private static double ComputeChainedLeftValueAxisReserve(IReadOnlyList<ChartValueAxisStripMeasure> strips)
    {
        double reserve = PptxChartMetricRules.BarValueAxisLabelFrameIndent + strips[0].MaxLabelWidth;
        for (int stripIndex = 1; stripIndex < strips.Count; stripIndex++)
        {
            double sideGap = Math.Max(3d, strips[stripIndex].FontSize * PptxChartMetricRules.ValueAxisLabelSideGapFactor);
            reserve += sideGap + PptxChartMetricRules.BarDualValueAxisInterAxisGap + strips[stripIndex].MaxLabelWidth;
        }

        return reserve + strips[strips.Count - 1].FontSize * PptxChartMetricRules.BarValueAxisLabelGapFactor;
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
    private static bool HasRenderableCategoryLabels(PptxSceneChartPlot? barPlot, XElement barChart, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        foreach (ChartIndexedTextPoint? labelPoint in ReadSceneOrXmlCategoryLabelVector(barPlot, barChart, workbook, plotVisibleOnly).DensePoints())
        {
            if (!string.IsNullOrWhiteSpace(labelPoint?.Text))
            {
                return true;
            }
        }

        return false;
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
        PptxSceneChart? sceneChart,
        XDocument chartXml,
        PptxSceneChartPlot? barPlot,
        XElement barChart)
    {
        // Office keeps a content-sized bottom margin for untitled bottom-legend columns.
        // Four same-frame probes with varied fonts pin an additive plane (all within 0.06):
        // 55.02 over 18pt legend plus 8pt cats, 47.77 over 12pt plus 8pt, 66.10 over 18pt
        // plus 14pt, 58.86 over 12pt plus 14pt. The single-knob forms are dead (the error
        // flips sign across probes), and bottom-minus-cat-baseline scales cleanly at 1.566
        // catFs inside the total without its own usable law, so the plane ships as such.
        ChartTextStyle legendStyle = ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml);
        ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
        ChartTextStyle tickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, categoryAxis.SceneAxis, chartXml, categoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
        double bottomReserve = ComputeNoTitleBottomLegendReserve(legendStyle.FontSize, tickStyle.FontSize);
        double presetTop = frame.Y + frame.Height * (PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxYRatio + PptxChartMetricRules.BarNoTitleBottomLegendPlotBoxHeightRatio);
        double y = frame.Y + bottomReserve;
        return new ChartPlotBox(plotBox.X, y, plotBox.Width, Math.Max(1d, presetTop - y));
    }

    // Regime gate for stacked value-axis reserves: the measured indent-plus-font-gap rule
    // is calibrated on single-plot single-axis charts only; multi-plot charts (dual-axis
    // compact/overlay probes in corpus) keep the shared estimator.
    private static bool UseMeasuredStackedValueAxisReserve(int barPlotCount)
    {
        return barPlotCount == 1;
    }
    // Left reserve for horizontal-bar category labels: widest label plus the shared Office
    // axis-label-to-plot gap (bar-stacked-port decomposes exactly to frame + 50.0pt label +
    // 23.2pt, the same gap calibrated for value-axis labels).
    // Left reserve for single-plot non-stacked vertical-bar value labels: frame indent plus
    // widest tick label plus a font-relative gap (Office origins decompose to frame plus
    // 6.5pt plus tick width plus 0.92 times tick font size; the old shared 23.2pt gap
    // overshot narrow-tick titled columns like the composite port by 5.7pt). Stacked,
    // multi-axis and horizontal paths keep their own estimators.
    private static double ComputeBarValueAxisLeftReserve(double maxValueLabelWidth, double tickFontSize)
    {
        return maxValueLabelWidth +
            PptxChartMetricRules.BarValueAxisLabelFrameIndent +
            tickFontSize * PptxChartMetricRules.BarValueAxisLabelGapFactor;
    }

    // Bottom reserve for the label strip below a bar plot: frame margin plus descent
    // plus our emitted label gap (Office baselines frame plus 7.0 plus descent on
    // vertical category and horizontal value labels alike).
    private static double ComputeBarLabelStripBottomReserve(double labelFontSize)
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

    private static double ComputeHorizontalBarCategoryLeftReserve(double maxCategoryLabelWidth)
    {
        return maxCategoryLabelWidth + PptxChartMetricRules.LineRightLegendValueAxisPadding;
    }

    private readonly record struct ChartValueAxisStripMeasure(double MaxLabelWidth, double FontSize, double StripWidth);

    private static double EstimateVerticalValueAxisLabelStripWidth(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml, XElement? valueAxis, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits units, string? defaultNumberFormat, PresentationFontResolver? fontResolver)
    {
        return MeasureVerticalValueAxisLabelStrip(theme, sceneChart, chartXml, valueAxis, sceneAxis, extents, units, defaultNumberFormat, fontResolver).StripWidth;
    }

    private static ChartValueAxisStripMeasure MeasureVerticalValueAxisLabelStrip(PptxTheme theme, PptxSceneChart? sceneChart, XDocument chartXml, XElement? valueAxis, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits units, string? defaultNumberFormat, PresentationFontResolver? fontResolver)
    {
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
        double fontSize = style.FontSize;
        var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
        IReadOnlyList<double> tickValues = GetChartAxisTickValues(extents, units.MajorUnit, includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount);
        double maxLabelWidth = tickValues
            .Select(value => FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis, defaultNumberFormat))
            .DefaultIfEmpty("0")
            .Max(label => textMeasurer.Measure(label, style));
        double labelWidth = Math.Max(
            fontSize * PptxChartMetricRules.ValueAxisMinimumLabelWidthFactor,
            maxLabelWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
        double sideGap = Math.Max(3d, fontSize * PptxChartMetricRules.ValueAxisLabelSideGapFactor);
        return new ChartValueAxisStripMeasure(maxLabelWidth, fontSize, labelWidth + sideGap);
    }

    // Near-maximum headroom (the shared 0.96 rule) applies to vertical-bar value axes:
    // Office ceilings dataMax 68 to 80 on the ladder column probe. Horizontal bars keep
    // the legacy frozen behavior their tuned right-margin laws were calibrated against,
    // and percent stacks keep theirs (their max is pinned downstream anyway).
    private static bool ResolveBarValueAxisHeadroom(bool horizontalBars, bool percentStacked)
    {
        return !horizontalBars && !percentStacked;
    }

    // Dual-axis strip fallback: strips must measure the owning plot data extents, not
    // the legacy (0,1) dummy (auto-max dual probes otherwise render stale strips while the
    // render path nices real data; explicit axis bounds ignore the fallback downstream, so
    // explicit-max dual configs are untouched by construction).
    private static ChartValueExtents GetDualAxisStripFallbackExtents(PptxSceneChart? sceneChart, XDocument chartXml, IReadOnlyList<PptxSceneChartPlot> barPlots, IReadOnlyList<XElement> barCharts, string stripAxisId, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        for (int plotIndex = 0; plotIndex < Math.Max(barCharts.Count, barPlots.Count); plotIndex++)
        {
            PptxSceneChartPlot? stripPlot = plotIndex < barPlots.Count ? barPlots[plotIndex] : null;
            XElement? stripChart = plotIndex < barCharts.Count ? barCharts[plotIndex] : null;
            if (stripChart is null)
            {
                continue;
            }

            bool ownsAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, stripPlot, chartXml, stripChart)
                .Any(source => string.Equals(source.SceneAxis?.Id ?? ReadChartAxisId(source.XmlAxis) ?? string.Empty, stripAxisId, StringComparison.Ordinal));
            if (!ownsAxis)
            {
                continue;
            }

            ChartBarPlotOptions stripOptions = ReadSceneOrXmlChartBarOptions(stripPlot, stripChart, PptxSceneChartGrouping.Clustered);
            return GetBarChartValueExtents(ReadSceneOrXmlChartSeriesVectors(stripPlot, stripChart, workbook, plotVisibleOnly), stripOptions.Grouping);
        }

        return new ChartValueExtents(0d, 1d);
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

        // Office defaults gridlines to round joins (caps stay butt).
        SetChartStroke(graphics, stroke with { Join = stroke.Join ?? 1 });
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

        // Office defaults gridlines to round joins (caps stay butt).
        SetChartStroke(graphics, stroke with { Join = stroke.Join ?? 1 });
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

    private static bool IsDarkChartPattern(string patternPreset)
    {
        return patternPreset.StartsWith("dk", StringComparison.OrdinalIgnoreCase);
    }

    private static void RenderClusteredHorizontalBars(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, double plotX, double plotY, double plotWidth, double plotHeight, IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> series, int categoryCount, ChartValueExtents valueExtents, bool valueAxisReversed, double zeroX, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, bool varyColors, double gapWidthPercent, double overlapPercent)
    {
        double categoryHeight = plotHeight / categoryCount;
        double barHeight = GetClusteredBarWidth(categoryHeight, series.Count, gapWidthPercent, overlapPercent);
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

                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, values.Count, false, value);
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
                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, values.Count, true, value);
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
                ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, series.Count, varyColors, seriesFills, pointFills, values.Count, false, value);
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

}
