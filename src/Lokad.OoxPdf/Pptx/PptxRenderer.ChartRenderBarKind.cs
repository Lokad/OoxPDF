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
    private static async ValueTask<bool> TryRenderBarChartKindAsync(
        PdfGraphicsBuilder graphics,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        ChartColorPalette chartPalette,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart sceneChart,
        ChartWorkbookData? workbook,
        List<PdfFontResource> fonts,
        PresentationFontResolver fontResolver,
        bool plotVisibleOnly,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        IReadOnlyList<PptxSceneChartPlot> barPlots = ReadSceneChartPlots(sceneChart, PptxSceneChartPlotKind.Bar);
        if (barPlots.FirstOrDefault() is { } barPlot)
        {
            XElement barChart = barPlot.Source;
            IReadOnlyList<ChartIndexedNumberVector> barSeriesVectors = ReadSharedChartSeriesVectors(barPlot, workbook, plotVisibleOnly);
            int barSeriesCount = CountRenderableSeries(barSeriesVectors);
            if (barSeriesCount != 0)
            {
                ChartBarPlotOptions barOptions = ReadSceneChartBarOptions(barPlot, PptxSceneChartGrouping.Clustered);
                ChartColorPalette barPalette = WithSourceSeries(chartPalette, barPlot);
                bool horizontalBars = barOptions.BarDirection == PptxSceneChartBarDirection.Bar;
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneSeriesFills(barPlot);
                ChartAxesStyle axesStyle = ReadSceneChartAxesStyle(sceneChart, barPlot);
                ChartShapeStyle plotAreaStyle = ToChartShapeStyle(sceneChart.PlotAreaStyle);
                ChartAxisSource valueAxisSource = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
                XElement? valueAxis = valueAxisSource.XmlAxis;
                PptxSceneChartAxis? valueSceneAxis = valueAxisSource.SceneAxis;
                bool percentStacked = IsPercentStackedChartGrouping(barOptions.Grouping);
                bool sharedValueAxis = barPlot is not null && HasSharedCartesianChartValueAxis(sceneChart, valueSceneAxis);
                ChartValueExtents dataExtents = GetSharedBarChartValueExtents(barPlot, barOptions.Grouping, workbook, plotVisibleOnly);
                if (sharedValueAxis) dataExtents = GetSharedCartesianChartValueExtents(sceneChart, barPlot!, dataExtents,
                    workbook, plotVisibleOnly, context.CancellationToken);
                ChartValueExtents valueExtents = ReadCartesianChartValueAxisExtents(sceneChart, valueSceneAxis, valueAxis, dataExtents, percentStacked, ResolveBarValueAxisHeadroom(horizontalBars, percentStacked), PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(valueSceneAxis, valueAxis, theme, valueExtents, percentStacked);
                valueAxisOptions = valueAxisOptions with { TickPolicy = ReadCartesianChartAxisTickPolicy(sceneChart, valueSceneAxis, valueExtents, percentStacked) };
                ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
                bool ownsColumnCategories = barOptions.Grouping == PptxSceneChartGrouping.Clustered &&
                    CanOwnColumnCategoryPlacement(sceneChart, categoryAxis, valueAxisSource);
                barOptions = barOptions with
                {
                    CategoryAxisReversed = (horizontalBars || ownsColumnCategories) &&
                        barOptions.Grouping == PptxSceneChartGrouping.Clustered &&
                        ReadSceneValueAxisReversed(categoryAxis.SceneAxis),
                    CategoryPointsOnTickMarks = !horizontalBars && ownsColumnCategories &&
                        ResolveSceneCategoryAxisLabelsOnTickMarks(valueSceneAxis)
                };
                ChartAxisPlacement? resolvedCategory = horizontalBars
                    ? ReadHorizontalBarCategoryPlacement(sceneChart, categoryAxis, valueAxisSource, valueExtents, valueAxisOptions.Reversed) :
                    ReadChartAxisPlacement(categoryAxis.SceneAxis, valueExtents, valueAxisOptions.Reversed);
                if (resolvedCategory is { } resolved)
                {
                    valueAxisOptions = valueAxisOptions with { CrossingValue = resolved.CrossingValue };
                    axesStyle = horizontalBars
                        ? axesStyle with { CategoryAxisRightSide = resolved.AxisHighSide }
                        : axesStyle with { CategoryAxisTopSide = resolved.AxisHighSide };
                }
                ChartLegendLayout sourceLegend = ReadSceneChartLegendLayout(sceneChart);
                ChartAxisTitleReserveSides axisTitles = ReadDefaultAxisTitleReserveSides(sceneChart);
                ChartAxisPlacement? numericPlacement = horizontalBars &&
                    barOptions.Grouping == PptxSceneChartGrouping.Clustered && resolvedCategory is not null &&
                    UsesAutomaticBarFrame(barOptions, sceneChart, sourceLegend) && !sceneChart.Title.Layout.HasGeometryValues &&
                    !axisTitles.HasHorizontalTitle && !axisTitles.HasVerticalTitle &&
                    valueSceneAxis is { CrossesAt: null, CrossesKind: not PptxSceneChartAxisCrosses.Unknown }
                    // Numeric X bounds do not describe this crossing: it lies on
                    // the paired ordinal Y axis. Reuse its direction and edges.
                    ? ReadChartAxisPlacement(valueAxisSource.SceneAxis, new(0d, Math.Max(1, MaxDensePointCount(barSeriesVectors))),
                        barOptions.CategoryAxisReversed)
                    : null;
                if (numericPlacement is { } numeric)
                    axesStyle = axesStyle with { ValueAxisBottomSide = !numeric.AxisHighSide };
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneSeriesStrokes(barPlot, ChartFilledSeriesInheritedStrokeWidth);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills = ReadSceneSeriesPointFills(barPlot);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes = ReadSceneSeriesPointStrokes(barPlot);
                bool categoryLegend = barSeriesVectors.Count == 1 && sceneChart.Plots.Sum(plot => plot.Series.Count) == 1 &&
                    (barOptions.VaryColors.Value || barPlot?.Series.FirstOrDefault()?.PointStyles.Any(point => point.HasShapeProperties) == true);
                var legendEntries = new List<ChartLegendEntry>(categoryLegend
                    ? BuildBarCategoryLegendEntries(theme, colorMap, barPalette, barPlot, barChart, barSeriesVectors[0], barOptions, seriesFills, seriesStrokes, pointFills, pointStrokes, workbook, plotVisibleOnly)
                    : BuildFillLegendEntries(theme, colorMap, barPalette, barPlot, barChart, seriesFills, seriesStrokes, paletteOffset: 0, workbook: workbook));
                int? barStyleId = ReadChartStyleId(chartXml);
                int legendSeriesOffset = barSeriesCount;
                // Collect every painted plot before measuring rows. A prepared
                // legend box owns its complete entries, not a mutable prefix.
                foreach (PptxSceneChartPlot plot in barPlots.Skip(1))
                {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    XElement element = plot.Source;
                    int count = CountRenderableSeries(ReadSharedChartSeriesVectors(plot, workbook, plotVisibleOnly));
                    if (count == 0) continue;
                    var fills = ReadSceneSeriesFills(plot);
                    var strokes = ReadSceneSeriesStrokes(plot, ChartFilledSeriesInheritedStrokeWidth);
                    legendEntries.AddRange(BuildFillLegendEntries(theme, colorMap, WithSourceSeries(chartPalette, plot),
                        plot, element, fills, strokes, legendSeriesOffset, workbook));
                    legendSeriesOffset += count;
                }
                foreach (PptxSceneChartPlot plot in ReadSceneChartPlots(sceneChart, PptxSceneChartPlotKind.Line))
                {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    XElement element = plot.Source;
                    if (CountRenderableSeries(ReadSharedChartSeriesVectors(plot, workbook, plotVisibleOnly)) == 0) continue;
                    var options = ReadSceneChartLineOptions(sceneChart, plot, PptxSceneChartGrouping.Standard);
                    var strokes = ReadSceneSeriesStrokes(plot, ResolveLineKindInheritedStrokeWidth(barStyleId));
                    var markers = ReadSceneMarkerStyles(plot);
                    legendEntries.AddRange(BuildStrokeLegendEntries(theme, colorMap, chartPalette, plot, element,
                        strokes, markers, reverseOrder: options.Stacked, workbook: workbook, chartStyleId: barStyleId));
                }
                ChartLayout chartLayout = (await GetBarChartLayoutAsync(document, theme, bounds, chartXml, sceneChart, colorMap, barPlot, barChart, barOptions, workbook, plotVisibleOnly, fontResolver, resolvedCategory, sourceLegend, numericPlacement, context: context,
                    resolvedValueScale: sharedValueAxis || horizontalBars && CanOwnCartesianChartSourceRange(sceneChart, valueSceneAxis)
                        ? (valueExtents, valueAxisOptions.Units) : null).ConfigureAwait(false));
                ChartTextStyle legendTextStyle = ReadSceneChartLegendTextStyle(theme, sceneChart);
                ChartLegendBox? preparedLegendBox = null;
                double? measuredLegendAllocation = null;
                ChartLayout unreservedLayout = chartLayout;
                bool measuredHorizontalLegend = UsesMeasuredBarLegend(barOptions, sceneChart, chartLayout.Legend);
                bool manualPlotBottomLegend = chartLayout.ManualPlotLayoutApplied &&
                    chartLayout.Legend.PositionKind == PptxSceneChartLegendPosition.Bottom &&
                    !chartLayout.Legend.Layout.HasGeometryValues;
                // Outer captions consume their authored area, independently of
                // the legend's row container. Prepare it before plot admission
                // so a narrower data rectangle cannot truncate its names.
                if (categoryLegend || measuredHorizontalLegend || manualPlotBottomLegend || !horizontalBars &&
                    sceneChart.PlotAreaLayout.LayoutTargetKind == PptxSceneChartManualLayoutTarget.Outer)
                {
                    var prepared = (await LayoutChartLegendRowsAsync(chartLayout, legendEntries, legendTextStyle, fontResolver, context,
                        chartLayout.Frame.Width, measuredHorizontalLegend || manualPlotBottomLegend).ConfigureAwait(false));
                    chartLayout = prepared.Layout;
                    preparedLegendBox = prepared.Box;
                    measuredLegendAllocation = prepared.AllocationBand;
                }
                double legendBand = unreservedLayout.PlotBox.Height - chartLayout.PlotBox.Height;
                double? scaleLegendBand = legendBand;
                if (sceneChart.PlotAreaLayout.HasGeometryValues)
                    scaleLegendBand = !chartLayout.ManualPlotLayoutApplied ? null :
                        !chartLayout.Legend.Visible || chartLayout.Legend.Overlay ? 0d : measuredLegendAllocation;
                if (await MeasureAutomaticColumnScaleAsync(theme, sceneChart, chartXml, valueAxisSource, categoryAxis,
                    dataExtents, chartLayout, scaleLegendBand, fontResolver).ConfigureAwait(false) is { } scale &&
                    (scale.Extents != valueExtents || scale.Units != valueAxisOptions.Units))
                {
                    valueExtents = scale.Extents;
                    valueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(valueSceneAxis, valueAxis, theme,
                        valueExtents, percentStacked) with { Units = scale.Units };
                    resolvedCategory = ReadChartAxisPlacement(categoryAxis.SceneAxis, valueExtents, valueAxisOptions.Reversed);
                    if (resolvedCategory is { } scaledCategory)
                    {
                        valueAxisOptions = valueAxisOptions with { CrossingValue = scaledCategory.CrossingValue };
                        axesStyle = axesStyle with { CategoryAxisTopSide = scaledCategory.AxisHighSide };
                    }
                    // Bounds and units change numeric strip widths together. Reuse
                    // packed legend frames; only the dependent plot measurement reruns.
                    ChartLayout resized = await GetBarChartLayoutAsync(document, theme, bounds, chartXml, sceneChart,
                        colorMap, barPlot, barChart, barOptions, workbook, plotVisibleOnly, fontResolver,
                        resolvedCategory, sourceLegend, numericPlacement, context, scale).ConfigureAwait(false);
                    chartLayout = resized with
                    {
                        PlotBox = resized.PlotBox with { Y = resized.PlotBox.Y + chartLayout.PlotBox.Y - unreservedLayout.PlotBox.Y,
                            Height = resized.PlotBox.Height - legendBand },
                        PlotAreaBox = resized.PlotAreaBox with { Y = resized.PlotAreaBox.Y + chartLayout.PlotAreaBox.Y - unreservedLayout.PlotAreaBox.Y,
                            Height = Math.Max(1d, resized.PlotAreaBox.Height - legendBand) }
                    };
                }
                var categoryLabels = ReadSharedCategoryLabels(barPlot, workbook, plotVisibleOnly);
                var categoryPlacement = resolvedCategory is { } placed ? (placed.LabelPlotRatio, placed.LabelsHighSide) :
                    ResolveChartAxisLabelPlacement(categoryAxis.SceneAxis, categoryAxis.XmlAxis,
                        GetChartValuePlotRatio(valueExtents, valueAxisOptions.CrossingValue ?? 0d, valueAxisOptions.Reversed), valueAxisOptions.Reversed);
                if (!horizontalBars)
                    chartLayout = await PrepareCartesianCategoryFramesAsync(chartLayout, theme, sceneChart, chartXml,
                        valueAxisSource, valueExtents, valueAxisOptions,
                        GetValueAxisAutoTickTargetCount(false, true, chartLayout.ManualPlotLayoutApplied), axesStyle.ValueAxisRightSide,
                        categoryAxis, categoryLabels, categoryPlacement.LabelPlotRatio, categoryPlacement.LabelsHighSide,
                        fontResolver, context, preparedLegendBox, percentStacked ? "0%" : null).ConfigureAwait(false);
                RenderChartAreaStyle(graphics, document, bounds, sceneChart);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                // Gallery style 2 fills unstyled axis-family strokes on single-plot charts
                // (multi-plot dual-axis charts keep legacy output); explicit styles win by
                // construction since only null entries are replaced.
                if (barPlots.Count == 1 && ResolveGalleryAxisFamilyDefault(sceneChart.StyleId) is { } galleryAxisDefault)
                {
                    // SecondaryValueAxis stays untouched: its null-ness gates secondary-axis
                    // emission downstream, and filling it would conjure duplicate axis lines.
                    axesStyle = axesStyle with
                    {
                        ValueAxis = axesStyle.ValueAxis ?? galleryAxisDefault,
                        CategoryAxis = axesStyle.CategoryAxis ?? galleryAxisDefault,
                    };
                    valueAxisOptions = valueAxisOptions with
                    {
                        GridlineStyle = valueAxisOptions.GridlineStyle with { Major = valueAxisOptions.GridlineStyle.Major ?? galleryAxisDefault },
                    };
                }
                bool valueAxisLabelsVisible = IsSceneOrXmlChartAxisLabelVisible(valueSceneAxis, valueAxis);
                RenderBarChart(graphics, theme, colorMap, barPalette, chartLayout.PlotAreaBox, plotBox, barSeriesVectors, horizontalBars, barOptions, seriesFills, pointFills, pointStrokes, valueAxisOptions, axesStyle, plotAreaStyle, valueExtents, valueAxisLabelsVisible, chartLayout.ManualPlotLayoutApplied, barStyleId, resolvedCategory, context: context);
                XElement? secondaryValueAxis = null;
                PptxSceneChartAxis? secondaryValueSceneAxis = null;
                ChartValueExtents secondaryValueExtents = default;
                ChartAxisUnits secondaryAxisUnits = default;
                bool secondaryAxisReversed = false;
                foreach (PptxSceneChartPlot extraBarPlot in barPlots.Skip(1))
                {
                    XElement extraBarChart = extraBarPlot.Source;
                    IReadOnlyList<ChartIndexedNumberVector> extraSeriesVectors = ReadSharedChartSeriesVectors(extraBarPlot, workbook, plotVisibleOnly);
                    if (CountRenderableSeries(extraSeriesVectors) == 0)
                    {
                        continue;
                    }

                    ChartBarPlotOptions extraBarOptions = ReadSceneChartBarOptions(extraBarPlot, PptxSceneChartGrouping.Clustered);
                    ChartColorPalette extraPalette = WithSourceSeries(chartPalette, extraBarPlot);
                    bool extraHorizontalBars = extraBarOptions.BarDirection == PptxSceneChartBarDirection.Bar;
                    ChartAxisSource extraCategoryAxis = ReadChartCategoryAxisForPlot(sceneChart, extraBarPlot);
                    extraBarOptions = extraBarOptions with { CategoryAxisReversed = extraHorizontalBars &&
                        extraBarOptions.Grouping == PptxSceneChartGrouping.Clustered &&
                        ReadSceneValueAxisReversed(extraCategoryAxis.SceneAxis) };
                    ChartAxisSource extraValueAxisSource = ReadChartValueAxesForPlot(sceneChart, extraBarPlot).FirstOrDefault();
                    XElement? extraValueAxis = extraValueAxisSource.XmlAxis;
                    PptxSceneChartAxis? extraValueSceneAxis = extraValueAxisSource.SceneAxis;
                    bool extraPercentStacked = IsPercentStackedChartGrouping(extraBarOptions.Grouping);
                    ChartValueExtents extraValueExtents = sharedValueAxis ? valueExtents : ReadPercentStackedAwareValueAxisExtents(extraValueSceneAxis, extraValueAxis, GetSharedBarChartValueExtents(extraBarPlot, extraBarOptions.Grouping, workbook, plotVisibleOnly), extraPercentStacked, ResolveBarValueAxisHeadroom(extraHorizontalBars, extraPercentStacked), PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                    ChartValueAxisRenderOptions extraValueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(extraValueSceneAxis, extraValueAxis, theme, extraValueExtents, extraPercentStacked);
                    IReadOnlyList<ChartSeriesFill?> extraSeriesFills = ReadSceneSeriesFills(extraBarPlot);
                    IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> extraPointFills = ReadSceneSeriesPointFills(extraBarPlot);
                    IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> extraPointStrokes = ReadSceneSeriesPointStrokes(extraBarPlot);
                    if (!extraHorizontalBars && secondaryValueAxis is null &&
                        !IsSameChartAxis(valueAxisSource, extraValueAxisSource) &&
                        IsSceneOrXmlVisibleValueAxis(extraValueSceneAxis, extraValueAxis))
                    {
                        secondaryValueAxis = extraValueAxis;
                        secondaryValueSceneAxis = extraValueSceneAxis;
                        secondaryValueExtents = extraValueExtents;
                        secondaryAxisUnits = extraValueAxisOptions.Units;
                        secondaryAxisReversed = extraValueAxisOptions.Reversed;
                    }

                    RenderBarChart(
                        graphics,
                        theme,
                        colorMap,
                        extraPalette,
                        chartLayout.PlotAreaBox,
                        plotBox,
                        extraSeriesVectors,
                        extraHorizontalBars,
                        extraBarOptions,
                        extraSeriesFills,
                        extraPointFills,
                        extraPointStrokes,
                        extraValueAxisOptions with { MajorGridlines = false, MinorGridlines = false, GridlineStyle = ChartGridlineStyle.Empty },
                        axesStyle with { ValueAxisVisible = false, CategoryAxisVisible = false },
                        ChartShapeStyle.Empty,
                        extraValueExtents,
                        valueAxisLabelsVisible: false,
                        manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied,
                        barStyleId, categoryPlacement: sharedValueAxis && extraHorizontalBars ? resolvedCategory : null, context: context);
                    await RenderBarDataLabelsAsync(
                        theme,
                        colorMap,
                        graphics,
                        chartLayout.Frame,
                        plotBox,
                        extraSeriesVectors,
                        extraPalette,
                        extraValueExtents,
                        extraBarOptions,
                        extraValueAxisOptions,
                        extraSeriesFills,
                        extraPointFills,
                        ReadSceneDataLabelOptions(sceneChart, extraBarPlot),
                        ReadSceneSeriesDataLabelOptions(sceneChart, extraBarPlot),
                        ReadSceneCategoryLabelVector(extraBarPlot, workbook, plotVisibleOnly),
                        ReadSharedChartSeriesNames(extraBarPlot, workbook), fontResolver, fonts, context, sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds,
                        categoryPlacement: sharedValueAxis && extraHorizontalBars ? resolvedCategory : null).ConfigureAwait(false);
                }

                foreach (PptxSceneChartPlot linePlot in ReadSceneChartPlots(sceneChart, PptxSceneChartPlotKind.Line))
                {
                    XElement comboLineChart = linePlot.Source;
                    IReadOnlyList<ChartIndexedNumberVector> lineSeriesVectors = ReadSharedChartSeriesVectors(linePlot, workbook, plotVisibleOnly);
                    if (CountRenderableSeries(lineSeriesVectors) == 0)
                    {
                        continue;
                    }

                    ChartAxisSource lineValueAxisSource = ReadChartValueAxesForPlot(sceneChart, linePlot).FirstOrDefault();
                    XElement? lineValueAxis = lineValueAxisSource.XmlAxis;
                    XElement? lineValueAxisForScale = lineValueAxis ?? valueAxis;
                    PptxSceneChartAxis? lineValueSceneAxis = lineValueAxisSource.SceneAxis;
                    ChartLinePlotOptions lineOptions = ReadSceneChartLineOptions(sceneChart, linePlot, PptxSceneChartGrouping.Standard);
                    ChartValueExtents lineValueExtents = sharedValueAxis ? valueExtents : ReadPercentStackedAwareValueAxisExtents(lineValueSceneAxis, lineValueAxisForScale, GetSharedLineChartValueExtents(linePlot, lineOptions.Stacked, lineOptions.PercentStacked, workbook, plotVisibleOnly), lineOptions.PercentStacked, useNearMaximumHeadroom: !lineOptions.PercentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                    ChartValueAxisRenderOptions lineValueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(lineValueSceneAxis, lineValueAxisForScale, theme, lineValueExtents, lineOptions.PercentStacked);
                    IReadOnlyList<ChartSeriesStroke?> lineSeriesStrokes = ReadSceneSeriesStrokes(linePlot, ResolveLineKindInheritedStrokeWidth(barStyleId));
                    IReadOnlyList<ChartMarkerStyle> lineMarkerStyles = ReadSceneMarkerStyles(linePlot);
                    if (secondaryValueAxis is null && !IsSameChartAxis(valueAxisSource, lineValueAxisSource) &&
                        IsSceneOrXmlVisibleValueAxis(lineValueSceneAxis, lineValueAxis))
                    {
                        secondaryValueAxis = lineValueAxis;
                        secondaryValueSceneAxis = lineValueSceneAxis;
                        secondaryValueExtents = lineValueExtents;
                        secondaryAxisUnits = lineValueAxisOptions.Units;
                        secondaryAxisReversed = lineValueAxisOptions.Reversed;
                    }

                    ChartAxisSource comboCategoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
                    // Keep the independently qualified composed-chart defaults while
                    // sharing the painter's resolved appearance contract.
                    IReadOnlyList<ChartLineAppearance> lineAppearance = ReadLineSeriesAppearance(
                        linePlot.Series, lineOptions,
                        theme, colorMap, chartPalette, barStyleId, themeWidth: null,
                        pointColors: false, context.CancellationToken);
                    RenderLineChart(
                        graphics,
                        chartLayout.PlotAreaBox,
                        plotBox,
                        lineSeriesVectors,
                        lineOptions,
                        lineAppearance,
                        lineValueAxisOptions with { MajorGridlines = false, MinorGridlines = false, GridlineStyle = ChartGridlineStyle.Empty },
                        axesStyle with { ValueAxisVisible = false, CategoryAxisVisible = false },
                        ChartShapeStyle.Empty,
                        lineValueExtents,
                        categoryPointsOnTickMarks: false,
                        ReadSceneChartTextStyle(theme, sceneChart, comboCategoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis").FontSize, barStyleId, context: context);
                    await RenderCategoryPointDataLabelsAsync(
                        theme,
                        colorMap,
                        chartPalette,
                        graphics,
                        chartLayout.Frame,
                        plotBox,
                        lineSeriesVectors,
                        lineValueExtents,
                        lineValueAxisOptions.Reversed,
                        // Preserve the composed-chart label fallback until its
                        // secondary-axis/category contract is independently qualified.
                        categoryPointsOnTickMarks: true,
                        lineSeriesStrokes,
                        lineMarkerStyles,
                        ReadSceneDataLabelOptions(sceneChart, linePlot),
                        ReadSceneSeriesDataLabelOptions(sceneChart, linePlot),
                        ReadSceneCategoryLabelVector(linePlot, workbook, plotVisibleOnly),
                        ReadSharedChartSeriesNames(linePlot, workbook), fontResolver, fonts, context, sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds, chartRootFontSize: sceneChart.TextStyle.FontSize).ConfigureAwait(false);
                }

                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    double? categoryLabelCoordinate = horizontalBars
                        ? resolvedCategory is { } horizontalPlacement ? plotBox.X + plotBox.Width * horizontalPlacement.LabelPlotRatio : null
                        : plotBox.Y + plotBox.Height * categoryPlacement.LabelPlotRatio;
                    await RenderChartCategoryLabelsAsync(theme, graphics, chartLayout.Frame, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, categoryLabels, horizontalBars, categoryLabelCoordinate, categoryLabelsOnTickMarks: ResolveSceneCategoryAxisLabelsOnTickMarks(valueSceneAxis), labelsHighSide: (!horizontalBars || resolvedCategory is not null) && categoryPlacement.LabelsHighSide, fontResolver: fontResolver, chartFonts: fonts, categoryLabelOffsetAdjustment: chartLayout.CategoryLabelOffsetAdjustment, prepared: chartLayout.CategoryLabels, categoryAxisReversed: barOptions.CategoryAxisReversed).ConfigureAwait(false);
                }

                if (axesStyle.ValueAxisVisible)
                {
                    bool sameSideSecondaryValueAxis = !horizontalBars &&
                        secondaryValueAxis is not null &&
                        IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis) &&
                        GetValueAxisSideSlot(sceneChart,
                            valueSceneAxis,
                            valueAxis,
                            secondaryValueSceneAxis,
                            secondaryValueAxis,
                            defaultPrimaryRightSide: axesStyle.ValueAxisRightSide,
                            defaultSecondaryRightSide: axesStyle.SecondaryValueAxisRightSide) > 0;
                    if (valueAxisLabelsVisible)
                    {
                        await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, valueAxis, valueSceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars, rightSide: axesStyle.ValueAxisRightSide, axisSideSlot: 0, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied, useTextSizedWidth: sameSideSecondaryValueAxis, defaultNumberFormat: percentStacked ? "0%" : null, fontResolver: fontResolver, chartFonts: fonts, tickPolicy: valueAxisOptions.TickPolicy, placement: numericPlacement, context: context).ConfigureAwait(false);
                    }

                    if (!horizontalBars)
                    {
                        if (secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                        {
                            bool secondaryValueAxisRightSide = axesStyle.SecondaryValueAxisRightSide;
                            int sideSlot = GetValueAxisSideSlot(sceneChart, valueSceneAxis, valueAxis, secondaryValueSceneAxis, secondaryValueAxis, defaultPrimaryRightSide: axesStyle.ValueAxisRightSide, defaultSecondaryRightSide: secondaryValueAxisRightSide);
                            double? secondaryInkAnchor = null;
                            bool secondaryLabelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(sceneChart, secondaryValueSceneAxis, secondaryValueAxis, secondaryValueAxisRightSide);
                            if (!secondaryLabelsRightSide && sideSlot == 1 && valueAxisLabelsVisible)
                            {
                                ChartValueAxisStripMeasure primaryStrip = (await MeasureVerticalValueAxisLabelStripAsync(theme, sceneChart, chartXml, valueAxis, valueSceneAxis, valueExtents, valueAxisOptions.Units, defaultNumberFormat: percentStacked ? "0%" : null, fontResolver: fontResolver, context: context).ConfigureAwait(false));
                                secondaryInkAnchor = plotBox.X - primaryStrip.SideGap - primaryStrip.MaxLabelWidth;
                            }
                            await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, secondaryAxisReversed, horizontalBars: false, rightSide: secondaryValueAxisRightSide, axisSideSlot: sideSlot, useTextSizedWidth: sideSlot > 0, manualPlotLayoutApplied: false, defaultNumberFormat: null, fontResolver: fontResolver, chartFonts: fonts, innerStripInkEdge: secondaryInkAnchor, context: context).ConfigureAwait(false);
                        }
                        else
                        {
                            await RenderSecondaryChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, GetSharedBarChartValueExtents(barPlot, barOptions.Grouping, workbook, plotVisibleOnly), fontResolver, chartFonts: fonts, primarySceneAxis: valueSceneAxis, primaryXmlAxis: valueAxis, context: context).ConfigureAwait(false);
                        }
                    }
                }
                else if (!horizontalBars && secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                {
                    await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, secondaryAxisReversed, horizontalBars: false, rightSide: axesStyle.SecondaryValueAxisRightSide, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver, chartFonts: fonts, context: context).ConfigureAwait(false);
                }
                // Major tick marks for vertical bars (Office honors out/in/cross at category
                // slot boundaries and value gridlines; horizontal bars keep legacy output
                // for lack of tick evidence).
                if (!horizontalBars)
                {
                    ChartAxisSource tickCategoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
                    PptxSceneChartAxisTickMark categoryTickMark = ReadSceneChartAxisMajorTickMark(tickCategoryAxis.SceneAxis);
                    if (axesStyle.CategoryAxisVisible && categoryTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke categoryTickStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
                        if (categoryTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle categoryTickStyle = ReadSceneChartTextStyle(theme, sceneChart, tickCategoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                            int tickCategoryCount = Math.Max(1, MaxDensePointCount(barSeriesVectors));
                            double categoryAxisY = ChartValueToPlotCoordinate(valueExtents, valueAxisOptions.CrossingValue, plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                            SetChartStroke(graphics, categoryTickStroke);
                            DrawCartesianChartCategoryAxisMajorTicks(graphics, plotBox.X, plotBox.Width,
                                tickCategoryCount, categoryAxisY, categoryTickMark, categoryTickStyle.FontSize,
                                barOptions.CategoryPointsOnTickMarks, axesStyle.CategoryAxisTopSide);
                        }
                    }

                    PptxSceneChartAxisTickMark valueTickMark = ReadSceneChartAxisMajorTickMark(valueSceneAxis);
                    if (axesStyle.ValueAxisVisible && valueTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke valueTickStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
                        if (valueTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle valueTickStyle = ReadSceneChartTextStyle(theme, sceneChart, valueSceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                            IReadOnlyList<double> tickValues = GetChartAxisTickValuesCore(valueExtents, valueAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: false, valueAxisLabelsVisible: valueAxisLabelsVisible, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied), valueAxisOptions.TickPolicy, context: context);
                            double[] tickEdges = new double[tickValues.Count];
                            for (int tickIndex = 0; tickIndex < tickValues.Count; tickIndex++)
                            {
                                tickEdges[tickIndex] = ChartValueToPlotCoordinate(valueExtents, tickValues[tickIndex], plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                            }

                            bool valueLabelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(sceneChart, valueSceneAxis, valueAxis, axesStyle.ValueAxisRightSide);
                            double valueAxisX = axesStyle.ValueAxisRightSide ? plotBox.X + plotBox.Width : plotBox.X;
                            SetChartStroke(graphics, valueTickStroke);
                            StrokeMajorTickSegments(graphics, tickEdges, valueAxisX, verticalSegments: false, outwardIsLowSide: !valueLabelsRightSide, valueTickMark, valueTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                        }
                    }
                }
                // Major tick marks for horizontal bars (Office honors out/in/cross at value
                // gridlines and category slot boundaries; lengths share the 0.315fs law).
                if (horizontalBars)
                {
                    ChartAxisSource tickValueAxis = ReadChartValueAxesForPlot(sceneChart, barPlot).FirstOrDefault();
                    PptxSceneChartAxisTickMark horizontalValueTickMark = ReadSceneChartAxisMajorTickMark(tickValueAxis.SceneAxis);
                    if (axesStyle.CategoryAxisVisible && horizontalValueTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke horizontalValueTickStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
                        if (horizontalValueTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle horizontalValueTickStyle = ReadSceneChartTextStyle(theme, sceneChart, tickValueAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                            IReadOnlyList<double> horizontalTickValues = GetChartAxisTickValuesCore(valueExtents, valueAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: true, valueAxisLabelsVisible: valueAxisLabelsVisible, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied), valueAxisOptions.TickPolicy, context: context);
                            double[] horizontalTickEdges = new double[horizontalTickValues.Count];
                            for (int horizontalTickIndex = 0; horizontalTickIndex < horizontalTickValues.Count; horizontalTickIndex++)
                            {
                                horizontalTickEdges[horizontalTickIndex] = ChartValueToPlotCoordinate(valueExtents, horizontalTickValues[horizontalTickIndex], plotBox.X, plotBox.Width, valueAxisOptions.Reversed);
                            }

                            double horizontalValueAxisY = axesStyle.ValueAxisBottomSide ? plotBox.Y : plotBox.Y + plotBox.Height;
                            SetChartStroke(graphics, horizontalValueTickStroke);
                            StrokeMajorTickSegments(graphics, horizontalTickEdges, horizontalValueAxisY, verticalSegments: true, outwardIsLowSide: axesStyle.ValueAxisBottomSide, horizontalValueTickMark, horizontalValueTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                        }
                    }

                    ChartAxisSource tickCategoryAxis = ReadChartCategoryAxisForPlot(sceneChart, barPlot);
                    PptxSceneChartAxisTickMark horizontalCategoryTickMark = ReadSceneChartAxisMajorTickMark(tickCategoryAxis.SceneAxis);
                    if (axesStyle.ValueAxisVisible && horizontalCategoryTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke horizontalCategoryTickStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
                        if (horizontalCategoryTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle horizontalCategoryTickStyle = ReadSceneChartTextStyle(theme, sceneChart, tickCategoryAxis.SceneAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                            int horizontalCategoryCount = Math.Max(1, MaxDensePointCount(barSeriesVectors));
                            double horizontalSlotHeight = plotBox.Height / horizontalCategoryCount;
                            double[] horizontalCategoryEdges = new double[horizontalCategoryCount + 1];
                            for (int horizontalCategoryIndex = 0; horizontalCategoryIndex <= horizontalCategoryCount; horizontalCategoryIndex++)
                            {
                                horizontalCategoryEdges[horizontalCategoryIndex] = plotBox.Y + horizontalSlotHeight * horizontalCategoryIndex;
                            }

                            double horizontalCategoryAxisX = resolvedCategory is { } categoryTickPlacement
                                ? ChartValueToPlotCoordinate(valueExtents, categoryTickPlacement.CrossingValue, plotBox.X, plotBox.Width, valueAxisOptions.Reversed)
                                : axesStyle.CategoryAxisRightSide ? plotBox.X + plotBox.Width : plotBox.X;
                            SetChartStroke(graphics, horizontalCategoryTickStroke);
                            StrokeMajorTickSegments(graphics, horizontalCategoryEdges, horizontalCategoryAxisX, verticalSegments: false, outwardIsLowSide: !axesStyle.CategoryAxisRightSide, horizontalCategoryTickMark, horizontalCategoryTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                        }
                    }
                }
                await RenderDefaultChartAxisTitlesAsync(theme, graphics, chartLayout, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                await RenderChartLegendAsync(graphics, chartLayout.Frame, plotBox, legendEntries, chartLayout.Legend, legendTextStyle, fontResolver, ChartLegendPlacement.Default, chartFonts: fonts, preparedBox: preparedLegendBox).ConfigureAwait(false);
                await RenderBarDataLabelsAsync(
                    theme,
                    colorMap,
                    graphics,
                    chartLayout.Frame,
                    plotBox,
                    barSeriesVectors,
                    barPalette,
                    valueExtents,
                    barOptions,
                    valueAxisOptions,
                    seriesFills,
                    pointFills,
                    ReadSceneDataLabelOptions(sceneChart, barPlot),
                    ReadSceneSeriesDataLabelOptions(sceneChart, barPlot),
                    ReadSceneCategoryLabelVector(barPlot, workbook, plotVisibleOnly),
                    ReadSharedChartSeriesNames(barPlot, workbook), fontResolver, fonts, context, sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds, categoryPlacement: resolvedCategory).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }
}
