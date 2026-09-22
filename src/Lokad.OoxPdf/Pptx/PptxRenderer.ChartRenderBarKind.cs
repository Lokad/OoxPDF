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
    private static bool TryRenderBarChartKind(
        PdfGraphicsBuilder graphics,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        IReadOnlyList<RgbColor>? chartPalette,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        ChartWorkbookData? workbook,
        List<PdfFontResource> fonts,
        PresentationFontResolver fontResolver,
        bool plotVisibleOnly,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        IReadOnlyList<XElement> barCharts = ReadSceneOrXmlChartPlotElements(sceneChart, chartXml, PptxSceneChartPlotKind.Bar);
        XElement? barChart = barCharts.FirstOrDefault();
        if (barChart is not null)
        {
            PptxSceneChartPlot? barPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Bar, 0);
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
                ChartValueExtents valueExtents = ReadPercentStackedAwareValueAxisExtents(valueSceneAxis, valueAxis, GetBarChartValueExtents(barSeriesVectors, barOptions.Grouping), percentStacked, ResolveBarValueAxisHeadroom(horizontalBars, percentStacked), PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartValueAxisRenderOptions valueAxisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(valueSceneAxis, valueAxis, theme, valueExtents, percentStacked);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(barPlot, barChart, theme, colorMap, ChartFilledSeriesInheritedStrokeWidth);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills = ReadSceneOrXmlSeriesPointFills(barPlot, barChart, theme, colorMap);
                IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes = ReadSceneOrXmlSeriesPointStrokes(barPlot, barChart, theme, colorMap);
                var legendEntries = new List<ChartLegendEntry>(BuildFillLegendEntries(theme, colorMap, chartPalette, barPlot, barChart, seriesFills, seriesStrokes, paletteOffset: 0, workbook: workbook));
                ChartLayout chartLayout = GetBarChartLayout(document, theme, bounds, chartXml, sceneChart, colorMap, barPlot, barChart, barOptions, workbook, plotVisibleOnly, fontResolver);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPlotBox plotBox = chartLayout.PlotBox;
                // Gallery style 2 fills unstyled axis-family strokes on single-plot charts
                // (multi-plot dual-axis charts keep legacy output); explicit styles win by
                // construction since only null entries are replaced.
                if (barCharts.Count == 1 && ResolveGalleryAxisFamilyDefault(sceneChart?.StyleId, sceneChart?.StylePart.IsDefined == true) is { } galleryAxisDefault)
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
                RenderBarChart(graphics, theme, colorMap, chartPalette, chartLayout.PlotAreaBox, plotBox, barSeriesVectors, horizontalBars, barOptions, seriesFills, pointFills, pointStrokes, valueAxisOptions, axesStyle, plotAreaStyle, valueExtents, valueAxisLabelsVisible, chartLayout.ManualPlotLayoutApplied);
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
                    ChartValueExtents extraValueExtents = ReadPercentStackedAwareValueAxisExtents(extraValueSceneAxis, extraValueAxis, GetBarChartValueExtents(extraSeriesVectors, extraBarOptions.Grouping), extraPercentStacked, ResolveBarValueAxisHeadroom(extraHorizontalBars, extraPercentStacked), PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
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
                        extraBarOptions,
                        extraSeriesFills,
                        extraPointFills,
                        extraPointStrokes,
                        extraValueAxisOptions with { MajorGridlines = false, MinorGridlines = false, GridlineStyle = ChartGridlineStyle.Empty },
                        axesStyle with { ValueAxisVisible = false, CategoryAxisVisible = false },
                        ChartShapeStyle.Empty,
                        extraValueExtents,
                        valueAxisLabelsVisible: false,
                        manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied);
                    RenderBarDataLabels(
                        theme,
                        colorMap,
                        graphics,
                        plotBox,
                        extraSeriesVectors,
                        chartPalette,
                        extraValueExtents,
                        extraBarOptions,
                        extraValueAxisOptions,
                        extraSeriesFills,
                        extraPointFills,
                        ReadSceneOrXmlDataLabelOptions(sceneChart, extraBarPlot, extraBarChart, theme, colorMap),
                        ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, extraBarPlot, extraBarChart, theme, colorMap),
                        ReadSceneOrXmlCategoryLabelVector(extraBarPlot, extraBarChart, workbook, plotVisibleOnly),
                        ReadSharedChartSeriesNames(extraBarPlot, extraBarChart, workbook), fontResolver, fonts, context, sceneChart?.Relationships, linkAnnotations, reportedHyperlinkIds);
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
                    ChartValueExtents lineValueExtents = ReadPercentStackedAwareValueAxisExtents(lineValueSceneAxis, lineValueAxisForScale, GetLineChartValueExtents(lineSeriesVectors, lineOptions.Stacked, lineOptions.PercentStacked), lineOptions.PercentStacked, useNearMaximumHeadroom: !lineOptions.PercentStacked, nearMaximumHeadroomRatio: PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
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
                    ChartAxisSource comboCategoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
                    RenderLineChart(
                        graphics,
                        theme,
                        colorMap,
                        chartPalette,
                        chartLayout.PlotAreaBox,
                        plotBox,
                        lineSeriesVectors,
                        lineOptions,
                        lineSeriesStrokes,
                        lineMarkerStyles,
                        lineValueAxisOptions with { MajorGridlines = false, MinorGridlines = false, GridlineStyle = ChartGridlineStyle.Empty },
                        axesStyle with { ValueAxisVisible = false, CategoryAxisVisible = false },
                        ChartShapeStyle.Empty,
                        lineValueExtents,
                        ReadSceneOrXmlChartTextStyle(theme, sceneChart, comboCategoryAxis.SceneAxis, chartXml, comboCategoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis").FontSize);
                    RenderLineDataLabels(
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
                        ReadSharedChartSeriesNames(linePlot, comboLineChart, workbook), fontResolver, fonts, context, sceneChart?.Relationships, linkAnnotations, reportedHyperlinkIds);
                    lineChartIndex++;
                }

                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
                if (axesStyle.CategoryAxisVisible && IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    double? categoryLabelAxisY = horizontalBars
                        ? null
                        : (axesStyle.CategoryAxisTopSide ? plotBox.Y + plotBox.Height : ChartValueToPlotCoordinate(valueExtents, valueAxisOptions.CrossingValue, plotBox.Y, plotBox.Height, valueAxisOptions.Reversed));
                    RenderChartCategoryLabels(document, theme, graphics, plotBox, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSharedCategoryLabels(barPlot, barChart, workbook, plotVisibleOnly), horizontalBars, categoryLabelAxisY, categoryLabelsOnTickMarks: ResolveSceneOrXmlCategoryAxisLabelsOnTickMarks(valueSceneAxis, valueAxis), categoryLabelsTopSide: ResolveSceneOrXmlCategoryAxisTopSide(categoryAxis.SceneAxis, categoryAxis.XmlAxis, defaultTopSide: false), fontResolver: fontResolver, chartFonts: fonts);
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
                        RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, valueAxis, valueSceneAxis, valueExtents, valueAxisOptions.Units, valueAxisOptions.Reversed, horizontalBars, rightSide: axesStyle.ValueAxisRightSide, axisSideSlot: 0, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied, useTextSizedWidth: sameSideSecondaryValueAxis, defaultNumberFormat: percentStacked ? "0%" : null, fontResolver: fontResolver, chartFonts: fonts);
                    }

                    if (!horizontalBars)
                    {
                        if (secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                        {
                            bool secondaryValueAxisRightSide = ResolveSceneOrXmlValueAxisRightSide(secondaryValueSceneAxis, secondaryValueAxis, axesStyle.SecondaryValueAxisRightSide);
                            int sideSlot = GetValueAxisSideSlot(valueSceneAxis, valueAxis, secondaryValueSceneAxis, secondaryValueAxis, defaultPrimaryRightSide: axesStyle.ValueAxisRightSide, defaultSecondaryRightSide: secondaryValueAxisRightSide);
                            double? secondaryInkAnchor = null;
                            bool secondaryLabelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(secondaryValueSceneAxis, secondaryValueAxis, secondaryValueAxisRightSide);
                            if (!secondaryLabelsRightSide && sideSlot == 1 && valueAxisLabelsVisible)
                            {
                                ChartValueAxisStripMeasure primaryStrip = MeasureVerticalValueAxisLabelStrip(theme, sceneChart, chartXml, valueAxis, valueSceneAxis, valueExtents, valueAxisOptions.Units, defaultNumberFormat: percentStacked ? "0%" : null, fontResolver: fontResolver);
                                double primarySideGap = Math.Max(3d, primaryStrip.FontSize * PptxChartMetricRules.ValueAxisLabelSideGapFactor);
                                secondaryInkAnchor = plotBox.X - primarySideGap - primaryStrip.MaxLabelWidth;
                            }
                            RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, secondaryAxisReversed, horizontalBars: false, rightSide: secondaryValueAxisRightSide, axisSideSlot: sideSlot, useTextSizedWidth: sideSlot > 0, manualPlotLayoutApplied: false, defaultNumberFormat: null, fontResolver: fontResolver, chartFonts: fonts, innerStripInkEdge: secondaryInkAnchor);
                        }
                        else
                        {
                            RenderSecondaryChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, GetBarChartValueExtents(barSeriesVectors, barOptions.Grouping), fontResolver, chartFonts: fonts, primarySceneAxis: valueSceneAxis, primaryXmlAxis: valueAxis);
                        }
                    }
                }
                else if (!horizontalBars && secondaryValueAxis is not null && IsSceneOrXmlChartAxisLabelVisible(secondaryValueSceneAxis, secondaryValueAxis))
                {
                    RenderChartValueAxisLabels(document, theme, graphics, plotBox, chartXml, sceneChart, secondaryValueAxis, secondaryValueSceneAxis, secondaryValueExtents, secondaryAxisUnits, secondaryAxisReversed, horizontalBars: false, rightSide: ResolveSceneOrXmlValueAxisRightSide(secondaryValueSceneAxis, secondaryValueAxis, axesStyle.SecondaryValueAxisRightSide), axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver, chartFonts: fonts);
                }
                // Major tick marks for vertical bars (Office honors out/in/cross at category
                // slot boundaries and value gridlines; horizontal bars keep legacy output
                // for lack of tick evidence).
                if (!horizontalBars)
                {
                    ChartAxisSource tickCategoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
                    PptxSceneChartAxisTickMark categoryTickMark = ReadSceneOrXmlChartAxisMajorTickMark(tickCategoryAxis.SceneAxis, tickCategoryAxis.XmlAxis);
                    if (axesStyle.CategoryAxisVisible && categoryTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke categoryTickStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
                        if (categoryTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle categoryTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, tickCategoryAxis.SceneAxis, chartXml, tickCategoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                            int tickCategoryCount = Math.Max(1, DensifyChartPointSeries(barSeriesVectors).Max(values => values.Count));
                            double tickSlotWidth = plotBox.Width / tickCategoryCount;
                            double[] tickEdges = new double[tickCategoryCount + 1];
                            for (int tickIndex = 0; tickIndex <= tickCategoryCount; tickIndex++)
                            {
                                tickEdges[tickIndex] = plotBox.X + tickSlotWidth * tickIndex;
                            }

                            double categoryAxisY = ChartValueToPlotCoordinate(valueExtents, valueAxisOptions.CrossingValue, plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                            SetChartStroke(graphics, categoryTickStroke);
                            StrokeMajorTickSegments(graphics, tickEdges, categoryAxisY, verticalSegments: true, outwardIsLowSide: !axesStyle.CategoryAxisTopSide, categoryTickMark, categoryTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                        }
                    }

                    PptxSceneChartAxisTickMark valueTickMark = ReadSceneOrXmlChartAxisMajorTickMark(valueSceneAxis, valueAxis);
                    if (axesStyle.ValueAxisVisible && valueTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke valueTickStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
                        if (valueTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle valueTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, valueSceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                            IReadOnlyList<double> tickValues = GetChartAxisTickValues(valueExtents, valueAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: false, valueAxisLabelsVisible: valueAxisLabelsVisible, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied));
                            double[] tickEdges = new double[tickValues.Count];
                            for (int tickIndex = 0; tickIndex < tickValues.Count; tickIndex++)
                            {
                                tickEdges[tickIndex] = ChartValueToPlotCoordinate(valueExtents, tickValues[tickIndex], plotBox.Y, plotBox.Height, valueAxisOptions.Reversed);
                            }

                            bool valueLabelsRightSide = ResolveSceneOrXmlValueAxisLabelsRightSide(valueSceneAxis, valueAxis, axesStyle.ValueAxisRightSide);
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
                    ChartAxisSource tickValueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, barPlot, chartXml, barChart).FirstOrDefault();
                    PptxSceneChartAxisTickMark horizontalValueTickMark = ReadSceneOrXmlChartAxisMajorTickMark(tickValueAxis.SceneAxis, tickValueAxis.XmlAxis);
                    if (axesStyle.CategoryAxisVisible && horizontalValueTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke horizontalValueTickStroke = axesStyle.ValueAxis ?? ChartAxisDefaultStroke;
                        if (horizontalValueTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle horizontalValueTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, tickValueAxis.SceneAxis, chartXml, tickValueAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
                            IReadOnlyList<double> horizontalTickValues = GetChartAxisTickValues(valueExtents, valueAxisOptions.Units.MajorUnit, includeEndpoints: true, GetValueAxisAutoTickTargetCount(horizontalBars: true, valueAxisLabelsVisible: valueAxisLabelsVisible, manualPlotLayoutApplied: chartLayout.ManualPlotLayoutApplied));
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

                    ChartAxisSource tickCategoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, barPlot, chartXml, barChart);
                    PptxSceneChartAxisTickMark horizontalCategoryTickMark = ReadSceneOrXmlChartAxisMajorTickMark(tickCategoryAxis.SceneAxis, tickCategoryAxis.XmlAxis);
                    if (axesStyle.ValueAxisVisible && horizontalCategoryTickMark != PptxSceneChartAxisTickMark.None)
                    {
                        ChartSeriesStroke horizontalCategoryTickStroke = axesStyle.CategoryAxis ?? ChartAxisDefaultStroke;
                        if (horizontalCategoryTickStroke.Alpha > 0.001d)
                        {
                            ChartTextStyle horizontalCategoryTickStyle = ReadSceneOrXmlChartTextStyle(theme, sceneChart, tickCategoryAxis.SceneAxis, chartXml, tickCategoryAxis.XmlAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
                            int horizontalCategoryCount = Math.Max(1, DensifyChartPointSeries(barSeriesVectors).Max(values => values.Count));
                            double horizontalSlotHeight = plotBox.Height / horizontalCategoryCount;
                            double[] horizontalCategoryEdges = new double[horizontalCategoryCount + 1];
                            for (int horizontalCategoryIndex = 0; horizontalCategoryIndex <= horizontalCategoryCount; horizontalCategoryIndex++)
                            {
                                horizontalCategoryEdges[horizontalCategoryIndex] = plotBox.Y + horizontalSlotHeight * horizontalCategoryIndex;
                            }

                            double horizontalCategoryAxisX = axesStyle.CategoryAxisRightSide ? plotBox.X + plotBox.Width : plotBox.X;
                            SetChartStroke(graphics, horizontalCategoryTickStroke);
                            StrokeMajorTickSegments(graphics, horizontalCategoryEdges, horizontalCategoryAxisX, verticalSegments: false, outwardIsLowSide: !axesStyle.CategoryAxisRightSide, horizontalCategoryTickMark, horizontalCategoryTickStyle.FontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
                        }
                    }
                }
                RenderDefaultChartAxisTitles(theme, colorMap, graphics, chartLayout, chartXml, sceneChart, fontResolver, fonts, context, linkAnnotations, reportedHyperlinkIds);
                RenderChartLegend(graphics, chartLayout.Frame, plotBox, legendEntries, chartLayout.Legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default, chartFonts: fonts);
                RenderBarDataLabels(
                    theme,
                    colorMap,
                    graphics,
                    plotBox,
                    barSeriesVectors,
                    chartPalette,
                    valueExtents,
                    barOptions,
                    valueAxisOptions,
                    seriesFills,
                    pointFills,
                    ReadSceneOrXmlDataLabelOptions(sceneChart, barPlot, barChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, barPlot, barChart, theme, colorMap),
                    ReadSceneOrXmlCategoryLabelVector(barPlot, barChart, workbook, plotVisibleOnly),
                    ReadSharedChartSeriesNames(barPlot, barChart, workbook), fontResolver, fonts, context, sceneChart?.Relationships, linkAnnotations, reportedHyperlinkIds);
                return true;
            }
        }
        return false;
    }
}
