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
        PptxSceneChartGrouping grouping,
        double gapWidthPercent,
        IReadOnlyList<ChartSeriesFill?> seriesFills,
        IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills,
        bool varyColors,
        ChartDataLabelOptions labelOptions,
        IReadOnlyList<ChartDataLabelOptions> seriesLabelOptions,
        ChartIndexedTextVector categoryLabels,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames,
        PresentationFontResolver? fontResolver)
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
        bool stacked = IsStackedChartGrouping(grouping);
        bool percentStacked = IsPercentStackedChartGrouping(grouping);
        if (horizontalBars)
        {
            double categoryHeight = plotBox.Height / categoryCount;
            double barSlot = stacked
                ? GetStackedBarWidth(categoryHeight, gapWidthPercent)
                : categoryHeight * PptxChartMetricRules.BarDataLabelSlotFillRatio / Math.Max(1, densePointSeries.Count);
            double labelWidth = Math.Max(PptxChartMetricRules.CartesianDataLabelMinimumWidth, plotBox.Width * PptxChartMetricRules.HorizontalBarDataLabelWidthRatio);
            double[] positiveValues = new double[categoryCount];
            double[] negativeValues = new double[categoryCount];
            double[] positiveTotals = stacked ? GetCategoryPositiveTotals(densePointSeries, categoryCount, percentStacked) : [];
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryY = stacked
                    ? plotBox.Y + category * categoryHeight + (categoryHeight - barSlot) / 2d
                    : plotBox.Y + category * categoryHeight + categoryHeight * PptxChartMetricRules.BarDataLabelCategoryInsetRatio;
                for (int seriesIndex = 0; seriesIndex < densePointSeries.Count; seriesIndex++)
                {
                    IReadOnlyList<ChartIndexedNumberPoint?> points = densePointSeries[seriesIndex];
                    if (category >= points.Count || points[category]?.Value is not double value)
                    {
                        continue;
                    }

                    double normalizedValue = stacked ? NormalizeStackedValue(value, positiveTotals[category], percentStacked) : value;
                    (double barBase, double barEnd) = stacked
                        ? ResolveStackedBarSegmentValues(positiveValues, negativeValues, category, normalizedValue)
                        : (0d, value);
                    double barBaseX = stacked
                        ? ChartValueToPlotCoordinate(extents, barBase, plotBox.X, plotBox.Width, valueAxisReversed)
                        : zeroX;
                    double barEndX = ChartValueToPlotCoordinate(extents, barEnd, plotBox.X, plotBox.Width, valueAxisReversed);
                    ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), category);
                    if (!effectiveOptions.HasVisibleContent)
                    {
                        continue;
                    }

                    ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                    double fontSize = style.FontSize;
                    double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                    PptxSceneChartDataLabelPosition labelPosition = ResolveStackedBarDataLabelPosition(effectiveOptions.PositionKind, stacked);
                    double x = ResolveHorizontalBarDataLabelX(labelPosition, barBaseX, barEndX, labelWidth);
                    double y = categoryY + (stacked ? (barSlot - labelHeight) / 2d : seriesIndex * barSlot + barSlot * PptxChartMetricRules.HorizontalBarDataLabelSlotCenterRatio - labelHeight / 2d);
                    ChartIndexedNumberPoint point = points[category] ?? default;
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
            double barSlot = stacked
                ? GetStackedBarWidth(categoryWidth, gapWidthPercent)
                : categoryWidth * PptxChartMetricRules.BarDataLabelSlotFillRatio / Math.Max(1, densePointSeries.Count);
            double[] positiveValues = new double[categoryCount];
            double[] negativeValues = new double[categoryCount];
            double[] positiveTotals = stacked ? GetCategoryPositiveTotals(densePointSeries, categoryCount, percentStacked) : [];
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryX = stacked
                    ? plotBox.X + category * categoryWidth + (categoryWidth - barSlot) / 2d
                    : plotBox.X + category * categoryWidth + categoryWidth * PptxChartMetricRules.BarDataLabelCategoryInsetRatio;
                for (int seriesIndex = 0; seriesIndex < densePointSeries.Count; seriesIndex++)
                {
                    IReadOnlyList<ChartIndexedNumberPoint?> points = densePointSeries[seriesIndex];
                    if (category >= points.Count || points[category]?.Value is not double value)
                    {
                        continue;
                    }

                    double normalizedValue = stacked ? NormalizeStackedValue(value, positiveTotals[category], percentStacked) : value;
                    (double barBase, double barEnd) = stacked
                        ? ResolveStackedBarSegmentValues(positiveValues, negativeValues, category, normalizedValue)
                        : (0d, value);
                    double x = categoryX + (stacked ? 0d : seriesIndex * barSlot);
                    double barBaseY = stacked
                        ? ChartValueToPlotCoordinate(extents, barBase, plotBox.Y, plotBox.Height, valueAxisReversed)
                        : zeroY;
                    double barEndY = ChartValueToPlotCoordinate(extents, barEnd, plotBox.Y, plotBox.Height, valueAxisReversed);
                    ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), category);
                    if (!effectiveOptions.HasVisibleContent)
                    {
                        continue;
                    }

                    ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                    double fontSize = style.FontSize;
                    double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                    PptxSceneChartDataLabelPosition labelPosition = ResolveStackedBarDataLabelPosition(effectiveOptions.PositionKind, stacked);
                    double y = ResolveVerticalBarDataLabelY(labelPosition, barBaseY, barEndY, labelHeight);
                    ChartIndexedNumberPoint point = points[category] ?? default;
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
}
