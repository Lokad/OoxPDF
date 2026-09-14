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
    // Swatch+text unit left for clustered vertical-bar legend-key labels (unit centered
    // on the bar middle plus the calibrated offset).
    private static double ComputeBarLegendKeyUnitLeft(double barCenterX, double swatchSize, double textWidth)
    {
        return barCenterX + PptxChartMetricRules.BarLegendKeyUnitCenterOffset - (swatchSize + swatchSize + textWidth) / 2d;
    }

    // Out-end baseline gap for clustered vertical-bar legend-key labels
    // (Office baselines barTop plus 6.74 at 8pt, plus 8.98 at 16pt).
    private static double ComputeBarLegendKeyOutEndGap(double fontSize)
    {
        return PptxChartMetricRules.BarLegendKeyOutEndBoxPad +
            fontSize * PptxChartMetricRules.BarLegendKeyOutEndGapFontFactor +
            PptxChartMetricRules.BarLegendKeyOutEndGapConstant;
    }

    // Swatch gap for legend-key labels (Office 4.50 at 8pt, 6.63 at 16pt).
    private static double ComputeBarLegendKeySwatchGap(double fontSize)
    {
        return fontSize * PptxChartMetricRules.BarLegendKeySwatchGapFactor +
            PptxChartMetricRules.BarLegendKeySwatchGapConstant;
    }

    // Swatch left from text left (Office swatchX equals textX minus G minus S).
    private static double ComputeBarLegendKeySwatchX(double textX, double fontSize)
    {
        double swatchSize = fontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
        return textX - ComputeBarLegendKeySwatchGap(fontSize) - swatchSize;
    }

    // Swatch top from baseline (Office centers baseline plus 0.34fs).
    private static double ComputeBarLegendKeySwatchY(double baselineY, double fontSize)
    {
        double swatchSize = fontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
        return baselineY + fontSize * PptxChartMetricRules.BarLegendKeySwatchCenterOffsetFactor - swatchSize / 2d;
    }

    private static void RenderBarDataLabels(
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<ChartIndexedNumberVector> series,
        IReadOnlyList<RgbColor>? chartPalette,
        ChartValueExtents extents,
        ChartBarPlotOptions barOptions,
        ChartValueAxisRenderOptions valueAxisOptions,
        IReadOnlyList<ChartSeriesFill?> seriesFills,
        IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills,
        ChartDataLabelOptions labelOptions,
        IReadOnlyList<ChartDataLabelOptions> seriesLabelOptions,
        ChartIndexedTextVector categoryLabels,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames,
        PresentationFontResolver? fontResolver,
        List<PdfFontResource> chartFonts,
        PptxRenderContext context,
        IReadOnlyDictionary<string, OoxRelationship>? chartRelationships,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        bool horizontalBars = barOptions.BarDirection == PptxSceneChartBarDirection.Bar;
        bool valueAxisReversed = valueAxisOptions.Reversed;
        PptxSceneChartGrouping grouping = barOptions.Grouping;
        double gapWidthPercent = barOptions.GapWidth;
        bool varyColors = barOptions.VaryColors.Value;
        if ((!labelOptions.HasVisibleContent && !seriesLabelOptions.Any(options => options.HasVisibleContent)) || series.Count == 0)
        {
            return;
        }

        IReadOnlyList<IReadOnlyList<ChartIndexedNumberPoint?>> densePointSeries = DensifyChartPointSeries(series);
        if (densePointSeries.Count == 0)
        {
            return;
        }

        int categoryCount = Math.Max(1, densePointSeries.Max(values => values.Count));
        double zeroX = ChartValueToPlotCoordinate(extents, 0d, plotBox.X, plotBox.Width, valueAxisReversed);
        double zeroY = ChartValueToPlotCoordinate(extents, 0d, plotBox.Y, plotBox.Height, valueAxisReversed);
        var runs = new List<TextRun>();
        List<ChartTextRunLink>? labelLinks = chartRelationships is null ? null : new List<ChartTextRunLink>();
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
                            ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, densePointSeries.Count, varyColors, seriesFills, pointFills, points.Count, !horizontalBars, value);
                            double legendKeyWidth = RenderFillDataLabelLegendKey(graphics, labelBox, fontSize, fill);
                            textX += legendKeyWidth;
                            textWidth = Math.Max(1d, textWidth - legendKeyWidth);
                        }

                        if (!string.IsNullOrEmpty(label))
                        {
                            AddChartLabelRuns(runs, label, effectiveOptions, textX, labelBox.Y, textWidth, labelBox.Height, plotBox, style, TextAlignment.Left, fontResolver, labelLinks);
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
                    PptxSceneChartDataLabelPosition barLegendKeyGapPosition = ResolveChartDataLabelPosition(effectiveOptions.PositionKind);
                    if (effectiveOptions.ShowLegendKey && !stacked && barEndY >= barBaseY &&
                        barLegendKeyGapPosition == PptxSceneChartDataLabelPosition.OutsideEnd)
                    {
                        // Office lifts legend-key out-end baselines by box pad plus
                        // bottom-pad law (other positions and plain labels keep legacy 1.0).
                        y = barEndY + ComputeBarLegendKeyOutEndGap(fontSize);
                    }
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
                        double barLegendKeySwatch = fontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
                        bool barLegendKeyUnitCenter = effectiveOptions.ShowLegendKey && !stacked;
                        if (barLegendKeyUnitCenter)
                        {
                            // Office centers the swatch+text unit on the bar middle (other
                            // groupings and plain labels keep the legacy slot math, unobserved).
                            double barLegendKeyTextWidth = string.IsNullOrEmpty(label) ? 0d : Math.Max(0d, new ChartTextMeasurer(fontResolver).Measure(label, style));
                            double barLegendKeyCategoryWidth = plotBox.Width / categoryCount;
                            double barLegendKeyBarWidth = GetClusteredBarWidth(barLegendKeyCategoryWidth, densePointSeries.Count, barOptions.GapWidth, barOptions.Overlap);
                            double barLegendKeyStep = GetClusteredBarStep(barLegendKeyBarWidth, barOptions.Overlap);
                            double barLegendKeyClusterWidth = barLegendKeyBarWidth + Math.Max(0, densePointSeries.Count - 1) * barLegendKeyStep;
                            double barLegendKeyBarCenterX = plotBox.X + category * barLegendKeyCategoryWidth + (barLegendKeyCategoryWidth - barLegendKeyClusterWidth) / 2d + seriesIndex * barLegendKeyStep + barLegendKeyBarWidth / 2d;
                            double barLegendKeyUnitLeft = ComputeBarLegendKeyUnitLeft(barLegendKeyBarCenterX, barLegendKeySwatch, barLegendKeyTextWidth);
                            labelBox = new ChartLayoutBox(barLegendKeyUnitLeft, labelBox.Y, labelBox.Width, labelBox.Height);
                            textX = barLegendKeyUnitLeft + barLegendKeySwatch + barLegendKeySwatch;
                        }
                        if (effectiveOptions.ShowLegendKey)
                        {
                            ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, densePointSeries.Count, varyColors, seriesFills, pointFills, points.Count, !horizontalBars, value);
                            if (barLegendKeyUnitCenter)
                            {
                                textX = labelBox.X + barLegendKeySwatch + barLegendKeySwatch;
                                textWidth = Math.Max(1d, labelBox.X + labelBox.Width - textX);
                                // Office anchors swatches to text (swatchX equals textX minus G minus S,
                                // center equals baseline plus 0.34fs), not box-centered.
                                double barLegendKeySwatchX = ComputeBarLegendKeySwatchX(textX, fontSize);
                                double barLegendKeySwatchY = ComputeBarLegendKeySwatchY(labelBox.Y, fontSize);
                                ChartLayoutBox barLegendKeySwatchBox = new ChartLayoutBox(
                                    barLegendKeySwatchX,
                                    barLegendKeySwatchY,
                                    barLegendKeySwatch,
                                    barLegendKeySwatch);
                                RenderFillDataLabelLegendKey(graphics, barLegendKeySwatchBox, fontSize, fill);
                            }
                            else
                            {
                                double consumedWidth = RenderFillDataLabelLegendKey(graphics, labelBox, fontSize, fill);
                                textX += consumedWidth;
                                textWidth = Math.Max(1d, textWidth - consumedWidth);
                            }
                            alignment = TextAlignment.Left;
                        }

                        if (!string.IsNullOrEmpty(label))
                        {
                            AddChartLabelRuns(runs, label, effectiveOptions, textX, labelBox.Y, textWidth, labelBox.Height, plotBox, style, alignment, fontResolver, labelLinks);
                        }
                    }
                }
            }
        }

        RenderedFonts labelFonts = RenderChartTextRuns(runs, graphics, chartFonts, "CBD", fontResolver, diagnosticSink);
        AddChartTextRunHyperlinkAnnotations(labelLinks, chartRelationships, context, labelFonts, linkAnnotations, reportedHyperlinkIds);
    }
}
