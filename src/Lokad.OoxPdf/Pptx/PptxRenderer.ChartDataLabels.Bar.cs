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

    private static async ValueTask RenderBarDataLabelsAsync(
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ChartFrameBox chartFrame,
        ChartPlotBox plotBox,
        IReadOnlyList<ChartIndexedNumberVector> series,
        ChartColorPalette chartPalette,
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
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        ChartAxisPlacement? categoryPlacement = null)
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

        IReadOnlyList<IReadOnlyList<double?>> denseValueSeries = DensifyChartValueSeries(series);
        if (denseValueSeries.Count == 0)
        {
            return;
        }

        // A lone point-styled series uses category keys even without variation.
        // Series keys keep series paint and negative inversion; category keys
        // retain their authored point styling.
        IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> legendPointFills =
            chartPalette.SeriesCount == 1 || varyColors && denseValueSeries.Count == 1 ? pointFills : [];

        // index workbook points once per series instead of filtering plus
        // linear-scanning per rendered label below.
        Dictionary<int, ChartIndexedNumberPoint>[] workbookIndexes = series.Select(BuildWorkbookPointIndex).ToArray();

        // RV20: index dense-slot provenance once per series (last-wins, like
        // DensePoints) instead of materializing dense arrays per label below.
        Dictionary<int, ChartIndexedNumberPoint>[] densePointIndexes = series.Select(BuildDensePointIndex).ToArray();

        // RV14: index category labels once per frame instead of linear-scanning
        // per rendered label below.
        Dictionary<int, string> categoryLabelIndex = BuildCategoryLabelIndex(categoryLabels);

        int categoryCount = Math.Max(1, denseValueSeries.Max(values => values.Count));
        double originX = ChartValueToPlotCoordinate(extents, categoryPlacement?.CrossingValue ?? 0d, plotBox.X, plotBox.Width, valueAxisReversed);
        double zeroY = ChartValueToPlotCoordinate(extents, 0d, plotBox.Y, plotBox.Height, valueAxisReversed);
        var runs = new List<TextRun>();
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        List<ChartTextRunLink>? labelLinks = chartRelationships is null ? null : new List<ChartTextRunLink>();
        bool stacked = IsStackedChartGrouping(grouping);
        bool percentStacked = IsPercentStackedChartGrouping(grouping);
        if (horizontalBars)
        {
            double categoryHeight = plotBox.Height / categoryCount;
            double paintedBarHeight = stacked ? GetStackedBarWidth(categoryHeight, gapWidthPercent)
                : GetClusteredBarWidth(categoryHeight, denseValueSeries.Count, gapWidthPercent, barOptions.Overlap);
            double paintedStep = stacked ? 0d : GetClusteredBarStep(paintedBarHeight, barOptions.Overlap);
            double paintedClusterHeight = paintedBarHeight + Math.Max(0, denseValueSeries.Count - 1) * paintedStep;
            double barSlot = stacked
                ? GetStackedBarWidth(categoryHeight, gapWidthPercent)
                : categoryHeight * PptxChartMetricRules.BarDataLabelSlotFillRatio / Math.Max(1, denseValueSeries.Count);
            double labelWidth = Math.Max(PptxChartMetricRules.CartesianDataLabelMinimumWidth, plotBox.Width * PptxChartMetricRules.HorizontalBarDataLabelWidthRatio);
            double[] positiveValues = new double[categoryCount];
            double[] negativeValues = new double[categoryCount];
            double[] positiveTotals = stacked ? GetCategoryPositiveTotals(denseValueSeries, categoryCount, percentStacked) : [];
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryY = stacked
                    ? plotBox.Y + category * categoryHeight + (categoryHeight - barSlot) / 2d
                    : plotBox.Y + category * categoryHeight + categoryHeight * PptxChartMetricRules.BarDataLabelCategoryInsetRatio;
                for (int seriesIndex = 0; seriesIndex < denseValueSeries.Count; seriesIndex++)
                {
                    IReadOnlyList<double?> values = denseValueSeries[seriesIndex];
                    if (category >= values.Count || values[category] is not double value)
                    {
                        continue;
                    }

                    double normalizedValue = stacked ? NormalizeStackedValue(value, positiveTotals[category], percentStacked) : value;
                    (double barBase, double barEnd) = stacked
                        ? ResolveStackedBarSegmentValues(positiveValues, negativeValues, category, normalizedValue)
                        : (0d, value);
                    double barBaseX = stacked
                        ? ChartValueToPlotCoordinate(extents, barBase, plotBox.X, plotBox.Width, valueAxisReversed)
                        : originX;
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
                    if (barOptions.CategoryAxisReversed)
                        y = ChartCategoryToPlotCoordinate(category, categoryCount, plotBox.Y, plotBox.Height,
                            onTickMarks: false, reversed: true,
                            slotOffset: categoryHeight * (PptxChartMetricRules.BarDataLabelCategoryInsetRatio - 0.5d) +
                                seriesIndex * barSlot + barSlot * PptxChartMetricRules.HorizontalBarDataLabelSlotCenterRatio) - labelHeight / 2d;
                    ChartIndexedNumberPoint point = densePointIndexes[seriesIndex].TryGetValue(category, out ChartIndexedNumberPoint densePoint) ? densePoint : default;
                    string label = FormatCartesianDataLabel(value, seriesIndex, category, point, (workbookIndexes[seriesIndex].TryGetValue(point.Index, out ChartIndexedNumberPoint workbookPoint) ? workbookPoint : null), series[seriesIndex].FormatCode, effectiveOptions, categoryLabelIndex, seriesNames);
                    if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
                    {
                        ChartSeriesFill? keyFill = effectiveOptions.ShowLegendKey
                            ? ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, denseValueSeries.Count, varyColors, seriesFills, legendPointFills, values.Count,
                                value - (stacked ? 0d : categoryPlacement?.CrossingValue ?? 0d), barOptions.InvertsNegativePoint(seriesIndex, category)) : null;
                        ChartLayoutBox labelBox;
                        ChartLayoutBox? keyBox = null;
                        double textX, textWidth, baseline, textHeight, descent = 0d;
                        bool measuredCenter = !string.IsNullOrEmpty(label) &&
                            (labelPosition is PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit) &&
                            !effectiveOptions.Layout.HasLayout && effectiveOptions.CustomTextRuns.Count == 0;
                        if (measuredCenter)
                        {
                            // The occupied key/text frame is centered on the same
                            // rectangle painting uses, including grouped spacing.
                            var line = await textMeasurer.MeasureLineBoxAsync(style).ConfigureAwait(false);
                            double advance = await textMeasurer.MeasureAsync(label, style).ConfigureAwait(false);
                            double horizontalPadding = PptxChartMetricRules.DataLabelHorizontalPadding;
                            double verticalPadding = PptxChartMetricRules.DataLabelVerticalPadding;
                            double swatch = fontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
                            double keyReserve = keyFill is not null ? 2d * swatch + horizontalPadding : 0d;
                            double width = advance + 2d * horizontalPadding + keyReserve;
                            double height = line.Height + 2d * verticalPadding;
                            double centerY = plotBox.Y + category * categoryHeight + (categoryHeight - paintedClusterHeight) / 2d +
                                seriesIndex * paintedStep + paintedBarHeight / 2d;
                            if (barOptions.CategoryAxisReversed)
                                centerY = ChartCategoryToPlotCoordinate(category, categoryCount, plotBox.Y, plotBox.Height,
                                    onTickMarks: false, reversed: true,
                                    slotOffset: seriesIndex * paintedStep + (paintedBarHeight - paintedClusterHeight) / 2d);
                            labelBox = new ChartLayoutBox((barBaseX + barEndX - width) / 2d, centerY - height / 2d, width, height);
                            baseline = labelBox.Y + verticalPadding + line.Descent;
                            descent = line.Descent;
                            textHeight = line.Height;
                            textX = labelBox.X + horizontalPadding + keyReserve;
                            textWidth = Math.Max(1d, advance);
                            if (keyFill is not null) keyBox = new ChartLayoutBox(ComputeBarLegendKeySwatchX(textX, fontSize),
                                ComputeBarLegendKeySwatchY(baseline, fontSize), swatch, swatch);
                        }
                        else
                        {
                            // Authored boxes/rich runs and other positions keep
                            // their independently qualified placement fallback.
                            labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, x, y, labelWidth, labelHeight);
                            double keyReserve = keyFill is not null ? fontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor +
                                fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor : 0d;
                            textX = labelBox.X + keyReserve;
                            textWidth = Math.Max(1d, labelBox.Width - keyReserve);
                            baseline = labelBox.Y;
                            textHeight = labelBox.Height;
                            if (keyFill is not null) keyBox = labelBox;
                        }

                        RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                        if (keyFill is { } fill && keyBox is { } key) RenderFillDataLabelLegendKey(graphics, key, fontSize, fill);
                        if (!string.IsNullOrEmpty(label))
                        {
                            await AddChartLabelRunsAsync(runs, label, effectiveOptions, textX, baseline, textWidth, textHeight, plotBox, style, TextAlignment.Left, fontResolver, labelLinks, descent).ConfigureAwait(false);
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
                : categoryWidth * PptxChartMetricRules.BarDataLabelSlotFillRatio / Math.Max(1, denseValueSeries.Count);
            double[] positiveValues = new double[categoryCount];
            double[] negativeValues = new double[categoryCount];
            double[] positiveTotals = stacked ? GetCategoryPositiveTotals(denseValueSeries, categoryCount, percentStacked) : [];
            for (int category = 0; category < categoryCount; category++)
            {
                double categoryX = stacked
                    ? plotBox.X + category * categoryWidth + (categoryWidth - barSlot) / 2d
                    : plotBox.X + category * categoryWidth + categoryWidth * PptxChartMetricRules.BarDataLabelCategoryInsetRatio;
                for (int seriesIndex = 0; seriesIndex < denseValueSeries.Count; seriesIndex++)
                {
                    IReadOnlyList<double?> values = denseValueSeries[seriesIndex];
                    if (category >= values.Count || values[category] is not double value)
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
                    ChartIndexedNumberPoint point = densePointIndexes[seriesIndex].TryGetValue(category, out ChartIndexedNumberPoint densePoint) ? densePoint : default;
                    string label = FormatCartesianDataLabel(value, seriesIndex, category, point, (workbookIndexes[seriesIndex].TryGetValue(point.Index, out ChartIndexedNumberPoint workbookPoint) ? workbookPoint : null), series[seriesIndex].FormatCode, effectiveOptions, categoryLabelIndex, seriesNames);
                    if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
                    {
                        if (!stacked && barEndY >= barBaseY &&
                            ResolveChartDataLabelPosition(labelPosition) == PptxSceneChartDataLabelPosition.OutsideEnd &&
                            CanWrapAutomaticColumnLabel(effectiveOptions) && !string.IsNullOrEmpty(label))
                        {
                            double barWidth = GetClusteredBarWidth(categoryWidth, denseValueSeries.Count, barOptions.GapWidth, barOptions.Overlap);
                            double step = GetClusteredBarStep(barWidth, barOptions.Overlap);
                            double clusterWidth = barWidth + Math.Max(0, denseValueSeries.Count - 1) * step;
                            double centerX = plotBox.X + category * categoryWidth + (categoryWidth - clusterWidth) / 2d + seriesIndex * step + barWidth / 2d;
                            ChartSeriesFill? keyFill = effectiveOptions.ShowLegendKey
                                ? ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, denseValueSeries.Count, varyColors, seriesFills, legendPointFills, values.Count, value, barOptions.InvertsNegativePoint(seriesIndex, category)) : null;
                            await AddAutomaticColumnLabelAsync(runs, graphics, label, effectiveOptions, style,
                                textMeasurer, chartFrame, centerX, barEndY, keyFill, context.CancellationToken).ConfigureAwait(false);
                            continue;
                        }
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
                            double barLegendKeyTextWidth = string.IsNullOrEmpty(label) ? 0d : Math.Max(0d, (await new ChartTextMeasurer(fontResolver).MeasureAsync(label, style).ConfigureAwait(false)));
                            double barLegendKeyCategoryWidth = plotBox.Width / categoryCount;
                            double barLegendKeyBarWidth = GetClusteredBarWidth(barLegendKeyCategoryWidth, denseValueSeries.Count, barOptions.GapWidth, barOptions.Overlap);
                            double barLegendKeyStep = GetClusteredBarStep(barLegendKeyBarWidth, barOptions.Overlap);
                            double barLegendKeyClusterWidth = barLegendKeyBarWidth + Math.Max(0, denseValueSeries.Count - 1) * barLegendKeyStep;
                            double barLegendKeyBarCenterX = plotBox.X + category * barLegendKeyCategoryWidth + (barLegendKeyCategoryWidth - barLegendKeyClusterWidth) / 2d + seriesIndex * barLegendKeyStep + barLegendKeyBarWidth / 2d;
                            double barLegendKeyUnitLeft = ComputeBarLegendKeyUnitLeft(barLegendKeyBarCenterX, barLegendKeySwatch, barLegendKeyTextWidth);
                            labelBox = new ChartLayoutBox(barLegendKeyUnitLeft, labelBox.Y, labelBox.Width, labelBox.Height);
                            textX = barLegendKeyUnitLeft + barLegendKeySwatch + barLegendKeySwatch;
                        }
                        if (effectiveOptions.ShowLegendKey)
                        {
                            ChartSeriesFill fill = ResolveBarPointFill(theme, colorMap, chartPalette, seriesIndex, category, denseValueSeries.Count, varyColors, seriesFills, legendPointFills, values.Count, value, barOptions.InvertsNegativePoint(seriesIndex, category));
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
                            await AddChartLabelRunsAsync(runs, label, effectiveOptions, textX, labelBox.Y, textWidth, labelBox.Height, plotBox, style, alignment, fontResolver, labelLinks).ConfigureAwait(false);
                        }
                    }
                }
            }
        }

        RenderedFonts labelFonts = (await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CBD", fontResolver, diagnosticSink).ConfigureAwait(false));
        AddChartTextRunHyperlinkAnnotations(labelLinks, chartRelationships, context, labelFonts, linkAnnotations, reportedHyperlinkIds);
    }
}
