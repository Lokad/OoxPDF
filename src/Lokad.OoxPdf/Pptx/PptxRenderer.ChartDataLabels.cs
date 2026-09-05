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
    private static IReadOnlyList<PdfFontResource> RenderPieDataLabels(PptxTheme theme, PptxColorMap colorMap, PdfGraphicsBuilder graphics, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<ChartIndexedPieSlice> slices, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, string? valueFormatCode, ChartDataLabelOptions labelOptions, ChartIndexedTextVector categoryLabels, IReadOnlyList<ChartSeriesNameRecord> seriesNames, PresentationFontResolver? fontResolver)
    {
        if (!labelOptions.HasVisibleContent || slices.Count == 0)
        {
            return [];
        }

        double total = slices.Sum(slice => slice.Value);
        if (total <= 0d)
        {
            return [];
        }

        ChartPlotBox plotBox = layout.PlotBox;
        ChartPolarGeometry geometry = layout.Geometry;
        double labelRadius = geometry.Radius * (holeSize > 0d ? Math.Max(PptxChartMetricRules.PieDataLabelRadiusRatio, (1d + holeSize) / 2d) : PptxChartMetricRules.PieDataLabelRadiusRatio);
        double labelWidth = Math.Max(PptxChartMetricRules.PieDataLabelMinimumWidth, geometry.Radius * PptxChartMetricRules.PieDataLabelWidthRatio);
        var runs = new List<TextRun>(slices.Count);
        double angle = -90d;
        foreach (ChartIndexedPieSlice slice in slices)
        {
            ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(labelOptions, slice.Index);
            if (!effectiveOptions.HasVisibleContent)
            {
                continue;
            }

            ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
            double fontSize = style.FontSize;
            double labelHeight = fontSize * PptxChartMetricRules.PieDataLabelHeightFactor;
            double sweep = slice.Value / total * 360d;
            double mid = (angle + sweep / 2d) * Math.PI / 180d;
            double explosion = pointExplosions.TryGetValue(slice.Index, out double offset) ? Math.Clamp(offset, 0d, 1d) * geometry.Radius * PptxChartMetricRules.PieExplosionLabelRadiusRatio : 0d;
            double labelX = geometry.CenterX + Math.Cos(mid) * (labelRadius + explosion) - labelWidth / 2d;
            double labelY = geometry.CenterY + Math.Sin(mid) * (labelRadius + explosion) - labelHeight / 2d;
            IReadOnlyList<string> labelParts = FormatPieDataLabelParts(slice.Value, total, slice.Index, slice.WorkbookPoint, valueFormatCode, effectiveOptions, categoryLabels, seriesNames);
            string label = JoinChartDataLabelParts(labelParts, effectiveOptions);
            if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
            {
                ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, labelWidth, labelHeight);
                RenderPieDataLabelLeaderLine(graphics, geometry, mid, explosion, labelBox, effectiveOptions);
                RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                double textX = labelBox.X;
                double textWidth = labelBox.Width;
                TextAlignment alignment = TextAlignment.Center;
                if (effectiveOptions.ShowLegendKey)
                {
                    double swatchSize = fontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
                    double swatchGap = fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
                    double swatchY = labelBox.Y + Math.Max(0d, (labelBox.Height - swatchSize) / 2d);
                    ChartSeriesFill fill = pointFills.TryGetValue(slice.Index, out ChartSeriesFill explicitFill)
                        ? explicitFill
                        : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, slice.Index), 1d, null, null);
                    FillChartRectangle(graphics, labelBox.X, swatchY, swatchSize, swatchSize, fill);
                    textX += swatchSize + swatchGap;
                    textWidth = Math.Max(1d, textWidth - swatchSize - swatchGap);
                    alignment = TextAlignment.Left;
                }

                if (!string.IsNullOrEmpty(label))
                {
                    AddPolarChartLabelRuns(runs, labelParts, label, effectiveOptions, textX, labelBox.Y, textWidth, labelBox.Height, plotBox, style, alignment, fontResolver);
                }
            }
            angle += sweep;
        }

        return RenderTextRuns(runs, graphics, "CP", fontResolver);
    }

    private static void RenderPieDataLabelLeaderLine(PdfGraphicsBuilder graphics, ChartPolarGeometry geometry, double angleRadians, double explosion, ChartLayoutBox labelBox, ChartDataLabelOptions options)
    {
        if (!options.ShowLeaderLines)
        {
            return;
        }

        ChartSeriesStroke stroke = options.LeaderLines.Stroke ?? ChartDataLabelLeaderLineDefaultStroke;
        if (stroke.Alpha <= 0.001d || stroke.Width <= 0d)
        {
            return;
        }

        double startX = geometry.CenterX + Math.Cos(angleRadians) * (geometry.Radius + explosion);
        double startY = geometry.CenterY + Math.Sin(angleRadians) * (geometry.Radius + explosion);
        double labelCenterX = labelBox.X + labelBox.Width / 2d;
        double labelCenterY = labelBox.Y + labelBox.Height / 2d;
        bool labelIsLeft = labelCenterX < geometry.CenterX;
        double labelEdgeX = labelIsLeft ? labelBox.X + labelBox.Width : labelBox.X;
        double tail = Math.Min(Math.Max(stroke.Width * 2d, 4d), Math.Max(4d, labelBox.Width * 0.18d));
        double elbowX = labelIsLeft ? labelEdgeX + tail : labelEdgeX - tail;

        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        graphics.MoveTo(startX, startY);
        graphics.LineTo(elbowX, labelCenterY);
        graphics.LineTo(labelEdgeX, labelCenterY);
        graphics.StrokeCurrentPath();

        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }

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

    private static IReadOnlyList<PdfFontResource> RenderLineDataLabels(
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<ChartIndexedNumberVector> series,
        ChartValueExtents extents,
        bool valueAxisReversed,
        IReadOnlyList<ChartSeriesStroke?> seriesStrokes,
        IReadOnlyList<ChartMarkerStyle> markerStyles,
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

        int pointCount = Math.Max(1, densePointSeries.Max(values => values.Count));
        double labelWidth = Math.Max(
            PptxChartMetricRules.CartesianDataLabelMinimumWidth,
            plotBox.Width / Math.Max(PptxChartMetricRules.LineDataLabelMinimumPointSpan, pointCount * PptxChartMetricRules.LineDataLabelPointWidthFactor));
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        var runs = new List<TextRun>();
        for (int seriesIndex = 0; seriesIndex < densePointSeries.Count; seriesIndex++)
        {
            IReadOnlyList<ChartIndexedNumberPoint?> points = densePointSeries[seriesIndex];
            for (int i = 0; i < points.Count; i++)
            {
                if (points[i]?.Value is not double value)
                {
                    continue;
                }

                double pointX = plotBox.X + (pointCount == 1 ? plotBox.Width / 2d : plotBox.Width * i / (pointCount - 1));
                double pointY = ChartValueToPlotCoordinate(extents, value, plotBox.Y, plotBox.Height, valueAxisReversed);
                ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), i);
                if (!effectiveOptions.HasVisibleContent)
                {
                    continue;
                }

                ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                double fontSize = style.FontSize;
                double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                ChartIndexedNumberPoint point = points[i] ?? default;
                string label = FormatCartesianDataLabel(value, seriesIndex, i, point, series[seriesIndex].WorkbookPointForIndex(point.Index), series[seriesIndex].FormatCode, effectiveOptions, categoryLabels, seriesNames);
                if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
                {
                    double legendKeyWidth = effectiveOptions.ShowLegendKey
                        ? GetStrokeMarkerDataLabelLegendKeyWidth(fontSize, ChartMarker(seriesIndex, markerStyles))
                        : 0d;
                    double effectiveLabelWidth = effectiveOptions.ShowLegendKey
                        ? Math.Max(labelWidth, textMeasurer.Measure(label, style) + legendKeyWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor)
                        : labelWidth;
                    (double labelX, double labelY, TextAlignment alignment) = ResolveLineDataLabelPosition(
                        effectiveOptions.PositionKind,
                        pointX,
                        pointY,
                        effectiveLabelWidth,
                        labelHeight);
                    ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, effectiveLabelWidth, labelHeight);
                    RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                    double textX = labelBox.X;
                    double textWidth = labelBox.Width;
                    if (effectiveOptions.ShowLegendKey)
                    {
                        ChartSeriesStroke stroke = ChartSeriesStrokeColor(theme, colorMap, null, seriesIndex, seriesStrokes, 1.2d);
                        ChartMarkerStyle marker = ChartMarker(seriesIndex, markerStyles);
                        double consumedWidth = RenderStrokeMarkerDataLabelLegendKey(graphics, labelBox, fontSize, marker, stroke);
                        textX += consumedWidth;
                        textWidth = Math.Max(1d, textWidth - consumedWidth);
                        alignment = TextAlignment.Left;
                    }

                    if (!string.IsNullOrEmpty(label))
                    {
                        AddChartLabelRuns(
                            runs,
                            label,
                            effectiveOptions,
                            textX,
                            labelBox.Y,
                            textWidth,
                            labelBox.Height,
                            plotBox,
                            style,
                            alignment,
                            fontResolver);
                    }
                }
            }
        }

        return RenderTextRuns(runs, graphics, "CLD", fontResolver);
    }

    private static double GetFillDataLabelLegendKeyWidth(double fontSize)
    {
        return fontSize * (PptxChartMetricRules.DataLabelLegendKeySizeFactor + PptxChartMetricRules.DataLabelLegendKeyTextGapFactor);
    }

    private static double RenderFillDataLabelLegendKey(PdfGraphicsBuilder graphics, ChartLayoutBox labelBox, double fontSize, ChartSeriesFill fill)
    {
        double swatchSize = fontSize * PptxChartMetricRules.DataLabelLegendKeySizeFactor;
        double swatchGap = fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
        double swatchY = labelBox.Y + Math.Max(0d, (labelBox.Height - swatchSize) / 2d);
        FillChartRectangle(graphics, labelBox.X, swatchY, swatchSize, swatchSize, fill);
        var stroke = new ChartSeriesStroke(fill.Color, fill.Alpha, ChartFilledSeriesInheritedStrokeWidth);
        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        graphics.StrokeRectangle(labelBox.X, swatchY, swatchSize, swatchSize);
        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }

        return swatchSize + swatchGap;
    }

    private static double GetStrokeMarkerDataLabelLegendKeyWidth(double fontSize, ChartMarkerStyle marker)
    {
        return Math.Max(marker.Size * 2d, fontSize * 1.05d) + fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
    }

    private static double RenderStrokeMarkerDataLabelLegendKey(PdfGraphicsBuilder graphics, ChartLayoutBox labelBox, double fontSize, ChartMarkerStyle marker, ChartSeriesStroke stroke)
    {
        double segmentLength = Math.Max(marker.Size * 2d, fontSize * 1.05d);
        double segmentY = labelBox.Y + labelBox.Height / 2d;
        SetChartStroke(graphics, stroke);
        graphics.StrokeLine(labelBox.X, segmentY, labelBox.X + segmentLength, segmentY);
        DrawChartMarker(graphics, labelBox.X + segmentLength / 2d, segmentY, marker, stroke.Color, stroke.Color);
        return segmentLength + fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
    }

    private static ChartLayoutBox ResolveDataLabelBox(ChartPlotBox plotBox, ChartDataLabelOptions options, double x, double y, double width, double height)
    {
        ChartLayoutBox defaultBox = new(x, y, width, height);
        if (!options.Layout.HasLayout)
        {
            return defaultBox;
        }

        ChartFrameBox frame = new(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height);
        return TryBuildManualLayoutBox(options.Layout, frame, defaultBox, out ChartLayoutBox manualBox, clampToFrame: false, missingPositionModesAreFactor: true)
            ? manualBox
            : defaultBox;
    }

    private static double ResolveHorizontalBarDataLabelX(PptxSceneChartDataLabelPosition position, double barBase, double barEnd, double labelWidth)
    {
        position = ResolveChartDataLabelPosition(position);
        bool extendsRight = barEnd >= barBase;
        return position switch
        {
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (barBase + barEnd - labelWidth) / 2d,
            PptxSceneChartDataLabelPosition.InsideBase => extendsRight ? barBase + PptxChartMetricRules.BarDataLabelHorizontalGap : barBase - labelWidth - PptxChartMetricRules.BarDataLabelHorizontalGap,
            PptxSceneChartDataLabelPosition.InsideEnd or PptxSceneChartDataLabelPosition.Left or PptxSceneChartDataLabelPosition.Right => extendsRight ? barEnd - labelWidth - PptxChartMetricRules.BarDataLabelHorizontalGap : barEnd + PptxChartMetricRules.BarDataLabelHorizontalGap,
            PptxSceneChartDataLabelPosition.OutsideEnd or _ => extendsRight ? barEnd + PptxChartMetricRules.BarDataLabelHorizontalGap : barEnd - labelWidth - PptxChartMetricRules.BarDataLabelHorizontalGap
        };
    }

    private static double ResolveVerticalBarDataLabelY(PptxSceneChartDataLabelPosition position, double barBase, double barEnd, double labelHeight)
    {
        position = ResolveChartDataLabelPosition(position);
        bool extendsUp = barEnd >= barBase;
        return position switch
        {
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (barBase + barEnd - labelHeight) / 2d,
            PptxSceneChartDataLabelPosition.InsideBase => extendsUp ? barBase + PptxChartMetricRules.BarDataLabelVerticalGap : barBase - labelHeight - PptxChartMetricRules.BarDataLabelVerticalGap,
            PptxSceneChartDataLabelPosition.InsideEnd or PptxSceneChartDataLabelPosition.Top or PptxSceneChartDataLabelPosition.Bottom => extendsUp ? barEnd - labelHeight - PptxChartMetricRules.BarDataLabelVerticalGap : barEnd + PptxChartMetricRules.BarDataLabelVerticalGap,
            PptxSceneChartDataLabelPosition.OutsideEnd or _ => extendsUp ? barEnd + PptxChartMetricRules.BarDataLabelVerticalGap : barEnd - labelHeight - PptxChartMetricRules.BarDataLabelVerticalGap
        };
    }

    private static PptxSceneChartDataLabelPosition ResolveStackedBarDataLabelPosition(PptxSceneChartDataLabelPosition position, bool stacked)
    {
        return stacked && position == PptxSceneChartDataLabelPosition.Unknown
            ? PptxSceneChartDataLabelPosition.Center
            : position;
    }

    private static (double Start, double End) ResolveStackedBarSegmentValues(double[] positiveValues, double[] negativeValues, int category, double value)
    {
        if (value >= 0d)
        {
            double start = positiveValues[category];
            positiveValues[category] += value;
            return (start, positiveValues[category]);
        }

        double negativeStart = negativeValues[category];
        negativeValues[category] += value;
        return (negativeStart, negativeValues[category]);
    }

    private static (double X, double Y, TextAlignment Alignment) ResolveLineDataLabelPosition(PptxSceneChartDataLabelPosition position, double pointX, double pointY, double labelWidth, double labelHeight)
    {
        position = ResolveChartDataLabelPosition(position);
        return position switch
        {
            PptxSceneChartDataLabelPosition.Bottom => (pointX - labelWidth / 2d, pointY - labelHeight * PptxChartMetricRules.LineDataLabelBelowOffsetFactor, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Left => (pointX - labelWidth - PptxChartMetricRules.LineDataLabelSideGap, pointY - labelHeight / 2d, TextAlignment.Right),
            PptxSceneChartDataLabelPosition.Right => (pointX + PptxChartMetricRules.LineDataLabelSideGap, pointY - labelHeight / 2d, TextAlignment.Left),
            PptxSceneChartDataLabelPosition.Center or PptxSceneChartDataLabelPosition.BestFit => (pointX - labelWidth / 2d, pointY - labelHeight / 2d, TextAlignment.Center),
            PptxSceneChartDataLabelPosition.Top or PptxSceneChartDataLabelPosition.OutsideEnd or _ => (pointX - labelWidth / 2d, pointY + labelHeight * PptxChartMetricRules.LineDataLabelAboveOffsetFactor, TextAlignment.Center)
        };
    }

    private static PptxSceneChartDataLabelPosition ResolveChartDataLabelPosition(PptxSceneChartDataLabelPosition position)
    {
        return position == PptxSceneChartDataLabelPosition.Unknown
            ? PptxSceneChartDataLabelPosition.OutsideEnd
            : position;
    }

    private static IReadOnlyList<PdfFontResource> RenderScatterDataLabels(
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<ScatterSeries> series,
        bool bubble,
        ChartValueExtents xExtents,
        ChartValueExtents yExtents,
        IReadOnlyList<ChartSeriesFill?> seriesFills,
        ChartDataLabelOptions labelOptions,
        IReadOnlyList<ChartDataLabelOptions> seriesLabelOptions,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames,
        PresentationFontResolver? fontResolver)
    {
        if ((!labelOptions.HasVisibleContent && !seriesLabelOptions.Any(options => options.HasVisibleContent)) || series.Count == 0)
        {
            return [];
        }

        double maxBubbleSize = Math.Max(1d, series.SelectMany(item => item.Points).DefaultIfEmpty().Max(point => point.Size));
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        var runs = new List<TextRun>();
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
        {
            foreach (ScatterPoint point in series[seriesIndex].Points)
            {
                ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(
                    ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex),
                    point.Index);
                if (!effectiveOptions.HasVisibleContent)
                {
                    continue;
                }

                ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                double fontSize = style.FontSize;
                double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                string label = FormatScatterDataLabel(point, seriesIndex, effectiveOptions, seriesNames);
                if (string.IsNullOrEmpty(label) && !effectiveOptions.ShowLegendKey)
                {
                    continue;
                }

                double legendKeyWidth = effectiveOptions.ShowLegendKey
                    ? GetFillDataLabelLegendKeyWidth(fontSize)
                    : 0d;
                double labelWidth = Math.Max(
                    PptxChartMetricRules.CartesianDataLabelMinimumWidth,
                    textMeasurer.Measure(label, style) + legendKeyWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
                (double pointX, double pointY, double radius) = ResolveScatterPointGeometry(
                    plotBox,
                    point,
                    bubble,
                    xExtents,
                    yExtents,
                    maxBubbleSize);
                (double labelX, double labelY, TextAlignment alignment) = ResolveLineDataLabelPosition(
                    effectiveOptions.PositionKind,
                    pointX,
                    bubble ? pointY + radius : pointY,
                    labelWidth,
                    labelHeight);
                ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, labelWidth, labelHeight);
                RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                double textX = labelBox.X;
                double textWidth = labelBox.Width;
                if (effectiveOptions.ShowLegendKey)
                {
                    double consumedWidth = RenderFillDataLabelLegendKey(graphics, labelBox, fontSize, ChartSeriesColor(theme, colorMap, null, seriesIndex, seriesFills, 1d));
                    textX += consumedWidth;
                    textWidth = Math.Max(1d, textWidth - consumedWidth);
                    alignment = TextAlignment.Left;
                }

                if (!string.IsNullOrEmpty(label))
                {
                    AddChartLabelRuns(
                        runs,
                        label,
                        effectiveOptions,
                        textX,
                        labelBox.Y,
                        textWidth,
                        labelBox.Height,
                        plotBox,
                        style,
                        alignment,
                        fontResolver);
                }
            }
        }

        return RenderTextRuns(runs, graphics, "CSD", fontResolver);
    }

    private static string FormatScatterDataLabel(
        ScatterPoint point,
        int seriesIndex,
        ChartDataLabelOptions options,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames)
    {
        if (!string.IsNullOrWhiteSpace(options.CustomText))
        {
            return options.CustomText;
        }

        var parts = new List<string>(3);
        string seriesName = GetActiveSeriesName(seriesNames, seriesIndex);
        if (options.ShowSeriesName && !string.IsNullOrWhiteSpace(seriesName))
        {
            parts.Add(seriesName);
        }

        if (options.ShowValue)
        {
            parts.Add(FormatChartDataLabelValue(point.Y, options, point.YWorkbookPoint ?? point.YPoint, point.YFormatCode));
        }

        if (options.ShowBubbleSize && ResolveBubbleSizeValue(point) is { } bubbleSize)
        {
            parts.Add(FormatBubbleSizeDataLabelValue(bubbleSize, point, options));
        }

        return string.Join(GetChartDataLabelSeparator(options), parts);
    }

    private static double? ResolveBubbleSizeValue(ScatterPoint point)
    {
        return point.BubbleSizeWorkbookPoint?.Value ??
            point.BubbleSizePoint?.Value;
    }

    private static string FormatBubbleSizeDataLabelValue(double value, ScatterPoint point, ChartDataLabelOptions options)
    {
        if (options.NumberFormatInfo.IsDefined || !string.IsNullOrWhiteSpace(options.NumberFormat))
        {
            return FormatChartDataLabelValue(value, options, point.BubbleSizeWorkbookPoint ?? point.BubbleSizePoint, point.BubbleSizeFormatCode);
        }

        string formatCode = point.BubbleSizeFormatCode ?? string.Empty;
        return !string.IsNullOrWhiteSpace(formatCode) &&
            !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)
            ? FormatChartNumber(value, formatCode)
            : FormatChartAxisLabel(value, null);
    }

    private static (double X, double Y, double Radius) ResolveScatterPointGeometry(
        ChartPlotBox plotBox,
        ScatterPoint point,
        bool bubble,
        ChartValueExtents xExtents,
        ChartValueExtents yExtents,
        double maxBubbleSize)
    {
        double xRange = Math.Max(1d, xExtents.Max - xExtents.Min);
        double yRange = Math.Max(1d, yExtents.Max - yExtents.Min);
        double pointX = plotBox.X + (point.X - xExtents.Min) / xRange * plotBox.Width;
        double pointY = plotBox.Y + (point.Y - yExtents.Min) / yRange * plotBox.Height;
        double radius = bubble
            ? Math.Sqrt(Math.Max(0d, point.Size) / Math.Max(1d, maxBubbleSize)) * Math.Min(plotBox.Width, plotBox.Height) * PptxChartMetricRules.BubbleRadiusPlotRatio
            : 3d;
        return (pointX, pointY, radius);
    }

    private static ChartTextStyle ResolveChartDataLabelTextStyle(PptxTheme theme, PptxColorMap colorMap, ChartDataLabelOptions options)
    {
        RgbColor fallbackColor = theme.TryResolveColor("tx1", colorMap, out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        PptxThemeTypefaceResolution typeface = ResolveChartThemeTypeface(theme);
        ChartTextStyle style = new(
            typeface.Typeface,
            PptxChartMetricRules.DataLabelFallbackFontSize,
            0d,
            fallbackColor,
            Alpha: 1d,
            Bold: false,
            Italic: false,
            Underline: false,
            Strike: false,
            typeface.RequestedTypeface,
            typeface.Typeface is null ? null : typeface.Source);
        return MergeChartTextStyle(style, options.TextStyle);
    }

    private static ChartDataLabelOptions ResolveChartDataLabelOptions(ChartDataLabelOptions options, int index)
    {
        if (!options.Overrides.TryGetValue(index, out ChartDataLabelOverride dataLabel))
        {
            return options;
        }

        if (dataLabel.IsDeleted == true)
        {
            return ChartDataLabelOptions.None;
        }

        ChartTextStyleOverride textStyle = new(
            dataLabel.TextStyle.FontFamily ?? options.TextStyle.FontFamily,
            dataLabel.TextStyle.FontSize ?? options.TextStyle.FontSize,
            dataLabel.TextStyle.CharacterSpacing ?? options.TextStyle.CharacterSpacing,
            dataLabel.TextStyle.Color ?? options.TextStyle.Color,
            dataLabel.TextStyle.Alpha ?? options.TextStyle.Alpha,
            dataLabel.TextStyle.Bold ?? options.TextStyle.Bold,
            dataLabel.TextStyle.Italic ?? options.TextStyle.Italic,
            dataLabel.TextStyle.Underline ?? options.TextStyle.Underline,
            dataLabel.TextStyle.Strike ?? options.TextStyle.Strike,
            dataLabel.TextStyle.FontFamily is null ? options.TextStyle.RequestedTypeface : dataLabel.TextStyle.RequestedTypeface,
            dataLabel.TextStyle.FontFamily is null ? options.TextStyle.TypefaceSource : dataLabel.TextStyle.TypefaceSource);
        return options with
        {
            ShowValue = dataLabel.ShowValue ?? options.ShowValue,
            ShowPercent = dataLabel.ShowPercent ?? options.ShowPercent,
            ShowCategoryName = dataLabel.ShowCategoryName ?? options.ShowCategoryName,
            ShowSeriesName = dataLabel.ShowSeriesName ?? options.ShowSeriesName,
            ShowLeaderLines = dataLabel.ShowLeaderLines ?? options.ShowLeaderLines,
            ShowLegendKey = dataLabel.ShowLegendKey ?? options.ShowLegendKey,
            ShowBubbleSize = dataLabel.ShowBubbleSize ?? options.ShowBubbleSize,
            LeaderLines = dataLabel.LeaderLines.IsDefined ? dataLabel.LeaderLines : options.LeaderLines,
            CustomText = string.IsNullOrEmpty(dataLabel.CustomText) ? options.CustomText : dataLabel.CustomText,
            CustomTextRuns = dataLabel.CustomTextRuns.Count == 0 ? options.CustomTextRuns : dataLabel.CustomTextRuns,
            PositionKind = string.IsNullOrEmpty(dataLabel.Position) ? options.PositionKind : dataLabel.PositionKind,
            Position = string.IsNullOrEmpty(dataLabel.Position) ? options.Position : dataLabel.Position,
            Separator = string.IsNullOrEmpty(dataLabel.Separator) ? options.Separator : dataLabel.Separator,
            NumberFormat = string.IsNullOrEmpty(dataLabel.NumberFormat) ? options.NumberFormat : dataLabel.NumberFormat,
            NumberFormatInfo = dataLabel.NumberFormatInfo.IsDefined ? dataLabel.NumberFormatInfo : options.NumberFormatInfo,
            Layout = dataLabel.Layout.HasLayout ? dataLabel.Layout : options.Layout,
            TextStyle = textStyle,
            TextBodyProperties = IsChartTextBodyPropertiesEmpty(dataLabel.TextBodyProperties) ? options.TextBodyProperties : dataLabel.TextBodyProperties,
            ShapeStyle = dataLabel.ShapeStyle.IsEmpty ? options.ShapeStyle : dataLabel.ShapeStyle,
            FlagOptions = ResolveChartDataLabelFlagOptions(options.FlagOptions, dataLabel.FlagOptions),
            Overrides = EmptyChartDataLabelOverrides
        };
    }

    private static ChartDataLabelOptions ResolveChartDataLabelOptionsForSeries(ChartDataLabelOptions plotOptions, IReadOnlyList<ChartDataLabelOptions> seriesOptions, int seriesIndex)
    {
        if (seriesIndex >= seriesOptions.Count || !seriesOptions[seriesIndex].IsDefined)
        {
            return plotOptions;
        }

        return MergeChartDataLabelOptions(plotOptions, seriesOptions[seriesIndex]);
    }

    private static ChartDataLabelOptions MergeChartDataLabelOptions(ChartDataLabelOptions baseOptions, ChartDataLabelOptions overrideOptions)
    {
        IReadOnlyDictionary<string, ChartBooleanOption> flags = ResolveChartDataLabelFlagOptions(baseOptions.FlagOptions, overrideOptions.FlagOptions);
        ChartTextStyleOverride textStyle = new(
            overrideOptions.TextStyle.FontFamily ?? baseOptions.TextStyle.FontFamily,
            overrideOptions.TextStyle.FontSize ?? baseOptions.TextStyle.FontSize,
            overrideOptions.TextStyle.CharacterSpacing ?? baseOptions.TextStyle.CharacterSpacing,
            overrideOptions.TextStyle.Color ?? baseOptions.TextStyle.Color,
            overrideOptions.TextStyle.Alpha ?? baseOptions.TextStyle.Alpha,
            overrideOptions.TextStyle.Bold ?? baseOptions.TextStyle.Bold,
            overrideOptions.TextStyle.Italic ?? baseOptions.TextStyle.Italic,
            overrideOptions.TextStyle.Underline ?? baseOptions.TextStyle.Underline,
            overrideOptions.TextStyle.Strike ?? baseOptions.TextStyle.Strike,
            overrideOptions.TextStyle.FontFamily is null ? baseOptions.TextStyle.RequestedTypeface : overrideOptions.TextStyle.RequestedTypeface,
            overrideOptions.TextStyle.FontFamily is null ? baseOptions.TextStyle.TypefaceSource : overrideOptions.TextStyle.TypefaceSource);

        return baseOptions with
        {
            ShowValue = ChartDataLabelFlagValue(flags, "showVal"),
            ShowPercent = ChartDataLabelFlagValue(flags, "showPercent"),
            ShowCategoryName = ChartDataLabelFlagValue(flags, "showCatName"),
            ShowSeriesName = ChartDataLabelFlagValue(flags, "showSerName"),
            ShowLeaderLines = ChartDataLabelFlagValue(flags, "showLeaderLines"),
            ShowLegendKey = ChartDataLabelFlagValue(flags, "showLegendKey"),
            ShowBubbleSize = ChartDataLabelFlagValue(flags, "showBubbleSize"),
            LeaderLines = overrideOptions.LeaderLines.IsDefined ? overrideOptions.LeaderLines : baseOptions.LeaderLines,
            CustomText = string.IsNullOrEmpty(overrideOptions.CustomText) ? baseOptions.CustomText : overrideOptions.CustomText,
            CustomTextRuns = overrideOptions.CustomTextRuns.Count == 0 ? baseOptions.CustomTextRuns : overrideOptions.CustomTextRuns,
            PositionKind = string.IsNullOrEmpty(overrideOptions.Position) ? baseOptions.PositionKind : overrideOptions.PositionKind,
            Position = string.IsNullOrEmpty(overrideOptions.Position) ? baseOptions.Position : overrideOptions.Position,
            Separator = string.IsNullOrEmpty(overrideOptions.Separator) ? baseOptions.Separator : overrideOptions.Separator,
            NumberFormat = string.IsNullOrEmpty(overrideOptions.NumberFormat) ? baseOptions.NumberFormat : overrideOptions.NumberFormat,
            NumberFormatInfo = overrideOptions.NumberFormatInfo.IsDefined ? overrideOptions.NumberFormatInfo : baseOptions.NumberFormatInfo,
            Layout = overrideOptions.Layout.HasLayout ? overrideOptions.Layout : baseOptions.Layout,
            TextStyle = textStyle,
            TextBodyProperties = IsChartTextBodyPropertiesEmpty(overrideOptions.TextBodyProperties) ? baseOptions.TextBodyProperties : overrideOptions.TextBodyProperties,
            ShapeStyle = overrideOptions.ShapeStyle.IsEmpty ? baseOptions.ShapeStyle : overrideOptions.ShapeStyle,
            FlagOptions = flags,
            Overrides = MergeChartDataLabelOverrides(baseOptions.Overrides, overrideOptions.Overrides),
            IsDefined = baseOptions.IsDefined || overrideOptions.IsDefined
        };
    }

    private static bool ChartDataLabelFlagValue(IReadOnlyDictionary<string, ChartBooleanOption> flags, string name)
    {
        return flags.TryGetValue(name, out ChartBooleanOption option) && option.Value;
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> MergeChartDataLabelOverrides(IReadOnlyDictionary<int, ChartDataLabelOverride> baseOverrides, IReadOnlyDictionary<int, ChartDataLabelOverride> overrideOverrides)
    {
        if (baseOverrides.Count == 0)
        {
            return overrideOverrides;
        }

        if (overrideOverrides.Count == 0)
        {
            return baseOverrides;
        }

        var merged = new Dictionary<int, ChartDataLabelOverride>(baseOverrides);
        foreach (KeyValuePair<int, ChartDataLabelOverride> item in overrideOverrides)
        {
            merged[item.Key] = item.Value;
        }

        return merged;
    }

    private static bool IsChartTextBodyPropertiesEmpty(PptxSceneChartTextBodyProperties properties)
    {
        return properties.RotationDegrees is null &&
            string.IsNullOrEmpty(properties.RotationValue) &&
            string.IsNullOrEmpty(properties.OrientationValue) &&
            string.IsNullOrEmpty(properties.VerticalOverflowValue);
    }

    private static TextRun CreateChartLabelRun(string text, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment)
    {
        return CreateChartTextRun(text, x, y, width, height, plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height, style, alignment);
    }

    private static void AddChartLabelRuns(List<TextRun> runs, string text, ChartDataLabelOptions options, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver)
    {
        ChartLayoutBox clipBox = ResolveDataLabelTextClipBox(plotBox, options, x, y, width, height);
        if (options.CustomTextRuns.Count == 0)
        {
            runs.Add(CreateChartTextRun(text, x, y, width, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, style, alignment));
            return;
        }

        AddChartRichTextRuns(runs, options.CustomTextRuns, text, x, y, width, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, style, alignment, fontResolver);
    }

    private static void AddPolarChartLabelRuns(List<TextRun> runs, IReadOnlyList<string> parts, string fallbackText, ChartDataLabelOptions options, double x, double y, double width, double height, ChartPlotBox plotBox, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver)
    {
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        ChartLayoutBox clipBox = ResolveDataLabelTextClipBox(plotBox, options, x, y, width, height);
        if (parts.Count <= 1 || options.CustomTextRuns.Count > 0 || !ShouldSplitPolarDataLabelParts(options))
        {
            AddChartLabelRuns(runs, fallbackText, options, x, y, width, height, plotBox, style, alignment, fontResolver);
            return;
        }

        ChartTextRunLayout[] labelRuns = parts
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .Select(part => new ChartTextRunLayout(part, style, Math.Max(0d, textMeasurer.Measure(part, style))))
            .Where(run => run.Width > 0d)
            .ToArray();
        if (labelRuns.Length <= 1)
        {
            AddChartLabelRuns(runs, fallbackText, options, x, y, width, height, plotBox, style, alignment, fontResolver);
            return;
        }

        double separatorWidth = textMeasurer.Measure(GetChartDataLabelSeparator(options), style);
        double totalWidth = labelRuns.Sum(run => run.Width) + separatorWidth * Math.Max(0, labelRuns.Length - 1);
        double cursor = alignment switch
        {
            TextAlignment.Right => x + Math.Max(1d, width) - totalWidth,
            TextAlignment.Center => x + (Math.Max(1d, width) - totalWidth) / 2d,
            _ => x
        };

        foreach (ChartTextRunLayout run in labelRuns)
        {
            double runWidth = Math.Max(0.1d, run.Width);
            runs.Add(CreateChartTextRun(run.Text, cursor, y, runWidth, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, run.Style, TextAlignment.Left) with { PreventCoalesce = true });
            cursor += runWidth + separatorWidth;
        }
    }

    private static bool ShouldSplitPolarDataLabelParts(ChartDataLabelOptions options)
    {
        string separator = GetChartDataLabelSeparator(options);
        return separator.Length > 0 && separator.All(char.IsWhiteSpace);
    }

    private static ChartLayoutBox ResolveDataLabelTextClipBox(ChartPlotBox plotBox, ChartDataLabelOptions options, double x, double y, double width, double height)
    {
        if (!options.Layout.HasLayout)
        {
            return new ChartLayoutBox(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height);
        }

        double left = Math.Min(plotBox.X, x);
        double top = Math.Min(plotBox.Y, y);
        double right = Math.Max(plotBox.X + plotBox.Width, x + width);
        double bottom = Math.Max(plotBox.Y + plotBox.Height, y + height);
        return new ChartLayoutBox(left, top, Math.Max(1d, right - left), Math.Max(1d, bottom - top));
    }

    private static void AddChartRichTextRuns(List<TextRun> runs, IReadOnlyList<ChartTextRunOverride> richTextRuns, string fallbackText, double x, double y, double width, double height, double clipX, double clipY, double clipWidth, double clipHeight, ChartTextStyle style, TextAlignment alignment, PresentationFontResolver? fontResolver)
    {
        if (richTextRuns.Count == 0)
        {
            runs.Add(CreateChartTextRun(fallbackText, x, y, width, height, clipX, clipY, clipWidth, clipHeight, style, alignment));
            return;
        }

        var textMeasurer = new ChartTextMeasurer(fontResolver);
        ChartTextRunLayout[] richRuns = richTextRuns
            .Where(run => !string.IsNullOrEmpty(run.Text))
            .Select(run =>
            {
                ChartTextStyle runStyle = MergeChartTextStyle(style, run.TextStyle);
                return new ChartTextRunLayout(run.Text, runStyle, Math.Max(0d, textMeasurer.Measure(run.Text, runStyle)));
            })
            .Where(run => run.Width > 0d)
            .ToArray();
        if (richRuns.Length == 0)
        {
            runs.Add(CreateChartTextRun(fallbackText, x, y, width, height, clipX, clipY, clipWidth, clipHeight, style, alignment));
            return;
        }

        double totalWidth = richRuns.Sum(run => run.Width);
        double cursor = alignment switch
        {
            TextAlignment.Right => x + Math.Max(1d, width) - totalWidth,
            TextAlignment.Center => x + (Math.Max(1d, width) - totalWidth) / 2d,
            _ => x
        };

        foreach (ChartTextRunLayout run in richRuns)
        {
            double runWidth = Math.Max(0.1d, run.Width);
            runs.Add(CreateChartTextRun(run.Text, cursor, y, runWidth, height, clipX, clipY, clipWidth, clipHeight, run.Style, TextAlignment.Left) with { PreventCoalesce = true });
            cursor += runWidth;
        }
    }

    private static TextRun CreateChartTextRun(string text, double x, double y, double width, double height, double clipX, double clipY, double clipWidth, double clipHeight, ChartTextStyle style, TextAlignment alignment)
    {
        return new TextRun(
            text,
            x,
            y,
            Math.Max(1d, width),
            height,
            clipX,
            clipY,
            clipWidth,
            clipHeight,
            style.FontSize,
            style.CharacterSpacing,
            0d,
            style.Color,
            style.Alpha,
            null,
            Bold: style.Bold,
            Italic: style.Italic,
            Underline: style.Underline,
            Strike: style.Strike,
            KerningEnabled: true,
            alignment,
            FontFamily: style.FontFamily,
            RotationDegrees: 0d,
            RotationCenterX: 0d,
            RotationCenterY: 0d,
            FlipHorizontal: false,
            FlipVertical: false);
    }

    private static ChartTextStyle ReadChartTextStyle(PptxTheme theme, XDocument chartXml, XElement? element, double fallbackFontSize)
    {
        return ReadChartTextStyle(theme, PptxColorMap.Default, chartXml, element, fallbackFontSize);
    }

    private static ChartTextStyle ReadChartTextStyle(PptxTheme theme, PptxColorMap colorMap, XDocument chartXml, XElement? element, double fallbackFontSize)
    {
        ChartTextStyle style = CreateDefaultChartTextStyle(theme, colorMap, fallbackFontSize);
        style = MergeChartTextStyle(style, ToChartTextStyleOverride(PptxSceneBuilder.ReadChartTextStyleOverride(chartXml.Root, theme, colorMap)));
        style = MergeChartTextStyle(style, ToChartTextStyleOverride(PptxSceneBuilder.ReadChartTextStyleOverride(element, theme, colorMap)));
        return style;
    }

    private static ChartTextStyle ReadSceneOrXmlChartTextStyle(PptxTheme theme, PptxSceneChart? sceneChart, PptxSceneChartAxis? sceneAxis, XDocument chartXml, XElement? element, double fallbackFontSize, string? chartStyleRole)
    {
        if (sceneChart is null)
        {
            return ReadChartTextStyle(theme, chartXml, element, fallbackFontSize);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, sceneChart.ColorMap, fallbackFontSize);
        return MergeChartTextStyle(style, ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartAxisTextStyleOverride(sceneChart, sceneAxis, chartStyleRole)));
    }

    private static ChartTextStyle CreateDefaultChartTextStyle(PptxTheme theme, double fallbackFontSize)
    {
        return CreateDefaultChartTextStyle(theme, PptxColorMap.Default, fallbackFontSize);
    }

    private static ChartTextStyle CreateDefaultChartTextStyle(PptxTheme theme, PptxColorMap colorMap, double fallbackFontSize)
    {
        RgbColor fallbackColor = theme.TryResolveColor("tx1", colorMap, out RgbColor themeText)
            ? themeText
            : new RgbColor(0, 0, 0);
        PptxThemeTypefaceResolution typeface = ResolveChartThemeTypeface(theme);
        return new ChartTextStyle(
            typeface.Typeface,
            fallbackFontSize,
            0d,
            fallbackColor,
            Alpha: 1d,
            Bold: false,
            Italic: false,
            Underline: false,
            Strike: false,
            typeface.RequestedTypeface,
            typeface.Typeface is null ? null : typeface.Source);
    }

    private static ChartTextStyleOverride ToChartTextStyleOverride(PptxSceneChartTextStyleOverride style)
    {
        return new ChartTextStyleOverride(
            style.FontFamily,
            style.FontSize,
            style.CharacterSpacing,
            style.Color,
            style.Alpha,
            style.Bold,
            style.Italic,
            style.Underline,
            style.Strike,
            style.RequestedTypeface,
            style.TypefaceSource);
    }

    private static PptxThemeTypefaceResolution ResolveChartThemeTypeface(PptxTheme theme)
    {
        PptxThemeTypefaceResolution minorLatin = theme.ResolveTypefaceWithSource("+mn-lt");
        if (minorLatin.Typeface is not null)
        {
            return minorLatin;
        }

        PptxThemeTypefaceResolution majorLatin = theme.ResolveTypefaceWithSource("+mj-lt");
        return majorLatin.Typeface is null ? default : majorLatin;
    }

    private static ChartTextStyle MergeChartTextStyle(ChartTextStyle style, ChartTextStyleOverride next)
    {
        return new ChartTextStyle(
            next.FontFamily ?? style.FontFamily,
            next.FontSize ?? style.FontSize,
            next.CharacterSpacing ?? style.CharacterSpacing,
            next.Color ?? style.Color,
            next.Alpha ?? style.Alpha,
            next.Bold ?? style.Bold,
            next.Italic ?? style.Italic,
            next.Underline ?? style.Underline,
            next.Strike ?? style.Strike,
            next.FontFamily is null ? style.RequestedTypeface : next.RequestedTypeface,
            next.FontFamily is null ? style.TypefaceSource : next.TypefaceSource);
    }

    private static ChartDataLabelOptions ReadChartDataLabelOptions(XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        PptxSceneChartDataLabels labels = PptxSceneBuilder.ReadChartDataLabels(chartElement, theme, colorMap);
        return labels.IsDefined
            ? ToChartDataLabelOptions(sceneChart: null, labels)
            : ChartDataLabelOptions.None;
    }

    private static ChartDataLabelOptions ReadSceneOrXmlDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        return plot is null
            ? ReadChartDataLabelOptions(chartElement, theme, colorMap)
            : new ChartDataLabelOptions(
                plot.DataLabels.ShowValue == true,
                plot.DataLabels.ShowPercent == true,
                plot.DataLabels.ShowCategoryName == true,
                plot.DataLabels.ShowSeriesName == true,
                plot.DataLabels.ShowLeaderLines == true,
                plot.DataLabels.ShowLegendKey == true,
                plot.DataLabels.ShowBubbleSize == true,
                ToChartDataLabelLeaderLines(plot.DataLabels.LeaderLines),
                string.Empty,
                [],
                plot.DataLabels.PositionKind,
                plot.DataLabels.Position,
                plot.DataLabels.Separator,
                plot.DataLabels.NumberFormat,
                ToChartNumberFormat(plot.DataLabels.NumberFormatInfo),
                plot.DataLabels.Layout,
                ToChartDataLabelTextStyleOverride(sceneChart, plot.DataLabels),
                plot.DataLabels.TextBodyProperties,
                ToChartShapeStyle(plot.DataLabels.ShapeStyle),
                ToChartDataLabelFlagOptions(plot.DataLabels),
                ToChartDataLabelOverrides(plot.DataLabels.Overrides),
                plot.DataLabels.IsDefined);
    }

    private static IReadOnlyList<ChartDataLabelOptions> ReadSceneOrXmlSeriesDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartPlot? plot, XElement chartElement, PptxTheme theme, PptxColorMap colorMap)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select(series => ToChartDataLabelOptions(sceneChart, series.DataLabels))
                .ToArray();
        }

        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select(series => ReadChartDataLabelOptions(series, theme, colorMap))
            .ToArray();
    }

    private static ChartDataLabelOptions ToChartDataLabelOptions(PptxSceneChart? sceneChart, PptxSceneChartDataLabels labels)
    {
        return new ChartDataLabelOptions(
            labels.ShowValue == true,
            labels.ShowPercent == true,
            labels.ShowCategoryName == true,
            labels.ShowSeriesName == true,
            labels.ShowLeaderLines == true,
            labels.ShowLegendKey == true,
            labels.ShowBubbleSize == true,
            ToChartDataLabelLeaderLines(labels.LeaderLines),
            string.Empty,
            [],
            labels.PositionKind,
            labels.Position,
            labels.Separator,
            labels.NumberFormat,
            ToChartNumberFormat(labels.NumberFormatInfo),
            labels.Layout,
            ToChartDataLabelTextStyleOverride(sceneChart, labels),
            labels.TextBodyProperties,
            ToChartShapeStyle(labels.ShapeStyle),
            ToChartDataLabelFlagOptions(labels),
            ToChartDataLabelOverrides(labels.Overrides),
            labels.IsDefined);
    }

    private static ChartTextStyleOverride ToChartDataLabelTextStyleOverride(PptxSceneChart? sceneChart, PptxSceneChartDataLabels labels)
    {
        ChartTextStyleOverride style = ChartTextStyleOverride.Empty;
        if (sceneChart is not null)
        {
            style = ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartDataLabelTextStyleOverride(sceneChart, labels));
            return style;
        }

        return ToChartTextStyleOverride(labels.TextStyle);
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> ToChartDataLabelOverrides(IReadOnlyList<PptxSceneChartDataLabelOverride> overrides)
    {
        if (overrides.Count == 0)
        {
            return EmptyChartDataLabelOverrides;
        }

        var result = new Dictionary<int, ChartDataLabelOverride>(overrides.Count);
        foreach (PptxSceneChartDataLabelOverride dataLabel in overrides)
        {
            result[dataLabel.Index] = new ChartDataLabelOverride(
                dataLabel.IsDeleted,
                dataLabel.ShowValue,
                dataLabel.ShowPercent,
                dataLabel.ShowCategoryName,
                dataLabel.ShowSeriesName,
                dataLabel.ShowLeaderLines,
                dataLabel.ShowLegendKey,
                dataLabel.ShowBubbleSize,
                ToChartDataLabelLeaderLines(dataLabel.LeaderLines),
                dataLabel.CustomText,
                ToChartTextRuns(dataLabel.CustomTextRuns),
                dataLabel.PositionKind,
                dataLabel.Position,
                dataLabel.Separator,
                dataLabel.NumberFormat,
                ToChartNumberFormat(dataLabel.NumberFormatInfo),
                dataLabel.Layout,
                ToChartTextStyleOverride(dataLabel.TextStyle),
                dataLabel.TextBodyProperties,
                ToChartShapeStyle(dataLabel.ShapeStyle),
                ToChartDataLabelOverrideFlagOptions(dataLabel));
        }

        return result;
    }

    private static ChartNumberFormat ToChartNumberFormat(PptxSceneChartNumberFormat numberFormat)
    {
        return numberFormat.IsDefined
            ? new ChartNumberFormat(numberFormat.IsDefined, numberFormat.FormatCode, numberFormat.SourceLinked, numberFormat.SourceLinkedValue)
            : default;
    }

    private static ChartDataLabelLeaderLines ToChartDataLabelLeaderLines(PptxSceneChartLeaderLines leaderLines)
    {
        return leaderLines.IsDefined
            ? new ChartDataLabelLeaderLines(IsDefined: true, ToChartSeriesStroke(leaderLines.Line, null))
            : ChartDataLabelLeaderLines.Empty;
    }

    private static IReadOnlyList<ChartTextRunOverride> ToChartTextRuns(IReadOnlyList<PptxSceneChartTextRun> runs)
    {
        return runs.Count == 0
            ? []
            : runs.Select(run => new ChartTextRunOverride(run.Text, ToChartTextStyleOverride(run.TextStyle))).ToArray();
    }

    private static IReadOnlyDictionary<string, ChartBooleanOption> ToChartDataLabelFlagOptions(PptxSceneChartDataLabels labels)
    {
        return new Dictionary<string, ChartBooleanOption>(StringComparer.Ordinal)
        {
            ["showVal"] = new(labels.ShowValue == true, labels.ShowValueValue, labels.ShowValue is not null),
            ["showPercent"] = new(labels.ShowPercent == true, labels.ShowPercentValue, labels.ShowPercent is not null),
            ["showCatName"] = new(labels.ShowCategoryName == true, labels.ShowCategoryNameValue, labels.ShowCategoryName is not null),
            ["showSerName"] = new(labels.ShowSeriesName == true, labels.ShowSeriesNameValue, labels.ShowSeriesName is not null),
            ["showLeaderLines"] = new(labels.ShowLeaderLines == true, labels.ShowLeaderLinesValue, labels.ShowLeaderLines is not null),
            ["showLegendKey"] = new(labels.ShowLegendKey == true, labels.ShowLegendKeyValue, labels.ShowLegendKey is not null),
            ["showBubbleSize"] = new(labels.ShowBubbleSize == true, labels.ShowBubbleSizeValue, labels.ShowBubbleSize is not null)
        };
    }

    private static IReadOnlyDictionary<string, ChartBooleanOption> ToChartDataLabelOverrideFlagOptions(PptxSceneChartDataLabelOverride label)
    {
        return new Dictionary<string, ChartBooleanOption>(StringComparer.Ordinal)
        {
            ["showVal"] = new(label.ShowValue == true, label.ShowValueValue, label.ShowValue is not null),
            ["showPercent"] = new(label.ShowPercent == true, label.ShowPercentValue, label.ShowPercent is not null),
            ["showCatName"] = new(label.ShowCategoryName == true, label.ShowCategoryNameValue, label.ShowCategoryName is not null),
            ["showSerName"] = new(label.ShowSeriesName == true, label.ShowSeriesNameValue, label.ShowSeriesName is not null),
            ["showLeaderLines"] = new(label.ShowLeaderLines == true, label.ShowLeaderLinesValue, label.ShowLeaderLines is not null),
            ["showLegendKey"] = new(label.ShowLegendKey == true, label.ShowLegendKeyValue, label.ShowLegendKey is not null),
            ["showBubbleSize"] = new(label.ShowBubbleSize == true, label.ShowBubbleSizeValue, label.ShowBubbleSize is not null)
        };
    }

    private static IReadOnlyDictionary<string, ChartBooleanOption> ResolveChartDataLabelFlagOptions(
        IReadOnlyDictionary<string, ChartBooleanOption> baseFlags,
        IReadOnlyDictionary<string, ChartBooleanOption> overrideFlags)
    {
        if (overrideFlags.Count == 0)
        {
            return baseFlags;
        }

        var resolved = new Dictionary<string, ChartBooleanOption>(ChartDataLabelFlagNames.Length, StringComparer.Ordinal);
        foreach (string flagName in ChartDataLabelFlagNames)
        {
            if (overrideFlags.TryGetValue(flagName, out ChartBooleanOption overrideOption) && overrideOption.IsDefined)
            {
                resolved[flagName] = overrideOption;
            }
            else if (baseFlags.TryGetValue(flagName, out ChartBooleanOption baseOption))
            {
                resolved[flagName] = baseOption;
            }
            else
            {
                resolved[flagName] = new ChartBooleanOption(false, string.Empty, false);
            }
        }

        return resolved;
    }

    private static string FormatChartPercentageLabel(double fraction)
    {
        return (fraction * 100d).ToString("0.#", CultureInfo.InvariantCulture) + "%";
    }

    private static IReadOnlyList<string> FormatPieDataLabelParts(double value, double total, int categoryIndex, ChartIndexedNumberPoint? workbookPoint, string? valueFormatCode, ChartDataLabelOptions options, ChartIndexedTextVector categoryLabels, IReadOnlyList<ChartSeriesNameRecord> seriesNames)
    {
        if (!string.IsNullOrWhiteSpace(options.CustomText))
        {
            return [options.CustomText];
        }

        var parts = new List<string>(4);
        string seriesName = GetActiveSeriesName(seriesNames, 0);
        if (options.ShowSeriesName && !string.IsNullOrWhiteSpace(seriesName))
        {
            parts.Add(seriesName);
        }

        string categoryLabel = GetIndexedCategoryLabel(categoryLabels, categoryIndex);
        if (options.ShowCategoryName && !string.IsNullOrWhiteSpace(categoryLabel))
        {
            parts.Add(categoryLabel);
        }

        if (options.ShowValue)
        {
            parts.Add(FormatChartDataLabelValue(value, options, workbookPoint, valueFormatCode));
        }

        if (options.ShowPercent)
        {
            parts.Add(FormatChartPercentageLabel(value / total));
        }

        return parts;
    }

    private static string JoinChartDataLabelParts(IReadOnlyList<string> parts, ChartDataLabelOptions options)
    {
        return string.Join(GetChartDataLabelSeparator(options), parts);
    }

    private static string FormatCartesianDataLabel(
        double value,
        int seriesIndex,
        int categoryIndex,
        ChartIndexedNumberPoint point,
        ChartIndexedNumberPoint? workbookPoint,
        string? valueFormatCode,
        ChartDataLabelOptions options,
        ChartIndexedTextVector categoryLabels,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames)
    {
        if (!string.IsNullOrWhiteSpace(options.CustomText))
        {
            return options.CustomText;
        }

        var parts = new List<string>(3);
        string seriesName = GetActiveSeriesName(seriesNames, seriesIndex);
        if (options.ShowSeriesName && !string.IsNullOrWhiteSpace(seriesName))
        {
            parts.Add(seriesName);
        }

        string categoryLabel = GetIndexedCategoryLabel(categoryLabels, categoryIndex);
        if (options.ShowCategoryName && !string.IsNullOrWhiteSpace(categoryLabel))
        {
            parts.Add(categoryLabel);
        }

        if (options.ShowValue)
        {
            parts.Add(FormatChartDataLabelValue(value, options, workbookPoint ?? point, valueFormatCode));
        }

        return string.Join(GetChartDataLabelSeparator(options), parts);
    }

    private static string GetIndexedCategoryLabel(ChartIndexedTextVector categoryLabels, int categoryIndex)
    {
        return (categoryLabels.Points ?? [])
            .Where(point => point.HasText && point.Index == categoryIndex)
            .Select(point => point.Text)
            .FirstOrDefault() ?? string.Empty;
    }

    private static string GetActiveSeriesName(IReadOnlyList<ChartSeriesNameRecord> seriesNames, int seriesIndex)
    {
        return seriesIndex >= 0 && seriesIndex < seriesNames.Count
            ? seriesNames[seriesIndex].ActiveName
            : string.Empty;
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options)
    {
        return FormatChartDataLabelValue(value, options.NumberFormatInfo, options.NumberFormat);
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options, ChartIndexedNumberPoint? sourcePoint)
    {
        return FormatChartDataLabelValue(value, options, sourcePoint, sourceFormatCode: null);
    }

    private static string FormatChartDataLabelValue(double value, ChartDataLabelOptions options, ChartIndexedNumberPoint? sourcePoint, string? sourceFormatCode)
    {
        if (ResolveSourceLinkedChartNumberFormatCode(options.NumberFormatInfo, sourcePoint) is { } sourceFormat)
        {
            return FormatChartNumber(value, sourceFormat);
        }

        return FormatChartDataLabelValue(value, options.NumberFormatInfo, options.NumberFormat, sourceFormatCode);
    }

    private static string FormatChartDataLabelValue(double value, ChartNumberFormat numberFormat, string legacyNumberFormat)
    {
        return FormatChartDataLabelValue(value, numberFormat, legacyNumberFormat, sourceFormatCode: null);
    }

    private static string FormatChartDataLabelValue(double value, ChartNumberFormat numberFormat, string legacyNumberFormat, string? sourceFormatCode)
    {
        if (IsRenderableChartNumberFormat(numberFormat))
        {
            return FormatChartNumber(value, numberFormat.FormatCode);
        }

        return !string.IsNullOrWhiteSpace(legacyNumberFormat) &&
            !string.Equals(legacyNumberFormat, "General", StringComparison.OrdinalIgnoreCase)
            ? FormatChartNumber(value, legacyNumberFormat)
            : IsRenderableChartFormatCode(sourceFormatCode)
                ? FormatChartNumber(value, sourceFormatCode)
            : FormatChartAxisLabel(value, null);
    }

    private static string? ResolveSourceLinkedChartNumberFormatCode(ChartNumberFormat numberFormat, ChartIndexedNumberPoint? sourcePoint)
    {
        if (sourcePoint is null)
        {
            return null;
        }

        ChartWorkbookRangeCell cell = sourcePoint.Value.WorkbookCell;
        return ResolveSourceLinkedChartNumberFormatCode(
            numberFormat,
            cell.StyleNumberFormatCode,
            cell.StyleAppliesNumberFormat,
            cell.StyleNumberFormatIsDateLike);
    }

    private static string? ResolveSourceLinkedChartNumberFormatCode(
        ChartNumberFormat numberFormat,
        string? workbookFormatCode,
        bool? workbookAppliesNumberFormat,
        bool workbookFormatIsDateLike)
    {
        if (numberFormat.SourceLinked != true ||
            workbookAppliesNumberFormat == false ||
            workbookFormatIsDateLike ||
            !IsRenderableChartFormatCode(workbookFormatCode))
        {
            return null;
        }

        return workbookFormatCode;
    }

    private static string GetChartDataLabelSeparator(ChartDataLabelOptions options)
    {
        return string.IsNullOrEmpty(options.Separator) ? ", " : options.Separator;
    }
}
