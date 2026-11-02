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
    private static async ValueTask RenderCategoryPointDataLabelsAsync(
        PptxTheme theme,
        PptxColorMap colorMap,
        ChartColorPalette chartPalette,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<ChartIndexedNumberVector> series,
        ChartValueExtents extents,
        bool valueAxisReversed,
        bool categoryPointsOnTickMarks,
        IReadOnlyList<ChartSeriesStroke?> seriesStrokes,
        IReadOnlyList<ChartMarkerStyle> markerStyles,
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
        Func<int, int, ChartLineAppearance>? pointAppearance = null, IReadOnlyList<ChartLineAppearance>? seriesAppearance = null, bool categoryAxisReversed = false,
        ChartAreaPlotOptions? areaOptions = null, IReadOnlyList<ChartSeriesFill?>? areaSeriesFills = null, double? areaCrossingValue = null)
    {
        if ((!labelOptions.HasVisibleContent && !seriesLabelOptions.Any(options => options.HasVisibleContent)) || series.Count == 0)
        {
            return;
        }

        IReadOnlyList<IReadOnlyList<double?>> denseValueSeries = DensifyChartValueSeries(series);
        if (denseValueSeries.Count == 0)
        {
            return;
        }

        // index workbook points once per series instead of filtering plus
        // linear-scanning per rendered label below.
        Dictionary<int, ChartIndexedNumberPoint>[] workbookIndexes = series.Select(BuildWorkbookPointIndex).ToArray();

        // RV20: index dense-slot provenance once per series (last-wins, like
        // DensePoints) instead of materializing dense arrays per label below.
        Dictionary<int, ChartIndexedNumberPoint>[] densePointIndexes = series.Select(BuildDensePointIndex).ToArray();

        // RV14: index category labels once per frame instead of linear-scanning
        // per rendered label below.
        Dictionary<int, string> categoryLabelIndex = BuildCategoryLabelIndex(categoryLabels);

        int pointCount = Math.Max(1, denseValueSeries.Max(values => values.Count));
        double[]? areaLower = areaOptions is { Stacked: true } ? new double[pointCount] : null;
        double crossingY = ChartValueToPlotCoordinate(extents, areaCrossingValue, plotBox.Y, plotBox.Height, valueAxisReversed);
        double labelWidth = Math.Max(
            PptxChartMetricRules.CartesianDataLabelMinimumWidth,
            plotBox.Width / Math.Max(PptxChartMetricRules.LineDataLabelMinimumPointSpan, pointCount * PptxChartMetricRules.LineDataLabelPointWidthFactor));
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        var runs = new List<TextRun>();
        List<ChartTextRunLink>? labelLinks = chartRelationships is null ? null : new List<ChartTextRunLink>();
        for (int seriesIndex = 0; seriesIndex < denseValueSeries.Count; seriesIndex++)
        {
            IReadOnlyList<double?> values = denseValueSeries[seriesIndex];
            for (int i = 0; i < values.Count; i++)
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                double value;
                if (values[i] is { } sourceValue) value = sourceValue;
                else if (areaOptions is { } missingArea && (missingArea.Stacked || missingArea.DisplayBlanksAs == PptxSceneChartDisplayBlanksAs.Zero))
                {
                    // Zero-valued logical slots are visible labels in Office;
                    // retain the absent cached point in source inspection.
                    value = 0d;
                }
                else continue;

                double pointX = ChartCategoryToPlotCoordinate(i, pointCount, plotBox.X, plotBox.Width, categoryPointsOnTickMarks, categoryAxisReversed);
                double pointY = ChartValueToPlotCoordinate(extents, value, plotBox.Y, plotBox.Height, valueAxisReversed);
                if (areaOptions is { } area)
                {
                    var span = ResolveAreaPointVerticalSpan(plotBox, extents, area, valueAxisReversed,
                        seriesIndex, value, areaLower?[i] ?? 0d,
                        GetCategoryPositiveTotal(denseValueSeries, i, area.PercentStacked), crossingY);
                    pointY = (span.LowerY + span.UpperY) / 2d;
                    if (areaLower is not null) areaLower[i] = span.UpperValue;
                }
                ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(ResolveChartDataLabelOptionsForSeries(labelOptions, seriesLabelOptions, seriesIndex), i);
                if (!effectiveOptions.HasVisibleContent)
                {
                    continue;
                }

                ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                double fontSize = style.FontSize;
                double labelHeight = fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor;
                bool hasPoint = densePointIndexes[seriesIndex].TryGetValue(i, out ChartIndexedNumberPoint point);
                string label = FormatCartesianDataLabel(value, seriesIndex, i, point, (workbookIndexes[seriesIndex].TryGetValue(hasPoint ? point.Index : i, out ChartIndexedNumberPoint workbookPoint) ? workbookPoint : null), series[seriesIndex].FormatCode, effectiveOptions, categoryLabelIndex, seriesNames);
                if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
                {
                    ChartLineAppearance? appearance = pointAppearance?.Invoke(seriesIndex, i) ?? seriesAppearance?[seriesIndex];
                    ChartMarkerStyle labelMarker = appearance?.Marker ?? ChartMarker(seriesIndex, markerStyles);
                    labelMarker = labelMarker with { Size = Math.Min(labelMarker.Size, PptxChartMarkerMetricRules.DataLabelMarkerSizeLimit) };
                    double legendKeyWidth = effectiveOptions.ShowLegendKey
                        ? areaOptions is not null
                            ? fontSize * (PptxChartMetricRules.DataLabelLegendKeySizeFactor + PptxChartMetricRules.DataLabelLegendKeyTextGapFactor)
                            : GetStrokeMarkerDataLabelLegendKeyWidth(fontSize, labelMarker)
                        : 0d;
                    double effectiveLabelWidth = Math.Max(labelWidth,
                        (await textMeasurer.MeasureAsync(label, style).ConfigureAwait(false)) + legendKeyWidth +
                        (effectiveOptions.ShowLegendKey ? fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor : 0d));
                    (double labelX, double labelY, TextAlignment alignment) = ResolveLineDataLabelPosition(
                        areaOptions is not null && effectiveOptions.PositionKind == PptxSceneChartDataLabelPosition.Unknown
                            ? PptxSceneChartDataLabelPosition.Center : effectiveOptions.PositionKind,
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
                        double consumedWidth;
                        if (areaOptions is not null)
                        {
                            consumedWidth = RenderFillDataLabelLegendKey(graphics, labelBox, fontSize,
                                ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, areaSeriesFills ?? [], 1d));
                        }
                        else
                        {
                            ChartSeriesStroke stroke = appearance?.Stroke ?? ChartSeriesStrokeColor(theme, colorMap, chartPalette, seriesIndex, seriesStrokes, 1.2d);
                            ChartMarkerStyle marker = labelMarker;
                            if (appearance is { } pointStyle && marker.Stroke is null)
                                marker = marker with { Stroke = new ChartSeriesStroke(pointStyle.MarkerOutlineColor, 1d, pointStyle.MarkerOutlineWidth ?? PptxChartMarkerMetricRules.DefaultMarkerOutlineWidth) with { Join = 1 } };
                            consumedWidth = RenderStrokeMarkerDataLabelLegendKey(graphics, labelBox, fontSize, marker, stroke, appearance?.MarkerColor);
                        }
                        textX += consumedWidth;
                        textWidth = Math.Max(1d, textWidth - consumedWidth);
                        alignment = TextAlignment.Left;
                    }

                    if (!string.IsNullOrEmpty(label))
                    {
                        await AddChartLabelRunsAsync(
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
                            fontResolver, labelLinks).ConfigureAwait(false);
                    }
                }
            }
        }

        RenderedFonts labelFonts = (await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CLD", fontResolver, diagnosticSink).ConfigureAwait(false));
        AddChartTextRunHyperlinkAnnotations(labelLinks, chartRelationships, context, labelFonts, linkAnnotations, reportedHyperlinkIds);

        double GetStrokeMarkerDataLabelLegendKeyWidth(double fontSize, ChartMarkerStyle marker)
        {
            return Math.Max(marker.Size * 2d, fontSize * 1.05d) + fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
        }
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

    private static double RenderStrokeMarkerDataLabelLegendKey(PdfGraphicsBuilder graphics, ChartLayoutBox labelBox, double fontSize, ChartMarkerStyle marker, ChartSeriesStroke stroke, RgbColor? markerColor = null)
    {
        double segmentLength = Math.Max(marker.Size * 2d, fontSize * 1.05d);
        double segmentY = labelBox.Y + labelBox.Height / 2d;
        SetChartStroke(graphics, stroke);
        graphics.StrokeLine(labelBox.X, segmentY, labelBox.X + segmentLength, segmentY);
        DrawChartMarker(graphics, labelBox.X + segmentLength / 2d, segmentY, marker, markerColor ?? stroke.Color, stroke.Color);
        return segmentLength + fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
    }
}
