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
        ChartFrameBox chartFrame,
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
                var lineBox = await textMeasurer.MeasureLineBoxAsync(style).ConfigureAwait(false);
                double descent = lineBox.Height - lineBox.Ascent;
                // Authored extents retain the existing sizing fallback. Offset-only
                // layouts translate the same measured frame as automatic labels.
                bool authoredSize = effectiveOptions.Layout.Width.HasValue || effectiveOptions.Layout.Height.HasValue;
                double labelHeight = authoredSize
                    ? fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor : lineBox.Height;
                bool hasPoint = densePointIndexes[seriesIndex].TryGetValue(i, out ChartIndexedNumberPoint point);
                string label = FormatCartesianDataLabel(value, seriesIndex, i, point, (workbookIndexes[seriesIndex].TryGetValue(hasPoint ? point.Index : i, out ChartIndexedNumberPoint workbookPoint) ? workbookPoint : null), series[seriesIndex].FormatCode, effectiveOptions, categoryLabelIndex, seriesNames);
                if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
                {
                    ChartLineAppearance? appearance = pointAppearance?.Invoke(seriesIndex, i) ?? seriesAppearance?[seriesIndex];
                    ChartMarkerStyle labelMarker = appearance?.Marker ?? ChartMarker(seriesIndex, markerStyles);
                    double pointMarkerRadius = areaOptions is null ? labelMarker.Size / 2d : 0d;
                    labelMarker = labelMarker with { Size = PptxChartMarkerMetricRules.QuantizeMarkerExtent(Math.Min(labelMarker.Size, PptxChartMarkerMetricRules.DataLabelMarkerSizeLimit)) };
                    double strokeKeyLength = lineBox.Height * PptxChartMetricRules.CategoryPointLabelStrokeKeyLengthFactor;
                    double strokeKeyInset = fontSize * PptxChartMetricRules.CategoryPointLabelKeyInsetFactor;
                    double legendKeyWidth = effectiveOptions.ShowLegendKey
                        ? areaOptions is not null
                            ? fontSize * (PptxChartMetricRules.DataLabelLegendKeySizeFactor + PptxChartMetricRules.DataLabelLegendKeyTextGapFactor)
                            : strokeKeyInset + strokeKeyLength + PptxChartMetricRules.CategoryPointLabelKeyTextGap
                        : 0d;
                    double textWidth = await textMeasurer.MeasureAsync(label, style).ConfigureAwait(false);
                    double textPadding = authoredSize ? 0d : PptxChartMetricRules.CategoryPointLabelHorizontalTextPadding;
                    double effectiveLabelWidth = authoredSize
                        ? Math.Max(labelWidth, textWidth + legendKeyWidth +
                            (effectiveOptions.ShowLegendKey ? fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor : 0d))
                        : textWidth + textPadding * 2d + legendKeyWidth;
                    PptxSceneChartDataLabelPosition position = effectiveOptions.PositionKind;
                    if (position == PptxSceneChartDataLabelPosition.Unknown)
                        position = areaOptions is not null ? PptxSceneChartDataLabelPosition.Center
                            : authoredSize ? position : PptxSceneChartDataLabelPosition.Right;
                    (double labelX, double labelY, TextAlignment alignment) = authoredSize
                        ? ResolveLineDataLabelPosition(position, pointX, pointY, effectiveLabelWidth, labelHeight)
                        : ResolveCategoryPointLabelPosition(position, pointX, pointY, effectiveLabelWidth,
                            lineBox.Ascent, lineBox.Height, pointMarkerRadius, fontSize);
                    ChartLayoutBox labelBox = authoredSize
                        ? ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, effectiveLabelWidth, labelHeight)
                        : new ChartLayoutBox(labelX, labelY, effectiveLabelWidth, labelHeight);
                    double verticalPadding = authoredSize ? 0d : PptxChartMetricRules.CategoryPointLabelVerticalTextPadding;
                    ChartLayoutBox textBox = authoredSize ? labelBox : new ChartLayoutBox(
                        labelBox.X + legendKeyWidth, labelBox.Y - descent - verticalPadding,
                        Math.Max(1d, labelBox.Width - legendKeyWidth), lineBox.Height + verticalPadding * 2d);
                    if (!authoredSize && effectiveOptions.Layout.HasLayout)
                    {
                        // Office applies factor offsets to automatic occupied bounds
                        // in the outer chart frame. Edge coordinates locate that
                        // frame's top-left; translate text, key, paint and links alike.
                        ChartLayoutBox occupied = new(labelBox.X, textBox.Y, labelBox.Width, textBox.Height);
                        if (TryBuildManualLayoutBox(effectiveOptions.Layout, chartFrame, occupied,
                            out ChartLayoutBox manual, clampToFrame: false, missingPositionModesAreFactor: true))
                        {
                            double dx = manual.X - occupied.X, dy = manual.Y - occupied.Y;
                            labelBox = labelBox with { X = labelBox.X + dx, Y = labelBox.Y + dy };
                            textBox = textBox with { X = textBox.X + dx, Y = textBox.Y + dy };
                        }
                    }
                    RenderChartShapeStyle(graphics, textBox.X, textBox.Y, textBox.Width, textBox.Height, effectiveOptions.ShapeStyle);
                    double textX = labelBox.X + legendKeyWidth + textPadding;
                    double textRunWidth = Math.Max(1d, labelBox.Width - legendKeyWidth - textPadding * 2d);
                    if (effectiveOptions.ShowLegendKey)
                    {
                        ChartLayoutBox keyBox = labelBox with
                        {
                            X = labelBox.X + (areaOptions is null ? strokeKeyInset : 0d),
                            Y = labelBox.Y + lineBox.Ascent - lineBox.Height / 2d - labelBox.Height / 2d
                        };
                        if (areaOptions is not null)
                        {
                            RenderFillDataLabelLegendKey(graphics, keyBox, fontSize,
                                ChartSeriesColor(theme, colorMap, chartPalette, seriesIndex, areaSeriesFills ?? [], 1d));
                        }
                        else
                        {
                            ChartSeriesStroke stroke = appearance?.Stroke ?? ChartSeriesStrokeColor(theme, colorMap, chartPalette, seriesIndex, seriesStrokes, 1.2d);
                            ChartMarkerStyle marker = labelMarker;
                            if (appearance is { } pointStyle && marker.Stroke is null)
                                marker = marker with { Stroke = new ChartSeriesStroke(pointStyle.MarkerOutlineColor, 1d, pointStyle.MarkerOutlineWidth ?? PptxChartMarkerMetricRules.DefaultMarkerOutlineWidth) with { Join = 1 } };
                            RenderStrokeMarkerDataLabelLegendKey(graphics, keyBox, marker, stroke, strokeKeyLength, appearance?.MarkerColor);
                        }
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
                            textRunWidth,
                            textBox.Height,
                            plotBox,
                            style,
                            alignment,
                            fontResolver, labelLinks, authoredSize ? 0d : descent + verticalPadding).ConfigureAwait(false);
                    }
                }
            }
        }

        RenderedFonts labelFonts = (await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CLD", fontResolver, diagnosticSink).ConfigureAwait(false));
        AddChartTextRunHyperlinkAnnotations(labelLinks, chartRelationships, context, labelFonts, linkAnnotations, reportedHyperlinkIds);

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

    private static void RenderStrokeMarkerDataLabelLegendKey(PdfGraphicsBuilder graphics, ChartLayoutBox labelBox, ChartMarkerStyle marker, ChartSeriesStroke stroke, double segmentLength, RgbColor? markerColor = null)
    {
        double segmentY = labelBox.Y + labelBox.Height / 2d;
        SetChartStroke(graphics, stroke);
        graphics.StrokeLine(labelBox.X, segmentY, labelBox.X + segmentLength, segmentY);
        DrawChartMarker(graphics, labelBox.X + segmentLength / 2d, segmentY, marker, markerColor ?? stroke.Color, stroke.Color);
    }
}
