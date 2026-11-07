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
        ChartAreaPlotOptions? areaOptions = null, IReadOnlyList<ChartSeriesFill?>? areaSeriesFills = null, double? areaCrossingValue = null,
        double? chartRootFontSize = null)
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

                PptxSceneChartManualLayout layout = effectiveOptions.Layout;
                bool authoredSize = layout.Width.HasValue || layout.Height.HasValue;
                if (authoredSize)
                {
                    // Office rejects negative extents/out-of-frame edge origins as
                    // a layout declaration, retaining the automatic label. A zero
                    // or absent extent instead leaves no occupied label to paint.
                    if (layout.Width < 0d || layout.Height < 0d ||
                        IsManualLayoutEdgeMode(layout.XModeKind) && layout.X is < 0d or > 1d ||
                        IsManualLayoutEdgeMode(layout.YModeKind) && layout.Y is < 0d or > 1d)
                    {
                        layout = default;
                        effectiveOptions = effectiveOptions with { Layout = layout };
                    }
                    else if (layout.Width.GetValueOrDefault() == 0d || layout.Height.GetValueOrDefault() == 0d)
                        continue;
                }
                authoredSize = layout.Width.HasValue || layout.Height.HasValue;
                // Mixed edge extents retain the existing approximate fallback;
                // their Office anchor/extent coupling is independently unresolved.
                bool edgeSize = authoredSize && (IsManualLayoutEdgeMode(layout.WidthModeKind) || IsManualLayoutEdgeMode(layout.HeightModeKind));

                ChartTextStyle style = ResolveChartDataLabelTextStyle(theme, colorMap, effectiveOptions);
                double fontSize = style.FontSize;
                var lineBox = await textMeasurer.MeasureLineBoxAsync(style).ConfigureAwait(false);
                double descent = lineBox.Height - lineBox.Ascent;
                double labelHeight = edgeSize
                    ? fontSize * PptxChartMetricRules.CartesianDataLabelHeightFactor : lineBox.Height;
                bool hasPoint = densePointIndexes[seriesIndex].TryGetValue(i, out ChartIndexedNumberPoint point);
                string label = FormatCartesianDataLabel(value, seriesIndex, i, point, (workbookIndexes[seriesIndex].TryGetValue(hasPoint ? point.Index : i, out ChartIndexedNumberPoint workbookPoint) ? workbookPoint : null), series[seriesIndex].FormatCode, effectiveOptions, categoryLabelIndex, seriesNames);
                if (!string.IsNullOrEmpty(label) || effectiveOptions.ShowLegendKey)
                {
                    bool measured = areaOptions is null && CanMeasurePlainChartLabelFrame(effectiveOptions);
                    ChartLineAppearance? appearance = pointAppearance?.Invoke(seriesIndex, i) ?? seriesAppearance?[seriesIndex];
                    ChartMarkerStyle labelMarker = appearance?.Marker ?? ChartMarker(seriesIndex, markerStyles);
                    double pointMarkerRadius = areaOptions is null ? labelMarker.Size / 2d : 0d;
                    double keyScale = measured && chartRootFontSize is > 0d
                        ? PptxChartMetricRules.CategoryPointLabelKeyReferenceFontSize / chartRootFontSize.Value : 1d;
                    double markerLimit = measured ? Math.Round(PptxChartMarkerMetricRules.DataLabelMarkerSizeLimit / keyScale,
                        MidpointRounding.AwayFromZero) : PptxChartMarkerMetricRules.DataLabelMarkerSizeLimit;
                    labelMarker = labelMarker with { Size = PptxChartMarkerMetricRules.QuantizeMarkerExtent(Math.Min(labelMarker.Size, markerLimit)) };
                    // Implicit text bodies have visible fallback glyphs but no
                    // paragraph default-run metrics for their line key. Keep
                    // that key owner distinct from the measured glyph frame.
                    double keyLineHeight = measured && effectiveOptions.TextScope == PptxSceneChartPointTextScope.ImplicitTextBody
                        ? 0d : lineBox.Height;
                    double strokeKeyLength = keyLineHeight * PptxChartMetricRules.CategoryPointLabelStrokeKeyLengthFactor;
                    if (measured) strokeKeyLength *= keyScale;
                    double strokeKeyInset = fontSize * PptxChartMetricRules.CategoryPointLabelKeyInsetFactor;
                    double legendKeyWidth = effectiveOptions.ShowLegendKey
                        ? areaOptions is not null
                            ? fontSize * (PptxChartMetricRules.DataLabelLegendKeySizeFactor + PptxChartMetricRules.DataLabelLegendKeyTextGapFactor)
                            : strokeKeyInset + strokeKeyLength + PptxChartMetricRules.CategoryPointLabelKeyTextGap
                        : 0d;
                    if (measured && effectiveOptions.ShowLegendKey)
                    {
                        // The key's reference span includes symmetric space around
                        // its stroke. Root typography scales the key independently
                        // of the label's own font; stroke width does not move it.
                        legendKeyWidth = strokeKeyLength * PptxChartMetricRules.CategoryPointLabelKeyStrokeReserveFactor +
                            PptxChartMetricRules.DataLabelHorizontalPadding;
                        strokeKeyInset = (legendKeyWidth - strokeKeyLength) / 2d +
                            PptxChartMetricRules.CategoryPointLabelKeyCenterOffset;
                    }
                    double textPadding = edgeSize ? 0d : PptxChartMetricRules.DataLabelHorizontalPadding;
                    double verticalPadding = edgeSize ? 0d : PptxChartMetricRules.DataLabelVerticalPadding;
                    double cap = Math.Max(1d, (authoredSize ? Math.Min(1d, layout.Width.GetValueOrDefault()) * chartFrame.Width
                        : chartFrame.Width / 5d) - textPadding * 2d);
                    List<(string Text, double Width)>? lines = measured
                        ? await MeasureChartTextLinesAsync(label, cap, style, textMeasurer, context.CancellationToken,
                            effectiveOptions.TextBodyProperties.WrapValue != "none").ConfigureAwait(false) : null;
                    double lineAdvance = measured
                        ? (await textMeasurer.MeasureLineBoxAsync(style, horizontalHeader: true).ConfigureAwait(false)).Height : lineBox.Height;
                    double? lastAlignmentWidth = null;
                    if (lines is not null && authoredSize)
                    {
                        int limit = Math.Max(1, (int)Math.Min(lines.Count, Math.Max(0d,
                            Math.Floor((Math.Min(1d, layout.Height.GetValueOrDefault()) * chartFrame.Height - verticalPadding * 2d) / lineAdvance))));
                        if (limit < lines.Count)
                        {
                            var fitted = await FitChartLabelLineAsync(lines[limit - 1].Text, cap, style,
                                textMeasurer, context.CancellationToken).ConfigureAwait(false);
                            lines[limit - 1] = (fitted.Text, fitted.Width);
                            lines.RemoveRange(limit, lines.Count - limit);
                            lastAlignmentWidth = limit == 1 ? fitted.PrefixWidth : fitted.Width;
                        }
                    }
                    double textWidth = lines is null ? await textMeasurer.MeasureAsync(label, style).ConfigureAwait(false)
                        : Math.Min(cap, lines.Max(line => line.Width));
                    double paragraphExtraHeight = lines is null ? 0d : (lines.Count - 1) * lineAdvance;
                    double paragraphHeight = lineBox.Height + paragraphExtraHeight;
                    double occupiedHeight = authoredSize && !edgeSize
                        ? Math.Min(1d, layout.Height.GetValueOrDefault()) * chartFrame.Height : paragraphHeight + verticalPadding * 2d;
                    double verticalExpansion = (occupiedHeight - paragraphHeight - verticalPadding * 2d) / 2d;
                    double effectiveLabelWidth = edgeSize
                        ? Math.Max(labelWidth, textWidth + legendKeyWidth +
                            (effectiveOptions.ShowLegendKey ? fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor : 0d))
                        : authoredSize ? Math.Min(1d, layout.Width.GetValueOrDefault()) * chartFrame.Width + legendKeyWidth
                            : textWidth + textPadding * 2d + legendKeyWidth;
                    PptxSceneChartDataLabelPosition position = effectiveOptions.PositionKind;
                    if (position == PptxSceneChartDataLabelPosition.Unknown)
                        position = areaOptions is not null ? PptxSceneChartDataLabelPosition.Center
                            : edgeSize ? position : PptxSceneChartDataLabelPosition.Right;
                    (double labelX, double labelY, TextAlignment alignment) = edgeSize
                        ? ResolveLineDataLabelPosition(position, pointX, pointY, effectiveLabelWidth, labelHeight)
                        : ResolveCategoryPointLabelPosition(position, pointX, pointY, effectiveLabelWidth,
                            lineBox.Ascent, paragraphHeight, pointMarkerRadius, verticalExpansion);
                    if (authoredSize && !edgeSize)
                    {
                        labelX += legendKeyWidth / 2d;
                        alignment = TextAlignment.Center;
                    }
                    ChartLayoutBox labelBox = edgeSize
                        ? ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, effectiveLabelWidth, labelHeight)
                        : new ChartLayoutBox(labelX, labelY, effectiveLabelWidth, labelHeight);
                    ChartLayoutBox textBox = edgeSize ? labelBox : new ChartLayoutBox(
                        labelBox.X + legendKeyWidth, labelBox.Y - descent - verticalPadding - verticalExpansion - paragraphExtraHeight,
                        Math.Max(1d, labelBox.Width - legendKeyWidth), occupiedHeight);
                    if (!edgeSize && layout.HasLayout)
                    {
                        // Office applies factor offsets to automatic occupied bounds
                        // in the outer chart frame. Edge coordinates locate that
                        // frame's top-left; translate text, key, paint and links alike.
                        ChartLayoutBox occupied = new(labelBox.X, textBox.Y, labelBox.Width, textBox.Height);
                        ChartLayoutBox placement = occupied;
                        if (authoredSize)
                        {
                            // Admit the automatic sized box inside the chart before
                            // applying offsets. Authored dimensions belong to this
                            // box; the shared layout helper only places it.
                            placement = occupied with
                            {
                                X = Math.Clamp(occupied.X, chartFrame.X, Math.Max(chartFrame.X, chartFrame.X + chartFrame.Width - occupied.Width)),
                                Y = Math.Clamp(occupied.Y, chartFrame.Y, Math.Max(chartFrame.Y, chartFrame.Y + chartFrame.Height - occupied.Height))
                            };
                        }
                        PptxSceneChartManualLayout offsets = authoredSize ? layout with { Width = null, Height = null,
                            WidthModeKind = PptxSceneChartManualLayoutMode.Unknown, HeightModeKind = PptxSceneChartManualLayoutMode.Unknown } : layout;
                        if (TryBuildManualLayoutBox(offsets, chartFrame, placement,
                            out ChartLayoutBox manual, clampToFrame: false, missingPositionModesAreFactor: true))
                            placement = manual;
                        double dx = placement.X - occupied.X, dy = placement.Y - occupied.Y;
                        labelBox = labelBox with { X = labelBox.X + dx, Y = labelBox.Y + dy };
                        textBox = textBox with { X = textBox.X + dx, Y = textBox.Y + dy };
                    }
                    RenderChartShapeStyle(graphics, textBox.X, textBox.Y, textBox.Width, textBox.Height, effectiveOptions.ShapeStyle);
                    double textX = labelBox.X + legendKeyWidth + textPadding;
                    double textRunWidth = Math.Max(1d, labelBox.Width - legendKeyWidth - textPadding * 2d);
                    if (effectiveOptions.ShowLegendKey)
                    {
                        ChartLayoutBox keyBox = labelBox with
                        {
                            X = labelBox.X + (areaOptions is null ? strokeKeyInset : 0d),
                            Y = edgeSize ? labelBox.Y + lineBox.Ascent - lineBox.Height / 2d - labelBox.Height / 2d
                                : textBox.Y + textBox.Height - (keyLineHeight + verticalPadding * 2d) / 2d - labelBox.Height / 2d
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
                        if (lines is null && (!authoredSize || edgeSize)) alignment = TextAlignment.Left;
                    }

                    if (!string.IsNullOrEmpty(label))
                    {
                        if (lines is not null)
                        {
                            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++)
                            {
                                context.CancellationToken.ThrowIfCancellationRequested();
                                var fragment = lines[lineIndex];
                                double alignedWidth = lineIndex == lines.Count - 1 && lastAlignmentWidth is { } finalWidth
                                    ? finalWidth : fragment.Width;
                                double lineX = textX + (textRunWidth - alignedWidth) / 2d;
                                await AddChartLabelRunsAsync(runs, fragment.Text, effectiveOptions, lineX,
                                    labelBox.Y - lineIndex * lineAdvance, Math.Max(1d, fragment.Width), lineBox.Height,
                                    plotBox, style, TextAlignment.Left, fontResolver, labelLinks, descent).ConfigureAwait(false);
                            }
                        }
                        else await AddChartLabelRunsAsync(
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
                            fontResolver, labelLinks, edgeSize ? 0d : descent + verticalPadding + verticalExpansion).ConfigureAwait(false);
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
