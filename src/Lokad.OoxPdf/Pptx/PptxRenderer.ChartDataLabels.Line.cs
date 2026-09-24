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
    private static void RenderLineDataLabels(
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
        PresentationFontResolver? fontResolver,
        List<PdfFontResource> chartFonts,
        PptxRenderContext context,
        IReadOnlyDictionary<string, OoxRelationship>? chartRelationships,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
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
                if (values[i] is not double value)
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
                ChartIndexedNumberPoint point = densePointIndexes[seriesIndex].TryGetValue(i, out ChartIndexedNumberPoint densePoint) ? densePoint : default;
                string label = FormatCartesianDataLabel(value, seriesIndex, i, point, (workbookIndexes[seriesIndex].TryGetValue(point.Index, out ChartIndexedNumberPoint workbookPoint) ? workbookPoint : null), series[seriesIndex].FormatCode, effectiveOptions, categoryLabelIndex, seriesNames);
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
                            fontResolver, labelLinks);
                    }
                }
            }
        }

        RenderedFonts labelFonts = RenderChartTextRuns(runs, graphics, chartFonts, "CLD", fontResolver, diagnosticSink);
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

    private static double RenderStrokeMarkerDataLabelLegendKey(PdfGraphicsBuilder graphics, ChartLayoutBox labelBox, double fontSize, ChartMarkerStyle marker, ChartSeriesStroke stroke)
    {
        double segmentLength = Math.Max(marker.Size * 2d, fontSize * 1.05d);
        double segmentY = labelBox.Y + labelBox.Height / 2d;
        SetChartStroke(graphics, stroke);
        graphics.StrokeLine(labelBox.X, segmentY, labelBox.X + segmentLength, segmentY);
        DrawChartMarker(graphics, labelBox.X + segmentLength / 2d, segmentY, marker, stroke.Color, stroke.Color);
        return segmentLength + fontSize * PptxChartMetricRules.DataLabelLegendKeyTextGapFactor;
    }
}
