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
    private static async ValueTask RenderScatterDataLabelsAsync(
        PptxTheme theme,
        PptxColorMap colorMap,
        PdfGraphicsBuilder graphics,
        ChartPlotBox plotBox,
        IReadOnlyList<ScatterSeries> series,
        bool bubble,
        ChartValueExtents xExtents,
        ChartValueExtents yExtents,
        bool xReversed,
        bool yReversed,
        IReadOnlyList<ChartSeriesFill?> seriesFills,
        ChartDataLabelOptions labelOptions,
        IReadOnlyList<ChartDataLabelOptions> seriesLabelOptions,
        IReadOnlyList<ChartSeriesNameRecord> seriesNames,
        PresentationFontResolver? fontResolver,
        List<PdfFontResource> chartFonts,
        PptxRenderContext context,
        IReadOnlyDictionary<string, OoxRelationship>? chartRelationships,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        ChartBubbleSizing? bubbleSizing = null)
    {
        if ((!labelOptions.HasVisibleContent && !seriesLabelOptions.Any(options => options.HasVisibleContent)) || series.Count == 0)
        {
            return;
        }

        double maxBubbleSize = Math.Max(1d, series.SelectMany(item => item.Points).DefaultIfEmpty().Max(point => point.Size));
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        var runs = new List<TextRun>();
        List<ChartTextRunLink>? labelLinks = chartRelationships is null ? null : new List<ChartTextRunLink>();
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
                    (await textMeasurer.MeasureAsync(label, style).ConfigureAwait(false)) + legendKeyWidth + fontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor);
                (double pointX, double pointY, double radius) = ResolveScatterPointGeometry(
                    plotBox,
                    point,
                    bubble,
                    xExtents,
                    yExtents,
                    maxBubbleSize, xReversed, yReversed, bubbleSizing);
                // Office defaults scatter/bubble labels to the right. Bubble
                // placement uses the circle edge on the requested side; center
                // placement remains at the data point.
                PptxSceneChartDataLabelPosition position = effectiveOptions.PositionKind == PptxSceneChartDataLabelPosition.Unknown
                    ? PptxSceneChartDataLabelPosition.Right
                    : effectiveOptions.PositionKind;
                double anchorX = pointX;
                double anchorY = pointY;
                if (bubble)
                {
                    switch (position)
                    {
                        case PptxSceneChartDataLabelPosition.Left: anchorX -= radius; break;
                        case PptxSceneChartDataLabelPosition.Right: anchorX += radius; break;
                        case PptxSceneChartDataLabelPosition.Bottom: anchorY -= radius; break;
                        case PptxSceneChartDataLabelPosition.Center:
                        case PptxSceneChartDataLabelPosition.BestFit: break;
                        default: anchorY += radius; break;
                    }
                }
                (double labelX, double labelY, TextAlignment alignment) = ResolveLineDataLabelPosition(
                    position,
                    anchorX,
                    anchorY,
                    labelWidth,
                    labelHeight);
                ChartLayoutBox labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, labelWidth, labelHeight);
                RenderChartShapeStyle(graphics, labelBox.X, labelBox.Y, labelBox.Width, labelBox.Height, effectiveOptions.ShapeStyle);
                double textX = labelBox.X;
                double textWidth = labelBox.Width;
                if (effectiveOptions.ShowLegendKey)
                {
                    double consumedWidth = RenderFillDataLabelLegendKey(graphics, labelBox, fontSize, ChartSeriesColor(theme, colorMap, default, seriesIndex, seriesFills, 1d));
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

        RenderedFonts labelFonts = (await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CSD", fontResolver, diagnosticSink).ConfigureAwait(false));
        AddChartTextRunHyperlinkAnnotations(labelLinks, chartRelationships, context, labelFonts, linkAnnotations, reportedHyperlinkIds);

        double GetFillDataLabelLegendKeyWidth(double fontSize)
        {
            return fontSize * (PptxChartMetricRules.DataLabelLegendKeySizeFactor + PptxChartMetricRules.DataLabelLegendKeyTextGapFactor);
        }
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

        var parts = new List<string>(4);
        string seriesName = GetActiveSeriesName(seriesNames, seriesIndex);
        if (options.ShowSeriesName && !string.IsNullOrWhiteSpace(seriesName))
        {
            parts.Add(seriesName);
        }

        if (options.ShowCategoryName)
        {
            parts.Add(FormatChartDataLabelValue(point.X, options, point.XWorkbookPoint ?? point.XPoint, point.XFormatCode));
        }

        if (options.ShowValue)
        {
            parts.Add(FormatChartDataLabelValue(point.Y, options, point.YWorkbookPoint ?? point.YPoint, point.YFormatCode));
        }

        string FormatBubbleSizeDataLabelValue(double value, ScatterPoint point, ChartDataLabelOptions options)
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

        double? ResolveBubbleSizeValue()
        {
            return point.BubbleSizeWorkbookPoint?.Value ??
                point.BubbleSizePoint?.Value;
        }

        if (options.ShowBubbleSize && ResolveBubbleSizeValue() is { } bubbleSize)
        {
            parts.Add(FormatBubbleSizeDataLabelValue(bubbleSize, point, options));
        }

        return string.Join(GetChartDataLabelSeparator(options), parts);
    }

    private static (double X, double Y, double Radius) ResolveScatterPointGeometry(
        ChartPlotBox plotBox,
        ScatterPoint point,
        bool bubble,
        ChartValueExtents xExtents,
        ChartValueExtents yExtents,
        double maxBubbleSize,
        bool xReversed,
        bool yReversed,
        ChartBubbleSizing? bubbleSizing = null)
    {
        double xRange = xExtents.Max > xExtents.Min ? xExtents.Max - xExtents.Min : 1d;
        double yRange = yExtents.Max > yExtents.Min ? yExtents.Max - yExtents.Min : 1d;
        // Keep off-plot points outside the clip; clamping them would create
        // spurious edge markers and change the connecting path.
        double xRatio = (point.X - xExtents.Min) / xRange;
        double yRatio = (point.Y - yExtents.Min) / yRange;
        double pointX = plotBox.X + (xReversed ? 1d - xRatio : xRatio) * plotBox.Width;
        double pointY = plotBox.Y + (yReversed ? 1d - yRatio : yRatio) * plotBox.Height;
        double radius = bubble
            ? bubbleSizing is { } prepared
                ? Math.Sqrt(Math.Max(0d, point.Size) / prepared.MaximumSize) * prepared.Radius
                : Math.Sqrt(Math.Max(0d, point.Size) / Math.Max(1d, maxBubbleSize)) * Math.Min(plotBox.Width, plotBox.Height) * PptxChartMetricRules.BubbleRadiusPlotRatio
            : 3d;
        return (pointX, pointY, radius);
    }
}
