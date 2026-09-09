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
    private static void RenderPieDataLabels(PptxTheme theme, PptxColorMap colorMap, PdfGraphicsBuilder graphics, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<ChartIndexedPieSlice> slices, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, string? valueFormatCode, ChartDataLabelOptions labelOptions, ChartIndexedTextVector categoryLabels, IReadOnlyList<ChartSeriesNameRecord> seriesNames, PresentationFontResolver? fontResolver, List<PdfFontResource> chartFonts, PptxRenderContext context, IReadOnlyDictionary<string, OoxRelationship>? chartRelationships, List<PdfLinkAnnotation> linkAnnotations, HashSet<string> reportedHyperlinkIds, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        if (!labelOptions.HasVisibleContent || slices.Count == 0)
        {
            return;
        }

        double total = slices.Sum(slice => slice.Value);
        if (total <= 0d)
        {
            return;
        }

        ChartPlotBox plotBox = layout.PlotBox;
        ChartPolarGeometry geometry = layout.Geometry;
        double labelRadius = geometry.Radius * (holeSize > 0d ? Math.Max(PptxChartMetricRules.PieDataLabelRadiusRatio, (1d + holeSize) / 2d) : PptxChartMetricRules.PieDataLabelRadiusRatio);
        double labelWidth = Math.Max(PptxChartMetricRules.PieDataLabelMinimumWidth, geometry.Radius * PptxChartMetricRules.PieDataLabelWidthRatio);
        var runs = new List<TextRun>(slices.Count);
        List<ChartTextRunLink>? labelLinks = chartRelationships is null ? null : new List<ChartTextRunLink>();
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
            IReadOnlyList<string> labelParts = FormatPieDataLabelParts(slice.Value, slice.Index, slice.WorkbookPoint, effectiveOptions);
            string JoinChartDataLabelParts()
            {
                return string.Join(GetChartDataLabelSeparator(effectiveOptions), labelParts);
            }

            string label = JoinChartDataLabelParts();
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
                    AddPolarChartLabelRuns(labelParts, label, effectiveOptions, textX, labelBox.Y, textWidth, labelBox.Height, style, alignment);
                }
            }
            angle += sweep;
        }
        RenderedFonts labelFonts = RenderChartTextRuns(runs, graphics, chartFonts, "CP", fontResolver, diagnosticSink);
        AddChartTextRunHyperlinkAnnotations(labelLinks, chartRelationships, context, labelFonts, linkAnnotations, reportedHyperlinkIds);

        IReadOnlyList<string> FormatPieDataLabelParts(double value, int categoryIndex, ChartIndexedNumberPoint? workbookPoint, ChartDataLabelOptions options)
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

            string FormatChartPercentageLabel(double fraction)
            {
                return (fraction * 100d).ToString("0.#", CultureInfo.InvariantCulture) + "%";
            }

            if (options.ShowPercent)
            {
                parts.Add(FormatChartPercentageLabel(value / total));
            }

            return parts;
        }

        void AddPolarChartLabelRuns(IReadOnlyList<string> parts, string fallbackText, ChartDataLabelOptions options, double x, double y, double width, double height, ChartTextStyle style, TextAlignment alignment)
        {
            var textMeasurer = new ChartTextMeasurer(fontResolver);
            ChartLayoutBox clipBox = ResolveDataLabelTextClipBox(plotBox, options, x, y, width, height);
            bool ShouldSplitPolarDataLabelParts()
            {
                string separator = GetChartDataLabelSeparator(options);
                return separator.Length > 0 && separator.All(char.IsWhiteSpace);
            }

            if (parts.Count <= 1 || options.CustomTextRuns.Count > 0 || !ShouldSplitPolarDataLabelParts())
            {
                AddChartLabelRuns(runs, fallbackText, options, x, y, width, height, plotBox, style, alignment, fontResolver, labelLinks);
                return;
            }

            ChartTextRunLayout[] labelRuns = parts
                .Where(part => !string.IsNullOrWhiteSpace(part))
                .Select(part => new ChartTextRunLayout(part, style, Math.Max(0d, textMeasurer.Measure(part, style))))
                .Where(run => run.Width > 0d)
                .ToArray();
            if (labelRuns.Length <= 1)
            {
                AddChartLabelRuns(runs, fallbackText, options, x, y, width, height, plotBox, style, alignment, fontResolver, labelLinks);
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
}
