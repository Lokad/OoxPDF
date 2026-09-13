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
    // Start angle for pie data labels in math convention (y-up): top, shifted clockwise
    // by the first-slice angle, exactly like the slice loop.
    private static double GetPieDataLabelStartAngle(double firstSliceAngle)
    {
        return Math.PI / 2d - firstSliceAngle * Math.PI / 180d;
    }

    // Midpoint angle for a pie data label: start minus half the clockwise share, mirroring
    // the slice loop exactly.
    private static double GetPieDataLabelMidpointAngle(double startAngleRadians, double sliceValue, double total)
    {
        return startAngleRadians - sliceValue / total * Math.PI;
    }

    private static void RenderPieDataLabels(PptxTheme theme, PptxColorMap colorMap, PdfGraphicsBuilder graphics, IReadOnlyList<RgbColor>? chartPalette, ChartPolarLayout layout, IReadOnlyList<ChartIndexedPieSlice> slices, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, IReadOnlyDictionary<int, double> pointExplosions, double holeSize, double firstSliceAngle, string? valueFormatCode, ChartDataLabelOptions labelOptions, ChartIndexedTextVector categoryLabels, IReadOnlyList<ChartSeriesNameRecord> seriesNames, PresentationFontResolver? fontResolver, List<PdfFontResource> chartFonts, PptxRenderContext context, IReadOnlyDictionary<string, OoxRelationship>? chartRelationships, List<PdfLinkAnnotation> linkAnnotations, HashSet<string> reportedHyperlinkIds, Action<OoxPdfDiagnostic>? diagnosticSink = null)
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
        double labelRadius = geometry.Radius * (holeSize > 0d ? Math.Max(PptxChartMetricRules.PieDataLabelRadiusRatio, (1d + holeSize) / 2d) : PptxChartMetricRules.PieAutoDataLabelRadiusRatio);
        double labelWidth = Math.Max(PptxChartMetricRules.PieDataLabelMinimumWidth, geometry.Radius * PptxChartMetricRules.PieDataLabelWidthRatio);
        var runs = new List<TextRun>(slices.Count);
        List<ChartTextRunLink>? labelLinks = chartRelationships is null ? null : new List<ChartTextRunLink>();
        double angle = GetPieDataLabelStartAngle(firstSliceAngle);
        // Office draws leader lines only for same-side narrow manual pie labels (the
        // box center sits on the wedge half and the box fits 85.3pt): one leader on
        // the base probes (Gamma/West), two on rotated small-factor probes
        // (Beta+Gamma), none when the same-side box runs wide (graded/long-Gamma).
        // Charts without any valid manual label keep the legacy per-label behavior.
        bool pieChartHasManualLeaderCandidate = HasPieManualLeaderCandidate(labelOptions, slices);
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
            double sweep = -slice.Value / total * Math.PI * 2d;
            double mid = GetPieDataLabelMidpointAngle(angle, slice.Value, total);
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
                // Factor-mode manuals resolve on a circle past the rim with content-tight
                // boxes clamped into the plot; out-of-range manuals fall back to auto,
                // edge-mode and layout-free labels keep the legacy path.
                List<ChartTextRunLayout[]>? pieWrapLines = null;
                double pieWrapLongest = 0d;
                string pieWrapSeparator = GetChartDataLabelSeparator(effectiveOptions);
                if (labelParts.Count > 1 && effectiveOptions.CustomTextRuns.Count == 0 && pieWrapSeparator.Length > 0 && pieWrapSeparator.All(char.IsWhiteSpace))
                {
                    var pieWrapMeasurer = new ChartTextMeasurer(fontResolver);
                    var pieWrapTexts = new List<string>(labelParts.Count);
                    var pieWrapWidths = new List<double>(labelParts.Count);
                    var pieWrapSeps = new List<double>(labelParts.Count);
                    double pieWrapSepWidth = pieWrapMeasurer.Measure(pieWrapSeparator, style);
                    double pieWrapCapWidth = plotBox.Width * PptxChartMetricRules.PieDataLabelWrapWidthFactor;
                    foreach (string part in labelParts)
                    {
                        if (string.IsNullOrWhiteSpace(part))
                        {
                            continue;
                        }
                        double partWidth = Math.Max(0d, pieWrapMeasurer.Measure(part, style));
                        if (partWidth <= 0d)
                        {
                            continue;
                        }
                        if (partWidth > pieWrapCapWidth && part.Length > 1)
                        {
                            List<int> pieChunks = SplitPieLabelLongWord(MeasurePieLabelChars(part, style, pieWrapMeasurer), pieWrapCapWidth);
                            int pieOffset = 0;
                            foreach (int pieChunkLength in pieChunks)
                            {
                                pieWrapTexts.Add(part.Substring(pieOffset, pieChunkLength));
                                pieWrapWidths.Add(0d);
                                pieWrapSeps.Add(0d);
                                pieOffset += pieChunkLength;
                            }
                            pieWrapSeps[^1] = pieWrapSepWidth;
                        }
                        else
                        {
                            pieWrapTexts.Add(part);
                            pieWrapWidths.Add(partWidth);
                            pieWrapSeps.Add(pieWrapSepWidth);
                        }
                    }
                    if (pieWrapSeps.Count > 0)
                    {
                        pieWrapSeps[^1] = 0d;
                    }
                    for (int pieMeasureIndex = 0; pieMeasureIndex < pieWrapTexts.Count; pieMeasureIndex++)
                    {
                        if (pieWrapWidths[pieMeasureIndex] <= 0d)
                        {
                            pieWrapWidths[pieMeasureIndex] = Math.Max(0d, pieWrapMeasurer.Measure(pieWrapTexts[pieMeasureIndex], style));
                        }
                    }
                    List<int[]> pieSplitLines = SplitPieLabelWordLines(pieWrapWidths, pieWrapSeps, pieWrapCapWidth);
                    if (pieSplitLines.Count > 1)
                    {
                        var pieBuiltLines = new List<ChartTextRunLayout[]>(pieSplitLines.Count);
                        foreach (int[] pieLine in pieSplitLines)
                        {
                            var pieBuiltRuns = new List<ChartTextRunLayout>(pieLine.Length);
                            double pieLineWidth = 0d;
                            for (int piePosition = 0; piePosition < pieLine.Length; piePosition++)
                            {
                                int pieIndex = pieLine[piePosition];
                                pieBuiltRuns.Add(new ChartTextRunLayout(pieWrapTexts[pieIndex], style, pieWrapWidths[pieIndex]));
                                pieLineWidth += pieWrapWidths[pieIndex] + (piePosition < pieLine.Length - 1 ? pieWrapSeps[pieIndex] : 0d);
                            }
                            pieBuiltLines.Add(pieBuiltRuns.ToArray());
                            pieWrapLongest = Math.Max(pieWrapLongest, pieLineWidth);
                        }
                        pieWrapLines = pieBuiltLines;
                    }
                }
                ChartLayoutBox labelBox;
                bool pieManualBottomAnchor = false;
                double pieUnwrappedWidth = 0d;
                int pieManualLineCount = 1;
                bool pieIsFactorManual = TryGetPieManualLeaderFactorX(effectiveOptions.Layout, out _);
                if (pieIsFactorManual)
                {
                    double pieCircleGap = plotBox.Width * PptxChartMetricRules.PieManualLabelCircleGapPlotWidthFactor;
                    double pieCircleX = geometry.CenterX + Math.Cos(mid) * (geometry.Radius + explosion + pieCircleGap);
                    double pieCircleY = geometry.CenterY + Math.Sin(mid) * (geometry.Radius + explosion + pieCircleGap);
                    pieUnwrappedWidth = MeasurePieLabelPartsWidth(labelParts, pieWrapSeparator, style, fontResolver);
                    double pieContentWidth = (pieWrapLines is not null && pieWrapLines.Count > 1 && pieWrapLongest > 0d)
                        ? pieWrapLongest
                        : pieUnwrappedWidth;
                    pieManualLineCount = pieWrapLines is not null && pieWrapLines.Count > 1 ? pieWrapLines.Count : 1;
                    double pieManualWidth = pieContentWidth + 2d * PptxChartMetricRules.PieManualLabelBoxSidePad;
                    double pieManualHeight = ComputePieManualLabelBoxHeight(fontSize, pieManualLineCount);
                    double pieSingleHeight = ComputePieManualLabelBoxHeight(fontSize, 1);
                    labelBox = ResolvePieManualDataLabelBox(plotBox, effectiveOptions.Layout, pieCircleX, pieCircleY, Math.Sin(mid), Math.Cos(mid), geometry.CenterX, pieManualWidth, pieManualHeight, pieSingleHeight);
                    pieManualBottomAnchor = true;
                }
                else if (effectiveOptions.Layout.HasLayout && (effectiveOptions.Layout.XModeKind == PptxSceneChartManualLayoutMode.Edge || effectiveOptions.Layout.YModeKind == PptxSceneChartManualLayoutMode.Edge))
                {
                    labelBox = ResolveDataLabelBox(plotBox, effectiveOptions, labelX, labelY, labelWidth, labelHeight);
                }
                else
                {
                    labelBox = new ChartLayoutBox(labelX, labelY, labelWidth, labelHeight);
                }
                bool pieDrawLeader = !pieChartHasManualLeaderCandidate
                    || (pieIsFactorManual && ShouldDrawPieManualLeaderLabel(labelBox.X + labelBox.Width / 2d, geometry.CenterX, Math.Cos(mid), pieUnwrappedWidth, plotBox.Width * PptxChartMetricRules.PieDataLabelWrapWidthFactor, plotBox.Width, plotBox.Height));
                if (pieDrawLeader)
                {
                    RenderPieDataLabelLeaderLine(graphics, geometry, mid, explosion, labelBox, effectiveOptions, fontSize, pieManualLineCount, pieManualBottomAnchor);
                }
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
                double pieManualTextX = textX;
                if (pieManualBottomAnchor && pieWrapLines is not null && pieWrapLines.Count > 1)
                {
                    pieManualTextX = textX + PptxChartMetricRules.PieManualLabelBoxSidePad;
                }
                    AddPolarChartLabelRuns(labelParts, label, effectiveOptions, pieManualTextX, labelBox.Y, textWidth, labelBox.Height, style, alignment, pieWrapLines, pieWrapLongest, pieManualBottomAnchor);
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

        void AddPolarChartLabelRuns(IReadOnlyList<string> parts, string fallbackText, ChartDataLabelOptions options, double x, double y, double width, double height, ChartTextStyle style, TextAlignment alignment, List<ChartTextRunLayout[]>? wrapLines, double wrapLongest, bool bottomAnchor = false)
        {
            var textMeasurer = new ChartTextMeasurer(fontResolver);
            ChartLayoutBox clipBox = ResolveDataLabelTextClipBox(plotBox, options, x, y, width, height);
            if (bottomAnchor)
            {
                int pieAnchorLines = wrapLines is not null && wrapLines.Count > 1 ? wrapLines.Count : 1;
                y += ComputePieManualLabelBaselinePad(style.FontSize)
                    + style.FontSize * PptxChartMetricRules.PieDataLabelLinePitchFactor * (pieAnchorLines - 1);
            }
            else
            {
                y += style.FontSize * PptxChartMetricRules.PieDataLabelBaselineFactor;
            }
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
            clipBox = ExpandPieLabelClipToText(clipBox, x, y, width, height, totalWidth);
            if (wrapLines is not null && wrapLines.Count > 1)
            {
                clipBox = ExpandPieLabelClipToText(clipBox, x, y, width, height + (wrapLines.Count - 1) * height, totalWidth);
                double lineY = y;
                foreach (ChartTextRunLayout[] wrapLine in wrapLines)
                {
                    double lineWidth = separatorWidth * Math.Max(0, wrapLine.Length - 1);
                    foreach (ChartTextRunLayout wrapRun in wrapLine)
                    {
                        lineWidth += wrapRun.Width;
                    }
                    double lineCursor = alignment switch
                    {
                        TextAlignment.Right => x + Math.Max(1d, wrapLongest) - lineWidth,
                        TextAlignment.Center => x + (Math.Max(1d, wrapLongest) - lineWidth) / 2d,
                        _ => x
                    };
                    for (int wrapPosition = 0; wrapPosition < wrapLine.Length; wrapPosition++)
                    {
                        ChartTextRunLayout wrapRun = wrapLine[wrapPosition];
                        double wrapRunWidth = Math.Max(0.1d, wrapRun.Width);
                        runs.Add(CreateChartTextRun(wrapRun.Text, lineCursor, lineY, wrapRunWidth, height, clipBox.X, clipBox.Y, clipBox.Width, clipBox.Height, wrapRun.Style, TextAlignment.Left) with { PreventCoalesce = true });
                        lineCursor += wrapRunWidth + (wrapPosition < wrapLine.Length - 1 ? separatorWidth : 0d);
                    }
                    lineY -= style.FontSize * PptxChartMetricRules.PieDataLabelLinePitchFactor;
                }
                return;
            }
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

    // Centered clip expansion for overflowing pie label text: single-line text wider
    // than its box (Office wraps instead, still open) would otherwise lose glyphs at
    // the box clip. Fitting text returns the clip unchanged; all alignments stay inside
    // the centered expansion.
    private static ChartLayoutBox ExpandPieLabelClipToText(ChartLayoutBox clip, double x, double y, double width, double height, double textWidth)
    {
        if (textWidth <= width)
        {
            return clip;
        }
        double left = x + (width - textWidth) / 2d;
        double right = Math.Max(clip.X + clip.Width, left + textWidth);
        double newLeft = Math.Min(clip.X, left);
        return new ChartLayoutBox(newLeft, clip.Y, Math.Max(1d, right - newLeft), clip.Height);
    }
    // Greedy pie label word wrap: pack words into lines within maxWidth, counting each
    // part's following separator (zero inside mid-word fragments); overlong single words
    // ride whole on their own line here and are pre-split by the caller when needed.
    // Returns part-index groups; more than one group means the label wraps.
    private static List<int[]> SplitPieLabelWordLines(IReadOnlyList<double> wordWidths, IReadOnlyList<double> separatorAfters, double maxWidth)
    {
        var lines = new List<int[]>();
        var current = new List<int>();
        double lineWidth = 0d;
        for (int i = 0; i < wordWidths.Count; i++)
        {
            double add = (current.Count == 0 ? 0d : separatorAfters[i - 1]) + wordWidths[i];
            if (current.Count > 0 && lineWidth + add > maxWidth)
            {
                lines.Add(current.ToArray());
                current = new List<int>();
                lineWidth = 0d;
                add = wordWidths[i];
            }
            current.Add(i);
            lineWidth += add;
        }
        if (current.Count > 0)
        {
            lines.Add(current.ToArray());
        }
        return lines;
    }

    // Greedy mid-word split: longest prefix runs fitting maxWidth by per-character
    // advances. Returns successive chunk char counts covering the whole text.
    private static List<int> SplitPieLabelLongWord(IReadOnlyList<double> charWidths, double maxWidth)
    {
        var chunks = new List<int>();
        int start = 0;
        while (start < charWidths.Count)
        {
            double width = 0d;
            int end = start;
            while (end < charWidths.Count && width + charWidths[end] <= maxWidth)
            {
                width += charWidths[end];
                end++;
            }
            if (end == start)
            {
                end = start + 1;
            }
            chunks.Add(end - start);
            start = end;
        }
        return chunks;
    }

    private static List<double> MeasurePieLabelChars(string text, ChartTextStyle style, ChartTextMeasurer measurer)
    {
        var widths = new List<double>(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            widths.Add(Math.Max(0d, measurer.Measure(text.Substring(i, 1), style)));
        }
        return widths;
    }
    // Valid manual pie-label layouts for leader selection: factor-mode x/y within
    // unit range (missing modes count as factor, matching ResolveDataLabelBox).
    private static bool TryGetPieManualLeaderFactorX(PptxSceneChartManualLayout layout, out double factorX)
    {
        factorX = 0d;
        if (!layout.HasLayout)
        {
            return false;
        }
        if (layout.XModeKind == PptxSceneChartManualLayoutMode.Edge || layout.YModeKind == PptxSceneChartManualLayoutMode.Edge)
        {
            return false;
        }
        double x = layout.X ?? 0d;
        double y = layout.Y ?? 0d;
        if (Math.Abs(x) > 1d || Math.Abs(y) > 1d)
        {
            return false;
        }
        factorX = x;
        return true;
    }

    // Whether the chart carries any valid manual pie label (factor-mode, in range):
    // charts without one keep legacy per-label leader emission.
    private static bool HasPieManualLeaderCandidate(ChartDataLabelOptions labelOptions, IReadOnlyList<ChartIndexedPieSlice> slices)
    {
        foreach (ChartIndexedPieSlice slice in slices)
        {
            ChartDataLabelOptions effectiveOptions = ResolveChartDataLabelOptions(labelOptions, slice.Index);
            if (!effectiveOptions.HasVisibleContent)
            {
                continue;
            }
            if (TryGetPieManualLeaderFactorX(effectiveOptions.Layout, out double _))
            {
                return true;
            }
        }
        return false;
    }
    // Per-label leader pick for charts with valid manuals: the box center must sit on
    // the wedge half of the pie (near-cardinal rims count as on-axis below 0.01), the
    // unwrapped text must fit 1.12 wrap caps, and the plot must be wide (squares draw
    // none).
    private static bool ShouldDrawPieManualLeaderLabel(double boxCenterX, double centerX, double cosTheta, double unwrappedWidth, double capWidth, double plotWidth, double plotHeight)
    {
        if (plotWidth <= plotHeight * PptxChartMetricRules.PieManualLabelLeaderMinPlotAspect)
        {
            return false;
        }
        bool east = cosTheta > 0.01d;
        bool west = cosTheta < -0.01d;
        bool sameSide = (east && boxCenterX > centerX) || (west && boxCenterX < centerX);
        return sameSide && unwrappedWidth <= capWidth * PptxChartMetricRules.PieManualLabelLeaderUnwrappedCapRatio;
    }

    // Circle-anchored box for factor-mode manual pie labels: the anchor rides a
    // circle past the rim (R plus the plot-width gap, explosion included), the box
    // hangs off it with its rim-side edge at the circle in X and its center H/2 out
    // along the radial in Y (bottom = circleY + (H/2)(sin-1) - fy plotH), then factor
    // offsets apply in plot units like Office. Exact-vertical labels shift toward the
    // center; the box clamps into the plot on all four edges (Office freezes label
    // edges at the plot rect: top-clamp kink at fy=-1, Gamma triple pinned at 120).
    private static ChartLayoutBox ResolvePieManualDataLabelBox(ChartPlotBox plotBox, PptxSceneChartManualLayout layout, double circleX, double circleY, double sinTheta, double cosTheta, double centerX, double labelWidth, double labelHeight, double singleLineHeight)
    {
        double edge = circleX - centerX;
        double factorX = layout.X ?? 0d;
        double factorY = layout.Y ?? 0d;
        double left = Math.Abs(edge) <= plotBox.Width * 0.01d
            ? circleX - labelWidth / 2d
            : edge > 0d
                ? circleX
                : circleX - labelWidth;
        left += factorX * plotBox.Width;
        double bottom;
        if (Math.Abs(cosTheta) < 0.01d)
        {
            double shift = PptxChartMetricRules.PieManualLabelCardinalVerticalShift;
            bottom = sinTheta > 0d
                ? circleY - factorY * plotBox.Height - shift
                : circleY - singleLineHeight - factorY * plotBox.Height + shift;
        }
        else
        {
            bottom = circleY + labelHeight / 2d * (sinTheta - 1d) - factorY * plotBox.Height;
        }
        double clampedLeft = ClampPieManualLabelEdge(left, labelWidth, plotBox.X, plotBox.Width);
        double clampedBottom = ClampPieManualLabelEdge(bottom, labelHeight, plotBox.Y, plotBox.Height);
        return new ChartLayoutBox(clampedLeft, clampedBottom, labelWidth, labelHeight);
    }
    // Plot clamp for one manual-label box edge: boxes wider than the plot pin to the
    // plot origin; otherwise the edge freezes at the plot rect like Office.
    private static double ClampPieManualLabelEdge(double edge, double size, double plotMin, double plotSize)
    {
        if (size >= plotSize)
        {
            return plotMin;
        }
        return Math.Min(Math.Max(edge, plotMin), plotMin + plotSize - size);
    }
    // Anchor-circle gap for manual pie labels: R plus this plot-width fraction.
    private static double ComputePieManualLabelCircleGap(double plotWidth)
    {
        return plotWidth * PptxChartMetricRules.PieManualLabelCircleGapPlotWidthFactor;
    }
    // Content-tight manual box height: line pitch times lines plus the total pad.
    private static double ComputePieManualLabelBoxHeight(double fontSize, int lineCount)
    {
        return fontSize * PptxChartMetricRules.PieDataLabelLinePitchFactor * lineCount + PptxChartMetricRules.PieManualLabelBoxHeightPad;
    }
    // Bottom-anchored manual baseline pad above the box bottom.
    private static double ComputePieManualLabelBaselinePad(double fontSize)
    {
        return fontSize * PptxChartMetricRules.PieManualLabelBaselinePadFontFactor + PptxChartMetricRules.PieManualLabelBaselinePadConstant;
    }
    // Rendered width of pie label parts joined by the separator (wrap-block skip
    // rules: blank and zero-width parts contribute nothing).
    private static double MeasurePieLabelPartsWidth(IReadOnlyList<string> parts, string separator, ChartTextStyle style, PresentationFontResolver? fontResolver)
    {
        var measurer = new ChartTextMeasurer(fontResolver);
        double separatorWidth = measurer.Measure(separator, style);
        double total = 0d;
        bool first = true;
        foreach (string part in parts)
        {
            if (string.IsNullOrWhiteSpace(part))
            {
                continue;
            }
            double partWidth = Math.Max(0d, measurer.Measure(part, style));
            if (partWidth <= 0d)
            {
                continue;
            }
            if (!first)
            {
                total += separatorWidth;
            }
            total += partWidth;
            first = false;
        }
        return total;
    }
    // Manual-leader foot height: single-line feet sit on the text baseline (2 Office
    // samples exact); wrapped feet sit 2.35pt above the first baseline at 18pt (3 Office
    // samples exact, single-size evidence so the rise scales with font size).
    private static double ComputePieLeaderFootY(double boxBottom, double boxHeight, double fontSize, int lineCount)
    {
        if (lineCount > 1)
        {
            double firstBaseline = boxBottom + boxHeight - (0.94d * fontSize + 1.5d);
            return firstBaseline + PptxChartMetricRules.PieManualLabelLeaderWrappedFootRise * fontSize / 18d;
        }
        return boxBottom + ComputePieManualLabelBaselinePad(fontSize);
    }
    private static void RenderPieDataLabelLeaderLine(PdfGraphicsBuilder graphics, ChartPolarGeometry geometry, double angleRadians, double explosion, ChartLayoutBox labelBox, ChartDataLabelOptions options, double fontSize, int lineCount, bool bottomAnchor)
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
        double footY = labelCenterY;
        double outInset;
        double inInset;
        bool verticalDrop = false;
        if (bottomAnchor)
        {
            footY = ComputePieLeaderFootY(labelBox.Y, labelBox.Height, fontSize, lineCount);
            outInset = 3d;
            inInset = 1.5d;
            verticalDrop = startX >= labelBox.X && startX <= labelBox.X + labelBox.Width && startY > labelBox.Y + labelBox.Height;
        }
        else
        {
            outInset = Math.Min(Math.Max(stroke.Width * 2d, 4d), Math.Max(4d, labelBox.Width * 0.18d));
            inInset = 0d;
        }
        double elbowX = labelIsLeft ? labelEdgeX + outInset : labelEdgeX - outInset;
        double tipX = labelIsLeft ? labelEdgeX - inInset : labelEdgeX + inInset;

        if (stroke.Alpha < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, stroke.Alpha);
        }

        SetChartStroke(graphics, stroke);
        graphics.MoveTo(startX, startY);
        if (verticalDrop)
        {
            graphics.LineTo(startX, labelBox.Y + labelBox.Height + outInset);
            graphics.LineTo(startX, labelBox.Y + labelBox.Height - inInset);
        }
        else
        {
            graphics.LineTo(elbowX, footY);
            graphics.LineTo(tipX, footY);
        }
        graphics.StrokeCurrentPath();

        if (stroke.Alpha < 1d)
        {
            graphics.RestoreState();
        }
    }
}
