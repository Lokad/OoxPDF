using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static async ValueTask<bool> TryRenderRadarChartKindAsync(
        PdfGraphicsBuilder graphics,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        ChartColorPalette chartPalette,
        ShapeBounds bounds,
        PptxSceneChart sceneChart,
        ChartWorkbookData? workbook,
        List<PdfFontResource> fonts,
        PresentationFontResolver fontResolver,
        bool plotVisibleOnly, PptxRenderContext? context = null)
    {
        PptxSceneChartPlot? radarPlot = ReadFirstSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Radar);
        if (radarPlot is not null)
        {
            IReadOnlyList<ChartIndexedNumberVector> radarSeriesVectors = ReadSharedChartSeriesVectors(radarPlot, workbook, plotVisibleOnly);
            double?[][] radarSeries = BuildRadarSeries(radarSeriesVectors);
            if (radarSeries.Any(values => values.Length != 0))
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneSeriesFills(radarPlot);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneSeriesStrokes(radarPlot, null);
                ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, radarPlot).FirstOrDefault();
                ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, radarPlot);
                ChartValueExtents valueExtents = ReadSceneChartValueAxisExtents(valueAxis.SceneAxis, GetRadarChartValueExtents(radarSeries), false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartAxisUnits axisUnits = ReadSceneChartValueAxisUnits(valueAxis.SceneAxis);
                ChartFrameBox frame = GetChartFrameBox(document, bounds);
                ChartPlotBox defaultPlotBox = new(frame.X, frame.Y, frame.Width, frame.Height);
                bool radarManualPlot = TryBuildManualPlotLayout(sceneChart.PlotAreaLayout, frame, defaultPlotBox, out ChartPlotLayout radarManualLayout);
                ChartPlotBox plotBox = radarManualPlot ? radarManualLayout.PlotBox : defaultPlotBox;
                // An authored inner plot bounds the web. Axis labels use the enclosing chart frame.
                ChartPlotBox captionClip = radarManualLayout.ManualLayoutTargetKind == PptxSceneChartManualLayoutTarget.Inner
                    ? defaultPlotBox : plotBox;
                PptxSceneChartRadarStyle radarStyle = ReadSceneChartRadarStyle(radarPlot);
                bool radarHasTitle = !string.IsNullOrWhiteSpace(sceneChart.Title.Text);
                ChartRadarLayout radarLayout = ResolveRadarLayout(frame, plotBox, captionClip, radarStyle, radarSeries, radarHasTitle, radarManualPlot);
                IReadOnlyList<double> ticks = GetChartAxisTickValues(valueExtents, axisUnits.MajorUnit,
                    includeEndpoints: true, PptxChartMetricRules.AxisNiceTickTargetCount, context);
                IReadOnlyList<ChartIndexedTextPoint?> categorySource = IsSceneChartAxisLabelVisible(categoryAxis.SceneAxis)
                    ? ReadSharedCategoryLabels(radarPlot, workbook, plotVisibleOnly) : [];
                IReadOnlyList<double> labeledTicks = IsSceneChartAxisLabelVisible(valueAxis.SceneAxis) ? ticks : [];
                List<TextRun> categoryLabels = await PrepareRadarCategoryLabelsAsync(theme, radarLayout, sceneChart,
                    categoryAxis.SceneAxis, categorySource, fontResolver).ConfigureAwait(false);
                List<TextRun> valueLabels = await PrepareRadarValueAxisLabelsAsync(theme, radarLayout, sceneChart,
                    valueAxis.SceneAxis, valueExtents, labeledTicks, fontResolver).ConfigureAwait(false);

                ChartLegendLayout legend = ReadSceneChartLegendLayout(sceneChart);
                ChartLegendBox? legendBox = null;
                ChartTextStyle legendStyle = default;
                IReadOnlyList<ChartLegendEntry> legendEntries = [];
                double plotShift = 0d;
                if (legend.Visible)
                {
                    if (legend.PositionKind is PptxSceneChartLegendPosition.Left or PptxSceneChartLegendPosition.Right &&
                        sceneChart.Legend.DeletedEntryIndices is not { Count: > 0 })
                    {
                        bool legendOnLeft = legend.PositionKind == PptxSceneChartLegendPosition.Left;
                        var horizontalLimits = (Left: Math.Max(frame.X, 0d),
                            Right: Math.Min(frame.X + frame.Width, document.SlideWidthPoints));
                        legendEntries = radarLayout.IsFilled
                            ? BuildFillLegendEntries(theme, colorMap, chartPalette, radarPlot,
                                seriesFills, seriesStrokes, paletteOffset: 0, workbook)
                            : ReadSharedChartSeriesNames(radarPlot, workbook).Select((name, index) => new ChartLegendEntry(
                                name.ActiveName, null, ResolveRadarSeriesStroke(theme, colorMap, chartPalette, seriesStrokes, index),
                                null, name, LineHidden: false)).ToArray();
                        legendStyle = ReadSceneChartLegendTextStyle(theme, sceneChart);
                        legendBox = await PrepareRadarSideLegendAsync(frame, plotBox, legend, legendEntries,
                            legendStyle, fontResolver, horizontalLimits, document.SlideHeightPoints, legendOnLeft).ConfigureAwait(false);
                        if (legendBox is { } box && !legend.Overlay)
                        {
                            double strokeWidth = Math.Max(PptxChartMetricRules.RadarSeriesOutlineWidth,
                                seriesStrokes.Max(stroke => stroke?.Width ?? 0d));
                            if (!TryPlaceRadarPlotBesideLegend(radarLayout, radarManualPlot, box, legendStyle,
                                    strokeWidth, categoryLabels.Concat(valueLabels), horizontalLimits, legendOnLeft, out plotShift))
                            {
                                if (!radarManualPlot && TryFitRadarWebBesideLegend(radarLayout, box, strokeWidth,
                                        categoryLabels, valueLabels, horizontalLimits, legendOnLeft, out var fitted))
                                {
                                    radarLayout = fitted;
                                    categoryLabels = await PrepareRadarCategoryLabelsAsync(theme, radarLayout, sceneChart,
                                        categoryAxis.SceneAxis, categorySource, fontResolver).ConfigureAwait(false);
                                    valueLabels = await PrepareRadarValueAxisLabelsAsync(theme, radarLayout, sceneChart,
                                        valueAxis.SceneAxis, valueExtents, labeledTicks, fontResolver).ConfigureAwait(false);
                                }
                                else
                                {
                                    legendBox = null;
                                }
                            }
                        }
                    }
                    if (legendBox is null)
                    {
                        EmitChartDiagnostic(context?.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning,
                            "The radar legend placement is unsupported or does not fit within its chart frame.",
                            sceneChart.TargetPartName, context?.SlideNumber ?? 0, "Legend ignored");
                    }
                }

                RenderChartAreaStyle(graphics, document, bounds, sceneChart);
                if (plotShift != 0d)
                {
                    graphics.SaveState();
                    graphics.Transform(1d, 0d, 0d, 1d, plotShift, 0d);
                }
                int? radarStyleId = sceneChart.SeriesStyleId;
                RenderRadarChart(graphics, theme, colorMap, chartPalette, radarLayout, radarSeries, seriesFills, seriesStrokes, valueExtents, ticks, radarStyleId, context);
                if (categoryLabels.Count != 0)
                {
                    await RenderChartTextRunsAsync(categoryLabels, graphics, fonts, "RCA", fontResolver, diagnosticSink: null).ConfigureAwait(false);
                }

                if (valueLabels.Count != 0)
                {
                    await RenderChartTextRunsAsync(valueLabels, graphics, fonts, "RVA", fontResolver, diagnosticSink: null).ConfigureAwait(false);
                }
                if (plotShift != 0d) graphics.RestoreState();

                if (legendBox is { } prepared)
                {
                    await RenderChartLegendAsync(graphics, frame, plotBox, legendEntries, legend, legendStyle,
                        fontResolver, ChartLegendPlacement.Default, fonts, context?.DiagnosticSink, preparedBox: prepared).ConfigureAwait(false);
                }

                return true;
            }
        }
        return false;
    }

    private static async ValueTask<ChartLegendBox?> PrepareRadarSideLegendAsync(
        ChartFrameBox frame, ChartPlotBox plot, ChartLegendLayout legend, IReadOnlyList<ChartLegendEntry> entries,
        ChartTextStyle style, PresentationFontResolver resolver, (double Left, double Right) horizontalLimits,
        double slideHeight, bool legendOnLeft)
    {
        if (entries.Count == 0) return null;
        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        await measure.PrepareAsync(entries.Select(entry => entry.Name), style).ConfigureAwait(false);
        var box = ResolveChartLegendBox(frame, plot, entries, legend, style, measure, ChartLegendPlacement.Default);
        double requiredWidth = box.MarkerWidth + box.TextGap + entries.Max(entry => measure.Measure(entry.Name, style));
        if (!legend.Layout.HasGeometryValues)
        {
            double inset = PptxChartMetricRules.ChartContentFrameInset;
            double width = Math.Max(box.Width, requiredWidth + inset);
            double x = legendOnLeft ? horizontalLimits.Left + inset : horizontalLimits.Right - inset - width;
            box = box with { X = x, Width = width };
        }

        double bottom = Math.Max(frame.Y, 0d);
        double top = Math.Min(frame.Y + frame.Height, slideHeight);
        double lastBaseline = ResolveLegendEntryBaseline(box, entries.Count - 1);
        return box.Width >= requiredWidth && box.X >= horizontalLimits.Left && box.X + box.Width <= horizontalLimits.Right &&
            box.ClipY >= bottom && box.ClipY + box.ClipHeight <= top &&
            lastBaseline >= box.ClipY && lastBaseline - style.FontSize >= bottom && box.FirstY + style.FontSize <= top
            ? box : null;
    }

    private static bool TryPlaceRadarPlotBesideLegend(ChartRadarLayout radar, bool manualPlot, ChartLegendBox legend,
        ChartTextStyle legendStyle, double strokeWidth, IEnumerable<TextRun> labels,
        (double Left, double Right) horizontalLimits, bool legendOnLeft, out double shift)
    {
        // Preserve the web radius and the axis label positions within the web.
        // A non-overlay legend can take a free band or translate an automatic web.
        // In this coordinate system the legend is always on the high side.
        double direction = legendOnLeft ? -1d : 1d;
        double center = direction * radar.Geometry.CenterX;
        double farEdge = center - radar.Geometry.Radius - strokeWidth;
        double nearEdge = center + radar.Geometry.Radius + strokeWidth;
        double bottom = legend.ClipY - legendStyle.FontSize;
        double top = legend.ClipY + legend.ClipHeight + legendStyle.FontSize;
        foreach (TextRun label in labels)
        {
            double labelLeft = Math.Max(label.X, label.ClipX);
            double labelRight = Math.Min(label.X + label.Width, label.ClipX + label.ClipWidth);
            farEdge = Math.Min(farEdge, legendOnLeft ? -labelRight : labelLeft);
            if (label.Y + label.Height > bottom && label.Y - label.FontSize < top)
                nearEdge = Math.Max(nearEdge, legendOnLeft ? -labelLeft : labelRight);
        }

        double legendEdge = legendOnLeft ? -(legend.X + legend.Width) : legend.X;
        double farLimit = legendOnLeft ? -horizontalLimits.Right : horizontalLimits.Left;
        double displacement = Math.Min(0d, legendEdge - PptxChartMetricRules.ChartContentFrameInset - nearEdge);
        if (displacement == 0d || !manualPlot && farEdge + displacement >= farLimit)
        {
            shift = direction * displacement;
            return true;
        }
        shift = 0d;
        return false;
    }

    private static bool TryFitRadarWebBesideLegend(ChartRadarLayout radar, ChartLegendBox legend, double strokeWidth,
        IReadOnlyList<TextRun> categoryLabels, IReadOnlyList<TextRun> valueLabels,
        (double Left, double Right) horizontalLimits, bool legendOnLeft, out ChartRadarLayout fitted)
    {
        fitted = radar;
        double inset = PptxChartMetricRules.ChartContentFrameInset;
        double left = (legendOnLeft ? legend.X + legend.Width : horizontalLimits.Left) + inset;
        double right = (legendOnLeft ? horizontalLimits.Right : legend.X) - inset;
        double gapFactor = radar.LabelRules.CategoryHorizontalGapSideFactor;
        // The horizontal category anchor grows with the web radius and its
        // proportional gap. The full prepared label width bounds either alignment.
        double categoryBand = categoryLabels.Count == 0 ? 0d :
            categoryLabels.Max(label => label.Width) + 2d * gapFactor * PptxChartMetricRules.RadarWebRadiusPenHalf;
        double side = Math.Max(strokeWidth, categoryBand);
        double valueBand = valueLabels.Count == 0 ? 0d :
            radar.Geometry.CenterX - valueLabels.Min(label => label.X);
        double leftBand = Math.Max(side, valueBand);
        double radius = Math.Min(radar.Geometry.Radius, (right - left - leftBand - side) / (2d * (1d + 2d * gapFactor)));
        if (!(radius >= 1d)) return false;
        double center = (left + leftBand + right - side) / 2d;
        fitted = radar with { Geometry = radar.Geometry with { CenterX = center, Radius = radius } };
        return true;
    }
}
