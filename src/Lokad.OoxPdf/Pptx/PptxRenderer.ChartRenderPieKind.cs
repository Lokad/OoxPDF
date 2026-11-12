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
    private static async ValueTask<bool> TryRenderPieChartKindAsync(
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
        bool plotVisibleOnly,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        PptxSceneChartPlot? piePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Pie, 0);
        if (piePlot is not null)
        {
            XElement pieChart = piePlot.Source;
            IReadOnlyList<ChartIndexedNumberVector> pieSeriesVectors = ReadSharedChartSeriesVectors(piePlot, workbook, plotVisibleOnly);
            IReadOnlyList<ChartIndexedPieSlice> pieSlices = pieSeriesVectors.Count == 0 ? [] : BuildChartIndexedPieSlices(pieSeriesVectors[0]);
            if (pieSlices.Count != 0)
            {
                ChartIndexedTextVector categoryLabels = ReadSceneCategoryLabelVector(piePlot, workbook, plotVisibleOnly);
                IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSharedChartSeriesNames(piePlot, workbook);
                ChartDataLabelOptions labelOptions = ResolveChartDataLabelOptionsForSeries(
                    ReadSceneDataLabelOptions(sceneChart, piePlot),
                    ReadSceneSeriesDataLabelOptions(sceneChart, piePlot),
                    seriesIndex: 0);
                ChartPolarPointOptions polarPoints = ReadSceneChartPolarPointOptions(piePlot, workbook);
                ChartLayout sourceLayout = GetPolarChartLayout(document, bounds, sceneChart);
                ChartFrameBox frame = sourceLayout.Frame;
                ChartLegendLayout legend = sourceLayout.Legend;
                ChartPlotBox plotBox = sourceLayout.PlotBox;
                IReadOnlyList<ChartLegendEntry> legendEntries = BuildCategoryFillLegendEntries(theme, colorMap, chartPalette,
                    piePlot, pieChart, polarPoints.PointFills, workbook, plotVisibleOnly, valuePointCount: pieSlices.Count);
                ChartTextStyle legendStyle = ReadSceneChartLegendTextStyle(theme, sceneChart);
                var prepared = await PreparePolarChartLayoutAsync(sourceLayout, theme, sceneChart,
                    legendEntries, legendStyle, fontResolver, context).ConfigureAwait(false);
                RenderChartAreaStyle(graphics, document, bounds, sceneChart);
                bool hasVisiblePieDataLabels = labelOptions.HasVisibleText || pieSlices.Any(slice => ResolveChartDataLabelOptions(labelOptions, slice.Index).HasVisibleText);
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Pie, plotBox, polarPoints.PointExplosions, legend, hasVisibleDataLabels: hasVisiblePieDataLabels, hasTitle: !string.IsNullOrWhiteSpace(sceneChart.Title.Text));
                if (sourceLayout.ManualPlotLayoutApplied && sceneChart.PlotAreaLayout.X is not null &&
                    IsManualLayoutFactorMode(sceneChart.PlotAreaLayout.XModeKind))
                {
                    // Office's factor offset follows the automatic center even
                    // when an authored width resizes the pie. Edge placement
                    // still belongs to the authored rectangle. Reuse the same
                    // default geometry and resolved, bounded source offset;
                    // don't infer a new side-legend anchor.
                    ChartPolarLayout automatic = ResolvePieOrDoughnutLayout(ChartPolarKind.Pie,
                        new ChartPlotBox(frame.X, frame.Y, frame.Width, frame.Height), polarPoints.PointExplosions,
                        legend, hasVisibleDataLabels: hasVisiblePieDataLabels, hasTitle: !string.IsNullOrWhiteSpace(sceneChart.Title.Text));
                    polarLayout = polarLayout with { Geometry = polarLayout.Geometry with
                        { CenterX = automatic.Geometry.CenterX + plotBox.X - frame.X } };
                }
                if (UsesAutomaticPolarFrame(sourceLayout, sceneChart))
                    polarLayout = FitPieLayoutToFrame(polarLayout, prepared.Layout.PlotBox, hasVisiblePieDataLabels);
                RenderPieChart(graphics, theme, colorMap, chartPalette, polarLayout, pieSlices, polarPoints.PointFills, polarPoints.PointStrokes, polarPoints.PointExplosions, polarPoints.FirstSliceAngle);
                await RenderPieDataLabelsAsync(theme, colorMap, graphics, chartPalette, polarLayout, pieSlices, polarPoints.PointFills, polarPoints.PointExplosions, 0d, polarPoints.FirstSliceAngle, pieSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver, fonts, context, sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                await RenderChartLegendAsync(graphics, frame, plotBox, legendEntries, legend, legendStyle,
                    fontResolver, ChartLegendPlacement.Default, chartFonts: fonts, preparedBox: prepared.LegendBox).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }

    private static async ValueTask<bool> TryRenderDoughnutChartKindAsync(
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
        bool plotVisibleOnly,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        PptxSceneChartPlot? doughnutPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Doughnut, 0);
        if (doughnutPlot is not null)
        {
            XElement doughnutChart = doughnutPlot.Source;
            IReadOnlyList<ChartIndexedNumberVector> doughnutSeriesVectors = ReadSharedChartSeriesVectors(doughnutPlot, workbook, plotVisibleOnly);
            IReadOnlyList<ChartIndexedPieSlice> doughnutSlices = doughnutSeriesVectors.Count == 0 ? [] : BuildChartIndexedPieSlices(doughnutSeriesVectors[0]);
            if (doughnutSlices.Count != 0)
            {
                ChartIndexedTextVector categoryLabels = ReadSceneCategoryLabelVector(doughnutPlot, workbook, plotVisibleOnly);
                IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSharedChartSeriesNames(doughnutPlot, workbook);
                ChartDataLabelOptions labelOptions = ResolveChartDataLabelOptionsForSeries(
                    ReadSceneDataLabelOptions(sceneChart, doughnutPlot),
                    ReadSceneSeriesDataLabelOptions(sceneChart, doughnutPlot),
                    seriesIndex: 0);
                ChartDoughnutPlotOptions doughnutOptions = ReadSceneChartDoughnutOptions(doughnutPlot, workbook);
                ChartLayout sourceLayout = GetPolarChartLayout(document, bounds, sceneChart);
                ChartFrameBox frame = sourceLayout.Frame;
                ChartLegendLayout legend = sourceLayout.Legend;
                ChartPlotBox plotBox = sourceLayout.PlotBox;
                RenderChartAreaStyle(graphics, document, bounds, sceneChart);
                ChartPolarPointOptions polarPoints = doughnutOptions.PolarPoints;
                IReadOnlyList<ChartLegendEntry> doughnutLegendEntries = BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, doughnutPlot, doughnutChart, polarPoints.PointFills, workbook, plotVisibleOnly, valuePointCount: doughnutSlices.Count);
                ChartTextStyle doughnutLegendStyle = ReadSceneChartLegendTextStyle(theme, sceneChart);
                var prepared = await PreparePolarChartLayoutAsync(sourceLayout, theme, sceneChart,
                    doughnutLegendEntries, doughnutLegendStyle, fontResolver, context).ConfigureAwait(false);
                ChartPlotBox doughnutGeometryPlotBox = plotBox;
                // Exploded charts keep the full-frame plot only while the legend fits the slack:
                // narrow names show no Office shrink, but mid/wide Office rings translate rigidly
                // left (1:2 ring:legend rule), so wide legends force a content-driven reserve that
                // translates the ring by half the reserve with identical radius. The legend box
                // below re-anchors to the same tail edge.
                bool isDoughnutFillLegend = legend.Visible && !legend.Overlay && !doughnutLegendEntries.All(entry => entry.Stroke is not null && entry.Fill is null);
                bool isDoughnutRightFillLegend = isDoughnutFillLegend && legend.PositionKind == PptxSceneChartLegendPosition.Right;
                bool isDoughnutLeftFillLegend = isDoughnutFillLegend && legend.PositionKind == PptxSceneChartLegendPosition.Left;
                bool isExplodedDoughnut = polarPoints.PointExplosions.Count != 0;
                double explodedDoughnutLegendReserve = 0d;
                bool hasTitle = !string.IsNullOrWhiteSpace(sceneChart.Title.Text);
                if (isDoughnutRightFillLegend && !isExplodedDoughnut)
                {
                    var legendMeasurer = new ChartTextMeasurer(fontResolver);
                    await legendMeasurer.PrepareAsync(doughnutLegendEntries.Select(entry => entry.Name), doughnutLegendStyle).ConfigureAwait(false);
                    double doughnutLegendReserve = ComputeDoughnutRightLegendReserve(doughnutLegendEntries, doughnutLegendStyle, legendMeasurer);
                    doughnutGeometryPlotBox = new ChartPlotBox(plotBox.X, plotBox.Y, Math.Max(1d, plotBox.Width - doughnutLegendReserve), plotBox.Height);
                }
                else if (isDoughnutRightFillLegend)
                {
                    // Gap-driven reserve needs the unshrunk ring edge: lay out once without reserve,
                    // measure the legend box the renderer will draw, then translate by the shortfall.
                    var legendMeasurer = new ChartTextMeasurer(fontResolver);
                    await legendMeasurer.PrepareAsync(doughnutLegendEntries.Select(entry => entry.Name), doughnutLegendStyle).ConfigureAwait(false);
                    double explodedBoxContent = MeasureDoughnutLegendContentWidth(doughnutLegendEntries, doughnutLegendStyle, legendMeasurer, PptxChartMetricRules.AreaRightLegendTextGap);
                    double explodedBoxLeft = frame.X + frame.Width - PptxChartMetricRules.DoughnutRightLegendTail - explodedBoxContent;
                    double maxExplosion = polarPoints.PointExplosions.Values.Max();
                    ChartPolarLayout unshrunkLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Doughnut, doughnutGeometryPlotBox, polarPoints.PointExplosions, legend, hasVisibleDataLabels: false, hasLegendReserve: false, explodedRightLegendReserve: 0d, leftLegendBoxRight: 0d, hasTitle: hasTitle);
                    double ringRightUnshrunk = unshrunkLayout.Geometry.CenterX + unshrunkLayout.Geometry.Radius * (1d + maxExplosion);
                    explodedDoughnutLegendReserve = ComputeExplodedDoughnutRightLegendReserve(explodedBoxLeft, ringRightUnshrunk);
                }
                bool hasDoughnutLegendReserve = doughnutGeometryPlotBox.Width < plotBox.Width;
                // Left-legend geometry centers the ring past the legend box edge, measured with the
                // same content the legend box will use; manual layouts pass zero and keep legacy.
                bool doughnutLegendSharesFrame = Math.Abs(plotBox.X - frame.X) <= 0.01d && Math.Abs(plotBox.Y - frame.Y) <= 0.01d && Math.Abs(plotBox.Width - frame.Width) <= 0.01d && Math.Abs(plotBox.Height - frame.Height) <= 0.01d;
                double leftLegendBoxRight = 0d;
                if (isDoughnutLeftFillLegend && doughnutLegendSharesFrame)
                {
                    var legendMeasurer = new ChartTextMeasurer(fontResolver);
                    await legendMeasurer.PrepareAsync(doughnutLegendEntries.Select(entry => entry.Name), doughnutLegendStyle).ConfigureAwait(false);
                    double leftContentWidth = MeasureDoughnutLegendContentWidth(doughnutLegendEntries, doughnutLegendStyle, legendMeasurer, PptxChartMetricRules.AreaRightLegendTextGap);
                    leftLegendBoxRight = frame.X + PptxChartMetricRules.DoughnutLeftLegendHeadInset + leftContentWidth;
                }
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Doughnut, doughnutGeometryPlotBox, polarPoints.PointExplosions, legend, hasVisibleDataLabels: false, hasLegendReserve: hasDoughnutLegendReserve, explodedRightLegendReserve: explodedDoughnutLegendReserve, leftLegendBoxRight: leftLegendBoxRight, hasTitle: hasTitle);
                RenderDoughnutChart(graphics, theme, colorMap, chartPalette, polarLayout, doughnutSlices, polarPoints.PointFills, polarPoints.PointStrokes, polarPoints.PointExplosions, doughnutOptions.HoleSize, polarPoints.FirstSliceAngle);
                await RenderPieDataLabelsAsync(theme, colorMap, graphics, chartPalette, polarLayout, doughnutSlices, polarPoints.PointFills, polarPoints.PointExplosions, doughnutOptions.HoleSize, polarPoints.FirstSliceAngle, doughnutSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver, fonts, context, sceneChart.Relationships, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                await RenderChartLegendAsync(graphics, frame, plotBox, doughnutLegendEntries, legend, doughnutLegendStyle, fontResolver, ChartLegendPlacement.Default, chartFonts: fonts, doughnutRightLegend: isDoughnutRightFillLegend, doughnutLeftLegend: isDoughnutLeftFillLegend, doughnutRingCenterY: polarLayout.Geometry.CenterY, preparedBox: prepared.LegendBox).ConfigureAwait(false);
                return true;
            }
        }
        return false;
    }
}
