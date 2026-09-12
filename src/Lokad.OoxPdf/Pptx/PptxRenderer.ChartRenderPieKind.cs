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
    private static bool TryRenderPieChartKind(
        PdfGraphicsBuilder graphics,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        IReadOnlyList<RgbColor>? chartPalette,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        ChartWorkbookData? workbook,
        List<PdfFontResource> fonts,
        PresentationFontResolver fontResolver,
        bool plotVisibleOnly,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        XElement? pieChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Pie);
        if (pieChart is not null)
        {
            PptxSceneChartPlot? piePlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Pie, 0);
            IReadOnlyList<ChartIndexedNumberVector> pieSeriesVectors = ReadSceneOrXmlChartSeriesVectors(piePlot, pieChart, workbook, plotVisibleOnly);
            IReadOnlyList<ChartIndexedPieSlice> pieSlices = pieSeriesVectors.Count == 0 ? [] : BuildChartIndexedPieSlices(pieSeriesVectors[0]);
            if (pieSlices.Count != 0)
            {
                ChartIndexedTextVector categoryLabels = ReadSceneOrXmlCategoryLabelVector(piePlot, pieChart, workbook, plotVisibleOnly);
                IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSceneOrXmlChartSeriesNameRecords(piePlot, pieChart, workbook);
                ChartDataLabelOptions labelOptions = ResolveChartDataLabelOptionsForSeries(
                    ReadSceneOrXmlDataLabelOptions(sceneChart, piePlot, pieChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, piePlot, pieChart, theme, colorMap),
                    seriesIndex: 0);
                ChartPolarPointOptions polarPoints = ReadSceneOrXmlChartPolarPointOptions(piePlot, pieChart, theme, colorMap, workbook);
                ChartFrameBox frame = GetChartFrameBox(document, bounds);
                ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
                ChartPlotBox plotBox = GetPolarChartPlotBox(document, bounds, chartXml, sceneChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                bool hasVisiblePieDataLabels = labelOptions.HasVisibleContent || pieSlices.Any(slice => ResolveChartDataLabelOptions(labelOptions, slice.Index).HasVisibleContent);
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Pie, plotBox, polarPoints.PointExplosions, legend, hasVisibleDataLabels: hasVisiblePieDataLabels);
                RenderPieChart(graphics, theme, colorMap, chartPalette, polarLayout, pieSlices, polarPoints.PointFills, polarPoints.PointStrokes, polarPoints.PointExplosions, polarPoints.FirstSliceAngle);
                RenderPieDataLabels(theme, colorMap, graphics, chartPalette, polarLayout, pieSlices, polarPoints.PointFills, polarPoints.PointExplosions, 0d, polarPoints.FirstSliceAngle, pieSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver, fonts, context, sceneChart?.Relationships, linkAnnotations, reportedHyperlinkIds);
                RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, piePlot, pieChart, polarPoints.PointFills, workbook, plotVisibleOnly), legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default, chartFonts: fonts);
                return true;
            }
        }
        return false;
    }

    private static bool TryRenderDoughnutChartKind(
        PdfGraphicsBuilder graphics,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        IReadOnlyList<RgbColor>? chartPalette,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        ChartWorkbookData? workbook,
        List<PdfFontResource> fonts,
        PresentationFontResolver fontResolver,
        bool plotVisibleOnly,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        XElement? doughnutChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Doughnut);
        if (doughnutChart is not null)
        {
            PptxSceneChartPlot? doughnutPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Doughnut, 0);
            IReadOnlyList<ChartIndexedNumberVector> doughnutSeriesVectors = ReadSceneOrXmlChartSeriesVectors(doughnutPlot, doughnutChart, workbook, plotVisibleOnly);
            IReadOnlyList<ChartIndexedPieSlice> doughnutSlices = doughnutSeriesVectors.Count == 0 ? [] : BuildChartIndexedPieSlices(doughnutSeriesVectors[0]);
            if (doughnutSlices.Count != 0)
            {
                ChartIndexedTextVector categoryLabels = ReadSceneOrXmlCategoryLabelVector(doughnutPlot, doughnutChart, workbook, plotVisibleOnly);
                IReadOnlyList<ChartSeriesNameRecord> seriesNames = ReadSceneOrXmlChartSeriesNameRecords(doughnutPlot, doughnutChart, workbook);
                ChartDataLabelOptions labelOptions = ResolveChartDataLabelOptionsForSeries(
                    ReadSceneOrXmlDataLabelOptions(sceneChart, doughnutPlot, doughnutChart, theme, colorMap),
                    ReadSceneOrXmlSeriesDataLabelOptions(sceneChart, doughnutPlot, doughnutChart, theme, colorMap),
                    seriesIndex: 0);
                ChartDoughnutPlotOptions doughnutOptions = ReadSceneOrXmlChartDoughnutOptions(doughnutPlot, doughnutChart, theme, colorMap, workbook);
                ChartFrameBox frame = GetChartFrameBox(document, bounds);
                ChartLegendLayout legend = ReadSceneOrXmlChartLegendLayout(theme, colorMap, sceneChart, chartXml);
                ChartPlotBox plotBox = GetPolarChartPlotBox(document, bounds, chartXml, sceneChart);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                ChartPolarPointOptions polarPoints = doughnutOptions.PolarPoints;
                IReadOnlyList<ChartLegendEntry> doughnutLegendEntries = BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, doughnutPlot, doughnutChart, polarPoints.PointFills, workbook, plotVisibleOnly);
                ChartTextStyle doughnutLegendStyle = ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml);
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
                if (isDoughnutRightFillLegend && !isExplodedDoughnut)
                {
                    double doughnutLegendReserve = ComputeDoughnutRightLegendReserve(doughnutLegendEntries, doughnutLegendStyle, new ChartTextMeasurer(fontResolver));
                    doughnutGeometryPlotBox = new ChartPlotBox(plotBox.X, plotBox.Y, Math.Max(1d, plotBox.Width - doughnutLegendReserve), plotBox.Height);
                }
                else if (isDoughnutRightFillLegend)
                {
                    explodedDoughnutLegendReserve = ComputeExplodedDoughnutRightLegendReserve(MeasureDoughnutLegendContentWidth(doughnutLegendEntries, doughnutLegendStyle, new ChartTextMeasurer(fontResolver)));
                }
                bool hasDoughnutLegendReserve = doughnutGeometryPlotBox.Width < plotBox.Width;
                // Left-legend geometry centers the ring past the legend box edge, measured with the
                // same content the legend box will use; manual layouts pass zero and keep legacy.
                bool doughnutLegendSharesFrame = Math.Abs(plotBox.X - frame.X) <= 0.01d && Math.Abs(plotBox.Y - frame.Y) <= 0.01d && Math.Abs(plotBox.Width - frame.Width) <= 0.01d && Math.Abs(plotBox.Height - frame.Height) <= 0.01d;
                double leftLegendBoxRight = 0d;
                if (isDoughnutLeftFillLegend && doughnutLegendSharesFrame)
                {
                    double leftMarkerWidth = doughnutLegendStyle.FontSize * PptxChartMetricRules.LegendMarkerSizeFactor;
                    double leftContentWidth = 0d;
                    var leftMeasurer = new ChartTextMeasurer(fontResolver);
                    foreach (ChartLegendEntry entry in doughnutLegendEntries)
                    {
                        leftContentWidth = Math.Max(leftContentWidth, leftMarkerWidth + PptxChartMetricRules.AreaRightLegendTextGap + leftMeasurer.Measure(entry.Name, doughnutLegendStyle));
                    }
                    leftLegendBoxRight = frame.X + PptxChartMetricRules.DoughnutLeftLegendHeadInset + leftContentWidth;
                }
                bool doughnutHasTitle = !string.IsNullOrWhiteSpace(ReadSceneOrXmlChartTitleText(sceneChart, chartXml));
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Doughnut, doughnutGeometryPlotBox, polarPoints.PointExplosions, legend, hasVisibleDataLabels: false, hasLegendReserve: hasDoughnutLegendReserve, explodedRightLegendReserve: explodedDoughnutLegendReserve, leftLegendBoxRight: leftLegendBoxRight, doughnutHasTitle: doughnutHasTitle);
                RenderDoughnutChart(graphics, theme, colorMap, chartPalette, polarLayout, doughnutSlices, polarPoints.PointFills, polarPoints.PointStrokes, polarPoints.PointExplosions, doughnutOptions.HoleSize, polarPoints.FirstSliceAngle);
                RenderPieDataLabels(theme, colorMap, graphics, chartPalette, polarLayout, doughnutSlices, polarPoints.PointFills, polarPoints.PointExplosions, doughnutOptions.HoleSize, polarPoints.FirstSliceAngle, doughnutSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver, fonts, context, sceneChart?.Relationships, linkAnnotations, reportedHyperlinkIds);
                RenderChartLegend(graphics, frame, plotBox, doughnutLegendEntries, legend, doughnutLegendStyle, fontResolver, ChartLegendPlacement.Default, chartFonts: fonts, doughnutRightLegend: isDoughnutRightFillLegend, doughnutLeftLegend: isDoughnutLeftFillLegend, doughnutRingCenterY: polarLayout.Geometry.CenterY);
                return true;
            }
        }
        return false;
    }
}
