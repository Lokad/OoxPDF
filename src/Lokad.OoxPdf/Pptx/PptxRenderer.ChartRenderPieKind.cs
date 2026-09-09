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
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Pie, plotBox, polarPoints.PointExplosions, legend);
                RenderPieChart(graphics, theme, colorMap, chartPalette, polarLayout, pieSlices, polarPoints.PointFills, polarPoints.PointStrokes, polarPoints.PointExplosions, polarPoints.FirstSliceAngle);
                RenderPieDataLabels(theme, colorMap, graphics, chartPalette, polarLayout, pieSlices, polarPoints.PointFills, polarPoints.PointExplosions, 0d, pieSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver, fonts, context, sceneChart?.Relationships, linkAnnotations, reportedHyperlinkIds);
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
                ChartPolarLayout polarLayout = ResolvePieOrDoughnutLayout(ChartPolarKind.Doughnut, plotBox, polarPoints.PointExplosions, legend);
                RenderDoughnutChart(graphics, theme, colorMap, chartPalette, polarLayout, doughnutSlices, polarPoints.PointFills, polarPoints.PointStrokes, polarPoints.PointExplosions, doughnutOptions.HoleSize, polarPoints.FirstSliceAngle);
                RenderPieDataLabels(theme, colorMap, graphics, chartPalette, polarLayout, doughnutSlices, polarPoints.PointFills, polarPoints.PointExplosions, doughnutOptions.HoleSize, doughnutSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver, fonts, context, sceneChart?.Relationships, linkAnnotations, reportedHyperlinkIds);
                RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, doughnutPlot, doughnutChart, polarPoints.PointFills, workbook, plotVisibleOnly), legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default, chartFonts: fonts);
                return true;
            }
        }
        return false;
    }
}
