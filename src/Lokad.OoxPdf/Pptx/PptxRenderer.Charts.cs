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
    private const string WorkbookRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string WorksheetRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    private const string SharedStringsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    private const string SpreadsheetStylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    private const string SpreadsheetTableRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
    private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    private const string SharedStringsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    private const string SpreadsheetStylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    private const double ChartLineDefaultStrokeWidth = 2.25d;
    private const double ChartSeriesInheritedStrokeWidth = 3d;
    private const double ChartFilledSeriesInheritedStrokeWidth = 0.75d;
    private const double ChartMarkerInheritedStrokeWidth = 0.75d;

    private static bool TryRenderChart(PdfGraphicsBuilder graphics, PptxDocument document, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, ChartWorkbookData? workbook, List<PdfFontResource> fonts, PresentationFontResolver fontResolver)
    {
        bool plotVisibleOnly = ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, chartXml);
        if (TryRenderBarChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly))
        {
            return true;
        }

        if (TryRenderLineChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly))
        {
            return true;
        }

        if (TryRenderAreaChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly))
        {
            return true;
        }

        if (TryRenderScatterChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly))
        {
            return true;
        }

        if (TryRenderBubbleChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly))
        {
            return true;
        }

        if (TryRenderRadarChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly))
        {
            return true;
        }

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
                fonts.AddRange(RenderPieDataLabels(theme, colorMap, graphics, chartPalette, polarLayout, pieSlices, polarPoints.PointFills, polarPoints.PointExplosions, 0d, pieSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, piePlot, pieChart, polarPoints.PointFills, workbook, plotVisibleOnly), legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default));
                return true;
            }
        }

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
                fonts.AddRange(RenderPieDataLabels(theme, colorMap, graphics, chartPalette, polarLayout, doughnutSlices, polarPoints.PointFills, polarPoints.PointExplosions, doughnutOptions.HoleSize, doughnutSeriesVectors[0].FormatCode, labelOptions, categoryLabels, seriesNames, fontResolver));
                fonts.AddRange(RenderChartLegend(graphics, frame, plotBox, BuildCategoryFillLegendEntries(theme, colorMap, chartPalette, doughnutPlot, doughnutChart, polarPoints.PointFills, workbook, plotVisibleOnly), legend, ReadSceneOrXmlChartLegendTextStyle(theme, colorMap, sceneChart, chartXml), fontResolver, ChartLegendPlacement.Default));
                return true;
            }
        }

        return false;
    }

    private static RgbColor ChartPalette(IReadOnlyList<RgbColor>? chartPalette, PptxTheme? theme, int index)
    {
        return ChartPalette(chartPalette, theme, PptxColorMap.Default, index);
    }

    private static RgbColor ChartPalette(IReadOnlyList<RgbColor>? chartPalette, PptxTheme? theme, PptxColorMap colorMap, int index)
    {
        if (chartPalette is { Count: > 0 })
        {
            return chartPalette[index % chartPalette.Count];
        }

        if (theme is not null && theme.TryResolveColor("accent" + (index % 6 + 1).ToString(CultureInfo.InvariantCulture), colorMap, out RgbColor themeColor))
        {
            return themeColor;
        }

        RgbColor[] palette =
        [
            new RgbColor(68, 114, 196),
            new RgbColor(237, 125, 49),
            new RgbColor(165, 165, 165),
            new RgbColor(255, 192, 0),
            new RgbColor(91, 155, 213),
            new RgbColor(112, 173, 71)
        ];
        return palette[index % palette.Length];
    }

    private static RgbColor ChartPalette(int index)
    {
        return ChartPalette(null, null, index);
    }

    private static RgbColor ChartPalette(PptxTheme? theme, int index)
    {
        return ChartPalette(null, theme, index);
    }

    private static void StrokeChartPointRectangleInPlotClip(PdfGraphicsBuilder graphics, ChartPlotBox plotBox, int seriesIndex, int categoryIndex, IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes, double x, double y, double width, double height, ChartSeriesStroke? fallbackStroke)
    {
        bool hasExplicitStroke = seriesIndex < pointStrokes.Count && pointStrokes[seriesIndex].ContainsKey(categoryIndex);
        if (!hasExplicitStroke && fallbackStroke is null)
        {
            return;
        }

        RenderInChartPlotAreaClip(graphics, plotBox, () => StrokeChartPointRectangle(graphics, seriesIndex, categoryIndex, pointStrokes, x, y, width, height, fallbackStroke));
    }


}
