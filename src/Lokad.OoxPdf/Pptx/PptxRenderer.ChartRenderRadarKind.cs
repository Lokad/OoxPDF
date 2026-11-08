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
    private static async ValueTask<bool> TryRenderRadarChartKindAsync(
        PdfGraphicsBuilder graphics,
        PptxDocument document,
        PptxTheme theme,
        PptxColorMap colorMap,
        ChartColorPalette chartPalette,
        ShapeBounds bounds,
        XDocument chartXml,
        PptxSceneChart sceneChart,
        ChartWorkbookData? workbook,
        List<PdfFontResource> fonts,
        PresentationFontResolver fontResolver,
        bool plotVisibleOnly, PptxRenderContext? context = null)
    {
        PptxSceneChartPlot? radarPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Radar, 0);
        if (radarPlot is not null)
        {
            XElement radarChart = radarPlot.Source;
            IReadOnlyList<ChartIndexedNumberVector> radarSeriesVectors = ReadSharedChartSeriesVectors(radarPlot, workbook, plotVisibleOnly);
            IReadOnlyList<ChartRadarSeries> radarSeries = BuildRadarSeries(radarSeriesVectors);
            if (radarSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneSeriesFills(radarPlot);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneSeriesStrokes(radarPlot, null);
                ChartAxisSource valueAxis = ReadChartValueAxesForPlot(sceneChart, radarPlot).FirstOrDefault();
                ChartAxisSource categoryAxis = ReadChartCategoryAxisForPlot(sceneChart, radarPlot);
                ChartValueExtents valueExtents = ReadSceneOrXmlChartValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetRadarChartValueExtents(radarSeries), false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartAxisUnits axisUnits = ReadSceneChartValueAxisUnits(valueAxis.SceneAxis);
                ChartFrameBox frame = GetChartFrameBox(document, bounds);
                ChartPlotBox defaultPlotBox = new(frame.X, frame.Y, frame.Width, frame.Height);
                bool radarManualPlot = TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout radarManualLayout);
                ChartPlotBox plotBox = radarManualPlot ? radarManualLayout.PlotBox : defaultPlotBox;
                PptxSceneChartRadarStyle radarStyle = ReadSceneChartRadarStyle(radarPlot);
                bool radarHasTitle = !string.IsNullOrWhiteSpace(sceneChart.Title.Text);
                ChartRadarLayout radarLayout = ResolveRadarLayout(frame, plotBox, radarStyle, radarSeries, radarHasTitle, radarManualPlot);
                RenderChartAreaStyle(graphics, document, bounds, sceneChart);
                int? radarStyleId = ReadChartStyleId(chartXml);
                RenderRadarChart(graphics, theme, colorMap, chartPalette, radarLayout, radarSeries, seriesFills, seriesStrokes, valueExtents, axisUnits, radarStyleId, context: context);
                if (IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    await RenderRadarCategoryLabelsAsync(theme, graphics, radarLayout, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSharedCategoryLabels(radarPlot, workbook, plotVisibleOnly), fontResolver, chartFonts: fonts).ConfigureAwait(false);
                }

                if (IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    await RenderRadarValueAxisLabelsAsync(theme, graphics, radarLayout, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, axisUnits, fontResolver, chartFonts: fonts, context: context).ConfigureAwait(false);
                }

                return true;
            }
        }
        return false;
    }
}
