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
    private static bool TryRenderRadarChartKind(
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
        bool plotVisibleOnly)
    {
        XElement? radarChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Radar);
        if (radarChart is not null)
        {
            PptxSceneChartPlot? radarPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Radar, 0);
            IReadOnlyList<ChartIndexedNumberVector> radarSeriesVectors = ReadSceneOrXmlChartSeriesVectors(radarPlot, radarChart, workbook, plotVisibleOnly);
            IReadOnlyList<ChartRadarSeries> radarSeries = BuildRadarSeries(radarSeriesVectors);
            if (radarSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(radarPlot, radarChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(radarPlot, radarChart, theme, colorMap, null);
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, radarPlot, chartXml, radarChart).FirstOrDefault();
                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, radarPlot, chartXml, radarChart);
                ChartValueExtents valueExtents = ReadSceneOrXmlChartValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetRadarChartValueExtents(radarSeries), false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartAxisUnits axisUnits = ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxis.XmlAxis);
                ChartPlotBox plotBox = GetPolarChartPlotBox(document, bounds, chartXml, sceneChart);
                ChartRadarPlotOptions radarOptions = ReadSceneOrXmlChartRadarOptions(radarPlot, radarChart);
                ChartRadarLayout radarLayout = ResolveRadarLayout(plotBox, radarOptions.RadarStyle, radarSeries);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                RenderRadarChart(graphics, theme, colorMap, chartPalette, radarLayout, radarSeries, seriesFills, seriesStrokes, valueExtents, axisUnits);
                if (IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    fonts.AddRange(RenderRadarCategoryLabels(theme, graphics, radarLayout, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSceneOrXmlCategoryLabelVector(radarPlot, radarChart, workbook, plotVisibleOnly), fontResolver));
                }

                if (IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    fonts.AddRange(RenderRadarValueAxisLabels(theme, graphics, radarLayout, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, axisUnits, fontResolver));
                }

                return true;
            }
        }
        return false;
    }
}
