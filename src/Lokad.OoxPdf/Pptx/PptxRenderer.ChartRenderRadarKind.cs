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
        bool plotVisibleOnly,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        XElement? radarChart = ReadSceneOrXmlFirstChartPlotElement(sceneChart, chartXml, PptxSceneChartPlotKind.Radar);
        if (radarChart is not null)
        {
            PptxSceneChartPlot? radarPlot = ReadSceneChartPlot(sceneChart, PptxSceneChartPlotKind.Radar, 0);
            IReadOnlyList<ChartIndexedNumberVector> radarSeriesVectors = ReadSharedChartSeriesVectors(radarPlot, radarChart, workbook, plotVisibleOnly);
            IReadOnlyList<ChartRadarSeries> radarSeries = BuildRadarSeries(radarSeriesVectors);
            if (radarSeries.Count != 0)
            {
                IReadOnlyList<ChartSeriesFill?> seriesFills = ReadSceneOrXmlSeriesFills(radarPlot, radarChart, theme, colorMap);
                IReadOnlyList<ChartSeriesStroke?> seriesStrokes = ReadSceneOrXmlSeriesStrokes(radarPlot, radarChart, theme, colorMap, null);
                ChartAxisSource valueAxis = ReadSceneOrXmlChartValueAxesForPlot(sceneChart, radarPlot, chartXml, radarChart).FirstOrDefault();
                ChartAxisSource categoryAxis = ReadSceneOrXmlChartCategoryAxisForPlot(sceneChart, radarPlot, chartXml, radarChart);
                ChartValueExtents valueExtents = ReadSceneOrXmlChartValueAxisExtents(valueAxis.SceneAxis, valueAxis.XmlAxis, GetRadarChartValueExtents(radarSeries), false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
                ChartAxisUnits axisUnits = ReadSceneOrXmlChartValueAxisUnits(valueAxis.SceneAxis, valueAxis.XmlAxis);
                ChartFrameBox frame = GetChartFrameBox(document, bounds);
                ChartPlotBox defaultPlotBox = new(frame.X, frame.Y, frame.Width, frame.Height);
                bool radarManualPlot = TryReadSceneOrXmlManualPlotLayout(sceneChart, chartXml, frame, defaultPlotBox, out ChartPlotLayout radarManualLayout);
                ChartPlotBox plotBox = radarManualPlot ? radarManualLayout.PlotBox : defaultPlotBox;
                ChartRadarPlotOptions radarOptions = ReadSceneOrXmlChartRadarOptions(radarPlot, radarChart);
                bool radarHasTitle = !string.IsNullOrWhiteSpace(ReadSceneOrXmlChartTitleText(sceneChart, chartXml));
                ChartRadarLayout radarLayout = ResolveRadarLayout(frame, plotBox, radarOptions.RadarStyle, radarSeries, radarHasTitle, radarManualPlot);
                RenderChartAreaStyle(graphics, document, bounds, chartXml, sceneChart, theme, colorMap);
                RenderRadarChart(graphics, theme, colorMap, chartPalette, radarLayout, radarSeries, seriesFills, seriesStrokes, valueExtents, axisUnits);
                if (IsSceneOrXmlChartAxisLabelVisible(categoryAxis.SceneAxis, categoryAxis.XmlAxis))
                {
                    RenderRadarCategoryLabels(theme, graphics, radarLayout, chartXml, sceneChart, categoryAxis.SceneAxis, categoryAxis.XmlAxis, ReadSharedCategoryLabels(radarPlot, radarChart, workbook, plotVisibleOnly), fontResolver, chartFonts: fonts);
                }

                if (IsSceneOrXmlChartAxisLabelVisible(valueAxis.SceneAxis, valueAxis.XmlAxis))
                {
                    RenderRadarValueAxisLabels(theme, graphics, radarLayout, chartXml, sceneChart, valueAxis.XmlAxis, valueAxis.SceneAxis, valueExtents, axisUnits, fontResolver, chartFonts: fonts);
                }

                return true;
            }
        }
        return false;
    }
}
