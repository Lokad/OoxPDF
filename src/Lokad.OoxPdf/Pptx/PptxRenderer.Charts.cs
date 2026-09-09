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

    private static bool TryRenderChart(PdfGraphicsBuilder graphics, PptxDocument document, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ShapeBounds bounds, XDocument chartXml, PptxSceneChart? sceneChart, ChartWorkbookData? workbook, List<PdfFontResource> fonts, PresentationFontResolver fontResolver,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        bool plotVisibleOnly = ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, chartXml);
        if (TryRenderBarChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds))
        {
            return true;
        }

        if (TryRenderLineChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds))
        {
            return true;
        }

        if (TryRenderAreaChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds))
        {
            return true;
        }

        if (TryRenderScatterChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds))
        {
            return true;
        }

        if (TryRenderBubbleChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds))
        {
            return true;
        }

        if (TryRenderRadarChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds))
        {
            return true;
        }

        if (TryRenderPieChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds))
        {
            return true;
        }

        if (TryRenderDoughnutChartKind(
            graphics, document, theme, colorMap, chartPalette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds))
        {
            return true;
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
