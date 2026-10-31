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
    private const double ChartLineDefaultStrokeWidth = 2.25d;
    private const double ChartSeriesInheritedStrokeWidth = 3d;
    private const double ChartFilledSeriesInheritedStrokeWidth = 0.75d;
    // Explicit marker outlines without a width stroke at 1pt (COM smooth and legend
    // decks: red rims all measure 1pt); unstyled marker outlines keep the 0.75pt default
    // through ChartMarkerOutlineStroke, which only applies when no explicit line exists.
    private const double ChartMarkerInheritedStrokeWidth = 1d;

    private static async ValueTask<bool> TryRenderChartAsync(PdfGraphicsBuilder graphics, PptxDocument document, PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, ShapeBounds bounds, XDocument chartXml, PptxSceneChart sceneChart, ChartWorkbookData? workbook, List<PdfFontResource> fonts, PresentationFontResolver fontResolver,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        ChartColorPalette palette = ResolveChartColorPalette(sceneChart, chartPalette, theme, colorMap);
        bool plotVisibleOnly = ReadSceneOrXmlChartPlotVisibleOnly(sceneChart, chartXml);
        if ((await TryRenderBarChartKindAsync(
            graphics, document, theme, colorMap, palette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false)))
        {
            return true;
        }

        if ((await TryRenderLineChartKindAsync(
            graphics, document, theme, colorMap, palette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false)))
        {
            return true;
        }

        if ((await TryRenderAreaChartKindAsync(
            graphics, document, theme, colorMap, palette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false)))
        {
            return true;
        }

        if ((await TryRenderScatterChartKindAsync(
            graphics, document, theme, colorMap, palette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false)))
        {
            return true;
        }

        if ((await TryRenderBubbleChartKindAsync(
            graphics, document, theme, colorMap, palette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false)))
        {
            return true;
        }

        if ((await TryRenderRadarChartKindAsync(
            graphics, document, theme, colorMap, palette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly).ConfigureAwait(false)))
        {
            return true;
        }

        if ((await TryRenderPieChartKindAsync(
            graphics, document, theme, colorMap, palette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false)))
        {
            return true;
        }

        if ((await TryRenderDoughnutChartKindAsync(
            graphics, document, theme, colorMap, palette, bounds, chartXml, sceneChart, workbook, fonts, fontResolver, plotVisibleOnly, context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false)))
        {
            return true;
        }

        return false;
    }

    private static RgbColor ChartPalette(ChartColorPalette chartPalette, PptxTheme? theme, PptxColorMap colorMap, int index, int count = 0)
    {
        if (chartPalette.SingleAccent is { } accent)
        {
            int colorCount = Math.Max(1, count > 0 ? count : chartPalette.SeriesCount);
            return ResolveSingleAccentColor(accent, index % colorCount, colorCount);
        }

        if (chartPalette.Colors is { Count: > 0 } colors)
        {
            return colors[index % colors.Count];
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
