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
    private static IReadOnlyList<ChartLegendEntry> BuildFillLegendEntries(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?>? seriesStrokes, int paletteOffset, ChartWorkbookData? workbook)
    {
        IReadOnlyList<ChartSeriesNameRecord> names = ReadSceneOrXmlChartSeriesNameRecords(plot, chartElement, workbook);
        var entries = new List<ChartLegendEntry>(names.Count);
        for (int i = 0; i < names.Count; i++)
        {
            ChartSeriesFill fill = i < seriesFills.Count && seriesFills[i] is { } explicitFill
                ? explicitFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, i + paletteOffset), 1d, null, null);
            ChartSeriesStroke? stroke = seriesStrokes is not null && i < seriesStrokes.Count
                ? seriesStrokes[i]
                : null;
            if (stroke is { } explicitStroke &&
                PptxSceneBuilder.ParseChartPlotKind(chartElement.Name.LocalName) == PptxSceneChartPlotKind.Bar &&
                Math.Abs(explicitStroke.Width - ChartSeriesInheritedStrokeWidth) <= 0.001d)
            {
                stroke = explicitStroke with { Width = ChartFilledSeriesInheritedStrokeWidth };
            }

            entries.Add(new ChartLegendEntry(names[i].ActiveName, fill, stroke, null, names[i]));
        }

        return entries;
    }

    private static IReadOnlyList<ChartLegendEntry> BuildCategoryFillLegendEntries(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        ChartIndexedTextVector labels = ReadSceneOrXmlCategoryLabelVector(plot, chartElement, workbook, plotVisibleOnly);
        IReadOnlyList<ChartIndexedTextPoint> points = labels.DensePoints()
            .OfType<ChartIndexedTextPoint>()
            .ToArray();
        var entries = new List<ChartLegendEntry>(points.Count);
        foreach (ChartIndexedTextPoint point in points.OrderBy(point => point.Index))
        {
            if (!point.HasText || string.IsNullOrWhiteSpace(point.Text))
            {
                continue;
            }

            ChartSeriesFill fill = pointFills.TryGetValue(point.Index, out ChartSeriesFill pointFill)
                ? pointFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, point.Index), 1d, null, null);
            entries.Add(new ChartLegendEntry(point.Text, fill, null, null, null));
        }

        return entries;
    }

    private static IReadOnlyList<ChartLegendEntry> BuildStrokeLegendEntries(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle>? markerStyles, bool reverseOrder, ChartWorkbookData? workbook)
    {
        IReadOnlyList<ChartSeriesNameRecord> names = ReadSceneOrXmlChartSeriesNameRecords(plot, chartElement, workbook);
        var entries = new List<ChartLegendEntry>(names.Count);
        for (int i = 0; i < names.Count; i++)
        {
            ChartMarkerStyle? marker = markerStyles is not null && i < markerStyles.Count
                ? markerStyles[i]
                : null;
            entries.Add(new ChartLegendEntry(names[i].ActiveName, null, ChartSeriesStrokeColor(theme, colorMap, chartPalette, i, seriesStrokes, ChartLineDefaultStrokeWidth), marker, names[i]));
        }

        if (reverseOrder)
        {
            entries.Reverse();
        }

        return entries;
    }

    private static IReadOnlyList<ChartSeriesNameRecord> ReadSceneOrXmlChartSeriesNameRecords(PptxSceneChartPlot? plot, XElement chartElement, ChartWorkbookData? workbook)
    {
        if (plot is not null)
        {
            return plot.Series
                .Select((series, index) =>
                {
                    IReadOnlyList<ChartIndexedTextPoint> workbookPoints = ReadWorkbookTextPoints(workbook, series.DataSources.Name);
                    if (!string.IsNullOrWhiteSpace(series.Name))
                    {
                        return new ChartSeriesNameRecord(
                            series.Name.Trim(),
                            series.Name.Trim(),
                            ChartSeriesNameSource.Cache,
                            series.DataSources.Name,
                            workbookPoints);
                    }

                    string? workbookName = workbookPoints
                        .Where(value => value.HasText)
                        .Select(value => value.Text)
                        .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
                    return new ChartSeriesNameRecord(
                        string.IsNullOrWhiteSpace(workbookName) ? $"Series {index + 1}" : workbookName.Trim(),
                        string.Empty,
                        string.IsNullOrWhiteSpace(workbookName) ? ChartSeriesNameSource.Default : ChartSeriesNameSource.Workbook,
                        series.DataSources.Name,
                        workbookPoints);
                })
                .ToArray();
        }

        return ReadChartSeriesNameRecords(chartElement, workbook);
    }

    private static IReadOnlyList<ChartSeriesNameRecord> ReadChartSeriesNameRecords(XElement chartElement, ChartWorkbookData? workbook)
    {
        return chartElement
            .Elements(ChartNamespace + "ser")
            .Select((series, index) =>
            {
                string cacheName = ReadChartSeriesName(series)?.Trim() ?? string.Empty;
                string activeName = string.IsNullOrWhiteSpace(cacheName) ? $"Series {index + 1}" : cacheName;
                PptxSceneChartDataSource ReadChartSeriesNameDataSource()
                {
                    return PptxSceneBuilder.ReadChartDataSource(series.Element(ChartNamespace + "tx"), "strRef", "numRef");
                }

                PptxSceneChartDataSource source = ReadChartSeriesNameDataSource();
                IReadOnlyList<ChartIndexedTextPoint> workbookPoints = ReadWorkbookTextPoints(workbook, source);
                ChartSeriesNameSource activeNameSource = !string.IsNullOrWhiteSpace(cacheName)
                    ? ChartSeriesNameSource.Cache
                    : ChartSeriesNameSource.Default;
                return new ChartSeriesNameRecord(activeName, cacheName, activeNameSource, source, workbookPoints);
            })
            .ToArray();
    }

    private static string? ReadChartSeriesName(XElement series)
    {
        return PptxSceneBuilder.ReadChartText(series.Element(ChartNamespace + "tx"), trimLiteral: true);
    }

    private static ChartLegendLayout ReadChartLegendLayout(PptxTheme theme, PptxColorMap colorMap, XDocument chartXml)
    {
        PptxSceneChartLegend legend = PptxSceneBuilder.ReadChartLegend(chartXml, theme, colorMap);
        if (!legend.IsDefined || legend.IsDeleted == true)
        {
            return ChartLegendLayout.Hidden;
        }

        string position = string.IsNullOrEmpty(legend.Position) ? "r" : legend.Position;
        return new ChartLegendLayout(
            ResolveChartLegendPosition(PptxSceneBuilder.ParseChartLegendPosition(position)),
            position,
            legend.Overlay == true,
            Visible: true,
            legend.Layout,
            legend.TextBodyProperties,
            ToChartShapeStyle(legend.ShapeStyle));
    }

    private static ChartLegendLayout ReadSceneOrXmlChartLegendLayout(PptxTheme theme, PptxColorMap colorMap, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is null)
        {
            return ReadChartLegendLayout(theme, colorMap, chartXml);
        }

        if (!sceneChart.Legend.IsDefined)
        {
            return ChartLegendLayout.Hidden;
        }

        return sceneChart.Legend.IsDeleted != true
            ? new ChartLegendLayout(ResolveChartLegendPosition(sceneChart.Legend.PositionKind), sceneChart.Legend.Position, sceneChart.Legend.Overlay == true, Visible: true, sceneChart.Legend.Layout, sceneChart.Legend.TextBodyProperties, ToChartShapeStyle(sceneChart.Legend.ShapeStyle))
            : ChartLegendLayout.Hidden;
    }

    private static PptxSceneChartLegendPosition ResolveChartLegendPosition(PptxSceneChartLegendPosition position)
    {
        return position == PptxSceneChartLegendPosition.Unknown
            ? PptxSceneChartLegendPosition.Right
            : position;
    }

    private static ChartTextStyle ReadSceneOrXmlChartLegendTextStyle(PptxTheme theme, PptxColorMap colorMap, PptxSceneChart? sceneChart, XDocument chartXml)
    {
        if (sceneChart is null)
        {
            XElement? legend = chartXml.Descendants(ChartNamespace + "legend").FirstOrDefault();
            return ReadChartTextStyle(theme, colorMap, chartXml, legend, fallbackFontSize: PptxChartMetricRules.LegendFallbackFontSize);
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, sceneChart.ColorMap, fallbackFontSize: PptxChartMetricRules.LegendFallbackFontSize);
        return MergeChartTextStyle(style, ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartLegendTextStyleOverride(sceneChart)));
    }

    private static IReadOnlyList<PdfFontResource> RenderChartLegend(PdfGraphicsBuilder graphics, ChartFrameBox frame, ChartPlotBox plotBox, IReadOnlyList<ChartLegendEntry> entries, ChartLegendLayout layout, ChartTextStyle style, PresentationFontResolver? fontResolver, ChartLegendPlacement placement)
    {
        if (!layout.Visible || entries.Count == 0)
        {
            return [];
        }

        var textMeasurer = new ChartTextMeasurer(fontResolver);
        ChartLegendBox legendBox = ResolveChartLegendBox(frame, plotBox, entries, layout, style, textMeasurer, placement);

        RenderChartShapeStyle(graphics, legendBox.X, legendBox.ClipY, legendBox.Width, legendBox.ClipHeight, layout.ShapeStyle);

        var runs = new List<TextRun>(entries.Count);
        for (int i = 0; i < entries.Count; i++)
        {
            ChartLegendEntry entry = entries[i];
            double GetPackedHorizontalLegendEntryX(double markerSize, double legendX, int entryIndex)
            {
            double x = legendX;
            for (int i = 0; i < entryIndex; i++)
            {
                x += GetPackedHorizontalLegendEntryWidth(entries[i].Name, style, textMeasurer, markerSize);
            }
    
            return x;
            }

            double entryX = legendBox.Horizontal ? GetPackedHorizontalLegendEntryX(legendBox.MarkerSize, legendBox.X, i) : legendBox.X;
            double entryWidth = legendBox.Horizontal ? GetPackedHorizontalLegendEntryWidth(entries[i].Name, style, textMeasurer, legendBox.MarkerSize) : legendBox.Width;
            double y = legendBox.Horizontal ? legendBox.FirstY : legendBox.FirstY - i * legendBox.LineHeight;
            double markerBaselineFactor = legendBox.Horizontal || legendBox.SideStrokeLegend
                ? PptxChartMetricRules.LegendHorizontalMarkerBaselineFactor
                : entry.Fill is not null
                    ? PptxChartMetricRules.LegendSideFillMarkerBaselineFactor
                    : PptxChartMetricRules.LegendMarkerBaselineFactor;
            double markerY = y + legendBox.LineHeight * markerBaselineFactor;
            if (entry.Fill is { } fill)
            {
                FillChartRectangle(graphics, entryX, markerY, legendBox.MarkerSize, legendBox.MarkerSize, fill);
                if (entry.Stroke is { } fillStroke && fillStroke.Alpha > 0d)
                {
                    if (fillStroke.Alpha < 1d)
                    {
                        graphics.SaveState();
                        graphics.SetAlpha(1d, fillStroke.Alpha);
                    }

                    ChartSeriesStroke ResolveFilledLegendKeyStroke()
                    {
                        return fillStroke.Width > ChartFilledSeriesInheritedStrokeWidth
                            ? fillStroke with { Width = ChartFilledSeriesInheritedStrokeWidth }
                            : fillStroke;
                    }

                    SetChartStroke(graphics, ResolveFilledLegendKeyStroke());
                    graphics.StrokeRectangle(entryX, markerY, legendBox.MarkerSize, legendBox.MarkerSize);
                    if (fillStroke.Alpha < 1d)
                    {
                        graphics.RestoreState();
                    }
                }
            }
            else if (entry.Stroke is { } stroke)
            {
                double lineY = markerY + legendBox.MarkerSize / 2d;
                SetChartStroke(graphics, stroke);
                graphics.StrokeLine(entryX, lineY, entryX + legendBox.MarkerWidth, lineY);
                if (entry.Marker is { } marker)
                {
                    DrawChartMarker(graphics, entryX + legendBox.MarkerWidth / 2d, lineY, marker, stroke.Color, stroke.Color);
                }
            }

            runs.Add(new TextRun(
                entry.Name,
                entryX + legendBox.MarkerWidth + legendBox.TextGap,
                y,
                Math.Max(1d, entryWidth - legendBox.MarkerWidth - legendBox.TextGap),
                legendBox.LineHeight,
                legendBox.X,
                legendBox.ClipY,
                legendBox.Width,
                legendBox.ClipHeight,
                style.FontSize,
                style.CharacterSpacing,
                0d,
                style.Color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: style.Underline,
                Strike: style.Strike,
                KerningEnabled: true,
                TextAlignment.Left,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false, PreventCoalesce: false, Outline: null, StrictClip: false));
        }

        return RenderTextRuns(runs, graphics, "CL", fontResolver);
    }

    private static ChartLegendBox ResolveChartLegendBox(ChartFrameBox frame, ChartPlotBox plotBox, IReadOnlyList<ChartLegendEntry> entries, ChartLegendLayout layout, ChartTextStyle style, ChartTextMeasurer textMeasurer, ChartLegendPlacement placement)
    {
        double fontSize = style.FontSize;
        double markerSize = fontSize * PptxChartMetricRules.LegendMarkerSizeFactor;
        bool horizontal = IsHorizontalLegendPosition(layout.PositionKind);
        bool sideStrokeLegend = !horizontal && entries.All(entry => entry.Stroke is not null && entry.Fill is null);
        bool IsSameChartBox()
        {
            const double tolerance = 0.01d;
            return Math.Abs(plotBox.X - frame.X) <= tolerance &&
                Math.Abs(plotBox.Y - frame.Y) <= tolerance &&
                Math.Abs(plotBox.Width - frame.Width) <= tolerance &&
                Math.Abs(plotBox.Height - frame.Height) <= tolerance;
        }

        bool fillLegendInFullFrame = !sideStrokeLegend && IsSameChartBox();
        bool sideFillLegend = !horizontal && !fillLegendInFullFrame && !sideStrokeLegend && entries.Any(entry => entry.Fill is not null);
        bool sideFillLegendInFullFrame = !horizontal && fillLegendInFullFrame;
        double lineHeight = fontSize * (sideStrokeLegend
            || sideFillLegend
            || sideFillLegendInFullFrame
            ? PptxChartMetricRules.LegendSideStrokeLineHeightFactor
            : PptxChartMetricRules.LegendLineHeightFactor);
        double sideGap = sideStrokeLegend
            ? fontSize * PptxChartMetricRules.LegendSideStrokeGapFactor
            : PptxChartMetricRules.LegendSideGap;
        double markerWidth = sideStrokeLegend
            ? fontSize * PptxChartMetricRules.LegendSideStrokeMarkerWidthFactor
            : markerSize;
        double textGap = sideStrokeLegend
            ? fontSize * PptxChartMetricRules.LegendSideStrokeTextGapFactor
            : PptxChartMetricRules.LegendTextGap;
        double GetSideLegendContentWidth()
        {
            double contentWidth = entries.Count == 0
                ? 0d
                : entries.Max(entry => markerWidth + textGap + textMeasurer.Measure(entry.Name, style));
            return Math.Max(style.FontSize * PptxChartMetricRules.LegendSideFillMinimumWidthFactor, contentWidth);
        }

        double GetPackedHorizontalLegendWidth()
        {
            double packedWidth = 0d;
            foreach (ChartLegendEntry entry in entries)
            {
                packedWidth += GetPackedHorizontalLegendEntryWidth(entry.Name, style, textMeasurer, markerSize);
            }

            return Math.Max(1d, packedWidth);
        }

        double width = horizontal
            ? Math.Min(plotBox.Width, GetPackedHorizontalLegendWidth())
            : Math.Max(
                sideStrokeLegend || sideFillLegendInFullFrame
                    ? 0d
                    : Math.Max(PptxChartMetricRules.LegendMinimumSideWidth, plotBox.Width * PptxChartMetricRules.LegendSideWidthRatio),
                GetSideLegendContentWidth());
        double x = layout.PositionKind switch
        {
            PptxSceneChartLegendPosition.Left when sideFillLegendInFullFrame => frame.X + frame.Width * PptxChartMetricRules.LegendFullFrameSideInsetRatio,
            PptxSceneChartLegendPosition.Left => Math.Max(0d, plotBox.X - width - sideGap),
            _ when horizontal => plotBox.X + (plotBox.Width - width) / 2d,
            _ when sideFillLegendInFullFrame => frame.X + frame.Width - width,
            _ when sideFillLegend && placement == ChartLegendPlacement.BubbleTitleRightLegend => frame.X + frame.Width * PptxChartMetricRules.BubbleTitleRightLegendSwatchXRatio,
            _ when sideFillLegend => plotBox.X + plotBox.Width + sideGap + frame.Width * PptxChartMetricRules.LegendSideFillContentBoxReservedBandOffsetFactor,
            _ when !sideStrokeLegend => plotBox.X + plotBox.Width + sideGap + frame.Width * PptxChartMetricRules.LegendSideFillReservedBandOffsetFactor,
            _ => plotBox.X + plotBox.Width + sideGap
        };
        double GetLegendSideStrokeBaselineCenterOffsetFactor()
        {
            return entries.Any(entry => entry.Marker is { } marker &&
                marker.Size >= PptxChartMarkerMetricRules.StyledLineChartMarkerSize - PptxChartMetricRules.AxisValueEpsilon)
                ? PptxChartMetricRules.LegendSideStrokeStyledMarkerBaselineCenterOffsetFactor
                : PptxChartMetricRules.LegendSideStrokeBaselineCenterOffsetFactor;
        }

        double firstY = layout.PositionKind switch
        {
            PptxSceneChartLegendPosition.Bottom when fillLegendInFullFrame => frame.Y + lineHeight * PptxChartMetricRules.LegendFullFrameBottomBaselineFactor,
            PptxSceneChartLegendPosition.Top when fillLegendInFullFrame => frame.Y + frame.Height - lineHeight * PptxChartMetricRules.LegendFullFrameTopBaselineFactor,
            PptxSceneChartLegendPosition.Bottom => Math.Max(0d, plotBox.Y - lineHeight * PptxChartMetricRules.LegendBottomOffsetFactor),
            PptxSceneChartLegendPosition.Top => plotBox.Y + plotBox.Height + lineHeight * PptxChartMetricRules.LegendTopOffsetFactor,
            _ when sideStrokeLegend => plotBox.Y + plotBox.Height / 2d -
                fontSize * GetLegendSideStrokeBaselineCenterOffsetFactor() +
                (entries.Count - 1) * lineHeight / 2d,
            _ when sideFillLegend && placement == ChartLegendPlacement.BubbleTitleRightLegend => frame.Y + frame.Height * PptxChartMetricRules.BubbleTitleRightLegendSwatchYRatio,
            _ when sideFillLegend => frame.Y + frame.Height / 2d -
                fontSize * PptxChartMetricRules.LegendSideFillBaselineCenterOffsetFactor +
                (entries.Count - 1) * lineHeight / 2d,
            _ when !sideStrokeLegend && !horizontal => frame.Y + frame.Height / 2d +
                (entries.Count - 1) * lineHeight / 2d,
            _ when !horizontal => plotBox.Y + plotBox.Height / 2d +
                fontSize * PptxChartMetricRules.LegendMarkerBaselineFactor +
                (entries.Count - 1) * lineHeight / 2d,
            _ => plotBox.Y + plotBox.Height - lineHeight
        };
        double clipHeight = horizontal ? lineHeight * PptxChartMetricRules.LegendHorizontalClipHeightFactor : Math.Max(lineHeight, entries.Count * lineHeight);
        double clipY = horizontal ? firstY : Math.Max(0d, firstY - (entries.Count - 1) * lineHeight);
        if (TryBuildManualLayoutBox(layout.Layout, frame, new ChartLayoutBox(x, clipY, width, clipHeight), out ChartLayoutBox manualBox, true, false))
        {
            x = manualBox.X;
            width = manualBox.Width;
            clipY = manualBox.Y;
            clipHeight = manualBox.Height;
            firstY = horizontal
                ? manualBox.Y
                : manualBox.Y + manualBox.Height - lineHeight;
        }

        return new ChartLegendBox(x, clipY, width, clipHeight, firstY, lineHeight, markerSize, markerWidth, textGap, horizontal, sideStrokeLegend);
    }

    private static double GetPackedHorizontalLegendEntryWidth(string name, ChartTextStyle style, ChartTextMeasurer textMeasurer, double markerSize)
    {
        return markerSize + PptxChartMetricRules.LegendTextGap + textMeasurer.Measure(name, style) + PptxChartMetricRules.LegendHorizontalEntryPadding;
    }

    private static bool IsOoxmlTrue(string? value)
    {
        return OoxBoolean.IsTrue(value);
    }

    private static bool? ReadOoxmlBooleanAttribute(XElement? element, string name)
    {
        return element?.Attribute(name) is { } attribute
            ? IsOoxmlTrue(attribute.Value)
            : null;
    }
}
