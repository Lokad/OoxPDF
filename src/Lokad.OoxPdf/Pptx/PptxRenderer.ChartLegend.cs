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
    private static IReadOnlyList<ChartLegendEntry> BuildFillLegendEntries(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesFill?> seriesFills, IReadOnlyList<ChartSeriesStroke?>? seriesStrokes, int paletteOffset, ChartWorkbookData? workbook, bool reverseOrder = false)
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

            entries.Add(new ChartLegendEntry(names[i].ActiveName, fill, stroke, null, names[i], LineHidden: false));
        }

        // Stacked area legends list top-of-stack first (area-stacked Office reference
        // shows Support/Service/Product against file order Product/Service/Support, while
        // standard areas, stacked columns and stacked lines keep their own established
        // orders; stacked-right-column would discriminate a position-driven rival).
        if (reverseOrder)
        {
            entries.Reverse();
        }

        return entries;
    }

    private static IReadOnlyList<ChartLegendEntry> BuildCategoryFillLegendEntries(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyDictionary<int, ChartSeriesFill> pointFills, ChartWorkbookData? workbook, bool plotVisibleOnly, int valuePointCount = 0)
    {
        ChartIndexedTextVector labels = ReadSceneOrXmlCategoryLabelVector(plot, chartElement, workbook, plotVisibleOnly);
        IReadOnlyList<ChartIndexedTextPoint> points = labels.DensePoints()
            .OfType<ChartIndexedTextPoint>()
            .ToArray();
        var entries = new List<ChartLegendEntry>(points.Count);
        var usedIndexes = new HashSet<int>();
        foreach (ChartIndexedTextPoint point in points.OrderBy(point => point.Index))
        {
            if (!point.HasText || string.IsNullOrWhiteSpace(point.Text))
            {
                continue;
            }

            usedIndexes.Add(point.Index);
            ChartSeriesFill fill = pointFills.TryGetValue(point.Index, out ChartSeriesFill pointFill)
                ? pointFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, point.Index), 1d, null, null);
            entries.Add(new ChartLegendEntry(point.Text, fill, null, null, null, LineHidden: false));
        }

        // Value points without category text still draw swatch-only legend rows
        // (doughnut overlay probe: four values over three categories emits a textless
        // fourth swatch; pies share the rule, unobserved).
        for (int valueIndex = 0; valueIndex < valuePointCount; valueIndex++)
        {
            if (!usedIndexes.Add(valueIndex))
            {
                continue;
            }

            ChartSeriesFill valueFill = pointFills.TryGetValue(valueIndex, out ChartSeriesFill valuePointFill)
                ? valuePointFill
                : new ChartSeriesFill(ChartPalette(chartPalette, theme, colorMap, valueIndex), 1d, null, null);
            entries.Add(new ChartLegendEntry(string.Empty, valueFill, null, null, null, LineHidden: false));
        }

        return entries;
    }

    private static IReadOnlyList<ChartLegendEntry> BuildStrokeLegendEntries(PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? chartPalette, PptxSceneChartPlot? plot, XElement chartElement, IReadOnlyList<ChartSeriesStroke?> seriesStrokes, IReadOnlyList<ChartMarkerStyle>? markerStyles, bool reverseOrder, ChartWorkbookData? workbook, IReadOnlyList<bool>? seriesLineHidden = null)
    {
        IReadOnlyList<ChartSeriesNameRecord> names = ReadSceneOrXmlChartSeriesNameRecords(plot, chartElement, workbook);
        var entries = new List<ChartLegendEntry>(names.Count);
        for (int i = 0; i < names.Count; i++)
        {
            ChartMarkerStyle? marker = markerStyles is not null && i < markerStyles.Count
                ? markerStyles[i]
                : null;
            bool lineHidden = seriesLineHidden is not null && i < seriesLineHidden.Count && seriesLineHidden[i];
            ChartSeriesStroke keyStroke = ChartSeriesStrokeColor(theme, colorMap, chartPalette, i, seriesStrokes, ChartLineDefaultStrokeWidth);
            // Unstyled key lines default to round caps/joins; explicitly styled keys keep
            // DrawingML attr defaults (same rule as series lines).
            if (i >= seriesStrokes.Count || seriesStrokes[i] is null)
            {
                keyStroke = keyStroke with { Cap = keyStroke.Cap ?? 1, Join = keyStroke.Join ?? 1 };
            }
            entries.Add(new ChartLegendEntry(names[i].ActiveName, null, keyStroke, marker, names[i], LineHidden: lineHidden));
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
        return style.Merge(ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartLegendTextStyleOverride(sceneChart)));
    }

    private static void RenderChartLegend(PdfGraphicsBuilder graphics, ChartFrameBox frame, ChartPlotBox plotBox, IReadOnlyList<ChartLegendEntry> entries, ChartLegendLayout layout, ChartTextStyle style, PresentationFontResolver? fontResolver, ChartLegendPlacement placement, List<PdfFontResource> chartFonts, Action<OoxPdfDiagnostic>? diagnosticSink = null, double legendLeadExtra = 0d, bool doughnutRightLegend = false, bool doughnutLeftLegend = false, double doughnutRingCenterY = 0d, bool scatterLegend = false)
    {
        if (!layout.Visible || entries.Count == 0)
        {
            return;
        }

        // Legend boxes measure unkerned advances like the emission (see reserve note).
        var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
        ChartLegendBox legendBox = ResolveChartLegendBox(frame, plotBox, entries, layout, style, textMeasurer, placement, legendLeadExtra, doughnutRightLegend, doughnutLeftLegend, doughnutRingCenterY, scatterLegend);

        RenderChartShapeStyle(graphics, legendBox.X, legendBox.ClipY, legendBox.Width, legendBox.ClipHeight, layout.ShapeStyle);

        bool emitBottomLegend = layout.PositionKind == PptxSceneChartLegendPosition.Bottom;
        double emitTextGap = ComputeHorizontalLegendTextGap(style.FontSize, emitBottomLegend);
        double emitEntryPadding = ComputeHorizontalLegendInterEntryGap(style.FontSize, ComputeHorizontalLegendAdvanceRange(entries, style, textMeasurer), emitBottomLegend);
        var runs = new List<TextRun>(entries.Count);
        for (int i = 0; i < entries.Count; i++)
        {
            ChartLegendEntry entry = entries[i];
            double GetPackedHorizontalLegendEntryX(double markerSize, double legendX, int entryIndex)
            {
            double x = legendX;
            for (int i = 0; i < entryIndex; i++)
            {
                x += GetPackedHorizontalLegendEntryWidth(entries[i].Name, style, textMeasurer, markerSize, emitTextGap, emitEntryPadding);
            }
    
            return x;
            }

            double entryX = legendBox.Horizontal ? GetPackedHorizontalLegendEntryX(legendBox.MarkerSize, legendBox.X, i) : legendBox.X;
            double entryWidth = legendBox.Horizontal ? GetPackedHorizontalLegendEntryWidth(entries[i].Name, style, textMeasurer, legendBox.MarkerSize, emitTextGap, emitEntryPadding) : legendBox.Width;
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
                // An explicit series noFill suppresses the key line sample: Office shows
                // marker-only keys there. The stroke stays on the entry so box layout and
                // the plot reserve do not move. Without a marker, keep the sample rather
                // than emitting an empty key.
                if (!entry.LineHidden || entry.Marker is null)
                {
                    SetChartStroke(graphics, stroke);
                    graphics.StrokeLine(entryX, lineY, entryX + legendBox.MarkerWidth, lineY);
                }
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
                // Office applies no pair kerning to legend entries (Markets 37.6 vs 4, Monthly
                // Trend 84 vs 5, Base 6.8 vs 0 against our GPOS pairs; titles keep kerning, like
                // the cartesian axes and radar labels which already pass kerning off).
                KerningEnabled: false,
                TextAlignment.Left,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false, PreventCoalesce: false, Outline: null, StrictClip: false));
        }

        RenderChartTextRuns(runs, graphics, chartFonts, "CL", fontResolver, diagnosticSink);
    }

    // Fill-swatch side keys share the 4.66pt Office entry text gap: area right legends
    // measure 4.65-4.67 on five references, doughnut 4.58-4.66 on eleven renders, and the
    // bubble title-right ladder port 4.64. Stroke keys keep their font-scaled gap; every
    // other side legend keeps the legacy 3pt gap for lack of evidence.
    private static double ResolveSideLegendTextGap(
        ChartLegendPlacement placement,
        bool useDoughnutRightAnchor,
        bool useDoughnutLeftAnchor,
        bool sideStrokeLegend,
        double fontSize)
    {
        return sideStrokeLegend
            ? fontSize * PptxChartMetricRules.LegendSideStrokeTextGapFactor
            : placement == ChartLegendPlacement.AreaRightLegend
                || placement == ChartLegendPlacement.BubbleTitleRightLegend
                || useDoughnutRightAnchor
                || useDoughnutLeftAnchor
                ? PptxChartMetricRules.AreaRightLegendTextGap
                : PptxChartMetricRules.LegendTextGap;
    }

    // The styled-line center offset is for explicitly shaped line markers only: sparse
    // scatter defaults (9.9) trip the size threshold while Office centers them exactly
    // like plain keys (scatter, line and line-markers blocks all center at 282.77 on
    // three ladder ports), so scatter legends keep the plain offset.
    private static double ResolveSideStrokeLegendCenterOffsetFactor(IReadOnlyList<double> markerSizes, bool scatterLegend)
    {
        return !scatterLegend && markerSizes.Any(size =>
            size >= PptxChartMarkerMetricRules.StyledLineChartMarkerSize - PptxChartMetricRules.AxisValueEpsilon)
            ? PptxChartMetricRules.LegendSideStrokeStyledMarkerBaselineCenterOffsetFactor
            : PptxChartMetricRules.LegendSideStrokeBaselineCenterOffsetFactor;
    }

    // Right stroke-key legend X from the axis-based plot right edge plus the clip
    // allowance (Office anchors to the clip edge; see the constant evidence note).
    private static double ResolveStrokeLegendRightX(double plotRight, double sideGap, double legendLeadExtra)
    {
        return plotRight + sideGap + legendLeadExtra + PptxChartMetricRules.LegendStrokeRightClipAllowance;
    }

    private static ChartLegendBox ResolveChartLegendBox(ChartFrameBox frame, ChartPlotBox plotBox, IReadOnlyList<ChartLegendEntry> entries, ChartLegendLayout layout, ChartTextStyle style, ChartTextMeasurer textMeasurer, ChartLegendPlacement placement, double legendLeadExtra = 0d, bool doughnutRightLegend = false, bool doughnutLeftLegend = false, double doughnutRingCenterY = 0d, bool scatterLegend = false)
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
        // Doughnut right legends (exploded or not) share the measured vertical anchor; manual-layout
        // charts keep the legacy placement for lack of evidence.
        bool useDoughnutRightAnchor = doughnutRightLegend && sideFillLegendInFullFrame;
        // Doughnut left legends share the pattern with their own measured shift; the left x anchor
        // stays legacy.
        bool useDoughnutLeftAnchor = doughnutLeftLegend && sideFillLegendInFullFrame;
        // Shared vertical anchor: right and left blocks sit below the plot middle by their own
        // measured shifts; every other side legend keeps the centered anchor.
        double doughnutLegendVerticalShift = useDoughnutRightAnchor
            ? PptxChartMetricRules.DoughnutRightLegendVerticalShift
            : useDoughnutLeftAnchor ? PptxChartMetricRules.DoughnutLeftLegendVerticalShift : 0d;
        // The x re-anchor below now covers every doughnut right legend (exploded or not): the
        // Office block right-anchors to the 10pt tail edge content-tight on eight renders.
        double lineHeight = ComputeFullFrameFillLegendLineHeight(fontSize, sideFillLegendInFullFrame, sideStrokeLegend, sideFillLegend);
        double sideGap = sideStrokeLegend
            ? fontSize * PptxChartMetricRules.LegendSideStrokeGapFactor
            : PptxChartMetricRules.LegendSideGap;
        double markerWidth = sideStrokeLegend
            ? fontSize * PptxChartMetricRules.LegendSideStrokeMarkerWidthFactor
            : markerSize;
        // Fill-swatch side keys share the resolved 4.66pt Office text gap (area, doughnut
        // and bubble branches per the helper evidence note).
        double textGap = ResolveSideLegendTextGap(placement, useDoughnutRightAnchor, useDoughnutLeftAnchor, sideStrokeLegend, fontSize);
        double GetSideLegendContentWidth()
        {
            double contentWidth = entries.Count == 0
                ? 0d
                : entries.Max(entry => markerWidth + textGap + textMeasurer.Measure(entry.Name, style));
            // Exploded doughnut boxes run content-tight like the Office block (narrow Office box
            // is 25pt with ~10pt names, well under the minimum factor); no other path changes.
            return useDoughnutRightAnchor ? contentWidth : Math.Max(style.FontSize * PptxChartMetricRules.LegendSideFillMinimumWidthFactor, contentWidth);
        }

        bool bottomLegend = layout.PositionKind == PptxSceneChartLegendPosition.Bottom;
        double horizontalTextGap = ComputeHorizontalLegendTextGap(fontSize, bottomLegend);
        double horizontalEntryPadding = ComputeHorizontalLegendInterEntryGap(fontSize, ComputeHorizontalLegendAdvanceRange(entries, style, textMeasurer), bottomLegend);
        if (horizontal && bottomLegend)
        {
            textGap = horizontalTextGap;
        }

        double GetPackedHorizontalLegendWidth()
        {
            double packedWidth = 0d;
            foreach (ChartLegendEntry entry in entries)
            {
                packedWidth += GetPackedHorizontalLegendEntryWidth(entry.Name, style, textMeasurer, markerSize, horizontalTextGap, horizontalEntryPadding);
            }

            // Office rows carry no trailing padding after the last entry (botleg block
            // matches the untrailed sum within 0.02); the padding still separates entries.
            if (entries.Count > 0)
            {
                packedWidth -= horizontalEntryPadding;
            }

            return Math.Max(1d, packedWidth);
        }

        // Area right legends right-anchor to the frame tail, so their box must be
        // content-tight like the Office block; the plot-relative ratio floor would
        // over-widen it by ~33pt.
        bool areaRightLegend = sideFillLegend && placement == ChartLegendPlacement.AreaRightLegend;
        double width = horizontal
            ? Math.Min(plotBox.Width, GetPackedHorizontalLegendWidth())
            : Math.Max(
                sideStrokeLegend || sideFillLegendInFullFrame || areaRightLegend
                    ? 0d
                    : Math.Max(PptxChartMetricRules.LegendMinimumSideWidth, plotBox.Width * PptxChartMetricRules.LegendSideWidthRatio),
                GetSideLegendContentWidth());
        double x = layout.PositionKind switch
        {
            // Doughnut left blocks grow rightward from a fixed head edge (four Office renders).
            _ when useDoughnutLeftAnchor => frame.X + PptxChartMetricRules.DoughnutLeftLegendHeadInset,
            PptxSceneChartLegendPosition.Left when sideFillLegendInFullFrame => frame.X + frame.Width * PptxChartMetricRules.LegendFullFrameSideInsetRatio,
            PptxSceneChartLegendPosition.Left => Math.Max(0d, plotBox.X - width - sideGap),
            _ when horizontal => plotBox.X + (plotBox.Width - width) / 2d,
            // Doughnut right blocks right-anchor to the tail edge like the Office block (right
            // edge 10pt inside the frame on eight Office renders, same tail as the geometry).
            _ when useDoughnutRightAnchor => frame.X + frame.Width - PptxChartMetricRules.DoughnutRightLegendTail - width,
            _ when sideFillLegendInFullFrame => frame.X + frame.Width - width,
            _ when sideFillLegend && placement == ChartLegendPlacement.BubbleTitleRightLegend => frame.X + frame.Width * PptxChartMetricRules.BubbleTitleRightLegendSwatchXRatio,
            _ when sideFillLegend && placement == ChartLegendPlacement.AreaRightLegend && !layout.Overlay && layout.PositionKind == PptxSceneChartLegendPosition.Right => frame.X + frame.Width - PptxChartMetricRules.AreaRightLegendTail - width,
            _ when sideFillLegend => plotBox.X + plotBox.Width + sideGap + frame.Width * PptxChartMetricRules.LegendSideFillContentBoxReservedBandOffsetFactor,
            _ when !sideStrokeLegend => plotBox.X + plotBox.Width + sideGap + frame.Width * PptxChartMetricRules.LegendSideFillReservedBandOffsetFactor,
            _ => ResolveStrokeLegendRightX(plotBox.X + plotBox.Width, sideGap, legendLeadExtra)
        };

        List<double> legendMarkerSizes = new(entries.Count);
        foreach (ChartLegendEntry legendEntry in entries)
        {
            if (legendEntry.Marker is { } legendMarker)
            {
                legendMarkerSizes.Add(legendMarker.Size);
            }
        }

                double firstY = layout.PositionKind switch
        {
            PptxSceneChartLegendPosition.Bottom when fillLegendInFullFrame => frame.Y + lineHeight * PptxChartMetricRules.LegendFullFrameBottomBaselineFactor,
            PptxSceneChartLegendPosition.Top when fillLegendInFullFrame => frame.Y + frame.Height - lineHeight * PptxChartMetricRules.LegendFullFrameTopBaselineFactor,
            PptxSceneChartLegendPosition.Bottom => frame.Y + PptxChartMetricRules.LegendBottomFramePad + fontSize * PptxChartMetricRules.LegendBottomBaselineFontFactor,
            PptxSceneChartLegendPosition.Top => plotBox.Y + plotBox.Height + lineHeight * PptxChartMetricRules.LegendTopOffsetFactor,
            _ when sideStrokeLegend => plotBox.Y + plotBox.Height / 2d -
                fontSize * ResolveSideStrokeLegendCenterOffsetFactor(legendMarkerSizes, scatterLegend) +
                (entries.Count - 1) * lineHeight / 2d,
            _ when sideFillLegend && placement == ChartLegendPlacement.BubbleTitleRightLegend => frame.Y + frame.Height * PptxChartMetricRules.BubbleTitleRightLegendSwatchYRatio,
            _ when sideFillLegend => frame.Y + frame.Height / 2d -
                fontSize * PptxChartMetricRules.LegendSideFillBaselineCenterOffsetFactor +
                (entries.Count - 1) * lineHeight / 2d,
            // Left blocks hang below the ring center (exact on landscape probes, 0.2 residual on
            // square); right blocks keep the plot-relative shift proven by the short-plot probe.
            _ when !sideStrokeLegend && !horizontal => (useDoughnutLeftAnchor && doughnutRingCenterY > 0d ? doughnutRingCenterY : frame.Y + frame.Height / 2d) -
                doughnutLegendVerticalShift - ComputeOverlayLegendVerticalShift(markerSize, layout.Overlay) +
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

    private static double ComputeOverlayLegendVerticalShift(double markerSize, bool overlayLegend)
    {
        // Overlay legend blocks center a half marker below the frame middle on two Office
        // renders (n=4 shift 5.01, n=3 shift 5.03, fs24 middle shift 6.64): markerSize/2 lands
        // within 0.06 on all three. Sole overlay legend in corpus is doughnut-right; other
        // kinds keep legacy output through the flag, unobserved.
        return overlayLegend ? markerSize / 2d : 0d;
    }

    private static double ComputeFullFrameFillLegendLineHeight(double fontSize, bool fullFrameFillLegend, bool sideStrokeLegend, bool sideFillLegend)
    {
        return fullFrameFillLegend
            ? PptxChartMetricRules.FullFrameFillLegendLineHeightPerFontSize * fontSize + PptxChartMetricRules.FullFrameFillLegendLineHeightBase
            : fontSize * ((sideStrokeLegend || sideFillLegend)
                ? PptxChartMetricRules.LegendSideStrokeLineHeightFactor
                : PptxChartMetricRules.LegendLineHeightFactor);
    }

    private static double ComputeHorizontalLegendTextGap(double fontSize, bool bottomLegend)
    {
        return bottomLegend
            ? PptxChartMetricRules.LegendHorizontalTextGapFactor * fontSize + PptxChartMetricRules.LegendHorizontalTextGapIntercept
            : PptxChartMetricRules.LegendTextGap;
    }

    private static double ComputeHorizontalLegendInterEntryGap(double fontSize, double advanceRange, bool bottomLegend)
    {
        return bottomLegend
            ? PptxChartMetricRules.LegendHorizontalInterEntryGapFactor * fontSize - PptxChartMetricRules.LegendHorizontalInterEntryGapRangeFactor * advanceRange + PptxChartMetricRules.LegendHorizontalInterEntryGapBase
            : PptxChartMetricRules.LegendHorizontalEntryPadding;
    }

    private static double ComputeHorizontalLegendAdvanceRange(IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle style, ChartTextMeasurer textMeasurer)
    {
        double minAdvance = double.MaxValue;
        double maxAdvance = 0d;
        foreach (ChartLegendEntry entry in entries)
        {
            double advance = textMeasurer.Measure(entry.Name, style);
            if (advance < minAdvance)
            {
                minAdvance = advance;
            }

            if (advance > maxAdvance)
            {
                maxAdvance = advance;
            }
        }

        return entries.Count < 2 ? 0d : Math.Max(0d, maxAdvance - minAdvance);
    }

    private static double GetPackedHorizontalLegendEntryWidth(string name, ChartTextStyle style, ChartTextMeasurer textMeasurer, double markerSize, double textGap, double entryPadding)
    {
        return markerSize + textGap + textMeasurer.Measure(name, style) + entryPadding;
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
