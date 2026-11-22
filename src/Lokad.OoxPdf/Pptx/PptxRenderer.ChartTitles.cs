using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static async ValueTask RenderChartTitleAsync(PptxDocument document, PptxTheme theme, PdfGraphicsBuilder graphics, ShapeBounds bounds, PptxSceneChart sceneChart, PresentationFontResolver fontResolver, List<PdfFontResource> chartFonts, PptxRenderContext context, List<PdfLinkAnnotation> linkAnnotations, HashSet<string> reportedHyperlinkIds, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        string? title = sceneChart.Title.Text;
        if (string.IsNullOrWhiteSpace(title))
        {
            return;
        }

        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        double x = frame.X;
        double y = frame.Y;
        double width = frame.Width;
        double height = frame.Height;
        bool titleManualLayoutApplied = false;
        if (sceneChart.Title.Layout.HasLayout &&
            TryBuildManualLayoutBox(sceneChart.Title.Layout, frame, new ChartLayoutBox(frame.X, frame.Y, frame.Width, frame.Height), out ChartLayoutBox titleBox, true, false))
        {
            x = titleBox.X;
            y = titleBox.Y;
            width = titleBox.Width;
            height = titleBox.Height;
            titleManualLayoutApplied = true;
        }

        var (style, titleLines, measuredHeight) = await MeasureChartTitleAsync(theme, sceneChart, fontResolver).ConfigureAwait(false);
        var runs = new List<TextRun>();
        List<ChartTextRunLink>? titleLinks = sceneChart.Relationships is null ? null : new List<ChartTextRunLink>();
        await AddMeasuredChartTitleRunsAsync(runs, title, titleLines, measuredHeight,
            new ChartLayoutBox(x, y, width, height), titleManualLayoutApplied, style, fontResolver, titleLinks).ConfigureAwait(false);
        ChartShapeStyle titleShape = ToChartShapeStyle(sceneChart.Title.ShapeStyle);
        if (!titleShape.IsEmpty)
        {
            ChartLayoutBox paintBox = titleManualLayoutApplied
                ? new ChartLayoutBox(x, y, width, height)
                : await MeasureChartTitlePaintBoxAsync(runs, style, fontResolver).ConfigureAwait(false);
            RenderChartShapeStyle(graphics, paintBox.X, paintBox.Y, paintBox.Width, paintBox.Height, titleShape);
        }
        RenderedFonts titleFonts = (await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CT", fontResolver, diagnosticSink).ConfigureAwait(false));
        if (titleLinks is not null && sceneChart.Relationships is not null)
        {
            foreach (ChartTextRunLink titleLink in titleLinks)
            {
                await AddChartTextRunHyperlinkAnnotationAsync(titleLink, sceneChart.Relationships, context, sceneChart.TargetPartName, linkAnnotations, reportedHyperlinkIds, titleFonts.Fonts).ConfigureAwait(false);
            }
        }
    }

    private static async ValueTask AddMeasuredChartTitleRunsAsync(List<TextRun> runs, string title,
        IReadOnlyList<MeasuredChartTextLine> titleLines, double measuredHeight, ChartLayoutBox box,
        bool manualLayout, ChartTextStyle style, PresentationFontResolver? fontResolver, List<ChartTextRunLink>? titleLinks)
    {
        // A manual title keeps its first-line anchor if the line block fits.
        // An automatic chart title starts at the chart header.
        double top = box.Y + box.Height - PptxChartMetricRules.ChartContentFrameInset -
            PptxChartMetricRules.ChartTitleLinePadding;
        double? singleLineBaseline = null;
        if (manualLayout)
        {
            double anchor = box.Y + box.Height * PptxChartMetricRules.TitleBaselineYRatio;
            var first = titleLines[0];
            if (titleLines.Count > 1 || first.Runs.Count > 0 && first.Runs[0].LineIndex > 0)
                top = Math.Min(box.Y + box.Height, Math.Max(box.Y + measuredHeight, anchor + first.BaselineFromTop));
            else
                singleLineBaseline = anchor;
        }
        double titleX = box.X + box.Width * PptxChartMetricRules.TitleXInsetRatio;
        double titleWidth = box.Width * PptxChartMetricRules.TitleWidthRatio;
        await AddChartTitleLinesAsync(runs, title, titleLines, style, titleX, titleWidth, top,
            singleLineBaseline, box, fontResolver, titleLinks).ConfigureAwait(false);
    }

    private static async ValueTask AddChartTitleLinesAsync(List<TextRun> runs, string title,
        IReadOnlyList<MeasuredChartTextLine> titleLines, ChartTextStyle style, double x, double width,
        double top, double? singleLineBaseline, ChartLayoutBox clip, PresentationFontResolver? fontResolver,
        List<ChartTextRunLink>? titleLinks)
    {
        double titleHeight = style.FontSize * PptxChartMetricRules.TitleHeightFactor;
        foreach (MeasuredChartTextLine line in titleLines)
        {
            fontResolver?.CancellationToken.ThrowIfCancellationRequested();
            string lineText = titleLines.Count == 1 ? title : string.Concat(line.Runs.Select(run => run.Text));
            await AddChartRichTextRunsAsync(runs, line.Runs, lineText.Trim(), x,
                singleLineBaseline ?? top - line.BaselineFromTop, width, titleHeight,
                clip.X, clip.Y, clip.Width, clip.Height, style, TextAlignment.Center, fontResolver, titleLinks).ConfigureAwait(false);
        }
    }

    private static async ValueTask<ChartLayoutBox> MeasureChartTitlePaintBoxAsync(
        IReadOnlyList<TextRun> runs, ChartTextStyle style, PresentationFontResolver resolver)
    {
        var measure = new ChartTextMeasurer(resolver);
        double left = double.PositiveInfinity, bottom = double.PositiveInfinity;
        double right = double.NegativeInfinity, top = double.NegativeInfinity;
        foreach (TextRun run in runs)
        {
            ChartTextStyle runStyle = style with
            {
                FontFamily = run.FontFamily, FontSize = run.FontSize,
                Bold = run.Bold, Italic = run.Italic, CharacterSpacing = run.CharacterSpacing
            };
            double advance = await measure.MeasureAsync(run.Text, runStyle).ConfigureAwait(false);
            var line = await measure.MeasureTextLineBoxAsync(run.Text, runStyle).ConfigureAwait(false);
            double x = run.X + (run.Alignment switch
            {
                TextAlignment.Center => Math.Max(0d, run.Width - advance) / 2d,
                TextAlignment.Right => Math.Max(0d, run.Width - advance),
                _ => 0d
            });
            double baseline = run.Y + run.BaselineOffset;
            left = Math.Min(left, Math.Min(x, x + advance));
            right = Math.Max(right, Math.Max(x, x + advance));
            top = Math.Max(top, baseline + line.Ascent);
            bottom = Math.Min(bottom, baseline + line.Ascent - line.Height);
        }
        double horizontal = PptxChartMetricRules.ChartTitleHorizontalPadding;
        double vertical = PptxChartMetricRules.ChartTitleLinePadding;
        // Paint the measured caption box. Text and links retain their existing
        // frame clip, including glyph overhang.
        return new ChartLayoutBox(left - horizontal, bottom - vertical,
            Math.Max(1d, right - left + 2d * horizontal), Math.Max(1d, top - bottom + 2d * vertical));
    }

    private static async ValueTask RenderManualChartAxisTitlesAsync(
        PptxDocument document,
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ShapeBounds bounds,
        PptxSceneChart sceneChart,
        PresentationFontResolver fontResolver,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        List<PdfFontResource> chartFonts,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds)
    {
        ChartFrameBox frame = GetChartFrameBox(document, bounds);
        foreach (PptxSceneChartAxis axis in sceneChart.Axes)
        {
            if (!string.IsNullOrWhiteSpace(axis.Title.Text) && !axis.Title.Layout.HasLayout)
            {
                continue;
            }

            await RenderManualChartAxisTitleAsync(
                theme,
                graphics,
                frame,
                axis.Title.Text,
                ToChartTextRuns(axis.Title.TextRuns),
                axis.Title.Layout,
                ToChartShapeStyle(axis.Title.ShapeStyle),
                ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartElementTextStyleOverride(sceneChart, axis.Title.TextStyle, GetChartAxisStyleRole(axis.AxisKind))),
                fontResolver, chartFonts, context, sceneChart, linkAnnotations, reportedHyperlinkIds, diagnosticSink).ConfigureAwait(false);
        }
    }

    private static async ValueTask RenderDefaultChartAxisTitlesAsync(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartLayout layout,
        PptxSceneChart sceneChart,
        PresentationFontResolver fontResolver,
        List<PdfFontResource> chartFonts,
        PptxRenderContext context,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        for (int index = 0; index < sceneChart.Axes.Count; index++)
        {
            PptxSceneChartAxis axis = sceneChart.Axes[index];
            if (GetDefaultAxisTitleStatus(axis) != DefaultAxisTitleStatus.Supported)
            {
                continue;
            }

            if (layout.AxisTitleBands?.Titles.TryGetValue(index, out var band) == true)
            {
                await RenderDefaultAxisTitleBandAsync(graphics, sceneChart, axis, band, fontResolver, chartFonts,
                    context, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
                continue;
            }

            await RenderDefaultChartAxisTitleAsync(
                graphics,
                layout,
                axis.Title.Text!,
                ToChartTextRuns(axis.Title.TextRuns),
                axis.PositionKind,
                ToChartShapeStyle(axis.Title.ShapeStyle),
                ReadDefaultAxisTitleStyle(theme, sceneChart, axis),
                fontResolver, chartFonts, context, sceneChart, linkAnnotations, reportedHyperlinkIds, diagnosticSink).ConfigureAwait(false);
        }
    }

    private static void EmitUnrenderedDefaultChartAxisTitleDiagnostics(
        PptxSceneChart sceneChart,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        string? chartPartName,
        int slideIndex)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        foreach (PptxSceneChartAxis axis in sceneChart.Axes)
        {
            if (GetDefaultAxisTitleStatus(axis) != DefaultAxisTitleStatus.Unsupported)
            {
                continue;
            }

            EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART_AXIS_TITLE_AXIS_POSITION", OoxPdfSeverity.Warning, "Default-placement chart axis title has an unsupported or missing axis kind/position.", chartPartName, slideIndex, "Ignored");
        }
    }

    private static ChartTextStyle ReadDefaultAxisTitleStyle(PptxTheme theme, PptxSceneChart chart, PptxSceneChartAxis axis)
    {
        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize);
        PptxSceneChartTextStyleOverride title = axis.Title.TextStyle;
        style = style.Merge(ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartElementTextStyleOverride(
            chart, title, GetChartAxisStyleRole(axis.AxisKind))));
        if (title.Bold is null) style = style with { Bold = true };
        if (title.Color is null) style = style with { Color = new RgbColor(0, 0, 0) };
        return style;
    }

    private static async ValueTask RenderDefaultChartAxisTitleAsync(
        PdfGraphicsBuilder graphics,
        ChartLayout layout,
        string text,
        IReadOnlyList<ChartTextRunOverride> textRuns,
        PptxSceneChartAxisPosition positionKind,
        ChartShapeStyle shapeStyle,
        ChartTextStyle style,
        PresentationFontResolver fontResolver,
        List<PdfFontResource> chartFonts,
        PptxRenderContext context,
        PptxSceneChart? chart,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        string trimmed = text.Trim();
        double titleHeight = style.FontSize * PptxChartMetricRules.TitleHeightFactor;
        var textMeasurer = new ChartTextMeasurer(fontResolver);
        double textWidth = Math.Max(style.FontSize, (await textMeasurer.MeasureAsync(trimmed, style).ConfigureAwait(false)));
        if (positionKind is PptxSceneChartAxisPosition.Bottom or PptxSceneChartAxisPosition.Top)
        {
            await RenderDefaultHorizontalChartAxisTitleAsync(graphics, layout, trimmed, textRuns, style, shapeStyle, titleHeight, textWidth, positionKind, fontResolver, chartFonts, context, chart, linkAnnotations, reportedHyperlinkIds, diagnosticSink).ConfigureAwait(false);
        }
        else if (positionKind is PptxSceneChartAxisPosition.Left or PptxSceneChartAxisPosition.Right)
        {
            await RenderDefaultVerticalChartAxisTitleAsync(graphics, layout, trimmed, textRuns, style, shapeStyle, titleHeight, textWidth, positionKind, fontResolver, chartFonts, context, chart, linkAnnotations, reportedHyperlinkIds, diagnosticSink).ConfigureAwait(false);
        }
    }

    // Bottom default axis-title baseline: frame bottom plus 0.338 times the
    // plot-to-frame reserve (two Office renders agree within 0.01pt).
    private static double ComputeDefaultBottomAxisTitleBaselineY(double frameBottomY, double plotBottomY)
    {
        double bottomReserve = Math.Max(0d, plotBottomY - frameBottomY);
        return frameBottomY + bottomReserve * PptxChartMetricRules.DefaultAxisTitleBandBaselineRatio;
    }

    private static double ComputeDefaultHorizontalAxisTitleBaselineY(ChartLayout layout, PptxSceneChartAxisPosition side)
    {
        ChartFrameBox frame = layout.Frame;
        ChartPlotBox plot = layout.PlotBox;
        if (side == PptxSceneChartAxisPosition.Bottom)
            return ComputeDefaultBottomAxisTitleBaselineY(frame.Y, plot.Y);

        double plotTop = plot.Y + plot.Height;
        double topReserve = Math.Max(0d, frame.Y + frame.Height - plot.Y - plot.Height);
        return plotTop + topReserve * PptxChartMetricRules.DefaultAxisTitleTopBandBaselineRatio;
    }

    private enum DefaultAxisTitleStatus
    {
        None,
        Supported,
        Unsupported
    }

    private static DefaultAxisTitleStatus GetDefaultAxisTitleStatus(PptxSceneChartAxis axis)
    {
        if (axis.IsDeleted == true || string.IsNullOrWhiteSpace(axis.Title.Text) || axis.Title.Layout.HasLayout)
        {
            return DefaultAxisTitleStatus.None;
        }

        PptxSceneChartAxisKind axisKind = axis.AxisKind;
        bool supported = axis.PositionKind switch
        {
            PptxSceneChartAxisPosition.Bottom or PptxSceneChartAxisPosition.Top =>
                axisKind is PptxSceneChartAxisKind.Category or PptxSceneChartAxisKind.Date or PptxSceneChartAxisKind.Series or PptxSceneChartAxisKind.Value,
            PptxSceneChartAxisPosition.Left or PptxSceneChartAxisPosition.Right =>
                axisKind is PptxSceneChartAxisKind.Value or PptxSceneChartAxisKind.Category or PptxSceneChartAxisKind.Date,
            _ => false
        };
        return supported ? DefaultAxisTitleStatus.Supported : DefaultAxisTitleStatus.Unsupported;
    }

    private static string GetChartAxisStyleRole(PptxSceneChartAxisKind axisKind)
    {
        return axisKind switch
        {
            PptxSceneChartAxisKind.Value => "valueAxis",
            PptxSceneChartAxisKind.Category or PptxSceneChartAxisKind.Date or PptxSceneChartAxisKind.Series => "categoryAxis",
            _ => string.Empty
        };
    }

    private static async ValueTask RenderDefaultHorizontalChartAxisTitleAsync(
        PdfGraphicsBuilder graphics,
        ChartLayout layout,
        string text,
        IReadOnlyList<ChartTextRunOverride> textRuns,
        ChartTextStyle style,
        ChartShapeStyle shapeStyle,
        double titleHeight,
        double textWidth,
        PptxSceneChartAxisPosition positionKind,
        PresentationFontResolver? fontResolver,
        List<PdfFontResource> chartFonts,
        PptxRenderContext context,
        PptxSceneChart? chart,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        ChartFrameBox frame = layout.Frame;
        ChartPlotBox plotBox = layout.PlotBox;
        double titleX = plotBox.X + plotBox.Width / 2d - textWidth / 2d;
        double baselineY = ComputeDefaultHorizontalAxisTitleBaselineY(layout, positionKind);
        double boxBottom = positionKind == PptxSceneChartAxisPosition.Top ? plotBox.Y + plotBox.Height : frame.Y;
        double boxY = Math.Max(boxBottom, baselineY - titleHeight);

        double boxX = Math.Max(frame.X, titleX - style.FontSize * 0.5d);
        double boxWidth = Math.Min(frame.X + frame.Width - boxX, textWidth + style.FontSize);
        RenderChartShapeStyle(graphics, boxX, boxY, boxWidth, titleHeight, shapeStyle);
        var runs = new List<TextRun>();
        List<ChartTextRunLink>? axisLinks = chart?.Relationships is null ? null : new List<ChartTextRunLink>();
        await AddChartRichTextRunsAsync(
            runs,
            textRuns,
            text,
            titleX,
            baselineY,
            textWidth,
            titleHeight,
            frame.X,
            frame.Y,
            frame.Width,
            frame.Height,
            style,
            TextAlignment.Center,
            fontResolver,
            axisLinks).ConfigureAwait(false);
        RenderedFonts axisFonts = (await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CAT", fontResolver, diagnosticSink).ConfigureAwait(false));
        await AddChartTextRunHyperlinkAnnotationsAsync(axisLinks, chart, context, axisFonts, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
    }

    private static async ValueTask RenderDefaultVerticalChartAxisTitleAsync(
        PdfGraphicsBuilder graphics,
        ChartLayout layout,
        string text,
        IReadOnlyList<ChartTextRunOverride> textRuns,
        ChartTextStyle style,
        ChartShapeStyle shapeStyle,
        double titleHeight,
        double textWidth,
        PptxSceneChartAxisPosition positionKind,
        PresentationFontResolver? fontResolver,
        List<PdfFontResource> chartFonts,
        PptxRenderContext context,
        PptxSceneChart? chart,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        ChartFrameBox frame = layout.Frame;
        ChartPlotBox plotBox = layout.PlotBox;
        double baselineX;
        if (positionKind == PptxSceneChartAxisPosition.Right)
        {
            double rightReserve = Math.Max(0d, frame.X + frame.Width - plotBox.X - plotBox.Width);
            baselineX = plotBox.X + plotBox.Width + rightReserve * PptxChartMetricRules.DefaultAxisTitleRightSideBaselineRatio;
        }
        else
        {
            baselineX = frame.X + PptxChartMetricRules.DefaultAxisTitleLeftColumnOffset;
        }
        // Office centers the title ink on the plot middle: the TJ origin sits exactly
        // at mid minus half the EMITTED width on three cached references (bar 253.9 vs
        // 253.7, column 280.95 vs 280.8, top-right 245.09 vs 245.0). The caller-measured
        // textWidth uses the chart-level size. Use the same measured run styles for
        // centering and emission.
        List<ChartTextRunLayout> measuredRuns = await MeasureChartRichTextRunsAsync(
            textRuns, style, new ChartTextMeasurer(fontResolver)).ConfigureAwait(false);
        double emittedTitleWidth = measuredRuns.Sum(run => run.Width);

        double centeredTitleWidth = emittedTitleWidth > 0d ? emittedTitleWidth : textWidth;
        double baselineY = plotBox.Y + plotBox.Height / 2d - centeredTitleWidth / 2d;
        double boxX = Math.Max(frame.X, Math.Min(frame.X + frame.Width - titleHeight, baselineX - titleHeight * 0.5d));
        double boxY = Math.Max(frame.Y, Math.Min(frame.Y + frame.Height - textWidth, baselineY));
        RenderChartShapeStyle(graphics, boxX, boxY, titleHeight, textWidth, shapeStyle);
        var runs = new List<TextRun>();
        List<ChartTextRunLink>? axisLinks = chart?.Relationships is null ? null : new List<ChartTextRunLink>();
        AddMeasuredChartTextRuns(runs, measuredRuns, emittedTitleWidth, text, baselineX, baselineY, textWidth,
            titleHeight, new ChartLayoutBox(frame.X, frame.Y, frame.Width, frame.Height), style, TextAlignment.Left, axisLinks);
        RotateChartAxisTitle(runs, axisLinks, baselineX, baselineY);
        RenderedFonts axisFonts = (await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CAT", fontResolver, diagnosticSink).ConfigureAwait(false));
        await AddChartTextRunHyperlinkAnnotationsAsync(axisLinks, chart, context, axisFonts, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
    }

    private static void RotateChartAxisTitle(List<TextRun> runs, List<ChartTextRunLink>? axisLinks,
        double pivotX, double pivotY)
    {
        for (int i = 0; i < runs.Count; i++)
        {
            runs[i] = PlaceInAxisFrame(runs[i]);
        }

        if (axisLinks is not null)
        {
            for (int linkIndex = 0; linkIndex < axisLinks.Count; linkIndex++)
            {
                ChartTextRunLink axisLink = axisLinks[linkIndex];
                axisLinks[linkIndex] = axisLink with { Run = PlaceInAxisFrame(axisLink.Run) };
            }
        }

        // The same physical frame owns all runs and their linked ink.
        TextRun PlaceInAxisFrame(TextRun run) => run with
        {
            RotationDegrees = -90d,
            RotationCenterX = pivotX,
            RotationCenterY = pivotY
        };
    }

    private static async ValueTask RenderManualChartAxisTitleAsync(
        PptxTheme theme,
        PdfGraphicsBuilder graphics,
        ChartFrameBox frame,
        string? text,
        IReadOnlyList<ChartTextRunOverride> textRuns,
        PptxSceneChartManualLayout layout,
        ChartShapeStyle shapeStyle,
        ChartTextStyleOverride resolvedStyle,
        PresentationFontResolver? fontResolver,
        List<PdfFontResource> chartFonts,
        PptxRenderContext context,
        PptxSceneChart? chart,
        List<PdfLinkAnnotation> linkAnnotations,
        HashSet<string> reportedHyperlinkIds,
        Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        if (string.IsNullOrWhiteSpace(text) ||
            !TryBuildManualLayoutBox(layout, frame, new ChartLayoutBox(frame.X, frame.Y, frame.Width, frame.Height), out ChartLayoutBox titleBox, true, false))
        {
            return;
        }

        ChartTextStyle style = CreateDefaultChartTextStyle(theme, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize)
            .Merge(resolvedStyle);
        RenderChartShapeStyle(graphics, titleBox.X, titleBox.Y, titleBox.Width, titleBox.Height, shapeStyle);
        var runs = new List<TextRun>();
        List<ChartTextRunLink>? axisLinks = chart?.Relationships is null ? null : new List<ChartTextRunLink>();
        var (titleLines, measuredHeight) = await MeasureChartRichTextLinesAsync(text, textRuns, style, fontResolver).ConfigureAwait(false);
        await AddMeasuredChartTitleRunsAsync(runs, text, titleLines, measuredHeight, titleBox,
            manualLayout: true, style, fontResolver, axisLinks).ConfigureAwait(false);
        RenderedFonts axisFonts = (await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CAT", fontResolver, diagnosticSink).ConfigureAwait(false));
        await AddChartTextRunHyperlinkAnnotationsAsync(axisLinks, chart, context, axisFonts, linkAnnotations, reportedHyperlinkIds).ConfigureAwait(false);
    }

    private static ChartAxisTitleReserveSides ReadDefaultAxisTitleReserveSides(PptxSceneChart sceneChart)
    {
        var reserveSides = new ChartAxisTitleReserveSides();
        foreach (PptxSceneChartAxis axis in sceneChart.Axes)
        {
            if (GetDefaultAxisTitleStatus(axis) != DefaultAxisTitleStatus.Supported)
            {
                continue;
            }

            reserveSides = reserveSides.With(axis.PositionKind);
        }

        return reserveSides;
    }

    private static async ValueTask<(double Reserve, ChartAxisSource ValueAxis, double NumericClearance)?> TryMeasureCartesianHeaderAsync(
        PptxTheme theme, PptxSceneChart chart, ChartLegendLayout legend,
        PptxSceneChartPlot? plot, PresentationFontResolver? resolver,
        ChartAxisPlacement? categoryPlacement = null, bool automaticScaleAllocation = false)
    {
        // Authored frames and furniture above/beside the plot keep their existing
        // placement owner. This path measures the ordinary numeric-axis header.
        ChartAxisTitleReserveSides titles = ReadDefaultAxisTitleReserveSides(chart);
        if (chart.PlotAreaLayout.HasGeometryValues &&
                !(automaticScaleAllocation && chart.PlotAreaLayout.LayoutTargetKind == PptxSceneChartManualLayoutTarget.Inner) ||
            chart.Title.Layout.HasGeometryValues || legend.Layout.HasGeometryValues ||
            titles.HasHorizontalTitle || titles.HasVerticalTitle ||
            legend.Visible && !legend.Overlay && legend.PositionKind != PptxSceneChartLegendPosition.Bottom)
            return null;
        IReadOnlyList<ChartAxisSource> axes = ReadChartValueAxesForPlot(chart, plot);
        if (axes.Count != 1 || axes[0].SceneAxis is not { } axis ||
            axis.PositionKind is not (PptxSceneChartAxisPosition.Left or PptxSceneChartAxisPosition.Right))
            return null;
        // One reciprocal shared scale also owns one measured header and
        // category band. Unsupported composed axes retain their fallback.
        if (chart.Plots.Count != 1 && !HasSharedCartesianChartValueAxis(chart, axis)) return null;
        ChartAxisSource category = ReadChartCategoryAxisForPlot(chart, plot);
        if (categoryPlacement is null && IsSceneChartAxisLabelVisible(category.SceneAxis) &&
            (category.SceneAxis is not { PositionKind: PptxSceneChartAxisPosition.Bottom } categoryAxis ||
             categoryAxis.TickLabelPositionKind == PptxSceneChartTickLabelPosition.High ||
             ReadSceneValueAxisReversed(axis) && categoryAxis.TickLabelPositionKind != PptxSceneChartTickLabelPosition.Low))
            return null;
        double lineHeight = 0d;
        if (IsSceneChartAxisLabelVisible(axis))
        {
            ChartTextStyle style = ReadSceneChartTextStyle(theme, chart, axis,
                fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
            var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
            if (await measure.TryMeasureNormalLineHeightAsync(style).ConfigureAwait(false) is not { } measured)
                return null;
            lineHeight = measured;
        }
        // Keep the numeric end clearance available to both ends of the frame;
        // the caption contributes only to the header.
        var header = await MeasureCartesianHeaderReserveAsync(theme, chart, resolver, lineHeight).ConfigureAwait(false);
        return (header.Reserve, axes[0], header.NumericClearance);
    }

    private static async ValueTask<(double Reserve, double NumericClearance)> MeasureCartesianHeaderReserveAsync(
        PptxTheme theme, PptxSceneChart chart, PresentationFontResolver? resolver, double numericLineHeight)
    {
        double numericClearance = CartesianAxisEndClearance(numericLineHeight);
        double reserve = numericClearance;
        if (!string.IsNullOrWhiteSpace(chart.Title.Text) && chart.Title.Overlay != true)
        {
            var title = await MeasureChartTitleAsync(theme, chart, resolver).ConfigureAwait(false);
            reserve += PptxChartMetricRules.ChartContentFrameInset +
                2d * PptxChartMetricRules.ChartTitleLinePadding + title.Height;
        }
        return (reserve, numericClearance);
    }

    private static double CartesianAxisEndClearance(double lineHeight) =>
        Math.Max(PptxChartMetricRules.HorizontalBarPlotTopReserve,
            PptxChartMetricRules.AutomaticCartesianFramePadding + lineHeight / 2d);

    private static async ValueTask<double?> MeasureCategoryFittingClearanceAsync(PptxTheme theme,
        PptxSceneChart chart, PresentationFontResolver? resolver, double numericClearance)
    {
        // Built-in charts retain a default fitting reference when overridden
        // role/global fonts change. Those unused sizes must not alter the frame.
        // Actual numeric typography still consumes space at both ends; painting
        // retains its effective fonts. External/custom styles keep their fallback.
        ChartTextStyle reference = chart.GalleryStyle.HasValue
            ? CreateDefaultChartTextStyle(theme, chart.ColorMap, PptxChartMetricRules.BuiltInAxisFittingReferenceFontSize)
            : ReadSceneChartTextStyle(theme, chart, null,
                PptxChartMetricRules.CategoryAxisFallbackFontSize, "categoryAxis");
        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        if (await measure.TryMeasureNormalLineHeightAsync(reference).ConfigureAwait(false) is not { } lineHeight)
            return null;
        return 2d * CartesianAxisEndClearance(lineHeight) - numericClearance;
    }

    private readonly record struct MeasuredChartTextLine(
        IReadOnlyList<ChartTextRunOverride> Runs, double BaselineFromTop, double Ascent);

    private static async ValueTask<(ChartTextStyle Style, IReadOnlyList<MeasuredChartTextLine> Lines, double Height)> MeasureChartTitleAsync(
        PptxTheme theme, PptxSceneChart sceneChart, PresentationFontResolver? resolver)
    {
        ChartTextStyle style = ReadChartTitleTextStyle(theme, sceneChart);
        var (lines, height) = await MeasureChartRichTextLinesAsync(sceneChart.Title.Text ?? string.Empty,
            ReadChartTitleTextRuns(sceneChart), style, resolver).ConfigureAwait(false);
        return (style, lines, height);
    }

    private static async ValueTask<(IReadOnlyList<MeasuredChartTextLine> Lines, double Height)> MeasureChartRichTextLinesAsync(
        string text, IReadOnlyList<ChartTextRunOverride> runs, ChartTextStyle style, PresentationFontResolver? resolver)
    {
        var measure = new ChartTextMeasurer(resolver);
        if (runs.Count == 0)
        {
            var line = await measure.MeasureTextLineBoxAsync(text, style).ConfigureAwait(false);
            return ([new([], line.Ascent, line.Ascent)], line.Height);
        }

        var lines = new List<MeasuredChartTextLine>();
        double height = 0d;
        double? emptyLineHeight = null;
        int nextLine = 0;
        bool hasLineBreaks = runs[^1].LineIndex > 0;
        foreach (var line in runs.GroupBy(run => run.LineIndex))
        {
            if (line.Key > nextLine)
            {
                emptyLineHeight ??= (await measure.MeasureLineBoxAsync(style, horizontalHeader: true).ConfigureAwait(false)).Height;
                height += (line.Key - nextLine) * emptyLineHeight.Value;
            }
            double ascent = 0d, descent = 0d, lineHeight = 0d;
            foreach (ChartTextRunOverride run in line)
            {
                resolver?.CancellationToken.ThrowIfCancellationRequested();
                var metrics = await measure.MeasureTextLineBoxAsync(run.Text, style.Merge(run.TextStyle)).ConfigureAwait(false);
                ascent = Math.Max(ascent, metrics.Ascent);
                descent = Math.Max(descent, metrics.Height - metrics.Ascent);
                lineHeight = Math.Max(lineHeight, metrics.Height);
            }
            lines.Add(new(line.ToArray(), height + ascent, ascent));
            height += hasLineBreaks ? ascent + descent : lineHeight;
            nextLine = line.Key + 1;
        }
        return (lines, height);
    }

    private static ChartTextStyle ReadChartTitleTextStyle(PptxTheme theme, PptxSceneChart sceneChart)
    {
        ChartTextStyle style = CreateDefaultChartTextStyle(theme, sceneChart.ColorMap, fallbackFontSize: PptxChartMetricRules.TitleFallbackFontSize);
        style = style.Merge(ToChartTextStyleOverride(PptxSceneBuilder.ResolveChartTitleTextStyleOverride(sceneChart)));
        if (!sceneChart.Title.IsAutoGenerated)
        {
            return style;
        }

        return style with
        {
            Bold = true
        };
    }

    private static IReadOnlyList<ChartTextRunOverride> ReadChartTitleTextRuns(PptxSceneChart sceneChart)
    {
        return sceneChart.Title.IsAutoGenerated ? [] : ToChartTextRuns(sceneChart.Title.TextRuns);
    }

}
