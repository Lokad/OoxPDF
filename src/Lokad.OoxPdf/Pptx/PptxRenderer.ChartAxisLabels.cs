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
    private static async ValueTask<double?> TryMeasureCategoryLabelOffsetAdjustmentAsync(
        ChartTextMeasurer measurer, ChartTextStyle style, double offsetScale)
    {
        if (offsetScale == 1d || await measurer.TryMeasureNormalLineHeightAsync(style).ConfigureAwait(false) is not { } lineHeight)
            return null;
        return lineHeight * PptxChartMetricRules.CategoryAxisLabelOffsetNormalLineHeightRatio * (offsetScale - 1d);
    }

    private static async ValueTask RenderChartCategoryLabelsAsync(PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, PptxSceneChartAxis? sceneAxis, XElement? categoryAxis, IReadOnlyList<ChartIndexedTextPoint?> labels, bool horizontalBars, double? axisCoordinate, bool categoryLabelsOnTickMarks, bool labelsHighSide, PresentationFontResolver? fontResolver, List<PdfFontResource> chartFonts, Action<OoxPdfDiagnostic>? diagnosticSink = null, double? categoryLabelOffsetAdjustment = null, ChartCategoryLabelLayout? prepared = null, bool categoryAxisReversed = false)
    {
        if (labels.Count == 0)
        {
            return;
        }

        ChartTextStyle style = prepared?.Style ?? ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, categoryAxis, fallbackFontSize: PptxChartMetricRules.CategoryAxisFallbackFontSize, chartStyleRole: "categoryAxis");
        double fontSize = style.FontSize;
        RgbColor color = style.Color;
        double labelOffsetScale = ResolveSceneOrXmlCategoryAxisLabelOffsetScale(sceneAxis, categoryAxis);
        int tickLabelSkip = prepared?.TickLabelSkip ?? ResolveSceneOrXmlCategoryAxisTickLabelSkip(sceneAxis, categoryAxis);
        double widestCategoryLabel = 0d;
        if (horizontalBars)
        {
            var stripMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
            for (int labelIndex = 0; labelIndex < labels.Count; labelIndex++)
            {
                if (labelIndex % tickLabelSkip != 0)
                {
                    continue;
                }

                string? stripLabel = labels[labelIndex]?.Text;
                if (string.IsNullOrWhiteSpace(stripLabel))
                {
                    continue;
                }

                widestCategoryLabel = Math.Max(widestCategoryLabel, (await stripMeasurer.MeasureAsync(stripLabel, style).ConfigureAwait(false)));
            }
        }
        var runs = new List<TextRun>(labels.Count);
        for (int i = 0; i < labels.Count; i++)
        {
            if (i % tickLabelSkip != 0)
            {
                continue;
            }

            string? label = labels[i]?.Text;
            if (string.IsNullOrWhiteSpace(label))
            {
                continue;
            }

            if (prepared is { RotationDegrees: not 0d } rotated && rotated.Labels.TryGetValue(i, out var rotatedParts))
            {
                foreach (var part in rotatedParts)
                {
                    if (part.Text.Length == 0) continue;
                    double center = ChartCategoryToPlotCoordinate(i, labels.Count, plotBox.X, plotBox.Width, categoryLabelsOnTickMarks, categoryAxisReversed);
                    double axisY = axisCoordinate ?? plotBox.Y;
                    double gap = rotated.AxisGap;
                    double radians = -rotated.RotationDegrees * Math.PI / 180d;
                    double sin = Math.Sin(radians), cos = Math.Cos(radians);
                    bool rightAnchor = (radians > 0d) != labelsHighSide;
                    // A trailing marker hangs toward the axis only when the
                    // trailing end anchors there. At the opposite end it moves
                    // with the caption band. Literal ellipses have no overhang.
                    double overhang = rightAnchor ? part.OverflowAdvance * Math.Abs(sin) : 0d;
                    double labelAxisY = axisY + (labelsHighSide ? -overhang : overhang);
                    var horizontal = ChartCategoryCaptionHorizontalFrame(center, part.Width, rotated.Ascent, rotated.Descent, sin, cos, rightAnchor);
                    double pivotX = horizontal.Pivot;
                    double pivotY = labelsHighSide ? labelAxisY + gap + rotated.Descent * cos : labelAxisY - gap - rotated.Ascent * cos;
                    double localX = rightAnchor ? pivotX - part.Width : pivotX;
                    double clipLeft = horizontal.Left;
                    double clipRight = horizontal.Right;
                    double extent = part.Width * Math.Abs(sin) + rotated.LineHeight * cos;
                    double clipBottom = labelsHighSide ? labelAxisY + gap : labelAxisY - gap - extent;
                    TextRun rotatedRun = CreateChartTextRun(part.Text, localX, pivotY, part.Width, rotated.LineHeight,
                        clipLeft, clipBottom, Math.Max(1d, clipRight - clipLeft), extent,
                        style with { Alpha = 1d }, TextAlignment.Left, kerningEnabled: false);
                    runs.Add(rotatedRun with { RotationDegrees = rotated.RotationDegrees, RotationCenterX = pivotX, RotationCenterY = pivotY });
                }
                continue;
            }

            double x;
            double y;
            double width;
            double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
            TextAlignment alignment;
            if (horizontalBars)
            {
                double anchor = axisCoordinate ?? plotBox.X;
                double gap = PptxChartMetricRules.HorizontalBarCategoryLabelPlotGapFactor * fontSize * labelOffsetScale;
                y = ChartCategoryToPlotCoordinate(i, labels.Count, plotBox.Y, plotBox.Height,
                    onTickMarks: false, categoryAxisReversed) - height * PptxChartMetricRules.CategoryAxisHorizontalBaselineRatio;
                width = Math.Max(1d, widestCategoryLabel);
                x = Math.Max(0d, labelsHighSide ? anchor + gap : anchor - gap - width);
                alignment = labelsHighSide ? TextAlignment.Left : TextAlignment.Right;
            }
            else
            {
                double slotWidth = plotBox.Width / labels.Count;
                // Skipped categories leave their label space available to the
                // next visible tick. Keep source tick positions unchanged.
                width = slotWidth * Math.Min(tickLabelSkip, labels.Count) * PptxChartMetricRules.CategoryAxisVerticalWidthFactor;
                double labelCenterX = ChartCategoryToPlotCoordinate(i, labels.Count, plotBox.X, plotBox.Width, categoryLabelsOnTickMarks, categoryAxisReversed);
                x = labelCenterX - width / 2d;
                double axisY = axisCoordinate ?? plotBox.Y;
                double baseline = ChartCategoryLabelBaselineDistance(fontSize, labelsHighSide, labelOffsetScale, categoryLabelOffsetAdjustment);
                y = labelsHighSide ? axisY + baseline : axisY - baseline;
                alignment = TextAlignment.Center;
            }

            double labelWidth = Math.Max(1d, width);
            TextRun run = new TextRun(
                label,
                x,
                y,
                labelWidth,
                height,
                x,
                y - height * PptxChartMetricRules.AxisLabelClipTopOffsetFactor,
                labelWidth,
                height * PptxChartMetricRules.AxisLabelClipHeightFactor,
                fontSize,
                style.CharacterSpacing,
                0d,
                color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: style.Underline,
                Strike: style.Strike,
                KerningEnabled: false,
                alignment,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false, PreventCoalesce: false, Outline: null, StrictClip: false);
            if (prepared is not null && prepared.Labels.TryGetValue(i, out var parts))
            {
                for (int lineIndex = 0; lineIndex < parts.Length; lineIndex++)
                {
                    double shift = lineIndex * prepared.LineHeight * (labelsHighSide ? 1d : -1d);
                    runs.Add(run with { Text = parts[lineIndex].Text, Y = run.Y + shift, ClipY = run.ClipY + shift });
                }
            }
            else runs.Add(run);
        }

        await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CCA", fontResolver, diagnosticSink).ConfigureAwait(false);
    }

    private static async ValueTask<(IReadOnlyList<double> Values, double MaximumWidth, double SlotWidth, double SideGap)> MeasureChartValueAxisCaptionsAsync(
        ChartTextMeasurer measurer, ChartTextStyle style, ChartAxisSource source, ChartValueExtents extents,
        ChartAxisUnits units, double target, ChartAxisTickPolicy policy, string? defaultNumberFormat, bool reserveEndpoints, PptxRenderContext? context = null)
    {
        IReadOnlyList<double> values = GetChartAxisTickValuesCore(extents, units.MajorUnit, includeEndpoints: true, target, policy, context: context);
        IEnumerable<double> measured = reserveEndpoints ? values.Append(extents.Min).Append(extents.Max) : values;
        double maximum = 0d, normalHeight = 0d;
        foreach (string label in measured.Select(value => FormatSceneOrXmlChartAxisLabel(value, source.SceneAxis, source.XmlAxis, defaultNumberFormat)).DefaultIfEmpty("0"))
        {
            maximum = Math.Max(maximum, await measurer.MeasureAsync(label, style).ConfigureAwait(false));
            var line = await measurer.MeasureTextLineBoxAsync(label, style, horizontalHeader: false).ConfigureAwait(false);
            if (line.HasMetrics) normalHeight = Math.Max(normalHeight, line.Height);
        }
        double sideGap = Math.Max(3d, normalHeight > 0d
            ? normalHeight * PptxChartMetricRules.ValueAxisLabelSideGapLineHeightRatio
            : style.FontSize * PptxChartMetricRules.ValueAxisLabelSideGapFallbackFontSizeRatio);
        return (values, maximum, Math.Max(style.FontSize * PptxChartMetricRules.ValueAxisMinimumLabelWidthFactor,
            maximum + style.FontSize * PptxChartMetricRules.ValueAxisLabelPaddingFactor), sideGap);
    }

    private static async ValueTask RenderChartValueAxisLabelsAsync(PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, XElement? valueAxis, PptxSceneChartAxis? sceneAxis, ChartValueExtents extents, ChartAxisUnits axisUnits, bool valueAxisReversed, bool horizontalBars, bool rightSide, int axisSideSlot, bool manualPlotLayoutApplied, bool useTextSizedWidth, string? defaultNumberFormat, PresentationFontResolver? fontResolver, List<PdfFontResource> chartFonts, Action<OoxPdfDiagnostic>? diagnosticSink = null, double? innerStripInkEdge = null, ChartAxisTickPolicy tickPolicy = ChartAxisTickPolicy.ZeroAlignedWithEndpoints, ChartAxisPlacement? placement = null, PptxRenderContext? context = null)
    {
        ChartTextStyle style = ReadSceneOrXmlChartTextStyle(theme, sceneChart, sceneAxis, chartXml, valueAxis, fallbackFontSize: PptxChartMetricRules.ValueAxisFallbackFontSize, chartStyleRole: "valueAxis");
        double fontSize = style.FontSize;
        double height = fontSize * PptxChartMetricRules.AxisLabelHeightFactor;
        RgbColor color = style.Color;
        var textMeasurer = new ChartTextMeasurer(fontResolver, kerningEnabled: false);
        double autoTickTargetCount = GetValueAxisAutoTickTargetCount(horizontalBars, valueAxisLabelsVisible: true, manualPlotLayoutApplied);
        var captions = await MeasureChartValueAxisCaptionsAsync(textMeasurer, style, new(sceneAxis, valueAxis), extents,
            axisUnits, autoTickTargetCount, tickPolicy, defaultNumberFormat, reserveEndpoints: false, context: context).ConfigureAwait(false);
        IReadOnlyList<double> tickValues = captions.Values;
        double valueAxisLabelWidth = captions.SlotWidth;
        double sideGap = captions.SideGap;
        var runs = new List<TextRun>(tickValues.Count);
        foreach (double value in tickValues)
        {
            string label = FormatSceneOrXmlChartAxisLabel(value, sceneAxis, valueAxis, defaultNumberFormat);
            double offset = GetChartValuePlotRatio(extents, value, valueAxisReversed);
            double x;
            double y;
            double width;
            TextAlignment alignment;
            if (horizontalBars)
            {
                // A numeric caption is one measured fragment, even when wider
                // than its nominal slot. Keep its center on the value coordinate.
                width = Math.Max(plotBox.Width / PptxChartMetricRules.HorizontalValueAxisSlotCount, textMeasurer.Measure(label, style));
                x = plotBox.X + plotBox.Width * offset - width / 2d;
                double axisY = plotBox.Y + plotBox.Height * (placement?.LabelPlotRatio ?? 0d);
                y = placement?.LabelsHighSide == true
                    ? axisY + height * PptxChartMetricRules.CategoryAxisVerticalTopSideOffsetFactor
                    : axisY - height * PptxChartMetricRules.HorizontalValueAxisTopOffsetFactor;
                alignment = TextAlignment.Center;
            }
            else
            {
                bool labelsRightSide = placement?.LabelsHighSide ?? ResolveSceneOrXmlValueAxisLabelsRightSide(sceneChart, sceneAxis, valueAxis, rightSide);
                double axisX = placement is { } resolved ? plotBox.X + plotBox.Width * resolved.LabelPlotRatio
                    : labelsRightSide ? plotBox.X + plotBox.Width : plotBox.X;
                width = useTextSizedWidth ? valueAxisLabelWidth : Math.Max(valueAxisLabelWidth,
                    plotBox.Width * PptxChartMetricRules.VerticalValueAxisWidthRatio);
                // Outer same-side strips anchor to the inner strip ink edge: Office keeps a
                // fixed sideGap-plus-inter-axis gap between the ink blocks instead of chaining
                // text-sized slot boxes (three dual-column probes agree within 0.13 with zero fitting).
                // Right-side and deeper slots keep the legacy slot boxes (unobserved).
                x = labelsRightSide
                    ? axisX + sideGap + axisSideSlot * (width + sideGap)
                    : innerStripInkEdge.HasValue && axisSideSlot == 1
                        ? innerStripInkEdge.Value - sideGap - PptxChartMetricRules.BarDualValueAxisInterAxisGap - width
                        : Math.Max(0d, axisX - (axisSideSlot + 1) * (width + sideGap));
                y = plotBox.Y + plotBox.Height * offset - height * PptxChartMetricRules.VerticalValueAxisBaselineRatio;
                alignment = labelsRightSide ? TextAlignment.Left : TextAlignment.Right;
            }

            double labelWidth = Math.Max(1d, width);
            double clipY = horizontalBars ? y : plotBox.Y - height;
            double clipHeight = horizontalBars ? height : plotBox.Height + height * PptxChartMetricRules.ValueAxisVerticalClipExtraHeightFactor;
            runs.Add(new TextRun(
                label,
                x,
                y,
                labelWidth,
                height,
                x,
                clipY,
                labelWidth,
                clipHeight,
                fontSize,
                style.CharacterSpacing,
                0d,
                color,
                1d,
                null,
                Bold: style.Bold,
                Italic: style.Italic,
                Underline: style.Underline,
                Strike: style.Strike,
                KerningEnabled: false,
                alignment,
                FontFamily: style.FontFamily,
                RotationDegrees: 0d,
                RotationCenterX: 0d,
                RotationCenterY: 0d,
                FlipHorizontal: false,
                FlipVertical: false, PreventCoalesce: false, Outline: null, StrictClip: false));
        }

        if (!horizontalBars && axisUnits.MajorUnit is null && runs.Count > 2)
        {
            // Automatic labels may be thinned when the actual glyph ink would
            // overlap. Preserve both range endpoints and authored intervals.
            // Grid/scale ownership remains with the existing axis calculation.
            var ink = new List<(int Index, double Bottom, double Top)>();
            ChartTextStyle emissionStyle = style with { FontSize = OfficePdfTextEmissionProfile.FontSize(style.FontSize) };
            for (int index = 0; index < runs.Count; index++)
            {
                if (await textMeasurer.TryMeasureInkBoundsAsync(runs[index].Text, emissionStyle).ConfigureAwait(false) is not { } bounds)
                {
                    ink.Clear();
                    break;
                }
                ink.Add((index, runs[index].Y + bounds.Bottom, runs[index].Y + bounds.Top));
            }
            if (ink.Count == runs.Count)
            {
                ink.Sort((left, right) => runs[left.Index].Y.CompareTo(runs[right.Index].Y));
                var admitted = new HashSet<int> { ink[0].Index, ink[^1].Index };
                double previousTop = ink[0].Top;
                foreach (var candidate in ink.Skip(1).SkipLast(1))
                {
                    if (candidate.Bottom < previousTop || candidate.Top > ink[^1].Bottom) continue;
                    admitted.Add(candidate.Index);
                    previousTop = candidate.Top;
                }
                if (admitted.Count != runs.Count)
                    runs = runs.Where((_, index) => admitted.Contains(index)).ToList();
            }
        }
        await RenderChartTextRunsAsync(runs, graphics, chartFonts, "CVA", fontResolver, diagnosticSink).ConfigureAwait(false);
    }

    private static async ValueTask RenderSecondaryChartValueAxisLabelsAsync(PptxTheme theme, PdfGraphicsBuilder graphics, ChartPlotBox plotBox, XDocument chartXml, PptxSceneChart? sceneChart, ChartValueExtents fallback, PresentationFontResolver fontResolver, List<PdfFontResource> chartFonts, Action<OoxPdfDiagnostic>? diagnosticSink = null, PptxSceneChartAxis? primarySceneAxis = null, XElement? primaryXmlAxis = null, PptxRenderContext? context = null)
    {
        ChartAxisSource rightValueAxisSource = ReadSceneOrXmlSecondaryRightValueAxis(sceneChart, chartXml);
        if (primarySceneAxis is not null && rightValueAxisSource.SceneAxis is not null && string.Equals(primarySceneAxis.Id, rightValueAxisSource.SceneAxis.Id, StringComparison.Ordinal))
        {
            return;
        }

        if (primaryXmlAxis is not null && rightValueAxisSource.XmlAxis is not null && string.Equals(ReadChartAxisId(primaryXmlAxis), ReadChartAxisId(rightValueAxisSource.XmlAxis), StringComparison.Ordinal))
        {
            return;
        }
        XElement? rightValueAxis = rightValueAxisSource.XmlAxis;
        PptxSceneChartAxis? rightSceneAxis = rightValueAxisSource.SceneAxis;
        if (rightValueAxis is null && rightSceneAxis is null ||
            !IsSceneOrXmlChartAxisLabelVisible(rightSceneAxis, rightValueAxis))
        {
            return;
        }

        ChartValueExtents extents = ReadSceneOrXmlChartValueAxisExtents(rightSceneAxis, rightValueAxis, fallback, false, PptxChartMetricRules.AxisNiceNearMaximumHeadroomRatio);
        ChartValueAxisRenderOptions axisOptions = ReadSceneOrXmlChartValueAxisRenderOptions(rightSceneAxis, rightValueAxis, theme, extents, percentStacked: false);
        await RenderChartValueAxisLabelsAsync(theme, graphics, plotBox, chartXml, sceneChart, rightValueAxis, rightSceneAxis, extents, axisOptions.Units, axisOptions.Reversed, horizontalBars: false, rightSide: true, axisSideSlot: 0, manualPlotLayoutApplied: false, useTextSizedWidth: false, defaultNumberFormat: null, fontResolver: fontResolver, chartFonts: chartFonts, diagnosticSink: diagnosticSink, context: context).ConfigureAwait(false);
    }
}
