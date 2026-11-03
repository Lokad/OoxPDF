using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private readonly record struct ChartBubbleSizing(double MaximumSize, double Radius);

    private sealed record ChartBubbleGeometry(ChartLayout Layout, ChartValueExtents X, ChartValueExtents Y,
        ChartValueAxisRenderOptions XOptions, ChartValueAxisRenderOptions YOptions,
        ChartBubbleSizing Sizing, ChartLegendBox? LegendBox);

    private static async ValueTask<ChartBubbleGeometry?> TryPrepareBubbleGeometryAsync(PptxTheme theme,
        PptxSceneChart chart, XDocument xml, ChartLayout fallback, PptxSceneChartPlot? plot,
        IReadOnlyList<ScatterSeries> series, IReadOnlyList<ChartLegendEntry> entries,
        ChartAxisSource xAxis, ChartAxisSource yAxis, ChartValueExtents dataX, ChartValueExtents dataY,
        PresentationFontResolver resolver, PptxRenderContext? context = null)
    {
        // Authored/composed frames, width-sized/negative bubbles and rejected
        // scales retain their existing approximation until independently qualified.
        if (plot is null || chart.Plots.Count != 1 || chart.PlotAreaLayout.HasGeometryValues ||
            chart.Title.Layout.HasGeometryValues || fallback.Legend.Layout.HasGeometryValues ||
            fallback.Legend.Visible && (fallback.Legend.Overlay || fallback.Legend.PositionKind != PptxSceneChartLegendPosition.Right) ||
            !string.IsNullOrEmpty(plot.BubbleSizeRepresents) && !plot.BubbleSizeRepresents.Equals("area", StringComparison.OrdinalIgnoreCase) ||
            !string.IsNullOrEmpty(plot.ShowNegativeBubblesValue) && plot.ShowNegativeBubblesValue is not ("0" or "false") ||
            ReadChartPointAxisDeclaration(xAxis) is not { } xSource || ReadChartPointAxisDeclaration(yAxis) is not { } ySource ||
            xSource.Minimum is null || xSource.Maximum is null || xSource.MajorUnit is null)
            return null;
        // Automatic horizontal interval density is unresolved. Keep its whole
        // geometry on the prior path rather than mixing two range owners.
        double scale = 1d;
        if (plot.BubbleScale is { } authoredScale)
        {
            if (!double.IsFinite(authoredScale) || authoredScale < 0d || authoredScale > 300d || Math.Truncate(authoredScale) != authoredScale) return null;
            scale = authoredScale / 100d;
        }
        else if (!string.IsNullOrEmpty(plot.BubbleScaleValue)) return null;

        var token = resolver.CancellationToken;
        double maximumSize = 0d;
        foreach (var item in series)
        foreach (var point in item.Points)
        {
            token.ThrowIfCancellationRequested();
            if (!(point.Size > 0d) || !double.IsFinite(point.Size)) return null;
            maximumSize = Math.Max(maximumSize, point.Size);
        }
        if (!(maximumSize > 0d) || !TryResolvePointAxisRange(dataX, xSource.Minimum, xSource.Maximum, xSource.MajorUnit,
                PptxChartMetricRules.AxisNiceTickTargetCount, out var x, out double xUnit, headroomRatio: 0d) ||
            !TryResolvePointAxisRange(dataY, ySource.Minimum, ySource.Maximum, ySource.MajorUnit,
                PptxChartMetricRules.AxisNiceTickTargetCount, out var y, out double yUnit, headroomRatio: 0d)) return null;

        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        ChartTextStyle xStyle = ReadSceneOrXmlChartTextStyle(theme, chart, xAxis.SceneAxis, xml, xAxis.XmlAxis,
            PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
        ChartTextStyle yStyle = ReadSceneOrXmlChartTextStyle(theme, chart, yAxis.SceneAxis, xml, yAxis.XmlAxis,
            PptxChartMetricRules.ValueAxisFallbackFontSize, "valueAxis");
        bool xLabels = IsSceneOrXmlChartAxisLabelVisible(xAxis.SceneAxis, xAxis.XmlAxis);
        bool yLabels = IsSceneOrXmlChartAxisLabelVisible(yAxis.SceneAxis, yAxis.XmlAxis);
        double? yLine = yLabels ? await measure.TryMeasureNormalLineHeightAsync(yStyle).ConfigureAwait(false) : 0d;
        if (yLine is null) return null;
        ChartTextStyle legendStyle = ReadSceneOrXmlChartLegendTextStyle(theme, chart.ColorMap, chart, xml);
        double legendWidth = 0d, legendLine = 0d;
        if (fallback.Legend.Visible && entries.Count != 0)
        {
            if (await MeasureSideLegendFrameAsync(measure, legendStyle, entries).ConfigureAwait(false) is not { } legend) return null;
            legendWidth = legend.Width;
            legendLine = legend.LineHeight;
        }
        ChartFrameBox frame = fallback.Frame;
        double inset = PptxChartMetricRules.ChartContentFrameInset;
        double margin = PptxChartMetricRules.ColumnPlotRightReserve;
        var outerHeader = await MeasureCartesianHeaderReserveAsync(theme, chart, resolver, 0d).ConfigureAwait(false);
        var innerHeader = await MeasureCartesianHeaderReserveAsync(theme, chart, resolver, yLine.Value).ConfigureAwait(false);
        double right = legendWidth > 0d ? margin + inset + legendWidth : margin;
        var outer = new ChartLayoutBox(frame.X + margin, frame.Y + margin,
            frame.Width - margin - right, frame.Height - outerHeader.Reserve - margin);
        if (!(outer.Width > 0d) || !(outer.Height > 0d)) return null;
        // Reserve the diameter from the nominal container before applying the
        // thirty-percent scale: R=.15*s*(N-2R). All painting uses this one radius.
        double radius = (Math.Min(outer.Width, outer.Height) + 2d * inset) * (.15d * scale) / (1d + .3d * scale);
        if (!double.IsFinite(radius)) return null;
        var sizing = new ChartBubbleSizing(maximumSize, radius);

        // Range nicing and caption admission are discrete layout dependencies.
        // Keep non-convergent declarations on the existing fallback; this is a
        // bounded layout pass, not a caller quota or a replacement work budget.
        for (int pass = 0; pass < 16; pass++)
        {
            token.ThrowIfCancellationRequested();
            var xOptions = ReadSceneOrXmlChartBubbleValueAxisOptions(xAxis.SceneAxis, xAxis.XmlAxis, theme, x,
                PptxChartMetricRules.AxisNiceTickTargetCount);
            var yOptions = ReadSceneOrXmlChartBubbleValueAxisOptions(yAxis.SceneAxis, yAxis.XmlAxis, theme, y,
                PptxChartMetricRules.AxisNiceTickTargetCount);
            xOptions = xOptions with { Units = xOptions.Units with { MajorUnit = xUnit }, TickPolicy = ChartAxisTickPolicy.SourceMinimumIntervals };
            yOptions = yOptions with { Units = yOptions.Units with { MajorUnit = yUnit }, TickPolicy = ChartAxisTickPolicy.SourceMinimumIntervals };
            ChartAxisPlacement xPlacement = ReadChartAxisPlacement(xAxis, y, yOptions.Reversed) ?? new(y.Min, 0d, false, false);
            ChartAxisPlacement yPlacement = ReadChartAxisPlacement(yAxis, x, xOptions.Reversed, horizontal: false) ?? new(x.Min, 0d, false, false);
            double firstWidth = 0d, lastWidth = 0d, sideSlot = 0d;
            if (xLabels)
            {
                firstWidth = await measure.MeasureAsync(FormatSceneOrXmlChartAxisLabel(x.Min, xAxis.SceneAxis, xAxis.XmlAxis, null), xStyle).ConfigureAwait(false);
                lastWidth = await measure.MeasureAsync(FormatSceneOrXmlChartAxisLabel(x.Max, xAxis.SceneAxis, xAxis.XmlAxis, null), xStyle).ConfigureAwait(false);
                if (xOptions.Reversed) (firstWidth, lastWidth) = (lastWidth, firstWidth);
            }
            if (yLabels)
            {
                var captions = await MeasureChartValueAxisCaptionsAsync(measure, yStyle, yAxis, y, yOptions.Units,
                    PptxChartMetricRules.AxisNiceTickTargetCount, yOptions.TickPolicy, null, reserveEndpoints: true, context: context).ConfigureAwait(false);
                sideSlot = Math.Max(0d, ComputeCartesianAxisLabelSideReserve(captions.MaximumWidth, yStyle.FontSize) - margin);
            }
            double captionReserve = xLabels ? ComputeCartesianLabelStripBottomReserve(xStyle.FontSize) : innerHeader.NumericClearance;
            var vertical = ResolveCartesianCategoryReserves(frame.Height, innerHeader.Reserve, captionReserve,
                innerHeader.NumericClearance, xPlacement);
            var desired = new ChartPlotBox(outer.X, frame.Y + vertical.Bottom, outer.Width,
                frame.Height - vertical.Top - vertical.Bottom);
            var admissionFrame = new ChartFrameBox(outer.X - inset, frame.Y, outer.Width + 2d * inset, frame.Height);
            ChartPlotBox admitted = AdmitChartPlotCaptions(admissionFrame, desired, firstWidth, lastWidth, sideSlot, yPlacement);
            if (!(admitted.Width > 0d) || !(admitted.Height > 0d)) return null;
            if (!TryResolveBubbleAxis(series, horizontal: true, xSource, dataX, x, radius / admitted.Width, maximumSize, token,
                    out var nextX, out double nextXUnit) ||
                !TryResolveBubbleAxis(series, horizontal: false, ySource, dataY, y, radius / admitted.Height, maximumSize, token,
                    out var nextY, out double nextYUnit)) return null;
            if (nextX == x && nextY == y && nextXUnit == xUnit && nextYUnit == yUnit)
            {
                ChartLayout layout = fallback with { PlotBox = admitted, PlotAreaBox = outer };
                ChartLegendBox? legendBox = fallback.Legend.Visible && entries.Count != 0
                    ? await PrepareBubbleLegendAsync(frame, outer, entries, legendStyle, legendWidth, legendLine, measure).ConfigureAwait(false) : null;
                return new(layout, x, y, xOptions, yOptions, sizing, legendBox);
            }
            (x, y, xUnit, yUnit) = (nextX, nextY, nextXUnit, nextYUnit);
        }
        return null;
    }

    private static bool TryResolveBubbleAxis(IReadOnlyList<ScatterSeries> series, bool horizontal,
        ChartPointAxisDeclaration source, ChartValueExtents data, ChartValueExtents current, double relativeRadius,
        double maximumSize, CancellationToken token, out ChartValueExtents range, out double unit)
    {
        range = default; unit = 0d;
        double margin = relativeRadius * (current.Max - current.Min);
        if (!double.IsFinite(margin)) return false;
        double low = double.PositiveInfinity, high = double.NegativeInfinity;
        foreach (var item in series)
        foreach (var point in item.Points)
        {
            token.ThrowIfCancellationRequested();
            double value = horizontal ? point.X : point.Y;
            double pointMargin = Math.Sqrt(point.Size / maximumSize) * margin;
            low = Math.Min(low, value - pointMargin); high = Math.Max(high, value + pointMargin);
        }
        var occupied = IncludePointAxisZero(new(low, high), source.Minimum, source.Maximum);
        double span = occupied.Max - occupied.Min;
        if (!(span > 0d) || !double.IsFinite(span)) return false;
        double inferred = source.MajorUnit ?? ChooseChartAxisMajorUnit(span, PptxChartMetricRules.AxisNiceTickTargetCount);
        return TryResolvePointAxisRange(new(data.Min - margin, data.Max + margin), source.Minimum, source.Maximum,
            inferred, PptxChartMetricRules.AxisNiceTickTargetCount, out range, out unit, headroomRatio: 0d);
    }

    private static async ValueTask<ChartLegendBox> PrepareBubbleLegendAsync(ChartFrameBox frame, ChartLayoutBox outer,
        IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle style, double width, double lineHeight, ChartTextMeasurer measure)
    {
        double markerSize = .45d * lineHeight;
        double rowHeight = lineHeight + HorizontalLegendRowPadding;
        double center = outer.Y + outer.Height / 2d;
        double textRight = frame.X + frame.Width - PptxChartMetricRules.AreaRightLegendTail;
        double textWidth = width - .9d * lineHeight - HorizontalLegendRowPadding;
        double textX = textRight - textWidth;
        double markerX = textX - 1.5d * markerSize;
        double clipHeight = entries.Count * rowHeight;
        double ascent = 0d, descent = 0d;
        var frames = new List<ChartLegendEntryFrame>(entries.Count);
        for (int i = 0; i < entries.Count; i++)
        {
            double rowCenter = center + ((entries.Count - 1d) / 2d - i) * rowHeight;
            var ink = await measure.TryMeasureInkBoundsAsync(entries[i].Name, style).ConfigureAwait(false);
            if (ink is { } extent)
            {
                ascent = Math.Max(ascent, extent.Top);
                descent = Math.Max(descent, -extent.Bottom);
            }
            else
            {
                var line = await measure.MeasureLineBoxAsync(style).ConfigureAwait(false);
                ascent = Math.Max(ascent, line.Ascent);
                descent = Math.Max(descent, line.Descent);
            }
            double baseline = rowCenter - (ink is { } bounds ? (bounds.Bottom + bounds.Top) / 2d : style.FontSize * .3d);
            frames.Add(new(markerX, baseline, markerSize * 1.5d + textWidth, rowCenter - markerSize / 2d));
        }
        return new(markerX, center - clipHeight / 2d, textRight - markerX + PptxChartMetricRules.ChartContentFrameInset,
            clipHeight, frames[0].Y, rowHeight, markerSize, markerSize, markerSize / 2d, false, false, frames, ascent, descent);
    }
}
