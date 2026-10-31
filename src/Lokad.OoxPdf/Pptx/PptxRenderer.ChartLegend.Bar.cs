using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Office controls across Aptos/Calibri/Arial/Times, sizes and wrapped rows.
    // Row advance, occupied height and frame inset are separate measurements.
    private const double HorizontalLegendLineBoxScale = 0.99d;
    private const double HorizontalLegendRowPadding = 6d;
    private const double HorizontalLegendFrameInset = 6d;
    private const double HorizontalLegendWrappedHeightTrim = 1.5d;

    private static bool UsesAutomaticHorizontalBarFrame(ChartBarPlotOptions options, PptxSceneChart chart, ChartLegendLayout legend) =>
        options.BarDirection == PptxSceneChartBarDirection.Bar &&
        !chart.PlotAreaLayout.HasLayout && !legend.Layout.HasLayout &&
        (!legend.Visible || legend.Overlay ||
            legend.PositionKind is PptxSceneChartLegendPosition.Top or PptxSceneChartLegendPosition.Bottom);

    private static bool UsesMeasuredHorizontalBarLegend(ChartBarPlotOptions options, PptxSceneChart chart, ChartLegendLayout legend) =>
        UsesAutomaticHorizontalBarFrame(options, chart, legend) && legend.Visible && !legend.Overlay;

    private static bool IsLegendBaselineClipped(ChartLegendBox box, double baseline) =>
        baseline - box.ContentDescent < box.ClipY - 0.001d ||
        baseline + (box.ContentAscent ?? box.LineHeight) > box.ClipY + box.ClipHeight + 0.001d;

    private static IReadOnlyList<ChartLegendEntry> BuildBarCategoryLegendEntries(
        PptxTheme theme, PptxColorMap colorMap, ChartColorPalette palette,
        PptxSceneChartPlot? plot, XElement chart, ChartIndexedNumberVector vector,
        ChartBarPlotOptions options, IReadOnlyList<ChartSeriesFill?> seriesFills,
        IReadOnlyList<ChartSeriesStroke?> seriesStrokes,
        IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesFill>> pointFills,
        IReadOnlyList<IReadOnlyDictionary<int, ChartSeriesStroke>> pointStrokes,
        ChartWorkbookData? workbook, bool plotVisibleOnly)
    {
        int count = vector.DensePointCount();
        Dictionary<int, ChartIndexedNumberPoint> points = BuildDensePointIndex(vector);
        double Value(int index) => points.TryGetValue(index, out var point) ? point.Value ?? 0d : 0d;
        ChartSeriesFill Fill(int index) => ResolveBarPointFill(theme, colorMap, palette, 0, index, 1,
            options.VaryColors.Value, seriesFills, pointFills, count, Value(index), options.InvertsNegativePoint(0, index));
        ChartSeriesStroke? Stroke(int index) => pointStrokes.Count > 0 && pointStrokes[0].TryGetValue(index, out var stroke)
            ? stroke : seriesStrokes.Count > 0 && seriesStrokes[0] is { } seriesStroke
                ? seriesStroke : ResolveNegativeBarFallbackStroke(pointStrokes, 0, index, Value(index), options.InvertsNegativePoint(0, index));
        return BuildCategoryFillLegendEntries(theme, colorMap, palette, plot, chart, pointFills[0],
            workbook, plotVisibleOnly, count, Fill, Stroke);
    }

    // Measure once, reserve automatic plot space, then give painting the same rows.
    // Manual frames and overlay legends retain their authored placement contract.
    private static async ValueTask<(ChartLayout Layout, ChartLegendBox Box)> LayoutBarLegendAsync(
        ChartLayout layout, IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle style,
        PresentationFontResolver resolver, PptxRenderContext context, bool measuredBand)
    {
        if (!layout.Legend.Visible || entries.Count == 0) { return (layout, default); }
        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        await measure.PrepareAsync(entries.Select(entry => entry.Name), style).ConfigureAwait(false);
        ChartLegendBox box = ResolveChartLegendBox(layout.Frame, layout.PlotBox, entries,
            layout.Legend, style, measure, ChartLegendPlacement.Default);
        if (!box.Horizontal) { return (layout, box); }
        bool bottom = layout.Legend.PositionKind == PptxSceneChartLegendPosition.Bottom;
        bool packedAsBottom = bottom || measuredBand;
        double gap = ComputeHorizontalLegendTextGap(style.FontSize, packedAsBottom);
        double padding = ResolveHorizontalLegendInterEntryGap(style.FontSize, entries, style, measure, packedAsBottom);
        var lineBox = measuredBand ? await measure.MeasureLineBoxAsync(style).ConfigureAwait(false) : default;
        if (measuredBand)
            box = box with { LineHeight = lineBox.Height * HorizontalLegendLineBoxScale + HorizontalLegendRowPadding,
                TextGap = gap, ContentAscent = Math.Max(lineBox.Ascent, box.MarkerSize), ContentDescent = lineBox.Descent };
        // Horizontal legends belong to the chart frame, including the axis-label
        // reserve beside the plot. A manual legend keeps its declared width.
        double availableWidth = layout.Legend.Layout.HasLayout ? box.Width : layout.Frame.Width;
        double[] widths = new double[entries.Count];
        var rows = new List<(int Start, int End, double Width)>();
        int start = 0;
        double width = 0d;
        for (int i = 0; i < entries.Count; i++)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            widths[i] = GetPackedHorizontalLegendEntryWidth(entries[i].Name, style, measure, box.MarkerSize, gap, padding);
            if (i > start && width + widths[i] - padding > availableWidth)
            {
                rows.Add((start, i, Math.Max(0d, width - padding)));
                start = i;
                width = 0d;
            }
            width += widths[i];
        }
        rows.Add((start, entries.Count, Math.Max(0d, width - padding)));
        if (measuredBand && rows.Count > 1 && widths.Max() <= availableWidth + padding)
        {
            // Greedy packing establishes the required row count. Minimize its
            // widest row so wrapped legends stay balanced without reordering.
            int rowCount = rows.Count;
            int CountRows(double limit)
            {
                int count = 1;
                double used = 0d;
                foreach (double entryWidth in widths)
                {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    if (used > 0d && used + entryWidth > limit) { count++; used = 0d; }
                    used += entryWidth;
                }
                return count;
            }
            double lower = widths.Max();
            double upper = availableWidth + padding;
            for (int step = 0; step < 32 && upper - lower > 0.01d; step++)
            {
                double middle = (lower + upper) / 2d;
                if (CountRows(middle) <= rowCount) { upper = middle; } else { lower = middle; }
            }
            rows.Clear();
            start = 0;
            width = 0d;
            for (int i = 0; i < entries.Count; i++)
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                if (i > start && width + widths[i] > upper)
                {
                    rows.Add((start, i, Math.Max(0d, width - padding)));
                    start = i;
                    width = 0d;
                }
                width += widths[i];
            }
            rows.Add((start, entries.Count, Math.Max(0d, width - padding)));
        }
        if (!layout.Legend.Layout.HasLayout)
        {
            double packedWidth = Math.Min(availableWidth, rows.Max(row => row.Width));
            double x = packedAsBottom
                ? ResolveBottomLegendFrameAnchorX(layout.Frame.X, layout.Frame.Width, packedWidth, box.MarkerSize)
                : layout.Frame.X + (layout.Frame.Width - packedWidth) / 2d;
            box = box with { X = x, Width = packedWidth };
        }
        double firstY, clipY, clipHeight;
        if (measuredBand)
        {
            double occupied = rows.Count * box.LineHeight - (rows.Count > 1 ? HorizontalLegendWrappedHeightTrim : 0d);
            double requiredBand = occupied + HorizontalLegendFrameInset;
            double reserved = Math.Min(requiredBand, Math.Max(0d, layout.PlotBox.Height - 1d));
            double topContent = Math.Max(0d, layout.Frame.Y + layout.Frame.Height -
                PptxChartMetricRules.HorizontalBarPlotTopReserve - layout.PlotBox.Y - layout.PlotBox.Height);
            double legendTop = bottom
                ? Math.Min(layout.Frame.Y + HorizontalLegendFrameInset + occupied,
                    layout.Frame.Y + layout.Frame.Height - HorizontalLegendFrameInset - topContent)
                : layout.Frame.Y + layout.Frame.Height - HorizontalLegendFrameInset - topContent;
            clipY = bottom ? layout.Frame.Y + HorizontalLegendFrameInset :
                Math.Max(layout.Frame.Y + HorizontalLegendFrameInset, legendTop - occupied);
            clipHeight = Math.Max(0d, legendTop - clipY);
            double baselineInset = PptxChartMetricRules.LegendBottomFramePad +
                style.FontSize * PptxChartMetricRules.LegendBottomBaselineFontFactor - HorizontalLegendFrameInset;
            firstY = legendTop - box.LineHeight + baselineInset;
            double shift = bottom ? reserved : 0d;
            layout = layout with
            {
                PlotBox = layout.PlotBox with { Y = layout.PlotBox.Y + shift, Height = layout.PlotBox.Height - reserved },
                PlotAreaBox = layout.PlotAreaBox with { Y = layout.PlotAreaBox.Y + shift, Height = Math.Max(1d, layout.PlotAreaBox.Height - reserved) }
            };
        }
        else
        {
            double extra = (rows.Count - 1) * box.LineHeight;
            bool automatic = !layout.Legend.Overlay && !layout.Legend.Layout.HasLayout && !layout.ManualPlotLayoutApplied;
            double reserved = automatic ? Math.Min(extra, Math.Max(0d, layout.PlotBox.Height - 1d)) : 0d;
            if (reserved > 0d)
            {
                double shift = bottom ? reserved : 0d;
                layout = layout with
                {
                    PlotBox = layout.PlotBox with { Y = layout.PlotBox.Y + shift, Height = layout.PlotBox.Height - reserved },
                    PlotAreaBox = layout.PlotAreaBox with { Y = layout.PlotAreaBox.Y + shift, Height = Math.Max(1d, layout.PlotAreaBox.Height - reserved) }
                };
            }
            double admitted = automatic ? reserved : extra;
            firstY = box.FirstY + (bottom ? admitted : 0d);
            clipY = box.ClipY - (bottom ? 0d : admitted);
            clipHeight = box.ClipHeight + admitted;
            if (layout.Legend.Layout.HasLayout)
            {
                clipY = box.ClipY;
                clipHeight = box.ClipHeight;
            }
        }
        double clipTop = Math.Min(layout.Frame.Y + layout.Frame.Height, clipY + clipHeight);
        clipY = Math.Max(layout.Frame.Y, clipY);
        clipHeight = Math.Max(0d, clipTop - clipY);
        box = box with { ClipY = clipY, ClipHeight = clipHeight };
        var frames = new ChartLegendEntryFrame[entries.Count];
        bool clipped = false;
        for (int row = 0; row < rows.Count; row++)
        {
            double x = box.X + Math.Max(0d, (box.Width - rows[row].Width) / 2d);
            double y = firstY - row * box.LineHeight;
            for (int i = rows[row].Start; i < rows[row].End; i++)
            {
                frames[i] = new ChartLegendEntryFrame(x, y, widths[i]);
                clipped |= IsLegendBaselineClipped(box, y) || widths[i] - padding > box.Width;
                x += widths[i];
            }
        }
        if (clipped)
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning,
                "Legend exceeds its available frame; overflowing entries are clipped.", context.SlidePartName, context.SlideNumber, "Approximate");
        }
        return (layout, box with { FirstY = firstY, EntryFrames = frames });
    }
}
