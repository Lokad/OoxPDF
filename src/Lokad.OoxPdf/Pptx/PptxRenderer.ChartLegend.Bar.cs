using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static IReadOnlyList<ChartLegendEntry> BuildBarCategoryLegendEntries(
        PptxTheme theme, PptxColorMap colorMap, IReadOnlyList<RgbColor>? palette,
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
    private static async ValueTask<(ChartLayout Layout, ChartLegendBox Box)> LayoutBarCategoryLegendAsync(
        ChartLayout layout, IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle style,
        PresentationFontResolver resolver, PptxRenderContext context)
    {
        if (!layout.Legend.Visible || entries.Count == 0) { return (layout, default); }
        var measure = new ChartTextMeasurer(resolver, kerningEnabled: false);
        await measure.PrepareAsync(entries.Select(entry => entry.Name), style).ConfigureAwait(false);
        ChartLegendBox box = ResolveChartLegendBox(layout.Frame, layout.PlotBox, entries,
            layout.Legend, style, measure, ChartLegendPlacement.Default);
        if (!box.Horizontal) { return (layout, box); }
        bool bottom = layout.Legend.PositionKind == PptxSceneChartLegendPosition.Bottom;
        double gap = ComputeHorizontalLegendTextGap(style.FontSize, bottom);
        double padding = ResolveHorizontalLegendInterEntryGap(style.FontSize, entries, style, measure, bottom);
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
        if (!layout.Legend.Layout.HasLayout)
        {
            double packedWidth = Math.Min(availableWidth, rows.Max(row => row.Width));
            double x = bottom
                ? ResolveBottomLegendFrameAnchorX(layout.Frame.X, layout.Frame.Width, packedWidth, box.MarkerSize)
                : layout.Frame.X + (layout.Frame.Width - packedWidth) / 2d;
            box = box with { X = x, Width = packedWidth };
        }
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
        double firstY = box.FirstY + (bottom ? admitted : 0d);
        double clipY = box.ClipY - (bottom ? 0d : admitted);
        double clipHeight = box.ClipHeight + admitted;
        if (layout.Legend.Layout.HasLayout)
        {
            clipY = box.ClipY;
            clipHeight = box.ClipHeight;
        }
        double clipTop = Math.Min(layout.Frame.Y + layout.Frame.Height, clipY + clipHeight);
        clipY = Math.Max(layout.Frame.Y, clipY);
        clipHeight = Math.Max(0d, clipTop - clipY);
        var frames = new ChartLegendEntryFrame[entries.Count];
        bool clipped = false;
        for (int row = 0; row < rows.Count; row++)
        {
            double x = box.X + Math.Max(0d, (box.Width - rows[row].Width) / 2d);
            double y = firstY - row * box.LineHeight;
            for (int i = rows[row].Start; i < rows[row].End; i++)
            {
                frames[i] = new ChartLegendEntryFrame(x, y, widths[i]);
                clipped |= y < clipY - 0.001d || y + box.LineHeight > clipY + clipHeight + 0.001d || widths[i] - padding > box.Width;
                x += widths[i];
            }
        }
        if (clipped)
        {
            EmitChartDiagnostic(context.DiagnosticSink, "PPTX_UNSUPPORTED_CHART", OoxPdfSeverity.Warning,
                "Category legend exceeds its available frame; overflowing entries are clipped.", context.SlidePartName, context.SlideNumber, "Approximate");
        }
        return (layout, box with { FirstY = firstY, ClipY = clipY, ClipHeight = clipHeight, EntryFrames = frames });
    }
}
