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

    private static bool UsesAutomaticBarFrame(ChartBarPlotOptions options, PptxSceneChart chart, ChartLegendLayout legend)
    {
        // Column bottom legends share the measured row owner. Composed plots,
        // axis-title furniture and other legend sides retain their source layout.
        if (options.BarDirection != PptxSceneChartBarDirection.Bar)
        {
            ChartAxisTitleReserveSides titles = ReadDefaultAxisTitleReserveSides(chart);
            if (chart.Plots.Count != 1 || legend.PositionKind != PptxSceneChartLegendPosition.Bottom ||
                titles.HasHorizontalTitle || titles.HasVerticalTitle)
                return false;
        }
        return !chart.PlotAreaLayout.HasGeometryValues && !legend.Layout.HasGeometryValues &&
            (!legend.Visible || legend.Overlay ||
                legend.PositionKind is PptxSceneChartLegendPosition.Top or PptxSceneChartLegendPosition.Bottom);
    }

    private static bool UsesMeasuredBarLegend(ChartBarPlotOptions options, PptxSceneChart chart, ChartLegendLayout legend) =>
        UsesAutomaticBarFrame(options, chart, legend) && legend.Visible;

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
    // Authored data frames and overlay legends do not reserve plot space.
    private static async ValueTask<(ChartLayout Layout, ChartLegendBox Box)> LayoutChartLegendRowsAsync(
        ChartLayout layout, IReadOnlyList<ChartLegendEntry> entries, ChartTextStyle style,
        PresentationFontResolver resolver, PptxRenderContext context, double maximumRowWidth, bool measuredBand)
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
        // The container supplies its row constraint; bar frames include the
        // axis strips, while line rows follow their plot width. Manual legends
        // keep their authored width. Painting consumes these same packed frames.
        double availableWidth = layout.Legend.Layout.HasGeometryValues ? box.Width : Math.Min(layout.Frame.Width, maximumRowWidth);
        double[] widths = new double[entries.Count];
        var rows = new List<(int Start, int End, double Width)>();
        bool automaticLegend = !layout.Legend.Layout.HasGeometryValues &&
            (measuredBand || !layout.ManualPlotLayoutApplied);
        // Filled keys share natural-row admission and wrapped columns. Series
        // cells retain their measured key/text gap; category cells also own
        // the independently observed marker-sized tail.
        bool measuredFillGrid = measuredBand && bottom && automaticLegend &&
            entries.All(entry => entry.Fill is not null);
        bool categoryFillGrid = measuredFillGrid && entries.All(entry => entry.SeriesName is null);
        double naturalWidth = 0d, widestCell = 0d, widestNaturalCell = 0d;
        for (int i = 0; i < entries.Count; i++)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            widths[i] = GetPackedHorizontalLegendEntryWidth(entries[i].Name, style, measure, box.MarkerWidth, gap, padding);
            if (measuredFillGrid)
            {
                double advance = measure.Measure(entries[i].Name, style) + box.MarkerWidth;
                double naturalCell = advance + box.MarkerSize;
                double cell = categoryFillGrid ? naturalCell : widths[i];
                naturalWidth += naturalCell;
                widestNaturalCell = Math.Max(widestNaturalCell, naturalCell);
                widestCell = Math.Max(widestCell, cell);
            }
        }
        if (measuredFillGrid)
        {
            // Office's automatic fill legend first admits a natural single row.
            // Wrapped rows use equal cells, including the incomplete last row;
            // alternating short/long names still occupy the widest cell. The
            // independently observed frame limit also owns one-row admission.
            availableWidth = Math.Min(availableWidth, layout.Frame.Width * 0.9d);
            if (naturalWidth <= availableWidth)
            {
                double packedWidth = widths.Sum() - padding;
                if (entries.Count > 1 && packedWidth > availableWidth)
                {
                    // Keep preferred spacing while it fits. When it does not,
                    // share the remaining whitespace between the measured keys
                    // and text, so a naturally fitting row cannot lose names.
                    double reduction = (packedWidth - availableWidth) / (entries.Count - 1);
                    padding -= reduction;
                    for (int i = 0; i < widths.Length; i++)
                    {
                        context.CancellationToken.ThrowIfCancellationRequested();
                        widths[i] -= reduction;
                    }
                    packedWidth = availableWidth;
                }
                rows.Add((0, entries.Count, Math.Max(0d, packedWidth)));
            }
            else
            {
                int capacity = (int)Math.Min(entries.Count, Math.Max(1d, Math.Floor(availableWidth / widestNaturalCell)));
                int rowCount = (entries.Count - 1) / capacity + 1;
                int columns = (entries.Count - 1) / rowCount + 1;
                Array.Fill(widths, widestCell);
                if (categoryFillGrid) padding = 0d;
                for (int start = 0; start < entries.Count; start += columns)
                {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    rows.Add((start, Math.Min(entries.Count, start + columns), columns * widestCell - padding));
                }
            }
        }
        else
        {
            int start = 0;
            double width = 0d;
            for (int i = 0; i < entries.Count; i++)
            {
                context.CancellationToken.ThrowIfCancellationRequested();
                if (i > start && width + widths[i] - padding > availableWidth)
                {
                    rows.Add((start, i, Math.Max(0d, width - padding)));
                    start = i;
                    width = 0d;
                }
                width += widths[i];
            }
            rows.Add((start, entries.Count, Math.Max(0d, width - padding)));
        }
        if (!measuredFillGrid && measuredBand && rows.Count > 1 && widths.Max() <= availableWidth + padding)
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
            int start = 0;
            double width = 0d;
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
        // Office limits automatic legends to whole rows within half the chart
        // height, including overlay legends. Authored geometry keeps its bounds.
        int admittedRowCount = automaticLegend
            ? (int)Math.Min(rows.Count, Math.Max(0d, Math.Floor(layout.Frame.Height / (2d * box.LineHeight))))
            : rows.Count;
        // Truncated Office legends allocate one further row at the tail; only
        // the admitted rows paint. Complete legends keep their existing extent.
        int placementRowCount = admittedRowCount + (admittedRowCount > 0 && admittedRowCount < rows.Count ? 1 : 0);
        if (!layout.Legend.Layout.HasGeometryValues)
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
            double occupied = placementRowCount * box.LineHeight - (placementRowCount > 1 ? HorizontalLegendWrappedHeightTrim : 0d);
            double requiredBand = admittedRowCount == 0 ? 0d : occupied + HorizontalLegendFrameInset;
            double reserved = layout.Legend.Overlay || layout.ManualPlotLayoutApplied
                ? 0d : Math.Min(requiredBand, Math.Max(0d, layout.PlotBox.Height - 1d));
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
            double extra = Math.Max(0, placementRowCount - 1) * box.LineHeight;
            bool automatic = !layout.Legend.Overlay && !layout.Legend.Layout.HasGeometryValues && !layout.ManualPlotLayoutApplied;
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
            if (layout.Legend.Layout.HasGeometryValues)
            {
                clipY = box.ClipY;
                clipHeight = box.ClipHeight;
            }
        }
        double clipTop = Math.Min(layout.Frame.Y + layout.Frame.Height, clipY + clipHeight);
        clipY = Math.Max(layout.Frame.Y, clipY);
        clipHeight = Math.Max(0d, clipTop - clipY);
        box = box with { ClipY = clipY, ClipHeight = clipHeight };
        var frames = new ChartLegendEntryFrame[admittedRowCount == 0 ? 0 : rows[admittedRowCount - 1].End];
        bool clipped = admittedRowCount < rows.Count;
        for (int row = 0; row < admittedRowCount; row++)
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
