namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Source-owned estimates keep painting/snapshots in one page-local grant;
    // retained descriptors transfer to their separate staging owner.
    private static IMemoryReservation? ReservePageConstruction(PptxSceneSlide slide, CancellationToken token, out long axisConstructionSites, out long svgSourceBytes)
    {
        axisConstructionSites = 0;
        svgSourceBytes = 0;
        IMemoryScope? owner = OoxConversionBudget.Current?.Memory;
        if (owner is null) return null;
        // Background painting is one solid rectangle. Unused master/layout style
        // definitions do not increase construction; select nodes as painting does.
        long capacity = 64L * 1024;
        long chartExpansion = 0;
        long axisSites = 0;
        long svgBytes = 0;
        if (slide.ShowMasterShapes) Nodes(slide.MasterNodes, renderPlaceholders: false);
        Nodes(slide.LayoutNodes, renderPlaceholders: false);
        Nodes(slide.SlideNodes, renderPlaceholders: true);
        token.ThrowIfCancellationRequested();
        axisConstructionSites = Math.Max(1L, axisSites);
        svgSourceBytes = svgBytes;
        // Chart frames paint sequentially and clear their temporary memos on exit.
        // Share one expansion window instead of adding each frame's dense peak.
        return OoxMemory.Reserve(owner, checked(capacity + chartExpansion));

        void Nodes(IReadOnlyList<PptxSceneNode> nodes, bool renderPlaceholders)
        {
            foreach (var node in EnumerateRenderedSceneNodes(nodes, renderPlaceholders))
            {
                token.ThrowIfCancellationRequested();
                capacity = checked(capacity + 4096);
                if (node.Kind == PptxSceneNodeKind.Group) continue; // Its children are yielded once by the shared traversal.
                capacity = checked(capacity + 8 * OoxMemory.EstimateMarkupCapacity(node.Source, token));
                if (node.Chart is { ChartXml: { } chart } sceneChart)
                {
                    long markup = checked(8 * OoxMemory.EstimateMarkupCapacity(chart, token));
                    capacity = checked(capacity + markup);
                    chartExpansion = Math.Max(chartExpansion, Math.Max(0L, EstimateChartSlotConstruction(sceneChart, token) - markup));
                    // Numeric axis lists, labels and their painted paths can
                    // coexist across frames. Repeated measurement/painting
                    // shares one high-water window, never a work counter.
                    axisSites = checked(axisSites + Math.Max(1L, sceneChart.Axes.Count) * Math.Max(1, sceneChart.Plots.Count));
                    foreach (var plot in sceneChart.Plots)
                    {
                        if (plot.PlotKind != PptxSceneChartPlotKind.Radar) continue;
                        int vertices = 0;
                        foreach (var series in plot.Series)
                        {
                            token.ThrowIfCancellationRequested();
                            vertices = Math.Max(vertices, CountChartSlots(series.CategoryPointCount, series.Categories.Count, series.CategoryPoints, static p => p.Index, token));
                            vertices = Math.Max(vertices, CountChartSlots(series.ValuePointCount, series.Values.Count, series.ValuePoints, static p => p.Index, token));
                        }
                        axisSites = checked(axisSites + vertices);
                    }
                }
                // A placeholder may inherit text absent from its own source XML.
                if (node.TextBody is { } text)
                    foreach (var paragraph in text.Paragraphs)
                        foreach (var run in paragraph.Runs)
                        {
                            token.ThrowIfCancellationRequested();
                            capacity = checked(capacity + 512L * run.Text.Length);
                        }
                Image(node.Picture?.Resource);
                Image(node.Shape?.PictureFill.Resource);
            }
        }
        void Image(PptxSceneImageResource? image)
        {
            if (image is not null && image.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
            {
                capacity = checked(capacity + 64L * image.Bytes.Length);
                svgBytes = checked(svgBytes + image.Bytes.Length);
            }
        }
    }

    private static long EstimateChartSlotConstruction(PptxSceneChart chart, CancellationToken token)
    {
        long capacity = 0;
        foreach (var plot in chart.Plots)
        {
            token.ThrowIfCancellationRequested();
            bool numericSlots = plot.PlotKind is PptxSceneChartPlotKind.Bar or PptxSceneChartPlotKind.Line or PptxSceneChartPlotKind.Area or PptxSceneChartPlotKind.Radar;
            bool categorySlots = numericSlots || plot.PlotKind is PptxSceneChartPlotKind.Pie or PptxSceneChartPlotKind.Doughnut;
            if (!categorySlots) continue;
            bool renderable = false, categorySource = false;
            long values = 0;
            int categories = 0;
            foreach (var series in plot.Series)
            {
                token.ThrowIfCancellationRequested();
                if (!categorySource && (series.CategoryPointCount is not null || series.CategoryPoints.Count != 0 || series.Categories.Count != 0 || series.CategoryLevels.Any(level => level.Count != 0)))
                {
                    categorySource = true;
                    categories = CountChartSlots(series.CategoryPointCount, series.Categories.Count, series.CategoryPoints, static p => p.Index, token);
                    foreach (var level in series.CategoryLevels)
                        categories = Math.Max(categories, CountChartSlots(null, 0, level, static p => p.Index, token));
                }
                if (series.Values.Count == 0 && !series.ValuePoints.Any(point => point.Value is not null)) continue;
                renderable = true;
                if (numericSlots)
                    values = checked(values + CountChartSlots(series.ValuePointCount, series.Values.Count, series.ValuePoints, static p => p.Index, token));
            }
            // Coarse construction headroom for nullable values, coordinate/stack
            // buffers and reference category slots. Source markup already funds
            // the real point/provenance records; holes must not duplicate them.
            if (renderable) capacity = checked(capacity + 128L * values + 32L * categories);
        }
        return capacity;
    }

    private static int CountChartSlots<T>(int? declared, int ordinalCount, IReadOnlyList<T> points, Func<T, int> index, CancellationToken token)
    {
        int count = Math.Clamp(Math.Max(declared ?? 0, ordinalCount), 0, MaxChartDensePoints);
        foreach (T point in points)
        {
            token.ThrowIfCancellationRequested();
            int slot = index(point);
            if (slot >= 0 && slot < MaxChartDensePoints) count = Math.Max(count, slot + 1);
        }
        // This estimate does not validate or relax the existing dense bound.
        return count;
    }

    private static void AdmitChartAxisConstruction(PptxRenderContext? context, int ticks)
    {
        if (context?.PageConstruction is not { } construction) return;
        context.CancellationToken.ThrowIfCancellationRequested();
        // Coarse family headroom: numeric/filtered arrays, caption runs and
        // generated path/text operators plus their construction snapshots.
        // This is admission to significant storage, not a managed heap ceiling.
        long capacity = checked(context.PageConstructionBaseBytes + 512L * ticks * context.PageAxisConstructionSites);
        if (capacity > construction.Bytes) OoxMemory.Resize(construction, capacity);
        context.CancellationToken.ThrowIfCancellationRequested();
    }
}
