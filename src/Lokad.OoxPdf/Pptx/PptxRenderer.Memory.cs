namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Source-owned estimates keep painting/snapshots in one page-local grant;
    // retained descriptors transfer to their separate staging owner.
    private static IMemoryReservation? ReservePageConstruction(PptxSceneSlide slide, CancellationToken token)
    {
        IMemoryScope? owner = OoxConversionBudget.Current?.Memory;
        if (owner is null) return null;
        // Background painting is one solid rectangle. Unused master/layout style
        // definitions do not increase construction; select nodes as painting does.
        long capacity = 64L * 1024;
        long chartExpansion = 0;
        if (slide.ShowMasterShapes) Nodes(slide.MasterNodes, renderPlaceholders: false);
        Nodes(slide.LayoutNodes, renderPlaceholders: false);
        Nodes(slide.SlideNodes, renderPlaceholders: true);
        token.ThrowIfCancellationRequested();
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
                capacity = checked(capacity + 64L * image.Bytes.Length);
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
                    categories = Slots(series.CategoryPointCount, series.Categories.Count, series.CategoryPoints, static p => p.Index);
                    foreach (var level in series.CategoryLevels)
                        categories = Math.Max(categories, Slots(null, 0, level, static p => p.Index));
                }
                if (series.Values.Count == 0 && !series.ValuePoints.Any(point => point.Value is not null)) continue;
                renderable = true;
                if (numericSlots)
                    values = checked(values + Slots(series.ValuePointCount, series.Values.Count, series.ValuePoints, static p => p.Index));
            }
            // Coarse construction headroom for nullable values, coordinate/stack
            // buffers and reference category slots. Source markup already funds
            // the real point/provenance records; holes must not duplicate them.
            if (renderable) capacity = checked(capacity + 128L * values + 32L * categories);
        }
        return capacity;

        int Slots<T>(int? declared, int ordinalCount, IReadOnlyList<T> points, Func<T, int> index)
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
    }
}
