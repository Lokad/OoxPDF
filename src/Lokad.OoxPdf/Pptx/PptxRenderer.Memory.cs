namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    // Source-owned estimates keep painting/snapshots in one page-local grant;
    // retained descriptors transfer to their separate staging owner.
    private static IMemoryReservation? ReservePageConstruction(PptxSceneSlide slide, CancellationToken token)
    {
        IMemoryScope? owner = OoxConversionBudget.Current?.Memory;
        if (owner is null) return null;
        long capacity = checked(8 * OoxMemory.EstimateMarkupCapacity(slide.SlideXml, token));
        if (slide.LayoutXml is { } layout) capacity = checked(capacity + 8 * OoxMemory.EstimateMarkupCapacity(layout, token));
        if (slide.MasterXml is { } master) capacity = checked(capacity + 8 * OoxMemory.EstimateMarkupCapacity(master, token));
        if (slide.ShowMasterShapes) Resources(slide.MasterNodes);
        Resources(slide.LayoutNodes); Resources(slide.SlideNodes);
        token.ThrowIfCancellationRequested();
        return OoxMemory.Reserve(owner, capacity);

        void Resources(IReadOnlyList<PptxSceneNode> nodes)
        {
            foreach (var node in nodes)
            {
                token.ThrowIfCancellationRequested();
                if (node.Chart?.ChartXml is { } chart) capacity = checked(capacity + 8 * OoxMemory.EstimateMarkupCapacity(chart, token));
                Image(node.Picture?.Resource);
                Image(node.Shape?.PictureFill.Resource);
                Resources(node.Children);
            }
        }
        void Image(PptxSceneImageResource? image)
        {
            if (image is not null && image.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
                capacity = checked(capacity + 64L * image.Bytes.Length);
        }
    }
}
