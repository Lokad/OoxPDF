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
        if (slide.ShowMasterShapes) Nodes(slide.MasterNodes, renderPlaceholders: false);
        Nodes(slide.LayoutNodes, renderPlaceholders: false);
        Nodes(slide.SlideNodes, renderPlaceholders: true);
        token.ThrowIfCancellationRequested();
        return OoxMemory.Reserve(owner, capacity);

        void Nodes(IReadOnlyList<PptxSceneNode> nodes, bool renderPlaceholders)
        {
            foreach (var node in EnumerateRenderedSceneNodes(nodes, renderPlaceholders))
            {
                token.ThrowIfCancellationRequested();
                capacity = checked(capacity + 4096);
                if (node.Kind == PptxSceneNodeKind.Group) continue; // Its children are yielded once by the shared traversal.
                capacity = checked(capacity + 8 * OoxMemory.EstimateMarkupCapacity(node.Source, token));
                if (node.Chart?.ChartXml is { } chart) capacity = checked(capacity + 8 * OoxMemory.EstimateMarkupCapacity(chart, token));
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
}
