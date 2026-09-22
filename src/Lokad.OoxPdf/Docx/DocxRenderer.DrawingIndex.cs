namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    // R12: page drawing index built once per render pass. Both drawing lists were
    // filtered for every page (and layer), and each page re-sorted its matches by
    // z-order: O(pages x drawings) filtering plus repeated sorts. The index
    // partitions each list by anchor page a single time, preserving document order
    // within a page, with one stable z-order sort per (page, layer). Lookups return
    // the exact legacy sequences (null anchors match no page; missing pages yield
    // shared empty sets), so rendering is byte-identical. The layout is final at
    // render time, so no repagination invalidation applies here.
    internal sealed class FloatingDrawingPageIndex
    {
        private static readonly PageDrawings EmptyPage = new([], [], []);

        private readonly Dictionary<int, PageDrawings> pages;

        private FloatingDrawingPageIndex(Dictionary<int, PageDrawings> pages)
        {
            this.pages = pages;
        }

        internal sealed record PageDrawings(
            DocxFloatingDrawingLayout[] All,
            DocxFloatingDrawingLayout[] Behind,
            DocxFloatingDrawingLayout[] Ahead);

        // Both render lists share one build per pass: the render path consumes the
        // per-list layer slices while balloon/text paths consume the concatenated
        // page sequences below.
        internal sealed record PageIndexPair(FloatingDrawingPageIndex Floating, FloatingDrawingPageIndex Static)
        {
            public DocxFloatingDrawingLayout[] PageAll(int pageIndex)
            {
                return Floating.Get(pageIndex).All.Concat(Static.Get(pageIndex).All).ToArray();
            }
        }

        public static PageIndexPair BuildPair(DocxLayout layout, CancellationToken cancellationToken)
        {
            return new PageIndexPair(
                Build(layout.FloatingDrawings, cancellationToken),
                Build(layout.StaticFloatingDrawings, cancellationToken));
        }

        public static FloatingDrawingPageIndex Build(IReadOnlyList<DocxFloatingDrawingLayout> drawings, CancellationToken cancellationToken)
        {
            var byPage = new Dictionary<int, List<DocxFloatingDrawingLayout>>();
            foreach (DocxFloatingDrawingLayout drawing in drawings)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (drawing.AnchorPageIndex is not int pageIndex)
                {
                    continue;
                }

                if (!byPage.TryGetValue(pageIndex, out List<DocxFloatingDrawingLayout>? list))
                {
                    list = new List<DocxFloatingDrawingLayout>();
                    byPage[pageIndex] = list;
                }

                list.Add(drawing);
            }

            var pages = new Dictionary<int, PageDrawings>(byPage.Count);
            foreach ((int pageIndex, List<DocxFloatingDrawingLayout> list) in byPage)
            {
                cancellationToken.ThrowIfCancellationRequested();
                pages[pageIndex] = new PageDrawings(
                    list.ToArray(),
                    list.Where(drawing => IsBehindDocument(drawing.Drawing)).OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue)).ToArray(),
                    list.Where(drawing => !IsBehindDocument(drawing.Drawing)).OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue)).ToArray());
            }

            return new FloatingDrawingPageIndex(pages);
        }

        public PageDrawings Get(int pageIndex)
        {
            return pages.TryGetValue(pageIndex, out PageDrawings? drawings) ? drawings : EmptyPage;
        }
    }
}
