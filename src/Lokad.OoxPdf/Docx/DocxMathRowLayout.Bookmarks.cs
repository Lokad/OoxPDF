namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathRowLayout
{
    // Source boundaries survive flattening; they are not mathematical operands.
    internal static IEnumerable<(DocxMathNode? Node, DocxMathBookmark? Bookmark)> EnumerateItems(
        DocxMathRow source, CancellationToken token, bool allowArgumentSize = false)
    {
        var pending = new Stack<(DocxMathNode? Node, DocxMathBookmark? Bookmark)>();
        pending.Push((source, null));
        while (pending.TryPop(out var item))
        {
            token.ThrowIfCancellationRequested();
            if (item.Node is DocxMathRow row && (row.HasSupportedProperties ||
                ReferenceEquals(row, source) && allowArgumentSize && row.HasKnownArgumentProperties))
            {
                int bookmarkIndex = row.Bookmarks.Count - 1;
                for (int child = row.Children.Count; child >= 0; child--)
                {
                    token.ThrowIfCancellationRequested();
                    if (child < row.Children.Count) pending.Push((row.Children[child], null));
                    while (bookmarkIndex >= 0 && row.Bookmarks[bookmarkIndex].ChildIndex == child)
                    {
                        token.ThrowIfCancellationRequested();
                        pending.Push((null, row.Bookmarks[bookmarkIndex--].Source));
                    }
                }
            }
            else yield return item;
        }
    }
}
