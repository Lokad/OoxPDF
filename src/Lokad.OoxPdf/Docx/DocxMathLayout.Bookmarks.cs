namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathLayout
{
    private static void AppendBookmarkPoints(ref List<DocxMathBookmarkPoint>? points, DocxMathBox child,
        double x, double baseline, CancellationToken token)
    {
        foreach (var point in child.BookmarkPoints)
        {
            token.ThrowIfCancellationRequested();
            (points ??= []).Add(point with { X = point.X + x, Top = point.Top + baseline });
        }
    }

    private static DocxMathBox WithStandaloneBookmarks(DocxEquation equation, DocxMathBox box, CancellationToken token)
    {
        // A standalone control owns one frame; surrounding source boundaries
        // select its start/end. Literal and compound rows publish their own points.
        if (GetSingleSupportedChild(equation) is not IDocxMathControl)
            return box;
        List<DocxMathBookmarkPoint>? points = null;
        DocxMathRow? row = equation.Content;
        while (row is not null)
        {
            token.ThrowIfCancellationRequested();
            foreach (var boundary in row.Bookmarks)
            {
                token.ThrowIfCancellationRequested();
                points ??= new(box.BookmarkPoints);
                points.Add(new(boundary.Source, boundary.ChildIndex == 0 ? 0d : box.AdvanceWidth, box.AboveBaseline));
            }
            row = row.Children is [DocxMathRow nested] ? nested : null;
        }
        return points is null ? box : box with { BookmarkPoints = points };
    }
}
