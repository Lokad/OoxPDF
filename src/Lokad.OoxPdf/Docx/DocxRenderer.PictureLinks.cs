using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private static IEnumerable<PdfLinkAnnotation> CreateInlinePictureAnnotations(DocxLayoutPage page,
        IReadOnlyList<PdfImageResource> resources, IReadOnlyDictionary<string, PdfImageXObject?> imageCache,
        DocxMarkupContext context, CancellationToken token)
    {
        if (resources.Count == 0) yield break;
        var pageClip = new PdfRectangle(0d, 0d, page.Width, page.Height);
        HashSet<PdfImageXObject>? painted = null;
        foreach (var (image, clip) in EnumerateImages())
        {
            token.ThrowIfCancellationRequested();
            if (image.Image.ClickTarget is not { } uri || image.Image.RotationDegrees != 0d ||
                GetImageCacheKey(image.Image) is not { } key || !imageCache.TryGetValue(key, out var resource) || resource is null)
                continue;
            var frame = new PdfRectangle(image.X, image.Y, image.Width, image.Height);
            if (!ContainsFrame(clip, frame) || image.ClipBounds is { } own && !ContainsFrame(own, frame) ||
                image.LineFrame is { } line && !(frame.Y >= line.Bottom - 1e-6 && frame.Y + frame.Height <= line.Bottom + line.Height + 1e-6))
                continue;
            painted ??= new HashSet<PdfImageXObject>(resources.Select(item => item.Image), ReferenceEqualityComparer.Instance);
            if (painted.Contains(resource))
                yield return PdfLinkAnnotation.ToUri(frame.X, frame.Y, frame.Width, frame.Height, uri);
        }

        IEnumerable<(DocxInlineImageLayout Image, PdfRectangle Clip)> EnumerateImages()
        {
            foreach (var item in page.Items)
            {
                token.ThrowIfCancellationRequested();
                if (item is DocxInlineImageLayout image && image.Image.ClickTarget is not null)
                    yield return (PlaceInlineImageForPaint(image, context), pageClip);
                else if (item is DocxTableRowLayout row)
                    foreach (var child in CellImages(row, pageClip)) yield return child;
            }
        }

        IEnumerable<(DocxInlineImageLayout Image, PdfRectangle Clip)> CellImages(DocxTableRowLayout row, PdfRectangle parent)
        {
            token.ThrowIfCancellationRequested();
            // Complete unmerged owners have one physical cell clip. Fragmented
            // and merged ownership retains its staged navigation fallback.
            if (row.FragmentCount != 1) yield break;
            double xOffset = ResolveTextEmissionXOffset(context), yOffset = ResolveTextEmissionBaselineOffset(context);
            DocxTableRowLayout? geometryRow = null;
            for (int index = 0; index < row.Cells.Count; index++)
            {
                token.ThrowIfCancellationRequested();
                var cell = row.Cells[index];
                if (cell.VisualOwnership != DocxTableCellVisualOwnership.OwnCell || cell.Cell.HasVerticalMerge) continue;
                if (!cell.InlineImages.Any(image => image.Image.ClickTarget is not null) && cell.NestedRows.Count == 0) continue;
                geometryRow ??= xOffset == 0d && yOffset == 0d ? row : ShiftTableRowGeometry(row, xOffset, yOffset);
                var cellClip = ResolveTableCellContentClip(geometryRow, geometryRow.Cells[index]);
                double left = Math.Max(parent.X, cellClip.X), bottom = Math.Max(parent.Y, cellClip.Y);
                var clip = new PdfRectangle(left, bottom, Math.Min(parent.X + parent.Width, cellClip.X + cellClip.Width) - left,
                    Math.Min(parent.Y + parent.Height, cellClip.Y + cellClip.Height) - bottom);
                foreach (var image in cell.InlineImages)
                {
                    token.ThrowIfCancellationRequested();
                    if (image.Image.ClickTarget is not null)
                        yield return (PlaceCellImageForPaint(image, xOffset, yOffset, cellClip), clip);
                }
                foreach (var nested in cell.NestedRows)
                    foreach (var image in CellImages(nested, clip)) yield return image;
            }
        }

        static bool ContainsFrame(PdfRectangle outer, PdfRectangle frame) =>
            double.IsFinite(outer.X) && double.IsFinite(outer.Y) && double.IsFinite(outer.Width) && double.IsFinite(outer.Height) &&
            double.IsFinite(frame.X) && double.IsFinite(frame.Y) && double.IsFinite(frame.Width) && double.IsFinite(frame.Height) &&
            frame.Width > 0d && frame.Height > 0d && outer.Width > 0d && outer.Height > 0d &&
            frame.X >= outer.X - 1e-6 && frame.Y >= outer.Y - 1e-6 &&
            frame.X + frame.Width <= outer.X + outer.Width + 1e-6 && frame.Y + frame.Height <= outer.Y + outer.Height + 1e-6;
    }
}
