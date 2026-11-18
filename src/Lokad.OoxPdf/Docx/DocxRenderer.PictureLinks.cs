using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    // Carry the active paint viewport to links whose geometry depends on
    // painting, including pictures in boxes and rotated inline pictures.
    private readonly record struct PictureNavigation(
        PdfRectangle Viewport, HashSet<DocxInlineImage> InlinePictures, List<PdfLinkAnnotation> DrawingUris,
        bool InsideTextBox = false, bool AllowsRotatedInlineUris = false)
    {
        public PictureNavigation ClipTo(PdfRectangle clip) => this with
        {
            Viewport = IntersectPictureViewports(Viewport, clip)
        };

        public PictureNavigation EnterTextBox(PdfRectangle clip) => ClipTo(clip) with { InsideTextBox = true };

        public void RecordInline(DocxInlineImageLayout image, PdfRectangle? frame, bool usesImageClip, bool usesRotatedLineFrame)
        {
            if (!InsideTextBox)
                InlinePictures.Add(image.Image);

            if ((InsideTextBox || AllowsRotatedInlineUris && usesRotatedLineFrame) &&
                frame is { } painted && ContainsInlinePictureFrame(image, painted, usesImageClip))
                RecordUri(image.Image, painted);
        }

        public void RecordUri(DocxInlineImage image, PdfRectangle? frame)
        {
            if (image.ClickTarget is { } uri && frame is { } painted && ContainsPictureFrame(Viewport, painted))
                DrawingUris.Add(PdfLinkAnnotation.ToUri(painted.X, painted.Y, painted.Width, painted.Height, uri));
        }
    }

    private static bool OwnsUnmergedPictureCell(DocxTableCellLayout cell) =>
        cell.VisualOwnership == DocxTableCellVisualOwnership.OwnCell && !cell.Cell.HasVerticalMerge;

    private static IEnumerable<PdfLinkAnnotation> CreateInlinePictureAnnotations(DocxLayoutPage page, int pageNumber, IReadOnlyList<DocxLayoutItem> staticItems,
        IReadOnlySet<DocxInlineImage> paintedPictures,
        DocxMarkupContext context, CancellationToken token)
    {
        foreach (var image in EnumerateLinkedPictureFrames(page, pageNumber, staticItems, paintedPictures,
            context, static image => image.ClickTarget is not null, token))
            yield return PdfLinkAnnotation.ToUri(image.X, image.Y, image.Width, image.Height, image.Image.ClickTarget!);
    }

    // Ordinary placement also provides bounds for mixed source ranges.
    // Rotated drawing URIs use the actual paint frame instead.
    private static IEnumerable<DocxInlineImageLayout> EnumerateLinkedPictureFrames(DocxLayoutPage page, int pageNumber,
        IReadOnlyList<DocxLayoutItem> staticItems, IReadOnlySet<DocxInlineImage> paintedPictures,
        DocxMarkupContext context,
        Func<DocxInlineImage, bool> select, CancellationToken token)
    {
        if (paintedPictures.Count == 0) yield break;
        var pageClip = new PdfRectangle(0d, 0d, page.Width, page.Height);
        foreach (var (image, clip) in EnumeratePlacedPictureFrames(page, pageNumber, staticItems, context, select, token))
        {
            token.ThrowIfCancellationRequested();
            if (image.Image.RotationDegrees != 0d || !paintedPictures.Contains(image.Image))
                continue;
            var frame = new PdfRectangle(image.X, image.Y, image.Width, image.Height);
            if (!ContainsPictureFrame(pageClip, frame) || !ContainsPictureFrame(clip, frame) ||
                !ContainsInlinePictureFrame(image, frame, usesImageClip: true))
                continue;
            yield return image;
        }
    }

    private static bool ContainsInlinePictureFrame(
        DocxInlineImageLayout image, PdfRectangle frame, bool usesImageClip) =>
        (!usesImageClip || image.ClipBounds is not { } own || ContainsPictureFrame(own, frame)) &&
        (image.LineFrame is not { } line ||
            frame.Y >= line.Bottom - 1e-6 && frame.Y + frame.Height <= line.Bottom + line.Height + 1e-6);

    private static PdfRectangle IntersectPictureViewports(PdfRectangle parent, PdfRectangle viewport)
    {
        double x = Math.Max(parent.X, viewport.X), y = Math.Max(parent.Y, viewport.Y);
        return new PdfRectangle(x, y, Math.Min(parent.X + parent.Width, viewport.X + viewport.Width) - x,
            Math.Min(parent.Y + parent.Height, viewport.Y + viewport.Height) - y);
    }

    private static bool ContainsPictureFrame(PdfRectangle outer, PdfRectangle frame) =>
        double.IsFinite(outer.X) && double.IsFinite(outer.Y) && double.IsFinite(outer.Width) && double.IsFinite(outer.Height) &&
        double.IsFinite(frame.X) && double.IsFinite(frame.Y) && double.IsFinite(frame.Width) && double.IsFinite(frame.Height) &&
        frame.Width > 0d && frame.Height > 0d && outer.Width > 0d && outer.Height > 0d &&
        frame.X >= outer.X - 1e-6 && frame.Y >= outer.Y - 1e-6 &&
        frame.X + frame.Width <= outer.X + outer.Width + 1e-6 && frame.Y + frame.Height <= outer.Y + outer.Height + 1e-6;

    // Source destinations and URI links borrow the same physical placement.
    // URI links additionally require successful paint and complete visible clips.
    private static IEnumerable<(DocxInlineImageLayout Image, PdfRectangle Clip)> EnumeratePlacedPictureFrames(
        DocxLayoutPage page, int pageNumber, IEnumerable<DocxLayoutItem> staticItems,
        DocxMarkupContext context, Func<DocxInlineImage, bool> select, CancellationToken token)
    {
        var pageClip = new PdfRectangle(0d, 0d, page.Width, page.Height);
        foreach (var item in staticItems.Concat(page.Items))
        {
            token.ThrowIfCancellationRequested();
            if (item is DocxInlineImageLayout image && select(image.Image))
                yield return (PlaceInlineImageForPaint(image, context), pageClip);
            else if (item is DocxTableRowLayout row)
                foreach (var child in CellImages(row, pageClip)) yield return child;
        }
        foreach (var story in page.PlacedRelatedStories)
        {
            token.ThrowIfCancellationRequested();
            if (!story.StoryLayout.Story.IsNormalStoryType ||
                story.StoryLayout.Story.Kind is not (DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote)) continue;
            var (mapped, map, clip, tableClip) = ResolvePlacedStoryPaintFrame(story, context, page.Height);
            foreach (var image in story.InlineImages)
            {
                token.ThrowIfCancellationRequested();
                if (select(image.Image))
                    yield return (PlaceInlineImageForPaint(mapped ? map.PrecompensateImage(image, pageNumber) : image, context), clip);
            }
            // Drawing URI targets follow the row's chosen physical paint frame.
            foreach (var row in story.TableRows)
                foreach (var image in CellImages(ResolvePlacedStoryTablePaintFrame(
                    story, row, mapped ? map : null, token), tableClip)) yield return image;
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
                if (!OwnsUnmergedPictureCell(cell)) continue;
                if (!cell.InlineImages.Any(image => select(image.Image)) && cell.NestedRows.Count == 0) continue;
                geometryRow ??= xOffset == 0d && yOffset == 0d ? row : ShiftTableRowGeometry(row, xOffset, yOffset);
                var cellClip = ResolveTableCellContentClip(geometryRow, geometryRow.Cells[index]);
                var clip = IntersectPictureViewports(parent, new PdfRectangle(cellClip.X, cellClip.Y, cellClip.Width, cellClip.Height));
                foreach (var image in cell.InlineImages)
                {
                    token.ThrowIfCancellationRequested();
                    if (select(image.Image))
                        yield return (PlaceCellImageForPaint(image, xOffset, yOffset, cellClip), clip);
                }
                foreach (var nested in cell.NestedRows)
                    foreach (var image in CellImages(nested, clip)) yield return image;
            }
        }

    }
}
