using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    private readonly record struct PlacedNoteTableFrame(IReadOnlyList<DocxLayoutItem> Items, PdfRectangle Clip);

    private static bool IsNormalNoteStory(DocxRelatedStory story) =>
        story.IsNormalStoryType && story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote;

    private static PdfRectangle ResolvePlacedStoryDesignClip(DocxPlacedRelatedStoryLayout story) =>
        new(story.X, story.TopY - story.Height, story.Width, story.Height);

    private static (bool Mapped, FloatingTextBoxEmissionMap Map, PdfRectangle Clip, PdfRectangle TableClip) ResolvePlacedStoryPaintFrame(
        DocxPlacedRelatedStoryLayout story, DocxMarkupContext context, double pageHeight)
    {
        FloatingTextBoxEmissionMap map = default;
        bool note = story.StoryLayout.Story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote;
        bool mapped = note &&
            FloatingTextBoxEmissionMap.TryCreate(context, pageHeight, out map);
        PdfRectangle clip;
        if (mapped)
        {
            var (x, y, width, height) = map.MapRectangle(story.X, story.TopY, story.Width, story.Height);
            clip = new PdfRectangle(x, y, width, height);
        }
        else clip = ResolvePlacedStoryDesignClip(story);

        // Flow admission bounds the selected text/image slice. Table cells own
        // their content clips; an unfragmented row can extend into the margin,
        // where Office clips it at the paper edge instead of the body bottom.
        // Keep the horizontal/top story frame and every cell/ancestor clip.
        PdfRectangle tableClip = note && story.StoryLayout.Story.IsNormalStoryType
            ? new PdfRectangle(clip.X, 0d, clip.Width, Math.Clamp(clip.Y + clip.Height, 0d, pageHeight))
            : clip;
        return (mapped, map, clip, tableClip);
    }

    // Complete fitting note tables share their story's design-to-print map.
    // Unfit, exact, fragmented or unresolved owners keep their staged paint path.
    // Drawing links consume this same row instead of reconstructing its frame.
    private static DocxTableRowLayout ResolvePlacedStoryTablePaintFrame(
        DocxPlacedRelatedStoryLayout story, DocxTableRowLayout row,
        FloatingTextBoxEmissionMap? map, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (map is not { } mapping || !story.StoryLayout.Story.IsNormalStoryType)
            return row;
        var bounds = ResolvePlacedStoryDesignClip(story);
        return Fits(row, bounds) ? mapping.PrecompensateRow(row) : row;

        bool Fits(DocxTableRowLayout current, PdfRectangle parent)
        {
            token.ThrowIfCancellationRequested();
            const double roundoff = .000001d;
            if (current.FragmentCount != 1 || current.Cells.Count == 0 ||
                !DocxLayoutEngine.HasAutomaticTableRowHeight(current.DeclaredHeightPoints, current.HeightRuleValue) &&
                !DocxLayoutEngine.HasMinimumTableRowHeight(current.DeclaredHeightPoints, current.HeightRuleValue) ||
                !Contains(parent, parent.X, current.Y, parent.Width, current.Height) ||
                current.Y < -roundoff || current.Y + current.Height > mapping.PageHeight + roundoff ||
                mapping.MapEmissionY(current.Y) < -roundoff ||
                mapping.MapEmissionY(current.Y + current.Height) > mapping.PageHeight + roundoff) return false;
            foreach (var cell in current.Cells)
            {
                token.ThrowIfCancellationRequested();
                if (cell.VisualOwnership != DocxTableCellVisualOwnership.OwnCell ||
                    !DocxMathParagraphPolicy.IsSupportedCellProperties(cell.VisualCell) ||
                    !Contains(parent, cell.X, cell.Y, cell.Width, cell.Height)) return false;
                var cellBounds = new PdfRectangle(cell.X, cell.Y, cell.Width, cell.Height);
                foreach (var child in cell.NestedRows)
                    if (!Fits(child, cellBounds)) return false;
            }
            return true;
        }

        static bool Contains(PdfRectangle frame, double x, double y, double width, double height)
        {
            const double roundoff = .000001d;
            return double.IsFinite(frame.X) && double.IsFinite(frame.Y) &&
                double.IsFinite(frame.Width) && double.IsFinite(frame.Height) && frame.Width > 0d && frame.Height > 0d &&
                double.IsFinite(x) && double.IsFinite(y) && double.IsFinite(width) && double.IsFinite(height) &&
                width > 0d && height > 0d && x >= frame.X - roundoff && y >= frame.Y - roundoff &&
                x + width <= frame.X + frame.Width + roundoff && y + height <= frame.Y + frame.Height + roundoff;
        }
    }

    private static PlacedNoteTableFrame? RenderPlacedRelatedStory(
        DocxPlacedRelatedStoryLayout story,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        PictureNavigation navigation)
    {
        // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote and endnote
        // stories lay out unscaled in design space while emission maps uniformly
        // (WC first baseline 236.62 = affine-mapped design). Other stories keep
        // the body-shift path. Complete fitting table subtrees join this map;
        // their unsupported ownership/overflow cases retain the legacy frame.
        var (mapStory, storyMap, clip, tableClip) = ResolvePlacedStoryPaintFrame(story, markupContext, navigation.Viewport.Height);
        bool normalNote = IsNormalNoteStory(story.StoryLayout.Story);
        var storyNavigation = navigation with
        {
            AllowsRotatedInlineUris = navigation.AllowsRotatedInlineUris && normalNote
        };
        if (story.SeparatorY is { } separatorY)
        {
            graphics.SetFillRgb(0, 0, 0);
            graphics.FillRectangle(
                mapStory ? storyMap.MapEmissionX(story.X) : story.X,
                mapStory ? storyMap.MapEmissionY(separatorY) : separatorY,
                Math.Min(mapStory ? storyMap.ScaleExtent(story.SeparatorWidth) : story.SeparatorWidth, mapStory ? storyMap.ScaleExtent(story.Width) : story.Width),
                mapStory ? storyMap.ScaleExtent(story.SeparatorThickness) : story.SeparatorThickness);
        }

        graphics.SaveState();
        graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);

        // RV06 height-model probes (Word COM reference edge-endsepheight-text): Office
        // drops non-mark separator text while drawing the rule. Skip text-bearing separator lines
        // at emission (layout keeps them so placement boxes stay Office-true); endnotes only,
        // footnotes keep legacy behavior until separately probed.
        bool suppressSeparatorText = story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && story.StoryLayout.Story.Type == DocxRelatedStoryType.Separator;
        IEnumerable<DocxTextLineLayout> storyTextLines = suppressSeparatorText
            ? story.TextLines.Select(line => RewriteSeparatorTextLineAsMarkSpace(line, story, fontResources))
            : story.TextLines;
        var items = (mapStory ? storyTextLines.Select(line => storyMap.PrecompensateLine(line, 0d, 0d)) : storyTextLines)
            .Cast<DocxLayoutItem>()
            .Concat(mapStory ? story.InlineImages.Select(image => storyMap.PrecompensateImage(image, pageNumber)) : story.InlineImages)
            .Concat(story.TableRows.Select(row => ResolvePlacedStoryTablePaintFrame(story, row, mapStory ? storyMap : null, cancellationToken)))
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        RenderParagraphBackgrounds(items, graphics, markupContext, cancellationToken);
        for (int itemIndex = 0; itemIndex < items.Length; itemIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutItem item = items[itemIndex];
            bool tableViewport = item is DocxTableRowLayout && tableClip != clip;
            if (tableViewport) SetClip(tableClip);
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Length ? items[itemIndex + 1] as DocxTableRowLayout : null;
            if (diagnosticSink is not null && item is DocxTableRowLayout row &&
                story.StoryLayout.Story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
                HasClippedNoteTableContent(row, tableClip, ResolveTextEmissionBaselineOffset(markupContext), cancellationToken))
                diagnosticSink(new OoxPdfDiagnostic("DOCX_CLIPPED_NOTE_TABLE_ROW", OoxPdfSeverity.Warning,
                    "A note-table row exceeds its available frame; content is clipped. Row fragmentation for this content is unsupported.",
                    story.StoryLayout.Story.PartName, PageIndex: pageNumber, Feature: "note table row", Fallback: "Clipped"));
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex,
                storyNavigation.ClipTo(tableViewport ? tableClip : clip));
            if (tableViewport) SetClip(clip);
        }

        graphics.RestoreState();

        // Text and images between rows matter to fragment adjacency. Pass the
        // ordered paint list to this page's annotation pass.
        return normalNote && story.TableRows.Count != 0 ? new(items, tableClip) : null;

        void SetClip(PdfRectangle viewport)
        {
            // PDF clips intersect. Restore the surrounding state before
            // selecting another container, preserving sibling paint order.
            graphics.RestoreState();
            graphics.SaveState();
            graphics.ClipRectangle(viewport.X, viewport.Y, viewport.Width, viewport.Height);
        }
    }

    private static bool HasClippedNoteTableContent(DocxTableRowLayout row, PdfRectangle clip,
        double baselineOffset, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        // Authored exact heights and rows with real continuation fragments keep
        // their intentional clipping. The cell iterator excludes text already
        // clipped by a cell/ancestor, including nested exact-height cells.
        if (row.FragmentCount != 1 ||
            !DocxLayoutEngine.HasAutomaticTableRowHeight(row.DeclaredHeightPoints, row.HeightRuleValue)) return false;
        foreach (var (line, _) in DocxTableCellClipGeometry.EnumerateCellLines(row, null, token))
        {
            if (!string.IsNullOrWhiteSpace(line.Text) &&
                DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double bottom, out _) &&
                bottom + baselineOffset < clip.Y - .000001d) return true;
        }
        // Cell pictures paint with the cell's baseline translation. Only a
        // complete source-contained frame demonstrates loss at the note clip;
        // authored cell/ancestor/view clipping retains its existing policy.
        bool HasClippedImages(DocxTableRowLayout current)
        {
            foreach (var cell in current.Cells)
            {
                token.ThrowIfCancellationRequested();
                foreach (var image in cell.InlineImages)
                {
                    token.ThrowIfCancellationRequested();
                    if (image.Image.RotationDegrees == 0d && image.Y - baselineOffset < clip.Y - .000001d) return true;
                }
                foreach (var nested in cell.NestedRows)
                    if (HasClippedImages(nested)) return true;
            }
            return false;
        }
        return HasClippedImages(row) && DocxTableCellClipGeometry.HasCompleteMeasuredRows([row], token);
    }

    // Separator text lines render as a single pilcrow-sized mark space at the rule tab
    // position (Office drops the text but keeps one space); whitespace-only mark lines pass
    // through unchanged. Layout keeps the text lines so placement boxes stay Office-true.
    private static DocxTextLineLayout RewriteSeparatorTextLineAsMarkSpace(DocxTextLineLayout line, DocxPlacedRelatedStoryLayout story, DocxFontResources fontResources)
    {
        if (!line.Text.Any(character => !char.IsWhiteSpace(character)))
        {
            return line;
        }
        double pilcrowSize = line.SourceParagraph?.ParagraphMarkFontSize ?? line.FontSize;
        double markX = story.X + Math.Min(144d, story.Width);
        if (line.Segments.Count == 0)
        {
            return line with { Text = " ", X = markX };
        }
        DocxTextSegmentLayout firstSegment = line.Segments[0];
        double spaceWidth = fontResources.TextMeasurer?.MeasureText(firstSegment.StyleRun, " ", pilcrowSize) ?? firstSegment.Width;
        DocxTextSegmentLayout markSpace = firstSegment with { Text = " ", X = markX, Width = spaceWidth, FontSize = pilcrowSize };
        return line with { Text = " ", Width = spaceWidth, Segments = [markSpace] };
    }

    private static void RenderFloatingDrawings(
        IReadOnlyList<DocxFloatingDrawingLayout> pageDrawings,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        PictureNavigation navigation)
    {
        // R12: pageDrawings arrives pre-filtered by page/layer and pre-sorted by
        // z-order from the once-per-render page index.
        foreach (DocxFloatingDrawingLayout drawing in pageDrawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, navigation);
        }
    }

    private static void RenderPlacedRelatedStoryDrawings(
        DocxPlacedRelatedStoryLayout story,
        bool behindDocument,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        PictureNavigation navigation)
    {
        DocxFloatingDrawingLayout[] drawings = story.FloatingDrawings
            .Where(drawing => IsBehindDocument(drawing.Drawing) == behindDocument)
            .OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue))
            .ToArray();
        if (drawings.Length == 0)
        {
            return;
        }

        var (mapped, map, clip, _) = ResolvePlacedStoryPaintFrame(story, markupContext, navigation.Viewport.Height);
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, navigation, clip, mapped ? map : null);
        }
    }

    // Word's in-cell layout flag applies to wrapped shapes. A wrap-none
    // drawing still uses its paragraph's anchor and overlays the page.
    private static PdfRectangle? ResolveFloatingCellViewport(DocxFloatingDrawingLayout drawing) =>
        drawing.Drawing.WrapKind is not (null or DocxFloatingWrapKind.None) &&
        OoxBoolean.IsTrue(drawing.Drawing.LayoutInCellValue) ? drawing.StaticCellClip : null;

    private static void RenderFloatingDrawing(
        DocxFloatingDrawingLayout drawing,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        PictureNavigation navigation,
        PdfRectangle? storyClip = null,
        FloatingTextBoxEmissionMap? storyMap = null)
    {
        if (drawing.PlacedX is not { } placedX ||
            drawing.PlacedTop is not { } placedTop ||
            drawing.ExtentWidthPoints is not { } width ||
            drawing.ExtentHeightPoints is not { } height)
        {
            return;
        }

        PdfRectangle? drawingClip = ResolveFloatingCellViewport(drawing) ?? storyClip;
        if (drawing.Drawing.Image is { } image)
        {
            if (drawingClip is { } clip)
            {
                graphics.SaveState();
                graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);
            }
            if (storyMap is null && drawing.StaticAnchorY is not null &&
                FloatingTextBoxEmissionMap.TryCreate(markupContext, navigation.Viewport.Height, out var staticMap))
                storyMap = staticMap;
            var pictureFrame = storyMap is { } pictureMap
                ? pictureMap.MapRectangle(placedX, placedTop, width, height)
                : (X: placedX, Y: placedTop - height, Width: width, Height: height);
            bool painted = RenderImage(image, pictureFrame.X, pictureFrame.Y, pictureFrame.Width, pictureFrame.Height, graphics, pageImages,
                diagnosticSink, drawing.AnchorPageIndex ?? 0, cancellationToken, imageCache, ref imageIndex,
                out PdfRectangle? axisAlignedFrame, svgTextFonts: fontResources.SvgTextFonts);
            if (drawingClip is not null) graphics.RestoreState();
            if (painted)
                (drawingClip is { } viewport ? navigation.ClipTo(viewport) : navigation).RecordUri(image, axisAlignedFrame);
        }

        if (drawing.TextBoxLayout is { } textBoxLayout)
        {
            RenderFloatingTextBox(drawing, textBoxLayout, placedX, placedTop, width, height, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, navigation, drawingClip);
        }
    }

    private readonly record struct FloatingTextBoxPaintFrame(double X, double Top, PdfRectangle Clip, FloatingTextBoxEmissionMap? Map)
    {
        public DocxTextLineLayout PlaceLine(DocxTextLineLayout line) => Map is { } map
            ? map.PrecompensateLine(line, X, Top)
            : line.Translate(X, Top);

        public DocxTableRowLayout PlaceRow(DocxTableRowLayout row)
        {
            var translated = DocxLayoutEngine.ShiftTableRow(row, Top, X);
            return Map is { } map ? map.PrecompensateRow(translated) : translated;
        }
    }

    private static FloatingTextBoxPaintFrame ResolveFloatingTextBoxPaintFrame(DocxFloatingDrawing drawing,
        double placedX, double placedTop, double width, double height, DocxMarkupContext context, double pageHeight)
    {
        DocxLayoutEngine.ResolveTextBoxContentInsets(drawing, out double left, out double top, out double right, out double bottom);
        double contentX = placedX + left;
        double contentTop = placedTop - top;
        double contentWidth = Math.Max(0d, width - left - right);
        double contentHeight = Math.Max(0d, height - top - bottom);
        if (FloatingTextBoxEmissionMap.TryCreate(context, pageHeight, out var map))
        {
            var (clipX, clipY, clipWidth, clipHeight) = map.MapRectangle(contentX, contentTop, contentWidth, contentHeight);
            return new(contentX, contentTop, new PdfRectangle(clipX, clipY, clipWidth, clipHeight), map);
        }

        return new(contentX, contentTop, new PdfRectangle(contentX, contentTop - contentHeight, contentWidth, contentHeight), null);
    }

    private static void RenderFloatingTextBox(
        DocxFloatingDrawingLayout drawing,
        DocxRelatedStoryLayout textBoxLayout,
        double placedX,
        double placedTop,
        double width,
        double height,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        int pageNumber,
        int pageCount,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        PictureNavigation navigation,
        PdfRectangle? storyClip)
    {
        if (textBoxLayout.TextLines.Count == 0 &&
            textBoxLayout.InlineImages.Count == 0 &&
            textBoxLayout.TableRows.Count == 0)
        {
            return;
        }

        var paintFrame = ResolveFloatingTextBoxPaintFrame(
            drawing.Drawing, placedX, placedTop, width, height, markupContext, navigation.Viewport.Height);
        var (_, _, clip, mapping) = paintFrame;
        var boxNavigation = (storyClip is { } viewport ? navigation.ClipTo(viewport) : navigation).EnterTextBox(clip);
        bool uniformMap = mapping.HasValue;
        var map = mapping.GetValueOrDefault();
        graphics.SaveState();
        if (storyClip is { } ancestor)
            graphics.ClipRectangle(ancestor.X, ancestor.Y, ancestor.Width, ancestor.Height);
        // The frame owns its outer stroke; only its children acquire the inset
        // content clip, as for inline text boxes. Clipping first hides the frame.
        if (TryResolveTextBoxStroke(drawing.Drawing.TextBoxStrokeWidthEmuValue, drawing.Drawing.TextBoxStrokeColorHex,
                width, height, out double strokeWidth, out RgbColor strokeColor))
        {
            var frame = uniformMap ? map.MapRectangle(placedX, placedTop, width, height)
                : (X: placedX, Y: placedTop - height, Width: width, Height: height);
            graphics.SetStrokeRgb(strokeColor.Red, strokeColor.Green, strokeColor.Blue);
            graphics.SetLineWidth(uniformMap ? map.ScaleExtent(strokeWidth) : strokeWidth);
            graphics.StrokeRectangle(frame.X, frame.Y, frame.Width, frame.Height);
        }
        graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);
        var items = PlaceFloatingTextBoxItems(textBoxLayout, paintFrame, drawing.AnchorPageIndex);
        RenderParagraphBackgrounds(items, graphics, markupContext, cancellationToken);
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex, boxNavigation);
        }

        graphics.RestoreState();
    }

    private static IReadOnlyList<DocxLayoutItem> PlaceFloatingTextBoxItems(DocxRelatedStoryLayout box,
        FloatingTextBoxPaintFrame frame, int? pageIndex) => box.TextLines.Select(frame.PlaceLine)
        .Cast<DocxLayoutItem>()
        .Concat(box.InlineImages.Select(image => frame.Map is { } map
            ? map.PrecompensateImage(image, frame.X, frame.Top, pageIndex ?? image.PageIndex)
            : image.Relocate(frame.X + image.X, frame.Top + image.Y, pageIndex)))
        .Concat(box.TableRows.Select(frame.PlaceRow))
        .OrderByDescending(ResolveLayoutItemTop)
        .ToArray();

    private static bool TryResolveTextBoxStroke(string? strokeWidthEmuValue, string? strokeColorHex,
        double width, double height, out double strokeWidth, out RgbColor strokeColor)
    {
        strokeWidth = 0d;
        strokeColor = default;
        if (width <= 0d || height <= 0d ||
            !long.TryParse(strokeWidthEmuValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out long strokeEmu) || strokeEmu <= 0L ||
            !RgbColor.TryParse(strokeColorHex, out strokeColor))
        {
            return false;
        }

        strokeWidth = OoxUnits.EmuToPoints(strokeEmu);
        return strokeWidth > 0d;
    }

    // Body boxes use the body emission offset. Physical static boxes already own their Y.
    private static (double X, double Top) ResolveInlineTextBoxPaintOrigin(
        DocxInlineTextBoxLayout box, DocxMarkupContext context) =>
        (box.BoxX + ResolveTextEmissionXOffset(context),
            box.BoxTop - (box.StaticStoryUsesPhysicalY ? 0d : ResolveTextEmissionBaselineOffset(context)));

    private static void RenderInlineTextBox(
        DocxInlineTextBoxLayout box,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int pageNumber,
        int pageCount,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        PictureNavigation navigation,
        DocxTableCellClipSnapshot? cellClip = null,
        RgbColor? background = null)
    {
        if (diagnosticSink is not null)
            foreach (var element in box.TextBox.BodyElements)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (element is not DocxParagraphElement { Paragraph.InlineTextBoxes.Count: > 0 }) continue;
                diagnosticSink(new OoxPdfDiagnostic("DOCX_UNSUPPORTED_NESTED_INLINE_TEXTBOX", OoxPdfSeverity.Warning,
                    "Nested inline text boxes in a text-box paragraph were ignored; surrounding supported content is retained.",
                    PartName: null, SlideIndex: null, PageIndex: pageNumber - 1,
                    Feature: "nested-inline-textbox", Fallback: "Ignored"));
                break;
            }
        if (box.TextLines.Count == 0 &&
            box.InlineImages.Count == 0 &&
            box.TableRows.Count == 0)
        {
            return;
        }

        var (clipX, clipTop) = ResolveInlineTextBoxPaintOrigin(box, markupContext);
        graphics.SaveState();
        // RV05 inline-box probe (Word COM reference edge-inlinebox): stroked inline
        // frames paint their outline under the text like floating boxes.
        if (TryResolveTextBoxStroke(box.TextBox.TextBoxStrokeWidthEmuValue, box.TextBox.TextBoxStrokeColorHex,
                box.BoxWidth, box.BoxHeight, out double inlineStrokeWidth, out RgbColor inlineStrokeColor))
        {
            double strokeWidth = inlineStrokeWidth * box.PrintScale;
            if (cellClip is not null)
            {
                graphics.SaveState();
                var outlineClip = DocxTableCellClipGeometry.ResolveInlineOutline(cellClip,
                    new PdfRectangle(clipX, clipTop - box.BoxHeight, box.BoxWidth, box.BoxHeight), strokeWidth, box.OutlineViewportBottom);
                graphics.ClipRectangle(outlineClip.X, outlineClip.Y, outlineClip.Width, outlineClip.Height);
            }
            double viewportBottom = box.OutlineViewportBottom.GetValueOrDefault(clipTop - box.BoxHeight);
            bool hasViewport = viewportBottom > clipTop - box.BoxHeight;
            if (hasViewport)
            {
                graphics.SaveState();
                // This viewport supplies a lower bound; the outline's own
                // envelope must not acquire a new lateral or upper content clip.
                graphics.ClipRectangle(clipX - strokeWidth / 2d, viewportBottom,
                    box.BoxWidth + strokeWidth, Math.Max(0d, clipTop + strokeWidth / 2d - viewportBottom));
            }
            graphics.SetStrokeRgb(inlineStrokeColor.Red, inlineStrokeColor.Green, inlineStrokeColor.Blue);
            graphics.SetLineWidth(strokeWidth);
            graphics.StrokeRectangle(clipX, clipTop - box.BoxHeight, box.BoxWidth, box.BoxHeight);
            if (hasViewport) graphics.RestoreState();
            if (cellClip is not null) graphics.RestoreState();
        }
        if (cellClip is not null) graphics.ClipRectangle(cellClip.X, cellClip.Y, cellClip.Width, cellClip.Height);
        // The containing drawing viewport owns the outline. Only children
        // acquire this interior clip; clipping before stroking halves the border.
        var clip = new PdfRectangle(clipX, clipTop - box.BoxHeight, box.BoxWidth, box.BoxHeight);
        graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);
        var boxNavigation = navigation.EnterTextBox(clip);
        var items = GetInlineTextBoxPaintItems(box, markupContext);
        RenderParagraphBackgrounds(items, graphics, markupContext, cancellationToken);
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex, boxNavigation, background);
        }

        graphics.RestoreState();
    }

    private static IReadOnlyList<DocxLayoutItem> GetInlineTextBoxPaintItems(DocxInlineTextBoxLayout box, DocxMarkupContext context) =>
        box.TextLines.Cast<DocxLayoutItem>().Concat(box.InlineImages).Concat(box.TableRows)
            .OrderByDescending(ResolveLayoutItemTop)
            .Select(item => item.StaticStoryUsesPhysicalY ? AdjustStaticLayoutItem(item, context) : item)
            .ToArray();

    private static double ResolveLayoutItemTop(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout textLine => textLine.BaselineY,
            DocxInlineImageLayout image => image.Y + image.Height,
            DocxTableRowLayout row => row.Y + row.Height,
            DocxInlineTextBoxLayout box => box.BoxTop,
            _ => 0d
        };
    }

    private static bool IsBehindDocument(DocxFloatingDrawing drawing)
    {
        // Keep the legacy empty-attribute default; share XML lexical normalization.
        return drawing.BehindDocumentValue is { } value && (value.Length == 0 || OoxBoolean.IsTrue(value));
    }

    private static long ReadZOrder(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long zOrder)
            ? zOrder
            : 0L;
    }
}
