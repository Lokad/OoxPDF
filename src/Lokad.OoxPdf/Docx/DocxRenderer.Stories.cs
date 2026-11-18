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

    private static void RenderPlacedRelatedStory(
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
        double pageHeight,
        HashSet<DocxInlineImage> paintedPictures)
    {
        // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote and endnote
        // stories lay out unscaled in design space while emission maps uniformly
        // (WC first baseline 236.62 = affine-mapped design). Other stories keep
        // the body-shift path. Complete fitting table subtrees join this map;
        // their unsupported ownership/overflow cases retain the legacy frame.
        var (mapStory, storyMap, clip, tableClip) = ResolvePlacedStoryPaintFrame(story, markupContext, pageHeight);
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
        System.Collections.Generic.IEnumerable<DocxTextLineLayout> storyTextLines = suppressSeparatorText ? story.TextLines.Select(line => RewriteSeparatorTextLineAsMarkSpace(line, story, fontResources)) : story.TextLines;
        IReadOnlyList<DocxLayoutItem> items = (mapStory ? storyTextLines.Select(line => storyMap.PrecompensateLine(line, 0d, 0d)) : storyTextLines)
            .Cast<DocxLayoutItem>()
            .Concat(mapStory ? story.InlineImages.Select(image => storyMap.PrecompensateImage(image, pageNumber)) : story.InlineImages)
            .Concat(story.TableRows.Select(row => ResolvePlacedStoryTablePaintFrame(
                story, row, mapStory ? storyMap : null, cancellationToken)))
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        RenderParagraphBackgrounds(items, graphics, markupContext, cancellationToken);
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutItem item = items[itemIndex];
            bool tableViewport = item is DocxTableRowLayout && tableClip != clip;
            if (tableViewport) SetClip(tableClip);
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            if (diagnosticSink is not null && item is DocxTableRowLayout row &&
                story.StoryLayout.Story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
                HasClippedNoteTableContent(row, tableClip, ResolveTextEmissionBaselineOffset(markupContext), cancellationToken))
                diagnosticSink(new OoxPdfDiagnostic("DOCX_CLIPPED_NOTE_TABLE_ROW", OoxPdfSeverity.Warning,
                    "A note-table row exceeds its available frame; content is clipped. Row fragmentation for this content is unsupported.",
                    story.StoryLayout.Story.PartName, PageIndex: pageNumber, Feature: "note table row", Fallback: "Clipped"));
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex, paintedPictures: paintedPictures);
            if (tableViewport) SetClip(clip);
        }

        graphics.RestoreState();

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
        PdfRectangle pageClip,
        ref List<PdfLinkAnnotation>? pictureLinks)
    {
        // R12: pageDrawings arrives pre-filtered by page/layer and pre-sorted by
        // z-order from the once-per-render page index.
        foreach (DocxFloatingDrawingLayout drawing in pageDrawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, pageClip, ref pictureLinks);
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
        PdfRectangle pageClip,
        ref List<PdfLinkAnnotation>? pictureLinks)
    {
        DocxFloatingDrawingLayout[] drawings = story.FloatingDrawings
            .Where(drawing => IsBehindDocument(drawing.Drawing) == behindDocument)
            .OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue))
            .ToArray();
        if (drawings.Length == 0)
        {
            return;
        }

        var imageClip = ResolvePlacedStoryDesignClip(story);
        // Floating boxes join the printed story frame. Floating pictures keep
        // the viewport of their existing design-space placement.
        var (_, _, textBoxClip, _) = ResolvePlacedStoryPaintFrame(story, markupContext, pageClip.Height);
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, pageClip, ref pictureLinks, imageClip, textBoxClip);
        }
    }

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
        PdfRectangle pageClip,
        ref List<PdfLinkAnnotation>? pictureLinks,
        PdfRectangle? imageClip = null,
        PdfRectangle? textBoxClip = null)
    {
        if (drawing.PlacedX is not { } placedX ||
            drawing.PlacedTop is not { } placedTop ||
            drawing.ExtentWidthPoints is not { } width ||
            drawing.ExtentHeightPoints is not { } height)
        {
            return;
        }

        if (drawing.Drawing.Image is { } image)
        {
            if (imageClip is { } clip)
            {
                graphics.SaveState();
                graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);
            }
            var pictureFrame = drawing.StaticAnchorY is not null &&
                FloatingTextBoxEmissionMap.TryCreate(markupContext, pageClip.Height, out var pictureMap)
                ? pictureMap.MapRectangle(placedX, placedTop, width, height)
                : (X: placedX, Y: placedTop - height, Width: width, Height: height);
            bool painted = RenderImage(image, pictureFrame.X, pictureFrame.Y, pictureFrame.Width, pictureFrame.Height, graphics, pageImages,
                diagnosticSink, drawing.AnchorPageIndex ?? 0, cancellationToken, imageCache, ref imageIndex,
                out PdfRectangle? axisAlignedFrame, svgTextFonts: fontResources.SvgTextFonts);
            if (imageClip is not null) graphics.RestoreState();
            if (painted && image.ClickTarget is { } uri && axisAlignedFrame is { } frame &&
                ContainsPictureFrame(pageClip, frame) && (imageClip is not { } viewport || ContainsPictureFrame(viewport, frame)))
            {
                (pictureLinks ??= []).Add(PdfLinkAnnotation.ToUri(frame.X, frame.Y, frame.Width, frame.Height, uri));
            }
        }

        if (drawing.TextBoxLayout is { } textBoxLayout)
        {
            RenderFloatingTextBox(drawing, textBoxLayout, placedX, placedTop, width, height, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, pageClip.Height, textBoxClip);
        }
    }

    private readonly record struct FloatingTextBoxPaintFrame(double X, double Top, PdfRectangle Clip, FloatingTextBoxEmissionMap? Map)
    {
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
        double pageHeight,
        PdfRectangle? storyClip)
    {
        if (textBoxLayout.TextLines.Count == 0 &&
            textBoxLayout.InlineImages.Count == 0 &&
            textBoxLayout.TableRows.Count == 0)
        {
            return;
        }

        var paintFrame = ResolveFloatingTextBoxPaintFrame(
            drawing.Drawing, placedX, placedTop, width, height, markupContext, pageHeight);
        var (contentX, contentTop, clip, mapping) = paintFrame;
        bool uniformMap = mapping.HasValue;
        var map = mapping.GetValueOrDefault();
        graphics.SaveState();
        if (storyClip is { } ancestor)
            graphics.ClipRectangle(ancestor.X, ancestor.Y, ancestor.Width, ancestor.Height);
        // The frame owns its outer stroke; only its children acquire the inset
        // content clip, as for inline text boxes. Clipping first hides the frame.
        if (TryResolveTextBoxStroke(drawing.Drawing, width, height, out double strokeWidth, out byte strokeRed, out byte strokeGreen, out byte strokeBlue))
        {
            var frame = uniformMap ? map.MapRectangle(placedX, placedTop, width, height)
                : (X: placedX, Y: placedTop - height, Width: width, Height: height);
            graphics.SetStrokeRgb(strokeRed, strokeGreen, strokeBlue);
            graphics.SetLineWidth(uniformMap ? map.ScaleExtent(strokeWidth) : strokeWidth);
            graphics.StrokeRectangle(frame.X, frame.Y, frame.Width, frame.Height);
        }
        graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);
        IReadOnlyList<DocxLayoutItem> items = textBoxLayout.TextLines
            .Select(line => uniformMap
                ? map.PrecompensateLine(line, contentX, contentTop)
                : TranslateTextLine(line, contentX, contentTop))
            .Cast<DocxLayoutItem>()
            .Concat(textBoxLayout.InlineImages.Select(image => uniformMap
                ? map.PrecompensateImage(image, contentX, contentTop, drawing.AnchorPageIndex ?? image.PageIndex)
                : image.Relocate(contentX + image.X, contentTop + image.Y, drawing.AnchorPageIndex)))
            .Concat(textBoxLayout.TableRows.Select(paintFrame.PlaceRow))
            .OrderByDescending(ResolveLayoutItemTop)
            .ToArray();
        RenderParagraphBackgrounds(items, graphics, markupContext, cancellationToken);
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex);
        }

        graphics.RestoreState();
    }

    private static bool TryResolveTextBoxStroke(DocxFloatingDrawing drawing, double width, double height, out double strokeWidth, out byte strokeRed, out byte strokeGreen, out byte strokeBlue)
    {
        return TryResolveTextBoxStrokeGeometry(drawing.TextBoxStrokeWidthEmuValue, drawing.TextBoxStrokeColorHex, width, height, out strokeWidth, out strokeRed, out strokeGreen, out strokeBlue);
    }

    private static bool TryResolveInlineTextBoxStroke(DocxInlineTextBox textBox, double width, double height, out double strokeWidth, out byte strokeRed, out byte strokeGreen, out byte strokeBlue)
    {
        return TryResolveTextBoxStrokeGeometry(textBox.TextBoxStrokeWidthEmuValue, textBox.TextBoxStrokeColorHex, width, height, out strokeWidth, out strokeRed, out strokeGreen, out strokeBlue);
    }

    private static bool TryResolveTextBoxStrokeGeometry(string? strokeWidthEmuValue, string? strokeColorHex, double width, double height, out double strokeWidth, out byte strokeRed, out byte strokeGreen, out byte strokeBlue)
    {
        strokeWidth = 0d;
        strokeRed = 0;
        strokeGreen = 0;
        strokeBlue = 0;
        if (width <= 0d || height <= 0d || string.IsNullOrEmpty(strokeColorHex) || strokeColorHex.Length != 6)
        {
            return false;
        }

        if (!long.TryParse(strokeWidthEmuValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out long strokeEmu) || strokeEmu <= 0L)
        {
            return false;
        }

        if (!byte.TryParse(strokeColorHex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out strokeRed) ||
            !byte.TryParse(strokeColorHex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out strokeGreen) ||
            !byte.TryParse(strokeColorHex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out strokeBlue))
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
        if (TryResolveInlineTextBoxStroke(box.TextBox, box.BoxWidth, box.BoxHeight, out double inlineStrokeWidth, out byte inlineStrokeRed, out byte inlineStrokeGreen, out byte inlineStrokeBlue))
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
            graphics.SetStrokeRgb(inlineStrokeRed, inlineStrokeGreen, inlineStrokeBlue);
            graphics.SetLineWidth(strokeWidth);
            graphics.StrokeRectangle(clipX, clipTop - box.BoxHeight, box.BoxWidth, box.BoxHeight);
            if (hasViewport) graphics.RestoreState();
            if (cellClip is not null) graphics.RestoreState();
        }
        if (cellClip is not null) graphics.ClipRectangle(cellClip.X, cellClip.Y, cellClip.Width, cellClip.Height);
        // The containing drawing viewport owns the outline. Only children
        // acquire this interior clip; clipping before stroking halves the border.
        graphics.ClipRectangle(clipX, clipTop - box.BoxHeight, box.BoxWidth, box.BoxHeight);
        IReadOnlyList<DocxLayoutItem> items = box.TextLines
            .Cast<DocxLayoutItem>()
            .Concat(box.InlineImages)
            .Concat(box.TableRows)
            .OrderByDescending(ResolveLayoutItemTop)
            .Select(item => item.StaticStoryUsesPhysicalY
                ? AdjustStaticLayoutItem(item, markupContext) : item)
            .ToArray();
        RenderParagraphBackgrounds(items, graphics, markupContext, cancellationToken);
        for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxLayoutItem item = items[itemIndex];
            DocxTableRowLayout? previousRow = itemIndex > 0 ? items[itemIndex - 1] as DocxTableRowLayout : null;
            DocxTableRowLayout? nextRow = itemIndex + 1 < items.Count ? items[itemIndex + 1] as DocxTableRowLayout : null;
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex, background);
        }

        graphics.RestoreState();
    }

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

    private static DocxTextLineLayout TranslateTextLine(DocxTextLineLayout line, double deltaX, double deltaY)
    {
        return line with
        {
            X = line.X + deltaX,
            BaselineY = line.BaselineY + deltaY,
            Segments = line.Segments
                .Select(segment => segment with { X = segment.X + deltaX })
                .ToArray()
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
