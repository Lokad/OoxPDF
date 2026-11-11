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
            var (x, y, width, height) = map.MapClipRectangle(story.X, story.TopY, story.Width, story.Height);
            clip = new PdfRectangle(x, y, width, height);
        }
        else clip = new PdfRectangle(story.X, story.TopY - story.Height, story.Width, story.Height);

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
        var bounds = new PdfRectangle(story.X, story.TopY - story.Height, story.Width, story.Height);
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
        double pageHeight)
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
            RenderLayoutItem(item, previousRow, nextRow, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex);
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
        double pageHeight)
    {
        // R12: pageDrawings arrives pre-filtered by page/layer and pre-sorted by
        // z-order from the once-per-render page index.
        foreach (DocxFloatingDrawingLayout drawing in pageDrawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, pageHeight);
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
        double pageHeight)
    {
        DocxFloatingDrawingLayout[] drawings = story.FloatingDrawings
            .Where(drawing => IsBehindDocument(drawing.Drawing) == behindDocument)
            .OrderBy(drawing => ReadZOrder(drawing.Drawing.RelativeHeightValue))
            .ToArray();
        if (drawings.Length == 0)
        {
            return;
        }

        graphics.SaveState();
        graphics.ClipRectangle(story.X, story.TopY - story.Height, story.Width, story.Height);
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            RenderFloatingDrawing(drawing, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, pageHeight);
        }

        graphics.RestoreState();
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
        double pageHeight)
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
            RenderImage(image, placedX, placedTop - height, width, height, graphics, pageImages,
                diagnosticSink, drawing.AnchorPageIndex ?? 0, cancellationToken, imageCache, ref imageIndex);
        }

        if (drawing.TextBoxLayout is { } textBoxLayout)
        {
            RenderFloatingTextBox(drawing, textBoxLayout, placedX, placedTop, width, height, graphics, pageImages, fontResources, markupContext, pageNumber, pageCount, diagnosticSink, cancellationToken, imageCache, ref imageIndex, pageHeight);
        }
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
        double pageHeight)
    {
        if (textBoxLayout.TextLines.Count == 0 &&
            textBoxLayout.InlineImages.Count == 0 &&
            textBoxLayout.TableRows.Count == 0)
        {
            return;
        }

        // W6-d: content sits in the inset content box (explicit bodyPr insets, else Office
        // defaults), matching the layout content width.
        DocxLayoutEngine.ResolveTextBoxContentInsets(drawing.Drawing, out double insetLeft, out double insetTop, out double insetRight, out double insetBottom);
        double contentX = placedX + insetLeft;
        double contentTop = placedTop - insetTop;
        double contentWidth = Math.Max(0d, width - insetLeft - insetRight);
        double contentHeight = Math.Max(0d, height - insetTop - insetBottom);
        // Scaled WC pages map design-space content uniformly to emission space; every
        // other mode keeps the translated layout coordinates (FloatingTextBoxEmissionMap).
        bool uniformMap = FloatingTextBoxEmissionMap.TryCreate(markupContext, pageHeight, out FloatingTextBoxEmissionMap map);
        graphics.SaveState();
        if (uniformMap)
        {
            (double clipX, double clipY, double clipWidth, double clipHeight) = map.MapClipRectangle(contentX, contentTop, contentWidth, contentHeight);
            graphics.ClipRectangle(clipX, clipY, clipWidth, clipHeight);
        }
        else
        {
            graphics.ClipRectangle(contentX, contentTop - contentHeight, contentWidth, contentHeight);
        }
        // RV05 floatbox probe (Word COM reference edge-floatbox): stroked frames paint
        // their outline under the text. Uniform-map pages keep legacy borderless output
        // until scaled stroking is probed.
        if (!uniformMap && TryResolveTextBoxStroke(drawing.Drawing, width, height, out double strokeWidth, out byte strokeRed, out byte strokeGreen, out byte strokeBlue))
        {
            graphics.SetStrokeRgb(strokeRed, strokeGreen, strokeBlue);
            graphics.SetLineWidth(strokeWidth);
            graphics.StrokeRectangle(placedX, placedTop - height, width, height);
        }

        IReadOnlyList<DocxLayoutItem> items = textBoxLayout.TextLines
            .Select(line => uniformMap
                ? map.PrecompensateLine(line, contentX, contentTop)
                : TranslateTextLine(line, contentX, contentTop))
            .Cast<DocxLayoutItem>()
            .Concat(textBoxLayout.InlineImages.Select(image => uniformMap
                ? map.PrecompensateImage(image, contentX, contentTop, drawing.AnchorPageIndex ?? image.PageIndex)
                : image.Relocate(contentX + image.X, contentTop + image.Y, drawing.AnchorPageIndex)))
            .Concat(textBoxLayout.TableRows.Select(row => uniformMap
                ? map.PrecompensateRow(DocxLayoutEngine.ShiftTableRow(row, contentTop, contentX))
                : DocxLayoutEngine.ShiftTableRow(row, contentTop, contentX)))
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
        if (box.TextLines.Count == 0 &&
            box.InlineImages.Count == 0 &&
            box.TableRows.Count == 0)
        {
            return;
        }

        // Body boxes carry the body emission offset. Mapped static boxes already
        // own physical Y; their children consume that same frame when painting.
        double clipX = box.BoxX + ResolveTextEmissionXOffset(markupContext);
        double clipTop = box.BoxTop - (box.StaticStoryUsesPhysicalY ? 0d : ResolveTextEmissionBaselineOffset(markupContext));
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
        return IsOnOffTrue(drawing.BehindDocumentValue);

        bool IsOnOffTrue(string? value)
        {
            return value is not null &&
                (value.Length == 0 ||
                value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("on", StringComparison.OrdinalIgnoreCase));
        }
    }

    private static long ReadZOrder(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long zOrder)
            ? zOrder
            : 0L;
    }
}
