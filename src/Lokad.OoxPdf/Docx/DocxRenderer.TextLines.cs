using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxRenderer
{
    internal static IEnumerable<DocxTextLineLayout> EnumerateBodyTextLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            switch (item)
            {
                case DocxTextLineLayout line:
                    yield return line;
                    break;
                case DocxTableRowLayout row:
                    foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                    {
                        yield return cellLine;
                    }

                    break;
            }
        }
    }

    internal static (double? BaselineY, double SourceBaselineOffset) ResolveBodyPrintOrigin(DocxLayoutPage page)
    {
        double? baseline = null;
        double sourceOffset = 0d;
        foreach (var line in EnumerateBodyTextLines(page))
        {
            double anchor = line.BaselineY;
            double offset = line.SourceBaselineOffset;
            if (line.InlineFrame is { TextBaselineInset: { } textInset } frame)
            {
                // Object-only lines own their occupied top. Text/math lines
                // retain the independently measured typographic anchor.
                bool hasTextOrMath = DocxInlineLineGeometry.HasTextOrMath(line);
                anchor += frame.BaselineInset - (hasTextOrMath ? textInset : 0d);
            }
            if (baseline is null || anchor > baseline.Value)
                (baseline, sourceOffset) = (anchor, offset);
        }
        // Leading standalone objects have their own measured frame origin. A later
        // caption cannot supply the source origin for already scaled content.
        double? objectTop = page.Items.FirstOrDefault() switch
        {
            DocxEquationLayout equation => equation.BaselineY + equation.Math.BaselineInset * equation.Scale,
            DocxInlineImageLayout image => image.Y + image.Height,
            DocxTableRowLayout row when !EnumerateTableRowTextLines(row).Any() =>
                row.Cells.SelectMany(cell => cell.InlineImages).Select(image => (double?)(image.Y + image.Height)).Max(),
            _ => null
        };
        return objectTop is { } top && (baseline is null || top > baseline.Value)
            ? (top, 0d) : (baseline, sourceOffset);

    }

    private static IEnumerable<DocxTextLineLayout> EnumerateRenderedPageTextLines(
        FloatingDrawingPageIndex.PageIndexPair drawingPages,
        DocxLayoutPage page,
        int pageIndex,
        DocxMarkupContext markupContext,
        double pageHeight)
    {
        return EnumerateRenderedPageLineSources(page, drawingPages.PageAll(pageIndex), markupContext, pageHeight)
            .Select(source => source.Line);
    }

    private readonly record struct LinkPaintViewport(PdfRectangle Bounds, bool PreserveFittingPadding = false);

    // Every mapped line retains its placement container. A cell/textbox's local
    // story kind cannot replace the header/footer/note that owns its print frame.
    private static IEnumerable<(DocxTextLineLayout Line, DocxStoryId Owner, LinkPaintViewport? Viewport)> EnumerateRenderedPageLineSources(
        DocxLayoutPage page, IEnumerable<DocxFloatingDrawingLayout> floatingDrawings,
        DocxMarkupContext markupContext, double pageHeight, IReadOnlyList<DocxLayoutItem>? staticPaintItems = null,
        IReadOnlyList<PlacedNoteTableFrame?>? noteTables = null, CancellationToken token = default)
    {
        FloatingTextBoxEmissionMap? map = TryCreateFloatingTextBoxEmissionMap(markupContext, pageHeight);
        bool clipCellLinks = staticPaintItems is not null;
        foreach (var source in EnumerateStaticLineSources(page, true, markupContext, includeTableRows: !clipCellLinks))
            yield return (source.Line, source.Owner, null);
        if (staticPaintItems is not null)
        {
            foreach (var source in EnumerateInlineTextBoxTableLinkLines(staticPaintItems, markupContext, token))
                yield return source;
            foreach (var (line, clip, owner) in EnumerateLinkLinesInItems(staticPaintItems, null, markupContext, token, includeDirectLines: false))
                yield return (line, owner ?? DocxStoryId.Body(), clip is { } bounds ? new(bounds, PreserveFittingPadding: true) : null);
            foreach (var (line, clip, _) in EnumerateLinkLinesInItems(page.Items, null, markupContext, token))
                yield return (line, DocxStoryId.Body(), clip is { } bounds ? new(bounds, PreserveFittingPadding: true) : null);
        }
        else
        {
            foreach (var line in EnumerateBodyTextLines(page)) yield return (line, DocxStoryId.Body(), null);
        }
        for (int index = 0; index < page.PlacedRelatedStories.Count; index++)
        {
            var story = page.PlacedRelatedStories[index];
            var owner = DocxStoryId.Related(story.StoryLayout.Story.Kind);
            var noteTable = noteTables?[index];
            foreach (var line in EnumerateMappedPlacedNoteStoryTextLines(story, map, includeTableRows: noteTable is null))
                yield return (line, owner, null);
            if (noteTable is { } frame)
            {
                foreach (var (line, clip, _) in EnumerateLinkLinesInItems(frame.Items, frame.Clip, markupContext, token, includeDirectLines: false))
                    yield return (line, owner, new(clip ?? frame.Clip, PreserveFittingPadding: true));
            }
        }
        foreach (var line in EnumerateInlineTextBoxTextLines(page, includeTableRows: !clipCellLinks))
            yield return (line, DocxStoryId.Body(), null);
        if (clipCellLinks)
            foreach (var source in EnumerateInlineTextBoxTableLinkLines(page.Items, markupContext, token))
                yield return source;
        foreach (var drawing in floatingDrawings)
            foreach (var (line, clip) in EnumerateMappedFloatingDrawingTextBoxTextLines([drawing], markupContext, pageHeight,
                clipCellLinks: clipCellLinks, token: token))
                yield return (line, drawing.Story ?? DocxStoryId.Body(), new(clip));
        foreach (var story in page.PlacedRelatedStories)
        {
            if (story.FloatingDrawings.Count == 0) continue;
            var (_, _, storyClip, _) = ResolvePlacedStoryPaintFrame(story, markupContext, pageHeight);
            foreach (var (line, clip) in EnumerateMappedFloatingDrawingTextBoxTextLines(story.FloatingDrawings, markupContext,
                pageHeight, storyClip, clipCellLinks, token))
                yield return (line, DocxStoryId.Related(story.StoryLayout.Story.Kind), new(clip));
        }
    }

    private static IEnumerable<(DocxTextLineLayout Line, DocxStoryId Owner, LinkPaintViewport? Viewport)> EnumerateInlineTextBoxTableLinkLines(
        IReadOnlyList<DocxLayoutItem> items, DocxMarkupContext context, CancellationToken token)
    {
        foreach (var item in items)
        {
            token.ThrowIfCancellationRequested();
            if (item is not DocxInlineTextBoxLayout { TableRows.Count: > 0 } box) continue;
            var (x, top) = ResolveInlineTextBoxPaintOrigin(box, context);
            var viewport = new PdfRectangle(x, top - box.BoxHeight, box.BoxWidth, box.BoxHeight);
            var boxItems = GetInlineTextBoxPaintItems(box, context);
            foreach (var (line, clip, _) in EnumerateLinkLinesInItems(boxItems, viewport, context, token, includeDirectLines: false))
                yield return (line, box.Story ?? DocxStoryId.Body(), new(clip ?? viewport, PreserveFittingPadding: true));
        }
    }

    private static FloatingTextBoxEmissionMap? TryCreateFloatingTextBoxEmissionMap(DocxMarkupContext markupContext, double pageHeight)
    {
        return FloatingTextBoxEmissionMap.TryCreate(markupContext, pageHeight, out FloatingTextBoxEmissionMap map)
            ? map
            : null;
    }

    // Mapped legs translate by the inset content origin (like the renderer) and
    // pre-compensate design coordinates into body-layout space on scaled WC pages; every
    // other mode keeps the legacy frame-origin translation bit-identically.
    // Box-only walk for comment suppression: unlike EnumerateTableRowTextLines it yields
    // no regular cell lines, so body-cell comments keep ballooning while box-anchored
    // ones suppress by paragraph identity.
    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextBoxLines(DocxLayoutPage page)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            if (item is not DocxTableRowLayout row)
            {
                continue;
            }

            foreach (DocxTextLineLayout line in EnumerateTableRowTextBoxLines(row))
            {
                yield return line;
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextBoxLines(DocxTableRowLayout row)
    {
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxInlineTextBoxLayout box in cell.InlineTextBoxes)
            {
                foreach (DocxTextLineLayout boxLine in box.TextLines)
                {
                    yield return boxLine;
                }

                foreach (DocxTableRowLayout boxRow in box.TableRows)
                {
                    foreach (DocxTextLineLayout boxCellLine in EnumerateTableRowTextLines(boxRow))
                    {
                        yield return boxCellLine;
                    }

                    foreach (DocxTextLineLayout nestedBoxLine in EnumerateTableRowTextBoxLines(boxRow))
                    {
                        yield return nestedBoxLine;
                    }
                }
            }

            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTextLineLayout nestedBoxLine in EnumerateTableRowTextBoxLines(nestedRow))
                {
                    yield return nestedBoxLine;
                }
            }
        }
    }
    private static IEnumerable<(DocxTextLineLayout Line, PdfRectangle Clip)> EnumerateMappedFloatingDrawingTextBoxTextLines(
        IEnumerable<DocxFloatingDrawingLayout> drawings,
        DocxMarkupContext context, double pageHeight, PdfRectangle? storyClip = null,
        bool clipCellLinks = false, CancellationToken token = default)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            if (drawing.PlacedX is not { } placedX ||
                drawing.PlacedTop is not { } placedTop ||
                drawing.ExtentWidthPoints is not { } width || drawing.ExtentHeightPoints is not { } height ||
                drawing.TextBoxLayout is not { } textBoxLayout)
            {
                continue;
            }

            var frame = ResolveFloatingTextBoxPaintFrame(drawing.Drawing, placedX, placedTop, width, height, context, pageHeight);
            var cellViewport = ResolveFloatingCellViewport(drawing);
            var viewport = (cellViewport ?? storyClip) is { } parent
                ? IntersectPaintViewports(parent, frame.Clip) : frame.Clip;
            DocxTableCellClipSnapshot? clip = cellViewport is { } cell ? new(true, cell.X, cell.Y, cell.Width, cell.Height) : null;

            var paintItems = clipCellLinks && textBoxLayout.TableRows.Count != 0
                ? PlaceFloatingTextBoxItems(textBoxLayout, frame, drawing.AnchorPageIndex) : null;
            var lines = paintItems is null ? textBoxLayout.TextLines.Select(frame.PlaceLine) : paintItems.OfType<DocxTextLineLayout>();
            foreach (var placed in lines)
            {
                // A line can cross the cell clip while its linked fragment
                // stays visible. Link admission intersects the fragment with
                // the paint viewport after resolving its text or math bounds.
                if (clipCellLinks || InsideViewport(placed)) yield return (placed, viewport);
            }

            if (paintItems is not null)
            {
                foreach (var (line, tableClip, _) in EnumerateLinkLinesInItems(paintItems, viewport, context, token, includeDirectLines: false))
                    yield return (line, tableClip ?? viewport);
                continue;
            }

            foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
            {
                foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                {
                    var placed = frame.PlaceLine(cellLine);
                    if (InsideViewport(placed)) yield return (placed, viewport);
                }
            }

            bool InsideViewport(DocxTextLineLayout line)
            {
                if (clip is null) return true;
                if (!DocxInlineLineGeometry.TryGetOccupiedBounds(line, out double bottom, out double top))
                    bottom = top = line.BaselineY;
                double yOffset = frame.Map?.YOffset ?? 0d;
                return DocxTableCellClipGeometry.ContainsHorizontalRange(clip, line.X + (frame.Map?.XOffset ?? 0d), line.Width) &&
                    DocxTableCellClipGeometry.ContainsVerticalRange(clip, bottom - yOffset, top - yOffset);
            }
        }
    }

    // Link sources use the same placed rows, fragment admission and cell
    // viewports as paint. Keep every ancestor clip for nested content.
    private static IEnumerable<(DocxTextLineLayout Line, PdfRectangle? Clip, DocxStoryId? Owner)> EnumerateLinkLinesInItems(
        IReadOnlyList<DocxLayoutItem> items, PdfRectangle? viewport, DocxMarkupContext context, CancellationToken token,
        bool includeDirectLines = true)
    {
        double xOffset = ResolveTextEmissionXOffset(context);
        double yOffset = ResolveTextEmissionBaselineOffset(context);
        return Enumerate(items, viewport, includeDirectLines, null);

        IEnumerable<(DocxTextLineLayout Line, PdfRectangle? Clip, DocxStoryId? Owner)> Enumerate(
            IReadOnlyList<DocxLayoutItem> current, PdfRectangle? parent, bool includeLines, DocxStoryId? owner)
        {
            for (int index = 0; index < current.Count; index++)
            {
                token.ThrowIfCancellationRequested();
                if (includeLines && current[index] is DocxTextLineLayout direct)
                {
                    yield return (direct, parent, owner ?? direct.Story);
                    continue;
                }
                if (current[index] is not DocxTableRowLayout row) continue;
                var rowOwner = owner ?? row.Story;
                var geometry = ResolveTablePaintGeometry(row, xOffset, yOffset);
                for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
                {
                    token.ThrowIfCancellationRequested();
                    var cell = row.Cells[cellIndex];
                    if (!ShouldRenderTableCellVisualFragment(cell)) continue;
                    var clip = DocxTableCellClipGeometry.Resolve(geometry, geometry.Cells[cellIndex]);
                    var cellFrame = new PdfRectangle(clip.X, clip.Y, clip.Width, clip.Height);
                    var cellViewport = parent is { } enclosing ? IntersectPaintViewports(enclosing, cellFrame) : cellFrame;
                    foreach (var line in cell.TextLines) yield return (line, cellViewport, rowOwner);
                    foreach (var nested in Enumerate(cell.NestedRows, cellViewport, includeLines: false, rowOwner)) yield return nested;
                    foreach (var box in cell.InlineTextBoxes)
                    {
                        token.ThrowIfCancellationRequested();
                        var (x, top) = ResolveInlineTextBoxPaintOrigin(box, context);
                        var boxViewport = IntersectPaintViewports(cellViewport,
                            new PdfRectangle(x, top - box.BoxHeight, box.BoxWidth, box.BoxHeight));
                        var boxItems = GetInlineTextBoxPaintItems(box, context);
                        foreach (var line in boxItems.OfType<DocxTextLineLayout>()) yield return (line, boxViewport, rowOwner);
                        foreach (var nested in Enumerate(boxItems, boxViewport, includeLines: false, rowOwner)) yield return nested;
                    }
                }
            }
        }
    }

    // Office A/B (w6-tbxctl probe): Word never balloons body-flow floating-textbox
    // comments; static header/footer floatings keep the legacy path (unprobed).
    internal static bool IsStaticStoryFloatingDrawing(DocxFloatingDrawingLayout drawing)
    {
        return drawing.Story?.Kind is DocxStoryKind.Header or DocxStoryKind.Footer;
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateMarkupBalloonAnchorTextLines(
        DocxLayoutPage page,
        IReadOnlyList<DocxFloatingDrawingLayout> floatingDrawings,
        DocxMarkupContext markupContext,
        double pageHeight)
    {
        return EnumerateRenderedPageLineSources(page, floatingDrawings, markupContext, pageHeight)
            .Select(source => source.Line);
    }

    // RV06 anchor probe (edge-footanchor-5, Word 16.0): footnote and endnote stories
    // lay out unscaled in design space, so balloon anchors from their lines map
    // uniformly like their emission; other stories keep legacy coordinates.
    private static IEnumerable<DocxTextLineLayout> EnumerateMappedPlacedNoteStoryTextLines(
        DocxLayoutPage page, FloatingTextBoxEmissionMap? map) =>
        page.PlacedRelatedStories.SelectMany(story => EnumerateMappedPlacedNoteStoryTextLines(story, map));

    private static IEnumerable<DocxTextLineLayout> EnumerateMappedPlacedNoteStoryTextLines(
        DocxPlacedRelatedStoryLayout story, FloatingTextBoxEmissionMap? map, bool includeTableRows = true)
    {
        FloatingTextBoxEmissionMap? storyMap = null;
        if (map is { } availableMap &&
            story.StoryLayout.Story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote)
            storyMap = availableMap;
        foreach (DocxTextLineLayout line in story.TextLines)
            yield return storyMap is { } lineMap ? lineMap.PrecompensateLine(line, 0d, 0d) : line;
        if (!includeTableRows) yield break;
        foreach (DocxTableRowLayout row in story.TableRows)
            foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                yield return storyMap is { } cellMap ? cellMap.PrecompensateLine(cellLine, 0d, 0d) : cellLine;
    }

    // Inline-textbox content lives in absolute flow space (placed at layout time), so
    // anchor/snapshot/link consumers take box lines directly with the body offsets.
    private static IEnumerable<DocxTextLineLayout> EnumerateInlineTextBoxTextLines(DocxLayoutPage page, bool includeTableRows = true)
    {
        foreach (DocxLayoutItem item in page.Items)
        {
            if (item is not DocxInlineTextBoxLayout box)
            {
                continue;
            }

            foreach (DocxTextLineLayout line in box.TextLines)
            {
                yield return line;
            }

            if (!includeTableRows) continue;
            foreach (DocxTableRowLayout row in box.TableRows)
            {
                foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
                {
                    yield return cellLine;
                }
            }
        }
    }
    private sealed record DocxTextEmissionLineSource(
        DocxTextLineLayout Line,
        bool IsStaticStory,
        string StoryKind,
        string? StoryVariantType,
        string? ContainerStoryKind,
        string? ContainerStoryVariantType);

    private static IEnumerable<DocxTextEmissionLineSource> EnumerateRenderedFloatingDrawingTextBoxTextLines(
        FloatingDrawingPageIndex.PageIndexPair drawingPages,
        DocxLayoutPage page,
        int pageIndex,
        DocxMarkupContext markupContext,
        double pageHeight)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawingPages.PageAll(pageIndex))
        {
            bool isStaticStory = IsStaticStoryFloatingDrawing(drawing);
            foreach (var (line, _) in EnumerateMappedFloatingDrawingTextBoxTextLines([drawing], markupContext, pageHeight))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    isStaticStory,
                    "TextBox",
                    line.Story?.VariantType,
                    drawing.Story?.ToKindString() ?? "Body",
                    drawing.Story?.VariantType);
            }
        }

        foreach (DocxPlacedRelatedStoryLayout story in page.PlacedRelatedStories)
        {
            foreach (var (line, _) in EnumerateMappedFloatingDrawingTextBoxTextLines(story.FloatingDrawings, markupContext, pageHeight))
            {
                yield return new DocxTextEmissionLineSource(
                    line,
                    IsStaticStory: false,
                    "TextBox",
                    line.Story?.VariantType,
                    story.StoryLayout.Story.Kind.ToValueString(),
                    story.StoryLayout.Story.Id);
            }
        }

        foreach (DocxTextLineLayout line in EnumerateInlineTextBoxTextLines(page))
        {
            yield return new DocxTextEmissionLineSource(
                line,
                IsStaticStory: false,
                "TextBox",
                line.Story?.VariantType,
                "Body",
                null);
        }
    }

    private static IEnumerable<DocxLayoutItem> EnumerateStaticLayoutItems(DocxLayoutPage page, DocxMarkupContext markupContext)
    {
        return page.StaticTextLines.Cast<DocxLayoutItem>()
            .Concat(page.StaticInlineImages)
            .Concat(page.StaticInlineTextBoxes)
            .Concat(page.StaticTableRows)
            .Select(item => AdjustStaticLayoutItem(item, markupContext))
            .OrderByDescending(item => item switch
            {
                DocxTextLineLayout textLine => textLine.BaselineY,
                DocxInlineImageLayout image => image.Y + image.Height,
                DocxTableRowLayout row => row.Y + row.Height,
                _ => 0d
            });
    }

    // Physical static frames already own their Y. Precompensate the existing
    // text/table painters once, regardless of which container supplies the item.
    private static DocxLayoutItem AdjustStaticLayoutItem(DocxLayoutItem item, DocxMarkupContext markupContext) => item switch
    {
        DocxTextLineLayout line => AdjustStaticTextLine(line, markupContext),
        DocxTableRowLayout row => AdjustStaticTableRow(row, markupContext),
        _ => item
    };

    private static DocxTableRowLayout AdjustStaticTableRow(DocxTableRowLayout row, DocxMarkupContext markupContext) =>
        row.StaticStoryUsesPhysicalY ? PrecompensateStaticTableRow(row, markupContext.WordCompatibleTextYOffset) : row;

    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextLines(DocxLayoutPage page, DocxMarkupContext? markupContext = null)
    {
        return EnumerateStaticTextLines(page, includeTextBoxes: true, markupContext);
    }

    private static DocxTextLineLayout AdjustStaticTextLine(DocxTextLineLayout line, DocxMarkupContext? markupContext)
    {
        double adjustment = line.StaticStoryUsesPhysicalY
            ? markupContext?.WordCompatibleTextYOffset ?? 0d
            : markupContext?.WordCompatibleStaticTextYAdjustmentPoints ?? 0d;
        return line.Translate(0d, adjustment);
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextLines(DocxLayoutPage page, bool includeTextBoxes, DocxMarkupContext? markupContext = null) =>
        EnumerateStaticLineSources(page, includeTextBoxes, markupContext).Select(source => source.Line);

    private static IEnumerable<(DocxTextLineLayout Line, DocxStoryId Owner)> EnumerateStaticLineSources(
        DocxLayoutPage page, bool includeTextBoxes, DocxMarkupContext? markupContext, bool includeTableRows = true)
    {
        foreach (DocxTextLineLayout line in page.StaticTextLines)
            yield return (AdjustStaticTextLine(line, markupContext), line.Story ?? DocxStoryId.Body());
        if (includeTextBoxes)
            foreach (var box in page.StaticInlineTextBoxes)
            {
                var owner = box.Story ?? DocxStoryId.Body();
                foreach (var line in box.TextLines) yield return (AdjustStaticTextLine(line, markupContext), owner);
                if (!includeTableRows) continue;
                foreach (var row in box.TableRows)
                    foreach (var line in EnumerateTableRowTextLines(row)) yield return (AdjustStaticTextLine(line, markupContext), owner);
            }
        if (!includeTableRows) yield break;
        foreach (DocxTableRowLayout row in page.StaticTableRows)
            foreach (DocxTextLineLayout line in EnumerateTableRowTextLines(row))
                yield return (line.StaticStoryUsesPhysicalY ? AdjustStaticTextLine(line, markupContext) : line,
                    row.Story ?? DocxStoryId.Body());
    }

    // Box-only static walk for comment suppression: static text comments keep
    // ballooning while static-box ones suppress by paragraph identity.
    private static IEnumerable<DocxTextLineLayout> EnumerateStaticTextBoxLines(DocxLayoutPage page)
    {
        foreach (DocxInlineTextBoxLayout box in page.StaticInlineTextBoxes)
        {
            foreach (DocxTextLineLayout boxLine in box.TextLines)
            {
                yield return boxLine;
            }

            foreach (DocxTableRowLayout boxRow in box.TableRows)
            {
                foreach (DocxTextLineLayout boxCellLine in EnumerateTableRowTextLines(boxRow))
                {
                    yield return boxCellLine;
                }
            }
        }
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateFloatingDrawingTextBoxTextLines(DocxFloatingDrawingLayout drawing)
    {
        if (drawing.PlacedX is not { } placedX ||
            drawing.PlacedTop is not { } placedTop ||
            drawing.TextBoxLayout is not { } textBoxLayout)
        {
            yield break;
        }

        foreach (DocxTextLineLayout line in textBoxLayout.TextLines)
        {
            yield return line.Translate(placedX, placedTop);
        }

        foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
        {
            foreach (DocxTextLineLayout cellLine in EnumerateTableRowTextLines(row))
            {
                yield return cellLine.Translate(placedX, placedTop);
            }
        }
    }

    private static IEnumerable<DocxTableRowLayout> EnumerateFloatingDrawingTextBoxTableRows(IEnumerable<DocxFloatingDrawingLayout> drawings)
    {
        foreach (DocxFloatingDrawingLayout drawing in drawings)
        {
            if (drawing.PlacedX is not { } placedX ||
                drawing.PlacedTop is not { } placedTop ||
                drawing.TextBoxLayout is not { } textBoxLayout)
            {
                continue;
            }

            foreach (DocxTableRowLayout row in textBoxLayout.TableRows)
            {
                foreach (DocxTableRowLayout nested in EnumerateTableRows(DocxLayoutEngine.ShiftTableRow(row, placedTop, placedX)))
                {
                    yield return nested;
                }
            }
        }
    }

    // Static rows already own physical geometry. Precompensate only consumers
    // of the body paint offset; inline boxes perform this at their own boundary.
    private static DocxTableRowLayout PrecompensateStaticTableRow(DocxTableRowLayout row, double deltaY)
    {
        return row with
        {
            Y = row.Y + deltaY,
            Cells = row.Cells
                .Select(cell => PrecompensateStaticTableCell(cell, deltaY))
                .ToArray()
        };
    }

    private static DocxTableCellLayout PrecompensateStaticTableCell(DocxTableCellLayout cell, double deltaY)
    {
        return cell with
        {
            Y = cell.Y + deltaY,
            TextLines = cell.TextLines
                .Select(line => line.Translate(0d, deltaY))
                .ToArray(),
            InlineImages = cell.InlineImages
                .Select(image => image.Translate(0d, deltaY))
                .ToArray(),
            NestedTableRows = cell.NestedRows
                .Select(row => PrecompensateStaticTableRow(row, deltaY))
                .ToArray()
        };
    }

    private static IEnumerable<DocxTextLineLayout> EnumerateTableRowTextLines(DocxTableRowLayout row)
    {
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            foreach (DocxTextLineLayout line in cell.TextLines)
            {
                yield return line;
            }

            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxTextLineLayout nestedLine in EnumerateTableRowTextLines(nestedRow))
                {
                    yield return nestedLine;
                }
            }

            foreach (DocxInlineTextBoxLayout box in cell.InlineTextBoxes)
            {
                foreach (DocxTextLineLayout boxLine in box.TextLines)
                {
                    yield return boxLine;
                }

                foreach (DocxTableRowLayout boxRow in box.TableRows)
                {
                    foreach (DocxTextLineLayout boxCellLine in EnumerateTableRowTextLines(boxRow))
                    {
                        yield return boxCellLine;
                    }
                }
            }
        }
    }

    private static DocxTextEmissionLineSnapshot ToTextEmissionLineSnapshot(
        int pageIndex,
        bool isStaticStory,
        string storyKind,
        string? storyVariantType,
        string? containerStoryKind,
        string? containerStoryVariantType,
        DocxTextLineLayout line,
        DocxFontResources fontResources,
        DocxPageNumber pageNumber,
        int pageCount,
        double fontScale,
        double baselineOffsetY,
        double xOffset,
        bool useWordCompatibleTextProfile)
    {
        DocxTextEmissionSegmentSnapshot[] segments = CreateTextEmissionSegments(line, fontResources, pageNumber, pageCount, fontScale, baselineOffsetY, xOffset, useWordCompatibleTextProfile)
            .Select(segment => ToTextEmissionSegmentSnapshot(segment, line))
            .ToArray();
        return new DocxTextEmissionLineSnapshot(
            pageIndex,
            isStaticStory,
            storyKind,
            storyVariantType,
            containerStoryKind,
            containerStoryVariantType,
            line.SourceBlockIndex,
            line.SourceParagraphIndex,
            line.SourceLineIndex,
            line.EndsWithIntraTokenBreak,
            segments.Length,
            segments.Sum(segment => segment.TextLength),
            segments.Count(segment => segment.RevisionKind is not null),
            segments.Count(segment => segment.RevisionKind == "Insertion"),
            segments.Count(segment => segment.RevisionKind == "Deletion"),
            segments.Count(segment => segment.RevisionKind == "MoveFrom"),
            segments.Count(segment => segment.RevisionKind == "MoveTo"),
            segments.Count(segment => segment.RevisionKind is not null &&
                segment.RevisionKind != "Insertion" &&
                segment.RevisionKind != "Deletion" &&
                segment.RevisionKind != "MoveFrom" &&
                segment.RevisionKind != "MoveTo"),
            line.SourceParagraph?.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment) ?? 0,
            segments.Count(segment => segment.IsTerminalLineSpace),
            segments.Count(segment => Math.Abs(segment.PdfCharacterSpacing) > 0.0001d),
            segments);
    }

    private static DocxTextEmissionSegmentSnapshot ToTextEmissionSegmentSnapshot(
        DocxTextEmissionSegment segment,
        DocxTextLineLayout line)
    {
        DocxTextEmissionPlan plan = DocxTextEmissionPlanner.CreateForEmissionSegment(segment);
        return new DocxTextEmissionSegmentSnapshot(
            segment.Text.Length,
            line.SourceBlockIndex,
            line.SourceParagraphIndex,
            line.SourceLineIndex,
            segment.Role.ToString(),
            segment.X,
            segment.BaselineY,
            segment.Width,
            segment.FontSize,
            plan.PdfFontSize,
            segment.StyleRun.EffectiveProperties.CharacterSpacingPoints,
            plan.PdfCharacterSpacing,
            plan.PdfCharacterSpacingSource.ToString(),
            plan.PositioningCharacterSpacing,
            plan.CompensatePdfCharacterSpacing,
            DocxTextEmissionPlanner.ClassifyText(segment.Text),
            DocxTextEmissionPlanner.MeasureAdvanceProfile(segment.Text, segment.Resource?.Embedded, segment.FallbackFace?.Font, segment.Width, plan),
            DocxTextEmissionPlanner.CreateGlyphAdvanceSignature(segment.Text, segment.Resource?.Embedded, segment.FallbackFace?.Font, plan.PairKerningEnabled),
            segment.IsTerminalLineSpace,
            segment.Resource?.Name ?? segment.FallbackFace?.ResourceName ?? string.Empty,
            segment.SyntheticBold,
            segment.SyntheticItalic,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleId,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleFound,
            segment.StyleRun.EffectiveProperties.StyleResolution.CharacterStyleDepth,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasDocumentDefaultRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasParagraphStyleRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasCharacterStyleRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasDirectRunProperties,
            segment.StyleRun.EffectiveProperties.StyleResolution.HasTableStyleRunProperties,
            segment.StyleRun.Revision?.Kind.ToValueString(),
            segment.StyleRun.Revision?.SourceElement);
    }
}
