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
    private static void AddPlacedTableBookmarkDestinations(
        DocxLayoutPage page, IReadOnlyList<DocxFloatingDrawingLayout> pageDrawings, int pageIndex, DocxMarkupContext context,
        Dictionary<string, PdfLinkDestination> destinations, CancellationToken token)
    {
        double xOffset = ResolveTextEmissionXOffset(context);
        double yOffset = ResolveTextEmissionBaselineOffset(context);
        var pageClip = new DocxTableCellClipSnapshot(true, 0d, 0d, page.Width, page.Height);
        foreach (var item in page.Items)
        {
            token.ThrowIfCancellationRequested();
            if (item is DocxTableRowLayout row && HasTarget(row)) AddTargets(row, null);
            if (item is DocxInlineTextBoxLayout box) AddBoxTargets(box, pageClip);
        }
        foreach (var row in page.StaticTableRows)
            if (HasTarget(row)) AddTargets(AdjustStaticTableRow(row, context), pageClip);
        foreach (var box in page.StaticInlineTextBoxes)
            AddBoxTargets(box, pageClip);
        foreach (var drawing in pageDrawings)
            AddFloatingBoxTargets(drawing, pageClip);
        foreach (var story in page.PlacedRelatedStories)
        {
            token.ThrowIfCancellationRequested();
            bool hasNoteRows = story.TableRows.Count > 0 && story.StoryLayout.Story.IsNormalStoryType &&
                story.StoryLayout.Story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote;
            if (!hasNoteRows && story.FloatingDrawings.Count == 0)
                continue;
            var (map, boxViewport, clip) = ResolvePlacedStoryPaintFrame(story, context, page.Height);
            if (hasNoteRows)
            {
                var storyClip = new DocxTableCellClipSnapshot(true, clip.X, clip.Y, clip.Width, clip.Height);
                foreach (var row in story.TableRows)
                    if (HasTarget(row))
                        AddTargets(ResolvePlacedStoryTablePaintFrame(story, row, map, token), storyClip);
            }
            if (story.FloatingDrawings.Count > 0)
            {
                var storyClip = DocxTableCellClipGeometry.Intersect(
                    new(true, boxViewport.X, boxViewport.Y, boxViewport.Width, boxViewport.Height), pageClip);
                foreach (var drawing in story.FloatingDrawings)
                    AddFloatingBoxTargets(drawing, storyClip);
            }
        }

        bool HasTarget(DocxTableRowLayout row)
        {
            token.ThrowIfCancellationRequested();
            return row.BookmarkAnchors is { Count: > 0 } ||
                row.Cells.Any(cell => cell.NestedRows.Any(HasTarget) ||
                    cell.InlineTextBoxes.Any(box => box.TableRows.Any(HasTarget)));
        }

        void AddBoxTargets(DocxInlineTextBoxLayout box, DocxTableCellClipSnapshot parentClip)
        {
            token.ThrowIfCancellationRequested();
            DocxTableCellClipSnapshot? clip = null;
            foreach (var row in box.TableRows)
            {
                if (!HasTarget(row)) continue;
                if (clip is null)
                {
                    var (x, top) = ResolveInlineTextBoxPaintOrigin(box, context);
                    clip = DocxTableCellClipGeometry.Intersect(
                        new(true, x, top - box.BoxHeight, box.BoxWidth, box.BoxHeight), parentClip);
                }
                AddTargets(AdjustStaticTableRow(row, context), clip);
            }
        }

        void AddFloatingBoxTargets(DocxFloatingDrawingLayout drawing, DocxTableCellClipSnapshot parentClip)
        {
            token.ThrowIfCancellationRequested();
            if (drawing.TextBoxLayout is not { TableRows.Count: > 0 } textBox ||
                drawing.PlacedX is not { } placedX || drawing.PlacedTop is not { } placedTop ||
                drawing.ExtentWidthPoints is not { } width || drawing.ExtentHeightPoints is not { } height)
                return;

            FloatingTextBoxPaintFrame? frame = null;
            DocxTableCellClipSnapshot? clip = null;
            if (ResolveFloatingCellViewport(drawing) is { } cell)
                parentClip = DocxTableCellClipGeometry.Intersect(
                    new(true, cell.X, cell.Y, cell.Width, cell.Height), parentClip);
            foreach (var row in textBox.TableRows)
            {
                if (!HasTarget(row)) continue;
                if (frame is null)
                {
                    frame = ResolveFloatingTextBoxPaintFrame(drawing.Drawing, placedX, placedTop, width, height, context, page.Height);
                    PdfRectangle viewport = frame.Value.Clip;
                    clip = DocxTableCellClipGeometry.Intersect(
                        new(true, viewport.X, viewport.Y, viewport.Width, viewport.Height), parentClip);
                }
                AddTargets(frame.Value.PlaceRow(row), clip);
            }
        }

        void AddTargets(DocxTableRowLayout row, DocxTableCellClipSnapshot? parentClip)
        {
            token.ThrowIfCancellationRequested();
            if (row.Cells.Count == 0 || parentClip is { Width: <= 0d } or { Height: <= 0d })
                return;
            if (row.FragmentIndex == 0 && row.BookmarkAnchors is { } bookmarks)
            {
                // Use the same precompensated row origin and paint offsets as
                // the cell fill. A nested target also needs its ancestor clip.
                double x = row.Cells[0].X + xOffset;
                double top = row.Y + row.Height - yOffset;
                if (DocxTableCellClipGeometry.ContainsHorizontalRange(parentClip, x, 0d) &&
                    DocxTableCellClipGeometry.ContainsVerticalRange(parentClip, top, top))
                    foreach (var bookmark in bookmarks)
                    {
                        token.ThrowIfCancellationRequested();
                        if (!string.IsNullOrEmpty(bookmark.Name))
                            destinations.TryAdd(bookmark.Name, new PdfLinkDestination(pageIndex, x, top, Zoom: null));
                    }
            }

            DocxTableRowLayout? geometryRow = null;
            for (int i = 0; i < row.Cells.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                var cell = row.Cells[i];
                if ((cell.NestedRows.Count == 0 && cell.InlineTextBoxes.Count == 0) ||
                    cell.VisualOwnership != DocxTableCellVisualOwnership.OwnCell)
                    continue;
                geometryRow ??= ResolveTablePaintGeometry(row, xOffset, yOffset);
                var geometryCell = geometryRow.Cells[i];
                var viewport = parentClip ?? pageClip;
                if (cell.Cell.HasVerticalMerge)
                {
                    // Borrowed rows never own a destination. For a restart,
                    // descend only when its full painted frame fits the viewport.
                    viewport = DocxTableCellClipGeometry.Intersect(viewport, pageClip);
                    if (row.FragmentCount != 1 ||
                        !DocxTableCellClipGeometry.ContainsHorizontalRange(viewport, geometryCell.X, geometryCell.Width) ||
                        !DocxTableCellClipGeometry.ContainsVerticalRange(viewport, geometryCell.Y, geometryCell.Y + geometryCell.Height))
                        continue;
                }
                var clip = DocxTableCellClipGeometry.Resolve(geometryRow, geometryCell, viewport);
                foreach (var nested in cell.NestedRows)
                    AddTargets(nested, clip);
                foreach (var box in cell.InlineTextBoxes)
                    AddBoxTargets(box, clip);
            }
        }
    }

    private static void RenderTableRow(
        DocxTableRowLayout row,
        DocxTableRowLayout? previousRow,
        DocxTableRowLayout? nextRow,
        ref DocxTableRowLayout? nextRowGeometry,
        PdfGraphicsBuilder graphics,
        List<PdfImageResource> pageImages,
        DocxFontResources fontResources,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        DocxPageNumber pageNumber,
        int pageCount,
        CancellationToken cancellationToken,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int imageIndex,
        PictureNavigation navigation,
        RgbColor? background = null)
    {
        // Office A/B (comment-table and w6-celltuckborders probes, Word-COM rendered):
        // the first-pin emission shift moves text but table geometry was painted raw, so
        // cell text rendered a full offset below its rows and clips. Geometry follows the
        // same text offset (a no-op outside the Word-compatible profile, where it is 0).
        cancellationToken.ThrowIfCancellationRequested();
        double geometryXOffset = ResolveTextEmissionXOffset(markupContext);
        double geometryYOffset = ResolveTextEmissionBaselineOffset(markupContext);
        // Each sweep visits the immediate next item in order. Only an adjacent
        // next row is projected, and it becomes the next iteration's current row.
        DocxTableRowLayout geometryRow = nextRowGeometry ?? ResolveTablePaintGeometry(row, geometryXOffset, geometryYOffset);
        DocxTableRowLayout? geometryNextRow = nextRow is null ? null : ResolveTablePaintGeometry(nextRow, geometryXOffset, geometryYOffset);
        nextRowGeometry = geometryNextRow;
        foreach (DocxTableCellLayout cellLayout in geometryRow.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ShouldRenderTableCellVisualFragment(cellLayout))
            {
                continue;
            }

            DocxTableCell cell = cellLayout.VisualCell;
            RenderShadingFill(cell.FillHex, cell.ShadingValue, cell.ShadingColor, graphics, cellLayout.X, cellLayout.Y, cellLayout.Width, cellLayout.Height);
        }

        double printScale = row.ReviewBorderPrintScale;
        bool scaleBorders = DocxTableBorderGeometry.TryResolveReviewPaintScale(printScale, out double borderScale) &&
            geometryRow.Cells.Count != 0 &&
            (HasVisibleBorder(geometryRow) || previousRow is not null && HasVisibleBorder(previousRow) ||
             geometryNextRow is not null && HasVisibleBorder(geometryNextRow));
        double borderOriginX = scaleBorders ? geometryRow.Cells[0].X : 0d;
        double borderOriginY = scaleBorders ? geometryRow.Y : 0d;
        DocxTableRowLayout borderRow = scaleBorders ? ProjectBorderGeometry(geometryRow) : geometryRow;
        DocxTableRowLayout? borderNextRow = scaleBorders && geometryNextRow is not null ? ProjectBorderGeometry(geometryNextRow) : geometryNextRow;
        if (scaleBorders)
        {
            graphics.SaveState();
            graphics.Transform(borderScale, 0d, 0d, borderScale, borderOriginX, borderOriginY);
        }
        RowPairBorderPlan? sharedBorderPlan = RowPairBorderPlan.TryBuild(borderRow, borderNextRow, cancellationToken);
        bool hasPreviousRow = previousRow is not null, hasNextRow = nextRow is not null;
        RenderTableRowBorders(borderRow, hasPreviousRow, hasNextRow, graphics, cancellationToken, sharedBorderPlan);
        RenderTableBorderJunctions(borderRow, hasPreviousRow, hasNextRow, graphics, cancellationToken, sharedBorderPlan);
        if (scaleBorders) graphics.RestoreState();
        RenderTableRowMarkupIndicators(row, graphics, markupContext);

        static bool HasVisibleBorder(DocxTableRowLayout input) => input.Cells.Any(cell =>
            cell.VisualCell.Borders.Any(border => DocxTableBorderGeometry.ResolveVisibleWidth(border) > 0d));

        // The layout has already shrunk positions and lengths. Project only border
        // coordinates through a local inverse, then let the PDF transform scale
        // widths, patterns and junctions together without scaling text or shading.
        DocxTableRowLayout ProjectBorderGeometry(DocxTableRowLayout input) => input with
        {
            Y = (input.Y - borderOriginY) / borderScale,
            Height = input.Height / borderScale,
            Cells = input.Cells.Select(cell => cell with
            {
                X = (cell.X - borderOriginX) / borderScale,
                Y = (cell.Y - borderOriginY) / borderScale,
                Width = cell.Width / borderScale,
                Height = cell.Height / borderScale
            }).ToArray()
        };

        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxTableCellLayout cellLayout = row.Cells[cellIndex];
            if (!ShouldRenderTableCellVisualFragment(cellLayout))
            {
                continue;
            }

            var cell = cellLayout.VisualCell;
            RgbColor? cellBackground = ResolveShadingBackground(cell.FillHex, cell.ShadingValue, cell.ShadingColor, background);
            if (cellLayout.TextLines.Count != 0 || cellLayout.InlineImages.Count != 0 || cellLayout.InlineTextBoxes.Count != 0 || cellLayout.NestedRows.Count != 0)
            {
                DocxTableCellLayout geometryCell = geometryRow.Cells[cellIndex];
                graphics.SaveState();
                DocxTableCellClipSnapshot clip = DocxTableCellClipGeometry.Resolve(geometryRow, geometryCell);
                graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);
                bool ownsCompleteCell = row.FragmentCount == 1 && OwnsUnmergedPictureCell(cellLayout);
                var cellNavigation = navigation.ClipTo(new PdfRectangle(clip.X, clip.Y, clip.Width, clip.Height)) with
                {
                    AllowsRotatedInlineUris = navigation.AllowsRotatedInlineUris && ownsCompleteCell,
                    // A fragmented or merged ancestor may replay an inner restart.
                    OwnsAdmittedMergeSource = navigation.AllowsRotatedInlineUris && cellLayout.OwnsAdmittedMergeSource
                };
                RenderParagraphBackgrounds(cellLayout.TextLines.Cast<DocxLayoutItem>().Concat(cellLayout.InlineImages), graphics, markupContext, cancellationToken);
                foreach (DocxTextLineLayout line in cellLayout.TextLines)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    RenderTextLine(line, graphics, fontResources, markupContext, pageNumber, pageCount, cancellationToken, cellBackground);
                }

                foreach (DocxInlineImageLayout image in cellLayout.InlineImages)
                {
                    // Pictures share the cell clip used by text, including aligned
                    // and repeated paragraphs.
                    DocxInlineImageLayout placedImage = PlaceCellImageForPaint(image, geometryXOffset, geometryYOffset, clip);
                    RenderInlineImage(placedImage, graphics, pageImages, diagnosticSink, cancellationToken, imageCache, ref imageIndex, fontResources, markupContext, cellNavigation);
                }

                // Only complete, unmerged cells can give an admitted outline its
                // own envelope. Each box restores the exact child content clip;
                // ancestor, fragmented and merged-cell clips remain in force.
                bool separateBoxClip = cellLayout.InlineTextBoxes.Count != 0 && ownsCompleteCell;
                if (separateBoxClip) graphics.RestoreState();
                foreach (DocxInlineTextBoxLayout textBox in cellLayout.InlineTextBoxes)
                {
                    RenderInlineTextBox(textBox, graphics, pageImages, fontResources, markupContext, diagnosticSink, pageNumber, pageCount, cancellationToken, imageCache, ref imageIndex,
                        cellNavigation, separateBoxClip ? clip : null, cellBackground);
                }
                if (separateBoxClip)
                {
                    graphics.SaveState();
                    graphics.ClipRectangle(clip.X, clip.Y, clip.Width, clip.Height);
                }

                DocxTableRowLayout? nextNestedGeometry = null;
                for (int nestedRowIndex = 0; nestedRowIndex < cellLayout.NestedRows.Count; nestedRowIndex++)
                {
                    DocxTableRowLayout nestedRow = cellLayout.NestedRows[nestedRowIndex];
                    DocxTableRowLayout? previousNestedRow = nestedRowIndex > 0 ? cellLayout.NestedRows[nestedRowIndex - 1] : null;
                    DocxTableRowLayout? nextNestedRow = nestedRowIndex + 1 < cellLayout.NestedRows.Count ? cellLayout.NestedRows[nestedRowIndex + 1] : null;
                    RenderTableRow(
                        nestedRow,
                        IsAdjacentTableRow(previousNestedRow, nestedRow) ? previousNestedRow : null,
                        IsAdjacentTableRow(nestedRow, nextNestedRow) ? nextNestedRow : null,
                        ref nextNestedGeometry,
                        graphics,
                        pageImages,
                        fontResources,
                        markupContext,
                        diagnosticSink,
                        pageNumber,
                        pageCount,
                        cancellationToken,
                        imageCache, ref imageIndex, cellNavigation, cellBackground);
                }

                graphics.RestoreState();
            }
        }
    }

    private static DocxTableRowLayout ShiftTableRowGeometry(DocxTableRowLayout row, double xOffset, double yOffset)
    {
        return row with
        {
            Y = row.Y - yOffset,
            Cells = row.Cells.Select(cell => cell with { X = cell.X + xOffset, Y = cell.Y - yOffset }).ToArray(),
        };
    }

    private static DocxTableRowLayout ResolveTablePaintGeometry(DocxTableRowLayout row, double xOffset, double yOffset) =>
        xOffset == 0d && yOffset == 0d ? row : ShiftTableRowGeometry(row, xOffset, yOffset);
    private static void RenderTableRowMarkupIndicators(
        DocxTableRowLayout row,
        PdfGraphicsBuilder graphics,
        DocxMarkupContext markupContext)
    {
        if (!markupContext.DrawsChangeBars ||
            row.RevisionCount == 0 ||
            UsesWordCompatibleAllMarkupTextProfile(markupContext))
        {
            return;
        }

        double height = Math.Max(6d, row.Height);
        DocxMarkupBalloonRgb color = ResolveRevisionAuthorColor(row.Revisions ?? []);
        graphics.SetFillRgb(color.Red, color.Green, color.Blue);
        graphics.FillRectangle(Math.Max(0d, row.Table.TableX - 7d), row.Y, 1.5d, height);
    }

    private static void RenderTableBorderJunctions(
        DocxTableRowLayout row,
        bool hasPreviousRow,
        bool hasNextRow,
        PdfGraphicsBuilder graphics,
        CancellationToken cancellationToken,
        RowPairBorderPlan? sharedPlan)
    {
        DocxTableBorderBoundary[] rowBoundaries = ResolveVisibleVerticalBoundaries(row, cancellationToken);
        if (rowBoundaries.Length == 0)
        {
            return;
        }

        OrderedBoundaryIndex rowBoundaryIndex = OrderedBoundaryIndex.Build(rowBoundaries);
        var emittedJunctions = new HashSet<(double X, double Y)>();
        foreach (DocxTableCellLayout cellLayout in row.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!hasPreviousRow && ShouldRenderTableCellVisualFragment(cellLayout) && ShouldRenderTableRowEdge(row, top: true))
            {
                RenderHorizontalBorderJunctions(cellLayout.X, cellLayout.X + cellLayout.Width, cellLayout.Y + cellLayout.Height, cellLayout.VisualCell, "top", rowBoundaryIndex.SliceRaw(cellLayout.X, cellLayout.X + cellLayout.Width, cancellationToken), graphics, emittedJunctions);
            }

            if (!hasNextRow && OwnsHorizontalBottom(cellLayout) && ShouldRenderTableRowEdge(row, top: false))
            {
                RenderHorizontalBorderJunctions(cellLayout.X, cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.VisualCell, "bottom", rowBoundaryIndex.SliceRaw(cellLayout.X, cellLayout.X + cellLayout.Width, cancellationToken), graphics, emittedJunctions);
            }
        }

        if (!hasPreviousRow && ShouldRenderTableRowEdge(row, top: true))
        {
            RenderOuterHorizontalFragmentCornerJunctions(row, rowBoundaries, "top", graphics);
        }

        if (!hasNextRow && ShouldRenderTableRowEdge(row, top: false))
        {
            RenderOuterHorizontalFragmentCornerJunctions(row, rowBoundaries, "bottom", graphics);
        }

        if (sharedPlan is not null)
        {
            RenderSharedHorizontalBorderJunctions(sharedPlan, rowBoundaries, rowBoundaryIndex, graphics, emittedJunctions, cancellationToken);
        }
    }

    private static void RenderOuterHorizontalFragmentCornerJunctions(
        DocxTableRowLayout row,
        IReadOnlyList<DocxTableBorderBoundary> boundaries,
        string edge,
        PdfGraphicsBuilder graphics)
    {
        bool top = string.Equals(edge, "top", StringComparison.Ordinal);
        Func<DocxTableCellLayout, bool> ownsEdge = top ? ShouldRenderTableCellVisualFragment : OwnsHorizontalBottom;
        DocxTableCellLayout? firstCell = row.Cells.FirstOrDefault(ownsEdge);
        DocxTableCellLayout? lastCell = row.Cells.LastOrDefault(ownsEdge);
        if (firstCell is null || lastCell is null)
        {
            return;
        }

        DocxTableCellBorder? firstHorizontal = DocxTableBorderGeometry.Find(firstCell.VisualCell.Borders, edge);
        DocxTableCellBorder? lastHorizontal = DocxTableBorderGeometry.Find(lastCell.VisualCell.Borders, edge);
        DocxTableBorderBoundary? firstBoundary = boundaries.OrderBy(boundary => boundary.X).FirstOrDefault();
        DocxTableBorderBoundary? lastBoundary = boundaries.OrderByDescending(boundary => boundary.X).FirstOrDefault();
        double y = top ? firstCell.Y + firstCell.Height : firstCell.Y;

        if (firstBoundary is not null && firstHorizontal is not null && !DocxTableBorderGeometry.IsSuppressed(firstHorizontal))
        {
            RenderBorderJunctions([firstBoundary], y, firstHorizontal, graphics, []);
        }

        if (lastBoundary is not null && lastHorizontal is not null && !DocxTableBorderGeometry.IsSuppressed(lastHorizontal))
        {
            RenderBorderJunctions([lastBoundary], y, lastHorizontal, graphics, []);
        }
    }

    private static void RenderSharedHorizontalBorderJunctions(
        RowPairBorderPlan plan,
        IReadOnlyList<DocxTableBorderBoundary> rowBoundaries,
        OrderedBoundaryIndex rowBoundaryIndex,
        PdfGraphicsBuilder graphics,
        HashSet<(double X, double Y)> emittedJunctions,
        CancellationToken cancellationToken)
    {
        DocxTableBorderBoundary[] nextRowBoundaries = ResolveVisibleVerticalBoundaries(plan.NextRow, cancellationToken);
        OrderedBoundaryIndex mergedBoundaries = OrderedBoundaryIndex.Build(rowBoundaries, nextRowBoundaries);
        int pairIndex = 0;
        for (int cellIndex = 0; cellIndex < plan.Row.Cells.Count; cellIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!plan.CurrentBottomEdges[cellIndex])
            {
                continue;
            }

            DocxTableCellLayout cellLayout = plan.Row.Cells[cellIndex];
            if (plan.OverlapCounts[cellIndex] == 0)
            {
                RenderHorizontalBorderJunctions(cellLayout.X, cellLayout.X + cellLayout.Width, cellLayout.Y, cellLayout.VisualCell, "bottom", rowBoundaryIndex.SliceRaw(cellLayout.X, cellLayout.X + cellLayout.Width, cancellationToken), graphics, emittedJunctions);
                continue;
            }

            for (int k = 0; k < plan.OverlapCounts[cellIndex]; k++, pairIndex++)
            {
                (DocxTableCellLayout current, DocxTableCellLayout nextRowCell, double x, double right) = plan.Overlaps[pairIndex];
                DocxTableCellBorder? horizontal = ResolveSharedHorizontalBorder(current, nextRowCell);
                if (horizontal is null)
                {
                    continue;
                }

                DocxTableBorderBoundary[] boundaries = mergedBoundaries.SliceGrouped(x, right, cancellationToken);
                RenderBorderJunctions(boundaries, current.Y - DocxTableBorderGeometry.ResolveVisibleWidth(horizontal) / 2d, horizontal, graphics, emittedJunctions);
            }
        }
    }

    private static void RenderHorizontalBorderJunctions(
        double x,
        double right,
        double y,
        DocxTableCell cell,
        string edge,
        IReadOnlyList<DocxTableBorderBoundary> boundaries,
        PdfGraphicsBuilder graphics,
        HashSet<(double X, double Y)> emittedJunctions)
    {
        DocxTableCellBorder? horizontal = DocxTableBorderGeometry.Find(cell.Borders, edge);
        if (horizontal is null || DocxTableBorderGeometry.IsSuppressed(horizontal))
        {
            return;
        }

        DocxTableBorderBoundary[] crossingBoundaries = boundaries
            .Where(boundary => boundary.X >= x - 0.001d && boundary.X <= right + 0.001d)
            .ToArray();
        RenderBorderJunctions(crossingBoundaries, y, horizontal, graphics, emittedJunctions);
    }

    private static void RenderBorderJunctions(
        IReadOnlyList<DocxTableBorderBoundary> boundaries,
        double y,
        DocxTableCellBorder horizontal,
        PdfGraphicsBuilder graphics,
        HashSet<(double X, double Y)> emittedJunctions)
    {
        double horizontalWidth = DocxTableBorderGeometry.ResolveVisibleWidth(horizontal);
        if (horizontalWidth <= 0d)
        {
            return;
        }

        foreach (DocxTableBorderBoundary boundary in boundaries)
        {
            if (!emittedJunctions.Add((Math.Round(boundary.X, 3), Math.Round(y, 3))))
            {
                continue;
            }

            DocxTableCellBorder? border = DocxTableBorderGeometry.SelectStronger(horizontal, boundary.Border);
            if (border is null)
            {
                continue;
            }

            RgbColor color = ReadColor(border.Color);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            RenderTableBorderStrip(graphics, border, boundary.X, y, boundary.Width, horizontalWidth, DocxTableBorderOrientation.Horizontal);
        }
    }

    private static DocxTableBorderBoundary[] ResolveVisibleVerticalBoundaries(DocxTableRowLayout row, CancellationToken cancellationToken)
    {
        var boundaries = new List<DocxTableBorderBoundary>();
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxTableCellLayout cellLayout = row.Cells[cellIndex];
            if (cellLayout.VisualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner)
            {
                continue;
            }

            DocxTableCell visualCell = cellLayout.VisualCell;
            DocxTableCellBorder? left = DocxTableBorderGeometry.Find(visualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "start");
            if (cellIndex == 0)
            {
                AddVerticalBoundary(boundaries, cellLayout.X, left);
            }

            DocxTableCellBorder? right = DocxTableBorderGeometry.Find(visualCell.Borders, "right") ?? DocxTableBorderGeometry.Find(visualCell.Borders, "end");
            if (cellIndex == row.Cells.Count - 1)
            {
                AddVerticalBoundary(boundaries, cellLayout.X + cellLayout.Width, right);
                continue;
            }

            DocxTableCellLayout nextCell = row.Cells[cellIndex + 1];
            DocxTableCell nextVisualCell = nextCell.VisualCell;
            DocxTableCellBorder? nextLeft = DocxTableBorderGeometry.Find(nextVisualCell.Borders, "left") ?? DocxTableBorderGeometry.Find(nextVisualCell.Borders, "start");
            if (!DocxTableBorderGeometry.IsSuppressed(right) && !DocxTableBorderGeometry.IsSuppressed(nextLeft))
            {
                AddVerticalBoundary(boundaries, cellLayout.X + cellLayout.Width, DocxTableBorderGeometry.SelectStronger(right, nextLeft));
            }
        }

        return boundaries.ToArray();
    }

    private static void AddVerticalBoundary(List<DocxTableBorderBoundary> boundaries, double x, DocxTableCellBorder? border)
    {
        double width = DocxTableBorderGeometry.ResolveVisibleWidth(border);
        if (width <= 0d || border is null)
        {
            return;
        }

        // Office A/B (w63-w68 probes): junctions ride the centered vertical bands.
        boundaries.Add(new DocxTableBorderBoundary(x - width / 2d, width, border));
    }
}
