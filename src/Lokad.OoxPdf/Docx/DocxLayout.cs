using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxLayout(
    IReadOnlyList<DocxLayoutPage> Pages,
    IReadOnlyList<DocxFloatingDrawingLayout> FloatingDrawings,
    IReadOnlyList<DocxFloatingDrawingLayout> StaticFloatingDrawings,
    IReadOnlyList<DocxRelatedStoryLayout> RelatedStories);

internal sealed record DocxRelatedStoryLayout(
    DocxRelatedStory Story,
    int StoryIndex,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    IReadOnlyList<DocxTableRowLayout> TableRows,
    double ContentHeight);

internal sealed record DocxFloatingDrawingLayout(
    DocxFloatingDrawing Drawing,
    int? PageStartIndex,
    int? PageEndIndex,
    int? AnchorPageIndex,
    int? AnchorColumnIndex,
    double? AnchorBlockVerticalTop,
    double? AnchorBlockVerticalBottom,
    double? ExtentWidthPoints,
    double? ExtentHeightPoints,
    double? HorizontalOffsetPoints,
    double? VerticalOffsetPoints,
    double? DistanceTopPoints,
    double? DistanceBottomPoints,
    double? DistanceLeftPoints,
    double? DistanceRightPoints,
    double? HorizontalReferenceX,
    double? HorizontalReferenceWidth,
    double? VerticalReferenceTop,
    double? VerticalReferenceBottom,
    double? PlacedX,
    double? PlacedTop,
    DocxAnchorPlacementSource? HorizontalPlacementSource,
    DocxAnchorPlacementSource? VerticalPlacementSource,
    double? WrapExclusionX,
    double? WrapExclusionTop,
    double? WrapExclusionWidth,
    double? WrapExclusionHeight,
    string? StoryKind = null,
    string? StoryVariantType = null);

internal sealed record DocxWrapExclusionFrame(
    double X,
    double Top,
    double Width,
    double Height);

internal enum DocxAnchorPlacementSource
{
    Align,
    Offset,
    Unsupported,
    MissingReferenceOrExtent
}

internal sealed record DocxAnchorPlacement(double? Position, DocxAnchorPlacementSource Source);

internal sealed record DocxLineHeightProfile(
    double LineHeight,
    double? SingleLineHeight,
    double? ListLabelSingleLineHeight,
    double? BodyWindowsLineHeight,
    double? ListLabelWindowsLineHeight,
    double? EffectiveLineSpacingFactor,
    bool LineSpacingFactorFloorApplied);

internal sealed record DocxParagraphSpacingProfile(
    double PendingAfterSpacing,
    double ParagraphBeforeSpacing,
    double ParagraphAfterSpacing,
    double AppliedBeforeSpacing,
    bool ContextualSpacingSuppressed);

internal sealed record DocxSectionLayoutProperties(
    string? BreakTypeValue,
    string? ColumnCountValue,
    string? ColumnEqualWidthValue,
    string? ColumnSpaceValue,
    int? ColumnCount,
    double? ColumnSpacePoints,
    IReadOnlyList<DocxSectionColumnLayoutProperties> ColumnDefinitions);

internal sealed record DocxSectionColumnLayoutProperties(
    string? WidthValue,
    string? SpaceValue,
    double? WidthPoints,
    double? SpacePoints);

internal sealed record DocxLayoutColumnFrame(
    int Index,
    double X,
    double Width,
    double? GutterAfterPoints);

internal static class DocxLayoutColumnOwnership
{
    public static int? ResolveColumnIndex(IReadOnlyList<DocxLayoutColumnFrame> frames, double x, double width)
    {
        if (frames.Count == 0)
        {
            return null;
        }

        double center = x + Math.Max(0d, width) / 2d;
        DocxLayoutColumnFrame? containingFrame = frames.FirstOrDefault(frame => center >= frame.X && center <= frame.X + frame.Width);
        if (containingFrame is not null)
        {
            return containingFrame.Index;
        }

        return frames
            .OrderBy(frame => Math.Abs(center - (frame.X + frame.Width / 2d)))
            .First()
            .Index;
    }
}

internal sealed record DocxLayoutSnapshot(
    IReadOnlyList<DocxLayoutPageSnapshot> Pages,
    IReadOnlyList<DocxTableSnapshot> Tables,
    IReadOnlyList<DocxLayoutSourceBlockSnapshot> SourceBlocks,
    IReadOnlyList<DocxFloatingDrawingLayoutSnapshot> FloatingDrawings,
    IReadOnlyList<DocxFloatingDrawingLayoutSnapshot> StaticFloatingDrawings,
    IReadOnlyList<DocxRelatedStoryLayoutSnapshot> RelatedStories)
{
    public static DocxLayoutSnapshot FromLayout(DocxLayout layout)
    {
        DocxLayoutPageSnapshot[] pages = layout.Pages.Select(ToSnapshot).ToArray();
        IReadOnlyList<DocxLayoutSourceBlockSnapshot> sourceBlocks = ToSourceBlockSnapshots(pages);
        return new DocxLayoutSnapshot(
            pages,
            ToTableSnapshots(pages),
            sourceBlocks,
            ToFloatingDrawingSnapshots(layout.FloatingDrawings),
            ToFloatingDrawingSnapshots(layout.StaticFloatingDrawings),
            ToRelatedStorySnapshots(layout.RelatedStories));
    }

    private static IReadOnlyList<DocxRelatedStoryLayoutSnapshot> ToRelatedStorySnapshots(
        IReadOnlyList<DocxRelatedStoryLayout> stories)
    {
        return stories
            .Select(story =>
            {
                int tableCellTextLineCount = CountTableCellTextLines(story.TableRows);
                int inlineImageCount = story.InlineImages.Count + CountTableCellInlineImages(story.TableRows);
                IReadOnlyList<DocxLayoutItemSnapshot> items = ToRelatedStoryItemSnapshots(story);
                return new DocxRelatedStoryLayoutSnapshot(
                    story.StoryIndex,
                    story.Story.Kind,
                    story.Story.PartName,
                    story.Story.Id,
                    story.Story.BodyElements.Count,
                    story.Story.BodyElements.OfType<DocxParagraphElement>().Count(),
                    DocxBlockTraversal.EnumerateBodyTables(story.Story).Count(),
                    story.TextLines.Count,
                    tableCellTextLineCount,
                    story.TableRows.Count,
                    inlineImageCount,
                    story.Story.FloatingDrawings.Count,
                    CountBodyTextLength(story.Story.BodyElements),
                    story.ContentHeight,
                    items,
                    ToRelatedStorySourceBlockSnapshots(items),
                    story.TableRows.Select((row, rowIndex) => ToTableRowSnapshot(row, rowIndex, pageMarginBottom: 0d)).ToArray());
            })
            .ToArray();
    }

    private static IReadOnlyList<DocxRelatedStorySourceBlockSnapshot> ToRelatedStorySourceBlockSnapshots(
        IReadOnlyList<DocxLayoutItemSnapshot> items)
    {
        return items
            .Where(item => item.SourceBlockIndex is not null)
            .GroupBy(item => item.SourceBlockIndex!.Value)
            .OrderBy(group => group.Key)
            .Select(group =>
            {
                DocxLayoutItemSnapshot[] blockItems = group.ToArray();
                double verticalTop = blockItems.Max(item => item.Y + item.Height);
                double verticalBottom = blockItems.Min(item => item.Y);
                return new DocxRelatedStorySourceBlockSnapshot(
                    group.Key,
                    ResolveSourceBlockKind(blockItems),
                    blockItems.Length,
                    blockItems.Count(item => item.Kind == "TextLine"),
                    blockItems.Count(item => item.Kind == "InlineImage"),
                    blockItems.Count(item => item.Kind == "TableRow"),
                    blockItems.Sum(item => item.TextLength),
                    verticalTop,
                    verticalBottom,
                    Math.Max(0d, verticalTop - verticalBottom),
                    blockItems.Sum(item => item.AppliedBeforeSpacingPoints ?? 0d));
            })
            .ToArray();
    }

    private static IReadOnlyList<DocxLayoutItemSnapshot> ToRelatedStoryItemSnapshots(DocxRelatedStoryLayout story)
    {
        return story.TextLines
            .Cast<DocxLayoutItem>()
            .Concat(story.InlineImages)
            .Concat(story.TableRows)
            .OrderBy(item => GetSourceBlockIndex(item) ?? int.MaxValue)
            .ThenByDescending(item => GetVerticalBounds(item).Y)
            .Select(item => ToSnapshot(item, []))
            .ToArray();
    }

    private static int? GetSourceBlockIndex(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => text.SourceBlockIndex,
            DocxInlineImageLayout image => image.SourceBlockIndex,
            DocxTableRowLayout row => row.Table.SourceBlockIndex,
            _ => null
        };
    }

    private static (double Y, double Height) GetVerticalBounds(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => (text.BaselineY, text.FontSize),
            DocxInlineImageLayout image => (image.Y, image.Height),
            DocxTableRowLayout row => (row.Y, row.Height),
            _ => (0d, 0d)
        };
    }

    private static int CountBodyTextLength(IReadOnlyList<DocxBodyElement> elements)
    {
        int length = 0;
        foreach (DocxBodyElement element in elements)
        {
            if (element is DocxParagraphElement paragraphElement)
            {
                length += paragraphElement.Paragraph.Runs.Sum(run => run.Text.Length);
                continue;
            }

            if (element is DocxTableElement tableElement)
            {
                foreach (DocxTableRow row in tableElement.Table.Rows)
                {
                    foreach (DocxTableCell cell in row.Cells)
                    {
                        length += CountBodyTextLength(DocxTableCellContent.GetBodyElements(cell));
                    }
                }
            }
        }

        return length;
    }

    private static int CountTableCellTextLines(IReadOnlyList<DocxTableRowLayout> rows)
    {
        int count = 0;
        foreach (DocxTableRowLayout row in rows)
        {
            foreach (DocxTableCellLayout cell in row.Cells)
            {
                count += cell.TextLines.Count;
                count += CountTableCellTextLines(cell.NestedRows);
            }
        }

        return count;
    }

    private static int CountTableCellInlineImages(IReadOnlyList<DocxTableRowLayout> rows)
    {
        int count = 0;
        foreach (DocxTableRowLayout row in rows)
        {
            foreach (DocxTableCellLayout cell in row.Cells)
            {
                count += cell.InlineImages.Count;
                count += CountTableCellInlineImages(cell.NestedRows);
            }
        }

        return count;
    }

    private static IReadOnlyList<DocxFloatingDrawingLayoutSnapshot> ToFloatingDrawingSnapshots(
        IReadOnlyList<DocxFloatingDrawingLayout> drawings)
    {
        return drawings
            .Select((drawing, index) =>
            {
                return new DocxFloatingDrawingLayoutSnapshot(
                    index,
                    drawing.Drawing.SourceBlockIndex,
                    drawing.Drawing.SourceParagraphIndex,
                    drawing.PageStartIndex,
                    drawing.PageEndIndex,
                    drawing.AnchorPageIndex,
                    drawing.AnchorColumnIndex,
                    drawing.AnchorBlockVerticalTop,
                    drawing.AnchorBlockVerticalBottom,
                    drawing.ExtentWidthPoints,
                    drawing.ExtentHeightPoints,
                    drawing.HorizontalOffsetPoints,
                    drawing.VerticalOffsetPoints,
                    drawing.DistanceTopPoints,
                    drawing.DistanceBottomPoints,
                    drawing.DistanceLeftPoints,
                    drawing.DistanceRightPoints,
                    drawing.HorizontalReferenceX,
                    drawing.HorizontalReferenceWidth,
                    drawing.VerticalReferenceTop,
                    drawing.VerticalReferenceBottom,
                    drawing.PlacedX,
                    drawing.PlacedTop,
                    drawing.HorizontalPlacementSource?.ToString(),
                    drawing.VerticalPlacementSource?.ToString(),
                    drawing.WrapExclusionX,
                    drawing.WrapExclusionTop,
                    drawing.WrapExclusionWidth,
                    drawing.WrapExclusionHeight,
                    drawing.Drawing.WrapKind,
                    drawing.Drawing.WrapTextValue,
                    drawing.Drawing.HorizontalRelativeFromValue,
                    drawing.Drawing.HorizontalAlignValue,
                    drawing.Drawing.HorizontalOffsetValue,
                    drawing.Drawing.VerticalRelativeFromValue,
                    drawing.Drawing.VerticalAlignValue,
                    drawing.Drawing.VerticalOffsetValue,
                    drawing.Drawing.ImageRelationshipId,
                    drawing.Drawing.Image?.PartName,
                    drawing.Drawing.Image?.ContentType,
                    drawing.Drawing.Image?.WidthPoints,
                    drawing.Drawing.Image?.HeightPoints,
                    drawing.StoryKind,
                    drawing.StoryVariantType);
            })
            .ToArray();
    }

    private static IReadOnlyList<DocxTableSnapshot> ToTableSnapshots(IReadOnlyList<DocxLayoutPageSnapshot> pages)
    {
        return pages
            .SelectMany((page, pageIndex) => page.TableRows.Select(row => (pageIndex, row)))
            .GroupBy(entry => new TableSnapshotKey(
                entry.row.StoryKind,
                entry.row.StoryVariantType,
                entry.row.TableIndex,
                entry.row.SourceBlockIndex))
            .OrderBy(group => group.Key.StoryKind ?? string.Empty, StringComparer.Ordinal)
            .ThenBy(group => group.Key.StoryVariantType ?? string.Empty, StringComparer.Ordinal)
            .ThenBy(group => group.Key.TableIndex)
            .ThenBy(group => group.Key.SourceBlockIndex)
            .Select(group =>
            {
                DocxTableRowSnapshot first = group.First().row;
                DocxTableRowSnapshot[] distinctRows = group
                    .Select(entry => entry.row)
                    .GroupBy(row => row.RowIndex)
                    .Select(rowGroup => rowGroup.First())
                    .ToArray();
                DocxTableCellSnapshot[] authoredCells = distinctRows.SelectMany(row => row.Cells).ToArray();
                DocxTableCellSnapshot[] laidOutCells = group.SelectMany(entry => entry.row.Cells).ToArray();
                return new DocxTableSnapshot(
                    group.Key.TableIndex,
                    first.SourceBlockIndex,
                    group.Min(entry => entry.pageIndex),
                    group.Max(entry => entry.pageIndex),
                    first.TableRowCount,
                    group.Count(),
                    group.Count(entry => entry.row.IsHeader),
                    distinctRows.Count(row => row.IsHeader),
                    first.GridColumnCount,
                    first.GridColumnsWidthSum,
                    first.HasExplicitGrid,
                    first.ResolvedColumnWidths,
                    first.ResolvedTableWidth,
                    first.TableX,
                    first.PreferredTableWidthPoints,
                    first.PreferredTableWidthValue,
                    first.PreferredTableWidthType,
                    first.IndentPoints,
                    first.CellSpacingPoints,
                    first.LayoutValue,
                    distinctRows.Count(row => row.DeclaredHeightPoints is not null),
                    distinctRows.Count(row => string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase)),
                    distinctRows.Count(row => string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase)),
                    distinctRows.Count(row => row.CantSplit),
                    distinctRows.Count(row => row.FragmentCount > 1),
                    group.Count(entry => entry.row.FragmentCount > 1),
                    group.Max(entry => entry.row.FragmentCount),
                    authoredCells.Any(cell => cell.HasVerticalMerge),
                    authoredCells.Count(cell => cell.HasVerticalMerge),
                    authoredCells.Count(cell => string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase)),
                    authoredCells.Count(cell => cell.IsVerticalMergeContinuation),
                    laidOutCells.Count(cell => cell.IsVerticalMergeContinuation),
                    laidOutCells.Count(cell => string.Equals(cell.VisualOwnership, DocxTableCellVisualOwnership.MissingVerticalMergeOwner.ToString(), StringComparison.Ordinal)),
                    first.StoryKind,
                    first.StoryVariantType);
            })
            .ToArray();
    }

    private sealed record TableSnapshotKey(string? StoryKind, string? StoryVariantType, int TableIndex, int SourceBlockIndex);

    private static IReadOnlyList<DocxLayoutSourceBlockSnapshot> ToSourceBlockSnapshots(IReadOnlyList<DocxLayoutPageSnapshot> pages)
    {
        return pages
            .SelectMany((page, pageIndex) => page.Items
                .Where(item => item.SourceBlockIndex is not null)
                .Select(item => (pageIndex, item)))
            .GroupBy(entry => entry.item.SourceBlockIndex!.Value)
            .OrderBy(group => group.Key)
            .Select(group => new DocxLayoutSourceBlockSnapshot(
                group.Key,
                ResolveSourceBlockKind(group.Select(entry => entry.item)),
                group.Min(entry => entry.pageIndex),
                group.Max(entry => entry.pageIndex),
                group.Select(entry => entry.item.ColumnIndex).Where(index => index is not null).DefaultIfEmpty().Min(),
                group.Select(entry => entry.item.ColumnIndex).Where(index => index is not null).DefaultIfEmpty().Max(),
                group.Max(entry => entry.item.Y + entry.item.Height),
                group.Min(entry => entry.item.Y),
                group.Count(),
                group.Count(entry => entry.item.Kind == "TextLine"),
                group.Count(entry => entry.item.Kind == "TableRow"),
                group.Count(entry => entry.item.Kind == "InlineImage"),
                group.Sum(entry => entry.item.TextLength),
                group.Sum(entry => entry.item.LineHeightPoints ?? entry.item.Height),
                group.Sum(entry => entry.item.AppliedBeforeSpacingPoints ?? 0d)))
            .ToArray();
    }

    private static string ResolveSourceBlockKind(IEnumerable<DocxLayoutItemSnapshot> items)
    {
        bool hasTextLine = false;
        bool hasTableRow = false;
        bool hasInlineImage = false;
        foreach (DocxLayoutItemSnapshot item in items)
        {
            hasTextLine |= item.Kind == "TextLine";
            hasTableRow |= item.Kind == "TableRow";
            hasInlineImage |= item.Kind == "InlineImage";
        }

        int kindCount = (hasTextLine ? 1 : 0) + (hasTableRow ? 1 : 0) + (hasInlineImage ? 1 : 0);
        if (kindCount > 1)
        {
            return "Mixed";
        }

        if (hasTableRow)
        {
            return "Table";
        }

        if (hasTextLine)
        {
            return "Paragraph";
        }

        return hasInlineImage ? "InlineImage" : "Unknown";
    }

    private static DocxLayoutPageSnapshot ToSnapshot(DocxLayoutPage page)
    {
        IReadOnlyList<DocxLayoutColumnFrameSnapshot> columnFrames = page.ColumnFrames.Select(frame => new DocxLayoutColumnFrameSnapshot(
            frame.Index,
            frame.X,
            frame.Width,
            frame.GutterAfterPoints)).ToArray();
        IReadOnlyList<DocxLayoutItemSnapshot> items = page.Items.Select(item => ToSnapshot(item, page.ColumnFrames)).ToArray();
        IReadOnlyList<DocxLayoutItemSnapshot> staticItems = page.StaticTextLines
            .Select(ToStaticSnapshot)
            .Concat(page.StaticInlineImages.Select(ToStaticSnapshot))
            .Concat(page.StaticTableRows.Select(ToStaticSnapshot))
            .ToArray();
        int?[] sourceBlockIndexes = items
            .Select(item => item.SourceBlockIndex)
            .Where(index => index is not null)
            .ToArray();
        IReadOnlyList<DocxTableRowSnapshot> tableRows = page.Items
            .OfType<DocxTableRowLayout>()
            .Concat(page.StaticTableRows)
            .Select((row, rowIndex) => ToTableRowSnapshot(row, rowIndex, page.MarginBottom))
            .ToArray();
        double verticalTop = items.Count == 0 ? 0d : items.Max(item => item.Y + item.Height);
        double verticalBottom = items.Count == 0 ? 0d : items.Min(item => item.Y);
        return new DocxLayoutPageSnapshot(
            page.Width,
            page.Height,
            page.MarginLeft,
            page.MarginRight,
            page.MarginTop,
            page.MarginBottom,
            page.SectionProperties.BreakTypeValue,
            page.SectionProperties.ColumnCountValue,
            page.SectionProperties.ColumnEqualWidthValue,
            page.SectionProperties.ColumnSpaceValue,
            page.SectionProperties.ColumnCount,
            page.SectionProperties.ColumnSpacePoints,
            page.SectionProperties.ColumnDefinitions.Count,
            page.SectionProperties.ColumnDefinitions.Sum(column => column.WidthPoints ?? 0d),
            page.SectionProperties.ColumnDefinitions.Sum(column => column.SpacePoints ?? 0d),
            page.ColumnFrames.Count,
            page.ColumnFrames.Sum(frame => frame.Width),
            page.ColumnFrames.Sum(frame => frame.GutterAfterPoints ?? 0d),
            columnFrames,
            items.Count,
            page.StaticTextLines.Count,
            page.StaticInlineImages.Count,
            page.StaticTableRows.Count,
            items.Count(item => item.Kind == "TextLine"),
            items.Count(item => item.Kind == "InlineImage"),
            items.Count(item => item.Kind == "TableRow"),
            sourceBlockIndexes.Distinct().Count(),
            sourceBlockIndexes.FirstOrDefault(),
            sourceBlockIndexes.LastOrDefault(),
            Math.Max(0d, verticalTop - verticalBottom),
            items.Where(item => item.Kind == "TextLine").Sum(item => item.Height),
            items.Where(item => item.Kind == "InlineImage").Sum(item => item.Height),
            items.Where(item => item.Kind == "TableRow").Sum(item => item.Height),
            ToStaticStorySnapshots(staticItems),
            staticItems,
            items,
            tableRows);
    }

    private static IReadOnlyList<DocxStaticStoryLayoutSnapshot> ToStaticStorySnapshots(
        IReadOnlyList<DocxLayoutItemSnapshot> staticItems)
    {
        return staticItems
            .GroupBy(item => new StaticStorySnapshotKey(StaticStoryKind(item.Kind), item.StoryVariantType))
            .OrderBy(group => group.Key.Kind, StringComparer.Ordinal)
            .ThenBy(group => group.Key.VariantType ?? string.Empty, StringComparer.Ordinal)
            .Select(group =>
            {
                DocxLayoutItemSnapshot[] storyItems = group.ToArray();
                int[] paragraphIndexes = storyItems
                    .Select(item => item.SourceParagraphIndex)
                    .Where(index => index is not null)
                    .Select(index => index!.Value)
                    .Distinct()
                    .OrderBy(index => index)
                    .ToArray();
                int[] lineIndexes = storyItems
                    .Select(item => item.SourceLineIndex)
                    .Where(index => index is not null)
                    .Select(index => index!.Value)
                    .Distinct()
                    .OrderBy(index => index)
                    .ToArray();
                return new DocxStaticStoryLayoutSnapshot(
                    group.Key.Kind,
                    group.Key.VariantType,
                    storyItems.Count(item => item.Kind.EndsWith("TextLine", StringComparison.Ordinal)),
                    storyItems.Count(item => item.Kind.EndsWith("InlineImage", StringComparison.Ordinal)),
                    storyItems.Count(item => item.Kind.EndsWith("TableRow", StringComparison.Ordinal)),
                    paragraphIndexes.Length,
                    lineIndexes.Length,
                    storyItems.Sum(item => item.TextLength),
                    storyItems.Count(item => item.IsFirstParagraphLine == true),
                    storyItems.Min(item => item.Y),
                    storyItems.Max(item => item.Y + item.Height),
                    paragraphIndexes.Length == 0 ? null : paragraphIndexes.First(),
                    paragraphIndexes.Length == 0 ? null : paragraphIndexes.Last(),
                    lineIndexes.Length == 0 ? null : lineIndexes.First(),
                    lineIndexes.Length == 0 ? null : lineIndexes.Last(),
                    storyItems);
            })
            .ToArray();
    }

    private static string StaticStoryKind(string itemKind)
    {
        return itemKind switch
        {
            "StaticHeaderTextLine" => "Header",
            "StaticFooterTextLine" => "Footer",
            "StaticHeaderInlineImage" => "Header",
            "StaticFooterInlineImage" => "Footer",
            "StaticHeaderTableRow" => "Header",
            "StaticFooterTableRow" => "Footer",
            _ => "Static"
        };
    }

    private sealed record StaticStorySnapshotKey(string Kind, string? VariantType);

    private static DocxLayoutItemSnapshot ToStaticSnapshot(DocxTextLineLayout text)
    {
        string kind = text.StoryKind switch
        {
            "Header" => "StaticHeaderTextLine",
            "Footer" => "StaticFooterTextLine",
            _ => "StaticTextLine"
        };
        return ToSnapshot(text, []) with { Kind = kind };
    }

    private static DocxLayoutItemSnapshot ToStaticSnapshot(DocxInlineImageLayout image)
    {
        string kind = image.StoryKind switch
        {
            "Header" => "StaticHeaderInlineImage",
            "Footer" => "StaticFooterInlineImage",
            _ => "StaticInlineImage"
        };
        return ToSnapshot(image, []) with { Kind = kind };
    }

    private static DocxLayoutItemSnapshot ToStaticSnapshot(DocxTableRowLayout row)
    {
        string kind = row.StoryKind switch
        {
            "Header" => "StaticHeaderTableRow",
            "Footer" => "StaticFooterTableRow",
            _ => "StaticTableRow"
        };
        return ToSnapshot(row, []) with { Kind = kind };
    }

    private static DocxLayoutItemSnapshot ToSnapshot(DocxLayoutItem item, IReadOnlyList<DocxLayoutColumnFrame> columnFrames)
    {
        (double itemX, double itemWidth) = GetHorizontalBounds(item);
        int? columnIndex = DocxLayoutColumnOwnership.ResolveColumnIndex(columnFrames, itemX, itemWidth);
        return item switch
        {
            DocxTextLineLayout text => new DocxLayoutItemSnapshot(
                "TextLine",
                text.X,
                text.BaselineY,
                text.Width,
                text.FontSize,
                TextLength: text.Text.Length,
                CellCount: 0,
                columnIndex,
                text.SourceBlockIndex,
                text.SourceParagraphIndex,
                text.SourceLineIndex,
                text.LineHeight,
                text.AppliedBeforeSpacing,
                text.SingleLineHeight,
                text.ListLabelSingleLineHeight,
                text.BodyWindowsLineHeight,
                text.ListLabelWindowsLineHeight,
                text.EffectiveLineSpacingFactor,
                text.LineSpacingFactorFloorApplied,
                text.IsFirstParagraphLine,
                text.PendingAfterSpacing,
                text.ParagraphBeforeSpacing,
                text.ParagraphAfterSpacing,
                text.ContextualSpacingSuppressed,
                text.StoryVariantType,
                ParagraphStyleId: text.SourceParagraph?.StyleResolution.StyleId,
                ParagraphStyleFound: text.SourceParagraph?.StyleResolution.StyleFound,
                ParagraphStyleDepth: text.SourceParagraph?.StyleResolution.StyleDepth,
                HasDocumentDefaultParagraphProperties: text.SourceParagraph?.StyleResolution.HasDocumentDefaultParagraphProperties,
                HasDirectParagraphProperties: text.SourceParagraph?.StyleResolution.HasDirectParagraphProperties,
                HasTableStyleParagraphProperties: text.SourceParagraph?.StyleResolution.HasTableStyleParagraphProperties,
                CharacterStyleTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasCharacterStyleRunProperties),
                DirectRunPropertyTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasDirectRunProperties),
                ParagraphStyleRunPropertyTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasParagraphStyleRunProperties),
                TableStyleRunPropertyTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasTableStyleRunProperties),
                DocumentDefaultRunPropertyTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasDocumentDefaultRunProperties)),
            DocxInlineImageLayout image => new DocxLayoutItemSnapshot(
                "InlineImage",
                image.X,
                image.Y,
                image.Width,
                image.Height,
                TextLength: 0,
                CellCount: 0,
                columnIndex,
                image.SourceBlockIndex,
                image.SourceParagraphIndex,
                SourceLineIndex: null,
                LineHeightPoints: null,
                AppliedBeforeSpacingPoints: null,
                SingleLineHeightPoints: null,
                ListLabelSingleLineHeightPoints: null,
                BodyWindowsLineHeightPoints: null,
                ListLabelWindowsLineHeightPoints: null,
                EffectiveLineSpacingFactor: null,
                LineSpacingFactorFloorApplied: null,
                IsFirstParagraphLine: null,
                PendingAfterSpacingPoints: null,
                ParagraphBeforeSpacingPoints: null,
                ParagraphAfterSpacingPoints: null,
                ContextualSpacingSuppressed: null,
                image.StoryVariantType),
            DocxTableRowLayout row => new DocxLayoutItemSnapshot(
                "TableRow",
                row.Cells.Count == 0 ? 0d : row.Cells.Min(cell => cell.X),
                row.Y,
                row.Cells.Sum(cell => cell.Width),
                row.Height,
                TextLength: SumTableRowTextLength(row),
                CellCount: row.Cells.Count,
                columnIndex,
                SourceBlockIndex: row.Table.SourceBlockIndex,
                SourceParagraphIndex: null,
                SourceLineIndex: null,
                LineHeightPoints: null,
                AppliedBeforeSpacingPoints: null,
                SingleLineHeightPoints: null,
                ListLabelSingleLineHeightPoints: null,
                BodyWindowsLineHeightPoints: null,
                ListLabelWindowsLineHeightPoints: null,
                EffectiveLineSpacingFactor: null,
                LineSpacingFactorFloorApplied: null,
                IsFirstParagraphLine: null,
                PendingAfterSpacingPoints: null,
                ParagraphBeforeSpacingPoints: null,
                ParagraphAfterSpacingPoints: null,
                ContextualSpacingSuppressed: null,
                row.StoryVariantType,
                ToTableRowTextLineSnapshots(row)),
            _ => new DocxLayoutItemSnapshot("Unknown", 0d, 0d, 0d, 0d, 0, 0, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null)
        };
    }

    private static int CountTextSegments(DocxTextLineLayout text, Func<DocxRunStyleResolution, bool> predicate)
    {
        return text.Segments.Count(segment => predicate(segment.StyleRun.StyleResolution));
    }

    private static IReadOnlyList<DocxLayoutItemSnapshot> ToTableRowTextLineSnapshots(DocxTableRowLayout row)
    {
        List<DocxLayoutItemSnapshot> lines = [];
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            lines.AddRange(cell.TextLines.Select(line => ToSnapshot(line, [])));
            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                lines.AddRange(ToTableRowTextLineSnapshots(nestedRow));
            }
        }

        return lines
            .OrderByDescending(line => line.Y)
            .ThenBy(line => line.X)
            .ToArray();
    }

    private static (double X, double Width) GetHorizontalBounds(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => (text.X, text.Width),
            DocxInlineImageLayout image => (image.X, image.Width),
            DocxTableRowLayout row => (row.Table.TableX, row.Table.ResolvedTableWidth),
            _ => (0d, 0d)
        };
    }

    private static DocxTableRowSnapshot ToTableRowSnapshot(DocxTableRowLayout row, int rowIndex, double pageMarginBottom)
    {
        IReadOnlyList<DocxTableCellSnapshot> cells = row.Cells
            .Select(ToTableCellSnapshot)
            .ToArray();
        double? firstBaselineY = cells
            .Select(cell => cell.FirstBaselineY)
            .Where(baseline => baseline is not null)
            .DefaultIfEmpty(null)
            .Min();
        double? lastBaselineY = cells
            .Select(cell => cell.LastBaselineY)
            .Where(baseline => baseline is not null)
            .DefaultIfEmpty(null)
            .Min();
        return new DocxTableRowSnapshot(
            row.Table.TableIndex,
            row.Table.SourceBlockIndex,
            rowIndex,
            row.RowIndex,
            row.Table.RowCount,
            row.Table.GridColumnCount,
            row.Table.GridColumnsWidthSum,
            row.Table.HasExplicitGrid,
            row.Table.ResolvedColumnWidths,
            row.Table.ResolvedTableWidth,
            row.Table.TableX,
            row.Table.PreferredWidthPoints,
            row.Table.PreferredWidthValue,
            row.Table.PreferredWidthType,
            row.Table.IndentPoints,
            row.Table.CellSpacingPoints,
            row.Table.LayoutValue,
            row.HeightValue,
            row.HeightRuleValue,
            row.DeclaredHeightPoints,
            row.FragmentIndex,
            row.FragmentCount,
            row.FragmentReason,
            row.FullRowHeight,
            row.FragmentOffsetFromRowTop,
            row.Cells.Count == 0 ? 0d : row.Cells.Min(cell => cell.X),
            row.Y,
            row.Cells.Sum(cell => cell.Width),
            row.Height,
            Math.Max(0d, pageMarginBottom - row.Y),
            firstBaselineY,
            lastBaselineY,
            cells.Count,
            SumTableRowTextLineCount(row),
            SumTableRowTextLength(row),
            cells.Select(cell => cell.MaxFontSize).DefaultIfEmpty(0d).Max(),
            row.IsHeader,
            row.HeaderValue,
            row.HasTablePropertyExceptionCellMargins,
            row.CantSplit,
            row.CantSplitValue,
            row.StoryKind,
            row.StoryVariantType,
            cells);
    }

    private static int SumTableRowTextLineCount(DocxTableRowLayout row)
    {
        return row.Cells.Sum(SumTableCellTextLineCount);
    }

    private static int SumTableRowTextLength(DocxTableRowLayout row)
    {
        return row.Cells.Sum(SumTableCellTextLength);
    }

    private static int SumTableCellTextLineCount(DocxTableCellLayout cell)
    {
        return cell.TextLines.Count + cell.NestedRows.Sum(SumTableRowTextLineCount);
    }

    private static int SumTableCellTextLength(DocxTableCellLayout cell)
    {
        return cell.TextLines.Sum(line => line.Text.Length) + cell.NestedRows.Sum(SumTableRowTextLength);
    }

    private static DocxTableCellSnapshot ToTableCellSnapshot(DocxTableCellLayout cellLayout, int cellIndex)
    {
        DocxTableCell cell = cellLayout.Cell;
        DocxTableCell visualCell = cellLayout.VisualCell;
        DocxTextLineLayout? firstLine = cellLayout.TextLines.FirstOrDefault();
        DocxTextLineLayout? lastLine = cellLayout.TextLines.LastOrDefault();
        IReadOnlyList<DocxBodyElement> bodyElements = DocxTableCellContent.GetBodyElements(cell);
        IReadOnlyList<DocxParagraph> paragraphs = GetParagraphsFromBodyElements(bodyElements);
        IReadOnlyList<DocxBodyElement> visualBodyElements = DocxTableCellContent.GetBodyElements(visualCell);
        IReadOnlyList<DocxParagraph> visualParagraphs = GetParagraphsFromBodyElements(visualBodyElements);
        IReadOnlyList<double> spacingBeforePoints = paragraphs.Select(paragraph => paragraph.SpacingBeforePoints).ToArray();
        IReadOnlyList<double> spacingAfterPoints = paragraphs.Select(paragraph => paragraph.SpacingAfterPoints).ToArray();
        string cellText = string.Concat(paragraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.Text));
        string visualCellText = string.Concat(visualParagraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.Text));
        TextProfile textProfile = BuildTextProfile(cellText);
        return new DocxTableCellSnapshot(
            cellIndex,
            cellLayout.X,
            cellLayout.Y,
            cellLayout.Width,
            cellLayout.Height,
            cellLayout.X + cellLayout.ContentPaddingLeft,
            cellLayout.Y + cellLayout.ContentPaddingBottom,
            Math.Max(0d, cellLayout.Width - cellLayout.ContentPaddingLeft - cellLayout.ContentPaddingRight),
            Math.Max(0d, cellLayout.Height - cellLayout.ContentPaddingTop - cellLayout.ContentPaddingBottom),
            cellLayout.ContentPaddingLeft,
            cellLayout.ContentPaddingTop,
            cellLayout.ContentPaddingRight,
            cellLayout.ContentPaddingBottom,
            SumTableCellTextLineCount(cellLayout),
            SumTableCellTextLength(cellLayout),
            cellLayout.TextLines.Count == 0 ? 0d : cellLayout.TextLines.Max(line => line.FontSize),
            firstLine?.X,
            firstLine?.BaselineY,
            DocxLineMetrics.ResolveTableCellFirstBaselineInset(visualParagraphs),
            lastLine?.BaselineY,
            cellLayout.InlineImages.Count,
            paragraphs.Count,
            paragraphs.Count(paragraph => HasBeforeSpacingToken(paragraph.Spacing)),
            paragraphs.Count(paragraph => HasAfterSpacingToken(paragraph.Spacing)),
            paragraphs.Count(paragraph => string.Equals(paragraph.Spacing.BeforeValue, "0", StringComparison.Ordinal)),
            paragraphs.Count(paragraph => string.Equals(paragraph.Spacing.AfterValue, "0", StringComparison.Ordinal)),
            spacingBeforePoints.Count == 0 ? null : spacingBeforePoints.Min(),
            spacingBeforePoints.Count == 0 ? null : spacingBeforePoints.Max(),
            spacingAfterPoints.Count == 0 ? null : spacingAfterPoints.Min(),
            spacingAfterPoints.Count == 0 ? null : spacingAfterPoints.Max(),
            cell.GridSpan,
            cell.GridSpanValue,
            cell.PreferredWidthPoints,
            cell.PreferredWidthValue,
            cell.PreferredWidthType,
            cell.VerticalAlignmentValue,
            cell.Margins.TopPoints,
            cell.Margins.RightPoints,
            cell.Margins.BottomPoints,
            cell.Margins.LeftPoints,
            cell.Borders.Count,
            cell.FillHex is not null,
            cell.ShadingValue is not null,
            cell.ConditionalFormat?.IsDefined == true,
            cell.HasVerticalMerge,
            cell.VerticalMergeValue,
            cellLayout.IsVerticalMergeContinuation,
            cellLayout.VisualOwnership.ToString(),
            cellLayout.VerticalMergeOwner?.RowIndex,
            cellLayout.VerticalMergeOwner?.GridColumnIndex,
            visualParagraphs.Count,
            visualCellText.Length,
            visualParagraphs.Sum(paragraph => paragraph.Images.Count),
            textProfile.SpaceCharacterCount,
            textProfile.NonAsciiCharacterCount,
            textProfile.PunctuationCharacterCount,
            textProfile.DigitCharacterCount,
            textProfile.UppercaseCharacterCount,
            textProfile.LowercaseCharacterCount,
            textProfile.LongestBreakableTokenLength,
            bodyElements.Count,
            bodyElements.OfType<DocxManualBreakElement>().Count(),
            bodyElements.OfType<DocxPageBreakElement>().Count(),
            bodyElements.OfType<DocxTableElement>().Count());
    }

    private static IReadOnlyList<DocxParagraph> GetParagraphsFromBodyElements(IReadOnlyList<DocxBodyElement> bodyElements)
    {
        return bodyElements
            .OfType<DocxParagraphElement>()
            .Select(element => element.Paragraph)
            .ToArray();
    }

    private static TextProfile BuildTextProfile(string text)
    {
        int spaceCharacterCount = 0;
        int nonAsciiCharacterCount = 0;
        int punctuationCharacterCount = 0;
        int digitCharacterCount = 0;
        int uppercaseCharacterCount = 0;
        int lowercaseCharacterCount = 0;
        foreach (char value in text)
        {
            if (char.IsWhiteSpace(value))
            {
                spaceCharacterCount++;
            }

            if (value > 0x7f)
            {
                nonAsciiCharacterCount++;
            }

            if (char.IsPunctuation(value))
            {
                punctuationCharacterCount++;
            }

            if (char.IsDigit(value))
            {
                digitCharacterCount++;
            }

            if (char.IsUpper(value))
            {
                uppercaseCharacterCount++;
            }

            if (char.IsLower(value))
            {
                lowercaseCharacterCount++;
            }
        }

        int longestBreakableTokenLength = ResolveLongestBreakableTokenLength(text);
        return new TextProfile(
            spaceCharacterCount,
            nonAsciiCharacterCount,
            punctuationCharacterCount,
            digitCharacterCount,
            uppercaseCharacterCount,
            lowercaseCharacterCount,
            longestBreakableTokenLength);
    }

    private static int ResolveLongestBreakableTokenLength(string text)
    {
        int longest = 0;
        int current = 0;
        foreach (char value in text)
        {
            if (char.IsWhiteSpace(value) &&
                value != '\u00A0' &&
                value != '\u202F' &&
                value != '\u2007')
            {
                longest = Math.Max(longest, current);
                current = 0;
            }
            else
            {
                current++;
            }
        }

        return Math.Max(longest, current);
    }

    private static bool HasBeforeSpacingToken(DocxParagraphSpacing spacing)
    {
        return spacing.BeforeValue is not null ||
            spacing.BeforeLinesValue is not null ||
            spacing.BeforeAutoSpacingValue is not null;
    }

    private static bool HasAfterSpacingToken(DocxParagraphSpacing spacing)
    {
        return spacing.AfterValue is not null ||
            spacing.AfterLinesValue is not null ||
            spacing.AfterAutoSpacingValue is not null;
    }
}

internal sealed record DocxLayoutPageSnapshot(
    double Width,
    double Height,
    double MarginLeft,
    double MarginRight,
    double MarginTop,
    double MarginBottom,
    string? SectionBreakTypeValue,
    string? SectionColumnCountValue,
    string? SectionColumnEqualWidthValue,
    string? SectionColumnSpaceValue,
    int? SectionColumnCount,
    double? SectionColumnSpacePoints,
    int SectionColumnDefinitionCount,
    double SectionColumnDefinitionWidthSum,
    double SectionColumnDefinitionSpaceSum,
    int ColumnFrameCount,
    double ColumnFrameWidthSum,
    double ColumnGutterWidthSum,
    IReadOnlyList<DocxLayoutColumnFrameSnapshot> ColumnFrames,
    int ItemCount,
    int StaticTextLineCount,
    int StaticInlineImageCount,
    int StaticTableRowCount,
    int TextLineCount,
    int InlineImageCount,
    int TableRowCount,
    int SourceBlockCount,
    int? FirstSourceBlockIndex,
    int? LastSourceBlockIndex,
    double VerticalUsed,
    double TextLineHeightSum,
    double InlineImageHeightSum,
    double TableRowHeightSum,
    IReadOnlyList<DocxStaticStoryLayoutSnapshot> StaticStories,
    IReadOnlyList<DocxLayoutItemSnapshot> StaticItems,
    IReadOnlyList<DocxLayoutItemSnapshot> Items,
    IReadOnlyList<DocxTableRowSnapshot> TableRows);

internal sealed record DocxLayoutColumnFrameSnapshot(
    int Index,
    double X,
    double Width,
    double? GutterAfterPoints);

internal sealed record DocxFloatingDrawingLayoutSnapshot(
    int Index,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    int? PageStartIndex,
    int? PageEndIndex,
    int? AnchorPageIndex,
    int? AnchorColumnIndex,
    double? AnchorBlockVerticalTop,
    double? AnchorBlockVerticalBottom,
    double? ExtentWidthPoints,
    double? ExtentHeightPoints,
    double? HorizontalOffsetPoints,
    double? VerticalOffsetPoints,
    double? DistanceTopPoints,
    double? DistanceBottomPoints,
    double? DistanceLeftPoints,
    double? DistanceRightPoints,
    double? HorizontalReferenceX,
    double? HorizontalReferenceWidth,
    double? VerticalReferenceTop,
    double? VerticalReferenceBottom,
    double? PlacedX,
    double? PlacedTop,
    string? HorizontalPlacementSource,
    string? VerticalPlacementSource,
    double? WrapExclusionX,
    double? WrapExclusionTop,
    double? WrapExclusionWidth,
    double? WrapExclusionHeight,
    string? WrapKind,
    string? WrapTextValue,
    string? HorizontalRelativeFromValue,
    string? HorizontalAlignValue,
    string? HorizontalOffsetValue,
    string? VerticalRelativeFromValue,
    string? VerticalAlignValue,
    string? VerticalOffsetValue,
    string? ImageRelationshipId,
    string? ImagePartName,
    string? ImageContentType,
    double? ImageWidthPoints,
    double? ImageHeightPoints,
    string? StoryKind = null,
    string? StoryVariantType = null);

internal sealed record DocxStaticStoryLayoutSnapshot(
    string Kind,
    string? VariantType,
    int TextLineCount,
    int InlineImageCount,
    int TableRowCount,
    int ParagraphCount,
    int SourceLineCount,
    int TextLength,
    int FirstParagraphLineCount,
    double VerticalTop,
    double VerticalBottom,
    int? FirstSourceParagraphIndex,
    int? LastSourceParagraphIndex,
    int? FirstSourceLineIndex,
    int? LastSourceLineIndex,
    IReadOnlyList<DocxLayoutItemSnapshot> Items);

internal sealed record DocxRelatedStoryLayoutSnapshot(
    int StoryIndex,
    string Kind,
    string PartName,
    string? Id,
    int BlockCount,
    int ParagraphCount,
    int TableCount,
    int TextLineCount,
    int TableCellTextLineCount,
    int TableRowCount,
    int InlineImageCount,
    int FloatingDrawingCount,
    int TextLength,
    double ContentHeight,
    IReadOnlyList<DocxLayoutItemSnapshot> Items,
    IReadOnlyList<DocxRelatedStorySourceBlockSnapshot> SourceBlocks,
    IReadOnlyList<DocxTableRowSnapshot> TableRows);

internal sealed record DocxRelatedStorySourceBlockSnapshot(
    int SourceBlockIndex,
    string Kind,
    int ItemCount,
    int TextLineCount,
    int InlineImageCount,
    int TableRowCount,
    int TextLength,
    double VerticalTop,
    double VerticalBottom,
    double ConsumedHeight,
    double AppliedBeforeSpacingSum);

internal sealed record DocxLayoutSourceBlockSnapshot(
    int SourceBlockIndex,
    string Kind,
    int FirstPageIndex,
    int LastPageIndex,
    int? FirstColumnIndex,
    int? LastColumnIndex,
    double VerticalTop,
    double VerticalBottom,
    int ItemCount,
    int TextLineCount,
    int TableRowCount,
    int InlineImageCount,
    int TextLength,
    double ConsumedHeightSum,
    double AppliedBeforeSpacingSum);

internal sealed record DocxLayoutItemSnapshot(
    string Kind,
    double X,
    double Y,
    double Width,
    double Height,
    int TextLength,
    int CellCount,
    int? ColumnIndex,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    int? SourceLineIndex,
    double? LineHeightPoints,
    double? AppliedBeforeSpacingPoints,
    double? SingleLineHeightPoints,
    double? ListLabelSingleLineHeightPoints,
    double? BodyWindowsLineHeightPoints,
    double? ListLabelWindowsLineHeightPoints,
    double? EffectiveLineSpacingFactor,
    bool? LineSpacingFactorFloorApplied,
    bool? IsFirstParagraphLine,
    double? PendingAfterSpacingPoints,
    double? ParagraphBeforeSpacingPoints,
    double? ParagraphAfterSpacingPoints,
    bool? ContextualSpacingSuppressed,
    string? StoryVariantType = null,
    IReadOnlyList<DocxLayoutItemSnapshot>? TextLines = null,
    string? ParagraphStyleId = null,
    bool? ParagraphStyleFound = null,
    int? ParagraphStyleDepth = null,
    bool? HasDocumentDefaultParagraphProperties = null,
    bool? HasDirectParagraphProperties = null,
    bool? HasTableStyleParagraphProperties = null,
    int? CharacterStyleTextSegmentCount = null,
    int? DirectRunPropertyTextSegmentCount = null,
    int? ParagraphStyleRunPropertyTextSegmentCount = null,
    int? TableStyleRunPropertyTextSegmentCount = null,
    int? DocumentDefaultRunPropertyTextSegmentCount = null);

internal sealed record DocxTableRowSnapshot(
    int TableIndex,
    int SourceBlockIndex,
    int PageRowIndex,
    int RowIndex,
    int TableRowCount,
    int GridColumnCount,
    double GridColumnsWidthSum,
    bool HasExplicitGrid,
    IReadOnlyList<double> ResolvedColumnWidths,
    double ResolvedTableWidth,
    double TableX,
    double? PreferredTableWidthPoints,
    string? PreferredTableWidthValue,
    string? PreferredTableWidthType,
    double? IndentPoints,
    double? CellSpacingPoints,
    string? LayoutValue,
    string? HeightValue,
    string? HeightRuleValue,
    double? DeclaredHeightPoints,
    int FragmentIndex,
    int FragmentCount,
    string FragmentReason,
    double FullRowHeight,
    double FragmentOffsetFromRowTop,
    double X,
    double Y,
    double Width,
    double Height,
    double BottomOverflowPoints,
    double? FirstBaselineY,
    double? LastBaselineY,
    int CellCount,
    int TextLineCount,
    int TextLength,
    double MaxFontSize,
    bool IsHeader,
    string? HeaderValue,
    bool HasTablePropertyExceptionCellMargins,
    bool CantSplit,
    string? CantSplitValue,
    string? StoryKind,
    string? StoryVariantType,
    IReadOnlyList<DocxTableCellSnapshot> Cells);

internal sealed record DocxTableSnapshot(
    int TableIndex,
    int SourceBlockIndex,
    int PageStartIndex,
    int PageEndIndex,
    int RowCount,
    int LaidOutRowCount,
    int HeaderRowLayoutCount,
    int AuthoredHeaderRowCount,
    int GridColumnCount,
    double GridColumnsWidthSum,
    bool HasExplicitGrid,
    IReadOnlyList<double> ResolvedColumnWidths,
    double ResolvedTableWidth,
    double X,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    double? IndentPoints,
    double? CellSpacingPoints,
    string? LayoutValue,
    int DeclaredHeightRowCount,
    int ExactHeightRowCount,
    int AtLeastHeightRowCount,
    int CantSplitRowCount,
    int FragmentedRowCount,
    int FragmentedRowLayoutCount,
    int MaxRowFragmentCount,
    bool HasVerticalMerge,
    int AuthoredVerticalMergeCellCount,
    int AuthoredVerticalMergeRestartCellCount,
    int AuthoredVerticalMergeContinuationCellCount,
    int LaidOutVerticalMergeContinuationCellCount,
    int MissingVerticalMergeOwnerCellCount,
    string? StoryKind = null,
    string? StoryVariantType = null);

internal sealed record DocxTableCellSnapshot(
    int CellIndex,
    double X,
    double Y,
    double Width,
    double Height,
    double ContentBoxX,
    double ContentBoxY,
    double ContentBoxWidth,
    double ContentBoxHeight,
    double ResolvedPaddingLeftPoints,
    double ResolvedPaddingTopPoints,
    double ResolvedPaddingRightPoints,
    double ResolvedPaddingBottomPoints,
    int TextLineCount,
    int TextLength,
    double MaxFontSize,
    double? FirstTextLineX,
    double? FirstBaselineY,
    double FirstBaselineInset,
    double? LastBaselineY,
    int InlineImageCount,
    int ParagraphCount,
    int ParagraphsWithBeforeSpacingToken,
    int ParagraphsWithAfterSpacingToken,
    int ParagraphsWithZeroBeforeSpacing,
    int ParagraphsWithZeroAfterSpacing,
    double? MinSpacingBeforePoints,
    double? MaxSpacingBeforePoints,
    double? MinSpacingAfterPoints,
    double? MaxSpacingAfterPoints,
    int GridSpan,
    string? GridSpanValue,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    string? VerticalAlignmentValue,
    double? MarginTopPoints,
    double? MarginRightPoints,
    double? MarginBottomPoints,
    double? MarginLeftPoints,
    int BorderCount,
    bool HasFill,
    bool HasShadingValue,
    bool HasConditionalFormat,
    bool HasVerticalMerge,
    string? VerticalMergeValue,
    bool IsVerticalMergeContinuation,
    string VisualOwnership,
    int? VerticalMergeOwnerRowIndex,
    int? VerticalMergeOwnerGridColumnIndex,
    int VisualParagraphCount,
    int VisualTextLength,
    int VisualInlineImageCount,
    int SpaceCharacterCount,
    int NonAsciiCharacterCount,
    int PunctuationCharacterCount,
    int DigitCharacterCount,
    int UppercaseCharacterCount,
    int LowercaseCharacterCount,
    int LongestBreakableTokenLength,
    int BodyElementCount = 0,
    int ManualBreakElementCount = 0,
    int PageBreakElementCount = 0,
    int NestedTableElementCount = 0);

internal sealed record DocxTextEmissionSnapshot(
    int LineCount,
    int SegmentCount,
    int TerminalSpaceSegmentCount,
    int NonzeroPdfCharacterSpacingSegmentCount,
    int CompensatedCharacterSpacingSegmentCount,
    IReadOnlyList<DocxTextEmissionLineSnapshot> Lines);

internal sealed record DocxTextEmissionLineSnapshot(
    int PageIndex,
    bool IsStaticStory,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    int? SourceLineIndex,
    bool EndsWithIntraTokenBreak,
    int SegmentCount,
    int TextLength,
    int TerminalSpaceSegmentCount,
    int NonzeroPdfCharacterSpacingSegmentCount,
    IReadOnlyList<DocxTextEmissionSegmentSnapshot> Segments);

internal sealed record DocxTextEmissionSegmentSnapshot(
    int TextLength,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    int? SourceLineIndex,
    string Role,
    double X,
    double BaselineY,
    double Width,
    double FontSize,
    double PdfFontSize,
    double LayoutCharacterSpacing,
    double PdfCharacterSpacing,
    string PdfCharacterSpacingSource,
    double PositioningCharacterSpacing,
    bool CompensatePdfCharacterSpacing,
    DocxTextEmissionCharacterProfile CharacterProfile,
    DocxTextEmissionAdvanceProfile AdvanceProfile,
    DocxTextEmissionGlyphAdvanceSignature GlyphAdvanceSignature,
    bool IsTerminalLineSpace,
    string? FontResourceName,
    bool SyntheticBold,
    bool SyntheticItalic,
    string? CharacterStyleId = null,
    bool CharacterStyleFound = false,
    int CharacterStyleDepth = 0,
    bool HasDocumentDefaultRunProperties = false,
    bool HasParagraphStyleRunProperties = false,
    bool HasCharacterStyleRunProperties = false,
    bool HasDirectRunProperties = false,
    bool HasTableStyleRunProperties = false);

internal sealed record TextProfile(
    int SpaceCharacterCount,
    int NonAsciiCharacterCount,
    int PunctuationCharacterCount,
    int DigitCharacterCount,
    int UppercaseCharacterCount,
    int LowercaseCharacterCount,
    int LongestBreakableTokenLength);

internal sealed record DocxLayoutPage(
    double Width,
    double Height,
    double MarginLeft,
    double MarginRight,
    double MarginTop,
    double MarginBottom,
    DocxPageSettings PageSettings,
    DocxSectionLayoutProperties SectionProperties,
    IReadOnlyList<DocxLayoutColumnFrame> ColumnFrames,
    IReadOnlyList<DocxTextLineLayout> StaticTextLines,
    IReadOnlyList<DocxInlineImageLayout> StaticInlineImages,
    IReadOnlyList<DocxTableRowLayout> StaticTableRows,
    IReadOnlyList<DocxLayoutItem> Items);

internal abstract record DocxLayoutItem;

internal sealed record DocxTextLineLayout(
    string Text,
    DocxTextRun StyleRun,
    double FontSize,
    double X,
    double BaselineY,
    double Width,
    IReadOnlyList<DocxTextSegmentLayout> Segments,
    int? SourceBlockIndex = null,
    int? SourceParagraphIndex = null,
    int? SourceLineIndex = null,
    string? StoryKind = null,
    double? LineHeight = null,
    double? AppliedBeforeSpacing = null,
    bool? IsFirstParagraphLine = null,
    bool EndsWithIntraTokenBreak = false,
    double? SingleLineHeight = null,
    double? ListLabelSingleLineHeight = null,
    double? BodyWindowsLineHeight = null,
    double? ListLabelWindowsLineHeight = null,
    double? EffectiveLineSpacingFactor = null,
    bool? LineSpacingFactorFloorApplied = null,
    double? PendingAfterSpacing = null,
    double? ParagraphBeforeSpacing = null,
    double? ParagraphAfterSpacing = null,
    bool? ContextualSpacingSuppressed = null,
    DocxParagraph? SourceParagraph = null,
    string? StoryVariantType = null) : DocxLayoutItem;

internal sealed record DocxTextSegmentLayout(
    string Text,
    DocxTextRun StyleRun,
    double X,
    double Width,
    double? FontSize = null,
    double BaselineOffsetY = 0d,
    double PdfCharacterSpacing = 0d,
    DocxTextStateCharacterSpacingSource PdfCharacterSpacingSource = DocxTextStateCharacterSpacingSource.None,
    bool CompensatePdfCharacterSpacing = true,
    int SourceTextRunIndex = -1,
    DocxTextSegmentRole Role = DocxTextSegmentRole.Text);

internal enum DocxTextSegmentRole
{
    Text,
    ListLabel,
    ListSeparator
}

internal sealed record DocxTextSpan(
    string Text,
    DocxTextRun StyleRun,
    int SourceTextRunIndex);

internal sealed record DocxWrappedTextLine(
    string Text,
    IReadOnlyList<DocxTextSpan> Spans,
    bool EndsWithIntraTokenBreak = false);

internal sealed record DocxInlineImageLayout(
    DocxInlineImage Image,
    double X,
    double Y,
    double Width,
    double Height,
    int PageIndex,
    int? SourceBlockIndex = null,
    int? SourceParagraphIndex = null,
    string? StoryKind = null,
    string? StoryVariantType = null) : DocxLayoutItem;

internal sealed record DocxTableRowLayout(
    DocxTableLayoutContext Table,
    int RowIndex,
    int FragmentIndex,
    int FragmentCount,
    string FragmentReason,
    double FullRowHeight,
    double FragmentOffsetFromRowTop,
    IReadOnlyList<DocxTableCellLayout> Cells,
    double Y,
    double Height,
    double? DeclaredHeightPoints,
    string? HeightValue,
    string? HeightRuleValue,
    bool IsHeader,
    string? HeaderValue,
    bool HasTablePropertyExceptionCellMargins,
    bool CantSplit,
    string? CantSplitValue,
    string? StoryKind = null,
    string? StoryVariantType = null) : DocxLayoutItem;

internal sealed record DocxTableLayoutContext(
    int TableIndex,
    int SourceBlockIndex,
    int RowCount,
    int GridColumnCount,
    double GridColumnsWidthSum,
    bool HasExplicitGrid,
    IReadOnlyList<double> ResolvedColumnWidths,
    double ResolvedTableWidth,
    double TableX,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    double? IndentPoints,
    double? CellSpacingPoints,
    string? LayoutValue);

internal sealed record DocxTableCellLayout(
    DocxTableCell Cell,
    double X,
    double Y,
    double Width,
    double Height,
    double ContentPaddingLeft,
    double ContentPaddingTop,
    double ContentPaddingRight,
    double ContentPaddingBottom,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    bool IsVerticalMergeContinuation = false,
    DocxTableCell? VerticalMergeOwnerCell = null,
    DocxVerticalMergeOwner? VerticalMergeOwner = null,
    DocxTableCellVisualOwnership VisualOwnership = DocxTableCellVisualOwnership.OwnCell,
    IReadOnlyList<DocxTableRowLayout>? NestedTableRows = null)
{
    public IReadOnlyList<DocxTableRowLayout> NestedRows => NestedTableRows ?? [];

    public DocxTableCell VisualCell =>
        VisualOwnership == DocxTableCellVisualOwnership.VerticalMergeOwner && VerticalMergeOwner is not null
            ? VerticalMergeOwner.Cell
            : Cell;
}

internal sealed record DocxVerticalMergeOwner(DocxTableCell Cell, int RowIndex, int GridColumnIndex);

internal enum DocxTableCellVisualOwnership
{
    OwnCell,
    VerticalMergeOwner,
    MissingVerticalMergeOwner
}

internal sealed record DocxRunFontResource(string Name, PdfEmbeddedFont Embedded, FontFaceResolution Resolution);

internal sealed record DocxFontResources(
    DocxFontPlan Plan,
    IDocxTextMeasurer? TextMeasurer,
    IReadOnlyList<PdfFontResource> Resources,
    IReadOnlyDictionary<DocxTextRun, DocxRunFontResource> RunResources,
    DocxRunFontResource? Fallback);

internal sealed record DocxTextEmissionSegment(
    string Text,
    DocxTextRun StyleRun,
    DocxRunFontResource Resource,
    RgbColor Color,
    double X,
    double BaselineY,
    double Width,
    double FontSize,
    double PdfCharacterSpacing,
    DocxTextStateCharacterSpacingSource PdfCharacterSpacingSource,
    bool CompensatePdfCharacterSpacing,
    bool SyntheticBold,
    bool SyntheticItalic,
    bool IsTerminalLineSpace,
    int SourceTextRunIndex = -1,
    DocxTextSegmentRole Role = DocxTextSegmentRole.Text);

internal readonly record struct DocxKeepBlockEstimate(
    double Height,
    int ParagraphCount,
    int FirstTableRowCount);

internal interface IDocxTextMeasurer
{
    double MeasureText(DocxTextRun? run, string text, double fontSize);
}

internal interface IDocxLineMetricsProvider
{
    double MeasureSingleLineHeight(DocxTextRun? run, double fontSize);
}

internal interface IDocxStaticTextMetricsProvider
{
    double MeasureWindowsAscender(DocxTextRun? run, double fontSize);

    double MeasureWindowsDescender(DocxTextRun? run, double fontSize);
}

internal static class DocxTextSpacing
{
    public static double AddCharacterSpacing(double measuredWidth, DocxTextRun? run, string text)
    {
        return measuredWidth + CountCharacterSpacingGaps(text) * (run?.CharacterSpacingPoints ?? 0d);
    }

    public static double BoundarySpacing(DocxTextRun? left, string leftText, string rightText)
    {
        return leftText.Length == 0 ||
            rightText.Length == 0 ||
            leftText[^1] == '\t' ||
            rightText[0] == '\t'
            ? 0d
            : left?.CharacterSpacingPoints ?? 0d;
    }

    private static int CountCharacterSpacingGaps(string text)
    {
        int count = 0;
        foreach (Rune _ in text.EnumerateRunes())
        {
            count++;
        }

        return Math.Max(0, count - 1);
    }
}

internal static class DocxLineMetrics
{
    private const double WordSingleLineMinimumEm = 1.15d;
    private const double WordAutoLineBaselineOffsetEm = 0.94d;
    private const double WordExactLineTextBottomInsetEm = 0.299d;

    public static double MeasureOpenTypeSingleLineHeight(OpenTypeFont font, double fontSize)
    {
        if (font.UnitsPerEm == 0)
        {
            return fontSize * WordSingleLineMinimumEm;
        }

        double units = font.Os2.TypographicAscender - font.Os2.TypographicDescender + font.Os2.TypographicLineGap;

        return Math.Max(fontSize * WordSingleLineMinimumEm, units * fontSize / font.UnitsPerEm);
    }

    public static double MeasureWindowsAscender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? fontSize
            : font.Os2.WindowsAscender * fontSize / font.UnitsPerEm;
    }

    public static double MeasureWindowsDescender(OpenTypeFont font, double fontSize)
    {
        return font.UnitsPerEm == 0
            ? 0d
            : font.Os2.WindowsDescender * fontSize / font.UnitsPerEm;
    }

    public static double ResolveBodyBaselineOffset(double fontSize, double lineHeight, bool hasExplicitLineSpacing)
    {
        return hasExplicitLineSpacing
            ? Math.Max(0d, lineHeight - fontSize * WordExactLineTextBottomInsetEm)
            : fontSize * WordAutoLineBaselineOffsetEm;
    }

    public static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs)
    {
        DocxParagraph? firstTextParagraph = paragraphs.FirstOrDefault(paragraph => paragraph.Runs.Count != 0);
        return firstTextParagraph is null ? 0d : firstTextParagraph.Runs.Max(run => run.FontSize);
    }
}

internal static class DocxVerticalAlignMetrics
{
    private const double SuperscriptSubscriptScale = 2d / 3d;
    private const double HalfPointGrid = 2d;
    private const double SubscriptBaselineShiftEm = -0.06d;

    public static double ResolveFontSize(double nominalFontSize, DocxTextRun run)
    {
        if (!IsSuperscript(run) && !IsSubscript(run))
        {
            return nominalFontSize;
        }

        return Math.Max(0.5d, Math.Floor(nominalFontSize * SuperscriptSubscriptScale * HalfPointGrid) / HalfPointGrid);
    }

    public static double ResolveBaselineOffset(double nominalFontSize, double layoutFontSize, DocxTextRun run)
    {
        if (IsSuperscript(run))
        {
            return Math.Max(0d, nominalFontSize - layoutFontSize);
        }

        return IsSubscript(run) ? nominalFontSize * SubscriptBaselineShiftEm : 0d;
    }

    private static bool IsSuperscript(DocxTextRun run)
    {
        return run.VerticalAlignmentValue?.Equals("superscript", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool IsSubscript(DocxTextRun run)
    {
        return run.VerticalAlignmentValue?.Equals("subscript", StringComparison.OrdinalIgnoreCase) == true;
    }
}

internal sealed class DocxEmbeddedTextMeasurer(PdfEmbeddedFont embedded) : IDocxTextMeasurer, IDocxLineMetricsProvider, IDocxStaticTextMetricsProvider
{
    public double MeasureText(DocxTextRun? run, string text, double fontSize)
    {
        return DocxTextSpacing.AddCharacterSpacing(embedded.MeasureTextPoints(text, fontSize), run, text);
    }

    public double MeasureSingleLineHeight(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureOpenTypeSingleLineHeight(embedded.Font, fontSize);
    }

    public double MeasureWindowsAscender(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureWindowsAscender(embedded.Font, fontSize);
    }

    public double MeasureWindowsDescender(DocxTextRun? run, double fontSize)
    {
        return DocxLineMetrics.MeasureWindowsDescender(embedded.Font, fontSize);
    }
}

internal sealed class DocxLayoutEngine
{
    private const double WordDefaultTabStopPoints = 36d;
    private const double InlineImageParagraphGapPoints = 6d;
    private const double WordListMinimumAutoLineSpacingFactor = 1.19d;
    private const double UnpagedRelatedStoryCanvasHeightPoints = 100000d;

    private sealed record DocxPageGeometry(
        double Width,
        double Height,
        double MarginLeft,
        double MarginRight,
        double MarginTop,
        double MarginBottom,
        DocxPageSettings PageSettings,
        DocxSectionLayoutProperties SectionProperties,
        IReadOnlyList<DocxLayoutColumnFrame> ColumnFrames)
    {
        public double BodyWidth => Math.Max(1d, Width - MarginLeft - MarginRight);
    }

    private sealed record DocxEffectiveSectionSettings(
        DocxPageSettings PageSettings,
        DocxSectionLayoutProperties SectionProperties);

    private sealed record DocxTableLayoutFrame(
        DocxTableLayoutContext Context,
        IReadOnlyList<double> EffectiveColumns,
        double Scale,
        IReadOnlyList<double> RowHeights,
        double PageContentHeight,
        double TableX);

    public DocxLayout Create(DocxDocument document, PdfEmbeddedFont? embedded)
    {
        IDocxTextMeasurer? textMeasurer = embedded is null ? null : new DocxEmbeddedTextMeasurer(embedded);
        return Create(document, textMeasurer);
    }

    internal DocxLayout Create(DocxDocument document, IDocxTextMeasurer? textMeasurer)
    {
        var pages = new List<DocxLayoutPage>();
        var currentItems = new List<DocxLayoutItem>();
        IReadOnlyDictionary<int, DocxEffectiveSectionSettings> sectionSettingsByElementIndex = BuildEffectiveSectionSettings(document, out DocxEffectiveSectionSettings finalSectionSettings);
        DocxPageGeometry page = ResolveSectionGeometry(document, FindSectionSettingsAtOrAfter(document.BodyElements, 0, sectionSettingsByElementIndex) ?? finalSectionSettings);
        int activeColumnIndex = 0;
        double x = ResolveActiveColumnFrame(page, activeColumnIndex).X;
        double width = ResolveActiveColumnFrame(page, activeColumnIndex).Width;
        double cursorY = page.Height - page.MarginTop;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        bool activeColumnHasContent = false;
        int tableIndex = 0;
        double defaultTabStopPoints = document.Settings.DefaultTabStopPoints ?? WordDefaultTabStopPoints;

        void ApplyActiveColumnFrame()
        {
            DocxLayoutColumnFrame frame = ResolveActiveColumnFrame(page, activeColumnIndex);
            x = frame.X;
            width = frame.Width;
        }

        void FinishPage()
        {
            pages.Add(new DocxLayoutPage(
                page.Width,
                page.Height,
                page.MarginLeft,
                page.MarginRight,
                page.MarginTop,
                page.MarginBottom,
                page.PageSettings,
                page.SectionProperties,
                page.ColumnFrames,
                [],
                [],
                [],
                currentItems.ToArray()));
            currentItems = [];
            activeColumnIndex = 0;
            ApplyActiveColumnFrame();
            cursorY = page.Height - page.MarginTop;
            pendingSpacingAfter = 0d;
            previousParagraph = null;
            activeColumnHasContent = false;
        }

        void AdvanceColumnOrPage()
        {
            if (activeColumnIndex + 1 < page.ColumnFrames.Count)
            {
                activeColumnIndex++;
                ApplyActiveColumnFrame();
                cursorY = page.Height - page.MarginTop;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                activeColumnHasContent = false;
                return;
            }

            FinishPage();
        }

        void ApplySectionAfterBreak(int elementIndex)
        {
            page = ResolveSectionGeometry(document, FindSectionSettingsAtOrAfter(document.BodyElements, elementIndex + 1, sectionSettingsByElementIndex) ?? finalSectionSettings);
            activeColumnIndex = 0;
            ApplyActiveColumnFrame();
            activeColumnHasContent = false;
            if (!HasPageContent())
            {
                cursorY = page.Height - page.MarginTop;
            }
        }

        bool HasPageContent() => currentItems.Count > 0;
        bool HasCurrentColumnContent() => activeColumnHasContent;

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            DocxBodyElement element = document.BodyElements[elementIndex];
            if (element is DocxPageBreakElement pageBreak)
            {
                if (pageBreak.BreakParagraph is { } breakParagraph)
                {
                    DocxParagraphSpacingProfile breakSpacingProfile = ResolveParagraphSpacingProfile(previousParagraph, breakParagraph, pendingSpacingAfter);
                    double breakFontSize = GetParagraphFontSize(breakParagraph);
                    double breakLineHeight = ResolveLineHeight(breakParagraph, breakFontSize, textMeasurer);
                    double paragraphAdvance = breakSpacingProfile.AppliedBeforeSpacing + breakLineHeight;
                    if (cursorY - paragraphAdvance < page.MarginBottom && HasCurrentColumnContent())
                    {
                        AdvanceColumnOrPage();
                    }

                    cursorY -= paragraphAdvance;
                    activeColumnHasContent = true;
                }

                if (HasPageContent() || pageBreak.BreakParagraph is not null)
                {
                    FinishPage();
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxManualBreakElement manualBreak)
            {
                if (manualBreak.Value?.Equals("column", StringComparison.OrdinalIgnoreCase) == true)
                {
                    AdvanceColumnOrPage();
                    continue;
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxSectionBreakElement sectionBreak)
            {
                bool startsNewPage = ShouldStartNewPageForSectionBreak(sectionBreak);
                if (startsNewPage && HasPageContent())
                {
                    FinishPage();
                }

                if (ShouldInsertParityBlankPage(sectionBreak, pages.Count + 1))
                {
                    FinishPage();
                }

                if (startsNewPage || (IsContinuousSectionBreak(sectionBreak) && !HasPageContent()))
                {
                    ApplySectionAfterBreak(elementIndex);
                }

                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            if (element is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                int itemCountBeforeTable = currentItems.Count;
                int currentTableIndex = tableIndex++;
                bool tableActiveFrameHasContent = false;
                void MarkTableBoundaryContent()
                {
                    tableActiveFrameHasContent = true;
                }

                Action advanceTableBoundary = page.ColumnFrames.Count > 1
                    ? () =>
                    {
                        AdvanceColumnOrPage();
                        tableActiveFrameHasContent = false;
                    }
                    : FinishPage;
                Func<bool> hasTableBoundaryContent = page.ColumnFrames.Count > 1
                    ? () => HasCurrentColumnContent() || tableActiveFrameHasContent
                    : HasPageContent;
                DocxTableLayoutFrame ResolveCurrentTableFrame()
                {
                    return CreateTableLayoutFrame(
                        tableElement.Table,
                        currentTableIndex,
                        elementIndex,
                        x,
                        width,
                        page.Height - page.MarginTop - page.MarginBottom,
                        textMeasurer,
                        defaultTabStopPoints);
                }

                LayoutTable(tableElement.Table, page.MarginBottom, textMeasurer, defaultTabStopPoints, () => pages.Count + 1, ref currentItems, ref cursorY, ResolveCurrentTableFrame, advanceTableBoundary, hasTableBoundaryContent, MarkTableBoundaryContent);
                if (currentItems.Count > itemCountBeforeTable)
                {
                    activeColumnHasContent = true;
                }

                continue;
            }

            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double paragraphFontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, paragraphFontSize, textMeasurer);
            double lineHeight = lineHeightProfile.LineHeight;
            if (textMeasurer is not null &&
                HasPageContent() &&
                ShouldKeepParagraphBlockTogether(paragraph) &&
                cursorY - EstimateKeptParagraphBlock(document.BodyElements, elementIndex, width, textMeasurer, defaultTabStopPoints).Height <= page.MarginBottom)
            {
                AdvanceColumnOrPage();
            }

            IReadOnlyList<DocxTextSpan> textSpans = textMeasurer is null ? [] : CreateTextSpans(paragraph.Runs);
            if (textMeasurer is not null && textSpans.Count > 0)
            {
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, paragraphFontSize, textMeasurer);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
                double paragraphX = x + textStartOffset;
                double paragraphWidth = Math.Max(1d, width - textStartOffset - GetParagraphRightInset(paragraph));
                DocxTextRun firstRun = paragraph.Runs[0];
                bool firstLine = true;
                double continuationParagraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: false).ToArray();
                if (ShouldMoveParagraphForWidowControl(paragraph, lines.Length, cursorY, lineHeight, page.MarginBottom, HasCurrentColumnContent()))
                {
                    AdvanceColumnOrPage();
                }

                for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    DocxWrappedTextLine line = lines[lineIndex];
                    if (cursorY - lineHeight < page.MarginBottom && HasCurrentColumnContent())
                    {
                        AdvanceColumnOrPage();
                    }

                    double lineWidth = MeasureTextSpans(line.Spans, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double drawableLineWidth = MeasureDrawableTextSpans(line.Spans, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double lineX = paragraph.Alignment switch
                    {
                        DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                        DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                        _ => paragraphX
                    };
                    double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(paragraphFontSize, lineHeight, paragraph.LineSpacingPoints is not null);
                    bool justifyLine = (paragraph.ListLabel is null || !firstLine) &&
                        ShouldJustifyTextLine(paragraph.Alignment, lineIndex == lines.Length - 1, drawableLineWidth, paragraphWidth, line.Spans);
                    IReadOnlyList<DocxTextSegmentLayout> segments = firstLine && paragraph.ListLabel is not null
                        ? CreateNumberedLineSegments(paragraph.ListLabel, line.Spans, firstRun, x + labelStartOffset, lineX, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                        : justifyLine
                            ? CreateJustifiedTextSegments(line.Spans, lineX, drawableLineWidth, paragraphWidth, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                            : CreateTextSegments(line.Spans, lineX, paragraphFontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double effectiveX = firstLine && paragraph.ListLabel is not null ? x + labelStartOffset : lineX;
                    double effectiveWidth = firstLine && paragraph.ListLabel is not null
                        ? Math.Max(lineX + lineWidth, x + labelStartOffset + MeasureListLabel(paragraph.ListLabel, firstRun, paragraphFontSize, textMeasurer)) - (x + labelStartOffset)
                        : justifyLine
                            ? paragraphWidth
                            : lineWidth;
                    currentItems.Add(new DocxTextLineLayout(
                        firstLine && paragraph.ListLabel is not null ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line.Text : line.Text,
                        firstRun,
                        paragraphFontSize,
                        effectiveX,
                        cursorY - baselineOffset,
                        effectiveWidth,
                        segments,
                        SourceBlockIndex: elementIndex,
                        SourceParagraphIndex: 0,
                        SourceLineIndex: lineIndex,
                        LineHeight: lineHeight,
                        AppliedBeforeSpacing: firstLine ? spacingProfile.AppliedBeforeSpacing : 0d,
                        IsFirstParagraphLine: firstLine,
                        EndsWithIntraTokenBreak: line.EndsWithIntraTokenBreak,
                        SingleLineHeight: lineHeightProfile.SingleLineHeight,
                        ListLabelSingleLineHeight: lineHeightProfile.ListLabelSingleLineHeight,
                        BodyWindowsLineHeight: lineHeightProfile.BodyWindowsLineHeight,
                        ListLabelWindowsLineHeight: lineHeightProfile.ListLabelWindowsLineHeight,
                        EffectiveLineSpacingFactor: lineHeightProfile.EffectiveLineSpacingFactor,
                        LineSpacingFactorFloorApplied: lineHeightProfile.LineSpacingFactorFloorApplied,
                        PendingAfterSpacing: firstLine ? spacingProfile.PendingAfterSpacing : null,
                        ParagraphBeforeSpacing: firstLine ? spacingProfile.ParagraphBeforeSpacing : null,
                        ParagraphAfterSpacing: firstLine ? spacingProfile.ParagraphAfterSpacing : null,
                        ContextualSpacingSuppressed: firstLine ? spacingProfile.ContextualSpacingSuppressed : null,
                        SourceParagraph: paragraph,
                        StoryKind: "Body"));
                    activeColumnHasContent = true;
                    firstLine = false;
                    paragraphX = x + continuationTextStartOffset;
                    paragraphWidth = Math.Max(1d, width - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                    cursorY -= lineHeight;
                }
            }
            else if (paragraph.Images.Count == 0)
            {
                if (cursorY - lineHeight < page.MarginBottom && HasCurrentColumnContent())
                {
                    AdvanceColumnOrPage();
                }

                cursorY -= lineHeight;
                activeColumnHasContent = true;
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(width, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                if (cursorY - imageHeight < page.MarginBottom && HasCurrentColumnContent())
                {
                    AdvanceColumnOrPage();
                }

                double imageX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => x + Math.Max(0, width - imageWidth) / 2d,
                    DocxTextAlignment.Right => x + Math.Max(0, width - imageWidth),
                    _ => x
                };
                currentItems.Add(new DocxInlineImageLayout(
                    image,
                    imageX,
                    cursorY - imageHeight,
                    imageWidth,
                    imageHeight,
                    pages.Count + 1,
                    SourceBlockIndex: elementIndex,
                    SourceParagraphIndex: 0));
                activeColumnHasContent = true;
                cursorY -= imageHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        if (HasPageContent() || pages.Count == 0)
        {
            FinishPage();
        }

        DocxLayoutPage[] pagesWithStaticText = AddStaticContent(pages, textMeasurer, defaultTabStopPoints).ToArray();
        return new DocxLayout(
            pagesWithStaticText,
            CreateFloatingDrawingLayouts(document.FloatingDrawings, pagesWithStaticText),
            CreateStaticFloatingDrawingLayouts(pagesWithStaticText),
            CreateRelatedStoryLayouts(document.RelatedStories, page.BodyWidth, textMeasurer, defaultTabStopPoints));
    }

    private static IReadOnlyList<DocxRelatedStoryLayout> CreateRelatedStoryLayouts(
        IReadOnlyList<DocxRelatedStory> stories,
        double bodyWidth,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints)
    {
        if (textMeasurer is null || stories.Count == 0)
        {
            return stories
                .Select((story, index) => new DocxRelatedStoryLayout(story, index, [], [], [], 0d))
                .ToArray();
        }

        return stories
            .Select((story, index) => CreateRelatedStoryLayout(story, index, bodyWidth, textMeasurer, defaultTabStopPoints))
            .ToArray();
    }

    private static DocxRelatedStoryLayout CreateRelatedStoryLayout(
        DocxRelatedStory story,
        int storyIndex,
        double bodyWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints)
    {
        var textLines = new List<DocxTextLineLayout>();
        var inlineImages = new List<DocxInlineImageLayout>();
        var tableRows = new List<DocxTableRowLayout>();
        double cursorY = 0d;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int paragraphIndex = 0;
        int tableIndex = 0;

        for (int elementIndex = 0; elementIndex < story.BodyElements.Count; elementIndex++)
        {
            DocxBodyElement element = story.BodyElements[elementIndex];
            if (element is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                DocxTableLayoutFrame frame = CreateTableLayoutFrame(
                    tableElement.Table,
                    tableIndex++,
                    elementIndex,
                    0d,
                    bodyWidth,
                    UnpagedRelatedStoryCanvasHeightPoints,
                    textMeasurer,
                    defaultTabStopPoints);
                for (int rowIndex = 0; rowIndex < tableElement.Table.Rows.Count; rowIndex++)
                {
                    double rowHeight = frame.RowHeights[rowIndex];
                    tableRows.Add(CreateTableRowLayout(
                        tableElement.Table,
                        frame.Context,
                        tableElement.Table.Rows[rowIndex],
                        rowIndex,
                        frame.RowHeights,
                        frame.EffectiveColumns,
                        frame.Scale,
                        textMeasurer,
                        defaultTabStopPoints,
                        () => 0,
                        cursorY,
                        rowHeight,
                        cursorY,
                        FragmentIndex: 0,
                        FragmentCount: 1,
                        FragmentReason: "None"));
                    cursorY -= rowHeight;
                }

                continue;
            }

            if (element is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            IReadOnlyList<DocxTextLineLayout> paragraphLines = LayoutRelatedStoryParagraphTextLines(
                paragraph,
                elementIndex,
                paragraphIndex,
                story.Kind,
                bodyWidth,
                cursorY,
                spacingProfile,
                textMeasurer,
                defaultTabStopPoints);
            textLines.AddRange(paragraphLines);
            cursorY -= paragraphLines.Sum(line => line.LineHeight ?? 0d);
            if (paragraphLines.Count == 0 && paragraph.Images.Count == 0)
            {
                double fontSize = GetParagraphFontSize(paragraph);
                cursorY -= ResolveLineHeight(paragraph, fontSize, textMeasurer);
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(bodyWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                double imageX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => Math.Max(0, bodyWidth - imageWidth) / 2d,
                    DocxTextAlignment.Right => Math.Max(0, bodyWidth - imageWidth),
                    _ => 0d
                };
                inlineImages.Add(new DocxInlineImageLayout(
                    image,
                    imageX,
                    cursorY - imageHeight,
                    imageWidth,
                    imageHeight,
                    PageIndex: 0,
                    SourceBlockIndex: elementIndex,
                    SourceParagraphIndex: paragraphIndex));
                cursorY -= imageHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        cursorY -= pendingSpacingAfter;
        return new DocxRelatedStoryLayout(story, storyIndex, textLines.ToArray(), inlineImages.ToArray(), tableRows.ToArray(), Math.Abs(cursorY));
    }

    private static IReadOnlyList<DocxTextLineLayout> LayoutRelatedStoryParagraphTextLines(
        DocxParagraph paragraph,
        int sourceBlockIndex,
        int sourceParagraphIndex,
        string storyKind,
        double bodyWidth,
        double cursorY,
        DocxParagraphSpacingProfile spacingProfile,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints)
    {
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
        if (textSpans.Count == 0)
        {
            return [];
        }

        double fontSize = GetParagraphFontSize(paragraph);
        DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer);
        double lineHeight = lineHeightProfile.LineHeight;
        double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
        double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
        double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
        double paragraphX = textStartOffset;
        double paragraphWidth = Math.Max(1d, bodyWidth - textStartOffset - GetParagraphRightInset(paragraph));
        double continuationParagraphWidth = Math.Max(1d, bodyWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
        DocxTextRun firstRun = paragraph.Runs[0];
        DocxWrappedTextLine[] lines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: false).ToArray();
        var layouts = new List<DocxTextLineLayout>(lines.Length);
        bool firstLine = true;
        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            DocxWrappedTextLine line = lines[lineIndex];
            double lineWidth = MeasureTextSpans(line.Spans, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
            double drawableLineWidth = MeasureDrawableTextSpans(line.Spans, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
            double lineX = paragraph.Alignment switch
            {
                DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                _ => paragraphX
            };
            double baselineOffset = DocxLineMetrics.ResolveBodyBaselineOffset(fontSize, lineHeight, paragraph.LineSpacingPoints is not null);
            bool justifyLine = (paragraph.ListLabel is null || !firstLine) &&
                ShouldJustifyTextLine(paragraph.Alignment, lineIndex == lines.Length - 1, drawableLineWidth, paragraphWidth, line.Spans);
            IReadOnlyList<DocxTextSegmentLayout> segments = firstLine && paragraph.ListLabel is not null
                ? CreateNumberedLineSegments(paragraph.ListLabel, line.Spans, firstRun, labelStartOffset, lineX, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                : justifyLine
                    ? CreateJustifiedTextSegments(line.Spans, lineX, drawableLineWidth, paragraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                    : CreateTextSegments(line.Spans, lineX, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
            double effectiveX = firstLine && paragraph.ListLabel is not null ? labelStartOffset : lineX;
            double effectiveWidth = firstLine && paragraph.ListLabel is not null
                ? Math.Max(lineX + lineWidth, labelStartOffset + MeasureListLabel(paragraph.ListLabel, firstRun, fontSize, textMeasurer)) - labelStartOffset
                : justifyLine
                    ? paragraphWidth
                    : lineWidth;
            layouts.Add(new DocxTextLineLayout(
                firstLine && paragraph.ListLabel is not null ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line.Text : line.Text,
                firstRun,
                fontSize,
                effectiveX,
                cursorY - baselineOffset,
                effectiveWidth,
                segments,
                SourceBlockIndex: sourceBlockIndex,
                SourceParagraphIndex: sourceParagraphIndex,
                SourceLineIndex: lineIndex,
                StoryKind: storyKind,
                LineHeight: lineHeight,
                AppliedBeforeSpacing: firstLine ? spacingProfile.AppliedBeforeSpacing : 0d,
                IsFirstParagraphLine: firstLine,
                EndsWithIntraTokenBreak: line.EndsWithIntraTokenBreak,
                SingleLineHeight: lineHeightProfile.SingleLineHeight,
                ListLabelSingleLineHeight: lineHeightProfile.ListLabelSingleLineHeight,
                BodyWindowsLineHeight: lineHeightProfile.BodyWindowsLineHeight,
                ListLabelWindowsLineHeight: lineHeightProfile.ListLabelWindowsLineHeight,
                EffectiveLineSpacingFactor: lineHeightProfile.EffectiveLineSpacingFactor,
                LineSpacingFactorFloorApplied: lineHeightProfile.LineSpacingFactorFloorApplied,
                PendingAfterSpacing: firstLine ? spacingProfile.PendingAfterSpacing : null,
                ParagraphBeforeSpacing: firstLine ? spacingProfile.ParagraphBeforeSpacing : null,
                ParagraphAfterSpacing: firstLine ? spacingProfile.ParagraphAfterSpacing : null,
                ContextualSpacingSuppressed: firstLine ? spacingProfile.ContextualSpacingSuppressed : null,
                SourceParagraph: paragraph));
            firstLine = false;
            paragraphX = continuationTextStartOffset;
            paragraphWidth = Math.Max(1d, bodyWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
            cursorY -= lineHeight;
        }

        return layouts;
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateFloatingDrawingLayouts(
        IReadOnlyList<DocxFloatingDrawing> drawings,
        IReadOnlyList<DocxLayoutPage> pages)
    {
        return drawings
            .Select(drawing =>
            {
                DocxLayoutSourceBlockBounds? sourceBlock = drawing.SourceBlockIndex is null
                    ? null
                    : FindSourceBlockBounds(pages, drawing.SourceBlockIndex.Value);
                DocxLayoutPage? anchorPage = sourceBlock is null
                    ? pages.FirstOrDefault()
                    : pages[sourceBlock.FirstPageIndex];
                return CreateFloatingDrawingLayout(
                    drawing,
                    anchorPage,
                    sourceBlock?.FirstPageIndex,
                    sourceBlock?.LastPageIndex,
                    sourceBlock?.FirstPageIndex,
                    sourceBlock?.FirstColumnIndex,
                    sourceBlock);
            })
            .ToArray();
    }

    private static IReadOnlyList<DocxFloatingDrawingLayout> CreateStaticFloatingDrawingLayouts(IReadOnlyList<DocxLayoutPage> pages)
    {
        var layouts = new List<DocxFloatingDrawingLayout>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            DocxLayoutPage page = pages[pageIndex];
            int pageNumber = pageIndex + 1;
            DocxSelectedStaticDrawings selectedHeader = SelectStaticHeaderFooterDrawings(
                page.PageSettings.HeaderFloatingDrawingsByType,
                page.PageSettings,
                pageNumber);
            DocxSelectedStaticDrawings selectedFooter = SelectStaticHeaderFooterDrawings(
                page.PageSettings.FooterFloatingDrawingsByType,
                page.PageSettings,
                pageNumber);
            layouts.AddRange(CreateStaticFloatingDrawingLayouts(selectedHeader, "Header", page, pageIndex));
            layouts.AddRange(CreateStaticFloatingDrawingLayouts(selectedFooter, "Footer", page, pageIndex));
        }

        return layouts;
    }

    private static IEnumerable<DocxFloatingDrawingLayout> CreateStaticFloatingDrawingLayouts(
        DocxSelectedStaticDrawings selectedDrawings,
        string storyKind,
        DocxLayoutPage page,
        int pageIndex)
    {
        foreach (DocxFloatingDrawing drawing in selectedDrawings.Drawings)
        {
            yield return CreateFloatingDrawingLayout(
                drawing,
                page,
                pageStartIndex: pageIndex,
                pageEndIndex: pageIndex,
                anchorPageIndex: pageIndex,
                anchorColumnIndex: null,
                sourceBlock: null,
                storyKind,
                selectedDrawings.VariantType);
        }
    }

    private static DocxFloatingDrawingLayout CreateFloatingDrawingLayout(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? anchorPage,
        int? pageStartIndex,
        int? pageEndIndex,
        int? anchorPageIndex,
        int? anchorColumnIndex,
        DocxLayoutSourceBlockBounds? sourceBlock,
        string? storyKind = null,
        string? storyVariantType = null)
    {
        DocxAnchorReferenceFrame? horizontalReference = ResolveHorizontalReferenceFrame(drawing, anchorPage, sourceBlock);
        DocxAnchorReferenceFrame? verticalReference = ResolveVerticalReferenceFrame(drawing, anchorPage, sourceBlock);
        double? extentWidth = ReadEmuPoints(drawing.ExtentCxValue);
        double? extentHeight = ReadEmuPoints(drawing.ExtentCyValue);
        double? horizontalOffset = ReadEmuPoints(drawing.HorizontalOffsetValue);
        double? verticalOffset = ReadEmuPoints(drawing.VerticalOffsetValue);
        double? distanceTop = ReadEmuPoints(drawing.DistanceTopValue);
        double? distanceBottom = ReadEmuPoints(drawing.DistanceBottomValue);
        double? distanceLeft = ReadEmuPoints(drawing.DistanceLeftValue);
        double? distanceRight = ReadEmuPoints(drawing.DistanceRightValue);
        DocxAnchorPlacement horizontalPlacement = ResolveHorizontalPlacement(horizontalReference, extentWidth, drawing.HorizontalAlignValue, horizontalOffset);
        DocxAnchorPlacement verticalPlacement = ResolveVerticalPlacement(verticalReference, extentHeight, drawing.VerticalAlignValue, verticalOffset);
        double? placedX = horizontalPlacement.Position;
        double? placedTop = verticalPlacement.Position;
        DocxWrapExclusionFrame? wrapExclusion = CreateWrapExclusionFrame(drawing, placedX, placedTop, extentWidth, extentHeight, distanceTop, distanceBottom, distanceLeft, distanceRight);
        return new DocxFloatingDrawingLayout(
            drawing,
            pageStartIndex,
            pageEndIndex,
            anchorPageIndex,
            anchorColumnIndex,
            sourceBlock?.VerticalTop,
            sourceBlock?.VerticalBottom,
            extentWidth,
            extentHeight,
            horizontalOffset,
            verticalOffset,
            distanceTop,
            distanceBottom,
            distanceLeft,
            distanceRight,
            horizontalReference?.Start,
            horizontalReference?.Size,
            verticalReference?.Start,
            verticalReference?.End,
            placedX,
            placedTop,
            horizontalPlacement.Source,
            verticalPlacement.Source,
            wrapExclusion?.X,
            wrapExclusion?.Top,
            wrapExclusion?.Width,
            wrapExclusion?.Height,
            storyKind,
            storyVariantType);
    }

    private static DocxWrapExclusionFrame? CreateWrapExclusionFrame(
        DocxFloatingDrawing drawing,
        double? placedX,
        double? placedTop,
        double? extentWidth,
        double? extentHeight,
        double? distanceTop,
        double? distanceBottom,
        double? distanceLeft,
        double? distanceRight)
    {
        if (!IsWrapExclusionKind(drawing.WrapKind) ||
            placedX is not { } x ||
            placedTop is not { } top ||
            extentWidth is not { } width ||
            extentHeight is not { } height)
        {
            return null;
        }

        double leftDistance = distanceLeft ?? 0d;
        double rightDistance = distanceRight ?? 0d;
        double topDistance = distanceTop ?? 0d;
        double bottomDistance = distanceBottom ?? 0d;
        return new DocxWrapExclusionFrame(
            x - leftDistance,
            top + topDistance,
            width + leftDistance + rightDistance,
            height + topDistance + bottomDistance);
    }

    private static bool IsWrapExclusionKind(string? wrapKind)
    {
        return wrapKind is not null &&
            (wrapKind.Equals("wrapSquare", StringComparison.OrdinalIgnoreCase) ||
            wrapKind.Equals("wrapTight", StringComparison.OrdinalIgnoreCase) ||
            wrapKind.Equals("wrapThrough", StringComparison.OrdinalIgnoreCase) ||
            wrapKind.Equals("wrapTopAndBottom", StringComparison.OrdinalIgnoreCase));
    }

    private static DocxAnchorPlacement ResolveHorizontalPlacement(
        DocxAnchorReferenceFrame? reference,
        double? extentWidth,
        string? alignValue,
        double? offset)
    {
        if (reference is null || extentWidth is null)
        {
            return new DocxAnchorPlacement(null, DocxAnchorPlacementSource.MissingReferenceOrExtent);
        }

        return alignValue?.ToLowerInvariant() switch
        {
            "left" => new DocxAnchorPlacement(reference.Start, DocxAnchorPlacementSource.Align),
            "center" => new DocxAnchorPlacement(reference.Start + Math.Max(0d, reference.Size - extentWidth.Value) / 2d, DocxAnchorPlacementSource.Align),
            "right" => new DocxAnchorPlacement(reference.End - extentWidth.Value, DocxAnchorPlacementSource.Align),
            null when offset is not null => new DocxAnchorPlacement(reference.Start + offset.Value, DocxAnchorPlacementSource.Offset),
            "" when offset is not null => new DocxAnchorPlacement(reference.Start + offset.Value, DocxAnchorPlacementSource.Offset),
            _ => new DocxAnchorPlacement(null, DocxAnchorPlacementSource.Unsupported)
        };
    }

    private static DocxAnchorPlacement ResolveVerticalPlacement(
        DocxAnchorReferenceFrame? reference,
        double? extentHeight,
        string? alignValue,
        double? offset)
    {
        if (reference is null || extentHeight is null)
        {
            return new DocxAnchorPlacement(null, DocxAnchorPlacementSource.MissingReferenceOrExtent);
        }

        return alignValue?.ToLowerInvariant() switch
        {
            "top" => new DocxAnchorPlacement(reference.Start, DocxAnchorPlacementSource.Align),
            "center" => new DocxAnchorPlacement(reference.End + (reference.Size + extentHeight.Value) / 2d, DocxAnchorPlacementSource.Align),
            "bottom" => new DocxAnchorPlacement(reference.End + extentHeight.Value, DocxAnchorPlacementSource.Align),
            null when offset is not null => new DocxAnchorPlacement(reference.Start - offset.Value, DocxAnchorPlacementSource.Offset),
            "" when offset is not null => new DocxAnchorPlacement(reference.Start - offset.Value, DocxAnchorPlacementSource.Offset),
            _ => new DocxAnchorPlacement(null, DocxAnchorPlacementSource.Unsupported)
        };
    }

    private static DocxAnchorReferenceFrame? ResolveHorizontalReferenceFrame(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? page,
        DocxLayoutSourceBlockBounds? sourceBlock)
    {
        if (page is null)
        {
            return null;
        }

        return drawing.HorizontalRelativeFromValue?.ToLowerInvariant() switch
        {
            "page" => new DocxAnchorReferenceFrame(0d, page.Width),
            "margin" => new DocxAnchorReferenceFrame(page.MarginLeft, page.Width - page.MarginRight),
            "column" when page.ColumnFrames.Count == 1 => new DocxAnchorReferenceFrame(page.ColumnFrames[0].X, page.ColumnFrames[0].X + page.ColumnFrames[0].Width),
            "column" when sourceBlock?.FirstColumnIndex is { } columnIndex && columnIndex >= 0 && columnIndex < page.ColumnFrames.Count =>
                new DocxAnchorReferenceFrame(page.ColumnFrames[columnIndex].X, page.ColumnFrames[columnIndex].X + page.ColumnFrames[columnIndex].Width),
            _ => null
        };
    }

    private static DocxAnchorReferenceFrame? ResolveVerticalReferenceFrame(
        DocxFloatingDrawing drawing,
        DocxLayoutPage? page,
        DocxLayoutSourceBlockBounds? sourceBlock)
    {
        if (page is null)
        {
            return null;
        }

        return drawing.VerticalRelativeFromValue?.ToLowerInvariant() switch
        {
            "page" => new DocxAnchorReferenceFrame(page.Height, 0d),
            "margin" => new DocxAnchorReferenceFrame(page.Height - page.MarginTop, page.MarginBottom),
            "paragraph" when sourceBlock is not null => new DocxAnchorReferenceFrame(sourceBlock.VerticalTop, sourceBlock.VerticalBottom),
            _ => null
        };
    }

    private static double? ReadEmuPoints(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu)
            ? OoxUnits.EmuToPoints(emu)
            : null;
    }

    private static DocxLayoutSourceBlockBounds? FindSourceBlockBounds(IReadOnlyList<DocxLayoutPage> pages, int sourceBlockIndex)
    {
        int? firstPageIndex = null;
        int? lastPageIndex = null;
        int? firstColumnIndex = null;
        int? lastColumnIndex = null;
        double verticalTop = double.NegativeInfinity;
        double verticalBottom = double.PositiveInfinity;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            DocxLayoutPage page = pages[pageIndex];
            foreach (DocxLayoutItem item in pages[pageIndex].Items)
            {
                if (GetSourceBlockIndex(item) != sourceBlockIndex)
                {
                    continue;
                }

                firstPageIndex ??= pageIndex;
                lastPageIndex = pageIndex;
                if (ResolveItemColumnIndex(page, item) is { } columnIndex)
                {
                    firstColumnIndex = firstColumnIndex is null ? columnIndex : Math.Min(firstColumnIndex.Value, columnIndex);
                    lastColumnIndex = lastColumnIndex is null ? columnIndex : Math.Max(lastColumnIndex.Value, columnIndex);
                }

                (double y, double height) = GetVerticalBounds(item);
                verticalTop = Math.Max(verticalTop, y + height);
                verticalBottom = Math.Min(verticalBottom, y);
            }
        }

        return firstPageIndex is null || lastPageIndex is null
            ? null
            : new DocxLayoutSourceBlockBounds(firstPageIndex.Value, lastPageIndex.Value, firstColumnIndex, lastColumnIndex, verticalTop, verticalBottom);
    }

    private static int? GetSourceBlockIndex(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => text.SourceBlockIndex,
            DocxInlineImageLayout image => image.SourceBlockIndex,
            DocxTableRowLayout row => row.Table.SourceBlockIndex,
            _ => null
        };
    }

    private static (double Y, double Height) GetVerticalBounds(DocxLayoutItem item)
    {
        return item switch
        {
            DocxTextLineLayout text => (text.BaselineY, text.FontSize),
            DocxInlineImageLayout image => (image.Y, image.Height),
            DocxTableRowLayout row => (row.Y, row.Height),
            _ => (0d, 0d)
        };
    }

    private static int? ResolveItemColumnIndex(DocxLayoutPage page, DocxLayoutItem item)
    {
        return ResolveItemColumnIndex(page.ColumnFrames, item);
    }

    private static int? ResolveItemColumnIndex(IReadOnlyList<DocxLayoutColumnFrame> frames, DocxLayoutItem item)
    {
        (double x, double width) = item switch
        {
            DocxTextLineLayout text => (text.X, text.Width),
            DocxInlineImageLayout image => (image.X, image.Width),
            DocxTableRowLayout row => (row.Table.TableX, row.Table.ResolvedTableWidth),
            _ => (0d, 0d)
        };
        return DocxLayoutColumnOwnership.ResolveColumnIndex(frames, x, width);
    }

    private sealed record DocxLayoutSourceBlockBounds(
        int FirstPageIndex,
        int LastPageIndex,
        int? FirstColumnIndex,
        int? LastColumnIndex,
        double VerticalTop,
        double VerticalBottom);

    private sealed record DocxAnchorReferenceFrame(double Start, double End)
    {
        public double Size => Math.Abs(Start - End);
    }

    private sealed record DocxStaticStoryLayoutResult(
        IReadOnlyList<DocxTextLineLayout> TextLines,
        IReadOnlyList<DocxInlineImageLayout> InlineImages,
        IReadOnlyList<DocxTableRowLayout> TableRows);

    private static IReadOnlyList<DocxLayoutPage> AddStaticContent(
        IReadOnlyList<DocxLayoutPage> pages,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints)
    {
        if (textMeasurer is not IDocxStaticTextMetricsProvider staticMetrics)
        {
            return pages;
        }

        var pagesWithStaticText = new DocxLayoutPage[pages.Count];
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++)
        {
            DocxLayoutPage page = pages[pageIndex];
            int pageNumber = pageIndex + 1;
            double bodyWidth = Math.Max(1d, page.Width - page.MarginLeft - page.MarginRight);
            DocxSelectedStaticStory selectedHeader = SelectStaticHeaderFooter(
                page.PageSettings.HeaderBodyElementsByType,
                page.PageSettings.HeaderParagraphsByType,
                page.PageSettings,
                pageNumber);
            DocxSelectedStaticStory selectedFooter = SelectStaticHeaderFooter(
                page.PageSettings.FooterBodyElementsByType,
                page.PageSettings.FooterParagraphsByType,
                page.PageSettings,
                pageNumber);
            DocxStaticStoryLayoutResult headerLayout = CreateStaticStoryLayout(
                    selectedHeader,
                    page.MarginLeft,
                    bodyWidth,
                    page.Height - ResolveHeaderDistance(page),
                    true,
                    pageNumber,
                    pages.Count,
                    textMeasurer,
                    staticMetrics,
                    defaultTabStopPoints);
            DocxStaticStoryLayoutResult footerLayout = CreateStaticStoryLayout(
                    selectedFooter,
                    page.MarginLeft,
                    bodyWidth,
                    ResolveFooterDistance(page),
                    false,
                    pageNumber,
                    pages.Count,
                    textMeasurer,
                    staticMetrics,
                    defaultTabStopPoints);
            pagesWithStaticText[pageIndex] = page with
            {
                StaticTextLines = headerLayout.TextLines.Concat(footerLayout.TextLines).ToArray(),
                StaticInlineImages = headerLayout.InlineImages.Concat(footerLayout.InlineImages).ToArray(),
                StaticTableRows = headerLayout.TableRows.Concat(footerLayout.TableRows).ToArray()
            };
        }

        return pagesWithStaticText;
    }

    private static DocxStaticStoryLayoutResult CreateStaticStoryLayout(
        DocxSelectedStaticStory story,
        double x,
        double width,
        double startY,
        bool isHeader,
        int pageNumber,
        int pageCount,
        IDocxTextMeasurer textMeasurer,
        IDocxStaticTextMetricsProvider staticMetrics,
        double defaultTabStopPoints)
    {
        var lines = new List<DocxTextLineLayout>();
        var images = new List<DocxInlineImageLayout>();
        var tableRows = new List<DocxTableRowLayout>();
        double cursorY = startY;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int paragraphIndex = 0;
        int tableIndex = 0;
        for (int elementIndex = 0; elementIndex < story.BodyElements.Count; elementIndex++)
        {
            if (story.BodyElements[elementIndex] is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                DocxTableLayoutFrame frame = CreateTableLayoutFrame(
                    tableElement.Table,
                    tableIndex++,
                    elementIndex,
                    x,
                    width,
                    UnpagedRelatedStoryCanvasHeightPoints,
                    textMeasurer,
                    defaultTabStopPoints);
                for (int rowIndex = 0; rowIndex < tableElement.Table.Rows.Count; rowIndex++)
                {
                    double rowHeight = frame.RowHeights[rowIndex];
                    tableRows.Add(CreateTableRowLayout(
                        tableElement.Table,
                        frame.Context,
                        tableElement.Table.Rows[rowIndex],
                        rowIndex,
                        frame.RowHeights,
                        frame.EffectiveColumns,
                        frame.Scale,
                        textMeasurer,
                        defaultTabStopPoints,
                        () => pageNumber - 1,
                        cursorY,
                        rowHeight,
                        cursorY,
                        FragmentIndex: 0,
                        FragmentCount: 1,
                        FragmentReason: "None",
                        StoryKind: isHeader ? "Header" : "Footer",
                        StoryVariantType: story.VariantType));
                    cursorY -= rowHeight;
                }

                continue;
            }

            if (story.BodyElements[elementIndex] is not DocxParagraphElement paragraphElement)
            {
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            DocxTextSpan[] spans = CreateStaticTextSpans(paragraph.Runs, pageNumber, pageCount);
            int sourceLineIndex = 0;
            if (spans.Length != 0)
            {
                foreach (DocxWrappedTextLine line in WrapStaticTextLines(spans, width, textMeasurer))
                {
                    if (line.Spans.Count == 0)
                    {
                        continue;
                    }

                    double lineWidth = MeasureStaticTextSpans(line.Spans, textMeasurer);
                    double lineX = paragraph.Alignment switch
                    {
                        DocxTextAlignment.Center => x + Math.Max(0d, width - lineWidth) / 2d,
                        DocxTextAlignment.Right => x + Math.Max(0d, width - lineWidth),
                        _ => x
                    };
                    double ascender = line.Spans.Max(span => staticMetrics.MeasureWindowsAscender(span.StyleRun, span.StyleRun.FontSize));
                    double descender = line.Spans.Max(span => staticMetrics.MeasureWindowsDescender(span.StyleRun, span.StyleRun.FontSize));
                    double baselineY = isHeader ? cursorY - ascender : cursorY + descender;
                    IReadOnlyList<DocxTextSegmentLayout> segments = CreateStaticTextSegments(line.Spans, lineX, textMeasurer);
                    lines.Add(new DocxTextLineLayout(
                        line.Text,
                        line.Spans[0].StyleRun,
                        line.Spans.Max(span => span.StyleRun.FontSize),
                        lineX,
                        baselineY,
                        lineWidth,
                        segments,
                        LineHeight: ascender + descender,
                        AppliedBeforeSpacing: sourceLineIndex == 0 ? spacingProfile.AppliedBeforeSpacing : 0d,
                        IsFirstParagraphLine: sourceLineIndex == 0,
                        SourceLineIndex: sourceLineIndex,
                        PendingAfterSpacing: sourceLineIndex == 0 ? spacingProfile.PendingAfterSpacing : null,
                        ParagraphBeforeSpacing: sourceLineIndex == 0 ? spacingProfile.ParagraphBeforeSpacing : null,
                        ParagraphAfterSpacing: sourceLineIndex == 0 ? spacingProfile.ParagraphAfterSpacing : null,
                        ContextualSpacingSuppressed: sourceLineIndex == 0 ? spacingProfile.ContextualSpacingSuppressed : null,
                        SourceParagraph: paragraph,
                        SourceParagraphIndex: paragraphIndex,
                        StoryKind: isHeader ? "Header" : "Footer",
                        StoryVariantType: story.VariantType));
                    sourceLineIndex++;
                    cursorY -= ascender + descender;
                }
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(width, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                double imageX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => x + Math.Max(0d, width - imageWidth) / 2d,
                    DocxTextAlignment.Right => x + Math.Max(0d, width - imageWidth),
                    _ => x
                };
                images.Add(new DocxInlineImageLayout(
                    image,
                    imageX,
                    cursorY - imageHeight,
                    imageWidth,
                    imageHeight,
                    pageNumber,
                    SourceParagraphIndex: paragraphIndex,
                    StoryKind: isHeader ? "Header" : "Footer",
                    StoryVariantType: story.VariantType));
                cursorY -= imageHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        return new DocxStaticStoryLayoutResult(lines.ToArray(), images.ToArray(), tableRows.ToArray());
    }

    private static DocxTextSpan[] CreateStaticTextSpans(IReadOnlyList<DocxTextRun> runs, int pageNumber, int pageCount)
    {
        if (runs.Count != 0 && runs.All(run => run.Text.Length == 0 || run.Hidden))
        {
            for (int i = 0; i < runs.Count; i++)
            {
                if (!runs[i].Hidden)
                {
                    return [new DocxTextSpan(" ", runs[i], i)];
                }
            }

            return [];
        }

        return runs
            .Select((run, index) => (run, index))
            .Where(item => !item.run.Hidden)
            .Select(item => new DocxTextSpan(ResolveStaticFieldPlaceholders(item.run.Text, pageNumber, pageCount), item.run, item.index))
            .Where(span => span.Text.Length != 0)
            .ToArray();
    }

    private static IEnumerable<DocxWrappedTextLine> WrapStaticTextLines(
        IReadOnlyList<DocxTextSpan> spans,
        double maxWidth,
        IDocxTextMeasurer textMeasurer)
    {
        string text = string.Concat(spans.Select(span => span.Text));
        int segmentStart = 0;
        while (segmentStart <= text.Length)
        {
            int breakIndex = text.IndexOf('\n', segmentStart);
            int segmentLength = breakIndex < 0 ? text.Length - segmentStart : breakIndex - segmentStart;
            bool yielded = false;
            foreach (DocxWrappedTextLine line in WrapStaticWords(text, spans, segmentStart, segmentLength, maxWidth, textMeasurer))
            {
                yielded = true;
                yield return line;
            }

            if (!yielded && segmentLength == 0)
            {
                yield return new DocxWrappedTextLine(string.Empty, []);
            }

            if (breakIndex < 0)
            {
                yield break;
            }

            segmentStart = breakIndex + 1;
        }
    }

    private static IEnumerable<DocxWrappedTextLine> WrapStaticWords(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int segmentStart,
        int segmentLength,
        double maxWidth,
        IDocxTextMeasurer textMeasurer)
    {
        IReadOnlyList<TextToken> tokens = TokenizeSpaces(text, segmentStart, segmentLength);
        if (tokens.Count == 0)
        {
            yield break;
        }

        int lineStart = tokens[0].Start;
        int lineLength = 0;
        foreach (TextToken token in tokens)
        {
            int candidateLength = token.Start + token.Length - lineStart;
            bool lineHasNonWhitespace = HasNonWhitespace(text, lineStart, lineLength);
            if (lineLength > 0 &&
                lineHasNonWhitespace &&
                !token.IsBreakableWhitespace &&
                MeasureStaticTextSpans(SliceTextSpans(spans, lineStart, candidateLength), textMeasurer) > maxWidth)
            {
                yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
                lineStart = token.Start;
                lineLength = token.Length;
            }
            else
            {
                lineLength = candidateLength;
            }
        }

        if (lineLength > 0)
        {
            yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
        }
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateStaticTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        IDocxTextMeasurer textMeasurer)
    {
        var segments = new List<DocxTextSegmentLayout>(spans.Count);
        double segmentX = lineX;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double nominalFontSize = span.StyleRun.FontSize;
            double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
            double baselineOffset = DocxVerticalAlignMetrics.ResolveBaselineOffset(nominalFontSize, layoutFontSize, span.StyleRun);
            double width = textMeasurer.MeasureText(span.StyleRun, span.Text, layoutFontSize);
            segments.Add(new DocxTextSegmentLayout(span.Text, span.StyleRun, segmentX, width, layoutFontSize, baselineOffset, SourceTextRunIndex: span.SourceTextRunIndex));
            segmentX += width;
            if (i + 1 < spans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return segments;
    }

    private static double MeasureStaticTextSpans(IReadOnlyList<DocxTextSpan> spans, IDocxTextMeasurer textMeasurer)
    {
        double width = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(span.StyleRun.FontSize, span.StyleRun);
            width += textMeasurer.MeasureText(span.StyleRun, span.Text, layoutFontSize);
            if (i + 1 < spans.Count)
            {
                width += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return width;
    }

    private static DocxSelectedStaticStory SelectStaticHeaderFooter(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        DocxPageSettings settings,
        int pageNumber)
    {
        if (settings.TitlePage == true &&
            pageNumber == 1 &&
            TryGetStaticStoryBodyElements("first", bodyElementsByType, paragraphsByType, out IReadOnlyList<DocxBodyElement>? first))
        {
            return new DocxSelectedStaticStory(first, "first");
        }

        if (settings.EvenAndOddHeaders == true &&
            pageNumber % 2 == 0 &&
            TryGetStaticStoryBodyElements("even", bodyElementsByType, paragraphsByType, out IReadOnlyList<DocxBodyElement>? even))
        {
            return new DocxSelectedStaticStory(even, "even");
        }

        return TryGetStaticStoryBodyElements("default", bodyElementsByType, paragraphsByType, out IReadOnlyList<DocxBodyElement>? defaults)
            ? new DocxSelectedStaticStory(defaults, "default")
            : new DocxSelectedStaticStory([], null);
    }

    private static bool TryGetStaticStoryBodyElements(
        string variantType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        out IReadOnlyList<DocxBodyElement> bodyElements)
    {
        if (bodyElementsByType.TryGetValue(variantType, out bodyElements!))
        {
            return true;
        }

        if (paragraphsByType.TryGetValue(variantType, out IReadOnlyList<DocxParagraph>? paragraphs))
        {
            bodyElements = paragraphs.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>().ToArray();
            return true;
        }

        bodyElements = [];
        return false;
    }

    private sealed record DocxSelectedStaticStory(IReadOnlyList<DocxBodyElement> BodyElements, string? VariantType);

    private static DocxSelectedStaticDrawings SelectStaticHeaderFooterDrawings(
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> drawingsByType,
        DocxPageSettings settings,
        int pageNumber)
    {
        if (settings.TitlePage == true &&
            pageNumber == 1 &&
            drawingsByType.TryGetValue("first", out IReadOnlyList<DocxFloatingDrawing>? first))
        {
            return new DocxSelectedStaticDrawings(first, "first");
        }

        if (settings.EvenAndOddHeaders == true &&
            pageNumber % 2 == 0 &&
            drawingsByType.TryGetValue("even", out IReadOnlyList<DocxFloatingDrawing>? even))
        {
            return new DocxSelectedStaticDrawings(even, "even");
        }

        return drawingsByType.TryGetValue("default", out IReadOnlyList<DocxFloatingDrawing>? defaults)
            ? new DocxSelectedStaticDrawings(defaults, "default")
            : new DocxSelectedStaticDrawings([], null);
    }

    private sealed record DocxSelectedStaticDrawings(IReadOnlyList<DocxFloatingDrawing> Drawings, string? VariantType);

    private static double ResolveHeaderDistance(DocxLayoutPage page)
    {
        return page.PageSettings.HeaderDistancePoints ?? Math.Max(18d, page.MarginTop / 2d);
    }

    private static double ResolveFooterDistance(DocxLayoutPage page)
    {
        return page.PageSettings.FooterDistancePoints ?? Math.Max(18d, page.MarginBottom / 2d);
    }

    private static string ResolveStaticFieldPlaceholders(string text, int pageNumber, int pageCount)
    {
        return text
            .Replace("{NUMPAGES}", pageCount.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
            .Replace("{PAGE}", pageNumber.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal);
    }

    private static IReadOnlyDictionary<int, DocxEffectiveSectionSettings> BuildEffectiveSectionSettings(DocxDocument document, out DocxEffectiveSectionSettings finalSectionSettings)
    {
        var sectionSettingsByElementIndex = new Dictionary<int, DocxEffectiveSectionSettings>();
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedHeadersByType =
            new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedFootersByType =
            new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedHeaderBodyElementsByType =
            new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedFooterBodyElementsByType =
            new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);

        for (int elementIndex = 0; elementIndex < document.BodyElements.Count; elementIndex++)
        {
            if (document.BodyElements[elementIndex] is not DocxSectionBreakElement sectionBreak)
            {
                continue;
            }

            DocxPageSettings effectiveSettings = ResolveEffectiveSectionSettings(
                sectionBreak.PageSettings,
                inheritedHeadersByType,
                inheritedFootersByType,
                inheritedHeaderBodyElementsByType,
                inheritedFooterBodyElementsByType);
            sectionSettingsByElementIndex[elementIndex] = new DocxEffectiveSectionSettings(
                effectiveSettings,
                CreateSectionLayoutProperties(sectionBreak));
            inheritedHeadersByType = effectiveSettings.HeaderParagraphsByType;
            inheritedFootersByType = effectiveSettings.FooterParagraphsByType;
            inheritedHeaderBodyElementsByType = effectiveSettings.HeaderBodyElementsByType;
            inheritedFooterBodyElementsByType = effectiveSettings.FooterBodyElementsByType;
        }

        finalSectionSettings = new DocxEffectiveSectionSettings(
            ResolveEffectiveSectionSettings(
                BuildFinalSectionSettings(document),
                inheritedHeadersByType,
                inheritedFootersByType,
                inheritedHeaderBodyElementsByType,
                inheritedFooterBodyElementsByType),
            document.FinalSectionBreak is null
                ? new DocxSectionLayoutProperties(null, null, null, null, null, null, [])
                : CreateSectionLayoutProperties(document.FinalSectionBreak));
        return sectionSettingsByElementIndex;
    }

    private static DocxSectionLayoutProperties CreateSectionLayoutProperties(DocxSectionBreakElement sectionBreak)
    {
        return new DocxSectionLayoutProperties(
            sectionBreak.TypeValue,
            sectionBreak.ColumnCountValue,
            sectionBreak.ColumnEqualWidthValue,
            sectionBreak.ColumnSpaceValue,
            ReadOptionalInt32Value(sectionBreak.ColumnCountValue),
            ReadOptionalTwipsValue(sectionBreak.ColumnSpaceValue),
            sectionBreak.ColumnDefinitions
                .Select(column => new DocxSectionColumnLayoutProperties(
                    column.WidthValue,
                    column.SpaceValue,
                    ReadOptionalTwipsValue(column.WidthValue),
                    ReadOptionalTwipsValue(column.SpaceValue)))
                .ToArray());
    }

    private static DocxPageSettings BuildFinalSectionSettings(DocxDocument document)
    {
        DocxPageSettings settings = document.PageSettings;
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = settings.HeaderParagraphsByType.Count == 0 && document.HeaderParagraphsByType.Count > 0
            ? document.HeaderParagraphsByType
            : settings.HeaderParagraphsByType;
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = settings.FooterParagraphsByType.Count == 0 && document.FooterParagraphsByType.Count > 0
            ? document.FooterParagraphsByType
            : settings.FooterParagraphsByType;
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> headerBodyElementsByType = settings.HeaderBodyElementsByType.Count == 0 && document.HeaderBodyElementsByType.Count > 0
            ? document.HeaderBodyElementsByType
            : settings.HeaderBodyElementsByType;
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> footerBodyElementsByType = settings.FooterBodyElementsByType.Count == 0 && document.FooterBodyElementsByType.Count > 0
            ? document.FooterBodyElementsByType
            : settings.FooterBodyElementsByType;

        if (headersByType.Count == 0 && document.HeaderParagraphs.Count > 0)
        {
            headersByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = document.HeaderParagraphs
            };
            headerBodyElementsByType = ToStaticBodyElementsByType(headersByType);
        }

        if (footersByType.Count == 0 && document.FooterParagraphs.Count > 0)
        {
            footersByType = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = document.FooterParagraphs
            };
            footerBodyElementsByType = ToStaticBodyElementsByType(footersByType);
        }

        return settings with
        {
            HeaderParagraphsByType = headersByType,
            FooterParagraphsByType = footersByType,
            HeaderBodyElementsByType = headerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(headersByType) : headerBodyElementsByType,
            FooterBodyElementsByType = footerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(footersByType) : footerBodyElementsByType
        };
    }

    private static DocxPageSettings ResolveEffectiveSectionSettings(
        DocxPageSettings settings,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedHeadersByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedFootersByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedHeaderBodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedFooterBodyElementsByType)
    {
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = MergeInheritedStaticParagraphs(inheritedHeadersByType, settings.HeaderParagraphsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = MergeInheritedStaticParagraphs(inheritedFootersByType, settings.FooterParagraphsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> headerBodyElementsByType = MergeInheritedStaticBodyElements(inheritedHeaderBodyElementsByType, settings.HeaderBodyElementsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> footerBodyElementsByType = MergeInheritedStaticBodyElements(inheritedFooterBodyElementsByType, settings.FooterBodyElementsByType);

        return settings with
        {
            HeaderParagraphsByType = headersByType,
            FooterParagraphsByType = footersByType,
            HeaderBodyElementsByType = headerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(headersByType) : headerBodyElementsByType,
            FooterBodyElementsByType = footerBodyElementsByType.Count == 0 ? ToStaticBodyElementsByType(footersByType) : footerBodyElementsByType
        };
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> ToStaticBodyElementsByType(
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType)
    {
        return paragraphsByType.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyList<DocxBodyElement>)pair.Value.Select(paragraph => new DocxParagraphElement(paragraph)).Cast<DocxBodyElement>().ToArray(),
            StringComparer.OrdinalIgnoreCase);
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> MergeInheritedStaticBodyElements(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> inheritedBodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> localBodyElementsByType)
    {
        if (inheritedBodyElementsByType.Count == 0)
        {
            return localBodyElementsByType;
        }

        if (localBodyElementsByType.Count == 0)
        {
            return inheritedBodyElementsByType;
        }

        var merged = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(inheritedBodyElementsByType, StringComparer.OrdinalIgnoreCase);
        foreach ((string type, IReadOnlyList<DocxBodyElement> bodyElements) in localBodyElementsByType)
        {
            merged[type] = bodyElements;
        }

        return merged;
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> MergeInheritedStaticParagraphs(
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> inheritedParagraphsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> localParagraphsByType)
    {
        if (inheritedParagraphsByType.Count == 0)
        {
            return localParagraphsByType;
        }

        if (localParagraphsByType.Count == 0)
        {
            return inheritedParagraphsByType;
        }

        var merged = new Dictionary<string, IReadOnlyList<DocxParagraph>>(inheritedParagraphsByType, StringComparer.OrdinalIgnoreCase);
        foreach ((string type, IReadOnlyList<DocxParagraph> paragraphs) in localParagraphsByType)
        {
            merged[type] = paragraphs;
        }

        return merged;
    }

    private static DocxEffectiveSectionSettings? FindSectionSettingsAtOrAfter(
        IReadOnlyList<DocxBodyElement> elements,
        int startIndex,
        IReadOnlyDictionary<int, DocxEffectiveSectionSettings> sectionSettingsByElementIndex)
    {
        for (int i = Math.Max(0, startIndex); i < elements.Count; i++)
        {
            if (elements[i] is DocxSectionBreakElement && sectionSettingsByElementIndex.TryGetValue(i, out DocxEffectiveSectionSettings? settings))
            {
                return settings;
            }
        }

        return null;
    }

    private static DocxPageGeometry ResolveSectionGeometry(DocxDocument document, DocxEffectiveSectionSettings section)
    {
        DocxPageSettings effectiveSettings = section.PageSettings;
        double width = ReadTwipsValue(effectiveSettings.WidthValue, document.PageWidthPoints);
        double height = ReadTwipsValue(effectiveSettings.HeightValue, document.PageHeightPoints);
        (width, height) = NormalizePageSize(width, height);
        if (effectiveSettings.OrientationValue?.Equals("landscape", StringComparison.OrdinalIgnoreCase) == true && height > width)
        {
            (width, height) = (height, width);
        }

        double marginLeft = ReadTwipsValue(effectiveSettings.MarginLeftValue, document.MarginLeftPoints);
        double marginRight = ReadTwipsValue(effectiveSettings.MarginRightValue, document.MarginRightPoints);
        double marginTop = ReadTwipsValue(effectiveSettings.MarginTopValue, document.MarginTopPoints);
        double marginBottom = ReadTwipsValue(effectiveSettings.MarginBottomValue, document.MarginBottomPoints);

        return new DocxPageGeometry(
            width,
            height,
            marginLeft,
            marginRight,
            marginTop,
            marginBottom,
            effectiveSettings,
            section.SectionProperties,
            CreateColumnFrames(
                width,
                marginLeft,
                marginRight,
                section.SectionProperties));
    }

    private static IReadOnlyList<DocxLayoutColumnFrame> CreateColumnFrames(
        double pageWidth,
        double marginLeft,
        double marginRight,
        DocxSectionLayoutProperties section)
    {
        double bodyWidth = Math.Max(1d, pageWidth - marginLeft - marginRight);
        int columnCount = Math.Max(1, section.ColumnCount ?? 1);
        if (columnCount == 1)
        {
            return [new DocxLayoutColumnFrame(0, marginLeft, bodyWidth, null)];
        }

        if (string.Equals(section.ColumnEqualWidthValue, "0", StringComparison.OrdinalIgnoreCase))
        {
            if (section.ColumnDefinitions.Count == 0)
            {
                return [];
            }

            double x = marginLeft;
            var frames = new List<DocxLayoutColumnFrame>();
            for (int index = 0; index < section.ColumnDefinitions.Count; index++)
            {
                DocxSectionColumnLayoutProperties column = section.ColumnDefinitions[index];
                double customColumnWidth = Math.Max(1d, column.WidthPoints ?? 0d);
                double? customGutter = index + 1 < section.ColumnDefinitions.Count
                    ? Math.Max(0d, column.SpacePoints ?? 0d)
                    : null;
                frames.Add(new DocxLayoutColumnFrame(index, x, customColumnWidth, customGutter));
                x += customColumnWidth + (customGutter ?? 0d);
            }

            return frames;
        }

        double gutter = Math.Max(0d, section.ColumnSpacePoints ?? 0d);
        double columnWidth = Math.Max(1d, (bodyWidth - gutter * (columnCount - 1)) / columnCount);
        return Enumerable.Range(0, columnCount)
            .Select(index => new DocxLayoutColumnFrame(
                index,
                marginLeft + index * (columnWidth + gutter),
                columnWidth,
                index + 1 < columnCount ? gutter : null))
            .ToArray();
    }

    private static DocxLayoutColumnFrame ResolveActiveColumnFrame(DocxPageGeometry page, int activeColumnIndex)
    {
        if (page.ColumnFrames.Count == 0)
        {
            return new DocxLayoutColumnFrame(0, page.MarginLeft, page.BodyWidth, null);
        }

        int index = Math.Clamp(activeColumnIndex, 0, page.ColumnFrames.Count - 1);
        return page.ColumnFrames[index];
    }

    private static double ReadTwipsValue(string? value, double fallback)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long twips)
            ? OoxUnits.TwipsToPoints(twips)
            : fallback;
    }

    private static double? ReadOptionalTwipsValue(string? value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long twips)
            ? OoxUnits.TwipsToPoints(twips)
            : null;
    }

    private static int? ReadOptionalInt32Value(string? value)
    {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result)
            ? result
            : null;
    }

    private static (double Width, double Height) NormalizePageSize(double width, double height)
    {
        if (Math.Abs(width - 595d) < 0.01d && Math.Abs(height - 842d) < 0.01d)
        {
            return (594.96d, 842.04d);
        }

        return (width, height);
    }

    private static bool ShouldStartNewPageForSectionBreak(DocxSectionBreakElement sectionBreak)
    {
        return sectionBreak.TypeValue is null ||
            sectionBreak.TypeValue.Equals("nextPage", StringComparison.OrdinalIgnoreCase) ||
            sectionBreak.TypeValue.Equals("oddPage", StringComparison.OrdinalIgnoreCase) ||
            sectionBreak.TypeValue.Equals("evenPage", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsContinuousSectionBreak(DocxSectionBreakElement sectionBreak)
    {
        return sectionBreak.TypeValue?.Equals("continuous", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static bool ShouldInsertParityBlankPage(DocxSectionBreakElement sectionBreak, int nextPageNumber)
    {
        if (sectionBreak.TypeValue?.Equals("oddPage", StringComparison.OrdinalIgnoreCase) == true)
        {
            return nextPageNumber % 2 == 0;
        }

        if (sectionBreak.TypeValue?.Equals("evenPage", StringComparison.OrdinalIgnoreCase) == true)
        {
            return nextPageNumber % 2 != 0;
        }

        return false;
    }

    private static bool ShouldKeepParagraphBlockTogether(DocxParagraph paragraph)
    {
        return paragraph.KeepRules.KeepLines == true || paragraph.KeepRules.KeepNext == true;
    }

    private static double ResolveLineHeight(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer? textMeasurer)
    {
        return ResolveLineHeightProfile(paragraph, fontSize, textMeasurer).LineHeight;
    }

    private static DocxLineHeightProfile ResolveLineHeightProfile(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer? textMeasurer)
    {
        if (paragraph.LineSpacingPoints is { } exactLineHeight)
        {
            return new DocxLineHeightProfile(
                exactLineHeight,
                SingleLineHeight: null,
                ListLabelSingleLineHeight: null,
                BodyWindowsLineHeight: null,
                ListLabelWindowsLineHeight: null,
                EffectiveLineSpacingFactor: null,
                LineSpacingFactorFloorApplied: false);
        }

        DocxTextRun? bodyRun = paragraph.Runs.FirstOrDefault();
        DocxTextRun? listLabelRun = paragraph.ListLabel is null
            ? null
            : CreateListLabelRun(paragraph.ListLabel, bodyRun, fontSize);
        IDocxLineMetricsProvider? metricsProvider = textMeasurer as IDocxLineMetricsProvider;
        IDocxStaticTextMetricsProvider? staticMetrics = textMeasurer as IDocxStaticTextMetricsProvider;
        double singleLineHeight = metricsProvider is not null
            ? metricsProvider.MeasureSingleLineHeight(bodyRun, fontSize)
            : fontSize;
        double? listLabelSingleLineHeight = metricsProvider is not null && listLabelRun is not null
            ? metricsProvider.MeasureSingleLineHeight(listLabelRun, listLabelRun.FontSize)
            : null;
        double? bodyWindowsLineHeight = staticMetrics is not null
            ? staticMetrics.MeasureWindowsAscender(bodyRun, fontSize) + staticMetrics.MeasureWindowsDescender(bodyRun, fontSize)
            : null;
        double? listLabelWindowsLineHeight = staticMetrics is not null && listLabelRun is not null
            ? staticMetrics.MeasureWindowsAscender(listLabelRun, listLabelRun.FontSize) + staticMetrics.MeasureWindowsDescender(listLabelRun, listLabelRun.FontSize)
            : null;
        double effectiveLineSpacingFactor = ResolveAutoLineSpacingFactor(paragraph, out bool floorApplied);
        return new DocxLineHeightProfile(
            singleLineHeight * effectiveLineSpacingFactor,
            singleLineHeight,
            listLabelSingleLineHeight,
            bodyWindowsLineHeight,
            listLabelWindowsLineHeight,
            effectiveLineSpacingFactor,
            floorApplied);
    }

    private static double ResolveAutoLineSpacingFactor(DocxParagraph paragraph, out bool floorApplied)
    {
        if (paragraph.ListLabel is not null &&
            paragraph.LineSpacingPoints is null &&
            paragraph.SpacingBeforePoints > 0d &&
            paragraph.Spacing.LineRuleValue?.Equals("auto", StringComparison.OrdinalIgnoreCase) == true)
        {
            double effectiveFactor = Math.Max(paragraph.LineSpacingFactor, WordListMinimumAutoLineSpacingFactor);
            floorApplied = effectiveFactor > paragraph.LineSpacingFactor;
            return effectiveFactor;
        }

        floorApplied = false;
        return paragraph.LineSpacingFactor;
    }

    private static double RoundToTwips(double points)
    {
        return Math.Round(points * 20d, MidpointRounding.AwayFromZero) / 20d;
    }

    private static double QuantizeTableCellWrappedLineHeight(double lineHeight, int wrappedLineCount)
    {
        return wrappedLineCount > 1 ? RoundToTwips(lineHeight) : lineHeight;
    }

    private static bool ShouldMoveParagraphForWidowControl(
        DocxParagraph paragraph,
        int lineCount,
        double cursorY,
        double lineHeight,
        double marginBottom,
        bool hasPageContent)
    {
        if (paragraph.KeepRules.WidowControl == false ||
            lineCount <= 1 ||
            !hasPageContent)
        {
            return false;
        }

        int fittingLineCount = (int)Math.Floor(Math.Max(0d, cursorY - marginBottom) / lineHeight);
        return fittingLineCount > 0 &&
            fittingLineCount < lineCount &&
            (fittingLineCount == 1 || lineCount - fittingLineCount == 1);
    }

    private static bool ShouldSuppressContextualSpacing(DocxParagraph? previousParagraph, DocxParagraph paragraph)
    {
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        DocxEffectiveParagraphProperties? previousEffective = previousParagraph?.EffectiveProperties;
        return effective.Spacing.ContextualSpacing == true &&
            previousEffective?.StyleId is not null &&
            effective.StyleId is not null &&
            string.Equals(previousEffective.StyleId, effective.StyleId, StringComparison.Ordinal);
    }

    private static DocxParagraphSpacingProfile ResolveParagraphSpacingProfile(
        DocxParagraph? previousParagraph,
        DocxParagraph paragraph,
        double pendingAfterSpacing)
    {
        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        bool suppress = ShouldSuppressContextualSpacing(previousParagraph, paragraph);
        double appliedBefore = suppress
            ? 0d
            : Math.Max(pendingAfterSpacing, effective.SpacingBeforePoints);
        return new DocxParagraphSpacingProfile(
            pendingAfterSpacing,
            effective.SpacingBeforePoints,
            effective.SpacingAfterPoints,
            appliedBefore,
            suppress);
    }

    private static DocxKeepBlockEstimate EstimateKeptParagraphBlock(
        IReadOnlyList<DocxBodyElement> elements,
        int elementIndex,
        double availableWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints)
    {
        if (elements[elementIndex] is not DocxParagraphElement paragraphElement)
        {
            return new DocxKeepBlockEstimate(0d, 0, 0);
        }

        DocxParagraph paragraph = paragraphElement.Paragraph;
        double height = EstimateParagraphContentHeight(paragraph, availableWidth, textMeasurer, defaultTabStopPoints);
        int paragraphCount = 1;
        int firstTableRowCount = 0;
        int nextSearchIndex = elementIndex + 1;
        while (paragraph.EffectiveProperties.KeepRules.KeepNext == true &&
            TryFindNextKeepTarget(elements, nextSearchIndex, out int nextIndex, out DocxBodyElement? next))
        {
            if (next is DocxParagraphElement nextParagraph)
            {
                DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
                DocxEffectiveParagraphProperties nextEffective = nextParagraph.Paragraph.EffectiveProperties;
                height += Math.Max(effective.SpacingAfterPoints, nextEffective.SpacingBeforePoints);
                height += EstimateParagraphContentHeight(nextParagraph.Paragraph, availableWidth, textMeasurer, defaultTabStopPoints);
                paragraphCount++;
                paragraph = nextParagraph.Paragraph;
                nextSearchIndex = nextIndex + 1;
                continue;
            }

            if (next is DocxTableElement nextTable)
            {
                height += paragraph.EffectiveProperties.SpacingAfterPoints;
                height += EstimateFirstTableRowHeight(nextTable.Table, availableWidth, textMeasurer, defaultTabStopPoints);
                firstTableRowCount++;
            }

            break;
        }

        return new DocxKeepBlockEstimate(height, paragraphCount, firstTableRowCount);
    }

    private static bool TryFindNextKeepTarget(IReadOnlyList<DocxBodyElement> elements, int startIndex, out DocxBodyElement? target)
    {
        bool found = TryFindNextKeepTarget(elements, startIndex, out _, out DocxBodyElement? indexedTarget);
        target = indexedTarget;
        return found;
    }

    private static bool TryFindNextKeepTarget(IReadOnlyList<DocxBodyElement> elements, int startIndex, out int targetIndex, out DocxBodyElement? target)
    {
        for (int i = startIndex; i < elements.Count; i++)
        {
            if (elements[i] is DocxParagraphElement or DocxTableElement)
            {
                targetIndex = i;
                target = elements[i];
                return true;
            }

            if (elements[i] is DocxPageBreakElement or DocxManualBreakElement or DocxSectionBreakElement)
            {
                break;
            }
        }

        targetIndex = -1;
        target = null;
        return false;
    }

    private static double EstimateParagraphContentHeight(DocxParagraph paragraph, double availableWidth, IDocxTextMeasurer textMeasurer, double defaultTabStopPoints)
    {
        double height = 0d;
        double fontSize = GetParagraphFontSize(paragraph);
        double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
        if (textSpans.Count != 0)
        {
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double firstParagraphWidth = Math.Max(1d, availableWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = Math.Max(1d, availableWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            height += WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: false).Count() * lineHeight;
        }
        else if (paragraph.Images.Count == 0)
        {
            height += lineHeight;
        }

        foreach (DocxInlineImage image in paragraph.Images)
        {
            double imageWidth = Math.Min(availableWidth, image.WidthPoints);
            double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
            height += imageHeight + InlineImageParagraphGapPoints;
        }

        return height;
    }

    private static double EstimateFirstTableRowHeight(DocxTable table, double availableWidth, IDocxTextMeasurer textMeasurer, double defaultTabStopPoints)
    {
        DocxTableRow? row = table.Rows.FirstOrDefault();
        if (row is null)
        {
            return 0d;
        }

        double tableAvailableWidth = Math.Max(1d, availableWidth - Math.Max(0d, table.IndentPoints ?? 0d));
        double gridTableWidth = table.ColumnWidthsPoints.Sum();
        double fallbackTableWidth = table.HasExplicitGrid && gridTableWidth > 0d ? gridTableWidth : tableAvailableWidth;
        double targetTableWidth = ResolveTargetTableWidth(table, tableAvailableWidth, fallbackTableWidth);
        IReadOnlyList<double> effectiveColumns = GetEffectiveTableColumnWidths(table, targetTableWidth);
        double rawTableWidth = effectiveColumns.Sum();
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
        double contentHeight = row.Cells
            .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer, defaultTabStopPoints, rowTopPadding))
            .DefaultIfEmpty(0d)
            .Max();
        return ResolveTableRowHeight(row, contentHeight);
    }

    private static double GetParagraphTextStartOffset(DocxParagraph paragraph)
    {
        if (paragraph.ListLabel is null)
        {
            return Math.Max(0d, paragraph.Indent.LeftPoints ?? 0d);
        }

        DocxNumberingIndent indent = paragraph.ListLabel.Indent;
        double left = indent.LeftPoints ?? 0d;
        double firstLine = indent.FirstLinePoints ?? 0d;
        return Math.Max(0d, left + firstLine);
    }

    private static double GetParagraphFirstLineTextStartOffset(DocxParagraph paragraph, double fontSize, IDocxTextMeasurer textMeasurer)
    {
        if (paragraph.ListLabel is null)
        {
            return GetParagraphFirstLineIndentOffset(paragraph);
        }

        if (IsNumberingTabSuffix(paragraph.ListLabel))
        {
            return GetParagraphTextStartOffset(paragraph);
        }

        double gap = IsNumberingSpaceSuffix(paragraph.ListLabel)
            ? textMeasurer.MeasureText(paragraph.Runs.FirstOrDefault(), " ", fontSize)
            : 0d;
        DocxTextRun labelRun = CreateListLabelRun(paragraph.ListLabel, paragraph.Runs.FirstOrDefault(), fontSize);
        return Math.Max(
            0d,
            GetParagraphLabelStartOffset(paragraph) + textMeasurer.MeasureText(labelRun, paragraph.ListLabel.Text, labelRun.FontSize) + gap);
    }

    private static double GetParagraphLabelStartOffset(DocxParagraph paragraph)
    {
        if (paragraph.ListLabel is null)
        {
            return 0d;
        }

        DocxNumberingIndent indent = paragraph.ListLabel.Indent;
        double left = indent.LeftPoints ?? 0d;
        double hanging = indent.HangingPoints ?? 0d;
        double firstLine = indent.FirstLinePoints ?? 0d;
        return Math.Max(0d, left - hanging + firstLine);
    }

    private static double GetParagraphStartOffset(DocxParagraph paragraph)
    {
        return paragraph.ListLabel is null
            ? GetParagraphFirstLineIndentOffset(paragraph)
            : GetParagraphLabelStartOffset(paragraph);
    }

    private static double GetParagraphRightInset(DocxParagraph paragraph)
    {
        return paragraph.ListLabel?.Indent.RightPoints ?? paragraph.Indent.RightPoints ?? 0d;
    }

    private static double GetParagraphFirstLineIndentOffset(DocxParagraph paragraph)
    {
        double left = paragraph.Indent.LeftPoints ?? 0d;
        double firstLine = paragraph.Indent.FirstLinePoints ?? 0d;
        double hanging = paragraph.Indent.HangingPoints ?? 0d;
        return Math.Max(0d, left + firstLine - hanging);
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateNumberedLineSegments(
        DocxListLabel label,
        IReadOnlyList<DocxTextSpan> lineSpans,
        DocxTextRun styleRun,
        double labelX,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        DocxTextRun labelRun = CreateListLabelRun(label, styleRun, fontSize);
        double labelWidth = textMeasurer.MeasureText(labelRun, label.Text, labelRun.FontSize);
        DocxTextEmissionPlan labelPlan = DocxTextEmissionPlanner.CreateForListLabel(labelRun, label);
        var segments = new List<DocxTextSegmentLayout>
        {
            new(
                label.Text,
                labelRun,
                labelX,
                labelWidth,
                labelRun.FontSize,
                PdfCharacterSpacing: labelPlan.PdfCharacterSpacing,
                PdfCharacterSpacingSource: labelPlan.PdfCharacterSpacingSource,
                CompensatePdfCharacterSpacing: labelPlan.CompensatePdfCharacterSpacing,
                SourceTextRunIndex: -1,
                Role: DocxTextSegmentRole.ListLabel)
        };

        string separator = GetListLabelPdfSeparator(label);
        if (separator.Length != 0)
        {
            double separatorX = labelX + labelWidth;
            double separatorWidth = textMeasurer.MeasureText(labelRun, separator, labelRun.FontSize);
            segments.Add(new DocxTextSegmentLayout(separator, labelRun, separatorX, separatorWidth, labelRun.FontSize, SourceTextRunIndex: -1, Role: DocxTextSegmentRole.ListSeparator));
        }

        segments.AddRange(CreateTextSegments(lineSpans, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints));
        return segments;
    }

    private static double MeasureListLabel(DocxListLabel label, DocxTextRun? baseRun, double fontSize, IDocxTextMeasurer textMeasurer)
    {
        DocxTextRun labelRun = CreateListLabelRun(label, baseRun, fontSize);
        return textMeasurer.MeasureText(labelRun, label.Text, labelRun.FontSize);
    }

    internal static DocxTextRun CreateListLabelRun(DocxListLabel label, DocxTextRun? baseRun, double fontSize)
    {
        return label.Style.ApplyTo(baseRun, label.Text, fontSize);
    }

    private static string GetListLabelTextSeparator(DocxListLabel label)
    {
        return label.SuffixValue switch
        {
            "nothing" => string.Empty,
            "space" => " ",
            _ => "\t"
        };
    }

    private static string GetListLabelPdfSeparator(DocxListLabel label)
    {
        return label.SuffixValue.Equals("nothing", StringComparison.OrdinalIgnoreCase) ? string.Empty : " ";
    }

    private static bool IsNumberingTabSuffix(DocxListLabel label)
    {
        return string.IsNullOrEmpty(label.SuffixValue) ||
            label.SuffixValue.Equals("tab", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNumberingSpaceSuffix(DocxListLabel label)
    {
        return label.SuffixValue.Equals("space", StringComparison.OrdinalIgnoreCase);
    }

    private static void LayoutTable(
        DocxTable table,
        double marginBottom,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        Func<DocxTableLayoutFrame> resolveFrame,
        Action finishPage,
        Func<bool> hasPageContent,
        Action markBoundaryContent)
    {
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows = table.Rows
            .Select((row, rowIndex) => (row, rowIndex))
            .TakeWhile(entry => entry.row.IsHeader)
            .Select(entry => (entry.row, entry.rowIndex))
            .ToArray();
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            DocxTableLayoutFrame frame = resolveFrame();
            DocxTableRow row = table.Rows[rowIndex];
            IReadOnlyList<double> rowHeights = frame.RowHeights;
            double rowHeight = rowHeights[rowIndex];
            double remainingPageHeight = Math.Max(0d, cursorY - marginBottom);
            if (!row.CantSplit &&
                TryResolveExplicitTableCellPageBreakBoundaries(row, frame.EffectiveColumns, frame.Scale, rowHeight, textMeasurer, defaultTabStopPoints, out IReadOnlyList<double> explicitBreakBoundaries))
            {
                double explicitBreakFragmentHeight = explicitBreakBoundaries[0];
                if (explicitBreakFragmentHeight > remainingPageHeight && hasPageContent())
                {
                    finishPage();
                    if (!row.IsHeader)
                    {
                        frame = resolveFrame();
                        rowHeights = frame.RowHeights;
                        AddRepeatedTableHeaderRows(table, frame.Context, rowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, frame.TableX);
                        markBoundaryContent();
                    }
                }

                AddSplitTableRowLayout(table, row, rowIndex, headerRows, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, resolveFrame, explicitBreakBoundaries, "CellPageBreak", finishPage);
                markBoundaryContent();
                continue;
            }

            if (!row.CantSplit &&
                rowHeight > remainingPageHeight &&
                remainingPageHeight > 0.001d &&
                CanSplitTableRowAtPageBoundary(row, frame.EffectiveColumns, frame.Scale, rowHeight, remainingPageHeight, textMeasurer, defaultTabStopPoints))
            {
                AddSplitTableRowLayout(table, row, rowIndex, headerRows, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, resolveFrame, remainingPageHeight, "PageBoundary", finishPage);
                markBoundaryContent();
                continue;
            }

            if (cursorY - rowHeight < marginBottom && hasPageContent())
            {
                finishPage();
                if (!row.IsHeader)
                {
                    frame = resolveFrame();
                    rowHeights = frame.RowHeights;
                    AddRepeatedTableHeaderRows(table, frame.Context, rowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, frame.TableX);
                    markBoundaryContent();
                }
            }

            frame = resolveFrame();
            rowHeights = frame.RowHeights;
            AddTableRowLayout(table, frame.Context, row, rowIndex, rowHeights, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, frame.TableX);
            markBoundaryContent();
        }
    }

    private static DocxTableLayoutFrame CreateTableLayoutFrame(
        DocxTable table,
        int tableIndex,
        int sourceBlockIndex,
        double x,
        double availableWidth,
        double pageContentHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints)
    {
        double tableX = x + Math.Max(0d, table.IndentPoints ?? 0d);
        double tableAvailableWidth = Math.Max(1d, availableWidth - Math.Max(0d, table.IndentPoints ?? 0d));
        double gridTableWidth = table.ColumnWidthsPoints.Sum();
        double fallbackTableWidth = table.HasExplicitGrid && gridTableWidth > 0d ? gridTableWidth : tableAvailableWidth;
        double targetTableWidth = ResolveTargetTableWidth(table, tableAvailableWidth, fallbackTableWidth);
        IReadOnlyList<double> effectiveColumns = GetEffectiveTableColumnWidths(table, targetTableWidth);
        double rawTableWidth = effectiveColumns.Sum();
        double scale = rawTableWidth <= 0d ? 1d : targetTableWidth / rawTableWidth;
        var tableContext = new DocxTableLayoutContext(
            tableIndex,
            sourceBlockIndex,
            table.Rows.Count,
            table.ColumnWidthsPoints.Count,
            table.ColumnWidthsPoints.Sum(),
            table.HasExplicitGrid,
            effectiveColumns.Select(width => width * scale).ToArray(),
            targetTableWidth,
            tableX,
            table.PreferredWidthPoints,
            table.PreferredWidthValue,
            table.PreferredWidthType,
            table.IndentPoints,
            table.CellSpacingPoints,
            table.LayoutValue);
        double[] rowHeights = table.Rows
            .Select(row => MeasureTableRowHeight(table, row, effectiveColumns, scale, textMeasurer, defaultTabStopPoints))
            .ToArray();
        return new DocxTableLayoutFrame(tableContext, effectiveColumns, scale, rowHeights, pageContentHeight, tableX);
    }

    private static IReadOnlyList<double> GetEffectiveTableColumnWidths(DocxTable table, double preferredTableWidth)
    {
        int columnCount = table.ColumnWidthsPoints.Count;
        if (columnCount == 0)
        {
            int inferredColumnCount = GetMaxGridColumnCount(table);
            return inferredColumnCount > 0
                ? Enumerable.Repeat(preferredTableWidth / inferredColumnCount, inferredColumnCount).ToArray()
                : table.ColumnWidthsPoints;
        }

        double?[] preferredWidths = new double?[columnCount];
        foreach (DocxTableRow row in table.Rows)
        {
            int gridColumnIndex = 0;
            foreach (DocxTableCell cell in row.Cells)
            {
                int span = Math.Max(1, cell.GridSpan);
                double? preferredWidth = ResolvePreferredCellWidth(cell, preferredTableWidth);
                if (span == 1 &&
                    gridColumnIndex < columnCount &&
                    preferredWidth is > 0d)
                {
                    preferredWidths[gridColumnIndex] = preferredWidth.Value;
                }

                gridColumnIndex += span;
            }

            if (preferredWidths.All(width => width is > 0d))
            {
                return preferredWidths.Select(width => width!.Value).ToArray();
            }
        }

        if (!table.HasExplicitGrid)
        {
            int inferredColumnCount = columnCount == 0 ? GetMaxGridColumnCount(table) : columnCount;
            if (inferredColumnCount > 0)
            {
                return Enumerable.Repeat(preferredTableWidth / inferredColumnCount, inferredColumnCount).ToArray();
            }
        }

        return table.ColumnWidthsPoints;
    }

    private static int GetMaxGridColumnCount(DocxTable table)
    {
        return table.Rows
            .Select(row => row.Cells.Sum(cell => Math.Max(1, cell.GridSpan)))
            .DefaultIfEmpty(0)
            .Max();
    }

    private static double? ResolvePreferredCellWidth(DocxTableCell cell, double preferredTableWidth)
    {
        if (cell.PreferredWidthPoints is { } points)
        {
            return points;
        }

        if (cell.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true &&
            int.TryParse(cell.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
        {
            return Math.Max(0d, preferredTableWidth * fiftiethsPercent / 5000d);
        }

        return null;
    }

    private static double? ResolvePreferredTableWidth(DocxTable table, double availableWidth)
    {
        if (table.PreferredWidthPoints is { } points)
        {
            return points;
        }

        if (table.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true &&
            int.TryParse(table.PreferredWidthValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fiftiethsPercent))
        {
            double normalPercentageWidth = availableWidth * fiftiethsPercent / 5000d;
            double explicitGridWidth = table.HasExplicitGrid ? table.ColumnWidthsPoints.Sum() : 0d;
            double percentageBasis = explicitGridWidth > 0d && explicitGridWidth < normalPercentageWidth - 0.001d
                ? availableWidth + ResolveOuterTableCellContentInset(table)
                : availableWidth;
            return Math.Max(0d, percentageBasis * fiftiethsPercent / 5000d);
        }

        return null;
    }

    private static double ResolveOuterTableCellContentInset(DocxTable table)
    {
        DocxTableRow? firstRow = table.Rows.FirstOrDefault();
        if (firstRow is null || firstRow.Cells.Count == 0)
        {
            return 0d;
        }

        DocxTableCell firstCell = firstRow.Cells[0];
        DocxTableCell lastCell = firstRow.Cells[^1];
        return ResolveTableCellHorizontalPadding(firstCell.Margins.LeftPoints) +
            ResolveTableCellBorderContentInset(firstCell, "left") +
            ResolveTableCellHorizontalPadding(lastCell.Margins.RightPoints) +
            ResolveTableCellBorderContentInset(lastCell, "right");
    }

    private static double ResolveTargetTableWidth(DocxTable table, double availableWidth, double fallbackWidth)
    {
        double preferredWidth = ResolvePreferredTableWidth(table, availableWidth) ?? fallbackWidth;
        return table.PreferredWidthType?.Equals("dxa", StringComparison.OrdinalIgnoreCase) == true ||
            table.PreferredWidthType?.Equals("pct", StringComparison.OrdinalIgnoreCase) == true
            ? Math.Max(1d, preferredWidth)
            : Math.Min(availableWidth, preferredWidth);
    }

    private static double MeasureTableRowHeight(
        DocxTable table,
        DocxTableRow row,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints)
    {
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
        double contentHeight = textMeasurer is null
            ? 0d
            : row.Cells
                .Select((cell, columnIndex) => MeasureTableCellContentHeight(cell, cellWidths[columnIndex], textMeasurer, defaultTabStopPoints, rowTopPadding))
                .DefaultIfEmpty(0d)
                .Max();
        return ResolveTableRowHeight(row, contentHeight);
    }

    private static double ResolveTableRowHeight(DocxTableRow row, double contentHeight)
    {
        if (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) &&
            row.HeightPoints is { } exactHeight)
        {
            return Math.Max(1d, exactHeight);
        }

        double declaredHeight = string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase)
            ? 0d
            : row.HeightPoints ?? 0d;
        if (row.HeightPoints is not null &&
            !string.Equals(row.HeightRuleValue, "auto", StringComparison.OrdinalIgnoreCase))
        {
            declaredHeight += ResolveTableRowTopPadding(row);
        }

        double height = Math.Max(declaredHeight, contentHeight);
        height += ResolveTableRowCollapsedHorizontalBorderAdvance(row);
        return Math.Max(1d, height);
    }

    private static double ResolveTableRowCollapsedHorizontalBorderAdvance(DocxTableRow row)
    {
        double maxBottom = row.Cells
            .Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "bottom")))
            .DefaultIfEmpty(0d)
            .Max();
        if (maxBottom > 0d)
        {
            return maxBottom;
        }

        return row.Cells
            .Select(cell => DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, "top")))
            .DefaultIfEmpty(0d)
            .Max();
    }

    private static void AddTableRowLayout(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x)
    {
        double rowHeight = rowHeights[rowIndex];
        currentItems.Add(CreateTableRowLayout(
            table,
            tableContext,
            row,
            rowIndex,
            rowHeights,
            effectiveColumns,
            scale,
            textMeasurer,
            defaultTabStopPoints,
            getPageIndex,
            cursorY,
            rowHeight,
            logicalRowTopY: cursorY,
            FragmentIndex: 0,
            FragmentCount: 1,
            FragmentReason: "None"));
        cursorY -= rowHeight;
    }

    private static void AddSplitTableRowLayout(
        DocxTable table,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        Func<DocxTableLayoutFrame> resolveFrame,
        double firstFragmentHeight,
        string fragmentReason,
        Action finishPage)
    {
        AddSplitTableRowLayout(
            table,
            row,
            rowIndex,
            headerRows,
            textMeasurer,
            defaultTabStopPoints,
            getPageIndex,
            ref currentItems,
            ref cursorY,
            resolveFrame,
            [firstFragmentHeight],
            fragmentReason,
            finishPage);
    }

    private static void AddSplitTableRowLayout(
        DocxTable table,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        Func<DocxTableLayoutFrame> resolveFrame,
        IReadOnlyList<double> fragmentBoundariesFromRowTop,
        string fragmentReason,
        Action finishPage)
    {
        DocxTableLayoutFrame initialFrame = resolveFrame();
        IReadOnlyList<double> initialRowHeights = initialFrame.RowHeights;
        double rowHeight = initialRowHeights[rowIndex];
        double continuationContentHeight = row.IsHeader
            ? initialFrame.PageContentHeight
            : Math.Max(1d, initialFrame.PageContentHeight - SumRepeatedTableHeaderRowsHeight(initialRowHeights, headerRows));
        IReadOnlyList<double> fragmentHeights = ComputeTableRowFragmentHeights(rowHeight, fragmentBoundariesFromRowTop, continuationContentHeight);
        double consumedHeight = 0d;
        for (int fragmentIndex = 0; fragmentIndex < fragmentHeights.Count; fragmentIndex++)
        {
            DocxTableLayoutFrame frame = resolveFrame();
            double fragmentHeight = fragmentHeights[fragmentIndex];
            currentItems.Add(CreateTableRowLayout(
                table,
                frame.Context,
                row,
                rowIndex,
                frame.RowHeights,
                frame.EffectiveColumns,
                frame.Scale,
                textMeasurer,
                defaultTabStopPoints,
                getPageIndex,
                cursorY,
                fragmentHeight,
                logicalRowTopY: cursorY + consumedHeight,
                FragmentIndex: fragmentIndex,
                FragmentCount: fragmentHeights.Count,
                FragmentReason: fragmentReason));
            cursorY -= fragmentHeight;
            consumedHeight += fragmentHeight;

            if (fragmentIndex + 1 < fragmentHeights.Count)
            {
                finishPage();
                if (!row.IsHeader)
                {
                    frame = resolveFrame();
                    AddRepeatedTableHeaderRows(table, frame.Context, frame.RowHeights, headerRows, frame.EffectiveColumns, frame.Scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, frame.TableX);
                }
            }
        }
    }

    private static void AddRepeatedTableHeaderRows(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        ref List<DocxLayoutItem> currentItems,
        ref double cursorY,
        double x)
    {
        foreach ((DocxTableRow headerRow, int headerRowIndex) in headerRows)
        {
            AddTableRowLayout(table, tableContext, headerRow, headerRowIndex, rowHeights, effectiveColumns, scale, textMeasurer, defaultTabStopPoints, getPageIndex, ref currentItems, ref cursorY, x);
        }
    }

    private static double SumRepeatedTableHeaderRowsHeight(
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<(DocxTableRow Row, int RowIndex)> headerRows)
    {
        return headerRows.Sum(entry => entry.RowIndex >= 0 && entry.RowIndex < rowHeights.Count ? rowHeights[entry.RowIndex] : 0d);
    }

    private static bool CanSplitTableRowAtPageBoundary(
        DocxTableRow row,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        double rowHeight,
        double firstFragmentHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints)
    {
        if (textMeasurer is null)
        {
            return false;
        }

        double fragmentBottomY = rowHeight - firstFragmentHeight;
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCell cell = row.Cells[cellIndex];
            if (IsVerticalMergeContinuation(cell))
            {
                continue;
            }

            IReadOnlyList<DocxTextLineLayout> textLines = LayoutTableCellTextLines(cell, 0d, 0d, cellWidths[cellIndex], rowHeight, rowTopPadding, textMeasurer, defaultTabStopPoints);
            bool hasLineInFirstFragment = textLines.Any(line => firstFragmentHeight >= line.LineHeight && line.BaselineY >= fragmentBottomY);
            bool hasLineInContinuation = textLines.Any(line => line.BaselineY < fragmentBottomY);
            if (hasLineInFirstFragment && hasLineInContinuation)
            {
                return true;
            }
        }

        return false;
    }

    private static bool TryResolveExplicitTableCellPageBreakBoundaries(
        DocxTableRow row,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        double rowHeight,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        out IReadOnlyList<double> breakBoundariesFromRowTop)
    {
        breakBoundariesFromRowTop = [];
        if (textMeasurer is null || rowHeight <= 1.001d)
        {
            return false;
        }

        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
        var breakHeights = new List<double>();
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCell cell = row.Cells[cellIndex];
            if (IsVerticalMergeContinuation(cell) ||
                !TryMeasureTableCellHeightBeforePageBreak(cell, cellWidths[cellIndex], textMeasurer, defaultTabStopPoints, rowTopPadding, out double heightBeforeBreak))
            {
                continue;
            }

            double breakHeight = Math.Min(rowHeight - 1d, Math.Max(1d, heightBeforeBreak));
            if (breakHeight < rowHeight - 0.001d)
            {
                breakHeights.Add(breakHeight);
            }
        }

        if (breakHeights.Count == 0)
        {
            return false;
        }

        breakBoundariesFromRowTop = breakHeights
            .Order()
            .Aggregate(new List<double>(), (boundaries, breakHeight) =>
            {
                if (boundaries.Count == 0 || Math.Abs(boundaries[^1] - breakHeight) > 0.001d)
                {
                    boundaries.Add(breakHeight);
                }

                return boundaries;
            });
        return breakBoundariesFromRowTop.Count != 0;
    }

    private static bool TryMeasureTableCellHeightBeforePageBreak(
        DocxTableCell cell,
        double cellWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        double rowTopPadding,
        out double heightBeforeBreak)
    {
        heightBeforeBreak = 0d;
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return false;
        }

        int pageBreakIndex = -1;
        for (int index = 0; index < bodyElements.Count; index++)
        {
            if (bodyElements[index] is DocxPageBreakElement)
            {
                pageBreakIndex = index;
                break;
            }
        }

        if (pageBreakIndex <= 0 || !bodyElements.Skip(pageBreakIndex + 1).Any(IsRenderableTableCellBodyElement))
        {
            return false;
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        heightBeforeBreak = rowTopPadding;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        for (int index = 0; index < pageBreakIndex; index++)
        {
            DocxBodyElement bodyElement = bodyElements[index];
            if (bodyElement is DocxTableElement tableElement)
            {
                heightBeforeBreak += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                heightBeforeBreak += MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints);
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            heightBeforeBreak += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            heightBeforeBreak += MeasureTableCellParagraphContentHeight(paragraph, textWidth, textMeasurer, defaultTabStopPoints);
            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        heightBeforeBreak += pendingSpacingAfter;
        return heightBeforeBreak > rowTopPadding + 0.001d;
    }

    private static bool IsRenderableTableCellBodyElement(DocxBodyElement bodyElement)
    {
        return bodyElement switch
        {
            DocxParagraphElement paragraphElement => paragraphElement.Paragraph.Runs.Count != 0 ||
                paragraphElement.Paragraph.Images.Count != 0 ||
                paragraphElement.Paragraph.ListLabel is not null,
            DocxTableElement => true,
            _ => false
        };
    }

    private static IReadOnlyList<double> ComputeTableRowFragmentHeights(double rowHeight, double firstFragmentHeight, double pageContentHeight)
    {
        return ComputeTableRowFragmentHeights(rowHeight, [firstFragmentHeight], pageContentHeight);
    }

    private static IReadOnlyList<double> ComputeTableRowFragmentHeights(double rowHeight, IReadOnlyList<double> fragmentBoundariesFromRowTop, double pageContentHeight)
    {
        var fragments = new List<double>();
        double consumedHeight = 0d;
        double fullPageHeight = Math.Max(1d, pageContentHeight);
        foreach (double boundary in fragmentBoundariesFromRowTop.Order())
        {
            double clampedBoundary = Math.Min(rowHeight, Math.Max(0d, boundary));
            if (clampedBoundary <= consumedHeight + 0.001d)
            {
                continue;
            }

            AddTableRowFragmentSegmentHeights(fragments, clampedBoundary - consumedHeight, fullPageHeight);
            consumedHeight = clampedBoundary;
        }

        AddTableRowFragmentSegmentHeights(fragments, rowHeight - consumedHeight, fullPageHeight);
        return fragments.Count == 0 ? [Math.Max(1d, rowHeight)] : fragments;
    }

    private static void AddTableRowFragmentSegmentHeights(List<double> fragments, double segmentHeight, double fullPageHeight)
    {
        double remainingHeight = segmentHeight;
        while (remainingHeight > fullPageHeight + 0.001d)
        {
            fragments.Add(fullPageHeight);
            remainingHeight -= fullPageHeight;
        }

        if (remainingHeight > 0.001d)
        {
            fragments.Add(remainingHeight);
        }
    }

    private static DocxTableRowLayout CreateTableRowLayout(
        DocxTable table,
        DocxTableLayoutContext tableContext,
        DocxTableRow row,
        int rowIndex,
        IReadOnlyList<double> rowHeights,
        IReadOnlyList<double> effectiveColumns,
        double scale,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        Func<int> getPageIndex,
        double cursorY,
        double rowHeight,
        double logicalRowTopY,
        int FragmentIndex,
        int FragmentCount,
        string FragmentReason,
        string? StoryKind = null,
        string? StoryVariantType = null)
    {
        double[] cellWidths = GetTableRowCellWidths(row, effectiveColumns, scale);
        double rowTopPadding = ResolveTableRowTopPadding(row);
        double fullRowHeight = rowHeights[rowIndex];
        double fragmentOffsetFromRowTop = logicalRowTopY - cursorY;
        double cellX = tableContext.TableX;
        double cellY = cursorY - rowHeight;
        double fullCellY = logicalRowTopY - fullRowHeight;
        var cells = new List<DocxTableCellLayout>(row.Cells.Count);
        int gridColumnIndex = 0;
        for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
        {
            double cellWidth = cellWidths[columnIndex];
            DocxTableCell cell = row.Cells[columnIndex];
            bool isVerticalMergeContinuation = IsVerticalMergeContinuation(cell);
            DocxVerticalMergeOwner? verticalMergeOwner = isVerticalMergeContinuation
                ? FindVerticalMergeRestartOwner(table, rowIndex, gridColumnIndex)
                : null;
            DocxTableCell? verticalMergeOwnerCell = verticalMergeOwner?.Cell;
            DocxTableCellVisualOwnership visualOwnership = isVerticalMergeContinuation
                ? verticalMergeOwnerCell is null
                    ? DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                    : DocxTableCellVisualOwnership.VerticalMergeOwner
                : DocxTableCellVisualOwnership.OwnCell;
            double visualHeight = rowHeight;
            double visualY = cellY;
            double fullVisualHeight = fullRowHeight;
            double fullVisualY = fullCellY;
            if (IsVerticalMergeRestart(cell))
            {
                fullVisualHeight = GetVerticalMergeSpanHeight(table, rowIndex, gridColumnIndex, rowHeights);
                fullVisualY = cursorY - fullVisualHeight;
                if (FragmentCount == 1)
                {
                    visualHeight = fullVisualHeight;
                    visualY = fullVisualY;
                }
            }

            bool useCellPageBreakBoundaryPartition = FragmentReason == "CellPageBreak" && textMeasurer is not null;
            int cellPageBreakLowerParagraphBoundaryIndex = 0;
            int? cellPageBreakUpperParagraphBoundaryIndex = null;
            int cellPageBreakLowerNestedTableBoundaryIndex = 0;
            int? cellPageBreakUpperNestedTableBoundaryIndex = null;
            bool cellPageBreakLowerBoundaryInsideNestedTable = false;
            bool cellPageBreakUpperBoundaryInsideNestedTable = false;
            if (useCellPageBreakBoundaryPartition)
            {
                if (fragmentOffsetFromRowTop > 0.001d &&
                    TryResolveTableCellParagraphBoundaryIndex(cell, cellWidth, rowTopPadding, fragmentOffsetFromRowTop, textMeasurer!, defaultTabStopPoints, out int lowerParagraphBoundaryIndex))
                {
                    cellPageBreakLowerParagraphBoundaryIndex = lowerParagraphBoundaryIndex;
                }

                if (fragmentOffsetFromRowTop > 0.001d &&
                    TryResolveTableCellNestedTableBoundary(cell, cellWidth, rowTopPadding, fragmentOffsetFromRowTop, textMeasurer!, defaultTabStopPoints, out DocxNestedTableBoundary lowerNestedTableBoundary))
                {
                    cellPageBreakLowerNestedTableBoundaryIndex = lowerNestedTableBoundary.BoundaryIndex;
                    cellPageBreakLowerBoundaryInsideNestedTable = lowerNestedTableBoundary.IsInsideNestedTable;
                }

                double fragmentEndFromRowTop = fragmentOffsetFromRowTop + rowHeight;
                if (fragmentEndFromRowTop < fullRowHeight - 0.001d &&
                    TryResolveTableCellParagraphBoundaryIndex(cell, cellWidth, rowTopPadding, fragmentEndFromRowTop, textMeasurer!, defaultTabStopPoints, out int upperParagraphBoundaryIndex))
                {
                    cellPageBreakUpperParagraphBoundaryIndex = upperParagraphBoundaryIndex;
                }

                if (fragmentEndFromRowTop < fullRowHeight - 0.001d &&
                    TryResolveTableCellNestedTableBoundary(cell, cellWidth, rowTopPadding, fragmentEndFromRowTop, textMeasurer!, defaultTabStopPoints, out DocxNestedTableBoundary upperNestedTableBoundary))
                {
                    cellPageBreakUpperNestedTableBoundaryIndex = upperNestedTableBoundary.BoundaryIndex + (upperNestedTableBoundary.IsInsideNestedTable ? 1 : 0);
                    cellPageBreakUpperBoundaryInsideNestedTable = upperNestedTableBoundary.IsInsideNestedTable;
                }
            }

            bool cellPageBreakBoundaryInsideNestedTable =
                cellPageBreakLowerBoundaryInsideNestedTable ||
                cellPageBreakUpperBoundaryInsideNestedTable;
            bool cellPageBreakAlignsWithNestedTableBlock =
                useCellPageBreakBoundaryPartition &&
                !cellPageBreakBoundaryInsideNestedTable;

            DocxTableCell contentCell = visualOwnership == DocxTableCellVisualOwnership.VerticalMergeOwner && verticalMergeOwnerCell is not null
                ? verticalMergeOwnerCell
                : cell;
            double contentY = isVerticalMergeContinuation ? visualY : fullVisualY;
            double contentHeight = isVerticalMergeContinuation ? visualHeight : fullVisualHeight;
            double contentPaddingLeft = ResolveTableCellHorizontalPadding(contentCell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(contentCell, "left");
            double contentPaddingTop = rowTopPadding;
            double contentPaddingRight = ResolveTableCellHorizontalPadding(contentCell.Margins.RightPoints) + ResolveTableCellBorderContentInset(contentCell, "right");
            double contentPaddingBottom = ResolveTableCellVerticalPadding(contentCell.Margins.BottomPoints);
            IReadOnlyList<DocxTextLineLayout> textLines = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? []
                : LayoutTableCellTextLines(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints)
                    .Where(line => IsTextLineOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, cell, line, FragmentIndex, FragmentCount))
                    .Where(line => IsTextLineVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, line, visualY, visualHeight, FragmentIndex, FragmentCount))
                    .ToArray();
            IReadOnlyList<DocxInlineImageLayout> inlineImages = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? []
                : LayoutTableCellInlineImages(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, getPageIndex())
                    .Where(image => IsInlineImageOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerParagraphBoundaryIndex, cellPageBreakUpperParagraphBoundaryIndex, cell, image, FragmentIndex, FragmentCount))
                    .Where(image => IsInlineImageVisibleInCellFragmentGeometry(useCellPageBreakBoundaryPartition, image, visualY, visualHeight, FragmentIndex, FragmentCount))
                    .ToArray();
            IReadOnlyList<DocxTableRowLayout> nestedTableRows = visualOwnership == DocxTableCellVisualOwnership.MissingVerticalMergeOwner
                ? []
                : LayoutTableCellNestedTables(contentCell, cellX, contentY, cellWidth, contentHeight, rowTopPadding, textMeasurer, defaultTabStopPoints, getPageIndex())
                    .Where(rowLayout => IsNestedTableRowOnVisibleSideOfCellPageBreak(useCellPageBreakBoundaryPartition, cellPageBreakLowerNestedTableBoundaryIndex, cellPageBreakUpperNestedTableBoundaryIndex, rowLayout, FragmentCount))
                    .Where(rowLayout => IsNestedTableRowVisibleInCellFragmentGeometry(cellPageBreakAlignsWithNestedTableBlock, rowLayout, visualY, visualHeight, FragmentIndex, FragmentCount))
                    .ToArray();
            cells.Add(new DocxTableCellLayout(
                cell,
                cellX,
                visualY,
                cellWidth,
                visualHeight,
                contentPaddingLeft,
                contentPaddingTop,
                contentPaddingRight,
                contentPaddingBottom,
                textLines,
                inlineImages,
                isVerticalMergeContinuation,
                verticalMergeOwnerCell,
                verticalMergeOwner,
                visualOwnership,
                nestedTableRows));
            cellX += cellWidth + (table.CellSpacingPoints ?? 0d);
            gridColumnIndex += Math.Max(1, cell.GridSpan);
        }

        return new DocxTableRowLayout(
            tableContext,
            rowIndex,
            FragmentIndex,
            FragmentCount,
            FragmentReason,
            fullRowHeight,
            fragmentOffsetFromRowTop,
            cells.ToArray(),
            cellY,
            rowHeight,
            row.HeightPoints,
            row.HeightValue,
            row.HeightRuleValue,
            row.IsHeader,
            row.HeaderValue,
            row.TablePropertyExceptionCellMargins is not null,
            row.CantSplit,
            row.CantSplitValue,
            StoryKind,
            StoryVariantType);
    }

    private static bool IsVerticalMergeRestart(DocxTableCell cell)
    {
        return cell.HasVerticalMerge &&
            string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsTextLineVisibleInCellFragment(
        DocxTextLineLayout line,
        double cellY,
        double cellHeight,
        int fragmentIndex,
        int fragmentCount)
    {
        if (fragmentCount <= 1)
        {
            return true;
        }

        double bottom = fragmentIndex == 0 ? cellY - 0.001d : cellY + 0.001d;
        return line.BaselineY >= bottom && line.BaselineY <= cellY + cellHeight + 0.001d;
    }

    private static bool IsTextLineVisibleInCellFragmentGeometry(
        bool cellPageBreakAlignsWithFragmentBoundary,
        DocxTextLineLayout line,
        double cellY,
        double cellHeight,
        int fragmentIndex,
        int fragmentCount)
    {
        return cellPageBreakAlignsWithFragmentBoundary
            ? true
            : IsTextLineVisibleInCellFragment(line, cellY, cellHeight, fragmentIndex, fragmentCount);
    }

    private static bool IsInlineImageVisibleInCellFragmentGeometry(
        bool cellPageBreakAlignsWithFragmentBoundary,
        DocxInlineImageLayout image,
        double cellY,
        double cellHeight,
        int fragmentIndex,
        int fragmentCount)
    {
        return cellPageBreakAlignsWithFragmentBoundary
            ? true
            : VerticalOverlap(image.Y, image.Height, cellY, cellHeight) > 0.001d;
    }

    private static bool IsNestedTableRowVisibleInCellFragmentGeometry(
        bool cellPageBreakAlignsWithFragmentBoundary,
        DocxTableRowLayout row,
        double cellY,
        double cellHeight,
        int fragmentIndex,
        int fragmentCount)
    {
        return cellPageBreakAlignsWithFragmentBoundary
            ? true
            : VerticalOverlap(row.Y, row.Height, cellY, cellHeight) > 0.001d;
    }

    private static bool TryResolveTableCellParagraphBoundaryIndex(
        DocxTableCell cell,
        double cellWidth,
        double rowTopPadding,
        double fragmentBoundaryFromRowTop,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        out int paragraphBoundaryIndex)
    {
        paragraphBoundaryIndex = 0;
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return false;
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double consumedHeight = rowTopPadding;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                consumedHeight += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                consumedHeight += MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints);
                if (consumedHeight >= fragmentBoundaryFromRowTop - 0.001d)
                {
                    return true;
                }

                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            consumedHeight += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            consumedHeight += MeasureTableCellParagraphContentHeight(paragraph, textWidth, textMeasurer, defaultTabStopPoints);
            paragraphBoundaryIndex++;
            if (consumedHeight >= fragmentBoundaryFromRowTop - 0.001d)
            {
                return true;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        return paragraphBoundaryIndex > 0;
    }

    private sealed record DocxNestedTableBoundary(int BoundaryIndex, bool IsInsideNestedTable);

    private static bool TryResolveTableCellNestedTableBoundary(
        DocxTableCell cell,
        double cellWidth,
        double rowTopPadding,
        double fragmentBoundaryFromRowTop,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        out DocxNestedTableBoundary boundary)
    {
        boundary = new DocxNestedTableBoundary(0, IsInsideNestedTable: false);
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return false;
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double consumedHeight = rowTopPadding;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int nestedTableIndex = 0;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                consumedHeight += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                double tableTop = consumedHeight;
                double tableHeight = MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints);
                consumedHeight += tableHeight;
                if (consumedHeight >= fragmentBoundaryFromRowTop - 0.001d)
                {
                    bool boundaryInsideTable =
                        fragmentBoundaryFromRowTop > tableTop + 0.001d &&
                        fragmentBoundaryFromRowTop < consumedHeight - 0.001d;
                    boundary = new DocxNestedTableBoundary(
                        boundaryInsideTable ? nestedTableIndex : nestedTableIndex + 1,
                        boundaryInsideTable);
                    return true;
                }

                nestedTableIndex++;
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            consumedHeight += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            consumedHeight += MeasureTableCellParagraphContentHeight(paragraph, textWidth, textMeasurer, defaultTabStopPoints);
            if (consumedHeight >= fragmentBoundaryFromRowTop - 0.001d)
            {
                boundary = new DocxNestedTableBoundary(nestedTableIndex, IsInsideNestedTable: false);
                return true;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        boundary = new DocxNestedTableBoundary(nestedTableIndex, IsInsideNestedTable: false);
        return nestedTableIndex > 0;
    }

    private static bool IsTextLineOnVisibleSideOfCellPageBreak(
        bool useCellPageBreakBoundaryPartition,
        int lowerParagraphBoundaryIndex,
        int? upperParagraphBoundaryIndex,
        DocxTableCell cell,
        DocxTextLineLayout line,
        int fragmentIndex,
        int fragmentCount)
    {
        if (!useCellPageBreakBoundaryPartition ||
            fragmentCount <= 1 ||
            line.SourceParagraphIndex is not { } paragraphIndex)
        {
            return true;
        }

        return paragraphIndex >= lowerParagraphBoundaryIndex &&
            (upperParagraphBoundaryIndex is not { } upper || paragraphIndex < upper);
    }

    private static bool IsInlineImageOnVisibleSideOfCellPageBreak(
        bool useCellPageBreakBoundaryPartition,
        int lowerParagraphBoundaryIndex,
        int? upperParagraphBoundaryIndex,
        DocxTableCell cell,
        DocxInlineImageLayout image,
        int fragmentIndex,
        int fragmentCount)
    {
        if (!useCellPageBreakBoundaryPartition ||
            fragmentCount <= 1 ||
            image.SourceParagraphIndex is not { } paragraphIndex)
        {
            return true;
        }

        return paragraphIndex >= lowerParagraphBoundaryIndex &&
            (upperParagraphBoundaryIndex is not { } upper || paragraphIndex < upper);
    }

    private static bool IsNestedTableRowOnVisibleSideOfCellPageBreak(
        bool useCellPageBreakBoundaryPartition,
        int lowerNestedTableBoundaryIndex,
        int? upperNestedTableBoundaryIndex,
        DocxTableRowLayout row,
        int fragmentCount)
    {
        if (!useCellPageBreakBoundaryPartition || fragmentCount <= 1)
        {
            return true;
        }

        return row.Table.TableIndex >= lowerNestedTableBoundaryIndex &&
            (upperNestedTableBoundaryIndex is not { } upper || row.Table.TableIndex < upper);
    }

    private static double VerticalOverlap(double firstY, double firstHeight, double secondY, double secondHeight)
    {
        return Math.Min(firstY + firstHeight, secondY + secondHeight) - Math.Max(firstY, secondY);
    }

    private static bool IsVerticalMergeContinuation(DocxTableCell cell)
    {
        return cell.HasVerticalMerge && !IsVerticalMergeRestart(cell);
    }

    private static DocxVerticalMergeOwner? FindVerticalMergeRestartOwner(
        DocxTable table,
        int rowIndex,
        int gridColumnIndex)
    {
        for (int previousRowIndex = rowIndex - 1; previousRowIndex >= 0; previousRowIndex--)
        {
            if (!TryGetCellAtGridColumn(table.Rows[previousRowIndex], gridColumnIndex, out DocxTableCell? previousCell) ||
                previousCell is null)
            {
                return null;
            }

            if (IsVerticalMergeRestart(previousCell))
            {
                return new DocxVerticalMergeOwner(previousCell, previousRowIndex, gridColumnIndex);
            }

            if (!IsVerticalMergeContinuation(previousCell))
            {
                return null;
            }
        }

        return null;
    }

    private static double GetVerticalMergeSpanHeight(
        DocxTable table,
        int rowIndex,
        int gridColumnIndex,
        IReadOnlyList<double> rowHeights)
    {
        double height = rowHeights[rowIndex];
        for (int nextRowIndex = rowIndex + 1; nextRowIndex < table.Rows.Count; nextRowIndex++)
        {
            if (!TryGetCellAtGridColumn(table.Rows[nextRowIndex], gridColumnIndex, out DocxTableCell? nextCell) ||
                nextCell is null ||
                !IsVerticalMergeContinuation(nextCell))
            {
                break;
            }

            height += rowHeights[nextRowIndex];
        }

        return height;
    }

    private static bool TryGetCellAtGridColumn(DocxTableRow row, int gridColumnIndex, out DocxTableCell? cell)
    {
        int currentGridColumnIndex = 0;
        foreach (DocxTableCell candidate in row.Cells)
        {
            int span = Math.Max(1, candidate.GridSpan);
            if (gridColumnIndex >= currentGridColumnIndex && gridColumnIndex < currentGridColumnIndex + span)
            {
                cell = candidate;
                return true;
            }

            currentGridColumnIndex += span;
        }

        cell = null;
        return false;
    }

    private static double[] GetTableRowCellWidths(DocxTableRow row, IReadOnlyList<double> effectiveColumns, double scale)
    {
        var widths = new double[row.Cells.Count];
        int gridColumnIndex = 0;
        for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++)
        {
            DocxTableCell cell = row.Cells[cellIndex];
            int span = Math.Max(1, cell.GridSpan);
            double width = 0d;
            for (int spanIndex = 0; spanIndex < span; spanIndex++)
            {
                width += effectiveColumns[Math.Min(gridColumnIndex + spanIndex, effectiveColumns.Count - 1)] * scale;
            }

            widths[cellIndex] = width;
            gridColumnIndex += span;
        }

        return widths;
    }

    private static IReadOnlyList<DocxBodyElement> GetTableCellLayoutBodyElements(DocxTableCell cell)
    {
        IReadOnlyList<DocxBodyElement> authoredElements = DocxTableCellContent.GetBodyElements(cell);
        if (!authoredElements.Any(IsTableCellColumnBreakElement))
        {
            return authoredElements;
        }

        var layoutElements = new List<DocxBodyElement>(authoredElements.Count);
        for (int index = 0; index < authoredElements.Count; index++)
        {
            if (authoredElements[index] is not DocxParagraphElement paragraphElement)
            {
                if (!IsTableCellColumnBreakElement(authoredElements[index]))
                {
                    layoutElements.Add(authoredElements[index]);
                }

                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            while (index + 2 < authoredElements.Count &&
                IsTableCellColumnBreakElement(authoredElements[index + 1]) &&
                authoredElements[index + 2] is DocxParagraphElement continuationElement)
            {
                paragraph = MergeTableCellColumnBreakParagraphs(paragraph, continuationElement.Paragraph);
                index += 2;
            }

            layoutElements.Add(new DocxParagraphElement(paragraph));
        }

        return layoutElements;
    }

    private static IReadOnlyList<DocxParagraph> GetParagraphsFromBodyElements(IReadOnlyList<DocxBodyElement> bodyElements)
    {
        return bodyElements
            .OfType<DocxParagraphElement>()
            .Select(element => element.Paragraph)
            .ToArray();
    }

    private static bool IsTableCellColumnBreakElement(DocxBodyElement element)
    {
        return element is DocxManualBreakElement manualBreak &&
            manualBreak.Value?.Equals("column", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static DocxParagraph MergeTableCellColumnBreakParagraphs(DocxParagraph first, DocxParagraph second)
    {
        return first with
        {
            Runs = first.Runs.Concat(second.Runs).ToArray(),
            Images = first.Images.Concat(second.Images).ToArray(),
            SpacingAfterPoints = second.SpacingAfterPoints,
            Spacing = second.Spacing
        };
    }

    private static double MeasureTableCellContentHeight(
        DocxTableCell cell,
        double cellWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints,
        double? rowTopPadding = null)
    {
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (bodyElements.Count == 0)
        {
            return 0d;
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double paddingTop = rowTopPadding ?? ResolveTableCellVerticalPadding(cell.Margins.TopPoints);
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double contentHeight = paddingTop + paddingBottom;
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                contentHeight += pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                contentHeight += MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints);
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            contentHeight += spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double fontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer);
            double lineHeight = lineHeightProfile.LineHeight;
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
            if (textSpans.Count != 0)
            {
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                double firstParagraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                double continuationParagraphWidth = Math.Max(1d, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                int lineCount = WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true).Count();
                lineHeight = QuantizeTableCellWrappedLineHeight(lineHeight, lineCount);
                contentHeight += lineCount * lineHeight;
            }
            else if (paragraph.Images.Count == 0)
            {
                contentHeight += lineHeight;
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double imageWidth = Math.Min(textWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                contentHeight += imageHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        contentHeight += pendingSpacingAfter;
        return contentHeight;
    }

    private static double MeasureNestedTableHeight(
        DocxTable table,
        double availableWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints)
    {
        DocxTableLayoutFrame frame = CreateTableLayoutFrame(
            table,
            tableIndex: -1,
            sourceBlockIndex: -1,
            x: 0d,
            availableWidth: availableWidth,
            pageContentHeight: double.MaxValue / 4d,
            textMeasurer,
            defaultTabStopPoints);
        return frame.RowHeights.Sum();
    }

    private static IReadOnlyList<DocxTextLineLayout> LayoutTableCellTextLines(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        double rowTopPadding,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints)
    {
        if (textMeasurer is null)
        {
            return [];
        }

        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        IReadOnlyList<DocxParagraph> paragraphs = GetParagraphsFromBodyElements(bodyElements);
        if (paragraphs.Count == 0)
        {
            return [];
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double paddingTop = rowTopPadding;
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints);
        double baselineInset = ResolveTableCellFirstBaselineInset(paragraphs);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - baselineInset - paddingTop;
        double cursorY = startBaselineY;
        var lines = new List<DocxTextLineLayout>();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int paragraphIndex = 0;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                cursorY -= MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints);
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            double fontSize = GetParagraphFontSize(paragraph);
            DocxLineHeightProfile lineHeightProfile = ResolveLineHeightProfile(paragraph, fontSize, textMeasurer);
            double lineHeight = lineHeightProfile.LineHeight;
            IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
            if (textSpans.Count == 0)
            {
                if (paragraph.Images.Count == 0)
                {
                    cursorY -= lineHeight;
                }
            }
            else
            {
                DocxTextRun firstRun = paragraph.Runs[0];
                double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                double continuationTextStartOffset = GetParagraphTextStartOffset(paragraph);
                double labelStartOffset = GetParagraphLabelStartOffset(paragraph);
                double paragraphX = cellX + paddingLeft + textStartOffset;
                double paragraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                double continuationParagraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                bool firstLine = true;
                DocxWrappedTextLine[] wrappedLines = WrapTextLines(textSpans, paragraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true).ToArray();
                lineHeight = QuantizeTableCellWrappedLineHeight(lineHeight, wrappedLines.Length);
                for (int lineIndex = 0; lineIndex < wrappedLines.Length; lineIndex++)
                {
                    DocxWrappedTextLine line = wrappedLines[lineIndex];
                    double lineWidth = MeasureTextSpans(line.Spans, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double drawableLineWidth = MeasureDrawableTextSpans(line.Spans, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double lineX = paragraph.Alignment switch
                    {
                        DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - lineWidth) / 2d,
                        DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - lineWidth),
                        _ => paragraphX
                    };
                    bool justifyLine = (paragraph.ListLabel is null || !firstLine) &&
                        ShouldJustifyTextLine(paragraph.Alignment, lineIndex == wrappedLines.Length - 1, drawableLineWidth, paragraphWidth, line.Spans);
                    IReadOnlyList<DocxTextSegmentLayout> segments = firstLine && paragraph.ListLabel is not null
                        ? CreateNumberedLineSegments(paragraph.ListLabel, line.Spans, firstRun, cellX + paddingLeft + labelStartOffset, lineX, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                        : justifyLine
                            ? CreateJustifiedTextSegments(line.Spans, lineX, drawableLineWidth, paragraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints)
                            : CreateTextSegments(line.Spans, lineX, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints);
                    double effectiveX = firstLine && paragraph.ListLabel is not null ? cellX + paddingLeft + labelStartOffset : lineX;
                    double effectiveWidth = firstLine && paragraph.ListLabel is not null
                        ? Math.Max(lineX + lineWidth, cellX + paddingLeft + labelStartOffset + MeasureListLabel(paragraph.ListLabel, firstRun, fontSize, textMeasurer)) - (cellX + paddingLeft + labelStartOffset)
                        : justifyLine
                            ? paragraphWidth
                            : lineWidth;
                    lines.Add(new DocxTextLineLayout(
                        firstLine && paragraph.ListLabel is not null ? paragraph.ListLabel.Text + GetListLabelTextSeparator(paragraph.ListLabel) + line.Text : line.Text,
                        firstRun,
                        fontSize,
                        effectiveX,
                        cursorY,
                        effectiveWidth,
                        segments,
                        SourceBlockIndex: null,
                        SourceParagraphIndex: paragraphIndex,
                        SourceLineIndex: lineIndex,
                        StoryKind: "TableCell",
                        LineHeight: lineHeight,
                        AppliedBeforeSpacing: firstLine ? spacingProfile.AppliedBeforeSpacing : 0d,
                        IsFirstParagraphLine: firstLine,
                        EndsWithIntraTokenBreak: line.EndsWithIntraTokenBreak,
                        SingleLineHeight: lineHeightProfile.SingleLineHeight,
                        ListLabelSingleLineHeight: lineHeightProfile.ListLabelSingleLineHeight,
                        BodyWindowsLineHeight: lineHeightProfile.BodyWindowsLineHeight,
                        ListLabelWindowsLineHeight: lineHeightProfile.ListLabelWindowsLineHeight,
                        EffectiveLineSpacingFactor: lineHeightProfile.EffectiveLineSpacingFactor,
                        LineSpacingFactorFloorApplied: lineHeightProfile.LineSpacingFactorFloorApplied,
                        PendingAfterSpacing: firstLine ? spacingProfile.PendingAfterSpacing : null,
                        ParagraphBeforeSpacing: firstLine ? spacingProfile.ParagraphBeforeSpacing : null,
                        ParagraphAfterSpacing: firstLine ? spacingProfile.ParagraphAfterSpacing : null,
                        ContextualSpacingSuppressed: firstLine ? spacingProfile.ContextualSpacingSuppressed : null,
                        SourceParagraph: paragraph));
                    firstLine = false;
                    paragraphX = cellX + paddingLeft + continuationTextStartOffset;
                    paragraphWidth = Math.Max(1d, textWidth - continuationTextStartOffset - GetParagraphRightInset(paragraph));
                    cursorY -= lineHeight;
                }
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        cursorY -= pendingSpacingAfter;
        if (lines.Count == 0)
        {
            return lines;
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - baselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = cell.VerticalAlignmentValue?.Equals("bottom", StringComparison.OrdinalIgnoreCase) == true
            ? extra
            : cell.VerticalAlignmentValue?.Equals("center", StringComparison.OrdinalIgnoreCase) == true
                ? extra / 2d
                : 0d;
        return verticalOffset == 0d ? lines : ShiftTextLines(lines, -verticalOffset);
    }

    private static IReadOnlyList<DocxInlineImageLayout> LayoutTableCellInlineImages(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        double rowTopPadding,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int pageIndex)
    {
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        IReadOnlyList<DocxParagraph> paragraphs = GetParagraphsFromBodyElements(bodyElements);
        if (paragraphs.Count == 0 || !paragraphs.Any(paragraph => paragraph.Images.Count != 0))
        {
            return [];
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double paddingTop = rowTopPadding;
        double paddingBottom = ResolveTableCellVerticalPadding(cell.Margins.BottomPoints);
        double baselineInset = ResolveTableCellFirstBaselineInset(paragraphs);
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double startBaselineY = cellY + cellHeight - baselineInset - paddingTop;
        double cursorY = startBaselineY;
        var images = new List<DocxInlineImageLayout>();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int paragraphIndex = 0;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                if (textMeasurer is not null)
                {
                    cursorY -= MeasureNestedTableHeight(tableElement.Table, textWidth, textMeasurer, defaultTabStopPoints);
                }

                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            if (textMeasurer is not null)
            {
                double fontSize = GetParagraphFontSize(paragraph);
                double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
                IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
                if (textSpans.Count != 0)
                {
                    double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
                    double firstParagraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
                    double continuationParagraphWidth = Math.Max(1d, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                    cursorY -= WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true).Count() * lineHeight;
                }
                else if (paragraph.Images.Count == 0)
                {
                    cursorY -= lineHeight;
                }
            }

            foreach (DocxInlineImage image in paragraph.Images)
            {
                double paragraphX = cellX + paddingLeft + GetParagraphStartOffset(paragraph);
                double paragraphWidth = Math.Max(1d, textWidth - GetParagraphStartOffset(paragraph) - GetParagraphRightInset(paragraph));
                double imageWidth = Math.Min(paragraphWidth, image.WidthPoints);
                double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
                double imageX = paragraph.Alignment switch
                {
                    DocxTextAlignment.Center => paragraphX + Math.Max(0, paragraphWidth - imageWidth) / 2d,
                    DocxTextAlignment.Right => paragraphX + Math.Max(0, paragraphWidth - imageWidth),
                    _ => paragraphX
                };
                images.Add(new DocxInlineImageLayout(image, imageX, cursorY - imageHeight, imageWidth, imageHeight, pageIndex, SourceParagraphIndex: paragraphIndex));
                cursorY -= imageHeight + InlineImageParagraphGapPoints;
            }

            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
            paragraphIndex++;
        }

        cursorY -= pendingSpacingAfter;
        if (images.Count == 0)
        {
            return images;
        }

        double usedHeight = Math.Max(0d, startBaselineY - cursorY);
        double availableHeight = Math.Max(0d, cellHeight - paddingTop - paddingBottom - baselineInset);
        double extra = Math.Max(0d, availableHeight - usedHeight);
        double verticalOffset = cell.VerticalAlignmentValue?.Equals("bottom", StringComparison.OrdinalIgnoreCase) == true
            ? extra
            : cell.VerticalAlignmentValue?.Equals("center", StringComparison.OrdinalIgnoreCase) == true
                ? extra / 2d
                : 0d;
        return verticalOffset == 0d
            ? images
            : images.Select(image => image with { Y = image.Y - verticalOffset }).ToArray();
    }

    private static IReadOnlyList<DocxTableRowLayout> LayoutTableCellNestedTables(
        DocxTableCell cell,
        double cellX,
        double cellY,
        double cellWidth,
        double cellHeight,
        double rowTopPadding,
        IDocxTextMeasurer? textMeasurer,
        double defaultTabStopPoints,
        int pageIndex)
    {
        IReadOnlyList<DocxBodyElement> bodyElements = GetTableCellLayoutBodyElements(cell);
        if (textMeasurer is null || !bodyElements.OfType<DocxTableElement>().Any())
        {
            return [];
        }

        double paddingLeft = ResolveTableCellHorizontalPadding(cell.Margins.LeftPoints) + ResolveTableCellBorderContentInset(cell, "left");
        double paddingRight = ResolveTableCellHorizontalPadding(cell.Margins.RightPoints) + ResolveTableCellBorderContentInset(cell, "right");
        double textWidth = Math.Max(1d, cellWidth - paddingLeft - paddingRight);
        double cursorY = cellY + cellHeight - rowTopPadding;
        var nestedRows = new List<DocxTableRowLayout>();
        double pendingSpacingAfter = 0d;
        DocxParagraph? previousParagraph = null;
        int nestedTableIndex = 0;
        foreach (DocxBodyElement bodyElement in bodyElements)
        {
            if (bodyElement is DocxTableElement tableElement)
            {
                cursorY -= pendingSpacingAfter;
                pendingSpacingAfter = 0d;
                previousParagraph = null;
                DocxTableLayoutFrame frame = CreateTableLayoutFrame(
                    tableElement.Table,
                    nestedTableIndex,
                    sourceBlockIndex: -1,
                    cellX + paddingLeft,
                    textWidth,
                    cellHeight,
                    textMeasurer,
                    defaultTabStopPoints);
                for (int rowIndex = 0; rowIndex < tableElement.Table.Rows.Count; rowIndex++)
                {
                    double rowHeight = frame.RowHeights[rowIndex];
                    nestedRows.Add(CreateTableRowLayout(
                        tableElement.Table,
                        frame.Context,
                        tableElement.Table.Rows[rowIndex],
                        rowIndex,
                        frame.RowHeights,
                        frame.EffectiveColumns,
                        frame.Scale,
                        textMeasurer,
                        defaultTabStopPoints,
                        () => pageIndex,
                        cursorY,
                        rowHeight,
                        cursorY,
                        FragmentIndex: 0,
                        FragmentCount: 1,
                        FragmentReason: "None"));
                    cursorY -= rowHeight;
                }

                nestedTableIndex++;
                continue;
            }

            if (bodyElement is not DocxParagraphElement paragraphElement)
            {
                continue;
            }

            DocxParagraph paragraph = paragraphElement.Paragraph;
            DocxParagraphSpacingProfile spacingProfile = ResolveParagraphSpacingProfile(previousParagraph, paragraph, pendingSpacingAfter);
            cursorY -= spacingProfile.AppliedBeforeSpacing;
            pendingSpacingAfter = 0d;
            cursorY -= MeasureTableCellParagraphContentHeight(paragraph, textWidth, textMeasurer, defaultTabStopPoints);
            pendingSpacingAfter = spacingProfile.ParagraphAfterSpacing;
            previousParagraph = paragraph;
        }

        return nestedRows;
    }

    private static double MeasureTableCellParagraphContentHeight(
        DocxParagraph paragraph,
        double textWidth,
        IDocxTextMeasurer textMeasurer,
        double defaultTabStopPoints)
    {
        double height = 0d;
        double fontSize = GetParagraphFontSize(paragraph);
        double lineHeight = ResolveLineHeight(paragraph, fontSize, textMeasurer);
        IReadOnlyList<DocxTextSpan> textSpans = CreateTextSpans(paragraph.Runs);
        if (textSpans.Count != 0)
        {
            double textStartOffset = GetParagraphFirstLineTextStartOffset(paragraph, fontSize, textMeasurer);
            double firstParagraphWidth = Math.Max(1d, textWidth - textStartOffset - GetParagraphRightInset(paragraph));
            double continuationParagraphWidth = Math.Max(1d, textWidth - GetParagraphTextStartOffset(paragraph) - GetParagraphRightInset(paragraph));
            height += WrapTextLines(textSpans, firstParagraphWidth, continuationParagraphWidth, fontSize, textMeasurer, paragraph.TabStops, defaultTabStopPoints, allowOverwideTokenBreaks: true).Count() * lineHeight;
        }
        else if (paragraph.Images.Count == 0)
        {
            height += lineHeight;
        }

        foreach (DocxInlineImage image in paragraph.Images)
        {
            double imageWidth = Math.Min(textWidth, image.WidthPoints);
            double imageHeight = image.HeightPoints * imageWidth / Math.Max(1d, image.WidthPoints);
            height += imageHeight + InlineImageParagraphGapPoints;
        }

        return height;
    }

    private static IReadOnlyList<DocxTextLineLayout> ShiftTextLines(IReadOnlyList<DocxTextLineLayout> lines, double deltaY)
    {
        return lines
            .Select(line => line with { BaselineY = line.BaselineY + deltaY })
            .ToArray();
    }

    private static double ResolveTableCellHorizontalPadding(double? points)
    {
        return Math.Max(0d, points ?? 0d);
    }

    private static double ResolveTableCellBorderContentInset(DocxTableCell cell, string edge)
    {
        return DocxTableBorderGeometry.ResolveVisibleWidth(DocxTableBorderGeometry.Find(cell.Borders, edge)) / 2d;
    }

    private static double ResolveTableCellVerticalPadding(double? points)
    {
        return Math.Max(0d, points ?? 0d);
    }

    private static double ResolveTableRowTopPadding(DocxTableRow row)
    {
        return row.Cells
            .Select(cell => ResolveTableCellVerticalPadding(cell.Margins.TopPoints))
            .DefaultIfEmpty(0d)
            .Max();
    }

    private static double ResolveTableCellFirstBaselineInset(IReadOnlyList<DocxParagraph> paragraphs)
    {
        return DocxLineMetrics.ResolveTableCellFirstBaselineInset(paragraphs);
    }

    private static IReadOnlyList<DocxTextSpan> CreateTextSpans(IReadOnlyList<DocxTextRun> runs)
    {
        if (runs.Count != 0 && runs.All(run => run.Text.Length == 0 || run.Hidden))
        {
            for (int i = 0; i < runs.Count; i++)
            {
                if (!runs[i].Hidden)
                {
                    return [new DocxTextSpan(" ", runs[i], i)];
                }
            }

            return [];
        }

        return runs
            .Select((run, index) => (run, index))
            .Where(item => item.run.Text.Length != 0 && !item.run.Hidden)
            .Select(item => new DocxTextSpan(item.run.Text, item.run, item.index))
            .ToArray();
    }

    private static double GetParagraphFontSize(DocxParagraph paragraph)
    {
        return paragraph.Runs.Count == 0 ? 11d : paragraph.Runs.Max(run => run.FontSize);
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        var segments = new List<DocxTextSegmentLayout>(spans.Count);
        double segmentX = lineX;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            segmentX = AddTextSegments(segments, span, segmentX, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
            if (i + 1 < spans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return segments;
    }

    private static double MeasureTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        double width = 0d;
        for (int i = 0; i < spans.Count; i++)
        {
            DocxTextSpan span = spans[i];
            width = MeasureTextSpanAdvance(span, width, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
            if (i + 1 < spans.Count)
            {
                width += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, spans[i + 1].Text);
            }
        }

        return width;
    }

    private static double MeasureDrawableTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        int length = spans.Sum(span => span.Text.Length);
        int drawableLength = FindDrawableTextLength(spans);
        return drawableLength == length
            ? MeasureTextSpans(spans, fontSize, textMeasurer, tabStops, defaultTabStopPoints)
            : MeasureTextSpans(SliceTextSpans(spans, 0, drawableLength), fontSize, textMeasurer, tabStops, defaultTabStopPoints);
    }

    private static bool ShouldJustifyTextLine(
        DocxTextAlignment alignment,
        bool isLastLine,
        double drawableLineWidth,
        double paragraphWidth,
        IReadOnlyList<DocxTextSpan> spans)
    {
        return alignment == DocxTextAlignment.Justified &&
            !isLastLine &&
            paragraphWidth - drawableLineWidth > 0.001d &&
            CountStretchableJustificationSpaces(spans) > 0 &&
            !spans.Any(span => span.Text.IndexOf('\t') >= 0);
    }

    private static IReadOnlyList<DocxTextSegmentLayout> CreateJustifiedTextSegments(
        IReadOnlyList<DocxTextSpan> spans,
        double lineX,
        double drawableLineWidth,
        double paragraphWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        int stretchableSpaces = CountStretchableJustificationSpaces(spans);
        if (stretchableSpaces == 0 || spans.Any(span => span.Text.IndexOf('\t') >= 0))
        {
            return CreateTextSegments(spans, lineX, fontSize, textMeasurer, tabStops, defaultTabStopPoints);
        }

        double extraPerSpace = Math.Max(0d, paragraphWidth - drawableLineWidth) / stretchableSpaces;
        int drawableLength = FindDrawableTextLength(spans);
        IReadOnlyList<DocxTextSpan> drawableSpans = SliceTextSpans(spans, 0, drawableLength);
        var segments = new List<DocxTextSegmentLayout>(drawableSpans.Count);
        double segmentX = lineX;
        for (int i = 0; i < drawableSpans.Count; i++)
        {
            DocxTextSpan span = drawableSpans[i];
            segmentX = AddJustifiedSpanSegments(segments, span, segmentX, fontSize, textMeasurer, extraPerSpace);
            if (i + 1 < drawableSpans.Count)
            {
                segmentX += DocxTextSpacing.BoundarySpacing(span.StyleRun, span.Text, drawableSpans[i + 1].Text);
            }
        }

        return segments;
    }

    private static double AddJustifiedSpanSegments(
        List<DocxTextSegmentLayout> segments,
        DocxTextSpan span,
        double segmentX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        double extraPerSpace)
    {
        double spanFontSize = GetTextSpanFontSize(span, fontSize);
        double baselineOffset = GetTextSpanBaselineOffset(span, fontSize);
        int start = 0;
        while (start < span.Text.Length)
        {
            bool isSpace = IsJustificationSpace(span.Text[start]);
            int end = start + 1;
            while (end < span.Text.Length && IsJustificationSpace(span.Text[end]) == isSpace)
            {
                end++;
            }

            string text = span.Text[start..end];
            double width = textMeasurer.MeasureText(span.StyleRun, text, spanFontSize);
            if (isSpace)
            {
                segmentX += width + text.Length * extraPerSpace;
            }
            else
            {
                segments.Add(new DocxTextSegmentLayout(text, span.StyleRun, segmentX, width, spanFontSize, baselineOffset, SourceTextRunIndex: span.SourceTextRunIndex));
                segmentX += width;
            }

            start = end;
        }

        return segmentX;
    }

    private static int CountStretchableJustificationSpaces(IReadOnlyList<DocxTextSpan> spans)
    {
        int drawableLength = FindDrawableTextLength(spans);
        int seen = 0;
        int count = 0;
        foreach (DocxTextSpan span in spans)
        {
            foreach (char c in span.Text)
            {
                if (seen++ >= drawableLength)
                {
                    return count;
                }

                if (IsJustificationSpace(c))
                {
                    count++;
                }
            }
        }

        return count;
    }

    private static int FindDrawableTextLength(IReadOnlyList<DocxTextSpan> spans)
    {
        int length = spans.Sum(span => span.Text.Length);
        int index = length;
        for (int spanIndex = spans.Count - 1; spanIndex >= 0; spanIndex--)
        {
            string text = spans[spanIndex].Text;
            for (int i = text.Length - 1; i >= 0; i--)
            {
                index--;
                if (!IsBreakableWhitespaceChar(text[i]))
                {
                    return index + 1;
                }
            }
        }

        return 0;
    }

    private static bool IsJustificationSpace(char c)
    {
        return c == ' ';
    }

    private static double AddTextSegments(
        List<DocxTextSegmentLayout> segments,
        DocxTextSpan span,
        double segmentX,
        double lineX,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        double spanFontSize = GetTextSpanFontSize(span, fontSize);
        double baselineOffset = GetTextSpanBaselineOffset(span, fontSize);
        int start = 0;
        for (int i = 0; i <= span.Text.Length; i++)
        {
            if (i < span.Text.Length && span.Text[i] != '\t')
            {
                continue;
            }

            if (i > start)
            {
                string text = span.Text[start..i];
                segmentX = AddTextSegment(segments, span.StyleRun, span.SourceTextRunIndex, text, segmentX, spanFontSize, baselineOffset, textMeasurer);
            }

            if (i < span.Text.Length)
            {
                segmentX = lineX + AdvanceToNextTabStop(segmentX - lineX, tabStops, defaultTabStopPoints);
                start = i + 1;
            }
        }

        return segmentX;
    }

    private static double AddTextSegment(
        List<DocxTextSegmentLayout> segments,
        DocxTextRun styleRun,
        int sourceTextRunIndex,
        string text,
        double segmentX,
        double fontSize,
        double baselineOffset,
        IDocxTextMeasurer textMeasurer)
    {
        int leadingSpaces = CountLeadingOfficeSeparatedSpaces(text);
        if (leadingSpaces == 0 || leadingSpaces == text.Length)
        {
            double width = textMeasurer.MeasureText(styleRun, text, fontSize);
            segments.Add(new DocxTextSegmentLayout(text, styleRun, segmentX, width, fontSize, baselineOffset, SourceTextRunIndex: sourceTextRunIndex));
            return segmentX + width;
        }

        string spaceText = text[..leadingSpaces];
        double spaceWidth = textMeasurer.MeasureText(styleRun, spaceText, fontSize);
        segments.Add(new DocxTextSegmentLayout(spaceText, styleRun, segmentX, spaceWidth, fontSize, baselineOffset, SourceTextRunIndex: sourceTextRunIndex));
        segmentX += spaceWidth + DocxTextSpacing.BoundarySpacing(styleRun, spaceText, text[leadingSpaces..]);

        string bodyText = text[leadingSpaces..];
        double bodyWidth = textMeasurer.MeasureText(styleRun, bodyText, fontSize);
        segments.Add(new DocxTextSegmentLayout(bodyText, styleRun, segmentX, bodyWidth, fontSize, baselineOffset, SourceTextRunIndex: sourceTextRunIndex));
        return segmentX + bodyWidth;
    }

    private static int CountLeadingOfficeSeparatedSpaces(string text)
    {
        int count = 0;
        while (count < text.Length && text[count] == ' ')
        {
            count++;
        }

        return count;
    }

    private static double MeasureTextSpanAdvance(
        DocxTextSpan span,
        double currentWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints)
    {
        double spanFontSize = GetTextSpanFontSize(span, fontSize);
        int start = 0;
        for (int i = 0; i <= span.Text.Length; i++)
        {
            if (i < span.Text.Length && span.Text[i] != '\t')
            {
                continue;
            }

            if (i > start)
            {
                currentWidth += textMeasurer.MeasureText(span.StyleRun, span.Text[start..i], spanFontSize);
            }

            if (i < span.Text.Length)
            {
                currentWidth = AdvanceToNextTabStop(currentWidth, tabStops, defaultTabStopPoints);
                start = i + 1;
            }
        }

        return currentWidth;
    }

    private static double GetTextSpanFontSize(DocxTextSpan span, double fallbackFontSize)
    {
        double nominalFontSize = span.StyleRun.FontSize > 0d ? span.StyleRun.FontSize : fallbackFontSize;
        return DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
    }

    private static double GetTextSpanBaselineOffset(DocxTextSpan span, double fallbackFontSize)
    {
        double nominalFontSize = span.StyleRun.FontSize > 0d ? span.StyleRun.FontSize : fallbackFontSize;
        double layoutFontSize = DocxVerticalAlignMetrics.ResolveFontSize(nominalFontSize, span.StyleRun);
        return DocxVerticalAlignMetrics.ResolveBaselineOffset(nominalFontSize, layoutFontSize, span.StyleRun);
    }

    private static double AdvanceToNextDefaultTabStop(double width, double defaultTabStopPoints)
    {
        double tabStop = defaultTabStopPoints > 0d ? defaultTabStopPoints : WordDefaultTabStopPoints;
        return (Math.Floor(width / tabStop) + 1d) * tabStop;
    }

    private static double AdvanceToNextTabStop(double width, IReadOnlyList<DocxTabStop> tabStops, double defaultTabStopPoints)
    {
        foreach (DocxTabStop tabStop in tabStops
            .Where(tabStop => tabStop.PositionPoints is not null && IsPositioningTabStop(tabStop))
            .OrderBy(tabStop => tabStop.PositionPoints!.Value))
        {
            if (tabStop.PositionPoints!.Value > width + 0.001d)
            {
                return tabStop.PositionPoints.Value;
            }
        }

        return AdvanceToNextDefaultTabStop(width, defaultTabStopPoints);
    }

    private static bool IsPositioningTabStop(DocxTabStop tabStop)
    {
        return !string.Equals(tabStop.Value, "bar", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(tabStop.Value, "clear", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<DocxWrappedTextLine> WrapTextLines(
        IReadOnlyList<DocxTextSpan> spans,
        double firstLineMaxWidth,
        double continuationLineMaxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        bool allowOverwideTokenBreaks)
    {
        string text = string.Concat(spans.Select(span => span.Text));
        int lineIndex = 0;
        int segmentStart = 0;
        while (segmentStart <= text.Length)
        {
            int breakIndex = text.IndexOf('\n', segmentStart);
            int segmentLength = breakIndex < 0 ? text.Length - segmentStart : breakIndex - segmentStart;
            bool yielded = false;
            foreach (DocxWrappedTextLine line in WrapWords(text, spans, segmentStart, segmentLength, index => index == 0 && lineIndex == 0 ? firstLineMaxWidth : continuationLineMaxWidth, fontSize, textMeasurer, tabStops, defaultTabStopPoints, allowOverwideTokenBreaks))
            {
                yielded = true;
                yield return line;
                lineIndex++;
            }

            if (!yielded && segmentLength == 0)
            {
                yield return new DocxWrappedTextLine(string.Empty, []);
                lineIndex++;
            }

            if (breakIndex < 0)
            {
                yield break;
            }

            segmentStart = breakIndex + 1;
        }
    }

    private static IEnumerable<DocxWrappedTextLine> WrapWords(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int segmentStart,
        int segmentLength,
        Func<int, double> maxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        bool allowOverwideTokenBreaks)
    {
        IReadOnlyList<TextToken> tokens = TokenizeSpaces(text, segmentStart, segmentLength);
        if (tokens.Count == 0)
        {
            yield break;
        }

        int lineStart = tokens[0].Start;
        int lineLength = 0;
        int lineIndex = 0;
        for (int tokenIndex = 0; tokenIndex < tokens.Count; tokenIndex++)
        {
            TextToken token = tokens[tokenIndex];
            int candidateLength = token.Start + token.Length - lineStart;
            bool lineHasNonWhitespace = HasNonWhitespace(text, lineStart, lineLength);
            if (lineLength > 0 &&
                lineHasNonWhitespace &&
                !token.IsBreakableWhitespace &&
                MeasureTextSpans(SliceTextSpans(spans, lineStart, candidateLength), fontSize, textMeasurer, tabStops, defaultTabStopPoints) > maxWidth(lineIndex))
            {
                yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
                lineIndex++;
                lineStart = token.Start;
                lineLength = 0;
                tokenIndex--;
                continue;
            }

            if (lineLength == 0 &&
                allowOverwideTokenBreaks &&
                !token.IsBreakableWhitespace &&
                TryFindOverwideTokenBreak(text, spans, token, maxWidth(lineIndex), fontSize, textMeasurer, tabStops, defaultTabStopPoints, out int breakLength))
            {
                yield return CreateWrappedTextLine(text, spans, token.Start, breakLength, endsWithIntraTokenBreak: true);
                lineIndex++;
                lineStart = token.Start + breakLength;
                lineLength = 0;
                tokens = ReplaceToken(tokens, tokenIndex, new TextToken(text.Substring(lineStart, token.Length - breakLength), lineStart, token.Length - breakLength));
                tokenIndex--;
            }
            else
            {
                lineLength = candidateLength;
            }
        }

        if (lineLength > 0)
        {
            yield return CreateWrappedTextLine(text, spans, lineStart, lineLength);
        }
    }

    private static bool TryFindOverwideTokenBreak(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        TextToken token,
        double maxWidth,
        double fontSize,
        IDocxTextMeasurer textMeasurer,
        IReadOnlyList<DocxTabStop> tabStops,
        double defaultTabStopPoints,
        out int breakLength)
    {
        breakLength = 0;
        double tokenWidth = MeasureTextSpans(SliceTextSpans(spans, token.Start, token.Length), fontSize, textMeasurer, tabStops, defaultTabStopPoints);
        if (tokenWidth <= maxWidth)
        {
            return false;
        }

        for (int length = token.Length - 1; length > 0; length--)
        {
            int absoluteIndex = token.Start + length - 1;
            if (!DocxLineBreakOpportunities.IsOpportunityAfter(text[absoluteIndex]))
            {
                continue;
            }

            double prefixWidth = MeasureTextSpans(SliceTextSpans(spans, token.Start, length), fontSize, textMeasurer, tabStops, defaultTabStopPoints);
            if (prefixWidth <= maxWidth)
            {
                breakLength = length;
                return true;
            }
        }

        for (int length = token.Length - 1; length > 0; length--)
        {
            if (!IsSafeEmergencyTokenBreak(text, token, length))
            {
                continue;
            }

            double prefixWidth = MeasureTextSpans(SliceTextSpans(spans, token.Start, length), fontSize, textMeasurer, tabStops, defaultTabStopPoints);
            if (prefixWidth <= maxWidth)
            {
                breakLength = length;
                return true;
            }
        }

        return false;
    }

    private static bool IsSafeEmergencyTokenBreak(string text, TextToken token, int length)
    {
        if (length <= 0 || length >= token.Length)
        {
            return false;
        }

        int breakIndex = token.Start + length;
        char before = text[breakIndex - 1];
        char after = text[breakIndex];
        if (char.IsHighSurrogate(before) && char.IsLowSurrogate(after))
        {
            return false;
        }

        UnicodeCategory afterCategory = char.GetUnicodeCategory(after);
        return afterCategory is not UnicodeCategory.NonSpacingMark and
            not UnicodeCategory.SpacingCombiningMark and
            not UnicodeCategory.EnclosingMark;
    }

    private static IReadOnlyList<TextToken> ReplaceToken(IReadOnlyList<TextToken> tokens, int index, TextToken replacement)
    {
        var result = new List<TextToken>(tokens.Count);
        for (int i = 0; i < tokens.Count; i++)
        {
            result.Add(i == index ? replacement : tokens[i]);
        }

        return result;
    }

    private static DocxWrappedTextLine CreateWrappedTextLine(
        string text,
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length,
        bool endsWithIntraTokenBreak = false)
    {
        return new DocxWrappedTextLine(text.Substring(start, length), SliceTextSpans(spans, start, length), endsWithIntraTokenBreak);
    }

    private static IReadOnlyList<DocxTextSpan> SliceTextSpans(
        IReadOnlyList<DocxTextSpan> spans,
        int start,
        int length)
    {
        if (length == 0)
        {
            return [];
        }

        var sliced = new List<DocxTextSpan>();
        int spanStart = 0;
        int end = start + length;
        foreach (DocxTextSpan span in spans)
        {
            int spanEnd = spanStart + span.Text.Length;
            int sliceStart = Math.Max(start, spanStart);
            int sliceEnd = Math.Min(end, spanEnd);
            if (sliceStart < sliceEnd)
            {
                sliced.Add(new DocxTextSpan(span.Text[(sliceStart - spanStart)..(sliceEnd - spanStart)], span.StyleRun, span.SourceTextRunIndex));
            }

            if (spanEnd >= end)
            {
                break;
            }

            spanStart = spanEnd;
        }

        return sliced;
    }

    private static IReadOnlyList<TextToken> TokenizeSpaces(string text, int start, int length)
    {
        if (length == 0)
        {
            return [];
        }

        string segment = text.Substring(start, length);
        return TokenizeSpaces(segment)
            .Select(token => new TextToken(token.Text, start + token.Start, token.Length))
            .ToArray();
    }

    private static bool HasNonWhitespace(string text, int start, int length)
    {
        for (int i = start; i < start + length; i++)
        {
            if (!char.IsWhiteSpace(text[i]))
            {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<TextToken> TokenizeSpaces(string text)
    {
        if (text.Length == 0)
        {
            return [];
        }

        var tokens = new List<TextToken>();
        int start = 0;
        bool inBreakableWhitespace = IsBreakableWhitespaceChar(text[0]);
        for (int i = 1; i < text.Length; i++)
        {
            bool breakableWhitespace = IsBreakableWhitespaceChar(text[i]);
            if (breakableWhitespace == inBreakableWhitespace)
            {
                continue;
            }

            tokens.Add(new TextToken(text[start..i], start, i - start));
            start = i;
            inBreakableWhitespace = breakableWhitespace;
        }

        tokens.Add(new TextToken(text[start..], start, text.Length - start));
        return tokens;
    }

    private static bool IsBreakableWhitespaceChar(char value)
    {
        return char.IsWhiteSpace(value) &&
            value != '\u00A0' &&
            value != '\u202F' &&
            value != '\u2007';
    }

    private static class DocxLineBreakOpportunities
    {
        public static bool IsOpportunityAfter(char value)
        {
            return value is '-' or '/' or '\\' or '\u2010' or '\u2011' or '\u2012' or '\u2013' or '\u2014';
        }
    }

    private readonly record struct TextToken(string Text, int Start, int Length)
    {
        public bool IsBreakableWhitespace => Text.All(IsBreakableWhitespaceChar);
    }
}
