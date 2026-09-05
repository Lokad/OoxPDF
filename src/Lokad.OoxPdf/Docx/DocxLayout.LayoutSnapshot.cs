using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxLayoutSnapshot(
    string MarkupMode,
    string MarkupGeometryMode,
    double MarkupMarginReservePoints,
    int RevisionItemCount,
    int RevisionCount,
    int InsertionRevisionCount,
    int DeletionRevisionCount,
    int MoveFromRevisionCount,
    int MoveToRevisionCount,
    int OtherRevisionCount,
    int CommentReferenceItemCount,
    int CommentReferenceCount,
    IReadOnlyList<DocxLayoutPageSnapshot> Pages,
    IReadOnlyList<DocxTableSnapshot> Tables,
    IReadOnlyList<DocxLayoutSourceBlockSnapshot> SourceBlocks,
    IReadOnlyList<DocxFloatingDrawingLayoutSnapshot> FloatingDrawings,
    IReadOnlyList<DocxFloatingDrawingLayoutSnapshot> StaticFloatingDrawings,
    IReadOnlyList<DocxRelatedStoryLayoutSnapshot> RelatedStories)
{
    public static DocxLayoutSnapshot FromLayout(DocxLayout layout)
    {
        return FromLayout(layout, OoxPdfDocxMarkupMode.Final, OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout);
    }

    public static DocxLayoutSnapshot FromLayout(
        DocxLayout layout,
        OoxPdfDocxMarkupMode markupMode,
        OoxPdfDocxMarkupGeometryMode markupGeometryMode)
    {
        DocxLayoutPageSnapshot[] pages = layout.Pages.Select(ToSnapshot).ToArray();
        IReadOnlyList<DocxLayoutSourceBlockSnapshot> sourceBlocks = ToSourceBlockSnapshots(pages);
        DocxLayoutItemSnapshot[] items = pages.SelectMany(EnumerateSnapshotItems).ToArray();
        return new DocxLayoutSnapshot(
            markupMode.ToString(),
            markupGeometryMode.ToString(),
            pages.Select(page => page.MarkupMarginReservePoints).DefaultIfEmpty(0d).Max(),
            items.Count(item => item.RevisionCount != 0),
            items.Sum(item => item.RevisionCount),
            items.Sum(item => item.InsertionRevisionCount),
            items.Sum(item => item.DeletionRevisionCount),
            items.Sum(item => item.MoveFromRevisionCount),
            items.Sum(item => item.MoveToRevisionCount),
            items.Sum(item => item.OtherRevisionCount),
            items.Count(item => item.CommentReferenceCount != 0),
            items.Sum(item => item.CommentReferenceCount),
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
                    story.Story.Kind.ToValueString(),
                    story.Story.PartName,
                    story.Story.Id,
                    story.Story.Type,
                    story.Story.BodyElements.Count,
                    story.Story.BodyElements.OfType<DocxParagraphElement>().Count(),
                    DocxBlockTraversal.EnumerateBodyTables(story.Story).Count(),
                    story.TextLines.Count,
                    tableCellTextLineCount,
                    story.TableRows.Count,
                    inlineImageCount,
                    story.FloatingDrawings.Count,
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
            .GroupBy(item => item.SourceBlockIndex ?? int.MaxValue)
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
                    drawing.Drawing.WrapKind?.ToValueString(),
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
                    drawing.TextBoxLayout?.TextLines.Count ?? 0,
                    drawing.TextBoxLayout?.InlineImages.Count ?? 0,
                    drawing.TextBoxLayout?.TableRows.Count ?? 0,
                    drawing.TextBoxLayout?.ContentHeight ?? 0d,
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
            .GroupBy(entry => entry.item.SourceBlockIndex ?? int.MaxValue)
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

    private static bool IsNormalRelatedStory(DocxRelatedStory story)
    {
        return string.IsNullOrEmpty(story.Type) ||
            string.Equals(story.Type, "normal", StringComparison.OrdinalIgnoreCase);
    }

    private static DocxLayoutPageSnapshot ToSnapshot(DocxLayoutPage page)
    {
        IReadOnlyList<DocxLayoutColumnFrameSnapshot> columnFrames = page.ColumnFrames.Select(frame => new DocxLayoutColumnFrameSnapshot(
            frame.Index,
            frame.X,
            frame.Width,
            frame.GutterAfterPoints)).ToArray();
        IReadOnlyList<DocxLayoutItemSnapshot> items = page.Items.Select(item => ToSnapshot(item, page.ColumnFrames)).ToArray();
        IReadOnlyList<DocxLayoutItemSnapshot> placedRelatedItems = page.PlacedRelatedStories
            .SelectMany(story => story.TextLines
                .Cast<DocxLayoutItem>()
                .Concat(story.InlineImages)
                .Concat(story.TableRows)
                .Select(ToPlacedRelatedSnapshot))
            .ToArray();
        IReadOnlyList<DocxPlacedRelatedStoryLayoutSnapshot> placedRelatedStories = page.PlacedRelatedStories
            .Select(story => new DocxPlacedRelatedStoryLayoutSnapshot(
                story.StoryLayout.Story.Kind.ToValueString(),
                story.StoryLayout.Story.PartName,
                story.StoryLayout.Story.Id,
                story.StoryLayout.Story.Type,
                story.SourceBlockIndex,
                story.X,
                story.TopY,
                story.Width,
                story.Height,
                story.ContentTopOffset,
                story.ContentHeight,
                story.SeparatorY,
                story.SeparatorWidth,
                story.SeparatorThickness,
                story.TextLines.Count,
                story.InlineImages.Count,
                story.FloatingDrawings.Count,
                story.TableRows.Count))
            .ToArray();
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
            page.MarkupMarginReservePoints,
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
            page.PageSettings.FootnoteReferenceSettings.PositionValue,
            page.PageSettings.FootnoteReferenceSettings.NumberRestartValue,
            page.PageSettings.EndnoteReferenceSettings.PositionValue,
            page.PageSettings.EndnoteReferenceSettings.NumberRestartValue,
            page.ColumnFrames.Count,
            page.ColumnFrames.Sum(frame => frame.Width),
            page.ColumnFrames.Sum(frame => frame.GutterAfterPoints ?? 0d),
            columnFrames,
            items.Count,
            page.StaticTextLines.Count,
            page.StaticInlineImages.Count,
            page.StaticTableRows.Count,
            page.PlacedRelatedStories.Count,
            page.PlacedRelatedStories.Count(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Footnote && IsNormalRelatedStory(story.StoryLayout.Story)),
            page.PlacedRelatedStories.Count(story => story.StoryLayout.Story.Kind == DocxRelatedStoryKind.Endnote && IsNormalRelatedStory(story.StoryLayout.Story)),
            page.PlacedRelatedStories.Sum(story => story.TextLines.Count),
            page.PlacedRelatedStories.Sum(story => story.InlineImages.Count),
            page.PlacedRelatedStories.Sum(story => story.TableRows.Count),
            items.Count(item => item.Kind == "TextLine"),
            items.Count(item => item.Kind == "InlineImage"),
            items.Count(item => item.Kind == "TableRow"),
            items.Count(item => item.RevisionCount != 0),
            items.Sum(item => item.RevisionCount),
            items.Sum(item => item.InsertionRevisionCount),
            items.Sum(item => item.DeletionRevisionCount),
            items.Sum(item => item.MoveFromRevisionCount),
            items.Sum(item => item.MoveToRevisionCount),
            items.Sum(item => item.OtherRevisionCount),
            items.Count(item => item.CommentReferenceCount != 0),
            items.Sum(item => item.CommentReferenceCount),
            sourceBlockIndexes.Distinct().Count(),
            sourceBlockIndexes.FirstOrDefault(),
            sourceBlockIndexes.LastOrDefault(),
            Math.Max(0d, verticalTop - verticalBottom),
            items.Where(item => item.Kind == "TextLine").Sum(item => item.Height),
            items.Where(item => item.Kind == "InlineImage").Sum(item => item.Height),
            items.Where(item => item.Kind == "TableRow").Sum(item => item.Height),
            ToStaticStorySnapshots(staticItems),
            staticItems,
            placedRelatedStories,
            placedRelatedItems,
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
                    .OfType<int>()
                    .Distinct()
                    .OrderBy(index => index)
                    .ToArray();
                int[] lineIndexes = storyItems
                    .Select(item => item.SourceLineIndex)
                    .Where(index => index is not null)
                    .OfType<int>()
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

    private static DocxLayoutItemSnapshot ToPlacedRelatedSnapshot(DocxLayoutItem item)
    {
        DocxLayoutItemSnapshot snapshot = ToSnapshot(item, []);
        return snapshot with { Kind = "Placed" + snapshot.Kind };
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
                TextLines: null,
                ParagraphStyleId: text.SourceParagraph?.EffectiveProperties.StyleResolution.StyleId,
                ParagraphStyleFound: text.SourceParagraph?.EffectiveProperties.StyleResolution.StyleFound,
                ParagraphStyleDepth: text.SourceParagraph?.EffectiveProperties.StyleResolution.StyleDepth,
                HasDocumentDefaultParagraphProperties: text.SourceParagraph?.EffectiveProperties.StyleResolution.HasDocumentDefaultParagraphProperties,
                HasDirectParagraphProperties: text.SourceParagraph?.EffectiveProperties.StyleResolution.HasDirectParagraphProperties,
                HasTableStyleParagraphProperties: text.SourceParagraph?.EffectiveProperties.StyleResolution.HasTableStyleParagraphProperties,
                CharacterStyleTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasCharacterStyleRunProperties),
                DirectRunPropertyTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasDirectRunProperties),
                ParagraphStyleRunPropertyTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasParagraphStyleRunProperties),
                TableStyleRunPropertyTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasTableStyleRunProperties),
                DocumentDefaultRunPropertyTextSegmentCount: CountTextSegments(text, static resolution => resolution.HasDocumentDefaultRunProperties),
                LineHeightSource: text.LineHeightSource?.ToString(),
                RevisionCount: CountRevisions(text.SourceParagraph),
                InsertionRevisionCount: CountRevisions(text.SourceParagraph, DocxRevisionKind.Insertion),
                DeletionRevisionCount: CountRevisions(text.SourceParagraph, DocxRevisionKind.Deletion),
                MoveFromRevisionCount: CountRevisions(text.SourceParagraph, DocxRevisionKind.MoveFrom),
                MoveToRevisionCount: CountRevisions(text.SourceParagraph, DocxRevisionKind.MoveTo),
                OtherRevisionCount: CountOtherRevisions(text.SourceParagraph),
                CommentReferenceCount: CountCommentReferences(text.SourceParagraph)),
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
                image.StoryVariantType, TextLines: null, ParagraphStyleId: null, ParagraphStyleFound: null, ParagraphStyleDepth: null, HasDocumentDefaultParagraphProperties: null, HasDirectParagraphProperties: null, HasTableStyleParagraphProperties: null, CharacterStyleTextSegmentCount: null, DirectRunPropertyTextSegmentCount: null, ParagraphStyleRunPropertyTextSegmentCount: null, TableStyleRunPropertyTextSegmentCount: null, DocumentDefaultRunPropertyTextSegmentCount: null, LineHeightSource: null, RevisionCount: 0, InsertionRevisionCount: 0, DeletionRevisionCount: 0, MoveFromRevisionCount: 0, MoveToRevisionCount: 0, OtherRevisionCount: 0, CommentReferenceCount: 0),
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
                ToTableRowTextLineSnapshots(row), ParagraphStyleId: null, ParagraphStyleFound: null, ParagraphStyleDepth: null, HasDocumentDefaultParagraphProperties: null, HasDirectParagraphProperties: null, HasTableStyleParagraphProperties: null, CharacterStyleTextSegmentCount: null, DirectRunPropertyTextSegmentCount: null, ParagraphStyleRunPropertyTextSegmentCount: null, TableStyleRunPropertyTextSegmentCount: null, DocumentDefaultRunPropertyTextSegmentCount: null, LineHeightSource: null,
                RevisionCount: row.RevisionCount + TableRowParagraphs(row).Sum(CountRevisions),
                InsertionRevisionCount: TableRowParagraphs(row).Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.Insertion)),
                DeletionRevisionCount: TableRowParagraphs(row).Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.Deletion)),
                MoveFromRevisionCount: TableRowParagraphs(row).Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.MoveFrom)),
                MoveToRevisionCount: TableRowParagraphs(row).Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.MoveTo)),
                OtherRevisionCount: TableRowParagraphs(row).Sum(CountOtherRevisions),
                CommentReferenceCount: TableRowParagraphs(row).Sum(CountCommentReferences)),
            _ => new DocxLayoutItemSnapshot("Unknown", 0d, 0d, 0d, 0d, 0, 0, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, 0, 0, 0, 0, 0, 0, 0)
        };

        int CountCommentReferences(DocxParagraph? paragraph)
        {
            return paragraph?.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment) ?? 0;
        }

        int CountTextSegments(DocxTextLineLayout text, Func<DocxRunStyleResolution, bool> predicate)
        {
            return text.Segments.Count(segment => predicate(segment.StyleRun.StyleResolution));
        }
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

    private static IEnumerable<DocxLayoutItemSnapshot> EnumerateSnapshotItems(DocxLayoutPageSnapshot page)
    {
        foreach (DocxLayoutItemSnapshot item in page.StaticItems)
        {
            yield return item;
        }

        foreach (DocxLayoutItemSnapshot item in page.PlacedRelatedItems)
        {
            yield return item;
        }

        foreach (DocxLayoutItemSnapshot item in page.Items)
        {
            yield return item;
        }
    }

    private static IReadOnlyList<DocxParagraph> TableRowParagraphs(DocxTableRowLayout row)
    {
        var paragraphs = new List<DocxParagraph>();
        foreach (DocxTableCellLayout cell in row.Cells)
        {
            AddParagraphs(cell.TextLines);
            foreach (DocxTableRowLayout nestedRow in cell.NestedRows)
            {
                foreach (DocxParagraph paragraph in TableRowParagraphs(nestedRow))
                {
                    AddParagraph(paragraph);
                }
            }
        }

        return paragraphs;

        void AddParagraphs(IEnumerable<DocxTextLineLayout> lines)
        {
            foreach (DocxTextLineLayout line in lines)
            {
                if (line.SourceParagraph is not null)
                {
                    AddParagraph(line.SourceParagraph);
                }
            }
        }

        void AddParagraph(DocxParagraph paragraph)
        {
            if (!paragraphs.Any(existing => ReferenceEquals(existing, paragraph)))
            {
                paragraphs.Add(paragraph);
            }
        }
    }

    private static int CountRevisions(DocxParagraph? paragraph)
    {
        return paragraph?.Revisions.Count ?? 0;
    }

    private static int CountRevisions(DocxParagraph? paragraph, DocxRevisionKind kind)
    {
        return paragraph?.Revisions.Count(revision => revision.Kind == kind) ?? 0;
    }

    private static int CountOtherRevisions(DocxParagraph? paragraph)
    {
        return paragraph?.Revisions.Count(revision =>
            revision.Kind is not (DocxRevisionKind.Insertion or DocxRevisionKind.Deletion or DocxRevisionKind.MoveFrom or DocxRevisionKind.MoveTo)) ?? 0;
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
        IReadOnlyList<double> spacingBeforePoints = paragraphs.Select(paragraph => paragraph.EffectiveProperties.SpacingBeforePoints).ToArray();
        IReadOnlyList<double> spacingAfterPoints = paragraphs.Select(paragraph => paragraph.EffectiveProperties.SpacingAfterPoints).ToArray();
        string cellText = string.Concat(paragraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.Text));
        string visualCellText = string.Concat(visualParagraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.Text));
        TextProfile textProfile = BuildTextProfile();
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
            paragraphs.Count(paragraph => HasBeforeSpacingToken(paragraph.EffectiveProperties.Spacing)),
            paragraphs.Count(paragraph => HasAfterSpacingToken(paragraph.EffectiveProperties.Spacing)),
            paragraphs.Count(paragraph => string.Equals(paragraph.EffectiveProperties.Spacing.BeforeValue, "0", StringComparison.Ordinal)),
            paragraphs.Count(paragraph => string.Equals(paragraph.EffectiveProperties.Spacing.AfterValue, "0", StringComparison.Ordinal)),
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
            cell.NoWrap,
            cell.NoWrapValue,
            cell.FitText,
            cell.FitTextValue,
            cell.TextDirectionValue,
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

        TextProfile BuildTextProfile()
        {
            int spaceCharacterCount = 0;
            int nonAsciiCharacterCount = 0;
            int punctuationCharacterCount = 0;
            int digitCharacterCount = 0;
            int uppercaseCharacterCount = 0;
            int lowercaseCharacterCount = 0;
            foreach (char value in cellText)
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

            int longestBreakableTokenLength = ResolveLongestBreakableTokenLength(cellText);
            return new TextProfile(
                spaceCharacterCount,
                nonAsciiCharacterCount,
                punctuationCharacterCount,
                digitCharacterCount,
                uppercaseCharacterCount,
                lowercaseCharacterCount,
                longestBreakableTokenLength);
        }
    }

    private static IReadOnlyList<DocxParagraph> GetParagraphsFromBodyElements(IReadOnlyList<DocxBodyElement> bodyElements)
    {
        return DocxBlockTraversal.EnumerateDirectParagraphs(bodyElements).ToArray();
    }

    private static int ResolveLongestBreakableTokenLength(string text)
    {
        int longest = 0;
        int current = 0;
        foreach (char value in text)
        {
            if (IsTextProfileBreakableWhitespaceChar(value))
            {
                longest = Math.Max(longest, current);
                current = 0;
                continue;
            }

            if (IsTextProfileHiddenBreakCharacter(value))
            {
                longest = Math.Max(longest, current);
                current = 0;
                continue;
            }

            current++;
            if (IsTextProfileLineBreakOpportunityAfter(value))
            {
                longest = Math.Max(longest, current);
                current = 0;
            }
        }

        return Math.Max(longest, current);
    }

    private static bool IsTextProfileBreakableWhitespaceChar(char value)
    {
        return char.IsWhiteSpace(value) &&
            !IsTextProfileNoBreakWhitespaceChar(value);
    }

    private static bool IsTextProfileNoBreakWhitespaceChar(char value)
    {
        return value is '\u00A0' or '\u202F' or '\u2007';
    }

    private static bool IsTextProfileHiddenBreakCharacter(char value)
    {
        return value is '\u00AD' or '\u200B';
    }

    private static bool IsTextProfileLineBreakOpportunityAfter(char value)
    {
        return value is '-' or '/' or '\\' or '\u2010' or '\u2012' or '\u2013' or '\u2014';
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
