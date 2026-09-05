using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal sealed partial record DocxStructureSnapshot(
    string MarkupMode,
    int BlockCount,
    int ParagraphBlockCount,
    int TableBlockCount,
    int PageBreakBlockCount,
    int ManualBreakBlockCount,
    int SectionBreakBlockCount,
    int ContinuousSectionBreakBlockCount,
    int PageStartingSectionBreakBlockCount,
    int DefaultSectionBreakBlockCount,
    int ColumnSectionBreakBlockCount,
    int BodyTextLength,
    int RevisionCount,
    int InsertionRevisionCount,
    int DeletionRevisionCount,
    int MoveFromRevisionCount,
    int MoveToRevisionCount,
    int OtherRevisionCount,
    int InlineImageCount,
    int InlineReferenceCount,
    int CommentReferenceCount,
    int AnchoredInlineReferenceCount,
    int ResolvedInlineReferenceCount,
    int MaxInlineReferenceTextOffsetInRun,
    int FieldReferenceCount,
    int PageFieldReferenceCount,
    int NumPagesFieldReferenceCount,
    int OtherFieldReferenceCount,
    int ComplexFieldReferenceCount,
    int CachedResultFieldReferenceCount,
    int RenderedCachedResultFieldReferenceCount,
    int PlaceholderFieldReferenceCount,
    int NestedFieldReferenceCount,
    int BookmarkAnchorCount,
    int HyperlinkCount,
    int ExternalHyperlinkCount,
    int InternalHyperlinkCount,
    int FloatingDrawingCount,
    IReadOnlyList<DocxStructureBlockSnapshot> Blocks,
    IReadOnlyList<DocxStructureStorySnapshot> Stories,
    IReadOnlyList<DocxStructureInlineReferenceSnapshot> InlineReferences,
    IReadOnlyList<DocxStructureFloatingDrawingSnapshot> FloatingDrawings,
    DocxStyleCatalog StyleCatalog,
    IReadOnlyList<DocxStructureStyleUsageSnapshot> StyleUsages,
    IReadOnlyList<DocxStructureListUsageSnapshot> ListUsages,
    IReadOnlyList<DocxStructureTableSnapshot> Tables,
    IReadOnlyList<DocxStructureTableAdjacencySnapshot> TableAdjacency,
    IReadOnlyList<DocxStructureCommentRangeSnapshot>? CommentRanges,
    IReadOnlyList<DocxStructureRevisionRangeSnapshot>? RevisionRanges,
    int FormattingRevisionCount,
    int RunFormattingRevisionCount,
    int ParagraphFormattingRevisionCount,
    int TableFormattingRevisionCount,
    int RowFormattingRevisionCount,
    int CellFormattingRevisionCount,
    int SectionFormattingRevisionCount,
    IReadOnlyList<DocxStructureFormattingRevisionPropertySnapshot>? FormattingRevisionProperties,
    int PackageCommentAnchorIdCount,
    int HiddenCommentAnchorIdCount,
    int ResolvedCommentStoryAnchorCount,
    int HiddenCommentStoryAnchorCount,
    int OrphanedCommentStoryAnchorCount,
    int UnsupportedCommentStoryAnchorCount,
    IReadOnlyList<DocxStructureCommentStoryAnchorSnapshot>? CommentStoryAnchors,
    int DynamicFieldReferenceCount,
    int DynamicPlaceholderFieldReferenceCount,
    int DynamicComplexWithoutCachedResultFieldReferenceCount,
    int DynamicCachedResultNotRenderedFieldReferenceCount)
{
    public static DocxStructureSnapshot FromDocument(DocxDocument document)
    {
            bool IsContinuousSectionBreak(string? typeValue)
            {
                return string.Equals(typeValue, "continuous", StringComparison.OrdinalIgnoreCase);
            }

            bool StartsNewPageSectionBreak(string? typeValue)
            {
                return typeValue is null
                    || string.Equals(typeValue, "nextPage", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(typeValue, "oddPage", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(typeValue, "evenPage", StringComparison.OrdinalIgnoreCase);
            }

            bool HasSectionColumns(DocxStructureBlockSnapshot block)
            {
                return block.SectionColumnCountValue is not null
                    || block.SectionColumnEqualWidthValue is not null
                    || block.SectionColumnSpaceValue is not null
                    || (block.SectionColumnDefinitionCount ?? 0) > 0;
            }

            string GetBlockKind(DocxBodyElement element)
            {
                return element switch
                {
                    DocxParagraphElement => "Paragraph",
                    DocxTableElement => "Table",
                    DocxImplicitParagraphElement => "ImplicitParagraph",
                    DocxPageBreakElement => "PageBreak",
                    DocxManualBreakElement => "ManualBreak",
                    DocxSectionBreakElement => "SectionBreak",
                    _ => "Unknown"
                };
            }

        var blocks = new List<DocxStructureBlockSnapshot>(document.BodyElements.Count);
        var tables = new List<DocxStructureTableSnapshot>();
        var adjacency = new List<DocxStructureTableAdjacencySnapshot>();
        int tableIndex = 0;
        for (int blockIndex = 0; blockIndex < document.BodyElements.Count; blockIndex++)
        {
            DocxBodyElement element = document.BodyElements[blockIndex];
            string? previousKind = blockIndex == 0 ? null : GetBlockKind(document.BodyElements[blockIndex - 1]);
            string? nextKind = blockIndex + 1 >= document.BodyElements.Count ? null : GetBlockKind(document.BodyElements[blockIndex + 1]);
            switch (element)
            {
                case DocxParagraphElement paragraph:
                    blocks.Add(FromParagraph(blockIndex, previousKind, nextKind, paragraph.Paragraph, document.RelatedStories));
                    break;
                case DocxTableElement table:
                    blocks.Add(FromTable(blockIndex, previousKind, nextKind, table.Table, tableIndex, document.RelatedStories));
                    tables.Add(ToTableSnapshot(table.Table, tableIndex, blockIndex));
                    adjacency.Add(ToTableAdjacencySnapshot(document.BodyElements, table.Table, tableIndex, blockIndex, previousKind, nextKind));
                    tableIndex++;
                    break;
                case DocxImplicitParagraphElement implicitParagraph:
                    blocks.Add(FromImplicitParagraph(blockIndex, previousKind, nextKind, implicitParagraph));
                    break;
                case DocxPageBreakElement pageBreak:
                    blocks.Add(FromPageBreak(blockIndex, previousKind, nextKind, pageBreak));
                    break;
                case DocxManualBreakElement manualBreak:
                    blocks.Add(FromManualBreak(blockIndex, previousKind, nextKind, manualBreak));
                    break;
                case DocxSectionBreakElement sectionBreak:
                    blocks.Add(FromSectionBreak(blockIndex, previousKind, nextKind, sectionBreak));
                    break;
                default:
                    blocks.Add(DocxStructureBlockSnapshot.ForUnknown(blockIndex, previousKind, nextKind));
                    break;
            }
        }

        DocxParagraph[] allParagraphs = EnumerateParagraphs(document).ToArray();
        DocxRevisionInfo[] documentRevisions = EnumerateDocumentRevisions(document).ToArray();
        DocxStructureCommentStoryAnchorSnapshot[] commentStoryAnchors = ToCommentStoryAnchorSnapshots(document).ToArray();
        return new DocxStructureSnapshot(
            document.MarkupMode.ToString(),
            document.BodyElements.Count,
            blocks.Count(block => block.Kind == "Paragraph"),
            tables.Count,
            blocks.Count(block => block.Kind == "PageBreak"),
            blocks.Count(block => block.Kind == "ManualBreak"),
            blocks.Count(block => block.Kind == "SectionBreak"),
            blocks.Count(block => block.Kind == "SectionBreak" && IsContinuousSectionBreak(block.SectionBreakTypeValue)),
            blocks.Count(block => block.Kind == "SectionBreak" && StartsNewPageSectionBreak(block.SectionBreakTypeValue)),
            blocks.Count(block => block.Kind == "SectionBreak" && block.SectionBreakTypeValue is null),
            blocks.Count(block => block.Kind == "SectionBreak" && HasSectionColumns(block)),
            blocks.Sum(block => block.TextLength),
            blocks.Sum(block => block.RevisionCount),
            blocks.Sum(block => block.InsertionRevisionCount),
            blocks.Sum(block => block.DeletionRevisionCount),
            blocks.Sum(block => block.MoveFromRevisionCount),
            blocks.Sum(block => block.MoveToRevisionCount),
            blocks.Sum(block => block.OtherRevisionCount),
            blocks.Sum(block => block.InlineImageCount),
            blocks.Sum(block => block.InlineReferenceCount),
            blocks.Sum(block => block.CommentReferenceCount),
            blocks.Sum(block => block.AnchoredInlineReferenceCount),
            blocks.Sum(block => block.ResolvedInlineReferenceCount),
            blocks.Select(block => block.MaxInlineReferenceTextOffsetInRun).DefaultIfEmpty(0).Max(),
            allParagraphs.Sum(ParagraphFieldReferenceCount),
            allParagraphs.Sum(ParagraphPageFieldReferenceCount),
            allParagraphs.Sum(ParagraphNumPagesFieldReferenceCount),
            allParagraphs.Sum(ParagraphOtherFieldReferenceCount),
            allParagraphs.Sum(ParagraphComplexFieldReferenceCount),
            allParagraphs.Sum(ParagraphCachedResultFieldReferenceCount),
            allParagraphs.Sum(ParagraphRenderedCachedResultFieldReferenceCount),
            allParagraphs.Sum(ParagraphPlaceholderFieldReferenceCount),
            allParagraphs.Sum(ParagraphNestedFieldReferenceCount),
            allParagraphs.Sum(ParagraphBookmarkAnchorCount),
            blocks.Sum(block => block.HyperlinkCount),
            blocks.Sum(block => block.ExternalHyperlinkCount),
            blocks.Sum(block => block.InternalHyperlinkCount),
            document.FloatingDrawings.Count,
            blocks,
            ToStorySnapshots(document, blocks),
            ToInlineReferenceSnapshots(document),
            document.FloatingDrawings.Select((drawing, index) => ToFloatingDrawingSnapshot(drawing, index)).ToArray(),
            document.StyleCatalog,
            ToStyleUsages(document),
            ToListUsages(document),
            tables,
            adjacency,
            ToCommentRangeSnapshots(document),
            ToRevisionRangeSnapshots(document),
            FormattingRevisionCount: CountFormattingRevisions(documentRevisions),
            RunFormattingRevisionCount: CountFormattingRevisions(documentRevisions, DocxRevisionPropertyFamily.Run),
            ParagraphFormattingRevisionCount: CountFormattingRevisions(documentRevisions, DocxRevisionPropertyFamily.Paragraph),
            TableFormattingRevisionCount: CountFormattingRevisions(documentRevisions, DocxRevisionPropertyFamily.Table),
            RowFormattingRevisionCount: CountFormattingRevisions(documentRevisions, DocxRevisionPropertyFamily.Row),
            CellFormattingRevisionCount: CountFormattingRevisions(documentRevisions, DocxRevisionPropertyFamily.Cell),
            SectionFormattingRevisionCount: CountFormattingRevisions(documentRevisions, DocxRevisionPropertyFamily.Section),
            FormattingRevisionProperties: ToFormattingRevisionPropertySnapshots(documentRevisions),
            PackageCommentAnchorIdCount: document.PackageCommentAnchorIds.Count,
            HiddenCommentAnchorIdCount: document.HiddenCommentAnchorIds.Count,
            ResolvedCommentStoryAnchorCount: commentStoryAnchors.Count(anchor => anchor.Status == "Visible"),
            HiddenCommentStoryAnchorCount: commentStoryAnchors.Count(anchor => anchor.Status == "HiddenByMarkupMode"),
            OrphanedCommentStoryAnchorCount: commentStoryAnchors.Count(anchor => anchor.Status == "Orphaned"),
            UnsupportedCommentStoryAnchorCount: commentStoryAnchors.Count(anchor => anchor.Status == "Unsupported"),
            CommentStoryAnchors: commentStoryAnchors,
            DynamicFieldReferenceCount: allParagraphs.Sum(ParagraphDynamicFieldReferenceCount),
            DynamicPlaceholderFieldReferenceCount: allParagraphs.Sum(ParagraphDynamicPlaceholderFieldReferenceCount),
            DynamicComplexWithoutCachedResultFieldReferenceCount: allParagraphs.Sum(ParagraphDynamicComplexWithoutCachedResultFieldReferenceCount),
            DynamicCachedResultNotRenderedFieldReferenceCount: allParagraphs.Sum(ParagraphDynamicCachedResultNotRenderedFieldReferenceCount));
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static IReadOnlyList<DocxStructureStorySnapshot> ToStorySnapshots(
        DocxDocument document,
        IReadOnlyList<DocxStructureBlockSnapshot> bodyBlocks)
    {
        DocxParagraph[] bodyParagraphs = DocxBlockTraversal.EnumerateBodyParagraphs(document).ToArray();
        var stories = new List<DocxStructureStorySnapshot>
        {
            new(
                "Body",
                "document",
                null,
                null,
                bodyBlocks.Count,
                bodyBlocks.Count(block => block.Kind == "Paragraph"),
                bodyBlocks.Count(block => block.Kind == "Table"),
                bodyBlocks.Sum(block => block.TextLength),
                bodyBlocks.Sum(block => block.RevisionCount),
                bodyBlocks.Sum(block => block.InsertionRevisionCount),
                bodyBlocks.Sum(block => block.DeletionRevisionCount),
                bodyBlocks.Sum(block => block.MoveFromRevisionCount),
                bodyBlocks.Sum(block => block.MoveToRevisionCount),
                bodyBlocks.Sum(block => block.OtherRevisionCount),
                bodyBlocks.Sum(block => block.InlineImageCount),
                bodyBlocks.Sum(block => block.InlineReferenceCount),
                bodyBlocks.Sum(block => block.CommentReferenceCount),
                bodyBlocks.Sum(block => block.ResolvedInlineReferenceCount),
                bodyParagraphs.Sum(ParagraphFieldReferenceCount),
                bodyParagraphs.Sum(ParagraphComplexFieldReferenceCount),
                bodyParagraphs.Sum(ParagraphCachedResultFieldReferenceCount),
                bodyParagraphs.Sum(ParagraphRenderedCachedResultFieldReferenceCount),
                bodyParagraphs.Sum(ParagraphPlaceholderFieldReferenceCount),
                bodyParagraphs.Sum(ParagraphNestedFieldReferenceCount),
                bodyParagraphs.Sum(ParagraphBookmarkAnchorCount),
                bodyBlocks.Sum(block => block.HyperlinkCount),
                bodyBlocks.Sum(block => block.ExternalHyperlinkCount),
                bodyBlocks.Sum(block => block.InternalHyperlinkCount),
                document.FloatingDrawings.Count,
                false,
                false,
                false,
                null,
                null,
                null,
                null)
        };

        AddStaticStories(stories, "Header", "document", null, document.HeaderBodyElementsByType, document.HeaderParagraphsByType, document.HeaderFloatingDrawingsByType, document.RelatedStories);
        AddStaticStories(stories, "Footer", "document", null, document.FooterBodyElementsByType, document.FooterParagraphsByType, document.FooterFloatingDrawingsByType, document.RelatedStories);
        AddRelatedStories(stories, document.RelatedStories);
        for (int blockIndex = 0; blockIndex < document.BodyElements.Count; blockIndex++)
        {
            if (document.BodyElements[blockIndex] is not DocxSectionBreakElement sectionBreak)
            {
                continue;
            }

            string scope = "section@" + blockIndex.ToString(CultureInfo.InvariantCulture);
            AddStaticStories(stories, "Header", scope, blockIndex, sectionBreak.PageSettings.HeaderBodyElementsByType, sectionBreak.PageSettings.HeaderParagraphsByType, sectionBreak.PageSettings.HeaderFloatingDrawingsByType, document.RelatedStories);
            AddStaticStories(stories, "Footer", scope, blockIndex, sectionBreak.PageSettings.FooterBodyElementsByType, sectionBreak.PageSettings.FooterParagraphsByType, sectionBreak.PageSettings.FooterFloatingDrawingsByType, document.RelatedStories);
        }

        return stories;
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static void AddStaticStories(
        List<DocxStructureStorySnapshot> stories,
        string kind,
        string scope,
        int? sectionBreakBlockIndex,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> drawingsByType,
        IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        string[] variantTypes = bodyElementsByType.Keys
            .Concat(paragraphsByType.Keys)
            .Concat(drawingsByType.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(type => type, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        foreach (string variantType in variantTypes)
        {
            IReadOnlyList<DocxBodyElement> bodyElements = DocxBlockTraversal.GetStaticStoryBodyElements(variantType, bodyElementsByType, paragraphsByType);
            DocxParagraph[] paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(bodyElements).ToArray();
            int directParagraphCount = bodyElements.OfType<DocxParagraphElement>().Count();
            int tableCount = DocxBlockTraversal.EnumerateBodyTables(bodyElements).Count();
            IReadOnlyList<DocxFloatingDrawing> drawings = drawingsByType.TryGetValue(variantType, out IReadOnlyList<DocxFloatingDrawing>? drawingList)
                ? drawingList
                : [];
            stories.Add(new DocxStructureStorySnapshot(
                kind,
                scope,
                sectionBreakBlockIndex,
                variantType,
                bodyElements.Count,
                directParagraphCount,
                tableCount,
                paragraphs.Sum(TextLength),
                paragraphs.Sum(CountRevisions),
                paragraphs.Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.Insertion)),
                paragraphs.Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.Deletion)),
                paragraphs.Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.MoveFrom)),
                paragraphs.Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.MoveTo)),
                paragraphs.Sum(CountOtherRevisions),
                paragraphs.Sum(paragraph => paragraph.Images.Count),
                paragraphs.Sum(ParagraphInlineReferenceCount),
                paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment)),
                paragraphs.Sum(paragraph => ParagraphResolvedInlineReferenceCount(paragraph, relatedStories)),
                paragraphs.Sum(ParagraphFieldReferenceCount),
                paragraphs.Sum(ParagraphComplexFieldReferenceCount),
                paragraphs.Sum(ParagraphCachedResultFieldReferenceCount),
                paragraphs.Sum(ParagraphRenderedCachedResultFieldReferenceCount),
                paragraphs.Sum(ParagraphPlaceholderFieldReferenceCount),
                paragraphs.Sum(ParagraphNestedFieldReferenceCount),
                paragraphs.Sum(ParagraphBookmarkAnchorCount),
                paragraphs.Sum(ParagraphHyperlinkCount),
                paragraphs.Sum(ParagraphExternalHyperlinkCount),
                paragraphs.Sum(ParagraphInternalHyperlinkCount),
                drawings.Count,
                false,
                false,
                false,
                null,
                null,
                null,
                null));
        }
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static void AddRelatedStories(List<DocxStructureStorySnapshot> stories, IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        foreach (DocxRelatedStory story in relatedStories
            .OrderBy(story => story.Kind.ToValueString(), StringComparer.Ordinal)
            .ThenBy(story => story.PartName, StringComparer.Ordinal)
            .ThenBy(story => story.Id, StringComparer.Ordinal))
        {
            DocxParagraph[] paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(story).ToArray();
            DocxTable[] tables = DocxBlockTraversal.EnumerateBodyTables(story).ToArray();
            int directParagraphCount = story.BodyElements.OfType<DocxParagraphElement>().Count();
            stories.Add(new DocxStructureStorySnapshot(
                story.Kind.ToValueString(),
                story.PartName,
                null,
                story.Id,
                story.BodyElements.Count,
                directParagraphCount,
                tables.Length,
                paragraphs.Sum(TextLength),
                paragraphs.Sum(CountRevisions),
                paragraphs.Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.Insertion)),
                paragraphs.Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.Deletion)),
                paragraphs.Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.MoveFrom)),
                paragraphs.Sum(paragraph => CountRevisions(paragraph, DocxRevisionKind.MoveTo)),
                paragraphs.Sum(CountOtherRevisions),
                paragraphs.Sum(paragraph => paragraph.Images.Count),
                paragraphs.Sum(ParagraphInlineReferenceCount),
                paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment)),
                paragraphs.Sum(paragraph => ParagraphResolvedInlineReferenceCount(paragraph, relatedStories)),
                paragraphs.Sum(ParagraphFieldReferenceCount),
                paragraphs.Sum(ParagraphComplexFieldReferenceCount),
                paragraphs.Sum(ParagraphCachedResultFieldReferenceCount),
                paragraphs.Sum(ParagraphRenderedCachedResultFieldReferenceCount),
                paragraphs.Sum(ParagraphPlaceholderFieldReferenceCount),
                paragraphs.Sum(ParagraphNestedFieldReferenceCount),
                paragraphs.Sum(ParagraphBookmarkAnchorCount),
                paragraphs.Sum(ParagraphHyperlinkCount),
                paragraphs.Sum(ParagraphExternalHyperlinkCount),
                paragraphs.Sum(ParagraphInternalHyperlinkCount),
                story.FloatingDrawings.Count,
                HasCommentAuthor: !string.IsNullOrWhiteSpace(story.CommentMetadata?.Author),
                HasCommentInitials: !string.IsNullOrWhiteSpace(story.CommentMetadata?.Initials),
                HasCommentDate: !string.IsNullOrWhiteSpace(story.CommentMetadata?.Date),
                CommentParagraphId: story.CommentMetadata?.ParagraphId,
                CommentParentParagraphId: story.CommentMetadata?.ParentParagraphId,
                CommentParentId: story.CommentMetadata?.ParentCommentId,
                CommentResolved: story.CommentMetadata?.IsResolved));
        }
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static IReadOnlyList<DocxStructureInlineReferenceSnapshot> ToInlineReferenceSnapshots(DocxDocument document)
    {
            (int? Index, DocxRelatedStory? Story) ResolveInlineReferenceStory(
                DocxInlineReference reference,
                IReadOnlyList<DocxRelatedStory> relatedStories)
            {
                return ResolveRelatedStory(reference.Kind, reference.Id, relatedStories);
            }

        DocxRelatedStory[] relatedStories = document.RelatedStories
            .OrderBy(story => story.Kind.ToValueString(), StringComparer.Ordinal)
            .ThenBy(story => story.PartName, StringComparer.Ordinal)
            .ThenBy(story => story.Id, StringComparer.Ordinal)
            .ToArray();
        var references = new List<DocxStructureInlineReferenceSnapshot>();
        foreach ((int blockIndex, string blockKind, int paragraphIndex, DocxParagraph paragraph) in EnumerateBodyReferenceParagraphs(document.BodyElements))
        {
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                (int? storyIndex, DocxRelatedStory? story) = ResolveInlineReferenceStory(reference, relatedStories);
                references.Add(new DocxStructureInlineReferenceSnapshot(
                    blockIndex,
                    blockKind,
                    paragraphIndex,
                    reference.Kind.ToValueString(),
                    reference.Id,
                    reference.CustomMarkFollowsValue,
                    reference.DisplayText,
                    reference.SourceRunIndex,
                    reference.RunChildIndex,
                    reference.TextOffsetInRun,
                    storyIndex,
                    story?.Kind.ToValueString(),
                    story?.PartName,
                    story?.Id,
                    story?.BodyElements.Count,
                    story is null ? null : DocxBlockTraversal.EnumerateBodyParagraphs(story).Sum(TextLength),
                    reference.Revisions.Count,
                    reference.Revision?.Kind.ToValueString(),
                    reference.Revision?.SourceElement));
            }
        }

        return references;
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static IReadOnlyList<DocxStructureCommentRangeSnapshot> ToCommentRangeSnapshots(DocxDocument document)
    {
        DocxRelatedStory[] relatedStories = document.RelatedStories
            .OrderBy(story => story.Kind.ToValueString(), StringComparer.Ordinal)
            .ThenBy(story => story.PartName, StringComparer.Ordinal)
            .ThenBy(story => story.Id, StringComparer.Ordinal)
            .ToArray();
        var ranges = new List<DocxStructureCommentRangeSnapshot>();
        foreach ((int blockIndex, string blockKind, int paragraphIndex, DocxParagraph paragraph) in EnumerateBodyReferenceParagraphs(document.BodyElements))
        {
            foreach (DocxCommentRange range in paragraph.CommentRanges)
            {
                (int? storyIndex, DocxRelatedStory? story) = ResolveRelatedStory(DocxRelatedStoryKind.Comment, range.Id, relatedStories);
                ranges.Add(new DocxStructureCommentRangeSnapshot(
                    blockIndex,
                    blockKind,
                    paragraphIndex,
                    range.Id,
                    range.StartSourceRunIndex,
                    range.StartTextOffset,
                    range.EndSourceRunIndex,
                    range.EndTextOffset,
                    range.ReferenceSourceRunIndex,
                    range.ReferenceTextOffset,
                    storyIndex,
                    story?.PartName,
                    story?.Id,
                    story?.BodyElements.Count,
                    story is null ? null : DocxBlockTraversal.EnumerateBodyParagraphs(story).Sum(TextLength)));
            }
        }

        return ranges;
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static IReadOnlyList<DocxStructureCommentStoryAnchorSnapshot> ToCommentStoryAnchorSnapshots(DocxDocument document)
    {
        Dictionary<string, int> visibleInlineCounts = EnumerateParagraphs(document)
            .SelectMany(paragraph => paragraph.InlineReferences)
            .Where(reference => reference.Kind == DocxRelatedStoryKind.Comment && !string.IsNullOrWhiteSpace(reference.Id))
            .GroupBy(reference => reference.Id ?? string.Empty, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        Dictionary<string, int> visibleRangeCounts = EnumerateParagraphs(document)
            .SelectMany(paragraph => paragraph.CommentRanges)
            .Where(range => !string.IsNullOrWhiteSpace(range.Id))
            .GroupBy(range => range.Id ?? string.Empty, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var packageAnchorIds = new HashSet<string>(document.PackageCommentAnchorIds, StringComparer.Ordinal);
        var hiddenAnchorIds = new HashSet<string>(document.HiddenCommentAnchorIds, StringComparer.Ordinal);
        return document.RelatedStories
            .Where(story => story.Kind == DocxRelatedStoryKind.Comment)
            .OrderBy(story => story.Id, StringComparer.Ordinal)
            .Select(story =>
            {
                string? id = story.Id;
                bool hasId = !string.IsNullOrWhiteSpace(id);
                int inlineCount = hasId && id is not null && visibleInlineCounts.TryGetValue(id, out int foundInlineCount) ? foundInlineCount : 0;
                int rangeCount = hasId && id is not null && visibleRangeCounts.TryGetValue(id, out int foundRangeCount) ? foundRangeCount : 0;
                bool hasPackageAnchor = hasId && id is not null && packageAnchorIds.Contains(id);
                bool hasHiddenAnchor = hasId && id is not null && hiddenAnchorIds.Contains(id);
                string status = !hasId
                    ? "Unsupported"
                    : inlineCount != 0 || rangeCount != 0
                        ? "Visible"
                        : hasHiddenAnchor ? "HiddenByMarkupMode" : "Orphaned";
                return new DocxStructureCommentStoryAnchorSnapshot(
                    id,
                    status,
                    hasPackageAnchor,
                    hasHiddenAnchor,
                    inlineCount,
                    rangeCount);
            })
            .ToArray();
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static IReadOnlyList<DocxStructureRevisionRangeSnapshot> ToRevisionRangeSnapshots(DocxDocument document)
    {
        var ranges = new List<DocxStructureRevisionRangeSnapshot>();
        foreach ((int blockIndex, string blockKind, int paragraphIndex, DocxParagraph paragraph) in EnumerateBodyReferenceParagraphs(document.BodyElements))
        {
            foreach (DocxRevisionRange range in paragraph.RevisionRanges)
            {
                ranges.Add(new DocxStructureRevisionRangeSnapshot(
                    blockIndex,
                    blockKind,
                    paragraphIndex,
                    range.Kind.ToValueString(),
                    range.Id,
                    !string.IsNullOrWhiteSpace(range.Name),
                    !string.IsNullOrWhiteSpace(range.Author),
                    !string.IsNullOrWhiteSpace(range.Date),
                    range.StartSourceRunIndex,
                    range.StartTextOffset,
                    range.EndSourceRunIndex,
                    range.EndTextOffset,
                    range.StartSourceRunIndex is not null && range.EndSourceRunIndex is not null));
            }
        }

        return LinkCrossBlockRevisionRanges(ranges);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static IReadOnlyList<DocxStructureRevisionRangeSnapshot> LinkCrossBlockRevisionRanges(IReadOnlyList<DocxStructureRevisionRangeSnapshot> ranges)
    {
        if (ranges.Count == 0)
        {
            return ranges;
        }

        DocxStructureRevisionRangeSnapshot[] linked = ranges.ToArray();
        var openStarts = new Dictionary<(string Kind, string? Id), int>();
        for (int index = 0; index < linked.Length; index++)
        {
            DocxStructureRevisionRangeSnapshot range = linked[index];
            if (range.IsClosed || string.IsNullOrWhiteSpace(range.Id))
            {
                continue;
            }

            var key = (range.Kind, range.Id);
            if (range.StartSourceRunIndex is not null && range.EndSourceRunIndex is null)
            {
                openStarts[key] = index;
                continue;
            }

            if (range.StartSourceRunIndex is not null || range.EndSourceRunIndex is null ||
                !openStarts.TryGetValue(key, out int startIndex))
            {
                continue;
            }

            DocxStructureRevisionRangeSnapshot start = linked[startIndex];
            linked[startIndex] = start with
            {
                IsLinkedAcrossBlocks = true,
                LinkedSourceBlockIndex = range.SourceBlockIndex,
                LinkedSourceParagraphIndex = range.SourceParagraphIndex
            };
            linked[index] = range with
            {
                IsLinkedAcrossBlocks = true,
                LinkedSourceBlockIndex = start.SourceBlockIndex,
                LinkedSourceParagraphIndex = start.SourceParagraphIndex
            };
            openStarts.Remove(key);
        }

        return linked;
    }

    private static IEnumerable<(int BlockIndex, string BlockKind, int ParagraphIndex, DocxParagraph Paragraph)> EnumerateBodyReferenceParagraphs(IReadOnlyList<DocxBodyElement> elements)
    {
        for (int blockIndex = 0; blockIndex < elements.Count; blockIndex++)
        {
            switch (elements[blockIndex])
            {
                case DocxParagraphElement paragraph:
                    yield return (blockIndex, "Paragraph", 0, paragraph.Paragraph);
                    break;
                case DocxTableElement table:
                    int paragraphIndex = 0;
                    foreach (DocxParagraph cellParagraph in DocxBlockTraversal.EnumerateTableParagraphs(table.Table))
                    {
                        yield return (blockIndex, "Table", paragraphIndex++, cellParagraph);
                    }

                    break;
            }
        }
    }

    private static (int? Index, DocxRelatedStory? Story) ResolveRelatedStory(
        DocxRelatedStoryKind kind,
        string? id,
        IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        if (id is null)
        {
            return (null, null);
        }

        for (int index = 0; index < relatedStories.Count; index++)
        {
            DocxRelatedStory story = relatedStories[index];
            if (story.Kind == kind &&
                string.Equals(story.Id, id, StringComparison.Ordinal))
            {
                return (index, story);
            }
        }

        return (null, null);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromParagraph(
        int blockIndex,
        string? previousKind,
        string? nextKind,
        DocxParagraph paragraph,
        IReadOnlyList<DocxRelatedStory> relatedStories)
    {
            bool HasVisibleText(DocxParagraph paragraph)
            {
                return paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text));
            }

        DocxEffectiveParagraphProperties effective = paragraph.EffectiveProperties;
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "Paragraph",
            previousKind,
            nextKind,
            ParagraphStyleId: effective.StyleId,
            ParagraphStyleFound: effective.StyleResolution.StyleFound,
            ParagraphStyleDepth: effective.StyleResolution.StyleDepth,
            HasDocumentDefaultParagraphProperties: effective.StyleResolution.HasDocumentDefaultParagraphProperties,
            HasDirectParagraphProperties: effective.StyleResolution.HasDirectParagraphProperties,
            HasTableStyleParagraphProperties: effective.StyleResolution.HasTableStyleParagraphProperties,
            RunCount: paragraph.Runs.Count,
            TextLength: TextLength(paragraph),
            HasVisibleText: HasVisibleText(paragraph),
            RevisionCount: CountRevisions(paragraph),
            InsertionRevisionCount: CountRevisions(paragraph, DocxRevisionKind.Insertion),
            DeletionRevisionCount: CountRevisions(paragraph, DocxRevisionKind.Deletion),
            MoveFromRevisionCount: CountRevisions(paragraph, DocxRevisionKind.MoveFrom),
            MoveToRevisionCount: CountRevisions(paragraph, DocxRevisionKind.MoveTo),
            OtherRevisionCount: CountOtherRevisions(paragraph),
            RevisionRangeCount: paragraph.RevisionRanges.Count,
            ListFormatValue: paragraph.ListLabel?.FormatValue,
            InlineImageCount: paragraph.Images.Count,
            InlineReferenceCount: paragraph.InlineReferences.Count,
            AnchoredInlineReferenceCount: paragraph.InlineReferences.Count(HasInlineReferenceAnchor),
            ResolvedInlineReferenceCount: ParagraphResolvedInlineReferenceCount(paragraph, relatedStories),
            MaxInlineReferenceTextOffsetInRun: paragraph.InlineReferences.Select(reference => reference.TextOffsetInRun).DefaultIfEmpty(0).Max(),
            FieldReferenceCount: paragraph.FieldReferences.Count,
            PageFieldReferenceCount: paragraph.FieldReferences.Count(reference => reference.Kind == DocxFieldKind.Page),
            NumPagesFieldReferenceCount: paragraph.FieldReferences.Count(reference => reference.Kind == DocxFieldKind.NumPages),
            OtherFieldReferenceCount: paragraph.FieldReferences.Count(reference => reference.Kind == DocxFieldKind.Other),
            ComplexFieldReferenceCount: ParagraphComplexFieldReferenceCount(paragraph),
            CachedResultFieldReferenceCount: ParagraphCachedResultFieldReferenceCount(paragraph),
            RenderedCachedResultFieldReferenceCount: ParagraphRenderedCachedResultFieldReferenceCount(paragraph),
            PlaceholderFieldReferenceCount: ParagraphPlaceholderFieldReferenceCount(paragraph),
            NestedFieldReferenceCount: ParagraphNestedFieldReferenceCount(paragraph),
            BookmarkAnchorCount: paragraph.BookmarkAnchors.Count,
            CommentReferenceCount: paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment),
            FootnoteReferenceCount: paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Footnote),
            EndnoteReferenceCount: paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Endnote),
            HyperlinkCount: paragraph.Hyperlinks.Count,
            ExternalHyperlinkCount: paragraph.Hyperlinks.Count(link => string.Equals(link.TargetMode, "External", StringComparison.OrdinalIgnoreCase)),
            InternalHyperlinkCount: paragraph.Hyperlinks.Count(link => link.Anchor is not null || link.ResolvedTarget is not null),
            SpacingBeforePoints: effective.SpacingBeforePoints,
            SpacingAfterPoints: effective.SpacingAfterPoints,
            LineSpacingPoints: effective.LineSpacingPoints,
            LineSpacingFactor: effective.LineSpacingFactor,
            HasBeforeSpacingToken: HasBeforeSpacingToken(effective.Spacing),
            HasAfterSpacingToken: HasAfterSpacingToken(effective.Spacing),
            BeforeAutoSpacingValue: effective.Spacing.BeforeAutoSpacingValue,
            AfterAutoSpacingValue: effective.Spacing.AfterAutoSpacingValue,
            ContextualSpacing: effective.Spacing.ContextualSpacing,
            KeepNext: effective.KeepRules.KeepNext,
            KeepLines: effective.KeepRules.KeepLines,
            WidowControl: effective.KeepRules.WidowControl,
            WordWrap: effective.WordWrap,
            WordWrapValue: effective.WordWrapValue,
            ParagraphIndentLeftPoints: effective.Indent.LeftPoints,
            ParagraphIndentRightPoints: effective.Indent.RightPoints,
            ParagraphIndentFirstLinePoints: effective.Indent.FirstLinePoints,
            ParagraphIndentHangingPoints: effective.Indent.HangingPoints,
            WhitespaceDelimitedTokenCount: CountWhitespaceDelimitedTokens(paragraph),
            LongestWhitespaceDelimitedTokenLength: LongestWhitespaceDelimitedTokenLength(paragraph),
            SpaceCharacterCount: CountCharacters(paragraph, static c => c == ' '),
            NonAsciiCharacterCount: CountCharacters(paragraph, static c => c > 127),
            PunctuationCharacterCount: CountCharacters(paragraph, char.IsPunctuation),
            DigitCharacterCount: CountCharacters(paragraph, char.IsDigit),
            UppercaseCharacterCount: CountCharacters(paragraph, char.IsUpper),
            LowercaseCharacterCount: CountCharacters(paragraph, char.IsLower),
            TabStopCount: effective.TabStops.Count,
            SnapToGrid: effective.SnapToGrid,
            SnapToGridValue: effective.SnapToGridValue,
            TableRowCount: 0,
            TableIndex: null,
            TableMaxColumnCount: null,
            TablePreferredWidthPoints: null,
            TablePreferredWidthType: null,
            TableIndentPoints: null,
            TableCellSpacingPoints: null,
            TableLayoutValue: null,
            PageBreakSourceKind: null,
            PageBreakValue: null,
            ManualBreakSourceKind: null,
            ManualBreakValue: null,
            ImplicitSourceKind: null,
            SectionBreakTypeValue: null,
            SectionColumnCountValue: null,
            SectionColumnEqualWidthValue: null,
            SectionColumnSpaceValue: null,
            SectionColumnDefinitionCount: null,
            SectionColumnDefinitionWidthTokenCount: null,
            SectionColumnDefinitionSpaceTokenCount: null,
            PageBreakConsumesParagraphLine: null,
            PageBreakLineSpacingPoints: null,
            PageBreakLineSpacingFactor: null,
            ManualBreakConsumesParagraphLine: null,
            ManualBreakLineSpacingPoints: null,
            ManualBreakLineSpacingFactor: null);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromTable(
        int blockIndex,
        string? previousKind,
        string? nextKind,
        DocxTable table,
        int tableIndex,
        IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        DocxParagraph[] paragraphs = DocxBlockTraversal.EnumerateTableParagraphs(table).ToArray();
        DocxRevisionInfo[] revisions = EnumerateTableRevisions(table).ToArray();
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "Table",
            previousKind,
            nextKind,
            RevisionCount: CountRevisions(revisions),
            InsertionRevisionCount: CountRevisions(revisions, DocxRevisionKind.Insertion),
            DeletionRevisionCount: CountRevisions(revisions, DocxRevisionKind.Deletion),
            MoveFromRevisionCount: CountRevisions(revisions, DocxRevisionKind.MoveFrom),
            MoveToRevisionCount: CountRevisions(revisions, DocxRevisionKind.MoveTo),
            OtherRevisionCount: CountOtherRevisions(revisions),
            RevisionRangeCount: paragraphs.Sum(paragraph => paragraph.RevisionRanges.Count),
            InlineReferenceCount: paragraphs.Sum(ParagraphInlineReferenceCount),
            AnchoredInlineReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(HasInlineReferenceAnchor)),
            ResolvedInlineReferenceCount: paragraphs.Sum(paragraph => ParagraphResolvedInlineReferenceCount(paragraph, relatedStories)),
            MaxInlineReferenceTextOffsetInRun: paragraphs.SelectMany(paragraph => paragraph.InlineReferences).Select(reference => reference.TextOffsetInRun).DefaultIfEmpty(0).Max(),
            CommentReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Comment)),
            FootnoteReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Footnote)),
            EndnoteReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == DocxRelatedStoryKind.Endnote)),
            FieldReferenceCount: paragraphs.Sum(ParagraphFieldReferenceCount),
            PageFieldReferenceCount: paragraphs.Sum(ParagraphPageFieldReferenceCount),
            NumPagesFieldReferenceCount: paragraphs.Sum(ParagraphNumPagesFieldReferenceCount),
            OtherFieldReferenceCount: paragraphs.Sum(ParagraphOtherFieldReferenceCount),
            ComplexFieldReferenceCount: paragraphs.Sum(ParagraphComplexFieldReferenceCount),
            CachedResultFieldReferenceCount: paragraphs.Sum(ParagraphCachedResultFieldReferenceCount),
            RenderedCachedResultFieldReferenceCount: paragraphs.Sum(ParagraphRenderedCachedResultFieldReferenceCount),
            PlaceholderFieldReferenceCount: paragraphs.Sum(ParagraphPlaceholderFieldReferenceCount),
            NestedFieldReferenceCount: paragraphs.Sum(ParagraphNestedFieldReferenceCount),
            BookmarkAnchorCount: paragraphs.Sum(ParagraphBookmarkAnchorCount),
            HyperlinkCount: paragraphs.Sum(ParagraphHyperlinkCount),
            ExternalHyperlinkCount: paragraphs.Sum(ParagraphExternalHyperlinkCount),
            InternalHyperlinkCount: paragraphs.Sum(ParagraphInternalHyperlinkCount),
            TableRowCount: table.Rows.Count,
            TableIndex: tableIndex,
            TableMaxColumnCount: MaxColumnCount(table),
            TablePreferredWidthPoints: table.PreferredWidthPoints,
            TablePreferredWidthType: table.PreferredWidthType,
            TableIndentPoints: table.IndentPoints,
            TableCellSpacingPoints: table.CellSpacingPoints,
            TableLayoutValue: table.LayoutValue,
            ParagraphStyleId: null,
            ParagraphStyleFound: null,
            ParagraphStyleDepth: null,
            HasDocumentDefaultParagraphProperties: null,
            HasDirectParagraphProperties: null,
            HasTableStyleParagraphProperties: null,
            RunCount: 0,
            TextLength: 0,
            HasVisibleText: false,
            ListFormatValue: null,
            InlineImageCount: 0,
            SpacingBeforePoints: null,
            SpacingAfterPoints: null,
            LineSpacingPoints: null,
            LineSpacingFactor: null,
            HasBeforeSpacingToken: false,
            HasAfterSpacingToken: false,
            BeforeAutoSpacingValue: null,
            AfterAutoSpacingValue: null,
            ContextualSpacing: null,
            KeepNext: null,
            KeepLines: null,
            WidowControl: null,
            WordWrap: null,
            WordWrapValue: null,
            PageBreakSourceKind: null,
            PageBreakValue: null,
            ManualBreakSourceKind: null,
            ManualBreakValue: null,
            ImplicitSourceKind: null,
            SectionBreakTypeValue: null,
            SectionColumnCountValue: null,
            SectionColumnEqualWidthValue: null,
            SectionColumnSpaceValue: null,
            SectionColumnDefinitionCount: null,
            SectionColumnDefinitionWidthTokenCount: null,
            SectionColumnDefinitionSpaceTokenCount: null,
            ParagraphIndentLeftPoints: null,
            ParagraphIndentRightPoints: null,
            ParagraphIndentFirstLinePoints: null,
            ParagraphIndentHangingPoints: null,
            WhitespaceDelimitedTokenCount: null,
            LongestWhitespaceDelimitedTokenLength: null,
            SpaceCharacterCount: null,
            NonAsciiCharacterCount: null,
            PunctuationCharacterCount: null,
            DigitCharacterCount: null,
            UppercaseCharacterCount: null,
            LowercaseCharacterCount: null,
            TabStopCount: null,
            SnapToGrid: null,
            SnapToGridValue: null,
            PageBreakConsumesParagraphLine: null,
            PageBreakLineSpacingPoints: null,
            PageBreakLineSpacingFactor: null,
            ManualBreakConsumesParagraphLine: null,
            ManualBreakLineSpacingPoints: null,
            ManualBreakLineSpacingFactor: null);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromPageBreak(int blockIndex, string? previousKind, string? nextKind, DocxPageBreakElement pageBreak)
    {
        return DocxStructureBlockSnapshot.ForPageBreak(blockIndex, previousKind, nextKind, pageBreak);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromImplicitParagraph(int blockIndex, string? previousKind, string? nextKind, DocxImplicitParagraphElement implicitParagraph)
    {
        return DocxStructureBlockSnapshot.ForImplicitParagraph(blockIndex, previousKind, nextKind, implicitParagraph);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromManualBreak(int blockIndex, string? previousKind, string? nextKind, DocxManualBreakElement manualBreak)
    {
        return DocxStructureBlockSnapshot.ForManualBreak(blockIndex, previousKind, nextKind, manualBreak);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureBlockSnapshot FromSectionBreak(int blockIndex, string? previousKind, string? nextKind, DocxSectionBreakElement sectionBreak)
    {
        return DocxStructureBlockSnapshot.ForSectionBreak(blockIndex, previousKind, nextKind, CountRevisions(sectionBreak.Revisions), CountRevisions(sectionBreak.Revisions, DocxRevisionKind.Insertion), CountRevisions(sectionBreak.Revisions, DocxRevisionKind.Deletion), CountRevisions(sectionBreak.Revisions, DocxRevisionKind.MoveFrom), CountRevisions(sectionBreak.Revisions, DocxRevisionKind.MoveTo), CountOtherRevisions(sectionBreak.Revisions), sectionBreak.TypeValue, sectionBreak.ColumnCountValue, sectionBreak.ColumnEqualWidthValue, sectionBreak.ColumnSpaceValue, sectionBreak.ColumnDefinitions.Count, sectionBreak.ColumnDefinitions.Count(column => column.WidthValue is not null), sectionBreak.ColumnDefinitions.Count(column => column.SpaceValue is not null));
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureTableSnapshot ToTableSnapshot(DocxTable table, int tableIndex, int blockIndex)
    {
        int cellCount = table.Rows.Sum(row => row.Cells.Count);
        DocxParagraph[] tableParagraphs = DocxBlockTraversal.EnumerateTableParagraphs(table).ToArray();
        DocxRevisionInfo[] tableRevisions = EnumerateTableRevisions(table).ToArray();
        DocxTableCellBorder[] tableCellBorders = table.Rows.SelectMany(row => row.Cells.SelectMany(cell => cell.Borders)).ToArray();
        int paragraphCount = tableParagraphs.Length;
        return new DocxStructureTableSnapshot(
            tableIndex,
            blockIndex,
            table.StyleId,
            table.Rows.Count,
            MaxColumnCount(table),
            table.ColumnWidthsPoints.Count,
            table.ColumnWidthsPoints.Sum(),
            table.HasExplicitGrid,
            table.PreferredWidthPoints,
            table.PreferredWidthValue,
            table.PreferredWidthType,
            table.IndentPoints,
            table.IndentValue,
            table.IndentType,
            table.CellSpacingPoints,
            table.CellSpacingValue,
            table.CellSpacingType,
            table.LayoutValue,
            table.Rows.Count(row => row.IsHeader),
            table.Rows.Count(row => row.CantSplit),
            table.Rows.Count(row => row.HeightPoints is not null),
            table.Rows.Count(row => string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase)),
            table.Rows.Count(row => string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase)),
            table.Rows.Count(row => row.TablePropertyExceptionCellMargins is not null),
            cellCount,
            table.Rows.Sum(row => row.Cells.Count(cell => cell.GridSpan > 1)),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.HasVerticalMerge)),
            table.Rows.Sum(row => row.Cells.Count(cell => string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase))),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.FillHex is not null || cell.ShadingValue is not null)),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.VerticalAlignmentValue is not null)),
            table.Rows.Sum(row => row.Cells.Count(cell => cell.PreferredWidthPoints is not null || cell.PreferredWidthValue is not null)),
            tableCellBorders.Count(IsVisibleBorder),
            tableCellBorders.Count(border => IsBorderStyle(border, "single")),
            tableCellBorders.Count(border => IsBorderStyle(border, "thick")),
            tableCellBorders.Count(border => IsBorderStyle(border, "double")),
            tableCellBorders.Count(border => IsBorderStyle(border, "dotted")),
            tableCellBorders.Count(IsDashedBorderStyle),
            tableCellBorders.Count(IsSuppressedBorder),
            tableCellBorders.Count(border => IsVisibleBorder(border) &&
                !IsBorderStyle(border, "single") &&
                !IsBorderStyle(border, "thick") &&
                !IsBorderStyle(border, "double") &&
                !IsBorderStyle(border, "triple") &&
                !IsBorderStyle(border, "dotted") &&
                !IsDashedBorderStyle(border) &&
                !IsCompoundBorderStyle(border) &&
                !IsThreeDBorderStyle(border) &&
                !IsWaveBorderStyle(border) &&
                !IsSupportedSolidBorderStyle(border)),
            paragraphCount,
            tableParagraphs.Sum(paragraph => paragraph.Runs.Count),
            tableParagraphs.Sum(paragraph => TextLength(paragraph)),
            CountRevisions(tableRevisions),
            CountRevisions(tableRevisions, DocxRevisionKind.Insertion),
            CountRevisions(tableRevisions, DocxRevisionKind.Deletion),
            CountRevisions(tableRevisions, DocxRevisionKind.MoveFrom),
            CountRevisions(tableRevisions, DocxRevisionKind.MoveTo),
            CountOtherRevisions(tableRevisions),
            tableParagraphs.Sum(paragraph => CountWhitespaceDelimitedTokens(paragraph)),
            tableParagraphs.Select(paragraph => LongestWhitespaceDelimitedTokenLength(paragraph)).DefaultIfEmpty(0).Max(),
            tableParagraphs.Sum(paragraph => paragraph.Images.Count),
            tableParagraphs.Sum(ParagraphInlineReferenceCount),
            tableParagraphs.Sum(ParagraphHyperlinkCount),
            tableParagraphs.Sum(ParagraphExternalHyperlinkCount),
            tableParagraphs.Sum(ParagraphInternalHyperlinkCount),
            tableParagraphs.Count(paragraph => paragraph.ListLabel is not null),
            tableParagraphs.Count(HasEffectiveKeepConstraint),
            table.Look?.FirstRow,
            table.Look?.FirstColumn,
            table.Look?.NoHorizontalBand,
            table.Look?.NoVerticalBand,
            table.Rows.Select((row, rowIndex) => ToTableRowSnapshot(row, rowIndex)).ToArray());
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureTableRowSnapshot ToTableRowSnapshot(DocxTableRow row, int rowIndex)
    {
        DocxStructureTableCellSnapshot[] cells = row.Cells.Select(ToTableCellSnapshot).ToArray();
        DocxRevisionInfo[] rowRevisions = row.Revisions
            .Concat(row.Cells.SelectMany(EnumerateTableCellRevisions))
            .ToArray();
        return new DocxStructureTableRowSnapshot(
            rowIndex,
            row.Cells.Count,
            row.Cells.Sum(cell => Math.Max(1, cell.GridSpan)),
            row.IsHeader,
            row.HeaderValue,
            row.CantSplit,
            row.CantSplitValue,
            row.HeightPoints,
            row.HeightValue,
            row.HeightRuleValue,
            row.TablePropertyExceptionCellMargins is not null,
            cells.Count(cell => cell.GridSpan > 1),
            cells.Count(cell => cell.HasVerticalMerge),
            cells.Count(cell => string.Equals(cell.VerticalMergeValue, "restart", StringComparison.OrdinalIgnoreCase)),
            cells.Count(cell => cell.HasShading),
            cells.Count(cell => cell.VerticalAlignmentValue is not null),
            cells.Count(cell => cell.HasPreferredWidth),
            cells.Sum(cell => cell.VisibleBorderCount),
            cells.Sum(cell => cell.ParagraphCount),
            cells.Sum(cell => cell.RunCount),
            cells.Sum(cell => cell.TextLength),
            CountRevisions(rowRevisions),
            CountRevisions(rowRevisions, DocxRevisionKind.Insertion),
            CountRevisions(rowRevisions, DocxRevisionKind.Deletion),
            CountRevisions(rowRevisions, DocxRevisionKind.MoveFrom),
            CountRevisions(rowRevisions, DocxRevisionKind.MoveTo),
            CountOtherRevisions(rowRevisions),
            cells.Sum(cell => cell.WhitespaceDelimitedTokenCount),
            cells.Select(cell => cell.LongestWhitespaceDelimitedTokenLength).DefaultIfEmpty(0).Max(),
            cells.Sum(cell => cell.InlineImageCount),
            cells.Sum(cell => cell.InlineReferenceCount),
            cells.Sum(cell => cell.HyperlinkCount),
            cells.Sum(cell => cell.ExternalHyperlinkCount),
            cells.Sum(cell => cell.InternalHyperlinkCount),
            cells.Sum(cell => cell.NumberedParagraphCount),
            cells.Sum(cell => cell.KeepRuleParagraphCount),
            cells.Sum(cell => cell.BeforeSpacingTokenParagraphCount),
            cells.Sum(cell => cell.AfterSpacingTokenParagraphCount),
            cells.Select(cell => cell.MaxFontSize).DefaultIfEmpty(0d).Max(),
            cells);
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureTableCellSnapshot ToTableCellSnapshot(DocxTableCell cell, int cellIndex)
    {
        DocxParagraph[] paragraphs = DocxBlockTraversal
            .EnumerateBodyParagraphs(DocxTableCellContent.GetBodyElements(cell))
            .ToArray();
        DocxRevisionInfo[] cellRevisions = cell.Revisions
            .Concat(paragraphs.SelectMany(paragraph => paragraph.Revisions))
            .ToArray();
        return new DocxStructureTableCellSnapshot(
            cellIndex,
            Math.Max(1, cell.GridSpan),
            cell.GridSpanValue,
            cell.HasVerticalMerge,
            cell.VerticalMergeValue,
            cell.FillHex is not null || cell.ShadingValue is not null,
            cell.ShadingValue,
            cell.VerticalAlignmentValue,
            cell.PreferredWidthPoints is not null || cell.PreferredWidthValue is not null,
            cell.PreferredWidthPoints,
            cell.PreferredWidthValue,
            cell.PreferredWidthType,
            cell.Borders.Count(border => !string.Equals(border.Value, "nil", StringComparison.OrdinalIgnoreCase) && !string.Equals(border.Value, "none", StringComparison.OrdinalIgnoreCase)),
            paragraphs.Length,
            paragraphs.Sum(paragraph => paragraph.Runs.Count),
            paragraphs.Sum(paragraph => TextLength(paragraph)),
            CountRevisions(cellRevisions),
            CountRevisions(cellRevisions, DocxRevisionKind.Insertion),
            CountRevisions(cellRevisions, DocxRevisionKind.Deletion),
            CountRevisions(cellRevisions, DocxRevisionKind.MoveFrom),
            CountRevisions(cellRevisions, DocxRevisionKind.MoveTo),
            CountOtherRevisions(cellRevisions),
            paragraphs.Sum(paragraph => CountWhitespaceDelimitedTokens(paragraph)),
            paragraphs.Select(paragraph => LongestWhitespaceDelimitedTokenLength(paragraph)).DefaultIfEmpty(0).Max(),
            paragraphs.Sum(paragraph => paragraph.Images.Count),
            paragraphs.Sum(ParagraphInlineReferenceCount),
            paragraphs.Sum(ParagraphHyperlinkCount),
            paragraphs.Sum(ParagraphExternalHyperlinkCount),
            paragraphs.Sum(ParagraphInternalHyperlinkCount),
            paragraphs.Count(paragraph => paragraph.ListLabel is not null),
            paragraphs.Count(HasEffectiveKeepConstraint),
            paragraphs.Count(paragraph => HasBeforeSpacingToken(paragraph.EffectiveProperties.Spacing)),
            paragraphs.Count(paragraph => HasAfterSpacingToken(paragraph.EffectiveProperties.Spacing)),
            paragraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.EffectiveProperties.FontSize).DefaultIfEmpty(0d).Max(),
            CountCharacters(paragraphs, static c => c == ' '),
            CountCharacters(paragraphs, static c => c > 127),
            CountCharacters(paragraphs, char.IsPunctuation),
            CountCharacters(paragraphs, char.IsDigit),
            CountCharacters(paragraphs, char.IsUpper),
            CountCharacters(paragraphs, char.IsLower),
            cell.BodyElements.Count,
            cell.BodyElements.OfType<DocxManualBreakElement>().Count(),
            cell.BodyElements.OfType<DocxPageBreakElement>().Count(),
            cell.BodyElements.OfType<DocxTableElement>().Count());
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureFloatingDrawingSnapshot ToFloatingDrawingSnapshot(DocxFloatingDrawing drawing, int index)
    {
        return new DocxStructureFloatingDrawingSnapshot(
            index,
            drawing.WrapKind?.ToValueString(),
            drawing.WrapTextValue,
            drawing.BehindDocumentValue,
            drawing.LayoutInCellValue,
            drawing.AllowOverlapValue,
            drawing.HorizontalRelativeFromValue,
            drawing.HorizontalAlignValue,
            drawing.HorizontalOffsetValue,
            drawing.VerticalRelativeFromValue,
            drawing.VerticalAlignValue,
            drawing.VerticalOffsetValue,
            drawing.ExtentCxValue,
            drawing.ExtentCyValue,
            drawing.DistanceTopValue,
            drawing.DistanceBottomValue,
            drawing.DistanceLeftValue,
            drawing.DistanceRightValue,
            drawing.SimplePositionValue,
            drawing.RelativeHeightValue,
            drawing.LockedValue,
            drawing.ImageRelationshipId,
            drawing.Image?.PartName,
            drawing.Image?.ContentType,
            drawing.Image?.WidthPoints,
            drawing.Image?.HeightPoints,
            drawing.SourceParagraphIndex,
            drawing.SourceBlockIndex,
            drawing.TextBoxBodyElements.Count,
            DocxBlockTraversal.EnumerateBodyParagraphs(drawing.TextBoxBodyElements).Count(),
            DocxBlockTraversal.EnumerateBodyParagraphs(drawing.TextBoxBodyElements).Sum(TextLength),
            CountRevisions(drawing.Revisions),
            CountRevisions(drawing.Revisions, DocxRevisionKind.Insertion),
            CountRevisions(drawing.Revisions, DocxRevisionKind.Deletion),
            CountRevisions(drawing.Revisions, DocxRevisionKind.MoveFrom),
            CountRevisions(drawing.Revisions, DocxRevisionKind.MoveTo),
            CountOtherRevisions(drawing.Revisions));
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static IReadOnlyList<DocxStructureStyleUsageSnapshot> ToStyleUsages(DocxDocument document)
    {
            bool IsExactLineSpacing(DocxParagraphSpacing spacing)
            {
                return spacing.LineValue is not null &&
                    string.Equals(spacing.LineRuleValue, "exact", StringComparison.OrdinalIgnoreCase);
            }

            bool IsAtLeastLineSpacing(DocxParagraphSpacing spacing)
            {
                return spacing.LineValue is not null &&
                    string.Equals(spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase);
            }

            bool IsAutoLineSpacing(DocxParagraphSpacing spacing)
            {
                return spacing.LineValue is not null &&
                    (spacing.LineRuleValue is null || string.Equals(spacing.LineRuleValue, "auto", StringComparison.OrdinalIgnoreCase));
            }

        DocxStructureStyleUsageSnapshot[] paragraphStyles = EnumerateParagraphs(document)
            .GroupBy(paragraph => paragraph.EffectiveProperties.StyleId, StringComparer.Ordinal)
            .Select(group => new DocxStructureStyleUsageSnapshot(
                "Paragraph",
                group.Key,
                group.Count(),
                group.Count(),
                0,
                group.Sum(TextLength),
                group.Count(paragraph => HasBeforeSpacingToken(paragraph.EffectiveProperties.Spacing)),
                group.Count(paragraph => HasAfterSpacingToken(paragraph.EffectiveProperties.Spacing)),
                group.Count(paragraph => paragraph.EffectiveProperties.Spacing.BeforeAutoSpacingValue is not null),
                group.Count(paragraph => paragraph.EffectiveProperties.Spacing.AfterAutoSpacingValue is not null),
                group.Count(paragraph => paragraph.EffectiveProperties.Spacing.BeforeLinesValue is not null),
                group.Count(paragraph => paragraph.EffectiveProperties.Spacing.AfterLinesValue is not null),
                group.Count(paragraph => paragraph.EffectiveProperties.Spacing.ContextualSpacing == true),
                group.Count(paragraph => IsExactLineSpacing(paragraph.EffectiveProperties.Spacing)),
                group.Count(paragraph => IsAtLeastLineSpacing(paragraph.EffectiveProperties.Spacing)),
                group.Count(paragraph => IsAutoLineSpacing(paragraph.EffectiveProperties.Spacing)),
                group.Count(paragraph => paragraph.EffectiveProperties.StyleResolution.HasTableStyleParagraphProperties)))
            .ToArray();
        DocxStructureStyleUsageSnapshot[] tableStyles = DocxBlockTraversal.EnumerateBodyTables(document)
            .Concat(document.RelatedStories.SelectMany(DocxBlockTraversal.EnumerateBodyTables))
            .GroupBy(table => table.StyleId, StringComparer.Ordinal)
            .Select(group => new DocxStructureStyleUsageSnapshot(
                "Table",
                group.Key,
                group.Count(),
                0,
                group.Count(),
                group.Sum(table => DocxBlockTraversal.EnumerateTableParagraphs(table).Sum(paragraph => TextLength(paragraph))),
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0,
                0))
            .ToArray();
        return paragraphStyles.Concat(tableStyles).ToArray();
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static IReadOnlyList<DocxStructureListUsageSnapshot> ToListUsages(DocxDocument document)
    {
            bool HasParagraphIndentOverride(DocxParagraphIndent indent)
            {
                return indent.LeftPoints is not null ||
                    indent.RightPoints is not null ||
                    indent.FirstLinePoints is not null ||
                    indent.HangingPoints is not null ||
                    indent.LeftValue is not null ||
                    indent.RightValue is not null ||
                    indent.FirstLineValue is not null ||
                    indent.HangingValue is not null;
            }

        return EnumerateParagraphs(document)
            .SelectMany(paragraph => paragraph.ListLabel is { } label ? new[] { (Paragraph: paragraph, Label: label) } : [])
            .GroupBy(item => new
            {
                item.Label.NumberId,
                item.Label.Level,
                item.Label.FormatValue,
                item.Label.SuffixValue
            })
            .Select(group => new DocxStructureListUsageSnapshot(
                group.Key.NumberId,
                group.Key.Level,
                group.Key.FormatValue,
                group.Key.SuffixValue,
                group.Count(),
                group.Sum(item => TextLength(item.Paragraph)),
                group.Count(item => item.Label.Indent.LeftPoints is not null),
                group.Count(item => item.Label.Indent.RightPoints is not null),
                group.Count(item => item.Label.Indent.FirstLinePoints is not null),
                group.Count(item => item.Label.Indent.HangingPoints is not null),
                group.Count(item => item.Label.Indent.NumberingTabPositionPoints is not null),
                group.Count(item => HasParagraphIndentOverride(item.Paragraph.EffectiveProperties.Indent)),
                group.Count(item => item.Paragraph.EffectiveProperties.TabStops.Any(tab => string.Equals(tab.Value, "num", StringComparison.OrdinalIgnoreCase)))))
            .ToArray();
    }

    private static IEnumerable<DocxParagraph> EnumerateParagraphs(DocxDocument document)
    {
        return DocxBlockTraversal.EnumerateBodyParagraphs(document)
            .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.HeaderBodyElementsByType, document.HeaderParagraphsByType))
            .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.FooterBodyElementsByType, document.FooterParagraphsByType))
            .Concat(DocxBlockTraversal.EnumerateStaticStoryParagraphs(document.PageSettings))
            .Concat(document.BodyElements
                .OfType<DocxSectionBreakElement>()
                .SelectMany(sectionBreak => DocxBlockTraversal.EnumerateStaticStoryParagraphs(sectionBreak.PageSettings)))
            .Concat(document.RelatedStories.SelectMany(DocxBlockTraversal.EnumerateBodyParagraphs));
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static IReadOnlyList<DocxStructureFormattingRevisionPropertySnapshot> ToFormattingRevisionPropertySnapshots(IEnumerable<DocxRevisionInfo> revisions)
    {
        return revisions
            .Where(revision => revision.PropertyChangeFamily is not null)
            .SelectMany(revision => revision.PropertyElementNames.Select(name => new
            {
                Family = revision.PropertyChangeFamily?.ToValueString() ?? string.Empty,
                revision.SourceElement,
                PropertyElementName = name
            }))
            .GroupBy(item => (item.Family, item.SourceElement, item.PropertyElementName))
            .Select(group => new DocxStructureFormattingRevisionPropertySnapshot(
                group.Key.Family,
                group.Key.SourceElement,
                group.Key.PropertyElementName,
                group.Count()))
            .OrderBy(snapshot => snapshot.Family, StringComparer.Ordinal)
            .ThenBy(snapshot => snapshot.SourceElement, StringComparer.Ordinal)
            .ThenBy(snapshot => snapshot.PropertyElementName, StringComparer.Ordinal)
            .ToArray();
    }

    // Single-caller pipeline stage; kept static.
    private static IEnumerable<DocxRevisionInfo> EnumerateDocumentRevisions(DocxDocument document)
    {
        foreach (DocxRevisionInfo revision in EnumerateBodyElementRevisions(document.BodyElements))
        {
            yield return revision;
        }

        foreach (DocxFloatingDrawing drawing in document.FloatingDrawings)
        {
            foreach (DocxRevisionInfo revision in EnumerateFloatingDrawingRevisions(drawing))
            {
                yield return revision;
            }
        }

        foreach (DocxRevisionInfo revision in EnumerateStaticStoryRevisions(
            document.HeaderBodyElementsByType,
            document.HeaderParagraphsByType,
            document.HeaderFloatingDrawingsByType))
        {
            yield return revision;
        }

        foreach (DocxRevisionInfo revision in EnumerateStaticStoryRevisions(
            document.FooterBodyElementsByType,
            document.FooterParagraphsByType,
            document.FooterFloatingDrawingsByType))
        {
            yield return revision;
        }

        foreach (DocxRevisionInfo revision in EnumeratePageSettingsRevisions(document.PageSettings))
        {
            yield return revision;
        }

        foreach (DocxSectionBreakElement sectionBreak in document.BodyElements.OfType<DocxSectionBreakElement>())
        {
            foreach (DocxRevisionInfo revision in EnumeratePageSettingsRevisions(sectionBreak.PageSettings))
            {
                yield return revision;
            }
        }

        foreach (DocxRelatedStory story in document.RelatedStories)
        {
            foreach (DocxRevisionInfo revision in EnumerateBodyElementRevisions(story.BodyElements))
            {
                yield return revision;
            }

            foreach (DocxFloatingDrawing drawing in story.FloatingDrawings)
            {
                foreach (DocxRevisionInfo revision in EnumerateFloatingDrawingRevisions(drawing))
                {
                    yield return revision;
                }
            }
        }

        if (document.FinalSectionBreak is { } finalSectionBreak)
        {
            foreach (DocxRevisionInfo revision in finalSectionBreak.Revisions)
            {
                yield return revision;
            }
        }
    }

    // Single-caller pipeline stage; kept static.
    private static IEnumerable<DocxRevisionInfo> EnumeratePageSettingsRevisions(DocxPageSettings settings)
    {
        foreach (DocxRevisionInfo revision in EnumerateStaticStoryRevisions(
            settings.HeaderBodyElementsByType,
            settings.HeaderParagraphsByType,
            settings.HeaderFloatingDrawingsByType))
        {
            yield return revision;
        }

        foreach (DocxRevisionInfo revision in EnumerateStaticStoryRevisions(
            settings.FooterBodyElementsByType,
            settings.FooterParagraphsByType,
            settings.FooterFloatingDrawingsByType))
        {
            yield return revision;
        }
    }

    private static IEnumerable<DocxRevisionInfo> EnumerateStaticStoryRevisions(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> fallbackParagraphsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> floatingDrawingsByType)
    {
        if (bodyElementsByType.Count == 0)
        {
            foreach (DocxParagraph paragraph in fallbackParagraphsByType.Values.SelectMany(paragraphs => paragraphs))
            {
                foreach (DocxRevisionInfo revision in paragraph.Revisions)
                {
                    yield return revision;
                }
            }
        }
        else
        {
            foreach (IReadOnlyList<DocxBodyElement> bodyElements in bodyElementsByType.Values)
            {
                foreach (DocxRevisionInfo revision in EnumerateBodyElementRevisions(bodyElements))
                {
                    yield return revision;
                }
            }
        }

        foreach (DocxFloatingDrawing drawing in floatingDrawingsByType.Values.SelectMany(drawings => drawings))
        {
            foreach (DocxRevisionInfo revision in EnumerateFloatingDrawingRevisions(drawing))
            {
                yield return revision;
            }
        }
    }

    private static IEnumerable<DocxRevisionInfo> EnumerateFloatingDrawingRevisions(DocxFloatingDrawing drawing)
    {
        foreach (DocxRevisionInfo revision in drawing.Revisions)
        {
            yield return revision;
        }

        foreach (DocxRevisionInfo revision in EnumerateBodyElementRevisions(drawing.TextBoxBodyElements))
        {
            yield return revision;
        }
    }

    private static IEnumerable<DocxRevisionInfo> EnumerateBodyElementRevisions(IEnumerable<DocxBodyElement> bodyElements)
    {
        foreach (DocxBodyElement element in bodyElements)
        {
            switch (element)
            {
                case DocxParagraphElement paragraph:
                    foreach (DocxRevisionInfo revision in paragraph.Paragraph.Revisions)
                    {
                        yield return revision;
                    }

                    break;
                case DocxTableElement table:
                    foreach (DocxRevisionInfo revision in EnumerateTableRevisions(table.Table))
                    {
                        yield return revision;
                    }

                    break;
                case DocxSectionBreakElement sectionBreak:
                    foreach (DocxRevisionInfo revision in sectionBreak.Revisions)
                    {
                        yield return revision;
                    }

                    break;
            }
        }
    }

    private static IEnumerable<DocxRevisionInfo> EnumerateTableRevisions(DocxTable table)
    {
        foreach (DocxRevisionInfo revision in table.Revisions)
        {
            yield return revision;
        }

        foreach (DocxTableRow row in table.Rows)
        {
            foreach (DocxRevisionInfo revision in row.Revisions)
            {
                yield return revision;
            }

            foreach (DocxTableCell cell in row.Cells)
            {
                foreach (DocxRevisionInfo revision in EnumerateTableCellRevisions(cell))
                {
                    yield return revision;
                }
            }
        }

    }

    private static IEnumerable<DocxRevisionInfo> EnumerateTableCellRevisions(DocxTableCell cell)
    {
        foreach (DocxRevisionInfo revision in cell.Revisions)
        {
            yield return revision;
        }

        foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(DocxTableCellContent.GetBodyElements(cell)))
        {
            foreach (DocxRevisionInfo revision in paragraph.Revisions)
            {
                yield return revision;
            }
        }
    }

    // Single caller; kept static: future Phase-1 split unit, not a local candidate.
    private static DocxStructureTableAdjacencySnapshot ToTableAdjacencySnapshot(
        IReadOnlyList<DocxBodyElement> elements,
        DocxTable table,
        int tableIndex,
        int blockIndex,
        string? previousKind,
        string? nextKind)
    {
            DocxParagraph? TryGetAdjacentParagraph(IReadOnlyList<DocxBodyElement> elements, int index)
            {
                return index >= 0 &&
                    index < elements.Count &&
                    elements[index] is DocxParagraphElement paragraph
                        ? paragraph.Paragraph
                        : null;
            }

        DocxParagraph? previousParagraph = TryGetAdjacentParagraph(elements, blockIndex - 1);
        DocxParagraph? nextParagraph = TryGetAdjacentParagraph(elements, blockIndex + 1);
        return new DocxStructureTableAdjacencySnapshot(
            tableIndex,
            blockIndex,
            previousKind,
            nextKind,
            table.Rows.Count,
            MaxColumnCount(table),
            previousParagraph?.EffectiveProperties.StyleId,
            previousParagraph?.EffectiveProperties.SpacingAfterPoints,
            previousParagraph is null ? null : HasAfterSpacingToken(previousParagraph.EffectiveProperties.Spacing),
            nextParagraph?.EffectiveProperties.StyleId,
            nextParagraph?.EffectiveProperties.SpacingBeforePoints,
            nextParagraph is null ? null : HasBeforeSpacingToken(nextParagraph.EffectiveProperties.Spacing),
            nextParagraph is null ? null : TextLength(nextParagraph),
            nextParagraph?.ListLabel is not null,
            nextParagraph?.EffectiveProperties.KeepRules.KeepNext,
            nextParagraph?.EffectiveProperties.KeepRules.KeepLines);
    }
}
