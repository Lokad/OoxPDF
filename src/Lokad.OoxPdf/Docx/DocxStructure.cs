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
                group.Count(paragraph => DocxParagraphSpacing.HasBeforeSpacingSide(paragraph.EffectiveProperties.Spacing)),
                group.Count(paragraph => DocxParagraphSpacing.HasAfterSpacingSide(paragraph.EffectiveProperties.Spacing)),
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

}
