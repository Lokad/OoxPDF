using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxStructureSnapshot(
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
    int InlineImageCount,
    int InlineReferenceCount,
    int AnchoredInlineReferenceCount,
    int ResolvedInlineReferenceCount,
    int MaxInlineReferenceTextOffsetInRun,
    int FieldReferenceCount,
    int PageFieldReferenceCount,
    int NumPagesFieldReferenceCount,
    int OtherFieldReferenceCount,
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
    IReadOnlyList<DocxStructureTableAdjacencySnapshot> TableAdjacency)
{
    public static DocxStructureSnapshot FromDocument(DocxDocument document)
    {
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
                    blocks.Add(new DocxStructureBlockSnapshot(blockIndex, "Unknown", previousKind, nextKind));
                    break;
            }
        }

        DocxParagraph[] allParagraphs = EnumerateParagraphs(document).ToArray();
        return new DocxStructureSnapshot(
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
            blocks.Sum(block => block.InlineImageCount),
            blocks.Sum(block => block.InlineReferenceCount),
            blocks.Sum(block => block.AnchoredInlineReferenceCount),
            blocks.Sum(block => block.ResolvedInlineReferenceCount),
            blocks.Select(block => block.MaxInlineReferenceTextOffsetInRun).DefaultIfEmpty(0).Max(),
            allParagraphs.Sum(ParagraphFieldReferenceCount),
            allParagraphs.Sum(ParagraphPageFieldReferenceCount),
            allParagraphs.Sum(ParagraphNumPagesFieldReferenceCount),
            allParagraphs.Sum(ParagraphOtherFieldReferenceCount),
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
            adjacency);
    }

    private static bool IsContinuousSectionBreak(string? typeValue)
    {
        return string.Equals(typeValue, "continuous", StringComparison.OrdinalIgnoreCase);
    }

    private static bool StartsNewPageSectionBreak(string? typeValue)
    {
        return typeValue is null
            || string.Equals(typeValue, "nextPage", StringComparison.OrdinalIgnoreCase)
            || string.Equals(typeValue, "oddPage", StringComparison.OrdinalIgnoreCase)
            || string.Equals(typeValue, "evenPage", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasSectionColumns(DocxStructureBlockSnapshot block)
    {
        return block.SectionColumnCountValue is not null
            || block.SectionColumnEqualWidthValue is not null
            || block.SectionColumnSpaceValue is not null
            || (block.SectionColumnDefinitionCount ?? 0) > 0;
    }

    private static IReadOnlyList<DocxStructureStorySnapshot> ToStorySnapshots(
        DocxDocument document,
        IReadOnlyList<DocxStructureBlockSnapshot> bodyBlocks)
    {
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
                bodyBlocks.Sum(block => block.InlineImageCount),
                bodyBlocks.Sum(block => block.InlineReferenceCount),
                bodyBlocks.Sum(block => block.ResolvedInlineReferenceCount),
                DocxBlockTraversal.EnumerateBodyParagraphs(document).Sum(ParagraphFieldReferenceCount),
                DocxBlockTraversal.EnumerateBodyParagraphs(document).Sum(ParagraphBookmarkAnchorCount),
                bodyBlocks.Sum(block => block.HyperlinkCount),
                bodyBlocks.Sum(block => block.ExternalHyperlinkCount),
                bodyBlocks.Sum(block => block.InternalHyperlinkCount),
                document.FloatingDrawings.Count)
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
                paragraphs.Sum(paragraph => paragraph.Images.Count),
                paragraphs.Sum(ParagraphInlineReferenceCount),
                paragraphs.Sum(paragraph => ParagraphResolvedInlineReferenceCount(paragraph, relatedStories)),
                paragraphs.Sum(ParagraphFieldReferenceCount),
                paragraphs.Sum(ParagraphBookmarkAnchorCount),
                paragraphs.Sum(ParagraphHyperlinkCount),
                paragraphs.Sum(ParagraphExternalHyperlinkCount),
                paragraphs.Sum(ParagraphInternalHyperlinkCount),
                drawings.Count));
        }
    }

    private static void AddRelatedStories(List<DocxStructureStorySnapshot> stories, IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        foreach (DocxRelatedStory story in relatedStories
            .OrderBy(story => story.Kind, StringComparer.Ordinal)
            .ThenBy(story => story.PartName, StringComparer.Ordinal)
            .ThenBy(story => story.Id, StringComparer.Ordinal))
        {
            DocxParagraph[] paragraphs = DocxBlockTraversal.EnumerateBodyParagraphs(story).ToArray();
            DocxTable[] tables = DocxBlockTraversal.EnumerateBodyTables(story).ToArray();
            int directParagraphCount = story.BodyElements.OfType<DocxParagraphElement>().Count();
            stories.Add(new DocxStructureStorySnapshot(
                story.Kind,
                story.PartName,
                null,
                story.Id,
                story.BodyElements.Count,
                directParagraphCount,
                tables.Length,
                paragraphs.Sum(TextLength),
                paragraphs.Sum(paragraph => paragraph.Images.Count),
                paragraphs.Sum(ParagraphInlineReferenceCount),
                paragraphs.Sum(paragraph => ParagraphResolvedInlineReferenceCount(paragraph, relatedStories)),
                paragraphs.Sum(ParagraphFieldReferenceCount),
                paragraphs.Sum(ParagraphBookmarkAnchorCount),
                paragraphs.Sum(ParagraphHyperlinkCount),
                paragraphs.Sum(ParagraphExternalHyperlinkCount),
                paragraphs.Sum(ParagraphInternalHyperlinkCount),
                story.FloatingDrawings.Count));
        }
    }

    private static IReadOnlyList<DocxStructureInlineReferenceSnapshot> ToInlineReferenceSnapshots(DocxDocument document)
    {
        DocxRelatedStory[] relatedStories = document.RelatedStories
            .OrderBy(story => story.Kind, StringComparer.Ordinal)
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
                    reference.Kind,
                    reference.Id,
                    reference.CustomMarkFollowsValue,
                    reference.SourceRunIndex,
                    reference.RunChildIndex,
                    reference.TextOffsetInRun,
                    storyIndex,
                    story?.Kind,
                    story?.PartName,
                    story?.Id,
                    story?.BodyElements.Count,
                    story is null ? null : DocxBlockTraversal.EnumerateBodyParagraphs(story).Sum(TextLength)));
            }
        }

        return references;
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

    private static (int? Index, DocxRelatedStory? Story) ResolveInlineReferenceStory(
        DocxInlineReference reference,
        IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        if (reference.Id is null)
        {
            return (null, null);
        }

        for (int index = 0; index < relatedStories.Count; index++)
        {
            DocxRelatedStory story = relatedStories[index];
            if (string.Equals(story.Kind, reference.Kind, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(story.Id, reference.Id, StringComparison.Ordinal))
            {
                return (index, story);
            }
        }

        return (null, null);
    }

    private static DocxStructureBlockSnapshot FromParagraph(
        int blockIndex,
        string? previousKind,
        string? nextKind,
        DocxParagraph paragraph,
        IReadOnlyList<DocxRelatedStory> relatedStories)
    {
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
            ListFormatValue: paragraph.ListLabel?.FormatValue,
            InlineImageCount: paragraph.Images.Count,
            InlineReferenceCount: paragraph.InlineReferences.Count,
            AnchoredInlineReferenceCount: paragraph.InlineReferences.Count(HasInlineReferenceAnchor),
            ResolvedInlineReferenceCount: ParagraphResolvedInlineReferenceCount(paragraph, relatedStories),
            MaxInlineReferenceTextOffsetInRun: paragraph.InlineReferences.Select(reference => reference.TextOffsetInRun).DefaultIfEmpty(0).Max(),
            FieldReferenceCount: paragraph.FieldReferences.Count,
            PageFieldReferenceCount: paragraph.FieldReferences.Count(reference => reference.Kind == "Page"),
            NumPagesFieldReferenceCount: paragraph.FieldReferences.Count(reference => reference.Kind == "NumPages"),
            OtherFieldReferenceCount: paragraph.FieldReferences.Count(reference => reference.Kind == "Other"),
            BookmarkAnchorCount: paragraph.BookmarkAnchors.Count,
            CommentReferenceCount: paragraph.InlineReferences.Count(reference => reference.Kind == "Comment"),
            FootnoteReferenceCount: paragraph.InlineReferences.Count(reference => reference.Kind == "Footnote"),
            EndnoteReferenceCount: paragraph.InlineReferences.Count(reference => reference.Kind == "Endnote"),
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
            SnapToGridValue: effective.SnapToGridValue);
    }

    private static DocxStructureBlockSnapshot FromTable(
        int blockIndex,
        string? previousKind,
        string? nextKind,
        DocxTable table,
        int tableIndex,
        IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        DocxParagraph[] paragraphs = DocxBlockTraversal.EnumerateTableParagraphs(table).ToArray();
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "Table",
            previousKind,
            nextKind,
            InlineReferenceCount: paragraphs.Sum(ParagraphInlineReferenceCount),
            AnchoredInlineReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(HasInlineReferenceAnchor)),
            ResolvedInlineReferenceCount: paragraphs.Sum(paragraph => ParagraphResolvedInlineReferenceCount(paragraph, relatedStories)),
            MaxInlineReferenceTextOffsetInRun: paragraphs.SelectMany(paragraph => paragraph.InlineReferences).Select(reference => reference.TextOffsetInRun).DefaultIfEmpty(0).Max(),
            CommentReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == "Comment")),
            FootnoteReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == "Footnote")),
            EndnoteReferenceCount: paragraphs.Sum(paragraph => paragraph.InlineReferences.Count(reference => reference.Kind == "Endnote")),
            FieldReferenceCount: paragraphs.Sum(ParagraphFieldReferenceCount),
            PageFieldReferenceCount: paragraphs.Sum(ParagraphPageFieldReferenceCount),
            NumPagesFieldReferenceCount: paragraphs.Sum(ParagraphNumPagesFieldReferenceCount),
            OtherFieldReferenceCount: paragraphs.Sum(ParagraphOtherFieldReferenceCount),
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
            TableLayoutValue: table.LayoutValue);
    }

    private static DocxStructureBlockSnapshot FromPageBreak(int blockIndex, string? previousKind, string? nextKind, DocxPageBreakElement pageBreak)
    {
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "PageBreak",
            previousKind,
            nextKind,
            PageBreakSourceKind: pageBreak.SourceKind,
            PageBreakValue: pageBreak.Value,
            PageBreakConsumesParagraphLine: pageBreak.BreakParagraph is not null,
            PageBreakLineSpacingPoints: pageBreak.BreakParagraph?.EffectiveProperties.LineSpacingPoints,
            PageBreakLineSpacingFactor: pageBreak.BreakParagraph?.EffectiveProperties.LineSpacingFactor);
    }

    private static DocxStructureBlockSnapshot FromManualBreak(int blockIndex, string? previousKind, string? nextKind, DocxManualBreakElement manualBreak)
    {
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "ManualBreak",
            previousKind,
            nextKind,
            ManualBreakSourceKind: manualBreak.SourceKind,
            ManualBreakValue: manualBreak.Value,
            ManualBreakConsumesParagraphLine: manualBreak.BreakParagraph is not null,
            ManualBreakLineSpacingPoints: manualBreak.BreakParagraph?.EffectiveProperties.LineSpacingPoints,
            ManualBreakLineSpacingFactor: manualBreak.BreakParagraph?.EffectiveProperties.LineSpacingFactor);
    }

    private static DocxStructureBlockSnapshot FromSectionBreak(int blockIndex, string? previousKind, string? nextKind, DocxSectionBreakElement sectionBreak)
    {
        return new DocxStructureBlockSnapshot(
            blockIndex,
            "SectionBreak",
            previousKind,
            nextKind,
            SectionBreakTypeValue: sectionBreak.TypeValue,
            SectionColumnCountValue: sectionBreak.ColumnCountValue,
            SectionColumnEqualWidthValue: sectionBreak.ColumnEqualWidthValue,
            SectionColumnSpaceValue: sectionBreak.ColumnSpaceValue,
            SectionColumnDefinitionCount: sectionBreak.ColumnDefinitions.Count,
            SectionColumnDefinitionWidthTokenCount: sectionBreak.ColumnDefinitions.Count(column => column.WidthValue is not null),
            SectionColumnDefinitionSpaceTokenCount: sectionBreak.ColumnDefinitions.Count(column => column.SpaceValue is not null));
    }

    private static DocxStructureTableSnapshot ToTableSnapshot(DocxTable table, int tableIndex, int blockIndex)
    {
        int cellCount = table.Rows.Sum(row => row.Cells.Count);
        DocxParagraph[] tableParagraphs = DocxBlockTraversal.EnumerateTableParagraphs(table).ToArray();
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
            table.Rows.Sum(row => row.Cells.Sum(cell => cell.Borders.Count(border => !string.Equals(border.Value, "nil", StringComparison.OrdinalIgnoreCase) && !string.Equals(border.Value, "none", StringComparison.OrdinalIgnoreCase)))),
            paragraphCount,
            tableParagraphs.Sum(paragraph => paragraph.Runs.Count),
            tableParagraphs.Sum(paragraph => TextLength(paragraph)),
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

    private static DocxStructureTableRowSnapshot ToTableRowSnapshot(DocxTableRow row, int rowIndex)
    {
        DocxStructureTableCellSnapshot[] cells = row.Cells.Select(ToTableCellSnapshot).ToArray();
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

    private static DocxStructureTableCellSnapshot ToTableCellSnapshot(DocxTableCell cell, int cellIndex)
    {
        DocxParagraph[] paragraphs = DocxBlockTraversal
            .EnumerateBodyParagraphs(DocxTableCellContent.GetBodyElements(cell))
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
            paragraphs.Sum(paragraph => CountWhitespaceDelimitedTokens(paragraph)),
            paragraphs.Select(paragraph => LongestWhitespaceDelimitedTokenLength(paragraph)).DefaultIfEmpty(0).Max(),
            paragraphs.Sum(paragraph => paragraph.Images.Count),
            paragraphs.Sum(ParagraphInlineReferenceCount),
            paragraphs.Sum(ParagraphHyperlinkCount),
            paragraphs.Sum(ParagraphExternalHyperlinkCount),
            paragraphs.Sum(ParagraphInternalHyperlinkCount),
            paragraphs.Count(paragraph => paragraph.ListLabel is not null),
            paragraphs.Count(HasEffectiveKeepConstraint),
            paragraphs.Count(paragraph => HasBeforeSpacingToken(paragraph.Spacing)),
            paragraphs.Count(paragraph => HasAfterSpacingToken(paragraph.Spacing)),
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

    private static DocxStructureFloatingDrawingSnapshot ToFloatingDrawingSnapshot(DocxFloatingDrawing drawing, int index)
    {
        return new DocxStructureFloatingDrawingSnapshot(
            index,
            drawing.WrapKind,
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
            drawing.SourceBlockIndex);
    }

    private static bool HasEffectiveKeepConstraint(DocxParagraph paragraph)
    {
        DocxParagraphKeepRules keepRules = paragraph.EffectiveProperties.KeepRules;
        return keepRules.KeepNext == true || keepRules.KeepLines == true;
    }

    private static IReadOnlyList<DocxStructureStyleUsageSnapshot> ToStyleUsages(DocxDocument document)
    {
        DocxStructureStyleUsageSnapshot[] paragraphStyles = EnumerateParagraphs(document)
            .GroupBy(paragraph => paragraph.EffectiveProperties.StyleId, StringComparer.Ordinal)
            .Select(group => new DocxStructureStyleUsageSnapshot(
                "Paragraph",
                group.Key,
                group.Count(),
                group.Count(),
                0,
                group.Sum(TextLength)))
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
                group.Sum(table => DocxBlockTraversal.EnumerateTableParagraphs(table).Sum(paragraph => TextLength(paragraph)))))
            .ToArray();
        return paragraphStyles.Concat(tableStyles).ToArray();
    }

    private static IReadOnlyList<DocxStructureListUsageSnapshot> ToListUsages(DocxDocument document)
    {
        return EnumerateParagraphs(document)
            .Where(paragraph => paragraph.ListLabel is not null)
            .GroupBy(paragraph => new
            {
                paragraph.ListLabel!.NumberId,
                paragraph.ListLabel.Level,
                paragraph.ListLabel.FormatValue,
                paragraph.ListLabel.SuffixValue
            })
            .Select(group => new DocxStructureListUsageSnapshot(
                group.Key.NumberId,
                group.Key.Level,
                group.Key.FormatValue,
                group.Key.SuffixValue,
                group.Count(),
                group.Sum(TextLength)))
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

    private static int ParagraphInlineReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.InlineReferences.Count;
    }

    private static int ParagraphResolvedInlineReferenceCount(DocxParagraph paragraph, IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        return paragraph.InlineReferences.Count(reference => InlineReferenceResolves(reference, relatedStories));
    }

    private static int ParagraphFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count;
    }

    private static int ParagraphPageFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.Kind == "Page");
    }

    private static int ParagraphNumPagesFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.Kind == "NumPages");
    }

    private static int ParagraphOtherFieldReferenceCount(DocxParagraph paragraph)
    {
        return paragraph.FieldReferences.Count(reference => reference.Kind == "Other");
    }

    private static int ParagraphBookmarkAnchorCount(DocxParagraph paragraph)
    {
        return paragraph.BookmarkAnchors.Count;
    }

    private static int ParagraphHyperlinkCount(DocxParagraph paragraph)
    {
        return paragraph.Hyperlinks.Count;
    }

    private static int ParagraphExternalHyperlinkCount(DocxParagraph paragraph)
    {
        return paragraph.Hyperlinks.Count(link => string.Equals(link.TargetMode, "External", StringComparison.OrdinalIgnoreCase));
    }

    private static int ParagraphInternalHyperlinkCount(DocxParagraph paragraph)
    {
        return paragraph.Hyperlinks.Count(link => link.Anchor is not null || link.ResolvedTarget is not null);
    }

    private static bool HasInlineReferenceAnchor(DocxInlineReference reference)
    {
        return reference.SourceRunIndex >= 0 && reference.RunChildIndex >= 0;
    }

    private static bool InlineReferenceResolves(DocxInlineReference reference, IReadOnlyList<DocxRelatedStory> relatedStories)
    {
        return reference.Id is not null &&
            relatedStories.Any(story =>
                string.Equals(story.Kind, reference.Kind, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(story.Id, reference.Id, StringComparison.Ordinal));
    }

    private static DocxStructureTableAdjacencySnapshot ToTableAdjacencySnapshot(
        IReadOnlyList<DocxBodyElement> elements,
        DocxTable table,
        int tableIndex,
        int blockIndex,
        string? previousKind,
        string? nextKind)
    {
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

    private static string GetBlockKind(DocxBodyElement element)
    {
        return element switch
        {
            DocxParagraphElement => "Paragraph",
            DocxTableElement => "Table",
            DocxPageBreakElement => "PageBreak",
            DocxManualBreakElement => "ManualBreak",
            DocxSectionBreakElement => "SectionBreak",
            _ => "Unknown"
        };
    }

    private static DocxParagraph? TryGetAdjacentParagraph(IReadOnlyList<DocxBodyElement> elements, int index)
    {
        return index >= 0 &&
            index < elements.Count &&
            elements[index] is DocxParagraphElement paragraph
                ? paragraph.Paragraph
                : null;
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

    private static int TextLength(DocxParagraph paragraph)
    {
        return paragraph.Runs.Sum(run => run.Text.Length);
    }

    private static int CountCharacters(DocxParagraph paragraph, Func<char, bool> predicate)
    {
        return paragraph.Runs.Sum(run => run.Text.Count(predicate));
    }

    private static int CountWhitespaceDelimitedTokens(DocxParagraph paragraph)
    {
        return EnumerateWhitespaceDelimitedTokenLengths(paragraph).Count();
    }

    private static int LongestWhitespaceDelimitedTokenLength(DocxParagraph paragraph)
    {
        return EnumerateWhitespaceDelimitedTokenLengths(paragraph).DefaultIfEmpty(0).Max();
    }

    private static IEnumerable<int> EnumerateWhitespaceDelimitedTokenLengths(DocxParagraph paragraph)
    {
        int currentLength = 0;
        foreach (DocxTextRun run in paragraph.Runs)
        {
            foreach (char c in run.Text)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (currentLength > 0)
                    {
                        yield return currentLength;
                        currentLength = 0;
                    }

                    continue;
                }

                currentLength++;
            }
        }

        if (currentLength > 0)
        {
            yield return currentLength;
        }
    }

    private static int CountCharacters(IEnumerable<DocxParagraph> paragraphs, Func<char, bool> predicate)
    {
        return paragraphs.Sum(paragraph => CountCharacters(paragraph, predicate));
    }

    private static bool HasVisibleText(DocxParagraph paragraph)
    {
        return paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text));
    }

    private static int MaxColumnCount(DocxTable table)
    {
        return table.Rows.Select(row => row.Cells.Sum(cell => Math.Max(1, cell.GridSpan))).DefaultIfEmpty(0).Max();
    }
}

internal sealed record DocxStructureBlockSnapshot(
    int Index,
    string Kind,
    string? PreviousKind,
    string? NextKind,
    string? ParagraphStyleId = null,
    bool? ParagraphStyleFound = null,
    int? ParagraphStyleDepth = null,
    bool? HasDocumentDefaultParagraphProperties = null,
    bool? HasDirectParagraphProperties = null,
    bool? HasTableStyleParagraphProperties = null,
    int RunCount = 0,
    int TextLength = 0,
    bool HasVisibleText = false,
    string? ListFormatValue = null,
    int InlineImageCount = 0,
    int InlineReferenceCount = 0,
    int AnchoredInlineReferenceCount = 0,
    int ResolvedInlineReferenceCount = 0,
    int MaxInlineReferenceTextOffsetInRun = 0,
    int CommentReferenceCount = 0,
    int FootnoteReferenceCount = 0,
    int EndnoteReferenceCount = 0,
    int FieldReferenceCount = 0,
    int PageFieldReferenceCount = 0,
    int NumPagesFieldReferenceCount = 0,
    int OtherFieldReferenceCount = 0,
    int BookmarkAnchorCount = 0,
    int HyperlinkCount = 0,
    int ExternalHyperlinkCount = 0,
    int InternalHyperlinkCount = 0,
    int TableRowCount = 0,
    int? TableIndex = null,
    double? SpacingBeforePoints = null,
    double? SpacingAfterPoints = null,
    double? LineSpacingPoints = null,
    double? LineSpacingFactor = null,
    bool HasBeforeSpacingToken = false,
    bool HasAfterSpacingToken = false,
    string? BeforeAutoSpacingValue = null,
    string? AfterAutoSpacingValue = null,
    bool? ContextualSpacing = null,
    bool? KeepNext = null,
    bool? KeepLines = null,
    bool? WidowControl = null,
    int? TableMaxColumnCount = null,
    double? TablePreferredWidthPoints = null,
    string? TablePreferredWidthType = null,
    double? TableIndentPoints = null,
    double? TableCellSpacingPoints = null,
    string? TableLayoutValue = null,
    string? PageBreakSourceKind = null,
    string? PageBreakValue = null,
    string? ManualBreakSourceKind = null,
    string? ManualBreakValue = null,
    string? SectionBreakTypeValue = null,
    string? SectionColumnCountValue = null,
    string? SectionColumnEqualWidthValue = null,
    string? SectionColumnSpaceValue = null,
    int? SectionColumnDefinitionCount = null,
    int? SectionColumnDefinitionWidthTokenCount = null,
    int? SectionColumnDefinitionSpaceTokenCount = null,
    double? ParagraphIndentLeftPoints = null,
    double? ParagraphIndentRightPoints = null,
    double? ParagraphIndentFirstLinePoints = null,
    double? ParagraphIndentHangingPoints = null,
    int? WhitespaceDelimitedTokenCount = null,
    int? LongestWhitespaceDelimitedTokenLength = null,
    int? SpaceCharacterCount = null,
    int? NonAsciiCharacterCount = null,
    int? PunctuationCharacterCount = null,
    int? DigitCharacterCount = null,
    int? UppercaseCharacterCount = null,
    int? LowercaseCharacterCount = null,
    int? TabStopCount = null,
    bool? SnapToGrid = null,
    string? SnapToGridValue = null,
    bool? PageBreakConsumesParagraphLine = null,
    double? PageBreakLineSpacingPoints = null,
    double? PageBreakLineSpacingFactor = null,
    bool? ManualBreakConsumesParagraphLine = null,
    double? ManualBreakLineSpacingPoints = null,
    double? ManualBreakLineSpacingFactor = null);

internal sealed record DocxStructureStorySnapshot(
    string Kind,
    string Scope,
    int? SectionBreakBlockIndex,
    string? VariantType,
    int BlockCount,
    int ParagraphCount,
    int TableCount,
    int TextLength,
    int InlineImageCount,
    int InlineReferenceCount,
    int ResolvedInlineReferenceCount,
    int FieldReferenceCount,
    int BookmarkAnchorCount,
    int HyperlinkCount,
    int ExternalHyperlinkCount,
    int InternalHyperlinkCount,
    int FloatingDrawingCount);

internal sealed record DocxStructureInlineReferenceSnapshot(
    int SourceBlockIndex,
    string SourceBlockKind,
    int SourceParagraphIndex,
    string Kind,
    string? Id,
    string? CustomMarkFollowsValue,
    int SourceRunIndex,
    int RunChildIndex,
    int TextOffsetInRun,
    int? ResolvedStoryIndex,
    string? ResolvedStoryKind,
    string? ResolvedStoryPartName,
    string? ResolvedStoryId,
    int? ResolvedStoryBlockCount,
    int? ResolvedStoryTextLength);

internal sealed record DocxStructureFloatingDrawingSnapshot(
    int Index,
    string? WrapKind,
    string? WrapTextValue,
    string? BehindDocumentValue,
    string? LayoutInCellValue,
    string? AllowOverlapValue,
    string? HorizontalRelativeFromValue,
    string? HorizontalAlignValue,
    string? HorizontalOffsetValue,
    string? VerticalRelativeFromValue,
    string? VerticalAlignValue,
    string? VerticalOffsetValue,
    string? ExtentCxValue,
    string? ExtentCyValue,
    string? DistanceTopValue,
    string? DistanceBottomValue,
    string? DistanceLeftValue,
    string? DistanceRightValue,
    string? SimplePositionValue,
    string? RelativeHeightValue,
    string? LockedValue,
    string? ImageRelationshipId,
    string? ImagePartName,
    string? ImageContentType,
    double? ImageWidthPoints,
    double? ImageHeightPoints,
    int? SourceParagraphIndex,
    int? SourceBlockIndex);

internal sealed record DocxStructureStyleUsageSnapshot(
    string Kind,
    string? StyleId,
    int Count,
    int ParagraphCount,
    int TableCount,
    int TextLength);

internal sealed record DocxStructureListUsageSnapshot(
    string NumberId,
    int Level,
    string FormatValue,
    string SuffixValue,
    int ParagraphCount,
    int TextLength);

internal sealed record DocxStructureTableSnapshot(
    int TableIndex,
    int BlockIndex,
    string? StyleId,
    int RowCount,
    int MaxColumnCount,
    int GridColumnCount,
    double GridColumnsWidthSum,
    bool HasExplicitGrid,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    double? IndentPoints,
    string? IndentValue,
    string? IndentType,
    double? CellSpacingPoints,
    string? CellSpacingValue,
    string? CellSpacingType,
    string? LayoutValue,
    int HeaderRowCount,
    int CantSplitRowCount,
    int DeclaredHeightRowCount,
    int ExactHeightRowCount,
    int AtLeastHeightRowCount,
    int RowPropertyExceptionCount,
    int CellCount,
    int GridSpanCellCount,
    int VerticalMergeCellCount,
    int VerticalMergeRestartCellCount,
    int ShadedCellCount,
    int VerticalAlignmentCellCount,
    int PreferredWidthCellCount,
    int VisibleBorderCount,
    int ParagraphCount,
    int RunCount,
    int TextLength,
    int WhitespaceDelimitedTokenCount,
    int LongestWhitespaceDelimitedTokenLength,
    int InlineImageCount,
    int InlineReferenceCount,
    int HyperlinkCount,
    int ExternalHyperlinkCount,
    int InternalHyperlinkCount,
    int NumberedParagraphCount,
    int KeepRuleParagraphCount,
    bool? LookFirstRow,
    bool? LookFirstColumn,
    bool? LookNoHorizontalBand,
    bool? LookNoVerticalBand,
    IReadOnlyList<DocxStructureTableRowSnapshot> Rows);

internal sealed record DocxStructureTableRowSnapshot(
    int RowIndex,
    int CellCount,
    int LogicalGridSpan,
    bool IsHeader,
    string? HeaderValue,
    bool CantSplit,
    string? CantSplitValue,
    double? HeightPoints,
    string? HeightValue,
    string? HeightRuleValue,
    bool HasTablePropertyExceptionCellMargins,
    int GridSpanCellCount,
    int VerticalMergeCellCount,
    int VerticalMergeRestartCellCount,
    int ShadedCellCount,
    int VerticalAlignmentCellCount,
    int PreferredWidthCellCount,
    int VisibleBorderCount,
    int ParagraphCount,
    int RunCount,
    int TextLength,
    int WhitespaceDelimitedTokenCount,
    int LongestWhitespaceDelimitedTokenLength,
    int InlineImageCount,
    int InlineReferenceCount,
    int HyperlinkCount,
    int ExternalHyperlinkCount,
    int InternalHyperlinkCount,
    int NumberedParagraphCount,
    int KeepRuleParagraphCount,
    int BeforeSpacingTokenParagraphCount,
    int AfterSpacingTokenParagraphCount,
    double MaxFontSize,
    IReadOnlyList<DocxStructureTableCellSnapshot> Cells);

internal sealed record DocxStructureTableCellSnapshot(
    int CellIndex,
    int GridSpan,
    string? GridSpanValue,
    bool HasVerticalMerge,
    string? VerticalMergeValue,
    bool HasShading,
    string? ShadingValue,
    string? VerticalAlignmentValue,
    bool HasPreferredWidth,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    int VisibleBorderCount,
    int ParagraphCount,
    int RunCount,
    int TextLength,
    int WhitespaceDelimitedTokenCount,
    int LongestWhitespaceDelimitedTokenLength,
    int InlineImageCount,
    int InlineReferenceCount,
    int HyperlinkCount,
    int ExternalHyperlinkCount,
    int InternalHyperlinkCount,
    int NumberedParagraphCount,
    int KeepRuleParagraphCount,
    int BeforeSpacingTokenParagraphCount,
    int AfterSpacingTokenParagraphCount,
    double MaxFontSize,
    int SpaceCharacterCount,
    int NonAsciiCharacterCount,
    int PunctuationCharacterCount,
    int DigitCharacterCount,
    int UppercaseCharacterCount,
    int LowercaseCharacterCount,
    int BodyElementCount = 0,
    int ManualBreakElementCount = 0,
    int PageBreakElementCount = 0,
    int NestedTableElementCount = 0);

internal sealed record DocxStructureTableAdjacencySnapshot(
    int TableIndex,
    int BlockIndex,
    string? PreviousKind,
    string? NextKind,
    int RowCount,
    int MaxColumnCount,
    string? PreviousParagraphStyleId,
    double? PreviousParagraphSpacingAfterPoints,
    bool? PreviousParagraphHasAfterSpacingToken,
    string? NextParagraphStyleId,
    double? NextParagraphSpacingBeforePoints,
    bool? NextParagraphHasBeforeSpacingToken,
    int? NextParagraphTextLength,
    bool? NextParagraphHasListLabel,
    bool? NextParagraphKeepNext,
    bool? NextParagraphKeepLines);
