namespace Lokad.OoxPdf.Docx;

internal sealed record DocxDocument(
    double PageWidthPoints,
    double PageHeightPoints,
    double MarginLeftPoints,
    double MarginRightPoints,
    double MarginTopPoints,
    double MarginBottomPoints,
    DocxPageSettings PageSettings,
    IReadOnlyList<DocxFloatingDrawing> FloatingDrawings,
    IReadOnlyList<DocxParagraph> HeaderParagraphs,
    IReadOnlyList<DocxParagraph> FooterParagraphs,
    IReadOnlyList<DocxBodyElement> BodyElements,
    IReadOnlyList<DocxParagraph> FallbackParagraphs,
    IReadOnlyList<DocxTable> FallbackTables)
{
    public DocxFontCatalog FontCatalog { get; init; } = DocxFontCatalog.Empty;
    public DocxStyleCatalog StyleCatalog { get; init; } = DocxStyleCatalog.Empty;
    public IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> HeaderParagraphsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> FooterParagraphsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> HeaderBodyElementsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> FooterBodyElementsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> HeaderFloatingDrawingsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> FooterFloatingDrawingsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);
    public IReadOnlyList<DocxRelatedStory> RelatedStories { get; init; } = [];
    public IReadOnlyList<string> PackageCommentAnchorIds { get; init; } = [];
    public IReadOnlyList<string> HiddenCommentAnchorIds { get; init; } = [];
    public DocxDocumentSettings Settings { get; init; } = DocxDocumentSettings.Empty;
    public OoxPdfDocxMarkupMode MarkupMode { get; init; } = OoxPdfDocxMarkupMode.Final;
    public DocxSectionBreakElement? FinalSectionBreak { get; init; }
    public IReadOnlyList<DocxParagraph> Paragraphs => BodyElements.Count == 0
        ? FallbackParagraphs
        : DocxBlockTraversal.EnumerateDirectParagraphs(BodyElements).ToArray();
    public IReadOnlyList<DocxTable> Tables => BodyElements.Count == 0
        ? FallbackTables
        : DocxBlockTraversal.EnumerateBodyTables(BodyElements).ToArray();

    public DocxDocument(double pageWidthPoints, double pageHeightPoints)
        : this(pageWidthPoints, pageHeightPoints, 72d, 72d, 72d, 72d, DocxPageSettings.Empty, [], [], [], [], [], [])
    {
    }
}

internal sealed record DocxRelatedStory(
    DocxRelatedStoryKind Kind,
    string PartName,
    string? Id,
    IReadOnlyList<DocxBodyElement> BodyElements,
    IReadOnlyList<DocxParagraph> FallbackParagraphs,
    IReadOnlyList<DocxTable> FallbackTables,
    DocxRelatedStoryType? Type)
{
    public IReadOnlyList<DocxFloatingDrawing> FloatingDrawings { get; init; } = [];
    public bool IsNormalStoryType => Type is null || Type == DocxRelatedStoryType.Normal;
    public DocxCommentMetadata? CommentMetadata { get; init; }
    public IReadOnlyList<DocxParagraph> Paragraphs => BodyElements.Count == 0
        ? FallbackParagraphs
        : DocxBlockTraversal.EnumerateDirectParagraphs(BodyElements).ToArray();
    public IReadOnlyList<DocxTable> Tables => BodyElements.Count == 0
        ? FallbackTables
        : DocxBlockTraversal.EnumerateBodyTables(BodyElements).ToArray();
}

internal sealed record DocxCommentMetadata(
    string? Author,
    string? Initials,
    string? Date,
    string? ParagraphId,
    string? ParentParagraphId,
    string? ParentCommentId,
    bool? IsResolved);

internal static class DocxBlockTraversal
{
    public static IEnumerable<DocxParagraph> EnumerateBodyParagraphs(DocxDocument document)
    {
        return EnumerateBodyParagraphs(document.BodyElements);
    }

    public static IEnumerable<DocxParagraph> EnumerateBodyParagraphs(DocxRelatedStory story)
    {
        return EnumerateBodyParagraphs(story.BodyElements);
    }

    public static IEnumerable<DocxParagraph> EnumerateStaticStoryParagraphs(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> fallbackParagraphsByType)
    {
        return bodyElementsByType.Count == 0
            ? fallbackParagraphsByType.Values.SelectMany(paragraphs => paragraphs)
            : bodyElementsByType.Values.SelectMany(EnumerateBodyParagraphs);
    }

    public static IEnumerable<DocxParagraph> EnumerateStaticStoryParagraphs(DocxPageSettings settings)
    {
        return EnumerateStaticStoryParagraphs(settings.HeaderBodyElementsByType, settings.HeaderParagraphsByType)
            .Concat(EnumerateStaticStoryParagraphs(settings.FooterBodyElementsByType, settings.FooterParagraphsByType));
    }

    public static IEnumerable<DocxParagraph> EnumerateReferencedStaticStoryParagraphs(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType,
        IReadOnlyList<DocxParagraph> fallbackParagraphs)
    {
        if (bodyElementsByType.Count != 0)
        {
            return bodyElementsByType.Values.SelectMany(EnumerateBodyParagraphs);
        }

        return paragraphsByType.Count == 0
            ? fallbackParagraphs
            : paragraphsByType.Values.SelectMany(paragraphs => paragraphs);
    }

    public static IReadOnlyList<DocxBodyElement> GetStaticStoryBodyElements(
        string variantType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> fallbackParagraphsByType)
    {
        if (bodyElementsByType.TryGetValue(variantType, out IReadOnlyList<DocxBodyElement>? bodyElements))
        {
            return bodyElements;
        }

        return fallbackParagraphsByType.TryGetValue(variantType, out IReadOnlyList<DocxParagraph>? paragraphs)
            ? paragraphs.Select(DocxBodyElementFactory.CreateParagraph).Cast<DocxBodyElement>().ToArray()
            : [];
    }

    public static bool TryGetStaticStoryBodyElements(
        string variantType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType,
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> fallbackParagraphsByType,
        out IReadOnlyList<DocxBodyElement> bodyElements)
    {
        if (bodyElementsByType.TryGetValue(variantType, out IReadOnlyList<DocxBodyElement>? foundElements))
        {
            bodyElements = foundElements;
            return true;
        }

        if (fallbackParagraphsByType.TryGetValue(variantType, out IReadOnlyList<DocxParagraph>? paragraphs))
        {
            bodyElements = paragraphs.Select(DocxBodyElementFactory.CreateParagraph).Cast<DocxBodyElement>().ToArray();
            return true;
        }

        bodyElements = [];
        return false;
    }

    public static IEnumerable<DocxParagraph> EnumerateBodyParagraphs(IEnumerable<DocxBodyElement> bodyElements)
    {
        foreach (DocxBodyElement element in bodyElements)
        {
            switch (element)
            {
                case DocxParagraphElement paragraph:
                    yield return paragraph.Paragraph;
                    break;
                case DocxTableElement table:
                    foreach (DocxParagraph cellParagraph in EnumerateTableParagraphs(table.Table))
                    {
                        yield return cellParagraph;
                    }

                    break;
            }
        }
    }

    public static IEnumerable<DocxParagraph> EnumerateDirectParagraphs(IEnumerable<DocxBodyElement> bodyElements)
    {
        return bodyElements
            .OfType<DocxParagraphElement>()
            .Select(element => element.Paragraph);
    }

    public static IEnumerable<DocxTable> EnumerateBodyTables(DocxDocument document)
    {
        return EnumerateBodyTables(document.BodyElements);
    }

    public static IEnumerable<DocxTable> EnumerateBodyTables(DocxRelatedStory story)
    {
        return EnumerateBodyTables(story.BodyElements);
    }

    public static IEnumerable<DocxTable> EnumerateBodyTables(IEnumerable<DocxBodyElement> bodyElements)
    {
        foreach (DocxTableElement table in bodyElements.OfType<DocxTableElement>())
        {
            yield return table.Table;
            foreach (DocxTable nestedTable in EnumerateTableTables(table.Table))
            {
                yield return nestedTable;
            }
        }

        IEnumerable<DocxTable> EnumerateTableTables(DocxTable table)
        {
            return table.Rows
                .SelectMany(row => row.Cells)
                .SelectMany(cell => EnumerateBodyTables(DocxTableCellContent.GetBodyElements(cell)));
        }
    }

    public static IEnumerable<DocxParagraph> EnumerateTableParagraphs(DocxTable table)
    {
        return table.Rows
            .SelectMany(row => row.Cells)
            .SelectMany(cell => EnumerateBodyParagraphs(DocxTableCellContent.GetBodyElements(cell)));
    }
}

internal sealed record DocxDocumentSettings(
    string? CharacterSpacingControlValue,
    string? DefaultTabStopValue,
    double? DefaultTabStopPoints,
    bool? UseFELayout,
    string? UseFELayoutValue,
    DocxRevisionViewSettings RevisionViewSettings,
    DocxTrackChangesSettings TrackChangesSettings,
    DocxNoteReferenceSettings FootnoteReferenceSettings,
    DocxNoteReferenceSettings EndnoteReferenceSettings,
    string? MirrorMarginsValue,
    bool? MirrorMargins,
    IReadOnlyList<DocxCompatSetting> CompatSettings)
{
    public static DocxDocumentSettings Empty { get; } = new(null, null, null, null, null, DocxRevisionViewSettings.Empty, DocxTrackChangesSettings.Empty, DocxNoteReferenceSettings.Empty, DocxNoteReferenceSettings.Empty, null, null, []);
}

internal sealed record DocxRevisionViewSettings(
    string? MarkupValue,
    bool? ShowMarkup,
    string? CommentsValue,
    bool? ShowComments,
    string? InsertionsAndDeletionsValue,
    bool? ShowInsertionsAndDeletions,
    string? FormattingValue,
    bool? ShowFormatting,
    string? InkAnnotationsValue,
    bool? ShowInkAnnotations)
{
    public static DocxRevisionViewSettings Empty { get; } = new(null, null, null, null, null, null, null, null, null, null);
}

internal sealed record DocxTrackChangesSettings(
    string? TrackRevisionsValue,
    bool? TrackRevisions,
    string? DoNotTrackMovesValue,
    bool? DoNotTrackMoves,
    string? DoNotTrackFormattingValue,
    bool? DoNotTrackFormatting)
{
    public static DocxTrackChangesSettings Empty { get; } = new(null, null, null, null, null, null);
}

internal sealed record DocxNoteReferenceSettings(
    string? PositionValue,
    string? NumberFormatValue,
    string? NumberStartValue,
    int? NumberStart,
    string? NumberRestartValue)
{
    public static DocxNoteReferenceSettings Empty { get; } = new(null, null, null, null, null);
}

internal sealed record DocxCompatSetting(
    string? Name,
    string? Uri,
    string? Value);

internal sealed record DocxFontCatalog(
    IReadOnlyList<DocxFontTableEntry> Entries,
    DocxThemeFonts ThemeFonts)
{
    public static DocxFontCatalog Empty { get; } = new([], DocxThemeFonts.Empty);
}

internal sealed record DocxFontTableEntry(
    string Name,
    string? AlternateName,
    string? FamilyValue,
    string? PitchValue,
    string? PanoseValue,
    string? CharsetValue);

internal sealed record DocxThemeFonts(
    string? MajorLatinTypeface,
    string? MinorLatinTypeface,
    string? MajorComplexScriptTypeface,
    string? MinorComplexScriptTypeface,
    string? MajorEastAsiaTypeface,
    string? MinorEastAsiaTypeface)
{
    public static DocxThemeFonts Empty { get; } = new(null, null, null, null, null, null);
}

internal sealed record DocxStyleCatalog(
    bool HasRunDefaults,
    bool HasParagraphDefaults,
    string? DefaultTableStyleId,
    IReadOnlyList<DocxStyleDefinitionSummary> ParagraphStyles,
    IReadOnlyList<DocxStyleDefinitionSummary> CharacterStyles,
    IReadOnlyList<DocxTableStyleDefinitionSummary> TableStyles)
{
    public static DocxStyleCatalog Empty { get; } = new(false, false, null, [], [], []);
}

internal sealed record DocxStyleDefinitionSummary(
    string StyleId,
    string? BasedOnStyleId,
    bool HasParagraphProperties,
    bool HasRunProperties);

internal sealed record DocxTableStyleDefinitionSummary(
    string StyleId,
    string? BasedOnStyleId,
    bool HasTableProperties,
    bool HasCellProperties,
    bool HasParagraphProperties,
    bool HasRunProperties,
    int BorderCount,
    int ConditionalRegionCount);

internal sealed record DocxPageSettings(
    string? WidthValue,
    string? HeightValue,
    string? OrientationValue,
    string? MarginTopValue,
    string? MarginRightValue,
    string? MarginBottomValue,
    string? MarginLeftValue,
    double? HeaderDistancePoints,
    double? FooterDistancePoints,
    string? HeaderDistanceValue,
    string? FooterDistanceValue,
    bool? TitlePage,
    string? TitlePageValue,
    bool? EvenAndOddHeaders,
    string? EvenAndOddHeadersValue)
{
    public static DocxPageSettings Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

    public double? DocGridLinePitchPoints { get; init; }
    public string? DocGridLinePitchValue { get; init; }
    public double? GutterDistancePoints { get; init; }
    public string? GutterDistanceValue { get; init; }
    public DocxNoteReferenceSettings FootnoteReferenceSettings { get; init; } = DocxNoteReferenceSettings.Empty;
    public DocxNoteReferenceSettings EndnoteReferenceSettings { get; init; } = DocxNoteReferenceSettings.Empty;

    public IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> HeaderParagraphsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> FooterParagraphsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> HeaderBodyElementsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> FooterBodyElementsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> HeaderFloatingDrawingsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> FooterFloatingDrawingsByType { get; init; } =
        new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);
}

internal sealed record DocxFloatingDrawing(
    string? DistanceTopValue,
    string? DistanceBottomValue,
    string? DistanceLeftValue,
    string? DistanceRightValue,
    string? SimplePositionValue,
    string? RelativeHeightValue,
    string? BehindDocumentValue,
    string? LockedValue,
    string? LayoutInCellValue,
    string? AllowOverlapValue,
    string? ExtentCxValue,
    string? ExtentCyValue,
    string? HorizontalRelativeFromValue,
    string? HorizontalAlignValue,
    string? HorizontalOffsetValue,
    string? VerticalRelativeFromValue,
    string? VerticalAlignValue,
    string? VerticalOffsetValue,
    DocxFloatingWrapKind? WrapKind,
    string? WrapTextValue,
    string? ImageRelationshipId,
    DocxInlineImage? Image,
    int? SourceParagraphIndex,
    int? SourceBlockIndex)
{
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
    public IReadOnlyList<DocxBodyElement> TextBoxBodyElements { get; init; } = [];
}

// One block-level w:body child in document order. Implementors wrap a single
// source construct (paragraph, table, section break, page/manual break, or an
// implicit paragraph synthesized by the reader); the base Revisions list carries
// the tracked changes attached to that block, and layout consumes implementors
// in sequence without reordering.
internal abstract record DocxBodyElement
{
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal sealed record DocxParagraphElement(DocxParagraph Paragraph) : DocxBodyElement;

internal sealed record DocxTableElement(DocxTable Table) : DocxBodyElement;

internal sealed record DocxImplicitParagraphElement(DocxBreakSourceKind SourceKind) : DocxBodyElement
{
    public static DocxTextRun CreateParagraphMarkRun()
    {
        return new DocxTextRun(string.Empty, DocxDefaults.FontSizePoints, null, false, false, false, null, null);
    }
}

internal sealed record DocxPageBreakElement(DocxBreakSourceKind SourceKind, string? Value, DocxParagraph? BreakParagraph) : DocxBodyElement;

internal sealed record DocxManualBreakElement(DocxBreakSourceKind SourceKind, string? Value, DocxParagraph? BreakParagraph) : DocxBodyElement;

internal sealed record DocxSectionColumn(
    string? WidthValue,
    string? SpaceValue);

internal sealed record DocxSectionBreakElement(
    DocxPageSettings PageSettings,
    DocxSectionBreakType? TypeValue,
    string? ColumnCountValue,
    string? ColumnEqualWidthValue,
    string? ColumnSpaceValue,
    IReadOnlyList<DocxSectionColumn> ColumnDefinitions) : DocxBodyElement;

internal static class DocxBodyElementFactory
{
    public static DocxParagraphElement CreateParagraph(DocxParagraph paragraph)
    {
        return new DocxParagraphElement(paragraph)
        {
            Revisions = paragraph.Revisions
        };
    }

    public static DocxTableElement CreateTable(DocxTable table)
    {
        return new DocxTableElement(table)
        {
            Revisions = table.Revisions
        };
    }

    public static DocxPageBreakElement CreatePageBreak(
        DocxBreakSourceKind sourceKind,
        string? value,
        DocxParagraph? breakParagraph,
        IReadOnlyList<DocxRevisionInfo>? revisions)
    {
        return new DocxPageBreakElement(sourceKind, value, breakParagraph)
        {
            Revisions = revisions ?? breakParagraph?.Revisions ?? []
        };
    }

    public static DocxManualBreakElement CreateManualBreak(DocxBreakSourceKind sourceKind, string? value, DocxParagraph? breakParagraph)
    {
        return new DocxManualBreakElement(sourceKind, value, breakParagraph)
        {
            Revisions = breakParagraph?.Revisions ?? []
        };
    }
}

internal sealed record DocxParagraph(
    IReadOnlyList<DocxTextRun> Runs,
    IReadOnlyList<DocxInlineImage> Images,
    string? StyleId,
    DocxTextAlignment Alignment,
    string? AlignmentValue,
    double SpacingBeforePoints,
    double SpacingAfterPoints,
    double LineSpacingFactor,
    double? LineSpacingPoints,
    DocxParagraphSpacing Spacing,
    DocxParagraphKeepRules KeepRules,
    DocxListLabel? ListLabel)
{
    public DocxParagraphIndent Indent { get; init; } = DocxParagraphIndent.Empty;
    public IReadOnlyList<DocxTabStop> TabStops { get; init; } = [];
    public bool? SnapToGrid { get; init; }
    public string? SnapToGridValue { get; init; }
    public bool? WordWrap { get; init; }
    public string? WordWrapValue { get; init; }
    public DocxParagraphStyleResolution StyleResolution { get; init; } = DocxParagraphStyleResolution.Empty;
    public IReadOnlyList<DocxInlineReference> InlineReferences { get; init; } = [];
    public IReadOnlyList<DocxCommentRange> CommentRanges { get; init; } = [];
    public IReadOnlyList<DocxRevisionRange> RevisionRanges { get; init; } = [];
    public IReadOnlyList<DocxFieldReference> FieldReferences { get; init; } = [];
    public IReadOnlyList<DocxHyperlinkSpan> Hyperlinks { get; init; } = [];
    public IReadOnlyList<DocxBookmarkAnchor> BookmarkAnchors { get; init; } = [];
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
    public bool HasDeletedParagraphMark { get; init; }

    public DocxEffectiveParagraphProperties EffectiveProperties => new(
        StyleId,
        Alignment,
        AlignmentValue,
        SpacingBeforePoints,
        SpacingAfterPoints,
        LineSpacingFactor,
        LineSpacingPoints,
        Spacing,
        KeepRules,
        Indent,
        TabStops,
        SnapToGrid,
        SnapToGridValue,
        WordWrap,
        WordWrapValue,
        StyleResolution);
}

internal sealed record DocxEffectiveParagraphProperties(
    string? StyleId,
    DocxTextAlignment Alignment,
    string? AlignmentValue,
    double SpacingBeforePoints,
    double SpacingAfterPoints,
    double LineSpacingFactor,
    double? LineSpacingPoints,
    DocxParagraphSpacing Spacing,
    DocxParagraphKeepRules KeepRules,
    DocxParagraphIndent Indent,
    IReadOnlyList<DocxTabStop> TabStops,
    bool? SnapToGrid,
    string? SnapToGridValue,
    bool? WordWrap,
    string? WordWrapValue,
    DocxParagraphStyleResolution StyleResolution);

internal sealed record DocxParagraphStyleResolution(
    string? StyleId,
    bool StyleFound,
    int StyleDepth,
    bool HasDocumentDefaultParagraphProperties,
    bool HasDirectParagraphProperties,
    bool HasTableStyleParagraphProperties)
{
    public static DocxParagraphStyleResolution Empty { get; } = new(null, false, 0, false, false, false);
}

internal sealed record DocxBookmarkAnchor(
    string? Id,
    string? Name,
    int SourceRunIndex,
    int TextRunIndex,
    int TextOffset);

internal sealed record DocxInlineReference(
    DocxRelatedStoryKind Kind,
    string? Id,
    string? CustomMarkFollowsValue,
    string? DisplayText,
    int SourceRunIndex,
    int RunChildIndex,
    int TextOffsetInRun)
{
    public DocxRevisionInfo? Revision { get; init; }
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal sealed record DocxCommentRange(
    string? Id,
    int? StartSourceRunIndex,
    int? StartTextOffset,
    int? EndSourceRunIndex,
    int? EndTextOffset,
    int? ReferenceSourceRunIndex,
    int? ReferenceTextOffset);

internal sealed record DocxRevisionRange(
    DocxRevisionKind Kind,
    string? Id,
    string? Name,
    string? Author,
    string? Date,
    int? StartSourceRunIndex,
    int? StartTextOffset,
    int? EndSourceRunIndex,
    int? EndTextOffset);

internal sealed record DocxFieldReference(
    DocxFieldKind Kind,
    DocxFieldSourceKind SourceKind,
    string? Instruction,
    string? Placeholder,
    int SourceRunIndex,
    int TextRunIndex,
    int TextRunCount,
    int TextLength)
{
    public bool HasSeparate { get; init; }
    public bool HasCachedResult { get; init; }
    public bool RendersCachedResult { get; init; }
    public bool UsesPlaceholder { get; init; }
    public int NestingDepth { get; init; }
    public int InstructionRunCount { get; init; }
    public int ResultRunCount { get; init; }
}

internal sealed record DocxHyperlinkSpan(
    string? RelationshipId,
    string? Anchor,
    string? Tooltip,
    string? HistoryValue,
    string? Target,
    string? TargetMode,
    string? ResolvedTarget,
    int SourceRunStartIndex,
    int SourceRunCount,
    int TextRunStartIndex,
    int TextRunCount,
    int TextLength);

internal sealed record DocxTabStop(
    double? PositionPoints,
    string? PositionValue,
    string? Value,
    string? LeaderValue);

internal sealed record DocxParagraphSpacing(
    string? BeforeValue,
    string? AfterValue,
    string? BeforeLinesValue,
    string? AfterLinesValue,
    string? BeforeAutoSpacingValue,
    string? AfterAutoSpacingValue,
    string? LineValue,
    string? LineRuleValue,
    bool? ContextualSpacing)
{
    public static DocxParagraphSpacing Empty { get; } = new(null, null, null, null, null, null, null, null, null);
}

internal sealed record DocxParagraphKeepRules(
    bool? KeepNext,
    string? KeepNextValue,
    bool? KeepLines,
    string? KeepLinesValue,
    bool? WidowControl,
    string? WidowControlValue)
{
    public static DocxParagraphKeepRules Empty { get; } = new(null, null, null, null, null, null);
}

internal sealed record DocxParagraphIndent(
    double? LeftPoints,
    double? RightPoints,
    double? FirstLinePoints,
    double? HangingPoints,
    string? LeftValue,
    string? RightValue,
    string? FirstLineValue,
    string? HangingValue)
{
    public static DocxParagraphIndent Empty { get; } = new(null, null, null, null, null, null, null, null);
}

internal sealed record DocxListLabel(
    string Text,
    string FormatValue,
    string LevelTextValue,
    string SuffixValue,
    string NumberId,
    int Level,
    DocxNumberingIndent Indent,
    DocxTextRunStyle Style);

internal sealed record DocxNumberingIndent(
    double? LeftPoints,
    double? RightPoints,
    double? FirstLinePoints,
    double? HangingPoints,
    double? NumberingTabPositionPoints,
    string? LeftValue,
    string? RightValue,
    string? FirstLineValue,
    string? HangingValue,
    string? NumberingTabValue,
    string? NumberingTabPositionValue)
{
    public static DocxNumberingIndent Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null);
}

internal sealed record DocxTextRun(
    string Text,
    double FontSize,
    string? ColorHex,
    bool Bold,
    bool Italic,
    bool Underline,
    string? UnderlineValue,
    string? FontFamily,
    double CharacterSpacingPoints,
    bool AllCaps,
    string? VerticalAlignmentValue,
    bool Strike,
    string? StrikeValue,
    bool DoubleStrike,
    string? DoubleStrikeValue,
    string? HighlightValue,
    string? ShadingFillHex,
    string? ShadingValue,
    string? ShadingColor,
    bool SmallCaps,
    string? SmallCapsValue,
    bool Hidden,
    string? HiddenValue,
    string? UnderlineColorHex)
{

    public DocxTextRun(
        string Text,
        double FontSize,
        string? ColorHex,
        bool Bold,
        bool Italic,
        bool Underline,
        string? UnderlineValue,
        string? FontFamily)
        : this(Text, FontSize, ColorHex, Bold, Italic, Underline, UnderlineValue, FontFamily, 0d, false, null, false, null, false, null, null, null, null, null, false, null, false, null, null)
    {
    }

    public DocxTextRun(
        string Text,
        double FontSize,
        string? ColorHex,
        bool Bold,
        bool Italic,
        bool Underline,
        string? UnderlineValue,
        string? FontFamily,
        double CharacterSpacingPoints)
        : this(Text, FontSize, ColorHex, Bold, Italic, Underline, UnderlineValue, FontFamily, CharacterSpacingPoints, false, null, false, null, false, null, null, null, null, null, false, null, false, null, null)
    {
    }
    public DocxRunFonts Fonts { get; init; } = DocxRunFonts.Empty;
    public DocxRunStyleResolution StyleResolution { get; init; } = DocxRunStyleResolution.Empty;
    public int SourceRunIndex { get; init; } = -1;
    public int SourceTextOffsetInRun { get; init; } = 0;
    public DocxRevisionInfo? Revision { get; init; }
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];

    public DocxEffectiveRunProperties EffectiveProperties => new(
        FontSize,
        ColorHex,
        Bold,
        Italic,
        Underline,
        UnderlineValue,
        FontFamily,
        Fonts,
        CharacterSpacingPoints,
        AllCaps,
        VerticalAlignmentValue,
        Strike,
        StrikeValue,
        DoubleStrike,
        DoubleStrikeValue,
        HighlightValue,
        ShadingFillHex,
        ShadingValue,
        ShadingColor,
        SmallCaps,
        SmallCapsValue,
        Hidden,
        HiddenValue,
        UnderlineColorHex,
        StyleResolution);
}

internal sealed record DocxEffectiveRunProperties(
    double FontSize,
    string? ColorHex,
    bool Bold,
    bool Italic,
    bool Underline,
    string? UnderlineValue,
    string? FontFamily,
    DocxRunFonts Fonts,
    double CharacterSpacingPoints,
    bool AllCaps,
    string? VerticalAlignmentValue,
    bool Strike,
    string? StrikeValue,
    bool DoubleStrike,
    string? DoubleStrikeValue,
    string? HighlightValue,
    string? ShadingFillHex,
    string? ShadingValue,
    string? ShadingColor,
    bool SmallCaps,
    string? SmallCapsValue,
    bool Hidden,
    string? HiddenValue,
    string? UnderlineColorHex,
    DocxRunStyleResolution StyleResolution);

internal sealed record DocxRunStyleResolution(
    string? CharacterStyleId,
    bool CharacterStyleFound,
    int CharacterStyleDepth,
    bool HasDocumentDefaultRunProperties,
    bool HasParagraphStyleRunProperties,
    bool HasCharacterStyleRunProperties,
    bool HasDirectRunProperties,
    bool HasTableStyleRunProperties)
{
    public static DocxRunStyleResolution Empty { get; } = new(null, false, 0, false, false, false, false, false);
}

internal sealed record DocxRevisionInfo
{
    public DocxRevisionInfo(
        DocxRevisionKind kind,
        string? id,
        string? author,
        string? date,
        string sourceElement,
        DocxRevisionPropertyFamily? propertyChangeFamily,
        IReadOnlyList<string> propertyElementNames)
    {
        Kind = kind;
        Id = id;
        Author = author;
        Date = date;
        SourceElement = sourceElement;
        PropertyChangeFamily = propertyChangeFamily;
        PropertyElementNames = propertyElementNames.ToArray();
    }

    public DocxRevisionKind Kind { get; init; }

    public string? Id { get; init; }

    public string? Author { get; init; }

    public string? Date { get; init; }

    public string SourceElement { get; init; }

    public DocxRevisionPropertyFamily? PropertyChangeFamily { get; init; }

    public IReadOnlyList<string> PropertyElementNames { get; init; }
}

internal sealed record DocxTextRunStyle(
    double? FontSize,
    string? ColorHex,
    bool? Bold,
    bool? Italic,
    bool? Underline,
    string? UnderlineValue,
    string? FontFamily,
    DocxRunFonts Fonts,
    double? CharacterSpacingPoints,
    bool? AllCaps,
    string? VerticalAlignmentValue,
    bool? Strike,
    string? StrikeValue,
    bool? DoubleStrike,
    string? DoubleStrikeValue,
    string? HighlightValue,
    string? ShadingFillHex,
    string? ShadingValue,
    string? ShadingColor,
    bool? SmallCaps,
    string? SmallCapsValue,
    bool? Hidden,
    string? HiddenValue,
    string? UnderlineColorHex)
{

    public DocxTextRunStyle(
        double? FontSize,
        string? ColorHex,
        bool? Bold,
        bool? Italic,
        bool? Underline,
        string? UnderlineValue,
        string? FontFamily,
        DocxRunFonts Fonts)
        : this(FontSize, ColorHex, Bold, Italic, Underline, UnderlineValue, FontFamily, Fonts, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null)
    {
    }
    public static DocxTextRunStyle Empty { get; } = new(null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

    public DocxTextRun ApplyTo(DocxTextRun? baseRun, string text, double fallbackFontSize)
    {
        var source = baseRun ?? new DocxTextRun(string.Empty, fallbackFontSize, null, false, false, false, null, null);
        return new DocxTextRun(
            text,
            FontSize ?? source.FontSize,
            ColorHex ?? source.ColorHex,
            Bold ?? source.Bold,
            Italic ?? source.Italic,
            Underline ?? source.Underline,
            UnderlineValue ?? source.UnderlineValue,
            FontFamily ?? source.FontFamily,
            CharacterSpacingPoints ?? source.CharacterSpacingPoints,
            AllCaps ?? source.AllCaps,
            VerticalAlignmentValue ?? source.VerticalAlignmentValue,
            Strike ?? source.Strike,
            StrikeValue ?? source.StrikeValue,
            DoubleStrike ?? source.DoubleStrike,
            DoubleStrikeValue ?? source.DoubleStrikeValue,
            HighlightValue ?? source.HighlightValue,
            ShadingFillHex ?? source.ShadingFillHex,
            ShadingValue ?? source.ShadingValue,
            ShadingColor ?? source.ShadingColor,
            SmallCaps ?? source.SmallCaps,
            SmallCapsValue ?? source.SmallCapsValue,
            Hidden ?? source.Hidden,
            HiddenValue ?? source.HiddenValue,
            UnderlineColorHex ?? source.UnderlineColorHex)
        {
            Fonts = source.Fonts.Merge(Fonts),
            StyleResolution = source.StyleResolution,
            SourceRunIndex = source.SourceRunIndex,
            SourceTextOffsetInRun = source.SourceTextOffsetInRun,
            Revision = source.Revision,
            Revisions = source.Revisions
        };
    }
}

internal sealed record DocxRunFonts(
    string? Ascii,
    string? HighAnsi,
    string? EastAsia,
    string? ComplexScript,
    string? AsciiTheme,
    string? HighAnsiTheme,
    string? EastAsiaTheme,
    string? ComplexScriptTheme)
{
    public static DocxRunFonts Empty { get; } = new(null, null, null, null, null, null, null, null);

    public DocxRunFonts Merge(DocxRunFonts other)
    {
        return new DocxRunFonts(
            other.Ascii ?? Ascii,
            other.HighAnsi ?? HighAnsi,
            other.EastAsia ?? EastAsia,
            other.ComplexScript ?? ComplexScript,
            other.AsciiTheme ?? AsciiTheme,
            other.HighAnsiTheme ?? HighAnsiTheme,
            other.EastAsiaTheme ?? EastAsiaTheme,
            other.ComplexScriptTheme ?? ComplexScriptTheme);
    }
}

internal sealed record DocxInlineImage(double WidthPoints, double HeightPoints, string ContentType, byte[] Bytes, string? PartName)
{
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal sealed record DocxTable(
    string? LayoutValue,
    IReadOnlyList<double> ColumnWidthsPoints,
    IReadOnlyList<DocxTableRow> Rows,
    string? StyleId,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    double? IndentPoints,
    string? IndentValue,
    string? IndentType,
    double? CellSpacingPoints,
    string? CellSpacingValue,
    string? CellSpacingType,
    DocxTableLook? Look,
    bool HasExplicitGrid)
{

    public DocxTable(
        string? LayoutValue,
        IReadOnlyList<double> ColumnWidthsPoints,
        IReadOnlyList<DocxTableRow> Rows)
        : this(LayoutValue, ColumnWidthsPoints, Rows, null, null, null, null, null, null, null, null, null, null, null, true)
    {
    }
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal sealed record DocxTableLook(
    string? Value,
    bool? FirstRow,
    string? FirstRowValue,
    bool? LastRow,
    string? LastRowValue,
    bool? FirstColumn,
    string? FirstColumnValue,
    bool? LastColumn,
    string? LastColumnValue,
    bool? NoHorizontalBand,
    string? NoHorizontalBandValue,
    bool? NoVerticalBand,
    string? NoVerticalBandValue)
{
    public static DocxTableLook Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null, null);
}

internal sealed record DocxTableRow(
    IReadOnlyList<DocxTableCell> Cells,
    double? HeightPoints,
    bool IsHeader,
    string? HeaderValue,
    string? HeightValue,
    string? HeightRuleValue,
    DocxTableCellMargins? TablePropertyExceptionCellMargins,
    bool CantSplit,
    string? CantSplitValue)
{

    public DocxTableRow(
        IReadOnlyList<DocxTableCell> Cells,
        double? HeightPoints)
        : this(Cells, HeightPoints, false, null, null, null, null, false, null)
    {
    }
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal sealed record DocxTableCell(
    string Text,
    IReadOnlyList<DocxParagraph> Paragraphs,
    string? FillHex,
    string? ShadingValue,
    string? ShadingColor,
    string? VerticalAlignmentValue,
    IReadOnlyList<DocxTableCellBorder> Borders,
    DocxTableCellMargins Margins,
    double? PreferredWidthPoints,
    string? PreferredWidthValue,
    string? PreferredWidthType,
    int GridSpan,
    string? GridSpanValue,
    DocxTableCellConditionalFormat? ConditionalFormat,
    bool HasVerticalMerge,
    string? VerticalMergeValue,
    bool NoWrap,
    string? NoWrapValue,
    bool FitText,
    string? FitTextValue,
    string? TextDirectionValue)
{

    public DocxTableCell(
        string Text,
        IReadOnlyList<DocxParagraph> Paragraphs,
        string? FillHex,
        string? ShadingValue,
        string? ShadingColor,
        string? VerticalAlignmentValue,
        IReadOnlyList<DocxTableCellBorder> Borders,
        DocxTableCellMargins Margins)
        : this(Text, Paragraphs, FillHex, ShadingValue, ShadingColor, VerticalAlignmentValue, Borders, Margins, null, null, null, 1, null, null, false, null, false, null, false, null, null)
    {
    }
    public IReadOnlyList<DocxBodyElement> BodyElements { get; init; } = [];
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal sealed record DocxTableCellConditionalFormat(
    string? Value,
    bool? FirstRow,
    string? FirstRowValue,
    bool? LastRow,
    string? LastRowValue,
    bool? FirstColumn,
    string? FirstColumnValue,
    bool? LastColumn,
    string? LastColumnValue,
    bool? OddHorizontalBand,
    string? OddHorizontalBandValue,
    bool? EvenHorizontalBand,
    string? EvenHorizontalBandValue,
    bool? OddVerticalBand,
    string? OddVerticalBandValue,
    bool? EvenVerticalBand,
    string? EvenVerticalBandValue,
    bool? FirstRowFirstColumn,
    string? FirstRowFirstColumnValue,
    bool? FirstRowLastColumn,
    string? FirstRowLastColumnValue,
    bool? LastRowFirstColumn,
    string? LastRowFirstColumnValue,
    bool? LastRowLastColumn,
    string? LastRowLastColumnValue)
{
    public bool IsDefined =>
        FirstRow is not null ||
        LastRow is not null ||
        FirstColumn is not null ||
        LastColumn is not null ||
        OddHorizontalBand is not null ||
        EvenHorizontalBand is not null ||
        OddVerticalBand is not null ||
        EvenVerticalBand is not null ||
        FirstRowFirstColumn is not null ||
        FirstRowLastColumn is not null ||
        LastRowFirstColumn is not null ||
        LastRowLastColumn is not null;
}

internal sealed record DocxTableCellBorder(string Edge, string? Value, string? Color, string? SizeValue);

internal sealed record DocxTableCellMargins(
    double? TopPoints,
    double? RightPoints,
    double? BottomPoints,
    double? LeftPoints,
    string? TopValue,
    string? RightValue,
    string? BottomValue,
    string? LeftValue)
{
    public static DocxTableCellMargins Empty { get; } = new(null, null, null, null, null, null, null, null);
}

internal enum DocxTextAlignment
{
    Left,
    Center,
    Right,
    Justified
}