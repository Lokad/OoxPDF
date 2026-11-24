namespace Lokad.OoxPdf.Docx;

internal sealed record DocxStructureBlockSnapshot(int Index, string Kind, string? PreviousKind, string? NextKind)
{
    public string? ParagraphStyleId { get; init; }
    public bool? ParagraphStyleFound { get; init; }
    public int? ParagraphStyleDepth { get; init; }
    public bool? HasDocumentDefaultParagraphProperties { get; init; }
    public bool? HasDirectParagraphProperties { get; init; }
    public bool? HasTableStyleParagraphProperties { get; init; }
    public int RunCount { get; init; }
    public int TextLength { get; init; }
    public bool HasVisibleText { get; init; }
    public int RevisionCount { get; init; }
    public int InsertionRevisionCount { get; init; }
    public int DeletionRevisionCount { get; init; }
    public int MoveFromRevisionCount { get; init; }
    public int MoveToRevisionCount { get; init; }
    public int OtherRevisionCount { get; init; }
    public int RevisionRangeCount { get; init; }
    public string? ListFormatValue { get; init; }
    public int InlineImageCount { get; init; }
    public int InlineReferenceCount { get; init; }
    public int AnchoredInlineReferenceCount { get; init; }
    public int ResolvedInlineReferenceCount { get; init; }
    public int MaxInlineReferenceTextOffsetInRun { get; init; }
    public int CommentReferenceCount { get; init; }
    public int FootnoteReferenceCount { get; init; }
    public int EndnoteReferenceCount { get; init; }
    public int FieldReferenceCount { get; init; }
    public int PageFieldReferenceCount { get; init; }
    public int NumPagesFieldReferenceCount { get; init; }
    public int OtherFieldReferenceCount { get; init; }
    public int ComplexFieldReferenceCount { get; init; }
    public int CachedResultFieldReferenceCount { get; init; }
    public int RenderedCachedResultFieldReferenceCount { get; init; }
    public int PlaceholderFieldReferenceCount { get; init; }
    public int NestedFieldReferenceCount { get; init; }
    public int BookmarkAnchorCount { get; init; }
    public int HyperlinkCount { get; init; }
    public int ExternalHyperlinkCount { get; init; }
    public int InternalHyperlinkCount { get; init; }
    public int TableRowCount { get; init; }
    public int? TableIndex { get; init; }
    public double? SpacingBeforePoints { get; init; }
    public double? SpacingAfterPoints { get; init; }
    public double? LineSpacingPoints { get; init; }
    public double? LineSpacingFactor { get; init; }
    public bool HasBeforeSpacingToken { get; init; }
    public bool HasAfterSpacingToken { get; init; }
    public string? BeforeAutoSpacingValue { get; init; }
    public string? AfterAutoSpacingValue { get; init; }
    public bool? ContextualSpacing { get; init; }
    public bool? KeepNext { get; init; }
    public bool? KeepLines { get; init; }
    public bool? WidowControl { get; init; }
    public bool? WordWrap { get; init; }
    public string? WordWrapValue { get; init; }
    public int? TableMaxColumnCount { get; init; }
    public double? TablePreferredWidthPoints { get; init; }
    public string? TablePreferredWidthType { get; init; }
    public double? TableIndentPoints { get; init; }
    public double? TableCellSpacingPoints { get; init; }
    public string? TableLayoutValue { get; init; }
    public string? PageBreakSourceKind { get; init; }
    public string? PageBreakValue { get; init; }
    public string? ManualBreakSourceKind { get; init; }
    public string? ManualBreakValue { get; init; }
    public string? ImplicitSourceKind { get; init; }
    public string? SectionBreakTypeValue { get; init; }
    public string? SectionColumnCountValue { get; init; }
    public string? SectionColumnEqualWidthValue { get; init; }
    public string? SectionColumnSpaceValue { get; init; }
    public int? SectionColumnDefinitionCount { get; init; }
    public int? SectionColumnDefinitionWidthTokenCount { get; init; }
    public int? SectionColumnDefinitionSpaceTokenCount { get; init; }
    public double? ParagraphIndentLeftPoints { get; init; }
    public double? ParagraphIndentRightPoints { get; init; }
    public double? ParagraphIndentFirstLinePoints { get; init; }
    public double? ParagraphIndentHangingPoints { get; init; }
    public int? WhitespaceDelimitedTokenCount { get; init; }
    public int? LongestWhitespaceDelimitedTokenLength { get; init; }
    public int? SpaceCharacterCount { get; init; }
    public int? NonAsciiCharacterCount { get; init; }
    public int? PunctuationCharacterCount { get; init; }
    public int? DigitCharacterCount { get; init; }
    public int? UppercaseCharacterCount { get; init; }
    public int? LowercaseCharacterCount { get; init; }
    public int? TabStopCount { get; init; }
    public bool? SnapToGrid { get; init; }
    public string? SnapToGridValue { get; init; }
    public bool? PageBreakConsumesParagraphLine { get; init; }
    public double? PageBreakLineSpacingPoints { get; init; }
    public double? PageBreakLineSpacingFactor { get; init; }
    public bool? ManualBreakConsumesParagraphLine { get; init; }
    public double? ManualBreakLineSpacingPoints { get; init; }
    public double? ManualBreakLineSpacingFactor { get; init; }
    public int? SourceBreakCount { get; init; }
}

internal sealed record DocxStructureStorySnapshot(
    string Kind,
    string Scope,
    int? SectionBreakBlockIndex,
    string? VariantType,
    int BlockCount,
    int ParagraphCount,
    int TableCount,
    int TextLength,
    int RevisionCount,
    int InsertionRevisionCount,
    int DeletionRevisionCount,
    int MoveFromRevisionCount,
    int MoveToRevisionCount,
    int OtherRevisionCount,
    int InlineImageCount,
    int InlineReferenceCount,
    int CommentReferenceCount,
    int ResolvedInlineReferenceCount,
    int FieldReferenceCount,
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
    bool HasCommentAuthor,
    bool HasCommentInitials,
    bool HasCommentDate,
    string? CommentParagraphId,
    string? CommentParentParagraphId,
    string? CommentParentId,
    bool? CommentResolved);

internal sealed record DocxStructureInlineReferenceSnapshot(
    int SourceBlockIndex,
    string SourceBlockKind,
    int SourceParagraphIndex,
    string Kind,
    string? Id,
    string? CustomMarkFollowsValue,
    string? DisplayText,
    int SourceRunIndex,
    int RunChildIndex,
    int TextOffsetInRun,
    int? ResolvedStoryIndex,
    string? ResolvedStoryKind,
    string? ResolvedStoryPartName,
    string? ResolvedStoryId,
    int? ResolvedStoryBlockCount,
    int? ResolvedStoryTextLength,
    int RevisionCount,
    string? RevisionKind,
    string? RevisionSourceElement);

internal sealed record DocxStructureCommentRangeSnapshot(
    int SourceBlockIndex,
    string SourceBlockKind,
    int SourceParagraphIndex,
    string? Id,
    int? StartSourceRunIndex,
    int? StartTextOffset,
    int? EndSourceRunIndex,
    int? EndTextOffset,
    int? ReferenceSourceRunIndex,
    int? ReferenceTextOffset,
    int? ResolvedStoryIndex,
    string? ResolvedStoryPartName,
    string? ResolvedStoryId,
    int? ResolvedStoryBlockCount,
    int? ResolvedStoryTextLength);

internal sealed record DocxStructureCommentStoryAnchorSnapshot(
    string? Id,
    string Status,
    bool HasPackageAnchor,
    bool HasHiddenAnchor,
    int VisibleInlineReferenceCount,
    int VisibleRangeCount);

internal sealed record DocxStructureRevisionRangeSnapshot(
    int SourceBlockIndex,
    string SourceBlockKind,
    int SourceParagraphIndex,
    string Kind,
    string? Id,
    bool HasName,
    bool HasAuthor,
    bool HasDate,
    int? StartSourceRunIndex,
    int? StartTextOffset,
    int? EndSourceRunIndex,
    int? EndTextOffset,
    bool IsClosed)
{
    public bool IsLinkedAcrossBlocks { get; init; }
    public int? LinkedSourceBlockIndex { get; init; }
    public int? LinkedSourceParagraphIndex { get; init; }
}

internal sealed record DocxStructureFormattingRevisionPropertySnapshot(
    string Family,
    string SourceElement,
    string PropertyElementName,
    int Count);

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
    int? SourceBlockIndex,
    int TextBoxBlockCount,
    int TextBoxParagraphCount,
    int TextBoxTextLength,
    int RevisionCount,
    int InsertionRevisionCount,
    int DeletionRevisionCount,
    int MoveFromRevisionCount,
    int MoveToRevisionCount,
    int OtherRevisionCount);

internal sealed record DocxStructureStyleUsageSnapshot(
    string Kind,
    string? StyleId,
    int Count,
    int ParagraphCount,
    int TableCount,
    int TextLength,
    int BeforeSpacingTokenParagraphCount,
    int AfterSpacingTokenParagraphCount,
    int BeforeAutoSpacingParagraphCount,
    int AfterAutoSpacingParagraphCount,
    int BeforeLinesSpacingParagraphCount,
    int AfterLinesSpacingParagraphCount,
    int ContextualSpacingParagraphCount,
    int ExactLineSpacingParagraphCount,
    int AtLeastLineSpacingParagraphCount,
    int AutoLineSpacingParagraphCount,
    int TableStyleParagraphPropertiesCount);

internal sealed record DocxStructureListUsageSnapshot(
    string NumberId,
    int Level,
    string FormatValue,
    string SuffixValue,
    int ParagraphCount,
    int TextLength,
    int LeftIndentParagraphCount,
    int RightIndentParagraphCount,
    int FirstLineIndentParagraphCount,
    int HangingIndentParagraphCount,
    int NumberingTabParagraphCount,
    int ParagraphIndentOverrideCount,
    int ParagraphNumberingTabStopCount);

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
    int SingleBorderCount,
    int ThickBorderCount,
    int DoubleBorderCount,
    int DottedBorderCount,
    int DashedBorderCount,
    int SuppressedBorderCount,
    int OtherBorderStyleCount,
    int ParagraphCount,
    int RunCount,
    int TextLength,
    int RevisionCount,
    int InsertionRevisionCount,
    int DeletionRevisionCount,
    int MoveFromRevisionCount,
    int MoveToRevisionCount,
    int OtherRevisionCount,
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
    int RevisionCount,
    int InsertionRevisionCount,
    int DeletionRevisionCount,
    int MoveFromRevisionCount,
    int MoveToRevisionCount,
    int OtherRevisionCount,
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
    int RevisionCount,
    int InsertionRevisionCount,
    int DeletionRevisionCount,
    int MoveFromRevisionCount,
    int MoveToRevisionCount,
    int OtherRevisionCount,
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
    int BodyElementCount,
    int ManualBreakElementCount,
    int PageBreakElementCount,
    int NestedTableElementCount);

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
