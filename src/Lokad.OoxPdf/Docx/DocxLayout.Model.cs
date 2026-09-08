using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal sealed record DocxLayoutPage(
    double Width,
    double Height,
    double MarginLeft,
    double MarginRight,
    double MarkupMarginReservePoints,
    double MarginTop,
    double MarginBottom,
    DocxPageSettings PageSettings,
    DocxSectionLayoutProperties SectionProperties,
    IReadOnlyList<DocxLayoutColumnFrame> ColumnFrames,
    IReadOnlyList<DocxTextLineLayout> StaticTextLines,
    IReadOnlyList<DocxInlineImageLayout> StaticInlineImages,
    IReadOnlyList<DocxTableRowLayout> StaticTableRows,
    IReadOnlyList<DocxPlacedRelatedStoryLayout> PlacedRelatedStories,
    IReadOnlyList<DocxLayoutItem> Items);

// One positioned content item on a laid-out page, emitted in the order the
// renderer draws it. Implementors (text line, inline image, table row fragment)
// carry page-space coordinates plus source/story provenance; layout owns
// fragmentation (wrapped lines, row fragments), never the renderer.
internal abstract record DocxLayoutItem;

internal sealed record DocxTextLineLayout(
    string Text,
    DocxTextRun StyleRun,
    double FontSize,
    double X,
    double BaselineY,
    double Width,
    IReadOnlyList<DocxTextSegmentLayout> Segments,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    int? SourceLineIndex,
    string? StoryKind,
    double? LineHeight,
    double? AppliedBeforeSpacing,
    bool? IsFirstParagraphLine,
    bool EndsWithIntraTokenBreak,
    double? SingleLineHeight,
    double? ListLabelSingleLineHeight,
    double? BodyWindowsLineHeight,
    double? ListLabelWindowsLineHeight,
    double? EffectiveLineSpacingFactor,
    bool? LineSpacingFactorFloorApplied,
    double? PendingAfterSpacing,
    double? ParagraphBeforeSpacing,
    double? ParagraphAfterSpacing,
    bool? ContextualSpacingSuppressed,
    DocxParagraph? SourceParagraph,
    string? StoryVariantType,
    DocxLineHeightSource? LineHeightSource,
    bool EmitsTerminalParagraphMark) : DocxLayoutItem;

internal sealed record DocxTextSegmentLayout(
    string Text,
    DocxTextRun StyleRun,
    double X,
    double Width,
    double? FontSize,
    double BaselineOffsetY,
    double PdfCharacterSpacing,
    DocxTextStateCharacterSpacingSource PdfCharacterSpacingSource,
    bool CompensatePdfCharacterSpacing,
    int SourceTextRunIndex,
    int SourceTextOffsetInRun,
    DocxTextSegmentRole Role);

internal enum DocxTextSegmentRole
{
    Text,
    ListLabel,
    ListSeparator
}

internal sealed record DocxTextSpan(
    string Text,
    DocxTextRun StyleRun,
    int SourceTextRunIndex,
    int SourceTextOffsetInRun);

internal sealed record DocxWrappedTextLine(
    string Text,
    IReadOnlyList<DocxTextSpan> Spans,
    bool EndsWithIntraTokenBreak);

internal sealed record DocxInlineImageLayout(
    DocxInlineImage Image,
    double X,
    double Y,
    double Width,
    double Height,
    int PageIndex,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    string? StoryKind,
    string? StoryVariantType) : DocxLayoutItem;

// An inline textbox laid out as a body-flow block. Coordinates are absolute flow
// space (the box origin is known at layout time), so consumers use them directly
// with the body offsets; no emission map applies.
internal sealed record DocxInlineTextBoxLayout(
    DocxInlineTextBox TextBox,
    double BoxX,
    double BoxTop,
    double BoxWidth,
    double BoxHeight,
    IReadOnlyList<DocxTextLineLayout> TextLines,
    IReadOnlyList<DocxInlineImageLayout> InlineImages,
    IReadOnlyList<DocxTableRowLayout> TableRows,
    int? SourceBlockIndex,
    int? SourceParagraphIndex,
    string? StoryKind,
    string? StoryVariantType) : DocxLayoutItem;

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
    int RevisionCount,
    IReadOnlyList<DocxRevisionInfo>? Revisions,
    string? StoryKind,
    string? StoryVariantType) : DocxLayoutItem;

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
    string? LayoutValue,
    IReadOnlyList<DocxRevisionInfo>? Revisions);

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
    bool IsVerticalMergeContinuation,
    DocxTableCell? VerticalMergeOwnerCell,
    DocxVerticalMergeOwner? VerticalMergeOwner,
    DocxTableCellVisualOwnership VisualOwnership,
    IReadOnlyList<DocxTableRowLayout>? NestedTableRows)
{
    public IReadOnlyList<DocxInlineTextBoxLayout> InlineTextBoxes { get; init; } = [];
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

internal sealed record DocxFallbackFontEntry(OpenTypeFont Font, DocxRunFontResource Resource);

internal sealed record DocxFontResources(
    DocxFontPlan Plan,
    IDocxTextMeasurer? TextMeasurer,
    IReadOnlyList<PdfFontResource> Resources,
    IReadOnlyDictionary<DocxTextRun, DocxRunFontResource> RunResources,
    DocxRunFontResource? Fallback,
    IReadOnlyDictionary<DocxTextRun, IReadOnlyList<DocxFallbackFontEntry>> FallbackChains);

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
    int SourceTextRunIndex,
    int SourceTextOffsetInRun,
    DocxTextSegmentRole Role);

internal readonly record struct DocxKeepBlockEstimate(
    double Height,
    int ParagraphCount,
    int FirstTableRowCount);
