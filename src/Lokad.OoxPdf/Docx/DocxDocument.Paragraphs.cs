namespace Lokad.OoxPdf.Docx;

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
    public IReadOnlyList<DocxInlineTextBox> InlineTextBoxes { get; init; } = [];
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

    public DocxParagraphSpacing Merge(DocxParagraphSpacing other)
    {
        bool hasOtherBeforeSide = HasBeforeSpacingSide(other);
        bool hasOtherAfterSide = HasAfterSpacingSide(other);
        return new DocxParagraphSpacing(
            hasOtherBeforeSide ? other.BeforeValue : BeforeValue,
            hasOtherAfterSide ? other.AfterValue : AfterValue,
            hasOtherBeforeSide ? other.BeforeLinesValue : BeforeLinesValue,
            hasOtherAfterSide ? other.AfterLinesValue : AfterLinesValue,
            hasOtherBeforeSide ? other.BeforeAutoSpacingValue : BeforeAutoSpacingValue,
            hasOtherAfterSide ? other.AfterAutoSpacingValue : AfterAutoSpacingValue,
            other.LineValue ?? LineValue,
            other.LineRuleValue ?? LineRuleValue,
            other.ContextualSpacing ?? ContextualSpacing);
    }

    internal static bool HasBeforeSpacingSide(DocxParagraphSpacing spacing)
    {
        return spacing.BeforeValue is not null ||
            spacing.BeforeLinesValue is not null ||
            spacing.BeforeAutoSpacingValue is not null;
    }

    internal static bool HasAfterSpacingSide(DocxParagraphSpacing spacing)
    {
        return spacing.AfterValue is not null ||
            spacing.AfterLinesValue is not null ||
            spacing.AfterAutoSpacingValue is not null;
    }
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

    public DocxParagraphKeepRules Merge(DocxParagraphKeepRules other)
    {
        return new DocxParagraphKeepRules(
            other.KeepNext ?? KeepNext,
            other.KeepNextValue ?? KeepNextValue,
            other.KeepLines ?? KeepLines,
            other.KeepLinesValue ?? KeepLinesValue,
            other.WidowControl ?? WidowControl,
            other.WidowControlValue ?? WidowControlValue);
    }
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

    public DocxParagraphIndent Merge(DocxParagraphIndent other)
    {
        bool hasOtherFirstLineSide = other.FirstLineValue is not null || other.HangingValue is not null;
        return new DocxParagraphIndent(
            other.LeftPoints ?? LeftPoints,
            other.RightPoints ?? RightPoints,
            hasOtherFirstLineSide ? other.FirstLinePoints : FirstLinePoints,
            hasOtherFirstLineSide ? other.HangingPoints : HangingPoints,
            other.LeftValue ?? LeftValue,
            other.RightValue ?? RightValue,
            hasOtherFirstLineSide ? other.FirstLineValue : FirstLineValue,
            hasOtherFirstLineSide ? other.HangingValue : HangingValue);
    }
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

    // Typed field identity for PAGE/NUMPAGES placeholders. Null for ordinary text, including
    // literal {PAGE}/{NUMPAGES} content, so downstream substitution touches only actual fields.
    public DocxFieldKind? FieldKind { get; init; }

    // Office A/B (w56 superscript plus w57b subscript size curves, Word-COM rendered):
    // style-less unsized super/subscript runs scale glyphs from the 11pt document
    // default, not the 12pt body fallback, while paragraph layout keeps the resolved
    // nominal (max-nominal stays 12 there). Null means nominal (explicit, styled, or
    // table-linked sizes, plus all non-script runs).
    public double? ScriptBaseFontSize { get; init; }
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

// An inline DrawingML textbox (wp:inline plus wps:txbx, no blip): unlike floating
// drawings it has no page placement and joins body flow as a block. Extents and
// insets are file values; the layout scales them into its space.
internal sealed record DocxInlineTextBox(
    string? ExtentCxValue,
    string? ExtentCyValue,
    string? TextBoxInsetLeftValue = null,
    string? TextBoxInsetTopValue = null,
    string? TextBoxInsetRightValue = null,
    string? TextBoxInsetBottomValue = null)
{
    public IReadOnlyList<DocxBodyElement> BodyElements { get; init; } = [];
    public IReadOnlyList<DocxRevisionInfo> Revisions { get; init; } = [];
}

internal enum DocxTextAlignment
{
    Left,
    Center,
    Right,
    Justified
}
