using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    internal readonly record struct TextRun(
        string Text,
        double X,
        double Y,
        double Width,
        double Height,
        double ClipX,
        double ClipY,
        double ClipWidth,
        double ClipHeight,
        double FontSize,
        double CharacterSpacing,
        double BaselineOffset,
        RgbColor Color,
        double Alpha,
        RgbColor? HighlightColor,
        bool Bold,
        bool Italic,
        bool Underline,
        bool Strike,
        bool KerningEnabled,
        TextAlignment Alignment,
        string? FontFamily,
        double RotationDegrees,
        double RotationCenterX,
        double RotationCenterY,
        bool FlipHorizontal,
        bool FlipVertical,
        bool PreventCoalesce,
        TextOutline? Outline,
        bool StrictClip,
        int GlyphRotationQuarterTurns = 0);

    private sealed record TextGlyphRun(
        TextRun Source,
        string ResourceName,
        PdfEmbeddedFont Font,
        string GlyphHex,
        string? PositioningArray,
        IReadOnlyList<TextGlyphAtom> Glyphs,
        double X,
        double BaselineY,
        double Width,
        double PdfFontSize,
        double PdfCharacterSpacing,
        bool SyntheticBold,
        bool SyntheticItalic);

    private readonly record struct TextHighlightRectangle(double X, double Y, double Width, double Height);

    private readonly record struct TextDecorationRectangle(double X, double Y, double Width, double Height);

    internal sealed record PptxPositionedTextSpan(
        PptxTextRunModel? SourceRun,
        PptxTextLineBoxLayout? LineBox,
        int FrameIndex,
        int ParagraphIndex,
        int? SourceRunIndex,
        PptxParagraphBulletKind ParagraphBulletKind,
        string? ParagraphAutoNumberType,
        int? ParagraphAutoNumberStartAt,
        int LineIndex,
        int SpanIndex,
        int LineSpanCount,
        double FrameFontScale,
        double FrameShapeX,
        double FrameShapeTopY,
        double FrameShapeWidth,
        double FrameShapeHeight,
        int? TableRowIndex,
        int? TableColumnIndex,
        int? TableRowSpan,
        int? TableColumnSpan,
        double FrameInsetLeft,
        double FrameInsetRight,
        double FrameInsetTop,
        double FrameInsetBottom,
        PptxTextWrapMode FrameWrapMode,
        string? FrameWrapValue,
        PptxTextVerticalOverflow FrameVerticalOverflowMode,
        string? FrameVerticalOverflowValue,
        PptxTextBodyPropertySource FrameVerticalOverflowSource,
        PptxTextAutofitMode FrameAutofitMode,
        string FrameAutofitModeValue,
        double FrameTextX,
        double FrameTextWidth,
        double FrameTextWrapWidth,
        double FrameTextHeight,
        double FrameClipX,
        double FrameClipWidth,
        double FrameClipY,
        double FrameClipHeight,
        int FrameColumnCount,
        double FrameColumnSpacing,
        TextAlignment SourceAlignment,
        TextRun Run,
        double EndX,
        IReadOnlyList<PptxTextAtomLayout> Atoms,
        PptxTextGlyphSpanLayout GlyphSpan,
        double? PdfCharacterSpacingOverride);

    private readonly record struct PptxPdfTextEmissionContext(
        double LayoutFontSize,
        double BaselineY,
        int FrameIndex,
        int ParagraphIndex,
        int LineIndex,
        int SpanIndex,
        int LineSpanCount,
        double FrameFontScale,
        double FrameShapeX,
        double FrameShapeTopY,
        double FrameShapeWidth,
        double FrameShapeHeight,
        int? TableRowIndex,
        int? TableColumnIndex,
        int? TableRowSpan,
        int? TableColumnSpan,
        double FrameInsetLeft,
        double FrameInsetRight,
        double FrameInsetTop,
        double FrameInsetBottom,
        string FrameWrapMode,
        string? FrameWrapValue,
        string FrameVerticalOverflowMode,
        string? FrameVerticalOverflowValue,
        string FrameVerticalOverflowSource,
        string FrameAutofitMode,
        double FrameTextX,
        double FrameTextWidth,
        double FrameTextWrapWidth,
        double FrameTextHeight,
        double FrameClipX,
        double FrameClipWidth,
        double FrameClipY,
        double FrameClipHeight,
        int FrameColumnCount,
        double FrameColumnSpacing,
        double LineTopY,
        double LineAdvance,
        double LineMaxFontSize);

    private sealed record TextFontUse(
        string FamilyName,
        bool Bold,
        bool Italic,
        IReadOnlyList<int> CodePoints);

    private sealed record TextGlyphAtom(
        int CodePoint,
        string? Typeface,
        PptxGlyphTypefaceResolutionSource TypefaceResolutionSource,
        ushort GlyphId,
        double Advance,
        double AdjustmentBefore);

    internal readonly record struct TextOutline(RgbColor Color, double Alpha, double Width);

    private readonly record struct TextCapsFragment(string Text, double FontScale);

    private readonly record struct PptxTextLineMetrics(double BaselineOffset, double LineAdvance, string Source);

    private readonly record struct PptxTextFlowSegment(
        string Text,
        string AdvanceText,
        PptxTextFlowSegmentKind Kind,
        bool Draw,
        bool PreventCoalesce,
        double FontScale);

    private readonly record struct ResolvedParagraphTextStyle(
        TextAlignment Alignment,
        string? AlignmentValue,
        XElement? Properties,
        XElement? DefaultRunProperties,
        double FontSize,
        double SpacingBefore,
        double SpacingAfter,
        LineSpacing LineSpacing,
        ParagraphIndent Indent,
        IReadOnlyList<double> TabStops);

    private readonly record struct ResolvedEndParagraphTextStyle(
        double FontSize,
        string? Typeface,
        bool Bold,
        bool Italic);

    internal readonly record struct ResolvedRunTextStyle(
        double NominalFontSize,
        double FontSize,
        double CharacterSpacing,
        double BaselineOffset,
        RgbColor Color,
        PptxRunTextColorSource ColorSource,
        double Alpha,
        TextOutline? Outline,
        RgbColor? Highlight,
        bool HasHyperlinkClick,
        string? HyperlinkClickId,
        string? HyperlinkClickAction,
        bool Bold,
        bool Italic,
        bool Underline,
        string? UnderlineValue,
        bool Strike,
        string? StrikeValue,
        string? CapsValue,
        bool KerningEnabled,
        PptxThemeTypefaceSource TypefaceSource,
        string? Typeface);

    internal enum PptxRunTextColorSource
    {
        RunNoFill,
        RunSolidFill,
        TableTextStyle,
        ThemeHyperlink,
        ShapeFontRef,
        DefaultNoFill,
        DefaultSolidFill,
        FallbackBlack
    }

    private sealed record PptxTextFrameModel(
        XElement Shape,
        XElement TextBody,
        XElement? InheritedTextBody,
        int InheritedPlaceholderCount,
        bool UsesInheritedShapeBounds,
        PptxTheme Theme,
        PptxTextBodyProperties BodyProperties,
        ShapeBounds Bounds,
        int? TableRowIndex,
        int? TableColumnIndex,
        int? TableRowSpan,
        int? TableColumnSpan,
        double? TableDeclaredRowHeight,
        double? TableDeclaredRowSpanHeight,
        double? TableDeclaredHeight,
        double? TableHeightSlackFactor,
        TextInsets Insets,
        double FontScale,
        double LineSpacingScale,
        double TextX,
        double TextWidth,
        double TextWrapWidth,
        double TextHeight,
        double TextClipX,
        double TextClipWidth,
        double TextClipY,
        double TextClipHeight,
        int ColumnCount,
        double ColumnSpacing,
        double RotationCenterX,
        double RotationCenterY,
        double TextRotationDegrees,
        bool TextFlipHorizontal,
        bool TextFlipVertical,
        bool UseOfficeBaselineFloor,
        double FlowYTop,
        double VerticalOffset,
        PptxTextOrientation Orientation,
        RgbColor? ShapeFontColor,
        IReadOnlyList<PptxTextParagraphModel> Paragraphs);

    internal sealed record PptxTableCellTextFrame(
        XElement TextBody,
        double X,
        double Y,
        double Width,
        double Height,
        int RowIndex,
        int ColumnIndex,
        int RowSpan,
        int ColumnSpan,
        double DeclaredRowHeight,
        double DeclaredRowSpanHeight,
        double DeclaredTableHeight,
        double TableHeightSlackFactor,
        TextInsets Insets,
        TextInsetSources InsetSources,
        TextInsetValues InsetValues,
        TextVerticalAnchor VerticalAnchor,
        string? VerticalAnchorValue,
        PptxTextBodyPropertySource VerticalAnchorSource,
        PptxColorMap ColorMap,
        PptxSceneTableCellTextStyle TextStyle);

    private readonly record struct PptxTextBodyProperties(
        TextInsets Insets,
        TextInsetSources InsetSources,
        TextInsetValues InsetValues,
        PptxTextOrientation Orientation,
        string? OrientationValue,
        PptxTextBodyPropertySource OrientationSource,
        TextVerticalAnchor VerticalAnchor,
        string? VerticalAnchorValue,
        PptxTextBodyPropertySource VerticalAnchorSource,
        bool? AnchorCenter,
        string? AnchorCenterValue,
        PptxTextBodyPropertySource AnchorCenterSource,
        PptxTextWrapMode WrapMode,
        string? WrapValue,
        PptxTextBodyPropertySource WrapSource,
        PptxTextVerticalOverflow VerticalOverflow,
        string? VerticalOverflowValue,
        PptxTextBodyPropertySource VerticalOverflowSource,
        int ColumnCount,
        double ColumnSpacing,
        PptxTextBodyPropertySource ColumnSource,
        PptxTextBodyPropertySource ColumnCountSource,
        PptxTextBodyPropertySource ColumnSpacingSource,
        string? ColumnCountValue,
        string? ColumnSpacingValue,
        string AutofitModeValue,
        PptxTextBodyPropertySource AutofitModeSource,
        double FontScale,
        string? FontScaleValue,
        PptxTextBodyPropertySource FontScaleSource,
        double LineSpacingScale,
        string? LineSpacingReductionValue,
        PptxTextBodyPropertySource LineSpacingScaleSource,
        bool CompatibleLineSpacing,
        string? CompatibleLineSpacingValue,
        PptxTextBodyPropertySource CompatibleLineSpacingSource,
        double? RotationDegrees,
        string? RotationValue,
        PptxTextBodyPropertySource RotationDegreesSource,
        double? ExplicitWrapWidth)
    {
        public PptxTextAutofitMode AutofitMode { get; init; } = PptxTextAutofitMode.Absent;
    }

    internal enum PptxTextBodyPropertySource
    {
        DirectBodyPr,
        InheritedBodyPr,
        TableCellProperties,
        TableCellStyle,
        DefaultValue
    }

    internal readonly record struct TextInsetSources(
        PptxTextBodyPropertySource Left,
        PptxTextBodyPropertySource Right,
        PptxTextBodyPropertySource Top,
        PptxTextBodyPropertySource Bottom);

    internal readonly record struct TextInsetValues(
        string? Left,
        string? Right,
        string? Top,
        string? Bottom);

    private sealed record PptxTextParagraphModel(
        XElement Source,
        XElement? Properties,
        XElement? EndParagraphProperties,
        ResolvedEndParagraphTextStyle EndParagraphStyle,
        double EmptySpacingBefore,
        double EmptySpacingAfter,
        bool HasLayoutContent,
        bool HasVisibleContent,
        bool HasManualLineBreak,
        double FirstLineFallbackFontSize,
        XElement? DefaultProperties,
        int Level,
        PptxParagraphStyleCascade Cascade,
        PptxParagraphStyleCascade ResolvedStyleCascade,
        ResolvedParagraphTextStyle Style,
        PptxParagraphBulletModel Bullet,
        IReadOnlyList<PptxTextRunModel> Runs);

    private sealed record PptxParagraphBulletModel(
        PptxParagraphBulletKind Kind,
        string? Character,
        string? ResolvedCharacter,
        string? AutoNumberType,
        string? AutoNumberStartAtValue,
        int? AutoNumberStartAt,
        string? FontTypeface,
        string? FontCharset,
        string? ResolvedFontTypeface,
        PptxThemeTypefaceSource FontTypefaceSource,
        RgbColor? Color,
        PptxParagraphBulletSizeKind SizeKind,
        string? SizeValue);

    internal enum PptxParagraphBulletKind
    {
        None,
        Character,
        AutoNumber,
        Blip
    }

    private enum PptxParagraphBulletSizeKind
    {
        Text,
        Percent,
        Points
    }

    private sealed record PptxParagraphStyleCascade(
        string LevelName,
        IReadOnlyList<PptxParagraphStyleLayer> Layers)
    {
        public IReadOnlyList<XElement?> Sources => Layers.Select(layer => layer.Source).ToArray();

        public XElement? ResolveDefaultProperties()
        {
            return MergeParagraphProperties(Sources.ToArray());
        }
    }

    private enum PptxParagraphStyleLayerKind
    {
        ShapeListStyle,
        InheritedPlaceholderListStyle,
        MasterPlaceholderListStyle,
        LayoutPlaceholderListStyle,
        InheritedTextStyle,
        DefaultTextStyle,
        ParagraphProperties
    }

    private sealed record PptxParagraphStyleLayer(
        string Name,
        PptxParagraphStyleLayerKind Kind,
        XElement? Source);

    internal sealed record PptxTextRunModel(
        int RunIndex,
        PptxTextRunKind Kind,
        XElement Source,
        XElement? Properties,
        PptxRunStyleCascade Cascade,
        string Text,
        ResolvedRunTextStyle Style);

    internal sealed record PptxRunStyleCascade(
        IReadOnlyList<PptxRunStyleLayer> Layers,
        XElement? ResolvedDefaultProperties)
    {
        public IReadOnlyList<XElement?> Sources => Layers.Select(layer => layer.Source).ToArray();

        public XElement? DirectProperties => Layers.FirstOrDefault(layer => layer.Kind == PptxRunStyleLayerKind.RunProperties)?.Source;
    }

    internal enum PptxRunStyleLayerKind
    {
        RunProperties,
        ParagraphDefaultRunProperties,
        ParagraphPropertiesDefaultRunProperties,
        ShapeListStyleDefaultRunProperties,
        InheritedPlaceholderDefaultRunProperties,
        MasterPlaceholderDefaultRunProperties,
        LayoutPlaceholderDefaultRunProperties,
        InheritedTextStyleDefaultRunProperties,
        DefaultTextStyleDefaultRunProperties
    }

    internal sealed record PptxRunStyleLayer(
        string Name,
        PptxRunStyleLayerKind Kind,
        XElement? Source);
}
