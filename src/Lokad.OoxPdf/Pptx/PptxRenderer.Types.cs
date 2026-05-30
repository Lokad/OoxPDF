using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private readonly record struct ShapeBounds(
        long X,
        long Y,
        long Width,
        long Height,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical);

    private readonly record struct GroupTransform(
        long OffsetX,
        long OffsetY,
        long Width,
        long Height,
        long ChildOffsetX,
        long ChildOffsetY,
        double ScaleX,
        double ScaleY,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical)
    {
        public static GroupTransform Identity { get; } = new(0, 0, 0, 0, 0, 0, 1d, 1d, 0d, FlipHorizontal: false, FlipVertical: false);

        public ShapeBounds Apply(ShapeBounds bounds)
        {
            long width = (long)Math.Round(bounds.Width * ScaleX);
            long height = (long)Math.Round(bounds.Height * ScaleY);
            long localX = (long)Math.Round((bounds.X - ChildOffsetX) * ScaleX);
            long localY = (long)Math.Round((bounds.Y - ChildOffsetY) * ScaleY);
            long x = FlipHorizontal
                ? OffsetX + Width - localX - width
                : OffsetX + localX;
            long y = FlipVertical
                ? OffsetY + Height - localY - height
                : OffsetY + localY;
            double rotationDegrees = NormalizeRotationDegrees(bounds.RotationDegrees + RotationDegrees);
            if (Math.Abs(RotationDegrees) > PptxTextMetricRules.TextStateTolerance && Width > 0 && Height > 0)
            {
                double groupCenterX = OffsetX + Width / 2d;
                double groupCenterY = OffsetY + Height / 2d;
                double boundsCenterX = x + width / 2d;
                double boundsCenterY = y + height / 2d;
                double radians = RotationDegrees * Math.PI / 180d;
                double cos = Math.Cos(radians);
                double sin = Math.Sin(radians);
                double dx = boundsCenterX - groupCenterX;
                double dy = boundsCenterY - groupCenterY;
                double rotatedCenterX = groupCenterX + dx * cos - dy * sin;
                double rotatedCenterY = groupCenterY + dx * sin + dy * cos;
                x = (long)Math.Round(rotatedCenterX - width / 2d);
                y = (long)Math.Round(rotatedCenterY - height / 2d);
            }

            return new ShapeBounds(
                x,
                y,
                width,
                height,
                rotationDegrees,
                bounds.FlipHorizontal ^ FlipHorizontal,
                bounds.FlipVertical ^ FlipVertical);
        }

        public GroupTransform Combine(GroupTransform child)
        {
            ShapeBounds childBounds = Apply(new ShapeBounds(
                child.OffsetX,
                child.OffsetY,
                child.Width,
                child.Height,
                child.RotationDegrees,
                FlipHorizontal: false,
                FlipVertical: false));
            return new GroupTransform(
                childBounds.X,
                childBounds.Y,
                childBounds.Width,
                childBounds.Height,
                child.ChildOffsetX,
                child.ChildOffsetY,
                ScaleX * child.ScaleX,
                ScaleY * child.ScaleY,
                childBounds.RotationDegrees,
                FlipHorizontal ^ child.FlipHorizontal,
                FlipVertical ^ child.FlipVertical);
        }
    }

    private readonly record struct BezierSegment(
        double StartX,
        double StartY,
        double Control1X,
        double Control1Y,
        double Control2X,
        double Control2Y,
        double EndX,
        double EndY);

    private readonly record struct CurveSample(
        double X,
        double Y,
        double TangentX,
        double TangentY);

    private readonly record struct CurvedConnectorFillPath(
        IReadOnlyList<(double X, double Y)> Points,
        IReadOnlyList<(double X, double Y)>? TailSubpath,
        double TipX,
        double TipY,
        double DirectionX,
        double DirectionY,
        double NormalX,
        double NormalY);

    private readonly record struct TextRun(
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
        bool PreventCoalesce = false,
        TextOutline? Outline = null,
        bool StrictClip = false);

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

    private sealed record PptxPositionedTextSpan(
        PptxTextRunModel? SourceRun,
        PptxTextLineBoxLayout? LineBox,
        int FrameIndex,
        int ParagraphIndex,
        int? SourceRunIndex,
        string ParagraphBulletKind,
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

    private readonly record struct TextOutline(RgbColor Color, double Alpha, double Width);

    private readonly record struct TextCapsFragment(string Text, double FontScale);

    private readonly record struct PptxTextLineMetrics(double BaselineOffset, double LineAdvance, string Source);

    private readonly record struct PptxTextFlowSegment(
        string Text,
        string AdvanceText,
        PptxTextFlowSegmentKind Kind,
        bool Draw,
        bool PreventCoalesce,
        double FontScale = 1d);

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

    private readonly record struct ResolvedRunTextStyle(
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

    private enum PptxRunTextColorSource
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

    private sealed record PptxTableCellTextFrame(
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
        double? ExplicitWrapWidth);

    private enum PptxTextBodyPropertySource
    {
        DirectBodyPr,
        InheritedBodyPr,
        TableCellProperties,
        TableCellStyle,
        DefaultValue
    }

    private readonly record struct TextInsetSources(
        PptxTextBodyPropertySource Left,
        PptxTextBodyPropertySource Right,
        PptxTextBodyPropertySource Top,
        PptxTextBodyPropertySource Bottom);

    private readonly record struct TextInsetValues(
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

    private enum PptxParagraphBulletKind
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

    private sealed record PptxTextRunModel(
        int RunIndex,
        PptxTextRunKind Kind,
        XElement Source,
        XElement? Properties,
        PptxRunStyleCascade Cascade,
        string Text,
        ResolvedRunTextStyle Style);

    private sealed record PptxRunStyleCascade(
        IReadOnlyList<PptxRunStyleLayer> Layers,
        XElement? ResolvedDefaultProperties)
    {
        public IReadOnlyList<XElement?> Sources => Layers.Select(layer => layer.Source).ToArray();

        public XElement? DirectProperties => Layers.FirstOrDefault(layer => layer.Kind == PptxRunStyleLayerKind.RunProperties)?.Source;
    }

    private enum PptxRunStyleLayerKind
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

    private sealed record PptxRunStyleLayer(
        string Name,
        PptxRunStyleLayerKind Kind,
        XElement? Source);

    private sealed record PptxTextFlowModel(IReadOnlyList<PptxTextFlowFrame> Frames);

    private sealed record PptxTextFlowFrame(
        PptxTextFrameModel Model,
        PptxTextFlowBox Box,
        IReadOnlyList<PptxTextFlowParagraph> Paragraphs);

    private sealed record PptxTextFlowBox(
        double YTop,
        double CursorTop,
        double TextX,
        double TextWidth,
        double TextWrapWidth,
        double TextHeight,
        double ClipX,
        double ClipWidth,
        double ClipY,
        double ClipHeight,
        double RotationCenterX,
        double RotationCenterY);

    private sealed record PptxTextFlowParagraph(
        PptxTextParagraphModel Model,
        ResolvedParagraphTextStyle Style,
        IReadOnlyList<PptxTextFlowRun> Runs);

    private sealed record PptxTextFlowRun(
        PptxTextRunModel Source,
        ResolvedRunTextStyle Style,
        IReadOnlyList<PptxTextFlowSegment> Segments);

    private sealed record PptxTextLayoutModel(IReadOnlyList<PptxTextFrameLayout> Frames);

    private sealed record PptxTextFrameLayout(
        PptxTextFrameModel Model,
        IReadOnlyList<PptxTextParagraphLayout> Paragraphs);

    private sealed record PptxTextParagraphLayout(
        PptxTextParagraphModel Model,
        IReadOnlyList<PptxTextLineLayout> Lines);

    private sealed record PptxTextLineLayout(
        PptxTextLineBoxLayout Box,
        double StartX,
        double EndX,
        double NaturalEndX,
        TextAlignment Alignment,
        IReadOnlyList<PptxTextSpanLayout> Spans);

    private sealed record PptxTextLineBoxLayout(
        double TopY,
        double BaselineY,
        double Advance,
        double BaselineOffset,
        double MaxFontSize,
        LineSpacing LineSpacing,
        PptxTextBaselineMetricLayout BaselineMetric);

    private sealed record PptxTextBaselineMetricLayout(
        string Source,
        string? Typeface,
        bool Bold,
        bool Italic,
        double FontSize,
        double Ratio,
        int UnitsPerEm,
        int WindowsAscender,
        int WindowsDescender,
        int TypographicAscender,
        int TypographicDescender,
        int TypographicLineGap);

    private sealed record PptxTextSpanLayout(
        PptxTextRunModel? SourceRun,
        TextRun Run,
        double EndX,
        IReadOnlyList<PptxTextAtomLayout> Atoms,
        PptxTextGlyphSpanLayout GlyphSpan);

    private sealed record PptxTextGlyphSpanLayout(
        string Text,
        string? Typeface,
        bool Bold,
        bool Italic,
        double FontSize,
        double CharacterSpacing,
        bool KerningEnabled,
        double LeadingAdjustment,
        double NaturalWidth,
        double LayoutWidth,
        IReadOnlyList<PptxTextGlyphLayout> Glyphs)
    {
        public static PptxTextGlyphSpanLayout Empty(TextRun run)
        {
            return new PptxTextGlyphSpanLayout(
                run.Text,
                run.FontFamily,
                run.Bold,
                run.Italic,
                run.FontSize,
                run.CharacterSpacing,
                run.KerningEnabled,
                0d,
                run.Width,
                run.Width,
                []);
        }
    }

    private sealed record PptxTextGlyphLayout(
        int CodePoint,
        string? Typeface,
        PptxGlyphTypefaceResolutionSource TypefaceResolutionSource,
        ushort GlyphId,
        double Advance,
        double AdjustmentBefore);

    private sealed record PptxTextAtomLayout(
        PptxTextAtomKind Kind,
        string Text,
        double X,
        double Width,
        bool Draw);

    private enum PptxTextAtomKind
    {
        Word,
        Space,
        Tab,
        HiddenAdvance
    }

    private enum PptxTextFlowSegmentKind
    {
        Text,
        Tab,
        HiddenAdvance,
        NoBreakHiddenAdvance,
        BoundaryPunctuation,
        Break
    }

    private enum PptxTextRunKind
    {
        Text,
        Break,
        Field
    }

    private sealed class TextLayoutLine(double startX)
    {
        private double startX = startX;

        public List<PptxTextSpanLayout> Spans { get; } = [];

        public double EndX { get; private set; } = startX;

        public void Add(PptxTextRunModel? sourceRun, TextRun run, double endX, IReadOnlyList<PptxTextAtomLayout>? atoms = null, PptxTextGlyphSpanLayout? glyphSpan = null)
        {
            Spans.Add(new PptxTextSpanLayout(
                sourceRun,
                run,
                endX,
                atoms ?? [new PptxTextAtomLayout(PptxTextAtomKind.Word, run.Text, run.X, run.Width, Draw: true)],
                glyphSpan ?? PptxTextGlyphSpanLayout.Empty(run)));
            AdvanceTo(endX);
        }

        public void AdvanceTo(double x)
        {
            EndX = Math.Max(EndX, x);
        }

        public bool TryRemoveLastSpan([NotNullWhen(true)] out PptxTextSpanLayout? span)
        {
            if (Spans.Count == 0)
            {
                span = null;
                return false;
            }

            int lastIndex = Spans.Count - 1;
            span = Spans[lastIndex];
            Spans.RemoveAt(lastIndex);
            EndX = Spans.Count == 0 ? startX : Spans.Max(item => item.EndX);
            return true;
        }

        public void Reset(double startX)
        {
            this.startX = startX;
            Spans.Clear();
            EndX = startX;
        }
    }

    private readonly record struct TextInsets(double Left, double Right, double Top, double Bottom)
    {
        public static TextInsets Empty { get; } = new(0d, 0d, 0d, 0d);

        public bool IsEmpty => Left == 0d && Right == 0d && Top == 0d && Bottom == 0d;
    }

    private readonly record struct ParagraphIndent(double MarginLeft, double Hanging);

    private readonly record struct RenderedFont(string ResourceName, PdfEmbeddedFont Font, bool SyntheticBold, bool SyntheticItalic);

    private readonly record struct RenderedFonts(IReadOnlyDictionary<string, RenderedFont> Fonts, IReadOnlyList<PdfFontResource> Resources);

    private readonly record struct BulletStyle(double FontSize, RgbColor Color, string? Typeface);

    private readonly record struct LineSpacing(double Value, bool IsAbsolute, bool IsExplicit, bool UseNormalLineAdvance)
    {
        public static LineSpacing Absolute(double points) => new(points, true, true, false);

        public static LineSpacing Multiple(double factor, bool isExplicit, bool useNormalLineAdvance = true) => new(factor, false, isExplicit, useNormalLineAdvance);

        public LineSpacing ScaleExplicit(double factor)
        {
            return IsExplicit
                ? new LineSpacing(Value * factor, IsAbsolute, IsExplicit, UseNormalLineAdvance)
                : this;
        }

        public double Resolve(double fontSize)
        {
            return IsAbsolute ? Value : fontSize * Value;
        }
    }

    private sealed class TextAdvanceEstimator
    {
        private readonly PresentationFontResolver resolver;
        private readonly Dictionary<string, OpenTypeFont?> fonts = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, FontResolution?> resolutions = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ResolvedGlyphFont?> glyphFonts = new(StringComparer.OrdinalIgnoreCase);

        public TextAdvanceEstimator(PresentationFontResolver? resolver = null)
        {
            this.resolver = resolver ?? new PresentationFontResolver();
        }

        public double Measure(string text, double fontSize, string? familyName, bool bold = false, bool italic = false, double characterSpacing = 0d, bool kerningEnabled = true)
        {
            if (fontSize <= 0d)
            {
                return 0d;
            }

            double points = 0d;
            OpenTypeFont? previousFont = null;
            ushort previousGlyph = 0;
            int runeCount = 0;
            foreach (Rune rune in text.EnumerateRunes())
            {
                runeCount++;
                ResolvedGlyphFont? resolved = ResolveGlyphFont(familyName, bold, italic, rune.Value);
                if (resolved is null)
                {
                    continue;
                }

                OpenTypeFont font = resolved.Font;
                ushort glyph = font.MapCodePoint(rune.Value);
                if (glyph == 0)
                {
                    continue;
                }

                if (kerningEnabled && previousFont == font && previousGlyph != 0)
                {
                    points += font.GetKerning(previousGlyph, glyph) * fontSize / font.UnitsPerEm;
                }

                if (previousFont == font && previousGlyph != 0 && resolved.SyntheticBold)
                {
                    points -= PptxTextMetricRules.OfficeSyntheticBoldAdvanceTightening(fontSize);
                }

                points += font.GetAdvanceWidth(glyph) * fontSize / font.UnitsPerEm;
                previousFont = font;
                previousGlyph = glyph;
            }

            if (previousFont is null)
            {
                return PptxTextMetricRules.FallbackAdvanceWidth(text.Length, runeCount, fontSize, characterSpacing);
            }

            return Math.Max(0d, points + Math.Max(0, runeCount - 1) * characterSpacing);
        }

        public double MeasureBoundaryAdvance(int previousCodePoint, int nextCodePoint, double fontSize, string? familyName, bool bold = false, bool italic = false, double characterSpacing = 0d, bool kerningEnabled = true)
        {
            ResolvedGlyphFont? previousResolved = ResolveGlyphFont(familyName, bold, italic, previousCodePoint);
            ResolvedGlyphFont? nextResolved = ResolveGlyphFont(familyName, bold, italic, nextCodePoint);
            if (previousResolved is null || nextResolved is null || previousResolved.Font != nextResolved.Font)
            {
                return characterSpacing;
            }

            OpenTypeFont font = previousResolved.Font;
            ushort previousGlyph = font.MapCodePoint(previousCodePoint);
            ushort nextGlyph = font.MapCodePoint(nextCodePoint);
            double units = kerningEnabled && previousGlyph != 0 && nextGlyph != 0
                ? font.GetKerning(previousGlyph, nextGlyph)
                : 0d;
            return units * fontSize / font.UnitsPerEm + characterSpacing;
        }

        public ResolvedGlyphFont? ResolveGlyphFont(string? familyName, bool bold, bool italic, int codePoint)
        {
            string requestedFamily = PptxFontFallbackRules.ResolveDefaultLatinTypeface(familyName);
            string key = requestedFamily + "\u001f" + bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + italic.ToString(CultureInfo.InvariantCulture) + "\u001f" + codePoint.ToString(CultureInfo.InvariantCulture);
            if (glyphFonts.TryGetValue(key, out ResolvedGlyphFont? cached))
            {
                return cached;
            }

            FontResolution? primaryResolution = ResolveFontResolution(requestedFamily, bold, italic);
            OpenTypeFont? primaryFont = LoadFont(primaryResolution);
            if (primaryResolution is not null && primaryFont is not null && primaryFont.MapCodePoint(codePoint) != 0)
            {
                cached = new ResolvedGlyphFont(requestedFamily, requestedFamily, PptxGlyphTypefaceResolutionSource.Primary, primaryFont, bold && !primaryResolution.Bold);
                glyphFonts[key] = cached;
                return cached;
            }

            foreach (FontResolution resolution in resolver.GetDiscoveredFonts()
                         .Where(f => !f.HasMathTable && f.FontFilePath is not null)
                         .OrderBy(f => f.Bold == bold ? 0 : 1000)
                         .ThenBy(f => f.Italic == italic ? 0 : 1000)
                         .ThenBy(f => Math.Abs(f.WeightClass - (bold ? 700 : 400)))
                         .ThenBy(f => f.FamilyName, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(f => f.FontFilePath, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(f => f.FontFaceIndex))
            {
                if (primaryResolution is not null &&
                    resolution.FontFaceIndex == primaryResolution.FontFaceIndex &&
                    string.Equals(resolution.FontFilePath, primaryResolution.FontFilePath, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                OpenTypeFont? font = LoadFont(resolution);
                if (font is not null && font.MapCodePoint(codePoint) != 0)
                {
                    cached = new ResolvedGlyphFont(requestedFamily, resolution.FamilyName, PptxGlyphTypefaceResolutionSource.Fallback, font, bold && !resolution.Bold);
                    glyphFonts[key] = cached;
                    return cached;
                }
            }

            glyphFonts[key] = null;
            return null;
        }

        public OpenTypeFont? ResolveOpenTypeFont(string? familyName, bool bold = false, bool italic = false)
        {
            return ResolveFont(PptxFontFallbackRules.ResolveDefaultLatinTypeface(familyName), bold, italic);
        }

        public bool RequestedStyleRequiresSyntheticBold(string? familyName, bool bold, bool italic)
        {
            if (!bold)
            {
                return false;
            }

            FontResolution? resolution = ResolveFontResolution(PptxFontFallbackRules.ResolveDefaultLatinTypeface(familyName), bold, italic);
            return resolution is not null && !resolution.Bold;
        }

        public bool RequestedStyleRequiresSyntheticItalic(string? familyName, bool bold, bool italic)
        {
            if (!italic)
            {
                return false;
            }

            FontResolution? resolution = ResolveFontResolution(PptxFontFallbackRules.ResolveDefaultLatinTypeface(familyName), bold, italic);
            return resolution is not null && !resolution.Italic;
        }

        private OpenTypeFont? ResolveFont(string familyName, bool bold, bool italic)
        {
            return LoadFont(ResolveFontResolution(familyName, bold, italic));
        }

        private FontResolution? ResolveFontResolution(string familyName, bool bold, bool italic)
        {
            string key = familyName + "\u001f" + bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + italic.ToString(CultureInfo.InvariantCulture);
            if (resolutions.TryGetValue(key, out FontResolution? cached))
            {
                return cached;
            }

            try
            {
                cached = resolver.ResolvePresentationTextFace(new FontRequest(familyName, bold, italic));
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
            {
                cached = null;
            }

            resolutions[key] = cached;
            return cached;
        }

        private OpenTypeFont? LoadFont(FontResolution? resolution)
        {
            if (resolution is null || resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
            {
                return null;
            }

            string key = resolution.FontFilePath + "\u001f" + resolution.FontFaceIndex.ToString(CultureInfo.InvariantCulture);
            if (fonts.TryGetValue(key, out OpenTypeFont? cached))
            {
                return cached;
            }

            try
            {
                cached = OpenTypeFont.Load(resolution.FontFilePath, resolution.FontFaceIndex);
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
            {
                cached = null;
            }

            fonts[key] = cached;
            return cached;
        }
    }

    private sealed record ResolvedGlyphFont(
        string RequestedTypeface,
        string Typeface,
        PptxGlyphTypefaceResolutionSource Source,
        OpenTypeFont Font,
        bool SyntheticBold);

    private enum PptxGlyphTypefaceResolutionSource
    {
        Primary,
        Fallback
    }

    private enum TextAlignment
    {
        Left,
        Center,
        Right,
        Justify,
        Distributed,
        JustLow,
        ThaiDistributed
    }

    private static class PptxTextMetricRules
    {
        public const double CoordinateTolerance = 0.01d;
        public const double TextStateTolerance = 0.001d;
        public const double MinimumDrawableDimension = 1d;
        public const double MinimumStrokeWidth = 0.5d;
        public const double OfficeSuperscriptSubscriptScale = 2d / 3d;
        public const double CssNormalLineHeightFallback = 1.2d;
        public const double OfficeCompatibleDefaultLineSpacingFactor = 1.1d;
        public const double OfficeCompatibleNoAutoFitDefaultLineSpacingFactor = 1.2d;
        public const double MiddleVerticalAnchorSlackMultiplier = 0.5d;
        public const double OfficeManualBreakDefaultLineHeightFallback = 1.24d;
        public const double OfficeManualBreakBaselineFallback = 0.9344d;
        public const double OfficeBaselineFallback = 0.974d;
        public const double MinimumBaselineMetricRatio = 0.75d;
        public const double MaximumBaselineMetricRatio = 1.05d;
        public const double MaximumOfficeBaselineWindowsAscenderRatio = 1d;
        public const double MinimumOfficeTypographicBaselineAscenderRatio = 0.93d;
        public const double MaximumOfficeTypographicBaselineFontSize = 20d;
        public const double OfficeBaselineFloorMetricThreshold = 0.94d;
        public const double OfficeBaselineFloorMaximumWindowsDescenderRatio = 0.24d;
        public const double MinimumFontLineBoxMetricRatio = 0.75d;
        public const double MaximumFontLineBoxMetricRatio = 1.5d;
        public const double AbsoluteLineBaselineGapFallback = 0.374d;
        public const double ExplicitLineBaselineGapFallback = 0.234d;
        public const double MinimumLineSpacing = 0.1d;
        public const double MinimumAutofitScale = 0.01d;
        public const double MaximumAutofitScale = 10d;
        public const double MaximumLineSpacingReduction = 0.99d;
        public const double SuperscriptSubscriptMinimumBaselineRatio = 0.2d;
        public const double SmallCapsFallbackScale = 0.8d;
        public const double DefaultTextOutlineWidth = 0.75d;
        public const double SyntheticBoldStrokeWidthRatio = 1d / 35d;
        public const double OfficeSyntheticBoldAdvanceTighteningEm = 0.007d;
        public const double OfficeSyntheticBoldItalicCharacterSpacingEm = 0.01545d;
        public const double OfficeHighlightContinuationCharacterSpacingEm = -0.003d;
        public const double OfficeAutofitNumberedDenseCharacterSpacingEm = -0.002d;
        public const double OfficeAutofitNumberedDefaultCharacterSpacingEm = -0.004d;
        public const double OfficeStrikePositionFontScale = 0.211d;
        public const double StrikeThicknessFallback = 0.05d;
        public const double HighlightDescenderPaddingFontUnits = 32d;
        public const double HighlightMaximumDescentFontScale = 0.23d;
        public const double HighlightMaximumHeightFontScale = 1.18d;
        public const double AdjacentTextCoalesceGapFontScale = 0.2d;
        public const double AdjacentUnderlineCoalesceGapFontScale = 0.08d;
        public const double WrapFitToleranceFontScale = 0.16d;
        public const double BulletWrapFitToleranceFontScale = 0.2d;
        public const double CenteredTableCellWrapToleranceWidthScale = 0.02d;
        public const double FinalWordWrapToleranceWidthScale = 0.02d;
        public const double ShapeAutoFitWrapToleranceWidthScale = 0.011d;
        public const double FallbackAdvanceWidthScale = 0.42d;
        public const double EllipseTextRectInsetRatio = 0.1464466094067262d;
        public const double RoundRectDefaultAdjustment = 16667d;
        public const double RoundRectTextRectRadiusInsetFactor = 0.2928932188134525d;
        public const int ShapeAutoFitSearchIterations = 10;

        public static double ClampNonNegative(double value) => Math.Max(0d, value);

        public static double MinimumWidth(double width) => Math.Max(MinimumDrawableDimension, width);

        public static double SuperscriptSubscriptFontSize(double nominalFontSize) => nominalFontSize * OfficeSuperscriptSubscriptScale;

        public static bool ShouldScaleSuperscriptSubscript(double baselineOffset, double nominalFontSize) =>
            Math.Abs(baselineOffset) >= nominalFontSize * SuperscriptSubscriptMinimumBaselineRatio;

        public static double SmallCapsFontScale() => SmallCapsFallbackScale;

        public static double TextOutlineWidth(double? width) => Math.Max(MinimumStrokeWidth, width ?? DefaultTextOutlineWidth);

        public static double SyntheticBoldStrokeWidth(double fontSize) => Math.Max(0d, fontSize * SyntheticBoldStrokeWidthRatio);

        public static double OfficeSyntheticBoldAdvanceTightening(double fontSize) =>
            Math.Max(0d, fontSize * OfficeSyntheticBoldAdvanceTighteningEm);

        public static double OfficeSyntheticBoldItalicCharacterSpacing(double fontSize) =>
            Math.Max(0d, fontSize * OfficeSyntheticBoldItalicCharacterSpacingEm);

        public static double OfficeHighlightContinuationCharacterSpacing(double fontSize) =>
            fontSize * OfficeHighlightContinuationCharacterSpacingEm;

        public static double OfficeAutofitNumberedDenseCharacterSpacing(double fontSize) =>
            fontSize * OfficeAutofitNumberedDenseCharacterSpacingEm;

        public static double OfficeAutofitNumberedDefaultCharacterSpacing(double fontSize) =>
            fontSize * OfficeAutofitNumberedDefaultCharacterSpacingEm;

        public static double StrikeY(PdfEmbeddedFont embedded, double baselineY, double fontSize)
        {
            return baselineY + fontSize * OfficeStrikePositionFontScale;
        }

        public static double StrikeThickness(PdfEmbeddedFont embedded, double fontSize)
        {
            double fontScale = fontSize / embedded.Font.UnitsPerEm;
            double strikeoutSize = embedded.Font.Os2.StrikeoutSize * fontScale;
            return Math.Max(MinimumStrokeWidth, strikeoutSize > 0d ? strikeoutSize : fontSize * StrikeThicknessFallback);
        }

        public static double HighlightDescent(PdfEmbeddedFont embedded, double fontSize, double fontScale)
        {
            double metricDescent = (embedded.Font.Os2.WindowsDescender + HighlightDescenderPaddingFontUnits) * fontScale;
            return Math.Min(metricDescent, fontSize * HighlightMaximumDescentFontScale);
        }

        public static double HighlightHeight(PdfEmbeddedFont embedded, double fontSize, double fontScale)
        {
            double metricHeight = (embedded.Font.Os2.WindowsAscender + embedded.Font.Os2.WindowsDescender) * fontScale;
            return Math.Min(metricHeight, fontSize * HighlightMaximumHeightFontScale);
        }

        public static double TextCoalesceGap(double fontSize) => Math.Max(MinimumDrawableDimension, fontSize * AdjacentTextCoalesceGapFontScale);

        public static double UnderlineCoalesceGap(double fontSize) => Math.Max(MinimumDrawableDimension, fontSize * AdjacentUnderlineCoalesceGapFontScale);

        public static double WrapFitTolerance(double fontSize) => Math.Max(CoordinateTolerance, fontSize * WrapFitToleranceFontScale);

        public static double BulletWrapFitTolerance(double fontSize) => Math.Max(WrapFitTolerance(fontSize), fontSize * BulletWrapFitToleranceFontScale);

        public static double CenteredTableCellWrapTolerance(double fontSize, double availableWidth) => Math.Max(WrapFitTolerance(fontSize), availableWidth * CenteredTableCellWrapToleranceWidthScale);

        public static double FinalWordWrapTolerance(double fontSize, double availableWidth) => Math.Max(WrapFitTolerance(fontSize), availableWidth * FinalWordWrapToleranceWidthScale);

        public static double ShapeAutoFitWrapTolerance(double fontSize, double availableWidth) => Math.Max(WrapFitTolerance(fontSize), availableWidth * ShapeAutoFitWrapToleranceWidthScale);

        public static double FallbackAdvanceWidth(int codeUnitCount, int runeCount, double fontSize, double characterSpacing)
        {
            return Math.Max(0d, codeUnitCount * fontSize * FallbackAdvanceWidthScale + Math.Max(0, runeCount - 1) * characterSpacing);
        }
    }

    private static class PptxPdfTextEmissionProfile
    {
        private const double OfficeExportFontGridDpi = 600d;
        private const double PointsPerInch = 72d;

        public static double FontSize(PptxPdfTextEmissionContext context) => FontSize(context.LayoutFontSize);

        public static double CharacterSpacing(PptxPdfTextEmissionContext context, double layoutCharacterSpacing) => layoutCharacterSpacing;

        public static double FontSize(double layoutFontSize)
        {
            double deviceUnits = Math.Round(
                layoutFontSize * OfficeExportFontGridDpi / PointsPerInch,
                MidpointRounding.AwayFromZero);

            return deviceUnits * PointsPerInch / OfficeExportFontGridDpi;
        }
    }

    private static class PptxChartMetricRules
    {
        public const double TitleFallbackFontSize = 12d;
        public const double LegendFallbackFontSize = 9d;
        public const double CategoryAxisFallbackFontSize = 9d;
        public const double ValueAxisFallbackFontSize = 8.5d;
        public const double DataLabelFallbackFontSize = 8.5d;
        public const double AxisSingleValueHeadroomFactor = 1.2d;
        public const double AxisValueEpsilon = 0.0001d;
        public const double AxisNiceTickTargetCount = 8d;
        public const double AxisNiceHorizontalValueTickTargetCount = 10d;
        public const double BubbleAxisBoundsTickTargetCount = 5d;
        public const double BubbleAxisNiceTickTargetCount = 10d;
        public const double AxisNiceTickStepSmall = 1d;
        public const double AxisNiceTickStepMedium = 2d;
        public const double AxisNiceTickStepLarge = 5d;
        public const double AxisNiceTickStepMaximum = 10d;
        public const double AxisNiceNearMaximumHeadroomRatio = 0.96d;
        public const double AreaChartStackedAxisNearMaximumHeadroomRatio = 0.95d;
        public const double BubbleRadiusPlotRatio = 0.128d;
        public const double DoughnutHoleMinimumRatio = 0.1d;
        public const double DoughnutHoleMaximumRatio = 0.9d;
        public const double DoughnutHoleFallbackRatio = 0.56d;
        public const double DefaultPlotBoxXRatio = 0.12d;
        public const double DefaultPlotBoxYRatio = 0.16d;
        public const double DefaultPlotBoxWidthRatio = 0.76d;
        public const double DefaultPlotBoxHeightRatio = 0.68d;
        public const double BarDefaultPlotBoxXRatio = 0.1d;
        public const double BarDefaultPlotBoxYRatio = 0.14d;
        public const double BarDefaultPlotBoxWidthRatio = 0.82d;
        public const double BarDefaultPlotBoxHeightRatio = 0.81d;
        public const double BarOverlayOnlyPlotBoxXRatio = 0.0576d;
        public const double BarOverlayOnlyPlotBoxYRatio = 0.0924d;
        public const double BarOverlayOnlyPlotBoxWidthRatio = 0.9272d;
        public const double BarOverlayOnlyPlotBoxHeightRatio = 0.8706d;
        public const double BarNoTitleBottomLegendPlotBoxXRatio = 0.0415d;
        public const double BarNoTitleBottomLegendPlotBoxYRatio = 0.151d;
        public const double BarNoTitleBottomLegendPlotBoxWidthRatio = 0.9406d;
        public const double BarNoTitleBottomLegendPlotBoxHeightRatio = 0.815d;
        public const double BarTitleNoLegendPlotBoxXRatio = 0.1106d;
        public const double BarTitleNoLegendPlotBoxYRatio = 0.1008d;
        public const double BarTitleNoLegendPlotBoxWidthRatio = 0.8691d;
        public const double BarTitleNoLegendPlotBoxHeightRatio = 0.7696d;
        public const double BarTitleNoLegendInsideCrossingPlotBoxXRatio = 0.0652d;
        public const double BarTitleNoLegendInsideCrossingPlotBoxYRatio = 0.0370d;
        public const double BarTitleNoLegendInsideCrossingPlotBoxWidthRatio = 0.9195d;
        public const double BarTitleNoLegendInsideCrossingPlotBoxHeightRatio = 0.8441d;
        public const double BarMultiValueAxisPrimaryStripFactor = 1.85d;
        public const double BarMultiValueAxisSecondaryStripFactor = 1.2d;
        public const double HorizontalBarTitleNoLegendPlotBoxXRatio = 0.1524d;
        public const double HorizontalBarTitleNoLegendPlotBoxYRatio = 0.0924d;
        public const double HorizontalBarTitleNoLegendPlotBoxWidthRatio = 0.8196d;
        public const double HorizontalBarTitleNoLegendPlotBoxHeightRatio = 0.8003d;
        public const double HorizontalBarManualLayoutTargetPlotBoxXRatio = 0.179d;
        public const double HorizontalBarManualLayoutTargetPlotBoxYRatio = 0.1005d;
        public const double LineNoTitleRightLegendPlotBoxXRatio = 0.0828d;
        public const double LineNoTitleRightLegendPlotBoxYRatio = 0.0908d;
        public const double LineNoTitleRightLegendPlotBoxWidthRatio = 0.7687d;
        public const double LineNoTitleRightLegendPlotBoxHeightRatio = 0.8722d;
        public const double LineNoTitleRightLegendExplicitScalePlotBoxYRatio = 0.10128571428571429d;
        public const double LineNoTitleRightLegendExplicitScalePlotBoxHeightRatio = 0.8463333333333334d;
        public const double LineRightLegendValueAxisPadding = 35.7d;
        public const double LineRightLegendValueAxisFrameWidthPaddingRatio = 0.05d;
        public const double LineRightLegendExtraValueLabelCharacterPadding = 5.3d;
        public const double LineRightLegendReservePadding = 43.8d;
        public const double LineRightLegendExtraLegendCharacterPadding = 8.4d;
        public const double AreaRightLegendReserveFrameWidthFactor = 0.025d;
        public const double LineTitleRightLegendPlotBoxXRatio = 0.0639d;
        public const double LineTitleRightLegendPlotBoxYRatio = 0.0924d;
        public const double LineTitleRightLegendPlotBoxWidthRatio = 0.7391d;
        public const double LineTitleRightLegendPlotBoxHeightRatio = 0.7888d;
        public const double BubbleTitleRightLegendPlotBoxWidthRatio = 0.7863333333333333d;
        public const double BubbleTitleRightLegendSwatchXRatio = 0.8818333333333334d;
        public const double BubbleTitleRightLegendSwatchYRatio = 0.4476157407407407d;
        public const double PieCenterXRatio = 0.42d;
        public const double PieNoLegendCenterXRatio = 0.5d;
        public const double DoughnutRightLegendCenterXRatio = 0.3988d;
        public const double DoughnutLeftLegendCenterXRatio = 0.56743d;
        public const double DoughnutHorizontalLegendCenterXRatio = 0.5d;
        public const double DoughnutTopLegendCenterYRatio = 0.46095d;
        public const double DoughnutBottomLegendCenterYRatio = 0.53907d;
        public const double DoughnutHorizontalLegendRadiusRatio = 0.4355d;
        public const double DoughnutNoLegendCenterYRatio = 0.5d;
        public const double DoughnutNoLegendRadiusRatio = 0.4746d;
        public const double DoughnutExplosionCenterOffsetRatio = 1.0d;
        public const double PieCenterYRatio = 0.458d;
        public const double PieRadiusRatio = 0.434d;
        public const double PieDataLabelRadiusRatio = 0.62d;
        public const double PieDataLabelWidthRatio = 0.55d;
        public const double PieDataLabelMinimumWidth = 18d;
        public const double PieDataLabelHeightFactor = 1.35d;
        public const double DataLabelLegendKeySizeFactor = 0.55d;
        public const double DataLabelLegendKeyTextGapFactor = 0.35d;
        public const double PieExplosionLabelRadiusRatio = 0.22d;
        public const double CartesianDataLabelHeightFactor = 1.35d;
        public const double CartesianDataLabelMinimumWidth = 18d;
        public const double BarDataLabelHorizontalGap = 2d;
        public const double BarDataLabelVerticalGap = 1d;
        public const double BarDataLabelSlotFillRatio = 0.82d;
        public const double BarDataLabelCategoryInsetRatio = 0.09d;
        public const double HorizontalBarDataLabelWidthRatio = 0.1d;
        public const double HorizontalBarDataLabelSlotCenterRatio = 0.43d;
        public const double VerticalBarDataLabelWidthRatio = 0.86d;
        public const double LineDataLabelMinimumPointSpan = 5d;
        public const double LineDataLabelPointWidthFactor = 1.5d;
        public const double LineDataLabelSideGap = 2d;
        public const double LineDataLabelBelowOffsetFactor = 1.25d;
        public const double LineDataLabelAboveOffsetFactor = 0.35d;
        public const double AxisLabelHeightFactor = 1.35d;
        public const int CategoryAxisDefaultLabelOffset = 100;
        public const int CategoryAxisMinimumLabelOffset = 0;
        public const int CategoryAxisMaximumLabelOffset = 1000;
        public const int CategoryAxisDefaultTickLabelSkip = 1;
        public const double CategoryAxisHorizontalLeftOffsetRatio = 0.1882d;
        public const double CategoryAxisHorizontalWidthRatio = 0.16d;
        public const double CategoryAxisHorizontalBaselineRatio = 0.217d;
        public const double CategoryAxisVerticalWidthFactor = 1.35d;
        public const double CategoryAxisVerticalTopOffsetFactor = 1.18d;
        public const double CategoryAxisVerticalTopSideOffsetFactor = 0.70d;
        public const double CategoryAxisMajorTickLength = 4d;
        public const double AxisLabelClipTopOffsetFactor = 0.25d;
        public const double AxisLabelClipHeightFactor = 1.6d;
        public const double ValueAxisMinimumLabelWidthFactor = 1.6d;
        public const double ValueAxisLabelPaddingFactor = 0.45d;
        public const double ValueAxisLabelSideGapFactor = 0.93d;
        public const double ValueAxisVerticalClipExtraHeightFactor = 2d;
        public const double HorizontalValueAxisSlotCount = 5d;
        public const double HorizontalValueAxisTopOffsetFactor = 1.18d;
        public const double VerticalValueAxisWidthRatio = 0.12d;
        public const double VerticalValueAxisBaselineRatio = 0.215d;
        public const double TitleXInsetRatio = 0.08d;
        public const double TitleBaselineYRatio = 0.88d;
        public const double PolarTitleBaselineYRatio = 0.935d;
        public const double AutoTitleFontScale = 1.2d;
        public const double TitleAbovePlotBaselineOffsetFactor = 0.8483333333333334d;
        public const double AutoBarTitleAbovePlotBaselineOffsetFactor = 1.08d;
        public const double TitleWidthRatio = 0.84d;
        public const double TitleHeightFactor = 1.4d;
        public const double DefaultAxisTitleBandBaselineRatio = 0.23d;
        public const double DefaultAxisTitleSideBaselineRatio = 0.28d;
        public const double DefaultAxisTitleTopBandBaselineRatio = 0.485d;
        public const double DefaultAxisTitleLeftSideBaselineRatio = 0.50d;
        public const double DefaultAxisTitleRightSideBaselineRatio = 0.67d;
        public const double DefaultAxisTitlePlotSideReserveRatio = 0.089d;
        public const double DefaultAxisTitlePlotOppositeSideReserveRatio = 0.020d;
        public const double DefaultAxisTitlePlotBandReserveRatio = 0.126d;
        public const double DefaultAxisTitlePlotOppositeBandReserveRatio = 0.030d;
        public const double DefaultAxisTitleHorizontalBarPlotSideReserveRatio = 0.114d;
        public const double DefaultAxisTitleHorizontalBarPlotOppositeSideReserveRatio = 0.028d;
        public const double DefaultAxisTitleHorizontalBarPlotBandReserveRatio = 0.126d;
        public const double LegendLineHeightFactor = 1.45d;
        public const double LegendSideStrokeLineHeightFactor = 1.5433333333333332d;
        public const double LegendMarkerSizeFactor = 0.55d;
        public const double LegendMinimumSideWidth = 36d;
        public const double LegendSideFillMinimumWidthFactor = 1.95d;
        public const double LegendSideWidthRatio = 0.22d;
        public const double LegendSideGap = 8d;
        public const double LegendSideStrokeGapFactor = 0.8333333333333334d;
        public const double LegendSideStrokeMarkerWidthFactor = 1.0666666666666667d;
        public const double LegendSideStrokeTextGapFactor = 0.11666666666666667d;
        public const double LegendSideStrokeBaselineCenterOffsetFactor = 0.955d;
        public const double LegendSideStrokeStyledMarkerBaselineCenterOffsetFactor = 0.561d;
        public const double LegendSideFillReservedBandOffsetFactor = 0.04d;
        public const double LegendSideFillContentBoxReservedBandOffsetFactor = 0.06d;
        public const double LegendSideFillBaselineCenterOffsetFactor = 0.276d;
        public const double LegendFullFrameSideInsetRatio = 0.02d;
        public const double LegendFullFrameBottomBaselineFactor = 0.57d;
        public const double LegendFullFrameTopBaselineFactor = 0.95d;
        public const double LegendBottomOffsetFactor = 2.39d;
        public const double LegendTopOffsetFactor = 0.15d;
        public const double LegendHorizontalClipHeightFactor = 1.25d;
        public const double LegendMarkerBaselineFactor = 0.35d;
        public const double LegendSideFillMarkerBaselineFactor = 0d;
        public const double LegendHorizontalMarkerBaselineFactor = 0d;
        public const double LegendTextGap = 3d;
    }

    private readonly record struct ChartPolarGeometry(double CenterX, double CenterY, double Radius);

    private enum ChartPolarKind
    {
        Pie,
        Doughnut
    }

    private readonly record struct ChartPolarLayout(
        ChartPolarKind Kind,
        ChartPlotBox PlotBox,
        ChartPolarGeometry Geometry,
        double ExplosionReserve,
        bool HasLegend);

    private enum ChartRadarStyle
    {
        Marker,
        Filled
    }

    private readonly record struct ChartRadarGeometryRule(double CenterXRatio, double CenterYRatio, double RadiusRatio);

    private readonly record struct ChartRadarLabelRules(
        double CategoryVerticalGapFactor,
        double CategoryHorizontalGapFactor,
        double CategoryBaselineBaseFactor,
        double CategoryBaselineSineFactor,
        double CategoryBaselineSineSquaredFactor,
        double ValueGapFactor,
        double ValueBaselineOffsetFactor,
        double ValueWidthFactor);

    private readonly record struct ChartRadarLayout(
        ChartPlotBox PlotBox,
        ChartPolarGeometry Geometry,
        ChartRadarStyle Style,
        int PointCount,
        ChartRadarLabelRules LabelRules)
    {
        public bool IsFilled => Style == ChartRadarStyle.Filled;
    }

    private readonly record struct ChartRadarLabelFrame(
        double X,
        double Y,
        double Width,
        double Height,
        TextAlignment Alignment);

    private enum ChartPlotBoxPreset
    {
        DefaultCartesian,
        BarDefault,
        BarOverlayOnly,
        BarNoTitleBottomLegend,
        BarTitleNoLegend,
        BarTitleNoLegendInsideCrossing,
        HorizontalBarTitleNoLegend,
        LineNoTitleRightLegend,
        LineTitleRightLegend
    }

    private readonly record struct ChartPlotBoxRatios(double Left, double Top, double Width, double Height)
    {
        public double Right => Left + Width;

        public double Bottom => Top + Height;
    }

    private enum TextVerticalAnchor
    {
        Unknown,
        Top,
        Middle,
        Bottom
    }

    private enum PptxTextWrapMode
    {
        Unknown,
        Square,
        None
    }

    private enum PptxTextVerticalOverflow
    {
        Unknown,
        Overflow,
        Clip,
        Ellipsis
    }

    private enum PptxTextOrientation
    {
        Unknown,
        Horizontal,
        Vertical,
        Vertical270,
        EastAsianVertical,
        MongolianVertical,
        WordArtVertical,
        WordArtVerticalRightToLeft
    }

    private enum LineEndKind
    {
        None,
        Triangle,
        Arrow,
        Stealth,
        Diamond,
        Oval
    }

    private readonly record struct LineEndStyle(LineEndKind Kind, double WidthScale, double LengthScale)
    {
        public bool IsNone => Kind == LineEndKind.None;
    }

    private readonly record struct LineStyle(bool HasLine, RgbColor Color, double Width, double Alpha, IReadOnlyList<double> DashPattern, int? Cap, int? Join)
    {
        public bool HasDash => DashPattern is { Count: > 0 };
    }

    private readonly record struct FillStyle(bool HasFill, RgbColor Color, double Alpha);

    private sealed record GradientFill(double AngleDegrees, IReadOnlyList<GradientStop> Stops);

    private readonly record struct GradientStop(double Offset, RgbColor Color, double Alpha);

    private readonly record struct CropRect(double Left, double Top, double Right, double Bottom)
    {
        public bool IsEmpty => Left == 0d && Top == 0d && Right == 0d && Bottom == 0d;
    }

    private readonly record struct FillRect(double Left, double Top, double Right, double Bottom);

    private readonly record struct ShapePatternFill(string Preset, RgbColor Foreground, RgbColor Background, double Alpha);

    private readonly record struct ShapePictureFill(
        string RelationshipId,
        string? TargetPartName,
        PptxSceneImageResource? Resource,
        CropRect Crop,
        FillRect Fill,
        double Alpha);

    private readonly record struct Glow(RgbColor Color, double Alpha, double Radius);

    private readonly record struct OuterShadow(RgbColor Color, double Alpha, double OffsetX, double OffsetY, double BlurRadius);

    private readonly record struct SvgPaint(RgbColor? Color, SvgGradient? Gradient);

    private sealed record SvgGradient(double X1, double Y1, double X2, double Y2, IReadOnlyList<SvgGradientStop> Stops);

    private readonly record struct SvgGradientStop(double Offset, RgbColor Color);

    private readonly record struct SvgPathBounds(double MinX, double MinY, double MaxX, double MaxY)
    {
        public double CenterX => (MinX + MaxX) / 2d;

        public double CenterY => (MinY + MaxY) / 2d;
    }
}
