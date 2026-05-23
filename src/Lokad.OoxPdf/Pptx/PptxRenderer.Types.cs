using System.Globalization;
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
        TextOutline? Outline = null);

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
        bool SyntheticBold,
        bool SyntheticItalic);

    private sealed record PptxPositionedTextSpan(
        PptxTextRunModel? SourceRun,
        PptxTextLineBoxLayout? LineBox,
        TextRun Run,
        double EndX,
        IReadOnlyList<PptxTextAtomLayout> Atoms,
        PptxTextGlyphSpanLayout GlyphSpan);

    private sealed record TextGlyphAtom(
        int CodePoint,
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
        XElement? Properties,
        XElement? DefaultRunProperties,
        double FontSize,
        double SpacingBefore,
        double SpacingAfter,
        LineSpacing LineSpacing,
        ParagraphIndent Indent,
        IReadOnlyList<double> TabStops);

    private readonly record struct ResolvedRunTextStyle(
        double NominalFontSize,
        double FontSize,
        double CharacterSpacing,
        double BaselineOffset,
        RgbColor Color,
        double Alpha,
        TextOutline? Outline,
        RgbColor? Highlight,
        bool Bold,
        bool Italic,
        bool Underline,
        bool Strike,
        bool KerningEnabled,
        string? Typeface);

    private sealed record PptxTextFrameModel(
        XElement Shape,
        XElement TextBody,
        XElement? InheritedTextBody,
        PptxTheme Theme,
        ShapeBounds Bounds,
        TextInsets Insets,
        double FontScale,
        double LineSpacingScale,
        double TextX,
        double TextWidth,
        double TextWrapWidth,
        double TextHeight,
        double TextClipY,
        double TextClipHeight,
        int ColumnCount,
        double ColumnSpacing,
        double RotationCenterX,
        double RotationCenterY,
        double TextRotationDegrees,
        bool TextFlipHorizontal,
        bool TextFlipVertical,
        double FlowYTop,
        double VerticalOffset,
        PptxTextOrientation Orientation,
        RgbColor? ShapeFontColor,
        IReadOnlyList<PptxTextParagraphModel> Paragraphs);

    private sealed record PptxTextParagraphModel(
        XElement Source,
        XElement? Properties,
        XElement? DefaultProperties,
        int Level,
        PptxParagraphStyleCascade Cascade,
        ResolvedParagraphTextStyle Style,
        IReadOnlyList<PptxTextRunModel> Runs);

    private sealed record PptxParagraphStyleCascade(
        string LevelName,
        IReadOnlyList<PptxParagraphStyleLayer> Layers)
    {
        public IReadOnlyList<XElement?> Sources => Layers.Select(layer => layer.Source).ToArray();
    }

    private sealed record PptxParagraphStyleLayer(
        string Name,
        XElement? Source);

    private sealed record PptxTextRunModel(
        PptxTextRunKind Kind,
        XElement Source,
        XElement? Properties,
        string Text,
        ResolvedRunTextStyle Style);

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
                run.Width,
                run.Width,
                []);
        }
    }

    private sealed record PptxTextGlyphLayout(
        int CodePoint,
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

        public void Reset(double startX)
        {
            Spans.Clear();
            EndX = startX;
        }
    }

    private readonly record struct TextInsets(double Left, double Right, double Top, double Bottom);

    private readonly record struct ParagraphIndent(double MarginLeft, double Hanging);

    private readonly record struct RenderedFont(string ResourceName, PdfEmbeddedFont Font, bool SyntheticBold, bool SyntheticItalic);

    private readonly record struct RenderedFonts(IReadOnlyDictionary<string, RenderedFont> Fonts, IReadOnlyList<PdfFontResource> Resources);

    private readonly record struct BulletStyle(double FontSize, RgbColor Color, string? Typeface);

    private readonly record struct LineSpacing(double Value, bool IsAbsolute, bool IsExplicit)
    {
        public static LineSpacing Absolute(double points) => new(points, true, true);

        public static LineSpacing Multiple(double factor, bool isExplicit) => new(factor, false, isExplicit);

        public LineSpacing ScaleExplicit(double factor)
        {
            return IsExplicit
                ? new LineSpacing(Value * factor, IsAbsolute, IsExplicit)
                : this;
        }

        public double Resolve(double fontSize)
        {
            return IsAbsolute ? Value : fontSize * Value;
        }
    }

    private sealed class TextAdvanceEstimator
    {
        private readonly WindowsFontResolver resolver = new();
        private readonly Dictionary<string, OpenTypeFont?> fonts = new(StringComparer.OrdinalIgnoreCase);

        public double Measure(string text, double fontSize, string? familyName, bool bold = false, bool italic = false, double characterSpacing = 0d, bool kerningEnabled = true)
        {
            OpenTypeFont? font = ResolveFont(string.IsNullOrWhiteSpace(familyName) ? "Arial" : familyName, bold, italic);
            if (font is null)
            {
                int fallbackRuneCount = text.EnumerateRunes().Count();
                return PptxTextMetricRules.FallbackAdvanceWidth(text.Length, fallbackRuneCount, fontSize, characterSpacing);
            }

            double units = 0d;
            ushort previousGlyph = 0;
            foreach (Rune rune in text.EnumerateRunes())
            {
                ushort glyph = font.MapCodePoint(rune.Value);
                if (kerningEnabled && previousGlyph != 0 && glyph != 0)
                {
                    units += font.GetKerning(previousGlyph, glyph);
                }

                units += font.GetAdvanceWidth(glyph);
                previousGlyph = glyph;
            }

            int runeCount = text.EnumerateRunes().Count();
            return Math.Max(0d, units * fontSize / font.UnitsPerEm + Math.Max(0, runeCount - 1) * characterSpacing);
        }

        public OpenTypeFont? ResolveOpenTypeFont(string? familyName, bool bold = false, bool italic = false)
        {
            return ResolveFont(string.IsNullOrWhiteSpace(familyName) ? "Arial" : familyName, bold, italic);
        }

        private OpenTypeFont? ResolveFont(string familyName, bool bold, bool italic)
        {
            string key = familyName + "\u001f" + bold.ToString(CultureInfo.InvariantCulture) + "\u001f" + italic.ToString(CultureInfo.InvariantCulture);
            if (fonts.TryGetValue(key, out OpenTypeFont? cached))
            {
                return cached;
            }

            try
            {
                FontResolution resolution = resolver.ResolvePresentationTextFace(new FontRequest(familyName, bold, italic));
                cached = resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath)
                    ? null
                    : OpenTypeFont.Load(resolution.FontFilePath, resolution.FontFaceIndex);
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException)
            {
                cached = null;
            }

            fonts[key] = cached;
            return cached;
        }
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
        public const double CssSuperscriptSubscriptScale = 0.65d;
        public const double CssNormalLineHeightFallback = 1.2d;
        public const double OfficeCompatibleDefaultLineSpacingFactor = 1.1d;
        public const double OfficeManualBreakDefaultLineHeightFallback = 1.24d;
        public const double OfficeManualBreakBaselineFallback = 0.9344d;
        public const double OfficeBaselineFallback = 0.974d;
        public const double MinimumBaselineMetricRatio = 0.75d;
        public const double MaximumBaselineMetricRatio = 1.05d;
        public const double LargeTextBaselineMinimumFontSize = 24d;
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
        public const double StrikePositionFallback = 0.211d;
        public const double StrikeThicknessFallback = 0.05d;
        public const double HighlightDescenderPaddingFontUnits = 32d;
        public const double HighlightMaximumDescentFontScale = 0.23d;
        public const double HighlightMaximumHeightFontScale = 1.18d;
        public const double AdjacentTextCoalesceGapFontScale = 0.2d;
        public const double AdjacentUnderlineCoalesceGapFontScale = 0.08d;
        public const double WrapFitToleranceFontScale = 0.16d;
        public const double FinalWordWrapToleranceWidthScale = 0.02d;
        public const double FallbackAdvanceWidthScale = 0.42d;
        public const int ShapeAutoFitSearchIterations = 10;

        public static double ClampNonNegative(double value) => Math.Max(0d, value);

        public static double MinimumWidth(double width) => Math.Max(MinimumDrawableDimension, width);

        public static double SuperscriptSubscriptFontSize(double nominalFontSize) => nominalFontSize * CssSuperscriptSubscriptScale;

        public static bool ShouldScaleSuperscriptSubscript(double baselineOffset, double nominalFontSize) =>
            Math.Abs(baselineOffset) >= nominalFontSize * SuperscriptSubscriptMinimumBaselineRatio;

        public static double SmallCapsFontScale() => SmallCapsFallbackScale;

        public static double TextOutlineWidth(double? width) => Math.Max(MinimumStrokeWidth, width ?? DefaultTextOutlineWidth);

        public static double SyntheticBoldStrokeWidth(double fontSize) => Math.Max(0d, fontSize * SyntheticBoldStrokeWidthRatio);

        public static double StrikeY(double baselineY, double fontSize) => baselineY + fontSize * StrikePositionFallback;

        public static double StrikeThickness(double fontSize) => Math.Max(MinimumStrokeWidth, fontSize * StrikeThicknessFallback);

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

        public static double FinalWordWrapTolerance(double fontSize, double availableWidth) => Math.Max(WrapFitTolerance(fontSize), availableWidth * FinalWordWrapToleranceWidthScale);

        public static double FallbackAdvanceWidth(int codeUnitCount, int runeCount, double fontSize, double characterSpacing)
        {
            return Math.Max(0d, codeUnitCount * fontSize * FallbackAdvanceWidthScale + Math.Max(0, runeCount - 1) * characterSpacing);
        }
    }

    private enum TextVerticalAnchor
    {
        Top,
        Middle,
        Bottom
    }

    private enum PptxTextOrientation
    {
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

    private readonly record struct CropRect(double Left, double Top, double Right, double Bottom)
    {
        public bool IsEmpty => Left == 0d && Top == 0d && Right == 0d && Bottom == 0d;
    }

    private readonly record struct FillRect(double Left, double Top, double Right, double Bottom);

    private readonly record struct ShapePatternFill(string Preset, RgbColor Foreground, RgbColor Background, double Alpha);

    private readonly record struct Glow(RgbColor Color, double Alpha, double Radius);

    private readonly record struct OuterShadow(RgbColor Color, double Alpha, double OffsetX, double OffsetY);

    private readonly record struct SvgPaint(RgbColor? Color, SvgGradient? Gradient);

    private sealed record SvgGradient(double X1, double Y1, double X2, double Y2, IReadOnlyList<SvgGradientStop> Stops);

    private readonly record struct SvgGradientStop(double Offset, RgbColor Color);

    private readonly record struct SvgPathBounds(double MinX, double MinY, double MaxX, double MaxY)
    {
        public double CenterX => (MinX + MaxX) / 2d;

        public double CenterY => (MinY + MaxY) / 2d;
    }
}
