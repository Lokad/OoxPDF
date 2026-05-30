using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.PptxInspect <input.pptx> <output-directory> [--slide <number>]... [--include-text]");
    return 2;
}

string inputPath = Path.GetFullPath(args[0]);
string outputDirectory = Path.GetFullPath(args[1]);
bool includeText = args.Any(arg => string.Equals(arg, "--include-text", StringComparison.Ordinal));
HashSet<int>? slideFilter = ReadSlideFilter(args);

Directory.CreateDirectory(outputDirectory);

using FileStream stream = File.OpenRead(inputPath);
OoxPackage package = OoxPackage.Open(stream);
PptxDocument document = new PptxReader().Read(package);
IEnumerable<PptxSlide> slides = document.Slides;
if (slideFilter is not null)
{
    slides = slides.Where(slide => slideFilter.Contains(slide.Index + 1));
}

var records = new List<PptxGlyphRunRecord>();
var frameRecords = new List<PptxTextFrameRecord>();
var tableFrameRecords = new List<PptxTextFrameRecord>();
var paragraphRecords = new List<PptxTextParagraphRecord>();
var tableParagraphRecords = new List<PptxTextParagraphRecord>();
var lineRecords = new List<PptxTextLineRecord>();
foreach (PptxSlide slide in slides)
{
    int slideNumber = slide.Index + 1;
    IReadOnlyList<PptxTextFrameModelSnapshot> frameModels = PptxRenderer.InspectTextFrameModels(document, package, slide.Index);
    for (int frameIndex = 0; frameIndex < frameModels.Count; frameIndex++)
    {
        PptxTextFrameModelSnapshot frame = frameModels[frameIndex];
        frameRecords.Add(ToFrameRecord(slideNumber, frameIndex, frame));
        for (int paragraphIndex = 0; paragraphIndex < frame.Paragraphs.Count; paragraphIndex++)
        {
            PptxTextParagraphModelSnapshot paragraph = frame.Paragraphs[paragraphIndex];
            paragraphRecords.Add(ToParagraphRecord(slideNumber, frameIndex, paragraphIndex, paragraph, includeText));
        }
    }

    IReadOnlyList<PptxTextFrameModelSnapshot> tableFrameModels = PptxRenderer.InspectTableTextFrameModels(document, package, slide.Index);
    for (int frameIndex = 0; frameIndex < tableFrameModels.Count; frameIndex++)
    {
        PptxTextFrameModelSnapshot frame = tableFrameModels[frameIndex];
        tableFrameRecords.Add(ToFrameRecord(slideNumber, frameIndex, frame));
        for (int paragraphIndex = 0; paragraphIndex < frame.Paragraphs.Count; paragraphIndex++)
        {
            tableParagraphRecords.Add(ToParagraphRecord(slideNumber, frameIndex, paragraphIndex, frame.Paragraphs[paragraphIndex], includeText));
        }
    }

    PptxTextLayoutSnapshot textLayout = PptxRenderer.InspectTextLayout(document, package, slide.Index);
    for (int frameIndex = 0; frameIndex < textLayout.Frames.Count; frameIndex++)
    {
        PptxTextFrameLayoutSnapshot frame = textLayout.Frames[frameIndex];
        for (int paragraphIndex = 0; paragraphIndex < frame.Paragraphs.Count; paragraphIndex++)
        {
            PptxTextParagraphLayoutSnapshot paragraph = frame.Paragraphs[paragraphIndex];
            for (int lineIndex = 0; lineIndex < paragraph.Lines.Count; lineIndex++)
            {
                PptxTextLineLayoutSnapshot line = paragraph.Lines[lineIndex];
                lineRecords.Add(new PptxTextLineRecord(
                    slideNumber,
                    frameIndex,
                    paragraphIndex,
                    lineIndex,
                    paragraph.Level,
                    Round(line.TopY),
                    Round(line.BaselineY),
                    Round(line.Advance),
                    Round(line.BaselineOffset),
                    Round(line.MaxFontSize),
                    line.LineSpacingKind,
                    line.BaselineMetric.Source,
                    line.BaselineMetric.Typeface,
                    line.BaselineMetric.Bold,
                    line.BaselineMetric.Italic,
                    Round(line.BaselineMetric.FontSize),
                    Round(line.BaselineMetric.Ratio),
                    line.BaselineMetric.UnitsPerEm,
                    line.BaselineMetric.WindowsAscender,
                    line.BaselineMetric.WindowsDescender,
                    line.BaselineMetric.TypographicAscender,
                    line.BaselineMetric.TypographicDescender,
                    line.BaselineMetric.TypographicLineGap,
                    Round(line.StartX),
                    Round(line.EndX),
                    Round(line.NaturalEndX),
                    line.Alignment,
                    line.Spans.Count,
                    line.Spans.Sum(span => span.Text.Length),
                    line.Spans.Sum(span => span.GlyphSpan.GlyphCount),
                    line.Spans.Select((span, spanIndex) => new PptxTextSpanRecord(
                        spanIndex,
                        includeText ? span.Text : null,
                        span.Text.Length,
                        Round(span.X),
                        Round(span.Y),
                        Round(span.Width),
                        Round(span.FontSize),
                        span.GlyphSpan.Typeface,
                        Round(span.GlyphSpan.LeadingAdjustment),
                        Round(span.GlyphSpan.NaturalWidth),
                        Round(span.GlyphSpan.LayoutWidth),
                        span.GlyphSpan.GlyphCount,
                        Round(span.GlyphSpan.FirstAdjustmentAfterOrigin))).ToArray()));
            }
        }
    }

    foreach (PptxTextGlyphRunSnapshot run in PptxRenderer.InspectTextGlyphRuns(document, package, slide.Index))
    {
        GlyphCategoryCounts categories = CountGlyphCategories(run.Text);
        records.Add(new PptxGlyphRunRecord(
            slideNumber,
            includeText ? run.Text : null,
            run.Text.Length,
            Round(run.X),
            Round(run.BaselineY),
            Round(run.Width),
            Round(run.NaturalWidth),
            Round(run.LayoutWidth),
            run.FontFamily,
            run.Bold,
            run.Italic,
            run.Underline,
            run.Strike,
            run.SyntheticBold,
            run.SyntheticItalic,
            FormatColor(run.HighlightColor),
            RoundNullable(run.HighlightX),
            RoundNullable(run.HighlightY),
            RoundNullable(run.HighlightWidth),
            RoundNullable(run.HighlightHeight),
            RoundNullable(run.UnderlineX),
            RoundNullable(run.UnderlineY),
            RoundNullable(run.UnderlineWidth),
            RoundNullable(run.UnderlineHeight),
            RoundNullable(run.StrikeX),
            RoundNullable(run.StrikeY),
            RoundNullable(run.StrikeWidth),
            RoundNullable(run.StrikeHeight),
            run.FrameIndex,
            run.ParagraphIndex,
            run.SourceRunIndex,
            run.ParagraphBulletKind,
            run.ParagraphAutoNumberType,
            run.ParagraphAutoNumberStartAt,
            run.LineIndex,
            run.SpanIndex,
            run.LineSpanCount,
            Round(run.FrameFontScale),
            Round(run.FrameShapeX),
            Round(run.FrameShapeTopY),
            Round(run.FrameShapeWidth),
            Round(run.FrameShapeHeight),
            run.TableRowIndex,
            run.TableColumnIndex,
            run.TableRowSpan,
            run.TableColumnSpan,
            Round(run.FrameInsetLeft),
            Round(run.FrameInsetRight),
            Round(run.FrameInsetTop),
            Round(run.FrameInsetBottom),
            run.FrameWrapMode,
            run.FrameWrapValue,
            run.FrameVerticalOverflowMode,
            run.FrameVerticalOverflowValue,
            run.FrameVerticalOverflowSource,
            run.FrameAutofitMode,
            Round(run.FrameTextX),
            Round(run.FrameTextWidth),
            Round(run.FrameTextWrapWidth),
            Round(run.FrameTextHeight),
            Round(run.FrameClipX),
            Round(run.FrameClipWidth),
            Round(run.FrameClipY),
            Round(run.FrameClipHeight),
            run.FrameColumnCount,
            Round(run.FrameColumnSpacing),
            Round(run.LineTopY),
            Round(document.SlideHeightPoints - run.LineTopY - run.FrameShapeTopY),
            Round(document.SlideHeightPoints - run.LineTopY - run.FrameShapeTopY - run.FrameInsetTop),
            Round(document.SlideHeightPoints - run.BaselineY - run.FrameShapeTopY),
            Round(document.SlideHeightPoints - (run.LineTopY - run.LineAdvance) - (run.FrameShapeTopY + run.FrameShapeHeight)),
            Round(run.LineAdvance),
            Round(run.LineMaxFontSize),
            Round(run.LayoutFontSize),
            Round(run.PdfFontSize),
            Round(run.LayoutCharacterSpacing),
            Round(run.PdfCharacterSpacing),
            run.GlyphCount,
            Round(run.FirstAdjustmentAfterOrigin),
            categories.LetterCount,
            categories.DecimalDigitCount,
            categories.PunctuationCount,
            categories.SymbolCount,
            categories.SpaceCount,
            categories.OtherCount));
    }
}

var options = new JsonSerializerOptions
{
    WriteIndented = true,
    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};
string outputPath = Path.Combine(outputDirectory, "glyph-runs.json");
File.WriteAllText(outputPath, JsonSerializer.Serialize(records, options), Encoding.UTF8);
string frameOutputPath = Path.Combine(outputDirectory, "text-frame-models.json");
File.WriteAllText(frameOutputPath, JsonSerializer.Serialize(frameRecords, options), Encoding.UTF8);
string tableFrameOutputPath = Path.Combine(outputDirectory, "table-text-frame-models.json");
File.WriteAllText(tableFrameOutputPath, JsonSerializer.Serialize(tableFrameRecords, options), Encoding.UTF8);
string paragraphOutputPath = Path.Combine(outputDirectory, "text-paragraph-models.json");
File.WriteAllText(paragraphOutputPath, JsonSerializer.Serialize(paragraphRecords, options), Encoding.UTF8);
string tableParagraphOutputPath = Path.Combine(outputDirectory, "table-text-paragraph-models.json");
File.WriteAllText(tableParagraphOutputPath, JsonSerializer.Serialize(tableParagraphRecords, options), Encoding.UTF8);
string lineOutputPath = Path.Combine(outputDirectory, "text-layout-lines.json");
File.WriteAllText(lineOutputPath, JsonSerializer.Serialize(lineRecords, options), Encoding.UTF8);

Console.WriteLine(FormattableString.Invariant($"PPTX: {inputPath}"));
Console.WriteLine(FormattableString.Invariant($"Slides: {document.Slides.Count}"));
Console.WriteLine(FormattableString.Invariant($"Glyph runs: {records.Count}"));
Console.WriteLine(FormattableString.Invariant($"Output: {outputPath}"));
Console.WriteLine(FormattableString.Invariant($"Frame output: {frameOutputPath}"));
Console.WriteLine(FormattableString.Invariant($"Table frame output: {tableFrameOutputPath}"));
Console.WriteLine(FormattableString.Invariant($"Paragraph output: {paragraphOutputPath}"));
Console.WriteLine(FormattableString.Invariant($"Table paragraph output: {tableParagraphOutputPath}"));
Console.WriteLine(FormattableString.Invariant($"Layout output: {lineOutputPath}"));

return 0;

static PptxTextFrameRecord ToFrameRecord(int slideNumber, int frameIndex, PptxTextFrameModelSnapshot frame)
{
    return new PptxTextFrameRecord(
        slideNumber,
        frameIndex,
        Round(frame.FrameX),
        Round(frame.FrameY),
        Round(frame.FrameWidth),
        Round(frame.FrameHeight),
        Round(frame.TextX),
        Round(frame.TextWidth),
        Round(frame.TextWrapWidth),
        Round(frame.TextHeight),
        frame.TableRowIndex,
        frame.TableColumnIndex,
        frame.TableRowSpan,
        frame.TableColumnSpan,
        Round(frame.VerticalOffset),
        Round(frame.InsetLeft),
        Round(frame.InsetRight),
        Round(frame.InsetTop),
        Round(frame.InsetBottom),
        frame.InsetLeftValue,
        frame.InsetRightValue,
        frame.InsetTopValue,
        frame.InsetBottomValue,
        frame.InsetLeftSource,
        frame.InsetRightSource,
        frame.InsetTopSource,
        frame.InsetBottomSource,
        Round(frame.FontScale),
        frame.FontScaleValue,
        frame.FontScaleSource,
        Round(frame.LineSpacingScale),
        frame.LineSpacingReductionValue,
        frame.LineSpacingScaleSource,
        frame.CompatibleLineSpacing,
        frame.CompatibleLineSpacingSource,
        RoundNullable(frame.RotationDegrees),
        frame.RotationValue,
        frame.RotationDegreesSource,
        frame.InheritedPlaceholderCount,
        frame.HasInheritedTextBody,
        frame.UsesInheritedShapeBounds,
        frame.Orientation,
        frame.OrientationValue,
        frame.OrientationSource,
        frame.VerticalAnchor,
        frame.VerticalAnchorValue,
        frame.VerticalAnchorSource,
        frame.AnchorCenter,
        frame.AnchorCenterValue,
        frame.AnchorCenterSource,
        frame.WrapMode,
        frame.WrapValue,
        frame.WrapSource,
        frame.VerticalOverflow,
        frame.VerticalOverflowValue,
        frame.VerticalOverflowSource,
        frame.ColumnCount,
        Round(frame.ColumnSpacing),
        frame.ColumnSource,
        frame.ColumnCountSource,
        frame.ColumnSpacingSource,
        frame.ColumnCountValue,
        frame.ColumnSpacingValue,
        frame.AutofitModeValue,
        frame.AutofitModeSource,
        frame.Paragraphs.Count,
        frame.Paragraphs.Sum(paragraph => paragraph.Runs.Count),
        frame.Paragraphs.Sum(paragraph => paragraph.Runs.Sum(run => run.Text.Length)));
}

static PptxTextParagraphRecord ToParagraphRecord(int slideNumber, int frameIndex, int paragraphIndex, PptxTextParagraphModelSnapshot paragraph, bool includeText)
{
    return new PptxTextParagraphRecord(
        slideNumber,
        frameIndex,
        paragraphIndex,
        paragraph.Level,
        paragraph.CascadeLevelName,
        paragraph.ResolvedCascadeSourceCount,
        paragraph.CascadeLayerNames,
        paragraph.CascadeLayerKinds,
        paragraph.ResolvedStyleSourceCount,
        paragraph.ResolvedStyleLayerNames,
        paragraph.ResolvedStyleLayerKinds,
        paragraph.Alignment,
        paragraph.AlignmentValue,
        Round(paragraph.FontSize),
        paragraph.BulletKind,
        paragraph.BulletCharacter,
        paragraph.BulletResolvedCharacter,
        paragraph.BulletAutoNumberType,
        paragraph.BulletAutoNumberStartAtValue,
        paragraph.BulletFontTypeface,
        paragraph.BulletFontCharset,
        paragraph.BulletResolvedFontTypeface,
        paragraph.BulletFontTypefaceSource,
        paragraph.BulletColor,
        paragraph.BulletSizeKind,
        paragraph.BulletSizeValue,
        Round(paragraph.SpacingBefore),
        Round(paragraph.SpacingAfter),
        Round(paragraph.LineSpacingValue),
        paragraph.LineSpacingKind,
        paragraph.LineSpacingUseNormalLineAdvance,
        Round(paragraph.MarginLeft),
        Round(paragraph.HangingIndent),
        paragraph.Runs.Count,
        paragraph.Runs.Sum(run => run.Text.Length),
        paragraph.Runs.Select(run => new PptxTextRunRecord(
            run.RunIndex,
            run.Kind,
            includeText ? run.Text : null,
            run.Text.Length,
            run.ResolvedCascadeSourceCount,
            run.CascadeLayerNames,
            run.CascadeLayerKinds,
            Round(run.FontSize),
            Round(run.CharacterSpacing),
            run.Typeface,
            run.ColorSource,
            run.HasHyperlinkClick,
            run.HyperlinkClickId,
            run.Underline,
            run.UnderlineValue,
            run.Strike,
            run.StrikeValue,
            run.CapsValue,
            FormatColor(run.Highlight))).ToArray());
}

static HashSet<int>? ReadSlideFilter(string[] args)
{
    var slides = new HashSet<int>();
    for (int i = 2; i < args.Length; i++)
    {
        if (!string.Equals(args[i], "--slide", StringComparison.Ordinal))
        {
            continue;
        }

        if (i + 1 >= args.Length ||
            !int.TryParse(args[i + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int slide) ||
            slide <= 0)
        {
            throw new ArgumentException("--slide expects a positive one-based slide number.");
        }

        slides.Add(slide);
        i++;
    }

    return slides.Count == 0 ? null : slides;
}

static double Round(double value)
{
    return Math.Round(value, 6, MidpointRounding.AwayFromZero);
}

static double? RoundNullable(double? value)
{
    return value is null ? null : Round(value.Value);
}

static GlyphCategoryCounts CountGlyphCategories(string text)
{
    int letterCount = 0;
    int decimalDigitCount = 0;
    int punctuationCount = 0;
    int symbolCount = 0;
    int spaceCount = 0;
    int otherCount = 0;
    foreach (Rune rune in text.EnumerateRunes())
    {
        switch (Rune.GetUnicodeCategory(rune))
        {
            case UnicodeCategory.UppercaseLetter:
            case UnicodeCategory.LowercaseLetter:
            case UnicodeCategory.TitlecaseLetter:
            case UnicodeCategory.ModifierLetter:
            case UnicodeCategory.OtherLetter:
                letterCount++;
                break;
            case UnicodeCategory.DecimalDigitNumber:
                decimalDigitCount++;
                break;
            case UnicodeCategory.ConnectorPunctuation:
            case UnicodeCategory.DashPunctuation:
            case UnicodeCategory.OpenPunctuation:
            case UnicodeCategory.ClosePunctuation:
            case UnicodeCategory.InitialQuotePunctuation:
            case UnicodeCategory.FinalQuotePunctuation:
            case UnicodeCategory.OtherPunctuation:
                punctuationCount++;
                break;
            case UnicodeCategory.MathSymbol:
            case UnicodeCategory.CurrencySymbol:
            case UnicodeCategory.ModifierSymbol:
            case UnicodeCategory.OtherSymbol:
                symbolCount++;
                break;
            case UnicodeCategory.SpaceSeparator:
            case UnicodeCategory.LineSeparator:
            case UnicodeCategory.ParagraphSeparator:
                spaceCount++;
                break;
            default:
                otherCount++;
                break;
        }
    }

    return new GlyphCategoryCounts(letterCount, decimalDigitCount, punctuationCount, symbolCount, spaceCount, otherCount);
}

static string? FormatColor(RgbColor? color)
{
    return color is null
        ? null
        : string.Create(
            CultureInfo.InvariantCulture,
            $"{color.Value.Red:X2}{color.Value.Green:X2}{color.Value.Blue:X2}");
}

internal sealed record PptxGlyphRunRecord(
    int Slide,
    string? Text,
    int TextLength,
    double X,
    double BaselineY,
    double Width,
    double NaturalWidth,
    double LayoutWidth,
    string? FontFamily,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Strike,
    bool SyntheticBold,
    bool SyntheticItalic,
    string? HighlightColor,
    double? HighlightX,
    double? HighlightY,
    double? HighlightWidth,
    double? HighlightHeight,
    double? UnderlineX,
    double? UnderlineY,
    double? UnderlineWidth,
    double? UnderlineHeight,
    double? StrikeX,
    double? StrikeY,
    double? StrikeWidth,
    double? StrikeHeight,
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
    double LineTopY,
    double LineTopFromShapeTop,
    double LineTopFromTextTop,
    double BaselineFromShapeTop,
    double LineBottomFromShapeBottom,
    double LineAdvance,
    double LineMaxFontSize,
    double LayoutFontSize,
    double PdfFontSize,
    double LayoutCharacterSpacing,
    double PdfCharacterSpacing,
    int GlyphCount,
    double FirstAdjustmentAfterOrigin,
    int LetterCount,
    int DecimalDigitCount,
    int PunctuationCount,
    int SymbolCount,
    int SpaceCount,
    int OtherCount);

internal readonly record struct GlyphCategoryCounts(
    int LetterCount,
    int DecimalDigitCount,
    int PunctuationCount,
    int SymbolCount,
    int SpaceCount,
    int OtherCount);

internal sealed record PptxTextFrameRecord(
    int Slide,
    int FrameIndex,
    double FrameX,
    double FrameY,
    double FrameWidth,
    double FrameHeight,
    double TextX,
    double TextWidth,
    double TextWrapWidth,
    double TextHeight,
    int? TableRowIndex,
    int? TableColumnIndex,
    int? TableRowSpan,
    int? TableColumnSpan,
    double VerticalOffset,
    double InsetLeft,
    double InsetRight,
    double InsetTop,
    double InsetBottom,
    string? InsetLeftValue,
    string? InsetRightValue,
    string? InsetTopValue,
    string? InsetBottomValue,
    string InsetLeftSource,
    string InsetRightSource,
    string InsetTopSource,
    string InsetBottomSource,
    double FontScale,
    string? FontScaleValue,
    string FontScaleSource,
    double LineSpacingScale,
    string? LineSpacingReductionValue,
    string LineSpacingScaleSource,
    bool CompatibleLineSpacing,
    string CompatibleLineSpacingSource,
    double? RotationDegrees,
    string? RotationValue,
    string RotationDegreesSource,
    int InheritedPlaceholderCount,
    bool HasInheritedTextBody,
    bool UsesInheritedShapeBounds,
    string Orientation,
    string? OrientationValue,
    string OrientationSource,
    string VerticalAnchor,
    string? VerticalAnchorValue,
    string VerticalAnchorSource,
    bool? AnchorCenter,
    string? AnchorCenterValue,
    string AnchorCenterSource,
    string WrapMode,
    string? WrapValue,
    string WrapSource,
    string VerticalOverflow,
    string? VerticalOverflowValue,
    string VerticalOverflowSource,
    int ColumnCount,
    double ColumnSpacing,
    string ColumnSource,
    string ColumnCountSource,
    string ColumnSpacingSource,
    string? ColumnCountValue,
    string? ColumnSpacingValue,
    string AutofitModeValue,
    string AutofitModeSource,
    int ParagraphCount,
    int RunCount,
    int TextLength);

internal sealed record PptxTextParagraphRecord(
    int Slide,
    int FrameIndex,
    int ParagraphIndex,
    int Level,
    string CascadeLevelName,
    int ResolvedCascadeSourceCount,
    IReadOnlyList<string> CascadeLayerNames,
    IReadOnlyList<string> CascadeLayerKinds,
    int ResolvedStyleSourceCount,
    IReadOnlyList<string> ResolvedStyleLayerNames,
    IReadOnlyList<string> ResolvedStyleLayerKinds,
    string Alignment,
    string? AlignmentValue,
    double FontSize,
    string BulletKind,
    string? BulletCharacter,
    string? BulletResolvedCharacter,
    string? BulletAutoNumberType,
    string? BulletAutoNumberStartAtValue,
    string? BulletFontTypeface,
    string? BulletFontCharset,
    string? BulletResolvedFontTypeface,
    string BulletFontTypefaceSource,
    string? BulletColor,
    string BulletSizeKind,
    string? BulletSizeValue,
    double SpacingBefore,
    double SpacingAfter,
    double LineSpacingValue,
    string LineSpacingKind,
    bool LineSpacingUseNormalLineAdvance,
    double MarginLeft,
    double HangingIndent,
    int RunCount,
    int TextLength,
    IReadOnlyList<PptxTextRunRecord> Runs);

internal sealed record PptxTextRunRecord(
    int RunIndex,
    string Kind,
    string? Text,
    int TextLength,
    int ResolvedCascadeSourceCount,
    IReadOnlyList<string> CascadeLayerNames,
    IReadOnlyList<string> CascadeLayerKinds,
    double FontSize,
    double CharacterSpacing,
    string? Typeface,
    string ColorSource,
    bool HasHyperlinkClick,
    string? HyperlinkClickId,
    bool Underline,
    string? UnderlineValue,
    bool Strike,
    string? StrikeValue,
    string? CapsValue,
    string? HighlightColor);

internal sealed record PptxTextLineRecord(
    int Slide,
    int FrameIndex,
    int ParagraphIndex,
    int LineIndex,
    int Level,
    double TopY,
    double BaselineY,
    double Advance,
    double BaselineOffset,
    double MaxFontSize,
    string LineSpacingKind,
    string BaselineMetricSource,
    string? BaselineMetricTypeface,
    bool BaselineMetricBold,
    bool BaselineMetricItalic,
    double BaselineMetricFontSize,
    double BaselineMetricRatio,
    int BaselineMetricUnitsPerEm,
    int BaselineMetricWindowsAscender,
    int BaselineMetricWindowsDescender,
    int BaselineMetricTypographicAscender,
    int BaselineMetricTypographicDescender,
    int BaselineMetricTypographicLineGap,
    double StartX,
    double EndX,
    double NaturalEndX,
    string Alignment,
    int SpanCount,
    int TextLength,
    int GlyphCount,
    IReadOnlyList<PptxTextSpanRecord> Spans);

internal sealed record PptxTextSpanRecord(
    int SpanIndex,
    string? Text,
    int TextLength,
    double X,
    double Y,
    double Width,
    double FontSize,
    string? Typeface,
    double LeadingAdjustment,
    double NaturalWidth,
    double LayoutWidth,
    int GlyphCount,
    double FirstAdjustmentAfterOrigin);
