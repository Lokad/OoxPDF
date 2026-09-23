using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal readonly record struct PptxScenePictureTile(
    bool HasTile,
    string? TileValue,
    string? AlignmentValue,
    string? FlipValue,
    string? ScaleXValue,
    string? ScaleYValue,
    string? OffsetXValue,
    string? OffsetYValue);

internal enum PptxSceneImageRecolorKind
{
    None,
    Luminance,
    Duotone,
    Grayscale,
    BiLevel
}

internal readonly record struct PptxSceneImageRecolor(
    PptxSceneImageRecolorKind Kind,
    double Brightness,
    double Contrast,
    RgbColor Dark,
    RgbColor Light,
    double Threshold,
    string? KindValue,
    string? BrightnessValue,
    string? ContrastValue,
    string? ThresholdValue)
{
    public static PptxSceneImageRecolor None { get; } = new(PptxSceneImageRecolorKind.None, 0d, 0d, default, default, 0d, null, null, null, null);

    public static PptxSceneImageRecolor Luminance(
        double brightness,
        double contrast,
        string? kindValue,
        string? brightnessValue,
        string? contrastValue)
    {
        return new PptxSceneImageRecolor(
            PptxSceneImageRecolorKind.Luminance,
            Math.Clamp(brightness, -1d, 1d),
            Math.Clamp(contrast, -1d, 1d),
            default,
            default,
            0d,
            kindValue,
            brightnessValue,
            contrastValue,
            null);
    }

    public static PptxSceneImageRecolor Duotone(RgbColor dark, RgbColor light, string? kindValue)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.Duotone, 0d, 0d, dark, light, 0d, kindValue, null, null, null);
    }

    public static PptxSceneImageRecolor Grayscale(string? kindValue)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.Grayscale, 0d, 0d, default, default, 0d, kindValue, null, null, null);
    }

    public static PptxSceneImageRecolor BiLevel(double threshold, string? kindValue, string? thresholdValue)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.BiLevel, 0d, 0d, default, default, Math.Clamp(threshold, 0d, 1d), kindValue, null, null, thresholdValue);
    }
}

internal sealed record PptxSceneTextBody(
    XElement? BodyProperties,
    XElement? ListStyle,
    bool HasUnsupportedTextOrientation,
    bool HasUnsupportedVerticalOverflow,
    IReadOnlyList<PptxSceneTextParagraph> Paragraphs);

internal sealed record PptxSceneCascadeLayer(
    string Name,
    string Kind,
    XElement? Source);

internal sealed record PptxSceneTextParagraph(
    XElement? Properties,
    XElement? EndParagraphProperties,
    int Level,
    PptxSceneParagraphStyle ResolvedStyle,
    IReadOnlyList<PptxSceneTextRun> Runs,
    XElement? DefaultParagraphProperties,
    XElement? DefaultRunProperties,
    IReadOnlyList<PptxSceneCascadeLayer> CascadeLayers);

internal sealed record PptxSceneTextRun(
    PptxSceneTextRunKind Kind,
    string Text,
    XElement? Properties,
    PptxSceneRunStyle ResolvedStyle,
    XElement Source);

internal sealed record PptxSceneParagraphStyle(
    int Level,
    string Alignment,
    double FontSize,
    RgbColor Color,
    double Alpha,
    string? Typeface,
    PptxThemeTypefaceSource TypefaceSource,
    bool Bold,
    bool Italic,
    double CharacterSpacing);

internal sealed record PptxSceneRunStyle(
    double FontSize,
    RgbColor Color,
    double Alpha,
    string? Typeface,
    PptxThemeTypefaceSource TypefaceSource,
    bool Bold,
    bool Italic,
    bool Underline,
    string? UnderlineValue,
    bool Strike,
    string? StrikeValue,
    string? CapsValue,
    double CharacterSpacing,
    double BaselineOffset,
    RgbColor? Highlight);

internal enum PptxSceneTextRunKind
{
    Text,
    Break,
    Field
}

internal sealed record PptxSceneBounds(
    long XEmu,
    long YEmu,
    long WidthEmu,
    long HeightEmu,
    double RotationDegrees,
    bool FlipHorizontal,
    bool FlipVertical)
{
    public double X => OoxUnits.EmuToPoints(XEmu);
    public double Y => OoxUnits.EmuToPoints(YEmu);
    public double Width => OoxUnits.EmuToPoints(WidthEmu);
    public double Height => OoxUnits.EmuToPoints(HeightEmu);
}

internal enum PptxSceneNodeKind
{
    Shape,
    Picture,
    Table,
    Chart,
    Group,
    Connector,
    UnknownGraphicFrame,
    Unknown
}
