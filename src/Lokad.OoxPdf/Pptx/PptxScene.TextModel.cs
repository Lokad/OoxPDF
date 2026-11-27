using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

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
    double Brightness = 0d,
    double Contrast = 0d,
    RgbColor Dark = default,
    RgbColor Light = default,
    double Threshold = 0d);

internal sealed record PptxSceneTextBody(
    bool HasUnsupportedTextOrientation,
    IReadOnlyList<PptxSceneTextParagraph> Paragraphs);

internal sealed record PptxSceneCascadeLayer(
    string Name,
    string Kind,
    XElement? Source);

internal sealed record PptxSceneTextParagraph(
    XElement? Properties,
    XElement? EndParagraphProperties,
    int Level,
    IReadOnlyList<PptxSceneTextRun> Runs,
    IReadOnlyList<PptxSceneCascadeLayer> CascadeLayers,
    XElement? Source);

internal sealed record PptxSceneTextRun(string Text);

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
