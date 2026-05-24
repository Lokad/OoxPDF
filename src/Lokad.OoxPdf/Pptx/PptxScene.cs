using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxScene(PptxDocument Document, PptxTheme Theme, IReadOnlyList<PptxSceneSlide> Slides);

internal sealed record PptxTextRunSnapshot(
    string Text,
    double X,
    double Y,
    double Width,
    double FontSize,
    double CharacterSpacing,
    RgbColor Color,
    double Alpha,
    RgbColor? Highlight,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Strike,
    string Alignment,
    string? FontFamily);

internal sealed record PptxTextFrameModelSnapshot(
    double TextX,
    double TextWidth,
    double FontScale,
    IReadOnlyList<PptxTextParagraphModelSnapshot> Paragraphs);

internal sealed record PptxTextParagraphModelSnapshot(
    int Level,
    string CascadeLevelName,
    int ResolvedCascadeSourceCount,
    IReadOnlyList<string> CascadeLayerNames,
    string Alignment,
    double FontSize,
    IReadOnlyList<PptxTextRunModelSnapshot> Runs);

internal sealed record PptxTextRunModelSnapshot(
    string Kind,
    string Text,
    double FontSize,
    double CharacterSpacing,
    string? Typeface,
    bool Underline,
    RgbColor? Highlight);

internal sealed record PptxTextFlowSnapshot(IReadOnlyList<PptxTextFlowFrameSnapshot> Frames);

internal sealed record PptxTextFlowFrameSnapshot(
    double TextX,
    double TextWidth,
    double TextHeight,
    double ClipY,
    double ClipHeight,
    double CursorTop,
    IReadOnlyList<PptxTextFlowParagraphSnapshot> Paragraphs);

internal sealed record PptxTextFlowParagraphSnapshot(
    int Level,
    string Alignment,
    double FontSize,
    IReadOnlyList<PptxTextFlowRunSnapshot> Runs);

internal sealed record PptxTextFlowRunSnapshot(
    string SourceKind,
    string SourceText,
    double FontSize,
    string? Typeface,
    IReadOnlyList<PptxTextFlowSegmentSnapshot> Segments);

internal sealed record PptxTextFlowSegmentSnapshot(
    string Kind,
    string Text,
    string AdvanceText,
    bool Draw,
    bool PreventCoalesce,
    double FontScale);

internal sealed record PptxTextLayoutSnapshot(IReadOnlyList<PptxTextFrameLayoutSnapshot> Frames);

internal sealed record PptxTextFrameLayoutSnapshot(IReadOnlyList<PptxTextParagraphLayoutSnapshot> Paragraphs);

internal sealed record PptxTextParagraphLayoutSnapshot(
    int Level,
    IReadOnlyList<PptxTextLineLayoutSnapshot> Lines);

internal sealed record PptxTextLineLayoutSnapshot(
    double TopY,
    double BaselineY,
    double Advance,
    double BaselineOffset,
    double MaxFontSize,
    string LineSpacingKind,
    PptxTextBaselineMetricSnapshot BaselineMetric,
    double StartX,
    double EndX,
    string Alignment,
    IReadOnlyList<PptxTextSpanLayoutSnapshot> Spans);

internal sealed record PptxTextBaselineMetricSnapshot(
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

internal sealed record PptxTextSpanLayoutSnapshot(
    string? SourceText,
    string Text,
    double X,
    double Y,
    double Width,
    double FontSize,
    IReadOnlyList<PptxTextAtomLayoutSnapshot> Atoms,
    PptxTextGlyphSpanLayoutSnapshot GlyphSpan);

internal sealed record PptxTextGlyphSpanLayoutSnapshot(
    string Text,
    string? Typeface,
    double FontSize,
    double NaturalWidth,
    double LayoutWidth,
    int GlyphCount,
    double FirstAdjustmentAfterOrigin,
    IReadOnlyList<PptxTextGlyphLayoutSnapshot> Glyphs);

internal sealed record PptxTextGlyphLayoutSnapshot(
    int CodePoint,
    ushort GlyphId,
    double Advance,
    double AdjustmentBefore);

internal sealed record PptxTextAtomLayoutSnapshot(
    string Kind,
    string Text,
    double X,
    double Width,
    bool Draw);

internal sealed record PptxTextGlyphRunSnapshot(
    string Text,
    double X,
    double BaselineY,
    double Width,
    int GlyphCount,
    double FirstAdjustmentAfterOrigin);

internal sealed record PptxSceneSlide(
    int Index,
    string PartName,
    string? MasterPartName,
    string? LayoutPartName,
    XDocument SlideXml,
    PptxSceneBackground MasterBackground,
    PptxSceneBackground LayoutBackground,
    PptxSceneBackground SlideBackground,
    IReadOnlyList<PptxSceneNode> MasterNodes,
    IReadOnlyList<PptxSceneNode> LayoutNodes,
    IReadOnlyList<PptxSceneNode> SlideNodes);

internal readonly record struct PptxSceneBackground(
    bool HasFill,
    RgbColor Color,
    double Alpha);

internal sealed record PptxSceneNode(
    PptxSceneNodeKind Kind,
    string Id,
    string Name,
    bool IsPlaceholder,
    PptxSceneBounds? Bounds,
    PptxSceneShape? Shape,
    PptxSceneTextBody? TextBody,
    PptxScenePicture? Picture,
    PptxSceneTable? Table,
    PptxSceneChart? Chart,
    PptxSceneGroupTransform GroupTransform,
    IReadOnlyList<PptxSceneNode> Children,
    XElement Source);

internal sealed record PptxSceneShape(
    string Preset,
    IReadOnlyDictionary<string, double> PresetAdjustments,
    bool HasCustomGeometry,
    PptxSceneCustomGeometry CustomGeometry,
    PptxSceneFillStyle Fill,
    PptxSceneGradientFill GradientFill,
    PptxScenePatternFill PatternFill,
    PptxSceneShapePictureFill PictureFill,
    PptxSceneGlow Glow,
    PptxSceneOuterShadow OuterShadow,
    PptxSceneLineStyle Line,
    PptxSceneLineEnd HeadEnd,
    PptxSceneLineEnd TailEnd);

internal sealed record PptxSceneCustomGeometry(
    bool HasGeometry,
    IReadOnlyList<PptxSceneCustomGuide> Guides,
    IReadOnlyList<PptxSceneCustomPath> Paths);

internal sealed record PptxSceneCustomGuide(
    string Name,
    string Formula);

internal sealed record PptxSceneCustomPath(
    double Width,
    double Height,
    bool AllowsFill,
    bool AllowsStroke,
    IReadOnlyList<PptxSceneCustomCommand> Commands);

internal sealed record PptxSceneCustomCommand(
    PptxSceneCustomCommandKind Kind,
    IReadOnlyList<PptxSceneCustomPoint> Points,
    string RadiusX,
    string RadiusY,
    string StartAngle,
    string SweepAngle);

internal enum PptxSceneCustomCommandKind
{
    MoveTo,
    LineTo,
    CubicBezierTo,
    QuadraticBezierTo,
    ArcTo,
    Close
}

internal readonly record struct PptxSceneCustomPoint(
    string X,
    string Y);

internal readonly record struct PptxSceneFillStyle(
    bool HasFill,
    RgbColor Color,
    double Alpha);

internal sealed record PptxSceneGradientFill(
    bool HasGradient,
    double AngleDegrees,
    IReadOnlyList<PptxSceneGradientStop> Stops);

internal readonly record struct PptxSceneGradientStop(
    double Offset,
    RgbColor Color,
    double Alpha);

internal readonly record struct PptxScenePatternFill(
    bool HasPattern,
    string Preset,
    RgbColor Foreground,
    RgbColor Background,
    double Alpha);

internal readonly record struct PptxSceneShapePictureFill(
    bool HasPicture,
    string RelationshipId,
    PptxSceneRect Crop,
    PptxSceneRect Fill);

internal readonly record struct PptxSceneGlow(
    bool HasGlow,
    RgbColor Color,
    double Alpha,
    double Radius);

internal readonly record struct PptxSceneOuterShadow(
    bool HasShadow,
    RgbColor Color,
    double Alpha,
    double OffsetX,
    double OffsetY);

internal readonly record struct PptxSceneLineStyle(
    bool HasLine,
    RgbColor Color,
    double Width,
    double Alpha,
    IReadOnlyList<double> DashPattern,
    int? Cap,
    int? Join)
{
    public bool HasDash => DashPattern is { Count: > 0 };
}

internal enum PptxSceneLineEndKind
{
    None,
    Triangle,
    Arrow,
    Stealth,
    Diamond,
    Oval
}

internal readonly record struct PptxSceneLineEnd(PptxSceneLineEndKind Kind, double WidthScale, double LengthScale)
{
    public bool IsNone => Kind == PptxSceneLineEndKind.None;
}

internal sealed record PptxScenePicture(
    string? RelationshipId,
    PptxSceneRect Crop,
    PptxSceneRect Fill,
    double Alpha,
    PptxSceneImageRecolor Recolor);

internal sealed record PptxSceneChart(
    string? RelationshipId,
    string? TargetPartName,
    XDocument? ChartXml,
    IReadOnlyList<RgbColor>? PaletteColors,
    string StyleId,
    IReadOnlyList<PptxSceneChartPlot> Plots,
    IReadOnlyList<PptxSceneChartAxis> Axes,
    PptxSceneChartTitle Title,
    PptxSceneChartLegend Legend,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartManualLayout PlotAreaLayout,
    PptxSceneChartShapeStyle ChartAreaStyle,
    PptxSceneChartShapeStyle PlotAreaStyle);

internal sealed record PptxSceneChartPlot(
    string Kind,
    int SeriesCount,
    IReadOnlyList<string> AxisIds,
    IReadOnlyList<PptxSceneChartSeries> Series,
    string Grouping,
    string BarDirection,
    string ScatterStyle,
    bool VaryColors,
    double? GapWidth,
    double? Overlap,
    double? HoleSize,
    PptxSceneChartDataLabels DataLabels);

internal sealed record PptxSceneChartDataLabels(
    bool ShowValue,
    bool ShowPercent,
    bool ShowCategoryName,
    bool ShowSeriesName,
    bool ShowLeaderLines,
    string Position,
    string Separator,
    string NumberFormat,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartShapeStyle ShapeStyle,
    IReadOnlyList<PptxSceneChartDataLabelOverride> Overrides,
    bool IsDefined);

internal sealed record PptxSceneChartDataLabelOverride(
    int Index,
    bool? ShowValue,
    bool? ShowPercent,
    bool? ShowCategoryName,
    bool? ShowSeriesName,
    bool? ShowLeaderLines,
    string CustomText,
    string Position,
    string Separator,
    string NumberFormat,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartShapeStyle ShapeStyle);

internal sealed record PptxSceneChartShapeStyle(
    bool NoFill,
    PptxSceneFillStyle Fill,
    PptxScenePatternFill PatternFill,
    PptxSceneLineStyle Line);

internal readonly record struct PptxSceneChartTextStyleOverride(
    string? FontFamily,
    double? FontSize,
    RgbColor? Color,
    bool? Bold,
    bool? Italic);

internal readonly record struct PptxSceneChartManualLayout(
    bool HasLayout,
    double X,
    double Y,
    double Width,
    double Height,
    string LayoutTarget,
    string XMode,
    string YMode,
    string WidthMode,
    string HeightMode);

internal sealed record PptxSceneChartSeries(
    string? Name,
    IReadOnlyList<double> Values,
    IReadOnlyList<string> Categories,
    IReadOnlyList<double> XValues,
    IReadOnlyList<double> YValues,
    IReadOnlyList<double> BubbleSizes,
    PptxSceneFillStyle Fill,
    PptxScenePatternFill PatternFill,
    PptxSceneLineStyle Line,
    PptxSceneChartMarker Marker,
    IReadOnlyList<PptxSceneChartPointStyle> PointStyles,
    bool Smooth,
    PptxSceneChartDataLabels DataLabels);

internal sealed record PptxSceneChartMarker(
    string Symbol,
    double Size,
    PptxSceneFillStyle Fill,
    PptxSceneLineStyle Line);

internal sealed record PptxSceneChartPointStyle(
    int Index,
    PptxSceneFillStyle Fill,
    PptxScenePatternFill PatternFill,
    PptxSceneLineStyle Line,
    double? Explosion);

internal sealed record PptxSceneChartAxis(
    string Id,
    string Kind,
    string Position,
    string CrossAxisId,
    string Crosses,
    double? CrossesAt,
    string CrossBetween,
    bool IsReversed,
    bool IsDeleted,
    bool HasScaling,
    double? Minimum,
    double? Maximum,
    double? MajorUnit,
    double? MinorUnit,
    bool HasMajorGridlines,
    bool HasMinorGridlines,
    PptxSceneLineStyle Line,
    PptxSceneLineStyle MajorGridlineLine,
    PptxSceneLineStyle MinorGridlineLine,
    PptxSceneChartTextStyleOverride TextStyle,
    string TickLabelPosition,
    string MajorTickMark,
    string MinorTickMark,
    int? LabelOffset,
    int? TickLabelSkip,
    int? TickMarkSkip,
    bool NoMultiLevelLabels,
    string? NumberFormat);

internal sealed record PptxSceneChartTitle(
    string? Text,
    bool IsAutoDeleted,
    PptxSceneChartTextStyleOverride TextStyle);

internal sealed record PptxSceneChartLegend(
    string Position,
    bool Overlay,
    bool IsVisible,
    PptxSceneChartTextStyleOverride TextStyle);

internal sealed record PptxSceneTable(
    IReadOnlyList<double> ColumnWidths,
    IReadOnlyList<double> RowHeights,
    IReadOnlyList<PptxSceneTableRow> Rows,
    PptxSceneTableStyle Style,
    XElement? Source);

internal readonly record struct PptxSceneTableStyle(
    string? StyleId,
    string Name,
    string Accent,
    bool IsSupported,
    bool FirstRow,
    bool LastRow,
    bool FirstColumn,
    bool LastColumn,
    bool BandRow,
    bool BandColumn)
{
    public bool HasStyle => StyleId is not null;
}

internal readonly record struct PptxBuiltInTableStyle(string Name, string Accent);

internal static class PptxBuiltInTableStyles
{
    public static bool TryGet(string? styleId, out PptxBuiltInTableStyle style)
    {
        return Styles.TryGetValue(styleId ?? string.Empty, out style);
    }

    private static IReadOnlyDictionary<string, PptxBuiltInTableStyle> Styles { get; } =
        new Dictionary<string, PptxBuiltInTableStyle>(StringComparer.OrdinalIgnoreCase)
        {
            ["{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}"] = new("Light-Style-1", "tx1"),
            ["{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"] = new("Light-Style-1", "accent1"),
            ["{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}"] = new("Light-Style-1", "accent2"),
            ["{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}"] = new("Light-Style-1", "accent3"),
            ["{D27102A9-8310-4765-A935-A1911B00CA55}"] = new("Light-Style-1", "accent4"),
            ["{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}"] = new("Light-Style-1", "accent5"),
            ["{68D230F3-CF80-4859-8CE7-A43EE81993B5}"] = new("Light-Style-1", "accent6"),
            ["{E8034E78-7F5D-4C2E-B375-FC64B27BC917}"] = new("Dark-Style-1", "dk1"),
            ["{125E5076-3810-47DD-B79F-674D7AD40C01}"] = new("Dark-Style-1", "accent1"),
            ["{37CE84F3-28C3-443E-9E96-99CF82512B78}"] = new("Dark-Style-1", "accent2"),
            ["{D03447BB-5D67-496B-8E87-E561075AD55C}"] = new("Dark-Style-1", "accent3"),
            ["{E929F9F4-4A8F-4326-A1B4-22849713DDAB}"] = new("Dark-Style-1", "accent4"),
            ["{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}"] = new("Dark-Style-1", "accent5"),
            ["{AF606853-7671-496A-8E4F-DF71F8EC918B}"] = new("Dark-Style-1", "accent6"),
            ["{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"] = new("Medium-Style-2", "tx1"),
            ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"] = new("Medium-Style-2", "accent1"),
            ["{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}"] = new("Medium-Style-2", "accent2"),
            ["{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}"] = new("Medium-Style-2", "accent3"),
            ["{00A15C55-8517-42AA-B614-E9B94910E393}"] = new("Medium-Style-2", "accent4"),
            ["{7DF18680-E054-41AD-8BC1-D1AEF772440D}"] = new("Medium-Style-2", "accent5"),
            ["{93296810-A885-4BE3-A3E7-6D5BEEA58F35}"] = new("Medium-Style-2", "accent6")
        };
}

internal sealed record PptxSceneTableRow(IReadOnlyList<PptxSceneTableCell> Cells);

internal readonly record struct PptxSceneTableCell(
    int ColumnSpan,
    int RowSpan,
    bool IsMergedContinuation,
    PptxSceneTextInsets TextInsets,
    PptxSceneTableCellVerticalAnchor VerticalAnchor,
    PptxSceneFillStyle Fill,
    PptxSceneTableCellBorders Borders,
    PptxSceneFillStyle StyleFill,
    PptxSceneTableCellTextStyle StyleText,
    XElement? TextBody);

internal readonly record struct PptxSceneTableCellTextStyle(RgbColor? Color, bool Bold);

internal readonly record struct PptxSceneTableCellBorders(
    PptxSceneTableCellBorder Left,
    PptxSceneTableCellBorder Right,
    PptxSceneTableCellBorder Top,
    PptxSceneTableCellBorder Bottom)
{
    public bool HasExplicitBorder => Left.IsSpecified || Right.IsSpecified || Top.IsSpecified || Bottom.IsSpecified;
}

internal readonly record struct PptxSceneTableCellBorder(bool IsSpecified, PptxSceneLineStyle Line);

internal readonly record struct PptxSceneTextInsets(double Left, double Right, double Top, double Bottom);

internal enum PptxSceneTableCellVerticalAnchor
{
    Top,
    Middle,
    Bottom
}

internal readonly record struct PptxSceneGroupTransform(
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
    public static PptxSceneGroupTransform Identity { get; } = new(0, 0, 0, 0, 0, 0, 1d, 1d, 0d, FlipHorizontal: false, FlipVertical: false);
}

internal readonly record struct PptxSceneRect(double Left, double Top, double Right, double Bottom)
{
    public bool IsEmpty => Left == 0d && Top == 0d && Right == 0d && Bottom == 0d;
}

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
    double Threshold)
{
    public static PptxSceneImageRecolor None { get; } = new(PptxSceneImageRecolorKind.None, 0d, 0d, default, default, 0d);

    public static PptxSceneImageRecolor Luminance(double brightness, double contrast)
    {
        return new PptxSceneImageRecolor(
            PptxSceneImageRecolorKind.Luminance,
            Math.Clamp(brightness, -1d, 1d),
            Math.Clamp(contrast, -1d, 1d),
            default,
            default,
            0d);
    }

    public static PptxSceneImageRecolor Duotone(RgbColor dark, RgbColor light)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.Duotone, 0d, 0d, dark, light, 0d);
    }

    public static PptxSceneImageRecolor Grayscale()
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.Grayscale, 0d, 0d, default, default, 0d);
    }

    public static PptxSceneImageRecolor BiLevel(double threshold)
    {
        return new PptxSceneImageRecolor(PptxSceneImageRecolorKind.BiLevel, 0d, 0d, default, default, Math.Clamp(threshold, 0d, 1d));
    }
}

internal sealed record PptxSceneTextBody(
    XElement? BodyProperties,
    XElement? ListStyle,
    IReadOnlyList<PptxSceneTextParagraph> Paragraphs);

internal sealed record PptxSceneTextParagraph(
    XElement? Properties,
    XElement? EndParagraphProperties,
    int Level,
    PptxSceneParagraphStyle ResolvedStyle,
    IReadOnlyList<PptxSceneTextRun> Runs);

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
    bool Bold,
    bool Italic,
    double CharacterSpacing);

internal sealed record PptxSceneRunStyle(
    double FontSize,
    RgbColor Color,
    double Alpha,
    string? Typeface,
    bool Bold,
    bool Italic,
    bool Underline,
    bool Strike,
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

internal sealed class PptxSceneBuilder
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const double MinimumStrokeWidth = 0.1d;
    private const double SceneEffectTolerance = 0.001d;
    private const string SlideLayoutRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
    private const string ChartColorStyleRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

    public PptxScene Build(PptxDocument document, OoxPackage package)
    {
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        var slides = new List<PptxSceneSlide>(document.Slides.Count);
        foreach (PptxSlide slide in document.Slides)
        {
            OoxPart? slidePart = package.GetPart(slide.PartName);
            if (slidePart is null)
            {
                slides.Add(new PptxSceneSlide(slide.Index, slide.PartName, null, null, new XDocument(), default, default, default, [], [], []));
                continue;
            }

            XDocument slideXml = LoadXml(slidePart);
            OoxPart? layoutPart = GetRelatedPart(package, slide.PartName, SlideLayoutRelationshipType);
            OoxPart? masterPart = layoutPart is null ? null : GetRelatedPart(package, layoutPart.Name, SlideMasterRelationshipType);
            XDocument? masterXml = masterPart is null ? null : LoadXml(masterPart);
            XDocument? layoutXml = layoutPart is null ? null : LoadXml(layoutPart);
            IReadOnlyDictionary<string, OoxRelationship> masterRelationships = masterPart is null ? new Dictionary<string, OoxRelationship>() : ReadRelationships(package, masterPart.Name);
            IReadOnlyDictionary<string, OoxRelationship> layoutRelationships = layoutPart is null ? new Dictionary<string, OoxRelationship>() : ReadRelationships(package, layoutPart.Name);
            IReadOnlyDictionary<string, OoxRelationship> slideRelationships = ReadRelationships(package, slide.PartName);
            IReadOnlyList<XDocument> layoutSources = masterXml is null ? [] : [masterXml];
            IReadOnlyList<XDocument> slideSources = masterXml is null
                ? layoutXml is null ? [] : [layoutXml]
                : layoutXml is null ? [masterXml] : [masterXml, layoutXml];
            slides.Add(new PptxSceneSlide(
                slide.Index,
                slide.PartName,
                masterPart?.Name,
                layoutPart?.Name,
                slideXml,
                ReadBackground(masterXml, theme),
                ReadBackground(layoutXml, theme),
                ReadBackground(slideXml, theme),
                masterXml is null ? [] : ReadNodes(masterXml, [], theme, package, masterRelationships),
                layoutXml is null ? [] : ReadNodes(layoutXml, layoutSources, theme, package, layoutRelationships),
                ReadNodes(slideXml, slideSources, theme, package, slideRelationships)));
        }

        return new PptxScene(document, theme, slides);
    }

    private static XDocument LoadXml(OoxPart part)
    {
        using Stream stream = part.OpenRead();
        return SafeXml.Load(stream);
    }

    private static OoxPart? GetRelatedPart(OoxPackage package, string sourcePartName, string relationshipType)
    {
        OoxRelationship? relationship = package.GetRelationships(sourcePartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null ? null : package.GetPart(relationship.ResolvedTarget);
    }

    private static IReadOnlyDictionary<string, OoxRelationship> ReadRelationships(OoxPackage package, string sourcePartName)
    {
        return package.GetRelationships(sourcePartName)
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
    }

    private static IReadOnlyList<PptxSceneNode> ReadNodes(
        XDocument xml,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var nodes = new List<PptxSceneNode>();
        foreach (XElement shapeTree in xml.Descendants(PresentationNamespace + "spTree"))
        {
            nodes.AddRange(ReadChildNodes(shapeTree, placeholderSources, theme, package, relationships));
        }

        return nodes;
    }

    private static PptxSceneBackground ReadBackground(XDocument? xml, PptxTheme theme)
    {
        XElement? background = xml?.Root?
            .Element(PresentationNamespace + "cSld")?
            .Element(PresentationNamespace + "bg")?
            .Element(PresentationNamespace + "bgPr");
        return TryReadSolidColorWithAlpha(background, theme, out RgbColor color, out double alpha)
            ? new PptxSceneBackground(true, color, alpha)
            : default;
    }

    private static IReadOnlyList<PptxSceneNode> ReadChildNodes(
        XElement container,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var nodes = new List<PptxSceneNode>();
        foreach (XElement child in container.Elements())
        {
            PptxSceneNodeKind kind = ReadNodeKind(child);
            if (kind == PptxSceneNodeKind.Unknown)
            {
                continue;
            }

            (string id, string name) = ReadNonVisualProperties(child);
            nodes.Add(new PptxSceneNode(
                kind,
                id,
                name,
                IsPlaceholder(child),
                ReadBounds(child),
                kind is PptxSceneNodeKind.Shape or PptxSceneNodeKind.Connector ? ReadShape(child, theme) : null,
                ReadTextBody(child, placeholderSources, theme),
                kind == PptxSceneNodeKind.Picture ? ReadPicture(child, theme) : null,
                kind == PptxSceneNodeKind.Table ? ReadTable(child, theme) : null,
                kind == PptxSceneNodeKind.Chart ? ReadChart(child, package, theme, relationships) : null,
                kind == PptxSceneNodeKind.Group ? ReadGroupTransform(child) : PptxSceneGroupTransform.Identity,
                kind == PptxSceneNodeKind.Group ? ReadChildNodes(child, placeholderSources, theme, package, relationships) : [],
                child));
        }

        return nodes;
    }

    internal static PptxSceneNodeKind ReadNodeKind(XElement element)
    {
        if (element.Name == PresentationNamespace + "sp")
        {
            return PptxSceneNodeKind.Shape;
        }

        if (element.Name == PresentationNamespace + "cxnSp")
        {
            return PptxSceneNodeKind.Connector;
        }

        if (element.Name == PresentationNamespace + "pic")
        {
            return PptxSceneNodeKind.Picture;
        }

        if (element.Name == PresentationNamespace + "grpSp")
        {
            return PptxSceneNodeKind.Group;
        }

        if (element.Name != PresentationNamespace + "graphicFrame")
        {
            return PptxSceneNodeKind.Unknown;
        }

        XElement? graphicData = element
            .Descendants(DrawingNamespace + "graphicData")
            .FirstOrDefault();
        string uri = (string?)graphicData?.Attribute("uri") ?? string.Empty;
        if (graphicData?.Descendants(DrawingNamespace + "tbl").Any() == true)
        {
            return PptxSceneNodeKind.Table;
        }

        return uri.Contains("chart", StringComparison.OrdinalIgnoreCase)
            ? PptxSceneNodeKind.Chart
            : PptxSceneNodeKind.UnknownGraphicFrame;
    }

    private static (string Id, string Name) ReadNonVisualProperties(XElement element)
    {
        XElement? nonVisual = element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "cNvPr");
        return ((string?)nonVisual?.Attribute("id") ?? string.Empty, (string?)nonVisual?.Attribute("name") ?? string.Empty);
    }

    private static bool IsPlaceholder(XElement element)
    {
        return element
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName is "nvSpPr" or "nvPicPr" or "nvGrpSpPr" or "nvGraphicFramePr" or "nvCxnSpPr")
            ?.Elements()
            .FirstOrDefault(e => e.Name.LocalName == "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
    }

    private static PptxSceneBounds? ReadBounds(XElement element)
    {
        XElement? transform = element
            .Element(PresentationNamespace + "spPr")?
            .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "grpSpPr")?
                .Element(DrawingNamespace + "xfrm") ??
            element.Element(PresentationNamespace + "xfrm");
        if (transform is null)
        {
            return null;
        }

        XElement? offset = transform.Element(DrawingNamespace + "off");
        XElement? extents = transform.Element(DrawingNamespace + "ext");
        if (offset is null || extents is null)
        {
            return null;
        }

        return new PptxSceneBounds(
            ReadLong(offset, "x"),
            ReadLong(offset, "y"),
            ReadLong(extents, "cx"),
            ReadLong(extents, "cy"),
            transform.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            ReadBool(transform, "flipH"),
            ReadBool(transform, "flipV"));
    }

    internal static string? ReadPictureRelationshipId(XElement picture)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        return (string?)blip?.Attribute(RelationshipsNamespace + "embed") ??
            blip?.Descendants()
                .FirstOrDefault(element => element.Name.LocalName == "svgBlip")
                ?.Attribute(RelationshipsNamespace + "embed")
                ?.Value;
    }

    private static PptxScenePicture ReadPicture(XElement picture, PptxTheme theme)
    {
        return new PptxScenePicture(
            ReadPictureRelationshipId(picture),
            ReadPictureCrop(picture),
            ReadPictureFill(picture),
            ReadPictureAlpha(picture),
            ReadImageRecolor(picture, theme));
    }

    private static PptxSceneChart ReadChart(
        XElement frame,
        OoxPackage package,
        PptxTheme theme,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement? graphicData = frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData");
        string? relationshipId = (string?)graphicData
            ?.Element(ChartNamespace + "chart")
            ?.Attribute(RelationshipsNamespace + "id");
        string? targetPartName = relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship)
            ? relationship.ResolvedTarget
            : null;
        OoxPart? chartPart = targetPartName is null ? null : package.GetPart(targetPartName);
        XDocument? chartXml = chartPart is null ? null : LoadXml(chartPart);
        IReadOnlyList<RgbColor>? paletteColors = chartPart is null ? null : ReadChartPaletteColors(package, chartPart.Name, theme);
        return new PptxSceneChart(
            relationshipId,
            targetPartName,
            chartXml,
            paletteColors,
            (string?)chartXml?.Root?.Element(ChartNamespace + "style")?.Attribute("val") ?? string.Empty,
            ReadChartPlots(chartXml, theme),
            ReadChartAxes(chartXml, theme),
            ReadChartTitle(chartXml, theme),
            ReadChartLegend(chartXml, theme),
            ReadChartTextStyleOverride(chartXml?.Root, theme),
            ReadChartPlotAreaManualLayout(chartXml),
            ReadChartShapeStyle(chartXml?.Root?.Element(ChartNamespace + "spPr"), theme),
            ReadChartShapeStyle(chartXml?
                .Descendants(ChartNamespace + "plotArea")
                .FirstOrDefault()
                ?.Element(ChartNamespace + "spPr"), theme));
    }

    private static IReadOnlyList<PptxSceneChartPlot> ReadChartPlots(XDocument? chartXml, PptxTheme theme)
    {
        XElement? plotArea = chartXml?
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        if (plotArea is null)
        {
            return [];
        }

        var plots = new List<PptxSceneChartPlot>();
        foreach (XElement plot in plotArea.Elements().Where(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal)))
        {
            string[] axisIds = plot
                .Elements(ChartNamespace + "axId")
                .Select(axis => (string?)axis.Attribute("val") ?? string.Empty)
                .Where(value => value.Length != 0)
                .ToArray();
            plots.Add(new PptxSceneChartPlot(
                plot.Name.LocalName,
                plot.Elements(ChartNamespace + "ser").Count(),
                axisIds,
                ReadChartSeries(plot, theme),
                ReadChartElementValue(plot, "grouping"),
                ReadChartElementValue(plot, "barDir"),
                ReadChartElementValue(plot, "scatterStyle"),
                ReadChartPlotVaryColors(plot),
                ReadChartElementDouble(plot, "gapWidth"),
                ReadChartElementDouble(plot, "overlap"),
                ReadChartElementDouble(plot, "holeSize"),
                ReadChartDataLabels(plot, theme)));
        }

        return plots;
    }

    private static PptxSceneChartDataLabels ReadChartDataLabels(XElement plot, PptxTheme theme)
    {
        XElement? labels = plot.Element(ChartNamespace + "dLbls") ??
            plot.Elements(ChartNamespace + "ser")
                .Select(series => series.Element(ChartNamespace + "dLbls"))
                .FirstOrDefault(element => element is not null);
        return labels is null
            ? new PptxSceneChartDataLabels(
                ShowValue: false,
                ShowPercent: false,
                ShowCategoryName: false,
                ShowSeriesName: false,
                ShowLeaderLines: false,
                Position: string.Empty,
                Separator: string.Empty,
                NumberFormat: string.Empty,
                TextStyle: new PptxSceneChartTextStyleOverride(null, null, null, null, null),
                ShapeStyle: new PptxSceneChartShapeStyle(false, default, default, default),
                Overrides: [],
                IsDefined: false)
            : new PptxSceneChartDataLabels(
                IsOoxmlBooleanElementEnabled(labels.Element(ChartNamespace + "showVal")),
                IsOoxmlBooleanElementEnabled(labels.Element(ChartNamespace + "showPercent")),
                IsOoxmlBooleanElementEnabled(labels.Element(ChartNamespace + "showCatName")),
                IsOoxmlBooleanElementEnabled(labels.Element(ChartNamespace + "showSerName")),
                IsOoxmlBooleanElementEnabled(labels.Element(ChartNamespace + "showLeaderLines")),
                ReadChartElementValue(labels, "dLblPos"),
                labels.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                labels.Element(ChartNamespace + "numFmt")?.Attribute("formatCode")?.Value ?? string.Empty,
                ReadChartTextStyleOverride(labels, theme),
                ReadChartShapeStyle(labels.Element(ChartNamespace + "spPr"), theme),
                ReadChartDataLabelOverrides(labels, theme),
                IsDefined: true);
    }

    private static IReadOnlyList<PptxSceneChartDataLabelOverride> ReadChartDataLabelOverrides(XElement labels, PptxTheme theme)
    {
        var overrides = new List<PptxSceneChartDataLabelOverride>();
        foreach (XElement label in labels.Elements(ChartNamespace + "dLbl"))
        {
            if (!int.TryParse(label.Element(ChartNamespace + "idx")?.Attribute("val")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) ||
                index < 0)
            {
                continue;
            }

            overrides.Add(new PptxSceneChartDataLabelOverride(
                index,
                ReadOptionalOoxmlBooleanElement(label, "showVal"),
                ReadOptionalOoxmlBooleanElement(label, "showPercent"),
                ReadOptionalOoxmlBooleanElement(label, "showCatName"),
                ReadOptionalOoxmlBooleanElement(label, "showSerName"),
                ReadOptionalOoxmlBooleanElement(label, "showLeaderLines"),
                ReadChartText(label.Element(ChartNamespace + "tx")) ?? string.Empty,
                ReadChartElementValue(label, "dLblPos"),
                label.Element(ChartNamespace + "separator")?.Value ?? string.Empty,
                label.Element(ChartNamespace + "numFmt")?.Attribute("formatCode")?.Value ?? string.Empty,
                ReadChartTextStyleOverride(label, theme),
                ReadChartShapeStyle(label.Element(ChartNamespace + "spPr"), theme)));
        }

        return overrides;
    }

    private static bool? ReadOptionalOoxmlBooleanElement(XElement parent, string elementName)
    {
        XElement? element = parent.Element(ChartNamespace + elementName);
        return element is null ? null : IsOoxmlBooleanElementEnabled(element);
    }

    private static PptxSceneChartShapeStyle ReadChartShapeStyle(XElement? shapeProperties, PptxTheme theme)
    {
        bool noFill = shapeProperties?.Element(DrawingNamespace + "noFill") is not null;
        PptxSceneFillStyle fill = !noFill && TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
            ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
            : default;
        return new PptxSceneChartShapeStyle(
            noFill,
            fill,
            noFill ? default : ReadChartPatternFill(shapeProperties, theme),
            ReadChartLine(shapeProperties, theme));
    }

    private static string ReadChartElementValue(XElement element, string childName)
    {
        return (string?)element.Element(ChartNamespace + childName)?.Attribute("val") ?? string.Empty;
    }

    private static double? ReadChartElementDouble(XElement element, string childName)
    {
        string? value = (string?)element.Element(ChartNamespace + childName)?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? parsed
            : null;
    }

    private static int? ReadChartElementInt(XElement element, string childName)
    {
        string? value = (string?)element.Element(ChartNamespace + childName)?.Attribute("val");
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : null;
    }

    private static bool ReadChartPlotVaryColors(XElement plot)
    {
        return IsOoxmlBooleanElementEnabled(plot.Element(ChartNamespace + "varyColors"), defaultValue: true);
    }

    private static IReadOnlyList<PptxSceneChartSeries> ReadChartSeries(XElement plot, PptxTheme theme)
    {
        var series = new List<PptxSceneChartSeries>();
        foreach (XElement seriesElement in plot.Elements(ChartNamespace + "ser"))
        {
            series.Add(new PptxSceneChartSeries(
                ReadChartSeriesName(seriesElement),
                ReadChartSeriesValues(seriesElement),
                ReadChartSeriesCategories(seriesElement),
                ReadChartSeriesNumbers(seriesElement, "xVal"),
                ReadChartSeriesNumbers(seriesElement, "yVal"),
                ReadChartSeriesNumbers(seriesElement, "bubbleSize"),
                ReadChartSeriesFill(seriesElement, theme),
                ReadChartSeriesPatternFill(seriesElement, theme),
                ReadChartSeriesLine(seriesElement, theme),
                ReadChartMarker(seriesElement, theme),
                ReadChartPointStyles(seriesElement, theme),
                ReadChartSeriesSmooth(seriesElement),
                ReadChartDataLabels(seriesElement, theme)));
        }

        return series;
    }

    private static PptxSceneFillStyle ReadChartSeriesFill(XElement series, PptxTheme theme)
    {
        XElement? shapeProperties = series.Element(ChartNamespace + "spPr");
        return TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    private static PptxSceneLineStyle ReadChartSeriesLine(XElement series, PptxTheme theme)
    {
        return ReadChartLine(series.Element(ChartNamespace + "spPr"), theme);
    }

    private static PptxScenePatternFill ReadChartSeriesPatternFill(XElement series, PptxTheme theme)
    {
        return ReadChartPatternFill(series.Element(ChartNamespace + "spPr"), theme);
    }

    private static PptxScenePatternFill ReadChartPatternFill(XElement? shapeProperties, PptxTheme theme)
    {
        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        if (patternFill is null)
        {
            return default;
        }

        RgbColor foreground = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "fgClr"), theme, out RgbColor foregroundColor, out _)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "bgClr"), theme, out RgbColor backgroundColor, out _)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        return new PptxScenePatternFill(
            HasPattern: true,
            (string?)patternFill.Attribute("prst") ?? "pct50",
            foreground,
            background,
            Alpha: 1d);
    }

    private static PptxSceneChartMarker ReadChartMarker(XElement series, PptxTheme theme)
    {
        XElement? marker = series.Element(ChartNamespace + "marker");
        string symbol = (string?)marker?.Element(ChartNamespace + "symbol")?.Attribute("val") ?? "circle";
        double size = marker?.Element(ChartNamespace + "size")?.Attribute("val") is { } value &&
            double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
                ? Math.Clamp(parsed, 2d, 30d)
                : 4d;
        XElement? shapeProperties = marker?.Element(ChartNamespace + "spPr");
        PptxSceneFillStyle fill = TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
            ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
            : default;
        return new PptxSceneChartMarker(symbol, size, fill, ReadChartLine(shapeProperties, theme));
    }

    private static IReadOnlyList<PptxSceneChartPointStyle> ReadChartPointStyles(XElement series, PptxTheme theme)
    {
        var styles = new List<PptxSceneChartPointStyle>();
        foreach (XElement point in series.Elements(ChartNamespace + "dPt"))
        {
            if (point.Element(ChartNamespace + "idx")?.Attribute("val") is not { } indexAttribute ||
                !int.TryParse(indexAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                continue;
            }

            XElement? shapeProperties = point.Element(ChartNamespace + "spPr");
            styles.Add(new PptxSceneChartPointStyle(
                index,
                ReadChartPointFill(shapeProperties, theme),
                ReadChartPointPatternFill(shapeProperties, theme),
                ReadChartLine(shapeProperties, theme),
                ReadChartPointExplosion(point)));
        }

        return styles;
    }

    private static PptxSceneFillStyle ReadChartPointFill(XElement? shapeProperties, PptxTheme theme)
    {
        return TryReadSolidColorWithAlpha(shapeProperties, theme, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    private static PptxScenePatternFill ReadChartPointPatternFill(XElement? shapeProperties, PptxTheme theme)
    {
        return ReadChartPatternFill(shapeProperties, theme);
    }

    private static double? ReadChartPointExplosion(XElement point)
    {
        return point.Element(ChartNamespace + "explosion")?.Attribute("val") is { } explosionAttribute &&
            double.TryParse(explosionAttribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double explosion)
                ? explosion
                : null;
    }

    private static PptxSceneLineStyle ReadChartLine(XElement? shapeProperties, PptxTheme theme)
    {
        return shapeProperties is not null &&
            TryReadLineWithAlpha(shapeProperties, theme, out RgbColor color, out double lineWidth, out double alpha)
                ? new PptxSceneLineStyle(
                    true,
                    color,
                    lineWidth,
                    alpha,
                    TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern) ? dashPattern : [],
                    ReadLineCap(shapeProperties) switch
                    {
                        "rnd" => 1,
                        "sq" => 2,
                        _ => null
                    },
                    ReadLineJoin(shapeProperties))
                : default;
    }

    private static bool ReadChartSeriesSmooth(XElement series)
    {
        string? value = (string?)series.Element(ChartNamespace + "smooth")?.Attribute("val");
        return value is "1" or "true";
    }

    private static string? ReadChartSeriesName(XElement series)
    {
        return ReadChartText(series.Element(ChartNamespace + "tx"));
    }

    private static string? ReadChartText(XElement? text)
    {
        string? literal = text?
            .Descendants(ChartNamespace + "v")
            .Select(value => value.Value)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (!string.IsNullOrWhiteSpace(literal))
        {
            return literal;
        }

        string? richText = text?
            .Descendants(DrawingNamespace + "t")
            .Aggregate(string.Empty, (current, textElement) => current + textElement.Value);
        return string.IsNullOrWhiteSpace(richText) ? null : richText;
    }

    private static IReadOnlyList<double> ReadChartSeriesValues(XElement series)
    {
        return ReadChartSeriesNumbers(series, "val");
    }

    private static IReadOnlyList<double> ReadChartSeriesNumbers(XElement series, string elementName)
    {
        return series
            .Elements(ChartNamespace + elementName)
            .Descendants(ChartNamespace + "pt")
            .Select(point => (string?)point.Element(ChartNamespace + "v"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ? parsed : double.NaN)
            .Where(value => !double.IsNaN(value))
            .ToArray();
    }

    private static IReadOnlyList<string> ReadChartSeriesCategories(XElement series)
    {
        return series
            .Elements(ChartNamespace + "cat")
            .Descendants(ChartNamespace + "pt")
            .Select(point => (string?)point.Element(ChartNamespace + "v"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value!)
            .ToArray();
    }

    private static IReadOnlyList<PptxSceneChartAxis> ReadChartAxes(XDocument? chartXml, PptxTheme theme)
    {
        XElement? plotArea = chartXml?
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault();
        if (plotArea is null)
        {
            return [];
        }

        var axes = new List<PptxSceneChartAxis>();
        foreach (XElement axis in plotArea.Elements().Where(element => element.Name.Namespace == ChartNamespace && element.Name.LocalName.EndsWith("Ax", StringComparison.Ordinal)))
        {
            string id = (string?)axis.Element(ChartNamespace + "axId")?.Attribute("val") ?? string.Empty;
            if (id.Length == 0)
            {
                continue;
            }

            axes.Add(new PptxSceneChartAxis(
                id,
                axis.Name.LocalName,
                (string?)axis.Element(ChartNamespace + "axPos")?.Attribute("val") ?? string.Empty,
                (string?)axis.Element(ChartNamespace + "crossAx")?.Attribute("val") ?? string.Empty,
                (string?)axis.Element(ChartNamespace + "crosses")?.Attribute("val") ?? string.Empty,
                ReadChartElementDouble(axis, "crossesAt"),
                (string?)axis.Element(ChartNamespace + "crossBetween")?.Attribute("val") ?? string.Empty,
                string.Equals(
                    (string?)axis.Element(ChartNamespace + "scaling")?.Element(ChartNamespace + "orientation")?.Attribute("val"),
                    "maxMin",
                    StringComparison.Ordinal),
                IsOoxmlBooleanElementEnabled(axis.Element(ChartNamespace + "delete")),
                axis.Element(ChartNamespace + "scaling") is not null,
                ReadChartAxisScalingValue(axis, "min"),
                ReadChartAxisScalingValue(axis, "max"),
                ReadChartAxisUnitValue(axis, "majorUnit"),
                ReadChartAxisUnitValue(axis, "minorUnit"),
                IsChartGridlineVisible(axis.Element(ChartNamespace + "majorGridlines")),
                IsChartGridlineVisible(axis.Element(ChartNamespace + "minorGridlines")),
                ReadChartAxisLine(axis, theme),
                ReadChartGridlineLine(axis.Element(ChartNamespace + "majorGridlines"), theme),
                ReadChartGridlineLine(axis.Element(ChartNamespace + "minorGridlines"), theme),
                ReadChartTextStyleOverride(axis, theme),
                (string?)axis.Element(ChartNamespace + "tickLblPos")?.Attribute("val") ?? string.Empty,
                ReadChartElementValue(axis, "majorTickMark"),
                ReadChartElementValue(axis, "minorTickMark"),
                ReadChartElementInt(axis, "lblOffset"),
                ReadChartElementInt(axis, "tickLblSkip"),
                ReadChartElementInt(axis, "tickMarkSkip"),
                IsOoxmlBooleanElementEnabled(axis.Element(ChartNamespace + "noMultiLvlLbl")),
                ReadChartAxisNumberFormat(axis)));
        }

        return axes;
    }

    private static PptxSceneLineStyle ReadChartAxisLine(XElement axis, PptxTheme theme)
    {
        XElement? shapeProperties = axis.Element(ChartNamespace + "spPr");
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new PptxSceneLineStyle(true, new RgbColor(0, 0, 0), 0d, 0d, [], null, null);
        }

        return ReadChartLine(shapeProperties, theme);
    }

    private static PptxSceneChartTextStyleOverride ReadChartTextStyleOverride(XElement? parent, PptxTheme theme)
    {
        XElement? defaultRunProperties = parent?
            .Element(ChartNamespace + "txPr")?
            .Elements(DrawingNamespace + "p")
            .Select(paragraph => paragraph.Element(DrawingNamespace + "pPr")?.Element(DrawingNamespace + "defRPr"))
            .FirstOrDefault(element => element is not null);
        if (defaultRunProperties is null)
        {
            return default;
        }

        string? typeface = (string?)defaultRunProperties.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)defaultRunProperties.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)defaultRunProperties.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        string? fontFamily = string.IsNullOrWhiteSpace(typeface)
            ? null
            : theme.ResolveTypeface(typeface);
        double? fontSize = defaultRunProperties.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0
                ? sizeHundredths / 100d
                : null;
        RgbColor? color = TryReadSolidColorWithAlpha(defaultRunProperties.Element(DrawingNamespace + "solidFill"), theme, out RgbColor parsedColor, out _)
            ? parsedColor
            : null;
        bool? bold = ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "b");
        bool? italic = ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "i");
        return new PptxSceneChartTextStyleOverride(fontFamily, fontSize, color, bold, italic);
    }

    private static PptxSceneChartManualLayout ReadChartPlotAreaManualLayout(XDocument? chartXml)
    {
        XElement? manualLayout = chartXml?
            .Descendants(ChartNamespace + "plotArea")
            .FirstOrDefault()
            ?.Element(ChartNamespace + "layout")
            ?.Element(ChartNamespace + "manualLayout");
        if (manualLayout is null)
        {
            return default;
        }

        double? x = ReadChartManualLayoutFactor(manualLayout, "x");
        double? y = ReadChartManualLayoutFactor(manualLayout, "y");
        double? width = ReadChartManualLayoutFactor(manualLayout, "w");
        double? height = ReadChartManualLayoutFactor(manualLayout, "h");
        return x is null || y is null || width is null || height is null
            ? default
            : new PptxSceneChartManualLayout(
                true,
                x.Value,
                y.Value,
                width.Value,
                height.Value,
                ReadChartElementValue(manualLayout, "layoutTarget"),
                ReadChartElementValue(manualLayout, "xMode"),
                ReadChartElementValue(manualLayout, "yMode"),
                ReadChartElementValue(manualLayout, "wMode"),
                ReadChartElementValue(manualLayout, "hMode"));
    }

    private static double? ReadChartManualLayoutFactor(XElement manualLayout, string elementName)
    {
        string? value = (string?)manualLayout.Element(ChartNamespace + elementName)?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? parsed
            : null;
    }

    private static PptxSceneLineStyle ReadChartGridlineLine(XElement? gridlines, PptxTheme theme)
    {
        XElement? shapeProperties = gridlines?.Element(ChartNamespace + "spPr");
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "noFill") is not null)
        {
            return new PptxSceneLineStyle(true, new RgbColor(0, 0, 0), 0d, 0d, [], null, null);
        }

        return ReadChartLine(shapeProperties, theme);
    }

    private static double? ReadChartAxisScalingValue(XElement axis, string elementName)
    {
        string? value = (string?)axis
            .Element(ChartNamespace + "scaling")
            ?.Element(ChartNamespace + elementName)
            ?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? parsed
            : null;
    }

    private static double? ReadChartAxisUnitValue(XElement axis, string elementName)
    {
        string? value = (string?)axis.Element(ChartNamespace + elementName)?.Attribute("val");
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) && parsed > 0d
            ? parsed
            : null;
    }

    private static bool IsChartGridlineVisible(XElement? gridlines)
    {
        XElement? line = gridlines?
            .Element(ChartNamespace + "spPr")
            ?.Element(DrawingNamespace + "ln");
        return gridlines is not null && line?.Element(DrawingNamespace + "noFill") is null;
    }

    private static string? ReadChartAxisNumberFormat(XElement axis)
    {
        string? format = (string?)axis
            .Element(ChartNamespace + "numFmt")
            ?.Attribute("formatCode");
        return string.IsNullOrWhiteSpace(format) ? null : format;
    }

    private static PptxSceneChartTitle ReadChartTitle(XDocument? chartXml, PptxTheme theme)
    {
        XElement? chart = chartXml?
            .Descendants(ChartNamespace + "chart")
            .FirstOrDefault();
        if (chart is null)
        {
            return new PptxSceneChartTitle(null, IsAutoDeleted: false, default);
        }

        bool isAutoDeleted = IsOoxmlBooleanElementEnabled(chart.Element(ChartNamespace + "autoTitleDeleted"));
        XElement? title = chart.Element(ChartNamespace + "title");
        string? text = title?
            .Descendants(DrawingNamespace + "t")
            .Aggregate(string.Empty, (current, textElement) => current + textElement.Value);
        if (string.IsNullOrWhiteSpace(text))
        {
            text = (string?)title?
                .Descendants(ChartNamespace + "v")
                .FirstOrDefault();
        }

        return new PptxSceneChartTitle(string.IsNullOrWhiteSpace(text) ? null : text, isAutoDeleted, ReadChartTextStyleOverride(title, theme));
    }

    private static PptxSceneChartLegend ReadChartLegend(XDocument? chartXml, PptxTheme theme)
    {
        XElement? legend = chartXml?
            .Descendants(ChartNamespace + "legend")
            .FirstOrDefault();
        if (legend is null)
        {
            return new PptxSceneChartLegend("r", Overlay: false, IsVisible: false, default);
        }

        return new PptxSceneChartLegend(
            (string?)legend.Element(ChartNamespace + "legendPos")?.Attribute("val") ?? "r",
            IsOoxmlBooleanElementEnabled(legend.Element(ChartNamespace + "overlay")),
            !IsOoxmlBooleanElementEnabled(legend.Element(ChartNamespace + "delete")),
            ReadChartTextStyleOverride(legend, theme));
    }

    private static IReadOnlyList<RgbColor>? ReadChartPaletteColors(OoxPackage package, string chartPartName, PptxTheme theme)
    {
        OoxRelationship? colorRelationship = package.GetRelationships(chartPartName)
            .FirstOrDefault(relationship => !relationship.IsExternal &&
                relationship.Type == ChartColorStyleRelationshipType &&
                relationship.ResolvedTarget is not null);
        if (colorRelationship?.ResolvedTarget is null)
        {
            return null;
        }

        OoxPart? colorPart = package.GetPart(colorRelationship.ResolvedTarget);
        if (colorPart is null)
        {
            return null;
        }

        XDocument document = LoadXml(colorPart);
        var colors = new List<RgbColor>();
        foreach (XElement colorElement in document.Root?.Elements().Where(element => element.Name.Namespace == DrawingNamespace) ?? [])
        {
            var wrapper = new XElement(DrawingNamespace + "solidFill", new XElement(colorElement));
            if (TryReadSolidColorWithAlpha(wrapper, theme, out RgbColor color, out _))
            {
                colors.Add(color);
            }
        }

        return colors.Count == 0 ? null : colors;
    }

    private static bool IsOoxmlBooleanElementEnabled(XElement? element)
    {
        return IsOoxmlBooleanElementEnabled(element, defaultValue: false);
    }

    private static bool IsOoxmlBooleanElementEnabled(XElement? element, bool defaultValue)
    {
        if (element is null)
        {
            return defaultValue;
        }

        string? value = (string?)element.Attribute("val");
        return value is null ||
            value == "1" ||
            value.Equals("true", StringComparison.OrdinalIgnoreCase);
    }

    private static bool? ReadOptionalOoxmlBooleanAttribute(XElement element, string attributeName)
    {
        XAttribute? attribute = element.Attribute(attributeName);
        if (attribute is null)
        {
            return null;
        }

        return attribute.Value == "1" ||
            attribute.Value.Equals("true", StringComparison.OrdinalIgnoreCase);
    }

    private static PptxSceneTable ReadTable(XElement frame, PptxTheme theme)
    {
        XElement? table = ReadTableElement(frame);
        IReadOnlyList<double> columnWidths = ReadTableColumnWidths(table);
        IReadOnlyList<double> rowHeights = ReadTableRowHeights(table);
        PptxSceneTableStyle style = ReadTableStyle(table);
        return new PptxSceneTable(
            columnWidths,
            rowHeights,
            ReadTableRows(table, theme, style, rowHeights.Count, columnWidths.Count),
            style,
            table);
    }

    internal static XElement? ReadTableElement(XElement frame)
    {
        return frame
            .Element(DrawingNamespace + "graphic")
            ?.Element(DrawingNamespace + "graphicData")
            ?.Element(DrawingNamespace + "tbl");
    }

    internal static IReadOnlyList<double> ReadTableColumnWidths(XElement? table)
    {
        return table
            ?.Element(DrawingNamespace + "tblGrid")
            ?.Elements(DrawingNamespace + "gridCol")
            .Select(column => Math.Max(1d, ReadLong(column, "w", 1)))
            .ToArray() ?? [];
    }

    internal static IReadOnlyList<double> ReadTableRowHeights(XElement? table)
    {
        return table
            ?.Elements(DrawingNamespace + "tr")
            .Select(row => Math.Max(1d, ReadLong(row, "h", 1)))
            .ToArray() ?? [];
    }

    internal static IReadOnlyList<PptxSceneTableRow> ReadTableRows(XElement? table, PptxTheme theme, PptxSceneTableStyle tableStyle, int rowCount, int columnCount)
    {
        if (table is null)
        {
            return [];
        }

        var rows = new List<PptxSceneTableRow>();
        int rowIndex = 0;
        foreach (XElement row in table.Elements(DrawingNamespace + "tr"))
        {
            var cells = new List<PptxSceneTableCell>();
            int columnIndex = 0;
            foreach (XElement cell in row.Elements(DrawingNamespace + "tc"))
            {
                PptxSceneTableCell sceneCell = ReadTableCell(cell, theme, tableStyle, rowIndex, columnIndex, rowCount, columnCount);
                cells.Add(sceneCell);
                columnIndex += sceneCell.IsMergedContinuation ? 1 : sceneCell.ColumnSpan;
            }

            rows.Add(new PptxSceneTableRow(cells));
            rowIndex++;
        }

        return rows;
    }

    internal static PptxSceneTableStyle ReadTableStyle(XElement? table)
    {
        XElement? tableProperties = table?.Element(DrawingNamespace + "tblPr");
        string? styleId = (string?)tableProperties?.Element(DrawingNamespace + "tableStyleId");
        bool supported = PptxBuiltInTableStyles.TryGet(styleId, out PptxBuiltInTableStyle style);
        return new PptxSceneTableStyle(
            styleId,
            supported ? style.Name : string.Empty,
            supported ? style.Accent : string.Empty,
            supported,
            ReadTablePropertyFlag(tableProperties, "firstRow"),
            ReadTablePropertyFlag(tableProperties, "lastRow"),
            ReadTablePropertyFlag(tableProperties, "firstCol"),
            ReadTablePropertyFlag(tableProperties, "lastCol"),
            ReadTablePropertyFlag(tableProperties, "bandRow"),
            ReadTablePropertyFlag(tableProperties, "bandCol"));
    }

    private static bool ReadTablePropertyFlag(XElement? tableProperties, string name)
    {
        if (tableProperties is null)
        {
            return false;
        }

        if (tableProperties.Attribute(name) is { } attribute)
        {
            return attribute.Value == "1" ||
                attribute.Value.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        return tableProperties.Element(DrawingNamespace + name) is not null;
    }

    internal static PptxSceneTableCell ReadTableCell(XElement cell, PptxTheme theme, PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount)
    {
        return new PptxSceneTableCell(
            ReadTableCellColumnSpan(cell),
            ReadTableCellRowSpan(cell),
            IsMergedTableCellContinuation(cell),
            ReadTableCellTextInsets(cell),
            ReadTableCellVerticalAnchor(cell),
            ReadTableCellFill(cell, theme),
            ReadTableCellBorders(cell, theme),
            PptxTableStyleResolver.ReadCellFill(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme),
            PptxTableStyleResolver.ReadCellTextStyle(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme),
            cell.Element(DrawingNamespace + "txBody"));
    }

    internal static bool IsMergedTableCellContinuation(XElement cell)
    {
        return ReadBool(cell, "hMerge") || ReadBool(cell, "vMerge");
    }

    internal static int ReadTableCellColumnSpan(XElement cell)
    {
        return ReadTableCellSpan(cell, "gridSpan");
    }

    internal static int ReadTableCellRowSpan(XElement cell)
    {
        return ReadTableCellSpan(cell, "rowSpan");
    }

    private static int ReadTableCellSpan(XElement cell, string attributeName)
    {
        return cell.Attribute(attributeName) is { } spanAttribute &&
            int.TryParse(spanAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            ? Math.Max(1, span)
            : 1;
    }

    internal static PptxSceneTextInsets ReadTableCellTextInsets(XElement cell)
    {
        XElement? textBody = cell.Element(DrawingNamespace + "txBody");
        XElement? bodyProperties = textBody?.Element(DrawingNamespace + "bodyPr");
        PptxSceneTextInsets bodyInsets = new(
            ReadInset(bodyProperties, "lIns", 91440),
            ReadInset(bodyProperties, "rIns", 91440),
            ReadInset(bodyProperties, "tIns", 45720),
            ReadInset(bodyProperties, "bIns", 45720));
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return new PptxSceneTextInsets(
            ReadTableCellMargin(cellProperties, "marL", bodyInsets.Left),
            ReadTableCellMargin(cellProperties, "marR", bodyInsets.Right),
            ReadTableCellMargin(cellProperties, "marT", bodyInsets.Top),
            ReadTableCellMargin(cellProperties, "marB", bodyInsets.Bottom));
    }

    private static double ReadTableCellMargin(XElement? cellProperties, string attributeName, double fallback)
    {
        return cellProperties?.Attribute(attributeName) is { } margin &&
            long.TryParse(margin.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emus)
                ? OoxUnits.EmuToPoints(emus)
                : fallback;
    }

    private static double ReadInset(XElement? element, string attributeName, long defaultEmu)
    {
        long emu = element?.Attribute(attributeName) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultEmu;
        return OoxUnits.EmuToPoints(emu);
    }

    internal static PptxSceneTableCellVerticalAnchor ReadTableCellVerticalAnchor(XElement cell)
    {
        string? anchor = (string?)cell
            .Element(DrawingNamespace + "tcPr")
            ?.Attribute("anchor");
        return anchor switch
        {
            "ctr" => PptxSceneTableCellVerticalAnchor.Middle,
            "b" => PptxSceneTableCellVerticalAnchor.Bottom,
            _ => PptxSceneTableCellVerticalAnchor.Top
        };
    }

    internal static PptxSceneFillStyle ReadTableCellFill(XElement cell, PptxTheme theme)
    {
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return TryReadSolidColorWithAlpha(cellProperties, theme, out RgbColor color, out double alpha)
            ? new PptxSceneFillStyle(true, color, alpha)
            : default;
    }

    internal static PptxSceneTableCellBorders ReadTableCellBorders(XElement cell, PptxTheme theme)
    {
        XElement? cellProperties = cell.Element(DrawingNamespace + "tcPr");
        return new PptxSceneTableCellBorders(
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnL"), theme),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnR"), theme),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnT"), theme),
            ReadTableCellBorder(cellProperties?.Element(DrawingNamespace + "lnB"), theme));
    }

    private static PptxSceneTableCellBorder ReadTableCellBorder(XElement? line, PptxTheme theme)
    {
        if (line is null)
        {
            return default;
        }

        if (line.Element(DrawingNamespace + "noFill") is not null ||
            !TryReadSolidColorWithAlpha(line, theme, out RgbColor color, out double alpha))
        {
            return new PptxSceneTableCellBorder(IsSpecified: true, default);
        }

        double lineWidth = line.Attribute("w") is { } widthAttribute
            ? Math.Max(1d, OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture)) / 2d)
            : 0.75d;
        return new PptxSceneTableCellBorder(
            IsSpecified: true,
            new PptxSceneLineStyle(true, color, lineWidth, alpha, [], null, null));
    }

    internal static PptxSceneGroupTransform ReadGroupTransform(XElement group)
    {
        XElement? transform = group
            .Element(PresentationNamespace + "grpSpPr")
            ?.Element(DrawingNamespace + "xfrm");
        XElement? offset = transform?.Element(DrawingNamespace + "off");
        XElement? extents = transform?.Element(DrawingNamespace + "ext");
        XElement? childOffset = transform?.Element(DrawingNamespace + "chOff");
        XElement? childExtents = transform?.Element(DrawingNamespace + "chExt");
        if (offset is null || extents is null || childOffset is null || childExtents is null)
        {
            return PptxSceneGroupTransform.Identity;
        }

        long width = ReadLong(extents, "cx");
        long height = ReadLong(extents, "cy");
        long childWidth = Math.Max(1, ReadLong(childExtents, "cx"));
        long childHeight = Math.Max(1, ReadLong(childExtents, "cy"));
        return new PptxSceneGroupTransform(
            ReadLong(offset, "x"),
            ReadLong(offset, "y"),
            width,
            height,
            ReadLong(childOffset, "x"),
            ReadLong(childOffset, "y"),
            width / (double)childWidth,
            height / (double)childHeight,
            transform!.Attribute("rot") is { } rotation ? long.Parse(rotation.Value, CultureInfo.InvariantCulture) / 60000d : 0d,
            ReadBool(transform, "flipH"),
            ReadBool(transform, "flipV"));
    }

    private static PptxSceneShape ReadShape(XElement shape, PptxTheme theme)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        PptxSceneLineStyle line = TryReadShapeLine(shape, shapeProperties, theme, out RgbColor lineColor, out double lineWidth, out double lineAlpha)
            ? new PptxSceneLineStyle(
                true,
                lineColor,
                lineWidth,
                lineAlpha,
                TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern) ? dashPattern : [],
                ReadLineCap(shapeProperties) switch
                {
                    "rnd" => 1,
                    "sq" => 2,
                    _ => null
                },
                ReadLineJoin(shapeProperties))
            : default;
        return new PptxSceneShape(
            ReadShapePreset(shapeProperties),
            ReadPresetAdjustments(shapeProperties),
            shapeProperties?.Element(DrawingNamespace + "custGeom") is not null,
            ReadCustomGeometry(shapeProperties),
            TryReadShapeFill(shape, shapeProperties, theme, out RgbColor fillColor, out double fillAlpha)
                ? new PptxSceneFillStyle(true, fillColor, fillAlpha)
                : default,
            TryReadShapeGradientFill(shapeProperties, theme, out PptxSceneGradientFill gradientFill) ? gradientFill : new PptxSceneGradientFill(false, 0d, []),
            TryReadShapePatternFill(shapeProperties, theme, out PptxScenePatternFill patternFill) ? patternFill : default,
            ReadShapePictureFill(shapeProperties),
            TryReadGlow(shapeProperties, theme, out PptxSceneGlow glow) ? glow : default,
            TryReadOuterShadow(shapeProperties, theme, out PptxSceneOuterShadow outerShadow) ? outerShadow : default,
            line,
            ReadLineEnd(shapeProperties, "headEnd"),
            ReadLineEnd(shapeProperties, "tailEnd"));
    }

    private static bool TryReadShapeGradientFill(XElement? shapeProperties, PptxTheme theme, out PptxSceneGradientFill fill)
    {
        XElement? gradientFill = shapeProperties?.Element(DrawingNamespace + "gradFill");
        XElement? gradientStopList = gradientFill?.Element(DrawingNamespace + "gsLst");
        if (gradientFill is null || gradientStopList is null)
        {
            fill = new PptxSceneGradientFill(false, 0d, []);
            return false;
        }

        PptxSceneGradientStop[] stops = gradientStopList
            .Elements(DrawingNamespace + "gs")
            .Select(stop => TryReadGradientStop(stop, theme, out PptxSceneGradientStop parsed) ? parsed : (PptxSceneGradientStop?)null)
            .Where(stop => stop is not null)
            .Select(stop => stop!.Value)
            .OrderBy(stop => stop.Offset)
            .ToArray();
        if (stops.Length < 2 ||
            stops.Any(stop => stop.Alpha < 0.999d) ||
            gradientFill.Element(DrawingNamespace + "lin") is not { } linear)
        {
            fill = new PptxSceneGradientFill(false, 0d, []);
            return false;
        }

        double angleDegrees = ReadLong(linear, "ang", 0) / 60000d;
        fill = new PptxSceneGradientFill(true, angleDegrees, stops);
        return true;
    }

    private static bool TryReadGradientStop(XElement gradientStop, PptxTheme theme, out PptxSceneGradientStop stop)
    {
        if (!TryReadSolidColorWithAlpha(gradientStop, theme, out RgbColor color, out double alpha))
        {
            stop = default;
            return false;
        }

        stop = new PptxSceneGradientStop(ParseGradientPercentage(gradientStop, "pos"), color, alpha);
        return true;
    }

    private static double ParseGradientPercentage(XElement element, string attribute)
    {
        return element.Attribute(attribute) is { } value
            ? Math.Clamp(int.Parse(value.Value, CultureInfo.InvariantCulture) / 100000d, 0d, 1d)
            : 0d;
    }

    private static IReadOnlyDictionary<string, double> ReadPresetAdjustments(XElement? shapeProperties)
    {
        Dictionary<string, double> adjustments = new(StringComparer.Ordinal);
        foreach (XElement guide in shapeProperties
                     ?.Element(DrawingNamespace + "prstGeom")
                     ?.Element(DrawingNamespace + "avLst")
                     ?.Elements(DrawingNamespace + "gd") ?? [])
        {
            string? name = (string?)guide.Attribute("name");
            string? formula = (string?)guide.Attribute("fmla");
            if (!string.IsNullOrWhiteSpace(name) &&
                formula is not null &&
                formula.StartsWith("val ", StringComparison.Ordinal) &&
                double.TryParse(formula[4..], NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            {
                adjustments[name] = value;
            }
        }

        return adjustments;
    }

    private static PptxSceneCustomGeometry ReadCustomGeometry(XElement? shapeProperties)
    {
        XElement? customGeometry = shapeProperties?.Element(DrawingNamespace + "custGeom");
        if (customGeometry is null)
        {
            return new PptxSceneCustomGeometry(false, [], []);
        }

        IReadOnlyList<PptxSceneCustomGuide> guides = customGeometry
            .Element(DrawingNamespace + "gdLst")
            ?.Elements(DrawingNamespace + "gd")
            .Select(guide => new PptxSceneCustomGuide(
                (string?)guide.Attribute("name") ?? string.Empty,
                (string?)guide.Attribute("fmla") ?? string.Empty))
            .Where(guide => !string.IsNullOrWhiteSpace(guide.Name) && !string.IsNullOrWhiteSpace(guide.Formula))
            .ToArray() ?? [];

        IReadOnlyList<PptxSceneCustomPath> paths = customGeometry
            .Element(DrawingNamespace + "pathLst")
            ?.Elements(DrawingNamespace + "path")
            .Select(ReadCustomPath)
            .Where(path => path.Commands.Count > 0)
            .ToArray() ?? [];

        return new PptxSceneCustomGeometry(paths.Count > 0, guides, paths);
    }

    private static PptxSceneCustomPath ReadCustomPath(XElement path)
    {
        PptxSceneCustomCommand?[] commands = path.Elements().Select(ReadCustomCommand).ToArray();
        if (commands.Any(command => command is null))
        {
            return new PptxSceneCustomPath(0d, 0d, false, false, []);
        }

        return new PptxSceneCustomPath(
            ParseOptionalDoubleAttribute(path, "w", 21600d),
            ParseOptionalDoubleAttribute(path, "h", 21600d),
            !string.Equals((string?)path.Attribute("fill"), "none", StringComparison.Ordinal),
            ParseBoolAttribute(path, "stroke", defaultValue: true),
            commands
                .Cast<PptxSceneCustomCommand>()
                .ToArray());
    }

    private static PptxSceneCustomCommand? ReadCustomCommand(XElement command)
    {
        return command.Name.LocalName switch
        {
            "moveTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.MoveTo,
                ReadCustomPoints(command).Take(1).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "lnTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.LineTo,
                ReadCustomPoints(command).Take(1).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "cubicBezTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.CubicBezierTo,
                ReadCustomPoints(command).Take(3).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "quadBezTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.QuadraticBezierTo,
                ReadCustomPoints(command).Take(2).ToArray(),
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            "arcTo" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.ArcTo,
                [],
                (string?)command.Attribute("wR") ?? string.Empty,
                (string?)command.Attribute("hR") ?? string.Empty,
                (string?)command.Attribute("stAng") ?? string.Empty,
                (string?)command.Attribute("swAng") ?? string.Empty),
            "close" => new PptxSceneCustomCommand(
                PptxSceneCustomCommandKind.Close,
                [],
                string.Empty,
                string.Empty,
                string.Empty,
                string.Empty),
            _ => null
        };
    }

    private static IReadOnlyList<PptxSceneCustomPoint> ReadCustomPoints(XElement command)
    {
        XElement[] points = command.Elements(DrawingNamespace + "pt").ToArray();
        if (points.Length == 0 && command.Name.LocalName is "moveTo" or "lnTo")
        {
            points = [command];
        }

        return points
            .Select(point => new PptxSceneCustomPoint(
                (string?)point.Attribute("x") ?? string.Empty,
                (string?)point.Attribute("y") ?? string.Empty))
            .ToArray();
    }

    private static PptxSceneShapePictureFill ReadShapePictureFill(XElement? shapeProperties)
    {
        XElement? blip = shapeProperties
            ?.Element(DrawingNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        string? relationshipId = (string?)blip?.Attribute(RelationshipsNamespace + "embed");
        return relationshipId is null
            ? default
            : new PptxSceneShapePictureFill(
                true,
                relationshipId,
                ReadPictureCrop(shapeProperties!),
                ReadPictureFill(shapeProperties!));
    }

    private static bool TryReadGlow(XElement? shapeProperties, PptxTheme theme, out PptxSceneGlow glow)
    {
        XElement? glowElement = shapeProperties
            ?.Element(DrawingNamespace + "effectLst")
            ?.Element(DrawingNamespace + "glow");
        if (glowElement is null)
        {
            glow = default;
            return false;
        }

        XElement? colorElement = glowElement.Elements().FirstOrDefault(element =>
            element.Name.LocalName is "srgbClr" or "schemeClr" or "prstClr");
        if (colorElement is not null &&
            TryReadImageRecolorColor(colorElement, theme, out RgbColor color))
        {
            double radius = OoxUnits.EmuToPoints(ReadLong(glowElement, "rad", 0));
            glow = new PptxSceneGlow(
                true,
                color,
                ReadAlpha(new XElement(DrawingNamespace + "solidFill", new XElement(colorElement))),
                radius);
            return radius > SceneEffectTolerance;
        }

        glow = default;
        return false;
    }

    private static bool TryReadOuterShadow(XElement? shapeProperties, PptxTheme theme, out PptxSceneOuterShadow shadow)
    {
        XElement? outerShadow = shapeProperties
            ?.Element(DrawingNamespace + "effectLst")
            ?.Element(DrawingNamespace + "outerShdw");
        if (outerShadow is null)
        {
            shadow = default;
            return false;
        }

        XElement? colorElement = outerShadow.Elements().FirstOrDefault(element =>
            element.Name.LocalName is "srgbClr" or "schemeClr" or "prstClr");
        if (colorElement is not null &&
            TryReadImageRecolorColor(colorElement, theme, out RgbColor color))
        {
            double alpha = ReadAlpha(new XElement(DrawingNamespace + "solidFill", new XElement(colorElement)));
            double distance = OoxUnits.EmuToPoints(ReadLong(outerShadow, "dist", 0));
            double direction = ReadLong(outerShadow, "dir", 0) / 60000d * Math.PI / 180d;
            shadow = new PptxSceneOuterShadow(
                true,
                color,
                alpha,
                distance * Math.Cos(direction),
                -distance * Math.Sin(direction));
            return true;
        }

        shadow = default;
        return false;
    }

    private static bool TryReadShapePatternFill(XElement? shapeProperties, PptxTheme theme, out PptxScenePatternFill fill)
    {
        XElement? patternFill = shapeProperties?.Element(DrawingNamespace + "pattFill");
        string? preset = (string?)patternFill?.Attribute("prst");
        if (patternFill is null || !IsSupportedDiagonalPatternFill(preset))
        {
            fill = default;
            return false;
        }

        RgbColor foreground = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "fgClr"), theme, out RgbColor foregroundColor, out _)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = TryReadSolidColorWithAlpha(patternFill.Element(DrawingNamespace + "bgClr"), theme, out RgbColor backgroundColor, out _)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        fill = new PptxScenePatternFill(true, preset!, foreground, background, 1d);
        return true;
    }

    private static bool IsSupportedDiagonalPatternFill(string? preset)
    {
        return preset is not null &&
            (preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase) ||
             preset.Contains("DnDiag", StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryReadShapeFill(XElement shape, XElement? shapeProperties, PptxTheme theme, out RgbColor color, out double alpha)
    {
        if (shapeProperties?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            alpha = 1d;
            return false;
        }

        if (TryReadSolidColorWithAlpha(shapeProperties, theme, out color, out alpha))
        {
            return true;
        }

        XElement? fillRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fillRef");
        int fillIndex = ParseOptionalIntAttribute(fillRef, "idx", 0);
        if (fillIndex > 0 &&
            theme.TryGetFillStyle(fillIndex, out XElement fillStyle) &&
            TryReadSolidColorWithAlpha(fillStyle, theme, fillRef, out color, out alpha))
        {
            return true;
        }

        return fillIndex > 0 && TryReadSolidColorWithAlpha(fillRef, theme, out color, out alpha);
    }

    private static bool TryReadShapeLine(XElement shape, XElement? shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth, out double alpha)
    {
        XElement? explicitLine = shapeProperties?.Element(DrawingNamespace + "ln");
        if (explicitLine?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        if (shapeProperties is not null && explicitLine is not null && TryReadLineWithAlpha(shapeProperties, theme, out color, out lineWidth, out alpha))
        {
            return true;
        }

        XElement? lineRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "lnRef");
        int lineIndex = ParseOptionalIntAttribute(lineRef, "idx", 0);
        if (lineIndex <= 0 || !theme.TryGetLineStyle(lineIndex, out XElement lineStyle))
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        lineWidth = lineStyle.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColorWithAlpha(lineStyle, theme, lineRef, out color, out alpha);
    }

    private static bool TryReadLineWithAlpha(XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth, out double alpha)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColorWithAlpha(line, theme, out color, out alpha);
    }

    private static bool TryReadPresetDash(XElement? shapeProperties, double lineWidth, out IReadOnlyList<double> dashPattern)
    {
        string? presetDash = (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + "prstDash")
            ?.Attribute("val");
        double w = Math.Max(lineWidth, MinimumStrokeWidth);
        dashPattern = presetDash switch
        {
            "dot" or "sysDot" => [w, w * 2d],
            "dash" or "sysDash" => [w * 4d, w * 3d],
            "lgDash" => [w * 8d, w * 3d],
            "dashDot" or "sysDashDot" => [w * 4d, w * 3d, w, w * 3d],
            "lgDashDot" => [w * 8d, w * 3d, w, w * 3d],
            "lgDashDotDot" or "sysDashDotDot" => [w * 8d, w * 3d, w, w * 3d, w, w * 3d],
            _ => []
        };
        return dashPattern.Count > 0;
    }

    private static string? ReadLineCap(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Attribute("cap");
    }

    private static int? ReadLineJoin(XElement? shapeProperties)
    {
        XElement? line = shapeProperties?.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return 1;
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return 2;
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return 0;
        }

        return null;
    }

    private static string ReadShapePreset(XElement? shapeProperties)
    {
        return (string?)shapeProperties
            ?.Element(DrawingNamespace + "prstGeom")
            ?.Attribute("prst") ?? "rect";
    }

    private static PptxSceneLineEnd ReadLineEnd(XElement? shapeProperties, string elementName)
    {
        XElement? end = shapeProperties
            ?.Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + elementName);
        return new PptxSceneLineEnd(
            ReadLineEndKind((string?)end?.Attribute("type")),
            ReadLineEndScale((string?)end?.Attribute("w")),
            ReadLineEndScale((string?)end?.Attribute("len")));
    }

    private static PptxSceneLineEndKind ReadLineEndKind(string? type)
    {
        return type switch
        {
            "triangle" => PptxSceneLineEndKind.Triangle,
            "arrow" => PptxSceneLineEndKind.Arrow,
            "stealth" => PptxSceneLineEndKind.Stealth,
            "diamond" => PptxSceneLineEndKind.Diamond,
            "oval" => PptxSceneLineEndKind.Oval,
            _ => PptxSceneLineEndKind.None
        };
    }

    private static double ReadLineEndScale(string? value)
    {
        return value switch
        {
            "sm" => 0.5d,
            "lg" => 1.5d,
            _ => 1d
        };
    }

    internal static PptxSceneRect ReadPictureCrop(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? sourceRectangle = blipFill?.Element(DrawingNamespace + "srcRect");
        return sourceRectangle is null
            ? default
            : new PptxSceneRect(
                ParsePercentage(sourceRectangle, "l"),
                ParsePercentage(sourceRectangle, "t"),
                ParsePercentage(sourceRectangle, "r"),
                ParsePercentage(sourceRectangle, "b"));
    }

    internal static PptxSceneRect ReadPictureFill(XElement picture)
    {
        XElement? blipFill = picture.Element(PresentationNamespace + "blipFill") ??
            picture.Element(DrawingNamespace + "blipFill");
        XElement? fillRectangle = blipFill
            ?.Element(DrawingNamespace + "stretch")
            ?.Element(DrawingNamespace + "fillRect");
        return fillRectangle is null
            ? default
            : new PptxSceneRect(
                ParsePercentage(fillRectangle, "l"),
                ParsePercentage(fillRectangle, "t"),
                ParsePercentage(fillRectangle, "r"),
                ParsePercentage(fillRectangle, "b"));
    }

    internal static double ReadPictureAlpha(XElement picture)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        XElement? alphaModFix = blip?.Element(DrawingNamespace + "alphaModFix");
        if (alphaModFix?.Attribute("amt") is { } amount &&
            int.TryParse(amount.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedAmount))
        {
            return Math.Clamp(parsedAmount / 100000d, 0d, 1d);
        }

        return 1d;
    }

    internal static PptxSceneImageRecolor ReadImageRecolor(XElement picture, PptxTheme theme)
    {
        XElement? blip = picture
            .Element(PresentationNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip");
        if (blip is null)
        {
            return PptxSceneImageRecolor.None;
        }

        if (blip.Element(DrawingNamespace + "grayscl") is not null)
        {
            return PptxSceneImageRecolor.Grayscale();
        }

        XElement? biLevel = blip.Element(DrawingNamespace + "biLevel");
        if (biLevel is not null)
        {
            double threshold = biLevel.Attribute("thresh") is { } thresholdAttribute
                ? Math.Clamp(int.Parse(thresholdAttribute.Value, CultureInfo.InvariantCulture) / 100000d, 0d, 1d)
                : 0.5d;
            return PptxSceneImageRecolor.BiLevel(threshold);
        }

        XElement? luminance = blip.Element(DrawingNamespace + "lum");
        if (luminance is not null)
        {
            double brightness = luminance.Attribute("bright") is { } brightnessAttribute
                ? Math.Clamp(int.Parse(brightnessAttribute.Value, CultureInfo.InvariantCulture) / 100000d, -1d, 1d)
                : 0d;
            double contrast = luminance.Attribute("contrast") is { } contrastAttribute
                ? Math.Clamp(int.Parse(contrastAttribute.Value, CultureInfo.InvariantCulture) / 100000d, -1d, 1d)
                : 0d;
            return PptxSceneImageRecolor.Luminance(brightness, contrast);
        }

        XElement? duotone = blip.Element(DrawingNamespace + "duotone");
        if (duotone is not null)
        {
            XElement[] colors = duotone.Elements().Take(2).ToArray();
            if (colors.Length == 2 &&
                TryReadImageRecolorColor(colors[0], theme, out RgbColor dark) &&
                TryReadImageRecolorColor(colors[1], theme, out RgbColor light))
            {
                return PptxSceneImageRecolor.Duotone(dark, light);
            }
        }

        return PptxSceneImageRecolor.None;
    }

    private static bool TryReadImageRecolorColor(XElement colorElement, PptxTheme theme, out RgbColor color)
    {
        if (colorElement.Name == DrawingNamespace + "prstClr")
        {
            string? preset = (string?)colorElement.Attribute("val");
            color = preset switch
            {
                "black" => new RgbColor(0, 0, 0),
                "white" => new RgbColor(255, 255, 255),
                _ => default
            };
            return preset is "black" or "white";
        }

        XElement wrapper = new(DrawingNamespace + "solidFill", new XElement(colorElement));
        return TryReadSolidColorWithAlpha(wrapper, theme, out color, out _);
    }

    private static double ParsePercentage(XElement element, string attribute)
    {
        return element.Attribute(attribute) is { } value
            ? Math.Clamp(int.Parse(value.Value, CultureInfo.InvariantCulture) / 100000d, 0d, 0.999d)
            : 0d;
    }

    private static PptxSceneTextBody? ReadTextBody(XElement element, IReadOnlyList<XDocument> placeholderSources, PptxTheme theme)
    {
        XElement? textBody = element.Element(PresentationNamespace + "txBody");
        if (textBody is null)
        {
            return null;
        }

        IReadOnlyList<XElement> inheritedTextBodies = FindInheritedPlaceholderShapes(element, placeholderSources)
            .Select(shape => shape.Element(PresentationNamespace + "txBody"))
            .Where(textBody => textBody is not null)
            .Cast<XElement>()
            .ToArray();
        return new PptxSceneTextBody(
            textBody.Element(DrawingNamespace + "bodyPr"),
            textBody.Element(DrawingNamespace + "lstStyle"),
            textBody.Elements(DrawingNamespace + "p").Select(paragraph => ReadParagraph(paragraph, element, textBody, inheritedTextBodies, placeholderSources, theme)).ToArray());
    }

    private static PptxSceneTextParagraph ReadParagraph(
        XElement paragraph,
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedTextBodies,
        IReadOnlyList<XDocument> placeholderSources,
        PptxTheme theme)
    {
        XElement? properties = paragraph.Element(DrawingNamespace + "pPr");
        int level = properties?.Attribute("lvl") is { } levelAttribute
            ? int.Parse(levelAttribute.Value, CultureInfo.InvariantCulture)
            : 0;
        XElement? defaultParagraphProperties = ResolveDefaultParagraphProperties(
            level,
            shape,
            textBody,
            inheritedTextBodies,
            placeholderSources);
        XElement? defaultRunProperties = properties?.Element(DrawingNamespace + "defRPr") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
        PptxSceneParagraphStyle resolvedStyle = ResolveParagraphStyle(level, properties, defaultParagraphProperties, defaultRunProperties, shape, theme);
        return new PptxSceneTextParagraph(
            properties,
            paragraph.Element(DrawingNamespace + "endParaRPr"),
            level,
            resolvedStyle,
            paragraph.Elements().Select(run => ReadRun(run, defaultRunProperties, resolvedStyle, theme)).Where(run => run is not null).Cast<PptxSceneTextRun>().ToArray());
    }

    private static PptxSceneTextRun? ReadRun(XElement element, XElement? defaultRunProperties, PptxSceneParagraphStyle paragraphStyle, PptxTheme theme)
    {
        if (element.Name == DrawingNamespace + "r")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(
                PptxSceneTextRunKind.Text,
                (string?)element.Element(DrawingNamespace + "t") ?? string.Empty,
                runProperties,
                ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme),
                element);
        }

        if (element.Name == DrawingNamespace + "br")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(PptxSceneTextRunKind.Break, "\n", runProperties, ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme), element);
        }

        if (element.Name == DrawingNamespace + "fld")
        {
            XElement? runProperties = element.Element(DrawingNamespace + "rPr");
            return new PptxSceneTextRun(
                PptxSceneTextRunKind.Field,
                (string?)element.Element(DrawingNamespace + "t") ?? string.Empty,
                runProperties,
                ResolveRunStyle(runProperties, defaultRunProperties, paragraphStyle, theme),
                element);
        }

        return null;
    }

    private static XElement? ResolveDefaultParagraphProperties(
        int level,
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedTextBodies,
        IReadOnlyList<XDocument> placeholderSources)
    {
        string levelName = $"lvl{Math.Clamp(level + 1, 1, 9).ToString(CultureInfo.InvariantCulture)}pPr";
        var sources = new List<XElement?>();
        sources.Add(textBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName));
        sources.AddRange(inheritedTextBodies
            .Reverse()
            .Select(inheritedTextBody => inheritedTextBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)));
        sources.Add(FindInheritedTextStyle(shape, placeholderSources, levelName));
        sources.Add(FindDefaultTextStyle(placeholderSources, levelName));
        return MergeParagraphProperties(sources.ToArray());
    }

    private static XElement? MergeParagraphProperties(params XElement?[] sources)
    {
        XElement? merged = null;
        foreach (XElement source in sources.Reverse().Where(source => source is not null).Cast<XElement>())
        {
            merged ??= new XElement(source.Name);
            foreach (XAttribute attribute in source.Attributes())
            {
                merged.SetAttributeValue(attribute.Name, attribute.Value);
            }

            foreach (XElement child in source.Elements())
            {
                XElement? existing = merged.Element(child.Name);
                if (existing is null)
                {
                    merged.Add(new XElement(child));
                }
                else
                {
                    MergeElementInto(existing, child);
                }
            }
        }

        return merged;
    }

    private static void MergeElementInto(XElement target, XElement source)
    {
        foreach (XAttribute attribute in source.Attributes())
        {
            target.SetAttributeValue(attribute.Name, attribute.Value);
        }

        foreach (XElement child in source.Elements())
        {
            target.Elements(child.Name).Remove();
            target.Add(new XElement(child));
        }
    }

    private static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        string? placeholderType = (string?)shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph")
            ?.Attribute("type");
        string styleName = placeholderType switch
        {
            "title" or "ctrTitle" => "titleStyle",
            "body" or "subTitle" => "bodyStyle",
            _ => "otherStyle"
        };

        foreach (XDocument source in placeholderSources)
        {
            XElement? style = source.Root?
                .Element(PresentationNamespace + "txStyles")
                ?.Element(PresentationNamespace + styleName);
            XElement? level = style?.Element(DrawingNamespace + levelName) ??
                style?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    private static XElement? FindDefaultTextStyle(IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        foreach (XDocument source in placeholderSources)
        {
            XElement? defaultTextStyle = source.Root?.Element(PresentationNamespace + "defaultTextStyle");
            XElement? level = defaultTextStyle?.Element(DrawingNamespace + levelName) ??
                defaultTextStyle?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    private static IReadOnlyList<XElement> FindInheritedPlaceholderShapes(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        XElement? placeholder = shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph");
        if (placeholder is null)
        {
            return [];
        }

        var matches = new List<XElement>();
        string? type = (string?)placeholder.Attribute("type");
        string? index = (string?)placeholder.Attribute("idx");
        foreach (XDocument source in placeholderSources)
        {
            foreach (XElement candidate in source.Descendants(PresentationNamespace + "sp"))
            {
                XElement? candidatePlaceholder = candidate
                    .Element(PresentationNamespace + "nvSpPr")
                    ?.Element(PresentationNamespace + "nvPr")
                    ?.Element(PresentationNamespace + "ph");
                if (candidatePlaceholder is null)
                {
                    continue;
                }

                string? candidateType = (string?)candidatePlaceholder.Attribute("type");
                string? candidateIndex = (string?)candidatePlaceholder.Attribute("idx");
                bool indexMatches = index is not null && candidateIndex == index;
                bool typeMatches = index is null && type is not null && candidateType == type;
                if (indexMatches || typeMatches)
                {
                    matches.Add(candidate);
                    break;
                }
            }
        }

        return matches;
    }

    private static PptxSceneParagraphStyle ResolveParagraphStyle(
        int level,
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        XElement? defaultRunProperties,
        XElement shape,
        PptxTheme theme)
    {
        double fontSize = ReadFontSize(defaultRunProperties, null);
        RgbColor color = TryReadSolidColorWithAlpha(defaultRunProperties, theme, out RgbColor defaultColor, out double alpha)
            ? defaultColor
            : TryReadShapeFontColor(shape, theme, out RgbColor shapeColor)
                ? shapeColor
                : new RgbColor(0, 0, 0);
        if (!TryReadSolidColorWithAlpha(defaultRunProperties, theme, out _, out alpha))
        {
            alpha = 1d;
        }

        string? typeface = theme.ResolveTypeface((string?)defaultRunProperties?.Element(DrawingNamespace + "latin")?.Attribute("typeface"));
        return new PptxSceneParagraphStyle(
            level,
            (string?)(paragraphProperties?.Attribute("algn") ?? defaultParagraphProperties?.Attribute("algn")) ?? "l",
            fontSize,
            color,
            alpha,
            typeface,
            ParseOptionalBoolAttribute(defaultRunProperties, "b"),
            ParseOptionalBoolAttribute(defaultRunProperties, "i"),
            ReadCharacterSpacing(defaultRunProperties, null));
    }

    private static PptxSceneRunStyle ResolveRunStyle(XElement? runProperties, XElement? defaultRunProperties, PptxSceneParagraphStyle paragraphStyle, PptxTheme theme)
    {
        double fontSize = ReadFontSize(runProperties, defaultRunProperties);
        double alpha = paragraphStyle.Alpha;
        RgbColor color = paragraphStyle.Color;
        if (TryReadSolidColorWithAlpha(runProperties, theme, out RgbColor runColor, out double runAlpha))
        {
            color = runColor;
            alpha = runAlpha;
        }
        else if (TryReadSolidColorWithAlpha(defaultRunProperties, theme, out RgbColor defaultColor, out double defaultAlpha))
        {
            color = defaultColor;
            alpha = defaultAlpha;
        }

        string? typeface = theme.ResolveTypeface((string?)(runProperties?.Element(DrawingNamespace + "latin") ??
            defaultRunProperties?.Element(DrawingNamespace + "latin"))?.Attribute("typeface")) ?? paragraphStyle.Typeface;
        bool bold = ParseOptionalBoolAttribute(runProperties, "b") ||
            (runProperties?.Attribute("b") is null && paragraphStyle.Bold);
        bool italic = ParseOptionalBoolAttribute(runProperties, "i") ||
            (runProperties?.Attribute("i") is null && paragraphStyle.Italic);
        bool underline = ((string?)(runProperties?.Attribute("u") ?? defaultRunProperties?.Attribute("u"))) is { } underlineValue &&
            !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
        bool strike = IsStrikeEnabled(runProperties, defaultRunProperties);
        return new PptxSceneRunStyle(
            fontSize,
            color,
            alpha,
            typeface,
            bold,
            italic,
            underline,
            strike,
            ReadCharacterSpacing(runProperties, defaultRunProperties),
            ReadBaselineOffset(runProperties, defaultRunProperties, fontSize),
            TryReadHighlightColor(runProperties, out RgbColor highlight) ? highlight : null);
    }

    private static double ReadFontSize(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("sz") ?? defaultRunProperties?.Attribute("sz")) is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
            : 18d;
    }

    private static double ReadCharacterSpacing(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("spc") ?? defaultRunProperties?.Attribute("spc")) is { } spacing
            ? int.Parse(spacing.Value, CultureInfo.InvariantCulture) / 100d
            : 0d;
    }

    private static double ReadBaselineOffset(XElement? runProperties, XElement? defaultRunProperties, double fontSize)
    {
        return (runProperties?.Attribute("baseline") ?? defaultRunProperties?.Attribute("baseline")) is { } baseline
            ? fontSize * int.Parse(baseline.Value, CultureInfo.InvariantCulture) / 100000d
            : 0d;
    }

    private static bool IsStrikeEnabled(XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = (string?)(runProperties?.Attribute("strike") ?? defaultRunProperties?.Attribute("strike"));
        return value is not null && !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ParseOptionalBoolAttribute(XElement? element, string attributeName)
    {
        return ParseBoolAttribute(element, attributeName, defaultValue: false);
    }

    private static bool ParseBoolAttribute(XElement? element, string attributeName, bool defaultValue)
    {
        string? value = (string?)element?.Attribute(attributeName);
        return value is null
            ? defaultValue
            : value is "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }

    private static int ParseOptionalIntAttribute(XElement? element, string attributeName, int defaultValue)
    {
        return element?.Attribute(attributeName) is { } attribute &&
            int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)
            ? value
            : defaultValue;
    }

    private static double ParseOptionalDoubleAttribute(XElement element, string attributeName, double defaultValue)
    {
        return element.Attribute(attributeName) is { } attribute &&
            double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
            ? value
            : defaultValue;
    }

    private static bool TryReadHighlightColor(XElement? runProperties, out RgbColor color)
    {
        XElement? highlight = runProperties?.Element(DrawingNamespace + "highlight");
        string? hex = (string?)highlight?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        return RgbColor.TryParse(hex, out color);
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, out RgbColor color)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return TryReadSolidColorWithAlpha(fontRef, theme, out color, out _);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, out RgbColor color, out double alpha)
    {
        return TryReadSolidColorWithAlpha(element, theme, placeholderColorContainer: null, out color, out alpha);
    }

    private static bool TryReadSolidColorWithAlpha(XElement? element, PptxTheme theme, XElement? placeholderColorContainer, out RgbColor color, out double alpha)
    {
        XElement? solidFill = element?.Element(DrawingNamespace + "solidFill");
        XElement? colorContainer = element?.Name == DrawingNamespace + "solidFill"
            ? element
            : solidFill ?? element;
        alpha = ReadAlpha(colorContainer);
        XElement? srgbColor = colorContainer?.Element(DrawingNamespace + "srgbClr");
        if (RgbColor.TryParse((string?)srgbColor?.Attribute("val"), out color))
        {
            color = ApplyColorTransforms(srgbColor, color);
            return true;
        }

        XElement? schemeColorElement = colorContainer?.Element(DrawingNamespace + "schemeClr");
        string? schemeColor = (string?)schemeColorElement?.Attribute("val");
        if (schemeColor == "phClr" &&
            placeholderColorContainer is not null &&
            TryReadSolidColorWithAlpha(placeholderColorContainer, theme, placeholderColorContainer: null, out color, out double placeholderAlpha))
        {
            color = ApplyColorTransforms(schemeColorElement, color);
            alpha *= placeholderAlpha;
            return true;
        }

        if (schemeColor is not null && theme.TryResolveColor(schemeColor, out color))
        {
            color = ApplyColorTransforms(schemeColorElement, color);
            return true;
        }

        XElement? systemColor = colorContainer?.Element(DrawingNamespace + "sysClr");
        string? systemHex = (string?)systemColor?.Attribute("lastClr") ?? (string?)systemColor?.Attribute("val");
        if (RgbColor.TryParse(systemHex, out color))
        {
            color = ApplyColorTransforms(systemColor, color);
            return true;
        }

        XElement? presetColor = colorContainer?.Element(DrawingNamespace + "prstClr");
        if (TryResolvePresetColor((string?)presetColor?.Attribute("val"), out color))
        {
            color = ApplyColorTransforms(presetColor, color);
            return true;
        }

        XElement? scRgbColor = colorContainer?.Element(DrawingNamespace + "scrgbClr");
        if (scRgbColor is not null)
        {
            color = new RgbColor(
                ReadPercentageByte(scRgbColor, "r"),
                ReadPercentageByte(scRgbColor, "g"),
                ReadPercentageByte(scRgbColor, "b"));
            color = ApplyColorTransforms(scRgbColor, color);
            return true;
        }

        XElement? hslColor = colorContainer?.Element(DrawingNamespace + "hslClr");
        if (hslColor is not null)
        {
            color = ReadHslColor(hslColor);
            color = ApplyColorTransforms(hslColor, color);
            return true;
        }

        return false;
    }

    private static double ReadAlpha(XElement? colorContainer)
    {
        XElement? alpha = colorContainer?
            .Elements()
            .FirstOrDefault(element => element.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr" or "scrgbClr" or "hslClr")
            ?.Element(DrawingNamespace + "alpha");
        if (alpha?.Attribute("val") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed))
        {
            return Math.Clamp(parsed / 100000d, 0d, 1d);
        }

        return 1d;
    }

    private static RgbColor ApplyColorTransforms(XElement? colorElement, RgbColor color)
    {
        if (colorElement is null)
        {
            return color;
        }

        double red = color.Red;
        double green = color.Green;
        double blue = color.Blue;
        foreach (XElement transform in colorElement.Elements())
        {
            double value = ReadLong(transform, "val", 100000) / 100000d;
            switch (transform.Name.LocalName)
            {
                case "lumMod":
                case "shade":
                    red *= value;
                    green *= value;
                    blue *= value;
                    break;
                case "lumOff":
                    red += 255d * value;
                    green += 255d * value;
                    blue += 255d * value;
                    break;
                case "tint":
                    red += (255d - red) * value;
                    green += (255d - green) * value;
                    blue += (255d - blue) * value;
                    break;
            }
        }

        return new RgbColor(ToByte(red), ToByte(green), ToByte(blue));
    }

    private static bool TryResolvePresetColor(string? name, out RgbColor color)
    {
        return PptxPresetColors.TryResolve(name, out color);
    }

    private static byte ReadPercentageByte(XElement element, string attributeName)
    {
        double ratio = ReadLong(element, attributeName, 0) / 100000d;
        return ToByte(255d * ratio);
    }

    private static RgbColor ReadHslColor(XElement element)
    {
        double hue = (ReadLong(element, "hue", 0) / 60000d) % 360d;
        double saturation = Math.Clamp(ReadLong(element, "sat", 0) / 100000d, 0d, 1d);
        double luminosity = Math.Clamp(ReadLong(element, "lum", 0) / 100000d, 0d, 1d);
        double chroma = (1d - Math.Abs(2d * luminosity - 1d)) * saturation;
        double segment = hue / 60d;
        double second = chroma * (1d - Math.Abs(segment % 2d - 1d));
        (double r1, double g1, double b1) = segment switch
        {
            >= 0d and < 1d => (chroma, second, 0d),
            >= 1d and < 2d => (second, chroma, 0d),
            >= 2d and < 3d => (0d, chroma, second),
            >= 3d and < 4d => (0d, second, chroma),
            >= 4d and < 5d => (second, 0d, chroma),
            _ => (chroma, 0d, second)
        };
        double match = luminosity - chroma / 2d;
        return new RgbColor(ToByte((r1 + match) * 255d), ToByte((g1 + match) * 255d), ToByte((b1 + match) * 255d));
    }

    private static byte ToByte(double value)
    {
        return (byte)Math.Clamp((int)Math.Round(value, MidpointRounding.AwayFromZero), 0, 255);
    }

    private static long ReadLong(XElement element, string name)
    {
        return element.Attribute(name) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : 0L;
    }

    private static long ReadLong(XElement element, string name, long defaultValue)
    {
        return element.Attribute(name) is { } attribute
            ? long.Parse(attribute.Value, CultureInfo.InvariantCulture)
            : defaultValue;
    }

    private static bool ReadBool(XElement element, string name)
    {
        string? value = (string?)element.Attribute(name);
        return value is "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
    }
}
