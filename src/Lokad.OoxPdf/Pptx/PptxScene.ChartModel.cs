using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxSceneChart(
    string? RelationshipId,
    string? TargetPartName,
    XDocument? ChartXml,
    PptxColorMap ColorMap,
    PptxSceneChartExternalData ExternalData,
    PptxSceneChartOptions Options,
    IReadOnlyList<RgbColor>? PaletteColors,
    PptxSceneChartColorStyle ColorStyle,
    PptxSceneChartStyle StylePart,
    string StyleId,
    IReadOnlyList<PptxSceneChartPlot> Plots,
    IReadOnlyList<PptxSceneChartAxis> Axes,
    PptxSceneChartTitle Title,
    PptxSceneChartLegend Legend,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartManualLayout PlotAreaLayout,
    PptxSceneChartShapeStyle ChartAreaStyle,
    PptxSceneChartShapeStyle PlotAreaStyle);

internal readonly record struct PptxSceneChartExternalData(
    bool IsDefined,
    string? RelationshipId,
    string? TargetPartName,
    PptxScenePackageResource? Resource,
    bool? AutoUpdate,
    string AutoUpdateValue);

internal readonly record struct PptxSceneChartOptions(
    bool? Date1904,
    string Date1904Value,
    bool? RoundedCorners,
    string RoundedCornersValue,
    bool? PlotVisibleOnly,
    string PlotVisibleOnlyValue,
    bool? ShowDataLabelsOverMaximum,
    string ShowDataLabelsOverMaximumValue,
    PptxSceneChartDisplayBlanksAs DisplayBlanksAsKind,
    string DisplayBlanksAs);

internal enum PptxSceneChartDisplayBlanksAs
{
    Unknown,
    Gap,
    Span,
    Zero
}

internal readonly record struct PptxSceneChartColorStyle(
    bool IsDefined,
    string? PartName,
    string Method,
    string Id,
    IReadOnlyList<RgbColor> Colors,
    int VariationCount,
    IReadOnlyList<PptxSceneChartColorDeclaration> Declarations,
    IReadOnlyList<PptxSceneChartColorDeclaration> RootDeclarations,
    IReadOnlyList<PptxSceneChartColorVariation> Variations,
    XDocument? ColorStyleXml);

internal readonly record struct PptxSceneChartColorVariation(
    int Index,
    IReadOnlyList<PptxSceneChartColorDeclaration> Declarations,
    IReadOnlyList<RgbColor> Colors);

internal readonly record struct PptxSceneChartColorDeclaration(
    string Kind,
    string Value,
    int? VariationIndex,
    bool IsResolved,
    RgbColor? Color,
    double Alpha);

internal sealed record PptxSceneChartStyle(
    bool IsDefined,
    string? PartName,
    string Id,
    XDocument? StyleXml,
    IReadOnlyList<PptxSceneChartStyleEntry> Entries);

internal readonly record struct PptxSceneChartStyleEntry(
    string Role,
    int SourceIndex,
    string NamespaceUri,
    int? LineReferenceIndex,
    string LineReferenceIndexValue,
    int? FillReferenceIndex,
    string FillReferenceIndexValue,
    PptxSceneFillStyle FillReferenceFill,
    int? EffectReferenceIndex,
    string EffectReferenceIndexValue,
    PptxSceneChartEffectFamily EffectReferenceEffects,
    string FontReferenceIndex,
    PptxSceneLineStyle Line,
    PptxSceneChartShapeStyle ShapeStyle,
    PptxSceneLineStyle ShapeLine,
    PptxSceneChartTextStyleOverride TextStyle);

internal sealed record PptxSceneChartPlot(
    PptxSceneChartPlotKind PlotKind,
    string Kind,
    int PlotAreaIndex,
    int KindIndex,
    int SeriesCount,
    IReadOnlyList<string> AxisIds,
    IReadOnlyList<PptxSceneChartSeries> Series,
    PptxSceneChartGrouping GroupingKind,
    string Grouping,
    PptxSceneChartBarDirection BarDirectionKind,
    string BarDirection,
    PptxSceneChartScatterStyle ScatterStyleKind,
    string ScatterStyle,
    PptxSceneChartRadarStyle RadarStyleKind,
    string RadarStyle,
    bool? MarkersEnabled,
    string MarkersEnabledValue,
    bool? VaryColors,
    string VaryColorsValue,
    double? GapWidth,
    string GapWidthValue,
    double? Overlap,
    string OverlapValue,
    double? HoleSize,
    string HoleSizeValue,
    double? FirstSliceAngle,
    string FirstSliceAngleValue,
    PptxSceneChartDataLabels DataLabels,
    XElement Source);

internal enum PptxSceneChartGrouping
{
    Clustered,
    PercentStacked,
    Stacked,
    Standard,
    Unknown
}

internal enum PptxSceneChartBarDirection
{
    Bar,
    Column,
    Unknown
}

internal enum PptxSceneChartScatterStyle
{
    Line,
    LineMarker,
    Marker,
    None,
    Smooth,
    SmoothMarker,
    Unknown
}

internal enum PptxSceneChartRadarStyle
{
    Filled,
    Marker,
    Standard,
    Unknown
}

internal enum PptxSceneChartPlotKind
{
    Area,
    Bar,
    Bubble,
    Doughnut,
    Line,
    Pie,
    Radar,
    Scatter,
    Unknown
}

internal sealed record PptxSceneChartDataLabels(
    bool? ShowValue,
    string ShowValueValue,
    bool? ShowPercent,
    string ShowPercentValue,
    bool? ShowCategoryName,
    string ShowCategoryNameValue,
    bool? ShowSeriesName,
    string ShowSeriesNameValue,
    bool? ShowLeaderLines,
    string ShowLeaderLinesValue,
    bool? ShowLegendKey,
    string ShowLegendKeyValue,
    bool? ShowBubbleSize,
    string ShowBubbleSizeValue,
    PptxSceneChartLeaderLines LeaderLines,
    PptxSceneChartDataLabelPosition PositionKind,
    string Position,
    string Separator,
    string NumberFormat,
    PptxSceneChartNumberFormat NumberFormatInfo,
    PptxSceneChartManualLayout Layout,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartTextBodyProperties TextBodyProperties,
    PptxSceneChartShapeStyle ShapeStyle,
    IReadOnlyList<string> RejectedOverrideIndexValues,
    IReadOnlyList<PptxSceneChartDataLabelOverride> Overrides,
    bool IsDefined);

internal sealed record PptxSceneChartDataLabelOverride(
    int Index,
    string IndexValue,
    bool? IsDeleted,
    string IsDeletedValue,
    bool? ShowValue,
    string ShowValueValue,
    bool? ShowPercent,
    string ShowPercentValue,
    bool? ShowCategoryName,
    string ShowCategoryNameValue,
    bool? ShowSeriesName,
    string ShowSeriesNameValue,
    bool? ShowLeaderLines,
    string ShowLeaderLinesValue,
    bool? ShowLegendKey,
    string ShowLegendKeyValue,
    bool? ShowBubbleSize,
    string ShowBubbleSizeValue,
    PptxSceneChartLeaderLines LeaderLines,
    string CustomText,
    IReadOnlyList<PptxSceneChartTextRun> CustomTextRuns,
    PptxSceneChartDataLabelPosition PositionKind,
    string Position,
    string Separator,
    string NumberFormat,
    PptxSceneChartNumberFormat NumberFormatInfo,
    PptxSceneChartManualLayout Layout,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartTextBodyProperties TextBodyProperties,
    PptxSceneChartShapeStyle ShapeStyle);

internal readonly record struct PptxSceneChartTextBodyProperties(
    double? RotationDegrees,
    string RotationValue,
    string OrientationValue,
    string VerticalOverflowValue);

internal readonly record struct PptxSceneChartNumberFormat(
    bool IsDefined,
    string FormatCode,
    bool? SourceLinked,
    string SourceLinkedValue);

internal readonly record struct PptxSceneChartLeaderLines(
    bool IsDefined,
    PptxSceneLineStyle Line);

internal sealed record PptxSceneChartTextRun(
    string Text,
    PptxSceneChartTextStyleOverride TextStyle);

internal enum PptxSceneChartDataLabelPosition
{
    BestFit,
    Bottom,
    Center,
    InsideBase,
    InsideEnd,
    Left,
    OutsideEnd,
    Right,
    Top,
    Unknown
}

internal sealed record PptxSceneChartShapeStyle(
    bool NoFill,
    PptxSceneFillStyle Fill,
    PptxSceneGradientFill? GradientFill,
    PptxScenePatternFill PatternFill,
    PptxSceneShapePictureFill PictureFill,
    PptxSceneLineStyle Line,
    PptxSceneGlow Glow,
    PptxSceneOuterShadow OuterShadow,
    PptxSceneChartEffectFamily Effects);

internal readonly record struct PptxSceneChartEffectFamily(
    bool HasEffectList,
    bool HasEffectDag,
    IReadOnlyList<string> UnsupportedEffectNames);

internal readonly record struct PptxSceneChartTextStyleOverride(
    string? FontFamily,
    string? RequestedTypeface,
    PptxThemeTypefaceSource? TypefaceSource,
    double? FontSize,
    double? CharacterSpacing,
    RgbColor? Color,
    double? Alpha,
    bool? Bold,
    bool? Italic,
    bool? Underline,
    bool? Strike)
{
    public PptxSceneChartTextStyleOverride(
        string? fontFamily,
        double? fontSize,
        RgbColor? color,
        double? alpha,
        bool? bold,
        bool? italic,
        bool? underline,
        bool? strike)
        : this(fontFamily, null, null, fontSize, null, color, alpha, bold, italic, underline, strike)
    {
    }
}

internal readonly record struct PptxSceneChartManualLayout(
    bool HasLayout,
    double? X,
    string XValue,
    double? Y,
    string YValue,
    double? Width,
    string WidthValue,
    double? Height,
    string HeightValue,
    PptxSceneChartManualLayoutTarget LayoutTargetKind,
    string LayoutTarget,
    PptxSceneChartManualLayoutMode XModeKind,
    string XMode,
    PptxSceneChartManualLayoutMode YModeKind,
    string YMode,
    PptxSceneChartManualLayoutMode WidthModeKind,
    string WidthMode,
    PptxSceneChartManualLayoutMode HeightModeKind,
    string HeightMode);

internal enum PptxSceneChartManualLayoutTarget
{
    Unknown,
    Inner,
    Outer
}

internal enum PptxSceneChartManualLayoutMode
{
    Unknown,
    Edge,
    Factor
}

internal sealed record PptxSceneChartSeries(
    int? Index,
    string IndexValue,
    int? Order,
    string OrderValue,
    string? Name,
    PptxSceneChartSeriesDataSources DataSources,
    IReadOnlyList<double> Values,
    IReadOnlyList<PptxSceneChartNumberPoint> ValuePoints,
    int? ValuePointCount,
    string ValuePointCountValue,
    string? ValueFormatCode,
    IReadOnlyList<string> Categories,
    IReadOnlyList<PptxSceneChartStringPoint> CategoryPoints,
    int? CategoryPointCount,
    string CategoryPointCountValue,
    IReadOnlyList<IReadOnlyList<PptxSceneChartStringPoint>> CategoryLevels,
    IReadOnlyList<double> XValues,
    IReadOnlyList<PptxSceneChartNumberPoint> XValuePoints,
    int? XValuePointCount,
    string XValuePointCountValue,
    string? XValueFormatCode,
    IReadOnlyList<double> YValues,
    IReadOnlyList<PptxSceneChartNumberPoint> YValuePoints,
    int? YValuePointCount,
    string YValuePointCountValue,
    string? YValueFormatCode,
    IReadOnlyList<double> BubbleSizes,
    IReadOnlyList<PptxSceneChartNumberPoint> BubbleSizePoints,
    int? BubbleSizePointCount,
    string BubbleSizePointCountValue,
    string? BubbleSizeFormatCode,
    PptxSceneFillStyle Fill,
    PptxScenePatternFill PatternFill,
    PptxSceneLineStyle Line,
    PptxSceneChartEffectFamily Effects,
    PptxSceneChartMarker Marker,
    IReadOnlyList<PptxSceneChartPointStyle> PointStyles,
    IReadOnlyList<string> RejectedPointStyleIndexValues,
    double? Explosion,
    string ExplosionValue,
    bool? Smooth,
    string SmoothValue,
    PptxSceneChartDataLabels DataLabels);

internal readonly record struct PptxSceneChartNumberPoint(
    int Index,
    string IndexValue,
    bool HasParsedIndex,
    double? Value,
    string Text,
    bool HasValueElement);

internal readonly record struct PptxSceneChartStringPoint(
    int Index,
    string IndexValue,
    bool HasParsedIndex,
    string Text,
    bool HasText);

internal sealed record PptxSceneChartSeriesDataSources(
    PptxSceneChartDataSource Name,
    PptxSceneChartDataSource Values,
    PptxSceneChartDataSource Categories,
    PptxSceneChartDataSource XValues,
    PptxSceneChartDataSource YValues,
    PptxSceneChartDataSource BubbleSizes);

internal readonly record struct PptxSceneChartDataSource(
    string? Formula,
    PptxSceneChartDataSourceReferenceKind ReferenceKindValue,
    string ReferenceKind,
    PptxSceneChartDataSourceCacheKind CacheKindValue,
    string CacheKind,
    bool HasCachedPoints);

internal enum PptxSceneChartDataSourceReferenceKind
{
    Unknown,
    StringReference,
    NumberReference,
    MultiLevelStringReference
}

internal enum PptxSceneChartDataSourceCacheKind
{
    Unknown,
    StringCache,
    NumberCache,
    MultiLevelStringCache
}

internal sealed record PptxSceneChartMarker(
    bool IsDefined,
    PptxSceneChartMarkerSymbol SymbolKind,
    string Symbol,
    string? SizeValue,
    double Size,
    PptxSceneFillStyle Fill,
    PptxSceneLineStyle Line);

internal enum PptxSceneChartMarkerSymbol
{
    Circle,
    Dash,
    Diamond,
    Dot,
    None,
    Plus,
    Square,
    Star,
    Triangle,
    X,
    Unknown
}

internal sealed record PptxSceneChartPointStyle(
    int Index,
    string IndexValue,
    PptxSceneFillStyle Fill,
    PptxScenePatternFill PatternFill,
    PptxSceneLineStyle Line,
    PptxSceneChartEffectFamily Effects,
    double? Explosion,
    string ExplosionValue);

internal sealed record PptxSceneChartAxis(
    string Id,
    PptxSceneChartAxisKind AxisKind,
    string Kind,
    PptxSceneChartAxisPosition PositionKind,
    string Position,
    string CrossAxisId,
    PptxSceneChartAxisCrosses CrossesKind,
    string Crosses,
    double? CrossesAt,
    string CrossesAtValue,
    PptxSceneChartAxisCrossBetween CrossBetweenKind,
    string CrossBetween,
    PptxSceneChartAxisOrientation OrientationKind,
    string Orientation,
    bool IsReversed,
    bool? IsDeleted,
    string IsDeletedValue,
    bool HasScaling,
    double? Minimum,
    string MinimumValue,
    double? Maximum,
    string MaximumValue,
    double? MajorUnit,
    string MajorUnitValue,
    double? MinorUnit,
    string MinorUnitValue,
    bool HasMajorGridlines,
    bool HasMinorGridlines,
    bool HasMajorGridlineElement,
    bool HasMinorGridlineElement,
    PptxSceneLineStyle Line,
    PptxSceneLineStyle MajorGridlineLine,
    PptxSceneLineStyle MinorGridlineLine,
    PptxSceneLineStyle MajorGridlineStyleLine,
    PptxSceneLineStyle MinorGridlineStyleLine,
    PptxSceneChartTextStyleOverride TextStyle,
    PptxSceneChartTickLabelPosition TickLabelPositionKind,
    string TickLabelPosition,
    PptxSceneChartAxisTickMark MajorTickMarkKind,
    string MajorTickMark,
    PptxSceneChartAxisTickMark MinorTickMarkKind,
    string MinorTickMark,
    int? LabelOffset,
    string LabelOffsetValue,
    int? TickLabelSkip,
    string TickLabelSkipValue,
    int? TickMarkSkip,
    string TickMarkSkipValue,
    bool? NoMultiLevelLabels,
    string NoMultiLevelLabelsValue,
    string? NumberFormat,
    PptxSceneChartNumberFormat NumberFormatInfo,
    PptxSceneChartTitle Title);

internal enum PptxSceneChartAxisKind
{
    Category,
    Date,
    Series,
    Value,
    Unknown
}

internal enum PptxSceneChartAxisCrosses
{
    AutoZero,
    Maximum,
    Minimum,
    Unknown
}

internal enum PptxSceneChartAxisCrossBetween
{
    Between,
    MidpointCategory,
    Unknown
}

internal enum PptxSceneChartAxisOrientation
{
    MinimumMaximum,
    MaximumMinimum,
    Unknown
}

internal enum PptxSceneChartAxisTickMark
{
    Cross,
    Inside,
    None,
    Outside,
    Unknown
}

internal enum PptxSceneChartTickLabelPosition
{
    High,
    Low,
    NextTo,
    None,
    Unknown
}

internal enum PptxSceneChartAxisPosition
{
    Bottom,
    Left,
    Right,
    Top,
    Unknown
}

internal sealed record PptxSceneChartTitle(
    string? Text,
    IReadOnlyList<PptxSceneChartTextRun> TextRuns,
    bool? IsAutoDeleted,
    string IsAutoDeletedValue,
    bool IsAutoGenerated,
    bool? Overlay,
    string OverlayValue,
    PptxSceneChartManualLayout Layout,
    PptxSceneChartShapeStyle ShapeStyle,
    PptxSceneChartTextBodyProperties TextBodyProperties,
    PptxSceneChartTextStyleOverride TextStyle);

internal sealed record PptxSceneChartLegend(
    PptxSceneChartLegendPosition PositionKind,
    string Position,
    bool? Overlay,
    string OverlayValue,
    bool IsDefined,
    bool? IsDeleted,
    string IsDeletedValue,
    PptxSceneChartManualLayout Layout,
    PptxSceneChartShapeStyle ShapeStyle,
    PptxSceneChartTextBodyProperties TextBodyProperties,
    PptxSceneChartTextStyleOverride TextStyle);

internal enum PptxSceneChartLegendPosition
{
    Bottom,
    Left,
    Right,
    Top,
    TopRight,
    Unknown
}
