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
    PptxSceneChartShapeStyle PlotAreaStyle,
    IReadOnlyDictionary<string, OoxRelationship>? Relationships = null);

// T02: the defined/undefined states build only through the factories below, so the
// IsDefined flag cannot disagree with construction. Undefined is the default value;
// Defined carries whatever the package yielded (possibly-unresolved targets stay null
// and read as missing downstream, exactly as before).
internal readonly record struct PptxSceneChartExternalData
{
    public bool IsDefined { get; }

    public string? RelationshipId { get; }

    public string? TargetPartName { get; }

    public PptxScenePackageResource? Resource { get; }

    public bool? AutoUpdate { get; }

    public string AutoUpdateValue { get; }

    public static PptxSceneChartExternalData Undefined => default;

    private PptxSceneChartExternalData(
        bool isDefined,
        string? relationshipId,
        string? targetPartName,
        PptxScenePackageResource? resource,
        bool? autoUpdate,
        string autoUpdateValue)
    {
        IsDefined = isDefined;
        RelationshipId = relationshipId;
        TargetPartName = targetPartName;
        Resource = resource;
        AutoUpdate = autoUpdate;
        AutoUpdateValue = autoUpdateValue;
    }

    public static PptxSceneChartExternalData Defined(
        string? relationshipId,
        string? targetPartName,
        PptxScenePackageResource? resource,
        bool? autoUpdate,
        string autoUpdateValue)
    {
        ArgumentNullException.ThrowIfNull(autoUpdateValue);
        return new PptxSceneChartExternalData(true, relationshipId, targetPartName, resource, autoUpdate, autoUpdateValue);
    }
}

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

// T02: defined color styles always resolve a part (producers pass the loaded part
// name); undefined carries an optional known target for diagnostics. Both states build
// only through the factories, so three drifting hand-built undefined literals collapse
// to one spelling.
internal readonly record struct PptxSceneChartColorStyle
{
    public bool IsDefined { get; }

    public string? PartName { get; }

    public string Method { get; }

    public string Id { get; }

    public IReadOnlyList<RgbColor> Colors { get; }

    public int VariationCount { get; }

    public IReadOnlyList<PptxSceneChartColorDeclaration> Declarations { get; }

    public IReadOnlyList<PptxSceneChartColorDeclaration> RootDeclarations { get; }

    public IReadOnlyList<PptxSceneChartColorVariation> Variations { get; }

    public XDocument? ColorStyleXml { get; }

    private PptxSceneChartColorStyle(
        bool isDefined,
        string? partName,
        string method,
        string id,
        IReadOnlyList<RgbColor> colors,
        int variationCount,
        IReadOnlyList<PptxSceneChartColorDeclaration> declarations,
        IReadOnlyList<PptxSceneChartColorDeclaration> rootDeclarations,
        IReadOnlyList<PptxSceneChartColorVariation> variations,
        XDocument? colorStyleXml)
    {
        IsDefined = isDefined;
        PartName = partName;
        Method = method;
        Id = id;
        Colors = colors;
        VariationCount = variationCount;
        Declarations = declarations;
        RootDeclarations = rootDeclarations;
        Variations = variations;
        ColorStyleXml = colorStyleXml;
    }

    public static PptxSceneChartColorStyle Undefined(string? partName)
    {
        return new PptxSceneChartColorStyle(false, partName, string.Empty, string.Empty, [], 0, [], [], [], null);
    }

    public static PptxSceneChartColorStyle Defined(
        string partName,
        string method,
        string id,
        IReadOnlyList<RgbColor> colors,
        int variationCount,
        IReadOnlyList<PptxSceneChartColorDeclaration> declarations,
        IReadOnlyList<PptxSceneChartColorDeclaration> rootDeclarations,
        IReadOnlyList<PptxSceneChartColorVariation> variations,
        XDocument? colorStyleXml)
    {
        ArgumentNullException.ThrowIfNull(partName);
        return new PptxSceneChartColorStyle(true, partName, method, id, colors, variationCount, declarations, rootDeclarations, variations, colorStyleXml);
    }
}

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

// T02: same defined/undefined discipline as the color style above.
internal sealed record PptxSceneChartStyle
{
    public bool IsDefined { get; }

    public string? PartName { get; }

    public string Id { get; }

    public XDocument? StyleXml { get; }

    public IReadOnlyList<PptxSceneChartStyleEntry> Entries { get; }

    private PptxSceneChartStyle(
        bool isDefined,
        string? partName,
        string id,
        XDocument? styleXml,
        IReadOnlyList<PptxSceneChartStyleEntry> entries)
    {
        IsDefined = isDefined;
        PartName = partName;
        Id = id;
        StyleXml = styleXml;
        Entries = entries;
    }

    public static PptxSceneChartStyle Undefined(string? partName)
    {
        return new PptxSceneChartStyle(false, partName, string.Empty, null, []);
    }

    public static PptxSceneChartStyle Defined(
        string partName,
        string id,
        XDocument? styleXml,
        IReadOnlyList<PptxSceneChartStyleEntry> entries)
    {
        ArgumentNullException.ThrowIfNull(partName);
        ArgumentNullException.ThrowIfNull(entries);
        return new PptxSceneChartStyle(true, partName, id, styleXml, entries);
    }
}

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

// T02: kind enums derive from their raw spellings inside Defined, so a plot cannot
// carry a grouping/direction/style enum that disagrees with its stored spelling.
internal sealed record PptxSceneChartPlot
{
    public PptxSceneChartPlotKind PlotKind { get; }

    public string Kind { get; }

    public int PlotAreaIndex { get; }

    public int KindIndex { get; }

    public int SeriesCount { get; }

    public IReadOnlyList<string> AxisIds { get; }

    public IReadOnlyList<PptxSceneChartSeries> Series { get; }

    public PptxSceneChartGrouping GroupingKind { get; }

    public string Grouping { get; }

    public PptxSceneChartBarDirection BarDirectionKind { get; }

    public string BarDirection { get; }

    public PptxSceneChartScatterStyle ScatterStyleKind { get; }

    public string ScatterStyle { get; }

    public PptxSceneChartRadarStyle RadarStyleKind { get; }

    public string RadarStyle { get; }

    public bool? MarkersEnabled { get; }

    public string MarkersEnabledValue { get; }

    public bool? VaryColors { get; }

    public string VaryColorsValue { get; }

    public double? GapWidth { get; }

    public string GapWidthValue { get; }

    public double? Overlap { get; }

    public string OverlapValue { get; }

    public double? HoleSize { get; }

    public string HoleSizeValue { get; }

    public double? FirstSliceAngle { get; }

    public string FirstSliceAngleValue { get; }

    public PptxSceneChartDataLabels DataLabels { get; }

    public XElement Source { get; }

    private PptxSceneChartPlot(
        PptxSceneChartPlotKind plotKind,
        string kind,
        int plotAreaIndex,
        int kindIndex,
        int seriesCount,
        IReadOnlyList<string> axisIds,
        IReadOnlyList<PptxSceneChartSeries> series,
        PptxSceneChartGrouping groupingKind,
        string grouping,
        PptxSceneChartBarDirection barDirectionKind,
        string barDirection,
        PptxSceneChartScatterStyle scatterStyleKind,
        string scatterStyle,
        PptxSceneChartRadarStyle radarStyleKind,
        string radarStyle,
        bool? markersEnabled,
        string markersEnabledValue,
        bool? varyColors,
        string varyColorsValue,
        double? gapWidth,
        string gapWidthValue,
        double? overlap,
        string overlapValue,
        double? holeSize,
        string holeSizeValue,
        double? firstSliceAngle,
        string firstSliceAngleValue,
        PptxSceneChartDataLabels dataLabels,
        XElement source)
    {
        PlotKind = plotKind;
        Kind = kind;
        PlotAreaIndex = plotAreaIndex;
        KindIndex = kindIndex;
        SeriesCount = seriesCount;
        AxisIds = axisIds;
        Series = series;
        GroupingKind = groupingKind;
        Grouping = grouping;
        BarDirectionKind = barDirectionKind;
        BarDirection = barDirection;
        ScatterStyleKind = scatterStyleKind;
        ScatterStyle = scatterStyle;
        RadarStyleKind = radarStyleKind;
        RadarStyle = radarStyle;
        MarkersEnabled = markersEnabled;
        MarkersEnabledValue = markersEnabledValue;
        VaryColors = varyColors;
        VaryColorsValue = varyColorsValue;
        GapWidth = gapWidth;
        GapWidthValue = gapWidthValue;
        Overlap = overlap;
        OverlapValue = overlapValue;
        HoleSize = holeSize;
        HoleSizeValue = holeSizeValue;
        FirstSliceAngle = firstSliceAngle;
        FirstSliceAngleValue = firstSliceAngleValue;
        DataLabels = dataLabels;
        Source = source;
    }

    public static PptxSceneChartPlot Defined(
        string kind,
        int plotAreaIndex,
        int kindIndex,
        int seriesCount,
        IReadOnlyList<string> axisIds,
        IReadOnlyList<PptxSceneChartSeries> series,
        string grouping,
        string barDirection,
        string scatterStyle,
        string radarStyle,
        bool? markersEnabled,
        string markersEnabledValue,
        bool? varyColors,
        string varyColorsValue,
        double? gapWidth,
        string gapWidthValue,
        double? overlap,
        string overlapValue,
        double? holeSize,
        string holeSizeValue,
        double? firstSliceAngle,
        string firstSliceAngleValue,
        PptxSceneChartDataLabels dataLabels,
        XElement source)
    {
        return new PptxSceneChartPlot(
            PptxSceneBuilder.ParseChartPlotKind(kind),
            kind,
            plotAreaIndex,
            kindIndex,
            seriesCount,
            axisIds,
            series,
            PptxSceneBuilder.ParseChartGrouping(grouping),
            grouping,
            PptxSceneBuilder.ParseChartBarDirection(barDirection),
            barDirection,
            PptxSceneBuilder.ParseChartScatterStyle(scatterStyle),
            scatterStyle,
            PptxSceneBuilder.ParseChartRadarStyle(radarStyle),
            radarStyle,
            markersEnabled,
            markersEnabledValue,
            varyColors,
            varyColorsValue,
            gapWidth,
            gapWidthValue,
            overlap,
            overlapValue,
            holeSize,
            holeSizeValue,
            firstSliceAngle,
            firstSliceAngleValue,
            dataLabels,
            source);
    }
}

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

// T02: the position enum derives from its raw spelling inside Defined, so labels
// cannot carry a position enum that disagrees with the stored spelling.
internal sealed record PptxSceneChartDataLabels
{
    public bool? ShowValue { get; }

    public string ShowValueValue { get; }

    public bool? ShowPercent { get; }

    public string ShowPercentValue { get; }

    public bool? ShowCategoryName { get; }

    public string ShowCategoryNameValue { get; }

    public bool? ShowSeriesName { get; }

    public string ShowSeriesNameValue { get; }

    public bool? ShowLeaderLines { get; }

    public string ShowLeaderLinesValue { get; }

    public bool? ShowLegendKey { get; }

    public string ShowLegendKeyValue { get; }

    public bool? ShowBubbleSize { get; }

    public string ShowBubbleSizeValue { get; }

    public PptxSceneChartLeaderLines LeaderLines { get; }

    public PptxSceneChartDataLabelPosition PositionKind { get; }

    public string Position { get; }

    public string Separator { get; }

    public string NumberFormat { get; }

    public PptxSceneChartNumberFormat NumberFormatInfo { get; }

    public PptxSceneChartManualLayout Layout { get; }

    public PptxSceneChartTextStyleOverride TextStyle { get; }

    public PptxSceneChartTextBodyProperties TextBodyProperties { get; }

    public PptxSceneChartShapeStyle ShapeStyle { get; }

    public IReadOnlyList<string> RejectedOverrideIndexValues { get; }

    public IReadOnlyList<PptxSceneChartDataLabelOverride> Overrides { get; }

    public bool IsDefined { get; }

    private PptxSceneChartDataLabels(
        bool? showValue,
        string showValueValue,
        bool? showPercent,
        string showPercentValue,
        bool? showCategoryName,
        string showCategoryNameValue,
        bool? showSeriesName,
        string showSeriesNameValue,
        bool? showLeaderLines,
        string showLeaderLinesValue,
        bool? showLegendKey,
        string showLegendKeyValue,
        bool? showBubbleSize,
        string showBubbleSizeValue,
        PptxSceneChartLeaderLines leaderLines,
        PptxSceneChartDataLabelPosition positionKind,
        string position,
        string separator,
        string numberFormat,
        PptxSceneChartNumberFormat numberFormatInfo,
        PptxSceneChartManualLayout layout,
        PptxSceneChartTextStyleOverride textStyle,
        PptxSceneChartTextBodyProperties textBodyProperties,
        PptxSceneChartShapeStyle shapeStyle,
        IReadOnlyList<string> rejectedOverrideIndexValues,
        IReadOnlyList<PptxSceneChartDataLabelOverride> overrides,
        bool isDefined)
    {
        ShowValue = showValue;
        ShowValueValue = showValueValue;
        ShowPercent = showPercent;
        ShowPercentValue = showPercentValue;
        ShowCategoryName = showCategoryName;
        ShowCategoryNameValue = showCategoryNameValue;
        ShowSeriesName = showSeriesName;
        ShowSeriesNameValue = showSeriesNameValue;
        ShowLeaderLines = showLeaderLines;
        ShowLeaderLinesValue = showLeaderLinesValue;
        ShowLegendKey = showLegendKey;
        ShowLegendKeyValue = showLegendKeyValue;
        ShowBubbleSize = showBubbleSize;
        ShowBubbleSizeValue = showBubbleSizeValue;
        LeaderLines = leaderLines;
        PositionKind = positionKind;
        Position = position;
        Separator = separator;
        NumberFormat = numberFormat;
        NumberFormatInfo = numberFormatInfo;
        Layout = layout;
        TextStyle = textStyle;
        TextBodyProperties = textBodyProperties;
        ShapeStyle = shapeStyle;
        RejectedOverrideIndexValues = rejectedOverrideIndexValues;
        Overrides = overrides;
        IsDefined = isDefined;
    }

    public static PptxSceneChartDataLabels Defined(
        bool? showValue,
        string showValueValue,
        bool? showPercent,
        string showPercentValue,
        bool? showCategoryName,
        string showCategoryNameValue,
        bool? showSeriesName,
        string showSeriesNameValue,
        bool? showLeaderLines,
        string showLeaderLinesValue,
        bool? showLegendKey,
        string showLegendKeyValue,
        bool? showBubbleSize,
        string showBubbleSizeValue,
        PptxSceneChartLeaderLines leaderLines,
        string position,
        string separator,
        string numberFormat,
        PptxSceneChartNumberFormat numberFormatInfo,
        PptxSceneChartManualLayout layout,
        PptxSceneChartTextStyleOverride textStyle,
        PptxSceneChartTextBodyProperties textBodyProperties,
        PptxSceneChartShapeStyle shapeStyle,
        IReadOnlyList<string> rejectedOverrideIndexValues,
        IReadOnlyList<PptxSceneChartDataLabelOverride> overrides,
        bool isDefined)
    {
        return new PptxSceneChartDataLabels(
            showValue,
            showValueValue,
            showPercent,
            showPercentValue,
            showCategoryName,
            showCategoryNameValue,
            showSeriesName,
            showSeriesNameValue,
            showLeaderLines,
            showLeaderLinesValue,
            showLegendKey,
            showLegendKeyValue,
            showBubbleSize,
            showBubbleSizeValue,
            leaderLines,
            PptxSceneBuilder.ParseChartDataLabelPosition(position),
            position,
            separator,
            numberFormat,
            numberFormatInfo,
            layout,
            textStyle,
            textBodyProperties,
            shapeStyle,
            rejectedOverrideIndexValues,
            overrides,
            isDefined);
    }
}

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
    PptxSceneChartTextStyleOverride TextStyle,
    string? HyperlinkClickId = null,
    string? HyperlinkClickAction = null);

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

// T02: axis kind/position/crossing/orientation/tick enums derive from their raw
// spellings inside Defined, so an axis cannot carry an enum that disagrees with
// its stored spelling.
internal sealed record PptxSceneChartAxis
{
    public string Id { get; }

    public PptxSceneChartAxisKind AxisKind { get; }

    public string Kind { get; }

    public PptxSceneChartAxisPosition PositionKind { get; }

    public string Position { get; }

    public string CrossAxisId { get; }

    public PptxSceneChartAxisCrosses CrossesKind { get; }

    public string Crosses { get; }

    public double? CrossesAt { get; }

    public string CrossesAtValue { get; }

    public PptxSceneChartAxisCrossBetween CrossBetweenKind { get; }

    public string CrossBetween { get; }

    public PptxSceneChartAxisOrientation OrientationKind { get; }

    public string Orientation { get; }

    public bool IsReversed { get; }

    public bool? IsDeleted { get; }

    public string IsDeletedValue { get; }

    public bool HasScaling { get; }

    public double? Minimum { get; }

    public string MinimumValue { get; }

    public double? Maximum { get; }

    public string MaximumValue { get; }

    public double? MajorUnit { get; }

    public string MajorUnitValue { get; }

    public double? MinorUnit { get; }

    public string MinorUnitValue { get; }

    public bool HasMajorGridlines { get; }

    public bool HasMinorGridlines { get; }

    public bool HasMajorGridlineElement { get; }

    public bool HasMinorGridlineElement { get; }

    public PptxSceneLineStyle Line { get; }

    public PptxSceneLineStyle MajorGridlineLine { get; }

    public PptxSceneLineStyle MinorGridlineLine { get; }

    public PptxSceneLineStyle MajorGridlineStyleLine { get; }

    public PptxSceneLineStyle MinorGridlineStyleLine { get; }

    public PptxSceneChartTextStyleOverride TextStyle { get; }

    public PptxSceneChartTickLabelPosition TickLabelPositionKind { get; }

    public string TickLabelPosition { get; }

    public PptxSceneChartAxisTickMark MajorTickMarkKind { get; }

    public string MajorTickMark { get; }

    public PptxSceneChartAxisTickMark MinorTickMarkKind { get; }

    public string MinorTickMark { get; }

    public int? LabelOffset { get; }

    public string LabelOffsetValue { get; }

    public int? TickLabelSkip { get; }

    public string TickLabelSkipValue { get; }

    public int? TickMarkSkip { get; }

    public string TickMarkSkipValue { get; }

    public bool? NoMultiLevelLabels { get; }

    public string NoMultiLevelLabelsValue { get; }

    public string? NumberFormat { get; }

    public PptxSceneChartNumberFormat NumberFormatInfo { get; }

    public PptxSceneChartTitle Title { get; }

    private PptxSceneChartAxis(
        string id,
        PptxSceneChartAxisKind axisKind,
        string kind,
        PptxSceneChartAxisPosition positionKind,
        string position,
        string crossAxisId,
        PptxSceneChartAxisCrosses crossesKind,
        string crosses,
        double? crossesAt,
        string crossesAtValue,
        PptxSceneChartAxisCrossBetween crossBetweenKind,
        string crossBetween,
        PptxSceneChartAxisOrientation orientationKind,
        string orientation,
        bool isReversed,
        bool? isDeleted,
        string isDeletedValue,
        bool hasScaling,
        double? minimum,
        string minimumValue,
        double? maximum,
        string maximumValue,
        double? majorUnit,
        string majorUnitValue,
        double? minorUnit,
        string minorUnitValue,
        bool hasMajorGridlines,
        bool hasMinorGridlines,
        bool hasMajorGridlineElement,
        bool hasMinorGridlineElement,
        PptxSceneLineStyle line,
        PptxSceneLineStyle majorGridlineLine,
        PptxSceneLineStyle minorGridlineLine,
        PptxSceneLineStyle majorGridlineStyleLine,
        PptxSceneLineStyle minorGridlineStyleLine,
        PptxSceneChartTextStyleOverride textStyle,
        PptxSceneChartTickLabelPosition tickLabelPositionKind,
        string tickLabelPosition,
        PptxSceneChartAxisTickMark majorTickMarkKind,
        string majorTickMark,
        PptxSceneChartAxisTickMark minorTickMarkKind,
        string minorTickMark,
        int? labelOffset,
        string labelOffsetValue,
        int? tickLabelSkip,
        string tickLabelSkipValue,
        int? tickMarkSkip,
        string tickMarkSkipValue,
        bool? noMultiLevelLabels,
        string noMultiLevelLabelsValue,
        string? numberFormat,
        PptxSceneChartNumberFormat numberFormatInfo,
        PptxSceneChartTitle title)
    {
        Id = id;
        AxisKind = axisKind;
        Kind = kind;
        PositionKind = positionKind;
        Position = position;
        CrossAxisId = crossAxisId;
        CrossesKind = crossesKind;
        Crosses = crosses;
        CrossesAt = crossesAt;
        CrossesAtValue = crossesAtValue;
        CrossBetweenKind = crossBetweenKind;
        CrossBetween = crossBetween;
        OrientationKind = orientationKind;
        Orientation = orientation;
        IsReversed = isReversed;
        IsDeleted = isDeleted;
        IsDeletedValue = isDeletedValue;
        HasScaling = hasScaling;
        Minimum = minimum;
        MinimumValue = minimumValue;
        Maximum = maximum;
        MaximumValue = maximumValue;
        MajorUnit = majorUnit;
        MajorUnitValue = majorUnitValue;
        MinorUnit = minorUnit;
        MinorUnitValue = minorUnitValue;
        HasMajorGridlines = hasMajorGridlines;
        HasMinorGridlines = hasMinorGridlines;
        HasMajorGridlineElement = hasMajorGridlineElement;
        HasMinorGridlineElement = hasMinorGridlineElement;
        Line = line;
        MajorGridlineLine = majorGridlineLine;
        MinorGridlineLine = minorGridlineLine;
        MajorGridlineStyleLine = majorGridlineStyleLine;
        MinorGridlineStyleLine = minorGridlineStyleLine;
        TextStyle = textStyle;
        TickLabelPositionKind = tickLabelPositionKind;
        TickLabelPosition = tickLabelPosition;
        MajorTickMarkKind = majorTickMarkKind;
        MajorTickMark = majorTickMark;
        MinorTickMarkKind = minorTickMarkKind;
        MinorTickMark = minorTickMark;
        LabelOffset = labelOffset;
        LabelOffsetValue = labelOffsetValue;
        TickLabelSkip = tickLabelSkip;
        TickLabelSkipValue = tickLabelSkipValue;
        TickMarkSkip = tickMarkSkip;
        TickMarkSkipValue = tickMarkSkipValue;
        NoMultiLevelLabels = noMultiLevelLabels;
        NoMultiLevelLabelsValue = noMultiLevelLabelsValue;
        NumberFormat = numberFormat;
        NumberFormatInfo = numberFormatInfo;
        Title = title;
    }

    public static PptxSceneChartAxis Defined(
        string id,
        string kind,
        string position,
        string crossAxisId,
        string crosses,
        double? crossesAt,
        string crossesAtValue,
        string crossBetween,
        string orientation,
        bool? isDeleted,
        string isDeletedValue,
        bool hasScaling,
        double? minimum,
        string minimumValue,
        double? maximum,
        string maximumValue,
        double? majorUnit,
        string majorUnitValue,
        double? minorUnit,
        string minorUnitValue,
        bool hasMajorGridlines,
        bool hasMinorGridlines,
        bool hasMajorGridlineElement,
        bool hasMinorGridlineElement,
        PptxSceneLineStyle line,
        PptxSceneLineStyle majorGridlineLine,
        PptxSceneLineStyle minorGridlineLine,
        PptxSceneLineStyle majorGridlineStyleLine,
        PptxSceneLineStyle minorGridlineStyleLine,
        PptxSceneChartTextStyleOverride textStyle,
        string tickLabelPosition,
        string majorTickMark,
        string minorTickMark,
        int? labelOffset,
        string labelOffsetValue,
        int? tickLabelSkip,
        string tickLabelSkipValue,
        int? tickMarkSkip,
        string tickMarkSkipValue,
        bool? noMultiLevelLabels,
        string noMultiLevelLabelsValue,
        string? numberFormat,
        PptxSceneChartNumberFormat numberFormatInfo,
        PptxSceneChartTitle title)
    {
        return new PptxSceneChartAxis(
            id,
            PptxSceneBuilder.ParseChartAxisKind(kind),
            kind,
            PptxSceneBuilder.ParseChartAxisPosition(position),
            position,
            crossAxisId,
            PptxSceneBuilder.ParseChartAxisCrosses(crosses),
            crosses,
            crossesAt,
            crossesAtValue,
            PptxSceneBuilder.ParseChartAxisCrossBetween(crossBetween),
            crossBetween,
            PptxSceneBuilder.ParseChartAxisOrientation(orientation),
            orientation,
            PptxSceneBuilder.ParseChartAxisOrientation(orientation) == PptxSceneChartAxisOrientation.MaximumMinimum,
            isDeleted,
            isDeletedValue,
            hasScaling,
            minimum,
            minimumValue,
            maximum,
            maximumValue,
            majorUnit,
            majorUnitValue,
            minorUnit,
            minorUnitValue,
            hasMajorGridlines,
            hasMinorGridlines,
            hasMajorGridlineElement,
            hasMinorGridlineElement,
            line,
            majorGridlineLine,
            minorGridlineLine,
            majorGridlineStyleLine,
            minorGridlineStyleLine,
            textStyle,
            PptxSceneBuilder.ParseChartTickLabelPosition(tickLabelPosition),
            tickLabelPosition,
            PptxSceneBuilder.ParseChartAxisTickMark(majorTickMark),
            majorTickMark,
            PptxSceneBuilder.ParseChartAxisTickMark(minorTickMark),
            minorTickMark,
            labelOffset,
            labelOffsetValue,
            tickLabelSkip,
            tickLabelSkipValue,
            tickMarkSkip,
            tickMarkSkipValue,
            noMultiLevelLabels,
            noMultiLevelLabelsValue,
            numberFormat,
            numberFormatInfo,
            title);
    }
}

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
