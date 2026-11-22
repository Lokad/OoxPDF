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

    private PptxSceneChartExternalData(
        string? relationshipId,
        string? targetPartName,
        PptxScenePackageResource? resource)
    {
        IsDefined = true;
        RelationshipId = relationshipId;
        TargetPartName = targetPartName;
        Resource = resource;
    }

    public static PptxSceneChartExternalData Defined(
        string? relationshipId,
        string? targetPartName,
        PptxScenePackageResource? resource) => new(relationshipId, targetPartName, resource);
}

internal readonly record struct PptxSceneChartOptions(
    bool? Date1904,
    bool? PlotVisibleOnly,
    PptxSceneChartDisplayBlanksAs DisplayBlanksAsKind);

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

    private PptxSceneChartColorStyle(
        bool isDefined,
        string? partName,
        string method,
        string id,
        IReadOnlyList<RgbColor> colors,
        int variationCount,
        IReadOnlyList<PptxSceneChartColorDeclaration> declarations,
        IReadOnlyList<PptxSceneChartColorDeclaration> rootDeclarations,
        IReadOnlyList<PptxSceneChartColorVariation> variations)
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
    }

    public static PptxSceneChartColorStyle Undefined(string? partName)
    {
        return new PptxSceneChartColorStyle(false, partName, string.Empty, string.Empty, [], 0, [], [], []);
    }

    public static PptxSceneChartColorStyle Defined(
        string partName,
        string method,
        string id,
        IReadOnlyList<RgbColor> colors,
        int variationCount,
        IReadOnlyList<PptxSceneChartColorDeclaration> declarations,
        IReadOnlyList<PptxSceneChartColorDeclaration> rootDeclarations,
        IReadOnlyList<PptxSceneChartColorVariation> variations)
    {
        ArgumentNullException.ThrowIfNull(partName);
        return new PptxSceneChartColorStyle(true, partName, method, id, colors, variationCount, declarations, rootDeclarations, variations);
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

    public IReadOnlyList<PptxSceneChartStyleEntry> Entries { get; }

    private PptxSceneChartStyle(
        bool isDefined,
        string? partName,
        string id,
        IReadOnlyList<PptxSceneChartStyleEntry> entries)
    {
        IsDefined = isDefined;
        PartName = partName;
        Id = id;
        Entries = entries;
    }

    public static PptxSceneChartStyle Undefined(string? partName)
    {
        return new PptxSceneChartStyle(false, partName, string.Empty, []);
    }

    public static PptxSceneChartStyle Defined(
        string partName,
        string id,
        IReadOnlyList<PptxSceneChartStyleEntry> entries)
    {
        ArgumentNullException.ThrowIfNull(partName);
        ArgumentNullException.ThrowIfNull(entries);
        return new PptxSceneChartStyle(true, partName, id, entries);
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

internal sealed record PptxSceneChartPlot
{
    public PptxSceneChartPlotKind PlotKind { get; }

    public string Kind { get; }


    public IReadOnlyList<string> AxisIds { get; }

    public IReadOnlyList<PptxSceneChartSeries> Series { get; }

    public PptxSceneChartGrouping GroupingKind { get; }

    public PptxSceneChartBarDirection BarDirectionKind { get; }

    public PptxSceneChartScatterStyle ScatterStyleKind { get; }

    public string ScatterStyle { get; }

    public double? BubbleScale { get; }

    public string BubbleScaleValue { get; }

    public string BubbleSizeRepresents { get; }

    public string ShowNegativeBubblesValue { get; }

    public PptxSceneChartRadarStyle RadarStyleKind { get; }

    public bool? MarkersEnabled { get; }

    public bool? VaryColors { get; }

    public double? GapWidth { get; }

    public double? Overlap { get; }

    public double? HoleSize { get; }

    public double? FirstSliceAngle { get; }

    public PptxSceneChartDataLabels DataLabels { get; }

    public XElement Source { get; }

    public PptxSceneChartPlot(
        PptxSceneChartPlotKind plotKind,
        string kind,
        IReadOnlyList<string> axisIds,
        IReadOnlyList<PptxSceneChartSeries> series,
        string grouping,
        string barDirection,
        string scatterStyle,
        string radarStyle,
        bool? markersEnabled,
        bool? varyColors,
        double? gapWidth,
        double? overlap,
        double? holeSize,
        double? firstSliceAngle,
        PptxSceneChartDataLabels dataLabels,
        XElement source)
    {
        PlotKind = plotKind;
        Kind = kind;
        AxisIds = axisIds;
        Series = series;
        GroupingKind = PptxSceneBuilder.ParseChartGrouping(grouping);
        BarDirectionKind = PptxSceneBuilder.ParseChartBarDirection(barDirection);
        ScatterStyleKind = PptxSceneBuilder.ParseChartScatterStyle(scatterStyle);
        ScatterStyle = scatterStyle;
        (BubbleScale, BubbleScaleValue) = PptxSceneBuilder.ReadChartElementDoubleWithValue(source, "bubbleScale");
        BubbleSizeRepresents = PptxSceneBuilder.ReadChartElementValue(source, "sizeRepresents");
        ShowNegativeBubblesValue = PptxSceneBuilder.ReadChartElementValue(source, "showNegBubbles");
        RadarStyleKind = PptxSceneBuilder.ParseChartRadarStyle(radarStyle);
        MarkersEnabled = markersEnabled;
        VaryColors = varyColors;
        GapWidth = gapWidth;
        Overlap = overlap;
        HoleSize = holeSize;
        FirstSliceAngle = firstSliceAngle;
        DataLabels = dataLabels;
        Source = source;
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
    public bool? IsDeleted { get; }

    public bool? ShowValue { get; }

    public bool? ShowPercent { get; }

    public bool? ShowCategoryName { get; }

    public bool? ShowSeriesName { get; }

    public bool? ShowLeaderLines { get; }

    public bool? ShowLegendKey { get; }

    public bool? ShowBubbleSize { get; }

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
        bool? showPercent,
        bool? showCategoryName,
        bool? showSeriesName,
        bool? showLeaderLines,
        bool? showLegendKey,
        bool? showBubbleSize,
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
        bool isDefined,
        bool? isDeleted)
    {
        IsDeleted = isDeleted;
        ShowValue = showValue;
        ShowPercent = showPercent;
        ShowCategoryName = showCategoryName;
        ShowSeriesName = showSeriesName;
        ShowLeaderLines = showLeaderLines;
        ShowLegendKey = showLegendKey;
        ShowBubbleSize = showBubbleSize;
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
        bool? showPercent,
        bool? showCategoryName,
        bool? showSeriesName,
        bool? showLeaderLines,
        bool? showLegendKey,
        bool? showBubbleSize,
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
        bool isDefined,
        bool? isDeleted = null)
    {
        return new PptxSceneChartDataLabels(
            showValue,
            showPercent,
            showCategoryName,
            showSeriesName,
            showLeaderLines,
            showLegendKey,
            showBubbleSize,
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
            isDefined,
            isDeleted);
    }
}

internal enum PptxSceneChartPointTextScope { Container, Chart, ImplicitTextBody }

internal sealed record PptxSceneChartDataLabelOverride(
    int Index,
    string IndexValue,
    bool? IsDeleted,
    bool? ShowValue,
    bool? ShowPercent,
    bool? ShowCategoryName,
    bool? ShowSeriesName,
    bool? ShowLeaderLines,
    bool? ShowLegendKey,
    bool? ShowBubbleSize,
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
    PptxSceneChartShapeStyle ShapeStyle,
    PptxSceneChartPointTextScope TextScope);

internal readonly record struct PptxSceneChartTextBodyProperties(
    double? RotationDegrees,
    string RotationValue,
    string OrientationValue,
    string VerticalOverflowValue,
    string WrapValue = "");

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
    string? HyperlinkClickAction = null,
    int LineIndex = 0);

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
    string HeightMode)
{
    // Raw presence stays inspectable. Empty layout metadata supplies no plot
    // coordinates; malformed authored values retain their existing fallback.
    public bool HasGeometryValues => HasLayout &&
        (X.HasValue || Y.HasValue || Width.HasValue || Height.HasValue ||
         !string.IsNullOrEmpty(XValue) || !string.IsNullOrEmpty(YValue) ||
         !string.IsNullOrEmpty(WidthValue) || !string.IsNullOrEmpty(HeightValue));
}

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
    IReadOnlyList<PptxSceneChartNumberPoint> ValuePoints,
    int? ValuePointCount,
    string ValuePointCountValue,
    string? ValueFormatCode,
    IReadOnlyList<PptxSceneChartStringPoint> CategoryPoints,
    int? CategoryPointCount,
    string CategoryPointCountValue,
    IReadOnlyList<IReadOnlyList<PptxSceneChartStringPoint>> CategoryLevels,
    IReadOnlyList<PptxSceneChartNumberPoint> XValuePoints,
    int? XValuePointCount,
    string XValuePointCountValue,
    string? XValueFormatCode,
    IReadOnlyList<PptxSceneChartNumberPoint> YValuePoints,
    int? YValuePointCount,
    string YValuePointCountValue,
    string? YValueFormatCode,
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
    PptxSceneChartDataLabels DataLabels)
{
    public PptxSceneChartNegativeInversion NegativeInversion { get; init; }
}

// One source-owned policy for both series defaults and point overrides.
internal readonly record struct PptxSceneChartNegativeInversion(
    bool? SeriesValue,
    IReadOnlyDictionary<int, bool>? PointValues)
{
    public bool ForPoint(int index) =>
        PointValues is not null && PointValues.TryGetValue(index, out bool value)
            ? value
            : SeriesValue ?? true;
}

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
    PptxSceneLineStyle Line)
{
    public bool SymbolSpecified { get; init; }
}

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
    string ExplosionValue)
{
    // An explicit empty appearance also gives this point a category identity.
    public bool HasShapeProperties { get; init; }

    public PptxSceneChartMarker? Marker { get; init; }

    public bool HasMarkerAppearance => Marker is { } marker &&
        (marker.SymbolSpecified || marker.SizeValue is not null || marker.Fill.HasFill ||
         marker.Line.HasLine || marker.Line.NoFill || marker.Line.WidthSpecified);
}

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

    public PptxSceneChartAxisCrossBetween CrossBetweenKind { get; }

    public string CrossBetween { get; }

    public PptxSceneChartAxisOrientation OrientationKind { get; }

    public string Orientation { get; }

    public bool? IsDeleted { get; }

    public bool HasScaling { get; }

    public bool HasLogarithmicScale { get; private init; }

    public double? Minimum { get; }

    public string MinimumValue { get; }

    public double? Maximum { get; }

    public string MaximumValue { get; }

    public double? MajorUnit { get; }

    public string MajorUnitValue { get; }

    public double? MinorUnit { get; }

    public bool HasMajorGridlines { get; }

    public bool HasMinorGridlines { get; }

    public bool HasMajorGridlineElement { get; }

    public PptxSceneLineStyle Line { get; }

    public PptxSceneLineStyle MajorGridlineLine { get; }

    public PptxSceneLineStyle MinorGridlineLine { get; }

    public PptxSceneLineStyle MajorGridlineStyleLine { get; }

    public PptxSceneLineStyle MinorGridlineStyleLine { get; }

    public PptxSceneChartTextStyleOverride TextStyle { get; }

    public PptxSceneChartTextBodyProperties TextBodyProperties { get; init; }

    public PptxSceneChartTickLabelPosition TickLabelPositionKind { get; }

    public string TickLabelPosition { get; }

    public PptxSceneChartAxisTickMark MajorTickMarkKind { get; }

    public string MajorTickMark { get; }

    public int? LabelOffset { get; }

    public int? TickLabelSkip { get; }

    public string TickLabelSkipValue { get; }

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
        PptxSceneChartAxisCrossBetween crossBetweenKind,
        string crossBetween,
        PptxSceneChartAxisOrientation orientationKind,
        string orientation,
        bool? isDeleted,
        bool hasScaling,
        double? minimum,
        string minimumValue,
        double? maximum,
        string maximumValue,
        double? majorUnit,
        string majorUnitValue,
        double? minorUnit,
        bool hasMajorGridlines,
        bool hasMinorGridlines,
        bool hasMajorGridlineElement,
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
        int? labelOffset,
        int? tickLabelSkip,
        string tickLabelSkipValue,
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
        CrossBetweenKind = crossBetweenKind;
        CrossBetween = crossBetween;
        OrientationKind = orientationKind;
        Orientation = orientation;
        IsDeleted = isDeleted;
        HasScaling = hasScaling;
        Minimum = minimum;
        MinimumValue = minimumValue;
        Maximum = maximum;
        MaximumValue = maximumValue;
        MajorUnit = majorUnit;
        MajorUnitValue = majorUnitValue;
        MinorUnit = minorUnit;
        HasMajorGridlines = hasMajorGridlines;
        HasMinorGridlines = hasMinorGridlines;
        HasMajorGridlineElement = hasMajorGridlineElement;
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
        LabelOffset = labelOffset;
        TickLabelSkip = tickLabelSkip;
        TickLabelSkipValue = tickLabelSkipValue;
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
        string crossBetween,
        string orientation,
        bool? isDeleted,
        bool hasScaling,
        double? minimum,
        string minimumValue,
        double? maximum,
        string maximumValue,
        double? majorUnit,
        string majorUnitValue,
        double? minorUnit,
        bool hasMajorGridlines,
        bool hasMinorGridlines,
        bool hasMajorGridlineElement,
        PptxSceneLineStyle line,
        PptxSceneLineStyle majorGridlineLine,
        PptxSceneLineStyle minorGridlineLine,
        PptxSceneLineStyle majorGridlineStyleLine,
        PptxSceneLineStyle minorGridlineStyleLine,
        PptxSceneChartTextStyleOverride textStyle,
        string tickLabelPosition,
        string majorTickMark,
        int? labelOffset,
        int? tickLabelSkip,
        string tickLabelSkipValue,
        string? numberFormat,
        PptxSceneChartNumberFormat numberFormatInfo,
        PptxSceneChartTitle title,
        bool hasLogarithmicScale = false)
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
            PptxSceneBuilder.ParseChartAxisCrossBetween(crossBetween),
            crossBetween,
            PptxSceneBuilder.ParseChartAxisOrientation(orientation),
            orientation,
            isDeleted,
            hasScaling,
            minimum,
            minimumValue,
            maximum,
            maximumValue,
            majorUnit,
            majorUnitValue,
            minorUnit,
            hasMajorGridlines,
            hasMinorGridlines,
            hasMajorGridlineElement,
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
            labelOffset,
            tickLabelSkip,
            tickLabelSkipValue,
            numberFormat,
            numberFormatInfo,
            title) { HasLogarithmicScale = hasLogarithmicScale };
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
    PptxSceneChartTextStyleOverride TextStyle)
{
    public bool HasMultipleTextLines => TextRuns.Count != 0 && TextRuns[^1].LineIndex > 0;
}

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
    PptxSceneChartTextStyleOverride TextStyle,
    IReadOnlySet<int>? DeletedEntryIndices = null);

internal enum PptxSceneChartLegendPosition
{
    Bottom,
    Left,
    Right,
    Top,
    TopRight,
    Unknown
}
