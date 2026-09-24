using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

// Chart layout and style model. Split from the chart renderer file:
// plot/box/axis/legend/label option records and shared series primitives.
internal sealed partial class PptxRenderer
{
    private readonly record struct ScatterSeries(IReadOnlyList<ScatterPoint> Points, ChartIndexedScatterSeries Source);

    private readonly record struct ChartRadarSeries(IReadOnlyList<double?> Values, ChartIndexedNumberVector Source);

    private readonly record struct ScatterPoint(
        double X,
        double Y,
        double Size,
        int Index,
        ChartIndexedNumberPoint XPoint,
        ChartIndexedNumberPoint YPoint,
        ChartIndexedNumberPoint? BubbleSizePoint,
        ChartIndexedNumberPoint? XWorkbookPoint,
        ChartIndexedNumberPoint? YWorkbookPoint,
        ChartIndexedNumberPoint? BubbleSizeWorkbookPoint,
        string? YFormatCode,
        string? BubbleSizeFormatCode);

    private readonly record struct ChartIndexedScatterSeries(
        ChartIndexedNumberVector XValues,
        ChartIndexedNumberVector YValues,
        ChartIndexedNumberVector BubbleSizes,
        bool ReadBubbleSize);

    private readonly record struct ChartIndexedPieSlice(
        int Index,
        double Value,
        ChartIndexedNumberPoint Point,
        ChartIndexedNumberPoint? WorkbookPoint);

    // R16: internal so the per-frame shared series-vector memo carries the vector type.
    internal readonly record struct ChartIndexedNumberVector(
        IReadOnlyList<ChartIndexedNumberPoint> Points,
        int? PointCount,
        string? Formula,
        string? FormatCode,
        PptxSceneChartDataSource Source,
        IReadOnlyList<ChartIndexedNumberPoint> WorkbookPoints,
        bool PlotVisibleOnly)
    {
        public IReadOnlyList<ChartIndexedNumberPoint?> DensePoints()
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = Points ?? [];
            int pointCount = ResolveDensePointCount(PointCount, points);
            if (pointCount <= 0)
            {
                return [];
            }

            OoxConversionBudget.Current?.ChargeChartRangeCells(pointCount);
            var values = new ChartIndexedNumberPoint?[pointCount];
            foreach (ChartIndexedNumberPoint point in points)
            {
                if (point.Index >= 0 && point.Index < pointCount && point.Value is not null)
                {
                    values[point.Index] = point;
                }
            }

            return values;
        }

        // RV20: compact dense value/index/validity view. Carries the same slot
        // values as DensePoints (same resolution, charge and validity
        // predicate, last-wins duplicates) at one nullable double per slot
        // instead of one fat provenance struct, for value-only consumers.
        // Provenance stays once in Points; use GetDensePoint for lossless
        // per-slot provenance.
        public double?[] DenseValues()
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = Points ?? [];
            int pointCount = ResolveDensePointCount(PointCount, points);
            if (pointCount <= 0)
            {
                return [];
            }

            OoxConversionBudget.Current?.ChargeChartRangeCells(pointCount);
            var values = new double?[pointCount];
            foreach (ChartIndexedNumberPoint point in points)
            {
                if (point.Index >= 0 && point.Index < pointCount && point.Value is not null)
                {
                    values[point.Index] = point.Value;
                }
            }

            return values;
        }

        // RV20: lossless per-slot provenance lookup over the sparse source:
        // returns exactly what DensePoints would hold at index (last-wins
        // duplicates, holes for missing/null/negative/out-of-range), without
        // materializing the dense array. Same resolution and caps as
        // DensePoints; lookups allocate nothing proportional and charge nothing.
        public ChartIndexedNumberPoint? GetDensePoint(int index)
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = Points ?? [];
            int pointCount = ResolveDensePointCount(PointCount, points);
            if (index < 0 || index >= pointCount)
            {
                return null;
            }

            ChartIndexedNumberPoint? found = null;
            foreach (ChartIndexedNumberPoint point in points)
            {
                if (point.Index == index && point.Value is not null)
                {
                    found = point;
                }
            }

            return found;
        }

        public bool HasAnyValue()
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = Points ?? [];
            int pointCount = ResolveDensePointCount(PointCount, points);
            if (pointCount <= 0)
            {
                return false;
            }

            foreach (ChartIndexedNumberPoint point in points)
            {
                if (point.Index >= 0 && point.Index < pointCount && point.Value is not null)
                {
                    return true;
                }
            }

            return false;
        }

        public bool HasAnyDenseSlot()
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = Points ?? [];
            return ResolveDensePointCount(PointCount, points) > 0;
        }

        // R15: sparse dense length behind the same resolution as DensePoints so
        // count-only passes never materialize (or charge) dense arrays.
        public int DensePointCount()
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = Points ?? [];
            return ResolveDensePointCount(PointCount, points);
        }

        public IReadOnlyList<ChartIndexedNumberPoint> WorkbookPointsForPlotVisibility(bool plotVisibleOnly)
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = WorkbookPoints ?? [];
            return plotVisibleOnly
                ? points.Where(point => IsWorkbookPointVisible(point.WorkbookCell)).ToArray()
                : points;
        }

    }

    internal const int MaxChartDensePoints = 100_000;

    private static int ResolveDensePointCount(int? declaredCount, IReadOnlyList<ChartIndexedNumberPoint> points)
    {
        int inferred = InferCheckedPointCount(points);
        int declared = declaredCount ?? 0;
        if (declared < 0)
        {
            throw new OoxPdfLimitExceededException("Chart point count is negative.");
        }

        if (declared > MaxChartDensePoints)
        {
            throw new OoxPdfLimitExceededException(
                "Chart point count exceeds the maximum supported point count of " + MaxChartDensePoints + ".");
        }

        long combined = Math.Max((long)declared, inferred);
        if (combined > MaxChartDensePoints)
        {
            throw new OoxPdfLimitExceededException(
                "Chart point index exceeds the maximum supported point count of " + MaxChartDensePoints + ".");
        }

        return (int)combined;
    }

    private static int ResolveDensePointCount(int? declaredCount, IReadOnlyList<ChartIndexedTextPoint> points)
    {
        int inferred = InferCheckedPointCount(points);
        int declared = declaredCount ?? 0;
        if (declared < 0)
        {
            throw new OoxPdfLimitExceededException("Chart point count is negative.");
        }

        if (declared > MaxChartDensePoints)
        {
            throw new OoxPdfLimitExceededException(
                "Chart point count exceeds the maximum supported point count of " + MaxChartDensePoints + ".");
        }

        long combined = Math.Max((long)declared, inferred);
        if (combined > MaxChartDensePoints)
        {
            throw new OoxPdfLimitExceededException(
                "Chart point index exceeds the maximum supported point count of " + MaxChartDensePoints + ".");
        }

        return (int)combined;
    }

    private static int InferCheckedPointCount(IReadOnlyList<ChartIndexedNumberPoint> points)
    {
        int max = -1;
        foreach (ChartIndexedNumberPoint point in points)
        {
            if (point.Index < 0)
            {
                continue;
            }

            // Checked index+1: a sparse index of int.MaxValue previously overflowed
            // to an empty dense result (silent loss). Reject it instead (M02).
            long candidate;
            try
            {
                candidate = checked((long)point.Index + 1);
            }
            catch (OverflowException ex)
            {
                throw new OoxPdfLimitExceededException("Chart point index overflows.", ex);
            }

            if (candidate > MaxChartDensePoints)
            {
                throw new OoxPdfLimitExceededException(
                    "Chart point index exceeds the maximum supported point count of " + MaxChartDensePoints + ".");
            }

            max = Math.Max(max, (int)(candidate - 1));
        }

        return max < 0 ? 0 : checked(max + 1);
    }

    private static int InferCheckedPointCount(IReadOnlyList<ChartIndexedTextPoint> points)
    {
        int max = -1;
        foreach (ChartIndexedTextPoint point in points)
        {
            if (point.Index < 0)
            {
                continue;
            }

            long candidate;
            try
            {
                candidate = checked((long)point.Index + 1);
            }
            catch (OverflowException ex)
            {
                throw new OoxPdfLimitExceededException("Chart point index overflows.", ex);
            }

            if (candidate > MaxChartDensePoints)
            {
                throw new OoxPdfLimitExceededException(
                    "Chart point index exceeds the maximum supported point count of " + MaxChartDensePoints + ".");
            }

            max = Math.Max(max, (int)(candidate - 1));
        }

        return max < 0 ? 0 : checked(max + 1);
    }

    internal readonly record struct ChartIndexedNumberPoint(
        int Index,
        ChartPointIndexSource IndexSource,
        double? Value,
        string Text,
        bool HasValue,
        ChartWorkbookRangeCell WorkbookCell);

    private readonly record struct ChartIndexedTextVector(
        IReadOnlyList<ChartIndexedTextPoint> Points,
        int? PointCount,
        IReadOnlyList<IReadOnlyList<ChartIndexedTextPoint>> Levels,
        string? Formula,
        PptxSceneChartDataSource Source,
        IReadOnlyList<ChartIndexedTextPoint> WorkbookPoints,
        bool PlotVisibleOnly)
    {
        public IReadOnlyList<ChartIndexedTextPoint?> DensePoints()
        {
            IReadOnlyList<ChartIndexedTextPoint> points = Points ?? [];
            int pointCount = ResolveDensePointCount(PointCount, points);
            if (pointCount <= 0)
            {
                return [];
            }

            OoxConversionBudget.Current?.ChargeChartRangeCells(pointCount);
            var values = new ChartIndexedTextPoint?[pointCount];
            foreach (ChartIndexedTextPoint point in points)
            {
                if (point.Index >= 0 && point.Index < pointCount && point.HasText)
                {
                    values[point.Index] = point;
                }
            }

            return values;
        }

        public bool HasAnyText()
        {
            IReadOnlyList<ChartIndexedTextPoint> points = Points ?? [];
            int pointCount = ResolveDensePointCount(PointCount, points);
            if (pointCount <= 0)
            {
                return false;
            }

            foreach (ChartIndexedTextPoint point in points)
            {
                if (point.Index >= 0 && point.Index < pointCount && point.HasText)
                {
                    return true;
                }
            }

            return false;
        }

        public bool HasAnyDenseSlot()
        {
            IReadOnlyList<ChartIndexedTextPoint> points = Points ?? [];
            return ResolveDensePointCount(PointCount, points) > 0;
        }

        // R15: sparse existence check for label rendering probes. Matches the dense
        // scan exactly (in-range indexes, text-bearing points, non-whitespace text)
        // without materializing the slot array.
        public bool HasAnyVisibleText()
        {
            IReadOnlyList<ChartIndexedTextPoint> points = Points ?? [];
            int pointCount = ResolveDensePointCount(PointCount, points);
            if (pointCount <= 0)
            {
                return false;
            }

            foreach (ChartIndexedTextPoint point in points)
            {
                if (point.Index >= 0 && point.Index < pointCount && point.HasText && !string.IsNullOrWhiteSpace(point.Text))
                {
                    return true;
                }
            }

            return false;
        }

        public IReadOnlyList<ChartIndexedTextPoint> WorkbookPointsForPlotVisibility(bool plotVisibleOnly)
        {
            IReadOnlyList<ChartIndexedTextPoint> points = WorkbookPoints ?? [];
            return plotVisibleOnly
                ? points.Where(point => IsWorkbookPointVisible(point.WorkbookCell)).ToArray()
                : points;
        }
    }

    private static bool IsWorkbookPointVisible(ChartWorkbookRangeCell cell)
    {
        return !cell.RowHidden && !cell.ColumnHidden;
    }

    // R16: internal so the per-frame shared dense label memo carries the point type.
    internal readonly record struct ChartIndexedTextPoint(
        int Index,
        ChartPointIndexSource IndexSource,
        string Text,
        bool HasText,
        ChartWorkbookRangeCell WorkbookCell);

    internal enum ChartPointIndexSource
    {
        OrdinalFallback,
        OoxmlIndex,
        WorkbookRange
    }

    private readonly record struct ChartSeriesFill(RgbColor Color, double Alpha, string? PatternPreset, RgbColor? BackgroundColor);

    private readonly record struct ChartRectangle(double X, double Y, double Width, double Height);

    private readonly record struct ChartStackedBarSegment(int SeriesIndex, int CategoryIndex, double X, double Y, double Width, double Height, double Value);

    private readonly record struct ChartStackedBarFillRun(ChartSeriesFill Fill, List<ChartRectangle> Rectangles);

    private readonly record struct ChartSeriesStroke(
        RgbColor Color,
        double Alpha,
        double Width,
        IReadOnlyList<double>? DashPattern,
        int? Cap,
        int? Join,
        PptxSceneLineCompound? Compound)
    {
        public ChartSeriesStroke(RgbColor color, double alpha, double width)
            : this(color, alpha, width, null, null, null, null)
        {
        }
    }

    // Default axis/tick stroke: Office draws unstyled axes and ticks black across
    // bar, column, line, scatter, and area refs (0.75 and 1.0 widths); explicitly
    // styled axes keep their colors.
    private static ChartSeriesStroke ChartAxisDefaultStroke { get; } = new(new RgbColor(0, 0, 0), 1d, 0.75d);

    private static ChartSeriesStroke ChartNegativeBarDefaultStroke { get; } = new(new RgbColor(0, 0, 0), 1d, 0.75d);

    private static ChartSeriesStroke ChartDataLabelLeaderLineDefaultStroke { get; } = new(new RgbColor(89, 89, 89), 1d, 0.75d);

    private static ChartSeriesStroke RadarGridlineDefaultStroke { get; } = new(new RgbColor(134, 134, 134), 1d, 0.75d, null, 0, 1, null);

    private static void DrawLineChartCategoryAxisMajorTicks(PdfGraphicsBuilder graphics, double plotX, double plotWidth, int pointCount, double axisY, PptxSceneChartAxisTickMark majorTickMark, double tickFontSize)
    {
        if (pointCount <= 0)
        {
            return;
        }

        double slotWidth = plotWidth / pointCount;
        double[] edges = new double[pointCount + 1];
        for (int i = 0; i <= pointCount; i++)
        {
            edges[i] = plotX + slotWidth * i;
        }

        StrokeMajorTickSegments(graphics, edges, axisY, verticalSegments: true, outwardIsLowSide: true, majorTickMark, tickFontSize * PptxChartMetricRules.ChartAxisMajorTickLengthFactor);
    }

    // Shared cartesian major-tick segments for category slot boundaries and value
    // gridlines: Office honors out/inside/cross on every cartesian axis with the same
    // font-relative length. outwardIsLowSide selects the out side for bottom/left
    // axes (top/right mirrors are definitional).
    private static void StrokeMajorTickSegments(PdfGraphicsBuilder graphics, IReadOnlyList<double> edgePositions, double axisCoordinate, bool verticalSegments, bool outwardIsLowSide, PptxSceneChartAxisTickMark tickMark, double tickLength)
    {
        if (tickMark == PptxSceneChartAxisTickMark.None || edgePositions.Count == 0 || tickLength <= 0d)
        {
            return;
        }

        double outward = tickMark == PptxSceneChartAxisTickMark.Cross ? tickLength / 2d : tickLength;
        double inward = tickMark == PptxSceneChartAxisTickMark.Inside || tickMark == PptxSceneChartAxisTickMark.Cross ? tickLength / 2d : 0d;
        double direction = outwardIsLowSide ? -1d : 1d;
        foreach (double position in edgePositions)
        {
            if (verticalSegments)
            {
                graphics.StrokeLine(position, axisCoordinate + direction * outward, position, axisCoordinate - direction * inward);
            }
            else
            {
                graphics.StrokeLine(axisCoordinate + direction * outward, position, axisCoordinate - direction * inward, position);
            }
        }
    }

    private readonly record struct ChartAxesStyle(ChartSeriesStroke? ValueAxis, ChartSeriesStroke? SecondaryValueAxis, ChartSeriesStroke? CategoryAxis, bool ValueAxisRightSide, bool SecondaryValueAxisRightSide, bool ValueAxisBottomSide, bool CategoryAxisRightSide, bool ValueAxisVisible, bool CategoryAxisVisible, PptxSceneChartAxisTickMark CategoryAxisMajorTickMark, bool CategoryAxisTopSide);

    private readonly record struct ChartGridlineStyle(ChartSeriesStroke? Major, ChartSeriesStroke? Minor)
    {
        public static ChartGridlineStyle Empty { get; } = new(null, null);
    }

    private readonly record struct ChartValueAxisRenderOptions(ChartAxisUnits Units, bool Reversed, double? CrossingValue, bool MajorGridlines, bool MinorGridlines, ChartGridlineStyle GridlineStyle);

    private readonly record struct ChartBubbleValueAxisOptions(ChartAxisUnits Units, ChartGridlineStyle GridlineStyle);

    private readonly record struct ChartLayout(ChartFrameBox Frame, ChartLayoutBox PlotAreaBox, ChartPlotBox PlotBox, bool ManualPlotLayoutApplied, string? Title, PptxSceneChartTextBodyProperties TitleTextBodyProperties, ChartLegendLayout Legend);

    private readonly record struct ChartFrameBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartLayoutBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartPlotBox(double X, double Y, double Width, double Height);

    private readonly record struct ChartLegendBox(
        double X,
        double ClipY,
        double Width,
        double ClipHeight,
        double FirstY,
        double LineHeight,
        double MarkerSize,
        double MarkerWidth,
        double TextGap,
        bool Horizontal,
        bool SideStrokeLegend);

    private readonly record struct ChartRightLegendReserve(
        double Width,
        double FontSize,
        double MaxTextWidth,
        int MaxTextLength,
        bool IncludesAreaReserve);

    private readonly record struct ChartAxisSource(PptxSceneChartAxis? SceneAxis, XElement? XmlAxis);

    private readonly record struct ChartPlotLayout(ChartLayoutBox PlotAreaBox, ChartPlotBox PlotBox, PptxSceneChartManualLayoutTarget? ManualLayoutTargetKind)
    {
        public static ChartPlotLayout FromPlotBox(ChartPlotBox plotBox)
        {
            return new ChartPlotLayout(new ChartLayoutBox(plotBox.X, plotBox.Y, plotBox.Width, plotBox.Height), plotBox, null);
        }
    }

    private readonly record struct ChartAxisTitleReserveSides(bool Left, bool Right, bool Bottom, bool Top)
    {
        public bool HasHorizontalTitle => Bottom || Top;

        public bool HasVerticalTitle => Left || Right;

        public ChartAxisTitleReserveSides With(PptxSceneChartAxisPosition positionKind)
        {
            return positionKind switch
            {
                PptxSceneChartAxisPosition.Left => this with { Left = true },
                PptxSceneChartAxisPosition.Right => this with { Right = true },
                PptxSceneChartAxisPosition.Bottom => this with { Bottom = true },
                PptxSceneChartAxisPosition.Top => this with { Top = true },
                _ => this
            };
        }
    }

    // R15: internal so the per-frame shared extent memo carries the type.
    internal readonly record struct ChartValueExtents(double Min, double Max);

    private enum ChartAxisScalingBound
    {
        Minimum,
        Maximum
    }

    private readonly record struct ChartAxisUnits(double? MajorUnit, double? MinorUnit)
    {
        public static ChartAxisUnits Empty { get; } = new(null, null);
    }

    private readonly record struct ChartTextStyle(
        string? FontFamily,
        double FontSize,
        double CharacterSpacing,
        RgbColor Color,
        double Alpha,
        bool Bold,
        bool Italic,
        bool Underline,
        bool Strike,
        string? RequestedTypeface,
        PptxThemeTypefaceSource? TypefaceSource)
    {
        public ChartTextStyle Merge(ChartTextStyleOverride next)
        {
            return new ChartTextStyle(
                next.FontFamily ?? FontFamily,
                next.FontSize ?? FontSize,
                next.CharacterSpacing ?? CharacterSpacing,
                next.Color ?? Color,
                next.Alpha ?? Alpha,
                next.Bold ?? Bold,
                next.Italic ?? Italic,
                next.Underline ?? Underline,
                next.Strike ?? Strike,
                next.FontFamily is null ? RequestedTypeface : next.RequestedTypeface,
                next.FontFamily is null ? TypefaceSource : next.TypefaceSource);
        }
    }

    private readonly record struct ChartTextStyleOverride(
        string? FontFamily,
        double? FontSize,
        double? CharacterSpacing,
        RgbColor? Color,
        double? Alpha,
        bool? Bold,
        bool? Italic,
        bool? Underline,
        bool? Strike,
        string? RequestedTypeface,
        PptxThemeTypefaceSource? TypefaceSource)
    {
        public static ChartTextStyleOverride Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null);
    }

    private static IReadOnlyDictionary<int, ChartDataLabelOverride> EmptyChartDataLabelOverrides { get; } = new Dictionary<int, ChartDataLabelOverride>();

    private static IReadOnlyDictionary<string, ChartBooleanOption> EmptyChartDataLabelFlagOptions { get; } = new Dictionary<string, ChartBooleanOption>();

    private static readonly string[] ChartDataLabelFlagNames =
    [
        "showVal",
        "showPercent",
        "showCatName",
        "showSerName",
        "showLeaderLines",
        "showLegendKey",
        "showBubbleSize"
    ];

    private readonly record struct ChartDataLabelOptions(bool ShowValue, bool ShowPercent, bool ShowCategoryName, bool ShowSeriesName, bool ShowLeaderLines, bool ShowLegendKey, bool ShowBubbleSize, ChartDataLabelLeaderLines LeaderLines, string CustomText, IReadOnlyList<ChartTextRunOverride> CustomTextRuns, PptxSceneChartDataLabelPosition PositionKind, string Position, string Separator, string NumberFormat, ChartNumberFormat NumberFormatInfo, PptxSceneChartManualLayout Layout, ChartTextStyleOverride TextStyle, PptxSceneChartTextBodyProperties TextBodyProperties, ChartShapeStyle ShapeStyle, IReadOnlyDictionary<string, ChartBooleanOption> FlagOptions, IReadOnlyDictionary<int, ChartDataLabelOverride> Overrides, bool IsDefined, bool Date1904 = false)
    {
        public static ChartDataLabelOptions None { get; } = new(ShowValue: false, ShowPercent: false, ShowCategoryName: false, ShowSeriesName: false, ShowLeaderLines: false, ShowLegendKey: false, ShowBubbleSize: false, LeaderLines: ChartDataLabelLeaderLines.Empty, CustomText: string.Empty, CustomTextRuns: [], PositionKind: PptxSceneChartDataLabelPosition.Unknown, Position: string.Empty, Separator: string.Empty, NumberFormat: string.Empty, NumberFormatInfo: default, Layout: default, TextStyle: ChartTextStyleOverride.Empty, TextBodyProperties: default, ShapeStyle: ChartShapeStyle.Empty, FlagOptions: EmptyChartDataLabelFlagOptions, Overrides: EmptyChartDataLabelOverrides, IsDefined: false);

        public bool HasVisibleText => ShowValue || ShowPercent || ShowCategoryName || ShowSeriesName || ShowBubbleSize ||
            !string.IsNullOrWhiteSpace(CustomText) ||
            Overrides.Values.Any(label => label.ShowValue == true || label.ShowPercent == true || label.ShowCategoryName == true || label.ShowSeriesName == true || label.ShowBubbleSize == true || !string.IsNullOrWhiteSpace(label.CustomText));

        public bool HasVisibleContent => HasVisibleText || ShowLegendKey ||
            Overrides.Values.Any(label => label.ShowLegendKey == true);
    }

    private readonly record struct ChartDataLabelLeaderLines(bool IsDefined, ChartSeriesStroke? Stroke)
    {
        public static ChartDataLabelLeaderLines Empty { get; } = new(IsDefined: false, Stroke: null);
    }

    private readonly record struct ChartTextRunOverride(string Text, ChartTextStyleOverride TextStyle, string? HyperlinkClickId = null, string? HyperlinkClickAction = null);

    private readonly record struct ChartTextRunLink(TextRun Run, string? RelationshipId, string? Action);

    private readonly record struct ChartTextRunLayout(string Text, ChartTextStyle Style, double Width, string? HyperlinkClickId = null, string? HyperlinkClickAction = null);

    private readonly record struct ChartDataLabelOverride(bool? IsDeleted, bool? ShowValue, bool? ShowPercent, bool? ShowCategoryName, bool? ShowSeriesName, bool? ShowLeaderLines, bool? ShowLegendKey, bool? ShowBubbleSize, ChartDataLabelLeaderLines LeaderLines, string CustomText, IReadOnlyList<ChartTextRunOverride> CustomTextRuns, PptxSceneChartDataLabelPosition PositionKind, string Position, string Separator, string NumberFormat, ChartNumberFormat NumberFormatInfo, PptxSceneChartManualLayout Layout, ChartTextStyleOverride TextStyle, PptxSceneChartTextBodyProperties TextBodyProperties, ChartShapeStyle ShapeStyle, IReadOnlyDictionary<string, ChartBooleanOption> FlagOptions);

    private readonly record struct ChartNumberFormat(bool IsDefined, string FormatCode, bool? SourceLinked, string SourceLinkedValue);

    internal enum ChartSeriesNameSource
    {
        Default,
        Cache,
        Workbook
    }

    // R16: internal so the per-frame shared series-name memo carries the record type.
    internal readonly record struct ChartSeriesNameRecord(
        string ActiveName,
        string CacheName,
        ChartSeriesNameSource ActiveNameSource,
        PptxSceneChartDataSource Source,
        IReadOnlyList<ChartIndexedTextPoint> WorkbookPoints);

    private readonly record struct ChartLegendEntry(string Name, ChartSeriesFill? Fill, ChartSeriesStroke? Stroke, ChartMarkerStyle? Marker, ChartSeriesNameRecord? SeriesName, bool LineHidden);

    private enum ChartLegendPlacement
    {
        Default,
        BubbleTitleRightLegend,
        AreaRightLegend
    }

    private readonly record struct ChartBooleanOption(bool Value, string RawValue, bool IsDefined);

    private readonly record struct ChartLegendLayout(PptxSceneChartLegendPosition PositionKind, string Position, bool Overlay, bool Visible, PptxSceneChartManualLayout Layout, PptxSceneChartTextBodyProperties TextBodyProperties, ChartShapeStyle ShapeStyle)
    {
        public static ChartLegendLayout Hidden { get; } = new(PptxSceneChartLegendPosition.Right, "r", Overlay: false, Visible: false, default, default, ChartShapeStyle.Empty);
    }

    private readonly record struct ChartShapeStyle(ChartSeriesFill? Fill, GradientFill? GradientFill, ChartSeriesStroke? Stroke, PptxSceneGlow Glow, PptxSceneOuterShadow OuterShadow)
    {
        public static ChartShapeStyle Empty { get; } = new(null, null, null, default, default);

        public bool IsEmpty => Fill is null && GradientFill is null && Stroke is null && !Glow.HasGlow && !OuterShadow.HasShadow;
    }

    private readonly record struct ChartMarkerStyle(PptxSceneChartMarkerSymbol SymbolKind, string Symbol, string? SizeValue, double Size, ChartSeriesFill? Fill, ChartSeriesStroke? Stroke, bool IsDefined)
    {
        public static ChartMarkerStyle Default { get; } = new(PptxSceneChartMarkerSymbol.Circle, "circle", null, PptxChartMarkerMetricRules.DefaultChartMarkerSize, null, null, false);
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
        double CategoryHorizontalGapSideFactor,
        double CategoryVerticalGapSideFactor,
        double CategoryVerticalGapFontFactor,
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
}
