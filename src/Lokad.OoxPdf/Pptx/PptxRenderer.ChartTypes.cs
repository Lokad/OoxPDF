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

    private readonly record struct ChartRadarSeries(IReadOnlyList<ChartIndexedNumberPoint?> Points, ChartIndexedNumberVector Source);

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

    private readonly record struct ChartIndexedNumberVector(
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
            int pointCount = Math.Max(PointCount ?? 0, InferPointCount(points) ?? 0);
            if (pointCount <= 0)
            {
                return [];
            }

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

        public IReadOnlyList<ChartIndexedNumberPoint> WorkbookPointsForPlotVisibility(bool plotVisibleOnly)
        {
            IReadOnlyList<ChartIndexedNumberPoint> points = WorkbookPoints ?? [];
            return plotVisibleOnly
                ? points.Where(point => IsWorkbookPointVisible(point.WorkbookCell)).ToArray()
                : points;
        }

        public ChartIndexedNumberPoint? WorkbookPointForIndex(int index)
        {
            foreach (ChartIndexedNumberPoint point in WorkbookPointsForPlotVisibility(PlotVisibleOnly))
            {
                if (point.Index == index)
                {
                    return point;
                }
            }

            return null;
        }
    }

    private readonly record struct ChartIndexedNumberPoint(
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
            int pointCount = Math.Max(PointCount ?? 0, InferPointCount(points) ?? 0);
            if (pointCount <= 0)
            {
                return [];
            }

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

    private readonly record struct ChartIndexedTextPoint(
        int Index,
        ChartPointIndexSource IndexSource,
        string Text,
        bool HasText,
        ChartWorkbookRangeCell WorkbookCell);

    private enum ChartPointIndexSource
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

    private static ChartSeriesStroke ChartAxisDefaultStroke { get; } = new(new RgbColor(90, 90, 90), 1d, 0.75d);

    private static ChartSeriesStroke ChartNegativeBarDefaultStroke { get; } = new(new RgbColor(0, 0, 0), 1d, 0.75d);

    private static ChartSeriesStroke ChartDataLabelLeaderLineDefaultStroke { get; } = new(new RgbColor(89, 89, 89), 1d, 0.75d);

    private static ChartSeriesStroke RadarGridlineDefaultStroke { get; } = new(new RgbColor(134, 134, 134), 1d, 0.75d, null, 0, 1, null);

    private static void DrawLineChartCategoryAxisMajorTicks(PdfGraphicsBuilder graphics, double plotX, double plotWidth, int pointCount, double axisY, PptxSceneChartAxisTickMark majorTickMark)
    {
        if (pointCount <= 0 || majorTickMark == PptxSceneChartAxisTickMark.None)
        {
            return;
        }

        double outward = majorTickMark == PptxSceneChartAxisTickMark.Cross
            ? PptxChartMetricRules.CategoryAxisMajorTickLength / 2d
            : PptxChartMetricRules.CategoryAxisMajorTickLength;
        double inward = majorTickMark == PptxSceneChartAxisTickMark.Inside || majorTickMark == PptxSceneChartAxisTickMark.Cross
            ? PptxChartMetricRules.CategoryAxisMajorTickLength / 2d
            : 0d;
        double slotWidth = plotWidth / pointCount;
        for (int i = 0; i <= pointCount; i++)
        {
            double x = plotX + slotWidth * i;
            graphics.StrokeLine(x, axisY - outward, x, axisY + inward);
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

    private readonly record struct ChartValueExtents(double Min, double Max);

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

    private enum ChartSeriesNameSource
    {
        Default,
        Cache,
        Workbook
    }

    private readonly record struct ChartSeriesNameRecord(
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
