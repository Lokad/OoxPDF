using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static class PptxTextMetricRules
    {
        public const double CoordinateTolerance = 0.01d;
        public const double TextStateTolerance = 0.001d;
        public const double MinimumDrawableDimension = 1d;
        public const double MinimumStrokeWidth = 0.5d;
        public const double OfficeSuperscriptSubscriptScale = 2d / 3d;
        public const double CssNormalLineHeightFallback = 1.2d;
        public const double OfficeCompatibleDefaultLineSpacingFactor = 1.1d;
        public const double OfficeCompatibleNoAutoFitDefaultLineSpacingFactor = 1.2d;
        public const double MiddleVerticalAnchorSlackMultiplier = 0.5d;
        public const double OfficeManualBreakDefaultLineHeightFallback = 1.24d;
        public const double OfficeManualBreakBaselineFallback = 0.9344d;
        public const double OfficeBaselineFallback = 0.974d;
        public const double MinimumBaselineMetricRatio = 0.75d;
        public const double MaximumBaselineMetricRatio = 1.05d;
        public const double MaximumOfficeBaselineWindowsAscenderRatio = 1d;
        public const double MinimumOfficeTypographicBaselineAscenderRatio = 0.93d;
        public const double MaximumOfficeTypographicBaselineFontSize = 20d;
        public const double OfficeBaselineFloorMetricThreshold = 0.94d;
        public const double OfficeBaselineFloorMaximumWindowsDescenderRatio = 0.24d;
        public const double MinimumFontLineBoxMetricRatio = 0.75d;
        public const double MaximumFontLineBoxMetricRatio = 1.5d;
        public const double MaximumTableAnchorCompressedFontBoxRatio = 1.18d;
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
        public const double OfficeSyntheticBoldAdvanceTighteningEm = 0.007d;
        public const double OfficeSyntheticBoldItalicCharacterSpacingEm = 0.01545d;
        public const double OfficeHighlightContinuationCharacterSpacingEm = -0.003d;
        public const double OfficeAutofitNumberedDenseCharacterSpacingEm = -0.002d;
        public const double OfficeAutofitNumberedDefaultCharacterSpacingEm = -0.004d;
        public const double OfficeStrikePositionFontScale = 0.211d;
        public const double StrikeThicknessFallback = 0.05d;
        public const double UnderlineThicknessFallback = 0.0125d;
        public const double OfficeUnderlineThicknessMetricScale = 0.25d;
        public const double HighlightDescenderPaddingFontUnits = 32d;
        public const double HighlightMaximumDescentFontScale = 0.23d;
        public const double HighlightMaximumHeightFontScale = 1.18d;
        public const double AdjacentTextCoalesceGapFontScale = 0.2d;
        public const double AdjacentUnderlineCoalesceGapFontScale = 0.08d;
        public const double WrapFitToleranceFontScale = 0.16d;
        public const double BulletWrapFitToleranceFontScale = 0.2d;
        public const double CenteredTableCellWrapToleranceWidthScale = 0.02d;
        public const double FinalWordWrapToleranceWidthScale = 0.02d;
        public const double ShapeAutoFitWrapToleranceWidthScale = 0.011d;
        public const double FallbackAdvanceWidthScale = 0.42d;
        public const double EllipseTextRectInsetRatio = 0.1464466094067262d;
        public const double RoundRectDefaultAdjustment = 16667d;
        public const double RoundRectTextRectRadiusInsetFactor = 0.2928932188134525d;
        public const int ShapeAutoFitSearchIterations = 10;

        public static double ClampNonNegative(double value) => Math.Max(0d, value);

        public static double MinimumWidth(double width) => Math.Max(MinimumDrawableDimension, width);

        public static double SuperscriptSubscriptFontSize(double nominalFontSize) => nominalFontSize * OfficeSuperscriptSubscriptScale;

        public static bool ShouldScaleSuperscriptSubscript(double baselineOffset, double nominalFontSize) =>
            Math.Abs(baselineOffset) >= nominalFontSize * SuperscriptSubscriptMinimumBaselineRatio;

        public static double SmallCapsFontScale() => SmallCapsFallbackScale;

        public static double TextOutlineWidth(double? width) => Math.Max(MinimumStrokeWidth, width ?? DefaultTextOutlineWidth);

        public static double SyntheticBoldStrokeWidth(double fontSize) => Math.Max(0d, fontSize * SyntheticBoldStrokeWidthRatio);

        public static double OfficeSyntheticBoldAdvanceTightening(double fontSize) =>
            Math.Max(0d, fontSize * OfficeSyntheticBoldAdvanceTighteningEm);

        public static double OfficeSyntheticBoldItalicCharacterSpacing(double fontSize) =>
            Math.Max(0d, fontSize * OfficeSyntheticBoldItalicCharacterSpacingEm);

        public static double OfficeHighlightContinuationCharacterSpacing(double fontSize) =>
            fontSize * OfficeHighlightContinuationCharacterSpacingEm;

        public static double OfficeAutofitNumberedDenseCharacterSpacing(double fontSize) =>
            fontSize * OfficeAutofitNumberedDenseCharacterSpacingEm;

        public static double OfficeAutofitNumberedDefaultCharacterSpacing(double fontSize) =>
            fontSize * OfficeAutofitNumberedDefaultCharacterSpacingEm;

        public static double StrikeY(PdfEmbeddedFont embedded, double baselineY, double fontSize)
        {
            return baselineY + fontSize * OfficeStrikePositionFontScale;
        }

        public static double StrikeThickness(PdfEmbeddedFont embedded, double fontSize)
        {
            double fontScale = fontSize / embedded.Font.UnitsPerEm;
            double strikeoutSize = embedded.Font.Os2.StrikeoutSize * fontScale;
            return Math.Max(MinimumStrokeWidth, strikeoutSize > 0d ? strikeoutSize : fontSize * StrikeThicknessFallback);
        }

        public static double UnderlineThickness(PdfEmbeddedFont embedded, double fontSize)
        {
            double fontScale = fontSize / embedded.Font.UnitsPerEm;
            double underlineSize = Math.Abs(embedded.Font.Post.UnderlineThickness) * fontScale;
            return underlineSize > TextStateTolerance
                ? underlineSize * OfficeUnderlineThicknessMetricScale
                : fontSize * UnderlineThicknessFallback;
        }

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

        public static double BulletWrapFitTolerance(double fontSize) => Math.Max(WrapFitTolerance(fontSize), fontSize * BulletWrapFitToleranceFontScale);

        public static double CenteredTableCellWrapTolerance(double fontSize, double availableWidth) => Math.Max(WrapFitTolerance(fontSize), availableWidth * CenteredTableCellWrapToleranceWidthScale);

        public static double FinalWordWrapTolerance(double fontSize, double availableWidth) => Math.Max(WrapFitTolerance(fontSize), availableWidth * FinalWordWrapToleranceWidthScale);

        public static double ShapeAutoFitWrapTolerance(double fontSize, double availableWidth) => Math.Max(WrapFitTolerance(fontSize), availableWidth * ShapeAutoFitWrapToleranceWidthScale);

        public static double FallbackAdvanceWidth(int codeUnitCount, int runeCount, double fontSize, double characterSpacing)
        {
            return Math.Max(0d, codeUnitCount * fontSize * FallbackAdvanceWidthScale + Math.Max(0, runeCount - 1) * characterSpacing);
        }
    }

    private static class PptxPdfTextEmissionProfile
    {
        public static double FontSize(PptxPdfTextEmissionContext context) => FontSize(context.LayoutFontSize);

        public static double CharacterSpacing(PptxPdfTextEmissionContext context, double layoutCharacterSpacing) => layoutCharacterSpacing;

        public static double FontSize(double layoutFontSize) => OfficePdfTextEmissionProfile.FontSize(layoutFontSize);
    }

    private static class PptxChartMetricRules
    {
        public const double TitleFallbackFontSize = 12d;
        public const double LegendFallbackFontSize = 9d;
        public const double CategoryAxisFallbackFontSize = 9d;
        public const double ValueAxisFallbackFontSize = 8.5d;
        public const double DataLabelFallbackFontSize = 8.5d;
        public const double AxisSingleValueHeadroomFactor = 1.2d;
        public const double AxisValueEpsilon = 0.0001d;
        public const double AxisNiceTickTargetCount = 9d;
        public const double AxisNiceVerticalValueTickTargetCount = 10d;
        public const double AxisNiceHorizontalValueTickTargetCount = 10d;
        public const double BubbleAxisBoundsTickTargetCount = 5d;
        public const double BubbleAxisNiceTickTargetCount = 10d;
        // Scatter-X tick units halve one step earlier than Y (unit 1 through range 8,
        // unit 2 from range 9; cached X-scaling probes over Xmax 8/9/9.5/12). Y keeps 10.
        public const double ScatterXAxisNiceTickTargetCount = 8d;
        public const double AxisNiceTickStepSmall = 1d;
        public const double AxisNiceTickStepMedium = 2d;
        public const double AxisNiceTickStepLarge = 5d;
        public const double AxisNiceTickStepMaximum = 10d;
        public const double AxisNiceNearMaximumHeadroomRatio = 0.96d;
        public const double AreaChartStackedAxisNearMaximumHeadroomRatio = 0.95d;
        // Max bubble radius as a fraction of the smaller plot dimension: Office renders
        // the largest bubble at 0.1306 of min(plot) against 0.128 shipped (bubble port:
        // Office max diameter 89.22 over plot min 341.52; relative sqrt sizing already exact
        // across all four bubbles, so only the scale moves).
        public const double BubbleRadiusPlotRatio = 0.131d;
        // Default radar series outline width. Office draws 3.75pt outlines on unfilled
        // radars without an explicit line (both ladder radar ports share style 118, so
        // revisit if other styles diverge); filled radars omit the outline instead.
        public const double RadarSeriesOutlineWidth = 3.75d;
        // Radar web square: Office bottom-anchors an S-by-S plot square at 32.76pt
        // above the frame bottom with a 33.36pt top reserve untitled (series-clip rects
        // on 432H and 324H Office renders, both styles byte-identical), so the available
        // side is the frame height minus this 66.12pt total; a single-line 21.6pt auto
        // title adds an exact 35.4pt band (two 1-series Office renders). Narrow frames
        // and explicit or multi-line titles are unobserved.
        public const double RadarPlotVerticalReserveTotal = 66.12d;
        public const double RadarTitledPlotBand = 35.4d;
        // Radar web radius stops half a web-graticule width (0.75pt) inside the square.
        public const double RadarWebRadiusPenHalf = 0.375d;
        // Radar category-label gaps: the horizontal gap is 0.0202 times the web-side
        // length (7.39/6.68/5.21pt across the 432H untitled/titled and 324H Office
        // renders, killing the legacy 0.35/0.41 style split); the vertical gap adds
        // 0.28 times the label font size on top (5.04-5.05pt at 18pt on both heights).
        public const double RadarCategoryGapSideFactor = 0.0202d;
        public const double RadarCategoryVerticalGapFontFactor = 0.28d;
        public const double DoughnutHoleMinimumRatio = 0.1d;
        public const double DoughnutHoleMaximumRatio = 0.9d;
        public const double DoughnutHoleFallbackRatio = 0.56d;
        public const double DefaultPlotBoxXRatio = 0.12d;
        public const double DefaultPlotBoxYRatio = 0.16d;
        public const double DefaultPlotBoxWidthRatio = 0.76d;
        public const double DefaultPlotBoxHeightRatio = 0.68d;
        public const double BarDefaultPlotBoxXRatio = 0.1d;
        public const double BarDefaultPlotBoxYRatio = 0.14d;
        public const double BarDefaultPlotBoxWidthRatio = 0.82d;
        public const double BarDefaultPlotBoxHeightRatio = 0.81d;
        public const double BarOverlayOnlyPlotBoxXRatio = 0.0576d;
        public const double BarOverlayOnlyPlotBoxYRatio = 0.0924d;
        public const double BarOverlayOnlyPlotBoxWidthRatio = 0.9272d;
        public const double BarOverlayOnlyPlotBoxHeightRatio = 0.8706d;
        public const double BarNoTitleBottomLegendPlotBoxXRatio = 0.0415d;
        public const double BarNoTitleBottomLegendPlotBoxYRatio = 0.213d;
        public const double BarNoTitleBottomLegendPlotBoxWidthRatio = 0.9406d;
        public const double BarNoTitleBottomLegendPlotBoxHeightRatio = 0.736d;
        // Untitled bottom-legend plot-top anchor: Office holds the plot top fixed
        // under downward frame growth (composite base+tall clips agree bit-identically
        // at 419.76 with frame top 432), while the 0.949H ratio top falls 5.1 per 100H.
        // Single-value inset with invariance proof; max-guarded so exact tops keep legacy.
        public const double BarNoTitleBottomLegendPlotBoxTopInset = 12.24d;
        // Bottom reserve below the legend and tick rows for untitled bottom-legend columns,
        // calibrated as an additive plane over legend and tick font sizes (four same-frame
        // Office probes: 55.02, 47.77, 66.10 and 58.86; all four land within 0.06).
        public const double BarNoTitleBottomLegendReserveLegendFactor = 1.21d;
        public const double BarNoTitleBottomLegendReserveTickFactor = 1.85d;
        public const double BarNoTitleBottomLegendReserveBase = 18.43d;
        public const double StackedColumnBottomLegendPlotBoxRightPadding = 3.9d;
        public const double BarTitleNoLegendPlotBoxXRatio = 0.1106d;
        public const double BarTitleNoLegendPlotBoxYRatio = 0.1008d;
        public const double BarTitleNoLegendPlotBoxWidthRatio = 0.8691d;
        public const double BarTitleNoLegendPlotBoxHeightRatio = 0.7696d;
        public const double BarTitleNoLegendInsideCrossingPlotBoxXRatio = 0.0652d;
        public const double BarTitleNoLegendInsideCrossingPlotBoxYRatio = 0.0370d;
        public const double BarTitleNoLegendInsideCrossingPlotBoxWidthRatio = 0.9195d;
        public const double BarTitleNoLegendInsideCrossingPlotBoxHeightRatio = 0.8441d;
        public const double BarMultiValueAxisPrimaryStripFactor = 1.85d;
        public const double BarMultiValueAxisSecondaryStripFactor = 1.2d;
        // Inter-axis gap addend for chained same-side value axes: the middle gap reads
        // sideGap plus this on three Office renders (7/9/12pt, within 0.08pt).
        public const double BarDualValueAxisInterAxisGap = 1.46d;
        public const double HorizontalBarTitleNoLegendPlotBoxXRatio = 0.1524d;
        public const double HorizontalBarTitleNoLegendPlotBoxYRatio = 0.0924d;
        public const double HorizontalBarTitleNoLegendPlotBoxWidthRatio = 0.8196d;
        public const double HorizontalBarTitleNoLegendPlotBoxHeightRatio = 0.8003d;
        public const double HorizontalBarManualLayoutTargetPlotBoxXRatio = 0.179d;
        public const double HorizontalBarManualLayoutTargetPlotBoxYRatio = 0.1005d;
        public const double LineNoTitleRightLegendPlotBoxXRatio = 0.0828d;
        // Untitled right-legend plot bottom sits on the Office category-axis line (111.90pt
        // on 432pt frames across 19 cached references, all chart kinds; the clip padding below
        // it is not reproduced). The top stays fixed.
        public const double LineNoTitleRightLegendPlotBoxYRatio = 0.09236111111111111d;
        public const double LineNoTitleRightLegendPlotBoxWidthRatio = 0.7687d;
        public const double LineNoTitleRightLegendPlotBoxHeightRatio = 0.87064814814814812d;
        public const double LineNoTitleRightLegendExplicitScalePlotBoxYRatio = 0.10128571428571429d;
        public const double LineNoTitleRightLegendExplicitScalePlotBoxHeightRatio = 0.8463333333333334d;
        // Reserve between the widest value-axis tick label and the plot area for untitled
        // right-legend line/area/scatter charts. Calibrated from cached Office references:
        // plot-left minus our measured widest-label width decomposes to 23.2pt on 720pt
        // frames for scatter-clusters (single-digit), line-markers/stacked and area ports
        // (three-digit), and line-3series (four-digit, which needs no extra character term).
        public const double LineRightLegendValueAxisPadding = 23.2d;
        // Tail past the centered edge tick on horizontal-bar value axes. Calibrated from
        // cached Office references: axis-relative right margin minus half the widest
        // bottom label decomposes to 11.0pt on stacked, clustered, and 9pt-value probes
        // (10.994 to 11.001, sigma 0.004; clip-relative math reads 10.31 because the
        // Office plot clip pads the axis rect by about 0.7pt on the right).
        public const double HorizontalBarValueAxisRightPadding = 11.0d;
        // Top reserve for untitled horizontal-bar plots: the Office axis top sits
        // frameTop minus 11.0pt on stacked, short-frame, and 9pt-value probes (all
        // exact; font- and frame-invariant, unlike the vertical top; titled paths keep
        // their own presets).
        public const double HorizontalBarPlotTopReserve = 11.0d;
        // Preset-floor slop for horizontal-bar plot top and bottom: the measured
        // reserve replaces the preset only on clear disagreement (stacked bottom
        // 0.29 keeps the preset axis-exact bottom; short-frame top 0.99 replaces).
        public const double HorizontalBarPlotFloorSlop = 0.5d;
        // Plot-clip pad past the axis-bounded plot rect for bar/column charts: the Office
        // clip extends past our axis-coincident clip on the bottom and right edges only
        // (horizontal stacked -0.68/+0.68, clustered -0.68/+0.72, axis-titles -0.71/+0.69,
        // shifted-frame -0.71/+0.65; vertical column-stacked -0.68/+0.68, column-clustered
        // -0.68/+0.68; linewidth-independent across 0.75 and 1.0 axis strokes while axes
        // and gridlines already match within 0.04; dual-axis plots keep larger strip
        // residuals on top). Clip-only: geometry, labels, legends, and titles keep the
        // unpadded plot box.
        public const double BarPlotClipPad = 0.69d;
        // Left indent of vertical-bar value tick labels from the chart frame, then a
        // font-relative gap to the axis (Office origins decompose to frame plus 6.5pt
        // plus tick width plus 0.92 times tick font size on column-stacked, column-clustered,
        // dashboard, and composite ports, within 0.1; two exact decompositions read 0.924-0.928,
        // kept at 0.92 until wider axis samples justify the bump without golden churn).
        public const double BarValueAxisLabelFrameIndent = 6.5d;
        public const double BarValueAxisLabelGapFactor = 0.92d;
        // Preset-floor slop for the single-value-axis strip: the measured reserve
        // replaces the preset only on clear disagreement (legend-keys 4.45 vs
        // column-clustered 0.06); inside slop the preset stands untouched.
        public const double BarValueAxisPresetFloorSlop = 1d;
        // Minimum Office top clearance for untitled legendless columns (10.9pt on three
        // short-frame renders; the frame ratio stands above it).
        public const double ColumnPlotTopMarginFloor = 10.9d;
        // Office right reserve for column plots: the plot AXIS right edge keeps 11.0pt
        // inside the frame (axis strokes, not clips: the prior 10.3 came from clip boxes,
        // which pad the axis rect by the known 0.7; compact/overlay/bottom-legend Office
        // axes all sit at 11.0 while our clip-bound rights overshot by 0.66-0.70).
        public const double ColumnPlotRightReserve = 11.0d;
        public const double LineRightLegendValueAxisFrameWidthPaddingRatio = 0.05d;
        // Tail reserve past the legend marker block for untitled right-legend line/scatter
        // charts. Calibrated from cached Office references: reserve minus our measured marker
        // block and widest legend text decomposes to 10.6-11.0pt on 720pt frames for scatter,
        // line-3series, line-markers and line-stacked ports (the character-count extra term
        // double-counts length the widest name already spans).
        public const double LineScatterRightLegendReservePadding = 10.8d;
        // Tail past the legend marker block for untitled right-legend scatter charts with
        // line-sample keys (Office 10.01-10.03 across seven cached references; line keeps
        // its own tail, marker-only keys keep the legacy layout untouched).
        public const double ScatterRightLegendReserveTail = 10.0d;
        // Area right-legend reserve block past the widest legend text and half the last
        // category label. Calibrated from cached Office references (two public ports plus
        // six purpose-built probes): reserve minus our measured widest text and half the
        // last category label decomposes to 42.05pt on all eight (within 0.06), covering
        // the fill key block, the 17.5pt swatch lead past the last category label end,
        // and the tail below.
        public const double AreaRightLegendFixedBlock = 42.05d;
        // Tail past the longest legend entry for area right legends (Office 9.85-10.04
        // across the same eight references); the content box right-anchors to it.
        public const double AreaRightLegendTail = 10.0d;
        // Gap between the fill swatch and the entry text for area right legends (Office
        // 4.65-4.67 across five same-branch references; stroke keys keep their own gap).
        public const double AreaRightLegendTextGap = 4.66d;
        public const double LineTitleRightLegendPlotBoxXRatio = 0.0639d;
        public const double LineTitleRightLegendPlotBoxYRatio = 0.0924d;
        public const double LineTitleRightLegendPlotBoxWidthRatio = 0.7391d;
        public const double LineTitleRightLegendPlotBoxHeightRatio = 0.7888d;
        public const double BubbleTitleRightLegendPlotBoxWidthRatio = 0.7863333333333333d;
        public const double BubbleTitleRightLegendSwatchXRatio = 0.8818333333333334d;
        public const double BubbleTitleRightLegendSwatchYRatio = 0.4476157407407407d;
        public const double PieCenterXRatio = 0.42d;
        public const double PieNoLegendCenterXRatio = 0.5d;
        public const double DoughnutRightLegendCenterXRatio = 0.3988d;
        public const double DoughnutLeftLegendCenterXRatio = 0.56743d;
        public const double DoughnutHorizontalLegendCenterXRatio = 0.5d;
        public const double DoughnutTopLegendCenterYRatio = 0.46095d;
        public const double DoughnutBottomLegendCenterYRatio = 0.53907d;
        public const double DoughnutHorizontalLegendRadiusRatio = 0.4355d;
        public const double DoughnutNoLegendCenterYRatio = 0.5d;
        public const double DoughnutNoLegendRadiusRatio = 0.4746d;
        public const double DoughnutExplosionCenterOffsetRatio = 1.0d;
        // Doughnut ring-to-plot-edge side margin: portrait/square Office rings fit
        // (plotW - 2m)/2 exactly (m = 10.97pt both sides, two samples), capped by the
        // height rule on wide plots.
        public const double DoughnutPlotSideMargin = 10.97d;
        // Uniform frame tail past the longest right-legend entry for doughnut plots
        // (10.0 to 10.1pt over three Office renders, same phenomenon as the cartesian
        // 10pt tails).
        public const double DoughnutRightLegendTail = 10.0d;
        // Exploded right-legend slack: Office keeps plotRight minus legendBoxLeft at 35pt on
        // Exploded right-legend clearance: Office keeps 86.3pt clear between the legend box
        // left edge and the unshrunk ring right edge (narrow/mid/wide/short-mid Office renders
        // agree within 0.05); narrow sits at the boundary, guarded by the exploded-port tripwires.
        // Replaces the content-slack form, which died on short plots where identical content needs
        // no reserve. Geometry translation only; the plot box stays full-frame.
        public const double DoughnutExplodedLegendClearance = 86.3d;
        // plot middle on eight Office renders (narrow/mid/wide/notitle/unexploded/portrait/square
        // at n=3 plus a 5-entry probe, sigma 0.05) with identical pitch, so the (n-1) block
        // scaling stands and only the anchor moves. Left legends keep their own offset (minus
        // 5.0pt on one sample) for lack of evidence.
        public const double DoughnutRightLegendVerticalShift = 22.92d;
        // Doughnut left-legend vertical shift: the block middle sits 5.0pt below the plot middle
        // (two Office renders at n=3 and n=5, same middle, sigma 0.01); the left x anchor stays
        // legacy on a single box sample.
        public const double DoughnutLeftLegendVerticalShift = 5.0d;
        // Doughnut left-legend head inset: the Office box left edge sits at frame.X plus 13.19pt
        // on four renders (narrow/wide/square/exploded-left, exact to 0.01) regardless of content
        // width, i.e. the box grows rightward from a fixed head edge with the shared 4.65pt gap.
        public const double DoughnutLeftLegendHeadInset = 13.19d;
        // Left-ring center lead past the remaining-space middle (narrow/mid/wide/square Office
        // box widths agree within 0.02).
        public const double DoughnutLeftRingCenterLead = 2.05d;
        // Titled doughnut side legends center the ring a measured band below the plot middle
        // (four 270.32 renders against four 288.0 untitled); the band holds the title zone.
        public const double DoughnutTitledCenterYOffset = 17.68d;
        // Single-series vary-colors shade: Office darkens per-point bar/column fills by
        // 0.88x once the series holds six or more points (4/5pt raw on neg4/neg5/dash5,
        // 6pt dark on neg6/allpos6/dash6, 7pt dark plus a light slot-7 tint on dash7;
        // factor mean 0.8797 over 18 channels, HLS-lumMod killed by accent6). Horizontal
        // bars keep legacy output (no 6pt horizontal sample); pies/doughnuts/markers and
        // explicit or negative fills keep theirs (pie ports green raw, markers exact).
        public const double SingleSeriesVaryColorsShadeFactor = 0.88d;
        // Unstyled line-chart series-stroke tint: Office modulates the raw theme-accent
        // luminance by 0.975 in HSL (3 cached Office refs, 9/9 bytes exact across blue,
        // red, and green series), while fills and explicitly styled strokes keep raw
        // colors. Linear per-channel scaling is killed (predicts 77 vs observed 74).
        public const double UnstyledLineStrokeLuminanceFactor = 0.975d;
        public const int SingleSeriesVaryColorsShadePointThreshold = 6;
        // Second variation regime at twelve-plus points (dash12 twice plus dash13):
        // slots 1-6 shade at 0.82 (maxabs 1 over 18 channels) and slots 7-12 go raw
        // (byte-exact accents); slot-13 pale teal, slot-14 dusty pink, slot-15 mint and slot-16 periwinkle are fixed with slot-17-plus on
        // shaded-cycling fallback (cyan single sample in the 12-17 window; eighteen-plus counts take the third-regime rows instead; confound noted: all probes share one frame).
        public const double SingleSeriesVaryColorsSecondRegimeShadeFactor = 0.82d;
        public const int SingleSeriesVaryColorsSecondRegimePointThreshold = 12;
        // Third variation regime at eighteen-plus points (dash18/dash19 Office fills
        // agree): slots 1-6 shade at 0.78, slots 7-12 at 0.93 (maxabs 2 and 1 over 18
        // channels each) and slots 13-18 take a fixed light row (six byte-exact vectors);
        // slots 19-22 (pale sky, dusty mauve, pistachio, lilac) are fixed with slot-23-plus on first-regime cycling fallback (cyan single sample in the 18-23 window;
        // twenty-four-plus counts take the fourth-regime rows instead).
        public const double SingleSeriesVaryColorsThirdRegimeDarkShadeFactor = 0.78d;
        public const double SingleSeriesVaryColorsThirdRegimeMidShadeFactor = 0.93d;
        public const int SingleSeriesVaryColorsThirdRegimePointThreshold = 18;
        public static readonly RgbColor SingleSeriesVaryColorsThirdRegimeSlot13Fill = new(126, 155, 200);
        public static readonly RgbColor SingleSeriesVaryColorsThirdRegimeSlot14Fill = new(202, 126, 125);
        public static readonly RgbColor SingleSeriesVaryColorsThirdRegimeSlot15Fill = new(174, 198, 131);
        public static readonly RgbColor SingleSeriesVaryColorsThirdRegimeSlot16Fill = new(155, 137, 179);
        public static readonly RgbColor SingleSeriesVaryColorsThirdRegimeSlot17Fill = new(124, 187, 207);
        public static readonly RgbColor SingleSeriesVaryColorsThirdRegimeSlot18Fill = new(248, 170, 121);
        // Single-series vary-colors slot-7 overflow: Office paints the 7th point
        // light steel (147,169,207) instead of a shaded accent (dash7/dash8 Office
        // renders agree bit-identically at 0.576/0.663/0.812.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot7Fill = new(147, 169, 207);
        // Single-series vary-colors slot-8 overflow: Office paints the 8th point
        // dusty rose (209,147,146) instead of a shaded accent (dash8/dash9 Office
        // renders agree at 0.82/0.576/0.573, stable across the 8-to-9 transition.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot8Fill = new(209, 147, 146);
        // Single-series vary-colors slot-9 overflow: Office paints the 9th point light
        // green (185,205,150) instead of a shaded accent (dash9/dash10 Office renders
        // agree at 0.725/0.804/0.588, killing shaded-cycle and raw-restart rivals.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot9Fill = new(185, 205, 150);
        // Single-series vary-colors slot-10 overflow: Office paints the 10th point
        // lavender (169,155,189) instead of a shaded accent (dash10/dash11 Office renders
        // agree at 0.663/0.608/0.741, killing shaded-accent and raw-accent rivals.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot10Fill = new(169, 155, 189);
        // Regime-2 fixed tail: Office paints the 13th point pale teal (170,186,215)
        // (dash13/dash14 Office fills agree at 0.667/0.729/0.843.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot13Fill = new(170, 186, 215);
        // Regime-2 fixed tail, second entry: Office paints the 14th point dusty pink
        // (217,170,169) (dash14/dash15 Office fills agree at 0.851/0.667/0.663.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot14Fill = new(217, 170, 169);
        // Regime-2 fixed tail, third entry: Office paints the 15th point mint
        // (198,214,172) (dash15/dash16 Office fills agree at 0.776/0.839/0.675.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot15Fill = new(198, 214, 172);
        // Regime-2 fixed tail, fourth entry: Office paints the 16th point periwinkle
        // (186,176,201) (dash16/dash17 Office fills agree at 0.729/0.69/0.788.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot16Fill = new(186, 176, 201);
        // Post-light-row fixed entry: Office paints the 19th point pale sky
        // (182,195,220) (dash19/dash20 Office fills agree at 0.714/0.765/0.863.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot19Fill = new(182, 195, 220);
        // Post-light-row fixed entry, second point: Office paints the 20th point dusty mauve
        // (221,182,181) (dash20/dash21 Office fills agree at 0.867/0.714/0.71.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot20Fill = new(221, 182, 181);
        // Post-light-row fixed entry, third point: Office paints the 21st point pale pistachio
        // (205,219,184) (dash21/dash22 Office fills agree at 0.804/0.859/0.722.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot21Fill = new(205, 219, 184);
        // Post-light-row fixed entry, fourth point: Office paints the 22nd point pale lilac
        // (195,186,208) (dash22/dash23 Office fills agree at 0.765/0.729/0.816); slot-23
        // pale cyan is a single sample, theme dependence and 24-plus cycling unobserved).
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot22Fill = new(195, 186, 208);
        // Fourth variation regime at twenty-four-plus points (dash24/dash25 Office fills
        // agree on all twenty-four vectors): slots 1-6 take a fixed dark row (six byte-exact vectors near a 0.74 shade, no pure-linear fit within 2 units);
        // slots 7-12 shade at 0.88 (maxabs 2 over 18 channels); slots 13-18 go raw (byte-exact accents);
        // slots 19-24 replay the first-regime fixed tail byte-exact (steel/rose/green/lavender/teal/peach);
        // slot-25 periwinkle-blue is fixed with slot-26-plus on first-regime cycling fallback (rose single sample).
        public const int SingleSeriesVaryColorsFourthRegimePointThreshold = 24;
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot1Fill = new(57, 96, 142);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot2Fill = new(144, 58, 56);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot3Fill = new(116, 140, 65);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot4Fill = new(95, 73, 121);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot5Fill = new(54, 129, 149);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot6Fill = new(186, 112, 50);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot19Fill = new(147, 169, 207);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot20Fill = new(209, 147, 146);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot21Fill = new(185, 205, 150);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot22Fill = new(169, 155, 189);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot23Fill = new(145, 195, 213);
        public static readonly RgbColor SingleSeriesVaryColorsFourthRegimeSlot24Fill = new(249, 181, 144);
        // Post-replay fixed entry: Office paints the 25th point periwinkle-blue
        // (188,200,223) (dash25/dash26 Office fills agree at 0.737/0.784/0.875.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot25Fill = new(188, 200, 223);
        // Post-replay fixed entry, second point: Office paints the 26th point pale rose
        // (224,188,188) (dash26/dash27 Office fills agree at 0.878/0.737/0.737.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot26Fill = new(224, 188, 188);
        // Post-replay fixed entry, third point: Office paints the 27th point pale mint
        // (209,222,190) (dash27/dash28 Office fills agree at 0.82/0.871/0.745.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot27Fill = new(209, 222, 190);
        // Post-replay fixed entry, fourth point: Office paints the 28th point pale grape
        // (200,192,212) (dash28/dash29 Office fills agree at 0.784/0.753/0.831); slot-29
        // pale azure is a single sample, theme dependence and 30-plus cycling unobserved).
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot28Fill = new(200, 192, 212);
        // Fifth variation regime at thirty-plus points (dash30/dash31 Office fills
        // agree on all thirty vectors): slots 1-6 take a fixed dark row (six byte-exact vectors);
        // slots 7-12 shade at 0.85 (maxabs 2 over 18 channels) and slots 13-18 at 0.95 (maxabs 1);
        // slots 19-24 and 25-30 take fixed rows (twelve byte-exact vectors);
        // slot-31 lavender-blue is fixed with slot-32-plus on first-regime cycling fallback (coral single sample).
        public const int SingleSeriesVaryColorsFifthRegimePointThreshold = 30;
        public const double SingleSeriesVaryColorsFifthRegimeMidShadeFactor = 0.85d;
        public const double SingleSeriesVaryColorsFifthRegimeLightShadeFactor = 0.95d;
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot1Fill = new(56, 93, 138);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot2Fill = new(140, 56, 54);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot3Fill = new(113, 137, 63);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot4Fill = new(92, 71, 118);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot5Fill = new(53, 125, 145);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot6Fill = new(182, 109, 49);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot19Fill = new(115, 148, 197);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot20Fill = new(200, 115, 114);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot21Fill = new(169, 195, 121);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot22Fill = new(148, 128, 174);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot23Fill = new(112, 183, 205);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot24Fill = new(248, 165, 110);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot25Fill = new(161, 180, 212);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot26Fill = new(214, 161, 160);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot27Fill = new(192, 210, 164);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot28Fill = new(179, 168, 196);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot29Fill = new(160, 202, 217);
        public static readonly RgbColor SingleSeriesVaryColorsFifthRegimeSlot30Fill = new(249, 190, 158);
        // Post-palest fixed entry: Office paints the 31st point lavender-blue
        // (194,205,225) (dash31/dash32 Office fills agree at 0.761/0.804/0.882.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot31Fill = new(194, 205, 225);
        // Post-palest fixed entry, second point: Office paints the 32nd point pale coral
        // (226,194,194) (dash32/dash33 Office fills agree at 0.886/0.761/0.761.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot32Fill = new(226, 194, 194);
        // Post-palest fixed entry, third point: Office paints the 33rd point pale spring
        // (213,224,196) (dash33/dash34 Office fills agree at 0.835/0.878/0.769.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot33Fill = new(213, 224, 196);
        // Post-palest fixed entry, fourth point: Office paints the 34th point pale violet
        // (205,198,215) (dash34/dash35 Office fills agree at 0.804/0.776/0.843); slot-35
        // pale sky-cyan is a single sample, theme dependence and 36-plus cycling unobserved).
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot34Fill = new(205, 198, 215);
        // Sixth variation regime at thirty-six-plus points (dash36/dash37 Office fills
        // agree on all thirty-six vectors): slots 1-6 take a fixed dark row (six byte-exact vectors, no round linear fit within 2);
        // slots 7-12 shade at the second-regime 0.82 (maxabs 1) and slots 13-18 at 0.91 (maxabs 1);
        // slots 19-24 go raw (byte-exact accents); slots 25-30 take a fixed row and slots 31-36 replay the regime-2 tail, both byte-exact;
        // slot-37 periwinkle is fixed with slot-38-plus on first-regime cycling fallback (terracotta single sample).
        public const int SingleSeriesVaryColorsSixthRegimePointThreshold = 36;
        public const double SingleSeriesVaryColorsSixthRegimeLightShadeFactor = 0.91d;
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot1Fill = new(54, 90, 134);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot2Fill = new(136, 55, 52);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot3Fill = new(109, 133, 61);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot4Fill = new(90, 69, 114);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot5Fill = new(51, 122, 141);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot6Fill = new(177, 106, 47);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot25Fill = new(133, 160, 202);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot26Fill = new(205, 134, 132);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot27Fill = new(177, 201, 138);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot28Fill = new(160, 143, 182);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot29Fill = new(131, 190, 209);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot30Fill = new(248, 174, 129);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot31Fill = new(170, 186, 215);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot32Fill = new(217, 170, 169);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot33Fill = new(198, 214, 172);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot34Fill = new(186, 176, 201);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot35Fill = new(169, 206, 220);
        public static readonly RgbColor SingleSeriesVaryColorsSixthRegimeSlot36Fill = new(250, 195, 168);
        // Post-replay fixed entry: Office paints the 37th point periwinkle
        // (197,207,226) (dash37/dash38 Office fills agree at 0.773/0.812/0.886.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot37Fill = new(197, 207, 226);
        // Post-replay fixed entry, second point: Office paints the 38th point pale terracotta
        // (228,197,197) (dash38/dash39 Office fills agree at 0.894/0.773/0.773.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot38Fill = new(228, 197, 197);
        // Post-replay fixed entry, third point: Office paints the 39th point pale lime
        // (215,226,199) (dash39/dash40 Office fills agree at 0.843/0.886/0.78.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot39Fill = new(215, 226, 199);
        // Post-replay fixed entry, fourth point: Office paints the 40th point pale heather
        // (207,201,217) (dash40/dash41 Office fills agree at 0.812/0.788/0.851.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot40Fill = new(207, 201, 217);
        // Post-replay fixed entry, fifth point: Office paints the 43rd point sky-blue
        // (200,209,228) (dash43/dash44 Office fills agree at 0.784/0.82/0.894.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot43Fill = new(200, 209, 228);
        // Post-replay fixed entry, sixth point: Office paints the 44th point pale salmon
        // (229,200,199) (dash44/dash45 Office fills agree at 0.898/0.784/0.78.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot44Fill = new(229, 200, 199);
        // Post-replay fixed entry, seventh point: Office paints the 45th point pale apple
        // (217,227,201) (dash45/dash46 Office fills agree at 0.851/0.89/0.788.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot45Fill = new(217, 227, 201);
        // Post-replay fixed entry, eighth point: Office paints the 46th point pale lilac-gray
        // (209,203,219) (dash46/dash47 Office fills agree at 0.82/0.796/0.859); slot-47
        // pale cyan-blue is a single sample, theme dependence and 48-plus cycling unobserved).
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot46Fill = new(209, 203, 219);
        // Seventh variation regime at forty-two-plus points (dash42/dash43 Office fills
        // agree on all forty-two vectors): slots 1-12 take fixed dark and mid rows (twelve byte-exact vectors, no round linear fits within 2);
        // slots 13-18 reuse the 0.88 shade (byte-exact first-regime dark strings) and slots 19-24 shade at 0.96 (maxabs 1);
        // slots 25-30 take a fixed row and slots 31-36 replay the first-regime tail, both byte-exact; slots 37-42 take a fixed palest row;
        // slot-43-plus falls back to first-regime cycling (sky-blue single sample).
        public const int SingleSeriesVaryColorsSeventhRegimePointThreshold = 42;
        public const double SingleSeriesVaryColorsSeventhRegimeLightShadeFactor = 0.96d;
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot1Fill = new(53, 89, 132);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot2Fill = new(134, 53, 51);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot3Fill = new(107, 130, 60);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot4Fill = new(88, 68, 112);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot5Fill = new(50, 119, 138);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot6Fill = new(173, 104, 46);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot7Fill = new(62, 102, 151);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot8Fill = new(154, 62, 60);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot9Fill = new(124, 150, 70);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot10Fill = new(102, 79, 129);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot11Fill = new(58, 137, 159);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot12Fill = new(198, 119, 54);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot25Fill = new(106, 143, 195);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot26Fill = new(197, 106, 104);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot27Fill = new(165, 193, 112);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot28Fill = new(142, 120, 171);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot29Fill = new(103, 180, 203);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot30Fill = new(248, 160, 100);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot31Fill = new(147, 169, 207);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot32Fill = new(209, 147, 146);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot33Fill = new(185, 205, 150);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot34Fill = new(169, 155, 189);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot35Fill = new(145, 195, 213);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot36Fill = new(249, 181, 144);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot37Fill = new(175, 190, 217);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot38Fill = new(219, 175, 175);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot39Fill = new(201, 216, 177);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot40Fill = new(190, 180, 204);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot41Fill = new(174, 209, 222);
        public static readonly RgbColor SingleSeriesVaryColorsSeventhRegimeSlot42Fill = new(250, 199, 173);
        // Eighth variation regime at forty-eight-plus points (dash48/dash49 Office fills
        // agree on all forty-eight vectors): slots 1-6 take a fixed dark row (six byte-exact vectors, no round linear fit within 2);
        // slots 7-12 reuse the third-regime 0.78 dark row and slots 19-24 the third-regime 0.93 mid row (maxabs 1);
        // slots 13-18 shade at 0.86 (maxabs 2); slots 25-30 go raw (byte-exact accents);
        // slots 31-36 reuse the third-light row, slots 37-42 take a fixed row and slots 43-48 a fixed palest row, all byte-exact;
        // slot-49-plus falls back to first-regime cycling (steel-blue single sample).
        public const int SingleSeriesVaryColorsEighthRegimePointThreshold = 48;
        public const double SingleSeriesVaryColorsEighthRegimeMidShadeFactor = 0.86d;
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot1Fill = new(52, 88, 130);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot2Fill = new(132, 53, 51);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot3Fill = new(106, 129, 59);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot4Fill = new(87, 67, 111);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot5Fill = new(49, 118, 137);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot6Fill = new(171, 102, 46);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot31Fill = new(126, 155, 200);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot32Fill = new(202, 126, 125);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot33Fill = new(174, 198, 131);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot34Fill = new(155, 137, 179);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot35Fill = new(124, 187, 207);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot36Fill = new(248, 170, 121);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot37Fill = new(157, 177, 210);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot38Fill = new(212, 157, 156);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot39Fill = new(190, 209, 160);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot40Fill = new(176, 164, 194);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot41Fill = new(156, 200, 216);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot42Fill = new(249, 187, 154);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot43Fill = new(182, 195, 220);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot44Fill = new(221, 182, 181);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot45Fill = new(205, 219, 184);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot46Fill = new(195, 186, 208);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot47Fill = new(181, 212, 224);
        public static readonly RgbColor SingleSeriesVaryColorsEighthRegimeSlot48Fill = new(250, 203, 180);
        // Post-palest fixed entry: Office paints the 49th point steel-blue
        // (203,212,229) (dash49/dash50 Office fills agree at 0.796/0.831/0.898.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot49Fill = new(203, 212, 229);
        // Post-palest fixed entry, second point: Office paints the 50th point pale terracotta-rose
        // (230,203,202) (dash50/dash51 Office fills agree at 0.902/0.796/0.792.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot50Fill = new(230, 203, 202);
        // Post-palest fixed entry, third point: Office paints the 51st point pale honeydew
        // (218,228,204) (dash51/dash52 Office fills agree at 0.855/0.894/0.8.)
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot51Fill = new(218, 228, 204);
        // Post-palest fixed entry, fourth point: Office paints the 52nd point pale lavender-gray
        // (212,206,220) (dash52/dash53 Office fills agree at 0.831/0.808/0.863); slot-53
        // pale sky is a single sample, theme dependence and 54-plus cycling unobserved).
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot52Fill = new(212, 206, 220);
        // Ninth variation regime at fifty-four-plus points (dash54/dash55 Office fills
        // agree on all fifty-four vectors): slots 1-6 take a fixed dark row (six byte-exact vectors, no round linear fit within 2);
        // slots 7-12, 13-18 and 19-24 shade at 0.76, 0.84 and 0.90 (maxabs 2 over 18 channels each); slots 25-30 reuse the 0.96 shade (maxabs 2);
        // slots 31-36, 37-42 and 43-48 take fixed rows and slots 49-54 a fixed ninth row, all byte-exact;
        // slot-55-plus falls back to first-regime cycling (periwinkle-blue single sample).
        public const int SingleSeriesVaryColorsNinthRegimePointThreshold = 54;
        public const double SingleSeriesVaryColorsNinthRegimeSecondRowShadeFactor = 0.76d;
        public const double SingleSeriesVaryColorsNinthRegimeThirdRowShadeFactor = 0.84d;
        public const double SingleSeriesVaryColorsNinthRegimeFourthRowShadeFactor = 0.90d;
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot1Fill = new(51, 86, 127);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot2Fill = new(130, 51, 49);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot3Fill = new(104, 126, 58);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot4Fill = new(85, 65, 109);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot5Fill = new(48, 116, 134);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot6Fill = new(168, 100, 45);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot31Fill = new(102, 141, 194);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot32Fill = new(197, 103, 101);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot33Fill = new(163, 192, 109);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot34Fill = new(140, 118, 170);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot35Fill = new(100, 178, 202);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot36Fill = new(247, 159, 96);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot37Fill = new(138, 163, 204);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot38Fill = new(206, 138, 137);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot39Fill = new(180, 202, 142);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot40Fill = new(163, 147, 185);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot41Fill = new(136, 192, 210);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot42Fill = new(249, 177, 134);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot43Fill = new(163, 181, 212);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot44Fill = new(214, 163, 162);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot45Fill = new(193, 211, 166);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot46Fill = new(180, 169, 197);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot47Fill = new(161, 203, 218);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot48Fill = new(250, 191, 160);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot49Fill = new(185, 198, 221);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot50Fill = new(223, 185, 184);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot51Fill = new(207, 220, 187);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot52Fill = new(197, 189, 210);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot53Fill = new(184, 214, 225);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeSlot54Fill = new(251, 205, 183);
        // Ninth-regime +6 replay row at fifty-five-plus points (dash55 through dash58 Office fills agree): slots 55-58 replay the eighth-regime overflow singles at idx48-51 within maxabs 2 (5/4/3/2 samples, bit-identical repeats); dedicated consts carry the observed bytes. Slot-59-plus stays open.
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeReplaySlot55Fill = new(204, 213, 230);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeReplaySlot56Fill = new(231, 204, 204);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeReplaySlot57Fill = new(219, 229, 205);
        public static readonly RgbColor SingleSeriesVaryColorsNinthRegimeReplaySlot58Fill = new(213, 207, 221);
        // Tenth variation regime at sixty-plus points (dash60/dash61 Office fills agree on all sixty vectors): slots 1-6 take a fixed dark row (six byte-exact vectors, within 2 of the ninth dark row); slots 7-12 replay fourth-regime slots 1-6 byte-exact; slots 13-18, 19-24 and 25-30 shade at 0.82, 0.88 and 0.94 (maxabs 1/2/1); slots 31-36 take raw accents byte-exact; slots 37-42 take a fixed row (six byte-exact vectors); slots 43-48 replay seventh-regime slots 31-36 byte-exact; slots 49-54 replay sixth-regime slots 31-36 byte-exact; slots 55-58 replay overflow slots 25-28 byte-exact; slots 59-64 take a fixed tail (six byte-exact vectors, 3/2/2/2/2/2 samples); slot-65-plus falls back to overflow/0.88 cycling (slot59/slot60 still single-samples, slot61 new single).
        public const int SingleSeriesVaryColorsTenthRegimePointThreshold = 60;
        public const double SingleSeriesVaryColorsTenthRegimeFifthRowShadeFactor = 0.94d;
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot1Fill = new(50, 85, 126);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot2Fill = new(128, 51, 49);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot3Fill = new(103, 125, 57);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot4Fill = new(84, 65, 107);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot5Fill = new(47, 114, 132);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot6Fill = new(166, 99, 44);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot37Fill = new(118, 150, 198);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot38Fill = new(200, 118, 116);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot39Fill = new(170, 196, 123);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot40Fill = new(149, 130, 176);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot41Fill = new(115, 184, 205);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot42Fill = new(248, 166, 113);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot59Fill = new(187, 215, 227);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot60Fill = new(251, 207, 186);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot61Fill = new(205, 214, 230);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot62Fill = new(231, 205, 205);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot63Fill = new(220, 230, 207);
        public static readonly RgbColor SingleSeriesVaryColorsTenthRegimeSlot64Fill = new(214, 208, 222);
        // Eleventh variation regime at sixty-six-plus points (dash66/dash67 Office fills agree on all sixty-six vectors): slots 1-6 take a fixed dark row (six byte-exact vectors, within 2 of the tenth dark row); slots 7-12 take a fixed second row (six byte-exact vectors, best round shade 0.73 at maxabs 3); slots 13-18, 19-24, 25-30 and 31-36 shade at 0.80, 0.87, 0.92 and 0.97 (maxabs 2/2/1/1, re-verified on the independent dash67 vectors); slots 37-66 take five fixed rows (thirty byte-exact vectors, no clean tint fit, no const matches); slot-67-plus takes a fixed tail (four byte-exact vectors, 2/2/2/2 samples); slot-71-plus falls back to overflow/0.88 cycling (slot71 new single).
        public const int SingleSeriesVaryColorsEleventhRegimePointThreshold = 66;
        public const double SingleSeriesVaryColorsEleventhRegimeThirdRowShadeFactor = 0.80d;
        public const double SingleSeriesVaryColorsEleventhRegimeFourthRowShadeFactor = 0.87d;
        public const double SingleSeriesVaryColorsEleventhRegimeFifthRowShadeFactor = 0.92d;
        public const double SingleSeriesVaryColorsEleventhRegimeSixthRowShadeFactor = 0.97d;
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot1Fill = new(49, 84, 125);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot2Fill = new(127, 50, 48);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot3Fill = new(101, 123, 56);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot4Fill = new(83, 64, 106);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot5Fill = new(47, 113, 131);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot6Fill = new(164, 98, 43);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot7Fill = new(56, 94, 139);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot8Fill = new(142, 57, 55);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot9Fill = new(114, 138, 64);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot10Fill = new(93, 72, 119);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot11Fill = new(53, 126, 146);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot12Fill = new(183, 110, 49);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot37Fill = new(99, 139, 193);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot38Fill = new(196, 100, 97);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot39Fill = new(162, 191, 106);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot40Fill = new(138, 115, 168);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot41Fill = new(96, 177, 201);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot42Fill = new(247, 158, 92);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot43Fill = new(131, 159, 202);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot44Fill = new(204, 131, 130);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot45Fill = new(176, 200, 135);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot46Fill = new(158, 141, 181);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot47Fill = new(129, 189, 209);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot48Fill = new(248, 173, 127);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot49Fill = new(153, 174, 209);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot50Fill = new(211, 153, 152);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot51Fill = new(188, 208, 156);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot52Fill = new(173, 161, 192);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot53Fill = new(152, 198, 215);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot54Fill = new(249, 185, 150);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot55Fill = new(173, 189, 217);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot56Fill = new(218, 173, 173);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot57Fill = new(200, 215, 176);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot58Fill = new(189, 179, 203);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot59Fill = new(172, 208, 221);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot60Fill = new(250, 198, 171);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot61Fill = new(191, 203, 224);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot62Fill = new(225, 191, 191);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot63Fill = new(211, 223, 193);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot64Fill = new(202, 195, 213);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot65Fill = new(190, 217, 228);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot66Fill = new(251, 209, 189);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot67Fill = new(207, 215, 231);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot68Fill = new(232, 207, 206);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot69Fill = new(221, 230, 208);
        public static readonly RgbColor SingleSeriesVaryColorsEleventhRegimeSlot70Fill = new(215, 210, 223);
        // Twelfth variation regime at seventy-two-plus points (dash72/dash73 Office fills agree on all seventy-two vectors): slots 1-6 replay eleventh-regime slots 1-6 byte-exact; slots 7-12 and 13-18 shade at 0.73 and 0.79 (maxabs 2 each, re-verified on the independent dash73 vectors); slots 19-24, 25-30 and 31-36 reuse the 0.85, 0.90 and 0.95 shades (maxabs 2/2/1); slots 37-42 take raw accents byte-exact; slots 43-48 replay fifth-regime slots 19-24 byte-exact; slots 49-54 take a fixed row (six byte-exact vectors); slots 55-60 replay fifth-regime slots 25-30 byte-exact; slots 61-66 take a fixed row (six byte-exact vectors); slots 67-70 replay overflow slots 31-34 byte-exact; slots 71-72 take fixed tail singles (two byte-exact vectors, second-sampled in dash73); slot-73-plus falls back to overflow/0.88 cycling (slot73 new single).
        public const int SingleSeriesVaryColorsTwelfthRegimePointThreshold = 72;
        public const double SingleSeriesVaryColorsTwelfthRegimeSecondRowShadeFactor = 0.73d;
        public const double SingleSeriesVaryColorsTwelfthRegimeThirdRowShadeFactor = 0.79d;
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot49Fill = new(140, 165, 204);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot50Fill = new(207, 140, 139);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot51Fill = new(181, 203, 144);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot52Fill = new(164, 149, 186);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot53Fill = new(139, 192, 211);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot54Fill = new(249, 178, 137);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot61Fill = new(178, 193, 219);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot62Fill = new(220, 179, 178);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot63Fill = new(203, 218, 181);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot64Fill = new(192, 184, 206);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot65Fill = new(177, 210, 223);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot66Fill = new(250, 201, 176);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot71Fill = new(193, 219, 229);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot72Fill = new(251, 211, 193);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot73Fill = new(208, 216, 232);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot74Fill = new(232, 208, 208);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot75Fill = new(222, 231, 209);
        public static readonly RgbColor SingleSeriesVaryColorsTwelfthRegimeSlot76Fill = new(216, 211, 224);
        // Thirteenth variation regime at seventy-eight-plus points (dash78/dash79 Office fills agree on all seventy-eight vectors): slots 1-6 replay eleventh-regime slots 1-6 byte-exact; slots 7-12 take a fixed row (six byte-exact vectors, best round shade 0.71 at maxabs 3); slots 13-18 and 19-24 shade at 0.78 and 0.83 (maxabs 2 each, re-verified on the independent dash79 vectors); slots 25-30 reuse the 0.88 shade (maxabs 2); slots 31-36 and 37-42 shade at 0.93 and 0.98 (maxabs 1 each); slots 43-48 take a fixed row (six byte-exact vectors); slots 49-54 replay eighth-regime slots 31-36 byte-exact; slots 55-60 replay seventh-regime slots 31-36 byte-exact; slots 61-66 take a fixed row (six byte-exact vectors); slots 67-72 replay eighth-regime slots 43-48 byte-exact; slots 73-78 take a fixed row (six byte-exact vectors); slot-79 takes a fixed tail single (one byte-exact vector, 2 samples); slot-80-plus falls back to overflow/0.88 cycling (slot80 new single).
        public const int SingleSeriesVaryColorsThirteenthRegimePointThreshold = 78;
        public const double SingleSeriesVaryColorsThirteenthRegimeThirdRowShadeFactor = 0.78d;
        public const double SingleSeriesVaryColorsThirteenthRegimeFourthRowShadeFactor = 0.83d;
        public const double SingleSeriesVaryColorsThirteenthRegimeSixthRowShadeFactor = 0.93d;
        public const double SingleSeriesVaryColorsThirteenthRegimeSeventhRowShadeFactor = 0.98d;
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot7Fill = new(54, 91, 136);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot8Fill = new(138, 55, 53);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot9Fill = new(110, 134, 62);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot10Fill = new(91, 70, 116);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot11Fill = new(52, 123, 142);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot12Fill = new(178, 107, 48);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot43Fill = new(95, 137, 192);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot44Fill = new(195, 96, 94);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot45Fill = new(161, 190, 103);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot46Fill = new(136, 112, 167);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot47Fill = new(92, 176, 201);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot48Fill = new(247, 156, 89);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot61Fill = new(164, 182, 213);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot62Fill = new(215, 165, 164);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot63Fill = new(195, 212, 167);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot64Fill = new(182, 171, 198);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot65Fill = new(163, 204, 218);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot66Fill = new(250, 192, 162);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot73Fill = new(195, 206, 226);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot74Fill = new(227, 196, 195);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot75Fill = new(214, 225, 197);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot76Fill = new(206, 199, 216);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot77Fill = new(195, 220, 229);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot78Fill = new(251, 212, 194);
        public static readonly RgbColor SingleSeriesVaryColorsThirteenthRegimeSlot79Fill = new(208, 216, 232);
        public const double PieCenterYRatio = 0.458d;
        public const double PieRadiusRatio = 0.434d;
        // Labeled pies (any visible data labels, no legend): Office centers the pie
        // in the plot middle with a height-gated radius (360H renders at 0.40694 over
        // probe-1/auto/square/tall-offset; 300H renders at 0.40083 over offset and
        // short-probe1 with swapped content, so height drives it, not labels or width).
        // Unlabeled pies keep the constants above (5-categories port: 0.4595 center,
        // 0.4327 radius).
        public const double PieLabeledCenterYRatio = 0.5d;
        public const double PieLabeledRadiusRatio = 0.40694d;
        // Short-plot labeled-pie radius ratio (300H Office renders; the 330H cutoff
        // splits the only observed heights with maximum margin; form beyond is open).
        public const double PieShortLabeledRadiusRatio = 0.40083d;
        public const double PieShortLabeledPlotHeightCutoff = 330d;
        // Manual pie/doughnut labels anchor on a circle past the rim: radius R plus
        // 0.0566 R (joint X/Y fit over 19 unclamped Office single-line boxes, rms
        // 0.68/1.1; square-plot Beta/Delta pairs fit within 0.12 here against 1.06+
        // for the plot-width framing, which is identical on the constant-aspect
        // corpus; near-identical on probe 2 either way).
        public const double PieManualLabelCircleGapRadiusFactor = 0.0566d;
        // Manual pie/doughnut boxes are content-tight: 3.0pt side pads (9 Office box
        // fills agree within 0.15pt) around the unwrapped text, or the longest wrapped
        // line past the 0.19 plot-width wrap cap.
        public const double PieManualLabelBoxSidePad = 3.0d;
        // Manual pie/doughnut box height rides the line pitch plus this total pad
        // (7 Office boxes: H = 1.22 fs nLines + 3.0 within 0.06pt).
        public const double PieManualLabelBoxHeightPad = 3.0d;
        // Manual pie/doughnut text is bottom-anchored: the last baseline sits this far
        // above the box bottom (4 Office sizes 14/18/24/36pt fit 0.28 fs + 1.5 within
        // 0.04pt; auto labels keep the top-anchored ascent rule below).
        public const double PieManualLabelBaselinePadFontFactor = 0.28d;
        public const double PieManualLabelBaselinePadConstant = 1.5d;
        // Exact-vertical (north/south) manual labels shift toward the pie center by
        // this much (3 Office boxes within 1.1pt; south uses the single-line height
        // even when wrapped; near-cardinals like West 1.8deg off keep the radial rule).
        public const double PieManualLabelCardinalVerticalShift = 2.42d;
        // Manual pie/doughnut leaders go to same-side labels whose unwrapped text fits
        // 1.12 wrap caps on wide plots (27 Office configs: West at 96.9% of cap leads
        // while graded Gamma at 121.9% does not; narrow-line wrapped boxes like G2 at
        // 265% of cap never lead). Square plots (plot no wider than 1.1 heights) draw
        // no leaders at all (2 narrow same-side square qualifiers leadless).
        public const double PieManualLabelLeaderUnwrappedCapRatio = 1.12d;
        // Wide-plot gate for manual pie/doughnut leaders (square plots draw none).
        public const double PieManualLabelLeaderMinPlotAspect = 1.1d;
        // Wrapped manual-leader foot rise above the first baseline at 18pt (3 Office
        // samples exact at 150.79; single-size evidence so it scales with font size).
        public const double PieManualLabelLeaderWrappedFootRise = 2.35d;
        // Office emits a chart-area placeholder border by default (11 kind refs carry
        // the G:0 0.14pt rect with a fully transparent stroke and no chartArea markup
        // anywhere in the corpus); an explicit area stroke keeps winning and skips it.
        public const double ChartAreaDefaultBorderWidth = 0.14d;
        // Pie labels wrap past this fraction of the plot width (8 Office samples: every
        // wrapped label exceeds it, every single-line label stays below; tightest margins
        // are West 0.6pt and Gamma 2.8pt. A square-plot probe kills the radius basis
        // (Beta/Delta wrap at 73-80pt against 0.67R = 98.2) while 0.19 plotW holds at 68.4.
        public const double PieDataLabelWrapWidthFactor = 0.19d;
        // Wrapped pie label line pitch as a fraction of the tick font size (Office stacks
        // wrapped lines 21.9pt apart at 18pt; single 18pt sample, Gamma/North/South agree
        // to 0.1pt; our 24.3pt box height stays for boxes and clips).
        public const double PieDataLabelLinePitchFactor = 1.22d;
        public const double PieDataLabelRadiusRatio = 0.62d;
        // Pie auto labels sit at this fraction of the pie radius (9 Office samples across
        // two pies: mean 0.735, range 0.67 to 0.85 excluding bestFit overflow; doughnuts
        // keep the shared constant above through the hole-size max).
        public const double PieAutoDataLabelRadiusRatio = 0.74d;
        // Pie label baselines sit this fraction of the font size below the box top
        // (single-line Office labels: 16.0pt at 18pt). Bar labels are baseline-anchored
        // by construction and never take this path.
        public const double PieDataLabelBaselineFactor = 0.88d;
        public const double PieDataLabelWidthRatio = 0.55d;
        public const double PieDataLabelMinimumWidth = 18d;
        public const double PieDataLabelHeightFactor = 1.35d;
        public const double DataLabelLegendKeySizeFactor = 0.55d;
        public const double DataLabelLegendKeyTextGapFactor = 0.35d;
        // Swatch+text unit centering for clustered vertical-bar legend-key labels:
        // Office centers the unit on the bar middle plus this much (9 Office labels
        // across two gap widths at 8pt: sigma 0.03; consistent at 16pt with the
        // calibrated swatch gap (textX exact within 0.06), so textX stays on this law).
        public const double BarLegendKeyUnitCenterOffset = 1.45d;
        // Bottom reserve for vertical-bar category labels: frame margin plus
        // category descent plus our emitted category gap (Office baselines sit
        // frame plus 7.0 plus descent within 0.3 over 7, 14, and 18pt cats;
        // plots land within 1.0 using our 1.18 gap, column-clustered stays
        // on preset via slop; bottoms hold at 156.72 across value-range, tick-unit,
        // and tall-frame probes with identical category labels).
        public const double BarCategoryBottomMargin = 7.0d;
        public const double BarCategoryDescentFactor = 0.22d;
        // Out-end gap for legend-key data labels: 4.5 plus 0.28fs (Office baselines
        // sit barTop plus 6.66 at 8pt, 7.83 at 12pt, and 8.99 at 16pt: 20 labels over
        // three sizes within 0.08 of the law; the slope matches pie manual bottom-pad.
        // Bar tops must map in the axis frame (bars bottom on the category axis line
        // with axis scale); clip-frame mapping fakes a minus 0.05 per unit tilt and
        // cost one wrong constant bump already (do not repeat it).
        public const double BarLegendKeyOutEndBoxPad = 3.0d;
        public const double BarLegendKeyOutEndGapFontFactor = 0.28d;
        public const double BarLegendKeyOutEndGapConstant = 1.5d;
        // Swatch gap and vertical anchor for legend-key labels: Office gaps read
        // 4.50 at 8pt, 5.56 at 12pt, and 6.63 at 16pt (kills G equals S at 16pt,
        // 6.63 versus 8.8); swatch centers sit baseline plus 0.34fs (all sizes
        // within 0.11, Y anchor exact at 12pt). Three-size evidence.
        public const double BarLegendKeySwatchGapFactor = 0.266d;
        public const double BarLegendKeySwatchGapConstant = 2.37d;
        public const double BarLegendKeySwatchCenterOffsetFactor = 0.34d;
        public const double PieExplosionLabelRadiusRatio = 0.22d;
        public const double CartesianDataLabelHeightFactor = 1.35d;
        public const double CartesianDataLabelMinimumWidth = 18d;
        public const double BarDataLabelHorizontalGap = 2d;
        public const double BarDataLabelVerticalGap = 1d;
        public const double BarDataLabelSlotFillRatio = 0.82d;
        public const double BarDataLabelCategoryInsetRatio = 0.09d;
        public const double HorizontalBarDataLabelWidthRatio = 0.1d;
        public const double HorizontalBarDataLabelSlotCenterRatio = 0.43d;
        public const double VerticalBarDataLabelWidthRatio = 0.86d;
        public const double LineDataLabelMinimumPointSpan = 5d;
        public const double LineDataLabelPointWidthFactor = 1.5d;
        public const double LineDataLabelSideGap = 2d;
        public const double LineDataLabelBelowOffsetFactor = 1.25d;
        public const double LineDataLabelAboveOffsetFactor = 0.35d;
        public const double AxisLabelHeightFactor = 1.35d;
        public const int CategoryAxisDefaultLabelOffset = 100;
        public const int CategoryAxisMinimumLabelOffset = 0;
        public const int CategoryAxisMaximumLabelOffset = 1000;
        public const int CategoryAxisDefaultTickLabelSkip = 1;
        public const double CategoryAxisHorizontalLeftOffsetRatio = 0.1882d;
        public const double CategoryAxisHorizontalWidthRatio = 0.16d;
        public const double CategoryAxisHorizontalBaselineRatio = 0.217d;
        // Font-relative gap between horizontal-bar category label ink and the plot edge.
        // Calibrated from cached Office references: plot-left minus the common label right
        // edge over label font size gives 0.928 at 18pt (clustered 181.73-165.0, stacked
        // 145.20-128.5, inner 330.5-313.8, outer 443.2-426.5) and 0.922 at 9pt (axis-titles
        // probe 181.4-173.1), so the bar-default probe with 9pt labels sits at 8.3pt, not 16.7pt.
        public const double HorizontalBarCategoryLabelPlotGapFactor = 0.925d;
        // Pad between the outer manual area edge and horizontal-bar category label ink.
        // Calibrated from cached Office references: outer carve (plot-left minus area-left)
        // is constant 104.8pt across five probes (x10/x18/x24/x30/w50), decomposing exactly
        // to widest label plus the font-relative gap above plus 1.6pt here.
        public const double HorizontalBarOuterAreaLabelPad = 1.6d;
        // Right reserve between the outer manual area edge and the plot edge for
        // horizontal bars. Calibrated from cached Office references: area-right minus
        // plot-right is constant 15.1pt across the same five outer probes.
        public const double HorizontalBarOuterPlotRightReserve = 15.1d;
        public const double CategoryAxisVerticalWidthFactor = 1.35d;
        public const double CategoryAxisVerticalTopOffsetFactor = 1.18d;
        public const double CategoryAxisVerticalTopSideOffsetFactor = 0.70d;
        public const double CategoryAxisMajorTickLength = 4d;
        // Major tick-mark ink length scales with the tick label size: Office draws 5.71pt
        // ticks at 18pt labels (line trend/markers, scatter clusters) and 2.82pt at 9pt
        // (bar default-axis-titles), i.e. about 0.315fs; the fixed 4d above stays for
        // label-strip reserves (safe bound, not ink).
        public const double ChartAxisMajorTickLengthFactor = 0.315d;
        public const double AxisLabelClipTopOffsetFactor = 0.25d;
        public const double AxisLabelClipHeightFactor = 1.6d;
        public const double ValueAxisMinimumLabelWidthFactor = 1.6d;
        public const double ValueAxisLabelPaddingFactor = 0.45d;
        public const double ValueAxisLabelSideGapFactor = 0.93d;
        public const double ValueAxisVerticalClipExtraHeightFactor = 2d;
        public const double HorizontalValueAxisSlotCount = 5d;
        public const double HorizontalValueAxisTopOffsetFactor = 1.18d;
        public const double VerticalValueAxisWidthRatio = 0.12d;
        public const double VerticalValueAxisBaselineRatio = 0.215d;
        public const double TitleXInsetRatio = 0.08d;
        public const double TitleBaselineYRatio = 0.88d;
        public const double PolarTitleBaselineYRatio = 0.935d;
        // Polar auto-title inset below the title-box top: Office baselines sit 28.0pt under the top
        // on full, short, and moved frames alike (sigma 0.05); plot-top versus shape-top stays
        // confounded on full-frame plots. Explicit and manual titles keep the ratio fallback.
        public const double PolarTitleTopOffset = 28.0d;
        public const double AutoTitleFontScale = 1.2d;
        public const double TitleAbovePlotBaselineOffsetFactor = 0.8483333333333334d;
        public const double AutoBarTitleAbovePlotBaselineOffsetFactor = 1.08d;
        public const double TitleWidthRatio = 0.84d;
        public const double TitleHeightFactor = 1.4d;
        // Bottom default axis-title baseline: two Office renders (column category title
        // and bar value title, same frame, both 12pt) agree on frame bottom plus 0.338
        // times the plot-to-frame reserve within 0.01pt; the legacy 0.23 had no sliced
        // provenance. Different-frame reserves stay unobserved.
        public const double DefaultAxisTitleBandBaselineRatio = 0.338d;
        public const double DefaultAxisTitleSideBaselineRatio = 0.28d;
        public const double DefaultAxisTitleTopBandBaselineRatio = 0.485d;
        public const double DefaultAxisTitleLeftSideBaselineRatio = 0.50d;
        // Left vertical axis title column: identical 143.81pt on the bar, column, and
        // long-label probes sharing one 540pt frame while plots (181.4/168.1/227.4) and
        // label strips (22/10/69pt wide) all move, so the column is frame-anchored, not
        // plot- or strip-anchored. Same-frame corpus caveat: needs a different-frame probe.
        public const double DefaultAxisTitleLeftColumnOffset = 23.8d;
        public const double DefaultAxisTitleRightSideBaselineRatio = 0.67d;
        public const double DefaultAxisTitlePlotSideReserveRatio = 0.089d;
        public const double DefaultAxisTitlePlotOppositeSideReserveRatio = 0.020d;
        public const double DefaultAxisTitlePlotBandReserveRatio = 0.126d;
        public const double DefaultAxisTitlePlotOppositeBandReserveRatio = 0.030d;
        public const double DefaultAxisTitleHorizontalBarPlotSideReserveRatio = 0.114d;
        public const double DefaultAxisTitleHorizontalBarPlotOppositeSideReserveRatio = 0.028d;
        public const double DefaultAxisTitleHorizontalBarPlotBandReserveRatio = 0.126d;
        // Horizontal-bar default-axis-titles content reserves: three Office renders per side
        // overconstrain the linear form (residuals under 0.03pt, one spare DOF each). Top grows
        // 1.833pt per value-tick point (9/12/14.04pt ticks give 46.77/52.31/56.00); bottom grows
        // 1.222pt per chart-title point (21.6/17.28/14.4pt titles give 46.37/41.10/37.57). The
        // symmetric 0.126 band it replaces under-read both sides unevenly (top minus 0.65,
        // bottom minus 0.24); side-dependent content reserves follow the bottom-reserve-plane
        // method. Horizontals only: the vertical branch keeps its own ratios, unobserved.
        public const double DefaultAxisTitleHorizontalBarTopReserveBase = 30.29d;
        public const double DefaultAxisTitleHorizontalBarTopReservePerTickFontSize = 1.833d;
        public const double DefaultAxisTitleHorizontalBarBottomReserveBase = 19.98d;
        public const double DefaultAxisTitleHorizontalBarBottomReservePerTitleFontSize = 1.222d;
        // Reference title size baked into the horizontal-top law (all three top knots carry
        // 12pt value titles): the vertical-bottom composition re-bases from it to the actual
        // category-title size, reusing both fitted slopes with zero new constants.
        public const double DefaultAxisTitleHorizontalBarReferenceTitleFontSize = 12d;
        // Full-frame fill-legend row pitch: three Office overlay knots (20.51/27.77/35.00
        // at 12/18/24pt) fit 1.208fs plus 6.03 within 0.022 (one spare DOF); the shared 1.5433
        // stroke factor fits fs18 but misses fs24 by 2.0 per row. Scoped to full-frame fill
        // legends (overlay shape, sole evidence); other fill legends keep the stroke factor,
        // unobserved at other sizes (they agree within 0.02 at fs18).
        public const double FullFrameFillLegendLineHeightPerFontSize = 1.208d;
        public const double FullFrameFillLegendLineHeightBase = 6.03d;
        // Bottom-legend horizontal packing: six Office renders overconstrain both terms.
        // Swatch-text gap is content-independent (4.59/4.59/4.61 at fs18 across three fixtures)
        // and linear in font size (2.96/4.59/6.25 at 12/18/24pt, residuals under 0.01, four spare
        // DOF); the negative intercept is fitted, mechanism open. Inter-entry gap carries a
        // content term in entry-advance range (max minus min measured advance): 9.60/14.27/18.90
        // same-entries ladder plus 13.54/14.79/14.77 content knots fit 0.813fs minus 0.0276range
        // plus 0.285 within 0.031 on all six (three spare DOF); longest/total/average forms all go
        // non-monotonic over per-point entries and are killed there. Top legends keep legacy constants, unobserved.
        public const double LegendHorizontalTextGapFactor = 0.275d;
        public const double LegendHorizontalTextGapIntercept = -0.35d;
        public const double LegendHorizontalInterEntryGapFactor = 0.813d;
        public const double LegendHorizontalInterEntryGapRangeFactor = 0.0276d;
        public const double LegendHorizontalInterEntryGapBase = 0.285d;
        // Per-series bottom entries (multi-series names) answer to mean advance instead:
        // composite base/leg18/longnames/shortnames plus NSW fit 0.45fs plus 0.148avgW
        // plus 0.195 within 0.062 on all five (two spare DOF); no range term (shortnames
        // range 17.2 absorbed within 0.04, dedicated varied-range per-series knots unobserved).
        // Series-tagged fill-only entries take this arm; point and stroke-key entries keep
        // the range law above.
        public const double LegendPerSeriesInterEntryGapBase = 0.195d;
        public const double LegendPerSeriesInterEntryGapFontSizeFactor = 0.45d;
        public const double LegendPerSeriesInterEntryGapMeanAdvanceFactor = 0.148d;
        public const double LegendLineHeightFactor = 1.45d;
        public const double LegendSideStrokeLineHeightFactor = 1.5433333333333332d;
        public const double LegendMarkerSizeFactor = 0.55d;
        public const double LegendMinimumSideWidth = 36d;
        public const double LegendSideFillMinimumWidthFactor = 1.95d;
        public const double LegendSideWidthRatio = 0.22d;
        public const double LegendSideGap = 8d;
        public const double LegendSideStrokeGapFactor = 0.8333333333333334d;
        public const double LegendSideStrokeMarkerWidthFactor = 1.0666666666666667d;
        public const double LegendSideStrokeTextGapFactor = 0.11666666666666667d;
        // Right stroke-key legend blocks anchor to the plot CLIP edge rather than the
        // axis line: legendLeft minus axisRight reads a fixed 0.65 over sideGap plus leadExtra
        // on line-3series, line-markers, line-trend and scatter-smooth (same 0.65-0.7 band as
        // the shipped axis-vs-clip finding). Marker-only scatter keeps its separate residual.
        public const double LegendStrokeRightClipAllowance = 0.65d;
        public const double LegendSideStrokeBaselineCenterOffsetFactor = 0.955d;
        public const double LegendSideStrokeStyledMarkerBaselineCenterOffsetFactor = 0.561d;
        public const double LegendSideFillReservedBandOffsetFactor = 0.04d;
        public const double LegendSideFillContentBoxReservedBandOffsetFactor = 0.06d;
        public const double LegendSideFillBaselineCenterOffsetFactor = 0.276d;
        public const double LegendFullFrameSideInsetRatio = 0.02d;
        public const double LegendFullFrameBottomBaselineFactor = 0.57d;
        public const double LegendFullFrameTopBaselineFactor = 0.95d;
        // Bottom-legend text baseline rides the frame, not the plot: four Office renders
        // (two frames, 12/18pt legend fonts, titled/untitled, plus a plot-shifted control
        // that leaves the legend fixed) agree on frame bottom plus 8.6 plus 0.35 times
        // the legend font size within 0.1pt. Doughnut bottom legends keep their own
        // full-frame branch and are untouched.
        public const double LegendBottomFramePad = 8.6d;
        public const double LegendBottomBaselineFontFactor = 0.35d;
        // Bottom-legend block X anchors to the frame center plus markerSize over 4
        // plus a fitted 0.65 bias (twelve Office renders: botleg family, composite
        // family, doughnut bottom; block centers within 0.05 of frame-center plus bias
        // while plot-centeredness residuals run 1.8 to 4.5; the val14 splitter proves
        // frame anchoring decisively with the plot moved plus-6 and the block unmoved).
        // Top-positioned legends keep legacy; mechanism behind the bias stays open.
        public const double LegendBottomFrameAnchorBias = 0.65d;
        public const double LegendTopOffsetFactor = 0.15d;
        public const double LegendHorizontalClipHeightFactor = 1.25d;
        public const double LegendMarkerBaselineFactor = 0.35d;
        public const double LegendSideFillMarkerBaselineFactor = 0d;
        public const double LegendHorizontalMarkerBaselineFactor = 0d;
        public const double LegendTextGap = 3d;
        public const double LegendHorizontalEntryPadding = 8d;
    }

}
