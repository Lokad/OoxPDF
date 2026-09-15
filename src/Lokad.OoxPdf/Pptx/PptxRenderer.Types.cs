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
        // (byte-exact accents); slot-13 pale teal, slot-14 dusty pink and slot-15 mint are fixed with slot-16-plus on
        // shaded-cycling fallback (periwinkle single sample; confound noted: all probes share one frame).
        public const double SingleSeriesVaryColorsSecondRegimeShadeFactor = 0.82d;
        public const int SingleSeriesVaryColorsSecondRegimePointThreshold = 12;
        // Third variation regime at eighteen-plus points (dash18/dash19 Office fills
        // agree): slots 1-6 shade at 0.78, slots 7-12 at 0.93 (maxabs 2 and 1 over 18
        // channels each) and slots 13-18 take a fixed light row (six byte-exact vectors);
        // slot-19-plus falls back to first-regime cycling (pale-sky single sample).
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
        // (198,214,172) (dash15/dash16 Office fills agree at 0.776/0.839/0.675); slot-16
        // periwinkle is a single sample, theme dependence and 17-plus cycling unobserved).
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot15Fill = new(198, 214, 172);
        // Regime-2 fixed tail, fourth entry: Office paints the 16th point periwinkle
        // (186,176,201) (dash16/dash17 Office fills agree at 0.729/0.69/0.788); slot-17
        // bright cyan is a single sample, theme dependence and 18-plus cycling unobserved).
        public static readonly RgbColor SingleSeriesVaryColorsOverflowSlot16Fill = new(186, 176, 201);
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
