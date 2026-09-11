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
        public const double BubbleRadiusPlotRatio = 0.128d;
        // Default radar series outline width. Office draws 3.75pt outlines on unfilled
        // radars without an explicit line (both ladder radar ports share style 118, so
        // revisit if other styles diverge); filled radars omit the outline instead.
        public const double RadarSeriesOutlineWidth = 3.75d;
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
        // Bottom reserve below the legend and tick rows for untitled bottom-legend columns.
        // Calibrated to the 54.4pt Office bottom margin measured on 360pt and 216pt frames
        // (12pt and 18pt legends over 12pt and 8pt ticks respectively).
        public const double BarNoTitleBottomLegendContentGap = 19.2d;
        public const double StackedColumnBottomLegendPlotBoxLeftPadding = 1.8d;
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
        // cached Office references: right margin minus half the widest bottom label
        // decomposes to 11.0pt on stacked (300) and clustered (50) bar ports.
        public const double HorizontalBarValueAxisRightPadding = 11.0d;
        // Left indent of vertical-bar value tick labels from the chart frame, then a
        // font-relative gap to the plot edge (Office origins decompose to frame plus 6.5pt
        // plus tick width plus 0.92 times tick font size on column-stacked, column-clustered,
        // dashboard, and composite ports, within 0.1).
        public const double BarValueAxisLabelFrameIndent = 6.5d;
        public const double BarValueAxisLabelGapFactor = 0.92d;
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
        public const double PieCenterYRatio = 0.458d;
        public const double PieRadiusRatio = 0.434d;
        // Labeled pies (any visible data labels, no legend): Office centers the pie
        // in the plot with a smaller radius (probe-1/offset/auto/noleader renders:
        // centers exact at 0.5, radii 0.4007-0.4069 of plot height). Unlabeled pies
        // keep the constants above (5-categories port: 0.4595 center, 0.4327 radius).
        public const double PieLabeledCenterYRatio = 0.5d;
        public const double PieLabeledRadiusRatio = 0.404d;
        // Slice-anchored manual pie labels sit this fraction of the tick font size past
        // the rim midpoint (7 Office samples at 18pt average 9.0pt, range 6.4 to 11.8).
        public const double PieManualLabelEdgeGapFactor = 0.5d;
        // Pie labels wrap past this fraction of the plot width (8 Office samples: every
        // wrapped label exceeds it, every single-line label stays below; tightest margins
        // are West 0.6pt and Gamma 2.8pt. Plot-width, height, and radius bases are
        // indistinguishable on the constant-aspect corpus, so width wins as the natural
        // horizontal basis; revisit if a different-aspect pie disagrees.
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
        public const double AutoTitleFontScale = 1.2d;
        public const double TitleAbovePlotBaselineOffsetFactor = 0.8483333333333334d;
        public const double AutoBarTitleAbovePlotBaselineOffsetFactor = 1.08d;
        public const double TitleWidthRatio = 0.84d;
        public const double TitleHeightFactor = 1.4d;
        public const double DefaultAxisTitleBandBaselineRatio = 0.23d;
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
        public const double LegendSideStrokeBaselineCenterOffsetFactor = 0.955d;
        public const double LegendSideStrokeStyledMarkerBaselineCenterOffsetFactor = 0.561d;
        public const double LegendSideFillReservedBandOffsetFactor = 0.04d;
        public const double LegendSideFillContentBoxReservedBandOffsetFactor = 0.06d;
        public const double LegendSideFillBaselineCenterOffsetFactor = 0.276d;
        public const double LegendFullFrameSideInsetRatio = 0.02d;
        public const double LegendFullFrameBottomBaselineFactor = 0.57d;
        public const double LegendFullFrameTopBaselineFactor = 0.95d;
        public const double LegendBottomOffsetFactor = 2.61d;
        public const double LegendTopOffsetFactor = 0.15d;
        public const double LegendHorizontalClipHeightFactor = 1.25d;
        public const double LegendMarkerBaselineFactor = 0.35d;
        public const double LegendSideFillMarkerBaselineFactor = 0d;
        public const double LegendHorizontalMarkerBaselineFactor = 0d;
        public const double LegendTextGap = 3d;
        public const double LegendHorizontalEntryPadding = 8d;
    }

}
