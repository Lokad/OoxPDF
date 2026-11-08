using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxChartsTests
{
    public static void PptxSyntheticFullFrameFillLegendLineHeight()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeFullFrameFillLegendLineHeight",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected fill line height to remain inspectable by the Office evidence guard.");

        // Office overlay pitches 20.51/27.77/35.00 at 12/18/24pt (residuals under 0.025);
        // non-full-frame paths keep their factors.
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [12d, true, false, false])! - 20.51d) < 0.05d, "Expected fs12 fill pitch.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, true, false, false])! - 27.77d) < 0.05d, "Expected fs18 fill pitch.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [24d, true, false, false])! - 35.00d) < 0.05d, "Expected fs24 fill pitch.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, false, false, true])! - 27.78d) < 0.05d, "Expected side-fill legacy pitch.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, false, true, false])! - 27.78d) < 0.05d, "Expected stroke legacy pitch.");
    }

    public static void PptxSyntheticOverlayLegendTextGap()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ResolveSideLegendTextGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected side text gap to remain inspectable by the Office evidence guard.");

        // Overlay side keys take the measured 4.66pt fill-swatch gap like doughnut-right
        // (Office overlay swatch-to-text 4.66 exact); everything else keeps legacy routing.
        var placementType = typeof(PptxRenderer).GetNestedType("ChartLegendPlacement", System.Reflection.BindingFlags.NonPublic);
        object? placement = System.Enum.Parse(placementType!, "Default");
        TestAssert.Equal(4.66d, (double)method!.Invoke(null, [placement, false, false, false, 18d, true])!);
        TestAssert.Equal(3.0d, (double)method.Invoke(null, [placement, false, false, false, 18d, false])!);
    }

    public static void PptxSyntheticOverlayLegendVerticalShift()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeOverlayLegendVerticalShift",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected overlay shift to remain inspectable by the Office evidence guard.");

        // Office middles sit markerSize/2 below frame middle (5.01/5.03 at fs18, 6.64 at fs24).
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [9.9d, true])! - 4.95d) < 0.01d, "Expected fs18 overlay shift.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [13.2d, true])! - 6.6d) < 0.01d, "Expected fs24 overlay shift.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [9.9d, false])!) < 0.01d, "Expected no shift off-overlay.");
    }

    public static void PptxSyntheticHorizontalLegendPackingGaps()
    {
        var rendererType = typeof(PptxRenderer);
        var textGap = rendererType.GetMethod(
            "ComputeHorizontalLegendTextGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        var interGap = rendererType.GetMethod(
            "ComputeHorizontalLegendInterEntryGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(textGap is not null && interGap is not null, "Expected horizontal packing helpers to remain inspectable by the Office evidence guard.");

        // Swatch-text gap: legacy 3.0 off-bottom; content-independent line 0.275fs-0.35
        // (2.96/4.60/6.25 at 12/18/24pt) on bottom legends.
        TestAssert.True(System.Math.Abs((double)textGap!.Invoke(null, [12d, false])! - 3.0d) < 0.01d, "Expected legacy text gap.");
        TestAssert.True(System.Math.Abs((double)textGap.Invoke(null, [12.02d, true])! - 2.96d) < 0.05d, "Expected leg12 text gap.");
        TestAssert.True(System.Math.Abs((double)textGap.Invoke(null, [18d, true])! - 4.60d) < 0.05d, "Expected leg18 text gap.");
        TestAssert.True(System.Math.Abs((double)textGap.Invoke(null, [24d, true])! - 6.25d) < 0.05d, "Expected leg24 text gap.");

        // Inter-entry gap: legacy 8.0 off-bottom; 0.813fs-0.0276range+0.285 on bottom
        // legends (six Office knots within 0.031).
        TestAssert.True(System.Math.Abs((double)interGap!.Invoke(null, [18d, 24.19d, false])! - 8.0d) < 0.01d, "Expected legacy inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [12.02d, 16.18d, true])! - 9.60d) < 0.05d, "Expected leg12 inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [18d, 24.19d, true])! - 14.27d) < 0.05d, "Expected leg18 inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [24d, 32.30d, true])! - 18.90d) < 0.05d, "Expected leg24 inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [18d, 50.56d, true])! - 13.54d) < 0.05d, "Expected longnames inter gap.");
        TestAssert.True(System.Math.Abs((double)interGap.Invoke(null, [18d, 5.80d, true])! - 14.79d) < 0.05d, "Expected nsw inter gap.");
    }

    public static void PptxSyntheticNoTitleBottomLegendAnchoredTop()
    {
        var rendererType = typeof(PptxRenderer);
        var method = rendererType.GetMethod(
            "ResolveNoTitleBottomLegendAnchoredTop",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected bottom-legend anchored top to remain inspectable by the Office evidence guard.");
        var boxType = rendererType.GetNestedType("ChartFrameBox", System.Reflection.BindingFlags.NonPublic);
        TestAssert.True(boxType is not null, "Expected chart frame box to remain resolvable for the anchor pin.");

        // Office composite clips agree bit-identically base+tall (plot top 419.76
        // under plus-100H downward growth); the 0.949H ratio top falls 5.1 per 100H.
        object? baseFrame = System.Activator.CreateInstance(boxType!, 72d, 72d, 576d, 360d);
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [baseFrame])! - 419.76d) < 0.01d, "Expected base anchored top.");
        object? tallFrame = System.Activator.CreateInstance(boxType!, 72d, -28d, 576d, 460d);
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [tallFrame])! - 419.76d) < 0.01d, "Expected tall anchored top.");
    }

    public static void PptxSyntheticPerSeriesHorizontalLegendInterEntryGap()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputePerSeriesHorizontalLegendInterEntryGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected per-series inter law to remain inspectable by the Office evidence guard.");

        // Office per-series knots (fs, mean advance, inter): composite base plus
        // leg18 plus longnames plus shortnames plus NSW plus rangednames fit 0.45fs
        // plus 0.148avgW plus 0.195 within 0.062 with three spare DOF (wide ranges to 42 absorbed).
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [12d, 40.99d])! - 11.66d) < 0.1d, "Expected composite-base inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, 61.57d])! - 17.39d) < 0.1d, "Expected leg18 inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [12d, 85.40d])! - 18.24d) < 0.1d, "Expected longnames inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, 37.77d])! - 13.85d) < 0.1d, "Expected shortnames inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [18d, 43.47d])! - 14.79d) < 0.1d, "Expected nsw inter.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [12d, 40.40d])! - 11.55d) < 0.1d, "Expected rangednames inter (wide range absorbed).");
    }

    public static void PptxSyntheticBottomLegendFrameAnchorX()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ResolveBottomLegendFrameAnchorX",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected bottom-legend frame anchor to remain inspectable by the Office evidence guard.");

        // Office bottom-legend blocks sit at frame-center plus markerSize over 4 plus
        // 0.65 (twelve renders within 0.05); the val14 splitter proves frame anchoring
        // with the plot moved plus-6 and the block unmoved.
        TestAssert.True(System.Math.Abs((double)method!.Invoke(null, [72d, 576d, 174.82d, 6.57d])! - 274.87d) < 0.1d, "Expected composite anchored X.");
        TestAssert.True(System.Math.Abs((double)method.Invoke(null, [120d, 360d, 128.05d, 6.59d])! - 238.23d) < 0.1d, "Expected botleg anchored X.");
    }

    public static void PptxSyntheticChartNoTitleRightLegendLeftInsetKeepsOfficePadding()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeNoTitleRightLegendLeftInset",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart left-inset helper to remain inspectable by the Office evidence guard.");

        double singleDigit = (double)method!.Invoke(null, [9.12d, 720d])!;
        double threeDigit = (double)method.Invoke(null, [27.37d, 720d])!;
        double fourDigit = (double)method.Invoke(null, [36.49d, 720d])!;
        double narrowFrame = (double)method.Invoke(null, [9.12d, 400d])!;

        TestAssert.True(Math.Abs(singleDigit - 32.32d) < 0.0001d, "Single-digit inset should be label plus 23.2pt Office padding. Got " + singleDigit);
        TestAssert.True(Math.Abs(threeDigit - 50.57d) < 0.0001d, "Three-digit inset should use the same padding. Got " + threeDigit);
        TestAssert.True(Math.Abs(fourDigit - 59.69d) < 0.0001d, "Four-digit inset must not add a character-count extra. Got " + fourDigit);
        TestAssert.True(Math.Abs(narrowFrame - 29.12d) < 0.0001d, "Narrow frames keep the width-ratio padding cap. Got " + narrowFrame);
    }

    public static void PptxSyntheticChartRightLegendReserveKeepsOfficeTail()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeRightLegendReservePadding",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected chart right-reserve helper to remain inspectable by the Office evidence guard.");

        double lineScatter = (double)method!.Invoke(null, [18d, false, 0d, 0d])!;
        double smallFont = (double)method.Invoke(null, [9d, false, 0d, 0d])!;
        double scatterWideX = (double)method.Invoke(null, [18d, false, 0d, 18.25d])!;
        double areaWideCats = (double)method.Invoke(null, [18d, true, 36.56d, 0d])!;
        double areaNarrowCats = (double)method.Invoke(null, [18d, true, 21.25d, 0d])!;
        double areaNoCats = (double)method.Invoke(null, [18d, true, 0d, 0d])!;

        TestAssert.True(Math.Abs(lineScatter - 47.1d) < 0.001d, "Line/scatter tail should be marker block plus 10.8pt Office padding. Got " + lineScatter);
        TestAssert.True(Math.Abs(smallFont - 28.95d) < 0.001d, "Smaller legend fonts shrink the marker block while retaining the Office tail. Got " + smallFont);
        TestAssert.True(Math.Abs(scatterWideX - 55.425d) < 0.001d, "Scatter padding should add half the overhanging edge label with its own tail. Got " + scatterWideX);
        TestAssert.True(Math.Abs(areaWideCats - 60.33d) < 0.001d, "Area padding should be the 42.05pt block plus half the last category label. Got " + areaWideCats);
        TestAssert.True(Math.Abs(areaNarrowCats - 52.675d) < 0.001d, "Narrow edge labels should shrink the area padding. Got " + areaNarrowCats);
        TestAssert.True(Math.Abs(areaNoCats - 42.05d) < 0.001d, "Missing edge labels should leave the bare block. Got " + areaNoCats);
    }

    public static void PptxSyntheticBarLegendKeyUnitLeftCentersOnBar()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeyUnitLeft",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected bar legend-key unit helper to remain inspectable by the Office evidence guard.");
        double twoDigit = (double)method!.Invoke(null, [500.65d, 4.4d, 8.63d])!;
        double singleDigit = (double)method.Invoke(null, [200.25d, 4.4d, 4.2934d])!;
        TestAssert.True(Math.Abs(twoDigit - 493.385d) < 0.1d, "Two-digit swatches should start at the unit left (Office Delta 493.4). Got " + twoDigit);
        TestAssert.True(Math.Abs(singleDigit - 195.157d) < 0.1d, "Single-digit swatches should start at the unit left (Office Alpha 195.1). Got " + singleDigit);
    }

    public static void PptxSyntheticBarLegendKeyOutEndGapLiftsWithFontSize()
    {
        var method = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeyOutEndGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(method is not null, "Expected bar out-end gap helper to remain inspectable by the Office evidence guard.");
        double eight = (double)method!.Invoke(null, [8d])!;
        double twelve = (double)method!.Invoke(null, [12d])!;
        double sixteen = (double)method!.Invoke(null, [16d])!;
        TestAssert.True(Math.Abs(eight - 6.74d) < 0.05d, "Eight-point labels should lift 6.74 above bar tops (Office 6.66). Got " + eight);
        TestAssert.True(Math.Abs(twelve - 7.86d) < 0.05d, "Twelve-point labels should lift 7.86 above bar tops (Office 7.83). Got " + twelve);
        TestAssert.True(Math.Abs(sixteen - 8.98d) < 0.05d, "Sixteen-point labels should lift 8.98 above bar tops (Office 8.99). Got " + sixteen);
    }

    public static void PptxSyntheticBarLegendKeySwatchAnchorsToText()
    {
        var gapMethod = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeySwatchGap",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        var xMethod = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeySwatchX",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        var yMethod = typeof(PptxRenderer).GetMethod(
            "ComputeBarLegendKeySwatchY",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        TestAssert.True(gapMethod is not null && xMethod is not null && yMethod is not null, "Expected bar swatch helpers to remain inspectable by the Office evidence guard.");
        double gapEight = (double)gapMethod!.Invoke(null, [8d])!;
        double gapSixteen = (double)gapMethod!.Invoke(null, [16d])!;
        TestAssert.True(Math.Abs(gapEight - 4.5d) < 0.05d, "Eight-point swatch gap should read 4.50 (Office 4.47). Got " + gapEight);
        TestAssert.True(Math.Abs(gapSixteen - 6.63d) < 0.05d, "Sixteen-point swatch gap should read 6.63 (Office 6.64, kills G equals S). Got " + gapSixteen);
        double swatchXEight = (double)xMethod!.Invoke(null, [203.98d, 8.04d])!;
        double swatchYEight = (double)yMethod!.Invoke(null, [247.94d, 8.04d])!;
        TestAssert.True(Math.Abs(swatchXEight - 195.12d) < 0.15d, "Eight-point swatch should start at the Office left (195.12). Got " + swatchXEight);
        TestAssert.True(Math.Abs(swatchYEight - 248.54d) < 0.15d, "Eight-point swatch should sit at the Office top (248.54). Got " + swatchYEight);
        double swatchXSixteen = (double)xMethod!.Invoke(null, [334.97d, 15.96d])!;
        double swatchYSixteen = (double)yMethod!.Invoke(null, [271.25d, 15.96d])!;
        TestAssert.True(Math.Abs(swatchXSixteen - 319.54d) < 0.15d, "Sixteen-point swatch should start at the Office left (319.54). Got " + swatchXSixteen);
        TestAssert.True(Math.Abs(swatchYSixteen - 272.25d) < 0.15d, "Sixteen-point swatch should sit at the Office top (272.25). Got " + swatchYSixteen);
    }

    public static void PptxSyntheticDoughnutLegendConstantsKeepOfficeCalibration()
    {
        var rules = typeof(PptxRenderer).GetNestedType("PptxChartMetricRules", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected metric rules.");
        double clearance = (double)rules.GetField("DoughnutExplodedLegendClearance")!.GetValue(null)!;
        double shift = (double)rules.GetField("DoughnutRightLegendVerticalShift")!.GetValue(null)!;
        double leftShift = (double)rules.GetField("DoughnutLeftLegendVerticalShift")!.GetValue(null)!;
        TestAssert.True(Math.Abs(clearance - 86.3d) < 0.000001d, "Exploded doughnuts should keep the 86.3pt ring clearance (narrow/mid/wide/short renders). Got " + clearance);
        TestAssert.True(Math.Abs(shift - 22.92d) < 0.000001d, "Doughnut right legends should keep the 22.92pt vertical shift (8 Office renders). Got " + shift);
        TestAssert.True(Math.Abs(leftShift - 5.0d) < 0.000001d, "Doughnut left legends should keep the 5.0pt vertical shift (2 Office renders). Got " + leftShift);
        double headInset = (double)rules.GetField("DoughnutLeftLegendHeadInset")!.GetValue(null)!;
        TestAssert.True(Math.Abs(headInset - 13.19d) < 0.000001d, "Doughnut left legends should keep the 13.19pt head inset (4 Office renders). Got " + headInset);
        double lead = (double)rules.GetField("DoughnutLeftRingCenterLead")!.GetValue(null)!;
        TestAssert.True(Math.Abs(lead - 2.05d) < 0.000001d, "Left rings should keep the 2.05pt center lead (4 Office box widths). Got " + lead);
        double band = (double)rules.GetField("DoughnutTitledCenterYOffset")!.GetValue(null)!;
        TestAssert.True(Math.Abs(band - 17.68d) < 0.000001d, "Titled doughnuts should keep the 17.68pt center band (8 Office renders). Got " + band);
    }

    public static void PptxSyntheticDoughnutLeftGeometryFollowsLegendBox()
    {
        var renderer = typeof(PptxRenderer);
        var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static;
        var method = renderer.GetMethod("ResolvePieOrDoughnutLayout", flags);
        TestAssert.True(method is not null, "Expected polar layout helper to remain inspectable by the Office evidence guard.");
        var kindType = renderer.GetNestedType("ChartPolarKind", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected polar kind.");
        var plotBoxType = renderer.GetNestedType("ChartPlotBox", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected plot box.");
        var legendType = renderer.GetNestedType("ChartLegendLayout", System.Reflection.BindingFlags.NonPublic) ?? throw new InvalidOperationException("Expected legend layout.");
        var instanceFlags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance;
        object doughnut = System.Enum.ToObject(kindType, 1);
        object plotBox = System.Activator.CreateInstance(plotBoxType, [144d, 72d, 576d, 432d])!;
        object hidden = legendType.GetProperty("Hidden", flags | System.Reflection.BindingFlags.Public)!.GetValue(null)!;
        object manual = legendType.GetProperty("Layout", instanceFlags)!.GetValue(hidden)!;
        object textBody = legendType.GetProperty("TextBodyProperties", instanceFlags)!.GetValue(hidden)!;
        object shape = legendType.GetProperty("ShapeStyle", instanceFlags)!.GetValue(hidden)!;
        object leftVisible = System.Activator.CreateInstance(legendType, [PptxSceneChartLegendPosition.Left, "l", false, true, manual, textBody, shape])!;
        var empty = new Dictionary<int, double>();
        object narrow = method!.Invoke(null, [doughnut, plotBox, empty, leftVisible, false, false, 0d, 217.67d, false])!;
        object wide = method.Invoke(null, [doughnut, plotBox, empty, leftVisible, false, false, 0d, 296.57d, false])!;
        (double nx, double ny, double nr) = ReadPolarGeometry(narrow);
        (double wx, double wy, double wr) = ReadPolarGeometry(wide);
        TestAssert.True(Math.Abs((wx - nx) - (296.57d - 217.67d) / 2d) < 0.01d, "The ring must translate by half the box growth (1:2 rule). Got " + (wx - nx));
        TestAssert.True(Math.Abs(nr - 205.027d) < 0.05d, "Narrow left rings keep the baseline radius. Got " + nr);
        TestAssert.True(wr < nr, "Wide left rings must bind the frame margin, not keep baseline. Got " + wr + " vs " + nr);
    }
}
