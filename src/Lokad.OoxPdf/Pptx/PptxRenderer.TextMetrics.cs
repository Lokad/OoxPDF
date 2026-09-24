using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static double LineBaselineOffset(double fontSize, LineSpacing lineSpacing, bool useOfficeBaselineFloor, bool useExplicitMultipleBaselineOffset)
    {
        if (lineSpacing.IsAbsolute)
        {
            return Math.Max(BaselineOffset(fontSize), lineSpacing.Value - fontSize * PptxTextMetricRules.AbsoluteLineBaselineGapFallback);
        }

        return lineSpacing.IsExplicit && useExplicitMultipleBaselineOffset
            ? ReadExplicitMultipleBaselineOffset(lineSpacing, fontSize)
            : BaselineOffset(fontSize);
    }

    private static double LineBaselineOffset(
        double fontSize,
        LineSpacing lineSpacing,
        ResolvedRunTextStyle? style,
        TextAdvanceEstimator advanceEstimator,
        bool useOfficeBaselineFloor,
        bool useExplicitMultipleBaselineOffset)
    {
        if (lineSpacing.IsAbsolute)
        {
            return Math.Max(BaselineOffset(fontSize, style, advanceEstimator, useOfficeBaselineFloor: false, lineSpacing), lineSpacing.Value - fontSize * PptxTextMetricRules.AbsoluteLineBaselineGapFallback);
        }

        return lineSpacing.IsExplicit && useExplicitMultipleBaselineOffset
            ? ReadExplicitMultipleBaselineOffset(lineSpacing, fontSize)
            : BaselineOffset(fontSize, style, advanceEstimator, useOfficeBaselineFloor, lineSpacing);
    }

    private static double ManualBreakBaselineOffset(double fontSize, LineSpacing lineSpacing, bool useOfficeBaselineFloor, bool useExplicitMultipleBaselineOffset)
    {
        return lineSpacing.IsExplicit ? LineBaselineOffset(fontSize, lineSpacing, useOfficeBaselineFloor, useExplicitMultipleBaselineOffset) : fontSize * PptxTextMetricRules.OfficeManualBreakBaselineFallback;
    }

    private static bool ShouldUseExplicitMultipleBaselineOffset(PptxTextFrameModel frame, LineSpacing lineSpacing)
    {
        if (!lineSpacing.IsExplicit ||
            lineSpacing.IsAbsolute ||
            lineSpacing.Value <= 1d + PptxTextMetricRules.CoordinateTolerance)
        {
            return true;
        }

        return frame.ColumnCount <= 1 ||
            frame.BodyProperties.VerticalOverflow != PptxTextVerticalOverflow.Overflow ||
            !HasNoAutoFit(frame.BodyProperties);
    }

    private static double ReadExplicitMultipleBaselineOffset(LineSpacing lineSpacing, double fontSize)
    {
        double baseline = BaselineOffset(fontSize);
        if (lineSpacing.UseNormalLineAdvance &&
            lineSpacing.Value < 1d - PptxTextMetricRules.CoordinateTolerance)
        {
            double normalAdvance = fontSize * PptxTextMetricRules.CssNormalLineHeightFallback;
            double compressedAdvance = normalAdvance * lineSpacing.Value;
            double compressedBaseline = baseline - (normalAdvance - compressedAdvance);
            return Math.Max(fontSize * PptxTextMetricRules.MinimumBaselineMetricRatio, compressedBaseline);
        }

        return baseline * lineSpacing.Value;
    }

    private static double BaselineOffset(double fontSize)
    {
        return fontSize * PptxTextMetricRules.OfficeBaselineFallback;
    }

    // RV18: painting shares the inspection metric core so font-metric choice and
    // capping live once; the emitted value is the metric record value.
    private static double BaselineOffset(double fontSize, ResolvedRunTextStyle? style, TextAdvanceEstimator advanceEstimator, bool useOfficeBaselineFloor, LineSpacing lineSpacing)
    {
        return ReadBaselineMetric(fontSize, style, advanceEstimator, useOfficeBaselineFloor, lineSpacing).Value;
    }

    private static PptxTextBaselineMetricLayout ReadBaselineMetric(double fontSize, ResolvedRunTextStyle? style, TextAdvanceEstimator advanceEstimator, bool useOfficeBaselineFloor, LineSpacing lineSpacing)
    {
        const double fallbackRatio = PptxTextMetricRules.OfficeBaselineFallback;
        if (style is null)
        {
            return new PptxTextBaselineMetricLayout("Fallback", null, false, false, fontSize, fallbackRatio, 0, 0, 0, 0, 0, 0);
        }

        ResolvedRunTextStyle runStyle = style.Value;
        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(runStyle.Typeface, runStyle.Bold, runStyle.Italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return new PptxTextBaselineMetricLayout("Fallback", runStyle.Typeface, runStyle.Bold, runStyle.Italic, fontSize, fallbackRatio, 0, 0, 0, 0, 0, 0);
        }

        double ascenderRatio = font.Os2.WindowsAscender / (double)font.UnitsPerEm;
        string source;
        double ratio;
        if (ascenderRatio > 0d && ascenderRatio <= PptxTextMetricRules.MaximumBaselineMetricRatio)
        {
            ratio = ResolveOfficeBaselineMetricRatio(font, ascenderRatio, fontSize, out source);
        }
        else
        {
            ratio = fallbackRatio;
            source = "Fallback";
        }

        if (useOfficeBaselineFloor && TextMetricUsesOfficeBaselineFloor(font, runStyle, advanceEstimator, ascenderRatio))
        {
            ratio = Math.Max(fallbackRatio, ratio);
        }

        // Tall-font cap (Office probes 2026-09-06): only fonts whose Windows ascent exceeds the em box (Aptos Display 2068/2048) cap the first baseline at lineAdvance minus Windows descent; normal fonts (Calibri 1950/2048) keep winAscent and fallback metrics never cap.
        if (!lineSpacing.IsExplicit && fontSize > 0d && ascenderRatio > PptxTextMetricRules.MaximumOfficeBaselineWindowsAscenderRatio && ascenderRatio <= PptxTextMetricRules.MaximumBaselineMetricRatio)
        {
            double lineAdvance = ReadLineAdvance(lineSpacing, fontSize);
            double windowsDescender = font.Os2.WindowsDescender / (double)font.UnitsPerEm * fontSize;
            ratio = Math.Min(ratio, (lineAdvance - windowsDescender) / fontSize);
        }
        return new PptxTextBaselineMetricLayout(
            source,
            runStyle.Typeface,
            runStyle.Bold,
            runStyle.Italic,
            fontSize,
            ratio,
            font.UnitsPerEm,
            font.Os2.WindowsAscender,
            font.Os2.WindowsDescender,
            font.Os2.TypographicAscender,
            font.Os2.TypographicDescender,
            font.Os2.TypographicLineGap);
    }

    private static double ResolveOfficeBaselineMetricRatio(OpenTypeFont font, double windowsAscenderRatio, double fontSize, out string source)
    {
        double ratio = windowsAscenderRatio;
        source = "OS/2 usWinAscent";
        double typographicAscenderRatio = font.Os2.TypographicAscender / (double)font.UnitsPerEm;
        if (windowsAscenderRatio > PptxTextMetricRules.MaximumOfficeBaselineWindowsAscenderRatio &&
            fontSize <= PptxTextMetricRules.MaximumOfficeTypographicBaselineFontSize &&
            typographicAscenderRatio > 0d &&
            typographicAscenderRatio >= PptxTextMetricRules.MinimumOfficeTypographicBaselineAscenderRatio &&
            typographicAscenderRatio <= PptxTextMetricRules.MaximumBaselineMetricRatio)
        {
            ratio = typographicAscenderRatio;
            source = "OS/2 sTypoAscender";
        }

        return Math.Max(ratio, PptxTextMetricRules.MinimumBaselineMetricRatio);
    }

    private static bool TextMetricUsesOfficeBaselineFloor(OpenTypeFont font, ResolvedRunTextStyle runStyle, TextAdvanceEstimator advanceEstimator, double ascenderRatio)
    {
        return ascenderRatio <= 0d ||
            ascenderRatio > PptxTextMetricRules.MaximumBaselineMetricRatio ||
            ascenderRatio < PptxTextMetricRules.OfficeBaselineFloorMetricThreshold ||
            font.Os2.WindowsDescender / (double)font.UnitsPerEm <= PptxTextMetricRules.OfficeBaselineFloorMaximumWindowsDescenderRatio;
    }
}
