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
    private static PptxTextOrientation ParseTextOrientation(string? orientation)
    {
        return orientation switch
        {
            null or "" => PptxTextOrientation.Horizontal,
            "horz" => PptxTextOrientation.Horizontal,
            "vert" => PptxTextOrientation.Vertical,
            "vert270" => PptxTextOrientation.Vertical270,
            "eaVert" => PptxTextOrientation.EastAsianVertical,
            "mongolianVert" => PptxTextOrientation.MongolianVertical,
            "wordArtVert" => PptxTextOrientation.WordArtVertical,
            "wordArtVertRtl" => PptxTextOrientation.WordArtVerticalRightToLeft,
            _ when orientation.Equals("horz", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.Horizontal,
            _ when orientation.Equals("vert", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.Vertical,
            _ when orientation.Equals("vert270", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.Vertical270,
            _ when orientation.Equals("eaVert", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.EastAsianVertical,
            _ when orientation.Equals("mongolianVert", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.MongolianVertical,
            _ when orientation.Equals("wordArtVert", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.WordArtVertical,
            _ when orientation.Equals("wordArtVertRtl", StringComparison.OrdinalIgnoreCase) => PptxTextOrientation.WordArtVerticalRightToLeft,
            _ => PptxTextOrientation.Unknown
        };
    }

    private static double TextOrientationRotationDegrees(PptxTextOrientation orientation)
    {
        return orientation switch
        {
            PptxTextOrientation.Vertical270 => 270d,
            PptxTextOrientation.Vertical or
            PptxTextOrientation.EastAsianVertical or
            PptxTextOrientation.MongolianVertical or
            PptxTextOrientation.WordArtVertical or
            PptxTextOrientation.WordArtVerticalRightToLeft => 90d,
            _ => 0d
        };
    }

    private static double? ParseTextBodyRotationDegrees(string? rotation)
    {
        return rotation is not null &&
            long.TryParse(rotation, NumberStyles.Integer, CultureInfo.InvariantCulture, out long rotationValue)
                ? rotationValue / 60000d
                : null;
    }

    private static (
        int Count,
        double Spacing,
        PptxTextBodyPropertySource CountSource,
        PptxTextBodyPropertySource SpacingSource,
        string? CountValue,
        string? SpacingValue) ReadTextColumns(XElement textBody, XElement? inheritedTextBody)
    {
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        XElement? inheritedBodyProperties = inheritedTextBody?.Element(DrawingNamespace + "bodyPr");
        (int count, PptxTextBodyPropertySource countSource, string? countValue) = ReadTextColumnCount(bodyProperties, inheritedBodyProperties);
        (double spacing, PptxTextBodyPropertySource spacingSource, string? spacingValue) = ReadTextColumnSpacing(bodyProperties, inheritedBodyProperties);
        return (count, spacing, countSource, spacingSource, countValue, spacingValue);
    }

    private static (int Count, PptxTextBodyPropertySource Source, string? Value) ReadTextColumnCount(
        XElement? bodyProperties,
        XElement? inheritedBodyProperties)
    {
        if (bodyProperties?.Attribute("numCol") is { } directAttribute)
        {
            if (int.TryParse(directAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int directCount))
            {
                return (Math.Clamp(directCount, 1, 16), PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
            }

            return (1, PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
        }

        if (inheritedBodyProperties?.Attribute("numCol") is { } inheritedAttribute)
        {
            if (int.TryParse(inheritedAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int inheritedCount))
            {
                return (Math.Clamp(inheritedCount, 1, 16), PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
            }

            return (1, PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
        }

        return (1, PptxTextBodyPropertySource.DefaultValue, null);
    }

    private static (double Spacing, PptxTextBodyPropertySource Source, string? Value) ReadTextColumnSpacing(
        XElement? bodyProperties,
        XElement? inheritedBodyProperties)
    {
        if (bodyProperties?.Attribute("spcCol") is { } directAttribute)
        {
            if (long.TryParse(directAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long directSpacing))
            {
                return (Math.Max(0d, OoxUnits.EmuToPoints(directSpacing)), PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
            }

            return (0d, PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
        }

        if (inheritedBodyProperties?.Attribute("spcCol") is { } inheritedAttribute)
        {
            if (long.TryParse(inheritedAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long inheritedSpacing))
            {
                return (Math.Max(0d, OoxUnits.EmuToPoints(inheritedSpacing)), PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
            }

            return (0d, PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
        }

        return (0d, PptxTextBodyPropertySource.DefaultValue, null);
    }

    private static double NormalizeRotationDegrees(double rotationDegrees)
    {
        double normalized = rotationDegrees % 360d;
        return normalized < 0d ? normalized + 360d : normalized;
    }

    private static (double Scale, PptxTextBodyPropertySource Source, string? Value) ReadNormAutofitFontScale(
        XElement? autofit,
        PptxTextBodyPropertySource autofitSource)
    {
        if (autofit?.Name.LocalName != "normAutofit" ||
            autofit.Attribute("fontScale") is not { } fontScale)
        {
            return (1d, PptxTextBodyPropertySource.DefaultValue, null);
        }

        if (!int.TryParse(fontScale.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fontScaleValue))
        {
            return (1d, autofitSource, fontScale.Value);
        }

        double scale = Math.Clamp(
            fontScaleValue / 100000d,
            PptxTextMetricRules.MinimumAutofitScale,
            PptxTextMetricRules.MaximumAutofitScale);
        return (scale, autofitSource, fontScale.Value);
    }

    private static (double Scale, PptxTextBodyPropertySource Source, string? Value) ReadNormAutofitLineSpacingScale(
        XElement? autofit,
        PptxTextBodyPropertySource autofitSource)
    {
        if (autofit?.Name.LocalName != "normAutofit" ||
            autofit.Attribute("lnSpcReduction") is not { } reduction)
        {
            return (1d, PptxTextBodyPropertySource.DefaultValue, null);
        }

        if (!int.TryParse(reduction.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int reductionValue))
        {
            return (1d, autofitSource, reduction.Value);
        }

        double reductionRatio = Math.Clamp(reductionValue / 100000d, 0d, PptxTextMetricRules.MaximumLineSpacingReduction);
        return (1d - reductionRatio, autofitSource, reduction.Value);
    }

    private static PptxTextWrapMode ParseTextWrapMode(string? wrap)
    {
        return wrap switch
        {
            null or "" => PptxTextWrapMode.Square,
            "none" => PptxTextWrapMode.None,
            "square" => PptxTextWrapMode.Square,
            _ when wrap.Equals("none", StringComparison.OrdinalIgnoreCase) => PptxTextWrapMode.None,
            _ when wrap.Equals("square", StringComparison.OrdinalIgnoreCase) => PptxTextWrapMode.Square,
            _ => PptxTextWrapMode.Unknown
        };
    }

    private static PptxTextVerticalOverflow ParseTextVerticalOverflow(string? overflow)
    {
        return overflow switch
        {
            "clip" => PptxTextVerticalOverflow.Clip,
            "ellipsis" => PptxTextVerticalOverflow.Ellipsis,
            _ when overflow?.Equals("clip", StringComparison.OrdinalIgnoreCase) == true => PptxTextVerticalOverflow.Clip,
            _ when overflow?.Equals("ellipsis", StringComparison.OrdinalIgnoreCase) == true => PptxTextVerticalOverflow.Ellipsis,
            _ when overflow?.Equals("overflow", StringComparison.OrdinalIgnoreCase) == true => PptxTextVerticalOverflow.Overflow,
            _ when !string.IsNullOrEmpty(overflow) => PptxTextVerticalOverflow.Unknown,
            _ => PptxTextVerticalOverflow.Overflow
        };
    }

    private static (double Value, PptxTextBodyPropertySource Source, string? RawValue) ReadInset(
        XElement? bodyProperties,
        XElement? inheritedBodyProperties,
        string attributeName,
        long defaultEmu)
    {
        if (bodyProperties?.Attribute(attributeName) is { } directAttribute)
        {
            if (long.TryParse(directAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu))
            {
                return (OoxUnits.EmuToPoints(emu), PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
            }

            return (OoxUnits.EmuToPoints(defaultEmu), PptxTextBodyPropertySource.DirectBodyPr, directAttribute.Value);
        }

        if (inheritedBodyProperties?.Attribute(attributeName) is { } inheritedAttribute)
        {
            if (long.TryParse(inheritedAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu))
            {
                return (OoxUnits.EmuToPoints(emu), PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
            }

            return (OoxUnits.EmuToPoints(defaultEmu), PptxTextBodyPropertySource.InheritedBodyPr, inheritedAttribute.Value);
        }

        return (OoxUnits.EmuToPoints(defaultEmu), PptxTextBodyPropertySource.DefaultValue, null);
    }

    private static double ReadParagraphSpacing(
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        string elementName,
        double referenceFontSize)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + elementName) ??
            defaultParagraphProperties?.Element(DrawingNamespace + elementName);
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d;
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return referenceFontSize * PptxTextMetricRules.ClampNonNegative(int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d);
        }

        return 0d;
    }

    private static ParagraphIndent ReadParagraphIndent(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        return new ParagraphIndent(
            ReadParagraphEmuAttribute(paragraphProperties, defaultParagraphProperties, "marL"),
            ReadParagraphEmuAttribute(paragraphProperties, defaultParagraphProperties, "indent"));
    }

    private static double ReadParagraphEmuAttribute(XElement? paragraphProperties, XElement? defaultParagraphProperties, string attributeName)
    {
        return paragraphProperties?.Attribute(attributeName) is { } attribute
            ? OoxUnits.EmuToPoints(long.Parse(attribute.Value, CultureInfo.InvariantCulture))
            : defaultParagraphProperties?.Attribute(attributeName) is { } defaultAttribute
                ? OoxUnits.EmuToPoints(long.Parse(defaultAttribute.Value, CultureInfo.InvariantCulture))
                : 0d;
    }

    private static IReadOnlyList<double> ReadTabStops(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        XElement? tabList = paragraphProperties?.Element(DrawingNamespace + "tabLst") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "tabLst");
        if (tabList is null)
        {
            return Array.Empty<double>();
        }

        return tabList
            .Elements(DrawingNamespace + "tab")
            .Select(tab => tab.Attribute("pos") is { } position
                ? OoxUnits.EmuToPoints(long.Parse(position.Value, CultureInfo.InvariantCulture))
                : double.NaN)
            .Where(position => !double.IsNaN(position))
            .Order()
            .ToArray();
    }

    private static double ResolveNextTabX(double cursorX, double paragraphTextX, IReadOnlyList<double> tabStops)
    {
        double current = cursorX - paragraphTextX;
        foreach (double tabStop in tabStops)
        {
            if (tabStop > current + PptxTextMetricRules.CoordinateTolerance)
            {
                return paragraphTextX + tabStop;
            }
        }

        const long defaultTabStopEmus = 914400;
        double defaultTabStop = OoxUnits.EmuToPoints(defaultTabStopEmus);
        return paragraphTextX + Math.Ceiling((current + PptxTextMetricRules.CoordinateTolerance) / defaultTabStop) * defaultTabStop;
    }

    private static LineSpacing ReadLineSpacing(XElement? paragraphProperties, XElement? defaultParagraphProperties)
    {
        XElement? spacing = paragraphProperties?.Element(DrawingNamespace + "lnSpc") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "lnSpc");
        if (spacing?.Element(DrawingNamespace + "spcPts")?.Attribute("val") is { } points)
        {
            return LineSpacing.Absolute(Math.Max(PptxTextMetricRules.MinimumLineSpacing, int.Parse(points.Value, CultureInfo.InvariantCulture) / 100d));
        }

        if (spacing?.Element(DrawingNamespace + "spcPct")?.Attribute("val") is { } percent)
        {
            return LineSpacing.Multiple(Math.Max(PptxTextMetricRules.MinimumLineSpacing, int.Parse(percent.Value, CultureInfo.InvariantCulture) / 100000d), true, true);
        }

        return LineSpacing.Multiple(1d, false, true);
    }

    private static LineSpacing ApplyCompatibleLineSpacing(LineSpacing lineSpacing, bool compatibleLineSpacing, double defaultLineSpacingFactor)
    {
        return compatibleLineSpacing && !lineSpacing.IsExplicit
            ? LineSpacing.Multiple(defaultLineSpacingFactor, isExplicit: true, useNormalLineAdvance: false)
            : lineSpacing;
    }

    private static double ReadParagraphAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return ReadLineAdvance(lineSpacing, fontSize);
    }

    private static double ReadLineAdvance(LineSpacing lineSpacing, double fontSize)
    {
        if (lineSpacing.IsAbsolute)
        {
            return lineSpacing.Resolve(fontSize);
        }

        double normalAdvance = fontSize * PptxTextMetricRules.CssNormalLineHeightFallback;
        return lineSpacing.IsExplicit
            ? (lineSpacing.UseNormalLineAdvance ? normalAdvance : fontSize) * lineSpacing.Value
            : normalAdvance;
    }

    private static double ReadFirstLineBaselineOffset(
        PptxTextParagraphModel paragraph,
        LineSpacing lineSpacing,
        TextAdvanceEstimator advanceEstimator,
        bool useOfficeBaselineFloor,
        bool useExplicitMultipleBaselineOffset)
    {
        bool startsWithManualLineBreak = paragraph.Runs.FirstOrDefault()?.Kind == PptxTextRunKind.Break;
        PptxTextRunModel? firstRun = paragraph.Runs.FirstOrDefault(run => run.Kind != PptxTextRunKind.Break);
        double fontSize = firstRun?.Style.NominalFontSize ?? paragraph.FirstLineFallbackFontSize;
        return startsWithManualLineBreak || paragraph.HasManualLineBreak
            ? ManualBreakBaselineOffset(fontSize, lineSpacing, useOfficeBaselineFloor, useExplicitMultipleBaselineOffset)
            : LineBaselineOffset(fontSize, lineSpacing, firstRun?.Style, advanceEstimator, useOfficeBaselineFloor, useExplicitMultipleBaselineOffset);
    }

    private static double ReadManualBreakLineAdvance(LineSpacing lineSpacing, double fontSize)
    {
        return lineSpacing.IsExplicit ? ReadLineAdvance(lineSpacing, fontSize) : fontSize * PptxTextMetricRules.OfficeManualBreakDefaultLineHeightFallback;
    }

    private static double ReadFirstParagraphFontSize(XElement paragraph, XElement? defaultRunProperties)
    {
        const double defaultFontSize = 18d;
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == DrawingNamespace + "br")
            {
                return defaultFontSize;
            }

            if (!IsTextRunElement(child))
            {
                continue;
            }

            XElement? runProperties = child.Element(DrawingNamespace + "rPr");
            if (runProperties?.Attribute("sz") is { } size)
            {
                return int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d;
            }

            if (defaultRunProperties?.Attribute("sz") is { } defaultSize)
            {
                return int.Parse(defaultSize.Value, CultureInfo.InvariantCulture) / 100d;
            }

            return defaultFontSize;
        }

        return defaultFontSize;
    }

    private static string? ReadBulletText(PptxParagraphBulletModel bullet, ref int autoNumberValue)
    {
        if (bullet.Kind == PptxParagraphBulletKind.None || bullet.Kind == PptxParagraphBulletKind.Blip)
        {
            return null;
        }

        if (bullet.Kind == PptxParagraphBulletKind.Character)
        {
            return bullet.ResolvedCharacter;
        }

        if (bullet.Kind != PptxParagraphBulletKind.AutoNumber)
        {
            return null;
        }

        if (bullet.AutoNumberStartAt is { } start)
        {
            autoNumberValue = start;
        }

        string result = FormatAutoNumber(autoNumberValue, bullet.AutoNumberType);
        autoNumberValue++;
        return result;

        string FormatAutoNumber(int value, string? type)
        {
            return type switch
            {
                "arabicParenBoth" => $"({value})",
                "arabicParenR" => $"{value})",
                "alphaLcPeriod" => $"{FormatAlphaNumber(value, upper: false)}.",
                "alphaUcPeriod" => $"{FormatAlphaNumber(value, upper: true)}.",
                "alphaLcParenR" => $"{FormatAlphaNumber(value, upper: false)})",
                "alphaUcParenR" => $"{FormatAlphaNumber(value, upper: true)})",
                "romanLcPeriod" => $"{FormatRomanNumber(value, upper: false)}.",
                "romanUcPeriod" => $"{FormatRomanNumber(value, upper: true)}.",
                "romanLcParenR" => $"{FormatRomanNumber(value, upper: false)})",
                "romanUcParenR" => $"{FormatRomanNumber(value, upper: true)})",
                _ => $"{value}."
            };
        }
    }

    private static bool IsSymbolBulletFont(XElement? bulletFont)
    {
        string? charset = (string?)bulletFont?.Attribute("charset");
        return charset is "2" or "-2";
    }

    private static string MapSymbolBulletText(string bullet)
    {
        Span<char> mapped = bullet.Length <= 256
            ? stackalloc char[bullet.Length]
            : new char[bullet.Length];
        for (int i = 0; i < bullet.Length; i++)
        {
            char ch = bullet[i];
            mapped[i] = ch <= 0x00FF
                ? (char)(0xF000 + ch)
                : ch;
        }

        return new string(mapped);
    }

    private static string FormatAlphaNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        return OoxNumbering.ToAlphabetic(value, upper);
    }

    private static string FormatRomanNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return string.Empty;
        }

        string result = OoxNumbering.ToRomanUpper(value);
        return upper ? result : result.ToLowerInvariant();
    }

    private static BulletStyle ReadBulletStyle(PptxParagraphBulletModel bullet, double textFontSize, RgbColor textColor, string? textTypeface)
    {
        RgbColor color = bullet.Color ?? textColor;
        double fontSize = textFontSize;
        if (bullet.SizeKind == PptxParagraphBulletSizeKind.Percent &&
            int.TryParse(bullet.SizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizePercent))
        {
            fontSize = textFontSize * Math.Max(0.1d, sizePercent / 100000d);
        }
        else if (bullet.SizeKind == PptxParagraphBulletSizeKind.Points &&
            int.TryParse(bullet.SizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizePoints))
        {
            fontSize = Math.Max(0.1d, sizePoints / 100d);
        }

        return new BulletStyle(fontSize, color, bullet.ResolvedFontTypeface ?? textTypeface);
    }

    private static bool TryReadHighlightColor(XElement? runProperties, out RgbColor color)
    {
        XElement? highlight = runProperties?.Element(DrawingNamespace + "highlight");
        string? hex = (string?)highlight?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        return RgbColor.TryParse(hex, out color);
    }

    private static TextVerticalAnchor ParseTextVerticalAnchor(string? anchor)
    {
        return anchor switch
        {
            null or "" => TextVerticalAnchor.Top,
            "t" => TextVerticalAnchor.Top,
            "ctr" => TextVerticalAnchor.Middle,
            "b" => TextVerticalAnchor.Bottom,
            _ when anchor.Equals("t", StringComparison.OrdinalIgnoreCase) => TextVerticalAnchor.Top,
            _ when anchor?.Equals("ctr", StringComparison.OrdinalIgnoreCase) == true => TextVerticalAnchor.Middle,
            _ when anchor?.Equals("b", StringComparison.OrdinalIgnoreCase) == true => TextVerticalAnchor.Bottom,
            _ when !string.IsNullOrEmpty(anchor) => TextVerticalAnchor.Unknown,
            _ => TextVerticalAnchor.Top
        };
    }

    private static double EstimateTextHeight(
        IReadOnlyList<PptxTextParagraphModel> paragraphs,
        double textWidth,
        PptxTextBodyProperties bodyProperties)
    {
        double height = 0d;
        var advanceEstimator = new TextAdvanceEstimator(null, CancellationToken.None);
        bool allowWrapping = TextBodyAllowsWrapping(bodyProperties);
        bool attachSpacesToFollowingWord = HasNoAutoFit(bodyProperties);
        bool useWindowsFontBoxForDefaultLineSpacing = !IsTableCellVerticalAnchorSource(bodyProperties.VerticalAnchorSource);
        bool hasEstimatedParagraph = false;
        double pendingSpacingAfter = 0d;
        foreach (PptxTextParagraphModel paragraph in paragraphs)
        {
            ResolvedParagraphTextStyle paragraphStyle = paragraph.Style;
            LineSpacing lineSpacing = paragraphStyle.LineSpacing;
            if (!paragraph.HasVisibleContent)
            {
                if (paragraph.HasLayoutContent)
                {
                    double emptyFontSize = paragraph.EndParagraphProperties is null
                        ? paragraphStyle.FontSize
                        : paragraph.EndParagraphStyle.FontSize;
                    if (hasEstimatedParagraph)
                    {
                        height += pendingSpacingAfter + paragraph.EmptySpacingBefore;
                    }

                    string? emptyTypeface = paragraph.EndParagraphStyle.Typeface;
                    bool emptyBold = paragraph.EndParagraphStyle.Bold;
                    bool emptyItalic = paragraph.EndParagraphStyle.Italic;
                    height += ReadEstimatedAnchorEmptyLineAdvance(lineSpacing, emptyFontSize, emptyTypeface, emptyBold, emptyItalic, useWindowsFontBoxForDefaultLineSpacing, advanceEstimator);
                    pendingSpacingAfter = paragraph.EmptySpacingAfter;
                    hasEstimatedParagraph = true;
                }

                continue;
            }

            double paragraphFontSize = paragraphStyle.FontSize;
            if (hasEstimatedParagraph)
            {
                height += pendingSpacingAfter + paragraphStyle.SpacingBefore;
            }

            PptxTextFlowParagraph flowParagraph = BuildTextFlowParagraph(paragraph, attachSpacesToFollowingWord, advanceEstimator);
            double maxFontSize = 0d;
            string? lineTypeface = null;
            bool lineBold = false;
            bool lineItalic = false;
            double lineWidth = 0d;
            bool hasLineContent = false;
            foreach (PptxTextFlowRun flowRun in flowParagraph.Runs)
            {
                ResolvedRunTextStyle runStyle = flowRun.Style;
                if (flowRun.Source.Kind == PptxTextRunKind.Break)
                {
                    height += ReadEstimatedAnchorLineAdvance(lineSpacing, ResolveLineFontSize(maxFontSize, paragraphFontSize), lineTypeface, lineBold, lineItalic, bodyProperties, advanceEstimator);
                    maxFontSize = 0d;
                    lineTypeface = null;
                    lineBold = false;
                    lineItalic = false;
                    lineWidth = 0d;
                    hasLineContent = false;
                    continue;
                }

                foreach (PptxTextFlowSegment segment in flowRun.Segments)
                {
                    if (segment.Kind == PptxTextFlowSegmentKind.Break)
                    {
                        continue;
                    }

                    double fontSize = runStyle.FontSize * segment.FontScale;
                    double advance = advanceEstimator.Measure(
                        segment.AdvanceText,
                        fontSize,
                        runStyle.Typeface,
                        runStyle.Bold,
                        runStyle.Italic,
                        runStyle.CharacterSpacing,
                        runStyle.KerningEnabled);
                    if (allowWrapping &&
                        !string.IsNullOrWhiteSpace(segment.AdvanceText) &&
                        lineWidth > PptxTextMetricRules.TextStateTolerance &&
                        lineWidth + advance > textWidth)
                    {
                        height += ReadEstimatedAnchorLineAdvance(lineSpacing, ResolveLineFontSize(maxFontSize, paragraphFontSize), lineTypeface, lineBold, lineItalic, bodyProperties, advanceEstimator);
                        maxFontSize = fontSize;
                        lineTypeface = runStyle.Typeface;
                        lineBold = runStyle.Bold;
                        lineItalic = runStyle.Italic;
                        lineWidth = 0d;
                        hasLineContent = false;
                    }

                    if (string.IsNullOrWhiteSpace(segment.AdvanceText) && lineWidth <= PptxTextMetricRules.TextStateTolerance)
                    {
                        continue;
                    }

                    if (fontSize >= maxFontSize)
                    {
                        maxFontSize = fontSize;
                        lineTypeface = runStyle.Typeface;
                        lineBold = runStyle.Bold;
                        lineItalic = runStyle.Italic;
                    }

                    lineWidth += advance;
                    hasLineContent |= segment.Draw && segment.Text.Length > 0;
                }
            }

            if (hasLineContent || maxFontSize > PptxTextMetricRules.TextStateTolerance)
            {
                height += ReadEstimatedAnchorLineAdvance(lineSpacing, ResolveLineFontSize(maxFontSize, paragraphFontSize), lineTypeface, lineBold, lineItalic, bodyProperties, advanceEstimator);
            }

            pendingSpacingAfter = paragraphStyle.SpacingAfter;
            hasEstimatedParagraph = true;
        }

        return height;
    }

    private static bool IsTableCellVerticalAnchorSource(PptxTextBodyPropertySource source)
    {
        return source is PptxTextBodyPropertySource.TableCellProperties or PptxTextBodyPropertySource.TableCellStyle;
    }

    private static double ReadEstimatedAnchorLineAdvance(
        LineSpacing lineSpacing,
        double fontSize,
        string? typeface,
        bool bold,
        bool italic,
        PptxTextBodyProperties bodyProperties,
        TextAdvanceEstimator advanceEstimator)
    {
        if (lineSpacing.IsExplicit)
        {
            return ReadLineAdvance(lineSpacing, fontSize);
        }
        // Office centers absent-autofit content by line advances (North-clone plus ellipse probes 2026-09-06); explicit modes keep the font-box estimate below.
        if (HasAbsentAutofit(bodyProperties) && !IsTableCellVerticalAnchorSource(bodyProperties.VerticalAnchorSource))
        {
            return ReadLineAdvance(lineSpacing, fontSize);
        }

        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(typeface, bold, italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return ReadLineAdvance(lineSpacing, fontSize);
        }

        bool useWindowsFontBoxForDefaultLineSpacing = !IsTableCellVerticalAnchorSource(bodyProperties.VerticalAnchorSource);
        double metricUnits = useWindowsFontBoxForDefaultLineSpacing
            ? font.Os2.WindowsAscender + font.Os2.WindowsDescender
            : font.Os2.TypographicAscender - font.Os2.TypographicDescender + font.Os2.TypographicLineGap;
        double metricRatio = metricUnits / font.UnitsPerEm;
        if (metricRatio <= PptxTextMetricRules.MinimumFontLineBoxMetricRatio ||
            metricRatio > PptxTextMetricRules.MaximumFontLineBoxMetricRatio)
        {
            return ReadLineAdvance(lineSpacing, fontSize);
        }

        double metricAdvance = fontSize * metricRatio;
        double windowsAscenderRatio = font.Os2.WindowsAscender / (double)font.UnitsPerEm;
        if (useWindowsFontBoxForDefaultLineSpacing)
        {
            return Math.Max(ReadLineAdvance(lineSpacing, fontSize), metricAdvance);
        }

        if (!useWindowsFontBoxForDefaultLineSpacing &&
            windowsAscenderRatio > PptxTextMetricRules.MaximumOfficeBaselineWindowsAscenderRatio &&
            metricRatio <= PptxTextMetricRules.MaximumTableAnchorCompressedFontBoxRatio)
        {
            // Office does not let pathological compressed typographic boxes over-center table-cell text.
            return Math.Max(metricAdvance, ReadLineAdvance(lineSpacing, fontSize));
        }

        return metricAdvance;
    }

    private static double ReadEstimatedAnchorEmptyLineAdvance(
        LineSpacing lineSpacing,
        double fontSize,
        string? typeface,
        bool bold,
        bool italic,
        bool useWindowsFontBoxForDefaultLineSpacing,
        TextAdvanceEstimator advanceEstimator)
    {
        double paragraphAdvance = ReadLineAdvance(lineSpacing, fontSize);
        if (lineSpacing.IsExplicit || useWindowsFontBoxForDefaultLineSpacing)
        {
            return paragraphAdvance;
        }

        OpenTypeFont? font = advanceEstimator.ResolveOpenTypeFont(typeface, bold, italic);
        if (font is null || font.UnitsPerEm == 0)
        {
            return paragraphAdvance;
        }

        double metricUnits = font.Os2.WindowsAscender + font.Os2.WindowsDescender;
        double metricRatio = metricUnits / font.UnitsPerEm;
        if (metricRatio <= PptxTextMetricRules.MinimumFontLineBoxMetricRatio ||
            metricRatio > PptxTextMetricRules.MaximumFontLineBoxMetricRatio)
        {
            return paragraphAdvance;
        }

        double fontBoxAdvance = fontSize * metricRatio;
        return (paragraphAdvance + fontBoxAdvance) / 2d;
    }
}
