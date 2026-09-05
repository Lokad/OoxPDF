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
    private static XElement? MergeParagraphProperties(params XElement?[] sources)
    {
        return PptxParagraphPropertyMerger.MergeRendererDefaultProperties(DrawingNamespace + "defRPr", sources);
    }

    private static ResolvedParagraphTextStyle ResolveParagraphTextStyle(
        XElement paragraph,
        XElement? paragraphProperties,
        XElement? defaultParagraphProperties,
        double fontScale,
        double lineSpacingScale,
        bool compatibleLineSpacing,
        double compatibleDefaultLineSpacingFactor)
    {
        XElement? defaultRunProperties = paragraphProperties?.Element(DrawingNamespace + "defRPr") ??
            defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
        double fontSize = ReadFirstParagraphFontSize(paragraph, defaultRunProperties) * fontScale;
        string? alignmentValue = ReadAlignmentValue(paragraph, defaultParagraphProperties);
        return new ResolvedParagraphTextStyle(
            ParseAlignment(alignmentValue),
            alignmentValue,
            paragraphProperties,
            defaultRunProperties,
            fontSize,
            ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcBef", fontSize),
            ReadParagraphSpacing(paragraphProperties, defaultParagraphProperties, "spcAft", fontSize),
            ApplyCompatibleLineSpacing(
                ReadLineSpacing(paragraphProperties, defaultParagraphProperties),
                compatibleLineSpacing && HasManualLineBreak(paragraph) && !HasExplicitParagraphSpacing(paragraphProperties),
                compatibleDefaultLineSpacingFactor).ScaleExplicit(lineSpacingScale),
            ReadParagraphIndent(paragraphProperties, defaultParagraphProperties),
            ReadTabStops(paragraphProperties, defaultParagraphProperties));
    }

    private static bool HasManualLineBreak(XElement paragraph)
    {
        return paragraph.Elements(DrawingNamespace + "br").Any() ||
            paragraph.Elements(DrawingNamespace + "r")
                .Elements(DrawingNamespace + "t")
                .Any(text => TextContainsManualLineBreak(text.Value));
    }

    private static bool HasExplicitParagraphSpacing(XElement? paragraphProperties)
    {
        return paragraphProperties?.Element(DrawingNamespace + "spcBef") is not null ||
            paragraphProperties?.Element(DrawingNamespace + "spcAft") is not null;
    }

    private static ResolvedRunTextStyle ResolveRunTextStyle(
        PptxRunStyleCascade cascade,
        RgbColor? shapeFontColor,
        PptxTheme theme,
        PptxColorMap colorMap,
        double fontScale,
        PptxSceneTableCellTextStyle tableStyleTextStyle)
    {
        return ResolveRunTextStyle(
            cascade.DirectProperties,
            cascade.ResolvedDefaultProperties,
            shapeFontColor,
            theme,
            colorMap,
            fontScale,
            tableStyleTextStyle);
    }

    private static ResolvedRunTextStyle ResolveRunTextStyle(
        XElement? runProperties,
        XElement? defaultRunProperties,
        RgbColor? shapeFontColor,
        PptxTheme theme,
        PptxColorMap colorMap,
        double fontScale,
        PptxSceneTableCellTextStyle tableStyleTextStyle)
    {
        double nominalFontSize = ReadFontSize(runProperties, defaultRunProperties) * fontScale;
        double baselineOffset = ReadBaselineOffset(runProperties, defaultRunProperties, nominalFontSize);
        double fontSize = PptxTextMetricRules.ShouldScaleSuperscriptSubscript(baselineOffset, nominalFontSize)
            ? PptxTextMetricRules.SuperscriptSubscriptFontSize(nominalFontSize)
            : nominalFontSize;
        double alpha = 1d;
        RgbColor color;
        PptxRunTextColorSource colorSource;
        if (HasTextNoFill(runProperties))
        {
            color = new RgbColor(0, 0, 0);
            colorSource = PptxRunTextColorSource.RunNoFill;
            alpha = 0d;
        }
        else if (HasHyperlinkClick(runProperties) && theme.TryResolveColor("hlink", colorMap, out RgbColor hyperlinkColor))
        {
            color = hyperlinkColor;
            colorSource = PptxRunTextColorSource.ThemeHyperlink;
        }
        else if (TryReadSolidColorWithAlpha(runProperties, theme, colorMap, out RgbColor runColor, out double runAlpha))
        {
            color = runColor;
            colorSource = PptxRunTextColorSource.RunSolidFill;
            alpha = runAlpha;
        }
        else if (tableStyleTextStyle.Color is { } tableTextColor && !HasTextFill())
        {
            color = tableTextColor;
            colorSource = PptxRunTextColorSource.TableTextStyle;
        }
        else if (shapeFontColor is { } fontRefColor)
        {
            color = fontRefColor;
            colorSource = PptxRunTextColorSource.ShapeFontRef;
        }
        else if (HasTextNoFill(defaultRunProperties))
        {
            color = new RgbColor(0, 0, 0);
            colorSource = PptxRunTextColorSource.DefaultNoFill;
            alpha = 0d;
        }
        else if (TryReadSolidColorWithAlpha(defaultRunProperties, theme, colorMap, out RgbColor defaultColor, out double defaultAlpha))
        {
            color = defaultColor;
            colorSource = PptxRunTextColorSource.DefaultSolidFill;
            alpha = defaultAlpha;
        }
        else
        {
            color = new RgbColor(0, 0, 0);
            colorSource = PptxRunTextColorSource.FallbackBlack;
        }

        PptxThemeTypefaceResolution typeface = ReadRunTypeface(runProperties, defaultRunProperties, theme);
        bool bold = OoxXml.ParseOptionalBool(runProperties, "b") ||
            (runProperties?.Attribute("b") is null && tableStyleTextStyle.Bold) ||
            (runProperties?.Attribute("b") is null && OoxXml.ParseOptionalBool(defaultRunProperties, "b"));
        bool italic = OoxXml.ParseOptionalBool(runProperties, "i") ||
            (runProperties?.Attribute("i") is null && OoxXml.ParseOptionalBool(defaultRunProperties, "i"));
        bool hasHyperlinkClick = HasHyperlinkClick(runProperties);
        string? underlineValue = ReadUnderlineValue(runProperties, defaultRunProperties);
        string? strikeValue = ReadStrikeValue(runProperties, defaultRunProperties);
        string? capsValue = ReadTextCapsValue(runProperties, defaultRunProperties);
        bool underline = underlineValue is null
            ? hasHyperlinkClick
            : !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);

        return new ResolvedRunTextStyle(
            nominalFontSize,
            fontSize,
            ReadCharacterSpacing(runProperties, defaultRunProperties),
            baselineOffset,
            color,
            colorSource,
            alpha,
            TryReadTextOutline(runProperties, defaultRunProperties, theme, colorMap, out TextOutline outline) ? outline : null,
            TryReadHighlightColor(runProperties, out RgbColor highlightColor) ? highlightColor : null,
            hasHyperlinkClick,
            ReadHyperlinkClickId(runProperties),
            bold,
            italic,
            underline,
            underlineValue ?? (hasHyperlinkClick ? "sng" : null),
            IsStrikeEnabled(strikeValue),
            strikeValue,
            capsValue,
            IsKerningEnabled(),
            typeface.Source,
            typeface.Typeface);

        bool HasTextFill()
        {
            return runProperties?.Element(DrawingNamespace + "solidFill") is not null ||
                runProperties?.Element(DrawingNamespace + "noFill") is not null ||
                runProperties?.Element(DrawingNamespace + "gradFill") is not null;
        }

        bool IsKerningEnabled()
        {
            XAttribute? threshold = runProperties?.Attribute("kern") ?? defaultRunProperties?.Attribute("kern");
            if (threshold is null)
            {
                return false;
            }

            double minimumFontSize = int.Parse(threshold.Value, CultureInfo.InvariantCulture) / 100d;
            return minimumFontSize <= 0d || fontSize >= minimumFontSize;
        }
    }

    private static bool HasTextNoFill(XElement? runProperties)
    {
        return runProperties?.Element(DrawingNamespace + "noFill") is not null;
    }

    private static bool TryReadTextOutline(XElement? runProperties, XElement? defaultRunProperties, PptxTheme theme, PptxColorMap colorMap, out TextOutline outline)
    {
        return TryReadTextOutline(runProperties, theme, colorMap, out outline) ||
            TryReadTextOutline(defaultRunProperties, theme, colorMap, out outline);
    }

    private static bool TryReadTextOutline(XElement? runProperties, PptxTheme theme, PptxColorMap colorMap, out TextOutline outline)
    {
        outline = default;
        XElement? line = runProperties?.Element(DrawingNamespace + "ln");
        if (line is null || line.Element(DrawingNamespace + "noFill") is not null)
        {
            return false;
        }

        double? width = line.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : null;
        if (!TryReadSolidColorWithAlpha(line, theme, colorMap, out RgbColor color, out double alpha))
        {
            return false;
        }

        outline = new TextOutline(color, alpha, PptxTextMetricRules.TextOutlineWidth(width));
        return true;
    }

    private static bool HasHyperlinkClick(XElement? runProperties)
    {
        return runProperties?.Element(DrawingNamespace + "hlinkClick") is not null;
    }

    private static string? ReadHyperlinkClickId(XElement? runProperties)
    {
        return (string?)runProperties
            ?.Element(DrawingNamespace + "hlinkClick")
            ?.Attribute(RelationshipsNamespace + "id");
    }

    private static PptxThemeTypefaceResolution ReadRunTypeface(XElement? runProperties, XElement? defaultRunProperties, PptxTheme theme)
    {
        return theme.ResolveTypefaceWithSource(ReadTypeface(runProperties) ?? ReadTypeface(defaultRunProperties));
    }

    private static string? ReadTypeface(XElement? runProperties)
    {
        return (string?)(runProperties?.Element(DrawingNamespace + "latin") ??
            runProperties?.Element(DrawingNamespace + "ea") ??
            runProperties?.Element(DrawingNamespace + "cs"))
            ?.Attribute("typeface");
    }

    private static bool IsTextRunElement(XElement element)
    {
        return element.Name == DrawingNamespace + "r" ||
            element.Name == DrawingNamespace + "fld";
    }

    private static string ReadTextElementText(XElement element, int slideNumber)
    {
        if (slideNumber > 0 &&
            element.Name == DrawingNamespace + "fld" &&
            string.Equals((string?)element.Attribute("type"), "slidenum", StringComparison.OrdinalIgnoreCase))
        {
            return slideNumber.ToString(CultureInfo.InvariantCulture);
        }

        return NormalizeText((string?)element.Element(DrawingNamespace + "t") ?? string.Empty);
    }

    private static string NormalizeText(string text)
    {
        return text;
    }

    private static double ReadFontSize(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("sz") ?? defaultRunProperties?.Attribute("sz")) is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
            : 18d;
    }

    private static double ReadCharacterSpacing(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (runProperties?.Attribute("spc") ?? defaultRunProperties?.Attribute("spc")) is { } spacing
            ? int.Parse(spacing.Value, CultureInfo.InvariantCulture) / 100d
            : 0d;
    }

    private static double ReadBaselineOffset(XElement? runProperties, XElement? defaultRunProperties, double fontSize)
    {
        return (runProperties?.Attribute("baseline") ?? defaultRunProperties?.Attribute("baseline")) is { } baseline
            ? fontSize * int.Parse(baseline.Value, CultureInfo.InvariantCulture) / 100000d
            : 0d;
    }

    private static bool IsStrikeEnabled(XElement? runProperties, XElement? defaultRunProperties)
    {
        return IsStrikeEnabled(ReadStrikeValue(runProperties, defaultRunProperties));
    }

    private static bool IsStrikeEnabled(string? value)
    {
        return value is not null && !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }

    private static string? ReadUnderlineValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("u") ?? defaultRunProperties?.Attribute("u"));
    }

    private static string? ReadStrikeValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("strike") ?? defaultRunProperties?.Attribute("strike"));
    }

    private static string? ReadTextCapsValue(XElement? runProperties, XElement? defaultRunProperties)
    {
        return (string?)(runProperties?.Attribute("cap") ?? defaultRunProperties?.Attribute("cap"));
    }

    private static IReadOnlyList<TextCapsFragment> ApplyTextCaps(string text, XElement? runProperties, XElement? defaultRunProperties)
    {
        string? value = ReadTextCapsValue(runProperties, defaultRunProperties);
        if (text.Length == 0)
        {
            return [];
        }

        if (value is "all")
        {
            return [new TextCapsFragment(text.ToUpperInvariant(), 1d)];
        }

        if (value is not "small")
        {
            return [new TextCapsFragment(text, 1d)];
        }

        var fragments = new List<TextCapsFragment>();
        var builder = new StringBuilder();
        bool? currentSmall = null;
        foreach (char character in text)
        {
            bool isSmall = char.IsLetter(character) && char.IsLower(character);
            if (currentSmall is not null && currentSmall != isSmall)
            {
                fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? PptxTextMetricRules.SmallCapsFontScale() : 1d));
                builder.Clear();
            }

            currentSmall = isSmall;
            builder.Append(char.ToUpperInvariant(character));
        }

        if (builder.Length > 0 && currentSmall is not null)
        {
            fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? PptxTextMetricRules.SmallCapsFontScale() : 1d));
        }

        return fragments;
    }

    private static (TextInsets Insets, TextInsetSources Sources, TextInsetValues Values) ReadTextInsets(
        XElement textBody,
        XElement? inheritedTextBody)
    {
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        XElement? inheritedBodyProperties = inheritedTextBody?.Element(DrawingNamespace + "bodyPr");
        (double left, PptxTextBodyPropertySource leftSource, string? leftValue) = ReadInset(
            bodyProperties, inheritedBodyProperties, "lIns", 91440);
        (double right, PptxTextBodyPropertySource rightSource, string? rightValue) = ReadInset(
            bodyProperties, inheritedBodyProperties, "rIns", 91440);
        (double top, PptxTextBodyPropertySource topSource, string? topValue) = ReadInset(
            bodyProperties, inheritedBodyProperties, "tIns", 45720);
        (double bottom, PptxTextBodyPropertySource bottomSource, string? bottomValue) = ReadInset(
            bodyProperties, inheritedBodyProperties, "bIns", 45720);

        return (
            new TextInsets(left, right, top, bottom),
            new TextInsetSources(leftSource, rightSource, topSource, bottomSource),
            new TextInsetValues(leftValue, rightValue, topValue, bottomValue));
    }
}
