using System.Globalization;
using System.Text;
using System.Xml.Linq;

using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;
using static Lokad.OoxPdf.Pptx.PptxRunTextAttributeReaders;

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
        // The caller supplies the complete cascade. A partial authored defRPr
        // must not replace inherited size/typeface with renderer defaults.
        XElement? defaultRunProperties = defaultParagraphProperties?.Element(DrawingNamespace + "defRPr");
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
        (RgbColor Color, double Alpha)? shapeFontPaint,
        PptxTheme theme,
        PptxColorMap colorMap,
        double fontScale,
        PptxSceneTableCellTextStyle tableStyleTextStyle,
        bool disableZeroKerning = false)
    {
        XElement? paragraphDefaults = cascade.Layers
            .FirstOrDefault(layer => layer.Kind == PptxRunStyleLayerKind.ParagraphPropertiesDefaultRunProperties)?.Source;
        tableStyleTextStyle = BelowParagraphDefaults(tableStyleTextStyle, paragraphDefaults);
        XElement? runProperties = cascade.DirectProperties;
        XElement? defaultRunProperties = cascade.ResolvedDefaultProperties;
        PptxSceneHyperlinkClick hyperlink = PptxSceneBuilder.ReadHyperlinkClick(runProperties);
        double nominalFontSize = ReadFontSize(runProperties, defaultRunProperties) * fontScale;
        double baselineOffset = ReadBaselineOffset(runProperties, defaultRunProperties, nominalFontSize);
        double fontSize = PptxTextMetricRules.ShouldScaleSuperscriptSubscript(baselineOffset, nominalFontSize)
            ? PptxTextMetricRules.SuperscriptSubscriptFontSize(nominalFontSize)
            : nominalFontSize;
        double alpha = 1d;
        RgbColor color;
        PptxRunTextColorSource colorSource;
        if (PptxRunTextAttributeReaders.HasTextNoFill(runProperties))
        {
            color = new RgbColor(0, 0, 0);
            colorSource = PptxRunTextColorSource.RunNoFill;
            alpha = 0d;
        }
        else if (hyperlink.IsDefined && theme.TryResolveColor("hlink", colorMap, out RgbColor hyperlinkColor))
        {
            color = hyperlinkColor;
            colorSource = PptxRunTextColorSource.ThemeHyperlink;
        }
        else if (PptxColorResolver.TryReadSolidColorWithAlpha(runProperties, theme, colorMap, out RgbColor runColor, out double runAlpha))
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
        // Authored paragraph paint outranks the shape font reference. Merged
        // presentation/theme defaults remain below it; keep the source owner.
        else if (ReadDefaultTextPaint(paragraphDefaults, theme, colorMap) is { } paragraphPaint)
        {
            (color, alpha, colorSource) = paragraphPaint;
        }
        else if (shapeFontPaint is { } fontRefPaint)
        {
            (color, alpha) = fontRefPaint;
            colorSource = PptxRunTextColorSource.ShapeFontRef;
        }
        else if (ReadDefaultTextPaint(defaultRunProperties, theme, colorMap) is { } defaultPaint)
        {
            (color, alpha, colorSource) = defaultPaint;
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
        string? underlineValue = ReadUnderlineValue(runProperties, defaultRunProperties);
        string? strikeValue = ReadStrikeValue(runProperties, defaultRunProperties);
        string? capsValue = ReadTextCapsValue(runProperties, defaultRunProperties);
        bool underline = underlineValue is null
            ? hyperlink.IsDefined
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
            ReadHighlightPaint(runProperties, defaultRunProperties, theme, colorMap),
            hyperlink.IsDefined,
            hyperlink.RelationshipId,
            hyperlink.Action,
            bold,
            italic,
            underline,
            underlineValue ?? (hyperlink.IsDefined ? "sng" : null),
            IsStrikeEnabled(strikeValue),
            strikeValue,
            capsValue,
            IsKerningEnabled(),
            typeface.Source,
            typeface.Typeface);

        bool HasTextFill() => HasRunTextFill(runProperties);

        bool IsKerningEnabled()
        {
            XAttribute? threshold = runProperties?.Attribute("kern") ?? defaultRunProperties?.Attribute("kern");
            // Office uses its 12pt threshold when the source cascade omits
            // kerning. Explicit zero keeps the qualified no-wrap nominal rule.
            double minimumFontSize = threshold is null
                ? PptxTextMetricRules.OfficeDefaultKerningMinimumFontSizePoints
                : int.Parse(threshold.Value, CultureInfo.InvariantCulture) / 100d;
            if (disableZeroKerning && minimumFontSize == 0d)
            {
                return false;
            }

            return minimumFontSize <= 0d || fontSize >= minimumFontSize;
        }
    }

    private static (RgbColor Color, double Alpha, PptxRunTextColorSource Source)? ReadDefaultTextPaint(
        XElement? properties, PptxTheme theme, PptxColorMap colorMap)
    {
        if (HasTextNoFill(properties)) return (default, 0d, PptxRunTextColorSource.DefaultNoFill);
        return PptxColorResolver.TryReadSolidColorWithAlpha(properties, theme, colorMap, out RgbColor color, out double alpha)
            ? (color, alpha, PptxRunTextColorSource.DefaultSolidFill) : null;
    }

    private static bool TryReadTextOutline(XElement? runProperties, XElement? defaultRunProperties, PptxTheme theme, PptxColorMap colorMap, out TextOutline outline)
    {
        // Explicit removal terminates inheritance. An unreadable authored paint
        // still retains the established default-outline fallback below.
        outline = default;
        if (runProperties?.Element(DrawingNamespace + "ln")?.Element(DrawingNamespace + "noFill") is not null)
            return false;
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
        if (!PptxColorResolver.TryReadSolidColorWithAlpha(line, theme, colorMap, out RgbColor color, out double alpha))
        {
            return false;
        }

        outline = new TextOutline(color, alpha, PptxTextMetricRules.TextOutlineWidth(width));
        return true;
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

    private static string ReadTextElementText(XElement element, long slideNumber)
    {
        if (slideNumber >= 0 &&
            element.Name == DrawingNamespace + "fld" &&
            string.Equals((string?)element.Attribute("type"), "slidenum", StringComparison.OrdinalIgnoreCase))
        {
            return slideNumber.ToString(CultureInfo.InvariantCulture);
        }

        return (string?)element.Element(DrawingNamespace + "t") ?? string.Empty;
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
        Span<char> uppercase = stackalloc char[2];
        bool? currentSmall = null;
        foreach (Rune character in text.EnumerateRunes())
        {
            bool isSmall = Rune.IsLetter(character) && Rune.IsLower(character);
            if (currentSmall is not null && currentSmall != isSmall)
            {
                fragments.Add(new TextCapsFragment(builder.ToString(), currentSmall.Value ? PptxTextMetricRules.SmallCapsFontScale() : 1d));
                builder.Clear();
            }

            currentSmall = isSmall;
            int length = Rune.ToUpperInvariant(character).EncodeToUtf16(uppercase);
            builder.Append(uppercase[..length]);
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
