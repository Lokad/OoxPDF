using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    internal static PptxSceneChartTextStyleOverride ReadChartTextStyleOverride(XElement? parent, PptxTheme theme)
    {
        return ReadChartTextStyleOverride(parent, theme, PptxColorMap.Default);
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartTitleTextStyleOverride(PptxSceneChart chart)
    {
        return ResolveChartElementTextStyleOverride(chart, chart.Title.TextStyle, "title");
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartLegendTextStyleOverride(PptxSceneChart chart)
    {
        return ResolveChartElementTextStyleOverride(chart, chart.Legend.TextStyle, "legend");
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartAxisTextStyleOverride(PptxSceneChart chart, PptxSceneChartAxis? axis, string? chartStyleRole)
    {
        return ResolveChartElementTextStyleOverride(chart, axis?.TextStyle ?? default, chartStyleRole);
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartElementTextStyleOverride(PptxSceneChart chart, PptxSceneChartTextStyleOverride elementTextStyle, string? chartStyleRole)
    {
        PptxSceneChartTextStyleOverride style = chart.TextStyle;
        if (!string.IsNullOrWhiteSpace(chartStyleRole))
        {
            style = MergeChartTextStyleOverride(style, ReadChartStyleRoleTextStyleOverride(chart.StylePart, chartStyleRole));
        }

        return MergeChartTextStyleOverride(style, elementTextStyle);
    }

    internal static PptxSceneChartTextStyleOverride ResolveChartDataLabelTextStyleOverride(PptxSceneChart chart, PptxSceneChartDataLabels labels)
    {
        return ResolveChartElementTextStyleOverride(chart, labels.TextStyle, "dataLabel");
    }

    internal static PptxSceneChartTextStyleOverride ReadChartStyleRoleTextStyleOverride(PptxSceneChartStyle stylePart, string role)
    {
        PptxSceneChartStyleEntry entry = stylePart.Entries.FirstOrDefault(item => item.Role == role);
        return entry.TextStyle;
    }

    internal static PptxSceneChartTextStyleOverride MergeChartTextStyleOverride(PptxSceneChartTextStyleOverride style, PptxSceneChartTextStyleOverride next)
    {
        return new PptxSceneChartTextStyleOverride(
            next.FontFamily ?? style.FontFamily,
            next.FontFamily is null ? style.RequestedTypeface : next.RequestedTypeface,
            next.FontFamily is null ? style.TypefaceSource : next.TypefaceSource,
            next.FontSize ?? style.FontSize,
            next.CharacterSpacing ?? style.CharacterSpacing,
            next.Color ?? style.Color,
            next.Alpha ?? style.Alpha,
            next.Bold ?? style.Bold,
            next.Italic ?? style.Italic,
            next.Underline ?? style.Underline,
            next.Strike ?? style.Strike);
    }

    internal static PptxSceneChartTextStyleOverride ReadChartTextStyleOverride(XElement? parent, PptxTheme theme, PptxColorMap colorMap)
    {
        XElement? defaultRunProperties = parent?
            .Element(ChartNamespace + "txPr")?
            .Elements(DrawingNamespace + "p")
            .Select(paragraph => paragraph.Element(DrawingNamespace + "pPr")?.Element(DrawingNamespace + "defRPr"))
            .FirstOrDefault(element => element is not null);
        if (defaultRunProperties is null)
        {
            return default;
        }

        string? typeface = (string?)defaultRunProperties.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)defaultRunProperties.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)defaultRunProperties.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        PptxThemeTypefaceResolution typefaceResolution = string.IsNullOrWhiteSpace(typeface)
            ? default
            : theme.ResolveTypefaceWithSource(typeface);
        string? fontFamily = typefaceResolution.Typeface;
        double? fontSize = defaultRunProperties.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0
                ? sizeHundredths / 100d
                : null;
        double? characterSpacing = ReadOptionalChartCharacterSpacing(defaultRunProperties);
        RgbColor? color = PptxColorResolver.TryReadSolidColorWithAlpha(defaultRunProperties.Element(DrawingNamespace + "solidFill"), theme, colorMap, out RgbColor parsedColor, out double alpha)
            ? parsedColor
            : null;
        bool? bold = ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "b");
        bool? italic = ReadOptionalOoxmlBooleanAttribute(defaultRunProperties, "i");
        bool? underline = ReadChartUnderline(defaultRunProperties);
        bool? strike = ReadChartStrike(defaultRunProperties);
        return new PptxSceneChartTextStyleOverride(
            fontFamily,
            typefaceResolution.RequestedTypeface,
            typefaceResolution.RequestedTypeface is null ? null : typefaceResolution.Source,
            fontSize,
            characterSpacing,
            color,
            color is null ? null : alpha,
            bold,
            italic,
            underline,
            strike);
    }

    internal static PptxSceneChartTextBodyProperties ReadChartTextBodyProperties(XElement? parent)
    {
        XElement? bodyProperties = parent?
            .Element(ChartNamespace + "txPr")?
            .Element(DrawingNamespace + "bodyPr");
        string rotation = (string?)bodyProperties?.Attribute("rot") ?? string.Empty;
        return new PptxSceneChartTextBodyProperties(
            ParseOptionalOoxmlAngle(rotation),
            rotation,
            (string?)bodyProperties?.Attribute("vert") ?? string.Empty,
            (string?)bodyProperties?.Attribute("vertOverflow") ?? string.Empty);
    }

    private static double? ParseOptionalOoxmlAngle(string value)
    {
        return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long rawAngle)
            ? rawAngle / 60000d
            : null;
    }

    private static PptxSceneChartTextStyleOverride ReadChartTextRunStyle(XElement? runProperties, PptxTheme theme)
    {
        return ReadChartTextRunStyle(runProperties, theme, PptxColorMap.Default);
    }

    private static PptxSceneChartTextStyleOverride ReadChartTextRunStyle(XElement? runProperties, PptxTheme theme, PptxColorMap colorMap)
    {
        if (runProperties is null)
        {
            return default;
        }

        string? typeface = (string?)runProperties.Element(DrawingNamespace + "latin")?.Attribute("typeface") ??
            (string?)runProperties.Element(DrawingNamespace + "ea")?.Attribute("typeface") ??
            (string?)runProperties.Element(DrawingNamespace + "cs")?.Attribute("typeface");
        PptxThemeTypefaceResolution typefaceResolution = string.IsNullOrWhiteSpace(typeface)
            ? default
            : theme.ResolveTypefaceWithSource(typeface);
        string? fontFamily = typefaceResolution.Typeface;
        double? fontSize = runProperties.Attribute("sz") is { } sizeAttribute &&
            int.TryParse(sizeAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sizeHundredths) &&
            sizeHundredths > 0
                ? sizeHundredths / 100d
                : null;
        double? characterSpacing = ReadOptionalChartCharacterSpacing(runProperties);
        RgbColor? color = PptxColorResolver.TryReadSolidColorWithAlpha(runProperties.Element(DrawingNamespace + "solidFill"), theme, colorMap, out RgbColor parsedColor, out double alpha)
            ? parsedColor
            : null;
        bool? bold = ReadOptionalOoxmlBooleanAttribute(runProperties, "b");
        bool? italic = ReadOptionalOoxmlBooleanAttribute(runProperties, "i");
        bool? underline = ReadChartUnderline(runProperties);
        bool? strike = ReadChartStrike(runProperties);
        return new PptxSceneChartTextStyleOverride(
            fontFamily,
            typefaceResolution.RequestedTypeface,
            typefaceResolution.RequestedTypeface is null ? null : typefaceResolution.Source,
            fontSize,
            characterSpacing,
            color,
            color is null ? null : alpha,
            bold,
            italic,
            underline,
            strike);
    }

    private static bool? ReadChartUnderline(XElement runProperties)
    {
        string? value = (string?)runProperties.Attribute("u");
        return value is null ? null : !value.Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static bool? ReadChartStrike(XElement runProperties)
    {
        string? value = (string?)runProperties.Attribute("strike");
        return value is null ? null : !value.Equals("noStrike", StringComparison.OrdinalIgnoreCase);
    }
}
