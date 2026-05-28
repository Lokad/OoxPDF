using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed class PptxTheme
{
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string ThemeRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    private const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

    private PptxTheme(
        IReadOnlyDictionary<string, RgbColor> colors,
        string? majorLatinFont,
        string? majorEastAsianFont,
        string? majorComplexScriptFont,
        string? minorLatinFont,
        string? minorEastAsianFont,
        string? minorComplexScriptFont,
        IReadOnlyList<XElement> fillStyles,
        IReadOnlyList<XElement> lineStyles,
        IReadOnlyList<XElement> effectStyles)
    {
        Colors = colors;
        MajorLatinFont = majorLatinFont;
        MajorEastAsianFont = majorEastAsianFont;
        MajorComplexScriptFont = majorComplexScriptFont;
        MinorLatinFont = minorLatinFont;
        MinorEastAsianFont = minorEastAsianFont;
        MinorComplexScriptFont = minorComplexScriptFont;
        FillStyles = fillStyles;
        LineStyles = lineStyles;
        EffectStyles = effectStyles;
    }

    public IReadOnlyDictionary<string, RgbColor> Colors { get; }

    public string? MajorLatinFont { get; }

    public string? MajorEastAsianFont { get; }

    public string? MajorComplexScriptFont { get; }

    public string? MinorLatinFont { get; }

    public string? MinorEastAsianFont { get; }

    public string? MinorComplexScriptFont { get; }

    public IReadOnlyList<XElement> FillStyles { get; }

    public IReadOnlyList<XElement> LineStyles { get; }

    public IReadOnlyList<XElement> EffectStyles { get; }

    public static PptxTheme Empty { get; } = new(new Dictionary<string, RgbColor>(), null, null, null, null, null, null, [], [], []);

    public static PptxTheme Load(OoxPackage package, string presentationPartName)
    {
        OoxRelationship? themeRelationship = package.GetRelationships(presentationPartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == ThemeRelationshipType && r.ResolvedTarget is not null);
        string? themePartName = themeRelationship?.ResolvedTarget;
        if (themePartName is null)
        {
            OoxRelationship? masterRelationship = package.GetRelationships(presentationPartName)
                .FirstOrDefault(r => !r.IsExternal && r.Type == SlideMasterRelationshipType && r.ResolvedTarget is not null);
            themePartName = masterRelationship?.ResolvedTarget is null
                ? null
                : package.GetRelationships(masterRelationship.ResolvedTarget)
                    .FirstOrDefault(r => !r.IsExternal && r.Type == ThemeRelationshipType && r.ResolvedTarget is not null)
                    ?.ResolvedTarget;
        }

        if (themePartName is null)
        {
            return Empty;
        }

        OoxPart? themePart = package.GetPart(themePartName);
        if (themePart is null)
        {
            return Empty;
        }

        using Stream stream = themePart.OpenRead();
        XDocument document = SafeXml.Load(stream);
        var colors = new Dictionary<string, RgbColor>(StringComparer.Ordinal);
        XElement? colorScheme = document.Descendants(DrawingNamespace + "clrScheme").FirstOrDefault();
        if (colorScheme is not null)
        {
            foreach (XElement colorElement in colorScheme.Elements())
            {
                if (TryReadColorElement(colorElement, out RgbColor color))
                {
                    colors[colorElement.Name.LocalName] = color;
                }
            }
        }

        XElement? fontScheme = document.Descendants(DrawingNamespace + "fontScheme").FirstOrDefault();
        string? majorLatin = (string?)fontScheme?
            .Element(DrawingNamespace + "majorFont")?
            .Element(DrawingNamespace + "latin")?
            .Attribute("typeface");
        string? majorEastAsian = (string?)fontScheme?
            .Element(DrawingNamespace + "majorFont")?
            .Element(DrawingNamespace + "ea")?
            .Attribute("typeface");
        string? majorComplexScript = (string?)fontScheme?
            .Element(DrawingNamespace + "majorFont")?
            .Element(DrawingNamespace + "cs")?
            .Attribute("typeface");
        string? minorLatin = (string?)fontScheme?
            .Element(DrawingNamespace + "minorFont")?
            .Element(DrawingNamespace + "latin")?
            .Attribute("typeface");
        string? minorEastAsian = (string?)fontScheme?
            .Element(DrawingNamespace + "minorFont")?
            .Element(DrawingNamespace + "ea")?
            .Attribute("typeface");
        string? minorComplexScript = (string?)fontScheme?
            .Element(DrawingNamespace + "minorFont")?
            .Element(DrawingNamespace + "cs")?
            .Attribute("typeface");

        XElement? formatScheme = document.Descendants(DrawingNamespace + "fmtScheme").FirstOrDefault();
        IReadOnlyList<XElement> fillStyles = formatScheme?
            .Element(DrawingNamespace + "fillStyleLst")?
            .Elements()
            .Select(element => new XElement(element))
            .ToArray() ?? [];
        IReadOnlyList<XElement> lineStyles = formatScheme?
            .Element(DrawingNamespace + "lnStyleLst")?
            .Elements(DrawingNamespace + "ln")
            .Select(element => new XElement(element))
            .ToArray() ?? [];
        IReadOnlyList<XElement> effectStyles = formatScheme?
            .Element(DrawingNamespace + "effectStyleLst")?
            .Elements(DrawingNamespace + "effectStyle")
            .Select(element => new XElement(element))
            .ToArray() ?? [];

        return new PptxTheme(colors, majorLatin, majorEastAsian, majorComplexScript, minorLatin, minorEastAsian, minorComplexScript, fillStyles, lineStyles, effectStyles);
    }

    public bool TryResolveColor(string schemeColor, out RgbColor color)
    {
        return TryResolveColor(schemeColor, PptxColorMap.Default, out color);
    }

    public bool TryResolveColor(string schemeColor, PptxColorMap colorMap, out RgbColor color)
    {
        string mappedColor = colorMap.ResolveSchemeColor(schemeColor);
        return Colors.TryGetValue(mappedColor, out color) ||
            (!string.Equals(mappedColor, schemeColor, StringComparison.Ordinal) && Colors.TryGetValue(schemeColor, out color));
    }

    public string? ResolveTypeface(string? typeface)
    {
        return ResolveTypefaceWithSource(typeface).Typeface;
    }

    public PptxThemeTypefaceResolution ResolveTypefaceWithSource(string? typeface)
    {
        return typeface switch
        {
            "+mj-lt" => new(typeface, MajorLatinFont, PptxThemeTypefaceSource.MajorLatin),
            "+mj-ea" when MajorEastAsianFont is not null => new(typeface, MajorEastAsianFont, PptxThemeTypefaceSource.MajorEastAsian),
            "+mj-ea" => new(typeface, MajorLatinFont, PptxThemeTypefaceSource.MajorEastAsianFallbackLatin),
            "+mj-cs" when MajorComplexScriptFont is not null => new(typeface, MajorComplexScriptFont, PptxThemeTypefaceSource.MajorComplexScript),
            "+mj-cs" => new(typeface, MajorLatinFont, PptxThemeTypefaceSource.MajorComplexScriptFallbackLatin),
            "+mn-lt" => new(typeface, MinorLatinFont, PptxThemeTypefaceSource.MinorLatin),
            "+mn-ea" when MinorEastAsianFont is not null => new(typeface, MinorEastAsianFont, PptxThemeTypefaceSource.MinorEastAsian),
            "+mn-ea" => new(typeface, MinorLatinFont, PptxThemeTypefaceSource.MinorEastAsianFallbackLatin),
            "+mn-cs" when MinorComplexScriptFont is not null => new(typeface, MinorComplexScriptFont, PptxThemeTypefaceSource.MinorComplexScript),
            "+mn-cs" => new(typeface, MinorLatinFont, PptxThemeTypefaceSource.MinorComplexScriptFallbackLatin),
            null or "" => new(typeface, MinorLatinFont, PptxThemeTypefaceSource.DefaultMinorLatin),
            _ => new(typeface, typeface, PptxThemeTypefaceSource.Direct)
        };
    }

    public bool TryGetFillStyle(int index, out XElement fillStyle)
    {
        if (index > 0 && index <= FillStyles.Count)
        {
            fillStyle = FillStyles[index - 1];
            return true;
        }

        fillStyle = null!;
        return false;
    }

    public bool TryGetLineStyle(int index, out XElement lineStyle)
    {
        if (index > 0 && index <= LineStyles.Count)
        {
            lineStyle = LineStyles[index - 1];
            return true;
        }

        lineStyle = null!;
        return false;
    }

    public bool TryGetEffectStyle(int index, out XElement effectStyle)
    {
        if (index > 0 && index <= EffectStyles.Count)
        {
            effectStyle = EffectStyles[index - 1];
            return true;
        }

        effectStyle = null!;
        return false;
    }

    private static bool TryReadColorElement(XElement colorElement, out RgbColor color)
    {
        XElement? srgb = colorElement.Element(DrawingNamespace + "srgbClr");
        string? hex = (string?)srgb?.Attribute("val");
        if (hex is null)
        {
            hex = (string?)colorElement.Element(DrawingNamespace + "sysClr")?.Attribute("lastClr");
        }

        return RgbColor.TryParse(hex, out color);
    }

}

internal sealed class PptxColorMap
{
    private static readonly string[] SchemeSlots =
    [
        "bg1",
        "tx1",
        "bg2",
        "tx2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink"
    ];

    private PptxColorMap(IReadOnlyDictionary<string, string> mappings)
    {
        Mappings = mappings;
    }

    public IReadOnlyDictionary<string, string> Mappings { get; }

    public static PptxColorMap Default { get; } = new(new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["bg1"] = "lt1",
        ["tx1"] = "dk1",
        ["bg2"] = "lt2",
        ["tx2"] = "dk2",
        ["accent1"] = "accent1",
        ["accent2"] = "accent2",
        ["accent3"] = "accent3",
        ["accent4"] = "accent4",
        ["accent5"] = "accent5",
        ["accent6"] = "accent6",
        ["hlink"] = "hlink",
        ["folHlink"] = "folHlink"
    });

    public string ResolveSchemeColor(string schemeColor)
    {
        return Mappings.TryGetValue(schemeColor, out string? mappedColor) && !string.IsNullOrWhiteSpace(mappedColor)
            ? mappedColor
            : schemeColor;
    }

    public static PptxColorMap FromElement(XElement? colorMapElement)
    {
        return FromElement(colorMapElement, Default);
    }

    public static PptxColorMap FromElement(XElement? colorMapElement, PptxColorMap inheritedMap)
    {
        if (colorMapElement is null)
        {
            return inheritedMap;
        }

        var mappings = new Dictionary<string, string>(inheritedMap.Mappings, StringComparer.Ordinal);
        foreach (string slot in SchemeSlots)
        {
            string? mappedColor = (string?)colorMapElement.Attribute(slot);
            if (!string.IsNullOrWhiteSpace(mappedColor))
            {
                mappings[slot] = mappedColor;
            }
        }

        return mappings.Count == 0
            ? inheritedMap
            : new PptxColorMap(mappings);
    }
}

internal readonly record struct PptxThemeTypefaceResolution(
    string? RequestedTypeface,
    string? Typeface,
    PptxThemeTypefaceSource Source);

internal enum PptxThemeTypefaceSource
{
    Direct,
    DefaultMinorLatin,
    MajorLatin,
    MajorEastAsian,
    MajorEastAsianFallbackLatin,
    MajorComplexScript,
    MajorComplexScriptFallbackLatin,
    MinorLatin,
    MinorEastAsian,
    MinorEastAsianFallbackLatin,
    MinorComplexScript,
    MinorComplexScriptFallbackLatin
}
