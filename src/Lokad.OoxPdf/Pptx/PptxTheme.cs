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
        string? minorLatinFont,
        IReadOnlyList<XElement> fillStyles,
        IReadOnlyList<XElement> lineStyles)
    {
        Colors = colors;
        MajorLatinFont = majorLatinFont;
        MinorLatinFont = minorLatinFont;
        FillStyles = fillStyles;
        LineStyles = lineStyles;
    }

    public IReadOnlyDictionary<string, RgbColor> Colors { get; }

    public string? MajorLatinFont { get; }

    public string? MinorLatinFont { get; }

    public IReadOnlyList<XElement> FillStyles { get; }

    public IReadOnlyList<XElement> LineStyles { get; }

    public static PptxTheme Empty { get; } = new(new Dictionary<string, RgbColor>(), null, null, [], []);

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
        string? minorLatin = (string?)fontScheme?
            .Element(DrawingNamespace + "minorFont")?
            .Element(DrawingNamespace + "latin")?
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

        return new PptxTheme(colors, majorLatin, minorLatin, fillStyles, lineStyles);
    }

    public bool TryResolveColor(string schemeColor, out RgbColor color)
    {
        return Colors.TryGetValue(schemeColor, out color) ||
            Colors.TryGetValue(ResolveSchemeAlias(schemeColor), out color);
    }

    public string? ResolveTypeface(string? typeface)
    {
        return typeface switch
        {
            "+mj-lt" => MajorLatinFont,
            "+mn-lt" => MinorLatinFont,
            null or "" => MinorLatinFont,
            _ => typeface
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

    private static string ResolveSchemeAlias(string schemeColor)
    {
        return schemeColor switch
        {
            "bg1" => "lt1",
            "tx1" => "dk1",
            "bg2" => "lt2",
            "tx2" => "dk2",
            _ => schemeColor
        };
    }
}
