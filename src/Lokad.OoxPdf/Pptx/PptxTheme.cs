using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed class PptxTheme
{
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string ThemeRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

    private PptxTheme(IReadOnlyDictionary<string, RgbColor> colors, string? majorLatinFont, string? minorLatinFont)
    {
        Colors = colors;
        MajorLatinFont = majorLatinFont;
        MinorLatinFont = minorLatinFont;
    }

    public IReadOnlyDictionary<string, RgbColor> Colors { get; }

    public string? MajorLatinFont { get; }

    public string? MinorLatinFont { get; }

    public static PptxTheme Empty { get; } = new(new Dictionary<string, RgbColor>(), null, null);

    public static PptxTheme Load(OoxPackage package, string presentationPartName)
    {
        OoxRelationship? themeRelationship = package.GetRelationships(presentationPartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == ThemeRelationshipType && r.ResolvedTarget is not null);
        if (themeRelationship?.ResolvedTarget is null)
        {
            return Empty;
        }

        OoxPart? themePart = package.GetPart(themeRelationship.ResolvedTarget);
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

        return new PptxTheme(colors, majorLatin, minorLatin);
    }

    public bool TryResolveColor(string schemeColor, out RgbColor color)
    {
        return Colors.TryGetValue(schemeColor, out color);
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
