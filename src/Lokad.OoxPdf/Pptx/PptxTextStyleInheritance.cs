using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxTextStyleInheritance
{

    public static XElement? FindInheritedTextStyle(XElement shape, IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        string styleName = ReadPlaceholderTextStyleName(shape);
        foreach (XDocument source in placeholderSources)
        {
            XElement? style = source.Root?
                .Element(PresentationNamespace + "txStyles")
                ?.Element(PresentationNamespace + styleName);
            XElement? level = style?.Element(DrawingNamespace + levelName) ??
                style?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    public static string PlaceholderListStyleLayerKindName(XElement placeholder, int sourceIndex, int sourceCount)
    {
        string? sourceRootName = placeholder.Document?.Root?.Name.LocalName;
        if (sourceRootName == "sldMaster")
        {
            return "MasterPlaceholderListStyle";
        }

        if (sourceRootName == "sldLayout")
        {
            return "LayoutPlaceholderListStyle";
        }

        return (sourceCount, sourceIndex) switch
        {
            (2, 0) => "MasterPlaceholderListStyle",
            (2, _) => "LayoutPlaceholderListStyle",
            _ => "InheritedPlaceholderListStyle",
        };
    }

    public static string PlaceholderListStyleLayerName(XElement placeholder, int sourceIndex, int sourceCount)
    {
        string? sourceRootName = placeholder.Document?.Root?.Name.LocalName;
        if (sourceRootName == "sldMaster")
        {
            return "master.placeholder.lstStyle";
        }

        if (sourceRootName == "sldLayout")
        {
            return "layout.placeholder.lstStyle";
        }

        return (sourceCount, sourceIndex) switch
        {
            (2, 0) => "master.placeholder.lstStyle",
            (2, _) => "layout.placeholder.lstStyle",
            _ => $"inherited.placeholder[{sourceIndex}].lstStyle",
        };
    }
    public static XElement? FindDefaultTextStyle(IReadOnlyList<XDocument> placeholderSources, string levelName)
    {
        foreach (XDocument source in placeholderSources)
        {
            XElement? defaultTextStyle = source.Root?.Element(PresentationNamespace + "defaultTextStyle");
            XElement? level = defaultTextStyle?.Element(DrawingNamespace + levelName) ??
                defaultTextStyle?.Element(DrawingNamespace + "defPPr");
            if (level is not null)
            {
                return level;
            }
        }

        return null;
    }

    private static string ReadPlaceholderTextStyleName(XElement shape)
    {
        string? placeholderType = (string?)shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph")
            ?.Attribute("type");
        return placeholderType switch
        {
            "title" or "ctrTitle" => "titleStyle",
            "body" or "subTitle" => "bodyStyle",
            _ => "otherStyle"
        };
    }
}
