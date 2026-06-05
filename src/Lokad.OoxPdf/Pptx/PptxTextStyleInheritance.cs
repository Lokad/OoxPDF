using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxTextStyleInheritance
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

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
