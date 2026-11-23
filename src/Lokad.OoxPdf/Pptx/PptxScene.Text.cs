using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxSceneBuilder
{
    private static PptxSceneTextBody? ReadTextBody(XElement element, IReadOnlyList<XDocument> placeholderSources, XElement? defaultTextStyle)
    {
        XElement? textBody = element.Element(PresentationNamespace + "txBody");
        if (textBody is null)
        {
            return null;
        }

        IReadOnlyList<XElement> inheritedPlaceholderShapes = PptxPlaceholderMatcher.FindInheritedPlaceholderShapes(element, placeholderSources);
        XElement? bodyProperties = textBody.Element(DrawingNamespace + "bodyPr");
        string? orientation = (string?)bodyProperties?.Attribute("vert");
        for (int index = inheritedPlaceholderShapes.Count - 1; orientation is null && index >= 0; index--)
        {
            orientation = (string?)inheritedPlaceholderShapes[index].Element(PresentationNamespace + "txBody")
                ?.Element(DrawingNamespace + "bodyPr")?.Attribute("vert");
        }
        return new PptxSceneTextBody(
            bodyProperties,
            PptxRenderer.HasUnsupportedTextOrientation(orientation),
            HasUnsupportedTextVerticalOverflow(bodyProperties),
            textBody.Elements(DrawingNamespace + "p").Select(paragraph => ReadParagraph(paragraph, element, textBody, inheritedPlaceholderShapes, placeholderSources, defaultTextStyle)).ToArray());
    }

    private static bool HasUnsupportedTextVerticalOverflow(XElement? bodyProperties)
    {
        string? overflow = (string?)bodyProperties?.Attribute("vertOverflow");
        return overflow?.Equals("ellipsis", StringComparison.OrdinalIgnoreCase) == true;
    }

    private static PptxSceneTextParagraph ReadParagraph(
        XElement paragraph,
        XElement shape,
        XElement textBody,
        IReadOnlyList<XElement> inheritedPlaceholderShapes,
        IReadOnlyList<XDocument> placeholderSources,
        XElement? defaultTextStyle)
    {
        XElement? properties = paragraph.Element(DrawingNamespace + "pPr");
        int level = PptxTextStyleInheritance.ReadParagraphLevel(properties);
        string levelName = PptxTextStyleInheritance.ParagraphLevelName(level);
        var cascadeLayers = new List<PptxSceneCascadeLayer>
        {
            new("shape.lstStyle", "ShapeListStyle", textBody.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)),
        };
        cascadeLayers.AddRange(inheritedPlaceholderShapes
            .Select((placeholder, index) => new PptxSceneCascadeLayer(
                PptxTextStyleInheritance.PlaceholderListStyleLayerName(placeholder, index, placeholderSources.Count),
                PptxTextStyleInheritance.PlaceholderListStyleLayerKindName(placeholder, index, placeholderSources.Count),
                placeholder.Element(PresentationNamespace + "txBody")?.Element(DrawingNamespace + "lstStyle")?.Element(DrawingNamespace + levelName)))
            .Reverse());
        cascadeLayers.Add(new PptxSceneCascadeLayer(
            "inherited.txStyle",
            "InheritedTextStyle",
            PptxTextStyleInheritance.FindInheritedTextStyle(shape, placeholderSources, levelName))
);
        cascadeLayers.Add(new PptxSceneCascadeLayer(
            "defaultTextStyle",
            "DefaultTextStyle",
            PptxTextStyleInheritance.FindDefaultTextStyle(defaultTextStyle, levelName))
);
        return new PptxSceneTextParagraph(
            properties,
            paragraph.Element(DrawingNamespace + "endParaRPr"),
            level,
            paragraph.Elements().Select(run => ReadRun(run)).Where(run => run is not null).Cast<PptxSceneTextRun>().ToArray(),
            cascadeLayers,
            paragraph);
    }

    private static PptxSceneTextRun? ReadRun(XElement element)
    {
        if (element.Name == DrawingNamespace + "r" || element.Name == DrawingNamespace + "fld")
        {
            return new PptxSceneTextRun((string?)element.Element(DrawingNamespace + "t") ?? string.Empty);
        }

        if (element.Name == DrawingNamespace + "br")
        {
            return new PptxSceneTextRun("\n");
        }

        return null;
    }
}
