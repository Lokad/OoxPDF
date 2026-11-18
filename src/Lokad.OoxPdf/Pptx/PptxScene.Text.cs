using System.Globalization;
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
        return new PptxSceneTextBody(
            bodyProperties,
            textBody.Element(DrawingNamespace + "lstStyle"),
            HasUnsupportedTextOrientation(bodyProperties),
            HasUnsupportedTextVerticalOverflow(bodyProperties),
            textBody.Elements(DrawingNamespace + "p").Select(paragraph => ReadParagraph(paragraph, element, textBody, inheritedPlaceholderShapes, placeholderSources, defaultTextStyle)).ToArray());
    }

    private static bool HasUnsupportedTextOrientation(XElement? bodyProperties)
    {
        string? orientation = (string?)bodyProperties?.Attribute("vert");
        return !string.IsNullOrEmpty(orientation) &&
            !orientation.Equals("horz", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("vert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("vert270", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("eaVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("mongolianVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("wordArtVert", StringComparison.OrdinalIgnoreCase) &&
            !orientation.Equals("wordArtVertRtl", StringComparison.OrdinalIgnoreCase);
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
        int level = properties?.Attribute("lvl") is { } levelAttribute
            ? int.Parse(levelAttribute.Value, CultureInfo.InvariantCulture)
            : 0;
        string levelName = "lvl" + Math.Clamp(level + 1, 1, 9).ToString(CultureInfo.InvariantCulture) + "pPr";
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

    // Cell paragraphs retain the same source/cascade contract as shape paragraphs;
    // the renderer resolves table-style defaults along with the live text style.
    internal static IReadOnlyList<PptxSceneTextParagraph> ReadTableCellParagraphs(
        XElement textBody,
        IReadOnlyList<XDocument> placeholderSources,
        XElement? defaultTextStyle)
    {
        return textBody.Elements(DrawingNamespace + "p")
            .Select(paragraph => ReadParagraph(paragraph, textBody, textBody, [], placeholderSources, defaultTextStyle))
            .ToArray();
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

    private static double? ReadOptionalChartCharacterSpacing(XElement? runProperties)
    {
        return runProperties?.Attribute("spc") is { } spacing &&
            int.TryParse(spacing.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int spacingHundredths)
                ? spacingHundredths / 100d
                : null;
    }

}
