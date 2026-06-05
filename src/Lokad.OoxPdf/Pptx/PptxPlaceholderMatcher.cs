using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxPlaceholderMatcher
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";

    internal static IReadOnlyList<XElement> FindInheritedPlaceholderShapes(XElement shape, IReadOnlyList<XDocument> placeholderSources)
    {
        XElement? placeholder = shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph");
        if (placeholder is null)
        {
            return [];
        }

        var matches = new List<XElement>();
        string? type = (string?)placeholder.Attribute("type");
        string? index = (string?)placeholder.Attribute("idx");
        foreach (XDocument source in placeholderSources)
        {
            XElement[] candidates = source.Descendants(PresentationNamespace + "sp").ToArray();
            XElement? match = FindMatchingPlaceholderShape(candidates, type, index);
            if (match is not null)
            {
                matches.Add(match);
            }
        }

        return matches;
    }

    private static XElement? FindMatchingPlaceholderShape(IReadOnlyList<XElement> candidates, string? type, string? index)
    {
        XElement? indexMatch = null;
        XElement? typeMatch = null;
        string normalizedType = NormalizePlaceholderType(type);
        foreach (XElement candidate in candidates)
        {
            XElement? candidatePlaceholder = candidate
                .Element(PresentationNamespace + "nvSpPr")
                ?.Element(PresentationNamespace + "nvPr")
                ?.Element(PresentationNamespace + "ph");
            if (candidatePlaceholder is null)
            {
                continue;
            }

            string? candidateType = (string?)candidatePlaceholder.Attribute("type");
            string? candidateIndex = (string?)candidatePlaceholder.Attribute("idx");
            bool sameIndex = index is not null && candidateIndex == index;
            bool exactType = type is not null && candidateType == type;
            bool sameType = NormalizePlaceholderType(candidateType) == normalizedType;
            if (sameIndex && exactType)
            {
                return candidate;
            }

            if (sameIndex)
            {
                indexMatch ??= candidate;
            }

            if (sameType)
            {
                typeMatch ??= candidate;
            }
        }

        return indexMatch ?? typeMatch;
    }

    private static string NormalizePlaceholderType(string? type)
    {
        return string.IsNullOrEmpty(type) ? "body" : type;
    }
}
