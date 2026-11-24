using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxPlaceholderMatcher
{

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
        bool hasNumericIndex = uint.TryParse(index.AsSpan().Trim(), NumberStyles.None,
            CultureInfo.InvariantCulture, out uint numericIndex);
        foreach (XDocument source in placeholderSources)
        {
            if (PptxSceneBuilder.GetShapeTree(source) is { } tree && FindMatchingPlaceholderShape(tree) is { } match)
                matches.Add(match);
        }

        return matches;

        XElement? FindMatchingPlaceholderShape(XElement tree)
        {
            XElement? indexMatch = null;
            XElement? typeMatch = null;
            string normalizedType = NormalizePlaceholderType(type);
            foreach (XElement candidate in EnumerateShapes(tree))
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
                // Keep absent/malformed recovery unchanged. Valid unsigned
                // indices compare by value; their authored spelling survives.
                bool sameIndex = index is not null && (candidateIndex == index ||
                    hasNumericIndex && uint.TryParse(candidateIndex.AsSpan().Trim(), NumberStyles.None,
                        CultureInfo.InvariantCulture, out uint candidateNumericIndex) && candidateNumericIndex == numericIndex);
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
    }

    // Template placeholders can provide defaults even when their own paint is
    // suppressed. Walk drawing groups without entering unrelated markup.
    private static IEnumerable<XElement> EnumerateShapes(XElement owner)
    {
        foreach (XElement element in owner.Elements())
        {
            if (element.Name == PresentationNamespace + "sp")
                yield return element;
            else if (element.Name == PresentationNamespace + "grpSp")
            {
                foreach (XElement shape in EnumerateShapes(element))
                    yield return shape;
            }
        }
    }

    private static string NormalizePlaceholderType(string? type)
    {
        return string.IsNullOrEmpty(type) ? "body" : type;
    }
}
