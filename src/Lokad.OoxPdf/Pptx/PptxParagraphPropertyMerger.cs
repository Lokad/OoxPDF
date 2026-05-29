using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal static class PptxParagraphPropertyMerger
{
    public static XElement? MergeRendererDefaultProperties(XName defaultRunPropertiesName, params XElement?[] sources)
    {
        XElement? merged = null;
        foreach (XElement? source in sources.Reverse())
        {
            if (source is null)
            {
                continue;
            }

            merged = merged is null
                ? new XElement(source)
                : OverlayRendererParagraphProperties(defaultRunPropertiesName, source, merged);
        }

        return merged;
    }

    public static XElement? MergeSceneDefaultProperties(params XElement?[] sources)
    {
        XElement? merged = null;
        foreach (XElement source in sources.Reverse().Where(source => source is not null).Cast<XElement>())
        {
            merged ??= new XElement(source.Name);
            foreach (XAttribute attribute in source.Attributes())
            {
                merged.SetAttributeValue(attribute.Name, attribute.Value);
            }

            foreach (XElement child in source.Elements())
            {
                XElement? existing = merged.Element(child.Name);
                if (existing is null)
                {
                    merged.Add(new XElement(child));
                    continue;
                }

                OverlaySceneChildElement(existing, child);
            }
        }

        return merged;
    }

    private static XElement OverlayRendererParagraphProperties(XName defaultRunPropertiesName, XElement primary, XElement fallback)
    {
        XElement merged = new(primary);
        foreach (XAttribute attribute in fallback.Attributes())
        {
            if (merged.Attribute(attribute.Name) is null)
            {
                merged.Add(new XAttribute(attribute));
            }
        }

        MergeRendererChildElement(defaultRunPropertiesName, merged, fallback);
        MergeRendererBulletProperties(merged, fallback);
        return merged;
    }

    private static void MergeRendererBulletProperties(XElement primaryParent, XElement fallbackParent)
    {
        XNamespace ns = primaryParent.Name.Namespace;
        XName[] bulletKindNames =
        [
            ns + "buNone",
            ns + "buChar",
            ns + "buAutoNum",
            ns + "buBlip"
        ];
        XName[] bulletPropertyNames =
        [
            ns + "buFont",
            ns + "buClr",
            ns + "buSzPct",
            ns + "buSzPts"
        ];

        foreach (XName propertyName in bulletPropertyNames)
        {
            if (primaryParent.Element(propertyName) is null &&
                fallbackParent.Element(propertyName) is { } fallbackProperty)
            {
                primaryParent.Add(new XElement(fallbackProperty));
            }
        }

        if (bulletKindNames.Any(name => primaryParent.Element(name) is not null))
        {
            return;
        }

        foreach (XName kindName in bulletKindNames)
        {
            if (fallbackParent.Element(kindName) is { } fallbackKind)
            {
                primaryParent.Add(new XElement(fallbackKind));
                return;
            }
        }
    }

    private static void MergeRendererChildElement(XName childName, XElement primaryParent, XElement fallbackParent)
    {
        XElement? primary = primaryParent.Element(childName);
        XElement? fallback = fallbackParent.Element(childName);
        if (fallback is null)
        {
            return;
        }

        if (primary is null)
        {
            primaryParent.Add(new XElement(fallback));
            return;
        }

        foreach (XAttribute attribute in fallback.Attributes())
        {
            if (primary.Attribute(attribute.Name) is null)
            {
                primary.Add(new XAttribute(attribute));
            }
        }

        foreach (XElement child in fallback.Elements())
        {
            if (primary.Element(child.Name) is null)
            {
                primary.Add(new XElement(child));
            }
        }
    }

    private static void OverlaySceneChildElement(XElement target, XElement source)
    {
        foreach (XAttribute attribute in source.Attributes())
        {
            target.SetAttributeValue(attribute.Name, attribute.Value);
        }

        foreach (XElement child in source.Elements())
        {
            XElement? existing = target.Element(child.Name);
            if (existing is null)
            {
                target.Add(new XElement(child));
                continue;
            }

            existing.ReplaceWith(new XElement(child));
        }
    }
}
