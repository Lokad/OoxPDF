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
        XName[][] bulletPropertyChoices =
        [
            [ns + "buClrTx", ns + "buClr"],
            [ns + "buSzTx", ns + "buSzPct", ns + "buSzPts"],
            [ns + "buFontTx", ns + "buFont"]
        ];

        foreach (XName[] choices in bulletPropertyChoices)
        {
            if (FindProperty(primaryParent, choices) is null &&
                FindProperty(fallbackParent, choices) is { } fallbackProperty)
            {
                AddBeforeBulletOrTail(fallbackProperty);
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
                AddBeforeBulletOrTail(fallbackKind);
                return;
            }
        }

        XElement? FindProperty(XElement parent, XName[] choices)
        {
            // Authored properties after a marker retain their existing ignored
            // fallback. Inherited properties must join the same pre-marker scope.
            return parent.Elements().TakeWhile(element => !bulletKindNames.Contains(element.Name))
                .FirstOrDefault(element => choices.Contains(element.Name));
        }

        void AddBeforeBulletOrTail(XElement property)
        {
            XElement copy = new(property);
            XElement? next = primaryParent.Elements().FirstOrDefault(element =>
                bulletKindNames.Contains(element.Name) || element.Name == ns + "tabLst" ||
                element.Name == ns + "defRPr" || element.Name == ns + "extLst");
            if (next is not null) next.AddBeforeSelf(copy);
            else primaryParent.Add(copy);
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

}
