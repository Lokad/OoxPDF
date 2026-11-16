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

        // Supported paragraph properties inherit as complete choices. An absent
        // tab list inherits; an authored empty list clears the inherited stops.
        InheritParagraphChild("lnSpc");
        InheritParagraphChild("spcBef");
        InheritParagraphChild("spcAft");
        InheritParagraphChild("tabLst");
        MergeRendererChildElement(defaultRunPropertiesName, merged, fallback);
        MergeRendererBulletProperties(merged, fallback);
        return merged;

        void InheritParagraphChild(string name)
        {
            XName childName = merged.Name.Namespace + name;
            if (merged.Element(childName) is not null || fallback.Element(childName) is not { } inherited) return;
            AddParagraphChild(merged, inherited);
        }
    }

    private static void AddParagraphChild(XElement parent, XElement source)
    {
        XElement copy = new(source);
        int order = ParagraphChildOrder(source.Name.LocalName);
        XElement? next = parent.Elements().FirstOrDefault(child =>
            child.Name.Namespace == parent.Name.Namespace && ParagraphChildOrder(child.Name.LocalName) > order);
        if (next is null) parent.Add(copy);
        else next.AddBeforeSelf(copy);
    }

    private static int ParagraphChildOrder(string name) => name switch
    {
        "lnSpc" => 0,
        "spcBef" => 1,
        "spcAft" => 2,
        "buClrTx" or "buClr" => 3,
        "buSzTx" or "buSzPct" or "buSzPts" => 4,
        "buFontTx" or "buFont" => 5,
        "buNone" or "buChar" or "buAutoNum" or "buBlip" => 6,
        "tabLst" => 7,
        "defRPr" => 8,
        _ => 9
    };

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
            AddParagraphChild(primaryParent, fallback);
            return;
        }

        foreach (XAttribute attribute in fallback.Attributes())
        {
            if (primary.Attribute(attribute.Name) is null)
            {
                primary.Add(new XAttribute(attribute));
            }
        }

        bool hasPrimaryFill = primary.Elements().Any(IsFill);
        foreach (XElement child in fallback.Elements())
        {
            if (hasPrimaryFill && IsFill(child)) continue;
            if (primary.Element(child.Name) is null)
            {
                primary.Add(new XElement(child));
            }
        }

        bool IsFill(XElement child) => child.Name.Namespace == childName.Namespace &&
            child.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "blipFill" or "pattFill" or "grpFill";
    }

}
