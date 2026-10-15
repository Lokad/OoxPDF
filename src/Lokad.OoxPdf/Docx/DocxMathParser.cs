using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static class DocxMathParser
{
    // The package XML limits still apply. Use an explicit traversal stack so
    // accepted XML nesting does not also consume the managed call stack.
    public static DocxMathRow Read(
        XElement equation,
        Func<XElement, string, DocxTextRun> readRun,
        CancellationToken cancellationToken)
    {
        var pending = new Stack<(XElement Element, bool Expanded)>();
        var nodes = new Dictionary<XElement, DocxMathNode>();
        pending.Push((equation, false));
        while (pending.TryPop(out var item))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = item.Element;
            string name = element.Name.LocalName;
            if (element.Name.Namespace != MathNamespace || !IsContainer(name))
            {
                nodes.Add(element, ReadLeaf(element, readRun));
                continue;
            }

            if (!item.Expanded)
            {
                pending.Push((element, true));
                foreach (XElement child in ContentChildren(element).Reverse())
                {
                    pending.Push((child, false));
                }
                continue;
            }

            DocxMathNode node;
            if (IsRow(name) && (name != "deg" || element.Parent?.Name == MathNamespace + "rad"))
            {
                node = new DocxMathRow(ContentChildren(element).Select(child => nodes[child]).ToArray())
                { ArgumentProperties = ReadArgumentProperties(element) };
            }
            else if (name == "f" && Role(element, "num") is { } numerator && Role(element, "den") is { } denominator &&
                HasOnlyRoles(element, "num", "den"))
            {
                node = new DocxMathFraction(numerator, denominator, PropertyValue(element, "fPr", "type"))
                {
                    HasUnsupportedProperties = element.Elements().Any(child => IsProperty(child) && child.Name != MathNamespace + "fPr") ||
                        element.Element(MathNamespace + "fPr")?.Elements().Any(child => child.Name != MathNamespace + "type") == true
                };
            }
            else if (name is "sSub" or "sSup" or "sSubSup" && Role(element, "e") is { } basis &&
                (name == "sSup" || Role(element, "sub") is not null) &&
                (name == "sSub" || Role(element, "sup") is not null) &&
                HasOnlyRoles(element, name == "sSub" ? ["e", "sub"] : name == "sSup" ? ["e", "sup"] : ["e", "sub", "sup"]))
            {
                node = new DocxMathScripts(basis, Role(element, "sub"), Role(element, "sup"));
            }
            else if (name == "rad" && Role(element, "e") is { } radicand && Role(element, "deg") is { } degree &&
                HasOnlyRoles(element, "deg", "e"))
            {
                XElement[] properties = element.Elements().Where(IsProperty).ToArray();
                XElement[] visibility = element.Element(MathNamespace + "radPr")?.Elements(MathNamespace + "degHide").ToArray() ?? [];
                string? value = (string?)visibility.FirstOrDefault()?.Attribute(MathNamespace + "val");
                bool valid = properties.Length <= 1 && properties.All(p => p.Name == MathNamespace + "radPr" &&
                    p.Attributes().All(a => a.IsNamespaceDeclaration) &&
                    string.IsNullOrWhiteSpace(string.Concat(p.Nodes().OfType<XText>().Select(t => t.Value))) &&
                    p.Elements().All(child => child.Name == MathNamespace + "degHide")) &&
                    visibility.Length <= 1 && visibility.All(v => !v.HasElements && string.IsNullOrWhiteSpace(v.Value) &&
                        v.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == MathNamespace + "val")) &&
                    (value is null || Lokad.OoxPdf.Ooxml.OoxBoolean.IsTrue(value) || Lokad.OoxPdf.Ooxml.OoxBoolean.IsOff(value));
                node = new DocxMathRadical(radicand, degree, visibility.Length == 0 ? null :
                    Lokad.OoxPdf.Ooxml.OoxBoolean.ParseElement(visibility[0], false, valueAttributeName: MathNamespace + "val"))
                    { HasUnsupportedProperties = !valid };
            }
            else
            {
                // Do not serialize already traversed subtrees again: malformed
                // nested constructs must retain linear work and storage.
                node = new DocxMathUnsupported(name, element.Name.NamespaceName, null,
                    ContentChildren(element).Select(child => new DocxMathChild(
                        child.Name.LocalName, child.Name.NamespaceName, nodes[child])).ToArray());
            }
            nodes.Add(element, node with { PropertiesXml = ReadProperties(element) });
        }

        return nodes[equation] as DocxMathRow ?? new DocxMathRow([nodes[equation]]);

        DocxMathRow? Role(XElement element, string name)
        {
            XElement[] matches = element.Elements(MathNamespace + name).ToArray();
            return matches.Length == 1 ? nodes[matches[0]] as DocxMathRow : null;
        }
    }

    private static bool IsRow(string name) => name is "oMath" or "e" or "sub" or "sup" or "num" or "den" or "deg";
    private static bool IsContainer(string name) => IsRow(name) || name is "f" or "sSub" or "sSup" or "sSubSup" or "rad";
    private static bool IsProperty(XElement element) => element.Name.Namespace == MathNamespace && element.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal);
    private static IEnumerable<XElement> ContentChildren(XElement element) => element.Elements().Where(child => !IsProperty(child));
    private static bool HasOnlyRoles(XElement element, params string[] roles)
    {
        XElement[] children = ContentChildren(element).ToArray();
        return children.Length == roles.Length && children.All(child => child.Name.Namespace == MathNamespace && roles.Contains(child.Name.LocalName));
    }
    private static string? PropertyValue(XElement element, string property, string value) =>
        (string?)element.Element(MathNamespace + property)?.Element(MathNamespace + value)?.Attribute(MathNamespace + "val");
    private static string[] ReadProperties(XElement element) => element.Elements().Where(IsProperty)
        .Select(child => child.ToString(SaveOptions.DisableFormatting)).ToArray();

    private static DocxMathArgumentProperties? ReadArgumentProperties(XElement element)
    {
        XElement[] properties = element.Elements().Where(IsProperty).ToArray();
        if (properties.Length == 0) { return null; }
        if (properties.Length != 1 || properties[0].Name != MathNamespace + "argPr")
            return new(null, true);
        XElement argument = properties[0];
        XElement[] children = argument.Elements().ToArray();
        if (argument.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            !string.IsNullOrWhiteSpace(string.Concat(argument.Nodes().OfType<XText>().Select(text => text.Value))))
            return new(null, true);
        if (children.Length == 0) { return new(null, false); }
        if (children.Length != 1 || children[0].Name != MathNamespace + "argSz") { return new(null, true); }
        XElement size = children[0];
        if (size.HasElements || !string.IsNullOrWhiteSpace(size.Value) ||
            size.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != MathNamespace + "val"))
            return new(null, true);
        string? source = (string?)size.Attribute(MathNamespace + "val");
        int value = 0;
        if (source is not null && (!int.TryParse(source,
            System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value) || value is < -2 or > 2))
            return new(null, true);
        return new(value, false);
    }

    private static DocxMathNode ReadLeaf(XElement element, Func<XElement, string, DocxTextRun> readRun)
    {
        if (element.Name == MathNamespace + "r" && element.Elements().All(child =>
            child.Name == MathNamespace + "t" || child.Name == MathNamespace + "rPr" || child.Name == WordprocessingNamespace + "rPr"))
        {
            return new DocxMathRun(
                readRun(element, string.Concat(element.Elements(MathNamespace + "t").Select(text => text.Value))),
                PropertyValue(element, "rPr", "sty"), PropertyValue(element, "rPr", "scr"))
            {
                PropertiesXml = ReadProperties(element),
                NormalTextValue = element.Element(MathNamespace + "rPr")?.Element(MathNamespace + "nor") is { } authoredNormal
                    ? Lokad.OoxPdf.Ooxml.OoxBoolean.ParseElement(authoredNormal, false, valueAttributeName: MathNamespace + "val") : null,
                ExplicitFontFamily = (string?)element.Element(WordprocessingNamespace + "rPr")?.Element(WordprocessingNamespace + "rFonts")?.Attribute(WordprocessingNamespace + "ascii")
                    ?? (string?)element.Element(WordprocessingNamespace + "rPr")?.Element(WordprocessingNamespace + "rFonts")?.Attribute(WordprocessingNamespace + "hAnsi"),
                HasUnsupportedProperties = element.Element(MathNamespace + "rPr")?.Elements().Any(property =>
                    property.Name != MathNamespace + "sty" && property.Name != MathNamespace + "scr" &&
                    property.Name != MathNamespace + "nor") == true || !HasSupportedNormalTextProperty(element)
            };
        }
        return Unsupported(element);
    }

    private static bool HasSupportedNormalTextProperty(XElement element)
    {
        XElement[] properties = element.Element(MathNamespace + "rPr")?.Elements(MathNamespace + "nor").ToArray() ?? [];
        if (properties.Length == 0) { return true; }
        if (properties.Length != 1 || properties[0].HasElements || !string.IsNullOrWhiteSpace(properties[0].Value) ||
            properties[0].Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != MathNamespace + "val"))
            return false;
        string? value = (string?)properties[0].Attribute(MathNamespace + "val");
        return value is null || Lokad.OoxPdf.Ooxml.OoxBoolean.IsTrue(value) || Lokad.OoxPdf.Ooxml.OoxBoolean.IsOff(value);
    }

    private static DocxMathUnsupported Unsupported(XElement element) =>
        new(element.Name.LocalName, element.Name.NamespaceName, element.ToString(SaveOptions.DisableFormatting), []);
}
