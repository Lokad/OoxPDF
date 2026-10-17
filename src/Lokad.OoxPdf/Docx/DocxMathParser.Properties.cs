using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathParser
{
    private static DocxMathFraction ReadFraction(XElement element, DocxMathRow numerator, DocxMathRow denominator,
        Func<XElement, string, DocxTextRun> readRun, CancellationToken token)
    {
        bool valid = ReadPropertyOwner(element, "fPr", out XElement? owner, out XElement[] children);
        foreach (XElement property in children)
        {
            token.ThrowIfCancellationRequested();
            valid &= property.Name == MathNamespace + "ctrlPr" ? HasKnownControl(property) :
                property.Name == MathNamespace + "type" && HasKnownValueProperty(property);
        }
        XElement? control = owner?.Element(MathNamespace + "ctrlPr");
        DocxTextRun text = readRun(control ?? element, string.Empty);
        return new(numerator, denominator, PropertyValue(element, "fPr", "type"))
        { HasUnsupportedProperties = !valid, ControlText = text, ControlFontFamily = ControlFontFamily(control, text) };
    }

    private static DocxMathScripts ReadScripts(XElement element, DocxMathRow basis, DocxMathRow? subscript, DocxMathRow? superscript,
        Func<XElement, string, DocxTextRun> readRun, CancellationToken token)
    {
        bool valid = ReadPropertyOwner(element, element.Name.LocalName + "Pr", out XElement? owner, out XElement[] children);
        foreach (XElement property in children)
        { token.ThrowIfCancellationRequested(); valid &= property.Name == MathNamespace + "ctrlPr" && HasKnownControl(property); }
        XElement? control = owner?.Element(MathNamespace + "ctrlPr");
        DocxTextRun text = readRun(control ?? element, string.Empty);
        return new(basis, subscript, superscript)
        { HasUnsupportedProperties = !valid, ControlText = valid ? text : null, ControlFontFamily = ControlFontFamily(control, text) };
    }

    private static DocxMathFunction ReadFunction(XElement element, DocxMathRow name, DocxMathRow argument,
        Func<XElement, string, DocxTextRun> readRun, CancellationToken token)
    {
        bool valid = ReadPropertyOwner(element, "funcPr", out XElement? owner, out XElement[] children);
        foreach (var property in children)
        { token.ThrowIfCancellationRequested(); valid &= property.Name == MathNamespace + "ctrlPr" && HasKnownControl(property); }
        foreach (var role in ContentChildren(element))
        { token.ThrowIfCancellationRequested(); valid &= role.Attributes().All(a => a.IsNamespaceDeclaration) && HasOnlyWhitespaceText(role); }
        XElement? control = owner?.Element(MathNamespace + "ctrlPr");
        DocxTextRun text = readRun(control ?? element, string.Empty);
        return new(name, argument)
        { HasUnsupportedProperties = !valid, ControlText = text, ControlFontFamily = ControlFontFamily(control, text) };
    }

    private static bool HasKnownControl(XElement control) =>
        control.Attributes().All(a => a.IsNamespaceDeclaration) &&
        HasOnlyWhitespaceText(control) && control.Elements().Count() <= 1 &&
        control.Elements().All(p => p.Name == WordprocessingNamespace + "rPr" &&
            p.Attributes().All(a => a.IsNamespaceDeclaration) && HasOnlyWhitespaceText(p) &&
            p.Elements().All(HasKnownControlProperty));

    private static bool HasOnlyWhitespaceText(XElement element) =>
        string.IsNullOrWhiteSpace(string.Concat(element.Nodes().OfType<XText>().Select(t => t.Value)));

    private static string? ControlFontFamily(XElement? control, DocxTextRun text) =>
        control?.Element(WordprocessingNamespace + "rPr") is { } properties &&
        (properties.Element(WordprocessingNamespace + "rFonts") is not null ||
         properties.Element(WordprocessingNamespace + "rStyle") is not null) ? text.FontFamily : null;

    private static DocxMathDelimiter ReadDelimiter(XElement element, IReadOnlyList<DocxMathRow> arguments,
        Func<XElement, string, DocxTextRun> readRun, CancellationToken token)
    {
        bool valid = ReadPropertyOwner(element, "dPr", out XElement? owner, out XElement[] children);
        XElement? control = owner?.Element(MathNamespace + "ctrlPr");
        DocxTextRun text = readRun(control ?? element, string.Empty);
        foreach (XElement property in children)
        {
            token.ThrowIfCancellationRequested();
            if (property.Name == MathNamespace + "ctrlPr") { valid &= HasKnownControl(property); continue; }
            string? value = (string?)property.Attribute(MathNamespace + "val");
            valid &= HasKnownValueProperty(property) &&
                (property.Name.LocalName switch
                {
                    "begChr" or "endChr" or "sepChr" => HasCharacterValue(value),
                    "grow" => HasBooleanValue(value),
                    "shp" => value is null or "centered" or "match",
                    _ => false
                });
        }
        return new(arguments)
        {
            BeginCharacterValue = ReadCharacter(owner, "begChr"), EndCharacterValue = ReadCharacter(owner, "endChr"),
            SeparatorCharacterValue = ReadCharacter(owner, "sepChr"), GrowValue = ReadBoolean(owner, "grow"),
            ShapeValue = (string?)owner?.Element(MathNamespace + "shp")?.Attribute(MathNamespace + "val"),
            HasUnsupportedProperties = !valid, ControlText = text, ControlFontFamily = ControlFontFamily(control, text)
        };
    }

    private static DocxMathNary ReadNary(XElement element, DocxMathRow body, DocxMathRow lower, DocxMathRow upper,
        Func<XElement, string, DocxTextRun> readRun, CancellationToken token)
    {
        bool valid = ReadPropertyOwner(element, "naryPr", out XElement? owner, out XElement[] children);
        XElement? control = owner?.Element(MathNamespace + "ctrlPr");
        DocxTextRun text = readRun(control ?? element, string.Empty);
        foreach (XElement property in children)
        {
            token.ThrowIfCancellationRequested();
            if (property.Name == MathNamespace + "ctrlPr") { valid &= HasKnownControl(property); continue; }
            string? value = (string?)property.Attribute(MathNamespace + "val");
            valid &= HasKnownValueProperty(property) && (property.Name.LocalName switch
            {
                "chr" => HasCharacterValue(value),
                "grow" or "subHide" or "supHide" => HasBooleanValue(value),
                "limLoc" => value is null or "subSup" or "undOvr",
                _ => false
            });
        }
        return new(body, lower, upper)
        {
            CharacterValue = ReadCharacter(owner, "chr"), GrowValue = ReadBoolean(owner, "grow"),
            SubscriptHiddenValue = ReadBoolean(owner, "subHide"), SuperscriptHiddenValue = ReadBoolean(owner, "supHide"),
            LimitLocationValue = owner?.Element(MathNamespace + "limLoc") is { } location
                ? (string?)location.Attribute(MathNamespace + "val") ?? "undOvr" : null,
            HasUnsupportedProperties = !valid, ControlText = text, ControlFontFamily = ControlFontFamily(control, text)
        };
    }

    private static bool ReadPropertyOwner(XElement element, string name, out XElement? owner, out XElement[] children)
    {
        XElement[] properties = element.Elements().Where(IsProperty).ToArray();
        owner = element.Element(MathNamespace + name);children = owner?.Elements().ToArray() ?? [];
        return element.Attributes().All(a => a.IsNamespaceDeclaration) && HasOnlyWhitespaceText(element) &&
            properties.Length <= 1 && properties.All(p => p.Name == MathNamespace + name &&
                p.Attributes().All(a => a.IsNamespaceDeclaration) && HasOnlyWhitespaceText(p)) &&
            children.Select(p => p.Name).Distinct().Count() == children.Length;
    }

    internal static bool HasKnownValueProperty(XElement property) => property.Name.Namespace == MathNamespace &&
        !property.HasElements && string.IsNullOrWhiteSpace(property.Value) &&
        property.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == MathNamespace + "val");
    private static bool HasCharacterValue(string? value) => value is null || value.Length <= 2 && value.EnumerateRunes().Count() <= 1;
    private static bool HasBooleanValue(string? value) => value is null || OoxBoolean.IsTrue(value) || OoxBoolean.IsOff(value);
    private static string? ReadCharacter(XElement? owner, string name) => owner?.Element(MathNamespace + name) is { } character
        ? (string?)character.Attribute(MathNamespace + "val") ?? string.Empty : null;
    private static bool? ReadBoolean(XElement? owner, string name) => owner?.Element(MathNamespace + name) is { } value
        ? OoxBoolean.ParseElement(value, false, valueAttributeName: MathNamespace + "val") : null;
}
