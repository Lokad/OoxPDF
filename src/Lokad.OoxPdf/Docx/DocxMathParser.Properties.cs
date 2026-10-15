using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static partial class DocxMathParser
{
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
        XElement[] properties = element.Elements().Where(IsProperty).ToArray();
        XElement? owner = element.Element(MathNamespace + "dPr");
        XElement[] children = owner?.Elements().ToArray() ?? [];
        XElement? control = owner?.Element(MathNamespace + "ctrlPr");
        DocxTextRun text = readRun(control ?? element, string.Empty);
        bool valid = element.Attributes().All(a => a.IsNamespaceDeclaration) && HasOnlyWhitespaceText(element) &&
            properties.Length <= 1 && properties.All(p => p.Name == MathNamespace + "dPr" &&
                p.Attributes().All(a => a.IsNamespaceDeclaration) && HasOnlyWhitespaceText(p)) &&
            children.Select(p => p.Name).Distinct().Count() == children.Length;
        foreach (XElement property in children)
        {
            token.ThrowIfCancellationRequested();
            if (property.Name == MathNamespace + "ctrlPr") { valid &= HasKnownControl(property); continue; }
            string? value = (string?)property.Attribute(MathNamespace + "val");
            valid &= property.Name.Namespace == MathNamespace && !property.HasElements &&
                string.IsNullOrWhiteSpace(property.Value) &&
                property.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == MathNamespace + "val") &&
                (property.Name.LocalName switch
                {
                    "begChr" or "endChr" or "sepChr" => value is null || value.Length <= 2 && value.EnumerateRunes().Count() <= 1,
                    "grow" => value is null || OoxBoolean.IsTrue(value) || OoxBoolean.IsOff(value),
                    "shp" => value is null or "centered" or "match",
                    _ => false
                });
        }
        string? Character(string name) => owner?.Element(MathNamespace + name) is { } character
            ? (string?)character.Attribute(MathNamespace + "val") ?? string.Empty : null;
        XElement? grow = owner?.Element(MathNamespace + "grow");
        return new(arguments)
        {
            BeginCharacterValue = Character("begChr"), EndCharacterValue = Character("endChr"),
            SeparatorCharacterValue = Character("sepChr"),
            GrowValue = grow is null ? null : OoxBoolean.ParseElement(grow, false, valueAttributeName: MathNamespace + "val"),
            ShapeValue = (string?)owner?.Element(MathNamespace + "shp")?.Attribute(MathNamespace + "val"),
            HasUnsupportedProperties = !valid, ControlText = text, ControlFontFamily = ControlFontFamily(control, text)
        };
    }
}
