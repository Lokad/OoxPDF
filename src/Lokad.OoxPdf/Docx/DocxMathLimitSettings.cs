using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

// One immutable document owner; authored node properties remain separate.
internal sealed record DocxMathLimitSettings(string? NaryValue, string? IntegralValue,
    bool UnsupportedNary, bool UnsupportedIntegral, IReadOnlyList<string> PropertiesXml)
{
    public static DocxMathLimitSettings Empty { get; } = new(null, null, false, false, []);

    public static DocxMathLimitSettings Read(XElement? settings, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        string? nary = null, integral = null;
        int owners = 0, naryCount = 0, integralCount = 0;
        bool unsupportedNary = false, unsupportedIntegral = false, unsupportedOwner = false;
        var xml = new List<string>();
        if (settings is not null)
        foreach (var owner in settings.Elements(MathNamespace + "mathPr"))
        {
            token.ThrowIfCancellationRequested(); owners++;
            unsupportedOwner |= owner.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
                owner.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value));
            xml.Add(owner.ToString(SaveOptions.DisableFormatting));
            foreach (var property in owner.Elements())
            {
                token.ThrowIfCancellationRequested();
                if (property.Name.LocalName is not ("naryLim" or "intLim")) { continue; }
                string? value = (string?)property.Attribute(MathNamespace + "val");
                bool invalid = !DocxMathParser.HasKnownValueProperty(property) || value is not (null or "subSup" or "undOvr");
                if (property.Name.LocalName == "naryLim")
                { if (naryCount++ == 0) { nary = value; } unsupportedNary |= invalid; }
                else { if (integralCount++ == 0) { integral = value; } unsupportedIntegral |= invalid; }
            }
        }
        if (owners == 0) { return Empty; }
        return new(nary, integral, unsupportedNary || unsupportedOwner || owners > 1 || naryCount > 1,
            unsupportedIntegral || unsupportedOwner || owners > 1 || integralCount > 1, xml);
    }
}
