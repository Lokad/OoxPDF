using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static class DocxSourceXml
{
    // These wrappers contribute stored content without changing inline formatting.
    // Revisions, fields, links and directional/unknown owners need their own rules.
    public static bool IsInlineMetadataContainer(XElement element) =>
        element.Name == WordprocessingNamespace + "customXml" ||
        element.Name == WordprocessingNamespace + "smartTag";

    public static bool IsTransparentInlineContainer(XElement element) =>
        IsInlineMetadataContainer(element) || element.Name == WordprocessingNamespace + "sdt" ||
        element.Name == WordprocessingNamespace + "sdtContent";

    public static IEnumerable<XElement> StoredRuns(XElement source, CancellationToken token)
    {
        foreach (XElement child in StoredChildren(source))
        {
            token.ThrowIfCancellationRequested();
            if (child.Name == WordprocessingNamespace + "r") yield return child;
            else if (IsTransparentInlineContainer(child))
                foreach (XElement run in StoredRuns(child, token)) yield return run;
        }
    }

    // Editing boundaries and stored math pagination positions have no payload.
    // Content-bearing, unknown and misplaced children retain their source owner.
    public static bool IsNonPrintingMarker(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value) &&
        (element.Name.LocalName is "proofErr" or "permStart" or "permEnd" ||
            element.Name.LocalName == "lastRenderedPageBreak" && element.Parent?.Name == MathNamespace + "r" &&
            element.Attributes().All(attribute => attribute.IsNamespaceDeclaration));

    // Container properties describe the source, not stored inline payload.
    // Ignore only the property's actual owner; unknown/misplaced children stay.
    public static IEnumerable<XElement> StoredChildren(XElement element) =>
        element.Name == WordprocessingNamespace + "sdt"
            ? element.Elements(WordprocessingNamespace + "sdtContent")
            : element.Name == WordprocessingNamespace + "customXml"
                ? element.Elements().Where(child => child.Name != WordprocessingNamespace + "customXmlPr")
                : element.Name == WordprocessingNamespace + "smartTag"
                    ? element.Elements().Where(child => child.Name != WordprocessingNamespace + "smartTagPr")
                    : element.Elements();
}
