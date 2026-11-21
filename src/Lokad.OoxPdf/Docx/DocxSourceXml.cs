using System.Diagnostics.CodeAnalysis;
using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static class DocxSourceXml
{
    public static bool IsRevisionContainer(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo";

    public static bool TryResolveRevisionRangeMarker(XElement element, [NotNullWhen(true)] out DocxRevisionKind? kind, out bool isStart)
    {
        (DocxRevisionKind? Kind, bool IsStart) marker = element.Name.Namespace == WordprocessingNamespace
            ? element.Name.LocalName switch
            {
                "moveFromRangeStart" => (DocxRevisionKind.MoveFrom, true),
                "moveFromRangeEnd" => (DocxRevisionKind.MoveFrom, false),
                "moveToRangeStart" => (DocxRevisionKind.MoveTo, true),
                "moveToRangeEnd" => (DocxRevisionKind.MoveTo, false),
                _ => (null, false)
            }
            : (null, false);
        (kind, isStart) = marker;
        return kind is not null;
    }

    public static bool IsNonPrintingRevisionRangeBoundary(XElement element) =>
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value) &&
        !string.IsNullOrWhiteSpace((string?)element.Attribute(WordprocessingNamespace + "id")) &&
        element.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name.Namespace == WordprocessingNamespace &&
            attribute.Name.LocalName is "id" or "name" or "author" or "date");

    // These wrappers contribute stored content without changing inline formatting.
    // Revisions, fields, links and directional/unknown owners need their own rules.
    public static bool IsInlineMetadataContainer(XElement element) =>
        element.Name == WordprocessingNamespace + "customXml" ||
        element.Name == WordprocessingNamespace + "smartTag";

    public static bool IsTransparentInlineContainer(XElement element) =>
        IsInlineMetadataContainer(element) || element.Name == WordprocessingNamespace + "sdt" ||
        element.Name == WordprocessingNamespace + "sdtContent";

    // Editing boundaries and stored math pagination positions have no payload.
    // Content-bearing, unknown and misplaced children retain their source owner.
    public static bool IsNonPrintingMarker(XElement element) =>
        element.Parent?.Name == MathNamespace + "r" && IsLastRenderedPageBreakHint(element) ||
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "proofErr" or "permStart" or "permEnd" &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value);

    public static bool IsLastRenderedPageBreakHint(XElement element) =>
        element.Name == WordprocessingNamespace + "lastRenderedPageBreak" &&
        (element.Parent?.Name == WordprocessingNamespace + "r" || element.Parent?.Name == MathNamespace + "r") &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value) &&
        element.Attributes().All(attribute => attribute.IsNamespaceDeclaration);

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
