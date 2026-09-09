using System.Xml.Linq;

using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Ooxml;

// ISO 29500 Markup Compatibility subset (O02): exactly one AlternateContent
// representation survives load, so Descendants traversals can neither duplicate
// Choice and Fallback content nor silently keep both. The first Choice whose
// Requires prefixes all resolve to understood namespaces wins, else the Fallback,
// else the block is removed. Ignorable content needs no handling (unknown elements
// are already skipped, which subsumes ProcessContent); layout/master parts stay open.
internal static class OoxMarkupCompatibility
{
    private static readonly HashSet<string> UnderstoodNamespaces = new(
        [
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/chart",

            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",

            "urn:schemas-microsoft-com:vml"
        ],
        StringComparer.Ordinal);

    // MustUnderstand scan (O02): true when an element demands a namespace outside
    // the understood set. Readers warn once per document; slides and DOCX story parts warn with their part names.
    public static bool HasUnrecognizedMustUnderstand(XDocument? document)
    {
        if (document?.Root is null)
        {
            return false;
        }

        XNamespace compatibility = OoxNamespaces.MarkupCompatibilityNamespace;
        foreach (XElement element in document.Descendants())
        {
            string? demand = (string?)element.Attribute(compatibility + "MustUnderstand");
            if (!string.Equals(demand, "1", StringComparison.Ordinal) &&
                !string.Equals(demand, "true", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!UnderstoodNamespaces.Contains(element.Name.NamespaceName))
            {
                return true;
            }
        }

        return false;
    }

    // O02: one must-understand warning per part per document. Slides warn with
    // their slide index; shared master/layout parts warn once no matter how many
    // slides use them; DOCX story parts warn without one. Repeat loads of the
    // same part stay quiet through the gate.
    public static void WarnMustUnderstandOnce(
        XDocument? document,
        string? partName,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<string>? warnedParts,
        int? slideIndex = null)
    {
        if (diagnosticSink is null || warnedParts is null || !HasUnrecognizedMustUnderstand(document))
        {
            return;
        }

        if (!warnedParts.Add(partName ?? ""))
        {
            return;
        }

        diagnosticSink(new OoxPdfDiagnostic(
            "OOXML_MUST_UNDERSTAND",
            OoxPdfSeverity.Warning,
            "Content marked must-understand uses unsupported namespaces and was ignored.",
            partName,
            SlideIndex: slideIndex,
            PageIndex: null,
            Feature: "must-understand",
            Fallback: "Ignored"));
    }

    public static void ResolveAlternateContent(XDocument? document)
    {
        if (document?.Root is null)
        {
            return;
        }

        // Reversed document order normalizes innermost blocks before their parents.
        foreach (XElement alternate in document.Descendants(OoxNamespaces.MarkupCompatibilityNamespace + "AlternateContent").Reverse().ToArray())
        {
            ResolveAlternateContentBlock(alternate);
        }
    }

    private static void ResolveAlternateContentBlock(XElement alternate)
    {
        XNamespace compatibility = OoxNamespaces.MarkupCompatibilityNamespace;
        foreach (XElement choice in alternate.Elements(compatibility + "Choice"))
        {
            if (IsUnderstoodChoice(choice))
            {
                alternate.ReplaceWith(choice.Elements().ToArray());
                return;
            }
        }

        XElement? fallback = alternate.Element(compatibility + "Fallback");
        if (fallback is null)
        {
            alternate.Remove();
            return;
        }

        alternate.ReplaceWith(fallback.Elements().ToArray());
    }

    private static bool IsUnderstoodChoice(XElement choice)
    {
        string requires = (string?)choice.Attribute("Requires") ?? string.Empty;
        string[] prefixes = requires.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        if (prefixes.Length == 0)
        {
            return true;
        }

        foreach (string prefix in prefixes)
        {
            if (!UnderstoodNamespaces.Contains(choice.GetNamespaceOfPrefix(prefix)?.NamespaceName ?? string.Empty))
            {
                return false;
            }
        }

        return true;
    }
}
