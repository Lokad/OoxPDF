using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static IReadOnlyList<DocxRevisionInfo> ReadRunPropertyChangeRevisions(
        XElement? properties, string? paragraphStyleId, string? characterStyleId,
        DocxStyleSet styles, DocxResolvedRunProperties? tableStyle,
        DocxResolvedRunProperties current)
    {
        var revisions = ReadPropertyChangeRevisions(properties);
        if (properties is null || revisions.Count == 0) return revisions;
        return revisions.Select(revision => revision with
        {
            RunFormattingChange = ResolveChange(revision)
        }).ToArray();

        DocxRunFormattingChange? ResolveChange(DocxRevisionInfo revision)
        {
            // A complete supported change can use Word's visible description.
            // Other property families, uncertain source values and character-style
            // changes retain the existing generic summary, including author/date.
            if (revision.Kind != DocxRevisionKind.RunPropertiesChange ||
                revision.PropertyElementNames.Count == 0 ||
                revision.PropertyElementNames.Any(name => name is not ("b" or "i")) ||
                characterStyleId is not null)
                return null;
            var changes = properties.Elements(WordprocessingNamespace + "rPrChange")
                .Where(element => (string?)element.Attribute(WordprocessingNamespace + "id") == revision.Id).Take(2).ToArray();
            if (changes.Length != 1) return null;
            var previous = changes[0].Element(WordprocessingNamespace + "rPr");
            if (previous is null || !HasKnownBooleanValues(previous) || !HasKnownBooleanValues(properties) ||
                properties.Elements().Any(element => element.Name.Namespace != WordprocessingNamespace ||
                    element.Name.LocalName is not ("b" or "i" or "bCs" or "iCs" or "caps" or "smallCaps" or
                    "vanish" or "strike" or "dstrike" or "spacing" or "vertAlign" or "highlight" or "shd" or
                    "u" or "sz" or "color" or "rFonts" or "kern" or "rPrChange"))) return null;
            var before = ResolveRunProperties(previous, paragraphStyleId, null, styles, tableStyle);
            if ((before with { Bold = null, Italic = null }) != (current with { Bold = null, Italic = null })) return null;
            bool? bold = (before.Bold == true) == (current.Bold == true) ? null : current.Bold == true;
            bool? italic = (before.Italic == true) == (current.Italic == true) ? null : current.Italic == true;
            return bold is null && italic is null ? null : new(bold, italic);
        }

        static bool HasKnownBooleanValues(XElement properties)
        {
            foreach (string name in new[] { "b", "i" })
            {
                var values = properties.Elements(WordprocessingNamespace + name).ToArray();
                if (values.Length > 1) return false;
                if (values.SingleOrDefault() is { } value &&
                    (string?)value.Attribute(WordprocessingNamespace + "val") is { } token &&
                    token is not ("0" or "1" or "true" or "false" or "on" or "off")) return false;
            }
            return true;
        }
    }
}
