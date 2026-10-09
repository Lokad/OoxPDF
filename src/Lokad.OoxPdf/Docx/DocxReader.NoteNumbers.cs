using System.Xml.Linq;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static DocxDocumentSettings ResolveSingleSectionNoteReferenceSettings(
        XDocument document, DocxDocumentSettings settings)
    {
        XElement[] sections = document.Descendants(Ooxml.OoxNamespaces.WordprocessingNamespace + "sectPr").Take(2).ToArray();
        if (sections.Length != 1)
        {
            return settings;
        }

        // The body counters remain document-wide. Apply explicit single-section
        // overrides; section restarts and document-setting precedence keep fallback.
        return settings with
        {
            FootnoteReferenceSettings = Apply(settings.FootnoteReferenceSettings, "footnotePr"),
            EndnoteReferenceSettings = Apply(settings.EndnoteReferenceSettings, "endnotePr"),
        };

        DocxNoteReferenceSettings Apply(DocxNoteReferenceSettings authored, string name)
        {
            DocxNoteReferenceSettings section = ReadNoteReferenceSettings(
                sections[0].Element(Ooxml.OoxNamespaces.WordprocessingNamespace + name));
            if (section.NumberRestartValue is not (null or "continuous"))
            {
                return authored;
            }
            bool hasStart = section.NumberStart is > 0;
            bool hasFormat = section.NumberFormatValue is "decimal" or "decimalZero" or
                "lowerRoman" or "upperRoman" or "lowerLetter" or "upperLetter";
            if (!hasStart && !hasFormat)
            {
                return authored;
            }

            return authored with
            {
                NumberStartValue = hasStart ? section.NumberStartValue : authored.NumberStartValue,
                NumberStart = hasStart ? section.NumberStart : authored.NumberStart,
                NumberFormatValue = hasFormat ? section.NumberFormatValue : authored.NumberFormatValue,
            };
        }
    }
}
