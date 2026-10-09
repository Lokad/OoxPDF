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

        // The body counters remain document-wide. Word uses section starts/formats
        // and their defaults rather than the corresponding document settings.
        // Multi-section and restart numbering still retain the prior fallback.
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
            // Admit omitted or supported section properties only. Malformed starts,
            // unsupported formats and document restart requests keep the previous
            // partial override behavior, including its authored-property fallback.
            bool useSectionDefaults = (section.NumberStartValue is null || hasStart) &&
                (section.NumberFormatValue is null || hasFormat) &&
                authored.NumberRestartValue is null or "continuous";

            return authored with
            {
                NumberStartValue = hasStart ? section.NumberStartValue : useSectionDefaults ? null : authored.NumberStartValue,
                NumberStart = hasStart ? section.NumberStart : useSectionDefaults ? null : authored.NumberStart,
                NumberFormatValue = hasFormat ? section.NumberFormatValue : useSectionDefaults ? null : authored.NumberFormatValue,
            };
        }
    }
}
