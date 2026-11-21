using System.Xml.Linq;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private sealed record DocxNoteReferenceSection(
        DocxDocumentSettings Settings, bool Footnotes, bool Endnotes);

    private sealed class DocxNoteReferenceCounters
    {
        private readonly Dictionary<DocxRelatedStoryKind, int> _current = [];
        private readonly Dictionary<DocxRelatedStoryKind, int> _offsets = [];

        public void EnterSection(DocxNoteReferenceSection section)
        {
            Enter(DocxRelatedStoryKind.Footnote, section.Footnotes, section.Settings.FootnoteReferenceSettings);
            Enter(DocxRelatedStoryKind.Endnote, section.Endnotes, section.Settings.EndnoteReferenceSettings);
        }

        private void Enter(DocxRelatedStoryKind kind, bool admitted, DocxNoteReferenceSettings settings)
        {
            if (admitted)
            {
                _current.TryGetValue(kind, out int current);
                _offsets[kind] = settings.NumberRestartValue == "eachSect" ? current : 0;
            }
        }

        public int Next(DocxRelatedStoryKind kind, DocxNoteReferenceSettings settings)
        {
            _current.TryGetValue(kind, out int current);
            if (_offsets.TryGetValue(kind, out int offset))
            {
                // Office applies the current section's start to an occurrence index,
                // not to the previous section's visible number. Continuous numbering
                // counts earlier notes even when those sections restarted their labels.
                _current[kind] = current + 1;
                return (settings.NumberStart ?? 1) + current - offset;
            }

            int next = current == 0 ? settings.NumberStart ?? 1 : current + 1;
            _current[kind] = next;
            return next;
        }
    }

    private static IReadOnlyDictionary<XElement, DocxNoteReferenceSection> ResolveSectionNoteReferenceSettings(
        XDocument document, DocxDocumentSettings settings, DocxDocumentSettings fallback,
        OoxPdfDocxMarkupMode markupMode, CancellationToken cancellationToken)
    {
        var result = new Dictionary<XElement, DocxNoteReferenceSection>();
        var word = Ooxml.OoxNamespaces.WordprocessingNamespace;
        XName sectionName = word + "sectPr";
        XName sectionChangeName = word + "sectPrChange";
        XName importedContentName = word + "altChunk";
        XElement? body = document.Root?.Element(word + "body");
        XElement? terminal = body?.Elements().LastOrDefault();
        if (body is null || terminal?.Name != sectionName)
        {
            return result;
        }
        foreach (XElement element in body.Descendants())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (DocxSourceXml.IsRevisionContainer(element) || element.Name == sectionChangeName)
                return result;
        }

        var sections = new List<(XElement? First, XElement Properties)>();
        XElement? first = null;
        // Use the same original block owners as body reading through transparent
        // containers. Unexpanded imports and section history retain the fallback.
        foreach (var scope in EnumerateRevisionScopedChildren(body.Elements(), markupMode, cancellationToken,
                     [word + "p", word + "tbl", sectionName, importedContentName]))
        {
            XElement child = scope.Element;
            if (child.Name == importedContentName || child.Name == sectionName && child != terminal)
                return result;

            if (child.Name != sectionName) first ??= child;
            XElement? properties = child == terminal ? child : child.Element(word + "pPr")?.Element(sectionName);
            if (properties is not null)
            {
                sections.Add((first, properties));
                first = null;
            }
        }
        // A section inside table or otherwise unplaced content has no body owner.
        if (!document.Descendants(sectionName).SequenceEqual(sections.Select(s => s.Properties)))
        {
            return result;
        }

        bool footnotes = Admits("footnotePr", settings.FootnoteReferenceSettings);
        bool endnotes = Admits("endnotePr", settings.EndnoteReferenceSettings);
        foreach ((XElement? start, XElement properties) in sections)
        {
            if (start is null) continue;
            result.Add(start, new DocxNoteReferenceSection(settings with
            {
                FootnoteReferenceSettings = footnotes ? Effective(properties, "footnotePr", settings.FootnoteReferenceSettings) : fallback.FootnoteReferenceSettings,
                EndnoteReferenceSettings = endnotes ? Effective(properties, "endnotePr", settings.EndnoteReferenceSettings) : fallback.EndnoteReferenceSettings,
            }, footnotes, endnotes));
        }
        return result;

        bool Admits(string name, DocxNoteReferenceSettings authored) =>
            // Document restart metadata does not supply omitted section values.
            // Section eachPage still requires rendered page ownership and keeps
            // the existing fallback below; document eachPage is not that request.
            authored.HasKnownNumberRestart &&
            sections.All(s =>
            {
                DocxNoteReferenceSettings section = ReadNoteReferenceSettings(s.Properties.Element(word + name));
                return (section.NumberRestartValue is null or "continuous" or "eachSect") &&
                    (section.NumberStartValue is null || section.NumberStart is > 0 and <= 32767) &&
                    (section.NumberFormatValue is null or "decimal" or "decimalZero" or
                        "lowerRoman" or "upperRoman" or "lowerLetter" or "upperLetter");
            });

        DocxNoteReferenceSettings Effective(XElement properties, string name, DocxNoteReferenceSettings authored)
        {
            DocxNoteReferenceSettings section = ReadNoteReferenceSettings(properties.Element(word + name));
            return authored with
            {
                NumberStartValue = section.NumberStartValue,
                NumberStart = section.NumberStart,
                NumberFormatValue = section.NumberFormatValue,
                NumberRestartValue = section.NumberRestartValue,
            };
        }
    }

    private static DocxDocumentSettings ResolveSingleSectionNoteReferenceSettings(
        XDocument document, DocxDocumentSettings settings)
    {
        XElement[] sections = document.Descendants(Ooxml.OoxNamespaces.WordprocessingNamespace + "sectPr").Take(2).ToArray();
        if (sections.Length != 1)
        {
            return settings;
        }

        // Preserve the prior single-section partial overrides when the bounded
        // section-occurrence policy is unavailable. This fallback never resets
        // counters or resolves multiple sections.
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
