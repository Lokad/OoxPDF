using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static string? ResolveFieldPlaceholder(string? instruction)
    {
        return ResolveFieldKind(instruction) switch
        {
            DocxFieldKind.NumPages => "{NUMPAGES}",
            DocxFieldKind.Page => "{PAGE}",
            _ => null
        };
    }

    private static DocxFieldKind ResolveFieldKind(string? instruction)
    {
            string? ReadFieldOpcode(string? instruction)
            {
                if (string.IsNullOrWhiteSpace(instruction))
                {
                    return null;
                }

                ReadOnlySpan<char> trimmed = instruction.AsSpan().TrimStart();
                int length = 0;
                while (length < trimmed.Length && char.IsLetter(trimmed[length]))
                {
                    length++;
                }

                return length == 0 ? null : trimmed[..length].ToString().ToUpperInvariant();
            }

        string? opcode = ReadFieldOpcode(instruction);
        return opcode switch
        {
            "PAGE" => DocxFieldKind.Page,
            "NUMPAGES" => DocxFieldKind.NumPages,
            _ => DocxFieldKind.Other
        };
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static string? ReadCharacterStyleId(XElement run)
    {
        return (string?)run
            .Element(WordprocessingNamespace + "rPr")
            ?.Element(WordprocessingNamespace + "rStyle")
            ?.Attribute(WordprocessingNamespace + "val");
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static bool IsComplexFieldMarkupElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "fldChar" ||
            element.Name == WordprocessingNamespace + "instrText";
    }

    private sealed record DocxVmlTextBoxContent(string Text, IReadOnlyList<DocxRevisionInfo> Revisions);

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IEnumerable<DocxVmlTextBoxContent> ReadVmlTextBoxContents(XElement run)
    {
        foreach (XElement textBox in run.Descendants(VmlNamespace + "textbox"))
        {
            foreach (XElement content in textBox.Descendants(WordprocessingNamespace + "txbxContent"))
            {
                DocxVmlTextBoxContent textBoxContent = ReadTextBoxContent(content);
                if (textBoxContent.Text.Length != 0)
                {
                    yield return textBoxContent;
                }
            }
        }
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxVmlTextBoxContent ReadTextBoxContent(XElement content)
    {
        var paragraphs = new List<string>();
        var revisions = new List<DocxRevisionInfo>();
        foreach (XElement paragraph in content.Elements(WordprocessingNamespace + "p"))
        {
            AddRevisions(revisions, ReadPropertyChangeRevisions(paragraph.Element(WordprocessingNamespace + "pPr")));
            string paragraphText = string.Concat(
                paragraph
                    .Descendants(WordprocessingNamespace + "r")
                    .Select(ReadRunText));
            if (paragraphText.Length != 0)
            {
                paragraphs.Add(paragraphText);
            }
        }

        return new DocxVmlTextBoxContent(string.Join("\n", paragraphs), revisions);
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static bool FieldHasCachedResultText(XElement field)
    {
        return field
            .Descendants()
            .Any(element => ReadRunTextChild(element).Length != 0);
    }
}
