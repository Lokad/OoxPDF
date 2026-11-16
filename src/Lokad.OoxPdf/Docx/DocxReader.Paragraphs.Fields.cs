using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    // A lock preserves the stored result, including an empty result, while the
    // opcode remains available for source inspection. Nested fields own their locks.
    private static string? ResolveFieldPlaceholder(string? instruction, bool isLocked = false,
        bool requireOpcodeTerminator = false)
        => isLocked ? null : DocxFieldText.Placeholder(ResolveFieldKind(instruction, requireOpcodeTerminator));

    private static bool UsesDefaultDynamicFieldFormatting(string? instruction, DocxFieldKind kind) => kind switch
    {
        DocxFieldKind.Page => instruction.AsSpan().Trim().Equals("PAGE", StringComparison.OrdinalIgnoreCase),
        DocxFieldKind.NumPages => instruction.AsSpan().Trim().Equals("NUMPAGES", StringComparison.OrdinalIgnoreCase),
        _ => false
    };

    private static DocxFieldKind ResolveFieldKind(string? instruction, bool requireOpcodeTerminator = false)
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

                return length == 0 || requireOpcodeTerminator && length == trimmed.Length
                    ? null : trimmed[..length].ToString().ToUpperInvariant();
            }

        string? opcode = ReadFieldOpcode(instruction);
        return opcode switch
        {
            "PAGE" => DocxFieldKind.Page,
            "NUMPAGES" => DocxFieldKind.NumPages,
            _ => DocxFieldKind.Other
        };
    }

    private static string? ReadCharacterStyleId(XElement run)
    {
        return (string?)run
            .Element(WordprocessingNamespace + "rPr")
            ?.Element(WordprocessingNamespace + "rStyle")
            ?.Attribute(WordprocessingNamespace + "val");
    }

    private static bool IsComplexFieldMarkupElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "fldChar" ||
            element.Name == WordprocessingNamespace + "instrText";
    }

    // Math and pictures are result items, rather than ordinary w:r text children.
    // Both field forms and the diagnostic scan use this source-content rule.
    private static bool IsFieldResultContent(XElement element) =>
        element.Name == MathNamespace + "oMath" || element.Name == MathNamespace + "oMathPara" ||
        ReadRunTextChild(element).Length != 0 ||
        (element.Name == WordprocessingNamespace + "drawing" || element.Name == WordprocessingNamespace + "pict") &&
            element.Descendants().Any(content => content.Name == DrawingNamespace + "blip" || content.Name == VmlNamespace + "imagedata");

    private static bool FieldHasCachedResultContent(XElement field) => field.Descendants().Any(IsFieldResultContent);
}
