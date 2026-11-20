using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    // A lock preserves the stored result, including an empty result, while the
    // opcode remains available for source inspection. Nested fields own their locks.
    private static string? ResolveFieldPlaceholder(string? instruction, bool isLocked = false,
        bool requireOpcodeTerminator = false)
        => isLocked ? null : DocxFieldText.Placeholder(ResolveFieldKind(instruction, requireOpcodeTerminator));

    private static string? ResolveSimpleFieldPlaceholder(XElement field, DocxFieldKind kind) =>
        OoxBoolean.IsTrue((string?)field.Attribute(WordprocessingNamespace + "fldLock"))
            ? null : DocxFieldText.Placeholder(kind);

    private static bool HasSimpleFieldPlaceholder(XElement element) =>
        element.Name == WordprocessingNamespace + "fldSimple" &&
        ResolveSimpleFieldPlaceholder(element, ResolveFieldKind((string?)element.Attribute(WordprocessingNamespace + "instr"))) is not null;

    private sealed record ComputedSimpleFieldFragment(XElement Source, bool OwnsCounter);

    private static bool UsesDefaultDynamicFieldFormatting(string? instruction, DocxFieldKind kind) => kind switch
    {
        DocxFieldKind.Page => instruction.AsSpan().Trim().Equals("PAGE", StringComparison.OrdinalIgnoreCase),
        DocxFieldKind.NumPages => instruction.AsSpan().Trim().Equals("NUMPAGES", StringComparison.OrdinalIgnoreCase),
        _ => false
    };

    private static DocxFieldKind ResolveFieldKind(string? instruction, bool requireOpcodeTerminator = false)
    {
        ReadOnlySpan<char> trimmed = instruction.AsSpan().TrimStart();
        int length = 0;
        // A switch or whitespace can delimit the opcode. Other characters
        // are part of the token, so PAGE2 and NUMPAGES_ARCHIVE are stored fields.
        while (length < trimmed.Length && !char.IsWhiteSpace(trimmed[length]) && trimmed[length] != '\\')
        {
            length++;
        }
        if (requireOpcodeTerminator && length == trimmed.Length) return DocxFieldKind.Other;

        ReadOnlySpan<char> opcode = trimmed[..length];
        if (opcode.Equals("PAGE", StringComparison.OrdinalIgnoreCase)) return DocxFieldKind.Page;
        if (opcode.Equals("NUMPAGES", StringComparison.OrdinalIgnoreCase)) return DocxFieldKind.NumPages;
        return DocxFieldKind.Other;
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
        IsCachedFieldDrawing(element) &&
            element.Descendants().Any(content => content.Name == DrawingNamespace + "blip" || content.Name == VmlNamespace + "imagedata");

    private static bool IsCachedFieldDrawing(XElement element) =>
        element.Name == WordprocessingNamespace + "drawing" || element.Name == WordprocessingNamespace + "pict";

    private static bool FieldHasCachedResultContent(XElement field) => field.Descendants().Any(IsFieldResultContent);
}
