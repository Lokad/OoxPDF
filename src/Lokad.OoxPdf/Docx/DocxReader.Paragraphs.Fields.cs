using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    // A lock or an unclosed field keeps the result literal. The opcode remains
    // available for source inspection.
    private static string? ResolveFieldPlaceholder(string? instruction, bool preserveStoredResult = false,
        bool requireOpcodeTerminator = false)
        => preserveStoredResult ? null : DocxFieldText.Placeholder(ResolveFieldKind(instruction, requireOpcodeTerminator));

    private static string? ResolveSimpleFieldPlaceholder(XElement field, DocxFieldKind kind) =>
        OoxBoolean.IsTrue((string?)field.Attribute(WordprocessingNamespace + "fldLock"))
            ? null : DocxFieldText.Placeholder(kind);

    private static bool HasSimpleFieldPlaceholder(XElement element) =>
        element.Name == WordprocessingNamespace + "fldSimple" &&
        ResolveSimpleFieldPlaceholder(element, ResolveFieldKind((string?)element.Attribute(WordprocessingNamespace + "instr"))) is not null;

    private sealed record ComputedSimpleFieldFragment(XElement Source, bool OwnsCounter);

    private sealed record ComplexFieldClosure(OoxPdfDocxMarkupMode Mode, IReadOnlySet<XElement> Unclosed);
    private sealed class UnclosedFieldBegin;

    // Pair visible field boundaries in their authored story before a counter
    // can replace its stored result. A generated page fragment is not a story.
    private static IReadOnlySet<XElement>? ReadUnclosedFieldBegins(XElement? story, OoxPdfDocxMarkupMode markupMode,
        CancellationToken token)
    {
        if (story is null) return null;
        foreach (var cached in story.Annotations<ComplexFieldClosure>())
            if (cached.Mode == markupMode) return cached.Unclosed;

        var open = new Stack<XElement>();
        foreach (XElement marker in story.Descendants())
        {
            token.ThrowIfCancellationRequested();
            if (marker.Name != WordprocessingNamespace + "fldChar" || marker.Parent?.Name != WordprocessingNamespace + "r" ||
                DocxSourceXml.FindStoryRoot(marker) != story ||
                IsInsideExcludedRevisionContainer(marker, markupMode)) continue;
            XElement? container = marker.Parent.Parent;
            while (container is not null && IsVisibleRunContainer(container, markupMode))
                container = container.Parent;
            if (container?.Name != WordprocessingNamespace + "p") continue;

            string? type = (string?)marker.Attribute(WordprocessingNamespace + "fldCharType");
            if (string.Equals(type, "begin", StringComparison.OrdinalIgnoreCase))
                open.Push(marker);
            else if (string.Equals(type, "end", StringComparison.OrdinalIgnoreCase))
                open.TryPop(out _);
        }
        token.ThrowIfCancellationRequested();
        var unclosed = open.ToHashSet();
        story.AddAnnotation(new ComplexFieldClosure(markupMode, unclosed));
        return unclosed;
    }

    private static XElement CopyFieldAwareFragment(XElement source, IReadOnlySet<XElement>? unclosed)
    {
        var copy = new XElement(source);
        if (unclosed is not { Count: > 0 }) return copy;
        foreach (var (original, marker) in source.DescendantsAndSelf(WordprocessingNamespace + "fldChar")
                     .Zip(copy.DescendantsAndSelf(WordprocessingNamespace + "fldChar")))
            if (unclosed.Contains(original)) marker.AddAnnotation(new UnclosedFieldBegin());
        return copy;
    }

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

    private static bool IsFieldInstructionText(XElement element) =>
        element.Name == WordprocessingNamespace + "instrText" ||
        element.Name == WordprocessingNamespace + "delInstrText";

    // Review uses the current opcode on a current field. A deleted field
    // shown in Review keeps the opcode it had before the deletion.
    private static bool UsesOriginalFieldInstruction(OoxPdfDocxMarkupMode mode, bool deletedField) =>
        mode == OoxPdfDocxMarkupMode.Original || deletedField;

    private static bool IncludesFieldInstruction(XElement instruction, bool usesOriginalInstruction, bool inserted) =>
        instruction.Name == WordprocessingNamespace + "delInstrText" ? usesOriginalInstruction : !usesOriginalInstruction || !inserted;

    private static bool IsComplexFieldMarkupElement(XElement element) =>
        element.Name == WordprocessingNamespace + "fldChar" || IsFieldInstructionText(element);

    // Math and pictures are result items, rather than ordinary w:r text children.
    // Both field forms and the diagnostic scan use this source-content rule.
    private static bool IsFieldResultContent(XElement element, bool inFieldResult = false) =>
        element.Name == MathNamespace + "oMath" || element.Name == MathNamespace + "oMathPara" ||
        ReadRunTextChild(element, inFieldResult).Length != 0 ||
        IsCachedFieldDrawing(element) &&
            element.Descendants().Any(content => content.Name == DrawingNamespace + "blip" || content.Name == VmlNamespace + "imagedata");

    private static bool IsCachedFieldDrawing(XElement element) =>
        element.Name == WordprocessingNamespace + "drawing" || element.Name == WordprocessingNamespace + "pict";

    private static bool FieldHasCachedResultContent(XElement field) => field.Descendants().Any(element => IsFieldResultContent(element));
}
