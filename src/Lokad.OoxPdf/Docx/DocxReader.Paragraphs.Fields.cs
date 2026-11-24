using System.Text;
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

    private sealed record ComplexFieldFlow(OoxPdfDocxMarkupMode Mode, IReadOnlySet<XElement> Unclosed, bool HasUnsupported);
    private sealed record DocxComplexFieldScanState(XElement Begin, StringBuilder Instruction,
        bool UsesOriginalInstruction, bool HasSeparate, bool HasCachedResult);
    private sealed class UnclosedFieldBegin;

    private static bool HasUnsupportedComplexFields(XDocument document, OoxPdfDocxMarkupMode markupMode, CancellationToken token)
    {
        var inspected = new HashSet<XElement>();
        foreach (XElement marker in document.Descendants())
        {
            token.ThrowIfCancellationRequested();
            if (!IsComplexFieldMarkupElement(marker) || IsInsideExcludedRevisionContainer(marker, markupMode)) continue;
            if (marker.Parent?.Name != WordprocessingNamespace + "r") return true;
            XElement? container = marker.Parent.Parent;
            while (container is not null && IsVisibleRunContainer(container, markupMode))
                container = container.Parent;
            if (container?.Name != WordprocessingNamespace + "p") return true;
            if (DocxSourceXml.FindFieldFlowOwner(container) is { } owner && inspected.Add(owner) &&
                ReadComplexFieldFlow(owner, markupMode, token).HasUnsupported) return true;
        }
        return false;
    }

    // A field can span projected paragraphs and generated page fragments.
    // A cell, nested story or intervening nonparagraph block ends that flow.
    private static ComplexFieldFlow ReadComplexFieldFlow(XElement owner, OoxPdfDocxMarkupMode markupMode,
        CancellationToken token)
    {
        foreach (var cached in owner.Annotations<ComplexFieldFlow>())
            if (cached.Mode == markupMode) return cached;

        var fields = new List<DocxComplexFieldScanState>();
        var unclosed = new HashSet<XElement>();
        bool unsupported = false;
        foreach (var block in EnumerateRevisionScopedChildren(owner.Elements(), markupMode, token, ParagraphFlowBlocks))
        {
            if (block.Element.Name == WordprocessingNamespace + "p")
                ScanFieldChildren(block.Element);
            else
                FinishFlow();
        }
        FinishFlow();
        var result = new ComplexFieldFlow(markupMode, unclosed, unsupported);
        owner.AddAnnotation(result);
        return result;

        void FinishFlow()
        {
            foreach (var field in fields)
            {
                token.ThrowIfCancellationRequested();
                unclosed.Add(field.Begin);
                if ((!field.HasSeparate || !field.HasCachedResult) && ResolveFieldPlaceholder(field.Instruction.ToString()) is null)
                    unsupported = true;
            }
            fields.Clear();
        }

        void ScanFieldChildren(XElement container)
        {
            foreach (XElement child in DocxSourceXml.StoredChildren(container))
            {
                token.ThrowIfCancellationRequested();
                if (child.Name == WordprocessingNamespace + "r")
                {
                    foreach (XElement part in child.Elements())
                    {
                        token.ThrowIfCancellationRequested();
                        unsupported |= ProcessComplexFieldRunChild(part, fields, markupMode);
                    }
                }
                else if (fields.Count != 0 && IsFieldResultContent(child))
                {
                    // Equations live beside runs and share the active result.
                    ProcessComplexFieldRunChild(child, fields, markupMode);
                }
                else if (IsVisibleRunContainer(child, markupMode) && !HasSimpleFieldPlaceholder(child))
                {
                    ScanFieldChildren(child);
                }
            }
        }
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

    private static bool ProcessComplexFieldRunChild(
        XElement child,
        List<DocxComplexFieldScanState> fields,
        OoxPdfDocxMarkupMode markupMode)
    {
        if (child.Name == WordprocessingNamespace + "fldChar")
        {
            string? fieldCharType = (string?)child.Attribute(WordprocessingNamespace + "fldCharType");
            if (string.Equals(fieldCharType, "begin", StringComparison.OrdinalIgnoreCase))
            {
                bool nestedInInstruction = fields.Count != 0 && !fields[^1].HasSeparate;
                bool deletedField = child.Ancestors(WordprocessingNamespace + "del").Any();
                fields.Add(new DocxComplexFieldScanState(child, new StringBuilder(), UsesOriginalFieldInstruction(markupMode, deletedField),
                    HasSeparate: false, HasCachedResult: false));
                return nestedInInstruction;
            }

            if (string.Equals(fieldCharType, "separate", StringComparison.OrdinalIgnoreCase))
            {
                if (fields.Count == 0)
                {
                    return true;
                }

                fields[^1] = fields[^1] with { HasSeparate = true };
                return false;
            }

            if (string.Equals(fieldCharType, "end", StringComparison.OrdinalIgnoreCase))
            {
                if (fields.Count == 0)
                {
                    return true;
                }

                var field = fields[^1];
                // A closed result may be intentionally empty. An instruction
                // without a result boundary is still unsupported.
                fields.RemoveAt(fields.Count - 1);
                return !field.HasSeparate && ResolveFieldPlaceholder(field.Instruction.ToString()) is null;
            }

            return true;
        }

        bool inFieldResult = fields.Count != 0 && fields[^1].HasSeparate;
        if (fields.Count != 0 && !inFieldResult && IsFieldInstructionText(child))
        {
            var field = fields[^1];
            bool inserted = field.UsesOriginalInstruction && child.Ancestors().Any(element =>
                element.Name == WordprocessingNamespace + "ins" || element.Name == WordprocessingNamespace + "moveTo");
            if (IncludesFieldInstruction(child, field.UsesOriginalInstruction, inserted))
                field.Instruction.Append((string?)child);
            return false;
        }

        if (fields.Count != 0 && IsFieldResultContent(child, inFieldResult))
        {
            for (int fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                var field = fields[fieldIndex];
                if (field.HasSeparate && !field.HasCachedResult)
                    fields[fieldIndex] = field with { HasCachedResult = true };
            }
        }

        return false;
    }
}
