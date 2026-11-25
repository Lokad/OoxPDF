using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static bool IsDisplayMathRun(XElement element) =>
        element.Name == MathNamespace + "r" && element.Parent is { } owner &&
        (owner.Name == MathNamespace + "oMathPara" ||
         DocxSourceXml.IsRevisionContainer(owner) && owner.Parent?.Name == MathNamespace + "oMathPara");

    private static int CountEquationPayloads(XElement element, CancellationToken token) =>
        element.Name == MathNamespace + "oMath" ? 1 :
        IsDisplayMathRun(element) ? ClassifyDisplayMathRun(element, token).EquationCount : 0;

    private static (XElement Run, XElement? FirstBreakInRun)? FindFirstTrailingDisplayMathBreak(
        XElement equation, bool allowsDeletedText, CancellationToken cancellationToken)
    {
        if (equation.Name != MathNamespace + "oMath") return null;
        if (equation.LastNode is XElement last && !IsDisplayMathLineBreakRun(last) && !IsBoundary(last))
            return ReadPackedBreak(last);

        (XElement Run, XElement? FirstBreakInRun)? first = null;
        foreach (XElement child in equation.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsDisplayMathLineBreakRun(child)) first ??= (child, null);
            else if (!IsBoundary(child)) first = ReadPackedBreak(child);
        }
        return first;

        // A trailing break can share the last operand's run. Keep the
        // readable prefix under its original formatting owner.
        (XElement, XElement?)? ReadPackedBreak(XElement run) =>
            run.Name == MathNamespace + "r" && run.Elements().FirstOrDefault(IsDisplayMathBreak) is { } split &&
            HasOnlyWhitespaceRunText(run) && split.ElementsAfterSelf().Prepend(split).All(IsBreakOrNonPrintingMathControl) &&
            DocxMathParser.IsReadableRun(run, allowsDeletedText, split) ? (run, split) : null;

        static bool IsBoundary(XElement element) =>
            DocxMathParser.IsNonPrintingSourceMarker(element) ||
            DocxMathParser.IsBookmarkMarker(element) ||
            DocxMathParser.IsCommentBoundary(element) ||
            (element.Name == MathNamespace + "r" &&
             element.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
             element.Elements().All(DocxSourceXml.IsNonPrintingMathRunControl) &&
             HasOnlyWhitespaceRunText(element));
    }

    private static (bool UsesParagraphLayout, int EquationCount) ClassifyDisplayMathRun(XElement run, CancellationToken token)
    {
        // Only drawings and wrapping breaks require a direct run to use
        // paragraph layout. Its contiguous math operands keep math layout.
        bool usesParagraphLayout = false;
        bool inMath = false;
        int equationCount = 0;
        foreach (var child in run.Elements())
        {
            token.ThrowIfCancellationRequested();
            if (child.Name == MathNamespace + "rPr" || child.Name == WordprocessingNamespace + "rPr" ||
                DocxSourceXml.IsNonPrintingMathRunControl(child))
                continue;
            if (child.Name == MathNamespace + "t" && !child.HasElements)
            {
                if (!inMath) equationCount++;
                inMath = true;
                continue;
            }
            inMath = false;
            if (child.Name == WordprocessingNamespace + "drawing" || IsDisplayMathBreak(child))
            {
                usesParagraphLayout = true;
            }
            else if (!(child.Name == WordprocessingNamespace + "t" && !child.HasElements) &&
                !DocxMathParser.TryReadSymbol(child, out _, out _))
            {
                return (false, 1);
            }
        }
        if (!usesParagraphLayout || !HasOnlyWhitespaceRunText(run)) return (false, 1);
        if (equationCount == 0)
            return run.Element(MathNamespace + "rPr") is null ? (true, 0) : (false, 1);
        if (run.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            run.Elements(MathNamespace + "rPr").Skip(1).Any() ||
            run.Elements(WordprocessingNamespace + "rPr").Skip(1).Any()) return (false, 1);
        return (true, equationCount);
    }

    private static IEnumerable<XElement> SplitDisplayMathRun(XElement source, CancellationToken token)
    {
        XElement? mathProperties = source.Element(MathNamespace + "rPr");
        XElement? wordProperties = source.Element(WordprocessingNamespace + "rPr");
        XElement? run = null;
        foreach (var child in source.Elements())
        {
            token.ThrowIfCancellationRequested();
            if (child == mathProperties || child == wordProperties || DocxSourceXml.IsNonPrintingMathRunControl(child))
                continue;
            bool math = child.Name == MathNamespace + "t";
            XName name = (math ? MathNamespace : WordprocessingNamespace) + "r";
            if (run?.Name != name)
            {
                if (run is not null) yield return WrapMathRun(run);
                run = new XElement(name, math ? mathProperties : null, wordProperties);
            }
            run.Add(child);
        }
        if (run is not null) yield return WrapMathRun(run);

        static XElement WrapMathRun(XElement run) => run.Name.Namespace == MathNamespace
            ? new XElement(MathNamespace + "oMath", run) : run;
    }

    private static XElement AsParagraphRun(XElement run, XElement? firstChild = null) => new(WordprocessingNamespace + "r",
        (firstChild is null ? run.Elements() : firstChild.ElementsAfterSelf().Prepend(firstChild))
            .Where(child => !DocxSourceXml.IsNonPrintingMathRunControl(child)));

    // Within an equation, only a pure trailing break can move to paragraph layout.
    private static bool IsDisplayMathLineBreakRun(XElement run) =>
        run.Name == MathNamespace + "r" &&
            run.Elements().Any(IsDisplayMathBreak) &&
            run.Elements().All(child => child.Name == WordprocessingNamespace + "rPr" ||
                IsBreakOrNonPrintingMathControl(child)) &&
            HasOnlyWhitespaceRunText(run);

    private static bool HasOnlyWhitespaceRunText(XElement run) =>
        run.Nodes().OfType<XText>().All(text => string.IsNullOrWhiteSpace(text.Value));

    private static bool IsBreakOrNonPrintingMathControl(XElement element) =>
        DocxSourceXml.IsNonPrintingMathRunControl(element) || IsDisplayMathBreak(element);

    // Keep the authored break attributes when passing the control to paragraph
    // layout. Only a recognized, empty source control can leave the math run.
    private static bool IsDisplayMathBreak(XElement element) =>
        DocxSourceXml.IsTextWrappingBreak(element) &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value) &&
        element.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            element.Name == WordprocessingNamespace + "br" &&
            (attribute.Name == WordprocessingNamespace + "type" && attribute.Value == "textWrapping" ||
             attribute.Name == WordprocessingNamespace + "clear" && attribute.Value is "none" or "left" or "right" or "all"));

    private static IEnumerable<(XElement Element, bool IsEquation, DocxRevisionInfo? Revision, bool RemoveDeletedText)> EnumerateMathParagraphSources(
        XElement container, DocxRevisionInfo? inheritedRevision, bool removeDeletedText,
        OoxPdfDocxMarkupMode mode, CancellationToken token)
    {
        if (container.Name == MathNamespace + "oMath")
        {
            token.ThrowIfCancellationRequested();
            yield return (container, true, inheritedRevision, removeDeletedText);
            yield break;
        }
        var pending = new Stack<(XElement Element, DocxRevisionInfo? Revision, bool RemoveDeletedText, bool HasWrapper)>();
        foreach (var child in container.Elements().Reverse())
            pending.Push((child, inheritedRevision, removeDeletedText, inheritedRevision is not null || removeDeletedText));
        while (pending.TryPop(out var source))
        {
            token.ThrowIfCancellationRequested();
            // Math and Wordprocessing runs keep their authored order
            // alongside the displayed equations and range boundaries.
            if (IsDisplayMathRun(source.Element))
            {
                var layout = ClassifyDisplayMathRun(source.Element, token);
                if (!layout.UsesParagraphLayout)
                    yield return (source.Element, true, source.Revision, source.RemoveDeletedText);
                else if (layout.EquationCount == 0)
                    yield return (AsParagraphRun(source.Element), false, source.Revision, source.RemoveDeletedText);
                else
                    foreach (var fragment in SplitDisplayMathRun(source.Element, token))
                        yield return (fragment, fragment.Name == MathNamespace + "oMath", source.Revision, source.RemoveDeletedText);
                continue;
            }
            if (source.Element.Name == MathNamespace + "oMath" ||
                source.Element.Name == WordprocessingNamespace + "r" ||
                source.Element.Name == WordprocessingNamespace + "bookmarkStart" ||
                source.Element.Name == WordprocessingNamespace + "bookmarkEnd" ||
                source.Element.Name == WordprocessingNamespace + "commentRangeStart" ||
                source.Element.Name == WordprocessingNamespace + "commentRangeEnd")
            {
                yield return (source.Element, source.Element.Name == MathNamespace + "oMath", source.Revision, source.RemoveDeletedText);
                continue;
            }
            // One actual revision owner is supported. Unknown/nested wrapper
            // composition retains the existing diagnosed source fallback.
            if (!DocxSourceXml.IsRevisionContainer(source.Element) || source.HasWrapper) continue;
            var revision = source.Revision;
            bool remove = source.RemoveDeletedText;
            if (!IsIncludedRevisionContainer(source.Element, mode))
            {
                if (!RetainsDeletedMathStructure(source.Element, mode)) continue;
                revision = null; remove = true;
            }
            else revision = CreateRevisionInfo(source.Element) ?? revision;
            foreach (var child in source.Element.Elements().Reverse())
                pending.Push((child, revision, remove, true));
        }
    }
}
