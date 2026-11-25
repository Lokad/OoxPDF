using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private readonly record struct DisplayMathRunLayout(
        bool UsesParagraphLayout, int EquationCount, bool StartsWithMath, bool EndsWithMath);

    private static readonly DisplayMathRunLayout UnsplitMathRun = new(false, 1, true, true);

    private static bool IsDisplayMathRun(XElement element) =>
        element.Name == MathNamespace + "r" && HasDisplayMathParagraphOwner(element);

    private static bool HasDisplayMathParagraphOwner(XElement element) =>
        element.Parent is { } owner &&
        (owner.Name == MathNamespace + "oMathPara" ||
         DocxSourceXml.IsRevisionContainer(owner) && owner.Parent?.Name == MathNamespace + "oMathPara");

    private static int CountEquationPayloads(XElement element, CancellationToken token) =>
        element.Name == MathNamespace + "oMath" ? CountDisplayDrawingRow(element, token) ?? 1 :
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
            IsMathRangeBoundary(element) ||
            (element.Name == MathNamespace + "r" &&
             element.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
             element.Elements().All(DocxSourceXml.IsNonPrintingMathRunControl) &&
             HasOnlyWhitespaceRunText(element));
    }

    private static bool IsMathRangeBoundary(XElement element) =>
        DocxMathParser.IsBookmarkMarker(element) || DocxMathParser.IsCommentBoundary(element);

    private static DisplayMathRunLayout ClassifyDisplayMathRun(XElement run, CancellationToken token, bool insideEquation = false)
    {
        // Only drawings and wrapping breaks require a direct run to use
        // paragraph layout. Its contiguous math operands keep math layout.
        bool usesParagraphLayout = false;
        bool inMath = false;
        bool visitedContent = false;
        bool startsWithMath = false;
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
                startsWithMath |= !visitedContent;
                visitedContent = true;
                continue;
            }
            visitedContent = true;
            inMath = false;
            if (child.Name == WordprocessingNamespace + "drawing")
            {
                // A row can decompose around an inline drawing. Anchors still
                // belong to the paragraph's floating-drawing source admission.
                if (insideEquation && (child.Element(WordprocessingDrawingNamespace + "inline") is null ||
                    child.Elements().Skip(1).Any())) return UnsplitMathRun;
                usesParagraphLayout = true;
            }
            else if (!insideEquation && IsDisplayMathBreak(child))
            {
                usesParagraphLayout = true;
            }
            else if (!IsDisplayMathWordText(child, run) &&
                !DocxMathParser.TryReadSymbol(child, out _, out _))
            {
                return UnsplitMathRun;
            }
        }
        if (!usesParagraphLayout || !HasOnlyWhitespaceRunText(run)) return UnsplitMathRun;
        var layout = new DisplayMathRunLayout(true, equationCount, startsWithMath, inMath);
        if (equationCount == 0 && !insideEquation)
            return run.Element(MathNamespace + "rPr") is null ? layout : UnsplitMathRun;
        if (run.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            run.Elements(MathNamespace + "rPr").Skip(1).Any() ||
            run.Elements(WordprocessingNamespace + "rPr").Skip(1).Any() ||
            equationCount == 0 && run.Element(MathNamespace + "rPr") is not null) return UnsplitMathRun;
        return layout;
    }

    private static bool IsDisplayMathWordText(XElement element, XElement run) =>
        !element.HasElements &&
        (element.Name == WordprocessingNamespace + "t" ||
         element.Name == WordprocessingNamespace + "delText" &&
         run.Ancestors().FirstOrDefault(DocxSourceXml.IsRevisionContainer)?.Name == WordprocessingNamespace + "del");

    // Keep consecutive mathematical atoms in one row on each side of an
    // inline drawing. The existing trailing-break path still owns math breaks.
    private static int? CountDisplayDrawingRow(XElement equation, CancellationToken token)
    {
        if (!HasDisplayMathParagraphOwner(equation) ||
            equation.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
            !HasOnlyWhitespaceRunText(equation)) return null;

        bool hasDrawing = false;
        bool inMath = false;
        int count = 0;
        foreach (var child in equation.Elements())
        {
            token.ThrowIfCancellationRequested();
            if (IsMathRangeBoundary(child) || DocxMathParser.IsNonPrintingSourceMarker(child)) continue;
            if (child.Name == MathNamespace + "r" && child.Element(WordprocessingNamespace + "drawing") is not null)
            {
                var layout = ClassifyDisplayMathRun(child, token, insideEquation: true);
                if (!layout.UsesParagraphLayout) return null;
                count += layout.EquationCount - (inMath && layout.StartsWithMath ? 1 : 0);
                inMath = layout.EndsWithMath;
                hasDrawing = true;
                continue;
            }
            if (child.Name != MathNamespace + "r" && !DocxMathParser.IsMathControl(child.Name) ||
                child.Name == MathNamespace + "r" && child.Elements().Any(DocxSourceXml.IsTextWrappingBreak)) return null;
            if (!inMath) count++;
            inMath = true;
        }
        return hasDrawing && count != 0 ? count : null;
    }

    private static IEnumerable<XElement> SplitDisplayDrawingRow(XElement equation, CancellationToken token)
    {
        XElement? group = null;
        List<XElement>? boundaries = null;
        foreach (var child in equation.Elements())
        {
            token.ThrowIfCancellationRequested();
            if (DocxMathParser.IsNonPrintingSourceMarker(child)) continue;
            if (IsMathRangeBoundary(child))
            {
                // A comment ending just after Wordprocessing content closes
                // at that inline boundary, before a later math group begins.
                if (group is null && child.Name == WordprocessingNamespace + "commentRangeEnd")
                {
                    foreach (var boundary in TakeBoundaries())
                        yield return boundary;
                    yield return child;
                }
                else
                    (boundaries ??= []).Add(child);
                continue;
            }
            if (child.Name != MathNamespace + "r" || child.Element(WordprocessingNamespace + "drawing") is null)
            {
                AddMath(child);
                continue;
            }
            foreach (var part in SplitDisplayMathRun(child, token))
            {
                if (part.Name.Namespace == MathNamespace)
                {
                    AddMath(part);
                    continue;
                }
                if (group is not null)
                {
                    yield return group;
                    group = null;
                }
                foreach (var boundary in TakeBoundaries())
                    yield return boundary;
                yield return part;
            }
        }
        if (group is not null)
        {
            group.Add(TakeBoundaries());
            yield return group;
        }
        else
        {
            foreach (var boundary in TakeBoundaries())
                yield return boundary;
        }

        void AddMath(XElement element)
        {
            group ??= new XElement(equation.Name, equation.Attributes());
            group.Add(TakeBoundaries());
            group.Add(element);
        }

        // Boundaries between atoms retain their position in the group.
        // Boundaries adjacent to Wordprocessing content use paragraph order.
        IEnumerable<XElement> TakeBoundaries()
        {
            if (boundaries is null) yield break;
            foreach (var boundary in boundaries)
            {
                token.ThrowIfCancellationRequested();
                yield return boundary;
            }
            boundaries.Clear();
        }
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
                if (run is not null) yield return run;
                run = new XElement(name, math ? mathProperties : null, wordProperties);
            }
            run.Add(child);
        }
        if (run is not null) yield return run;
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
                    {
                        bool math = fragment.Name.Namespace == MathNamespace;
                        yield return (math ? new XElement(MathNamespace + "oMath", fragment) : fragment,
                            math, source.Revision, source.RemoveDeletedText);
                    }
                continue;
            }
            if (source.Element.Name == MathNamespace + "oMath" && CountDisplayDrawingRow(source.Element, token) is not null)
            {
                foreach (var fragment in SplitDisplayDrawingRow(source.Element, token))
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
