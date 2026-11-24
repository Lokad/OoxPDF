using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static bool IsDisplayMathRun(XElement element) =>
        element.Name == MathNamespace + "r" && element.Parent is { } owner &&
        (owner.Name == MathNamespace + "oMathPara" ||
         DocxSourceXml.IsRevisionContainer(owner) && owner.Parent?.Name == MathNamespace + "oMathPara");

    private static bool IsEquationPayload(XElement element) =>
        element.Name == MathNamespace + "oMath" ||
        IsDisplayMathRun(element) && !IsDisplayMathLineBreakRun(element, allowCompanions: true);

    private static XElement? FindFirstTrailingDisplayMathBreak(XElement equation, CancellationToken cancellationToken)
    {
        if (equation.LastNode is XElement last &&
            !IsDisplayMathLineBreakRun(last) && !IsBoundary(last)) return null;

        XElement? first = null;
        foreach (XElement child in equation.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsDisplayMathLineBreakRun(child)) first ??= child;
            else if (!IsBoundary(child)) first = null;
        }
        return first;

        static bool IsBoundary(XElement element) =>
            DocxSourceXml.IsNonPrintingMarker(element) ||
            DocxMathParser.IsBookmarkMarker(element) ||
            DocxMathParser.IsCommentBoundary(element) ||
            (element.Name == MathNamespace + "r" &&
             element.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
             element.Elements().All(DocxSourceXml.IsLastRenderedPageBreakHint) &&
             element.Nodes().OfType<XText>().All(text => string.IsNullOrWhiteSpace(text.Value)));
    }

    private static bool IsDisplayMathLineBreakRun(XElement run, bool allowCompanions = false)
    {
        // Direct display runs can carry ordinary paragraph content alongside
        // a break. A run inside the equation can lift only its pure break.
        return run.Name == MathNamespace + "r" &&
            run.Element(WordprocessingNamespace + "br") is not null &&
            run.Elements().All(child => child.Name == WordprocessingNamespace + "rPr" ||
                DocxSourceXml.IsLastRenderedPageBreakHint(child) || IsLineBreak(child) ||
                allowCompanions && (child.Name == WordprocessingNamespace + "drawing" ||
                    child.Name == WordprocessingNamespace + "t" && !child.HasElements)) &&
            run.Nodes().OfType<XText>().All(text => string.IsNullOrWhiteSpace(text.Value));

        static bool IsLineBreak(XElement element) =>
            element.Name == WordprocessingNamespace + "br" &&
            !element.HasElements && string.IsNullOrWhiteSpace(element.Value) &&
            element.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
                attribute.Name == WordprocessingNamespace + "type" && attribute.Value == "textWrapping" ||
                attribute.Name == WordprocessingNamespace + "clear" && attribute.Value == "none");
    }

    private static IEnumerable<(XElement Element, DocxRevisionInfo? Revision, bool RemoveDeletedText)> EnumerateMathParagraphSources(
        XElement container, DocxRevisionInfo? inheritedRevision, bool removeDeletedText,
        OoxPdfDocxMarkupMode mode, CancellationToken token)
    {
        if (container.Name == MathNamespace + "oMath")
        {
            token.ThrowIfCancellationRequested();
            yield return (container, inheritedRevision, removeDeletedText);
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
            if (source.Element.Name == MathNamespace + "oMath" || IsDisplayMathRun(source.Element) ||
                source.Element.Name == WordprocessingNamespace + "r" ||
                source.Element.Name == WordprocessingNamespace + "bookmarkStart" ||
                source.Element.Name == WordprocessingNamespace + "bookmarkEnd" ||
                source.Element.Name == WordprocessingNamespace + "commentRangeStart" ||
                source.Element.Name == WordprocessingNamespace + "commentRangeEnd")
            {
                yield return (source.Element, source.Revision, source.RemoveDeletedText);
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
