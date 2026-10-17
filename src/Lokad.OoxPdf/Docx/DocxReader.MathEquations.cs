using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static IEnumerable<(XElement Element, DocxRevisionInfo? Revision, bool RemoveDeletedText)> EnumerateMathEquations(
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
            if (source.Element.Name == MathNamespace + "oMath")
            {
                yield return (source.Element, source.Revision, source.RemoveDeletedText);
                continue;
            }
            // One actual revision owner is supported. Unknown/nested wrapper
            // composition retains the existing diagnosed source fallback.
            if (!IsRevisionContainer(source.Element) || source.HasWrapper) continue;
            var revision = source.Revision;
            bool remove = source.RemoveDeletedText;
            if (!IsIncludedRevisionContainer(source.Element, mode))
            {
                if (mode != OoxPdfDocxMarkupMode.Final || source.Element.Name != WordprocessingNamespace + "del") continue;
                revision = null; remove = true;
            }
            else revision = CreateRevisionInfo(source.Element) ?? revision;
            foreach (var child in source.Element.Elements().Reverse())
                pending.Push((child, revision, remove, true));
        }
    }
}
