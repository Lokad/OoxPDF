using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using System.Diagnostics.CodeAnalysis;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    private static bool IsIncludedRevisionContainer(XElement element, OoxPdfDocxMarkupMode markupMode)
    {
        DocxMarkupContext markupContext = DocxMarkupContext.FromMode(markupMode);
        if (element.Name == WordprocessingNamespace + "ins" ||
            element.Name == WordprocessingNamespace + "moveTo")
        {
            return element.Name == WordprocessingNamespace + "ins"
                ? markupContext.IncludesInsertions
                : markupContext.IncludesMoveTo;
        }

        if (element.Name == WordprocessingNamespace + "del" ||
            element.Name == WordprocessingNamespace + "moveFrom")
        {
            return element.Name == WordprocessingNamespace + "del"
                ? markupContext.IncludesDeletions
                : markupContext.IncludesMoveFrom;
        }

        return false;
    }

    private static bool IsRevisionContainer(XElement element)
    {
        return element.Name == WordprocessingNamespace + "ins" ||
            element.Name == WordprocessingNamespace + "del" ||
            element.Name == WordprocessingNamespace + "moveFrom" ||
            element.Name == WordprocessingNamespace + "moveTo";
    }

    private static IEnumerable<DocxRevisionScopedElement> EnumerateRevisionScopedChildren(
        IEnumerable<XElement> elements,
        OoxPdfDocxMarkupMode markupMode,
        params XName[] includedNames)
    {
        return EnumerateRevisionScopedChildren(elements, markupMode, inheritedRevision: null, includedNames);
    }

    private static IEnumerable<DocxRevisionScopedElement> EnumerateRevisionScopedChildren(
        IEnumerable<XElement> elements,
        OoxPdfDocxMarkupMode markupMode,
        DocxRevisionInfo? inheritedRevision,
        params XName[] includedNames)
    {
        foreach (XElement element in elements)
        {
            if (includedNames.Contains(element.Name))
            {
                yield return new DocxRevisionScopedElement(element, inheritedRevision);
                continue;
            }

            if (element.Name == WordprocessingNamespace + "sdt")
            {
                foreach (XElement content in element.Elements(WordprocessingNamespace + "sdtContent"))
                {
                    foreach (DocxRevisionScopedElement child in EnumerateRevisionScopedChildren(content.Elements(), markupMode, inheritedRevision, includedNames))
                    {
                        yield return child;
                    }
                }

                continue;
            }

            if (element.Name == WordprocessingNamespace + "sdtContent")
            {
                foreach (DocxRevisionScopedElement child in EnumerateRevisionScopedChildren(element.Elements(), markupMode, inheritedRevision, includedNames))
                {
                    yield return child;
                }

                continue;
            }

            if (!IsRevisionContainer(element) || !IsIncludedRevisionContainer(element, markupMode))
            {
                continue;
            }

            DocxRevisionInfo? revision = CreateRevisionInfo(element) ?? inheritedRevision;
            foreach (DocxRevisionScopedElement child in EnumerateRevisionScopedChildren(element.Elements(), markupMode, revision, includedNames))
            {
                yield return child;
            }
        }
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static bool IsRevisionMarkerElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "moveFromRangeStart" ||
            element.Name == WordprocessingNamespace + "moveFromRangeEnd" ||
            element.Name == WordprocessingNamespace + "moveToRangeStart" ||
            element.Name == WordprocessingNamespace + "moveToRangeEnd";
    }

    private static bool TryResolveRevisionRangeMarker(XElement element, [NotNullWhen(true)] out DocxRevisionKind? kind, out bool isStart)
    {
        if (element.Name == WordprocessingNamespace + "moveFromRangeStart")
        {
            kind = DocxRevisionKind.MoveFrom;
            isStart = true;
            return true;
        }

        if (element.Name == WordprocessingNamespace + "moveFromRangeEnd")
        {
            kind = DocxRevisionKind.MoveFrom;
            isStart = false;
            return true;
        }

        if (element.Name == WordprocessingNamespace + "moveToRangeStart")
        {
            kind = DocxRevisionKind.MoveTo;
            isStart = true;
            return true;
        }

        if (element.Name == WordprocessingNamespace + "moveToRangeEnd")
        {
            kind = DocxRevisionKind.MoveTo;
            isStart = false;
            return true;
        }

        kind = null;
        isStart = false;
        return false;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxRevisionKind? RevisionKind(XElement element)
    {
        if (element.Name == WordprocessingNamespace + "ins")
        {
            return DocxRevisionKind.Insertion;
        }

        if (element.Name == WordprocessingNamespace + "del")
        {
            return DocxRevisionKind.Deletion;
        }

        if (element.Name == WordprocessingNamespace + "moveFrom")
        {
            return DocxRevisionKind.MoveFrom;
        }

        if (element.Name == WordprocessingNamespace + "moveTo")
        {
            return DocxRevisionKind.MoveTo;
        }

        if (element.Name == WordprocessingNamespace + "rPrChange")
        {
            return DocxRevisionKind.RunPropertiesChange;
        }

        if (element.Name == WordprocessingNamespace + "pPrChange")
        {
            return DocxRevisionKind.ParagraphPropertiesChange;
        }

        if (element.Name == WordprocessingNamespace + "tblPrChange")
        {
            return DocxRevisionKind.TablePropertiesChange;
        }

        if (element.Name == WordprocessingNamespace + "trPrChange")
        {
            return DocxRevisionKind.TableRowPropertiesChange;
        }

        if (element.Name == WordprocessingNamespace + "tcPrChange")
        {
            return DocxRevisionKind.TableCellPropertiesChange;
        }

        if (element.Name == WordprocessingNamespace + "sectPrChange")
        {
            return DocxRevisionKind.SectionPropertiesChange;
        }

        if (element.Name == WordprocessingNamespace + "moveFromRangeStart")
        {
            return DocxRevisionKind.MoveFromRangeStart;
        }

        if (element.Name == WordprocessingNamespace + "moveFromRangeEnd")
        {
            return DocxRevisionKind.MoveFromRangeEnd;
        }

        if (element.Name == WordprocessingNamespace + "moveToRangeStart")
        {
            return DocxRevisionKind.MoveToRangeStart;
        }

        return element.Name == WordprocessingNamespace + "moveToRangeEnd" ? DocxRevisionKind.MoveToRangeEnd : null;
    }

    private static DocxRevisionInfo? CreateRevisionInfo(XElement element)
    {
        if (RevisionKind(element) is not { } kind)
        {
            return null;
        }

        return new DocxRevisionInfo(
                kind,
                (string?)element.Attribute(WordprocessingNamespace + "id"),
                (string?)element.Attribute(WordprocessingNamespace + "author"),
                (string?)element.Attribute(WordprocessingNamespace + "date"),
                element.Name.LocalName,
                RevisionPropertyChangeFamily(element),
                ReadRevisionPropertyElementNames(element));
    }

    private static IReadOnlyList<DocxRevisionInfo> ReadPropertyChangeRevisions(XElement? properties)
    {
        if (properties is null)
        {
            return [];
        }

        var revisions = new List<DocxRevisionInfo>();
        foreach (XElement change in properties.Elements())
        {
            DocxRevisionInfo? revision = CreateRevisionInfo(change);
            if (revision is null)
            {
                continue;
            }

            // Office A/B (header-void and body-void probes): a property change whose
            // recorded original properties are identical to the current properties is
            // semantically void (for example bold added to an already-bold run) and
            // sustains no Word balloon, so it is dropped at the reader boundary. Balloon
            // emission, the lane-fit scale trigger, and revision counts all key off the
            // retained model.
            if (IsSemanticallyVoidPropertyChange(properties, change, revision))
            {
                continue;
            }

            revisions.Add(revision);
        }

        return revisions;
    }

    // Single caller; kept static: semantic-void property-change comparison.
    private static bool IsSemanticallyVoidPropertyChange(XElement properties, XElement change, DocxRevisionInfo revision)
    {
        if (revision.PropertyChangeFamily is null || revision.PropertyElementNames.Count == 0)
        {
            return false;
        }

        XElement? originalProperties = null;
        foreach (XElement child in change.Elements())
        {
            if (child.Name.Namespace != WordprocessingNamespace)
            {
                continue;
            }

            if (originalProperties is not null)
            {
                return false;
            }

            originalProperties = child;
        }

        if (originalProperties is null)
        {
            return false;
        }

        return OoxPropertyElementsEqual(originalProperties, properties, excludeChangeMarkers: true);
    }

    // Single caller; kept static: structural OOXML property comparison ignoring
    // namespace-declaration spelling and insignificant whitespace, so a recorded
    // original that only re-serializes the current properties compares void.
    private static bool OoxPropertyElementsEqual(XElement original, XElement current, bool excludeChangeMarkers)
    {
        if (original.Name != current.Name)
        {
            return false;
        }

        List<XAttribute> originalAttributes = original.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToList();
        List<XAttribute> currentAttributes = current.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).ToList();
        if (originalAttributes.Count != currentAttributes.Count)
        {
            return false;
        }

        foreach (XAttribute attribute in originalAttributes)
        {
            XAttribute? match = current.Attribute(attribute.Name);
            if (match is null || !string.Equals(match.Value, attribute.Value, StringComparison.Ordinal))
            {
                return false;
            }
        }

        List<XNode> originalContent = SignificantOoxContent(original).ToList();
        List<XNode> currentContent = SignificantOoxContent(current)
            .Where(node => !excludeChangeMarkers || node is not XElement element || RevisionPropertyChangeFamily(element) is null)
            .ToList();
        if (originalContent.Count != currentContent.Count)
        {
            return false;
        }

        for (int index = 0; index < originalContent.Count; index++)
        {
            if (originalContent[index] is XElement originalChild && currentContent[index] is XElement currentChild)
            {
                if (!OoxPropertyElementsEqual(originalChild, currentChild, excludeChangeMarkers: false))
                {
                    return false;
                }

                continue;
            }

            if (originalContent[index] is XText originalText && currentContent[index] is XText currentText)
            {
                if (!string.Equals(originalText.Value, currentText.Value, StringComparison.Ordinal))
                {
                    return false;
                }

                continue;
            }

            return false;
        }

        return true;
    }

    // Single caller; kept static: significant OOXML content (elements plus
    // non-whitespace text; comments and processing instructions never decide voidness).
    private static IEnumerable<XNode> SignificantOoxContent(XElement element)
    {
        foreach (XNode node in element.Nodes())
        {
            if (node is XText text)
            {
                if (!string.IsNullOrWhiteSpace(text.Value))
                {
                    yield return node;
                }
            }
            else if (node is XElement)
            {
                yield return node;
            }
        }
    }

    private static DocxRevisionPropertyFamily? RevisionPropertyChangeFamily(XElement element)
    {
        if (element.Name == WordprocessingNamespace + "rPrChange")
        {
            return DocxRevisionPropertyFamily.Run;
        }

        if (element.Name == WordprocessingNamespace + "pPrChange")
        {
            return DocxRevisionPropertyFamily.Paragraph;
        }

        if (element.Name == WordprocessingNamespace + "tblPrChange")
        {
            return DocxRevisionPropertyFamily.Table;
        }

        if (element.Name == WordprocessingNamespace + "trPrChange")
        {
            return DocxRevisionPropertyFamily.Row;
        }

        if (element.Name == WordprocessingNamespace + "tcPrChange")
        {
            return DocxRevisionPropertyFamily.Cell;
        }

        return element.Name == WordprocessingNamespace + "sectPrChange" ? DocxRevisionPropertyFamily.Section : null;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IReadOnlyList<string> ReadRevisionPropertyElementNames(XElement element)
    {
        if (RevisionPropertyChangeFamily(element) is null)
        {
            return [];
        }

        return element
            .Elements()
            .Where(child => child.Name.Namespace == WordprocessingNamespace)
            .SelectMany(child => child.Elements().Any()
                ? child.Elements().Where(grandchild => grandchild.Name.Namespace == WordprocessingNamespace)
                : [child])
            .Select(child => child.Name.LocalName)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
    }

    private static void AddRevision(List<DocxRevisionInfo> revisions, DocxRevisionInfo? revision)
    {
        if (revision is not null)
        {
            revisions.Add(revision);
        }
    }

    private static void AddRevisions(List<DocxRevisionInfo> revisions, IReadOnlyList<DocxRevisionInfo> added)
    {
        if (added.Count != 0)
        {
            revisions.AddRange(added);
        }
    }

    private static IReadOnlyList<DocxRevisionInfo> RevisionList(DocxRevisionInfo? revision)
    {
        return revision is null ? [] : [revision];
    }

    private static IReadOnlyList<DocxRevisionInfo> MergeRevisionLists(DocxRevisionInfo? primary, IReadOnlyList<DocxRevisionInfo> secondary)
    {
        if (primary is null)
        {
            return secondary;
        }

        return secondary.Count == 0
            ? [primary]
            : [primary, .. secondary];
    }

    // Single caller; kept static: index-finder trio kept together.
    private static DocxRevisionInfo? FindInheritedRevision(XElement element, OoxPdfDocxMarkupMode markupMode)
    {
        XElement? revisionContainer = element.Ancestors().FirstOrDefault(IsRevisionContainer);
        return revisionContainer is not null && IsIncludedRevisionContainer(revisionContainer, markupMode)
            ? CreateRevisionInfo(revisionContainer)
            : null;
    }

    private static bool IsInsideExcludedRevisionContainer(XElement element, OoxPdfDocxMarkupMode markupMode)
    {
        return element
            .Ancestors()
            .Any(ancestor => IsRevisionContainer(ancestor) && !IsIncludedRevisionContainer(ancestor, markupMode));
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxResolvedRunProperties ApplyMarkupRevisionStyle(
        DocxResolvedRunProperties run,
        DocxRevisionInfo? revision,
        OoxPdfDocxMarkupMode markupMode)
    {
        DocxMarkupContext markupContext = DocxMarkupContext.FromMode(markupMode);
        if (!markupContext.AppliesInlineRevisionStyle || revision is null)
        {
            return run;
        }

        return revision.Kind switch
        {
            DocxRevisionKind.Insertion => run with { ColorHex = run.ColorHex ?? "0000FF", Underline = true, UnderlineValue = run.UnderlineValue ?? "single" },
            DocxRevisionKind.Deletion => run with { ColorHex = run.ColorHex ?? "C00000", Strike = true, StrikeValue = run.StrikeValue ?? "true" },
            DocxRevisionKind.MoveFrom => run with { ColorHex = run.ColorHex ?? "C00000", DoubleStrike = true, DoubleStrikeValue = run.DoubleStrikeValue ?? "true" },
            DocxRevisionKind.MoveTo => run with { ColorHex = run.ColorHex ?? "008000", Underline = true, UnderlineValue = run.UnderlineValue ?? "single" },
            _ => run
        };
    }

    private sealed record DocxRevisionScopedElement(XElement Element, DocxRevisionInfo? Revision);

    private sealed record DocxCommentRangeStart(string? Id, int SourceRunIndex, int TextOffset);

    private sealed record DocxRevisionRangeStart(
        DocxRevisionKind Kind,
        string? Id,
        string? Name,
        string? Author,
        string? Date,
        int SourceRunIndex,
        int TextOffset);
}
