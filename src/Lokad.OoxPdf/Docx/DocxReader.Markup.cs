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
        return properties
            ?.Elements()
            .Select(CreateRevisionInfo)
            .OfType<DocxRevisionInfo>()
            .ToArray() ?? [];
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
