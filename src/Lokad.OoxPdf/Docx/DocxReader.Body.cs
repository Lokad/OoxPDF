using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    // Single caller; kept static: entry-point stage, not a local candidate.
    private static IReadOnlyList<DocxBodyElement> ReadBodyElements(
        XDocument document,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        XDocument? settings,
        DocxDocumentSettings documentSettings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
            bool IsRunPageBreakOnlyParagraph(XElement paragraph, OoxPdfDocxMarkupMode markupMode)
            {
                return IsRunBreakOnlyParagraph(paragraph, IsPageBreak, markupMode);
            }

            void AddImplicitTerminalTableParagraph(List<DocxBodyElement> elements)
            {
                if (elements.Count == 0 ||
                    elements[^1] is not DocxTableElement)
                {
                    return;
                }

                elements.Add(new DocxImplicitParagraphElement(DocxBreakSourceKind.TerminalTable));
            }

        var elements = new List<DocxBodyElement>();
        var numberingCounters = new Dictionary<(string NumId, int Level), int>();
        var inlineReferenceCounters = new Dictionary<DocxRelatedStoryKind, int>();
        IEnumerable<XElement> bodyChildren = document.Descendants(WordprocessingNamespace + "body").Elements();
        foreach (DocxRevisionScopedElement scopedElement in EnumerateRevisionScopedChildren(bodyChildren, markupMode, WordprocessingNamespace + "p", WordprocessingNamespace + "tbl"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = scopedElement.Element;
            DocxRevisionInfo? inheritedRevision = scopedElement.Revision;
            if (element.Name == WordprocessingNamespace + "p")
            {
                XElement? paragraphProperties = element.Element(WordprocessingNamespace + "pPr");
                DocxResolvedParagraphProperties resolvedParagraph = ResolveParagraphProperties(
                    paragraphProperties,
                    ReadParagraphStyleId(paragraphProperties),
                    styles,
                    null);
                if (resolvedParagraph.PageBreakBefore == true)
                {
                    elements.Add(DocxBodyElementFactory.CreatePageBreak(
                        DocxBreakSourceKind.PageBreakBefore,
                        resolvedParagraph.PageBreakBeforeValue,
                        null,
                        revisions: inheritedRevision is null ? [] : [inheritedRevision]));
                }

                if (IsRunPageBreakOnlyParagraph(element, markupMode))
                {
                    DocxParagraph? breakParagraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
                    elements.Add(DocxBodyElementFactory.CreatePageBreak(DocxBreakSourceKind.RunBreak, "page", breakParagraph, null));
                    XElement? breakParagraphSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (breakParagraphSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(breakParagraphSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, inheritedRevision));
                    }

                    continue;
                }

                if (IsRunColumnBreakOnlyParagraph(element, markupMode))
                {
                    DocxParagraph? breakParagraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
                    elements.Add(DocxBodyElementFactory.CreateManualBreak(DocxBreakSourceKind.RunBreak, "column", breakParagraph));
                    XElement? breakParagraphSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (breakParagraphSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(breakParagraphSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, inheritedRevision));
                    }

                    continue;
                }

                if (HasRunPageOrColumnBreak(element, markupMode))
                {
                    foreach (ParagraphBreakPart part in SplitParagraphAtRunBreaks(element, markupMode))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (part.BreakValue is not null)
                        {
                            if (string.Equals(part.BreakValue, "column", StringComparison.OrdinalIgnoreCase))
                            {
                                elements.Add(DocxBodyElementFactory.CreateManualBreak(DocxBreakSourceKind.RunBreak, "column", null));
                            }
                            else
                            {
                                elements.Add(DocxBodyElementFactory.CreatePageBreak(DocxBreakSourceKind.RunBreak, part.BreakValue, null, null));
                            }

                            continue;
                        }

                        if (part.Paragraph is null)
                        {
                            continue;
                        }

                        DocxParagraph? splitParagraph = ReadParagraph(part.Paragraph, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
                        if (splitParagraph is not null)
                        {
                            elements.Add(DocxBodyElementFactory.CreateParagraph(AdjustBreakParagraphFragment(splitParagraph, part)));
                        }
                    }

                    XElement? splitSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (splitSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(splitSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, inheritedRevision));
                    }

                    continue;
                }

                DocxParagraph? paragraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
                if (paragraph is not null)
                {
                    elements.Add(DocxBodyElementFactory.CreateParagraph(paragraph));
                }

                XElement? sectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                if (sectionProperties is not null)
                {
                    elements.Add(ReadSectionBreak(sectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, inheritedRevision));
                }
            }
            else if (element.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? table = ReadTable(element, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters, documentSettings, markupMode, cancellationToken, inheritedRevision);
                if (table is not null)
                {
                    elements.Add(DocxBodyElementFactory.CreateTable(table));
                }
            }
        }

        elements = NormalizeDeletedParagraphMarkElements(elements, markupMode).ToList();
        AddImplicitTerminalTableParagraph(elements);
        return elements;
    }

    private static IReadOnlyList<DocxBodyElement> NormalizeDeletedParagraphMarkElements(
        IReadOnlyList<DocxBodyElement> elements,
        OoxPdfDocxMarkupMode markupMode)
    {
        if (DocxMarkupContext.FromMode(markupMode).IncludesDeletions || elements.Count < 2)
        {
            return elements;
        }

        var output = new List<DocxBodyElement>(elements.Count);
        foreach (DocxBodyElement element in elements)
        {
            if (element is DocxParagraphElement paragraphElement &&
                output.Count != 0 &&
                output[^1] is DocxParagraphElement previousParagraphElement &&
                previousParagraphElement.Paragraph.HasDeletedParagraphMark)
            {
                output[^1] = DocxBodyElementFactory.CreateParagraph(MergeParagraphsAcrossDeletedMark(
                    previousParagraphElement.Paragraph,
                    paragraphElement.Paragraph));
                continue;
            }

            output.Add(element);
        }

        return output;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxParagraph MergeParagraphsAcrossDeletedMark(DocxParagraph first, DocxParagraph second)
    {
            int MaxSourceRunIndex(DocxParagraph paragraph)
            {
                return paragraph.Runs
                    .Select(run => run.SourceRunIndex)
                    .Concat(paragraph.InlineReferences.Select(reference => reference.SourceRunIndex))
                    .Concat(paragraph.CommentRanges.SelectMany(range => new[] { range.StartSourceRunIndex, range.EndSourceRunIndex, range.ReferenceSourceRunIndex }).OfType<int>())
                    .Concat(paragraph.RevisionRanges.SelectMany(range => new[] { range.StartSourceRunIndex, range.EndSourceRunIndex }).OfType<int>())
                    .Concat(paragraph.FieldReferences.Select(field => field.SourceRunIndex))
                    .Concat(paragraph.Hyperlinks.Select(link => link.SourceRunStartIndex))
                    .Concat(paragraph.BookmarkAnchors.Select(anchor => anchor.SourceRunIndex))
                    .Where(index => index >= 0)
                    .DefaultIfEmpty(-1)
                    .Max();
            }

        int sourceRunOffset = MaxSourceRunIndex(first) + 1;
        int textRunOffset = first.Runs.Count;
        int textOffset = first.Runs.Sum(run => run.Text.Length);
        return first with
        {
            Runs = first.Runs.Concat(second.Runs.Select(run => ShiftRun(run, sourceRunOffset))).ToArray(),
            Images = first.Images.Concat(second.Images).ToArray(),
            SpacingAfterPoints = second.SpacingAfterPoints,
            InlineReferences = first.InlineReferences.Concat(second.InlineReferences.Select(reference => ShiftInlineReference(reference, sourceRunOffset))).ToArray(),
            CommentRanges = first.CommentRanges.Concat(second.CommentRanges.Select(range => ShiftCommentRange(range, sourceRunOffset, textOffset))).ToArray(),
            RevisionRanges = first.RevisionRanges.Concat(second.RevisionRanges.Select(range => ShiftRevisionRange(range, sourceRunOffset, textOffset))).ToArray(),
            FieldReferences = first.FieldReferences.Concat(second.FieldReferences.Select(field => ShiftFieldReference(field, sourceRunOffset, textRunOffset))).ToArray(),
            Hyperlinks = first.Hyperlinks.Concat(second.Hyperlinks.Select(link => ShiftHyperlink(link, sourceRunOffset, textRunOffset))).ToArray(),
            BookmarkAnchors = first.BookmarkAnchors.Concat(second.BookmarkAnchors.Select(anchor => ShiftBookmarkAnchor(anchor, sourceRunOffset, textRunOffset, textOffset))).ToArray(),
            Revisions = first.Revisions.Concat(second.Revisions).ToArray(),
            HasDeletedParagraphMark = second.HasDeletedParagraphMark
        };
    }


    private static int ShiftSourceRunIndex(int sourceRunIndex, int offset)
    {
        return sourceRunIndex < 0 ? sourceRunIndex : sourceRunIndex + offset;
    }

    private static int? ShiftSourceRunIndex(int? sourceRunIndex, int offset)
    {
        return sourceRunIndex is null || sourceRunIndex < 0 ? sourceRunIndex : sourceRunIndex + offset;
    }

    private static int? ShiftTextOffset(int? textOffset, int offset)
    {
        return textOffset is null ? null : textOffset + offset;
    }

    // Single caller; kept static: index-shift family kept together.
    private static DocxTextRun ShiftRun(DocxTextRun run, int sourceRunOffset)
    {
        return run with { SourceRunIndex = ShiftSourceRunIndex(run.SourceRunIndex, sourceRunOffset) };
    }

    // Single caller; kept static: index-shift family kept together.
    private static DocxInlineReference ShiftInlineReference(DocxInlineReference reference, int sourceRunOffset)
    {
        return reference with { SourceRunIndex = ShiftSourceRunIndex(reference.SourceRunIndex, sourceRunOffset) };
    }

    // Single caller; kept static: index-shift family kept together.
    private static DocxCommentRange ShiftCommentRange(DocxCommentRange range, int sourceRunOffset, int textOffset)
    {
        return range with
        {
            StartSourceRunIndex = ShiftSourceRunIndex(range.StartSourceRunIndex, sourceRunOffset),
            StartTextOffset = ShiftTextOffset(range.StartTextOffset, textOffset),
            EndSourceRunIndex = ShiftSourceRunIndex(range.EndSourceRunIndex, sourceRunOffset),
            EndTextOffset = ShiftTextOffset(range.EndTextOffset, textOffset),
            ReferenceSourceRunIndex = ShiftSourceRunIndex(range.ReferenceSourceRunIndex, sourceRunOffset),
            ReferenceTextOffset = ShiftTextOffset(range.ReferenceTextOffset, textOffset)
        };
    }

    // Single caller; kept static: index-shift family kept together.
    private static DocxRevisionRange ShiftRevisionRange(DocxRevisionRange range, int sourceRunOffset, int textOffset)
    {
        return range with
        {
            StartSourceRunIndex = ShiftSourceRunIndex(range.StartSourceRunIndex, sourceRunOffset),
            StartTextOffset = ShiftTextOffset(range.StartTextOffset, textOffset),
            EndSourceRunIndex = ShiftSourceRunIndex(range.EndSourceRunIndex, sourceRunOffset),
            EndTextOffset = ShiftTextOffset(range.EndTextOffset, textOffset)
        };
    }

    // Single caller; kept static: index-shift family kept together.
    private static DocxFieldReference ShiftFieldReference(DocxFieldReference field, int sourceRunOffset, int textRunOffset)
    {
        return field with
        {
            SourceRunIndex = ShiftSourceRunIndex(field.SourceRunIndex, sourceRunOffset),
            TextRunIndex = field.TextRunIndex + textRunOffset
        };
    }

    // Single caller; kept static: index-shift family kept together.
    private static DocxHyperlinkSpan ShiftHyperlink(DocxHyperlinkSpan hyperlink, int sourceRunOffset, int textRunOffset)
    {
        return hyperlink with
        {
            SourceRunStartIndex = ShiftSourceRunIndex(hyperlink.SourceRunStartIndex, sourceRunOffset),
            TextRunStartIndex = hyperlink.TextRunStartIndex + textRunOffset
        };
    }

    // Single caller; kept static: index-shift family kept together.
    private static DocxBookmarkAnchor ShiftBookmarkAnchor(DocxBookmarkAnchor anchor, int sourceRunOffset, int textRunOffset, int textOffset)
    {
        return anchor with
        {
            SourceRunIndex = ShiftSourceRunIndex(anchor.SourceRunIndex, sourceRunOffset),
            TextRunIndex = anchor.TextRunIndex + textRunOffset,
            TextOffset = anchor.TextOffset + textOffset
        };
    }

    private static bool HasRunPageOrColumnBreak(XElement paragraph, OoxPdfDocxMarkupMode markupMode)
    {
        return paragraph
            .Elements()
            .Any(element => HasVisibleRunPageOrColumnBreak(element, markupMode));
    }

    private static DocxParagraph AdjustBreakParagraphFragment(DocxParagraph paragraph, ParagraphBreakPart part)
    {
        return paragraph with
        {
            SpacingBeforePoints = part.StartsAfterBreak ? 0d : paragraph.SpacingBeforePoints,
            SpacingAfterPoints = part.EndsBeforeBreak ? 0d : paragraph.SpacingAfterPoints,
            ListLabel = part.StartsAfterBreak ? null : paragraph.ListLabel
        };
    }

    private static IReadOnlyList<ParagraphBreakPart> SplitParagraphAtRunBreaks(XElement paragraph, OoxPdfDocxMarkupMode markupMode)
    {
        var parts = new List<ParagraphBreakPart>();
        var currentChildren = new List<XElement>();
        XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
        bool startsAfterBreak = false;

        void AddParagraphPart(bool endsBeforeBreak)
        {
            if (currentChildren.Count == 0)
            {
                return;
            }

            var splitParagraph = new XElement(WordprocessingNamespace + "p");
            if (paragraphProperties is not null)
            {
                splitParagraph.Add(new XElement(paragraphProperties));
            }

            splitParagraph.Add(currentChildren.Select(child => new XElement(child)));
            parts.Add(new ParagraphBreakPart(splitParagraph, null, startsAfterBreak, endsBeforeBreak));
            currentChildren.Clear();
            startsAfterBreak = false;
        }

        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == WordprocessingNamespace + "pPr")
            {
                continue;
            }

            if (TrySplitRunBreakContainer(child, currentChildren, AddParagraphPart, parts, ref startsAfterBreak, markupMode))
            {
                continue;
            }

            if (child.Name != WordprocessingNamespace + "r")
            {
                currentChildren.Add(new XElement(child));
                continue;
            }

            XElement? runProperties = child.Element(WordprocessingNamespace + "rPr");
            var runChildren = new List<XElement>();
            if (runProperties is not null)
            {
                runChildren.Add(new XElement(runProperties));
            }

            foreach (XElement runChild in child.Elements())
            {
                if (runChild.Name == WordprocessingNamespace + "rPr")
                {
                    continue;
                }

                if (runChild.Name == WordprocessingNamespace + "br" && IsPageOrColumnBreak(runChild))
                {
                    AddRunPart(currentChildren, runProperties, runChildren);
                    AddParagraphPart(endsBeforeBreak: true);
                    parts.Add(new ParagraphBreakPart(null, (string?)runChild.Attribute(WordprocessingNamespace + "type"), false, false));
                    startsAfterBreak = true;
                    runChildren.Clear();
                    if (runProperties is not null)
                    {
                        runChildren.Add(new XElement(runProperties));
                    }

                    continue;
                }

                runChildren.Add(new XElement(runChild));
            }

            AddRunPart(currentChildren, runProperties, runChildren);
        }

        AddParagraphPart(endsBeforeBreak: false);
        return parts;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static bool TrySplitRunBreakContainer(
        XElement child,
        List<XElement> paragraphChildren,
        Action<bool> addParagraphPart,
        List<ParagraphBreakPart> parts,
        ref bool startsAfterBreak,
        OoxPdfDocxMarkupMode markupMode)
    {
        if (!IsVisibleRunContainer(child, markupMode) || !HasVisibleRunPageOrColumnBreak(child, markupMode))
        {
            return false;
        }

        SplitRunBreakContainer(child, paragraphChildren, addParagraphPart, parts, ref startsAfterBreak, markupMode);
        return true;
    }

    private static void SplitRunBreakContainer(
        XElement sourceContainer,
        List<XElement> ownerChildren,
        Action<bool> flushOwnerAndParagraph,
        List<ParagraphBreakPart> parts,
        ref bool startsAfterBreak,
        OoxPdfDocxMarkupMode markupMode)
    {
        var containerChildren = new List<XElement>();

        void FlushThisContainerAndParagraph(bool endsBeforeBreak)
        {
            AddContainerPart(ownerChildren, sourceContainer, containerChildren);
            flushOwnerAndParagraph(endsBeforeBreak);
        }

        foreach (XElement containerChild in sourceContainer.Elements())
        {
            if (containerChild.Name != WordprocessingNamespace + "r")
            {
                if (IsVisibleRunContainer(containerChild, markupMode) && HasVisibleRunPageOrColumnBreak(containerChild, markupMode))
                {
                    SplitRunBreakContainer(containerChild, containerChildren, FlushThisContainerAndParagraph, parts, ref startsAfterBreak, markupMode);
                }
                else
                {
                    containerChildren.Add(new XElement(containerChild));
                }

                continue;
            }

            XElement? runProperties = containerChild.Element(WordprocessingNamespace + "rPr");
            var runChildren = new List<XElement>();
            if (runProperties is not null)
            {
                runChildren.Add(new XElement(runProperties));
            }

            foreach (XElement runChild in containerChild.Elements())
            {
                if (runChild.Name == WordprocessingNamespace + "rPr")
                {
                    continue;
                }

                if (runChild.Name == WordprocessingNamespace + "br" && IsPageOrColumnBreak(runChild))
                {
                    AddRunPart(containerChildren, runProperties, runChildren);
                    FlushThisContainerAndParagraph(true);
                    parts.Add(new ParagraphBreakPart(null, (string?)runChild.Attribute(WordprocessingNamespace + "type"), false, false));
                    startsAfterBreak = true;
                    runChildren.Clear();
                    if (runProperties is not null)
                    {
                        runChildren.Add(new XElement(runProperties));
                    }

                    continue;
                }

                runChildren.Add(new XElement(runChild));
            }

            AddRunPart(containerChildren, runProperties, runChildren);
        }

        AddContainerPart(ownerChildren, sourceContainer, containerChildren);
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static void AddContainerPart(List<XElement> paragraphChildren, XElement sourceContainer, List<XElement> containerChildren)
    {
        if (containerChildren.Count == 0)
        {
            return;
        }

        var container = new XElement(sourceContainer.Name, sourceContainer.Attributes());
        container.Add(containerChildren.Select(child => new XElement(child)));
        paragraphChildren.Add(container);
        containerChildren.Clear();
    }

    private static void AddRunPart(List<XElement> paragraphChildren, XElement? runProperties, List<XElement> runChildren)
    {
        int contentOffset = runProperties is null ? 0 : 1;
        if (runChildren.Count <= contentOffset)
        {
            return;
        }

        paragraphChildren.Add(new XElement(WordprocessingNamespace + "r", runChildren.Select(child => new XElement(child))));
        runChildren.Clear();
        if (runProperties is not null)
        {
            runChildren.Add(new XElement(runProperties));
        }
    }


    private static bool IsRunColumnBreakOnlyParagraph(XElement paragraph, OoxPdfDocxMarkupMode markupMode)
    {
        return IsRunBreakOnlyParagraph(paragraph, IsColumnBreak, markupMode);
    }

    private static bool IsRunBreakOnlyParagraph(XElement paragraph, Func<XElement, bool> isBreak, OoxPdfDocxMarkupMode markupMode)
    {
        bool hasBreak = paragraph
            .Elements()
            .Any(element => HasVisibleRunBreak(element, isBreak, markupMode));
        if (!hasBreak)
        {
            return false;
        }

        return paragraph.Elements().All(element =>
            element.Name == WordprocessingNamespace + "pPr" ||
            IsBreakOnlyInlineElement(element, isBreak, markupMode));
    }

    private static bool HasVisibleRunBreak(XElement element, Func<XElement, bool> isBreak, OoxPdfDocxMarkupMode markupMode)
    {
        if (element.Name == WordprocessingNamespace + "r")
        {
            return element.Elements(WordprocessingNamespace + "br").Any(isBreak);
        }

        return IsVisibleRunContainer(element, markupMode) &&
            element.Elements().Any(child => HasVisibleRunBreak(child, isBreak, markupMode));
    }

    private static bool IsBreakOnlyInlineElement(XElement element, Func<XElement, bool> isBreak, OoxPdfDocxMarkupMode markupMode)
    {
        if (element.Name == WordprocessingNamespace + "r")
        {
            return element.Elements().All(child =>
                child.Name == WordprocessingNamespace + "rPr" ||
                child.Name == WordprocessingNamespace + "br" && isBreak(child));
        }

        if (IsIgnorableBreakOnlyContainerChild(element))
        {
            return true;
        }

        return IsVisibleRunContainer(element, markupMode) &&
            element.Elements().All(child => IsBreakOnlyInlineElement(child, isBreak, markupMode));
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static bool IsIgnorableBreakOnlyContainerChild(XElement element)
    {
        return element.Name == WordprocessingNamespace + "bookmarkStart" ||
            element.Name == WordprocessingNamespace + "bookmarkEnd" ||
            element.Name == WordprocessingNamespace + "commentRangeStart" ||
            element.Name == WordprocessingNamespace + "commentRangeEnd" ||
            element.Name == WordprocessingNamespace + "proofErr" ||
            element.Name == WordprocessingNamespace + "sdtPr" ||
            element.Name == WordprocessingNamespace + "sdtEndPr";
    }

    private static bool HasVisibleRunPageOrColumnBreak(XElement element, OoxPdfDocxMarkupMode markupMode)
    {
        if (element.Name == WordprocessingNamespace + "r")
        {
            return element.Elements(WordprocessingNamespace + "br").Any(IsPageOrColumnBreak);
        }

        return IsVisibleRunContainer(element, markupMode) &&
            element.Elements().Any(child => HasVisibleRunPageOrColumnBreak(child, markupMode));
    }

    private static bool IsPageOrColumnBreak(XElement breakElement)
    {
        return IsPageBreak(breakElement) || IsColumnBreak(breakElement);
    }

    private static bool IsVisibleRunContainer(XElement element)
    {
        return IsVisibleRunContainer(element, OoxPdfDocxMarkupMode.Final);
    }

    private static bool IsVisibleRunContainer(XElement element, OoxPdfDocxMarkupMode markupMode)
    {
        return element.Name == WordprocessingNamespace + "fldSimple" ||
            element.Name == WordprocessingNamespace + "hyperlink" ||
            element.Name == WordprocessingNamespace + "sdt" ||
            element.Name == WordprocessingNamespace + "sdtContent" ||
            IsIncludedRevisionContainer(element, markupMode);
    }

    private static bool IsPageBreak(XElement breakElement)
    {
        return string.Equals((string?)breakElement.Attribute(WordprocessingNamespace + "type"), "page", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsColumnBreak(XElement breakElement)
    {
        return string.Equals((string?)breakElement.Attribute(WordprocessingNamespace + "type"), "column", StringComparison.OrdinalIgnoreCase);
    }
}
