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
    private static IReadOnlyList<DocxBodyElement> ReadBodyElements(
        XDocument document,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        XDocument? settings,
        DocxDocumentSettings documentSettings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        RelatedStoryPartCache storyCache)
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

                // Office sizes the terminal mark like unsized runs (edge-mark14c: Normal
                // 14pt wins over hardcoded 11; style-less docs fall back to 12pt).
                double markFontSize = ResolveRunProperties(null, null, null, styles, null).FontSize
                    ?? DocxDefaults.UnstyledRunFontSizePoints;
                elements.Add(new DocxImplicitParagraphElement(DocxBreakSourceKind.TerminalTable)
                {
                    MarkFontSizePoints = markFontSize
                });
            }

        var elements = new List<DocxBodyElement>();
        var numberingCounters = new DocxNumberingCounters();
        var inlineReferenceCounters = new DocxNoteReferenceCounters();
        DocxDocumentSettings authoredSettings = documentSettings;
        documentSettings = ResolveSingleSectionNoteReferenceSettings(document, documentSettings);
        IReadOnlyDictionary<XElement, DocxNoteReferenceSection> noteSections = ResolveSectionNoteReferenceSettings(document, authoredSettings, documentSettings);
        IEnumerable<XElement> bodyChildren = document.Descendants(WordprocessingNamespace + "body").Elements();
        foreach (var (scopedElement, tableFragments) in EnumerateLogicalTableBlocks(bodyChildren, styles, markupMode, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = scopedElement.Element;
            if (noteSections.TryGetValue(element, out DocxNoteReferenceSection? noteSection))
            {
                documentSettings = noteSection.Settings;
                inlineReferenceCounters.EnterSection(noteSection);
            }
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
                        elements.Add(ReadSectionBreak(breakParagraphSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, storyCache, inheritedRevision));
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
                        elements.Add(ReadSectionBreak(breakParagraphSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, storyCache, inheritedRevision));
                    }

                    continue;
                }

                if (element.Elements().Any(child => HasVisibleRunBreak(child, IsPageOrColumnBreak, markupMode)))
                {
                    foreach (ParagraphBreakPart part in SplitParagraphAtRunBreaks(element, markupMode, retainEmptyColumnTail: true))
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
                        elements.Add(ReadSectionBreak(splitSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, storyCache, inheritedRevision));
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
                    elements.Add(ReadSectionBreak(sectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, storyCache, inheritedRevision));
                }
            }
            else if (element.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? table = ReadTable(scopedElement, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters, documentSettings, markupMode, cancellationToken, tableFragments);
                if (table is not null)
                {
                    elements.Add(DocxBodyElementFactory.CreateTable(table));
                }
            }
        }

        elements = NormalizeDeletedParagraphMarkElements(elements, markupMode, cancellationToken).ToList();
        AddImplicitTerminalTableParagraph(elements);
        return elements;
    }

    private static IReadOnlyList<DocxBodyElement> NormalizeDeletedParagraphMarkElements(
        IReadOnlyList<DocxBodyElement> elements,
        OoxPdfDocxMarkupMode markupMode, CancellationToken cancellationToken)
    {
        if (DocxMarkupContext.FromMode(markupMode).IncludesDeletions || elements.Count < 2)
        {
            return elements;
        }

        var output = new List<DocxBodyElement>(elements.Count);
        foreach (DocxBodyElement element in elements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (element is DocxParagraphElement paragraphElement &&
                output.Count != 0 &&
                output[^1] is DocxParagraphElement previousParagraphElement &&
                previousParagraphElement.Paragraph.HasDeletedParagraphMark)
            {
                output[^1] = DocxBodyElementFactory.CreateParagraph(MergeParagraphsAcrossDeletedMark(
                    previousParagraphElement.Paragraph,
                    paragraphElement.Paragraph, cancellationToken));
                continue;
            }

            output.Add(element);
        }

        return output;
    }

    private static DocxParagraph MergeParagraphsAcrossDeletedMark(DocxParagraph first, DocxParagraph second,
        CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        int sourceRunOffset = MaxSourceRunIndex(first) + 1;
        int textRunOffset = first.Runs.Count;
        int textOffset = first.Runs.Sum(run => run.Text.Length);
        var runs = new Dictionary<DocxTextRun, DocxTextRun>(ReferenceEqualityComparer.Instance);
        var images = new Dictionary<DocxInlineImage, DocxInlineImage>(ReferenceEqualityComparer.Instance);
        var equations = new Dictionary<DocxEquation, DocxEquation>(ReferenceEqualityComparer.Instance);
        var boxes = new Dictionary<DocxInlineTextBox, DocxInlineTextBox>(ReferenceEqualityComparer.Instance);

        // Rebind every reference to the same shifted source object. Lists,
        // measured items and generated note references must not diverge.
        DocxTextRun ShiftRun(DocxTextRun source)
        {
            token.ThrowIfCancellationRequested();
            if (!runs.TryGetValue(source, out var shifted))
                runs.Add(source, shifted = source with { SourceRunIndex = ShiftSourceRunIndex(source.SourceRunIndex, sourceRunOffset) });
            return shifted;
        }
        DocxInlineImage ShiftImage(DocxInlineImage source)
        {
            token.ThrowIfCancellationRequested();
            if (!images.TryGetValue(source, out var shifted))
                images.Add(source, shifted = source with {
                    SourceRunIndex = ShiftSourceRunIndex(source.SourceRunIndex, sourceRunOffset),
                    SourceRunStyle = source.SourceRunStyle is { } style ? ShiftRun(style) : null });
            return shifted;
        }
        DocxEquation ShiftEquation(DocxEquation source)
        {
            token.ThrowIfCancellationRequested();
            if (!equations.TryGetValue(source, out var shifted))
                equations.Add(source, shifted = source with {
                    SourceRunIndex = ShiftSourceRunIndex(source.SourceRunIndex, sourceRunOffset),
                    TextRunIndex = source.TextRunIndex + textRunOffset });
            return shifted;
        }
        DocxInlineTextBox ShiftBox(DocxInlineTextBox source)
        {
            token.ThrowIfCancellationRequested();
            if (!boxes.TryGetValue(source, out var shifted))
                boxes.Add(source, shifted = source with { SourceRunIndex = ShiftSourceRunIndex(source.SourceRunIndex, sourceRunOffset) });
            return shifted;
        }
        DocxInlineItem ShiftItem(DocxInlineItem source) => source switch {
            DocxInlineText text => text with { Run = ShiftRun(text.Run) },
            DocxInlineTab tab => tab with { Run = ShiftRun(tab.Run) },
            DocxInlineBreak lineBreak => lineBreak with { Run = ShiftRun(lineBreak.Run) },
            DocxInlinePicture picture => picture with { Image = ShiftImage(picture.Image) },
            DocxInlineEquation math => math with { Equation = ShiftEquation(math.Equation) },
            DocxInlineBox box => box with { Box = ShiftBox(box.Box) },
            _ => source
        };
        var firstItems = first.InlineContent;
        var secondItems = second.InlineContent;
        // Plain prefixes/suffixes need no permanent source view until an object
        // or measured annotation requires one. Never reconstruct missing math.
        if (firstItems is not null || secondItems is not null)
        {
            if (firstItems is null && first.Equations.Count == 0)
                firstItems = DocxInlineContent.FromPrefix(first.Runs, first.Images, token, first.InlineTextBoxes);
            if (secondItems is null && second.Equations.Count == 0)
                secondItems = DocxInlineContent.FromPrefix(second.Runs, second.Images, token, second.InlineTextBoxes);
        }
        bool hasItems = firstItems is not null && secondItems is not null;
        int itemOffset = firstItems?.Count ?? 0;
        DocxBookmarkAnchor BindTerminalBookmark(DocxParagraph paragraph, DocxBookmarkAnchor anchor, int? itemCount) =>
            anchor.InlineItemIndex is null && itemCount is not null && anchor.TextRunIndex == paragraph.Runs.Count &&
                paragraph.Equations.Count == 0 && paragraph.Images.Count == 0 && paragraph.InlineTextBoxes.Count == 0
                ? anchor with { InlineItemIndex = itemCount } : anchor;
        return first with
        {
            Runs = first.Runs.Concat(second.Runs.Select(ShiftRun)).ToArray(),
            Images = first.Images.Concat(second.Images.Select(ShiftImage)).ToArray(),
            Equations = first.Equations.Concat(second.Equations.Select(ShiftEquation)).ToArray(),
            InlineTextBoxes = first.InlineTextBoxes.Concat(second.InlineTextBoxes.Select(ShiftBox)).ToArray(),
            InlineContent = hasItems ? firstItems!.Concat(secondItems!.Select(ShiftItem)).ToArray() : null,
            SpacingAfterPoints = second.SpacingAfterPoints,
            InlineReferences = first.InlineReferences.Concat(second.InlineReferences.Select(reference => reference with {
                SourceRunIndex = ShiftSourceRunIndex(reference.SourceRunIndex, sourceRunOffset),
                DisplayRun = reference.DisplayRun is { } run ? ShiftRun(run) : null })).ToArray(),
            CommentRanges = first.CommentRanges.Concat(second.CommentRanges.Select(range => ShiftCommentRange(range, sourceRunOffset, textOffset, itemOffset))).ToArray(),
            RevisionRanges = first.RevisionRanges.Concat(second.RevisionRanges.Select(range => ShiftRevisionRange(range, sourceRunOffset, textOffset))).ToArray(),
            FieldReferences = first.FieldReferences.Concat(second.FieldReferences.Select(field => ShiftFieldReference(field, sourceRunOffset, textRunOffset))).ToArray(),
            Hyperlinks = first.Hyperlinks.Concat(second.Hyperlinks.Select(link => ShiftHyperlink(link, sourceRunOffset, textRunOffset, itemOffset))).ToArray(),
            BookmarkAnchors = first.BookmarkAnchors.Select(anchor => BindTerminalBookmark(first, anchor, hasItems ? itemOffset : null))
                .Concat(second.BookmarkAnchors.Select(anchor => ShiftBookmarkAnchor(
                    BindTerminalBookmark(second, anchor, hasItems ? secondItems!.Count : null), sourceRunOffset, textRunOffset, itemOffset))).ToArray(),
            Revisions = first.Revisions.Concat(second.Revisions).ToArray(),
            NonRenderingRunOwners = first.NonRenderingRunOwners.Concat(second.NonRenderingRunOwners.Select(owner =>
                owner with { SourceRunIndex = ShiftSourceRunIndex(owner.SourceRunIndex, sourceRunOffset) })).ToArray(),
            ConsumedParagraphMarks = first.ConsumedParagraphMarks
                .Concat(first.HasDeletedParagraphMark && first.ParagraphMarkRun is { } mark ? [mark] : [])
                .Concat(second.ConsumedParagraphMarks).ToArray(),
            ParagraphMarkRun = second.ParagraphMarkRun,
            HasDeletedParagraphMark = second.HasDeletedParagraphMark,
            DeletedText = first.DeletedText + second.DeletedText
        };

        int MaxSourceRunIndex(DocxParagraph paragraph)
        {
            token.ThrowIfCancellationRequested();
            return paragraph.Runs.Select(run => run.SourceRunIndex)
                .Concat(paragraph.Images.Select(image => image.SourceRunIndex))
                .Concat(paragraph.Equations.Select(equation => equation.SourceRunIndex))
                .Concat(paragraph.InlineTextBoxes.Select(box => box.SourceRunIndex))
                .Concat(paragraph.NonRenderingRunOwners.Select(owner => owner.SourceRunIndex))
                .Concat(paragraph.InlineReferences.Select(reference => reference.SourceRunIndex))
                .Concat(paragraph.CommentRanges.SelectMany(range => new[] { range.StartSourceRunIndex, range.EndSourceRunIndex, range.ReferenceSourceRunIndex }).OfType<int>())
                .Concat(paragraph.RevisionRanges.SelectMany(range => new[] { range.StartSourceRunIndex, range.EndSourceRunIndex }).OfType<int>())
                .Concat(paragraph.FieldReferences.Select(field => field.SourceRunIndex))
                .Concat(paragraph.Hyperlinks.Select(link => link.SourceRunStartIndex))
                .Concat(paragraph.BookmarkAnchors.Select(anchor => anchor.SourceRunIndex))
                .Where(index => index >= 0).DefaultIfEmpty(-1).Max();
        }
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

    private static DocxCommentRange ShiftCommentRange(DocxCommentRange range, int sourceRunOffset, int textOffset, int itemOffset)
    {
        return range with
        {
            StartSourceRunIndex = ShiftSourceRunIndex(range.StartSourceRunIndex, sourceRunOffset),
            StartTextOffset = ShiftTextOffset(range.StartTextOffset, textOffset),
            EndSourceRunIndex = ShiftSourceRunIndex(range.EndSourceRunIndex, sourceRunOffset),
            EndTextOffset = ShiftTextOffset(range.EndTextOffset, textOffset),
            ReferenceSourceRunIndex = ShiftSourceRunIndex(range.ReferenceSourceRunIndex, sourceRunOffset),
            ReferenceTextOffset = ShiftTextOffset(range.ReferenceTextOffset, textOffset),
            StartInlineItemIndex = ShiftSourceRunIndex(range.StartInlineItemIndex, itemOffset),
            EndInlineItemIndex = ShiftSourceRunIndex(range.EndInlineItemIndex, itemOffset)
        };
    }

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

    private static DocxFieldReference ShiftFieldReference(DocxFieldReference field, int sourceRunOffset, int textRunOffset)
    {
        return field with
        {
            SourceRunIndex = ShiftSourceRunIndex(field.SourceRunIndex, sourceRunOffset),
            TextRunIndex = field.TextRunIndex + textRunOffset
        };
    }

    private static DocxHyperlinkSpan ShiftHyperlink(DocxHyperlinkSpan hyperlink, int sourceRunOffset, int textRunOffset, int itemOffset)
    {
        return hyperlink with
        {
            SourceRunStartIndex = ShiftSourceRunIndex(hyperlink.SourceRunStartIndex, sourceRunOffset),
            TextRunStartIndex = hyperlink.TextRunStartIndex + textRunOffset,
            StartInlineItemIndex = ShiftSourceRunIndex(hyperlink.StartInlineItemIndex, itemOffset),
            EndInlineItemIndex = ShiftSourceRunIndex(hyperlink.EndInlineItemIndex, itemOffset)
        };
    }

    private static DocxBookmarkAnchor ShiftBookmarkAnchor(DocxBookmarkAnchor anchor, int sourceRunOffset, int textRunOffset, int itemOffset)
    {
        return anchor with
        {
            SourceRunIndex = ShiftSourceRunIndex(anchor.SourceRunIndex, sourceRunOffset),
            TextRunIndex = anchor.TextRunIndex + textRunOffset,
            InlineItemIndex = ShiftSourceRunIndex(anchor.InlineItemIndex, itemOffset)
        };
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

    private static IReadOnlyList<ParagraphBreakPart> SplitParagraphAtRunBreaks(XElement paragraph, OoxPdfDocxMarkupMode markupMode,
        bool retainEmptyColumnTail = false, Func<XElement, bool>? isBreak = null)
    {
        isBreak ??= IsPageOrColumnBreak;
        var parts = new List<ParagraphBreakPart>();
        var currentChildren = new List<XElement>();
        XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
        bool startsAfterBreak = false;

        void AddParagraphPart(bool endsBeforeBreak)
        {
            // Word retains the paragraph mark after a trailing body column break.
            // Its style, pitch and after-spacing belong to the continuation frame.
            bool emptyColumnTail = retainEmptyColumnTail && !endsBeforeBreak && startsAfterBreak &&
                string.Equals(parts.LastOrDefault()?.BreakValue, "column", StringComparison.OrdinalIgnoreCase);
            if (currentChildren.Count == 0 && !emptyColumnTail)
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

            if (TrySplitRunBreakContainer(child, currentChildren, AddParagraphPart, parts, ref startsAfterBreak, markupMode, isBreak))
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

                if (runChild.Name == WordprocessingNamespace + "br" && isBreak(runChild))
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

    private static bool TrySplitRunBreakContainer(
        XElement child,
        List<XElement> paragraphChildren,
        Action<bool> addParagraphPart,
        List<ParagraphBreakPart> parts,
        ref bool startsAfterBreak,
        OoxPdfDocxMarkupMode markupMode, Func<XElement, bool> isBreak)
    {
        if (!IsVisibleRunContainer(child, markupMode) || !HasVisibleRunBreak(child, isBreak, markupMode))
        {
            return false;
        }

        SplitRunBreakContainer(child, paragraphChildren, addParagraphPart, parts, ref startsAfterBreak, markupMode, isBreak);
        return true;
    }

    private static void SplitRunBreakContainer(
        XElement sourceContainer,
        List<XElement> ownerChildren,
        Action<bool> flushOwnerAndParagraph,
        List<ParagraphBreakPart> parts,
        ref bool startsAfterBreak,
        OoxPdfDocxMarkupMode markupMode, Func<XElement, bool> isBreak)
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
                if (IsVisibleRunContainer(containerChild, markupMode) && HasVisibleRunBreak(containerChild, isBreak, markupMode))
                {
                    SplitRunBreakContainer(containerChild, containerChildren, FlushThisContainerAndParagraph, parts, ref startsAfterBreak, markupMode, isBreak);
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

                if (runChild.Name == WordprocessingNamespace + "br" && isBreak(runChild))
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
        return IsInlineMetadataContainer(element) || element.Name == WordprocessingNamespace + "fldSimple" ||
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
