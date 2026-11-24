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
        RelatedStoryPartCache storyCache,
        Dictionary<XElement, int>? floatingBlocks,
        bool projectCommentSelections)
    {
            void AddImplicitTerminalTableParagraph(List<DocxBodyElement> elements)
            {
                if (elements.Count == 0 ||
                    elements[^1] is not DocxTableElement)
                {
                    return;
                }

                // Office sizes the terminal mark like unsized runs (edge-mark14c: Normal
                // 14pt wins over hardcoded 11; style-less docs fall back to 12pt).
                double markFontSize = ResolveRunProperties(null, null, null, styles, null, out _).FontSize
                    ?? DocxDefaults.UnstyledRunFontSizePoints;
                elements.Add(new DocxImplicitParagraphElement(DocxBreakSourceKind.TerminalTable)
                {
                    MarkFontSizePoints = markFontSize
                });
            }

        var elements = new List<DocxBodyElement>();
        var comments = projectCommentSelections ? new DocxCommentProjection(elements) : null;
        bool mergeDeletedMarks = !DocxMarkupContext.FromMode(markupMode).IncludesDeletions;
        var numberingCounters = new DocxNumberingCounters();
        var inlineReferenceCounters = new DocxNoteReferenceCounters();
        DocxDocumentSettings authoredSettings = documentSettings;
        documentSettings = ResolveSingleSectionNoteReferenceSettings(document, documentSettings);
        IReadOnlyDictionary<XElement, DocxNoteReferenceSection> noteSections = ResolveSectionNoteReferenceSettings(
            document, authoredSettings, documentSettings, markupMode, cancellationToken);
        IEnumerable<XElement> bodyChildren = document.Descendants(WordprocessingNamespace + "body").Elements();
        foreach (var (scopedElement, tableFragments) in EnumerateLogicalTableBlocks(
            bodyChildren, styles, markupMode, cancellationToken, readCommentBoundaries: projectCommentSelections))
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
                var leadingBookmarks = scopedElement.LeadingBookmarks;
                bool HiddenBreak(XElement source) => IsHiddenHardRunBreak(source, styles, null);
                bool VisibleBreak(XElement source) => IsPageOrColumnBreak(source) && !HiddenBreak(source);
                bool VisiblePageBreak(XElement source) => IsPageBreak(source) && VisibleBreak(source);
                bool VisibleColumnBreak(XElement source) => IsColumnBreak(source) && VisibleBreak(source);
                XElement? paragraphProperties = element.Element(WordprocessingNamespace + "pPr");
                string? paragraphStyleId = ReadEffectiveParagraphStyleId(paragraphProperties, styles);
                DocxResolvedParagraphProperties resolvedParagraph = ResolveParagraphProperties(
                    ReadParagraphProperties(paragraphProperties),
                    paragraphStyleId,
                    styles,
                    null);
                bool hasVisibleBreak = element.Elements().Any(child => HasVisibleRunBreak(child, VisibleBreak, markupMode));

                // Numbering belongs to the paragraph, outside its run children.
                // A visible marker keeps the source first line even when its only
                // run payload is a break; ordinary paragraph layout paints it.
                var sourceNumbering = hasVisibleBreak ? ResolveParagraphNumbering(paragraphProperties, resolvedParagraph,
                    paragraphStyleId, styles, numbering, cancellationToken) : null;
                bool hasListMarker = sourceNumbering is { } source && !string.IsNullOrWhiteSpace(source.Definition.Text);
                bool isPageBreakOnly = hasVisibleBreak && !hasListMarker &&
                    IsRunBreakOnlyParagraph(element, VisiblePageBreak, HiddenBreak, markupMode, keepsBreakParagraph: true);
                bool isColumnBreakOnly = hasVisibleBreak && !hasListMarker && !isPageBreakOnly &&
                    IsRunBreakOnlyParagraph(element, VisibleColumnBreak, HiddenBreak, markupMode, keepsBreakParagraph: true);
                var parts = hasVisibleBreak && !isPageBreakOnly && !isColumnBreakOnly
                    ? SplitParagraphAtRunBreaks(element, markupMode, VisibleBreak, HiddenBreak, cancellationToken,
                        retainEmptyColumnTail: true, retainFirstLine: hasListMarker)
                    : null;

                // A first fragment can consume the preceding paragraph mark
                // before the authored break places the rest of the paragraph.
                bool canMergeLeadingFlow = !hasVisibleBreak || parts is [{ Paragraph: not null }, ..];
                var inlineCommentBoundaries = comments?.BeginParagraph(scopedElement.LeadingCommentBoundaries,
                    resolvedParagraph.PageBreakBefore != true && canMergeLeadingFlow &&
                    FindMergePredecessor(elements, mergeDeletedMarks) is not null,
                    cancellationToken);
                if (resolvedParagraph.PageBreakBefore == true)
                {
                    elements.Add(DocxBodyElementFactory.CreatePageBreak(
                        DocxBreakSourceKind.PageBreakBefore,
                        resolvedParagraph.PageBreakBeforeValue,
                        null,
                        revisions: inheritedRevision is null ? [] : [inheritedRevision]));
                }

                if (isPageBreakOnly)
                {
                    DocxParagraph? breakParagraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken,
                        leadingBookmarks: leadingBookmarks, retainUnmatchedCommentStarts: projectCommentSelections);
                    int breakCount = element.Elements().SelectMany(child => EnumerateVisibleRunBreaks(child, VisiblePageBreak, markupMode)).Count();
                    elements.Add(DocxBodyElementFactory.CreatePageBreak(DocxBreakSourceKind.RunBreak, "page", breakParagraph, null, breakCount));
                    comments?.TrackParagraph(breakParagraph, cancellationToken);
                    XElement? breakParagraphSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (breakParagraphSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(breakParagraphSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, storyCache, inheritedRevision));
                    }

                    continue;
                }

                if (isColumnBreakOnly)
                {
                    DocxParagraph? breakParagraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken,
                        leadingBookmarks: leadingBookmarks, retainUnmatchedCommentStarts: projectCommentSelections);
                    int breakCount = element.Elements().SelectMany(child => EnumerateVisibleRunBreaks(child, VisibleColumnBreak, markupMode)).Count();
                    elements.Add(DocxBodyElementFactory.CreateManualBreak(DocxBreakSourceKind.RunBreak, "column", breakParagraph, breakCount));
                    comments?.TrackParagraph(breakParagraph, cancellationToken);
                    XElement? breakParagraphSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (breakParagraphSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(breakParagraphSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, storyCache, inheritedRevision));
                    }

                    continue;
                }

                if (parts is not null)
                {
                    using IEnumerator<XElement>? sourceAnchors = floatingBlocks is null ? null :
                        EnumerateSplitFloatingAnchors(element).GetEnumerator();
                    bool matchesSourceAnchors = sourceAnchors is not null;
                    var bookmarksForFragment = leadingBookmarks;
                    var splitFields = new List<DocxComplexFieldState>();
                    var paragraphs = new DocxParagraph?[parts.Count];
                    ParagraphBreakPart? lastParagraph = parts.LastOrDefault(part => part.Paragraph is not null);
                    // A layout break does not end an authored field. Finish its
                    // instruction and result before publishing the fragments.
                    for (int i = 0; i < parts.Count; i++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        ParagraphBreakPart part = parts[i];
                        if (part.Paragraph is null) continue;
                        bool retainColumnMark = !part.EndsBeforeBreak && i > 0 &&
                            string.Equals(parts[i - 1].BreakValue, "column", StringComparison.OrdinalIgnoreCase);
                        paragraphs[i] = ReadParagraph(part.Paragraph, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken,
                            isContinuationFragment: !part.IsFirstFragment, leadingBookmarks: bookmarksForFragment,
                            splitFields: splitFields, continuesAfterBreak: !ReferenceEquals(part, lastParagraph),
                            retainEmptyColumnMark: retainColumnMark, retainUnmatchedCommentStarts: projectCommentSelections,
                            leadingInlineCommentBoundaries: i == 0 ? inlineCommentBoundaries : null);
                        bookmarksForFragment = null;
                    }
                    for (int i = 0; i < parts.Count; i++)
                    {
                        ParagraphBreakPart part = parts[i];
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

                        int? blockIndex = null;
                        if (paragraphs[i] is { } splitParagraph)
                        {
                            blockIndex = AppendFlowParagraph(elements,
                                AdjustBreakParagraphFragment(splitParagraph, part, sourceNumbering?.Definition.Indent),
                                mergeDeletedMarks, cancellationToken, comments);
                        }
                        if (matchesSourceAnchors)
                        {
                            // The splitter clones the run payloads in source order.
                            // Match each cloned anchor to its authored XML element.
                            foreach (XElement clone in EnumerateSplitFloatingAnchors(part.Paragraph))
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                if (!sourceAnchors!.MoveNext() || !XNode.DeepEquals(sourceAnchors.Current, clone))
                                {
                                    matchesSourceAnchors = false;
                                    break;
                                }
                                if (blockIndex is { } index)
                                {
                                    floatingBlocks![sourceAnchors.Current] = index;
                                }
                            }
                        }
                    }

                    XElement? splitSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (splitSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(splitSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, storyCache, inheritedRevision));
                    }

                    continue;
                }

                DocxParagraph? paragraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken,
                    leadingBookmarks: leadingBookmarks, retainUnmatchedCommentStarts: projectCommentSelections,
                    leadingInlineCommentBoundaries: inlineCommentBoundaries);
                if (paragraph is not null)
                {
                    int blockIndex = AppendFlowParagraph(elements, paragraph, mergeDeletedMarks, cancellationToken, comments);
                    BindFloatingDrawingBlocks(element, null, floatingBlocks, blockIndex, cancellationToken);
                }

                XElement? sectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                if (sectionProperties is not null)
                {
                    elements.Add(ReadSectionBreak(sectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, storyCache, inheritedRevision));
                }
            }
            else
            {
                // A table or external import interrupts the projected paragraph
                // sequence. Its content needs its own selection ownership.
                if (element.Name == WordprocessingNamespace + "tbl")
                {
                    DocxTable? table = ReadTable(scopedElement, styles, numbering, numberingCounters, package, relationships,
                        inlineReferenceCounters, documentSettings, markupMode, cancellationToken, tableFragments, projectCommentSelections);
                    if (table is not null)
                    {
                        BindFloatingDrawingBlocks(element, tableFragments, floatingBlocks, elements.Count, cancellationToken);
                        elements.Add(DocxBodyElementFactory.CreateTable(table));
                    }
                }
                comments?.FinishGroup(cancellationToken);
            }
        }

        comments?.FinishGroup(cancellationToken);
        AddImplicitTerminalTableParagraph(elements);
        return elements;

        IEnumerable<XElement> EnumerateSplitFloatingAnchors(XElement paragraph)
        {
            foreach (XElement anchor in paragraph.Descendants(WordprocessingDrawingNamespace + "anchor"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (markupMode != OoxPdfDocxMarkupMode.Original || !IsInsideExcludedRevisionContainer(anchor, markupMode))
                    yield return anchor;
            }
        }
    }

    private static void BindFloatingDrawingBlocks(XElement source, IReadOnlyList<DocxRevisionScopedElement>? tableFragments,
        Dictionary<XElement, int>? floatingBlocks, int blockIndex, CancellationToken cancellationToken)
    {
        if (floatingBlocks is null) return;
        if (tableFragments is null)
        {
            Bind(source);
        }
        else
        {
            foreach (DocxRevisionScopedElement fragment in tableFragments)
            {
                Bind(fragment.Element);
            }
        }

        void Bind(XElement scope)
        {
            foreach (XElement anchor in scope.Descendants(WordprocessingDrawingNamespace + "anchor"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                floatingBlocks[anchor] = blockIndex;
            }
        }
    }

    private static DocxParagraphElement? FindMergePredecessor(List<DocxBodyElement> elements, bool mergeDeletedMarks) =>
        mergeDeletedMarks && elements.Count > 0 && elements[^1] is DocxParagraphElement { Paragraph.HasDeletedParagraphMark: true } previous
            ? previous : null;

    private static int AppendFlowParagraph(List<DocxBodyElement> elements, DocxParagraph paragraph,
        bool mergeDeletedMarks, CancellationToken cancellationToken, DocxCommentProjection? comments = null)
    {
        if (FindMergePredecessor(elements, mergeDeletedMarks) is { } previous)
        {
            elements[^1] = DocxBodyElementFactory.CreateParagraph(
                MergeParagraphsAcrossDeletedMark(previous.Paragraph, paragraph, cancellationToken));
        }
        else
        {
            elements.Add(DocxBodyElementFactory.CreateParagraph(paragraph));
        }
        comments?.TrackParagraph(paragraph, cancellationToken);
        return elements.Count - 1;
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
            StaticDrawingOwners = first.StaticDrawingOwners.Count == 0 ? second.StaticDrawingOwners :
                second.StaticDrawingOwners.Count == 0 ? first.StaticDrawingOwners : [.. first.StaticDrawingOwners, .. second.StaticDrawingOwners],
            HiddenDrawingEquationSourceCount = first.HiddenDrawingEquationSourceCount + second.HiddenDrawingEquationSourceCount,
            InlineContent = hasItems ? firstItems!.Concat(secondItems!.Select(ShiftItem)).ToArray() : null,
            SpacingAfterPoints = second.SpacingAfterPoints,
            InlineReferences = first.InlineReferences.Concat(second.InlineReferences.Select(reference => reference with {
                SourceRunIndex = ShiftSourceRunIndex(reference.SourceRunIndex, sourceRunOffset),
                DisplayRun = reference.DisplayRun is { } run ? ShiftRun(run) : null })).ToArray(),
            CommentRanges = first.CommentRanges.Concat(second.CommentRanges.Select(range => ShiftCommentRange(range, sourceRunOffset, itemOffset))).ToArray(),
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

    private static DocxCommentRange ShiftCommentRange(DocxCommentRange range, int sourceRunOffset, int itemOffset)
    {
        // Text offsets belong to their source run and survive renumbering.
        return range with
        {
            StartSourceRunIndex = ShiftSourceRunIndex(range.StartSourceRunIndex, sourceRunOffset),
            EndSourceRunIndex = ShiftSourceRunIndex(range.EndSourceRunIndex, sourceRunOffset),
            ReferenceSourceRunIndex = ShiftSourceRunIndex(range.ReferenceSourceRunIndex, sourceRunOffset),
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

    private static DocxParagraph AdjustBreakParagraphFragment(DocxParagraph paragraph, ParagraphBreakPart part,
        DocxNumberingIndent? sourceListIndent = null)
    {
        DocxParagraphIndent indent = paragraph.Indent;
        if (part.StartsAfterBreak)
        {
            // A source paragraph keeps its continuation insets across a page or
            // column. Suppressing the repeated label must not discard its level's
            // indentation, or apply the paragraph's first-line offset again.
            sourceListIndent ??= paragraph.ListLabel?.Indent;
            var left = paragraph.ResolveLeftIndent(sourceListIndent);
            var right = paragraph.ResolveRightIndent(sourceListIndent);
            indent = indent with
            {
                LeftPoints = left.Points,
                LeftValue = left.Value,
                RightPoints = right.Points,
                RightValue = right.Value,
                FirstLinePoints = null, FirstLineValue = null,
                HangingPoints = null, HangingValue = null
            };
        }
        return paragraph with
        {
            SpacingBeforePoints = part.StartsAfterBreak ? 0d : paragraph.SpacingBeforePoints,
            SpacingAfterPoints = part.EndsBeforeBreak ? 0d : paragraph.SpacingAfterPoints,
            ListLabel = part.StartsAfterBreak ? null : paragraph.ListLabel,
            Indent = indent
        };
    }

    private static IReadOnlyList<ParagraphBreakPart> SplitParagraphAtRunBreaks(XElement paragraph, OoxPdfDocxMarkupMode markupMode,
        Func<XElement, bool> isBreak, Func<XElement, bool> isHiddenBreak, CancellationToken cancellationToken,
        bool retainEmptyColumnTail = false, bool retainFirstLine = false,
        IReadOnlyList<XElement>? leadingCommentBoundaries = null)
    {
        var parts = new List<ParagraphBreakPart>();
        var currentChildren = new List<XElement>();
        var authoredStory = DocxSourceXml.FindStoryRoot(paragraph) is { } story ? new DocxSourceXml.StoryOrigin(story) : null;
        var unclosed = paragraph.Descendants(WordprocessingNamespace + "fldChar").Any()
            ? ReadUnclosedFieldBegins(authoredStory?.Story, markupMode, cancellationToken) : null;
        XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
        bool startsAfterBreak = false;
        bool hasParagraphPart = false;

        void AddParagraphPart(bool endsBeforeBreak)
        {
            // Word retains the paragraph mark after a trailing body column break.
            // Its style, pitch and after-spacing belong to the continuation frame.
            bool emptyColumnTail = retainEmptyColumnTail && !endsBeforeBreak && startsAfterBreak &&
                string.Equals(parts.LastOrDefault()?.BreakValue, "column", StringComparison.OrdinalIgnoreCase);
            bool firstMarkerLine = retainFirstLine && endsBeforeBreak && parts.Count == 0;
            if (currentChildren.Count == 0 && !emptyColumnTail && !firstMarkerLine)
            {
                return;
            }

            var splitParagraph = new XElement(WordprocessingNamespace + "p");
            if (authoredStory is not null) splitParagraph.AddAnnotation(authoredStory);
            if (paragraphProperties is not null)
            {
                splitParagraph.Add(new XElement(paragraphProperties));
            }

            // These children are detached copies. Attaching them retains source annotations.
            splitParagraph.Add(currentChildren);
            parts.Add(new ParagraphBreakPart(splitParagraph, null, startsAfterBreak, endsBeforeBreak, !hasParagraphPart));
            hasParagraphPart = true;
            currentChildren.Clear();
            startsAfterBreak = false;
        }

        IEnumerable<XElement> children = EnumerateBreakSplitChildren(paragraph, isHiddenBreak, markupMode);
        if (leadingCommentBoundaries is not null)
            children = leadingCommentBoundaries.Concat(children);
        foreach (XElement child in children)
        {
            if (child.Name == WordprocessingNamespace + "pPr")
            {
                continue;
            }

            if (IsVisibleRunContainer(child, markupMode) && HasVisibleRunBreak(child, isBreak, markupMode))
            {
                SplitRunBreakContainer(child, currentChildren, AddParagraphPart, parts, ref startsAfterBreak,
                    markupMode, isBreak, isHiddenBreak, unclosed);
                continue;
            }

            if (child.Name != WordprocessingNamespace + "r")
            {
                currentChildren.Add(CopyFieldAwareFragment(child, unclosed));
                continue;
            }

            SplitRunAtBreaks(child, currentChildren, AddParagraphPart, parts, ref startsAfterBreak, isBreak, isHiddenBreak, unclosed);
        }

        AddParagraphPart(endsBeforeBreak: false);
        AttachNonprintingMetadataToContinuation(parts, retainFirstLine, isHiddenBreak, markupMode);
        return parts;
    }

    private static void AttachNonprintingMetadataToContinuation(List<ParagraphBreakPart> parts, bool retainFirstLine,
        Func<XElement, bool> isHiddenBreak, OoxPdfDocxMarkupMode markupMode)
    {
        int continuationIndex = -1;
        bool removedFragment = false;
        for (int i = parts.Count - 1; i >= 0; i--)
        {
            ParagraphBreakPart part = parts[i];
            if (part.Paragraph is not { } paragraph) continue;
            var children = DocxSourceXml.StoredChildren(paragraph)
                .Where(child => child.Name != WordprocessingNamespace + "pPr");
            if (!children.Any()) continue;
            if (!ContainsOnlyNonprintingBreakMetadata(children, isHiddenBreak, markupMode))
            {
                continuationIndex = i;
                continue;
            }

            // Closed ranges and empty inline containers can keep their metadata on this
            // paragraph's next fragment. Numbered first lines keep their frame.
            if (!part.EndsBeforeBreak || continuationIndex < 0 || retainFirstLine && part.IsFirstFragment)
                continue;

            ParagraphBreakPart continuation = parts[continuationIndex];
            XElement target = continuation.Paragraph!;
            if (target.Element(WordprocessingNamespace + "pPr") is { } properties)
                properties.AddAfterSelf(children);
            else
                target.AddFirst(children);
            if (part.IsFirstFragment)
                parts[continuationIndex] = continuation with { IsFirstFragment = true };
            parts[i] = part with { Paragraph = null };
            removedFragment = true;
        }

        if (removedFragment)
            parts.RemoveAll(part => part.Paragraph is null && part.BreakValue is null);
    }

    private static bool ContainsOnlyNonprintingBreakMetadata(IEnumerable<XElement> children,
        Func<XElement, bool> isHiddenBreak, OoxPdfDocxMarkupMode markupMode)
    {
        HashSet<(DocxRevisionKind Kind, string Id)>? open = null;
        bool hasEmptyContainer = false;
        return Visit(children) && (open is { Count: 0 } || open is null && hasEmptyContainer);

        bool Visit(IEnumerable<XElement> content)
        {
            foreach (XElement child in content)
            {
                if (DocxSourceXml.IsTransparentInlineContainer(child))
                {
                    if (!Visit(DocxSourceXml.StoredChildren(child))) return false;
                    continue;
                }

                if (DocxSourceXml.IsNonPrintingMarker(child)) continue;
                if ((child.Name == WordprocessingNamespace + "fldSimple" || child.Name == WordprocessingNamespace + "hyperlink") &&
                    !HasBreakSplitPayload(child, isHiddenBreak, markupMode, keepEmptyFields: false))
                {
                    hasEmptyContainer = true;
                    continue;
                }
                if (!TrackNonPrintingMoveRangeBoundary(child, ref open))
                    return false;
            }

            return true;
        }
    }

    private static bool TrackNonPrintingMoveRangeBoundary(XElement child, ref HashSet<(DocxRevisionKind Kind, string Id)>? open)
    {
        if (!DocxSourceXml.TryResolveRevisionRangeMarker(child, out DocxRevisionKind? kind, out bool isStart) ||
            !DocxSourceXml.IsNonPrintingRevisionRangeBoundary(child))
            return false;

        var marker = (kind.Value, child.Attribute(WordprocessingNamespace + "id")!.Value);
        open ??= [];
        return isStart ? open.Add(marker) : open.Remove(marker);
    }

    private static void SplitRunBreakContainer(
        XElement sourceContainer,
        List<XElement> ownerChildren,
        Action<bool> flushOwnerAndParagraph,
        List<ParagraphBreakPart> parts,
        ref bool startsAfterBreak,
        OoxPdfDocxMarkupMode markupMode, Func<XElement, bool> isBreak, Func<XElement, bool> isHiddenBreak,
        IReadOnlySet<XElement>? unclosed)
    {
        var containerChildren = new List<XElement>();
        XElement? computedField = HasSimpleFieldPlaceholder(sourceContainer) ? sourceContainer : null;
        bool ownsCounter = computedField is not null;

        void AddContainerPart()
        {
            if (!ownsCounter && (containerChildren.Count == 0 ||
                computedField is not null && !containerChildren.Any(child => child.DescendantsAndSelf().Any(IsCachedFieldDrawing))))
            {
                containerChildren.Clear();
                return;
            }

            var fragment = new XElement(sourceContainer.Name, sourceContainer.Attributes());
            fragment.Add(containerChildren);
            if (computedField is not null)
            {
                // Continuations retain their cached drawings. The first frame owns the
                // author's one instruction, including when the cache starts with a break.
                fragment.AddAnnotation(new ComputedSimpleFieldFragment(computedField, ownsCounter));
                ownsCounter = false;
            }
            ownerChildren.Add(fragment);
            containerChildren.Clear();
        }

        void FlushThisContainerAndParagraph(bool endsBeforeBreak)
        {
            AddContainerPart();
            flushOwnerAndParagraph(endsBeforeBreak);
        }

        foreach (XElement containerChild in EnumerateBreakSplitChildren(sourceContainer, isHiddenBreak, markupMode))
        {
            if (containerChild.Name != WordprocessingNamespace + "r")
            {
                if (IsVisibleRunContainer(containerChild, markupMode) && HasVisibleRunBreak(containerChild, isBreak, markupMode))
                {
                    SplitRunBreakContainer(containerChild, containerChildren, FlushThisContainerAndParagraph, parts, ref startsAfterBreak,
                        markupMode, isBreak, isHiddenBreak, unclosed);
                }
                else
                {
                    containerChildren.Add(CopyFieldAwareFragment(containerChild, unclosed));
                }

                continue;
            }

            SplitRunAtBreaks(containerChild, containerChildren, FlushThisContainerAndParagraph, parts, ref startsAfterBreak,
                isBreak, isHiddenBreak, unclosed);
        }

        AddContainerPart();
    }

    private static IEnumerable<XElement> EnumerateBreakSplitChildren(XElement owner, Func<XElement, bool> isHiddenBreak,
        OoxPdfDocxMarkupMode markupMode) =>
        DocxSourceXml.StoredChildren(owner).Where(child => HasBreakSplitPayload(child, isHiddenBreak, markupMode));

    // Revisions need stored payload too, but their authored runs can own revision
    // state without visible text. Preserve fields during splitting so their
    // metadata can be assigned to a fragment after break positions are known.
    private static bool HasBreakSplitPayload(XElement child, Func<XElement, bool> isHiddenBreak,
        OoxPdfDocxMarkupMode markupMode, bool preserveRevisionRuns = false, bool keepEmptyFields = true)
    {
        if (child.Name == WordprocessingNamespace + "r")
        {
            return preserveRevisionRuns || EnumerateBreakSplitRunChildren(child, isHiddenBreak).Any();
        }

        bool isRevision = DocxSourceXml.IsRevisionContainer(child);
        // Original excludes unselected insertions and moves even from measurement.
        // Other modes can need excluded deletions for math and table sizing.
        if (isRevision && markupMode == OoxPdfDocxMarkupMode.Original && !IsIncludedRevisionContainer(child, markupMode))
            return false;

        bool isStoredField = !keepEmptyFields && child.Name == WordprocessingNamespace + "fldSimple";
        bool isHyperlink = child.Name == WordprocessingNamespace + "hyperlink";
        if (!keepEmptyFields && (isStoredField || isHyperlink) &&
            (HasSimpleFieldPlaceholder(child) || child.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))))
            return true;

        if (isRevision || isStoredField || isHyperlink || DocxSourceXml.IsTransparentInlineContainer(child))
        {
            return DocxSourceXml.StoredChildren(child)
                .Any(nested => HasBreakSplitPayload(nested, isHiddenBreak, markupMode, preserveRevisionRuns || isRevision, keepEmptyFields));
        }

        return !DocxSourceXml.IsNonPrintingMarker(child);
    }

    private static void SplitRunAtBreaks(XElement run, List<XElement> ownerChildren,
        Action<bool> flushOwnerAndParagraph, List<ParagraphBreakPart> parts, ref bool startsAfterBreak,
        Func<XElement, bool> isBreak, Func<XElement, bool> isHiddenBreak, IReadOnlySet<XElement>? unclosed)
    {
        XElement? properties = run.Element(WordprocessingNamespace + "rPr");
        var content = new List<XElement>();

        void AddRunPart()
        {
            if (content.Count != 0)
            {
                // Adding attached children clones them, keeping the source run intact.
                ownerChildren.Add(new XElement(run.Name, properties, unclosed is { Count: > 0 }
                    ? content.Select(child => CopyFieldAwareFragment(child, unclosed)) : content));
                content.Clear();
            }
        }

        foreach (XElement child in EnumerateBreakSplitRunChildren(run, isHiddenBreak))
        {
            if (child.Name != WordprocessingNamespace + "br" || !isBreak(child))
            {
                content.Add(child);
                continue;
            }

            AddRunPart();
            flushOwnerAndParagraph(true);
            parts.Add(new ParagraphBreakPart(null, (string?)child.Attribute(WordprocessingNamespace + "type"), false, false));
            startsAfterBreak = true;
        }
        AddRunPart();
    }

    // Cached positions and hidden hard breaks contribute no fragment payload.
    private static IEnumerable<XElement> EnumerateBreakSplitRunChildren(XElement run, Func<XElement, bool> isHiddenBreak) =>
        run.Elements().Where(child => child.Name != WordprocessingNamespace + "rPr" &&
            !DocxSourceXml.IsLastRenderedPageBreakHint(child) && !isHiddenBreak(child));

    private static bool IsHiddenHardRunBreak(XElement source, DocxStyleSet styles, DocxTableCellStyle? tableStyle)
    {
        if (source.Name != WordprocessingNamespace + "br" || !IsPageOrColumnBreak(source))
        {
            return false;
        }

        // Structural admission uses the same resolved visibility as literal run
        // layout, including inherited styles and explicit visible overrides.
        XElement run = source.Parent!;
        XElement? properties = run.Ancestors(WordprocessingNamespace + "p").FirstOrDefault()?
            .Element(WordprocessingNamespace + "pPr");
        return ResolveRunProperties(run.Element(WordprocessingNamespace + "rPr"),
            ReadEffectiveParagraphStyleId(properties, styles), ReadCharacterStyleId(run), styles, tableStyle?.Run, out _).Hidden == true;
    }

    private static bool IsRunBreakOnlyParagraph(XElement paragraph, Func<XElement, bool> isBreak, Func<XElement, bool> isHiddenBreak,
        OoxPdfDocxMarkupMode markupMode, bool keepsBreakParagraph = false)
    {
        bool hasBreak = paragraph
            .Elements()
            .Any(element => HasVisibleRunBreak(element, isBreak, markupMode));
        if (!hasBreak)
        {
            return false;
        }

        HashSet<(DocxRevisionKind Kind, string Id)>? openMoveRanges = null;
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name != WordprocessingNamespace + "pPr" &&
                !IsBreakOnlyInlineElement(child, isBreak, isHiddenBreak, markupMode, keepsBreakParagraph, ref openMoveRanges))
                return false;
        }
        return openMoveRanges is not { Count: > 0 };
    }

    private static bool HasVisibleRunBreak(XElement element, Func<XElement, bool> isBreak, OoxPdfDocxMarkupMode markupMode)
        => EnumerateVisibleRunBreaks(element, isBreak, markupMode).Any();

    private static IEnumerable<XElement> EnumerateVisibleRunBreaks(XElement element, Func<XElement, bool> isBreak, OoxPdfDocxMarkupMode markupMode)
    {
        if (element.Name == WordprocessingNamespace + "r")
        {
            return element.Elements(WordprocessingNamespace + "br").Where(isBreak);
        }

        return IsVisibleRunContainer(element, markupMode)
            ? DocxSourceXml.StoredChildren(element).SelectMany(child => EnumerateVisibleRunBreaks(child, isBreak, markupMode))
            : [];
    }

    private static bool IsBreakOnlyInlineElement(XElement element, Func<XElement, bool> isBreak, Func<XElement, bool> isHiddenBreak,
        OoxPdfDocxMarkupMode markupMode, bool keepsBreakParagraph, ref HashSet<(DocxRevisionKind Kind, string Id)>? openMoveRanges)
    {
        if (element.Name == WordprocessingNamespace + "r")
        {
            foreach (XElement child in EnumerateBreakSplitRunChildren(element, isHiddenBreak))
            {
                // A range crossing a visible break keeps its original fragments.
                if (child.Name != WordprocessingNamespace + "br" || !isBreak(child) || openMoveRanges is { Count: > 0 })
                    return false;
            }
            return true;
        }

        if (!HasBreakSplitPayload(element, isHiddenBreak, markupMode) ||
            element.Name == WordprocessingNamespace + "bookmarkStart" ||
            element.Name == WordprocessingNamespace + "bookmarkEnd" ||
            element.Name == WordprocessingNamespace + "commentRangeStart" ||
            element.Name == WordprocessingNamespace + "commentRangeEnd")
        {
            return true;
        }

        // Body breaks own a source paragraph. Cell spill breaks do not.
        if (keepsBreakParagraph && IsRevisionMarkerElement(element))
            return TrackNonPrintingMoveRangeBoundary(element, ref openMoveRanges);

        // An unlocked counter supplies its own text even with an empty stored result.
        if (HasSimpleFieldPlaceholder(element))
            return false;

        if (!IsVisibleRunContainer(element, markupMode))
            return false;

        foreach (XElement child in DocxSourceXml.StoredChildren(element))
        {
            if (!IsBreakOnlyInlineElement(child, isBreak, isHiddenBreak, markupMode, keepsBreakParagraph, ref openMoveRanges))
                return false;
        }
        return true;
    }

    private static bool IsPageOrColumnBreak(XElement breakElement)
    {
        return IsPageBreak(breakElement) || IsColumnBreak(breakElement);
    }

    private static bool IsVisibleRunContainer(XElement element, OoxPdfDocxMarkupMode markupMode)
    {
        return DocxSourceXml.IsTransparentInlineContainer(element) || element.Name == WordprocessingNamespace + "fldSimple" ||
            element.Name == WordprocessingNamespace + "hyperlink" ||
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
