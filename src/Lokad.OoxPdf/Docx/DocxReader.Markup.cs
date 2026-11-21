using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxReader
{
    // Expected visible payloads drive omission diagnostics, not a raw XML census.
    // Empty display payloads still diagnose; visibility shares the reader's
    // existing revision and row decisions.
    private static int CountSelectedEquationSources(XContainer source, OoxPdfDocxMarkupMode mode,
        CancellationToken token, out bool hasSource)
    {
        token.ThrowIfCancellationRequested();
        int count = 0;
        hasSource = false;
        foreach (XElement element in source.Descendants())
        {
            token.ThrowIfCancellationRequested();
            bool equation = element.Name == MathNamespace + "oMath";
            if (!equation && element.Name != MathNamespace + "oMathPara") continue;
            hasSource = true;
            if (!equation && element.Descendants(MathNamespace + "oMath").Any()) continue;
            bool included = true;
            bool retainsDeletedStructure = true;
            foreach (XElement ancestor in element.Ancestors())
            {
                token.ThrowIfCancellationRequested();
                if (DocxSourceXml.IsRevisionContainer(ancestor) && !IsIncludedRevisionContainer(ancestor, mode) &&
                        !(retainsDeletedStructure && RetainsDeletedMathStructure(ancestor, mode)) ||
                    ancestor.Name == WordprocessingNamespace + "tr" && !IsIncludedTableRow(ancestor, mode, token))
                { included = false; break; }
                // Deleted math may cross its metadata wrappers, not a paragraph,
                // drawing run or a second revision boundary. Those owners are
                // omitted before paragraph-level math reading occurs.
                retainsDeletedStructure &= ancestor.Name == MathNamespace + "oMathPara" ||
                    DocxSourceXml.IsTransparentInlineContainer(ancestor);
            }
            if (included) count++;
        }
        return count;
    }

    private static bool RetainsDeletedMathStructure(XElement container, OoxPdfDocxMarkupMode mode) =>
        mode == OoxPdfDocxMarkupMode.Final && container.Name == WordprocessingNamespace + "del";

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

    private static IEnumerable<DocxRevisionScopedElement> EnumerateRevisionScopedChildren(
        IEnumerable<XElement> elements,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        XName[] includedNames,
        DocxRevisionInfo? inheritedRevision = null,
        bool inlineContent = false)
    {
        cancellationToken.ThrowIfCancellationRequested();
        foreach (XElement element in elements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (includedNames.Contains(element.Name))
            {
                yield return new DocxRevisionScopedElement(element, inheritedRevision);
                continue;
            }

            bool transparent = DocxSourceXml.IsTransparentInlineContainer(element);
            // The cached run projection keeps drawing payloads inside hyperlinks.
            // The paragraph reader owns hyperlink spans for the computed result.
            if (inlineContent
                ? transparent || element.Name == WordprocessingNamespace + "hyperlink"
                : transparent && element.Name != WordprocessingNamespace + "smartTag")
            {
                foreach (DocxRevisionScopedElement child in EnumerateRevisionScopedChildren(
                    DocxSourceXml.StoredChildren(element), markupMode, cancellationToken, includedNames, inheritedRevision, inlineContent))
                {
                    yield return child;
                }

                continue;
            }

            if (!DocxSourceXml.IsRevisionContainer(element) || !IsIncludedRevisionContainer(element, markupMode))
            {
                continue;
            }

            DocxRevisionInfo? revision = CreateRevisionInfo(element) ?? inheritedRevision;
            foreach (DocxRevisionScopedElement child in EnumerateRevisionScopedChildren(
                element.Elements(), markupMode, cancellationToken, includedNames, revision, inlineContent))
            {
                yield return child;
            }
        }
    }

    private static bool IsRevisionMarkerElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "moveFromRangeStart" ||
            element.Name == WordprocessingNamespace + "moveFromRangeEnd" ||
            element.Name == WordprocessingNamespace + "moveToRangeStart" ||
            element.Name == WordprocessingNamespace + "moveToRangeEnd";
    }

    private static DocxRevisionKind? RevisionKind(XElement element)
    {
        if (element.Parent?.Name == WordprocessingNamespace + "tcPr")
        {
            if (element.Name == WordprocessingNamespace + "cellIns") return DocxRevisionKind.TableCellInsertion;
            if (element.Name == WordprocessingNamespace + "cellDel") return DocxRevisionKind.TableCellDeletion;
        }
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

    // Semantic-void property-change comparison.
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

    // Structural OOXML property comparison ignoring
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

    // Significant OOXML content (elements plus
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

    private static DocxRevisionInfo? FindInheritedRevision(XElement element, OoxPdfDocxMarkupMode markupMode)
    {
        XElement? revisionContainer = element.Ancestors().FirstOrDefault(DocxSourceXml.IsRevisionContainer);
        return revisionContainer is not null && IsIncludedRevisionContainer(revisionContainer, markupMode)
            ? CreateRevisionInfo(revisionContainer)
            : null;
    }

    private static bool IsInsideExcludedRevisionContainer(XElement element, OoxPdfDocxMarkupMode markupMode)
    {
        return element
            .Ancestors()
            .Any(ancestor => DocxSourceXml.IsRevisionContainer(ancestor) && !IsIncludedRevisionContainer(ancestor, markupMode));
    }

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

    private sealed record DocxRevisionScopedElement(XElement Element, DocxRevisionInfo? Revision)
    {
        public IReadOnlyList<DocxBookmarkAnchor>? LeadingBookmarks { get; init; }
        public DocxTableColumnBookmark[]? LeadingColumnBookmarks { get; init; }
        public IReadOnlyList<XElement>? LeadingCommentBoundaries { get; init; }
    }

    private readonly record struct DocxTableColumnBookmark(DocxBookmarkAnchor Anchor, int First, int Last);

    private static bool NeedsCommentBinding(DocxCommentRange range) =>
        range.StartSourceRunIndex is not null && range.ReferenceSourceRunIndex is null;

    private static bool IsUnmatchedCommentStart(DocxCommentRange range) =>
        NeedsCommentBinding(range) && range.EndSourceRunIndex is null;

    private readonly record struct DocxBlockCommentBoundary(int BeforePart, string Id, bool IsStart)
    {
        public bool Consumed { get; init; }
    }

    // Each flow or cell resolves its own ranges. A nested table or an import
    // ends the surrounding group without moving its table or paragraph owners.
    private sealed class DocxCommentProjection(List<DocxBodyElement> elements)
    {
        private int groupStart;
        private int paragraphBoundary;
        private bool hasPendingRange;
        private List<DocxBlockCommentBoundary>? boundaries;

        public IReadOnlyList<XElement>? BeginParagraph(IReadOnlyList<XElement>? markers, bool mergesWithPrevious, CancellationToken token)
        {
            paragraphBoundary = elements.Count;
            if (markers is null || mergesWithPrevious) return markers;
            foreach (XElement marker in markers)
            {
                token.ThrowIfCancellationRequested();
                if (NormalizeDecimalId((string?)marker.Attribute(WordprocessingNamespace + "id")) is { Length: > 0 } id)
                    (boundaries ??= []).Add(new(paragraphBoundary, id, marker.Name == WordprocessingNamespace + "commentRangeStart"));
            }
            return null;
        }

        public void TrackParagraph(DocxParagraph? paragraph, CancellationToken token)
        {
            // A split fragment can merge without a single inline entry point.
            // Keep those block boundaries unresolved.
            if (elements.Count <= paragraphBoundary && boundaries is not null)
            {
                for (int index = boundaries.Count - 1; index >= 0 && boundaries[index].BeforePart == paragraphBoundary; index--)
                {
                    token.ThrowIfCancellationRequested();
                    boundaries[index] = boundaries[index] with { Consumed = true };
                }
            }
            if (!hasPendingRange && paragraph is not null)
                hasPendingRange = paragraph.CommentRanges.Any(NeedsCommentBinding);
        }

        public void FinishGroup(CancellationToken token)
        {
            if (hasPendingRange || boundaries is { Count: > 0 })
                ProjectCommentRanges(elements, groupStart, boundaries, token);
            groupStart = elements.Count;
            boundaries?.Clear();
            hasPendingRange = false;
        }
    }

    private static void ProjectCommentRanges(List<DocxBodyElement> elements, int firstPart,
        IReadOnlyList<DocxBlockCommentBoundary>? blockBoundaries, CancellationToken token)
    {
        var starts = new List<(int Part, DocxCommentRange Start)>();
        var startIds = new HashSet<string>(StringComparer.Ordinal);
        var lastParts = new Dictionary<string, (int Part, DocxCommentRange End)>(StringComparer.Ordinal);
        var referenceParts = new Dictionary<string, int>(StringComparer.Ordinal);
        var ambiguous = new HashSet<string>(StringComparer.Ordinal);
        int nextBoundary = 0;
        int boundaryCount = blockBoundaries?.Count ?? 0;
        for (int part = firstPart; part <= elements.Count; part++)
        {
            token.ThrowIfCancellationRequested();
            while (nextBoundary < boundaryCount && blockBoundaries![nextBoundary].BeforePart == part)
            {
                token.ThrowIfCancellationRequested();
                var boundary = blockBoundaries[nextBoundary++];
                if (boundary.Consumed)
                {
                    ambiguous.Add(boundary.Id);
                }
                else if (boundary.IsStart)
                {
                    AddStart(boundary.Id, part, new(boundary.Id, 0, 0, null, null, null, null)
                    {
                        StartInlineItemIndex = 0
                    });
                }
                else if (!lastParts.TryAdd(boundary.Id, (part - 1,
                             new(boundary.Id, null, null, int.MaxValue, null, null, null))))
                {
                    ambiguous.Add(boundary.Id);
                }
            }
            if (part == elements.Count) break;
            if (GetParagraph(elements[part]) is not { } paragraph) continue;
            foreach (DocxCommentRange range in paragraph.CommentRanges)
            {
                token.ThrowIfCancellationRequested();
                if (range.Id is not { Length: > 0 } id) continue;
                if (NeedsCommentBinding(range))
                {
                    AddStart(id, part, range);
                }
                else if (range.StartSourceRunIndex is not null)
                {
                    ambiguous.Add(id);
                    continue;
                }
                if (range.EndSourceRunIndex is not null && !lastParts.TryAdd(id, (part, range)))
                    ambiguous.Add(id);
            }
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                token.ThrowIfCancellationRequested();
                if (reference.Kind == DocxRelatedStoryKind.Comment && reference.Id is { } id &&
                    !referenceParts.TryAdd(id, part))
                    ambiguous.Add(id);
            }
        }

        // Resolve a closed selection and its unique original reference within
        // the paragraph sequence. Breaks affect placement but preserve source order.
        var ranges = new List<DocxSplitCommentRange>();
        int[]? boundaries = null;
        foreach (var (part, start) in starts)
        {
            token.ThrowIfCancellationRequested();
            string id = start.Id!;
            if (ambiguous.Contains(id) || !lastParts.TryGetValue(id, out var end) ||
                end.Part < part || !referenceParts.TryGetValue(id, out int referencePart) ||
                referencePart < end.Part || referencePart == end.Part && end.End.ReferenceSourceRunIndex is null)
                continue;

            ranges.Add(new(part, end.Part, start.StartSourceRunIndex!.Value,
                start.StartTextOffset!.Value, start.StartInlineItemIndex, end.End));
            boundaries ??= new int[elements.Count - firstPart + 1];
            boundaries[part - firstPart]++;
            boundaries[end.Part - firstPart + 1]--;
        }

        void AddStart(string id, int part, DocxCommentRange range)
        {
            if (startIds.Add(id))
                starts.Add((part, range));
            else
                ambiguous.Add(id);
        }

        var selections = ranges.Count == 0 ? null : new DocxSplitCommentRanges(ranges, token);
        int active = 0;
        for (int part = firstPart; part < elements.Count; part++)
        {
            token.ThrowIfCancellationRequested();
            active += boundaries?[part - firstPart] ?? 0;
            DocxBodyElement element = elements[part];
            if (GetParagraph(element) is not { } paragraph) continue;
            var localRanges = paragraph.CommentRanges;
            if (localRanges.Any(IsUnmatchedCommentStart))
                localRanges = localRanges.Where(range => !IsUnmatchedCommentStart(range)).ToArray();
            if (active == 0 && ReferenceEquals(localRanges, paragraph.CommentRanges)) continue;

            paragraph = paragraph with
            {
                CommentRanges = localRanges,
                SplitComments = active == 0 ? null : new DocxSplitCommentFragment(part, selections!)
            };
            elements[part] = element switch
            {
                DocxParagraphElement flow => flow with { Paragraph = paragraph },
                DocxPageBreakElement page => page with { BreakParagraph = paragraph },
                DocxManualBreakElement column => column with { BreakParagraph = paragraph },
                _ => element
            };
        }

        static DocxParagraph? GetParagraph(DocxBodyElement element) => element switch
        {
            DocxParagraphElement flow => flow.Paragraph,
            DocxPageBreakElement page => page.BreakParagraph,
            DocxManualBreakElement column => column.BreakParagraph,
            _ => null
        };
    }

    private sealed record DocxCommentRangeStart(string? Id, int SourceRunIndex, int TextOffset)
    {
        public int? InlineItemIndex { get; init; }
        public XElement? FractionNumerator { get; init; }
    }

    private sealed record DocxRevisionRangeStart(
        DocxRevisionKind Kind,
        string? Id,
        string? Name,
        string? Author,
        string? Date,
        int SourceRunIndex,
        int TextOffset)
    {
        public long ContentSourcePosition { get; init; }
        public bool IsNonPrintingBoundary { get; init; }
        public DocxRevisionInfo? MarkerRevision { get; init; }
    }
}
