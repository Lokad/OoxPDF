using System.Diagnostics.CodeAnalysis;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static class DocxSourceXml
{
    public static bool IsRevisionContainer(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo";

    public static bool IsPackedMathRunRevision(XElement? element) =>
        element is not null && element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "ins" or "del" && element.Parent?.Name == MathNamespace + "r";

    public static bool TryResolveRevisionRangeMarker(XElement element, [NotNullWhen(true)] out DocxRevisionKind? kind, out bool isStart)
    {
        (DocxRevisionKind? Kind, bool IsStart) marker = element.Name.Namespace == WordprocessingNamespace
            ? element.Name.LocalName switch
            {
                "moveFromRangeStart" => (DocxRevisionKind.MoveFrom, true),
                "moveFromRangeEnd" => (DocxRevisionKind.MoveFrom, false),
                "moveToRangeStart" => (DocxRevisionKind.MoveTo, true),
                "moveToRangeEnd" => (DocxRevisionKind.MoveTo, false),
                _ => (null, false)
            }
            : (null, false);
        (kind, isStart) = marker;
        return kind is not null;
    }

    public static bool IsNonPrintingRevisionRangeBoundary(XElement element) =>
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value) &&
        !string.IsNullOrWhiteSpace((string?)element.Attribute(WordprocessingNamespace + "id")) &&
        element.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name.Namespace == WordprocessingNamespace &&
            attribute.Name.LocalName is "id" or "name" or "author" or "date");

    public sealed record StoryOrigin(XElement Story);

    private sealed record PairedMoveOwners(IReadOnlySet<(XName Name, string Id)> Ids, IReadOnlySet<XElement> MathSources);
    private sealed record MoveRange(string Name, XElement Start, List<XElement> Owners);

    public static XElement? FindStoryRoot(XElement source) =>
        source.AncestorsAndSelf().FirstOrDefault(IsStoryRoot);

    public static XElement? FindFieldFlowOwner(XElement source) =>
        source.AncestorsAndSelf().FirstOrDefault(element =>
            element.Name == WordprocessingNamespace + "tc" || IsStoryRoot(element));

    private static bool IsStoryRoot(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "body" or "hdr" or "ftr" or "footnote" or "endnote" or "comment" or "txbxContent";

    public static IReadOnlySet<(XName Name, string Id)> ReadPairedMoveOwnerIds(
        XElement equation, XElement? authoredStory, CancellationToken token)
    {
        XElement? story = FindStoryRoot(equation) ?? authoredStory;
        return story is null ? CollectPairedMoveOwners(equation, token).Ids : GetPairedMoveOwners(story, token).Ids;
    }

    public static bool IsAdmittedMathRevisionOwner(XElement element, XElement equation, XElement? authoredStory,
        ref IReadOnlySet<(XName Name, string Id)>? moveOwners, CancellationToken token) =>
        IsRevisionContainer(element) &&
        (element.Name.LocalName is not ("moveFrom" or "moveTo") ||
            IsPairedMoveOwner(element, equation, authoredStory, ref moveOwners, token));

    public static bool IsPairedMoveOwner(XElement owner, XElement source, XElement? authoredStory,
        ref IReadOnlySet<(XName Name, string Id)>? moveOwners, CancellationToken token) =>
        OoxXml.NormalizeDecimalId((string?)owner.Attribute(WordprocessingNamespace + "id")) is { } id &&
        (moveOwners ??= ReadPairedMoveOwnerIds(source, authoredStory, token)).Contains((owner.Name, id));

    public static bool IsPairedMathMoveSource(XElement element, CancellationToken token) =>
        IsInEquation(element) && FindStoryRoot(element) is { } story &&
        GetPairedMoveOwners(story, token).MathSources.Contains(element);

    public static DocxInlineMoveOwnership GetInlineMoveOwnership(XElement start, XElement end, DocxRevisionKind kind,
        XElement? authoredStory, CancellationToken token)
    {
        var name = WordprocessingNamespace + (kind == DocxRevisionKind.MoveFrom ? "moveFrom" : "moveTo");
        if (!ReferenceEquals(start.Parent, end.Parent)) return DocxInlineMoveOwnership.Unresolved;
        IReadOnlySet<(XName Name, string Id)>? pairedOwners = null;
        Stack<XElement>? stored = null;
        bool hasOwner = false, hasOtherContent = false;
        // A partial pair establishes local boundaries and explicit source
        // owners. Other content still has unresolved range ownership.
        foreach (var sibling in start.ElementsAfterSelf())
        {
            token.ThrowIfCancellationRequested();
            if (ReferenceEquals(sibling, end))
                return !hasOwner ? DocxInlineMoveOwnership.Unresolved :
                    hasOtherContent ? DocxInlineMoveOwnership.Partial : DocxInlineMoveOwnership.Complete;

            // Use the parser's stored payload rule for metadata inside the
            // range. Only actual move wrappers can prove revision ownership.
            XElement child = sibling;
            while (true)
            {
                token.ThrowIfCancellationRequested();
                if (IsTransparentInlineContainer(child))
                {
                    stored ??= new();
                    foreach (var payload in StoredChildren(child))
                    {
                        token.ThrowIfCancellationRequested();
                        stored.Push(payload);
                    }
                }
                else if (!(IsNonPrintingMarker(child) || IsBookmarkMarker(child) || IsCommentBoundary(child)))
                {
                    if (child.Name == name && IsPairedMoveOwner(child, start, authoredStory, ref pairedOwners, token))
                        hasOwner = true;
                    else hasOtherContent = true;
                }
                if (stored is null || !stored.TryPop(out var next)) break;
                child = next;
            }
        }
        return DocxInlineMoveOwnership.Unresolved;
    }

    private static bool IsInEquation(XElement element) => element.Ancestors(MathNamespace + "oMath").Any();

    private static PairedMoveOwners GetPairedMoveOwners(XElement story, CancellationToken token)
    {
        if (story.Annotation<PairedMoveOwners>() is { } cached) return cached;
        var paired = CollectPairedMoveOwners(story, token);
        story.AddAnnotation(paired);
        return paired;
    }

    private static PairedMoveOwners CollectPairedMoveOwners(XElement story, CancellationToken token)
    {
        var open = new Dictionary<(DocxRevisionKind Kind, string Id), MoveRange?>();
        var closed = new Dictionary<(DocxRevisionKind Kind, string Name), (MoveRange Range, XElement End)?>();
        var declaredIds = new HashSet<(XName Name, string Id)>();
        var duplicateIds = new HashSet<(XName Name, string Id)>();
        var pending = new Stack<XElement>(StoredChildren(story).Reverse());
        while (pending.TryPop(out var child))
        {
            token.ThrowIfCancellationRequested();
            if (IsStoryRoot(child)) continue;
            foreach (var item in StoredChildren(child).Reverse()) pending.Push(item);
            if ((string?)child.Attribute(WordprocessingNamespace + "id") is not { Length: > 0 } authoredId) continue;

            if (TryResolveRevisionRangeMarker(child, out var kind, out bool isStart))
            {
                var key = (kind.Value, OoxXml.NormalizeDecimalId(authoredId)!);
                bool boundary = IsNonPrintingRevisionRangeBoundary(child);
                if (isStart)
                {
                    // An overlap invalidates the active range. Keep each end
                    // independent so later, disjoint pairs can still resolve.
                    if (open.Count == 1) open[open.Keys.First()] = null;
                    MoveRange? range = open.Count == 0 && boundary &&
                        (string?)child.Attribute(WordprocessingNamespace + "name") is { Length: > 0 } name
                        ? new(name, child, []) : null;
                    open.TryAdd(key, range);
                }
                else if (open.Remove(key, out var range) && boundary && range is not null)
                {
                    var named = (kind.Value, range.Name);
                    if (!closed.TryAdd(named, (range, child))) closed[named] = null;
                }
            }
            else if (child.Name == WordprocessingNamespace + "moveFrom" || child.Name == WordprocessingNamespace + "moveTo")
            {
                var owner = OwnerId(child);
                if (!declaredIds.Add(owner)) duplicateIds.Add(owner);
                if (open.Count != 1) continue;
                var range = open.First();
                DocxRevisionKind side = child.Name.LocalName == "moveFrom" ? DocxRevisionKind.MoveFrom : DocxRevisionKind.MoveTo;
                if (range.Key.Kind == side) range.Value?.Owners.Add(child);
            }
        }

        // A move has a closed source and destination with the same name.
        // Unique revision IDs carry these owners through cloned page-break fragments.
        var owners = new HashSet<(XName Name, string Id)>();
        var mathSources = new HashSet<XElement>();
        foreach (var (key, from) in closed)
        {
            token.ThrowIfCancellationRequested();
            if (key.Kind != DocxRevisionKind.MoveFrom || from is not { } source ||
                !closed.TryGetValue((DocxRevisionKind.MoveTo, key.Name), out var to) || to is not { } destination ||
                !HasUniqueOwners(source.Range) || !HasUniqueOwners(destination.Range))
            {
                continue;
            }
            owners.UnionWith(source.Range.Owners.Select(OwnerId));
            owners.UnionWith(destination.Range.Owners.Select(OwnerId));
            // Math has its own fallback diagnostics at placement. The broad
            // revision probe can defer to them when both move ranges belong to it.
            if (IsMathRange(source) && IsMathRange(destination))
            {
                AddMathSources(source);
                AddMathSources(destination);
            }
        }
        return new(owners, mathSources);

        bool HasUniqueOwners(MoveRange range) =>
            range.Owners.Count != 0 && !range.Owners.Select(OwnerId).Any(duplicateIds.Contains);
        static (XName Name, string Id) OwnerId(XElement owner) =>
            (owner.Name, OoxXml.NormalizeDecimalId(owner.Attribute(WordprocessingNamespace + "id")!.Value)!);
        static bool IsMathRange((MoveRange Range, XElement End) move) =>
            IsInEquation(move.Range.Start) && IsInEquation(move.End) && move.Range.Owners.All(IsInEquation);
        void AddMathSources((MoveRange Range, XElement End) move)
        {
            mathSources.Add(move.Range.Start);
            mathSources.UnionWith(move.Range.Owners);
            mathSources.Add(move.End);
        }
    }

    // These wrappers contribute stored content without changing inline formatting.
    // Revisions, fields, links and directional/unknown owners need their own rules.
    public static bool IsInlineMetadataContainer(XElement element) =>
        element.Name == WordprocessingNamespace + "customXml" ||
        element.Name == WordprocessingNamespace + "smartTag";

    public static bool IsTransparentInlineContainer(XElement element) =>
        IsInlineMetadataContainer(element) || element.Name == WordprocessingNamespace + "sdt" ||
        element.Name == WordprocessingNamespace + "sdtContent";

    public static bool IsDirectionalInlineContainer(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace && element.Name.LocalName is "dir" or "bdo";

    // Directional markup retains stored content through the source-order fallback.
    // Only transparent wrappers carry structural math through an excluded deletion.
    public static bool IsInlineContentContainer(XElement element) =>
        IsTransparentInlineContainer(element) || IsDirectionalInlineContainer(element);

    public static bool IsTextWrappingBreak(XElement element) =>
        element.Name == WordprocessingNamespace + "cr" ||
        element.Name == WordprocessingNamespace + "br" &&
        (string?)element.Attribute(WordprocessingNamespace + "type") is null or "" or "textWrapping";

    // Editing boundaries and ignored math-run controls have no payload.
    // Content-bearing, unknown and misplaced children retain their source owner.
    public static bool IsNonPrintingMarker(XElement element) =>
        IsNonPrintingMathRunControl(element) ||
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "proofErr" or "permStart" or "permEnd" &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value);

    // The math run owns these controls even when its revision stores them.
    // Ordinary WordprocessingML runs still use tabs for positioning.
    public static bool IsNonPrintingMathRunControl(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "lastRenderedPageBreak" or "tab" &&
        (element.Parent?.Name == MathNamespace + "r" || IsPackedMathRunRevision(element.Parent)) &&
        IsEmptyRunControl(element);

    public static bool IsLastRenderedPageBreakHint(XElement element) =>
        element.Name == WordprocessingNamespace + "lastRenderedPageBreak" &&
        (element.Parent?.Name == WordprocessingNamespace + "r" || element.Parent?.Name == MathNamespace + "r") &&
        IsEmptyRunControl(element);

    public static bool IsBookmarkMarker(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "bookmarkStart" or "bookmarkEnd" &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value);

    public static bool IsCommentBoundary(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "commentRangeStart" or "commentRangeEnd" &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value);

    private static bool IsEmptyRunControl(XElement element) =>
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value) &&
        element.Attributes().All(attribute => attribute.IsNamespaceDeclaration);

    // Container properties describe the source, not stored inline payload.
    // Ignore only the property's actual owner; unknown/misplaced children stay.
    public static IEnumerable<XElement> StoredChildren(XElement element) =>
        element.Name == WordprocessingNamespace + "sdt"
            ? element.Elements(WordprocessingNamespace + "sdtContent")
            : element.Name == WordprocessingNamespace + "customXml"
                ? element.Elements().Where(child => child.Name != WordprocessingNamespace + "customXmlPr")
                : element.Name == WordprocessingNamespace + "smartTag"
                    ? element.Elements().Where(child => child.Name != WordprocessingNamespace + "smartTagPr")
                    : element.Elements();
}
