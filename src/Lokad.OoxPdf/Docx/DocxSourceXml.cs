using System.Diagnostics.CodeAnalysis;
using System.Xml.Linq;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Docx;

internal static class DocxSourceXml
{
    public static bool IsRevisionContainer(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo";

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

    private sealed record PairedMoveOwners(IReadOnlySet<(XName Name, string Id)> Ids);
    private sealed record MoveRange(string Name, List<(XName Name, string Id)> Owners);

    public static XElement? FindStoryRoot(XElement source) =>
        source.AncestorsAndSelf().FirstOrDefault(IsStoryRoot);

    private static bool IsStoryRoot(XElement element) =>
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "body" or "hdr" or "ftr" or "footnote" or "endnote" or "comment" or "txbxContent";

    public static IReadOnlySet<(XName Name, string Id)> ReadPairedMoveOwnerIds(
        XElement equation, XElement? authoredStory, CancellationToken token)
    {
        XElement? story = FindStoryRoot(equation) ?? authoredStory;
        if (story?.Annotation<PairedMoveOwners>() is { } cached) return cached.Ids;
        var ids = CollectPairedMoveOwnerIds(story ?? equation, token);
        story?.AddAnnotation(new PairedMoveOwners(ids));
        return ids;
    }

    private static HashSet<(XName Name, string Id)> CollectPairedMoveOwnerIds(XElement story, CancellationToken token)
    {
        var open = new Dictionary<(DocxRevisionKind Kind, string Id), MoveRange?>();
        var closed = new Dictionary<(DocxRevisionKind Kind, string Name), MoveRange?>();
        var declaredIds = new HashSet<(XName Name, string Id)>();
        var duplicateIds = new HashSet<(XName Name, string Id)>();
        var pending = new Stack<XElement>(StoredChildren(story).Reverse());
        while (pending.TryPop(out var child))
        {
            token.ThrowIfCancellationRequested();
            if (IsStoryRoot(child)) continue;
            foreach (var item in StoredChildren(child).Reverse()) pending.Push(item);
            if ((string?)child.Attribute(WordprocessingNamespace + "id") is not { Length: > 0 } id) continue;

            if (TryResolveRevisionRangeMarker(child, out var kind, out bool isStart))
            {
                var key = (kind.Value, id);
                bool boundary = IsNonPrintingRevisionRangeBoundary(child);
                if (isStart)
                {
                    // An overlap invalidates the active range. Keep each end
                    // independent so later, disjoint pairs can still resolve.
                    if (open.Count == 1) open[open.Keys.First()] = null;
                    MoveRange? range = open.Count == 0 && boundary &&
                        (string?)child.Attribute(WordprocessingNamespace + "name") is { Length: > 0 } name
                        ? new(name, []) : null;
                    open.TryAdd(key, range);
                }
                else if (open.Remove(key, out var range) && boundary && range is not null)
                {
                    var named = (kind.Value, range.Name);
                    if (!closed.TryAdd(named, range)) closed[named] = null;
                }
            }
            else if (child.Name == WordprocessingNamespace + "moveFrom" || child.Name == WordprocessingNamespace + "moveTo")
            {
                var owner = (child.Name, id);
                if (!declaredIds.Add(owner)) duplicateIds.Add(owner);
                if (open.Count != 1) continue;
                var range = open.First();
                DocxRevisionKind side = child.Name.LocalName == "moveFrom" ? DocxRevisionKind.MoveFrom : DocxRevisionKind.MoveTo;
                if (range.Key.Kind == side) range.Value?.Owners.Add(owner);
            }
        }

        // A move has a closed source and destination with the same name.
        // Unique revision IDs carry these owners through cloned page-break fragments.
        var owners = new HashSet<(XName Name, string Id)>();
        foreach (var (key, from) in closed)
        {
            token.ThrowIfCancellationRequested();
            if (key.Kind == DocxRevisionKind.MoveFrom && from is { Owners.Count: > 0 } &&
                closed.TryGetValue((DocxRevisionKind.MoveTo, key.Name), out var to) && to is { Owners.Count: > 0 } &&
                !from.Owners.Any(duplicateIds.Contains) && !to.Owners.Any(duplicateIds.Contains))
            {
                owners.UnionWith(from.Owners);
                owners.UnionWith(to.Owners);
            }
        }
        return owners;
    }

    // These wrappers contribute stored content without changing inline formatting.
    // Revisions, fields, links and directional/unknown owners need their own rules.
    public static bool IsInlineMetadataContainer(XElement element) =>
        element.Name == WordprocessingNamespace + "customXml" ||
        element.Name == WordprocessingNamespace + "smartTag";

    public static bool IsTransparentInlineContainer(XElement element) =>
        IsInlineMetadataContainer(element) || element.Name == WordprocessingNamespace + "sdt" ||
        element.Name == WordprocessingNamespace + "sdtContent";

    // Editing boundaries and stored math pagination positions have no payload.
    // Content-bearing, unknown and misplaced children retain their source owner.
    public static bool IsNonPrintingMarker(XElement element) =>
        element.Parent?.Name == MathNamespace + "r" && IsLastRenderedPageBreakHint(element) ||
        element.Name.Namespace == WordprocessingNamespace &&
        element.Name.LocalName is "proofErr" or "permStart" or "permEnd" &&
        !element.HasElements && string.IsNullOrWhiteSpace(element.Value);

    public static bool IsLastRenderedPageBreakHint(XElement element) =>
        element.Name == WordprocessingNamespace + "lastRenderedPageBreak" &&
        (element.Parent?.Name == WordprocessingNamespace + "r" || element.Parent?.Name == MathNamespace + "r") &&
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
