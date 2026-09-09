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
    private static IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> ReadReferencedHeaderFooterBodyElementsByType(
        XContainer referenceRoot,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        string relationshipType,
        string referenceElementName,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        HashSet<string>? warnedParts = null)
    {
        var bodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);
        foreach (XElement reference in referenceRoot.Descendants(WordprocessingNamespace + referenceElementName))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? relationshipId = (string?)reference.Attribute(RelationshipsNamespace + "id");
            if (relationshipId is null || !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.Type != relationshipType || relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxPart? part = package.GetPart(relationship.ResolvedTarget);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument partXml = SafeXml.Load(stream, cancellationToken);
            OoxMarkupCompatibility.WarnMustUnderstandOnce(partXml, part.Name, diagnosticSink, warnedParts);
            string type = (string?)reference.Attribute(WordprocessingNamespace + "type") ?? "default";
            IReadOnlyDictionary<string, OoxRelationship> partRelationships = package.GetRelationships(part.Name, cancellationToken)
                .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
                .ToDictionary(r => r.Id, StringComparer.Ordinal);
            bodyElementsByType[type] = ReadRelatedStoryBodyElements(
                partXml.Root?.Elements() ?? [],
                styles,
                numbering,
                new Dictionary<(string NumId, int Level), int>(),
                package,
                partRelationships,
                markupMode,
                cancellationToken);
        }

        return bodyElementsByType;
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> ToStaticParagraphsByType(
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> bodyElementsByType)
    {
        return bodyElementsByType.ToDictionary(
            pair => pair.Key,
            pair => (IReadOnlyList<DocxParagraph>)DocxBlockTraversal.EnumerateDirectParagraphs(pair.Value).ToArray(),
            StringComparer.OrdinalIgnoreCase);
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> ReadReferencedHeaderFooterFloatingDrawingsByType(
        XContainer referenceRoot,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        string relationshipType,
        string referenceElementName,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        HashSet<string>? warnedParts = null)
    {
        var drawings = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);
        foreach (XElement reference in referenceRoot.Descendants(WordprocessingNamespace + referenceElementName))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? relationshipId = (string?)reference.Attribute(RelationshipsNamespace + "id");
            if (relationshipId is null || !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.Type != relationshipType || relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxPart? part = package.GetPart(relationship.ResolvedTarget);
            if (part is null)
            {
                continue;
            }

            using Stream stream = part.OpenRead();
            XDocument partXml = SafeXml.Load(stream, cancellationToken);
            OoxMarkupCompatibility.WarnMustUnderstandOnce(partXml, part.Name, diagnosticSink, warnedParts);
            IReadOnlyDictionary<string, OoxRelationship> partRelationships = package.GetRelationships(part.Name, cancellationToken)
                .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
                .ToDictionary(r => r.Id, StringComparer.Ordinal);
            string type = (string?)reference.Attribute(WordprocessingNamespace + "type") ?? "default";
            drawings[type] = ReadFloatingDrawings(partXml, package, partRelationships, styles, numbering, markupMode, cancellationToken);
        }

        return drawings;
    }

    // Single caller; kept static: entry-point stage, not a local candidate.
    private static IReadOnlyList<DocxParagraph> SelectDefaultHeaderFooterParagraphs(IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> paragraphsByType)
    {
        return paragraphsByType.TryGetValue("default", out IReadOnlyList<DocxParagraph>? paragraphs)
            ? paragraphs
            : [];
    }

    private static IReadOnlyList<DocxRelatedStory> ReadRelatedStories(
        OoxPackage package,
        string documentPartName,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        HashSet<string>? warnedParts = null)
    {
        return ReadCommentStories(package, documentPartName, styles, numbering, markupMode, cancellationToken, diagnosticSink, warnedParts)
            .Concat(ReadRelatedStories(package, documentPartName, styles, numbering, FootnotesRelationshipType, FootnotesContentType, DocxRelatedStoryKind.Footnote, "footnote", markupMode, cancellationToken, null, diagnosticSink, warnedParts))
            .Concat(ReadRelatedStories(package, documentPartName, styles, numbering, EndnotesRelationshipType, EndnotesContentType, DocxRelatedStoryKind.Endnote, "endnote", markupMode, cancellationToken, null, diagnosticSink, warnedParts))
            .ToArray();
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IReadOnlyList<DocxRelatedStory> ReadCommentStories(
        OoxPackage package,
        string documentPartName,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        HashSet<string>? warnedParts = null)
    {
        IReadOnlyDictionary<string, DocxCommentThreadMetadata> threadMetadataByParagraphId =
            ReadCommentThreadMetadata(package, documentPartName, cancellationToken);
        IReadOnlyList<DocxRelatedStory> stories = ReadRelatedStories(
            package,
            documentPartName,
            styles,
            numbering,
            CommentsRelationshipType,
            CommentsContentType,
            DocxRelatedStoryKind.Comment,
            "comment",
            markupMode,
            cancellationToken,
            threadMetadataByParagraphId,
            diagnosticSink,
            warnedParts);
        return ResolveCommentThreadParents(stories);
    }

    private static IReadOnlyList<DocxRelatedStory> ReadRelatedStories(
        OoxPackage package,
        string documentPartName,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        string relationshipType,
        string contentType,
        DocxRelatedStoryKind kind,
        string storyElementName,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        IReadOnlyDictionary<string, DocxCommentThreadMetadata>? commentThreadMetadataByParagraphId,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        HashSet<string>? warnedParts = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? part = FindRelatedPart(package, documentPartName, relationshipType, contentType, cancellationToken);
        if (part is null)
        {
            return [];
        }

        using Stream stream = part.OpenRead();
        XDocument partXml = SafeXml.Load(stream, cancellationToken);
        OoxMarkupCompatibility.WarnMustUnderstandOnce(partXml, part.Name, diagnosticSink, warnedParts);
        IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(part.Name, cancellationToken)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        var numberingCounters = new Dictionary<(string NumId, int Level), int>();
        var stories = new List<DocxRelatedStory>();
        foreach (XElement storyElement in partXml.Root?.Elements(WordprocessingNamespace + storyElementName) ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRelatedStory story = ReadRelatedStory(kind, part.Name, storyElement, styles, numbering, numberingCounters, package, relationships, markupMode, cancellationToken, commentThreadMetadataByParagraphId);
            if (story.BodyElements.Count > 0)
            {
                stories.Add(story);
            }
        }

        return stories;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxRelatedStory ReadRelatedStory(
        DocxRelatedStoryKind kind,
        string partName,
        XElement story,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        IReadOnlyDictionary<string, DocxCommentThreadMetadata>? commentThreadMetadataByParagraphId)
    {
            string? ReadCommentParagraphId(XElement story)
            {
                return story
                    .Descendants(WordprocessingNamespace + "p")
                    .Select(paragraph => (string?)paragraph.Attribute(Office2010WordNamespace + "paraId"))
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            }

        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<DocxBodyElement> bodyElements = ReadRelatedStoryBodyElements(
            story.Elements(),
            styles,
            numbering,
            numberingCounters,
            package,
            relationships,
            markupMode,
            cancellationToken);
        IReadOnlyList<DocxFloatingDrawing> floatingDrawings = ReadFloatingDrawings(story, package, relationships, styles, numbering, markupMode, cancellationToken);
        string? paragraphId = kind == DocxRelatedStoryKind.Comment ? ReadCommentParagraphId(story) : null;
        DocxCommentThreadMetadata? threadMetadata = paragraphId is not null && commentThreadMetadataByParagraphId is not null && commentThreadMetadataByParagraphId.TryGetValue(paragraphId, out DocxCommentThreadMetadata? metadata)
            ? metadata
            : null;
        return new DocxRelatedStory(
            kind,
            partName,
            (string?)story.Attribute(WordprocessingNamespace + "id"),
            bodyElements,
            [],
            [],
            ParseRelatedStoryType((string?)story.Attribute(WordprocessingNamespace + "type")))
        {
            FloatingDrawings = floatingDrawings,
            CommentMetadata = kind == DocxRelatedStoryKind.Comment
                ? new DocxCommentMetadata(
                    (string?)story.Attribute(WordprocessingNamespace + "author"),
                    (string?)story.Attribute(WordprocessingNamespace + "initials"),
                    (string?)story.Attribute(WordprocessingNamespace + "date"),
                    paragraphId,
                    threadMetadata?.ParentParagraphId,
                    null,
                    threadMetadata?.IsResolved)
                : null
        };

        DocxRelatedStoryType? ParseRelatedStoryType(string? value)
        {
            if (value is null)
            {
                return null;
            }

            if (value.Equals("normal", StringComparison.OrdinalIgnoreCase))
            {
                return DocxRelatedStoryType.Normal;
            }

            if (value.Equals("separator", StringComparison.OrdinalIgnoreCase))
            {
                return DocxRelatedStoryType.Separator;
            }

            if (value.Equals("continuationSeparator", StringComparison.OrdinalIgnoreCase))
            {
                return DocxRelatedStoryType.ContinuationSeparator;
            }

            if (value.Equals("continuationNotice", StringComparison.OrdinalIgnoreCase))
            {
                return DocxRelatedStoryType.ContinuationNotice;
            }

            return DocxRelatedStoryType.Unknown;
        }
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IReadOnlyDictionary<string, DocxCommentThreadMetadata> ReadCommentThreadMetadata(
        OoxPackage package,
        string documentPartName,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? part = FindRelatedPart(package, documentPartName, CommentsExtendedRelationshipType, CommentsExtendedContentType, cancellationToken);
        if (part is null)
        {
            return new Dictionary<string, DocxCommentThreadMetadata>(StringComparer.OrdinalIgnoreCase);
        }

        using Stream stream = part.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        var metadataByParagraphId = new Dictionary<string, DocxCommentThreadMetadata>(StringComparer.OrdinalIgnoreCase);
        foreach (XElement element in document.Root?.Elements(Office2012WordNamespace + "commentEx") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? paragraphId = (string?)element.Attribute(Office2012WordNamespace + "paraId");
            if (string.IsNullOrWhiteSpace(paragraphId))
            {
                continue;
            }

            metadataByParagraphId[paragraphId] = new DocxCommentThreadMetadata(
                (string?)element.Attribute(Office2012WordNamespace + "paraIdParent"),
                element.Attribute(Office2012WordNamespace + "done") is { } done ? OoxBoolean.IsTrue(done.Value) : null);
        }

        return metadataByParagraphId;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IReadOnlyList<DocxRelatedStory> ResolveCommentThreadParents(IReadOnlyList<DocxRelatedStory> stories)
    {
        Dictionary<string, string> commentIdByParagraphId = stories
            .Where(story => story.CommentMetadata?.ParagraphId is not null && story.Id is not null)
            .GroupBy(story => story.CommentMetadata?.ParagraphId ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First().Id ?? string.Empty, StringComparer.OrdinalIgnoreCase);
        return stories
            .Select(story =>
            {
                DocxCommentMetadata? metadata = story.CommentMetadata;
                if (metadata?.ParentParagraphId is null ||
                    !commentIdByParagraphId.TryGetValue(metadata.ParentParagraphId, out string? parentCommentId))
                {
                    return story;
                }

                return story with
                {
                    CommentMetadata = metadata with { ParentCommentId = parentCommentId }
                };
            })
            .ToArray();
    }

}
