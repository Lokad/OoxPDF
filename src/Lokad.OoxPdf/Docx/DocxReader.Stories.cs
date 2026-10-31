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
    // header/footer body and drawing extraction loaded the same parts
    // separately and repeated across references and sections. One entry per unique
    // part per Read: XML comes from the shared package cache, body elements and
    // drawings are each built once, and every referencing type reuses them. Results
    // are deterministic functions of (part bytes, styles, numbering, markup mode),
    // all constant within a Read, so sharing is exact. Diagnostics still fire once
    // per part via the shared warned-parts set.
    private sealed class RelatedStoryPartCache(IMemoryReservation? modelStorage)
    {
        private readonly Dictionary<string, ParsedRelatedStoryPart> parts = new(StringComparer.OrdinalIgnoreCase);

        public IEnumerable<DocxEquationPartSource> EnumerateEquationParts(CancellationToken token)
        {
            foreach (var part in parts)
            {
                token.ThrowIfCancellationRequested();
                if (part.Value.BodyElements is not { } body || part.Value.Xml is not { } xml) { continue; }
                int count = 0;
                foreach (XElement equation in xml.Descendants(MathNamespace + "oMath"))
                { token.ThrowIfCancellationRequested(); count++; }
                if (count > 0) { yield return new(part.Key, count, body); }
            }
        }

        public void Admit(XDocument source, CancellationToken token) => OoxMemory.GrowModel(modelStorage, source, token);

        public ParsedRelatedStoryPart GetOrCreate(string partName)
        {
            if (!parts.TryGetValue(partName, out ParsedRelatedStoryPart? parsed))
            {
                parsed = new ParsedRelatedStoryPart();
                parts[partName] = parsed;
            }

            return parsed;
        }
    }

    private sealed class ParsedRelatedStoryPart
    {
        public XDocument? Xml;
        public bool MustUnderstandVerdictRendered;
        public IReadOnlyDictionary<string, OoxRelationship>? Relationships;
        public IReadOnlyList<DocxBodyElement>? BodyElements;
        public IReadOnlyList<DocxFloatingDrawing>? Drawings;
    }

    // Section-break page settings parse stories without a diagnostic sink while the
    // document pass warns live. Firing the warning only on parse would let a sink-less
    // first touch swallow it; firing per reference would rescan shared DOMs. Render the
    // verdict once per part on the first touch carrying a live sink instead.
    private static void WarnStoryMustUnderstandOnce(
        ParsedRelatedStoryPart parsed,
        string partName,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        HashSet<string>? warnedParts)
    {
        if (diagnosticSink is null || warnedParts is null || parsed.MustUnderstandVerdictRendered || parsed.Xml is null)
        {
            return;
        }

        parsed.MustUnderstandVerdictRendered = true;
        OoxMarkupCompatibility.WarnMustUnderstandOnce(parsed.Xml, partName, diagnosticSink, warnedParts);
    }

    private static IEnumerable<(XElement Reference, OoxRelationship Relationship)> EnumerateReferencedStaticStories(
        XContainer referenceRoot,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        string relationshipType,
        string referenceElementName,
        CancellationToken cancellationToken)
    {
        foreach (XElement reference in referenceRoot.Descendants(WordprocessingNamespace + referenceElementName))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? id = (string?)reference.Attribute(RelationshipsNamespace + "id");
            if (id is not null && relationships.TryGetValue(id, out OoxRelationship? relationship) &&
                !relationship.IsExternal && relationship.Type == relationshipType && relationship.ResolvedTarget is not null)
                yield return (reference, relationship);
        }
    }

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
        RelatedStoryPartCache storyCache,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        HashSet<string>? warnedParts = null)
    {
        var bodyElementsByType = new Dictionary<string, IReadOnlyList<DocxBodyElement>>(StringComparer.OrdinalIgnoreCase);
        foreach (var (reference, relationship) in EnumerateReferencedStaticStories(
            referenceRoot, relationships, relationshipType, referenceElementName, cancellationToken))
        {
            OoxPart? part = package.GetPart(relationship.ResolvedTarget!);
            if (part is null)
            {
                continue;
            }

            ParsedRelatedStoryPart parsed = storyCache.GetOrCreate(part.Name);
            if (parsed.Xml is null)
            {
                parsed.Xml = package.LoadXml(part, cancellationToken);
                storyCache.Admit(parsed.Xml, cancellationToken);
            }

            WarnStoryMustUnderstandOnce(parsed, part.Name, diagnosticSink, warnedParts);
            parsed.Relationships ??= package.GetRelationships(part.Name, cancellationToken)
                .ToDictionary(r => r.Id, StringComparer.Ordinal);
            parsed.BodyElements ??= ReadRelatedStoryBodyElements(
                parsed.Xml.Root?.Elements() ?? [],
                styles,
                numbering,
                new DocxNumberingCounters(),
                package,
                parsed.Relationships,
                markupMode,
                cancellationToken);
            string type = (string?)reference.Attribute(WordprocessingNamespace + "type") ?? "default";
            bodyElementsByType[type] = parsed.BodyElements;
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
        RelatedStoryPartCache storyCache,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        HashSet<string>? warnedParts = null)
    {
        var drawings = new Dictionary<string, IReadOnlyList<DocxFloatingDrawing>>(StringComparer.OrdinalIgnoreCase);
        foreach (var (reference, relationship) in EnumerateReferencedStaticStories(
            referenceRoot, relationships, relationshipType, referenceElementName, cancellationToken))
        {
            OoxPart? part = package.GetPart(relationship.ResolvedTarget!);
            if (part is null)
            {
                continue;
            }

            ParsedRelatedStoryPart parsed = storyCache.GetOrCreate(part.Name);
            if (parsed.Xml is null)
            {
                parsed.Xml = package.LoadXml(part, cancellationToken);
                storyCache.Admit(parsed.Xml, cancellationToken);
            }

            WarnStoryMustUnderstandOnce(parsed, part.Name, diagnosticSink, warnedParts);
            parsed.Relationships ??= package.GetRelationships(part.Name, cancellationToken)
                .ToDictionary(r => r.Id, StringComparer.Ordinal);
            parsed.Drawings ??= ReadFloatingDrawings(parsed.Xml, package, parsed.Relationships, styles, numbering, markupMode, cancellationToken);
            string type = (string?)reference.Attribute(WordprocessingNamespace + "type") ?? "default";
            drawings[type] = parsed.Drawings;
        }

        return drawings;
    }

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
        HashSet<string>? warnedParts = null,
        IReadOnlyList<DocxBodyElement>? mainBodyElements = null,
        XDocument? settings = null, IMemoryReservation? modelStorage = null)
    {
        var noteReferences = new Dictionary<(DocxRelatedStoryKind Kind, string Id), (string? Label, int Order)>();
        foreach (DocxParagraph paragraph in DocxBlockTraversal.EnumerateBodyParagraphs(mainBodyElements ?? []))
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (DocxInlineReference reference in paragraph.InlineReferences)
            {
                if (reference.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
                    reference.Id is { } id)
                {
                    noteReferences.TryAdd((reference.Kind, id), (reference.DisplayText, noteReferences.Count));
                }
            }
        }
        return ReadCommentStories(package, documentPartName, styles, numbering, markupMode, cancellationToken, diagnosticSink, warnedParts, modelStorage)
            .Concat(ReadRelatedStories(package, documentPartName, styles, numbering, FootnotesRelationshipType, DocxRelatedStoryKind.Footnote, "footnote", markupMode, cancellationToken, null, diagnosticSink, warnedParts, noteReferences, settings, modelStorage))
            .Concat(ReadRelatedStories(package, documentPartName, styles, numbering, EndnotesRelationshipType, DocxRelatedStoryKind.Endnote, "endnote", markupMode, cancellationToken, null, diagnosticSink, warnedParts, noteReferences, settings, modelStorage))
            .ToArray();
    }

    private static IReadOnlyList<DocxRelatedStory> ReadCommentStories(
        OoxPackage package,
        string documentPartName,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        HashSet<string>? warnedParts = null, IMemoryReservation? modelStorage = null)
    {
        IReadOnlyDictionary<string, DocxCommentThreadMetadata> threadMetadataByParagraphId =
            ReadCommentThreadMetadata(package, documentPartName, cancellationToken, modelStorage);
        IReadOnlyList<DocxRelatedStory> stories = ReadRelatedStories(
            package,
            documentPartName,
            styles,
            numbering,
            CommentsRelationshipType,
            DocxRelatedStoryKind.Comment,
            "comment",
            markupMode,
            cancellationToken,
            threadMetadataByParagraphId,
            diagnosticSink,
            warnedParts, modelStorage: modelStorage);
        return ResolveCommentThreadParents(stories);
    }

    private static IReadOnlyList<DocxRelatedStory> ReadRelatedStories(
        OoxPackage package,
        string documentPartName,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        string relationshipType,
        DocxRelatedStoryKind kind,
        string storyElementName,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        IReadOnlyDictionary<string, DocxCommentThreadMetadata>? commentThreadMetadataByParagraphId,
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        HashSet<string>? warnedParts = null,
        IReadOnlyDictionary<(DocxRelatedStoryKind Kind, string Id), (string? Label, int Order)>? noteReferences = null,
        XDocument? settings = null, IMemoryReservation? modelStorage = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? part = FindRelatedPart(package, documentPartName, relationshipType, cancellationToken);
        if (part is null)
        {
            return [];
        }

        XDocument partXml = package.LoadXml(part, cancellationToken);
        OoxMemory.GrowModel(modelStorage, partXml, cancellationToken);
        OoxMarkupCompatibility.WarnMustUnderstandOnce(partXml, part.Name, diagnosticSink, warnedParts);
        IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(part.Name, cancellationToken)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        var numberingCounters = new DocxNumberingCounters();
        var parsedStories = new List<(int Index, DocxRelatedStory Story)>();
        HashSet<string> registeredSeparators = (settings?.Root?
            .Element(WordprocessingNamespace + (storyElementName + "Pr"))?
            .Elements(WordprocessingNamespace + storyElementName)
            .Select(reference => NormalizeDecimalId((string?)reference.Attribute(WordprocessingNamespace + "id")))
            .OfType<string>() ?? []).ToHashSet(StringComparer.Ordinal);
        IEnumerable<(XElement Story, int Index)> sourceStories =
            (partXml.Root?.Elements(WordprocessingNamespace + storyElementName) ?? [])
                .Select((story, index) => (story, index));
        if (noteReferences is not null)
        {
            // List instances continue across referenced notes in document order.
            // Unreferenced/special stories remain readable, but cannot advance
            // counters before visible notes. Sorting leaves cached XML untouched.
            sourceStories = sourceStories.OrderBy(source =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                XElement story = source.Story;
                return (string?)story.Attribute(WordprocessingNamespace + "type") is null or "normal" &&
                    NormalizeDecimalId((string?)story.Attribute(WordprocessingNamespace + "id")) is { } id &&
                    noteReferences.TryGetValue((kind, id), out var reference)
                        ? reference.Order : int.MaxValue;
            });
        }
        foreach (var source in sourceStories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement sourceStory = source.Story;
            XElement storyElement = sourceStory;
            if (kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
                (string?)storyElement.Attribute(WordprocessingNamespace + "type") is null or "normal" &&
                NormalizeDecimalId((string?)storyElement.Attribute(WordprocessingNamespace + "id")) is { } id &&
                noteReferences is not null && noteReferences.TryGetValue((kind, id), out var noteReference) &&
                !string.IsNullOrEmpty(noteReference.Label))
            {
                // Cached XML stays read-only. Detach this note before replacing explicit
                // automatic marks, retaining run properties and source child positions.
                // Numbering comes from the filtered main story, including table cells.
                storyElement = new XElement(sourceStory);
                foreach (XElement mark in storyElement.Descendants(WordprocessingNamespace + (storyElementName + "Ref")).ToArray())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var text = new XElement(WordprocessingNamespace + "t", noteReference.Label);
                    text.AddAnnotation(AutomaticNoteStoryMarker.Instance);
                    mark.ReplaceWith(text);
                }
            }
            XElement effectiveStory = storyElement;
            bool usesDefaultSeparator = kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote &&
                (string?)storyElement.Attribute(WordprocessingNamespace + "type") is "separator" or "continuationSeparator" &&
                NormalizeDecimalId((string?)storyElement.Attribute(WordprocessingNamespace + "id")) is { } separatorId &&
                !registeredSeparators.Contains(separatorId);
            if (usesDefaultSeparator)
            {
                // Word uses a default separator unless settings register this
                // typed special story. Unregistered formatting is not visible.
                string type = (string)storyElement.Attribute(WordprocessingNamespace + "type")!;
                effectiveStory = new XElement(storyElement.Name, storyElement.Attributes(),
                    new XElement(WordprocessingNamespace + "p",
                        new XElement(WordprocessingNamespace + "pPr",
                            new XElement(WordprocessingNamespace + "spacing",
                                new XAttribute(WordprocessingNamespace + "after", "0"),
                                new XAttribute(WordprocessingNamespace + "line", "240"),
                                new XAttribute(WordprocessingNamespace + "lineRule", "auto"))),
                        new XElement(WordprocessingNamespace + "r", new XElement(WordprocessingNamespace + type))));
            }
            DocxRelatedStory story = ReadRelatedStory(kind, part.Name, effectiveStory, styles, numbering, numberingCounters, package, relationships, markupMode, cancellationToken, commentThreadMetadataByParagraphId);
            if (usesDefaultSeparator) story = story with { IsSynthetic = true };
            if (story.BodyElements.Count > 0)
            {
                parsedStories.Add((source.Index, story));
            }
        }

        // Restore the part's source order for inspection and font/resource
        // discovery; only list-counter evaluation follows the references.
        IEnumerable<(int Index, DocxRelatedStory Story)> sourceOrder = parsedStories;
        if (noteReferences is not null) sourceOrder = sourceOrder.OrderBy(source => source.Index);
        var stories = sourceOrder.Select(source => source.Story).ToList();

        // RV06 height-model probes (absent-separator endnotes, Word COM reference
        // edge-endsepheight-absent): Word normalizes separator-less endnote parts with
        // a default separator (absent renders identically to mark: same rule, same
        // placement), while the renderer skips placement entirely. Synthesize the missing
        // separator through the normal story parser from an empty paragraph, so cascade
        // sizing matches author-written empty separators exactly. Footnote parts keep
        // legacy absent behavior (validated corpus). Id 0 is Word's separator convention
        // and separator matching requires a non-null Id.
        if (kind == DocxRelatedStoryKind.Endnote &&
            stories.Count > 0 &&
            stories.All(story => story.Type != DocxRelatedStoryType.Separator))
        {
            XElement syntheticSeparator = new(
                WordprocessingNamespace + storyElementName,
                new XAttribute(WordprocessingNamespace + "type", "separator"),
                new XAttribute(WordprocessingNamespace + "id", "0"),
                new XElement(WordprocessingNamespace + "p"));
            DocxRelatedStory defaultSeparator = ReadRelatedStory(kind, part.Name, syntheticSeparator, styles, numbering, numberingCounters, package, relationships, markupMode, cancellationToken, commentThreadMetadataByParagraphId);
            defaultSeparator = defaultSeparator with { IsSynthetic = true };
            if (defaultSeparator.BodyElements.Count > 0)
            {
                stories.Add(defaultSeparator);
            }
        }

        return stories;
    }

    private sealed class AutomaticNoteStoryMarker
    {
        internal static readonly AutomaticNoteStoryMarker Instance = new();
    }

    private static DocxRelatedStory ReadRelatedStory(
        DocxRelatedStoryKind kind,
        string partName,
        XElement story,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxNumberingCounters numberingCounters,
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
            cancellationToken,
            retainBalloonParagraphMark: kind == DocxRelatedStoryKind.Comment && markupMode == OoxPdfDocxMarkupMode.AllMarkup &&
                story.HasElements && story.Elements().All(element => element.Name == WordprocessingNamespace + "p") &&
                !story.Descendants(WordprocessingNamespace + "drawing").Any() && !story.Descendants(WordprocessingNamespace + "pict").Any());
        // Unqualified stories retain their former font plan:
        // unused marks would otherwise shift surviving PDF resource numbers.
        // Existing one-to-four paragraph fallback resources are left intact.
        if (kind == DocxRelatedStoryKind.Comment && bodyElements.Count >= 5 &&
            !DocxBalloonParagraphPolicy.CanRetainPlainParagraphMarks(bodyElements))
        {
            bodyElements = bodyElements.Select(element => element is DocxParagraphElement paragraph
                ? paragraph with { Paragraph = paragraph.Paragraph with { ParagraphMarkRun = null } }
                : element).ToArray();
        }
        IReadOnlyList<DocxFloatingDrawing> floatingDrawings = ReadFloatingDrawings(story, package, relationships, styles, numbering, markupMode, cancellationToken);
        string? paragraphId = kind == DocxRelatedStoryKind.Comment ? ReadCommentParagraphId(story) : null;
        DocxCommentThreadMetadata? threadMetadata = paragraphId is not null && commentThreadMetadataByParagraphId is not null && commentThreadMetadataByParagraphId.TryGetValue(paragraphId, out DocxCommentThreadMetadata? metadata)
            ? metadata
            : null;
        int sourceEquationCount = 0;
        foreach (XElement equation in story.Descendants(MathNamespace + "oMath"))
        { cancellationToken.ThrowIfCancellationRequested(); sourceEquationCount++; }
        string? storyId = (string?)story.Attribute(WordprocessingNamespace + "id");
        if (kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote)
        {
            storyId = NormalizeDecimalId(storyId);
        }
        return new DocxRelatedStory(
            kind,
            partName,
            storyId,
            bodyElements,
            [],
            [],
            ParseRelatedStoryType((string?)story.Attribute(WordprocessingNamespace + "type")))
        {
            FloatingDrawings = floatingDrawings,
            SourceEquationCount = sourceEquationCount,
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

    private static IReadOnlyDictionary<string, DocxCommentThreadMetadata> ReadCommentThreadMetadata(
        OoxPackage package,
        string documentPartName,
        CancellationToken cancellationToken, IMemoryReservation? modelStorage = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? part = FindRelatedPart(package, documentPartName, CommentsExtendedRelationshipType, cancellationToken);
        if (part is null)
        {
            return new Dictionary<string, DocxCommentThreadMetadata>(StringComparer.OrdinalIgnoreCase);
        }

        XDocument document = package.LoadXml(part, cancellationToken);
        OoxMemory.GrowModel(modelStorage, document, cancellationToken);
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
