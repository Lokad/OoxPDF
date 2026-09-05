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
    private const double WordUntokenedAutoLineSpacingFactor = 1.2d;
    private const double WordSpacingTokenAutoLineSpacingFactor = 1.2d;
    private const double WordDefaultSpacingAfterPoints = 8d;

    private sealed class DocxComplexFieldState
    {
        public DocxComplexFieldState(int sourceRunIndex, int textRunIndex, int textLengthStart, int nestingDepth)
        {
            SourceRunIndex = sourceRunIndex;
            TextRunIndex = textRunIndex;
            TextLengthStart = textLengthStart;
            NestingDepth = nestingDepth;
        }

        public StringBuilder Instruction { get; } = new();
        public int SourceRunIndex { get; }
        public int InstructionSourceRunIndex { get; set; } = -1;
        public int TextRunIndex { get; private set; }
        public int TextLengthStart { get; private set; }
        public int NestingDepth { get; }
        public bool HasSeparate { get; set; }
        public bool InResult { get; set; }
        public bool HasCachedResult { get; set; }
        public bool RendersCachedResult { get; set; }
        public bool PlaceholderEmitted { get; set; }
        public int InstructionRunCount { get; set; }
        public int ResultRunCount { get; set; }

        public void EnsureTextSpan(int textRunIndex, int textLengthStart)
        {
            if (ResultRunCount == 0 && !PlaceholderEmitted && !RendersCachedResult)
            {
                TextRunIndex = textRunIndex;
                TextLengthStart = textLengthStart;
            }
        }
    }

    // Scan-phase projection of complex-field open/close tracking, used only by the
    // unsupported-field pre-scan (HasUnsupportedComplexFields family). The full parse
    // carries richer per-field indices in DocxComplexFieldState; this stays separate
    // because the pre-scan runs before (and independently of) paragraph parsing.
    private sealed record DocxComplexFieldScanState(
        StringBuilder Instruction,
        bool HasSeparate,
        bool InResult,
        bool HasCachedResult);

    private const string MainDocumentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string OfficeDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string StylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    private const string NumberingRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
    private const string HeaderRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
    private const string FooterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
    private const string SettingsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    private const string FontTableRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
    private const string ThemeRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    private const string CommentsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    private const string CommentsExtendedRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
    private const string FootnotesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    private const string EndnotesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    private const string StylesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    private const string NumberingContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    private const string SettingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    private const string FontTableContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
    private const string ThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";
    private const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    private const string CommentsExtendedContentType = "application/vnd.ms-word.commentsExtended+xml";
    private const string FootnotesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    private const string EndnotesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    private const double WordAutomaticParagraphSpacingPoints = 14d;
    public DocxDocument Read(
        OoxPackage package,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        OoxPdfDocxMarkupMode markupMode)
    {
            OoxPart FindDocumentPart(OoxPackage package, CancellationToken cancellationToken)
            {
                cancellationToken.ThrowIfCancellationRequested();
                OoxRelationship? packageRelationship = package.GetRelationships("/", cancellationToken)
                    .FirstOrDefault(r => !r.IsExternal && r.Type == OfficeDocumentRelationshipType && r.ResolvedTarget is not null);
                if (packageRelationship?.ResolvedTarget is not null)
                {
                    OoxPart? relatedPart = package.GetPart(packageRelationship.ResolvedTarget);
                    if (relatedPart is not null)
                    {
                        return relatedPart;
                    }
                }

                OoxPart? contentTypePart = package.Parts.FirstOrDefault(p => p.ContentType == MainDocumentContentType);
                return contentTypePart ?? throw new InvalidDataException("DOCX package does not contain a main document part.");
            }

        cancellationToken.ThrowIfCancellationRequested();
        OoxPart documentPart = FindDocumentPart(package, cancellationToken);
        using Stream stream = documentPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);
        IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(documentPart.Name, cancellationToken)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        IReadOnlyDictionary<string, OoxRelationship> internalRelationships = relationships.Values
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        XDocument? settings = LoadRelatedXmlPart(package, documentPart.Name, SettingsRelationshipType, SettingsContentType, out _, cancellationToken);
        DocxDocumentSettings documentSettings = ReadDocumentSettings(settings);
        OoxPdfDocxMarkupMode revisionFilteringMarkupMode = ResolveRevisionFilteringMarkupMode(markupMode, documentSettings);
        DocxMarkupContext markupContext = DocxMarkupContext.FromMode(markupMode).ApplyDocumentSettings(documentSettings);
        DocxCommentAnchorInventory commentAnchorInventory = ReadCommentAnchorInventory(document, revisionFilteringMarkupMode);
        EmitUnsupportedFeatureDiagnostics(package, document, documentPart.Name, relationships, markupContext, diagnosticSink, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        XElement? sectionProperties = document.Descendants(WordprocessingNamespace + "sectPr").LastOrDefault();
        XElement? pageSize = sectionProperties?.Element(WordprocessingNamespace + "pgSz");
        XElement? pageMargins = sectionProperties?.Element(WordprocessingNamespace + "pgMar");

        DocxFontCatalog fontCatalog = LoadFontCatalog(package, documentPart.Name, cancellationToken);
        DocxStyleSet styles = LoadStyles(package, documentPart.Name, cancellationToken);
        DocxNumberingSet numbering = LoadNumbering(package, documentPart.Name, fontCatalog, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        DocxSectionBreakElement? finalSectionBreak = sectionProperties is null
            ? null
            : ReadSectionBreak(sectionProperties, package, internalRelationships, styles, numbering, settings, revisionFilteringMarkupMode, cancellationToken, null);
        IReadOnlyList<DocxBodyElement> bodyElements = ReadBodyElements(document, styles, numbering, package, relationships, settings, documentSettings, revisionFilteringMarkupMode, cancellationToken);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> headerBodyElementsByType = ReadReferencedHeaderFooterBodyElementsByType(document, package, internalRelationships, styles, numbering, HeaderRelationshipType, "headerReference", revisionFilteringMarkupMode, cancellationToken);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> footerBodyElementsByType = ReadReferencedHeaderFooterBodyElementsByType(document, package, internalRelationships, styles, numbering, FooterRelationshipType, "footerReference", revisionFilteringMarkupMode, cancellationToken);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = ToStaticParagraphsByType(headerBodyElementsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = ToStaticParagraphsByType(footerBodyElementsByType);
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> headerDrawingsByType = ReadReferencedHeaderFooterFloatingDrawingsByType(document, package, internalRelationships, styles, numbering, HeaderRelationshipType, "headerReference", revisionFilteringMarkupMode, cancellationToken);
        IReadOnlyDictionary<string, IReadOnlyList<DocxFloatingDrawing>> footerDrawingsByType = ReadReferencedHeaderFooterFloatingDrawingsByType(document, package, internalRelationships, styles, numbering, FooterRelationshipType, "footerReference", revisionFilteringMarkupMode, cancellationToken);
        IReadOnlyList<DocxParagraph> headers = SelectDefaultHeaderFooterParagraphs(headersByType);
        IReadOnlyList<DocxParagraph> footers = SelectDefaultHeaderFooterParagraphs(footersByType);
        IReadOnlyList<DocxRelatedStory> relatedStories = ReadRelatedStories(package, documentPart.Name, styles, numbering, revisionFilteringMarkupMode, cancellationToken);
        IReadOnlyList<DocxFloatingDrawing> floatingDrawings = ReadFloatingDrawings(document, package, relationships, styles, numbering, revisionFilteringMarkupMode, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        if (pageSize is null)
        {
            return new DocxDocument(
                612d,
                792d,
                72d,
                72d,
                72d,
                72d,
                ReadPageSettings(pageSize, pageMargins, sectionProperties, settings, package, internalRelationships, styles, numbering, revisionFilteringMarkupMode, cancellationToken),
                floatingDrawings,
                headers,
                footers,
                bodyElements,
                [],
                [])
            {
                FontCatalog = fontCatalog,
                StyleCatalog = ToStyleCatalog(styles),
                HeaderParagraphsByType = headersByType,
                FooterParagraphsByType = footersByType,
                HeaderBodyElementsByType = headerBodyElementsByType,
                FooterBodyElementsByType = footerBodyElementsByType,
                HeaderFloatingDrawingsByType = headerDrawingsByType,
                FooterFloatingDrawingsByType = footerDrawingsByType,
                RelatedStories = relatedStories,
                PackageCommentAnchorIds = commentAnchorInventory.PackageAnchorIds,
                HiddenCommentAnchorIds = commentAnchorInventory.HiddenAnchorIds,
                Settings = documentSettings,
                MarkupMode = markupMode,
                FinalSectionBreak = finalSectionBreak
            };
        }

        double width = OoxUnits.TwipsToPoints(OoxXml.ParseRequiredLong(pageSize, WordprocessingNamespace + "w", "DOCX"));
        double height = OoxUnits.TwipsToPoints(OoxXml.ParseRequiredLong(pageSize, WordprocessingNamespace + "h", "DOCX"));
        (width, height) = NormalizePageSize(width, height);
        string? orientation = (string?)pageSize.Attribute(WordprocessingNamespace + "orient");
        if (orientation?.Equals("landscape", StringComparison.OrdinalIgnoreCase) == true && height > width)
        {
            (width, height) = (height, width);
        }

        double left = ReadMargin(pageMargins, WordprocessingNamespace + "left", 72d);
        double right = ReadMargin(pageMargins, WordprocessingNamespace + "right", 72d);
        double top = ReadMargin(pageMargins, WordprocessingNamespace + "top", 72d);
        double bottom = ReadMargin(pageMargins, WordprocessingNamespace + "bottom", 72d);
        return new DocxDocument(
            width,
            height,
            left,
            right,
            top,
            bottom,
            ReadPageSettings(pageSize, pageMargins, sectionProperties, settings, package, internalRelationships, styles, numbering, revisionFilteringMarkupMode, cancellationToken),
            floatingDrawings,
            headers,
            footers,
            bodyElements,
            [],
            [])
        {
            FontCatalog = fontCatalog,
            StyleCatalog = ToStyleCatalog(styles),
            HeaderParagraphsByType = headersByType,
            FooterParagraphsByType = footersByType,
            HeaderBodyElementsByType = headerBodyElementsByType,
            FooterBodyElementsByType = footerBodyElementsByType,
            HeaderFloatingDrawingsByType = headerDrawingsByType,
            FooterFloatingDrawingsByType = footerDrawingsByType,
            RelatedStories = relatedStories,
            PackageCommentAnchorIds = commentAnchorInventory.PackageAnchorIds,
            HiddenCommentAnchorIds = commentAnchorInventory.HiddenAnchorIds,
            Settings = documentSettings,
            MarkupMode = markupMode,
            FinalSectionBreak = finalSectionBreak
        };
    }

    // Single caller; kept static: entry-point stage, not a local candidate.
    private static OoxPdfDocxMarkupMode ResolveRevisionFilteringMarkupMode(
        OoxPdfDocxMarkupMode markupMode,
        DocxDocumentSettings settings)
    {
        DocxRevisionViewSettings revisionView = settings.RevisionViewSettings;
        if (revisionView.ShowMarkup == false || revisionView.ShowInsertionsAndDeletions == false)
        {
            return OoxPdfDocxMarkupMode.Final;
        }

        return markupMode;
    }

    // Single caller; kept static: entry-point stage, not a local candidate.
    private static DocxDocumentSettings ReadDocumentSettings(XDocument? settings)
    {
            DocxRevisionViewSettings ReadRevisionViewSettings(XElement? revisionView)
            {
                return new DocxRevisionViewSettings(
                    (string?)revisionView?.Attribute(WordprocessingNamespace + "markup"),
                    OoxBoolean.ParseOptionalAttribute(revisionView, WordprocessingNamespace + "markup"),
                    (string?)revisionView?.Attribute(WordprocessingNamespace + "comments"),
                    OoxBoolean.ParseOptionalAttribute(revisionView, WordprocessingNamespace + "comments"),
                    (string?)revisionView?.Attribute(WordprocessingNamespace + "insDel"),
                    OoxBoolean.ParseOptionalAttribute(revisionView, WordprocessingNamespace + "insDel"),
                    (string?)revisionView?.Attribute(WordprocessingNamespace + "formatting"),
                    OoxBoolean.ParseOptionalAttribute(revisionView, WordprocessingNamespace + "formatting"),
                    (string?)revisionView?.Attribute(WordprocessingNamespace + "inkAnnotations"),
                    OoxBoolean.ParseOptionalAttribute(revisionView, WordprocessingNamespace + "inkAnnotations"));
            }

        XElement? root = settings?.Root;
        if (root is null)
        {
            return DocxDocumentSettings.Empty;
        }

        XElement? characterSpacingControl = root.Element(WordprocessingNamespace + "characterSpacingControl");
        XElement? defaultTabStop = root.Element(WordprocessingNamespace + "defaultTabStop");
        XElement? revisionView = root.Element(WordprocessingNamespace + "revisionView");
        XElement? trackRevisions = root.Element(WordprocessingNamespace + "trackRevisions");
        XElement? doNotTrackMoves = root.Element(WordprocessingNamespace + "doNotTrackMoves");
        XElement? doNotTrackFormatting = root.Element(WordprocessingNamespace + "doNotTrackFormatting");
        XElement? mirrorMargins = root.Element(WordprocessingNamespace + "mirrorMargins");
        XElement? useFELayout = root.Element(WordprocessingNamespace + "compat")?.Element(WordprocessingNamespace + "useFELayout");
        DocxNoteReferenceSettings footnoteReferenceSettings = ReadNoteReferenceSettings(root.Element(WordprocessingNamespace + "footnotePr"));
        DocxNoteReferenceSettings endnoteReferenceSettings = ReadNoteReferenceSettings(root.Element(WordprocessingNamespace + "endnotePr"));
        IReadOnlyList<DocxCompatSetting> compatSettings = root
            .Element(WordprocessingNamespace + "compat")?
            .Elements(WordprocessingNamespace + "compatSetting")
            .Select(setting => new DocxCompatSetting(
                (string?)setting.Attribute(WordprocessingNamespace + "name"),
                (string?)setting.Attribute(WordprocessingNamespace + "uri"),
                (string?)setting.Attribute(WordprocessingNamespace + "val")))
            .ToArray() ?? [];

        return new DocxDocumentSettings(
            (string?)characterSpacingControl?.Attribute(WordprocessingNamespace + "val"),
            (string?)defaultTabStop?.Attribute(WordprocessingNamespace + "val"),
            ReadTwipsAttribute(defaultTabStop, WordprocessingNamespace + "val"),
            ReadOnOff(useFELayout),
            (string?)useFELayout?.Attribute(WordprocessingNamespace + "val"),
            ReadRevisionViewSettings(revisionView),
            new DocxTrackChangesSettings(
                (string?)trackRevisions?.Attribute(WordprocessingNamespace + "val"),
                ReadOnOff(trackRevisions),
                (string?)doNotTrackMoves?.Attribute(WordprocessingNamespace + "val"),
                ReadOnOff(doNotTrackMoves),
                (string?)doNotTrackFormatting?.Attribute(WordprocessingNamespace + "val"),
                ReadOnOff(doNotTrackFormatting)),
            footnoteReferenceSettings,
            endnoteReferenceSettings,
            (string?)mirrorMargins?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(mirrorMargins),
            compatSettings);
    }


    private static DocxNoteReferenceSettings ReadNoteReferenceSettings(XElement? properties)
    {
        XElement? numberFormat = properties?.Element(WordprocessingNamespace + "numFmt");
        XElement? numberStart = properties?.Element(WordprocessingNamespace + "numStart");
        XElement? numberRestart = properties?.Element(WordprocessingNamespace + "numRestart");
        XElement? position = properties?.Element(WordprocessingNamespace + "pos");
        string? startValue = (string?)numberStart?.Attribute(WordprocessingNamespace + "val");
        return new DocxNoteReferenceSettings(
            (string?)position?.Attribute(WordprocessingNamespace + "val"),
            (string?)numberFormat?.Attribute(WordprocessingNamespace + "val"),
            startValue,
            int.TryParse(startValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int start) ? start : null,
            (string?)numberRestart?.Attribute(WordprocessingNamespace + "val"));
    }

    // Single caller; kept static: entry-point stage, not a local candidate.
    private static DocxCommentAnchorInventory ReadCommentAnchorInventory(XDocument document, OoxPdfDocxMarkupMode markupMode)
    {
            bool IsCommentAnchorElement(XElement element)
            {
                return element.Name == WordprocessingNamespace + "commentReference" ||
                    element.Name == WordprocessingNamespace + "commentRangeStart" ||
                    element.Name == WordprocessingNamespace + "commentRangeEnd";
            }

            string? ReadCommentAnchorId(XElement element)
            {
                return (string?)element.Attribute(WordprocessingNamespace + "id");
            }

        XElement[] anchors = document
            .Descendants()
            .Where(IsCommentAnchorElement)
            .ToArray();
        string[] packageAnchorIds = anchors
            .Select(ReadCommentAnchorId)
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        string[] hiddenAnchorIds = anchors
            .Where(anchor => IsInsideExcludedRevisionContainer(anchor, markupMode))
            .Select(ReadCommentAnchorId)
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        return new DocxCommentAnchorInventory(packageAnchorIds, hiddenAnchorIds);
    }



    private static IReadOnlyList<DocxFloatingDrawing> ReadFloatingDrawings(
        XDocument document,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
            bool IsBodyBlockElementOrRevisionContainer(XElement element)
            {
                return IsBodyBlockElement(element) || IsRevisionContainer(element);
            }

        cancellationToken.ThrowIfCancellationRequested();
        XElement[] paragraphs = document
            .Descendants(WordprocessingNamespace + "p")
            .ToArray();
        XElement[] bodyBlocks = document
            .Root?
            .Element(WordprocessingNamespace + "body")?
            .Elements()
            .Where(IsBodyBlockElementOrRevisionContainer)
            .ToArray() ?? [];
        var drawings = new List<DocxFloatingDrawing>();
        foreach (XElement anchor in document.Descendants(WordprocessingDrawingNamespace + "anchor"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRevisionInfo? revision = FindInheritedRevision(anchor, markupMode);
            if (IsInsideExcludedRevisionContainer(anchor, markupMode))
            {
                continue;
            }

            drawings.Add(ReadFloatingDrawing(
                anchor,
                package,
                relationships,
                styles,
                numbering,
                FindSourceParagraphIndex(anchor, paragraphs),
                FindSourceBlockIndex(anchor, bodyBlocks),
                revision,
                markupMode,
                cancellationToken));
        }

        return drawings;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxFloatingDrawing ReadFloatingDrawing(
        XElement anchor,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        int? sourceParagraphIndex,
        int? sourceBlockIndex,
        DocxRevisionInfo? revision,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        XElement? extent = anchor.Element(WordprocessingDrawingNamespace + "extent");
        XElement? positionH = anchor.Element(WordprocessingDrawingNamespace + "positionH");
        XElement? positionV = anchor.Element(WordprocessingDrawingNamespace + "positionV");
        XElement? wrap = anchor
            .Elements()
            .FirstOrDefault(element =>
            element.Name.Namespace == WordprocessingDrawingNamespace &&
            element.Name.LocalName.StartsWith("wrap", StringComparison.Ordinal));
        string? relationshipId = ReadDrawingImageRelationshipId(anchor);
        DocxInlineImage? image = ReadDrawingImage(anchor, package, relationships, revision);
        IReadOnlyList<DocxBodyElement> textBoxBodyElements = ReadTextBoxBodyElements(
            anchor,
            styles,
            numbering,
            package,
            relationships,
            markupMode,
            cancellationToken);

        return new DocxFloatingDrawing(
            (string?)anchor.Attribute("distT"),
            (string?)anchor.Attribute("distB"),
            (string?)anchor.Attribute("distL"),
            (string?)anchor.Attribute("distR"),
            (string?)anchor.Attribute("simplePos"),
            (string?)anchor.Attribute("relativeHeight"),
            (string?)anchor.Attribute("behindDoc"),
            (string?)anchor.Attribute("locked"),
            (string?)anchor.Attribute("layoutInCell"),
            (string?)anchor.Attribute("allowOverlap"),
            (string?)extent?.Attribute("cx"),
            (string?)extent?.Attribute("cy"),
            (string?)positionH?.Attribute("relativeFrom"),
            (string?)positionH?.Element(WordprocessingDrawingNamespace + "align"),
            (string?)positionH?.Element(WordprocessingDrawingNamespace + "posOffset"),
            (string?)positionV?.Attribute("relativeFrom"),
            (string?)positionV?.Element(WordprocessingDrawingNamespace + "align"),
            (string?)positionV?.Element(WordprocessingDrawingNamespace + "posOffset"),
            DocxFloatingWrapKindExtensions.FromLocalName(wrap?.Name.LocalName),
            (string?)wrap?.Attribute("wrapText"),
            relationshipId,
            image,
            sourceParagraphIndex,
            sourceBlockIndex)
        {
            Revisions = RevisionList(revision),
            TextBoxBodyElements = textBoxBodyElements
        };
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IReadOnlyList<DocxBodyElement> ReadTextBoxBodyElements(
        XElement anchor,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        XElement? textBoxContent = anchor
            .Descendants(WordprocessingNamespace + "txbxContent")
            .FirstOrDefault();
        if (textBoxContent is null)
        {
            return [];
        }

        return ReadRelatedStoryBodyElements(
            textBoxContent.Elements(),
            styles,
            numbering,
            new Dictionary<(string NumId, int Level), int>(),
            package,
            relationships,
            markupMode,
            cancellationToken);
    }

    private static bool IsBodyBlockElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "p" ||
            element.Name == WordprocessingNamespace + "tbl" ||
            element.Name == WordprocessingNamespace + "sectPr";
    }


    // Single caller; kept static: index-finder trio kept together.
    private static int? FindSourceParagraphIndex(XElement element, IReadOnlyList<XElement> paragraphs)
    {
        XElement? paragraph = element.Ancestors(WordprocessingNamespace + "p").FirstOrDefault();
        if (paragraph is null)
        {
            return null;
        }

        for (int index = 0; index < paragraphs.Count; index++)
        {
            if (ReferenceEquals(paragraphs[index], paragraph))
            {
                return index;
            }
        }

        return null;
    }

    // Single caller; kept static: index-finder trio kept together.
    private static int? FindSourceBlockIndex(XElement element, IReadOnlyList<XElement> bodyBlocks)
    {
        XElement? bodyChild = element
            .Ancestors()
            .FirstOrDefault(ancestor => bodyBlocks.Any(block => ReferenceEquals(block, ancestor)));
        if (bodyChild is null)
        {
            return null;
        }

        for (int index = 0; index < bodyBlocks.Count; index++)
        {
            if (ReferenceEquals(bodyBlocks[index], bodyChild))
            {
                return index;
            }
        }

        return null;
    }

    private static DocxPageSettings ReadPageSettings(
        XElement? pageSize,
        XElement? pageMargins,
        XElement? sectionProperties,
        XDocument? settings,
        OoxPackage? package,
        IReadOnlyDictionary<string, OoxRelationship>? relationships,
        DocxStyleSet? styles,
        DocxNumberingSet? numbering,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? titlePage = sectionProperties?.Element(WordprocessingNamespace + "titlePg");
        XElement? evenAndOddHeaders = settings?.Root?.Element(WordprocessingNamespace + "evenAndOddHeaders");
        XElement? docGrid = sectionProperties?.Element(WordprocessingNamespace + "docGrid");
        DocxPageSettings pageSettings = new(
            (string?)pageSize?.Attribute(WordprocessingNamespace + "w"),
            (string?)pageSize?.Attribute(WordprocessingNamespace + "h"),
            (string?)pageSize?.Attribute(WordprocessingNamespace + "orient"),
            (string?)pageMargins?.Attribute(WordprocessingNamespace + "top"),
            (string?)pageMargins?.Attribute(WordprocessingNamespace + "right"),
            (string?)pageMargins?.Attribute(WordprocessingNamespace + "bottom"),
            (string?)pageMargins?.Attribute(WordprocessingNamespace + "left"),
            ReadTwipsAttribute(pageMargins, WordprocessingNamespace + "header"),
            ReadTwipsAttribute(pageMargins, WordprocessingNamespace + "footer"),
            (string?)pageMargins?.Attribute(WordprocessingNamespace + "header"),
            (string?)pageMargins?.Attribute(WordprocessingNamespace + "footer"),
            ReadOnOff(titlePage),
            (string?)titlePage?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(evenAndOddHeaders),
            (string?)evenAndOddHeaders?.Attribute(WordprocessingNamespace + "val"))
        {
            DocGridLinePitchPoints = ReadTwipsAttribute(docGrid, WordprocessingNamespace + "linePitch"),
            DocGridLinePitchValue = (string?)docGrid?.Attribute(WordprocessingNamespace + "linePitch"),
            GutterDistancePoints = ReadTwipsAttribute(pageMargins, WordprocessingNamespace + "gutter"),
            GutterDistanceValue = (string?)pageMargins?.Attribute(WordprocessingNamespace + "gutter"),
            FootnoteReferenceSettings = ReadNoteReferenceSettings(sectionProperties?.Element(WordprocessingNamespace + "footnotePr")),
            EndnoteReferenceSettings = ReadNoteReferenceSettings(sectionProperties?.Element(WordprocessingNamespace + "endnotePr"))
        };
        if (sectionProperties is null || package is null || relationships is null || styles is null || numbering is null)
        {
            return pageSettings;
        }

        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> headerBodyElementsByType = ReadReferencedHeaderFooterBodyElementsByType(
            sectionProperties,
            package,
            relationships,
            styles,
            numbering,
            HeaderRelationshipType,
            "headerReference",
            markupMode,
            cancellationToken);
        IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> footerBodyElementsByType = ReadReferencedHeaderFooterBodyElementsByType(
            sectionProperties,
            package,
            relationships,
            styles,
            numbering,
            FooterRelationshipType,
            "footerReference",
            markupMode,
            cancellationToken);

        return pageSettings with
        {
            HeaderBodyElementsByType = headerBodyElementsByType,
            FooterBodyElementsByType = footerBodyElementsByType,
            HeaderParagraphsByType = ToStaticParagraphsByType(headerBodyElementsByType),
            FooterParagraphsByType = ToStaticParagraphsByType(footerBodyElementsByType),
            HeaderFloatingDrawingsByType = ReadReferencedHeaderFooterFloatingDrawingsByType(
                sectionProperties,
                package,
                relationships,
                styles,
                numbering,
                HeaderRelationshipType,
                "headerReference",
                markupMode,
                cancellationToken),
            FooterFloatingDrawingsByType = ReadReferencedHeaderFooterFloatingDrawingsByType(
                sectionProperties,
                package,
                relationships,
                styles,
                numbering,
                FooterRelationshipType,
                "footerReference",
                markupMode,
                cancellationToken)
        };
    }

    // Single caller; kept static: entry-point stage, not a local candidate.
    private static (double Width, double Height) NormalizePageSize(double width, double height)
    {
        if (Math.Abs(width - 595d) < 0.01d && Math.Abs(height - 842d) < 0.01d)
        {
            return (594.96d, 842.04d);
        }

        return (width, height);
    }

    // Single caller; kept static: entry-point stage, not a local candidate.
    private static void EmitUnsupportedFeatureDiagnostics(
        OoxPackage package,
        XDocument document,
        string partName,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        void EmitUnsupported(string id, string feature)
        {
            Emit(id, feature, "", "Ignored", false);
        }

        void Emit(string id, string feature, string diagnosticPartName, string fallback, bool approximated)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!emitted.Add(id))
            {
                return;
            }

            diagnosticSink(new OoxPdfDiagnostic(
                id,
                OoxPdfSeverity.Warning,
                approximated
                    ? $"DOCX feature '{feature}' was detected and approximated."
                    : $"Unsupported DOCX feature '{feature}' was detected and ignored or approximated.",
                diagnosticPartName.Length == 0 ? partName : diagnosticPartName,
                Feature: feature,
                Fallback: fallback));
        }

        if (document.Descendants(WordprocessingNamespace + "commentRangeStart").Any() ||
            document.Descendants(WordprocessingNamespace + "commentReference").Any())
        {
            string diagnosticPartName = ResolveRelatedPartNameOrDefault(package, partName, CommentsRelationshipType, CommentsContentType, cancellationToken);
            if (markupContext.ApproximatesComments)
            {
                Emit(
                    "DOCX_APPROXIMATED_COMMENTS",
                    "comments",
                    diagnosticPartName,
                    fallback: "Approximated",
                    approximated: true);
            }
            else
            {
                Emit(
                    "DOCX_UNSUPPORTED_COMMENTS",
                    "comments",
                    diagnosticPartName,
                    "Ignored",
                    false);
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedTrackedChanges(document))
        {
            if (markupContext.ApproximatesTrackedChanges)
            {
                Emit(
                    "DOCX_APPROXIMATED_TRACKED_CHANGES",
                    "tracked changes",
                    diagnosticPartName: "",
                    fallback: "Approximated",
                    approximated: true);
            }
            else
            {
                EmitUnsupported("DOCX_UNSUPPORTED_TRACKED_CHANGES", "tracked changes");
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasFormattingTrackedChanges(document))
        {
            Emit(
                markupContext.ApproximatesFormattingRevisions
                    ? "DOCX_APPROXIMATED_FORMATTING_REVISIONS"
                    : "DOCX_UNSUPPORTED_FORMATTING_REVISIONS",
                "formatting revisions",
                diagnosticPartName: "",
                fallback: markupContext.ApproximatesFormattingRevisions ? "Approximated" : "Ignored",
                approximated: markupContext.ApproximatesFormattingRevisions);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedComplexFields(document))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_COMPLEX_FIELD", "complex field");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(MathNamespace + "oMath").Any() ||
            document.Descendants(MathNamespace + "oMathPara").Any())
        {
            EmitUnsupported("DOCX_UNSUPPORTED_EQUATION", "equation");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "object").Any())
        {
            EmitUnsupported("DOCX_UNSUPPORTED_OLE_OBJECT", "OLE object");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingDrawingNamespace + "anchor").Any(anchor => IsUnsupportedFloatingDrawingAnchor(anchor, relationships)))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_FLOATING_DRAWING", "floating drawing");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(ChartNamespace + "chart").Any())
        {
            EmitUnsupported("DOCX_UNSUPPORTED_CHART", "chart drawing payload");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants().Any(element => element.Name.Namespace == DiagramNamespace))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_SMARTART", "SmartArt diagram payload");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedVml(document, relationships))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_VML", "VML drawing payload");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasExternalDrawingImage(document, relationships))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_EXTERNAL_IMAGE", "external drawing image");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "footnoteReference").Any())
        {
            Emit(
                "DOCX_APPROXIMATED_FOOTNOTE",
                "footnote",
                ResolveRelatedPartNameOrDefault(package, partName, FootnotesRelationshipType, FootnotesContentType, cancellationToken),
                "Approximated", false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "endnoteReference").Any())
        {
            Emit(
                "DOCX_APPROXIMATED_ENDNOTE",
                "endnote",
                ResolveRelatedPartNameOrDefault(package, partName, EndnotesRelationshipType, EndnotesContentType, cancellationToken),
                "Approximated", false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedMultiColumnSection(document))
        {
            Emit("DOCX_UNSUPPORTED_MULTI_COLUMN", "multi-column balancing or in-flow section columns", diagnosticPartName: "", fallback: "Explicit break-only column flow is supported", approximated: false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "br").Any(IsUnsupportedColumnBreak))
        {
            Emit("DOCX_UNSUPPORTED_MANUAL_BREAK", "unsupported manual column break container", diagnosticPartName: "", fallback: "Visible body column breaks are supported", approximated: false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "pPr")
            .Elements(WordprocessingNamespace + "sectPr")
            .Any(IsUnsupportedParagraphSectionBreak))
        {
            Emit("DOCX_UNSUPPORTED_SECTION_BREAK", "continuous or unknown paragraph section break", diagnosticPartName: "", fallback: "Partially supported", approximated: false);
        }

        XDocument? styles = LoadRelatedXmlPart(package, partName, StylesRelationshipType, StylesContentType, out string? stylesPartName, cancellationToken);
        if (styles is not null)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (styles.Descendants(WordprocessingNamespace + "style")
                .Elements(WordprocessingNamespace + "pPr")
                .Elements(WordprocessingNamespace + "spacing")
                .Any(HasUnsupportedParagraphSpacingVariant))
            {
                Emit("DOCX_STYLE_PARAGRAPH_SPACING", "style paragraph spacing variant", stylesPartName ?? partName, "Approximated", false);
            }

            if (HasUnsupportedTableBorderStyle(styles))
            {
                Emit("DOCX_TABLE_BORDER_STYLE", "table border style", stylesPartName ?? partName, "Approximated", false);
            }

            if (HasUnsupportedTableCellTextDirection(styles))
            {
                Emit("DOCX_TABLE_TEXT_DIRECTION", "table cell text direction", stylesPartName ?? partName, "Approximated", approximated: true);
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedTableBorderStyle(document))
        {
            Emit("DOCX_TABLE_BORDER_STYLE", "table border style", partName, "Approximated", false);
        }

        if (HasUnsupportedTableCellTextDirection(document))
        {
            Emit("DOCX_TABLE_TEXT_DIRECTION", "table cell text direction", partName, "Approximated", approximated: true);
        }

        XDocument? numbering = LoadRelatedXmlPart(package, partName, NumberingRelationshipType, NumberingContentType, out string? numberingPartName, cancellationToken);
        if (numbering is not null &&
            numbering.Descendants(WordprocessingNamespace + "lvl")
                .Any(HasUnsupportedNumberingIndent))
        {
            Emit("DOCX_NUMBERING_INDENT", "numbering level indent", numberingPartName ?? partName, "Approximated", false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "ind").Any(HasCharacterUnitIndent) ||
            styles?.Descendants(WordprocessingNamespace + "ind").Any(HasCharacterUnitIndent) == true ||
            numbering?.Descendants(WordprocessingNamespace + "ind").Any(HasCharacterUnitIndent) == true)
        {
            EmitUnsupported("DOCX_UNSUPPORTED_CHARACTER_UNIT_INDENT", "character-unit paragraph indent");
        }

        if (package.Parts.Any(p => p.Name.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase) ||
            p.ContentType.Contains("vbaProject", StringComparison.OrdinalIgnoreCase)))
        {
            EmitUnsupported("DOCX_UNSUPPORTED_MACRO", "macro");
        }
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool IsUnsupportedFloatingDrawingAnchor(XElement anchor, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement? extent = anchor.Element(WordprocessingDrawingNamespace + "extent");
        if (!HasPositiveEmuAttribute(extent, "cx") || !HasPositiveEmuAttribute(extent, "cy"))
        {
            return true;
        }

        if (!HasSupportedFloatingDrawingImage(anchor, relationships) &&
            !HasTextBoxContent(anchor))
        {
            return true;
        }

        return !IsSupportedHorizontalAnchorPosition(anchor.Element(WordprocessingDrawingNamespace + "positionH")) ||
            !IsSupportedVerticalAnchorPosition(anchor.Element(WordprocessingDrawingNamespace + "positionV")) ||
            !IsSupportedFloatingWrap(anchor);
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool HasSupportedFloatingDrawingImage(XElement anchor, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        string? relationshipId = ReadDrawingImageRelationshipId(anchor);
        return relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) &&
            !relationship.IsExternal &&
            relationship.ResolvedTarget is not null;
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool HasExternalDrawingImage(XDocument document, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        foreach (XElement blip in document.Descendants(DrawingNamespace + "blip"))
        {
            string? relationshipId =
                (string?)blip.Attribute(RelationshipsNamespace + "link") ??
                (string?)blip.Attribute(RelationshipsNamespace + "embed");
            if (relationshipId is not null &&
                relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) &&
                relationship.IsExternal)
            {
                return true;
            }
        }

        return false;
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool HasUnsupportedVml(XDocument document, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        foreach (XElement element in document.Descendants().Where(element => element.Name.Namespace == VmlNamespace))
        {
            if (!IsSupportedVmlElement(element, relationships))
            {
                return true;
            }
        }

        return false;
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsSupportedVmlElement(XElement element, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        return IsSupportedVmlImageElement(element, relationships) ||
            IsSupportedVmlTextBoxElement(element) ||
            IsInertVmlDefinitionElement(element);
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsSupportedVmlImageElement(XElement element, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        if (element.Name == VmlNamespace + "shape")
        {
            return IsSupportedVmlImageShape(element, relationships);
        }

        if (element.Name == VmlNamespace + "imagedata")
        {
            return element.Ancestors(VmlNamespace + "shape").Any(shape => IsSupportedVmlImageShape(shape, relationships));
        }

        return false;
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsSupportedVmlTextBoxElement(XElement element)
    {
        if (element.Name == VmlNamespace + "shape")
        {
            return IsSupportedVmlTextBoxShape(element);
        }

        if (element.Name == VmlNamespace + "textbox")
        {
            return element.Ancestors(VmlNamespace + "shape").Any(IsSupportedVmlTextBoxShape);
        }

        return false;
    }

    private static bool IsSupportedVmlTextBoxShape(XElement shape)
    {
        return shape.Descendants(WordprocessingNamespace + "txbxContent").Any() &&
            !shape.Descendants(WordprocessingNamespace + "tbl").Any() &&
            shape
                .Descendants()
                .Where(element => element.Name.Namespace == VmlNamespace)
                .All(element => element.Name == VmlNamespace + "textbox");
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsInertVmlDefinitionElement(XElement element)
    {
        return element.Name == VmlNamespace + "shapetype" ||
            element.Name == VmlNamespace + "stroke" ||
            element.Name == VmlNamespace + "path";
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsSupportedVmlImageShape(XElement shape, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        return TryReadVmlImageShape(
                shape,
                relationships,
                out _,
                out _,
                out _) &&
            shape
                .Descendants()
                .Where(element => element.Name.Namespace == VmlNamespace)
                .All(element => element.Name == VmlNamespace + "imagedata");
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool HasTextBoxContent(XElement anchor)
    {
        return anchor
            .Descendants(WordprocessingNamespace + "txbxContent")
            .Any(content => content.Elements().Any());
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsSupportedHorizontalAnchorPosition(XElement? position)
    {
        return IsSupportedAnchorPosition(
            position,
            static relativeFrom => relativeFrom is "page" or "margin" or "column",
            static align => align is "left" or "center" or "right");
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsSupportedVerticalAnchorPosition(XElement? position)
    {
        return IsSupportedAnchorPosition(
            position,
            static relativeFrom => relativeFrom is "page" or "margin" or "paragraph",
            static align => align is "top" or "center" or "bottom");
    }

    private static bool IsSupportedAnchorPosition(
        XElement? position,
        Func<string, bool> supportsRelativeFrom,
        Func<string, bool> supportsAlign)
    {
        if (position is null)
        {
            return false;
        }

        string? relativeFrom = ((string?)position.Attribute("relativeFrom"))?.ToLowerInvariant();
        if (relativeFrom is null || !supportsRelativeFrom(relativeFrom))
        {
            return false;
        }

        string? align = ((string?)position.Element(WordprocessingDrawingNamespace + "align"))?.ToLowerInvariant();
        if (align is not null)
        {
            return supportsAlign(align);
        }

        return long.TryParse(
            (string?)position.Element(WordprocessingDrawingNamespace + "posOffset"),
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out _);
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsSupportedFloatingWrap(XElement anchor)
    {
        XElement? wrap = anchor.Elements()
            .FirstOrDefault(element =>
                element.Name.Namespace == WordprocessingDrawingNamespace &&
                element.Name.LocalName.StartsWith("wrap", StringComparison.Ordinal));
        return wrap?.Name.LocalName is null ||
            wrap.Name.LocalName.Equals("wrapNone", StringComparison.OrdinalIgnoreCase) ||
            wrap.Name.LocalName.Equals("wrapSquare", StringComparison.OrdinalIgnoreCase) ||
            wrap.Name.LocalName.Equals("wrapTight", StringComparison.OrdinalIgnoreCase) ||
            wrap.Name.LocalName.Equals("wrapThrough", StringComparison.OrdinalIgnoreCase) ||
            wrap.Name.LocalName.Equals("wrapTopAndBottom", StringComparison.OrdinalIgnoreCase);
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool HasPositiveEmuAttribute(XElement? element, string attributeName)
    {
        return long.TryParse(
            (string?)element?.Attribute(attributeName),
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out long value) &&
            value > 0;
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool IsUnsupportedParagraphSectionBreak(XElement sectionProperties)
    {
        string? typeValue = (string?)sectionProperties
            .Element(WordprocessingNamespace + "type")
            ?.Attribute(WordprocessingNamespace + "val");
        return typeValue is not null &&
            !typeValue.Equals("nextPage", StringComparison.OrdinalIgnoreCase) &&
            !typeValue.Equals("oddPage", StringComparison.OrdinalIgnoreCase) &&
            !typeValue.Equals("evenPage", StringComparison.OrdinalIgnoreCase);
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool HasUnsupportedMultiColumnSection(XDocument document)
    {
            bool IsMultiColumnDeclaration(XElement columns)
            {
                return columns.Attribute(WordprocessingNamespace + "num") is { } value &&
                    int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int columnCount) &&
                    columnCount > 1;
            }

        XElement[] multiColumnDeclarations = document
            .Descendants(WordprocessingNamespace + "cols")
            .Where(IsMultiColumnDeclaration)
            .ToArray();
        if (multiColumnDeclarations.Length == 0)
        {
            return false;
        }

        if (document.Descendants(WordprocessingNamespace + "br").Any(IsUnsupportedColumnBreak))
        {
            return true;
        }

        return !IsSupportedExplicitFinalSectionColumnFlow(document, multiColumnDeclarations);
    }


    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsSupportedExplicitFinalSectionColumnFlow(XDocument document, IReadOnlyList<XElement> multiColumnDeclarations)
    {
        XElement? body = document.Root?.Element(WordprocessingNamespace + "body");
        XElement? finalSectionProperties = body?.Element(WordprocessingNamespace + "sectPr");
        XElement? finalColumns = finalSectionProperties?.Element(WordprocessingNamespace + "cols");
        if (body is null ||
            finalColumns is null ||
            multiColumnDeclarations.Count != 1 ||
            !ReferenceEquals(finalColumns, multiColumnDeclarations[0]))
        {
            return false;
        }

        int columnCount = (int)OoxXml.ParseRequiredLong(finalColumns, WordprocessingNamespace + "num", "column count");
        int supportedColumnBreaks = body
            .Elements(WordprocessingNamespace + "p")
            .SelectMany(paragraph => paragraph.Descendants(WordprocessingNamespace + "br"))
            .Count(breakElement => IsColumnBreak(breakElement) && !IsUnsupportedColumnBreak(breakElement));
        return supportedColumnBreaks == columnCount - 1;
    }

    private static bool IsUnsupportedColumnBreak(XElement breakElement)
    {
        if (!IsColumnBreak(breakElement))
        {
            return false;
        }

        XElement? paragraph = breakElement.Ancestors(WordprocessingNamespace + "p").FirstOrDefault();
        if (paragraph is null)
        {
            return true;
        }

        bool isBodyParagraph = paragraph.Parent?.Name == WordprocessingNamespace + "body";
        bool isTableCellParagraph = paragraph.Ancestors(WordprocessingNamespace + "tc").Any();
        if (!isBodyParagraph && !isTableCellParagraph)
        {
            return true;
        }

        XElement? visibleOwner = breakElement
            .AncestorsAndSelf()
            .TakeWhile(element => element != paragraph)
            .FirstOrDefault(element => element.Name == WordprocessingNamespace + "r" || IsVisibleRunContainer(element));
        return visibleOwner is null;
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool HasUnsupportedTableBorderStyle(XDocument document)
    {
        return document
            .Descendants()
            .Any(element => element.Name.Namespace == WordprocessingNamespace &&
                IsTableBorderContainer(element.Parent) &&
                IsUnsupportedVisibleBorderStyle((string?)element.Attribute(WordprocessingNamespace + "val")));
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool HasUnsupportedTableCellTextDirection(XDocument document)
    {
        return document
            .Descendants(WordprocessingNamespace + "textDirection")
            .Any(IsUnsupportedTableCellTextDirection);
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsUnsupportedTableCellTextDirection(XElement textDirection)
    {
        string? value = (string?)textDirection.Attribute(WordprocessingNamespace + "val");
        return value is not null && !value.Equals("lrTb", StringComparison.OrdinalIgnoreCase);
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool HasUnsupportedTrackedChanges(XDocument document)
    {
            bool IsSupportedVisibleInlineContainerChild(XElement element)
            {
                return element.Name == WordprocessingNamespace + "r" ||
                    element.Name == WordprocessingNamespace + "bookmarkStart" ||
                    element.Name == WordprocessingNamespace + "sdt" ||
                    IsVisibleInlineContainer(element);
            }

            bool IsSupportedInlineContainerParent(XElement? element)
            {
                return element?.Name == WordprocessingNamespace + "p" ||
                    element?.Name == WordprocessingNamespace + "sdtContent" ||
                    IsVisibleInlineContainer(element);
            }

        XName[] unsupportedTrackChangeContainers =
        [
            WordprocessingNamespace + "moveFrom",
            WordprocessingNamespace + "moveFromRangeStart",
            WordprocessingNamespace + "moveFromRangeEnd",
            WordprocessingNamespace + "moveToRangeStart",
            WordprocessingNamespace + "moveToRangeEnd"
        ];
        if (unsupportedTrackChangeContainers.Any(name => document.Descendants(name).Any()))
        {
            return true;
        }

        XName[] visibleTrackedChangeContainers =
        [
            WordprocessingNamespace + "ins",
            WordprocessingNamespace + "moveTo"
        ];
        return visibleTrackedChangeContainers.Any(containerName =>
            document.Descendants(containerName).Any(container =>
                !IsSupportedInlineContainerParent(container.Parent) ||
                container.Elements().Any(child => !IsSupportedVisibleInlineContainerChild(child))));
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool HasFormattingTrackedChanges(XDocument document)
    {
        XName[] formattingRevisionElements =
        [
            WordprocessingNamespace + "rPrChange",
            WordprocessingNamespace + "pPrChange",
            WordprocessingNamespace + "tblPrChange",
            WordprocessingNamespace + "trPrChange",
            WordprocessingNamespace + "tcPrChange",
            WordprocessingNamespace + "sectPrChange"
        ];
        return formattingRevisionElements.Any(name => document.Descendants(name).Any());
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static bool HasUnsupportedComplexFields(XDocument document)
    {
            bool IsSupportedComplexFieldRunChild(XElement element)
            {
                XElement? run = element.Parent;
                XElement? container = run?.Parent;
                return run?.Name == WordprocessingNamespace + "r" &&
                    (container?.Name == WordprocessingNamespace + "p" || IsComplexFieldInlineContainer(container));
            }

            bool IsUnsupportedUnclosedComplexField(
                DocxComplexFieldScanState field)
            {
                return ResolveFieldPlaceholder(field.Instruction.ToString()) is null &&
                    (!field.HasSeparate || !field.HasCachedResult);
            }

        var fields = new List<DocxComplexFieldScanState>();
        foreach (XElement paragraph in document.Descendants(WordprocessingNamespace + "p"))
        {
            if (HasUnsupportedComplexFieldInInlineChildren(paragraph, fields))
            {
                return true;
            }
        }

        if (fields.Any(IsUnsupportedUnclosedComplexField))
        {
            return true;
        }

        return document.Descendants(WordprocessingNamespace + "fldChar").Any(fieldChar => !IsSupportedComplexFieldRunChild(fieldChar)) ||
            document.Descendants(WordprocessingNamespace + "instrText").Any(instruction => !IsSupportedComplexFieldRunChild(instruction));
    }

    private static bool IsVisibleInlineContainer(XElement? element)
    {
        return element is not null &&
            (element.Name == WordprocessingNamespace + "hyperlink" ||
            element.Name == WordprocessingNamespace + "fldSimple" ||
            element.Name == WordprocessingNamespace + "ins" ||
            element.Name == WordprocessingNamespace + "moveTo");
    }

    private static bool IsComplexFieldInlineContainer(XElement? element)
    {
        return IsVisibleInlineContainer(element) ||
            element?.Name == WordprocessingNamespace + "del" ||
            element?.Name == WordprocessingNamespace + "moveFrom" ||
            element?.Name == WordprocessingNamespace + "sdtContent";
    }





    private static bool HasUnsupportedComplexFieldInInlineChildren(
        XElement container,
        List<DocxComplexFieldScanState> fields)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == WordprocessingNamespace + "r")
            {
                foreach (XElement runChild in child.Elements())
                {
                    if (ProcessComplexFieldRunChild(runChild, fields))
                    {
                        return true;
                    }
                }

                continue;
            }

            if (child.Name == WordprocessingNamespace + "sdt")
            {
                foreach (XElement content in child.Elements(WordprocessingNamespace + "sdtContent"))
                {
                    if (HasUnsupportedComplexFieldInInlineChildren(content, fields))
                    {
                        return true;
                    }
                }

                continue;
            }

            if (IsComplexFieldInlineContainer(child) &&
                HasUnsupportedComplexFieldInInlineChildren(child, fields))
            {
                return true;
            }
        }

        return false;
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool ProcessComplexFieldRunChild(
        XElement child,
        List<DocxComplexFieldScanState> fields)
    {
        if (child.Name == WordprocessingNamespace + "fldChar")
        {
            string? fieldCharType = (string?)child.Attribute(WordprocessingNamespace + "fldCharType");
            if (string.Equals(fieldCharType, "begin", StringComparison.OrdinalIgnoreCase))
            {
                if (fields.Count != 0 && !fields[^1].InResult)
                {
                    return true;
                }

                fields.Add(new DocxComplexFieldScanState(new StringBuilder(), HasSeparate: false, InResult: false, HasCachedResult: false));
                return false;
            }

            if (string.Equals(fieldCharType, "separate", StringComparison.OrdinalIgnoreCase))
            {
                if (fields.Count == 0)
                {
                    return true;
                }

                (StringBuilder instruction, _, _, bool hasCachedResult) = fields[^1];
                fields[^1] = new DocxComplexFieldScanState(instruction, HasSeparate: true, InResult: true, hasCachedResult);
                return false;
            }

            if (string.Equals(fieldCharType, "end", StringComparison.OrdinalIgnoreCase))
            {
                if (fields.Count == 0)
                {
                    return true;
                }

                (StringBuilder instruction, bool hasSeparate, _, bool hasCachedResult) = fields[^1];
                if (ResolveFieldPlaceholder(instruction.ToString()) is null && (!hasSeparate || !hasCachedResult))
                {
                    return true;
                }

                fields.RemoveAt(fields.Count - 1);
                return false;
            }

            return true;
        }

        if (child.Name == WordprocessingNamespace + "instrText")
        {
            if (fields.Count == 0 || fields[^1].InResult)
            {
                return true;
            }

            fields[^1].Instruction.Append((string?)child);
            return false;
        }

        if (ReadRunTextChild(child).Length != 0)
        {
            for (int fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                (StringBuilder instruction, bool hasSeparate, bool inResult, bool _) = fields[fieldIndex];
                if (inResult)
                {
                    fields[fieldIndex] = new DocxComplexFieldScanState(instruction, hasSeparate, inResult, HasCachedResult: true);
                }
            }
        }

        return false;
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsTableBorderContainer(XElement? element)
    {
        return element is not null &&
            element.Name.Namespace == WordprocessingNamespace &&
            (element.Name.LocalName.Equals("tblBorders", StringComparison.OrdinalIgnoreCase) ||
                element.Name.LocalName.Equals("tcBorders", StringComparison.OrdinalIgnoreCase));
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsUnsupportedVisibleBorderStyle(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) ||
            value.Equals("nil", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return !IsSupportedVisibleBorderStyle(value);
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool IsSupportedVisibleBorderStyle(string value)
    {
        return value.Equals("single", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thick", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("double", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dotted", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dashed", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dashSmallGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dashDotStroked", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dotDash", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("dotDotDash", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("triple", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickSmallGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thickThinSmallGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickThinSmallGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickMediumGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thickThinMediumGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickThinMediumGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickLargeGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thickThinLargeGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("thinThickThinLargeGap", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("threeDEmboss", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("threeDEngrave", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("wave", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("doubleWave", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("outset", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("inset", StringComparison.OrdinalIgnoreCase);
    }

    private static XDocument? LoadRelatedXmlPart(OoxPackage package, string documentPartName, string relationshipType, string contentType, out string? relatedPartName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? part = FindRelatedPart(package, documentPartName, relationshipType, contentType, cancellationToken);
        relatedPartName = part?.Name;
        if (part is null)
        {
            return null;
        }

        using Stream stream = part.OpenRead();
        return SafeXml.Load(stream, cancellationToken);
    }

    // Single caller; kept static: diagnostic probe in the unsupported-feature checklist.
    private static string ResolveRelatedPartNameOrDefault(OoxPackage package, string documentPartName, string relationshipType, string contentType, CancellationToken cancellationToken)
    {
        return FindRelatedPart(package, documentPartName, relationshipType, contentType, cancellationToken)?.Name ?? documentPartName;
    }

    private static OoxPart? FindRelatedPart(OoxPackage package, string documentPartName, string relationshipType, string contentType, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? relationship = package.GetRelationships(documentPartName, cancellationToken)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null
            ? package.Parts.FirstOrDefault(p => p.ContentType == contentType)
            : package.GetPart(relationship.ResolvedTarget);
    }

    // Single caller; kept static: entry-point stage, not a local candidate.
    private static DocxFontCatalog LoadFontCatalog(OoxPackage package, string documentPartName, CancellationToken cancellationToken)
    {
        XDocument? fontTable = LoadRelatedXmlPart(package, documentPartName, FontTableRelationshipType, FontTableContentType, out _, cancellationToken);
        XDocument? theme = LoadRelatedXmlPart(package, documentPartName, ThemeRelationshipType, ThemeContentType, out _, cancellationToken);
        return new DocxFontCatalog(ReadFontTableEntries(fontTable), ReadThemeFonts(theme));
    }

    // Single caller; kept static: font-catalog assembly pair kept together.
    private static IReadOnlyList<DocxFontTableEntry> ReadFontTableEntries(XDocument? fontTable)
    {
        if (fontTable is null)
        {
            return [];
        }

        return fontTable
            .Descendants(WordprocessingNamespace + "font")
            .Select(font => new DocxFontTableEntry(
                (string?)font.Attribute(WordprocessingNamespace + "name") ?? string.Empty,
                (string?)font.Element(WordprocessingNamespace + "altName")?.Attribute(WordprocessingNamespace + "val"),
                (string?)font.Element(WordprocessingNamespace + "family")?.Attribute(WordprocessingNamespace + "val"),
                (string?)font.Element(WordprocessingNamespace + "pitch")?.Attribute(WordprocessingNamespace + "val"),
                (string?)font.Element(WordprocessingNamespace + "panose1")?.Attribute(WordprocessingNamespace + "val"),
                (string?)font.Element(WordprocessingNamespace + "charset")?.Attribute(WordprocessingNamespace + "val")))
            .Where(entry => entry.Name.Length != 0)
            .ToArray();
    }

    // Single caller; kept static: font-catalog assembly pair kept together.
    private static DocxThemeFonts ReadThemeFonts(XDocument? theme)
    {
        if (theme is null)
        {
            return DocxThemeFonts.Empty;
        }

        XElement? fontScheme = theme
            .Descendants(DrawingNamespace + "fontScheme")
            .FirstOrDefault();
        return new DocxThemeFonts(
            MajorLatinTypeface: (string?)fontScheme
                ?.Element(DrawingNamespace + "majorFont")
                ?.Element(DrawingNamespace + "latin")
                ?.Attribute("typeface"),
            MinorLatinTypeface: (string?)fontScheme
                ?.Element(DrawingNamespace + "minorFont")
                ?.Element(DrawingNamespace + "latin")
                ?.Attribute("typeface"),
            MajorComplexScriptTypeface: (string?)fontScheme
                ?.Element(DrawingNamespace + "majorFont")
                ?.Element(DrawingNamespace + "cs")
                ?.Attribute("typeface"),
            MinorComplexScriptTypeface: (string?)fontScheme
                ?.Element(DrawingNamespace + "minorFont")
                ?.Element(DrawingNamespace + "cs")
                ?.Attribute("typeface"),
            MajorEastAsiaTypeface: (string?)fontScheme
                ?.Element(DrawingNamespace + "majorFont")
                ?.Element(DrawingNamespace + "ea")
                ?.Attribute("typeface"),
            MinorEastAsiaTypeface: (string?)fontScheme
                ?.Element(DrawingNamespace + "minorFont")
                ?.Element(DrawingNamespace + "ea")
                ?.Attribute("typeface"));
    }

    private static DocxParagraph? ReadParagraph(
        XElement paragraph,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxTableCellStyle? tableCellStyle,
        Dictionary<DocxRelatedStoryKind, int>? inlineReferenceCounters,
        DocxDocumentSettings? documentSettings,
        DocxRevisionInfo? inheritedRevision,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
        string? paragraphStyleId = ReadParagraphStyleId(paragraphProperties);
        DocxResolvedParagraphProperties resolvedParagraph = ResolveParagraphProperties(
            paragraphProperties,
            paragraphStyleId,
            styles,
            tableCellStyle?.Paragraph);
        DocxParagraphStyleResolution styleResolution = CreateParagraphStyleResolution(
            paragraphProperties,
            paragraphStyleId,
            styles,
            tableCellStyle?.Paragraph);
        var runs = new List<DocxTextRun>();
        var images = new List<DocxInlineImage>();
        var inlineReferences = new List<DocxInlineReference>();
        var commentRanges = new List<DocxCommentRange>();
        var openCommentRanges = new List<DocxCommentRangeStart>();
        var revisionRanges = new List<DocxRevisionRange>();
        var openRevisionRanges = new List<DocxRevisionRangeStart>();
        var fieldReferences = new List<DocxFieldReference>();
        var hyperlinkSpans = new List<DocxHyperlinkSpan>();
        var bookmarkAnchors = new List<DocxBookmarkAnchor>();
        var paragraphRevisions = new List<DocxRevisionInfo>();
        AddRevision(paragraphRevisions, inheritedRevision);
        AddRevisions(paragraphRevisions, ReadPropertyChangeRevisions(paragraphProperties));
        XElement? paragraphMarkRunProperties = paragraphProperties?.Element(WordprocessingNamespace + "rPr");
        IReadOnlyList<DocxRevisionInfo> paragraphMarkRevisions = ReadPropertyChangeRevisions(paragraphMarkRunProperties);
        bool hasDeletedParagraphMark = paragraphMarkRevisions.Any(revision => revision.Kind == DocxRevisionKind.Deletion && revision.SourceElement == "del");
        AddRevisions(paragraphRevisions, paragraphMarkRevisions);
        bool pageInstructionSeen = false;
        var complexFieldStack = new List<DocxComplexFieldState>();
        int sourceRunIndex = 0;
        foreach (XElement child in paragraph.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (child.Name == WordprocessingNamespace + "r")
            {
                AddParagraphRun(child, ref pageInstructionSeen, inheritedRevision);
            }
            else if (child.Name == WordprocessingNamespace + "fldSimple")
            {
                AddSimpleField(child, inheritedRevision);
            }
            else if (IsRevisionContainer(child))
            {
                AddRevisionRunContainer(child, inheritedRevision);
            }
            else if (child.Name == WordprocessingNamespace + "hyperlink")
            {
                AddHyperlinkContainer(child, inheritedRevision);
            }
            else if (child.Name == WordprocessingNamespace + "sdt")
            {
                AddContentControl(child, inheritedRevision);
            }
            else if (child.Name == WordprocessingNamespace + "bookmarkStart")
            {
                AddBookmarkAnchor(child);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeStart")
            {
                AddCommentRangeStart(child);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeEnd")
            {
                AddCommentRangeEnd(child);
            }
            else if (IsRevisionMarkerElement(child))
            {
                AddRevisionMarker(child);
            }
        }

        void AddInlineContainerChild(XElement child, DocxRevisionInfo? revision)
        {
            if (child.Name == WordprocessingNamespace + "r")
            {
                AddParagraphRun(child, ref pageInstructionSeen, revision);
            }
            else if (child.Name == WordprocessingNamespace + "bookmarkStart")
            {
                AddBookmarkAnchor(child);
            }
            else if (child.Name == WordprocessingNamespace + "fldSimple")
            {
                AddSimpleField(child, revision);
            }
            else if (IsRevisionContainer(child))
            {
                AddRevisionRunContainer(child, inheritedRevision: revision);
            }
            else if (child.Name == WordprocessingNamespace + "hyperlink")
            {
                AddHyperlinkContainer(child, revision);
            }
            else if (child.Name == WordprocessingNamespace + "sdt")
            {
                AddContentControl(child, revision);
            }
            else if (child.Name == WordprocessingNamespace + "sdtContent")
            {
                AddVisibleRunContainer(child, revision);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeStart")
            {
                AddCommentRangeStart(child);
            }
            else if (child.Name == WordprocessingNamespace + "commentRangeEnd")
            {
                AddCommentRangeEnd(child);
            }
            else if (IsRevisionMarkerElement(child))
            {
                AddRevisionMarker(child);
            }
        }

        void AddVisibleRunContainer(XElement container, DocxRevisionInfo? revision)
        {
            foreach (XElement containerChild in container.Elements())
            {
                AddInlineContainerChild(containerChild, revision);
            }
        }

        void AddContentControl(XElement contentControl, DocxRevisionInfo? revision)
        {
            foreach (XElement content in contentControl.Elements(WordprocessingNamespace + "sdtContent"))
            {
                AddVisibleRunContainer(content, revision);
            }
        }

        void AddRevisionRunContainer(XElement container, DocxRevisionInfo? inheritedRevision)
        {
            if (!IsIncludedRevisionContainer(container, markupMode))
            {
                return;
            }

            DocxRevisionInfo? revision = CreateRevisionInfo(container) ?? inheritedRevision;
            if (revision is not null)
            {
                paragraphRevisions.Add(revision);
            }

            AddVisibleRunContainer(container, revision);
        }

        FinalizeOpenComplexFields();

        if (runs.Count == 0 && images.Count == 0)
        {
            DocxResolvedRunProperties paragraphMarkRun = ResolveRunProperties(
                paragraphMarkRunProperties,
                paragraphStyleId,
                characterStyleId: null,
                styles,
                tableCellStyle?.Run);
            DocxRunStyleResolution paragraphMarkStyleResolution = CreateRunStyleResolution(
                paragraphMarkRunProperties,
                paragraphStyleId,
                characterStyleId: null,
                styles,
                tableCellStyle?.Run);
            AddResolvedTextRun(
                runs,
                string.Empty,
                paragraphMarkRun,
                paragraphMarkStyleResolution,
                complexScript: false,
                sourceRunIndex: -1,
                sourceTextOffsetInRun: 0,
                revision: inheritedRevision,
                revisions: MergeRevisionLists(inheritedRevision, paragraphMarkRevisions));
        }

        foreach (DocxRevisionRangeStart openRange in openRevisionRanges)
        {
            revisionRanges.Add(new DocxRevisionRange(
                openRange.Kind,
                openRange.Id,
                openRange.Name,
                openRange.Author,
                openRange.Date,
                openRange.SourceRunIndex,
                openRange.TextOffset,
                EndSourceRunIndex: null,
                EndTextOffset: null));
        }

        double paragraphFontSize = runs.Count == 0 ? DocxDefaults.FontSizePoints : runs.Max(run => run.FontSize);
        double lineSpacingFactor = resolvedParagraph.LineSpacingFactor ?? ResolveDefaultAutoLineSpacingFactor(resolvedParagraph);
        double paragraphLineHeight = resolvedParagraph.LineSpacingPoints ?? paragraphFontSize * lineSpacingFactor;

        return new DocxParagraph(
            runs,
            images,
            paragraphStyleId,
            resolvedParagraph.Alignment ?? DocxTextAlignment.Left,
            resolvedParagraph.AlignmentValue,
            ResolveSpacingBeforePoints(resolvedParagraph, paragraphLineHeight),
            ResolveSpacingAfterPoints(resolvedParagraph, paragraphLineHeight),
            lineSpacingFactor,
            resolvedParagraph.LineSpacingPoints,
            resolvedParagraph.Spacing,
            resolvedParagraph.KeepRules,
            CreateListLabel(paragraphProperties, numbering, numberingCounters))
        {
            Indent = resolvedParagraph.Indent,
            TabStops = resolvedParagraph.TabStops,
            SnapToGrid = resolvedParagraph.SnapToGrid,
            SnapToGridValue = resolvedParagraph.SnapToGridValue,
            WordWrap = resolvedParagraph.WordWrap,
            WordWrapValue = resolvedParagraph.WordWrapValue,
            StyleResolution = styleResolution,
            InlineReferences = inlineReferences,
            CommentRanges = commentRanges,
            RevisionRanges = revisionRanges,
            FieldReferences = fieldReferences,
            Hyperlinks = hyperlinkSpans,
            BookmarkAnchors = bookmarkAnchors,
            Revisions = paragraphRevisions,
            HasDeletedParagraphMark = hasDeletedParagraphMark
        };

        void AddBookmarkAnchor(XElement bookmarkStart)
        {
            bookmarkAnchors.Add(new DocxBookmarkAnchor(
                (string?)bookmarkStart.Attribute(WordprocessingNamespace + "id"),
                (string?)bookmarkStart.Attribute(WordprocessingNamespace + "name"),
                sourceRunIndex,
                runs.Count,
                runs.Sum(run => run.Text.Length)));
        }

        void AddCommentRangeStart(XElement rangeStart)
        {
            openCommentRanges.Add(new DocxCommentRangeStart(
                (string?)rangeStart.Attribute(WordprocessingNamespace + "id"),
                sourceRunIndex,
                runs.Sum(run => run.Text.Length)));
        }

        void AddCommentRangeEnd(XElement rangeEnd)
        {
            string? id = (string?)rangeEnd.Attribute(WordprocessingNamespace + "id");
            int startIndex = openCommentRanges.FindLastIndex(start => string.Equals(start.Id, id, StringComparison.Ordinal));
            DocxCommentRangeStart? start = startIndex < 0 ? null : openCommentRanges[startIndex];
            if (startIndex >= 0)
            {
                openCommentRanges.RemoveAt(startIndex);
            }

            commentRanges.Add(new DocxCommentRange(
                id,
                start?.SourceRunIndex,
                start?.TextOffset,
                sourceRunIndex,
                runs.Sum(run => run.Text.Length),
                ReferenceSourceRunIndex: null,
                ReferenceTextOffset: null));
        }

        void AddRevisionMarker(XElement marker)
        {
            AddRevision(paragraphRevisions, CreateRevisionInfo(marker));
            if (!TryResolveRevisionRangeMarker(marker, out DocxRevisionKind? resolvedKind, out bool isStart) ||
                resolvedKind is not { } kind)
            {
                return;
            }

            if (isStart)
            {
                openRevisionRanges.Add(new DocxRevisionRangeStart(
                    kind,
                    (string?)marker.Attribute(WordprocessingNamespace + "id"),
                    (string?)marker.Attribute(WordprocessingNamespace + "name"),
                    (string?)marker.Attribute(WordprocessingNamespace + "author"),
                    (string?)marker.Attribute(WordprocessingNamespace + "date"),
                    sourceRunIndex,
                    runs.Sum(run => run.Text.Length)));
                return;
            }

            string? id = (string?)marker.Attribute(WordprocessingNamespace + "id");
            int startIndex = openRevisionRanges.FindLastIndex(start =>
                start.Kind == kind &&
                string.Equals(start.Id, id, StringComparison.Ordinal));
            DocxRevisionRangeStart? startRange = startIndex < 0 ? null : openRevisionRanges[startIndex];
            if (startIndex >= 0)
            {
                openRevisionRanges.RemoveAt(startIndex);
            }

            revisionRanges.Add(new DocxRevisionRange(
                kind,
                id,
                startRange?.Name,
                startRange?.Author,
                startRange?.Date,
                startRange?.SourceRunIndex,
                startRange?.TextOffset,
                sourceRunIndex,
                runs.Sum(run => run.Text.Length)));
        }

        void AddHyperlinkSpan(
            XElement hyperlink,
            int sourceRunStartIndex,
            int sourceRunCount,
            int textRunStartIndex,
            int textRunCount,
            int textLength)
        {
            string? relationshipId = (string?)hyperlink.Attribute(RelationshipsNamespace + "id");
            relationships.TryGetValue(relationshipId ?? string.Empty, out OoxRelationship? relationship);
            hyperlinkSpans.Add(new DocxHyperlinkSpan(
                relationshipId,
                (string?)hyperlink.Attribute(WordprocessingNamespace + "anchor"),
                (string?)hyperlink.Attribute(WordprocessingNamespace + "tooltip"),
                (string?)hyperlink.Attribute(WordprocessingNamespace + "history"),
                relationship?.Target,
                relationship?.TargetMode,
                relationship?.ResolvedTarget,
                sourceRunStartIndex,
                sourceRunCount,
                textRunStartIndex,
                textRunCount,
                textLength));
        }

        void AddSimpleField(XElement field, DocxRevisionInfo? revision)
        {
            string? instruction = (string?)field.Attribute(WordprocessingNamespace + "instr");
            DocxFieldKind kind = ResolveFieldKind(instruction);
            string? placeholder = ResolveFieldPlaceholder(instruction);
            int fieldSourceRunIndex = sourceRunIndex;
            int fieldTextRunIndex = runs.Count;
            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
            bool hasCachedResult = FieldHasCachedResultText(field);
            if (placeholder is not null)
            {
                XElement? firstRun = field.Elements(WordprocessingNamespace + "r").FirstOrDefault();
                if (firstRun is null)
                {
                    runs.Add(new DocxTextRun(placeholder, DocxDefaults.FontSizePoints, null, false, false, false, null, null)
                    {
                        SourceRunIndex = fieldSourceRunIndex,
                        Revision = revision,
                        Revisions = RevisionList(revision)
                    });
                    AddFieldReference(
                        kind,
                        DocxFieldSourceKind.Simple,
                        instruction,
                        placeholder,
                        fieldSourceRunIndex,
                        fieldTextRunIndex,
                        fieldTextLengthStart,
                        hasCachedResult,
                        rendersCachedResult: false,
                        usesPlaceholder: true);
                    return;
                }

                AddFieldPlaceholderRun(firstRun, placeholder, fieldSourceRunIndex, revision);
                images.AddRange(ReadInlineImages(firstRun, package, relationships, revision));
                AddFieldReference(
                    kind,
                    DocxFieldSourceKind.Simple,
                    instruction,
                    placeholder,
                    fieldSourceRunIndex,
                    fieldTextRunIndex,
                    fieldTextLengthStart,
                    hasCachedResult,
                    rendersCachedResult: false,
                    usesPlaceholder: true);
                return;
            }

            foreach (XElement fieldChild in field.Elements())
            {
                AddInlineContainerChild(fieldChild, revision);
            }

            AddFieldReference(
                kind,
                DocxFieldSourceKind.Simple,
                instruction,
                placeholder,
                fieldSourceRunIndex,
                fieldTextRunIndex,
                fieldTextLengthStart,
                hasCachedResult,
                rendersCachedResult: hasCachedResult);
        }

        void AddFieldReference(
            DocxFieldKind kind,
            DocxFieldSourceKind sourceKind,
            string? instruction,
            string? placeholder,
            int fieldSourceRunIndex,
            int fieldTextRunIndex,
            int fieldTextLengthStart,
            bool hasCachedResult = false,
            bool rendersCachedResult = false,
            bool usesPlaceholder = false,
            bool hasSeparate = false,
            int nestingDepth = 0,
            int instructionRunCount = 0,
            int resultRunCount = 0)
        {
            fieldReferences.Add(new DocxFieldReference(
                kind,
                sourceKind,
                instruction,
                placeholder,
                fieldSourceRunIndex,
                fieldTextRunIndex,
                runs.Count - fieldTextRunIndex,
                runs.Sum(run => run.Text.Length) - fieldTextLengthStart)
            {
                HasSeparate = hasSeparate,
                HasCachedResult = hasCachedResult,
                RendersCachedResult = rendersCachedResult,
                UsesPlaceholder = usesPlaceholder,
                NestingDepth = nestingDepth,
                InstructionRunCount = instructionRunCount,
                ResultRunCount = resultRunCount
            });
        }

        void AddHyperlinkContainer(XElement hyperlink, DocxRevisionInfo? revision)
        {
            int sourceRunStartIndex = sourceRunIndex;
            int textRunStartIndex = runs.Count;
            int textLengthStart = runs.Sum(run => run.Text.Length);
            foreach (XElement hyperlinkChild in hyperlink.Elements())
            {
                AddInlineContainerChild(hyperlinkChild, revision);
            }

            AddHyperlinkSpan(
                hyperlink,
                sourceRunStartIndex,
                sourceRunIndex - sourceRunStartIndex,
                textRunStartIndex,
                runs.Count - textRunStartIndex,
                runs.Sum(run => run.Text.Length) - textLengthStart);
        }

        void AddFieldPlaceholderRun(XElement run, string text, int sourceRunIndex, DocxRevisionInfo? revision)
        {
            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            string? characterStyleId = ReadCharacterStyleId(run);
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            DocxRunStyleResolution runStyleResolution = CreateRunStyleResolution(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            IReadOnlyList<DocxRevisionInfo> runRevisions = ReadPropertyChangeRevisions(runProperties);
            AddRevisions(paragraphRevisions, runRevisions);
            AddResolvedTextRuns(
                runs,
                resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text,
                ApplyMarkupRevisionStyle(resolvedRun, revision, markupMode),
                runStyleResolution,
                sourceRunIndex,
                sourceTextOffsetInRun: 0,
                revision,
                MergeRevisionLists(revision, runRevisions));
        }

        void AddParagraphRun(XElement run, ref bool currentPageInstructionSeen, DocxRevisionInfo? revision)
        {
            int currentSourceRunIndex = sourceRunIndex++;
            string text = ReadRunText(run);
            string? fieldInstruction = run
                .Elements(WordprocessingNamespace + "instrText")
                .Select(instruction => (string?)instruction)
                .FirstOrDefault(value => value is not null);
            string? placeholder = ResolveFieldPlaceholder(fieldInstruction);
            int fieldTextRunIndex = runs.Count;
            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            string? characterStyleId = ReadCharacterStyleId(run);
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            DocxRunStyleResolution runStyleResolution = CreateRunStyleResolution(
                runProperties,
                paragraphStyleId,
                characterStyleId,
                styles,
                tableCellStyle?.Run);
            IReadOnlyList<DocxRevisionInfo> runRevisions = ReadPropertyChangeRevisions(runProperties);
            AddRevisions(paragraphRevisions, runRevisions);
            IReadOnlyList<DocxRevisionInfo> effectiveRevisions = MergeRevisionLists(revision, runRevisions);
            resolvedRun = ApplyMarkupRevisionStyle(resolvedRun, revision, markupMode);

            if (run.Elements().Any(IsComplexFieldMarkupElement))
            {
                AddComplexFieldAwareRun(
                    run,
                    currentSourceRunIndex,
                    resolvedRun,
                    runStyleResolution,
                    revision,
                    effectiveRevisions,
                    ref currentPageInstructionSeen);
                images.AddRange(ReadInlineImages(run, package, relationships, revision));
                return;
            }

            if (placeholder is not null)
            {
                text = placeholder;
                currentPageInstructionSeen = true;
            }
            else if (currentPageInstructionSeen && text.Trim().All(char.IsDigit))
            {
                foreach (DocxComplexFieldState field in ActiveComplexResultFields())
                {
                    field.HasCachedResult = true;
                }

                text = string.Empty;
            }

            if (placeholder is null &&
                fieldInstruction is null &&
                run.Elements().Any(IsInlineReferenceElement))
            {
                AddOrderedRunTextAndReferences(run, currentSourceRunIndex, resolvedRun, runStyleResolution, revision, effectiveRevisions);
            }
            else if (text.Length != 0)
            {
                AddParagraphDisplayText(text, currentSourceRunIndex, sourceTextOffset: 0, resolvedRun, runStyleResolution, revision, effectiveRevisions);
                AddInlineReferences(run, currentSourceRunIndex, resolvedRun, runStyleResolution, emitDisplayRuns: false, revision, effectiveRevisions);
            }
            else
            {
                AddInlineReferences(run, currentSourceRunIndex, resolvedRun, runStyleResolution, emitDisplayRuns: false, revision, effectiveRevisions);
            }

            foreach (DocxVmlTextBoxContent textBoxContent in ReadVmlTextBoxContents(run))
            {
                AddRevisions(paragraphRevisions, textBoxContent.Revisions);
                AddParagraphDisplayText(textBoxContent.Text, currentSourceRunIndex, sourceTextOffset: 0, resolvedRun, runStyleResolution, revision, effectiveRevisions);
            }

            if (fieldInstruction is not null)
            {
                AddFieldReference(
                    ResolveFieldKind(fieldInstruction),
                    DocxFieldSourceKind.ComplexInstruction,
                    fieldInstruction,
                    placeholder,
                    currentSourceRunIndex,
                    fieldTextRunIndex,
                    fieldTextLengthStart);
            }

            images.AddRange(ReadInlineImages(run, package, relationships, revision));
        }

        void AddParagraphDisplayText(
            string text,
            int currentSourceRunIndex,
            int sourceTextOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> effectiveRevisions)
        {
            if (text.Length == 0)
            {
                return;
            }

            string displayText = resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text;
            DocxComplexFieldState[] resultFields = ActiveComplexResultFields();
            if (resultFields.Length != 0)
            {
                int activeFieldTextRunIndex = runs.Count;
                int activeFieldTextLengthStart = runs.Sum(run => run.Text.Length);
                foreach (DocxComplexFieldState field in resultFields)
                {
                    field.HasCachedResult = true;
                    field.EnsureTextSpan(activeFieldTextRunIndex, activeFieldTextLengthStart);
                }
            }

            int runsBefore = runs.Count;
            AddResolvedTextRuns(runs, displayText, resolvedRun, runStyleResolution, currentSourceRunIndex, sourceTextOffset, revision, effectiveRevisions);
            int addedRuns = runs.Count - runsBefore;
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.RendersCachedResult = true;
                field.ResultRunCount += addedRuns;
            }
        }

        void AddComplexFieldAwareRun(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            ref bool currentPageInstructionSeen)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                if (child.Name == WordprocessingNamespace + "fldChar")
                {
                    ApplyComplexFieldChar(child, currentSourceRunIndex);
                }
                else if (child.Name == WordprocessingNamespace + "instrText")
                {
                    AddComplexFieldInstruction(child, currentSourceRunIndex, resolvedRun, runStyleResolution, revision, revisions, ref currentPageInstructionSeen);
                }
                else if (IsInlineReferenceElement(child))
                {
                    AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRun: true, revision, revisions);
                }
                else
                {
                    string childText = ReadRunTextChild(child);
                    if (childText.Length != 0)
                    {
                        AddComplexFieldText(
                            childText,
                            currentSourceRunIndex,
                            textOffset,
                            resolvedRun,
                            runStyleResolution,
                            revision,
                            revisions,
                            ref currentPageInstructionSeen);
                    }
                }

                textOffset += ReadRunTextChild(child).Length;
                childIndex++;
            }
        }

        void ApplyComplexFieldChar(XElement fieldChar, int currentSourceRunIndex)
        {
            string? fieldCharType = (string?)fieldChar.Attribute(WordprocessingNamespace + "fldCharType");
            if (string.Equals(fieldCharType, "begin", StringComparison.OrdinalIgnoreCase))
            {
                complexFieldStack.Add(new DocxComplexFieldState(
                    currentSourceRunIndex,
                    runs.Count,
                    runs.Sum(run => run.Text.Length),
                    complexFieldStack.Count));
                return;
            }

            if (string.Equals(fieldCharType, "separate", StringComparison.OrdinalIgnoreCase))
            {
                DocxComplexFieldState? field = CurrentComplexField();
                if (field is not null)
                {
                    field.HasSeparate = true;
                    field.InResult = true;
                    field.EnsureTextSpan(runs.Count, runs.Sum(run => run.Text.Length));
                }

                return;
            }

            if (string.Equals(fieldCharType, "end", StringComparison.OrdinalIgnoreCase))
            {
                DocxComplexFieldState? field = CurrentComplexField();
                if (field is null)
                {
                    return;
                }

                complexFieldStack.RemoveAt(complexFieldStack.Count - 1);
                AddComplexFieldReference(field, null, null);
            }
        }

        void FinalizeOpenComplexFields()
        {
            for (int fieldIndex = complexFieldStack.Count - 1; fieldIndex >= 0; fieldIndex--)
            {
                DocxComplexFieldState field = complexFieldStack[fieldIndex];
                string instruction = field.Instruction.ToString();
                string? placeholder = ResolveFieldPlaceholder(instruction);
                if (placeholder is null && (!field.HasSeparate || !field.HasCachedResult))
                {
                    continue;
                }

                AddComplexFieldReference(field, instruction, placeholder);
            }

            complexFieldStack.Clear();
        }

        void AddComplexFieldReference(DocxComplexFieldState field, string? instruction, string? placeholder)
        {
            instruction ??= field.Instruction.ToString();
            placeholder ??= ResolveFieldPlaceholder(instruction);
            int textRunIndex = field.ResultRunCount == 0 && !field.PlaceholderEmitted
                ? runs.Count
                : field.TextRunIndex;
            int textLengthStart = field.ResultRunCount == 0 && !field.PlaceholderEmitted
                ? runs.Sum(run => run.Text.Length)
                : field.TextLengthStart;
            AddFieldReference(
                ResolveFieldKind(instruction),
                DocxFieldSourceKind.ComplexInstruction,
                instruction,
                placeholder,
                field.InstructionSourceRunIndex >= 0 ? field.InstructionSourceRunIndex : field.SourceRunIndex,
                textRunIndex,
                textLengthStart,
                field.HasCachedResult,
                field.RendersCachedResult,
                field.PlaceholderEmitted,
                field.HasSeparate,
                field.NestingDepth,
                field.InstructionRunCount,
                field.ResultRunCount);
        }

        void AddComplexFieldInstruction(
            XElement instruction,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            ref bool currentPageInstructionSeen)
        {
            string instructionText = (string?)instruction ?? string.Empty;
            DocxComplexFieldState? field = CurrentComplexField();
            if (field is null)
            {
                int fieldTextRunIndex = runs.Count;
                int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
                string? placeholder = ResolveFieldPlaceholder(instructionText);
                if (placeholder is not null)
                {
                    AddResolvedTextRuns(
                        runs,
                        resolvedRun.AllCaps == true ? placeholder.ToUpperInvariant() : placeholder,
                        resolvedRun,
                        runStyleResolution,
                        currentSourceRunIndex,
                        sourceTextOffsetInRun: 0,
                        revision,
                        revisions);
                    currentPageInstructionSeen = true;
                }

                AddFieldReference(
                    ResolveFieldKind(instructionText),
                    DocxFieldSourceKind.ComplexInstruction,
                    instructionText,
                    placeholder,
                    currentSourceRunIndex,
                    fieldTextRunIndex,
                    fieldTextLengthStart,
                    usesPlaceholder: placeholder is not null,
                    instructionRunCount: 1);
                return;
            }

            if (field.InstructionSourceRunIndex < 0)
            {
                field.InstructionSourceRunIndex = currentSourceRunIndex;
            }

            field.Instruction.Append(instructionText);
            field.InstructionRunCount++;
            string? fieldPlaceholder = ResolveFieldPlaceholder(field.Instruction.ToString());
            if (fieldPlaceholder is not null && !field.PlaceholderEmitted)
            {
                field.EnsureTextSpan(runs.Count, runs.Sum(run => run.Text.Length));
                int runsBefore = runs.Count;
                AddResolvedTextRuns(
                    runs,
                    resolvedRun.AllCaps == true ? fieldPlaceholder.ToUpperInvariant() : fieldPlaceholder,
                    resolvedRun,
                    runStyleResolution,
                    currentSourceRunIndex,
                    sourceTextOffsetInRun: 0,
                    revision,
                    revisions);
                field.ResultRunCount += runs.Count - runsBefore;
                field.PlaceholderEmitted = true;
                currentPageInstructionSeen = true;
            }
        }

        void AddComplexFieldText(
            string text,
            int currentSourceRunIndex,
            int sourceTextOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions,
            ref bool currentPageInstructionSeen)
        {
            DocxComplexFieldState[] resultFields = complexFieldStack
                .Where(field => field.InResult)
                .ToArray();
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.HasCachedResult = true;
            }

            string displayText = resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text;
            bool suppressPageNumberCache = currentPageInstructionSeen && displayText.Trim().All(char.IsDigit);
            if (suppressPageNumberCache)
            {
                return;
            }

            if (resultFields.Length != 0)
            {
                int fieldTextRunIndex = runs.Count;
                int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
                foreach (DocxComplexFieldState field in resultFields)
                {
                    field.EnsureTextSpan(fieldTextRunIndex, fieldTextLengthStart);
                }
            }

            int runsBefore = runs.Count;
            AddResolvedTextRuns(
                runs,
                displayText,
                resolvedRun,
                runStyleResolution,
                currentSourceRunIndex,
                sourceTextOffset,
                revision,
                revisions);
            int addedRuns = runs.Count - runsBefore;
            foreach (DocxComplexFieldState field in resultFields)
            {
                field.RendersCachedResult = true;
                field.ResultRunCount += addedRuns;
            }
        }

        DocxComplexFieldState? CurrentComplexField()
        {
            return complexFieldStack.Count == 0 ? null : complexFieldStack[^1];
        }

        DocxComplexFieldState[] ActiveComplexResultFields()
        {
            return complexFieldStack
                .Where(field => field.InResult)
                .ToArray();
        }

        void AddOrderedRunTextAndReferences(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                if (IsInlineReferenceElement(child))
                {
                    AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRun: true, revision, revisions);
                }
                else
                {
                    string childText = ReadRunTextChild(child);
                    if (childText.Length != 0)
                    {
                        DocxComplexFieldState[] resultFields = ActiveComplexResultFields();
                        if (resultFields.Length != 0)
                        {
                            int fieldTextRunIndex = runs.Count;
                            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
                            foreach (DocxComplexFieldState field in resultFields)
                            {
                                field.HasCachedResult = true;
                                field.EnsureTextSpan(fieldTextRunIndex, fieldTextLengthStart);
                            }
                        }

                        int runsBefore = runs.Count;
                        AddResolvedTextRuns(
                            runs,
                            resolvedRun.AllCaps == true ? childText.ToUpperInvariant() : childText,
                            resolvedRun,
                            runStyleResolution,
                            currentSourceRunIndex,
                            textOffset,
                            revision,
                            revisions);
                        int addedRuns = runs.Count - runsBefore;
                        foreach (DocxComplexFieldState field in resultFields)
                        {
                            field.RendersCachedResult = true;
                            field.ResultRunCount += addedRuns;
                        }
                    }

                    textOffset += childText.Length;
                }

                childIndex++;
            }
        }

        void AddInlineReferences(
            XElement run,
            int currentSourceRunIndex,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            bool emitDisplayRuns,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                AddInlineReference(child, currentSourceRunIndex, childIndex, textOffset, resolvedRun, runStyleResolution, emitDisplayRuns, revision, revisions);

                textOffset += ReadRunTextChild(child).Length;
                childIndex++;
            }
        }

        void AddInlineReference(
            XElement child,
            int currentSourceRunIndex,
            int childIndex,
            int textOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            bool emitDisplayRun,
            DocxRevisionInfo? revision,
            IReadOnlyList<DocxRevisionInfo> revisions)
        {
            if (ResolveInlineReferenceKind(child) is not { } kind)
            {
                return;
            }

            string? customMarkFollows = kind == DocxRelatedStoryKind.Footnote || kind == DocxRelatedStoryKind.Endnote
                ? (string?)child.Attribute(WordprocessingNamespace + "customMarkFollows")
                : null;
            string? displayText = ResolveInlineReferenceDisplayText(kind, customMarkFollows);
            inlineReferences.Add(new DocxInlineReference(
                kind,
                (string?)child.Attribute(WordprocessingNamespace + "id"),
                customMarkFollows,
                displayText,
                currentSourceRunIndex,
                childIndex,
                textOffset)
            {
                Revision = revision,
                Revisions = revisions
            });
            if (kind == DocxRelatedStoryKind.Comment)
            {
                AddCommentReferenceRange((string?)child.Attribute(WordprocessingNamespace + "id"), currentSourceRunIndex, textOffset);
            }

            if (emitDisplayRun && displayText is not null)
            {
                AddInlineReferenceDisplayRun(displayText, currentSourceRunIndex, textOffset, resolvedRun, runStyleResolution, revision);
            }
        }

        void AddCommentReferenceRange(string? id, int currentSourceRunIndex, int textOffset)
        {
            int rangeIndex = commentRanges.FindIndex(range =>
                string.Equals(range.Id, id, StringComparison.Ordinal) &&
                range.ReferenceSourceRunIndex is null);
            if (rangeIndex >= 0)
            {
                commentRanges[rangeIndex] = commentRanges[rangeIndex] with
                {
                    ReferenceSourceRunIndex = currentSourceRunIndex,
                    ReferenceTextOffset = textOffset
                };
                return;
            }

            int openRangeIndex = openCommentRanges.FindLastIndex(start => string.Equals(start.Id, id, StringComparison.Ordinal));
            if (openRangeIndex >= 0)
            {
                DocxCommentRangeStart start = openCommentRanges[openRangeIndex];
                openCommentRanges.RemoveAt(openRangeIndex);
                commentRanges.Add(new DocxCommentRange(
                    id,
                    start.SourceRunIndex,
                    start.TextOffset,
                    EndSourceRunIndex: null,
                    EndTextOffset: null,
                    currentSourceRunIndex,
                    textOffset));
                return;
            }

            commentRanges.Add(new DocxCommentRange(
                id,
                StartSourceRunIndex: null,
                StartTextOffset: null,
                EndSourceRunIndex: null,
                EndTextOffset: null,
                currentSourceRunIndex,
                textOffset));
        }

        string? ResolveInlineReferenceDisplayText(DocxRelatedStoryKind kind, string? customMarkFollows)
        {
            if (!string.IsNullOrEmpty(customMarkFollows) || (kind != DocxRelatedStoryKind.Footnote && kind != DocxRelatedStoryKind.Endnote))
            {
                return null;
            }

            if (inlineReferenceCounters is null)
            {
                return null;
            }

            DocxNoteReferenceSettings settings = kind == DocxRelatedStoryKind.Endnote
                ? (documentSettings ?? DocxDocumentSettings.Empty).EndnoteReferenceSettings
                : (documentSettings ?? DocxDocumentSettings.Empty).FootnoteReferenceSettings;
            inlineReferenceCounters.TryGetValue(kind, out int current);
            int next = current == 0 ? settings.NumberStart ?? 1 : current + 1;
            inlineReferenceCounters[kind] = next;
            return FormatNoteReferenceNumber(next, settings.NumberFormatValue);
        }

        void AddInlineReferenceDisplayRun(
            string displayText,
            int currentSourceRunIndex,
            int textOffset,
            DocxResolvedRunProperties resolvedRun,
            DocxRunStyleResolution runStyleResolution,
            DocxRevisionInfo? revision)
        {
            AddResolvedTextRuns(
                runs,
                displayText,
                resolvedRun with { VerticalAlignmentValue = "superscript" },
                runStyleResolution,
                currentSourceRunIndex,
                textOffset,
                revision,
                revisions: null);
        }

        static bool IsInlineReferenceElement(XElement element)
        {
            return ResolveInlineReferenceKind(element) is not null;
        }

        static DocxRelatedStoryKind? ResolveInlineReferenceKind(XElement element)
        {
            if (element.Name == WordprocessingNamespace + "commentReference")
            {
                return DocxRelatedStoryKind.Comment;
            }

            if (element.Name == WordprocessingNamespace + "footnoteReference")
            {
                return DocxRelatedStoryKind.Footnote;
            }

            return element.Name == WordprocessingNamespace + "endnoteReference" ? DocxRelatedStoryKind.Endnote : null;
        }
    }

    private static string? ResolveFieldPlaceholder(string? instruction)
    {
        return ResolveFieldKind(instruction) switch
        {
            DocxFieldKind.NumPages => "{NUMPAGES}",
            DocxFieldKind.Page => "{PAGE}",
            _ => null
        };
    }

    private static DocxFieldKind ResolveFieldKind(string? instruction)
    {
            string? ReadFieldOpcode(string? instruction)
            {
                if (string.IsNullOrWhiteSpace(instruction))
                {
                    return null;
                }

                ReadOnlySpan<char> trimmed = instruction.AsSpan().TrimStart();
                int length = 0;
                while (length < trimmed.Length && char.IsLetter(trimmed[length]))
                {
                    length++;
                }

                return length == 0 ? null : trimmed[..length].ToString().ToUpperInvariant();
            }

        string? opcode = ReadFieldOpcode(instruction);
        return opcode switch
        {
            "PAGE" => DocxFieldKind.Page,
            "NUMPAGES" => DocxFieldKind.NumPages,
            _ => DocxFieldKind.Other
        };
    }


    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static string? ReadCharacterStyleId(XElement run)
    {
        return (string?)run
            .Element(WordprocessingNamespace + "rPr")
            ?.Element(WordprocessingNamespace + "rStyle")
            ?.Attribute(WordprocessingNamespace + "val");
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static bool IsComplexFieldMarkupElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "fldChar" ||
            element.Name == WordprocessingNamespace + "instrText";
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static void AddResolvedTextRuns(
        List<DocxTextRun> runs,
        string text,
        DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution styleResolution,
        int sourceRunIndex,
        int sourceTextOffsetInRun,
        DocxRevisionInfo? revision,
        IReadOnlyList<DocxRevisionInfo>? revisions)
    {
        var segment = new StringBuilder();
        bool? currentComplexScript = null;
        int segmentSourceOffset = sourceTextOffsetInRun;
        foreach (Rune rune in text.EnumerateRunes())
        {
            bool complexScript = DocxScriptClassifier.IsComplexScriptRune(rune.Value);
            if (currentComplexScript is not null && currentComplexScript.Value != complexScript)
            {
                string segmentText = segment.ToString();
                AddResolvedTextRun(runs, segmentText, resolvedRun, styleResolution, currentComplexScript.Value, sourceRunIndex, segmentSourceOffset, revision, revisions);
                segmentSourceOffset += segmentText.Length;
                segment.Clear();
            }

            segment.Append(rune);
            currentComplexScript = complexScript;
        }

        if (segment.Length != 0 && currentComplexScript is not null)
        {
            AddResolvedTextRun(runs, segment.ToString(), resolvedRun, styleResolution, currentComplexScript.Value, sourceRunIndex, segmentSourceOffset, revision, revisions);
        }
    }

    private static void AddResolvedTextRun(
        List<DocxTextRun> runs,
        string text,
        DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution styleResolution,
        bool complexScript,
        int sourceRunIndex,
        int sourceTextOffsetInRun,
        DocxRevisionInfo? revision,
        IReadOnlyList<DocxRevisionInfo>? revisions)
    {
        bool bold = complexScript
            ? resolvedRun.ComplexScriptBold ?? resolvedRun.Bold ?? false
            : resolvedRun.Bold ?? false;
        bool italic = complexScript
            ? resolvedRun.ComplexScriptItalic ?? resolvedRun.Italic ?? false
            : resolvedRun.Italic ?? false;
        string? fontFamily = complexScript
            ? FirstNonEmpty(resolvedRun.Fonts.ComplexScript, resolvedRun.FontFamily)
            : resolvedRun.FontFamily;
        runs.Add(new DocxTextRun(
            text,
            resolvedRun.FontSize ?? DocxDefaults.FontSizePoints,
            resolvedRun.ColorHex,
            bold,
            italic,
            resolvedRun.Underline ?? false,
            resolvedRun.UnderlineValue,
            fontFamily,
            resolvedRun.CharacterSpacingPoints ?? 0d,
            resolvedRun.AllCaps ?? false,
            resolvedRun.VerticalAlignmentValue,
            resolvedRun.Strike ?? false,
            resolvedRun.StrikeValue,
            resolvedRun.DoubleStrike ?? false,
            resolvedRun.DoubleStrikeValue,
            resolvedRun.HighlightValue,
            resolvedRun.ShadingFillHex,
            resolvedRun.ShadingValue,
            resolvedRun.ShadingColor,
            resolvedRun.SmallCaps ?? false,
            resolvedRun.SmallCapsValue,
            resolvedRun.Hidden ?? false,
            resolvedRun.HiddenValue,
            resolvedRun.UnderlineColorHex)
        {
            Fonts = resolvedRun.Fonts,
            StyleResolution = styleResolution,
            SourceRunIndex = sourceRunIndex,
            SourceTextOffsetInRun = sourceTextOffsetInRun,
            Revision = revision,
            Revisions = revisions ?? RevisionList(revision)
        });
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static double ResolveSpacingBeforePoints(DocxResolvedParagraphProperties paragraph, double lineHeight)
    {
        if (paragraph.SpacingBeforePoints is { } points)
        {
            return points;
        }

        if (OoxBoolean.IsTrue(paragraph.Spacing.BeforeAutoSpacingValue))
        {
            return WordAutomaticParagraphSpacingPoints;
        }

        return TryReadLineBasedSpacing(paragraph.Spacing.BeforeLinesValue, lineHeight, out double linePoints)
            ? linePoints
            : 0d;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static double ResolveDefaultAutoLineSpacingFactor(DocxResolvedParagraphProperties paragraph)
    {
        return HasBeforeSpacingSide(paragraph.Spacing) || HasAfterSpacingSide(paragraph.Spacing)
            ? WordSpacingTokenAutoLineSpacingFactor
            : WordUntokenedAutoLineSpacingFactor;
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static double ResolveSpacingAfterPoints(DocxResolvedParagraphProperties paragraph, double lineHeight)
    {
        if (paragraph.SpacingAfterPoints is { } points)
        {
            return points;
        }

        if (OoxBoolean.IsTrue(paragraph.Spacing.AfterAutoSpacingValue))
        {
            return WordAutomaticParagraphSpacingPoints;
        }

        return TryReadLineBasedSpacing(paragraph.Spacing.AfterLinesValue, lineHeight, out double linePoints)
            ? linePoints
            : WordDefaultSpacingAfterPoints;
    }

    private static bool TryReadLineBasedSpacing(string? value, double lineHeight, out double points)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int hundredthsOfLine))
        {
            points = lineHeight * hundredthsOfLine / 100d;
            return true;
        }

        points = 0d;
        return false;
    }

    private static string ReadRunText(XElement run)
    {
        var text = new System.Text.StringBuilder();
        foreach (XElement child in run.Elements())
        {
            text.Append(ReadRunTextChild(child));
        }

        return text.ToString();
    }

    private sealed record DocxVmlTextBoxContent(string Text, IReadOnlyList<DocxRevisionInfo> Revisions);

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IEnumerable<DocxVmlTextBoxContent> ReadVmlTextBoxContents(XElement run)
    {
        foreach (XElement textBox in run.Descendants(VmlNamespace + "textbox"))
        {
            foreach (XElement content in textBox.Descendants(WordprocessingNamespace + "txbxContent"))
            {
                DocxVmlTextBoxContent textBoxContent = ReadTextBoxContent(content);
                if (textBoxContent.Text.Length != 0)
                {
                    yield return textBoxContent;
                }
            }
        }
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxVmlTextBoxContent ReadTextBoxContent(XElement content)
    {
        var paragraphs = new List<string>();
        var revisions = new List<DocxRevisionInfo>();
        foreach (XElement paragraph in content.Elements(WordprocessingNamespace + "p"))
        {
            AddRevisions(revisions, ReadPropertyChangeRevisions(paragraph.Element(WordprocessingNamespace + "pPr")));
            string paragraphText = string.Concat(
                paragraph
                    .Descendants(WordprocessingNamespace + "r")
                    .Select(ReadRunText));
            if (paragraphText.Length != 0)
            {
                paragraphs.Add(paragraphText);
            }
        }

        return new DocxVmlTextBoxContent(string.Join("\n", paragraphs), revisions);
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static bool FieldHasCachedResultText(XElement field)
    {
        return field
            .Descendants()
            .Any(element => ReadRunTextChild(element).Length != 0);
    }

    private static string ReadRunTextChild(XElement child)
    {
        if (child.Name == WordprocessingNamespace + "t" ||
            child.Name == WordprocessingNamespace + "delText")
        {
            return (string?)child ?? string.Empty;
        }

        if (child.Name == WordprocessingNamespace + "tab")
        {
            return "\t";
        }

        if (child.Name == WordprocessingNamespace + "noBreakHyphen")
        {
            return "\u2011";
        }

        if (child.Name == WordprocessingNamespace + "softHyphen")
        {
            return "\u00AD";
        }

        if (child.Name == WordprocessingNamespace + "cr")
        {
            return "\n";
        }

        if (child.Name == WordprocessingNamespace + "br" &&
            string.IsNullOrEmpty((string?)child.Attribute(WordprocessingNamespace + "type")))
        {
            return "\n";
        }

        return string.Empty;
    }

    private static string? ReadParagraphStyleId(XElement? paragraphProperties)
    {
        return (string?)paragraphProperties?
            .Element(WordprocessingNamespace + "pStyle")
            ?.Attribute(WordprocessingNamespace + "val");
    }

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

    private static DocxSectionBreakElement ReadSectionBreak(
        XElement sectionProperties,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        XDocument? settings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        DocxRevisionInfo? inheritedRevision)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? columns = sectionProperties.Element(WordprocessingNamespace + "cols");
        return new DocxSectionBreakElement(
            ReadPageSettings(
                sectionProperties.Element(WordprocessingNamespace + "pgSz"),
                sectionProperties.Element(WordprocessingNamespace + "pgMar"),
                sectionProperties,
                settings,
                package,
                relationships,
                styles,
                numbering,
                markupMode,
                cancellationToken),
            DocxSectionBreakTypeExtensions.FromValue((string?)sectionProperties
                .Element(WordprocessingNamespace + "type")
                ?.Attribute(WordprocessingNamespace + "val")),
            (string?)columns?.Attribute(WordprocessingNamespace + "num"),
            (string?)columns?.Attribute(WordprocessingNamespace + "equalWidth"),
            (string?)columns?.Attribute(WordprocessingNamespace + "space"),
            ReadSectionColumns(columns))
        {
            Revisions = MergeRevisionLists(inheritedRevision, ReadPropertyChangeRevisions(sectionProperties))
        };
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IReadOnlyList<DocxSectionColumn> ReadSectionColumns(XElement? columns)
    {
        return columns
            ?.Elements(WordprocessingNamespace + "col")
            .Select(column => new DocxSectionColumn(
                (string?)column.Attribute(WordprocessingNamespace + "w"),
                (string?)column.Attribute(WordprocessingNamespace + "space")))
            .ToArray() ?? [];
    }

    private static IReadOnlyList<DocxFloatingDrawing> ReadFloatingDrawings(
        XElement story,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement[] paragraphs = story
            .Descendants(WordprocessingNamespace + "p")
            .ToArray();
        XElement[] bodyBlocks = story
            .Elements()
            .Where(IsBodyBlockElement)
            .ToArray();
        var drawings = new List<DocxFloatingDrawing>();
        foreach (XElement anchor in story.Descendants(WordprocessingDrawingNamespace + "anchor"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxRevisionInfo? revision = FindInheritedRevision(anchor, markupMode);
            if (IsInsideExcludedRevisionContainer(anchor, markupMode))
            {
                continue;
            }

            drawings.Add(ReadFloatingDrawing(
                anchor,
                package,
                relationships,
                styles,
                numbering,
                FindSourceParagraphIndex(anchor, paragraphs),
                FindSourceBlockIndex(anchor, bodyBlocks),
                revision,
                markupMode,
                cancellationToken));
        }

        return drawings;
    }

    private static IReadOnlyList<DocxBodyElement> ReadRelatedStoryBodyElements(
        IEnumerable<XElement> elements,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        var bodyElements = new List<DocxBodyElement>();
        foreach (DocxRevisionScopedElement scopedElement in EnumerateRevisionScopedChildren(elements, markupMode, WordprocessingNamespace + "p", WordprocessingNamespace + "tbl"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = scopedElement.Element;
            DocxRevisionInfo? inheritedRevision = scopedElement.Revision;
            if (element.Name == WordprocessingNamespace + "p")
            {
                DocxParagraph? paragraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, tableCellStyle: null, inlineReferenceCounters: null, documentSettings: null, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
                if (paragraph is not null)
                {
                    bodyElements.Add(DocxBodyElementFactory.CreateParagraph(paragraph));
                }
            }
            else if (element.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? table = ReadTable(element, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters: null, documentSettings: null, markupMode: markupMode, cancellationToken: cancellationToken, inheritedRevision: inheritedRevision);
                if (table is not null)
                {
                    bodyElements.Add(DocxBodyElementFactory.CreateTable(table));
                }
            }
        }

        return NormalizeDeletedParagraphMarkElements(bodyElements, markupMode);
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static IReadOnlyList<DocxInlineImage> ReadInlineImages(
        XElement run,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision)
    {
        var images = new List<DocxInlineImage>();
        foreach (XElement inline in run.Descendants(WordprocessingDrawingNamespace + "inline"))
        {
            DocxInlineImage? image = ReadDrawingImage(inline, package, relationships, revision);
            if (image is null)
            {
                continue;
            }

            images.Add(image);
        }

        foreach (XElement shape in run.Descendants(VmlNamespace + "shape"))
        {
            DocxInlineImage? image = ReadVmlInlineImage(shape, package, relationships, revision);
            if (image is null)
            {
                continue;
            }

            images.Add(image);
        }

        return images;
    }

    private static DocxInlineImage? ReadDrawingImage(
        XElement drawing,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision)
    {
        XElement? extent = drawing.Element(WordprocessingDrawingNamespace + "extent");
        string? relationshipId = ReadDrawingImageRelationshipId(drawing);
        if (extent is null || relationshipId is null || !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.ResolvedTarget is null)
        {
            return null;
        }

        OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
        if (imagePart is null)
        {
            return null;
        }

        return new DocxInlineImage(
            OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(extent, "cx", "DOCX")),
            OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(extent, "cy", "DOCX")),
            imagePart.ContentType,
            imagePart.Bytes,
            imagePart.Name)
        {
            Revisions = RevisionList(revision)
        };
    }

    private static string? ReadDrawingImageRelationshipId(XElement drawing)
    {
        return (string?)drawing
            .Descendants(DrawingNamespace + "blip")
            .FirstOrDefault()
            ?.Attribute(RelationshipsNamespace + "embed");
    }

    // Single caller; kept static: used once by its pipeline stage; kept for navigability.
    private static DocxInlineImage? ReadVmlInlineImage(
        XElement shape,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision)
    {
        if (!TryReadVmlImageShape(
                shape,
                relationships,
                out OoxRelationship? relationship,
                out double widthPoints,
                out double heightPoints) ||
            relationship is null ||
            relationship.ResolvedTarget is null)
        {
            return null;
        }

        OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
        if (imagePart is null)
        {
            return null;
        }

        return new DocxInlineImage(
            widthPoints,
            heightPoints,
            imagePart.ContentType,
            imagePart.Bytes,
            imagePart.Name)
        {
            Revisions = RevisionList(revision)
        };
    }

    private static bool TryReadVmlImageShape(
        XElement shape,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        out OoxRelationship? relationship,
        out double widthPoints,
        out double heightPoints)
    {
        relationship = null;
        widthPoints = 0d;
        heightPoints = 0d;
        XElement? imageData = shape.Descendants(VmlNamespace + "imagedata").FirstOrDefault();
        string? relationshipId = (string?)imageData?.Attribute(RelationshipsNamespace + "id");
        if (relationshipId is null ||
            !relationships.TryGetValue(relationshipId, out OoxRelationship? resolvedRelationship) ||
            resolvedRelationship.IsExternal ||
            resolvedRelationship.ResolvedTarget is null ||
            !TryReadVmlShapeSizePoints(shape, out widthPoints, out heightPoints))
        {
            return false;
        }

        relationship = resolvedRelationship;
        return true;
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool TryReadVmlShapeSizePoints(XElement shape, out double widthPoints, out double heightPoints)
    {
        widthPoints = 0d;
        heightPoints = 0d;
        double? width = null;
        double? height = null;
        string? style = (string?)shape.Attribute("style");
        if (!string.IsNullOrWhiteSpace(style))
        {
            foreach (string declaration in style.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                int separator = declaration.IndexOf(':');
                if (separator <= 0 || separator == declaration.Length - 1)
                {
                    continue;
                }

                string name = declaration[..separator].Trim();
                string value = declaration[(separator + 1)..].Trim();
                if (name.Equals("width", StringComparison.OrdinalIgnoreCase) &&
                    TryParseVmlLengthPoints(value, out double parsedWidth))
                {
                    width = parsedWidth;
                }
                else if (name.Equals("height", StringComparison.OrdinalIgnoreCase) &&
                    TryParseVmlLengthPoints(value, out double parsedHeight))
                {
                    height = parsedHeight;
                }
            }
        }

        if (width is null &&
            shape.Attribute("width") is { } widthAttribute &&
            TryParseVmlLengthPoints(widthAttribute.Value, out double attributeWidth))
        {
            width = attributeWidth;
        }

        if (height is null &&
            shape.Attribute("height") is { } heightAttribute &&
            TryParseVmlLengthPoints(heightAttribute.Value, out double attributeHeight))
        {
            height = attributeHeight;
        }

        if (width is not > 0d || height is not > 0d)
        {
            return false;
        }

        widthPoints = width.Value;
        heightPoints = height.Value;
        return true;
    }

    // Single caller; kept static: probe-chain link in the unsupported-feature checklist.
    private static bool TryParseVmlLengthPoints(string value, out double points)
    {
        points = 0d;
        string trimmed = value.Trim().ToLowerInvariant();
        if (trimmed.Length == 0)
        {
            return false;
        }

        int unitStart = trimmed.Length;
        while (unitStart > 0 && char.IsLetter(trimmed[unitStart - 1]))
        {
            unitStart--;
        }

        string numberText = trimmed[..unitStart].Trim();
        string unit = trimmed[unitStart..];
        if (!double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric))
        {
            return false;
        }

        points = unit switch
        {
            "" or "pt" => numeric,
            "in" => numeric * 72d,
            "cm" => numeric * 72d / 2.54d,
            "mm" => numeric * 72d / 25.4d,
            "pc" => numeric * 12d,
            "px" => numeric * 0.75d,
            _ => 0d
        };
        return points > 0d;
    }


    private readonly record struct DocxResolvedRunProperties(
        double? FontSize,
        string? ColorHex,
        string? FontFamily,
        bool? Bold,
        bool? Italic,
        bool? ComplexScriptBold,
        bool? ComplexScriptItalic,
        bool? Underline,
        string? UnderlineValue,
        DocxRunFonts Fonts,
        double? CharacterSpacingPoints,
        bool? AllCaps,
        string? VerticalAlignmentValue,
        bool? Strike,
        string? StrikeValue,
        bool? DoubleStrike,
        string? DoubleStrikeValue,
        string? HighlightValue,
        string? ShadingFillHex,
        string? ShadingValue,
        string? ShadingColor,
        bool? SmallCaps,
        string? SmallCapsValue,
        bool? Hidden,
        string? HiddenValue,
        string? UnderlineColorHex)
    {
        public static DocxResolvedRunProperties Empty { get; } = new(null, null, null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

        public DocxResolvedRunProperties Merge(DocxResolvedRunProperties other)
        {
            return new DocxResolvedRunProperties(
                other.FontSize ?? FontSize,
                other.ColorHex ?? ColorHex,
                other.FontFamily ?? FontFamily,
                other.Bold ?? Bold,
                other.Italic ?? Italic,
                other.ComplexScriptBold ?? ComplexScriptBold,
                other.ComplexScriptItalic ?? ComplexScriptItalic,
                other.Underline ?? Underline,
                other.UnderlineValue ?? UnderlineValue,
                MergeRunFonts(Fonts, other.Fonts),
                other.CharacterSpacingPoints ?? CharacterSpacingPoints,
                other.AllCaps ?? AllCaps,
                other.VerticalAlignmentValue ?? VerticalAlignmentValue,
                other.Strike ?? Strike,
                other.StrikeValue ?? StrikeValue,
                other.DoubleStrike ?? DoubleStrike,
                other.DoubleStrikeValue ?? DoubleStrikeValue,
                other.HighlightValue ?? HighlightValue,
                other.ShadingFillHex ?? ShadingFillHex,
                other.ShadingValue ?? ShadingValue,
                other.ShadingColor ?? ShadingColor,
                other.SmallCaps ?? SmallCaps,
                other.SmallCapsValue ?? SmallCapsValue,
                other.Hidden ?? Hidden,
                other.HiddenValue ?? HiddenValue,
                other.UnderlineColorHex ?? UnderlineColorHex);
        }
    }

    private sealed record DocxCommentAnchorInventory(
        IReadOnlyList<string> PackageAnchorIds,
        IReadOnlyList<string> HiddenAnchorIds);

    private sealed record DocxCommentThreadMetadata(string? ParentParagraphId, bool? IsResolved);
}
