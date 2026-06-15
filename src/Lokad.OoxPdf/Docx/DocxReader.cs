using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxReader
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

    private static readonly XNamespace WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace WordprocessingDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace DiagramNamespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private static readonly XNamespace MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace VmlNamespace = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace Office2010WordNamespace = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace Office2012WordNamespace = "http://schemas.microsoft.com/office/word/2012/wordml";

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
        Action<OoxPdfDiagnostic>? diagnosticSink = null,
        CancellationToken cancellationToken = default,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final)
    {
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
            : ReadSectionBreak(sectionProperties, package, internalRelationships, styles, numbering, settings, revisionFilteringMarkupMode, cancellationToken);
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

        double width = OoxUnits.TwipsToPoints(ParseLongAttribute(pageSize, WordprocessingNamespace + "w"));
        double height = OoxUnits.TwipsToPoints(ParseLongAttribute(pageSize, WordprocessingNamespace + "h"));
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

    private static DocxDocumentSettings ReadDocumentSettings(XDocument? settings)
    {
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

    private static DocxRevisionViewSettings ReadRevisionViewSettings(XElement? revisionView)
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

    private static DocxCommentAnchorInventory ReadCommentAnchorInventory(XDocument document, OoxPdfDocxMarkupMode markupMode)
    {
        XElement[] anchors = document
            .Descendants()
            .Where(IsCommentAnchorElement)
            .ToArray();
        string[] packageAnchorIds = anchors
            .Select(ReadCommentAnchorId)
            .Where(id => id is not null)
            .Select(id => id!)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        string[] hiddenAnchorIds = anchors
            .Where(anchor => IsInsideExcludedRevisionContainer(anchor, markupMode))
            .Select(ReadCommentAnchorId)
            .Where(id => id is not null)
            .Select(id => id!)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray();
        return new DocxCommentAnchorInventory(packageAnchorIds, hiddenAnchorIds);
    }

    private static bool IsCommentAnchorElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "commentReference" ||
            element.Name == WordprocessingNamespace + "commentRangeStart" ||
            element.Name == WordprocessingNamespace + "commentRangeEnd";
    }

    private static string? ReadCommentAnchorId(XElement element)
    {
        return (string?)element.Attribute(WordprocessingNamespace + "id");
    }

    private static IReadOnlyList<DocxFloatingDrawing> ReadFloatingDrawings(
        XDocument document,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
    {
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

    private static DocxFloatingDrawing ReadFloatingDrawing(
        XElement anchor,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        int? sourceParagraphIndex,
        int? sourceBlockIndex,
        DocxRevisionInfo? revision = null,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
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
            wrap?.Name.LocalName,
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

    private static bool IsBodyBlockElementOrRevisionContainer(XElement element)
    {
        return IsBodyBlockElement(element) || IsRevisionContainer(element);
    }

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
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
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

    private static (double Width, double Height) NormalizePageSize(double width, double height)
    {
        if (Math.Abs(width - 595d) < 0.01d && Math.Abs(height - 842d) < 0.01d)
        {
            return (594.96d, 842.04d);
        }

        return (width, height);
    }

    private static void EmitUnsupportedFeatureDiagnostics(
        OoxPackage package,
        XDocument document,
        string partName,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxMarkupContext markupContext,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken = default)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        void Emit(string id, string feature, string diagnosticPartName = "", string fallback = "Ignored", bool approximated = false)
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
                    diagnosticPartName);
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
                    fallback: "Approximated",
                    approximated: true);
            }
            else
            {
                Emit("DOCX_UNSUPPORTED_TRACKED_CHANGES", "tracked changes");
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
                fallback: markupContext.ApproximatesFormattingRevisions ? "Approximated" : "Ignored",
                approximated: markupContext.ApproximatesFormattingRevisions);
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedComplexFields(document))
        {
            Emit("DOCX_UNSUPPORTED_COMPLEX_FIELD", "complex field");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(MathNamespace + "oMath").Any() ||
            document.Descendants(MathNamespace + "oMathPara").Any())
        {
            Emit("DOCX_UNSUPPORTED_EQUATION", "equation");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "object").Any())
        {
            Emit("DOCX_UNSUPPORTED_OLE_OBJECT", "OLE object");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingDrawingNamespace + "anchor").Any(anchor => IsUnsupportedFloatingDrawingAnchor(anchor, relationships)))
        {
            Emit("DOCX_UNSUPPORTED_FLOATING_DRAWING", "floating drawing");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(ChartNamespace + "chart").Any())
        {
            Emit("DOCX_UNSUPPORTED_CHART", "chart drawing payload");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants().Any(element => element.Name.Namespace == DiagramNamespace))
        {
            Emit("DOCX_UNSUPPORTED_SMARTART", "SmartArt diagram payload");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedVml(document, relationships))
        {
            Emit("DOCX_UNSUPPORTED_VML", "VML drawing payload");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasExternalDrawingImage(document, relationships))
        {
            Emit("DOCX_UNSUPPORTED_EXTERNAL_IMAGE", "external drawing image");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "footnoteReference").Any())
        {
            Emit(
                "DOCX_APPROXIMATED_FOOTNOTE",
                "footnote",
                ResolveRelatedPartNameOrDefault(package, partName, FootnotesRelationshipType, FootnotesContentType, cancellationToken),
                "Approximated");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "endnoteReference").Any())
        {
            Emit(
                "DOCX_APPROXIMATED_ENDNOTE",
                "endnote",
                ResolveRelatedPartNameOrDefault(package, partName, EndnotesRelationshipType, EndnotesContentType, cancellationToken),
                "Approximated");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedMultiColumnSection(document))
        {
            Emit("DOCX_UNSUPPORTED_MULTI_COLUMN", "multi-column balancing or in-flow section columns", fallback: "Explicit break-only column flow is supported");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "br").Any(IsUnsupportedColumnBreak))
        {
            Emit("DOCX_UNSUPPORTED_MANUAL_BREAK", "unsupported manual column break container", fallback: "Visible body column breaks are supported");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "pPr")
            .Elements(WordprocessingNamespace + "sectPr")
            .Any(IsUnsupportedParagraphSectionBreak))
        {
            Emit("DOCX_UNSUPPORTED_SECTION_BREAK", "continuous or unknown paragraph section break", fallback: "Partially supported");
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
                Emit("DOCX_STYLE_PARAGRAPH_SPACING", "style paragraph spacing variant", stylesPartName ?? partName, "Approximated");
            }

            if (HasUnsupportedTableBorderStyle(styles))
            {
                Emit("DOCX_TABLE_BORDER_STYLE", "table border style", stylesPartName ?? partName, "Approximated");
            }

            if (HasUnsupportedTableCellTextDirection(styles))
            {
                Emit("DOCX_TABLE_TEXT_DIRECTION", "table cell text direction", stylesPartName ?? partName, "Approximated", approximated: true);
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (HasUnsupportedTableBorderStyle(document))
        {
            Emit("DOCX_TABLE_BORDER_STYLE", "table border style", partName, "Approximated");
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
            Emit("DOCX_NUMBERING_INDENT", "numbering level indent", numberingPartName ?? partName, "Approximated");
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (document.Descendants(WordprocessingNamespace + "ind").Any(HasCharacterUnitIndent) ||
            styles?.Descendants(WordprocessingNamespace + "ind").Any(HasCharacterUnitIndent) == true ||
            numbering?.Descendants(WordprocessingNamespace + "ind").Any(HasCharacterUnitIndent) == true)
        {
            Emit("DOCX_UNSUPPORTED_CHARACTER_UNIT_INDENT", "character-unit paragraph indent");
        }

        if (package.Parts.Any(p => p.Name.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase) ||
            p.ContentType.Contains("vbaProject", StringComparison.OrdinalIgnoreCase)))
        {
            Emit("DOCX_UNSUPPORTED_MACRO", "macro");
        }
    }

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

    private static bool HasSupportedFloatingDrawingImage(XElement anchor, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        string? relationshipId = ReadDrawingImageRelationshipId(anchor);
        return relationshipId is not null &&
            relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) &&
            !relationship.IsExternal &&
            relationship.ResolvedTarget is not null;
    }

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

    private static bool IsSupportedVmlElement(XElement element, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        return IsSupportedVmlImageElement(element, relationships) ||
            IsSupportedVmlTextBoxElement(element) ||
            IsInertVmlDefinitionElement(element);
    }

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

    private static bool IsInertVmlDefinitionElement(XElement element)
    {
        return element.Name == VmlNamespace + "shapetype" ||
            element.Name == VmlNamespace + "stroke" ||
            element.Name == VmlNamespace + "path";
    }

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

    private static bool HasTextBoxContent(XElement anchor)
    {
        return anchor
            .Descendants(WordprocessingNamespace + "txbxContent")
            .Any(content => content.Elements().Any());
    }

    private static bool IsSupportedHorizontalAnchorPosition(XElement? position)
    {
        return IsSupportedAnchorPosition(
            position,
            static relativeFrom => relativeFrom is "page" or "margin" or "column",
            static align => align is "left" or "center" or "right");
    }

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

    private static bool HasPositiveEmuAttribute(XElement? element, string attributeName)
    {
        return long.TryParse(
            (string?)element?.Attribute(attributeName),
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out long value) &&
            value > 0;
    }

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

    private static bool HasUnsupportedMultiColumnSection(XDocument document)
    {
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

    private static bool IsMultiColumnDeclaration(XElement columns)
    {
        return columns.Attribute(WordprocessingNamespace + "num") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int columnCount) &&
            columnCount > 1;
    }

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

        int columnCount = int.Parse(finalColumns.Attribute(WordprocessingNamespace + "num")!.Value, CultureInfo.InvariantCulture);
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

    private static bool HasUnsupportedTableBorderStyle(XDocument document)
    {
        return document
            .Descendants()
            .Any(element => element.Name.Namespace == WordprocessingNamespace &&
                IsTableBorderContainer(element.Parent) &&
                IsUnsupportedVisibleBorderStyle((string?)element.Attribute(WordprocessingNamespace + "val")));
    }

    private static bool HasUnsupportedTableCellTextDirection(XDocument document)
    {
        return document
            .Descendants(WordprocessingNamespace + "textDirection")
            .Any(IsUnsupportedTableCellTextDirection);
    }

    private static bool IsUnsupportedTableCellTextDirection(XElement textDirection)
    {
        string? value = (string?)textDirection.Attribute(WordprocessingNamespace + "val");
        return value is not null && !value.Equals("lrTb", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnsupportedTrackedChanges(XDocument document)
    {
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

    private static bool HasUnsupportedComplexFields(XDocument document)
    {
        var fields = new List<(StringBuilder Instruction, bool HasSeparate, bool InResult, bool HasCachedResult)>();
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

    private static bool IsSupportedInlineContainerParent(XElement? element)
    {
        return element?.Name == WordprocessingNamespace + "p" ||
            element?.Name == WordprocessingNamespace + "sdtContent" ||
            IsVisibleInlineContainer(element);
    }

    private static bool IsSupportedVisibleInlineContainerChild(XElement element)
    {
        return element.Name == WordprocessingNamespace + "r" ||
            element.Name == WordprocessingNamespace + "bookmarkStart" ||
            element.Name == WordprocessingNamespace + "sdt" ||
            IsVisibleInlineContainer(element);
    }

    private static bool IsSupportedInlineRunChild(XElement element)
    {
        XElement? run = element.Parent;
        XElement? container = run?.Parent;
        return run?.Name == WordprocessingNamespace + "r" &&
            (container?.Name == WordprocessingNamespace + "p" || IsVisibleInlineContainer(container));
    }

    private static bool IsSupportedComplexFieldRunChild(XElement element)
    {
        XElement? run = element.Parent;
        XElement? container = run?.Parent;
        return run?.Name == WordprocessingNamespace + "r" &&
            (container?.Name == WordprocessingNamespace + "p" || IsComplexFieldInlineContainer(container));
    }

    private static bool IsUnsupportedUnclosedComplexField(
        (StringBuilder Instruction, bool HasSeparate, bool InResult, bool HasCachedResult) field)
    {
        return ResolveFieldPlaceholder(field.Instruction.ToString()) is null &&
            (!field.HasSeparate || !field.HasCachedResult);
    }

    private static bool HasUnsupportedComplexFieldInInlineChildren(
        XElement container,
        List<(StringBuilder Instruction, bool HasSeparate, bool InResult, bool HasCachedResult)> fields)
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

    private static bool ProcessComplexFieldRunChild(
        XElement child,
        List<(StringBuilder Instruction, bool HasSeparate, bool InResult, bool HasCachedResult)> fields)
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

                fields.Add((new StringBuilder(), HasSeparate: false, InResult: false, HasCachedResult: false));
                return false;
            }

            if (string.Equals(fieldCharType, "separate", StringComparison.OrdinalIgnoreCase))
            {
                if (fields.Count == 0)
                {
                    return true;
                }

                (StringBuilder instruction, _, _, bool hasCachedResult) = fields[^1];
                fields[^1] = (instruction, HasSeparate: true, InResult: true, hasCachedResult);
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
                    fields[fieldIndex] = (instruction, hasSeparate, inResult, HasCachedResult: true);
                }
            }
        }

        return false;
    }

    private static bool IsTableBorderContainer(XElement? element)
    {
        return element is not null &&
            element.Name.Namespace == WordprocessingNamespace &&
            (element.Name.LocalName.Equals("tblBorders", StringComparison.OrdinalIgnoreCase) ||
                element.Name.LocalName.Equals("tcBorders", StringComparison.OrdinalIgnoreCase));
    }

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

    private static XDocument? LoadRelatedXmlPart(OoxPackage package, string documentPartName, string relationshipType, string contentType, out string? relatedPartName, CancellationToken cancellationToken = default)
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

    private static string ResolveRelatedPartNameOrDefault(OoxPackage package, string documentPartName, string relationshipType, string contentType, CancellationToken cancellationToken = default)
    {
        return FindRelatedPart(package, documentPartName, relationshipType, contentType, cancellationToken)?.Name ?? documentPartName;
    }

    private static OoxPart? FindRelatedPart(OoxPackage package, string documentPartName, string relationshipType, string contentType, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? relationship = package.GetRelationships(documentPartName, cancellationToken)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null
            ? package.Parts.FirstOrDefault(p => p.ContentType == contentType)
            : package.GetPart(relationship.ResolvedTarget);
    }

    private static DocxFontCatalog LoadFontCatalog(OoxPackage package, string documentPartName, CancellationToken cancellationToken = default)
    {
        XDocument? fontTable = LoadRelatedXmlPart(package, documentPartName, FontTableRelationshipType, FontTableContentType, out _, cancellationToken);
        XDocument? theme = LoadRelatedXmlPart(package, documentPartName, ThemeRelationshipType, ThemeContentType, out _, cancellationToken);
        return new DocxFontCatalog(ReadFontTableEntries(fontTable), ReadThemeFonts(theme));
    }

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

    private static IReadOnlyList<DocxParagraph> ReadParagraphs(
        XDocument document,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
    {
        var paragraphs = new List<DocxParagraph>();
        var numberingCounters = new Dictionary<(string NumId, int Level), int>();
        var inlineReferenceCounters = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (XElement paragraph in document.Descendants(WordprocessingNamespace + "body").Elements(WordprocessingNamespace + "p"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            DocxParagraph? parsed = ReadParagraph(paragraph, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters: inlineReferenceCounters, documentSettings: DocxDocumentSettings.Empty, markupMode: markupMode, cancellationToken: cancellationToken);
            if (parsed is not null)
            {
                paragraphs.Add(parsed);
            }
        }

        return paragraphs;
    }

    private static DocxParagraph? ReadParagraph(
        XElement paragraph,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxTableCellStyle? tableCellStyle = null,
        Dictionary<string, int>? inlineReferenceCounters = null,
        DocxDocumentSettings? documentSettings = null,
        DocxRevisionInfo? inheritedRevision = null,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
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
        bool hasDeletedParagraphMark = paragraphMarkRevisions.Any(revision => revision.Kind == "Deletion" && revision.SourceElement == "del");
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

        void AddRevisionRunContainer(XElement container, DocxRevisionInfo? inheritedRevision = null)
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
            if (!TryResolveRevisionRangeMarker(marker, out string kind, out bool isStart))
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
                string.Equals(start.Kind, kind, StringComparison.Ordinal) &&
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
            string kind = ResolveFieldKind(instruction);
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
                        "Simple",
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
                    "Simple",
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
                "Simple",
                instruction,
                placeholder,
                fieldSourceRunIndex,
                fieldTextRunIndex,
                fieldTextLengthStart,
                hasCachedResult,
                rendersCachedResult: hasCachedResult);
        }

        void AddFieldReference(
            string kind,
            string sourceKind,
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
                    "ComplexInstruction",
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
                AddComplexFieldReference(field);
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

        void AddComplexFieldReference(DocxComplexFieldState field, string? instruction = null, string? placeholder = null)
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
                "ComplexInstruction",
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
                    "ComplexInstruction",
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
            string? kind = ResolveInlineReferenceKind(child);
            if (kind is null)
            {
                return;
            }

            string? customMarkFollows = kind == "Footnote" || kind == "Endnote"
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
            if (kind == "Comment")
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

        string? ResolveInlineReferenceDisplayText(string kind, string? customMarkFollows)
        {
            if (!string.IsNullOrEmpty(customMarkFollows) || (kind != "Footnote" && kind != "Endnote"))
            {
                return null;
            }

            if (inlineReferenceCounters is null)
            {
                return null;
            }

            DocxNoteReferenceSettings settings = kind == "Endnote"
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
                revision);
        }

        static bool IsInlineReferenceElement(XElement element)
        {
            return ResolveInlineReferenceKind(element) is not null;
        }

        static string? ResolveInlineReferenceKind(XElement element)
        {
            if (element.Name == WordprocessingNamespace + "commentReference")
            {
                return "Comment";
            }

            if (element.Name == WordprocessingNamespace + "footnoteReference")
            {
                return "Footnote";
            }

            return element.Name == WordprocessingNamespace + "endnoteReference" ? "Endnote" : null;
        }
    }

    private static string FormatNoteReferenceNumber(int value, string? format)
    {
        return format switch
        {
            "lowerRoman" => ToRomanNumeral(value).ToLowerInvariant(),
            "upperRoman" => ToRomanNumeral(value),
            "lowerLetter" => ToAlphabeticNumber(value, upper: false),
            "upperLetter" => ToAlphabeticNumber(value, upper: true),
            _ => value.ToString(CultureInfo.InvariantCulture)
        };
    }

    private static string ToAlphabeticNumber(int value, bool upper)
    {
        if (value <= 0)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        var builder = new StringBuilder();
        int current = value;
        while (current > 0)
        {
            current--;
            char letter = (char)((upper ? 'A' : 'a') + current % 26);
            builder.Insert(0, letter);
            current /= 26;
        }

        return builder.ToString();
    }

    private static string ToRomanNumeral(int value)
    {
        if (value <= 0 || value > 3999)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        ReadOnlySpan<(int Value, string Text)> numerals =
        [
            (1000, "M"),
            (900, "CM"),
            (500, "D"),
            (400, "CD"),
            (100, "C"),
            (90, "XC"),
            (50, "L"),
            (40, "XL"),
            (10, "X"),
            (9, "IX"),
            (5, "V"),
            (4, "IV"),
            (1, "I")
        ];
        var builder = new StringBuilder();
        int current = value;
        foreach ((int numeralValue, string numeralText) in numerals)
        {
            while (current >= numeralValue)
            {
                builder.Append(numeralText);
                current -= numeralValue;
            }
        }

        return builder.ToString();
    }

    private static string? ResolveFieldPlaceholder(string? instruction)
    {
        return ResolveFieldKind(instruction) switch
        {
            "NumPages" => "{NUMPAGES}",
            "Page" => "{PAGE}",
            _ => null
        };
    }

    private static string ResolveFieldKind(string? instruction)
    {
        string? opcode = ReadFieldOpcode(instruction);
        return opcode switch
        {
            "PAGE" => "Page",
            "NUMPAGES" => "NumPages",
            _ => "Other"
        };
    }

    private static string? ReadFieldOpcode(string? instruction)
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

    private static string? ReadCharacterStyleId(XElement run)
    {
        return (string?)run
            .Element(WordprocessingNamespace + "rPr")
            ?.Element(WordprocessingNamespace + "rStyle")
            ?.Attribute(WordprocessingNamespace + "val");
    }

    private static bool IsComplexFieldMarkupElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "fldChar" ||
            element.Name == WordprocessingNamespace + "instrText";
    }

    private static void AddResolvedTextRuns(
        List<DocxTextRun> runs,
        string text,
        DocxResolvedRunProperties resolvedRun,
        DocxRunStyleResolution styleResolution,
        int sourceRunIndex = -1,
        int sourceTextOffsetInRun = 0,
        DocxRevisionInfo? revision = null,
        IReadOnlyList<DocxRevisionInfo>? revisions = null)
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
        DocxRevisionInfo? revision = null,
        IReadOnlyList<DocxRevisionInfo>? revisions = null)
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

    private static double ResolveDefaultAutoLineSpacingFactor(DocxResolvedParagraphProperties paragraph)
    {
        return HasBeforeSpacingSide(paragraph.Spacing) || HasAfterSpacingSide(paragraph.Spacing)
            ? WordSpacingTokenAutoLineSpacingFactor
            : WordUntokenedAutoLineSpacingFactor;
    }

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

    private static IReadOnlyList<DocxBodyElement> ReadBodyElements(
        XDocument document,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        XDocument? settings,
        DocxDocumentSettings documentSettings,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
    {
        var elements = new List<DocxBodyElement>();
        var numberingCounters = new Dictionary<(string NumId, int Level), int>();
        var inlineReferenceCounters = new Dictionary<string, int>(StringComparer.Ordinal);
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
                    styles);
                if (resolvedParagraph.PageBreakBefore == true)
                {
                    elements.Add(DocxBodyElementFactory.CreatePageBreak(
                        "pageBreakBefore",
                        resolvedParagraph.PageBreakBeforeValue,
                        revisions: inheritedRevision is null ? [] : [inheritedRevision]));
                }

                if (IsRunPageBreakOnlyParagraph(element, markupMode))
                {
                    DocxParagraph? breakParagraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
                    elements.Add(DocxBodyElementFactory.CreatePageBreak("runBreak", "page", breakParagraph));
                    XElement? breakParagraphSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (breakParagraphSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(breakParagraphSectionProperties, package, relationships, styles, numbering, settings, markupMode, cancellationToken, inheritedRevision));
                    }

                    continue;
                }

                if (IsRunColumnBreakOnlyParagraph(element, markupMode))
                {
                    DocxParagraph? breakParagraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
                    elements.Add(DocxBodyElementFactory.CreateManualBreak("runBreak", "column", breakParagraph));
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
                                elements.Add(DocxBodyElementFactory.CreateManualBreak("runBreak", "column"));
                            }
                            else
                            {
                                elements.Add(DocxBodyElementFactory.CreatePageBreak("runBreak", part.BreakValue));
                            }

                            continue;
                        }

                        if (part.Paragraph is null)
                        {
                            continue;
                        }

                        DocxParagraph? splitParagraph = ReadParagraph(part.Paragraph, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
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

                DocxParagraph? paragraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters: inlineReferenceCounters, documentSettings: documentSettings, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
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

    private static DocxParagraph MergeParagraphsAcrossDeletedMark(DocxParagraph first, DocxParagraph second)
    {
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

    private static int MaxSourceRunIndex(DocxParagraph paragraph)
    {
        return paragraph.Runs
            .Select(run => run.SourceRunIndex)
            .Concat(paragraph.InlineReferences.Select(reference => reference.SourceRunIndex))
            .Concat(paragraph.CommentRanges.SelectMany(range => new[] { range.StartSourceRunIndex, range.EndSourceRunIndex, range.ReferenceSourceRunIndex }).Where(index => index is not null).Select(index => index!.Value))
            .Concat(paragraph.RevisionRanges.SelectMany(range => new[] { range.StartSourceRunIndex, range.EndSourceRunIndex }).Where(index => index is not null).Select(index => index!.Value))
            .Concat(paragraph.FieldReferences.Select(field => field.SourceRunIndex))
            .Concat(paragraph.Hyperlinks.Select(link => link.SourceRunStartIndex))
            .Concat(paragraph.BookmarkAnchors.Select(anchor => anchor.SourceRunIndex))
            .Where(index => index >= 0)
            .DefaultIfEmpty(-1)
            .Max();
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

    private static DocxTextRun ShiftRun(DocxTextRun run, int sourceRunOffset)
    {
        return run with { SourceRunIndex = ShiftSourceRunIndex(run.SourceRunIndex, sourceRunOffset) };
    }

    private static DocxInlineReference ShiftInlineReference(DocxInlineReference reference, int sourceRunOffset)
    {
        return reference with { SourceRunIndex = ShiftSourceRunIndex(reference.SourceRunIndex, sourceRunOffset) };
    }

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

    private static DocxHyperlinkSpan ShiftHyperlink(DocxHyperlinkSpan hyperlink, int sourceRunOffset, int textRunOffset)
    {
        return hyperlink with
        {
            SourceRunStartIndex = ShiftSourceRunIndex(hyperlink.SourceRunStartIndex, sourceRunOffset),
            TextRunStartIndex = hyperlink.TextRunStartIndex + textRunOffset
        };
    }

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

    private static bool IsRunPageBreakOnlyParagraph(XElement paragraph, OoxPdfDocxMarkupMode markupMode)
    {
        return IsRunBreakOnlyParagraph(paragraph, IsPageBreak, markupMode);
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

    private static bool IsRevisionMarkerElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "moveFromRangeStart" ||
            element.Name == WordprocessingNamespace + "moveFromRangeEnd" ||
            element.Name == WordprocessingNamespace + "moveToRangeStart" ||
            element.Name == WordprocessingNamespace + "moveToRangeEnd";
    }

    private static bool TryResolveRevisionRangeMarker(XElement element, out string kind, out bool isStart)
    {
        if (element.Name == WordprocessingNamespace + "moveFromRangeStart")
        {
            kind = "MoveFrom";
            isStart = true;
            return true;
        }

        if (element.Name == WordprocessingNamespace + "moveFromRangeEnd")
        {
            kind = "MoveFrom";
            isStart = false;
            return true;
        }

        if (element.Name == WordprocessingNamespace + "moveToRangeStart")
        {
            kind = "MoveTo";
            isStart = true;
            return true;
        }

        if (element.Name == WordprocessingNamespace + "moveToRangeEnd")
        {
            kind = "MoveTo";
            isStart = false;
            return true;
        }

        kind = string.Empty;
        isStart = false;
        return false;
    }

    private static string? RevisionKind(XElement element)
    {
        if (element.Name == WordprocessingNamespace + "ins")
        {
            return "Insertion";
        }

        if (element.Name == WordprocessingNamespace + "del")
        {
            return "Deletion";
        }

        if (element.Name == WordprocessingNamespace + "moveFrom")
        {
            return "MoveFrom";
        }

        if (element.Name == WordprocessingNamespace + "moveTo")
        {
            return "MoveTo";
        }

        if (element.Name == WordprocessingNamespace + "rPrChange")
        {
            return "RunPropertiesChange";
        }

        if (element.Name == WordprocessingNamespace + "pPrChange")
        {
            return "ParagraphPropertiesChange";
        }

        if (element.Name == WordprocessingNamespace + "tblPrChange")
        {
            return "TablePropertiesChange";
        }

        if (element.Name == WordprocessingNamespace + "trPrChange")
        {
            return "TableRowPropertiesChange";
        }

        if (element.Name == WordprocessingNamespace + "tcPrChange")
        {
            return "TableCellPropertiesChange";
        }

        if (element.Name == WordprocessingNamespace + "sectPrChange")
        {
            return "SectionPropertiesChange";
        }

        if (element.Name == WordprocessingNamespace + "moveFromRangeStart")
        {
            return "MoveFromRangeStart";
        }

        if (element.Name == WordprocessingNamespace + "moveFromRangeEnd")
        {
            return "MoveFromRangeEnd";
        }

        if (element.Name == WordprocessingNamespace + "moveToRangeStart")
        {
            return "MoveToRangeStart";
        }

        return element.Name == WordprocessingNamespace + "moveToRangeEnd" ? "MoveToRangeEnd" : null;
    }

    private static DocxRevisionInfo? CreateRevisionInfo(XElement element)
    {
        string? kind = RevisionKind(element);
        return kind is null
            ? null
            : new DocxRevisionInfo(
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
            .Where(revision => revision is not null)
            .Select(revision => revision!)
            .ToArray() ?? [];
    }

    private static string? RevisionPropertyChangeFamily(XElement element)
    {
        if (element.Name == WordprocessingNamespace + "rPrChange")
        {
            return "Run";
        }

        if (element.Name == WordprocessingNamespace + "pPrChange")
        {
            return "Paragraph";
        }

        if (element.Name == WordprocessingNamespace + "tblPrChange")
        {
            return "Table";
        }

        if (element.Name == WordprocessingNamespace + "trPrChange")
        {
            return "Row";
        }

        if (element.Name == WordprocessingNamespace + "tcPrChange")
        {
            return "Cell";
        }

        return element.Name == WordprocessingNamespace + "sectPrChange" ? "Section" : null;
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
            "Insertion" => run with { ColorHex = run.ColorHex ?? "0000FF", Underline = true, UnderlineValue = run.UnderlineValue ?? "single" },
            "Deletion" => run with { ColorHex = run.ColorHex ?? "C00000", Strike = true, StrikeValue = run.StrikeValue ?? "true" },
            "MoveFrom" => run with { ColorHex = run.ColorHex ?? "C00000", DoubleStrike = true, DoubleStrikeValue = run.DoubleStrikeValue ?? "true" },
            "MoveTo" => run with { ColorHex = run.ColorHex ?? "008000", Underline = true, UnderlineValue = run.UnderlineValue ?? "single" },
            _ => run
        };
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
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default,
        DocxRevisionInfo? inheritedRevision = null)
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
            (string?)sectionProperties
                .Element(WordprocessingNamespace + "type")
                ?.Attribute(WordprocessingNamespace + "val"),
            (string?)columns?.Attribute(WordprocessingNamespace + "num"),
            (string?)columns?.Attribute(WordprocessingNamespace + "equalWidth"),
            (string?)columns?.Attribute(WordprocessingNamespace + "space"),
            ReadSectionColumns(columns))
        {
            Revisions = MergeRevisionLists(inheritedRevision, ReadPropertyChangeRevisions(sectionProperties))
        };
    }

    private static IReadOnlyList<DocxSectionColumn> ReadSectionColumns(XElement? columns)
    {
        return columns
            ?.Elements(WordprocessingNamespace + "col")
            .Select(column => new DocxSectionColumn(
                (string?)column.Attribute(WordprocessingNamespace + "w"),
                (string?)column.Attribute(WordprocessingNamespace + "space")))
            .ToArray() ?? [];
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxBodyElement>> ReadReferencedHeaderFooterBodyElementsByType(
        XContainer referenceRoot,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        string relationshipType,
        string referenceElementName,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
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
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
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
            IReadOnlyDictionary<string, OoxRelationship> partRelationships = package.GetRelationships(part.Name, cancellationToken)
                .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
                .ToDictionary(r => r.Id, StringComparer.Ordinal);
            string type = (string?)reference.Attribute(WordprocessingNamespace + "type") ?? "default";
            drawings[type] = ReadFloatingDrawings(partXml, package, partRelationships, styles, numbering, markupMode, cancellationToken);
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
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
    {
        return ReadCommentStories(package, documentPartName, styles, numbering, markupMode, cancellationToken)
            .Concat(ReadRelatedStories(package, documentPartName, styles, numbering, FootnotesRelationshipType, FootnotesContentType, "Footnote", "footnote", markupMode, cancellationToken))
            .Concat(ReadRelatedStories(package, documentPartName, styles, numbering, EndnotesRelationshipType, EndnotesContentType, "Endnote", "endnote", markupMode, cancellationToken))
            .ToArray();
    }

    private static IReadOnlyList<DocxRelatedStory> ReadCommentStories(
        OoxPackage package,
        string documentPartName,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
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
            "Comment",
            "comment",
            markupMode,
            cancellationToken,
            threadMetadataByParagraphId);
        return ResolveCommentThreadParents(stories);
    }

    private static IReadOnlyList<DocxRelatedStory> ReadRelatedStories(
        OoxPackage package,
        string documentPartName,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        string relationshipType,
        string contentType,
        string kind,
        string storyElementName,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default,
        IReadOnlyDictionary<string, DocxCommentThreadMetadata>? commentThreadMetadataByParagraphId = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? part = FindRelatedPart(package, documentPartName, relationshipType, contentType, cancellationToken);
        if (part is null)
        {
            return [];
        }

        using Stream stream = part.OpenRead();
        XDocument partXml = SafeXml.Load(stream, cancellationToken);
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

    private static DocxRelatedStory ReadRelatedStory(
        string kind,
        string partName,
        XElement story,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default,
        IReadOnlyDictionary<string, DocxCommentThreadMetadata>? commentThreadMetadataByParagraphId = null)
    {
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
        string? paragraphId = kind == "Comment" ? ReadCommentParagraphId(story) : null;
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
            (string?)story.Attribute(WordprocessingNamespace + "type"))
        {
            FloatingDrawings = floatingDrawings,
            CommentMetadata = kind == "Comment"
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
    }

    private static IReadOnlyDictionary<string, DocxCommentThreadMetadata> ReadCommentThreadMetadata(
        OoxPackage package,
        string documentPartName,
        CancellationToken cancellationToken = default)
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

    private static IReadOnlyList<DocxRelatedStory> ResolveCommentThreadParents(IReadOnlyList<DocxRelatedStory> stories)
    {
        Dictionary<string, string> commentIdByParagraphId = stories
            .Where(story => story.CommentMetadata?.ParagraphId is not null && story.Id is not null)
            .GroupBy(story => story.CommentMetadata!.ParagraphId!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First().Id!, StringComparer.OrdinalIgnoreCase);
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

    private static string? ReadCommentParagraphId(XElement story)
    {
        return story
            .Descendants(WordprocessingNamespace + "p")
            .Select(paragraph => (string?)paragraph.Attribute(Office2010WordNamespace + "paraId"))
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }

    private static IReadOnlyList<DocxFloatingDrawing> ReadFloatingDrawings(
        XElement story,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
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
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
    {
        var bodyElements = new List<DocxBodyElement>();
        foreach (DocxRevisionScopedElement scopedElement in EnumerateRevisionScopedChildren(elements, markupMode, WordprocessingNamespace + "p", WordprocessingNamespace + "tbl"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement element = scopedElement.Element;
            DocxRevisionInfo? inheritedRevision = scopedElement.Revision;
            if (element.Name == WordprocessingNamespace + "p")
            {
                DocxParagraph? paragraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships, inheritedRevision: inheritedRevision, markupMode: markupMode, cancellationToken: cancellationToken);
                if (paragraph is not null)
                {
                    bodyElements.Add(DocxBodyElementFactory.CreateParagraph(paragraph));
                }
            }
            else if (element.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? table = ReadTable(element, styles, numbering, numberingCounters, package, relationships, markupMode: markupMode, cancellationToken: cancellationToken, inheritedRevision: inheritedRevision);
                if (table is not null)
                {
                    bodyElements.Add(DocxBodyElementFactory.CreateTable(table));
                }
            }
        }

        return NormalizeDeletedParagraphMarkElements(bodyElements, markupMode);
    }

    private static IReadOnlyList<DocxParagraph> ReadParagraphElements(
        IEnumerable<XElement> paragraphElements,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default)
    {
        var wrapper = new XDocument(new XElement(WordprocessingNamespace + "document", new XElement(WordprocessingNamespace + "body", paragraphElements)));
        return ReadParagraphs(wrapper, styles, numbering, package, relationships, markupMode, cancellationToken);
    }

    private static DocxTable? ReadTable(
        XElement table,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        Dictionary<string, int>? inlineReferenceCounters = null,
        DocxDocumentSettings? documentSettings = null,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default,
        DocxRevisionInfo? inheritedRevision = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? tableProperties = table.Element(WordprocessingNamespace + "tblPr");
        string? layoutValue = (string?)tableProperties
            ?.Element(WordprocessingNamespace + "tblLayout")
            ?.Attribute(WordprocessingNamespace + "type");
        string? tableStyleId = (string?)tableProperties
            ?.Element(WordprocessingNamespace + "tblStyle")
            ?.Attribute(WordprocessingNamespace + "val");
        XElement? tableWidth = tableProperties
            ?.Element(WordprocessingNamespace + "tblW");
        XElement? tableIndent = tableProperties
            ?.Element(WordprocessingNamespace + "tblInd");
        XElement? tableCellSpacing = tableProperties
            ?.Element(WordprocessingNamespace + "tblCellSpacing");
        var tableRevisions = new List<DocxRevisionInfo>();
        AddRevision(tableRevisions, inheritedRevision);
        AddRevisions(tableRevisions, ReadPropertyChangeRevisions(tableProperties));
        DocxTableLook tableLook = ReadTableLook(tableProperties);
        IReadOnlyList<DocxTableCellBorder> tableBorders = ReadTableBorders(tableProperties);
        DocxTableStyle tableStyle = tableStyleId is not null && styles.TableStyles.TryGetValue(tableStyleId, out DocxTableStyle? parsedTableStyle)
            ? parsedTableStyle
            : styles.DefaultTableStyle ?? DocxTableStyle.Empty;
        DocxTableCellMargins tableCellMargins = ReadTableStyleCellMargins(tableProperties);
        IReadOnlyList<double> columns = table
            .Element(WordprocessingNamespace + "tblGrid")
            ?.Elements(WordprocessingNamespace + "gridCol")
            .Select(column => ReadTwipsAttribute(column, WordprocessingNamespace + "w") ?? 72d)
            .ToArray() ?? [];
        bool hasExplicitGrid = columns.Count != 0;
        var rows = new List<DocxTableRow>();
        DocxRevisionScopedElement[] rowElements = EnumerateRevisionScopedChildren(table.Elements(), markupMode, WordprocessingNamespace + "tr").ToArray();
        for (int rowIndex = 0; rowIndex < rowElements.Length; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement row = rowElements[rowIndex].Element;
            XElement? rowProperties = row.Element(WordprocessingNamespace + "trPr");
            var rowRevisions = new List<DocxRevisionInfo>();
            AddRevision(rowRevisions, rowElements[rowIndex].Revision);
            AddRevisions(rowRevisions, ReadPropertyChangeRevisions(rowProperties));
            DocxTableCellMargins rowExceptionMargins = ReadTablePropertyExceptionCellMargins(row);
            DocxTableCellMargins rowInheritedMargins = MergeTableCellMargins(rowExceptionMargins, MergeTableCellMargins(tableCellMargins, tableStyle.Cell.Margins));
            var cells = new List<DocxTableCell>();
            DocxRevisionScopedElement[] cellElements = EnumerateRevisionScopedChildren(row.Elements(), markupMode, WordprocessingNamespace + "tc").ToArray();
            for (int cellIndex = 0; cellIndex < cellElements.Length; cellIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                XElement cell = cellElements[cellIndex].Element;
                DocxRevisionInfo? inheritedCellRevision = cellElements[cellIndex].Revision ?? rowElements[rowIndex].Revision;
                XElement? cellProperties = cell.Element(WordprocessingNamespace + "tcPr");
                var cellRevisions = new List<DocxRevisionInfo>();
                AddRevision(cellRevisions, cellElements[cellIndex].Revision);
                AddRevisions(cellRevisions, ReadPropertyChangeRevisions(cellProperties));
                DocxTableCellConditionalFormat? conditionalFormat = ReadTableCellConditionalFormat(cellProperties);
                DocxTableCellStyle conditionalStyle = ResolveTableCellStyle(tableStyle, tableLook, conditionalFormat, rowIndex, cellIndex, rowElements.Length, cellElements.Length);
                IReadOnlyList<DocxBodyElement> cellBodyElements = ReadTableCellBodyElements(
                    cell,
                    styles,
                    numbering,
                    numberingCounters,
                    package,
                    relationships,
                    conditionalStyle,
                    inlineReferenceCounters,
                    documentSettings,
                    markupMode,
                    cancellationToken,
                    inheritedCellRevision);
                IReadOnlyList<DocxParagraph> paragraphs = DocxBlockTraversal.EnumerateDirectParagraphs(cellBodyElements).ToArray();
                string text = string.Join(" ", paragraphs
                    .Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text)))
                    .Where(t => t.Length != 0));
                XElement? shading = cellProperties?.Element(WordprocessingNamespace + "shd");
                string? fill = (string?)shading?.Attribute(WordprocessingNamespace + "fill") ?? conditionalStyle.FillHex;
                string? shadingValue = (string?)shading?.Attribute(WordprocessingNamespace + "val") ?? conditionalStyle.ShadingValue;
                string? shadingColor = (string?)shading?.Attribute(WordprocessingNamespace + "color") ?? conditionalStyle.ShadingColor;
                string? verticalAlignment = (string?)cellProperties
                    ?.Element(WordprocessingNamespace + "vAlign")
                    ?.Attribute(WordprocessingNamespace + "val")
                    ?? conditionalStyle.VerticalAlignmentValue;
                XElement? cellWidth = cellProperties?.Element(WordprocessingNamespace + "tcW");
                XElement? verticalMerge = cellProperties?.Element(WordprocessingNamespace + "vMerge");
                XElement? noWrap = cellProperties?.Element(WordprocessingNamespace + "noWrap");
                XElement? fitText = cellProperties?.Element(WordprocessingNamespace + "tcFitText");
                string? textDirectionValue = cellProperties?.Element(WordprocessingNamespace + "textDirection") is { } textDirection
                    ? (string?)textDirection.Attribute(WordprocessingNamespace + "val")
                    : conditionalStyle.TextDirectionValue;
                bool resolvedNoWrap = noWrap is not null
                    ? ReadOnOff(noWrap) == true
                    : conditionalStyle.NoWrap == true;
                string? resolvedNoWrapValue = noWrap is not null
                    ? (string?)noWrap.Attribute(WordprocessingNamespace + "val")
                    : conditionalStyle.NoWrapValue;
                bool resolvedFitText = fitText is not null
                    ? ReadOnOff(fitText) == true
                    : conditionalStyle.FitText == true;
                string? resolvedFitTextValue = fitText is not null
                    ? (string?)fitText.Attribute(WordprocessingNamespace + "val")
                    : conditionalStyle.FitTextValue;
                IReadOnlyList<DocxTableCellBorder> directBorders = ReadTableCellBorders(cellProperties);
                IReadOnlyList<DocxTableCellBorder> borders = ResolveTableCellBorders(
                    directBorders,
                    conditionalStyle.Borders,
                    tableStyle.TableBorders,
                    tableBorders,
                    rowIndex,
                    cellIndex,
                    rowElements.Length,
                    cellElements.Length);
                DocxTableCellMargins inheritedMargins = MergeTableCellMargins(rowInheritedMargins, conditionalStyle.Margins);
                DocxTableCellMargins margins = MergeTableCellMargins(ReadTableCellMargins(cellProperties), inheritedMargins);
                cells.Add(new DocxTableCell(
                    text,
                    paragraphs,
                    fill,
                    shadingValue,
                    shadingColor,
                    verticalAlignment,
                    borders,
                    margins,
                    ReadDxaWidth(cellWidth),
                    (string?)cellWidth?.Attribute(WordprocessingNamespace + "w"),
                    (string?)cellWidth?.Attribute(WordprocessingNamespace + "type"),
                    ReadGridSpan(cellProperties),
                    (string?)cellProperties
                        ?.Element(WordprocessingNamespace + "gridSpan")
                        ?.Attribute(WordprocessingNamespace + "val"),
                    conditionalFormat,
                    verticalMerge is not null,
                    (string?)verticalMerge?.Attribute(WordprocessingNamespace + "val"),
                    resolvedNoWrap,
                    resolvedNoWrapValue,
                    resolvedFitText,
                    resolvedFitTextValue,
                    textDirectionValue)
                {
                    BodyElements = cellBodyElements,
                    Revisions = cellRevisions
                });
            }

            if (cells.Count > 0)
            {
                XElement? header = rowProperties?.Element(WordprocessingNamespace + "tblHeader");
                XElement? cantSplit = rowProperties?.Element(WordprocessingNamespace + "cantSplit");
                XElement? rowHeight = rowProperties?.Element(WordprocessingNamespace + "trHeight");
                rows.Add(new DocxTableRow(
                    cells,
                    ReadTableRowHeight(rowHeight),
                    ReadOnOff(header) == true,
                    (string?)header?.Attribute(WordprocessingNamespace + "val"),
                    (string?)rowHeight?.Attribute(WordprocessingNamespace + "val"),
                    (string?)rowHeight?.Attribute(WordprocessingNamespace + "hRule"),
                    HasAnyTableCellMargin(rowExceptionMargins) ? rowExceptionMargins : null,
                    ReadOnOff(cantSplit) == true,
                    (string?)cantSplit?.Attribute(WordprocessingNamespace + "val"))
                {
                    Revisions = rowRevisions
                });
            }
        }

        if (rows.Count == 0)
        {
            return null;
        }

        if (columns.Count == 0)
        {
            int inferredGridColumns = rows
                .Select(row => row.Cells.Sum(cell => Math.Max(1, cell.GridSpan)))
                .DefaultIfEmpty(0)
                .Max();
            columns = Enumerable.Repeat(72d, inferredGridColumns).ToArray();
        }

        return new DocxTable(
            layoutValue ?? tableStyle.Table.LayoutValue,
            columns,
            rows,
            tableStyleId,
            tableWidth is not null ? ReadDxaWidth(tableWidth) : tableStyle.Table.PreferredWidthPoints,
            tableWidth is not null ? (string?)tableWidth.Attribute(WordprocessingNamespace + "w") : tableStyle.Table.PreferredWidthValue,
            tableWidth is not null ? (string?)tableWidth.Attribute(WordprocessingNamespace + "type") : tableStyle.Table.PreferredWidthType,
            tableIndent is not null ? ReadDxaWidth(tableIndent) : tableStyle.Table.IndentPoints,
            tableIndent is not null ? (string?)tableIndent.Attribute(WordprocessingNamespace + "w") : tableStyle.Table.IndentValue,
            tableIndent is not null ? (string?)tableIndent.Attribute(WordprocessingNamespace + "type") : tableStyle.Table.IndentType,
            tableCellSpacing is not null ? ReadDxaWidth(tableCellSpacing) : tableStyle.Table.CellSpacingPoints,
            tableCellSpacing is not null ? (string?)tableCellSpacing.Attribute(WordprocessingNamespace + "w") : tableStyle.Table.CellSpacingValue,
            tableCellSpacing is not null ? (string?)tableCellSpacing.Attribute(WordprocessingNamespace + "type") : tableStyle.Table.CellSpacingType,
            tableLook,
            hasExplicitGrid)
        {
            Revisions = tableRevisions
        };
    }

    private static double? ReadDxaWidth(XElement? width)
    {
        string? type = (string?)width?.Attribute(WordprocessingNamespace + "type");
        if (type is not null && !type.Equals("dxa", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return ReadTwipsAttribute(width, WordprocessingNamespace + "w");
    }

    private static int ReadGridSpan(XElement? cellProperties)
    {
        return cellProperties
            ?.Element(WordprocessingNamespace + "gridSpan")
            ?.Attribute(WordprocessingNamespace + "val") is { } span &&
            int.TryParse(span.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
                ? Math.Max(1, parsed)
                : 1;
    }

    private static IReadOnlyList<DocxBodyElement> ReadTableCellBodyElements(
        XElement cell,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxTableCellStyle tableCellStyle,
        Dictionary<string, int>? inlineReferenceCounters = null,
        DocxDocumentSettings? documentSettings = null,
        OoxPdfDocxMarkupMode markupMode = OoxPdfDocxMarkupMode.Final,
        CancellationToken cancellationToken = default,
        DocxRevisionInfo? inheritedRevision = null)
    {
        var elements = new List<DocxBodyElement>();
        foreach (DocxRevisionScopedElement scopedChild in EnumerateRevisionScopedChildren(cell.Elements(), markupMode, WordprocessingNamespace + "p", WordprocessingNamespace + "tbl"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement child = scopedChild.Element;
            DocxRevisionInfo? childRevision = scopedChild.Revision ?? inheritedRevision;
            if (child.Name == WordprocessingNamespace + "tcPr")
            {
                continue;
            }

            if (child.Name == WordprocessingNamespace + "p")
            {
                if (IsRunColumnBreakOnlyParagraph(child, markupMode))
                {
                    DocxParagraph? breakParagraph = ReadParagraph(
                        child,
                        styles,
                        numbering,
                        numberingCounters,
                        package,
                        relationships,
                        tableCellStyle,
                        inlineReferenceCounters: inlineReferenceCounters,
                        documentSettings: documentSettings,
                        inheritedRevision: childRevision,
                        markupMode: markupMode,
                        cancellationToken: cancellationToken);
                    elements.Add(DocxBodyElementFactory.CreateManualBreak("runBreak", "column", breakParagraph));
                    continue;
                }

                if (HasRunPageOrColumnBreak(child, markupMode))
                {
                    foreach (ParagraphBreakPart part in SplitParagraphAtRunBreaks(child, markupMode))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (part.BreakValue is not null)
                        {
                            elements.Add(string.Equals(part.BreakValue, "column", StringComparison.OrdinalIgnoreCase)
                                ? DocxBodyElementFactory.CreateManualBreak("runBreak", "column")
                                : DocxBodyElementFactory.CreatePageBreak("runBreak", part.BreakValue));
                            continue;
                        }

                        if (part.Paragraph is null)
                        {
                            continue;
                        }

                        DocxParagraph? splitParagraph = ReadParagraph(
                            part.Paragraph,
                            styles,
                            numbering,
                            numberingCounters,
                            package,
                            relationships,
                            tableCellStyle,
                            inlineReferenceCounters: inlineReferenceCounters,
                            documentSettings: documentSettings,
                            inheritedRevision: childRevision,
                            markupMode: markupMode,
                            cancellationToken: cancellationToken);
                        if (splitParagraph is not null)
                        {
                            elements.Add(DocxBodyElementFactory.CreateParagraph(AdjustBreakParagraphFragment(splitParagraph, part)));
                        }
                    }

                    continue;
                }

                DocxParagraph? parsed = ReadParagraph(
                    child,
                    styles,
                    numbering,
                    numberingCounters,
                    package,
                    relationships,
                    tableCellStyle,
                    inlineReferenceCounters: inlineReferenceCounters,
                    documentSettings: documentSettings,
                    inheritedRevision: childRevision,
                    markupMode: markupMode,
                    cancellationToken: cancellationToken);
                if (parsed is not null)
                {
                    elements.Add(DocxBodyElementFactory.CreateParagraph(parsed));
                }
            }
            else if (child.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? nestedTable = ReadTable(child, styles, numbering, numberingCounters, package, relationships, inlineReferenceCounters, documentSettings, markupMode, cancellationToken, childRevision);
                if (nestedTable is not null)
                {
                    elements.Add(DocxBodyElementFactory.CreateTable(nestedTable));
                }
            }
        }

        return NormalizeDeletedParagraphMarkElements(elements, markupMode);
    }

    private static void AddImplicitTerminalTableParagraph(List<DocxBodyElement> elements)
    {
        if (elements.Count == 0 ||
            elements[^1] is not DocxTableElement)
        {
            return;
        }

        elements.Add(new DocxImplicitParagraphElement("terminalTable"));
    }

    private static DocxTableCellMargins ReadTableCellMargins(XElement? cellProperties)
    {
        XElement? margins = cellProperties?.Element(WordprocessingNamespace + "tcMar");
        return ReadCellMargins(margins);
    }

    private static DocxTableCellMargins ReadTableStyleCellMargins(XElement? tableProperties)
    {
        XElement? margins = tableProperties?.Element(WordprocessingNamespace + "tblCellMar");
        return ReadCellMargins(margins);
    }

    private static DocxTableCellMargins ReadTablePropertyExceptionCellMargins(XElement row)
    {
        XElement? margins = row
            .Element(WordprocessingNamespace + "tblPrEx")
            ?.Element(WordprocessingNamespace + "tblCellMar");
        return ReadCellMargins(margins);
    }

    private static DocxTableCellMargins ReadCellMargins(XElement? margins)
    {
        return new DocxTableCellMargins(
            ReadMargin(margins, "top"),
            ReadMargin(margins, "right"),
            ReadMargin(margins, "bottom"),
            ReadMargin(margins, "left"),
            ReadMarginValue(margins, "top"),
            ReadMarginValue(margins, "right"),
            ReadMarginValue(margins, "bottom"),
            ReadMarginValue(margins, "left"));
    }

    private static DocxTableCellMargins MergeTableCellMargins(DocxTableCellMargins direct, DocxTableCellMargins inherited)
    {
        return new DocxTableCellMargins(
            direct.TopPoints ?? inherited.TopPoints,
            direct.RightPoints ?? inherited.RightPoints,
            direct.BottomPoints ?? inherited.BottomPoints,
            direct.LeftPoints ?? inherited.LeftPoints,
            direct.TopValue ?? inherited.TopValue,
            direct.RightValue ?? inherited.RightValue,
            direct.BottomValue ?? inherited.BottomValue,
            direct.LeftValue ?? inherited.LeftValue);
    }

    private static bool HasAnyTableCellMargin(DocxTableCellMargins margins)
    {
        return margins.TopValue is not null ||
            margins.RightValue is not null ||
            margins.BottomValue is not null ||
            margins.LeftValue is not null;
    }

    private static double? ReadMargin(XElement? margins, string edge)
    {
        XElement? margin = margins?.Element(WordprocessingNamespace + edge);
        string? type = (string?)margin?.Attribute(WordprocessingNamespace + "type");
        if (type is not null && !type.Equals("dxa", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return ReadTwipsAttribute(margin, WordprocessingNamespace + "w");
    }

    private static string? ReadMarginValue(XElement? margins, string edge)
    {
        return (string?)margins
            ?.Element(WordprocessingNamespace + edge)
            ?.Attribute(WordprocessingNamespace + "w");
    }

    private static IReadOnlyList<DocxTableCellBorder> ReadTableCellBorders(XElement? cellProperties)
    {
        return ReadBorderElements(cellProperties?.Element(WordprocessingNamespace + "tcBorders"));
    }

    private static IReadOnlyList<DocxTableCellBorder> ReadTableBorders(XElement? tableProperties)
    {
        return ReadBorderElements(tableProperties?.Element(WordprocessingNamespace + "tblBorders"));
    }

    private static DocxTableLook ReadTableLook(XElement? tableProperties)
    {
        XElement? look = tableProperties?.Element(WordprocessingNamespace + "tblLook");
        if (look is null)
        {
            return DocxTableLook.Empty;
        }

        return new DocxTableLook(
            (string?)look.Attribute(WordprocessingNamespace + "val"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "firstRow"),
            (string?)look.Attribute(WordprocessingNamespace + "firstRow"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "lastRow"),
            (string?)look.Attribute(WordprocessingNamespace + "lastRow"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "firstColumn"),
            (string?)look.Attribute(WordprocessingNamespace + "firstColumn"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "lastColumn"),
            (string?)look.Attribute(WordprocessingNamespace + "lastColumn"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "noHBand"),
            (string?)look.Attribute(WordprocessingNamespace + "noHBand"),
            OoxBoolean.ParseOptionalAttribute(look, WordprocessingNamespace + "noVBand"),
            (string?)look.Attribute(WordprocessingNamespace + "noVBand"));
    }

    private static DocxTableCellConditionalFormat? ReadTableCellConditionalFormat(XElement? cellProperties)
    {
        XElement? conditional = cellProperties?.Element(WordprocessingNamespace + "cnfStyle");
        if (conditional is null)
        {
            return null;
        }

        return new DocxTableCellConditionalFormat(
            (string?)conditional.Attribute(WordprocessingNamespace + "val"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRow"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstRow"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRow"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastRow"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "oddHBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "oddHBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "evenHBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "evenHBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "oddVBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "oddVBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "evenVBand"),
            (string?)conditional.Attribute(WordprocessingNamespace + "evenVBand"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRowFirstColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstRowFirstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "firstRowLastColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "firstRowLastColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRowFirstColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastRowFirstColumn"),
            OoxBoolean.ParseOptionalAttribute(conditional, WordprocessingNamespace + "lastRowLastColumn"),
            (string?)conditional.Attribute(WordprocessingNamespace + "lastRowLastColumn"));
    }

    private static IReadOnlyList<DocxTableCellBorder> ReadBorderElements(XElement? borders)
    {
        if (borders is null)
        {
            return [];
        }

        return borders
            .Elements()
            .Where(border => border.Name.Namespace == WordprocessingNamespace)
            .Select(border => new DocxTableCellBorder(
                border.Name.LocalName,
                (string?)border.Attribute(WordprocessingNamespace + "val"),
                (string?)border.Attribute(WordprocessingNamespace + "color"),
                (string?)border.Attribute(WordprocessingNamespace + "sz")))
            .ToArray();
    }

    private static IReadOnlyList<DocxTableCellBorder> ResolveTableCellBorders(
        IReadOnlyList<DocxTableCellBorder> directBorders,
        IReadOnlyList<DocxTableCellBorder> styleBorders,
        IReadOnlyList<DocxTableCellBorder> styleTableBorders,
        IReadOnlyList<DocxTableCellBorder> tableBorders,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        var resolved = new Dictionary<string, DocxTableCellBorder>(StringComparer.OrdinalIgnoreCase);
        AddBorders(resolved, ResolveTableBordersForCell(styleTableBorders, rowIndex, cellIndex, rowCount, cellCount));
        AddBorders(resolved, styleBorders);
        AddBorders(resolved, ResolveTableBordersForCell(tableBorders, rowIndex, cellIndex, rowCount, cellCount));
        AddBorders(resolved, directBorders);
        return new[] { "top", "bottom", "left", "right" }
            .Where(resolved.ContainsKey)
            .Select(edge => resolved[edge])
            .Concat(resolved.Values.Where(border => !IsCanonicalCellBorderEdge(border.Edge)))
            .ToArray();
    }

    private static IEnumerable<DocxTableCellBorder> ResolveTableBordersForCell(
        IReadOnlyList<DocxTableCellBorder> tableBorders,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        DocxTableCellBorder? top = FindBorder(tableBorders, rowIndex == 0 ? "top" : "insideH");
        if (top is not null)
        {
            yield return top with { Edge = "top" };
        }

        DocxTableCellBorder? bottom = FindBorder(tableBorders, rowIndex == rowCount - 1 ? "bottom" : "insideH");
        if (bottom is not null)
        {
            yield return bottom with { Edge = "bottom" };
        }

        DocxTableCellBorder? left = cellIndex == 0
            ? FindBorder(tableBorders, "left") ?? FindBorder(tableBorders, "start")
            : FindBorder(tableBorders, "insideV");
        if (left is not null)
        {
            yield return left with { Edge = "left" };
        }

        DocxTableCellBorder? right = cellIndex == cellCount - 1
            ? FindBorder(tableBorders, "right") ?? FindBorder(tableBorders, "end")
            : FindBorder(tableBorders, "insideV");
        if (right is not null)
        {
            yield return right with { Edge = "right" };
        }
    }

    private static void AddBorders(Dictionary<string, DocxTableCellBorder> target, IEnumerable<DocxTableCellBorder> borders)
    {
        foreach (DocxTableCellBorder border in borders)
        {
            target[border.Edge] = border;
        }
    }

    private static DocxTableCellBorder? FindBorder(IReadOnlyList<DocxTableCellBorder> borders, string edge)
    {
        return borders.FirstOrDefault(border => string.Equals(border.Edge, edge, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsCanonicalCellBorderEdge(string edge)
    {
        return edge.Equals("top", StringComparison.OrdinalIgnoreCase) ||
            edge.Equals("bottom", StringComparison.OrdinalIgnoreCase) ||
            edge.Equals("left", StringComparison.OrdinalIgnoreCase) ||
            edge.Equals("right", StringComparison.OrdinalIgnoreCase);
    }

    private static double? ReadTableRowHeight(XElement? height)
    {
        if (height?.Attribute(WordprocessingNamespace + "val") is not { } value)
        {
            return null;
        }

        return OoxUnits.TwipsToPoints(long.Parse(value.Value, CultureInfo.InvariantCulture));
    }

    private static IReadOnlyList<DocxInlineImage> ReadInlineImages(
        XElement run,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision = null)
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
        DocxRevisionInfo? revision = null)
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
            OoxUnits.EmuToPoints(ParseLongAttribute(extent, "cx")),
            OoxUnits.EmuToPoints(ParseLongAttribute(extent, "cy")),
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

    private static DocxInlineImage? ReadVmlInlineImage(
        XElement shape,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxRevisionInfo? revision = null)
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

    private static DocxListLabel? CreateListLabel(XElement? paragraphProperties, DocxNumberingSet numbering, Dictionary<(string NumId, int Level), int> counters)
    {
        XElement? numberingProperties = paragraphProperties?.Element(WordprocessingNamespace + "numPr");
        string? numId = (string?)numberingProperties?
            .Element(WordprocessingNamespace + "numId")
            ?.Attribute(WordprocessingNamespace + "val");
        int level = numberingProperties?
            .Element(WordprocessingNamespace + "ilvl")
            ?.Attribute(WordprocessingNamespace + "val") is { } levelAttribute
            ? int.Parse(levelAttribute.Value, CultureInfo.InvariantCulture)
            : 0;
        if (numId is null || !numbering.NumToAbstract.TryGetValue(numId, out string? abstractId))
        {
            return null;
        }

        DocxNumberingLevel? numberingLevel = numbering.LevelOverrides.TryGetValue((numId, level), out DocxNumberingLevel? concreteLevel)
            ? concreteLevel
            : numbering.Levels.TryGetValue((abstractId, level), out DocxNumberingLevel? abstractLevel)
                ? abstractLevel
                : null;
        if (numberingLevel is null)
        {
            return null;
        }

        if (numberingLevel.Format.Equals("bullet", StringComparison.OrdinalIgnoreCase))
        {
            string bulletText = string.IsNullOrEmpty(numberingLevel.Text) ? "\u2022" : numberingLevel.Text;
            return new DocxListLabel(bulletText, numberingLevel.Format, numberingLevel.Text, numberingLevel.Suffix, numId, level, numberingLevel.Indent, numberingLevel.Style);
        }

        var key = (numId, level);
        int start = numbering.StartOverrides.TryGetValue((numId, level), out int overriddenStart)
            ? overriddenStart
            : numberingLevel.Start;
        counters[key] = counters.TryGetValue(key, out int current) ? current + 1 : start;
        foreach (var resetKey in counters.Keys.Where(k => k.NumId == numId && k.Level > level).ToArray())
        {
            counters.Remove(resetKey);
        }

        string labelText = ResolveNumberingLevelText(numberingLevel.Text, numId, counters);
        return new DocxListLabel(labelText, numberingLevel.Format, numberingLevel.Text, numberingLevel.Suffix, numId, level, numberingLevel.Indent, numberingLevel.Style);
    }

    private static string ResolveNumberingLevelText(string text, string numId, IReadOnlyDictionary<(string NumId, int Level), int> counters)
    {
        string resolved = text;
        for (int level = 0; level < 9; level++)
        {
            string token = "%" + (level + 1).ToString(CultureInfo.InvariantCulture);
            if (resolved.Contains(token, StringComparison.Ordinal))
            {
                string value = counters.TryGetValue((numId, level), out int counter)
                    ? counter.ToString(CultureInfo.InvariantCulture)
                    : "0";
                resolved = resolved.Replace(token, value, StringComparison.Ordinal);
            }
        }

        return resolved;
    }

    private static DocxStyleSet LoadStyles(OoxPackage package, string documentPartName, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? styleRelationship = package.GetRelationships(documentPartName, cancellationToken)
            .FirstOrDefault(r => !r.IsExternal && r.Type == StylesRelationshipType && r.ResolvedTarget is not null);
        OoxPart? stylesPart = styleRelationship?.ResolvedTarget is null
            ? package.Parts.FirstOrDefault(p => p.ContentType == StylesContentType)
            : package.GetPart(styleRelationship.ResolvedTarget);
        if (stylesPart is null)
        {
            return DocxStyleSet.Empty;
        }

        using Stream stream = stylesPart.OpenRead();
        XDocument stylesXml = SafeXml.Load(stream, cancellationToken);
        DocxResolvedRunProperties runDefaults = ReadRunProperties(stylesXml
            .Root?
            .Element(WordprocessingNamespace + "docDefaults")
            ?.Element(WordprocessingNamespace + "rPrDefault")
            ?.Element(WordprocessingNamespace + "rPr"));
        DocxResolvedParagraphProperties paragraphDefaults = ReadParagraphProperties(stylesXml
            .Root?
            .Element(WordprocessingNamespace + "docDefaults")
            ?.Element(WordprocessingNamespace + "pPrDefault")
            ?.Element(WordprocessingNamespace + "pPr"));

        var paragraphStyles = new Dictionary<string, DocxStyle>(StringComparer.Ordinal);
        var characterStyles = new Dictionary<string, DocxStyle>(StringComparer.Ordinal);
        var tableStyles = new Dictionary<string, DocxTableStyle>(StringComparer.Ordinal);
        string? defaultTableStyleId = null;
        foreach (XElement style in stylesXml.Root?.Elements(WordprocessingNamespace + "style") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? styleId = (string?)style.Attribute(WordprocessingNamespace + "styleId");
            string? type = (string?)style.Attribute(WordprocessingNamespace + "type");
            if (string.IsNullOrWhiteSpace(styleId))
            {
                continue;
            }

            var parsed = new DocxStyle(
                (string?)style.Element(WordprocessingNamespace + "basedOn")?.Attribute(WordprocessingNamespace + "val"),
                ReadParagraphProperties(style.Element(WordprocessingNamespace + "pPr")),
                ReadRunProperties(style.Element(WordprocessingNamespace + "rPr")));
            if (type == "paragraph")
            {
                paragraphStyles[styleId] = parsed;
            }
            else if (type == "character")
            {
                characterStyles[styleId] = parsed;
            }
            else if (type == "table")
            {
                tableStyles[styleId] = ReadTableStyle(style);
                if (OoxBoolean.ParseAttribute(style, WordprocessingNamespace + "default"))
                {
                    defaultTableStyleId = styleId;
                }
            }
        }

        IReadOnlyDictionary<string, DocxTableStyle> resolvedTableStyles = ResolveTableStyles(tableStyles);
        DocxTableStyle? defaultTableStyle = defaultTableStyleId is not null && resolvedTableStyles.TryGetValue(defaultTableStyleId, out DocxTableStyle? resolvedDefault)
            ? resolvedDefault
            : null;
        return new DocxStyleSet(runDefaults, paragraphDefaults, paragraphStyles, characterStyles, resolvedTableStyles, defaultTableStyleId, defaultTableStyle);
    }

    private static DocxStyleCatalog ToStyleCatalog(DocxStyleSet styles)
    {
        return new DocxStyleCatalog(
            styles.RunDefaults != DocxResolvedRunProperties.Empty,
            styles.ParagraphDefaults != DocxResolvedParagraphProperties.Empty,
            styles.DefaultTableStyleId,
            styles.ParagraphStyles
                .OrderBy(style => style.Key, StringComparer.Ordinal)
                .Select(style => ToStyleDefinitionSummary(style.Key, style.Value))
                .ToArray(),
            styles.CharacterStyles
                .OrderBy(style => style.Key, StringComparer.Ordinal)
                .Select(style => ToStyleDefinitionSummary(style.Key, style.Value))
                .ToArray(),
            styles.TableStyles
                .OrderBy(style => style.Key, StringComparer.Ordinal)
                .Select(style => ToTableStyleDefinitionSummary(style.Key, style.Value))
                .ToArray());
    }

    private static DocxStyleDefinitionSummary ToStyleDefinitionSummary(string styleId, DocxStyle style)
    {
        return new DocxStyleDefinitionSummary(
            styleId,
            style.BasedOnStyleId,
            style.Paragraph != DocxResolvedParagraphProperties.Empty,
            style.Run != DocxResolvedRunProperties.Empty);
    }

    private static DocxTableStyleDefinitionSummary ToTableStyleDefinitionSummary(string styleId, DocxTableStyle style)
    {
        return new DocxTableStyleDefinitionSummary(
            styleId,
            style.BasedOnStyleId,
            style.Table != DocxTableStyleProperties.Empty,
            style.Cell != DocxTableCellStyle.Empty,
            style.Cell.Paragraph != DocxResolvedParagraphProperties.Empty,
            style.Cell.Run != DocxResolvedRunProperties.Empty,
            style.TableBorders.Count,
            style.ConditionalRegions.Count);
    }

    private static DocxNumberingSet LoadNumbering(OoxPackage package, string documentPartName, DocxFontCatalog fontCatalog, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? numberingRelationship = package.GetRelationships(documentPartName, cancellationToken)
            .FirstOrDefault(r => !r.IsExternal && r.Type == NumberingRelationshipType && r.ResolvedTarget is not null);
        OoxPart? numberingPart = numberingRelationship?.ResolvedTarget is null
            ? package.Parts.FirstOrDefault(p => p.ContentType == NumberingContentType)
            : package.GetPart(numberingRelationship.ResolvedTarget);
        if (numberingPart is null)
        {
            return DocxNumberingSet.Empty;
        }

        using Stream stream = numberingPart.OpenRead();
        XDocument numberingXml = SafeXml.Load(stream, cancellationToken);
        var levels = new Dictionary<(string AbstractId, int Level), DocxNumberingLevel>();
        foreach (XElement abstractNum in numberingXml.Root?.Elements(WordprocessingNamespace + "abstractNum") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? abstractId = (string?)abstractNum.Attribute(WordprocessingNamespace + "abstractNumId");
            if (abstractId is null)
            {
                continue;
            }

            foreach (XElement level in abstractNum.Elements(WordprocessingNamespace + "lvl"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                int levelIndex = level.Attribute(WordprocessingNamespace + "ilvl") is { } ilvl
                    ? int.Parse(ilvl.Value, CultureInfo.InvariantCulture)
                    : 0;
                levels[(abstractId, levelIndex)] = ReadNumberingLevel(level, levelIndex, fontCatalog);
            }
        }

        var numToAbstract = new Dictionary<string, string>(StringComparer.Ordinal);
        var startOverrides = new Dictionary<(string NumId, int Level), int>();
        var levelOverrides = new Dictionary<(string NumId, int Level), DocxNumberingLevel>();
        foreach (XElement num in numberingXml.Root?.Elements(WordprocessingNamespace + "num") ?? [])
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? numId = (string?)num.Attribute(WordprocessingNamespace + "numId");
            string? abstractId = (string?)num.Element(WordprocessingNamespace + "abstractNumId")?.Attribute(WordprocessingNamespace + "val");
            if (numId is not null && abstractId is not null)
            {
                numToAbstract[numId] = abstractId;
            }

            if (numId is null)
            {
                continue;
            }

            foreach (XElement overrideLevel in num.Elements(WordprocessingNamespace + "lvlOverride"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                int levelIndex = overrideLevel.Attribute(WordprocessingNamespace + "ilvl") is { } ilvl
                    ? int.Parse(ilvl.Value, CultureInfo.InvariantCulture)
                    : 0;
                if (overrideLevel.Element(WordprocessingNamespace + "startOverride")?.Attribute(WordprocessingNamespace + "val") is { } startValue)
                {
                    startOverrides[(numId, levelIndex)] = int.Parse(startValue.Value, CultureInfo.InvariantCulture);
                }

                XElement? concreteLevel = overrideLevel.Element(WordprocessingNamespace + "lvl");
                if (concreteLevel is not null)
                {
                    levelOverrides[(numId, levelIndex)] = ReadNumberingLevel(concreteLevel, levelIndex, fontCatalog);
                }
            }
        }

        return new DocxNumberingSet(numToAbstract, levels, startOverrides, levelOverrides);
    }

    private static DocxNumberingLevel ReadNumberingLevel(XElement level, int levelIndex, DocxFontCatalog fontCatalog)
    {
        int start = level.Element(WordprocessingNamespace + "start")?.Attribute(WordprocessingNamespace + "val") is { } startValue
            ? int.Parse(startValue.Value, CultureInfo.InvariantCulture)
            : 1;
        string format = (string?)level.Element(WordprocessingNamespace + "numFmt")?.Attribute(WordprocessingNamespace + "val") ?? "decimal";
        string text = (string?)level.Element(WordprocessingNamespace + "lvlText")?.Attribute(WordprocessingNamespace + "val") ??
            (format.Equals("bullet", StringComparison.OrdinalIgnoreCase) ? "\u2022" : "%" + (levelIndex + 1) + ".");
        string suffix = (string?)level.Element(WordprocessingNamespace + "suff")?.Attribute(WordprocessingNamespace + "val") ?? "tab";
        DocxTextRunStyle style = ReadTextRunStyle(level.Element(WordprocessingNamespace + "rPr"));
        return new DocxNumberingLevel(format, ResolveNumberingSymbolText(text, style, fontCatalog), suffix, start, ReadNumberingIndent(level), style);
    }

    private static string ResolveNumberingSymbolText(string text, DocxTextRunStyle style, DocxFontCatalog fontCatalog)
    {
        return UsesSymbolCharset(style, fontCatalog) ? MapSymbolCharsetText(text) : text;
    }

    private static bool UsesSymbolCharset(DocxTextRunStyle style, DocxFontCatalog fontCatalog)
    {
        string? family = FirstNonEmpty(style.Fonts.Ascii, style.Fonts.HighAnsi, style.FontFamily, style.Fonts.ComplexScript);
        if (family is null)
        {
            return false;
        }

        DocxFontTableEntry? entry = fontCatalog.Entries
            .FirstOrDefault(item => item.Name.Equals(family, StringComparison.OrdinalIgnoreCase));
        return entry?.CharsetValue is { } charset &&
            (charset.Equals("02", StringComparison.OrdinalIgnoreCase) || charset.Equals("2", StringComparison.OrdinalIgnoreCase));
    }

    private static string MapSymbolCharsetText(string text)
    {
        Span<char> mapped = text.Length <= 256
            ? stackalloc char[text.Length]
            : new char[text.Length];
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            mapped[i] = ch <= 0x00FF ? (char)(0xF000 + ch) : ch;
        }

        return new string(mapped);
    }

    private static DocxResolvedParagraphProperties ResolveParagraphProperties(
        XElement? directProperties,
        string? paragraphStyleId,
        DocxStyleSet styles,
        DocxResolvedParagraphProperties? tableStyleProperties = null)
    {
        DocxResolvedParagraphProperties result = styles.ParagraphDefaults;
        foreach (DocxStyle style in EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles))
        {
            result = result.Merge(style.Paragraph);
        }

        if (tableStyleProperties is { } tableProperties)
        {
            result = result.Merge(tableProperties);
        }

        return result.Merge(ReadParagraphProperties(directProperties));
    }

    private static DocxParagraphStyleResolution CreateParagraphStyleResolution(
        XElement? directProperties,
        string? paragraphStyleId,
        DocxStyleSet styles,
        DocxResolvedParagraphProperties? tableStyleProperties)
    {
        DocxStyle[] styleChain = EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles).ToArray();
        return new DocxParagraphStyleResolution(
            paragraphStyleId,
            paragraphStyleId is not null && styleChain.Length != 0,
            styleChain.Length,
            styles.ParagraphDefaults != DocxResolvedParagraphProperties.Empty,
            HasDirectParagraphProperties(directProperties),
            tableStyleProperties is not null && tableStyleProperties.Value != DocxResolvedParagraphProperties.Empty);
    }

    private static bool HasDirectParagraphProperties(XElement? properties)
    {
        return properties?.Elements().Any(element =>
            element.Name != WordprocessingNamespace + "pStyle" &&
            element.Name != WordprocessingNamespace + "rPr") == true;
    }

    private static DocxResolvedRunProperties ResolveRunProperties(
        XElement? directProperties,
        string? paragraphStyleId,
        string? characterStyleId,
        DocxStyleSet styles,
        DocxResolvedRunProperties? tableStyleProperties = null)
    {
        DocxResolvedRunProperties result = styles.RunDefaults;
        if (tableStyleProperties is { } tableProperties)
        {
            result = result.Merge(tableProperties);
        }

        foreach (DocxStyle paragraphStyle in EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles))
        {
            result = result.Merge(paragraphStyle.Run);
        }

        foreach (DocxStyle characterStyle in EnumerateStyleInheritance(characterStyleId, styles.CharacterStyles))
        {
            result = result.Merge(characterStyle.Run);
        }

        return result.Merge(ReadRunProperties(directProperties));
    }

    private static DocxRunStyleResolution CreateRunStyleResolution(
        XElement? directProperties,
        string? paragraphStyleId,
        string? characterStyleId,
        DocxStyleSet styles,
        DocxResolvedRunProperties? tableStyleProperties = null)
    {
        DocxStyle[] paragraphStyleChain = EnumerateStyleInheritance(paragraphStyleId, styles.ParagraphStyles).ToArray();
        DocxStyle[] characterStyleChain = EnumerateStyleInheritance(characterStyleId, styles.CharacterStyles).ToArray();
        return new DocxRunStyleResolution(
            characterStyleId,
            characterStyleId is not null && characterStyleChain.Length != 0,
            characterStyleChain.Length,
            styles.RunDefaults != DocxResolvedRunProperties.Empty,
            paragraphStyleChain.Any(style => style.Run != DocxResolvedRunProperties.Empty),
            characterStyleChain.Any(style => style.Run != DocxResolvedRunProperties.Empty),
            HasDirectRunProperties(directProperties),
            tableStyleProperties is not null && tableStyleProperties.Value != DocxResolvedRunProperties.Empty);
    }

    private static bool HasDirectRunProperties(XElement? properties)
    {
        return properties?.Elements().Any(element => element.Name != WordprocessingNamespace + "rStyle") == true;
    }

    private static IEnumerable<DocxStyle> EnumerateStyleInheritance(string? styleId, IReadOnlyDictionary<string, DocxStyle> styles)
    {
        if (styleId is null)
        {
            yield break;
        }

        var chain = new Stack<DocxStyle>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? currentStyleId = styleId;
        while (currentStyleId is not null && seen.Add(currentStyleId) && styles.TryGetValue(currentStyleId, out DocxStyle? style))
        {
            chain.Push(style);
            currentStyleId = style.BasedOnStyleId;
        }

        while (chain.Count != 0)
        {
            yield return chain.Pop();
        }
    }

    private static DocxResolvedParagraphProperties ReadParagraphProperties(XElement? properties)
    {
        string? alignmentValue = ReadAlignmentValue(properties);
        DocxTextAlignment? alignment = ReadAlignment(alignmentValue);
        XElement? spacing = properties?.Element(WordprocessingNamespace + "spacing");
        double? before = ReadTwipsAttribute(spacing, WordprocessingNamespace + "before");
        double? after = ReadTwipsAttribute(spacing, WordprocessingNamespace + "after");
        double? lineFactor = null;
        double? linePoints = null;
        if (spacing?.Attribute(WordprocessingNamespace + "line") is { } line)
        {
            string? lineRule = (string?)spacing.Attribute(WordprocessingNamespace + "lineRule");
            if (string.Equals(lineRule, "exact", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(lineRule, "atLeast", StringComparison.OrdinalIgnoreCase))
            {
                linePoints = OoxUnits.TwipsToPoints(long.Parse(line.Value, CultureInfo.InvariantCulture));
            }
            else
            {
                lineFactor = int.Parse(line.Value, CultureInfo.InvariantCulture) / 240d;
            }
        }

        DocxParagraphSpacing paragraphSpacing = new(
            (string?)spacing?.Attribute(WordprocessingNamespace + "before"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "after"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "beforeLines"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "afterLines"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "beforeAutospacing"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "afterAutospacing"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "line"),
            (string?)spacing?.Attribute(WordprocessingNamespace + "lineRule"),
            ReadOnOff(properties?.Element(WordprocessingNamespace + "contextualSpacing")));
        DocxParagraphKeepRules keepRules = new(
            ReadOnOff(properties?.Element(WordprocessingNamespace + "keepNext")),
            (string?)properties?.Element(WordprocessingNamespace + "keepNext")?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(properties?.Element(WordprocessingNamespace + "keepLines")),
            (string?)properties?.Element(WordprocessingNamespace + "keepLines")?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(properties?.Element(WordprocessingNamespace + "widowControl")),
            (string?)properties?.Element(WordprocessingNamespace + "widowControl")?.Attribute(WordprocessingNamespace + "val"));
        DocxParagraphIndent indent = ReadParagraphIndent(properties);
        IReadOnlyList<DocxTabStop> tabStops = ReadParagraphTabStops(properties);
        XElement? snapToGrid = properties?.Element(WordprocessingNamespace + "snapToGrid");
        XElement? pageBreakBefore = properties?.Element(WordprocessingNamespace + "pageBreakBefore");
        XElement? wordWrap = properties?.Element(WordprocessingNamespace + "wordWrap");

        return new DocxResolvedParagraphProperties(
            alignment,
            alignmentValue,
            before,
            after,
            lineFactor,
            linePoints,
            paragraphSpacing,
            keepRules,
            indent,
            tabStops,
            ReadOnOff(snapToGrid),
            (string?)snapToGrid?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(pageBreakBefore),
            (string?)pageBreakBefore?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(wordWrap),
            (string?)wordWrap?.Attribute(WordprocessingNamespace + "val"));
    }

    private static IReadOnlyList<DocxTabStop> ReadParagraphTabStops(XElement? properties)
    {
        return properties?
            .Element(WordprocessingNamespace + "tabs")
            ?.Elements(WordprocessingNamespace + "tab")
            .Select(tab => new DocxTabStop(
                ReadTwipsAttribute(tab, WordprocessingNamespace + "pos"),
                (string?)tab.Attribute(WordprocessingNamespace + "pos"),
                (string?)tab.Attribute(WordprocessingNamespace + "val"),
                (string?)tab.Attribute(WordprocessingNamespace + "leader")))
            .ToArray() ?? [];
    }

    private static DocxParagraphIndent ReadParagraphIndent(XElement? properties)
    {
        XElement? indent = properties?.Element(WordprocessingNamespace + "ind");
        return new DocxParagraphIndent(
            ReadLogicalStartTwips(indent),
            ReadLogicalEndTwips(indent),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "firstLine"),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "hanging"),
            ReadLogicalStartValue(indent),
            ReadLogicalEndValue(indent),
            (string?)indent?.Attribute(WordprocessingNamespace + "firstLine"),
            (string?)indent?.Attribute(WordprocessingNamespace + "hanging"));
    }

    private static string? ReadAlignmentValue(XElement? properties)
    {
        return (string?)properties
            ?.Element(WordprocessingNamespace + "jc")
            ?.Attribute(WordprocessingNamespace + "val");
    }

    private static DocxTextAlignment? ReadAlignment(string? value)
    {
        return value switch
        {
            "center" => DocxTextAlignment.Center,
            "right" => DocxTextAlignment.Right,
            "both" => DocxTextAlignment.Justified,
            null => null,
            _ => DocxTextAlignment.Left
        };
    }

    private static DocxResolvedRunProperties ReadRunProperties(XElement? properties)
    {
        double? fontSize = properties?
            .Element(WordprocessingNamespace + "sz")
            ?.Attribute(WordprocessingNamespace + "val") is { } size
            ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 2d
            : null;
        string? color = (string?)properties?
            .Element(WordprocessingNamespace + "color")
            ?.Attribute(WordprocessingNamespace + "val");
        string? fontFamily = (string?)properties?
            .Element(WordprocessingNamespace + "rFonts")
            ?.Attribute(WordprocessingNamespace + "ascii");
        bool? bold = ReadOnOff(properties?.Element(WordprocessingNamespace + "b"));
        bool? italic = ReadOnOff(properties?.Element(WordprocessingNamespace + "i"));
        bool? complexScriptBold = ReadOnOff(properties?.Element(WordprocessingNamespace + "bCs"));
        bool? complexScriptItalic = ReadOnOff(properties?.Element(WordprocessingNamespace + "iCs"));
        bool? allCaps = ReadOnOff(properties?.Element(WordprocessingNamespace + "caps"));
        XElement? smallCapsElement = properties?.Element(WordprocessingNamespace + "smallCaps");
        bool? smallCaps = ReadOnOff(smallCapsElement);
        string? smallCapsValue = (string?)smallCapsElement?.Attribute(WordprocessingNamespace + "val");
        XElement? hiddenElement = properties?.Element(WordprocessingNamespace + "vanish");
        bool? hidden = ReadOnOff(hiddenElement);
        string? hiddenValue = (string?)hiddenElement?.Attribute(WordprocessingNamespace + "val");
        XElement? strikeElement = properties?.Element(WordprocessingNamespace + "strike");
        XElement? doubleStrikeElement = properties?.Element(WordprocessingNamespace + "dstrike");
        bool? strike = ReadOnOff(strikeElement);
        bool? doubleStrike = ReadOnOff(doubleStrikeElement);
        string? strikeValue = (string?)strikeElement?.Attribute(WordprocessingNamespace + "val");
        string? doubleStrikeValue = (string?)doubleStrikeElement?.Attribute(WordprocessingNamespace + "val");
        double? characterSpacingPoints = ReadSignedTwipsElement(properties?.Element(WordprocessingNamespace + "spacing"));
        string? verticalAlignmentValue = (string?)properties?
            .Element(WordprocessingNamespace + "vertAlign")
            ?.Attribute(WordprocessingNamespace + "val");
        string? highlightValue = (string?)properties?
            .Element(WordprocessingNamespace + "highlight")
            ?.Attribute(WordprocessingNamespace + "val");
        XElement? shading = properties?.Element(WordprocessingNamespace + "shd");
        string? shadingFill = (string?)shading?.Attribute(WordprocessingNamespace + "fill");
        string? shadingValue = (string?)shading?.Attribute(WordprocessingNamespace + "val");
        string? shadingColor = (string?)shading?.Attribute(WordprocessingNamespace + "color");
        XElement? underlineElement = properties?.Element(WordprocessingNamespace + "u");
        string? underlineValue = (string?)underlineElement?.Attribute(WordprocessingNamespace + "val");
        string? underlineColor = (string?)underlineElement?.Attribute(WordprocessingNamespace + "color");
        bool? underline = underlineElement is not null
            ? !string.Equals(underlineValue, "none", StringComparison.OrdinalIgnoreCase)
            : null;
        return new DocxResolvedRunProperties(fontSize, color, fontFamily, bold, italic, complexScriptBold, complexScriptItalic, underline, underlineValue, ReadRunFonts(properties), characterSpacingPoints, allCaps, verticalAlignmentValue, strike, strikeValue, doubleStrike, doubleStrikeValue, highlightValue, shadingFill, shadingValue, shadingColor, smallCaps, smallCapsValue, hidden, hiddenValue, underlineColor);
    }

    private static DocxRunFonts ReadRunFonts(XElement? properties)
    {
        XElement? fonts = properties?.Element(WordprocessingNamespace + "rFonts");
        return fonts is null
            ? DocxRunFonts.Empty
            : new DocxRunFonts(
                (string?)fonts.Attribute(WordprocessingNamespace + "ascii"),
                (string?)fonts.Attribute(WordprocessingNamespace + "hAnsi"),
                (string?)fonts.Attribute(WordprocessingNamespace + "eastAsia"),
                (string?)fonts.Attribute(WordprocessingNamespace + "cs"),
                (string?)fonts.Attribute(WordprocessingNamespace + "asciiTheme"),
                (string?)fonts.Attribute(WordprocessingNamespace + "hAnsiTheme"),
                (string?)fonts.Attribute(WordprocessingNamespace + "eastAsiaTheme"),
                (string?)fonts.Attribute(WordprocessingNamespace + "csTheme"));
    }

    private static DocxRunFonts MergeRunFonts(DocxRunFonts current, DocxRunFonts other)
    {
        return new DocxRunFonts(
            other.Ascii ?? current.Ascii,
            other.HighAnsi ?? current.HighAnsi,
            other.EastAsia ?? current.EastAsia,
            other.ComplexScript ?? current.ComplexScript,
            other.AsciiTheme ?? current.AsciiTheme,
            other.HighAnsiTheme ?? current.HighAnsiTheme,
            other.EastAsiaTheme ?? current.EastAsiaTheme,
            other.ComplexScriptTheme ?? current.ComplexScriptTheme);
    }

    private static bool? ReadOnOff(XElement? element)
    {
        if (element is null)
        {
            return null;
        }

        return OoxBoolean.ParseElement(element, valueAttributeName: WordprocessingNamespace + "val");
    }

    private static double? ReadTwipsAttribute(XElement? element, XName name)
    {
        return element?.Attribute(name) is { } value
            ? OoxUnits.TwipsToPoints(long.Parse(value.Value, CultureInfo.InvariantCulture))
            : null;
    }

    private static double? ReadSignedTwipsElement(XElement? element)
    {
        return int.TryParse((string?)element?.Attribute(WordprocessingNamespace + "val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips)
            ? OoxUnits.TwipsToPoints(twips)
            : null;
    }

    private static double? ReadLogicalStartTwips(XElement? indent)
    {
        return ReadTwipsAttribute(indent, WordprocessingNamespace + "start") ??
            ReadTwipsAttribute(indent, WordprocessingNamespace + "left");
    }

    private static double? ReadLogicalEndTwips(XElement? indent)
    {
        return ReadTwipsAttribute(indent, WordprocessingNamespace + "end") ??
            ReadTwipsAttribute(indent, WordprocessingNamespace + "right");
    }

    private static string? ReadLogicalStartValue(XElement? indent)
    {
        return (string?)indent?.Attribute(WordprocessingNamespace + "start") ??
            (string?)indent?.Attribute(WordprocessingNamespace + "left");
    }

    private static string? ReadLogicalEndValue(XElement? indent)
    {
        return (string?)indent?.Attribute(WordprocessingNamespace + "end") ??
            (string?)indent?.Attribute(WordprocessingNamespace + "right");
    }

    private static bool HasCharacterUnitIndent(XElement indent)
    {
        return indent.Attribute(WordprocessingNamespace + "leftChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "rightChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "startChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "endChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "firstLineChars") is not null ||
            indent.Attribute(WordprocessingNamespace + "hangingChars") is not null;
    }

    private static bool HasUnsupportedNumberingIndent(XElement level)
    {
        XElement? indent = level
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "ind");
        return indent is not null && HasCharacterUnitIndent(indent);
    }

    private static bool HasUnsupportedParagraphSpacingVariant(XElement spacing)
    {
        string? lineRule = (string?)spacing.Attribute(WordprocessingNamespace + "lineRule");
        return lineRule is not null &&
            !lineRule.Equals("auto", StringComparison.OrdinalIgnoreCase) &&
            !lineRule.Equals("exact", StringComparison.OrdinalIgnoreCase) &&
            !lineRule.Equals("atLeast", StringComparison.OrdinalIgnoreCase);
    }

    private static int? ReadPositiveIntAttribute(XElement? element, XName name)
    {
        if (element?.Attribute(name) is not { } value ||
            !int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ||
            parsed <= 0)
        {
            return null;
        }

        return parsed;
    }

    private static double ReadMargin(XElement? margins, XName name, double defaultValue)
    {
        return margins?.Attribute(name) is { } margin
            ? OoxUnits.TwipsToPoints(long.Parse(margin.Value, CultureInfo.InvariantCulture))
            : defaultValue;
    }

    private static OoxPart FindDocumentPart(OoxPackage package, CancellationToken cancellationToken = default)
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

    private static long ParseLongAttribute(XElement element, XName name)
    {
        string value = (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required DOCX attribute '{name.LocalName}'.");
        return long.Parse(value, CultureInfo.InvariantCulture);
    }

    private sealed record DocxStyleSet(
        DocxResolvedRunProperties RunDefaults,
        DocxResolvedParagraphProperties ParagraphDefaults,
        IReadOnlyDictionary<string, DocxStyle> ParagraphStyles,
        IReadOnlyDictionary<string, DocxStyle> CharacterStyles,
        IReadOnlyDictionary<string, DocxTableStyle> TableStyles,
        string? DefaultTableStyleId,
        DocxTableStyle? DefaultTableStyle)
    {
        public static DocxStyleSet Empty { get; } = new(
            new DocxResolvedRunProperties(null, null, null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null),
            new DocxResolvedParagraphProperties(null, null, null, null, null, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, DocxParagraphIndent.Empty, [], null, null, null, null, null, null),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxTableStyle>(),
            null,
            null);
    }

    private sealed record DocxStyle(string? BasedOnStyleId, DocxResolvedParagraphProperties Paragraph, DocxResolvedRunProperties Run);

    private sealed record ParagraphBreakPart(
        XElement? Paragraph,
        string? BreakValue,
        bool StartsAfterBreak,
        bool EndsBeforeBreak);

    private sealed record DocxTableStyle(
        string? BasedOnStyleId,
        DocxTableStyleProperties Table,
        DocxTableCellStyle Cell,
        IReadOnlyList<DocxTableCellBorder> TableBorders,
        IReadOnlyDictionary<string, DocxTableCellStyle> ConditionalRegions)
    {
        public static DocxTableStyle Empty { get; } = new(null, DocxTableStyleProperties.Empty, DocxTableCellStyle.Empty, [], new Dictionary<string, DocxTableCellStyle>());

        public DocxTableStyle Merge(DocxTableStyle other)
        {
            var conditional = new Dictionary<string, DocxTableCellStyle>(ConditionalRegions, StringComparer.Ordinal);
            foreach ((string region, DocxTableCellStyle regionStyle) in other.ConditionalRegions)
            {
                conditional[region] = conditional.TryGetValue(region, out DocxTableCellStyle? inherited)
                    ? inherited.Merge(regionStyle)
                    : regionStyle;
            }

            return new DocxTableStyle(
                other.BasedOnStyleId ?? BasedOnStyleId,
                Table.Merge(other.Table),
                Cell.Merge(other.Cell),
                other.TableBorders.Count == 0 ? TableBorders : other.TableBorders,
                conditional);
        }
    }

    private sealed record DocxTableStyleProperties(
        string? LayoutValue,
        double? PreferredWidthPoints,
        string? PreferredWidthValue,
        string? PreferredWidthType,
        double? IndentPoints,
        string? IndentValue,
        string? IndentType,
        double? CellSpacingPoints,
        string? CellSpacingValue,
        string? CellSpacingType,
        int? RowBandSize,
        int? ColumnBandSize)
    {
        public static DocxTableStyleProperties Empty { get; } = new(null, null, null, null, null, null, null, null, null, null, null, null);

        public DocxTableStyleProperties Merge(DocxTableStyleProperties other)
        {
            return new DocxTableStyleProperties(
                other.LayoutValue ?? LayoutValue,
                other.PreferredWidthPoints ?? PreferredWidthPoints,
                other.PreferredWidthValue ?? PreferredWidthValue,
                other.PreferredWidthType ?? PreferredWidthType,
                other.IndentPoints ?? IndentPoints,
                other.IndentValue ?? IndentValue,
                other.IndentType ?? IndentType,
                other.CellSpacingPoints ?? CellSpacingPoints,
                other.CellSpacingValue ?? CellSpacingValue,
                other.CellSpacingType ?? CellSpacingType,
                other.RowBandSize ?? RowBandSize,
                other.ColumnBandSize ?? ColumnBandSize);
        }
    }

    private sealed record DocxTableCellStyle(
        DocxResolvedParagraphProperties Paragraph,
        DocxResolvedRunProperties Run,
        string? FillHex,
        string? ShadingValue,
        string? ShadingColor,
        string? VerticalAlignmentValue,
        IReadOnlyList<DocxTableCellBorder> Borders,
        DocxTableCellMargins Margins,
        bool? NoWrap = null,
        string? NoWrapValue = null,
        bool? FitText = null,
        string? FitTextValue = null,
        string? TextDirectionValue = null)
    {
        public static DocxTableCellStyle Empty { get; } = new(DocxResolvedParagraphProperties.Empty, DocxResolvedRunProperties.Empty, null, null, null, null, [], DocxTableCellMargins.Empty);

        public DocxTableCellStyle Merge(DocxTableCellStyle other)
        {
            return new DocxTableCellStyle(
                Paragraph.Merge(other.Paragraph),
                Run.Merge(other.Run),
                other.FillHex ?? FillHex,
                other.ShadingValue ?? ShadingValue,
                other.ShadingColor ?? ShadingColor,
                other.VerticalAlignmentValue ?? VerticalAlignmentValue,
                other.Borders.Count == 0 ? Borders : other.Borders,
                MergeTableCellMargins(other.Margins, Margins),
                other.NoWrap ?? NoWrap,
                other.NoWrapValue ?? NoWrapValue,
                other.FitText ?? FitText,
                other.FitTextValue ?? FitTextValue,
                other.TextDirectionValue ?? TextDirectionValue);
        }
    }

    private static DocxTableStyle ReadTableStyle(XElement style)
    {
        var conditional = new Dictionary<string, DocxTableCellStyle>(StringComparer.Ordinal);
        foreach (XElement region in style.Elements(WordprocessingNamespace + "tblStylePr"))
        {
            string? type = (string?)region.Attribute(WordprocessingNamespace + "type");
            if (type is not null)
            {
                conditional[type] = ReadTableCellStyle(
                    region.Element(WordprocessingNamespace + "tcPr"),
                    region.Element(WordprocessingNamespace + "pPr"),
                    region.Element(WordprocessingNamespace + "rPr"));
            }
        }

        XElement? tableProperties = style.Element(WordprocessingNamespace + "tblPr");
        return new DocxTableStyle(
            (string?)style.Element(WordprocessingNamespace + "basedOn")?.Attribute(WordprocessingNamespace + "val"),
            ReadTableStyleProperties(tableProperties),
            ReadTableCellStyle(
                style.Element(WordprocessingNamespace + "tcPr"),
                style.Element(WordprocessingNamespace + "pPr"),
                style.Element(WordprocessingNamespace + "rPr")) with
            {
                Margins = ReadTableStyleCellMargins(tableProperties)
            },
            ReadTableBorders(tableProperties),
            conditional);
    }

    private static DocxTableStyleProperties ReadTableStyleProperties(XElement? tableProperties)
    {
        XElement? tableWidth = tableProperties?.Element(WordprocessingNamespace + "tblW");
        XElement? tableIndent = tableProperties?.Element(WordprocessingNamespace + "tblInd");
        XElement? tableCellSpacing = tableProperties?.Element(WordprocessingNamespace + "tblCellSpacing");
        return new DocxTableStyleProperties(
            (string?)tableProperties
                ?.Element(WordprocessingNamespace + "tblLayout")
                ?.Attribute(WordprocessingNamespace + "type"),
            ReadDxaWidth(tableWidth),
            (string?)tableWidth?.Attribute(WordprocessingNamespace + "w"),
            (string?)tableWidth?.Attribute(WordprocessingNamespace + "type"),
            ReadDxaWidth(tableIndent),
            (string?)tableIndent?.Attribute(WordprocessingNamespace + "w"),
            (string?)tableIndent?.Attribute(WordprocessingNamespace + "type"),
            ReadDxaWidth(tableCellSpacing),
            (string?)tableCellSpacing?.Attribute(WordprocessingNamespace + "w"),
            (string?)tableCellSpacing?.Attribute(WordprocessingNamespace + "type"),
            ReadPositiveIntAttribute(tableProperties?.Element(WordprocessingNamespace + "tblStyleRowBandSize"), WordprocessingNamespace + "val"),
            ReadPositiveIntAttribute(tableProperties?.Element(WordprocessingNamespace + "tblStyleColBandSize"), WordprocessingNamespace + "val"));
    }

    private static IReadOnlyDictionary<string, DocxTableStyle> ResolveTableStyles(IReadOnlyDictionary<string, DocxTableStyle> tableStyles)
    {
        var resolved = new Dictionary<string, DocxTableStyle>(StringComparer.Ordinal);
        foreach (string styleId in tableStyles.Keys)
        {
            resolved[styleId] = ResolveTableStyle(styleId, tableStyles);
        }

        return resolved;
    }

    private static DocxTableStyle ResolveTableStyle(string styleId, IReadOnlyDictionary<string, DocxTableStyle> tableStyles)
    {
        var chain = new Stack<DocxTableStyle>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? currentStyleId = styleId;
        while (currentStyleId is not null && seen.Add(currentStyleId) && tableStyles.TryGetValue(currentStyleId, out DocxTableStyle? style))
        {
            chain.Push(style);
            currentStyleId = style.BasedOnStyleId;
        }

        DocxTableStyle resolved = DocxTableStyle.Empty;
        while (chain.Count != 0)
        {
            resolved = resolved.Merge(chain.Pop());
        }

        return resolved with
        {
            BasedOnStyleId = tableStyles.TryGetValue(styleId, out DocxTableStyle? source)
                ? source.BasedOnStyleId
                : null
        };
    }

    private static DocxTableCellStyle ReadTableCellStyle(
        XElement? cellProperties,
        XElement? paragraphProperties = null,
        XElement? runProperties = null)
    {
        XElement? shading = cellProperties?.Element(WordprocessingNamespace + "shd");
        XElement? noWrap = cellProperties?.Element(WordprocessingNamespace + "noWrap");
        XElement? fitText = cellProperties?.Element(WordprocessingNamespace + "tcFitText");
        return new DocxTableCellStyle(
            ReadParagraphProperties(paragraphProperties),
            ReadRunProperties(runProperties),
            (string?)shading?.Attribute(WordprocessingNamespace + "fill"),
            (string?)shading?.Attribute(WordprocessingNamespace + "val"),
            (string?)shading?.Attribute(WordprocessingNamespace + "color"),
            (string?)cellProperties
                ?.Element(WordprocessingNamespace + "vAlign")
                ?.Attribute(WordprocessingNamespace + "val"),
            ReadTableCellBorders(cellProperties),
            DocxTableCellMargins.Empty,
            ReadOnOff(noWrap),
            (string?)noWrap?.Attribute(WordprocessingNamespace + "val"),
            ReadOnOff(fitText),
            (string?)fitText?.Attribute(WordprocessingNamespace + "val"),
            (string?)cellProperties
                ?.Element(WordprocessingNamespace + "textDirection")
                ?.Attribute(WordprocessingNamespace + "val"));
    }

    private static DocxTableCellStyle ResolveTableCellStyle(
        DocxTableStyle tableStyle,
        DocxTableLook tableLook,
        DocxTableCellConditionalFormat? conditionalFormat,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        DocxTableCellStyle resolved = tableStyle.Cell;
        IEnumerable<string> regions = conditionalFormat?.IsDefined == true
            ? EnumerateTableStyleRegions(conditionalFormat)
            : EnumerateTableStyleRegions(tableLook, tableStyle.Table.RowBandSize, tableStyle.Table.ColumnBandSize, rowIndex, cellIndex, rowCount, cellCount);
        foreach (string region in regions)
        {
            if (tableStyle.ConditionalRegions.TryGetValue(region, out DocxTableCellStyle? style))
            {
                resolved = resolved.Merge(style);
            }
        }

        return resolved;
    }

    private static IEnumerable<string> EnumerateTableStyleRegions(DocxTableCellConditionalFormat conditionalFormat)
    {
        if (conditionalFormat.FirstRow == true)
        {
            yield return "firstRow";
        }

        if (conditionalFormat.LastRow == true)
        {
            yield return "lastRow";
        }

        if (conditionalFormat.FirstColumn == true)
        {
            yield return "firstCol";
        }

        if (conditionalFormat.LastColumn == true)
        {
            yield return "lastCol";
        }

        if (conditionalFormat.FirstRowFirstColumn == true)
        {
            yield return "nwCell";
        }

        if (conditionalFormat.FirstRowLastColumn == true)
        {
            yield return "neCell";
        }

        if (conditionalFormat.LastRowFirstColumn == true)
        {
            yield return "swCell";
        }

        if (conditionalFormat.LastRowLastColumn == true)
        {
            yield return "seCell";
        }

        if (conditionalFormat.OddHorizontalBand == true)
        {
            yield return "band1Horz";
        }

        if (conditionalFormat.EvenHorizontalBand == true)
        {
            yield return "band2Horz";
        }

        if (conditionalFormat.OddVerticalBand == true)
        {
            yield return "band1Vert";
        }

        if (conditionalFormat.EvenVerticalBand == true)
        {
            yield return "band2Vert";
        }
    }

    private static IEnumerable<string> EnumerateTableStyleRegions(
        DocxTableLook tableLook,
        int? rowBandSize,
        int? columnBandSize,
        int rowIndex,
        int cellIndex,
        int rowCount,
        int cellCount)
    {
        bool firstRow = tableLook.FirstRow != false;
        bool lastRow = tableLook.LastRow == true;
        bool firstColumn = tableLook.FirstColumn == true;
        bool lastColumn = tableLook.LastColumn == true;
        bool horizontalBand = tableLook.NoHorizontalBand != true;
        bool verticalBand = tableLook.NoVerticalBand != true;

        if (firstRow && rowIndex == 0)
        {
            yield return "firstRow";
        }

        if (lastRow && rowIndex == rowCount - 1)
        {
            yield return "lastRow";
        }

        if (firstColumn && cellIndex == 0)
        {
            yield return "firstCol";
        }

        if (lastColumn && cellIndex == cellCount - 1)
        {
            yield return "lastCol";
        }

        if (firstRow && firstColumn && rowIndex == 0 && cellIndex == 0)
        {
            yield return "nwCell";
        }

        if (firstRow && lastColumn && rowIndex == 0 && cellIndex == cellCount - 1)
        {
            yield return "neCell";
        }

        if (lastRow && firstColumn && rowIndex == rowCount - 1 && cellIndex == 0)
        {
            yield return "swCell";
        }

        if (lastRow && lastColumn && rowIndex == rowCount - 1 && cellIndex == cellCount - 1)
        {
            yield return "seCell";
        }

        string? horizontalBandRegion = ResolveBandRegion(rowIndex, rowBandSize ?? 1, "band1Horz", "band2Horz");
        if (horizontalBand && horizontalBandRegion is not null)
        {
            yield return horizontalBandRegion;
        }

        string? verticalBandRegion = ResolveBandRegion(cellIndex, columnBandSize ?? 1, "band1Vert", "band2Vert");
        if (verticalBand && verticalBandRegion is not null)
        {
            yield return verticalBandRegion;
        }
    }

    private static string? ResolveBandRegion(int index, int bandSize, string firstBand, string secondBand)
    {
        if (index == 0)
        {
            return null;
        }

        int effectiveBandSize = Math.Max(1, bandSize);
        int bandIndex = (index - 1) / effectiveBandSize;
        return bandIndex % 2 == 0 ? firstBand : secondBand;
    }

    private sealed record DocxNumberingSet(
        IReadOnlyDictionary<string, string> NumToAbstract,
        IReadOnlyDictionary<(string AbstractId, int Level), DocxNumberingLevel> Levels,
        IReadOnlyDictionary<(string NumId, int Level), int> StartOverrides,
        IReadOnlyDictionary<(string NumId, int Level), DocxNumberingLevel> LevelOverrides)
    {
        public static DocxNumberingSet Empty { get; } = new(
            new Dictionary<string, string>(),
            new Dictionary<(string AbstractId, int Level), DocxNumberingLevel>(),
            new Dictionary<(string NumId, int Level), int>(),
            new Dictionary<(string NumId, int Level), DocxNumberingLevel>());
    }

    private sealed record DocxNumberingLevel(string Format, string Text, string Suffix, int Start, DocxNumberingIndent Indent, DocxTextRunStyle Style);

    private static DocxTextRunStyle ReadTextRunStyle(XElement? properties)
    {
        DocxResolvedRunProperties run = ReadRunProperties(properties);
        return new DocxTextRunStyle(
            run.FontSize,
            run.ColorHex,
            run.Bold,
            run.Italic,
            run.Underline,
            run.UnderlineValue,
            run.FontFamily,
            run.Fonts,
            run.CharacterSpacingPoints,
            run.AllCaps,
            run.VerticalAlignmentValue,
            run.Strike,
            run.StrikeValue,
            run.DoubleStrike,
            run.DoubleStrikeValue,
            run.HighlightValue,
            run.ShadingFillHex,
            run.ShadingValue,
            run.ShadingColor,
            run.SmallCaps,
            run.SmallCapsValue,
            run.Hidden,
            run.HiddenValue,
            run.UnderlineColorHex);
    }

    private static DocxNumberingIndent ReadNumberingIndent(XElement level)
    {
        XElement? indent = level
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "ind");
        XElement? numberingTab = level
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "tabs")
            ?.Elements(WordprocessingNamespace + "tab")
            .FirstOrDefault(tab => string.Equals(
                (string?)tab.Attribute(WordprocessingNamespace + "val"),
                "num",
                StringComparison.OrdinalIgnoreCase));
        return new DocxNumberingIndent(
            ReadLogicalStartTwips(indent),
            ReadLogicalEndTwips(indent),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "firstLine"),
            ReadTwipsAttribute(indent, WordprocessingNamespace + "hanging"),
            ReadTwipsAttribute(numberingTab, WordprocessingNamespace + "pos"),
            ReadLogicalStartValue(indent),
            ReadLogicalEndValue(indent),
            (string?)indent?.Attribute(WordprocessingNamespace + "firstLine"),
            (string?)indent?.Attribute(WordprocessingNamespace + "hanging"),
            (string?)numberingTab?.Attribute(WordprocessingNamespace + "val"),
            (string?)numberingTab?.Attribute(WordprocessingNamespace + "pos"));
    }

    private readonly record struct DocxResolvedParagraphProperties(
        DocxTextAlignment? Alignment,
        string? AlignmentValue,
        double? SpacingBeforePoints,
        double? SpacingAfterPoints,
        double? LineSpacingFactor,
        double? LineSpacingPoints,
        DocxParagraphSpacing Spacing,
        DocxParagraphKeepRules KeepRules,
        DocxParagraphIndent Indent,
        IReadOnlyList<DocxTabStop> TabStops,
        bool? SnapToGrid,
        string? SnapToGridValue,
        bool? PageBreakBefore,
        string? PageBreakBeforeValue,
        bool? WordWrap,
        string? WordWrapValue)
    {
        public static DocxResolvedParagraphProperties Empty { get; } = new(null, null, null, null, null, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, DocxParagraphIndent.Empty, [], null, null, null, null, null, null);

        public DocxResolvedParagraphProperties Merge(DocxResolvedParagraphProperties other)
        {
            bool hasOtherBeforeSide = HasBeforeSpacingSide(other.Spacing);
            bool hasOtherAfterSide = HasAfterSpacingSide(other.Spacing);
            return new DocxResolvedParagraphProperties(
                other.Alignment ?? Alignment,
                other.AlignmentValue ?? AlignmentValue,
                hasOtherBeforeSide ? other.SpacingBeforePoints : other.SpacingBeforePoints ?? SpacingBeforePoints,
                hasOtherAfterSide ? other.SpacingAfterPoints : other.SpacingAfterPoints ?? SpacingAfterPoints,
                other.LineSpacingFactor ?? LineSpacingFactor,
                other.LineSpacingPoints ?? LineSpacingPoints,
                MergeSpacing(Spacing, other.Spacing),
                MergeKeepRules(KeepRules, other.KeepRules),
                MergeIndent(Indent, other.Indent),
                other.TabStops.Count != 0 ? other.TabStops : TabStops,
                other.SnapToGrid ?? SnapToGrid,
                other.SnapToGridValue ?? SnapToGridValue,
                other.PageBreakBefore ?? PageBreakBefore,
                other.PageBreakBeforeValue ?? PageBreakBeforeValue,
                other.WordWrap ?? WordWrap,
                other.WordWrapValue ?? WordWrapValue);
        }
    }

    private static DocxParagraphSpacing MergeSpacing(DocxParagraphSpacing current, DocxParagraphSpacing other)
    {
        bool hasOtherBeforeSide = HasBeforeSpacingSide(other);
        bool hasOtherAfterSide = HasAfterSpacingSide(other);
        return new DocxParagraphSpacing(
            hasOtherBeforeSide ? other.BeforeValue : current.BeforeValue,
            hasOtherAfterSide ? other.AfterValue : current.AfterValue,
            hasOtherBeforeSide ? other.BeforeLinesValue : current.BeforeLinesValue,
            hasOtherAfterSide ? other.AfterLinesValue : current.AfterLinesValue,
            hasOtherBeforeSide ? other.BeforeAutoSpacingValue : current.BeforeAutoSpacingValue,
            hasOtherAfterSide ? other.AfterAutoSpacingValue : current.AfterAutoSpacingValue,
            other.LineValue ?? current.LineValue,
            other.LineRuleValue ?? current.LineRuleValue,
            other.ContextualSpacing ?? current.ContextualSpacing);
    }

    private static bool HasBeforeSpacingSide(DocxParagraphSpacing spacing)
    {
        return spacing.BeforeValue is not null ||
            spacing.BeforeLinesValue is not null ||
            spacing.BeforeAutoSpacingValue is not null;
    }

    private static bool HasAfterSpacingSide(DocxParagraphSpacing spacing)
    {
        return spacing.AfterValue is not null ||
            spacing.AfterLinesValue is not null ||
            spacing.AfterAutoSpacingValue is not null;
    }

    private static DocxParagraphKeepRules MergeKeepRules(DocxParagraphKeepRules current, DocxParagraphKeepRules other)
    {
        return new DocxParagraphKeepRules(
            other.KeepNext ?? current.KeepNext,
            other.KeepNextValue ?? current.KeepNextValue,
            other.KeepLines ?? current.KeepLines,
            other.KeepLinesValue ?? current.KeepLinesValue,
            other.WidowControl ?? current.WidowControl,
            other.WidowControlValue ?? current.WidowControlValue);
    }

    private static DocxParagraphIndent MergeIndent(DocxParagraphIndent current, DocxParagraphIndent other)
    {
        bool hasOtherFirstLineSide = other.FirstLineValue is not null || other.HangingValue is not null;
        return new DocxParagraphIndent(
            other.LeftPoints ?? current.LeftPoints,
            other.RightPoints ?? current.RightPoints,
            hasOtherFirstLineSide ? other.FirstLinePoints : current.FirstLinePoints,
            hasOtherFirstLineSide ? other.HangingPoints : current.HangingPoints,
            other.LeftValue ?? current.LeftValue,
            other.RightValue ?? current.RightValue,
            hasOtherFirstLineSide ? other.FirstLineValue : current.FirstLineValue,
            hasOtherFirstLineSide ? other.HangingValue : current.HangingValue);
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
        string? VerticalAlignmentValue = null,
        bool? Strike = null,
        string? StrikeValue = null,
        bool? DoubleStrike = null,
        string? DoubleStrikeValue = null,
        string? HighlightValue = null,
        string? ShadingFillHex = null,
        string? ShadingValue = null,
        string? ShadingColor = null,
        bool? SmallCaps = null,
        string? SmallCapsValue = null,
        bool? Hidden = null,
        string? HiddenValue = null,
        string? UnderlineColorHex = null)
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

    private sealed record DocxRevisionScopedElement(XElement Element, DocxRevisionInfo? Revision);

    private sealed record DocxCommentRangeStart(string? Id, int SourceRunIndex, int TextOffset);

    private sealed record DocxCommentAnchorInventory(
        IReadOnlyList<string> PackageAnchorIds,
        IReadOnlyList<string> HiddenAnchorIds);

    private sealed record DocxRevisionRangeStart(
        string Kind,
        string? Id,
        string? Name,
        string? Author,
        string? Date,
        int SourceRunIndex,
        int TextOffset);

    private sealed record DocxCommentThreadMetadata(string? ParentParagraphId, bool? IsResolved);
}
