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
        XElement? textBoxBodyProperties = anchor
            .Descendants(WordprocessingShapeNamespace + "bodyPr")
            .FirstOrDefault();
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
            sourceBlockIndex,
            (string?)textBoxBodyProperties?.Attribute("lIns"),
            (string?)textBoxBodyProperties?.Attribute("tIns"),
            (string?)textBoxBodyProperties?.Attribute("rIns"),
            (string?)textBoxBodyProperties?.Attribute("bIns"))
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
}
