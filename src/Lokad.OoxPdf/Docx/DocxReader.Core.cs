using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pptx;
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
        XDocument document = package.LoadXml(documentPart, cancellationToken);
        IMemoryReservation? modelStorage = OoxMemory.ReserveModel(document, cancellationToken);
        DocxDocument? result = null;
        try
        {
            IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(documentPart.Name, cancellationToken)
                .ToDictionary(r => r.Id, StringComparer.Ordinal);
            IReadOnlyDictionary<string, OoxRelationship> internalRelationships = relationships.Values
                .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
                .ToDictionary(r => r.Id, StringComparer.Ordinal);
            var warnedMustUnderstandParts = new HashSet<string>(StringComparer.Ordinal);
            XDocument? settings = LoadRelatedXmlPart(package, documentPart.Name, SettingsRelationshipType, out _, cancellationToken, diagnosticSink, warnedMustUnderstandParts, modelStorage);
            DocxDocumentSettings documentSettings = ReadDocumentSettings(settings, cancellationToken);
            OoxPdfDocxMarkupMode revisionFilteringMarkupMode = ResolveRevisionFilteringMarkupMode(markupMode, documentSettings);
            DocxMarkupContext markupContext = DocxMarkupContext.FromMode(markupMode).ApplyDocumentSettings(documentSettings);
            DocxCommentAnchorInventory commentAnchorInventory = ReadCommentAnchorInventory(document, revisionFilteringMarkupMode);
            var equationDiagnostics = new List<OoxPdfDiagnostic>();
            int equationSourceCount = CountSelectedEquationSources(document, revisionFilteringMarkupMode, cancellationToken, out _);
            Action<OoxPdfDiagnostic>? featureSink = diagnosticSink;
            if (diagnosticSink is not null || equationSourceCount != 0 || document.Descendants(MathNamespace + "oMathPara").Any())
            {
                featureSink = diagnostic =>
                {
                    // Font availability and frame admission belong to rendering. Preserve
                    // probes until that stage can distinguish painted and ignored math.
                    if (diagnostic.Id == "DOCX_UNSUPPORTED_EQUATION") { equationDiagnostics.Add(diagnostic); }
                    else { diagnosticSink?.Invoke(diagnostic); }
                };
            }
            EmitUnsupportedFeatureDiagnostics(package, document, documentPart.Name, relationships, markupContext, revisionFilteringMarkupMode, featureSink, warnedMustUnderstandParts, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();

            // Current sections belong to the body or a paragraph's properties.
            // Revision snapshots also contain sectPr, but do not own page geometry.
            XElement? sectionProperties = document.Descendants(WordprocessingNamespace + "sectPr")
                .LastOrDefault(section => section.Parent?.Name == WordprocessingNamespace + "body" ||
                    (section.Parent?.Name == WordprocessingNamespace + "pPr" &&
                     section.Parent?.Parent?.Name == WordprocessingNamespace + "p"));
            XElement? pageSize = sectionProperties?.Element(WordprocessingNamespace + "pgSz");
            XElement? pageMargins = sectionProperties?.Element(WordprocessingNamespace + "pgMar");

            DocxFontCatalog fontCatalog = LoadFontCatalog(package, documentPart.Name, cancellationToken, diagnosticSink, warnedMustUnderstandParts, modelStorage);
            DocxStyleSet styles = LoadStyles(package, documentPart.Name, cancellationToken, diagnosticSink, warnedMustUnderstandParts, modelStorage);
            DocxNumberingSet numbering = LoadNumbering(package, documentPart.Name, fontCatalog, cancellationToken, diagnosticSink, warnedMustUnderstandParts, modelStorage);
            cancellationToken.ThrowIfCancellationRequested();
            // one related-story cache per Read, shared by document headers/footers,
            // body section breaks, and every section's page settings.
            var storyCache = new RelatedStoryPartCache(modelStorage, documentSettings);
            DocxSectionBreakElement? finalSectionBreak = sectionProperties is null
                ? null
                : ReadSectionBreak(sectionProperties, package, internalRelationships, styles, numbering, settings, revisionFilteringMarkupMode, cancellationToken, storyCache, null);
            Dictionary<XElement, int>? floatingBlocks = CreateFloatingBlockMap(document);
            IReadOnlyList<DocxBodyElement> bodyElements = ReadBodyElements(document, styles, numbering, package, relationships, settings, documentSettings, revisionFilteringMarkupMode, cancellationToken, storyCache, floatingBlocks);
            // An empty final section has a Word paragraph mark even when its
            // paragraph is absent from the XML. Resolve its defaults through
            // the same reader as an authored paragraph; placement decides use.
            DocxParagraph? emptyFinalSectionParagraph = bodyElements.LastOrDefault() is DocxSectionBreakElement &&
                finalSectionBreak is { TypeValue: null or DocxSectionBreakType.NextPage }
                ? ReadParagraph(new XElement(WordprocessingNamespace + "p"), styles, numbering,
                    new DocxNumberingCounters(), package, relationships, null, null, documentSettings,
                    null, revisionFilteringMarkupMode, cancellationToken) : null;
            var (headerBodyElementsByType, headerDrawingsByType) = ReadReferencedStaticStories(
                document, package, internalRelationships, styles, numbering, HeaderRelationshipType, "headerReference",
                revisionFilteringMarkupMode, cancellationToken, storyCache, diagnosticSink, warnedMustUnderstandParts);
            var (footerBodyElementsByType, footerDrawingsByType) = ReadReferencedStaticStories(
                document, package, internalRelationships, styles, numbering, FooterRelationshipType, "footerReference",
                revisionFilteringMarkupMode, cancellationToken, storyCache, diagnosticSink, warnedMustUnderstandParts);
            IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = ToStaticParagraphsByType(headerBodyElementsByType);
            IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = ToStaticParagraphsByType(footerBodyElementsByType);
            IReadOnlyList<DocxParagraph> headers = SelectDefaultHeaderFooterParagraphs(headersByType);
            IReadOnlyList<DocxParagraph> footers = SelectDefaultHeaderFooterParagraphs(footersByType);
            IReadOnlyList<DocxRelatedStory> relatedStories = ReadRelatedStories(package, documentPart.Name, styles, numbering, documentSettings, revisionFilteringMarkupMode, cancellationToken, diagnosticSink, warnedMustUnderstandParts, bodyElements, settings, modelStorage);
            IReadOnlyList<DocxFloatingDrawing> floatingDrawings = floatingBlocks is null ? [] :
                ReadFloatingDrawings(document, package, relationships, styles, numbering, documentSettings, revisionFilteringMarkupMode, cancellationToken, floatingBlocks);
            cancellationToken.ThrowIfCancellationRequested();

            IReadOnlyList<DocxEquationPartSource> ReadEquationParts()
            {
                var probedParts = equationDiagnostics.Select(d => d.PartName).ToHashSet(StringComparer.Ordinal);
                var parts = new[] { new DocxEquationPartSource(documentPart.Name, equationSourceCount, bodyElements) { FloatingDrawings = floatingDrawings } }
                    .Concat(storyCache.EnumerateEquationParts(revisionFilteringMarkupMode, cancellationToken))
                    .Concat(relatedStories.Where(story => story.Kind is DocxRelatedStoryKind.Footnote or DocxRelatedStoryKind.Endnote)
                        .GroupBy(story => story.PartName, StringComparer.Ordinal)
                        .Where(group => group.Any(story => story.SourceEquationCount != 0) || probedParts.Contains(group.Key))
                        .Select(group => new DocxEquationPartSource(group.Key, group.Sum(story => story.SourceEquationCount),
                            group.Where(story => story.IsNormalStoryType).SelectMany(story => story.BodyElements).ToArray())
                        {
                            FloatingDrawings = group.Where(story => story.IsNormalStoryType).SelectMany(story => story.FloatingDrawings).ToArray()
                        }))
                    .ToArray();
                // Sink-less readers retain story-only equation warnings from the
                // already parsed story cache, without enabling unrelated probes.
                foreach (DocxEquationPartSource part in parts)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (part.SourceCount == 0 || !probedParts.Add(part.PartName)) { continue; }
                    equationDiagnostics.Add(new OoxPdfDiagnostic("DOCX_UNSUPPORTED_EQUATION", OoxPdfSeverity.Warning,
                        "Unsupported DOCX feature 'equation' was detected and ignored or approximated.", part.PartName,
                        SlideIndex: null, PageIndex: null, Feature: "equation", Fallback: "Ignored"));
                }
                return parts;
            }

            double width = 612d;
            double height = 792d;
            double left = 72d;
            double right = 72d;
            double top = 72d;
            double bottom = 72d;
            if (pageSize is not null)
            {
                width = OoxUnits.TwipsToPoints(OoxXml.ParseRequiredLong(pageSize, WordprocessingNamespace + "w", "DOCX"));
                height = OoxUnits.TwipsToPoints(OoxXml.ParseRequiredLong(pageSize, WordprocessingNamespace + "h", "DOCX"));
                (width, height) = DocxPageSettings.NormalizeDimensions(width, height, (string?)pageSize.Attribute(WordprocessingNamespace + "orient"));

                left = ReadMargin(pageMargins, WordprocessingNamespace + "left", left);
                right = ReadMargin(pageMargins, WordprocessingNamespace + "right", right);
                top = ReadMargin(pageMargins, WordprocessingNamespace + "top", top);
                bottom = ReadMargin(pageMargins, WordprocessingNamespace + "bottom", bottom);
            }

            cancellationToken.ThrowIfCancellationRequested();
            DocxPageSettings pageSettings = finalSectionBreak?.PageSettings ?? ReadPageSettings(
                null, settings, package, internalRelationships, styles, numbering, revisionFilteringMarkupMode,
                cancellationToken, storyCache);
            return result = new DocxDocument(
                width,
                height,
                left,
                right,
                top,
                bottom,
                pageSettings,
                floatingDrawings,
                headers,
                footers,
                bodyElements,
                [],
                [])
            {
                FontCatalog = fontCatalog,
                MainPartName = documentPart.Name,
                EquationParts = ReadEquationParts(),
                EquationDiagnostics = equationDiagnostics,
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
                FinalSectionBreak = finalSectionBreak,
                EmptyFinalSectionParagraph = emptyFinalSectionParagraph
            };
        }
        finally
        {
            if (result is null) modelStorage?.Dispose();
            else if (modelStorage is not null) OoxMemory.Resize(modelStorage, modelStorage.Bytes / 2);
        }
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

    private static DocxDocumentSettings ReadDocumentSettings(XDocument? settings, CancellationToken cancellationToken)
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
            compatSettings)
        {
            MathFontFamily = (string?)root.Element(MathNamespace + "mathPr")?.Element(MathNamespace + "mathFont")?.Attribute(MathNamespace + "val"),
            MathParagraphAlignmentValue = (string?)root.Element(MathNamespace + "mathPr")?.Element(MathNamespace + "defJc")?.Attribute(MathNamespace + "val"),
            MathLimitSettings = DocxMathLimitSettings.Read(root, cancellationToken)
        };
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
            bool IsCommentAnchorElement(XElement element)
            {
                return element.Name == WordprocessingNamespace + "commentReference" ||
                    element.Name == WordprocessingNamespace + "commentRangeStart" ||
                    element.Name == WordprocessingNamespace + "commentRangeEnd";
            }

            string? ReadCommentAnchorId(XElement element)
            {
                return NormalizeDecimalId((string?)element.Attribute(WordprocessingNamespace + "id"));
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
        DocxDocumentSettings? documentSettings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        IReadOnlyDictionary<XElement, int>? floatingBlocks = null) =>
        ReadFloatingDrawings(document, document.Root?
            .Element(WordprocessingNamespace + "body")?
            .Elements()
            .Where(static block => IsBodyBlockElement(block) || IsRevisionContainer(block)) ?? [],
            package, relationships, styles, numbering, documentSettings, markupMode, cancellationToken, floatingBlocks);

    private static IReadOnlyList<DocxFloatingDrawing> ReadFloatingDrawings(
        XContainer source,
        IEnumerable<XElement> sourceBlocks,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxDocumentSettings? documentSettings,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        IReadOnlyDictionary<XElement, int>? floatingBlocks)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Dictionary<XElement, int>? paragraphIndexes = null;
        Dictionary<XElement, int>? blockIndexes = null;
        var drawings = new List<DocxFloatingDrawing>();
        foreach (XElement anchor in source.Descendants(WordprocessingDrawingNamespace + "anchor"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsInsideExcludedRevisionContainer(anchor, markupMode) ||
                anchor.Ancestors(WordprocessingNamespace + "tr")
                    .Any(row => !IsIncludedTableRow(row, markupMode, cancellationToken)))
            {
                continue;
            }

            // Ordinals belong to the authored source, including hidden siblings.
            paragraphIndexes ??= IndexSources(source.Descendants(WordprocessingNamespace + "p"));
            blockIndexes ??= IndexSources(sourceBlocks);
            drawings.Add(ReadFloatingDrawing(
                anchor,
                package,
                relationships,
                styles,
                numbering,
                documentSettings,
                FindSourceAncestorIndex(anchor, paragraphIndexes),
                FindSourceAncestorIndex(anchor, blockIndexes),
                floatingBlocks?.TryGetValue(anchor, out int blockIndex) == true ? blockIndex : null,
                FindInheritedRevision(anchor, markupMode),
                markupMode,
                cancellationToken));
        }

        return drawings;

        Dictionary<XElement, int> IndexSources(IEnumerable<XElement> elements)
        {
            var indexes = new Dictionary<XElement, int>(ReferenceEqualityComparer.Instance);
            foreach (XElement element in elements)
            {
                cancellationToken.ThrowIfCancellationRequested();
                indexes.Add(element, indexes.Count);
            }
            return indexes;
        }
    }

    private static Dictionary<XElement, int>? CreateFloatingBlockMap(XContainer source) =>
        source.Descendants(WordprocessingDrawingNamespace + "anchor").Any()
            ? new(ReferenceEqualityComparer.Instance)
            : null;

    private static DocxFloatingDrawing ReadFloatingDrawing(
        XElement anchor,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxDocumentSettings? documentSettings,
        int? sourceParagraphIndex,
        int? sourceBlockIndex,
        int? layoutBlockIndex,
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
        string? relationshipId = ReadDrawingImageRelationshipId(anchor, cancellationToken);
        DocxInlineImage? image = ReadDrawingImage(anchor, package, relationships, revision, cancellationToken);
        XElement? textBoxBodyProperties = EnumerateDrawingOwnedElements(anchor, WordprocessingShapeNamespace + "bodyPr", cancellationToken)
            .FirstOrDefault();
        // RV05 floatbox probe (Word COM reference edge-floatbox): stroked textbox
        // content starts inside the border, so the outline width joins the insets.
        // Nested drawings own their own properties, as for inline text boxes.
        XElement? textBoxOutline = EnumerateDrawingOwnedElements(anchor, DrawingNamespace + "ln", cancellationToken)
            .FirstOrDefault(element => element.Attribute("w") is not null);
        string? textBoxStrokeWidthEmuValue = (string?)textBoxOutline?.Attribute("w");
        string? textBoxStrokeColorHex = ResolveTextBoxOutlineColorHex(textBoxOutline);
        IReadOnlyList<DocxBodyElement> textBoxBodyElements = ReadTextBoxBodyElements(
            anchor,
            styles,
            numbering,
            documentSettings,
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
            (string?)textBoxBodyProperties?.Attribute("bIns"),
            TextBoxStrokeWidthEmuValue: textBoxStrokeWidthEmuValue,
            TextBoxStrokeColorHex: textBoxStrokeColorHex)
        {
            LayoutBlockIndex = layoutBlockIndex,
            StaticOwner = anchor.Annotation<DocxStaticDrawingOwner>(),
            Revisions = RevisionList(revision),
            TextBoxBodyElements = textBoxBodyElements
        };
    }

    // Resolved outline color hex (RRGGBB) for frame stroking.
    // Solid srgb colors resolve verbatim; presets resolve through the shared table.
    // Scheme/system colors and transforms stay out (null keeps legacy borderless
    // rendering) until separately probed.
    private static string? ResolveTextBoxOutlineColorHex(XElement? outline)
    {
        XElement? solidFill = outline?.Element(DrawingNamespace + "solidFill");
        if (solidFill is null)
        {
            return null;
        }

        string? srgb = (string?)solidFill.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        if (RgbColor.TryParse(srgb, out RgbColor parsed))
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:X2}{1:X2}{2:X2}", parsed.Red, parsed.Green, parsed.Blue);
        }

        string? preset = (string?)solidFill.Element(DrawingNamespace + "prstClr")?.Attribute("val");
        if (PptxPresetColors.TryResolve(preset, out RgbColor presetColor))
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:X2}{1:X2}{2:X2}", presetColor.Red, presetColor.Green, presetColor.Blue);
        }

        return null;
    }

    private static IReadOnlyList<DocxBodyElement> ReadTextBoxBodyElements(
        XElement anchor,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        DocxDocumentSettings? documentSettings,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken)
    {
        XElement? textBoxContent = EnumerateDrawingOwnedElements(anchor, WordprocessingNamespace + "txbxContent", cancellationToken)
            .FirstOrDefault();
        if (textBoxContent is null)
        {
            return [];
        }

        return ReadRelatedStoryBodyElements(
            textBoxContent.Elements(),
            styles,
            numbering,
            new DocxNumberingCounters(),
            package,
            relationships,
            documentSettings,
            markupMode,
            cancellationToken);
    }

    private static bool IsBodyBlockElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "p" ||
            element.Name == WordprocessingNamespace + "tbl" ||
            element.Name == WordprocessingNamespace + "sectPr";
    }


    private static int? FindSourceAncestorIndex(XElement element, Dictionary<XElement, int> indexes)
    {
        foreach (XElement ancestor in element.Ancestors())
        {
            if (indexes.TryGetValue(ancestor, out int index))
            {
                return index;
            }
        }

        return null;
    }

    private static DocxPageSettings ReadPageSettings(
        XElement? sectionProperties,
        XDocument? settings,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPdfDocxMarkupMode markupMode,
        CancellationToken cancellationToken,
        RelatedStoryPartCache storyCache)
    {
        cancellationToken.ThrowIfCancellationRequested();
        XElement? pageSize = sectionProperties?.Element(WordprocessingNamespace + "pgSz");
        XElement? pageMargins = sectionProperties?.Element(WordprocessingNamespace + "pgMar");
        XElement? titlePage = sectionProperties?.Element(WordprocessingNamespace + "titlePg");
        // Office A/B (w53 even-pages probe, Word-COM rendered): Word honors the
        // section-level flag; settings.xml stays as a fallback for packages that carry
        // it there (mirrors the titlePg section lookup above).
        XElement? evenAndOddHeaders = sectionProperties?.Element(WordprocessingNamespace + "evenAndOddHeaders")
            ?? settings?.Root?.Element(WordprocessingNamespace + "evenAndOddHeaders");
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
        if (sectionProperties is null)
        {
            return pageSettings;
        }

        var (headerBodyElementsByType, headerDrawingsByType) = ReadReferencedStaticStories(
            sectionProperties, package, relationships, styles, numbering, HeaderRelationshipType, "headerReference",
            markupMode, cancellationToken, storyCache);
        var (footerBodyElementsByType, footerDrawingsByType) = ReadReferencedStaticStories(
            sectionProperties, package, relationships, styles, numbering, FooterRelationshipType, "footerReference",
            markupMode, cancellationToken, storyCache);

        return pageSettings with
        {
            HeaderBodyElementsByType = headerBodyElementsByType,
            FooterBodyElementsByType = footerBodyElementsByType,
            HeaderParagraphsByType = ToStaticParagraphsByType(headerBodyElementsByType),
            FooterParagraphsByType = ToStaticParagraphsByType(footerBodyElementsByType),
            HeaderFloatingDrawingsByType = headerDrawingsByType,
            FooterFloatingDrawingsByType = footerDrawingsByType
        };
    }

    private static XDocument? LoadRelatedXmlPart(OoxPackage package, string documentPartName, string relationshipType, out string? relatedPartName, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink = null, HashSet<string>? warnedParts = null, IMemoryReservation? modelStorage = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart? part = FindRelatedPart(package, documentPartName, relationshipType, cancellationToken);
        relatedPartName = part?.Name;
        if (part is null)
        {
            return null;
        }

        XDocument related = package.LoadXml(part, cancellationToken);
        OoxMemory.GrowModel(modelStorage, related, cancellationToken);
        OoxMarkupCompatibility.WarnMustUnderstandOnce(related, part.Name, diagnosticSink, warnedParts);
        return related;
    }

    private static string ResolveRelatedPartNameOrDefault(OoxPackage package, string documentPartName, string relationshipType, CancellationToken cancellationToken)
    {
        return FindRelatedPart(package, documentPartName, relationshipType, cancellationToken)?.Name ?? documentPartName;
    }

    private static OoxPart? FindRelatedPart(OoxPackage package, string documentPartName, string relationshipType, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxRelationship? relationship = package.GetRelationships(documentPartName, cancellationToken)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        // Optional parts belong to this document through relationships. A
        // matching but unreferenced package payload does not configure the story.
        return relationship?.ResolvedTarget is null ? null : package.GetPart(relationship.ResolvedTarget);
    }

    private static DocxFontCatalog LoadFontCatalog(OoxPackage package, string documentPartName, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink = null, HashSet<string>? warnedParts = null, IMemoryReservation? modelStorage = null)
    {
        XDocument? fontTable = LoadRelatedXmlPart(package, documentPartName, FontTableRelationshipType, out _, cancellationToken, diagnosticSink, warnedParts, modelStorage);
        XDocument? theme = LoadRelatedXmlPart(package, documentPartName, ThemeRelationshipType, out _, cancellationToken, diagnosticSink, warnedParts, modelStorage);
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
}
