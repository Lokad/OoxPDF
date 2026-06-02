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

    private static readonly XNamespace WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace WordprocessingDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

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
    private const string FootnotesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
    private const string EndnotesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
    private const string StylesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    private const string NumberingContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
    private const string SettingsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
    private const string FontTableContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";
    private const string ThemeContentType = "application/vnd.openxmlformats-officedocument.theme+xml";
    private const string CommentsContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
    private const string FootnotesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
    private const string EndnotesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml";
    private const double WordAutomaticParagraphSpacingPoints = 14d;
    public DocxDocument Read(OoxPackage package, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        OoxPart documentPart = FindDocumentPart(package);
        using Stream stream = documentPart.OpenRead();
        XDocument document = SafeXml.Load(stream);
        EmitUnsupportedFeatureDiagnostics(package, document, documentPart.Name, diagnosticSink);

        XElement? sectionProperties = document.Descendants(WordprocessingNamespace + "sectPr").LastOrDefault();
        XElement? pageSize = sectionProperties?.Element(WordprocessingNamespace + "pgSz");
        XElement? pageMargins = sectionProperties?.Element(WordprocessingNamespace + "pgMar");

        DocxFontCatalog fontCatalog = LoadFontCatalog(package, documentPart.Name);
        DocxStyleSet styles = LoadStyles(package, documentPart.Name);
        DocxNumberingSet numbering = LoadNumbering(package, documentPart.Name, fontCatalog);
        XDocument? settings = LoadRelatedXmlPart(package, documentPart.Name, SettingsRelationshipType, SettingsContentType, out _);
        DocxDocumentSettings documentSettings = ReadDocumentSettings(settings);
        IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(documentPart.Name)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        IReadOnlyDictionary<string, OoxRelationship> internalRelationships = relationships.Values
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        IReadOnlyList<DocxBodyElement> bodyElements = ReadBodyElements(document, styles, numbering, package, relationships, settings);
        IReadOnlyList<DocxParagraph> paragraphs = bodyElements.OfType<DocxParagraphElement>().Select(e => e.Paragraph).ToArray();
        IReadOnlyList<DocxTable> tables = bodyElements.OfType<DocxTableElement>().Select(e => e.Table).ToArray();
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> headersByType = ReadReferencedHeaderFooterParagraphsByType(document, package, internalRelationships, styles, numbering, HeaderRelationshipType, "headerReference");
        IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> footersByType = ReadReferencedHeaderFooterParagraphsByType(document, package, internalRelationships, styles, numbering, FooterRelationshipType, "footerReference");
        IReadOnlyList<DocxParagraph> headers = SelectDefaultHeaderFooterParagraphs(headersByType);
        IReadOnlyList<DocxParagraph> footers = SelectDefaultHeaderFooterParagraphs(footersByType);
        IReadOnlyList<DocxRelatedStory> relatedStories = ReadRelatedStories(package, documentPart.Name, styles, numbering);
        IReadOnlyList<DocxFloatingDrawing> floatingDrawings = ReadFloatingDrawings(document, package, relationships);

        if (pageSize is null)
        {
            return new DocxDocument(
                612d,
                792d,
                72d,
                72d,
                72d,
                72d,
                ReadPageSettings(pageSize, pageMargins, sectionProperties, settings, package, internalRelationships, styles, numbering),
                floatingDrawings,
                headers,
                footers,
                bodyElements,
                paragraphs,
                tables)
            {
                FontCatalog = fontCatalog,
                HeaderParagraphsByType = headersByType,
                FooterParagraphsByType = footersByType,
                RelatedStories = relatedStories,
                Settings = documentSettings
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
            ReadPageSettings(pageSize, pageMargins, sectionProperties, settings, package, internalRelationships, styles, numbering),
            floatingDrawings,
            headers,
            footers,
            bodyElements,
            paragraphs,
            tables)
        {
            FontCatalog = fontCatalog,
            HeaderParagraphsByType = headersByType,
            FooterParagraphsByType = footersByType,
            RelatedStories = relatedStories,
            Settings = documentSettings
        };
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
        XElement? useFELayout = root.Element(WordprocessingNamespace + "compat")?.Element(WordprocessingNamespace + "useFELayout");
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
            compatSettings);
    }

    private static IReadOnlyList<DocxFloatingDrawing> ReadFloatingDrawings(
        XDocument document,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement[] paragraphs = document
            .Descendants(WordprocessingNamespace + "p")
            .ToArray();
        XElement[] bodyBlocks = document
            .Root?
            .Element(WordprocessingNamespace + "body")?
            .Elements()
            .Where(IsBodyBlockElement)
            .ToArray() ?? [];
        return document
            .Descendants(WordprocessingDrawingNamespace + "anchor")
            .Select(anchor => ReadFloatingDrawing(
                anchor,
                package,
                relationships,
                FindSourceParagraphIndex(anchor, paragraphs),
                FindSourceBlockIndex(anchor, bodyBlocks)))
            .ToArray();
    }

    private static DocxFloatingDrawing ReadFloatingDrawing(
        XElement anchor,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        int? sourceParagraphIndex,
        int? sourceBlockIndex)
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
        DocxInlineImage? image = ReadDrawingImage(anchor, package, relationships);

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
            sourceBlockIndex);
    }

    private static bool IsBodyBlockElement(XElement element)
    {
        return element.Name == WordprocessingNamespace + "p" ||
            element.Name == WordprocessingNamespace + "tbl" ||
            element.Name == WordprocessingNamespace + "sectPr";
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
            .FirstOrDefault(ancestor => ancestor.Parent?.Name == WordprocessingNamespace + "body");
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
        DocxNumberingSet? numbering)
    {
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
            DocGridLinePitchValue = (string?)docGrid?.Attribute(WordprocessingNamespace + "linePitch")
        };
        if (sectionProperties is null || package is null || relationships is null || styles is null || numbering is null)
        {
            return pageSettings;
        }

        return pageSettings with
        {
            HeaderParagraphsByType = ReadReferencedHeaderFooterParagraphsByType(
                sectionProperties,
                package,
                relationships,
                styles,
                numbering,
                HeaderRelationshipType,
                "headerReference"),
            FooterParagraphsByType = ReadReferencedHeaderFooterParagraphsByType(
                sectionProperties,
                package,
                relationships,
                styles,
                numbering,
                FooterRelationshipType,
                "footerReference")
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

    private static void EmitUnsupportedFeatureDiagnostics(OoxPackage package, XDocument document, string partName, Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        var emitted = new HashSet<string>(StringComparer.Ordinal);
        void Emit(string id, string feature, string diagnosticPartName = "", string fallback = "Ignored")
        {
            if (!emitted.Add(id))
            {
                return;
            }

            diagnosticSink(new OoxPdfDiagnostic(
                id,
                OoxPdfSeverity.Warning,
                $"Unsupported DOCX feature '{feature}' was detected and ignored or approximated.",
                diagnosticPartName.Length == 0 ? partName : diagnosticPartName,
                Feature: feature,
                Fallback: fallback));
        }

        if (document.Descendants(WordprocessingNamespace + "commentRangeStart").Any() ||
            document.Descendants(WordprocessingNamespace + "commentReference").Any())
        {
            Emit(
                "DOCX_UNSUPPORTED_COMMENTS",
                "comments",
                ResolveRelatedPartNameOrDefault(package, partName, CommentsRelationshipType, CommentsContentType));
        }

        if (HasUnsupportedTrackedChanges(document))
        {
            Emit("DOCX_UNSUPPORTED_TRACKED_CHANGES", "tracked changes");
        }

        if (HasUnsupportedComplexFields(document))
        {
            Emit("DOCX_UNSUPPORTED_COMPLEX_FIELD", "complex field");
        }

        if (document.Descendants(MathNamespace + "oMath").Any() ||
            document.Descendants(MathNamespace + "oMathPara").Any())
        {
            Emit("DOCX_UNSUPPORTED_EQUATION", "equation");
        }

        if (document.Descendants(WordprocessingNamespace + "object").Any())
        {
            Emit("DOCX_UNSUPPORTED_OLE_OBJECT", "OLE object");
        }

        if (document.Descendants(WordprocessingDrawingNamespace + "anchor").Any())
        {
            Emit("DOCX_UNSUPPORTED_FLOATING_DRAWING", "floating drawing");
        }

        if (document.Descendants(WordprocessingNamespace + "footnoteReference").Any())
        {
            Emit(
                "DOCX_UNSUPPORTED_FOOTNOTE",
                "footnote",
                ResolveRelatedPartNameOrDefault(package, partName, FootnotesRelationshipType, FootnotesContentType));
        }

        if (document.Descendants(WordprocessingNamespace + "endnoteReference").Any())
        {
            Emit(
                "DOCX_UNSUPPORTED_ENDNOTE",
                "endnote",
                ResolveRelatedPartNameOrDefault(package, partName, EndnotesRelationshipType, EndnotesContentType));
        }

        if (document.Descendants(WordprocessingNamespace + "cols").Any(cols =>
            cols.Attribute(WordprocessingNamespace + "num") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int columns) &&
            columns > 1))
        {
            Emit("DOCX_UNSUPPORTED_MULTI_COLUMN", "multi-column section");
        }

        if (document.Descendants(WordprocessingNamespace + "br").Any(IsUnsupportedColumnBreak))
        {
            Emit("DOCX_UNSUPPORTED_MANUAL_BREAK", "inline manual column break", fallback: "Break-only column paragraphs are supported");
        }

        if (document.Descendants().Any(IsUnsupportedParagraphKeepRule))
        {
            Emit("DOCX_UNSUPPORTED_PARAGRAPH_KEEP_RULE", "table-cell paragraph keep/widow-orphan rule", fallback: "Body paragraph keep rules are supported");
        }

        if (document.Descendants(WordprocessingNamespace + "pPr")
            .Elements(WordprocessingNamespace + "sectPr")
            .Any(IsUnsupportedParagraphSectionBreak))
        {
            Emit("DOCX_UNSUPPORTED_SECTION_BREAK", "continuous or unknown paragraph section break", fallback: "Partially supported");
        }

        XDocument? styles = LoadRelatedXmlPart(package, partName, StylesRelationshipType, StylesContentType, out string? stylesPartName);
        if (styles is not null)
        {
            if (styles.Descendants(WordprocessingNamespace + "style")
                .Elements(WordprocessingNamespace + "pPr")
                .Elements(WordprocessingNamespace + "spacing")
                .Any(spacing =>
                    spacing.Attribute(WordprocessingNamespace + "beforeAutospacing") is not null ||
                    spacing.Attribute(WordprocessingNamespace + "afterAutospacing") is not null))
            {
                Emit("DOCX_STYLE_PARAGRAPH_SPACING", "style paragraph spacing variant", stylesPartName ?? partName, "Approximated");
            }

            if (HasUnsupportedTableBorderStyle(styles))
            {
                Emit("DOCX_TABLE_BORDER_STYLE", "table border style", stylesPartName ?? partName, "Approximated");
            }
        }

        if (HasUnsupportedTableBorderStyle(document))
        {
            Emit("DOCX_TABLE_BORDER_STYLE", "table border style", partName, "Approximated");
        }

        XDocument? numbering = LoadRelatedXmlPart(package, partName, NumberingRelationshipType, NumberingContentType, out string? numberingPartName);
        if (numbering is not null &&
            numbering.Descendants(WordprocessingNamespace + "lvl")
                .Elements(WordprocessingNamespace + "pPr")
                .Elements(WordprocessingNamespace + "ind")
                .Any(ind =>
                    ind.Attribute(WordprocessingNamespace + "firstLine") is not null ||
                    HasCharacterUnitIndent(ind)))
        {
            Emit("DOCX_NUMBERING_INDENT", "numbering level indent", numberingPartName ?? partName, "Approximated");
        }

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

    private static bool IsUnsupportedColumnBreak(XElement breakElement)
    {
        if (!IsColumnBreak(breakElement))
        {
            return false;
        }

        XElement? paragraph = breakElement.Ancestors(WordprocessingNamespace + "p").FirstOrDefault();
        return paragraph is null || !IsRunColumnBreakOnlyParagraph(paragraph);
    }

    private static bool IsUnsupportedParagraphKeepRule(XElement element)
    {
        if (element.Name != WordprocessingNamespace + "keepNext" &&
            element.Name != WordprocessingNamespace + "keepLines" &&
            element.Name != WordprocessingNamespace + "widowControl")
        {
            return false;
        }

        XElement? paragraph = element.Ancestors(WordprocessingNamespace + "p").FirstOrDefault();
        return paragraph is null || paragraph.Ancestors(WordprocessingNamespace + "tc").Any();
    }

    private static bool HasUnsupportedTableBorderStyle(XDocument document)
    {
        return document
            .Descendants()
            .Any(element => element.Name.Namespace == WordprocessingNamespace &&
                IsTableBorderContainer(element.Parent) &&
                IsUnsupportedVisibleBorderStyle((string?)element.Attribute(WordprocessingNamespace + "val")));
    }

    private static bool HasUnsupportedTrackedChanges(XDocument document)
    {
        XName[] unsupportedTrackChangeContainers =
        [
            WordprocessingNamespace + "moveFrom",
            WordprocessingNamespace + "moveTo",
            WordprocessingNamespace + "moveFromRangeStart",
            WordprocessingNamespace + "moveFromRangeEnd",
            WordprocessingNamespace + "moveToRangeStart",
            WordprocessingNamespace + "moveToRangeEnd"
        ];
        if (unsupportedTrackChangeContainers.Any(name => document.Descendants(name).Any()))
        {
            return true;
        }

        return document.Descendants(WordprocessingNamespace + "ins").Any(insertion =>
            insertion.Parent?.Name != WordprocessingNamespace + "p" ||
            insertion.Elements().Any(child => child.Name != WordprocessingNamespace + "r"));
    }

    private static bool HasUnsupportedComplexFields(XDocument document)
    {
        foreach (XElement paragraph in document.Descendants(WordprocessingNamespace + "p"))
        {
            if (HasUnsupportedComplexFieldInInlineContainer(paragraph))
            {
                return true;
            }

            foreach (XElement container in paragraph.Descendants().Where(IsVisibleInlineContainer))
            {
                if (HasUnsupportedComplexFieldInInlineContainer(container))
                {
                    return true;
                }
            }
        }

        return document.Descendants(WordprocessingNamespace + "fldChar").Any(fieldChar => !IsSupportedInlineRunChild(fieldChar)) ||
            document.Descendants(WordprocessingNamespace + "instrText").Any(instruction => !IsSupportedInlineRunChild(instruction));
    }

    private static bool IsVisibleInlineContainer(XElement? element)
    {
        return element is not null &&
            (element.Name == WordprocessingNamespace + "hyperlink" ||
            element.Name == WordprocessingNamespace + "fldSimple" ||
            element.Name == WordprocessingNamespace + "ins");
    }

    private static bool IsSupportedInlineRunChild(XElement element)
    {
        XElement? run = element.Parent;
        XElement? container = run?.Parent;
        return run?.Name == WordprocessingNamespace + "r" &&
            (container?.Name == WordprocessingNamespace + "p" || IsVisibleInlineContainer(container));
    }

    private static bool HasUnsupportedComplexFieldInInlineContainer(XElement container)
    {
        bool inField = false;
        bool inResult = false;
        bool hasSeparate = false;
        bool hasCachedResult = false;
        var instruction = new StringBuilder();

        foreach (XElement run in container.Elements(WordprocessingNamespace + "r"))
        {
            foreach (XElement child in run.Elements())
            {
                if (child.Name == WordprocessingNamespace + "fldChar")
                {
                    string? fieldCharType = (string?)child.Attribute(WordprocessingNamespace + "fldCharType");
                    if (string.Equals(fieldCharType, "begin", StringComparison.OrdinalIgnoreCase))
                    {
                        if (inField)
                        {
                            return true;
                        }

                        inField = true;
                        inResult = false;
                        hasSeparate = false;
                        hasCachedResult = false;
                        instruction.Clear();
                        continue;
                    }

                    if (string.Equals(fieldCharType, "separate", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!inField)
                        {
                            return true;
                        }

                        hasSeparate = true;
                        inResult = true;
                        continue;
                    }

                    if (string.Equals(fieldCharType, "end", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!inField)
                        {
                            return true;
                        }

                        if (ResolveFieldPlaceholder(instruction.ToString()) is null && (!hasSeparate || !hasCachedResult))
                        {
                            return true;
                        }

                        inField = false;
                        inResult = false;
                        continue;
                    }

                    return true;
                }

                if (child.Name == WordprocessingNamespace + "instrText")
                {
                    if (!inField || inResult)
                    {
                        return true;
                    }

                    instruction.Append((string?)child);
                    continue;
                }

                if (inResult && child.Name == WordprocessingNamespace + "t")
                {
                    hasCachedResult = true;
                }
            }
        }

        return inField;
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
            value.Equals("single", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("nil", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    private static XDocument? LoadRelatedXmlPart(OoxPackage package, string documentPartName, string relationshipType, string contentType, out string? relatedPartName)
    {
        OoxPart? part = FindRelatedPart(package, documentPartName, relationshipType, contentType);
        relatedPartName = part?.Name;
        if (part is null)
        {
            return null;
        }

        using Stream stream = part.OpenRead();
        return SafeXml.Load(stream);
    }

    private static string ResolveRelatedPartNameOrDefault(OoxPackage package, string documentPartName, string relationshipType, string contentType)
    {
        return FindRelatedPart(package, documentPartName, relationshipType, contentType)?.Name ?? documentPartName;
    }

    private static OoxPart? FindRelatedPart(OoxPackage package, string documentPartName, string relationshipType, string contentType)
    {
        OoxRelationship? relationship = package.GetRelationships(documentPartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null
            ? package.Parts.FirstOrDefault(p => p.ContentType == contentType)
            : package.GetPart(relationship.ResolvedTarget);
    }

    private static DocxFontCatalog LoadFontCatalog(OoxPackage package, string documentPartName)
    {
        XDocument? fontTable = LoadRelatedXmlPart(package, documentPartName, FontTableRelationshipType, FontTableContentType, out _);
        XDocument? theme = LoadRelatedXmlPart(package, documentPartName, ThemeRelationshipType, ThemeContentType, out _);
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
            (string?)fontScheme
                ?.Element(DrawingNamespace + "majorFont")
                ?.Element(DrawingNamespace + "latin")
                ?.Attribute("typeface"),
            (string?)fontScheme
                ?.Element(DrawingNamespace + "minorFont")
                ?.Element(DrawingNamespace + "latin")
                ?.Attribute("typeface"),
            (string?)fontScheme
                ?.Element(DrawingNamespace + "majorFont")
                ?.Element(DrawingNamespace + "cs")
                ?.Attribute("typeface"),
            (string?)fontScheme
                ?.Element(DrawingNamespace + "minorFont")
                ?.Element(DrawingNamespace + "cs")
                ?.Attribute("typeface"));
    }

    private static IReadOnlyList<DocxParagraph> ReadParagraphs(
        XDocument document,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var paragraphs = new List<DocxParagraph>();
        var numberingCounters = new Dictionary<(string NumId, int Level), int>();
        foreach (XElement paragraph in document.Descendants(WordprocessingNamespace + "body").Elements(WordprocessingNamespace + "p"))
        {
            DocxParagraph? parsed = ReadParagraph(paragraph, styles, numbering, numberingCounters, package, relationships);
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
        DocxTableCellStyle? tableCellStyle = null)
    {
        XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
        string? paragraphStyleId = (string?)paragraphProperties?
            .Element(WordprocessingNamespace + "pStyle")
            ?.Attribute(WordprocessingNamespace + "val");
        DocxResolvedParagraphProperties resolvedParagraph = ResolveParagraphProperties(
            paragraphProperties,
            paragraphStyleId,
            styles,
            tableCellStyle?.Paragraph);
        var runs = new List<DocxTextRun>();
        var images = new List<DocxInlineImage>();
        var inlineReferences = new List<DocxInlineReference>();
        var fieldReferences = new List<DocxFieldReference>();
        var hyperlinkSpans = new List<DocxHyperlinkSpan>();
        var bookmarkAnchors = new List<DocxBookmarkAnchor>();
        bool pageInstructionSeen = false;
        int sourceRunIndex = 0;
        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == WordprocessingNamespace + "r")
            {
                AddParagraphRun(child, ref pageInstructionSeen);
            }
            else if (child.Name == WordprocessingNamespace + "fldSimple")
            {
                AddSimpleField(child);
            }
            else if (child.Name == WordprocessingNamespace + "ins")
            {
                foreach (XElement insertedChild in child.Elements())
                {
                    AddInlineContainerChild(insertedChild);
                }
            }
            else if (child.Name == WordprocessingNamespace + "hyperlink")
            {
                int sourceRunStartIndex = sourceRunIndex;
                int textRunStartIndex = runs.Count;
                int textLengthStart = runs.Sum(run => run.Text.Length);
                foreach (XElement hyperlinkChild in child.Elements())
                {
                    AddInlineContainerChild(hyperlinkChild);
                }

                AddHyperlinkSpan(child, sourceRunStartIndex, sourceRunIndex - sourceRunStartIndex, textRunStartIndex, runs.Count - textRunStartIndex, runs.Sum(run => run.Text.Length) - textLengthStart);
            }
            else if (child.Name == WordprocessingNamespace + "bookmarkStart")
            {
                AddBookmarkAnchor(child);
            }
        }

        void AddInlineContainerChild(XElement child)
        {
            if (child.Name == WordprocessingNamespace + "r")
            {
                AddParagraphRun(child, ref pageInstructionSeen);
            }
            else if (child.Name == WordprocessingNamespace + "bookmarkStart")
            {
                AddBookmarkAnchor(child);
            }
            else if (child.Name == WordprocessingNamespace + "fldSimple")
            {
                AddSimpleField(child);
            }
        }

        if (runs.Count == 0 && images.Count == 0)
        {
            DocxResolvedRunProperties paragraphMarkRun = ResolveRunProperties(
                paragraphProperties?.Element(WordprocessingNamespace + "rPr"),
                paragraphStyleId,
                characterStyleId: null,
                styles,
                tableCellStyle?.Run);
            AddResolvedTextRun(runs, string.Empty, paragraphMarkRun, complexScript: false);
        }

        double paragraphFontSize = runs.Count == 0 ? 11d : runs.Max(run => run.FontSize);
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
            InlineReferences = inlineReferences,
            FieldReferences = fieldReferences,
            Hyperlinks = hyperlinkSpans,
            BookmarkAnchors = bookmarkAnchors
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

        void AddSimpleField(XElement field)
        {
            string? instruction = (string?)field.Attribute(WordprocessingNamespace + "instr");
            string kind = ResolveFieldKind(instruction);
            string? placeholder = ResolveFieldPlaceholder(instruction);
            int fieldSourceRunIndex = sourceRunIndex;
            int fieldTextRunIndex = runs.Count;
            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);
            if (placeholder is not null)
            {
                XElement? firstRun = field.Elements(WordprocessingNamespace + "r").FirstOrDefault();
                if (firstRun is null)
                {
                    runs.Add(new DocxTextRun(placeholder, 11d, null, false, false, false, null, null));
                    AddFieldReference(kind, "Simple", instruction, placeholder, fieldSourceRunIndex, fieldTextRunIndex, fieldTextLengthStart);
                    return;
                }

                AddFieldPlaceholderRun(firstRun, placeholder);
                images.AddRange(ReadInlineImages(firstRun, package, relationships));
                AddFieldReference(kind, "Simple", instruction, placeholder, fieldSourceRunIndex, fieldTextRunIndex, fieldTextLengthStart);
                return;
            }

            foreach (XElement fieldRun in field.Elements(WordprocessingNamespace + "r"))
            {
                bool fieldPageInstructionSeen = false;
                AddParagraphRun(fieldRun, ref fieldPageInstructionSeen);
            }

            AddFieldReference(kind, "Simple", instruction, placeholder, fieldSourceRunIndex, fieldTextRunIndex, fieldTextLengthStart);
        }

        void AddFieldReference(
            string kind,
            string sourceKind,
            string? instruction,
            string? placeholder,
            int fieldSourceRunIndex,
            int fieldTextRunIndex,
            int fieldTextLengthStart)
        {
            fieldReferences.Add(new DocxFieldReference(
                kind,
                sourceKind,
                instruction,
                placeholder,
                fieldSourceRunIndex,
                fieldTextRunIndex,
                runs.Count - fieldTextRunIndex,
                runs.Sum(run => run.Text.Length) - fieldTextLengthStart));
        }

        void AddFieldPlaceholderRun(XElement run, string text)
        {
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                run.Element(WordprocessingNamespace + "rPr"),
                paragraphStyleId,
                ReadCharacterStyleId(run),
                styles,
                tableCellStyle?.Run);
            AddResolvedTextRuns(runs, resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text, resolvedRun);
        }

        void AddParagraphRun(XElement run, ref bool currentPageInstructionSeen)
        {
            int currentSourceRunIndex = sourceRunIndex++;
            AddInlineReferences(run, currentSourceRunIndex);
            string text = ReadRunText(run);
            string? fieldInstruction = run
                .Elements(WordprocessingNamespace + "instrText")
                .Select(instruction => (string?)instruction)
                .FirstOrDefault(value => value is not null);
            string? placeholder = ResolveFieldPlaceholder(fieldInstruction);
            int fieldTextRunIndex = runs.Count;
            int fieldTextLengthStart = runs.Sum(run => run.Text.Length);

            if (placeholder is not null)
            {
                text = placeholder;
                currentPageInstructionSeen = true;
            }
            else if (currentPageInstructionSeen && text.Trim().All(char.IsDigit))
            {
                text = string.Empty;
            }

            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(
                runProperties,
                paragraphStyleId,
                ReadCharacterStyleId(run),
                styles,
                tableCellStyle?.Run);
            if (text.Length != 0)
            {
                string displayText = resolvedRun.AllCaps == true ? text.ToUpperInvariant() : text;
                AddResolvedTextRuns(runs, displayText, resolvedRun);
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

            images.AddRange(ReadInlineImages(run, package, relationships));
        }

        void AddInlineReferences(XElement run, int currentSourceRunIndex)
        {
            int childIndex = 0;
            int textOffset = 0;
            foreach (XElement child in run.Elements())
            {
                if (child.Name == WordprocessingNamespace + "commentReference")
                {
                    inlineReferences.Add(new DocxInlineReference(
                        "Comment",
                        (string?)child.Attribute(WordprocessingNamespace + "id"),
                        null,
                        currentSourceRunIndex,
                        childIndex,
                        textOffset));
                }
                else if (child.Name == WordprocessingNamespace + "footnoteReference")
                {
                    inlineReferences.Add(new DocxInlineReference(
                        "Footnote",
                        (string?)child.Attribute(WordprocessingNamespace + "id"),
                        (string?)child.Attribute(WordprocessingNamespace + "customMarkFollows"),
                        currentSourceRunIndex,
                        childIndex,
                        textOffset));
                }
                else if (child.Name == WordprocessingNamespace + "endnoteReference")
                {
                    inlineReferences.Add(new DocxInlineReference(
                        "Endnote",
                        (string?)child.Attribute(WordprocessingNamespace + "id"),
                        (string?)child.Attribute(WordprocessingNamespace + "customMarkFollows"),
                        currentSourceRunIndex,
                        childIndex,
                        textOffset));
                }

                textOffset += ReadRunTextChild(child).Length;
                childIndex++;
            }
        }
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

    private static void AddResolvedTextRuns(List<DocxTextRun> runs, string text, DocxResolvedRunProperties resolvedRun)
    {
        var segment = new StringBuilder();
        bool? currentComplexScript = null;
        foreach (Rune rune in text.EnumerateRunes())
        {
            bool complexScript = DocxScriptClassifier.IsComplexScriptRune(rune.Value);
            if (currentComplexScript is not null && currentComplexScript.Value != complexScript)
            {
                AddResolvedTextRun(runs, segment.ToString(), resolvedRun, currentComplexScript.Value);
                segment.Clear();
            }

            segment.Append(rune);
            currentComplexScript = complexScript;
        }

        if (segment.Length != 0 && currentComplexScript is not null)
        {
            AddResolvedTextRun(runs, segment.ToString(), resolvedRun, currentComplexScript.Value);
        }
    }

    private static void AddResolvedTextRun(List<DocxTextRun> runs, string text, DocxResolvedRunProperties resolvedRun, bool complexScript)
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
            resolvedRun.FontSize ?? 11d,
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
            resolvedRun.HiddenValue)
        {
            Fonts = resolvedRun.Fonts
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

    private static string ReadRunTextChild(XElement child)
    {
        if (child.Name == WordprocessingNamespace + "t")
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

    private static IReadOnlyList<DocxBodyElement> ReadBodyElements(
        XDocument document,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        XDocument? settings)
    {
        var elements = new List<DocxBodyElement>();
        var numberingCounters = new Dictionary<(string NumId, int Level), int>();
        foreach (XElement element in document.Descendants(WordprocessingNamespace + "body").Elements())
        {
            if (element.Name == WordprocessingNamespace + "p")
            {
                XElement? paragraphProperties = element.Element(WordprocessingNamespace + "pPr");
                XElement? pageBreakBefore = paragraphProperties?.Element(WordprocessingNamespace + "pageBreakBefore");
                if (ReadOnOff(pageBreakBefore) == true)
                {
                    elements.Add(new DocxPageBreakElement("pageBreakBefore", (string?)pageBreakBefore?.Attribute(WordprocessingNamespace + "val")));
                }

                if (IsRunPageBreakOnlyParagraph(element))
                {
                    elements.Add(new DocxPageBreakElement(
                        "runBreak",
                        "page",
                        ReadParagraph(element, styles, numbering, numberingCounters, package, relationships)));
                    XElement? breakParagraphSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (breakParagraphSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(breakParagraphSectionProperties, package, relationships, styles, numbering, settings));
                    }

                    continue;
                }

                if (IsRunColumnBreakOnlyParagraph(element))
                {
                    elements.Add(new DocxManualBreakElement(
                        "runBreak",
                        "column",
                        ReadParagraph(element, styles, numbering, numberingCounters, package, relationships)));
                    XElement? breakParagraphSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (breakParagraphSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(breakParagraphSectionProperties, package, relationships, styles, numbering, settings));
                    }

                    continue;
                }

                if (HasRunPageBreak(element))
                {
                    foreach (ParagraphPageBreakPart part in SplitParagraphAtRunPageBreaks(element))
                    {
                        if (part.BreakValue is not null)
                        {
                            elements.Add(new DocxPageBreakElement("runBreak", part.BreakValue));
                            continue;
                        }

                        if (part.Paragraph is null)
                        {
                            continue;
                        }

                        DocxParagraph? splitParagraph = ReadParagraph(part.Paragraph, styles, numbering, numberingCounters, package, relationships);
                        if (splitParagraph is not null)
                        {
                            elements.Add(new DocxParagraphElement(AdjustPageBreakParagraphFragment(splitParagraph, part)));
                        }
                    }

                    XElement? splitSectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                    if (splitSectionProperties is not null)
                    {
                        elements.Add(ReadSectionBreak(splitSectionProperties, package, relationships, styles, numbering, settings));
                    }

                    continue;
                }

                DocxParagraph? paragraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships);
                if (paragraph is not null)
                {
                    elements.Add(new DocxParagraphElement(paragraph));
                }

                XElement? sectionProperties = paragraphProperties?.Element(WordprocessingNamespace + "sectPr");
                if (sectionProperties is not null)
                {
                    elements.Add(ReadSectionBreak(sectionProperties, package, relationships, styles, numbering, settings));
                }
            }
            else if (element.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? table = ReadTable(element, styles, numbering, numberingCounters, package, relationships);
                if (table is not null)
                {
                    elements.Add(new DocxTableElement(table));
                }
            }
        }

        return elements;
    }

    private static bool HasRunPageBreak(XElement paragraph)
    {
        return paragraph
            .Elements()
            .Any(HasVisibleRunPageBreak);
    }

    private static DocxParagraph AdjustPageBreakParagraphFragment(DocxParagraph paragraph, ParagraphPageBreakPart part)
    {
        return paragraph with
        {
            SpacingBeforePoints = part.StartsAfterBreak ? 0d : paragraph.SpacingBeforePoints,
            SpacingAfterPoints = part.EndsBeforeBreak ? 0d : paragraph.SpacingAfterPoints,
            ListLabel = part.StartsAfterBreak ? null : paragraph.ListLabel
        };
    }

    private static IReadOnlyList<ParagraphPageBreakPart> SplitParagraphAtRunPageBreaks(XElement paragraph)
    {
        var parts = new List<ParagraphPageBreakPart>();
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
            parts.Add(new ParagraphPageBreakPart(splitParagraph, null, startsAfterBreak, endsBeforeBreak));
            currentChildren.Clear();
            startsAfterBreak = false;
        }

        foreach (XElement child in paragraph.Elements())
        {
            if (child.Name == WordprocessingNamespace + "pPr")
            {
                continue;
            }

            if (TrySplitRunPageBreakContainer(child, currentChildren, AddParagraphPart, parts, ref startsAfterBreak))
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

                if (runChild.Name == WordprocessingNamespace + "br" && IsPageBreak(runChild))
                {
                    AddRunPart(currentChildren, runProperties, runChildren);
                    AddParagraphPart(endsBeforeBreak: true);
                    parts.Add(new ParagraphPageBreakPart(null, (string?)runChild.Attribute(WordprocessingNamespace + "type"), false, false));
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

    private static bool TrySplitRunPageBreakContainer(
        XElement child,
        List<XElement> paragraphChildren,
        Action<bool> addParagraphPart,
        List<ParagraphPageBreakPart> parts,
        ref bool startsAfterBreak)
    {
        if (!IsVisibleRunContainer(child) || !HasVisibleRunPageBreak(child))
        {
            return false;
        }

        var containerChildren = new List<XElement>();
        foreach (XElement containerChild in child.Elements())
        {
            if (containerChild.Name != WordprocessingNamespace + "r")
            {
                containerChildren.Add(new XElement(containerChild));
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

                if (runChild.Name == WordprocessingNamespace + "br" && IsPageBreak(runChild))
                {
                    AddRunPart(containerChildren, runProperties, runChildren);
                    AddContainerPart(paragraphChildren, child, containerChildren);
                    addParagraphPart(true);
                    parts.Add(new ParagraphPageBreakPart(null, (string?)runChild.Attribute(WordprocessingNamespace + "type"), false, false));
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

        AddContainerPart(paragraphChildren, child, containerChildren);
        return true;
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

    private static bool IsRunPageBreakOnlyParagraph(XElement paragraph)
    {
        return IsRunBreakOnlyParagraph(paragraph, IsPageBreak);
    }

    private static bool IsRunColumnBreakOnlyParagraph(XElement paragraph)
    {
        return IsRunBreakOnlyParagraph(paragraph, IsColumnBreak);
    }

    private static bool IsRunBreakOnlyParagraph(XElement paragraph, Func<XElement, bool> isBreak)
    {
        bool hasBreak = paragraph
            .Elements(WordprocessingNamespace + "r")
            .SelectMany(run => run.Elements(WordprocessingNamespace + "br"))
            .Any(isBreak);
        if (!hasBreak)
        {
            return false;
        }

        foreach (XElement run in paragraph.Elements(WordprocessingNamespace + "r"))
        {
            foreach (XElement child in run.Elements())
            {
                if (child.Name == WordprocessingNamespace + "rPr")
                {
                    continue;
                }

                if (child.Name == WordprocessingNamespace + "br" && isBreak(child))
                {
                    continue;
                }

                return false;
            }
        }

        return paragraph.Elements().All(element =>
            element.Name == WordprocessingNamespace + "pPr" ||
            element.Name == WordprocessingNamespace + "r");
    }

    private static bool HasVisibleRunPageBreak(XElement element)
    {
        if (element.Name == WordprocessingNamespace + "r")
        {
            return element.Elements(WordprocessingNamespace + "br").Any(IsPageBreak);
        }

        return IsVisibleRunContainer(element) &&
            element.Elements(WordprocessingNamespace + "r").Any(HasVisibleRunPageBreak);
    }

    private static bool IsVisibleRunContainer(XElement element)
    {
        return element.Name == WordprocessingNamespace + "fldSimple" ||
            element.Name == WordprocessingNamespace + "hyperlink" ||
            element.Name == WordprocessingNamespace + "ins";
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
        XDocument? settings)
    {
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
                numbering),
            (string?)sectionProperties
                .Element(WordprocessingNamespace + "type")
                ?.Attribute(WordprocessingNamespace + "val"),
            (string?)columns?.Attribute(WordprocessingNamespace + "num"),
            (string?)columns?.Attribute(WordprocessingNamespace + "equalWidth"),
            (string?)columns?.Attribute(WordprocessingNamespace + "space"),
            ReadSectionColumns(columns));
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

    private static IReadOnlyDictionary<string, IReadOnlyList<DocxParagraph>> ReadReferencedHeaderFooterParagraphsByType(
        XContainer referenceRoot,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        string relationshipType,
        string referenceElementName)
    {
        var paragraphs = new Dictionary<string, IReadOnlyList<DocxParagraph>>(StringComparer.OrdinalIgnoreCase);
        foreach (XElement reference in referenceRoot.Descendants(WordprocessingNamespace + referenceElementName))
        {
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
            XDocument partXml = SafeXml.Load(stream);
            string type = (string?)reference.Attribute(WordprocessingNamespace + "type") ?? "default";
            paragraphs[type] = ReadParagraphElements(
                partXml.Root?.Elements(WordprocessingNamespace + "p") ?? [],
                styles,
                numbering,
                package,
                package.GetRelationships(part.Name).Where(r => !r.IsExternal && r.ResolvedTarget is not null).ToDictionary(r => r.Id, StringComparer.Ordinal));
        }

        return paragraphs;
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
        DocxNumberingSet numbering)
    {
        return ReadRelatedStories(package, documentPartName, styles, numbering, CommentsRelationshipType, CommentsContentType, "Comment", "comment")
            .Concat(ReadRelatedStories(package, documentPartName, styles, numbering, FootnotesRelationshipType, FootnotesContentType, "Footnote", "footnote"))
            .Concat(ReadRelatedStories(package, documentPartName, styles, numbering, EndnotesRelationshipType, EndnotesContentType, "Endnote", "endnote"))
            .ToArray();
    }

    private static IReadOnlyList<DocxRelatedStory> ReadRelatedStories(
        OoxPackage package,
        string documentPartName,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        string relationshipType,
        string contentType,
        string kind,
        string storyElementName)
    {
        OoxPart? part = FindRelatedPart(package, documentPartName, relationshipType, contentType);
        if (part is null)
        {
            return [];
        }

        using Stream stream = part.OpenRead();
        XDocument partXml = SafeXml.Load(stream);
        IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(part.Name)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        var numberingCounters = new Dictionary<(string NumId, int Level), int>();
        return partXml.Root?
            .Elements(WordprocessingNamespace + storyElementName)
            .Select(story => ReadRelatedStory(kind, part.Name, story, styles, numbering, numberingCounters, package, relationships))
            .Where(story => story.BodyElements.Count > 0)
            .ToArray() ?? [];
    }

    private static DocxRelatedStory ReadRelatedStory(
        string kind,
        string partName,
        XElement story,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        IReadOnlyList<DocxBodyElement> bodyElements = ReadRelatedStoryBodyElements(
            story.Elements(),
            styles,
            numbering,
            numberingCounters,
            package,
            relationships);
        return new DocxRelatedStory(
            kind,
            partName,
            (string?)story.Attribute(WordprocessingNamespace + "id"),
            bodyElements,
            bodyElements.OfType<DocxParagraphElement>().Select(element => element.Paragraph).ToArray(),
            bodyElements.OfType<DocxTableElement>().Select(element => element.Table).ToArray());
    }

    private static IReadOnlyList<DocxBodyElement> ReadRelatedStoryBodyElements(
        IEnumerable<XElement> elements,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var bodyElements = new List<DocxBodyElement>();
        foreach (XElement element in elements)
        {
            if (element.Name == WordprocessingNamespace + "p")
            {
                DocxParagraph? paragraph = ReadParagraph(element, styles, numbering, numberingCounters, package, relationships);
                if (paragraph is not null)
                {
                    bodyElements.Add(new DocxParagraphElement(paragraph));
                }
            }
            else if (element.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? table = ReadTable(element, styles, numbering, numberingCounters, package, relationships);
                if (table is not null)
                {
                    bodyElements.Add(new DocxTableElement(table));
                }
            }
        }

        return bodyElements;
    }

    private static IReadOnlyList<DocxParagraph> ReadParagraphElements(
        IEnumerable<XElement> paragraphElements,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var wrapper = new XDocument(new XElement(WordprocessingNamespace + "document", new XElement(WordprocessingNamespace + "body", paragraphElements)));
        return ReadParagraphs(wrapper, styles, numbering, package, relationships);
    }

    private static DocxTable? ReadTable(
        XElement table,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
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
        XElement[] rowElements = table.Elements(WordprocessingNamespace + "tr").ToArray();
        for (int rowIndex = 0; rowIndex < rowElements.Length; rowIndex++)
        {
            XElement row = rowElements[rowIndex];
            DocxTableCellMargins rowExceptionMargins = ReadTablePropertyExceptionCellMargins(row);
            DocxTableCellMargins rowInheritedMargins = MergeTableCellMargins(rowExceptionMargins, MergeTableCellMargins(tableCellMargins, tableStyle.Cell.Margins));
            var cells = new List<DocxTableCell>();
            XElement[] cellElements = row.Elements(WordprocessingNamespace + "tc").ToArray();
            for (int cellIndex = 0; cellIndex < cellElements.Length; cellIndex++)
            {
                XElement cell = cellElements[cellIndex];
                XElement? cellProperties = cell.Element(WordprocessingNamespace + "tcPr");
                DocxTableCellConditionalFormat? conditionalFormat = ReadTableCellConditionalFormat(cellProperties);
                DocxTableCellStyle conditionalStyle = ResolveTableCellStyle(tableStyle, tableLook, conditionalFormat, rowIndex, cellIndex, rowElements.Length, cellElements.Length);
                IReadOnlyList<DocxParagraph> paragraphs = ReadTableCellParagraphs(
                    cell,
                    styles,
                    numbering,
                    numberingCounters,
                    package,
                    relationships,
                    conditionalStyle);
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
                    (string?)verticalMerge?.Attribute(WordprocessingNamespace + "val")));
            }

            if (cells.Count > 0)
            {
                XElement? header = row
                    .Element(WordprocessingNamespace + "trPr")
                    ?.Element(WordprocessingNamespace + "tblHeader");
                XElement? cantSplit = row
                    .Element(WordprocessingNamespace + "trPr")
                    ?.Element(WordprocessingNamespace + "cantSplit");
                XElement? rowHeight = row
                    .Element(WordprocessingNamespace + "trPr")
                    ?.Element(WordprocessingNamespace + "trHeight");
                rows.Add(new DocxTableRow(
                    cells,
                    ReadTableRowHeight(rowHeight),
                    ReadOnOff(header) == true,
                    (string?)header?.Attribute(WordprocessingNamespace + "val"),
                    (string?)rowHeight?.Attribute(WordprocessingNamespace + "val"),
                    (string?)rowHeight?.Attribute(WordprocessingNamespace + "hRule"),
                    HasAnyTableCellMargin(rowExceptionMargins) ? rowExceptionMargins : null,
                    ReadOnOff(cantSplit) == true,
                    (string?)cantSplit?.Attribute(WordprocessingNamespace + "val")));
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
            hasExplicitGrid);
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

    private static IReadOnlyList<DocxParagraph> ReadTableCellParagraphs(
        XElement cell,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        Dictionary<(string NumId, int Level), int> numberingCounters,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxTableCellStyle tableCellStyle)
    {
        var paragraphs = new List<DocxParagraph>();
        foreach (XElement paragraph in cell.Elements(WordprocessingNamespace + "p"))
        {
            DocxParagraph? parsed = ReadParagraph(
                paragraph,
                styles,
                numbering,
                numberingCounters,
                package,
                relationships,
                tableCellStyle);
            if (parsed is not null)
            {
                paragraphs.Add(parsed);
            }
        }

        return paragraphs;
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

    private static IReadOnlyList<DocxInlineImage> ReadInlineImages(XElement run, OoxPackage package, IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var images = new List<DocxInlineImage>();
        foreach (XElement inline in run.Descendants(WordprocessingDrawingNamespace + "inline"))
        {
            DocxInlineImage? image = ReadDrawingImage(inline, package, relationships);
            if (image is null)
            {
                continue;
            }

            images.Add(image);
        }

        return images;
    }

    private static DocxInlineImage? ReadDrawingImage(XElement drawing, OoxPackage package, IReadOnlyDictionary<string, OoxRelationship> relationships)
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
            imagePart.Name);
    }

    private static string? ReadDrawingImageRelationshipId(XElement drawing)
    {
        return (string?)drawing
            .Descendants(DrawingNamespace + "blip")
            .FirstOrDefault()
            ?.Attribute(RelationshipsNamespace + "embed");
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
        if (numId is null || !numbering.NumToAbstract.TryGetValue(numId, out string? abstractId) || !numbering.Levels.TryGetValue((abstractId, level), out DocxNumberingLevel? numberingLevel))
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

    private static DocxStyleSet LoadStyles(OoxPackage package, string documentPartName)
    {
        OoxRelationship? styleRelationship = package.GetRelationships(documentPartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == StylesRelationshipType && r.ResolvedTarget is not null);
        OoxPart? stylesPart = styleRelationship?.ResolvedTarget is null
            ? package.Parts.FirstOrDefault(p => p.ContentType == StylesContentType)
            : package.GetPart(styleRelationship.ResolvedTarget);
        if (stylesPart is null)
        {
            return DocxStyleSet.Empty;
        }

        using Stream stream = stylesPart.OpenRead();
        XDocument stylesXml = SafeXml.Load(stream);
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
        return new DocxStyleSet(runDefaults, paragraphDefaults, paragraphStyles, characterStyles, resolvedTableStyles, defaultTableStyle);
    }

    private static DocxNumberingSet LoadNumbering(OoxPackage package, string documentPartName, DocxFontCatalog fontCatalog)
    {
        OoxRelationship? numberingRelationship = package.GetRelationships(documentPartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == NumberingRelationshipType && r.ResolvedTarget is not null);
        OoxPart? numberingPart = numberingRelationship?.ResolvedTarget is null
            ? package.Parts.FirstOrDefault(p => p.ContentType == NumberingContentType)
            : package.GetPart(numberingRelationship.ResolvedTarget);
        if (numberingPart is null)
        {
            return DocxNumberingSet.Empty;
        }

        using Stream stream = numberingPart.OpenRead();
        XDocument numberingXml = SafeXml.Load(stream);
        var levels = new Dictionary<(string AbstractId, int Level), DocxNumberingLevel>();
        foreach (XElement abstractNum in numberingXml.Root?.Elements(WordprocessingNamespace + "abstractNum") ?? [])
        {
            string? abstractId = (string?)abstractNum.Attribute(WordprocessingNamespace + "abstractNumId");
            if (abstractId is null)
            {
                continue;
            }

            foreach (XElement level in abstractNum.Elements(WordprocessingNamespace + "lvl"))
            {
                int levelIndex = level.Attribute(WordprocessingNamespace + "ilvl") is { } ilvl
                    ? int.Parse(ilvl.Value, CultureInfo.InvariantCulture)
                    : 0;
                int start = level.Element(WordprocessingNamespace + "start")?.Attribute(WordprocessingNamespace + "val") is { } startValue
                    ? int.Parse(startValue.Value, CultureInfo.InvariantCulture)
                    : 1;
                string format = (string?)level.Element(WordprocessingNamespace + "numFmt")?.Attribute(WordprocessingNamespace + "val") ?? "decimal";
                string text = (string?)level.Element(WordprocessingNamespace + "lvlText")?.Attribute(WordprocessingNamespace + "val") ??
                    (format.Equals("bullet", StringComparison.OrdinalIgnoreCase) ? "\u2022" : "%" + (levelIndex + 1) + ".");
                string suffix = (string?)level.Element(WordprocessingNamespace + "suff")?.Attribute(WordprocessingNamespace + "val") ?? "tab";
                DocxTextRunStyle style = ReadTextRunStyle(level.Element(WordprocessingNamespace + "rPr"));
                levels[(abstractId, levelIndex)] = new DocxNumberingLevel(format, ResolveNumberingSymbolText(text, style, fontCatalog), suffix, start, ReadNumberingIndent(level), style);
            }
        }

        var numToAbstract = new Dictionary<string, string>(StringComparer.Ordinal);
        var startOverrides = new Dictionary<(string NumId, int Level), int>();
        foreach (XElement num in numberingXml.Root?.Elements(WordprocessingNamespace + "num") ?? [])
        {
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
                int levelIndex = overrideLevel.Attribute(WordprocessingNamespace + "ilvl") is { } ilvl
                    ? int.Parse(ilvl.Value, CultureInfo.InvariantCulture)
                    : 0;
                if (overrideLevel.Element(WordprocessingNamespace + "startOverride")?.Attribute(WordprocessingNamespace + "val") is { } startValue)
                {
                    startOverrides[(numId, levelIndex)] = int.Parse(startValue.Value, CultureInfo.InvariantCulture);
                }
            }
        }

        return new DocxNumberingSet(numToAbstract, levels, startOverrides);
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
            (string?)snapToGrid?.Attribute(WordprocessingNamespace + "val"));
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
        string? underlineValue = (string?)properties?
            .Element(WordprocessingNamespace + "u")
            ?.Attribute(WordprocessingNamespace + "val");
        bool? underline = properties?.Element(WordprocessingNamespace + "u") is not null
            ? !string.Equals(underlineValue, "none", StringComparison.OrdinalIgnoreCase)
            : null;
        return new DocxResolvedRunProperties(fontSize, color, fontFamily, bold, italic, complexScriptBold, complexScriptItalic, underline, underlineValue, ReadRunFonts(properties), characterSpacingPoints, allCaps, verticalAlignmentValue, strike, strikeValue, doubleStrike, doubleStrikeValue, highlightValue, shadingFill, shadingValue, shadingColor, smallCaps, smallCapsValue, hidden, hiddenValue);
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

    private static OoxPart FindDocumentPart(OoxPackage package)
    {
        OoxRelationship? packageRelationship = package.GetRelationships("/")
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
        DocxTableStyle? DefaultTableStyle)
    {
        public static DocxStyleSet Empty { get; } = new(
            new DocxResolvedRunProperties(null, null, null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null),
            new DocxResolvedParagraphProperties(null, null, null, null, null, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, DocxParagraphIndent.Empty, [], null, null),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxTableStyle>(),
            null);
    }

    private sealed record DocxStyle(string? BasedOnStyleId, DocxResolvedParagraphProperties Paragraph, DocxResolvedRunProperties Run);

    private sealed record ParagraphPageBreakPart(
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
        DocxTableCellMargins Margins)
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
                MergeTableCellMargins(other.Margins, Margins));
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

        return resolved with { BasedOnStyleId = null };
    }

    private static DocxTableCellStyle ReadTableCellStyle(
        XElement? cellProperties,
        XElement? paragraphProperties = null,
        XElement? runProperties = null)
    {
        XElement? shading = cellProperties?.Element(WordprocessingNamespace + "shd");
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
            DocxTableCellMargins.Empty);
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
        IReadOnlyDictionary<(string NumId, int Level), int> StartOverrides)
    {
        public static DocxNumberingSet Empty { get; } = new(
            new Dictionary<string, string>(),
            new Dictionary<(string AbstractId, int Level), DocxNumberingLevel>(),
            new Dictionary<(string NumId, int Level), int>());
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
            run.HiddenValue);
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
        string? SnapToGridValue)
    {
        public static DocxResolvedParagraphProperties Empty { get; } = new(null, null, null, null, null, null, DocxParagraphSpacing.Empty, DocxParagraphKeepRules.Empty, DocxParagraphIndent.Empty, [], null, null);

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
                other.SnapToGridValue ?? SnapToGridValue);
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
        string? HiddenValue = null)
    {
        public static DocxResolvedRunProperties Empty { get; } = new(null, null, null, null, null, null, null, null, null, DocxRunFonts.Empty, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

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
                other.HiddenValue ?? HiddenValue);
        }
    }
}
