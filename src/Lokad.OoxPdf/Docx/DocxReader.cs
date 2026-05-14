using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxReader
{
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
    private const string StylesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    private const string NumberingContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

    public DocxDocument Read(OoxPackage package, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        OoxPart documentPart = FindDocumentPart(package);
        using Stream stream = documentPart.OpenRead();
        XDocument document = SafeXml.Load(stream);
        EmitUnsupportedFeatureDiagnostics(package, document, documentPart.Name, diagnosticSink);

        XElement? pageSize = document
            .Descendants(WordprocessingNamespace + "sectPr")
            .Elements(WordprocessingNamespace + "pgSz")
            .LastOrDefault();
        XElement? pageMargins = document
            .Descendants(WordprocessingNamespace + "sectPr")
            .Elements(WordprocessingNamespace + "pgMar")
            .LastOrDefault();

        DocxStyleSet styles = LoadStyles(package, documentPart.Name);
        DocxNumberingSet numbering = LoadNumbering(package, documentPart.Name);
        IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(documentPart.Name)
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        IReadOnlyList<DocxBodyElement> bodyElements = ReadBodyElements(document, styles, numbering, package, relationships);
        IReadOnlyList<DocxParagraph> paragraphs = bodyElements.OfType<DocxParagraphElement>().Select(e => e.Paragraph).ToArray();
        IReadOnlyList<DocxTable> tables = bodyElements.OfType<DocxTableElement>().Select(e => e.Table).ToArray();
        IReadOnlyList<DocxParagraph> headers = ReadReferencedHeaderFooterParagraphs(document, package, relationships, styles, numbering, HeaderRelationshipType, "headerReference");
        IReadOnlyList<DocxParagraph> footers = ReadReferencedHeaderFooterParagraphs(document, package, relationships, styles, numbering, FooterRelationshipType, "footerReference");

        if (pageSize is null)
        {
            return new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, headers, footers, bodyElements, paragraphs, tables);
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
        return new DocxDocument(width, height, left, right, top, bottom, headers, footers, bodyElements, paragraphs, tables);
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
        void Emit(string id, string feature)
        {
            if (!emitted.Add(id))
            {
                return;
            }

            diagnosticSink(new OoxPdfDiagnostic(
                id,
                OoxPdfSeverity.Warning,
                $"Unsupported DOCX feature '{feature}' was detected and ignored or approximated.",
                partName,
                Feature: feature,
                Fallback: "Ignored"));
        }

        if (document.Descendants(WordprocessingNamespace + "commentRangeStart").Any() ||
            document.Descendants(WordprocessingNamespace + "commentReference").Any())
        {
            Emit("DOCX_UNSUPPORTED_COMMENTS", "comments");
        }

        if (document.Descendants(WordprocessingNamespace + "ins").Any() ||
            document.Descendants(WordprocessingNamespace + "del").Any())
        {
            Emit("DOCX_UNSUPPORTED_TRACKED_CHANGES", "tracked changes");
        }

        if (document.Descendants(WordprocessingNamespace + "fldChar").Any() ||
            document.Descendants(WordprocessingNamespace + "instrText").Any(instruction => !(((string?)instruction)?.Contains("PAGE", StringComparison.OrdinalIgnoreCase) == true)))
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
            Emit("DOCX_UNSUPPORTED_FOOTNOTE", "footnote");
        }

        if (document.Descendants(WordprocessingNamespace + "endnoteReference").Any())
        {
            Emit("DOCX_UNSUPPORTED_ENDNOTE", "endnote");
        }

        if (document.Descendants(WordprocessingNamespace + "cols").Any(cols =>
            cols.Attribute(WordprocessingNamespace + "num") is { } value &&
            int.TryParse(value.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int columns) &&
            columns > 1))
        {
            Emit("DOCX_UNSUPPORTED_MULTI_COLUMN", "multi-column section");
        }

        if (document.Descendants(WordprocessingNamespace + "br").Any(br =>
            string.Equals((string?)br.Attribute(WordprocessingNamespace + "type"), "column", StringComparison.OrdinalIgnoreCase)))
        {
            Emit("DOCX_UNSUPPORTED_MANUAL_BREAK", "manual column break");
        }

        if (document.Descendants(WordprocessingNamespace + "keepLines").Any() ||
            document.Descendants(WordprocessingNamespace + "widowControl").Any())
        {
            Emit("DOCX_UNSUPPORTED_PARAGRAPH_KEEP_RULE", "paragraph keep-lines/widow-orphan rule");
        }

        if (document.Descendants(WordprocessingNamespace + "pPr").Elements(WordprocessingNamespace + "sectPr").Any())
        {
            Emit("DOCX_UNSUPPORTED_SECTION_BREAK", "paragraph section break");
        }

        if (package.Parts.Any(p => p.Name.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase) ||
            p.ContentType.Contains("vbaProject", StringComparison.OrdinalIgnoreCase)))
        {
            Emit("DOCX_UNSUPPORTED_MACRO", "macro");
        }
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
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
        string? paragraphStyleId = (string?)paragraphProperties?
            .Element(WordprocessingNamespace + "pStyle")
            ?.Attribute(WordprocessingNamespace + "val");
        DocxResolvedParagraphProperties resolvedParagraph = ResolveParagraphProperties(paragraphProperties, paragraphStyleId, styles);
        var runs = new List<DocxTextRun>();
        var images = new List<DocxInlineImage>();
        bool pageInstructionSeen = false;
        foreach (XElement run in paragraph.Elements(WordprocessingNamespace + "r"))
        {
            string text = string.Concat(run.Elements(WordprocessingNamespace + "t").Select(t => (string?)t ?? string.Empty));
            if (run.Elements(WordprocessingNamespace + "instrText").Any(instruction => ((string?)instruction)?.Contains("PAGE", StringComparison.OrdinalIgnoreCase) == true))
            {
                text = "{PAGE}";
                pageInstructionSeen = true;
            }
            else if (pageInstructionSeen && text.Trim().All(char.IsDigit))
            {
                text = string.Empty;
            }

            XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
            string? characterStyleId = (string?)runProperties?
                .Element(WordprocessingNamespace + "rStyle")
                ?.Attribute(WordprocessingNamespace + "val");
            DocxResolvedRunProperties resolvedRun = ResolveRunProperties(runProperties, paragraphStyleId, characterStyleId, styles);
            if (text.Length != 0)
            {
                runs.Add(new DocxTextRun(
                    text,
                    resolvedRun.FontSize ?? 11d,
                    resolvedRun.ColorHex,
                    resolvedRun.Bold ?? false,
                    resolvedRun.Italic ?? false,
                    resolvedRun.Underline ?? false,
                    resolvedRun.FontFamily));
            }

            images.AddRange(ReadInlineImages(run, package, relationships));
        }

        foreach (XElement field in paragraph.Elements(WordprocessingNamespace + "fldSimple"))
        {
            string? instruction = (string?)field.Attribute(WordprocessingNamespace + "instr");
            if (instruction?.Contains("PAGE", StringComparison.OrdinalIgnoreCase) == true)
            {
                runs.Add(new DocxTextRun("{PAGE}", 11d, null, false, false, false, null));
            }
        }

        if (runs.Count == 0 && images.Count == 0)
        {
            return null;
        }

        return new DocxParagraph(
            runs,
            images,
            resolvedParagraph.Alignment ?? DocxTextAlignment.Left,
            resolvedParagraph.SpacingBeforePoints ?? 0d,
            resolvedParagraph.SpacingAfterPoints ?? 6d,
            resolvedParagraph.LineSpacingFactor ?? 1.25d,
            resolvedParagraph.LineSpacingPoints,
            resolvedParagraph.KeepWithNext ?? false,
            CreateListLabel(paragraphProperties, numbering, numberingCounters));
    }

    private static IReadOnlyList<DocxBodyElement> ReadBodyElements(
        XDocument document,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships)
    {
        var elements = new List<DocxBodyElement>();
        var numberingCounters = new Dictionary<(string NumId, int Level), int>();
        foreach (XElement element in document.Descendants(WordprocessingNamespace + "body").Elements())
        {
            if (element.Name == WordprocessingNamespace + "p")
            {
                if (element
                    .Element(WordprocessingNamespace + "pPr")
                    ?.Element(WordprocessingNamespace + "pageBreakBefore") is not null)
                {
                    elements.Add(new DocxPageBreakElement());
                }

                foreach (XElement? fragment in SplitParagraphByManualPageBreaks(element))
                {
                    if (fragment is null)
                    {
                        elements.Add(new DocxPageBreakElement());
                        continue;
                    }

                    DocxParagraph? paragraph = ReadParagraph(fragment, styles, numbering, numberingCounters, package, relationships);
                    if (paragraph is not null)
                    {
                        elements.Add(new DocxParagraphElement(paragraph));
                    }
                }
            }
            else if (element.Name == WordprocessingNamespace + "tbl")
            {
                DocxTable? table = ReadTable(element);
                if (table is not null)
                {
                    elements.Add(new DocxTableElement(table));
                }
            }
        }

        return elements;
    }

    private static IEnumerable<XElement?> SplitParagraphByManualPageBreaks(XElement paragraph)
    {
        XElement CreateParagraph() => new(
            paragraph.Name,
            paragraph.Attributes(),
            paragraph.Element(WordprocessingNamespace + "pPr") is { } properties ? new XElement(properties) : null);

        XElement currentParagraph = CreateParagraph();
        foreach (XElement child in paragraph.Elements().Where(e => e.Name != WordprocessingNamespace + "pPr"))
        {
            if (child.Name != WordprocessingNamespace + "r")
            {
                currentParagraph.Add(new XElement(child));
                continue;
            }

            XElement currentRun = CreateRunShell(child);
            foreach (XElement runChild in child.Elements())
            {
                if (runChild.Name == WordprocessingNamespace + "br" &&
                    string.Equals((string?)runChild.Attribute(WordprocessingNamespace + "type"), "page", StringComparison.OrdinalIgnoreCase))
                {
                    AddRunIfNotEmpty(currentParagraph, currentRun);
                    yield return currentParagraph;
                    yield return null;
                    currentParagraph = CreateParagraph();
                    currentRun = CreateRunShell(child);
                    continue;
                }

                currentRun.Add(new XElement(runChild));
            }

            AddRunIfNotEmpty(currentParagraph, currentRun);
        }

        yield return currentParagraph;
    }

    private static XElement CreateRunShell(XElement run)
    {
        var shell = new XElement(run.Name, run.Attributes());
        if (run.Element(WordprocessingNamespace + "rPr") is { } properties)
        {
            shell.Add(new XElement(properties));
        }

        return shell;
    }

    private static void AddRunIfNotEmpty(XElement paragraph, XElement run)
    {
        if (run.Elements().Any(e => e.Name != WordprocessingNamespace + "rPr"))
        {
            paragraph.Add(run);
        }
    }

    private static IReadOnlyList<DocxParagraph> ReadReferencedHeaderFooterParagraphs(
        XDocument document,
        OoxPackage package,
        IReadOnlyDictionary<string, OoxRelationship> relationships,
        DocxStyleSet styles,
        DocxNumberingSet numbering,
        string relationshipType,
        string referenceElementName)
    {
        var paragraphs = new List<DocxParagraph>();
        foreach (XElement reference in document.Descendants(WordprocessingNamespace + referenceElementName))
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
            paragraphs.AddRange(ReadParagraphElements(partXml.Root?.Elements(WordprocessingNamespace + "p") ?? [], styles, numbering, package, package.GetRelationships(part.Name).Where(r => !r.IsExternal && r.ResolvedTarget is not null).ToDictionary(r => r.Id, StringComparer.Ordinal)));
        }

        return paragraphs;
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

    private static IReadOnlyList<DocxTable> ReadTables(XDocument document)
    {
        var tables = new List<DocxTable>();
        foreach (XElement table in document.Descendants(WordprocessingNamespace + "body").Elements(WordprocessingNamespace + "tbl"))
        {
            DocxTable? parsed = ReadTable(table);
            if (parsed is not null)
            {
                tables.Add(parsed);
            }
        }

        return tables;
    }

    private static DocxTable? ReadTable(XElement table)
    {
        IReadOnlyList<double> columns = table
            .Element(WordprocessingNamespace + "tblGrid")
            ?.Elements(WordprocessingNamespace + "gridCol")
            .Select(column => ReadTwipsAttribute(column, WordprocessingNamespace + "w") ?? 72d)
            .ToArray() ?? [];
        var rows = new List<DocxTableRow>();
        foreach (XElement row in table.Elements(WordprocessingNamespace + "tr"))
        {
            var cells = new List<DocxTableCell>();
            foreach (XElement cell in row.Elements(WordprocessingNamespace + "tc"))
            {
                string text = string.Join(" ", cell
                    .Descendants(WordprocessingNamespace + "p")
                    .Select(p => string.Concat(p.Descendants(WordprocessingNamespace + "t").Select(t => (string?)t ?? string.Empty)))
                    .Where(t => t.Length != 0));
                string? fill = (string?)cell
                    .Element(WordprocessingNamespace + "tcPr")
                    ?.Element(WordprocessingNamespace + "shd")
                    ?.Attribute(WordprocessingNamespace + "fill");
                cells.Add(new DocxTableCell(text, fill));
            }

            if (cells.Count > 0)
            {
                rows.Add(new DocxTableRow(cells, ReadTableRowHeight(row)));
            }
        }

        if (rows.Count == 0)
        {
            return null;
        }

        if (columns.Count == 0)
        {
            int maxCells = rows.Max(r => r.Cells.Count);
            columns = Enumerable.Repeat(72d, maxCells).ToArray();
        }

        return new DocxTable(columns, rows);
    }

    private static double? ReadTableRowHeight(XElement row)
    {
        XElement? height = row
            .Element(WordprocessingNamespace + "trPr")
            ?.Element(WordprocessingNamespace + "trHeight");
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
            XElement? extent = inline.Element(WordprocessingDrawingNamespace + "extent");
            string? relationshipId = (string?)inline
                .Descendants(DrawingNamespace + "blip")
                .FirstOrDefault()
                ?.Attribute(RelationshipsNamespace + "embed");
            if (extent is null || relationshipId is null || !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) || relationship.ResolvedTarget is null)
            {
                continue;
            }

            OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
            if (imagePart is null)
            {
                continue;
            }

            images.Add(new DocxInlineImage(
                OoxUnits.EmuToPoints(ParseLongAttribute(extent, "cx")),
                OoxUnits.EmuToPoints(ParseLongAttribute(extent, "cy")),
                imagePart.ContentType,
                imagePart.Bytes,
                imagePart.Name));
        }

        return images;
    }

    private static string? CreateListLabel(XElement? paragraphProperties, DocxNumberingSet numbering, Dictionary<(string NumId, int Level), int> counters)
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
            return "\u2022";
        }

        var key = (numId, level);
        counters[key] = counters.TryGetValue(key, out int current) ? current + 1 : numberingLevel.Start;
        string numberText = counters[key].ToString(CultureInfo.InvariantCulture);
        return numberingLevel.Text.Replace("%" + (level + 1).ToString(CultureInfo.InvariantCulture), numberText, StringComparison.Ordinal);
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
        foreach (XElement style in stylesXml.Root?.Elements(WordprocessingNamespace + "style") ?? [])
        {
            string? styleId = (string?)style.Attribute(WordprocessingNamespace + "styleId");
            string? type = (string?)style.Attribute(WordprocessingNamespace + "type");
            if (string.IsNullOrWhiteSpace(styleId))
            {
                continue;
            }

            var parsed = new DocxStyle(
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
        }

        return new DocxStyleSet(runDefaults, paragraphDefaults, paragraphStyles, characterStyles);
    }

    private static DocxNumberingSet LoadNumbering(OoxPackage package, string documentPartName)
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
                string text = (string?)level.Element(WordprocessingNamespace + "lvlText")?.Attribute(WordprocessingNamespace + "val") ?? "%" + (levelIndex + 1) + ".";
                levels[(abstractId, levelIndex)] = new DocxNumberingLevel(format, text, start);
            }
        }

        var numToAbstract = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (XElement num in numberingXml.Root?.Elements(WordprocessingNamespace + "num") ?? [])
        {
            string? numId = (string?)num.Attribute(WordprocessingNamespace + "numId");
            string? abstractId = (string?)num.Element(WordprocessingNamespace + "abstractNumId")?.Attribute(WordprocessingNamespace + "val");
            if (numId is not null && abstractId is not null)
            {
                numToAbstract[numId] = abstractId;
            }
        }

        return new DocxNumberingSet(numToAbstract, levels);
    }

    private static DocxResolvedParagraphProperties ResolveParagraphProperties(XElement? directProperties, string? paragraphStyleId, DocxStyleSet styles)
    {
        DocxResolvedParagraphProperties result = styles.ParagraphDefaults;
        if (paragraphStyleId is not null && styles.ParagraphStyles.TryGetValue(paragraphStyleId, out DocxStyle? style))
        {
            result = result.Merge(style.Paragraph);
        }

        return result.Merge(ReadParagraphProperties(directProperties));
    }

    private static DocxResolvedRunProperties ResolveRunProperties(XElement? directProperties, string? paragraphStyleId, string? characterStyleId, DocxStyleSet styles)
    {
        DocxResolvedRunProperties result = styles.RunDefaults;
        if (paragraphStyleId is not null && styles.ParagraphStyles.TryGetValue(paragraphStyleId, out DocxStyle? paragraphStyle))
        {
            result = result.Merge(paragraphStyle.Run);
        }

        if (characterStyleId is not null && styles.CharacterStyles.TryGetValue(characterStyleId, out DocxStyle? characterStyle))
        {
            result = result.Merge(characterStyle.Run);
        }

        return result.Merge(ReadRunProperties(directProperties));
    }

    private static DocxResolvedParagraphProperties ReadParagraphProperties(XElement? properties)
    {
        DocxTextAlignment? alignment = ReadAlignment(properties);
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

        bool? keepWithNext = ReadOnOff(properties?.Element(WordprocessingNamespace + "keepNext"));
        return new DocxResolvedParagraphProperties(alignment, before, after, lineFactor, linePoints, keepWithNext);
    }

    private static DocxTextAlignment? ReadAlignment(XElement? properties)
    {
        string? value = (string?)properties
            ?.Element(WordprocessingNamespace + "jc")
            ?.Attribute(WordprocessingNamespace + "val");
        return value switch
        {
            "center" => DocxTextAlignment.Center,
            "right" => DocxTextAlignment.Right,
            "both" => DocxTextAlignment.Left,
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
        bool? underline = properties?.Element(WordprocessingNamespace + "u") is { } underlineElement
            ? !string.Equals((string?)underlineElement.Attribute(WordprocessingNamespace + "val"), "none", StringComparison.OrdinalIgnoreCase)
            : null;
        return new DocxResolvedRunProperties(fontSize, color, fontFamily, bold, italic, underline);
    }

    private static bool? ReadOnOff(XElement? element)
    {
        if (element is null)
        {
            return null;
        }

        string? value = (string?)element.Attribute(WordprocessingNamespace + "val");
        return value is null || value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase);
    }

    private static double? ReadTwipsAttribute(XElement? element, XName name)
    {
        return element?.Attribute(name) is { } value
            ? OoxUnits.TwipsToPoints(long.Parse(value.Value, CultureInfo.InvariantCulture))
            : null;
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
        IReadOnlyDictionary<string, DocxStyle> CharacterStyles)
    {
        public static DocxStyleSet Empty { get; } = new(
            new DocxResolvedRunProperties(null, null, null, null, null, null),
            new DocxResolvedParagraphProperties(null, null, null, null, null, null),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxStyle>());
    }

    private sealed record DocxStyle(DocxResolvedParagraphProperties Paragraph, DocxResolvedRunProperties Run);

    private sealed record DocxNumberingSet(
        IReadOnlyDictionary<string, string> NumToAbstract,
        IReadOnlyDictionary<(string AbstractId, int Level), DocxNumberingLevel> Levels)
    {
        public static DocxNumberingSet Empty { get; } = new(
            new Dictionary<string, string>(),
            new Dictionary<(string AbstractId, int Level), DocxNumberingLevel>());
    }

    private sealed record DocxNumberingLevel(string Format, string Text, int Start);

    private readonly record struct DocxResolvedParagraphProperties(
        DocxTextAlignment? Alignment,
        double? SpacingBeforePoints,
        double? SpacingAfterPoints,
        double? LineSpacingFactor,
        double? LineSpacingPoints,
        bool? KeepWithNext)
    {
        public DocxResolvedParagraphProperties Merge(DocxResolvedParagraphProperties other)
        {
            return new DocxResolvedParagraphProperties(
                other.Alignment ?? Alignment,
                other.SpacingBeforePoints ?? SpacingBeforePoints,
                other.SpacingAfterPoints ?? SpacingAfterPoints,
                other.LineSpacingFactor ?? LineSpacingFactor,
                other.LineSpacingPoints ?? LineSpacingPoints,
                other.KeepWithNext ?? KeepWithNext);
        }
    }

    private readonly record struct DocxResolvedRunProperties(
        double? FontSize,
        string? ColorHex,
        string? FontFamily,
        bool? Bold,
        bool? Italic,
        bool? Underline)
    {
        public DocxResolvedRunProperties Merge(DocxResolvedRunProperties other)
        {
            return new DocxResolvedRunProperties(
                other.FontSize ?? FontSize,
                other.ColorHex ?? ColorHex,
                other.FontFamily ?? FontFamily,
                other.Bold ?? Bold,
                other.Italic ?? Italic,
                other.Underline ?? Underline);
        }
    }
}
