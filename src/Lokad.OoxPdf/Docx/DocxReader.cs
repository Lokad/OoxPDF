using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxReader
{
    private static readonly XNamespace WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string MainDocumentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string OfficeDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string StylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    private const string StylesContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

    public DocxDocument Read(OoxPackage package)
    {
        OoxPart documentPart = FindDocumentPart(package);
        using Stream stream = documentPart.OpenRead();
        XDocument document = SafeXml.Load(stream);

        XElement? pageSize = document
            .Descendants(WordprocessingNamespace + "sectPr")
            .Elements(WordprocessingNamespace + "pgSz")
            .LastOrDefault();
        XElement? pageMargins = document
            .Descendants(WordprocessingNamespace + "sectPr")
            .Elements(WordprocessingNamespace + "pgMar")
            .LastOrDefault();

        DocxStyleSet styles = LoadStyles(package, documentPart.Name);
        IReadOnlyList<DocxParagraph> paragraphs = ReadParagraphs(document, styles);

        if (pageSize is null)
        {
            return new DocxDocument(612d, 792d, 72d, 72d, 72d, 72d, paragraphs);
        }

        double width = OoxUnits.TwipsToPoints(ParseLongAttribute(pageSize, WordprocessingNamespace + "w"));
        double height = OoxUnits.TwipsToPoints(ParseLongAttribute(pageSize, WordprocessingNamespace + "h"));
        string? orientation = (string?)pageSize.Attribute(WordprocessingNamespace + "orient");
        if (orientation?.Equals("landscape", StringComparison.OrdinalIgnoreCase) == true && height > width)
        {
            (width, height) = (height, width);
        }

        double left = ReadMargin(pageMargins, WordprocessingNamespace + "left", 72d);
        double right = ReadMargin(pageMargins, WordprocessingNamespace + "right", 72d);
        double top = ReadMargin(pageMargins, WordprocessingNamespace + "top", 72d);
        double bottom = ReadMargin(pageMargins, WordprocessingNamespace + "bottom", 72d);
        return new DocxDocument(width, height, left, right, top, bottom, paragraphs);
    }

    private static IReadOnlyList<DocxParagraph> ReadParagraphs(XDocument document, DocxStyleSet styles)
    {
        var paragraphs = new List<DocxParagraph>();
        foreach (XElement paragraph in document.Descendants(WordprocessingNamespace + "body").Elements(WordprocessingNamespace + "p"))
        {
            XElement? paragraphProperties = paragraph.Element(WordprocessingNamespace + "pPr");
            string? paragraphStyleId = (string?)paragraphProperties?
                .Element(WordprocessingNamespace + "pStyle")
                ?.Attribute(WordprocessingNamespace + "val");
            DocxResolvedParagraphProperties resolvedParagraph = ResolveParagraphProperties(paragraphProperties, paragraphStyleId, styles);
            var runs = new List<DocxTextRun>();
            foreach (XElement run in paragraph.Elements(WordprocessingNamespace + "r"))
            {
                string text = string.Concat(run.Elements(WordprocessingNamespace + "t").Select(t => (string?)t ?? string.Empty));
                if (text.Length == 0)
                {
                    continue;
                }

                XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
                string? characterStyleId = (string?)runProperties?
                    .Element(WordprocessingNamespace + "rStyle")
                    ?.Attribute(WordprocessingNamespace + "val");
                DocxResolvedRunProperties resolvedRun = ResolveRunProperties(runProperties, paragraphStyleId, characterStyleId, styles);
                runs.Add(new DocxTextRun(
                    text,
                    resolvedRun.FontSize ?? 11d,
                    resolvedRun.ColorHex,
                    resolvedRun.Bold ?? false,
                    resolvedRun.Italic ?? false,
                    resolvedRun.Underline ?? false,
                    resolvedRun.FontFamily));
            }

            if (runs.Count > 0)
            {
                paragraphs.Add(new DocxParagraph(
                    runs,
                    resolvedParagraph.Alignment ?? DocxTextAlignment.Left,
                    resolvedParagraph.SpacingBeforePoints ?? 0d,
                    resolvedParagraph.SpacingAfterPoints ?? 6d,
                    resolvedParagraph.LineSpacingFactor ?? 1.25d));
            }
        }

        return paragraphs;
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
        double? lineFactor = spacing?.Attribute(WordprocessingNamespace + "line") is { } line
            ? int.Parse(line.Value, CultureInfo.InvariantCulture) / 240d
            : null;
        return new DocxResolvedParagraphProperties(alignment, before, after, lineFactor);
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
            new DocxResolvedParagraphProperties(null, null, null, null),
            new Dictionary<string, DocxStyle>(),
            new Dictionary<string, DocxStyle>());
    }

    private sealed record DocxStyle(DocxResolvedParagraphProperties Paragraph, DocxResolvedRunProperties Run);

    private readonly record struct DocxResolvedParagraphProperties(
        DocxTextAlignment? Alignment,
        double? SpacingBeforePoints,
        double? SpacingAfterPoints,
        double? LineSpacingFactor)
    {
        public DocxResolvedParagraphProperties Merge(DocxResolvedParagraphProperties other)
        {
            return new DocxResolvedParagraphProperties(
                other.Alignment ?? Alignment,
                other.SpacingBeforePoints ?? SpacingBeforePoints,
                other.SpacingAfterPoints ?? SpacingAfterPoints,
                other.LineSpacingFactor ?? LineSpacingFactor);
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
