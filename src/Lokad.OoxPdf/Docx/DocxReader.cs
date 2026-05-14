using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Docx;

internal sealed class DocxReader
{
    private static readonly XNamespace WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string MainDocumentContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string OfficeDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

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

        IReadOnlyList<DocxParagraph> paragraphs = ReadParagraphs(document);

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

    private static IReadOnlyList<DocxParagraph> ReadParagraphs(XDocument document)
    {
        var paragraphs = new List<DocxParagraph>();
        foreach (XElement paragraph in document.Descendants(WordprocessingNamespace + "body").Elements(WordprocessingNamespace + "p"))
        {
            var runs = new List<DocxTextRun>();
            foreach (XElement run in paragraph.Elements(WordprocessingNamespace + "r"))
            {
                string text = string.Concat(run.Elements(WordprocessingNamespace + "t").Select(t => (string?)t ?? string.Empty));
                if (text.Length == 0)
                {
                    continue;
                }

                XElement? runProperties = run.Element(WordprocessingNamespace + "rPr");
                double fontSize = runProperties?
                    .Element(WordprocessingNamespace + "sz")
                    ?.Attribute(WordprocessingNamespace + "val") is { } size
                    ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 2d
                    : 11d;
                string? color = (string?)runProperties?
                    .Element(WordprocessingNamespace + "color")
                    ?.Attribute(WordprocessingNamespace + "val");
                string? fontFamily = (string?)runProperties?
                    .Element(WordprocessingNamespace + "rFonts")
                    ?.Attribute(WordprocessingNamespace + "ascii");
                bool bold = runProperties?.Element(WordprocessingNamespace + "b") is not null;
                bool italic = runProperties?.Element(WordprocessingNamespace + "i") is not null;
                bool underline = runProperties?.Element(WordprocessingNamespace + "u") is not null;
                runs.Add(new DocxTextRun(text, fontSize, color, bold, italic, underline, fontFamily));
            }

            if (runs.Count > 0)
            {
                paragraphs.Add(new DocxParagraph(runs, ReadAlignment(paragraph)));
            }
        }

        return paragraphs;
    }

    private static DocxTextAlignment ReadAlignment(XElement paragraph)
    {
        string? value = (string?)paragraph
            .Element(WordprocessingNamespace + "pPr")
            ?.Element(WordprocessingNamespace + "jc")
            ?.Attribute(WordprocessingNamespace + "val");
        return value switch
        {
            "center" => DocxTextAlignment.Center,
            "right" => DocxTextAlignment.Right,
            _ => DocxTextAlignment.Left
        };
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
}
