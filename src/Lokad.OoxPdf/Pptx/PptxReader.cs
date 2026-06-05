using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed class PptxReader
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace RelationshipsDocumentNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string PresentationContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    private const string OfficeDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string SlideRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    public PptxDocument Read(OoxPackage package, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart presentationPart = FindPresentationPart(package, cancellationToken);
        using Stream stream = presentationPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);

        XElement? size = document.Root?.Element(PresentationNamespace + "sldSz");
        double width = size is null ? 720d : OoxUnits.EmuToPoints(ParseLongAttribute(size, "cx"));
        double height = size is null ? 540d : OoxUnits.EmuToPoints(ParseLongAttribute(size, "cy"));

        IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(presentationPart.Name, cancellationToken)
            .Where(r => !r.IsExternal && r.Type == SlideRelationshipType && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);

        var slides = new List<PptxSlide>();
        IEnumerable<XElement> slideIds = document.Root?
            .Element(PresentationNamespace + "sldIdLst")?
            .Elements(PresentationNamespace + "sldId") ?? [];

        foreach (XElement slideId in slideIds)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? relationshipId = (string?)slideId.Attribute(RelationshipsDocumentNamespace + "id");
            if (relationshipId is not null && relationships.TryGetValue(relationshipId, out OoxRelationship? relationship))
            {
                slides.Add(new PptxSlide(relationship.ResolvedTarget!, slides.Count));
            }
        }

        if (slides.Count == 0)
        {
            slides.AddRange(relationships.Values.Select(r => new PptxSlide(r.ResolvedTarget!, slides.Count)));
        }

        return new PptxDocument(presentationPart.Name, slides, width, height);
    }

    private static OoxPart FindPresentationPart(OoxPackage package, CancellationToken cancellationToken)
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

        OoxPart? contentTypePart = package.Parts.FirstOrDefault(p => p.ContentType == PresentationContentType);
        return contentTypePart ?? throw new InvalidDataException("PPTX package does not contain a presentation part.");
    }

    private static long ParseLongAttribute(XElement element, string name)
    {
        string value = (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required PPTX attribute '{name}'.");
        return long.Parse(value, CultureInfo.InvariantCulture);
    }
}
