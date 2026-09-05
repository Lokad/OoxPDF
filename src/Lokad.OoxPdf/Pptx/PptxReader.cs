using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed class PptxReader
{
    private const string PresentationContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    private const string OfficeDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string SlideRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    public PptxDocument Read(OoxPackage package, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart presentationPart = FindPresentationPart();
        using Stream stream = presentationPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);

        XElement? size = document.Root?.Element(PresentationNamespace + "sldSz");
        double width = size is null ? 720d : OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(size, "cx", "PPTX"));
        double height = size is null ? 540d : OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(size, "cy", "PPTX"));

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
            string? relationshipId = (string?)slideId.Attribute(RelationshipsNamespace + "id");
            if (relationshipId is not null && relationships.TryGetValue(relationshipId, out OoxRelationship? relationship))
            {
                slides.Add(new PptxSlide(relationship.ResolvedTarget ?? throw new InvalidDataException("Slide relationship has no resolved target part."), slides.Count));
            }
        }

        if (slides.Count == 0)
        {
            slides.AddRange(relationships.Values.Select(r => new PptxSlide(r.ResolvedTarget ?? throw new InvalidDataException("Slide relationship has no resolved target part."), slides.Count)));
        }

        return new PptxDocument(presentationPart.Name, slides, width, height);

        OoxPart FindPresentationPart()
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
    }


}
