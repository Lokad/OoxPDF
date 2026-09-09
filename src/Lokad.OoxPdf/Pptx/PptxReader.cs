using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed class PptxReader
{
    private const string PresentationContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    private const string OfficeDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    internal const string SlideRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    public PptxDocument Read(OoxPackage package, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPart presentationPart = FindPresentationPart();
        using Stream stream = presentationPart.OpenRead();
        XDocument document = SafeXml.Load(stream, cancellationToken);

        // Strict OOXML (ISO 29500) parts read as blank under transitional queries;
        // fail visibly once per document instead of converting silently empty (O02).
        if (diagnosticSink is not null && OoxNamespaces.HasStrictOoxmlRoot(document))
        {
            diagnosticSink(new OoxPdfDiagnostic(
                "OOXML_STRICT_DIALECT",
                OoxPdfSeverity.Warning,
                "Strict OOXML content was detected; only the transitional dialect is supported and content may be missing.",
                presentationPart.Name,
                SlideIndex: null,
                PageIndex: null,
                Feature: "strict-dialect",
                Fallback: "Ignored"));
        }

        if (diagnosticSink is not null && OoxMarkupCompatibility.HasUnrecognizedMustUnderstand(document))
        {
            diagnosticSink(new OoxPdfDiagnostic(
                "OOXML_MUST_UNDERSTAND",
                OoxPdfSeverity.Warning,
                "Content marked must-understand uses unsupported namespaces and was ignored.",
                presentationPart.Name,
                SlideIndex: null,
                PageIndex: null,
                Feature: "must-understand",
                Fallback: "Ignored"));
        }

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
                string target = relationship.ResolvedTarget ?? throw new InvalidDataException("Slide relationship has no resolved target part.");
                if (!IsHiddenSlide(package, target, cancellationToken))
                {
                    slides.Add(new PptxSlide(target, slides.Count));
                }
            }
        }

        if (slides.Count == 0)
        {
            foreach (OoxRelationship relationship in relationships.Values)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string target = relationship.ResolvedTarget ?? throw new InvalidDataException("Slide relationship has no resolved target part.");
                if (!IsHiddenSlide(package, target, cancellationToken))
                {
                    slides.Add(new PptxSlide(target, slides.Count));
                }
            }
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

    // Hidden slides (unprefixed p:sld flag, verified against PowerPoint-saved
    // output) are excluded to match PowerPoint PDF export (S01). Missing parts
    // stay included; the scene renders them as blank pages.
    private static bool IsHiddenSlide(OoxPackage package, string partName, CancellationToken cancellationToken)
    {
        OoxPart? part = package.GetPart(partName);
        if (part is null)
        {
            return false;
        }

        using Stream stream = part.OpenRead();
        XDocument slideXml = SafeXml.Load(stream, cancellationToken);
        return OoxBoolean.IsOff((string?)slideXml.Root?.Attribute("show"));
    }


}
