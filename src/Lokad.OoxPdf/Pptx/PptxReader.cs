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
        XDocument document = package.LoadXml(presentationPart, cancellationToken);
        IMemoryReservation? modelStorage = OoxMemory.ReserveModel(document, cancellationToken);
        PptxDocument? result = null;
        try
        {

            // Strict OOXML (ISO 29500) parts read as blank under transitional queries;
            // fail visibly once per document instead of converting silently empty (O02).
            if (diagnosticSink is not null && OoxNamespaces.HasStrictOoxmlOrigin(document))
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
                    "A must-understand declaration requires unsupported or unresolved namespaces; conversion continues with partial compatibility handling.",
                    presentationPart.Name,
                    SlideIndex: null,
                    PageIndex: null,
                    Feature: "must-understand",
                    Fallback: "Partial compatibility"));
            }

            XElement? size = document.Root?.Element(PresentationNamespace + "sldSz");
            double width = size is null ? 720d : OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(size, "cx", "PPTX"));
            double height = size is null ? 540d : OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(size, "cy", "PPTX"));
            long firstSlideNumber = OoxXml.ReadOptionalLong(document.Root, "firstSlideNum", 1);
            if (firstSlideNumber is < 0 or > uint.MaxValue) firstSlideNumber = 1;

            IReadOnlyDictionary<string, OoxRelationship> relationships = package.GetRelationships(presentationPart.Name, cancellationToken)
                .Where(r => !r.IsExternal && r.Type == SlideRelationshipType && r.ResolvedTarget is not null)
                .ToDictionary(r => r.Id, StringComparer.Ordinal);

            var slides = new List<PptxSlide>();
            XElement? slideList = document.Root?.Element(PresentationNamespace + "sldIdLst");
            IEnumerable<XElement> slideIds = slideList?.Elements(PresentationNamespace + "sldId") ?? [];

            int sourceIndex = -1;
            foreach (XElement slideId in slideIds)
            {
                cancellationToken.ThrowIfCancellationRequested();
                sourceIndex++;
                string? relationshipId = (string?)slideId.Attribute(RelationshipsNamespace + "id");
                if (relationshipId is not null && relationships.TryGetValue(relationshipId, out OoxRelationship? relationship))
                {
                    string target = relationship.ResolvedTarget ?? throw new InvalidDataException("Slide relationship has no resolved target part.");
                    if (!IsHiddenSlide(package, target, cancellationToken))
                    {
                        slides.Add(new PptxSlide(target, slides.Count) { SourceIndex = sourceIndex });
                    }
                }
            }

            // Recover missing ordering metadata, never an explicitly empty or fully
            // hidden list: relationships alone do not make an unlisted slide printable.
            if (slideList is null)
            {
                foreach (OoxRelationship relationship in relationships.Values)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    sourceIndex++;
                    string target = relationship.ResolvedTarget ?? throw new InvalidDataException("Slide relationship has no resolved target part.");
                    if (!IsHiddenSlide(package, target, cancellationToken))
                    {
                        slides.Add(new PptxSlide(target, slides.Count) { SourceIndex = sourceIndex });
                    }
                }
            }

            return result = new PptxDocument(presentationPart.Name, slides, width, height) { FirstSlideNumber = firstSlideNumber };
        }
        finally
        {
            if (result is null) modelStorage?.Dispose();
            else if (modelStorage is not null) OoxMemory.Resize(modelStorage, modelStorage.Bytes / 2);
        }

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

        XDocument slideXml = package.LoadXml(part, cancellationToken);
        return OoxBoolean.IsOff((string?)slideXml.Root?.Attribute("show"));
    }


}
