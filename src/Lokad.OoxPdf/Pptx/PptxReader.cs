using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;

namespace Lokad.OoxPdf.Pptx;

internal sealed class PptxReader
{
    private const string PresentationContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    internal const string SlideRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

    public PptxDocument Read(OoxPackage package, CancellationToken cancellationToken, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        OoxPart presentationPart = ReadPresentationPart(package, cancellationToken);
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
                diagnosticSink(OoxMarkupCompatibility.CreateMustUnderstandWarning(presentationPart.Name));
            }

            XElement? size = document.Root?.Element(PresentationNamespace + "sldSz");
            double width = size is null ? 720d : OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(size, "cx", "PPTX"));
            double height = size is null ? 540d : OoxUnits.EmuToPoints(OoxXml.ParseRequiredLong(size, "cy", "PPTX"));
            long firstSlideNumber = OoxXml.ReadOptionalLong(document.Root, "firstSlideNum", 1);
            if (firstSlideNumber is < 0 or > uint.MaxValue) firstSlideNumber = 1;

            var slides = new List<PptxSlide>();
            foreach (var (target, sourceIndex) in ReadSlideSources(package, presentationPart, document, cancellationToken))
            {
                if (!IsHiddenSlide(package, target, cancellationToken))
                {
                    slides.Add(new PptxSlide(target, slides.Count) { SourceIndex = sourceIndex });
                }
            }

            return result = new PptxDocument(presentationPart.Name, slides, width, height)
            {
                FirstSlideNumber = firstSlideNumber,
                DefaultTextStyle = document.Root?.Element(PresentationNamespace + "defaultTextStyle")
            };
        }
        finally
        {
            if (result is null) modelStorage?.Dispose();
            else if (modelStorage is not null) OoxMemory.Resize(modelStorage, modelStorage.Bytes / 2);
        }
    }

    internal static OoxPart ReadPresentationPart(OoxPackage package, CancellationToken cancellationToken) =>
        package.FindMainPart(PresentationContentType, cancellationToken)
        ?? throw new InvalidDataException("PPTX package does not contain a presentation part.");

    internal static IEnumerable<(string PartName, int SourceIndex)> ReadSlideSources(
        OoxPackage package, OoxPart presentationPart, XDocument document, CancellationToken cancellationToken)
    {
        var relationships = package.GetRelationshipDictionary(presentationPart.Name, cancellationToken);
        var slideList = document.Root?.Element(PresentationNamespace + "sldIdLst");
        int sourceIndex = 0;
        if (slideList is not null)
        {
            // Every authored entry owns an index, including unresolved entries.
            foreach (XElement slideId in slideList.Elements(PresentationNamespace + "sldId"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                string? relationshipId = (string?)slideId.Attribute(RelationshipsNamespace + "id");
                if (relationshipId is not null && relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) &&
                    SlideTarget(relationship) is { } target)
                {
                    yield return (target, sourceIndex);
                }
                sourceIndex++;
            }
            yield break;
        }

        // With no slide list, only eligible relationships own a source position.
        foreach (OoxRelationship relationship in relationships.Values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (SlideTarget(relationship) is { } target)
            {
                yield return (target, sourceIndex++);
            }
        }

        static string? SlideTarget(OoxRelationship relationship) =>
            !relationship.IsExternal && relationship.Type == SlideRelationshipType ? relationship.ResolvedTarget : null;
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
