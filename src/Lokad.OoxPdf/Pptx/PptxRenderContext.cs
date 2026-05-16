using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxRenderContext(
    OoxPackage Package,
    PptxDocument Document,
    PptxTheme Theme,
    PptxSlide Slide,
    XDocument SlideXml,
    IReadOnlyList<XDocument> InheritedXml,
    IReadOnlyDictionary<string, OoxRelationship> SlideRelationships,
    Action<OoxPdfDiagnostic>? DiagnosticSink)
{
    public int SlideNumber => Slide.Index + 1;
}
