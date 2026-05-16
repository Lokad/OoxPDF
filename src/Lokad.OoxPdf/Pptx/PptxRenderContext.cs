using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxRenderContext(
    OoxPackage Package,
    PptxDocument Document,
    PptxTheme Theme,
    PptxSlide Slide,
    XDocument SlideXml,
    IReadOnlyList<XDocument> InheritedXml,
    IReadOnlyDictionary<string, OoxRelationship> SlideRelationships,
    Dictionary<string, PdfImageXObject?> ImageCache,
    Action<OoxPdfDiagnostic>? DiagnosticSink)
{
    public int SlideNumber => Slide.Index + 1;
}
