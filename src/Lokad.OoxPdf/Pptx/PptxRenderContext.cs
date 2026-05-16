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
    PptxSlideInheritance Inheritance,
    IReadOnlyList<XDocument> InheritedXml,
    IReadOnlyDictionary<string, OoxRelationship> SlideRelationships,
    Dictionary<string, PdfImageXObject?> ImageCache,
    Action<OoxPdfDiagnostic>? DiagnosticSink)
{
    public int SlideNumber => Slide.Index + 1;
}

internal sealed record PptxSlideInheritance(
    XDocument? MasterXml,
    XDocument? LayoutXml)
{
    public IReadOnlyList<XDocument> Sources => (MasterXml, LayoutXml) switch
    {
        ({ } master, { } layout) => [master, layout],
        ({ } master, null) => [master],
        (null, { } layout) => [layout],
        _ => []
    };
}
