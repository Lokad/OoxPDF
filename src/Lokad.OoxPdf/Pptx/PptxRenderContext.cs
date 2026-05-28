using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxRenderContext(
    PptxDocument Document,
    PptxTheme Theme,
    PptxSlide Slide,
    XDocument SlideXml,
    PptxSceneSlide SceneSlide,
    IReadOnlyList<XDocument> InheritedXml,
    PresentationFontResolver FontResolver,
    Dictionary<string, PdfImageXObject?> ImageCache,
    Action<OoxPdfDiagnostic>? DiagnosticSink)
{
    public int SlideNumber => Slide.Index + 1;
}
