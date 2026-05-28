using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
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

    public string SlidePartName => SceneSlide.PartName;

    public string? MasterPartName => SceneSlide.MasterPartName;

    public string? LayoutPartName => SceneSlide.LayoutPartName;

    public IReadOnlyDictionary<string, OoxRelationship> SlideRelationships => SceneSlide.SlideRelationships;

    public IReadOnlyDictionary<string, OoxRelationship> MasterRelationships => SceneSlide.MasterRelationships;

    public IReadOnlyDictionary<string, OoxRelationship> LayoutRelationships => SceneSlide.LayoutRelationships;

    public PptxColorMap SlideColorMap => SceneSlide.SlideColorMap;

    public PptxColorMap MasterColorMap => SceneSlide.MasterColorMap;

    public PptxColorMap LayoutColorMap => SceneSlide.LayoutColorMap;
}
