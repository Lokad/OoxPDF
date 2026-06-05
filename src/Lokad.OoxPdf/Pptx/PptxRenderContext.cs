using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal enum PptxRenderSourceKind
{
    Master,
    Layout,
    Slide
}

internal sealed record PptxRenderSource(
    PptxRenderSourceKind Kind,
    string? PartName,
    XDocument Xml,
    IReadOnlyDictionary<string, OoxRelationship> Relationships,
    PptxColorMap ColorMap);

internal sealed record PptxRenderContext(
    PptxDocument Document,
    PptxTheme Theme,
    PptxSlide Slide,
    PptxSceneSlide SceneSlide,
    PptxRenderSource SlideSource,
    IReadOnlyList<PptxRenderSource> InheritedSources,
    PresentationFontResolver FontResolver,
    Dictionary<string, PdfImageXObject?> ImageCache,
    Action<OoxPdfDiagnostic>? DiagnosticSink,
    CancellationToken CancellationToken)
{
    public IReadOnlyList<XDocument> InheritedXml { get; } = InheritedSources.Select(source => source.Xml).ToArray();

    public int SlideNumber => Slide.Index + 1;

    public XDocument SlideXml => SlideSource.Xml;

    public string SlidePartName => SlideSource.PartName ?? SceneSlide.PartName;

    public string? MasterPartName => MasterSource?.PartName ?? SceneSlide.MasterPartName;

    public string? LayoutPartName => LayoutSource?.PartName ?? SceneSlide.LayoutPartName;

    public PptxRenderSource? MasterSource => InheritedSources.FirstOrDefault(source => source.Kind == PptxRenderSourceKind.Master);

    public PptxRenderSource? LayoutSource => InheritedSources.FirstOrDefault(source => source.Kind == PptxRenderSourceKind.Layout);

    public IReadOnlyDictionary<string, OoxRelationship> SlideRelationships => SlideSource.Relationships;

    public IReadOnlyDictionary<string, OoxRelationship> MasterRelationships => MasterSource?.Relationships ?? SceneSlide.MasterRelationships;

    public IReadOnlyDictionary<string, OoxRelationship> LayoutRelationships => LayoutSource?.Relationships ?? SceneSlide.LayoutRelationships;

    public PptxColorMap SlideColorMap => SlideSource.ColorMap;

    public PptxColorMap MasterColorMap => MasterSource?.ColorMap ?? SceneSlide.MasterColorMap;

    public PptxColorMap LayoutColorMap => LayoutSource?.ColorMap ?? SceneSlide.LayoutColorMap;
}
