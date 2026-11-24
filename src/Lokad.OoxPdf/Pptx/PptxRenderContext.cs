using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed record PptxRenderContext(
    PptxDocument Document,
    PptxSlide Slide,
    PptxSceneSlide SceneSlide,
    PresentationFontResolver FontResolver,
    Dictionary<string, PdfImageXObject?>? ImageCache,
    Action<OoxPdfDiagnostic>? DiagnosticSink,
    CancellationToken CancellationToken,
    // R16: caches ride as typed dictionaries with explicit null-negative contracts
    // (a hit carrying null already resolved as absent; only a miss recomputes).
    Dictionary<string, PptxRenderer.ChartWorkbookData?>? WorkbookCache = null,
    Dictionary<PptxTextSpanMemoKey, IReadOnlyList<PptxRenderer.PptxPositionedTextSpan>>? TextSpanMemo = null,
    Dictionary<PptxTableFrameMemoKey, PptxRenderer.TableFrameLayout?>? TableFrameMemo = null)
{
    // Borrow the page owner's existing grant. Context copies never own/refund it.
    public IMemoryReservation? PageConstruction { get; init; }
    public long PageConstructionBaseBytes { get; init; }
    public long PageAxisConstructionSites { get; init; }
    public IReadOnlyDictionary<FontRequest, PptxRenderer.SvgTextFont?>? SvgTextFonts { get; init; }

    public PptxTheme Theme => SceneSlide.Theme;

    // Hidden master paint still contributes placeholder defaults before the layout.
    public IReadOnlyList<XDocument> InheritedXml { get; } = (SceneSlide.MasterXml, SceneSlide.LayoutXml) switch
    {
        ({ } master, { } layout) => [master, layout],
        ({ } master, null) => [master],
        (null, { } layout) => [layout],
        _ => []
    };

    public int SlideNumber => Slide.Index + 1;

    // Displayed fields follow the authored sequence, independently of the
    // compacted page index used by diagnostics and navigation.
    public long DisplayedSlideNumber => Document.FirstSlideNumber + (Slide.SourceIndex ?? Slide.Index);
}
