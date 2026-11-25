using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Diagnostics;
using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private readonly IFontResolver? primaryFonts;

    public PptxRenderer(IFontResolver? fontResolver)
    {
        primaryFonts = fontResolver;
    }

    // R06.3: yields pages progressively so the writer can spill content and retain
    // only descriptors; callers must drain inside the conversion budget scope.
    public async IAsyncEnumerable<PdfPage> RenderPagesAsync(PptxDocument document, OoxPackage package, Action<OoxPdfDiagnostic>? diagnosticSink, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var fontResolver = new PresentationFontResolver(primaryFonts ??
            await WindowsFontResolver.CreateAsync(cancellationToken).ConfigureAwait(false), cancellationToken);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, cancellationToken);
        var imageCache = new Dictionary<string, RasterPicture>(StringComparer.OrdinalIgnoreCase);
        // one bounded workbook model per embedded part, shared by every chart
        // frame of this conversion (like the image cache above). Each model is already
        // capped (workbook totals, range unions); the dictionary itself is bounded by
        // the package entry count and dies with this conversion.
        var chartWorkbookCache = new Dictionary<string, ChartWorkbookData?>(StringComparer.Ordinal);
        // R09: per-slide layout memoization shared by font collection (below) and
        // node painting. Owned inside the slide iteration so previous-slide glyph/layout
        // graphs become eligible for collection each slide; cross-slide Node references
        // never hit (fresh master/layout instances per slide), so per-slide ownership
        // preserves compute-once within a slide without retaining the whole deck.
        var warnedMustUnderstandParts = new HashSet<string>(StringComparer.Ordinal);
        for (int slideIndex = 0; slideIndex < document.Slides.Count; slideIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // R06.2: admit the upcoming page before it materializes; later slides
            // (and their fonts/images) never produce past the trip.
            OoxConversionBudget.Current?.ChargePdfPages(1);
            PptxSlide slide = document.Slides[slideIndex];
            PptxSceneSlide sceneSlide = scene.Slides[slideIndex];
            XDocument slideXml = sceneSlide.SlideXml;
            if (slideXml.Root is null)
            {
                yield return new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints);
                continue;
            }

            sceneSlide = PreparePictureSources(sceneSlide, imageCache, diagnosticSink, slideIndex + 1, cancellationToken);
            using IMemoryReservation? construction = ReservePageConstruction(sceneSlide, cancellationToken, out long axisConstructionSites, out long svgSourceBytes);
            cancellationToken.ThrowIfCancellationRequested();
            var graphics = new PdfGraphicsBuilder(construction, svgSourceBytes);
            var textSpanMemo = new Dictionary<PptxTextSpanMemoKey, IReadOnlyList<PptxPositionedTextSpan>>(PptxTextSpanMemoKeyComparer.Instance);
            var tableFrameMemo = new Dictionary<PptxTableFrameMemoKey, TableFrameLayout?>(PptxTableFrameMemoKeyComparer.Instance);
            cancellationToken.ThrowIfCancellationRequested();
            var context = new PptxRenderContext(document, slide, sceneSlide, fontResolver, imageCache, diagnosticSink,
                cancellationToken, chartWorkbookCache, textSpanMemo, tableFrameMemo);
            IReadOnlyList<PptxPositionedTextSpan> shapeTextSpans = await ReadSceneShapeTextSpansAsync(context).ConfigureAwait(false);
            EmitUnsupportedFeatureDiagnostics(sceneSlide, slideXml, slide.PartName, slideIndex + 1, diagnosticSink, warnedMustUnderstandParts, textSpanMemo);
            cancellationToken.ThrowIfCancellationRequested();
            context = context with
            {
                PageConstruction = construction,
                PageConstructionBaseBytes = construction?.Bytes ?? 0L,
                PageAxisConstructionSites = axisConstructionSites,
                SvgTextFonts = await PrepareSvgTextFontsAsync(EnumerateSvgPictureSources(sceneSlide), fontResolver,
                    fontResolver.GetOrLoadOpenTypeFontAsync, diagnosticSink, cancellationToken).ConfigureAwait(false)
            };

            RenderBackground(context, graphics);
            cancellationToken.ThrowIfCancellationRequested();
            var orderedImages = new List<PdfImageResource>();
            var linkAnnotations = new List<PdfLinkAnnotation>();
            var reportedHyperlinkIds = new HashSet<string>(StringComparer.Ordinal);
            var orderedChartFonts = new List<PdfFontResource>();
            int imageIndex = 1;
            var tableTextSpans = new List<PptxPositionedTextSpan>();
            await foreach (TableFrameLayout table in ReadSceneTableLayoutsAsync(context).ConfigureAwait(false))
                tableTextSpans.AddRange(table.TextSpans);
            cancellationToken.ThrowIfCancellationRequested();
            // Prepare the faces used by measured glyphs, rather than reporting a
            // primary-face gap for text that already has an allowed fallback.
            PptxPositionedTextSpan[] slideTextSpans = shapeTextSpans.Concat(tableTextSpans).ToArray();
            RenderedFonts renderedFonts = await CreateRenderedFontsAsync(slideTextSpans,
                fontResolver, "F", cancellationToken, diagnosticSink, includeFallbackFaces: false).ConfigureAwait(false);
            Dictionary<FontRequest, RenderedFont> slideFonts = new(renderedFonts.Fonts, FontRequestKeyComparer.OrdinalIgnoreCaseFamily);
            var slideResources = new List<PdfFontResource>(renderedFonts.Resources);
            await AddSplitFallbackFacesAsync(slideFonts, slideTextSpans, fontResolver, diagnosticSink, cancellationToken, slideResources, "F").ConfigureAwait(false);
            renderedFonts = new RenderedFonts(slideFonts, slideResources);

            foreach (var source in EnumerateScenePaintSources(sceneSlide))
            {
                imageIndex = await RenderOrderedSceneNodesAsync(source.Nodes, context, graphics, renderedFonts.Fonts, orderedImages,
                    orderedChartFonts, linkAnnotations, reportedHyperlinkIds, source.PartName, source.ColorMap, imageIndex,
                    GroupTransform.Identity, source.RenderPlaceholders, cancellationToken).ConfigureAwait(false);
            }
            cancellationToken.ThrowIfCancellationRequested();
            string content = graphics.ToString();
            IReadOnlySet<string> usedFonts = graphics.UsedFontResourceNames;
            var pageFonts = new List<PdfFontResource>();
            foreach (PdfFontResource font in renderedFonts.Resources.Concat(orderedChartFonts))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (usedFonts.Contains(font.ResourceName))
                {
                    pageFonts.Add(font);
                }
            }
            IReadOnlyList<PdfFallbackFontResource> fallbackFonts = PdfFallbackFont.GetUsedResources(usedFonts);
            // RV11: admit the emitted content bytes before yielding, so a small
            // content budget rejects before the writer encodes or stores the page.
            OoxConversionBudget.Current?.ChargePdfContentBytes(content.Length);
            yield return new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, content,
                pageFonts, orderedImages, graphics.ExtGStates.ToArray(), graphics.Shadings.ToArray(), graphics.Patterns.ToArray(),
                linkAnnotations, fallbackFonts, graphics.Groups.ToArray(), graphics.ShadingPatterns.ToArray());
        }
    }

    internal static PptxRenderContext? TryLoadTextInspectionContext(
        PptxDocument document,
        OoxPackage package,
        int slideIndex,
        PptxScene? sharedScene = null,
        PresentationFontResolver? inspectionFonts = null)
    {
        if (slideIndex < 0 || slideIndex >= document.Slides.Count)
        {
            return null;
        }

        PptxSlide slide = document.Slides[slideIndex];
        // inspection sessions share one scene per conversion instead of
        // rebuilding it per slide per entry point. Rendering still builds its own.
        PptxScene scene = sharedScene ?? new PptxSceneBuilder().Build(document, package, CancellationToken.None);
        PptxSceneSlide sceneSlide = scene.Slides[slideIndex];
        if (sceneSlide.SlideXml.Root is null)
        {
            return null;
        }

        return new PptxRenderContext(document, slide, sceneSlide, inspectionFonts ?? new PresentationFontResolver(null),
            ImageCache: null, DiagnosticSink: null, CancellationToken: CancellationToken.None,
            TextSpanMemo: new(PptxTextSpanMemoKeyComparer.Instance),
            TableFrameMemo: new(PptxTableFrameMemoKeyComparer.Instance));
    }

    private static ShapeBounds? ReadShapeBounds(XElement shape, IReadOnlyList<XElement> inheritedPlaceholders,
        out bool usesInheritedShapeBounds)
    {
        usesInheritedShapeBounds = false;
        if (ReadOwnerBounds(shape) is { } direct) return direct;
        // Inherited sources are ordered master then layout. An empty nearer
        // placeholder does not erase a usable frame supplied by a lower owner.
        for (int index = inheritedPlaceholders.Count - 1; index >= 0; index--)
        {
            if (ReadOwnerBounds(inheritedPlaceholders[index]) is { } inherited)
            {
                usesInheritedShapeBounds = true;
                return inherited;
            }
        }
        return null;

        static ShapeBounds? ReadOwnerBounds(XElement owner) =>
            owner.Element(PresentationNamespace + "spPr")?.Element(DrawingNamespace + "xfrm") is { } transform
                ? ReadBoundsFromTransform(transform) : null;
    }

    private static ShapeBounds? ReadBoundsFromTransform(XElement transform)
    {
        XElement? offset = transform.Element(DrawingNamespace + "off");
        XElement? extents = transform.Element(DrawingNamespace + "ext");
        if (offset is null || extents is null)
        {
            return null;
        }

        double rotationDegrees = transform.Attribute("rot") is { } rotationAttribute
            ? long.Parse(rotationAttribute.Value, CultureInfo.InvariantCulture) / 60000d
            : 0d;
        bool flipHorizontal = OoxXml.ParseOptionalBool(transform, "flipH");
        bool flipVertical = OoxXml.ParseOptionalBool(transform, "flipV");

        return new ShapeBounds(
            OoxXml.ParseRequiredLong(offset, "x", "PPTX shape"),
            OoxXml.ParseRequiredLong(offset, "y", "PPTX shape"),
            OoxXml.ParseRequiredLong(extents, "cx", "PPTX shape"),
            OoxXml.ParseRequiredLong(extents, "cy", "PPTX shape"),
            rotationDegrees,
            flipHorizontal,
            flipVertical);
    }

    private static bool IsPlaceholder(XElement shape)
    {
        return shape
            .Element(PresentationNamespace + "nvSpPr")
            ?.Element(PresentationNamespace + "nvPr")
            ?.Element(PresentationNamespace + "ph") is not null;
    }

    private static string? ReadAlignmentValue(XElement paragraph, XElement? defaultParagraphProperties)
    {
        return (string?)(paragraph.Element(DrawingNamespace + "pPr")?.Attribute("algn") ??
            defaultParagraphProperties?.Attribute("algn"));
    }

    private static TextAlignment ParseAlignment(string? value)
    {
        return value switch
        {
            "ctr" => TextAlignment.Center,
            "r" => TextAlignment.Right,
            "just" => TextAlignment.Justify,
            "dist" => TextAlignment.Distributed,
            "justLow" => TextAlignment.JustLow,
            "thaiDist" => TextAlignment.ThaiDistributed,
            _ => TextAlignment.Left
        };
    }

}
