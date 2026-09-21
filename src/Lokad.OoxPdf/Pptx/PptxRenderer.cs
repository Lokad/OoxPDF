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
    private readonly PresentationFontResolver fontResolver;

    public PptxRenderer(IFontResolver? fontResolver)
    {
        this.fontResolver = new PresentationFontResolver(fontResolver);
    }

    public IReadOnlyList<PdfPage> RenderBlankPages(PptxDocument document, CancellationToken cancellationToken)
    {
        var pages = new PdfPage[document.Slides.Count];
        for (int i = 0; i < pages.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            pages[i] = new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints);
        }

        return pages;
    }

    public IReadOnlyList<PdfPage> RenderPages(PptxDocument document, OoxPackage package, Action<OoxPdfDiagnostic>? diagnosticSink, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var pages = new List<PdfPage>(document.Slides.Count);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, cancellationToken);
        PptxTheme theme = scene.Theme;
        var imageCache = new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase);
        // PLAN W05: one bounded workbook model per embedded part, shared by every chart
        // frame of this conversion (like the image cache above). Each model is already
        // capped (workbook totals, range unions); the dictionary itself is bounded by
        // the package entry count and dies with this conversion.
        var chartWorkbookCache = new Dictionary<string, object?>(StringComparer.Ordinal);
        // PLAN W02: per-slide layout memoization shared by font collection (below) and
        // node painting. Owned per slide, dies with the conversion.
        var textSpanMemo = new Dictionary<PptxTextSpanMemoKey, object?>(PptxTextSpanMemoKeyComparer.Instance);
        var tableFrameMemo = new Dictionary<PptxTableFrameMemoKey, object?>(PptxTableFrameMemoKeyComparer.Instance);
        var warnedMustUnderstandParts = new HashSet<string>(StringComparer.Ordinal);
        for (int slideIndex = 0; slideIndex < document.Slides.Count; slideIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PptxSlide slide = document.Slides[slideIndex];
            PptxSceneSlide sceneSlide = scene.Slides[slideIndex];
            XDocument slideXml = sceneSlide.SlideXml;
            if (slideXml.Root is null)
            {
                pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints));
                continue;
            }

            EmitUnsupportedFeatureDiagnostics(sceneSlide, slideXml, slide.PartName, slideIndex + 1, diagnosticSink, warnedMustUnderstandParts);
            var graphics = new PdfGraphicsBuilder();
            PptxRenderContext context = CreateRenderContext(document, theme, slide, slideXml, sceneSlide, fontResolver, imageCache, diagnosticSink, cancellationToken, chartWorkbookCache, textSpanMemo, tableFrameMemo);

            bool masterBackgroundPainted = RenderBackground(context, context.SceneSlide.MasterBackground, graphics, defaultWhenMissing: false);
            bool layoutBackgroundPainted = RenderBackground(context, context.SceneSlide.LayoutBackground, graphics, defaultWhenMissing: false);
            bool slideBackgroundPainted = RenderBackground(context, context.SceneSlide.SlideBackground, graphics, defaultWhenMissing: false);
            if (!masterBackgroundPainted && !layoutBackgroundPainted && !slideBackgroundPainted)
            {
                RenderBackground(context, context.SceneSlide.SlideBackground, graphics, defaultWhenMissing: true);
            }
            cancellationToken.ThrowIfCancellationRequested();
            var orderedImages = new List<PdfImageResource>();
            var linkAnnotations = new List<PdfLinkAnnotation>();
            var reportedHyperlinkIds = new HashSet<string>(StringComparer.Ordinal);
            var orderedChartFonts = new List<PdfFontResource>();
            int imageIndex = 1;
            IReadOnlyList<PptxPositionedTextSpan> shapeTextSpans = ReadSceneShapeTextSpans(context, includeMasterNodes: context.SceneSlide.ShowMasterShapes);
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = ReadSceneTableTextSpans(context, includeMasterNodes: context.SceneSlide.ShowMasterShapes);
            cancellationToken.ThrowIfCancellationRequested();
            RenderedFonts renderedFonts = CreateRenderedFonts(shapeTextSpans.Concat(tableTextSpans).Select(span => span.Run).ToArray(), fontResolver, "F", cancellationToken, diagnosticSink);
            // Hidden master shapes stay unpainted when the slide opts out (S01).
            if (context.SceneSlide.ShowMasterShapes)
            {
                RenderOrderedSceneNodes(context.SceneSlide.MasterNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, linkAnnotations, reportedHyperlinkIds, context.MasterPartName, context.MasterColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false, cancellationToken: cancellationToken);
            }
            RenderOrderedSceneNodes(context.SceneSlide.LayoutNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, linkAnnotations, reportedHyperlinkIds, context.LayoutPartName, context.LayoutColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false, cancellationToken: cancellationToken);
            RenderOrderedSceneNodes(context.SceneSlide.SlideNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, linkAnnotations, reportedHyperlinkIds, context.SlidePartName, context.SlideColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: true, cancellationToken: cancellationToken);
            // Q02: drop orphan image/chart-font registrations left by rewound nodes so
            // failed nodes cannot accumulate uncharged serialized resources. Surviving
            // references are name-based, so pruning unreferenced names keeps them valid.
            string content = graphics.ToString();
            List<PdfImageResource> pageImages = PruneUnreferencedImages(content, orderedImages);
            List<PdfFontResource> pageChartFonts = PruneUnreferencedChartFonts(content, orderedChartFonts);
            pages.Add(new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, content, renderedFonts.Resources.Concat(pageChartFonts).ToArray(), pageImages, graphics.ExtGStates.ToArray(), graphics.Shadings.ToArray(), graphics.Patterns.ToArray(), linkAnnotations));

        }

        return pages;
    }

    private static PptxRenderContext? TryLoadRenderContext(
        PptxDocument document,
        OoxPackage package,
        int slideIndex,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        PptxScene? sharedScene = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (slideIndex < 0 || slideIndex >= document.Slides.Count)
        {
            return null;
        }

        PptxSlide slide = document.Slides[slideIndex];
        // PLAN Q06: inspection sessions share one scene per conversion instead of
        // rebuilding it per slide per entry point. Rendering still builds its own.
        PptxScene scene = sharedScene ?? new PptxSceneBuilder().Build(document, package, cancellationToken);
        PptxSceneSlide sceneSlide = scene.Slides[slideIndex];
        XDocument slideXml = sceneSlide.SlideXml;
        if (slideXml.Root is null)
        {
            return null;
        }

        return CreateRenderContext(document, scene.Theme, slide, slideXml, sceneSlide, new PresentationFontResolver(null), imageCache, diagnosticSink, cancellationToken);
    }

    private static PptxRenderContext CreateRenderContext(
        PptxDocument document,
        PptxTheme theme,
        PptxSlide slide,
        XDocument slideXml,
        PptxSceneSlide sceneSlide,
        PresentationFontResolver fontResolver,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken,
        Dictionary<string, object?>? chartWorkbookCache = null,
        Dictionary<PptxTextSpanMemoKey, object?>? textSpanMemo = null,
        Dictionary<PptxTableFrameMemoKey, object?>? tableFrameMemo = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        PptxRenderSource slideSource = new(
            PptxRenderSourceKind.Slide,
            sceneSlide.PartName,
            slideXml,
            sceneSlide.SlideRelationships,
            sceneSlide.SlideColorMap);
        return new PptxRenderContext(document, theme, slide, sceneSlide, slideSource, BuildInheritedSources(), fontResolver, imageCache, diagnosticSink, cancellationToken, chartWorkbookCache ?? new Dictionary<string, object?>(StringComparer.Ordinal), textSpanMemo ?? new Dictionary<PptxTextSpanMemoKey, object?>(PptxTextSpanMemoKeyComparer.Instance), tableFrameMemo ?? new Dictionary<PptxTableFrameMemoKey, object?>(PptxTableFrameMemoKeyComparer.Instance));

        IReadOnlyList<PptxRenderSource> BuildInheritedSources()
        {
            return (sceneSlide.MasterXml, sceneSlide.LayoutXml) switch
            {
                ({ } master, { } layout) => [BuildMasterSource(sceneSlide, master), BuildLayoutSource(sceneSlide, layout)],
                ({ } master, null) => [BuildMasterSource(sceneSlide, master)],
                (null, { } layout) => [BuildLayoutSource(sceneSlide, layout)],
                _ => []
            };
        }
    }

    private static PptxRenderSource BuildMasterSource(PptxSceneSlide sceneSlide, XDocument xml)
    {
        return new PptxRenderSource(
            PptxRenderSourceKind.Master,
            sceneSlide.MasterPartName,
            xml,
            sceneSlide.MasterRelationships,
            sceneSlide.MasterColorMap);
    }

    private static PptxRenderSource BuildLayoutSource(PptxSceneSlide sceneSlide, XDocument xml)
    {
        return new PptxRenderSource(
            PptxRenderSourceKind.Layout,
            sceneSlide.LayoutPartName,
            xml,
            sceneSlide.LayoutRelationships,
            sceneSlide.LayoutColorMap);
    }

    private static ShapeBounds? ReadBounds(XElement shapeProperties)
    {
        XElement? transform = shapeProperties.Element(DrawingNamespace + "xfrm");
        return transform is null ? null : ReadBoundsFromTransform(transform);
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

    private static TextAlignment ReadAlignment(XElement paragraph, XElement? defaultParagraphProperties)
    {
        return ParseAlignment(ReadAlignmentValue(paragraph, defaultParagraphProperties));
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
