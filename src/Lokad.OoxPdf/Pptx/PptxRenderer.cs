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
            PptxRenderContext context = CreateRenderContext(document, theme, slide, slideXml, sceneSlide, fontResolver, imageCache, diagnosticSink, cancellationToken);

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
            IReadOnlyList<PptxPositionedTextSpan> shapeTextSpans = ReadSceneShapeTextSpans(context);
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = ReadSceneTableTextSpans(context);
            cancellationToken.ThrowIfCancellationRequested();
            RenderedFonts renderedFonts = CreateRenderedFonts(shapeTextSpans.Concat(tableTextSpans).Select(span => span.Run).ToArray(), fontResolver, "F", cancellationToken, diagnosticSink);
            // Hidden master shapes stay unpainted when the slide opts out (S01).
            if (context.SceneSlide.ShowMasterShapes)
            {
                RenderOrderedSceneNodes(context.SceneSlide.MasterNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, linkAnnotations, reportedHyperlinkIds, context.MasterPartName, context.MasterColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false, cancellationToken: cancellationToken);
            }
            RenderOrderedSceneNodes(context.SceneSlide.LayoutNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, linkAnnotations, reportedHyperlinkIds, context.LayoutPartName, context.LayoutColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false, cancellationToken: cancellationToken);
            RenderOrderedSceneNodes(context.SceneSlide.SlideNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, linkAnnotations, reportedHyperlinkIds, context.SlidePartName, context.SlideColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: true, cancellationToken: cancellationToken);
            pages.Add(new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, graphics.ToString(), renderedFonts.Resources.Concat(orderedChartFonts).ToArray(), orderedImages, graphics.ExtGStates.ToArray(), graphics.Shadings.ToArray(), graphics.Patterns.ToArray(), linkAnnotations));

        }

        return pages;
    }

    private static PptxRenderContext? TryLoadRenderContext(
        PptxDocument document,
        OoxPackage package,
        int slideIndex,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (slideIndex < 0 || slideIndex >= document.Slides.Count)
        {
            return null;
        }

        PptxSlide slide = document.Slides[slideIndex];
        PptxScene scene = new PptxSceneBuilder().Build(document, package, cancellationToken);
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
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        PptxRenderSource slideSource = new(
            PptxRenderSourceKind.Slide,
            sceneSlide.PartName,
            slideXml,
            sceneSlide.SlideRelationships,
            sceneSlide.SlideColorMap);
        return new PptxRenderContext(document, theme, slide, sceneSlide, slideSource, BuildInheritedSources(), fontResolver, imageCache, diagnosticSink, cancellationToken);

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
