using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Diagnostics;
using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private readonly PresentationFontResolver fontResolver;

    public PptxRenderer(IFontResolver? fontResolver = null)
    {
        this.fontResolver = new PresentationFontResolver(fontResolver);
    }

    public IReadOnlyList<PdfPage> RenderBlankPages(PptxDocument document, CancellationToken cancellationToken = default)
    {
        var pages = new PdfPage[document.Slides.Count];
        for (int i = 0; i < pages.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            pages[i] = new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints);
        }

        return pages;
    }

    public IReadOnlyList<PdfPage> RenderPages(PptxDocument document, OoxPackage package, Action<OoxPdfDiagnostic>? diagnosticSink = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var pages = new List<PdfPage>(document.Slides.Count);
        PptxScene scene = new PptxSceneBuilder().Build(document, package, cancellationToken);
        PptxTheme theme = scene.Theme;
        var imageCache = new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase);
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

            EmitUnsupportedFeatureDiagnostics(sceneSlide, slideXml, slide.PartName, slideIndex + 1, diagnosticSink);
            var graphics = new PdfGraphicsBuilder();
            PptxRenderContext context = CreateRenderContext(document, theme, slide, slideXml, sceneSlide, fontResolver, imageCache, diagnosticSink, cancellationToken);

            RenderBackground(context, context.SceneSlide.MasterBackground, graphics, defaultWhenMissing: false);
            RenderBackground(context, context.SceneSlide.LayoutBackground, graphics, defaultWhenMissing: false);
            RenderBackground(context, context.SceneSlide.SlideBackground, graphics, defaultWhenMissing: true);
            cancellationToken.ThrowIfCancellationRequested();
            var orderedImages = new List<PdfImageResource>();
            var orderedChartFonts = new List<PdfFontResource>();
            int imageIndex = 1;
            IReadOnlyList<PptxPositionedTextSpan> shapeTextSpans = ReadSceneShapeTextSpans(context);
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = ReadSceneTableTextSpans(context);
            cancellationToken.ThrowIfCancellationRequested();
            RenderedFonts renderedFonts = CreateRenderedFonts(shapeTextSpans.Concat(tableTextSpans).Select(span => span.Run).ToArray(), fontResolver, cancellationToken: cancellationToken);
            RenderOrderedSceneNodes(context.SceneSlide.MasterNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, context.MasterPartName, context.MasterColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false, cancellationToken: cancellationToken);
            RenderOrderedSceneNodes(context.SceneSlide.LayoutNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, context.LayoutPartName, context.LayoutColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false, cancellationToken: cancellationToken);
            RenderOrderedSceneNodes(context.SceneSlide.SlideNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, context.SlidePartName, context.SlideColorMap, ref imageIndex, GroupTransform.Identity, renderPlaceholders: true, cancellationToken: cancellationToken);

            pages.Add(new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, graphics.ToString(), renderedFonts.Resources.Concat(orderedChartFonts).ToArray(), orderedImages, graphics.ExtGStates.ToArray(), graphics.Shadings.ToArray(), graphics.Patterns.ToArray()));
        }

        return pages;
    }

    private static PptxRenderContext? TryLoadRenderContext(
        PptxDocument document,
        OoxPackage package,
        int slideIndex,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        CancellationToken cancellationToken = default)
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

        return CreateRenderContext(document, scene.Theme, slide, slideXml, sceneSlide, new PresentationFontResolver(), imageCache, diagnosticSink, cancellationToken);
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
        return new PptxRenderContext(document, theme, slide, sceneSlide, slideSource, BuildInheritedSources(sceneSlide), fontResolver, imageCache, diagnosticSink, cancellationToken);
    }

    private static IReadOnlyList<PptxRenderSource> BuildInheritedSources(PptxSceneSlide sceneSlide)
    {
        return (sceneSlide.MasterXml, sceneSlide.LayoutXml) switch
        {
            ({ } master, { } layout) => [BuildMasterSource(sceneSlide, master), BuildLayoutSource(sceneSlide, layout)],
            ({ } master, null) => [BuildMasterSource(sceneSlide, master)],
            (null, { } layout) => [BuildLayoutSource(sceneSlide, layout)],
            _ => []
        };
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
        bool flipHorizontal = ParseBoolAttribute(transform, "flipH");
        bool flipVertical = ParseBoolAttribute(transform, "flipV");

        return new ShapeBounds(
            ParseLongAttribute(offset, "x"),
            ParseLongAttribute(offset, "y"),
            ParseLongAttribute(extents, "cx"),
            ParseLongAttribute(extents, "cy"),
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

    private static long ParseLongAttribute(XElement element, string name)
    {
        string value = (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required PPTX shape attribute '{name}'.");
        return long.Parse(value, CultureInfo.InvariantCulture);
    }

    private static long ParseOptionalLongAttribute(XElement element, string name, long defaultValue)
    {
        return element.Attribute(name) is { } value
            ? long.Parse(value.Value, CultureInfo.InvariantCulture)
            : defaultValue;
    }

    private static int ParseOptionalIntAttribute(XElement? element, string name, int defaultValue)
    {
        return element?.Attribute(name) is { } value
            ? int.Parse(value.Value, CultureInfo.InvariantCulture)
            : defaultValue;
    }

    private static bool ParseBoolAttribute(XElement element, string name)
    {
        return ParseBoolAttribute(element, name, defaultValue: false);
    }

    private static bool ParseBoolAttribute(XElement element, string name, bool defaultValue)
    {
        return OoxBoolean.ParseAttribute(element, name, defaultValue);
    }

    private static bool ParseOptionalBoolAttribute(XElement? element, string name)
    {
        return element is not null && ParseBoolAttribute(element, name);
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
