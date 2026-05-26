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
    private static readonly XNamespace OoxPdfInternalNamespace = "urn:lokad:ooxpdf:internal";
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public IReadOnlyList<PdfPage> RenderBlankPages(PptxDocument document)
    {
        return document.Slides
            .Select(_ => new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints))
            .ToArray();
    }

    public IReadOnlyList<PdfPage> RenderPages(PptxDocument document, OoxPackage package, Action<OoxPdfDiagnostic>? diagnosticSink = null)
    {
        var pages = new List<PdfPage>(document.Slides.Count);
        PptxScene scene = new PptxSceneBuilder().Build(document, package);
        PptxTheme theme = scene.Theme;
        var imageCache = new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase);
        for (int slideIndex = 0; slideIndex < document.Slides.Count; slideIndex++)
        {
            PptxSlide slide = document.Slides[slideIndex];
            PptxSceneSlide sceneSlide = scene.Slides[slideIndex];
            XDocument slideXml = sceneSlide.SlideXml;
            if (slideXml.Root is null)
            {
                pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints));
                continue;
            }

            EmitUnsupportedFeatureDiagnostics(slideXml, slide.PartName, slideIndex + 1, diagnosticSink);
            var graphics = new PdfGraphicsBuilder();
            PptxRenderContext context = CreateRenderContext(package, document, theme, slide, slideXml, sceneSlide, imageCache, diagnosticSink);

            RenderBackground(context, context.SceneSlide.MasterBackground, graphics, defaultWhenMissing: false);
            RenderBackground(context, context.SceneSlide.LayoutBackground, graphics, defaultWhenMissing: false);
            RenderBackground(context, context.SceneSlide.SlideBackground, graphics, defaultWhenMissing: true);
            var orderedImages = new List<PdfImageResource>();
            var orderedChartFonts = new List<PdfFontResource>();
            int imageIndex = 1;
            IReadOnlyList<PptxPositionedTextSpan> shapeTextSpans = ReadSceneShapeTextSpans(context);
            IReadOnlyList<PptxPositionedTextSpan> tableTextSpans = ReadSceneTableTextSpans(context);
            RenderedFonts renderedFonts = CreateRenderedFonts(shapeTextSpans.Concat(tableTextSpans).Select(span => span.Run).ToArray());
            RenderOrderedSceneNodes(context.SceneSlide.MasterNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, context.SceneSlide.MasterRelationships, context.SceneSlide.MasterPartName, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false);
            RenderOrderedSceneNodes(context.SceneSlide.LayoutNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, context.SceneSlide.LayoutRelationships, context.SceneSlide.LayoutPartName, ref imageIndex, GroupTransform.Identity, renderPlaceholders: false);
            RenderOrderedSceneNodes(context.SceneSlide.SlideNodes, context, graphics, renderedFonts.Fonts, orderedImages, orderedChartFonts, context.SceneSlide.SlideRelationships, context.Slide.PartName, ref imageIndex, GroupTransform.Identity, renderPlaceholders: true);

            pages.Add(new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, graphics.ToString(), renderedFonts.Resources.Concat(orderedChartFonts).ToArray(), orderedImages, graphics.ExtGStates.ToArray(), graphics.Shadings.ToArray()));
        }

        return pages;
    }

    private static PptxRenderContext? TryLoadRenderContext(
        PptxDocument document,
        OoxPackage package,
        int slideIndex,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        if (slideIndex < 0 || slideIndex >= document.Slides.Count)
        {
            return null;
        }

        PptxSlide slide = document.Slides[slideIndex];
        PptxScene scene = new PptxSceneBuilder().Build(document, package);
        PptxSceneSlide sceneSlide = scene.Slides[slideIndex];
        XDocument slideXml = sceneSlide.SlideXml;
        if (slideXml.Root is null)
        {
            return null;
        }

        return CreateRenderContext(package, document, scene.Theme, slide, slideXml, sceneSlide, imageCache, diagnosticSink);
    }

    private static PptxRenderContext CreateRenderContext(
        OoxPackage package,
        PptxDocument document,
        PptxTheme theme,
        PptxSlide slide,
        XDocument slideXml,
        PptxSceneSlide sceneSlide,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        PptxSlideInheritance inheritance = new(sceneSlide.MasterXml, sceneSlide.LayoutXml);
        IReadOnlyDictionary<string, OoxRelationship> slideRelationships = sceneSlide.SlideRelationships;
        return new PptxRenderContext(package, document, theme, slide, slideXml, sceneSlide, inheritance, inheritance.Sources, slideRelationships, imageCache, diagnosticSink);
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
        string? value = (string?)(paragraph.Element(DrawingNamespace + "pPr")?.Attribute("algn") ??
            defaultParagraphProperties?.Attribute("algn"));
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
