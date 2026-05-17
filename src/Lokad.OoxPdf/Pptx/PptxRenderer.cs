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
    private const string SlideLayoutRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
    private const string SlideMasterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
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
        PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);
        var imageCache = new Dictionary<string, PdfImageXObject?>(StringComparer.OrdinalIgnoreCase);
        for (int slideIndex = 0; slideIndex < document.Slides.Count; slideIndex++)
        {
            PptxSlide slide = document.Slides[slideIndex];
            OoxPart? slidePart = package.GetPart(slide.PartName);
            if (slidePart is null)
            {
                pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints));
                continue;
            }

            using Stream stream = slidePart.OpenRead();
            XDocument slideXml = SafeXml.Load(stream);
            EmitUnsupportedFeatureDiagnostics(slideXml, slide.PartName, slideIndex + 1, diagnosticSink);
            var graphics = new PdfGraphicsBuilder();
            PptxRenderContext context = CreateRenderContext(package, document, theme, slide, slideXml, imageCache, diagnosticSink);

            foreach (XDocument inherited in context.InheritedXml)
            {
                RenderBackground(context, inherited, graphics);
                RenderShapes(context, inherited, graphics, renderPlaceholders: false);
            }

            RenderBackground(context, context.SlideXml, graphics);
            if (CanRenderSlideInOrder(context.SlideXml))
            {
                var orderedImages = new List<PdfImageResource>();
                int imageIndex = 1;
                IReadOnlyList<PptxPositionedTextSpan> inheritedTextSpans = ReadInheritedTextSpans(context);
                IReadOnlyList<PptxPositionedTextSpan> slideTextSpans = ReadSlideTextSpans(context);
                IReadOnlyList<TextRun> slideTableTextRuns = RenderTables(context, context.SlideXml, new PdfGraphicsBuilder());
                RenderedFonts renderedFonts = CreateRenderedFonts(inheritedTextSpans.Concat(slideTextSpans).Select(span => span.Run).Concat(slideTableTextRuns).ToArray());
                DrawTextSpansWithFonts(inheritedTextSpans, graphics, renderedFonts.Fonts);
                foreach (XElement shapeTree in context.SlideXml.Descendants(PresentationNamespace + "spTree"))
                {
                    RenderOrderedShapeTextContainer(shapeTree, context, graphics, renderedFonts.Fonts, orderedImages, ref imageIndex, GroupTransform.Identity, renderPlaceholders: true);
                }

                pages.Add(new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, graphics.ToString(), renderedFonts.Resources, orderedImages, graphics.ExtGStates.ToArray()));
                continue;
            }

            IReadOnlyList<PdfImageResource> images = RenderPictures(context, graphics);
            RenderShapes(context, context.SlideXml, graphics, renderPlaceholders: true);
            IReadOnlyList<TextRun> tableTextRuns = context.InheritedXml
                .Append(context.SlideXml)
                .SelectMany(xml => RenderTables(context, xml, graphics))
                .ToArray();
            RenderCharts(context, graphics);
            IReadOnlyList<PptxPositionedTextSpan> textSpans = ReadInheritedTextSpans(context)
                .Concat(ReadSlideTextSpans(context))
                .ToArray();
            IReadOnlyList<PdfFontResource> fonts = RenderPositionedTextSpans(textSpans, tableTextRuns, graphics);
            pages.Add(new PdfPage(context.Document.SlideWidthPoints, context.Document.SlideHeightPoints, graphics.ToString(), fonts, images, graphics.ExtGStates.ToArray()));
        }

        return pages;
    }

    private static PptxRenderContext? TryLoadRenderContext(
        PptxDocument document,
        OoxPackage package,
        PptxTheme theme,
        int slideIndex,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        if (slideIndex < 0 || slideIndex >= document.Slides.Count)
        {
            return null;
        }

        PptxSlide slide = document.Slides[slideIndex];
        OoxPart? slidePart = package.GetPart(slide.PartName);
        if (slidePart is null)
        {
            return null;
        }

        using Stream stream = slidePart.OpenRead();
        XDocument slideXml = SafeXml.Load(stream);
        return CreateRenderContext(package, document, theme, slide, slideXml, imageCache, diagnosticSink);
    }

    private static PptxRenderContext CreateRenderContext(
        OoxPackage package,
        PptxDocument document,
        PptxTheme theme,
        PptxSlide slide,
        XDocument slideXml,
        Dictionary<string, PdfImageXObject?> imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink)
    {
        PptxSlideInheritance inheritance = LoadSlideInheritance(package, slide.PartName);
        IReadOnlyDictionary<string, OoxRelationship> slideRelationships = package.GetRelationships(slide.PartName)
            .Where(r => !r.IsExternal && r.ResolvedTarget is not null)
            .ToDictionary(r => r.Id, StringComparer.Ordinal);
        return new PptxRenderContext(package, document, theme, slide, slideXml, inheritance, inheritance.Sources, slideRelationships, imageCache, diagnosticSink);
    }

    private static IReadOnlyList<XDocument> LoadInheritedSlideXml(OoxPackage package, string slidePartName)
    {
        return LoadSlideInheritance(package, slidePartName).Sources;
    }

    private static PptxSlideInheritance LoadSlideInheritance(OoxPackage package, string slidePartName)
    {
        OoxPart? layoutPart = GetRelatedPart(package, slidePartName, SlideLayoutRelationshipType);
        OoxPart? masterPart = layoutPart is null ? null : GetRelatedPart(package, layoutPart.Name, SlideMasterRelationshipType);

        XDocument? masterXml = null;
        if (masterPart is not null)
        {
            using Stream masterStream = masterPart.OpenRead();
            masterXml = SafeXml.Load(masterStream);
        }

        XDocument? layoutXml = null;
        if (layoutPart is not null)
        {
            using Stream layoutStream = layoutPart.OpenRead();
            layoutXml = SafeXml.Load(layoutStream);
        }

        return new PptxSlideInheritance(masterXml, layoutXml);
    }

    private static OoxPart? GetRelatedPart(OoxPackage package, string sourcePartName, string relationshipType)
    {
        OoxRelationship? relationship = package.GetRelationships(sourcePartName)
            .FirstOrDefault(r => !r.IsExternal && r.Type == relationshipType && r.ResolvedTarget is not null);
        return relationship?.ResolvedTarget is null ? null : package.GetPart(relationship.ResolvedTarget);
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
        string? value = (string?)element.Attribute(name);
        return value is not null && (value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase));
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
