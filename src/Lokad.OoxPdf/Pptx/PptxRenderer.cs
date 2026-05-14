using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;
using System.Globalization;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed class PptxRenderer
{
    private static readonly XNamespace PresentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public IReadOnlyList<PdfPage> RenderBlankPages(PptxDocument document)
    {
        return document.Slides
            .Select(_ => new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints))
            .ToArray();
    }

    public IReadOnlyList<PdfPage> RenderPages(PptxDocument document, OoxPackage package)
    {
        var pages = new List<PdfPage>(document.Slides.Count);
        foreach (PptxSlide slide in document.Slides)
        {
            OoxPart? slidePart = package.GetPart(slide.PartName);
            if (slidePart is null)
            {
                pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints));
                continue;
            }

            using Stream stream = slidePart.OpenRead();
            XDocument slideXml = SafeXml.Load(stream);
            var graphics = new PdfGraphicsBuilder();

            RenderBackground(slideXml, document, graphics);
            RenderShapes(slideXml, document, graphics);
            pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints, graphics.ToString()));
        }

        return pages;
    }

    private static void RenderBackground(XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics)
    {
        XElement? background = slideXml.Root?
            .Element(PresentationNamespace + "cSld")?
            .Element(PresentationNamespace + "bg")?
            .Element(PresentationNamespace + "bgPr");
        if (TryReadSolidColor(background, out RgbColor color))
        {
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(0, 0, document.SlideWidthPoints, document.SlideHeightPoints);
        }
    }

    private static void RenderShapes(XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics)
    {
        IEnumerable<XElement> shapes = slideXml
            .Descendants(PresentationNamespace + "spTree")
            .Elements(PresentationNamespace + "sp");

        foreach (XElement shape in shapes)
        {
            XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
            if (shapeProperties is null)
            {
                continue;
            }

            ShapeBounds? bounds = ReadBounds(shapeProperties);
            if (bounds is null)
            {
                continue;
            }

            string preset = (string?)shapeProperties
                .Element(DrawingNamespace + "prstGeom")
                ?.Attribute("prst") ?? "rect";

            double x = OoxUnits.EmuToPoints(bounds.Value.X);
            double yTop = OoxUnits.EmuToPoints(bounds.Value.Y);
            double width = OoxUnits.EmuToPoints(bounds.Value.Width);
            double height = OoxUnits.EmuToPoints(bounds.Value.Height);
            double y = document.SlideHeightPoints - yTop - height;
            bool transformed = bounds.Value.RotationDegrees != 0d || bounds.Value.FlipHorizontal || bounds.Value.FlipVertical;

            bool hasFill = TryReadSolidColor(shapeProperties, out RgbColor fill);
            bool hasStroke = TryReadLine(shapeProperties, out RgbColor stroke, out double lineWidth);

            if (transformed)
            {
                graphics.SaveState();
                ApplyShapeTransform(graphics, x, y, width, height, bounds.Value);
            }

            if (preset == "line")
            {
                if (hasStroke)
                {
                    graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
                    graphics.SetLineWidth(lineWidth);
                    graphics.StrokeLine(x, document.SlideHeightPoints - yTop, x + width, document.SlideHeightPoints - yTop - height);
                }

                if (transformed)
                {
                    graphics.RestoreState();
                }

                continue;
            }

            if (hasFill)
            {
                graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
                if (preset == "ellipse")
                {
                    graphics.FillEllipse(x, y, width, height);
                }
                else
                {
                    graphics.FillRectangle(x, y, width, height);
                }
            }

            if (hasStroke)
            {
                graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
                graphics.SetLineWidth(lineWidth);
                if (preset == "ellipse")
                {
                    graphics.StrokeEllipse(x, y, width, height);
                }
                else
                {
                    graphics.StrokeRectangle(x, y, width, height);
                }
            }

            if (transformed)
            {
                graphics.RestoreState();
            }
        }
    }

    private static ShapeBounds? ReadBounds(XElement shapeProperties)
    {
        XElement? transform = shapeProperties.Element(DrawingNamespace + "xfrm");
        XElement? offset = transform?.Element(DrawingNamespace + "off");
        XElement? extents = transform?.Element(DrawingNamespace + "ext");
        if (transform is null || offset is null || extents is null)
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

    private static void ApplyShapeTransform(PdfGraphicsBuilder graphics, double x, double y, double width, double height, ShapeBounds bounds)
    {
        double radians = -bounds.RotationDegrees * Math.PI / 180d;
        double sx = bounds.FlipHorizontal ? -1d : 1d;
        double sy = bounds.FlipVertical ? -1d : 1d;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        double centerX = x + width / 2d;
        double centerY = y + height / 2d;

        double a = cos * sx;
        double b = sin * sx;
        double c = -sin * sy;
        double d = cos * sy;
        double e = centerX - a * centerX - c * centerY;
        double f = centerY - b * centerX - d * centerY;
        graphics.Transform(a, b, c, d, e, f);
    }

    private static bool TryReadSolidColor(XElement? element, out RgbColor color)
    {
        string? hex = (string?)element?
            .Element(DrawingNamespace + "solidFill")
            ?.Element(DrawingNamespace + "srgbClr")
            ?.Attribute("val");
        if (hex is null || hex.Length != 6)
        {
            color = default;
            return false;
        }

        color = new RgbColor(
            byte.Parse(hex.AsSpan(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            byte.Parse(hex.AsSpan(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
            byte.Parse(hex.AsSpan(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        return true;
    }

    private static bool TryReadLine(XElement shapeProperties, out RgbColor color, out double lineWidth)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColor(line, out color);
    }

    private static long ParseLongAttribute(XElement element, string name)
    {
        string value = (string?)element.Attribute(name)
            ?? throw new InvalidDataException($"Missing required PPTX shape attribute '{name}'.");
        return long.Parse(value, CultureInfo.InvariantCulture);
    }

    private static bool ParseBoolAttribute(XElement element, string name)
    {
        string? value = (string?)element.Attribute(name);
        return value is not null && (value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase));
    }

    private readonly record struct ShapeBounds(
        long X,
        long Y,
        long Width,
        long Height,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical);

    private readonly record struct RgbColor(byte Red, byte Green, byte Blue);
}
