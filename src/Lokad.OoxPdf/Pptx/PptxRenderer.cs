using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Fonts;
using System.Globalization;
using System.Text;
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
            PptxTheme theme = PptxTheme.Load(package, document.PresentationPartName);

            RenderBackground(slideXml, document, graphics, theme);
            RenderShapes(slideXml, document, graphics, theme);
            IReadOnlyList<TextRun> textRuns = ReadTextRuns(slideXml, document, theme);
            IReadOnlyList<PdfFontResource> fonts = RenderTextRuns(textRuns, graphics);
            pages.Add(new PdfPage(document.SlideWidthPoints, document.SlideHeightPoints, graphics.ToString(), fonts));
        }

        return pages;
    }

    private static void RenderBackground(XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme)
    {
        XElement? background = slideXml.Root?
            .Element(PresentationNamespace + "cSld")?
            .Element(PresentationNamespace + "bg")?
            .Element(PresentationNamespace + "bgPr");
        if (TryReadSolidColor(background, theme, out RgbColor color))
        {
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(0, 0, document.SlideWidthPoints, document.SlideHeightPoints);
        }
    }

    private static void RenderShapes(XDocument slideXml, PptxDocument document, PdfGraphicsBuilder graphics, PptxTheme theme)
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

            bool hasFill = TryReadSolidColor(shapeProperties, theme, out RgbColor fill);
            bool hasStroke = TryReadLine(shapeProperties, theme, out RgbColor stroke, out double lineWidth);

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

    private static bool TryReadSolidColor(XElement? element, PptxTheme theme, out RgbColor color)
    {
        XElement? solidFill = element?.Element(DrawingNamespace + "solidFill");
        string? hex = (string?)solidFill?.Element(DrawingNamespace + "srgbClr")?.Attribute("val");
        if (RgbColor.TryParse(hex, out color))
        {
            return true;
        }

        string? schemeColor = (string?)solidFill?.Element(DrawingNamespace + "schemeClr")?.Attribute("val");
        return schemeColor is not null && theme.TryResolveColor(schemeColor, out color);
    }

    private static bool TryReadLine(XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        lineWidth = line?.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColor(line, theme, out color);
    }

    private static IReadOnlyList<TextRun> ReadTextRuns(XDocument slideXml, PptxDocument document, PptxTheme theme)
    {
        var runs = new List<TextRun>();
        foreach (XElement shape in slideXml.Descendants(PresentationNamespace + "sp"))
        {
            XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
            XElement? textBody = shape.Element(PresentationNamespace + "txBody");
            ShapeBounds? bounds = shapeProperties is null ? null : ReadBounds(shapeProperties);
            if (bounds is null || textBody is null)
            {
                continue;
            }

            double x = OoxUnits.EmuToPoints(bounds.Value.X);
            double yTop = OoxUnits.EmuToPoints(bounds.Value.Y);
            double width = OoxUnits.EmuToPoints(bounds.Value.Width);
            double height = OoxUnits.EmuToPoints(bounds.Value.Height);
            double cursorY = document.SlideHeightPoints - yTop - 18d;

            foreach (XElement paragraph in textBody.Elements(DrawingNamespace + "p"))
            {
                TextAlignment alignment = ReadAlignment(paragraph);
                double cursorX = x + 4d;
                double maxFontSize = 18d;
                foreach (XElement run in paragraph.Elements(DrawingNamespace + "r"))
                {
                    string text = (string?)run.Element(DrawingNamespace + "t") ?? string.Empty;
                    if (text.Length == 0)
                    {
                        continue;
                    }

                    XElement? runProperties = run.Element(DrawingNamespace + "rPr");
                    double fontSize = runProperties?.Attribute("sz") is { } size
                        ? int.Parse(size.Value, CultureInfo.InvariantCulture) / 100d
                        : 18d;
                    maxFontSize = Math.Max(maxFontSize, fontSize);
                    RgbColor color = TryReadSolidColor(runProperties, theme, out RgbColor runColor)
                        ? runColor
                        : new RgbColor(0, 0, 0);
                    string? typeface = theme.ResolveTypeface((string?)runProperties?
                        .Element(DrawingNamespace + "latin")
                        ?.Attribute("typeface"));
                    bool bold = ParseOptionalBoolAttribute(runProperties, "b");
                    bool italic = ParseOptionalBoolAttribute(runProperties, "i");
                    bool underline = ((string?)runProperties?.Attribute("u")) is { } underlineValue
                        && !underlineValue.Equals("none", StringComparison.OrdinalIgnoreCase);
                    runs.Add(new TextRun(text, cursorX, cursorY, width, height, fontSize, color, bold, italic, underline, alignment, typeface));
                    cursorX += text.Length * fontSize * 0.55d;
                }

                cursorY -= maxFontSize * 1.2d;
            }
        }

        return runs;
    }

    private static IReadOnlyList<PdfFontResource> RenderTextRuns(IReadOnlyList<TextRun> textRuns, PdfGraphicsBuilder graphics)
    {
        if (textRuns.Count == 0)
        {
            return [];
        }

        string familyName = textRuns.Select(r => r.FontFamily).FirstOrDefault(f => !string.IsNullOrWhiteSpace(f)) ?? "Arial";
        FontResolution resolution = new WindowsFontResolver().Resolve(new FontRequest(familyName));
        if (resolution.FontFilePath is null || !File.Exists(resolution.FontFilePath))
        {
            return [];
        }

        OpenTypeFont font = OpenTypeFont.Load(resolution.FontFilePath);
        PdfEmbeddedFont embedded = PdfEmbeddedFont.Create(font, textRuns.SelectMany(r => r.Text.EnumerateRunes().Select(rune => rune.Value)));
        var resource = new PdfFontResource("F1", embedded);

        foreach (TextRun run in textRuns)
        {
            DrawWrappedRun(graphics, embedded, run);
        }

        return [resource];
    }

    private static void DrawWrappedRun(PdfGraphicsBuilder graphics, PdfEmbeddedFont embedded, TextRun run)
    {
        double cursorY = run.Y;
        double lineHeight = run.FontSize * 1.2d;
        foreach (string line in WrapWords(run.Text, run.Width, run.FontSize, embedded))
        {
            if (cursorY < run.Y - run.Height)
            {
                break;
            }

            string glyphHex = embedded.EncodeGlyphHex(line);
            if (glyphHex.Length != 0)
            {
                double lineWidth = embedded.MeasureTextPoints(line, run.FontSize);
                double x = run.Alignment switch
                {
                    TextAlignment.Center => run.X + Math.Max(0, run.Width - lineWidth) / 2d,
                    TextAlignment.Right => run.X + Math.Max(0, run.Width - lineWidth),
                    _ => run.X
                };

                graphics.DrawGlyphText("F1", run.FontSize, x, cursorY, run.Color.Red, run.Color.Green, run.Color.Blue, glyphHex, run.Italic);
                if (run.Bold)
                {
                    graphics.DrawGlyphText("F1", run.FontSize, x + 0.35d, cursorY, run.Color.Red, run.Color.Green, run.Color.Blue, glyphHex, run.Italic);
                }

                if (run.Underline)
                {
                    graphics.SetStrokeRgb(run.Color.Red, run.Color.Green, run.Color.Blue);
                    graphics.SetLineWidth(Math.Max(0.5d, run.FontSize / 18d));
                    graphics.StrokeLine(x, cursorY - run.FontSize * 0.12d, x + lineWidth, cursorY - run.FontSize * 0.12d);
                }
            }

            cursorY -= lineHeight;
        }
    }

    private static IEnumerable<string> WrapWords(string text, double maxWidth, double fontSize, PdfEmbeddedFont embedded)
    {
        string[] words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0)
        {
            yield break;
        }

        var line = new StringBuilder();
        foreach (string word in words)
        {
            string candidate = line.Length == 0 ? word : line + " " + word;
            if (line.Length > 0 && embedded.MeasureTextPoints(candidate, fontSize) > maxWidth)
            {
                yield return line.ToString();
                line.Clear();
                line.Append(word);
            }
            else
            {
                line.Clear();
                line.Append(candidate);
            }
        }

        if (line.Length > 0)
        {
            yield return line.ToString();
        }
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

    private static bool ParseOptionalBoolAttribute(XElement? element, string name)
    {
        return element is not null && ParseBoolAttribute(element, name);
    }

    private static TextAlignment ReadAlignment(XElement paragraph)
    {
        string? value = (string?)paragraph.Element(DrawingNamespace + "pPr")?.Attribute("algn");
        return value switch
        {
            "ctr" => TextAlignment.Center,
            "r" => TextAlignment.Right,
            _ => TextAlignment.Left
        };
    }

    private readonly record struct ShapeBounds(
        long X,
        long Y,
        long Width,
        long Height,
        double RotationDegrees,
        bool FlipHorizontal,
        bool FlipVertical);

    private readonly record struct TextRun(
        string Text,
        double X,
        double Y,
        double Width,
        double Height,
        double FontSize,
        RgbColor Color,
        bool Bold,
        bool Italic,
        bool Underline,
        TextAlignment Alignment,
        string? FontFamily);

    private enum TextAlignment
    {
        Left,
        Center,
        Right
    }
}
