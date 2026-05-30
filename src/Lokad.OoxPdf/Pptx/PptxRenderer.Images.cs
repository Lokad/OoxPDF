using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void RenderPicture(
        PptxSceneNode picture,
        PptxRenderContext context,
        PdfGraphicsBuilder graphics,
        GroupTransform transform,
        List<PdfImageResource> images,
        ref int index)
    {
        if (picture.Picture is null || picture.Bounds is null)
        {
            return;
        }

        RenderPicture(
            context.Document,
            graphics,
            context.DiagnosticSink,
            context.SlideNumber,
            transform,
            images,
            context.ImageCache,
            ref index,
            picture.Picture.TargetPartName,
            picture.Picture.Resource,
            ToShapeBounds(picture.Bounds),
            ToCropRect(picture.Picture.Crop),
            ToFillRect(picture.Picture.Fill),
            picture.Picture.Alpha,
            picture.Picture.Recolor,
            ToLineStyle(picture.Picture.Line),
            ToOuterShadow(picture.Picture.OuterShadow));
    }

    private static void RenderPicture(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        GroupTransform transform,
        List<PdfImageResource> images,
        Dictionary<string, PdfImageXObject?> imageCache,
        ref int index,
        string? targetPartName,
        PptxSceneImageResource? imageResource,
        ShapeBounds rawBounds,
        CropRect crop,
        FillRect fillRect,
        double alpha,
        PptxSceneImageRecolor recolor,
        LineStyle line,
        OuterShadow? outerShadow)
    {
        if (targetPartName is null)
        {
            return;
        }

        if (imageResource is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_PART",
                OoxPdfSeverity.Error,
                "Referenced image part was missing and the image was ignored.",
                targetPartName,
                SlideIndex: slideIndex,
                Feature: "image",
                Fallback: "Ignored"));
            return;
        }

        ShapeBounds transformedBounds = transform.Apply(rawBounds);
        double x = OoxUnits.EmuToPoints(transformedBounds.X);
        double yTop = OoxUnits.EmuToPoints(transformedBounds.Y);
        double width = OoxUnits.EmuToPoints(transformedBounds.Width);
        double height = OoxUnits.EmuToPoints(transformedBounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        bool hasTransform = Math.Abs(transformedBounds.RotationDegrees) > 0.001d || transformedBounds.FlipHorizontal || transformedBounds.FlipVertical;
        RenderPictureOuterShadow(document, graphics, transformedBounds, x, y, width, height, hasTransform, outerShadow, images, ref index);
        if (imageResource.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
        {
            RenderSvgPicture(graphics, document, transformedBounds, imageResource.Bytes, crop, fillRect);
            StrokePictureFrame(document, graphics, transformedBounds, x, y, width, height, line, hasTransform);
            return;
        }

        PdfImageXObject? image = null;
        if (!crop.IsEmpty)
        {
            image = GetOrCreateCroppedImage(imageResource, recolor, crop, imageCache, diagnosticSink, slideIndex);
            if (image is not null)
            {
                crop = default;
            }
        }

        image ??= GetOrCreateImage(imageResource, recolor, imageCache, diagnosticSink, slideIndex);
        if (image is null)
        {
            return;
        }

        string name = "Im" + index++;
        double imageX = x + fillRect.Left * width;
        double imageY = y + fillRect.Bottom * height;
        double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
        double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));
        bool transparent = alpha < 0.999d;
        if (hasTransform)
        {
            graphics.SaveState();
            ClipSlideBoundsEvenOdd(document, graphics);
            ApplyShapeTransform(graphics, x, y, width, height, transformedBounds);
        }

        if (transparent)
        {
            graphics.SaveState();
            graphics.SetAlpha(alpha, 1d);
        }

        graphics.SaveState();
        double clipX = imageX;
        double clipY = imageY;
        double clipWidth = imageWidth;
        double clipHeight = imageHeight;
        if (!hasTransform &&
            !TryIntersectWithSlideBounds(imageX, imageY, imageWidth, imageHeight, document, out clipX, out clipY, out clipWidth, out clipHeight))
        {
            graphics.RestoreState();
            if (transparent)
            {
                graphics.RestoreState();
            }

            return;
        }

        graphics.ClipRectangleEvenOdd(clipX, clipY, clipWidth, clipHeight);
        if (crop.IsEmpty)
        {
            graphics.DrawImage(name, imageX, imageY, imageWidth, imageHeight);
        }
        else
        {
            graphics.DrawImageCropped(name, imageX, imageY, imageWidth, imageHeight, crop.Left, crop.Top, crop.Right, crop.Bottom);
        }

        graphics.RestoreState();

        if (transparent)
        {
            graphics.RestoreState();
        }

        if (hasTransform)
        {
            graphics.RestoreState();
        }

        images.Add(new PdfImageResource(name, image));

        StrokePictureFrame(document, graphics, transformedBounds, x, y, width, height, line, hasTransform);
    }

    private static void RenderPictureOuterShadow(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        ShapeBounds bounds,
        double x,
        double y,
        double width,
        double height,
        bool transformed,
        OuterShadow? outerShadow,
        List<PdfImageResource> images,
        ref int imageIndex)
    {
        if (outerShadow is not { } shadow)
        {
            return;
        }

        graphics.SaveState();
        ClipSlideBoundsEvenOdd(document, graphics);
        if (transformed)
        {
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }

        if (shadow.BlurRadius > 0d)
        {
            DrawRasterOuterShadow(graphics, x, y, width, height, shadow, images, ref imageIndex);
        }
        else
        {
            DrawOuterShadow(graphics, "rect", x, y, width, height, shadow);
        }

        graphics.RestoreState();
    }

    private static void StrokePictureFrame(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        ShapeBounds bounds,
        double x,
        double y,
        double width,
        double height,
        LineStyle line,
        bool transformed)
    {
        if (!line.HasLine)
        {
            return;
        }

        if (transformed)
        {
            graphics.SaveState();
            ClipSlideBoundsEvenOdd(document, graphics);
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }
        else
        {
            ClipSlideBoundsEvenOdd(document, graphics);
        }

        bool transparentStroke = line.Alpha < 0.999d;
        if (transparentStroke)
        {
            graphics.SaveState();
            graphics.SetAlpha(1d, line.Alpha);
        }

        graphics.SetStrokeRgb(line.Color.Red, line.Color.Green, line.Color.Blue);
        graphics.SetLineWidth(line.Width);
        if (line.DashPattern is { Count: > 0 })
        {
            graphics.SetLineDash(line.DashPattern);
        }

        if (line.Cap is { } cap)
        {
            graphics.SetLineCap(cap);
        }

        if (line.Join is { } join)
        {
            graphics.SetLineJoin(join);
        }

        double outlineOutset = line.Width / 2d;
        graphics.StrokeRectangle(
            x - outlineOutset,
            y - outlineOutset,
            width + line.Width,
            height + line.Width);

        if (line.DashPattern is { Count: > 0 })
        {
            graphics.ClearLineDash();
        }

        if (line.Cap is not null)
        {
            graphics.SetLineCap(0);
        }

        if (line.Join is not null)
        {
            graphics.SetLineJoin(0);
        }

        if (transparentStroke)
        {
            graphics.RestoreState();
        }

        if (transformed)
        {
            graphics.RestoreState();
        }
    }

    private static ShapeBounds ToShapeBounds(PptxSceneBounds bounds)
    {
        return new ShapeBounds(
            bounds.XEmu,
            bounds.YEmu,
            bounds.WidthEmu,
            bounds.HeightEmu,
            bounds.RotationDegrees,
            bounds.FlipHorizontal,
            bounds.FlipVertical);
    }

    private static CropRect ToCropRect(PptxSceneRect rect)
    {
        return new CropRect(rect.Left, rect.Top, rect.Right, rect.Bottom);
    }

    private static FillRect ToFillRect(PptxSceneRect rect)
    {
        return new FillRect(rect.Left, rect.Top, rect.Right, rect.Bottom);
    }

    private static void RenderSvgPicture(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, byte[] bytes, CropRect crop, FillRect fillRect)
    {
        XDocument svg;
        using (var stream = new MemoryStream(bytes))
        {
            svg = SafeXml.Load(stream);
        }

        if (!TryReadSvgViewBox(svg.Root, out double minX, out double minY, out double viewWidth, out double viewHeight) ||
            viewWidth <= 0d ||
            viewHeight <= 0d)
        {
            return;
        }

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        double imageX = x + fillRect.Left * width;
        double imageY = y + fillRect.Bottom * height;
        double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
        double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));

        graphics.SaveState();
        if (!crop.IsEmpty || Math.Abs(bounds.RotationDegrees) > 0.001d || bounds.FlipHorizontal || bounds.FlipVertical)
        {
            graphics.ClipRectangle(imageX, imageY, imageWidth, imageHeight);
        }

        if (Math.Abs(bounds.RotationDegrees) > 0.001d || bounds.FlipHorizontal || bounds.FlipVertical)
        {
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }

        double scaleX = imageWidth / viewWidth;
        double scaleY = imageHeight / viewHeight;
        var gradients = ReadSvgGradients(svg);
        foreach (XElement path in svg.Descendants().Where(element => element.Name.LocalName == "path"))
        {
            string? data = (string?)path.Attribute("d");
            if (string.IsNullOrWhiteSpace(data) || !TryReadSvgFill(path, gradients, out SvgPaint paint))
            {
                continue;
            }

            if (paint.Gradient is { } gradient)
            {
                if (TryReadSvgPathBounds(data, out SvgPathBounds pathBounds))
                {
                    RenderSvgGradientPath(graphics, data, gradient, pathBounds, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY);
                }
            }
            else if (paint.Color is { } color)
            {
                graphics.SetFillRgb(color.Red, color.Green, color.Blue);
                if (TryAppendSvgPath(graphics, data, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY))
                {
                    graphics.FillCurrentPath();
                }
            }
        }

        graphics.RestoreState();
    }

    private static void RenderSvgGradientPath(
        PdfGraphicsBuilder graphics,
        string data,
        SvgGradient gradient,
        SvgPathBounds pathBounds,
        double minX,
        double minY,
        double imageX,
        double imageY,
        double imageHeight,
        double scaleX,
        double scaleY)
    {
        graphics.SaveState();
        if (!TryAppendSvgPath(graphics, data, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY))
        {
            graphics.RestoreState();
            return;
        }

        graphics.ClipCurrentPath();
        double pathWidth = Math.Max(0.001d, pathBounds.MaxX - pathBounds.MinX);
        int stripCount = Math.Clamp((int)Math.Ceiling(pathWidth / 2d), 16, 128);
        double stripSvgWidth = pathWidth / stripCount;
        double sampleY = pathBounds.CenterY;
        for (int strip = 0; strip < stripCount; strip++)
        {
            double stripMinX = pathBounds.MinX + strip * stripSvgWidth;
            double stripMaxX = strip == stripCount - 1 ? pathBounds.MaxX : stripMinX + stripSvgWidth;
            double sampleX = (stripMinX + stripMaxX) / 2d;
            RgbColor color = SampleSvgGradient(gradient, sampleX, sampleY);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(
                imageX + (stripMinX - minX) * scaleX,
                imageY,
                Math.Max(0.001d, (stripMaxX - stripMinX) * scaleX),
                imageHeight);
        }

        graphics.RestoreState();
    }

    private static bool TryReadSvgViewBox(XElement? root, out double minX, out double minY, out double width, out double height)
    {
        minX = 0d;
        minY = 0d;
        width = 0d;
        height = 0d;
        string? viewBox = (string?)root?.Attribute("viewBox");
        if (viewBox is null)
        {
            return false;
        }

        double[] values = SvgNumberRegex().Matches(viewBox)
            .Select(match => double.Parse(match.Value, CultureInfo.InvariantCulture))
            .ToArray();
        if (values.Length < 4)
        {
            return false;
        }

        minX = values[0];
        minY = values[1];
        width = values[2];
        height = values[3];
        return true;
    }

    private static IReadOnlyDictionary<string, SvgGradient> ReadSvgGradients(XDocument svg)
    {
        var gradients = new Dictionary<string, SvgGradient>(StringComparer.Ordinal);
        foreach (XElement gradient in svg.Descendants().Where(element => element.Name.LocalName == "linearGradient"))
        {
            string? id = (string?)gradient.Attribute("id");
            SvgGradientStop[] stops = gradient
                .Elements()
                .Where(element => element.Name.LocalName == "stop")
                .Select(ReadSvgGradientStop)
                .Where(stop => stop.Color is not null)
                .Select(stop => new SvgGradientStop(stop.Offset, stop.Color!.Value))
                .OrderBy(stop => stop.Offset)
                .ToArray();
            if (!string.IsNullOrWhiteSpace(id) && stops.Length > 0)
            {
                gradients[id] = new SvgGradient(
                    ReadSvgDoubleAttribute(gradient, "x1", 0d),
                    ReadSvgDoubleAttribute(gradient, "y1", 0d),
                    ReadSvgDoubleAttribute(gradient, "x2", 1d),
                    ReadSvgDoubleAttribute(gradient, "y2", 0d),
                    stops);
            }
        }

        return gradients;
    }

    private static (double Offset, RgbColor? Color) ReadSvgGradientStop(XElement stop)
    {
        return (ReadSvgOffset((string?)stop.Attribute("offset")), RgbColor.TryParse(((string?)stop.Attribute("stop-color"))?.TrimStart('#'), out RgbColor color) ? color : null);
    }

    private static double ReadSvgOffset(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0d;
        }

        string trimmed = value.Trim();
        if (trimmed.EndsWith("%", StringComparison.Ordinal))
        {
            return Math.Clamp(double.Parse(trimmed[..^1], CultureInfo.InvariantCulture) / 100d, 0d, 1d);
        }

        return Math.Clamp(double.Parse(trimmed, CultureInfo.InvariantCulture), 0d, 1d);
    }

    private static double ReadSvgDoubleAttribute(XElement element, string name, double fallback)
    {
        string? value = (string?)element.Attribute(name);
        return string.IsNullOrWhiteSpace(value)
            ? fallback
            : double.Parse(value, CultureInfo.InvariantCulture);
    }

    private static bool TryReadSvgFill(XElement path, IReadOnlyDictionary<string, SvgGradient> gradients, out SvgPaint paint)
    {
        string? fill = (string?)path.Attribute("fill");
        if (fill is null || fill.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            paint = default;
            return false;
        }

        Match gradient = Regex.Match(fill, @"url\(#(?<id>[^)]+)\)");
        if (gradient.Success && gradients.TryGetValue(gradient.Groups["id"].Value, out SvgGradient? svgGradient))
        {
            paint = new SvgPaint(null, svgGradient);
            return true;
        }

        if (RgbColor.TryParse(fill.TrimStart('#'), out RgbColor color))
        {
            paint = new SvgPaint(color, null);
            return true;
        }

        paint = default;
        return false;
    }

    private static RgbColor SampleSvgGradient(SvgGradient gradient, double x, double y)
    {
        double dx = gradient.X2 - gradient.X1;
        double dy = gradient.Y2 - gradient.Y1;
        double lengthSquared = dx * dx + dy * dy;
        double offset = lengthSquared <= PptxTextMetricRules.TextStateTolerance
            ? 0d
            : ((x - gradient.X1) * dx + (y - gradient.Y1) * dy) / lengthSquared;
        offset = Math.Clamp(offset, 0d, 1d);

        SvgGradientStop previous = gradient.Stops[0];
        foreach (SvgGradientStop next in gradient.Stops.Skip(1))
        {
            if (offset <= next.Offset)
            {
                double span = next.Offset - previous.Offset;
                double amount = span <= PptxTextMetricRules.TextStateTolerance ? 0d : (offset - previous.Offset) / span;
                return Interpolate(previous.Color, next.Color, Math.Clamp(amount, 0d, 1d));
            }

            previous = next;
        }

        return previous.Color;

        static RgbColor Interpolate(RgbColor left, RgbColor right, double amount)
        {
            return new RgbColor(
                ToByte(left.Red + (right.Red - left.Red) * amount),
                ToByte(left.Green + (right.Green - left.Green) * amount),
                ToByte(left.Blue + (right.Blue - left.Blue) * amount));
        }

        static byte ToByte(double value) => (byte)Math.Clamp((int)Math.Round(value), byte.MinValue, byte.MaxValue);
    }

    private static bool TryReadSvgPathBounds(string data, out SvgPathBounds bounds)
    {
        MatchCollection tokens = SvgPathTokenRegex().Matches(data);
        int index = 0;
        char command = '\0';
        double currentX = 0d;
        double currentY = 0d;
        double startX = 0d;
        double startY = 0d;
        double minX = double.PositiveInfinity;
        double minY = double.PositiveInfinity;
        double maxX = double.NegativeInfinity;
        double maxY = double.NegativeInfinity;

        void Include(double x, double y)
        {
            minX = Math.Min(minX, x);
            minY = Math.Min(minY, y);
            maxX = Math.Max(maxX, x);
            maxY = Math.Max(maxY, y);
        }

        while (index < tokens.Count)
        {
            string token = tokens[index].Value;
            if (token.Length == 1 && char.IsLetter(token[0]))
            {
                command = token[0];
                index++;
            }
            else if (command == '\0')
            {
                bounds = default;
                return false;
            }

            bool relative = char.IsLower(command);
            switch (char.ToUpperInvariant(command))
            {
                case 'M':
                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out currentX, out currentY))
                    {
                        bounds = default;
                        return false;
                    }

                    startX = currentX;
                    startY = currentY;
                    Include(currentX, currentY);
                    command = relative ? 'l' : 'L';
                    break;
                case 'L':
                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out currentX, out currentY))
                    {
                        bounds = default;
                        return false;
                    }

                    Include(currentX, currentY);
                    break;
                case 'H':
                    if (!TryReadSvgNumber(tokens, ref index, out double h))
                    {
                        bounds = default;
                        return false;
                    }

                    currentX = relative ? currentX + h : h;
                    Include(currentX, currentY);
                    break;
                case 'V':
                    if (!TryReadSvgNumber(tokens, ref index, out double v))
                    {
                        bounds = default;
                        return false;
                    }

                    currentY = relative ? currentY + v : v;
                    Include(currentX, currentY);
                    break;
                case 'C':
                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out double c1x, out double c1y) ||
                        !TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out double c2x, out double c2y) ||
                        !TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out currentX, out currentY))
                    {
                        bounds = default;
                        return false;
                    }

                    Include(c1x, c1y);
                    Include(c2x, c2y);
                    Include(currentX, currentY);
                    break;
                case 'Z':
                    currentX = startX;
                    currentY = startY;
                    break;
                default:
                    bounds = default;
                    return false;
            }
        }

        if (double.IsInfinity(minX) || double.IsInfinity(minY) || double.IsInfinity(maxX) || double.IsInfinity(maxY))
        {
            bounds = default;
            return false;
        }

        bounds = new SvgPathBounds(minX, minY, maxX, maxY);
        return true;
    }

    private static bool TryAppendSvgPath(PdfGraphicsBuilder graphics, string data, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY)
    {
        MatchCollection tokens = SvgPathTokenRegex().Matches(data);
        if (tokens.Count == 0)
        {
            return false;
        }

        int index = 0;
        char command = '\0';
        double currentX = 0d;
        double currentY = 0d;
        double startX = 0d;
        double startY = 0d;
        bool hasPath = false;
        while (index < tokens.Count)
        {
            string token = tokens[index].Value;
            if (char.IsLetter(token[0]))
            {
                command = token[0];
                index++;
            }

            bool relative = char.IsLower(command);
            switch (char.ToUpperInvariant(command))
            {
                case 'M':
                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out currentX, out currentY))
                    {
                        return hasPath;
                    }

                    startX = currentX;
                    startY = currentY;
                    graphics.MoveTo(SvgX(currentX), SvgY(currentY));
                    hasPath = true;
                    command = relative ? 'l' : 'L';
                    break;
                case 'L':
                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out currentX, out currentY))
                    {
                        return hasPath;
                    }

                    graphics.LineTo(SvgX(currentX), SvgY(currentY));
                    break;
                case 'H':
                    if (!TryReadSvgNumber(tokens, ref index, out double h))
                    {
                        return hasPath;
                    }

                    currentX = relative ? currentX + h : h;
                    graphics.LineTo(SvgX(currentX), SvgY(currentY));
                    break;
                case 'V':
                    if (!TryReadSvgNumber(tokens, ref index, out double v))
                    {
                        return hasPath;
                    }

                    currentY = relative ? currentY + v : v;
                    graphics.LineTo(SvgX(currentX), SvgY(currentY));
                    break;
                case 'C':
                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out double c1x, out double c1y) ||
                        !TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out double c2x, out double c2y) ||
                        !TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out double endX, out double endY))
                    {
                        return hasPath;
                    }

                    graphics.CurveTo(SvgX(c1x), SvgY(c1y), SvgX(c2x), SvgY(c2y), SvgX(endX), SvgY(endY));
                    currentX = endX;
                    currentY = endY;
                    break;
                case 'Z':
                    graphics.ClosePath();
                    currentX = startX;
                    currentY = startY;
                    break;
                default:
                    return hasPath;
            }
        }

        return hasPath;

        double SvgX(double value) => imageX + (value - minX) * scaleX;
        double SvgY(double value) => imageY + imageHeight - (value - minY) * scaleY;
    }

    private static bool TryReadSvgPoint(MatchCollection tokens, ref int index, bool relative, double currentX, double currentY, out double x, out double y)
    {
        if (!TryReadSvgNumber(tokens, ref index, out double rawX) ||
            !TryReadSvgNumber(tokens, ref index, out double rawY))
        {
            x = 0d;
            y = 0d;
            return false;
        }

        x = relative ? currentX + rawX : rawX;
        y = relative ? currentY + rawY : rawY;
        return true;
    }

    private static bool TryReadSvgNumber(MatchCollection tokens, ref int index, out double value)
    {
        value = 0d;
        if (index >= tokens.Count || char.IsLetter(tokens[index].Value[0]))
        {
            return false;
        }

        value = double.Parse(tokens[index].Value, CultureInfo.InvariantCulture);
        index++;
        return true;
    }

    private static PdfImageXObject? GetOrCreateImage(
        PptxSceneImageResource imageResource,
        PptxSceneImageRecolor recolor,
        Dictionary<string, PdfImageXObject?>? imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex)
    {
        return GetOrCreateImage(imageResource.PartName, imageResource.ContentType, imageResource.Bytes, recolor, imageCache, diagnosticSink, slideIndex);
    }

    private static PdfImageXObject? GetOrCreateImage(
        string partName,
        string contentType,
        byte[] bytes,
        PptxSceneImageRecolor recolor,
        Dictionary<string, PdfImageXObject?>? imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex)
    {
        string cacheKey = partName + "\u001f" + ImageRecolorCacheKey(recolor);
        if (imageCache is not null && imageCache.TryGetValue(cacheKey, out PdfImageXObject? cached))
        {
            return cached;
        }

        PdfImageXObject? image = CreateImage(partName, contentType, bytes, recolor, diagnosticSink, slideIndex);
        imageCache?.TryAdd(cacheKey, image);
        return image;
    }

    private static PdfImageXObject? GetOrCreateCroppedImage(
        PptxSceneImageResource imageResource,
        PptxSceneImageRecolor recolor,
        CropRect crop,
        Dictionary<string, PdfImageXObject?>? imageCache,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex)
    {
        string cacheKey = imageResource.PartName + "\u001f" + ImageRecolorCacheKey(recolor) + "\u001fcrop:" +
            crop.Left.ToString("R", CultureInfo.InvariantCulture) + "," +
            crop.Top.ToString("R", CultureInfo.InvariantCulture) + "," +
            crop.Right.ToString("R", CultureInfo.InvariantCulture) + "," +
            crop.Bottom.ToString("R", CultureInfo.InvariantCulture);
        if (imageCache is not null && imageCache.TryGetValue(cacheKey, out PdfImageXObject? cached))
        {
            return cached;
        }

        PdfImageXObject? image = CreateCroppedImage(
            imageResource.PartName,
            imageResource.ContentType,
            imageResource.Bytes,
            recolor,
            crop,
            diagnosticSink,
            slideIndex);
        imageCache?.TryAdd(cacheKey, image);
        return image;
    }

    private static PdfImageXObject? CreateImage(
        string partName,
        string contentType,
        byte[] bytes,
        PptxSceneImageRecolor recolor,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex)
    {
        try
        {
            if (contentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                contentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
            {
                JpegInfo info = JpegInfo.Read(bytes);
                if (!IsNoImageRecolor(recolor))
                {
                    if (info.IsBaselineDct && info.BitsPerComponent == 8 && info.ComponentCount is 1 or 3)
                    {
                        try
                        {
                            JpegImage jpeg = JpegImage.Read(bytes);
                            byte[] rgb = ApplyImageRecolor(jpeg.Rgb, recolor);
                            return PdfImageXObject.RgbPng(jpeg.Width, jpeg.Height, rgb, alpha: null);
                        }
                        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or IndexOutOfRangeException)
                        {
                        }
                    }

                    diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                        "PPTX_UNSUPPORTED_IMAGE_RECOLOR",
                        OoxPdfSeverity.Warning,
                        $"PPTX {ImageRecolorKindName(recolor)} image recolor could not be applied to {contentType} {info.FrameProfileName} image data and was ignored.",
                        partName,
                        SlideIndex: slideIndex,
                        Feature: "image recolor",
                        Fallback: "Original image"));
                }

                return PdfImageXObject.Jpeg(info.Width, info.Height, bytes, info.ComponentCount, info.BitsPerComponent);
            }

            if (contentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
            {
                PngImage png = PngImage.Read(bytes);
                byte[] rgb = ApplyImageRecolor(png.Rgb, recolor);
                return PdfImageXObject.RgbPng(png.Width, png.Height, rgb, png.Alpha);
            }

            if (contentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
                contentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase))
            {
                BmpImage bmp = BmpImage.Read(bytes);
                byte[] rgb = ApplyImageRecolor(bmp.Rgb, recolor);
                return PdfImageXObject.RgbPng(bmp.Width, bmp.Height, rgb, bmp.Alpha);
            }

            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{contentType}' could not be rendered and was ignored: Unsupported image content type.",
                partName,
                SlideIndex: slideIndex,
                Feature: contentType,
                Fallback: "Ignored"));
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Error,
                $"Image '{contentType}' could not be rendered and was ignored: {ex.Message}",
                partName,
                SlideIndex: slideIndex,
                Feature: contentType,
                Fallback: "Ignored"));
        }

        return null;
    }

    private static PdfImageXObject? CreateCroppedImage(
        string partName,
        string contentType,
        byte[] bytes,
        PptxSceneImageRecolor recolor,
        CropRect crop,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex)
    {
        try
        {
            if (contentType.Equals("image/png", StringComparison.OrdinalIgnoreCase))
            {
                PngImage png = PngImage.Read(bytes);
                return CreateCroppedRgbImage(png.Width, png.Height, png.Rgb, png.Alpha, recolor, crop);
            }

            if (contentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) ||
                contentType.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase))
            {
                BmpImage bmp = BmpImage.Read(bytes);
                return CreateCroppedRgbImage(bmp.Width, bmp.Height, bmp.Rgb, bmp.Alpha, recolor, crop);
            }

            if (contentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                contentType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase))
            {
                JpegImage jpeg = JpegImage.Read(bytes);
                return CreateCroppedRgbImage(jpeg.Width, jpeg.Height, jpeg.Rgb, alpha: null, recolor, crop);
            }
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or IndexOutOfRangeException)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_CROP_UNSUPPORTED_FORMAT",
                OoxPdfSeverity.Warning,
                $"Image '{partName}' could not be decoded for Office-style cropped image embedding on slide {slideIndex}; falling back to PDF clipping.",
                partName,
                SlideIndex: slideIndex,
                Feature: "image crop",
                Fallback: "PDF clipping"));
        }

        return null;
    }

    private static PdfImageXObject? CreateCroppedRgbImage(
        int width,
        int height,
        byte[] rgb,
        byte[]? alpha,
        PptxSceneImageRecolor recolor,
        CropRect crop)
    {
        int left = (int)Math.Floor(width * Math.Clamp(crop.Left, 0d, 1d));
        int top = (int)Math.Floor(height * Math.Clamp(crop.Top, 0d, 1d));
        int right = (int)Math.Ceiling(width * (1d - Math.Clamp(crop.Right, 0d, 1d)));
        int bottom = (int)Math.Ceiling(height * (1d - Math.Clamp(crop.Bottom, 0d, 1d)));

        left = Math.Clamp(left, 0, width);
        top = Math.Clamp(top, 0, height);
        right = Math.Clamp(right, left + 1, width);
        bottom = Math.Clamp(bottom, top + 1, height);

        int croppedWidth = right - left;
        int croppedHeight = bottom - top;
        if (croppedWidth <= 0 || croppedHeight <= 0)
        {
            return null;
        }

        byte[] croppedRgb = new byte[croppedWidth * croppedHeight * 3];
        byte[]? croppedAlpha = alpha is null ? null : new byte[croppedWidth * croppedHeight];
        for (int y = 0; y < croppedHeight; y++)
        {
            int sourceY = top + y;
            Buffer.BlockCopy(rgb, (sourceY * width + left) * 3, croppedRgb, y * croppedWidth * 3, croppedWidth * 3);
            if (alpha is not null && croppedAlpha is not null)
            {
                Buffer.BlockCopy(alpha, sourceY * width + left, croppedAlpha, y * croppedWidth, croppedWidth);
            }
        }

        byte[] recoloredRgb = ApplyImageRecolor(croppedRgb, recolor);
        return PdfImageXObject.RgbPng(croppedWidth, croppedHeight, recoloredRgb, croppedAlpha);
    }

    private static byte[] ApplyImageRecolor(byte[] rgb, PptxSceneImageRecolor recolor)
    {
        if (IsNoImageRecolor(recolor))
        {
            return rgb;
        }

        byte[] transformed = new byte[rgb.Length];
        for (int i = 0; i < rgb.Length; i += 3)
        {
            double red = rgb[i];
            double green = rgb[i + 1];
            double blue = rgb[i + 2];
            if (recolor.Kind == PptxSceneImageRecolorKind.Luminance)
            {
                transformed[i] = ApplyBrightnessContrast(red, recolor.Brightness, recolor.Contrast);
                transformed[i + 1] = ApplyBrightnessContrast(green, recolor.Brightness, recolor.Contrast);
                transformed[i + 2] = ApplyBrightnessContrast(blue, recolor.Brightness, recolor.Contrast);
                continue;
            }

            double luma = (0.2126d * red + 0.7152d * green + 0.0722d * blue) / 255d;
            if (recolor.Kind == PptxSceneImageRecolorKind.Grayscale)
            {
                byte gray = ToByte(luma * 255d);
                transformed[i] = gray;
                transformed[i + 1] = gray;
                transformed[i + 2] = gray;
                continue;
            }

            if (recolor.Kind == PptxSceneImageRecolorKind.BiLevel)
            {
                byte value = luma >= recolor.Threshold ? (byte)255 : (byte)0;
                transformed[i] = value;
                transformed[i + 1] = value;
                transformed[i + 2] = value;
                continue;
            }

            transformed[i] = Interpolate(recolor.Dark.Red, recolor.Light.Red, luma);
            transformed[i + 1] = Interpolate(recolor.Dark.Green, recolor.Light.Green, luma);
            transformed[i + 2] = Interpolate(recolor.Dark.Blue, recolor.Light.Blue, luma);
        }

        return transformed;
    }

    private static bool IsNoImageRecolor(PptxSceneImageRecolor recolor)
    {
        return recolor.Kind == PptxSceneImageRecolorKind.None;
    }

    private static string ImageRecolorKindName(PptxSceneImageRecolor recolor)
    {
        return recolor.Kind switch
        {
            PptxSceneImageRecolorKind.Luminance => "luminance",
            PptxSceneImageRecolorKind.Duotone => "duotone",
            PptxSceneImageRecolorKind.Grayscale => "grayscale",
            PptxSceneImageRecolorKind.BiLevel => "bi-level",
            _ => "none"
        };
    }

    private static string ImageRecolorCacheKey(PptxSceneImageRecolor recolor)
    {
        return recolor.Kind switch
        {
            PptxSceneImageRecolorKind.Luminance => FormattableString.Invariant($"lum:{recolor.Brightness:0.#####}:{recolor.Contrast:0.#####}"),
            PptxSceneImageRecolorKind.Duotone => FormattableString.Invariant($"duo:{recolor.Dark.Red:X2}{recolor.Dark.Green:X2}{recolor.Dark.Blue:X2}:{recolor.Light.Red:X2}{recolor.Light.Green:X2}{recolor.Light.Blue:X2}"),
            PptxSceneImageRecolorKind.Grayscale => "gray",
            PptxSceneImageRecolorKind.BiLevel => FormattableString.Invariant($"bi:{recolor.Threshold:0.#####}"),
            _ => "none"
        };
    }

    private static byte ApplyBrightnessContrast(double channel, double brightness, double contrast)
    {
        double value = channel / 255d;
        value = (value - 0.5d) * Math.Max(0d, 1d + contrast) + 0.5d;
        value += brightness;
        return ToByte(value * 255d);
    }

    private static byte Interpolate(byte from, byte to, double ratio)
    {
        return ToByte(from + (to - from) * Math.Clamp(ratio, 0d, 1d));
    }

    [GeneratedRegex(@"[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?", RegexOptions.CultureInvariant)]
    private static partial Regex SvgNumberRegex();

    [GeneratedRegex(@"[MmLlHhVvCcZz]|[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?", RegexOptions.CultureInvariant)]
    private static partial Regex SvgPathTokenRegex();
}
