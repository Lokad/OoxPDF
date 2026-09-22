using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void DrawOuterShadow(
        PdfGraphicsBuilder graphics,
        string preset,
        double x,
        double y,
        double width,
        double height,
        OuterShadow shadow)
    {
        graphics.SaveState();
        graphics.SetAlpha(shadow.Alpha, 1d);
        graphics.SetFillRgb(shadow.Color.Red, shadow.Color.Green, shadow.Color.Blue);
        graphics.FillRectangleEvenOdd(x + shadow.OffsetX, y + shadow.OffsetY, width, height);
        graphics.RestoreState();
    }

    private static void DrawOuterShadow(
        PdfGraphicsBuilder graphics,
        XElement shapeProperties,
        string preset,
        double x,
        double y,
        double width,
        double height,
        OuterShadow shadow,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride)
    {
        graphics.SaveState();
        graphics.SetAlpha(shadow.Alpha, 1d);
        graphics.SetFillRgb(shadow.Color.Red, shadow.Color.Green, shadow.Color.Blue);
        DrawPresetFill(graphics, shapeProperties, preset, x + shadow.OffsetX, y + shadow.OffsetY, width, height, presetAdjustmentsOverride);
        graphics.RestoreState();
    }

    private static void DrawRasterOuterShadow(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        OuterShadow shadow,
        List<PdfImageResource> images,
        ref int imageIndex)
    {
        if (shadow.BlurRadius <= 0d ||
            shadow.Alpha <= 0d ||
            width <= 0d ||
            height <= 0d)
        {
            return;
        }

        double shadowExtent = shadow.BlurRadius * OfficeOuterShadowRasterExtentFactor;
        double shadowX = x + shadow.OffsetX - shadowExtent;
        double shadowY = y + shadow.OffsetY - shadowExtent;
        double shadowWidth = width + 2d * shadowExtent;
        double shadowHeight = height + 2d * shadowExtent;
        int pixelWidth = Math.Clamp((int)Math.Ceiling(shadowWidth * OfficeGlowRasterPixelsPerPoint), 1, OfficeGlowRasterMaxPixelsPerSide);
        int pixelHeight = Math.Clamp((int)Math.Ceiling(shadowHeight * OfficeGlowRasterPixelsPerPoint), 1, OfficeGlowRasterMaxPixelsPerSide);
        double scaleX = pixelWidth / shadowWidth;
        OoxConversionBudget.Current?.ChargeImagesDecoded(1);
        double scaleY = pixelHeight / shadowHeight;

        // R02: effect rasters hold decoded-size planes plus compression scratch; reserve across rasterization and PDF compression.
        using var effectReservation = OoxConversionBudget.Current?.ReserveLiveImageBytes(checked((long)pixelWidth * pixelHeight * 4L));
        byte[] rgb = new byte[pixelWidth * pixelHeight * 3];
        byte[] alpha = new byte[pixelWidth * pixelHeight];
        for (int pixelY = 0; pixelY < pixelHeight; pixelY++)
        {
            double localY = (pixelY + 0.5d) / scaleY - shadowExtent;
            double dy = localY < 0d ? -localY : localY > height ? localY - height : 0d;
            for (int pixelX = 0; pixelX < pixelWidth; pixelX++)
            {
                double localX = (pixelX + 0.5d) / scaleX - shadowExtent;
                double dx = localX < 0d ? -localX : localX > width ? localX - width : 0d;
                int pixel = pixelY * pixelWidth + pixelX;
                int rgbOffset = pixel * 3;
                rgb[rgbOffset] = shadow.Color.Red;
                rgb[rgbOffset + 1] = shadow.Color.Green;
                rgb[rgbOffset + 2] = shadow.Color.Blue;

                double distance = Math.Sqrt(dx * dx + dy * dy);
                double t = Math.Clamp(distance / shadow.BlurRadius, 0d, 1d);
                double falloff = (1d - t) * (1d - t);
                alpha[pixel] = (byte)Math.Clamp((int)Math.Round(255d * falloff), 0, 255);
            }
        }

        var image = PdfImageXObject.RgbPng(pixelWidth, pixelHeight, rgb, alpha);
        string name = "Im" + imageIndex++;
        graphics.SaveState();
        graphics.SetAlpha(shadow.Alpha, 1d);
        graphics.DrawImage(name, shadowX, shadowY, shadowWidth, shadowHeight);
        graphics.RestoreState();
        images.Add(new PdfImageResource(name, image));
    }

    private static void DrawPresetFill(
        PdfGraphicsBuilder graphics,
        XElement shapeProperties,
        string preset,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride)
    {
        if (preset == "ellipse")
        {
            graphics.FillEllipseEvenOdd(x, y, width, height);
        }
        else if (preset == "roundRect")
        {
            graphics.FillRoundedRectangleEvenOdd(x, y, width, height, ReadRoundRectangleRadius(shapeProperties, presetAdjustmentsOverride, width, height));
        }
        else if (TryCreatePresetPolygonPoints(preset, x, y, width, height, out (double X, double Y)[] polygonPoints))
        {
            graphics.FillPolygonEvenOdd(polygonPoints);
        }
        else
        {
            graphics.FillRectangleEvenOdd(x, y, width, height);
        }
    }

    private static bool CanRenderGlowPreset(string preset)
    {
        return preset == "rect";
    }

    private static void DrawGlow(
        PdfGraphicsBuilder graphics,
        string preset,
        double x,
        double y,
        double width,
        double height,
        Glow glow)
    {
        graphics.SaveState();
        graphics.SetAlpha(glow.Alpha, 1d);
        graphics.SetFillRgb(glow.Color.Red, glow.Color.Green, glow.Color.Blue);
        graphics.FillRectangleEvenOdd(x - glow.Radius, y - glow.Radius, width + 2d * glow.Radius, height + 2d * glow.Radius);
        graphics.RestoreState();
    }

    private static void DrawRasterGlow(
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        Glow glow,
        List<PdfImageResource> images,
        ref int imageIndex)
    {
        if (glow.Radius <= 0d ||
            glow.Alpha <= 0d ||
            width <= 0d ||
            height <= 0d)
        {
            return;
        }

        double glowX = x - glow.Radius;
        double glowY = y - glow.Radius;
        double glowWidth = width + 2d * glow.Radius;
        double glowHeight = height + 2d * glow.Radius;
        int pixelWidth = Math.Clamp((int)Math.Ceiling(glowWidth * OfficeGlowRasterPixelsPerPoint), 1, OfficeGlowRasterMaxPixelsPerSide);
        int pixelHeight = Math.Clamp((int)Math.Ceiling(glowHeight * OfficeGlowRasterPixelsPerPoint), 1, OfficeGlowRasterMaxPixelsPerSide);
        double scaleX = pixelWidth / glowWidth;
        double scaleY = pixelHeight / glowHeight;

        OoxConversionBudget.Current?.ChargeImagesDecoded(1);
        // R02: effect rasters hold decoded-size planes plus compression scratch; reserve across rasterization and PDF compression.
        using var effectReservation = OoxConversionBudget.Current?.ReserveLiveImageBytes(checked((long)pixelWidth * pixelHeight * 4L));
        byte[] rgb = new byte[pixelWidth * pixelHeight * 3];
        byte[] alpha = new byte[pixelWidth * pixelHeight];
        for (int pixelY = 0; pixelY < pixelHeight; pixelY++)
        {
            double localY = (pixelY + 0.5d) / scaleY - glow.Radius;
            double dy = localY < 0d ? -localY : localY > height ? localY - height : 0d;
            for (int pixelX = 0; pixelX < pixelWidth; pixelX++)
            {
                double localX = (pixelX + 0.5d) / scaleX - glow.Radius;
                double dx = localX < 0d ? -localX : localX > width ? localX - width : 0d;
                int pixel = pixelY * pixelWidth + pixelX;
                int rgbOffset = pixel * 3;
                rgb[rgbOffset] = glow.Color.Red;
                rgb[rgbOffset + 1] = glow.Color.Green;
                rgb[rgbOffset + 2] = glow.Color.Blue;

                if (dx <= 0d && dy <= 0d)
                {
                    continue;
                }

                double distance = Math.Sqrt(dx * dx + dy * dy);
                if (distance >= glow.Radius)
                {
                    continue;
                }

                double t = distance / glow.Radius;
                double falloff = (1d - t) * (1d - t);
                alpha[pixel] = (byte)Math.Clamp((int)Math.Round(255d * falloff), 0, 255);
            }
        }

        var image = PdfImageXObject.RgbPng(pixelWidth, pixelHeight, rgb, alpha);
        string name = "Im" + imageIndex++;
        graphics.SaveState();
        graphics.SetAlpha(glow.Alpha, 1d);
        graphics.DrawImage(name, glowX, glowY, glowWidth, glowHeight);
        graphics.RestoreState();
        images.Add(new PdfImageResource(name, image));
    }

    private static void DrawLinearGradientFill(PdfGraphicsBuilder graphics, GradientFill gradient, double x, double y, double width, double height)
    {
        if (gradient.Stops.Count < 2)
        {
            return;
        }

        double radians = DegreesToRadians(gradient.AngleDegrees);
        double dx = Math.Cos(radians);
        // Office places the first gradient stop at the top for a 90-degree linear gradient;
        // emission space is y-up PDF, so the sine component is negated to keep stop 0 at the Office end.
        double dy = -Math.Sin(radians);
        double half = Math.Abs(dx) * width / 2d + Math.Abs(dy) * height / 2d;
        double centerX = x + width / 2d;
        double centerY = y + height / 2d;
        bool alphaState = PptxColorResolver.TryGetUniformGradientAlpha(gradient.Stops, static stop => stop.Alpha, out double alpha) && alpha < 0.999d;
        if (alphaState)
        {
            graphics.SaveState();
            graphics.SetAlpha(alpha, 1d);
        }

        graphics.PaintAxialShading(
            centerX - dx * half,
            centerY - dy * half,
            centerX + dx * half,
            centerY + dy * half,
            gradient.Stops.Select(stop => new PdfShadingStop(stop.Offset, stop.Color.Red, stop.Color.Green, stop.Color.Blue)).ToArray());
        if (alphaState)
        {
            graphics.RestoreState();
        }
    }

    private static void DrawPresetArcStroke(
        PdfGraphicsBuilder graphics,
        XElement shapeProperties,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride)
    {
        double startDegrees = ReadPresetGeometryGuide(shapeProperties, presetAdjustmentsOverride, "adj1", OfficePresetArcDefaultStartGuide) / 60000d;
        double endDegrees = ReadPresetGeometryGuide(shapeProperties, presetAdjustmentsOverride, "adj2", OfficePresetArcDefaultEndGuide) / 60000d;
        double sweepDegrees = endDegrees - startDegrees;
        while (sweepDegrees <= 0d)
        {
            sweepDegrees += 360d;
        }

        double centerX = x + width / 2d;
        double centerY = y + height / 2d;
        double radiusX = width / 2d;
        double radiusY = height / 2d;
        double remaining = sweepDegrees;
        double current = startDegrees;
        bool first = true;
        while (remaining > 0.0001d)
        {
            double segment = Math.Min(90d, remaining);
            AppendEllipseArcSegment(graphics, centerX, centerY, radiusX, radiusY, current, segment, first);
            first = false;
            current += segment;
            remaining -= segment;
        }

        graphics.StrokeCurrentPath();
    }

    private static bool TryFillPresetArcLineEndOutline(
        PdfGraphicsBuilder graphics,
        XElement shapeProperties,
        double x,
        double y,
        double width,
        double height,
        double lineWidth,
        LineEndStyle tailEnd,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride)
    {
        if (width <= 0d || height <= 0d || lineWidth <= 0d)
        {
            return false;
        }

        double startDegrees = ReadPresetGeometryGuide(shapeProperties, presetAdjustmentsOverride, "adj1", OfficePresetArcDefaultStartGuide) / 60000d;
        double endDegrees = ReadPresetGeometryGuide(shapeProperties, presetAdjustmentsOverride, "adj2", OfficePresetArcDefaultEndGuide) / 60000d;
        double sweepDegrees = endDegrees - startDegrees;
        while (sweepDegrees <= 0d)
        {
            sweepDegrees += 360d;
        }

        double centerX = x + width / 2d;
        double centerY = y + height / 2d;
        double radiusX = width / 2d;
        double radiusY = height / 2d;
        List<CurveSample> samples = SamplePresetArc(centerX, centerY, radiusX, radiusY, startDegrees, sweepDegrees);
        if (samples.Count < 2)
        {
            return false;
        }

        double[] cumulativeLengths = BuildCumulativeSampleLengths(samples, out double totalLength);
        double markerLength = tailEnd.Kind switch
        {
            LineEndKind.Arrow => lineWidth * OfficeArrowheadLengthFactor * tailEnd.LengthScale,
            LineEndKind.Stealth => lineWidth * OfficeStraightStealthLineEndLengthFactor * tailEnd.LengthScale,
            _ => Math.Max(OfficeTriangleTailMinimumLength, lineWidth * OfficeTriangleTailLengthFactor) * tailEnd.LengthScale
        };
        double baseDistance = Math.Max(0d, totalLength - markerLength);
        CurveSample baseSample = SampleAtDistance(samples, cumulativeLengths, baseDistance);
        List<CurveSample> bodySamples = SelectBodySamples(samples, cumulativeLengths, baseDistance, baseSample);
        if (bodySamples.Count < 2)
        {
            return false;
        }

        (double X, double Y) direction = Normalize(baseSample.TangentX, baseSample.TangentY);
        if (Math.Abs(direction.X) <= 0.000001d && Math.Abs(direction.Y) <= 0.000001d)
        {
            return false;
        }

        double halfWidth = Math.Max(0.1d, lineWidth / 2d);
        var points = new List<(double X, double Y)>(bodySamples.Count * 2);
        foreach (CurveSample sample in bodySamples)
        {
            (double X, double Y) normal = NormalForSample(sample);
            points.Add((sample.X + normal.X * halfWidth, sample.Y + normal.Y * halfWidth));
        }

        for (int i = bodySamples.Count - 1; i >= 0; i--)
        {
            CurveSample sample = bodySamples[i];
            (double X, double Y) normal = NormalForSample(sample);
            points.Add((sample.X - normal.X * halfWidth, sample.Y - normal.Y * halfWidth));
        }

        AppendClosedLinePath(graphics, points, explicitClosingLine: true);
        (double X, double Y) tip = (samples[^1].X, samples[^1].Y);
        (double X, double Y) normalAtBase = (-direction.Y, direction.X);
        if (tailEnd.Kind == LineEndKind.Arrow)
        {
            AppendOfficeArrowHeadPath(
                graphics,
                tip.X,
                tip.Y,
                -direction.X,
                -direction.Y,
                -normalAtBase.X,
                -normalAtBase.Y,
                lineWidth,
                splitTrailingCurve: true);
        }
        else if (tailEnd.Kind == LineEndKind.Stealth)
        {
            double markerWidth = lineWidth * OfficeStraightStealthLineEndWidthFactor * tailEnd.WidthScale;
            AppendClosedLinePath(graphics,
            [
                tip,
                (baseSample.X + normalAtBase.X * markerWidth / 2d, baseSample.Y + normalAtBase.Y * markerWidth / 2d),
                (tip.X - direction.X * markerLength * OfficeStraightStealthLineEndNotchFactor, tip.Y - direction.Y * markerLength * OfficeStraightStealthLineEndNotchFactor),
                (baseSample.X - normalAtBase.X * markerWidth / 2d, baseSample.Y - normalAtBase.Y * markerWidth / 2d)],
            false);
        }
        else
        {
            double arrowHalfWidth = markerLength * OfficeTriangleTailHalfWidthFactor * tailEnd.WidthScale;
            AppendClosedLinePath(graphics,
            [
                tip,
                (baseSample.X + normalAtBase.X * arrowHalfWidth, baseSample.Y + normalAtBase.Y * arrowHalfWidth),
                (baseSample.X - normalAtBase.X * arrowHalfWidth, baseSample.Y - normalAtBase.Y * arrowHalfWidth)],
            false);
        }

        graphics.FillCurrentPath();
        return true;
    }

    private static List<CurveSample> SamplePresetArc(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        double startDegrees,
        double sweepDegrees)
    {
        double maxRadius = Math.Max(radiusX, radiusY);
        double sweepRadians = Math.Abs(DegreesToRadians(sweepDegrees));
        int flatnessSegments = maxRadius > OfficePresetArcFlatteningTolerance
            ? (int)Math.Ceiling(sweepRadians / (2d * Math.Acos(Math.Max(-1d, 1d - OfficePresetArcFlatteningTolerance / maxRadius))))
            : 1;
        int angularSegments = (int)Math.Ceiling(Math.Abs(sweepDegrees) / OfficePresetArcMaximumFlatteningDegrees);
        int segmentCount = Math.Clamp(Math.Max(flatnessSegments, angularSegments), 2, 128);
        var samples = new List<CurveSample>(segmentCount + 1);
        for (int i = 0; i <= segmentCount; i++)
        {
            double degrees = startDegrees + sweepDegrees * i / segmentCount;
            double parameter = ConvertVisualAngleToEllipseParameter(degrees, radiusX, radiusY);
            double sampleX = centerX + radiusX * Math.Cos(parameter);
            double sampleY = centerY - radiusY * Math.Sin(parameter);
            double tangentX = -radiusX * Math.Sin(parameter);
            double tangentY = -radiusY * Math.Cos(parameter);
            samples.Add(new CurveSample(sampleX, sampleY, tangentX, tangentY));
        }

        return samples;
    }

    private static double ReadPresetGeometryGuide(
        XElement shapeProperties,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride,
        string name,
        double fallback)
    {
        if (presetAdjustmentsOverride is not null && presetAdjustmentsOverride.TryGetValue(name, out double resolved))
        {
            return resolved;
        }

        XElement? guide = shapeProperties
            .Element(DrawingNamespace + "prstGeom")
            ?.Element(DrawingNamespace + "avLst")
            ?.Elements(DrawingNamespace + "gd")
            .FirstOrDefault(element => string.Equals((string?)element.Attribute("name"), name, StringComparison.Ordinal));
        string? formula = (string?)guide?.Attribute("fmla");
        if (formula is null || !formula.StartsWith("val ", StringComparison.Ordinal))
        {
            return fallback;
        }

        return double.TryParse(formula[4..], NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
            ? value
            : fallback;
    }

    private static double ReadRoundRectangleRadius(
        XElement shapeProperties,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride,
        double width,
        double height)
    {
        double adjustment = ReadPresetGeometryGuide(shapeProperties, presetAdjustmentsOverride, "adj", 16667d);
        double factor = Math.Clamp(adjustment / 100000d, 0d, 0.5d);
        return Math.Min(width, height) * factor;
    }

    private static void AppendEllipseArcSegment(PdfGraphicsBuilder graphics, double centerX, double centerY, double radiusX, double radiusY, double startDegrees, double sweepDegrees, bool moveToStart)
    {
        BezierSegment segment = CreateEllipseArcSegment();
        if (moveToStart)
        {
            graphics.MoveTo(segment.StartX, segment.StartY);
        }

        graphics.CurveTo(segment.Control1X, segment.Control1Y, segment.Control2X, segment.Control2Y, segment.EndX, segment.EndY);

        BezierSegment CreateEllipseArcSegment()
        {
            double start = ConvertVisualAngleToEllipseParameter(startDegrees, radiusX, radiusY);
            double end = ConvertVisualAngleToEllipseParameter(startDegrees + sweepDegrees, radiusX, radiusY);
            double sweep = end - start;
            if (sweepDegrees >= 0d)
            {
                while (sweep <= 0d)
                {
                    sweep += Math.Tau;
                }
            }
            else
            {
                while (sweep >= 0d)
                {
                    sweep -= Math.Tau;
                }
            }

            double endParameter = start + sweep;
            double k = 4d / 3d * Math.Tan(sweep / 4d);

            double x0 = centerX + radiusX * Math.Cos(start);
            double y0 = centerY - radiusY * Math.Sin(start);
            double x3 = centerX + radiusX * Math.Cos(endParameter);
            double y3 = centerY - radiusY * Math.Sin(endParameter);
            double x1 = x0 - radiusX * k * Math.Sin(start);
            double y1 = y0 - radiusY * k * Math.Cos(start);
            double x2 = x3 + radiusX * k * Math.Sin(endParameter);
            double y2 = y3 + radiusY * k * Math.Cos(endParameter);

            return new BezierSegment(x0, y0, x1, y1, x2, y2, x3, y3);
        }
    }

    private static double ConvertVisualAngleToEllipseParameter(double degrees, double radiusX, double radiusY)
    {
        double angle = DegreesToRadians(degrees);
        if (radiusX <= 0d || radiusY <= 0d)
        {
            return angle;
        }

        return Math.Atan2(Math.Sin(angle) / radiusY, Math.Cos(angle) / radiusX);
    }

    private static string ReadPreset(XElement shapeProperties)
    {
        return (string?)shapeProperties
            .Element(DrawingNamespace + "prstGeom")
            ?.Attribute("prst") ?? "rect";
    }
}
