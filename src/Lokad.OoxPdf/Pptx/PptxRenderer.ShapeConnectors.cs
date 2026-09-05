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
    private static bool TryReadShapePictureFill(
        XElement shapeProperties,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        List<PdfImageResource>? images,
        Dictionary<string, PdfImageXObject?>? imageCache,
        ref int imageIndex,
        ShapePictureFill? pictureFillOverride,
        out string? name,
        out PdfImageXObject? image,
        out CropRect crop,
        out FillRect fillRect,
        out double alpha)
    {
        name = null;
        image = null;
        crop = default;
        fillRect = default;
        alpha = 1d;
        if (images is null || pictureFillOverride is null || !CanRenderPictureFillPreset(ReadPreset(shapeProperties)))
        {
            return false;
        }

        ShapePictureFill resolvedPictureFill = pictureFillOverride.Value;
        if (resolvedPictureFill.Resource is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_PART",
                OoxPdfSeverity.Error,
                "Referenced image part was missing and the image was ignored.",
                resolvedPictureFill.TargetPartName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "image",
                Fallback: "Ignored"));
            return false;
        }

        image = GetOrCreateImage(resolvedPictureFill.Resource, PptxSceneImageRecolor.None, imageCache, diagnosticSink, slideIndex);
        if (image is null)
        {
            return false;
        }

        crop = resolvedPictureFill.Crop;
        fillRect = resolvedPictureFill.Fill;
        alpha = resolvedPictureFill.Alpha;
        name = "Im" + imageIndex++;
        return true;
    }

    private static bool CanRenderPictureFillPreset(string preset)
    {
        return preset is "rect" or "ellipse" or "roundRect" ||
            TryCreatePresetPolygonPoints(preset, 0d, 0d, 1d, 1d, out _);
    }

    private static void ClipToPresetShape(
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
            graphics.ClipEllipse(x, y, width, height);
        }
        else if (preset == "roundRect")
        {
            graphics.ClipRoundedRectangle(x, y, width, height, ReadRoundRectangleRadius(shapeProperties, presetAdjustmentsOverride, width, height));
        }
        else if (TryCreatePresetPolygonPoints(preset, x, y, width, height, out (double X, double Y)[] polygonPoints))
        {
            graphics.ClipPolygon(polygonPoints);
        }
        else
        {
            graphics.ClipRectangle(x, y, width, height);
        }
    }

    private static void DrawImageFill(PdfGraphicsBuilder graphics, string imageName, double x, double y, double width, double height, CropRect crop)
    {
        if (crop.IsEmpty)
        {
            graphics.DrawImage(imageName, x, y, width, height);
        }
        else
        {
            graphics.DrawImageCropped(imageName, x, y, width, height, crop.Left, crop.Top, crop.Right, crop.Bottom);
        }
    }

    private static void FillLineArrowhead(PdfGraphicsBuilder graphics, double tipX, double tipY, double directionX, double directionY, double lineWidth)
    {
        double length = Math.Sqrt(directionX * directionX + directionY * directionY);
        if (length <= 0.001d)
        {
            return;
        }

        double ux = directionX / length;
        double uy = directionY / length;
        double nx = -uy;
        double ny = ux;
        double size = Math.Max(5d, lineWidth * 3.5d);
        double baseX = tipX - ux * size;
        double baseY = tipY - uy * size;
        double halfWidth = size * OfficeTriangleTailHalfWidthFactor;
        graphics.FillPolygon(
        [
            (tipX, tipY),
            (baseX + nx * halfWidth, baseY + ny * halfWidth),
            (baseX - nx * halfWidth, baseY - ny * halfWidth)
        ]);
    }

    private static void FillLineArrowhead(PdfGraphicsBuilder graphics, double tipX, double tipY, (double X, double Y) direction, double lineWidth)
    {
        FillLineArrowhead(graphics, tipX, tipY, direction.X, direction.Y, lineWidth);
    }

    private static void FillLineEndMarker(PdfGraphicsBuilder graphics, LineEndStyle style, double tipX, double tipY, double directionX, double directionY, double lineWidth)
    {
        if (style.IsNone || style.Kind is LineEndKind.Triangle or LineEndKind.Arrow)
        {
            return;
        }

        double length = Math.Sqrt(directionX * directionX + directionY * directionY);
        if (length <= 0.001d)
        {
            return;
        }

        double ux = directionX / length;
        double uy = directionY / length;
        double nx = -uy;
        double ny = ux;
        double markerLength = style.Kind == LineEndKind.Stealth
            ? GetStraightStealthLineEndLength(lineWidth, style)
            : Math.Max(6.5d, lineWidth * 4d) * style.LengthScale;
        double markerWidth = style.Kind == LineEndKind.Stealth
            ? GetStraightStealthLineEndWidth(lineWidth, style)
            : Math.Max(5d, lineWidth * 3.2d) * style.WidthScale;

        (double X, double Y) Point(double along, double normal)
        {
            return (tipX - ux * along + nx * normal, tipY - uy * along + ny * normal);
        }

        switch (style.Kind)
        {
            case LineEndKind.Stealth:
                graphics.FillPolygon(
                [
                    (tipX, tipY),
                    Point(markerLength, markerWidth / 2d),
                    Point(markerLength * OfficeStraightStealthLineEndNotchFactor, 0d),
                    Point(markerLength, -markerWidth / 2d)
                ]);
                break;
            case LineEndKind.Diamond:
                graphics.FillPolygon(
                [
                    (tipX, tipY),
                    Point(markerLength / 2d, markerWidth / 2d),
                    Point(markerLength, 0d),
                    Point(markerLength / 2d, -markerWidth / 2d)
                ]);
                break;
            case LineEndKind.Oval:
                (double X, double Y) center = Point(markerLength / 2d, 0d);
                graphics.FillEllipse(center.X - markerLength / 2d, center.Y - markerWidth / 2d, markerLength, markerWidth);
                break;
        }
    }

    private static void FillStealthEndedLine(PdfGraphicsBuilder graphics, double x1, double y1, double x2, double y2, double lineWidth, LineEndStyle headEnd, LineEndStyle tailEnd)
    {
        double dx = x2 - x1;
        double dy = y2 - y1;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length <= 0.001d)
        {
            return;
        }

        double ux = dx / length;
        double uy = dy / length;
        double nx = -uy;
        double ny = ux;
        double half = lineWidth / 2d;
        double headInset = headEnd.Kind == LineEndKind.Stealth
            ? GetStraightStealthLineEndLength(lineWidth, headEnd) * OfficeStraightStealthLineEndNotchFactor
            : 0d;
        double tailInset = tailEnd.Kind == LineEndKind.Stealth
            ? GetStraightStealthLineEndLength(lineWidth, tailEnd) * OfficeStraightStealthLineEndNotchFactor
            : 0d;
        double startX = x1 + ux * headInset;
        double startY = y1 + uy * headInset;
        double endX = x2 - ux * tailInset;
        double endY = y2 - uy * tailInset;

        if (Distance(startX, startY, endX, endY) > 0.001d)
        {
            AppendClosedPolygon(graphics,
            [
                (startX + nx * half, startY + ny * half),
                (endX + nx * half, endY + ny * half),
                (endX - nx * half, endY - ny * half),
                (startX - nx * half, startY - ny * half)
            ]);
        }

        if (headEnd.Kind == LineEndKind.Stealth)
        {
            AppendStealthLineEndMarker(graphics, headEnd, x1, y1, x1 - x2, y1 - y2, lineWidth);
        }

        if (tailEnd.Kind == LineEndKind.Stealth)
        {
            AppendStealthLineEndMarker(graphics, tailEnd, x2, y2, x2 - x1, y2 - y1, lineWidth);
        }

        graphics.FillCurrentPath();
    }

    private static void AppendStealthLineEndMarker(PdfGraphicsBuilder graphics, LineEndStyle style, double tipX, double tipY, double directionX, double directionY, double lineWidth)
    {
        double length = Math.Sqrt(directionX * directionX + directionY * directionY);
        if (length <= 0.001d)
        {
            return;
        }

        double ux = directionX / length;
        double uy = directionY / length;
        double nx = -uy;
        double ny = ux;
        double markerLength = GetStraightStealthLineEndLength(lineWidth, style);
        double markerWidth = GetStraightStealthLineEndWidth(lineWidth, style);

        (double X, double Y) Point(double along, double normal)
        {
            return (tipX - ux * along + nx * normal, tipY - uy * along + ny * normal);
        }

        AppendClosedPolygon(graphics,
        [
            (tipX, tipY),
            Point(markerLength, markerWidth / 2d),
            Point(markerLength * OfficeStraightStealthLineEndNotchFactor, 0d),
            Point(markerLength, -markerWidth / 2d)
        ]);
    }

    private static double GetStraightStealthLineEndLength(double lineWidth, LineEndStyle style)
    {
        return Math.Max(OfficeStraightStealthLineEndMinimumSize, lineWidth * OfficeStraightStealthLineEndLengthFactor) * style.LengthScale;
    }

    private static double GetStraightStealthLineEndWidth(double lineWidth, LineEndStyle style)
    {
        return Math.Max(OfficeStraightStealthLineEndMinimumSize, lineWidth * OfficeStraightStealthLineEndWidthFactor) * style.WidthScale;
    }

    private static void AppendClosedPolygon(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points)
    {
        if (points.Count == 0)
        {
            return;
        }

        graphics.MoveTo(points[0].X, points[0].Y);
        for (int i = 1; i < points.Count; i++)
        {
            graphics.LineTo(points[i].X, points[i].Y);
        }

        graphics.ClosePath();
    }

    private static void FillArrowedLine(
        PdfGraphicsBuilder graphics,
        double x1,
        double y1,
        double x2,
        double y2,
        double lineWidth,
        bool headArrow,
        bool tailArrow,
        double arrowHalfWidthFactor)
    {
        double dx = x2 - x1;
        double dy = y2 - y1;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length <= 0.001d)
        {
            return;
        }

        double ux = dx / length;
        double uy = dy / length;
        double nx = -uy;
        double ny = ux;
        double half = lineWidth / 2d;
        double arrowLength = lineWidth * OfficeStraightTriangleLineEndLengthFactor;
        double arrowHalfWidth = lineWidth * arrowHalfWidthFactor;
        double shaftInset = Math.Max(0d, arrowLength - lineWidth * OfficeStraightTriangleLineEndOverlapFactor);
        double startX = headArrow ? x1 + ux * shaftInset : x1;
        double startY = headArrow ? y1 + uy * shaftInset : y1;
        double endX = tailArrow ? x2 - ux * shaftInset : x2;
        double endY = tailArrow ? y2 - uy * shaftInset : y2;

        graphics.MoveTo(startX + nx * half, startY + ny * half);
        graphics.LineTo(endX + nx * half, endY + ny * half);
        graphics.LineTo(endX - nx * half, endY - ny * half);
        graphics.LineTo(startX - nx * half, startY - ny * half);
        graphics.ClosePath();

        if (headArrow)
        {
            AppendTriangle(graphics, x1, y1, x1 + ux * arrowLength, y1 + uy * arrowLength, nx, ny, arrowHalfWidth);
        }

        if (tailArrow)
        {
            AppendTriangle(graphics, x2, y2, x2 - ux * arrowLength, y2 - uy * arrowLength, nx, ny, arrowHalfWidth);
        }

        graphics.FillCurrentPath();
    }

    private static void FillTriangle(PdfGraphicsBuilder graphics, double tipX, double tipY, double baseX, double baseY, double nx, double ny, double halfWidth)
    {
        AppendTriangle(graphics, tipX, tipY, baseX, baseY, nx, ny, halfWidth);
        graphics.FillCurrentPath();
    }

    private static void AppendTriangle(PdfGraphicsBuilder graphics, double tipX, double tipY, double baseX, double baseY, double nx, double ny, double halfWidth)
    {
        graphics.MoveTo(tipX, tipY);
        graphics.LineTo(baseX + nx * halfWidth, baseY + ny * halfWidth);
        graphics.LineTo(baseX - nx * halfWidth, baseY - ny * halfWidth);
        graphics.ClosePath();
    }

    private static void FillOfficeArrowedLine(PdfGraphicsBuilder graphics, double x1, double y1, double x2, double y2, double lineWidth, bool headArrow, bool tailArrow)
    {
        double dx = x2 - x1;
        double dy = y2 - y1;
        double length = Math.Sqrt(dx * dx + dy * dy);
        if (length <= 0.001d)
        {
            return;
        }

        double ux = dx / length;
        double uy = dy / length;
        double nx = -uy;
        double ny = ux;
        double half = lineWidth / 2d;
        double shaftInset = lineWidth * 0.99d;
        double startX = headArrow ? x1 + ux * shaftInset : x1;
        double startY = headArrow ? y1 + uy * shaftInset : y1;
        double endX = tailArrow ? x2 - ux * shaftInset : x2;
        double endY = tailArrow ? y2 - uy * shaftInset : y2;

        graphics.MoveTo(startX + nx * half, startY + ny * half);
        graphics.LineTo(endX + nx * half, endY + ny * half);
        graphics.LineTo(endX - nx * half, endY - ny * half);
        graphics.LineTo(startX - nx * half, startY - ny * half);
        graphics.ClosePath();

        if (headArrow)
        {
            AppendOfficeArrowHeadPath(graphics, x1, y1, ux, uy, nx, ny, lineWidth, false);
        }

        if (tailArrow)
        {
            AppendOfficeArrowHeadPath(graphics, x2, y2, -ux, -uy, -nx, -ny, lineWidth, false);
        }

        graphics.FillCurrentPath();
    }

    private static void AppendOfficeArrowHeadPath(PdfGraphicsBuilder graphics, double tipX, double tipY, double ux, double uy, double nx, double ny, double lineWidth, bool splitTrailingCurve)
    {
        tipX -= ux * 0.01d;
        tipY -= uy * 0.01d;

        (double X, double Y) Point(double along, double normal)
        {
            return (tipX + ux * along + nx * normal, tipY + uy * along + ny * normal);
        }

        (double X, double Y) p = Point(lineWidth * 3.74d, -lineWidth * 2.181667d);
        graphics.MoveTo(p.X, p.Y);
        p = Point(0d, 0d);
        graphics.LineTo(p.X, p.Y);
        p = Point(lineWidth * 3.74d, lineWidth * 2.181667d);
        graphics.LineTo(p.X, p.Y);
        (double X, double Y) c1 = Point(lineWidth * 3.978333d, lineWidth * 2.321667d);
        (double X, double Y) c2 = Point(lineWidth * 4.285d, lineWidth * 2.24d);
        (double X, double Y) c3 = Point(lineWidth * 4.423333d, lineWidth * 2.001667d);
        graphics.CurveTo(c1.X, c1.Y, c2.X, c2.Y, c3.X, c3.Y);
        p = Point(lineWidth * 4.243333d, lineWidth * 1.318333d);
        graphics.LineTo(p.X, p.Y);
        p = Point(lineWidth * 1.243333d, -lineWidth * 0.431667d);
        graphics.LineTo(p.X, p.Y);
        p = Point(lineWidth * 1.243333d, lineWidth * 0.431667d);
        graphics.LineTo(p.X, p.Y);
        p = Point(lineWidth * 4.243333d, -lineWidth * 1.318333d);
        graphics.LineTo(p.X, p.Y);
        c1 = Point(lineWidth * 4.481667d, -lineWidth * 1.456667d);
        c2 = Point(lineWidth * 4.563333d, -lineWidth * 1.763333d);
        c3 = Point(lineWidth * 4.423333d, -lineWidth * 2.001667d);
        graphics.CurveTo(c1.X, c1.Y, c2.X, c2.Y, c3.X, c3.Y);
        c1 = Point(lineWidth * 4.285d, -lineWidth * 2.24d);
        c2 = Point(lineWidth * 3.978333d, -lineWidth * 2.321667d);
        c3 = Point(lineWidth * 3.74d, -lineWidth * 2.181667d);
        if (splitTrailingCurve)
        {
            (double X, double Y) start = Point(lineWidth * 4.423333d, -lineWidth * 2.001667d);
            ((double X, double Y) leftControl1, (double X, double Y) leftControl2, (double X, double Y) mid, (double X, double Y) rightControl1, (double X, double Y) rightControl2) = SplitCubicAtHalf(start, c1, c2, c3);
            graphics.CurveTo(leftControl1.X, leftControl1.Y, leftControl2.X, leftControl2.Y, mid.X, mid.Y);
            graphics.CurveTo(rightControl1.X, rightControl1.Y, rightControl2.X, rightControl2.Y, c3.X, c3.Y);
        }
        else
        {
            graphics.CurveTo(c1.X, c1.Y, c2.X, c2.Y, c3.X, c3.Y);
        }
        graphics.ClosePath();
    }

    private static ((double X, double Y) LeftControl1, (double X, double Y) LeftControl2, (double X, double Y) Mid, (double X, double Y) RightControl1, (double X, double Y) RightControl2) SplitCubicAtHalf(
        (double X, double Y) start,
        (double X, double Y) control1,
        (double X, double Y) control2,
        (double X, double Y) end)
    {
        (double X, double Y) p01 = Midpoint(start, control1);
        (double X, double Y) p12 = Midpoint(control1, control2);
        (double X, double Y) p23 = Midpoint(control2, end);
        (double X, double Y) p012 = Midpoint(p01, p12);
        (double X, double Y) p123 = Midpoint(p12, p23);
        (double X, double Y) p0123 = Midpoint(p012, p123);
        return (p01, p012, p0123, p123, p23);
    }

    private static (double X, double Y) Midpoint((double X, double Y) left, (double X, double Y) right)
    {
        return ((left.X + right.X) / 2d, (left.Y + right.Y) / 2d);
    }

    private static void DrawCurvedConnectorPreset(
        PdfGraphicsBuilder graphics,
        XElement shapeProperties,
        string preset,
        double x,
        double yTop,
        double width,
        double height,
        double slideHeight,
        RgbColor stroke,
        double lineWidth,
        LineEndStyle headEnd,
        LineEndStyle tailEnd,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride)
    {
        List<BezierSegment> segments = preset switch
        {
            "curvedConnector2" => CreateCurvedConnector2Segments(x, yTop, width, height, slideHeight),
            "curvedConnector3" => CreateCurvedConnector3Segments(shapeProperties, presetAdjustmentsOverride, x, yTop, width, height, slideHeight),
            _ => []
        };
        if (segments.Count == 0)
        {
            return;
        }

        BezierSegment first = segments[0];
        BezierSegment last = segments[^1];
        graphics.MoveTo(first.StartX, first.StartY);
        foreach (BezierSegment segment in segments)
        {
            graphics.CurveTo(segment.Control1X, segment.Control1Y, segment.Control2X, segment.Control2Y, segment.EndX, segment.EndY);
        }
        graphics.StrokeCurrentPath();

        if (!headEnd.IsNone)
        {
            FillCurvedConnectorEndMarker(
                graphics,
                headEnd,
                stroke,
                first.StartX,
                first.StartY,
                ResolveEndpointDirection(first.StartX, first.StartY, first.Control1X, first.Control1Y, first.Control2X, first.Control2Y),
                lineWidth);
        }

        if (!tailEnd.IsNone)
        {
            FillCurvedConnectorEndMarker(
                graphics,
                tailEnd,
                stroke,
                last.EndX,
                last.EndY,
                ResolveEndpointDirection(last.EndX, last.EndY, last.Control2X, last.Control2Y, last.Control1X, last.Control1Y),
                lineWidth);
        }
    }

    private static bool TryFillCurvedConnectorPreset(
        PdfGraphicsBuilder graphics,
        XElement shapeProperties,
        string preset,
        double x,
        double yTop,
        double width,
        double height,
        double slideHeight,
        double lineWidth,
        LineEndStyle tailEnd,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride)
    {
        List<BezierSegment> segments = preset switch
        {
            "curvedConnector2" => CreateCurvedConnector2Segments(x, yTop, width, height, slideHeight),
            "curvedConnector3" => CreateCurvedConnector3Segments(shapeProperties, presetAdjustmentsOverride, x, yTop, width, height, slideHeight),
            _ => []
        };
        if (segments.Count == 0)
        {
            return false;
        }

        return TryFillBezierConnectorPath(graphics, segments, lineWidth, tailEnd);
    }

    private static bool TryFillBezierConnectorPath(
        PdfGraphicsBuilder graphics,
        IReadOnlyList<BezierSegment> segments,
        double lineWidth,
        LineEndStyle tailEnd)
    {
        CurvedConnectorFillPath? path = BuildOfficeCurvedConnectorFillPath(segments, lineWidth, tailEnd);
        if (path is null)
        {
            return false;
        }

        LineEndKind tailKind = tailEnd.Kind;
        if (tailKind == LineEndKind.Arrow)
        {
            AppendClosedLinePath(graphics, path.Value.Points, explicitClosingLine: true);
            AppendOfficeArrowHeadPath(
                graphics,
                path.Value.TipX,
                path.Value.TipY,
                -path.Value.DirectionX,
                -path.Value.DirectionY,
                -path.Value.NormalX,
                -path.Value.NormalY,
                lineWidth,
                splitTrailingCurve: true);
            graphics.FillCurrentPath();
        }
        else
        {
            AppendClosedLinePath(graphics, path.Value.Points, explicitClosingLine: true);
            if (path.Value.TailSubpath is { } tailSubpath)
            {
                AppendClosedLinePath(graphics, tailSubpath, false);
            }

            graphics.FillCurrentPath();
        }

        return true;
    }

    private static CurvedConnectorFillPath? BuildOfficeCurvedConnectorFillPath(
        IReadOnlyList<BezierSegment> segments,
        double lineWidth,
        LineEndStyle tailEnd)
    {
        LineEndKind tailKind = tailEnd.Kind;
        int samplesPerSegment = tailKind switch
        {
            LineEndKind.Arrow => OfficeArrowTailConnectorSamplesPerSegment,
            LineEndKind.Stealth => OfficeStealthTailConnectorSamplesPerSegment,
            _ => OfficeTriangleTailConnectorSamplesPerSegment
        };
        List<CurveSample> samples = SampleBezierSegments(segments, samplesPerSegment);
        if (samples.Count < 2)
        {
            return null;
        }

        double[] cumulativeLengths = BuildCumulativeSampleLengths(samples, out double totalLength);
        double markerLength = tailKind switch
        {
            LineEndKind.Arrow => lineWidth * OfficeArrowheadLengthFactor * tailEnd.LengthScale,
            LineEndKind.Stealth => lineWidth * OfficeStraightStealthLineEndLengthFactor * tailEnd.LengthScale,
            _ => Math.Max(OfficeTriangleTailMinimumLength, lineWidth * OfficeTriangleTailLengthFactor) * tailEnd.LengthScale
        };
        double baseDistance = Math.Max(0d, totalLength - markerLength);
        CurveSample baseSample = SampleAtDistance(samples, cumulativeLengths, baseDistance);
        double halfWidth = Math.Max(0.1d, lineWidth / 2d);
        List<CurveSample> bodySamples = SelectBodySamples(samples, cumulativeLengths, baseDistance, baseSample);
        if (bodySamples.Count < 2)
        {
            return null;
        }

        (double X, double Y) direction = Normalize(baseSample.TangentX, baseSample.TangentY);
        if (Math.Abs(direction.X) <= 0.000001d && Math.Abs(direction.Y) <= 0.000001d)
        {
            return null;
        }

        (double X, double Y) normal = (-direction.Y, direction.X);
        (double X, double Y) tip = (samples[^1].X, samples[^1].Y);
        var points = new List<(double X, double Y)>(bodySamples.Count * 2);
        foreach (CurveSample sample in bodySamples)
        {
            (double X, double Y) sampleNormal = NormalForSample(sample);
            points.Add((sample.X + sampleNormal.X * halfWidth, sample.Y + sampleNormal.Y * halfWidth));
        }

        for (int i = bodySamples.Count - 1; i >= 0; i--)
        {
            CurveSample sample = bodySamples[i];
            (double X, double Y) sampleNormal = NormalForSample(sample);
            points.Add((sample.X - sampleNormal.X * halfWidth, sample.Y - sampleNormal.Y * halfWidth));
        }

        IReadOnlyList<(double X, double Y)>? tailSubpath = null;
        if (tailKind == LineEndKind.Stealth)
        {
            double markerWidth = lineWidth * OfficeStraightStealthLineEndWidthFactor * tailEnd.WidthScale;
            tailSubpath =
            [
                tip,
                (baseSample.X + normal.X * markerWidth / 2d, baseSample.Y + normal.Y * markerWidth / 2d),
                (tip.X - direction.X * markerLength * OfficeStraightStealthLineEndNotchFactor, tip.Y - direction.Y * markerLength * OfficeStraightStealthLineEndNotchFactor),
                (baseSample.X - normal.X * markerWidth / 2d, baseSample.Y - normal.Y * markerWidth / 2d)
            ];
        }
        else if (tailKind != LineEndKind.Arrow)
        {
            double arrowHalfWidth = markerLength *
                OfficeTriangleTailHalfWidthFactor *
                tailEnd.WidthScale;
            tailSubpath =
            [
                tip,
                (baseSample.X + normal.X * arrowHalfWidth, baseSample.Y + normal.Y * arrowHalfWidth),
                (baseSample.X - normal.X * arrowHalfWidth, baseSample.Y - normal.Y * arrowHalfWidth)
            ];
        }

        return new CurvedConnectorFillPath(points, tailSubpath, tip.X, tip.Y, direction.X, direction.Y, normal.X, normal.Y);
    }

    private static List<CurveSample> SampleBezierSegments(IReadOnlyList<BezierSegment> segments, int samplesPerSegment)
    {
        var samples = new List<CurveSample>(segments.Count * samplesPerSegment + 1);
        foreach (BezierSegment segment in segments)
        {
            int start = samples.Count == 0 ? 0 : 1;
            for (int i = start; i <= samplesPerSegment; i++)
            {
                samples.Add(SampleBezierSegment(segment, i / (double)samplesPerSegment));
            }
        }

        return samples;
    }

    private static double[] BuildCumulativeSampleLengths(IReadOnlyList<CurveSample> samples, out double totalLength)
    {
        totalLength = 0d;
        double[] cumulativeLengths = new double[samples.Count];
        for (int i = 1; i < samples.Count; i++)
        {
            totalLength += Distance(samples[i - 1].X, samples[i - 1].Y, samples[i].X, samples[i].Y);
            cumulativeLengths[i] = totalLength;
        }

        return cumulativeLengths;
    }

    private static List<CurveSample> SelectBodySamples(
        IReadOnlyList<CurveSample> samples,
        IReadOnlyList<double> cumulativeLengths,
        double baseDistance,
        CurveSample baseSample)
    {
        var bodySamples = new List<CurveSample>();
        for (int i = 0; i < samples.Count && cumulativeLengths[i] < baseDistance; i++)
        {
            bodySamples.Add(samples[i]);
        }

        bodySamples.Add(baseSample);
        return bodySamples;
    }

    private static void AppendClosedLinePath(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points, bool explicitClosingLine)
    {
        if (points.Count == 0)
        {
            return;
        }

        graphics.MoveTo(points[0].X, points[0].Y);
        for (int i = 1; i < points.Count; i++)
        {
            graphics.LineTo(points[i].X, points[i].Y);
        }

        if (explicitClosingLine)
        {
            graphics.LineTo(points[0].X, points[0].Y);
        }

        graphics.ClosePath();
    }

    private static CurveSample SampleBezierSegment(BezierSegment segment, double t)
    {
        double mt = 1d - t;
        double mt2 = mt * mt;
        double t2 = t * t;
        double x = mt2 * mt * segment.StartX +
            3d * mt2 * t * segment.Control1X +
            3d * mt * t2 * segment.Control2X +
            t2 * t * segment.EndX;
        double y = mt2 * mt * segment.StartY +
            3d * mt2 * t * segment.Control1Y +
            3d * mt * t2 * segment.Control2Y +
            t2 * t * segment.EndY;
        double dx = 3d * mt2 * (segment.Control1X - segment.StartX) +
            6d * mt * t * (segment.Control2X - segment.Control1X) +
            3d * t2 * (segment.EndX - segment.Control2X);
        double dy = 3d * mt2 * (segment.Control1Y - segment.StartY) +
            6d * mt * t * (segment.Control2Y - segment.Control1Y) +
            3d * t2 * (segment.EndY - segment.Control2Y);
        return new CurveSample(x, y, dx, dy);
    }

    private static CurveSample SampleAtDistance(IReadOnlyList<CurveSample> samples, IReadOnlyList<double> cumulativeLengths, double distance)
    {
        for (int i = 1; i < samples.Count; i++)
        {
            if (cumulativeLengths[i] < distance)
            {
                continue;
            }

            double previousLength = cumulativeLengths[i - 1];
            double segmentLength = cumulativeLengths[i] - previousLength;
            double t = segmentLength <= 0.000001d ? 0d : (distance - previousLength) / segmentLength;
            CurveSample a = samples[i - 1];
            CurveSample b = samples[i];
            return new CurveSample(
                Lerp(a.X, b.X, t),
                Lerp(a.Y, b.Y, t),
                Lerp(a.TangentX, b.TangentX, t),
                Lerp(a.TangentY, b.TangentY, t));
        }

        return samples[^1];
    }

    private static (double X, double Y) NormalForSample(CurveSample sample)
    {
        (double X, double Y) tangent = Normalize(sample.TangentX, sample.TangentY);
        return (-tangent.Y, tangent.X);
    }

    private static (double X, double Y) Normalize(double x, double y)
    {
        double length = Math.Sqrt(x * x + y * y);
        return length <= 0.000001d ? (0d, 0d) : (x / length, y / length);
    }

    private static double Distance(double x1, double y1, double x2, double y2)
    {
        double dx = x2 - x1;
        double dy = y2 - y1;
        return Math.Sqrt(dx * dx + dy * dy);
    }

    private static double Lerp(double a, double b, double t)
    {
        return a + (b - a) * t;
    }

    private static void FillCurvedConnectorEndMarker(
        PdfGraphicsBuilder graphics,
        LineEndStyle style,
        RgbColor stroke,
        double tipX,
        double tipY,
        (double X, double Y) awayDirection,
        double lineWidth)
    {
        if (style.Kind is LineEndKind.None)
        {
            return;
        }

        graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
        if (style.Kind == LineEndKind.Arrow)
        {
            double length = Math.Sqrt(awayDirection.X * awayDirection.X + awayDirection.Y * awayDirection.Y);
            if (length <= 0.001d)
            {
                return;
            }

            double ux = -awayDirection.X / length;
            double uy = -awayDirection.Y / length;
            double nx = -uy;
            double ny = ux;
            AppendOfficeArrowHeadPath(graphics, tipX, tipY, ux, uy, nx, ny, lineWidth, false);
            graphics.FillCurrentPath();
            return;
        }

        if (style.Kind == LineEndKind.Triangle)
        {
            FillLineArrowhead(graphics, tipX, tipY, awayDirection, lineWidth);
            return;
        }

        FillLineEndMarker(graphics, style, tipX, tipY, awayDirection.X, awayDirection.Y, lineWidth);
    }

    private static List<BezierSegment> CreateCurvedConnector2Segments(double x, double yTop, double width, double height, double slideHeight)
    {
        double y0 = slideHeight - yTop;
        double y1 = slideHeight - yTop - height;
        double k = 4d / 3d * (Math.Sqrt(2d) - 1d);
        return
        [
            new(
                x,
                y0,
                x + width * k,
                y0,
                x + width,
                y1 + height * k,
                x + width,
                y1)
        ];
    }

    private static List<BezierSegment> CreateCurvedConnector3Segments(
        XElement shapeProperties,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride,
        double x,
        double yTop,
        double width,
        double height,
        double slideHeight)
    {
        double adj1 = ReadPresetGeometryGuide(shapeProperties, presetAdjustmentsOverride, "adj1", 50000d) / 100000d;
        double x2 = width * adj1;
        double x1 = x2 / 2d;
        double x3 = (width + x2) / 2d;
        double vc = height / 2d;
        double hd4 = height / 4d;
        double y3 = height * 3d / 4d;
        return
        [
            ToSlideBezier(x, yTop, slideHeight, 0d, 0d, x1, 0d, x2, hd4, x2, vc),
            ToSlideBezier(x, yTop, slideHeight, x2, vc, x2, y3, x3, height, width, height)
        ];
    }

    private static BezierSegment ToSlideBezier(
        double x,
        double yTop,
        double slideHeight,
        double startX,
        double startY,
        double control1X,
        double control1Y,
        double control2X,
        double control2Y,
        double endX,
        double endY)
    {
        double MapY(double localY) => slideHeight - yTop - localY;
        return new BezierSegment(
            x + startX,
            MapY(startY),
            x + control1X,
            MapY(control1Y),
            x + control2X,
            MapY(control2Y),
            x + endX,
            MapY(endY));
    }

    private static (double X, double Y) ResolveEndpointDirection(double tipX, double tipY, double nearestControlX, double nearestControlY, double fallbackControlX, double fallbackControlY)
    {
        double dx = tipX - nearestControlX;
        double dy = tipY - nearestControlY;
        return dx * dx + dy * dy > 0.000001d
            ? (dx, dy)
            : (tipX - fallbackControlX, tipY - fallbackControlY);
    }

    private static LineEndStyle ReadLineEnd(XElement shapeProperties, string elementName)
    {
        XElement? end = shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + elementName);
        LineEndKind kind = ReadLineEndKind((string?)end?.Attribute("type"));
        return new LineEndStyle(
            kind,
            ReadLineEndScale((string?)end?.Attribute("w")),
            ReadLineEndScale((string?)end?.Attribute("len")));
    }

    private static LineEndStyle ToLineEndStyle(PptxSceneLineEnd lineEnd)
    {
        return new LineEndStyle(
            lineEnd.Kind switch
            {
                PptxSceneLineEndKind.Triangle => LineEndKind.Triangle,
                PptxSceneLineEndKind.Arrow => LineEndKind.Arrow,
                PptxSceneLineEndKind.Stealth => LineEndKind.Stealth,
                PptxSceneLineEndKind.Diamond => LineEndKind.Diamond,
                PptxSceneLineEndKind.Oval => LineEndKind.Oval,
                _ => LineEndKind.None
            },
            lineEnd.WidthScale,
            lineEnd.LengthScale);
    }

    private static LineStyle ToLineStyle(PptxSceneLineStyle line)
    {
        return new LineStyle(line.HasLine, line.Color, line.Width, line.Alpha, line.DashPattern ?? [], line.Cap, line.Join);
    }

    private static FillStyle ToFillStyle(PptxSceneFillStyle fill)
    {
        return new FillStyle(fill.HasFill, fill.Color, fill.Alpha);
    }
}
