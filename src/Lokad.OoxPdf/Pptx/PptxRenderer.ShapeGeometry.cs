using System.Globalization;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool TryRenderCustomGeometry(
        PptxSceneCustomGeometry customGeometry,
        PdfGraphicsBuilder graphics,
        double x,
        double y,
        double width,
        double height,
        bool hasFill,
        RgbColor fill,
        double fillAlpha,
        bool hasStroke,
        RgbColor stroke,
        double lineWidth,
        double strokeAlpha,
        bool hasDash,
        IReadOnlyList<double> dashPattern,
        int? lineCap,
        int? lineJoin,
        LineEndStyle headEnd,
        LineEndStyle tailEnd)
    {
        if (!customGeometry.HasGeometry || customGeometry.Paths.Count == 0)
        {
            return false;
        }

        if (hasFill)
        {
            bool transparentFill = fillAlpha < 0.999d;
            if (transparentFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(fillAlpha, 1d);
            }

            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
            foreach (PptxSceneCustomPath path in customGeometry.Paths.Where(path => path.AllowsFill))
            {
                AppendCustomGeometryPath(graphics, customGeometry, path, x, y, width, height);
                graphics.FillCurrentPathEvenOdd();
            }

            if (transparentFill)
            {
                graphics.RestoreState();
            }
        }

        if (hasStroke)
        {
            bool transparentStroke = strokeAlpha < 0.999d;
            if (transparentStroke)
            {
                graphics.SaveState();
                graphics.SetAlpha(strokeAlpha, strokeAlpha);
            }

            graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
            graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
            graphics.SetLineWidth(lineWidth);
            if (hasDash)
            {
                graphics.SetLineDash(dashPattern);
            }

            int? appliedLineJoin = lineJoin ?? (lineCap is null ? null : 1);
            if (lineCap is { } cap)
            {
                graphics.SetLineCap(cap);
            }

            if (appliedLineJoin is { } join)
            {
                graphics.SetLineJoin(join);
            }

            foreach (PptxSceneCustomPath path in customGeometry.Paths.Where(path => path.AllowsStroke))
            {
                if (TryFillCustomGeometryOpenLineEndPath(
                    graphics,
                    customGeometry,
                    path,
                    x,
                    y,
                    width,
                    height,
                    hasFill,
                    lineWidth,
                    hasDash,
                    lineCap,
                    lineJoin,
                    headEnd,
                    tailEnd))
                {
                    continue;
                }

                AppendCustomGeometryPath(graphics, customGeometry, path, x, y, width, height);
                graphics.StrokeCurrentPath();
            }

            if (hasDash)
            {
                graphics.ClearLineDash();
            }

            if (lineCap is not null)
            {
                graphics.SetLineCap(0);
            }

            if (appliedLineJoin is not null)
            {
                graphics.SetLineJoin(0);
            }

            if (transparentStroke)
            {
                graphics.RestoreState();
            }
        }

        return true;
    }

    private static bool TryFillCustomGeometryOpenLineEndPath(
        PdfGraphicsBuilder graphics,
        PptxSceneCustomGeometry geometry,
        PptxSceneCustomPath path,
        double x,
        double y,
        double width,
        double height,
        bool hasFill,
        double lineWidth,
        bool hasDash,
        int? lineCap,
        int? lineJoin,
        LineEndStyle headEnd,
        LineEndStyle tailEnd)
    {
        if (hasDash ||
            lineCap is not null ||
            lineJoin is not null ||
            !headEnd.IsNone ||
            tailEnd.Kind is not (LineEndKind.Triangle or LineEndKind.Arrow or LineEndKind.Stealth) ||
            hasFill && path.AllowsFill)
        {
            return false;
        }

        if (!TryCreateCustomGeometryBezierSegments(geometry, path, x, y, width, height, out List<BezierSegment> segments))
        {
            return false;
        }

        return TryFillBezierConnectorPath(graphics, segments, lineWidth, tailEnd);
    }

    private static bool TryCreateCustomGeometryBezierSegments(
        PptxSceneCustomGeometry geometry,
        PptxSceneCustomPath path,
        double x,
        double y,
        double width,
        double height,
        out List<BezierSegment> segments)
    {
        double coordinateWidth = Math.Max(1d, path.Width);
        double coordinateHeight = Math.Max(1d, path.Height);
        IReadOnlyDictionary<string, double> guides = BuildCustomGeometryGuides(
            geometry.Guides,
            coordinateWidth,
            coordinateHeight);
        segments = [];
        (double X, double Y)? current = null;
        bool hasMove = false;

        foreach (PptxSceneCustomCommand command in path.Commands)
        {
            if (command.Kind == PptxSceneCustomCommandKind.MoveTo)
            {
                if (hasMove)
                {
                    return false;
                }

                current = ReadCustomGeometryPoint(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                hasMove = true;
            }
            else if (command.Kind == PptxSceneCustomCommandKind.LineTo && current is { } lineStart)
            {
                (double X, double Y) end = ReadCustomGeometryPoint(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                segments.Add(CreateLineBezierSegment(lineStart, end));
                current = end;
            }
            else if (command.Kind == PptxSceneCustomCommandKind.CubicBezierTo && current is { } cubicStart)
            {
                List<(double X, double Y)> points = ReadCustomGeometryPoints(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                if (points.Count != 3)
                {
                    return false;
                }

                segments.Add(new BezierSegment(cubicStart.X, cubicStart.Y, points[0].X, points[0].Y, points[1].X, points[1].Y, points[2].X, points[2].Y));
                current = points[2];
            }
            else if (command.Kind == PptxSceneCustomCommandKind.QuadraticBezierTo && current is { } quadStart)
            {
                List<(double X, double Y)> points = ReadCustomGeometryPoints(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                if (points.Count != 2)
                {
                    return false;
                }

                segments.Add(CreateQuadraticBezierSegment(quadStart, points[0], points[1]));
                current = points[1];
            }
            else
            {
                return false;
            }
        }

        return hasMove && segments.Count > 0;
    }

    private static BezierSegment CreateLineBezierSegment((double X, double Y) start, (double X, double Y) end)
    {
        return new BezierSegment(
            start.X,
            start.Y,
            start.X + (end.X - start.X) / 3d,
            start.Y + (end.Y - start.Y) / 3d,
            start.X + (end.X - start.X) * 2d / 3d,
            start.Y + (end.Y - start.Y) * 2d / 3d,
            end.X,
            end.Y);
    }

    private static BezierSegment CreateQuadraticBezierSegment((double X, double Y) start, (double X, double Y) control, (double X, double Y) end)
    {
        return new BezierSegment(
            start.X,
            start.Y,
            start.X + (2d / 3d) * (control.X - start.X),
            start.Y + (2d / 3d) * (control.Y - start.Y),
            end.X + (2d / 3d) * (control.X - end.X),
            end.Y + (2d / 3d) * (control.Y - end.Y),
            end.X,
            end.Y);
    }

    private static void AppendCustomGeometryPath(
        PdfGraphicsBuilder graphics,
        PptxSceneCustomGeometry geometry,
        PptxSceneCustomPath path,
        double x,
        double y,
        double width,
        double height)
    {
        double coordinateWidth = Math.Max(1d, path.Width);
        double coordinateHeight = Math.Max(1d, path.Height);
        IReadOnlyDictionary<string, double> guides = BuildCustomGeometryGuides(
            geometry.Guides,
            coordinateWidth,
            coordinateHeight);
        (double X, double Y)? current = null;

        foreach (PptxSceneCustomCommand command in path.Commands)
        {
            if (command.Kind == PptxSceneCustomCommandKind.MoveTo)
            {
                current = ReadCustomGeometryPoint(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                graphics.MoveTo(current.Value.X, current.Value.Y);
            }
            else if (command.Kind == PptxSceneCustomCommandKind.LineTo)
            {
                current = ReadCustomGeometryPoint(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                graphics.LineTo(current.Value.X, current.Value.Y);
            }
            else if (command.Kind == PptxSceneCustomCommandKind.CubicBezierTo)
            {
                List<(double X, double Y)> points = ReadCustomGeometryPoints(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                if (points.Count == 3)
                {
                    graphics.CurveTo(points[0].X, points[0].Y, points[1].X, points[1].Y, points[2].X, points[2].Y);
                    current = points[2];
                }
            }
            else if (command.Kind == PptxSceneCustomCommandKind.QuadraticBezierTo)
            {
                List<(double X, double Y)> points = ReadCustomGeometryPoints(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                if (points.Count == 2 && current is { } start)
                {
                    (double X, double Y) control = points[0];
                    (double X, double Y) end = points[1];
                    graphics.CurveTo(
                        start.X + (2d / 3d) * (control.X - start.X),
                        start.Y + (2d / 3d) * (control.Y - start.Y),
                        end.X + (2d / 3d) * (control.X - end.X),
                        end.Y + (2d / 3d) * (control.Y - end.Y),
                        end.X,
                        end.Y);
                    current = end;
                }
            }
            else if (command.Kind == PptxSceneCustomCommandKind.ArcTo && current is { } arcStart)
            {
                current = AppendCustomGeometryArc(graphics, command, arcStart, x, y, width, height, coordinateWidth, coordinateHeight, guides);
            }
            else if (command.Kind == PptxSceneCustomCommandKind.Close)
            {
                graphics.ClosePath();
            }
        }
    }

    private static (double X, double Y) ReadCustomGeometryPoint(
        PptxSceneCustomCommand command,
        double x,
        double y,
        double width,
        double height,
        double coordinateWidth,
        double coordinateHeight,
        IReadOnlyDictionary<string, double> guides)
    {
        return command.Points.Count == 0
            ? (x, y + height)
            : MapCustomGeometryPoint(command.Points[0], x, y, width, height, coordinateWidth, coordinateHeight, guides);
    }

    private static List<(double X, double Y)> ReadCustomGeometryPoints(
        PptxSceneCustomCommand command,
        double x,
        double y,
        double width,
        double height,
        double coordinateWidth,
        double coordinateHeight,
        IReadOnlyDictionary<string, double> guides)
    {
        return command.Points
            .Select(point => MapCustomGeometryPoint(point, x, y, width, height, coordinateWidth, coordinateHeight, guides))
            .ToList();
    }

    private static (double X, double Y) MapCustomGeometryPoint(
        PptxSceneCustomPoint point,
        double x,
        double y,
        double width,
        double height,
        double coordinateWidth,
        double coordinateHeight,
        IReadOnlyDictionary<string, double> guides)
    {
        double pointX = ReadCustomGeometryValue(point.X, guides, 0d);
        double pointY = ReadCustomGeometryValue(point.Y, guides, 0d);
        return (x + pointX / coordinateWidth * width, y + height - pointY / coordinateHeight * height);
    }

    private static (double X, double Y) AppendCustomGeometryArc(
        PdfGraphicsBuilder graphics,
        PptxSceneCustomCommand arc,
        (double X, double Y) start,
        double x,
        double y,
        double width,
        double height,
        double coordinateWidth,
        double coordinateHeight,
        IReadOnlyDictionary<string, double> guides)
    {
        double radiusX = Math.Abs(ReadCustomGeometryValue(arc.RadiusX, guides, 0d) / coordinateWidth * width);
        double radiusY = Math.Abs(ReadCustomGeometryValue(arc.RadiusY, guides, 0d) / coordinateHeight * height);
        double startAngle = DegreesToRadians(ReadCustomGeometryValue(arc.StartAngle, guides, 0d) / 60000d);
        double sweepAngle = DegreesToRadians(ReadCustomGeometryValue(arc.SweepAngle, guides, 0d) / 60000d);
        if (radiusX <= PptxTextMetricRules.TextStateTolerance ||
            radiusY <= PptxTextMetricRules.TextStateTolerance ||
            Math.Abs(sweepAngle) <= PptxTextMetricRules.TextStateTolerance)
        {
            return start;
        }

        double centerX = start.X - radiusX * Math.Cos(startAngle);
        double centerY = start.Y + radiusY * Math.Sin(startAngle);
        int segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(sweepAngle) / (Math.PI / 2d)));
        double delta = sweepAngle / segments;
        double angle = startAngle;
        (double X, double Y) current = start;
        for (int i = 0; i < segments; i++)
        {
            double nextAngle = angle + delta;
            double k = 4d / 3d * Math.Tan(delta / 4d);
            (double X, double Y) end = (
                centerX + radiusX * Math.Cos(nextAngle),
                centerY - radiusY * Math.Sin(nextAngle));
            (double X, double Y) control1 = (
                current.X + k * -radiusX * Math.Sin(angle),
                current.Y + k * -radiusY * Math.Cos(angle));
            (double X, double Y) control2 = (
                end.X - k * -radiusX * Math.Sin(nextAngle),
                end.Y - k * -radiusY * Math.Cos(nextAngle));
            graphics.CurveTo(control1.X, control1.Y, control2.X, control2.Y, end.X, end.Y);
            current = end;
            angle = nextAngle;
        }

        return current;
    }

    private static IReadOnlyDictionary<string, double> BuildCustomGeometryGuides(IReadOnlyList<PptxSceneCustomGuide> customGuides, double width, double height)
    {
        var guides = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
        {
            ["w"] = width,
            ["h"] = height,
            ["ss"] = Math.Min(width, height),
            ["ls"] = Math.Max(width, height)
        };

        foreach (PptxSceneCustomGuide guide in customGuides)
        {
            if (!string.IsNullOrWhiteSpace(guide.Name) && !string.IsNullOrWhiteSpace(guide.Formula))
            {
                guides[guide.Name] = EvaluateCustomGeometryFormula(guide.Formula, guides);
            }
        }

        return guides;
    }

    private static double EvaluateCustomGeometryFormula(string formula, IReadOnlyDictionary<string, double> guides)
    {
        string[] parts = formula.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0)
        {
            return 0d;
        }

        return parts[0] switch
        {
            "val" when parts.Length >= 2 => ReadCustomGeometryValue(parts[1], guides, 0d),
            "+-" when parts.Length >= 4 => ReadCustomGeometryValue(parts[1], guides, 0d) +
                ReadCustomGeometryValue(parts[2], guides, 0d) -
                ReadCustomGeometryValue(parts[3], guides, 0d),
            "*/" when parts.Length >= 4 => ReadCustomGeometryValue(parts[3], guides, 0d) == 0d
                ? 0d
                : ReadCustomGeometryValue(parts[1], guides, 0d) *
                    ReadCustomGeometryValue(parts[2], guides, 0d) /
                    ReadCustomGeometryValue(parts[3], guides, 0d),
            "abs" when parts.Length >= 2 => Math.Abs(ReadCustomGeometryValue(parts[1], guides, 0d)),
            "min" when parts.Length >= 3 => Math.Min(ReadCustomGeometryValue(parts[1], guides, 0d), ReadCustomGeometryValue(parts[2], guides, 0d)),
            "max" when parts.Length >= 3 => Math.Max(ReadCustomGeometryValue(parts[1], guides, 0d), ReadCustomGeometryValue(parts[2], guides, 0d)),
            "?:" when parts.Length >= 4 => ReadCustomGeometryValue(parts[1], guides, 0d) > 0d
                ? ReadCustomGeometryValue(parts[2], guides, 0d)
                : ReadCustomGeometryValue(parts[3], guides, 0d),
            "sin" when parts.Length >= 3 => ReadCustomGeometryValue(parts[1], guides, 0d) *
                Math.Sin(DegreesToRadians(ReadCustomGeometryValue(parts[2], guides, 0d) / 60000d)),
            "cos" when parts.Length >= 3 => ReadCustomGeometryValue(parts[1], guides, 0d) *
                Math.Cos(DegreesToRadians(ReadCustomGeometryValue(parts[2], guides, 0d) / 60000d)),
            _ => 0d
        };
    }

    private static double ReadCustomGeometryValue(string? value, IReadOnlyDictionary<string, double> guides, double defaultValue)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return defaultValue;
        }

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed))
        {
            return parsed;
        }

        return guides.TryGetValue(value, out double guide) ? guide : defaultValue;
    }

    private static double DegreesToRadians(double degrees)
    {
        return degrees * Math.PI / 180d;
    }
}
