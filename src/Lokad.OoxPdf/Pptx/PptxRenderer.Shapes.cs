using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static void RenderBackground(PptxRenderContext context, XDocument slideXml, PdfGraphicsBuilder graphics)
    {
        XElement? background = slideXml.Root?
            .Element(PresentationNamespace + "cSld")?
            .Element(PresentationNamespace + "bg")?
            .Element(PresentationNamespace + "bgPr");
        if (TryReadSolidColor(background, context.Theme, out RgbColor color))
        {
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            graphics.FillRectangle(0, 0, context.Document.SlideWidthPoints, context.Document.SlideHeightPoints);
        }
    }

    private static void RenderShapes(PptxRenderContext context, XDocument slideXml, PdfGraphicsBuilder graphics, bool renderPlaceholders)
    {
        foreach (XElement shapeTree in slideXml.Descendants(PresentationNamespace + "spTree"))
        {
            RenderShapeContainer(context, shapeTree, graphics, GroupTransform.Identity, renderPlaceholders);
        }
    }

    private static void RenderShapeContainer(
        PptxRenderContext context,
        XElement container,
        PdfGraphicsBuilder graphics,
        GroupTransform transform,
        bool renderPlaceholders)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == PresentationNamespace + "sp")
            {
                if (renderPlaceholders || !IsPlaceholder(child))
                {
                    RenderShape(child, context.Document, graphics, context.Theme, transform);
                }

                continue;
            }

            if (child.Name == PresentationNamespace + "cxnSp")
            {
                RenderShape(child, context.Document, graphics, context.Theme, transform);
                continue;
            }

            if (child.Name == PresentationNamespace + "grpSp")
            {
                GroupTransform childTransform = transform.Combine(ReadGroupTransform(child));
                RenderShapeContainer(context, child, graphics, childTransform, renderPlaceholders);
            }
        }
    }

    private static void RenderShape(
        XElement shape,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        PptxTheme theme,
        GroupTransform groupTransform)
    {
        int imageIndex = 1;
        RenderShape(
            shape,
            relationships: null,
            package: null,
            document,
            graphics,
            diagnosticSink: null,
            slideIndex: 0,
            theme,
            groupTransform,
            images: null,
            imageCache: null,
            ref imageIndex);
    }
    private static void RenderShape(
        XElement shape,
        IReadOnlyDictionary<string, OoxRelationship>? relationships,
        OoxPackage? package,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        PptxTheme theme,
        GroupTransform groupTransform,
        List<PdfImageResource>? images,
        Dictionary<string, PdfImageXObject?>? imageCache,
        ref int imageIndex)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        if (shapeProperties is null)
        {
            return;
        }

        ShapeBounds? rawBounds = ReadBounds(shapeProperties);
        if (rawBounds is null)
        {
            return;
        }

        ShapeBounds bounds = groupTransform.Apply(rawBounds.Value);
        string preset = ReadPreset(shapeProperties);

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        bool transformed = bounds.RotationDegrees != 0d || bounds.FlipHorizontal || bounds.FlipVertical;

        bool hasFill = TryReadShapeFill(shape, shapeProperties, theme, out RgbColor fill, out double fillAlpha);
        bool hasPatternFill = TryReadShapePatternFill(shapeProperties, theme, out ShapePatternFill patternFill);
        bool hasStroke = TryReadShapeLine(shape, shapeProperties, theme, out RgbColor stroke, out double lineWidth, out double strokeAlpha);
        bool hasDash = TryReadPresetDash(shapeProperties, lineWidth, out IReadOnlyList<double> dashPattern);
        int? lineCap = ReadLineCap(shapeProperties) switch
        {
            "rnd" => 1,
            "sq" => 2,
            _ => null
        };
        int? lineJoin = ReadLineJoin(shapeProperties);
        bool hasPictureFill = TryReadShapePictureFill(
            shapeProperties,
            relationships,
            package,
            diagnosticSink,
            slideIndex,
            images,
            imageCache,
            ref imageIndex,
            out string? pictureFillName,
            out PdfImageXObject? pictureFillImage);
        XElement? customGeometry = shapeProperties.Element(DrawingNamespace + "custGeom");

        if (transformed)
        {
            graphics.SaveState();
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }

        if (TryReadGlow(shapeProperties, theme, out Glow glow) &&
            preset is not ("line" or "straightConnector1" or "curvedConnector2" or "curvedConnector3") &&
            customGeometry is null)
        {
            DrawGlow(graphics, preset, x, y, width, height, glow);
        }

        if (TryReadOuterShadow(shapeProperties, theme, out OuterShadow outerShadow) &&
            preset is not ("line" or "straightConnector1" or "curvedConnector2" or "curvedConnector3") &&
            customGeometry is null)
        {
            DrawOuterShadow(graphics, preset, x, y, width, height, outerShadow);
        }

        if (customGeometry is not null && TryRenderCustomGeometry(
                customGeometry,
                graphics,
                x,
                y,
                width,
                height,
                hasFill,
                fill,
                fillAlpha,
                hasStroke,
                stroke,
                lineWidth,
                strokeAlpha,
                hasDash,
                dashPattern,
                lineCap,
                lineJoin))
        {
            if (transformed)
            {
                graphics.RestoreState();
            }

            return;
        }

        if (preset is "line" or "straightConnector1")
        {
            if (hasStroke)
            {
                bool transparentStroke = strokeAlpha < 0.999d;
                if (transparentStroke)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(strokeAlpha, strokeAlpha);
                }

                double x1 = x;
                double y1 = document.SlideHeightPoints - yTop;
                double x2 = x + width;
                double y2 = document.SlideHeightPoints - yTop - height;
                LineEndStyle headEnd = ReadLineEnd(shapeProperties, "headEnd");
                LineEndStyle tailEnd = ReadLineEnd(shapeProperties, "tailEnd");
                bool hasHeadArrow = IsFilledTriangleArrow(headEnd);
                bool hasTailArrow = IsFilledTriangleArrow(tailEnd);
                if ((hasHeadArrow || hasTailArrow) && headEnd.Kind is LineEndKind.None or LineEndKind.Triangle or LineEndKind.Arrow && tailEnd.Kind is LineEndKind.None or LineEndKind.Triangle or LineEndKind.Arrow && !hasDash && lineCap is null)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    bool usesOfficeArrowType = headEnd.Kind == LineEndKind.Arrow || tailEnd.Kind == LineEndKind.Arrow;
                    if (usesOfficeArrowType)
                    {
                        FillOfficeArrowedLine(graphics, x1, y1, x2, y2, lineWidth, hasHeadArrow, hasTailArrow);
                    }
                    else
                    {
                        FillArrowedLine(graphics, x1, y1, x2, y2, lineWidth, hasHeadArrow, hasTailArrow);
                    }
                }
                else
                {
                    graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
                    graphics.SetLineWidth(lineWidth);
                    if (hasDash)
                    {
                        graphics.SetLineDash(dashPattern);
                    }

                    if (lineCap is { } cap)
                    {
                        graphics.SetLineCap(cap);
                        graphics.SetLineJoin(1);
                    }

                    graphics.StrokeLine(x1, y1, x2, y2);
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    FillLineEndMarker(graphics, headEnd, x1, y1, x1 - x2, y1 - y2, lineWidth);
                    FillLineEndMarker(graphics, tailEnd, x2, y2, x2 - x1, y2 - y1, lineWidth);
                    if (hasDash)
                    {
                        graphics.ClearLineDash();
                    }

                    if (lineCap is not null)
                    {
                        graphics.SetLineCap(0);
                        graphics.SetLineJoin(0);
                    }
                }

                if (transparentStroke)
                {
                    graphics.RestoreState();
                }
            }

            if (transformed)
            {
                graphics.RestoreState();
            }

            return;
        }

        if (preset is "curvedConnector2" or "curvedConnector3")
        {
            if (hasStroke)
            {
                bool transparentStroke = strokeAlpha < 0.999d;
                if (transparentStroke)
                {
                    graphics.SaveState();
                    graphics.SetAlpha(strokeAlpha, strokeAlpha);
                }

                graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
                graphics.SetLineWidth(lineWidth);
                if (hasDash)
                {
                    graphics.SetLineDash(dashPattern);
                }

                if (lineCap is { } cap)
                {
                    graphics.SetLineCap(cap);
                    graphics.SetLineJoin(1);
                }

                bool hasHeadArrow = IsFilledTriangleArrow(ReadLineEnd(shapeProperties, "headEnd"));
                bool hasTailArrow = IsFilledTriangleArrow(ReadLineEnd(shapeProperties, "tailEnd"));
                DrawCurvedConnector3(graphics, x, yTop, width, height, document.SlideHeightPoints, stroke, lineWidth, hasHeadArrow, hasTailArrow);

                if (hasDash)
                {
                    graphics.ClearLineDash();
                }

                if (lineCap is not null)
                {
                    graphics.SetLineCap(0);
                    graphics.SetLineJoin(0);
                }

                if (transparentStroke)
                {
                    graphics.RestoreState();
                }
            }

            if (transformed)
            {
                graphics.RestoreState();
            }

            return;
        }

        if (hasPictureFill && pictureFillName is not null && pictureFillImage is not null)
        {
            CropRect crop = ReadCrop(shapeProperties);
            FillRect fillRect = ReadFillRect(shapeProperties);
            double imageX = x + fillRect.Left * width;
            double imageY = y + fillRect.Bottom * height;
            double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
            double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));
            bool clippedToShape = preset != "rect";
            if (clippedToShape)
            {
                graphics.SaveState();
                ClipToPresetShape(graphics, preset, x, y, width, height);
            }

            DrawImageFill(graphics, pictureFillName, imageX, imageY, imageWidth, imageHeight, crop);

            if (clippedToShape)
            {
                graphics.RestoreState();
            }

            images?.Add(new PdfImageResource(pictureFillName, pictureFillImage));
        }
        else if (hasPatternFill)
        {
            bool transparentFill = patternFill.Alpha < 0.999d;
            if (transparentFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(patternFill.Alpha, 1d);
            }

            graphics.SetFillRgb(patternFill.Background.Red, patternFill.Background.Green, patternFill.Background.Blue);
            DrawPresetFill(graphics, preset, x, y, width, height);
            StrokeShapePatternFill(graphics, preset, x, y, width, height, patternFill);

            if (transparentFill)
            {
                graphics.RestoreState();
            }
        }
        else if (hasFill)
        {
            bool transparentFill = fillAlpha < 0.999d;
            if (transparentFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(fillAlpha, 1d);
            }

            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
            DrawPresetFill(graphics, preset, x, y, width, height);

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
                graphics.SetAlpha(1d, strokeAlpha);
            }

            graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
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

            if (preset == "arc")
            {
                DrawPresetArcStroke(graphics, shapeProperties, x, y, width, height);
            }
            else if (preset == "ellipse")
            {
                graphics.StrokeEllipse(x, y, width, height);
            }
            else if (preset == "roundRect")
            {
                graphics.StrokeRoundedRectangle(x, y, width, height, Math.Min(width, height) * 0.16d);
            }
            else if (TryCreatePresetPolygonPoints(preset, x, y, width, height, out (double X, double Y)[] polygonPoints))
            {
                graphics.StrokePolygon(polygonPoints);
            }
            else
            {
                graphics.StrokeRectangle(x, y, width, height);
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

        if (transformed)
        {
            graphics.RestoreState();
        }
    }

    private static bool HasUnsupportedPictureFill(XElement shapeProperties)
    {
        if (shapeProperties.Element(DrawingNamespace + "blipFill") is null)
        {
            return false;
        }

        return !CanRenderPictureFillPreset(ReadPreset(shapeProperties));
    }

    private static bool TryReadGlow(XElement shapeProperties, PptxTheme theme, out Glow glow)
    {
        XElement? glowElement = shapeProperties
            .Element(DrawingNamespace + "effectLst")
            ?.Element(DrawingNamespace + "glow");
        if (glowElement is null)
        {
            glow = default;
            return false;
        }

        XElement? colorElement = glowElement.Elements().FirstOrDefault(element =>
            element.Name.LocalName is "srgbClr" or "schemeClr" or "prstClr");
        if (colorElement is not null &&
            TryReadImageRecolorColor(colorElement, theme, out RgbColor color))
        {
            double radius = OoxUnits.EmuToPoints(ParseOptionalLongAttribute(glowElement, "rad", 0));
            glow = new Glow(
                color,
                ReadAlpha(new XElement(DrawingNamespace + "solidFill", new XElement(colorElement))),
                radius);
            return radius > PptxTextMetricRules.TextStateTolerance;
        }

        glow = default;
        return false;
    }

    private static bool TryReadOuterShadow(XElement shapeProperties, PptxTheme theme, out OuterShadow shadow)
    {
        XElement? outerShadow = shapeProperties
            .Element(DrawingNamespace + "effectLst")
            ?.Element(DrawingNamespace + "outerShdw");
        if (outerShadow is null)
        {
            shadow = default;
            return false;
        }

        XElement? colorElement = outerShadow.Elements().FirstOrDefault(element =>
            element.Name.LocalName is "srgbClr" or "schemeClr" or "prstClr");
        if (colorElement is not null &&
            TryReadImageRecolorColor(colorElement, theme, out RgbColor color))
        {
            double alpha = ReadAlpha(new XElement(DrawingNamespace + "solidFill", new XElement(colorElement)));
            double distance = OoxUnits.EmuToPoints(ParseOptionalLongAttribute(outerShadow, "dist", 0));
            double direction = ParseOptionalLongAttribute(outerShadow, "dir", 0) / 60000d * Math.PI / 180d;
            shadow = new OuterShadow(
                color,
                alpha,
                distance * Math.Cos(direction),
                -distance * Math.Sin(direction));
            return true;
        }

        shadow = default;
        return false;
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
        DrawPresetFill(graphics, preset, x - glow.Radius, y - glow.Radius, width + 2d * glow.Radius, height + 2d * glow.Radius);
        graphics.RestoreState();
    }

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
        DrawPresetFill(graphics, preset, x + shadow.OffsetX, y + shadow.OffsetY, width, height);
        graphics.RestoreState();
    }

    private static void DrawPresetFill(PdfGraphicsBuilder graphics, string preset, double x, double y, double width, double height)
    {
        if (preset == "ellipse")
        {
            graphics.FillEllipse(x, y, width, height);
        }
        else if (preset == "roundRect")
        {
            graphics.FillRoundedRectangle(x, y, width, height, Math.Min(width, height) * 0.16d);
        }
        else if (TryCreatePresetPolygonPoints(preset, x, y, width, height, out (double X, double Y)[] polygonPoints))
        {
            graphics.FillPolygon(polygonPoints);
        }
        else
        {
            graphics.FillRectangle(x, y, width, height);
        }
    }

    private static void DrawPresetArcStroke(PdfGraphicsBuilder graphics, XElement shapeProperties, double x, double y, double width, double height)
    {
        double startDegrees = ReadPresetGeometryGuide(shapeProperties, "adj1", 0d) / 60000d;
        double endDegrees = ReadPresetGeometryGuide(shapeProperties, "adj2", 90d) / 60000d;
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

    private static double ReadPresetGeometryGuide(XElement shapeProperties, string name, double fallback)
    {
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

    private static void AppendEllipseArcSegment(PdfGraphicsBuilder graphics, double centerX, double centerY, double radiusX, double radiusY, double startDegrees, double sweepDegrees, bool moveToStart)
    {
        double start = DegreesToRadians(startDegrees);
        double sweep = DegreesToRadians(sweepDegrees);
        double end = start + sweep;
        double k = 4d / 3d * Math.Tan(sweep / 4d);

        double x0 = centerX + radiusX * Math.Cos(start);
        double y0 = centerY + radiusY * Math.Sin(start);
        double x3 = centerX + radiusX * Math.Cos(end);
        double y3 = centerY + radiusY * Math.Sin(end);
        double x1 = x0 - radiusX * k * Math.Sin(start);
        double y1 = y0 + radiusY * k * Math.Cos(start);
        double x2 = x3 + radiusX * k * Math.Sin(end);
        double y2 = y3 - radiusY * k * Math.Cos(end);

        if (moveToStart)
        {
            graphics.MoveTo(x0, y0);
        }

        graphics.CurveTo(x1, y1, x2, y2, x3, y3);
    }

    private static string ReadPreset(XElement shapeProperties)
    {
        return (string?)shapeProperties
            .Element(DrawingNamespace + "prstGeom")
            ?.Attribute("prst") ?? "rect";
    }

    private static bool TryRenderCustomGeometry(
        XElement customGeometry,
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
        int? lineJoin)
    {
        List<XElement> paths = customGeometry
            .Element(DrawingNamespace + "pathLst")
            ?.Elements(DrawingNamespace + "path")
            .Where(CanRenderCustomGeometryPath)
            .ToList() ?? [];
        if (paths.Count == 0)
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
            foreach (XElement path in paths.Where(CustomGeometryPathAllowsFill))
            {
                AppendCustomGeometryPath(graphics, path, x, y, width, height);
                graphics.FillCurrentPath();
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
                graphics.SetAlpha(1d, strokeAlpha);
            }

            graphics.SetStrokeRgb(stroke.Red, stroke.Green, stroke.Blue);
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

            foreach (XElement path in paths.Where(CustomGeometryPathAllowsStroke))
            {
                AppendCustomGeometryPath(graphics, path, x, y, width, height);
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

    private static bool CanRenderCustomGeometryPath(XElement path)
    {
        return path.Elements().Any() && path.Elements().All(child =>
            child.Name == DrawingNamespace + "moveTo" ||
            child.Name == DrawingNamespace + "lnTo" ||
            child.Name == DrawingNamespace + "cubicBezTo" ||
            child.Name == DrawingNamespace + "quadBezTo" ||
            child.Name == DrawingNamespace + "arcTo" ||
            child.Name == DrawingNamespace + "close");
    }

    private static bool CustomGeometryPathAllowsFill(XElement path)
    {
        string? fill = (string?)path.Attribute("fill");
        return !string.Equals(fill, "none", StringComparison.Ordinal);
    }

    private static bool CustomGeometryPathAllowsStroke(XElement path)
    {
        return !string.Equals((string?)path.Attribute("stroke"), "false", StringComparison.Ordinal);
    }

    private static void AppendCustomGeometryPath(PdfGraphicsBuilder graphics, XElement path, double x, double y, double width, double height)
    {
        double coordinateWidth = Math.Max(1d, ParseOptionalDoubleAttribute(path, "w", 21600d));
        double coordinateHeight = Math.Max(1d, ParseOptionalDoubleAttribute(path, "h", 21600d));
        IReadOnlyDictionary<string, double> guides = BuildCustomGeometryGuides(
            path.Parent?.Parent,
            coordinateWidth,
            coordinateHeight);
        (double X, double Y)? current = null;

        foreach (XElement command in path.Elements())
        {
            if (command.Name == DrawingNamespace + "moveTo")
            {
                current = ReadCustomGeometryPoint(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                graphics.MoveTo(current.Value.X, current.Value.Y);
            }
            else if (command.Name == DrawingNamespace + "lnTo")
            {
                current = ReadCustomGeometryPoint(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                graphics.LineTo(current.Value.X, current.Value.Y);
            }
            else if (command.Name == DrawingNamespace + "cubicBezTo")
            {
                List<(double X, double Y)> points = ReadCustomGeometryPoints(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                if (points.Count == 3)
                {
                    graphics.CurveTo(points[0].X, points[0].Y, points[1].X, points[1].Y, points[2].X, points[2].Y);
                    current = points[2];
                }
            }
            else if (command.Name == DrawingNamespace + "quadBezTo")
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
            else if (command.Name == DrawingNamespace + "arcTo" && current is { } arcStart)
            {
                current = AppendCustomGeometryArc(graphics, command, arcStart, x, y, width, height, coordinateWidth, coordinateHeight, guides);
            }
            else if (command.Name == DrawingNamespace + "close")
            {
                graphics.ClosePath();
            }
        }
    }

    private static (double X, double Y) ReadCustomGeometryPoint(
        XElement command,
        double x,
        double y,
        double width,
        double height,
        double coordinateWidth,
        double coordinateHeight,
        IReadOnlyDictionary<string, double> guides)
    {
        XElement point = command.Element(DrawingNamespace + "pt") ?? command;
        return MapCustomGeometryPoint(point, x, y, width, height, coordinateWidth, coordinateHeight, guides);
    }

    private static List<(double X, double Y)> ReadCustomGeometryPoints(
        XElement command,
        double x,
        double y,
        double width,
        double height,
        double coordinateWidth,
        double coordinateHeight,
        IReadOnlyDictionary<string, double> guides)
    {
        return command
            .Elements(DrawingNamespace + "pt")
            .Select(point => MapCustomGeometryPoint(point, x, y, width, height, coordinateWidth, coordinateHeight, guides))
            .ToList();
    }

    private static (double X, double Y) MapCustomGeometryPoint(
        XElement point,
        double x,
        double y,
        double width,
        double height,
        double coordinateWidth,
        double coordinateHeight,
        IReadOnlyDictionary<string, double> guides)
    {
        double pointX = ReadCustomGeometryValue((string?)point.Attribute("x"), guides, 0d);
        double pointY = ReadCustomGeometryValue((string?)point.Attribute("y"), guides, 0d);
        return (x + pointX / coordinateWidth * width, y + height - pointY / coordinateHeight * height);
    }

    private static (double X, double Y) AppendCustomGeometryArc(
        PdfGraphicsBuilder graphics,
        XElement arc,
        (double X, double Y) start,
        double x,
        double y,
        double width,
        double height,
        double coordinateWidth,
        double coordinateHeight,
        IReadOnlyDictionary<string, double> guides)
    {
        double radiusX = Math.Abs(ReadCustomGeometryValue((string?)arc.Attribute("wR"), guides, 0d) / coordinateWidth * width);
        double radiusY = Math.Abs(ReadCustomGeometryValue((string?)arc.Attribute("hR"), guides, 0d) / coordinateHeight * height);
        double startAngle = DegreesToRadians(ReadCustomGeometryValue((string?)arc.Attribute("stAng"), guides, 0d) / 60000d);
        double sweepAngle = DegreesToRadians(ReadCustomGeometryValue((string?)arc.Attribute("swAng"), guides, 0d) / 60000d);
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

    private static IReadOnlyDictionary<string, double> BuildCustomGeometryGuides(XElement? customGeometry, double width, double height)
    {
        var guides = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
        {
            ["w"] = width,
            ["h"] = height,
            ["ss"] = Math.Min(width, height),
            ["ls"] = Math.Max(width, height)
        };

        foreach (XElement guide in customGeometry
                     ?.Element(DrawingNamespace + "gdLst")
                     ?.Elements(DrawingNamespace + "gd") ?? [])
        {
            string? name = (string?)guide.Attribute("name");
            string? formula = (string?)guide.Attribute("fmla");
            if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(formula))
            {
                guides[name] = EvaluateCustomGeometryFormula(formula, guides);
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

    private static double ParseOptionalDoubleAttribute(XElement element, string name, double defaultValue)
    {
        return element.Attribute(name) is { } value &&
            double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? parsed
            : defaultValue;
    }

    private static bool TryReadShapePictureFill(
        XElement shapeProperties,
        IReadOnlyDictionary<string, OoxRelationship>? relationships,
        OoxPackage? package,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        List<PdfImageResource>? images,
        Dictionary<string, PdfImageXObject?>? imageCache,
        ref int imageIndex,
        out string? name,
        out PdfImageXObject? image)
    {
        name = null;
        image = null;
        if (relationships is null || package is null || images is null || !CanRenderPictureFillPreset(ReadPreset(shapeProperties)))
        {
            return false;
        }

        string? relationshipId = (string?)shapeProperties
            .Element(DrawingNamespace + "blipFill")
            ?.Element(DrawingNamespace + "blip")
            ?.Attribute(RelationshipsNamespace + "embed");
        if (relationshipId is null ||
            !relationships.TryGetValue(relationshipId, out OoxRelationship? relationship) ||
            relationship.ResolvedTarget is null)
        {
            return false;
        }

        OoxPart? imagePart = package.GetPart(relationship.ResolvedTarget);
        if (imagePart is null)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "IMAGE_MISSING_PART",
                OoxPdfSeverity.Error,
                "Referenced image part was missing and the image was ignored.",
                relationship.ResolvedTarget,
                SlideIndex: slideIndex,
                Feature: "image",
                Fallback: "Ignored"));
            return false;
        }

        image = GetOrCreateImage(imagePart, ImageRecolor.None, imageCache, diagnosticSink, slideIndex);
        if (image is null)
        {
            return false;
        }

        name = "Im" + imageIndex++;
        return true;
    }

    private static bool CanRenderPictureFillPreset(string preset)
    {
        return preset is "rect" or "ellipse" or "roundRect" ||
            TryCreatePresetPolygonPoints(preset, 0d, 0d, 1d, 1d, out _);
    }

    private static void ClipToPresetShape(PdfGraphicsBuilder graphics, string preset, double x, double y, double width, double height)
    {
        if (preset == "ellipse")
        {
            graphics.ClipEllipse(x, y, width, height);
        }
        else if (preset == "roundRect")
        {
            graphics.ClipRoundedRectangle(x, y, width, height, Math.Min(width, height) * 0.16d);
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
        double halfWidth = size * 0.45d;
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
        double markerLength = Math.Max(6.5d, lineWidth * 4d) * style.LengthScale;
        double markerWidth = Math.Max(5d, lineWidth * 3.2d) * style.WidthScale;

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
                    Point(markerLength * 0.7d, 0d),
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

    private static void FillArrowedLine(PdfGraphicsBuilder graphics, double x1, double y1, double x2, double y2, double lineWidth, bool headArrow, bool tailArrow)
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
        double arrowLength = lineWidth * 3d;
        double arrowHalfWidth = lineWidth * 1.5d;
        double startX = headArrow ? x1 + ux * (arrowLength - half) : x1;
        double startY = headArrow ? y1 + uy * (arrowLength - half) : y1;
        double endX = tailArrow ? x2 - ux * (arrowLength - half) : x2;
        double endY = tailArrow ? y2 - uy * (arrowLength - half) : y2;

        graphics.FillPolygon(
        [
            (startX + nx * half, startY + ny * half),
            (endX + nx * half, endY + ny * half),
            (endX - nx * half, endY - ny * half),
            (startX - nx * half, startY - ny * half)
        ]);

        if (headArrow)
        {
            FillTriangle(graphics, x1, y1, x1 + ux * arrowLength, y1 + uy * arrowLength, nx, ny, arrowHalfWidth);
        }

        if (tailArrow)
        {
            FillTriangle(graphics, x2, y2, x2 - ux * arrowLength, y2 - uy * arrowLength, nx, ny, arrowHalfWidth);
        }
    }

    private static void FillTriangle(PdfGraphicsBuilder graphics, double tipX, double tipY, double baseX, double baseY, double nx, double ny, double halfWidth)
    {
        graphics.FillPolygon(
        [
            (tipX, tipY),
            (baseX + nx * halfWidth, baseY + ny * halfWidth),
            (baseX - nx * halfWidth, baseY - ny * halfWidth)
        ]);
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
            AppendOfficeArrowHeadPath(graphics, x1, y1, ux, uy, nx, ny, lineWidth);
        }

        if (tailArrow)
        {
            AppendOfficeArrowHeadPath(graphics, x2, y2, -ux, -uy, -nx, -ny, lineWidth);
        }

        graphics.FillCurrentPath();
    }

    private static void AppendOfficeArrowHeadPath(PdfGraphicsBuilder graphics, double tipX, double tipY, double ux, double uy, double nx, double ny, double lineWidth)
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
        graphics.CurveTo(c1.X, c1.Y, c2.X, c2.Y, c3.X, c3.Y);
        graphics.ClosePath();
    }

    private static void DrawCurvedConnector3(
        PdfGraphicsBuilder graphics,
        double x,
        double yTop,
        double width,
        double height,
        double slideHeight,
        RgbColor stroke,
        double lineWidth,
        bool headArrow,
        bool tailArrow)
    {
        double x1 = x;
        double y1 = slideHeight - yTop;
        double x2 = x + width;
        double y2 = slideHeight - yTop - height;
        double c1x = x2;
        double c1y = y1;
        double c2x = x2;
        double c2y = y2;

        graphics.MoveTo(x1, y1);
        graphics.CurveTo(c1x, c1y, c2x, c2y, x2, y2);
        graphics.StrokeCurrentPath();

        if (headArrow)
        {
            graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
            FillLineArrowhead(graphics, x1, y1, ResolveEndpointDirection(x1, y1, c1x, c1y, c2x, c2y), lineWidth);
        }

        if (tailArrow)
        {
            graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
            FillLineArrowhead(graphics, x2, y2, ResolveEndpointDirection(x2, y2, c2x, c2y, c1x, c1y), lineWidth);
        }
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

    private static LineEndKind ReadLineEndKind(string? type)
    {
        return type switch
        {
            "triangle" => LineEndKind.Triangle,
            "arrow" => LineEndKind.Arrow,
            "stealth" => LineEndKind.Stealth,
            "diamond" => LineEndKind.Diamond,
            "oval" => LineEndKind.Oval,
            _ => LineEndKind.None
        };
    }

    private static double ReadLineEndScale(string? value)
    {
        return value switch
        {
            "sm" => 0.5d,
            "lg" => 1.5d,
            _ => 1d
        };
    }

    private static bool IsFilledTriangleArrow(LineEndStyle style)
    {
        return style.Kind is LineEndKind.Arrow or LineEndKind.Triangle;
    }

    private static bool TryReadPresetDash(XElement shapeProperties, double lineWidth, out IReadOnlyList<double> dashPattern)
    {
        string? presetDash = (string?)shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + "prstDash")
            ?.Attribute("val");
        double w = Math.Max(lineWidth, PptxTextMetricRules.MinimumStrokeWidth);
        dashPattern = presetDash switch
        {
            "dot" or "sysDot" => [w, w * 2d],
            "dash" or "sysDash" => [w * 4d, w * 3d],
            "lgDash" => [w * 8d, w * 3d],
            "dashDot" or "sysDashDot" => [w * 4d, w * 3d, w, w * 3d],
            "lgDashDot" => [w * 8d, w * 3d, w, w * 3d],
            "lgDashDotDot" or "sysDashDotDot" => [w * 8d, w * 3d, w, w * 3d, w, w * 3d],
            _ => []
        };
        return dashPattern.Count > 0;
    }

    private static string? ReadLineCap(XElement shapeProperties)
    {
        return (string?)shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Attribute("cap");
    }

    private static int? ReadLineJoin(XElement shapeProperties)
    {
        XElement? line = shapeProperties.Element(DrawingNamespace + "ln");
        if (line?.Element(DrawingNamespace + "round") is not null)
        {
            return 1;
        }

        if (line?.Element(DrawingNamespace + "bevel") is not null)
        {
            return 2;
        }

        if (line?.Element(DrawingNamespace + "miter") is not null)
        {
            return 0;
        }

        return null;
    }

    private static bool TryCreatePresetPolygonPoints(string preset, double x, double y, double width, double height, out (double X, double Y)[] points)
    {
        points = preset switch
        {
            "triangle" => CreateTrianglePoints(x, y, width, height),
            "rtTriangle" => CreateRightTrianglePoints(x, y, width, height),
            "diamond" => CreateDiamondPoints(x, y, width, height),
            "pentagon" => CreatePentagonPoints(x, y, width, height),
            "hexagon" => CreateHexagonPoints(x, y, width, height),
            "octagon" => CreateOctagonPoints(x, y, width, height),
            "star4" => CreateStar4Points(x, y, width, height),
            "star5" => CreateStar5Points(x, y, width, height),
            "star6" => CreateStar6Points(x, y, width, height),
            "star8" => CreateStar8Points(x, y, width, height),
            "plus" => CreatePlusPoints(x, y, width, height),
            "chevron" => CreateChevronPoints(x, y, width, height),
            "homePlate" => CreateHomePlatePoints(x, y, width, height),
            "parallelogram" => CreateParallelogramPoints(x, y, width, height),
            "trapezoid" => CreateTrapezoidPoints(x, y, width, height),
            "downArrow" => CreateDownArrowPoints(x, y, width, height),
            "upArrow" => CreateUpArrowPoints(x, y, width, height),
            "leftArrow" => CreateLeftArrowPoints(x, y, width, height),
            "rightArrow" => CreateRightArrowPoints(x, y, width, height),
            "leftRightArrow" => CreateLeftRightArrowPoints(x, y, width, height),
            "upDownArrow" => CreateUpDownArrowPoints(x, y, width, height),
            "wedgeRectCallout" => CreateWedgeRectCalloutPoints(x, y, width, height),
            _ => []
        };
        return points.Length != 0;
    }

    private static (double X, double Y)[] CreateTrianglePoints(double x, double y, double width, double height)
    {
        return
        [
            (x + width / 2d, y + height),
            (x + width, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateRightTrianglePoints(double x, double y, double width, double height)
    {
        return
        [
            (x, y),
            (x, y + height),
            (x + width, y)
        ];
    }

    private static (double X, double Y)[] CreateDiamondPoints(double x, double y, double width, double height)
    {
        return
        [
            (x + width / 2d, y + height),
            (x + width, y + height / 2d),
            (x + width / 2d, y),
            (x, y + height / 2d)
        ];
    }

    private static (double X, double Y)[] CreateStar4Points(double x, double y, double width, double height)
    {
        double innerLeftX = RoundOfficeShapeCoordinate(width * 0.41161d);
        double innerRightX = RoundOfficeShapeCoordinate(width * 0.58839d);
        double innerBottomY = RoundOfficeShapeCoordinate(height * 0.41159d);
        double innerTopY = RoundOfficeShapeCoordinate(height * 0.58841d);
        return
        [
            (x, y + height / 2d),
            (x + innerLeftX, y + innerTopY),
            (x + width / 2d, y + height),
            (x + innerRightX, y + innerTopY),
            (x + width, y + height / 2d),
            (x + innerRightX, y + innerBottomY),
            (x + width / 2d, y),
            (x + innerLeftX, y + innerBottomY)
        ];
    }

    private static (double X, double Y)[] CreateStar5Points(double x, double y, double width, double height)
    {
        double topShoulderY = RoundOfficeShapeCoordinate(height * 0.61803d);
        double lowerShoulderY = RoundOfficeShapeCoordinate(height * 0.38197d);
        double innerBottomY = RoundOfficeShapeCoordinate(height * 0.23607d);
        double innerTopLeftX = RoundOfficeShapeCoordinate(width * 0.38194d);
        double innerTopRightX = RoundOfficeShapeCoordinate(width * 0.61806d);
        double innerLeftX = RoundOfficeShapeCoordinate(width * 0.30902d);
        double innerRightX = RoundOfficeShapeCoordinate(width * 0.69098d);
        double lowerLeftX = RoundOfficeShapeCoordinate(width * 0.19098d);
        double lowerRightX = RoundOfficeShapeCoordinate(width * 0.80902d);
        return
        [
            (x, y + topShoulderY),
            (x + innerTopLeftX, y + topShoulderY),
            (x + width / 2d, y + height),
            (x + innerTopRightX, y + topShoulderY),
            (x + width, y + topShoulderY),
            (x + innerRightX, y + lowerShoulderY),
            (x + lowerRightX, y),
            (x + width / 2d, y + innerBottomY),
            (x + lowerLeftX, y),
            (x + innerLeftX, y + lowerShoulderY)
        ];
    }

    private static (double X, double Y)[] CreateStar6Points(double x, double y, double width, double height)
    {
        double quarterWidth = width / 6d;
        double quarterHeight = height / 4d;
        return
        [
            (x, y + height * 0.75d),
            (x + width / 3d, y + height * 0.75d),
            (x + width / 2d, y + height),
            (x + width * 2d / 3d, y + height * 0.75d),
            (x + width, y + height * 0.75d),
            (x + width - quarterWidth, y + height / 2d),
            (x + width, y + height * 0.25d),
            (x + width * 2d / 3d, y + height * 0.25d),
            (x + width / 2d, y),
            (x + width / 3d, y + height * 0.25d),
            (x, y + height * 0.25d),
            (x + quarterWidth, y + height / 2d)
        ];
    }

    private static (double X, double Y)[] CreateStar8Points(double x, double y, double width, double height)
    {
        return
        [
            (x, y + height / 2d),
            (x + RoundOfficeShapeCoordinate(width * 0.15356d), y + RoundOfficeShapeCoordinate(height * 0.64348d)),
            (x + RoundOfficeShapeCoordinate(width * 0.14644d), y + RoundOfficeShapeCoordinate(height * 0.85356d)),
            (x + RoundOfficeShapeCoordinate(width * 0.3565d), y + RoundOfficeShapeCoordinate(height * 0.84644d)),
            (x + width / 2d, y + height),
            (x + RoundOfficeShapeCoordinate(width * 0.6435d), y + RoundOfficeShapeCoordinate(height * 0.84644d)),
            (x + RoundOfficeShapeCoordinate(width * 0.85356d), y + RoundOfficeShapeCoordinate(height * 0.85356d)),
            (x + RoundOfficeShapeCoordinate(width * 0.84644d), y + RoundOfficeShapeCoordinate(height * 0.64348d)),
            (x + width, y + height / 2d),
            (x + RoundOfficeShapeCoordinate(width * 0.84644d), y + RoundOfficeShapeCoordinate(height * 0.35652d)),
            (x + RoundOfficeShapeCoordinate(width * 0.85356d), y + RoundOfficeShapeCoordinate(height * 0.14644d)),
            (x + RoundOfficeShapeCoordinate(width * 0.6435d), y + RoundOfficeShapeCoordinate(height * 0.15356d)),
            (x + width / 2d, y),
            (x + RoundOfficeShapeCoordinate(width * 0.3565d), y + RoundOfficeShapeCoordinate(height * 0.15356d)),
            (x + RoundOfficeShapeCoordinate(width * 0.14644d), y + RoundOfficeShapeCoordinate(height * 0.14644d)),
            (x + RoundOfficeShapeCoordinate(width * 0.15356d), y + RoundOfficeShapeCoordinate(height * 0.35652d))
        ];
    }

    private static (double X, double Y)[] CreatePentagonPoints(double x, double y, double width, double height)
    {
        double golden = (Math.Sqrt(5d) - 1d) / 2d;
        double bottomInset = RoundOfficeShapeCoordinate(width * (1d - golden) / 2d);
        double shoulderY = RoundOfficeShapeCoordinate(y + height * golden);
        return
        [
            (x, shoulderY),
            (x + width / 2d, y + height),
            (x + width, shoulderY),
            (x + width - bottomInset, y),
            (x + bottomInset, y)
        ];
    }

    private static (double X, double Y)[] CreateHexagonPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width / 2d, height / 4d);
        return
        [
            (x, y + height / 2d),
            (x + inset, y + height),
            (x + width - inset, y + height),
            (x + width, y + height / 2d),
            (x + width - inset, y),
            (x + inset, y)
        ];
    }

    private static (double X, double Y)[] CreateOctagonPoints(double x, double y, double width, double height)
    {
        double inset = RoundOfficeShapeCoordinate(Math.Min(width, height) * (1d - Math.Sqrt(0.5d)));
        return
        [
            (x, y + height - inset),
            (x + inset, y + height),
            (x + width - inset, y + height),
            (x + width, y + height - inset),
            (x + width, y + inset),
            (x + width - inset, y),
            (x + inset, y),
            (x, y + inset)
        ];
    }

    private static double RoundOfficeShapeCoordinate(double value)
    {
        return Math.Round(value, 2, MidpointRounding.AwayFromZero);
    }

    private static (double X, double Y)[] CreatePlusPoints(double x, double y, double width, double height)
    {
        double armWidth = Math.Min(width, height) / 4d;
        double rightArmX = x + width - armWidth;
        double topArmY = y + height - armWidth;
        return
        [
            (x, topArmY),
            (x + armWidth, topArmY),
            (x + armWidth, y + height),
            (rightArmX, y + height),
            (rightArmX, topArmY),
            (x + width, topArmY),
            (x + width, y + armWidth),
            (rightArmX, y + armWidth),
            (rightArmX, y),
            (x + armWidth, y),
            (x + armWidth, y + armWidth),
            (x, y + armWidth)
        ];
    }

    private static (double X, double Y)[] CreateChevronPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width / 2d, height / 2d);
        return
        [
            (x, y + height),
            (x + width - inset, y + height),
            (x + width, y + height / 2d),
            (x + width - inset, y),
            (x, y),
            (x + inset, y + height / 2d)
        ];
    }

    private static (double X, double Y)[] CreateHomePlatePoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width, height / 2d);
        return
        [
            (x, y + height),
            (x + width - inset, y + height),
            (x + width, y + height / 2d),
            (x + width - inset, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateParallelogramPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width, height / 4d);
        return
        [
            (x + inset, y + height),
            (x + width, y + height),
            (x + width - inset, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateTrapezoidPoints(double x, double y, double width, double height)
    {
        double inset = Math.Min(width / 2d, height / 4d);
        return
        [
            (x + inset, y + height),
            (x + width - inset, y + height),
            (x + width, y),
            (x, y)
        ];
    }

    private static (double X, double Y)[] CreateDownArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderY = y + Math.Min(height, width / 2d);
        return
        [
            (x + width * 0.25d, y + height),
            (x + width * 0.75d, y + height),
            (x + width * 0.75d, arrowShoulderY),
            (x + width, arrowShoulderY),
            (x + width * 0.5d, y),
            (x, arrowShoulderY),
            (x + width * 0.25d, arrowShoulderY)
        ];
    }

    private static (double X, double Y)[] CreateUpArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderY = y + Math.Max(0d, height - width / 2d);
        return
        [
            (x + width * 0.25d, y),
            (x + width * 0.75d, y),
            (x + width * 0.75d, arrowShoulderY),
            (x + width, arrowShoulderY),
            (x + width * 0.5d, y + height),
            (x, arrowShoulderY),
            (x + width * 0.25d, arrowShoulderY)
        ];
    }

    private static (double X, double Y)[] CreateRightArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderX = x + Math.Max(0d, width - height / 2d);
        return
        [
            (x, y + height * 0.25d),
            (arrowShoulderX, y + height * 0.25d),
            (arrowShoulderX, y),
            (x + width, y + height * 0.5d),
            (arrowShoulderX, y + height),
            (arrowShoulderX, y + height * 0.75d),
            (x, y + height * 0.75d)
        ];
    }

    private static (double X, double Y)[] CreateLeftArrowPoints(double x, double y, double width, double height)
    {
        double arrowShoulderX = x + Math.Min(width, height / 2d);
        return
        [
            (x + width, y + height * 0.25d),
            (arrowShoulderX, y + height * 0.25d),
            (arrowShoulderX, y),
            (x, y + height * 0.5d),
            (arrowShoulderX, y + height),
            (arrowShoulderX, y + height * 0.75d),
            (x + width, y + height * 0.75d)
        ];
    }

    private static (double X, double Y)[] CreateLeftRightArrowPoints(double x, double y, double width, double height)
    {
        double headDepth = Math.Min(width / 2d, height / 2d);
        double shaftBottom = y + height * 0.25d;
        double shaftTop = y + height * 0.75d;
        return
        [
            (x, y + height / 2d),
            (x + headDepth, y + height),
            (x + headDepth, shaftTop),
            (x + width - headDepth, shaftTop),
            (x + width - headDepth, y + height),
            (x + width, y + height / 2d),
            (x + width - headDepth, y),
            (x + width - headDepth, shaftBottom),
            (x + headDepth, shaftBottom),
            (x + headDepth, y)
        ];
    }

    private static (double X, double Y)[] CreateUpDownArrowPoints(double x, double y, double width, double height)
    {
        double headDepth = Math.Min(height / 2d, width / 2d);
        double shaftLeft = x + width * 0.25d;
        double shaftRight = x + width * 0.75d;
        return
        [
            (x, y + height - headDepth),
            (x + width / 2d, y + height),
            (x + width, y + height - headDepth),
            (shaftRight, y + height - headDepth),
            (shaftRight, y + headDepth),
            (x + width, y + headDepth),
            (x + width / 2d, y),
            (x, y + headDepth),
            (shaftLeft, y + headDepth),
            (shaftLeft, y + height - headDepth)
        ];
    }

    private static (double X, double Y)[] CreateWedgeRectCalloutPoints(double x, double y, double width, double height)
    {
        double tailLeftX = x + width / 6d;
        double tailTipX = x + width * 7d / 24d;
        double tailRightX = x + width * 5d / 12d;
        double tailTipY = y - height / 8d;
        return
        [
            (x, y + height),
            (x + width, y + height),
            (x + width, y),
            (tailRightX, y),
            (tailTipX, tailTipY),
            (tailLeftX, y),
            (x, y)
        ];
    }

    private static GroupTransform ReadGroupTransform(XElement group)
    {
        XElement? transform = group
            .Element(PresentationNamespace + "grpSpPr")
            ?.Element(DrawingNamespace + "xfrm");
        XElement? offset = transform?.Element(DrawingNamespace + "off");
        XElement? extents = transform?.Element(DrawingNamespace + "ext");
        XElement? childOffset = transform?.Element(DrawingNamespace + "chOff");
        XElement? childExtents = transform?.Element(DrawingNamespace + "chExt");
        if (offset is null || extents is null || childOffset is null || childExtents is null)
        {
            return GroupTransform.Identity;
        }

        long chWidth = Math.Max(1, ParseLongAttribute(childExtents, "cx"));
        long chHeight = Math.Max(1, ParseLongAttribute(childExtents, "cy"));
        return new GroupTransform(
            ParseLongAttribute(offset, "x"),
            ParseLongAttribute(offset, "y"),
            ParseLongAttribute(extents, "cx"),
            ParseLongAttribute(extents, "cy"),
            ParseLongAttribute(childOffset, "x"),
            ParseLongAttribute(childOffset, "y"),
            ParseLongAttribute(extents, "cx") / (double)chWidth,
            ParseLongAttribute(extents, "cy") / (double)chHeight,
            transform!.Attribute("rot") is { } rotationAttribute
                ? long.Parse(rotationAttribute.Value, CultureInfo.InvariantCulture) / 60000d
                : 0d);
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

    private static bool TryReadShapeFill(XElement shape, XElement shapeProperties, PptxTheme theme, out RgbColor color, out double alpha)
    {
        if (shapeProperties.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            alpha = 1d;
            return false;
        }

        if (TryReadSolidColorWithAlpha(shapeProperties, theme, out color, out alpha))
        {
            return true;
        }

        XElement? fillRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fillRef");
        int fillIndex = ParseOptionalIntAttribute(fillRef, "idx", 0);
        if (fillIndex > 0 &&
            theme.TryGetFillStyle(fillIndex, out XElement fillStyle) &&
            TryReadSolidColorWithAlpha(fillStyle, theme, fillRef, out color, out alpha))
        {
            return true;
        }

        return fillIndex > 0 && TryReadSolidColorWithAlpha(fillRef, theme, out color, out alpha);
    }

    private static bool TryReadShapePatternFill(XElement shapeProperties, PptxTheme theme, out ShapePatternFill fill)
    {
        XElement? patternFill = shapeProperties.Element(DrawingNamespace + "pattFill");
        string? preset = (string?)patternFill?.Attribute("prst");
        if (patternFill is null || !IsSupportedDiagonalPatternFill(preset))
        {
            fill = default;
            return false;
        }

        RgbColor foreground = TryReadSolidColor(patternFill.Element(DrawingNamespace + "fgClr"), theme, out RgbColor foregroundColor)
            ? foregroundColor
            : new RgbColor(0, 0, 0);
        RgbColor background = TryReadSolidColor(patternFill.Element(DrawingNamespace + "bgClr"), theme, out RgbColor backgroundColor)
            ? backgroundColor
            : new RgbColor(255, 255, 255);
        fill = new ShapePatternFill(preset!, foreground, background, 1d);
        return true;
    }

    private static void StrokeShapePatternFill(PdfGraphicsBuilder graphics, string preset, double x, double y, double width, double height, ShapePatternFill fill)
    {
        if (width <= 0d || height <= 0d)
        {
            return;
        }

        graphics.SaveState();
        ClipToPresetShape(graphics, preset, x, y, width, height);
        graphics.SetStrokeRgb(fill.Foreground.Red, fill.Foreground.Green, fill.Foreground.Blue);
        graphics.SetLineWidth(IsDarkDiagonalPatternFill(fill.Preset) ? 1.0d : 0.5d);
        double spacing = IsDarkDiagonalPatternFill(fill.Preset) ? 4d : 5d;
        bool up = fill.Preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase);
        for (double offset = -height; offset <= width + height; offset += spacing)
        {
            if (up)
            {
                graphics.StrokeLine(x + offset, y, x + offset + height, y + height);
            }
            else
            {
                graphics.StrokeLine(x + offset, y + height, x + offset + height, y);
            }
        }

        graphics.RestoreState();
    }

    private static bool IsSupportedDiagonalPatternFill(string? preset)
    {
        return preset is not null &&
            (preset.Contains("UpDiag", StringComparison.OrdinalIgnoreCase) ||
             preset.Contains("DnDiag", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsDarkDiagonalPatternFill(string preset)
    {
        return preset.StartsWith("dk", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryReadShapeLine(XElement shape, XElement shapeProperties, PptxTheme theme, out RgbColor color, out double lineWidth, out double alpha)
    {
        XElement? explicitLine = shapeProperties.Element(DrawingNamespace + "ln");
        if (explicitLine?.Element(DrawingNamespace + "noFill") is not null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        if (explicitLine is not null && TryReadLineWithAlpha(shapeProperties, theme, out color, out lineWidth, out alpha))
        {
            return true;
        }

        XElement? lineRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "lnRef");
        int lineIndex = ParseOptionalIntAttribute(lineRef, "idx", 0);
        if (lineIndex <= 0 || !theme.TryGetLineStyle(lineIndex, out XElement lineStyle))
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        lineWidth = lineStyle.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColorWithAlpha(lineStyle, theme, lineRef, out color, out alpha);
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, out RgbColor color)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return TryReadSolidColor(fontRef, theme, out color);
    }
}
