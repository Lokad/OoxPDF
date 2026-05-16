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

        if (transformed)
        {
            graphics.SaveState();
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
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
                string? headArrowType = ReadLineEndType(shapeProperties, "headEnd");
                string? tailArrowType = ReadLineEndType(shapeProperties, "tailEnd");
                bool hasHeadArrow = IsFilledTriangleArrow(headArrowType);
                bool hasTailArrow = IsFilledTriangleArrow(tailArrowType);
                if ((hasHeadArrow || hasTailArrow) && !hasDash && lineCap is null)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    bool usesOfficeArrowType = headArrowType == "arrow" || tailArrowType == "arrow";
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
        else if (hasFill)
        {
            bool transparentFill = fillAlpha < 0.999d;
            if (transparentFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(fillAlpha, 1d);
            }

            graphics.SetFillRgb(fill.Red, fill.Green, fill.Blue);
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

            if (preset == "ellipse")
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

    private static string ReadPreset(XElement shapeProperties)
    {
        return (string?)shapeProperties
            .Element(DrawingNamespace + "prstGeom")
            ?.Attribute("prst") ?? "rect";
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

        image = GetOrCreateImage(imagePart, imageCache, diagnosticSink, slideIndex);
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

    private static string? ReadLineEndType(XElement shapeProperties, string elementName)
    {
        return (string?)shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + elementName)
            ?.Attribute("type");
    }

    private static bool IsFilledTriangleArrow(string? type)
    {
        return type is "arrow" or "triangle";
    }

    private static bool TryReadPresetDash(XElement shapeProperties, double lineWidth, out IReadOnlyList<double> dashPattern)
    {
        string? presetDash = (string?)shapeProperties
            .Element(DrawingNamespace + "ln")
            ?.Element(DrawingNamespace + "prstDash")
            ?.Attribute("val");
        if (presetDash == "dash")
        {
            dashPattern = [lineWidth * 4d, lineWidth * 3d];
            return true;
        }

        if (presetDash == "dashDot")
        {
            dashPattern = [lineWidth * 4d, lineWidth * 3d, lineWidth, lineWidth * 3d];
            return true;
        }

        dashPattern = [];
        return false;
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

    private static IReadOnlyList<TextRun> RenderTables(PptxRenderContext context, XDocument slideXml, PdfGraphicsBuilder graphics)
    {
        var textRuns = new List<TextRun>();
        foreach (XElement frame in slideXml.Descendants(PresentationNamespace + "graphicFrame"))
        {
            textRuns.AddRange(RenderTableFrame(context, frame, graphics));
        }

        return textRuns;
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
            ParseLongAttribute(childOffset, "x"),
            ParseLongAttribute(childOffset, "y"),
            ParseLongAttribute(extents, "cx") / (double)chWidth,
            ParseLongAttribute(extents, "cy") / (double)chHeight);
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