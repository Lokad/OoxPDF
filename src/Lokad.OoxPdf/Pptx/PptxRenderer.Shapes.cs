using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private const int OfficeArrowTailConnectorSamplesPerSegment = 61;
    private const int OfficeTriangleTailConnectorSamplesPerSegment = 54;
    private const int OfficeStealthTailConnectorSamplesPerSegment = 66;
    private const double OfficeArrowheadLengthFactor = 4.423333d;
    private const double OfficeTriangleTailMinimumLength = 5d;
    private const double OfficeTriangleTailLengthFactor = 3.5d;
    private const double OfficeTriangleTailHalfWidthFactor = 0.45d;
    private const double OfficeStraightTriangleLineEndLengthFactor = 4d;
    private const double OfficeStraightTriangleLineEndHalfWidthFactor = 1.5d;
    private const double OfficeStraightTriangleLineEndOverlapFactor = 2d / 3d;
    private const double OfficeStraightStealthLineEndLengthFactor = 3d;
    private const double OfficeStraightStealthLineEndWidthFactor = 3d;
    private const double OfficeStraightStealthLineEndNotchFactor = 2d / 3d;
    private const double OfficeGlowRasterPixelsPerPoint = 1d;
    private const int OfficeGlowRasterMaxPixelsPerSide = 2048;

    private static void RenderBackground(PptxRenderContext context, PptxSceneBackground background, PdfGraphicsBuilder graphics, bool defaultWhenMissing)
    {
        if (background.HasFill)
        {
            graphics.SaveState();
            ClipSlideBoundsEvenOdd(context.Document, graphics);
            if (background.Alpha < 0.999d)
            {
                graphics.SetAlpha(background.Alpha, 1d);
            }

            graphics.SetFillRgb(background.Color.Red, background.Color.Green, background.Color.Blue);
            graphics.FillRectangleEvenOdd(0, 0, context.Document.SlideWidthPoints, context.Document.SlideHeightPoints);
            graphics.RestoreState();
            return;
        }

        if (!defaultWhenMissing)
        {
            return;
        }

        graphics.SaveState();
        ClipSlideBoundsEvenOdd(context.Document, graphics);
        graphics.SetFillRgb(255, 255, 255);
        graphics.FillRectangleEvenOdd(0, 0, context.Document.SlideWidthPoints, context.Document.SlideHeightPoints);
        graphics.RestoreState();
    }

    private static void RenderShape(
        PptxSceneNode shape,
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
        if (shape.Shape is null || shape.Bounds is null)
        {
            return;
        }

        RenderShape(
            shape.Source,
            document,
            graphics,
            diagnosticSink,
            slideIndex,
            theme,
            groupTransform,
            images,
            imageCache,
            ref imageIndex,
            ToShapeBounds(shape.Bounds),
            shape.Shape.Preset,
            shape.Shape.PresetAdjustments,
            shape.Shape.HasCustomGeometry,
            shape.Shape.CustomGeometry,
            shape.Shape.NoFill,
            ToFillStyle(shape.Shape.Fill),
            ToGradientFill(shape.Shape.GradientFill),
            ToShapePatternFill(shape.Shape.PatternFill),
            ToShapePictureFill(shape.Shape.PictureFill),
            ToGlow(shape.Shape.Glow),
            ToOuterShadow(shape.Shape.OuterShadow),
            ToLineStyle(shape.Shape.Line),
            ToLineEndStyle(shape.Shape.HeadEnd),
            ToLineEndStyle(shape.Shape.TailEnd));
    }

    private static void RenderShape(
        XElement shape,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        PptxTheme theme,
        GroupTransform groupTransform,
        List<PdfImageResource>? images,
        Dictionary<string, PdfImageXObject?>? imageCache,
        ref int imageIndex,
        ShapeBounds rawBounds,
        string preset,
        IReadOnlyDictionary<string, double>? presetAdjustmentsOverride,
        bool hasCustomGeometry,
        PptxSceneCustomGeometry? customGeometryOverride,
        bool noFillOverride,
        FillStyle? fillOverride,
        GradientFill? gradientFillOverride,
        ShapePatternFill? patternFillOverride,
        ShapePictureFill? pictureFillOverride,
        Glow? glowOverride,
        OuterShadow? outerShadowOverride,
        LineStyle? lineOverride,
        LineEndStyle? headEndOverride,
        LineEndStyle? tailEndOverride)
    {
        XElement? shapeProperties = shape.Element(PresentationNamespace + "spPr");
        if (shapeProperties is null)
        {
            return;
        }

        ShapeBounds bounds = groupTransform.Apply(rawBounds);

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        bool transformed = bounds.RotationDegrees != 0d || bounds.FlipHorizontal || bounds.FlipVertical;

        RgbColor fill;
        double fillAlpha;
        bool hasFill;
        if (fillOverride is { } resolvedFill)
        {
            fill = resolvedFill.Color;
            fillAlpha = resolvedFill.Alpha;
            hasFill = resolvedFill.HasFill;
        }
        else if (noFillOverride)
        {
            fill = default;
            fillAlpha = 1d;
            hasFill = false;
        }
        else
        {
            fill = default;
            fillAlpha = 1d;
            hasFill = false;
        }
        GradientFill? gradientFill = gradientFillOverride;
        bool hasPatternFill = patternFillOverride is not null;
        ShapePatternFill patternFill;
        patternFill = patternFillOverride ?? default;
        RgbColor stroke;
        double lineWidth;
        double strokeAlpha;
        bool hasStroke;
        if (lineOverride is { } line)
        {
            stroke = line.Color;
            lineWidth = line.Width;
            strokeAlpha = line.Alpha;
            hasStroke = line.HasLine;
        }
        else
        {
            hasStroke = TryReadShapeLine(shape, shapeProperties, theme, out stroke, out lineWidth, out strokeAlpha);
        }
        bool hasDash;
        IReadOnlyList<double> dashPattern;
        int? lineCap;
        int? lineJoin;
        if (lineOverride is { } resolvedLine)
        {
            hasDash = resolvedLine.HasDash;
            dashPattern = resolvedLine.DashPattern;
            lineCap = resolvedLine.Cap;
            lineJoin = resolvedLine.Join;
        }
        else
        {
            hasDash = TryReadPresetDash(shapeProperties, lineWidth, out dashPattern);
            lineCap = ReadLineCap(shapeProperties) switch
            {
                "rnd" => 1,
                "sq" => 2,
                _ => null
            };
            lineJoin = ReadLineJoin(shapeProperties);
        }
        LineEndStyle headEnd = headEndOverride ?? ReadLineEnd(shapeProperties, "headEnd");
        LineEndStyle tailEnd = tailEndOverride ?? ReadLineEnd(shapeProperties, "tailEnd");
        bool hasPictureFill = TryReadShapePictureFill(
            shapeProperties,
            diagnosticSink,
            slideIndex,
            images,
            imageCache,
            ref imageIndex,
            pictureFillOverride,
            out string? pictureFillName,
            out PdfImageXObject? pictureFillImage,
            out CropRect pictureFillCrop,
            out FillRect pictureFillRect,
            out double pictureFillAlpha);
        PptxSceneCustomGeometry? sceneCustomGeometry = customGeometryOverride is { HasGeometry: true } ? customGeometryOverride : null;
        XElement? customGeometry = hasCustomGeometry && sceneCustomGeometry is null ? shapeProperties.Element(DrawingNamespace + "custGeom") : null;

        if (transformed)
        {
            graphics.SaveState();
            ApplyShapeTransform(graphics, x, y, width, height, bounds);
        }

        bool hasOuterShadow = outerShadowOverride is not null;
        OuterShadow outerShadow;
        outerShadow = outerShadowOverride ?? default;

        if (glowOverride is { } glow &&
            CanRenderGlowPreset(preset) &&
            images is not null)
        {
            DrawRasterGlow(graphics, x, y, width, height, glow, images, ref imageIndex);
        }

        if (hasOuterShadow &&
            preset is not ("line" or "straightConnector1" or "curvedConnector2" or "curvedConnector3") &&
            customGeometry is null && sceneCustomGeometry is null)
        {
            DrawOuterShadow(graphics, preset, x, y, width, height, outerShadow);
        }

        if ((sceneCustomGeometry is not null && TryRenderCustomGeometry(
                sceneCustomGeometry,
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
                lineJoin,
                headEnd,
                tailEnd)) ||
            (customGeometry is not null && TryRenderCustomGeometry(
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
                lineJoin,
                headEnd,
                tailEnd)))
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
                bool hasHeadArrow = IsFilledTriangleArrow(headEnd);
                bool hasTailArrow = IsFilledTriangleArrow(tailEnd);
                bool hasStealthEnd = headEnd.Kind == LineEndKind.Stealth || tailEnd.Kind == LineEndKind.Stealth;
                if (hasStealthEnd && headEnd.Kind is LineEndKind.None or LineEndKind.Stealth && tailEnd.Kind is LineEndKind.None or LineEndKind.Stealth && !hasDash && lineCap is null)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    FillStealthEndedLine(graphics, x1, y1, x2, y2, lineWidth, headEnd, tailEnd);
                }
                else if ((hasHeadArrow || hasTailArrow) && headEnd.Kind is LineEndKind.None or LineEndKind.Triangle or LineEndKind.Arrow && tailEnd.Kind is LineEndKind.None or LineEndKind.Triangle or LineEndKind.Arrow && !hasDash && lineCap is null)
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

                if (!hasDash &&
                    lineCap is null &&
                    headEnd.IsNone &&
                    tailEnd.Kind is LineEndKind.Triangle or LineEndKind.Arrow or LineEndKind.Stealth)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    if (!TryFillCurvedConnectorPreset(
                        graphics,
                        shapeProperties,
                        preset,
                        x,
                        yTop,
                        width,
                        height,
                        document.SlideHeightPoints,
                        lineWidth,
                        tailEnd,
                        presetAdjustmentsOverride))
                    {
                        DrawCurvedConnectorPreset(
                            graphics,
                            shapeProperties,
                            preset,
                            x,
                            yTop,
                            width,
                            height,
                            document.SlideHeightPoints,
                            stroke,
                            lineWidth,
                            headEnd,
                            tailEnd,
                            presetAdjustmentsOverride);
                    }
                }
                else
                {
                    DrawCurvedConnectorPreset(
                        graphics,
                        shapeProperties,
                        preset,
                        x,
                        yTop,
                        width,
                        height,
                        document.SlideHeightPoints,
                        stroke,
                        lineWidth,
                        headEnd,
                        tailEnd,
                        presetAdjustmentsOverride);
                }

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

        bool repeatSlideClipBeforeFill = hasStroke && hasFill && !transformed;
        if (repeatSlideClipBeforeFill)
        {
            ClipSlideBoundsEvenOdd(document, graphics);
        }

        if (gradientFill is not null)
        {
            graphics.SaveState();
            ClipToPresetShape(graphics, preset, x, y, width, height);

            DrawLinearGradientFill(graphics, gradientFill, x, y, width, height);

            graphics.RestoreState();
        }
        else if (hasPictureFill && pictureFillName is not null && pictureFillImage is not null)
        {
            CropRect crop = pictureFillCrop;
            FillRect fillRect = pictureFillRect;
            double imageX = x + fillRect.Left * width;
            double imageY = y + fillRect.Bottom * height;
            double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
            double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));
            bool transparentPictureFill = pictureFillAlpha < 0.999d;
            if (transparentPictureFill)
            {
                graphics.SaveState();
                graphics.SetAlpha(pictureFillAlpha, 1d);
            }

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

            if (transparentPictureFill)
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
            if (hasFill && !repeatSlideClipBeforeFill)
            {
                if (transformed)
                {
                    graphics.RestoreState();
                    ClipSlideBoundsEvenOdd(document, graphics);
                    graphics.SaveState();
                    ApplyShapeTransform(graphics, x, y, width, height, bounds);
                }
                else
                {
                    ClipSlideBoundsEvenOdd(document, graphics);
                }
            }

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
                DrawPresetArcStroke(graphics, shapeProperties, x, y, width, height, presetAdjustmentsOverride);
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
            graphics.FillEllipseEvenOdd(x, y, width, height);
        }
        else if (preset == "roundRect")
        {
            graphics.FillRoundedRectangleEvenOdd(x, y, width, height, Math.Min(width, height) * 0.16d);
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
        DrawPresetFill(graphics, preset, x - glow.Radius, y - glow.Radius, width + 2d * glow.Radius, height + 2d * glow.Radius);
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
                alpha[pixel] = (byte)Math.Clamp((int)Math.Round(255d * glow.Alpha * falloff), 0, 255);
            }
        }

        var image = PdfImageXObject.RgbPng(pixelWidth, pixelHeight, rgb, alpha);
        string name = "Im" + imageIndex++;
        graphics.SaveState();
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
        double dy = Math.Sin(radians);
        double half = Math.Abs(dx) * width / 2d + Math.Abs(dy) * height / 2d;
        double centerX = x + width / 2d;
        double centerY = y + height / 2d;
        bool alphaState = TryGetUniformGradientAlpha(gradient.Stops, out double alpha) && alpha < 0.999d;
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

    private static bool TryGetUniformGradientAlpha(IReadOnlyList<GradientStop> stops, out double alpha)
    {
        double candidate = stops[0].Alpha;
        alpha = candidate;
        return stops.All(stop => Math.Abs(stop.Alpha - candidate) <= 0.001d);
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
        double startDegrees = ReadPresetGeometryGuide(shapeProperties, presetAdjustmentsOverride, "adj1", 0d) / 60000d;
        double endDegrees = ReadPresetGeometryGuide(shapeProperties, presetAdjustmentsOverride, "adj2", 90d) / 60000d;
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
        int? lineJoin,
        LineEndStyle headEnd,
        LineEndStyle tailEnd)
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

            foreach (XElement path in paths.Where(CustomGeometryPathAllowsStroke))
            {
                if (TryFillCustomGeometryOpenLineEndPath(
                    graphics,
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
        return ParseBoolAttribute(path, "stroke", defaultValue: true);
    }

    private static bool TryFillCustomGeometryOpenLineEndPath(
        PdfGraphicsBuilder graphics,
        XElement path,
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
            hasFill && CustomGeometryPathAllowsFill(path))
        {
            return false;
        }

        if (!TryCreateCustomGeometryBezierSegments(path, x, y, width, height, out List<BezierSegment> segments))
        {
            return false;
        }

        return TryFillBezierConnectorPath(graphics, segments, lineWidth, tailEnd);
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
        XElement path,
        double x,
        double y,
        double width,
        double height,
        out List<BezierSegment> segments)
    {
        double coordinateWidth = Math.Max(1d, ParseOptionalDoubleAttribute(path, "w", 21600d));
        double coordinateHeight = Math.Max(1d, ParseOptionalDoubleAttribute(path, "h", 21600d));
        IReadOnlyDictionary<string, double> guides = BuildCustomGeometryGuides(
            path.Parent?.Parent,
            coordinateWidth,
            coordinateHeight);
        segments = [];
        (double X, double Y)? current = null;
        bool hasMove = false;

        foreach (XElement command in path.Elements())
        {
            if (command.Name == DrawingNamespace + "moveTo")
            {
                if (hasMove)
                {
                    return false;
                }

                current = ReadCustomGeometryPoint(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                hasMove = true;
            }
            else if (command.Name == DrawingNamespace + "lnTo" && current is { } lineStart)
            {
                (double X, double Y) end = ReadCustomGeometryPoint(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                segments.Add(CreateLineBezierSegment(lineStart, end));
                current = end;
            }
            else if (command.Name == DrawingNamespace + "cubicBezTo" && current is { } cubicStart)
            {
                List<(double X, double Y)> points = ReadCustomGeometryPoints(command, x, y, width, height, coordinateWidth, coordinateHeight, guides);
                if (points.Count != 3)
                {
                    return false;
                }

                segments.Add(new BezierSegment(cubicStart.X, cubicStart.Y, points[0].X, points[0].Y, points[1].X, points[1].Y, points[2].X, points[2].Y));
                current = points[2];
            }
            else if (command.Name == DrawingNamespace + "quadBezTo" && current is { } quadStart)
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

    private static double ParseOptionalDoubleAttribute(XElement element, string name, double defaultValue)
    {
        return element.Attribute(name) is { } value &&
            double.TryParse(value.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            ? parsed
            : defaultValue;
    }

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
            ? lineWidth * OfficeStraightStealthLineEndLengthFactor * style.LengthScale
            : Math.Max(6.5d, lineWidth * 4d) * style.LengthScale;
        double markerWidth = style.Kind == LineEndKind.Stealth
            ? lineWidth * OfficeStraightStealthLineEndWidthFactor * style.WidthScale
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
            ? lineWidth * OfficeStraightStealthLineEndLengthFactor * headEnd.LengthScale * OfficeStraightStealthLineEndNotchFactor
            : 0d;
        double tailInset = tailEnd.Kind == LineEndKind.Stealth
            ? lineWidth * OfficeStraightStealthLineEndLengthFactor * tailEnd.LengthScale * OfficeStraightStealthLineEndNotchFactor
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
        double markerLength = lineWidth * OfficeStraightStealthLineEndLengthFactor * style.LengthScale;
        double markerWidth = lineWidth * OfficeStraightStealthLineEndWidthFactor * style.WidthScale;

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
        double arrowLength = lineWidth * OfficeStraightTriangleLineEndLengthFactor;
        double arrowHalfWidth = lineWidth * OfficeStraightTriangleLineEndHalfWidthFactor;
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
            AppendOfficeArrowHeadPath(graphics, x1, y1, ux, uy, nx, ny, lineWidth);
        }

        if (tailArrow)
        {
            AppendOfficeArrowHeadPath(graphics, x2, y2, -ux, -uy, -nx, -ny, lineWidth);
        }

        graphics.FillCurrentPath();
    }

    private static void AppendOfficeArrowHeadPath(PdfGraphicsBuilder graphics, double tipX, double tipY, double ux, double uy, double nx, double ny, double lineWidth, bool splitTrailingCurve = false)
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
                AppendClosedLinePath(graphics, tailSubpath);
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

    private static void AppendClosedLinePath(PdfGraphicsBuilder graphics, IReadOnlyList<(double X, double Y)> points, bool explicitClosingLine = false)
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
            AppendOfficeArrowHeadPath(graphics, tipX, tipY, ux, uy, nx, ny, lineWidth);
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

    private static ShapePatternFill? ToShapePatternFill(PptxScenePatternFill fill)
    {
        return fill.HasPattern
            ? new ShapePatternFill(fill.Preset, fill.Foreground, fill.Background, fill.Alpha)
            : null;
    }

    private static GradientFill? ToGradientFill(PptxSceneGradientFill fill)
    {
        return fill.HasGradient
            ? new GradientFill(fill.AngleDegrees, fill.Stops.Select(stop => new GradientStop(stop.Offset, stop.Color, stop.Alpha)).ToArray())
            : null;
    }

    private static ShapePictureFill? ToShapePictureFill(PptxSceneShapePictureFill fill)
    {
        return fill.HasPicture
            ? new ShapePictureFill(fill.RelationshipId, fill.TargetPartName, fill.Resource, ToCropRect(fill.Crop), ToFillRect(fill.Fill), fill.Alpha)
            : null;
    }

    private static Glow? ToGlow(PptxSceneGlow glow)
    {
        return glow.HasGlow
            ? new Glow(glow.Color, glow.Alpha, glow.Radius)
            : null;
    }

    private static OuterShadow? ToOuterShadow(PptxSceneOuterShadow shadow)
    {
        return shadow.HasShadow
            ? new OuterShadow(shadow.Color, shadow.Alpha, shadow.OffsetX, shadow.OffsetY)
            : null;
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
        return ToGroupTransform(PptxSceneBuilder.ReadGroupTransform(group));
    }

    private static GroupTransform ToGroupTransform(PptxSceneGroupTransform group)
    {
        return new GroupTransform(
            group.OffsetX,
            group.OffsetY,
            group.Width,
            group.Height,
            group.ChildOffsetX,
            group.ChildOffsetY,
            group.ScaleX,
            group.ScaleY,
            group.RotationDegrees,
            group.FlipHorizontal,
            group.FlipVertical);
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

        double? styleLineWidth = TryReadStyleLineWidth(shape, theme, out double inheritedLineWidth)
            ? inheritedLineWidth
            : null;
        PptxFormatSchemeReference lineReference = PptxFormatSchemeResolver.ResolveLineReference(shape, theme);
        if (explicitLine is not null && TryReadLineWithAlpha(shapeProperties, theme, out color, out lineWidth, out alpha, styleLineWidth))
        {
            return true;
        }

        if (explicitLine?.Attribute("w") is { } explicitWidthAttribute &&
            lineReference.Style is not null &&
            TryReadSolidColorWithAlpha(lineReference.Style, theme, lineReference.Reference, out color, out alpha))
        {
            lineWidth = OoxUnits.EmuToPoints(long.Parse(explicitWidthAttribute.Value, CultureInfo.InvariantCulture));
            return true;
        }

        if (lineReference.Style is null)
        {
            color = default;
            lineWidth = 0d;
            alpha = 1d;
            return false;
        }

        lineWidth = lineReference.Style.Attribute("w") is { } widthAttribute
            ? OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture))
            : 1d;
        return TryReadSolidColorWithAlpha(lineReference.Style, theme, lineReference.Reference, out color, out alpha);
    }

    private static bool TryReadStyleLineWidth(XElement shape, PptxTheme theme, out double lineWidth)
    {
        PptxFormatSchemeReference lineReference = PptxFormatSchemeResolver.ResolveLineReference(shape, theme);
        if (lineReference.Style?.Attribute("w") is { } widthAttribute)
        {
            lineWidth = OoxUnits.EmuToPoints(long.Parse(widthAttribute.Value, CultureInfo.InvariantCulture));
            return true;
        }

        lineWidth = 0d;
        return false;
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, out RgbColor color)
    {
        return TryReadShapeFontColor(shape, theme, PptxColorMap.Default, out color);
    }

    private static bool TryReadShapeFontColor(XElement shape, PptxTheme theme, PptxColorMap colorMap, out RgbColor color)
    {
        XElement? fontRef = shape
            .Element(PresentationNamespace + "style")
            ?.Element(DrawingNamespace + "fontRef");
        return PptxColorResolver.TryReadSolidColor(fontRef, theme, colorMap, out color);
    }
}
