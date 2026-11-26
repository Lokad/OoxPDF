using System.Globalization;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
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
    private const double OfficeStraightConnectorTriangleLineEndHalfWidthFactor = 2d;
    private const double OfficeStraightTriangleLineEndOverlapFactor = 2d / 3d;
    private const double OfficeStraightStealthLineEndLengthFactor = 3d;
    private const double OfficeStraightStealthLineEndWidthFactor = 3d;
    private const double OfficeStraightStealthLineEndMinimumSize = 6d;
    private const double OfficeStraightStealthLineEndNotchFactor = 2d / 3d;
    private const double OfficePresetArcDefaultStartGuide = 16200000d;
    private const double OfficePresetArcDefaultEndGuide = 0d;
    private const double OfficePresetArcFlatteningTolerance = 0.015d;
    private const double OfficePresetArcMaximumFlatteningDegrees = 4.5d;
    private const double OfficeShapeEffectRasterPixelsPerPoint = 1d;
    private const int OfficeShapeEffectRasterMaxPixelsPerSide = 2048;
    private const double OfficeOuterShadowRasterExtentFactor = 1.25d;

    private static void RenderBackground(PptxRenderContext context, PdfGraphicsBuilder graphics)
    {
        // Backgrounds are inherited alternatives, not composited source layers.
        // A transparent override must not reveal the overridden master/layout.
        PptxSceneBackground background = context.SceneSlide.BackgroundCandidates().Last().Background;
        graphics.SaveState();
        ClipSlideBoundsEvenOdd(context.Document, graphics);
        double width = context.Document.SlideWidthPoints;
        double height = context.Document.SlideHeightPoints;
        if (background.GradientFill is { } sourceGradient && ToGradientFill(sourceGradient) is { } gradient)
        {
            DrawLinearGradientFill(graphics, gradient, 0d, 0d, width, height);
        }
        else
        {
            if (background.HasFill)
            {
                if (background.Alpha < 0.999d)
                {
                    graphics.SetAlpha(background.Alpha, 1d);
                }

                graphics.SetFillRgb(background.Color.Red, background.Color.Green, background.Color.Blue);
            }
            else
            {
                graphics.SetFillRgb(255, 255, 255);
            }
            graphics.FillRectangleEvenOdd(0, 0, width, height);
        }
        graphics.RestoreState();
    }

    private static PdfRectangle? RenderShape(
        PptxSceneNode shape,
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        GroupTransform groupTransform,
        List<PdfImageResource>? images,
        Dictionary<string, RasterPicture>? imageCache,
        ref int imageIndex,
        CancellationToken cancellationToken)
    {
        if (shape.Shape is null || shape.Bounds is null)
        {
            return null;
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (shape.Source.Element(PresentationNamespace + "spPr") is null)
        {
            return null;
        }

        return RenderShape(
            document,
            graphics,
            diagnosticSink,
            slideIndex,
            groupTransform,
            images,
            imageCache,
            ref imageIndex,
            ToShapeBounds(shape.Bounds),
            shape.Shape.Preset,
            shape.Shape.PresetAdjustments,
            shape.Shape.HasCustomGeometry,
            shape.Shape.CustomGeometry,
            shape.Shape.Fill,
            ToGradientFill(shape.Shape.GradientFill),
            shape.Shape.PatternFill,
            ToShapePictureFill(shape.Shape.PictureFill),
            ToGlow(shape.Shape.Glow),
            ToOuterShadow(shape.Shape.OuterShadow),
            ToLineStyle(shape.Shape.Line),
            shape.Shape.HeadEnd,
            shape.Shape.TailEnd,
            shape.HyperlinkClick.IsDefined,
            cancellationToken);
    }

    private static PdfRectangle? RenderShape(
        PptxDocument document,
        PdfGraphicsBuilder graphics,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex,
        GroupTransform groupTransform,
        List<PdfImageResource>? images,
        Dictionary<string, RasterPicture>? imageCache,
        ref int imageIndex,
        ShapeBounds rawBounds,
        string preset,
        IReadOnlyDictionary<string, double>? presetAdjustments,
        bool hasCustomGeometry,
        PptxSceneCustomGeometry? customGeometryOverride,
        PptxSceneFillStyle fillStyle,
        GradientFill? gradientFillOverride,
        PptxScenePatternFill patternFill,
        ShapePictureFill? pictureFillOverride,
        Glow? glowOverride,
        OuterShadow? outerShadowOverride,
        LineStyle line,
        PptxSceneLineEnd headEnd,
        PptxSceneLineEnd tailEnd,
        bool hasHyperlink,
        CancellationToken cancellationToken)
    {
        ShapeBounds bounds = groupTransform.Apply(rawBounds);

        double x = OoxUnits.EmuToPoints(bounds.X);
        double yTop = OoxUnits.EmuToPoints(bounds.Y);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - yTop - height;
        bool transformed = bounds.RotationDegrees != 0d || bounds.FlipHorizontal || bounds.FlipVertical;

        RgbColor fill = fillStyle.Color;
        double fillAlpha = fillStyle.Alpha;
        bool hasFill = fillStyle.HasFill;
        GradientFill? gradientFill = gradientFillOverride;
        RgbColor stroke = line.Color;
        double lineWidth = line.Width;
        double strokeAlpha = line.Alpha;
        bool hasStroke = line.HasLine;
        bool hasDash = line.HasDash;
        IReadOnlyList<double> dashPattern = line.DashPattern;
        int? lineCap = line.Cap;
        int? lineJoin = line.Join;
        PptxSceneCustomGeometry? sceneCustomGeometry = customGeometryOverride is { HasGeometry: true } ? customGeometryOverride : null;
        PptxSceneCustomPath? customPictureMask = null;
        PdfImageResource? pictureResource = null;
        if (images is not null && pictureFillOverride is { } pictureFill &&
            (sceneCustomGeometry is null
                ? CanRenderPictureFillPreset(preset)
                : TryResolveCustomPictureFillPath(sceneCustomGeometry, out customPictureMask) && customPictureMask is not null))
        {
            pictureResource = PrepareShapePictureFill(pictureFill, diagnosticSink, slideIndex, imageCache, ref imageIndex, cancellationToken);
        }

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
            DrawRasterGlow(graphics, x, y, width, height, glow, images, ref imageIndex, cancellationToken);
        }

        if (hasOuterShadow &&
            preset is not ("line" or "straightConnector1" or "curvedConnector2" or "curvedConnector3") &&
            sceneCustomGeometry is null && !hasCustomGeometry)
        {
            if (outerShadow.BlurRadius > 0d &&
                preset == "rect" &&
                images is not null)
            {
                DrawRasterOuterShadow(graphics, x, y, width, height, outerShadow, images, ref imageIndex,
                    cancellationToken: cancellationToken);
            }
            else
            {
                DrawOuterShadow(graphics, preset, x, y, width, height, outerShadow, presetAdjustments);
            }
        }

        if (sceneCustomGeometry is not null)
        {
            bool paintSolidFill = hasFill;
            if (customPictureMask is not null && pictureResource is not null &&
                pictureFillOverride is { } customFill && images is not null)
            {
                RenderShapePictureFill(graphics, images, pictureResource, customFill, preset,
                    x, y, width, height, presetAdjustments, (sceneCustomGeometry, customPictureMask));
                paintSolidFill = false;
            }

            RenderCustomGeometry(
                sceneCustomGeometry,
                graphics,
                x,
                y,
                width,
                height,
                paintSolidFill,
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
                tailEnd,
                cancellationToken);
            if (transformed)
            {
                graphics.RestoreState();
            }

            return null;
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
                bool hasStealthEnd = headEnd.Kind == PptxSceneLineEndKind.Stealth || tailEnd.Kind == PptxSceneLineEndKind.Stealth;
                double triangleHalfWidthFactor = preset == "straightConnector1"
                    ? OfficeStraightConnectorTriangleLineEndHalfWidthFactor
                    : OfficeStraightTriangleLineEndHalfWidthFactor;
                // Office fills stealth-tailed straight bodies (seven renders); unflipped straight
                // connectors take the filled body too while transformed (grouped or rotated)
                // shapes keep the legacy stroked path for lack of Office evidence there.
                bool useFilledStealthBody = string.Equals(preset, "line", StringComparison.Ordinal) ||
                    rawBounds.FlipHorizontal ||
                    rawBounds.FlipVertical ||
                    (string.Equals(preset, "straightConnector1", StringComparison.Ordinal) && !transformed);
                if (hasStealthEnd && useFilledStealthBody && headEnd.Kind is PptxSceneLineEndKind.None or PptxSceneLineEndKind.Stealth && tailEnd.Kind is PptxSceneLineEndKind.None or PptxSceneLineEndKind.Stealth && !hasDash && lineCap is null)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    FillStealthEndedLine(graphics, x1, y1, x2, y2, lineWidth, headEnd, tailEnd);
                }
                else if ((hasHeadArrow || hasTailArrow) && headEnd.Kind is PptxSceneLineEndKind.None or PptxSceneLineEndKind.Triangle or PptxSceneLineEndKind.Arrow && tailEnd.Kind is PptxSceneLineEndKind.None or PptxSceneLineEndKind.Triangle or PptxSceneLineEndKind.Arrow && !hasDash && lineCap is null)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    bool usesOfficeArrowType = headEnd.Kind == PptxSceneLineEndKind.Arrow || tailEnd.Kind == PptxSceneLineEndKind.Arrow;
                    FillStraightArrowedLine(graphics, x1, y1, x2, y2, lineWidth,
                        hasHeadArrow, hasTailArrow, usesOfficeArrowType, triangleHalfWidthFactor);
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
                    FillLineEndMarker(graphics, headEnd, x1, y1, (x1 - x2, y1 - y2), lineWidth, triangleHalfWidthFactor);
                    FillLineEndMarker(graphics, tailEnd, x2, y2, (x2 - x1, y2 - y1), lineWidth, triangleHalfWidthFactor);
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

            return CollapsedStrokeLinkBounds(lineCap);
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
                    tailEnd.Kind is PptxSceneLineEndKind.Triangle or PptxSceneLineEndKind.Arrow or PptxSceneLineEndKind.Stealth)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                    if (!TryFillCurvedConnectorPreset(
                        graphics,
                        preset,
                        x,
                        yTop,
                        width,
                        height,
                        document.SlideHeightPoints,
                        lineWidth,
                        tailEnd,
                        presetAdjustments))
                    {
                        DrawCurvedConnectorPreset(
                            graphics,
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
                            presetAdjustments);
                    }
                }
                else
                {
                    DrawCurvedConnectorPreset(
                        graphics,
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
                        presetAdjustments);
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

            return CollapsedStrokeLinkBounds(lineCap);
        }

        bool repeatSlideClipBeforeFill = hasStroke && hasFill && !transformed;
        if (repeatSlideClipBeforeFill)
        {
            ClipSlideBoundsEvenOdd(document, graphics);
        }

        if (gradientFill is not null)
        {
            graphics.SaveState();
            ClipToPresetShape(graphics, preset, x, y, width, height, presetAdjustments);

            DrawLinearGradientFill(graphics, gradientFill, x, y, width, height);

            graphics.RestoreState();
        }
        else if (pictureResource is not null && pictureFillOverride is { } presetFill && images is not null)
        {
            RenderShapePictureFill(graphics, images, pictureResource, presetFill, preset,
                x, y, width, height, presetAdjustments);
        }
        else if (patternFill.HasPattern)
        {
            graphics.SetFillRgb(patternFill.Background.Red, patternFill.Background.Green, patternFill.Background.Blue);
            DrawPresetFill(graphics, preset, x, y, width, height, presetAdjustments);
            StrokeShapePatternFill(graphics, preset, x, y, width, height, patternFill, presetAdjustments);
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
            DrawPresetFill(graphics, preset, x, y, width, height, presetAdjustments);

            if (transparentFill)
            {
                graphics.RestoreState();
            }
        }

        bool strokedFrame = false;
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
                bool useFilledLineEndOutline =
                    !hasDash &&
                    lineCap is null &&
                    headEnd.IsNone &&
                    tailEnd.Kind is PptxSceneLineEndKind.Triangle or PptxSceneLineEndKind.Arrow or PptxSceneLineEndKind.Stealth;
                if (useFilledLineEndOutline)
                {
                    graphics.SetFillRgb(stroke.Red, stroke.Green, stroke.Blue);
                }

                if (!useFilledLineEndOutline ||
                    !TryFillPresetArcLineEndOutline(graphics, x, y, width, height, lineWidth, tailEnd, presetAdjustments))
                {
                    DrawPresetArcStroke(graphics, x, y, width, height, presetAdjustments);
                }
            }
            else
            {
                strokedFrame = StrokePresetShape(graphics, preset, x, y, width, height, presetAdjustments);
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

        return strokedFrame ? CollapsedStrokeLinkBounds(cap: null) : null;

        PdfRectangle? CollapsedStrokeLinkBounds(int? cap)
        {
            if (!hasHyperlink || !hasStroke || strokeAlpha <= 0d || lineWidth <= 0d ||
                !((width == 0d && height > 0d) || (width > 0d && height == 0d)))
            {
                return null;
            }

            // A stroke supplies area around the collapsed frame. Use the
            // original frame as the pivot for the paint and link transforms.
            double halfStroke = lineWidth / 2d;
            double endCap = cap is 1 or 2 ? halfStroke : 0d;
            double padX = width == 0d ? halfStroke : endCap;
            double padY = height == 0d ? halfStroke : endCap;
            var frame = new PdfRectangle(x - padX, y - padY, width + 2d * padX, height + 2d * padY);
            return transformed ? TransformedRectangleBounds(frame, ShapeTransformMatrix(x, y, width, height, bounds)) : frame;
        }
    }
}
