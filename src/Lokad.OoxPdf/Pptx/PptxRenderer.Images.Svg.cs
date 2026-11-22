using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool TryRenderSvgPicture(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, byte[] bytes, CropRect crop, FillRect fillRect, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex, string? partName, CancellationToken cancellationToken,
        string picturePreset, IReadOnlyDictionary<string, double> pictureAdjustments,
        IReadOnlyDictionary<FontRequest, SvgTextFont?>? textFonts, double pictureOpacity)
    {
        double x = OoxUnits.EmuToPoints(bounds.X);
        double width = OoxUnits.EmuToPoints(bounds.Width);
        double height = OoxUnits.EmuToPoints(bounds.Height);
        double y = document.SlideHeightPoints - OoxUnits.EmuToPoints(bounds.Y) - height;
        return RenderSvgPictureCore(graphics, x, y, width, height, bounds, bytes, crop, fillRect, diagnosticSink, slideIndex, partName, cancellationToken, picturePreset, pictureAdjustments, textFonts, pictureOpacity);
    }

    // The vector paint pipeline owns no presentation state. DOCX supplies its
    // already resolved PDF rectangle and maps diagnostics to the document page.
    internal static bool TryRenderSvgImage(PdfGraphicsBuilder graphics, double x, double y, double width, double height,
        byte[] bytes, Action<OoxPdfDiagnostic>? diagnosticSink, int pageIndex, string? partName, CancellationToken cancellationToken,
        IReadOnlyDictionary<FontRequest, SvgTextFont?>? textFonts = null)
    {
        Action<OoxPdfDiagnostic>? pageSink = diagnosticSink is null ? null :
            diagnostic => diagnosticSink(diagnostic with { SlideIndex = null, PageIndex = pageIndex });
        graphics.SaveState();
        try
        {
            graphics.ClipRectangle(x, y, width, height);
            return RenderSvgPictureCore(graphics, x, y, width, height, default, bytes, default, default,
                pageSink, 0, partName, cancellationToken, textFonts: textFonts, hasViewportClip: true);
        }
        finally { graphics.RestoreState(); }
    }

    private static bool RenderSvgPictureCore(PdfGraphicsBuilder graphics, double x, double y, double width, double height,
        ShapeBounds bounds, byte[] bytes, CropRect crop, FillRect fillRect, Action<OoxPdfDiagnostic>? diagnosticSink,
        int slideIndex, string? partName, CancellationToken cancellationToken,
        string picturePreset = "rect", IReadOnlyDictionary<string, double>? pictureAdjustments = null,
        IReadOnlyDictionary<FontRequest, SvgTextFont?>? textFonts = null, double pictureOpacity = 1d, bool hasViewportClip = false)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // Coarse render-local DOM, gradient/path models and working headroom.
        // Encoded input is borrowed; retained PDF content/resources have their
        // separate lifetime and are not funded by this temporary reservation.
        using IMemoryReservation? scratch = OoxMemory.Reserve(OoxConversionBudget.Current?.Memory,
            checked(64L * 1024 + 32L * bytes.Length));
        cancellationToken.ThrowIfCancellationRequested();
        XDocument svg;
        try
        {
            using (var stream = new MemoryStream(bytes))
            {
                svg = SafeXml.Load(stream, cancellationToken);
            }
        }
        catch (InvalidDataException ex)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "SVG_UNSUPPORTED_CONTENT",
                OoxPdfSeverity.Error,
                "SVG picture could not be parsed and was ignored: " + ex.Message,
                partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "svg",
                Fallback: "Ignored"));
            return false;
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (!TryReadSvgSourceFrame(svg.Root, out double minX, out double minY, out double viewWidth, out double viewHeight) ||
            viewWidth <= 0d ||
            viewHeight <= 0d)
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "SVG_UNSUPPORTED_CONTENT",
                OoxPdfSeverity.Error,
                "SVG picture has no usable viewBox and was ignored.",
                partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "svg",
                Fallback: "Ignored"));
            return false;
        }


        double imageX = x + fillRect.Left * width;
        double imageY = y + fillRect.Bottom * height;
        double imageWidth = Math.Max(0.001d, width * (1d - fillRect.Left - fillRect.Right));
        double imageHeight = Math.Max(0.001d, height * (1d - fillRect.Top - fillRect.Bottom));
        double sourceLeft = Math.Clamp(crop.Left, 0d, 0.999d);
        double sourceTop = Math.Clamp(crop.Top, 0d, 0.999d);
        double sourceRight = Math.Clamp(crop.Right, 0d, 0.999d);
        double sourceBottom = Math.Clamp(crop.Bottom, 0d, 0.999d);
        double sourceMinX = minX + sourceLeft * viewWidth;
        double sourceMinY = minY + sourceTop * viewHeight;
        double sourceWidth = viewWidth * Math.Max(0.001d, 1d - sourceLeft - sourceRight);
        double sourceHeight = viewHeight * Math.Max(0.001d, 1d - sourceTop - sourceBottom);
        double scaleX = imageWidth / sourceWidth;
        double scaleY = imageHeight / sourceHeight;
        // RV07: subnormal viewBox dimensions overflow the picture scale past
        // double range; the picture cannot map and is ignored like a missing
        // viewBox instead of throwing non-finite PDF numbers downstream.
        if (!double.IsFinite(scaleX) || !double.IsFinite(scaleY))
        {
            diagnosticSink?.Invoke(new OoxPdfDiagnostic(
                "SVG_UNSUPPORTED_CONTENT",
                OoxPdfSeverity.Error,
                "SVG picture has no usable viewBox and was ignored.",
                partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "svg",
                Fallback: "Ignored"));
            return false;
        }

        SvgTransform? picturePlacement = null;
        graphics.SaveState();
        if (Math.Abs(bounds.RotationDegrees) > 0.001d || bounds.FlipHorizontal || bounds.FlipVertical)
        {
            var (a, b, c, d, e, f) = ShapeTransformMatrix(x, y, width, height, bounds);
            graphics.Transform(a, b, c, d, e, f);
            static double Coefficient(double value) => double.Parse(
                PdfDocumentWriter.FormatNumber(value, PdfGraphicsBuilder.TransformCoefficientFormat), CultureInfo.InvariantCulture);
            picturePlacement = new SvgTransform(Coefficient(a), Coefficient(b), Coefficient(c), Coefficient(d),
                PrintedSvgNumber(e), PrintedSvgNumber(f));
        }

        if (picturePreset != "rect")
            ClipToPresetShape(graphics, null, picturePreset, x, y, width, height, pictureAdjustments);

        string? overflow = !hasViewportClip && svg.Root is { } outerSvg
            ? ReadSvgPresentationAttribute(outerSvg, ReadSvgStyleDeclarations(outerSvg), "overflow")?.Trim()
            : null;
        bool clipViewport = !crop.IsEmpty || picturePlacement.HasValue ||
            string.Equals(overflow, "hidden", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(overflow, "scroll", StringComparison.OrdinalIgnoreCase);
        if (clipViewport)
        {
            // The viewport rotates with the picture, matching the raster path.
            graphics.ClipRectangle(imageX, imageY, imageWidth, imageHeight);
        }

        if (pictureOpacity <= 0d || svg.Root is { } root && IsFullyTransparentSvgContainer(root))
        {
            graphics.RestoreState();
            return true;
        }

        var (gradients, radialGradients) = ReadSvgGradients(svg, viewWidth, viewHeight, cancellationToken);
        double strokePercentageBasis = GetSvgNormalizedDiagonal(viewWidth, viewHeight);
        var elementIds = CreateSvgElementIdLookup(svg, cancellationToken);
        var paintedText = new HashSet<XElement>();
        bool hasStylesheet = HasSvgStylesheet(svg);
        var unsupportedCommands = new SortedSet<char>();
        int unreadablePaths = 0;
        var missingGradients = new SortedSet<string>(StringComparer.Ordinal);
        int unpaintablePaths = 0;
        int unsupportedTransforms = 0;
        int gradientStrokes = 0;
        int unpaintableStrokes = 0;
        int invalidStrokePresentations = 0;
        int focalRadialGradients = 0;
        int nonFiniteGradientPaths = 0;
        int vectorEffectStrokes = 0;
        HashSet<string>? usedGradientIds = diagnosticSink is null ? null : new(StringComparer.Ordinal);
        var isolation = new List<(XElement Container, PdfGraphicsBuilder Parent, PdfGraphicsBuilder Child, double Opacity)>();
        var opacityGroups = new List<(XElement Container, double Opacity)>();
        var admittedContainers = new HashSet<XElement>();
        var admittedGradientAlpha = new HashSet<string>(StringComparer.Ordinal);
        var fallbackGradientAlpha = new HashSet<string>(StringComparer.Ordinal);
        PdfGraphicsBuilder pageGraphics = graphics;
        int precedingCharacters = graphics.ContentLength;
        long sampledBytes = 0;
        void AdmitSamples(int samples)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // Retained generated paint plus builder/snapshot overlap. Keep this
            // image's total across opacity groups; each new picture starts fresh.
            sampledBytes = checked(sampledBytes + 1024L * samples);
            pageGraphics.AdmitSvgConstruction(sampledBytes, bytes.Length, precedingCharacters);
            cancellationToken.ThrowIfCancellationRequested();
        }
        var pictureFrame = new PdfRectangle(imageX, imageY, imageWidth, imageHeight);
        void OpenGroup(XElement owner, double opacity)
        {
            var child = new PdfGraphicsBuilder();
            isolation.Add((owner, graphics, child, opacity));
            graphics = child;
        }
        void CloseGroup(PdfRectangle? paintBounds = null)
        {
            var entry = isolation[^1];
            isolation.RemoveAt(isolation.Count - 1);
            graphics = entry.Parent;
            graphics.DrawTransparencyGroup(entry.Child, paintBounds ?? pictureFrame, entry.Opacity);
        }
        foreach (XElement element in svg.Descendants().Where(candidate => IsSvgPaintableElement(candidate) || candidate.Name.LocalName == "text"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            // RV07-D2: definition descendants supply reusable paint resources;
            // they do not paint into the picture or activate paint diagnostics.
            if (IsSvgDefinitionElement(element))
            {
                continue;
            }

            IReadOnlyDictionary<string, string> style = ReadSvgStyleDeclarations(element);
            if (IsHiddenSvgElement(element, style) ||
                !TryCollectSvgOpacityGroups(element, pictureOpacity, opacityGroups, out bool strokeContainersQualified))
            {
                continue;
            }

            int shared = 0;
            while (shared < isolation.Count && shared < opacityGroups.Count &&
                ReferenceEquals(isolation[shared].Container, opacityGroups[shared].Container)) { shared++; }
            while (isolation.Count > shared) { CloseGroup(); }
            for (int i = shared; i < opacityGroups.Count; i++)
            {
                var (container, opacity) = opacityGroups[i];
                admittedContainers.Add(container);
                OpenGroup(container, opacity);
            }
            if (element.Name.LocalName == "text")
            {
                if (!hasStylesheet && TryReadSvgPathTransform(element, out SvgTransform textTransform) &&
                    TryPaintSvgText(graphics, element, elementIds, textFonts, textTransform, sourceMinX, sourceMinY,
                        imageX, imageY, imageWidth, imageHeight, scaleX, scaleY, AdmitSamples, cancellationToken))
                    paintedText.Add(element);
                continue;
            }
            string? data = ConvertSvgShapeToPathData(element, viewWidth, viewHeight);
            if (string.IsNullOrWhiteSpace(data))
            {
                continue;
            }
            if (!TryReadSvgPathTransform(element, out SvgTransform transform))
            {
                unsupportedTransforms++;
                continue;
            }
            char? badCommand = FindFirstUnsupportedSvgPathCommand(data);
            if (badCommand is not null)
            {
                unsupportedCommands.Add(badCommand.Value);
            }
            string? opacityText = ReadSvgPresentationAttribute(element, style, "opacity");
            double? authoredOpacity = ParseSvgOpacity(opacityText);
            double elementOpacity = Math.Clamp(authoredOpacity ?? 1d, 0d, 1d);
            bool allowGradientStroke = strokeContainersQualified &&
                (string.IsNullOrWhiteSpace(opacityText) || authoredOpacity is >= 0d and <= 1d);
            bool hasFill = TryReadSvgFill(element, style, gradients, radialGradients, elementIds, elementOpacity,
                out SvgPaint paint, out SvgFillFailure fillFailure, out string? gradientId, out SvgRadialGradient? radial);
            SvgStroke stroke = ReadSvgStroke(element, style, gradients, radialGradients, elementIds,
                allowGradientStroke, elementOpacity, strokePercentageBasis, out SvgStrokeFailure strokeFailure,
                out bool invalidStrokePresentation, out bool hasVectorEffect, out bool nonScalingStroke);

            (IReadOnlyList<SvgPathCommand> pathCommands, bool completePath) =
                hasFill || stroke.HasPaint ? ReadTransformedSvgPath(data, transform) : ([], false);
            if ((hasFill || fillFailure != SvgFillFailure.None) && completePath && !SvgPathMayHaveInterior(pathCommands))
            {
                hasFill = false;
                paint = default;
                fillFailure = SvgFillFailure.None;
                gradientId = null;
                radial = null;
            }
            // A combined PDF fill-and-stroke operation knocks out the fill
            // under the stroke. A translucent SVG stroke paints over its fill.
            bool separateStroke = stroke.Gradient is not null || paint.Color is null ||
                paint.Opacity > 0d && stroke.Opacity != elementOpacity;
            bool gradientFill = radial is not null || paint.Gradient is not null;
            bool isolateObject = separateStroke && stroke.HasPaint && stroke.Opacity > 0d &&
                hasFill && (!gradientFill || completePath) && paint.Opacity > 0d && elementOpacity < 1d;
            if (isolateObject)
            {
                // Leave room for the fill mask when adding an object group.
                bool fillNeedsMask = paint.Gradient is { UniformStopOpacity: null } ||
                    radial is { UniformStopOpacity: null };
                if (isolation.Count >= PdfTransparencyGroup.MaxDepth - (fillNeedsMask ? 1 : 0))
                {
                    if (stroke.Gradient is not null)
                    {
                        stroke = default;
                        strokeFailure = SvgStrokeFailure.UnresolvedGradient;
                    }
                    else
                    {
                        separateStroke = paint.Color is null;
                    }
                    isolateObject = false;
                }
            }

            if (hasFill && gradientId is not null)
            {
                usedGradientIds?.Add(gradientId);
            }
            if (!hasFill)
            {
                if (fillFailure == SvgFillFailure.UnresolvedGradient && gradientId is not null)
                {
                    missingGradients.Add(gradientId);
                }
                else if (fillFailure == SvgFillFailure.UnparsableColor)
                {
                    unpaintablePaths++;
                }
            }
            SvgPathBounds gradientBounds = default;
            SvgPathBounds controlBounds = default;
            bool needObjectBounds = isolateObject && !hasViewportClip && !clipViewport;
            bool hasPathBounds = (needObjectBounds || completePath && (gradientFill || stroke.Gradient is not null)) &&
                TryReadSvgPathBounds(pathCommands, out gradientBounds, out controlBounds);
            bool hasGradientBounds = completePath && hasPathBounds;
            bool controlsMappable = hasPathBounds &&
                AreSvgGradientBoundsMappable(controlBounds, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY);
            PdfShadingPattern? strokePattern = null;
            if (stroke.Gradient is { } strokeGradient)
            {
                if (hasGradientBounds && controlsMappable)
                {
                    // A Form inherits placement through its invocation. A page
                    // pattern maps to the initial user space before the picture cm.
                    SvgTransform? patternPlacement = isolation.Count == 0 && !isolateObject ? picturePlacement : null;
                    strokePattern = TryReadSvgStrokePattern(gradientBounds, strokeGradient, transform, sourceMinX, sourceMinY,
                        imageX, imageY, imageHeight, scaleX, scaleY, patternPlacement);
                }
                if (strokePattern is null)
                {
                    stroke = default;
                    strokeFailure = SvgStrokeFailure.UnresolvedGradient;
                }
                else if (stroke.GradientId is { } strokeGradientId)
                {
                    usedGradientIds?.Add(strokeGradientId);
                }
            }
            if (strokeFailure == SvgStrokeFailure.UnresolvedGradient)
            {
                gradientStrokes++;
            }
            else if (strokeFailure != SvgStrokeFailure.None)
            {
                unpaintableStrokes++;
            }
            if (stroke.HasPaint && invalidStrokePresentation)
            {
                invalidStrokePresentations++;
            }
            if (stroke.HasPaint && hasVectorEffect)
            {
                vectorEffectStrokes++;
            }
            if (!hasFill && !stroke.HasPaint)
            {
                continue;
            }
            bool evenOddFill = ReadSvgInheritedAttribute(element, style, "fill-rule")?.Trim().Equals("evenodd", StringComparison.OrdinalIgnoreCase) == true;
            double viewportScale = (scaleX + scaleY) / 2d;
            double strokeScale = nonScalingStroke ? viewportScale : GetSvgStrokeElementScale(transform) * viewportScale;
            double strokeWidthPoints = Math.Max(0.001d, stroke.Width * strokeScale);
            double[]? dashPoints = null;
            double dashPhasePoints = 0d;
            if (stroke.DashPattern is { } userDash)
            {
                dashPoints = new double[userDash.Length];
                for (int dashIndex = 0; dashIndex < userDash.Length; dashIndex++)
                {
                    dashPoints[dashIndex] = userDash[dashIndex] * strokeScale;
                }
                // RV07-S5: Office square caps preserve the painted dash/gap
                // footprint. PDF caps extend each segment by a full width.
                // Retain the approximation when a dash is no longer than that
                // width: zero/negative segments do not match Office's paint.
                // Gradient strokes follow Office's cap extension into the gaps;
                // keep the earlier compensation specific to solid stroke paint.
                bool compensateSquareCaps = stroke.LineCap == 2 && stroke.Gradient is null;
                for (int dashIndex = 0; compensateSquareCaps && dashIndex < dashPoints.Length; dashIndex += 2)
                {
                    compensateSquareCaps = dashPoints[dashIndex] > strokeWidthPoints;
                }
                if (compensateSquareCaps)
                {
                    for (int dashIndex = 0; dashIndex < dashPoints.Length; dashIndex += 2)
                    {
                        dashPoints[dashIndex] -= strokeWidthPoints;
                        dashPoints[dashIndex + 1] += strokeWidthPoints;
                    }
                }
                // RV07-S4: Office interprets dash offset in stroke-width units,
                // while dash lengths remain user-space geometry.
                dashPhasePoints = stroke.DashOffset * strokeWidthPoints;
            }
            // RV07: stroke state past double range (huge transform area factors)
            // cannot set PDF line state; the stroke is omitted with its fill
            // intact instead of dropping the picture.
            if (stroke.HasPaint
                && (!double.IsFinite(strokeWidthPoints)
                    || !double.IsFinite(dashPhasePoints)
                    || (dashPoints is not null && !dashPoints.All(double.IsFinite))))
            {
                stroke = default;
                strokePattern = null;
                unpaintableStrokes++;
                if (!hasFill)
                {
                    continue;
                }
            }
            PdfRectangle? objectFrame = null;
            if (isolateObject && stroke.HasPaint)
            {
                objectFrame = !needObjectBounds ? pictureFrame : controlsMappable
                    ? ReadSvgObjectPaintBounds(pictureFrame, controlBounds, stroke, strokeWidthPoints,
                        sourceMinX, sourceMinY, scaleX, scaleY)
                    : null;
                if (objectFrame.HasValue)
                {
                    OpenGroup(element, elementOpacity);
                    paint = paint with { Opacity = paint.Opacity / elementOpacity };
                    stroke = stroke with { Opacity = stroke.Opacity / elementOpacity };
                }
                else if (stroke.Gradient is not null)
                {
                    stroke = default;
                    strokePattern = null;
                    gradientStrokes++;
                }
                else
                {
                    separateStroke = paint.Color is null;
                }
            }
            if (radial is { } radialGradient)
            {
                if (radialGradient.HasFocal)
                {
                    focalRadialGradients++;
                }
                bool varyingAlphaAdmitted = false;
                if (hasGradientBounds)
                {
                    if (!controlsMappable || !TryRenderSvgRadialGradientPath(graphics, pathCommands, radialGradient, transform, evenOddFill, paint.Opacity, gradientBounds, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY, isolation.Count < PdfTransparencyGroup.MaxDepth, out varyingAlphaAdmitted, AdmitSamples, cancellationToken))
                    {
                        nonFiniteGradientPaths++;
                    }
                }
                else if (badCommand is null)
                {
                    unreadablePaths++;
                }
                if (gradientId is not null)
                {
                    (varyingAlphaAdmitted ? admittedGradientAlpha : fallbackGradientAlpha).Add(gradientId);
                }
            }
            else if (paint.Gradient is { } gradient)
            {
                bool varyingAlphaAdmitted = false;
                if (hasGradientBounds)
                {
                    if (!controlsMappable || !TryRenderSvgGradientPath(graphics, pathCommands, gradient, transform, evenOddFill, paint.Opacity, gradientBounds, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY, isolation.Count < PdfTransparencyGroup.MaxDepth, out varyingAlphaAdmitted, AdmitSamples, cancellationToken))
                    {
                        nonFiniteGradientPaths++;
                    }
                }
                else if (badCommand is null)
                {
                    unreadablePaths++;
                }
                if (gradientId is not null)
                {
                    (varyingAlphaAdmitted ? admittedGradientAlpha : fallbackGradientAlpha).Add(gradientId);
                }
            }
            else if (paint.Color is { } color)
            {
                if (!separateStroke && stroke.Color is { } strokeColor)
                {
                    bool transparent = paint.Opacity < 1d || stroke.Opacity < 1d;
                    graphics.SaveState();
                    // RV07-S3: preserve one combined paint operation and both
                    // alpha values while applying the shared viewport stroke map.
                    SvgTransform? strokeViewport = ApplySvgStrokeViewport(graphics, imageX, imageY, imageHeight, scaleX, scaleY);
                    if (transparent)
                    {
                        graphics.SetAlpha(paint.Opacity, stroke.Opacity);
                    }
                    graphics.SetFillRgb(color.Red, color.Green, color.Blue);
                    graphics.SetStrokeRgb(strokeColor.Red, strokeColor.Green, strokeColor.Blue);
                    graphics.SetLineWidth(strokeWidthPoints);
                    if (stroke.LineCap != 0)
                    {
                        graphics.SetLineCap(stroke.LineCap);
                    }
                    if (stroke.LineJoin != 0)
                    {
                        graphics.SetLineJoin(stroke.LineJoin);
                    }
                    if (dashPoints is not null)
                    {
                        graphics.SetLineDash(dashPoints, dashPhasePoints);
                    }
                    if (stroke.LineJoin == 0)
                    {
                        graphics.SetMiterLimit(stroke.MiterLimit);
                    }
                    if (TryAppendSvgPath(graphics, pathCommands, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY, strokeViewport))
                    {
                        if (evenOddFill)
                        {
                            graphics.FillAndStrokeCurrentPathEvenOdd();
                        }
                        else
                        {
                            graphics.FillAndStrokeCurrentPath();
                        }
                    }
                    else if (badCommand is null)
                    {
                        unreadablePaths++;
                    }
                    graphics.RestoreState();
                }
                else
                {
                    if (paint.Opacity < 1d)
                    {
                        graphics.SaveState();
                        graphics.SetAlpha(paint.Opacity, 1d);
                    }

                    graphics.SetFillRgb(color.Red, color.Green, color.Blue);
                    if (TryAppendSvgPath(graphics, pathCommands, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY))
                    {
                        if (evenOddFill)
                        {
                            graphics.FillCurrentPathEvenOdd();
                        }
                        else
                        {
                            graphics.FillCurrentPath();
                        }
                    }
                    else if (badCommand is null)
                    {
                        unreadablePaths++;
                    }

                    if (paint.Opacity < 1d)
                    {
                        graphics.RestoreState();
                    }
                }
            }
            // Stroke patterns already require complete bounds. A solid stroke
            // can still paint the readable prefix when gradient fill is skipped.
            if (separateStroke && stroke.HasPaint)
            {
                if (!TryPaintSvgStrokePath(graphics, pathCommands, completePath, stroke.Color.GetValueOrDefault(), strokeWidthPoints, stroke.Opacity, stroke.LineCap, stroke.LineJoin, dashPoints, dashPhasePoints, stroke.MiterLimit, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY, strokePattern)
                    && badCommand is null)
                {
                    unreadablePaths++;
                }
            }
            if (objectFrame.HasValue) { CloseGroup(objectFrame); }
        }
        while (isolation.Count > 0) { CloseGroup(); }
        ReportUnsupportedSvgElements(svg, diagnosticSink, slideIndex, partName, paintedText);
        ReportSkippedSvgPaths(unsupportedCommands, unreadablePaths, missingGradients, unpaintablePaths, unsupportedTransforms, gradientStrokes, unpaintableStrokes, invalidStrokePresentations, vectorEffectStrokes, focalRadialGradients, nonFiniteGradientPaths, diagnosticSink, slideIndex, partName);
        ReportIgnoredSvgOpacity(svg, gradients, radialGradients, usedGradientIds, admittedContainers,
            admittedGradientAlpha, fallbackGradientAlpha, diagnosticSink, slideIndex, partName, cancellationToken);

        graphics.RestoreState();
        return true;
    }

    // RV07: unsupported rendered content must diagnose instead of silently
    // vanishing. One diagnostic per element kind (with occurrence count),
    // restricted to paintable content outside definitions (inert definitions
    // are correctly skipped).
    private static void ReportUnsupportedSvgElements(XDocument svg, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex, string? partName,
        HashSet<XElement> paintedText)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        var counts = new SortedDictionary<string, int>(StringComparer.Ordinal);
        foreach (XElement element in svg.Descendants())
        {
            string name = element.Name.LocalName;
            if (IsSupportedSvgElement(name) || paintedText.Contains(element) || IsSvgDefinitionElement(element) ||
                IsInFullyTransparentSvgContainer(element) || IsHiddenSvgElement(element))
            {
                continue;
            }
            counts.TryGetValue(name, out int count);
            counts[name] = count + 1;
        }
        foreach ((string name, int count) in counts)
        {
            diagnosticSink(new OoxPdfDiagnostic(
                "SVG_UNSUPPORTED_CONTENT",
                OoxPdfSeverity.Warning,
                "SVG picture omits " + count.ToString(CultureInfo.InvariantCulture) + " unsupported " + name + (count == 1 ? " element." : " elements."),
                partName,
                PageIndex: null,
                SlideIndex: slideIndex,
                Feature: "svg",
                Fallback: "Partial"));
        }
    }
    private static void ReportIgnoredSvgOpacity(XDocument svg, IReadOnlyDictionary<string, SvgGradient> gradients,
        IReadOnlyDictionary<string, SvgRadialGradient> radialGradients, HashSet<string>? usedGradientIds,
        HashSet<XElement> admittedContainers, HashSet<string> admittedGradientAlpha, HashSet<string> fallbackGradientAlpha,
        Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex, string? partName, CancellationToken cancellationToken)
    {
        if (diagnosticSink is null || usedGradientIds is null) { return; }
        int containerCount = 0;
        foreach (XElement element in svg.Descendants())
        {
            cancellationToken.ThrowIfCancellationRequested();
            string name = element.Name.LocalName;
            if (name is "svg" or "g" && !IsSvgDefinitionElement(element) && !IsInFullyTransparentSvgContainer(element))
            {
                IReadOnlyDictionary<string, string> style = ReadSvgStyleDeclarations(element);
                string? opacity = style.TryGetValue("opacity", out string? value) ? value : (string?)element.Attribute("opacity");
                if (ReadSvgOpacityValue(opacity) < 1d && !admittedContainers.Contains(element)) { containerCount++; }
            }
        }

        long stopCount = 0;
        foreach (string id in usedGradientIds)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!admittedGradientAlpha.Contains(id) || fallbackGradientAlpha.Contains(id))
            {
                stopCount += gradients.GetValueOrDefault(id)?.IgnoredStopOpacityCount ??
                    radialGradients.GetValueOrDefault(id)?.IgnoredStopOpacityCount ?? 0;
            }
        }
        if (containerCount > 0)
        {
            Report("SVG picture ignores opacity on " + containerCount.ToString(CultureInfo.InvariantCulture) + (containerCount == 1 ? " container." : " containers."));
        }
        if (stopCount > 0)
        {
            Report("SVG picture ignores stop-opacity on " + stopCount.ToString(CultureInfo.InvariantCulture) + (stopCount == 1 ? " gradient stop." : " gradient stops."));
        }
        void Report(string message) => diagnosticSink(new OoxPdfDiagnostic(
            "SVG_UNSUPPORTED_CONTENT", OoxPdfSeverity.Warning, message, partName,
            PageIndex: null, SlideIndex: slideIndex, Feature: "svg", Fallback: "Partial"));
    }

    private static SvgContainerOpacity ReadSvgContainerOpacity(XElement element)
    {
        string name = element.Name.LocalName;
        if (name is not ("svg" or "g"))
        {
            return new(null, IsTransparent: false, AllowsGradientStroke: true);
        }

        bool nestedSvg = name == "svg" && element.Parent is not null;
        string? attribute = (string?)element.Attribute("opacity");
        string? inlineOpacity = element.Attribute("style") is null
            ? null : ReadSvgStyleDeclarations(element).GetValueOrDefault("opacity");
        double? numeric = ReadUnitSvgOpacity(inlineOpacity ?? attribute);
        // A fully opaque inline style needs no group and admits the same stroke
        // paint as an opaque presentation attribute.
        if (inlineOpacity is not null && numeric != 1d) numeric = null;
        // A numeric zero also hides descendants of a nested SVG viewport.
        return new(nestedSvg ? null : numeric, numeric == 0d,
            !nestedSvg && (numeric.HasValue || (inlineOpacity is null && attribute is null)));
    }

    private readonly record struct SvgContainerOpacity(double? GroupOpacity, bool IsTransparent, bool AllowsGradientStroke);

    private static bool TryCollectSvgOpacityGroups(XElement element, double pictureOpacity,
        List<(XElement Container, double Opacity)> groups, out bool strokeContainersQualified)
    {
        groups.Clear();
        strokeContainersQualified = true;
        int authoredDepth = 0;
        foreach (XElement ancestor in element.Ancestors().Reverse())
        {
            SvgContainerOpacity opacity = ReadSvgContainerOpacity(ancestor);
            if (opacity.IsTransparent) return false;

            if (opacity.GroupOpacity is > 0d and < 1d) authoredDepth++;
            strokeContainersQualified &= opacity.AllowsGradientStroke && authoredDepth <= PdfTransparencyGroup.MaxDepth;

            // Picture alpha shares the root group. Stroke eligibility still follows
            // authored SVG containers, including their distinct nesting limit.
            double groupOpacity = opacity.GroupOpacity ?? 1d;
            if (pictureOpacity < 1d && ancestor.Parent is null) groupOpacity *= pictureOpacity;
            if (groupOpacity is > 0d and < 1d && groups.Count < PdfTransparencyGroup.MaxDepth)
            {
                groups.Add((ancestor, groupOpacity));
            }
        }
        return true;
    }

    private static double? ReadUnitSvgOpacity(string? text)
    {
        double? opacity = ParseSvgOpacity(text);
        return opacity is >= 0d and <= 1d ? opacity : null;
    }

    private static bool IsFullyTransparentSvgContainer(XElement element) =>
        element.Attribute("opacity") is not null && ReadSvgContainerOpacity(element).IsTransparent;

    private static bool IsInFullyTransparentSvgContainer(XElement element)
    {
        foreach (XElement ancestor in element.AncestorsAndSelf())
        {
            if (IsFullyTransparentSvgContainer(ancestor)) { return true; }
        }
        return false;
    }

    private static double? ReadNumericSvgStopOpacity(XElement stop, IReadOnlyDictionary<string, string> style,
        bool hasPaint)
    {
        string? opacity = (string?)stop.Attribute("stop-opacity");
        if (style.TryGetValue("stop-opacity", out string? styledOpacity))
        {
            // A discarded style-bearing stop still prevents alpha admission.
            if (!hasPaint) return null;
            opacity = styledOpacity;
        }
        return opacity is null ? 1d : ReadUnitSvgOpacity(opacity);
    }

    private static bool IsHiddenSvgElement(XElement element, IReadOnlyDictionary<string, string>? style = null)
    {
        style ??= ReadSvgStyleDeclarations(element);
        string? visibility = ReadSvgInheritedAttribute(element, style, "visibility");
        if (visibility?.Trim().ToLowerInvariant() is "hidden" or "collapse")
        {
            return true;
        }

        foreach (XElement source in element.AncestorsAndSelf())
        {
            var declarations = ReferenceEquals(source, element) ? style : ReadSvgStyleDeclarations(source);
            if (string.Equals(ReadSvgPresentationAttribute(source, declarations, "display")?.Trim(), "none", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsSupportedSvgElement(string name)
    {
        return name is "svg" or "defs" or "g" or "title" or "desc" or "metadata" or "style" or "path" or "rect" or "circle" or "ellipse" or "line" or "polyline" or "polygon" or "linearGradient" or "radialGradient" or "stop";
    }
    private static bool IsSvgDefinitionElement(XElement element)
    {
        for (XElement? parent = element.Parent; parent is not null; parent = parent.Parent)
        {
            string name = parent.Name.LocalName;
            if (name is "defs" or "linearGradient" or "radialGradient" or "style" || IsSvgMetadataElement(parent))
            {
                return true;
            }
        }
        return false;
    }
    // Report the first unsupported path command independently of partial path paint.
    private static char? FindFirstUnsupportedSvgPathCommand(string data)
    {
        // Report even when the decoder stops earlier or the path has no paint.
        // Skip scientific-notation exponents that merely look like commands.
        for (int i = 0; i < data.Length; i++)
        {
            char command = data[i];
            if (!char.IsLetter(command) || IsSupportedSvgPathCommand(command) || IsSvgExponentMarker(data, i))
            {
                continue;
            }
            return command;
        }
        return null;
    }
    private static bool IsSvgExponentMarker(string data, int index)
    {
        char marker = data[index];
        if ((marker == (char)101 || marker == (char)69) && index > 0)
        {
            char previous = data[index - 1];
            if ((previous >= (char)48 && previous <= (char)57) || previous == (char)46)
            {
                int next = index + 1;
                if (next < data.Length && (data[next] == (char)43 || data[next] == (char)45))
                {
                    next++;
                }
                if (next < data.Length && data[next] >= (char)48 && data[next] <= (char)57)
                {
                    return true;
                }
            }
        }
        return false;
    }

    private static bool IsSupportedSvgPathCommand(char command)
    {
        return "MLHVCSQTZ".IndexOf(char.ToUpperInvariant(command)) >= 0;
    }
    private static void ReportSkippedSvgPaths(SortedSet<char> unsupportedCommands, int unreadablePaths, SortedSet<string> missingGradients, int unpaintablePaths, int unsupportedTransforms, int gradientStrokes, int unpaintableStrokes, int invalidStrokePresentations, int vectorEffectStrokes, int focalRadialGradients, int nonFiniteGradientPaths, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex, string? partName)
    {
        if (diagnosticSink is null)
        {
            return;
        }
        foreach (char command in unsupportedCommands)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture omits paths using unsupported " + command + " command.");
        }
        if (unreadablePaths > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture omits " + unreadablePaths.ToString(CultureInfo.InvariantCulture) + " unreadable paths.");
        }
        foreach (string gradientId in missingGradients)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture omits paths using unresolvable gradient " + gradientId + ".");
        }
        if (unpaintablePaths > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture omits " + unpaintablePaths.ToString(CultureInfo.InvariantCulture) + " paths with unparsable paint.");
        }
        if (unsupportedTransforms > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture omits " + unsupportedTransforms.ToString(CultureInfo.InvariantCulture) + " paths with unsupported transforms.");
        }
        if (gradientStrokes > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture omits " + gradientStrokes.ToString(CultureInfo.InvariantCulture) + " gradient path strokes.");
        }
        if (unpaintableStrokes > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture omits " + unpaintableStrokes.ToString(CultureInfo.InvariantCulture) + " paths with unparsable stroke paint.");
        }
        if (invalidStrokePresentations > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture renders " + invalidStrokePresentations.ToString(CultureInfo.InvariantCulture) + " paths with default stroke effects for unparsable dash/cap/join/miter values.");
        }
        if (focalRadialGradients > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture ignores focal points on " + focalRadialGradients.ToString(CultureInfo.InvariantCulture) + " radial gradients.");
        }
        if (nonFiniteGradientPaths > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture ignores " + nonFiniteGradientPaths.ToString(CultureInfo.InvariantCulture) + " fills with non-finite gradient geometry.");
        }
        if (vectorEffectStrokes > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture ignores vector-effect on " + vectorEffectStrokes.ToString(CultureInfo.InvariantCulture) + " stroked paths.");
        }
    }
    private static void EmitSvgWarning(Action<OoxPdfDiagnostic> diagnosticSink, int slideIndex, string? partName, string message)
    {
        diagnosticSink(new OoxPdfDiagnostic(
            "SVG_UNSUPPORTED_CONTENT",
            OoxPdfSeverity.Warning,
            message,
            partName,
            PageIndex: null,
            SlideIndex: slideIndex,
            Feature: "svg",
            Fallback: "Partial"));
    }
    private enum SvgFillFailure
    {
        None,
        UnresolvedGradient,
        UnparsableColor
    }

    private static bool TryRenderSvgGradientPath(
        PdfGraphicsBuilder graphics,
        IReadOnlyList<SvgPathCommand> commands,
        SvgGradient gradient,
        SvgTransform transform,
        bool evenOddFill,
        double opacity,
        SvgPathBounds pathBounds,
        double minX,
        double minY,
        double imageX,
        double imageY,
        double imageHeight,
        double scaleX,
        double scaleY,
        bool allowAlphaMask,
        out bool varyingAlphaAdmitted, Action<int> admitSamples, CancellationToken cancellationToken)
    {
        varyingAlphaAdmitted = false;
        // RV07: objectBoundingBox vectors normalize into path space while
        // userSpaceOnUse vectors stay in user units; strips run along the
        // dominant gradient axis so vertical gradients vary top to bottom.
        double pathWidth = Math.Max(0.001d, pathBounds.MaxX - pathBounds.MinX);
        double pathHeight = Math.Max(0.001d, pathBounds.MaxY - pathBounds.MinY);
        // User-space endpoints live in the outer coordinate system, so the path
        // transform applies to them exactly; bounding-box fractions normalize
        // against the already transformed path bounds.
        double vectorMinX;
        double vectorMinY;
        double vectorMaxX;
        double vectorMaxY;
        if (gradient.IsUserSpace)
        {
            (vectorMinX, vectorMinY) = transform.Apply(gradient.X1, gradient.Y1);
            (vectorMaxX, vectorMaxY) = transform.Apply(gradient.X2, gradient.Y2);
        }
        else
        {
            vectorMinX = pathBounds.MinX + gradient.X1 * pathWidth;
            vectorMinY = pathBounds.MinY + gradient.Y1 * pathHeight;
            vectorMaxX = pathBounds.MinX + gradient.X2 * pathWidth;
            vectorMaxY = pathBounds.MinY + gradient.Y2 * pathHeight;
        }

        SvgGradient effective = new SvgGradient(vectorMinX, vectorMinY, vectorMaxX, vectorMaxY, gradient.Stops, gradient.IsUserSpace, gradient.Spread);
        double dx = vectorMaxX - vectorMinX;
        double dy = vectorMaxY - vectorMinY;
        double lengthSquared = dx * dx + dy * dy;
        double maxProjection = Math.Max(Math.Abs(pathBounds.MinX - vectorMinX), Math.Abs(pathBounds.MaxX - vectorMinX)) * Math.Abs(dx)
            + Math.Max(Math.Abs(pathBounds.MinY - vectorMinY), Math.Abs(pathBounds.MaxY - vectorMinY)) * Math.Abs(dy);
        // Finite inputs can overflow when spans, transforms and sampling vectors
        // compose. Reject before changing the graphics state or emitting a clip.
        if (!AreSvgGradientBoundsMappable(pathBounds, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY)
            || !double.IsFinite(vectorMinX) || !double.IsFinite(vectorMinY)
            || !double.IsFinite(vectorMaxX) || !double.IsFinite(vectorMaxY)
            || !double.IsFinite(lengthSquared)
            || (lengthSquared > PptxTextMetricRules.TextStateTolerance && (!double.IsFinite(maxProjection) || !double.IsFinite(maxProjection / lengthSquared))))
        {
            return false;
        }
        graphics.SaveState();
        if (!TryAppendSvgPath(graphics, commands, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY))
        {
            graphics.RestoreState();
            return false;
        }
        graphics.ClipCurrentPath(evenOddFill);
        if (opacity < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(opacity, 1d);
        }
        bool varyingAlpha = gradient.UniformStopOpacity is null;
        bool tryNative = varyingAlpha ? allowAlphaMask : gradient.HasHardStops;
        SvgTransform? axial = tryNative ? ReadSvgAxialShadingTransform(gradient, transform, effective,
            pathBounds, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY) : null;
        if (axial is { } matrix)
        {
            if (varyingAlpha)
            {
                var mask = new PdfGraphicsBuilder();
                mask.SaveState();
                mask.Transform(matrix.M11, matrix.M12, matrix.M21, matrix.M22, matrix.OffsetX, matrix.OffsetY);
                mask.PaintShading(CreateSvgAxialShading(gradient.Stops, opacityMask: true));
                mask.RestoreState();
                double maskX = imageX + (pathBounds.MinX - minX) * scaleX;
                double maskY = imageY + imageHeight - (pathBounds.MaxY - minY) * scaleY;
                graphics.SetVectorLuminositySoftMask(mask, new PdfRectangle(maskX, maskY,
                    Math.Max(0.001d, imageX + (pathBounds.MaxX - minX) * scaleX - maskX),
                    Math.Max(0.001d, imageY + imageHeight - (pathBounds.MinY - minY) * scaleY - maskY)), opacity, 1d);
                varyingAlphaAdmitted = true;
            }
            graphics.SaveState();
            graphics.Transform(matrix.M11, matrix.M12, matrix.M21, matrix.M22, matrix.OffsetX, matrix.OffsetY);
            graphics.PaintShading(CreateSvgAxialShading(gradient.Stops));
            graphics.RestoreState();
            if (opacity < 1d) { graphics.RestoreState(); }
            graphics.RestoreState();
            return true;
        }
        if (Math.Abs(effective.X2 - effective.X1) >= Math.Abs(effective.Y2 - effective.Y1))
        {
            int stripCount = Math.Clamp((int)Math.Ceiling(pathWidth / 2d), 16, 128);
            admitSamples(stripCount);
            double stripSvgWidth = pathWidth / stripCount;
            double sampleY = pathBounds.CenterY;
            for (int strip = 0; strip < stripCount; strip++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                double stripMinX = pathBounds.MinX + strip * stripSvgWidth;
                double stripMaxX = strip == stripCount - 1 ? pathBounds.MaxX : stripMinX + stripSvgWidth;
                double sampleX = (stripMinX + stripMaxX) / 2d;
                RgbColor color = SampleSvgGradient(effective, sampleX, sampleY);
                graphics.SetFillRgb(color.Red, color.Green, color.Blue);
                graphics.FillRectangle(
                    imageX + (stripMinX - minX) * scaleX,
                    imageY,
                    Math.Max(0.001d, (stripMaxX - stripMinX) * scaleX),
                    imageHeight);
            }
        }
        else
        {
            int stripCount = Math.Clamp((int)Math.Ceiling(pathHeight / 2d), 16, 128);
            admitSamples(stripCount);
            double stripSvgHeight = pathHeight / stripCount;
            double sampleX = pathBounds.CenterX;
            double stripX = imageX + (pathBounds.MinX - minX) * scaleX;
            double stripWidth = Math.Max(0.001d, pathWidth * scaleX);
            for (int strip = 0; strip < stripCount; strip++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                double stripMinY = pathBounds.MinY + strip * stripSvgHeight;
                double stripMaxY = strip == stripCount - 1 ? pathBounds.MaxY : stripMinY + stripSvgHeight;
                double sampleY = (stripMinY + stripMaxY) / 2d;
                RgbColor color = SampleSvgGradient(effective, sampleX, sampleY);
                graphics.SetFillRgb(color.Red, color.Green, color.Blue);
                graphics.FillRectangle(
                    stripX,
                    imageY + imageHeight - (stripMaxY - minY) * scaleY,
                    stripWidth,
                    Math.Max(0.001d, (stripMaxY - stripMinY) * scaleY));
            }
        }

        if (opacity < 1d)
        {
            graphics.RestoreState();
        }

        graphics.RestoreState();
        return true;
    }

    private static SvgTransform? ReadSvgAxialShadingTransform(SvgGradient gradient, SvgTransform pathTransform,
        SvgGradient effective, SvgPathBounds bounds, double minX, double minY, double imageX, double imageY,
        double imageHeight, double scaleX, double scaleY, SvgTransform? patternPlacement = null)
    {
        // Mapped endpoints still determine the perpendicular gradient axis
        // after a translation or positive uniform scale.
        if (!gradient.HasNumericStopOpacity ||
            !gradient.HasPositiveUniformGradientTransform || gradient.Spread != SvgGradientSpread.Pad ||
            (gradient.IsUserSpace && !pathTransform.HasPositiveUniformScale) || gradient.Stops.Count > 256 ||
            (!gradient.IsUserSpace && (pathTransform.M11 <= 0d || pathTransform.M22 <= 0d || pathTransform.M12 != 0d || pathTransform.M21 != 0d)) ||
            !HaveRepresentableSvgStops(gradient.Stops)) { return null; }
        double unitWidth = gradient.IsUserSpace ? 1d : bounds.MaxX - bounds.MinX;
        double unitHeight = gradient.IsUserSpace ? 1d : bounds.MaxY - bounds.MinY;
        double startX = gradient.IsUserSpace ? effective.X1 : gradient.X1;
        double startY = gradient.IsUserSpace ? effective.Y1 : gradient.Y1;
        double dx = gradient.IsUserSpace ? effective.X2 - effective.X1 : gradient.X2 - gradient.X1;
        double dy = gradient.IsUserSpace ? effective.Y2 - effective.Y1 : gradient.Y2 - gradient.Y1;
        // Office reversals have a distinct preview fallback; keep prior sampling.
        if (unitWidth <= 0d || unitHeight <= 0d || dx < 0d || dy < 0d) { return null; }
        double lengthSquared = dx * dx + dy * dy;
        // Keep the source perpendicular axis through anisotropic viewport mapping.
        double a = dx * unitWidth * scaleX, b = -dy * unitHeight * scaleY;
        double c = -dy * unitWidth * scaleX, d = -dx * unitHeight * scaleY;
        double e = imageX + (effective.X1 - minX) * scaleX;
        double f = imageY + imageHeight - (effective.Y1 - minY) * scaleY;
        var local = new SvgTransform(a, b, c, d, e, f);
        var placed = patternPlacement is { } placement ? SvgTransform.Compose(placement, local) : local;
        if (!placed.IsFinite || !double.IsFinite(lengthSquared) || lengthSquared <= 0d) { return null; }
        var matrix = new SvgTransform(PrintedSvgNumber(placed.M11), PrintedSvgNumber(placed.M12),
            PrintedSvgNumber(placed.M21), PrintedSvgNumber(placed.M22), PrintedSvgNumber(placed.OffsetX), PrintedSvgNumber(placed.OffsetY));
        double determinant = matrix.M11 * matrix.M22 - matrix.M12 * matrix.M21;
        if (!double.IsFinite(determinant) || determinant == 0d) { return null; }
        foreach ((double x, double y) in new[] { (bounds.MinX, bounds.MinY), (bounds.MinX, bounds.MaxY), (bounds.MaxX, bounds.MinY), (bounds.MaxX, bounds.MaxY) })
        {
            double coordinateX = gradient.IsUserSpace ? x : (x - bounds.MinX) / unitWidth;
            double coordinateY = gradient.IsUserSpace ? y : (y - bounds.MinY) / unitHeight;
            double source = ((coordinateX - startX) * dx + (coordinateY - startY) * dy) / lengthSquared;
            double printedX = PrintedSvgNumber(imageX + (x - minX) * scaleX);
            double printedY = PrintedSvgNumber(imageY + imageHeight - (y - minY) * scaleY);
            if (patternPlacement is { } device) { (printedX, printedY) = device.Apply(printedX, printedY); }
            double actual = (matrix.M22 * (printedX - matrix.OffsetX) - matrix.M21 * (printedY - matrix.OffsetY)) / determinant;
            // Projection error is affine: corner bounds cover the clipped path.
            if (!double.IsFinite(source) || !double.IsFinite(actual) || Math.Abs(actual - source) > 0.001d) { return null; }
        }
        return matrix;
    }

    private static double PrintedSvgNumber(double value) =>
        double.Parse(PdfDocumentWriter.FormatNumber(value), CultureInfo.InvariantCulture);

    private static bool AreSvgGradientBoundsMappable(SvgPathBounds bounds, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY)
    {
        return double.IsFinite(bounds.MaxX - bounds.MinX) && double.IsFinite(bounds.MaxY - bounds.MinY)
            && double.IsFinite(bounds.CenterX) && double.IsFinite(bounds.CenterY)
            && double.IsFinite((bounds.MaxX - bounds.MinX) * scaleX) && double.IsFinite((bounds.MaxY - bounds.MinY) * scaleY)
            && double.IsFinite(imageX + (bounds.MinX - minX) * scaleX) && double.IsFinite(imageX + (bounds.MaxX - minX) * scaleX)
            && double.IsFinite(imageY + imageHeight - (bounds.MinY - minY) * scaleY) && double.IsFinite(imageY + imageHeight - (bounds.MaxY - minY) * scaleY);
    }

    private static PdfShadingPattern? TryReadSvgStrokePattern(SvgPathBounds bounds, SvgGradient gradient, SvgTransform transform,
        double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY,
        SvgTransform? patternPlacement)
    {
        if (!AreSvgGradientBoundsMappable(bounds, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY)) { return null; }
        (double x1, double y1) = gradient.IsUserSpace ? transform.Apply(gradient.X1, gradient.Y1)
            : (bounds.MinX + gradient.X1 * (bounds.MaxX - bounds.MinX), bounds.MinY + gradient.Y1 * (bounds.MaxY - bounds.MinY));
        (double x2, double y2) = gradient.IsUserSpace ? transform.Apply(gradient.X2, gradient.Y2)
            : (bounds.MinX + gradient.X2 * (bounds.MaxX - bounds.MinX), bounds.MinY + gradient.Y2 * (bounds.MaxY - bounds.MinY));
        var effective = new SvgGradient(x1, y1, x2, y2, gradient.Stops, gradient.IsUserSpace, gradient.Spread);
        SvgTransform? projection = ReadSvgAxialShadingTransform(gradient, transform, effective, bounds, minX, minY,
            imageX, imageY, imageHeight, scaleX, scaleY, patternPlacement);
        if (projection is not { } matrix) { return null; }
        return new PdfShadingPattern(CreateSvgAxialShading(gradient.Stops),
            new PdfPatternMatrix(matrix.M11, matrix.M12, matrix.M21, matrix.M22, matrix.OffsetX, matrix.OffsetY));
    }

    private static bool HaveRepresentableSvgStops(IReadOnlyList<SvgGradientStop> stops)
    {
        if (stops.Count < 2 || stops[0].Offset is > 0d and < 0.001d ||
            stops[^1].Offset is > 0.999d and < 1d) return false;
        for (int i = 1; i < stops.Count; i++)
        {
            double interval = stops[i].Offset - stops[i - 1].Offset;
            if (interval != 0d && interval < 0.001d) return false;
        }
        return true;
    }

    private static PdfAxialShading CreateSvgAxialShading(IReadOnlyList<SvgGradientStop> stops, bool opacityMask = false)
    {
        // Move endpoint jumps inside the PDF shading domain. Its extensions
        // then use the first and last source colors on the outer sides.
        double start = stops[1].Offset == 0d ? -1d : 0d;
        double end = stops[^2].Offset == 1d ? 2d : 1d;
        return new PdfAxialShading(start, 0d, end, 0d,
            CreateSvgShadingStops(stops, start, end, opacityMask), preserveHardStops: true);
    }

    private static PdfShadingStop[] CreateSvgShadingStops(IReadOnlyList<SvgGradientStop> stops, double start, double end,
        bool opacityMask = false)
    {
        var result = new PdfShadingStop[stops.Count];
        for (int i = 0; i < stops.Count; i++)
        {
            SvgGradientStop stop = stops[i];
            double offset = (stop.Offset - start) / (end - start);
            RgbColor color = stop.Color;
            if (opacityMask)
            {
                byte alpha = (byte)Math.Round(stop.Opacity!.Value * 255d);
                color = new RgbColor(alpha, alpha, alpha);
            }
            result[i] = new PdfShadingStop(offset, color.Red, color.Green, color.Blue);
        }
        return result;
    }

    private static bool TryReadSvgSourceFrame(XElement? root, out double minX, out double minY, out double width, out double height)
    {
        minX = 0d;
        minY = 0d;
        width = 0d;
        height = 0d;
        if (root is null) return false;
        string? viewBox = (string?)root.Attribute("viewBox");
        if (viewBox is null)
        {
            if (root.Name.LocalName != "svg") return false;
            // Without a viewBox, fixed viewport dimensions define the source
            // coordinate system at the origin, with one user unit per pixel.
            var style = ReadSvgStyleDeclarations(root);
            return TryReadSvgNumberOrPixels(ReadSvgPresentationAttribute(root, style, "width"), out width) &&
                TryReadSvgNumberOrPixels(ReadSvgPresentationAttribute(root, style, "height"), out height);
        }

        double[] values = SvgNumberRegex().Matches(viewBox)
            .Select(match => double.Parse(match.Value, CultureInfo.InvariantCulture))
            .ToArray();
        // RV07: out-of-range viewBox numbers parse to infinity and cannot map;
        // they fail like a missing viewBox instead of poisoning the picture.
        if (values.Length < 4 || !values.All(double.IsFinite))
        {
            return false;
        }

        minX = values[0];
        minY = values[1];
        width = values[2];
        height = values[3];
        return true;
    }

    private static (Dictionary<string, SvgGradient> Linear, Dictionary<string, SvgRadialGradient> Radial)
        ReadSvgGradients(XDocument svg, double viewWidth, double viewHeight, CancellationToken cancellationToken)
    {
        var linear = new Dictionary<string, SvgGradient>(StringComparer.Ordinal);
        var radial = new Dictionary<string, SvgRadialGradient>(StringComparer.Ordinal);
        var sources = new List<(string Id, XElement Element)>();
        var byId = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (XElement element in svg.Descendants())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (element.Name.LocalName is not ("linearGradient" or "radialGradient")) continue;
            string? id = (string?)element.Attribute("id");
            if (!string.IsNullOrWhiteSpace(id))
            {
                sources.Add((id, element));
                byId.TryAdd(id, element);
            }
        }

        var resolved = new Dictionary<XElement, SvgGradientDefinition?>();
        var chain = new List<XElement>();
        foreach ((string id, XElement element) in sources)
        {
            cancellationToken.ThrowIfCancellationRequested();
            SvgGradientDefinition definition = ResolveSvgGradientDefinition(element, byId, resolved, chain, cancellationToken);
            if (definition.Ramp is not { Stops.Length: > 0 } ramp || definition.Transform is not { } transform)
            {
                continue;
            }

            bool userSpace = definition.IsUserSpace;
            double? width = userSpace ? viewWidth : null;
            double? height = userSpace ? viewHeight : null;
            SvgGradientSpread spread = definition.Spread;
            if (definition.IsRadial)
            {
                double? diagonal = userSpace ? GetSvgNormalizedDiagonal(viewWidth, viewHeight) : null;
                if (ReadSvgRadialGradient(definition, ramp, transform, width, height, diagonal, spread) is { } gradient)
                {
                    radial[id] = gradient;
                }
            }
            else if (ReadSvgLinearGradient(definition, ramp, transform, width, height, spread) is { } gradient)
            {
                linear[id] = gradient;
            }
        }
        return (linear, radial);
    }

    private static SvgGradientDefinition ResolveSvgGradientDefinition(XElement element, Dictionary<string, XElement> byId,
        Dictionary<XElement, SvgGradientDefinition?> resolved, List<XElement> chain, CancellationToken cancellationToken)
    {
        chain.Clear();
        XElement? current = element;
        while (current is not null && resolved.TryAdd(current, null))
        {
            cancellationToken.ThrowIfCancellationRequested();
            chain.Add(current);
            string? href = ((string?)current.Attribute("href") ??
                (string?)current.Attribute(XName.Get("href", "http://www.w3.org/1999/xlink")))?.Trim();
            current = href is { Length: > 1 } && href[0] == '#' ? byId.GetValueOrDefault(href[1..]) : null;
        }

        int index = chain.Count - 1;
        if (current is not null && !resolved[current].HasValue)
        {
            // Each cyclic reference is invalid. Keep the cycle members' local
            // declarations independently of traversal order.
            XElement member;
            do
            {
                cancellationToken.ThrowIfCancellationRequested();
                member = chain[index--];
                resolved[member] = ReadSvgGradientDefinition(member, SvgGradientDefinition.Empty, cancellationToken);
            }
            while (!ReferenceEquals(member, current));
        }

        SvgGradientDefinition inherited = current is null ? SvgGradientDefinition.Empty : resolved[current].GetValueOrDefault();
        for (; index >= 0; index--)
        {
            cancellationToken.ThrowIfCancellationRequested();
            XElement source = chain[index];
            inherited = ReadSvgGradientDefinition(source, inherited, cancellationToken);
            resolved[source] = inherited;
        }
        return resolved[element].GetValueOrDefault();
    }

    private static SvgGradientDefinition ReadSvgGradientDefinition(XElement gradient, SvgGradientDefinition inherited,
        CancellationToken cancellationToken)
    {
        bool radial = gradient.Name.LocalName == "radialGradient";
        SvgGradientDefinition geometry = radial == inherited.IsRadial ? inherited : default;
        SvgTransform? transform = inherited.Transform;
        if (gradient.Attribute("gradientTransform") is { } attribute)
        {
            transform = TryParseSvgTransformList(attribute.Value, out SvgTransform local) ? local : null;
        }

        // Local stops replace the palette, including when their paint cannot be
        // used. Inherited stops keep the color context of their authored source.
        return new SvgGradientDefinition(
            ReadSvgGradientStops(gradient, cancellationToken) ?? inherited.Ramp,
            radial,
            gradient.Attribute("gradientUnits") is { } units ? units.Value == "userSpaceOnUse" : inherited.IsUserSpace,
            transform,
            gradient.Attribute("spreadMethod") is { } spread ? ReadSvgGradientSpread(spread.Value) : inherited.Spread,
            radial ? default : ReadSvgGradientCoordinate(gradient, "x1", geometry.X1),
            radial ? default : ReadSvgGradientCoordinate(gradient, "y1", geometry.Y1),
            radial ? default : ReadSvgGradientCoordinate(gradient, "x2", geometry.X2),
            radial ? default : ReadSvgGradientCoordinate(gradient, "y2", geometry.Y2),
            radial ? ReadSvgGradientCoordinate(gradient, "cx", geometry.Cx) : default,
            radial ? ReadSvgGradientCoordinate(gradient, "cy", geometry.Cy) : default,
            radial ? ReadSvgGradientCoordinate(gradient, "r", geometry.Radius) : default,
            radial ? ReadSvgGradientCoordinate(gradient, "fx", geometry.Fx) : default,
            radial ? ReadSvgGradientCoordinate(gradient, "fy", geometry.Fy) : default);
    }

    private readonly record struct SvgGradientDefinition(SvgGradientRamp? Ramp, bool IsRadial, bool IsUserSpace,
        SvgTransform? Transform, SvgGradientSpread Spread, SvgGradientCoordinate X1, SvgGradientCoordinate Y1,
        SvgGradientCoordinate X2, SvgGradientCoordinate Y2, SvgGradientCoordinate Cx, SvgGradientCoordinate Cy,
        SvgGradientCoordinate Radius, SvgGradientCoordinate Fx, SvgGradientCoordinate Fy)
    {
        public static SvgGradientDefinition Empty => default(SvgGradientDefinition) with { Transform = SvgTransform.Identity };
    }

    private sealed record SvgGradientRamp(SvgGradientStop[] Stops, double? UniformOpacity, bool HasNumericOpacity,
        bool HasHardStops, int IgnoredOpacityCount);

    private static SvgGradient? ReadSvgLinearGradient(SvgGradientDefinition definition, SvgGradientRamp ramp,
        SvgTransform transform, double? width, double? height, SvgGradientSpread spread)
    {
        if (!definition.X1.TryResolve(0d, width, out double x1)
            || !definition.Y1.TryResolve(0d, height, out double y1)
            || !definition.X2.TryResolve(width ?? 1d, width, out double x2)
            || !definition.Y2.TryResolve(0d, height, out double y2))
        {
            return null;
        }
        (double rawX1, double rawY1) = transform.Apply(x1, y1);
        (double rawX2, double rawY2) = transform.Apply(x2, y2);
        if (!double.IsFinite(rawX1) || !double.IsFinite(rawY1) || !double.IsFinite(rawX2) || !double.IsFinite(rawY2))
        {
            return null;
        }
        return new SvgGradient(rawX1, rawY1, rawX2, rawY2, ramp.Stops, width.HasValue, spread)
        {
            UniformStopOpacity = ramp.UniformOpacity,
            HasNumericStopOpacity = ramp.HasNumericOpacity,
            // The mapped endpoints still determine the gradient normal after
            // a translation or positive uniform scale in gradient space.
            HasPositiveUniformGradientTransform = transform.HasPositiveUniformScale,
            HasHardStops = ramp.HasHardStops,
            IgnoredStopOpacityCount = ramp.IgnoredOpacityCount,
            // A single stop or coincident source endpoints define solid paint.
            ConstantStop = ramp.HasNumericOpacity && (ramp.Stops.Length == 1 || (x1 == x2 && y1 == y2))
                ? ramp.Stops[^1] : null
        };
    }

    private static SvgGradientRamp? ReadSvgGradientStops(XElement gradient, CancellationToken cancellationToken)
    {
        List<SvgGradientStop>? stops = null;
        double? firstNumericOpacity = null;
        bool hasNumericOpacity = true;
        bool hasUniformOpacity = true;
        bool hasHardStops = false;
        int fallbackOpacityCount = 0;
        foreach (XElement stop in gradient.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (stop.Name.LocalName != "stop") continue;
            stops ??= new List<SvgGradientStop>();
            IReadOnlyDictionary<string, string> style = ReadSvgStyleDeclarations(stop);
            bool hasPaint = TryReadSvgStopPaint(stop, style, out double offset, out RgbColor color);
            // Opacity admission considers all authored stops, including those discarded for paint.
            double? opacity = ReadNumericSvgStopOpacity(stop, style, hasPaint);
            hasNumericOpacity &= opacity.HasValue;
            firstNumericOpacity ??= opacity;
            hasUniformOpacity &= opacity == firstNumericOpacity;
            if ((opacity ?? ReadSvgOpacityValue(style.GetValueOrDefault("stop-opacity") ??
                (string?)stop.Attribute("stop-opacity"))) < 1d)
            {
                fallbackOpacityCount++;
            }

            if (!hasPaint) continue;
            if (stops.Count > 0)
            {
                // SVG raises a decreasing offset to the preceding one and
                // keeps both sides of the transition in source order.
                offset = Math.Max(offset, stops[^1].Offset);
                hasHardStops |= offset == stops[^1].Offset;
            }
            stops.Add(new SvgGradientStop(offset, color, opacity));
        }

        if (stops is null) return null;
        double? uniformOpacity = hasNumericOpacity && hasUniformOpacity ? firstNumericOpacity ?? 1d : null;
        return new SvgGradientRamp(stops.ToArray(), uniformOpacity, hasNumericOpacity, hasHardStops,
            uniformOpacity.HasValue ? 0 : fallbackOpacityCount);
    }

    private static bool TryReadSvgStopPaint(XElement stop, IReadOnlyDictionary<string, string> style,
        out double offset, out RgbColor color)
    {
        offset = ReadSvgOffset((string?)stop.Attribute("offset"));
        color = default;
        return !double.IsNaN(offset) && TryReadSvgStopColor(stop, style, out color);
    }

    private static bool TryReadSvgStopColor(XElement stop, IReadOnlyDictionary<string, string> style, out RgbColor color)
    {
        if (style.TryGetValue("stop-color", out string? inlineColor) &&
            TryReadSvgSolidColor(stop, style, inlineColor, out color))
        {
            return true;
        }

        // Keep the presentation color when the inline declaration cannot be used.
        string? attributeColor = (string?)stop.Attribute("stop-color");
        return TryReadSvgSolidColor(stop, style, string.IsNullOrWhiteSpace(attributeColor) ? "black" : attributeColor, out color);
    }

    private static double ReadSvgOffset(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0d;
        }

        string trimmed = value.Trim();
        if (trimmed.EndsWith("%", StringComparison.Ordinal)
            && double.TryParse(trimmed[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out double percent))
        {
            return Math.Clamp(percent / 100d, 0d, 1d);
        }

        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double offset))
        {
            return Math.Clamp(offset, 0d, 1d);
        }
        return double.NaN;
    }

    // RV07: reflect and repeat tile the gradient vector; unknown methods pad.
    private static SvgGradientSpread ReadSvgGradientSpread(string? value)
    {
        return value switch
        {
            "reflect" => SvgGradientSpread.Reflect,
            "repeat" => SvgGradientSpread.Repeat,
            _ => SvgGradientSpread.Pad,
        };
    }

    private enum SvgGradientUnit { Unspecified, Number, Percent, Pixel, Invalid }

    private readonly record struct SvgGradientCoordinate(double Value, SvgGradientUnit Unit)
    {
        // The receiving gradient supplies the coordinate system after templating.
        public bool TryResolve(double fallback, double? viewportBasis, out double value)
        {
            value = Unit switch
            {
                SvgGradientUnit.Unspecified => fallback,
                SvgGradientUnit.Percent => (Value / 100d) * (viewportBasis ?? 1d),
                _ => Value
            };
            return Unit != SvgGradientUnit.Invalid && (Unit != SvgGradientUnit.Pixel || viewportBasis.HasValue) && double.IsFinite(value);
        }
    }

    private static SvgGradientCoordinate ReadSvgGradientCoordinate(XElement gradient, string name, SvgGradientCoordinate inherited)
    {
        if (gradient.Attribute(name) is not { } attribute) return inherited;
        ReadOnlySpan<char> text = attribute.Value.AsSpan().Trim();
        if (text.IsEmpty) return default;

        SvgGradientUnit unit = SvgGradientUnit.Number;
        if (text.EndsWith("%", StringComparison.Ordinal))
        {
            unit = SvgGradientUnit.Percent;
            text = text[..^1];
        }
        else if (text.EndsWith("px", StringComparison.OrdinalIgnoreCase))
        {
            unit = SvgGradientUnit.Pixel;
            text = text[..^2];
        }
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) && double.IsFinite(value)
            ? new SvgGradientCoordinate(value, unit) : new SvgGradientCoordinate(0d, SvgGradientUnit.Invalid);
    }

    // RV07: CSS style declarations override presentation attributes; opacity
    // multiplies fill-opacity with whole-element opacity.
    private static IReadOnlyDictionary<string, string> ReadSvgStyleDeclarations(XElement path)
    {
        if ((string?)path.Attribute("style") is not { Length: > 0 } style)
        {
            return ReadOnlyDictionary<string, string>.Empty;
        }

        var declarations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        HashSet<string>? importantNames = null;
        foreach (string declaration in style.Split((char)59))
        {
            int colon = declaration.IndexOf((char)58);
            if (colon <= 0)
            {
                continue;
            }

            string name = declaration.Substring(0, colon).Trim();
            string value = declaration.Substring(colon + 1).Trim();
            if (name.Length == 0 || value.Length == 0)
            {
                continue;
            }

            int marker = value.LastIndexOf('!');
            if (marker >= 0 && value.AsSpan(marker + 1).Trim().Equals("important", StringComparison.OrdinalIgnoreCase))
            {
                value = value[..marker].TrimEnd();
                if (value.Length == 0)
                {
                    continue;
                }

                (importantNames ??= new(StringComparer.OrdinalIgnoreCase)).Add(name);
            }
            else if (importantNames?.Contains(name) == true)
            {
                continue;
            }

            declarations[name] = value;
        }

        return declarations;
    }

    private static double ReadSvgOpacityValue(string? text) =>
        Math.Clamp(ParseSvgOpacity(text) ?? 1d, 0d, 1d);

    private static double? ParseSvgOpacity(string? text)
    {
        ReadOnlySpan<char> token = text.AsSpan().Trim();
        bool percent = token.Length != 0 && token[^1] == '%';
        if (percent) token = token[..^1];
        if (!double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double opacity)) return null;
        return percent ? opacity / 100d : opacity;
    }

    private static bool TryReadSvgFill(XElement path, IReadOnlyDictionary<string, string> style,
        IReadOnlyDictionary<string, SvgGradient> gradients,
        IReadOnlyDictionary<string, SvgRadialGradient> radialGradients, Lazy<HashSet<string>> elementIds,
        double elementOpacity, out SvgPaint paint, out SvgFillFailure failure, out string? gradientId, out SvgRadialGradient? radial)
    {
        gradientId = null;
        radial = null;
        // A line has no interior, so its fill cannot affect stroke composition.
        if (path.Name.LocalName == "line")
        {
            paint = default;
            failure = SvgFillFailure.None;
            return false;
        }

        string? fill = ReadSvgInheritedAttribute(path, style, "fill");
        double opacity = ReadSvgOpacityValue(ReadSvgInheritedAttribute(path, style, "fill-opacity")) * elementOpacity;
        if (fill is null)
        {
            // SVG initial value: a path without fill paints black instead of vanishing.
            paint = new SvgPaint(new RgbColor(0, 0, 0), null, opacity);
            failure = SvgFillFailure.None;
            return true;
        }
        fill = ResolveSvgPaintFallback(path, style, fill, elementIds, out gradientId);
        if (fill.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            paint = default;
            failure = SvgFillFailure.None;
            return false;
        }
        if (gradientId is not null)
        {
            SvgGradient? svgGradient = gradients.GetValueOrDefault(gradientId);
            SvgRadialGradient? svgRadial = svgGradient is null ? radialGradients.GetValueOrDefault(gradientId) : null;
            if ((svgGradient?.ConstantStop ?? svgRadial?.ConstantStop) is { Opacity: double stopOpacity } stop)
            {
                paint = new SvgPaint(stop.Color, null, opacity * stopOpacity);
                gradientId = null;
                failure = SvgFillFailure.None;
                return true;
            }
            if (svgGradient is not null)
            {
                paint = new SvgPaint(null, svgGradient, opacity * (svgGradient.UniformStopOpacity ?? 1d));
                failure = SvgFillFailure.None;
                return true;
            }
            if (svgRadial is not null)
            {
                radial = svgRadial;
                paint = new SvgPaint(null, null, opacity * (svgRadial.UniformStopOpacity ?? 1d));
                failure = SvgFillFailure.None;
                return true;
            }
            paint = default;
            failure = SvgFillFailure.UnresolvedGradient;
            return false;
        }
        if (TryReadSvgSolidColor(path, style, fill, out RgbColor color))
        {
            paint = new SvgPaint(color, null, opacity);
            failure = SvgFillFailure.None;
            return true;
        }
        paint = default;
        failure = SvgFillFailure.UnparsableColor;
        return false;
    }

    private readonly record struct SvgStroke(RgbColor? Color, double Width, double Opacity, int LineCap, int LineJoin, double[]? DashPattern, double DashOffset, double MiterLimit, SvgGradient? Gradient = null, string? GradientId = null)
    {
        public bool HasPaint => Color is not null || Gradient is not null;
    }
    private enum SvgStrokeFailure
    {
        None,
        UnresolvedGradient,
        UnparsableColor,
        UnparsableWidth
    }
    // RV07: solid strokes paint over any fill; gradient strokes, unparsable
    // colors/widths and dash/cap/join effects diagnose instead of vanishing.
    // Office uses the largest singular value of the element transform for
    // a uniform stroke width; the viewport separately stretches its normal.
    // vector-effect=non-scaling-stroke skips the element-transform factor;
    // any other effect value stays diagnosed with the scaled stroke.
    private static SvgStroke ReadSvgStroke(XElement path, IReadOnlyDictionary<string, string> style,
        IReadOnlyDictionary<string, SvgGradient> gradients,
        IReadOnlyDictionary<string, SvgRadialGradient> radialGradients, Lazy<HashSet<string>> elementIds,
        bool allowGradientStroke, double elementOpacity, double percentageBasis, out SvgStrokeFailure failure,
        out bool invalidPresentation, out bool hasVectorEffect, out bool nonScalingStroke)
    {
        failure = SvgStrokeFailure.None;
        string? strokePaint = ReadSvgInheritedAttribute(path, style, "stroke");
        invalidPresentation = false;
        string? vectorEffect = ReadSvgPresentationAttribute(path, style, "vector-effect");
        nonScalingStroke = vectorEffect?.Trim().Equals("non-scaling-stroke", StringComparison.OrdinalIgnoreCase) == true;
        hasVectorEffect = !nonScalingStroke && !string.IsNullOrWhiteSpace(vectorEffect);
        if (string.IsNullOrWhiteSpace(strokePaint) || strokePaint.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return default;
        }
        string? strokeOpacity = ReadSvgInheritedAttribute(path, style, "stroke-opacity");
        double opacity = ReadSvgOpacityValue(strokeOpacity) * elementOpacity;
        if (!TryReadSvgStrokeWidth(style, path, percentageBasis, out double width))
        {
            failure = SvgLocalPaintServerRegex().IsMatch(strokePaint)
                ? SvgStrokeFailure.UnresolvedGradient : SvgStrokeFailure.UnparsableWidth;
            return default;
        }
        // Zero means no stroke, before paint-server lookup or any positive PDF
        // width clamp. A fill on the same source remains independently paintable.
        if (width == 0d) { return default; }
        strokePaint = ResolveSvgPaintFallback(path, style, strokePaint, elementIds, out string? gradientId);
        if (strokePaint.Equals("none", StringComparison.OrdinalIgnoreCase)) return default;
        SvgGradient? strokeGradient = null;
        RgbColor? color = null;
        if (gradientId is not null)
        {
            SvgGradient? gradient = gradients.GetValueOrDefault(gradientId);
            SvgGradientStop? constant = gradient is null
                ? radialGradients.GetValueOrDefault(gradientId)?.ConstantStop : gradient.ConstantStop;
            if (constant is { Opacity: double stopOpacity } stop)
            {
                color = stop.Color;
                opacity *= stopOpacity;
                gradientId = null;
            }
            else
            {
                if (!allowGradientStroke || (!string.IsNullOrWhiteSpace(strokeOpacity) && ReadUnitSvgOpacity(strokeOpacity) is null)
                    || gradient is null || gradient.UniformStopOpacity is null || !gradient.HasNumericStopOpacity
                    || !gradient.HasPositiveUniformGradientTransform || gradient.Spread != SvgGradientSpread.Pad)
                {
                    failure = SvgStrokeFailure.UnresolvedGradient;
                    return default;
                }
                strokeGradient = gradient;
                opacity *= gradient.UniformStopOpacity.Value;
            }
        }
        else if (TryReadSvgSolidColor(path, style, strokePaint, out RgbColor parsedColor))
        {
            color = parsedColor;
        }
        else
        {
            failure = SvgStrokeFailure.UnparsableColor;
            return default;
        }
        int lineCap = ReadSvgLineCap(ReadSvgInheritedAttribute(path, style, "stroke-linecap"), out bool capInvalid);
        int lineJoin = ReadSvgLineJoin(ReadSvgInheritedAttribute(path, style, "stroke-linejoin"), out bool joinInvalid);
        double[]? dash = ReadSvgDashPattern(path, ReadSvgInheritedAttribute(path, style, "stroke-dasharray"),
            percentageBasis, out bool dashInvalid);
        double offset = 0d;
        if (dash is not null && ReadSvgInheritedAttribute(path, style, "stroke-dashoffset") is { } dashOffset &&
            !TryReadSvgViewportLength(path, dashOffset, percentageBasis, out offset))
        {
            dash = null;
            dashInvalid = true;
        }
        invalidPresentation = capInvalid || joinInvalid || dashInvalid;
        double miterLimit = 4d;
        if (!TryReadSvgMiterLimit(ReadSvgInheritedAttribute(path, style, "stroke-miterlimit"), out miterLimit))
        {
            invalidPresentation = true;
            miterLimit = 4d;
        }
        return new SvgStroke(color, width, opacity, lineCap, lineJoin, dash, offset, miterLimit, strokeGradient, gradientId);
    }

    private static Lazy<HashSet<string>> CreateSvgElementIdLookup(XDocument svg, CancellationToken cancellationToken) =>
        new(() =>
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (XElement element in svg.Descendants())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if ((string?)element.Attribute("id") is { } id) ids.Add(id);
            }
            return ids;
        });

    [GeneratedRegex("""^\s*url\(\s*(?:"#(?<id>[^"\s\\]+)"|'#(?<id>[^'\s\\]+)'|#(?<id>[^()\s"'\\]+))\s*\)""",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex SvgLocalPaintServerRegex();

    private static string ResolveSvgPaintFallback(XElement element, IReadOnlyDictionary<string, string> style,
        string paint, Lazy<HashSet<string>> elementIds, out string? gradientId)
    {
        gradientId = null;
        if (!paint.AsSpan().TrimStart().StartsWith("url(", StringComparison.OrdinalIgnoreCase)) return paint;
        Match reference = SvgLocalPaintServerRegex().Match(paint);
        gradientId = reference.Success ? reference.Groups["id"].Value : null;
        if (gradientId is null) return paint;

        ReadOnlySpan<char> suffix = paint.AsSpan(reference.Length);
        if (suffix.IsEmpty || !char.IsWhiteSpace(suffix[0]) || gradientId.Contains('%')) return paint;

        // A present server still owns its existing paint and diagnostics, even
        // when the renderer cannot use it. Resolve only absent local targets.
        string fallback = suffix.Trim().ToString();
        if ((!fallback.Equals("none", StringComparison.OrdinalIgnoreCase) &&
             !TryReadSvgSolidColor(element, style, fallback, out _)) || elementIds.Value.Contains(gradientId)) return paint;
        gradientId = null;
        return fallback;
    }
    private static bool TryReadSvgStrokeWidth(IReadOnlyDictionary<string, string> style, XElement path,
        double percentageBasis, out double width)
    {
        width = 1d;
        string? text = ReadSvgInheritedAttribute(path, style, "stroke-width");
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }
        return TryReadSvgViewportLength(path, text, percentageBasis, out width) && width >= 0d;
    }
    // Resolve the winning declaration on each element before following inheritance.
    // Whole-element opacity stays with the element or its compositing group.
    private static string? ReadSvgInheritedAttribute(XElement path, IReadOnlyDictionary<string, string> style, string name)
    {
        foreach (XElement source in path.AncestorsAndSelf())
        {
            IReadOnlyDictionary<string, string> declarations = ReferenceEquals(source, path) ? style : ReadSvgStyleDeclarations(source);
            string? value = ReadSvgPresentationAttribute(source, declarations, name);
            if (!string.IsNullOrWhiteSpace(value) && !IsSvgInherit(value) &&
                (name != "color" || !IsSvgCurrentColor(value)))
            {
                return value;
            }
        }
        return null;
    }

    private static bool IsSvgInherit(string value) =>
        value.AsSpan().Trim().Equals("inherit", StringComparison.OrdinalIgnoreCase);

    private static bool IsSvgCurrentColor(string value) =>
        value.AsSpan().Trim().Equals("currentColor", StringComparison.OrdinalIgnoreCase);

    private static bool TryReadSvgSolidColor(XElement element, IReadOnlyDictionary<string, string> style,
        string value, out RgbColor color)
    {
        // A descendant can inherit the paint keyword and supply its own color.
        if (IsSvgCurrentColor(value))
            value = ReadSvgInheritedAttribute(element, style, "color") ?? "black";
        return RgbColor.TryParseCssColor(value, out color);
    }

    private static string? ReadSvgPresentationAttribute(XElement path, IReadOnlyDictionary<string, string> style, string name)
    {
        if (style.TryGetValue(name, out string? styleValue) && !string.IsNullOrWhiteSpace(styleValue))
        {
            return styleValue;
        }
        return (string?)path.Attribute(name);
    }
    private static int ReadSvgLineCap(string? value, out bool invalid)
    {
        invalid = false;
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0;
        }
        string trimmed = value.Trim();
        if (trimmed.Equals("butt", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        if (trimmed.Equals("round", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        if (trimmed.Equals("square", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        invalid = true;
        return 0;
    }
    private static int ReadSvgLineJoin(string? value, out bool invalid)
    {
        invalid = false;
        if (string.IsNullOrWhiteSpace(value))
        {
            return 0;
        }
        string trimmed = value.Trim();
        if (trimmed.Equals("miter", StringComparison.OrdinalIgnoreCase))
        {
            return 0;
        }
        if (trimmed.Equals("round", StringComparison.OrdinalIgnoreCase))
        {
            return 1;
        }
        if (trimmed.Equals("bevel", StringComparison.OrdinalIgnoreCase))
        {
            return 2;
        }
        invalid = true;
        return 0;
    }
    // Dash lengths and widths share the viewport diagonal. An odd list doubles
    // per SVG and an all-zero pattern means solid.
    private static double[]? ReadSvgDashPattern(XElement path, string? text, double percentageBasis, out bool invalid)
    {
        invalid = false;
        if (string.IsNullOrWhiteSpace(text) || text.Trim().Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }
        string[] parts = text.Split([',', ' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
        var lengths = new List<double>();
        foreach (string part in parts)
        {
            if (!TryReadSvgViewportLength(path, part, percentageBasis, out double length) || length < 0d)
            {
                invalid = true;
                return null;
            }
            lengths.Add(length);
        }
        if (lengths.Count == 0)
        {
            invalid = true;
            return null;
        }
        if (lengths.TrueForAll(length => length == 0d))
        {
            return null;
        }
        if (lengths.Count % 2 == 1)
        {
            lengths.AddRange(lengths.ToArray());
        }
        return lengths.ToArray();
    }
    // RV07: miter limits below 1 are invalid per SVG; the limit only shapes
    // miter joins, so other joins never emit it.
    private static bool TryReadSvgMiterLimit(string? text, out double miterLimit)
    {
        miterLimit = 4d;
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }
        return TryReadSvgNumberOrPixels(text, out miterLimit) && miterLimit >= 1d;
    }
    private static bool TryReadSvgNumberOrPixels(ReadOnlySpan<char> text, out double value)
    {
        text = text.Trim();
        if (text.EndsWith("px", StringComparison.OrdinalIgnoreCase))
        {
            text = text[..^2];
        }
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value) && double.IsFinite(value);
    }

    private static double GetSvgNormalizedDiagonal(double width, double height) =>
        double.Hypot(width / Math.Sqrt(2d), height / Math.Sqrt(2d));

    private static double GetSvgStrokeElementScale(SvgTransform transform)
    {
        // RV07-S2: PowerPoint scales a uniform element stroke by the largest
        // singular value, including shear. Preserve the previous degenerate
        // and overflow fallback before normalizing the finite matrix entries.
        double area = transform.M11 * transform.M22 - transform.M12 * transform.M21;
        if (!double.IsFinite(area) || area == 0d)
        {
            return Math.Sqrt(Math.Abs(area));
        }
        double largest = Math.Max(Math.Max(Math.Abs(transform.M11), Math.Abs(transform.M12)),
            Math.Max(Math.Abs(transform.M21), Math.Abs(transform.M22)));
        double a = transform.M11 / largest, b = transform.M12 / largest;
        double c = transform.M21 / largest, d = transform.M22 / largest;
        double first = Math.Sqrt((a + d) * (a + d) + (b - c) * (b - c));
        double second = Math.Sqrt((a - d) * (a - d) + (b + c) * (b + c));
        return largest * ((first + second) / 2d);
    }
    private static bool TryPaintSvgStrokePath(PdfGraphicsBuilder graphics, IReadOnlyList<SvgPathCommand> commands, bool completePath, RgbColor color, double widthPoints, double opacity, int lineCap, int lineJoin, double[]? dashPoints, double dashPhasePoints, double miterLimit, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY, PdfShadingPattern? shadingPattern = null)
    {
        graphics.SaveState();
        SvgTransform? strokeViewport = ApplySvgStrokeViewport(graphics, imageX, imageY, imageHeight, scaleX, scaleY);
        SvgRoundDashPaint? roundDashPaint = completePath && opacity == 1d && lineCap == 1 && dashPoints is { Length: 2 }
            ? ReadSvgRoundDashPaint(commands, widthPoints, dashPoints, dashPhasePoints, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY, strokeViewport)
            : null;
        List<SvgRoundDashCap>? roundDashCaps = roundDashPaint?.Caps;
        if (opacity < 1d)
        {
            graphics.SetAlpha(1d, opacity);
        }
        if (shadingPattern is null) { graphics.SetStrokeRgb(color.Red, color.Green, color.Blue); }
        else { graphics.SetStrokeShadingPattern(shadingPattern); }
        graphics.SetLineWidth(widthPoints);
        if (roundDashCaps is not null)
        {
            graphics.SetLineCap(0);
        }
        else if (lineCap != 0)
        {
            graphics.SetLineCap(lineCap);
        }
        if (lineJoin != 0)
        {
            graphics.SetLineJoin(lineJoin);
        }
        if (dashPoints is not null)
        {
            graphics.SetLineDash(dashPoints, dashPhasePoints);
        }
        if (lineJoin == 0)
        {
            graphics.SetMiterLimit(miterLimit);
        }
        bool painted = TryAppendSvgPath(graphics, commands, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY, strokeViewport);
        if (painted)
        {
            graphics.StrokeCurrentPath();
            if (roundDashCaps is not null)
            {
                if (shadingPattern is null) { graphics.SetFillRgb(color.Red, color.Green, color.Blue); }
                else { graphics.SetFillShadingPattern(shadingPattern); }
                foreach (SvgRoundDashCap cap in roundDashCaps)
                {
                    PaintSvgRoundDashCap(graphics, cap, widthPoints / 2d);
                }
            }
            if (roundDashPaint is { SolidLines: { Count: > 0 } } dashPaint)
            {
                // RV07-S7: Office paints an opaque rounded line when its
                // entire path falls in dash gaps. Retain mixed painted paths.
                graphics.SetLineDash(Array.Empty<double>(), 0d);
                graphics.SetLineCap(1);
                foreach ((double x, double y, double endX, double endY) in dashPaint.SolidLines)
                {
                    graphics.MoveTo(x, y);
                    graphics.LineTo(endX, endY);
                }
                graphics.StrokeCurrentPath();
            }
        }
        graphics.RestoreState();
        return painted;
    }
    private readonly record struct SvgRoundDashCap(double X, double Y, double DirectionX, double DirectionY);
    private readonly record struct SvgRoundDashPaint(List<SvgRoundDashCap>? Caps, List<(double X, double Y, double EndX, double EndY)> SolidLines);
    private static void PaintSvgRoundDashCap(PdfGraphicsBuilder graphics, SvgRoundDashCap cap, double radius)
    {
        double nx = -cap.DirectionY, ny = cap.DirectionX;
        double sx = cap.X + nx * radius, sy = cap.Y + ny * radius;
        double tx = cap.X + cap.DirectionX * radius, ty = cap.Y + cap.DirectionY * radius;
        double ex = cap.X - nx * radius, ey = cap.Y - ny * radius;
        double kappa = SvgCircleKappa * radius;
        graphics.MoveTo(sx, sy);
        graphics.CurveTo(sx + cap.DirectionX * kappa, sy + cap.DirectionY * kappa,
            tx + nx * kappa, ty + ny * kappa, tx, ty);
        graphics.CurveTo(tx - nx * kappa, ty - ny * kappa,
            ex + cap.DirectionX * kappa, ey + cap.DirectionY * kappa, ex, ey);
        graphics.ClosePath();
        graphics.FillCurrentPath();
    }
    private static SvgRoundDashPaint? ReadSvgRoundDashPaint(IReadOnlyList<SvgPathCommand> commands, double widthPoints, double[] dash, double phase, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY, SvgTransform? viewport)
    {
        // RV07-S6: Office rounds visible dash ends within half a stroke width
        // of an original endpoint. Outward semicircles retain partial dash gaps.
        // Qualify complete single-line subpaths with flat dash ends;
        // RV07-S7: a wholly unpainted straight path gets solid opaque paint.
        // Curves, joins, alpha and extreme geometry retain native fallback.
        double period = dash[0] + dash[1];
        if (dash[0] <= 0d || dash[1] <= 0d || !double.IsFinite(period)) return null;
        bool qualifiedCaps = dash[0] > widthPoints;
        double meanScale = (scaleX + scaleY) / 2d;
        if (viewport is null && Math.Abs(scaleX - scaleY) > meanScale * .000001d) return null;
        if (commands.Count == 0 || commands.Count > 256 || commands.Count % 2 != 0) return null;
        var caps = new List<SvgRoundDashCap>();
        var solidLines = new List<(double X, double Y, double EndX, double EndY)>();
        bool hasPaintedDash = false;
        phase %= period;
        if (phase < 0d) phase += period;
        for (int index = 0; index < commands.Count; index += 2)
        {
            SvgPathCommand move = commands[index], line = commands[index + 1];
            if (move.Kind != 'M' || line.Kind is not ('L' or 'H' or 'V')) return null;
            double endX = line.Kind == 'V' ? move.Arguments[0] : line.Arguments[0];
            double endY = line.Kind == 'H' ? move.Arguments[1] : line.Kind == 'V' ? line.Arguments[0] : line.Arguments[1];
            double sourceDx = endX - move.Arguments[0], sourceDy = endY - move.Arguments[1];
            double largest = Math.Max(Math.Abs(sourceDx), Math.Abs(sourceDy));
            if (!double.IsFinite(largest) || largest == 0d) return null;
            // Use source geometry for cycle boundaries, before the rounded
            // viewport CTM and translated path coordinates add numeric drift.
            double length = largest * Math.Sqrt(Math.Pow(sourceDx / largest, 2d) + Math.Pow(sourceDy / largest, 2d)) * meanScale;
            if (!TryMapSvgPicturePoint(move.Arguments[0], move.Arguments[1], minX, minY, imageX, imageY, imageHeight, scaleX, scaleY, viewport, out double x, out double y)
                || !TryMapSvgPicturePoint(endX, endY, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY, viewport, out endX, out endY)) return null;
            double dx = endX - x, dy = endY - y;
            if (!double.IsFinite(dx) || !double.IsFinite(dy) || !double.IsFinite(length)
                || length <= 0d || length / period > 1e12d || !double.IsFinite(length + phase)) return null;
            double mappedLargest = Math.Max(Math.Abs(dx), Math.Abs(dy));
            double mappedLength = mappedLargest * Math.Sqrt(Math.Pow(dx / mappedLargest, 2d) + Math.Pow(dy / mappedLargest, 2d));
            if (!double.IsFinite(mappedLength) || mappedLength == 0d) return null;
            double tolerance = Math.Max(1d, period) * 1e-12d;
            double first = phase < dash[0] ? 0d : period - phase;
            if (first >= length - tolerance)
            {
                double radius = widthPoints / 2d;
                if (!double.IsFinite(x - radius) || !double.IsFinite(x + radius)
                    || !double.IsFinite(y - radius) || !double.IsFinite(y + radius)
                    || !double.IsFinite(endX - radius) || !double.IsFinite(endX + radius)
                    || !double.IsFinite(endY - radius) || !double.IsFinite(endY + radius)) return null;
                solidLines.Add((x, y, endX, endY));
                qualifiedCaps = false;
                continue;
            }
            hasPaintedDash = true;
            double remainder = (length + phase) % period;
            if (remainder < tolerance || period - remainder < tolerance) remainder = 0d;
            double last = remainder == 0d ? length - dash[1] : remainder <= dash[0] ? length : length - (remainder - dash[0]);
            if (last <= first || last > length) return null;
            double firstEnd = phase < dash[0] ? dash[0] - phase : first + dash[0];
            bool firstCap = first <= widthPoints / 2d + tolerance;
            bool lastCap = length - last <= widthPoints / 2d + tolerance;
            // RV07-S8: one visible dash needs flat ends outside Office endpoint
            // regions. Keep native paint when both rounded ends already match.
            if (length <= widthPoints || (last <= firstEnd + tolerance && firstCap && lastCap))
            {
                qualifiedCaps = false;
                continue;
            }
            if (!qualifiedCaps) continue;
            if (firstCap && !AddCap(first, -1d)) return null;
            if (lastCap && !AddCap(last, 1d)) return null;

            bool AddCap(double distance, double direction)
            {
                double cx = x + dx * (distance / length), cy = y + dy * (distance / length);
                double radius = widthPoints / 2d * (1d + SvgCircleKappa);
                if (!double.IsFinite(cx - radius) || !double.IsFinite(cx + radius)
                    || !double.IsFinite(cy - radius) || !double.IsFinite(cy + radius)) return false;
                caps.Add(new SvgRoundDashCap(cx, cy, direction * dx / mappedLength, direction * dy / mappedLength));
                return true;
            }
        }
        if (hasPaintedDash) solidLines.Clear();
        return new SvgRoundDashPaint(qualifiedCaps ? caps : null, solidLines);
    }
    private static SvgTransform? ApplySvgStrokeViewport(PdfGraphicsBuilder graphics, double imageX, double imageY, double imageHeight, double scaleX, double scaleY)
    {
        SvgTransform? viewport = ReadSvgStrokeViewport(imageX, imageY, imageHeight, scaleX, scaleY);
        if (viewport is { } transform)
        {
            graphics.Transform(transform.M11, 0d, 0d, transform.M22, transform.OffsetX, transform.OffsetY);
        }
        return viewport;
    }

    private static SvgTransform? ReadSvgStrokeViewport(double imageX, double imageY, double imageHeight, double scaleX, double scaleY)
    {
        // RV07-S1: keep viewport anisotropy in the PDF CTM so it stretches
        // the stroke's normal, caps and dash lengths as well as the path.
        // Normalize around the existing mean scale to retain page-sized path
        // coordinates. Bound the error from the writer's three-decimal CTM;
        // extreme aspect ratios retain the previous scalar approximation.
        double meanScale = (scaleX + scaleY) / 2d;
        if (double.IsFinite(meanScale) && meanScale > 0d && Math.Abs(scaleX - scaleY) > meanScale * .000001d)
        {
            double factorX = scaleX / meanScale, factorY = scaleY / meanScale;
            double printedX = Math.Round(factorX, 3), printedY = Math.Round(factorY, 3);
            if (printedX > 0d && printedY > 0d
                && Math.Abs(printedX - factorX) <= factorX * .001d
                && Math.Abs(printedY - factorY) <= factorY * .001d
                && double.IsFinite(imageY + imageHeight))
            {
                return new SvgTransform(printedX, 0d, 0d, printedY, imageX, imageY + imageHeight);
            }
        }
        return null;
    }

    private static PdfRectangle? ReadSvgObjectPaintBounds(PdfRectangle frame, SvgPathBounds controls, SvgStroke stroke,
        double strokeWidth, double sourceMinX, double sourceMinY, double scaleX, double scaleY)
    {
        // A PDF Form clips to its bounding box. Where the picture has no
        // viewport clip, retain the full control hull and stroke envelope.
        var viewport = ReadSvgStrokeViewport(frame.X, frame.Y, frame.Height, scaleX, scaleY);
        double tip = Math.Max(stroke.LineJoin == 0 ? stroke.MiterLimit : 1d, stroke.LineCap == 2 ? Math.Sqrt(2d) : 1d);
        double radius = (strokeWidth / 2d) * tip;
        double marginX = radius * (viewport?.M11 ?? 1d) + .002d;
        double marginY = radius * (viewport?.M22 ?? 1d) + .002d;
        double top = frame.Y + frame.Height;
        double left = Math.Min(frame.X, frame.X + (controls.MinX - sourceMinX) * scaleX - marginX);
        double right = Math.Max(frame.X + frame.Width, frame.X + (controls.MaxX - sourceMinX) * scaleX + marginX);
        double bottom = Math.Min(frame.Y, top - (controls.MaxY - sourceMinY) * scaleY - marginY);
        top = Math.Max(top, top - (controls.MinY - sourceMinY) * scaleY + marginY);
        double width = right - left, height = top - bottom;
        if (!double.IsFinite(width) || !double.IsFinite(height)) { return null; }
        return new PdfRectangle(left, bottom, width, height);
    }
    // RV07: radial gradients paint concentric ellipse rings, largest first.
    // Focal points diagnose and render centered; radii resolve per axis like
    // the linear vector mapping, with the gradient transform baked at read.
    private static SvgRadialGradient? ReadSvgRadialGradient(SvgGradientDefinition definition, SvgGradientRamp ramp,
        SvgTransform transform, double? width, double? height, double? diagonal, SvgGradientSpread spread)
    {
        if (!definition.Cx.TryResolve((width ?? 1d) / 2d, width, out double cx)
            || !definition.Cy.TryResolve((height ?? 1d) / 2d, height, out double cy)
            || !definition.Radius.TryResolve((diagonal ?? 1d) / 2d, diagonal, out double radius)
            || !definition.Fx.TryResolve(cx, width, out double fx)
            || !definition.Fy.TryResolve(cy, height, out double fy)
            || radius <= 0d)
        {
            return null;
        }

        (double rawCx, double rawCy) = transform.Apply(cx, cy);
        double rawRadius = radius;
        if (width.HasValue)
        {
            (double edgeX, double edgeY) = transform.Apply(cx + radius, cy + radius);
            rawRadius = Math.Max(Math.Abs(edgeX - rawCx), Math.Abs(edgeY - rawCy));
        }
        if (!double.IsFinite(rawCx) || !double.IsFinite(rawCy) || !double.IsFinite(rawRadius))
        {
            return null;
        }

        return new SvgRadialGradient(rawCx, rawCy, rawRadius, fx != cx || fy != cy, ramp.Stops, width.HasValue, spread)
        {
            UniformStopOpacity = ramp.UniformOpacity,
            HasNumericStopOpacity = ramp.HasNumericOpacity,
            HasIdentityGradientTransform = transform.IsIdentity,
            IgnoredStopOpacityCount = ramp.IgnoredOpacityCount
        };
    }
    private static bool TryRenderSvgRadialGradientPath(PdfGraphicsBuilder graphics, IReadOnlyList<SvgPathCommand> commands, SvgRadialGradient radial, SvgTransform transform, bool evenOddFill, double opacity, SvgPathBounds pathBounds, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY, bool allowAlphaMask, out bool varyingAlphaAdmitted, Action<int> admitSamples, CancellationToken cancellationToken)
    {
        varyingAlphaAdmitted = false;
        double pathWidth = Math.Max(0.001d, pathBounds.MaxX - pathBounds.MinX);
        double pathHeight = Math.Max(0.001d, pathBounds.MaxY - pathBounds.MinY);
        double centerX;
        double centerY;
        double radiusX;
        double radiusY;
        bool uniformPathScale = transform.HasPositiveUniformScale;
        if (radial.IsUserSpace)
        {
            (centerX, centerY) = transform.Apply(radial.Cx, radial.Cy);
            radiusX = radial.Radius * (uniformPathScale ? transform.M11 : 1d);
            radiusY = radiusX;
        }
        else
        {
            centerX = pathBounds.MinX + radial.Cx * pathWidth;
            centerY = pathBounds.MinY + radial.Cy * pathHeight;
            radiusX = radial.Radius * pathWidth;
            radiusY = radial.Radius * pathHeight;
        }
        if (!AreSvgGradientBoundsMappable(pathBounds, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY)
            || !AreSvgGradientBoundsMappable(new SvgPathBounds(centerX - radiusX, centerY - radiusY, centerX + radiusX, centerY + radiusY), minX, minY, imageX, imageY, imageHeight, scaleX, scaleY)
            || !double.IsFinite(centerX - radiusX) || !double.IsFinite(centerX + radiusX)
            || !double.IsFinite(centerY - radiusY) || !double.IsFinite(centerY + radiusY)
            || !double.IsFinite(radiusX * scaleX) || !double.IsFinite(radiusY * scaleY))
        {
            return false;
        }
        graphics.SaveState();
        if (!TryAppendSvgPath(graphics, commands, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY))
        {
            graphics.RestoreState();
            return false;
        }
        graphics.ClipCurrentPath(evenOddFill);
        if (opacity < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(opacity, 1d);
        }
        double printedRadiusX = Math.Abs(radiusX * scaleX);
        double printedRadiusY = Math.Abs(radiusY * scaleY);
        // Retain sampling for sub-precision radii or distinct stop intervals.
        // One clipped shading applies alpha once instead of accumulating it
        // over nested sampled rings and the padded background.
        double padEnd = radial.Spread == SvgGradientSpread.Pad && radial.Stops.Count >= 2 &&
            radial.Stops[^2].Offset == 1d ? 2d : 1d;
        bool nativeShading = printedRadiusX >= 0.001d && printedRadiusY >= 0.001d &&
            double.IsFinite(printedRadiusX * padEnd) && double.IsFinite(printedRadiusY * padEnd) &&
            HaveRepresentableSvgStops(radial.Stops);
        int cycleCount = 1;
        if (nativeShading && radial.Spread != SvgGradientSpread.Pad)
        {
            double dx = Math.Max(Math.Abs(pathBounds.MinX - centerX), Math.Abs(pathBounds.MaxX - centerX)) / radiusX;
            double dy = Math.Max(Math.Abs(pathBounds.MinY - centerY), Math.Abs(pathBounds.MaxY - centerY)) / radiusY;
            double requiredCycles = Math.Max(1d, Math.Ceiling(Math.Sqrt(dx * dx + dy * dy)));
            double normalizedStopCount = radial.Stops.Count +
                (radial.Stops[0].Offset > 0d ? 1d : 0d) + (radial.Stops[^1].Offset < 1d ? 1d : 0d);
            // Bound stitched-function expansion independently of source units.
            nativeShading = double.IsFinite(requiredCycles) && requiredCycles <= 128d &&
                normalizedStopCount * requiredCycles <= 256d &&
                double.IsFinite(printedRadiusX * requiredCycles) && double.IsFinite(printedRadiusY * requiredCycles);
            if (nativeShading) { cycleCount = (int)requiredCycles; }
        }
        if (!nativeShading && radial.Spread == SvgGradientSpread.Pad)
        {
            // Pad extends the final stop through the clipped shape beyond the outer ring.
            RgbColor outerColor = radial.Stops[^1].Color;
            graphics.SetFillRgb(outerColor.Red, outerColor.Green, outerColor.Blue);
            if (TryAppendSvgPath(graphics, commands, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY))
            {
                graphics.FillCurrentPath();
            }
        }
        if (nativeShading)
        {
            double maskX = imageX + (pathBounds.MinX - minX) * scaleX;
            double maskY = imageY + imageHeight - (pathBounds.MaxY - minY) * scaleY;
            double maskWidth = Math.Max(0.001d, imageX + (pathBounds.MaxX - minX) * scaleX - maskX);
            double maskHeight = Math.Max(0.001d, imageY + imageHeight - (pathBounds.MinY - minY) * scaleY - maskY);
            if (allowAlphaMask && radial.UniformStopOpacity is null && radial.HasNumericStopOpacity &&
                radial.HasIdentityGradientTransform && !radial.HasFocal &&
                (!radial.IsUserSpace || uniformPathScale) &&
                radial.Spread == SvgGradientSpread.Pad && radial.Stops.Count <= 256 &&
                double.IsFinite(maskWidth) && double.IsFinite(maskHeight) &&
                double.IsFinite(maskX + maskWidth) && double.IsFinite(maskY + maskHeight))
            {
                var mask = new PdfGraphicsBuilder();
                mask.SaveState();
                mask.Transform(printedRadiusX * padEnd, 0d, 0d, printedRadiusY * padEnd,
                    imageX + (centerX - minX) * scaleX,
                    imageY + imageHeight - (centerY - minY) * scaleY);
                mask.PaintShading(new PdfRadialShading(CreateSvgShadingStops(radial.Stops, 0d, padEnd, opacityMask: true)));
                mask.RestoreState();
                graphics.SetVectorLuminositySoftMask(mask, new PdfRectangle(maskX, maskY, maskWidth, maskHeight), opacity, 1d);
                varyingAlphaAdmitted = true;
            }
            graphics.SaveState();
            graphics.Transform(printedRadiusX * cycleCount * padEnd, 0d, 0d, printedRadiusY * cycleCount * padEnd,
                imageX + (centerX - minX) * scaleX,
                imageY + imageHeight - (centerY - minY) * scaleY);
            graphics.PaintShading(new PdfRadialShading(CreateSvgShadingStops(radial.Stops, 0d, padEnd),
                cycleCount, radial.Spread == SvgGradientSpread.Reflect));
            graphics.RestoreState();
        }
        else
        {
            int ringCount = Math.Clamp((int)Math.Ceiling(Math.Max(radiusX, radiusY) / 2d), 16, 128);
            admitSamples(ringCount);
            for (int ring = 0; ring < ringCount; ring++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                double fraction = (double)(ringCount - ring) / ringCount;
                RgbColor color = SampleSvgStops(radial.Stops, (double)(ringCount - ring - 1) / (ringCount - 1), radial.Spread);
                graphics.SetFillRgb(color.Red, color.Green, color.Blue);
                string ringData = ConvertSvgEllipseBody(centerX, centerY, radiusX * fraction, radiusY * fraction);
                if (TryAppendSvgPath(graphics, ParseSvgPathData(ringData).Commands, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY))
                {
                    graphics.FillCurrentPath();
                }
            }
        }
        if (opacity < 1d)
        {
            graphics.RestoreState();
        }
        graphics.RestoreState();
        return true;
    }
    private static RgbColor SampleSvgStops(IReadOnlyList<SvgGradientStop> stops, double t, SvgGradientSpread spread)
    {
        double offset = spread switch
        {
            SvgGradientSpread.Repeat => t - Math.Floor(t),
            SvgGradientSpread.Reflect => ReflectSvgOffset(t),
            _ => t,
        };
        SvgGradientStop previous = stops[0];
        for (int index = 1; index < stops.Count; index++)
        {
            SvgGradientStop next = stops[index];
            if (offset < next.Offset)
            {
                double span = next.Offset - previous.Offset;
                double amount = span <= PptxTextMetricRules.TextStateTolerance ? 0d : (offset - previous.Offset) / span;
                return InterpolateSvgGradientStop(previous.Color, next.Color, Math.Clamp(amount, 0d, 1d));
            }
            previous = next;
        }
        return previous.Color;
    }
    private static double ReflectSvgOffset(double value)
    {
        double wrapped = value % 2d;
        if (wrapped < 0d)
        {
            wrapped += 2d;
        }
        return wrapped > 1d ? 2d - wrapped : wrapped;
    }
    private static RgbColor InterpolateSvgGradientStop(RgbColor left, RgbColor right, double amount)
    {
        return new RgbColor(
            ToSvgGradientByte(left.Red + (right.Red - left.Red) * amount),
            ToSvgGradientByte(left.Green + (right.Green - left.Green) * amount),
            ToSvgGradientByte(left.Blue + (right.Blue - left.Blue) * amount));
    }
    private static byte ToSvgGradientByte(double value) => (byte)Math.Clamp((int)Math.Round(value), byte.MinValue, byte.MaxValue);
    private static RgbColor SampleSvgGradient(SvgGradient gradient, double x, double y)
    {
        double dx = gradient.X2 - gradient.X1;
        double dy = gradient.Y2 - gradient.Y1;
        double lengthSquared = dx * dx + dy * dy;
        double offset = lengthSquared <= PptxTextMetricRules.TextStateTolerance
            ? 0d
            : ((x - gradient.X1) * dx + (y - gradient.Y1) * dy) / lengthSquared;
        return SampleSvgStops(gradient.Stops, offset, gradient.Spread);
    }

    private readonly record struct SvgPathCommand(char Kind, double[] Arguments);

    private static (List<SvgPathCommand> Commands, bool Complete) ReadTransformedSvgPath(string data, SvgTransform transform)
    {
        var (commands, complete) = ParseSvgPathData(data);
        return (TransformSvgCommands(commands, transform), complete);
    }

    private static bool SvgPathMayHaveInterior(IReadOnlyList<SvgPathCommand> commands)
    {
        if (commands.Count == 0 || commands[0].Kind != 'M') return true;

        bool hasSegment = false;
        // Closing a single segment traverses the same edge in reverse.
        // Two edges or a curve can enclose an area when filled.
        foreach (SvgPathCommand command in commands)
        {
            if (!command.Arguments.All(double.IsFinite)) return true;
            switch (command.Kind)
            {
                case 'M':
                case 'Z':
                    hasSegment = false;
                    break;
                case 'L':
                case 'H':
                case 'V':
                    if (hasSegment) return true;
                    hasSegment = true;
                    break;
                default:
                    return true;
            }
        }
        return false;
    }

    // Parses the longest supported prefix of SVG path data into absolute
    // commands. Complete is false when parsing stops early: missing numbers,
    // an unsupported command, or numbers trailing a close. Bounds require a
    // complete parse, while emission walks the prefix and reports whether a
    // move was emitted, preserving the previous partial-emission behavior.
    // Known divergence: numbers after Z used to spin forever because Z
    // consumed no tokens; the parser now stops there instead of hanging.
    private static (List<SvgPathCommand> Commands, bool Complete) ParseSvgPathData(string data)
    {
        MatchCollection tokens = SvgPathTokenRegex().Matches(data);
        var commands = new List<SvgPathCommand>();
        int index = 0;
        char command = '\0';
        char previousKind = '\0';
        double currentX = 0d;
        double currentY = 0d;
        double startX = 0d;
        double startY = 0d;
        double quadraticControlX = 0d;
        double quadraticControlY = 0d;
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
                return (commands, false);
            }

            bool relative = char.IsLower(command);
            char kind = char.ToUpperInvariant(command);
            switch (kind)
            {
                case 'M':
                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out currentX, out currentY))
                    {
                        return (commands, false);
                    }

                    startX = currentX;
                    startY = currentY;
                    commands.Add(new SvgPathCommand('M', new[] { currentX, currentY }));
                    command = relative ? 'l' : 'L';
                    break;
                case 'L':
                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out currentX, out currentY))
                    {
                        return (commands, false);
                    }

                    commands.Add(new SvgPathCommand('L', new[] { currentX, currentY }));
                    break;
                case 'H':
                    if (!TryReadSvgNumber(tokens, ref index, out double h))
                    {
                        return (commands, false);
                    }

                    currentX = relative ? currentX + h : h;
                    commands.Add(new SvgPathCommand('H', new[] { currentX }));
                    break;
                case 'V':
                    if (!TryReadSvgNumber(tokens, ref index, out double v))
                    {
                        return (commands, false);
                    }

                    currentY = relative ? currentY + v : v;
                    commands.Add(new SvgPathCommand('V', new[] { currentY }));
                    break;
                case 'C':
                case 'S':
                {
                    double c1x = currentX;
                    double c1y = currentY;
                    if (kind == 'C')
                    {
                        if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out c1x, out c1y))
                        {
                            return (commands, false);
                        }
                    }
                    else if (previousKind is 'C' or 'S')
                    {
                        // Reflection follows the authored curve family. Quadratics also emit C.
                        double[] previous = commands[^1].Arguments;
                        c1x += currentX - previous[2];
                        c1y += currentY - previous[3];
                    }

                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out double c2x, out double c2y) ||
                        !TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out currentX, out currentY))
                    {
                        return (commands, false);
                    }

                    commands.Add(new SvgPathCommand('C', new[] { c1x, c1y, c2x, c2y, currentX, currentY }));
                    break;
                }
                case 'Q':
                case 'T':
                {
                    double controlX = currentX;
                    double controlY = currentY;
                    if (kind == 'Q')
                    {
                        if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out controlX, out controlY))
                        {
                            return (commands, false);
                        }
                    }
                    else if (previousKind is 'Q' or 'T')
                    {
                        controlX += currentX - quadraticControlX;
                        controlY += currentY - quadraticControlY;
                    }

                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out double endX, out double endY))
                    {
                        return (commands, false);
                    }

                    BezierSegment curve = CreateQuadraticBezierSegment((currentX, currentY), (controlX, controlY), (endX, endY));
                    commands.Add(new SvgPathCommand('C', [curve.Control1X, curve.Control1Y, curve.Control2X, curve.Control2Y, endX, endY]));
                    quadraticControlX = controlX;
                    quadraticControlY = controlY;
                    currentX = endX;
                    currentY = endY;
                    break;
                }
                case 'Z':
                    if (index < tokens.Count)
                    {
                        string next = tokens[index].Value;
                        if (next.Length != 1 || !char.IsLetter(next[0]))
                        {
                            return (commands, false);
                        }
                    }

                    currentX = startX;
                    currentY = startY;
                    commands.Add(new SvgPathCommand('Z', Array.Empty<double>()));
                    break;
                default:
                    return (commands, false);
            }
            previousKind = kind;
        }

        return (commands, true);
    }

    private static bool TryReadSvgPathBounds(IReadOnlyList<SvgPathCommand> commands, out SvgPathBounds bounds, out SvgPathBounds controlBounds)
    {
        bounds = default;
        controlBounds = default;
        if (commands.Count == 0)
        {
            return false;
        }

        double currentX = 0d;
        double currentY = 0d;
        double startX = 0d;
        double startY = 0d;
        double minX = double.PositiveInfinity;
        double minY = double.PositiveInfinity;
        double maxX = double.NegativeInfinity;
        double maxY = double.NegativeInfinity;
        double minControlX = double.PositiveInfinity;
        double minControlY = double.PositiveInfinity;
        double maxControlX = double.NegativeInfinity;
        double maxControlY = double.NegativeInfinity;

        void Include(double x, double y)
        {
            minX = Math.Min(minX, x);
            minY = Math.Min(minY, y);
            maxX = Math.Max(maxX, x);
            maxY = Math.Max(maxY, y);
        }

        foreach (SvgPathCommand pathCommand in commands)
        {
            switch (pathCommand.Kind)
            {
                case 'M':
                    currentX = pathCommand.Arguments[0];
                    currentY = pathCommand.Arguments[1];
                    startX = currentX;
                    startY = currentY;
                    Include(currentX, currentY);
                    break;
                case 'L':
                    currentX = pathCommand.Arguments[0];
                    currentY = pathCommand.Arguments[1];
                    Include(currentX, currentY);
                    break;
                case 'H':
                    currentX = pathCommand.Arguments[0];
                    Include(currentX, currentY);
                    break;
                case 'V':
                    currentY = pathCommand.Arguments[0];
                    Include(currentX, currentY);
                    break;
                case 'C':
                {
                    double[] points = pathCommand.Arguments;
                    var (curveMinX, curveMaxX) = GetSvgCubicExtent(currentX, points[0], points[2], points[4]);
                    var (curveMinY, curveMaxY) = GetSvgCubicExtent(currentY, points[1], points[3], points[5]);
                    Include(curveMinX, curveMinY);
                    Include(curveMaxX, curveMaxY);
                    minControlX = Math.Min(minControlX, Math.Min(points[0], points[2]));
                    minControlY = Math.Min(minControlY, Math.Min(points[1], points[3]));
                    maxControlX = Math.Max(maxControlX, Math.Max(points[0], points[2]));
                    maxControlY = Math.Max(maxControlY, Math.Max(points[1], points[3]));
                    currentX = pathCommand.Arguments[4];
                    currentY = pathCommand.Arguments[5];
                    break;
                }
                case 'Z':
                    currentX = startX;
                    currentY = startY;
                    break;
                default:
                    bounds = default;
                    return false;
            }
        }

        // The PDF still emits the control points. Keep their existing finite
        // mapping check separate from the gradient's painted extent.
        controlBounds = new SvgPathBounds(Math.Min(minX, minControlX), Math.Min(minY, minControlY),
            Math.Max(maxX, maxControlX), Math.Max(maxY, maxControlY));
        if (double.IsInfinity(controlBounds.MinX) || double.IsInfinity(controlBounds.MinY) ||
            double.IsInfinity(controlBounds.MaxX) || double.IsInfinity(controlBounds.MaxY))
        {
            bounds = default;
            return false;
        }

        bounds = new SvgPathBounds(minX, minY, maxX, maxY);
        return true;
    }

    private static (double Min, double Max) GetSvgCubicExtent(double start, double first, double second, double end)
    {
        double min = Math.Min(start, end);
        double max = Math.Max(start, end);
        double hullMin = Math.Min(min, Math.Min(first, second));
        double hullMax = Math.Max(max, Math.Max(first, second));
        if ((min == hullMin && max == hullMax) || !double.IsFinite(hullMin) || !double.IsFinite(hullMax))
        {
            return (hullMin, hullMax);
        }

        // Solve the derivative on normalized coordinates so differences and
        // the discriminant remain finite for large authored curves.
        double scale = Math.Max(Math.Abs(hullMin), Math.Abs(hullMax));
        double p0 = start / scale, p1 = first / scale, p2 = second / scale, p3 = end / scale;
        double d0 = p1 - p0, d1 = p2 - p1, d2 = p3 - p2;
        double a = d0 - 2d * d1 + d2;
        double b = 2d * (d1 - d0);
        if (a == 0d)
        {
            if (b != 0d) IncludeAt(-d0 / b);
        }
        else
        {
            double discriminant = b * b - 4d * a * d0;
            if (discriminant > 0d)
            {
                // Recover the smaller root without cancellation in the numerator.
                double q = -0.5d * (b + Math.CopySign(Math.Sqrt(discriminant), b));
                IncludeAt(q / a);
                IncludeAt(d0 / q);
            }
        }
        return (min, max);

        void IncludeAt(double t)
        {
            if (t <= 0d || t >= 1d) return;
            double u = 1d - t;
            double normalized = u * u * u * p0 + 3d * u * u * t * p1 + 3d * u * t * t * p2 + t * t * t * p3;
            double value = Math.Clamp(normalized, -1d, 1d) * scale;
            min = Math.Min(min, value);
            max = Math.Max(max, value);
        }
    }

    private static bool TryAppendSvgPath(PdfGraphicsBuilder graphics, IReadOnlyList<SvgPathCommand> commands, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY, SvgTransform? strokeViewport = null)
    {
        double currentX = 0d;
        double currentY = 0d;
        double startX = 0d;
        double startY = 0d;
        bool hasPath = false;
        foreach (SvgPathCommand pathCommand in commands)
        {
            switch (pathCommand.Kind)
            {
                case 'M':
                    currentX = pathCommand.Arguments[0];
                    currentY = pathCommand.Arguments[1];
                    startX = currentX;
                    startY = currentY;
                    if (!TryMapSvgPoint(currentX, currentY, out double moveX, out double moveY))
                    {
                        return hasPath;
                    }

                    graphics.MoveTo(moveX, moveY);
                    hasPath = true;
                    break;
                case 'L':
                    {
                    currentX = pathCommand.Arguments[0];
                    currentY = pathCommand.Arguments[1];
                    if (!TryMapSvgPoint(currentX, currentY, out double lineX, out double lineY))
                    {
                        return hasPath;
                    }

                    graphics.LineTo(lineX, lineY);
                    break;
                    }
                case 'H':
                    {
                    currentX = pathCommand.Arguments[0];
                    if (!TryMapSvgPoint(currentX, currentY, out double lineX, out double lineY))
                    {
                        return hasPath;
                    }

                    graphics.LineTo(lineX, lineY);
                    break;
                    }
                case 'V':
                    {
                    currentY = pathCommand.Arguments[0];
                    if (!TryMapSvgPoint(currentX, currentY, out double lineX, out double lineY))
                    {
                        return hasPath;
                    }

                    graphics.LineTo(lineX, lineY);
                    break;
                    }
                case 'C':
                    if (!TryMapSvgPoint(pathCommand.Arguments[0], pathCommand.Arguments[1], out double curve1X, out double curve1Y)
                        || !TryMapSvgPoint(pathCommand.Arguments[2], pathCommand.Arguments[3], out double curve2X, out double curve2Y)
                        || !TryMapSvgPoint(pathCommand.Arguments[4], pathCommand.Arguments[5], out double curveEndX, out double curveEndY))
                    {
                        return hasPath;
                    }

                    graphics.CurveTo(curve1X, curve1Y, curve2X, curve2Y, curveEndX, curveEndY);
                    currentX = pathCommand.Arguments[4];
                    currentY = pathCommand.Arguments[5];
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

        bool TryMapSvgPoint(double x, double y, out double mappedX, out double mappedY)
        {
            return TryMapSvgPicturePoint(x, y, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY, strokeViewport, out mappedX, out mappedY);
        }
    }
    private static bool TryMapSvgPicturePoint(double x, double y, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY, SvgTransform? strokeViewport, out double mappedX, out double mappedY)
    {
        // Stop composed overflow before emitting a non-finite PDF number.
        mappedX = imageX + (x - minX) * scaleX;
        mappedY = imageY + imageHeight - (y - minY) * scaleY;
        if (!double.IsFinite(mappedX) || !double.IsFinite(mappedY)) return false;
        if (strokeViewport is SvgTransform viewport)
        {
            mappedX = (mappedX - viewport.OffsetX) / viewport.M11;
            mappedY = (mappedY - viewport.OffsetY) / viewport.M22;
        }
        return double.IsFinite(mappedX) && double.IsFinite(mappedY);
    }

    // RV07: basic shapes convert to path data and paint through the shared
    // path pipeline. Circles, ellipses and rounded corners use the kappa
    // bezier approximation as a platform-independent literal; spec-invalid
    // geometry returns null and renders nothing like Office. Ordinary text
    // shares this picture's placement and isolated vector paint above.
    private const double SvgCircleKappa = 0.5522847498;
    private static bool IsSvgPaintableElement(XElement element)
    {
        return element.Name.LocalName is "path" or "rect" or "circle" or "ellipse" or "line" or "polyline" or "polygon";
    }
    private static string? ConvertSvgShapeToPathData(XElement element, double viewWidth, double viewHeight)
    {
        return element.Name.LocalName switch
        {
            "rect" => ConvertSvgRectToPathData(element, viewWidth, viewHeight),
            "circle" => ConvertSvgCircleToPathData(element, viewWidth, viewHeight),
            "ellipse" => ConvertSvgEllipseToPathData(element, viewWidth, viewHeight),
            "line" => ConvertSvgLineToPathData(element, viewWidth, viewHeight),
            "polyline" => ConvertSvgPointsToPathData(element, false),
            "polygon" => ConvertSvgPointsToPathData(element, true),
            _ => (string?)element.Attribute("d"),
        };
    }
    private static bool TryReadSvgShapeCoordinate(XElement element, string name, double percentageBasis, out double value)
    {
        string? text = (string?)element.Attribute(name);
        if (string.IsNullOrWhiteSpace(text))
        {
            value = 0d;
            return true;
        }
        return TryReadSvgViewportLength(element, text, percentageBasis, out value);
    }

    private static bool TryReadSvgViewportLength(XElement element, ReadOnlySpan<char> text, double percentageBasis, out double value)
    {
        text = text.Trim();
        if (!text.EndsWith("%", StringComparison.Ordinal)) return TryReadSvgNumberOrPixels(text, out value);

        value = 0d;
        // Nested viewports require their own dimensions.
        if (text.Length == 1 || char.IsWhiteSpace(text[^2]) ||
            element.Ancestors().Any(ancestor => ancestor.Name.LocalName == "svg" && ancestor.Parent is not null) ||
            !double.TryParse(text[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out double percent))
        {
            return false;
        }

        value = (percent / 100d) * percentageBasis;
        return double.IsFinite(value);
    }

    private static string SvgPathNumber(double value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }

    private static bool TryReadSvgPairedRadii(XElement element, double viewWidth, double viewHeight, out double rx, out double ry)
    {
        string? horizontal = (string?)element.Attribute("rx");
        string? vertical = (string?)element.Attribute("ry");
        bool autoX = horizontal is null || horizontal.AsSpan().Trim().Equals("auto", StringComparison.OrdinalIgnoreCase);
        bool autoY = vertical is null || vertical.AsSpan().Trim().Equals("auto", StringComparison.OrdinalIgnoreCase);
        rx = 0d;
        ry = 0d;
        if (!autoX && (!TryReadSvgViewportLength(element, horizontal, viewWidth, out rx) || rx < 0d) ||
            !autoY && (!TryReadSvgViewportLength(element, vertical, viewHeight, out ry) || ry < 0d))
        {
            return false;
        }

        // Resolve before the rectangle's own per-axis clamps.
        if (autoX) rx = ry;
        if (autoY) ry = rx;
        return true;
    }

    private static string? ConvertSvgRectToPathData(XElement element, double viewWidth, double viewHeight)
    {
        if (!TryReadSvgShapeCoordinate(element, "x", viewWidth, out double x)
            || !TryReadSvgShapeCoordinate(element, "y", viewHeight, out double y)
            || !TryReadSvgShapeCoordinate(element, "width", viewWidth, out double width)
            || !TryReadSvgShapeCoordinate(element, "height", viewHeight, out double height))
        {
            return null;
        }
        if (width <= 0d || height <= 0d)
        {
            return null;
        }
        if (!TryReadSvgPairedRadii(element, viewWidth, viewHeight, out double rx, out double ry))
        {
            return null;
        }
        rx = Math.Min(rx, width / 2d);
        ry = Math.Min(ry, height / 2d);
        if (rx == 0d || ry == 0d)
        {
            return "M" + SvgPathNumber(x) + " " + SvgPathNumber(y) + "H" + SvgPathNumber(x + width) + "V" + SvgPathNumber(y + height) + "H" + SvgPathNumber(x) + "Z";
        }
        double kx = SvgCircleKappa * rx;
        double ky = SvgCircleKappa * ry;
        return "M" + SvgPathNumber(x + rx) + " " + SvgPathNumber(y)
            + "H" + SvgPathNumber(x + width - rx)
            + "C" + SvgPathNumber(x + width - rx + kx) + " " + SvgPathNumber(y) + " " + SvgPathNumber(x + width) + " " + SvgPathNumber(y + ry - ky) + " " + SvgPathNumber(x + width) + " " + SvgPathNumber(y + ry)
            + "V" + SvgPathNumber(y + height - ry)
            + "C" + SvgPathNumber(x + width) + " " + SvgPathNumber(y + height - ry + ky) + " " + SvgPathNumber(x + width - rx + kx) + " " + SvgPathNumber(y + height) + " " + SvgPathNumber(x + width - rx) + " " + SvgPathNumber(y + height)
            + "H" + SvgPathNumber(x + rx)
            + "C" + SvgPathNumber(x + rx - kx) + " " + SvgPathNumber(y + height) + " " + SvgPathNumber(x) + " " + SvgPathNumber(y + height - ry + ky) + " " + SvgPathNumber(x) + " " + SvgPathNumber(y + height - ry)
            + "V" + SvgPathNumber(y + ry)
            + "C" + SvgPathNumber(x) + " " + SvgPathNumber(y + ry - ky) + " " + SvgPathNumber(x + rx - kx) + " " + SvgPathNumber(y) + " " + SvgPathNumber(x + rx) + " " + SvgPathNumber(y)
            + "Z";
    }
    private static string? ConvertSvgCircleToPathData(XElement element, double viewWidth, double viewHeight)
    {
        if (!TryReadSvgShapeCoordinate(element, "cx", viewWidth, out double cx)
            || !TryReadSvgShapeCoordinate(element, "cy", viewHeight, out double cy)
            || !TryReadSvgShapeCoordinate(element, "r", GetSvgNormalizedDiagonal(viewWidth, viewHeight), out double radius)
            || radius <= 0d)
        {
            return null;
        }
        return ConvertSvgEllipseBody(cx, cy, radius, radius);
    }
    private static string? ConvertSvgEllipseToPathData(XElement element, double viewWidth, double viewHeight)
    {
        if (!TryReadSvgShapeCoordinate(element, "cx", viewWidth, out double cx)
            || !TryReadSvgShapeCoordinate(element, "cy", viewHeight, out double cy)
            || !TryReadSvgPairedRadii(element, viewWidth, viewHeight, out double rx, out double ry)
            || rx <= 0d || ry <= 0d)
        {
            return null;
        }
        return ConvertSvgEllipseBody(cx, cy, rx, ry);
    }
    private static string ConvertSvgEllipseBody(double cx, double cy, double rx, double ry)
    {
        double kx = SvgCircleKappa * rx;
        double ky = SvgCircleKappa * ry;
        return "M" + SvgPathNumber(cx + rx) + " " + SvgPathNumber(cy)
            + "C" + SvgPathNumber(cx + rx) + " " + SvgPathNumber(cy + ky) + " " + SvgPathNumber(cx + kx) + " " + SvgPathNumber(cy + ry) + " " + SvgPathNumber(cx) + " " + SvgPathNumber(cy + ry)
            + "C" + SvgPathNumber(cx - kx) + " " + SvgPathNumber(cy + ry) + " " + SvgPathNumber(cx - rx) + " " + SvgPathNumber(cy + ky) + " " + SvgPathNumber(cx - rx) + " " + SvgPathNumber(cy)
            + "C" + SvgPathNumber(cx - rx) + " " + SvgPathNumber(cy - ky) + " " + SvgPathNumber(cx - kx) + " " + SvgPathNumber(cy - ry) + " " + SvgPathNumber(cx) + " " + SvgPathNumber(cy - ry)
            + "C" + SvgPathNumber(cx + kx) + " " + SvgPathNumber(cy - ry) + " " + SvgPathNumber(cx + rx) + " " + SvgPathNumber(cy - ky) + " " + SvgPathNumber(cx + rx) + " " + SvgPathNumber(cy)
            + "Z";
    }
    private static string? ConvertSvgLineToPathData(XElement element, double viewWidth, double viewHeight)
    {
        if (!TryReadSvgShapeCoordinate(element, "x1", viewWidth, out double x1)
            || !TryReadSvgShapeCoordinate(element, "y1", viewHeight, out double y1)
            || !TryReadSvgShapeCoordinate(element, "x2", viewWidth, out double x2)
            || !TryReadSvgShapeCoordinate(element, "y2", viewHeight, out double y2))
        {
            return null;
        }
        return "M" + SvgPathNumber(x1) + " " + SvgPathNumber(y1) + "L" + SvgPathNumber(x2) + " " + SvgPathNumber(y2);
    }
    private static string? ConvertSvgPointsToPathData(XElement element, bool closed)
    {
        System.Text.RegularExpressions.MatchCollection numbers = SvgNumberRegex().Matches((string?)element.Attribute("points") ?? string.Empty);
        // A trailing unpaired coordinate does not discard the complete points.
        int coordinateCount = numbers.Count - numbers.Count % 2;
        if (coordinateCount < 4)
        {
            return null;
        }
        var points = new List<double>();
        for (int index = 0; index < coordinateCount; index++)
        {
            if (!double.TryParse(numbers[index].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double coordinate))
            {
                return null;
            }
            points.Add(coordinate);
        }
        System.Text.StringBuilder data = new System.Text.StringBuilder();
        data.Append("M").Append(SvgPathNumber(points[0])).Append((char)32).Append(SvgPathNumber(points[1]));
        for (int index = 2; index < points.Count; index += 2)
        {
            data.Append("L").Append(SvgPathNumber(points[index])).Append((char)32).Append(SvgPathNumber(points[index + 1]));
        }
        if (closed)
        {
            data.Append("Z");
        }
        return data.ToString();
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

        // RV07: e999-style literals parse to infinity and cannot paint; stopping
        // keeps the supported prefix like a truncated path instead of throwing
        // non-finite PDF numbers into per-node recovery.
        if (!double.TryParse(tokens[index].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
            || !double.IsFinite(value))
        {
            return false;
        }

        index++;
        return true;
    }

    [GeneratedRegex(@"[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?", RegexOptions.CultureInvariant)]
    private static partial Regex SvgNumberRegex();

    // Retain unsupported command letters so their operands cannot continue
    // the preceding command.
    [GeneratedRegex(@"\p{L}|[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?", RegexOptions.CultureInvariant)]
    private static partial Regex SvgPathTokenRegex();

    [GeneratedRegex(@"[A-Za-z]+|[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?|[(),]", RegexOptions.CultureInvariant)]
    private static partial Regex SvgTransformTokenRegex();

    // Affine transforms are shared by bounds and paint. The first transform
    // in a list establishes the outer coordinate system, like a parent group.
    private readonly record struct SvgTransform(double M11, double M12, double M21, double M22, double OffsetX, double OffsetY)
    {
        public static SvgTransform Identity => new SvgTransform(1d, 0d, 0d, 1d, 0d, 0d);

        public bool HasPositiveUniformScale => M11 > 0d && M11 == M22 && M12 == 0d && M21 == 0d;

        public bool IsIdentity => HasPositiveUniformScale && M11 == 1d && OffsetX == 0d && OffsetY == 0d;

        public bool IsFinite => double.IsFinite(M11) && double.IsFinite(M12) && double.IsFinite(M21) &&
            double.IsFinite(M22) && double.IsFinite(OffsetX) && double.IsFinite(OffsetY);

        public (double X, double Y) Apply(double x, double y) => (M11 * x + M21 * y + OffsetX, M12 * x + M22 * y + OffsetY);

        public static SvgTransform Compose(SvgTransform outer, SvgTransform inner) => new SvgTransform(
            outer.M11 * inner.M11 + outer.M21 * inner.M12,
            outer.M12 * inner.M11 + outer.M22 * inner.M12,
            outer.M11 * inner.M21 + outer.M21 * inner.M22,
            outer.M12 * inner.M21 + outer.M22 * inner.M22,
            outer.M11 * inner.OffsetX + outer.M21 * inner.OffsetY + outer.OffsetX,
            outer.M12 * inner.OffsetX + outer.M22 * inner.OffsetY + outer.OffsetY);

        public static SvgTransform Rotation(double degrees, double centerX, double centerY)
        {
            double radians = DegreesToRadians(degrees);
            double cosine = Math.Cos(radians);
            double sine = Math.Sin(radians);
            SvgTransform spin = new SvgTransform(cosine, sine, -sine, cosine, 0d, 0d);
            SvgTransform toOrigin = new SvgTransform(1d, 0d, 0d, 1d, -centerX, -centerY);
            SvgTransform back = new SvgTransform(1d, 0d, 0d, 1d, centerX, centerY);
            return Compose(back, Compose(spin, toOrigin));
        }
    }

    private static bool TryParseSvgTransformList(string? text, out SvgTransform transform)
    {
        transform = SvgTransform.Identity;
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }

        MatchCollection tokens = SvgTransformTokenRegex().Matches(text);
        int index = 0;
        SvgTransform accumulated = SvgTransform.Identity;
        while (index < tokens.Count)
        {
            if (!TryParseSvgTransform(tokens, ref index, out SvgTransform next))
            {
                transform = SvgTransform.Identity;
                return false;
            }

            accumulated = SvgTransform.Compose(accumulated, next);
        }

        transform = accumulated;
        return true;
    }

    private static bool TryParseSvgTransform(MatchCollection tokens, ref int index, out SvgTransform transform)
    {
        transform = SvgTransform.Identity;
        if (index >= tokens.Count || tokens[index].Value.Length == 0 || !char.IsLetter(tokens[index].Value[0]))
        {
            return false;
        }

        string name = tokens[index].Value;
        index++;
        if (index >= tokens.Count || tokens[index].Value != "(")
        {
            return false;
        }

        index++;
        var args = new List<double>();
        while (true)
        {
            if (index >= tokens.Count)
            {
                return false;
            }

            string token = tokens[index].Value;
            if (token == ")")
            {
                index++;
                break;
            }

            if (token == ",")
            {
                index++;
                continue;
            }

            if (!TryReadSvgNumber(tokens, ref index, out double value))
            {
                return false;
            }

            args.Add(value);
        }

        double[] values = args.ToArray();
        switch (name)
        {
            case "translate" when values.Length == 1:
                transform = new SvgTransform(1d, 0d, 0d, 1d, values[0], 0d);
                return true;
            case "translate" when values.Length == 2:
                transform = new SvgTransform(1d, 0d, 0d, 1d, values[0], values[1]);
                return true;
            case "scale" when values.Length == 1:
                transform = new SvgTransform(values[0], 0d, 0d, values[0], 0d, 0d);
                return true;
            case "scale" when values.Length == 2:
                transform = new SvgTransform(values[0], 0d, 0d, values[1], 0d, 0d);
                return true;
            case "rotate" when values.Length == 1:
                transform = SvgTransform.Rotation(values[0], 0d, 0d);
                return true;
            case "rotate" when values.Length == 3:
                transform = SvgTransform.Rotation(values[0], values[1], values[2]);
                return true;
            case "skewX" when values.Length == 1:
                transform = new SvgTransform(1d, 0d, Math.Tan(DegreesToRadians(values[0])), 1d, 0d, 0d);
                return true;
            case "skewY" when values.Length == 1:
                transform = new SvgTransform(1d, Math.Tan(DegreesToRadians(values[0])), 0d, 1d, 0d, 0d);
                return true;
            case "matrix" when values.Length == 6:
                transform = new SvgTransform(values[0], values[1], values[2], values[3], values[4], values[5]);
                return true;
            default:
                transform = SvgTransform.Identity;
                return false;
        }
    }

    // RV07: transformed path commands remapped once for bounds and emission.
    // H and V normalize to L since transformed axis-aligned segments skew.
    private static List<SvgPathCommand> TransformSvgCommands(List<SvgPathCommand> commands, SvgTransform transform)
    {
        if (transform.IsIdentity)
        {
            return commands;
        }

        var mapped = new List<SvgPathCommand>(commands.Count);
        double sourceCurrentX = 0d;
        double sourceCurrentY = 0d;
        double sourceStartX = 0d;
        double sourceStartY = 0d;
        foreach (SvgPathCommand command in commands)
        {
            char kind = command.Kind;
            if (kind is 'M' or 'L')
            {
                sourceCurrentX = command.Arguments[0];
                sourceCurrentY = command.Arguments[1];
                if (kind == 'M')
                {
                    sourceStartX = sourceCurrentX;
                    sourceStartY = sourceCurrentY;
                }

                var (x, y) = transform.Apply(sourceCurrentX, sourceCurrentY);
                mapped.Add(new SvgPathCommand(kind, [x, y]));
            }
            else if (kind == 'H')
            {
                sourceCurrentX = command.Arguments[0];
                var (x, y) = transform.Apply(sourceCurrentX, sourceCurrentY);
                mapped.Add(new SvgPathCommand('L', [x, y]));
            }
            else if (kind == 'V')
            {
                sourceCurrentY = command.Arguments[0];
                var (x, y) = transform.Apply(sourceCurrentX, sourceCurrentY);
                mapped.Add(new SvgPathCommand('L', [x, y]));
            }
            else if (kind == 'C')
            {
                var (x0, y0) = transform.Apply(command.Arguments[0], command.Arguments[1]);
                var (x1, y1) = transform.Apply(command.Arguments[2], command.Arguments[3]);
                sourceCurrentX = command.Arguments[4];
                sourceCurrentY = command.Arguments[5];
                var (x, y) = transform.Apply(sourceCurrentX, sourceCurrentY);
                mapped.Add(new SvgPathCommand(kind, [x0, y0, x1, y1, x, y]));
            }
            else
            {
                if (kind == 'Z')
                {
                    sourceCurrentX = sourceStartX;
                    sourceCurrentY = sourceStartY;
                }
                mapped.Add(command);
            }
        }

        return mapped;
    }

    private static bool TryReadSvgPathTransform(XElement path, out SvgTransform transform)
    {
        transform = SvgTransform.Identity;
        foreach (XElement ancestor in path.Ancestors().Where(element => element.Name.LocalName == "g").Reverse())
        {
            if (!TryParseSvgTransformList((string?)ancestor.Attribute("transform"), out SvgTransform ancestorTransform))
            {
                transform = SvgTransform.Identity;
                return false;
            }

            transform = SvgTransform.Compose(transform, ancestorTransform);
        }

        if (!TryParseSvgTransformList((string?)path.Attribute("transform"), out SvgTransform local))
        {
            transform = SvgTransform.Identity;
            return false;
        }

        transform = SvgTransform.Compose(transform, local);
        // RV07: entries composed past double range cannot map; the path is
        // omitted like an unparseable transform instead of throwing downstream.
        if (!double.IsFinite(transform.M11) || !double.IsFinite(transform.M12) || !double.IsFinite(transform.M21) || !double.IsFinite(transform.M22) || !double.IsFinite(transform.OffsetX) || !double.IsFinite(transform.OffsetY))
        {
            transform = SvgTransform.Identity;
            return false;
        }

        return true;
    }
}
