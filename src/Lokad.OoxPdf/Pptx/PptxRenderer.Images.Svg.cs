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
    private static void RenderSvgPicture(PdfGraphicsBuilder graphics, PptxDocument document, ShapeBounds bounds, byte[] bytes, CropRect crop, FillRect fillRect, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex, string? partName, CancellationToken cancellationToken)
    {
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
            return;
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (!TryReadSvgViewBox(svg.Root, out double minX, out double minY, out double viewWidth, out double viewHeight) ||
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
        var gradients = ReadSvgGradients(svg);
        var radialGradients = ReadSvgRadialGradients(svg);
        ReportUnsupportedSvgElements(svg, diagnosticSink, slideIndex, partName);
        var unsupportedCommands = new SortedSet<char>();
        int unreadablePaths = 0;
        var missingGradients = new SortedSet<string>(StringComparer.Ordinal);
        int unpaintablePaths = 0;
        int unsupportedTransforms = 0;
        int gradientStrokes = 0;
        int unpaintableStrokes = 0;
        int invalidStrokePresentations = 0;
        int focalRadialGradients = 0;
        int vectorEffectStrokes = 0;
        foreach (XElement element in svg.Descendants().Where(candidate => IsSvgPaintableElement(candidate)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? data = element.Name.LocalName == "path" ? (string?)element.Attribute("d") : ConvertSvgShapeToPathData(element);
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
            bool hasFill = TryReadSvgFill(element, gradients, radialGradients, out SvgPaint paint, out SvgFillFailure fillFailure, out string? gradientId, out SvgRadialGradient? radial);
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
            SvgStroke stroke = ReadSvgStroke(element, out SvgStrokeFailure strokeFailure, out bool invalidStrokePresentation, out bool hasVectorEffect);
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
            bool evenOddFill = ReadSvgInheritedAttribute(element, ReadSvgStyleDeclarations(element), "fill-rule")?.Trim().Equals("evenodd", StringComparison.OrdinalIgnoreCase) == true;
            double strokeScale = Math.Sqrt(Math.Abs(transform.M11 * transform.M22 - transform.M12 * transform.M21)) * (scaleX + scaleY) / 2d;
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
                dashPhasePoints = stroke.DashOffset * strokeScale;
            }
            if (radial is { } radialGradient)
            {
                if (radialGradient.HasFocal)
                {
                    focalRadialGradients++;
                }
                if (TryReadSvgPathBounds(data, transform, out SvgPathBounds radialBounds))
                {
                    RenderSvgRadialGradientPath(graphics, data, radialGradient, transform, paint.Opacity, radialBounds, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY);
                    if (stroke.Color is { } radialStrokeColor
                        && !TryPaintSvgStrokePath(graphics, data, transform, radialStrokeColor, strokeWidthPoints, stroke.Opacity, stroke.LineCap, stroke.LineJoin, dashPoints, dashPhasePoints, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY)
                        && badCommand is null)
                    {
                        unreadablePaths++;
                    }
                }
                else if (badCommand is null)
                {
                    unreadablePaths++;
                }
            }
            else if (paint.Gradient is { } gradient)
            {
                if (TryReadSvgPathBounds(data, transform, out SvgPathBounds pathBounds))
                {
                    RenderSvgGradientPath(graphics, data, gradient, transform, paint.Opacity, pathBounds, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY);
                    if (stroke.Color is { } gradientStrokeColor
                        && !TryPaintSvgStrokePath(graphics, data, transform, gradientStrokeColor, strokeWidthPoints, stroke.Opacity, stroke.LineCap, stroke.LineJoin, dashPoints, dashPhasePoints, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY)
                        && badCommand is null)
                    {
                        unreadablePaths++;
                    }
                }
                else if (badCommand is null)
                {
                    unreadablePaths++;
                }
            }
            else if (paint.Color is { } color)
            {
                if (stroke.Color is { } strokeColor)
                {
                    bool transparent = paint.Opacity < 1d || stroke.Opacity < 1d;
                    bool styled = stroke.HasStyledPresentation;
                    if (transparent || styled)
                    {
                        graphics.SaveState();
                    }
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
                    if (TryAppendSvgPath(graphics, data, transform, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY))
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
                    if (transparent || styled)
                    {
                        graphics.RestoreState();
                    }
                }
                else
                {
                    if (paint.Opacity < 1d)
                    {
                        graphics.SaveState();
                        graphics.SetAlpha(paint.Opacity, 1d);
                    }

                    graphics.SetFillRgb(color.Red, color.Green, color.Blue);
                    if (TryAppendSvgPath(graphics, data, transform, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY))
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
            else if (stroke.Color is { } strokeOnlyColor)
            {
                if (!TryPaintSvgStrokePath(graphics, data, transform, strokeOnlyColor, strokeWidthPoints, stroke.Opacity, stroke.LineCap, stroke.LineJoin, dashPoints, dashPhasePoints, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY)
                    && badCommand is null)
                {
                    unreadablePaths++;
                }
            }
        }
        ReportSkippedSvgPaths(unsupportedCommands, unreadablePaths, missingGradients, unpaintablePaths, unsupportedTransforms, gradientStrokes, unpaintableStrokes, invalidStrokePresentations, vectorEffectStrokes, focalRadialGradients, diagnosticSink, slideIndex, partName);

        graphics.RestoreState();
    }

    // RV07: unsupported rendered content must diagnose instead of silently
    // vanishing. One diagnostic per element kind (with occurrence count),
    // restricted to paintable content outside definitions (inert definitions
    // are correctly skipped).
    private static void ReportUnsupportedSvgElements(XDocument svg, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex, string? partName)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        var counts = new SortedDictionary<string, int>(StringComparer.Ordinal);
        foreach (XElement element in svg.Descendants())
        {
            string name = element.Name.LocalName;
            if (IsSupportedSvgElement(name) || IsSvgDefinitionElement(element))
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
    private static bool IsSupportedSvgElement(string name)
    {
        return name is "svg" or "defs" or "g" or "title" or "desc" or "metadata" or "style" or "path" or "rect" or "circle" or "ellipse" or "line" or "polyline" or "polygon" or "linearGradient" or "radialGradient" or "stop";
    }
    private static bool IsSvgDefinitionElement(XElement element)
    {
        for (XElement? parent = element.Parent; parent is not null; parent = parent.Parent)
        {
            string name = parent.Name.LocalName;
            if (name is "defs" or "linearGradient" or "radialGradient" or "style")
            {
                return true;
            }
        }
        return false;
    }
    // RV07: first unsupported path command for diagnostics (mirrors the M/L/H/V/C/Z
    // support in both path readers without reinterpreting data).
    private static char? FindFirstUnsupportedSvgPathCommand(string data)
    {
        // The path tokenizer only matches supported command letters, so unknown
        // commands never surface as tokens: scan raw data instead, skipping
        // scientific-notation exponents that merely look like letters.
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
        return "MLHVCZ".IndexOf(char.ToUpperInvariant(command)) >= 0;
    }
    private static void ReportSkippedSvgPaths(SortedSet<char> unsupportedCommands, int unreadablePaths, SortedSet<string> missingGradients, int unpaintablePaths, int unsupportedTransforms, int gradientStrokes, int unpaintableStrokes, int invalidStrokePresentations, int vectorEffectStrokes, int focalRadialGradients, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex, string? partName)
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
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture renders " + invalidStrokePresentations.ToString(CultureInfo.InvariantCulture) + " paths with default stroke effects for unparsable dash/cap/join values.");
        }
        if (focalRadialGradients > 0)
        {
            EmitSvgWarning(diagnosticSink, slideIndex, partName, "SVG picture ignores focal points on " + focalRadialGradients.ToString(CultureInfo.InvariantCulture) + " radial gradients.");
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

    private static void RenderSvgGradientPath(
        PdfGraphicsBuilder graphics,
        string data,
        SvgGradient gradient,
        SvgTransform transform,
        double opacity,
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
        if (!TryAppendSvgPath(graphics, data, transform, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY))
        {
            graphics.RestoreState();
            return;
        }

        graphics.ClipCurrentPath();
        if (opacity < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(opacity, 1d);
        }
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
        if (Math.Abs(effective.X2 - effective.X1) >= Math.Abs(effective.Y2 - effective.Y1))
        {
            int stripCount = Math.Clamp((int)Math.Ceiling(pathWidth / 2d), 16, 128);
            double stripSvgWidth = pathWidth / stripCount;
            double sampleY = pathBounds.CenterY;
            for (int strip = 0; strip < stripCount; strip++)
            {
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
            double stripSvgHeight = pathHeight / stripCount;
            double sampleX = pathBounds.CenterX;
            double stripX = imageX + (pathBounds.MinX - minX) * scaleX;
            double stripWidth = Math.Max(0.001d, pathWidth * scaleX);
            for (int strip = 0; strip < stripCount; strip++)
            {
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
                // RV07: a missing stop-color defaults to black (PowerPoint normalizes black
                // stops by dropping the attribute); present-but-unparseable colors still filter out.
                .Where(stop => stop.Color is not null || !stop.HasColorAttribute)
                .Select(stop => new SvgGradientStop(stop.Offset, stop.Color ?? new RgbColor(0, 0, 0)))
                .OrderBy(stop => stop.Offset)
                .ToArray();
            if (!string.IsNullOrWhiteSpace(id) && stops.Length > 0)
            {
                // RV07: gradientTransform composes onto the raw vector in its own
                // coordinate system; unparseable transforms leave no gradient so
                // referencing paths diagnose as unresolvable instead of
                // misrendering untransformed.
                if (!TryParseSvgTransformList((string?)gradient.Attribute("gradientTransform"), out SvgTransform gradientTransform))
                {
                    continue;
                }

                (double rawX1, double rawY1) = gradientTransform.Apply(
                    ReadSvgDoubleAttribute(gradient, "x1", 0d),
                    ReadSvgDoubleAttribute(gradient, "y1", 0d));
                (double rawX2, double rawY2) = gradientTransform.Apply(
                    ReadSvgDoubleAttribute(gradient, "x2", 1d),
                    ReadSvgDoubleAttribute(gradient, "y2", 0d));
                gradients[id] = new SvgGradient(
                    rawX1,
                    rawY1,
                    rawX2,
                    rawY2,
                    stops,
                    string.Equals((string?)gradient.Attribute("gradientUnits"), "userSpaceOnUse", StringComparison.Ordinal),
                    ReadSvgGradientSpread((string?)gradient.Attribute("spreadMethod")));
            }
        }

        return gradients;
    }

    private static (double Offset, RgbColor? Color, bool HasColorAttribute) ReadSvgGradientStop(XElement stop)
    {
        string? stopColorAttribute = (string?)stop.Attribute("stop-color");
        RgbColor? stopColor = RgbColor.TryParseCssColor(stopColorAttribute, out RgbColor parsedStopColor) ? parsedStopColor : null;
        return (ReadSvgOffset((string?)stop.Attribute("offset")), stopColor, !string.IsNullOrWhiteSpace(stopColorAttribute));
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

    private static double ReadSvgDoubleAttribute(XElement element, string name, double fallback)
    {
        string? value = (string?)element.Attribute(name);
        return string.IsNullOrWhiteSpace(value)
            ? fallback
            : double.Parse(value, CultureInfo.InvariantCulture);
    }

    // RV07: CSS style declarations override presentation attributes; opacity
    // multiplies fill-opacity with whole-element opacity.
    private static IReadOnlyDictionary<string, string> ReadSvgStyleDeclarations(XElement path)
    {
        var declarations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string declaration in (((string?)path.Attribute("style")) ?? string.Empty).Split((char)59))
        {
            int colon = declaration.IndexOf((char)58);
            if (colon <= 0)
            {
                continue;
            }

            string name = declaration.Substring(0, colon).Trim();
            string value = declaration.Substring(colon + 1).Trim();
            if (name.Length != 0 && value.Length != 0)
            {
                declarations[name] = value;
            }
        }

        return declarations;
    }

    private static double ReadSvgOpacityValue(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return 1d;
        }

        string trimmed = text.Trim();
        if (trimmed.EndsWith("%", StringComparison.Ordinal) && double.TryParse(trimmed.Substring(0, trimmed.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent))
        {
            return Math.Clamp(percent / 100d, 0d, 1d);
        }

        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
        {
            return Math.Clamp(value, 0d, 1d);
        }

        return 1d;
    }

    private static bool TryReadSvgFill(XElement path, IReadOnlyDictionary<string, SvgGradient> gradients, IReadOnlyDictionary<string, SvgRadialGradient> radialGradients, out SvgPaint paint, out SvgFillFailure failure, out string? gradientId, out SvgRadialGradient? radial)
    {
        gradientId = null;
        radial = null;
        IReadOnlyDictionary<string, string> style = ReadSvgStyleDeclarations(path);
        string? fill = ReadSvgInheritedAttribute(path, style, "fill");
        double opacity = ReadSvgOpacityValue(style.TryGetValue("fill-opacity", out string? styleFillOpacity) ? styleFillOpacity : (string?)path.Attribute("fill-opacity"))
            * ReadSvgOpacityValue(style.TryGetValue("opacity", out string? styleOpacity) ? styleOpacity : (string?)path.Attribute("opacity"));
        if (fill is null)
        {
            // SVG initial value: a path without fill paints black instead of vanishing.
            paint = new SvgPaint(new RgbColor(0, 0, 0), null, opacity);
            failure = SvgFillFailure.None;
            return true;
        }
        if (fill.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            paint = default;
            failure = SvgFillFailure.None;
            return false;
        }
        Match gradient = Regex.Match(fill, @"url\(#(?<id>[^)]+)\)");
        if (gradient.Success)
        {
            gradientId = gradient.Groups["id"].Value;
            if (gradients.TryGetValue(gradientId, out SvgGradient? svgGradient))
            {
                paint = new SvgPaint(null, svgGradient, opacity);
                failure = SvgFillFailure.None;
                return true;
            }
            if (radialGradients.TryGetValue(gradientId, out SvgRadialGradient? svgRadial))
            {
                radial = svgRadial;
                paint = new SvgPaint(null, null, opacity);
                failure = SvgFillFailure.None;
                return true;
            }
            paint = default;
            failure = SvgFillFailure.UnresolvedGradient;
            return false;
        }
        if (RgbColor.TryParseCssColor(fill, out RgbColor color))
        {
            paint = new SvgPaint(color, null, opacity);
            failure = SvgFillFailure.None;
            return true;
        }
        paint = default;
        failure = SvgFillFailure.UnparsableColor;
        return false;
    }

    private readonly record struct SvgStroke(RgbColor? Color, double Width, double Opacity, int LineCap, int LineJoin, double[]? DashPattern, double DashOffset)
    {
        public bool HasPaint => Color is not null;
        public bool HasStyledPresentation => DashPattern is not null || LineCap != 0 || LineJoin != 0;
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
    // Stroke widths are user units scaled by the path transform area scale
    // and the viewBox mapping; non-uniform mappings stay approximate.
    private static SvgStroke ReadSvgStroke(XElement path, out SvgStrokeFailure failure, out bool invalidPresentation, out bool hasVectorEffect)
    {
        failure = SvgStrokeFailure.None;
        IReadOnlyDictionary<string, string> style = ReadSvgStyleDeclarations(path);
        string? strokePaint = ReadSvgInheritedAttribute(path, style, "stroke");
        invalidPresentation = false;
        hasVectorEffect = !string.IsNullOrWhiteSpace(ReadSvgPresentationAttribute(path, style, "vector-effect"));
        if (string.IsNullOrWhiteSpace(strokePaint) || strokePaint.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return default;
        }
        double opacity = ReadSvgOpacityValue(ReadSvgPresentationAttribute(path, style, "stroke-opacity"))
            * ReadSvgOpacityValue(ReadSvgPresentationAttribute(path, style, "opacity"));
        Match gradient = Regex.Match(strokePaint, @"url\(#(?<id>[^)]+)\)");
        if (gradient.Success)
        {
            failure = SvgStrokeFailure.UnresolvedGradient;
            return default;
        }
        if (!RgbColor.TryParseCssColor(strokePaint, out RgbColor color))
        {
            failure = SvgStrokeFailure.UnparsableColor;
            return default;
        }
        if (!TryReadSvgStrokeWidth(style, path, out double width))
        {
            failure = SvgStrokeFailure.UnparsableWidth;
            return default;
        }
        int lineCap = ReadSvgLineCap(ReadSvgPresentationAttribute(path, style, "stroke-linecap"), out bool capInvalid);
        int lineJoin = ReadSvgLineJoin(ReadSvgPresentationAttribute(path, style, "stroke-linejoin"), out bool joinInvalid);
        double[]? dash = ReadSvgDashPattern(ReadSvgPresentationAttribute(path, style, "stroke-dasharray"), out bool dashInvalid);
        double offset = 0d;
        if (dash is not null && !TryReadSvgStrokeOffset(ReadSvgPresentationAttribute(path, style, "stroke-dashoffset"), out offset))
        {
            dash = null;
            dashInvalid = true;
        }
        invalidPresentation = capInvalid || joinInvalid || dashInvalid;
        return new SvgStroke(color, width, opacity, lineCap, lineJoin, dash, offset);
    }
    private static bool TryReadSvgStrokeWidth(IReadOnlyDictionary<string, string> style, XElement path, out double width)
    {
        width = 1d;
        string? text = ReadSvgPresentationAttribute(path, style, "stroke-width");
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }
        string trimmed = text.Trim();
        if (trimmed.EndsWith("px", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed.Substring(0, trimmed.Length - 2);
        }
        return double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out width) && width >= 0d;
    }
    // RV07: fill and stroke inherit through groups in cascade order (element
    // style, element attribute, then each ancestor outward the same way).
    // Widths, opacities and presentations stay element-local per SVG.
    private static string? ReadSvgInheritedAttribute(XElement path, IReadOnlyDictionary<string, string> style, string name)
    {
        if (style.TryGetValue(name, out string? styleValue) && !string.IsNullOrWhiteSpace(styleValue))
        {
            return styleValue;
        }
        string? attribute = (string?)path.Attribute(name);
        if (!string.IsNullOrWhiteSpace(attribute))
        {
            return attribute;
        }
        foreach (XElement ancestor in path.Ancestors())
        {
            IReadOnlyDictionary<string, string> ancestorStyle = ReadSvgStyleDeclarations(ancestor);
            if (ancestorStyle.TryGetValue(name, out string? ancestorStyleValue) && !string.IsNullOrWhiteSpace(ancestorStyleValue))
            {
                return ancestorStyleValue;
            }
            string? ancestorAttribute = (string?)ancestor.Attribute(name);
            if (!string.IsNullOrWhiteSpace(ancestorAttribute))
            {
                return ancestorAttribute;
            }
        }
        return null;
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
    // RV07: dash patterns scale like widths at emission; an odd count doubles
    // per SVG and an all-zero pattern means solid. Percentages stay diagnosed.
    private static double[]? ReadSvgDashPattern(string? text, out bool invalid)
    {
        invalid = false;
        if (string.IsNullOrWhiteSpace(text) || text.Trim().Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }
        string[] parts = text.Replace(",", " ").Split((char)32, StringSplitOptions.RemoveEmptyEntries);
        var lengths = new List<double>();
        foreach (string part in parts)
        {
            string item = part.Trim();
            if (item.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            {
                item = item.Substring(0, item.Length - 2);
            }
            if (!double.TryParse(item, NumberStyles.Float, CultureInfo.InvariantCulture, out double length) || length < 0d)
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
    private static bool TryReadSvgStrokeOffset(string? text, out double offset)
    {
        offset = 0d;
        if (string.IsNullOrWhiteSpace(text))
        {
            return true;
        }
        string trimmed = text.Trim();
        if (trimmed.EndsWith("px", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed.Substring(0, trimmed.Length - 2);
        }
        return double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out offset);
    }
    private static bool TryPaintSvgStrokePath(PdfGraphicsBuilder graphics, string data, SvgTransform transform, RgbColor color, double widthPoints, double opacity, int lineCap, int lineJoin, double[]? dashPoints, double dashPhasePoints, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY)
    {
        bool styled = dashPoints is not null || lineCap != 0 || lineJoin != 0;
        if (opacity < 1d || styled)
        {
            graphics.SaveState();
        }
        if (opacity < 1d)
        {
            graphics.SetAlpha(1d, opacity);
        }
        graphics.SetStrokeRgb(color.Red, color.Green, color.Blue);
        graphics.SetLineWidth(widthPoints);
        if (lineCap != 0)
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
        bool painted = TryAppendSvgPath(graphics, data, transform, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY);
        if (painted)
        {
            graphics.StrokeCurrentPath();
        }
        if (opacity < 1d || styled)
        {
            graphics.RestoreState();
        }
        return painted;
    }
    // RV07: radial gradients paint concentric ellipse rings, largest first.
    // Focal points diagnose and render centered; radii resolve per axis like
    // the linear vector mapping, with the gradient transform baked at read.
    private static IReadOnlyDictionary<string, SvgRadialGradient> ReadSvgRadialGradients(XDocument svg)
    {
        var gradients = new Dictionary<string, SvgRadialGradient>(StringComparer.Ordinal);
        foreach (XElement gradient in svg.Descendants().Where(element => element.Name.LocalName == "radialGradient"))
        {
            string? id = (string?)gradient.Attribute("id");
            SvgGradientStop[] stops = gradient
                .Elements()
                .Where(element => element.Name.LocalName == "stop")
                .Select(ReadSvgGradientStop)
                .Where(stop => stop.Color is not null || !stop.HasColorAttribute)
                .Select(stop => new SvgGradientStop(stop.Offset, stop.Color ?? new RgbColor(0, 0, 0)))
                .OrderBy(stop => stop.Offset)
                .ToArray();
            if (string.IsNullOrWhiteSpace(id) || stops.Length == 0)
            {
                continue;
            }
            if (!TryParseSvgTransformList((string?)gradient.Attribute("gradientTransform"), out SvgTransform gradientTransform))
            {
                continue;
            }
            bool userSpace = string.Equals((string?)gradient.Attribute("gradientUnits"), "userSpaceOnUse", StringComparison.Ordinal);
            double cx = ReadSvgDoubleAttribute(gradient, "cx", 0.5d);
            double cy = ReadSvgDoubleAttribute(gradient, "cy", 0.5d);
            double radius = ReadSvgDoubleAttribute(gradient, "r", 0.5d);
            if (radius <= 0d)
            {
                continue;
            }
            double fx = ReadSvgDoubleAttribute(gradient, "fx", cx);
            double fy = ReadSvgDoubleAttribute(gradient, "fy", cy);
            (double rawCx, double rawCy) = gradientTransform.Apply(cx, cy);
            double rawRadius = radius;
            if (userSpace)
            {
                (double edgeX, double edgeY) = gradientTransform.Apply(cx + radius, cy + radius);
                rawRadius = Math.Max(Math.Abs(edgeX - rawCx), Math.Abs(edgeY - rawCy));
            }
            gradients[id] = new SvgRadialGradient(rawCx, rawCy, rawRadius, fx != cx || fy != cy, stops, userSpace, ReadSvgGradientSpread((string?)gradient.Attribute("spreadMethod")));
        }
        return gradients;
    }
    private static void RenderSvgRadialGradientPath(PdfGraphicsBuilder graphics, string data, SvgRadialGradient radial, SvgTransform transform, double opacity, SvgPathBounds pathBounds, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY)
    {
        double pathWidth = Math.Max(0.001d, pathBounds.MaxX - pathBounds.MinX);
        double pathHeight = Math.Max(0.001d, pathBounds.MaxY - pathBounds.MinY);
        double centerX;
        double centerY;
        double radiusX;
        double radiusY;
        if (radial.IsUserSpace)
        {
            (centerX, centerY) = transform.Apply(radial.Cx, radial.Cy);
            radiusX = radial.Radius;
            radiusY = radial.Radius;
        }
        else
        {
            centerX = pathBounds.MinX + radial.Cx * pathWidth;
            centerY = pathBounds.MinY + radial.Cy * pathHeight;
            radiusX = radial.Radius * pathWidth;
            radiusY = radial.Radius * pathHeight;
        }
        graphics.SaveState();
        if (!TryAppendSvgPath(graphics, data, transform, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY))
        {
            graphics.RestoreState();
            return;
        }
        graphics.ClipCurrentPath();
        if (opacity < 1d)
        {
            graphics.SaveState();
            graphics.SetAlpha(opacity, 1d);
        }
        int ringCount = Math.Clamp((int)Math.Ceiling(Math.Max(radiusX, radiusY) / 2d), 16, 128);
        for (int ring = 0; ring < ringCount; ring++)
        {
            double fraction = (double)(ringCount - ring) / ringCount;
            RgbColor color = SampleSvgRadialStops(radial.Stops, (double)(ringCount - ring - 1) / (ringCount - 1), radial.Spread);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            string ringData = ConvertSvgEllipseBody(centerX, centerY, radiusX * fraction, radiusY * fraction);
            if (TryAppendSvgPath(graphics, ringData, SvgTransform.Identity, minX, minY, imageX, imageY, imageHeight, scaleX, scaleY))
            {
                graphics.FillCurrentPath();
            }
        }
        if (opacity < 1d)
        {
            graphics.RestoreState();
        }
        graphics.RestoreState();
    }
    private static RgbColor SampleSvgRadialStops(IReadOnlyList<SvgGradientStop> stops, double t, SvgGradientSpread spread)
    {
        double offset = spread switch
        {
            SvgGradientSpread.Repeat => t - Math.Floor(t),
            SvgGradientSpread.Reflect => ReflectSvgRadialOffset(t),
            _ => Math.Clamp(t, 0d, 1d),
        };
        SvgGradientStop previous = stops[0];
        foreach (SvgGradientStop next in stops.Skip(1))
        {
            if (offset <= next.Offset)
            {
                double span = next.Offset - previous.Offset;
                double amount = span <= PptxTextMetricRules.TextStateTolerance ? 0d : (offset - previous.Offset) / span;
                return InterpolateSvgGradientStop(previous.Color, next.Color, Math.Clamp(amount, 0d, 1d));
            }
            previous = next;
        }
        return previous.Color;
    }
    private static double ReflectSvgRadialOffset(double value)
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
        offset = gradient.Spread switch
        {
            SvgGradientSpread.Repeat => offset - Math.Floor(offset),
            SvgGradientSpread.Reflect => ReflectSvgGradientOffset(offset),
            _ => Math.Clamp(offset, 0d, 1d),
        };

        static double ReflectSvgGradientOffset(double value)
        {
            double wrapped = value % 2d;
            if (wrapped < 0d)
            {
                wrapped += 2d;
            }

            return wrapped > 1d ? 2d - wrapped : wrapped;
        }

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

    // RV18: one shared path representation for bounds and emission. The parser
    // resolves relative coordinates once, so both consumers walk the same
    // absolute commands instead of tokenizing the path data separately.
    private readonly record struct SvgPathCommand(char Kind, double[] Arguments);

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
        double currentX = 0d;
        double currentY = 0d;
        double startX = 0d;
        double startY = 0d;
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
            switch (char.ToUpperInvariant(command))
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
                    if (!TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out double c1x, out double c1y) ||
                        !TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out double c2x, out double c2y) ||
                        !TryReadSvgPoint(tokens, ref index, relative, currentX, currentY, out currentX, out currentY))
                    {
                        return (commands, false);
                    }

                    commands.Add(new SvgPathCommand('C', new[] { c1x, c1y, c2x, c2y, currentX, currentY }));
                    break;
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
        }

        return (commands, true);
    }

    private static bool TryReadSvgPathBounds(string data, SvgTransform transform, out SvgPathBounds bounds)
    {
        (List<SvgPathCommand> commands, bool complete) = ParseSvgPathData(data);
        commands = TransformSvgCommands(commands, transform);
        bounds = default;
        if (!complete || commands.Count == 0)
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
                    Include(pathCommand.Arguments[0], pathCommand.Arguments[1]);
                    Include(pathCommand.Arguments[2], pathCommand.Arguments[3]);
                    currentX = pathCommand.Arguments[4];
                    currentY = pathCommand.Arguments[5];
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

    private static bool TryAppendSvgPath(PdfGraphicsBuilder graphics, string data, SvgTransform transform, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY)
    {
        List<SvgPathCommand> commands = TransformSvgCommands(ParseSvgPathData(data).Commands, transform);
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
                    graphics.MoveTo(SvgX(currentX), SvgY(currentY));
                    hasPath = true;
                    break;
                case 'L':
                    currentX = pathCommand.Arguments[0];
                    currentY = pathCommand.Arguments[1];
                    graphics.LineTo(SvgX(currentX), SvgY(currentY));
                    break;
                case 'H':
                    currentX = pathCommand.Arguments[0];
                    graphics.LineTo(SvgX(currentX), SvgY(currentY));
                    break;
                case 'V':
                    currentY = pathCommand.Arguments[0];
                    graphics.LineTo(SvgX(currentX), SvgY(currentY));
                    break;
                case 'C':
                    graphics.CurveTo(SvgX(pathCommand.Arguments[0]), SvgY(pathCommand.Arguments[1]), SvgX(pathCommand.Arguments[2]), SvgY(pathCommand.Arguments[3]), SvgX(pathCommand.Arguments[4]), SvgY(pathCommand.Arguments[5]));
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

        double SvgX(double value) => imageX + (value - minX) * scaleX;
        double SvgY(double value) => imageY + imageHeight - (value - minY) * scaleY;
    }

    // RV07: basic shapes convert to path data and paint through the shared
    // path pipeline. Circles, ellipses and rounded corners use the kappa
    // bezier approximation as a platform-independent literal; spec-invalid
    // geometry returns null and renders nothing like Office. Text stays out
    // of scope behind the unsupported-element diagnostic.
    private const double SvgCircleKappa = 0.5522847498;
    private static bool IsSvgPaintableElement(XElement element)
    {
        return element.Name.LocalName is "path" or "rect" or "circle" or "ellipse" or "line" or "polyline" or "polygon";
    }
    private static string? ConvertSvgShapeToPathData(XElement element)
    {
        return element.Name.LocalName switch
        {
            "rect" => ConvertSvgRectToPathData(element),
            "circle" => ConvertSvgCircleToPathData(element),
            "ellipse" => ConvertSvgEllipseToPathData(element),
            "line" => ConvertSvgLineToPathData(element),
            "polyline" => ConvertSvgPointsToPathData(element, false),
            "polygon" => ConvertSvgPointsToPathData(element, true),
            _ => (string?)element.Attribute("d"),
        };
    }
    private static bool TryReadSvgShapeCoordinate(XElement element, string name, double fallback, out double value)
    {
        string? text = (string?)element.Attribute(name);
        if (string.IsNullOrWhiteSpace(text))
        {
            value = fallback;
            return true;
        }
        return double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
    }
    private static string SvgPathNumber(double value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
    private static string? ConvertSvgRectToPathData(XElement element)
    {
        if (!TryReadSvgShapeCoordinate(element, "x", 0d, out double x)
            || !TryReadSvgShapeCoordinate(element, "y", 0d, out double y)
            || !TryReadSvgShapeCoordinate(element, "width", 0d, out double width)
            || !TryReadSvgShapeCoordinate(element, "height", 0d, out double height))
        {
            return null;
        }
        if (width <= 0d || height <= 0d)
        {
            return null;
        }
        string? rxText = (string?)element.Attribute("rx");
        string? ryText = (string?)element.Attribute("ry");
        double rx = 0d;
        double ry = 0d;
        if (rxText is not null && !double.TryParse(rxText.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out rx))
        {
            return null;
        }
        if (ryText is not null && !double.TryParse(ryText.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out ry))
        {
            return null;
        }
        if (rxText is null && ryText is not null)
        {
            rx = ry;
        }
        if (ryText is null && rxText is not null)
        {
            ry = rx;
        }
        if (rx < 0d || ry < 0d)
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
    private static string? ConvertSvgCircleToPathData(XElement element)
    {
        if (!TryReadSvgShapeCoordinate(element, "cx", 0d, out double cx)
            || !TryReadSvgShapeCoordinate(element, "cy", 0d, out double cy)
            || !TryReadSvgShapeCoordinate(element, "r", 0d, out double radius)
            || radius <= 0d)
        {
            return null;
        }
        return ConvertSvgEllipseBody(cx, cy, radius, radius);
    }
    private static string? ConvertSvgEllipseToPathData(XElement element)
    {
        if (!TryReadSvgShapeCoordinate(element, "cx", 0d, out double cx)
            || !TryReadSvgShapeCoordinate(element, "cy", 0d, out double cy)
            || !TryReadSvgShapeCoordinate(element, "rx", 0d, out double rx)
            || !TryReadSvgShapeCoordinate(element, "ry", 0d, out double ry)
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
    private static string? ConvertSvgLineToPathData(XElement element)
    {
        if (!TryReadSvgShapeCoordinate(element, "x1", 0d, out double x1)
            || !TryReadSvgShapeCoordinate(element, "y1", 0d, out double y1)
            || !TryReadSvgShapeCoordinate(element, "x2", 0d, out double x2)
            || !TryReadSvgShapeCoordinate(element, "y2", 0d, out double y2))
        {
            return null;
        }
        return "M" + SvgPathNumber(x1) + " " + SvgPathNumber(y1) + "L" + SvgPathNumber(x2) + " " + SvgPathNumber(y2);
    }
    private static string? ConvertSvgPointsToPathData(XElement element, bool closed)
    {
        System.Text.RegularExpressions.MatchCollection numbers = SvgNumberRegex().Matches((string?)element.Attribute("points") ?? string.Empty);
        if (numbers.Count < 4 || numbers.Count % 2 == 1)
        {
            return null;
        }
        var points = new List<double>();
        foreach (System.Text.RegularExpressions.Match number in numbers)
        {
            if (!double.TryParse(number.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double coordinate))
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

        value = double.Parse(tokens[index].Value, CultureInfo.InvariantCulture);
        index++;
        return true;
    }

    [GeneratedRegex(@"[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?", RegexOptions.CultureInvariant)]
    private static partial Regex SvgNumberRegex();

    [GeneratedRegex(@"[MmLlHhVvCcZz]|[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?", RegexOptions.CultureInvariant)]
    private static partial Regex SvgPathTokenRegex();

    [GeneratedRegex(@"[A-Za-z]+|[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?|[(),]", RegexOptions.CultureInvariant)]
    private static partial Regex SvgTransformTokenRegex();

    // RV07: affine SVG transform lists shared by bounds and emission. translate,
    // scale, rotate, skew, and matrix compose left to right; anything else fails
    // so the path diagnoses instead of drawing untransformed.
    private readonly record struct SvgTransform(double M11, double M12, double M21, double M22, double OffsetX, double OffsetY)
    {
        public static SvgTransform Identity => new SvgTransform(1d, 0d, 0d, 1d, 0d, 0d);

        public bool IsIdentity => M11 == 1d && M12 == 0d && M21 == 0d && M22 == 1d && OffsetX == 0d && OffsetY == 0d;

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

            accumulated = SvgTransform.Compose(next, accumulated);
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
        double currentX = 0d;
        double currentY = 0d;
        double startX = 0d;
        double startY = 0d;
        double sourceCurrentX = 0d;
        double sourceCurrentY = 0d;
        double sourceStartX = 0d;
        double sourceStartY = 0d;
        foreach (SvgPathCommand command in commands)
        {
            string kind = command.Kind.ToString();
            if (kind == "M")
            {
                sourceCurrentX = command.Arguments[0];
                sourceCurrentY = command.Arguments[1];
                (currentX, currentY) = transform.Apply(sourceCurrentX, sourceCurrentY);
                startX = currentX;
                startY = currentY;
                sourceStartX = sourceCurrentX;
                sourceStartY = sourceCurrentY;
                mapped.Add(new SvgPathCommand(
                    command.Kind, new double[] { currentX, currentY }));
            }
            else if (kind == "L")
            {
                sourceCurrentX = command.Arguments[0];
                sourceCurrentY = command.Arguments[1];
                (currentX, currentY) = transform.Apply(sourceCurrentX, sourceCurrentY);
                mapped.Add(new SvgPathCommand(
                    command.Kind, new double[] { currentX, currentY }));
            }
            else if (kind == "H")
            {
                sourceCurrentX = command.Arguments[0];
                (currentX, currentY) = transform.Apply(sourceCurrentX, sourceCurrentY);
                mapped.Add(new SvgPathCommand(
                    "L"[0], new double[] { currentX, currentY }));
            }
            else if (kind == "V")
            {
                sourceCurrentY = command.Arguments[0];
                (currentX, currentY) = transform.Apply(sourceCurrentX, sourceCurrentY);
                mapped.Add(new SvgPathCommand(
                    "L"[0], new double[] { currentX, currentY }));
            }
            else if (kind == "C")
            {
                (double X0, double Y0) = transform.Apply(command.Arguments[0], command.Arguments[1]);
                (double X1, double Y1) = transform.Apply(command.Arguments[2], command.Arguments[3]);
                sourceCurrentX = command.Arguments[4];
                sourceCurrentY = command.Arguments[5];
                (currentX, currentY) = transform.Apply(sourceCurrentX, sourceCurrentY);
                mapped.Add(new SvgPathCommand(
                    command.Kind, new double[] { X0, Y0, X1, Y1, currentX, currentY }));
            }
            else if (kind == "Z")
            {
                currentX = startX;
                currentY = startY;
                sourceCurrentX = sourceStartX;
                sourceCurrentY = sourceStartY;
                mapped.Add(command);
            }
            else
            {
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

            transform = SvgTransform.Compose(ancestorTransform, transform);
        }

        if (!TryParseSvgTransformList((string?)path.Attribute("transform"), out SvgTransform local))
        {
            transform = SvgTransform.Identity;
            return false;
        }

        transform = SvgTransform.Compose(local, transform);
        return true;
    }
}
