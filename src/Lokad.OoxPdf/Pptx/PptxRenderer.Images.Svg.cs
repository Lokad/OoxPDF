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
        ReportUnsupportedSvgElements(svg, diagnosticSink, slideIndex, partName);
        var unsupportedCommands = new SortedSet<char>();
        int unreadablePaths = 0;
        var missingGradients = new SortedSet<string>(StringComparer.Ordinal);
        int unpaintablePaths = 0;
        foreach (XElement path in svg.Descendants().Where(element => element.Name.LocalName == "path"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? data = (string?)path.Attribute("d");
            if (string.IsNullOrWhiteSpace(data))
            {
                continue;
            }
            char? badCommand = FindFirstUnsupportedSvgPathCommand(data);
            if (badCommand is not null)
            {
                unsupportedCommands.Add(badCommand.Value);
            }
            if (!TryReadSvgFill(path, gradients, out SvgPaint paint, out SvgFillFailure fillFailure, out string? gradientId))
            {
                if (fillFailure == SvgFillFailure.UnresolvedGradient && gradientId is not null)
                {
                    missingGradients.Add(gradientId);
                }
                else if (fillFailure == SvgFillFailure.UnparsableColor)
                {
                    unpaintablePaths++;
                }
                continue;
            }
            if (paint.Gradient is { } gradient)
            {
                if (TryReadSvgPathBounds(data, out SvgPathBounds pathBounds))
                {
                    RenderSvgGradientPath(graphics, data, gradient, pathBounds, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY);
                }
                else if (badCommand is null)
                {
                    unreadablePaths++;
                }
            }
            else if (paint.Color is { } color)
            {
                graphics.SetFillRgb(color.Red, color.Green, color.Blue);
                if (TryAppendSvgPath(graphics, data, sourceMinX, sourceMinY, imageX, imageY, imageHeight, scaleX, scaleY))
                {
                    graphics.FillCurrentPath();
                }
                else if (badCommand is null)
                {
                    unreadablePaths++;
                }
            }
        }
        ReportSkippedSvgPaths(unsupportedCommands, unreadablePaths, missingGradients, unpaintablePaths, diagnosticSink, slideIndex, partName);

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
        return name is "svg" or "defs" or "g" or "title" or "desc" or "metadata" or "style" or "path" or "linearGradient" or "radialGradient" or "stop";
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
    private static void ReportSkippedSvgPaths(SortedSet<char> unsupportedCommands, int unreadablePaths, SortedSet<string> missingGradients, int unpaintablePaths, Action<OoxPdfDiagnostic>? diagnosticSink, int slideIndex, string? partName)
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
        // RV07: objectBoundingBox vectors normalize into path space while
        // userSpaceOnUse vectors stay in user units; strips run along the
        // dominant gradient axis so vertical gradients vary top to bottom.
        double pathWidth = Math.Max(0.001d, pathBounds.MaxX - pathBounds.MinX);
        double pathHeight = Math.Max(0.001d, pathBounds.MaxY - pathBounds.MinY);
        SvgGradient effective = gradient.IsUserSpace
            ? gradient
            : gradient with
            {
                X1 = pathBounds.MinX + gradient.X1 * pathWidth,
                Y1 = pathBounds.MinY + gradient.Y1 * pathHeight,
                X2 = pathBounds.MinX + gradient.X2 * pathWidth,
                Y2 = pathBounds.MinY + gradient.Y2 * pathHeight,
            };
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
                .Select(stop => new SvgGradientStop(stop.Offset, stop.Color ?? default))
                .OrderBy(stop => stop.Offset)
                .ToArray();
            if (!string.IsNullOrWhiteSpace(id) && stops.Length > 0)
            {
                gradients[id] = new SvgGradient(
                    ReadSvgDoubleAttribute(gradient, "x1", 0d),
                    ReadSvgDoubleAttribute(gradient, "y1", 0d),
                    ReadSvgDoubleAttribute(gradient, "x2", 1d),
                    ReadSvgDoubleAttribute(gradient, "y2", 0d),
                    stops,
                    string.Equals((string?)gradient.Attribute("gradientUnits"), "userSpaceOnUse", StringComparison.Ordinal));
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

    private static bool TryReadSvgFill(XElement path, IReadOnlyDictionary<string, SvgGradient> gradients, out SvgPaint paint, out SvgFillFailure failure, out string? gradientId)
    {
        gradientId = null;
        string? fill = (string?)path.Attribute("fill");
        if (fill is null || fill.Equals("none", StringComparison.OrdinalIgnoreCase))
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
                paint = new SvgPaint(null, svgGradient);
                failure = SvgFillFailure.None;
                return true;
            }
            paint = default;
            failure = SvgFillFailure.UnresolvedGradient;
            return false;
        }
        if (RgbColor.TryParse(fill.TrimStart((char)35), out RgbColor color))
        {
            paint = new SvgPaint(color, null);
            failure = SvgFillFailure.None;
            return true;
        }
        paint = default;
        failure = SvgFillFailure.UnparsableColor;
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

    private static bool TryReadSvgPathBounds(string data, out SvgPathBounds bounds)
    {
        (List<SvgPathCommand> commands, bool complete) = ParseSvgPathData(data);
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

    private static bool TryAppendSvgPath(PdfGraphicsBuilder graphics, string data, double minX, double minY, double imageX, double imageY, double imageHeight, double scaleX, double scaleY)
    {
        List<SvgPathCommand> commands = ParseSvgPathData(data).Commands;
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
}
