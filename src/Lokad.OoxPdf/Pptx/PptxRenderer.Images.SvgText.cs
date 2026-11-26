using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private const double SvgInitialFontSize = 16d;

    // Borrow the conversion's prepared font programs. SVG labels are picture
    // outlines, so even isolated alpha groups remain vector-only and need no
    // independent font subsets, PDF resources or font cache.
    internal sealed record SvgTextFont(OpenTypeFont Font, FontFaceResolution Resolution);

    private static IEnumerable<byte[]> EnumerateSvgPictureSources(PptxSceneSlide slide)
    {
        foreach (var source in EnumerateScenePaintSources(slide))
        foreach (var node in EnumerateRenderedSceneNodes(source.Nodes, source.RenderPlaceholders))
            if (node.Picture is { Alpha: > 0d, Resource: { } image } &&
                image.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
                yield return image.Bytes;
    }

    internal static async ValueTask<IReadOnlyDictionary<FontRequest, SvgTextFont?>> PrepareSvgTextFontsAsync(
        IEnumerable<byte[]> pictures, IFontResolver resolver,
        Func<FontFaceResolution, CancellationToken, ValueTask<OpenTypeFont?>> loadFont,
        Action<OoxPdfDiagnostic>? diagnostics, CancellationToken cancellationToken)
    {
        var fonts = new Dictionary<FontRequest, SvgTextFont?>(FontRequestKeyComparer.OrdinalIgnoreCaseFamily);
        var seen = new HashSet<byte[]>(ReferenceEqualityComparer.Instance);
        foreach (byte[] bytes in pictures)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!seen.Add(bytes)) continue;
            using IMemoryReservation? scratch = OoxMemory.Reserve(OoxConversionBudget.Current?.Memory,
                checked(64L * 1024 + 32L * bytes.Length));
            XDocument svg;
            try { using var stream = new MemoryStream(bytes); svg = SafeXml.Load(stream, cancellationToken); }
            catch (InvalidDataException) { continue; } // The paint path owns the existing parse diagnostic/fallback.
            XElement? root = ReadSvgRoot(svg);
            if (root is null || !TryReadSvgSourceFrame(root, out _, out _, out double viewWidth, out double viewHeight)) continue;
            var content = EnumerateSvgContent(root, cancellationToken);
            if (HasSvgStylesheet(content)) continue;
            var elements = content.OfType<XElement>();
            var paintServers = new SvgPaintServers(elements, viewWidth, viewHeight, cancellationToken);
            foreach (XElement element in elements.Where(e => e.Name.LocalName == "text"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (IsSvgDefinitionElement(element) || IsInFullyTransparentSvgContainer(element) ||
                    IsSvgDisplaySuppressed(element, ReadSvgStyleDeclarations(element)) ||
                    !TryReadSvgText(element, paintServers, viewWidth, viewHeight, cancellationToken, out SvgTextLabel? label) || label.Color is null ||
                    label.Opacity == 0d || fonts.ContainsKey(label.Font)) continue;
                FontFaceResolution resolution = resolver.Resolve(label.Font);
                OpenTypeFont? font = await loadFont(resolution, cancellationToken).ConfigureAwait(false);
                fonts.Add(label.Font, font is { HasTrueTypeOutlines: true } ? new SvgTextFont(font, resolution) : null);
                if (font is not null && (resolution.IsFallback || label.Font.Bold && !resolution.Bold || label.Font.Italic && !resolution.Italic))
                    diagnostics?.Invoke(new OoxPdfDiagnostic("SVG_UNSUPPORTED_CONTENT", OoxPdfSeverity.Warning,
                        $"SVG text font/style '{label.Font.FamilyName}' uses the available '{resolution.ResolvedFamily}' face.",
                        Feature: "svg", Fallback: "SubstitutedFont"));
            }
        }
        return fonts;
    }

    private sealed record SvgTextLabel(string Text, bool[]? HiddenCharacters, FontRequest Font, double FontSize, double X, double Y,
        string Anchor, RgbColor? Color, double Opacity, bool Kerning);

    private readonly record struct SvgTextPaint(RgbColor? Color, double Alpha);

    private readonly record struct SvgTextPosition(double X, double Y, double Dx, double Dy);

    private readonly record struct SvgTextMetrics(string FamilyList, string Slant, int Weight, double Size, bool Kerning)
    {
        public bool Matches(SvgTextMetrics other) =>
            Size == other.Size && Slant == other.Slant && Weight == other.Weight && Kerning == other.Kerning &&
            FamilyList.Equals(other.FamilyList, StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasSvgStylesheet(IEnumerable<XNode> content) =>
        content.Any(node => node switch
        {
            XProcessingInstruction instruction => instruction.Target == "xml-stylesheet",
            XElement element => element.Name.LocalName == "style" &&
                element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)),
            _ => false
        });

    private static bool IsSvgMetadataElement(XElement element) =>
        element.Name.LocalName is "title" or "desc" or "metadata";

    private static bool TryReadSvgText(XElement element, SvgPaintServers paintServers,
        double viewWidth, double viewHeight, CancellationToken cancellationToken, [NotNullWhen(true)] out SvgTextLabel? label)
    {
        label = null;
        // lengthAdjust selects the adjustment method for an authored textLength.
        if (element.Attribute("textLength") is not null ||
            (element.Attribute("rotate") is { } rotation && !IsZeroSvgTextPositionList(rotation.Value, allowLengths: false, cancellationToken))) return false;
        var style = ReadSvgStyleDeclarations(element);
        string? Inherited(string name) => ReadSvgInheritedAttribute(element, style, name)?.Trim();
        bool PreservesTextStyle(string name) => IsNeutralSvgTextStyle(name, Inherited(name));
        if (!PreservesTextStyle("direction") || !PreservesTextStyle("unicode-bidi") ||
            !PreservesTextStyle("writing-mode") || !PreservesTextStyle("dominant-baseline") ||
            !PreservesTextStyle("alignment-baseline") || !PreservesTextStyle("baseline-shift") ||
            !PreservesTextStyle("font-variant") || !PreservesTextStyle("font-stretch") ||
            !PreservesTextStyle("letter-spacing") || !PreservesTextStyle("word-spacing") ||
            !PreservesTextStyle("text-decoration") ||
            !PreservesTextStyle("clip-path") || !PreservesTextStyle("mask") || !PreservesTextStyle("filter")) return false;
        string? initialAnchor = ResolveSvgTextAnchor(Inherited("text-anchor"), "start");
        if (initialAnchor is null) return false;
        double size = SvgInitialFontSize;
        SvgTextWhitespace whitespace = SvgTextWhitespace.XmlDefault;
        foreach (XElement owner in element.AncestorsAndSelf().Reverse())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var declarations = ReferenceEquals(owner, element) ? style : ReadSvgStyleDeclarations(owner);
            if (declarations.ContainsKey("font")) return false;
            string? value = ReadSvgPresentationAttribute(owner, declarations, "font-size");
            if (!TryResolveSvgFontSize(value, size, out size)) return false;
            whitespace = ResolveSvgTextWhitespace(owner, declarations, whitespace);
        }
        if (!TryReadSvgTextMetrics(element, style, size, out var rootMetrics) ||
            !TryReadSvgTextPosition(element, default, viewWidth, viewHeight, size, startsLabel: true,
                cancellationToken, out var initialPosition)) return false;
        if (!TryReadPlainSvgText(element, style, rootMetrics, paintServers, IsSvgVisibilityHidden(Inherited("visibility")), whitespace,
                initialAnchor, initialPosition, viewWidth, viewHeight, cancellationToken,
                out string valueText, out bool[]? hiddenCharacters, out string anchor,
                out var position, out SvgTextPaint paint, out SvgTextMetrics metrics)) return false;
        double x = position.X + position.Dx;
        double y = position.Y + position.Dy;
        if (!double.IsFinite(x) || !double.IsFinite(y)) return false;
        if (OoxComplexScript.DetectNeeds(valueText) != OoxComplexScriptKind.None ||
            valueText.EnumerateRunes().Any(r => Rune.IsControl(r))) return false;
        string family = metrics.FamilyList.Split(',')[0].Trim().Trim('\'', '"');
        family = family.ToLowerInvariant() switch { "serif" => "Times New Roman", "sans-serif" => "Arial", "monospace" => "Courier New", _ => family };
        if (family.Length == 0) return false;
        double opacity = paint.Alpha * ReadSvgOpacityValue(ReadSvgElementOpacity(element, style));
        label = new(valueText, hiddenCharacters, new FontRequest(family, metrics.Weight >= 600, metrics.Slant != "normal"),
            metrics.Size, x, y, anchor, paint.Color, opacity, metrics.Kerning);
        return true;
    }

    private static bool TryReadSvgTextPosition(XElement owner, SvgTextPosition inherited,
        double viewWidth, double viewHeight, double size, bool startsLabel,
        CancellationToken cancellationToken, out SvgTextPosition position)
    {
        position = inherited;
        if (!TryRead("x", viewWidth, inherited.X, relative: false, out double x) ||
            !TryRead("y", viewHeight, inherited.Y, relative: false, out double y) ||
            !TryRead("dx", viewWidth, inherited.Dx, relative: true, out double dx) ||
            !TryRead("dy", viewHeight, inherited.Dy, relative: true, out double dy)) return false;
        position = new(x, y, dx, dy);
        return true;

        bool TryRead(string name, double percentageBasis, double previous, bool relative, out double value)
        {
            value = previous;
            if (owner.Attribute(name) is not { } attribute) return true;
            ReadOnlySpan<char> text = attribute.Value.AsSpan().Trim();
            // A later absolute position starts another chunk. Only later
            // relative zeros preserve the existing single-chunk geometry.
            if (!startsLabel) return relative && IsZeroSvgTextPositionList(text, allowLengths: true, cancellationToken);
            if (relative && (!TryReadSvgTextPositionToken(text, out text, out ReadOnlySpan<char> remaining) ||
                (!remaining.IsEmpty && !IsZeroSvgTextPositionList(remaining, allowLengths: true, cancellationToken)))) return false;
            return text.EndsWith("%", StringComparison.Ordinal)
                ? TryReadSvgViewportLength(owner, text, percentageBasis, out value)
                : TryReadSvgLength(text, out value, relativeSize: size);
        }
    }

    private static bool IsZeroSvgTextPositionList(ReadOnlySpan<char> values, bool allowLengths, CancellationToken cancellationToken)
    {
        values = values.Trim();
        if (values.IsEmpty) return false;
        while (!values.IsEmpty)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryReadSvgTextPositionToken(values, out ReadOnlySpan<char> token, out values)) return false;
            bool valid = allowLengths
                ? TryReadSvgLength(token, out double value, relativeSize: 1d, allowPercent: true)
                : double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
            if (!valid || value != 0d) return false;
        }
        return true;
    }

    private static bool TryReadSvgTextPositionToken(ReadOnlySpan<char> values, out ReadOnlySpan<char> token,
        out ReadOnlySpan<char> remaining)
    {
        int separator = values.IndexOfAny(" ,\t\r\n");
        token = separator < 0 ? values : values[..separator];
        remaining = separator < 0 ? default : values[separator..].TrimStart();
        if (!remaining.StartsWith(",")) return !token.IsEmpty;
        remaining = remaining[1..].TrimStart();
        return !token.IsEmpty && !remaining.IsEmpty;
    }

    private static bool TryReadSvgTextMetrics(XElement element, IReadOnlyDictionary<string, string> style,
        double size, out SvgTextMetrics metrics)
    {
        metrics = default;
        if (!double.IsFinite(size) || size <= 0d) return false;
        string slant = ResolveSvgTextFontValue(ReadSvgInheritedAttribute(element, style, "font-style"), "normal").ToLowerInvariant();
        if (slant is not ("normal" or "italic" or "oblique") ||
            !TryReadSvgTextWeight(ReadSvgInheritedAttribute(element, style, "font-weight"), out int weight) ||
            !TryReadSvgTextKerning(element, style, out bool kerning)) return false;
        string family = ResolveSvgTextFontValue(ReadSvgInheritedAttribute(element, style, "font-family"), "serif");
        metrics = new(family, slant, weight, size, kerning);
        return true;
    }

    private static bool TryReadSvgTextKerning(XElement element, IReadOnlyDictionary<string, string> style, out bool enabled)
    {
        ReadOnlySpan<char> fontKerning = ReadSvgInheritedAttribute(element, style, "font-kerning").AsSpan().Trim();
        ReadOnlySpan<char> legacyKerning = ReadSvgInheritedAttribute(element, style, "kerning").AsSpan().Trim();
        bool fontDisabled = fontKerning.Equals("none", StringComparison.OrdinalIgnoreCase);
        bool legacyDisabled = legacyKerning is "0" || legacyKerning.Equals("0px", StringComparison.OrdinalIgnoreCase);
        enabled = !fontDisabled && !legacyDisabled;
        return (fontDisabled || fontKerning.IsEmpty || fontKerning.Equals("auto", StringComparison.OrdinalIgnoreCase) ||
                fontKerning.Equals("normal", StringComparison.OrdinalIgnoreCase)) &&
            (legacyDisabled || legacyKerning.IsEmpty || legacyKerning.Equals("auto", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsNeutralSvgTextStyle(string name, string? value)
    {
        ReadOnlySpan<char> token = value.AsSpan().Trim();
        // A plain span's owners have already passed the default-style check.
        bool reset = value is null || token.IsEmpty || IsSvgInitialValue(token) || InheritsSvgPropertyValue(value);
        return name switch
        {
            "letter-spacing" or "word-spacing" => reset ||
                token.Equals("normal", StringComparison.OrdinalIgnoreCase) ||
                IsZeroLength(token),
            "font-variant" or "font-stretch" or "unicode-bidi" =>
                reset || token.Equals("normal", StringComparison.OrdinalIgnoreCase),
            "text-decoration" or "clip-path" or "mask" or "filter" =>
                reset || token.Equals("none", StringComparison.OrdinalIgnoreCase),
            "direction" => reset || token.Equals("ltr", StringComparison.OrdinalIgnoreCase),
            "writing-mode" => reset || token.Equals("horizontal-tb", StringComparison.OrdinalIgnoreCase) ||
                token.Equals("lr", StringComparison.OrdinalIgnoreCase) || token.Equals("lr-tb", StringComparison.OrdinalIgnoreCase),
            "dominant-baseline" => reset || token.Equals("auto", StringComparison.OrdinalIgnoreCase) ||
                token.Equals("alphabetic", StringComparison.OrdinalIgnoreCase),
            "alignment-baseline" => reset || token.Equals("auto", StringComparison.OrdinalIgnoreCase) ||
                token.Equals("baseline", StringComparison.OrdinalIgnoreCase) || token.Equals("alphabetic", StringComparison.OrdinalIgnoreCase),
            "baseline-shift" => reset || token.Equals("baseline", StringComparison.OrdinalIgnoreCase) ||
                IsZeroLength(token, allowPercent: true),
            _ => false
        };

        static bool IsZeroLength(ReadOnlySpan<char> token, bool allowPercent = false) =>
            TryReadSvgLength(token, out double length, relativeSize: 1d, allowPercent: allowPercent) && length == 0d;
    }

    private static string ResolveSvgTextFontValue(string? value, string initial) =>
        value is null || IsSvgInitialValue(value) ? initial : value.Trim();

    private static bool TryReadSvgTextWeight(string? authoredValue, out int weight)
    {
        ReadOnlySpan<char> value = ResolveSvgTextFontValue(authoredValue, "normal");
        if (value.Equals("normal", StringComparison.OrdinalIgnoreCase))
        {
            weight = 400;
            return true;
        }
        if (value.Equals("bold", StringComparison.OrdinalIgnoreCase))
        {
            weight = 700;
            return true;
        }
        return int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out weight) && weight is >= 100 and <= 900;
    }

    private static bool TryResolveSvgFontSize(string? value, double inheritedSize, out double size)
    {
        if (value is null || InheritsSvgPropertyValue(value))
        {
            size = inheritedSize;
            return true;
        }
        if (IsSvgInitialValue(value))
        {
            size = SvgInitialFontSize;
            return true;
        }
        return TryReadSvgLength(value, out size, relativeSize: inheritedSize, allowPercent: true);
    }

    private static bool TryReadSvgTextFill(XElement element, IReadOnlyDictionary<string, string> style,
        SvgPaintServers paintServers, out SvgTextPaint paint) =>
        TryReadSvgTextPaint(element, style, paintServers, ReadSvgInheritedAttribute(element, style, "fill") ?? "black", out paint);

    private static bool TryReadSvgTextPaint(XElement element, IReadOnlyDictionary<string, string> style,
        SvgPaintServers paintServers, string value, out SvgTextPaint paint)
    {
        value = ResolveSvgPaintFallback(element, style, value, paintServers, out string? gradientId);
        if (value.AsSpan().Trim().Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            paint = new(null, 1d);
            return true;
        }
        if (gradientId is not null)
        {
            var (linear, radial, constantStop) = paintServers.FindGradient(gradientId);
            if (linear?.UniformStopOpacity == 0d || radial?.UniformStopOpacity == 0d)
            {
                paint = new(null, 0d);
                return true;
            }
            if (constantStop is { Opacity: double opacity } stop)
            {
                paint = new(stop.Color, opacity);
                return true;
            }
        }
        else if (TryReadSvgSolidColor(element, style, value, out var color, out double alpha))
        {
            paint = new(color, alpha);
            return true;
        }
        paint = default;
        return false;
    }

    // A nonpainting stroke leaves the glyph fill intact. Any positive
    // percentage basis is enough to recognize a zero stroke width.
    private static bool HasSvgTextStroke(XElement element, IReadOnlyDictionary<string, string> style,
        SvgPaintServers paintServers)
    {
        if (ReadSvgInheritedAttribute(element, style, "stroke") is not { } stroke ||
            ReadSvgOpacityValue(ReadSvgInheritedAttribute(element, style, "stroke-opacity")) == 0d ||
            (TryReadSvgStrokeWidth(style, element, percentageBasis: 1d, out double width) && width == 0d))
            return false;
        return !TryReadSvgTextPaint(element, style, paintServers, stroke, out var paint) ||
            paint.Color is not null && paint.Alpha > 0d;
    }

    private static string? ResolveSvgTextAnchor(string? value, string inherited)
    {
        if (string.IsNullOrWhiteSpace(value) || InheritsSvgPropertyValue(value)) return inherited;
        string token = value.Trim().ToLowerInvariant();
        return token switch
        {
            "initial" => "start",
            "start" or "middle" or "end" => token,
            _ => null
        };
    }

    private static readonly XName SvgXmlSpace = XNamespace.Xml + "space";

    private enum SvgTextWhitespace
    {
        XmlDefault,
        XmlPreserve,
        CssCollapse,
        CssPreserve,
        CssPreLine
    }

    private static SvgTextWhitespace ResolveSvgTextWhitespace(XElement owner, IReadOnlyDictionary<string, string> declarations,
        SvgTextWhitespace inherited)
    {
        // Local CSS takes precedence over the legacy XML declaration.
        return ReadSvgPresentationAttribute(owner, declarations, "white-space")?.Trim().ToLowerInvariant() switch
        {
            "inherit" or "unset" => inherited,
            "initial" or "normal" or "nowrap" => SvgTextWhitespace.CssCollapse,
            "pre" or "pre-wrap" or "break-spaces" => SvgTextWhitespace.CssPreserve,
            "pre-line" => SvgTextWhitespace.CssPreLine,
            _ => (string?)owner.Attribute(SvgXmlSpace) switch
            {
                null => inherited,
                "preserve" => SvgTextWhitespace.XmlPreserve,
                _ => SvgTextWhitespace.XmlDefault
            }
        };
    }

    // Advancing characters share one font and size. Painted glyphs also
    // share one effective fill.
    private static bool TryReadPlainSvgText(XElement element, IReadOnlyDictionary<string, string> elementStyle, SvgTextMetrics rootMetrics,
        SvgPaintServers paintServers, bool hidden, SvgTextWhitespace whitespace, string initialAnchor,
        SvgTextPosition initialPosition, double viewWidth, double viewHeight,
        CancellationToken cancellationToken, out string text, out bool[]? hiddenCharacters, out string anchor,
        out SvgTextPosition firstPosition, out SvgTextPaint paint, out SvgTextMetrics metrics)
    {
        text = string.Empty;
        hiddenCharacters = null;
        anchor = initialAnchor;
        firstPosition = initialPosition;
        paint = default;
        metrics = rootMetrics;
        SvgTextPaint? commonPaint = null;
        SvgTextMetrics? commonMetrics = null;
        var characters = new StringBuilder();
        List<bool>? visibility = null;
        (bool Hidden, string Anchor, SvgTextMetrics Metrics)? pendingSpace = null;
        string? chunkAnchor = null;
        SvgTextPosition chunkPosition = initialPosition;
        XName xmlLanguageName = XNamespace.Xml + "lang";
        XName languageName = "lang";
        string? labelLanguage = null;
        XName spanName = element.Name.Namespace + "tspan";
        XName linkName = element.Name.Namespace + "a";
        if (!Append(element, elementStyle, rootMetrics, hidden, ownerTransparent: false, whitespace, initialAnchor, initialPosition)) return false;

        text = characters.ToString();
        hiddenCharacters = visibility?.ToArray();
        anchor = chunkAnchor ?? initialAnchor;
        firstPosition = chunkPosition;
        paint = commonPaint.GetValueOrDefault();
        metrics = commonMetrics.GetValueOrDefault(rootMetrics);
        return true;

        string? DeclaredLanguage(XElement owner) =>
            (string?)owner.Attribute(xmlLanguageName) ?? (string?)owner.Attribute(languageName);

        bool Append(XElement owner, IReadOnlyDictionary<string, string> ownerStyle, SvgTextMetrics ownerMetrics,
            bool ownerHidden, bool ownerTransparent, SvgTextWhitespace space, string ownerAnchor, SvgTextPosition ownerPosition)
        {
            bool preserveSpace = space is SvgTextWhitespace.XmlPreserve or SvgTextWhitespace.CssPreserve;
            bool hideOutlines = ownerHidden || ownerTransparent;
            bool? paintsGlyph = null;
            foreach (XNode node in owner.Nodes())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (node is XText content)
                {
                    foreach (char source in content.Value)
                    {
                        if (source is '\r' or '\n')
                        {
                            if (space == SvgTextWhitespace.XmlDefault) continue;
                            // Forced lines need their own geometry.
                            if (space is SvgTextWhitespace.CssPreserve or SvgTextWhitespace.CssPreLine) return false;
                        }
                        // A preserved CSS tab advances to a tab stop.
                        if (source == '\t' && space == SvgTextWhitespace.CssPreserve) return false;
                        char character = source is '\r' or '\n' or '\t' ? ' ' : source;
                        if (!preserveSpace && character == ' ')
                        {
                            // Hold a collapsed space until there is content on both sides.
                            if (commonMetrics is not null) pendingSpace ??= (hideOutlines, ownerAnchor, ownerMetrics);
                            continue;
                        }
                        bool characterHidden = hideOutlines;
                        if (!characterHidden && !char.IsWhiteSpace(character) &&
                            !OoxTextBreakRules.IsNonPrintingBreakControl(character))
                        {
                            if (paintsGlyph is null)
                            {
                                if (!TryAdmitOwnerFill(owner, ownerStyle, out bool ownerPaintsGlyph)) return false;
                                paintsGlyph = ownerPaintsGlyph;
                            }
                            characterHidden = !paintsGlyph.Value;
                        }
                        if (pendingSpace is { } pending)
                        {
                            if (!AppendCharacter(' ', pending.Hidden, pending.Anchor, pending.Metrics, ownerPosition)) return false;
                            pendingSpace = null;
                        }
                        if (!AppendCharacter(character, characterHidden, ownerAnchor, ownerMetrics, ownerPosition)) return false;
                    }
                    continue;
                }
                if (node is not XElement child || child.Name.Namespace != owner.Name.Namespace ||
                    IsSvgMetadataElement(child)) continue;
                var style = ReadSvgStyleDeclarations(child);
                if (IsSvgDisplayNone(child, style)) continue;
                bool isLink = child.Name == linkName;
                if (child.Name != spanName && !isLink) return false;
                string? localSpace = (string?)child.Attribute(SvgXmlSpace);
                if (localSpace is not (null or "default" or "preserve") ||
                    !IsPlainChild(child, style, isLink)) return false;
                string? localVisibility = ReadSvgPresentationAttribute(child, style, "visibility");
                bool childHidden = string.IsNullOrWhiteSpace(localVisibility) || InheritsSvgPropertyValue(localVisibility)
                    ? ownerHidden : IsSvgVisibilityHidden(localVisibility);
                string? authoredOpacity = ReadSvgElementOpacity(child, style);
                double? opacity = authoredOpacity is null ? 1d : ParseSvgOpacity(authoredOpacity);
                if (opacity is not double localOpacity || localOpacity is > 0d and < 1d) return false;
                bool childTransparent = ownerTransparent || localOpacity == 0d;
                SvgTextPosition childPosition = ownerPosition;
                string? childAnchor = ResolveSvgTextAnchor(ReadSvgPresentationAttribute(child, style, "text-anchor"), ownerAnchor);
                if (childAnchor is null ||
                    !TryResolveSvgFontSize(ReadSvgPresentationAttribute(child, style, "font-size"), ownerMetrics.Size, out double childSize) ||
                    !TryReadSvgTextMetrics(child, style, childSize, out var childMetrics) ||
                    (!isLink && !TryReadSvgTextPosition(child, ownerPosition, viewWidth, viewHeight, childSize,
                        commonMetrics is null, cancellationToken, out childPosition)) ||
                    !Append(child, style, childMetrics, childHidden, childTransparent,
                        ResolveSvgTextWhitespace(child, style, space), childAnchor, childPosition)) return false;
            }
            return true;
        }

        bool TryAdmitOwnerFill(XElement owner, IReadOnlyDictionary<string, string> ownerStyle, out bool paintsGlyph)
        {
            paintsGlyph = false;
            if (HasSvgTextStroke(owner, ownerStyle, paintServers)) return false;
            double fillOpacity = ReadSvgOpacityValue(ReadSvgInheritedAttribute(owner, ownerStyle, "fill-opacity"));
            if (fillOpacity == 0d) return true;
            if (!TryReadSvgTextFill(owner, ownerStyle, paintServers, out var fill)) return false;

            double alpha = fill.Alpha * fillOpacity;
            if (fill.Color is null || alpha == 0d) return true;

            var effective = new SvgTextPaint(fill.Color, alpha);
            if (commonPaint is { } existing && effective != existing) return false;
            commonPaint = effective;
            paintsGlyph = true;
            return true;
        }

        bool IsPlainChild(XElement child, IReadOnlyDictionary<string, string> style, bool isLink)
        {
            if (DeclaredLanguage(child) is { } language)
            {
                labelLanguage ??= element.AncestorsAndSelf().Select(DeclaredLanguage).FirstOrDefault(value => value is not null) ?? string.Empty;
                if (!string.Equals(language, labelLanguage, StringComparison.OrdinalIgnoreCase)) return false;
            }

            return child.Attributes().All(IsPlainAttribute) &&
                style.All(declaration => RetainsLabelStyle(declaration.Key.ToLowerInvariant(), declaration.Value));

            bool IsPlainAttribute(XAttribute attribute)
            {
                if (attribute.Name == SvgXmlSpace || attribute.Name == xmlLanguageName || attribute.Name == languageName ||
                    attribute.IsNamespaceDeclaration) return true;
                string name = attribute.Name.LocalName;
                XNamespace ns = attribute.Name.Namespace;
                bool isXLink = ns.NamespaceName == "http://www.w3.org/1999/xlink";
                // Link attributes describe the action, not the label's paint.
                // SVG link actions retain their separate diagnostic.
                if (isLink && (isXLink
                    ? name is "href" or "show" or "actuate" or "type" or "role" or "arcrole" or "title"
                    : ns == XNamespace.None && name is ("href" or "target" or "download" or "ping" or
                        "rel" or "hreflang" or "type" or "referrerpolicy")))
                    return true;
                // Unknown namespaced attributes carry no SVG presentation.
                if (ns != XNamespace.None) return ns != XNamespace.Xml && !isXLink;
                if (name is "x" or "y") return !isLink;
                if (name is "dx" or "dy")
                    return !isLink || IsZeroSvgTextPositionList(attribute.Value, allowLengths: true, cancellationToken);
                if (name == "rotate") return IsZeroSvgTextPositionList(attribute.Value, allowLengths: false, cancellationToken);
                return name is "id" or "class" or "role" or "style" or "tabindex" or "lengthAdjust" ||
                    name.StartsWith("aria-", StringComparison.Ordinal) ||
                    name.StartsWith("data-", StringComparison.Ordinal) ||
                    RetainsLabelStyle(name, style.GetValueOrDefault(name, attribute.Value));
            }

            bool RetainsLabelStyle(string name, string value) => name switch
            {
                "display" or "visibility" or "opacity" or "white-space" or "text-anchor" => true,
                // Pointer hit-testing leaves static text paint intact.
                "pointer-events" => true,
                // Descendants of SVG text content are not transformable (CSS Transforms §2).
                "transform" or "transform-origin" or "transform-box" => true,
                "font-family" or "font-style" or "font-weight" or "font-size" or "font-kerning" or "kerning" => true,
                // Effective paint is checked where the glyphs are owned.
                "fill" or "color" or "stroke" or "stroke-opacity" or "stroke-width" or "stroke-linecap" or "stroke-linejoin" or
                    "stroke-miterlimit" or "stroke-dasharray" or "stroke-dashoffset" or "paint-order" => true,
                "fill-opacity" => InheritsSvgPropertyValue(value) || ParseSvgOpacity(value) is not null,
                _ => IsNeutralSvgTextStyle(name, value),
            };
        }

        bool AppendCharacter(char character, bool characterHidden, string ownerAnchor, SvgTextMetrics ownerMetrics,
            SvgTextPosition ownerPosition)
        {
            if (!OoxTextBreakRules.IsNonPrintingBreakControl(character))
            {
                if (commonMetrics is { } existing && !ownerMetrics.Matches(existing)) return false;
                if (commonMetrics is null)
                {
                    commonMetrics = ownerMetrics;
                    chunkAnchor = ownerAnchor;
                    chunkPosition = ownerPosition;
                }
            }
            if (characterHidden && visibility is null) visibility = [.. Enumerable.Repeat(false, characters.Length)];
            characters.Append(character);
            visibility?.Add(characterHidden);
            return true;
        }
    }

    private static bool TryPaintSvgText(PdfGraphicsBuilder graphics, XElement element, SvgPaintServers paintServers,
        IReadOnlyDictionary<FontRequest, SvgTextFont?>? fonts, SvgTransform transform, double viewWidth, double viewHeight,
        double sourceMinX, double sourceMinY, double imageX, double imageY, double imageWidth, double imageHeight,
        double scaleX, double scaleY, Action<int> admitSamples, CancellationToken cancellationToken)
    {
        if (!TryReadSvgText(element, paintServers, viewWidth, viewHeight, cancellationToken, out SvgTextLabel? label)) return false;
        if (label.Color is not { } color || label.Opacity == 0d) return true;
        if (fonts is null || !fonts.TryGetValue(label.Font, out SvgTextFont? prepared) || prepared is null) return false;
        OpenTypeFont font = prepared.Font;
        double unitScale = label.FontSize / font.UnitsPerEm;
        var glyphs = new List<(OpenTypeFont.OpenTypeGlyphOutline? Outline, double X)>();
        double advance = 0d;
        ushort previous = 0;
        int characterIndex = 0;
        foreach (Rune rune in label.Text.EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            bool hidden = label.HiddenCharacters?[characterIndex] ?? false;
            characterIndex += rune.Utf16SequenceLength;
            // This label has no soft wraps. Keep its source whitespace and
            // visibility, with no advance or outline for a break control.
            if (OoxTextBreakRules.IsNonPrintingBreakControl(rune.Value)) continue;
            ushort glyph = font.MapCodePoint(rune.Value);
            // Hidden text still takes space. An unmapped hidden character uses
            // the face's missing-glyph advance and needs no outline.
            if (glyph == 0 && !hidden) return false;
            if (label.Kerning && previous != 0 && glyph != 0) advance += font.GetKerning(previous, glyph) * unitScale;
            OpenTypeFont.OpenTypeGlyphOutline? outline = null;
            if (!hidden && !Rune.IsWhiteSpace(rune))
            {
                if (!font.TryReadGlyphOutline(glyph, out var parsed) || parsed.Contours.Count == 0) return false;
                outline = parsed;
                int points = parsed.Contours.Sum(c => c.Points.Count);
                // The source font owns outlines; the picture's existing coarse
                // construction grant admits the generated paths before appending.
                admitSamples(checked(1 + points / 4));
            }
            glyphs.Add((outline, advance));
            advance += font.GetAdvanceWidth(glyph) * unitScale;
            if (!double.IsFinite(advance)) return false;
            previous = glyph;
        }
        double origin = label.X - (label.Anchor == "middle" ? advance / 2d : label.Anchor == "end" ? advance : 0d);
        double a = scaleX * transform.M11, b = -scaleY * transform.M12;
        double c = -scaleX * transform.M21, d = scaleY * transform.M22;
        double e = imageX + (transform.OffsetX - sourceMinX) * scaleX;
        double f = imageY + imageHeight - (transform.OffsetY - sourceMinY) * scaleY;
        if (!new[] { origin, unitScale, a, b, c, d, e, f }.All(double.IsFinite)) return false;
        double shear = label.Font.Italic && !prepared.Resolution.Italic ? PdfGraphicsBuilder.SyntheticItalicShear : 0d;
        foreach (var glyph in glyphs)
            if (glyph.Outline is { } outline)
                foreach (var contour in outline.Contours)
                    foreach (var point in contour.Points)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        double px = origin + glyph.X + (point.X + shear * point.Y) * unitScale;
                        double py = -label.Y + point.Y * unitScale;
                        if (!double.IsFinite(a * px + c * py + e) || !double.IsFinite(b * px + d * py + f)) return false;
                    }
        graphics.SaveState();
        try
        {
            graphics.ClipRectangle(imageX, imageY, imageWidth, imageHeight);
            graphics.Transform(a, b, c, d, e, f);
            graphics.SetAlpha(label.Opacity, 1d);
            graphics.SetFillRgb(color.Red, color.Green, color.Blue);
            foreach (var glyph in glyphs)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (glyph.Outline is not { } outline) continue;
                PdfGlyphOutlinePath.AppendGlyphPath(graphics, font, outline, origin + glyph.X, -label.Y, label.FontSize, shear);
                graphics.FillCurrentPath();
            }
        }
        finally { graphics.RestoreState(); }
        return true;
    }
}
