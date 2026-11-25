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
        IEnumerable<PptxSceneNode> nodes = EnumerateRenderedSceneNodes(slide.LayoutNodes, false)
            .Concat(EnumerateRenderedSceneNodes(slide.SlideNodes, true));
        if (slide.ShowMasterShapes) nodes = EnumerateRenderedSceneNodes(slide.MasterNodes, false).Concat(nodes);
        foreach (PptxSceneNode node in nodes)
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
            var content = EnumerateSvgContent(svg, cancellationToken);
            if (!TryReadSvgSourceFrame(svg.Root, out _, out _, out double viewWidth, out double viewHeight) ||
                viewWidth <= 0d || viewHeight <= 0d || HasSvgStylesheet(content)) continue;
            var elements = content.OfType<XElement>();
            var paintServerIds = CreateSvgPaintServerIdLookup(EnumerateSvgPaintServers(elements), cancellationToken);
            foreach (XElement element in elements.Where(e => e.Name.LocalName == "text"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (IsSvgDefinitionElement(element) || IsInFullyTransparentSvgContainer(element) ||
                    IsSvgDisplaySuppressed(element, ReadSvgStyleDeclarations(element)) ||
                    !TryReadSvgText(element, paintServerIds, viewWidth, viewHeight, cancellationToken, out SvgTextLabel? label) || label.Color is null ||
                    label.Opacity == 0d || !label.HasVisibleText || fonts.ContainsKey(label.Font)) continue;
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
        string Anchor, RgbColor? Color, double Opacity, bool Kerning)
    {
        public bool HasVisibleText => Text.Length != 0 && (HiddenCharacters is null || HiddenCharacters.AsSpan().Contains(false));
    }

    private readonly record struct SvgTextFill(RgbColor? Color, double Alpha);

    private readonly record struct SvgTextAppearance(string FamilyList, string Slant, int Weight, double Size,
        SvgTextFill Fill, double FillOpacity, bool Kerning);

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

    private static bool TryReadSvgText(XElement element, Lazy<HashSet<string>> paintServerIds,
        double viewWidth, double viewHeight, CancellationToken cancellationToken, [NotNullWhen(true)] out SvgTextLabel? label)
    {
        label = null;
        if (element.AncestorsAndSelf().Any(e => ReadSvgStyleDeclarations(e).ContainsKey("font")) ||
            element.Attributes().Any(a => a.Name.LocalName is "rotate" or "textLength" or "lengthAdjust")) return false;
        var style = ReadSvgStyleDeclarations(element);
        string? Inherited(string name) => ReadSvgInheritedAttribute(element, style, name)?.Trim();
        bool PreservesTextStyle(string name) => IsNeutralSvgTextStyle(name, Inherited(name));
        if (!PreservesTextStyle("direction") || !PreservesTextStyle("unicode-bidi") ||
            !PreservesTextStyle("writing-mode") || !PreservesTextStyle("dominant-baseline") ||
            !PreservesTextStyle("alignment-baseline") || !PreservesTextStyle("baseline-shift") ||
            !PreservesTextStyle("font-variant") || !PreservesTextStyle("font-stretch") ||
            !PreservesTextStyle("letter-spacing") || !PreservesTextStyle("word-spacing") ||
            !PreservesTextStyle("text-decoration") || !TryReadSvgTextKerning(element, style, out bool kerning) ||
            !PreservesTextStyle("clip-path") || !PreservesTextStyle("mask") || !PreservesTextStyle("filter")) return false;
        if (HasSvgTextStroke(element, style, paintServerIds)) return false;
        string anchor = Inherited("text-anchor")?.ToLowerInvariant() ?? "start";
        if (anchor is not ("start" or "middle" or "end")) return false;
        string fontStyle = Inherited("font-style")?.ToLowerInvariant() ?? "normal";
        if (fontStyle is not ("normal" or "italic" or "oblique")) return false;
        if (!TryReadSvgTextWeight(Inherited("font-weight") ?? "normal", out int weight)) return false;
        string familyList = Inherited("font-family") ?? "serif";
        double size = SvgInitialFontSize;
        foreach (XElement owner in element.AncestorsAndSelf().Reverse())
        {
            string? value = ReadSvgPresentationAttribute(owner, ReadSvgStyleDeclarations(owner), "font-size");
            if (!TryResolveSvgFontSize(value, size, out size)) return false;
        }
        if (size <= 0d || !TryReadPosition("x", viewWidth, out double x) ||
            !TryReadPosition("y", viewHeight, out double y) ||
            !TryReadPosition("dx", viewWidth, out double dx) ||
            !TryReadPosition("dy", viewHeight, out double dy)) return false;
        x += dx; y += dy;
        if (!double.IsFinite(x) || !double.IsFinite(y)) return false;
        if (!TryReadSvgTextFill(element, style, paintServerIds, out var paint)) return false;
        double fillOpacity = ReadSvgOpacityValue(Inherited("fill-opacity"));
        var appearance = new SvgTextAppearance(familyList, fontStyle, weight, size, paint, fillOpacity, kerning);
        if (!TryReadPlainSvgText(element, appearance, paintServerIds, IsSvgVisibilityHidden(Inherited("visibility")), cancellationToken,
                out string valueText, out bool[]? hiddenCharacters)) return false;
        if (OoxComplexScript.DetectNeeds(valueText) != OoxComplexScriptKind.None ||
            valueText.EnumerateRunes().Any(r => Rune.IsControl(r))) return false;
        string family = familyList.Split(',')[0].Trim().Trim('\'', '"');
        family = family.ToLowerInvariant() switch { "serif" => "Times New Roman", "sans-serif" => "Arial", "monospace" => "Courier New", _ => family };
        if (family.Length == 0) return false;
        double opacity = paint.Alpha * fillOpacity *
            ReadSvgOpacityValue(ReadSvgElementOpacity(element, style));
        label = new(valueText, hiddenCharacters, new FontRequest(family, weight >= 600, fontStyle != "normal"), size, x, y, anchor, paint.Color, opacity, kerning);
        return true;

        bool TryReadPosition(string name, double percentageBasis, out double value)
        {
            string text = ((string?)element.Attribute(name) ?? "0").Trim();
            return text.EndsWith('%')
                ? TryReadSvgViewportLength(element, text, percentageBasis, out value)
                : TryReadSvgLength(text, out value, relativeSize: size);
        }
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

    private static bool TryReadSvgTextWeight(ReadOnlySpan<char> value, out int weight)
    {
        value = value.Trim();
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
        Lazy<HashSet<string>> paintServerIds, out SvgTextFill fill)
    {
        string value = ResolveSvgPaintFallback(element, style,
            ReadSvgInheritedAttribute(element, style, "fill")?.Trim() ?? "black", paintServerIds, out _);
        if (value.AsSpan().Trim().Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            fill = new(null, 1d);
            return true;
        }
        if (TryReadSvgSolidColor(element, style, value, out var color, out double alpha))
        {
            fill = new(color, alpha);
            return true;
        }
        fill = default;
        return false;
    }

    // A nonpainting stroke leaves the glyph fill intact. Any positive
    // percentage basis is enough to recognize a zero stroke width.
    private static bool HasSvgTextStroke(XElement element, IReadOnlyDictionary<string, string> style,
        Lazy<HashSet<string>> paintServerIds)
    {
        if (ReadSvgInheritedAttribute(element, style, "stroke") is not { } stroke ||
            ReadSvgOpacityValue(ReadSvgInheritedAttribute(element, style, "stroke-opacity")) == 0d ||
            (TryReadSvgStrokeWidth(style, element, percentageBasis: 1d, out double width) && width == 0d))
            return false;
        string paint = ResolveSvgPaintFallback(element, style, stroke, paintServerIds, out _);
        return !paint.Equals("none", StringComparison.OrdinalIgnoreCase) &&
            (!TryReadSvgSolidColor(element, style, paint, out _, out double colorAlpha) || colorAlpha > 0d);
    }

    // Plain spans and link text share the label's metrics. Visibility and
    // zero opacity suppress outlines while retaining the glyph advances.
    private static bool TryReadPlainSvgText(XElement element, SvgTextAppearance appearance,
        Lazy<HashSet<string>> paintServerIds, bool hidden, CancellationToken cancellationToken,
        out string text, out bool[]? hiddenCharacters)
    {
        text = string.Empty;
        hiddenCharacters = null;
        var characters = new StringBuilder();
        List<bool>? visibility = null;
        bool? pendingSpaceHidden = null;
        XName spaceName = XNamespace.Xml + "space";
        bool preserve = element.AncestorsAndSelf().Select(e => (string?)e.Attribute(spaceName))
            .FirstOrDefault(value => value is not null) == "preserve";
        XName spanName = element.Name.Namespace + "tspan";
        XName linkName = element.Name.Namespace + "a";
        if (!Append(element, hidden, ownerTransparent: false, preserve)) return false;

        text = characters.ToString();
        hiddenCharacters = visibility?.ToArray();
        return true;

        bool Append(XElement owner, bool ownerHidden, bool ownerTransparent, bool preserveSpace)
        {
            foreach (XNode node in owner.Nodes())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (node is XText content)
                {
                    bool hideOutlines = ownerHidden || ownerTransparent;
                    foreach (char source in content.Value)
                    {
                        if (!preserveSpace && source is '\r' or '\n') continue;
                        char character = source is '\r' or '\n' or '\t' ? ' ' : source;
                        if (!preserveSpace && character == ' ')
                        {
                            // Hold a collapsed space until there is content on both sides.
                            if (characters.Length != 0) pendingSpaceHidden ??= hideOutlines;
                            continue;
                        }
                        if (pendingSpaceHidden is { } spaceHidden)
                        {
                            AppendCharacter(' ', spaceHidden);
                            pendingSpaceHidden = null;
                        }
                        AppendCharacter(character, hideOutlines);
                    }
                    continue;
                }
                if (node is not XElement child || child.Name.Namespace != owner.Name.Namespace ||
                    IsSvgMetadataElement(child)) continue;
                bool isLink = child.Name == linkName;
                if (child.Name != spanName && !isLink) return false;
                var style = ReadSvgStyleDeclarations(child);
                if (IsSvgDisplayNone(child, style)) continue;
                string? localSpace = (string?)child.Attribute(spaceName);
                if (localSpace is not (null or "default" or "preserve") ||
                    !IsPlainChild(child, style, isLink)) return false;
                string? localVisibility = ReadSvgPresentationAttribute(child, style, "visibility");
                bool childHidden = string.IsNullOrWhiteSpace(localVisibility) || InheritsSvgPropertyValue(localVisibility)
                    ? ownerHidden : IsSvgVisibilityHidden(localVisibility);
                string? authoredOpacity = ReadSvgElementOpacity(child, style);
                double? opacity = authoredOpacity is null ? 1d : ParseSvgOpacity(authoredOpacity);
                if (opacity is not double localOpacity || localOpacity is > 0d and < 1d) return false;
                bool childTransparent = ownerTransparent || localOpacity == 0d;
                if (!Append(child, childHidden, childTransparent, localSpace is null ? preserveSpace : localSpace == "preserve")) return false;
            }
            return true;
        }

        bool IsPlainChild(XElement child, IReadOnlyDictionary<string, string> style, bool isLink)
        {
            bool? hasLabelFill = null;
            bool? hasLabelKerning = null;
            bool? hasStroke = null;
            return child.Attributes().All(IsPlainAttribute) &&
                style.All(declaration => RetainsLabelStyle(declaration.Key.ToLowerInvariant(), declaration.Value));

            bool IsPlainAttribute(XAttribute attribute) =>
                attribute.Name == spaceName || attribute.IsNamespaceDeclaration ||
                (isLink && attribute.Name.LocalName == "href" &&
                 (attribute.Name.Namespace == XNamespace.None || attribute.Name.NamespaceName == "http://www.w3.org/1999/xlink")) ||
                attribute.Name.Namespace == XNamespace.None &&
                (attribute.Name.LocalName is "id" or "class" or "role" or "style" ||
                 attribute.Name.LocalName.StartsWith("aria-", StringComparison.Ordinal) ||
                 attribute.Name.LocalName.StartsWith("data-", StringComparison.Ordinal) ||
                 RetainsLabelStyle(attribute.Name.LocalName, style.GetValueOrDefault(attribute.Name.LocalName, attribute.Value)));

            bool RetainsLabelStyle(string name, string value) => name switch
            {
                "display" or "visibility" or "opacity" => true,
                "font-family" => InheritsSvgPropertyValue(value) ||
                    value.AsSpan().Trim().Equals(appearance.FamilyList, StringComparison.OrdinalIgnoreCase),
                "font-style" => InheritsSvgPropertyValue(value) ||
                    value.AsSpan().Trim().Equals(appearance.Slant, StringComparison.OrdinalIgnoreCase),
                "font-weight" => InheritsSvgPropertyValue(value) ||
                    TryReadSvgTextWeight(value, out int weight) && weight == appearance.Weight,
                "font-size" => TryResolveSvgFontSize(value, appearance.Size, out double size) && size == appearance.Size,
                "fill" or "color" =>
                    (hasLabelFill ??= TryReadSvgTextFill(child, style, paintServerIds, out var fill) &&
                        fill == appearance.Fill) == true && !PaintsStroke(),
                "fill-opacity" => InheritsSvgPropertyValue(value) || ParseSvgOpacity(value) == appearance.FillOpacity,
                "stroke" or "stroke-opacity" or "stroke-width" => !PaintsStroke(),
                "font-kerning" or "kerning" =>
                    (hasLabelKerning ??= TryReadSvgTextKerning(child, style, out bool kerning) && kerning == appearance.Kerning) == true,
                _ => IsNeutralSvgTextStyle(name, value),
            };

            bool PaintsStroke() => (hasStroke ??= HasSvgTextStroke(child, style, paintServerIds)) == true;
        }

        void AppendCharacter(char character, bool characterHidden)
        {
            if (characterHidden && visibility is null) visibility = [.. Enumerable.Repeat(false, characters.Length)];
            characters.Append(character);
            visibility?.Add(characterHidden);
        }
    }

    private static bool TryPaintSvgText(PdfGraphicsBuilder graphics, XElement element, Lazy<HashSet<string>> paintServerIds,
        IReadOnlyDictionary<FontRequest, SvgTextFont?>? fonts, SvgTransform transform, double viewWidth, double viewHeight,
        double sourceMinX, double sourceMinY, double imageX, double imageY, double imageWidth, double imageHeight,
        double scaleX, double scaleY, Action<int> admitSamples, CancellationToken cancellationToken)
    {
        if (!TryReadSvgText(element, paintServerIds, viewWidth, viewHeight, cancellationToken, out SvgTextLabel? label)) return false;
        if (label.Color is not { } color || label.Opacity == 0d || !label.HasVisibleText) return true;
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
