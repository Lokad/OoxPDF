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
            if (HasSvgStylesheet(svg) ||
                !TryReadSvgSourceFrame(svg.Root, out _, out _, out double viewWidth, out double viewHeight) ||
                viewWidth <= 0d || viewHeight <= 0d) continue;
            var paintServerIds = CreateSvgPaintServerIdLookup(svg, cancellationToken);
            foreach (XElement element in svg.Descendants().Where(e => e.Name.LocalName == "text"))
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

    private static bool HasSvgStylesheet(XDocument svg) =>
        svg.DescendantNodes().OfType<XProcessingInstruction>().Any(p => p.Target == "xml-stylesheet" && !p.Ancestors().Any(IsSvgMetadataElement)) ||
        svg.Descendants().Any(e => e.Name.LocalName == "style" && !string.IsNullOrWhiteSpace(e.Value) &&
            !e.Ancestors().Any(IsSvgMetadataElement));

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
        bool IsDefault(string name, params string[] accepted) => Inherited(name) is not { } value ||
            accepted.Contains(value, StringComparer.OrdinalIgnoreCase);
        if (!IsDefault("direction", "ltr") || !IsDefault("unicode-bidi", "normal") ||
            !IsDefault("writing-mode", "lr", "lr-tb", "horizontal-tb") ||
            !IsDefault("dominant-baseline", "auto", "alphabetic") ||
            !IsDefault("alignment-baseline", "auto", "baseline", "alphabetic") ||
            !IsDefault("baseline-shift", "baseline", "0", "0px") ||
            !IsDefault("font-variant", "normal") || !IsDefault("font-stretch", "normal") ||
            !IsDefault("letter-spacing", "normal", "0", "0px") ||
            !IsDefault("word-spacing", "normal", "0", "0px") ||
            !IsDefault("text-decoration", "none") || !IsDefault("font-kerning", "auto", "normal", "none") ||
            !IsDefault("kerning", "auto", "0", "0px") ||
            !IsDefault("clip-path", "none") || !IsDefault("mask", "none") || !IsDefault("filter", "none")) return false;
        // A nonpainting stroke leaves the supported glyph fill intact. The
        // width check needs only a positive percentage basis to recognize zero.
        if (Inherited("stroke") is { } stroke &&
            ReadSvgOpacityValue(Inherited("stroke-opacity")) != 0d &&
            !(TryReadSvgStrokeWidth(style, element, percentageBasis: 1d, out double strokeWidth) && strokeWidth == 0d))
        {
            string strokePaint = ResolveSvgPaintFallback(element, style, stroke, paintServerIds, out _);
            if (!strokePaint.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                (!TryReadSvgSolidColor(element, style, strokePaint, out _, out double strokeOpacity) || strokeOpacity > 0d))
                return false;
        }
        string anchor = Inherited("text-anchor")?.ToLowerInvariant() ?? "start";
        if (anchor is not ("start" or "middle" or "end")) return false;
        string fontStyle = Inherited("font-style")?.ToLowerInvariant() ?? "normal";
        if (fontStyle is not ("normal" or "italic" or "oblique")) return false;
        string weight = Inherited("font-weight")?.ToLowerInvariant() ?? "normal";
        bool bold = weight == "bold";
        if (weight is not ("normal" or "bold"))
        {
            if (!int.TryParse(weight, NumberStyles.None, CultureInfo.InvariantCulture, out int numeric) || numeric is < 100 or > 900) return false;
            bold = numeric >= 600;
        }
        const double initialFontSize = 16d;
        double size = initialFontSize;
        foreach (XElement owner in element.AncestorsAndSelf().Reverse())
        {
            string? value = ReadSvgPresentationAttribute(owner, ReadSvgStyleDeclarations(owner), "font-size");
            if (value is null || InheritsSvgPropertyValue(value)) continue;
            if (IsSvgInitialValue(value))
                size = initialFontSize;
            else if (!TryReadSvgLength(value, out size, relativeSize: size, allowPercent: true))
                return false;
        }
        if (size <= 0d || !TryReadPosition("x", viewWidth, out double x) ||
            !TryReadPosition("y", viewHeight, out double y) ||
            !TryReadPosition("dx", viewWidth, out double dx) ||
            !TryReadPosition("dy", viewHeight, out double dy)) return false;
        x += dx; y += dy;
        if (!double.IsFinite(x) || !double.IsFinite(y)) return false;
        string fill = ResolveSvgPaintFallback(element, style, Inherited("fill") ?? "black", paintServerIds, out _);
        RgbColor? color = null;
        double colorOpacity = 1d;
        if (!string.Equals(fill, "none", StringComparison.OrdinalIgnoreCase))
        {
            if (!TryReadSvgSolidColor(element, style, fill, out RgbColor parsed, out colorOpacity)) return false;
            color = parsed;
        }
        if (!TryReadPlainSvgText(element, IsSvgVisibilityHidden(Inherited("visibility")), cancellationToken,
                out string valueText, out bool[]? hiddenCharacters)) return false;
        if (OoxComplexScript.DetectNeeds(valueText) != OoxComplexScriptKind.None ||
            valueText.EnumerateRunes().Any(r => Rune.IsControl(r))) return false;
        string family = (Inherited("font-family") ?? "serif").Split(',')[0].Trim().Trim('\'', '"');
        family = family.ToLowerInvariant() switch { "serif" => "Times New Roman", "sans-serif" => "Arial", "monospace" => "Courier New", _ => family };
        if (family.Length == 0) return false;
        double opacity = colorOpacity * ReadSvgOpacityValue(Inherited("fill-opacity")) *
            ReadSvgOpacityValue(ReadSvgElementOpacity(element, style));
        label = new(valueText, hiddenCharacters, new FontRequest(family, bold, fontStyle != "normal"), size, x, y, anchor, color, opacity,
            !string.Equals(Inherited("font-kerning"), "none", StringComparison.OrdinalIgnoreCase) &&
            Inherited("kerning") is not ("0" or "0px"));
        return true;

        bool TryReadPosition(string name, double percentageBasis, out double value)
        {
            string text = ((string?)element.Attribute(name) ?? "0").Trim();
            return text.EndsWith('%')
                ? TryReadSvgViewportLength(element, text, percentageBasis, out value)
                : TryReadSvgLength(text, out value, relativeSize: size);
        }
    }

    // Plain spans and link text share the label's metrics. Visibility controls
    // which outlines paint. Only display:none removes characters from layout.
    private static bool TryReadPlainSvgText(XElement element, bool hidden, CancellationToken cancellationToken,
        out string text, out bool[]? hiddenCharacters)
    {
        text = string.Empty;
        hiddenCharacters = null;
        var characters = new StringBuilder();
        List<bool>? visibility = null;
        XName spaceName = XNamespace.Xml + "space";
        bool preserve = element.AncestorsAndSelf().Select(e => (string?)e.Attribute(spaceName))
            .FirstOrDefault(value => value is not null) == "preserve";
        string spacePolicy = preserve ? "preserve" : "default";
        XName spanName = element.Name.Namespace + "tspan";
        XName linkName = element.Name.Namespace + "a";
        if (!Append(element, hidden)) return false;

        string untrimmed = characters.ToString();
        ReadOnlySpan<char> normalized = untrimmed.AsSpan();
        int start = 0;
        if (!preserve)
        {
            // Tabs and XML line endings have already been normalized. Other
            // characters such as no-break spaces still own glyph advances.
            ReadOnlySpan<char> trimmed = normalized.TrimStart(' ');
            start = normalized.Length - trimmed.Length;
            normalized = trimmed.TrimEnd(' ');
        }
        text = normalized.Length == untrimmed.Length ? untrimmed : normalized.ToString();
        if (visibility is not null && text.Length != 0)
        {
            hiddenCharacters = new bool[text.Length];
            visibility.CopyTo(start, hiddenCharacters, 0, text.Length);
        }
        return true;

        bool Append(XElement owner, bool ownerHidden)
        {
            foreach (XNode node in owner.Nodes())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (node is XText content)
                {
                    foreach (char source in content.Value)
                    {
                        if (!preserve && source is '\r' or '\n') continue;
                        char character = source is '\r' or '\n' or '\t' ? ' ' : source;
                        if (!preserve && character == ' ' && characters.Length != 0 && characters[^1] == ' ') continue;
                        if (ownerHidden && visibility is null) visibility = [.. Enumerable.Repeat(false, characters.Length)];
                        characters.Append(character);
                        visibility?.Add(ownerHidden);
                    }
                    continue;
                }
                if (node is not XElement child || IsSvgMetadataElement(child)) continue;
                bool isLink = child.Name == linkName;
                if (child.Name != spanName && !isLink) return false;
                var style = ReadSvgStyleDeclarations(child);
                if (IsSvgDisplayNone(child, style)) continue;
                // A child can restate the label's existing whitespace policy.
                if (child.Attributes().Any(attribute => attribute.Name == spaceName
                        ? attribute.Value != spacePolicy : !IsPlainSvgTextChildAttribute(attribute, isLink)) ||
                    style.Keys.Any(property => property is not ("display" or "visibility"))) return false;
                string? localVisibility = ReadSvgPresentationAttribute(child, style, "visibility");
                bool childHidden = string.IsNullOrWhiteSpace(localVisibility) || InheritsSvgPropertyValue(localVisibility)
                    ? ownerHidden : IsSvgVisibilityHidden(localVisibility);
                if (!Append(child, childHidden)) return false;
            }
            return true;
        }
    }

    private static bool IsPlainSvgTextChildAttribute(XAttribute attribute, bool isLink) =>
        attribute.IsNamespaceDeclaration ||
        (isLink && attribute.Name.LocalName == "href" &&
         (attribute.Name.Namespace == XNamespace.None || attribute.Name.NamespaceName == "http://www.w3.org/1999/xlink")) ||
        attribute.Name.Namespace == XNamespace.None &&
        (attribute.Name.LocalName is "id" or "class" or "role" or "display" or "visibility" or "style" ||
         attribute.Name.LocalName.StartsWith("aria-", StringComparison.Ordinal) ||
         attribute.Name.LocalName.StartsWith("data-", StringComparison.Ordinal));

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
            ushort glyph = font.MapCodePoint(rune.Value);
            if (glyph == 0) return false;
            if (label.Kerning && previous != 0) advance += font.GetKerning(previous, glyph) * unitScale;
            OpenTypeFont.OpenTypeGlyphOutline? outline = null;
            bool hidden = label.HiddenCharacters?[characterIndex] ?? false;
            characterIndex += rune.Utf16SequenceLength;
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
