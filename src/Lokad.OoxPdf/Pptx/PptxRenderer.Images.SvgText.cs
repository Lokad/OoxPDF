using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Text.RegularExpressions;
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
            if (node.Picture?.Resource is { } image && image.ContentType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase))
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
            if (HasSvgStylesheet(svg)) continue;
            foreach (XElement element in svg.Descendants().Where(e => e.Name.LocalName == "text"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (IsSvgDefinitionElement(element) || IsInFullyTransparentSvgContainer(element) || IsHiddenSvgText(element) ||
                    !TryReadSvgText(element, out SvgTextLabel? label) || label.Color is null ||
                    label.Opacity == 0d || label.Text.Length == 0 || fonts.ContainsKey(label.Font)) continue;
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

    private sealed record SvgTextLabel(string Text, FontRequest Font, double FontSize, double X, double Y,
        string Anchor, RgbColor? Color, double Opacity, bool Kerning);

    private static bool HasSvgStylesheet(XDocument svg) =>
        svg.DescendantNodes().OfType<XProcessingInstruction>().Any(p => p.Target == "xml-stylesheet" && !p.Ancestors().Any(IsSvgMetadataElement)) ||
        svg.Descendants().Any(e => e.Name.LocalName == "style" && !string.IsNullOrWhiteSpace(e.Value) &&
            !e.Ancestors().Any(IsSvgMetadataElement));

    private static bool IsSvgMetadataElement(XElement element) =>
        element.Name.LocalName is "title" or "desc" or "metadata";

    private static bool IsHiddenSvgText(XElement element)
    {
        string? visibility = ReadSvgInheritedAttribute(element, ReadSvgStyleDeclarations(element), "visibility");
        return visibility?.Trim().ToLowerInvariant() is "hidden" or "collapse" ||
            element.AncestorsAndSelf().Any(e => string.Equals(
                ReadSvgPresentationAttribute(e, ReadSvgStyleDeclarations(e), "display")?.Trim(), "none", StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryReadSvgText(XElement element, [NotNullWhen(true)] out SvgTextLabel? label)
    {
        label = null;
        if (element.Elements().Any(child => !IsSvgMetadataElement(child)) ||
            element.AncestorsAndSelf().Any(e => ReadSvgStyleDeclarations(e).ContainsKey("font")) ||
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
            !IsDefault("stroke", "none") || !IsDefault("clip-path", "none") ||
            !IsDefault("mask", "none") || !IsDefault("filter", "none")) return false;
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
        double size = 16d;
        foreach (XElement owner in element.AncestorsAndSelf().Reverse())
        {
            string? value = ReadSvgPresentationAttribute(owner, ReadSvgStyleDeclarations(owner), "font-size");
            if (value is not null && !TryReadSvgTextLength(value, size, out size, allowPercent: true)) return false;
        }
        if (size <= 0d || !TryReadSvgTextLength((string?)element.Attribute("x") ?? "0", null, out double x) ||
            !TryReadSvgTextLength((string?)element.Attribute("y") ?? "0", null, out double y) ||
            !TryReadSvgTextLength((string?)element.Attribute("dx") ?? "0", size, out double dx) ||
            !TryReadSvgTextLength((string?)element.Attribute("dy") ?? "0", size, out double dy)) return false;
        x += dx; y += dy;
        if (!double.IsFinite(x) || !double.IsFinite(y)) return false;
        string? fill = Inherited("fill");
        RgbColor? color = null;
        if (!string.Equals(fill, "none", StringComparison.OrdinalIgnoreCase))
        {
            if (!RgbColor.TryParseCssColor(fill ?? "black", out RgbColor parsed)) return false;
            color = parsed;
        }
        string valueText = string.Concat(element.Nodes().OfType<XText>().Select(node => node.Value));
        bool preserve = element.AncestorsAndSelf().Select(e => (string?)e.Attribute(XNamespace.Xml + "space"))
            .FirstOrDefault(v => v is not null) == "preserve";
        valueText = preserve ? valueText.Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ') :
            Regex.Replace(valueText.Replace("\r", "").Replace("\n", "").Replace('\t', ' '), " +", " ").Trim();
        if (OoxComplexScript.DetectNeeds(valueText) != OoxComplexScriptKind.None ||
            valueText.EnumerateRunes().Any(r => Rune.IsControl(r))) return false;
        string family = (Inherited("font-family") ?? "serif").Split(',')[0].Trim().Trim('\'', '"');
        family = family.ToLowerInvariant() switch { "serif" => "Times New Roman", "sans-serif" => "Arial", "monospace" => "Courier New", _ => family };
        if (family.Length == 0) return false;
        double opacity = ReadSvgOpacityValue(Inherited("fill-opacity")) *
            ReadSvgOpacityValue(ReadSvgPresentationAttribute(element, style, "opacity"));
        label = new(valueText, new FontRequest(family, bold, fontStyle != "normal"), size, x, y, anchor, color, opacity,
            !string.Equals(Inherited("font-kerning"), "none", StringComparison.OrdinalIgnoreCase) &&
            Inherited("kerning") is not ("0" or "0px"));
        return true;
    }

    private static bool TryReadSvgTextLength(string text, double? relativeSize, out double value, bool allowPercent = false)
    {
        text = text.Trim();
        double scale = 1d;
        if (text.EndsWith("px", StringComparison.OrdinalIgnoreCase)) text = text[..^2];
        else if (text.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) { text = text[..^2]; scale = 96d / 72d; }
        else if (relativeSize is { } em && text.EndsWith("em", StringComparison.OrdinalIgnoreCase)) { text = text[..^2]; scale = em; }
        else if (allowPercent && relativeSize is { } percent && text.EndsWith('%')) { text = text[..^1]; scale = percent / 100d; }
        bool valid = double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        value *= scale;
        return valid && double.IsFinite(value);
    }

    private static bool TryPaintSvgText(PdfGraphicsBuilder graphics, XElement element,
        IReadOnlyDictionary<FontRequest, SvgTextFont?>? fonts, SvgTransform transform,
        double sourceMinX, double sourceMinY, double imageX, double imageY, double imageWidth, double imageHeight,
        double scaleX, double scaleY, Action<int> admitSamples, CancellationToken cancellationToken)
    {
        if (IsHiddenSvgText(element)) return true;
        if (!TryReadSvgText(element, out SvgTextLabel? label)) return false;
        if (label.Color is not { } color || label.Opacity == 0d || label.Text.Length == 0) return true;
        if (fonts is null || !fonts.TryGetValue(label.Font, out SvgTextFont? prepared) || prepared is null) return false;
        OpenTypeFont font = prepared.Font;
        double unitScale = label.FontSize / font.UnitsPerEm;
        var glyphs = new List<(OpenTypeFont.OpenTypeGlyphOutline? Outline, double X)>();
        double advance = 0d;
        ushort previous = 0;
        foreach (Rune rune in label.Text.EnumerateRunes())
        {
            cancellationToken.ThrowIfCancellationRequested();
            ushort glyph = font.MapCodePoint(rune.Value);
            if (glyph == 0) return false;
            if (label.Kerning && previous != 0) advance += font.GetKerning(previous, glyph) * unitScale;
            OpenTypeFont.OpenTypeGlyphOutline? outline = null;
            if (!Rune.IsWhiteSpace(rune))
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
