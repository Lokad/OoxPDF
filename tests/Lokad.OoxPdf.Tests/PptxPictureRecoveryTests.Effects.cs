using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxPictureRecoveryTests
{
    public static void PptxSvgPictureOpacityComposesWithRootAndNestedGroups()
    {
        foreach (bool nested in new[] { false, true })
        foreach (bool cropped in new[] { false, true })
        foreach (string? local in new string?[] { null, "0.5" })
        {
            var actual = SvgEffectParts("svg-only", new XElement(A + "alphaModFix", new XAttribute("amt", "25000")), local, nested, cropped);
            var expected = SvgEffectParts("svg-only", null, local is null ? "0.25" : "0.125", nested, cropped);
            byte[] actualPdf = ConvertFallback(actual, new RecoveryFontResolver(), out var diagnostics);
            byte[] expectedPdf = ConvertFallback(expected, new RecoveryFontResolver(), out var expectedDiagnostics);
            TestAssert.Equal(0, diagnostics.Count);
            TestAssert.Equal(0, expectedDiagnostics.Count);
            TestAssert.True(actualPdf.AsSpan().SequenceEqual(expectedPdf),
                "Picture opacity composites the whole SVG once, preserving local alpha, overlap, text and crop.");
        }
    }

    public static void PptxSvgPictureOpacityPreservesZeroAndFullEndpoints()
    {
        foreach (int opacity in new[] { 0, 100000 })
        {
            var actual = SvgEffectParts("svg-only", new XElement(A + "alphaModFix", new XAttribute("amt", opacity)));
            var expected = SvgEffectParts("svg-only", null, opacity == 0 ? "0" : null);
            byte[] actualPdf = ConvertFallback(actual, new RecoveryFontResolver(), out var diagnostics);
            byte[] expectedPdf = ConvertFallback(expected, new RecoveryFontResolver(), out var expectedDiagnostics);
            TestAssert.Equal(0, diagnostics.Count);
            TestAssert.Equal(0, expectedDiagnostics.Count);
            TestAssert.True(actualPdf.AsSpan().SequenceEqual(expectedPdf), "Zero/full picture opacity retains the corresponding SVG endpoint.");
        }
    }

    public static void PptxSvgPictureRecolorKeepsAnExplicitOriginalImageFallback()
    {
        string[] effects = ["<a:grayscl/>", "<a:biLevel thresh=\"50000\"/>",
            "<a:lum bright=\"10000\" contrast=\"20000\"/>",
            "<a:duotone><a:srgbClr val=\"000000\"/><a:srgbClr val=\"FFFFFF\"/></a:duotone>"];
        byte[] original = ConvertFallback(SvgEffectParts("svg-only", null), new RecoveryFontResolver(), out _);
        foreach (string effect in effects)
        {
            XElement transform = XElement.Parse($"<root xmlns:a=\"{A}\">{effect}</root>").Elements().Single();
            byte[] svg = ConvertFallback(SvgEffectParts("svg-only", transform), new RecoveryFontResolver(), out var diagnostics);
            TestAssert.True(svg.AsSpan().SequenceEqual(original), "Unsupported vector recolor retains the authored vector paint.");
            var warning = diagnostics.Single();
            TestAssert.Equal("PPTX_UNSUPPORTED_IMAGE_RECOLOR", warning.Id);
            TestAssert.Equal(OoxPdfSeverity.Warning, warning.Severity);
            TestAssert.Equal("Original image", warning.Fallback);

            // A usable authored raster preview remains the preferred source and
            // keeps the existing managed recolor implementation.
            ConvertFallback(SvgEffectParts("valid", new XElement(transform)), new RecoveryFontResolver(), out var previewDiagnostics);
            TestAssert.True(previewDiagnostics.All(d => d.Id != "PPTX_UNSUPPORTED_IMAGE_RECOLOR"),
                "Supported preview recoloring must not inherit the vector fallback diagnostic.");
        }
    }

    private static Dictionary<string, byte[]> SvgEffectParts(string variant, XElement? effect,
        string? rootOpacity = null, bool nested = false, bool cropped = false)
    {
        var parts = FallbackParts(variant, cropped: cropped);
        var slide = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
        if (effect is not null) slide.Descendants(A + "blip").Single().Add(effect);
        parts["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide.ToString());
        string opacity = rootOpacity is null ? "" : $" opacity=\"{rootOpacity}\"";
        string paint = "<rect width=\"70\" height=\"100\" fill=\"red\"/><rect x=\"30\" width=\"70\" height=\"100\" fill=\"red\"/>" +
            "<text x=\"4\" y=\"55\" font-family=\"RecoverySvg\" font-size=\"18\">LABEL</text>";
        if (nested) paint = "<g opacity=\"0.5\">" + paint + "</g>";
        parts["ppt/media/vector.svg"] = TestFixtures.Utf8("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"100\" height=\"100\" viewBox=\"0 0 100 100\"" + opacity + ">" + paint + "</svg>");
        return parts;
    }
}
