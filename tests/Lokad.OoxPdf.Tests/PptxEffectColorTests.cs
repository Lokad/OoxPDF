using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxEffectColorTests
{
    public static void PptxDrawingMlPresetFillsMatchNumericPaint()
    {
        (string Preset, string Hex, string Operator)[] swatches =
        [
            ("DKBlue", "00008B", "0 0 0.545 rg"),
            ("ltBlue", "ADD8E6", "0.678 0.847 0.902 rg"),
            ("medPurple", "9370DB", "0.576 0.439 0.859 rg"),
            ("dkSeaGreen", "8FBC8B", "0.561 0.737 0.545 rg"),
            ("darkSeaGreen", "8FBC8F", "0.561 0.737 0.561 rg"),
            ("ltGoldenrodYellow", "FAFA78", "0.98 0.98 0.471 rg"),
            ("lightGoldenrodYellow", "FAFAD2", "0.98 0.98 0.824 rg")
        ];
        PdfPage numeric = RenderSwatches("srgbClr", swatches.Select(swatch => swatch.Hex));
        PdfPage preset = RenderSwatches("prstClr", swatches.Select(swatch => swatch.Preset));
        foreach (var swatch in swatches)
            TestAssert.Contains(swatch.Operator, numeric.Content);
        TestAssert.True(numeric.ExtGStates.Any(state => state.FillAlpha == .5d), "Swatches retain their authored opacity.");
        TestAssert.Equal(numeric.Content, preset.Content);
        TestAssert.True(numeric.ExtGStates.SequenceEqual(preset.ExtGStates), "Presets retain the same paint state.");

        foreach (string invalid in new[] { "unknown-color", "dkRebeccaPurple", "RebeccaPurple", " dkBlue ", "#00008B", "rgb(0,0,139)" })
            TestAssert.True(!PptxPresetColors.TryResolve(invalid, out _), "Only DrawingML color keywords resolve here: " + invalid);
        TestAssert.True(RgbColor.TryParseCssColor(" RebeccaPurple ", out RgbColor css, out double opacity), "CSS retains its distinct keyword.");
        TestAssert.Equal(new RgbColor(102, 51, 153), css);
        TestAssert.Equal(1d, opacity);
        TestAssert.True(!RgbColor.TryParseCssColor("dkBlue", out _, out _), "DrawingML abbreviations stay scoped to DrawingML.");

        static PdfPage RenderSwatches(string colorElement, IEnumerable<string> values)
        {
            string shapes = string.Concat(values.Select((value, index) => $$"""
                <p:sp><p:nvSpPr><p:cNvPr id="{{index + 2}}" name="Swatch"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="{{914400 + 914400 * index}}" y="914400"/><a:ext cx="457200" cy="457200"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:solidFill><a:{{colorElement}} val="{{value}}"><a:alpha val="50000"/></a:{{colorElement}}></a:solidFill>
                    <a:ln><a:noFill/></a:ln>
                  </p:spPr>
                </p:sp>
                """));
            using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
            {
                ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
                ["_rels/.rels"] = PptxTests.PackageRelationship(),
                ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
                ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
                ["ppt/slides/slide1.xml"] = $$"""
                    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <p:cSld><p:spTree>{{shapes}}</p:spTree></p:cSld>
                    </p:sld>
                    """
            });
            using var package = OoxPackage.Open(input, default);
            var diagnostics = new List<OoxPdfDiagnostic>();
            var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
                .RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default).ToBlockingEnumerable().Single();
            TestAssert.Equal(0, diagnostics.Count);
            return page;
        }
    }

    public static void PptxEffectColorFormsKeepPaintAcrossSupportedOwners()
    {
        string[] colors =
        [
            "<a:srgbClr val=\"FF0000\">{0}</a:srgbClr>",
            "<a:scrgbClr r=\"100000\" g=\"0\" b=\"0\">{0}</a:scrgbClr>",
            "<a:sysClr val=\"windowText\" lastClr=\"FF0000\">{0}</a:sysClr>",
            "<a:prstClr val=\"red\">{0}</a:prstClr>",
            "<a:hslClr hue=\"0\" sat=\"100000\" lum=\"50000\">{0}</a:hslClr>"
        ];
        const string alpha = "<a:alpha val=\"60000\"/><a:alphaMod val=\"50000\"/><a:alphaOff val=\"20000\"/>";
        foreach (string owner in new[] { "shape-shadow", "shape-glow", "picture-shadow" })
        {
            PdfPage baseline = Render(owner, string.Format(System.Globalization.CultureInfo.InvariantCulture, colors[0], alpha));
            PdfImageXObject paint = baseline.Images.Select(resource => resource.Image).Single(image => image.Width > 1);
            byte[] rgb = TestFixtures.InflateZlib(paint.Bytes);
            TestAssert.True(Enumerable.Range(0, rgb.Length / 3).All(i => rgb[i * 3] == 255 && rgb[i * 3 + 1] == 0 && rgb[i * 3 + 2] == 0),
                "The actual effect image retains the independently specified red samples: " + owner);
            TestAssert.True(TestFixtures.InflateZlib(TestAssert.NotNull(paint.Alpha)).Any(value => value > 0), "The effect has visible mask samples.");
            TestAssert.True(baseline.ExtGStates.Any(state => state.FillAlpha == .5d), "Ordered opacity transforms produce half opacity.");
            foreach (string color in colors.Skip(1))
            {
                PdfPage actual = Render(owner, string.Format(System.Globalization.CultureInfo.InvariantCulture, color, alpha));
                TestAssert.True(HasSamePaint(baseline, actual), "Equivalent colors retain image samples and paint state: " + owner);
            }
            PdfPage green = Render(owner, $"<a:srgbClr val=\"00FF00\">{alpha}</a:srgbClr>");
            TestAssert.Equal(baseline.Content, green.Content);
            TestAssert.Equal(baseline.Images.Count, green.Images.Count);
            TestAssert.True(baseline.ExtGStates.SequenceEqual(green.ExtGStates), "The color control retains the same opacity.");
            TestAssert.True(!HasSamePaint(baseline, green), "The paint comparison detects changed pixels with identical drawing commands: " + owner);
        }
    }

    public static void PptxEffectColorFallbacksPreserveMissingPaintAndRecolorPolicy()
    {
        foreach (string owner in new[] { "shape-shadow", "shape-glow", "picture-shadow" })
        {
            PdfPage baseline = Render(owner, null);
            foreach (string color in new[] { "", "<a:prstClr val=\"unknown-color\"/>", "<a:srgbClr val=\"invalid\"/>", "<x:srgbClr xmlns:x=\"urn:ignored\" val=\"FF0000\"/>" })
            {
                PdfPage actual = Render(owner, color);
                TestAssert.True(HasSamePaint(baseline, actual), $"Unsupported effect color retains fallback paint: {owner}/{color}.");
            }
        }
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XElement props = XElement.Parse($$"""
            <a:spPr xmlns:a="{{a}}"><a:effectLst>
              <a:glow rad="38100"><a:prstClr val="red"/></a:glow>
              <a:outerShdw dist="76200"><a:scrgbClr r="100000" g="0" b="0"/></a:outerShdw>
            </a:effectLst></a:spPr>
            """);
        var style = PptxSceneBuilder.ReadChartShapeStyle(props, PptxTheme.Empty, PptxColorMap.Default);
        TestAssert.True(style.Glow.HasGlow && style.OuterShadow.HasShadow, "Chart furniture consumes the same supported effect colors.");
        TestAssert.Equal(new RgbColor(255, 0, 0), style.Glow.Color);
        TestAssert.Equal(new RgbColor(255, 0, 0), style.OuterShadow.Color);
        // This cleanup must not silently extend the separate duotone policy.
        XElement recolor = XElement.Parse($$"""
            <p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="{{a}}">
              <p:blipFill><a:blip><a:duotone><a:prstClr val="red"/><a:prstClr val="white"/></a:duotone></a:blip></p:blipFill>
            </p:pic>
            """);
        TestAssert.Equal(PptxSceneImageRecolorKind.None, PptxSceneBuilder.ReadImageRecolor(recolor, PptxTheme.Empty, PptxColorMap.Default).Kind);
    }

    private static PdfPage Render(string owner, string? color)
    {
        string effect = color is null ? "" : owner == "shape-glow"
            ? $"<a:effectLst><a:glow rad=\"38100\">{color}</a:glow></a:effectLst>"
            : $"<a:effectLst><a:outerShdw dist=\"76200\" dir=\"0\" blurRad=\"12700\">{color}</a:outerShdw></a:effectLst>";
        string geometry = "<a:xfrm><a:off x=\"914400\" y=\"914400\"/><a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"/>";
        string node = owner == "picture-shadow"
            ? $"<p:pic><p:blipFill><a:blip r:embed=\"image\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{geometry}{effect}</p:spPr></p:pic>"
            : $"<p:sp><p:spPr>{geometry}<a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill><a:ln><a:noFill/></a:ln>{effect}</p:spPr></p:sp>";
        var entries = new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes().Replace("</Types>", "<Default Extension=\"png\" ContentType=\"image/png\"/></Types>", StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"image\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/image.png\"/></Relationships>"),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8($"<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><p:cSld><p:spTree>{node}</p:spTree></p:cSld></p:sld>"),
            ["ppt/media/image.png"] = TestFixtures.CreateRgbPng(1, 1, [0, 0, 255])
        };
        using var input = TestFixtures.CreateZipPackage(entries);
        using var package = OoxPackage.Open(input, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        var page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default).ToBlockingEnumerable().Single();
        TestAssert.Equal(0, diagnostics.Count);
        return page;
    }

    private static bool HasSamePaint(PdfPage expected, PdfPage actual)
    {
        return expected.Content == actual.Content &&
            expected.ExtGStates.SequenceEqual(actual.ExtGStates) &&
            expected.Images.Count == actual.Images.Count &&
            expected.Images.Zip(actual.Images).All(pair =>
                pair.First.ResourceName == pair.Second.ResourceName &&
                pair.First.Image.HasIdenticalContent(pair.Second.Image));
    }
}
