using System.Globalization;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class PptxAuthoredColorTests
{
    public static void PptxPercentageTextFormsKeepVisibleBaselinesSpacingAndAutofit()
    {
        PdfPage Render(bool literal)
        {
            string Pct(int units) => literal ? (units / 1000m).ToString(CultureInfo.InvariantCulture) + "%"
                : units.ToString(CultureInfo.InvariantCulture);
            string[] words = ["Baseline", "Spacing", "Autofit", "Bullets"];
            string content = string.Concat(words.Select((word, index) =>
            {
                string autofit = index == 2 ? $"<a:normAutofit fontScale=\"{Pct(80000)}\" lnSpcReduction=\"{Pct(20000)}\"/>" : "<a:noAutofit/>";
                string paragraph = index == 1 ? $"<a:lnSpc><a:spcPct val=\"{Pct(125000)}\"/></a:lnSpc><a:spcBef><a:spcPct val=\"{Pct(50000)}\"/></a:spcBef>"
                    : index == 3 ? $"<a:buSzPct val=\"{Pct(60000)}\"/><a:buFont typeface=\"DejaVu Sans\"/><a:buChar char=\"•\"/>" : "";
                string body = $"<p:txBody><a:bodyPr>{autofit}</a:bodyPr><a:lstStyle/><a:p><a:pPr>{paragraph}</a:pPr><a:r><a:rPr sz=\"2400\" baseline=\"{Pct(index == 0 ? 30000 : 0)}\"><a:latin typeface=\"DejaVu Sans\"/></a:rPr><a:t>{word}</a:t></a:r></a:p></p:txBody>";
                return PercentageShape(index, "<a:noFill/>", body);
            }));
            var page = RenderPercentagePage(content, out var diagnostics);
            TestAssert.Equal(0, diagnostics.Count);
            foreach (string word in words)
                TestAssert.Equal(word.Length, PortableRenderingTests.ReadPaintedSourceLabel(page, word).Count);
            PdfContentValidator.ValidatePage(page, 0, default);
            return page;
        }
        PdfPage integer = Render(false), percent = Render(true);
        TestAssert.Equal(integer.Content, percent.Content);
        TestAssert.Equal(integer.Fonts.Count, percent.Fonts.Count);
    }

    public static void PptxPercentageColorsAndCropsKeepTypedSourceSemantics()
    {
        TestAssert.Equal(new RgbColor(188, 137, 0), ReadScRgbColor("<a:scrgbClr r=\"50%\" g=\"25%\" b=\"0%\"/>").Color);
        TestAssert.Equal(new RgbColor(0, 255, 0), ReadScRgbColor("<a:hslClr hue=\"7200000\" sat=\"100%\" lum=\"50%\"/>").Color);
        TestAssert.Equal(new RgbColor(0, 0, 188), ReadScRgbColor("<a:srgbClr val=\"0000FF\"><a:shade val=\"50%\"/></a:srgbClr>").Color);
        XNamespace a = OoxNamespaces.DrawingNamespace;
        XNamespace p = OoxNamespaces.PresentationNamespace;
        var picture = new XElement(p + "pic", new XElement(p + "blipFill",
            new XElement(a + "blip", new XElement(a + "alphaModFix", new XAttribute("amt", "25%")),
                new XElement(a + "lum", new XAttribute("bright", "-25%"), new XAttribute("contrast", "50%"))),
            new XElement(a + "srcRect", new XAttribute("l", "25%"), new XAttribute("r", "25%")),
            new XElement(a + "stretch", new XElement(a + "fillRect", new XAttribute("t", "12.5%")))));
        var crop = PptxSceneBuilder.ReadPictureCrop(picture);
        TestAssert.Equal(.25d, crop.Left); TestAssert.Equal(.25d, crop.Right);
        TestAssert.Equal(.125d, PptxSceneBuilder.ReadPictureFill(picture).Top);
        TestAssert.Equal(.25d, PptxSceneBuilder.ReadPictureAlpha(picture));
        var recolor = PptxSceneBuilder.ReadImageRecolor(picture, PptxTheme.Empty, PptxColorMap.Default);
        TestAssert.Equal(-.25d, recolor.Brightness); TestAssert.Equal(.5d, recolor.Contrast);
        var gradient = XElement.Parse($"<p:spPr xmlns:p=\"{p}\" xmlns:a=\"{a}\"><a:gradFill><a:gsLst><a:gs pos=\"0%\"><a:srgbClr val=\"FF0000\"/></a:gs><a:gs pos=\"100%\"><a:srgbClr val=\"0000FF\"/></a:gs></a:gsLst><a:lin/></a:gradFill></p:spPr>");
        var fill = PptxSceneBuilder.ReadLinearGradientFill(gradient, PptxTheme.Empty, PptxColorMap.Default);
        TestAssert.True(fill.HasGradient && !fill.HasUnsupportedGradient, "Percent stop offsets are supported source values.");
        TestAssert.Equal(0d, fill.Stops[0].Offset); TestAssert.Equal(1d, fill.Stops[1].Offset);
    }

    public static void PptxColorOpacityCompositionKeepsOrderAndGradientFallbacks()
    {
        foreach (var (transforms, expected) in new[]
        {
            ("", 1d), ("<a:alpha val=\"0%\"/>", 0d),
            ("<a:alpha val=\"50000\"/><a:alphaMod val=\"50000\"/>", .25d),
            ("<a:alpha val=\"50%\"/><a:alphaMod val=\"50%\"/>", .25d),
            ("<a:alphaMod val=\"50%\"/><a:alpha val=\"50%\"/>", .5d),
            ("<a:alpha val=\"50%\"/><a:alphaOff val=\"-50%\"/>", 0d),
            ("<a:alpha val=\"0%\"/><a:alphaOff val=\"50%\"/>", .5d),
            ("<a:alpha val=\"200%\"/><a:alphaMod val=\"50%\"/>", .5d),
            ("<a:alpha val=\"50%\"/><a:alphaOff val=\"100%\"/><a:alphaMod val=\"50%\"/>", .5d),
            ("<a:alpha val=\"50%\"/><a:alphaMod val=\"300%\"/><a:alphaOff val=\"-50%\"/>", .5d),
            ("<a:alpha val=\"50%\"/><a:alphaMod val=\"bad\"/><a:alphaOff val=\"bad\"/>", .5d),
            ("<a:alpha val=\"bad\"/>", 1d)
        })
        {
            var color = ReadScRgbColor("<a:srgbClr val=\"FF0000\">" + transforms + "</a:srgbClr>");
            TestAssert.Equal(new RgbColor(255, 0, 0), color.Color);
            TestAssert.Equal(expected, color.Alpha);
        }
        string quarter = "<a:alpha val=\"50%\"/><a:alphaMod val=\"50%\"/>";
        string gradient(bool uniform) => "<a:gradFill><a:gsLst><a:gs pos=\"0%\"><a:srgbClr val=\"FF0000\">" + quarter +
            "</a:srgbClr></a:gs><a:gs pos=\"100%\"><a:srgbClr val=\"0000FF\">" +
            (uniform ? "<a:alpha val=\"25%\"/>" : "") + "</a:srgbClr></a:gs></a:gsLst><a:lin/></a:gradFill>";
        string solid = "<a:solidFill><a:srgbClr val=\"FF0000\">" + quarter + "</a:srgbClr></a:solidFill>";
        var page = RenderPercentagePage(PercentageShape(0, solid) + PercentageShape(1, gradient(true)) + PercentageShape(2, gradient(false)), out var diagnostics);
        TestAssert.True(page.ExtGStates.Any(state => state.FillAlpha == .25d), "Effective quarter opacity must reach PDF painting.");
        TestAssert.Equal(1, diagnostics.Count(item => item.Id == "PPTX_UNSUPPORTED_GRADIENT_FILL"));
        TestAssert.Equal(1, diagnostics.Count(item => item.Id == "PPTX_UNSUPPORTED_TRANSPARENCY"));
        PdfContentValidator.ValidatePage(page, 0, default);
        XNamespace a = OoxNamespaces.DrawingNamespace;
        var stops = new[] { quarter, "<a:alpha val=\"25%\"/>" }.Select(alpha => XElement.Parse(
            $"<a:gs xmlns:a=\"{a}\"><a:srgbClr val=\"FF0000\">{alpha}</a:srgbClr></a:gs>")).ToArray();
        TestAssert.True(PptxColorResolver.HasSupportedGradientStopAlpha(stops), "Equivalent effective stop alpha shares the existing uniform-gradient rule.");
        XNamespace p = OoxNamespaces.PresentationNamespace;
        XElement ReadGradient(bool uniform) => XElement.Parse(gradient(uniform).Replace("<a:gradFill>", $"<a:gradFill xmlns:a=\"{a}\">", StringComparison.Ordinal));
        var mixedBackground = new XElement(p + "bgPr", ReadGradient(false));
        var owners = new XElement(p + "cSld",
            new XElement(p + "spPr", ReadGradient(true)),
            new XElement(p + "bgPr", ReadGradient(true)),
            mixedBackground);
        TestAssert.Equal(false, PptxSceneBuilder.HasUnsupportedAlpha(owners,
            color => !color.Ancestors().Contains(mixedBackground)));
        TestAssert.True(PptxSceneBuilder.HasUnsupportedAlpha(owners),
            "A following mixed background cannot inherit a preceding gradient's support.");
        var unsupported = XElement.Parse($"<p:spPr xmlns:p=\"{OoxNamespaces.PresentationNamespace}\" xmlns:a=\"{a}\"><a:pattFill><a:fgClr><a:srgbClr val=\"FF0000\"><a:alphaMod val=\"50%\"/></a:srgbClr></a:fgClr></a:pattFill></p:spPr>");
        TestAssert.True(PptxSceneBuilder.HasUnsupportedAlpha(unsupported), "Modulation alone must not hide unsupported pattern transparency.");
    }

    private static string PercentageShape(int index, string fill, string body = "") =>
        $"<p:sp><p:nvSpPr><p:cNvPr id=\"{index + 2}\" name=\"Percentage {index}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"{(72 + index % 2 * 288) * 12700}\" y=\"{(72 + index / 2 * 120) * 12700}\"/><a:ext cx=\"3048000\" cy=\"1219200\"/></a:xfrm><a:prstGeom prst=\"rect\"/>{fill}<a:ln><a:noFill/></a:ln></p:spPr>{body}</p:sp>";

    private static PdfPage RenderPercentagePage(string content, out List<OoxPdfDiagnostic> diagnostics)
    {
        using var input = TestFixtures.CreateZipPackage(new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = PptxTests.BasicContentTypes(),
            ["_rels/.rels"] = PptxTests.PackageRelationship(),
            ["ppt/_rels/presentation.xml.rels"] = PptxTests.PresentationRelationship(),
            ["ppt/presentation.xml"] = PptxTests.BasicPresentation(),
            ["ppt/slides/slide1.xml"] = $"<p:sld xmlns:p=\"{OoxNamespaces.PresentationNamespace}\" xmlns:a=\"{OoxNamespaces.DrawingNamespace}\"><p:cSld><p:spTree>{content}</p:spTree></p:cSld></p:sld>"
        });
        using var package = OoxPackage.Open(input, default);
        diagnostics = new();
        return new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default)
            .ToBlockingEnumerable().Single();
    }
}
