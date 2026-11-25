using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static class PptxPictureRecolorTests
{
    public static void PptxPictureRecolorKeepsPixelsAcrossPictureAndFillOwners()
    {
        (string Effect, byte[] Rgb)[] treatments =
        [
            ("", [255, 0, 0, 0, 0, 255]),
            ("<a:grayscl/>", [54, 54, 54, 18, 18, 18]),
            ("<a:biLevel thresh=\"50000\"/>", [0, 0, 0, 0, 0, 0]),
            ("<a:duotone><a:srgbClr val=\"000000\"/><a:srgbClr val=\"FFFFFF\"/></a:duotone>", [54, 54, 54, 18, 18, 18]),
            ("<a:lum bright=\"100000\"/>", [255, 255, 255, 255, 255, 255]),
            ("<a:grayscl/><a:alphaModFix amt=\"50000\"/>", [54, 54, 54, 18, 18, 18]),
            ("<a:biLevel thresh=\"50%\"/>", [0, 0, 0, 0, 0, 0]),
            ("<a:lum bright=\"100%\"/>", [255, 255, 255, 255, 255, 255]),
            ("<a:lum bright=\"0%\" contrast=\"0%\"/>", [255, 0, 0, 0, 0, 255]),
            ("<a:grayscl/><a:alphaModFix amt=\"50%\"/>", [54, 54, 54, 18, 18, 18]),
            ("<a:duotone><a:prstClr val=\"red\"/><a:prstClr val=\"white\"/></a:duotone>", [255, 54, 54, 255, 18, 18]),
            ("<a:duotone><a:prstClr val=\"black\"/><a:prstClr val=\"white\"><a:shade val=\"50000\"/></a:prstClr></a:duotone>",
                [40, 40, 40, 14, 14, 14])
        ];
        var nodes = new System.Text.StringBuilder();
        foreach (bool fill in new[] { false, true })
        for (int i = 0; i < treatments.Length; i++)
        {
            int id = 2 + (fill ? treatments.Length : 0) + i;
            string geometry = $"<a:xfrm><a:off x=\"{(10 + i % 6 * 110) * 12700}\" y=\"{(72 + (i / 6 + (fill ? 2 : 0)) * 88) * 12700}\"/><a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"/>";
            string blip = $"<a:blip r:embed=\"image\">{treatments[i].Effect}</a:blip><a:stretch><a:fillRect/></a:stretch>";
            nodes.Append(fill
                ? $"<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"Fill {i}\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>{geometry}<a:blipFill>{blip}</a:blipFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>"
                : $"<p:pic><p:nvPicPr><p:cNvPr id=\"{id}\" name=\"Picture {i}\"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>{blip}</p:blipFill><p:spPr>{geometry}</p:spPr></p:pic>");
        }
        var entries = new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = TestFixtures.Utf8(PptxTests.BasicContentTypes().Replace("</Types>", "<Default Extension=\"png\" ContentType=\"image/png\"/></Types>", StringComparison.Ordinal)),
            ["_rels/.rels"] = TestFixtures.Utf8(PptxTests.PackageRelationship()),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8(PptxTests.PresentationRelationship()),
            ["ppt/presentation.xml"] = TestFixtures.Utf8(PptxTests.BasicPresentation()),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image.png"/></Relationships>
                """),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8($$"""
                <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree>{{nodes}}</p:spTree></p:cSld></p:sld>
                """),
            ["ppt/media/image.png"] = TestFixtures.CreateRgbaPng(2, 1, [255, 0, 0, 255, 0, 0, 255, 128])
        };
        using var input = TestFixtures.CreateZipPackage(entries);
        using var package = OoxPackage.Open(input, default);
        var diagnostics = new List<OoxPdfDiagnostic>();
        PdfPage page = new PptxRenderer(new PortableRenderingTests.PinnedResolver())
            .RenderPagesAsync(new PptxReader().Read(package, default), package, diagnostics.Add, default)
            .ToBlockingEnumerable().Single();
        TestAssert.Equal(0, diagnostics.Count);
        TestAssert.Equal(2 * treatments.Length, page.Images.Count);
        for (int i = 0; i < page.Images.Count; i++)
        {
            var image = page.Images[i].Image;
            TestAssert.Equal(2, image.Width);
            TestAssert.Equal(1, image.Height);
            TestAssert.True(TestFixtures.InflateZlib(image.Bytes).AsSpan().SequenceEqual(treatments[i % treatments.Length].Rgb),
                $"Image {i} must retain the pixels for treatment {treatments[i % treatments.Length].Effect}.");
            TestAssert.True(TestFixtures.InflateZlib(TestAssert.NotNull(image.Alpha)).AsSpan().SequenceEqual(new byte[] { 255, 128 }),
                "Recoloring retains the image's authored alpha plane.");
        }
        TestAssert.True(page.ExtGStates.Any(state => state.FillAlpha == .5d), "Picture alpha remains a separate paint state.");
    }
}
