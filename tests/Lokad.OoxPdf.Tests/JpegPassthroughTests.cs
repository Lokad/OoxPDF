using System.Text;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;

namespace Lokad.OoxPdf.Tests;

internal static class JpegPassthroughTests
{
    // Existing 2x1 RGB baseline fixture and equivalent grayscale, CMYK and
    // progressive encodings. Fixtures only; no image library is needed at runtime.
    private static readonly string[] EncodedControls =
    [
        "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"
    ];

    public static void JpegPassthroughKeepsSupportedFramesAndBytes()
    {
        string[] colorSpaces = ["/DeviceRGB", "/DeviceGray", "/DeviceCMYK", "/DeviceRGB"];
        for (int i = 0; i < EncodedControls.Length; i++)
        {
            byte[] jpeg = Convert.FromBase64String(EncodedControls[i]);
            var image = OoxImageDecoder.Decode("image/jpeg", jpeg, static rgb => rgb);
            TestAssert.Equal(2, image.Width);
            TestAssert.Equal(1, image.Height);
            TestAssert.Equal(8, image.BitsPerComponent);
            TestAssert.Equal(colorSpaces[i], image.ColorSpace);
            TestAssert.True(ReferenceEquals(jpeg, image.Bytes), "Supported JPEGs must retain their encoded bytes.");
        }
    }

    public static void JpegPassthroughRejectsMalformedFrameMetadata()
    {
        foreach (byte[] jpeg in DamagedFrames())
        {
            TestAssert.Throws<InvalidDataException>(() => OoxImageDecoder.Decode("image/jpeg", jpeg, static rgb => rgb));
        }
        byte[] original = Convert.FromBase64String(EncodedControls[0]);
        int frame = original.AsSpan().IndexOf(new byte[] { 0xFF, 0xC0 });
        foreach (byte precision in new byte[] { 0, 12, 16 })
        {
            byte[] jpeg = (byte[])original.Clone();
            jpeg[frame + 4] = precision;
            TestAssert.Throws<NotSupportedException>(() => OoxImageDecoder.Decode("image/jpeg", jpeg, static rgb => rgb));
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => OoxImageDecoder.Decode("image/jpeg", original, static rgb => rgb, cancelled.Token));
    }

    public static void DocxMalformedJpegReportsErrorWithoutInvalidImageDictionary()
    {
        CheckPublicConversion(pptx: false, recolor: false);
    }

    public static void PptxMalformedJpegReportsErrorIncludingRecolorFallback()
    {
        CheckPublicConversion(pptx: true, recolor: false);
        CheckPublicConversion(pptx: true, recolor: true);
    }

    private static IEnumerable<byte[]> DamagedFrames()
    {
        byte[] original = Convert.FromBase64String(EncodedControls[0]);
        int frame = original.AsSpan().IndexOf(new byte[] { 0xFF, 0xC0 });
        foreach (int offset in new[] { 5, 7 })
        {
            byte[] jpeg = (byte[])original.Clone();
            jpeg[frame + offset] = jpeg[frame + offset + 1] = 0;
            yield return jpeg;
        }
        byte[] components = (byte[])original.Clone();
        components[frame + 9] = 2; // The unchanged table describes three components.
        yield return components;
        yield return original[..(frame + 12)]; // Incomplete SOF component table.
    }

    private static void CheckPublicConversion(bool pptx, bool recolor)
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        byte[] control = Convert.FromBase64String(EncodedControls[0]);
        byte[] badPrecision = (byte[])control.Clone();
        badPrecision[control.AsSpan().IndexOf(new byte[] { 0xFF, 0xC0 }) + 4] = 0;
        foreach (byte[] jpeg in new[] { control }.Concat(DamagedFrames()).Append(badPrecision))
        {
            bool valid = ReferenceEquals(control, jpeg);
            Dictionary<string, byte[]> entries = pptx ? PptxEntries() : DocxEntries();
            entries[pptx ? "ppt/media/probe.jpg" : "word/media/probe.jpg"] = jpeg;
            if (recolor)
            {
                string xml = Encoding.UTF8.GetString(entries["ppt/slides/slide1.xml"]);
                entries["ppt/slides/slide1.xml"] = TestFixtures.Utf8(xml.Replace("<a:blip r:embed=\"image\"/>",
                    "<a:blip r:embed=\"image\"><a:duotone><a:srgbClr val=\"000000\"/><a:srgbClr val=\"FFFFFF\"/></a:duotone></a:blip>"));
            }
            using var input = TestFixtures.CreateZipPackage(entries);
            using var output = new MemoryStream();
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions { FontResolver = resolver, DiagnosticSink = diagnostics.Add });
            string pdf = Encoding.Latin1.GetString(output.ToArray());
            TestAssert.Contains("%PDF-", pdf);
            TestAssert.Equal(valid ? 0 : 1, diagnostics.Count(d => d.Id == "IMAGE_UNSUPPORTED_FORMAT"));
            if (valid) TestAssert.Contains("/Subtype /Image", pdf);
            else
            {
                TestAssert.True(diagnostics.Any(d => d.Severity == OoxPdfSeverity.Error && d.Fallback == "Ignored"), "Malformed JPEG must be diagnosed as an omitted image.");
                TestAssert.DoesNotContain("/Subtype /Image", pdf);
            }
        }
    }

    private static Dictionary<string, byte[]> PptxEntries() => new()
    {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"""),
            ["_rels/.rels"] = TestFixtures.Utf8("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"""),
            ["ppt/presentation.xml"] = TestFixtures.Utf8("""<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldSz cx="9144000" cy="6858000"/><p:sldIdLst><p:sldId id="256" r:id="r1"/></p:sldIdLst></p:presentation>"""),
            ["ppt/_rels/presentation.xml.rels"] = TestFixtures.Utf8("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"""),
            ["ppt/slides/slide1.xml"] = TestFixtures.Utf8("""<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="2" name="Probe"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="2743200" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"""),
            ["ppt/slides/_rels/slide1.xml.rels"] = TestFixtures.Utf8("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/probe.jpg"/></Relationships>""")
    };

    private static Dictionary<string, byte[]> DocxEntries() => new()
    {
        ["[Content_Types].xml"] = TestFixtures.Utf8("""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""),
        ["_rels/.rels"] = TestFixtures.Utf8("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""),
        ["word/_rels/document.xml.rels"] = TestFixtures.Utf8("""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/probe.jpg"/></Relationships>"""),
        ["word/document.xml"] = TestFixtures.Utf8("""<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="image"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>""")
    };
}
