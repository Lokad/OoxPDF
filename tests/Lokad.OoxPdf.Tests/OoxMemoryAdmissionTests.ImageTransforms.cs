using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetImageTransformsKeepSourcePixelsAndRefundTemporaryPlanes()
    {
        foreach (string kind in new[] { "none", "gray", "luminance", "bilevel", "duotone", "crop", "crop-gray", "shadow", "glow" })
        {
            using var input = ImageTransformPackage(kind);
            using var package = OoxPackage.Open(input, default);
            var document = new PptxReader().Read(package, default);
            byte[] source = TestAssert.NotNull(package.GetPart("/ppt/media/image1.png")).Bytes;
            byte[] original = (byte[])source.Clone();
            var renderer = new PptxRenderer(new ProgramResolver(new MemoryFontProgramSource("transform-font", TestFontBuilder.CreateTestFont())));
            string? first = null;
            for (int pass = 0; pass < 2; pass++)
            {
                int admitted = 0;
                var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                    if (after == TransformPlaneBytes(kind)) admitted++;
                    return false;
                }};
                using (var conversion = OoxConversionBudget.BeginScope(null, pool))
                {
                    var diagnostics = new List<OoxPdfDiagnostic>();
                    PdfPage page = renderer.RenderPages(document, package, diagnostics.Add, default).First();
                    TestAssert.Equal(0, diagnostics.Count);
                    TestAssert.Equal(0, page.Fonts.Count);
                    PdfImageXObject image = page.Images.Single().Image;
                    TestAssert.Equal((long)image.Bytes.Length + (image.Alpha?.Length ?? 0), pool.Reserved);
                    if (TransformPlaneBytes(kind) > 0) TestAssert.True(admitted > 0, "The actual crop/effect plane is admitted: " + kind);
                    if (kind is not ("shadow" or "glow"))
                    {
                        byte[] expected = kind switch {
                            "gray" => [54, 54, 54, 18, 18, 18],
                            "luminance" => [255, 26, 26, 26, 26, 255],
                            "bilevel" => [0, 0, 0, 0, 0, 0],
                            "duotone" => [54, 0, 0, 18, 0, 0],
                            "crop" => [0, 0, 255],
                            "crop-gray" => [18, 18, 18],
                            _ => [255, 0, 0, 0, 0, 255]
                        };
                        TestAssert.True(InflateTransform(image.Bytes).SequenceEqual(expected), "Recolor/crop pixels are unchanged: " + kind);
                        TestAssert.True(InflateTransform(TestAssert.NotNull(image.Alpha)).SequenceEqual(kind.StartsWith("crop", StringComparison.Ordinal) ? new byte[] { 255 } : [128, 255]), "Source/cropped alpha remains exact.");
                    }
                    string fingerprint = image.ResourceKey + page.Content;
                    if (first is null) first = fingerprint;
                    else TestAssert.Equal(first, fingerprint);
                    TestAssert.True(source.SequenceEqual(original), "Repeated rendering never recolors the encoded source/cache.");
                }
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.Equal(0, pool.CallerDisposals);
            }
        }
    }

    public static void MemoryBudgetImageTransformsRefusalAndCancellationDoNotRecoverOrLeak()
    {
        foreach (string kind in new[] { "crop", "crop-gray", "shadow", "glow" })
        foreach (bool cancel in new[] { false, true })
        {
            using var input = ImageTransformPackage(kind);
            using var package = OoxPackage.Open(input, default);
            var document = new PptxReader().Read(package, default);
            using var cancellation = new CancellationTokenSource();
            int attempts = 0;
            var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                if (after != TransformPlaneBytes(kind)) return false;
                attempts++;
                if (cancel) { cancellation.Cancel(); return false; }
                return true;
            }};
            var diagnostics = new List<OoxPdfDiagnostic>();
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                var renderer = new PptxRenderer(new ProgramResolver(new MemoryFontProgramSource("transform-font", TestFontBuilder.CreateTestFont())));
                if (cancel)
                    TestAssert.Throws<OperationCanceledException>(() => renderer.RenderPages(document, package, diagnostics.Add, cancellation.Token).First());
                else
                {
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => renderer.RenderPages(document, package, diagnostics.Add, default).First());
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                }
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(1, attempts);
                TestAssert.True(diagnostics.All(d => d.Id is not "IMAGE_CROP_UNSUPPORTED_FORMAT" and not "PPTX_NODE_RENDER_FAILED"), "Refusal/cancellation cannot become clipping or node omission.");
            }
            TestAssert.Equal(0L, pool.Reserved);
        }
    }

    public static void MemoryBudgetImageTransformsPreservePublicationAndCallerStreams()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-transform-admission-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (string kind in new[] { "crop", "glow" })
            {
                using var package = ImageTransformPackage(kind);
                string input = Path.Combine(directory, kind + ".pptx");
                File.WriteAllBytes(input, package.ToArray());
                foreach (bool cancel in new[] { false, true })
                foreach (bool file in new[] { false, true })
                {
                    using var cancellation = new CancellationTokenSource();
                    int attempts = 0;
                    var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                        if (after != TransformPlaneBytes(kind)) return false;
                        attempts++;
                        if (cancel) { cancellation.Cancel(); return false; }
                        return true;
                    }};
                    var options = new OoxPdfOptions { InputKind=OoxPdfInputKind.Pptx, MemoryBudget=pool,
                        FontResolver=new ProgramResolver(new MemoryFontProgramSource("transform-font", TestFontBuilder.CreateTestFont())) };
                    if (file)
                    {
                        string output = Path.Combine(directory, "result.pdf");
                        File.WriteAllBytes(output, [7, 8, 9]);
                        Action convert = () => OoxPdfConverter.Convert(input, output, options, cancellation.Token);
                        if (cancel) TestAssert.Throws<OperationCanceledException>(convert);
                        else TestAssert.Throws<OoxPdfLimitExceededException>(convert);
                        TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Refusal/cancellation preserves the previous destination.");
                        TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                    }
                    else
                    {
                        using var source = File.OpenRead(input);
                        using var output = new MemoryStream();
                        Action convert = () => OoxPdfConverter.Convert(source, output, options, cancellation.Token);
                        if (cancel) TestAssert.Throws<OperationCanceledException>(convert);
                        else TestAssert.Throws<OoxPdfLimitExceededException>(convert);
                        TestAssert.Equal(0L, output.Length);
                        TestAssert.True(source.CanRead && output.CanWrite, "Caller streams stay open.");
                    }
                    TestAssert.Equal(1, attempts);
                    TestAssert.Equal(0L, pool.Reserved);
                    TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                    TestAssert.Equal(0, pool.CallerDisposals);
                }
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static long TransformPlaneBytes(string kind) => kind switch {
        "crop" or "crop-gray" => 4,
        "shadow" => 97L * 97 * 4,
        "glow" => 92L * 92 * 4,
        _ => 0
    };

    private static byte[] InflateTransform(byte[] bytes)
    {
        using var source = new MemoryStream(bytes);
        using var zlib = new ZLibStream(source, CompressionMode.Decompress);
        using var output = new MemoryStream();
        zlib.CopyTo(output);
        return output.ToArray();
    }

    private static MemoryStream ImageTransformPackage(string kind)
    {
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var parts = new Dictionary<string, byte[]>();
        using (var archive = ZipFile.OpenRead(Case("PictureGeometry/picture-geometry-controls.pptx")))
        foreach (var entry in archive.Entries)
        {
            using var stream = entry.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            parts[entry.FullName] = buffer.ToArray();
        }
        var slide = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
        var picture = new XElement(slide.Descendants(p + "pic").First());
        foreach (string name in parts.Keys.Where(name => name.EndsWith(".xml", StringComparison.Ordinal) &&
            (name.StartsWith("ppt/slides/", StringComparison.Ordinal) || name.StartsWith("ppt/slideLayouts/", StringComparison.Ordinal) || name.StartsWith("ppt/slideMasters/", StringComparison.Ordinal))).ToArray())
        {
            var xml = XDocument.Parse(Encoding.UTF8.GetString(parts[name]));
            foreach (var tree in xml.Descendants(p + "spTree")) tree.Elements().Where(element => element.Name.LocalName is "pic" or "sp" or "graphicFrame" or "grpSp" or "cxnSp").Remove();
            parts[name] = TestFixtures.Utf8(xml.ToString());
        }
        slide = XDocument.Parse(Encoding.UTF8.GetString(parts["ppt/slides/slide1.xml"]));
        if (kind is "shadow" or "glow")
        {
            string effect = kind == "shadow"
                ? "<a:outerShdw blurRad=\"127000\" dist=\"25400\" dir=\"2700000\"><a:srgbClr val=\"112233\"><a:alpha val=\"75000\"/></a:srgbClr></a:outerShdw>"
                : "<a:glow rad=\"127000\"><a:srgbClr val=\"112233\"><a:alpha val=\"75000\"/></a:srgbClr></a:glow>";
            slide.Descendants(p + "spTree").Single().Add(XElement.Parse($"""
                <p:sp xmlns:p="{p}" xmlns:a="{a}"><p:nvSpPr><p:cNvPr id="9000" name="Public effect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
                  <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="336699"/></a:solidFill>
                    <a:effectLst>{effect}</a:effectLst></p:spPr></p:sp>
                """));
        }
        else
        {
            var blip = picture.Descendants(a + "blip").Single();
            blip.RemoveNodes();
            blip.Parent!.Elements(a + "srcRect").Remove();
            if (kind.StartsWith("crop", StringComparison.Ordinal)) blip.AddAfterSelf(new XElement(a + "srcRect", new XAttribute("l", "50000")));
            if (kind is "gray" or "crop-gray") blip.Add(new XElement(a + "grayscl"));
            if (kind == "luminance") blip.Add(new XElement(a + "lum", new XAttribute("bright", "10000"), new XAttribute("contrast", "20000")));
            if (kind == "bilevel") blip.Add(new XElement(a + "biLevel", new XAttribute("thresh", "50000")));
            if (kind == "duotone") blip.Add(new XElement(a + "duotone", new XElement(a + "srgbClr", new XAttribute("val", "000000")), new XElement(a + "srgbClr", new XAttribute("val", "FF0000"))));
            slide.Descendants(p + "spTree").Single().Add(picture);
        }
        parts["ppt/slides/slide1.xml"] = TestFixtures.Utf8(slide.ToString());
        parts["ppt/media/image1.png"] = TestFixtures.CreateRgbaPng(2, 1, [255, 0, 0, 128, 0, 0, 255, 255]);
        return TestFixtures.CreateZipPackage(parts);
    }
}
