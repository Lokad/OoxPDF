using System.Text;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static class ImageRecoveryTests
{
    public static void BmpHeaderErrorsUseImageRecoveryAcrossContainers()
    {
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (bool pptx in new[] { false, true })
        foreach (int dibSize in new[] { 40, int.MaxValue })
        {
            byte[] bmp = TestFixtures.CreateRgbBmp(1, 1, [10, 20, 30]);
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(14, 4), dibSize);
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(10, 4), -1);
            using var input = TestFixtures.CreateZipPackage(PictureEntries(pptx, "", bmp, "bmp"));
            using var output = new MemoryStream();
            var diagnostics = new List<OoxPdfDiagnostic>();
            OoxPdfConverter.Convert(input, output, new OoxPdfOptions {
                InputKind = pptx ? OoxPdfInputKind.Pptx : OoxPdfInputKind.Docx,
                FontResolver = resolver, DiagnosticSink = diagnostics.Add
            });
            string pdf = Encoding.Latin1.GetString(output.ToArray());
            TestAssert.Contains("%PDF-", pdf);
            TestAssert.DoesNotContain("/Subtype /Image", pdf);
            TestAssert.Equal(1, diagnostics.Count(d => d.Id == "IMAGE_UNSUPPORTED_FORMAT" && d.Severity == OoxPdfSeverity.Error));
            TestAssert.True(diagnostics.All(d => d.Id != "PPTX_NODE_RENDER_FAILED"), "BMP errors must use image recovery.");
        }
    }

    public static void PngTransparentKeysAndEquivalentRgbaProduceIdenticalPdfAcrossContainersAndEffects()
    {
        var pairs = new[]
        {
            (Keyed: ImagingTests.BuildStoredPng(2, 1, 8, 2, 0, [0, 25, 50, 200, 0, 255, 0], transparency: [0, 0, 0, 255, 0, 0]),
             Rgba: ImagingTests.BuildStoredPng(2, 1, 8, 6, 0, [0, 25, 50, 200, 255, 0, 255, 0, 0])),
            (Keyed: ImagingTests.BuildStoredPng(2, 1, 1, 0, 0, [0, 0x40], transparency: [0, 1]),
             Rgba: ImagingTests.BuildStoredPng(2, 1, 8, 6, 0, [0, 0, 0, 0, 255, 255, 255, 255, 0]))
        };
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var (pptx, effect) in new[] { (false, ""), (true, ""), (true, "crop"), (true, "recolor") })
        foreach (var pair in pairs)
        {
            byte[] Convert(byte[] image)
            {
                using var input = TestFixtures.CreateZipPackage(PictureEntries(pptx, effect, image));
                using var output = new MemoryStream();
                var diagnostics = new List<OoxPdfDiagnostic>();
                OoxPdfConverter.Convert(input, output, new OoxPdfOptions
                {
                    InputKind = pptx ? OoxPdfInputKind.Pptx : OoxPdfInputKind.Docx,
                    FontResolver = resolver, DiagnosticSink = diagnostics.Add
                });
                TestAssert.Equal(0, diagnostics.Count);
                return output.ToArray();
            }
            TestAssert.True(Convert(pair.Keyed).SequenceEqual(Convert(pair.Rgba)),
                $"Equivalent PNG transparency must preserve the PDF: PPTX={pptx}, effect={effect}.");
        }
    }

    public static void PngTruncationUsesImageErrorsAcrossContainersAndEffects()
    {
        byte[][] images =
        [
            ImagingTests.BuildStoredPng(2, 2, 8, 2, 0, [0, 10, 20, 30, 40, 50, 60, 0, 70, 80, 90, 100, 110, 120]),
            ImagingTests.BuildStoredPng(2, 2, 8, 2, 0, [0, 10, 20, 30, 40, 50, 60]),
            ImagingTests.BuildStoredPng(2, 2, 8, 2, 0, [0, 10, 20, 30, 40, 50, 60, 0, 70, 80, 90]),
            ImagingTests.BuildStoredPng(1, 1, 8, 6, 1, [0, 10, 20, 30]),
            ImagingTests.BuildStoredPng(1, 1, 8, 2, 0, [], header: []),
            ImagingTests.BuildStoredPng(1, 1, 8, 2, 0, [], header: [0, 0, 0, 1, 0, 0, 0, 1, 8]),
            ImagingTests.BuildStoredPng(1, 1, 8, 2, 0, [0, 10, 20, 30], transparency: [0, 0, 0, 0, 0])
        ];
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (var (pptx, effect) in new[] { (false, ""), (true, ""), (true, "crop"), (true, "recolor") })
        {
            for (int i = 0; i < images.Length; i++)
            {
                using var input = TestFixtures.CreateZipPackage(PictureEntries(pptx, effect, images[i]));
                using var output = new MemoryStream();
                var diagnostics = new List<OoxPdfDiagnostic>();
                OoxPdfConverter.Convert(input, output, new OoxPdfOptions
                {
                    InputKind = pptx ? OoxPdfInputKind.Pptx : OoxPdfInputKind.Docx,
                    FontResolver = resolver, DiagnosticSink = diagnostics.Add
                });
                string pdf = Encoding.Latin1.GetString(output.ToArray());
                TestAssert.Contains("%PDF-", pdf);
                TestAssert.Equal(i == 0 ? 0 : 1, diagnostics.Count(d => d.Id == "IMAGE_UNSUPPORTED_FORMAT" && d.Severity == OoxPdfSeverity.Error));
                TestAssert.True(diagnostics.All(d => d.Id != "PPTX_NODE_RENDER_FAILED"), "Image failures belong to image recovery, not generic node recovery.");
                if (i == 0) TestAssert.Contains("/Subtype /Image", pdf);
                else TestAssert.DoesNotContain("/Subtype /Image", pdf);
            }
        }
    }

    public static void PngImageErrorsPreserveFilePublicationOnObserverFailureAndCancellation()
    {
        byte[][] malformed =
        [
            ImagingTests.BuildStoredPng(2, 2, 8, 2, 0, [0, 10, 20, 30, 40, 50, 60]),
            ImagingTests.BuildStoredPng(1, 1, 8, 2, 0, [], header: [])
        ];
        var resolver = new PortableRenderingTests.PinnedResolver();
        foreach (byte[] bad in malformed)
        {
            string input = TestFixtures.WriteTempPackage(".docx", PictureEntries(false, "", bad));
            foreach (bool cancel in new[] { false, true })
            {
                string output = Path.ChangeExtension(input, cancel ? ".cancel.pdf" : ".observer.pdf");
                File.WriteAllText(output, "previous PDF sentinel");
                using var cancellation = new CancellationTokenSource();
                var failure = new InvalidOperationException("Image observer failure.");
                var options = new OoxPdfOptions { FontResolver = resolver, DiagnosticSink = diagnostic =>
                {
                    if (diagnostic.Id != "IMAGE_UNSUPPORTED_FORMAT") return;
                    if (cancel) cancellation.Cancel();
                    else throw failure;
                } };
                if (cancel) TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(input, output, options, cancellation.Token));
                else TestAssert.True(ReferenceEquals(failure, TestAssert.Throws<InvalidOperationException>(() => OoxPdfConverter.Convert(input, output, options))), "Preserve the original observer exception.");
                TestAssert.Equal("previous PDF sentinel", File.ReadAllText(output));
                TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);
            }
        }
    }

    private static Dictionary<string, byte[]> PictureEntries(bool pptx, string effect, byte[] image, string extension = "png")
    {
        var entries = pptx ? JpegPassthroughTests.PptxEntries() : JpegPassthroughTests.DocxEntries();
        foreach (string name in entries.Keys.ToArray())
        {
            string xml = Encoding.UTF8.GetString(entries[name]).Replace("image/jpeg", "image/" + extension)
                .Replace("Extension=\"jpg\"", "Extension=\"" + extension + "\"").Replace("probe.jpg", "probe." + extension);
            if (name == "ppt/slides/slide1.xml")
            {
                if (effect == "crop") xml = xml.Replace("<a:stretch>", "<a:srcRect l=\"25000\" r=\"25000\"/><a:stretch>");
                if (effect == "recolor") xml = xml.Replace("<a:blip r:embed=\"image\"/>", "<a:blip r:embed=\"image\"><a:duotone><a:srgbClr val=\"000000\"/><a:srgbClr val=\"FFFFFF\"/></a:duotone></a:blip>");
            }
            entries[name] = TestFixtures.Utf8(xml);
        }
        entries[(pptx ? "ppt" : "word") + "/media/probe." + extension] = image;
        return entries;
    }
}
