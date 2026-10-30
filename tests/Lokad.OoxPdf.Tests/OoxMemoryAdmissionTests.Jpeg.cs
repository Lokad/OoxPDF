using System.Buffers.Binary;
using System.IO.Compression;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetJpegPlanesFundPaddedComponentsAndRgbUntilDisposal()
    {
        foreach (var source in JpegAdmissionSources())
        {
            byte[] original = (byte[])source.Bytes.Clone();
            var pool = new TestMemoryBudget();
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                var first = JpegImage.ReadOwned(source.Bytes);
                TestAssert.Equal(source.Width, first.Width);
                TestAssert.Equal(source.Height, first.Height);
                TestAssert.True(first.Rgb.All(value => value == 128), "Pixels match the independent baseline JPEG decoder.");
                long rgbBytes = first.Rgb.Length;
                TestAssert.Equal(rgbBytes, pool.Reserved);
                TestAssert.True(pool.Peak >= rgbBytes + source.ComponentBytes, "Construction admits padded sampling planes alongside RGB output.");
                using var second = JpegImage.ReadOwned(source.Bytes);
                TestAssert.Equal(2 * rgbBytes, pool.Reserved);
                first.Dispose();
                first.Dispose();
                TestAssert.Equal(rgbBytes, pool.Reserved);
                TestAssert.True(second.Rgb.All(value => value == 128), "Another decoded owner remains funded and readable.");
                second.Dispose();
                TestAssert.Equal(0L, pool.Reserved);
                bool recolored = false;
                int requests = pool.Requests;
                var passthrough = OoxImageDecoder.Decode("image/jpeg", source.Bytes, rgb => { recolored = true; return rgb; });
                TestAssert.True(!recolored && ReferenceEquals(source.Bytes, passthrough.Bytes), "Ordinary JPEG embedding borrows encoded bytes without decoding or charging them again.");
                TestAssert.Equal(requests, pool.Requests);
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.True(source.Bytes.SequenceEqual(original), "Encoded caller storage stays unchanged.");
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void MemoryBudgetJpegRefusalTruncationAndCancellationRefundConstruction()
    {
        foreach (var source in JpegAdmissionSources())
        {
            long peak = ImageDecodePeak("image/jpeg", source.Bytes);
            var refused = new TestMemoryBudget(peak - 1);
            using (var conversion = OoxConversionBudget.BeginScope(null, refused))
            {
                var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => JpegImage.ReadOwned(source.Bytes));
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                TestAssert.Equal(0L, refused.Reserved);
            }
            using var cancellation = new CancellationTokenSource();
            var canceled = new TestMemoryBudget { RefuseChange = (_, after) => { if (after > 0) cancellation.Cancel(); return false; } };
            using (var conversion = OoxConversionBudget.BeginScope(null, canceled))
            {
                TestAssert.Throws<OperationCanceledException>(() => JpegImage.ReadOwned(source.Bytes, cancellation.Token));
                TestAssert.True(canceled.Peak > 0, "Cancellation fires after admission and before component allocation.");
                TestAssert.Equal(0L, canceled.Reserved);
                int requests = canceled.Requests;
                TestAssert.Throws<OperationCanceledException>(() => JpegImage.ReadOwned(source.Bytes, cancellation.Token));
                TestAssert.Equal(requests, canceled.Requests);
            }
            int scan = source.Bytes.AsSpan().IndexOf([0xFF, 0xDA]);
            TestAssert.True(scan > 0, "The fixture contains a baseline scan.");
            int entropy = scan + 2 + BinaryPrimitives.ReadUInt16BigEndian(source.Bytes.AsSpan(scan + 2));
            byte[] truncated = source.Bytes[..entropy];
            var malformed = new TestMemoryBudget();
            using (var conversion = OoxConversionBudget.BeginScope(null, malformed))
            {
                TestAssert.Throws<InvalidDataException>(() => JpegImage.ReadOwned(truncated));
                TestAssert.True(malformed.Peak > source.Width * source.Height * 3, "Truncation occurs after component admission.");
                TestAssert.Equal(0L, malformed.Reserved);
                byte[] noScan = [..source.Bytes[..scan], 0xFF, 0xD9];
                TestAssert.Throws<InvalidDataException>(() => JpegImage.ReadOwned(noScan));
                TestAssert.Equal(0L, malformed.Reserved);
                using var good = JpegImage.ReadOwned(source.Bytes);
                TestAssert.Equal((long)good.Rgb.Length, malformed.Reserved);
            }
            TestAssert.Equal(0L, malformed.Reserved);
        }
    }

    public static void MemoryBudgetJpegRefusalRemainsTerminalForRecolorAndCrop()
    {
        byte[] jpeg = TestFixtures.CreateSubsampledBaselineJpeg();
        long construction = ImageDecodePeak("image/jpeg", jpeg);
        TestAssert.True(construction > 0, "Public refusal must target decoded JPEG storage.");
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-jpeg-admission-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (bool crop in new[] { false, true })
            {
                var parts = new Dictionary<string, byte[]>();
                using (var archive = ZipFile.OpenRead(Case("PictureGeometry/picture-geometry-controls.pptx")))
                foreach (var entry in archive.Entries)
                {
                    using var stream = entry.Open();
                    using var buffer = new MemoryStream();
                    stream.CopyTo(buffer);
                    parts[entry.FullName] = entry.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ? jpeg : buffer.ToArray();
                }
                var contentTypes = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
                foreach (var type in contentTypes.Root!.Elements().Where(element => (string?)element.Attribute("ContentType") == "image/png"))
                    type.SetAttributeValue("ContentType", "image/jpeg");
                parts["[Content_Types].xml"] = TestFixtures.Utf8(contentTypes.ToString());
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                int modified = 0;
                foreach (string part in parts.Keys.Where(name => name.StartsWith("ppt/slides/slide", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal)).ToArray())
                {
                    var slide = XDocument.Parse(System.Text.Encoding.UTF8.GetString(parts[part]));
                    foreach (var blip in slide.Descendants(a + "blip"))
                    {
                        if (crop)
                        {
                            blip.Parent!.Elements(a + "srcRect").Remove();
                            blip.AddAfterSelf(new XElement(a + "srcRect", new XAttribute("l", "25000")));
                        }
                        else blip.Add(new XElement(a + "grayscl"));
                        modified++;
                    }
                    parts[part] = TestFixtures.Utf8(slide.ToString());
                }
                TestAssert.True(modified > 0, "Exercise existing real picture relationships.");
                string input = Path.Combine(directory, crop ? "crop.pptx" : "recolor.pptx");
                using (var package = TestFixtures.CreateZipPackage(parts)) File.WriteAllBytes(input, package.ToArray());
                // The same source first succeeds through the actual decoder.
                var successDiagnostics = new List<OoxPdfDiagnostic>();
                using (var source = File.OpenRead(input))
                using (var output = new MemoryStream())
                {
                    OoxPdfConverter.Convert(source, output, new OoxPdfOptions { InputKind=OoxPdfInputKind.Pptx,
                        FontResolver=new PortableRenderingTests.PinnedResolver(mathFont: true), DiagnosticSink=successDiagnostics.Add });
                    TestAssert.True(output.Length > 0, "The baseline effect conversion succeeds.");
                }
                TestAssert.True(successDiagnostics.All(d => d.Id is not "PPTX_UNSUPPORTED_IMAGE_RECOLOR" and not "IMAGE_CROP_UNSUPPORTED_FORMAT" and not "PPTX_NODE_RENDER_FAILED"), "The positive control does not take effect recovery.");
                foreach (bool file in new[] { false, true })
                {
                    bool rendering = false;
                    int refusals = 0;
                    var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                        if (!rendering || after != construction) return false;
                        refusals++;
                        return true;
                    }};
                    var diagnostics = new List<OoxPdfDiagnostic>();
                    var options = new OoxPdfOptions { InputKind=OoxPdfInputKind.Pptx, MemoryBudget=pool,
                        FontResolver=new ObservedResolver(() => rendering=true), DiagnosticSink=diagnostics.Add };
                    OoxPdfLimitExceededException failure;
                    if (file)
                    {
                        string output = Path.Combine(directory, "result.pdf");
                        File.WriteAllBytes(output, [7, 8, 9]);
                        failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, options));
                        TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Late effect refusal preserves the destination.");
                        TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                    }
                    else
                    {
                        using var source = File.OpenRead(input);
                        using var output = new MemoryStream();
                        failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(source, output, options));
                        TestAssert.Equal(0L, output.Length);
                        TestAssert.True(source.CanRead && output.CanWrite, "Caller streams remain open.");
                    }
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                    TestAssert.Equal(1, refusals);
                    TestAssert.True(diagnostics.All(d => d.Id is not "PPTX_UNSUPPORTED_IMAGE_RECOLOR" and not "IMAGE_CROP_UNSUPPORTED_FORMAT" and not "PPTX_NODE_RENDER_FAILED"), "Refusal cannot recover to original JPEG, clipping or node omission.");
                    TestAssert.Equal(0L, pool.Reserved);
                    TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                    TestAssert.Equal(0, pool.CallerDisposals);
                }
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static (byte[] Bytes, int Width, int Height, int ComponentBytes)[] JpegAdmissionSources() => [
        (TestFixtures.CreateGrayscaleBaselineJpeg(), 13, 9, 16 * 16),
        (TestFixtures.CreateSubsampledBaselineJpeg(), 17, 9, 32 * 16 + 2 * 16 * 8)
    ];
}
