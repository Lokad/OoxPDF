using System.Buffers.Binary;
using System.IO.Compression;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Imaging;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetImagePlanesFollowDecodedOwnerAndCompressionBorrow()
    {
        foreach (var source in ImageAdmissionSources())
        {
            byte[] original = (byte[])source.Bytes.Clone();
            using var baseline = OoxImageDecoder.DecodePixelsOwned(source.Type, source.Bytes);
            long planes = baseline.Rgb.Length + (baseline.Alpha?.Length ?? 0);
            var pool = new TestMemoryBudget();
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                var first = OoxImageDecoder.DecodePixelsOwned(source.Type, source.Bytes);
                using var second = OoxImageDecoder.DecodePixelsOwned(source.Type, source.Bytes);
                TestAssert.Equal(2 * planes, pool.Reserved);
                TestAssert.True(first.Rgb.SequenceEqual(baseline.Rgb), "RGB bytes stay exact.");
                TestAssert.True(first.Alpha is null ? baseline.Alpha is null : first.Alpha.SequenceEqual(baseline.Alpha!), "Alpha stays exact.");
                first.Dispose();
                first.Dispose();
                TestAssert.Equal(planes, pool.Reserved);
                TestAssert.Throws<ObjectDisposedException>(() => _ = first.Rgb);
                TestAssert.Throws<ObjectDisposedException>(() => _ = first.Alpha);
                TestAssert.True(second.Rgb.SequenceEqual(baseline.Rgb), "Another owner remains readable and funded.");
                second.Dispose();
                TestAssert.Equal(0L, pool.Reserved);
                bool observed = false;
                var image = OoxImageDecoder.Decode(source.Type, source.Bytes, rgb => {
                    observed = true;
                    TestAssert.Equal(planes, pool.Reserved);
                    return rgb;
                });
                TestAssert.True(observed && image.Bytes.Length > 0, "Real compression borrows admitted decoded planes.");
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
            TestAssert.True(source.Bytes.SequenceEqual(original), "Input/metadata remain borrowed and unchanged.");
        }
    }

    public static void MemoryBudgetImageRefusalAndCancellationReleaseConstruction()
    {
        foreach (var source in ImageAdmissionSources())
        {
            long peak = ImageDecodePeak(source.Type, source.Bytes);
            foreach (bool cancel in new[] { false, true })
            {
                using var cancellation = new CancellationTokenSource();
                var pool = new TestMemoryBudget(cancel ? long.MaxValue : peak - 1);
                if (cancel) pool.RefuseChange = (_, after) => { if (after > 0) cancellation.Cancel(); return false; };
                using (var conversion = OoxConversionBudget.BeginScope(null, pool))
                {
                    if (cancel)
                        TestAssert.Throws<OperationCanceledException>(() => OoxImageDecoder.DecodePixelsOwned(source.Type, source.Bytes, cancellation.Token));
                    else
                    {
                        var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxImageDecoder.DecodePixelsOwned(source.Type, source.Bytes));
                        TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                        TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                    }
                    TestAssert.Equal(0L, pool.Reserved);
                    int requests = pool.Requests;
                    cancellation.Cancel();
                    TestAssert.Throws<OperationCanceledException>(() => OoxImageDecoder.DecodePixelsOwned(source.Type, source.Bytes, cancellation.Token));
                    TestAssert.Equal(requests, pool.Requests);
                }
                TestAssert.Equal(0L, pool.Reserved);
            }
        }
    }

    public static void MemoryBudgetPngTrailerFailureRefundsAdmittedConstruction()
    {
        byte[] png = TestFixtures.CreateTinyStoredRgbPng();
        // PNG CRC handling is unchanged; corrupt the zlib trailer, which the
        // live decoder validates after it has acquired the pixel planes.
        for (int offset = 8; offset + 12 <= png.Length;)
        {
            int length = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(offset));
            if (png.AsSpan(offset + 4, 4).SequenceEqual("IDAT"u8))
            {
                png[offset + 8 + length - 1] ^= 1;
                break;
            }
            offset += length + 12;
        }
        var pool = new TestMemoryBudget();
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            TestAssert.Throws<InvalidDataException>(() => PngImage.ReadOwned(png));
            TestAssert.True(pool.Peak > 6, "Failure occurs after construction admission.");
            TestAssert.Equal(0L, pool.Reserved);
            using var good = PngImage.ReadOwned(TestFixtures.CreateTinyStoredRgbPng());
            TestAssert.Equal(6L, pool.Reserved);
        }
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void MemoryBudgetImageOwnersShareCapacityAcrossConcurrentConversions()
    {
        byte[] bmp = TestFixtures.CreateRgbBmp(16, 16, new byte[16 * 16 * 3]);
        var pool = new TestMemoryBudget(16 * 16 * 3);
        using var ready = new ManualResetEventSlim();
        using var finish = new ManualResetEventSlim();
        Task first = Task.Run(() => {
            using var conversion = OoxConversionBudget.BeginScope(null, pool);
            using var image = BmpImage.ReadOwned(bmp);
            ready.Set();
            TestAssert.True(finish.Wait(TimeSpan.FromSeconds(30)), "Concurrent owner must be released.");
            TestAssert.Equal(16 * 16 * 3, image.Rgb.Length);
        });
        try
        {
            TestAssert.True(ready.Wait(TimeSpan.FromSeconds(30)), "First conversion acquires its planes.");
            Task.Run(() => {
                using var conversion = OoxConversionBudget.BeginScope(null, pool);
                TestAssert.Throws<OoxPdfLimitExceededException>(() => BmpImage.ReadOwned(bmp));
            }).GetAwaiter().GetResult();
            TestAssert.Equal(16L * 16 * 3, pool.Reserved);
            TestAssert.Equal(1, pool.RootsDisposed);
        }
        finally { finish.Set(); first.GetAwaiter().GetResult(); }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void MemoryBudgetImageRefusalIsTerminalAcrossPublicConverters()
    {
        byte[] png = TestFixtures.CreateRgbPng(113, 71, new byte[113 * 71 * 3]);
        long decodeRequest = ImageDecodePeak("image/png", png);
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-image-admission-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (string relative in new[] { "ImageRecovery/picture-present.docx", "PictureGeometry/picture-geometry-controls.pptx" })
            {
                string input = Path.Combine(directory, Path.GetFileName(relative));
                using (var archive = ZipFile.OpenRead(Case(relative)))
                {
                    var entries = new Dictionary<string, byte[]>();
                    int replacements = 0;
                    foreach (var entry in archive.Entries)
                    {
                        using var content = entry.Open();
                        using var buffer = new MemoryStream();
                        content.CopyTo(buffer);
                        entries[entry.FullName] = entry.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                            ? png : buffer.ToArray();
                        if (ReferenceEquals(entries[entry.FullName], png)) replacements++;
                    }
                    TestAssert.True(replacements > 0, "Use an existing real picture relationship.");
                    using var package = TestFixtures.CreateZipPackage(entries);
                    File.WriteAllBytes(input, package.ToArray());
                }
                foreach (bool file in new[] { false, true })
                {
                    int refusals = 0;
                    var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                        if (after != decodeRequest) return false;
                        refusals++;
                        return true;
                    }};
                    var diagnostics = new List<OoxPdfDiagnostic>();
                    var options = new OoxPdfOptions { MemoryBudget=pool, DiagnosticSink=diagnostics.Add,
                        FontResolver=new PortableRenderingTests.PinnedResolver(mathFont: true) };
                    OoxPdfLimitExceededException failure;
                    if (file)
                    {
                        string output = Path.Combine(directory, "result.pdf");
                        File.WriteAllBytes(output, [7, 8, 9]);
                        failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, options));
                        TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Image refusal preserves the old destination.");
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
                    TestAssert.Equal(1, refusals);
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                    TestAssert.True(diagnostics.All(d => d.Id is not "IMAGE_UNSUPPORTED_FORMAT" and not "PPTX_NODE_RENDER_FAILED"), "Refusal cannot become content omission.");
                    TestAssert.Equal(0L, pool.Reserved);
                    TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                    TestAssert.Equal(0, pool.CallerDisposals);
                }
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static (string Type, byte[] Bytes)[] ImageAdmissionSources() => [
        ("image/png", TestFixtures.CreateTinyStoredRgbPng()),
        ("image/png", TestFixtures.CreateRgbaPng(2, 1, [255, 0, 0, 128, 0, 0, 255, 255])),
        ("image/bmp", TestFixtures.CreateRgbBmp(2, 1, [255, 0, 0, 0, 0, 255]))
    ];

    private static long ImageDecodePeak(string type, byte[] bytes)
    {
        var pool = new TestMemoryBudget();
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        using (var image = OoxImageDecoder.DecodePixelsOwned(type, bytes)) { }
        TestAssert.Equal(0L, pool.Reserved);
        return pool.Peak;
    }
}
