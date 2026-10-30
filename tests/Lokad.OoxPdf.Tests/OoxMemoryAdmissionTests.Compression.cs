using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Imaging;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetCompressionPreservesBytesAndAdmitsActualCapacity()
    {
        foreach (int size in new[] { 0, 1, 8191, 8192, 65539, 150000 })
        foreach (bool random in new[] { false, true })
        {
            byte[] input = CompressionInput(size, random);
            byte[] expected = OriginalCompression(input);
            var pool = new TestMemoryBudget();
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            using (var storage = OoxMemory.Reserve(conversion.Budget.Memory, 13))
            {
                byte[] actual = PdfStreamCompression.Compress(input, storage, 13, default);
                TestAssert.True(actual.SequenceEqual(expected), "Encoded bytes preserve the original one-write zlib operation.");
                TestAssert.Equal(13L + actual.Length, pool.Reserved);
                TestAssert.True(pool.Peak >= 13L + 2L * actual.Length, "Final copying admits the live capacity alongside the retained result.");
                if (!random && size == 150000)
                    TestAssert.True(pool.Peak < size / 2, "Compressible data admits actual output capacity instead of worst-case source-sized storage.");
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void MemoryBudgetCompressionRefusalAndCancellationNeverRetry()
    {
        byte[] bytes = CompressionInput(150000, true);
        int requests = 0;
        var probe = new TestMemoryBudget { RefuseChange = (_, after) => { if (after > 13) requests++; return false; } };
        using (var conversion = OoxConversionBudget.BeginScope(null, probe))
        using (var storage = OoxMemory.Reserve(conversion.Budget.Memory, 13))
            _ = PdfStreamCompression.Compress(bytes, storage, 13, default);
        TestAssert.True(requests >= 3, "Exercise initial growth, replacement growth and final copying.");
        foreach (int target in new[] { 1, 2, requests })
        foreach (bool cancel in new[] { false, true })
        {
            int attempts = 0;
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                if (after <= 13 || ++attempts != target) return false;
                if (cancel) { cancellation.Cancel(); return false; }
                return true;
            }};
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            using (var storage = OoxMemory.Reserve(conversion.Budget.Memory, 13))
            {
                if (cancel)
                {
                    var failure = TestAssert.Throws<OperationCanceledException>(() => PdfStreamCompression.Compress(bytes, storage, 13, cancellation.Token));
                    TestAssert.Equal(cancellation.Token, failure.CancellationToken);
                }
                else
                {
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfStreamCompression.Compress(bytes, storage, 13, default));
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                }
                TestAssert.Equal(target, attempts);
                TestAssert.Equal(13L, pool.Reserved);
                TestAssert.Equal(13L, TestAssert.NotNull(storage).Bytes);
            }
            TestAssert.Equal(0L, pool.Reserved);
        }
    }

    public static void MemoryBudgetCompressionImagePayloadSurvivesEmissionAndAlphaFailureRefunds()
    {
        byte[] rgba = [255, 0, 0, 128, 0, 0, 255, 255];
        byte[] source = TestFixtures.CreateRgbaPng(2, 1, rgba);
        PdfImageXObject expected = OoxImageDecoder.Decode("image/png", source, static rgb => rgb);
        var pool = new TestMemoryBudget();
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            PdfImageXObject image = OoxImageDecoder.Decode("image/png", source, static rgb => rgb);
            TestAssert.True(image.HasIdenticalContent(expected), "Admission preserves RGB and alpha encoding.");
            long retained = image.Bytes.Length + TestAssert.NotNull(image.Alpha).Length;
            TestAssert.Equal(retained, pool.Reserved);
            using var output = new ObservedOutput(() => TestAssert.True(pool.Reserved >= retained,
                "Encoded RGB/alpha remain funded through actual image writes."));
            var page = new PdfPage(100, 100, "q 20 0 0 10 0 0 cm /Im1 Do Q\n", [], [new PdfImageResource("Im1", image)]);
            PdfDocumentWriter.WriteStagedAsync(output, [page], new OoxConversionLimits { MaxResidentPageContentBytesPerConversion=0 }, null, default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
            TestAssert.True(output.Writes > 0, "Exercise actual resource serialization.");
            TestAssert.Equal(retained, pool.Reserved);
        }
        TestAssert.Equal(0L, pool.Reserved);
        int refusals = 0;
        var refused = new TestMemoryBudget { RefuseChange = (before, after) => {
            if (before != expected.Bytes.Length || after <= before) return false;
            refusals++;
            return true;
        }};
        using (var conversion = OoxConversionBudget.BeginScope(null, refused))
        {
            TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxImageDecoder.Decode("image/png", source, static rgb => rgb));
            TestAssert.Equal(1, refusals);
            TestAssert.Equal(0L, refused.Reserved);
        }
    }

    public static void MemoryBudgetCompressionFontRefundsAfterItsActualStreamWrite()
    {
        PdfEmbeddedFont font = PdfEmbeddedFont.Create(OpenTypeFont.Load(TestFontBuilder.CreateTestFont()), [65, 66], default);
        byte[] expected = OriginalCompression(font.FontProgramBytes.Span);
        foreach (bool failOutput in new[] { false, true })
        {
            var pool = new TestMemoryBudget();
            bool payload = false, following = false;
            long staging = -1;
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                using var output = new CompressionObservedOutput(bytes => {
                    if (staging < 0) staging = pool.Reserved;
                    if (bytes.SequenceEqual(expected))
                    {
                        payload = true;
                        TestAssert.Equal(staging + expected.Length, pool.Reserved);
                        if (failOutput) throw new IOException("Output refused the font payload.");
                    }
                    else if (payload && bytes.StartsWith("<< "u8) && bytes.IndexOf("/Length"u8) >= 0 && bytes.IndexOf("/Filter"u8) < 0)
                    {
                        following = true;
                        TestAssert.Equal(staging, pool.Reserved);
                    }
                });
                var page = new PdfPage(100, 100, "BT /F1 12 Tf <0001> Tj ET\n", [new PdfFontResource("F1", font)]);
                if (failOutput)
                    TestAssert.Throws<IOException>(() => PdfDocumentWriter.WriteStagedAsync(output, [page], new OoxConversionLimits { MaxResidentPageContentBytesPerConversion=0 }, null, default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult());
                else
                    PdfDocumentWriter.WriteStagedAsync(output, [page], new OoxConversionLimits { MaxResidentPageContentBytesPerConversion=0 }, null, default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
                TestAssert.True(payload && (failOutput || following), "Observe real font emission and its following ToUnicode stream.");
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void MemoryBudgetCompressionPublicRefusalPreservesDestinationsAndCallerStreams()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-compression-admission-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (bool image in new[] { false, true })
            foreach (OoxPdfInputKind kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
            {
                string relative = kind == OoxPdfInputKind.Docx ? "AdjacentTables/adjacent.docx" : "pptx-ladder-02-plain-text.pptx";
                if (image) relative = kind == OoxPdfInputKind.Docx ? "ImageRecovery/picture-present.docx" : "PictureGeometry/picture-geometry-controls.pptx";
                string input = Path.Combine(directory, image ? "image" : "font") + (kind == OoxPdfInputKind.Docx ? ".docx" : ".pptx");
                using (var archive = ZipFile.OpenRead(Case(relative)))
                {
                    var parts = new Dictionary<string, byte[]>();
                    foreach (var entry in archive.Entries)
                    {
                        using var stream = entry.Open();
                        using var buffer = new MemoryStream();
                        stream.CopyTo(buffer);
                        byte[] content = buffer.ToArray();
                        if (!image && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                        {
                            var xml = XDocument.Parse(Encoding.UTF8.GetString(content));
                            foreach (var text in xml.Descendants().Where(element => element.Name.LocalName == "t")) text.Value = "A";
                            content = TestFixtures.Utf8(xml.ToString());
                        }
                        parts[entry.FullName] = content;
                    }
                    using var package = TestFixtures.CreateZipPackage(parts);
                    File.WriteAllBytes(input, package.ToArray());
                }
                foreach (bool file in new[] { false, true })
                {
                    int refusals = 0;
                    var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                        if (after != 256) return false;
                        refusals++;
                        return true;
                    }};
                    var diagnostics = new List<OoxPdfDiagnostic>();
                    var options = new OoxPdfOptions { InputKind=kind, MemoryBudget=pool, DiagnosticSink=diagnostics.Add,
                        FontResolver=new ProgramResolver(new MemoryFontProgramSource("caller-font", TestFontBuilder.CreateTestFont())) };
                    OoxPdfLimitExceededException failure;
                    if (file)
                    {
                        string output = Path.Combine(directory, "result.pdf");
                        File.WriteAllBytes(output, [7, 8, 9]);
                        failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, options));
                        TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Refusal preserves an existing file destination.");
                        TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                    }
                    else
                    {
                        using var source = File.OpenRead(input);
                        using var output = new MemoryStream();
                        failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(source, output, options));
                        TestAssert.True(image ? output.Length == 0 : output.Length > 0,
                            "Image refusal precedes publication; font-stream refusal can leave the documented PDF prefix.");
                        TestAssert.True(source.CanRead && output.CanWrite, "Caller streams remain open.");
                    }
                    TestAssert.Equal(1, refusals);
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                    TestAssert.True(diagnostics.All(d => d.Id is not "IMAGE_UNSUPPORTED_FORMAT" and not "PPTX_NODE_RENDER_FAILED"), "Admission cannot become content omission.");
                    TestAssert.Equal(0L, pool.Reserved);
                    TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                    TestAssert.Equal(0, pool.CallerDisposals);
                }
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static byte[] CompressionInput(int size, bool random)
    {
        var bytes = new byte[size];
        if (random) new Random(42).NextBytes(bytes);
        return bytes;
    }

    private static byte[] OriginalCompression(ReadOnlySpan<byte> bytes)
    {
        using var output = new MemoryStream();
        using (var zlib = new ZLibStream(output, CompressionLevel.SmallestSize, leaveOpen: true)) zlib.Write(bytes);
        return output.ToArray();
    }

    private delegate void CompressionWriteObserver(ReadOnlySpan<byte> bytes);
    private sealed class CompressionObservedOutput(CompressionWriteObserver observe) : MemoryStream
    {
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> bytes, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Write(bytes.Span);
            return ValueTask.CompletedTask;
        }

        public override void Write(ReadOnlySpan<byte> bytes)
        {
            observe(bytes);
            // Avoid MemoryStream's derived-type span forwarding observing twice.
            byte[] copy = bytes.ToArray();
            base.Write(copy, 0, copy.Length);
        }
        public override void Write(byte[] bytes, int offset, int count)
        {
            observe(bytes.AsSpan(offset, count));
            base.Write(bytes, offset, count);
        }
    }
}
