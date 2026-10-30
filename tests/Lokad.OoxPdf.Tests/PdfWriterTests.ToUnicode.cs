using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PdfWriterTests
{
    internal static string ReadToUnicodeCMap(PdfEmbeddedFont font)
    {
        using var output = new MemoryStream();
        font.WriteToUnicodeCMapAsync(new PdfObjectWriter(output, default, new byte[4096]), default).GetAwaiter().GetResult();
        return Encoding.ASCII.GetString(output.ToArray());
    }

    public static void PdfToUnicodeStreamingMatchesFrozenBytesAndLengths()
    {
        string fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "Cases", "to-unicode-frozen.json"));
        using var golden = JsonDocument.Parse(File.ReadAllText(fixture));
        var expected = golden.RootElement.GetProperty(OperatingSystem.IsWindows() ? "Windows" : "Unix");
        TestAssert.Equal(Environment.NewLine, expected.GetProperty("NewLine").GetString());
        var font = TestFontBuilder.LoadTestFont();
        CultureInfo previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ar-EG");
            foreach (var row in expected.GetProperty("Cases").EnumerateArray())
            {
                PdfEmbeddedFont embedded = row.GetProperty("Mappings").ValueKind == JsonValueKind.Array
                    ? PdfEmbeddedFont.CreateForGlyphMappings(font, row.GetProperty("Mappings").EnumerateArray()
                        .Select(item => (checked((ushort)item.GetProperty("Glyph").GetInt32()), item.GetProperty("Unicode").GetInt32())), default)
                    : PdfEmbeddedFont.Create(font, TestFontBuilder.AllCodePoints().Take(row.GetProperty("Count").GetInt32()), default);
                byte[] bytes = Encoding.ASCII.GetBytes(ReadToUnicodeCMap(embedded));
                TestAssert.True(bytes.AsSpan().SequenceEqual(Convert.FromBase64String(row.GetProperty("Base64").GetString()!)),
                    "The live encoder must preserve the frozen pre-streaming CMap bytes: " + row.GetProperty("Id").GetString());
                TestAssert.Equal(row.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)));
                TestAssert.Equal(bytes.Length, embedded.MeasureToUnicodeCMap(new byte[4096], default));
                using var output = new MemoryStream();
                var writer = new PdfObjectWriter(output, default, new byte[4096]);
                writer.WriteStreamObjectHeaderAsync(1, string.Empty, bytes.Length).GetAwaiter().GetResult();
                embedded.WriteToUnicodeCMapAsync(writer, default).GetAwaiter().GetResult();
                writer.WriteStreamObjectTrailerAsync().GetAwaiter().GetResult();
                string framed = Encoding.ASCII.GetString(output.ToArray());
                TestAssert.True(framed.EndsWith("\nendstream\nendobj\n", StringComparison.Ordinal), "Resource framing retains its established trailing separator.");
                TestAssert.Equal(output.Length, writer.Position);
                TestAssert.Equal(1, writer.Offsets.Count);
            }
        }
        finally { CultureInfo.CurrentCulture = previous; }
        var empty = PdfEmbeddedFont.Create(font, [], default);
        var cancelled = new CancellationToken(canceled: true);
        TestAssert.Throws<OperationCanceledException>(() => empty.MeasureToUnicodeCMap(new byte[4096], cancelled));
        using var untouched = new MemoryStream();
        TestAssert.Throws<OperationCanceledException>(() => empty.WriteToUnicodeCMapAsync(new PdfObjectWriter(untouched, cancelled, new byte[4096]), cancelled).GetAwaiter().GetResult());
        TestAssert.Equal(0L, untouched.Length);
    }

    public static void PdfToUnicodeStreamingBoundsChunksAndPreservesCancellationAndIoFailure()
    {
        var font = PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(), TestFontBuilder.AllCodePoints(), default);
        var page = new PdfPage(100, 100, "BT /F1 12 Tf <0001> Tj ET\n", [new PdfFontResource("F1", font)]);
        foreach (string mode in new[] { "success", "cancel", "io" })
        {
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget();
            bool inMap = false, mapSeen = false;
            int blocks = 0, largest = 0;
            var io = new IOException("Destination refused a CMap block.");
            using var output = new ToUnicodeObservedOutput(bytes => {
                if (bytes.StartsWith("/CIDInit"u8)) { inMap = mapSeen = true; }
                if (!inMap) return;
                largest = Math.Max(largest, bytes.Length);
                TestAssert.True(bytes.Length <= 2048, "A destination sees one bounded 100-entry CMap block, never the complete map.");
                if (bytes.IndexOf("beginbfchar"u8) >= 0)
                {
                    blocks++;
                    if (mode == "cancel") cancellation.Cancel();
                    if (mode == "io") throw io;
                }
                if (bytes.IndexOf("endcmap"u8) >= 0) inMap = false;
            });
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                void Write() => PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]), new OoxConversionLimits(), null, cancellation.Token, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
                if (mode == "cancel")
                    TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Write).CancellationToken);
                else if (mode == "io")
                    TestAssert.True(ReferenceEquals(io, TestAssert.Throws<IOException>(Write)), "Preserve the destination exception without retry or replacement.");
                else Write();
                TestAssert.True(mapSeen && largest > 0, "The real PDF font resource reaches CMap emission.");
                TestAssert.Equal(mode == "success" ? 2 : 1, blocks);
                TestAssert.True(output.CanWrite, "The writer keeps the caller's destination open on every path.");
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void PdfToUnicodeStreamingPreservesFontAndOutputLimits()
    {
        var font = PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(), TestFontBuilder.AllCodePoints(), default);
        var page = new PdfPage(100, 100, "BT /F1 12 Tf <0001> Tj ET\n", [new PdfFontResource("F1", font)]);
        long rawBytes = font.FontProgramBytes.Length + Encoding.ASCII.GetByteCount(ReadToUnicodeCMap(font));
        using var baseline = new MemoryStream();
        var exactFont = new OoxConversionLimits { MaxPdfFontBytesPerConversion = rawBytes };
        using (var scope = OoxConversionBudget.BeginScope(exactFont))
        {
            PdfDocumentWriter.WriteStagedAsync(baseline, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]), exactFont, null, default, DateTimeOffset.UnixEpoch, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
            TestAssert.Equal(rawBytes, scope.Budget.PdfFontBytes);
            TestAssert.Equal(baseline.Length, scope.Budget.PdfOutputBytes);
        }
        using var refused = new MemoryStream();
        var shortFont = new OoxConversionLimits { MaxPdfFontBytesPerConversion = rawBytes - 1 };
        using (var scope = OoxConversionBudget.BeginScope(shortFont))
        {
            TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfDocumentWriter.WriteStagedAsync(refused, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]), shortFont, null, default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult());
            TestAssert.DoesNotContain("/Filter /FlateDecode /Length1", Encoding.ASCII.GetString(refused.ToArray()));
            TestAssert.Equal(0L, scope.Budget.PdfFontBytes);
        }
        foreach (long allowance in new[] { baseline.Length, baseline.Length - 1 })
        {
            using var output = new MemoryStream();
            var limits = new OoxConversionLimits { MaxPdfFontBytesPerConversion = rawBytes, MaxOutputBytesPerConversion = allowance };
            using var scope = OoxConversionBudget.BeginScope(limits);
            if (allowance < baseline.Length)
                TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]), limits, null, default, DateTimeOffset.UnixEpoch, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult());
            else
            {
                PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]), limits, null, default, DateTimeOffset.UnixEpoch, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
                TestAssert.True(output.ToArray().AsSpan().SequenceEqual(baseline.ToArray()), "An exact output bound preserves the complete PDF.");
            }
            TestAssert.True(output.Length <= allowance && scope.Budget.PdfOutputBytes <= allowance,
                "Every generated CMap chunk passes the existing output bound before publication.");
        }
    }

    private delegate void ToUnicodeWriteObserver(ReadOnlySpan<byte> bytes);
    private sealed class ToUnicodeObservedOutput(ToUnicodeWriteObserver observe) : MemoryStream
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
            // The inherited array overload avoids observing a derived span write twice.
            base.Write(bytes.ToArray(), 0, bytes.Length);
        }
    }
}
