using System.Buffers.Binary;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PdfWriterTests
{
    internal static string ReadWidthArray(PdfEmbeddedFont font)
    {
        using var output = new MemoryStream();
        font.WriteWidthArray(new PdfObjectWriter(output, default), default);
        return Encoding.ASCII.GetString(output.ToArray());
    }

    public static void PdfWidthStreamingMatchesFrozenOrderMetricsAndBytes()
    {
        string cases = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases"));
        using var oracle = JsonDocument.Parse(File.ReadAllText(Path.Combine(cases, "widths-frozen.json")));
        CultureInfo previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ar-EG");
            foreach (var sample in oracle.RootElement.GetProperty("Cases").EnumerateArray())
            {
                string id = sample.GetProperty("Id").GetString()!;
                byte[] bytes = id == "compound" ? File.ReadAllBytes(Path.Combine(cases, "PortableFonts", "DejaVuSans.ttf")) : TestFontBuilder.CreateTestFont();
                if (id == "compound") TestAssert.Equal(oracle.RootElement.GetProperty("PortableFontSha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)));
                int head = TestFontBuilder.GetTableRange(bytes, "head").Offset;
                if (id == "rounding")
                {
                    BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(head + 18), 2000);
                    int hmtx = TestFontBuilder.GetTableRange(bytes, "hmtx").Offset;
                    foreach (var (glyph, advance) in new[] { (1, 0), (35, 1), (80, 3), (129, 65535), (130, 32767) })
                        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(hmtx + glyph * 4), (ushort)advance);
                }
                if (id == "whole-font") BinaryPrimitives.WriteInt16BigEndian(bytes.AsSpan(head + 50), 2);
                var font = OpenTypeFont.Load(bytes);
                PdfEmbeddedFont embedded = sample.GetProperty("Mappings").ValueKind == JsonValueKind.Array
                    ? PdfEmbeddedFont.CreateForGlyphMappings(font, sample.GetProperty("Mappings").EnumerateArray().Select(value =>
                        (checked((ushort)value.GetProperty("Glyph").GetInt32()), value.GetProperty("Unicode").GetInt32())), default)
                    : PdfEmbeddedFont.Create(font, id == "compound" ? new[] { 0x00C5, 0x00E9, 0x0301 } :
                        TestFontBuilder.AllCodePoints().Take(sample.GetProperty("Count").GetInt32()), default);
                TestAssert.Equal(sample.GetProperty("UsesSubset").GetBoolean(), embedded.UsesSubsetFontProgram);
                byte[] actual = Encoding.ASCII.GetBytes(ReadWidthArray(embedded));
                TestAssert.True(actual.AsSpan().SequenceEqual(Convert.FromBase64String(sample.GetProperty("Base64").GetString()!)),
                    "Live width emission preserves frozen original CID order, zero/fractional metrics and bytes: " + id);
                TestAssert.Equal(sample.GetProperty("Sha256").GetString(), Convert.ToHexString(SHA256.HashData(actual)));
                using var untouched = new MemoryStream();
                var cancelled = new CancellationToken(canceled: true);
                TestAssert.Throws<OperationCanceledException>(() => embedded.WriteWidthArray(new PdfObjectWriter(untouched, default), cancelled));
                TestAssert.Equal(0L, untouched.Length);
            }
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    public static void PdfWidthStreamingBoundsConstructionAndKeepsFailureOwnership()
    {
        var font = PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(), TestFontBuilder.AllCodePoints(), default);
        var page = new PdfPage(100, 100, "BT /F1 12 Tf <0001> Tj ET\n", [new PdfFontResource("F1", font)]);
        foreach (string mode in new[] { "success", "cancel", "io" })
        {
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget();
            bool inWidths = false, seen = false;
            int blocks = 0;
            var io = new IOException("Destination refused a width block.");
            using var output = new ToUnicodeObservedOutput(bytes => {
                if (bytes.IndexOf("/CIDToGIDMap /Identity /W "u8) >= 0) { inWidths = seen = true; }
                if (!inWidths) return;
                TestAssert.True(bytes.Length <= 512, "CID font widths reach the destination in bounded chunks, never a full width map/string.");
                if (bytes.IndexOf("/W "u8) < 0)
                {
                    blocks++;
                    if (mode == "cancel") cancellation.Cancel();
                    if (mode == "io") throw io;
                }
                if (bytes.EndsWith("]]"u8)) inWidths = false;
            });
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                void Write() => PdfDocumentWriter.WriteStaged(output, [page], new OoxConversionLimits(), null, cancellation.Token, DateTimeOffset.UnixEpoch);
                if (mode == "cancel") TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Write).CancellationToken);
                else if (mode == "io") TestAssert.True(ReferenceEquals(io, TestAssert.Throws<IOException>(Write)), "The original destination failure must escape without retry.");
                else Write();
                TestAssert.True(seen && (mode == "success" ? blocks > 1 : blocks == 1), "The real resource writer must exercise bounded widths and stop on failure.");
                TestAssert.True(output.CanWrite, "The caller retains its output stream on all paths.");
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void PdfWidthStreamingRejectsOutputInsideTheDictionary()
    {
        var font = PdfEmbeddedFont.Create(TestFontBuilder.LoadTestFont(), TestFontBuilder.AllCodePoints(), default);
        var page = new PdfPage(100, 100, "BT /F1 12 Tf <0001> Tj ET\n", [new PdfFontResource("F1", font)]);
        using var complete = new MemoryStream();
        PdfDocumentWriter.WriteStaged(complete, [page], new OoxConversionLimits(), null, default, DateTimeOffset.UnixEpoch);
        long allowance = Encoding.ASCII.GetString(complete.ToArray()).IndexOf("/W [", StringComparison.Ordinal) + 100;
        TestAssert.True(allowance > 100 && allowance < complete.Length, "The selected allowance must cut the real width array.");
        using var refused = new MemoryStream();
        var limits = new OoxConversionLimits { MaxOutputBytesPerConversion = allowance };
        using var scope = OoxConversionBudget.BeginScope(limits);
        TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfDocumentWriter.WriteStaged(refused, [page], limits, null, default));
        TestAssert.True(refused.Length <= allowance && scope.Budget.PdfOutputBytes <= allowance,
            "Each width chunk is admitted before reaching the caller's destination.");
        TestAssert.DoesNotContain("endobj\n", Encoding.ASCII.GetString(refused.ToArray()).Split("/CIDFontType2", StringSplitOptions.None)[1]);
    }
}
