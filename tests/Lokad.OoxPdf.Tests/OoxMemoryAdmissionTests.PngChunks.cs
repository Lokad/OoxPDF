using System.Buffers.Binary;
using Lokad.OoxPdf.Imaging;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetPngChunkMetadataDoesNotGrowBeforeAdmission()
    {
        byte[] source = TestFixtures.CreateTinyStoredRgbPng();
        byte[] chunked = RepartitionPngIdat(source, 1, 100000);
        // Warm decoder and callback plumbing before comparing the same payload.
        Measure(source);
        var control = Measure(source);
        var repeated = Measure(chunked);
        Console.WriteLine($"PNG chunk admission: control={control.Allocated}, repeated={repeated.Allocated}, request={repeated.Requested}");
        TestAssert.Equal(control.Requested, repeated.Requested);
        TestAssert.True(repeated.Allocated - control.Allocated < 128 * 1024,
            "Chunk partitioning must not create significant metadata before image admission.");

        static (long Allocated, long Requested) Measure(byte[] bytes)
        {
            long start = 0, allocated = -1, requested = -1;
            var pool = new TestMemoryBudget
            {
                RefuseChange = (_, after) => {
                    allocated = GC.GetAllocatedBytesForCurrentThread() - start;
                    requested = after;
                    return true;
                }
            };
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                start = GC.GetAllocatedBytesForCurrentThread();
                var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => PngImage.ReadOwned(bytes));
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                TestAssert.Equal(1, pool.Requests);
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.True(allocated >= 0 && requested > 0, "The measurement reaches actual image admission.");
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            return (allocated, requested);
        }
    }

    public static void MemoryBudgetPngChunkPartitionsKeepPixelsAndReservations()
    {
        byte[] rgba = Enumerable.Range(0, 24).Select(value => (byte)(value * 11)).ToArray();
        byte[][] sources = [
            TestFixtures.CreateTinyStoredRgbPng(),
            TestFixtures.CreateRgbaPng(3, 2, rgba),
            TestFixtures.CreateInterlacedRgbaPng(3, 2, rgba),
            TestFixtures.CreateGrayscalePng(3, 1, [0, 128, 255]),
            TestFixtures.CreatePackedIndexedPng(5, 1, 2,
                [255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255], [0, 1, 2, 3, 2])
        ];
        foreach (byte[] source in sources)
        {
            using var baseline = PngImage.ReadOwned(source);
            long peak = ImageDecodePeak("image/png", source);
            foreach (int chunkSize in new[] { 1, 2, 17 })
            foreach (int emptyChunks in new[] { 0, 31 })
            {
                byte[] bytes = RepartitionPngIdat(source, chunkSize, emptyChunks);
                byte[] original = (byte[])bytes.Clone();
                var pool = new TestMemoryBudget();
                using (var conversion = OoxConversionBudget.BeginScope(null, pool))
                {
                    using (var decoded = PngImage.ReadOwned(bytes))
                    {
                        TestAssert.Equal(baseline.Width, decoded.Width);
                        TestAssert.Equal(baseline.Height, decoded.Height);
                        TestAssert.True(decoded.Rgb.SequenceEqual(baseline.Rgb), "IDAT partitions preserve RGB pixels.");
                        TestAssert.True(decoded.Alpha is null ? baseline.Alpha is null : decoded.Alpha.SequenceEqual(baseline.Alpha!),
                            "IDAT partitions preserve alpha pixels.");
                        TestAssert.Equal((long)decoded.Rgb.Length + (decoded.Alpha?.Length ?? 0), pool.Reserved);
                        TestAssert.Equal(peak, pool.Peak);
                    }
                    TestAssert.Equal(0L, pool.Reserved);
                }
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.True(bytes.SequenceEqual(original), "Source bytes remain borrowed and unchanged.");
                // Existing decoding stops at IEND, even with an invalid trailing chunk.
                using var trailing = PngImage.ReadOwned(bytes.Concat(new byte[] { 127, 255, 255, 255, 73, 68, 65, 84 }).ToArray());
                TestAssert.True(trailing.Rgb.SequenceEqual(baseline.Rgb), "Copying data stays within the validated PNG prefix.");
            }
        }
    }

    public static void MemoryBudgetPngChunkRefusalAndCancellationStayTerminal()
    {
        byte[] bytes = RepartitionPngIdat(TestFixtures.CreateTinyStoredRgbPng(), 1, 100000);
        foreach (bool cancel in new[] { false, true })
        {
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget { RefuseChange = (_, _) => {
                if (cancel) cancellation.Cancel();
                return !cancel;
            } };
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                if (cancel)
                    TestAssert.Throws<OperationCanceledException>(() => PngImage.ReadOwned(bytes, cancellation.Token));
                else
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission,
                        TestAssert.Throws<OoxPdfLimitExceededException>(() => PngImage.ReadOwned(bytes)).Kind);
                TestAssert.Equal(1, pool.Requests);
                TestAssert.Equal(0L, pool.Reserved);
                cancellation.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => PngImage.ReadOwned(bytes, cancellation.Token));
                TestAssert.Equal(1, pool.Requests);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    private static byte[] RepartitionPngIdat(byte[] source, int chunkSize, int emptyChunks)
    {
        using var output = new MemoryStream();
        output.Write(source.AsSpan(0, 8));
        for (int offset = 8; offset + 12 <= source.Length;)
        {
            int length = BinaryPrimitives.ReadInt32BigEndian(source.AsSpan(offset, 4));
            if (source.AsSpan(offset + 4, 4).SequenceEqual("IDAT"u8))
            {
                for (int i = 0; i < emptyChunks; i++) TestFixtures.WritePngChunk(output, "IDAT", []);
                for (int start = 0; start < length; start += chunkSize)
                    TestFixtures.WritePngChunk(output, "IDAT", source.AsSpan(offset + 8 + start, Math.Min(chunkSize, length - start)));
            }
            else output.Write(source.AsSpan(offset, length + 12));
            offset += length + 12;
        }
        return output.ToArray();
    }
}
