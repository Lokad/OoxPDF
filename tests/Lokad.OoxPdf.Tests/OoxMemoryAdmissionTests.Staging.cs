using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void MemoryBudgetStagingRetainsCapacityThroughEmissionWithoutChangingPdf()
    {
        var pages = new[] { new PdfPage(612, 792, "q\nQ\n" + new string('\n', 90_000)), new PdfPage(612, 792, "q\nQ\n") };
        foreach (long window in new[] { 0L, 200_000L })
        {
            var limits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = window };
            var date = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
            using var baseline = new MemoryStream();
            PdfDocumentWriter.WriteStagedAsync(baseline, AsyncEnumerable.ToAsyncEnumerable<PdfPage>(pages), limits, null, default, date, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
            var pool = new TestMemoryBudget();
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                var staged = PdfDocumentWriter.ProduceStagedPagesAsync(AsyncEnumerable.ToAsyncEnumerable<PdfPage>(pages), limits, null, default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
                long expected = PdfPageContentStaging.ChunkByteCount + (window == 0 ? 0L : pages.Sum(page => (long)page.Content.Length));
                TestAssert.Equal(expected, pool.Reserved);
                byte[] scratch = staged.Staging.Scratch;
                using var output = new ObservedOutput(() => TestAssert.Equal(expected, pool.Reserved));
                PdfDocumentWriter.EmitStagedAsync(output, staged, default, date).GetAwaiter().GetResult();
                TestAssert.True(output.Writes > 0, "Observe the real emission lifetime.");
                TestAssert.True(ReferenceEquals(scratch, staged.Staging.Scratch), "Encoding and readback share one admitted scratch buffer.");
                TestAssert.True(baseline.ToArray().AsSpan().SequenceEqual(output.ToArray()), "Admission preserves PDF bytes in both backing modes.");
                staged.Dispose();
                staged.Dispose();
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void MemoryBudgetStagingRefusalPreservesPendingContentWithoutSpill()
    {
        // Distinguish a resident-copy refusal from a resident-window boundary
        // without a provider. Neither creates temporary backing.
        foreach (long window in new[] { 5L, 100L })
        {
            var pool = new TestMemoryBudget(5);
            using var conversion = OoxConversionBudget.BeginScope(null, pool);
            var diagnostics = new List<OoxPdfDiagnostic>();
            string unselected = Path.Combine(Path.GetTempPath(), "ooxpdf-unselected-storage-" + Guid.NewGuid().ToString("N"));
            using var storage = new PdfPageContentStaging(window, diagnostics.Add,
                temporaryStorage: window == 100 ? new OoxPdfFileTemporaryStorage(unselected) : null);
            storage.BeginPage(default);
            storage.AppendPageBytesAsync([1, 2, 3], 0, 3, default).GetAwaiter().GetResult();
            storage.EndPageAsync(default).GetAwaiter().GetResult();
            storage.BeginPage(default);
            storage.AppendPageBytesAsync([4, 5], 0, 2, default).GetAwaiter().GetResult();
            TestAssert.Equal(5L, pool.Reserved);
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => storage.AppendPageBytesAsync([6], 0, 1, default).GetAwaiter().GetResult());
            TestAssert.Equal(window == 5 ? OoxPdfLimitKind.DocumentBound : OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
            TestAssert.Equal(5L, pool.Reserved);
            TestAssert.True(!storage.UsesTemporaryStorage, "Refusal cannot select spill recovery or create a partially funded file.");
            TestAssert.Equal(0L, storage.SpilledBytes);
            TestAssert.Equal(0, diagnostics.Count);
            TestAssert.True(!Directory.Exists(unselected), "Admission refusal must not try the configured temporary provider as recovery.");
            storage.EndPageAsync(default).GetAwaiter().GetResult();
            byte[] bytes = new byte[3];
            TestAssert.Equal(3, storage.ReadPageBytesAsync(0, bytes, 0, 0, 3, default).GetAwaiter().GetResult());
            TestAssert.True(bytes.SequenceEqual(new byte[] { 1, 2, 3 }), "Previous pages survive refusal unchanged.");
            TestAssert.Equal(2, storage.ReadPageBytesAsync(1, bytes, 0, 0, 3, default).GetAwaiter().GetResult());
            TestAssert.True(bytes.Take(2).SequenceEqual(new byte[] { 4, 5 }), "Refused append does not change pending content.");
            storage.Dispose();
            TestAssert.Equal(0L, pool.Reserved);
        }
    }

    public static void MemoryBudgetStagingSpillAdmitsOverlapAndRefundsResidentBytes()
    {
        var pool = new TestMemoryBudget();
        using var conversion = OoxConversionBudget.BeginScope(null, pool);
        bool notified = false;
        using var storage = new PdfPageContentStaging(5, _ => {
            notified = true;
            TestAssert.Equal(0L, pool.Reserved); // Unbuffered file backing retains none of the transferred segments.
        }, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath()));
        TestAssert.Equal(0L, pool.Reserved);
        storage.BeginPage(default);
        storage.AppendPageBytesAsync([1, 2, 3], 0, 3, default).GetAwaiter().GetResult();
        storage.EndPageAsync(default).GetAwaiter().GetResult();
        storage.BeginPage(default);
        storage.AppendPageBytesAsync([4, 5], 0, 2, default).GetAwaiter().GetResult();
        storage.AppendPageBytesAsync([6], 0, 1, default).GetAwaiter().GetResult();
        storage.EndPageAsync(default).GetAwaiter().GetResult();
        TestAssert.True(notified && storage.UsesTemporaryStorage, "The configured window selects spilling independently of admission.");
        TestAssert.Equal(5L, pool.Peak);
        TestAssert.Equal(0L, pool.Reserved);
        byte[] bytes = new byte[3];
        TestAssert.Equal(3, storage.ReadPageBytesAsync(0, bytes, 0, 0, 3, default).GetAwaiter().GetResult());
        TestAssert.True(bytes.SequenceEqual(new byte[] { 1, 2, 3 }), "Closed page transfers before its refund.");
        TestAssert.Equal(3, storage.ReadPageBytesAsync(1, bytes, 0, 0, 3, default).GetAwaiter().GetResult());
        TestAssert.True(bytes.SequenceEqual(new byte[] { 4, 5, 6 }), "The page open across escalation transfers without losing its prefix.");
        storage.Dispose();
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void MemoryBudgetStagingFailuresRefundScratchAndSpillStorage()
    {
        foreach (string failure in new[] { "producer", "observer", "cancellation" })
        {
            var pool = new TestMemoryBudget();
            using var conversion = OoxConversionBudget.BeginScope(null, pool);
            using var cancellation = new CancellationTokenSource();
            IEnumerable<PdfPage> Pages()
            {
                yield return new PdfPage(612, 792, "q\nQ\n");
                if (failure == "cancellation") cancellation.Cancel();
                else throw new IOException("producer failed");
                yield return new PdfPage(612, 792, "q\nQ\n");
            }
            void Observe(OoxPdfDiagnostic diagnostic)
            {
                TestAssert.Equal((long)PdfPageContentStaging.ChunkByteCount, pool.Reserved);
                if (failure == "observer") throw new InvalidOperationException("observer failed");
            }
            using var output = new MemoryStream();
            void Write() => PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable<PdfPage>(Pages()),
                new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, Observe, cancellation.Token, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
            if (failure == "cancellation")
                TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Write).CancellationToken);
            else if (failure == "observer") TestAssert.Equal("observer failed", TestAssert.Throws<InvalidOperationException>(Write).Message);
            else TestAssert.Equal("producer failed", TestAssert.Throws<IOException>(Write).Message);
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(0L, output.Length);
        }

        var refusingPool = new TestMemoryBudget(PdfPageContentStaging.ChunkByteCount - 1);
        using var refusingConversion = OoxConversionBudget.BeginScope(null, refusingPool);
        using var destination = new MemoryStream();
        var refusal = TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfDocumentWriter.WriteStagedAsync(destination, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([new PdfPage(612, 792, "q\nQ\n")]), new(), null, default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult());
        TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, refusal.Kind);
        TestAssert.Equal(0L, refusingPool.Reserved);
        TestAssert.Equal(0L, destination.Length);
    }
}
