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
            PdfDocumentWriter.WriteStaged(baseline, pages, limits, null, default, date);
            var pool = new TestMemoryBudget();
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                var staged = PdfDocumentWriter.ProduceStagedPages(pages, limits, null, default);
                long expected = PdfPageContentStaging.ChunkByteCount + (window == 0 ? 4096 : pages.Sum(page => (long)page.Content.Length));
                TestAssert.Equal(expected, pool.Reserved);
                byte[] scratch = staged.Staging.Scratch;
                using var output = new ObservedOutput(() => TestAssert.Equal(expected, pool.Reserved));
                PdfDocumentWriter.EmitStaged(output, staged, default, date);
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
        // Exercise both a resident-copy refusal and refusal of the file-buffer
        // overlap selected by the configured window. Neither creates a file.
        foreach (long window in new[] { 5L, 100L })
        {
            var pool = new TestMemoryBudget(5);
            using var conversion = OoxConversionBudget.BeginScope(null, pool);
            var diagnostics = new List<OoxPdfDiagnostic>();
            using var storage = new PdfPageContentStaging(window, diagnostics.Add);
            storage.BeginPage(default);
            storage.AppendPageBytes([1, 2, 3], 0, 3, default);
            storage.EndPage(default);
            storage.BeginPage(default);
            storage.AppendPageBytes([4, 5], 0, 2, default);
            TestAssert.Equal(5L, pool.Reserved);
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => storage.AppendPageBytes([6], 0, 1, default));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
            TestAssert.Equal(5L, pool.Reserved);
            TestAssert.True(!storage.IsFileBacked, "Refusal cannot select spill recovery or create a partially funded file.");
            TestAssert.Equal(0L, storage.SpilledBytes);
            TestAssert.Equal(0, diagnostics.Count);
            storage.EndPage(default);
            byte[] bytes = new byte[3];
            TestAssert.Equal(3, storage.ReadPageBytes(0, bytes, 0, 0, 3, default));
            TestAssert.True(bytes.SequenceEqual(new byte[] { 1, 2, 3 }), "Previous pages survive refusal unchanged.");
            TestAssert.Equal(2, storage.ReadPageBytes(1, bytes, 0, 0, 3, default));
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
            TestAssert.Equal(4096L, pool.Reserved); // Buffer remains live; old segments no longer do.
        });
        TestAssert.Equal(0L, pool.Reserved);
        storage.BeginPage(default);
        storage.AppendPageBytes([1, 2, 3], 0, 3, default);
        storage.EndPage(default);
        storage.BeginPage(default);
        storage.AppendPageBytes([4, 5], 0, 2, default);
        storage.AppendPageBytes([6], 0, 1, default);
        storage.EndPage(default);
        TestAssert.True(notified && storage.IsFileBacked, "The configured window selects spilling independently of admission.");
        TestAssert.Equal(4101L, pool.Peak);
        TestAssert.Equal(4096L, pool.Reserved);
        byte[] bytes = new byte[3];
        TestAssert.Equal(3, storage.ReadPageBytes(0, bytes, 0, 0, 3, default));
        TestAssert.True(bytes.SequenceEqual(new byte[] { 1, 2, 3 }), "Closed page transfers before its refund.");
        TestAssert.Equal(3, storage.ReadPageBytes(1, bytes, 0, 0, 3, default));
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
                TestAssert.Equal(PdfPageContentStaging.ChunkByteCount + 4096L, pool.Reserved);
                if (failure == "observer") throw new InvalidOperationException("observer failed");
            }
            using var output = new MemoryStream();
            void Write() => PdfDocumentWriter.WriteStaged(output, Pages(),
                new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, Observe, cancellation.Token);
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
        var refusal = TestAssert.Throws<OoxPdfLimitExceededException>(() => PdfDocumentWriter.WriteStaged(destination,
            [new PdfPage(612, 792, "q\nQ\n")], new(), null, default));
        TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, refusal.Kind);
        TestAssert.Equal(0L, refusingPool.Reserved);
        TestAssert.Equal(0L, destination.Length);
    }
}
