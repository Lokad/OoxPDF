using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void VectorMemoryStagedResourcesSurviveSpillAndEmission()
    {
        PdfPage page = VectorMemoryPage();
        foreach (long window in new[] { 0L, 1_000_000L })
        {
            var limits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = window };
            using var baseline = new MemoryStream();
            PdfDocumentWriter.WriteStagedAsync(baseline, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page, page]), limits, null, default,
                DateTimeOffset.UnixEpoch, new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
            var pool = new TestMemoryBudget();
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                var staged = PdfDocumentWriter.ProduceStagedPagesAsync(AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page, page]), limits, null, default,
                    new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
                TestAssert.Equal(1, staged.GroupEntries.Count); // Existing shared identity is staged once.
                long content = window == 0 ? 0 : 2L * page.Content.Length + page.Groups[0].Group.Content.Length;
                long held = pool.Reserved;
                TestAssert.True(held > PdfPageContentStaging.ChunkByteCount + content, "Retained vector descriptors need funding after content spills or copies.");
                using var output = new ObservedOutput(() => TestAssert.Equal(held, pool.Reserved));
                PdfDocumentWriter.EmitStagedAsync(output, staged, default, DateTimeOffset.UnixEpoch).GetAwaiter().GetResult();
                TestAssert.True(baseline.ToArray().SequenceEqual(output.ToArray()), "Funding preserves emitted PDF bytes in both backing modes.");
                staged.Dispose(); staged.Dispose();
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void VectorMemoryRefusalAndCancellationDoNotSelectSpillRecovery()
    {
        PdfPage page = VectorMemoryPage();
        foreach (bool cancel in new[] { false, true })
        {
            string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-vector-refusal-" + Guid.NewGuid().ToString("N"));
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget(PdfPageContentStaging.ChunkByteCount);
            if (cancel) pool.RefuseChange = (_, after) => { if (after > 0 && after < PdfPageContentStaging.ChunkByteCount) cancellation.Cancel(); return false; };
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                using var output = new MemoryStream();
                void Write() => PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]),
                    new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, cancellation.Token,
                    temporaryStorage: new OoxPdfFileTemporaryStorage(directory)).GetAwaiter().GetResult();
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Write).CancellationToken);
                else TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(Write).Kind);
                TestAssert.True(pool.Peak > 0, "Descriptors hold capacity before the encoding window.");
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(0L, output.Length);
                TestAssert.True(!Directory.Exists(directory), "Refusal/cancellation must not choose temporary backing as recovery.");
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void VectorMemoryProducerAndSpillObserverFailuresRefundResources()
    {
        PdfPage page = VectorMemoryPage();
        foreach (bool observer in new[] { false, true })
        {
            var pool = new TestMemoryBudget();
            var sentinel = new InvalidOperationException("vector staging sentinel");
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                IEnumerable<PdfPage> Pages()
                {
                    yield return page;
                    TestAssert.True(pool.Reserved > PdfPageContentStaging.ChunkByteCount, "The staged group outlives its producer page.");
                    throw sentinel;
                }
                using var output = new MemoryStream();
                var actual = TestAssert.Throws<InvalidOperationException>(() => PdfDocumentWriter.WriteStagedAsync(output,
                    AsyncEnumerable.ToAsyncEnumerable<PdfPage>(Pages()), new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 },
                    observer ? _ => { TestAssert.True(pool.Reserved > PdfPageContentStaging.ChunkByteCount, "Spill keeps live resource funding."); throw sentinel; } : null,
                    default, temporaryStorage: new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult());
                TestAssert.Equal(sentinel, actual);
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(0L, output.Length);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void VectorMemoryEmissionFailureKeepsResourcesUntilOwnerDisposes()
    {
        PdfPage page = VectorMemoryPage();
        var pool = new TestMemoryBudget();
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            var staged = PdfDocumentWriter.ProduceStagedPagesAsync(AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]),
                new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, default,
                new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
            long held = pool.Reserved;
            var sentinel = new IOException("vector emission sentinel");
            using var output = new ObservedOutput(() => { TestAssert.True(pool.Reserved > PdfPageContentStaging.ChunkByteCount, "Emission borrows funded descriptors."); throw sentinel; });
            TestAssert.Equal(sentinel, TestAssert.Throws<IOException>(() => PdfDocumentWriter.EmitStagedAsync(output, staged, default).GetAwaiter().GetResult()));
            TestAssert.Equal(held, pool.Reserved);
            staged.Dispose();
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    private static PdfPage VectorMemoryPage()
    {
        var child = new PdfGraphicsBuilder();
        child.PaintAxialShading(0, 0, 100, 0, 255, 0, 0, 0, 0, 255);
        var graphics = new PdfGraphicsBuilder();
        graphics.DrawTransparencyGroup(child, new PdfRectangle(0, 0, 100, 100), .5);
        return new PdfPage(100, 100, graphics.ToString(), [], [], graphics.ExtGStates, graphics.Shadings, [], [], groups: graphics.Groups);
    }
}
