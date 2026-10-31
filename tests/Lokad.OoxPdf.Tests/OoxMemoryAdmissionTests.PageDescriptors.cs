using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    public static void StagedPageMemoryScalesWithRecordsAndSurvivesEmission()
    {
        foreach (long window in new[] { 0L, 1_000_000L })
        {
            var pool = new TestMemoryBudget();
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            {
                var pages = Enumerable.Repeat(new PdfPage(100, 100, "q\nQ\n"), 1000);
                var staged = PdfDocumentWriter.ProduceStagedPagesAsync(AsyncEnumerable.ToAsyncEnumerable(pages),
                    new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = window }, null, default,
                    new OoxPdfFileTemporaryStorage(Path.GetTempPath())).GetAwaiter().GetResult();
                long held = pool.Reserved;
                long payload = PdfPageContentStaging.ChunkByteCount + (window == 0 ? 0 : 4000);
                TestAssert.True(held >= payload + 128L * staged.Pages.Count, "Retained page/record storage must not disappear from admission when content is small or spilled.");
                using var output = new ObservedOutput(() => TestAssert.Equal(held, pool.Reserved));
                PdfDocumentWriter.EmitStagedAsync(output, staged, default).GetAwaiter().GetResult();
                TestAssert.True(output.Writes > 0, "Real planning/numbering and emission borrow the funded records.");
                staged.Dispose();
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void StagedPageMemoryRefusalAndCancellationStopBeforeSpillOrOutput()
    {
        foreach (bool cancel in new[] { false, true })
        {
            string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-page-record-" + Guid.NewGuid().ToString("N"));
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget(PdfPageContentStaging.ChunkByteCount);
            if (cancel) pool.RefuseChange = (_, after) => {
                if (after > 0 && after < PdfPageContentStaging.ChunkByteCount) cancellation.Cancel();
                return false;
            };
            bool ended = false;
            IEnumerable<PdfPage> Pages()
            {
                try { yield return new PdfPage(100, 100, "q\nQ\n"); }
                finally { ended = true; }
            }
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            {
                using var output = new MemoryStream();
                void Write() => PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable(Pages()),
                    new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 }, null, cancellation.Token,
                    temporaryStorage: new OoxPdfFileTemporaryStorage(directory)).GetAwaiter().GetResult();
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Write).CancellationToken);
                else TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(Write).Kind);
                TestAssert.True(ended, "Failure must dispose the page producer.");
                TestAssert.True(!Directory.Exists(directory), "Refusal/cancellation cannot choose spilling as recovery.");
                TestAssert.Equal(0L, output.Length);
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void StagedPageMemoryIncludesVariableAnnotationReferences()
    {
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            PdfPage plain = new(100, 100, "q\nQ\n");
            var links = Enumerable.Range(0, 1000).Select(index => PdfLinkAnnotation.ToUri(0, 0, 1, 1, "https://example.invalid/target")).ToArray();
            long Observe(PdfPage page)
            {
                using var staged = PdfDocumentWriter.ProduceStagedPagesAsync(AsyncEnumerable.ToAsyncEnumerable<PdfPage>([page]),
                    new(), null, default).GetAwaiter().GetResult();
                return pool.Reserved;
            }
            long baseline = Observe(plain);
            PdfPage annotated = new(100, 100, plain.Content, [], [], [], [], [], links);
            long withLinks = Observe(annotated);
            TestAssert.True(withLinks >= baseline + 4L * links.Length, "Annotation reference/numbering capacity is distinct from the page's fixed record.");
            TestAssert.Equal(0L, pool.Reserved);
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }
}
