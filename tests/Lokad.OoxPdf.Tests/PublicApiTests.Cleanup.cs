using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PublicApiTests
{
    private const string CleanupExceptionKey = "Lokad.OoxPdf.CleanupException";

    public static void ConvertAsyncStorageCleanupPreservesCancellationAndInvalidBacking()
    {
        foreach (bool cancel in new[] { true, false })
        foreach (bool cleanupThrows in new[] { false, true })
        {
            var pool = new TestMemoryBudget();
            var create = new EmissionGate("create");
            var dispose = new EmissionGate("dispose");
            var provider = new OwnedTemporaryStorage(create, shape: cancel ? null : "nonempty");
            var cleanup = new IOException("cleanup failed");
            using var cancellation = new CancellationTokenSource();
            using var input = new MemoryStream(ReadMinimalDocxBytes("<w:p><w:r><w:t>X</w:t></w:r></w:p>"));
            using var output = new MemoryStream();
            Task conversion = OoxPdfConverter.ConvertAsync(input, output, StorageOptions(pool, provider), cancellation.Token);
            try
            {
                create.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                var backing = TestAssert.NotNull(provider.Created);
                backing.CleanupFailure = cleanupThrows ? cleanup : null;
                backing.CleanupGate = dispose;
                if (cancel) cancellation.Cancel();
                create.Release.TrySetResult();
                dispose.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                TestAssert.True(!conversion.IsCompleted && pool.Reserved > 0, "A primary failure must await owned disposal before completion or refund.");
                TestAssert.Equal(0L, output.Length);
                dispose.Release.TrySetResult();
                Exception primary = TestAssert.Throws<Exception>(() => conversion.GetAwaiter().GetResult());
                if (cancel) TestAssert.Equal(cancellation.Token, ((OperationCanceledException)primary).CancellationToken);
                else TestAssert.True(primary is InvalidOperationException && primary.Message.Contains("empty", StringComparison.Ordinal), "Invalid backing retains its validation failure.");
                AssertCleanupMetadata(primary, cleanupThrows ? cleanup : null);
                AssertCleanupOwnership(pool, provider);
                TestAssert.True(input.CanRead && output.CanWrite, "Caller streams remain open after crossed failures.");
            }
            finally { create.Release.TrySetResult(); dispose.Release.TrySetResult(); try { conversion.GetAwaiter().GetResult(); } catch (Exception) { } }
        }
    }

    public static void PdfStagingCleanupPreservesProducerRefusalAndWriteFailures()
    {
        foreach (string phase in new[] { "producer", "write", "observer", "cancellation", "refusal" })
        foreach (bool cleanupThrows in new[] { false, true })
        {
            var pool = new TestMemoryBudget();
            var create = new EmissionGate("create");
            var io = new IOException("primary I/O failed");
            var observer = new InvalidOperationException("primary observer failed");
            var cleanup = new IOException("cleanup failed");
            var provider = new OwnedTemporaryStorage(create, failurePhase: phase == "write" ? "write" : null, failure: io);
            using var cancellation = new CancellationTokenSource();
            using var output = new MemoryStream();
            using (var scope = OoxConversionBudget.BeginScope(null, pool))
            {
                IEnumerable<PdfPage> Pages()
                {
                    yield return new PdfPage(100, 100, "q\nQ\n" + new string(' ', 64));
                    if (phase == "producer") throw io;
                    if (phase == "cancellation") cancellation.Cancel();
                    if (phase == "refusal") pool.RefuseChange = (before, after) => after > before;
                    yield return new PdfPage(100, 100, "q\nQ\n");
                }
                void Observe(OoxPdfDiagnostic diagnostic) { if (phase == "observer") throw observer; }
                Task write = PdfDocumentWriter.WriteStagedAsync(output, AsyncEnumerable.ToAsyncEnumerable<PdfPage>(Pages()),
                    new() { MaxResidentPageContentBytesPerConversion = 0 }, Observe, cancellation.Token, temporaryStorage: provider).AsTask();
                try
                {
                    create.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                    TestAssert.NotNull(provider.Created).CleanupFailure = cleanupThrows ? cleanup : null;
                    create.Release.TrySetResult();
                    Exception primary = TestAssert.Throws<Exception>(() => write.GetAwaiter().GetResult());
                    if (phase is "producer" or "write") TestAssert.True(ReferenceEquals(io, primary), "I/O exception identity survives cleanup.");
                    else if (phase == "observer") TestAssert.True(ReferenceEquals(observer, primary), "Producer observers retain their exception identity.");
                    else if (phase == "cancellation") TestAssert.Equal(cancellation.Token, ((OperationCanceledException)primary).CancellationToken);
                    else TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, ((OoxPdfLimitExceededException)primary).Kind);
                    AssertCleanupMetadata(primary, cleanupThrows ? cleanup : null);
                    TestAssert.Equal(0L, output.Length);
                    TestAssert.Equal(0L, pool.Reserved);
                }
                finally { create.Release.TrySetResult(); try { write.GetAwaiter().GetResult(); } catch (Exception) { } }
            }
            AssertCleanupOwnership(pool, provider);
        }
    }

    public static void ConvertAsyncStorageCleanupPreservesEmissionFailuresAndFileDestination()
    {
        foreach (bool file in new[] { false, true })
        foreach (bool cleanupThrows in new[] { false, true })
        {
            string? directory = file ? Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-cleanup-" + Guid.NewGuid().ToString("N"))).FullName : null;
            var pool = new TestMemoryBudget();
            var create = new EmissionGate("create");
            var primary = new IOException("emission failed");
            var earlier = new IOException("earlier cleanup failed");
            if (!file && cleanupThrows) primary.Data[CleanupExceptionKey] = earlier;
            var cleanup = new IOException("cleanup failed");
            var provider = new OwnedTemporaryStorage(create, failurePhase: file ? "read" : null, failure: primary);
            using var input = new MemoryStream(ReadMinimalDocxBytes("<w:p><w:r><w:t>X</w:t></w:r></w:p>"));
            using var output = new AsyncEmissionStream(failurePhase: "write", failure: primary);
            string? source = file ? Path.Combine(directory!, "input.docx") : null;
            string? destination = file ? Path.Combine(directory!, "output.pdf") : null;
            if (file) { File.WriteAllBytes(source!, input.ToArray()); File.WriteAllText(destination!, "existing destination"); }
            Task conversion = file ? OoxPdfConverter.ConvertAsync(source!, destination!, StorageOptions(pool, provider))
                : OoxPdfConverter.ConvertAsync(input, output, StorageOptions(pool, provider));
            try
            {
                create.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                TestAssert.NotNull(provider.Created).CleanupFailure = cleanupThrows ? cleanup : null;
                create.Release.TrySetResult();
                TestAssert.True(ReferenceEquals(primary, TestAssert.Throws<IOException>(() => conversion.GetAwaiter().GetResult())), "Emission retains the original exception object.");
                if (!file && cleanupThrows)
                {
                    var combined = (AggregateException)primary.Data[CleanupExceptionKey]!;
                    TestAssert.True(combined.InnerExceptions.SequenceEqual(new Exception[] { earlier, cleanup }), "Earlier cleanup information is preserved in order.");
                }
                else AssertCleanupMetadata(primary, cleanupThrows ? cleanup : null);
                if (file)
                {
                    TestAssert.Equal("existing destination", File.ReadAllText(destination!));
                    TestAssert.Equal(0, Directory.GetFiles(directory!, "output.pdf.tmp-*").Length);
                }
                TestAssert.Equal(0, output.AsyncDisposals + output.SyncDisposals);
                AssertCleanupOwnership(pool, provider);
            }
            finally { create.Release.TrySetResult(); try { conversion.GetAwaiter().GetResult(); } catch (Exception) { } if (directory is not null) Directory.Delete(directory, recursive: true); }
        }
    }

    public static void ConvertAsyncStorageCleanupKeepsObserverExceptionIdentity()
    {
        foreach (bool cleanupThrows in new[] { false, true })
        {
            var pool = new TestMemoryBudget(); var create = new EmissionGate("create");
            var provider = new OwnedTemporaryStorage(create);
            var primary = new InvalidOperationException("summary failed"); primary.Data["context"] = "retained";
            var cleanup = new IOException("cleanup failed");
            using var input = new MemoryStream(ReadMinimalDocxBytes("<w:p><w:r><w:t>X</w:t></w:r></w:p>"));
            using var output = new MemoryStream();
            void Observe(OoxPdfDiagnostic diagnostic) { if (diagnostic.Id == "CONVERSION_RESOURCE_SUMMARY") throw primary; }
            Task conversion = OoxPdfConverter.ConvertAsync(input, output, StorageOptions(pool, provider, summary: true, sink: Observe));
            try
            {
                create.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                TestAssert.NotNull(provider.Created).CleanupFailure = cleanupThrows ? cleanup : null;
                create.Release.TrySetResult();
                TestAssert.True(ReferenceEquals(primary, TestAssert.Throws<InvalidOperationException>(() => conversion.GetAwaiter().GetResult())), "Observer unwrapping preserves identity even after cleanup fails.");
                TestAssert.Equal("retained", primary.Data["context"]);
                AssertCleanupMetadata(primary, cleanupThrows ? cleanup : null);
                TestAssert.Equal(0L, output.Length);
                AssertCleanupOwnership(pool, provider);
            }
            finally { create.Release.TrySetResult(); try { conversion.GetAwaiter().GetResult(); } catch (Exception) { } }
        }
    }

    public static void ConvertAsyncStorageStandaloneCleanupFailurePreventsFilePublication()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        {
            string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-cleanup-only-" + Guid.NewGuid().ToString("N"))).FullName;
            var pool = new TestMemoryBudget(); var create = new EmissionGate("create");
            var provider = new OwnedTemporaryStorage(create);
            var cleanup = new IOException("only cleanup failed");
            string source = Path.Combine(directory, kind == OoxPdfInputKind.Docx ? "input.docx" : "input.pptx");
            string output = Path.Combine(directory, "output.pdf");
            File.WriteAllBytes(source, kind == OoxPdfInputKind.Docx ? ReadMinimalDocxBytes("<w:p><w:r><w:t>X</w:t></w:r></w:p>") : ReadMinimalPptxBytes());
            File.WriteAllText(output, "existing destination");
            Task conversion = OoxPdfConverter.ConvertAsync(source, output, StorageOptions(pool, provider, kind));
            try
            {
                create.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                TestAssert.NotNull(provider.Created).CleanupFailure = cleanup;
                create.Release.TrySetResult();
                TestAssert.True(ReferenceEquals(cleanup, TestAssert.Throws<IOException>(() => conversion.GetAwaiter().GetResult())), "A sole cleanup failure must still fail conversion.");
                AssertCleanupMetadata(cleanup, null);
                TestAssert.Equal("existing destination", File.ReadAllText(output));
                TestAssert.Equal(0, Directory.GetFiles(directory, "output.pdf.tmp-*").Length);
                AssertCleanupOwnership(pool, provider);
            }
            finally { create.Release.TrySetResult(); try { conversion.GetAwaiter().GetResult(); } catch (Exception) { } Directory.Delete(directory, recursive: true); }
        }
    }

    private static void AssertCleanupMetadata(Exception primary, Exception? cleanup)
    {
        if (cleanup is null) TestAssert.True(!primary.Data.Contains(CleanupExceptionKey), "Successful cleanup adds no secondary failure.");
        else TestAssert.True(ReferenceEquals(cleanup, primary.Data[CleanupExceptionKey]), "The original cleanup exception remains available to the caller.");
    }

    private static void AssertCleanupOwnership(TestMemoryBudget pool, OwnedTemporaryStorage provider)
    {
        var backing = TestAssert.NotNull(provider.Created);
        TestAssert.Equal(1, backing.AsyncDisposals);
        TestAssert.Equal(0, backing.SyncDisposals + provider.Disposals + pool.CallerDisposals);
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }
}
