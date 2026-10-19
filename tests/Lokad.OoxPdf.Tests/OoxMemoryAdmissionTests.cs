using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static class OoxMemoryAdmissionTests
{
    private const long ReservationBytes = 128L * 1024 * 1024;

    public static void MemoryAdmissionRefusalLeavesFileAndStreamUntouched()
    {
        var pool = new Pool(0);
        var admission = new Admission(pool);
        foreach (string extension in new[] { ".docx", ".pptx" })
        {
            string input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            File.WriteAllBytes(input, [1, 2, 3]); // Refusal must precede ZIP parsing.
            string output = Path.ChangeExtension(input, ".pdf");
            File.WriteAllBytes(output, [7, 8, 9]);
            TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output,
                new OoxPdfOptions { MemoryAdmission = admission }));
            TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Refusal preserves the destination.");
            TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);

            using var source = new MemoryStream([1, 2, 3]);
            using var destination = new MemoryStream();
            TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(source, destination,
                new OoxPdfOptions { InputKind = extension == ".docx" ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx, MemoryAdmission = admission }));
            TestAssert.Equal(0L, source.Position);
            TestAssert.Equal(0L, destination.Length);
        }
        TestAssert.Equal(4, admission.Requests);
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(0, pool.Releases);
    }

    public static void MemoryAdmissionCoversRenderingStagingAndEmission()
    {
        foreach (var (input, kind) in new[] {
            (Case("Composition/report.docx"), OoxPdfInputKind.Docx),
            (Case("pptx-ladder-07-image-crop.pptx"), OoxPdfInputKind.Pptx) })
        {
            var pool = new Pool(ReservationBytes);
            var admission = new Admission(pool);
            var resolver = new ObservedResolver(() => TestAssert.Equal(ReservationBytes, pool.Reserved));
            using var source = File.OpenRead(input);
            using var destination = new ObservedOutput(() => TestAssert.Equal(ReservationBytes, pool.Reserved));
            OoxPdfConverter.Convert(source, destination, new OoxPdfOptions {
                InputKind = kind, FontResolver = resolver, MemoryAdmission = admission,
                ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 },
                DiagnosticSink = _ => TestAssert.Equal(ReservationBytes, pool.Reserved), ReportResourceUsage = true });
            if (kind == OoxPdfInputKind.Docx) { TestAssert.True(resolver.Calls > 0, "Text-bearing DOCX exercises font preparation."); }
            TestAssert.True(destination.Writes > 0, "Staged emission was observed.");
            if (kind == OoxPdfInputKind.Pptx) { TestAssert.Contains("/Subtype /Image", System.Text.Encoding.Latin1.GetString(destination.ToArray())); }
            TestAssert.Equal(1, admission.Requests);
            TestAssert.Equal(1, pool.Releases);
            TestAssert.Equal(0L, pool.Reserved);
        }
    }

    public static void MemoryAdmissionRefundsOnFailureCancellationAndObserverException()
    {
        foreach (string failure in new[] { "parse", "font-cancel", "observer", "write", "entry-cancel" })
        {
            var pool = new Pool(ReservationBytes);
            using var cancellation = new CancellationTokenSource();
            var admission = new Admission(pool, failure == "entry-cancel" ? cancellation.Cancel : null);
            var resolver = new ObservedResolver(() => {
                TestAssert.Equal(ReservationBytes, pool.Reserved);
                if (failure == "font-cancel") { cancellation.Cancel(); }
            });
            using Stream source = failure == "parse" ? new MemoryStream([1, 2, 3]) : File.OpenRead(Case("Composition/report.docx"));
            using var destination = new ObservedOutput(() => {
                TestAssert.Equal(ReservationBytes, pool.Reserved);
                if (failure == "write") { throw new IOException("Injected output failure."); }
            });
            var options = new OoxPdfOptions {
                InputKind = OoxPdfInputKind.Docx, FontResolver = resolver, MemoryAdmission = admission,
                ReportResourceUsage = true, DiagnosticSink = d => {
                    TestAssert.Equal(ReservationBytes, pool.Reserved);
                    if (failure == "observer" && d.Id == "CONVERSION_RESOURCE_SUMMARY") { throw new InvalidOperationException("Injected observer failure."); }
                } };
            Exception? caught = null;
            try { OoxPdfConverter.Convert(source, destination, options, cancellation.Token); }
            catch (Exception exception) { caught = exception; }
            TestAssert.True(caught is not null, "The injected failure must escape: " + failure);
            if (failure is "font-cancel" or "entry-cancel") { TestAssert.True(caught is OperationCanceledException, "Cancellation retains its exception type."); }
            if (failure != "write") { TestAssert.Equal(0L, destination.Length); }
            TestAssert.Equal(1, admission.Requests);
            TestAssert.Equal(1, pool.Releases);
            TestAssert.Equal(0L, pool.Reserved);
        }
    }

    public static void MemoryAdmissionSharedCapacityRefusesConcurrentConversionAndRefunds()
    {
        var pool = new Pool(ReservationBytes);
        var admission = new Admission(pool);
        using var entered = new ManualResetEventSlim();
        using var resume = new ManualResetEventSlim();
        byte[] bytes = File.ReadAllBytes(Case("Composition/report.docx"));
        var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, MemoryAdmission = admission,
            FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true) };
        using var firstOutput = new MemoryStream();
        Task first = Task.Run(() => {
            using var source = new PausedInput(bytes, entered, resume);
            OoxPdfConverter.Convert(source, firstOutput, options);
        });
        try
        {
            TestAssert.True(entered.Wait(TimeSpan.FromSeconds(10)), "First conversion must enter its admitted input read.");
            using var secondInput = new MemoryStream(bytes);
            using var secondOutput = new MemoryStream();
            TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(secondInput, secondOutput, options));
            TestAssert.Equal(0L, secondInput.Position);
            TestAssert.Equal(0L, secondOutput.Length);
            TestAssert.Equal(ReservationBytes, pool.Reserved);
        }
        finally { resume.Set(); }
        first.GetAwaiter().GetResult();
        TestAssert.Equal(0L, pool.Reserved);
        using var retryInput = new MemoryStream(bytes);
        using var retryOutput = new MemoryStream();
        OoxPdfConverter.Convert(retryInput, retryOutput, options);
        TestAssert.True(firstOutput.ToArray().SequenceEqual(retryOutput.ToArray()), "Refunded capacity admits an identical retry.");
        TestAssert.Equal(3, admission.Requests);
        TestAssert.Equal(2, pool.Releases);
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void MemoryAdmissionNestedConversionsKeepIndependentOwnership()
    {
        var pool = new Pool(2 * ReservationBytes);
        var admission = new Admission(pool);
        bool nested = false;
        var resolver = new ObservedResolver(() => {
            TestAssert.Equal(ReservationBytes, pool.Reserved);
            if (nested) { return; }
            nested = true;
            using var source = File.OpenRead(Case("Composition/report.docx"));
            using var destination = new ObservedOutput(() => TestAssert.Equal(2 * ReservationBytes, pool.Reserved));
            OoxPdfConverter.Convert(source, destination, new OoxPdfOptions {
                InputKind = OoxPdfInputKind.Docx, MemoryAdmission = admission,
                FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true) });
            TestAssert.Equal(ReservationBytes, pool.Reserved);
        });
        using var input = File.OpenRead(Case("Composition/report.docx"));
        using var output = new MemoryStream();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions {
            InputKind = OoxPdfInputKind.Docx, FontResolver = resolver, MemoryAdmission = admission });
        TestAssert.True(nested, "Nested conversion must execute inside font preparation.");
        TestAssert.Equal(2, admission.Requests);
        TestAssert.Equal(2, pool.Releases);
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void MemoryAdmissionReleaseFailurePreventsFilePublication()
    {
        var pool = new Pool(ReservationBytes);
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        File.WriteAllBytes(output, [7, 8, 9]);
        var admission = new Admission(pool, onRelease: () => {
            TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Admission releases before publication.");
            throw new InvalidOperationException("Host violated the nonthrowing release contract.");
        });
        TestAssert.Throws<InvalidOperationException>(() => OoxPdfConverter.Convert(Case("Composition/report.docx"), output,
            new OoxPdfOptions { MemoryAdmission = admission, FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true) }));
        TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Release failure preserves the destination.");
        TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);
        TestAssert.Equal(1, pool.Releases);
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void MemoryAdmissionPreCancellationDoesNotAcquireAndAsyncRefunds()
    {
        var pool = new Pool(ReservationBytes);
        var admission = new Admission(pool);
        var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, MemoryAdmission = admission,
            FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true) };
        using var source = File.OpenRead(Case("Composition/report.docx"));
        using var destination = new MemoryStream();
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(source, destination, options, cancelled.Token));
        TestAssert.Equal(0, admission.Requests);
        OoxPdfConverter.ConvertAsync(source, destination, options).GetAwaiter().GetResult();
        TestAssert.True(destination.Length > 0, "Async conversion uses the same admission contract.");
        TestAssert.Equal(1, admission.Requests);
        TestAssert.Equal(1, pool.Releases);
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void StagedDocumentAndContentStoreReleaseOwnedReferencesOnDispose()
    {
        var page = new PdfPage(100d, 100d, "q\nQ\n");
        var staged = PdfDocumentWriter.ProduceStagedPages([page], new(), null, default);
        var storage = staged.Staging;
        TestAssert.Equal(1, staged.Pages.Count);
        staged.Dispose();
        staged.Dispose();
        TestAssert.Throws<ObjectDisposedException>(() => _ = staged.Pages);
        TestAssert.Throws<ObjectDisposedException>(() => storage.GetPageLength(0, default));
    }

    private static string Case(string relative) => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
        "..", "..", "..", "Cases", relative));

    private sealed class ObservedResolver(Action observe) : IFontResolver
    {
        private readonly IFontResolver inner = new PortableRenderingTests.PinnedResolver(mathFont: true);
        public int Calls;
        public FontFaceResolution Resolve(FontRequest request) { Calls++; observe(); return inner.Resolve(request); }
    }

    private sealed class ObservedOutput(Action observe) : MemoryStream
    {
        public int Writes;
        public override void Write(byte[] buffer, int offset, int count) { Writes++; observe(); base.Write(buffer, offset, count); }
        public override void Write(ReadOnlySpan<byte> buffer) { Writes++; observe(); base.Write(buffer); }
        public override void WriteByte(byte value) { Writes++; observe(); base.WriteByte(value); }
    }

    private sealed class PausedInput(byte[] bytes, ManualResetEventSlim entered, ManualResetEventSlim resume) : MemoryStream(bytes)
    {
        private int paused;
        private void PauseOnce()
        {
            if (Interlocked.Exchange(ref paused, 1) != 0) { return; }
            entered.Set();
            if (!resume.Wait(TimeSpan.FromSeconds(10))) { throw new TimeoutException("Concurrent admission probe did not resume."); }
        }
        public override int Read(byte[] buffer, int offset, int count) { PauseOnce(); return base.Read(buffer, offset, count); }
        public override int Read(Span<byte> buffer) { PauseOnce(); return base.Read(buffer); }
    }

    // A host adapter, deliberately outside the library. Its single coarse amount
    // includes storage and peak headroom; there is no per-object accounting.
    private sealed class Admission(Pool pool, Action? onAcquire = null, Action? onRelease = null) : IOoxPdfMemoryAdmission
    {
        public int Requests;
        public IDisposable? TryAcquire(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Interlocked.Increment(ref Requests);
            IDisposable? lease = pool.TryAcquire(onRelease);
            if (lease is not null) { onAcquire?.Invoke(); }
            return lease;
        }
    }

    private sealed class Pool(long capacity)
    {
        private readonly object sync = new();
        private long reserved;
        public int Releases;
        public long Reserved { get { lock (sync) { return reserved; } } }
        public IDisposable? TryAcquire(Action? onRelease)
        {
            lock (sync)
            {
                if (ReservationBytes > capacity - reserved) { return null; }
                reserved += ReservationBytes;
                return new Lease(this, onRelease);
            }
        }
        private sealed class Lease(Pool pool, Action? onRelease) : IDisposable
        {
            private Pool? owner = pool;
            public void Dispose()
            {
                Pool? held = Interlocked.Exchange(ref owner, null);
                if (held is null) { return; }
                lock (held.sync) { held.reserved -= ReservationBytes; held.Releases++; }
                onRelease?.Invoke();
            }
        }
    }
}
