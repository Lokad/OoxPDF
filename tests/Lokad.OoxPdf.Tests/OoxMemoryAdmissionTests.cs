using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{

    public static void MemoryAdmissionCallbackCannotRedirectRelativeInputOrOutput()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-admission-path-" + Guid.NewGuid().ToString("N"))).FullName;
        string elsewhere = Directory.CreateDirectory(Path.Combine(directory, "elsewhere")).FullName;
        string input = Path.Combine(directory, "source.docx");
        byte[] source = File.ReadAllBytes(Case("docx-ladder-00-blank.docx"));
        File.WriteAllBytes(input, source);
        var pool = new TestMemoryBudget(onCreate: () => Environment.CurrentDirectory = elsewhere);
        string previousDirectory = Environment.CurrentDirectory;
        try
        {
            Environment.CurrentDirectory = directory;
            OoxPdfConverter.Convert("source.docx", "result.pdf", new OoxPdfOptions { MemoryBudget = pool });
            TestAssert.Contains("%PDF-", System.Text.Encoding.ASCII.GetString(File.ReadAllBytes(Path.Combine(directory, "result.pdf"))));
            TestAssert.True(!File.Exists(Path.Combine(elsewhere, "result.pdf")), "Admission cannot redirect the destination.");
            TestAssert.True(source.SequenceEqual(File.ReadAllBytes(input)), "Admission must preserve input.");
            TestAssert.Equal(1, pool.RootsCreated);
            TestAssert.Equal(1, pool.RootsDisposed);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
        }
        finally
        {
            Environment.CurrentDirectory = previousDirectory;
            Directory.Delete(directory, recursive: true);
        }
    }

    public static void MemoryAdmissionRefusalLeavesFileAndStreamUntouched()
    {
        var pool = new TestMemoryBudget(0);
        foreach (string extension in new[] { ".docx", ".pptx" })
        {
            string input = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            File.WriteAllBytes(input, [1, 2, 3]); // Refusal must precede ZIP parsing.
            string output = Path.ChangeExtension(input, ".pdf");
            File.WriteAllBytes(output, [7, 8, 9]);
            var fileFailure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output,
                new OoxPdfOptions { MemoryBudget = pool }));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, fileFailure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, fileFailure.Scope);
            TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Refusal preserves the destination.");
            TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);

            using var source = new MemoryStream([1, 2, 3]);
            using var destination = new MemoryStream();
            var streamFailure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(source, destination,
                new OoxPdfOptions { InputKind = extension == ".docx" ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx, MemoryBudget = pool }));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, streamFailure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, streamFailure.Scope);
            TestAssert.Equal(0L, source.Position);
            TestAssert.Equal(0L, destination.Length);
        }
        TestAssert.Equal(4, pool.RootsCreated);
        TestAssert.Equal(4, pool.RootsDisposed);
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void MemoryAdmissionCoversRenderingStagingAndEmission()
    {
        foreach (var (input, kind) in new[] {
            (Case("Composition/report.docx"), OoxPdfInputKind.Docx),
            (Case("pptx-ladder-07-image-crop.pptx"), OoxPdfInputKind.Pptx) })
        {
            var pool = new TestMemoryBudget();
            bool observe = true;
            void Observe() { if (observe) TestAssert.True(pool.Reserved > 0, "Package source bytes remain funded while consumers read them."); }
            var resolver = new ObservedResolver(Observe);
            using var source = File.OpenRead(input);
            using var destination = new ObservedOutput(() => TestAssert.True(pool.Reserved > 0, "Package source bytes remain funded while consumers read them."));
            OoxPdfConverter.Convert(source, destination, new OoxPdfOptions {
                InputKind = kind, FontResolver = resolver, MemoryBudget = pool,
                ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 },
                DiagnosticSink = _ => TestAssert.True(pool.Reserved > 0, "Package source bytes remain funded while consumers read them."), ReportResourceUsage = true });
            if (kind == OoxPdfInputKind.Docx) { TestAssert.True(resolver.Calls > 0, "Text-bearing DOCX exercises font preparation."); }
            TestAssert.True(destination.Writes > 0, "Staged emission was observed.");
            if (kind == OoxPdfInputKind.Pptx) { TestAssert.Contains("/Subtype /Image", System.Text.Encoding.Latin1.GetString(destination.ToArray())); }
            TestAssert.Equal(1, pool.RootsCreated);
            TestAssert.Equal(1, pool.RootsDisposed);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0L, pool.Reserved);
            observe = false;
            using var ungovernedInput = File.OpenRead(input);
            using var ungovernedOutput = new MemoryStream();
            OoxPdfConverter.Convert(ungovernedInput, ungovernedOutput, new OoxPdfOptions {
                InputKind = kind, FontResolver = resolver,
                ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = 0 } });
            TestAssert.True(destination.ToArray().SequenceEqual(ungovernedOutput.ToArray()),
                "Admission does not change PDF bytes for identical explicit fonts and settings.");
            TestAssert.Equal(1, pool.RootsCreated);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void MemoryBudgetPackageStorageAndConstructionHaveSeparateLifetimes()
    {
        string path = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, byte[]> {
            ["[Content_Types].xml"] = TestFixtures.Utf8("""
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="bin" ContentType="application/octet-stream"/>
                </Types>
                """),
            ["padding.bin"] = new byte[2 * 1024 * 1024]
        });
        try
        {
            foreach (bool seekable in new[] { false, true })
            {
                var pool = new TestMemoryBudget();
                using (var conversion = OoxConversionBudget.BeginScope(null, pool))
                {
                    TestAssert.Equal(0L, pool.Reserved);
                    using Stream input = seekable ? File.OpenRead(path) : new ForwardInput(File.ReadAllBytes(path));
                    var package = OoxPackage.Open(input, default);
                    long retained = package.Parts.Sum(part => part.Bytes.LongLength);
                    TestAssert.Equal(retained, pool.Reserved);
                    TestAssert.True(pool.Peak >= 2 * retained, "The construction buffer and final array overlap before ownership transfer.");
                    package.Dispose();
                    package.Dispose();
                    TestAssert.Equal(0L, pool.Reserved);
                    TestAssert.True(input.CanRead, "Package disposal leaves its caller's input open.");
                    TestAssert.Throws<ObjectDisposedException>(() => package.GetPart("/padding.bin"));
                }
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.Equal(0, pool.CallerDisposals);
            }
        }
        finally { File.Delete(path); }
    }

    public static void MemoryBudgetGrowthRefusalIsTerminalAndRefundsAcquiredStorage()
    {
        var pool = new TestMemoryBudget { RefuseChange = (before, after) => before > 0 && after > before };
        using var input = File.OpenRead(Case("Composition/report.docx"));
        using var output = new MemoryStream();
        var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output,
            new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, MemoryBudget = pool }));
        TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
        TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
        TestAssert.True(input.Position > 0 && pool.Peak > 0, "Refusal occurs during real construction growth after intake has started.");
        TestAssert.Equal(0L, output.Length);
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }

    public static void MemoryBudgetNestedWorkbookRefusalEscapesChartRecovery()
    {
        var pool = new TestMemoryBudget();
        pool.RefuseChange = (_, after) => pool.ScopesCreated >= 3 && after > 0;
        var diagnostics = new List<Lokad.OoxPdf.Diagnostics.OoxPdfDiagnostic>();
        using var input = File.OpenRead(Case("pptx-ladder-11-chart-bar-clustered-port.pptx"));
        using var output = new MemoryStream();
        var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output,
            new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, MemoryBudget = pool,
                FontResolver = new PortableRenderingTests.PinnedResolver(), DiagnosticSink = diagnostics.Add }));
        TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
        TestAssert.True(diagnostics.All(item => item.Id != "PPTX_NODE_RENDER_FAILED"), "Shared refusal must not omit a chart and report success.");
        TestAssert.Equal(0L, output.Length);
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }

    public static void MemoryAdmissionRefundsOnFailureCancellationAndObserverException()
    {
        foreach (string failure in new[] { "parse", "font-cancel", "observer", "write", "entry-cancel" })
        {
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget(onCreate: failure == "entry-cancel" ? cancellation.Cancel : null);
            var resolver = new ObservedResolver(() => {
                TestAssert.True(pool.Reserved > 0, "Package source bytes remain funded while consumers read them.");
                if (failure == "font-cancel") { cancellation.Cancel(); }
            });
            using Stream source = failure == "parse" ? new MemoryStream([1, 2, 3]) : File.OpenRead(Case("Composition/report.docx"));
            using var destination = new ObservedOutput(() => {
                TestAssert.True(pool.Reserved > 0, "Package source bytes remain funded while consumers read them.");
                if (failure == "write") { throw new IOException("Injected output failure."); }
            });
            var options = new OoxPdfOptions {
                InputKind = OoxPdfInputKind.Docx, FontResolver = resolver, MemoryBudget = pool,
                ReportResourceUsage = true, DiagnosticSink = d => {
                    TestAssert.True(pool.Reserved > 0, "Package source bytes remain funded while consumers read them.");
                    if (failure == "observer" && d.Id == "CONVERSION_RESOURCE_SUMMARY") { throw new InvalidOperationException("Injected observer failure."); }
                } };
            Exception? caught = null;
            try { OoxPdfConverter.Convert(source, destination, options, cancellation.Token); }
            catch (Exception exception) { caught = exception; }
            TestAssert.True(caught is not null, "The injected failure must escape: " + failure);
            if (failure is "font-cancel" or "entry-cancel") { TestAssert.True(caught is OperationCanceledException, "Cancellation retains its exception type."); }
            if (failure != "write") { TestAssert.Equal(0L, destination.Length); }
            TestAssert.Equal(1, pool.RootsCreated);
            TestAssert.Equal(1, pool.RootsDisposed);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0L, pool.Reserved);
        }
    }

    public static void MemoryAdmissionSharedCapacityRefusesConcurrentConversionAndRefunds()
    {
        var pool = new TestMemoryBudget();
        using var entered = new ManualResetEventSlim();
        using var resume = new ManualResetEventSlim();
        int paused = 0;
        var resolver = new ObservedResolver(() => {
            if (Interlocked.Exchange(ref paused, 1) != 0) return;
            // Leave less than one intake buffer free while this package is held.
            pool.Capacity = pool.Reserved + 81919;
            entered.Set();
            if (!resume.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("Shared budget probe did not resume.");
        });
        byte[] bytes = File.ReadAllBytes(Case("Composition/report.docx"));
        var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx,
            MemoryBudget = pool, FontResolver = resolver };
        using var firstOutput = new MemoryStream();
        Task first = Task.Run(() => {
            using var source = new MemoryStream(bytes);
            OoxPdfConverter.Convert(source, firstOutput, options);
        });
        try
        {
            TestAssert.True(entered.Wait(TimeSpan.FromSeconds(10)), "First conversion must retain its acquired package.");
            long held = pool.Reserved;
            using var secondInput = new MemoryStream(bytes);
            using var secondOutput = new MemoryStream();
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(secondInput, secondOutput, options));
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(0L, secondInput.Position);
            TestAssert.Equal(0L, secondOutput.Length);
            TestAssert.Equal(held, pool.Reserved);
        }
        finally { resume.Set(); }
        first.GetAwaiter().GetResult();
        TestAssert.Equal(0L, pool.Reserved);
        pool.Capacity = 128L * 1024 * 1024;
        using var retryInput = new MemoryStream(bytes);
        using var retryOutput = new MemoryStream();
        OoxPdfConverter.Convert(retryInput, retryOutput, options);
        TestAssert.True(firstOutput.ToArray().SequenceEqual(retryOutput.ToArray()), "Refunded capacity admits an identical retry.");
        TestAssert.Equal(3, pool.RootsCreated);
        TestAssert.Equal(3, pool.RootsDisposed);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void MemoryAdmissionNestedConversionsKeepIndependentOwnership()
    {
        var pool = new TestMemoryBudget();
        bool nested = false;
        var resolver = new ObservedResolver(() => {
            TestAssert.True(pool.Reserved > 0, "Package source bytes remain funded while consumers read them.");
            if (nested) { return; }
            nested = true;
            long outer = pool.Reserved;
            using var source = File.OpenRead(Case("Composition/report.docx"));
            using var destination = new ObservedOutput(() => TestAssert.True(pool.Reserved > outer,
                "Nested package ownership shares the root without replacing its parent."));
            OoxPdfConverter.Convert(source, destination, new OoxPdfOptions {
                InputKind = OoxPdfInputKind.Docx, MemoryBudget = pool,
                FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true) });
            TestAssert.Equal(outer, pool.Reserved);
        });
        using var input = File.OpenRead(Case("Composition/report.docx"));
        using var output = new MemoryStream();
        OoxPdfConverter.Convert(input, output, new OoxPdfOptions {
            InputKind = OoxPdfInputKind.Docx, FontResolver = resolver, MemoryBudget = pool });
        TestAssert.True(nested, "Nested conversion must execute inside font preparation.");
        TestAssert.Equal(2, pool.RootsCreated);
        TestAssert.Equal(2, pool.RootsDisposed);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void MemoryAdmissionReleaseFailurePreventsFilePublication()
    {
        string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        File.WriteAllBytes(output, [7, 8, 9]);
        var pool = new TestMemoryBudget(onDispose: () => {
            TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Admission releases before publication.");
            throw new InvalidOperationException("Host violated the nonthrowing release contract.");
        });
        TestAssert.Throws<InvalidOperationException>(() => OoxPdfConverter.Convert(Case("Composition/report.docx"), output,
            new OoxPdfOptions { MemoryBudget = pool, FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true) }));
        TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Release failure preserves the destination.");
        TestAssert.Equal(0, Directory.GetFiles(Path.GetDirectoryName(output)!, Path.GetFileName(output) + ".tmp-*").Length);
        TestAssert.Equal(1, pool.RootsDisposed);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void MemoryAdmissionPreCancellationDoesNotAcquireAndAsyncRefunds()
    {
        var pool = new TestMemoryBudget();
        var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, MemoryBudget = pool,
            FontResolver = new PortableRenderingTests.PinnedResolver(mathFont: true) };
        using var source = File.OpenRead(Case("Composition/report.docx"));
        using var destination = new MemoryStream();
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(source, destination, options, cancelled.Token));
        TestAssert.Equal(0, pool.RootsCreated);
        OoxPdfConverter.ConvertAsync(source, destination, options).GetAwaiter().GetResult();
        TestAssert.True(destination.Length > 0, "Async conversion uses the same admission contract.");
        TestAssert.Equal(1, pool.RootsCreated);
        TestAssert.Equal(1, pool.RootsDisposed);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
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
        TestAssert.Throws<ObjectDisposedException>(() => storage.BeginPage(default));
        TestAssert.Throws<ObjectDisposedException>(() => storage.AppendPageBytes([1], 0, 1, default));
        TestAssert.Throws<ObjectDisposedException>(() => storage.EndPage(default));
        TestAssert.Throws<ObjectDisposedException>(() => storage.ReadPageBytes(0, new byte[1], 0, 0, 1, default));
        TestAssert.Throws<ObjectDisposedException>(() => _ = storage.Scratch);
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

    private sealed class ForwardInput(byte[] bytes) : Stream
    {
        private readonly MemoryStream source = new(bytes);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => source.Read(buffer, offset, Math.Min(count, 37));
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) source.Dispose(); base.Dispose(disposing); }
    }

}
