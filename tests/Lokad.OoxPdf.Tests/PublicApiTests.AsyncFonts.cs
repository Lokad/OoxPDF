using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class PublicApiTests
{
    public static void ConvertAsyncDocxReturnsWhileFontAcquisitionIsPending()
    {
        byte[] inputBytes = ReadMinimalDocxBytes("<w:p><w:r><w:t>ABC</w:t></w:r></w:p>");
        var pool = new TestMemoryBudget(); using var source = new AwaitedFontSource();
        using var cancellation = new CancellationTokenSource();
        using var input = new MemoryStream(inputBytes); using var output = new AsyncEmissionStream();
        using var returned = new ManualResetEventSlim();
        bool blocked = false;
        // A blocking parent must produce a finite behavioral failure, not hang
        // the runner. This watchdog is test-only; conversion is called directly.
        Task watchdog = Task.Run(() => {
            source.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
            if (!returned.Wait(TimeSpan.FromSeconds(2))) { blocked = true; source.Release.TrySetResult(); }
        });
        Task? conversion = null;
        try
        {
            conversion = OoxPdfConverter.ConvertAsync(input, output, AwaitedFontOptions(source, pool), cancellation.Token);
            returned.Set(); watchdog.GetAwaiter().GetResult();
            TestAssert.True(!blocked, "ConvertAsync blocked while its selected font acquisition was pending.");
            TestAssert.True(!conversion.IsCompleted && pool.Reserved > 0 && output.Length == 0,
                "Conversion returns an incomplete task while its admitted font data is borrowed.");
            TestAssert.Equal(cancellation.Token, source.ObservedToken);
            source.Release.TrySetResult(); conversion.GetAwaiter().GetResult();
            using var baselineInput = new MemoryStream(inputBytes); using var baseline = new MemoryStream();
            OoxPdfConverter.Convert(baselineInput, baseline, AwaitedFontOptions(new MemoryFontProgramSource(source.StableId, source.Template), null));
            TestAssert.True(output.Bytes().AsSpan().SequenceEqual(baseline.ToArray()), "Awaited font acquisition preserves PDF bytes.");
            TestAssert.Equal(1, source.Calls); TestAssert.Equal(1, source.LeasesDisposed);
            TestAssert.Equal(0L, pool.Reserved); TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, source.Disposals + pool.CallerDisposals + output.AsyncDisposals + output.SyncDisposals);
            TestAssert.True(input.CanRead && output.CanWrite, "Caller streams remain open.");
        }
        finally
        {
            source.Release.TrySetResult(); returned.Set();
            if (conversion is not null) conversion.GetAwaiter().GetResult();
            watchdog.GetAwaiter().GetResult();
        }
    }

    public static void ConvertAsyncDocxFontCancellationWaitsForTheAcquisitionBorrow()
    {
        var pool = new TestMemoryBudget(); using var source = new AwaitedFontSource();
        using var input = new MemoryStream(ReadMinimalDocxBytes("<w:p><w:r><w:t>ABC</w:t></w:r></w:p>"));
        using var output = new MemoryStream(); using var cancellation = new CancellationTokenSource();
        Task conversion = OoxPdfConverter.ConvertAsync(input, output, AwaitedFontOptions(source, pool), cancellation.Token);
        try
        {
            source.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
            cancellation.Cancel();
            TestAssert.True(!conversion.IsCompleted && pool.Reserved >= source.Template.Length,
                "Cancellation cannot abandon a source that still owns admitted acquisition storage.");
            source.Release.TrySetResult();
            TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => conversion.GetAwaiter().GetResult()).CancellationToken);
            TestAssert.Equal(1, source.LeasesDisposed); TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed); TestAssert.Equal(0L, output.Length);
            TestAssert.Equal(0, source.Disposals + pool.CallerDisposals);
            TestAssert.True(input.CanRead && output.CanWrite, "Cancellation preserves caller ownership.");
        }
        finally { source.Release.TrySetResult(); try { conversion.GetAwaiter().GetResult(); } catch (OperationCanceledException) { } }
    }

    public static void ConvertAsyncDocxFontCopyRefusalReleasesLeaseAndPreservesPublication()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-font-publication-" + Guid.NewGuid().ToString("N"))).FullName;
        string inputPath = Path.Combine(directory, "input.docx"), outputPath = Path.Combine(directory, "output.pdf");
        File.WriteAllBytes(inputPath, ReadMinimalDocxBytes("<w:p><w:r><w:t>ABC</w:t></w:r></w:p>"));
        File.WriteAllText(outputPath, "previous result");
        try
        {
            foreach (bool file in new[] { false, true })
            {
                var pool = new TestMemoryBudget();
                using var source = new AwaitedFontSource(() => pool.RefuseChange = (_, bytes) => bytes > 0);
                using var input = File.OpenRead(inputPath); using var output = new MemoryStream();
                Task conversion = file
                    ? OoxPdfConverter.ConvertAsync(inputPath, outputPath, AwaitedFontOptions(source, pool))
                    : OoxPdfConverter.ConvertAsync(input, output, AwaitedFontOptions(source, pool));
                source.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                source.Release.TrySetResult();
                var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => conversion.GetAwaiter().GetResult());
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind); TestAssert.Equal(OoxPdfLimitScope.Conversion, failure.Scope);
                TestAssert.Equal(1, source.Calls); TestAssert.Equal(1, source.LeasesDisposed);
                TestAssert.Equal(0L, pool.Reserved); TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.Equal("previous result", File.ReadAllText(outputPath)); TestAssert.Equal(0L, output.Length);
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                TestAssert.Equal(0, source.Disposals + pool.CallerDisposals);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static OoxPdfOptions AwaitedFontOptions(IFontProgramSource source, IMemoryBudget? pool) => new() {
        InputKind = OoxPdfInputKind.Docx, FontResolver = new AwaitedFontResolver(source),
        MemoryBudget = pool, FixedCreationDate = DateTimeOffset.UnixEpoch };

    private sealed class AwaitedFontResolver(IFontProgramSource source) : IFontResolver
    {
        public FontFaceResolution Resolve(FontRequest request) => new(request.FamilyName, "TestFont", new FontStyleKey(), source, false);
    }

    private sealed class AwaitedFontSource(Action? afterAcquisition = null) : IFontProgramSource, IDisposable
    {
        internal readonly byte[] Template = TestFontBuilder.CreateTestFont();
        internal readonly TaskCompletionSource Entered = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal readonly TaskCompletionSource Release = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal int Calls, LeasesDisposed, Disposals; internal CancellationToken ObservedToken;
        public string StableId => "awaited-font";
        public async ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)
        {
            Calls++; ObservedToken = ct;
            IMemoryReservation? reservation = OoxMemory.Reserve(operationBudget, Template.Length);
            FontProgramBuffer program;
            try { program = new FontProgramBuffer(Template.ToArray(), reservation); reservation = null; }
            finally { reservation?.Dispose(); }
            using (program)
            {
                Entered.TrySetResult();
                await Release.Task.ConfigureAwait(false); // The outstanding borrow deliberately ignores cancellation.
                afterAcquisition?.Invoke();
                return new ObservedFontLease(program.Acquire(), () => LeasesDisposed++);
            }
        }
        public void Dispose() => Disposals++;
    }

    private sealed class ObservedFontLease(IOoxPdfFontProgramLease inner, Action released) : IOoxPdfFontProgramLease
    {
        private IOoxPdfFontProgramLease? lease = inner;
        public ReadOnlyMemory<byte> Bytes => lease!.Bytes;
        public void Dispose() { if (lease is null) return; lease.Dispose(); lease = null; released(); }
    }
}
