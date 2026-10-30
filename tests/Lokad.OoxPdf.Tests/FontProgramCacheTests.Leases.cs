using System.Net;
using System.Security.Cryptography;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class FontProgramCacheTests
{
    public static void FontLeasesKeepEvictedReplacedAndShutdownStorageFunded()
    {
        var pool = new TestMemoryBudget();
        using var cache = new FontProgramCache(4, StringComparer.Ordinal, pool);
        IOoxPdfFontProgramLease Publish(string key, byte[] bytes)
        {
            using var program = new FontProgramBuffer(bytes, OoxMemory.Reserve(pool, bytes.Length));
            cache.Store(key, program, null);
            return program.Acquire();
        }
        using var original = Publish("a", [1, 2, 3, 4]);
        using var replacement = Publish("a", [5, 6, 7, 8]);
        TestAssert.Equal(8L, pool.Reserved);
        TestAssert.True(original.Bytes.Span.SequenceEqual(new byte[] { 1, 2, 3, 4 }), "Replacement preserves active readers.");
        using var next = Publish("b", [9, 10, 11, 12]);
        TestAssert.Equal(12L, pool.Reserved);
        TestAssert.Equal(4L, cache.RetainedBytes);
        cache.Dispose(); cache.Dispose();
        TestAssert.Equal(12L, pool.Reserved);
        TestAssert.Equal(0L, cache.RetainedBytes);
        TestAssert.Throws<ObjectDisposedException>(() => cache.TryAcquire("b", null, out _));
        original.Dispose(); original.Dispose(); TestAssert.Equal(8L, pool.Reserved);
        replacement.Dispose(); TestAssert.Equal(4L, pool.Reserved);
        TestAssert.Equal((byte)9, next.Bytes.Span[0]);
        next.Dispose(); TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Throws<ObjectDisposedException>(() => _ = next.Bytes);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void FontLeasesOversizedProgramsReleaseOnlyAfterFinalReader()
    {
        var pool = new TestMemoryBudget();
        using var cache = new FontProgramCache(2, StringComparer.Ordinal, pool);
        var program = new FontProgramBuffer(new byte[] { 1, 2, 3 }, OoxMemory.Reserve(pool, 3));
        cache.Store("large", program, null);
        using var a = program.Acquire(); using var b = program.Acquire(); program.Dispose();
        TestAssert.Equal(0L, cache.RetainedBytes); TestAssert.Equal(3L, pool.Reserved);
        a.Dispose(); cache.Dispose(); TestAssert.Equal(3L, pool.Reserved);
        b.Dispose(); TestAssert.Equal(0L, pool.Reserved);
    }

    public static void FontLeasesGovernedFilesDoNotRetainOperationBudgets()
    {
        string path = Path.GetTempFileName(); File.WriteAllBytes(path, [1, 2, 3, 4]);
        try
        {
            using var source = new FileFontProgramSource(path);
            var pool = new TestMemoryBudget();
            for (int index = 0; index < 2; index++)
            {
                using (var operation = pool.CreateScope())
                {
                    using var lease = source.AcquireAsync(operation).GetAwaiter().GetResult();
                    TestAssert.Equal(4L, pool.Reserved);
                    TestAssert.Equal((byte)4, lease.Bytes.Span[3]);
                    lease.Dispose(); TestAssert.Equal(0L, pool.Reserved);
                }
                TestAssert.Equal(0L, pool.Reserved);
            }
            var refused = new TestMemoryBudget(3);
            var error = TestAssert.Throws<OoxPdfLimitExceededException>(() => source.AcquireAsync(refused).GetAwaiter().GetResult());
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, error.Kind);
            TestAssert.Equal(OoxPdfLimitScope.Conversion, error.Scope);
            TestAssert.Equal(0L, refused.Reserved);
            using var funded = new FileFontProgramSource(path, pool);
            using var reader = funded.AcquireAsync(null).GetAwaiter().GetResult();
            funded.Dispose(); TestAssert.Equal(4L, pool.Reserved);
            reader.Dispose(); TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(0, pool.CallerDisposals + refused.CallerDisposals);
        }
        finally { File.Delete(path); }
    }

    public static void FontLeasesCanceledWaiterDoesNotInvalidateAnotherReader()
    {
        var pool = new TestMemoryBudget(); var gate = new FontAcquisitionGate("read");
        using var handler = new LeaseFontHandler(gate); using var client = new HttpClient(handler);
        using var owner = new OoxPdfFontPackResolver.FontPackFileSource("leases", new Uri("https://example.test/"), client, 4, pool);
        var source = owner.Create(PackFile("font.ttf", handler.Bytes));
        Task<IOoxPdfFontProgramLease> first = source.AcquireAsync(null).AsTask();
        try
        {
            gate.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
            using var cancellation = new CancellationTokenSource();
            Task<IOoxPdfFontProgramLease> waiter = source.AcquireAsync(null, cancellation.Token).AsTask();
            cancellation.Cancel();
            TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => waiter.GetAwaiter().GetResult()).CancellationToken);
            TestAssert.True(!first.IsCompleted && pool.Reserved >= 81920, "A canceled waiter cannot release another acquisition's read buffer.");
            gate.Release.TrySetResult();
            using var lease = first.GetAwaiter().GetResult();
            owner.Dispose(); TestAssert.Equal(4L, pool.Reserved);
            TestAssert.True(lease.Bytes.Span.SequenceEqual(handler.Bytes), "Shutdown preserves already acquired bytes.");
            TestAssert.Throws<ObjectDisposedException>(() => source.AcquireAsync(null).GetAwaiter().GetResult());
            lease.Dispose(); TestAssert.Equal(0L, pool.Reserved);
            using var stillCallerOwned = client.GetAsync("https://example.test/still-owned").GetAwaiter().GetResult();
            TestAssert.True(stillCallerOwned.IsSuccessStatusCode, "The resolver never disposes HttpClient.");
        }
        finally { gate.Release.TrySetResult(); if (first.IsCompletedSuccessfully) first.Result.Dispose(); }
    }

    public static void FontLeasesCancellationAwaitsActualReadAndOwnedDisposal()
    {
        foreach (string phase in new[] { "read", "dispose" })
        {
            var pool = new TestMemoryBudget(); var gate = new FontAcquisitionGate(phase);
            using var handler = new LeaseFontHandler(gate); using var client = new HttpClient(handler);
            using var owner = new OoxPdfFontPackResolver.FontPackFileSource("leases", new Uri("https://example.test/"), client, 4, pool);
            var source = owner.Create(PackFile("font.ttf", handler.Bytes));
            using var cancellation = new CancellationTokenSource();
            Task<IOoxPdfFontProgramLease> acquisition = source.AcquireAsync(null, cancellation.Token).AsTask();
            try
            {
                gate.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                cancellation.Cancel();
                TestAssert.True(!acquisition.IsCompleted && pool.Reserved > 0, "Cancellation must finish the pending I/O/disposal borrow before refund.");
                gate.Release.TrySetResult();
                TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => acquisition.GetAwaiter().GetResult()).CancellationToken);
                TestAssert.Equal(0L, pool.Reserved); TestAssert.Equal(0L, owner.RetainedBytes);
                TestAssert.Equal(1, handler.Requests); TestAssert.Equal(1, handler.LastStream!.AsyncDisposals);
            }
            finally { gate.Release.TrySetResult(); try { acquisition.GetAwaiter().GetResult().Dispose(); } catch (OperationCanceledException) { } }
        }
    }

    public static void FontLeasesShutdownDuringAcquisitionRefundsAfterItsRead()
    {
        var pool = new TestMemoryBudget(); var gate = new FontAcquisitionGate("read");
        using var handler = new LeaseFontHandler(gate); using var client = new HttpClient(handler);
        using var owner = new OoxPdfFontPackResolver.FontPackFileSource("leases", new Uri("https://example.test/"), client, 4, pool);
        var source = owner.Create(PackFile("font.ttf", handler.Bytes));
        Task<IOoxPdfFontProgramLease> acquisition = source.AcquireAsync(null).AsTask();
        try
        {
            gate.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
            owner.Dispose();
            TestAssert.True(!acquisition.IsCompleted && pool.Reserved >= 81920, "Shutdown cannot abandon an in-flight read.");
            gate.Release.TrySetResult();
            TestAssert.Throws<ObjectDisposedException>(() => acquisition.GetAwaiter().GetResult());
            TestAssert.Equal(0L, pool.Reserved); TestAssert.Equal(1, handler.Requests);
        }
        finally { gate.Release.TrySetResult(); try { acquisition.GetAwaiter().GetResult().Dispose(); } catch (ObjectDisposedException) { } }
    }

    public static void FontLeasesAdmissionRefusalNeverRetriesOrBecomesFontFallback()
    {
        foreach (long denied in new[] { 81924L, 81928L })
        {
            var pool = new TestMemoryBudget { RefuseChange = (_, bytes) => bytes == denied };
            using var handler = new LeaseFontHandler(null); using var client = new HttpClient(handler);
            using var owner = new OoxPdfFontPackResolver.FontPackFileSource("leases", new Uri("https://example.test/"), client, 4, pool);
            var source = owner.Create(PackFile("font.ttf", handler.Bytes));
            var face = new FontFaceResolution("Test", "Test", new FontStyleKey(), source, false);
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => FontProgramLoader.LoadAsync(face, default).GetAwaiter().GetResult());
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(OoxPdfLimitScope.SharedResource, failure.Scope);
            TestAssert.Equal(1, handler.Requests); TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(0L, owner.RetainedBytes); TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void FontLeasesParserOwnsItsCopyBeforeSourceBorrowEnds()
    {
        var pool = new TestMemoryBudget(); var source = new RecyclingFontSource();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            var face = new FontFaceResolution("Test", "Test", new FontStyleKey(), source, false);
            var font = TestAssert.NotNull(FontProgramLoader.LoadAsync(face, default).GetAwaiter().GetResult());
            TestAssert.Equal(1, source.Releases);
            TestAssert.True(source.LastBytes!.All(value => value == 0), "The source may recycle its bytes only after the lease ends.");
            TestAssert.Equal("TestFont", font.FamilyName); TestAssert.True(font.MapCodePoint(65) != 0, "The parsed font keeps its own valid program.");
            TestAssert.Equal((long)font.Bytes.Length, pool.Reserved);
            TestAssert.True(pool.Peak >= 2L * font.Bytes.Length, "Admit acquisition/copy overlap before parsing.");
        }
        TestAssert.Equal(0L, pool.Reserved); TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void FontLeasesCallerFundedMemoryIsBorrowedWithoutDuplicateAdmission()
    {
        byte[] bytes = [1, 2, 3]; var pool = new TestMemoryBudget(0);
        var source = new MemoryFontProgramSource("caller", bytes);
        using var lease = source.AcquireAsync(pool).GetAwaiter().GetResult();
        TestAssert.True(lease.Bytes.Span.SequenceEqual(bytes), "Caller-owned memory needs no new allocation.");
        TestAssert.Equal(0, pool.Requests); lease.Dispose(); lease.Dispose();
        TestAssert.True(bytes.SequenceEqual(new byte[] { 1, 2, 3 }), "Disposal ends the borrow without releasing caller storage.");
    }

    private sealed class RecyclingFontSource : IFontProgramSource
    {
        private readonly byte[] template = TestFontBuilder.CreateTestFont();
        internal byte[]? LastBytes; internal int Releases;
        public string StableId => "recycling-font";
        public ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            var reservation = OoxMemory.Reserve(operationBudget, template.Length);
            LastBytes = template.ToArray();
            return ValueTask.FromResult<IOoxPdfFontProgramLease>(new RecyclingLease(this, LastBytes, reservation));
        }
        private sealed class RecyclingLease(RecyclingFontSource source, byte[] bytes, IMemoryReservation? reservation) : IOoxPdfFontProgramLease
        {
            private bool disposed;
            public ReadOnlyMemory<byte> Bytes => !disposed ? bytes : throw new ObjectDisposedException(nameof(RecyclingLease));
            public void Dispose() { if (disposed) return; disposed = true; Array.Clear(bytes); reservation?.Dispose(); source.Releases++; }
        }
    }

    private sealed class FontAcquisitionGate(string phase)
    {
        internal readonly TaskCompletionSource Entered = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal readonly TaskCompletionSource Release = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal async ValueTask Pause(string operation, Memory<byte> bytes)
        {
            if (operation != phase || Entered.Task.IsCompleted) return;
            byte[] original = bytes.ToArray(); Entered.TrySetResult();
            await Release.Task.ConfigureAwait(false);
            TestAssert.True(bytes.Span.SequenceEqual(original), "Outstanding read buffers must not be reused.");
        }
    }

    private sealed class LeaseFontHandler(FontAcquisitionGate? gate) : HttpMessageHandler
    {
        internal readonly byte[] Bytes = [1, 2, 3, 4];
        internal int Requests; internal LeaseFontStream? LastStream;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests++;
            LastStream = new LeaseFontStream(Bytes, gate);
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StreamContent(LastStream) });
        }
    }

    private sealed class LeaseFontStream(byte[] bytes, FontAcquisitionGate? gate) : MemoryStream(bytes, writable: false)
    {
        private bool closed;
        internal int AsyncDisposals;
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Synchronous font read.");
        public override int Read(Span<byte> buffer) => throw new InvalidOperationException("Synchronous font read.");
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            if (gate is not null) await gate.Pause("read", buffer).ConfigureAwait(false);
            // MemoryStream's span override delegates to Stream for derived types,
            // which would call our deliberately rejected synchronous override.
            byte[] copy = new byte[Math.Min(2, buffer.Length)];
            int read = base.Read(copy, 0, copy.Length);
            copy.AsSpan(0, read).CopyTo(buffer.Span);
            return read; // Deliberately ignores cancellation while borrowing.
        }
        public override async ValueTask DisposeAsync()
        {
            if (closed) return;
            AsyncDisposals++;
            if (gate is not null) await gate.Pause("dispose", Memory<byte>.Empty).ConfigureAwait(false);
            closed = true; base.Dispose(true);
        }
        protected override void Dispose(bool disposing) { closed = true; base.Dispose(disposing); }
    }
}
