using System.Text;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PublicApiTests
{
    private static OoxPdfOptions StorageOptions(TestMemoryBudget pool, IOoxPdfTemporaryStorage provider,
        OoxPdfInputKind kind = OoxPdfInputKind.Docx) => new()
    {
        InputKind = kind, MemoryBudget = pool, TemporaryStorage = provider,
        ConversionLimits = new() { MaxResidentPageContentBytesPerConversion = 0 },
        FixedCreationDate = DateTimeOffset.UnixEpoch, FontResolver = new FixedFontResolver(TestFontBuilder.CreateTestFont())
    };

    public static void ConvertAsyncOutputUsesCallerAsyncWriteProtocol()
    {
        byte[] bytes = ReadMinimalDocxBytes("<w:p><w:r><w:t>async output</w:t></w:r></w:p>");
        using var input = new MemoryStream(bytes);
        using var output = new AsyncEmissionStream();
        OoxPdfConverter.ConvertAsync(input, output, AsyncInputOptions(OoxPdfInputKind.Docx)).GetAwaiter().GetResult();
        TestAssert.True(output.AsyncWrites > 0 && output.SyncWrites == 0, "Actual PDF bytes must use the caller's asynchronous write contract.");
        TestAssert.Equal(0, output.AsyncDisposals + output.SyncDisposals);
        ValidatePdfXref(output.Bytes());
    }

    public static void ConvertAsyncStorageNeedsProviderBeyondResidentWindow()
    {
        using var input = new MemoryStream(ReadMinimalDocxBytes("<w:p><w:r><w:t>resident window</w:t></w:r></w:p>"));
        using var output = new MemoryStream();
        // Uses only the existing public options: the parent silently selects a temp file.
        var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx,
            FontResolver = new FixedFontResolver(TestFontBuilder.CreateTestFont()),
            ConversionLimits = new() { MaxResidentPageContentBytesPerConversion = 0 } };
        var error = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.ConvertAsync(input, output, options).GetAwaiter().GetResult());
        TestAssert.Equal(OoxPdfLimitKind.DocumentBound, error.Kind);
        TestAssert.Contains("without a temporary-storage provider", error.Message);
        TestAssert.Equal(0L, output.Length);
        TestAssert.True(input.CanRead && output.CanWrite, "Refusal leaves both caller streams open.");
    }

    public static void ConvertAsyncStorageAwaitsBothFormatsAndKeepsExactPdf()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        {
            byte[] bytes = kind == OoxPdfInputKind.Docx ? ReadMinimalDocxBytes("<w:p><w:r><w:t>temporary storage</w:t></w:r></w:p>") : ReadMinimalPptxBytes();
            using var baseline = new MemoryStream();
            using (var ordinary = new MemoryStream(bytes)) OoxPdfConverter.Convert(ordinary, baseline, AsyncInputOptions(kind));
            var pool = new TestMemoryBudget(); var provider = new OwnedTemporaryStorage();
            using var input = new AsyncInputStream(bytes, seekable: false);
            using var output = new AsyncEmissionStream();
            OoxPdfConverter.ConvertAsync(input, output, StorageOptions(pool, provider, kind)).GetAwaiter().GetResult();
            var storage = TestAssert.NotNull(provider.Created);
            TestAssert.True(output.Bytes().AsSpan().SequenceEqual(baseline.ToArray()), "Host-selected asynchronous backing must not change PDF bytes.");
            TestAssert.True(storage.AsyncWrites > 0 && storage.AsyncReads > 0 && storage.SyncWrites == 0 && storage.SyncReads == 0,
                "Staging append, backpatching and bounded readback must all use asynchronous I/O.");
            TestAssert.Equal(1, provider.Calls); TestAssert.True(provider.SawBudget, "The provider borrows the conversion budget.");
            TestAssert.Equal(1, storage.AsyncDisposals); TestAssert.Equal(0, storage.SyncDisposals);
            TestAssert.Equal(0, provider.Disposals); TestAssert.Equal(0, output.AsyncDisposals + output.SyncDisposals + input.Disposals);
            TestAssert.Equal(0L, pool.Reserved); TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed); TestAssert.Equal(0, pool.CallerDisposals);
            ValidatePdfXref(output.Bytes());
        }
    }

    public static void ConvertAsyncStorageCancellationAwaitsCreateWriteReadAndDisposal()
    {
        foreach (string phase in new[] { "create", "write", "read", "dispose", "output", "font" })
        {
            var pool = new TestMemoryBudget(); var gate = new EmissionGate(phase);
            var provider = new OwnedTemporaryStorage(gate);
            using var input = new MemoryStream(ReadMinimalDocxBytes("<w:p><w:r><w:t>pending storage</w:t></w:r></w:p>"));
            using var output = new AsyncEmissionStream(gate: gate, observeMemory: () => pool.Reserved);
            using var cancellation = new CancellationTokenSource();
            Task conversion = OoxPdfConverter.ConvertAsync(input, output, StorageOptions(pool, provider), cancellation.Token);
            try
            {
                var token = gate.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                if (phase != "dispose") TestAssert.Equal(cancellation.Token, token);
                cancellation.Cancel();
                TestAssert.True(!conversion.IsCompleted && pool.Reserved > 0,
                    "Cancellation must await the outstanding creation/I/O/disposal borrow before completing or refunding.");
                if (phase == "font") TestAssert.True(pool.Reserved >= output.BeforeFontCompression + gate.BorrowedBytes,
                    "The compressed font buffer remains admitted throughout the caller's pending write.");
                gate.Release.TrySetResult();
                TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => conversion.GetAwaiter().GetResult()).CancellationToken);
                TestAssert.Equal(1, TestAssert.NotNull(provider.Created).AsyncDisposals);
                TestAssert.Equal(0, provider.Disposals + output.AsyncDisposals + output.SyncDisposals);
                TestAssert.Equal(0L, pool.Reserved); TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.Equal(0, pool.CallerDisposals);
            }
            finally
            {
                cancellation.Cancel(); gate.Release.TrySetResult();
                try { conversion.GetAwaiter().GetResult(); } catch (OperationCanceledException) { }
            }
        }
    }

    public static void ConvertAsyncStorageRejectsInvalidBackingAndTerminalAdmissionRefusal()
    {
        foreach (string shape in new[] { "unreadable", "unwritable", "unseekable", "nonempty", "offset", "refuse" })
        {
            var pool = new TestMemoryBudget();
            if (shape == "refuse") pool.RefuseChange = (_, bytes) => bytes == OwnedTemporaryStorage.Capacity;
            var provider = new OwnedTemporaryStorage(shape: shape);
            using var input = new MemoryStream(ReadMinimalDocxBytes("<w:p><w:r><w:t>backing contract</w:t></w:r></w:p>"));
            using var output = new MemoryStream();
            void Convert() => OoxPdfConverter.ConvertAsync(input, output, StorageOptions(pool, provider)).GetAwaiter().GetResult();
            if (shape == "refuse") TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(Convert).Kind);
            else TestAssert.Throws<InvalidOperationException>(Convert);
            TestAssert.Equal(1, provider.Calls);
            if (provider.Created is { } created) TestAssert.Equal(1, created.AsyncDisposals);
            TestAssert.Equal(0L, output.Length); TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed); TestAssert.Equal(0, provider.Disposals + pool.CallerDisposals);
        }
    }

    public static void ConvertAsyncStorageFailureKeepsFilePublicationAtomic()
    {
        foreach (string phase in new[] { "create", "write", "read", "dispose" })
        {
            string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-storage-failure-" + Guid.NewGuid().ToString("N"))).FullName;
            string input = Path.Combine(directory, "input.docx"), output = Path.Combine(directory, "output.pdf");
            File.WriteAllBytes(input, ReadMinimalDocxBytes("<w:p><w:r><w:t>publication</w:t></w:r></w:p>")); File.WriteAllText(output, "previous output");
            var failure = new IOException("temporary storage " + phase + " failed");
            var pool = new TestMemoryBudget(); var provider = new OwnedTemporaryStorage(failurePhase: phase, failure: failure);
            try
            {
                var actual = TestAssert.Throws<IOException>(() => OoxPdfConverter.ConvertAsync(input, output, StorageOptions(pool, provider)).GetAwaiter().GetResult());
                TestAssert.True(ReferenceEquals(failure, actual), "Preserve the original temporary-storage failure without retry.");
                TestAssert.Equal("previous output", File.ReadAllText(output));
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                TestAssert.Equal(0L, pool.Reserved); TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.Equal(1, provider.Calls); TestAssert.Equal(0, provider.Disposals + pool.CallerDisposals);
            }
            finally { Directory.Delete(directory, recursive: true); }
        }
    }

    public static void ConvertAsyncStorageRejectsInvalidRecordsDuringReadback()
    {
        foreach (string shape in new[] { "negative-record", "truncated-record" })
        {
            var pool = new TestMemoryBudget(); var provider = new OwnedTemporaryStorage(shape: shape);
            using var input = new MemoryStream(ReadMinimalDocxBytes("<w:p><w:r><w:t>record bounds</w:t></w:r></w:p>"));
            using var output = new MemoryStream();
            TestAssert.Throws<InvalidDataException>(() => OoxPdfConverter.ConvertAsync(input, output, StorageOptions(pool, provider)).GetAwaiter().GetResult());
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(1, TestAssert.NotNull(provider.Created).AsyncDisposals);
            TestAssert.DoesNotContain("%%EOF", Encoding.ASCII.GetString(output.ToArray()));
        }
    }

    public static void ConvertAsyncStorageFileConvenienceAndExplicitDirectoryCleanUp()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-storage-files-" + Guid.NewGuid().ToString("N"))).FullName;
        string input = Path.Combine(directory, "input.docx"), output = Path.Combine(directory, "output.pdf");
        File.WriteAllBytes(input, ReadMinimalDocxBytes("<w:p><w:r><w:t>filesystem backing</w:t></w:r></w:p>"));
        try
        {
            var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = new FixedFontResolver(TestFontBuilder.CreateTestFont()),
                ConversionLimits = new() { MaxResidentPageContentBytesPerConversion = 0 }, FixedCreationDate = DateTimeOffset.UnixEpoch };
            OoxPdfConverter.ConvertAsync(input, output, options).GetAwaiter().GetResult();
            byte[] expected = File.ReadAllBytes(output); ValidatePdfXref(expected);
            using var source = File.OpenRead(input); using var destination = new MemoryStream();
            var pool = new TestMemoryBudget();
            OoxPdfConverter.ConvertAsync(source, destination, StorageOptions(pool, new OoxPdfFileTemporaryStorage(directory))).GetAwaiter().GetResult();
            TestAssert.True(expected.AsSpan().SequenceEqual(destination.ToArray()), "Explicit filesystem backing and file convenience produce identical PDFs.");
            TestAssert.Equal(0, Directory.GetFiles(directory, "OoxPdfPages-*.tmp").Length);
            TestAssert.Equal(0L, pool.Reserved);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    public static void PdfAsyncAsciiWriterPreservesScalarEncodingAcrossChunks()
    {
        foreach (int size in new[] { 4096, 65536 })
        {
            string text = new string('A', size - 1) + "\U0001F600" + new string('Z', size + 1) + "\uD800";
            using var output = new AsyncEmissionStream();
            var writer = new PdfObjectWriter(output, default, new byte[size]);
            writer.WriteAsciiAsync(text).GetAwaiter().GetResult();
            TestAssert.True(output.Bytes().AsSpan().SequenceEqual(Encoding.ASCII.GetBytes(text)), "Chunking preserves the established ASCII scalar fallback, including split surrogates.");
            TestAssert.Equal((long)output.Bytes().Length, writer.Position);
            TestAssert.Equal(0, output.SyncWrites);
        }
    }

    private sealed class EmissionGate(string phase)
    {
        private bool used;
        public int BorrowedBytes;
        public readonly TaskCompletionSource<CancellationToken> Entered = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public readonly TaskCompletionSource Release = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public async ValueTask PauseAsync(string operation, ReadOnlyMemory<byte> bytes, CancellationToken token)
        {
            if (used || phase != operation || (operation is "write" or "read" && bytes.Length <= 8)) return;
            used = true; BorrowedBytes = bytes.Length;
            byte[] snapshot = bytes.ToArray(); Entered.TrySetResult(token);
            await Release.Task.ConfigureAwait(false); // Deliberately ignores cancellation while borrowing the buffer.
            TestAssert.True(bytes.Span.SequenceEqual(snapshot), "A pending I/O buffer must not be reused before its borrow ends.");
        }
    }

    private sealed class OwnedTemporaryStorage(EmissionGate? gate = null, string? shape = null,
        string? failurePhase = null, IOException? failure = null) : IOoxPdfTemporaryStorage, IDisposable
    {
        public const int Capacity = 8191;
        public int Calls, Disposals;
        public bool SawBudget;
        public AsyncEmissionStream? Created;
        public async ValueTask<Stream> CreateAsync(IMemoryBudget? memoryBudget, CancellationToken cancellationToken = default)
        {
            Calls++; SawBudget = memoryBudget is not null;
            if (failurePhase == "create") throw failure!;
            Created = new AsyncEmissionStream(ownedStorage: true, budget: memoryBudget, gate: gate, shape: shape, failurePhase: failurePhase, failure: failure);
            if (gate is not null) await gate.PauseAsync("create", ReadOnlyMemory<byte>.Empty, cancellationToken).ConfigureAwait(false);
            return Created;
        }
        public void Dispose() => Disposals++;
    }

    private sealed class AsyncEmissionStream : Stream
    {
        private MemoryStream? data;
        private readonly IMemoryReservation? reservation;
        private readonly bool ownedStorage;
        private readonly EmissionGate? gate;
        private readonly string? shape, failurePhase;
        private readonly IOException? failure;
        private readonly Func<long>? observeMemory;
        private bool fontPayload;
        public int AsyncReads, AsyncWrites, SyncReads, SyncWrites, AsyncDisposals, SyncDisposals;
        public long BeforeFontCompression;
        public AsyncEmissionStream(bool ownedStorage = false, IMemoryBudget? budget = null, EmissionGate? gate = null,
            string? shape = null, string? failurePhase = null, IOException? failure = null, Func<long>? observeMemory = null)
        {
            this.ownedStorage = ownedStorage; this.gate = gate; this.shape = shape;
            this.failurePhase = failurePhase; this.failure = failure; this.observeMemory = observeMemory;
            if (ownedStorage && budget is not null && !budget.TryReserve(OwnedTemporaryStorage.Capacity, out reservation))
                throw new OoxPdfLimitExceededException("Host refused temporary backing.", OoxPdfLimitKind.MemoryAdmission, OoxPdfLimitScope.Conversion);
            try { data = new MemoryStream(OwnedTemporaryStorage.Capacity); }
            catch { reservation?.Dispose(); throw; }
            if (shape == "nonempty") { data.WriteByte(42); data.Position = 0; }
            if (shape == "offset") data.Position = 1;
        }
        public byte[] Bytes() => TestAssert.NotNull(data).ToArray();
        public override bool CanRead => ownedStorage && shape != "unreadable" && data is not null;
        public override bool CanWrite => shape != "unwritable" && data is not null;
        public override bool CanSeek => ownedStorage && shape != "unseekable" && data is not null;
        public override long Length => TestAssert.NotNull(data).Length;
        public override long Position { get => TestAssert.NotNull(data).Position; set => TestAssert.NotNull(data).Position = value; }
        public override long Seek(long offset, SeekOrigin origin) => CanSeek ? TestAssert.NotNull(data).Seek(offset, origin) : throw new NotSupportedException();
        public override void SetLength(long value) => TestAssert.NotNull(data).SetLength(value);
        public override void Flush() => throw new InvalidOperationException("Synchronous storage flush was attempted.");
        public override int Read(byte[] buffer, int offset, int count) { SyncReads++; throw new InvalidOperationException("Synchronous storage read was attempted."); }
        public override int Read(Span<byte> buffer) { SyncReads++; throw new InvalidOperationException("Synchronous storage read was attempted."); }
        public override void Write(byte[] buffer, int offset, int count) { SyncWrites++; throw new InvalidOperationException("Synchronous output/staging write was attempted."); }
        public override void Write(ReadOnlySpan<byte> buffer) { SyncWrites++; throw new InvalidOperationException("Synchronous output/staging write was attempted."); }
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            AsyncWrites++;
            if (!ownedStorage)
            {
                if (buffer.Span.IndexOf("/Type /FontDescriptor"u8) >= 0 && observeMemory is not null) BeforeFontCompression = observeMemory();
                if (fontPayload)
                {
                    fontPayload = false;
                    if (gate is not null) await gate.PauseAsync("font", buffer, cancellationToken).ConfigureAwait(false);
                }
                else if (buffer.Span.IndexOf("/Filter /FlateDecode /Length1"u8) >= 0) fontPayload = true;
            }
            if (gate is not null) await gate.PauseAsync(ownedStorage ? "write" : "output", buffer, cancellationToken).ConfigureAwait(false);
            if (failurePhase == "write" && buffer.Length > 8) throw failure!;
            var target = TestAssert.NotNull(data);
            if (ownedStorage && target.Position + buffer.Length > target.Capacity) throw new IOException("The controlled backing capacity was exceeded.");
            target.Write(buffer.Span);
        }
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) => WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            AsyncReads++;
            if (gate is not null) await gate.PauseAsync("read", buffer, cancellationToken).ConfigureAwait(false);
            if (failurePhase == "read") throw failure!;
            if (shape == "truncated-record") return 0;
            if (shape == "negative-record" && buffer.Length == 8) { buffer.Span.Fill(255); return 8; }
            return TestAssert.NotNull(data).Read(buffer.Span[..Math.Min(11, buffer.Length)]);
        }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) => ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        public override async ValueTask DisposeAsync()
        {
            if (data is null) return;
            AsyncDisposals++;
            try
            {
                if (gate is not null) await gate.PauseAsync("dispose", ReadOnlyMemory<byte>.Empty, default).ConfigureAwait(false);
                if (failurePhase == "dispose") throw failure!;
            }
            finally { data.Dispose(); data = null; reservation?.Dispose(); }
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && data is not null) { SyncDisposals++; data.Dispose(); data = null; reservation?.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
