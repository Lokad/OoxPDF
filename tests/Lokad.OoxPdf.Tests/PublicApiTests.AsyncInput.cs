using System.Buffers.Binary;

namespace Lokad.OoxPdf.Tests;

internal static partial class PublicApiTests
{
    private static OoxPdfOptions AsyncInputOptions(OoxPdfInputKind kind, TestMemoryBudget? pool = null) => new()
    {
        InputKind = kind, MemoryBudget = pool, FixedCreationDate = DateTimeOffset.UnixEpoch,
        FontResolver = new FixedFontResolver(TestFontBuilder.CreateTestFont())
    };

    public static void ConvertAsyncInputUsesAwaitedShortReadsForDocxAndPptx()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        {
            byte[] bytes = kind == OoxPdfInputKind.Docx ? ReadMinimalDocxBytes("<w:p><w:r><w:t>async input</w:t></w:r></w:p>") : ReadMinimalPptxBytes();
            using var baseline = new MemoryStream();
            using (var source = new MemoryStream(bytes)) OoxPdfConverter.Convert(source, baseline, AsyncInputOptions(kind));
            foreach (var (seekable, lengthless, initialPosition) in new[] { (true, false, 0), (false, false, 0), (true, true, 0), (true, false, 17) })
            {
                var pool = new TestMemoryBudget();
                using var input = new AsyncInputStream(bytes, seekable, lengthless);
                input.Position = initialPosition;
                using var output = new MemoryStream();
                OoxPdfConverter.ConvertAsync(input, output, AsyncInputOptions(kind, pool)).GetAwaiter().GetResult();
                TestAssert.True(output.ToArray().AsSpan().SequenceEqual(baseline.ToArray()), "Async input protocol must preserve the matching synchronous PDF.");
                TestAssert.True(input.AsyncReads > 1 && input.SyncReads == 0, "Every external input read uses the asynchronous stream contract, including preflight and entries.");
                TestAssert.True((seekable && !lengthless) == (input.Seeks > 0), "Only a length-reporting seekable source permits ZIP preflight seeks on the caller stream.");
                TestAssert.Equal(0, input.Disposals);
                TestAssert.True(input.CanRead && output.CanWrite, "Caller streams stay open.");
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.Equal(0, pool.CallerDisposals);
                ValidatePdfXref(output.ToArray());
            }
        }
    }

    public static void ConvertAsyncInputCancellationAwaitsTheOutstandingBorrow()
    {
        foreach (bool seekable in new[] { false, true })
        {
            var pool = new TestMemoryBudget();
            using var input = new AsyncInputStream(ReadMinimalDocxBytes("<w:p/>"), seekable) { GateFirstRead = true };
            using var output = new MemoryStream();
            using var cancellation = new CancellationTokenSource();
            var callerScope = OoxConversionBudget.Current;
            Task conversion = OoxPdfConverter.ConvertAsync(input, output, AsyncInputOptions(OoxPdfInputKind.Docx, pool), cancellation.Token);
            try
            {
                TestAssert.Equal(cancellation.Token, input.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult());
                TestAssert.True(ReferenceEquals(callerScope, OoxConversionBudget.Current), "An awaiting conversion must not leak its AsyncLocal scope into its caller.");
                cancellation.Cancel();
                TestAssert.True(!conversion.IsCompleted && pool.Reserved >= 81920,
                    "Cancellation cannot complete/refund while a caller read still borrows the funded buffer.");
                TestAssert.Equal(0, input.Disposals);
                input.Release.TrySetResult();
                TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => conversion.GetAwaiter().GetResult()).CancellationToken);
                TestAssert.Equal(0L, output.Length);
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
                TestAssert.Equal(0, pool.CallerDisposals);
            }
            finally
            {
                cancellation.Cancel(); input.Release.TrySetResult();
                try { conversion.GetAwaiter().GetResult(); } catch (OperationCanceledException) { }
            }
        }
    }

    public static void ConvertAsyncInputRefusalPrecedesReadOrBufferGrowth()
    {
        foreach (long capacity in new[] { 0L, 81920L })
        {
            var pool = new TestMemoryBudget(capacity);
            using var input = new AsyncInputStream(ReadMinimalDocxBytes("<w:p/>"), seekable: false);
            using var output = new MemoryStream();
            var error = TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                OoxPdfConverter.ConvertAsync(input, output, AsyncInputOptions(OoxPdfInputKind.Docx, pool)).GetAwaiter().GetResult());
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, error.Kind);
            TestAssert.Equal(capacity == 0 ? 0 : 1, input.AsyncReads);
            TestAssert.Equal(0, input.SyncReads);
            TestAssert.Equal(0, input.Disposals);
            TestAssert.Equal(0L, output.Length);
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void ConvertAsyncInputTransportFailureEscapesWithoutRetry()
    {
        var pool = new TestMemoryBudget();
        var failure = new IOException("caller transport stopped during directory preflight");
        using var input = new AsyncInputStream(ReadMinimalDocxBytes("<w:p/>"), seekable: true)
            { ForbidSyncReads = false, FirstReadFailure = failure };
        using var output = new MemoryStream();
        var actual = TestAssert.Throws<IOException>(() =>
            OoxPdfConverter.ConvertAsync(input, output, AsyncInputOptions(OoxPdfInputKind.Docx, pool)).GetAwaiter().GetResult());
        TestAssert.True(ReferenceEquals(failure, actual), "The caller's first read failure must not become metadata fallback or a different exception.");
        TestAssert.Equal(1, input.AsyncReads);
        TestAssert.Equal(0, input.SyncReads);
        TestAssert.Equal(0L, output.Length);
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(0, input.Disposals);
    }

    public static void ConvertAsyncInputSharedRootKeepsPendingBorrowFunded()
    {
        var pool = new TestMemoryBudget(2 * 81920 - 1);
        byte[] bytes = ReadMinimalDocxBytes("<w:p/>");
        using var first = new AsyncInputStream(bytes, seekable: false) { GateFirstRead = true };
        using var second = new AsyncInputStream(bytes, seekable: false);
        using var firstOutput = new MemoryStream(); using var secondOutput = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        Task conversion = OoxPdfConverter.ConvertAsync(first, firstOutput, AsyncInputOptions(OoxPdfInputKind.Docx, pool), cancellation.Token);
        try
        {
            first.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
            cancellation.Cancel();
            var refused = TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                OoxPdfConverter.ConvertAsync(second, secondOutput, AsyncInputOptions(OoxPdfInputKind.Docx, pool)).GetAwaiter().GetResult());
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, refused.Kind);
            TestAssert.True(!conversion.IsCompleted && pool.Reserved == 81920, "A sibling refusal cannot refund the pending first read or borrow its capacity.");
            TestAssert.Equal(0, second.AsyncReads);
            first.Release.TrySetResult();
            TestAssert.Throws<OperationCanceledException>(() => conversion.GetAwaiter().GetResult());
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
            TestAssert.Equal(0L, firstOutput.Length + secondOutput.Length);
        }
        finally
        {
            cancellation.Cancel(); first.Release.TrySetResult();
            try { conversion.GetAwaiter().GetResult(); } catch (OperationCanceledException) { }
        }
    }

    public static void ConvertAsyncInputZip64PreflightKeepsBoundsAndContent()
    {
        byte[] original = ReadMinimalDocxBytes("<w:p><w:r><w:t>zip64</w:t></w:r></w:p>");
        foreach (string mode in new[] { "valid", "entries", "directory" })
        {
            using var input = new AsyncInputStream(Zip64(original, mode), seekable: true);
            using var output = new MemoryStream();
            void Convert() => OoxPdfConverter.ConvertAsync(input, output, AsyncInputOptions(OoxPdfInputKind.Docx)).GetAwaiter().GetResult();
            if (mode == "valid") { Convert(); ValidatePdfXref(output.ToArray()); }
            else
            {
                var error = TestAssert.Throws<OoxPdfLimitExceededException>(Convert);
                TestAssert.Contains(mode == "entries" ? "too many ZIP entries" : "central directory", error.Message);
                TestAssert.Equal(0L, output.Length);
            }
            TestAssert.Equal(0, input.SyncReads);
            TestAssert.True(input.AsyncReads > 3, "Short-read controls must traverse the ZIP64 locator and record through async reads.");
        }
    }

    private static byte[] Zip64(byte[] original, string mode)
    {
        int end = original.Length - 22;
        TestAssert.Equal(0x06054b50U, BinaryPrimitives.ReadUInt32LittleEndian(original.AsSpan(end)));
        var bytes = new byte[original.Length + 76];
        original.AsSpan(0, end).CopyTo(bytes);
        Span<byte> record = bytes.AsSpan(end, 56), locator = bytes.AsSpan(end + 56, 20);
        BinaryPrimitives.WriteUInt32LittleEndian(record, 0x06064b50); BinaryPrimitives.WriteUInt64LittleEndian(record[4..], 44);
        BinaryPrimitives.WriteUInt16LittleEndian(record[12..], 45); BinaryPrimitives.WriteUInt16LittleEndian(record[14..], 45);
        long entries = mode == "entries" ? 60000 : BinaryPrimitives.ReadUInt16LittleEndian(original.AsSpan(end + 10));
        BinaryPrimitives.WriteInt64LittleEndian(record[24..], entries); BinaryPrimitives.WriteInt64LittleEndian(record[32..], entries);
        BinaryPrimitives.WriteInt64LittleEndian(record[40..], mode == "directory" ? 64 * 1024 * 1024 : BinaryPrimitives.ReadUInt32LittleEndian(original.AsSpan(end + 12)));
        BinaryPrimitives.WriteInt64LittleEndian(record[48..], BinaryPrimitives.ReadUInt32LittleEndian(original.AsSpan(end + 16)));
        BinaryPrimitives.WriteUInt32LittleEndian(locator, 0x07064b50); BinaryPrimitives.WriteInt64LittleEndian(locator[8..], end);
        BinaryPrimitives.WriteUInt32LittleEndian(locator[16..], 1);
        original.AsSpan(end).CopyTo(bytes.AsSpan(end + 76));
        Span<byte> tail = bytes.AsSpan(end + 76);
        BinaryPrimitives.WriteUInt16LittleEndian(tail[8..], ushort.MaxValue); BinaryPrimitives.WriteUInt16LittleEndian(tail[10..], ushort.MaxValue);
        BinaryPrimitives.WriteUInt32LittleEndian(tail[12..], uint.MaxValue);
        return bytes;
    }

    private sealed class AsyncInputStream(byte[] bytes, bool seekable, bool lengthless = false) : Stream
    {
        private readonly MemoryStream data = new(bytes, writable: false);
        private int reads;
        public int AsyncReads, SyncReads, Seeks, Disposals;
        public bool ForbidSyncReads = true, GateFirstRead;
        public IOException? FirstReadFailure;
        public readonly TaskCompletionSource<CancellationToken> Entered = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public readonly TaskCompletionSource Release = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public override bool CanRead => data.CanRead;
        public override bool CanSeek => seekable && data.CanSeek;
        public override bool CanWrite => false;
        public override long Length => lengthless ? throw new NotSupportedException() : data.Length;
        public override long Position { get => data.Position; set => data.Position = value; }
        public override long Seek(long offset, SeekOrigin origin) { Seeks++; return seekable ? data.Seek(offset, origin) : throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => Read(buffer.AsSpan(offset, count));
        public override int Read(Span<byte> buffer)
        {
            SyncReads++;
            if (ForbidSyncReads) throw new InvalidOperationException("Synchronous input read was attempted.");
            if (++reads == 1 && FirstReadFailure is { } failure) throw failure;
            return data.Read(buffer[..Math.Min(13, buffer.Length)]);
        }
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            AsyncReads++;
            if (++reads == 1)
            {
                if (GateFirstRead)
                {
                    Entered.TrySetResult(cancellationToken);
                    // A caller may ignore cancellation while it still owns a read borrow.
                    await Release.Task.ConfigureAwait(false);
                }
                if (FirstReadFailure is { } failure) throw failure;
            }
            cancellationToken.ThrowIfCancellationRequested();
            return data.Read(buffer.Span[..Math.Min(13, buffer.Length)]);
        }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
            ReadAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();
        public override void Flush() => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) { Disposals++; data.Dispose(); } base.Dispose(disposing); }
    }
}
