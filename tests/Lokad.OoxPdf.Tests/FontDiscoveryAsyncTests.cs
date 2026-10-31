using System.Reflection;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class FontDiscoveryAsyncTests
{
    public static void FontDiscoveryAsyncShortReadsPreserveSfntAndCollectionMetadata()
    {
        foreach (byte[] original in new[] { TestFontBuilder.CreateTestFont(), TestFontBuilder.CreateCollection("First", "Second") })
        {
            byte[] padded = new byte[original.Length + 65536];
            original.CopyTo(padded, 0);
            using var input = new DiscoveryStream(padded);
            var result = ReadAsync(input).AsTask().GetAwaiter().GetResult();
            TestAssert.True(!result.Complete && result.Bytes.Length < padded.Length, "Discovery must keep excluding trailing unused data.");
            TestAssert.Equal(padded.Length, result.FileLength);
            int count = OpenTypeFont.GetCollectionFontCount(original);
            for (int face = 0; face < count; face++)
            {
                var expected = OpenTypeFont.ReadDiscoveryHeaders(original, face);
                var actual = OpenTypeFont.IsTrueTypeCollectionHeader(original)
                    ? OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(result.Bytes, face, result.FileLength)
                    : OpenTypeFont.ReadDiscoveryHeaders(result.Bytes, face, result.FileLength);
                TestAssert.Equal(expected.FamilyName, actual.FamilyName);
                TestAssert.Equal(expected.WeightClass, actual.WeightClass);
                TestAssert.Equal(expected.HasMathTable, actual.HasMathTable);
                TestAssert.Equal(expected.ItalicAngle, actual.ItalicAngle);
                TestAssert.Equal(string.Join("|", expected.FamilyNames), string.Join("|", actual.FamilyNames));
            }
            TestAssert.True(input.ReadCalls > 1 && input.CanRead, "Short async reads must complete without taking ownership of the stream.");
        }
    }

    public static void FontDiscoveryAsyncFallbackRetainsSizeAndEofBounds()
    {
        byte[] malformed = new byte[128]; // No usable table directory: full bounded read.
        using var input = new DiscoveryStream(malformed);
        var result = ReadAsync(input).AsTask().GetAwaiter().GetResult();
        TestAssert.True(result.Complete && result.Bytes.AsSpan().SequenceEqual(malformed), "Malformed metadata must use the existing full-read fallback.");

        using var oversized = new DiscoveryStream([], reportedLength: FileFontProgramSource.MaxLocalFontBytes + 1);
        TestAssert.Throws<InvalidDataException>(() => ReadAsync(oversized).AsTask().GetAwaiter().GetResult());
        TestAssert.Equal(0, oversized.ReadCalls);

        using var truncated = new DiscoveryStream(new byte[4], reportedLength: 12);
        TestAssert.Throws<EndOfStreamException>(() => ReadAsync(truncated).AsTask().GetAwaiter().GetResult());
    }

    public static void FontDiscoveryAsyncCancellationWaitsForPendingRead()
        => CheckPendingReadAsync().GetAwaiter().GetResult();

    private static async Task CheckPendingReadAsync()
    {
        using var cancellation = new CancellationTokenSource();
        using var input = new DiscoveryStream(TestFontBuilder.CreateTestFont(), pauseFirstRead: true);
        Task pending = ReadAsync(input, cancellation.Token).AsTask();
        try
        {
            await input.ReadStarted.Task.WaitAsync(TimeSpan.FromSeconds(10));
            TestAssert.True(!pending.IsCompleted, "Discovery must await the pending read.");
            cancellation.Cancel();
            TestAssert.True(!pending.IsCompleted, "Cancellation cannot finish while the borrowed read is still pending.");
        }
        finally { input.ReleaseRead.TrySetResult(); }
        try
        {
            await pending.WaitAsync(TimeSpan.FromSeconds(10));
            throw new InvalidOperationException("Expected cancellation after the read finished.");
        }
        catch (OperationCanceledException) { }
        TestAssert.True(input.CanRead, "Discovery borrows the stream; it must not dispose the caller's stream.");
    }

    public static void FontDiscoveryAsyncConcurrentInitializationPublishesOneMetadataSnapshot()
    {
        InDirectory(directory =>
        {
            string path = Path.Combine(directory, "faces.ttc");
            File.WriteAllBytes(path, TestFontBuilder.CreateCollection("First", "Second"));
            Task<WindowsFontResolver>[] calls = Enumerable.Range(0, 12).Select(_ => CreateAsync(directory).AsTask()).ToArray();
            TestAssert.True(Task.WaitAll(calls, TimeSpan.FromSeconds(10)), "Concurrent initialization must complete.");
            var first = calls[0].Result;
            File.Delete(path);
            foreach (var call in calls)
            {
                TestAssert.True(ReferenceEquals(first.GetDiscoveredFonts(), call.Result.GetDiscoveredFonts()), "Callers must share the published metadata snapshot.");
                var a = call.Result.Resolve(new("First"));
                var b = call.Result.Resolve(new("Second"));
                TestAssert.True(!a.IsFallback && !b.IsFallback, "Resolving prepared metadata must not reopen removed files.");
                TestAssert.True(ReferenceEquals(a.Source, b.Source), "Collection faces retain shared source ownership.");
                TestAssert.True(ReferenceEquals(a.Source, first.Resolve(new("First")).Source), "Concurrent publication must retain one source per file.");
            }
        });
    }

    public static void FontDiscoveryAsyncCanceledWaiterDoesNotPoisonInitialization()
    {
        InDirectory(directory =>
        {
            File.WriteAllBytes(Path.Combine(directory, "face.ttf"), TestFontBuilder.CreateTestFont());
            object state = typeof(WindowsFontResolver).GetMethod("GetOrCreateDiscovery", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, new object[] { new[] { directory } })!;
            var gate = (SemaphoreSlim)state.GetType().GetField("gate", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(state)!;
            using var cancellation = new CancellationTokenSource();
            gate.Wait();
            Task<WindowsFontResolver> survivor;
            try
            {
                Task canceled = CreateAsync(directory, cancellation.Token).AsTask();
                survivor = CreateAsync(directory).AsTask();
                TestAssert.True(!canceled.IsCompleted && !survivor.IsCompleted, "Both initializers must be waiting on shared preparation.");
                cancellation.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => canceled.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult());
                TestAssert.True(!survivor.IsCompleted, "Canceling a waiter must not cancel another initializer.");
            }
            finally { gate.Release(); }
            WindowsFontResolver prepared = survivor.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
            TestAssert.True(!prepared.Resolve(new("TestFont")).IsFallback, "A canceled waiter must leave discovery retryable.");
            TestAssert.True(ReferenceEquals(prepared.GetDiscoveredFonts(), CreateAsync(directory).AsTask().GetAwaiter().GetResult().GetDiscoveredFonts()),
                "The successful snapshot remains shared after a canceled waiter.");
        });
    }

    public static void WindowsFontAsyncFactoryPreparesMetadataAndHonorsPreCancellation()
    {
        WindowsFontResolver prepared = WindowsFontResolver.CreateAsync().AsTask().GetAwaiter().GetResult();
        TestAssert.True(ReferenceEquals(prepared.GetDiscoveredFonts(), new WindowsFontResolver().GetDiscoveredFonts()),
            "The sync convenience constructor must use the same prepared metadata as the async factory.");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => WindowsFontResolver.CreateAsync(cancellation.Token).AsTask().GetAwaiter().GetResult());
    }

    // Inspect the live private entry points, without production test-only adapters.
    private static ValueTask<WindowsFontResolver> CreateAsync(string directory, CancellationToken token = default)
        => (ValueTask<WindowsFontResolver>)typeof(WindowsFontResolver).GetMethod("CreateAsync", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, new object[] { new[] { directory }, token })!;

    private static ValueTask<(byte[] Bytes, bool Complete, long FileLength)> ReadAsync(Stream stream, CancellationToken token = default, IMemoryReservation? discovery = null, bool wholeFile = false)
        => (ValueTask<(byte[], bool, long)>)typeof(WindowsFontResolver).GetMethod("ReadDiscoveryStreamAsync", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, [stream, "synthetic.ttf", wholeFile, discovery, token])!;

    private static void InDirectory(Action<string> action)
    {
        string directory = Path.Combine(Path.GetTempPath(), "async-font-discovery-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try { action(directory); }
        finally { WindowsFontResolver.InvalidateDiscoveryCaches(); Directory.Delete(directory, recursive: true); }
    }

    private sealed class DiscoveryStream(byte[] bytes, bool pauseFirstRead = false, long? reportedLength = null, Action? observeRead = null) : MemoryStream(bytes, writable: false)
    {
        public readonly TaskCompletionSource ReadStarted = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public readonly TaskCompletionSource ReleaseRead = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public int ReadCalls { get; private set; }
        public override long Length => reportedLength ?? base.Length;
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("Unexpected synchronous discovery read.");
        public override int Read(Span<byte> buffer) => throw new InvalidOperationException("Unexpected synchronous discovery read.");
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            observeRead?.Invoke();
            if (++ReadCalls == 1 && pauseFirstRead)
            {
                ReadStarted.TrySetResult();
                await ReleaseRead.Task.ConfigureAwait(false); // Deliberately ignore cancellation while I/O is pending.
            }
            byte[] chunk = new byte[Math.Min(buffer.Length, 7)];
            int read = base.Read(chunk, 0, chunk.Length);
            chunk.AsMemory(0, read).CopyTo(buffer);
            return read;
        }
    }
}
