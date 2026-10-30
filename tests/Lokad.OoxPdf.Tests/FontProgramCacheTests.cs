using System.Net;
using System.Security.Cryptography;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class FontProgramCacheTests
{
    public static void OversizedFontIsBorrowedWithoutExceedingRetention()
    {
        using var cache = new FontProgramCache(3, StringComparer.Ordinal);
        using var small = new FontProgramBuffer(new byte[] { 1, 2 }, null);
        using var large = new FontProgramBuffer(new byte[] { 3, 4, 5, 6 }, null);
        using var replacement = new FontProgramBuffer(new byte[] { 7, 8, 9 }, null);
        cache.Store("small", small, null);
        TestAssert.True(cache.TryAcquire("small", null, out var borrowed), "The small entry must be retained.");
        cache.Store("large", large, null);
        TestAssert.Equal(2L, cache.RetainedBytes);
        TestAssert.True(!cache.TryAcquire("large", null, out _), "An entry over the aggregate budget must not be retained.");
        cache.Store("small", replacement, null);
        TestAssert.Equal(3L, cache.RetainedBytes);
        TestAssert.True(borrowed!.Bytes.Span.SequenceEqual(new byte[] { 1, 2 }), "Replacement must preserve an already borrowed array.");
        borrowed.Dispose();
    }

    public static void LocalLoadsAndReadsRemainValidDuringConcurrentEviction()
    {
        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-font-cache-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            byte[][] expected = Enumerable.Range(0, 4).Select(i => Enumerable.Repeat((byte)(i + 1), 4096).ToArray()).ToArray();
            using var cache = new FontProgramCache(8192, StringComparer.OrdinalIgnoreCase);
            FileFontProgramSource[] sources = expected.Select((bytes, i) =>
            {
                string path = Path.Combine(directory, i + ".ttf");
                File.WriteAllBytes(path, bytes);
                return new FileFontProgramSource(path, cache);
            }).ToArray();
            using var start = new ManualResetEventSlim();
            Task[] readers = Enumerable.Range(0, 12).Select(worker => Task.Run(() =>
            {
                start.Wait();
                for (int read = 0; read < 100; read++)
                {
                    int index = (worker + read) % sources.Length;
                    using var borrowed = sources[index].AcquireAsync(null).GetAwaiter().GetResult();
                    using var competing = sources[(index + 1) % sources.Length].AcquireAsync(null).GetAwaiter().GetResult();
                    TestAssert.True(borrowed!.Bytes.Span.SequenceEqual(expected[index]), "Reads must survive competing loads and eviction.");
                    TestAssert.True(cache.RetainedBytes <= 8192, "Retention must remain bounded after every publication.");
                }
            })).ToArray();
            start.Set();
            TestAssert.True(Task.WaitAll(readers, TimeSpan.FromSeconds(30)), "Concurrent readers must finish without deadlock.");
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    public static void PackPublicationCoalescesWhileAnotherFontIsEvicted()
    {
        byte[] a = [1, 2, 3, 4];
        byte[] b = [5, 6, 7, 8];
        using var handler = new PausedFontHandler(a, b);
        using var client = new HttpClient(handler);
        using var owner = new OoxPdfFontPackResolver.FontPackFileSource("test", new Uri("https://example.test/"), client, 4);
        IFontProgramSource sourceA = owner.Create(PackFile("a.ttf", a));
        IFontProgramSource sourceB = owner.Create(PackFile("b.ttf", b));
        Task<IOoxPdfFontProgramLease> first = sourceA.AcquireAsync(null).AsTask();
        TestAssert.True(handler.Started.Task.Wait(TimeSpan.FromSeconds(10)), "The first download must be in flight.");
        Task<IOoxPdfFontProgramLease>[] waiting = Enumerable.Range(0, 16).Select(_ => sourceA.AcquireAsync(null).AsTask()).ToArray();
        using var borrowedB = sourceB.AcquireAsync(null).GetAwaiter().GetResult();
        handler.Release.SetResult();
        TestAssert.True(Task.WaitAll(waiting.Append(first).ToArray(), TimeSpan.FromSeconds(10)), "Coalesced downloads must finish.");
        TestAssert.Equal(1, handler.ARequests);
        foreach (Task<IOoxPdfFontProgramLease> read in waiting.Append(first))
        {
            TestAssert.True(read.Result.Bytes.Equals(first.Result.Bytes), "Coalesced readers must borrow the same published memory.");
            TestAssert.True(read.Result.Bytes.Span.SequenceEqual(a), "All readers must receive the verified download.");
        }
        TestAssert.Equal(4L, owner.RetainedBytes);
        TestAssert.True(borrowedB.Bytes.Span.SequenceEqual(b), "A's publication must not invalidate evicted B bytes.");
        using var reloadedB = sourceB.AcquireAsync(null).GetAwaiter().GetResult();
        TestAssert.Equal(2, handler.BRequests);
        TestAssert.True(first.Result.Bytes.Span.SequenceEqual(a), "Reloading B must not invalidate borrowed A bytes.");
        TestAssert.Equal(4L, owner.RetainedBytes);
        foreach (var task in waiting.Append(first)) task.Result.Dispose();
    }

    private static OoxPdfFontPackResolver.FontPackFile PackFile(string path, byte[] bytes) =>
        new(path, bytes.Length, Convert.ToHexString(SHA256.HashData(bytes)));

    private sealed class PausedFontHandler(byte[] a, byte[] b) : HttpMessageHandler
    {
        internal readonly TaskCompletionSource Started = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal readonly TaskCompletionSource Release = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal int ARequests;
        internal int BRequests;

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            bool isA = request.RequestUri!.AbsolutePath == "/a.ttf";
            if (isA)
            {
                Interlocked.Increment(ref ARequests);
                Started.TrySetResult();
                await Release.Task.WaitAsync(cancellationToken);
            }
            else { Interlocked.Increment(ref BRequests); }
            return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(isA ? a : b) };
        }
    }
}
