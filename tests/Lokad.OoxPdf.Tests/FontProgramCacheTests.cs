using System.Net;
using System.Security.Cryptography;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class FontProgramCacheTests
{
    public static void OversizedFontIsBorrowedWithoutExceedingRetention()
    {
        var cache = new FontProgramCache(3, StringComparer.Ordinal);
        cache.Store("small", [1, 2]);
        TestAssert.True(cache.TryGet("small", out ReadOnlyMemory<byte> borrowed), "The small entry must be retained.");
        cache.Store("large", [3, 4, 5, 6]);
        TestAssert.Equal(2L, cache.RetainedBytes);
        TestAssert.True(!cache.TryGet("large", out _), "An entry over the aggregate budget must not be retained.");
        cache.Store("small", [7, 8, 9]);
        TestAssert.Equal(3L, cache.RetainedBytes);
        TestAssert.True(borrowed.Span.SequenceEqual(new byte[] { 1, 2 }), "Replacement must preserve an already borrowed array.");
    }

    public static void LocalLoadsAndReadsRemainValidDuringConcurrentEviction()
    {
        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-font-cache-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            byte[][] expected = Enumerable.Range(0, 4).Select(i => Enumerable.Repeat((byte)(i + 1), 4096).ToArray()).ToArray();
            var cache = new FontProgramCache(8192, StringComparer.OrdinalIgnoreCase);
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
                    ReadOnlyMemory<byte> borrowed = sources[index].GetBytesAsync().GetAwaiter().GetResult();
                    sources[(index + 1) % sources.Length].GetBytesAsync().GetAwaiter().GetResult();
                    TestAssert.True(borrowed.Span.SequenceEqual(expected[index]), "Reads must survive competing loads and eviction.");
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
        var owner = new OoxPdfFontPackResolver.FontPackFileSource("test", new Uri("https://example.test/"), client, 4);
        IFontProgramSource sourceA = owner.Create(PackFile("a.ttf", a));
        IFontProgramSource sourceB = owner.Create(PackFile("b.ttf", b));
        Task<ReadOnlyMemory<byte>> first = sourceA.GetBytesAsync(default).AsTask();
        TestAssert.True(handler.Started.Task.Wait(TimeSpan.FromSeconds(10)), "The first download must be in flight.");
        Task<ReadOnlyMemory<byte>>[] waiting = Enumerable.Range(0, 16).Select(_ => sourceA.GetBytesAsync(default).AsTask()).ToArray();
        ReadOnlyMemory<byte> borrowedB = sourceB.GetBytesAsync(default).GetAwaiter().GetResult();
        handler.Release.SetResult();
        TestAssert.True(Task.WaitAll(waiting.Append(first).ToArray(), TimeSpan.FromSeconds(10)), "Coalesced downloads must finish.");
        TestAssert.Equal(1, handler.ARequests);
        foreach (Task<ReadOnlyMemory<byte>> read in waiting.Append(first))
        {
            TestAssert.True(read.Result.Equals(first.Result), "Coalesced readers must borrow the same published memory.");
            TestAssert.True(read.Result.Span.SequenceEqual(a), "All readers must receive the verified download.");
        }
        TestAssert.Equal(4L, owner.RetainedBytes);
        TestAssert.True(borrowedB.Span.SequenceEqual(b), "A's publication must not invalidate evicted B bytes.");
        sourceB.GetBytesAsync(default).GetAwaiter().GetResult();
        TestAssert.Equal(2, handler.BRequests);
        TestAssert.True(first.Result.Span.SequenceEqual(a), "Reloading B must not invalidate borrowed A bytes.");
        TestAssert.Equal(4L, owner.RetainedBytes);
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
