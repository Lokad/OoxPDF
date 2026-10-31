using System.Collections;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class FontProgramCacheTests
{
    public static void FontPackMetadataAdmissionLivesWithResolver()
    {
        var pool = new TestMemoryBudget();
        using var handler = new MetadataFontHandler(MetadataManifest(4000));
        using var client = new HttpClient(handler);
        var resolver = OoxPdfFontPackResolver.CreateHttpAsync("metadata", new Uri("https://example.test/"), client, pool).GetAwaiter().GetResult();
        try
        {
            TestAssert.True(pool.Reserved > 1_000_000, "Significant indexed metadata stays admitted after its manifest download refunds.");
            TestAssert.True(pool.Peak > pool.Reserved, "Metadata construction headroom refunds after indexing.");
            TestAssert.Equal(4000, resolver.GetDiscoveredFonts().Count);
            var face = resolver.Resolve(new FontRequest("Family0123", false, false));
            TestAssert.Equal("Family0123", face.ResolvedFamily);
            TestAssert.Equal(1, handler.ManifestRequests);
            TestAssert.Equal(0, handler.ProgramRequests);
            resolver.Dispose(); resolver.Dispose();
            TestAssert.Equal(0L, pool.Reserved);
            foreach (string field in new[] { "faces", "textFaces", "fallbackFamilies", "facesByFamily", "requestCache" })
            {
                var owned = (ICollection)typeof(OoxPdfFontPackResolver).GetField(field, BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(resolver)!;
                TestAssert.Equal(0, owned.Count);
            }
            // Resolution values already returned are caller-owned; the disposed
            // resolver keeps no catalog graph and rejects further source use.
            TestAssert.Equal("Family0123", face.ResolvedFamily);
            TestAssert.Throws<ObjectDisposedException>(() => resolver.GetDiscoveredFonts());
            TestAssert.Throws<ObjectDisposedException>(() => face.Source.AcquireAsync(null).GetAwaiter().GetResult());
            using var health = client.GetAsync("https://example.test/health").GetAwaiter().GetResult();
            TestAssert.True(health.IsSuccessStatusCode, "HttpClient remains caller-owned.");
            TestAssert.Equal(0, pool.CallerDisposals);
        }
        finally { resolver.Dispose(); }
    }

    public static void FontPackMetadataRefusalAndCancellationPrecedeParsing()
    {
        foreach (bool cancel in new[] { false, true })
        {
            using var cancellation = new CancellationTokenSource();
            bool requested = false;
            // An invalid, tiny payload separates metadata admission from both
            // bounded download buffers and the later JSON/semantic failure.
            var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                if (after < 1024 * 1024) return false;
                requested = true;
                if (cancel) cancellation.Cancel();
                return !cancel;
            }};
            using var handler = new MetadataFontHandler("{]"u8.ToArray());
            using var client = new HttpClient(handler);
            Task<OoxPdfFontPackResolver> creation = OoxPdfFontPackResolver.CreateHttpAsync("metadata", new Uri("https://example.test/"), client, pool, cancellation.Token);
            if (cancel)
                TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => creation.GetAwaiter().GetResult()).CancellationToken);
            else
            {
                var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => creation.GetAwaiter().GetResult());
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                TestAssert.Equal(OoxPdfLimitScope.SharedResource, failure.Scope);
            }
            TestAssert.True(requested, "Admission happens before deserializing even an invalid manifest.");
            TestAssert.Equal(1, handler.ManifestRequests);
            TestAssert.Equal(0, handler.ProgramRequests);
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
        foreach (byte[] invalid in new[] { "{]"u8.ToArray(), "{\"packId\":\"metadata\",\"files\":[]}"u8.ToArray() })
        {
            var pool = new TestMemoryBudget();
            using var handler = new MetadataFontHandler(invalid);
            using var client = new HttpClient(handler);
            TestAssert.Throws<OoxPdfFontPackException>(() => OoxPdfFontPackResolver.CreateHttpAsync("metadata", new Uri("https://example.test/"), client, pool).GetAwaiter().GetResult());
            TestAssert.True(pool.Peak > 1_000_000, "Both JSON and semantic failures occur inside metadata ownership.");
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(1, handler.ManifestRequests);
        }
    }

    public static void FontPackMetadataWaitsForPendingReadAndDisposal()
    {
        foreach (string phase in new[] { "read", "dispose" })
        foreach (bool cancel in new[] { false, true })
        {
            var pool = new TestMemoryBudget();
            var gate = new FontAcquisitionGate(phase);
            using var handler = new MetadataFontHandler(MetadataManifest(100), gate);
            using var client = new HttpClient(handler);
            using var cancellation = new CancellationTokenSource();
            using var resolver = OoxPdfFontPackResolver.CreateHttpAsync("metadata", new Uri("https://example.test/"), client, pool).GetAwaiter().GetResult();
            long metadata = pool.Reserved;
            TestAssert.True(metadata > 0, "The resolver owns admitted metadata before acquisition starts.");
            var source = resolver.Resolve(new FontRequest("Family0000", false, false)).Source;
            Task<IOoxPdfFontProgramLease> acquisition = source.AcquireAsync(null, cancellation.Token).AsTask();
            try
            {
                gate.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5)).GetAwaiter().GetResult();
                long borrowing = pool.Reserved;
                if (cancel) cancellation.Cancel();
                resolver.Dispose();
                TestAssert.True(!acquisition.IsCompleted, "Shutdown/cancellation cannot end an outstanding HTTP buffer borrow.");
                TestAssert.Equal(borrowing, pool.Reserved);
                TestAssert.True(pool.Reserved >= metadata, "Retain metadata through pending read and asynchronous disposal.");
                TestAssert.Throws<ObjectDisposedException>(() => source.AcquireAsync(null).GetAwaiter().GetResult());
                TestAssert.Equal(1, handler.ProgramRequests);
            }
            finally { resolver.Dispose(); gate.Release.TrySetResult(); }
            if (cancel)
                TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(() => acquisition.GetAwaiter().GetResult()).CancellationToken);
            else
                TestAssert.Throws<ObjectDisposedException>(() => acquisition.GetAwaiter().GetResult());
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void FontPackMetadataShutdownKeepsOnlyOutstandingProgramBytes()
    {
        var pool = new TestMemoryBudget();
        using var handler = new MetadataFontHandler(MetadataManifest(100));
        using var client = new HttpClient(handler);
        using var resolver = OoxPdfFontPackResolver.CreateHttpAsync("metadata", new Uri("https://example.test/"), client, pool).GetAwaiter().GetResult();
        long metadata = pool.Reserved;
        TestAssert.True(metadata > 0, "Retained metadata has its own cache-budget capacity.");
        IReadOnlyList<FontFaceResolution> catalog = resolver.GetDiscoveredFonts();
        using var lease = catalog[0].Source.AcquireAsync(null).GetAwaiter().GetResult();
        TestAssert.Equal(metadata + MetadataFontBytes.Length, pool.Reserved);
        resolver.Dispose();
        TestAssert.Equal((long)MetadataFontBytes.Length, pool.Reserved);
        TestAssert.True(lease.Bytes.Span.SequenceEqual(MetadataFontBytes), "A live font lease retains independently funded immutable bytes.");
        TestAssert.Equal("Family0000", catalog[0].ResolvedFamily);
        TestAssert.Throws<ObjectDisposedException>(() => catalog[0].Source.AcquireAsync(null).GetAwaiter().GetResult());
        lease.Dispose();
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void FontPackMetadataSharedRootRefusalPreservesAnotherResolver()
    {
        var pool = new TestMemoryBudget();
        using var handler = new MetadataFontHandler(MetadataManifest(100));
        using var client = new HttpClient(handler);
        using var first = OoxPdfFontPackResolver.CreateHttpAsync("metadata", new Uri("https://example.test/"), client, pool).GetAwaiter().GetResult();
        long held = pool.Reserved;
        TestAssert.True(held > 0, "The first resolver occupies its shared pool.");
        pool.Capacity = held + 1024 * 1024; // Enough for this small download; not another catalog.
        var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfFontPackResolver.CreateHttpAsync("metadata", new Uri("https://example.test/"), client, pool).GetAwaiter().GetResult());
        TestAssert.Equal(OoxPdfLimitScope.SharedResource, failure.Scope);
        TestAssert.Equal(held, pool.Reserved);
        TestAssert.Equal("Family0001", first.Resolve(new FontRequest("Family0001", false, false)).ResolvedFamily);
        TestAssert.Equal(100, first.GetDiscoveredFonts().Count);
        first.Dispose();
        TestAssert.Equal(0L, pool.Reserved);
        pool.Capacity = 128L * 1024 * 1024;
        using var retry = OoxPdfFontPackResolver.CreateHttpAsync("metadata", new Uri("https://example.test/"), client, pool).GetAwaiter().GetResult();
        TestAssert.Equal(100, retry.GetDiscoveredFonts().Count);
        retry.Dispose();
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(3, handler.ManifestRequests);
        TestAssert.Equal(0, handler.ProgramRequests);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    private static readonly byte[] MetadataFontBytes = [1, 2, 3, 4];
    private static byte[] MetadataManifest(int count) => JsonSerializer.SerializeToUtf8Bytes(new {
        packId = "metadata",
        files = new[] { new { relativePath = "font.ttf", byteSize = MetadataFontBytes.Length, sha256 = Convert.ToHexString(SHA256.HashData(MetadataFontBytes)) } },
        families = Enumerable.Range(0, count).Select(index => new { requestedFamily = "Family" + index.ToString("D4"), resolvedFamily = "Family" + index.ToString("D4"), relativeFontFile = "font.ttf" }),
        fallbacks = new[] { new { family = "Family0000" } }
    });

    private sealed class MetadataFontHandler(byte[] manifest, FontAcquisitionGate? gate = null) : HttpMessageHandler
    {
        internal int ManifestRequests, ProgramRequests;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            HttpContent content;
            if (request.RequestUri!.AbsolutePath.EndsWith("/manifest.json", StringComparison.Ordinal))
            { ManifestRequests++; content = new ByteArrayContent(manifest); }
            else if (request.RequestUri.AbsolutePath == "/health") content = new ByteArrayContent([]);
            else { ProgramRequests++; content = new StreamContent(new LeaseFontStream(MetadataFontBytes, gate)); }
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = content });
        }
    }
}
