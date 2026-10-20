using System.Collections;
using System.Net;
using System.Reflection;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class FontResolverConcurrencyTests
{
    public static void WindowsResolverConcurrentRequestsSurviveCapacityEviction()
    {
        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-resolver-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            File.WriteAllBytes(Path.Combine(directory, "review.ttf"), TestFontBuilder.CreateCffKindFont("Review"));
            Exercise(new WindowsFontResolver(directory));
        }
        finally
        {
            WindowsFontResolver.InvalidateDiscoveryCaches();
            Directory.Delete(directory, recursive: true);
        }
    }

    public static void FontPackResolverConcurrentRequestsSurviveCapacityEviction()
    {
        using var client = new HttpClient(new ManifestHandler());
        var resolver = OoxPdfFontPackResolver.CreateHttpAsync(
            "review", new Uri("https://example.test/"), client).GetAwaiter().GetResult();
        Exercise(resolver);
    }

    private static void Exercise(IFontResolver resolver)
    {
        const int workers = 8;
        using var start = new Barrier(workers);
        Task[] tasks = Enumerable.Range(0, workers).Select(worker => Task.Factory.StartNew(() =>
        {
            TestAssert.True(start.SignalAndWait(TimeSpan.FromSeconds(10)), "All resolver workers must start.");
            string? source = null;
            // More than the full capacity per worker exercises repeated clear-all
            // eviction, while the second spelling also exercises concurrent hits.
            for (int i = 0; i < WindowsFontResolver.MaxCachedResolutions + 100; i++)
            {
                string family = $"Missing-{worker}-{i}";
                var request = new FontRequest(family, i % 2 == 0, i % 3 == 0);
                FontFaceResolution first = resolver.Resolve(request);
                FontFaceResolution second = resolver.Resolve(request with { FamilyName = family.ToLowerInvariant() });
                source ??= first.Source.StableId;
                TestAssert.True(first.IsFallback && second.IsFallback, "Unknown families must remain fallbacks.");
                TestAssert.Equal(family, first.RequestedFamily);
                TestAssert.Equal(family.ToLowerInvariant(), second.RequestedFamily);
                TestAssert.Equal("Review", first.ResolvedFamily);
                TestAssert.Equal(source, first.Source.StableId);
                TestAssert.Equal(source, second.Source.StableId);
            }
        }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default)).ToArray();
        TestAssert.True(Task.WaitAll(tasks, TimeSpan.FromSeconds(30)), "Shared resolver calls must complete without deadlock.");

        // Inspect retention after all workers finish; no production counter or
        // test-only accessor is needed for this bounded-cache contract.
        var cache = (IDictionary)resolver.GetType().GetField("requestCache", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(resolver)!;
        TestAssert.True(cache.Count is > 0 and <= WindowsFontResolver.MaxCachedResolutions,
            "Concurrent publication must retain the existing request-cache bound.");
    }

    private sealed class ManifestHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            TestAssert.True(request.RequestUri!.AbsolutePath.EndsWith("/manifest.json", StringComparison.Ordinal),
                "Resolving a face must not download its program.");
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent("""
                    {"packId":"review","files":[{"relativePath":"font.ttf","byteSize":4,"sha256":"0000000000000000000000000000000000000000000000000000000000000000"}],
                     "families":[{"requestedFamily":"Review","resolvedFamily":"Review","relativeFontFile":"font.ttf","weight":400,"italic":false,"faceIndex":0,"hasMathTable":false}],
                     "fallbacks":[{"family":"Review"}]}
                    """)
            });
        }
    }
}
