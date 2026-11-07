using System.Buffers.Binary;
using System.Reflection;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class FontDiscoveryAsyncTests
{
    private static ValueTask<WindowsFontResolver> CreateOwned(string directory, IMemoryBudget pool, CancellationToken token = default)
        => (ValueTask<WindowsFontResolver>)typeof(WindowsFontResolver).GetMethod("CreateOwnedAsync",
            BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, [new[] { directory }, pool, token])!;

    public static void FontDiscoveryOwnedCatalogFundsAliasesAndRefundsConstruction()
    {
        InDirectory(directory =>
        {
            byte[] bytes = FontFamilyAliasTests.SharedNameProgram();
            File.WriteAllBytes(Path.Combine(directory, "aliases.ttf"), bytes);
            var expected = OpenTypeFont.ReadDiscoveryHeaders(bytes, 0);
            var pool = new TestMemoryBudget(32L * 1024 * 1024);
            using (var resolver = CreateOwned(directory, pool).AsTask().GetAwaiter().GetResult())
            {
                TestAssert.True(pool.Reserved >= 2_560_000 && pool.Peak > pool.Reserved,
                    "Decoded aliases stay funded after publication;construction overlap refunds.");
                foreach (string alias in expected.FamilyNames)
                {
                    var face = resolver.Resolve(new(alias));
                    TestAssert.True(!face.IsFallback && face.ResolvedFamily == expected.FamilyName,
                        "Admission cannot discard aliases or change canonical/style selection.");
                }
                TestAssert.Equal(1, resolver.GetDiscoveredFonts().Count);
                TestAssert.True(!ReferenceEquals(resolver.GetDiscoveredFonts(), resolver.GetDiscoveredFonts()),
                    "Owned catalog queries return caller values rather than borrowing the retained array.");
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(0, pool.CallerDisposals);
        });
    }

    public static void FontDiscoveryOwnedCatalogRefusalCancellationAndRetryAreTerminal()
    {
        InDirectory(directory =>
        {
            File.WriteAllBytes(Path.Combine(directory, "aliases.ttf"), FontFamilyAliasTests.SharedNameProgram());
            foreach (bool cancel in new[] { false, true })
            {
                using var cancellation = new CancellationTokenSource();
                var pool = new TestMemoryBudget(cancel ? 32L * 1024 * 1024 : 256L * 1024);
                if (cancel) pool.RefuseChange = (_, after) => { if (after > 1024 * 1024) cancellation.Cancel(); return false; };
                void Prepare() => CreateOwned(directory, pool, cancellation.Token).AsTask().GetAwaiter().GetResult();
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Prepare).CancellationToken);
                else
                {
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(Prepare);
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(OoxPdfLimitScope.SharedResource, failure.Scope);
                }
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(0, pool.CallerDisposals);
            }
            var retryPool = new TestMemoryBudget(32L * 1024 * 1024);
            using (var retry = CreateOwned(directory, retryPool).AsTask().GetAwaiter().GetResult())
                TestAssert.Equal(1, retry.GetDiscoveredFonts().Count);
            TestAssert.Equal(0L, retryPool.Reserved);
        });
    }

    public static void FontDiscoveryOwnedCatalogClosesSourcesButPreservesOutstandingLeases()
    {
        InDirectory(directory =>
        {
            byte[] bytes = TestFontBuilder.CreateTestFont();
            File.WriteAllBytes(Path.Combine(directory, "face.ttf"), bytes);
            var cachePool = new TestMemoryBudget(); var operationPool = new TestMemoryBudget(0);
            var resolver = CreateOwned(directory, cachePool).AsTask().GetAwaiter().GetResult();
            var face = resolver.Resolve(new("TestFont"));
            using (var operation = OoxConversionBudget.BeginScope(null, operationPool))
            using (var lease = face.Source.AcquireAsync(operation.Budget.Memory).AsTask().GetAwaiter().GetResult())
            {
                TestAssert.Equal(0L, operationPool.Reserved);
                resolver.Dispose(); resolver.Dispose();
                TestAssert.Equal((long)bytes.Length, cachePool.Reserved);
                TestAssert.True(lease.Bytes.Span.SequenceEqual(bytes), "Shutdown cannot invalidate a borrowed immutable program.");
                TestAssert.Throws<ObjectDisposedException>(() => resolver.Resolve(new("TestFont")));
                TestAssert.Throws<ObjectDisposedException>(() => resolver.GetDiscoveredFonts());
                TestAssert.Throws<ObjectDisposedException>(() => face.Source.AcquireAsync(null).AsTask().GetAwaiter().GetResult());
            }
            TestAssert.Equal(0L, cachePool.Reserved);
            TestAssert.Equal(0, cachePool.CallerDisposals);
        });
    }

    public static void FontDiscoveryOwnedCatalogKeepsMalformedRecoveryAndEmptyCatalogs()
    {
        InDirectory(directory =>
        {
            var emptyPool = new TestMemoryBudget(0);
            using (var empty = CreateOwned(directory, emptyPool).AsTask().GetAwaiter().GetResult())
            {
                TestAssert.Equal(0, empty.GetDiscoveredFonts().Count);
                TestAssert.True(empty.Resolve(new("Absent")).IsFallback, "An empty catalog retains its ordinary missing-font fallback.");
                TestAssert.Equal(0L, emptyPool.Reserved);
            }
            byte[] malformed = FontFamilyAliasTests.SharedNameProgram();
            int name = TestFontBuilder.GetTableRange(malformed, "name").Offset;
            BinaryPrimitives.WriteUInt16BigEndian(malformed.AsSpan(name + 6 + 127 * 12 + 10, 2), ushort.MaxValue);
            File.WriteAllBytes(Path.Combine(directory, "bad.ttf"), malformed);
            var pool = new TestMemoryBudget(32L * 1024 * 1024);
            using (var resolver = CreateOwned(directory, pool).AsTask().GetAwaiter().GetResult())
                TestAssert.Equal(0, resolver.GetDiscoveredFonts().Count);
            TestAssert.Equal(0L, pool.Reserved);
            File.WriteAllBytes(Path.Combine(directory, "valid.ttc"), TestFontBuilder.CreateCollection("First", "Second"));
            using (var resolver = CreateOwned(directory, pool).AsTask().GetAwaiter().GetResult())
            {
                TestAssert.Equal(2, resolver.GetDiscoveredFonts().Count);
                var first = resolver.Resolve(new("First")); var second = resolver.Resolve(new("Second"));
                TestAssert.True(ReferenceEquals(first.Source, second.Source), "Collection faces share their live program source.");
                TestAssert.Equal(0, first.FaceIndex);
                TestAssert.Equal(1, second.FaceIndex);
            }
            TestAssert.Equal(0L, pool.Reserved);
        });
    }

    public static void FontDiscoveryOwnedCatalogRequestCapacityRefusalPreservesTheCatalog()
    {
        InDirectory(directory =>
        {
            File.WriteAllBytes(Path.Combine(directory, "face.ttf"), TestFontBuilder.CreateTestFont());
            var pool = new TestMemoryBudget();
            using (var resolver = CreateOwned(directory, pool).AsTask().GetAwaiter().GetResult())
            {
                long retained = pool.Reserved;
                pool.Capacity = retained;
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission,
                    TestAssert.Throws<OoxPdfLimitExceededException>(() => resolver.Resolve(new("TestFont"))).Kind);
                TestAssert.Equal(retained, pool.Reserved);
                pool.Capacity = 32L * 1024 * 1024;
                for (int i = 0; i < WindowsFontResolver.MaxCachedResolutions + 1; i++) resolver.Resolve(new("Alias" + i));
                TestAssert.True(!resolver.Resolve(new("TestFont")).IsFallback, "Refusal and bounded cache clearing keep the catalog usable.");
                TestAssert.True(pool.Reserved > retained, "Actual retained request capacity shares the catalog owner.");
            }
            TestAssert.Equal(0L, pool.Reserved);
        });
    }

    public static void FontDiscoveryOwnedCatalogQueuedAcquisitionFinishesAfterShutdown()
    {
        InDirectory(directory =>
        {
            File.WriteAllBytes(Path.Combine(directory, "face.ttf"), TestFontBuilder.CreateTestFont());
            var pool = new TestMemoryBudget();
            var resolver = CreateOwned(directory, pool).AsTask().GetAwaiter().GetResult();
            var face = resolver.Resolve(new("TestFont"));
            var gate = (SemaphoreSlim)typeof(FileFontProgramSource).GetField("gate",
                BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(face.Source)!;
            gate.Wait();
            Task<IOoxPdfFontProgramLease> pending;
            try
            {
                pending = face.Source.AcquireAsync(null).AsTask();
                TestAssert.True(!pending.IsCompleted, "The live source acquisition is queued.");
                resolver.Dispose();
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.True(!pending.IsCompleted, "Shutdown does not fabricate completion of a queued acquisition.");
            }
            finally { gate.Release(); }
            TestAssert.Throws<ObjectDisposedException>(() => pending.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult());
            TestAssert.Equal(0L, pool.Reserved);
        });
    }

    public static void FontDiscoveryOwnedCatalogRefusalPreservesPublicConverterDestinations()
    {
        InDirectory(directory =>
        {
            File.WriteAllBytes(Path.Combine(directory, "face.ttf"), TestFontBuilder.CreateTestFont());
            var pool = new TestMemoryBudget();
            using var resolver = CreateOwned(directory, pool).AsTask().GetAwaiter().GetResult();
            long retained = pool.Reserved;
            pool.Capacity = retained;
            foreach ((string name, OoxPdfInputKind kind) in new[] {
                ("Composition/report.docx", OoxPdfInputKind.Docx),
                ("pptx-ladder-02-plain-text.pptx", OoxPdfInputKind.Pptx) })
            foreach (bool file in new[] { false, true })
            {
                var diagnostics = new List<OoxPdfDiagnostic>();
                var options = new OoxPdfOptions { InputKind = kind, FontResolver = resolver, DiagnosticSink = diagnostics.Add };
                string input = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name));
                OoxPdfLimitExceededException failure;
                if (file)
                {
                    string output = Path.Combine(directory, "result.pdf");
                    File.WriteAllBytes(output, [7, 8, 9]);
                    failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(input, output, options));
                    TestAssert.True(File.ReadAllBytes(output).SequenceEqual(new byte[] { 7, 8, 9 }), "Refusal preserves the existing destination.");
                    TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                }
                else
                {
                    using var source = File.OpenRead(input);
                    using var output = new MemoryStream();
                    failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => OoxPdfConverter.Convert(source, output, options));
                    TestAssert.Equal(0L, output.Length);
                    TestAssert.True(source.CanRead && output.CanWrite, "Caller streams remain open.");
                }
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                TestAssert.Equal(OoxPdfLimitScope.SharedResource, failure.Scope);
                TestAssert.Equal(retained, pool.Reserved);
                TestAssert.True(diagnostics.All(d => d.Id is not "FONT_NO_USABLE_FACE" and not "PPTX_NODE_RENDER_FAILED"),
                    "Shared admission refusal must escape font and node recovery.");
            }
        });
    }

    public static void FontDiscoveryOwnedCatalogLeavesDefaultSharingIndependent()
    {
        InDirectory(directory =>
        {
            File.WriteAllBytes(Path.Combine(directory, "face.ttf"), TestFontBuilder.CreateTestFont());
            using var first = CreateAsync(directory).AsTask().GetAwaiter().GetResult();
            var shared = first.GetDiscoveredFonts();
            var pool = new TestMemoryBudget();
            using (var owned = CreateOwned(directory, pool).AsTask().GetAwaiter().GetResult())
                TestAssert.Equal(shared.Single().FamilyName, owned.GetDiscoveredFonts().Single().FamilyName);
            first.Dispose();
            using var second = CreateAsync(directory).AsTask().GetAwaiter().GetResult();
            TestAssert.True(ReferenceEquals(shared, second.GetDiscoveredFonts()), "Disposing a resolver cannot dispose a shared default snapshot.");
            TestAssert.True(!second.Resolve(new("TestFont")).IsFallback, "Shared publication remains usable.");
            TestAssert.Equal(0L, pool.Reserved);
        });
        var canceledPool = new TestMemoryBudget(0);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        TestAssert.Throws<OperationCanceledException>(() => WindowsFontResolver.CreateAsync(canceledPool, cancellation.Token).AsTask().GetAwaiter().GetResult());
        TestAssert.Equal(0, canceledPool.Requests);
        TestAssert.Throws<ArgumentNullException>(() => WindowsFontResolver.CreateAsync(null!, default).AsTask().GetAwaiter().GetResult());
    }
}
