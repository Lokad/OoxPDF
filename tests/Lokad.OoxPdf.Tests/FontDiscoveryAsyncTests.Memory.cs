using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static partial class FontDiscoveryAsyncTests
{
    public static void FontDiscoveryMemoryBuffersStayFundedThroughHeaderParsing()
    {
        foreach (byte[] font in new[] { TestFontBuilder.CreateTestFont(), TestFontBuilder.CreateCollection("First", "Second") })
        {
            byte[] padded = new byte[font.Length + 65536]; font.CopyTo(padded, 0);
            var pool = new TestMemoryBudget();
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            {
                using var discovery = OoxMemory.Reserve(operation.Budget.Memory, 0);
                using var input = new DiscoveryStream(padded, observeRead: () => TestAssert.True(pool.Reserved > 0, "Admit before async discovery reads."));
                var result = ReadAsync(input, discovery: discovery).AsTask().GetAwaiter().GetResult();
                TestAssert.True(discovery!.Bytes >= result.Bytes.LongLength, "Keep returned source bytes funded for the header parser.");
                int faces = OpenTypeFont.GetCollectionFontCount(font);
                for (int face = 0; face < faces; face++)
                {
                    var expected = OpenTypeFont.ReadDiscoveryHeaders(font, face);
                    var actual = OpenTypeFont.IsTrueTypeCollectionHeader(font)
                        ? OpenTypeFont.ReadCollectionFaceDiscoveryHeaders(result.Bytes, face, result.FileLength)
                        : OpenTypeFont.ReadDiscoveryHeaders(result.Bytes, face, result.FileLength);
                    TestAssert.Equal(expected.FamilyName, actual.FamilyName);
                    TestAssert.True(pool.Reserved >= result.Bytes.LongLength, "Parsing remains inside source ownership.");
                }
                TestAssert.True(!result.Complete && result.Bytes.Length < padded.Length, "Keep bounded selective reads.");
                TestAssert.True(input.CanRead, "The helper borrows its stream.");
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    public static void FontDiscoveryMemoryRefusalAndCancellationPrecedeReads()
    {
        foreach (bool cancel in new[] { false, true })
        {
            using var cancellation = new CancellationTokenSource();
            var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
                if (after == 0) return false;
                if (cancel) cancellation.Cancel();
                return !cancel;
            }};
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            {
                using var discovery = OoxMemory.Reserve(operation.Budget.Memory, 0);
                using var input = new DiscoveryStream(TestFontBuilder.CreateTestFont());
                void Read() => ReadAsync(input, cancellation.Token, discovery).AsTask().GetAwaiter().GetResult();
                if (cancel) TestAssert.Equal(cancellation.Token, TestAssert.Throws<OperationCanceledException>(Read).CancellationToken);
                else TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, TestAssert.Throws<OoxPdfLimitExceededException>(Read).Kind);
                TestAssert.Equal(0, input.ReadCalls);
                TestAssert.True(input.CanRead, "Early refusal/cancellation leaves the caller stream open.");
            }
            TestAssert.Equal(0L, pool.Reserved);
        }
    }

    public static void FontDiscoveryMemoryCollectionRefusalDoesNotRetryFullFile()
    {
        int refusals = 0;
        var pool = new TestMemoryBudget { RefuseChange = (_, after) => {
            if (after <= 128 * 1024 + 256) return false;
            refusals++; return true;
        }};
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            using var discovery = OoxMemory.Reserve(operation.Budget.Memory, 0);
            using var input = new MemoryStream(TestFontBuilder.CreateCollection("First", "Second"), writable: false);
            var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => ReadAsync(input, discovery: discovery).AsTask().GetAwaiter().GetResult());
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
            TestAssert.Equal(1, refusals);
            TestAssert.True(input.CanRead, "A refused TTC span cannot select full-file recovery or close the borrowed stream.");
        }
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void FontDiscoveryMemoryFallbackFundsLiveSpanAndReplacement()
    {
        byte[] font = TestFontBuilder.ReplaceTable(TestFontBuilder.CreateTestFont(), "JUNK", new byte[1024 * 1024]);
        byte[] padded = new byte[font.Length + 1024 * 1024]; font.CopyTo(padded, 0);
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            using var discovery = OoxMemory.Reserve(operation.Budget.Memory, 0);
            using var input = new MemoryStream(padded, writable: false);
            var span = ReadAsync(input, discovery: discovery).AsTask().GetAwaiter().GetResult();
            TestAssert.True(!span.Complete && span.Bytes.Length >= 1024 * 1024, "The source directory places needed metadata after significant unrelated data.");
            var full = ReadAsync(input, discovery: discovery, wholeFile: true).AsTask().GetAwaiter().GetResult();
            TestAssert.True(full.Complete && full.Bytes.AsSpan().SequenceEqual(padded), "Retain full-read byte/size semantics.");
            TestAssert.True(discovery!.Bytes >= (long)span.Bytes.Length + full.Bytes.Length, "A fallback replacement overlaps the previous returned span.");
            TestAssert.Equal(OpenTypeFont.ReadDiscoveryHeaders(span.Bytes, 0, span.FileLength).FamilyName,
                OpenTypeFont.ReadDiscoveryHeaders(full.Bytes, 0).FamilyName);
            TestAssert.True(input.CanRead, "Both reads borrow the same stream.");
        }
        TestAssert.Equal(0L, pool.Reserved);
    }

    public static void FontDiscoveryMemoryPendingCancellationKeepsBufferFunded()
        => CheckFundedPendingReadAsync().GetAwaiter().GetResult();

    private static async Task CheckFundedPendingReadAsync()
    {
        var pool = new TestMemoryBudget();
        using (var operation = OoxConversionBudget.BeginScope(null, pool))
        {
            using var cancellation = new CancellationTokenSource();
            using var input = new DiscoveryStream(TestFontBuilder.CreateTestFont(), pauseFirstRead: true);
            async Task ReadOwned()
            {
                using var discovery = OoxMemory.Reserve(operation.Budget.Memory, 0);
                await ReadAsync(input, cancellation.Token, discovery);
            }
            Task pending = ReadOwned();
            try
            {
                await input.ReadStarted.Task.WaitAsync(TimeSpan.FromSeconds(10));
                TestAssert.True(pool.Reserved > 0 && !pending.IsCompleted, "Pending I/O borrows an admitted buffer.");
                cancellation.Cancel();
                TestAssert.True(pool.Reserved > 0 && !pending.IsCompleted, "Cancellation must await the borrowed read before refunding.");
            }
            finally { input.ReleaseRead.TrySetResult(); }
            var failure = TestAssert.Throws<OperationCanceledException>(() => pending.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult());
            TestAssert.Equal(cancellation.Token, failure.CancellationToken);
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.True(input.CanRead, "Pending cancellation does not take stream ownership.");
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }

    public static void FontDiscoveryMemoryRefusalDoesNotPublishPartialCatalog()
    {
        InDirectory(directory => {
            File.WriteAllBytes(Path.Combine(directory, "a.ttf"), TestFontBuilder.CreateTestFont());
            File.WriteAllBytes(Path.Combine(directory, "b.ttf"), TestFontBuilder.ReplaceTable(
                TestFontBuilder.CreateFontWithFamilyNames((0x0409, "LargeFace")), "JUNK", new byte[1024 * 1024]));
            var pool = new TestMemoryBudget(256 * 1024);
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            {
                var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() => CreateAsync(directory).AsTask().GetAwaiter().GetResult());
                TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                TestAssert.Equal(0L, pool.Reserved);
                pool.Capacity = 16 * 1024 * 1024;
                var resolver = CreateAsync(directory).AsTask().GetAwaiter().GetResult();
                TestAssert.Equal(2, resolver.GetDiscoveredFonts().Count);
                TestAssert.True(!resolver.Resolve(new("LargeFace")).IsFallback, "Refusal cannot publish an incomplete catalog.");
                TestAssert.Equal(0L, pool.Reserved);
                pool.Capacity = 0;
                TestAssert.True(ReferenceEquals(resolver.GetDiscoveredFonts(), CreateAsync(directory).AsTask().GetAwaiter().GetResult().GetDiscoveredFonts()),
                    "Catalog hits do not reborrow discovery buffers or retain the invocation grant.");
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        });
    }

    public static void FontDiscoveryMemoryRepeatedFilesReuseCapacityAndKeepSnapshot()
    {
        InDirectory(directory => {
            File.WriteAllBytes(Path.Combine(directory, "a.ttf"), TestFontBuilder.CreateTestFont());
            File.WriteAllBytes(Path.Combine(directory, "b.ttc"), TestFontBuilder.CreateCollection("First", "Second"));
            var pool = new TestMemoryBudget();
            using (var operation = OoxConversionBudget.BeginScope(null, pool))
            {
                for (int attempt = 0; attempt < 3; attempt++)
                {
                    WindowsFontResolver.InvalidateDiscoveryCaches();
                    var resolver = CreateAsync(directory).AsTask().GetAwaiter().GetResult();
                    TestAssert.Equal(3, resolver.GetDiscoveredFonts().Count);
                    TestAssert.True(pool.Peak > 0, "The live factory must admit its discovery buffers.");
                    TestAssert.Equal(0L, pool.Reserved);
                    pool.Capacity = pool.Peak;
                    TestAssert.True(!resolver.Resolve(new("Second")).IsFallback, "Resolved metadata survives source-buffer refunds.");
                }
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        });
    }
}
