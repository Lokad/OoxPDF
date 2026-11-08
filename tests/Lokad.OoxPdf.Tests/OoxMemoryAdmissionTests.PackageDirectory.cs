using System.Buffers.Binary;
using System.IO.Compression;
using System.Xml;
using Lokad.OoxPdf.Ooxml;

namespace Lokad.OoxPdf.Tests;

internal static partial class OoxMemoryAdmissionTests
{
    private const string NamePackageTypes = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"bin\" ContentType=\"application/octet-stream\"/></Types>";
    private const string EmptyNamePackageTypes = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>";

    public static void MemoryBudgetPackageDirectoryRefusesBeforeIndexMaterialization()
    {
        using (var warmInput = new MemoryStream(NamePackage(1, 16)))
        using (var warm = OoxPackage.Open(warmInput, default)) { }
        byte[] bytes = NamePackage(1024, 1024);
        var pool = new TestMemoryBudget(1024 * 1024);
        using var input = new MemoryStream(bytes, writable: false);
        long start = GC.GetAllocatedBytesForCurrentThread();
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            var error = TestAssert.Throws<OoxPdfLimitExceededException>(() => {
                using var package = OoxPackage.Open(input, default);
            });
            TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, error.Kind);
        }
        long allocated = GC.GetAllocatedBytesForCurrentThread() - start;
        TestAssert.True(allocated < 256 * 1024,
            "A refused directory must not first allocate its many decoded names and entry objects.");
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.True(input.CanRead, "Refusal keeps the caller's input open.");
        Console.WriteLine($"Package directory refusal: {allocated} bytes allocated after caller-owned input creation.");
    }

    public static void MemoryBudgetPackageNamesStayFundedAfterArchiveDisposal()
    {
        byte[] bytes = NamePackage(2048, 4096);
        var pool = new TestMemoryBudget();
        using var input = new MemoryStream(bytes, writable: false);
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            using var package = OoxPackage.Open(input, default);
            TestAssert.Equal(2049, package.Parts.Count);
            long names = package.Parts.Sum(part => 2L * part.Name.Length);
            long payload = package.Parts.Sum(part => part.Bytes.LongLength);
            TestAssert.True(names > 16 * 1024 * 1024, "The fixture has significant retained names and almost no payload.");
            TestAssert.True(pool.Reserved >= names + payload, "The package must keep its selected names and index funded.");
            TestAssert.True(pool.Peak > pool.Reserved + names,
                "ZIP names and selected canonical names overlap before the archive closes.");
            Console.WriteLine($"Package directory ownership: names={names}, retained={pool.Reserved}, peak={pool.Peak}.");
            TestAssert.True(package.Parts.All(part => part.Name.StartsWith('/')),
                "The producer supplies canonical names to both the index and content-type lookup.");
            TestAssert.Equal("application/octet-stream", package.Parts.First(part => part.Name.EndsWith(".bin", StringComparison.Ordinal)).ContentType);
            package.Dispose();
            package.Dispose();
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.True(input.CanRead, "Package disposal leaves its caller's stream open.");
        }
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        TestAssert.Equal(0, pool.CallerDisposals);
    }

    public static void MemoryBudgetPackageIgnoredNamesRefundDirectoryConstruction()
    {
        byte[] bytes = NamePackage(128, 1024, EmptyNamePackageTypes);
        var pool = new TestMemoryBudget();
        using var input = new MemoryStream(bytes, writable: false);
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            using var package = OoxPackage.Open(input, default);
            TestAssert.Equal(1, package.Parts.Count);
            TestAssert.True(pool.Peak > 256 * 1024, "The discarded ZIP names still require funded construction.");
            TestAssert.True(pool.Reserved < 64 * 1024,
                "The package must not retain the directory window for names it did not keep.");
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }

    public static void MemoryBudgetPackageDirectoryCancellationAndFailuresRefund()
    {
        foreach (string mode in new[] { "cancel", "later-cancel", "observer", "missing-types", "invalid-types" })
        {
            using var cancellation = new CancellationTokenSource();
            byte[] bytes = NamePackage(128, 1024, mode == "missing-types" ? null : mode == "invalid-types" ? "<Types" : NamePackageTypes);
            var pool = new TestMemoryBudget();
            bool directoryRequested = false;
            pool.RefuseChange = (before, after) => {
                if (after > 512 * 1024)
                {
                    directoryRequested = true;
                    if (mode == "cancel") cancellation.Cancel();
                    if (mode == "observer") throw new InvalidOperationException("directory observer");
                }
                else if (directoryRequested && before > 0 && mode == "later-cancel") cancellation.Cancel();
                return false;
            };
            using var input = new MemoryStream(bytes, writable: false);
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                void Open() { using var package = OoxPackage.Open(input, cancellation.Token); }
                if (mode is "cancel" or "later-cancel") TestAssert.Throws<OperationCanceledException>(Open);
                else if (mode == "observer") TestAssert.Equal("directory observer", TestAssert.Throws<InvalidOperationException>(Open).Message);
                else if (mode == "missing-types") TestAssert.Throws<InvalidDataException>(Open);
                else TestAssert.Throws<XmlException>(Open);
                TestAssert.True(directoryRequested, "The failing operation must exercise funded directory ownership.");
                TestAssert.Equal(0L, pool.Reserved);
            }
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.True(input.CanRead, "Cancellation and failures do not close the caller's stream.");
        }
    }

    public static void MemoryBudgetPackageDirectoryMetadataFallbackKeepsNamesAndOwnership()
    {
        byte[] bytes = NamePackage(8, 128);
        var pool = new TestMemoryBudget();
        using var input = new PackagePositionUnavailableInput(bytes);
        using (var conversion = OoxConversionBudget.BeginScope(null, pool))
        {
            using var package = OoxPackage.Open(input, default);
            TestAssert.True(input.PositionQueryFailed, "The test exercises the existing unavailable-metadata fallback.");
            TestAssert.Equal(9, package.Parts.Count);
            TestAssert.True(pool.Reserved >= package.Parts.Sum(part => 2L * part.Name.Length + part.Bytes.LongLength),
                "Unavailable preflight metadata does not make surviving names free.");
        }
        TestAssert.Equal(0L, pool.Reserved);
        TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
    }

    public static void MemoryBudgetPackageDirectoryInputBoundsPrecedeAdmission()
    {
        foreach (bool count in new[] { true, false })
        {
            byte[] bytes = NamePackage(1, 16);
            if (count) BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(bytes.Length - 22 + 10), 60_000);
            else BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(bytes.Length - 22 + 12), 64 * 1024 * 1024);
            var pool = new TestMemoryBudget(81920);
            using var input = new MemoryStream(bytes, writable: false);
            using (var conversion = OoxConversionBudget.BeginScope(null, pool))
            {
                var error = TestAssert.Throws<OoxPdfLimitExceededException>(() => {
                    using var package = OoxPackage.Open(input, default);
                });
                TestAssert.Equal(OoxPdfLimitKind.DocumentBound, error.Kind);
                TestAssert.Contains(count ? "too many ZIP entries" : "central directory", error.Message);
            }
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
        }
    }

    private static byte[] NamePackage(int count, int nameLength, string? contentTypes = NamePackageTypes)
    {
        using var bytes = new MemoryStream();
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Create, leaveOpen: true))
        {
            if (contentTypes is not null)
            {
                var entry = zip.CreateEntry("[Content_Types].xml", CompressionLevel.NoCompression);
                entry.LastWriteTime = new(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var output = entry.Open();
                output.Write(TestFixtures.Utf8(contentTypes));
            }
            for (int index = 0; index < count; index++)
            {
                string prefix = index.ToString("D5", System.Globalization.CultureInfo.InvariantCulture) + "-";
                var entry = zip.CreateEntry(prefix + new string('a', nameLength - prefix.Length - 4) + ".bin", CompressionLevel.NoCompression);
                entry.LastWriteTime = new(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
            }
        }
        return bytes.ToArray();
    }

    private sealed class PackagePositionUnavailableInput(byte[] bytes) : MemoryStream(bytes, writable: false)
    {
        public bool PositionQueryFailed { get; private set; }
        public override long Position
        {
            get
            {
                if (!PositionQueryFailed) { PositionQueryFailed = true; throw new NotSupportedException("metadata position"); }
                return base.Position;
            }
            set => base.Position = value;
        }
    }
}
