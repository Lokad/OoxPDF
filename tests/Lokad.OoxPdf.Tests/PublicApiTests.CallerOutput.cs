using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static partial class PublicApiTests
{
    public static void ConvertAsyncCallerOutputFundingSurvivesDeferredPublication()
    {
        foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
        {
            byte[] inputBytes = kind == OoxPdfInputKind.Docx
                ? ReadMinimalDocxBytes("<w:p><w:r><w:t>ABC</w:t></w:r></w:p>") : ReadMinimalPptxBytes("ABC");
            using var baseline = new MemoryStream();
            using (var input = new MemoryStream(inputBytes))
                OoxPdfConverter.Convert(input, baseline, AsyncInputOptions(kind));
            byte[] expected = baseline.ToArray();
            var pool = new TestMemoryBudget();
            long capacity = pool.Capacity;
            using (IMemoryScope caller = pool.CreateScope())
            {
                TestAssert.True(caller.TryReserve(expected.Length, out var reservation), "The host admits its retained buffer before allocation.");
                using IMemoryReservation funding = TestAssert.NotNull(reservation);
                byte[] buffer = new byte[expected.Length];
                using var output = new MemoryStream(buffer, writable: true);
                output.SetLength(0);
                using (var input = new MemoryStream(inputBytes))
                {
                    OoxPdfConverter.ConvertAsync(input, output, AsyncInputOptions(kind, pool)).GetAwaiter().GetResult();
                    TestAssert.True(input.CanRead && output.CanWrite, "Conversion leaves caller streams open.");
                }
                TestAssert.True(buffer.AsSpan().SequenceEqual(expected), "Caller funding preserves exact PDF bytes.");
                TestAssert.Equal((long)buffer.Length, pool.Reserved);
                var gate = new EmissionGate("output");
                using var publication = new AsyncEmissionStream(gate: gate);
                Task pending = publication.WriteAsync(buffer.AsMemory()).AsTask();
                try
                {
                    gate.Entered.Task.WaitAsync(TimeSpan.FromSeconds(10)).GetAwaiter().GetResult();
                    TestAssert.True(!pending.IsCompleted && gate.BorrowedBytes == buffer.Length, "Deferred publication still borrows the retained output.");
                    TestAssert.Equal((long)buffer.Length, pool.Reserved);
                    pool.Capacity = pool.Reserved;
                    using var nextInput = new MemoryStream(inputBytes);
                    using var nextOutput = new MemoryStream();
                    var failure = TestAssert.Throws<OoxPdfLimitExceededException>(() =>
                        OoxPdfConverter.ConvertAsync(nextInput, nextOutput, AsyncInputOptions(kind, pool)).GetAwaiter().GetResult());
                    TestAssert.Equal(OoxPdfLimitKind.MemoryAdmission, failure.Kind);
                    TestAssert.Equal(0L, nextOutput.Length);
                    TestAssert.Equal((long)buffer.Length, pool.Reserved);
                    gate.Release.TrySetResult();
                    pending.GetAwaiter().GetResult();
                    TestAssert.True(publication.Bytes().AsSpan().SequenceEqual(expected), "Publication consumes the unchanged admitted buffer.");
                    TestAssert.Equal((long)buffer.Length, pool.Reserved);
                    TestAssert.Equal(0, pool.CallerDisposals);
                }
                finally { gate.Release.TrySetResult(); pending.GetAwaiter().GetResult(); }
            }
            TestAssert.Equal(0L, pool.Reserved);
            pool.Capacity = capacity;
            using var finalInput = new MemoryStream(inputBytes);
            using var finalOutput = new MemoryStream();
            OoxPdfConverter.ConvertAsync(finalInput, finalOutput, AsyncInputOptions(kind, pool)).GetAwaiter().GetResult();
            TestAssert.True(finalOutput.ToArray().AsSpan().SequenceEqual(expected), "Host release makes shared capacity available to the next conversion.");
            TestAssert.Equal(0L, pool.Reserved);
            TestAssert.Equal(pool.ScopesCreated, pool.ScopesDisposed);
            TestAssert.Equal(0, pool.CallerDisposals);
        }
    }

    public static void ConversionResourceSummaryPreservesRouteTimingAndProvenance()
    {
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-summary-timing-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (var kind in new[] { OoxPdfInputKind.Docx, OoxPdfInputKind.Pptx })
            foreach (bool file in new[] { false, true })
            {
                byte[] bytes = kind == OoxPdfInputKind.Docx
                    ? ReadMinimalDocxBytes("<w:p><w:r><w:t>ABC</w:t></w:r></w:p>") : ReadMinimalPptxBytes("ABC");
                string inputPath = Path.Combine(directory, kind == OoxPdfInputKind.Docx ? "input.docx" : "input.pptx");
                string outputPath = Path.Combine(directory, "output.pdf");
                File.WriteAllBytes(inputPath, bytes); File.WriteAllText(outputPath, "previous");
                using var input = new MemoryStream(bytes); using var output = new MemoryStream();
                var pool = new TestMemoryBudget(); int summaries = 0;
                var options = new OoxPdfOptions { InputKind = kind, MemoryBudget = pool,
                    FixedCreationDate = DateTimeOffset.UnixEpoch, FontResolver = AsyncInputOptions(kind).FontResolver,
                    ReportResourceUsage = true, DiagnosticSink = diagnostic =>
                    {
                        if (diagnostic.Id != "CONVERSION_RESOURCE_SUMMARY") return;
                        summaries++;
                        TestAssert.Equal(OoxPdfSeverity.Info, diagnostic.Severity);
                        TestAssert.True(diagnostic.PartName is null && diagnostic.SlideIndex is null && diagnostic.PageIndex is null && diagnostic.Feature is null && diagnostic.Fallback is null,
                            "The conversion summary retains its conversion-wide provenance.");
                        long count = long.Parse(System.Text.RegularExpressions.Regex.Match(diagnostic.Message, @"pdfOutputBytes=(\d+);").Groups[1].Value,
                            System.Globalization.CultureInfo.InvariantCulture);
                        if (file)
                        {
                            TestAssert.Equal("previous", File.ReadAllText(outputPath));
                            string staged = Directory.GetFiles(directory, "output.pdf.tmp-*").Single();
                            TestAssert.Equal(new FileInfo(staged).Length, count);
                            TestAssert.True(count > 0, "File totals describe completed staging before publication.");
                            TestAssert.Equal(0L, pool.Reserved);
                        }
                        else
                        {
                            TestAssert.Equal(0L, output.Length); TestAssert.Equal(0L, count);
                            TestAssert.True(pool.Reserved > 0, "A zero serialized-output total does not imply zero retained conversion memory.");
                        }
                    }};
                if (file) OoxPdfConverter.ConvertAsync(inputPath, outputPath, options).GetAwaiter().GetResult();
                else OoxPdfConverter.ConvertAsync(input, output, options).GetAwaiter().GetResult();
                TestAssert.Equal(1, summaries);
                ValidatePdfXref(file ? File.ReadAllBytes(outputPath) : output.ToArray());
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
                TestAssert.Equal(0L, pool.Reserved);
                TestAssert.Equal(0, pool.CallerDisposals);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }
}
