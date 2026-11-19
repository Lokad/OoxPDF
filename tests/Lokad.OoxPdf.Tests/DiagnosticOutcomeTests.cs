using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;
using System.IO.Compression;
using System.Runtime.CompilerServices;

namespace Lokad.OoxPdf.Tests;

internal static class DiagnosticOutcomeTests
{
    public static void DiagnosticCollectorBoundsSamplesWithTotals()
    {
        const int capacity = DiagnosticCollector.MaxRetainedDiagnostics;
        foreach (var (retainedSeverity, droppedSeverity) in new[]
        {
            (OoxPdfSeverity.Info, OoxPdfSeverity.Info),
            (OoxPdfSeverity.Info, OoxPdfSeverity.Warning),
            (OoxPdfSeverity.Info, OoxPdfSeverity.Error),
            (OoxPdfSeverity.Warning, OoxPdfSeverity.Info),
        })
        {
            var collector = new DiagnosticCollector();
            var retained = new OoxPdfDiagnostic("RETAINED", retainedSeverity, "retained sample");
            for (int i = 0; i < capacity; i++) collector.Add(retained);

            TestAssert.Equal(capacity, collector.DiagnosticsWithOverflow().Count);
            TestAssert.Equal(0, collector.DroppedCount);
            TestAssert.Equal(retainedSeverity != OoxPdfSeverity.Info, collector.HasWarningsOrErrors);

            collector.Add(new OoxPdfDiagnostic("DROPPED", droppedSeverity, "first discarded sample"));
            TestAssert.Equal(1, collector.DroppedCount);
            for (int i = 0; i < 99; i++) collector.Add(retained);

            bool observedSeverity = retainedSeverity != OoxPdfSeverity.Info || droppedSeverity != OoxPdfSeverity.Info;
            TestAssert.Equal(observedSeverity, collector.HasWarningsOrErrors);
            TestAssert.Equal(capacity, collector.Diagnostics.Count);
            TestAssert.True(collector.Diagnostics.All(item => item == retained), "Discarded entries must not replace retained samples.");
            TestAssert.Equal(100, collector.DroppedCount);
            TestAssert.Equal(capacity + 99, collector.OccurrenceCounts["RETAINED"]);
            TestAssert.Equal(1, collector.OccurrenceCounts["DROPPED"]);
            var serialized = collector.DiagnosticsWithOverflow();
            TestAssert.Equal(capacity + 1, serialized.Count);
            TestAssert.Equal("DIAGNOSTIC_OVERFLOW", serialized[^1].Id);
            TestAssert.Equal(observedSeverity ? OoxPdfSeverity.Error : OoxPdfSeverity.Info, serialized[^1].Severity);
        }
    }

    public static void ThrowingSummaryCallbackLeavesFileDestinationUntouched()
    {
        string input = FindCase("docx-tables.docx");
        string destination = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        File.WriteAllText(destination, "original");
        TestAssert.Throws<InvalidOperationException>(() => OoxPdfConverter.Convert(input, destination, new OoxPdfOptions
        {
            ReportResourceUsage = true,
            DiagnosticSink = ThrowOnSummary,
        }));
        TestAssert.Equal("original", File.ReadAllText(destination));
    }

    public static void ThrowingSummaryCallbackLeavesStreamOutputEmpty()
    {
        string input = FindCase("docx-tables.docx");
        using FileStream inputStream = File.OpenRead(input);
        using var output = new MemoryStream();
        TestAssert.Throws<InvalidOperationException>(() => OoxPdfConverter.Convert(inputStream, output, new OoxPdfOptions
        {
            InputKind = OoxPdfInputKind.Docx,
            ReportResourceUsage = true,
            DiagnosticSink = ThrowOnSummary,
        }));
        TestAssert.Equal(0, output.Length);
    }

    public static void FilePublicationFailureCleansStaging()
    {
        // R20: when the atomic move itself fails (destination is an existing
        // directory), the conversion propagates the failure, removes only owned
        // staging, and leaves the destination directory intact.
        string input = FindCase("docx-tables.docx");
        string directory = Path.Combine(Path.GetTempPath(), "oOxPdfPubFail" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            try
            {
                OoxPdfConverter.Convert(input, directory, new OoxPdfOptions
                {
                    InputKind = OoxPdfInputKind.Docx,
                });
                throw new InvalidOperationException("Expected publication failure.");
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Expected publication failure modes (platform move semantics differ);
                // any other exception type propagates and fails the test.
            }
            TestAssert.True(Directory.Exists(directory), "Publication failure must leave the destination directory intact.");
            string parent = TestAssert.NotNull(Path.GetDirectoryName(directory), "Destination must have a parent directory.");
            TestAssert.Equal(0, Directory.GetFiles(parent, Path.GetFileName(directory) + ".tmp-*").Length);
        }
        finally
        {
            Directory.Delete(directory);
        }
    }

    private static void ThrowOnSummary(OoxPdfDiagnostic diagnostic)
    {
        if (diagnostic.Id == "CONVERSION_RESOURCE_SUMMARY")
        {
            throw new InvalidOperationException("R20 probe callback.");
        }
    }

    public static void RecoverableDiagnosticExceptionsPreserveFileAndOriginalFailure()
    {
        WithMissingPicture((input, directory) =>
        {
            foreach (bool asynchronous in new[] { false, true })
            foreach (Exception expected in ObserverFailures())
            {
                string destination = Path.Combine(directory, "result.pdf");
                File.WriteAllText(destination, "original");
                var seen = new List<string>();
                var options = new OoxPdfOptions { DiagnosticSink = diagnostic =>
                {
                    seen.Add(diagnostic.Id);
                    if (diagnostic.Id == "IMAGE_MISSING_PART") { ThrowObserver(expected); }
                }};
                Exception actual = TestAssert.Throws<Exception>(() =>
                {
                    if (asynchronous) { OoxPdfConverter.ConvertAsync(input, destination, options).GetAwaiter().GetResult(); }
                    else { OoxPdfConverter.Convert(input, destination, options); }
                });
                AssertObserverFailure(expected, actual, seen);
                TestAssert.Equal("original", File.ReadAllText(destination));
                TestAssert.Equal(0, Directory.GetFiles(directory, "result.pdf.tmp-*").Length);
            }
        });
    }

    public static void RecoverableDiagnosticExceptionsLeaveStreamEmpty()
    {
        WithMissingPicture((input, _) =>
        {
            foreach (bool asynchronous in new[] { false, true })
            foreach (Exception expected in ObserverFailures())
            {
                using FileStream source = File.OpenRead(input);
                using var output = new MemoryStream();
                var seen = new List<string>();
                var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, DiagnosticSink = diagnostic =>
                {
                    seen.Add(diagnostic.Id);
                    if (diagnostic.Id == "IMAGE_MISSING_PART") { ThrowObserver(expected); }
                }};
                Exception actual = TestAssert.Throws<Exception>(() =>
                {
                    if (asynchronous) { OoxPdfConverter.ConvertAsync(source, output, options).GetAwaiter().GetResult(); }
                    else { OoxPdfConverter.Convert(source, output, options); }
                });
                AssertObserverFailure(expected, actual, seen);
                TestAssert.Equal(0L, output.Length);
                TestAssert.True(source.CanRead && output.CanWrite, "Caller streams must remain open.");
            }
        });
    }

    public static void DiagnosticCancellationPreservesCallerToken()
    {
        WithMissingPicture((input, directory) =>
        {
            using var cancellation = new CancellationTokenSource();
            string destination = Path.Combine(directory, "result.pdf");
            File.WriteAllText(destination, "original");
            OperationCanceledException actual = TestAssert.Throws<OperationCanceledException>(() =>
                OoxPdfConverter.Convert(input, destination, new OoxPdfOptions { DiagnosticSink = diagnostic =>
                {
                    if (diagnostic.Id == "IMAGE_MISSING_PART")
                    {
                        cancellation.Cancel();
                        cancellation.Token.ThrowIfCancellationRequested();
                    }
                }}, cancellation.Token));
            TestAssert.Equal(cancellation.Token, actual.CancellationToken);
            TestAssert.Equal("original", File.ReadAllText(destination));
            TestAssert.Equal(0, Directory.GetFiles(directory, "result.pdf.tmp-*").Length);
        });
    }

    public static void DiagnosticCallbacksCanRunNestedConversionsWithoutAffectingRecovery()
    {
        WithMissingPicture((input, directory) =>
        {
            int missingPictures = 0;
            string destination = Path.Combine(directory, "result.pdf");
            var expected = new FormatException("nested observer");
            OoxPdfConverter.Convert(input, destination, new OoxPdfOptions { DiagnosticSink = diagnostic =>
            {
                TestAssert.True(diagnostic.Id != "PPTX_NODE_RENDER_FAILED", "Observer failures are not node failures.");
                if (diagnostic.Id != "IMAGE_MISSING_PART") { return; }
                missingPictures++;
                using FileStream nestedInput = File.OpenRead(input);
                using var nestedOutput = new MemoryStream();
                Exception actual = TestAssert.Throws<FormatException>(() => OoxPdfConverter.Convert(nestedInput, nestedOutput,
                    new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, DiagnosticSink = nested =>
                    {
                        if (nested.Id == "IMAGE_MISSING_PART") { ThrowObserver(expected); }
                    }}));
                TestAssert.True(ReferenceEquals(expected, actual), "Nested conversions must restore the original failure.");
                TestAssert.Equal(0L, nestedOutput.Length);
            }});
            TestAssert.True(missingPictures > 0, "The missing-picture diagnostic must be observed.");
            TestAssert.True(File.ReadAllText(destination).StartsWith("%PDF-", StringComparison.Ordinal),
                "A missing image with a successful observer still produces a PDF.");
        });
    }

    public static void FinalSummaryCancellationPreventsPublication()
    {
        WithMissingPicture((input, directory) =>
        {
            foreach (bool asynchronous in new[] { false, true })
            {
                using var cancellation = new CancellationTokenSource();
                bool summarySeen = false;
                string destination = Path.Combine(directory, "result.pdf");
                File.WriteAllText(destination, "original");
                var options = new OoxPdfOptions { ReportResourceUsage = true, DiagnosticSink = diagnostic =>
                {
                    if (diagnostic.Id == "CONVERSION_RESOURCE_SUMMARY") { summarySeen = true; cancellation.Cancel(); }
                }};
                OperationCanceledException actual = TestAssert.Throws<OperationCanceledException>(() =>
                {
                    if (asynchronous) { OoxPdfConverter.ConvertAsync(input, destination, options, cancellation.Token).GetAwaiter().GetResult(); }
                    else { OoxPdfConverter.Convert(input, destination, options, cancellation.Token); }
                });
                TestAssert.True(summarySeen, "Cancellation must originate in the final observer.");
                TestAssert.Equal(cancellation.Token, actual.CancellationToken);
                TestAssert.Equal("original", File.ReadAllText(destination));
                TestAssert.Equal(0, Directory.GetFiles(directory, "result.pdf.tmp-*").Length);
            }
            using var streamCancellation = new CancellationTokenSource();
            using FileStream source = File.OpenRead(input);
            using var output = new MemoryStream();
            TestAssert.Throws<OperationCanceledException>(() => OoxPdfConverter.Convert(source, output,
                new OoxPdfOptions { InputKind = OoxPdfInputKind.Pptx, ReportResourceUsage = true, DiagnosticSink = diagnostic =>
                {
                    if (diagnostic.Id == "CONVERSION_RESOURCE_SUMMARY") { streamCancellation.Cancel(); }
                }}, streamCancellation.Token));
            TestAssert.Equal(0L, output.Length);
        });
    }

    private static Exception[] ObserverFailures() =>
    [new InvalidDataException("observer"), new FormatException("observer"),
     new NotSupportedException("observer"), new ArgumentException("observer"), new InvalidOperationException("observer")];

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ThrowObserver(Exception failure) => throw failure;

    private static void AssertObserverFailure(Exception expected, Exception actual, List<string> seen)
    {
        TestAssert.True(ReferenceEquals(expected, actual), "Preserve caller exception identity and type.");
        TestAssert.Contains(nameof(ThrowObserver), actual.StackTrace ?? "");
        TestAssert.True(seen.Contains("IMAGE_MISSING_PART"), "The selected observer must run.");
        TestAssert.True(!seen.Contains("PPTX_NODE_RENDER_FAILED"), "Observer failures are not malformed nodes.");
    }

    private static void WithMissingPicture(Action<string, string> check)
    {
        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-observer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string input = Path.Combine(directory, "missing-picture.pptx");
            File.Copy(FindCase("pptx-images.pptx"), input);
            File.SetAttributes(input, File.GetAttributes(input) & ~FileAttributes.ReadOnly);
            using (ZipArchive package = ZipFile.Open(input, ZipArchiveMode.Update))
            {
                ZipArchiveEntry image = package.Entries.First(entry => entry.FullName.StartsWith("ppt/media/", StringComparison.Ordinal));
                image.Delete();
            }
            check(input, directory);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static string FindCase(string name)
    {
        string[] candidates = new[]
        {
            Path.Combine("tests", "Lokad.OoxPdf.Tests", "Cases", name),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Cases", name),
        };
        foreach (string candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException("Case file not found: " + name);
    }
}
