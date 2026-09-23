using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static class DiagnosticOutcomeTests
{
    public static void OverflowPreservesErrorSeverity()
    {
        var collector = new DiagnosticCollector();
        for (int i = 0; i < DiagnosticCollector.MaxRetainedDiagnostics; i++)
        {
            collector.Add(new OoxPdfDiagnostic("INFO", OoxPdfSeverity.Info, "probe", null, null, null, null, null));
        }

        TestAssert.True(!collector.HasWarningsOrErrors, "Info-only overflow must not report warnings or errors.");
        collector.Add(new OoxPdfDiagnostic("ERROR", OoxPdfSeverity.Error, "probe", null, null, null, null, null));
        TestAssert.Equal(1, collector.DroppedCount);
        TestAssert.Equal(1, collector.OccurrenceCounts["ERROR"]);
        TestAssert.True(collector.HasWarningsOrErrors, "A dropped error must still report warnings or errors for CLI strict mode.");
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
            string? parent = Path.GetDirectoryName(directory);
            TestAssert.True(parent is not null, "Destination must have a parent directory.");
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
