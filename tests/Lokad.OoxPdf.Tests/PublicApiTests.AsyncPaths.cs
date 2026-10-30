using System.Diagnostics;
using Lokad.OoxPdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PublicApiTests
{
    public static void ConvertAsyncCapturesPathsBeforeHostCallbacks()
    {
        // Isolate process-wide directory changes from the remaining suite.
        const string childMarker = "OOXPDF_TEST_ASYNC_PATH_CHILD";
        if (Environment.GetEnvironmentVariable(childMarker) != "1")
        {
            var start = new ProcessStartInfo("dotnet")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };
            start.ArgumentList.Add(typeof(PublicApiTests).Assembly.Location);
            start.ArgumentList.Add("--test");
            start.ArgumentList.Add(nameof(ConvertAsyncCapturesPathsBeforeHostCallbacks));
            start.Environment[childMarker] = "1";
            using Process process = Process.Start(start) ?? throw new InvalidOperationException("Cannot start isolated async-path check.");
            Task<string> output = process.StandardOutput.ReadToEndAsync();
            Task<string> error = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(60_000))
            {
                process.Kill(entireProcessTree: true);
                process.WaitForExit();
                throw new TimeoutException("Isolated async-path check timed out.");
            }
            TestAssert.True(process.ExitCode == 0, output.GetAwaiter().GetResult() + error.GetAwaiter().GetResult());
            return;
        }

        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-async-path-" + Guid.NewGuid().ToString("N"));
        string invocation = Directory.CreateDirectory(Path.Combine(directory, "invocation")).FullName;
        string worker = Directory.CreateDirectory(Path.Combine(directory, "worker")).FullName;
        string fixture = WriteMinimalDocx("<w:p><w:r><w:t>invocation input</w:t></w:r></w:p>");
        File.Move(fixture, Path.Combine(invocation, "input.docx"));
        string destination = Path.Combine(invocation, "result.pdf");
        string unrelated = Path.Combine(worker, "result.pdf");
        File.WriteAllText(destination, "original destination");
        File.WriteAllText(unrelated, "unrelated destination");
        string previousDirectory = Environment.CurrentDirectory;
        var pool = new TestMemoryBudget(onCreate: () => Environment.CurrentDirectory = worker);
        var options = new OoxPdfOptions { MemoryBudget = pool, FontResolver = new FixedFontResolver(TestFontBuilder.CreateTestFont()) };
        try
        {
            Environment.CurrentDirectory = invocation;
            OoxPdfConverter.ConvertAsync("input.docx", "result.pdf", options).GetAwaiter().GetResult();
            TestAssert.Equal(worker, Environment.CurrentDirectory);
            TestAssert.Equal(0L, pool.Reserved);
            ValidatePdfXref(File.ReadAllBytes(destination));
            TestAssert.Equal("unrelated destination", File.ReadAllText(unrelated));
            TestAssert.True(!File.Exists(Path.Combine(worker, "input.docx")), "Input must be read from the invocation directory.");
            TestAssert.Equal(0, Directory.GetFiles(invocation, "*.tmp-*").Length);
            TestAssert.Equal(0, Directory.GetFiles(worker, "*.tmp-*").Length);
        }
        finally
        {
            Environment.CurrentDirectory = previousDirectory;
            Directory.Delete(directory, recursive: true);
        }
    }

    public static void ConvertAsyncKeepsValidationAndCancellationOnReturnedTask()
    {
        Task invalidPath = OoxPdfConverter.ConvertAsync(null!, "unused.pdf");
        TestAssert.Throws<ArgumentException>(() => invalidPath.GetAwaiter().GetResult());
        TestAssert.True(invalidPath.IsFaulted, "Invalid paths must fault the returned task.");

        var invalidOptions = new OoxPdfOptions { InputKind = (OoxPdfInputKind)int.MaxValue };
        Task invalidConfiguration = OoxPdfConverter.ConvertAsync("unused.docx", "unused.pdf", invalidOptions);
        TestAssert.Throws<ArgumentOutOfRangeException>(() => invalidConfiguration.GetAwaiter().GetResult());

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Task cancelled = OoxPdfConverter.ConvertAsync(inputPath: null!, outputPath: null!, invalidOptions, cancellation.Token);
        TestAssert.Throws<OperationCanceledException>(() => cancelled.GetAwaiter().GetResult());
        TestAssert.True(cancelled.IsCanceled, "Pre-cancellation must win over invalid paths/options without a synchronous throw.");
    }
}
