using System.Diagnostics;
using Lokad.OoxPdf;

namespace Lokad.OoxPdf.Tests;

internal static partial class PublicApiTests
{
    public static void ConvertAsyncCapturesPathsBeforeWorkerStarts()
    {
        // Isolate process-wide directory/thread-pool changes from the remaining suite.
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
            start.ArgumentList.Add(nameof(ConvertAsyncCapturesPathsBeforeWorkerStarts));
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
        var options = new OoxPdfOptions { FontResolver = new FixedFontResolver(TestFontBuilder.CreateTestFont()) };
        string previousDirectory = Environment.CurrentDirectory;
        ThreadPool.GetMinThreads(out int minWorkers, out int minIo);
        ThreadPool.GetMaxThreads(out int maxWorkers, out int maxIo);
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        Task? blocker = null;
        try
        {
            TestAssert.True(ThreadPool.SetMinThreads(1, 1) && ThreadPool.SetMaxThreads(1, 1), "The isolated process must allow one deterministic worker.");
            blocker = Task.Run(() => { entered.Set(); release.Wait(); });
            TestAssert.True(entered.Wait(TimeSpan.FromSeconds(10)), "The worker gate must enter.");
            Environment.CurrentDirectory = invocation;
            Task conversion = OoxPdfConverter.ConvertAsync("input.docx", "result.pdf", options);
            TestAssert.True(!conversion.IsCompleted, "Conversion must remain queued behind the occupied worker.");
            Environment.CurrentDirectory = worker;
            release.Set();
            TestAssert.True(conversion.Wait(TimeSpan.FromSeconds(30)), "Queued conversion must complete.");
            conversion.GetAwaiter().GetResult();

            ValidatePdfXref(File.ReadAllBytes(destination));
            TestAssert.Equal("unrelated destination", File.ReadAllText(unrelated));
            TestAssert.True(!File.Exists(Path.Combine(worker, "input.docx")), "Input must be read from the invocation directory.");
            TestAssert.Equal(0, Directory.GetFiles(invocation, "*.tmp-*").Length);
            TestAssert.Equal(0, Directory.GetFiles(worker, "*.tmp-*").Length);
        }
        finally
        {
            release.Set();
            blocker?.GetAwaiter().GetResult();
            Environment.CurrentDirectory = previousDirectory;
            ThreadPool.SetMaxThreads(maxWorkers, maxIo);
            ThreadPool.SetMinThreads(minWorkers, minIo);
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
        Task cancelled = OoxPdfConverter.ConvertAsync(null!, null!, invalidOptions, cancellation.Token);
        TestAssert.Throws<OperationCanceledException>(() => cancelled.GetAwaiter().GetResult());
        TestAssert.True(cancelled.IsCanceled, "Pre-cancellation must win over invalid paths/options without a synchronous throw.");
    }
}
