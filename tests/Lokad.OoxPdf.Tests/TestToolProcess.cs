using System.Diagnostics;
using System.Reflection;
using System.Runtime.Versioning;

namespace Lokad.OoxPdf.Tests;

internal static class TestToolProcess
{
    public static string ResolveAssembly(string binding, string projectDirectory)
    {
        if (Environment.GetEnvironmentVariable(binding) is { Length: > 0 } configured)
        {
            string path = Path.GetFullPath(configured);
            return File.Exists(path) ? path : throw new InvalidOperationException(
                $"The assembly specified by {binding} is missing: {path}");
        }

        if (Environment.GetEnvironmentVariable("OOXPDF_TEST_REQUIRE_BOUND_TOOLS") == "1")
            throw new InvalidOperationException($"Frozen qualification requires {binding} to name a prebuilt tool.");

        // Local console builds live under tests/<project>/bin/<configuration>/<framework>.
        string root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        string configuration = typeof(TestToolProcess).Assembly.GetCustomAttribute<AssemblyConfigurationAttribute>()!.Configuration;
        string framework = "net" + new FrameworkName(AppContext.TargetFrameworkName!).Version.ToString(2);
        string assembly = Path.Combine(root, projectDirectory, "bin", configuration, framework, Path.GetFileName(projectDirectory) + ".dll");
        if (File.Exists(Path.Combine(root, "Lokad.OoxPdf.slnx")) && File.Exists(assembly))
            return assembly;

        throw new InvalidOperationException($"Tool assembly is missing: {assembly}. Build Lokad.OoxPdf.slnx or set {binding}.");
    }

    public static TestToolResult Run(string assembly, params string[] args)
    {
        var start = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = Path.GetTempPath(),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add(assembly);
        foreach (string arg in args)
            start.ArgumentList.Add(arg);

        using Process process = Process.Start(start) ?? throw new InvalidOperationException($"Failed to start {assembly}.");
        Task<string> output = process.StandardOutput.ReadToEndAsync();
        Task<string> error = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(60_000))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            throw new TimeoutException($"Tool process timed out: {assembly}");
        }
        return new TestToolResult(process.ExitCode, output.GetAwaiter().GetResult(), error.GetAwaiter().GetResult());
    }
}

internal sealed record TestToolResult(int ExitCode, string Output, string Error);
