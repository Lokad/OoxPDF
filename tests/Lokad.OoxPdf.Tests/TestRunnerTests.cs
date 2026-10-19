using System.Text.Json;

namespace Lokad.OoxPdf.Tests;

internal static class TestRunnerTests
{
    public static void TestRunnerRejectsUnknownArgumentsBeforeEnumeratingTests()
    {
        foreach (string[] args in new string[][]
        {
            ["--filter", "unintended"], ["--list", "--filter", "unintended"],
            ["unexpected"], ["--skip-slow", "unexpected"], ["--LIST"]
        })
        {
            var result = RunCaptured(args, UnexpectedEnumeration());
            TestAssert.Equal(1, result.ExitCode);
            TestAssert.Contains("Unknown argument", result.Output);
        }
    }

    public static void TestRunnerRejectsMissingAndRepeatedOptionsBeforeEnumeratingTests()
    {
        foreach (string option in new[] { "--group", "--test", "--test-set", "--report" })
        {
            foreach (string[] args in new string[][] { [option], [option, "--list"], [option, " "] })
            {
                var result = RunCaptured(args, UnexpectedEnumeration());
                TestAssert.Equal(1, result.ExitCode);
                TestAssert.Contains("requires a value", result.Output);
            }
            var repeated = RunCaptured([option, "first", option, "second"], UnexpectedEnumeration());
            TestAssert.Equal(1, repeated.ExitCode);
            TestAssert.Contains("Repeated option", repeated.Output);
        }
        foreach (string flag in new[] { "--list", "--skip-slow", "--only-slow" })
        {
            var repeated = RunCaptured([flag, flag], UnexpectedEnumeration());
            TestAssert.Equal(1, repeated.ExitCode);
            TestAssert.Contains("Repeated option", repeated.Output);
        }
    }

    public static void TestRunnerPreservesExactSetsGroupsListingAndReports()
    {
        var calls = new List<int>();
        Action first = () => calls.Add(1);
        Action second = () => calls.Add(2);
        TestCase[] catalog = [.. TestCatalog.All, new("runner-probe", first), new("runner-probe", second)];
        var list = RunCaptured(["--group", "runner-probe", "--list"], catalog);
        TestAssert.Equal(0, list.ExitCode);
        TestAssert.Contains(first.Method.Name, list.Output);
        TestAssert.Contains(second.Method.Name, list.Output);
        TestAssert.Equal(0, calls.Count);
        TestAssert.Equal(0, RunCaptured(["--group", "runner-probe", "--skip-slow"], catalog).ExitCode);
        TestAssert.True(calls.SequenceEqual([1, 2]), "Group selection executes only its registered probes.");
        calls.Clear();
        TestAssert.Equal(0, RunCaptured(["--group", "runner-probe", "--only-slow"], catalog).ExitCode);
        TestAssert.Equal(0, calls.Count);
        TestAssert.Equal(0, RunCaptured(["--test", first.Method.Name], catalog).ExitCode);
        TestAssert.True(calls.SequenceEqual([1]), "The supported name filter keeps its selection.");
        calls.Clear();
        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-runner-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string set = Path.Combine(directory, "exact methods.json");
            string report = Path.Combine(directory, "report.json");
            File.WriteAllText(set, JsonSerializer.Serialize(new[] { second.Method.Name }));
            TestAssert.Equal(0, RunCaptured(["--test-set", set, "--report", report], catalog).ExitCode);
            TestAssert.True(calls.SequenceEqual([2]), "An exact set executes only its named method.");
            using var json = JsonDocument.Parse(File.ReadAllText(report));
            TestAssert.Equal(1, json.RootElement.GetProperty("passed").GetInt32());
            TestAssert.Equal(second.Method.Name, json.RootElement.GetProperty("tests")[0].GetProperty("name").GetString());
            calls.Clear();
            TestAssert.Equal(1, RunCaptured(["--test-set", set, "--group", "runner-probe"], catalog).ExitCode);
            TestAssert.Equal(0, calls.Count);
            TestAssert.Equal(1, RunCaptured(["--skip-slow", "--only-slow"], catalog).ExitCode);
            TestAssert.Equal(0, calls.Count);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static IEnumerable<TestCase> UnexpectedEnumeration() => Enumerable.Range(0, 1)
        .Select<int, TestCase>(_ => throw new InvalidOperationException("Invalid arguments must not enumerate tests."));

    private static (int ExitCode, string Output) RunCaptured(string[] args, IEnumerable<TestCase> tests)
    {
        TextWriter previous = Console.Out;
        using var writer = new StringWriter();
        try
        {
            Console.SetOut(writer);
            return (TestRunner.Run(args, tests), writer.ToString());
        }
        finally { Console.SetOut(previous); }
    }
}
