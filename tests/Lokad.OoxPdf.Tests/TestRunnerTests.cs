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
        var noSlow = RunCaptured(["--group", "runner-probe", "--only-slow"], catalog);
        TestAssert.Equal(1, noSlow.ExitCode);
        TestAssert.Contains("The speed filter excludes every selected test.", noSlow.Output);
        TestAssert.Equal(0, RunCaptured(["--group", "runner-probe", "--only-slow", "--list"], catalog).ExitCode);
        string slowName = nameof(DocxCommentsTests.DocxWordCompatibleMixedBalloonWrapsParagraphBodiesIndependently);
        var noFast = RunCaptured(["--test", slowName, "--skip-slow"], catalog);
        TestAssert.Equal(1, noFast.ExitCode);
        TestAssert.Contains("The speed filter excludes every selected test.", noFast.Output);
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
            var listWithReport = RunCaptured(["--list", "--report", report], UnexpectedEnumeration());
            TestAssert.Equal(1, listWithReport.ExitCode);
            TestAssert.Contains("--list and --report cannot be combined", listWithReport.Output);
            TestAssert.True(!File.Exists(report), "Listing with reporting must not create a report.");

            File.WriteAllText(report, "previous report");
            var reportWithList = RunCaptured(["--report", report, "--list"], UnexpectedEnumeration());
            TestAssert.Equal(1, reportWithList.ExitCode);
            TestAssert.Equal("previous report", File.ReadAllText(report));

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

            var noSlowExact = RunCaptured(["--test-set", set, "--only-slow"], catalog);
            TestAssert.Equal(1, noSlowExact.ExitCode);
            TestAssert.Contains("The speed filter excludes every selected test.", noSlowExact.Output);
            TestAssert.Equal(0, calls.Count);

            File.WriteAllText(set, JsonSerializer.Serialize(new[] { slowName, second.Method.Name }));
            TestAssert.Equal(0, RunCaptured(["--test-set", set, "--skip-slow", "--report", report], catalog).ExitCode);
            TestAssert.True(calls.SequenceEqual([2]), "A mixed exact set executes its one fast method.");
            using var mixed = JsonDocument.Parse(File.ReadAllText(report));
            TestAssert.Equal(1, mixed.RootElement.GetProperty("passed").GetInt32());
            TestAssert.Equal(1, mixed.RootElement.GetProperty("skipped").GetInt32());
            calls.Clear();

            File.WriteAllText(set, JsonSerializer.Serialize(new[] { second.Method.Name, first.Method.Name, second.Method.Name }));
            TestAssert.Equal(0, RunCaptured(["--test-set", set], catalog).ExitCode);
            TestAssert.True(calls.SequenceEqual([1, 2]), "Repeated exact requests execute once each in catalogue order.");
            calls.Clear();

            var ambiguous = RunCaptured(["--test-set", set], [.. catalog, new TestCase("runner-probe", second)]);
            TestAssert.Equal(1, ambiguous.ExitCode);
            TestAssert.Contains("Duplicate test registration", ambiguous.Output);
            TestAssert.Equal(0, calls.Count);

            File.WriteAllText(set, JsonSerializer.Serialize(new[] { first.Method.Name, "NotARegisteredExactMethod" }));
            var unknown = RunCaptured(["--test-set", set], catalog);
            TestAssert.Equal(1, unknown.ExitCode);
            TestAssert.Contains("Unknown exact test name 'NotARegisteredExactMethod'", unknown.Output);
            TestAssert.Equal(0, calls.Count);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    public static void TestRunnerRejectsDuplicateRegistrationsBeforeListingOrExecutingGroups()
    {
        int calls = 0;
        Action probe = () => calls++;
        TestCase[] catalog = [.. TestCatalog.All, new("runner-probe", probe), new("other-probe", probe)];
        foreach (string[] args in new string[][] { ["--list"], ["--group", "runner-probe"] })
        {
            var result = RunCaptured(args, catalog);
            TestAssert.Equal(1, result.ExitCode);
            TestAssert.Contains($"Duplicate test registration '{probe.Method.Name}'", result.Output);
            TestAssert.Equal(0, calls);
        }
    }

    public static void PublicCasePathsIgnoreWorkingDirectoryFixtures()
    {
        const string fixture = "docx-tables.docx";
        string pinned = TestFixtures.PublicCasePath(fixture);
        byte[] pinnedBytes = File.ReadAllBytes(pinned);
        string workingDirectory = Directory.CreateTempSubdirectory("working-directory-").FullName;
        string cases = Path.Combine(workingDirectory, "tests", "Lokad.OoxPdf.Tests", "Cases");
        Directory.CreateDirectory(cases);
        string alternate = TestFixtures.PublicCasePath("docx-basic-paragraphs.docx");
        File.Copy(alternate, Path.Combine(cases, fixture));
        const string missing = "missing-from-bound-public-case-tree.docx";
        File.Copy(alternate, Path.Combine(cases, missing));

        string previous = Environment.CurrentDirectory;
        try
        {
            Environment.CurrentDirectory = workingDirectory;
            string selected = TestFixtures.PublicCasePath(fixture);
            TestAssert.Equal(pinned, selected);
            TestAssert.True(pinnedBytes.AsSpan().SequenceEqual(File.ReadAllBytes(selected)),
                "The working directory fixture must never replace the bound input.");

            string missingPath = TestFixtures.PublicCasePath(missing);
            FileNotFoundException exception = TestAssert.Throws<FileNotFoundException>(() => File.ReadAllBytes(missingPath));
            TestAssert.Equal(missingPath, exception.FileName);
        }
        finally
        {
            Environment.CurrentDirectory = previous;
        }
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
