namespace Lokad.OoxPdf.Tests;

internal static class TestRunner
{
    private static readonly HashSet<string> SlowTests = new(StringComparer.Ordinal)
    {
        nameof(CliTests.CliConvertReturnsZeroOnSuccess),
        nameof(CliTests.CliReturnsTwoForInvalidArguments),
        nameof(CliTests.CliReturnsOneForConversionFailure),
        nameof(CliTests.CliStrictReturnsThreeWhenWarningsAreEmitted),
        nameof(PptxModelTests.PptxSceneBuilderBuildsResolvedNodeLists),
        nameof(PptxTypographyTests.PptxSyntheticRotatedTextBoxProducesTransform),
        nameof(PptxTypographyTests.PptxSyntheticTextBoxEmbedsFontAndDrawsGlyphs),
        nameof(FontTests.DiscoveryHeadersMatchFullLoadAcrossWindowsFonts)
    };

    public static int Run(params Action[] tests)
    {
        return Run([], tests);
    }

    public static int Run(string[] args, params Action[] tests)
    {
        return Run(args, tests.Select(test => new TestCase("default", test)));
    }

    public static int Run(string[] args, IEnumerable<TestCase> tests)
    {
        TestCase[] testCases = tests.ToArray();
        int passed = 0;
        int failed = 0;
        int skipped = 0;
        bool skipSlow = args.Contains("--skip-slow", StringComparer.Ordinal);
        bool onlySlow = args.Contains("--only-slow", StringComparer.Ordinal);
        if (skipSlow && onlySlow)
        {
            Console.WriteLine("Conflicting options: --skip-slow and --only-slow cannot be combined.");
            return 1;
        }
        bool list = args.Contains("--list", StringComparer.Ordinal);
        string? group = ReadOption(args, "--group");
        string? name = ReadOption(args, "--test");

        if (group is not null &&
            !testCases.Any(testCase => string.Equals(testCase.Group, group, StringComparison.Ordinal)))
        {
            string availableGroups = string.Join(", ", testCases
                .Select(testCase => testCase.Group)
                .Distinct(StringComparer.Ordinal)
                .Order(StringComparer.Ordinal));
            Console.WriteLine($"Unknown test group '{group}'. Available groups: {availableGroups}");
            return 1;
        }
        IReadOnlyList<string> catalogProblems = TestCatalogValidator.Validate(testCases);
        if (catalogProblems.Count != 0)
        {
            foreach (string problem in catalogProblems)
            {
                Console.WriteLine("Catalog problem: " + problem);
            }

            return 1;
        }

        List<TestCase> selected = new();
        foreach (TestCase testCase in testCases)
        {
            Action test = testCase.Action;
            bool isSlow = SlowTests.Contains(test.Method.Name);
            if (group is not null && !string.Equals(testCase.Group, group, StringComparison.Ordinal))
            {
                continue;
            }

            if (name is not null && !test.Method.Name.Contains(name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            selected.Add(testCase);
        }

        if ((group is not null || name is not null) && selected.Count == 0 && !list)
        {
            Console.WriteLine("No tests match the requested filter; refusing to report an empty pass.");
            return 1;
        }

        var results = new List<TestResult>();
        foreach (TestCase testCase in testCases)
        {
            Action test = testCase.Action;
            bool isSlow = SlowTests.Contains(test.Method.Name);
            if (group is not null && !string.Equals(testCase.Group, group, StringComparison.Ordinal))
            {
                continue;
            }

            if (name is not null && !test.Method.Name.Contains(name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (list)
            {
                Console.WriteLine($"{(isSlow ? "SLOW" : "FAST")} {testCase.Group} {test.Method.Name}");
                continue;
            }

            if ((skipSlow && isSlow) || (onlySlow && !isSlow))
            {
                skipped++;
                results.Add(new TestResult(testCase.Group, test.Method.Name, "skipped", 0, "slow filter"));
                Console.WriteLine($"SKIP {test.Method.Name} ({(isSlow ? "slow" : "fast")})");
                continue;
            }

            long start = Environment.TickCount64;
            try
            {
                test();
                passed++;
                results.Add(new TestResult(testCase.Group, test.Method.Name, "passed", Environment.TickCount64 - start, ""));
                Console.WriteLine($"PASS {test.Method.Name} ({Environment.TickCount64 - start} ms)");
            }
            catch (TestSkippedException ex)
            {
                skipped++;
                results.Add(new TestResult(testCase.Group, test.Method.Name, "skipped", Environment.TickCount64 - start, ex.Message));
                Console.WriteLine($"SKIP {test.Method.Name} ({Environment.TickCount64 - start} ms): {ex.Message}");
            }
            catch (Exception ex)
            {
                failed++;
                results.Add(new TestResult(testCase.Group, test.Method.Name, "failed", Environment.TickCount64 - start, ex.GetType().Name + ": " + ex.Message));
                Console.WriteLine($"FAIL {test.Method.Name} ({Environment.TickCount64 - start} ms): {ex.GetType().Name}: {ex.Message}");
            }
        }
        Console.WriteLine($"{passed} passed, {failed} failed, {skipped} skipped");
        string? reportPath = ReadOption(args, "--report");
        if (reportPath is not null)
        {
            var report = new System.Text.Json.Nodes.JsonObject
            {
                ["passed"] = passed,
                ["failed"] = failed,
                ["skipped"] = skipped,
                ["tests"] = new System.Text.Json.Nodes.JsonArray(results.Select(result => new System.Text.Json.Nodes.JsonObject
                {
                    ["group"] = result.Group,
                    ["name"] = result.Name,
                    ["outcome"] = result.Outcome,
                    ["milliseconds"] = result.Milliseconds,
                    ["detail"] = result.Detail,
                }).ToArray()),
            };

            System.IO.File.WriteAllText(reportPath, report.ToJsonString());
        }

        return failed == 0 ? 0 : 1;
    }

    private static string? ReadOption(string[] args, string option)
    {
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (string.Equals(args[i], option, StringComparison.Ordinal))
            {
                return args[i + 1];
            }
        }

        return null;
    }
}
