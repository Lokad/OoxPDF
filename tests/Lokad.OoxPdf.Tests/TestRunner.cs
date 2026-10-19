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
        Dictionary<string, string?>? options = ReadOptions(args);
        if (options is null) return 1;
        TestCase[] testCases = tests.ToArray();
        int passed = 0;
        int failed = 0;
        int skipped = 0;
        bool skipSlow = options.ContainsKey("--skip-slow");
        bool onlySlow = options.ContainsKey("--only-slow");
        if (skipSlow && onlySlow)
        {
            Console.WriteLine("Conflicting options: --skip-slow and --only-slow cannot be combined.");
            return 1;
        }
        bool list = options.ContainsKey("--list");
        string? group = options.GetValueOrDefault("--group");
        string? name = options.GetValueOrDefault("--test");
        string? testSetPath = options.GetValueOrDefault("--test-set");
        HashSet<string>? exactNames = null;
        if (options.ContainsKey("--test-set"))
        {
            if (testSetPath is null || testSetPath.StartsWith("--", StringComparison.Ordinal) ||
                group is not null || name is not null)
            {
                Console.WriteLine("--test-set requires a JSON array file and cannot be combined with --group or --test.");
                return 1;
            }
            try
            {
                string[]? requested = System.Text.Json.JsonSerializer.Deserialize<string[]>(File.ReadAllText(testSetPath));
                if (requested is null || requested.Length == 0 || requested.Any(string.IsNullOrWhiteSpace))
                {
                    Console.WriteLine("The exact test set must contain nonempty method names.");
                    return 1;
                }
                exactNames = new HashSet<string>(requested, StringComparer.Ordinal);
                foreach (string requestedName in exactNames)
                {
                    if (testCases.Count(testCase => testCase.Action.Method.Name == requestedName) != 1)
                    {
                        Console.WriteLine($"Unknown or ambiguous exact test name '{requestedName}'.");
                        return 1;
                    }
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException)
            {
                Console.WriteLine("Cannot read exact test set: " + ex.Message);
                return 1;
            }
        }

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
            if (group is not null && !string.Equals(testCase.Group, group, StringComparison.Ordinal))
            {
                continue;
            }

            if (name is not null && !test.Method.Name.Contains(name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (exactNames is not null && !exactNames.Contains(test.Method.Name))
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
        foreach (TestCase testCase in selected)
        {
            Action test = testCase.Action;
            bool isSlow = SlowTests.Contains(test.Method.Name);

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
        string? reportPath = options.GetValueOrDefault("--report");
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

    private static Dictionary<string, string?>? ReadOptions(string[] args)
    {
        var options = new Dictionary<string, string?>(StringComparer.Ordinal);
        for (int i = 0; i < args.Length; i++)
        {
            string option = args[i];
            bool needsValue = option is "--group" or "--test" or "--test-set" or "--report";
            if (!needsValue && option is not ("--skip-slow" or "--only-slow" or "--list"))
            {
                Console.WriteLine($"Unknown argument '{option}'. Use --test for a name filter or --test-set for exact methods.");
                return null;
            }
            string? value = null;
            if (needsValue)
            {
                if (++i >= args.Length || string.IsNullOrWhiteSpace(args[i]) || args[i].StartsWith("--", StringComparison.Ordinal))
                {
                    Console.WriteLine($"{option} requires a value.");
                    return null;
                }
                value = args[i];
            }
            if (!options.TryAdd(option, value))
            {
                Console.WriteLine($"Repeated option '{option}'.");
                return null;
            }
        }
        return options;
    }
}
