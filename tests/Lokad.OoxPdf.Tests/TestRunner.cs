namespace Lokad.OoxPdf.Tests;

internal static class TestRunner
{
    private static readonly HashSet<string> SlowTests = new(StringComparer.Ordinal)
    {
        nameof(CliTests.CliConvertReturnsZeroOnSuccess),
        nameof(CliTests.CliReturnsTwoForInvalidArguments),
        nameof(CliTests.CliReturnsOneForConversionFailure),
        nameof(CliTests.CliStrictReturnsThreeWhenWarningsAreEmitted),
        nameof(PptxTests.PptxSceneBuilderBuildsResolvedNodeLists),
        nameof(PptxTests.PptxSyntheticRotatedTextBoxProducesTransform),
        nameof(PptxTests.PptxSyntheticTextBoxEmbedsFontAndDrawsGlyphs)
    };

    public static int Run(params Action[] tests)
    {
        return Run([], tests);
    }

    public static int Run(string[] args, params Action[] tests)
    {
        int passed = 0;
        int failed = 0;
        int skipped = 0;
        bool skipSlow = args.Contains("--skip-slow", StringComparer.Ordinal);
        bool onlySlow = args.Contains("--only-slow", StringComparer.Ordinal);
        bool list = args.Contains("--list", StringComparer.Ordinal);

        foreach (Action test in tests)
        {
            bool isSlow = SlowTests.Contains(test.Method.Name);
            if (list)
            {
                Console.WriteLine($"{(isSlow ? "SLOW" : "FAST")} {test.Method.Name}");
                continue;
            }

            if ((skipSlow && isSlow) || (onlySlow && !isSlow))
            {
                skipped++;
                Console.WriteLine($"SKIP {test.Method.Name} ({(isSlow ? "slow" : "fast")})");
                continue;
            }

            long start = Environment.TickCount64;
            try
            {
                test();
                passed++;
                Console.WriteLine($"PASS {test.Method.Name} ({Environment.TickCount64 - start} ms)");
            }
            catch (Exception ex)
            {
                failed++;
                Console.WriteLine($"FAIL {test.Method.Name} ({Environment.TickCount64 - start} ms): {ex.Message}");
            }
        }

        if (list)
        {
            return 0;
        }

        Console.WriteLine($"{passed} passed, {failed} failed, {skipped} skipped");
        return failed == 0 ? 0 : 1;
    }
}
