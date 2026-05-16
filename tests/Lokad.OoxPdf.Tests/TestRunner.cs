namespace Lokad.OoxPdf.Tests;

internal static class TestRunner
{
    public static int Run(params Action[] tests)
    {
        int passed = 0;
        int failed = 0;

        foreach (Action test in tests)
        {
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

        Console.WriteLine($"{passed} passed, {failed} failed");
        return failed == 0 ? 0 : 1;
    }
}
