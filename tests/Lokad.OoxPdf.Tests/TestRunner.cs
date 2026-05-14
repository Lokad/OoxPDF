namespace Lokad.OoxPdf.Tests;

internal static class TestRunner
{
    public static int Run(params Action[] tests)
    {
        int passed = 0;
        int failed = 0;

        foreach (Action test in tests)
        {
            try
            {
                test();
                passed++;
                Console.WriteLine($"PASS {test.Method.Name}");
            }
            catch (Exception ex)
            {
                failed++;
                Console.WriteLine($"FAIL {test.Method.Name}: {ex.Message}");
            }
        }

        Console.WriteLine($"{passed} passed, {failed} failed");
        return failed == 0 ? 0 : 1;
    }
}
