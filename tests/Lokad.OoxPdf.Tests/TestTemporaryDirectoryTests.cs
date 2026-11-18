namespace Lokad.OoxPdf.Tests;

internal static class TestTemporaryDirectoryTests
{
    public static void TestTemporaryDirectoryRemovesSuccessfulScratchAndRestoresEnvironment()
    {
        string[] names = ["TMP", "TEMP", "TMPDIR"];
        string?[] previous = names.Select(Environment.GetEnvironmentVariable).ToArray();
        string parent = Path.TrimEndingDirectorySeparator(Path.GetFullPath(Path.GetTempPath()));
        string sentinel = Path.Combine(parent, "keep-" + Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(sentinel, "Retained caller storage");
        string? scratch = null;
        try
        {
            TestAssert.Equal(0, TestTemporaryDirectory.Run(() =>
            {
                scratch = Path.TrimEndingDirectorySeparator(Path.GetFullPath(Path.GetTempPath()));
                TestAssert.Equal(Path.TrimEndingDirectorySeparator(parent), Path.GetDirectoryName(scratch));
                foreach (string name in names) TestAssert.Equal(scratch, Environment.GetEnvironmentVariable(name));
                Directory.CreateDirectory(Path.Combine(scratch, "nested"));
                File.WriteAllText(Path.Combine(scratch, "nested", "scratch.txt"), "Disposable test output");
                return 0;
            }));
            TestAssert.True(scratch is not null && !Directory.Exists(scratch), "Successful scratch must be removed.");
            TestAssert.Equal("Retained caller storage", File.ReadAllText(sentinel));
            for (int index = 0; index < names.Length; index++)
                TestAssert.Equal(previous[index], Environment.GetEnvironmentVariable(names[index]));
        }
        finally { File.Delete(sentinel); }
    }

    public static void TestTemporaryDirectoryRetainsFailedAndThrownRunsWithoutMaskingResults()
    {
        string[] names = ["TMP", "TEMP", "TMPDIR"];
        string?[] previous = names.Select(Environment.GetEnvironmentVariable).ToArray();
        string parent = Path.TrimEndingDirectorySeparator(Path.GetFullPath(Path.GetTempPath()));
        foreach (bool throws in new[] { false, true })
        {
            string? scratch = null;
            try
            {
                int Invoke() => TestTemporaryDirectory.Run(() =>
                {
                    scratch = Path.GetTempPath();
                    File.WriteAllText(Path.Combine(scratch, "failure.txt"), "Retained failure input");
                    if (throws) throw new InvalidOperationException("Original failure");
                    return 7;
                });
                if (throws)
                    TestAssert.Equal("Original failure", TestAssert.Throws<InvalidOperationException>(() => Invoke()).Message);
                else
                    TestAssert.Equal(7, Invoke());
                TestAssert.Equal("Retained failure input", File.ReadAllText(Path.Combine(scratch!, "failure.txt")));
                for (int index = 0; index < names.Length; index++)
                    TestAssert.Equal(previous[index], Environment.GetEnvironmentVariable(names[index]));
            }
            finally
            {
                if (scratch is not null && IsOwnedChild(scratch, parent)) Directory.Delete(scratch, recursive: true);
            }
        }
    }

    public static void TestTemporaryDirectoryRetainsAliasesAndLeavesExternalTargetsUntouched()
    {
        string parent = Path.GetTempPath();
        string target = Path.Combine(parent, "caller-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(target);
        File.WriteAllText(Path.Combine(target, "keep.txt"), "External target");
        string? scratch = null;
        bool linked = false;
        try
        {
            TestAssert.Equal(0, TestTemporaryDirectory.Run(() =>
            {
                scratch = Path.GetTempPath();
                try { Directory.CreateSymbolicLink(Path.Combine(scratch, "alias"), target); linked = true; }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
                return 0;
            }));
            if (!linked) TestAssert.Skip("Directory symbolic links are unavailable for this account.");
            TestAssert.True(Directory.Exists(scratch), "A linked child must retain the owned scratch directory.");
            TestAssert.Equal("External target", File.ReadAllText(Path.Combine(target, "keep.txt")));
        }
        finally
        {
            if (linked) Directory.Delete(Path.Combine(scratch!, "alias"));
            if (scratch is not null && IsOwnedChild(scratch, Path.TrimEndingDirectorySeparator(Path.GetFullPath(parent))) && Directory.Exists(scratch))
                Directory.Delete(scratch);
            Directory.Delete(target, recursive: true);
        }
    }

    private static bool IsOwnedChild(string path, string parent) =>
        string.Equals(Path.GetDirectoryName(Path.TrimEndingDirectorySeparator(Path.GetFullPath(path))), parent,
            OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal) &&
        Path.GetFileName(Path.TrimEndingDirectorySeparator(path)).StartsWith("ooxpdf-tests-", StringComparison.Ordinal);
}
