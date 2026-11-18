namespace Lokad.OoxPdf.Tests;

// One console invocation owns this leaf. Nested runner probes use the same
// storage; individual fixtures need no registration or cleanup ledger.
internal static class TestTemporaryDirectory
{
    public static int Run(Func<int> run)
    {
        string parent = Path.TrimEndingDirectorySeparator(Path.GetFullPath(Path.GetTempPath()));
        RequirePhysicalAncestors(parent);
        string directory = Directory.CreateTempSubdirectory("ooxpdf-tests-").FullName;
        string[] names = ["TMP", "TEMP", "TMPDIR"];
        string?[] previous = names.Select(Environment.GetEnvironmentVariable).ToArray();
        int result = 1;
        try
        {
            foreach (string name in names) Environment.SetEnvironmentVariable(name, directory);
            result = run();
            return result;
        }
        finally
        {
            for (int index = 0; index < names.Length; index++)
                Environment.SetEnvironmentVariable(names[index], previous[index]);
            if (result != 0)
                Console.Error.WriteLine("Test temporary files retained: " + directory);
            else
            {
                try
                {
                    // Validate the absolute deletion target and its complete tree
                    // before removal. A fixture-created alias retains the leaf.
                    if (!string.Equals(Path.GetDirectoryName(Path.GetFullPath(directory)), parent,
                            OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
                        throw new IOException("Temporary storage escaped its declared parent.");
                    RequirePhysicalAncestors(directory);
                    RequirePhysicalTree(directory);
                    Directory.Delete(directory, recursive: true);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // Cleanup cannot turn a passing semantic result into failure.
                    Console.Error.WriteLine("Test temporary files retained: " + directory + ". " + ex.Message);
                }
            }
        }
    }

    private static void RequirePhysicalAncestors(string path)
    {
        for (DirectoryInfo? directory = new(path); directory is not null; directory = directory.Parent)
            if ((directory.Attributes & FileAttributes.ReparsePoint) != 0)
                throw new IOException("Test temporary storage requires physical directories.");
    }

    private static void RequirePhysicalTree(string path)
    {
        foreach (FileSystemInfo entry in new DirectoryInfo(path).EnumerateFileSystemInfos())
        {
            if ((entry.Attributes & FileAttributes.ReparsePoint) != 0)
                throw new IOException("A test temporary alias prevents cleanup.");
            if (entry is DirectoryInfo child) RequirePhysicalTree(child.FullName);
        }
    }
}
