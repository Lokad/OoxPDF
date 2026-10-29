using System.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static partial class PublicApiTests
{
    // These checks previously loaded the library into PowerShell, whose runtime
    // can be older than the selected library TFM. Keep API checks in this runner.
    public static void ConvertRejectsDirectoryAliasDestinations()
    {
        foreach (string kind in new[] { "direct", "ancestor", "target-ancestor" })
        {
            InPublicationDirectory(root =>
            {
                string real = Directory.CreateDirectory(Path.Combine(root, "real")).FullName;
                string nested = kind == "direct" ? real : Directory.CreateDirectory(Path.Combine(real, "nested")).FullName;
                string target = real;
                if (kind == "target-ancestor")
                {
                    target = Path.Combine(root, "target-alias");
                    CreatePublicationLink("Directory", target, real);
                    target = Path.Combine(target, "nested");
                }
                string alias = Path.Combine(root, "alias");
                CreatePublicationLink("Directory", alias, target);
                string destination = Path.Combine(kind == "ancestor" ? Path.Combine(alias, "nested") : alias, "source.docx");
                string input = Path.Combine(nested, "source.docx");
                byte[] source = ReadMinimalDocxBytes("<w:p/>");
                File.WriteAllBytes(input, source);
                string untouched = Path.Combine(nested, "result.pdf");
                File.WriteAllText(untouched, "Owned PDF sentinel");

                TestAssert.Throws<ArgumentException>(() => OoxPdfConverter.Convert(input, destination));

                TestAssert.True(File.ReadAllBytes(input).SequenceEqual(source), kind + ": input changed");
                TestAssert.Equal("Owned PDF sentinel", File.ReadAllText(untouched));
                AssertNoPublicationStaging(nested);
            });
        }
    }

    public static void ConvertReplacesDistinctLinkedDestinations()
    {
        foreach (string kind in new[] { "directory", "hard-link", "output-symbolic-link", "input-hard-link" })
        {
            InPublicationDirectory(root =>
            {
                byte[] source = ReadMinimalDocxBytes("<w:p/>");
                string input = Path.Combine(root, "source.docx");
                string destination = Path.Combine(root, "result.pdf");
                File.WriteAllBytes(input, source);
                if (kind == "directory")
                {
                    string alias = Path.Combine(root, "alias");
                    CreatePublicationLink("Directory", alias, root);
                    File.WriteAllText(destination, "Owned PDF sentinel");
                    destination = Path.Combine(alias, "result.pdf");
                }
                else if (kind == "input-hard-link")
                {
                    destination = input;
                    input = Path.Combine(root, "input-link.docx");
                    CreatePublicationLink("HardLink", input, destination);
                }
                else
                {
                    CreatePublicationLink(kind == "hard-link" ? "HardLink" : "SymbolicLink", destination, input);
                }

                OoxPdfConverter.Convert(input, destination);

                TestAssert.True(File.ReadAllBytes(input).SequenceEqual(source), kind + ": input changed");
                TestAssert.True(File.ReadAllBytes(destination).AsSpan().StartsWith("%PDF-"u8), kind + ": output is not PDF");
                AssertNoPublicationStaging(root);
            });
        }
    }

    public static void ConvertRejectsLinkedInputReferentDestinations()
    {
        foreach (string kind in new[] { "referent", "same-entry", "chain", "target-ancestor" })
        {
            InPublicationDirectory(root =>
            {
                byte[] source = ReadMinimalDocxBytes("<w:p/>");
                string input = Path.Combine(root, "source.docx");
                File.WriteAllBytes(input, source);
                string target = input;
                if (kind == "target-ancestor")
                {
                    string parent = Path.Combine(root, "parent-link");
                    CreatePublicationLink("Directory", parent, root);
                    target = Path.Combine(parent, "source.docx");
                }
                if (kind == "chain")
                {
                    string middle = Path.Combine(root, "middle.docx");
                    CreatePublicationLink("SymbolicLink", middle, target);
                    target = middle;
                }
                string alias = Path.Combine(root, "input-link.docx");
                CreatePublicationLink("SymbolicLink", alias, target);

                TestAssert.Throws<ArgumentException>(() => OoxPdfConverter.Convert(alias, kind == "same-entry" ? alias : input));

                TestAssert.True(File.ReadAllBytes(input).SequenceEqual(source), kind + ": referent changed");
                AssertNoPublicationStaging(root);
            });
        }
    }

    private static void InPublicationDirectory(Action<string> check)
    {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "publication-" + Guid.NewGuid().ToString("N"))).FullName;
        try { check(root); }
        finally
        {
            // Delete link entries before recursive cleanup, including cycles to
            // this owned root. Never traverse a directory-link target.
            foreach (string entry in Directory.EnumerateDirectories(root).Where(entry =>
                (File.GetAttributes(entry) & FileAttributes.ReparsePoint) != 0).ToArray())
            {
                Directory.Delete(entry);
            }
            Directory.Delete(root, recursive: true);
        }
    }

    private static void AssertNoPublicationStaging(string directory) =>
        TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);

    private static void CreatePublicationLink(string kind, string path, string target)
    {
        if (kind == "Directory" && !OperatingSystem.IsWindows()) { Directory.CreateSymbolicLink(path, target); return; }
        if (kind == "SymbolicLink")
        {
            try { File.CreateSymbolicLink(path, target); }
            catch (UnauthorizedAccessException) when (OperatingSystem.IsWindows()) { TestAssert.Skip("Windows symbolic-link privilege unavailable."); }
            return;
        }
        // Junctions preserve Windows coverage without symlink privilege; the
        // BCL has no portable hard-link creation API. Validation already uses pwsh.
        var start = new ProcessStartInfo("pwsh")
        {
            UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardOutput = true, RedirectStandardError = true
        };
        foreach (string argument in new[] { "-NoProfile", "-NonInteractive", "-Command",
            "New-Item -ItemType " + (kind == "Directory" ? "Junction" : "HardLink") +
            " -Path '" + path.Replace("'", "''") + "' -Target '" + target.Replace("'", "''") + "' -ErrorAction Stop | Out-Null" })
        { start.ArgumentList.Add(argument); }
        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Cannot start link fixture helper.");
        Task<string> output = process.StandardOutput.ReadToEndAsync();
        Task<string> error = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(15_000))
        {
            process.Kill(entireProcessTree: true); process.WaitForExit();
            throw new TimeoutException("Link fixture helper timed out.");
        }
        TestAssert.True(process.ExitCode == 0, output.GetAwaiter().GetResult() + error.GetAwaiter().GetResult());
    }
}
