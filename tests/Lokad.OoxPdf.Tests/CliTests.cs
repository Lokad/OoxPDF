using System.Diagnostics;

namespace Lokad.OoxPdf.Tests;

internal static class CliTests
{
    private static readonly Lazy<string> CliAssemblyPath = new(BuildCli);

    public static void CliConvertReturnsZeroOnSuccess()
    {
        string input = WriteBasicDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", input, output);

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "CLI should write the output PDF on success.");
    }

    public static void CliReturnsTwoForInvalidArguments()
    {
        CliResult result = RunCli("convert");

        TestAssert.Equal(2, result.ExitCode);
    }

    public static void CliReturnsOneForConversionFailure()
    {
        string missing = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");

        CliResult result = RunCli("convert", missing, output);

        TestAssert.Equal(1, result.ExitCode);
    }

    public static void CliStrictReturnsThreeWhenWarningsAreEmitted()
    {
        string input = WriteUnsupportedDocx();
        string output = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        string diagnostics = Path.ChangeExtension(Path.GetTempFileName(), ".json");

        CliResult result = RunCli("convert", input, output, "--diagnostics", diagnostics, "--strict");

        TestAssert.Equal(3, result.ExitCode);
        TestAssert.Contains("DOCX_UNSUPPORTED_COMMENTS", File.ReadAllText(diagnostics));
    }

    private static CliResult RunCli(params string[] args)
    {
        var start = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = FindRepositoryRoot(),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add(CliAssemblyPath.Value);
        foreach (string arg in args)
        {
            start.ArgumentList.Add(arg);
        }

        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Failed to start CLI process.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return new CliResult(process.ExitCode, output, error);
    }

    private static string BuildCli()
    {
        string repositoryRoot = FindRepositoryRoot();
        var start = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = repositoryRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add("build");
        start.ArgumentList.Add("src/Lokad.OoxPdf.Cli/Lokad.OoxPdf.Cli.csproj");
        start.ArgumentList.Add("--tl:off");
        start.ArgumentList.Add("--nologo");
        start.ArgumentList.Add("-v");
        start.ArgumentList.Add("minimal");

        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Failed to build CLI project.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"CLI build failed with exit code {process.ExitCode}.{Environment.NewLine}{output}{Environment.NewLine}{error}");
        }

        return Path.Combine(repositoryRoot, "src", "Lokad.OoxPdf.Cli", "bin", "Debug", "net10.0", "Lokad.OoxPdf.Cli.dll");
    }

    private static string FindRepositoryRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Lokad.OoxPdf.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Could not locate repository root.");
    }

    private static string WriteBasicDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """
        });
    }

    private static string WriteUnsupportedDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Text</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """
        });
    }

    private static string BasicContentTypes()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """;
    }

    private static string PackageRelationship()
    {
        return """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """;
    }

    private sealed record CliResult(int ExitCode, string Output, string Error);
}
