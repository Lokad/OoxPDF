namespace Lokad.OoxPdf.Tests;

internal static class CliTests
{
    private static readonly Lazy<string> CliAssemblyPath = new(() =>
        TestToolProcess.ResolveAssembly("OOXPDF_TEST_CLI_ASSEMBLY", "src/Lokad.OoxPdf.Cli"));

    public static void CliConvertReturnsZeroOnSuccess()
    {
        string input = WriteBasicDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", input, output);

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "CLI should write the output PDF on success.");
    }

    public static void CliReturnsTwoForInvalidArguments()
    {
        TestToolResult result = RunCli("convert");

        TestAssert.Equal(2, result.ExitCode);
    }

    public static void CliLongArgumentRefusalDrainsBothOutputStreams()
    {
        string argument = "--" + new string('x', 20_000);
        TestToolResult result = RunCli("convert", "absent.docx", "unused.pdf", argument);
        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains(argument, result.Error);
    }

    public static void CliReturnsOneForConversionFailure()
    {
        string missing = TestFixtures.NewTemporaryPath(".docx");
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", missing, output);

        TestAssert.Equal(1, result.ExitCode);
    }

    public static void CliStrictReturnsThreeWhenWarningsAreEmitted()
    {
        string input = WriteUnsupportedDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");
        string diagnostics = TestFixtures.NewTemporaryPath(".json");

        TestToolResult result = RunCli("convert", input, output, "--diagnostics", diagnostics, "--strict");

        TestAssert.Equal(3, result.ExitCode);
        TestAssert.Contains("DOCX_UNSUPPORTED_OLE_OBJECT", File.ReadAllText(diagnostics));

        byte[] expectedPdf = File.ReadAllBytes(output);
        var strictOnly = RunCli("convert", input, output, "--strict");
        TestAssert.Equal(3, strictOnly.ExitCode);
        TestAssert.True(expectedPdf.AsSpan().SequenceEqual(File.ReadAllBytes(output)),
            "Strict exit handling must preserve the PDF without diagnostic samples.");

        var plain = RunCli("convert", input, output);
        TestAssert.Equal(0, plain.ExitCode);
        TestAssert.True(expectedPdf.AsSpan().SequenceEqual(File.ReadAllBytes(output)),
            "Conversion without a diagnostic observer must preserve the PDF.");
    }

    public static void CliStrictCommentWarningsFollowTheSelectedView()
    {
        // Isolate strict exit handling from host font availability. Anchor prose
        // and comment bodies are covered by the reader/rendering view tests.
        string input = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = BasicContentTypes(),
            ["_rels/.rels"] = PackageRelationship(),
            ["word/document.xml"] = """
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body><w:p><w:commentRangeStart w:id="1"/>
                    <w:r><w:commentReference w:id="1"/></w:r><w:commentRangeEnd w:id="1"/>
                  </w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """
        });
        foreach (string mode in new[] { "final", "original", "simple", "all" })
        {
            string output = TestFixtures.NewTemporaryPath(".pdf");
            string diagnostics = TestFixtures.NewTemporaryPath(".json");
            TestToolResult result = RunCli("convert", input, output, "--diagnostics", diagnostics,
                "--strict", "--docx-markup", mode);
            bool requested = mode is "simple" or "all";
            TestAssert.Equal(requested ? 3 : 0, result.ExitCode);
            TestAssert.True(File.Exists(output) && new FileInfo(output).Length > 0,
                "Strict mode publishes the converted document even when requested markup is approximated.");
            string json = File.ReadAllText(diagnostics);
            TestAssert.True(!json.Contains("DOCX_UNSUPPORTED_COMMENTS", StringComparison.Ordinal),
                "A selected view does not make supported comment markup an unsupported feature.");
            TestAssert.Equal(requested, json.Contains("DOCX_APPROXIMATED_COMMENTS", StringComparison.Ordinal));
        }
    }

    public static void CliAcceptsDocxMarkupArgumentForDocxInput()
    {
        string input = WriteBasicDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", input, output, "--docx-markup", "all");

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "CLI should write the output PDF when DOCX markup mode is supplied.");
    }

    public static void CliAcceptsDocxMarkupGeometryArgumentForDocxInput()
    {
        string input = WriteBasicDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", input, output, "--docx-markup", "all", "--docx-markup-geometry", "reserve-margin");

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "CLI should write the output PDF when DOCX markup geometry is supplied.");
    }

    public static void CliAcceptsDocxWordCompatibleMarkupGeometryArgumentForDocxInput()
    {
        string input = WriteBasicDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", input, output, "--docx-markup", "all", "--docx-markup-geometry", "word-compatible");

        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(output), "CLI should write the output PDF when Word-compatible DOCX markup geometry is supplied.");
    }

    public static void CliRejectsInvalidDocxMarkupArgument()
    {
        string input = WriteBasicDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", input, output, "--docx-markup", "invalid");

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("Invalid DOCX markup mode", result.Error);
    }

    public static void CliRejectsInvalidDocxMarkupGeometryArgument()
    {
        string input = WriteBasicDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", input, output, "--docx-markup-geometry", "invalid");

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("Invalid DOCX markup geometry mode", result.Error);
    }

    public static void CliRejectsDocxMarkupArgumentForPptxInput()
    {
        string input = TestFixtures.NewTemporaryPath(".pptx");
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", input, output, "--docx-markup", "simple");

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("--docx-markup can only be used with DOCX input.", result.Error);
    }

    public static void CliRejectsDocxMarkupGeometryArgumentForPptxInput()
    {
        string input = TestFixtures.NewTemporaryPath(".pptx");
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", input, output, "--docx-markup-geometry", "reserve-margin");

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("--docx-markup-geometry can only be used with DOCX input.", result.Error);
    }

    public static void CliRejectsDiagnosticsPathEqualToInput()
    {
        string input = WriteBasicDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");
        long beforeLength = new FileInfo(input).Length;

        TestToolResult result = RunCli("convert", input, output, "--diagnostics", input);

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("Diagnostics output path must be different", result.Error);
        TestAssert.Equal(beforeLength, new FileInfo(input).Length);
    }

    public static void CliRejectsDiagnosticsPathEqualToOutput()
    {
        string input = WriteBasicDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");

        TestToolResult result = RunCli("convert", input, output, "--diagnostics", output);

        TestAssert.Equal(2, result.ExitCode);
        TestAssert.Contains("Diagnostics output path must be different", result.Error);
    }

    public static void CliAllowsCaseDistinctDiagnosticsOnCaseSensitiveFileSystems()
    {
        string input = WriteBasicDocx();
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-cli-case-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            string output = Path.Combine(directory, "out.pdf");
            string diagnostics = Path.Combine(directory, "OUT.PDF");
            File.WriteAllText(output, "destination sentinel");
            if (File.Exists(diagnostics)) { TestAssert.Skip("The temporary filesystem is case-insensitive."); }
            TestToolResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);
            TestAssert.Equal(0, result.ExitCode);
            TestAssert.True(File.ReadAllText(output).StartsWith("%PDF-", StringComparison.Ordinal), "CLI publishes the PDF.");
            TestAssert.True(File.ReadAllText(diagnostics).TrimStart().StartsWith("[", StringComparison.Ordinal), "CLI publishes separate JSON diagnostics.");
            TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
            File.Delete(input);
        }
    }

    public static void CliPreservesInputAndPdfForCaseAliases()
    {
        string original = WriteBasicDocx();
        byte[] source = File.ReadAllBytes(original);
        File.Delete(original);
        string directory = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ooxpdf-cli-alias-" + Guid.NewGuid().ToString("N"))).FullName;
        try
        {
            foreach (string target in new[] { "input", "output", "unresolved" })
            {
                string run = Directory.CreateDirectory(Path.Combine(directory, target)).FullName;
                string input = Path.Combine(run, "source.docx");
                string output = Path.Combine(run, "out.pdf");
                string diagnostics = Path.Combine(run, target == "input" ? "SOURCE.DOCX" : "OUT.PDF");
                File.WriteAllBytes(input, source);
                if (target != "unresolved") { File.WriteAllText(output, "Existing PDF sentinel"); }
                bool aliases = target == "unresolved" || File.Exists(diagnostics);
                if (!aliases) { File.WriteAllText(diagnostics, "Independent diagnostic destination"); }
                TestToolResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);
                TestAssert.Equal(aliases ? 2 : 0, result.ExitCode);
                TestAssert.True(source.SequenceEqual(File.ReadAllBytes(input)), "Diagnostics must preserve the input.");
                if (target == "unresolved")
                {
                    TestAssert.True(!File.Exists(output) && !File.Exists(diagnostics), "Unresolved case-only outputs refuse before conversion.");
                }
                else if (aliases)
                {
                    TestAssert.Equal("Existing PDF sentinel", File.ReadAllText(output));
                }
                else
                {
                    TestAssert.True(File.ReadAllText(output).StartsWith("%PDF-", StringComparison.Ordinal), "Distinct output remains a PDF.");
                    TestAssert.True(File.ReadAllText(diagnostics).TrimStart().StartsWith("[", StringComparison.Ordinal), "Distinct diagnostics remain JSON.");
                }
                TestAssert.Equal(0, Directory.GetFiles(run, "*.tmp-*").Length);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    public static void CliReportsMalformedNumericInputAsConversionFailure()
    {
        string input = WriteMalformedNumericDocx();
        string output = TestFixtures.NewTemporaryPath(".pdf");
        string diagnostics = TestFixtures.NewTemporaryPath(".json");

        TestToolResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);

        TestAssert.Equal(1, result.ExitCode);
        TestAssert.True(File.Exists(diagnostics), "Malformed numeric input should still write diagnostics JSON.");
    }

    public static void CliRejectsEarlyFailureDiagnosticsAliasingInput()
        => AssertEarlyFailureDiagnosticAlias(outputAlias: false);

    public static void CliRejectsEarlyFailureDiagnosticsAliasingOutput()
        => AssertEarlyFailureDiagnosticAlias(outputAlias: true);

    private static void AssertEarlyFailureDiagnosticAlias(bool outputAlias)
    {
        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-cli-alias-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            // Detection failure, malformed package and missing input must all reject
            // the alias before either success or failure diagnostics can be written.
            foreach (string name in new[] { "unsupported.txt", "malformed.docx", "missing.docx" })
            {
                string input = Path.Combine(directory, name);
                string output = Path.Combine(directory, "output.pdf");
                bool exists = name != "missing.docx";
                if (exists) { File.WriteAllText(input, "original input"); }
                File.WriteAllText(output, "original destination");
                string diagnostics = outputAlias ? output : Path.Combine(directory, ".", name);
                TestToolResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);
                TestAssert.Equal(2, result.ExitCode);
                TestAssert.Contains("Diagnostics output path must be different", result.Error);
                TestAssert.Equal("original destination", File.ReadAllText(output));
                TestAssert.Equal(exists, File.Exists(input));
                if (exists) { TestAssert.Equal("original input", File.ReadAllText(input)); }
                TestAssert.Equal(0, Directory.GetFiles(directory, "*.tmp-*").Length);
            }
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    public static void CliWritesSeparateDiagnosticsForUnsupportedInputWithoutChangingFiles()
    {
        string directory = Path.Combine(Path.GetTempPath(), "ooxpdf-cli-failure-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            string input = Path.Combine(directory, "unsupported.txt");
            string output = Path.Combine(directory, "output.pdf");
            string diagnostics = Path.Combine(directory, "diagnostics.json");
            File.WriteAllText(input, "original input");
            File.WriteAllText(output, "original destination");
            TestToolResult result = RunCli("convert", input, output, "--diagnostics", diagnostics);
            TestAssert.Equal(1, result.ExitCode);
            TestAssert.Equal("original input", File.ReadAllText(input));
            TestAssert.Equal("original destination", File.ReadAllText(output));
            TestAssert.Contains("OOXML_CONVERSION_FAILED", File.ReadAllText(diagnostics));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    private static string WriteMalformedNumericDocx()
    {
        return TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
                </Types>
                """,
            ["_rels/.rels"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
                """,
            ["word/document.xml"] = """
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p><w:pPr><w:spacing w:before="notanumber"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
                  </w:body>
                </w:document>
                """,
        });
    }

    private static TestToolResult RunCli(params string[] args) =>
        TestToolProcess.Run(CliAssemblyPath.Value, args);

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
                    <w:p><w:r><w:t>Text</w:t><w:object/></w:r></w:p>
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
}
