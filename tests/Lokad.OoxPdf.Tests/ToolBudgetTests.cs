using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace Lokad.OoxPdf.Tests;

internal static class ToolBudgetTests
{
    private static readonly Lazy<string> InspectAssemblyPath = new(() => BuildTool("tools/Lokad.OoxPdf.PdfInspect/Lokad.OoxPdf.PdfInspect.csproj", "Lokad.OoxPdf.PdfInspect.dll"));
    private static readonly Lazy<string> VisualDiffAssemblyPath = new(() => BuildTool("tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj", "Lokad.OoxPdf.VisualDiff.dll"));
    private static readonly Lazy<string> RasterizerAssemblyPath = new(() => BuildTool("tools/Lokad.OoxPdf.PdfiumRasterizer/Lokad.OoxPdf.PdfiumRasterizer.csproj", "Lokad.OoxPdf.PdfiumRasterizer.dll"));

    public static void InspectAggregateDecodedBudgetSkipsBeyondQuota()
    {
        // R21: three individually legal 10 KB streams exceed a 15 KB aggregate quota.
        // The first decodes; the rest keep raw bytes with an aggregate skip note.
        string input = WriteMultiFlatePdf();
        ToolResult result = RunTool(InspectAssemblyPath.Value, input, "--max-total-decoded-bytes", "15000");
        TestAssert.Equal(0, result.ExitCode);
        TestAssert.Contains("decoded", result.Output);
        TestAssert.Contains("aggregate", result.Output);
    }

    public static void InspectDefaultQuotaDecodesSmallStreams()
    {
        // R21: the default aggregate quota leaves small inspections untouched.
        string input = WriteMultiFlatePdf();
        ToolResult result = RunTool(InspectAssemblyPath.Value, input);
        TestAssert.Equal(0, result.ExitCode);
        TestAssert.Contains("decoded raw deflate", result.Output);
        TestAssert.DoesNotContain("aggregate", result.Output);
    }

    public static void InspectObjectCapRejects()
    {
        // R21: the aggregate object quota fails fast instead of parsing unboundedly.
        string input = WriteMultiFlatePdf();
        ToolResult result = RunTool(InspectAssemblyPath.Value, input, "--max-objects", "2");
        TestAssert.Equal(1, result.ExitCode);
        TestAssert.Contains("object count", result.Output + result.Error);
    }

    public static void VisualDiffCombinedPixelQuotaRejects()
    {
        // R21: a 10k+10k pixel pair exceeds a 100-pixel combined quota before either
        // RGBA buffer is decoded.
        string referenceDirectory = NewTempDirectory();
        string candidateDirectory = NewTempDirectory();
        WritePngPair(referenceDirectory, candidateDirectory);
        string outputDirectory = NewTempDirectory();
        ToolResult result = RunTool(VisualDiffAssemblyPath.Value, referenceDirectory, candidateDirectory, outputDirectory, "--max-combined-pixels", "100");
        TestAssert.Equal(1, result.ExitCode);
        TestAssert.Contains("exceeds the maximum", result.Output + result.Error);
    }

    public static void VisualDiffDefaultQuotaAcceptsSmallPair()
    {
        // R21: the default combined quota leaves small comparisons untouched.
        string referenceDirectory = NewTempDirectory();
        string candidateDirectory = NewTempDirectory();
        WritePngPair(referenceDirectory, candidateDirectory);
        string outputDirectory = NewTempDirectory();
        ToolResult result = RunTool(VisualDiffAssemblyPath.Value, referenceDirectory, candidateDirectory, outputDirectory);
        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(Path.Combine(outputDirectory, "metrics.json")), "VisualDiff must write metrics on success.");
    }

    public static void RasterizerTotalPixelQuotaRejects()
    {
        // R21: a single page exceeds a 1-pixel aggregate quota and aborts the run.
        RequirePdfium();
        string input = WriteSinglePagePdf();
        string outputDirectory = NewTempDirectory();
        ToolResult result = RunTool(RasterizerAssemblyPath.Value, input, outputDirectory, "--max-total-pixels", "1");
        TestAssert.Equal(1, result.ExitCode);
        TestAssert.Contains("exceeds the maximum", result.Output + result.Error);
    }

    public static void RasterizerDefaultQuotaRendersTinyPage()
    {
        // R21: the default aggregate quota renders a tiny page and exits, proving
        // termination and per-run cleanup.
        RequirePdfium();
        string input = WriteSinglePagePdf();
        string outputDirectory = NewTempDirectory();
        ToolResult result = RunTool(RasterizerAssemblyPath.Value, input, outputDirectory);
        TestAssert.Equal(0, result.ExitCode);
        TestAssert.True(File.Exists(Path.Combine(outputDirectory, "page-001.png")), "Rasterizer must render page 1 on success.");
    }

    private static string WriteMultiFlatePdf()
    {
        using var pdf = new MemoryStream();
        void AppendAscii(string text)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            pdf.Write(bytes, 0, bytes.Length);
        }

        AppendAscii("%PDF-1.4\n");
        for (int i = 1; i <= 3; i++)
        {
            byte[] inflated = new byte[10_000];
            Array.Fill(inflated, (byte)((byte)'A' + (byte)i));
            byte[] deflated;
            using (var raw = new MemoryStream())
            {
                using (var deflate = new DeflateStream(raw, CompressionLevel.Fastest, leaveOpen: true))
                {
                    deflate.Write(inflated, 0, inflated.Length);
                }

                deflated = raw.ToArray();
            }

            AppendAscii($"{i} 0 obj\n<< /Length {deflated.Length} /Filter /FlateDecode >>\nstream\n");
            pdf.Write(deflated, 0, deflated.Length);
            AppendAscii("\nendstream\nendobj\n");
        }

        AppendAscii("trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pdf");
        File.WriteAllBytes(path, pdf.ToArray());
        return path;
    }

    private static void WritePngPair(string referenceDirectory, string candidateDirectory)
    {
        byte[] rgb = new byte[100 * 100 * 3];
        for (int i = 0; i < rgb.Length; i++)
        {
            rgb[i] = (byte)(i % 251);
        }

        byte[] png = TestFixtures.CreateRgbPng(100, 100, rgb);
        File.WriteAllBytes(Path.Combine(referenceDirectory, "page-001.png"), png);
        File.WriteAllBytes(Path.Combine(candidateDirectory, "page-001.png"), png);
    }

    private static string WriteSinglePagePdf()
    {
        string docx = TestFixtures.WriteTempPackage(".docx", new Dictionary<string, string>
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
                  <w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
                </w:document>
                """,
        });
        string pdf = Path.ChangeExtension(Path.GetTempFileName(), ".pdf");
        Lokad.OoxPdf.OoxPdfConverter.Convert(docx, pdf, new Lokad.OoxPdf.OoxPdfOptions { InputKind = Lokad.OoxPdf.OoxPdfInputKind.Docx });
        return pdf;
    }

    private static void RequirePdfium()
    {
        if (!OperatingSystem.IsWindows())
        {
            TestAssert.Skip("PDFium rasterizer runs on Windows.");
        }

        string repositoryRoot = FindRepositoryRoot();
        string dll = Path.Combine(repositoryRoot, "tools", "vendor", "pdfium", "win-x64", "bin", "pdfium.dll");
        if (!File.Exists(dll))
        {
            TestAssert.Skip("PDFium native binary is not provisioned.");
        }
    }

    private static string NewTempDirectory()
    {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    private static ToolResult RunTool(string assemblyPath, params string[] args)
    {
        var start = new ProcessStartInfo("dotnet")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        start.ArgumentList.Add(assemblyPath);
        foreach (string arg in args)
        {
            start.ArgumentList.Add(arg);
        }

        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Failed to start tool process.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return new ToolResult(process.ExitCode, output, error);
    }

    private static string BuildTool(string projectPath, string dllName)
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
        start.ArgumentList.Add(projectPath);
        start.ArgumentList.Add("--nologo");

        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Failed to build tool project.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"Tool build failed with exit code {process.ExitCode}.{Environment.NewLine}{output}{Environment.NewLine}{error}");
        }

        string projectDirectory = Path.GetDirectoryName(projectPath) ?? throw new InvalidOperationException("Expected a project directory.");
        return Path.Combine(repositoryRoot, projectDirectory, "bin", "Debug", "net10.0", dllName);
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

    private sealed record ToolResult(int ExitCode, string Output, string Error);
}
