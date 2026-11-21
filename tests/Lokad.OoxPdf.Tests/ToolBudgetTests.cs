using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Text.Json;

namespace Lokad.OoxPdf.Tests;

internal static class ToolBudgetTests
{
    private static readonly Lazy<string> InspectAssemblyPath = new(() => BuildTool("tools/Lokad.OoxPdf.PdfInspect/Lokad.OoxPdf.PdfInspect.csproj", "Lokad.OoxPdf.PdfInspect.dll"));
    private static readonly Lazy<string> VisualDiffAssemblyPath = new(() => BuildTool("tools/Lokad.OoxPdf.VisualDiff/Lokad.OoxPdf.VisualDiff.csproj", "Lokad.OoxPdf.VisualDiff.dll"));
    private static readonly Lazy<string> RasterizerAssemblyPath = new(() => BuildTool("tools/Lokad.OoxPdf.PdfiumRasterizer/Lokad.OoxPdf.PdfiumRasterizer.csproj", "Lokad.OoxPdf.PdfiumRasterizer.dll"));

    public static void InspectToUnicodeSequentialSurrogatePairsRetainDistinctMathLetters()
    {
        var map = ReadToUnicode("1 beginbfrange\n<073D> <0740> <D835DC4E>\nendbfrange");
        TestAssert.Equal("\U0001D44E", map[0x073D]); TestAssert.Equal("\U0001D44F", map[0x073E]);
        TestAssert.Equal("\U0001D450", map[0x073F]); TestAssert.Equal("\U0001D451", map[0x0740]);
        TestAssert.Equal("abcd", string.Concat(Enumerable.Range(0x073D, 4).Select(i => map[i])).Normalize(NormalizationForm.FormKC));
    }

    public static void InspectToUnicodeRangeStringsAndExplicitArraysKeepIndependentTargets()
    {
        var map = ReadToUnicode("2 beginbfrange\n<0001> <0003> <00660069>\n<0004> <0005> [<D835DC4E> <0061>]\nendbfrange");
        TestAssert.Equal("fi", map[1]); TestAssert.Equal("fj", map[2]); TestAssert.Equal("fk", map[3]);
        TestAssert.Equal("\U0001D44E", map[4]); TestAssert.Equal("a", map[5]);
    }

    public static void InspectToUnicodeBmpRangesAndDirectAstralMappingsRemainExact()
    {
        var map = ReadToUnicode("1 beginbfchar\n<0004> <D83DDE00>\nendbfchar\n1 beginbfrange\n<0001> <0003> <0041>\nendbfrange");
        TestAssert.Equal("A", map[1]); TestAssert.Equal("B", map[2]); TestAssert.Equal("C", map[3]);
        TestAssert.Equal("\U0001F600", map[4]);
    }

    private static IReadOnlyDictionary<int, string> ReadToUnicode(string cmap)
    {
        var assembly = System.Reflection.Assembly.LoadFrom(InspectAssemblyPath.Value);
        var type = assembly.GetType("PdfToUnicodeMap") ?? throw new InvalidOperationException("Missing inspector CMap decoder.");
        var parse = type.GetMethod("Parse") ?? throw new InvalidOperationException("Missing inspector CMap parser.");
        object map = parse.Invoke(null, [cmap]) ?? throw new InvalidOperationException("Missing decoded CMap.");
        var entries = type.GetProperty("Entries")?.GetValue(map);
        return entries as IReadOnlyDictionary<int, string> ?? throw new InvalidOperationException("Missing decoded CMap entries.");
    }

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

    public static void VisualDiffComparesKeyedTransparencyWithExplicitAlpha()
    {
        string reference = NewTempDirectory();
        string candidate = NewTempDirectory();
        string output = NewTempDirectory();
        var expected = new List<(double ChangedRatio, double MeanError)>();

        // Packed grayscale sample 1 expands to 85. Upper tRNS bits are unused.
        AddComparisons(ImagingTests.BuildStoredPng(2, 1, 2, 0, 0, [0, 0x60], transparency: [0xA5, 0xFD]),
            [85, 85, 85, 0, 170, 170, 170, 255]);
        AddComparisons(ImagingTests.BuildStoredPng(2, 1, 8, 0, 0, [0, 30, 220], transparency: [0, 30]),
            [30, 30, 30, 0, 220, 220, 220, 255]);
        // Every near match differs in just one channel and must stay opaque.
        AddComparisons(ImagingTests.BuildStoredPng(4, 1, 8, 2, 0,
                [0, 22, 31, 40, 23, 31, 40, 22, 32, 40, 22, 31, 41], transparency: [0xA5, 22, 0xA5, 31, 0xA5, 40]),
            [22, 31, 40, 0, 23, 31, 40, 255, 22, 32, 40, 255, 22, 31, 41, 255]);
        AddComparisons(ImagingTests.BuildStoredPng(2, 1, 1, 3, 0, [0, 0x40],
                palette: [22, 31, 40, 210, 220, 230], transparency: [64]),
            [22, 31, 40, 64, 210, 220, 230, 255]);

        ToolResult result = RunTool(VisualDiffAssemblyPath.Value, reference, candidate, output);
        TestAssert.Equal(0, result.ExitCode);
        using JsonDocument report = JsonDocument.Parse(File.ReadAllText(Path.Combine(output, "metrics.json")));
        TestAssert.Equal(expected.Count, report.RootElement.GetArrayLength());
        foreach ((JsonElement metric, (double changedRatio, double error)) in report.RootElement.EnumerateArray().Zip(expected))
        {
            TestAssert.Equal(changedRatio, metric.GetProperty("ChangedPixelRatioAtThreshold16").GetDouble());
            TestAssert.Equal(error, metric.GetProperty("MeanAbsoluteError").GetDouble());
        }

        void AddComparisons(byte[] keyed, byte[] rgba)
        {
            int width = rgba.Length / 4;
            string equivalent = $"page-{expected.Count + 1:D3}.png";
            File.WriteAllBytes(Path.Combine(reference, equivalent), keyed);
            File.WriteAllBytes(Path.Combine(candidate, equivalent), TestFixtures.CreateRgbaPng(width, 1, rgba));
            expected.Add((0, 0));

            string opaque = $"page-{expected.Count + 1:D3}.png";
            File.WriteAllBytes(Path.Combine(reference, opaque), keyed);
            expected.Add((1d / width, (255d - rgba[3]) / rgba.Length));
            rgba[3] = 255;
            File.WriteAllBytes(Path.Combine(candidate, opaque), TestFixtures.CreateRgbaPng(width, 1, rgba));
        }
    }

    public static void VisualDiffRejectsInvalidTransparencyKeyLengths()
    {
        byte[] opaque = TestFixtures.CreateRgbPng(1, 1, [22, 31, 40]);
        foreach (byte[] malformed in new[]
        {
            ImagingTests.BuildStoredPng(1, 1, 8, 0, 0, [0, 30], transparency: [30]),
            ImagingTests.BuildStoredPng(1, 1, 8, 2, 0, [0, 22, 31, 40], transparency: [0, 22, 0, 31, 0]),
        })
        {
            string reference = NewTempDirectory();
            string candidate = NewTempDirectory();
            string output = NewTempDirectory();
            File.WriteAllBytes(Path.Combine(reference, "page-001.png"), opaque);
            File.WriteAllBytes(Path.Combine(candidate, "page-001.png"), malformed);
            ToolResult result = RunTool(VisualDiffAssemblyPath.Value, reference, candidate, output);
            TestAssert.Equal(1, result.ExitCode);
            TestAssert.Contains("PNG transparency key has the wrong length", result.Error);
            TestAssert.True(!File.Exists(Path.Combine(output, "metrics.json")), "Invalid input must not produce comparison metrics.");
        }
    }

    public static void VisualDiffForegroundRecallCoversSparseBlankAndRegionCases()
    {
        byte[] white = Enumerable.Repeat((byte)255, 100 * 100 * 3).ToArray();
        byte[] ink = (byte[])white.Clone();
        for (int y = 10; y < 13; y++)
            for (int x = 10; x < 14; x++)
                Array.Clear(ink, (y * 100 + x) * 3, 3);
        string blankDirectory = NewTempDirectory();
        string inkDirectory = NewTempDirectory();
        File.WriteAllBytes(Path.Combine(blankDirectory, "page-001.png"), TestFixtures.CreateRgbPng(100, 100, white));
        File.WriteAllBytes(Path.Combine(inkDirectory, "page-001.png"), TestFixtures.CreateRgbPng(100, 100, ink));
        string regions = Path.Combine(NewTempDirectory(), "regions.json");
        File.WriteAllText(regions, """[{"page":1,"region":"full","xRatio":0,"yRatio":0,"widthRatio":1,"heightRatio":1},{"page":1,"region":"bottom-half","xRatio":0,"yRatio":0.5,"widthRatio":1,"heightRatio":0.5}]""");
        (string Reference, string Candidate, bool Regional, double? Recall, double Histogram)[] cases =
        [
            (inkDirectory, blankDirectory, false, 0, 0),
            (blankDirectory, inkDirectory, false, null, 0),
            (inkDirectory, inkDirectory, false, 1, 1),
            (blankDirectory, blankDirectory, false, null, 1),
            (inkDirectory, blankDirectory, true, 0, 0),
            (inkDirectory, inkDirectory, true, 1, 1)
        ];
        foreach (var item in cases)
        {
            string output = NewTempDirectory();
            string[] arguments = item.Regional ? [item.Reference, item.Candidate, output, regions] : [item.Reference, item.Candidate, output];
            ToolResult result = RunTool(VisualDiffAssemblyPath.Value, arguments);
            TestAssert.Equal(0, result.ExitCode);
            using JsonDocument page = JsonDocument.Parse(File.ReadAllText(Path.Combine(output, "metrics.json")));
            TestAssert.Equal(1, page.RootElement.GetArrayLength());
            CheckRecall(page.RootElement[0], item.Recall);
            TestAssert.Equal(item.Histogram, page.RootElement[0].GetProperty("ForegroundColorHistogramCorrelation").GetDouble());
            if (!item.Regional) continue;
            using JsonDocument regional = JsonDocument.Parse(File.ReadAllText(Path.Combine(output, "region-metrics.json")));
            JsonElement full = regional.RootElement.EnumerateArray().Single(row => row.GetProperty("Region").GetString() == "full");
            JsonElement bottom = regional.RootElement.EnumerateArray().Single(row => row.GetProperty("Region").GetString() == "bottom-half");
            CheckRecall(full, item.Recall);
            CheckRecall(bottom, null);
            TestAssert.Equal(1d, bottom.GetProperty("ForegroundColorHistogramCorrelation").GetDouble());
        }
        static void CheckRecall(JsonElement metric, double? expected)
        {
            JsonElement actual = metric.GetProperty("ForegroundRecall");
            if (expected is { } recall) TestAssert.Equal(recall, actual.GetDouble());
            else TestAssert.True(actual.ValueKind == JsonValueKind.Null, "Empty-reference recall must remain inapplicable.");
        }
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

    private static string NewTempDirectory()
    {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    internal static void RasterizePdfAt144Dpi(string input, string outputDirectory)
    {
        ToolResult result = RunTool(RasterizerAssemblyPath.Value, input, outputDirectory, "144");
        TestAssert.True(result.ExitCode == 0, $"PDFium validation failed. {result.Output}{result.Error}");
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
        Task<string> outputRead = process.StandardOutput.ReadToEndAsync();
        Task<string> errorRead = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(60_000))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            throw new TimeoutException("Child process timed out.");
        }
        string output = outputRead.GetAwaiter().GetResult();
        string error = errorRead.GetAwaiter().GetResult();
        return new ToolResult(process.ExitCode, output, error);
    }

    private static string BuildTool(string projectPath, string dllName)
    {
        string binding = dllName switch
        {
            "Lokad.OoxPdf.PdfInspect.dll" => "OOXPDF_TEST_PDF_INSPECTOR",
            "Lokad.OoxPdf.VisualDiff.dll" => "OOXPDF_TEST_VISUAL_DIFF",
            "Lokad.OoxPdf.PdfiumRasterizer.dll" => "OOXPDF_TEST_PDF_RASTERIZER",
            _ => throw new InvalidOperationException("Unknown validation tool.")
        };
        if (Environment.GetEnvironmentVariable(binding) is { Length: > 0 } frozen)
        {
            string path = Path.GetFullPath(frozen);
            return File.Exists(path) ? path : throw new InvalidOperationException("Frozen validation tool is missing.");
        }
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
        string framework = "net" + new System.Runtime.Versioning.FrameworkName(AppContext.TargetFrameworkName!).Version.ToString(2);
        start.ArgumentList.Add("--framework");
        start.ArgumentList.Add(framework);

        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Failed to build tool project.");
        Task<string> outputRead = process.StandardOutput.ReadToEndAsync();
        Task<string> errorRead = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(120_000))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            throw new TimeoutException("Child process timed out.");
        }
        string output = outputRead.GetAwaiter().GetResult();
        string error = errorRead.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"Tool build failed with exit code {process.ExitCode}.{Environment.NewLine}{output}{Environment.NewLine}{error}");
        }

        string projectDirectory = Path.GetDirectoryName(projectPath) ?? throw new InvalidOperationException("Expected a project directory.");
        return Path.Combine(repositoryRoot, projectDirectory, "bin", "Debug", framework, dllName);
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
