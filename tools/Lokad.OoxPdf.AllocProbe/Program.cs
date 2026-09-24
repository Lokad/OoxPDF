using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

// Whole-pipeline allocation probe (G01 slice 1): measures calling-thread
// allocated bytes, GC counts, wall time, output size/pages, and byte
// stability across cold + warm conversions. Stage split (parse/layout/
// emission/write) and corpus generation stay open for later slices.
if (args.Any(arg => string.Equals(arg, "--help", StringComparison.Ordinal) || string.Equals(arg, "-h", StringComparison.Ordinal)))
{
    Console.WriteLine("Usage: Lokad.OoxPdf.AllocProbe --out <report.json> [--warmup <n>] [--iterations <n>] [--stages] [--resident-window bytes] <input...> | --self-test | --font-breadth <n,...> --out <report.json>");
    Console.WriteLine("  Measures one cold plus N warm conversions per input and writes a JSON report.");
    Console.WriteLine("  --output-mode buffer (default) keeps converter output in a MemoryStream;");
    Console.WriteLine("  --output-mode file writes converter output to a temp file (no output buffering attributed).");
    Console.WriteLine("  --isolate measures each input in a fresh child process (independent cold, per-input process peaks).");
    Console.WriteLine("  Allocation scope is the calling thread; see report/allocationScope.");
    Console.WriteLine("  --concurrency <n> runs n conversions of each input in parallel with batch peaks.");
    return 2;
}

if (args.Any(arg => string.Equals(arg, "--self-test", StringComparison.Ordinal)))
{
    return RunSelfTest();
}

if (ReadOption("--font-breadth") is string breadthSpec)
{
    return RunFontBreadth(breadthSpec, ReadOption("--font-file"), ReadOption("--out"));
}

string? reportPath = ReadOption("--out");
int warmup = ReadOption("--warmup") is null ? 1 : ReadIntOption("--warmup") ?? -1;
int iterations = ReadOption("--iterations") is null ? 3 : ReadIntOption("--iterations") ?? -1;
bool measureStages = args.Any(arg => string.Equals(arg, "--stages", StringComparison.Ordinal));
bool isolate = args.Any(arg => string.Equals(arg, "--isolate", StringComparison.Ordinal))
    && !args.Any(arg => string.Equals(arg, "--isolated-child", StringComparison.Ordinal));
string outputMode = ReadOption("--output-mode") ?? "buffer";
long residentWindowBytes = ReadOption("--resident-window") is null ? 67108864L : ReadLongOption("--resident-window") ?? -1L;
if (residentWindowBytes < 0)
{
    Console.Error.WriteLine("Invalid --resident-window: expected non-negative bytes.");
    return 2;
}
if (outputMode != "buffer" && outputMode != "file")
{
    Console.Error.WriteLine($"Invalid --output-mode '{outputMode}': expected buffer or file.");
    return 2;
}

string[] inputs = args.Where(arg => !arg.StartsWith("--", StringComparison.Ordinal) && !IsOptionValue(arg)).ToArray();
if (reportPath is null || inputs.Length == 0 || warmup < 0 || iterations < 1)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.AllocProbe --out <report.json> [--warmup <n>] [--iterations <n>] [--stages] [--output-mode buffer|file] [--resident-window bytes] [--isolate] <input...>");
    return 2;
}

if (isolate)
{
    return RunIsolated(reportPath, inputs, warmup, iterations, measureStages, outputMode, residentWindowBytes);
}

if (ReadOption("--concurrency") is string concurrencySpec)
{
    if (!int.TryParse(concurrencySpec, out int concurrency) || concurrency < 1)
    {
        Console.Error.WriteLine("Invalid --concurrency: expected a positive integer.");
        return 2;
    }

    return RunConcurrency(reportPath, inputs, outputMode, residentWindowBytes, concurrency);
}

var reports = new List<object>();
foreach (string input in inputs)
{
    byte[] inputBytes;
    try
    {
        inputBytes = File.ReadAllBytes(input);
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Cannot read '{input}': {ex.Message}");
        return 1;
    }

    try
    {
        string? peakNote = args.Any(arg => string.Equals(arg, "--isolated-child", StringComparison.Ordinal))
            ? "single-input isolated child process: peak is per-input"
            : null;
        reports.Add(MeasureInput(Path.GetFileName(input), inputBytes, warmup, iterations, measureStages, outputMode, peakNote, residentWindowBytes));
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Conversion failed for '{input}': {ex.Message}");
        return 1;
    }
}

var report = BuildReport(reports, outputMode, isolation: "in-process (cold is process-first; later inputs reuse static caches; process peak is a lifetime high-water mark)");
Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(reportPath)) ?? ".");
File.WriteAllText(reportPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Wrote {reportPath} ({reports.Count} inputs).");
return 0;

static object BuildReport(List<object> reports, string outputMode, string isolation)
{
    return new
    {
        tool = "Lokad.OoxPdf.AllocProbe",
        createdAtUtc = DateTime.UtcNow.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
        framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
#if DEBUG
        buildConfiguration = "Debug",
#else
        buildConfiguration = "Release",
#endif
        processBits = Environment.Is64BitProcess ? 64 : 32,
        serverGc = System.Runtime.GCSettings.IsServerGC,
        gcLatencyMode = System.Runtime.GCSettings.LatencyMode.ToString(),
        os = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
        processorCount = Environment.ProcessorCount,
        sourceRevision = GetSourceRevision(),
        outputMode,
        isolation,
        allocationScope = "calling-thread (GC.GetAllocatedBytesForCurrentThread); hashing, page-count checks, and report serialization run outside the counters; I/O buffers and background loaders outside this thread are not attributed",
        fontInventory = DescribeFontInventory(),
        inputs = reports.ToArray(),
    };
}

static string GetSourceRevision()
{
    // SourceLink stamps the library informational version in CI builds; local builds report unknown.
    string? version = typeof(OoxPdfConverter).Assembly
        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
    return string.IsNullOrWhiteSpace(version) ? "unknown (local build)" : version;
}

static object DescribeFontInventory()
{
    // Counted after all measurements so inventory enumeration never warms the
    // static discovery cache ahead of a cold conversion in this process.
    try
    {
        int faces = new WindowsFontResolver().GetDiscoveredFonts().Count;
        return new { resolver = "default WindowsFontResolver", discoveredFaces = faces };
    }
    catch (Exception ex)
    {
        return new { resolver = "default WindowsFontResolver", discoveredFaces = -1, error = ex.GetType().Name };
    }
}

static int RunIsolated(string reportPath, string[] inputs, int warmup, int iterations, bool measureStages, string outputMode, long residentWindowBytes)
{
    // cold inputs run independently in fresh child processes so static
    // caches cannot leak across inputs and each child reports its own process peak.
    string? entry = Assembly.GetEntryAssembly()?.Location;
    bool useDotnet = entry?.EndsWith(".dll", StringComparison.OrdinalIgnoreCase) == true;
    bool useExe = entry?.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) == true;
    if (!useDotnet && !useExe)
    {
        Console.Error.WriteLine("Cannot determine probe binary for --isolate.");
        return 1;
    }

    var merged = new List<object>();
    foreach (string input in inputs)
    {
        string childReport = Path.Combine(Path.GetTempPath(), "allocprobe-" + Guid.NewGuid().ToString("N") + ".json");
        var psi = new ProcessStartInfo(useDotnet ? "dotnet" : entry!)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
        };
        if (useDotnet)
        {
            psi.ArgumentList.Add(entry!);
        }

        psi.ArgumentList.Add("--out");
        psi.ArgumentList.Add(childReport);
        psi.ArgumentList.Add("--warmup");
        psi.ArgumentList.Add(warmup.ToString(System.Globalization.CultureInfo.InvariantCulture));
        psi.ArgumentList.Add("--iterations");
        psi.ArgumentList.Add(iterations.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (measureStages)
        {
            psi.ArgumentList.Add("--stages");
        }

        psi.ArgumentList.Add("--output-mode");
        psi.ArgumentList.Add(outputMode);
        psi.ArgumentList.Add("--resident-window");
        psi.ArgumentList.Add(residentWindowBytes.ToString(System.Globalization.CultureInfo.InvariantCulture));
        psi.ArgumentList.Add("--isolated-child");
        psi.ArgumentList.Add(input);
        try
        {
            using Process child = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start probe child process.");
            string childOut = child.StandardOutput.ReadToEnd();
            string childErr = child.StandardError.ReadToEnd();
            child.WaitForExit();
            if (child.ExitCode != 0)
            {
                Console.Error.WriteLine($"Isolated child failed for '{input}' (exit {child.ExitCode}): {childErr}{childOut}");
                return 1;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Isolated child failed for '{input}': {ex.Message}");
            return 1;
        }

        try
        {
            using JsonDocument childDoc = JsonDocument.Parse(File.ReadAllText(childReport));
            if (!childDoc.RootElement.TryGetProperty("inputs", out JsonElement childInputs) ||
                childInputs.GetArrayLength() != 1)
            {
                Console.Error.WriteLine($"Isolated child report is malformed for '{input}'.");
                return 1;
            }

            foreach (JsonElement element in childInputs.EnumerateArray())
            {
                merged.Add(element.Clone());
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Cannot read isolated child report for '{input}': {ex.Message}");
            return 1;
        }
        finally
        {
            try { File.Delete(childReport); } catch (IOException) { }
        }
    }

    object report = BuildReport(merged, outputMode, isolation: "per-input child processes (independent cold conversions; per-input peakWorkingSetBytes is that child's process peak)");
    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(reportPath)) ?? ".");
    File.WriteAllText(reportPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine($"Wrote {reportPath} ({merged.Count} inputs, isolated).");
    return 0;
}

bool IsOptionValue(string arg)
{
    int index = Array.IndexOf(args, arg);
    return index > 0 && (args[index - 1] == "--out" || args[index - 1] == "--warmup" || args[index - 1] == "--iterations" || args[index - 1] == "--output-mode" || args[index - 1] == "--resident-window" || args[index - 1] == "--concurrency");
}

string? ReadOption(string name)
{
    int index = Array.IndexOf(args, name);
    return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
}

int? ReadIntOption(string name)
{
    string? text = ReadOption(name);
    if (text is null)
    {
        return null;
    }

    if (!int.TryParse(text, out int value))
    {
        Console.Error.WriteLine($"Invalid integer for {name}: {text}");
        return null;
    }

    return value;
}

long? ReadLongOption(string name)
{
    string? text = ReadOption(name);
    if (text is null)
    {
        return null;
    }

    if (!long.TryParse(text, out long value))
    {
        Console.Error.WriteLine($"Invalid integer for {name}: {text}");
        return null;
    }

    return value;
}

static object MeasureInput(string name, byte[] inputBytes, int warmup, int iterations, bool measureStages, string outputMode, string? peakNoteOverride = null, long residentWindowBytes = 67108864L)
{
    string kind = name.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) ? "pptx"
        : name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ? "docx" : "unknown";
    string inputSha = Convert.ToHexString(SHA256.HashData(inputBytes)).ToLowerInvariant();
    var outputShas = new List<string>();
    var outputLengths = new List<int>();
    var pageCounts = new List<int>();

    var options = new OoxPdfOptions
    {
        InputKind = kind == "docx" ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx,
        ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = residentWindowBytes },
    };
    (object cold, string coldSha, int coldLength, int coldPages) = MeasureOnce(inputBytes, options, inputExtension: kind == "docx" ? ".docx" : ".pptx", outputMode, recordOutput: true);
    outputShas.Add(coldSha);
    outputLengths.Add(coldLength);
    pageCounts.Add(coldPages);
    for (int i = 0; i < warmup; i++)
    {
        MeasureOnce(inputBytes, options, inputExtension: kind == "docx" ? ".docx" : ".pptx", outputMode, recordOutput: false);
    }

    var warm = new List<object>();
    for (int i = 0; i < iterations; i++)
    {
        (object measured, string sha, int length, int pages) = MeasureOnce(inputBytes, options, inputExtension: kind == "docx" ? ".docx" : ".pptx", outputMode, recordOutput: true);
        warm.Add(measured);
        outputShas.Add(sha);
        outputLengths.Add(length);
        pageCounts.Add(pages);
    }

    bool stable = outputShas.Distinct(StringComparer.Ordinal).Count() == 1;
    if (!stable)
    {
        Console.Error.WriteLine($"WARNING: '{name}' produced {outputShas.Distinct(StringComparer.Ordinal).Count()} distinct outputs; measurements may compare different work.");
    }

    return new
    {
        name,
        kind,
        inputBytes = inputBytes.Length,
        inputSha256 = inputSha,
        outputBytes = outputLengths[0],
        outputSha256 = outputShas[0],
        outputStable = stable,
        pageCount = pageCounts[0],
        fontResolver = "default",
        outputMode,
        residentWindowBytes,
        cold,
        warm = warm.ToArray(),
        stages = measureStages ? MeasureStages(kind, inputBytes, outputShas[0], outputMode) : null,
        peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64,
        peakWorkingSetNote = peakNoteOverride ?? "process lifetime high-water mark, including earlier inputs and stage probes; pass --isolate for per-input peaks",
    };
}

// stage metrics attribute only the stage lambda. Verification output
// (hashes, page counts) is produced by callers outside these counters, and stage
// prerequisites are rebuilt by callers beforehand. Per-stage retained memory is not
// reported: stage outputs stay alive for downstream stages, so no isolated retained
// delta exists here; see whole-conversion retainedDeltaBytes instead.
static (T Value, object Metrics) MeasureStep<T>(Func<T> produce, string scope)
{
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();
    long startAllocated = GC.GetAllocatedBytesForCurrentThread();
    int startGen0 = GC.CollectionCount(0);
    int startGen1 = GC.CollectionCount(1);
    int startGen2 = GC.CollectionCount(2);
    var watch = Stopwatch.StartNew();
    T value = produce();
    watch.Stop();
    object metrics = new
    {
        scope,
        allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - startAllocated,
        gen0Collections = GC.CollectionCount(0) - startGen0,
        gen1Collections = GC.CollectionCount(1) - startGen1,
        gen2Collections = GC.CollectionCount(2) - startGen2,
        elapsedMilliseconds = watch.Elapsed.TotalMilliseconds,
    };
    return (value, metrics);
}

// Stage attribution (G01 slice 2): each stage rebuilds its prerequisites
// unmeasured so only the stage itself is attributed. Stages run after the
// whole-pipeline measurements, so shared static caches are warm; the whole
// control stays the cold/warm reference. Prerequisite objects are left for
// GC like the converter itself (no disposal).
static object MeasureStages(string kind, byte[] inputBytes, string wholeOutputSha, string outputMode)
{
    bool isDocx = string.Equals(kind, "docx", StringComparison.OrdinalIgnoreCase);
    var markupMode = OoxPdfDocxMarkupMode.Final;
    var geometryMode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;

    OoxPackage BuildPackage() => OoxPackage.Open(new MemoryStream(inputBytes, writable: false), CancellationToken.None);
    (_, object openMetrics) = MeasureStep(() => BuildPackage(), "exclusive: package open only");

    // Read prerequisites are built outside the measured lambda so only the read
    // itself is attributed (Q05: previously BuildPackage ran inside the counter).
    OoxPackage readPackage = BuildPackage();
    object? sceneMetrics = null;
    object readMetrics;
    if (isDocx)
    {
        (_, readMetrics) = MeasureStep(() => new DocxReader().Read(readPackage, diagnosticSink: null, CancellationToken.None, markupMode), "exclusive: document read from prebuilt package");
    }
    else
    {
        (_, readMetrics) = MeasureStep(() => new PptxReader().Read(readPackage, CancellationToken.None), "exclusive: document read from prebuilt package");
        OoxPackage scenePackage = BuildPackage();
        PptxDocument sceneDocument = new PptxReader().Read(scenePackage, CancellationToken.None);
        (_, sceneMetrics) = MeasureStep(() => new PptxSceneBuilder().Build(sceneDocument, scenePackage, CancellationToken.None), "exclusive: scene build from prebuilt package and document (pptx only)");
    }

    System.Collections.Generic.IReadOnlyList<Lokad.OoxPdf.Pdf.PdfPage> pages;
    object renderMetrics;
    string renderScope;
    {
        OoxPackage renderPackage = BuildPackage();
        if (isDocx)
        {
            DocxDocument renderDocument = new DocxReader().Read(renderPackage, diagnosticSink: null, CancellationToken.None, markupMode);
            renderScope = "exclusive: layout and emission from prebuilt package and document";
            (pages, renderMetrics) = MeasureStep(() => new DocxRenderer(fontResolver: null, markupMode, geometryMode).RenderBlankPages(renderDocument, diagnosticSink: null, CancellationToken.None).ToList(), renderScope);
        }
        else
        {
            PptxDocument renderDocument = new PptxReader().Read(renderPackage, CancellationToken.None);
            renderScope = "INCLUSIVE: PptxRenderer.RenderPages rebuilds the scene internally; package and document prerequisites are prebuilt";
            (pages, renderMetrics) = MeasureStep(() => new PptxRenderer(fontResolver: null).RenderPages(renderDocument, renderPackage, diagnosticSink: null, CancellationToken.None).ToList(), renderScope);
        }
    }

    (byte[] stagedBytes, object writeMetrics) = MeasureStep(() =>
    {
        using var output = new MemoryStream();
        Lokad.OoxPdf.Pdf.PdfDocumentWriter.WriteBlank(output, pages, CancellationToken.None, creationDate: null);
        return output.ToArray();
    }, "exclusive: PDF serialization of prebuilt pages into a memory buffer");
    string stagedSha = Convert.ToHexString(SHA256.HashData(stagedBytes)).ToLowerInvariant();

    return new
    {
        stageSet = isDocx ? "open/read/render/write" : "open/read/scene/render/write",
        stageOrderNote = "Stages run after whole-pipeline cold+warm; each stage rebuilds prerequisites unmeasured. Static caches are warm. PPTX render is inclusive of a scene rebuild; subtract scene.allocatedBytes for an approximate exclusive render figure.",
        open = openMetrics,
        read = readMetrics,
        scene = sceneMetrics,
        render = renderMetrics,
        write = writeMetrics,
        stagedOutputBytes = stagedBytes.Length,
        stagedMatchesWhole = string.Equals(stagedSha, wholeOutputSha, StringComparison.Ordinal),
    };
}

// the timer and allocation counters stop before output verification
// (hashing, page-count decoding) so verification work is never attributed to the
// conversion. retainedDeltaBytes is the post-full-GC heap delta against a pre-run
// baseline with the output released, i.e. surviving caches rather than live output.
static (object Metrics, string Sha, int Length, int Pages) MeasureOnce(byte[] inputBytes, OoxPdfOptions options, string inputExtension, string outputMode, bool recordOutput)
{
    string? stageDirectory = null;
    string? fileInput = null;
    string? fileOutput = null;
    if (outputMode == "file")
    {
        stageDirectory = Path.Combine(Path.GetTempPath(), "allocprobe-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(stageDirectory);
        fileInput = Path.Combine(stageDirectory, "input" + inputExtension);
        fileOutput = Path.Combine(stageDirectory, "output.pdf");
        File.WriteAllBytes(fileInput, inputBytes);
    }

    try
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long baselineRetained = GC.GetTotalMemory(forceFullCollection: false);
        long startAllocated = GC.GetAllocatedBytesForCurrentThread();
        int startGen0 = GC.CollectionCount(0);
        int startGen1 = GC.CollectionCount(1);
        int startGen2 = GC.CollectionCount(2);
        var watch = Stopwatch.StartNew();
        using var peakSampler = new PeakSampler();
        peakSampler.Start();
        byte[]? bufferedOutput = null;
        try
        {
            bufferedOutput = outputMode == "file"
                ? null
                : ConvertToBuffer(inputBytes, options);
            if (outputMode == "file")
            {
                ConvertToFile(fileInput!, fileOutput!, options);
            }
        }
        finally
        {
            peakSampler.Stop();
        }

        watch.Stop();
        long allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - startAllocated;
        int gen0 = GC.CollectionCount(0) - startGen0;
        int gen1 = GC.CollectionCount(1) - startGen1;
        int gen2 = GC.CollectionCount(2) - startGen2;

        // Verification runs outside the counters: hashing and page-count decoding
        // previously inflated the attributed conversion allocation.
        byte[] outputBytes = bufferedOutput ?? File.ReadAllBytes(fileOutput!);
        string sha = recordOutput ? Convert.ToHexString(SHA256.HashData(outputBytes)).ToLowerInvariant() : string.Empty;
        int pages = recordOutput ? ReadPageCount(outputBytes) : 0;
        int length = outputBytes.Length;
        bufferedOutput = null;
        outputBytes = null!;
        long retainedDeltaBytes = RetainedAfterFullGc() - baselineRetained;

        object metrics = new
        {
            outputMode,
            allocatedBytes,
            gen0Collections = gen0,
            gen1Collections = gen1,
            gen2Collections = gen2,
            elapsedMilliseconds = watch.Elapsed.TotalMilliseconds,
            retainedDeltaBytes,
            peakManagedHeapBytes = peakSampler.PeakManagedHeapBytes,
            peakPrivateBytes = peakSampler.PeakPrivateBytes,
            peakWorkingSetBytes = peakSampler.PeakWorkingSetBytes,
        };
        return (metrics, sha, length, pages);
    }
    finally
    {
        if (stageDirectory is not null)
        {
            try { Directory.Delete(stageDirectory, recursive: true); } catch (IOException) { }
        }
    }
}

// NoInlining scopes the conversion's short-lived buffers so the post-run full GC
// can actually release them and retainedDeltaBytes reflects surviving ownership.
[System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
static byte[] ConvertToBuffer(byte[] inputBytes, OoxPdfOptions options)
{
    using var input = new MemoryStream(inputBytes, writable: false);
    using var output = new MemoryStream();
    OoxPdfConverter.Convert(input, output, options);
    return output.ToArray();
}

[System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
static void ConvertToFile(string inputPath, string outputPath, OoxPdfOptions options)
{
    OoxPdfConverter.Convert(inputPath, outputPath, options);
}

static int ReadPageCount(byte[] pdfBytes)
{
    // Anchored to the writer's deterministic Pages object ("<< /Type /Pages /Count N ...").
    Match match = Regex.Match(Encoding.Latin1.GetString(pdfBytes), @"/Type\s*/Pages\s*/Count\s+(\d+)");
    return match.Success && int.TryParse(match.Groups[1].Value, out int count) ? count : -1;
}

static int RunSelfTest()
{
    var cases = new (string Name, byte[] Package)[]
    {
        ("self-test.pptx", BuildMinimalPptx()),
        ("self-test.docx", BuildMinimalDocx()),
    };
    bool ok = true;
    foreach ((string name, byte[] package) in cases)
    {
        try
        {
            dynamic report = MeasureInput(name, package, warmup: 0, iterations: 1, measureStages: false, outputMode: "buffer");
            bool stable = report.outputStable;
            int pages = report.pageCount;
            int bytes = report.outputBytes;
            Console.WriteLine($"PASS self-test {name}: bytes={bytes} pages={pages} stable={stable}");
            if (!stable || pages < 1 || bytes == 0)
            {
                ok = false;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FAIL self-test {name}: {ex.GetType().Name}: {ex.Message}");
            ok = false;
        }
    }

    // a deliberately allocating fake stage validates attribution: phase
    // counters must capture the planted allocation (no under-count from misplaced
    // boundaries) without wild over-count.
    try
    {
        const int planted = 8 * 1024 * 1024;
        (_, object first) = MeasureStep(() => PlantAllocation(planted), "self-test fake stage");
        (_, object second) = MeasureStep(() => PlantAllocation(planted), "self-test fake stage");
        long firstBytes = ((dynamic)first).allocatedBytes;
        long secondBytes = ((dynamic)second).allocatedBytes;
        Console.WriteLine($"self-test attribution: first={firstBytes} second={secondBytes} planted={planted}");
        if (firstBytes < planted || secondBytes < planted || firstBytes > planted + 16L * 1024 * 1024 || Math.Abs(firstBytes - secondBytes) > 4L * 1024 * 1024)
        {
            Console.WriteLine("FAIL self-test attribution: planted allocation not attributed within tolerance.");
            ok = false;
        }
        else
        {
            Console.WriteLine("PASS self-test attribution.");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"FAIL self-test attribution: {ex.GetType().Name}: {ex.Message}");
        ok = false;
    }

    // RV22: peak sampler must report positive, ordered high-water marks for a
    // real conversion (working set and private bytes both contain the heap).
    try
    {
        dynamic peakReport = MeasureInput("self-test-peaks.docx", BuildMinimalDocx(), warmup: 0, iterations: 1, measureStages: false, outputMode: "buffer");
        dynamic peakWarm = peakReport.warm[0];
        long peakHeap = peakWarm.peakManagedHeapBytes;
        long peakPrivate = peakWarm.peakPrivateBytes;
        long peakWorkingSet = peakWarm.peakWorkingSetBytes;
        Console.WriteLine("self-test peaks: heap=" + peakHeap + " private=" + peakPrivate + " workingSet=" + peakWorkingSet);
        if (peakHeap <= 0 || peakPrivate <= 0 || peakWorkingSet <= 0 || peakWorkingSet < peakHeap || peakPrivate < peakHeap)
        {
            Console.WriteLine("FAIL self-test peaks: high-water marks must be positive with working set and private bytes above the heap.");
            ok = false;
        }
        else
        {
            Console.WriteLine("PASS self-test peaks.");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine("FAIL self-test peaks: " + ex.GetType().Name + ": " + ex.Message);
        ok = false;
    }

    // RV22: concurrent conversions of one document must agree byte-for-byte
    // with positive batch peaks.
    try
    {
        dynamic concurrent = MeasureConcurrencyInput("self-test-concurrency.docx", BuildMinimalDocx(), "buffer", 67108864L, 2);
        bool concurrentStable = concurrent.outputsStable;
        int concurrentPages = concurrent.pageCount;
        long batchHeap = concurrent.batchPeakManagedHeapBytes;
        long batchPrivate = concurrent.batchPeakPrivateBytes;
        long batchWorkingSet = concurrent.batchPeakWorkingSetBytes;
        Console.WriteLine("self-test concurrency: stable=" + concurrentStable + " pages=" + concurrentPages + " heap=" + batchHeap + " private=" + batchPrivate + " workingSet=" + batchWorkingSet);
        if (!concurrentStable || concurrentPages < 1 || batchHeap <= 0 || batchPrivate <= 0 || batchWorkingSet <= 0)
        {
            Console.WriteLine("FAIL self-test concurrency: parallel conversions must agree with positive batch peaks.");
            ok = false;
        }
        else
        {
            Console.WriteLine("PASS self-test concurrency.");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine("FAIL self-test concurrency: " + ex.GetType().Name + ": " + ex.Message);
        ok = false;
    }

    return ok ? 0 : 1;
}

[System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
static int PlantAllocation(int bytes)
{
    var planted = new byte[bytes];
    planted[bytes - 1] = 1;
    return planted[bytes - 1];
}

static byte[] BuildMinimalPptx()
{
    return BuildZip(new Dictionary<string, string>
    {
        ["[Content_Types].xml"] = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
              <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
            </Types>
            """,
        ["_rels/.rels"] = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
            </Relationships>
            """,
        ["ppt/_rels/presentation.xml.rels"] = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
            </Relationships>
            """,
        ["ppt/presentation.xml"] = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <p:sldSz cx="9144000" cy="6858000"/>
              <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
            </p:presentation>
            """,
        ["ppt/slides/slide1.xml"] = """
            <?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <p:cSld><p:spTree><p:sp>
                <p:nvSpPr><p:cNvPr id="2" name="Probe"/><p:nvPr/></p:nvSpPr>
                <p:spPr>
                  <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="rect"/>
                  <a:solidFill><a:srgbClr val="808080"/></a:solidFill>
                </p:spPr>
                <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>Probe</a:t></a:r></a:p></p:txBody>
              </p:sp></p:spTree></p:cSld>
            </p:sld>
            """,
    });
}

static byte[] BuildMinimalDocx()
{
    return BuildZip(new Dictionary<string, string>
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
                <w:p><w:r><w:t>Probe</w:t></w:r></w:p>
                <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
              </w:body>
            </w:document>
            """,
    });
}

static byte[] BuildZip(IReadOnlyDictionary<string, string> entries)
{
    using var stream = new MemoryStream();
    using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
    {
        foreach ((string name, string content) in entries)
        {
            ZipArchiveEntry entry = archive.CreateEntry(name);
            using Stream entryStream = entry.Open();
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            entryStream.Write(bytes);
        }
    }

    return stream.ToArray();
}

// Font-breadth probe (F04 slice 6): N distinct file-backed fonts force N
// full font loads per conversion. Reports calling-thread allocated bytes
// and post-GC retained bytes per breadth, plus retained bytes across
// repeat conversions sharing one resolver (retention must stay flat).
static int RunFontBreadth(string spec, string? fontFile, string? reportPath)
{
    if (reportPath is null)
    {
        Console.Error.WriteLine("Usage: Lokad.OoxPdf.AllocProbe --font-breadth <n,...> [--font-file <ttf>] --out <report.json>");
        return 2;
    }
    int[] breadths;
    try
    {
        breadths = spec.Split(',').Select(part => int.Parse(part.Trim(), System.Globalization.CultureInfo.InvariantCulture)).Distinct().OrderBy(n => n).ToArray();
    }
    catch (Exception ex) when (ex is FormatException or OverflowException)
    {
        Console.Error.WriteLine($"Cannot parse --font-breadth '{spec}': expected comma-separated integers.");
        return 2;
    }
    if (breadths.Length == 0 || breadths.Any(n => n < 1))
    {
        Console.Error.WriteLine($"Cannot parse --font-breadth '{spec}': expected positive integers.");
        return 2;
    }
    fontFile ??= Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "arial.ttf");
    if (!File.Exists(fontFile))
    {
        Console.Error.WriteLine($"Font source is missing: '{fontFile}'. Pass --font-file <embeddable-truetype>.");
        return 1;
    }
    byte[] fontBytes = File.ReadAllBytes(fontFile);
    string stageDir = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(reportPath)) ?? ".", "breadth-fonts");
    Directory.CreateDirectory(stageDir);
    var rows = new List<object>();
    foreach (int breadth in breadths)
    {
        (long allocated, long retained, int outputBytes, string sha, int pages) = MeasureBreadthRow(breadth, fontBytes, stageDir);
        rows.Add(new { breadth, allocatedBytes = allocated, retainedBytes = retained, outputBytes, outputSha256 = sha, pageCount = pages });
        Console.WriteLine($"breadth {breadth}: allocated={allocated} retained={retained} output={outputBytes} pages={pages}");
    }
    for (int repeat = 0; repeat < 2; repeat++)
    {
        (long reAllocated, long reRetained, int reOutput, string reSha, int rePages) = MeasureBreadthRow(breadths.Max(), fontBytes, stageDir);
        rows.Add(new { breadth = breadths.Max(), allocatedBytes = reAllocated, retainedBytes = reRetained, outputBytes = reOutput, outputSha256 = reSha, pageCount = rePages });
        Console.WriteLine($"fresh-resolver repeat {repeat}: allocated={reAllocated} retained={reRetained}");
    }
    object repeats = MeasureBreadthRepeats(breadths.Max(), fontBytes, stageDir);
    rows.Add(repeats);
    var report = new
    {
        tool = "Lokad.OoxPdf.AllocProbe",
        mode = "font-breadth",
        createdAtUtc = DateTime.UtcNow.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
        framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
        serverGc = System.Runtime.GCSettings.IsServerGC,
        allocationScope = "calling-thread (GC.GetAllocatedBytesForCurrentThread)",
        fontFile,
        fontBytes = fontBytes.Length,
        rows = rows.ToArray(),
    };
    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(reportPath)) ?? ".");
    File.WriteAllText(reportPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine($"Wrote {reportPath} ({rows.Count} rows).");
    return 0;
}

// NoInlining lets the per-row resolver die before the retained read, so
// residue measures what conversions leave behind, not live sources.
[System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
static (long Allocated, long Retained, int OutputBytes, string Sha, int Pages) MeasureBreadthRow(int breadth, byte[] fontBytes, string stageDir)
{
    string[] copies = WriteBreadthCopies(breadth, fontBytes, stageDir);
    byte[] docx = BuildBreadthDocx(breadth);
    var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = new BreadthResolver(copies) };
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();
    long startAllocated = GC.GetAllocatedBytesForCurrentThread();
    byte[] output = ConvertBreadthDocx(docx, options);
    long allocated = GC.GetAllocatedBytesForCurrentThread() - startAllocated;
    options = null!;
    long retained = RetainedAfterFullGc();
    return (allocated, retained, output.Length, Convert.ToHexString(SHA256.HashData(output)).ToLowerInvariant(), ReadPageCount(output));
}

static object MeasureBreadthRepeats(int breadth, byte[] fontBytes, string stageDir)
{
    string[] copies = WriteBreadthCopies(breadth, fontBytes, stageDir);
    byte[] docx = BuildBreadthDocx(breadth);
    var resolver = new BreadthResolver(copies);
    var allocated = new List<long>();
    var retained = new List<long>();
    for (int i = 0; i < 3; i++)
    {
        var options = new OoxPdfOptions { InputKind = OoxPdfInputKind.Docx, FontResolver = resolver };
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long startAllocated = GC.GetAllocatedBytesForCurrentThread();
        ConvertBreadthDocx(docx, options);
        allocated.Add(GC.GetAllocatedBytesForCurrentThread() - startAllocated);
        retained.Add(RetainedAfterFullGc());
    }
    return new { mode = "shared-resolver-repeats", breadth, conversions = 3, allocatedBytes = allocated.ToArray(), retainedBytes = retained.ToArray() };
}

static long RetainedAfterFullGc()
{
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();
    return GC.GetTotalMemory(forceFullCollection: false);
}

static byte[] ConvertBreadthDocx(byte[] docx, OoxPdfOptions options)
{
    using var input = new MemoryStream(docx, writable: false);
    using var output = new MemoryStream();
    OoxPdfConverter.Convert(input, output, options);
    return output.ToArray();
}

static string[] WriteBreadthCopies(int breadth, byte[] fontBytes, string stageDir)
{
    var copies = new string[breadth];
    for (int i = 0; i < breadth; i++)
    {
        string path = Path.Combine(stageDir, $"breadth-{breadth}-{i}.ttf");
        File.WriteAllBytes(path, fontBytes);
        copies[i] = path;
    }
    return copies;
}

static byte[] BuildBreadthDocx(int breadth)
{
    var body = new StringBuilder();
    body.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>");
    for (int i = 0; i < breadth; i++)
    {
        body.Append("<w:p><w:r><w:rPr><w:rFonts w:ascii=\"Breadth" + i + "\" w:hAnsi=\"Breadth" + i + "\"/></w:rPr><w:t xml:space=\"preserve\">Breadth paragraph " + i + " 0123456789</w:t></w:r></w:p>");
    }
    body.Append("<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
    return BuildZip(new Dictionary<string, string>
    {
        ["[Content_Types].xml"] = """
            <?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
            """,
        ["_rels/.rels"] = """
            <?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
            """,
        ["word/document.xml"] = body.ToString(),
    });
}

// RV22: bounded-concurrency measurement. One sequential warmup validates the
// input and warms caches, then N conversions run concurrently on thread-pool
// threads with independent options and scopes; a batch sampler records process
// peaks over the whole batch while per-task walls expose overlap efficiency.
static int RunConcurrency(string reportPath, string[] inputs, string outputMode, long residentWindowBytes, int concurrency)
{
    var reports = new List<object>();
    foreach (string input in inputs)
    {
        byte[] inputBytes;
        try
        {
            inputBytes = File.ReadAllBytes(input);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Cannot read input: " + ex.Message);
            return 1;
        }

        try
        {
            reports.Add(MeasureConcurrencyInput(Path.GetFileName(input), inputBytes, outputMode, residentWindowBytes, concurrency));
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Concurrent conversion failed for " + input + ": " + ex.Message);
            return 1;
        }
    }

    object report = BuildReport(reports, outputMode, isolation: "concurrency " + concurrency.ToString(System.Globalization.CultureInfo.InvariantCulture) + " parallel in-process conversions sharing static caches; batch peaks cover the whole batch");
    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(reportPath)) ?? ".");
    File.WriteAllText(reportPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine("Wrote " + reportPath + " (" + reports.Count + " inputs, concurrency " + concurrency.ToString(System.Globalization.CultureInfo.InvariantCulture) + ").");
    return 0;
}
static object MeasureConcurrencyInput(string name, byte[] inputBytes, string outputMode, long residentWindowBytes, int concurrency)
{
    string kind = name.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) ? "pptx"
        : name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ? "docx" : "unknown";
    string extension = kind == "docx" ? ".docx" : ".pptx";
    OoxPdfOptions FreshOptions()
    {
        return new OoxPdfOptions
        {
            InputKind = kind == "docx" ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx,
            ConversionLimits = new OoxConversionLimits { MaxResidentPageContentBytesPerConversion = residentWindowBytes },
        };
    }

    ConvertOnce(inputBytes, FreshOptions(), extension, outputMode);

    // Pre-batch collection mirrors the MeasureOnce precondition so floating
    // warmup garbage does not inflate batch peaks.
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();

    using var batchPeaks = new PeakSampler();
    batchPeaks.Start();
    var batchWatch = Stopwatch.StartNew();
    Task<(string Sha, int Length, int Pages, double Milliseconds)>[] tasks = Enumerable.Range(0, concurrency).Select(_ => Task.Run(() =>
    {
        var taskWatch = Stopwatch.StartNew();
        (string sha, int length, int pages) = ConvertOnce(inputBytes, FreshOptions(), extension, outputMode);
        taskWatch.Stop();
        return (sha, length, pages, taskWatch.Elapsed.TotalMilliseconds);
    })).ToArray();

    try
    {
        Task.WaitAll(tasks);
    }
    finally
    {
        batchPeaks.Stop();
    }

    batchWatch.Stop();
    var results = tasks.Select(task => task.Result).ToArray();
    bool stable = results.Select(result => result.Sha).Distinct(StringComparer.Ordinal).Count() == 1;
    if (!stable)
    {
        Console.Error.WriteLine("WARNING: concurrent conversions of " + name + " produced distinct outputs.");
    }

    return new
    {
        name,
        kind,
        inputBytes = inputBytes.Length,
        concurrency,
        outputsStable = stable,
        outputBytes = results[0].Length,
        outputSha256 = results[0].Sha,
        pageCount = results[0].Pages,
        elapsedMilliseconds = batchWatch.Elapsed.TotalMilliseconds,
        batchPeakManagedHeapBytes = batchPeaks.PeakManagedHeapBytes,
        batchPeakPrivateBytes = batchPeaks.PeakPrivateBytes,
        batchPeakWorkingSetBytes = batchPeaks.PeakWorkingSetBytes,
        tasks = results.Select(result => new { outputBytes = result.Length, pageCount = result.Pages, elapsedMilliseconds = result.Milliseconds }).ToArray(),
    };
}
static (string Sha, int Length, int Pages) ConvertOnce(byte[] inputBytes, OoxPdfOptions options, string inputExtension, string outputMode)
{
    if (outputMode == "file")
    {
        string stageDirectory = Path.Combine(Path.GetTempPath(), "allocprobe-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(stageDirectory);
        try
        {
            string fileInput = Path.Combine(stageDirectory, "input" + inputExtension);
            string fileOutput = Path.Combine(stageDirectory, "output.pdf");
            File.WriteAllBytes(fileInput, inputBytes);
            ConvertToFile(fileInput, fileOutput, options);
            byte[] outputBytes = File.ReadAllBytes(fileOutput);
            return (Convert.ToHexString(SHA256.HashData(outputBytes)).ToLowerInvariant(), outputBytes.Length, ReadPageCount(outputBytes));
        }
        finally
        {
            try { Directory.Delete(stageDirectory, recursive: true); } catch (IOException) { }
        }
    }

    byte[] buffered = ConvertToBuffer(inputBytes, options);
    return (Convert.ToHexString(SHA256.HashData(buffered)).ToLowerInvariant(), buffered.Length, ReadPageCount(buffered));
}

// RV22: process-peak sampling for single conversions. Calling-thread volume
// counters cannot observe live peaks, so a background thread records the
// managed-heap, private-byte, and working-set high-water marks seen while one
// conversion runs. Samples are approximate (5 ms cadence); the sampler thread
// allocates outside the calling-thread attribution.
sealed class PeakSampler : IDisposable
{
    private readonly Process samplerProcess = Process.GetCurrentProcess();
    private readonly Thread samplerThread;
    private readonly object peakLock = new object();
    private long peakManagedHeapBytes;
    private long peakPrivateBytes;
    private long peakWorkingSetBytes;
    private volatile bool running;

    public PeakSampler()
    {
        samplerThread = new Thread(SampleLoop)
        {
            IsBackground = true,
            Name = "AllocProbe peak sampler",
        };
    }

    public long PeakManagedHeapBytes => ReadPeak(ref peakManagedHeapBytes);

    public long PeakPrivateBytes => ReadPeak(ref peakPrivateBytes);

    public long PeakWorkingSetBytes => ReadPeak(ref peakWorkingSetBytes);

    public void Start()
    {
        running = true;
        samplerThread.Start();
    }

    public void Stop()
    {
        running = false;
        samplerThread.Join();
    }

    public void Dispose()
    {
        running = false;
        if (samplerThread.IsAlive)
        {
            samplerThread.Join();
        }

        samplerProcess.Dispose();
        GC.SuppressFinalize(this);
    }

    private long ReadPeak(ref long field)
    {
        lock (peakLock)
        {
            return field;
        }
    }

    private void RecordPeak(ref long field, long value)
    {
        lock (peakLock)
        {
            if (value > field)
            {
                field = value;
            }
        }
    }

    private void SampleLoop()
    {
        while (running)
        {
            RecordPeak(ref peakManagedHeapBytes, GC.GetTotalMemory(forceFullCollection: false));
            samplerProcess.Refresh();
            RecordPeak(ref peakPrivateBytes, samplerProcess.PrivateMemorySize64);
            RecordPeak(ref peakWorkingSetBytes, samplerProcess.WorkingSet64);
            Thread.Sleep(5);
        }
    }
}

// One file-backed program source per breadth family: distinct StableIds
// force distinct font loads, so breadth N converts through N parsed fonts.
sealed class BreadthResolver : IFontResolver
{
    private readonly IFontProgramSource[] sources;
    public BreadthResolver(IReadOnlyList<string> paths)
    {
        sources = paths.Select(path => (IFontProgramSource)new FileFontProgramSource(path)).ToArray();
    }
    public FontFaceResolution Resolve(FontRequest request)
    {
        for (int i = 0; i < sources.Length; i++)
        {
            if (request.FamilyName.Equals("Breadth" + i, StringComparison.OrdinalIgnoreCase))
            {
                return new FontFaceResolution(request.FamilyName, "Breadth" + i, new FontStyleKey(Bold: false, Italic: false, WeightClass: 400, FaceIndex: 0, HasMathTable: false), sources[i], IsFallback: false);
            }
        }
        return new FontFaceResolution(request.FamilyName, "Breadth0", new FontStyleKey(Bold: false, Italic: false, WeightClass: 400, FaceIndex: 0, HasMathTable: false), sources[0], IsFallback: true);
    }
}
