using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

// Whole-pipeline allocation probe (G01 slice 1): measures calling-thread
// allocated bytes, GC counts, wall time, output size/pages, and byte
// stability across cold + warm conversions. Stage split (parse/layout/
// emission/write) and corpus generation stay open for later slices.
if (args.Any(arg => string.Equals(arg, "--help", StringComparison.Ordinal) || string.Equals(arg, "-h", StringComparison.Ordinal)))
{
    Console.WriteLine("Usage: Lokad.OoxPdf.AllocProbe --out <report.json> [--warmup <n>] [--iterations <n>] [--stages] <input...> | --self-test");
    Console.WriteLine("  Measures one cold plus N warm conversions per input and writes a JSON report.");
    Console.WriteLine("  Allocation scope is the calling thread; see report/allocationScope.");
    return 2;
}

if (args.Any(arg => string.Equals(arg, "--self-test", StringComparison.Ordinal)))
{
    return RunSelfTest();
}

string? reportPath = ReadOption("--out");
int warmup = ReadOption("--warmup") is null ? 1 : ReadIntOption("--warmup") ?? -1;
int iterations = ReadOption("--iterations") is null ? 3 : ReadIntOption("--iterations") ?? -1;
bool measureStages = args.Any(arg => string.Equals(arg, "--stages", StringComparison.Ordinal));
string[] inputs = args.Where(arg => !arg.StartsWith("--", StringComparison.Ordinal) && !IsOptionValue(arg)).ToArray();
if (reportPath is null || inputs.Length == 0 || warmup < 0 || iterations < 1)
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.AllocProbe --out <report.json> [--warmup <n>] [--iterations <n>] [--stages] <input...>");
    return 2;
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
        reports.Add(MeasureInput(Path.GetFileName(input), inputBytes, warmup, iterations, measureStages));
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Conversion failed for '{input}': {ex.Message}");
        return 1;
    }
}

var report = new
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
    allocationScope = "calling-thread (GC.GetAllocatedBytesForCurrentThread); I/O buffers and background loaders outside this thread are not attributed",
    inputs = reports.ToArray(),
};
Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(reportPath)) ?? ".");
File.WriteAllText(reportPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Wrote {reportPath} ({reports.Count} inputs).");
return 0;

bool IsOptionValue(string arg)
{
    int index = Array.IndexOf(args, arg);
    return index > 0 && (args[index - 1] == "--out" || args[index - 1] == "--warmup" || args[index - 1] == "--iterations");
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

static object MeasureInput(string name, byte[] inputBytes, int warmup, int iterations, bool measureStages)
{
    string kind = name.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) ? "pptx"
        : name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ? "docx" : "unknown";
    string inputSha = Convert.ToHexString(SHA256.HashData(inputBytes)).ToLowerInvariant();
    var outputShas = new List<string>();
    var outputLengths = new List<int>();
    var pageCounts = new List<int>();

    var options = new OoxPdfOptions { InputKind = kind == "docx" ? OoxPdfInputKind.Docx : OoxPdfInputKind.Pptx };
    object cold = MeasureOnce(inputBytes, options, recordOutput: true, outputShas, outputLengths, pageCounts);
    for (int i = 0; i < warmup; i++)
    {
        MeasureOnce(inputBytes, options, recordOutput: false, outputShas, outputLengths, pageCounts);
    }

    var warm = new List<object>();
    for (int i = 0; i < iterations; i++)
    {
        warm.Add(MeasureOnce(inputBytes, options, recordOutput: true, outputShas, outputLengths, pageCounts));
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
        cold,
        warm = warm.ToArray(),
        stages = measureStages ? MeasureStages(kind, inputBytes, outputShas[0]) : null,
        peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64,
    };
}

static (T Value, object Metrics) MeasureStep<T>(Func<T> produce)
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
    long retainedBytes = GC.GetTotalMemory(forceFullCollection: false);
    object metrics = new
    {
        allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - startAllocated,
        gen0Collections = GC.CollectionCount(0) - startGen0,
        gen1Collections = GC.CollectionCount(1) - startGen1,
        gen2Collections = GC.CollectionCount(2) - startGen2,
        elapsedMilliseconds = watch.Elapsed.TotalMilliseconds,
        retainedBytes,
    };
    return (value, metrics);
}

// Stage attribution (G01 slice 2): each stage rebuilds its prerequisites
// unmeasured so only the stage itself is attributed. Stages run after the
// whole-pipeline measurements, so shared static caches are warm; the whole
// control stays the cold/warm reference. Prerequisite objects are left for
// GC like the converter itself (no disposal).
static object MeasureStages(string kind, byte[] inputBytes, string wholeOutputSha)
{
    bool isDocx = string.Equals(kind, "docx", StringComparison.OrdinalIgnoreCase);
    var markupMode = OoxPdfDocxMarkupMode.Final;
    var geometryMode = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;

    OoxPackage BuildPackage() => OoxPackage.Open(new MemoryStream(inputBytes, writable: false), CancellationToken.None);
    (_, object openMetrics) = MeasureStep(() => BuildPackage());

    object? sceneMetrics = null;
    object readMetrics;
    if (isDocx)
    {
        (_, readMetrics) = MeasureStep(() => new DocxReader().Read(BuildPackage(), diagnosticSink: null, CancellationToken.None, markupMode));
    }
    else
    {
        (_, readMetrics) = MeasureStep(() => new PptxReader().Read(BuildPackage(), CancellationToken.None));
        OoxPackage scenePackage = BuildPackage();
        PptxDocument sceneDocument = new PptxReader().Read(scenePackage, CancellationToken.None);
        (_, sceneMetrics) = MeasureStep(() => new PptxSceneBuilder().Build(sceneDocument, scenePackage, CancellationToken.None));
    }

    System.Collections.Generic.IReadOnlyList<Lokad.OoxPdf.Pdf.PdfPage> pages;
    object renderMetrics;
    {
        OoxPackage renderPackage = BuildPackage();
        if (isDocx)
        {
            DocxDocument renderDocument = new DocxReader().Read(renderPackage, diagnosticSink: null, CancellationToken.None, markupMode);
            (pages, renderMetrics) = MeasureStep(() => new DocxRenderer(fontResolver: null, markupMode, geometryMode).RenderBlankPages(renderDocument, diagnosticSink: null, CancellationToken.None));
        }
        else
        {
            PptxDocument renderDocument = new PptxReader().Read(renderPackage, CancellationToken.None);
            (pages, renderMetrics) = MeasureStep(() => new PptxRenderer(fontResolver: null).RenderPages(renderDocument, renderPackage, diagnosticSink: null, CancellationToken.None));
        }
    }

    (byte[] stagedBytes, object writeMetrics) = MeasureStep(() =>
    {
        using var output = new MemoryStream();
        Lokad.OoxPdf.Pdf.PdfDocumentWriter.WriteBlank(output, pages, CancellationToken.None, creationDate: null);
        return output.ToArray();
    });
    string stagedSha = Convert.ToHexString(SHA256.HashData(stagedBytes)).ToLowerInvariant();

    return new
    {
        stageSet = isDocx ? "open/read/render/write" : "open/read/scene/render/write",
        stageOrderNote = "Stages run after whole-pipeline cold+warm; each stage rebuilds prerequisites unmeasured. Static caches are warm.",
        open = openMetrics,
        read = readMetrics,
        scene = sceneMetrics,
        render = renderMetrics,
        write = writeMetrics,
        stagedOutputBytes = stagedBytes.Length,
        stagedMatchesWhole = string.Equals(stagedSha, wholeOutputSha, StringComparison.Ordinal),
    };
}

static object MeasureOnce(byte[] inputBytes, OoxPdfOptions options, bool recordOutput, List<string> outputShas, List<int> outputLengths, List<int> pageCounts)
{
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();
    long startAllocated = GC.GetAllocatedBytesForCurrentThread();
    int startGen0 = GC.CollectionCount(0);
    int startGen1 = GC.CollectionCount(1);
    int startGen2 = GC.CollectionCount(2);
    var watch = Stopwatch.StartNew();
    byte[] outputBytes;
    using (var input = new MemoryStream(inputBytes, writable: false))
    using (var output = new MemoryStream())
    {
        OoxPdfConverter.Convert(input, output, options);
        outputBytes = output.ToArray();
    }

    watch.Stop();
    long retainedBytes = GC.GetTotalMemory(forceFullCollection: false);
    if (recordOutput)
    {
        outputShas.Add(Convert.ToHexString(SHA256.HashData(outputBytes)).ToLowerInvariant());
        outputLengths.Add(outputBytes.Length);
        pageCounts.Add(ReadPageCount(outputBytes));
    }

    return new
    {
        allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - startAllocated,
        gen0Collections = GC.CollectionCount(0) - startGen0,
        gen1Collections = GC.CollectionCount(1) - startGen1,
        gen2Collections = GC.CollectionCount(2) - startGen2,
        elapsedMilliseconds = watch.Elapsed.TotalMilliseconds,
        retainedBytes,
    };
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
            dynamic report = MeasureInput(name, package, warmup: 0, iterations: 1, measureStages: false);
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

    return ok ? 0 : 1;
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
