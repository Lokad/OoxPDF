using System.Text.Json;
using Lokad.OoxPdf;
using Lokad.OoxPdf.Diagnostics;

return Run(args);

static int Run(string[] args)
{
    if (args.Length < 1 || !args[0].Equals("convert", StringComparison.OrdinalIgnoreCase))
    {
        PrintUsage();
        return 2;
    }

    if (args.Length < 3)
    {
        PrintUsage();
        return 2;
    }

    string inputPath = args[1];
    string outputPath = args[2];
    string? diagnosticsPath = null;
    bool strict = false;

    for (int i = 3; i < args.Length; i++)
    {
        string arg = args[i];
        if (arg.Equals("--strict", StringComparison.OrdinalIgnoreCase))
        {
            strict = true;
            continue;
        }

        if (arg.Equals("--diagnostics", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
        {
            diagnosticsPath = args[++i];
            continue;
        }

        Console.Error.WriteLine($"Unknown or incomplete argument: {arg}");
        PrintUsage();
        return 2;
    }

    var collector = new DiagnosticCollector();

    try
    {
        OoxPdfConverter.Convert(inputPath, outputPath, new OoxPdfOptions
        {
            Strict = strict,
            DiagnosticSink = collector.Add
        });

        WriteDiagnostics(diagnosticsPath, collector.Diagnostics);
        return strict && collector.HasWarningsOrErrors ? 3 : 0;
    }
    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
    {
        collector.Add(new OoxPdfDiagnostic("OOXML_CONVERSION_FAILED", OoxPdfSeverity.Error, ex.Message));
        WriteDiagnostics(diagnosticsPath, collector.Diagnostics);
        Console.Error.WriteLine(ex.Message);
        return 1;
    }
}

static void PrintUsage()
{
    Console.Error.WriteLine("Usage: Lokad.OoxPdf.Cli convert <input.pptx|input.docx> <output.pdf> [--diagnostics <file>] [--strict]");
}

static void WriteDiagnostics(string? diagnosticsPath, IReadOnlyList<OoxPdfDiagnostic> diagnostics)
{
    if (diagnosticsPath is null)
    {
        return;
    }

    string? directory = Path.GetDirectoryName(Path.GetFullPath(diagnosticsPath));
    if (!string.IsNullOrEmpty(directory))
    {
        Directory.CreateDirectory(directory);
    }

    var options = new JsonSerializerOptions { WriteIndented = true };
    File.WriteAllText(diagnosticsPath, JsonSerializer.Serialize(diagnostics, options));
}
