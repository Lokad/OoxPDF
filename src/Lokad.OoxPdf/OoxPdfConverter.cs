using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf;

public static class OoxPdfConverter
{
    public static void Convert(string inputPath, string outputPath)
    {
        Convert(inputPath, outputPath, new OoxPdfOptions(), CancellationToken.None);
    }

    public static void Convert(string inputPath, string outputPath, CancellationToken cancellationToken)
    {
        Convert(inputPath, outputPath, new OoxPdfOptions(), cancellationToken);
    }

    public static void Convert(string inputPath, string outputPath, OoxPdfOptions? options)
    {
        Convert(inputPath, outputPath, options, CancellationToken.None);
    }

    public static void Convert(string inputPath, string outputPath, OoxPdfOptions? options, CancellationToken cancellationToken)
    {
        ConvertCore(inputPath, outputPath, options, cancellationToken);
    }

    public static Task ConvertAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
    {
        return ConvertAsync(inputPath, outputPath, new OoxPdfOptions(), cancellationToken);
    }

    public static Task ConvertAsync(string inputPath, string outputPath, OoxPdfOptions? options, CancellationToken cancellationToken = default)
    {
        return Task.Run(() => ConvertCore(inputPath, outputPath, options, cancellationToken), cancellationToken);
    }

    private static void ConvertCore(string inputPath, string outputPath, OoxPdfOptions? options, CancellationToken cancellationToken)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        options ??= new OoxPdfOptions();
        cancellationToken.ThrowIfCancellationRequested();

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input OOXML document was not found.", inputPath);
        }

        OoxPdfInputKind inputKind = DetectInputKind(inputPath, options.InputKind);
        cancellationToken.ThrowIfCancellationRequested();

        OoxPackage package = OoxPackage.Open(inputPath, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<PdfPage> pages = inputKind switch
        {
            OoxPdfInputKind.Pptx => new PptxRenderer(options.FontResolver).RenderPages(new PptxReader().Read(package, cancellationToken), package, options.DiagnosticSink, cancellationToken),
            OoxPdfInputKind.Docx => new DocxRenderer(options.FontResolver).RenderBlankPages(new DocxReader().Read(package, options.DiagnosticSink, cancellationToken), options.DiagnosticSink, cancellationToken),
            _ => throw new NotSupportedException($"Unsupported input kind '{inputKind}'.")
        };

        cancellationToken.ThrowIfCancellationRequested();
        string? outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrEmpty(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using FileStream output = File.Create(outputPath);
        PdfDocumentWriter.WriteBlank(output, pages, cancellationToken);
    }

    public static OoxPdfInputKind DetectInputKind(string inputPath, OoxPdfInputKind requestedKind = OoxPdfInputKind.Auto)
    {
        if (requestedKind is OoxPdfInputKind.Pptx or OoxPdfInputKind.Docx)
        {
            return requestedKind;
        }

        string extension = Path.GetExtension(inputPath);
        if (extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase))
        {
            return OoxPdfInputKind.Pptx;
        }

        if (extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            return OoxPdfInputKind.Docx;
        }

        throw new NotSupportedException($"Unsupported OOXML input extension '{extension}'. Expected .pptx or .docx.");
    }
}
