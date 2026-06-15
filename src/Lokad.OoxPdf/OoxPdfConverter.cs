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

    public static void Convert(Stream input, Stream output, OoxPdfOptions? options)
    {
        Convert(input, output, options, CancellationToken.None);
    }

    public static void Convert(Stream input, Stream output, OoxPdfOptions? options, CancellationToken cancellationToken)
    {
        ConvertCore(input, output, options, cancellationToken);
    }

    public static Task ConvertAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
    {
        return ConvertAsync(inputPath, outputPath, new OoxPdfOptions(), cancellationToken);
    }

    public static Task ConvertAsync(string inputPath, string outputPath, OoxPdfOptions? options, CancellationToken cancellationToken = default)
    {
        return Task.Run(() => ConvertCore(inputPath, outputPath, options, cancellationToken), cancellationToken);
    }

    public static Task ConvertAsync(Stream input, Stream output, OoxPdfOptions? options, CancellationToken cancellationToken = default)
    {
        return Task.Run(() => ConvertCore(input, output, options, cancellationToken), cancellationToken);
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

        using FileStream input = File.OpenRead(inputPath);
        IReadOnlyList<PdfPage> pages = RenderPages(input, inputKind, options, cancellationToken);

        cancellationToken.ThrowIfCancellationRequested();
        string? outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrEmpty(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using FileStream output = File.Create(outputPath);
        PdfDocumentWriter.WriteBlank(output, pages, cancellationToken);
    }

    private static void ConvertCore(Stream input, Stream output, OoxPdfOptions? options, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(output);

        if (!input.CanRead)
        {
            throw new ArgumentException("Input stream must be readable.", nameof(input));
        }

        if (!output.CanWrite)
        {
            throw new ArgumentException("Output stream must be writable.", nameof(output));
        }

        options ??= new OoxPdfOptions();
        cancellationToken.ThrowIfCancellationRequested();

        OoxPdfInputKind inputKind = RequireExplicitInputKind(options.InputKind);
        IReadOnlyList<PdfPage> pages = RenderPages(input, inputKind, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentWriter.WriteBlank(output, pages, cancellationToken);
    }

    private static IReadOnlyList<PdfPage> RenderPages(Stream input, OoxPdfInputKind inputKind, OoxPdfOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPackage package = OoxPackage.Open(input, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        return inputKind switch
        {
            OoxPdfInputKind.Pptx => new PptxRenderer(options.FontResolver).RenderPages(new PptxReader().Read(package, cancellationToken), package, options.DiagnosticSink, cancellationToken),
            OoxPdfInputKind.Docx => new DocxRenderer(options.FontResolver, options.DocxMarkupMode, options.DocxMarkupGeometryMode).RenderBlankPages(new DocxReader().Read(package, options.DiagnosticSink, cancellationToken, options.DocxMarkupMode), options.DiagnosticSink, cancellationToken),
            _ => throw new NotSupportedException($"Unsupported input kind '{inputKind}'.")
        };
    }

    private static OoxPdfInputKind RequireExplicitInputKind(OoxPdfInputKind requestedKind)
    {
        return requestedKind switch
        {
            OoxPdfInputKind.Pptx or OoxPdfInputKind.Docx => requestedKind,
            OoxPdfInputKind.Auto => throw new NotSupportedException("Stream input requires OoxPdfOptions.InputKind to be OoxPdfInputKind.Pptx or OoxPdfInputKind.Docx because no file extension is available for auto-detection."),
            _ => throw new NotSupportedException($"Unsupported input kind '{requestedKind}'.")
        };
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
