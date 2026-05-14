using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf;

public static class OoxPdfConverter
{
    public static void Convert(string inputPath, string outputPath)
    {
        Convert(inputPath, outputPath, new OoxPdfOptions());
    }

    public static void Convert(string inputPath, string outputPath, OoxPdfOptions? options)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        options ??= new OoxPdfOptions();

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input OOXML document was not found.", inputPath);
        }

        OoxPdfInputKind inputKind = DetectInputKind(inputPath, options.InputKind);

        OoxPackage package = OoxPackage.Open(inputPath);
        IReadOnlyList<PdfPage> pages = inputKind switch
        {
            OoxPdfInputKind.Pptx => new PptxRenderer().RenderPages(new PptxReader().Read(package), package, options.DiagnosticSink),
            OoxPdfInputKind.Docx => new DocxRenderer().RenderBlankPages(new DocxReader().Read(package, options.DiagnosticSink)),
            _ => throw new NotSupportedException($"Unsupported input kind '{inputKind}'.")
        };

        string? outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrEmpty(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using FileStream output = File.Create(outputPath);
        PdfDocumentWriter.WriteBlank(output, pages);
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
