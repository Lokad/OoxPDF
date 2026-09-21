using Lokad.OoxPdf.Docx;
using Lokad.OoxPdf.Ooxml;
using Lokad.OoxPdf.Pdf;
using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf;

/// <summary>
/// Dependency-free PPTX/DOCX to PDF conversion entry points.
/// </summary>
/// <remarks>
/// <para>Operational contracts for hosting services:</para>
/// <list type="bullet">
/// <item>Synchronous and asynchronous overloads perform the same conversion. The Async
/// variants offload the synchronous pipeline to the thread pool and perform no true
/// asynchronous I/O; awaiting one yields byte-identical output to the matching sync
/// overload for the same inputs.</item>
/// <item>Caller-owned streams are never closed or disposed. Input is read sequentially
/// and output needs only forward writes; non-seekable host streams are supported.
/// File outputs publish atomically: any failure, including cancellation, leaves a
/// pre-existing destination untouched and removes only the staging file.</item>
/// <item>Cancellation is observed cooperatively at stage boundaries and inside bounded
/// expansion loops. A cancelled conversion throws <see cref="OperationCanceledException"/>
/// and never produces a partial PDF.</item>
/// <item>Resolvers and font sources must be safe for concurrent use. Parallel
/// conversions may share one resolver instance and observe byte-identical output;
/// font-program bytes are treated as immutable once published.</item>
/// <item>Strict affects CLI exit policy, not library refusal to emit a PDF;
/// Deterministic is accepted for compatibility and output is deterministic by
/// construction (see <see cref="OoxPdfOptions"/>).</item>
/// </list>
/// </remarks>
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

    public static Task ConvertAsync(string inputPath, string outputPath)
    {
        return ConvertAsync(inputPath, outputPath, new OoxPdfOptions(), CancellationToken.None);
    }

    public static Task ConvertAsync(string inputPath, string outputPath, OoxPdfOptions? options)
    {
        return ConvertAsync(inputPath, outputPath, options, CancellationToken.None);
    }

    public static Task ConvertAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
    {
        return ConvertAsync(inputPath, outputPath, new OoxPdfOptions(), cancellationToken);
    }

    public static Task ConvertAsync(string inputPath, string outputPath, OoxPdfOptions? options, CancellationToken cancellationToken = default)
    {
        return Task.Run(() => ConvertCore(inputPath, outputPath, options, cancellationToken), cancellationToken);
    }

    public static Task ConvertAsync(Stream input, Stream output, OoxPdfOptions? options)
    {
        return ConvertAsync(input, output, options, CancellationToken.None);
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
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        if (string.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Input and output file paths must be different.", nameof(outputPath));
        }

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input OOXML document was not found.", inputPath);
        }

        OoxPdfInputKind inputKind = DetectInputKind(inputPath, options.InputKind);
        cancellationToken.ThrowIfCancellationRequested();

        using FileStream input = File.OpenRead(inputPath);
        // PLAN Q01: one explicitly scoped conversion budget per Convert call.
        // Totals are snapshotted before publication; the summary below only
        // emits after the atomic move succeeds, never for partial output.
        IReadOnlyList<PdfPage> pages;
        OoxConversionTotals totals;
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(options.ConversionLimits))
        {
            pages = RenderPages(input, inputKind, options, cancellationToken);
            totals = scope.Budget.Totals;
        }

        cancellationToken.ThrowIfCancellationRequested();
        string? outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrEmpty(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        // Publish atomically: render into a uniquely named temp file beside the destination
        // (same volume, so the final move is atomic) and move it over the destination once.
        // A failure at any point leaves a pre-existing destination untouched and removes
        // only the temporary file owned by this conversion.
        string stagingPath = Path.Combine(outputDirectory!, Path.GetFileName(outputPath) + ".tmp-" + Guid.NewGuid().ToString("N"));
        try
        {
            using (FileStream output = File.Create(stagingPath))
            {
                PdfDocumentWriter.WriteBlank(output, pages, cancellationToken, options.FixedCreationDate);
            }

            cancellationToken.ThrowIfCancellationRequested();
            File.Move(stagingPath, outputPath, overwrite: true);
            ReportResourceUsage(options, totals, pages.Count);
        }
        finally
        {
            try
            {
                if (File.Exists(stagingPath))
                {
                    File.Delete(stagingPath);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
            }
        }
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

        if (ReferenceEquals(input, output))
        {
            throw new ArgumentException("Input and output streams must be distinct instances.", nameof(output));
        }

        if (output.CanSeek)
        {
            if (output.Position != 0)
            {
                throw new ArgumentException("Output stream must be positioned at the start (Position 0) so PDF cross-reference offsets are valid.", nameof(output));
            }

            if (output.Length != 0)
            {
                throw new ArgumentException("Output stream must be empty (Length 0) so no trailing bytes remain after the PDF.", nameof(output));
            }
        }

        options ??= new OoxPdfOptions();
        options.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        OoxPdfInputKind inputKind = RequireExplicitInputKind(options.InputKind);
        IReadOnlyList<PdfPage> pages;
        OoxConversionTotals totals;
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(options.ConversionLimits))
        {
            pages = RenderPages(input, inputKind, options, cancellationToken);
            totals = scope.Budget.Totals;
        }

        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentWriter.WriteBlank(output, pages, cancellationToken, options.FixedCreationDate);
        ReportResourceUsage(options, totals, pages.Count);
    }

    private static void ReportResourceUsage(OoxPdfOptions options, OoxConversionTotals totals, int pageCount)
    {
        if (options.ReportResourceUsage)
        {
            options.DiagnosticSink?.Invoke(totals.ToSummaryDiagnostic(pageCount));
        }
    }

    private static IReadOnlyList<PdfPage> RenderPages(Stream input, OoxPdfInputKind inputKind, OoxPdfOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        OoxPackage package = OoxPackage.Open(input, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        return inputKind switch
        {
            OoxPdfInputKind.Pptx => new PptxRenderer(options.FontResolver).RenderPages(new PptxReader().Read(package, cancellationToken, options.DiagnosticSink), package, options.DiagnosticSink, cancellationToken),
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

    public static OoxPdfInputKind DetectInputKind(string inputPath)
    {
        return DetectInputKind(inputPath, OoxPdfInputKind.Auto);
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
