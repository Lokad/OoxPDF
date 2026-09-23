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
/// pre-existing destination untouched and removes only the staging file. Observer callbacks
/// (<see cref="OoxPdfOptions.DiagnosticSink"/>) run before publication, so a throwing observer fails
/// the conversion without publishing anything (a post-publication throw would falsely report
/// rollback of an already-replaced destination).</item>
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
        // one explicitly scoped conversion budget per Convert call.
        // Totals are snapshotted before publication; the summary below emits before the
        // atomic move succeeds, so a throwing observer fails the conversion while the
        // pre-existing destination is still untouched (R20).
        // R06.1: the conversion scope stays open through serialization so page and
        // content budgets bind the renderers (R06.2) and each output chunk is admitted against
        // the output budget before its write; totals snapshot after serialization so
        // the summary reports writer-stage fields, still before the atomic move (R20).
        OoxConversionTotals totals;
        int pageCount;
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
            using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(options.ConversionLimits))
            {
                // R06.3: the writer drains the producer enumerable once, spilling page
                // content past the resident window and retaining only descriptors.
                using (FileStream output = File.Create(stagingPath))
                {
                    PdfDocumentWriter.WriteStaged(output, RenderPages(input, inputKind, options, cancellationToken), options.ConversionLimits ?? new OoxConversionLimits(), options.DiagnosticSink, cancellationToken, options.FixedCreationDate);
                }
                totals = scope.Budget.Totals;
                pageCount = checked((int)totals.PdfPages);
            }

            cancellationToken.ThrowIfCancellationRequested();
            ReportResourceUsage(options, totals, pageCount);
            File.Move(stagingPath, outputPath, overwrite: true);
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
        OoxConversionTotals totals;
        // R06.1: like the file path, the scope stays open through serialization and
        // output chunks admit incrementally during WriteBlank; the
        // summary still reports pre-write totals (writer-stage fields stay zero here)
        // so a throwing observer leaves stream output untouched (R20), unlike the
        // file path which snapshots after serialization into its staging file.
        using (OoxConversionBudget.Scope scope = OoxConversionBudget.BeginScope(options.ConversionLimits))
        {
            // R06.3: produce (drain, spill, blank) first so the totals snapshot and the
            // observer report still precede serialization, preserving the stream rules;
            // emission follows from staging.
            var (blanked, staging) = PdfDocumentWriter.ProduceStagedPages(RenderPages(input, inputKind, options, cancellationToken), options.ConversionLimits ?? new OoxConversionLimits(), options.DiagnosticSink, cancellationToken);
            using (staging)
            {
                totals = scope.Budget.Totals;
                cancellationToken.ThrowIfCancellationRequested();
                ReportResourceUsage(options, totals, checked((int)scope.Budget.PdfPages));
                PdfDocumentWriter.EmitStaged(output, blanked, staging, cancellationToken, options.FixedCreationDate);
            }
        }
    }

    private static void ReportResourceUsage(OoxPdfOptions options, OoxConversionTotals totals, int pageCount)
    {
        if (options.ReportResourceUsage)
        {
            options.DiagnosticSink?.Invoke(totals.ToSummaryDiagnostic(pageCount));
        }
    }

    private static IEnumerable<PdfPage> RenderPages(Stream input, OoxPdfInputKind inputKind, OoxPdfOptions options, CancellationToken cancellationToken)
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
