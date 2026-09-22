using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf;

public sealed class OoxPdfOptions
{
    public OoxPdfInputKind InputKind { get; init; } = OoxPdfInputKind.Auto;

    /// <summary>
    /// Selects the DOCX review markup view. Ignored for PPTX inputs.
    /// </summary>
    public OoxPdfDocxMarkupMode DocxMarkupMode { get; init; } = OoxPdfDocxMarkupMode.Final;

    /// <summary>
    /// Selects whether DOCX markup balloons preserve the authored body layout or reserve a review margin. Ignored for PPTX inputs.
    /// </summary>
    public OoxPdfDocxMarkupGeometryMode DocxMarkupGeometryMode { get; init; } = OoxPdfDocxMarkupGeometryMode.PreserveDocumentLayout;

    /// <summary>
    /// Requests strict handling of diagnostics. The library validates options eagerly and
    /// reports unsupported content through <see cref="DiagnosticSink"/>; the CLI maps any
    /// warning or error diagnostic to process exit code 3 when this flag is set.
    /// </summary>
    public bool Strict { get; init; }

    /// <summary>
    /// Conversion output is deterministic by construction (no timestamps, random values, or
    /// absolute paths are emitted unless <see cref="FixedCreationDate"/> is set). The flag is
    /// accepted for compatibility and currently has no additional effect.
    /// </summary>
    public bool Deterministic { get; init; } = true;

    /// <summary>
    /// When set, the PDF document information dictionary records this value as CreationDate and
    /// ModDate. When null (the default), no dates are emitted and output stays deterministic.
    /// </summary>
    public DateTimeOffset? FixedCreationDate { get; init; }

    public IFontResolver? FontResolver { get; init; }

    public Action<OoxPdfDiagnostic>? DiagnosticSink { get; init; }

    /// <summary>
    /// Cumulative per-conversion work budgets. When null, generous
    /// built-in defaults apply (see <see cref="OoxConversionLimits"/>). Set tighter
    /// values to enforce a shared-process memory/work limit; budget crossings fail
    /// with <see cref="OoxPdfLimitExceededException"/> before further expansion and
    /// never produce a partial PDF.
    /// </summary>
    public OoxConversionLimits? ConversionLimits { get; init; }

    /// <summary>
    /// When true, a successful conversion emits one informational
    /// CONVERSION_RESOURCE_SUMMARY diagnostic reporting cumulative work counters
    /// (pages, chart cells, table fragments, images, font operations) plus the peak
    /// live image decode reservation so hosts can account concurrent conversions.
    /// Informational diagnostics never affect CLI strict exit codes. Disabled by default.
    /// </summary>
    public bool ReportResourceUsage { get; init; }


    internal void Validate()
    {
        if (!Enum.IsDefined(InputKind))
        {
            throw new ArgumentOutOfRangeException(nameof(InputKind), "Unsupported OOXML input kind.");
        }

        if (!Enum.IsDefined(DocxMarkupMode))
        {
            throw new ArgumentOutOfRangeException(nameof(DocxMarkupMode), "Unsupported DOCX markup mode.");
        }

        if (!Enum.IsDefined(DocxMarkupGeometryMode))
        {
            throw new ArgumentOutOfRangeException(nameof(DocxMarkupGeometryMode), "Unsupported DOCX markup geometry mode.");
        }

        ConversionLimits?.Validate();
    }
}
