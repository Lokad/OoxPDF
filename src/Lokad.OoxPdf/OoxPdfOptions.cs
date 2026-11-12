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
    /// When set, the PDF document information dictionary records this value as CreationDate and
    /// ModDate. When null (the default), no dates are emitted and output stays deterministic.
    /// </summary>
    public DateTimeOffset? FixedCreationDate { get; init; }

    /// <summary>
    /// Complete allowed font source. A custom resolver does not acquire additional
    /// system fonts implicitly. Null selects the local Windows font resolver.
    /// </summary>
    public IFontResolver? FontResolver { get; init; }

    public Action<OoxPdfDiagnostic>? DiagnosticSink { get; init; }

    /// <summary>
    /// Output-size/page limits and the resident page-content staging window.
    /// Null selects the built-in defaults. Individual archive, XML, image and
    /// font bounds apply independently.
    /// </summary>
    public OoxConversionLimits? ConversionLimits { get; init; }

    /// <summary>
    /// Optional caller-owned shared capacity for coarse package/model storage,
    /// page construction and selected PDF/font/image/vector resources. Producers
    /// reserve covered capacity before acquisition and refund it after final use.
    /// Coverage is partial, not a whole-heap limit. Longer-lived font caches use
    /// separately supplied cache budgets. Null leaves memory ungoverned;
    /// independent format/output bounds still apply.
    /// </summary>
    public IMemoryBudget? MemoryBudget { get; init; }

    /// <summary>
    /// Optional owned temporary backing for content beyond the resident staging window.
    /// Stream conversions without a provider stay in admitted memory or fail at that
    /// window. File conversions default to explicitly selected filesystem backing.
    /// The provider borrows the conversion budget and returns independently owned streams.
    /// </summary>
    public IOoxPdfTemporaryStorage? TemporaryStorage { get; init; }

    /// <summary>
    /// Emits an informational CONVERSION_RESOURCE_SUMMARY with page/content,
    /// serialized PDF output, font/image payload and spilled-content byte totals.
    /// Font totals count uncompressed embedded programs and ToUnicode data before
    /// PDF stream compression, excluding dictionaries and Standard-14 ToUnicode maps.
    /// Stream conversions report before serialization so observers can fail without
    /// writing output; file conversions report before atomic publication.
    /// This does not measure memory. Disabled by default.
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
