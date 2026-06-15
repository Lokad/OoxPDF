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

    public bool Strict { get; init; }

    public bool Deterministic { get; init; } = true;

    public DateTimeOffset? FixedCreationDate { get; init; }

    public IFontResolver? FontResolver { get; init; }

    public Action<OoxPdfDiagnostic>? DiagnosticSink { get; init; }
}
