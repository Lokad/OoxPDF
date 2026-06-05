using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf;

public sealed class OoxPdfOptions
{
    public OoxPdfInputKind InputKind { get; init; } = OoxPdfInputKind.Auto;

    public bool Strict { get; init; }

    public bool Deterministic { get; init; } = true;

    public DateTimeOffset? FixedCreationDate { get; init; }

    public IFontResolver? FontResolver { get; init; }

    public Action<OoxPdfDiagnostic>? DiagnosticSink { get; init; }
}
