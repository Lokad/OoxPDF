namespace Lokad.OoxPdf.Fonts;

public static class OoxPdfFontPackDiagnosticIds
{
    public const string FontPackDownloadFailed = "FONT_PACK_DOWNLOAD_FAILED";
    public const string FontPackHashMismatch = "FONT_PACK_HASH_MISMATCH";
    public const string FontPackInvalid = "FONT_PACK_INVALID";
    public const string FontPackMissing = "FONT_PACK_MISSING";
}

public sealed class OoxPdfFontPackException : Exception
{
    public OoxPdfFontPackException(string diagnosticId, string message)
        : base(message)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(diagnosticId);
        DiagnosticId = diagnosticId;
    }

    public OoxPdfFontPackException(string diagnosticId, string message, Exception innerException)
        : base(message, innerException)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(diagnosticId);
        DiagnosticId = diagnosticId;
    }

    public string DiagnosticId { get; }
}
