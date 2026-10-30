namespace Lokad.OoxPdf;

/// <summary>
/// Reports that a conversion exceeded an output/input bound or was refused host memory admission.
/// Derives from IOException so existing
/// IO-based failure handling treats it as a conversion failure, but renderers
/// must not swallow it as recoverable per-node content. It is deliberately not an InvalidDataException so PPTX per-node recovery does not swallow it.
/// </summary>
public sealed class OoxPdfLimitExceededException : IOException
{
    /// <summary>The document bound or shared memory admission category.</summary>
    public OoxPdfLimitKind Kind { get; }

    /// <summary>The conversion or shared resource owning the refused capacity.</summary>
    public OoxPdfLimitScope Scope { get; }

    public OoxPdfLimitExceededException(string message)
        : this(message, OoxPdfLimitKind.DocumentBound, OoxPdfLimitScope.Conversion)
    {
    }

    public OoxPdfLimitExceededException(string message, Exception? innerException)
        : this(message, OoxPdfLimitKind.DocumentBound, OoxPdfLimitScope.Conversion, innerException)
    {
    }

    /// <summary>Creates a structured refusal without changing fatal conversion handling.</summary>
    public OoxPdfLimitExceededException(string message, OoxPdfLimitKind kind,
        OoxPdfLimitScope scope, Exception? innerException = null)
        : base(message, innerException)
    {
        Kind = kind;
        Scope = scope;
    }
}
