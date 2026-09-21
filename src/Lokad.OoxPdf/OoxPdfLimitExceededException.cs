namespace Lokad.OoxPdf;

/// <summary>
/// Reports that a conversion exceeded a bounded resource or work budget
/// (PLAN Q01/Q02). Derives from IOException so existing
/// IO-based failure handling treats it as a conversion failure, but renderers
/// must not swallow it as recoverable per-node content. It is deliberately not an InvalidDataException so PPTX per-node recovery does not swallow it.
/// </summary>
public sealed class OoxPdfLimitExceededException : IOException
{
    public OoxPdfLimitExceededException(string message)
        : base(message)
    {
    }

    public OoxPdfLimitExceededException(string message, Exception? innerException)
        : base(message, innerException)
    {
    }
}
