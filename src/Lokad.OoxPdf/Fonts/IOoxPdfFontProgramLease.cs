namespace Lokad.OoxPdf.Fonts;

/// <summary>An owned borrow of immutable font-program bytes.</summary>
/// <remarks>
/// Keep the lease alive until all readers of Bytes finish. Disposal ends this
/// borrow and must not throw; it never disposes the caller's budget. Bytes must
/// not be used after disposal. Sharing sources refund or reuse storage only
/// after its final owner has released it.
/// </remarks>
public interface IOoxPdfFontProgramLease : IDisposable
{
    ReadOnlyMemory<byte> Bytes { get; }
}
