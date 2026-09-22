namespace Lokad.OoxPdf.Fonts;

internal static class FontProgramLoader
{
    public static OpenTypeFont? Load(FontFaceResolution? resolution, CancellationToken cancellationToken)
    {
        if (resolution is null)
        {
            return null;
        }

        OoxConversionBudget.Current?.ChargeFontWork(1);

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            // PLAN G03: synchronously completed sources (memory, cached files) must not
            // pay a Task allocation per load. The array copy stays: source bytes are
            // shared across concurrent conversions under the M09 immutability contract,
            // and the parsed font takes ownership of its own buffer at this boundary.
            ValueTask<ReadOnlyMemory<byte>> pending = resolution.Source.GetBytesAsync(cancellationToken);
            ReadOnlyMemory<byte> bytes = pending.IsCompletedSuccessfully
                ? pending.Result
                : pending.AsTask().GetAwaiter().GetResult();
            cancellationToken.ThrowIfCancellationRequested();
            return OpenTypeFont.Load(bytes.ToArray(), resolution.FontFaceIndex, cancellationToken);
        }
        catch (Exception ex) when (ex is not OoxPdfLimitExceededException && (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException))
        {
            return null;
        }
    }
}
