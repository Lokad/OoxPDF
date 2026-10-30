namespace Lokad.OoxPdf.Fonts;

internal static class FontProgramLoader
{
    public static OpenTypeFont? Load(FontFaceResolution? resolution, CancellationToken cancellationToken) =>
        LoadAsync(resolution, cancellationToken).GetAwaiter().GetResult();

    public static async ValueTask<OpenTypeFont?> LoadAsync(FontFaceResolution? resolution, CancellationToken cancellationToken)
    {
        if (resolution is null)
        {
            return null;
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            // Parsing takes an owned SFNT copy (or extracts one collection face).
            // Its source lease ends here; the conversion owns the parsed copy.
            using IOoxPdfFontProgramLease lease = await resolution.Source.AcquireAsync(
                OoxConversionBudget.Current?.Memory, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return OpenTypeFont.LoadShared(lease.Bytes.Span, resolution.FontFaceIndex, cancellationToken);
        }
        catch (Exception ex) when (ex is not OoxPdfLimitExceededException && (ex is IOException or InvalidDataException or NotSupportedException or ArgumentOutOfRangeException or UnauthorizedAccessException))
        {
            return null;
        }
    }
}
