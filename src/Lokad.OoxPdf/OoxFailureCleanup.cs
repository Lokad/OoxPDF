namespace Lokad.OoxPdf;

// Only used while unwinding an existing failure. Successful ownership scopes
// still propagate disposal failures normally.
internal static class OoxFailureCleanup
{
    internal const string SecondaryExceptionKey = "Lokad.OoxPdf.CleanupException";

    internal static async ValueTask DisposeAsync(IAsyncDisposable owner, Exception primary)
    {
        try { await owner.DisposeAsync().ConfigureAwait(false); }
        catch (Exception secondary)
        {
            primary.Data[SecondaryExceptionKey] = primary.Data[SecondaryExceptionKey] is Exception previous
                ? new AggregateException(previous, secondary) : secondary;
        }
    }
}
