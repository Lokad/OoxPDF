namespace Lokad.OoxPdf.Fonts;

// Actual ownership shared by the loader, retained cache entry and leases.
// Eviction releases cache retention, never an outstanding reader's storage.
internal sealed class FontProgramBuffer(ReadOnlyMemory<byte> bytes, IMemoryReservation? reservation) : IDisposable
{
    private ReadOnlyMemory<byte> bytes = bytes;
    private IMemoryReservation? reservation = reservation;
    private int references = 1, loaderReleased;
    internal ReadOnlyMemory<byte> Bytes => bytes;

    internal void Retain()
    {
        int count;
        do
        {
            count = Volatile.Read(ref references);
            ObjectDisposedException.ThrowIf(count == 0, this);
        }
        while (Interlocked.CompareExchange(ref references, checked(count + 1), count) != count);
    }

    internal IOoxPdfFontProgramLease Acquire()
    {
        Retain();
        return new Lease(this);
    }

    internal void Release()
    {
        if (Interlocked.Decrement(ref references) != 0) return;
        bytes = default;
        Interlocked.Exchange(ref reservation, null)?.Dispose();
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref loaderReleased, 1) == 0) Release();
    }

    private sealed class Lease(FontProgramBuffer owner) : IOoxPdfFontProgramLease
    {
        private FontProgramBuffer? owner = owner;
        public ReadOnlyMemory<byte> Bytes => (Volatile.Read(ref owner)
            ?? throw new ObjectDisposedException(nameof(IOoxPdfFontProgramLease))).Bytes;
        public void Dispose() => Interlocked.Exchange(ref owner, null)?.Release();
    }
}
