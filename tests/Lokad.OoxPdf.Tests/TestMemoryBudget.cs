namespace Lokad.OoxPdf.Tests;

// A caller implementation, deliberately outside the shipped library. Each
// scope owns an aggregate amount; there is no reservation graph or object ledger.
internal sealed class TestMemoryBudget(long capacity = 128L * 1024 * 1024,
    Action? onCreate = null, Action? onDispose = null) : IMemoryBudget, IDisposable
{
    private readonly object sync = new();
    private readonly Action? rootDisposed = onDispose;
    public long Capacity = capacity;
    public long Reserved { get; private set; }
    public long Peak { get; private set; }
    public int RootsCreated { get; private set; }
    public int RootsDisposed { get; private set; }
    public int ScopesCreated { get; private set; }
    public int ScopesDisposed { get; private set; }
    public int Requests { get; private set; }
    public int CallerDisposals { get; private set; }
    public Func<long, long, bool>? RefuseChange { get; set; }

    public IMemoryScope CreateScope()
    {
        var scope = CreateScopeCore(true);
        try { onCreate?.Invoke(); return scope; }
        catch { scope.Dispose(); throw; }
    }

    private Scope CreateScopeCore(bool root)
    {
        lock (sync) { ScopesCreated++; if (root) RootsCreated++; }
        return new Scope(this, root);
    }

    public bool TryReserve(long bytes, out IMemoryReservation? reservation) => Reserve(null, bytes, out reservation);

    private bool Reserve(Scope? scope, long bytes, out IMemoryReservation? reservation)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(bytes);
        lock (sync)
        {
            ObjectDisposedException.ThrowIf(scope?.Closed == true, scope!);
            Requests++;
            if (bytes > Capacity - Reserved || RefuseChange?.Invoke(0, bytes) == true)
            { reservation = null; return false; }
            Reserved += bytes;
            if (scope is not null) scope.Owned += bytes;
            Peak = Math.Max(Peak, Reserved);
            reservation = new Reservation(this, scope, bytes);
            return true;
        }
    }

    public void Dispose() { CallerDisposals++; }

    private sealed class Scope(TestMemoryBudget pool, bool root) : IMemoryScope
    {
        public long Owned;
        public bool Closed;
        public IMemoryScope CreateScope()
        {
            lock (pool.sync)
            {
                ObjectDisposedException.ThrowIf(Closed, this);
                return pool.CreateScopeCore(false);
            }
        }
        public bool TryReserve(long bytes, out IMemoryReservation? reservation) => pool.Reserve(this, bytes, out reservation);
        public void Dispose()
        {
            lock (pool.sync)
            {
                if (Closed) return;
                Closed = true;
                pool.Reserved -= Owned;
                Owned = 0;
                pool.ScopesDisposed++;
                if (root) pool.RootsDisposed++;
            }
            if (root) pool.rootDisposed?.Invoke();
        }
    }

    private sealed class Reservation(TestMemoryBudget pool, Scope? scope, long bytes) : IMemoryReservation
    {
        private long held = bytes;
        private bool closed;
        public long Bytes { get { lock (pool.sync) return scope?.Closed == true ? 0 : held; } }
        public bool TryResize(long bytes)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(bytes);
            lock (pool.sync)
            {
                ObjectDisposedException.ThrowIf(closed || scope?.Closed == true, this);
                long delta = bytes - held;
                pool.Requests++;
                if (delta > 0 && (delta > pool.Capacity - pool.Reserved || pool.RefuseChange?.Invoke(held, bytes) == true)) return false;
                pool.Reserved += delta;
                if (scope is not null) scope.Owned += delta;
                held = bytes;
                pool.Peak = Math.Max(pool.Peak, pool.Reserved);
                return true;
            }
        }
        public void Dispose()
        {
            lock (pool.sync)
            {
                if (closed) return;
                closed = true;
                if (scope?.Closed != true)
                {
                    pool.Reserved -= held;
                    if (scope is not null) scope.Owned -= held;
                }
                held = 0;
            }
        }
    }
}
