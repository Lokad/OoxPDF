namespace Lokad.OoxPdf.Fonts;

// Loading stays outside this lock. It owns cache retention; leases independently
// keep the immutable program and its admission alive through eviction/shutdown.
internal sealed class FontProgramCache(long maxBytes, StringComparer comparer, IMemoryBudget? budget = null) : IDisposable
{
    private readonly object sync = new();
    private readonly Dictionary<string, LinkedListNode<Entry>> entries = new(comparer);
    private readonly LinkedList<Entry> order = new();
    private long retainedBytes;
    private bool disposed;
    private sealed record Entry(string Key, FontProgramBuffer Program);
    internal IMemoryBudget? Budget => budget;
    internal long RetainedBytes { get { lock (sync) { return retainedBytes; } } }

    // Never lend an operation's short-lived budget to cache retention. Without a
    // cache budget, governed operations acquire private funded programs instead.
    internal bool CanRetain(IMemoryBudget? operationBudget) => budget is not null || operationBudget is null;

    internal bool TryAcquire(string key, IMemoryBudget? operationBudget, out IOoxPdfFontProgramLease? lease)
    {
        lock (sync)
        {
            ObjectDisposedException.ThrowIf(disposed, this);
            if (CanRetain(operationBudget) && entries.TryGetValue(key, out var node))
            {
                lease = node.Value.Program.Acquire();
                order.Remove(node); order.AddLast(node);
                return true;
            }
            lease = null;
            return false;
        }
    }

    internal void Store(string key, FontProgramBuffer program, IMemoryBudget? operationBudget)
    {
        List<FontProgramBuffer>? retired = null;
        try
        {
            lock (sync)
            {
                ObjectDisposedException.ThrowIf(disposed, this);
                if (!CanRetain(operationBudget)) return;
                if (entries.Remove(key, out var previous)) Retire(previous);
                long length = program.Bytes.Length;
                if (length > maxBytes) return;
                while (retainedBytes > maxBytes - length && order.First is { } eldest)
                {
                    entries.Remove(eldest.Value.Key); Retire(eldest);
                }
                program.Retain();
                entries.Add(key, order.AddLast(new Entry(key, program)));
                retainedBytes += length;
            }
        }
        finally { if (retired is not null) foreach (var old in retired) old.Release(); }

        void Retire(LinkedListNode<Entry> node)
        {
            order.Remove(node); retainedBytes -= node.Value.Program.Bytes.Length;
            (retired ??= []).Add(node.Value.Program);
        }
    }

    public void Dispose()
    {
        FontProgramBuffer[] retired;
        lock (sync)
        {
            if (disposed) return;
            disposed = true;
            retired = order.Select(entry => entry.Program).ToArray();
            entries.Clear(); order.Clear(); retainedBytes = 0;
        }
        foreach (var program in retired) program.Release();
    }
}
