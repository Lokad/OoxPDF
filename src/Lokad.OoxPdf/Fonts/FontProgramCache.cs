namespace Lokad.OoxPdf.Fonts;

// Owns retained arrays and their LRU accounting together. Loading happens outside
// this lock; sources serialize their own loads and recheck the cache first.
// Eviction releases only this owner's reference, never bytes borrowed by callers.
internal sealed class FontProgramCache(long maxBytes, StringComparer comparer)
{
    private readonly object sync = new();
    private readonly Dictionary<string, LinkedListNode<Entry>> entries = new(comparer);
    private readonly LinkedList<Entry> order = new();
    private long retainedBytes;

    private sealed record Entry(string Key, byte[] Bytes);

    internal long RetainedBytes
    {
        get { lock (sync) { return retainedBytes; } }
    }

    internal bool TryGet(string key, out ReadOnlyMemory<byte> bytes)
    {
        lock (sync)
        {
            if (entries.TryGetValue(key, out LinkedListNode<Entry>? node))
            {
                order.Remove(node);
                order.AddLast(node);
                bytes = node.Value.Bytes;
                return true;
            }

            bytes = default;
            return false;
        }
    }

    internal void Store(string key, byte[] bytes)
    {
        lock (sync)
        {
            if (entries.Remove(key, out LinkedListNode<Entry>? previous))
            {
                order.Remove(previous);
                retainedBytes -= previous.Value.Bytes.LongLength;
            }

            // An oversized entry may be returned by its loader but never retained.
            if (bytes.LongLength > maxBytes) { return; }
            while (retainedBytes > maxBytes - bytes.LongLength && order.First is { } eldest)
            {
                order.RemoveFirst();
                entries.Remove(eldest.Value.Key);
                retainedBytes -= eldest.Value.Bytes.LongLength;
            }

            entries.Add(key, order.AddLast(new Entry(key, bytes)));
            retainedBytes += bytes.LongLength;
        }
    }
}
