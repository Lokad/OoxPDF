using System.Collections;

namespace Lokad.OoxPdf.Fonts;

internal sealed partial class OpenTypeFont
{
    private sealed class KerningPairWriter(int sourceCount, CancellationToken token)
    {
        private readonly uint[] keys = sourceCount == 0 ? [] : new uint[sourceCount];
        private readonly short[] values = sourceCount == 0 ? [] : new short[sourceCount];
        private int written;
        public void Add(uint key, short value)
        {
            if (written >= sourceCount) throw new InvalidDataException("Kerning source contribution count changed during parsing.");
            if ((written & 1023) == 0) token.ThrowIfCancellationRequested();
            keys[written] = key;
            values[written] = value;
            written++;
        }
        public IReadOnlyDictionary<uint, short> Complete()
        {
            token.ThrowIfCancellationRequested();
            if (written != sourceCount) throw new InvalidDataException("Kerning source contribution count changed during parsing.");
            SortStable();
            int count = 0, index = 0;
            while (index < sourceCount)
            {
                uint key = keys[index];
                short latest = values[index++];
                while (index < sourceCount && keys[index] == key)
                {
                    if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
                    latest = values[index++];
                }
                if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
                keys[count] = key;
                values[count++] = latest;
            }
            token.ThrowIfCancellationRequested();
            return new KerningPairMap(keys, values, count);
        }

        private void SortStable()
        {
            if (sourceCount < 2) return;
            // Fixed-width glyph-pair keys admit a stable bounded radix sort.
            // Stability keeps the original last-wins order of equal keys.
            // Small maps use byte buckets to avoid a large fixed histogram.
            int bits = sourceCount < 65536 ? 8 : 16;
            int[] offsets = new int[1 << bits];
            uint mask = (uint)offsets.Length - 1;
            uint[] scratchKeys = new uint[sourceCount];
            short[] scratchValues = new short[sourceCount];
            uint[] sourceKeys = keys, targetKeys = scratchKeys;
            short[] sourceValues = values, targetValues = scratchValues;
            for (int shift = 0; shift < 32; shift += bits)
            {
                Array.Clear(offsets);
                for (int index = 0; index < sourceCount; index++)
                {
                    if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
                    offsets[(int)((sourceKeys[index] >> shift) & mask)]++;
                }
                int next = 0;
                for (int bucket = 0; bucket < offsets.Length; bucket++)
                {
                    if ((bucket & 1023) == 0) token.ThrowIfCancellationRequested();
                    int length = offsets[bucket]; offsets[bucket] = next; next += length;
                }
                for (int index = 0; index < sourceCount; index++)
                {
                    if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
                    int target = offsets[(int)((sourceKeys[index] >> shift) & mask)]++;
                    targetKeys[target] = sourceKeys[index];
                    targetValues[target] = sourceValues[index];
                }
                (sourceKeys, targetKeys) = (targetKeys, sourceKeys);
                (sourceValues, targetValues) = (targetValues, sourceValues);
            }
            // Both supported digit widths take an even number of passes, so
            // the complete result is back in the builder's owned arrays.
            token.ThrowIfCancellationRequested();
        }
    }

    // The builder transfers private arrays only after sorting/compaction. Font
    // readers share an immutable map; no lazy mutation or synchronization owner.
    private sealed class KerningPairMap(uint[] keys, short[] values, int count) : IReadOnlyDictionary<uint, short>
    {
        public int Count => count;

        public short this[uint key] => TryGetValue(key, out short value) ? value : throw new KeyNotFoundException();

        public bool ContainsKey(uint key) => Array.BinarySearch(keys, 0, count, key) >= 0;

        public bool TryGetValue(uint key, out short value)
        {
            int index = Array.BinarySearch(keys, 0, count, key);
            value = index >= 0 ? values[index] : (short)0;
            return index >= 0;
        }

        public IEnumerable<uint> Keys
        {
            get { for (int index = 0; index < count; index++) yield return keys[index]; }
        }

        public IEnumerable<short> Values
        {
            get { for (int index = 0; index < count; index++) yield return values[index]; }
        }

        public IEnumerator<KeyValuePair<uint, short>> GetEnumerator()
        {
            for (int index = 0; index < count; index++) yield return new(keys[index], values[index]);
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
