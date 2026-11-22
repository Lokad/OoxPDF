using System.Collections.Concurrent;
using Lokad.OoxPdf.Fonts;

namespace Lokad.OoxPdf.Tests;

internal static class TestFontPrograms
{
    // Inspectors that need bytes after acquisition keep an independent test copy.
    internal static async ValueTask<ReadOnlyMemory<byte>> CopyAsync(IFontProgramSource source, CancellationToken token = default)
    {
        using var lease = await source.AcquireAsync(null, token).ConfigureAwait(false);
        return lease.Bytes.ToArray();
    }

    internal static IOoxPdfFontProgramLease Borrow(ReadOnlyMemory<byte> bytes)
    {
        using var buffer = new FontProgramBuffer(bytes, null);
        return buffer.Acquire();
    }

    internal sealed class PerFamilyCountingResolver(ReadOnlyMemory<byte> fontBytes) : IFontResolver
    {
        private readonly ConcurrentDictionary<string, int> acquisitions = new(StringComparer.OrdinalIgnoreCase);

        internal IReadOnlyDictionary<string, int> Acquisitions => acquisitions;

        public FontFaceResolution Resolve(FontRequest request)
        {
            string family = string.IsNullOrWhiteSpace(request.FamilyName) ? "(default)" : request.FamilyName;
            var source = new CountingSource("counting:" + family, fontBytes, acquisitions);
            return new FontFaceResolution(request.FamilyName, family,
                new FontStyleKey(request.Bold, request.Italic, request.Bold ? 700 : 400, 0, false),
                source, IsFallback: false);
        }

        private sealed class CountingSource(string stableId, ReadOnlyMemory<byte> bytes,
            ConcurrentDictionary<string, int> acquisitions) : IFontProgramSource
        {
            public string StableId => stableId;

            public ValueTask<IOoxPdfFontProgramLease> AcquireAsync(IMemoryBudget? operationBudget, CancellationToken ct = default)
            {
                acquisitions.AddOrUpdate(stableId, 1, (_, count) => count + 1);
                ct.ThrowIfCancellationRequested();
                return ValueTask.FromResult(Borrow(bytes));
            }
        }
    }
}
