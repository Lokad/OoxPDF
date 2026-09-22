namespace Lokad.OoxPdf.Fonts;

/// <summary>
/// Provides the bytes of one font program behind an opaque stable identity.
/// </summary>
/// <remarks>
/// <para>
/// Font program ownership contract, relied upon by concurrent conversions:
/// </para>
/// <list type="bullet">
/// <item>Returned bytes are immutable after publication. Sources must return byte-identical
/// content for a StableId over their lifetime, and converters never write into returned
/// arrays (subsetters and repackagers always copy into fresh buffers).</item>
/// <item>Implementations must be safe for concurrent use across conversions sharing one
/// resolver. A source may cache and share a single backing array; callers must treat
/// the memory as read-only and must not dispose or pool it.</item>
/// <item>StableId is an opaque resolver-scoped identity, not a path: never normalize
/// its case and never compare it across resolver kinds.</item>
/// <item>The owning resolver (or caller for ad-hoc sources) defines retention.
/// Process-static snapshots retain their sources until explicitly invalidated;
/// see WindowsFontResolver.InvalidateDiscoveryCaches.</item>
/// </list>
/// </remarks>
public interface IFontProgramSource
{
    string StableId { get; }

    ValueTask<ReadOnlyMemory<byte>> GetBytesAsync(CancellationToken ct = default);
}
