namespace Lokad.OoxPdf;

/// <summary>
/// Output bounds and the resident page-content staging window. These are not
/// memory accounting or work quotas. Use host-supplied MemoryAdmission and
/// cooperative cancellation for shared-process admission and execution control.
/// Archive, XML and image bounds apply independently.
/// </summary>
public sealed class OoxConversionLimits
{
    /// <summary>Maximum pages per conversion (default 10000).</summary>
    public long MaxPagesPerConversion { get; init; } = 10_000;

    /// <summary>Maximum PDF content bytes per conversion (default 1073741824).</summary>
    public long MaxPdfContentBytesPerConversion { get; init; } = 1_073_741_824;

    /// <summary>Maximum output bytes per conversion (default 2147483648).</summary>
    public long MaxOutputBytesPerConversion { get; init; } = 2_147_483_648;

    /// <summary>Maximum serialized font bytes per conversion (default 268435456).</summary>
    public long MaxPdfFontBytesPerConversion { get; init; } = 268_435_456;

    /// <summary>Maximum serialized image bytes per conversion (default 536870912).</summary>
    public long MaxPdfImageBytesPerConversion { get; init; } = 536_870_912;

    /// <summary>
    /// Maximum page/group content retained in the staging store (default 64 MiB).
    /// Additional content spills to a delete-on-close file. This bounds the
    /// store's byte payload, not the complete conversion working set.
    /// </summary>
    public long MaxResidentPageContentBytesPerConversion { get; init; } = 67_108_864;

    internal void Validate()
    {
        if (MaxPagesPerConversion < 0) { throw new ArgumentOutOfRangeException(nameof(MaxPagesPerConversion), "Conversion limit must be non-negative."); }
        if (MaxPdfContentBytesPerConversion < 0) { throw new ArgumentOutOfRangeException(nameof(MaxPdfContentBytesPerConversion), "Conversion limit must be non-negative."); }
        if (MaxOutputBytesPerConversion < 0) { throw new ArgumentOutOfRangeException(nameof(MaxOutputBytesPerConversion), "Conversion limit must be non-negative."); }
        if (MaxPdfFontBytesPerConversion < 0) { throw new ArgumentOutOfRangeException(nameof(MaxPdfFontBytesPerConversion), "Conversion limit must be non-negative."); }
        if (MaxPdfImageBytesPerConversion < 0) { throw new ArgumentOutOfRangeException(nameof(MaxPdfImageBytesPerConversion), "Conversion limit must be non-negative."); }
        if (MaxResidentPageContentBytesPerConversion < 0) { throw new ArgumentOutOfRangeException(nameof(MaxResidentPageContentBytesPerConversion), "Conversion limit must be non-negative."); }
    }
}
