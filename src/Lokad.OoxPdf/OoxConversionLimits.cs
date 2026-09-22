namespace Lokad.OoxPdf;

/// <summary>
/// Cumulative per-conversion work budgets. Each cap bounds the total
/// covered work of one kind across a whole conversion, complementing the per-site caps
/// (per chart union, per table row, per image) that bound individual allocations.
/// Covered work includes chart range cells plus dense slots, table fragments,
/// content images (including crop/recolor variants and effect rasters), font loads
/// and subsets, and the live image reservation peak (R01-R04). Aggregate XML DOMs,
/// nested packages, worksheet models, pages/resources/output bytes, and writer work
/// are not yet bounded (R04-R06); the live peak is a reservation peak, not total
/// memory (R19). Defaults are generous multiples of the per-site caps; hosts with
/// a claimed memory/work limit should set tighter values and watch
/// <see cref="OoxPdfOptions.ReportResourceUsage"/> output while tuning.
/// </summary>
public sealed class OoxConversionLimits
{
    /// <summary>
    /// Maximum chart workbook range cells expanded per conversion (default 2,000,000,
    /// twenty times the per-union cap). Counts every expanded area cell, including
    /// blank cells materialized during union expansion, plus dense slots materialized
    /// from cached/literal chart vectors (R04).
    /// </summary>
    public long MaxChartRangeCellsPerConversion { get; init; } = 2_000_000;

    /// <summary>
    /// Maximum DOCX table row fragments constructed per conversion (default 20,000,
    /// twenty times the per-row cap).
    /// </summary>
    public long MaxTableFragmentsPerConversion { get; init; } = 20_000;

    /// <summary>
    /// Maximum content images decoded per conversion (default 500), including PPTX
    /// crop/recolor variants and effect rasters through the shared decoder boundary
    /// (R03). Each image is still individually pixel-capped; cache hits do not recharge.
    /// This bounds how many individually legal images one conversion may carry.
    /// </summary>
    public long MaxImagesDecodedPerConversion { get; init; } = 500;

    /// <summary>
    /// Maximum font program loads and subset builds per conversion (default 5,000)
    /// through the shared loader/subsetter boundary, covering DOCX and PPTX (R03).
    /// Ordinary conversions resolve dozens; thousands indicate reference churn.
    /// Cache hits do not recharge.
    /// </summary>
    public long MaxFontWorkPerConversion { get; init; } = 5_000;

    /// <summary>
    /// Maximum live image decode bytes reserved per conversion at any one time
    /// (default 536,870,912, i.e. 512 MiB). This is a reservation peak using a
    /// width-by-height-by-4 estimate per pixel decode (R01 reserves before JPEG
    /// component-plane allocation); it is not total live or process memory (R19).
    /// Simultaneous buffers (PNG inflated bytes plus planes, JPEG planes plus RGB,
    /// compression scratch), retained pixels/variants, and non-image work are outside
    /// this reservation. <see cref="MaxImagesDecodedPerConversion"/> bounds how many
    /// individually legal images one conversion may carry. Reservations fail with
    /// <see cref="OoxPdfLimitExceededException"/> before the decode allocates.
    /// JPEG passthrough retains the already-owned input bytes and holds no reservation.
    /// </summary>
    public long MaxLiveImageBytesPerConversion { get; init; } = 536_870_912;

    internal void Validate()
    {
        if (MaxChartRangeCellsPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxChartRangeCellsPerConversion), "Conversion chart range cell budget must be non-negative.");
        }

        if (MaxTableFragmentsPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxTableFragmentsPerConversion), "Conversion table fragment budget must be non-negative.");
        }

        if (MaxImagesDecodedPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxImagesDecodedPerConversion), "Conversion image decode budget must be non-negative.");
        }

        if (MaxFontWorkPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxFontWorkPerConversion), "Conversion font work budget must be non-negative.");
        }

        if (MaxLiveImageBytesPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxLiveImageBytesPerConversion), "Conversion live image byte budget must be non-negative.");
        }
    }
}
