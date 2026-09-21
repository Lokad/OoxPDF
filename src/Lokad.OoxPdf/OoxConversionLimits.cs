namespace Lokad.OoxPdf;

/// <summary>
/// Cumulative per-conversion work budgets (PLAN Q01). Each cap bounds the total
/// work of one kind across a whole conversion, complementing the per-site caps
/// (per chart union, per table row, per image) that bound individual allocations.
/// Defaults are generous multiples of the per-site caps; hosts with a claimed
/// memory/work limit should set tighter values and watch
/// <see cref="OoxPdfOptions.ReportResourceUsage"/> output while tuning.
/// </summary>
public sealed class OoxConversionLimits
{
    /// <summary>
    /// Maximum chart workbook range cells expanded per conversion (default 2,000,000,
    /// twenty times the per-union cap). Counts every expanded area cell, including
    /// blank cells materialized during union expansion.
    /// </summary>
    public long MaxChartRangeCellsPerConversion { get; init; } = 2_000_000;

    /// <summary>
    /// Maximum DOCX table row fragments constructed per conversion (default 20,000,
    /// twenty times the per-row cap).
    /// </summary>
    public long MaxTableFragmentsPerConversion { get; init; } = 20_000;

    /// <summary>
    /// Maximum content images decoded per conversion (default 500). Each image is
    /// still individually pixel-capped; this bounds how many individually legal
    /// images one conversion may carry.
    /// </summary>
    public long MaxImagesDecodedPerConversion { get; init; } = 500;

    /// <summary>
    /// Maximum font program loads and subset builds per conversion (default 5,000).
    /// Ordinary conversions resolve dozens; thousands indicate reference churn.
    /// </summary>
    public long MaxFontWorkPerConversion { get; init; } = 5_000;

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
    }
}
