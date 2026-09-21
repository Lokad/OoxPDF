using Lokad.OoxPdf.Diagnostics;

namespace Lokad.OoxPdf;

/// <summary>
/// Conversion-wide cumulative work accounting (PLAN Q01). Per-site caps bound
/// individual allocations; this budget bounds their sum across one conversion so
/// many individually legal expansions cannot jointly exhaust a shared process.
/// </summary>
/// <remarks>
/// <para>Explicitly scoped ownership: <see cref="OoxPdfConverter"/> opens exactly
/// one scope per conversion and restores the prior scope on exit, so parallel
/// conversions never share totals (the AsyncLocal only isolates asynchronous
/// flows; nesting restores correctly). Expansion sites read
/// <see cref="Current"/> and skip charging when no conversion scope is active
/// (validation tools, inspection snapshots); per-site caps apply regardless.
/// </para>
/// <para>Single-flow by design: renderers are single-threaded and each conversion
/// holds its own budget, so charge sites need no synchronization; concurrent
/// conversions hold separate budgets, which scoping guarantees.
/// </para>
/// </remarks>
internal sealed class OoxConversionBudget
{
    private static readonly AsyncLocal<OoxConversionBudget?> Storage = new();

    /// <summary>
    /// The ambient conversion budget, or null outside a conversion scope.
    /// </summary>
    public static OoxConversionBudget? Current => Storage.Value;

    private readonly OoxConversionLimits limits;

    internal OoxConversionBudget(OoxConversionLimits? limits)
    {
        this.limits = limits ?? new OoxConversionLimits();
    }

    public long ChartRangeCells { get; private set; }

    public long TableFragments { get; private set; }

    public long ImagesDecoded { get; private set; }

    public long FontWork { get; private set; }

    public OoxConversionTotals Totals => new(ChartRangeCells, TableFragments, ImagesDecoded, FontWork);

    public static Scope BeginScope(OoxConversionLimits? limits)
    {
        var budget = new OoxConversionBudget(limits);
        OoxConversionBudget? prior = Storage.Value;
        Storage.Value = budget;
        return new Scope(budget, prior);
    }

    public void ChargeChartRangeCells(long count)
    {
        if (count < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(count));
        }

        if (count > limits.MaxChartRangeCellsPerConversion - ChartRangeCells)
        {
            throw new OoxPdfLimitExceededException(
                $"Conversion exceeds the chart range cell budget of {limits.MaxChartRangeCellsPerConversion} cells.");
        }

        ChartRangeCells += count;
    }

    public void ChargeTableFragments(long count)
    {
        if (count < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(count));
        }

        if (count > limits.MaxTableFragmentsPerConversion - TableFragments)
        {
            throw new OoxPdfLimitExceededException(
                $"Conversion exceeds the table fragment budget of {limits.MaxTableFragmentsPerConversion} fragments.");
        }

        TableFragments += count;
    }

    public void ChargeImagesDecoded(long count)
    {
        if (count < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(count));
        }

        if (count > limits.MaxImagesDecodedPerConversion - ImagesDecoded)
        {
            throw new OoxPdfLimitExceededException(
                $"Conversion exceeds the image decode budget of {limits.MaxImagesDecodedPerConversion} images.");
        }

        ImagesDecoded += count;
    }

    public void ChargeFontWork(long count)
    {
        if (count < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(count));
        }

        if (count > limits.MaxFontWorkPerConversion - FontWork)
        {
            throw new OoxPdfLimitExceededException(
                $"Conversion exceeds the font work budget of {limits.MaxFontWorkPerConversion} operations.");
        }

        FontWork += count;
    }

    internal sealed class Scope : IDisposable
    {
        private readonly OoxConversionBudget? prior;
        private bool disposed;

        internal Scope(OoxConversionBudget budget, OoxConversionBudget? prior)
        {
            Budget = budget;
            this.prior = prior;
        }

        public OoxConversionBudget Budget { get; }

        public void Dispose()
        {
            if (!disposed)
            {
                disposed = true;
                Storage.Value = prior;
            }
        }
    }
}

/// <summary>
/// Snapshot of a conversion's cumulative work counters, reported through the
/// informational CONVERSION_RESOURCE_SUMMARY diagnostic when
/// <see cref="OoxPdfOptions.ReportResourceUsage"/> is set.
/// </summary>
internal readonly record struct OoxConversionTotals(
    long ChartRangeCells,
    long TableFragments,
    long ImagesDecoded,
    long FontWork)
{
    public OoxPdfDiagnostic ToSummaryDiagnostic(int pageCount)
    {
        string message = FormattableString.Invariant(
            $"Conversion resource totals: pages={pageCount}; chartRangeCells={ChartRangeCells}; tableFragments={TableFragments}; imagesDecoded={ImagesDecoded}; fontWork={FontWork}.");
        return new OoxPdfDiagnostic(
            "CONVERSION_RESOURCE_SUMMARY",
            OoxPdfSeverity.Info,
            message,
            PartName: null,
            SlideIndex: null,
            PageIndex: null,
            Feature: null,
            Fallback: null);
    }
}
