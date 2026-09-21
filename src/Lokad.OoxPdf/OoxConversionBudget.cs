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
/// <para>Live reservations (image decode scratch plus pixel planes) are tracked
/// separately from cumulative work: <see cref="ReserveLiveImageBytes"/> fails before
/// allocation when the live level would exceed its cap and releases on dispose, while
/// <see cref="PeakLiveImageBytes"/> records the high-water mark reported in the
/// CONVERSION_RESOURCE_SUMMARY diagnostic.
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

    /// <summary>
    /// Currently reserved live image decode bytes. Returns to zero when every
    /// reservation is disposed; single-threaded renderers hold at most one.
    /// </summary>
    public long LiveImageBytes { get; private set; }

    /// <summary>
    /// High-water mark of <see cref="LiveImageBytes"/> for this conversion scope.
    /// </summary>
    public long PeakLiveImageBytes { get; private set; }

    public OoxConversionTotals Totals => new(ChartRangeCells, TableFragments, ImagesDecoded, FontWork, PeakLiveImageBytes);

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

    /// <summary>
    /// Reserves live image decode bytes before the decode allocates, tracking the
    /// conversion peak. The returned reservation releases exactly once on dispose;
    /// callers hold it (typically via <c>using</c>) across the allocating work so the
    /// peak reflects attempted work while the current level always returns to
    /// baseline, including on failures. Outside a conversion scope image readers skip
    /// reservation entirely and per-image pixel caps still apply.
    /// </summary>
    public LiveReservation ReserveLiveImageBytes(long byteCount)
    {
        if (byteCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(byteCount));
        }

        if (byteCount > limits.MaxLiveImageBytesPerConversion - LiveImageBytes)
        {
            throw new OoxPdfLimitExceededException(
                $"Conversion exceeds the live image byte budget of {limits.MaxLiveImageBytesPerConversion} bytes.");
        }

        LiveImageBytes += byteCount;
        if (LiveImageBytes > PeakLiveImageBytes)
        {
            PeakLiveImageBytes = LiveImageBytes;
        }

        return new LiveReservation(this, byteCount);
    }

    private void ReleaseLiveImageBytes(long byteCount)
    {
        LiveImageBytes -= byteCount;
        if (LiveImageBytes < 0)
        {
            throw new InvalidOperationException("Unbalanced live image byte release.");
        }
    }

    /// <summary>
    /// An outstanding live image byte reservation. Disposing releases exactly once;
    /// double dispose is a no-op.
    /// </summary>
    internal sealed class LiveReservation : IDisposable
    {
        private OoxConversionBudget? budget;
        private readonly long byteCount;

        internal LiveReservation(OoxConversionBudget budget, long byteCount)
        {
            this.budget = budget;
            this.byteCount = byteCount;
        }

        public void Dispose()
        {
            if (budget is not null)
            {
                OoxConversionBudget owner = budget;
                budget = null;
                owner.ReleaseLiveImageBytes(byteCount);
            }
        }
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
/// Snapshot of a conversion's cumulative work counters plus the peak live image
/// reservation, reported through the informational CONVERSION_RESOURCE_SUMMARY
/// diagnostic when <see cref="OoxPdfOptions.ReportResourceUsage"/> is set.
/// </summary>
internal readonly record struct OoxConversionTotals(
    long ChartRangeCells,
    long TableFragments,
    long ImagesDecoded,
    long FontWork,
    long PeakLiveImageBytes)
{
    public OoxPdfDiagnostic ToSummaryDiagnostic(int pageCount)
    {
        string message = FormattableString.Invariant(
            $"Conversion resource totals: pages={pageCount}; chartRangeCells={ChartRangeCells}; tableFragments={TableFragments}; imagesDecoded={ImagesDecoded}; fontWork={FontWork}; peakLiveImageBytes={PeakLiveImageBytes}.");
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
