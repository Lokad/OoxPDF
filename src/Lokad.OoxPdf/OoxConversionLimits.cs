namespace Lokad.OoxPdf;

/// <summary>
/// Cumulative per-conversion work budgets. Each cap bounds the total
/// covered work of one kind across a whole conversion, complementing the per-site caps
/// (per chart union, per table row, per image) that bound individual allocations.
/// Covered work includes chart range cells plus dense slots, table fragments,
/// XML nodes, workbook cells, content images (including crop/recolor variants and
/// effect rasters), font loads and subsets, serialized pages/content/output bytes,
/// scene nodes, nested package bytes, workbook models, serialized font/image bytes,
/// and the live image reservation peak (R01-R06). Writer vector graphics (paths,
/// shadings, annotations) remain outside byte quotas; the live peak is a reservation
/// peak, not total
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
    /// Maximum XML nodes parsed per conversion (default 40,000,000, ten times the
    /// combined per-document object cap of 2,000,000 elements plus 2,000,000
    /// attributes). Counts elements plus attributes across every parsed part while
    /// cached parses do not recharge; each document still fails fast at its own
    /// caps before amplifying small input into a huge DOM (R04).
    /// </summary>
    public long MaxXmlNodesPerConversion { get; init; } = 40_000_000;

    /// <summary>
    /// Maximum chart workbook cells read per conversion (default 1,000,000, ten
    /// times the per-workbook cap). Counts cells across every embedded workbook
    /// while cached workbook models do not recharge (R04).
    /// </summary>
    public long MaxWorkbookCellsPerConversion { get; init; } = 1_000_000;

    /// <summary>
    /// Maximum PDF pages serialized per conversion (default 10,000). Counts every
    /// page once up front during serialization, bounding per-page writer overhead
    /// for repagination-heavy documents (R06).
    /// </summary>
    public long MaxPagesPerConversion { get; init; } = 10_000;

    /// <summary>
    /// Maximum PDF page-content bytes serialized per conversion (default
    /// 1,073,741,824, i.e. 1 GiB). Counts encoded content-stream bytes across all
    /// pages once up front during serialization (R06).
    /// </summary>
    public long MaxPdfContentBytesPerConversion { get; init; } = 1073741824;

    /// <summary>
    /// Maximum PDF output bytes published per conversion (default 2,147,483,648,
    /// i.e. 2 GiB). Charged from the measured serialized size after writing, while
    /// the conversion scope is still open (R06).
    /// </summary>
    public long MaxOutputBytesPerConversion { get; init; } = 2147483648;

    /// <summary>
    /// Maximum PDF font bytes serialized per conversion (default 268,435,456, i.e.
    /// 256 MiB, four times the per-file program cap). Counts embedded font programs
    /// plus ToUnicode maps as they serialize, complementing the font-work operation
    /// count (R06).
    /// </summary>
    public long MaxPdfFontBytesPerConversion { get; init; } = 268435456;

    /// <summary>
    /// Maximum PDF image bytes serialized per conversion (default 536,870,912, i.e.
    /// 512 MiB). Counts encoded image and soft-mask bytes as they serialize,
    /// complementing the image-decode count (R06).
    /// </summary>
    public long MaxPdfImageBytesPerConversion { get; init; } = 536870912;

    /// <summary>
    /// Maximum PPTX scene nodes built per conversion (default 2,000,000). Counts
    /// shapes across slides, masters, and layouts as they materialize, including
    /// nested group children; cached DOMs re-walked per slide recharge (R04).
    /// </summary>
    public long MaxSceneNodesPerConversion { get; init; } = 2_000_000;

    /// <summary>
    /// Maximum retained bytes across nested packages per conversion (default
    /// 536,870,912, i.e. 512 MiB, twice the per-package total). Counts embedded
    /// chart workbook packages, each already capped individually (R04).
    /// </summary>
    public long MaxNestedPackageBytesPerConversion { get; init; } = 536_870_912;

    /// <summary>
    /// Maximum chart workbook models read per conversion (default 100). Each model
    /// is already cell-capped; cached models do not recharge (R04).
    /// </summary>
    public long MaxWorkbookModelsPerConversion { get; init; } = 100;

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
    /// (default 536,870,912, i.e. 512 MiB). This is a reservation peak using
    /// per-format working-set estimates held by decoded-pixel ownership across
    /// decode/transform/compress work (R02): PNG reserves inflated bytes plus
    /// output planes, JPEG reserves sample planes plus RGB (R01 reserves before
    /// plane allocation), BMP reserves the output plane, and crop/recolor/effect
    /// transients reserve alongside the source pixels. It is not total live or
    /// process memory (R19). Pixels retained outside any live operation scope
    /// and non-image work are outside this reservation. <see cref="MaxImagesDecodedPerConversion"/> bounds how many
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

        if (MaxXmlNodesPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxXmlNodesPerConversion), "Conversion XML node budget must be non-negative.");
        }

        if (MaxWorkbookCellsPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxWorkbookCellsPerConversion), "Conversion workbook cell budget must be non-negative.");
        }

        if (MaxPagesPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxPagesPerConversion), "Conversion page budget must be non-negative.");
        }

        if (MaxPdfContentBytesPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxPdfContentBytesPerConversion), "Conversion PDF content byte budget must be non-negative.");
        }

        if (MaxOutputBytesPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxOutputBytesPerConversion), "Conversion output byte budget must be non-negative.");
        }

        if (MaxPdfFontBytesPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxPdfFontBytesPerConversion), "Conversion PDF font byte budget must be non-negative.");
        }

        if (MaxPdfImageBytesPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxPdfImageBytesPerConversion), "Conversion PDF image byte budget must be non-negative.");
        }

        if (MaxSceneNodesPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxSceneNodesPerConversion), "Conversion scene node budget must be non-negative.");
        }

        if (MaxNestedPackageBytesPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxNestedPackageBytesPerConversion), "Conversion nested package byte budget must be non-negative.");
        }

        if (MaxWorkbookModelsPerConversion < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(MaxWorkbookModelsPerConversion), "Conversion workbook model budget must be non-negative.");
        }
    }
}
