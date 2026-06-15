namespace Lokad.OoxPdf.VisualDiff;

internal sealed record RegionMetric(
    int Page,
    string Region,
    int X,
    int Y,
    int Width,
    int Height,
    int PixelCount,
    double? MeanAbsoluteError,
    double? RootMeanSquaredError,
    double? ChangedPixelRatioAtThreshold16,
    double? ChangedPixelRatioAtThreshold32,
    double? StructuralSimilarity,
    double? ForegroundColorHistogramCorrelation,
    bool DimensionsMatch);
