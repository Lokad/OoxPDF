namespace Lokad.OoxPdf.VisualDiff;

internal sealed record PageMetric(
    int Page,
    string? ReferenceFile,
    string? CandidateFile,
    int? ReferenceWidth,
    int? ReferenceHeight,
    int? CandidateWidth,
    int? CandidateHeight,
    double? MeanAbsoluteError,
    double? RootMeanSquaredError,
    double? ChangedPixelRatioAtThreshold16,
    double? ChangedPixelRatioAtThreshold32,
    double? StructuralSimilarity,
    double? ForegroundColorHistogramCorrelation,
    bool DimensionsMatch);
