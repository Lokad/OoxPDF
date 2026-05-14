namespace Lokad.OoxPdf.VisualDiff;

internal sealed record PixelMetric(
    double MeanAbsoluteError,
    double RootMeanSquaredError,
    double ChangedPixelRatioAtThreshold16,
    double ChangedPixelRatioAtThreshold32);
