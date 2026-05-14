namespace Lokad.OoxPdf.Fonts;

public sealed record FontResolution(
    string FamilyName,
    string? FontFilePath,
    bool IsFallback);
