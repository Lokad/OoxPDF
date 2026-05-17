namespace Lokad.OoxPdf.Fonts;

public sealed record FontResolution(
    string FamilyName,
    string? FontFilePath,
    bool IsFallback,
    bool Bold = false,
    bool Italic = false,
    int WeightClass = 400,
    int FontFaceIndex = 0,
    bool HasMathTable = false);
