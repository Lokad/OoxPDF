namespace Lokad.OoxPdf.Fonts;

public sealed record FontRequest(
    string FamilyName,
    bool Bold = false,
    bool Italic = false);
