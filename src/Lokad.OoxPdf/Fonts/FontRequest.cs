namespace Lokad.OoxPdf.Fonts;

public sealed record FontRequest(
    string FamilyName,
    bool Bold = false,
    bool Italic = false)
{
    public FontRequest(string familyName)
        : this(familyName, false, false)
    {
    }
}
