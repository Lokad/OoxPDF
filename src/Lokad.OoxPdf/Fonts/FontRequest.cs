namespace Lokad.OoxPdf.Fonts;

public sealed record FontRequest(
    string FamilyName,
    bool Bold,
    bool Italic)
{
    public FontRequest(string familyName)
        : this(familyName, false, false)
    {
    }
}
