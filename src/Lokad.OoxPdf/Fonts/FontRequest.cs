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

/// <summary>
/// Matches font requests the way resolvers do: family names compare
/// case-insensitively while bold/italic compare exactly. Used for
/// per-conversion resolution caches (G03).
/// </summary>
internal sealed class FontRequestKeyComparer : IEqualityComparer<FontRequest>
{
    public static readonly FontRequestKeyComparer OrdinalIgnoreCaseFamily = new();

    public bool Equals(FontRequest? left, FontRequest? right)
    {
        return left is not null &&
            right is not null &&
            left.Bold == right.Bold &&
            left.Italic == right.Italic &&
            string.Equals(left.FamilyName, right.FamilyName, StringComparison.OrdinalIgnoreCase);
    }

    public int GetHashCode(FontRequest request)
    {
        return HashCode.Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(request.FamilyName), request.Bold, request.Italic);
    }
}
