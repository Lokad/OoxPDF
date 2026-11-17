using System.Globalization;

namespace Lokad.OoxPdf.Pptx;

// Axis IDs are unsigned integers. Keep authored spellings in the scene and
// inspector, but use the same identity for lookup, pairing and deduplication.
// Malformed IDs retain the existing ordinal recovery, including exact matches.
internal sealed class PptxChartAxisIdComparer : IEqualityComparer<string>
{
    public static PptxChartAxisIdComparer Instance { get; } = new();

    private PptxChartAxisIdComparer() { }

    public bool Equals(string? left, string? right) =>
        string.Equals(left, right, StringComparison.Ordinal) ||
        TryParse(left, out uint leftId) && TryParse(right, out uint rightId) && leftId == rightId;

    public int GetHashCode(string value) => TryParse(value, out uint id)
        ? id.GetHashCode() : StringComparer.Ordinal.GetHashCode(value);

    private static bool TryParse(string? value, out uint id) =>
        uint.TryParse(value.AsSpan().Trim(" \t\r\n".AsSpan()), NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture, out id);
}
