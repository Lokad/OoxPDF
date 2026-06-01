using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxTableBorderGeometry
{
    private const double WordPdfBorderWidthScale = 0.96d;

    public static DocxTableCellBorder? Find(IReadOnlyList<DocxTableCellBorder> borders, string edge)
    {
        return borders.FirstOrDefault(border => string.Equals(border.Edge, edge, StringComparison.OrdinalIgnoreCase));
    }

    public static DocxTableCellBorder? SelectStronger(DocxTableCellBorder? first, DocxTableCellBorder? second)
    {
        if (first is null)
        {
            return second;
        }

        if (second is null)
        {
            return first;
        }

        return ResolveVisibleWidth(second) > ResolveVisibleWidth(first) ? second : first;
    }

    public static bool IsSuppressed(DocxTableCellBorder? border)
    {
        return border is not null &&
            (string.Equals(border.Value, "nil", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(border.Value, "none", StringComparison.OrdinalIgnoreCase));
    }

    public static double ResolveVisibleWidth(DocxTableCellBorder? border)
    {
        if (border is null || IsSuppressed(border))
        {
            return 0d;
        }

        double nominalWidth = int.TryParse(border.SizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int eighths)
            ? Math.Max(0.25d, eighths / 8d)
            : 0.75d;
        return nominalWidth * WordPdfBorderWidthScale;
    }
}
