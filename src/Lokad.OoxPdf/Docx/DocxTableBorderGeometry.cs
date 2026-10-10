using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal static class DocxTableBorderGeometry
{
    private const double WordPdfBorderWidthScale = 0.96d;

    public static bool TryResolveReviewPaintScale(double printScale, out double matrixScale)
    {
        matrixScale = Math.Round(printScale, 3);
        return double.IsFinite(printScale) && printScale > 0d && printScale < 1d && matrixScale > 0d &&
            Math.Abs(matrixScale - printScale) / printScale <= 0.001d;
    }

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
        return ResolveNominalWidth(border) * WordPdfBorderWidthScale;
    }

    public static double ResolveNominalWidth(DocxTableCellBorder? border)
    {
        if (border is null || IsSuppressed(border))
        {
            return 0d;
        }

        double lineWidth = int.TryParse(border.SizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int eighths)
            ? Math.Max(0.25d, eighths / 8d)
            : 0.75d;
        // Word's double border has two authored-width strokes and an equal gap.
        return border.UseDoubleBandWidth && string.Equals(border.Value, "double", StringComparison.OrdinalIgnoreCase)
            ? 3d * lineWidth : lineWidth;
    }

    public static DocxTable ResolveUniformDoubleBodyTable(DocxTable table, double scale, double availableWidth)
    {
        if (scale != 1d || table.UseLegacyTableGrid || table.CellSpacingPoints is > 0d ||
            table.Revisions.Count != 0 || table.ColumnWidthsPoints.Sum() > availableWidth ||
            table.Rows.Any(row => row.Revisions.Count != 0 || row.Cells.Any(cell =>
                cell.Revisions.Count != 0 || cell.GridSpan != 1 || cell.HasVerticalMerge ||
                cell.BodyElements.Any(element => element is DocxTableElement)))) return table;
        var borders = table.Rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Borders)
            .Where(border => ResolveVisibleWidth(border) > 0d).ToArray();
        if (borders.Length == 0 || borders.Any(border =>
            !string.Equals(border.Value, "double", StringComparison.OrdinalIgnoreCase) ||
            !int.TryParse(border.SizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int eighths) || eighths <= 0) ||
            borders.Select(ResolveNominalWidth).Distinct().Count() != 1) return table;
        return table with { Rows = table.Rows.Select(row => row with {
            Cells = row.Cells.Select(cell => cell with {
                Borders = cell.Borders.Select(border => border with {
                    UseDoubleBandWidth = string.Equals(border.Value, "double", StringComparison.OrdinalIgnoreCase)
                }).ToArray()
            }).ToArray()
        }).ToArray() };
    }
}
