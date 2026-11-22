using System.Globalization;

namespace Lokad.OoxPdf.Docx;

internal enum DocxTableBorderKind
{
    Single,
    Thick,
    Double,
    Triple,
    Dotted,
    Dashed,
    DashDotStroked,
    Compound,
    ThreeDimensional,
    Wave,
    InsetOrOutset,
    Suppressed,
    Unsupported,
    Count
}

internal static class DocxTableBorderGeometry
{
    private const double WordPdfBorderWidthScale = 0.96d;

    private static readonly Dictionary<string, DocxTableBorderKind> BorderKinds = new(StringComparer.OrdinalIgnoreCase)
    {
        ["single"] = DocxTableBorderKind.Single,
        ["thick"] = DocxTableBorderKind.Thick,
        ["double"] = DocxTableBorderKind.Double,
        ["triple"] = DocxTableBorderKind.Triple,
        ["dotted"] = DocxTableBorderKind.Dotted,
        ["dashed"] = DocxTableBorderKind.Dashed,
        ["dashSmallGap"] = DocxTableBorderKind.Dashed,
        ["dotDash"] = DocxTableBorderKind.Dashed,
        ["dotDotDash"] = DocxTableBorderKind.Dashed,
        ["dashDotStroked"] = DocxTableBorderKind.DashDotStroked,
        ["thinThickSmallGap"] = DocxTableBorderKind.Compound,
        ["thickThinSmallGap"] = DocxTableBorderKind.Compound,
        ["thinThickThinSmallGap"] = DocxTableBorderKind.Compound,
        ["thinThickMediumGap"] = DocxTableBorderKind.Compound,
        ["thickThinMediumGap"] = DocxTableBorderKind.Compound,
        ["thinThickThinMediumGap"] = DocxTableBorderKind.Compound,
        ["thinThickLargeGap"] = DocxTableBorderKind.Compound,
        ["thickThinLargeGap"] = DocxTableBorderKind.Compound,
        ["thinThickThinLargeGap"] = DocxTableBorderKind.Compound,
        ["threeDEmboss"] = DocxTableBorderKind.ThreeDimensional,
        ["threeDEngrave"] = DocxTableBorderKind.ThreeDimensional,
        ["wave"] = DocxTableBorderKind.Wave,
        ["doubleWave"] = DocxTableBorderKind.Wave,
        ["outset"] = DocxTableBorderKind.InsetOrOutset,
        ["inset"] = DocxTableBorderKind.InsetOrOutset,
        ["nil"] = DocxTableBorderKind.Suppressed,
        ["none"] = DocxTableBorderKind.Suppressed
    };

    public static DocxTableBorderKind Classify(string? value) => value is null
        ? DocxTableBorderKind.Single
        : BorderKinds.GetValueOrDefault(value, DocxTableBorderKind.Unsupported);

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
        return Classify(border?.Value) == DocxTableBorderKind.Suppressed;
    }

    public static double ResolveVisibleWidth(DocxTableCellBorder? border)
    {
        return ResolveNominalWidth(border) * WordPdfBorderWidthScale;
    }

    public static double ResolveNominalWidth(DocxTableCellBorder? border)
    {
        if (border is null)
        {
            return 0d;
        }

        DocxTableBorderKind kind = Classify(border.Value);
        if (kind == DocxTableBorderKind.Suppressed)
        {
            return 0d;
        }

        double lineWidth = int.TryParse(border.SizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int eighths)
            ? Math.Max(0.25d, eighths / 8d)
            : 0.75d;
        // Word's double border has two authored-width strokes and an equal gap.
        return border.UseDoubleBandWidth && kind == DocxTableBorderKind.Double
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
