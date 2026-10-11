namespace Lokad.OoxPdf.Docx;

internal sealed partial class DocxLayoutEngine
{
    private static bool IsEligibleDeclaredTopContentTable(
        DocxTable table, IDocxTextMeasurer? measurer, double fixedScale)
    {
        if (fixedScale != 1d || table.Rows.Count is < 1 or > 3 ||
            table.ColumnWidthsPoints.Count is < 1 or > 2)
            return false;
        // Mixed top rows reuse the qualified declared and automatic geometry
        // only in a complete two/three-row grid with one or two columns.
        bool mixedRows = table.Rows.Count is 2 or 3 && table.ColumnWidthsPoints.Count is 1 or 2 &&
            table.Rows.Any(row => HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue)) &&
            table.Rows.Any(row => row.HeightPoints is > 0d &&
                (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase)));
        return table.Rows.All(row =>
            row.Cells.Count == table.ColumnWidthsPoints.Count &&
            (mixedRows && HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue) ||
            row.HeightPoints is > 0d && double.IsFinite(row.HeightPoints.Value) &&
            (string.Equals(row.HeightRuleValue, "exact", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(row.HeightRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase))) &&
            row.Cells.All(cell => cell.VerticalAlignment == DocxTableCellVerticalAlignment.Top &&
                (cell.VerticalAlignmentValue is null ||
                    string.Equals(cell.VerticalAlignmentValue, "top", StringComparison.OrdinalIgnoreCase)) &&
                IsEligibleTopContentCellBorders(cell, limitWidth: table.Rows.Count == 3) &&
                IsEligibleTableMinimumCell(table, row, cell, measurer) &&
                cell.Paragraphs.All(p => p.Runs.Count != 0) &&
                cell.Paragraphs.Any(p => p.EffectiveProperties.LineSpacingPoints is > 0d &&
                    string.Equals(p.EffectiveProperties.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase))));
    }

    private static bool IsEligibleAutomaticTopContentTable(
        DocxTable table, IDocxTextMeasurer? measurer, double fixedScale)
    {
        if (fixedScale != 1d || table.Rows.Count is < 1 or > 3 ||
            table.ColumnWidthsPoints.Count is < 1 or > 2)
            return false;
        if (table.Rows.Any(row => HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue)) &&
            IsEligibleDeclaredTopContentTable(table, measurer, fixedScale))
            return true;
        // Top masks have independent admission. They never enable the
        // centered/bottom paragraph-slot alignment calculation.
        return table.Rows.All(row =>
            row.Cells.Count == table.ColumnWidthsPoints.Count &&
            HasAutomaticTableRowHeight(row.HeightPoints, row.HeightRuleValue) &&
            row.Cells.All(cell => cell.VerticalAlignment == DocxTableCellVerticalAlignment.Top &&
                (cell.VerticalAlignmentValue is null ||
                    string.Equals(cell.VerticalAlignmentValue, "top", StringComparison.OrdinalIgnoreCase)) &&
                (table.Rows.Count <= 2 || IsEligibleTopContentCellBorders(cell, limitWidth: true)) &&
                IsEligibleTableMinimumCell(table, row, cell, measurer) &&
                cell.Paragraphs.All(p => p.Runs.Count != 0) &&
                cell.Paragraphs.Any(p => p.EffectiveProperties.LineSpacingPoints is > 0d &&
                    string.Equals(p.EffectiveProperties.Spacing.LineRuleValue, "atLeast", StringComparison.OrdinalIgnoreCase))));
    }

    private static bool IsEligibleTopContentCellBorders(DocxTableCell cell, bool limitWidth) =>
        cell.Borders.All(border => DocxTableBorderGeometry.ResolveVisibleWidth(border) <= 0d ||
            // Wider bands accumulate unqualified error over three rows.
            (!limitWidth || DocxTableBorderGeometry.ResolveVisibleWidth(border) <=
                (border.UseDoubleBandWidth ? 2.88d : .96d)) &&
            (string.Equals(border.Value, "single", StringComparison.OrdinalIgnoreCase) ||
                border.UseDoubleBandWidth && string.Equals(border.Value, "double", StringComparison.OrdinalIgnoreCase)));

}
