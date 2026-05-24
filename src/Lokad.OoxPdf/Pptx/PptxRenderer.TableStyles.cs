namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static bool TryReadBuiltInTableStyleCellFill(PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme, out RgbColor color, out double alpha)
    {
        alpha = 1d;
        if (!tableStyle.IsSupported ||
            !theme.TryResolveColor(tableStyle.Accent, out RgbColor accent))
        {
            color = default;
            return false;
        }

        int bodyColumnIndex = columnIndex - (tableStyle.FirstColumn ? 1 : 0);
        if (string.Equals(tableStyle.Name, "Medium-Style-2", StringComparison.Ordinal) &&
            ((tableStyle.FirstRow && rowIndex == 0) ||
                (tableStyle.LastRow && rowIndex == rowCount - 1) ||
                (tableStyle.FirstColumn && columnIndex == 0) ||
                (tableStyle.LastColumn && columnIndex == columnCount - 1)))
        {
            color = accent;
            return true;
        }

        if (string.Equals(tableStyle.Name, "Light-Style-1", StringComparison.Ordinal) &&
            tableStyle.FirstRow &&
            rowIndex == 0)
        {
            color = accent;
            return true;
        }

        if (string.Equals(tableStyle.Name, "Dark-Style-1", StringComparison.Ordinal))
        {
            if (tableStyle.FirstRow && rowIndex == 0 && theme.TryResolveColor("dk1", out RgbColor dark))
            {
                color = dark;
                return true;
            }

            if ((tableStyle.FirstColumn && columnIndex == 0) ||
                (tableStyle.LastColumn && columnIndex == columnCount - 1))
            {
                color = ShadeColor(accent, 0.6d);
                return true;
            }

            if (tableStyle.LastRow && rowIndex == rowCount - 1)
            {
                color = accent;
                return true;
            }
        }

        int bodyRowIndex = rowIndex - (tableStyle.FirstRow ? 1 : 0);
        if (string.Equals(tableStyle.Name, "Light-Style-1", StringComparison.Ordinal))
        {
            if (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0)
            {
                color = accent;
                alpha = 0.4d;
                return true;
            }

            if (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0)
            {
                color = accent;
                alpha = 0.4d;
                return true;
            }

            color = default;
            return false;
        }

        if (string.Equals(tableStyle.Name, "Medium-Style-2", StringComparison.Ordinal))
        {
            bool banded = (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0) ||
                (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0);
            color = banded
                ? TintColor(accent, 0.4d)
                : TintColor(accent, 0.2d);
            return true;
        }

        if (string.Equals(tableStyle.Name, "Dark-Style-1", StringComparison.Ordinal))
        {
            bool banded = (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0) ||
                (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0);
            color = banded
                ? ShadeColor(accent, 0.4d)
                : ShadeColor(accent, 0.2d);
            return true;
        }

        color = default;
        return false;
    }

    private static TableCellTextStyle ReadBuiltInTableStyleTextStyle(PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme)
    {
        bool bold = false;
        RgbColor? color = null;
        bool supportedStyle = tableStyle.IsSupported &&
            (string.Equals(tableStyle.Name, "Medium-Style-2", StringComparison.Ordinal) ||
                string.Equals(tableStyle.Name, "Light-Style-1", StringComparison.Ordinal) ||
                string.Equals(tableStyle.Name, "Dark-Style-1", StringComparison.Ordinal));
        bool firstRow = tableStyle.FirstRow && rowIndex == 0;
        bool firstCol = tableStyle.FirstColumn && columnIndex == 0;
        bool lastRow = tableStyle.LastRow &&
            rowIndex == rowCount - 1;
        bool lastCol = tableStyle.LastColumn &&
            columnCount > 0 &&
            columnIndex == columnCount - 1;
        if (supportedStyle &&
            firstRow &&
            theme.TryResolveColor("lt1", out RgbColor firstRowColor))
        {
            color = firstRowColor;
            bold = true;
        }

        if (supportedStyle && (firstCol || lastRow || lastCol))
        {
            bold = true;
            if ((string.Equals(tableStyle.Name, "Medium-Style-2", StringComparison.Ordinal) ||
                    string.Equals(tableStyle.Name, "Dark-Style-1", StringComparison.Ordinal)) &&
                (lastRow || lastCol) &&
                theme.TryResolveColor("lt1", out RgbColor conditionalColor))
            {
                color = conditionalColor;
            }
        }

        return new TableCellTextStyle(color, bold);
    }
    private static RgbColor TintColor(RgbColor color, double tint)
    {
        return new RgbColor(
            ToByte(color.Red + (255d - color.Red) * tint),
            ToByte(color.Green + (255d - color.Green) * tint),
            ToByte(color.Blue + (255d - color.Blue) * tint));
    }

    private static RgbColor ShadeColor(RgbColor color, double shade)
    {
        return new RgbColor(
            ToByte(color.Red * shade),
            ToByte(color.Green * shade),
            ToByte(color.Blue * shade));
    }

    private readonly record struct TableCellTextStyle(RgbColor? Color, bool Bold);
}
