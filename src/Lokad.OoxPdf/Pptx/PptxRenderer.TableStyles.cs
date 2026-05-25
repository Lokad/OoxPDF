namespace Lokad.OoxPdf.Pptx;

internal static class PptxTableStyleResolver
{
    public static PptxSceneFillStyle ReadCellFill(PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme)
    {
        double alpha = 1d;
        if (!tableStyle.IsSupported ||
            !theme.TryResolveColor(tableStyle.Accent, out RgbColor accent))
        {
            return default;
        }

        int bodyColumnIndex = columnIndex - (tableStyle.FirstColumn ? 1 : 0);
        if (string.Equals(tableStyle.Name, "Medium-Style-2", StringComparison.Ordinal) &&
            ((tableStyle.FirstRow && rowIndex == 0) ||
                (tableStyle.LastRow && rowIndex == rowCount - 1) ||
                (tableStyle.FirstColumn && columnIndex == 0) ||
                (tableStyle.LastColumn && columnIndex == columnCount - 1)))
        {
            return new PptxSceneFillStyle(true, accent, alpha);
        }

        if (string.Equals(tableStyle.Name, "Light-Style-1", StringComparison.Ordinal) &&
            tableStyle.FirstRow &&
            rowIndex == 0)
        {
            return new PptxSceneFillStyle(true, accent, alpha);
        }

        if (string.Equals(tableStyle.Name, "Dark-Style-1", StringComparison.Ordinal))
        {
            if (tableStyle.FirstRow && rowIndex == 0 && theme.TryResolveColor("dk1", out RgbColor dark))
            {
                return new PptxSceneFillStyle(true, dark, alpha);
            }

            if ((tableStyle.FirstColumn && columnIndex == 0) ||
                (tableStyle.LastColumn && columnIndex == columnCount - 1))
            {
                return new PptxSceneFillStyle(true, ShadeColor(accent, 0.6d), alpha);
            }

            if (tableStyle.LastRow && rowIndex == rowCount - 1)
            {
                return new PptxSceneFillStyle(true, accent, alpha);
            }
        }

        int bodyRowIndex = rowIndex - (tableStyle.FirstRow ? 1 : 0);
        if (string.Equals(tableStyle.Name, "Light-Style-1", StringComparison.Ordinal))
        {
            if (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0)
            {
                return new PptxSceneFillStyle(true, accent, 0.4d);
            }

            if (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0)
            {
                return new PptxSceneFillStyle(true, accent, 0.4d);
            }

            return default;
        }

        if (string.Equals(tableStyle.Name, "Medium-Style-2", StringComparison.Ordinal))
        {
            bool banded = (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0) ||
                (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0);
            RgbColor color = banded
                ? TintOfficeTableAccent(accent, 0.71d, 0.75d)
                : TintOfficeTableAccent(accent, 0.86d, 0.75d);
            return new PptxSceneFillStyle(true, color, alpha);
        }

        if (string.Equals(tableStyle.Name, "Dark-Style-1", StringComparison.Ordinal))
        {
            bool banded = (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0) ||
                (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0);
            RgbColor color = banded
                ? ShadeColor(accent, 0.4d)
                : ShadeColor(accent, 0.2d);
            return new PptxSceneFillStyle(true, color, alpha);
        }

        return default;
    }

    public static PptxSceneTableCellTextStyle ReadCellTextStyle(PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme)
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

        return new PptxSceneTableCellTextStyle(color, bold);
    }

    private static RgbColor TintColor(RgbColor color, double tint)
    {
        return new RgbColor(
            ToByte(color.Red + (255d - color.Red) * tint),
            ToByte(color.Green + (255d - color.Green) * tint),
            ToByte(color.Blue + (255d - color.Blue) * tint));
    }

    private static RgbColor TintOfficeTableAccent(RgbColor color, double tint, double saturationFactor)
    {
        (double hue, double saturation, double luminance) = ToHsl(color);
        double tintedLuminance = luminance + (1d - luminance) * tint;
        return FromHsl(hue, saturation * saturationFactor, tintedLuminance);
    }

    private static RgbColor ShadeColor(RgbColor color, double shade)
    {
        return new RgbColor(
            ToByte(color.Red * shade),
            ToByte(color.Green * shade),
            ToByte(color.Blue * shade));
    }

    private static byte ToByte(double value)
    {
        return (byte)Math.Clamp((int)Math.Round(value, MidpointRounding.AwayFromZero), 0, 255);
    }

    private static (double Hue, double Saturation, double Luminance) ToHsl(RgbColor color)
    {
        double red = color.Red / 255d;
        double green = color.Green / 255d;
        double blue = color.Blue / 255d;
        double max = Math.Max(red, Math.Max(green, blue));
        double min = Math.Min(red, Math.Min(green, blue));
        double delta = max - min;
        double luminance = (max + min) / 2d;
        double saturation = delta == 0d ? 0d : delta / (1d - Math.Abs(2d * luminance - 1d));
        double hue = 0d;
        if (delta != 0d)
        {
            if (max == red)
            {
                hue = 60d * (((green - blue) / delta) % 6d);
            }
            else if (max == green)
            {
                hue = 60d * ((blue - red) / delta + 2d);
            }
            else
            {
                hue = 60d * ((red - green) / delta + 4d);
            }

            if (hue < 0d)
            {
                hue += 360d;
            }
        }

        return (hue, saturation, luminance);
    }

    private static RgbColor FromHsl(double hue, double saturation, double luminance)
    {
        saturation = Math.Clamp(saturation, 0d, 1d);
        luminance = Math.Clamp(luminance, 0d, 1d);
        double chroma = (1d - Math.Abs(2d * luminance - 1d)) * saturation;
        double segment = hue / 60d;
        double second = chroma * (1d - Math.Abs(segment % 2d - 1d));
        (double red, double green, double blue) = segment switch
        {
            >= 0d and < 1d => (chroma, second, 0d),
            >= 1d and < 2d => (second, chroma, 0d),
            >= 2d and < 3d => (0d, chroma, second),
            >= 3d and < 4d => (0d, second, chroma),
            >= 4d and < 5d => (second, 0d, chroma),
            _ => (chroma, 0d, second)
        };
        double match = luminance - chroma / 2d;
        return new RgbColor(ToByte((red + match) * 255d), ToByte((green + match) * 255d), ToByte((blue + match) * 255d));
    }
}
