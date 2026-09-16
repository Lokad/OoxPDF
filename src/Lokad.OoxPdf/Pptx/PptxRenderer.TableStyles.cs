namespace Lokad.OoxPdf.Pptx;

internal static class PptxTableStyleResolver
{
    public static PptxSceneFillStyle ReadCellFill(PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme)
    {
        return ReadCellFill(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, PptxColorMap.Default);
    }

    public static PptxSceneFillStyle ReadCellFill(PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme, PptxColorMap colorMap)
    {
        double alpha = 1d;
        if (!tableStyle.IsSupported ||
            !theme.TryResolveColor(tableStyle.Accent, colorMap, out RgbColor accent))
        {
            return default;
        }

        int bodyColumnIndex = columnIndex - (tableStyle.FirstColumn ? 1 : 0);
        if (tableStyle.Kind == PptxBuiltInTableStyleKind.MediumStyle2 &&
            ((tableStyle.FirstRow && rowIndex == 0) ||
                (tableStyle.LastRow && rowIndex == rowCount - 1) ||
                (tableStyle.FirstColumn && columnIndex == 0) ||
                (tableStyle.LastColumn && columnIndex == columnCount - 1)))
        {
            return new PptxSceneFillStyle(true, accent, alpha);
        }

        if (tableStyle.Kind == PptxBuiltInTableStyleKind.LightStyle1 &&
            tableStyle.FirstRow &&
            rowIndex == 0)
        {
            return new PptxSceneFillStyle(true, accent, alpha);
        }

        if (tableStyle.Kind == PptxBuiltInTableStyleKind.DarkStyle1)
        {
            if (tableStyle.FirstRow && rowIndex == 0 && theme.TryResolveColor("dk1", colorMap, out RgbColor dark))
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
        if (tableStyle.Kind == PptxBuiltInTableStyleKind.LightStyle1)
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

        if (tableStyle.Kind == PptxBuiltInTableStyleKind.MediumStyle2)
        {
            bool banded = (tableStyle.BandRow && bodyRowIndex >= 0 && bodyRowIndex % 2 == 0) ||
                (tableStyle.BandColumn && bodyColumnIndex >= 0 && bodyColumnIndex % 2 == 0);
            // MediumStyle2 band fills are linear-light tints toward white (Office-calibrated
            // exact weights 0.60 banded / 0.80 unbanded; see PLAN.md).
            RgbColor color = banded
                ? LinearTintWhite(accent, 0.60d)
                : LinearTintWhite(accent, 0.80d);
            return new PptxSceneFillStyle(true, color, alpha);
        }

        if (tableStyle.Kind == PptxBuiltInTableStyleKind.DarkStyle1)
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
        return ReadCellTextStyle(tableStyle, rowIndex, columnIndex, rowCount, columnCount, theme, PptxColorMap.Default);
    }

    public static PptxSceneTableCellTextStyle ReadCellTextStyle(PptxSceneTableStyle tableStyle, int rowIndex, int columnIndex, int rowCount, int columnCount, PptxTheme theme, PptxColorMap colorMap)
    {
        bool bold = false;
        RgbColor? color = null;
        bool supportedStyle = tableStyle.IsSupported &&
            tableStyle.Kind is PptxBuiltInTableStyleKind.MediumStyle2 or
                PptxBuiltInTableStyleKind.LightStyle1 or
                PptxBuiltInTableStyleKind.DarkStyle1;
        bool firstRow = tableStyle.FirstRow && rowIndex == 0;
        bool firstCol = tableStyle.FirstColumn && columnIndex == 0;
        bool lastRow = tableStyle.LastRow &&
            rowIndex == rowCount - 1;
        bool lastCol = tableStyle.LastColumn &&
            columnCount > 0 &&
            columnIndex == columnCount - 1;
        if (supportedStyle &&
            firstRow &&
            theme.TryResolveColor("lt1", colorMap, out RgbColor firstRowColor))
        {
            color = firstRowColor;
            bold = true;
        }

        if (supportedStyle && (firstCol || lastRow || lastCol))
        {
            bold = true;
            if (tableStyle.Kind is PptxBuiltInTableStyleKind.MediumStyle2 or PptxBuiltInTableStyleKind.DarkStyle1 &&
                (lastRow || lastCol) &&
                theme.TryResolveColor("lt1", colorMap, out RgbColor conditionalColor))
            {
                color = conditionalColor;
            }
        }

        return new PptxSceneTableCellTextStyle(color, bold);
    }

    // Linear-light tint toward white (Office-calibrated table band recipe).
    private static RgbColor LinearTintWhite(RgbColor color, double weight)
    {
        return new RgbColor(
            LinearToSrgbByte(SrgbToLinear(color.Red) + (1d - SrgbToLinear(color.Red)) * weight),
            LinearToSrgbByte(SrgbToLinear(color.Green) + (1d - SrgbToLinear(color.Green)) * weight),
            LinearToSrgbByte(SrgbToLinear(color.Blue) + (1d - SrgbToLinear(color.Blue)) * weight));
    }

    private static double SrgbToLinear(byte channel)
    {
        double c = channel / 255d;
        return c <= 0.04045d ? c / 12.92d : System.Math.Pow((c + 0.055d) / 1.055d, 2.4d);
    }

    private static byte LinearToSrgbByte(double linear)
    {
        double clamped = System.Math.Clamp(linear, 0d, 1d);
        double c = clamped <= 0.0031308d ? clamped * 12.92d : 1.055d * System.Math.Pow(clamped, 1d / 2.4d) - 0.055d;
        return ToByte(c * 255d);
    }

    private static RgbColor ShadeColor(RgbColor color, double shade)
    {
        return new RgbColor(
            ToByte(color.Red * shade),
            ToByte(color.Green * shade),
            ToByte(color.Blue * shade));
    }

    private static RgbColor TintColor(RgbColor color, double tint)
    {
        return new RgbColor(
            ToByte(color.Red + (255d - color.Red) * tint),
            ToByte(color.Green + (255d - color.Green) * tint),
            ToByte(color.Blue + (255d - color.Blue) * tint));
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

}
