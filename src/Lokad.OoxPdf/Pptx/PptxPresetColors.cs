namespace Lokad.OoxPdf.Pptx;

internal static class PptxPresetColors
{
    public static bool TryResolve(string? name, out RgbColor color)
    {
        color = default;
        if (name is null || name.Equals("rebeccapurple", StringComparison.OrdinalIgnoreCase))
            return false;

        if (RgbColor.TryResolveColorKeyword(name, out color))
            return true;

        // ECMA-376 Part 1, ST_PresetColorVal gives these two abbreviated
        // DrawingML values different RGB components from their long spellings.
        if (name.Equals("dkSeaGreen", StringComparison.OrdinalIgnoreCase))
        {
            color = new RgbColor(143, 188, 139);
            return true;
        }
        if (name.Equals("ltGoldenrodYellow", StringComparison.OrdinalIgnoreCase))
        {
            color = new RgbColor(250, 250, 120);
            return true;
        }

        if (name.StartsWith("dk", StringComparison.OrdinalIgnoreCase))
            return RgbColor.TryResolveColorKeyword(string.Concat("dark", name.AsSpan(2)), out color);
        if (name.StartsWith("lt", StringComparison.OrdinalIgnoreCase))
            return RgbColor.TryResolveColorKeyword(string.Concat("light", name.AsSpan(2)), out color);
        if (name.StartsWith("med", StringComparison.OrdinalIgnoreCase))
            return RgbColor.TryResolveColorKeyword(string.Concat("medium", name.AsSpan(3)), out color);
        return false;
    }
}
