using Lokad.OoxPdf.Pptx;

namespace Lokad.OoxPdf.Docx;

internal static class DocxTextColor
{
    // Null retains automatic color; malformed literal values retain the black fallback.
    public static RgbColor? Read(string? value) =>
        value is null || value.Equals("auto", StringComparison.OrdinalIgnoreCase) ? null :
        RgbColor.TryParse(value, out var color) ? color : default(RgbColor);

    public static RgbColor Automatic(RgbColor? background)
    {
        // Word 16 independent RGB controls pin normalized BT.601 luma at 60/256.
        // Highlights do not participate; only the containing/run shading does.
        var color = background ?? new RgbColor(255, 255, 255);
        int luma = 299 * color.Red + 587 * color.Green + 114 * color.Blue;
        return luma * 256 < 60 * 255000 ? new RgbColor(255, 255, 255) : default;
    }
}
